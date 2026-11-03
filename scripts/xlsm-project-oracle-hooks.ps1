# Public VBOM is the host project loader. No VBA binary-stream parser is claimed.
function Get-XlsmScalar($Value) {
    if ($null -eq $Value) { return [ordered]@{kind='Empty'} }
    if ($Value -is [string]) { return [ordered]@{kind='String';units=@($Value.ToCharArray() | ForEach-Object {[int]$_})} }
    return [ordered]@{kind=$Value.GetType().FullName;value=$Value}
}
function Get-XlsmCells($Workbook) {
    $sheets=@()
    foreach ($sheet in $Workbook.Worksheets) {
        $used=$sheet.UsedRange
        $cells=@()
        if ([long]$used.Cells.Count -gt 4096) { throw 'Fixture exceeded the bounded whole-used-range comparison' }
        for ($r=1;$r -le $used.Rows.Count;$r++) {
            for ($c=1;$c -le $used.Columns.Count;$c++) {
                $cell=$used.Cells.Item($r,$c)
                $cells+=,[ordered]@{row=[int]$cell.Row;column=[int]$cell.Column;value=(Get-XlsmScalar $cell.Value2);formula=(Get-XlsmScalar $cell.Formula);text=[string]$cell.Text;number_format=[string]$cell.NumberFormat;fill=[int]$cell.Interior.Color;bold=[bool]$cell.Font.Bold;merged=[bool]$cell.MergeCells}
                [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($cell)
            }
        }
        $sheets+=,[ordered]@{name=[string]$sheet.Name;code_name=[string]$sheet.CodeName;rows=[int]$used.Rows.Count;columns=[int]$used.Columns.Count;cells=$cells}
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($used)
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($sheet)
    }
    return ,$sheets
}
function Get-XlsmProject($Workbook,[string]$EntryModule) {
    $hosts=@{([string]$Workbook.CodeName)='Excel.Workbook'}
    foreach ($sheet in $Workbook.Worksheets) {
        $hosts[[string]$sheet.CodeName]='Excel.Worksheet'
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($sheet)
    }
    $project=$Workbook.VBProject
    $modules=@()
    foreach ($component in $project.VBComponents) {
        $source=if ($component.CodeModule.CountOfLines -gt 0) {[string]$component.CodeModule.Lines(1,$component.CodeModule.CountOfLines)} else {''}
        $hash=if ($source.Length -gt 0) {Get-ExcelOracleSha256 -Text $source} else {'sha256:e3b0c44298fc1c149afbf4c8996fb92427ae41e4649b934ca495991b7852b855'}
        $modules+=,[ordered]@{name=[string]$component.Name;kind=[int]$component.Type;source=$source;source_sha256=$hash;host_type=$hosts[[string]$component.Name]}
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($component)
    }
    $references=@()
    foreach ($reference in $project.References) {
        if ([int]$reference.Type -ne 0) {throw 'Project reference requires a separate host project loader'}
        $references+=,[ordered]@{name=[string]$reference.Name;guid=[string]$reference.GUID;major=[int]$reference.Major;minor=[int]$reference.Minor;broken=[bool]$reference.IsBroken;path=[string]$reference.FullPath}
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($reference)
    }
    $result=[ordered]@{name=[string]$project.Name;entry_module=$EntryModule;modules=$modules;references=$references}
    return $result
}
function Initialize-XlsmProject($Excel,[ref]$Workbook,[ref]$Project,[ref]$Component,[string]$Directory,[string]$CaseId,[string]$EntryModule,[string]$SourceRoot) {
    if ($CaseId -in @('events','mixed','application')) {
        $documentSources=,@([string]$Workbook.Value.CodeName,'ThisWorkbook.cls')
        if ($CaseId -ne 'application') {$documentSources+=,@([string]$Workbook.Value.Worksheets.Item(1).CodeName,'Sheet1.cls')}
        foreach ($pair in $documentSources) {
            $doc=$Project.Value.VBComponents.Item($pair[0])
            $source=Get-Content -Raw -LiteralPath (Join-Path $SourceRoot $pair[1])
            [void]$doc.CodeModule.AddFromString($source)
            if ((Get-ExcelOracleSha256 -Text $source) -cne (Get-ExcelOracleSha256 -Text ([string]$doc.CodeModule.Lines(1,$doc.CodeModule.CountOfLines)))) {throw 'Document source changed during injection'}
            [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($doc)
        }
    }
    $path=Join-Path $Directory ($CaseId+'.xlsm')
    $compileModule=[string]$Component.Value.Name
    $Excel.EnableEvents=$false
    $Workbook.Value.SaveAs($path,52)
    [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($Component.Value)
    [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($Project.Value)
    $Workbook.Value.Close($false)
    [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($Workbook.Value)
    $Workbook.Value=$Excel.Workbooks.Open($path)
    $Project.Value=$Workbook.Value.VBProject
    $Component.Value=$Project.Value.VBComponents.Item($compileModule)
    # Reinsert identical text so the explicit VBE compile command is enabled.
    $source=[string]$Component.Value.CodeModule.Lines(1,$Component.Value.CodeModule.CountOfLines)
    $Component.Value.CodeModule.DeleteLines(1,$Component.Value.CodeModule.CountOfLines)
    [void]$Component.Value.CodeModule.AddFromString($source)
    $snapshot=Get-XlsmProject $Workbook.Value $EntryModule
    $snapshot | ConvertTo-Json -Depth 20 | Set-Content -LiteralPath (Join-Path $Directory 'project.json') -Encoding utf8NoBOM
    $Project.Value=$Workbook.Value.VBProject
    $Component.Value=$Project.Value.VBComponents.Item($compileModule)
    $Excel.EnableEvents=$true
}
function Compare-XlsmRuntimes($Excel,[ref]$Workbook,[ref]$Project,[ref]$Component,[string]$Directory,[string]$EntryModule,[string]$QualifiedName,[string]$Engine,[int]$OwnedPid) {
    $referenceValue=$Excel.Run($QualifiedName)
    $referenceReturn=Get-XlsmScalar $referenceValue
    $referenceCells=Get-XlsmCells $Workbook.Value
    $referenceReturn | ConvertTo-Json -Depth 20 | Set-Content (Join-Path $Directory 'excel-return.json') -Encoding utf8NoBOM
    ConvertTo-Json -InputObject $referenceCells -Depth 30 | Set-Content (Join-Path $Directory 'excel-cells.json') -Encoding utf8NoBOM
    $baseline=[string]$Workbook.Value.FullName
    $snapshot=Get-Content -Raw (Join-Path $Directory 'project.json') | ConvertFrom-Json
    $comparisons=@()
    foreach ($backend in @('vm3','jit')) {
        $Excel.EnableEvents=$false
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($Component.Value)
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($Project.Value)
        $Workbook.Value.Close($false)
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($Workbook.Value)
        $backendDirectory=Join-Path $Directory $backend
        New-Item -ItemType Directory -Force $backendDirectory | Out-Null
        $copy=Join-Path $backendDirectory ([IO.Path]::GetFileName($baseline))
        Copy-Item -LiteralPath $baseline -Destination $copy
        $workbookSha=(Get-FileHash $copy -Algorithm SHA256).Hash.ToLowerInvariant()
        $Workbook.Value=$Excel.Workbooks.Open($copy)
        $actual=Get-XlsmProject $Workbook.Value $EntryModule
        $actualJson=$actual | ConvertTo-Json -Depth 20 -Compress
        if ($actualJson -cne ($snapshot | ConvertTo-Json -Depth 20 -Compress)) {throw 'Reopened source project differs from the reference snapshot'}
        $snapshotPath=Join-Path $backendDirectory 'project.json'
        $actual | ConvertTo-Json -Depth 20 | Set-Content $snapshotPath -Encoding utf8NoBOM
        $Project.Value=$Workbook.Value.VBProject
        # Remove the native VBA handlers only in memory; the saved XLSM retains
        # the complete original project. CalcVBA receives its unchanged source.
        foreach ($doc in $Project.Value.VBComponents) {
            if ($doc.CodeModule.CountOfLines -gt 0) {$doc.CodeModule.DeleteLines(1,$doc.CodeModule.CountOfLines)}
            [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($doc)
        }
        $Component.Value=$Project.Value.VBComponents.Item($EntryModule)
        $Excel.EnableEvents=$true
        $env:OXVBA_OWNED_EXCEL_WORKBOOK=[string]$Workbook.Value.FullName
        $env:OXVBA_OWNED_EXCEL_PID=[string]$OwnedPid
        $env:OXVBA_OWNED_EXCEL_HWND=[string]$Excel.Hwnd
        $env:OXVBA_XLSM_PROJECT=$snapshotPath
        $env:OXVBA_XLSM_BACKEND=$backend
        $env:OXVBA_XLSM_RESULT=Join-Path $backendDirectory 'engine-return.json'
        & $Engine --ignored --exact owned_xlsm_project --nocapture *> (Join-Path $backendDirectory 'engine.log')
        $exit=$LASTEXITCODE
        if ($exit -ne 0) {throw "Saved XLSM $backend execution failed: $exit"}
        $cells=Get-XlsmCells $Workbook.Value
        ConvertTo-Json -InputObject $cells -Depth 30 | Set-Content (Join-Path $backendDirectory 'engine-cells.json') -Encoding utf8NoBOM
        $result=Get-Content -Raw $env:OXVBA_XLSM_RESULT | ConvertFrom-Json
        $sameReturn=($result | ConvertTo-Json -Depth 20 -Compress) -ceq ($referenceReturn | ConvertTo-Json -Depth 20 -Compress)
        $sameCells=(ConvertTo-Json -InputObject $cells -Depth 30 -Compress) -ceq (ConvertTo-Json -InputObject $referenceCells -Depth 30 -Compress)
        $comparisons+=,[ordered]@{backend=$backend;exit_code=$exit;returned_value_equal=$sameReturn;all_used_cells_equal=$sameCells;workbook_sha256=$workbookSha;engine_sha256=(Get-FileHash $Engine -Algorithm SHA256).Hash.ToLowerInvariant();owned_excel_pid=$OwnedPid}
        $comparisons | ConvertTo-Json -Depth 10 | Set-Content (Join-Path $Directory 'comparisons.json') -Encoding utf8NoBOM
        if (-not $sameReturn -or -not $sameCells) {throw "Saved XLSM $backend differs: return=$sameReturn cells=$sameCells"}
    }
    $Excel.EnableEvents=$false
    return $referenceValue
}
