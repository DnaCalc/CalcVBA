# The engine reads saved source itself. VBOM clears native handlers and prepares
# visible VBE compilation, without supplying the engine's project sources.
function Invoke-OfficeProjectFileChecks {
    param(
        [Parameter(Mandatory=$true)][string]$RuntimeTestExecutable,
        [Parameter(Mandatory=$true)][string]$CliPath,
        [Parameter(Mandatory=$true)][string]$RepoRoot,
        [Parameter(Mandatory=$true)][string]$BatchDirectory
    )
    $Directory=[IO.Path]::GetFullPath($BatchDirectory)
    if (Test-Path -LiteralPath $Directory) {throw 'Use a new comparison directory'}
    [void][IO.Directory]::CreateDirectory($Directory)
    $excel.EnableEvents=$false
    while ($excel.Workbooks.Count) { $old=$excel.Workbooks.Item(1); $old.Close($false); Release-ComObject $old }
    . (Join-Path $RepoRoot 'scripts/xlsm-project-oracle-hooks.ps1')
    $fixtureRoot=Join-Path $RepoRoot 'crates/calcvba-office/tests/fixtures'
    $engine=$RuntimeTestExecutable
    $batchRoot=Join-Path $Directory 'reader'
    [void][IO.Directory]::CreateDirectory($batchRoot)
    $results=@()
    foreach ($fixture in (Get-ChildItem -LiteralPath $fixtureRoot -Directory | Sort-Object Name)) {
        $case=$fixture.Name
        $caseDir=Join-Path $batchRoot $case
        [void][IO.Directory]::CreateDirectory($caseDir)
        $snapshot=Get-Content -Raw -LiteralPath (Join-Path $fixture.FullName 'project.json') | ConvertFrom-Json
        foreach ($backend in @('vm3','jit')) {
            $excel.EnableEvents=$false
            $backendDir=Join-Path $caseDir $backend
            [void][IO.Directory]::CreateDirectory($backendDir)
            $copy=Join-Path $backendDir ($case+'.xlsm')
            Copy-Item -LiteralPath (Join-Path $fixture.FullName 'workbook.xlsm') -Destination $copy
            $book=$excel.Workbooks.Open($copy)
            $project=$book.VBProject
            foreach ($component in $project.VBComponents) {
                if ($component.CodeModule.CountOfLines) { $component.CodeModule.DeleteLines(1,$component.CodeModule.CountOfLines) }
                Release-ComObject $component
            }
            Release-ComObject $project
            $excel.EnableEvents=$true
            $env:OXVBA_OWNED_EXCEL_WORKBOOK=[string]$book.FullName
            $env:OXVBA_OWNED_EXCEL_PID=[string]$owned.Id
            $env:OXVBA_OWNED_EXCEL_HWND=[string]$excel.Hwnd
            $env:OXVBA_XLSM_PROJECT=Join-Path $fixture.FullName 'project.json'
            $env:OXVBA_XLSM_BACKEND=$backend
            $env:OXVBA_XLSM_RESULT=Join-Path $backendDir 'return.json'
            & $engine --ignored --exact owned_xlsm_project --nocapture *> (Join-Path $backendDir 'engine.log')
            if ($LASTEXITCODE -ne 0) {throw "File-loaded $case / $backend failed"}
            $actualReturn=Get-Content -Raw $env:OXVBA_XLSM_RESULT | ConvertFrom-Json
            $expectedReturn=Get-Content -Raw (Join-Path $fixture.FullName 'excel-return.json') | ConvertFrom-Json
            if (($actualReturn | ConvertTo-Json -Depth 20 -Compress) -cne ($expectedReturn | ConvertTo-Json -Depth 20 -Compress)) {throw "$case / $backend return differs"}
            $actualCells=Get-XlsmCells $book
            $expectedCells=Get-Content -Raw (Join-Path $fixture.FullName 'excel-cells.json') | ConvertFrom-Json -NoEnumerate
            $actualJson=ConvertTo-Json -InputObject $actualCells -Depth 30 -Compress
            $expectedJson=ConvertTo-Json -InputObject $expectedCells -Depth 30 -Compress
            if ($actualJson -cne $expectedJson) {
                $actualJson | Set-Content -LiteralPath (Join-Path $backendDir 'actual-cells.json')
                throw "$case / $backend cells differ"
            }
            $results+=,[ordered]@{case=$case;backend=$backend;passed=$true;source='workbook';cells=@($actualCells.cells).Count}
            $results | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath (Join-Path $batchRoot 'results.json') -Encoding utf8NoBOM
            $excel.EnableEvents=$false
            $book.Close($false)
            Release-ComObject $book
        }
    }
    if ($results.Count -ne 66) {throw "Expected 66 file-loaded comparisons, got $($results.Count)"}
    $readerResults=$results

    $writeRoot=Join-Path $Directory 'writeback'
    [void][IO.Directory]::CreateDirectory($writeRoot)
    $writerResults=@()
    foreach ($extension in @('xlsm','xlsb','xls','xlam')) {
        $workbookInput=Join-Path $RepoRoot ("crates/calcvba-office/tests/formats/FormsWorkbook.$extension")
        $folder=Join-Path $writeRoot $extension
        $manifest=(& $CliPath export-vba $workbookInput --out-dir $folder).Trim()
        $api=Join-Path $folder 'Api.bas'
        $source=[IO.File]::ReadAllText($api,[Text.UTF8Encoding]::new($false))
        if (-not $source.Contains('RunProbe = "ok"')) {throw 'Changed-procedure fixture anchor missing'}
        $source=$source.Replace('RunProbe = "ok"','RunProbe = "file-edited"')
        $source+="`r`nPublic Function SourceOnlyProbe() As Long`r`n SourceOnlyProbe = 1873`r`nEnd Function`r`n"
        [IO.File]::WriteAllText($api,$source,[Text.UTF8Encoding]::new($false))
        $output=Join-Path $writeRoot ("changed.$extension")
        & $CliPath write-vba $workbookInput --project $manifest --out $output | Out-Null
        $book=$excel.Workbooks.Open($output)
        # Before any VBE edit: this result must come from the new saved source.
        $value=$excel.Run("'$($book.Name)'!Api.RunProbe")
        $added=$excel.Run("'$($book.Name)'!Api.SourceOnlyProbe")
        if ($value -cne 'file-edited' -or $added -ne 1873) {throw "Saved source/p-code mismatch for $extension"}
        Invoke-OfficeCompile $book 'Api' -Force
        $project=$book.VBProject
        $form=$project.VBComponents.Item('ExampleForm')
        # The VBE properties collection is populated after loading the designer.
        $designer=$form.Designer
        $caption=[string]$form.Properties.Item('Caption').Value
        $text=[string]$designer.Controls.Item('TextBox1').Text
        $button=[string]$designer.Controls.Item('CommandButton1').Caption
        if ($caption -cne 'Café form' -or $text -cne 'original' -or $button -cne 'Apply') {throw "Form payload changed for $extension"}
        $writerResults+=,[ordered]@{format=$extension;changed_result=$value;new_procedure=$added;visible_compile=$true;caption=$caption;text=$text;button=$button}
        Release-ComObject $designer;Release-ComObject $form;Release-ComObject $project
        $book.Close($false);Release-ComObject $book
    }
    # A generated pair must also work with VBE Import after UTF-8 -> project CP conversion.
    $importDir=Join-Path $writeRoot 'vbe-import'
    [void][IO.Directory]::CreateDirectory($importDir)
    $frm=[IO.File]::ReadAllText((Join-Path $writeRoot 'xlsm/ExampleForm.frm'),[Text.UTF8Encoding]::new($false))
    [IO.File]::WriteAllBytes((Join-Path $importDir 'ExampleForm.frm'),[Text.Encoding]::GetEncoding(1252).GetBytes($frm))
    Copy-Item -LiteralPath (Join-Path $writeRoot 'xlsm/ExampleForm.frx') -Destination (Join-Path $importDir 'ExampleForm.frx')
    $book=$excel.Workbooks.Add();$project=$book.VBProject
    $form=$project.VBComponents.Import((Join-Path $importDir 'ExampleForm.frm'));$designer=$form.Designer
    if ([string]$form.Properties.Item('Caption').Value -cne 'Café form' -or [string]$designer.Controls.Item('TextBox1').Text -cne 'original') {throw 'VBE Import changed generated form data'}
    Invoke-OfficeCompile $book 'ExampleForm' -Force
    Release-ComObject $designer;Release-ComObject $form;Release-ComObject $project
    $book.Close($false);Release-ComObject $book
    # Project properties must survive source-only loading, before any VBE edit.
    $metadataInput=Join-Path $RepoRoot 'crates/calcvba-office/tests/formats/MetadataWorkbook.xlsm'
    $metadataFolder=Join-Path $writeRoot 'metadata'
    $metadataManifest=(& $CliPath export-vba $metadataInput --out-dir $metadataFolder).Trim()
    $metadataXml=[IO.File]::ReadAllText($metadataManifest)
    $metadataXml=$metadataXml.Replace('Enabled = -1 ; Number = 255','Enabled=True;Number=&amp;HFF')
    [IO.File]::WriteAllText($metadataManifest,$metadataXml,[Text.UTF8Encoding]::new($false))
    $metadataApi=Join-Path $metadataFolder 'Api.bas'
    [IO.File]::AppendAllText($metadataApi,"`r`nPublic Function MetadataProbe() As Long`r`n#If Enabled = -1 And Number = 255 Then`r`n MetadataProbe = 1873`r`n#Else`r`n MetadataProbe = -1`r`n#End If`r`nEnd Function`r`n",[Text.UTF8Encoding]::new($false))
    $metadataOutput=Join-Path $writeRoot 'metadata.xlsm'
    & $CliPath write-vba $metadataInput --project $metadataManifest --out $metadataOutput | Out-Null
    $book=$excel.Workbooks.Open($metadataOutput)
    $project=$book.VBProject
    $metadata=[ordered]@{description=[string]$project.Description;help_file=[string]$project.HelpFile;help_context=[string]$project.HelpContextID;conditional_result=$excel.Run("'$($book.Name)'!Api.MetadataProbe")}
    if ($metadata.description -cne 'Café "project"' -or $metadata.help_file -cne 'Guide.chm' -or $metadata.help_context -cne '1873' -or $metadata.conditional_result -ne 1873) {throw "Project metadata changed: $($metadata | ConvertTo-Json -Compress)"}
    Release-ComObject $project
    Invoke-OfficeCompile $book 'Api' -Force
    $book.Close($false);Release-ComObject $book
    $summary=[ordered]@{reader=$readerResults;writeback=$writerResults;vbe_form_import=$true;project_metadata=$metadata;excel_pid=$owned.Id;excel_launches=1}
    $summary | ConvertTo-Json -Depth 12 | Set-Content -LiteralPath (Join-Path $Directory 'results.json') -Encoding utf8NoBOM
    return $summary
}
