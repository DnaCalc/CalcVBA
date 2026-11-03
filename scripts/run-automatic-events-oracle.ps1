param([string]$RunId=('automatic_events_{0:yyyyMMddTHHmmssZ}' -f [DateTime]::UtcNow), [string]$RuntimeTestExecutable='')
$ErrorActionPreference='Stop'
$oracleScripts=$PSScriptRoot
if ($RuntimeTestExecutable) { $RuntimeTestExecutable=(Resolve-Path -LiteralPath $RuntimeTestExecutable).Path }
$sourceRoot=(Resolve-Path (Join-Path $PSScriptRoot '../crates/calcvba-host/tests/support/automatic')).Path
$script:preboundSource=Get-Content -Raw (Join-Path $sourceRoot 'Api.bas')
$script:watchSource=Get-Content -Raw (Join-Path $sourceRoot 'Watcher.cls')
$supervisor=Get-Content -Raw (Join-Path $oracleScripts 'run-excel-vba-oracle.ps1')
$anchor='. (Join-Path $PSScriptRoot "excel-vba-oracle-bootstrap.ps1")'
$override=@'
$script:preboundDescriptor=@(Get-ExcelOracleHarnessCases | Where-Object id -eq 'windows-scalar-interop')[0]
$script:preboundDescriptor.module_source=$script:preboundSource
$script:preboundDescriptor.module_name='Api'
$script:preboundDescriptor.run_procedure='Api.RunProbe'
$script:preboundDescriptor.macro_probe_target='Api.RunProbe'
$script:preboundDescriptor.expected_value='1:1:WA1;DA1;|nested=3:3:WA1;DA1;WA2;DA2;|rebind=4:WA3;DA3;|release=4:1:True:DA5;|close=1|idle=1:DZ1;'
$script:preboundDescriptor.purpose='Automatic document and WithEvents routing VBA observation'
$script:preboundDescriptor | Add-Member -NotePropertyName additional_modules -NotePropertyValue @([pscustomobject]@{name='Watcher';kind=2;source=$script:watchSource})
function Get-ExcelOracleHarnessCases {return @($script:preboundDescriptor)}
'@
$supervisor=$supervisor.Replace($anchor,$anchor+"`n"+$override)
$worker=Get-Content -Raw (Join-Path $oracleScripts 'excel-vba-oracle-worker.ps1')
$documentAnchor='        $excel.VBE.MainWindow.Visible = $true'
if (-not $worker.Contains($documentAnchor)) {throw 'Owned oracle document insertion point changed'}
$documents=@'
        $docSources=@(@{name=[string]$workbook.Worksheets.Item(1).CodeName;path='__SOURCE_ROOT__/Sheet1.cls'},@{name='ThisWorkbook';path='__SOURCE_ROOT__/ThisWorkbook.cls'})
        $docEvidence=@()
        foreach($item in $docSources){
            $doc=$project.VBComponents.Item($item.name)
            $source=Get-Content -Raw -LiteralPath $item.path
            [void]$doc.CodeModule.AddFromString($source)
            $injected=[string]$doc.CodeModule.Lines(1,$doc.CodeModule.CountOfLines)
            if ((Get-ExcelOracleSha256 -Text $source) -cne (Get-ExcelOracleSha256 -Text $injected)) { throw 'Document module source changed during injection' }
            $docEvidence+=@{name=$item.name;source=$source;injected=$injected;sha256=(Get-ExcelOracleSha256 -Text $injected)}
            [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($doc)
        }
        $docEvidence | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $caseDirectory 'documents.json') -Encoding utf8NoBOM
'@
$worker=$worker.Replace($documentAnchor,$documents+"`n"+$documentAnchor)
$worker=$worker.Replace('__SOURCE_ROOT__',$sourceRoot.Replace('\','/').Replace("'","''"))
$runtime=@'
                $excel.VBE.MainWindow.Visible=$false
                $runValue=$excel.Run($qualifiedName)
                [void]$excel.Run("'"+$workbook.Name+"'!Api.ResetIdle")
                $workbook.Worksheets.Item(1).Range('Z1').Value2=17
                $runValue+='|idle='+$excel.Run("'"+$workbook.Name+"'!Api.ReadIdle")
                $excel.EnableEvents=$false
                if ('__ENGINE__') {
                    # Retain the unchanged VBA source as evidence, then remove the
                    # native VBA document handlers from this owned execution target.
                    # CalcVBA selects and wires those same handlers itself.
                    foreach ($item in $docSources) {
                        $doc=$project.VBComponents.Item($item.name)
                        if ($doc.CodeModule.CountOfLines -gt 0) { $doc.CodeModule.DeleteLines(1,$doc.CodeModule.CountOfLines) }
                        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($doc)
                    }
                    $excel.EnableEvents=$true
                    $env:OXVBA_OWNED_EXCEL_WORKBOOK=[string]$workbook.FullName
                    $env:OXVBA_OWNED_EXCEL_HWND=[string]$excel.Hwnd
                    $env:OXVBA_OWNED_EXCEL_PID=[string]$ownedExcelProcess.Id
                    $engine='__ENGINE__'
                    & $engine --ignored --exact owned_excel_automatic_events --nocapture *> (Join-Path $caseDirectory 'engine-automatic-test.log')
                    $engineExit=$LASTEXITCODE
                    [ordered]@{executable=$engine;sha256=(Get-FileHash -LiteralPath $engine -Algorithm SHA256).Hash.ToLowerInvariant();exit_code=$engineExit;excel_pid=$ownedExcelProcess.Id;hwnd=[string]$excel.Hwnd;workbook=[string]$workbook.FullName;test='owned_excel_automatic_events'} | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $caseDirectory 'engine-automatic-test.json') -Encoding utf8NoBOM
                    $excel.EnableEvents=$false
                    if ($engineExit -ne 0) {throw "Automatic event runtime test failed: $engineExit"}
                }
'@
$runtime=$runtime.Replace('__ENGINE__',$RuntimeTestExecutable.Replace("'","''"))
$runAnchor='                $runValue = $excel.Run($qualifiedName)'
if (-not $worker.Contains($runAnchor)) {throw 'Worker run anchor changed'}
$worker=$worker.Replace($runAnchor,$runtime)
$worker=$worker.Replace('$PSScriptRoot',("'"+$oracleScripts.Replace("'","''")+"'"))
$workerPath=Join-Path ([IO.Path]::GetTempPath()) ('run14-auto-'+[Guid]::NewGuid().ToString('N')+'.ps1')
[IO.File]::WriteAllText($workerPath,$worker,[Text.UTF8Encoding]::new($false))
$supervisor=$supervisor.Replace('(Join-Path $PSScriptRoot "excel-vba-oracle-worker.ps1")',("'"+$workerPath.Replace("'","''")+"'"))
$supervisor=$supervisor.Replace('$PSScriptRoot',("'"+$oracleScripts.Replace("'","''")+"'"))
try {& ([scriptblock]::Create($supervisor)) -SafeMode -DiagnosticCaseId 'windows-scalar-interop' -RunId $RunId -OutputRoot 'temp/oracle' -TimeoutSeconds 360}
finally {Remove-Item -LiteralPath $workerPath -Force}
