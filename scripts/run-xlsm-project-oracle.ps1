param(
    [string]$CaseId='chains',
    [Parameter(Mandatory=$true)][string]$RuntimeTestExecutable,
    [string]$RunId=('xlsm_{0:yyyyMMddTHHmmssZ}' -f [DateTime]::UtcNow)
)
$ErrorActionPreference='Stop'
$oracleScripts=$PSScriptRoot
$engine=(Resolve-Path -LiteralPath $RuntimeTestExecutable).Path
$fixtures=(Resolve-Path (Join-Path $PSScriptRoot '../crates/oxvba-embed/tests/support')).Path
$extraCases=Get-Content -Raw (Join-Path $fixtures 'xlsm/cases.json') | ConvertFrom-Json
$extraCase=@($extraCases | Where-Object id -CEQ $CaseId)
if ($CaseId -notin @('chains','events','mixed','documents','arrays','bounds','application') -and $extraCase.Count -ne 1) { throw "Unknown saved workbook case: $CaseId" }
$entry=if ($CaseId -eq 'mixed') {'WorkbookMain'} else {'Api'}
$chainResult='A2:12:19:38:9:A2;E2:14;91:424:438;7:12:19:38:3:12:14'
$eventResult='1:1:WA1;DA1;|nested=3:3:WA1;DA1;WA2;DA2;|rebind=4:WA3;DA3;|release=4:1:True:DA5;|close=1'
$script:xlsmMain=Get-Content -Raw (Join-Path $fixtures $(if ($extraCase.Count -eq 1) {'xlsm/'+$CaseId+'-entry.bas'} elseif ($CaseId -in @('chains','documents','arrays','bounds')) {'xlsm/'+$CaseId+'-entry.bas'} elseif ($CaseId -eq 'application') {'application/Api.bas'} else {'automatic/Api.bas'}))
$script:xlsmExtras=@()
if ($CaseId -in @('events','mixed')) {$script:xlsmExtras+=,[pscustomobject]@{name='Watcher';kind=2;source=(Get-Content -Raw (Join-Path $fixtures 'automatic/Watcher.cls'))}}
if ($CaseId -eq 'application') {$script:xlsmExtras+=,[pscustomobject]@{name='AppWatcher';kind=2;source=(Get-Content -Raw (Join-Path $fixtures 'application/AppWatcher.cls'))}}
if ($CaseId -in @('chains','mixed')) {$script:xlsmExtras+=,[pscustomobject]@{name='ChainOps';kind=1;source=(Get-Content -Raw (Join-Path $fixtures 'excel_chains.bas'))}}
if ($CaseId -eq 'mixed') {$script:xlsmExtras+=,[pscustomobject]@{name='WorkbookMain';kind=1;source=(Get-Content -Raw (Join-Path $fixtures 'xlsm/mixed-entry.bas'))}}
$script:xlsmExpected=switch ($CaseId) {'chains' {$chainResult};'events' {$eventResult};'mixed' {$eventResult+';chains='+$chainResult};'documents' {'True:True:True:True:2'};'arrays' {'1:4:10'};'bounds' {'1:4:10'};'application' {'1:1:DA:True'}}
if ($extraCase.Count -eq 1) {
    $script:xlsmExpected=[string]$extraCase[0].expected
    foreach ($extra in $extraCase[0].extras) {
        $script:xlsmExtras+=,[pscustomobject]@{name=$extra.name;kind=$extra.kind;source=(Get-Content -Raw (Join-Path $fixtures ('xlsm/'+$extra.file)))}
    }
}
$script:xlsmEntry=$entry
$script:xlsmWorkbookName=$CaseId+'.xlsm'
$expectsError=$extraCase.Count -eq 1 -and $extraCase[0].PSObject.Properties['expected_error']
$script:xlsmOracleCase=if ($expectsError) {'runtime-unhandled-modal'} else {'windows-scalar-interop'}
if ($expectsError) { $script:xlsmExpected=$null }
$supervisor=Get-Content -Raw (Join-Path $PSScriptRoot 'run-excel-vba-oracle.ps1')
$anchor='. (Join-Path $PSScriptRoot "excel-vba-oracle-bootstrap.ps1")'
$override=@'
$script:xlsmDescriptor=@(Get-ExcelOracleHarnessCases | Where-Object id -eq $script:xlsmOracleCase)[0]
$script:xlsmDescriptor.module_source=$script:xlsmMain
$script:xlsmDescriptor.module_name='Api'
$script:xlsmDescriptor.run_procedure=$script:xlsmEntry+'.RunProbe'
$script:xlsmDescriptor.macro_probe_target=$script:xlsmEntry+'.RunProbe'
$script:xlsmDescriptor.expected_value=$script:xlsmExpected
$script:xlsmDescriptor.purpose='Saved and reopened real XLSM source project comparison'
if ($script:xlsmExtras.Count -gt 0) {$script:xlsmDescriptor | Add-Member -NotePropertyName additional_modules -NotePropertyValue $script:xlsmExtras}
$script:xlsmDescriptor | Add-Member -NotePropertyName saved_workbook_name -NotePropertyValue $script:xlsmWorkbookName
function Get-ExcelOracleHarnessCases {return @($script:xlsmDescriptor)}
'@
$supervisor=$supervisor.Replace($anchor,$anchor+"`n"+$override)
$worker=Get-Content -Raw (Join-Path $PSScriptRoot 'excel-vba-oracle-worker.ps1')
if ($expectsError) {
    # Compare-XlsmRuntimes captures the native failure before running both engine
    # backends. Preserve the same derived modal disposition as the base worker;
    # the supervisor still independently validates the exact owned modal/dismissal.
    $worker=$worker.Replace('$runStatus = "runtime-error-modal"', '$runStatus = "runtime-error-modal"; $macroDisposition = "non-macro-runtime-failure"')
}
if (-not $worker.Contains('        $excel.VBE.MainWindow.Visible = $true') -or
    -not $worker.Contains('                $runValue = $excel.Run($qualifiedName)')) {
    throw 'Owned oracle worker insertion points changed; saved-project comparison was not installed'
}
$initialize=@'
        . '__SCRIPTS__/xlsm-project-oracle-hooks.ps1'
        Initialize-XlsmProject $excel ([ref]$workbook) ([ref]$project) ([ref]$component) $caseDirectory '__CASE__' '__ENTRY__' '__DOCUMENT_ROOT__'
'@
$worker=$worker.Replace('        $excel.VBE.MainWindow.Visible = $true',$initialize+"`n"+'        $excel.VBE.MainWindow.Visible = $true')
$run=@'
                $excel.VBE.MainWindow.Visible=$false
                $runValue=Compare-XlsmRuntimes $excel ([ref]$workbook) ([ref]$project) ([ref]$component) $caseDirectory '__ENTRY__' $qualifiedName '__ENGINE__' $ownedExcelProcess.Id
'@
$worker=$worker.Replace('                $runValue = $excel.Run($qualifiedName)',$run)
$documentRoot=Join-Path $fixtures $(if ($CaseId -eq 'application') {'application'} else {'automatic'})
foreach ($pair in @(@('__SCRIPTS__',$oracleScripts),@('__FIXTURES__',$fixtures),@('__DOCUMENT_ROOT__',$documentRoot),@('__CASE__',$CaseId),@('__ENTRY__',$entry),@('__ENGINE__',$engine))) {
    $worker=$worker.Replace($pair[0],$pair[1].Replace('\','/').Replace("'","''"))
}
$worker=$worker.Replace('$PSScriptRoot',("'"+$oracleScripts.Replace("'","''")+"'"))
$workerPath=Join-Path ([IO.Path]::GetTempPath()) ('run14-xlsm-'+[Guid]::NewGuid().ToString('N')+'.ps1')
[IO.File]::WriteAllText($workerPath,$worker,[Text.UTF8Encoding]::new($false))
$supervisor=$supervisor.Replace('(Join-Path $PSScriptRoot "excel-vba-oracle-worker.ps1")',("'"+$workerPath.Replace("'","''")+"'"))
$supervisor=$supervisor.Replace('$PSScriptRoot',("'"+$oracleScripts.Replace("'","''")+"'"))
try {& ([scriptblock]::Create($supervisor)) -SafeMode -DiagnosticCaseId $script:xlsmOracleCase -RunId $RunId -OutputRoot 'temp/oracle' -TimeoutSeconds 360}
finally {Remove-Item -LiteralPath $workerPath -Force}
