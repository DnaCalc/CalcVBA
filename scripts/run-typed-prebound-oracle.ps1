param(
    [string]$RunId = ('typed_prebound_{0:yyyyMMddTHHmmssZ}' -f [DateTime]::UtcNow),
    [string]$OutputRoot = 'temp/oracle',
    [string]$RuntimeTestExecutable = ''
)
$ErrorActionPreference = 'Stop'
if ($RuntimeTestExecutable) { $RuntimeTestExecutable = (Resolve-Path -LiteralPath $RuntimeTestExecutable).Path }
$script:preboundSource = Get-Content -Raw -LiteralPath (Join-Path $PSScriptRoot '../crates/calcvba-host/tests/support/prebound_excel.bas')
$script:preboundSource += "`n" + (Get-Content -Raw -LiteralPath (Join-Path $PSScriptRoot '../crates/calcvba-host/tests/support/oracle/typed-prebound-run12.bas'))
$supervisor = Get-Content -Raw -LiteralPath (Join-Path $PSScriptRoot 'run-excel-vba-oracle.ps1')
$anchor = '. (Join-Path $PSScriptRoot "excel-vba-oracle-bootstrap.ps1")'
$override = @'
$script:preboundDescriptor = @(Get-ExcelOracleHarnessCases | Where-Object id -eq 'windows-scalar-interop')[0]
$script:preboundDescriptor.module_source = $script:preboundSource
$script:preboundDescriptor.expected_value = 'scalar=42:42:42:43|fill=42:42:42|arrays=1.25:True:alpha:Error 2042;1.25:True:alpha:Error 2042;1.25:True:alpha:Error 2042|functions=7:5:7:10|more=2.5:5:1004|bad=1004|date=7:5|currency=6:5|kinds=7:5|6:5|scalars=11:True:8:text:10:Error 2042:0|prepared=7:6'
$script:preboundDescriptor.purpose = 'Compiler-prebound pure dispinterface Range and WorksheetFunction development evidence'
function Get-ExcelOracleHarnessCases { return @($script:preboundDescriptor) }
'@
if (-not $supervisor.Contains($anchor)) { throw 'Supervisor import anchor changed' }
$supervisor = $supervisor.Replace($anchor, $anchor + "`n" + $override)
if ($RuntimeTestExecutable) {
    $worker = Get-Content -Raw -LiteralPath (Join-Path $PSScriptRoot 'excel-vba-oracle-worker.ps1')
    $runAnchor = '                $runValue = $excel.Run($qualifiedName)'
    $runtime = @'
                $excel.VBE.MainWindow.Visible = $false
                $runValue = $excel.Run($qualifiedName)
                $workbook.Worksheets.Item(1).Range('A1:B2').NumberFormat = 'General'
                $env:OXVBA_OWNED_EXCEL_WORKBOOK = [string]$workbook.FullName
                $env:OXVBA_OWNED_EXCEL_HWND = [string]$excel.Hwnd
                $env:OXVBA_OWNED_EXCEL_PID = [string]$ownedExcelProcess.Id
                $engineExecutable = '__ENGINE_EXECUTABLE__'
                $engineHash = (Get-FileHash -LiteralPath $engineExecutable -Algorithm SHA256).Hash.ToLowerInvariant()
                & $engineExecutable --ignored --exact owned_excel_imported_version_and_build_use_typed_slots --nocapture *> (Join-Path $caseDirectory 'engine-prebound-test.log')
                $engineExit = $LASTEXITCODE
                [ordered]@{ executable = $engineExecutable; sha256 = $engineHash; exit_code = $engineExit; excel_pid = $ownedExcelProcess.Id; hwnd = [string]$excel.Hwnd; workbook = [string]$workbook.FullName; test = 'owned_excel_imported_version_and_build_use_typed_slots' } | ConvertTo-Json -Depth 4 | Set-Content -LiteralPath (Join-Path $caseDirectory 'engine-prebound-test.json') -Encoding utf8NoBOM
                if ($engineExit -ne 0) { throw "Runtime prebound test failed: $engineExit" }
'@
    $runtime = $runtime.Replace('__ENGINE_EXECUTABLE__', $RuntimeTestExecutable.Replace("'", "''"))
    if (-not $worker.Contains($runAnchor)) { throw 'Worker runtime anchor changed' }
    $worker = $worker.Replace($runAnchor, $runtime)
    $worker = $worker.Replace('$PSScriptRoot', ("'" + $PSScriptRoot.Replace("'", "''") + "'"))
    $workerPath = Join-Path ([IO.Path]::GetTempPath()) ("calcvba-prebound-oracle-$([Guid]::NewGuid().ToString('N')).ps1")
    [IO.File]::WriteAllText($workerPath, $worker, [Text.UTF8Encoding]::new($false))
    $supervisor = $supervisor.Replace('(Join-Path $PSScriptRoot "excel-vba-oracle-worker.ps1")', ("'" + $workerPath.Replace("'", "''") + "'"))
}
$supervisor = $supervisor.Replace('$PSScriptRoot', ("'" + $PSScriptRoot.Replace("'", "''") + "'"))
try {
    & ([scriptblock]::Create($supervisor)) -SafeMode -DiagnosticCaseId 'windows-scalar-interop' -RunId $RunId -OutputRoot $OutputRoot -TimeoutSeconds 180
}
finally { if ($RuntimeTestExecutable) { Remove-Item -LiteralPath $workerPath -Force } }
