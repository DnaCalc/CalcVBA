param(
    [string]$RunId = ('typed_chains_{0:yyyyMMddTHHmmssZ}' -f [DateTime]::UtcNow),
    [string]$OutputRoot = 'temp/calcvba14/oracle',
    [string]$RuntimeTestExecutable = ''
)
$ErrorActionPreference = 'Stop'
if ($RuntimeTestExecutable) { $RuntimeTestExecutable = (Resolve-Path -LiteralPath $RuntimeTestExecutable).Path }
$script:preboundSource = Get-Content -Raw -LiteralPath (Join-Path $PSScriptRoot '../crates/oxvba-embed/tests/support/excel_chains.bas')
$script:preboundSource += "`n" + (Get-Content -Raw -LiteralPath (Join-Path $PSScriptRoot '../docs/evidence/programs/ideal-2026-07/windows-x64/WIN-4/typed-chains-run14.bas'))
$supervisor = Get-Content -Raw -LiteralPath (Join-Path $PSScriptRoot 'run-excel-vba-oracle.ps1')
$anchor = '. (Join-Path $PSScriptRoot "excel-vba-oracle-bootstrap.ps1")'
$override = @'
$script:preboundDescriptor = @(Get-ExcelOracleHarnessCases | Where-Object id -eq 'windows-scalar-interop')[0]
$script:preboundDescriptor.module_source = $script:preboundSource
$script:preboundDescriptor.expected_value = 'A2:12:19:38:9:A2;E2:14;91:424:438;7:12:19:38:3:12:14'
$script:preboundDescriptor.purpose = 'Unchanged genuine Object/Variant native Excel chains and errors on both owning runtimes'
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
                $env:OXVBA_OWNED_EXCEL_WORKBOOK = [string]$workbook.FullName
                $env:OXVBA_OWNED_EXCEL_HWND = [string]$excel.Hwnd
                $env:OXVBA_OWNED_EXCEL_PID = [string]$ownedExcelProcess.Id
                $engineExecutable = '__ENGINE_EXECUTABLE__'
                $engineHash = (Get-FileHash -LiteralPath $engineExecutable -Algorithm SHA256).Hash.ToLowerInvariant()
                & $engineExecutable --ignored --exact owned_excel_object_variant_chains --nocapture *> (Join-Path $caseDirectory 'engine-prebound-test.log')
                $engineExit = $LASTEXITCODE
                [ordered]@{ executable = $engineExecutable; sha256 = $engineHash; exit_code = $engineExit; excel_pid = $ownedExcelProcess.Id; hwnd = [string]$excel.Hwnd; workbook = [string]$workbook.FullName; test = 'owned_excel_object_variant_chains' } | ConvertTo-Json -Depth 4 | Set-Content -LiteralPath (Join-Path $caseDirectory 'engine-prebound-test.json') -Encoding utf8NoBOM
                if ($engineExit -ne 0) { throw "Runtime prebound test failed: $engineExit" }
'@
    $runtime = $runtime.Replace('__ENGINE_EXECUTABLE__', $RuntimeTestExecutable.Replace("'", "''"))
    if (-not $worker.Contains($runAnchor)) { throw 'Worker runtime anchor changed' }
    $worker = $worker.Replace($runAnchor, $runtime)
    $worker = $worker.Replace('$PSScriptRoot', ("'" + $PSScriptRoot.Replace("'", "''") + "'"))
    $workerPath = Join-Path ([IO.Path]::GetTempPath()) ("oxvba-prebound-oracle-$([Guid]::NewGuid().ToString('N')).ps1")
    [IO.File]::WriteAllText($workerPath, $worker, [Text.UTF8Encoding]::new($false))
    $supervisor = $supervisor.Replace('(Join-Path $PSScriptRoot "excel-vba-oracle-worker.ps1")', ("'" + $workerPath.Replace("'", "''") + "'"))
}
$supervisor = $supervisor.Replace('$PSScriptRoot', ("'" + $PSScriptRoot.Replace("'", "''") + "'"))
try {
    & ([scriptblock]::Create($supervisor)) -EnvironmentId 'win-x64-dev-oracle-2026-07' -NoMatrixUpdate -SafeMode -DiagnosticCaseId 'windows-scalar-interop' -RunId $RunId -OutputRoot $OutputRoot -TimeoutSeconds 180
}
finally { if ($RuntimeTestExecutable) { Remove-Item -LiteralPath $workerPath -Force } }
