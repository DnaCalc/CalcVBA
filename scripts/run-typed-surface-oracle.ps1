param(
    [string]$RunId = ('typed_surface_{0:yyyyMMddTHHmmssZ}' -f [DateTime]::UtcNow),
    [string]$OutputRoot = 'temp/oracle',
    [string]$RuntimeTestExecutable = ''
)
$ErrorActionPreference = 'Stop'
if ($RuntimeTestExecutable) { $RuntimeTestExecutable = (Resolve-Path -LiteralPath $RuntimeTestExecutable).Path }
$script:preboundSource = Get-Content -Raw -LiteralPath (Join-Path $PSScriptRoot '../crates/calcvba-host/tests/support/excel_surface.bas')
$script:preboundSource += "`n" + (Get-Content -Raw -LiteralPath (Join-Path $PSScriptRoot '../crates/calcvba-host/tests/support/oracle/typed-surface-run13.bas'))
$supervisor = Get-Content -Raw -LiteralPath (Join-Path $PSScriptRoot 'run-excel-vba-oracle.ps1')
$anchor = '. (Join-Path $PSScriptRoot "excel-vba-oracle-bootstrap.ps1")'
$override = @'
$script:preboundDescriptor = @(Get-ExcelOracleHarnessCases | Where-Object id -eq 'windows-scalar-interop')[0]
$script:preboundDescriptor.module_source = $script:preboundSource
$script:preboundDescriptor.expected_value = '7:5|6:5|$0.00|1.2346;shape=$F$2|$H$2:$I$3|$F$1:$F$3|2:2:4|F1:G2;formula==2+3:5:5.00|=G1*2:=RC[1]*2:8;app=False:False|5;range=F1:1;G1:11;F2:1;G2:11;nothing=True:9;first=F1:F1;ensure=2;sheets=2:Run13First;Run13Second;nothing=True;books=1:True;items=oracle-bootstrap.xlsx:Run13Second;nothing=424'
$script:preboundDescriptor.purpose = 'Typed native-root ordinary Excel members and NumberFormat locale provenance'
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
                & $engineExecutable --ignored --exact owned_excel_surface_from_native_window_uses_prebound_calls --nocapture *> (Join-Path $caseDirectory 'engine-prebound-test.log')
                $engineExit = $LASTEXITCODE
                [ordered]@{ executable = $engineExecutable; sha256 = $engineHash; exit_code = $engineExit; excel_pid = $ownedExcelProcess.Id; hwnd = [string]$excel.Hwnd; workbook = [string]$workbook.FullName; test = 'owned_excel_surface_from_native_window_uses_prebound_calls' } | ConvertTo-Json -Depth 4 | Set-Content -LiteralPath (Join-Path $caseDirectory 'engine-prebound-test.json') -Encoding utf8NoBOM
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
