param(
    [string]$RunId = ('typed_events_{0:yyyyMMddTHHmmssZ}' -f [DateTime]::UtcNow),
    [string]$OutputRoot = 'temp/oracle',
    [string]$RuntimeTestExecutable = ''
)
$ErrorActionPreference = 'Stop'
if ($RuntimeTestExecutable) { $RuntimeTestExecutable = (Resolve-Path -LiteralPath $RuntimeTestExecutable).Path }
$script:preboundSource = "Public oracleEventsActive As Boolean`n" + (Get-Content -Raw -LiteralPath (Join-Path $PSScriptRoot '../crates/oxvba-embed/tests/support/excel_events.bas'))
$script:preboundSource += "`n" + (Get-Content -Raw -LiteralPath (Join-Path $PSScriptRoot '../crates/oxvba-embed/tests/support/excel_surface.bas'))
$script:preboundSource += "`n" + (Get-Content -Raw -LiteralPath (Join-Path $PSScriptRoot '../crates/oxvba-embed/tests/support/oracle/typed-events-run13.bas'))
$supervisor = Get-Content -Raw -LiteralPath (Join-Path $PSScriptRoot 'run-excel-vba-oracle.ps1')
$anchor = '. (Join-Path $PSScriptRoot "excel-vba-oracle-bootstrap.ps1")'
$override = @'
$script:preboundDescriptor = @(Get-ExcelOracleHarnessCases | Where-Object id -eq 'windows-scalar-interop')[0]
$script:preboundDescriptor.module_source = $script:preboundSource
$script:preboundDescriptor.expected_value = 'off=1:True:20;on=2:True:20;close=1:oracle-bootstrap.xlsx'
$script:preboundDescriptor.purpose = 'Synchronous owned Excel Change reentry and BeforeClose Boolean cancellation'
function Get-ExcelOracleHarnessCases { return @($script:preboundDescriptor) }
'@
if (-not $supervisor.Contains($anchor)) { throw 'Supervisor import anchor changed' }
$supervisor = $supervisor.Replace($anchor, $anchor + "`n" + $override)
if ($RuntimeTestExecutable) {
    $worker = Get-Content -Raw -LiteralPath (Join-Path $PSScriptRoot 'excel-vba-oracle-worker.ps1')
    $documentAnchor = '        $excel.VBE.MainWindow.Visible = $true'
    $documents = @'
        $eventSheet = $workbook.Worksheets.Item(1)
        $eventSheetCodeName = [string]$eventSheet.CodeName
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($eventSheet)
        $eventModules = @(
            @{ name = $eventSheetCodeName; source = "Private Sub Worksheet_Change(ByVal changedRange As Excel.Range)`n    If $($Descriptor.module_name).oracleEventsActive Then $($Descriptor.module_name).Worksheet_Change changedRange`nEnd Sub`n" },
            @{ name = 'ThisWorkbook'; source = "Private Sub Workbook_BeforeClose(Cancel As Boolean)`n    If $($Descriptor.module_name).oracleEventsActive Then $($Descriptor.module_name).Workbook_BeforeClose Cancel`nEnd Sub`n" }
        )
        $documentEvidence = @()
        foreach ($eventModule in $eventModules) {
            $eventComponent = $project.VBComponents.Item($eventModule.name)
            [void]$eventComponent.CodeModule.AddFromString($eventModule.source)
            $eventInjected = [string]$eventComponent.CodeModule.Lines(1, $eventComponent.CodeModule.CountOfLines)
            $documentEvidence += [ordered]@{ name=$eventModule.name; selected_source=$eventModule.source; injected_source=$eventInjected; sha256=(Get-ExcelOracleSha256 -Text $eventInjected) }
            [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($eventComponent)
        }
        $documentEvidence | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $caseDirectory 'document-event-source.json') -Encoding utf8NoBOM
'@
    if (-not $worker.Contains($documentAnchor)) { throw 'Worker document injection anchor changed' }
    $worker = $worker.Replace($documentAnchor, $documents + "`n" + $documentAnchor)
    $runAnchor = '                $runValue = $excel.Run($qualifiedName)'
    $runtime = @'
                $excel.VBE.MainWindow.Visible = $false
                $runValue = $excel.Run($qualifiedName)
                $env:OXVBA_OWNED_EXCEL_WORKBOOK = [string]$workbook.FullName
                $env:OXVBA_OWNED_EXCEL_HWND = [string]$excel.Hwnd
                $env:OXVBA_OWNED_EXCEL_PID = [string]$ownedExcelProcess.Id
                $engineExecutable = '__ENGINE_EXECUTABLE__'
                $engineHash = (Get-FileHash -LiteralPath $engineExecutable -Algorithm SHA256).Hash.ToLowerInvariant()
                & $engineExecutable --ignored --exact owned_excel_events_reenter_and_cancel_before_native_return --nocapture *> (Join-Path $caseDirectory 'engine-prebound-test.log')
                $engineExit = $LASTEXITCODE
                [ordered]@{ executable = $engineExecutable; sha256 = $engineHash; exit_code = $engineExit; excel_pid = $ownedExcelProcess.Id; hwnd = [string]$excel.Hwnd; workbook = [string]$workbook.FullName; test = 'owned_excel_events_reenter_and_cancel_before_native_return' } | ConvertTo-Json -Depth 4 | Set-Content -LiteralPath (Join-Path $caseDirectory 'engine-prebound-test.json') -Encoding utf8NoBOM
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
    & ([scriptblock]::Create($supervisor)) -SafeMode -DiagnosticCaseId 'windows-scalar-interop' -RunId $RunId -OutputRoot $OutputRoot -TimeoutSeconds 180
}
finally { if ($RuntimeTestExecutable) { Remove-Item -LiteralPath $workerPath -Force } }
