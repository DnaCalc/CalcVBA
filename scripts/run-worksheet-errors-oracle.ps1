param(
    [string]$RunId = ('worksheet_errors_{0:yyyyMMddTHHmmssZ}' -f [DateTime]::UtcNow),
    [string]$OutputRoot = 'artifacts/windows-x64/worksheet-errors',
    [string]$RuntimeTestExecutable = '',
    [switch]$OwningObjectProbe,
    [ValidateSet('getobject_reads_owned_excel_workbook_values_on_both_backends', 'host_injected_owned_excel_range_uses_default_without_losing_identity', 'owned_excel_imported_version_and_build_use_typed_slots')]
    [string]$RuntimeTestName = 'getobject_reads_owned_excel_workbook_values_on_both_backends'
)
$ErrorActionPreference = 'Stop'
if ($RuntimeTestExecutable) {
    $RuntimeTestExecutable = (Resolve-Path -LiteralPath $RuntimeTestExecutable).Path
    if (-not [IO.Path]::IsPathFullyQualified($RuntimeTestExecutable) -or -not (Test-Path -LiteralPath $RuntimeTestExecutable -PathType Leaf)) { throw 'Runtime test executable must be an existing absolute file' }
}
# Reuse the exact process containment, compile authority and armed UIA guardian.
# Replace only the runtime action: no VBA macro is active during recalculation.
$script:worksheetSource = Get-Content -Raw -LiteralPath (Join-Path $PSScriptRoot '../crates/oxvba-embed/tests/support/oracle/worksheet-errors.bas')
$script:worksheetOwningObjectProbe = [bool]$OwningObjectProbe
$supervisor = Get-Content -Raw -LiteralPath (Join-Path $PSScriptRoot 'run-excel-vba-oracle.ps1')
$anchor = '. (Join-Path $PSScriptRoot "excel-vba-oracle-bootstrap.ps1")'
$override = @'
$script:worksheetDescriptor = @(Get-ExcelOracleHarnessCases | Where-Object id -eq 'windows-scalar-interop')[0]
$script:worksheetDescriptor.module_source = $script:worksheetSource
$script:worksheetDescriptor.expected_value = 'raise=#VALUE!|divide=#VALUE!|range-plus=42|literal-plus=42|range-times=82|identity=Range:5|text-plus=4|blank-plus=1|bad-plus=#VALUE!|multi-plus=#VALUE!|overflow-code=6|overflow-compare=6:True:True:7FF0000000000000'
if ($script:worksheetOwningObjectProbe) { $script:worksheetDescriptor.expected_value += '|object-owner=Range|Range|True|True|43|Range|True|13|7|7' }
$script:worksheetDescriptor.purpose = 'Idle worksheet unhandled errors and default-member arithmetic; development observation only'
$script:worksheetDescriptor.invocation_observation_prefix = $null
function Get-ExcelOracleHarnessCases { return @($script:worksheetDescriptor) }
'@
if (-not $supervisor.Contains($anchor)) { throw 'Supervisor import anchor changed' }
$supervisor = $supervisor.Replace($anchor, $anchor + "`n" + $override)
$worker = Get-Content -Raw -LiteralPath (Join-Path $PSScriptRoot 'excel-vba-oracle-worker.ps1')
$runAnchor = '                $runValue = $excel.Run($qualifiedName)'
$runtime = @'
                $excel.VBE.MainWindow.Visible = $false
                $excel.Calculation = -4135
                $sheet = $workbook.Worksheets.Item(1)
                $sheet.Range('A1').Value2 = 41.0
                $sheet.Range('A2').NumberFormat = '@'
                $sheet.Range('A2').Value2 = '3'
                [void]$sheet.Range('A3').ClearContents()
                $sheet.Range('A4').Value2 = 'bad'
                $inputs = @()
                foreach ($address in @('A1','A2','A3','A4')) {
                    $inputCell = $sheet.Range($address)
                    $inputValue = $inputCell.Value2
                    $inputs += [ordered]@{ address = $address; value = $inputValue; managed_type = if ($null -eq $inputValue) { $null } else { $inputValue.GetType().FullName }; number_format = [string]$inputCell.NumberFormat }
                }
                if ($sheet.Range('A2').Value2 -isnot [string]) { throw 'A2 must remain numeric text' }
                if ($null -ne $sheet.Range('A3').Value2) { throw 'A3 must remain blank' }
                $formulas = @('=RaiseMismatch()', '=DivideByZero(0)', '=UntypedPlus(A1)', '=UntypedPlus(41)', '=UntypedTimes(A1)', '=UntypedIdentity(A1)', '=UntypedPlus(A2)', '=UntypedPlus(A3)', '=UntypedPlus(A4)', '=UntypedPlus(A1:A2)', '=DoubleOverflowCode()', '=DoubleOverflowCompare()')
                $labels = @('raise','divide','range-plus','literal-plus','range-times','identity','text-plus','blank-plus','bad-plus','multi-plus','overflow-code','overflow-compare')
                for ($index = 0; $index -lt $formulas.Count; $index++) { $sheet.Cells.Item($index + 1, 3).Formula = $formulas[$index] }
                $excel.CalculateFullRebuild()
                $observations = @()
                $parts = @()
                for ($index = 0; $index -lt $formulas.Count; $index++) {
                    $cell = $sheet.Cells.Item($index + 1, 3)
                    $text = [string]$cell.Text
                    $raw = $cell.Value2
                    $observations += [ordered]@{ label = $labels[$index]; formula = $formulas[$index]; text = $text; value = [string]$raw; managed_type = if ($null -eq $raw) { $null } else { $raw.GetType().FullName } }
                    $parts += $labels[$index] + '=' + $text
                }
                [ordered]@{ schema = 'oxvba.worksheet-observations.v1'; calculation = 'COM CalculateFullRebuild while VBA idle'; vbe_visible_during_run = $false; inputs = $inputs; observations = $observations } | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath (Join-Path $caseDirectory 'worksheet-observations.json') -Encoding utf8NoBOM
                $runValue = $parts -join '|'
'@
if ($RuntimeTestExecutable) {
    # Reuse the live PID/HWND/workbook proof of the existing native integration
    # test. Run only its exact ignored case while this supervisor owns Excel.
    $runtimeProbe = @'
                $sheet.Range('A1').Value2 = 42.0
                $sheet.Range('B1').Value2 = 'alpha'
                $sheet.Range('A2').NumberFormat = 'General'
                $sheet.Range('A2').Value2 = $true
                $sheet.Range('B2').Formula = '=NA()'
                $excel.CalculateFullRebuild()
                $env:OXVBA_OWNED_EXCEL_WORKBOOK = [string]$workbook.FullName
                $env:OXVBA_OWNED_EXCEL_HWND = [string]$excel.Hwnd
                $env:OXVBA_OWNED_EXCEL_PID = [string]$ownedExcelProcess.Id
                $engineExecutable = '__ENGINE_EXECUTABLE__'
                $engineHash = (Get-FileHash -LiteralPath $engineExecutable -Algorithm SHA256).Hash.ToLowerInvariant()
                & $engineExecutable --ignored --exact __ENGINE_TEST_NAME__ --nocapture *> (Join-Path $caseDirectory 'engine-range-test.log')
                $engineExit = $LASTEXITCODE
                [ordered]@{ executable = $engineExecutable; sha256 = $engineHash; exit_code = $engineExit; excel_pid = $ownedExcelProcess.Id; hwnd = [string]$excel.Hwnd; workbook = [string]$workbook.FullName; test = '__ENGINE_TEST_NAME__' } | ConvertTo-Json -Depth 4 | Set-Content -LiteralPath (Join-Path $caseDirectory 'engine-range-test.json') -Encoding utf8NoBOM
                if ($engineExit -ne 0) { throw "Runtime Range test failed: $engineExit" }
'@
    $runtimeProbe = $runtimeProbe.Replace('__ENGINE_EXECUTABLE__', $RuntimeTestExecutable.Replace("'", "''"))
    $runtimeProbe = $runtimeProbe.Replace('__ENGINE_TEST_NAME__', $RuntimeTestName)
    $finishAnchor = "                `$runValue = `$parts -join '|'"
    if (-not $runtime.Contains($finishAnchor)) { throw 'Worksheet result anchor changed' }
    $runtime = $runtime.Replace($finishAnchor, $runtimeProbe + "`n" + $finishAnchor)
}
if ($OwningObjectProbe) {
    $ownerProbe = @'
                $owningName = "'" + $workbook.Name.Replace("'", "''") + "'!OwningObjectProbe"
                $owningValue = [string]$excel.Run($owningName)
                [ordered]@{ schema = 'oxvba.owning-object-observation.v1'; value = $owningValue; compile_authority = 'VBE command 578 before runtime' } | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $caseDirectory 'owning-object-observation.json') -Encoding utf8NoBOM
                $parts += 'object-owner=' + $owningValue
'@
    $finishAnchor = "                `$runValue = `$parts -join '|'"
    if (-not $runtime.Contains($finishAnchor)) { throw 'Worksheet result anchor changed' }
    $runtime = $runtime.Replace($finishAnchor, $ownerProbe + "`n" + $finishAnchor)
}
if (-not $worker.Contains($runAnchor)) { throw 'Worker runtime anchor changed' }
$worker = $worker.Replace($runAnchor, $runtime)
$worker = $worker.Replace('$PSScriptRoot', ("'" + $PSScriptRoot.Replace("'", "''") + "'"))
$workerPath = Join-Path ([IO.Path]::GetTempPath()) ("oxvba-worksheet-oracle-$([Guid]::NewGuid().ToString('N')).ps1")
[IO.File]::WriteAllText($workerPath, $worker, [Text.UTF8Encoding]::new($false))
$supervisor = $supervisor.Replace('(Join-Path $PSScriptRoot "excel-vba-oracle-worker.ps1")', ("'" + $workerPath.Replace("'", "''") + "'"))
$supervisor = $supervisor.Replace('$PSScriptRoot', ("'" + $PSScriptRoot.Replace("'", "''") + "'"))
try {
    & ([scriptblock]::Create($supervisor)) -SafeMode -DiagnosticCaseId 'windows-scalar-interop' -RunId $RunId -OutputRoot $OutputRoot -TimeoutSeconds 180
}
finally { Remove-Item -LiteralPath $workerPath -Force }
