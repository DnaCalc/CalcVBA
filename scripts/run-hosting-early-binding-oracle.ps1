param(
    [ValidateSet('public', 'private-module', 'dictionary')][string]$Case = 'public',
    [string]$RunId = ('hosting_early_{0:yyyyMMddTHHmmssZ}' -f [DateTime]::UtcNow),
    [string]$OutputRoot = 'artifacts/windows-x64/hosting-early'
)
$ErrorActionPreference = 'Stop'
# Reuse the owned-process supervisor, UIA guardian and explicit VBE compile.
# No Office binaries or security settings are inspected or changed.
$sourceName = if ($Case -eq 'dictionary') { 'early-dictionary.bas' } else { 'hosting-public.bas' }
$sourcePath = Join-Path $PSScriptRoot "../temp/oracle/programs/ideal-2026-07/windows-x64/WIN-4/$sourceName"
$script:hostingSource = Get-Content -Raw -LiteralPath $sourcePath
if ($Case -eq 'private-module') { $script:hostingSource = "Option Private Module`n" + $script:hostingSource }
$script:hostingCaseId = 'windows-scalar-interop'
# Frozen after the initial black-box readings; final receipts assert the exact
# tuple rather than accepting an arbitrary successful string from RunProbe.
$script:hostingExpected = if ($Case -eq 'dictionary') {
    'value=42|count=1|exists=True|missing=Empty|after=2|duplicate=457|source=VBAProject|description=This key is already associated with an element of this collection'
} else {
    'cell-public=2|cell-private=73|cell-sub=0|cell-text=3|cell-badtext=Error 2015|cell-cverr=Error 2042|cell-range=Range:5|cell-number=Double:5|run-long=Long:2|run-private=Long:73|run-sub=Empty:|run-cverr=Error:Error 2042|external-public=2|external-private=73'
}
$supervisor = Get-Content -Raw -LiteralPath (Join-Path $PSScriptRoot 'run-excel-vba-oracle.ps1')
$anchor = '. (Join-Path $PSScriptRoot "excel-vba-oracle-bootstrap.ps1")'
$override = @'
$script:hostingDescriptor = @(Get-ExcelOracleHarnessCases | Where-Object id -eq 'windows-scalar-interop')[0]
$script:hostingDescriptor.id = $script:hostingCaseId
$script:hostingDescriptor.module_source = $script:hostingSource
$script:hostingDescriptor.expected_value = $script:hostingExpected
$script:hostingDescriptor.purpose = 'Black-box procedure visibility, cell/Run coercion and errors; development observation only'
function Get-ExcelOracleHarnessCases { return @($script:hostingDescriptor) }
'@
if (-not $supervisor.Contains($anchor)) { throw 'Supervisor import anchor changed' }
$supervisor = $supervisor.Replace($anchor, $anchor + "`n" + $override)
$workerPath = $null
if ($Case -eq 'dictionary') {
    # Add a workbook-local reference to the installed public Scripting typelib
    # before the explicit compile. This reads registration; it never writes it.
    $worker = Get-Content -Raw -LiteralPath (Join-Path $PSScriptRoot 'excel-vba-oracle-worker.ps1')
    $projectAnchor = '        $project = $workbook.VBProject'
    if (-not $worker.Contains($projectAnchor)) { throw 'Worker project anchor changed' }
    $worker = $worker.Replace($projectAnchor, $projectAnchor + "`n" + '        [void]$project.References.AddFromGuid("{420B2830-E718-11CF-893D-00A0C9054228}", 1, 0)')
    $worker = $worker.Replace('$PSScriptRoot', ("'" + $PSScriptRoot.Replace("'", "''") + "'"))
    $workerPath = Join-Path ([IO.Path]::GetTempPath()) ("calcvba-early-oracle-$([Guid]::NewGuid().ToString('N')).ps1")
    [IO.File]::WriteAllText($workerPath, $worker, [Text.UTF8Encoding]::new($false))
    $supervisor = $supervisor.Replace('(Join-Path $PSScriptRoot "excel-vba-oracle-worker.ps1")', ("'" + $workerPath.Replace("'", "''") + "'"))
}
$supervisor = $supervisor.Replace('$PSScriptRoot', ("'" + $PSScriptRoot.Replace("'", "''") + "'"))
try {
    & ([scriptblock]::Create($supervisor)) -SafeMode -DiagnosticCaseId $script:hostingCaseId -RunId $RunId -OutputRoot $OutputRoot -TimeoutSeconds 180
}
finally { if ($workerPath) { Remove-Item -LiteralPath $workerPath -Force } }
