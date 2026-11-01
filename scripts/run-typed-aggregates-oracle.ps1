param(
    [ValidateSet('Aggregates','ParamArrayEdges')][string]$Case = 'Aggregates',
    [string]$RunId = ('typed_aggregates_{0:yyyyMMddTHHmmssZ}' -f [DateTime]::UtcNow),
    [string]$OutputRoot = 'temp/calcvba11/oracle'
)
$ErrorActionPreference = 'Stop'
$filename = if ($Case -eq 'Aggregates') { 'typed-aggregates-run11.bas' } else { 'typed-paramarray-edges-run11.bas' }
$script:aggregateSource = Get-Content -Raw -LiteralPath (Join-Path $PSScriptRoot "../docs/evidence/programs/ideal-2026-07/windows-x64/WIN-4/$filename")
if ($Case -eq 'Aggregates') {
    $script:aggregateSource = $script:aggregateSource.Replace('Public Function RunProbe() As String','Public Function OriginalProbe() As String').Replace('RunProbe =','OriginalProbe =').Replace('RunProbe &','OriginalProbe &')
    $script:aggregateSource += "`nPublic Function RunProbe() As String`n    RunProbe = OriginalProbe() & `"|returns=`" & RecordReturns() & `"|prefix=`" & SumFrom(10, 2, 3)`nEnd Function`n"
    $script:aggregateExpected = 'record=42:owned:7|copy=41|aliases=44|param=9:9:0|empty=0:-1|returns=42:42:owned:7|prefix=15'
} else {
    $script:aggregateExpected = 'private=5|repeat=9:9|copy=9:27:9|bad=13:9:9|nested=0:0:True|object=True'
}
$supervisor = Get-Content -Raw -LiteralPath (Join-Path $PSScriptRoot 'run-excel-vba-oracle.ps1')
$anchor = '. (Join-Path $PSScriptRoot "excel-vba-oracle-bootstrap.ps1")'
$override = @'
$script:aggregateDescriptor = @(Get-ExcelOracleHarnessCases | Where-Object id -eq 'windows-scalar-interop')[0]
$script:aggregateDescriptor.module_source = $script:aggregateSource
$script:aggregateDescriptor.expected_value = $script:aggregateExpected
$script:aggregateDescriptor.purpose = 'Typed record/ParamArray development observation only'
function Get-ExcelOracleHarnessCases { return @($script:aggregateDescriptor) }
'@
if (-not $supervisor.Contains($anchor)) { throw 'Supervisor import anchor changed' }
$supervisor = $supervisor.Replace($anchor, $anchor + "`n" + $override)
$supervisor = $supervisor.Replace('$PSScriptRoot', ("'" + $PSScriptRoot.Replace("'", "''") + "'"))
& ([scriptblock]::Create($supervisor)) -EnvironmentId 'win-x64-dev-oracle-2026-07' -NoMatrixUpdate -SafeMode -DiagnosticCaseId 'windows-scalar-interop' -RunId $RunId -OutputRoot $OutputRoot -TimeoutSeconds 180
