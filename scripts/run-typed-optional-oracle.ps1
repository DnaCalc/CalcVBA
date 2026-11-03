param(
    [string]$RunId = ('typed_optional_{0:yyyyMMddTHHmmssZ}' -f [DateTime]::UtcNow),
    [string]$OutputRoot = 'temp/oracle'
)
$ErrorActionPreference = 'Stop'
$script:defaultsSource = Get-Content -Raw -LiteralPath (Join-Path $PSScriptRoot '../crates/oxvba-embed/tests/support/oracle/typed-optional-run10.bas')
$script:defaultsExpected = 'default=42|alias=42:42|private=44:42|missing=True|present=False|zero=0|named=6.5'
$supervisor = Get-Content -Raw -LiteralPath (Join-Path $PSScriptRoot 'run-excel-vba-oracle.ps1')
$anchor = '. (Join-Path $PSScriptRoot "excel-vba-oracle-bootstrap.ps1")'
$override = @'
$script:defaultsDescriptor = @(Get-ExcelOracleHarnessCases | Where-Object id -eq 'windows-scalar-interop')[0]
$script:defaultsDescriptor.module_source = $script:defaultsSource
$script:defaultsDescriptor.expected_value = $script:defaultsExpected
$script:defaultsDescriptor.purpose = 'Optional/defaulted parameter and alias development observation only'
function Get-ExcelOracleHarnessCases { return @($script:defaultsDescriptor) }
'@
if (-not $supervisor.Contains($anchor)) { throw 'Supervisor import anchor changed' }
$supervisor = $supervisor.Replace($anchor, $anchor + "`n" + $override)
$supervisor = $supervisor.Replace('$PSScriptRoot', ("'" + $PSScriptRoot.Replace("'", "''") + "'"))
& ([scriptblock]::Create($supervisor)) -SafeMode -DiagnosticCaseId 'windows-scalar-interop' -RunId $RunId -OutputRoot $OutputRoot -TimeoutSeconds 180
