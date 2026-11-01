param(
    [string]$RunId = ('typed_identity_{0:yyyyMMddTHHmmssZ}' -f [DateTime]::UtcNow),
    [string]$OutputRoot = 'temp/calcvba11/oracle'
)
$ErrorActionPreference = 'Stop'
$script:identitySource = Get-Content -Raw -LiteralPath (Join-Path $PSScriptRoot '../docs/evidence/programs/ideal-2026-07/windows-x64/WIN-4/typed-identity-run11.bas')
$supervisor = Get-Content -Raw -LiteralPath (Join-Path $PSScriptRoot 'run-excel-vba-oracle.ps1')
$anchor = '. (Join-Path $PSScriptRoot "excel-vba-oracle-bootstrap.ps1")'
$override = @'
$script:identityDescriptor = @(Get-ExcelOracleHarnessCases | Where-Object id -eq 'windows-scalar-interop')[0]
$script:identityDescriptor.module_source = $script:identitySource
$script:identityDescriptor.expected_value = 'same=True|different=False|boxed=True|nothing=True|one=False|empty=424|null=424|long=424|text=424|error=424'
$script:identityDescriptor.purpose = 'Typed Is identity/Nothing/object-required development observation only'
function Get-ExcelOracleHarnessCases { return @($script:identityDescriptor) }
'@
if (-not $supervisor.Contains($anchor)) { throw 'Supervisor import anchor changed' }
$supervisor = $supervisor.Replace($anchor, $anchor + "`n" + $override)
$supervisor = $supervisor.Replace('$PSScriptRoot', ("'" + $PSScriptRoot.Replace("'", "''") + "'"))
& ([scriptblock]::Create($supervisor)) -EnvironmentId 'win-x64-dev-oracle-2026-07' -NoMatrixUpdate -SafeMode -DiagnosticCaseId 'windows-scalar-interop' -RunId $RunId -OutputRoot $OutputRoot -TimeoutSeconds 180
