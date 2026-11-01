param(
    [string]$RunId = ('typed_defaults_{0:yyyyMMddTHHmmssZ}' -f [DateTime]::UtcNow),
    [string]$OutputRoot = 'temp/calcvba10/oracle'
)
$ErrorActionPreference = 'Stop'
$script:defaultsSource = Get-Content -Raw -LiteralPath (Join-Path $PSScriptRoot '../docs/evidence/programs/ideal-2026-07/windows-x64/WIN-4/typed-defaults-run10.bas')
$script:defaultsExpected = 'add=42|text=v=41|number=41|equal=True|condition=1|identity=True|nothing=91|null=94|bad=13|required-default=450'
$supervisor = Get-Content -Raw -LiteralPath (Join-Path $PSScriptRoot 'run-excel-vba-oracle.ps1')
$anchor = '. (Join-Path $PSScriptRoot "excel-vba-oracle-bootstrap.ps1")'
$override = @'
$script:defaultsDescriptor = @(Get-ExcelOracleHarnessCases | Where-Object id -eq 'windows-scalar-interop')[0]
$script:defaultsDescriptor.module_source = $script:defaultsSource
$script:defaultsDescriptor.expected_value = $script:defaultsExpected
$script:defaultsDescriptor.purpose = 'Runtime object default coercions with retained identity; development observation only'
function Get-ExcelOracleHarnessCases { return @($script:defaultsDescriptor) }
'@
if (-not $supervisor.Contains($anchor)) { throw 'Supervisor import anchor changed' }
$supervisor = $supervisor.Replace($anchor, $anchor + "`n" + $override)
$supervisor = $supervisor.Replace('$PSScriptRoot', ("'" + $PSScriptRoot.Replace("'", "''") + "'"))
& ([scriptblock]::Create($supervisor)) -EnvironmentId 'win-x64-dev-oracle-2026-07' -NoMatrixUpdate -SafeMode -DiagnosticCaseId 'windows-scalar-interop' -RunId $RunId -OutputRoot $OutputRoot -TimeoutSeconds 180
