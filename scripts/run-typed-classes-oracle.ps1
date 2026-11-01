param([string]$RunId=('typed_classes_{0:yyyyMMddTHHmmssZ}' -f [DateTime]::UtcNow),[string]$OutputRoot='temp/calcvba11/oracle')
$ErrorActionPreference='Stop'
$script:aggregateSource=Get-Content -Raw -LiteralPath (Join-Path $PSScriptRoot '../docs/evidence/programs/ideal-2026-07/windows-x64/WIN-4/typed-classes-api-run11.bas')
$script:classSource=Get-Content -Raw -LiteralPath (Join-Path $PSScriptRoot '../docs/evidence/programs/ideal-2026-07/windows-x64/WIN-4/typed-classes-widget-run11.cls')
$script:aggregateExpected='same=True|start=40|next=42|distinct=False|journal=IIT42T40'
$supervisor = Get-Content -Raw -LiteralPath (Join-Path $PSScriptRoot 'run-excel-vba-oracle.ps1')
$anchor = '. (Join-Path $PSScriptRoot "excel-vba-oracle-bootstrap.ps1")'
$override = @'
$script:aggregateDescriptor = @(Get-ExcelOracleHarnessCases | Where-Object id -eq 'windows-scalar-interop')[0]
$script:aggregateDescriptor.module_source = $script:aggregateSource
$script:aggregateDescriptor.module_name = 'Api'
$script:aggregateDescriptor.run_procedure = 'Api.RunProbe'
$script:aggregateDescriptor.expected_value = $script:aggregateExpected
$script:aggregateDescriptor.purpose = 'Typed source class development observation only'
$script:aggregateDescriptor | Add-Member -NotePropertyName additional_modules -NotePropertyValue @([pscustomobject]@{ name='Widget'; kind=2; source=$script:classSource })
function Get-ExcelOracleHarnessCases { return @($script:aggregateDescriptor) }
'@
if (-not $supervisor.Contains($anchor)) { throw 'Supervisor import anchor changed' }
$supervisor = $supervisor.Replace($anchor, $anchor + "`n" + $override)
$supervisor = $supervisor.Replace('$PSScriptRoot', ("'" + $PSScriptRoot.Replace("'", "''") + "'"))
& ([scriptblock]::Create($supervisor)) -EnvironmentId 'win-x64-dev-oracle-2026-07' -NoMatrixUpdate -SafeMode -DiagnosticCaseId 'windows-scalar-interop' -RunId $RunId -OutputRoot $OutputRoot -TimeoutSeconds 180
