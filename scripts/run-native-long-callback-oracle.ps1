param(
    [string]$RunId = 'calcvba12-native-long-callback',
    [string]$OutputRoot = 'temp/calcvba12/oracle'
)
$ErrorActionPreference = 'Stop'
$root = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
$sourceRoot = Join-Path $root 'docs/evidence/programs/ideal-2026-07/core/CORE-7'
$buildRoot = Join-Path $root 'temp/calcvba12/native-long-callback'
[IO.Directory]::CreateDirectory($buildRoot) | Out-Null
$library = Join-Path $buildRoot 'native-long-callback.dll'
& rustc --crate-type cdylib --edition 2024 (Join-Path $sourceRoot 'native-long-callback-run12.rs') -o $library
if ($LASTEXITCODE -ne 0) { throw 'Native callback oracle DLL build failed' }
$script:aliasSource = (Get-Content -Raw -LiteralPath (Join-Path $sourceRoot 'native-long-callback-run12.bas')).Replace('__LIBRARY__', $library)
$supervisor = Get-Content -Raw -LiteralPath (Join-Path $PSScriptRoot 'run-excel-vba-oracle.ps1')
$anchor = '. (Join-Path $PSScriptRoot "excel-vba-oracle-bootstrap.ps1")'
$override = @'
$script:aliasDescriptor = @(Get-ExcelOracleHarnessCases | Where-Object id -eq 'windows-scalar-interop')[0]
$script:aliasDescriptor.module_source = $script:aliasSource
$script:aliasDescriptor.expected_value = '13:13'
$script:aliasDescriptor.purpose = 'Native exact-Long ByRef global visibility during synchronous callback'
function Get-ExcelOracleHarnessCases { return @($script:aliasDescriptor) }
'@
if (-not $supervisor.Contains($anchor)) { throw 'Supervisor import anchor changed' }
$supervisor = $supervisor.Replace($anchor, $anchor + "`n" + $override)
$supervisor = $supervisor.Replace('$PSScriptRoot', ("'" + $PSScriptRoot.Replace("'", "''") + "'"))
& ([scriptblock]::Create($supervisor)) -EnvironmentId 'win-x64-dev-oracle-2026-07' -NoMatrixUpdate -SafeMode -DiagnosticCaseId 'windows-scalar-interop' -RunId $RunId -OutputRoot $OutputRoot -TimeoutSeconds 180
