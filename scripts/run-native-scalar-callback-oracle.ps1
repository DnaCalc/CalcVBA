param(
    [string]$RunId = 'calcvba13-native-scalar-callback',
    [string]$OutputRoot = 'temp/oracle'
)
$ErrorActionPreference = 'Stop'
$root = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
$sourceRoot = Join-Path $root 'temp/oracle/programs/ideal-2026-07/core/CORE-7'
$buildRoot = Join-Path $root 'temp/calcvba13/native-scalar-callback'
[IO.Directory]::CreateDirectory($buildRoot) | Out-Null
$library = Join-Path $buildRoot 'native-scalar-callback.dll'
& rustc --crate-type cdylib --edition 2024 (Join-Path $sourceRoot 'native-scalar-callback-run13.rs') -o $library
if ($LASTEXITCODE -ne 0) { throw 'Native callback oracle DLL build failed' }
$receiptPath = Join-Path $buildRoot 'build.json'
[ordered]@{ schema='calcvba.native-scalar-oracle-build.v1'; created_utc=[DateTime]::UtcNow.ToString('o'); dll=$library; dll_sha256=(Get-FileHash -LiteralPath $library -Algorithm SHA256).Hash.ToLowerInvariant(); source_sha256=(Get-FileHash -LiteralPath (Join-Path $sourceRoot 'native-scalar-callback-run13.rs') -Algorithm SHA256).Hash.ToLowerInvariant(); callbacks_sha256=(Get-FileHash -LiteralPath (Join-Path $sourceRoot 'native-scalar-callback-bodies-run13.bas') -Algorithm SHA256).Hash.ToLowerInvariant(); rustc=((& rustc -Vv) -join [Environment]::NewLine) } | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $receiptPath -Encoding utf8NoBOM
$script:aliasSource = (Get-Content -Raw -LiteralPath (Join-Path $sourceRoot 'native-scalar-callback-run13.bas')).Replace('__LIBRARY__', $library)
$script:aliasSource = $script:aliasSource.Replace('__CALLBACK_BODIES__', (Get-Content -Raw -LiteralPath (Join-Path $sourceRoot 'native-scalar-callback-bodies-run13.bas')))
$supervisor = Get-Content -Raw -LiteralPath (Join-Path $PSScriptRoot 'run-excel-vba-oracle.ps1')
$anchor = '. (Join-Path $PSScriptRoot "excel-vba-oracle-bootstrap.ps1")'
$override = @'
$script:aliasDescriptor = @(Get-ExcelOracleHarnessCases | Where-Object id -eq 'windows-scalar-interop')[0]
$script:aliasDescriptor.module_source = $script:aliasSource
$script:aliasDescriptor.expected_value = 'double=12.5:12.5;boolean=False:False;not=True:False'
$script:aliasDescriptor.purpose = 'Native exact Double/Boolean ByRef visibility and repeated aliases during callback'
function Get-ExcelOracleHarnessCases { return @($script:aliasDescriptor) }
'@
if (-not $supervisor.Contains($anchor)) { throw 'Supervisor import anchor changed' }
$supervisor = $supervisor.Replace($anchor, $anchor + "`n" + $override)
$supervisor = $supervisor.Replace('$PSScriptRoot', ("'" + $PSScriptRoot.Replace("'", "''") + "'"))
& ([scriptblock]::Create($supervisor)) -SafeMode -DiagnosticCaseId 'windows-scalar-interop' -RunId $RunId -OutputRoot $OutputRoot -TimeoutSeconds 180

Copy-Item -LiteralPath $receiptPath -Destination (Join-Path (Join-Path $OutputRoot $RunId) 'native-build.json')
