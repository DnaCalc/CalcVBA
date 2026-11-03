param(
    [string]$RunId = ('typed_continuation_{0:yyyyMMddTHHmmssZ}' -f [DateTime]::UtcNow),
    [string]$OutputRoot = 'temp/oracle'
)
$ErrorActionPreference = 'Stop'
# The existing supervisor owns Excel and arms its PID-scoped UIA guardian before
# explicit VBE compile command 578. It preserves security settings and registration.
$sourcePath = Join-Path $PSScriptRoot '../crates/calcvba-host/tests/support/oracle/typed-continuation.bas'
$script:continuationSource = Get-Content -Raw -LiteralPath $sourcePath
$script:continuationExpected = 'alias=10:0,1,2,|catch=27:0,7,8,2,|retry=43:7,8,7,|skip=14:7,2,|object=42:$B$1'
$supervisor = Get-Content -Raw -LiteralPath (Join-Path $PSScriptRoot 'run-excel-vba-oracle.ps1')
$anchor = '. (Join-Path $PSScriptRoot "excel-vba-oracle-bootstrap.ps1")'
$override = @'
$script:continuationDescriptor = @(Get-ExcelOracleHarnessCases | Where-Object id -eq 'windows-scalar-interop')[0]
$script:continuationDescriptor.module_source = $script:continuationSource
$script:continuationDescriptor.expected_value = $script:continuationExpected
$script:continuationDescriptor.purpose = 'Nested ByRef aliases, propagation and explicit Resume semantics; development observation only'
function Get-ExcelOracleHarnessCases { return @($script:continuationDescriptor) }
'@
if (-not $supervisor.Contains($anchor)) { throw 'Supervisor import anchor changed' }
$supervisor = $supervisor.Replace($anchor, $anchor + "`n" + $override)
$supervisor = $supervisor.Replace('$PSScriptRoot', ("'" + $PSScriptRoot.Replace("'", "''") + "'"))
& ([scriptblock]::Create($supervisor)) -SafeMode -DiagnosticCaseId 'windows-scalar-interop' -RunId $RunId -OutputRoot $OutputRoot -TimeoutSeconds 180
