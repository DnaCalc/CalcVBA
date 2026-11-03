$ErrorActionPreference = "Stop"
$PSNativeCommandUseErrorActionPreference = $true

Push-Location (Join-Path $PSScriptRoot "..")
try {
    $sample = "conformance/tests/smoke.bas"
    if (-not (Test-Path $sample)) {
        throw "Missing smoke source: $sample"
    }

    cargo run -q -p calcvba-cli --bin calcvba-cli -- run $sample --dump-values
    Write-Host "smoke run: ok"
}
finally {
    Pop-Location
}
