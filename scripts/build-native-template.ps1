param([string]$OutDir = (Join-Path $PSScriptRoot '../assets/native'))
$ErrorActionPreference = 'Stop'
$PSNativeCommandUseErrorActionPreference = $true
# Developer/release infrastructure only. A user's native build reads the
# finished image; it never invokes this script, cargo or a system linker.
Push-Location (Join-Path $PSScriptRoot '..')
try {
    cargo build -p calcvba-native-template --release
    New-Item -ItemType Directory -Force $OutDir | Out-Null
    cargo run -p calcvba-pe --release -- finish target/release/calcvba_native_template.dll (Join-Path $OutDir 'calcvba-native-x64.dll')
    cargo run -p calcvba-pe --release -- finish target/release/calcvba-native-exe-template.exe (Join-Path $OutDir 'calcvba-native-x64.exe')
}
finally { Pop-Location }
