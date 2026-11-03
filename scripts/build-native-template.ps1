param([string]$OutDir = (Join-Path $PSScriptRoot '../assets/native'))
$ErrorActionPreference = 'Stop'
$PSNativeCommandUseErrorActionPreference = $true
# Developer/release infrastructure only. A user's NativeLibrary build reads the
# finished DLL; it never invokes this script, cargo or a system linker.
Push-Location (Join-Path $PSScriptRoot '..')
try {
    cargo build -p calcvba-native-template --release
    New-Item -ItemType Directory -Force $OutDir | Out-Null
    cargo run -p calcvba-pe --release -- finish target/release/calcvba_native_template.dll (Join-Path $OutDir 'calcvba-native-x64.dll')
}
finally { Pop-Location }
