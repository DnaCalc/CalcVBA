param([string]$OutDir = (Join-Path $PSScriptRoot '../assets/native'))
$ErrorActionPreference = 'Stop'
$PSNativeCommandUseErrorActionPreference = $true
# Developer/release infrastructure only. A user's NativeLibrary build reads the
# finished DLL; it never invokes this script, cargo or a system linker.
Push-Location (Join-Path $PSScriptRoot '..')
try {
    cargo build -p oxvba-native-template --release
    New-Item -ItemType Directory -Force $OutDir | Out-Null
    cargo run -p oxvba-pe --release -- finish target/release/oxvba_native_template.dll (Join-Path $OutDir 'calcvba-native-x64.dll')
}
finally { Pop-Location }
