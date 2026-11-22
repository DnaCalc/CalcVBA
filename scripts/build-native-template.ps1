param([string]$OutDir = (Join-Path $PSScriptRoot '../assets/native'))
$ErrorActionPreference = 'Stop'
$PSNativeCommandUseErrorActionPreference = $true
# Developer/release infrastructure only. A user's native build reads the
# finished image; it never invokes this script, cargo or a system linker.
Push-Location (Join-Path $PSScriptRoot '..')
try {
    New-Item -ItemType Directory -Force $OutDir | Out-Null
    if ($IsLinux) {
        cargo build -p calcvba-native-template --release --lib --bin calcvba-native-exe-template
        cargo run -p calcvba-pe --release -- finish-elf target/release/libcalcvba_native_template.so (Join-Path $OutDir 'calcvba-linux-x64.so')
        cargo run -p calcvba-pe --release -- finish-elf target/release/calcvba-native-exe-template (Join-Path $OutDir 'calcvba-linux-x64')
    }
    elseif ($IsWindows) {
        cargo build -p calcvba-native-template --release
        cargo run -p calcvba-pe --release -- finish target/release/calcvba_native_template.dll (Join-Path $OutDir 'calcvba-native-x64.dll')
        cargo run -p calcvba-pe --release -- finish target/release/calcvba-native-exe-template.exe (Join-Path $OutDir 'calcvba-native-x64.exe')
        cargo run -p calcvba-pe --release -- finish target/release/calcvba-native-com-exe-template.exe (Join-Path $OutDir 'calcvba-native-com-x64.exe')
    }
    else {
        throw 'Native templates require Windows x64 or GNU/Linux x86-64'
    }
}
finally { Pop-Location }
