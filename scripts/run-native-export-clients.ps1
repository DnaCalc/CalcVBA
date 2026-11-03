param(
    [string]$OutDir = 'temp/native-export-clients',
    [string]$CliPath = '',
    [switch]$Excel,
    [string]$RunId = ('native_exports_' + [Guid]::NewGuid().ToString('N'))
)
$ErrorActionPreference = 'Stop'
$PSNativeCommandUseErrorActionPreference = $true
$repoRoot = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
Push-Location $repoRoot
try {
    # Developer comparison harness. Only this C harness needs MSVC; the product
    # NativeLibrary build below uses its embedded template and no system linker.
    if (-not $CliPath) {
        cargo build -p oxvba-cli
        $CliPath = Join-Path $repoRoot 'target/debug/oxvba-cli.exe'
    }
    $CliPath = (Resolve-Path -LiteralPath $CliPath).Path
    New-Item -ItemType Directory -Force $OutDir | Out-Null
    $OutDir = (Resolve-Path -LiteralPath $OutDir).Path
    & $CliPath build examples/native/scalars/NativeScalars.basproj --target NativeLibrary --out-dir $OutDir
    $dll = Join-Path $OutDir 'NativeScalars.dll'
    $support = Join-Path $repoRoot 'crates/oxvba-build/tests/support'
    $expected = (Get-Content -LiteralPath (Join-Path $support 'native_scalar_expected.txt') -Raw).Trim()
    $vswhere = Join-Path ${env:ProgramFiles(x86)} 'Microsoft Visual Studio/Installer/vswhere.exe'
    $vs = & $vswhere -latest -products '*' -requires Microsoft.VisualStudio.Component.VC.Tools.x86.x64 -property installationPath
    if (-not $vs) { throw 'The C comparison needs MSVC Build Tools' }
    $vc = (Get-ChildItem -LiteralPath (Join-Path $vs 'VC/Tools/MSVC') -Directory | Sort-Object Name -Descending | Select-Object -First 1).FullName
    $sdk = Join-Path ${env:ProgramFiles(x86)} 'Windows Kits/10'
    $sdkVersion = (Get-ChildItem -LiteralPath (Join-Path $sdk 'Include') -Directory | Where-Object { Test-Path (Join-Path $_.FullName 'ucrt/stdio.h') } | Sort-Object Name -Descending | Select-Object -First 1).Name
    $client = Join-Path $OutDir 'native-client.exe'
    & (Join-Path $vc 'bin/Hostx64/x64/cl.exe') /nologo /W4 /WX /wd4191 /O2 /utf-8 (Join-Path $support 'native_scalar_client.c') "/Fo:$(Join-Path $OutDir 'native-client.obj')" "/Fe:$client" "/I$vc/include" "/I$sdk/Include/$sdkVersion/um" "/I$sdk/Include/$sdkVersion/shared" "/I$sdk/Include/$sdkVersion/ucrt" /link "/LIBPATH:$vc/lib/x64" "/LIBPATH:$sdk/Lib/$sdkVersion/um/x64" "/LIBPATH:$sdk/Lib/$sdkVersion/ucrt/x64"
    $actual = (& $client $dll).Trim()
    if ($actual -cne $expected) { throw "C client mismatch: $actual" }
    Write-Output "C:$actual"
    $script:nativeSource = (Get-Content -LiteralPath (Join-Path $support 'native_scalar_client.bas') -Raw).Replace('__DLL__', $dll)
    $sourcePath = Join-Path $OutDir 'native-client.bas'
    Set-Content -LiteralPath $sourcePath $script:nativeSource -Encoding utf8NoBOM
    foreach ($backend in @('vm3','jit')) {
        $values = & $CliPath run $sourcePath --backend $backend --profile windows-headless --dump-values
        if (-not ($values -match [regex]::Escape('string:"' + $expected + '"'))) { throw "CalcVBA $backend mismatch: $values" }
        Write-Output "CalcVBA/${backend}:$expected"
    }
    if ($Excel) {
        # All scalar cases and both math helpers run in ONE owned Excel process,
        # one visible VBE compile, and one VBA function. No per-case launches.
        $script:nativeExpected = $expected
        $supervisor = Get-Content -Raw -LiteralPath (Join-Path $PSScriptRoot 'run-excel-vba-oracle.ps1')
        $anchor = '. (Join-Path $PSScriptRoot "excel-vba-oracle-bootstrap.ps1")'
        $override = @'
$script:nativeDescriptor = @(Get-ExcelOracleHarnessCases | Where-Object id -eq 'windows-scalar-interop')[0]
$script:nativeDescriptor.module_source = $script:nativeSource
$script:nativeDescriptor.run_procedure = 'OracleSelfTest.ProbeNativeExports'
$script:nativeDescriptor.macro_probe_target = $script:nativeDescriptor.run_procedure
$script:nativeDescriptor.expected_value = $script:nativeExpected
$script:nativeDescriptor.purpose = 'Named AOT scalars: casts, bitwise operators, integer faults, bounded recursion and 32-argument cross-module calls'
function Get-ExcelOracleHarnessCases { return @($script:nativeDescriptor) }
'@
        if (-not $supervisor.Contains($anchor)) { throw 'Excel supervisor bootstrap anchor changed' }
        $supervisor = $supervisor.Replace($anchor, $anchor + "`n" + $override)
        $supervisor = $supervisor.Replace('$PSScriptRoot', ("'" + $PSScriptRoot.Replace("'", "''") + "'"))
        & ([scriptblock]::Create($supervisor)) -SafeMode -DiagnosticCaseId 'windows-scalar-interop' -RunId $RunId -OutputRoot (Join-Path $OutDir 'oracle') -TimeoutSeconds 180
        # Targeted oracle probes capture observations even when expectations
        # differ. This comparison command must fail on a compile/value mismatch.
        $results = Get-Content -Raw -LiteralPath (Join-Path $OutDir "oracle/$RunId/results.json") | ConvertFrom-Json
        $cases = @($results.cases)
        if (-not $results.passed -or $cases.Count -ne 1 -or
            $cases[0].compile_status -ne 'ok' -or $cases[0].run_status -ne 'ok' -or
            [string]$cases[0].run_value -cne $expected) {
            throw 'Excel native export comparison failed; inspect the oracle results'
        }
        Write-Output "Excel/VBA:$expected"
    }
}
finally { Pop-Location }
