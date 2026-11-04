param(
    [string]$OutDir = 'temp/native-record-clients',
    [string]$CliPath = '',
    [string]$OracleSession = '',
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
        cargo build -p calcvba-cli
        $CliPath = Join-Path $repoRoot 'target/debug/calcvba-cli.exe'
    }
    $CliPath = (Resolve-Path -LiteralPath $CliPath).Path
    New-Item -ItemType Directory -Force $OutDir | Out-Null
    $OutDir = (Resolve-Path -LiteralPath $OutDir).Path
    & $CliPath build examples/native/records/NativeRecords.basproj --target NativeLibrary --out-dir $OutDir
    $dll = Join-Path $OutDir 'NativeRecords.dll'
    $support = Join-Path $repoRoot 'crates/calcvba-build/tests/support'
    $expected = (Get-Content -LiteralPath (Join-Path $support 'native_record_expected.txt') -Raw).Trim()
    $vswhere = Join-Path ${env:ProgramFiles(x86)} 'Microsoft Visual Studio/Installer/vswhere.exe'
    $vs = & $vswhere -latest -products '*' -requires Microsoft.VisualStudio.Component.VC.Tools.x86.x64 -property installationPath
    if (-not $vs) { throw 'The C comparison needs MSVC Build Tools' }
    $vc = (Get-ChildItem -LiteralPath (Join-Path $vs 'VC/Tools/MSVC') -Directory | Sort-Object Name -Descending | Select-Object -First 1).FullName
    $sdk = Join-Path ${env:ProgramFiles(x86)} 'Windows Kits/10'
    $sdkVersion = (Get-ChildItem -LiteralPath (Join-Path $sdk 'Include') -Directory | Where-Object { Test-Path (Join-Path $_.FullName 'ucrt/stdio.h') } | Sort-Object Name -Descending | Select-Object -First 1).Name
    $client = Join-Path $OutDir 'native-client.exe'
    & (Join-Path $vc 'bin/Hostx64/x64/cl.exe') /nologo /W4 /WX /wd4191 /O2 /utf-8 (Join-Path $support 'native_record_client.c') "/Fo:$(Join-Path $OutDir 'native-client.obj')" "/Fe:$client" "/I$vc/include" "/I$sdk/Include/$sdkVersion/um" "/I$sdk/Include/$sdkVersion/shared" "/I$sdk/Include/$sdkVersion/ucrt" /link oleaut32.lib "/LIBPATH:$vc/lib/x64" "/LIBPATH:$sdk/Lib/$sdkVersion/um/x64" "/LIBPATH:$sdk/Lib/$sdkVersion/ucrt/x64"
    $actual = (& $client $dll).Trim()
    if ($actual -cne $expected) { throw "C client mismatch: $actual" }
    Write-Output "C:$actual"
    $script:nativeSource = (Get-Content -LiteralPath (Join-Path $support 'native_record_client.bas') -Raw).Replace('__DLL__', $dll)
    $sourcePath = Join-Path $OutDir 'native-client.bas'
    Set-Content -LiteralPath $sourcePath $script:nativeSource -Encoding utf8NoBOM
    $direct = (Get-Content -Raw examples/native/records/Records.bas) -replace '(?m)^Attribute .*\r?\n', ''
    $direct = $direct.Replace('Public Sub Main()', "Public Sub Main()`n    Dim observed As String`n    observed = FixedFieldSuite()")
    $direct += "`nPublic Function ProbeNativeExports() As String`n    ProbeNativeExports = FixedFieldSuite()`nEnd Function`n"
    $directPath = Join-Path $OutDir 'direct.bas'
    Set-Content -LiteralPath $directPath $direct -Encoding utf8NoBOM
    $directExpected = '29:41:left:42:abc:q  :record'
    foreach ($backend in @('vm3','jit')) {
        $values = & $CliPath run $sourcePath --backend $backend --profile windows-headless --dump-values
        if (-not ($values -match [regex]::Escape('string:"' + $expected + '"'))) { throw "CalcVBA $backend mismatch: $values" }
        Write-Output "CalcVBA/${backend}:$expected"
        $values = & $CliPath run $directPath --backend $backend --profile windows-headless --dump-values
        if (-not ($values -match [regex]::Escape('string:"' + $directExpected + '"'))) { throw "CalcVBA direct $backend mismatch: $values" }
        Write-Output "CalcVBA/direct/${backend}:$directExpected"
    }
    if ($OracleSession) {
        $OracleSession=(Resolve-Path -LiteralPath $OracleSession).Path
        $ready=Get-Content -Raw -LiteralPath (Join-Path $OracleSession 'ready.json') | ConvertFrom-Json
        if ($ready.launches -ne 1 -or $ready.pid -eq 3144) { throw 'Invalid owned oracle session' }
        foreach ($case in @(@{path=$sourcePath;expected=$expected}, @{path=$directPath;expected=$directExpected})) {
            $jobId='interop_'+[Guid]::NewGuid().ToString('N')
            [ordered]@{id=$jobId;source_path=$case.path;expected=$case.expected} | ConvertTo-Json |
                Set-Content -LiteralPath (Join-Path $OracleSession ('jobs/'+$jobId+'.json')) -Encoding utf8NoBOM
            $resultPath=Join-Path $OracleSession ('results/'+$jobId+'.json')
            $deadline=[DateTime]::UtcNow.AddSeconds(180)
            while (-not (Test-Path -LiteralPath $resultPath)) {
                if ([DateTime]::UtcNow -gt $deadline) { throw 'Owned Excel comparison timed out' }
                Start-Sleep -Milliseconds 100
            }
            $result=Get-Content -Raw -LiteralPath $resultPath | ConvertFrom-Json
            if ($result.compile_status -ne 'ok' -or $result.run_status -ne 'ok' -or $result.value -cne $case.expected -or $result.error) {
                throw ('Excel comparison failed: '+($result | ConvertTo-Json -Compress))
            }
            Write-Output "Excel/VBA:$($case.expected)"
        }
    }

}
finally { Pop-Location }
