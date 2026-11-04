param(
    [string]$OutDir = 'temp/native-xll-clients',
    [string]$CliPath = '',
    [string]$OracleSession = '',
    [string]$DnaCalcPath = '',
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
    if ($OracleSession) {
        $OracleSession=(Resolve-Path -LiteralPath $OracleSession).Path
        $ready=Get-Content -Raw -LiteralPath (Join-Path $OracleSession 'ready.json') | ConvertFrom-Json
        if ($ready.launches -ne 1 -or $ready.pid -eq 3144) { throw 'Invalid owned oracle session' }
    }
    New-Item -ItemType Directory -Force $OutDir | Out-Null
    $OutDir = (Resolve-Path -LiteralPath $OutDir).Path
    & $CliPath build examples/native/xll/PlainXll.basproj --target NativeLibrary --out-dir $OutDir
    $dll = Join-Path $OutDir 'PlainXll.dll'
    $support = Join-Path $repoRoot 'crates/calcvba-build/tests/support'
    $expected = '5.75|3.75|9|16385|1|0'
    $vswhere = Join-Path ${env:ProgramFiles(x86)} 'Microsoft Visual Studio/Installer/vswhere.exe'
    $vs = & $vswhere -latest -products '*' -requires Microsoft.VisualStudio.Component.VC.Tools.x86.x64 -property installationPath
    if (-not $vs) { throw 'The C comparison needs MSVC Build Tools' }
    $vc = (Get-ChildItem -LiteralPath (Join-Path $vs 'VC/Tools/MSVC') -Directory | Sort-Object Name -Descending | Select-Object -First 1).FullName
    $sdk = Join-Path ${env:ProgramFiles(x86)} 'Windows Kits/10'
    $sdkVersion = (Get-ChildItem -LiteralPath (Join-Path $sdk 'Include') -Directory | Where-Object { Test-Path (Join-Path $_.FullName 'ucrt/stdio.h') } | Sort-Object Name -Descending | Select-Object -First 1).Name
    $richExpected = (Get-Content -Raw -LiteralPath (Join-Path $support 'native_xll_rich_expected.txt')).Trim()
    foreach ($case in @(
        @{c='native_xll_client.c';vba='native_xll_client.bas';expected=$expected;name='numeric'},
        @{c='native_xll_rich_client.c';vba='native_xll_rich_client.bas';expected=$richExpected;name='rich'}
    )) {
        $client = Join-Path $OutDir ($case.name+'-client.exe')
        & (Join-Path $vc 'bin/Hostx64/x64/cl.exe') /nologo /W4 /WX /wd4191 /O2 /utf-8 (Join-Path $support $case.c) "/Fo:$(Join-Path $OutDir ($case.name+'-client.obj'))" "/Fe:$client" "/I$vc/include" "/I$sdk/Include/$sdkVersion/um" "/I$sdk/Include/$sdkVersion/shared" "/I$sdk/Include/$sdkVersion/ucrt" /link oleaut32.lib "/LIBPATH:$vc/lib/x64" "/LIBPATH:$sdk/Lib/$sdkVersion/um/x64" "/LIBPATH:$sdk/Lib/$sdkVersion/ucrt/x64"
        $actual = (& $client $dll).Trim()
        if ($actual -cne $case.expected) { throw "C client mismatch: $actual" }
        Write-Output "C/$($case.name):$actual"
        & $client $dll --bench | Set-Content -LiteralPath (Join-Path $OutDir ($case.name+'-bench.txt'))
        $sourcePath = Join-Path $OutDir ($case.name+'-client.bas')
        $source = (Get-Content -LiteralPath (Join-Path $support $case.vba) -Raw).Replace('__DLL__', $dll)
        Set-Content -LiteralPath $sourcePath $source -Encoding utf8NoBOM
        foreach ($backend in @('vm3','jit')) {
            $values = & $CliPath run $sourcePath --backend $backend --profile windows-headless --dump-values
            if (-not ($values -match [regex]::Escape('string:"' + $case.expected + '"'))) { throw "CalcVBA $backend mismatch: $values" }
            Write-Output "CalcVBA/$($case.name)/${backend}:$($case.expected)"
        }
        if ($OracleSession) {
            $jobId='xll_'+[Guid]::NewGuid().ToString('N')
            @{id=$jobId;source_path=$sourcePath;expected=$case.expected} | ConvertTo-Json |
                Set-Content -LiteralPath (Join-Path $OracleSession ('jobs/'+$jobId+'.json')) -Encoding utf8NoBOM
            $resultPath=Join-Path $OracleSession ('results/'+$jobId+'.json')
            $deadline=[DateTime]::UtcNow.AddSeconds(180)
            while (-not (Test-Path -LiteralPath $resultPath)) {
                if ([DateTime]::UtcNow -gt $deadline) { throw 'Owned Excel comparison timed out' }
                Start-Sleep -Milliseconds 100
            }
            $result=Get-Content -Raw -LiteralPath $resultPath | ConvertFrom-Json
            if ($result.compile_status -ne 'ok' -or $result.run_status -ne 'ok' -or $result.value -cne $case.expected -or $result.error) { throw ('Excel comparison failed: '+($result | ConvertTo-Json -Compress)) }
            Write-Output "Excel/VBA/$($case.name):$($case.expected)"
        }
    }
    $directExpected = (Get-Content -Raw -LiteralPath (Join-Path $support 'native_xll_value_expected.txt')).Trim()
    $libraryPath = Join-Path $repoRoot 'examples/native/xll/XlValues.bas'
    $probePath = Join-Path $support 'native_xll_value_probe.bas'
    $direct = ((Get-Content -Raw -LiteralPath $libraryPath) -replace '(?m)^Attribute .*\r?\n','') + "`n" + ((Get-Content -Raw -LiteralPath $probePath) -replace '(?m)^Option Explicit\r?\n','')
    $directPath = Join-Path $OutDir 'direct-values.bas'
    Set-Content -LiteralPath $directPath $direct -Encoding utf8NoBOM
    foreach ($backend in @('vm3','jit')) {
        $values = & $CliPath run $directPath --backend $backend --profile windows-headless --dump-values
        if (-not ($values -match [regex]::Escape('string:"' + $directExpected + '"'))) { throw "Direct library $backend mismatch: $values" }
        Write-Output "CalcVBA/direct/${backend}:$directExpected"
    }
    if ($OracleSession) {
        $modulesPath = Join-Path $OutDir 'direct-modules.json'
        @(@{name='XlValues';source_path=$libraryPath}, @{name='OracleSelfTest';source_path=$probePath}) |
            ConvertTo-Json | Set-Content -LiteralPath $modulesPath -Encoding utf8NoBOM
        $projectDir = Join-Path $OutDir ('direct-excel-'+[Guid]::NewGuid().ToString('N'))
        & (Join-Path $PSHOME 'pwsh.exe') -NoLogo -NoProfile -STA -File (Join-Path $PSScriptRoot 'native-export-oracle-project.ps1') -OracleSession $OracleSession -Directory $projectDir -ModulesFile $modulesPath
        $directResult = Get-Content -Raw -LiteralPath (Join-Path $projectDir 'result.json') | ConvertFrom-Json
        if ($directResult.value -cne $directExpected -or $directResult.error) { throw 'Direct Excel library comparison failed' }
        Write-Output "Excel/direct:$directExpected"
        $xll = Join-Path $OutDir 'PlainXll.xll'
        Copy-Item -LiteralPath $dll -Destination $xll
        $loadSource = (Get-Content -LiteralPath (Join-Path $support 'native_xll_excel.bas') -Raw).Replace('__XLL__', $xll)
        $cases = @(Get-Content -Raw -LiteralPath (Join-Path $support 'native_xll_cases.json') | ConvertFrom-Json)
        $vbaCases = [Collections.Generic.List[string]]::new()
        $vbaCases.Add('ReDim formulas(0 To '+($cases.Count-1)+'): ReDim expected(0 To '+($cases.Count-1)+')')
        for ($i=0; $i -lt $cases.Count; $i++) {
            $formula = ([string]$cases[$i].formula).Replace('"','""')
            $value = $cases[$i].value
            if ($value -is [string]) { $literal='"'+$value.Replace('"','""')+'"' }
            elseif ($value -is [bool]) { $literal=if($value){'True'}else{'False'} }
            elseif ($value.PSObject.Properties.Name -contains 'error') { $literal='CVErr(2042&)' }
            else { $literal=([double]$value).ToString('R',[Globalization.CultureInfo]::InvariantCulture)+'#' }
            $vbaCases.Add('formulas('+$i+') = "'+$formula+'": expected('+$i+') = '+$literal)
        }
        $loadSource=$loadSource.Replace("'__CASES__",($vbaCases -join "`n    "))
        $loadPath = Join-Path $OutDir 'excel-load.bas'
        Set-Content -LiteralPath $loadPath $loadSource -Encoding utf8NoBOM
        $jobId = 'xll_' + [Guid]::NewGuid().ToString('N')
        @{id=$jobId;source_path=$loadPath} | ConvertTo-Json |
            Set-Content -LiteralPath (Join-Path $OracleSession ('jobs/'+$jobId+'.json')) -Encoding utf8NoBOM
        $resultPath = Join-Path $OracleSession ('results/'+$jobId+'.json')
        $deadline = [DateTime]::UtcNow.AddSeconds(180)
        while (-not (Test-Path -LiteralPath $resultPath)) {
            if ([DateTime]::UtcNow -gt $deadline) { throw 'Owned Excel XLL check timed out' }
            Start-Sleep -Milliseconds 100
        }
        $result = Get-Content -Raw -LiteralPath $resultPath | ConvertFrom-Json
        if ($result.compile_status -ne 'ok' -or $result.run_status -ne 'ok' -or $result.error -or
            -not ([string]$result.value).StartsWith('5.75|9|0|Declare_us=') -or -not ([string]$result.value).EndsWith('|cases='+$cases.Count+'|buffers=0|closed=1')) {
            throw ('Excel XLL check failed: ' + ($result | ConvertTo-Json -Compress))
        }
        Write-Output "Excel/XLL:$($result.value)"
    }

    if ($DnaCalcPath) {
        $DnaCalcPath=(Resolve-Path -LiteralPath $DnaCalcPath).Path
        $hostDirectory=Join-Path $OutDir 'dna-host'
        New-Item -ItemType Directory -Path $hostDirectory -Force | Out-Null
        $hostExecutable=Join-Path $hostDirectory 'dna-calc-desktop.exe'
        Copy-Item -LiteralPath $DnaCalcPath -Destination $hostExecutable
        Copy-Item -LiteralPath (Join-Path (Split-Path $DnaCalcPath) 'xlcall32.dll') -Destination (Join-Path $hostDirectory 'XLCALL32.DLL')
        $xll=Join-Path $OutDir 'PlainXll.xll'
        if (-not (Test-Path -LiteralPath $xll)) { Copy-Item -LiteralPath $dll -Destination $xll }
        $hostOutput=Join-Path $hostDirectory 'result.json'
        $hostError=Join-Path $hostDirectory 'error.txt'
        # DNA Calc's current desktop harness reports MissingInput for blank A2
        # before invoking the XLL. Retain those Excel checks and share the rest.
        $cases=@(Get-Content -Raw -LiteralPath (Join-Path $support 'native_xll_cases.json') | ConvertFrom-Json | Where-Object { $_.host -ne 'excel' })
        $hostCases=Join-Path $hostDirectory 'cases.json'
        $cases | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath $hostCases -Encoding utf8NoBOM
        $hostProcess=Start-Process $hostExecutable -ArgumentList @('--xll',('"'+$xll+'"'),'--cases',('"'+$hostCases+'"')) -WindowStyle Hidden -PassThru -RedirectStandardOutput $hostOutput -RedirectStandardError $hostError
        @{pid=$hostProcess.Id;path=$hostExecutable;start_utc=$hostProcess.StartTime.ToUniversalTime().ToString('o')} | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $hostDirectory 'owned-process.json')
        if (-not $hostProcess.WaitForExit(60000)) {
            $live=Get-Process -Id $hostProcess.Id -ErrorAction SilentlyContinue
            if ($live -and $live.StartTime -eq $hostProcess.StartTime -and $live.Path -eq $hostProcess.Path) { Stop-Process -InputObject $live }
            throw 'DNA Calc XLL comparison timed out'
        }
        if ($hostProcess.ExitCode -ne 0) { throw ('DNA Calc: '+(Get-Content -Raw -LiteralPath $hostError)) }
        $result=Get-Content -Raw -LiteralPath $hostOutput | ConvertFrom-Json
        if ($result.results.Count -ne $cases.Count -or $result.live_callback_allocations -ne 0) { throw 'DNA Calc result count or callback ownership mismatch' }
        Write-Output "DNA Calc/XLL:$($cases.Count) matching calls; callback allocations zero"
    }
}
finally { Pop-Location }
