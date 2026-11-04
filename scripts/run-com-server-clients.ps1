param(
    [Parameter(Mandatory=$true)][string]$OracleSession,
    [string]$OutDir='temp/com-server-clients',
    [string]$CliPath='target/debug/calcvba-cli.exe'
)
$ErrorActionPreference='Stop'
$PSNativeCommandUseErrorActionPreference=$true
$repoRoot=(Resolve-Path (Join-Path $PSScriptRoot '..')).Path
Push-Location $repoRoot
try {
    $CliPath=(Resolve-Path -LiteralPath $CliPath).Path
    $OutDir=[IO.Path]::GetFullPath($OutDir)
    if(Test-Path -LiteralPath $OutDir){throw 'Use a new COM comparison directory'}
    [void][IO.Directory]::CreateDirectory($OutDir)
    $projectName='CalcVbaTest'+[guid]::NewGuid().ToString('N').Substring(0,16)
    $example=Join-Path $repoRoot 'examples/comserver'
    $support=Join-Path $repoRoot 'crates/calcvba-build/tests/support'
    $project=(Get-Content -Raw (Join-Path $example 'ComServer.basproj')).Replace('CalcVbaComExample',$projectName)
    foreach($name in @('Calculator','Counter','PublicItem','PrivateItem','Globals')) {
        $project=$project.Replace('Include="'+$name+'.cls"','Include="'+(Join-Path $example ($name+'.cls'))+'"')
    }
    $projectPath=Join-Path $OutDir 'Private.basproj'
    Set-Content -LiteralPath $projectPath $project -Encoding utf8NoBOM
    & $CliPath build $projectPath --target ComServer --out-dir $OutDir
    $descriptor=Get-Content -Raw (Join-Path $OutDir ($projectName+'.comserver.json')) | ConvertFrom-Json
    if($descriptor.classes.Count -ne 4 -or @($descriptor.classes | Where-Object {$_.implemented_interfaces.Count}).Count){throw 'Unexpected private test descriptor'}
    $dll=Join-Path $OutDir ($projectName+'.dll')
    $tlb=Join-Path $OutDir ($projectName+'.tlb')
    Add-Type -Path (Join-Path $support 'private_com_probe.cs')
    [CalcVbaPrivateComProbe]::Load($dll)
    $flags=[CalcVbaPrivateComProbe]::TypeFlags($tlb)
    if(($flags['Globals'] -band 1) -eq 0 -or ($flags['PublicItem'] -band 2) -ne 0 -or ($flags['Calculator'] -band 2) -eq 0 -or $flags.ContainsKey('PrivateItem')){throw 'Instancing type-library flags differ'}
    [CalcVbaPrivateComProbe]::Factories([string[]]$descriptor.classes.clsid,[bool[]]$descriptor.classes.creatable)
    Write-Output 'COM/factory:activation, identity, aggregation, locks and release pass'
    $expected='instances|ByRef|Variant|matrix|visibility|dual|global|events|error=457'
    $directClient=Join-Path $OutDir 'direct-client.bas'
    $directEvent=Join-Path $OutDir 'direct-events.cls'
    $directSource=(Get-Content -Raw (Join-Path $support 'com_server_client.bas')).Replace('__LIB__.','')
    $directSource=$directSource.Replace('Set first = New Calculator',"Dim globalObject As Globals`n    Set globalObject = New Globals`n    Set first = New Calculator")
    $directSource=$directSource.Replace('NativeGlobalAnswer()','globalObject.NativeGlobalAnswer()').Replace('NextGlobalValue()','globalObject.NextGlobalValue()')
    $directSource | Set-Content -LiteralPath $directClient -Encoding utf8NoBOM
    (Get-Content -Raw (Join-Path $support 'com_server_events.cls')).Replace('__LIB__.','') | Set-Content -LiteralPath $directEvent -Encoding utf8NoBOM
    $directModules=@(foreach($name in @('Calculator','Counter','PublicItem','PrivateItem','Globals')){
        @{name=$name;kind=2;source_path=(Join-Path $example ($name+'.cls'))}
    })
    $directModules+=@(@{name='EventClient';kind=2;source_path=$directEvent},@{name='OracleSelfTest';kind=1;source_path=$directClient})
    $directModulesPath=Join-Path $OutDir 'direct-modules.json'
    $directModules | ConvertTo-Json | Set-Content -LiteralPath $directModulesPath
    $directDirectory=Join-Path $OutDir 'direct-excel'
    & (Join-Path $PSHOME 'pwsh.exe') -NoLogo -NoProfile -STA -File (Join-Path $PSScriptRoot 'native-export-oracle-project.ps1') -OracleSession $OracleSession -Directory $directDirectory -ModulesFile $directModulesPath
    $directResult=Get-Content -Raw (Join-Path $directDirectory 'result.json') | ConvertFrom-Json
    if($directResult.compile_status -ne 'ok' -or $directResult.run_status -ne 'ok' -or $directResult.error -or $directResult.value -cne $expected){throw 'Ordinary Excel VBA source comparison failed'}
    Write-Output ('Excel/source:'+$directResult.value)
    $keys=[Collections.Generic.List[string]]::new()
    $keys.Add('Software\Classes\TypeLib\{'+$descriptor.libid+'}')
    foreach($class in $descriptor.classes) {
        if(-not $class.prog_id.StartsWith($projectName+'.')){throw 'Unexpected test ProgID'}
        $keys.Add('Software\Classes\CLSID\{'+$class.clsid+'}')
        $keys.Add('Software\Classes\'+$class.prog_id)
        $keys.Add('Software\Classes\Interface\{'+$class.default_interface_iid+'}')
        if($class.source_interface_iid){$keys.Add('Software\Classes\Interface\{'+$class.source_interface_iid+'}')}
    }
    foreach($key in $keys) {
        foreach($root in @([Microsoft.Win32.Registry]::CurrentUser,[Microsoft.Win32.Registry]::LocalMachine)) {
            $existing=$root.OpenSubKey($key)
            if($existing){$existing.Dispose();throw ('Test identity already registered: '+$key)}
        }
    }
    $keys | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $OutDir 'owned-registration.json')
    $attempted=$false
    try {
        $attempted=$true
        if([CalcVbaPrivateComProbe]::Call('DllRegisterServer') -ne 0){throw 'Per-user test registration failed'}
        foreach($class in $descriptor.classes) {
            $path='HKCU:\Software\Classes\CLSID\{'+$class.clsid+'}'
            if((Test-Path -LiteralPath $path) -ne [bool]$class.creatable){throw 'Registration ignores Instancing'}
            if($class.creatable -and (Get-ItemPropertyValue -LiteralPath ($path+'\InprocServer32') -Name ThreadingModel) -cne 'Apartment'){throw 'COM thread model differs'}
        }
        $clientPath=Join-Path $OutDir 'client.bas'
        $eventPath=Join-Path $OutDir 'events.cls'
        (Get-Content -Raw (Join-Path $support 'com_server_client.bas')).Replace('__LIB__',$projectName+'Lib') | Set-Content -LiteralPath $clientPath -Encoding utf8NoBOM
        (Get-Content -Raw (Join-Path $support 'com_server_events.cls')).Replace('__LIB__',$projectName+'Lib') | Set-Content -LiteralPath $eventPath -Encoding utf8NoBOM
        $modulesPath=Join-Path $OutDir 'modules.json'
        @(@{name='OracleSelfTest';source_path=$clientPath;kind=1},@{name='EventClient';source_path=$eventPath;kind=2}) | ConvertTo-Json | Set-Content -LiteralPath $modulesPath
        $excelDirectory=Join-Path $OutDir 'excel'
        & (Join-Path $PSHOME 'pwsh.exe') -NoLogo -NoProfile -STA -File (Join-Path $PSScriptRoot 'native-export-oracle-project.ps1') -OracleSession $OracleSession -Directory $excelDirectory -ModulesFile $modulesPath -ReferencePaths $tlb
        $result=Get-Content -Raw (Join-Path $excelDirectory 'result.json') | ConvertFrom-Json
        if($result.compile_status -ne 'ok' -or $result.run_status -ne 'ok' -or $result.error -or $result.value -cne $expected){throw 'Excel early binding comparison failed'}
        Write-Output ('Excel/COM:'+ $result.value)
    } finally {
        if($attempted -and [CalcVbaPrivateComProbe]::Call('DllUnregisterServer') -ne 0){throw 'Per-user test unregistration failed'}
        foreach($key in $keys) {
            $remaining=[Microsoft.Win32.Registry]::CurrentUser.OpenSubKey($key)
            if($remaining){$remaining.Dispose();throw ('Private test registration remains: '+$key)}
            $unexpected=[Microsoft.Win32.Registry]::LocalMachine.OpenSubKey($key)
            if($unexpected){$unexpected.Dispose();throw ('Unexpected machine registration: '+$key)}
        }
        'Private test identities unregistered' | Set-Content -LiteralPath (Join-Path $OutDir 'cleanup.txt')
    }
    Write-Output 'COM/registration:private HKCU identities removed; HKLM untouched'
} finally {Pop-Location}
