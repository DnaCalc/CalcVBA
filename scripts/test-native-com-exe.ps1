param(
    [Parameter(Mandatory=$true)][string]$ServerDirectory,
    [Parameter(Mandatory=$true)][string]$Directory,
    [string]$OracleSession,
    [string]$CliPath=(Join-Path $PSScriptRoot "../target/debug/calcvba-cli.exe")
)
$ErrorActionPreference='Stop'
$PSNativeCommandUseErrorActionPreference=$false
$ServerDirectory=(Resolve-Path -LiteralPath $ServerDirectory).Path
$Directory=[IO.Path]::GetFullPath($Directory)
if(Test-Path -LiteralPath $Directory){throw 'Use a new test directory'}
[void][IO.Directory]::CreateDirectory($Directory)
$descriptorFile=@(Get-ChildItem -LiteralPath $ServerDirectory -Filter '*.comserver.json')
if($descriptorFile.Count -ne 1){throw 'Expected one server descriptor'}
$descriptor=Get-Content -Raw -LiteralPath $descriptorFile[0].FullName|ConvertFrom-Json
$stem=[string]$descriptor.project_name
$executable=Join-Path $ServerDirectory ($stem+'.exe')
$typelib=Join-Path $ServerDirectory ($stem+'.tlb')
$class=$descriptor.classes[0]
if($descriptor.classes.Count -ne 1 -or $class.class_name -ne 'Counter'){throw 'Use the private native Counter fixture'}
$clsid='{'+$class.clsid.Trim('{}')+'}'
$iid='{'+$class.default_interface_iid.Trim('{}')+'}'
$libid='{'+$descriptor.libid.Trim('{}')+'}'
$keys=@(('CLSID\'+$clsid),('Interface\'+$iid),('TypeLib\'+$libid),[string]$class.prog_id)
$root=[Microsoft.Win32.Registry]::CurrentUser.OpenSubKey('Software\Classes',$true)
foreach($key in $keys){$existing=$root.OpenSubKey($key);if($existing){$existing.Dispose();$root.Dispose();throw "Refusing existing per-user key $key"}}
$source=@'
using System;
using System.Runtime.InteropServices;
using System.Threading;
[ComImport, Guid("__IID__"), InterfaceType(ComInterfaceType.InterfaceIsDual)]
public interface INativeCounter {
    [DispId(1)] int Add(int left, int right);
    [DispId(2)] [return: MarshalAs(UnmanagedType.BStr)] string Echo([MarshalAs(UnmanagedType.BStr)] string text);
    [DispId(3)] int Exchange(ref int value);
    [DispId(4)] int Value { get; set; }
    [DispId(5)] int ThreadId();
    [DispId(6)] int Fail();
}
public static class NativeComExeClient {
    [DllImport("oleaut32.dll",CharSet=CharSet.Unicode)] static extern int LoadTypeLibEx(string path,int kind,out IntPtr library);
    [DllImport("oleaut32.dll",CharSet=CharSet.Unicode)] static extern int RegisterTypeLibForUser(IntPtr library,string path,string help);
    [DllImport("oleaut32.dll")] static extern int UnRegisterTypeLibForUser(ref Guid libid,ushort major,ushort minor,uint lcid,int syskind);
    public static void Register(string path) { IntPtr library; Marshal.ThrowExceptionForHR(LoadTypeLibEx(path,2,out library));try{Marshal.ThrowExceptionForHR(RegisterTypeLibForUser(library,path,null));}finally{Marshal.Release(library);} }
    public static void Unregister(string id) { Guid libid=new Guid(id); Marshal.ThrowExceptionForHR(UnRegisterTypeLibForUser(ref libid,1,0,0,3)); }
    public static object Hold() { return Activator.CreateInstance(Type.GetTypeFromCLSID(new Guid("__CLSID__"),true)); }
    public static void Drop(object value) { Marshal.FinalReleaseComObject(value); }
    public static string Run() {
        object a=null,b=null;
        try {
            Type type=Type.GetTypeFromCLSID(new Guid("__CLSID__"),true);
            a=Activator.CreateInstance(type); b=Activator.CreateInstance(type);
            var first=(INativeCounter)a; var second=(INativeCounter)b;
            string result=first.Add(2,3)+":"+second.Value;
            first.Value=41; int value=7;
            result+=":"+first.Exchange(ref value)+":"+value+":"+first.Echo("hello");
            int owner=first.ThreadId();
            result+=":"+(owner==second.ThreadId());
            try{first.Fail();result+=":missing-error";}catch(COMException error){result+=":"+(error.ErrorCode&65535);}
            result+=":"+(string)type.InvokeMember("Echo",System.Reflection.BindingFlags.InvokeMethod,null,a,new object[]{"late"});
            return result+"|owner="+owner;
        } finally { if(b!=null)Marshal.FinalReleaseComObject(b); if(a!=null)Marshal.FinalReleaseComObject(a); }
    }
}
'@
$source=$source.Replace('__IID__',$iid.Trim('{}')).Replace('__CLSID__',$clsid.Trim('{}'))
Add-Type -TypeDefinition $source
$server=$null;$holder=$null;$registered=$false;$created=[Collections.Generic.List[string]]::new()
function Set-PrivateValue([string]$Path,[string]$Name,[string]$Value){$key=$root.CreateSubKey($Path);try{$key.SetValue($Name,$Value,[Microsoft.Win32.RegistryValueKind]::String)}finally{$key.Dispose()}}
try{
    foreach($key in $keys){$created.Add($key)}
    [NativeComExeClient]::Register($typelib);$registered=$true
    Set-PrivateValue ('CLSID\'+$clsid) '' ([string]$class.class_name)
    Set-PrivateValue ('CLSID\'+$clsid+'\LocalServer32') '' ('"'+$executable+'" /Embedding')
    Set-PrivateValue ('CLSID\'+$clsid+'\ProgID') '' ([string]$class.prog_id)
    Set-PrivateValue ([string]$class.prog_id+'\CLSID') '' $clsid
    Set-PrivateValue ('Interface\'+$iid) '' ([string]$class.default_interface_name)
    Set-PrivateValue ('Interface\'+$iid+'\ProxyStubClsid32') '' '{00020424-0000-0000-C000-000000000046}'
    Set-PrivateValue ('Interface\'+$iid+'\TypeLib') '' $libid
    Set-PrivateValue ('Interface\'+$iid+'\TypeLib') 'Version' '1.0'
    $server=Start-Process -FilePath $executable -ArgumentList '/Embedding' -WindowStyle Hidden -PassThru -RedirectStandardOutput (Join-Path $Directory 'server.out') -RedirectStandardError (Join-Path $Directory 'server.err')
    $identity=[ordered]@{pid=$server.Id;start_utc=$server.StartTime.ToUniversalTime().ToString('o');path=$server.Path;launches=1}
    $identity|ConvertTo-Json|Set-Content -LiteralPath (Join-Path $Directory 'owned-server.json')
    $deadline=[DateTime]::UtcNow.AddSeconds(30)
    while(-not ((Get-Content -Raw -LiteralPath (Join-Path $Directory 'server.err') -ErrorAction SilentlyContinue) -match 'COM server ready')){if($server.HasExited -or [DateTime]::UtcNow -ge $deadline){throw 'Owned server did not become ready'};Start-Sleep -Milliseconds 100}
    $holder=[NativeComExeClient]::Hold()
    $value=[NativeComExeClient]::Run()
    $value|Set-Content -LiteralPath (Join-Path $Directory 'early-client.txt')
    if($value -notmatch '^15:10:7:41:hello:42:True:513:late:42\|owner=(\d+)$'){throw "Early-bound client mismatch: $value"}
    $owner=$Matches[1]
    if((Get-Content -Raw -LiteralPath (Join-Path $Directory 'server.err')) -notmatch ('thread='+$owner+'\b')){throw 'VBA did not run on the owning message-loop thread'}
    $expected='15:10:7:41:hello:42:True:513:late:42'
    $clientSource=Get-Content -Raw -LiteralPath (Join-Path $PSScriptRoot '../examples/native/comexe/Client.bas')
    $clientSource+="`nPublic Sub Main()`nDim observed As String`nobserved = Probe()`nDebug.Print observed`nEnd Sub`n"
    $clientSource|Set-Content -LiteralPath (Join-Path $Directory 'Client.bas')
    $escapedTypelib=[Security.SecurityElement]::Escape($typelib)
    $clientProject=Join-Path $Directory 'Client.basproj'
    @"
<Project><PropertyGroup><OutputType>Exe</OutputType><ProjectName>NativeCounterClient</ProjectName><EntryPoint>ComExeClient.Main</EntryPoint></PropertyGroup><ItemGroup><Module Include="Client.bas"/><COMReference Include="${stem}Lib"><Guid>$libid</Guid><VersionMajor>1</VersionMajor><VersionMinor>0</VersionMinor><Lcid>0</Lcid><ImportLib>$escapedTypelib</ImportLib></COMReference></ItemGroup></Project>
"@|Set-Content -LiteralPath $clientProject
    foreach($backend in @('vm3','jit')){
        $output=& $CliPath run-project $clientProject --backend $backend --dump-values 2>&1
        $output|Set-Content -LiteralPath (Join-Path $Directory ($backend+'-client.txt'))
        if($LASTEXITCODE -ne 0 -or ($output -join "`n") -notlike ('*'+$expected+'*')){throw "Compiled server $backend client failed: $output"}
    }
    & $CliPath build $clientProject --target Exe --out-dir (Join-Path $Directory 'native-client')|Set-Content -LiteralPath (Join-Path $Directory 'native-client-build.txt')
    if($LASTEXITCODE -ne 0){throw 'AOT COM client build failed'}
    $output=& (Join-Path $Directory 'native-client/NativeCounterClient.exe') 2>&1
    $output|Set-Content -LiteralPath (Join-Path $Directory 'aot-client.txt')
    if($LASTEXITCODE -ne 0 -or ($output -join "`n").Trim() -cne $expected){throw "AOT COM client failed: $output"}
    if($OracleSession){
        # The held private instance keeps this same owned server alive for Office.
        $modules=@(@{name='ComExeClient';kind=1;source_path=(Join-Path $PSScriptRoot '../examples/native/comexe/Client.bas')})
        $modules|ConvertTo-Json -AsArray|Set-Content -LiteralPath (Join-Path $Directory 'office-modules.json')
        & (Join-Path $PSScriptRoot 'native-export-oracle-project.ps1') -OracleSession $OracleSession -Directory (Join-Path $Directory 'office') -ModulesFile (Join-Path $Directory 'office-modules.json') -ReferencePaths @($typelib) -Entry ComExeClient.Probe
        $result=Get-Content -Raw -LiteralPath (Join-Path $Directory 'office/result.json')|ConvertFrom-Json
        if($result.value -cne '15:10:7:41:hello:42:True:513:late:42'){throw 'Office early-bound server client mismatch'}
    }
    [NativeComExeClient]::Drop($holder);$holder=$null
    if(-not $server.WaitForExit(10000) -or $server.ExitCode -ne 0){throw 'Owned server did not shut down cleanly'}
    [ordered]@{early_bound=$value;owning_thread=[int]$owner;excel_launches=0;server_launches=1;passed=$true}|ConvertTo-Json|Set-Content -LiteralPath (Join-Path $Directory 'result.json')
}finally{
    if($holder){[NativeComExeClient]::Drop($holder);$holder=$null}
    if($server -and -not $server.HasExited){$live=Get-Process -Id $server.Id -ErrorAction SilentlyContinue;if($live -and $live.Path -ceq $executable -and $live.StartTime.ToUniversalTime() -eq $server.StartTime.ToUniversalTime()){Stop-Process -InputObject $live}}
    if($registered){[NativeComExeClient]::Unregister($libid)}
    foreach($key in $created){$root.DeleteSubKeyTree($key,$false)}
    $remaining=@($created|Where-Object {$key=$root.OpenSubKey($_);if($key){$key.Dispose();$true}})
    [ordered]@{remaining_keys=$remaining;owned_server_exited=($null -eq $server -or $server.HasExited)}|ConvertTo-Json|Set-Content -LiteralPath (Join-Path $Directory 'cleanup.json')
    $root.Dispose()
}
