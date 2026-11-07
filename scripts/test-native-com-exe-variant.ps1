param(
    [Parameter(Mandatory=$true)][string]$ServerDirectory,
    [Parameter(Mandatory=$true)][string]$Directory
)
# Private Boxer server (examples/native/comexe/Boxer.cls): Variant boundaries proved by an
# early-bound C# client against the registered type library. HKCU only; removed afterwards.
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
if($descriptor.classes.Count -ne 1 -or $class.class_name -ne 'Boxer'){throw 'Use the private native Boxer fixture'}
$clsid='{'+$class.clsid.Trim('{}')+'}'
$iid='{'+$class.default_interface_iid.Trim('{}')+'}'
$libid='{'+$descriptor.libid.Trim('{}')+'}'
$keys=@(('CLSID\'+$clsid),('Interface\'+$iid),('TypeLib\'+$libid),[string]$class.prog_id)
$root=[Microsoft.Win32.Registry]::CurrentUser.OpenSubKey('Software\Classes',$true)
foreach($key in $keys){$existing=$root.OpenSubKey($key);if($existing){$existing.Dispose();$root.Dispose();throw "Refusing existing per-user key $key"}}
$source=@'
using System;
using System.Runtime.InteropServices;
[ComImport, Guid("__IID__"), InterfaceType(ComInterfaceType.InterfaceIsDual)]
public interface IBoxer {
    [DispId(1)] [return: MarshalAs(UnmanagedType.Struct)] object EchoVariant([MarshalAs(UnmanagedType.Struct)] object value);
    [DispId(2)] [return: MarshalAs(UnmanagedType.BStr)] string Describe([MarshalAs(UnmanagedType.Struct)] object value);
    [DispId(3)] void Fill(ref object value);
    [DispId(4)] object Stored { [return: MarshalAs(UnmanagedType.Struct)] get; [param: MarshalAs(UnmanagedType.Struct)] set; }
    [DispId(5)] int CountOf([MarshalAs(UnmanagedType.IDispatch)] object value);
    [DispId(6)] [return: MarshalAs(UnmanagedType.IDispatch)] object Pass([MarshalAs(UnmanagedType.IDispatch)] object value);
    [DispId(7)] [return: MarshalAs(UnmanagedType.BStr)] string Pick(int a, [MarshalAs(UnmanagedType.Struct)] object b, [MarshalAs(UnmanagedType.Struct)] object c);
}
[ComVisible(true), ClassInterface(ClassInterfaceType.AutoDispatch)]
public class Counted { public int Count { get { return 3; } } }
public static class NativeComExeClient {
    [DllImport("oleaut32.dll",CharSet=CharSet.Unicode)] static extern int LoadTypeLibEx(string path,int kind,out IntPtr library);
    [DllImport("oleaut32.dll",CharSet=CharSet.Unicode)] static extern int RegisterTypeLibForUser(IntPtr library,string path,string help);
    [DllImport("oleaut32.dll")] static extern int UnRegisterTypeLibForUser(ref Guid libid,ushort major,ushort minor,uint lcid,int syskind);
    public static void Register(string path) { IntPtr library; Marshal.ThrowExceptionForHR(LoadTypeLibEx(path,2,out library));try{Marshal.ThrowExceptionForHR(RegisterTypeLibForUser(library,path,null));}finally{Marshal.Release(library);} }
    public static void Unregister(string id) { Guid libid=new Guid(id); Marshal.ThrowExceptionForHR(UnRegisterTypeLibForUser(ref libid,1,0,0,3)); }
    public static object Hold() { return Activator.CreateInstance(Type.GetTypeFromCLSID(new Guid("__CLSID__"),true)); }
    public static void Drop(object value) { Marshal.FinalReleaseComObject(value); }
    static string Text(object value) {
        if (value == null) return "null";
        if (value is Array) { var array=(Array)value; string text="array"+array.Length+"["; foreach (object item in array) text+=Text(item)+","; return text+"]"; }
        return value.GetType().Name+"="+value;
    }
    static string Step(string name, Func<string> action) {
        try { return "|"+action(); } catch (Exception error) { return "|"+name+" failed 0x"+Marshal.GetHRForException(error).ToString("X8"); }
    }
    public static string Run() {
        object a=null;
        try {
            a=Activator.CreateInstance(Type.GetTypeFromCLSID(new Guid("__CLSID__"),true));
            var boxer=(IBoxer)a;
            string result="";
            result+=Step("echo-int",()=>Text(boxer.EchoVariant(42)));
            result+=Step("echo-string",()=>Text(boxer.EchoVariant("text")));
            result+=Step("echo-double",()=>Text(boxer.EchoVariant(2.5)));
            result+=Step("echo-bool",()=>Text(boxer.EchoVariant(true)));
            result+=Step("echo-short",()=>Text(boxer.EchoVariant((short)7)));
            result+=Step("echo-null",()=>Text(boxer.EchoVariant(null)));
            result+=Step("echo-array",()=>Text(boxer.EchoVariant(new object[]{1,"two",3.5})));
            result+=Step("describe",()=>boxer.Describe(5)+","+boxer.Describe((short)5)+","+boxer.Describe("x")+","+boxer.Describe(2.5)+","+boxer.Describe(false));
            result+=Step("fill-int",()=>{object filled=7; boxer.Fill(ref filled); return Text(filled);});
            result+=Step("fill-string",()=>{object text="was text"; boxer.Fill(ref text); return Text(text);});
            result+=Step("stored-int",()=>{boxer.Stored=99; return Text(boxer.Stored);});
            result+=Step("stored-string",()=>{boxer.Stored="stored"; return Text(boxer.Stored);});
            result+=Step("count-of",()=>boxer.CountOf(new Counted()).ToString());
            result+=Step("pass",()=>{var counted=new Counted(); object back=boxer.Pass(counted); return (back==counted).ToString()+","+boxer.CountOf(back);});
            result+=Step("optional",()=>boxer.Pick(1,Type.Missing,Type.Missing)+","+boxer.Pick(1,2,Type.Missing)+","+boxer.Pick(1,Type.Missing,3)+","+boxer.Pick(1,2,3));
            return result;
        } finally { if(a!=null)Marshal.FinalReleaseComObject(a); }
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
    $expected='|Int32=42|String=text|Double=2.5|Boolean=True|Int16=7|null|array3[Int32=1,String=two,Double=3.5,]|Long:5,Integer:5,String:x,Double:2.5,Boolean:False|String=filled:Long|String=filled:String|Int32=99|String=stored|3|True,3|1:missing:missing,1:2:missing,1:missing:3,1:2:3'
    if($value -cne $expected){throw "Early-bound client mismatch: $value"}
    [NativeComExeClient]::Drop($holder);$holder=$null
    if(-not $server.WaitForExit(10000) -or $server.ExitCode -ne 0){throw 'Owned server did not shut down cleanly'}
    [ordered]@{early_bound=$value;excel_launches=0;server_launches=1;passed=$true}|ConvertTo-Json|Set-Content -LiteralPath (Join-Path $Directory 'result.json')
}finally{
    if($holder){[NativeComExeClient]::Drop($holder);$holder=$null}
    if($server -and -not $server.HasExited){$live=Get-Process -Id $server.Id -ErrorAction SilentlyContinue;if($live -and $live.Path -ceq $executable -and $live.StartTime.ToUniversalTime() -eq $server.StartTime.ToUniversalTime()){Stop-Process -InputObject $live}}
    if($registered){[NativeComExeClient]::Unregister($libid)}
    foreach($key in $created){$root.DeleteSubKeyTree($key,$false)}
    $remaining=@($created|Where-Object {$key=$root.OpenSubKey($_);if($key){$key.Dispose();$true}})
    [ordered]@{remaining_keys=$remaining;owned_server_exited=($null -eq $server -or $server.HasExited)}|ConvertTo-Json|Set-Content -LiteralPath (Join-Path $Directory 'cleanup.json')
    $root.Dispose()
}
