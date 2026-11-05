param(
    [Parameter(Mandatory=$true)][string]$OracleSession,
    [Parameter(Mandatory=$true)][string]$Directory
)
$ErrorActionPreference='Stop'
$PSNativeCommandUseErrorActionPreference=$true
$root=Split-Path $PSScriptRoot
$Directory=[IO.Path]::GetFullPath($Directory)
if (Test-Path -LiteralPath $Directory) { throw 'Use a new comparison directory' }
[void][IO.Directory]::CreateDirectory($Directory)
$source=Join-Path $Directory 'NativeExcel.bas'
Copy-Item -LiteralPath (Join-Path $root 'crates/calcvba-build/tests/support/native_members/NativeExcel.bas') -Destination $source
$moniker=Join-Path $Directory 'owned-running-excel'
$expected='Microsoft Excel:28:True'
@"
Attribute VB_Name = "RunnerMain"
Option Explicit
Public observed As String
Public Sub Main()
    Dim app As Excel.Application
    Set app = GetObject("$moniker").Application
    observed = NativeExcel.ExcelSuite(app)
    If observed <> "$expected" Then Err.Raise 513, , observed
End Sub
"@ | Set-Content -LiteralPath (Join-Path $Directory 'RunnerMain.bas')
$project=Join-Path $Directory 'Excel.basproj'
@'
<Project><PropertyGroup><OutputType>NativeLibrary</OutputType><ProjectName>NativeExcel</ProjectName></PropertyGroup><ItemGroup><Module Include="NativeExcel.bas"/><Module Include="RunnerMain.bas"/><COMReference Include="Excel"><Guid>{00020813-0000-0000-C000-000000000046}</Guid><VersionMajor>1</VersionMajor><VersionMinor>9</VersionMinor><Lcid>0</Lcid></COMReference><NativeExport Include="ExcelSuite"><Module>NativeExcel</Module><Procedure>ExcelSuite</Procedure></NativeExport></ItemGroup></Project>
'@ | Set-Content -LiteralPath $project
Push-Location $root
try {
    cargo build -p calcvba-cli
    $cli=Join-Path $root 'target/debug/calcvba-cli.exe'
    & $cli build $project --target NativeLibrary --out-dir (Join-Path $Directory 'out')
} finally { Pop-Location }
$dll=Join-Path $Directory 'out/NativeExcel.dll'
@"
Option Explicit
Private Declare PtrSafe Function NativeSuite Lib "$dll" Alias "ExcelSuite" (ByVal app As LongPtr) As LongPtr
Private Declare PtrSafe Function Shutdown Lib "$dll" Alias "CalcVbaShutdown" () As Long
Private Declare PtrSafe Function BStrLength Lib "oleaut32" Alias "SysStringLen" (ByVal pointer As LongPtr) As Long
Private Declare PtrSafe Sub FreeBStr Lib "oleaut32" Alias "SysFreeString" (ByVal pointer As LongPtr)
Private Declare PtrSafe Sub CopyBytes Lib "kernel32" Alias "RtlMoveMemory" (ByVal target As LongPtr, ByVal source As LongPtr, ByVal count As LongPtr)
Public Function ProbeNativeExports() As String
    Dim pointer As LongPtr, length As Long
    pointer=NativeSuite(ObjPtr(Application))
    length=BStrLength(pointer)
    ProbeNativeExports=String`$(length, ChrW`$(0))
    If length Then CopyBytes StrPtr(ProbeNativeExports), pointer, length*2&
    FreeBStr pointer
    If Shutdown()<>0 Then Err.Raise 5
End Function
"@ | Set-Content -LiteralPath (Join-Path $Directory 'OracleNative.bas')
@'
Option Explicit
Public Function ProbeNativeExports() As String
    ProbeNativeExports=NativeExcel.ExcelSuite(Application)
End Function
'@ | Set-Content -LiteralPath (Join-Path $Directory 'OracleSource.bas')
foreach ($mode in @('Native','Source')) {
    $modules=Join-Path $Directory ($mode+'-modules.json')
    @(@{name='NativeExcel';kind=1;source_path=$source},@{name='OracleSelfTest';kind=1;source_path=(Join-Path $Directory ('Oracle'+$mode+'.bas'))}) |
        ConvertTo-Json | Set-Content -LiteralPath $modules
    $resultDirectory=Join-Path $Directory $mode
    $options=@{OracleSession=$OracleSession;Directory=$resultDirectory;ModulesFile=$modules}
    if ($mode -eq 'Native') { $options.RunnerProject=$project; $options.RunningObjectPath=$moniker }
    & (Join-Path $PSScriptRoot 'native-export-oracle-project.ps1') @options
    $result=Get-Content -Raw -LiteralPath (Join-Path $resultDirectory 'result.json') | ConvertFrom-Json
    if ($result.value -cne $expected) { throw "$mode returned '$($result.value)'" }
}
foreach ($backend in @('vm3','jit')) {
    $values=Get-Content -Raw -LiteralPath (Join-Path $Directory ('Native/'+$backend+'.txt'))
    if (-not $values.Contains($expected)) { throw "$backend did not return $expected" }
}
