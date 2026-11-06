param(
    [Parameter(Mandatory=$true)][string]$OracleSession,
    [Parameter(Mandatory=$true)][string]$Directory,
    [Parameter(Mandatory=$true)][string]$ModulesFile,
    [string[]]$ReferencePaths=@(),
    [string]$Entry='OracleSelfTest.ProbeNativeExports',
    [string]$RunnerProject,
    [string]$RunningObjectPath
)
$ErrorActionPreference='Stop'
$scripts=Join-Path (Resolve-Path (Join-Path $PSScriptRoot '..')) 'scripts'
$Directory=[IO.Path]::GetFullPath($Directory)
if (Test-Path -LiteralPath $Directory) { throw 'Use a new project comparison directory' }
[void][IO.Directory]::CreateDirectory($Directory)
$OracleSession=(Resolve-Path -LiteralPath $OracleSession).Path
$ready=Get-Content -Raw -LiteralPath (Join-Path $OracleSession 'ready.json') | ConvertFrom-Json
$identityFile=Join-Path $OracleSession 'excel-identity.json'
$identity=Get-Content -Raw -LiteralPath $identityFile | ConvertFrom-Json
$RunId=[string]$ready.run_id
$owned=Get-Process -Id ([int]$ready.pid)
if ($owned.Id -eq 3144 -or $ready.launches -ne 1 -or $identity.ownership -ne 'owned-new-instance' -or
    $owned.Id -ne $identity.pid -or $RunId -cne $identity.run_id -or
    $owned.StartTime.ToUniversalTime() -ne ([datetime]$identity.process_start_utc).ToUniversalTime() -or $owned.Path -cne $identity.executable_path) {
    throw 'Owned Excel identity mismatch'
}
# Reuse the exact-PID attachment and UIA compile/dialog helpers. No Excel launch,
# security-setting changes, registration changes, or Excel process cleanup here.
$workerPath=Join-Path $scripts 'excel-vba-oracle-worker.ps1'
$tokens=$null; $errors=$null
$ast=[Management.Automation.Language.Parser]::ParseFile($workerPath,[ref]$tokens,[ref]$errors)
if ($errors.Count) { throw 'Excel worker does not parse' }
$body=Get-Content -Raw -LiteralPath $workerPath
$start=$ast.ParamBlock.Extent.EndOffset
$end=$ast.Find({param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq 'Invoke-HarnessCase'},$false).Extent.StartOffset
$definitions=$body.Substring($start,$end-$start).Replace('$PSScriptRoot',("'"+$scripts.Replace("'","''")+"'"))
. ([scriptblock]::Create($definitions))
$sessionPath=Join-Path $scripts 'run-native-export-oracle-session.ps1'
$ast=[Management.Automation.Language.Parser]::ParseFile($sessionPath,[ref]$tokens,[ref]$errors)
if ($errors.Count) { throw 'Excel session helper does not parse' }
foreach ($name in @('Start-OfficeGuardian','Set-OfficePhase','Invoke-OfficeCompile')) {
    $node=$ast.Find({param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $name},$false)
    . ([scriptblock]::Create($node.Extent.Text.Replace('$PSScriptRoot',("'"+$scripts.Replace("'","''")+"'"))))
}
$HelperOwnershipFile=Join-Path $Directory 'owned-helpers.jsonl'
$script:SelectedCaseIds=@('office-project')
$script:GuardianControlSequence=0L
$script:guardianGeneration=0
$controlFile=Join-Path $Directory 'guardian-control.json'
$eventsFile=Join-Path $Directory 'guardian-events.jsonl'
$readyFile=Join-Path $Directory 'guardian-ready.json'
$stopFile=Join-Path $Directory 'guardian-stop'
$guardian=$null; $excel=$null; $native=$null; $book=$null
$runningTable=$null; $runningMoniker=$null; $runningCookie=0
$record=[ordered]@{compile_status='not-run';run_status='not-run';value=$null;error=$null;pid=$owned.Id;launches=1}
try {
    $enumeration=[ExcelOracleNativeMethods]::EnumerateOwnedWindows([uint32]$owned.Id)
    if (-not (Test-ExcelOracleWindowEnumerationAuthority -Enumeration $enumeration -ExpectedProcessId $owned.Id)) { throw 'Incomplete owned-window enumeration' }
    foreach ($window in @($enumeration.Windows | Sort-Object @{Expression={if($_.ClassName -eq 'EXCEL7'){0}else{1}}},Hwnd)) {
        $candidate=$null; $application=$null
        $hr=[ExcelOracleNativeMethods]::TryGetNativeObjectFromWindow([IntPtr][int64]$window.Hwnd,[ref]$candidate)
        if ($candidate) { try { $application=$candidate.Application } catch { Release-ComObject $candidate; continue } }
        $applicationPid=if($application){Get-ExcelPidFromApplication $application}else{0}
        $decision=Resolve-ExcelOracleAttachmentCandidate -Enumeration $enumeration -ExpectedProcessId $owned.Id -Candidate $window -HResult $hr -NativeObjectPresent ($null -ne $candidate) -ApplicationPresent ($null -ne $application) -ApplicationPid $applicationPid
        if ($decision.attached) { $excel=$application; $native=$candidate; break }
        Release-ComObject $application; Release-ComObject $candidate
    }
    if (-not $excel) { throw 'Cannot attach to the recorded owned Excel' }
    if ($RunnerProject) {
        if (-not $RunningObjectPath) { throw 'RunnerProject requires a unique RunningObjectPath' }
        Add-Type @'
using System;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
[ComVisible(true), ClassInterface(ClassInterfaceType.AutoDispatch)]
public class NativeOracleRunningExcel {
    public object Application { get; set; }
}
public static class NativeOracleRunningObject {
    [DllImport("ole32.dll")] public static extern int GetRunningObjectTable(int reserved, out IRunningObjectTable table);
    [DllImport("ole32.dll", CharSet=CharSet.Unicode)] public static extern int CreateFileMoniker(string path, out IMoniker moniker);
    static IRunningObjectTable table;
    static IMoniker moniker;
    public static int Register(object value, string path) {
        Marshal.ThrowExceptionForHR(GetRunningObjectTable(0,out table));
        Marshal.ThrowExceptionForHR(CreateFileMoniker(path,out moniker));
        return table.Register(1,new NativeOracleRunningExcel { Application=value },moniker);
    }
    public static void Revoke(int cookie) { table.Revoke(cookie); Marshal.ReleaseComObject(moniker); Marshal.ReleaseComObject(table); }
}
'@
        $runningCookie=[NativeOracleRunningObject]::Register($excel.PSObject.BaseObject,[IO.Path]::GetFullPath($RunningObjectPath))
        $runner=Join-Path (Split-Path $scripts) 'target/debug/calcvba-cli.exe'
        foreach ($backend in @('vm3','jit')) {
            $output=& $runner run-project $RunnerProject --backend $backend --dump-values 2>&1
            $output | Set-Content -LiteralPath (Join-Path $Directory ($backend+'.txt'))
            if ($LASTEXITCODE -ne 0) { throw "Owned running-object $backend comparison failed: $output" }
        }
    }
    Start-OfficeGuardian
    [void](Set-OfficePhase run 'project-import')
    $book=$excel.Workbooks.Add()
    foreach ($path in $ReferencePaths) {
        $reference=$book.VBProject.References.AddFromFile((Resolve-Path -LiteralPath $path).Path)
        Release-ComObject $reference
    }
    foreach ($module in @(Get-Content -Raw -LiteralPath $ModulesFile | ConvertFrom-Json)) {
        $kind=if($module.PSObject.Properties.Name -contains 'kind'){[int]$module.kind}else{1}
        if($kind -notin @(1,2)){throw 'Only standard and class modules are supported'}
        $component=$book.VBProject.VBComponents.Add($kind)
        $component.Name=[string]$module.name
        $source=(Get-Content -Raw -LiteralPath ([string]$module.source_path)) -replace '(?m)^Attribute .*\r?\n',''
        $source=$source -replace '(?is)\AVERSION[^\r\n]*\r?\nBEGIN\r?\n.*?\r?\nEND\r?\n',''
        [void]$component.CodeModule.AddFromString($source)
        Release-ComObject $component
    }
    Invoke-OfficeCompile -ShowWindow -Workbook $book -ModuleName ($Entry.Split('.')[0])
    $record.compile_status='ok'
    $record.value=$excel.Run("'"+$book.Name+"'!"+$Entry)
    $record.run_status='ok'
} catch {
    $record.error=[string]$_
    if ($record.compile_status -eq 'not-run') { $record.compile_status='failed' } else { $record.run_status='failed' }
} finally {
    if ($runningCookie) { [NativeOracleRunningObject]::Revoke($runningCookie) }
    Release-ComObject $runningMoniker; Release-ComObject $runningTable
    if ($book) { $book.Close($false); Release-ComObject $book }
    if ($guardian -and -not $guardian.HasExited) {
        New-Item -ItemType File -Path $stopFile -Force | Out-Null
        if (-not $guardian.WaitForExit(5000)) { Stop-Process -InputObject $guardian; [void]$guardian.WaitForExit(5000) }
    }
    Release-ComObject $excel; Release-ComObject $native
    $record | ConvertTo-Json -Depth 12 | Set-Content -LiteralPath (Join-Path $Directory 'result.json') -Encoding utf8NoBOM
}
if ($record.error) { throw $record.error }
$record | ConvertTo-Json -Compress -Depth 12
