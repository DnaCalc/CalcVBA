# Runs probe modules in ONE owned Excel session and records compile and runtime outcomes.
#
# Spec JSON: {"cases":[{"id":"x","modules":[{"name":"M","kind":1,"path":"C:\\...\\M.bas"}],"entry":"M.Run"}]}
#   kind 1 = standard module, 2 = class module. Each case gets a fresh workbook; the modules are
#   imported, compiled through the visible VBE (Debug > Compile, command 578) with the PID-scoped
#   UIA guardian recording any dialog, and `entry` (a String function) is run only when the compile
#   passed. Results go to <Directory>\results.json.
#
# Run it under the private-desktop guard so no window reaches the owner's desktop:
#   pwsh -File on-private-desktop.ps1 -Command "pwsh -File scripts/run-excel-probe-session.ps1 ..." -SessionRoot <Directory> ...
# Safe Mode, owned PID only, no security or Trust Center change, exits by itself when done.
param(
    [Parameter(Mandatory=$true)][string]$Directory,
    [Parameter(Mandatory=$true)][string]$SpecPath,
    [string]$RunId=('probe-'+[Guid]::NewGuid().ToString('N'))
)
$ErrorActionPreference='Stop'
$Directory=[IO.Path]::GetFullPath($Directory)
if (Test-Path -LiteralPath $Directory) { throw 'Use a new session directory' }
[void][IO.Directory]::CreateDirectory($Directory)
$spec=Get-Content -Raw -LiteralPath $SpecPath | ConvertFrom-Json
$workerPath=Join-Path $PSScriptRoot 'excel-vba-oracle-worker.ps1'
$tokens=$null; $parseErrors=$null
$ast=[Management.Automation.Language.Parser]::ParseFile($workerPath,[ref]$tokens,[ref]$parseErrors)
if ($parseErrors.Count) { throw 'Excel helper script does not parse' }
$body=Get-Content -Raw -LiteralPath $workerPath
$start=$ast.ParamBlock.Extent.EndOffset
$end=$ast.Find({param($node) $node -is [Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq 'Invoke-HarnessCase'},$false).Extent.StartOffset
$definitions=$body.Substring($start,$end-$start)
$definitions=$definitions.Replace('$PSScriptRoot',("'"+$PSScriptRoot.Replace("'","''")+"'"))
$launchAnchor='$process = [Diagnostics.Process]::Start($startInfo)'
$definitions=$definitions.Replace($launchAnchor,$launchAnchor+@'

    $script:DirectProcess=$process
    [ordered]@{pid=$process.Id;start_utc=$process.StartTime.ToUniversalTime().ToString('o');path=$process.Path;launches=1} |
        ConvertTo-Json | Set-Content -LiteralPath (Join-Path $Directory 'launch.json') -Encoding utf8NoBOM
'@)
# On a private desktop EnumWindows reports failure (error 0) while Excel has no window yet; wait for
# the first window of the exact owned PID instead of refusing.
$enumAnchor='if (-not (Test-ExcelOracleWindowEnumerationAuthority -Enumeration $enumeration -ExpectedProcessId $process.Id)) {'
if (-not $definitions.Contains($enumAnchor)) { throw 'worker enumeration anchor moved' }
$definitions=$definitions.Replace($enumAnchor,'if ($enumeration.ErrorCode -eq 0 -and -not $enumeration.Truncated -and @($enumeration.Windows).Count -eq 0) { Start-Sleep -Milliseconds 200; continue }'+"`n        "+$enumAnchor)
. ([scriptblock]::Create($definitions))
# A private desktop shows the OS 'UAC Input Indicator' helper windows for the Excel process; they are
# not Excel startup dialogs.
function Test-ExcelStartupBlockingWindow {
    param([Parameter(Mandatory = $true)]$Window)
    if (-not [bool]$Window.IsTopLevel -or -not [bool]$Window.Visible) { return $false }
    return [string]$Window.ClassName -notin @('XLMAIN', 'MSOSPLASH', 'MsoSplash', 'UAC Input Indicator', 'UAC_InputIndicatorOverlayWnd')
}
$SafeMode=$true
$OwnershipFile=Join-Path $Directory 'owned-excel.jsonl'
$HelperOwnershipFile=Join-Path $Directory 'owned-helpers.jsonl'
$script:SelectedCaseIds=@('probe')
$script:GuardianControlSequence=0L
$controlFile=Join-Path $Directory 'guardian-control.json'
$readyFile=Join-Path $Directory 'guardian-ready.json'
$stopFile=Join-Path $Directory 'guardian-stop'
$identityFile=Join-Path $Directory 'excel-identity.json'
$guardian=$null; $excel=$null; $owned=$null; $script:DirectProcess=$null
$script:guardianGeneration=0
$results=[Collections.Generic.List[object]]::new()

function Start-ProbeGuardian {
    $script:guardianGeneration++
    $script:eventsFile=Join-Path $Directory ("guardian-events-$script:guardianGeneration.jsonl")
    $arguments=@('-NoLogo','-NoProfile','-NonInteractive','-STA','-File',(Join-Path $PSScriptRoot 'excel-vba-oracle-guardian.ps1'),
        '-ExcelPid',[string]$owned.Id,'-ExcelIdentityFile',$identityFile,'-RunId',$RunId,'-CaseId','probe',
        '-ControlFile',$controlFile,'-EventsFile',$script:eventsFile,'-ReadyFile',$readyFile,'-StopFile',$stopFile,'-MaxSeconds','3600')
    $script:guardian=Start-Process (Join-Path $PSHOME 'pwsh.exe') -ArgumentList $arguments -WindowStyle Hidden -PassThru `
        -RedirectStandardOutput (Join-Path $Directory ('guardian-'+[Guid]::NewGuid().ToString('N')+'.out')) `
        -RedirectStandardError (Join-Path $Directory ('guardian-'+[Guid]::NewGuid().ToString('N')+'.err'))
    [void](Add-HelperOwnershipRecord -Process $guardian -CaseId 'probe')
    $script:guardianReady=Wait-GuardianReady -ReadyFile $readyFile -Process $guardian
}
function Set-ProbePhase([string]$Phase,[string]$Operation) {
    Assert-GuardianLive -Process $guardian -ReadyRecord $guardianReady -Phase $Phase
    $control=Set-GuardianControl -Path $controlFile -CaseId 'probe' -OperationId $Operation -Phase $Phase -AllowDismiss $true
    Wait-GuardianOperationArmed -EventsFile $script:eventsFile -Control $control -Process $guardian -ReadyRecord $guardianReady
}
function Invoke-ProbeCase($Case) {
    $wb=$excel.Workbooks.Add()
    $record=[ordered]@{id=[string]$Case.id;compile_status='not-run';compile_dialogs=@();run_value=$null;run_error=$null;run_dialogs=@()}
    $first=$null
    try {
        foreach ($m in $Case.modules) {
            $comp=$wb.VBProject.VBComponents.Add([int]$m.kind)
            $comp.Name=[string]$m.name
            $text=[IO.File]::ReadAllText([string]$m.path)
            # Drop an exported header (Attribute VB_Name / VERSION / BEGIN blocks) so AddFromString gets plain code.
            $lines=@($text -split "`r?`n")
            $code=[Collections.Generic.List[string]]::new()
            $inBegin=$false
            foreach ($l in $lines) {
                if ($l -match '^VERSION ') { continue }
                if ($l -match '^BEGIN\s*$') { $inBegin=$true; continue }
                if ($inBegin) { if ($l -match '^END\s*$') { $inBegin=$false }; continue }
                if ($l -match '^Attribute VB_(Name|GlobalNameSpace|Creatable|PredeclaredId|Exposed)\b') { continue }
                $code.Add($l)
            }
            [void]$comp.CodeModule.AddFromString(($code -join "`r`n"))
            if ($null -eq $first) { $first=$comp }
        }
        $wb.Activate()
        $excel.VBE.MainWindow.Visible=$true
        $first.CodeModule.CodePane.Show()
        try { $excel.VBE.MainWindow.SetFocus() } catch { }
        Start-Sleep -Milliseconds 250
        $control=Get-VbeCompileControl -Vbe $excel.VBE
        if (-not $control -or -not $control.Enabled) { $record.compile_status='compile-disabled'; return $record }
        $op='compile-'+[Guid]::NewGuid().ToString('N')
        $armed=Set-ProbePhase compile $op
        $control.Execute()
        $events=@(Wait-GuardianEventFlush -EventsFile $script:eventsFile -OperationId $op -ArmRecord $armed -InvocationCompletedUtc ([DateTime]::UtcNow) -CompileControl $control)
        $dialogs=@($events | Where-Object event_type -eq 'dialog-observation')
        $record.compile_dialogs=@($dialogs)
        $kinds=@($dialogs | ForEach-Object { [string]$_.classification })
        if ($kinds -contains 'security-or-trust' -or $kinds -contains 'unrecognized-modal') { $record.compile_status='blocked'; return $record }
        if ($kinds -contains 'compile-error') { $record.compile_status='compile-error' }
        elseif (-not $control.Enabled) { $record.compile_status='ok' }
        else { $record.compile_status='no-dialog-unverified' }
        $excel.VBE.MainWindow.Visible=$false
        if ($record.compile_status -eq 'ok' -and $Case.entry) {
            $runOp='run-'+[Guid]::NewGuid().ToString('N')
            $runArm=Set-ProbePhase run $runOp
            try { $record.run_value=[string]$excel.Run("'$($wb.Name)'!$($Case.entry)") }
            catch { $record.run_error=$_.Exception.Message }
            $runEvents=@(Wait-GuardianEventFlush -EventsFile $script:eventsFile -OperationId $runOp -ArmRecord $runArm -InvocationCompletedUtc ([DateTime]::UtcNow))
            $record.run_dialogs=@($runEvents | Where-Object event_type -eq 'dialog-observation')
        }
    } finally {
        try { $excel.VBE.MainWindow.Visible=$false } catch { }
        try { $wb.Close($false) } catch { }
        Release-ComObject $wb
    }
    return $record
}
try {
    $before=@(Get-ExcelProcessIds)
    $bootstrap=New-ExcelOracleBootstrapWorkbook -Path (Join-Path $Directory 'bootstrap.xlsx')
    $launched=Start-OwnedExcelApplication -ExcelExecutable (Get-ExcelExecutablePath) -BootstrapWorkbook $bootstrap -DiagnosticPath (Join-Path $Directory 'attachment.json')
    $owned=$launched.process; $excel=$launched.application; $native=$launched.native_window
    if ($owned.Id -in $before -or (Get-ExcelPidFromApplication $excel) -ne $owned.Id) { throw 'Excel ownership mismatch' }
    $identity=Add-OwnershipRecord -Process $owned -BeforePids $before -CaseId 'probe'
    $identity | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $identityFile -Encoding utf8NoBOM
    $excel.Visible=$true; $excel.DisplayAlerts=$false; $excel.EnableEvents=$false
    [ordered]@{pid=$owned.Id;version=[string]$excel.Version;build=[string]$excel.Build;automation_security=[int]$excel.AutomationSecurity} |
        ConvertTo-Json | Set-Content -LiteralPath (Join-Path $Directory 'excel-environment.json') -Encoding utf8NoBOM
    Start-ProbeGuardian
    [void](Set-ProbePhase run 'session-ready')
    foreach ($case in $spec.cases) {
        $r=Invoke-ProbeCase $case
        $results.Add($r)
        $results | ConvertTo-Json -Depth 12 | Set-Content -LiteralPath (Join-Path $Directory 'results.json') -Encoding utf8NoBOM
    }
} finally {
    if ($excel -and $owned -and -not $owned.HasExited -and (Get-ExcelPidFromApplication $excel) -eq $owned.Id) {
        try { $excel.EnableEvents=$false; while ($excel.Workbooks.Count) { $book=$excel.Workbooks.Item(1); $book.Close($false); Release-ComObject $book }; $excel.Quit() } catch { }
    }
    if ($guardian -and -not $guardian.HasExited) {
        New-Item -ItemType File -Path $stopFile -Force | Out-Null
        [void]$guardian.WaitForExit(5000)
    }
    Release-ComObject $excel
    if (Get-Variable native -ErrorAction SilentlyContinue) { Release-ComObject $native }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    $cleanup=if ($owned) { $owned } else { $script:DirectProcess }
    if ($cleanup) {
        if (-not $cleanup.WaitForExit(5000)) {
            $live=Get-Process -Id $cleanup.Id -ErrorAction SilentlyContinue
            if ($live -and $live.StartTime.ToUniversalTime() -eq $cleanup.StartTime.ToUniversalTime() -and $live.Path -eq $cleanup.Path) { Stop-Process -Id $live.Id }
        }
        [ordered]@{pid=$cleanup.Id;exited=$cleanup.HasExited;launches=1} | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $Directory 'cleanup.json')
    }
}
