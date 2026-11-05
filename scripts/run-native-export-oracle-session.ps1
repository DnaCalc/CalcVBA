param(
    [Parameter(Mandatory=$true)][string]$Directory,
    [string]$RunId=('native-interop-'+[Guid]::NewGuid().ToString('N')),
    [switch]$Visible
)
$ErrorActionPreference='Stop'
$PSNativeCommandUseErrorActionPreference=$true
$Directory=[IO.Path]::GetFullPath($Directory)
$RepoRoot=(Resolve-Path (Join-Path $PSScriptRoot '..')).Path
if (Test-Path -LiteralPath $Directory) { throw 'Use a new session directory' }
[void][IO.Directory]::CreateDirectory($Directory)
# Reuse the repository's exact-PID attachment, compile and UIA dialog helpers.
# Load definitions only; the ordinary worker's per-case launch loop is excluded.
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
# Monitor startup can be delayed by concurrent compiler/linker work.
$definitions=$definitions.Replace('`$deadline = [DateTime]::UtcNow.AddSeconds(10)'.Replace('`',''), '`$deadline = [DateTime]::UtcNow.AddSeconds(60)'.Replace('`',''))
. ([scriptblock]::Create($definitions))
$SafeMode=$true
$OwnershipFile=Join-Path $Directory 'owned-excel.jsonl'
$HelperOwnershipFile=Join-Path $Directory 'owned-helpers.jsonl'
$script:SelectedCaseIds=@('office-project')
$script:GuardianControlSequence=0L
$controlFile=Join-Path $Directory 'guardian-control.json'
$eventsFile=Join-Path $Directory 'guardian-events.jsonl'
$readyFile=Join-Path $Directory 'guardian-ready.json'
$stopFile=Join-Path $Directory 'guardian-stop'
$identityFile=Join-Path $Directory 'excel-identity.json'
$guardian=$null; $excel=$null; $owned=$null; $script:DirectProcess=$null
$script:guardianGeneration=0

function Start-OfficeGuardian {
    $script:guardianGeneration++
    if (Test-Path -LiteralPath $stopFile) { Remove-Item -LiteralPath $stopFile }
    $script:eventsFile=Join-Path $Directory ("guardian-events-$script:guardianGeneration.jsonl")
    if (Test-Path -LiteralPath $readyFile) { Remove-Item -LiteralPath $readyFile }
    $arguments=@('-NoLogo','-NoProfile','-NonInteractive','-STA','-File',(Join-Path $PSScriptRoot 'excel-vba-oracle-guardian.ps1'),
        '-ExcelPid',[string]$owned.Id,'-ExcelIdentityFile',$identityFile,'-RunId',$RunId,'-CaseId','office-project',
        '-ControlFile',$controlFile,'-EventsFile',$eventsFile,'-ReadyFile',$readyFile,'-StopFile',$stopFile,'-MaxSeconds','3600')
    $script:guardian=Start-Process (Join-Path $PSHOME 'pwsh.exe') -ArgumentList $arguments -WindowStyle Hidden -PassThru `
        -RedirectStandardOutput (Join-Path $Directory ('guardian-'+[Guid]::NewGuid().ToString('N')+'.out')) `
        -RedirectStandardError (Join-Path $Directory ('guardian-'+[Guid]::NewGuid().ToString('N')+'.err'))
    [void](Add-HelperOwnershipRecord -Process $guardian -CaseId 'office-project')
    $script:guardianReady=Wait-GuardianReady -ReadyFile $readyFile -Process $guardian
}
function Set-OfficePhase([string]$Phase,[string]$Operation) {
    Assert-GuardianLive -Process $guardian -ReadyRecord $guardianReady -Phase $Phase
    $control=Set-GuardianControl -Path $controlFile -CaseId 'office-project' -OperationId $Operation -Phase $Phase -AllowDismiss $true
    Wait-GuardianOperationArmed -EventsFile $eventsFile -Control $control -Process $guardian -ReadyRecord $guardianReady
}
function Invoke-OfficeCompile($Workbook,[string]$ModuleName,[switch]$Force,[switch]$ShowWindow) {
    $project=$Workbook.VBProject
    $component=$project.VBComponents.Item($ModuleName)
    $excel.VBE.MainWindow.Visible=[bool]$ShowWindow
    if ($ShowWindow) { $component.CodeModule.CodePane.Show() }
    Start-Sleep -Milliseconds 200
    $control=Get-VbeCompileControl -Vbe $excel.VBE
    if (-not $control.Enabled -and $Force) {
        $text=[string]$component.CodeModule.Lines(1,$component.CodeModule.CountOfLines)
        $component.CodeModule.DeleteLines(1,$component.CodeModule.CountOfLines)
        [void]$component.CodeModule.AddFromString($text)
        Start-Sleep -Milliseconds 100
    }
    if (-not $control.Enabled) { throw "VBE Compile is disabled for $ModuleName" }
    $op='compile-'+[Guid]::NewGuid().ToString('N')
    $armed=Set-OfficePhase compile $op
    $control.Execute()
    $events=@(Wait-GuardianEventFlush -EventsFile $eventsFile -OperationId $op -ArmRecord $armed -InvocationCompletedUtc ([DateTime]::UtcNow) -CompileControl $control)
    $dialogs=@($events | Where-Object event_type -eq 'dialog-observation')
    if ($dialogs.Count -or $control.Enabled) { throw "VBE compile failed: $($dialogs | ConvertTo-Json -Depth 10 -Compress)" }
    [ordered]@{workbook=[string]$Workbook.FullName;module=$ModuleName;command_id=[int]$control.Id;passed=$true;forced=[bool]$Force;visible=[bool]$ShowWindow} |
        ConvertTo-Json -Compress | Add-Content -LiteralPath (Join-Path $Directory 'compiles.jsonl') -Encoding utf8NoBOM
    [void](Set-OfficePhase run ('run-'+[Guid]::NewGuid().ToString('N')))
    $excel.VBE.MainWindow.Visible=$false
    Release-ComObject $component
    Release-ComObject $project
}
try {
    $before=@(Get-ExcelProcessIds)
    $bootstrap=New-ExcelOracleBootstrapWorkbook -Path (Join-Path $Directory 'bootstrap.xlsx')
    $launched=Start-OwnedExcelApplication -ExcelExecutable (Get-ExcelExecutablePath) -BootstrapWorkbook $bootstrap -DiagnosticPath (Join-Path $Directory 'attachment.json') -NoDesktopInput
    $owned=$launched.process; $excel=$launched.application; $native=$launched.native_window
    if ($owned.Id -eq 3144 -or $owned.Id -in $before -or (Get-ExcelPidFromApplication $excel) -ne $owned.Id) { throw 'Excel ownership mismatch' }
    $identity=Add-OwnershipRecord -Process $owned -BeforePids $before -CaseId 'office-project'
    $identity | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $identityFile -Encoding utf8NoBOM
    $excel.Visible=[bool]$Visible; $excel.DisplayAlerts=$false; $excel.EnableEvents=$false
    # Start a PID-scoped monitor for each job; idle Excel needs no polling.
    [ordered]@{run_id=$RunId;pid=$owned.Id;hwnd=[string]$excel.Hwnd;launches=1;automation_security=[int]$excel.AutomationSecurity} |
        ConvertTo-Json | Set-Content -LiteralPath (Join-Path $Directory 'ready.json') -Encoding utf8NoBOM
    [void][IO.Directory]::CreateDirectory((Join-Path $Directory 'jobs'))
    [void][IO.Directory]::CreateDirectory((Join-Path $Directory 'results'))
    $processed=[Collections.Generic.HashSet[string]]::new()
    while (-not (Test-Path -LiteralPath (Join-Path $Directory 'stop'))) {
        foreach ($jobFile in @(Get-ChildItem -LiteralPath (Join-Path $Directory 'jobs') -Filter '*.json' | Sort-Object Name)) {
            if (-not $processed.Add($jobFile.Name)) { continue }
            $job=Get-Content -Raw -LiteralPath $jobFile.FullName | ConvertFrom-Json
            if ([string]$job.id -notmatch '^[a-zA-Z0-9_-]+$') { throw 'Invalid oracle job id' }
            $record=[ordered]@{id=[string]$job.id;compile_status='not-run';run_status='not-run';value=$null;error=$null;pid=$owned.Id;launches=1}
            $book=$null; $component=$null
            try {
                Start-OfficeGuardian
                [void](Set-OfficePhase run ('job-ready-'+$job.id))
                $book=$excel.Workbooks.Add()
                if ($job.PSObject.Properties.Name -contains 'modules') {
                    foreach ($path in $job.modules) {
                        $imported=$book.VBProject.VBComponents.Import([IO.Path]::GetFullPath([string]$path))
                        Release-ComObject $imported
                    }
                }
                $component=$book.VBProject.VBComponents.Add(1)
                $component.Name='OracleSelfTest'
                $source=Get-Content -Raw -LiteralPath ([string]$job.source_path)
                [void]$component.CodeModule.AddFromString($source)
                Invoke-OfficeCompile -Workbook $book -ModuleName 'OracleSelfTest' -ShowWindow:$Visible
                $record.compile_status='ok'
                $record.value=[string]$excel.Run("'"+$book.Name+"'!OracleSelfTest.ProbeNativeExports")
                $record.run_status='ok'
                if ($job.PSObject.Properties.Name -contains 'expected' -and $record.value -cne [string]$job.expected) { throw "Oracle mismatch: $($record.value)" }
            } catch {
                $record.error=[string]$_
                if ($record.compile_status -eq 'not-run') { $record.compile_status='failed' }
                else { $record.run_status='failed' }
            } finally {
                Release-ComObject $component
                if ($book) { $book.Close($false); Release-ComObject $book }
                $record | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath (Join-Path $Directory ('results/'+$job.id+'.json')) -Encoding utf8NoBOM
                if ($guardian -and -not $guardian.HasExited) {
                    New-Item -ItemType File -Path $stopFile -Force | Out-Null
                    if (-not $guardian.WaitForExit(5000)) { Stop-Process -InputObject $guardian; [void]$guardian.WaitForExit(5000) }
                }
            }
        }
        Start-Sleep -Milliseconds 200
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
            if ($live -and $live.Id -ne 3144 -and $live.StartTime.ToUniversalTime() -eq $cleanup.StartTime.ToUniversalTime() -and $live.Path -eq $cleanup.Path) { Stop-Process -Id $live.Id }
        }
        [ordered]@{pid=$cleanup.Id;exited=$cleanup.HasExited;launches=1} | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $Directory 'cleanup.json')
    }
}
