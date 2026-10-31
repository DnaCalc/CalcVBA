param(
    [int]$ExcelPid,
    [string]$ExcelIdentityFile,
    [string]$RunId,
    [string]$CaseId,
    [string]$ControlFile,
    [string]$EventsFile,
    [string]$ReadyFile,
    [string]$StopFile,
    [ValidateRange(25, 5000)][int]$PollMilliseconds = 100,
    [ValidateRange(1, 3600)][int]$MaxSeconds = 300,
    [switch]$PolicySelfTest
)

$ErrorActionPreference = "Stop"
Set-StrictMode -Version Latest
. (Join-Path $PSScriptRoot "excel-vba-oracle-contract.ps1")

function Assert-GuardianPolicy {
    param(
        [Parameter(Mandatory = $true)][string]$ExpectedKind,
        [Parameter(Mandatory = $true)][string]$ExpectedDisposition,
        [Parameter(Mandatory = $true)]$Actual
    )
    if ($Actual.kind -ne $ExpectedKind -or $Actual.disposition -ne $ExpectedDisposition) {
        throw "excel-vba-oracle-guardian: expected $ExpectedKind/$ExpectedDisposition, got $($Actual.kind)/$($Actual.disposition)"
    }
}

if ($PolicySelfTest) {
    Assert-GuardianPolicy "compile-error" "capture-then-dismiss" (Get-ExcelOracleDialogPolicy -Phase compile -WindowTitle "Microsoft Visual Basic for Applications" -Texts @("Compile error: Sub or Function not defined") -Buttons @("OK"))
    Assert-GuardianPolicy "runtime-error" "capture-then-dismiss" (Get-ExcelOracleDialogPolicy -Phase run -Texts @("Run-time error '13': Type mismatch") -Buttons @("End", "Debug", "Help"))
    Assert-GuardianPolicy "security-or-trust" "block-no-dismiss" (Get-ExcelOracleDialogPolicy -Phase compile -Texts @("Macros in this project are disabled") -Buttons @("Enable Content"))
    Assert-GuardianPolicy "ambiguous-macro-failure" "capture-then-dismiss" (Get-ExcelOracleDialogPolicy -Phase run -Texts @("Cannot run the macro. The macro may not be available or all macros may be disabled.") -Buttons @("OK"))
    Assert-GuardianPolicy "unrecognized-modal" "block-no-dismiss" (Get-ExcelOracleDialogPolicy -Phase run -Texts @("Unexpected prompt") -Buttons @("Yes", "No"))
    Assert-GuardianPolicy "unrecognized-modal" "block-no-dismiss" (Get-ExcelOracleDialogPolicy -Phase run -Texts @("Compile error: Sub or Function not defined") -Buttons @("OK"))
    Assert-GuardianPolicy "unrecognized-modal" "block-no-dismiss" (Get-ExcelOracleDialogPolicy -Phase compile -Texts @("Run-time error '13': Type mismatch") -Buttons @("End"))
    Write-Output "excel-vba-oracle-guardian: policy self-test passed"
    exit 0
}

foreach ($required in @{
        ExcelPid = $ExcelPid
        ExcelIdentityFile = $ExcelIdentityFile
        RunId = $RunId
        CaseId = $CaseId
        ControlFile = $ControlFile
        EventsFile = $EventsFile
        ReadyFile = $ReadyFile
        StopFile = $StopFile
    }.GetEnumerator()) {
    if ($required.Key -eq "ExcelPid") {
        if ([int]$required.Value -le 0) { throw "excel-vba-oracle-guardian: ExcelPid is required" }
    }
    elseif ([string]::IsNullOrWhiteSpace([string]$required.Value)) {
        throw "excel-vba-oracle-guardian: $($required.Key) is required"
    }
}

Add-Type -AssemblyName UIAutomationClient
Add-Type -AssemblyName UIAutomationTypes

if (-not (Test-Path -LiteralPath $ExcelIdentityFile)) {
    throw "excel-vba-oracle-guardian: Excel identity file does not exist"
}
$excelIdentity = ConvertFrom-ExcelOracleProcessIdentityJson `
    -Json (Get-Content -Raw -LiteralPath $ExcelIdentityFile) `
    -ExpectedSchema "oxvba.excel-vba-oracle-owned-process.v1"
if ([string]$excelIdentity.case_id -cne $CaseId -or
    -not (Test-ExcelOracleOwnedProcessRecord -Record $excelIdentity -BaselineExcelPids ([int[]]::new(0)) -RunId $RunId)) {
    throw "excel-vba-oracle-guardian: Excel ownership identity record did not validate"
}
$initialExcelProcess = Get-Process -Id $ExcelPid -ErrorAction Stop
if (-not (Test-ExcelOracleProcessIdentity -Record $excelIdentity -Process $initialExcelProcess -ExpectedProcessName "EXCEL" -RunId $RunId)) {
    throw "excel-vba-oracle-guardian: Excel PID/start/name/executable identity did not validate"
}

# Native enumeration also sees owned dialogs nested under the VBE UIA root.
# UIA desktop children alone omit those on some Office builds.
Add-Type @"
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;
public static class ExcelGuardianOwnedWindows {
    private delegate bool Visitor(IntPtr hwnd, IntPtr state);
    [DllImport("user32.dll")] private static extern bool EnumWindows(Visitor callback, IntPtr state);
    [DllImport("user32.dll")] private static extern bool EnumChildWindows(IntPtr parent, Visitor callback, IntPtr state);
    [DllImport("user32.dll")] private static extern bool IsWindowVisible(IntPtr hwnd);
    [DllImport("user32.dll")] private static extern bool IsWindow(IntPtr hwnd);
    [DllImport("user32.dll")] private static extern bool IsWindowEnabled(IntPtr hwnd);
    [DllImport("user32.dll")] private static extern bool IsChild(IntPtr parent, IntPtr child);
    [DllImport("user32.dll", CharSet=CharSet.Unicode)] private static extern IntPtr SendMessageTimeout(IntPtr hwnd, uint message, IntPtr wparam, IntPtr lparam, uint flags, uint timeout, out UIntPtr result);
    [DllImport("user32.dll")] private static extern uint GetWindowThreadProcessId(IntPtr hwnd, out uint pid);
    [DllImport("user32.dll", CharSet=CharSet.Unicode)] private static extern int GetWindowText(IntPtr hwnd, StringBuilder text, int size);
    [DllImport("user32.dll", CharSet=CharSet.Unicode)] private static extern int GetClassName(IntPtr hwnd, StringBuilder text, int size);
    public static string[] ChildText(IntPtr parent, uint expectedPid, bool buttons) {
        uint parentPid; GetWindowThreadProcessId(parent, out parentPid);
        var values = new List<string>();
        if (parentPid != expectedPid) return values.ToArray();
        EnumChildWindows(parent, (child, state) => {
            uint pid; GetWindowThreadProcessId(child, out pid);
            if (pid != expectedPid || !IsWindowVisible(child)) return true;
            var name = new StringBuilder(128); GetClassName(child, name, name.Capacity);
            if (name.ToString() == (buttons ? "Button" : "Static")) {
                var text = new StringBuilder(4096); GetWindowText(child, text, text.Capacity);
                if (text.Length != 0) values.Add(text.ToString().Replace("&", ""));
            }
            return true;
        }, IntPtr.Zero);
        return values.ToArray();
    }
    private static string WindowText(IntPtr hwnd) {
        var text = new StringBuilder(4096); GetWindowText(hwnd, text, text.Capacity); return text.ToString();
    }
    private static string WindowClass(IntPtr hwnd) {
        var text = new StringBuilder(128); GetClassName(hwnd, text, text.Capacity); return text.ToString();
    }
    private static string Normalize(string value) {
        return System.Text.RegularExpressions.Regex.Replace((value ?? "").Replace("&", ""), @"\s+", " ").Trim();
    }
    public static bool IsOwnedVisibleDialog(IntPtr hwnd, uint expectedPid) {
        uint pid; GetWindowThreadProcessId(hwnd, out pid);
        string kind = WindowClass(hwnd);
        return hwnd != IntPtr.Zero && IsWindow(hwnd) && IsWindowVisible(hwnd) && pid == expectedPid &&
            (kind == "#32770" || kind == "NUIDialog");
    }
    public static bool DialogGone(IntPtr hwnd, uint expectedPid) {
        uint pid; GetWindowThreadProcessId(hwnd, out pid);
        return !IsWindow(hwnd) || !IsWindowVisible(hwnd) || pid != expectedPid;
    }
    public static bool MatchesObservation(IntPtr parent, uint expectedPid, string title, string[] texts) {
        if (!IsOwnedVisibleDialog(parent, expectedPid) || WindowText(parent) != title || texts == null || texts.Length == 0) return false;
        foreach (string text in texts) if (String.IsNullOrWhiteSpace(text)) return false;
        string actual = Normalize(String.Join("\n", ChildText(parent, expectedPid, false)));
        return actual.Length != 0 && actual == Normalize(String.Join("\n", texts));
    }
    public static bool IsMatchingButton(IntPtr parent, IntPtr button, uint expectedPid, string label) {
        uint pid; GetWindowThreadProcessId(button, out pid);
        return pid == expectedPid && IsChild(parent, button) && IsWindowVisible(button) && IsWindowEnabled(button) &&
            WindowClass(button) == "Button" && String.Equals(Normalize(WindowText(button)), Normalize(label), StringComparison.OrdinalIgnoreCase);
    }
    public static IntPtr FindUniqueButton(IntPtr parent, uint expectedPid, string label) {
        var buttons = new List<IntPtr>();
        EnumChildWindows(parent, (child, state) => {
            if (IsMatchingButton(parent, child, expectedPid, label)) buttons.Add(child);
            return true;
        }, IntPtr.Zero);
        return buttons.Count == 1 ? buttons[0] : IntPtr.Zero;
    }
    // No desktop-wide button search or keyboard input: exactly one visible enabled Button
    // under the captured owned dialog, with unchanged title/message, must match the policy.
    public static string TryClickObservedButton(IntPtr parent, uint expectedPid, string title, string[] texts, string[] preferred) {
        if (!MatchesObservation(parent, expectedPid, title, texts)) return null;
        foreach (string label in preferred) {
            IntPtr button = FindUniqueButton(parent, expectedPid, label);
            if (button == IntPtr.Zero) continue;
            if (!MatchesObservation(parent, expectedPid, title, texts) || !IsMatchingButton(parent, button, expectedPid, label)) return null;
            UIntPtr result;
            if (SendMessageTimeout(button, 0x00F5, IntPtr.Zero, IntPtr.Zero, 2, 1000, out result) == IntPtr.Zero) return null;
            for (int i = 0; i < 40; i++) {
                if (DialogGone(parent, expectedPid)) return label;
                System.Threading.Thread.Sleep(25);
            }
            return null;
        }
        return null;
    }
    public static IntPtr[] Enumerate(uint expectedPid) {
        var handles = new HashSet<IntPtr>();
        EnumWindows((hwnd, state) => {
            uint pid; GetWindowThreadProcessId(hwnd, out pid);
            if (pid != expectedPid) return true;
            if (IsWindowVisible(hwnd)) handles.Add(hwnd);
            EnumChildWindows(hwnd, (child, unused) => {
                uint childPid; GetWindowThreadProcessId(child, out childPid);
                if (childPid != expectedPid || !IsWindowVisible(child)) return true;
                var name = new StringBuilder(128); GetClassName(child, name, name.Capacity);
                if (name.ToString() == "#32770" || name.ToString() == "NUIDialog") handles.Add(child);
                return true;
            }, IntPtr.Zero);
            return true;
        }, IntPtr.Zero);
        var result = new IntPtr[handles.Count]; handles.CopyTo(result); return result;
    }
}
"@

function Get-OwnedTopLevelWindows {
    $owned = [Collections.Generic.List[Windows.Automation.AutomationElement]]::new()
    foreach ($handle in [ExcelGuardianOwnedWindows]::Enumerate([uint32]$ExcelPid)) {
        try {
            $window = [Windows.Automation.AutomationElement]::FromHandle($handle)
            if ([int]$window.Current.ProcessId -eq $ExcelPid) { $owned.Add($window) }
        } catch { }
    }
    return @($owned)
}

function Get-ElementStrings {
    param(
        [Parameter(Mandatory = $true)][Windows.Automation.AutomationElement]$Root,
        [Parameter(Mandatory = $true)][Windows.Automation.ControlType]$ControlType
    )
    $condition = [Windows.Automation.PropertyCondition]::new(
        [Windows.Automation.AutomationElement]::ControlTypeProperty,
        $ControlType
    )
    $values = [Collections.Generic.List[string]]::new()
    $elements = $Root.FindAll([Windows.Automation.TreeScope]::Descendants, $condition)
    for ($elementIndex = 0; $elementIndex -lt $elements.Count; $elementIndex++) {
        $element = $elements.Item($elementIndex)
        try {
            $name = [string]$element.Current.Name
            if (-not [string]::IsNullOrWhiteSpace($name)) { $values.Add($name.Trim()) }
        }
        catch {
            # A stale descendant must not stop the guardian.
        }
    }
    if ($values.Count -eq 0) {
        # Some Office Win32 dialogs expose no UIA descendants. Read only their
        # PID-scoped Static/Button controls as a secondary observation.
        return @([ExcelGuardianOwnedWindows]::ChildText([IntPtr]$Root.Current.NativeWindowHandle, [uint32]$ExcelPid, $ControlType -eq [Windows.Automation.ControlType]::Button))
    }
    return @($values | Select-Object -Unique)
}

function Get-VbeSelection {
    $selection = [ordered]@{ selected_token = $null; expanded_line = $null }
    $documentCondition = [Windows.Automation.Condition]::TrueCondition
    foreach ($window in @(Get-OwnedTopLevelWindows)) {
        try {
            if ([string]$window.Current.Name -notlike "Microsoft Visual Basic for Applications*") { continue }
            $documents = $window.FindAll([Windows.Automation.TreeScope]::Descendants, $documentCondition)
            for ($documentIndex = 0; $documentIndex -lt $documents.Count; $documentIndex++) {
                $document = $documents.Item($documentIndex)
                try {
                $pattern = $document.GetCurrentPattern([Windows.Automation.TextPattern]::Pattern)
                $ranges = @($pattern.GetSelection())
                if ($ranges.Count -eq 0) { continue }
                $token = [string]$ranges[0].GetText(-1)
                $lineRange = $ranges[0].Clone()
                $lineRange.ExpandToEnclosingUnit([System.Windows.Automation.Text.TextUnit]::Line)
                $line = [string]$lineRange.GetText(-1)
                if (-not [string]::IsNullOrWhiteSpace($token) -or -not [string]::IsNullOrWhiteSpace($line)) {
                    $selection.selected_token = $token.Trim("`r", "`n")
                    $selection.expanded_line = $line.Trim("`r", "`n")
                    return [pscustomobject]$selection
                }
                } catch { continue }
            }
        }
        catch {
            # Stale windows and documents without TextPattern are nonfatal.
        }
    }
    return [pscustomobject]$selection
}

function Get-WindowIdentity {
    param([Parameter(Mandatory = $true)][Windows.Automation.AutomationElement]$Window)
    try { return [string]$Window.Current.NativeWindowHandle }
    catch { return "$($Window.Current.Name):$($Window.Current.AutomationId)" }
}

function Invoke-OwnedDialogButton {
    param(
        [Parameter(Mandatory = $true)][Windows.Automation.AutomationElement]$Window,
        [Parameter(Mandatory = $true)][string[]]$PreferredButtons,
        [Parameter(Mandatory = $true)][AllowEmptyString()][string]$ObservedTitle,
        [Parameter(Mandatory = $true)][string[]]$ObservedTexts
    )
    $script:lastDismissalMethod = $null
    $handle = [IntPtr]$Window.Current.NativeWindowHandle
    if (-not [ExcelGuardianOwnedWindows]::MatchesObservation($handle, [uint32]$ExcelPid, $ObservedTitle, $ObservedTexts)) { return $null }
    $buttonCondition = [Windows.Automation.PropertyCondition]::new(
        [Windows.Automation.AutomationElement]::ControlTypeProperty,
        [Windows.Automation.ControlType]::Button
    )
    $buttons = @()
    try {
        $foundButtons = $Window.FindAll([Windows.Automation.TreeScope]::Descendants, $buttonCondition)
        $buttons = @(for ($buttonIndex = 0; $buttonIndex -lt $foundButtons.Count; $buttonIndex++) { $foundButtons.Item($buttonIndex) })
    }
    catch { }
    foreach ($preferred in $PreferredButtons) {
        $nativeButton = [ExcelGuardianOwnedWindows]::FindUniqueButton($handle, [uint32]$ExcelPid, $preferred)
        if ($nativeButton -eq [IntPtr]::Zero) { continue }
        $button = @($buttons | Where-Object { [IntPtr]$_.Current.NativeWindowHandle -eq $nativeButton })
        if ($button.Count -ne 1) { continue }
        try {
            $invoke = $button[0].GetCurrentPattern([Windows.Automation.InvokePattern]::Pattern)
            if (-not [ExcelGuardianOwnedWindows]::MatchesObservation($handle, [uint32]$ExcelPid, $ObservedTitle, $ObservedTexts) -or
                -not [ExcelGuardianOwnedWindows]::IsMatchingButton($handle, $nativeButton, [uint32]$ExcelPid, $preferred)) { return $null }
            $invoke.Invoke()
            for ($attempt = 0; $attempt -lt 40; $attempt++) {
                if ([ExcelGuardianOwnedWindows]::DialogGone($handle, [uint32]$ExcelPid)) {
                    $script:lastDismissalMethod = "uia-invoke"
                    return $preferred
                }
                Start-Sleep -Milliseconds 25
            }
        }
        catch { }
    }
    $native = [ExcelGuardianOwnedWindows]::TryClickObservedButton($handle, [uint32]$ExcelPid, $ObservedTitle, $ObservedTexts, $PreferredButtons)
    if (-not [string]::IsNullOrWhiteSpace($native)) { $script:lastDismissalMethod = "native-owned-button" }
    return $native
}

function Add-GuardianEvent {
    param([Parameter(Mandatory = $true)]$Event)
    $script:eventSequence++
    $Event.event_sequence = $script:eventSequence
    $bytes = [Text.UTF8Encoding]::new($false).GetBytes("$(($Event | ConvertTo-Json -Compress -Depth 8))`n")
    $stream = [IO.FileStream]::new(
        $EventsFile,
        [IO.FileMode]::Append,
        [IO.FileAccess]::Write,
        [IO.FileShare]::Read,
        4096,
        [IO.FileOptions]::WriteThrough
    )
    try {
        $stream.Write($bytes, 0, $bytes.Length)
        $stream.Flush($true)
    }
    finally { $stream.Dispose() }
}

$readyParent = Split-Path -Parent $ReadyFile
if ($readyParent) { New-Item -ItemType Directory -Force -Path $readyParent | Out-Null }
$eventParent = Split-Path -Parent $EventsFile
if ($eventParent) { New-Item -ItemType Directory -Force -Path $eventParent | Out-Null }

$selfProcess = Get-Process -Id $PID
$ready = [ordered]@{
    schema = "oxvba.excel-vba-oracle-guardian-ready.v1"
    run_id = $RunId
    excel_pid = $ExcelPid
    excel_process_start_utc = [string]$excelIdentity.process_start_utc
    excel_executable_path = [string]$excelIdentity.executable_path
    guardian_pid = $PID
    pid = $PID
    process_name = [string]$selfProcess.ProcessName
    process_start_utc = $selfProcess.StartTime.ToUniversalTime().ToString("o")
    executable_path = [string]$selfProcess.Path
    ready_utc = [DateTime]::UtcNow.ToString("o")
}
$ready | ConvertTo-Json -Depth 4 | Set-Content -LiteralPath $ReadyFile -Encoding utf8NoBOM

$seen = [Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
$invalidControls = [Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
$armedControls = [Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
$script:eventSequence = 0L
$lastControlSequence = 0L
$deadline = [DateTime]::UtcNow.AddSeconds($MaxSeconds)
while ([DateTime]::UtcNow -lt $deadline -and -not (Test-Path -LiteralPath $StopFile)) {
    $currentExcelProcess = Get-Process -Id $ExcelPid -ErrorAction SilentlyContinue
    if (-not $currentExcelProcess -or -not (Test-ExcelOracleProcessIdentity -Record $excelIdentity -Process $currentExcelProcess -ExpectedProcessName "EXCEL" -RunId $RunId)) { break }
    $control = $null
    if (Test-Path -LiteralPath $ControlFile) {
        # Permit atomic replacement while retaining this exact file instance.
        $controlStream = [IO.FileStream]::new($ControlFile, [IO.FileMode]::Open, [IO.FileAccess]::Read, ([IO.FileShare]::ReadWrite -bor [IO.FileShare]::Delete))
        $controlReader = [IO.StreamReader]::new($controlStream)
        try { $controlJson = $controlReader.ReadToEnd() }
        finally { $controlReader.Dispose() }
        $controlHash = Get-ExcelOracleSha256 -Text $controlJson
        $parsedControl = ConvertFrom-ExcelOracleGuardianControl -Json $controlJson -RunId $RunId
        if (@($parsedControl.errors).Count -gt 0) {
            if ($invalidControls.Add($controlHash)) {
                Add-GuardianEvent -Event ([ordered]@{
                    schema = "oxvba.excel-vba-oracle-control-status.v1"
                    event_type = "invalid-control"
                    run_id = $RunId
                    excel_pid = $ExcelPid
                    event_sequence = 0L
                    observed_utc = [DateTime]::UtcNow.ToString("o")
                    control_sha256 = $controlHash
                    errors = @($parsedControl.errors)
                    valid = $false
                })
            }
            # An invalid control never arms an operation and can never authorize a dismissal.
            Start-Sleep -Milliseconds $PollMilliseconds
            continue
        }
        $control = $parsedControl.control
    }
    if ($null -eq $control -or [string]::IsNullOrWhiteSpace([string]$control.operation_id)) {
        Start-Sleep -Milliseconds $PollMilliseconds
        continue
    }
    $controlKey = "$([string]$control.case_id)|$([string]$control.operation_id)|$([long]$control.sequence)"
    if ([long]$control.sequence -lt $lastControlSequence) {
        if ($invalidControls.Add("stale:$controlKey")) {
            Add-GuardianEvent -Event ([ordered]@{
                schema = "oxvba.excel-vba-oracle-control-status.v1"; event_type = "invalid-control"; run_id = $RunId
                excel_pid = $ExcelPid
                event_sequence = 0L; observed_utc = [DateTime]::UtcNow.ToString("o"); control_sha256 = Get-ExcelOracleSha256 -Text ($control | ConvertTo-Json -Compress)
                errors = @("control sequence regressed"); valid = $false
            })
        }
        Start-Sleep -Milliseconds $PollMilliseconds
        continue
    }
    if ($armedControls.Add($controlKey)) {
        if ([long]$control.sequence -le $lastControlSequence) {
            Add-GuardianEvent -Event ([ordered]@{
                schema = "oxvba.excel-vba-oracle-control-status.v1"; event_type = "invalid-control"; run_id = $RunId
                excel_pid = $ExcelPid
                event_sequence = 0L; observed_utc = [DateTime]::UtcNow.ToString("o"); control_sha256 = Get-ExcelOracleSha256 -Text ($control | ConvertTo-Json -Compress)
                errors = @("control sequence was reused"); valid = $false
            })
            Start-Sleep -Milliseconds $PollMilliseconds
            continue
        }
        $lastControlSequence = [long]$control.sequence
        Add-GuardianEvent -Event ([ordered]@{
            schema = "oxvba.excel-vba-oracle-operation-state.v1"
            event_type = "operation-armed"
            run_id = $RunId
            case_id = [string]$control.case_id
            excel_pid = $ExcelPid
            operation_id = [string]$control.operation_id
            phase = [string]$control.phase
            control_sequence = [long]$control.sequence
            event_sequence = 0L
            observed_utc = [DateTime]::UtcNow.ToString("o")
        })
    }


    foreach ($window in @(Get-OwnedTopLevelWindows)) {
        try {
            $isModal = $false
            try {
                $windowPattern = $window.GetCurrentPattern([Windows.Automation.WindowPattern]::Pattern)
                $isModal = [bool]$windowPattern.Current.IsModal
            }
            catch { }
            $className = [string]$window.Current.ClassName
            $windowTitle = [string]$window.Current.Name
            $windowHandle = Get-WindowIdentity $window
            $observedProcessId = [int]$window.Current.ProcessId
            # Walking the entire Excel/VBE application tree can block UIA long
            # enough to miss the next modal. Their owned dialog HWNDs are
            # enumerated separately; inspect those, not the ribbon/editor tree.
            if (-not $isModal -and $className -in @("XLMAIN", "wndclass_desked_gsk")) {
                $texts = @()
                $buttons = @()
            } else {
                $texts = @(Get-ElementStrings -Root $window -ControlType ([Windows.Automation.ControlType]::Text))
                $buttons = @(Get-ElementStrings -Root $window -ControlType ([Windows.Automation.ControlType]::Button))
            }
            $policy = Get-ExcelOracleDialogPolicy -Phase ([string]$control.phase) -WindowTitle $windowTitle -Texts $texts -Buttons $buttons
            # VBE compile dialogs are not consistently exposed as IsModal/#32770 on
            # all Office builds. Recognized dialog text is authoritative. Unknown
            # non-modal application windows are recorded for the bounded audit but
            # ignored for policy/action purposes.
            $consideredDialog = $policy.kind -ne "unrecognized-modal" -or $isModal -or $className -eq "#32770"
            $key = "$($control.operation_id)|$windowHandle|$($policy.kind)|$($texts -join '|')"
            if (-not $seen.Add($key)) { continue }

            $vbeSelection = if ($consideredDialog) { Get-VbeSelection } else { [pscustomobject]@{ selected_token = $null; expanded_line = $null } }
            $observationId = [Guid]::NewGuid().ToString("D")
            $observationEvent = [ordered]@{
                schema = "oxvba.excel-vba-oracle-window-observation.v1"
                event_type = if ($consideredDialog) { "dialog-observation" } else { "ignored-top-level-window" }
                observation_id = $observationId
                run_id = $RunId
                case_id = [string]$control.case_id
                operation_id = [string]$control.operation_id
                control_sequence = [long]$control.sequence
                event_sequence = 0L
                phase = [string]$control.phase
                excel_pid = $ExcelPid
                observed_process_id = $observedProcessId
                observed_utc = [DateTime]::UtcNow.ToString("o")
                capture_completed_utc = [DateTime]::UtcNow.ToString("o")
                window_title = $windowTitle
                window_class = $className
                window_handle = $windowHandle
                is_modal = $isModal
                considered_dialog = $consideredDialog
                dialog_text = $texts
                visible_buttons = $buttons
                selected_token = $vbeSelection.selected_token
                expanded_line = $vbeSelection.expanded_line
                classification = [string]$policy.kind
                disposition = [string]$policy.disposition
            }
            # The immutable observation must be durable before any owned UI is
            # changed. A later crash can therefore never dismiss without capture.
            Add-GuardianEvent -Event $observationEvent

            if ($consideredDialog -and $policy.disposition -eq "capture-then-dismiss" -and [bool]$control.allow_dismiss -and
                ($policy.kind -ne "compile-error" -or (-not [string]::IsNullOrWhiteSpace([string]$vbeSelection.selected_token) -and -not [string]::IsNullOrWhiteSpace([string]$vbeSelection.expanded_line)))) {
                $dismissedButton = Invoke-OwnedDialogButton -Window $window -PreferredButtons @($policy.preferred_buttons) -ObservedTitle $windowTitle -ObservedTexts $texts
                $dismissalEvent = [ordered]@{
                    schema = "oxvba.excel-vba-oracle-dismissal-result.v1"
                    event_type = "dismissal-result"
                    observation_id = $observationId
                    run_id = $RunId
                    case_id = [string]$control.case_id
                    operation_id = [string]$control.operation_id
                    control_sequence = [long]$control.sequence
                    event_sequence = 0L
                    phase = [string]$control.phase
                    excel_pid = $ExcelPid
                    window_handle = $windowHandle
                    attempted_utc = [DateTime]::UtcNow.ToString("o")
                    requested_buttons = @($policy.preferred_buttons)
                    dismissed_button = $dismissedButton
                    dismissal_method = $script:lastDismissalMethod
                    succeeded = -not [string]::IsNullOrWhiteSpace([string]$dismissedButton)
                }
                Add-GuardianEvent -Event $dismissalEvent
            }
        }
        catch {
            # Stale top-level UIA children are expected and nonfatal per element.
        }
    }
    Add-GuardianEvent -Event ([ordered]@{
        schema = "oxvba.excel-vba-oracle-operation-state.v1"
        event_type = "guardian-heartbeat"
        run_id = $RunId
        case_id = [string]$control.case_id
        excel_pid = $ExcelPid
        operation_id = [string]$control.operation_id
        phase = [string]$control.phase
        control_sequence = [long]$control.sequence
        event_sequence = 0L
        observed_utc = [DateTime]::UtcNow.ToString("o")
    })
    Start-Sleep -Milliseconds $PollMilliseconds
}

$controlledStopObserved = Test-Path -LiteralPath $StopFile
$finalExcelProcess = Get-Process -Id $ExcelPid -ErrorAction SilentlyContinue
$excelIdentityLiveAtStop = $false
if ($finalExcelProcess) {
    try { $excelIdentityLiveAtStop = Test-ExcelOracleProcessIdentity -Record $excelIdentity -Process $finalExcelProcess -ExpectedProcessName "EXCEL" -RunId $RunId }
    catch { $excelIdentityLiveAtStop = $false }
}
$exitReason = if ($controlledStopObserved -and $excelIdentityLiveAtStop) { "controlled-stop" }
    elseif (-not $excelIdentityLiveAtStop) { "excel-identity-lost" }
    else { "deadline" }
Add-GuardianEvent -Event ([ordered]@{
    schema = "oxvba.excel-vba-oracle-final-state.v1"
    event_type = "guardian-stopped"
    run_id = $RunId
    case_id = $CaseId
    excel_pid = $ExcelPid
    event_sequence = 0L
    observed_utc = [DateTime]::UtcNow.ToString("o")
    guardian_pid = $PID
    process_name = [string]$selfProcess.ProcessName
    process_start_utc = $selfProcess.StartTime.ToUniversalTime().ToString("o")
    executable_path = [string]$selfProcess.Path
    controlled_stop_observed = [bool]$controlledStopObserved
    excel_identity_live_at_stop = [bool]$excelIdentityLiveAtStop
    exit_reason = $exitReason
})
if ($exitReason -ne "controlled-stop") { exit 2 }
