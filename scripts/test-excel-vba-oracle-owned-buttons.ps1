$ErrorActionPreference = "Stop"
Set-StrictMode -Version Latest
if (-not $IsWindows) { Write-Output "test-excel-vba-oracle-owned-buttons: Windows only"; return }

function Assert-OwnedButton {
    param([bool]$Condition, [string]$Message)
    if (-not $Condition) { throw "test-excel-vba-oracle-owned-buttons: $Message" }
}

# Compile the production native implementation, without running the guardian or Excel.
$source = Get-Content -Raw -LiteralPath (Join-Path $PSScriptRoot "excel-vba-oracle-guardian.ps1")
$native = [regex]::Match($source, '(?s)Add-Type @"\r?\n(using System;.*?public static class ExcelGuardianOwnedWindows.*?\r?\n})\r?\n"@')
Assert-OwnedButton $native.Success "production native class was not found"
if (-not ("ExcelGuardianOwnedWindows" -as [type])) { Add-Type -TypeDefinition $native.Groups[1].Value }
Add-Type @"
using System;
using System.Runtime.InteropServices;
using System.Text;
public static class OwnedButtonTestNative {
    [DllImport("user32.dll", CharSet=CharSet.Unicode)] public static extern IntPtr FindWindow(string kind, string title);
    [DllImport("user32.dll", CharSet=CharSet.Unicode)] public static extern IntPtr FindWindowEx(IntPtr parent, IntPtr after, string kind, string title);
    [DllImport("user32.dll", CharSet=CharSet.Unicode)] public static extern bool SetWindowText(IntPtr hwnd, string text);
    [DllImport("user32.dll")] public static extern bool EnableWindow(IntPtr hwnd, bool enabled);
}
"@
$fixtureSource = @'
using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;
public static class OwnedButtonDialogFixture {
    [DllImport("user32.dll", CharSet=CharSet.Unicode)] private static extern IntPtr FindWindow(string kind, string title);
    [DllImport("user32.dll", CharSet=CharSet.Unicode)] private static extern IntPtr CreateWindowEx(int ex, string kind, string text, uint style, int x, int y, int w, int h, IntPtr parent, IntPtr menu, IntPtr instance, IntPtr param);
    public static void Run(string caption, bool duplicate) {
        var timer = new Timer();
        timer.Interval = 25;
        timer.Tick += (sender, args) => {
            IntPtr dialog = FindWindow("#32770", caption);
            if (dialog == IntPtr.Zero) return;
            timer.Stop();
            if (duplicate) CreateWindowEx(0, "Button", "OK", 0x50000000, 5, 5, 50, 20, dialog, new IntPtr(99), IntPtr.Zero, IntPtr.Zero);
        };
        timer.Start();
        MessageBox.Show("Compile error: Sub or Function not defined", caption, MessageBoxButtons.OK);
        timer.Dispose();
    }
}
'@
$message = "Compile error: Sub or Function not defined"
$shell = Join-Path $env:WINDIR "System32/WindowsPowerShell/v1.0/powershell.exe"
foreach ($duplicate in @($false, $true)) {
    $caption = "OxVba owned button test $([Guid]::NewGuid())"
    $code = "Add-Type -AssemblyName System.Windows.Forms; Add-Type -TypeDefinition '$($fixtureSource.Replace("'", "''"))' -ReferencedAssemblies System.Windows.Forms; [OwnedButtonDialogFixture]::Run('$caption', `$$duplicate)"
    $encoded = [Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($code))
    $helper = Start-Process -FilePath $shell -ArgumentList @("-NoProfile", "-NonInteractive", "-STA", "-EncodedCommand", $encoded) -WindowStyle Hidden -PassThru
    $started = $helper.StartTime
    try {
        $deadline = [DateTime]::UtcNow.AddSeconds(15)
        do {
            $dialog = [OwnedButtonTestNative]::FindWindow("#32770", $caption)
            if ($dialog -ne [IntPtr]::Zero -and [ExcelGuardianOwnedWindows]::IsOwnedVisibleDialog($dialog, [uint32]$helper.Id)) { break }
            Start-Sleep -Milliseconds 50
        } while (-not $helper.HasExited -and [DateTime]::UtcNow -lt $deadline)
        Assert-OwnedButton ([ExcelGuardianOwnedWindows]::IsOwnedVisibleDialog($dialog, [uint32]$helper.Id)) "owned fixture did not appear"
        if ($duplicate) {
            do {
                $unique = [ExcelGuardianOwnedWindows]::FindUniqueButton($dialog, [uint32]$helper.Id, "OK")
                if ($unique -eq [IntPtr]::Zero) { break }
                Start-Sleep -Milliseconds 25
            } while ([DateTime]::UtcNow -lt $deadline)
            Assert-OwnedButton ($unique -eq [IntPtr]::Zero) "duplicate fixture must expose two OK buttons"
            $clicked = [ExcelGuardianOwnedWindows]::TryClickObservedButton($dialog, [uint32]$helper.Id, $caption, @($message), @("OK"))
            Assert-OwnedButton ([string]::IsNullOrEmpty($clicked)) "duplicate OK buttons must not be clicked"
        }
        else {
            foreach ($probe in @(
                @{ pid = [uint32]$PID; title = $caption; text = $message; button = "OK" },
                @{ pid = [uint32]$helper.Id; title = "wrong title"; text = $message; button = "OK" },
                @{ pid = [uint32]$helper.Id; title = $caption; text = "wrong message"; button = "OK" },
                @{ pid = [uint32]$helper.Id; title = $caption; text = $message; button = "End" }
            )) {
                $clicked = [ExcelGuardianOwnedWindows]::TryClickObservedButton($dialog, $probe.pid, $probe.title, @($probe.text), @($probe.button))
                Assert-OwnedButton ([string]::IsNullOrEmpty($clicked)) "mismatched owner/title/message/button must not be clicked"
                Assert-OwnedButton (-not [ExcelGuardianOwnedWindows]::DialogGone($dialog, [uint32]$helper.Id)) "rejected probe dismissed its dialog"
            }
            $button = [ExcelGuardianOwnedWindows]::FindUniqueButton($dialog, [uint32]$helper.Id, "OK")
            [void][OwnedButtonTestNative]::EnableWindow($button, $false)
            $clicked = [ExcelGuardianOwnedWindows]::TryClickObservedButton($dialog, [uint32]$helper.Id, $caption, @($message), @("OK"))
            Assert-OwnedButton ([string]::IsNullOrEmpty($clicked)) "disabled button must not be clicked"
            [void][OwnedButtonTestNative]::EnableWindow($button, $true)
            $text = [OwnedButtonTestNative]::FindWindowEx($dialog, [IntPtr]::Zero, "Static", $message)
            Assert-OwnedButton ($text -ne [IntPtr]::Zero) "fixture message control was not found"
            [void][OwnedButtonTestNative]::SetWindowText($text, "$message -- changed message")
            $clicked = [ExcelGuardianOwnedWindows]::TryClickObservedButton($dialog, [uint32]$helper.Id, $caption, @($message), @("OK"))
            Assert-OwnedButton ([string]::IsNullOrEmpty($clicked)) "appended message must not match the earlier observation"
            [void][OwnedButtonTestNative]::SetWindowText($text, $message)
            $clicked = [ExcelGuardianOwnedWindows]::TryClickObservedButton($dialog, [uint32]$helper.Id, $caption, @($message), @("OK"))
            Assert-OwnedButton ($clicked -ceq "OK" -and [ExcelGuardianOwnedWindows]::DialogGone($dialog, [uint32]$helper.Id)) "exact observed button must close the owned dialog"
        }
    }
    finally {
        # Retain and verify the launched instance; never search for or stop Excel.
        $helper.Refresh()
        if (-not $helper.HasExited -and $helper.StartTime -eq $started -and $helper.ProcessName -ceq "powershell") {
            $helper.Kill()
            [void]$helper.WaitForExit(5000)
        }
        $helper.Dispose()
    }
}
Write-Output "test-excel-vba-oracle-owned-buttons: PASS"
