Attribute VB_Name = "Fixture"
' Own fixture for capability triage: one function per capability, each with and without
' On Error Resume Next, plus one function Main never calls. The functions touch nothing that
' exists: under Safe the host is never reached, and under Full the paths and names are absent.
Option Explicit

Declare PtrSafe Function NativeGetTickCount Lib "kernel32" Alias "GetTickCount" () As Long

Public Function FsRead(ByVal p As String) As String
    Dim x As Variant
    x = FileLen(p)
    FsRead = "done"
End Function

Public Function FsWrite(ByVal p As String) As String
    Kill p
    FsWrite = "done"
End Function

Public Function RunProcess() As String
    Dim x As Variant
    x = Shell("echo capability-fixture")
    RunProcess = "done"
End Function

Public Function ReadEnv() As String
    Dim x As Variant
    x = Environ("CAPABILITY_FIXTURE_VAR")
    ReadEnv = "done"
End Function

Public Function ReadSetting() As String
    Dim x As Variant
    x = GetSetting("CapabilityFixture", "Section", "Key", "default")
    ReadSetting = "done"
End Function

Public Function CallNative() As String
    Dim x As Variant
    x = NativeGetTickCount()
    CallNative = "done"
End Function

Public Function Activate() As String
    Dim o As Object
    Set o = CreateObject("WinHttp.WinHttpRequest.5.1")
    Activate = "done"
End Function

Public Function PressKeys() As String
    SendKeys "x", False
    PressKeys = "done"
End Function

Public Function FsReadResume(ByVal p As String) As String
    On Error Resume Next
    Dim x As Variant
    x = FileLen(p)
    FsReadResume = "done"
End Function

Public Function FsWriteResume(ByVal p As String) As String
    On Error Resume Next
    Kill p
    FsWriteResume = "done"
End Function

Public Function RunProcessResume() As String
    On Error Resume Next
    Dim x As Variant
    x = Shell("echo capability-fixture")
    RunProcessResume = "done"
End Function

Public Function ReadEnvResume() As String
    On Error Resume Next
    Dim x As Variant
    x = Environ("CAPABILITY_FIXTURE_VAR")
    ReadEnvResume = "done"
End Function

Public Function ReadSettingResume() As String
    On Error Resume Next
    Dim x As Variant
    x = GetSetting("CapabilityFixture", "Section", "Key", "default")
    ReadSettingResume = "done"
End Function

Public Function CallNativeResume() As String
    On Error Resume Next
    Dim x As Variant
    x = NativeGetTickCount()
    CallNativeResume = "done"
End Function

Public Function ActivateResume() As String
    On Error Resume Next
    Dim o As Object
    Set o = CreateObject("WinHttp.WinHttpRequest.5.1")
    ActivateResume = "done"
End Function

Public Function PressKeysResume() As String
    On Error Resume Next
    SendKeys "x", False
    PressKeysResume = "done"
End Function

' Never called: triage lists its ProgID, the audit log never sees it.
Public Function NeverCalled() As String
    Dim o As Object
    Set o = CreateObject("Uncalled.Class")
    NeverCalled = "done"
End Function

Public Sub Main()
    Dim s As String
    Dim p As String
    p = "/capability-fixture/does-not-exist.txt"
    On Error Resume Next
    s = FsRead(p)
    s = FsWrite(p)
    s = RunProcess()
    s = ReadEnv()
    s = ReadSetting()
    s = CallNative()
    s = Activate()
    s = PressKeys()
    s = FsReadResume(p)
    s = FsWriteResume(p)
    s = RunProcessResume()
    s = ReadEnvResume()
    s = ReadSettingResume()
    s = CallNativeResume()
    s = ActivateResume()
    s = PressKeysResume()
End Sub
