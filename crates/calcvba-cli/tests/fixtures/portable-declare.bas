Option Explicit

Private Declare PtrSafe Sub Pause Lib "KERNEL32.DLL" Alias "Sleep" (ByVal milliseconds As Long)
Private Declare PtrSafe Function GetCurrentProcessId Lib "kernel32" () As Long
Private Declare PtrSafe Function GetTickCount Lib "kernel32" () As Long
Private Declare PtrSafe Function GetTickCount64 Lib "kernel32" () As LongLong
Private Declare PtrSafe Function QueryPerformanceCounter Lib "kernel32" (ByRef ticks As LongLong) As Long
Private Declare PtrSafe Function QueryPerformanceFrequency Lib "kernel32" (ByRef ticksPerSecond As Currency) As Long
Private Declare PtrSafe Function lstrlenW Lib "kernel32" (ByVal text As LongPtr) As Long
Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef dest As Any, ByRef source As Any, ByVal count As LongPtr)

Public Sub Main()
    Debug.Print Probe()
End Sub

Public Function Probe() As String
    Dim pid As Long, tick32 As Long, tick64 As LongLong
    Dim before As LongLong, after As LongLong, frequency As Currency
    Dim frequencyOk As Long, beforeOk As Long, afterOk As Long
    Dim message As String, data(0 To 2) As Byte

    pid = GetCurrentProcessId()
    frequencyOk = QueryPerformanceFrequency(frequency)
    beforeOk = QueryPerformanceCounter(before)
    tick32 = GetTickCount()
    tick64 = GetTickCount64()
    Pause 50
    afterOk = QueryPerformanceCounter(after)
    Probe = CStr(pid <> 0 And pid = GetCurrentProcessId()) & "|" & _
            CStr(frequencyOk = 1 And frequency > 0@) & "|" & _
            CStr(beforeOk = 1 And afterOk = 1 And after > before) & "|" & _
            CStr(GetTickCount() <> tick32 And GetTickCount64() > tick64)
    message = "a" & ChrW(&HD83D) & ChrW(&HDE00) & ChrW(0) & "b"
    Probe = Probe & "|" & CStr(lstrlenW(StrPtr(message))) & "|" & CStr(lstrlenW(0))
    data(0) = 41
    data(1) = 42
    data(2) = 43
    CopyMemory data(1), data(0), 2
    Probe = Probe & "|" & CStr(data(0)) & "|" & CStr(data(1)) & "|" & CStr(data(2))
End Function
