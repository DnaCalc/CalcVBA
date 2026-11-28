Sub Main()
    Dim shellObject As Object
    Dim initial As String
    Dim expanded As String
    Set shellObject = CreateObject("WScript.Shell")
    initial = shellObject.CurrentDirectory
    shellObject.CurrentDirectory = "C:\Data"
    expanded = shellObject.ExpandEnvironmentStrings("%TEMP%\file.txt|%CALCVBA_PORTABLE_UNMAPPED%")
    Debug.Print CStr(shellObject.CurrentDirectory = CurDir$) & "|" & CStr(CurDir$ = "C:\Data")
    Debug.Print CStr(expanded = Environ$("TEMP") & "\file.txt|%CALCVBA_PORTABLE_UNMAPPED%")
    shellObject.CurrentDirectory = initial
End Sub
