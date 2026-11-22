Option Explicit

Public Sub Main()
    Debug.Print Probe()
End Sub

Public Function Probe() As String
    Dim re As Object
    Set re = CreateObject("vBsCrIpT.rEgExP")
    re.Pattern = "ab+"
    Probe = CStr(re.Test("xxabbbxx")) & "|" & CStr(re.Test("xxaxx"))
End Function
