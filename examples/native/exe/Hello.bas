Attribute VB_Name = "Hello"
Option Explicit
Public observed As String
Public Sub Main()
    Dim values As Scripting.Dictionary
    Set values = New Scripting.Dictionary
    values.Add "base", 40
    observed = "native-exe:" & CStr(CLng(values.Item("base")) + NextValue())
    observed = observed & ":" & CStr(CLng(values.Item("base")) + NextValue()) & ":" & CStr(values.Count)
    Debug.Print observed
End Sub
Private Function NextValue() As Long
    Static value As Long
    value = value + 1
    NextValue = value
End Function
