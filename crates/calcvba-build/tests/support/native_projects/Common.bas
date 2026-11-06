Attribute VB_Name = "Common"
Option Explicit
Private total As Long
Private ended As Long
Private numbers(1 To 2) As Long
Public Sub Reset()
    total = 0
    ended = 0
End Sub
Public Function Increment() As Long
    total = total + 1
    Increment = total
End Function
Public Function TotalValue() As Long
    TotalValue = total
End Function
Public Function EndedValue() As Long
    EndedValue = ended
End Function
Public Sub Terminated()
    ended = ended + 1
End Sub
Public Function InitialState() As String
    InitialState = CStr(UBound(numbers)) & ":" & CStr(numbers(1))
End Function
Public Function MakeToken() As Token
    Set MakeToken = New Token
End Function
Public Sub Fail()
    Err.Raise 6
End Sub
