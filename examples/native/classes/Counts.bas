Attribute VB_Name = "Counts"
Option Explicit
Public created As Long
Public destroyed As Long
Public Function ClassSuite() As String
    Dim first As Counter, same As Object, lazy As New Counter
    Dim delta As Long, answer As Long
    On Error GoTo Failed
10 created = 0
20 destroyed = 0
30 Set first = New Counter
40 first.Value = 7
50 delta = 3
60 answer = first.Bump(delta)
70 Set same = first
80 ClassSuite = CStr(answer) & ":" & CStr(delta) & ":" & CStr(first Is same) & ":" & CStr(TypeOf first Is Counter)
90 ClassSuite = ClassSuite & ":" & TypeName(same)
100 lazy.Value = 11
110 Set lazy = Nothing
120 ClassSuite = ClassSuite & ":" & CStr(lazy.Value)
130 Set first = Nothing
140 Set same = Nothing
150 Set lazy = Nothing
160 ClassSuite = ClassSuite & ":" & CStr(created) & ":" & CStr(destroyed)
    Exit Function
Failed:
    ClassSuite="fault:" & CStr(Erl) & ":" & CStr(Err.Number) & ":" & Err.Description
End Function
Public Sub Main()
    Dim observed As String
    observed = ClassSuite()
End Sub
