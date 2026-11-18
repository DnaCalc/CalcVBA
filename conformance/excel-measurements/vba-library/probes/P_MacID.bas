Attribute VB_Name = "P_MacID"
Option Explicit
' MacID (Windows) (VBA.Conversion; MS-VBAL 6.1.2.x).
Private Const N As Long = 6

Private Function T(ByVal k As Long) As String
    Dim v As Variant, x As Variant, c As Currency, dt As Date, ds As Double, sg As Single
    Dim l As Long, i As Integer, by As Byte, b As Boolean, s As String, o As Object
    On Error GoTo fail
    Select Case k
        Case 1: v = MacID("TEXT")   ' Windows: error expected
        Case 2: v = MacID("")   ' empty
        Case 3: v = MacID("abcd")
        Case 4: v = MacID(Null)   ' 94
        Case 5: v = MacID(5)   ' numeric to string
        Case 6: v = MacID(Constant:="TEXT")   ' named argument
    End Select
    T = "MacID." & k & "=" & Sh(v)
    Exit Function
fail:
    T = "MacID." & k & "=ERR" & Err.Number
End Function

Public Function Run() As String
    Dim k As Long, s As String
    For k = 1 To N
        s = s & T(k) & vbLf
    Next
    Run = s
End Function
