Attribute VB_Name = "P_Error"
Option Explicit
' Error / Error$ (VBA.Conversion; MS-VBAL 6.1.2.x).
Private Const N As Long = 40

Private Function T(ByVal k As Long) As String
    Dim v As Variant, x As Variant, c As Currency, dt As Date, ds As Double, sg As Single
    Dim l As Long, i As Integer, by As Byte, b As Boolean, s As String, o As Object
    On Error GoTo fail
    Select Case k
        Case 1: v = Error(5)   ' Invalid procedure call
        Case 2: v = Error(13)   ' Type mismatch
        Case 3: v = Error(6)   ' Overflow
        Case 4: v = Error(9)
        Case 5: v = Error(11)
        Case 6: v = Error(7)
        Case 7: v = Error(91)
        Case 8: v = Error(94)
        Case 9: v = Error(0)   ' empty string
        Case 10: v = Error(1)   ' application-defined
        Case 11: v = Error(999)   ' unknown number
        Case 12: v = Error(65535)
        Case 13: v = Error(65536)   ' above 65535
        Case 14: v = Error(-1)   ' negative
        Case 15: v = Error(2147483647)
        Case 16: v = Error(2007)   ' Excel error code, not a VBA error number
        Case 17: v = Error(5.5)   ' rounds
        Case 18: v = Error("5")   ' numeric string
        Case 19: v = Error("abc")   ' 13
        Case 20: v = Error(Null)   ' 94
        Case 21: v = Error(Empty)
        Case 22: v = Error(True)
        Case 23: v = Error(CByte(13))
        Case 24: v = Error(ErrorNumber:=13)   ' named argument
        Case 25: v = Error$(ErrorNumber:=5)
        Case 26: v = Error$(5)   ' String form
        Case 27: v = Error$(0)
        Case 28: v = Error$(999)
        Case 29: v = Error$(-5)
        Case 30: v = Error$(Null)   ' 94
        Case 31: Err.Clear: v = Error   ' no argument after Clear
        Case 32: Err.Clear: v = Error$   ' no argument, String form
        Case 33: On Error Resume Next: Err.Raise 11: v = Error: On Error GoTo fail   ' no argument returns last error text
        Case 34: v = Error(CVErr(5))   ' 13
        Case 35: v = Error(3)   ' Return without GoSub
        Case 36: v = Error(10)
        Case 37: v = Error(20)
        Case 38: v = Error(53)
        Case 39: v = Error(76)
        Case 40: v = Error(429)
    End Select
    T = "Error." & k & "=" & Sh(v)
    Exit Function
fail:
    T = "Error." & k & "=ERR" & Err.Number
End Function

Public Function Run() As String
    Dim k As Long, s As String
    For k = 1 To N
        s = s & T(k) & vbLf
    Next
    Run = s
End Function
