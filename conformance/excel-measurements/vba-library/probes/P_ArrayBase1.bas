Attribute VB_Name = "P_ArrayBase1"
Option Explicit
Option Base 1
' Array() and array declarations under Option Base 1.
Private Const N As Long = 16

Private Function PA(ParamArray p() As Variant) As String
    PA = LBound(p) & "," & UBound(p)
End Function

Private Function T(ByVal k As Long) As String
    Dim v As Variant, a() As Variant, d(2, 3) As Long, f(3) As Long, g(2 To 4) As Long, h(0 To 2) As Long, i As Long, s As String
    On Error GoTo fail
    Select Case k
        Case 1: v = Array(1, 2, 3)                                         ' bounds start at 1
        Case 2: v = Array(): v = LBound(v) & "," & UBound(v)               ' empty array bounds
        Case 3: v = Array("a"): v = LBound(v) & "," & UBound(v)
        Case 4: v = Array(Array(1, 2), Array(3)): v = LBound(v) & LBound(v(1)) & "," & UBound(v(1))    ' nested bounds
        Case 5: v = LBound(f) & "," & UBound(f)                            ' Dim f(3)
        Case 6: v = LBound(g) & "," & UBound(g)                            ' explicit To is unaffected
        Case 7: v = LBound(h) & "," & UBound(h)                            ' 0 To 2
        Case 8: v = LBound(d, 1) & "," & UBound(d, 1) & "," & LBound(d, 2) & "," & UBound(d, 2)    ' two dimensions
        Case 9: ReDim a(3): v = LBound(a) & "," & UBound(a)                ' ReDim
        Case 10: a = Array(1, 2): v = LBound(a) & "," & UBound(a)          ' assigned to a dynamic array
        Case 11: v = Split("a b c", " "): v = LBound(v) & "," & UBound(v)  ' Split result base
        Case 12: v = PA(1, 2, 3)                                           ' ParamArray base
        Case 13: v = VBA.Array(7, 8): v = LBound(v) & "," & UBound(v) & "," & v(1)    ' qualified
        Case 14: v = Array(10, 20, 30): v = v(1) & "," & v(3)              ' element access
        Case 15: v = Array(10, 20, 30)(0)                                  ' index 0 is out of range
        Case 16: v = Filter(Array("a", "b", "ab"), "a"): v = LBound(v) & "," & UBound(v)    ' Filter result base
    End Select
    T = "ArrayBase1." & k & "=" & Sh(v)
    Exit Function
fail:
    T = "ArrayBase1." & k & "=ERR" & Err.Number
End Function

Public Function Run() As String
    Dim k As Long, s As String
    For k = 1 To N
        s = s & T(k) & vbLf
    Next
    Run = s
End Function
