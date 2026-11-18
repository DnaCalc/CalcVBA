Attribute VB_Name = "P_LBound"
Option Explicit
' LBound (language function; fixed, dynamic, unallocated and multi-dimensional arrays, errors 9 and 13).
Private Const N As Long = 43

Private Function LBH(a() As Long, ByVal d As Long) As Long
    LBH = LBound(a, d)
End Function
Private Function LBV(a As Variant) As Long
    LBV = LBound(a)
End Function

Private Function T(ByVal k As Long) As String
    Dim v As Variant
    Dim fa(1 To 5) As Long, g(5) As Long, m(1 To 3, 0 To 4, 2 To 6) As Long, da() As Long, va As Variant, vn(-5 To 5) As Double
    Dim nn As Long, bt() As Byte
    On Error GoTo fail
    Select Case k
        Case 1: v = LBound(fa)                        ' fixed 1 To 5
        Case 2: v = LBound(g)                         ' Dim g(5): 0
        Case 3: v = LBound(m)                         ' default dimension 1
        Case 4: v = LBound(m, 1)
        Case 5: v = LBound(m, 2)
        Case 6: v = LBound(m, 3)
        Case 7: nn = 4: v = LBound(m, nn)                      ' dimension past the last: error 9
        Case 8: nn = 0: v = LBound(m, nn)                      ' dimension 0: error 9
        Case 9: nn = -1: v = LBound(m, nn)                     ' negative dimension: error 9
        Case 10: v = LBound(m, Null)                   ' Null dimension: error 94
        Case 11: v = LBound(m, "3")                    ' numeric string dimension
        Case 12: v = LBound(m, 1.5)                    ' fractional dimension rounds to 2
        Case 13: v = LBound(m, 2.5)                    ' banker's rounding to 2
        Case 14: v = LBound(m, True)                   ' -1: error 9
        Case 15: v = LBound(m, Empty)                  ' 0: error 9
        Case 16: v = LBound(da)                        ' unallocated dynamic array: error 9
        Case 17: ReDim da(3 To 7): v = LBound(da)      ' allocated
        Case 18: ReDim da(2): Erase da: v = LBound(da) ' after Erase: error 9
        Case 19: ReDim da(1 To 3): ReDim Preserve da(1 To 6): v = LBound(da)   ' ReDim Preserve
        Case 20: ReDim da(0 To 0): v = LBound(da)      ' single element
        Case 21: nn = -1: ReDim da(0 To nn): v = LBound(da)   ' upper below lower: error 9
        Case 22: v = LBound(vn)                        ' negative lower bound
        Case 23: v = LBound(Array(1, 2, 3))            ' Array(): 0
        Case 24: v = LBound(Array())                   ' empty Array(): 0
        Case 25: v = LBound(Split("a b c", " "))       ' Split
        Case 26: v = LBound(Split(""))                 ' empty Split
        Case 27: va = 5: v = LBound(va)                ' non-array number: error 13
        Case 28: va = "abc": v = LBound(va)            ' string: error 13
        Case 29: va = Null: v = LBound(va)             ' Null: error 13
        Case 30: v = LBound(va)                        ' Empty Variant: error 13
        Case 31: Set va = New Collection: v = LBound(va)    ' object: error 13
        Case 32: ReDim va(2 To 4, 5 To 9): v = LBound(va, 1) & "," & LBound(va, 2)   ' two-dimensional Variant array
        Case 33: ReDim va(2 To 4, 5 To 9): v = LBound(va)      ' default dimension of a Variant array
        Case 34: v = LBound(Array(Array(1, 2), 3)(0))  ' nested array
        Case 35: v = TypeName(LBound(fa))              ' result type Long
        Case 36: v = LBound(fa, 1)                     ' explicit dimension 1
        Case 37: va = fa: v = LBound(va)               ' Variant copy of a fixed array
        Case 38: v = LBH(fa, 1)                        ' array passed to a procedure keeps its bounds
        Case 39: v = LBH(fa, 2)                        ' no second dimension: error 9
        Case 40: v = LBV(Array(1, 2, 3, 4))            ' Variant array passed on
        Case 41: v = LBound(StrConv("abc", vbFromUnicode))   ' Byte array
        Case 42: bt = StrConv("abcd", vbFromUnicode): v = LBound(bt)   ' Byte array in a variable
        Case 43: v = "compile-refused"
    End Select
    T = "LBound." & k & "=" & Sh(v)
    Exit Function
fail:
    T = "LBound." & k & "=ERR" & Err.Number
End Function

Public Function Run() As String
    Dim k As Long, s As String
    For k = 1 To N
        s = s & T(k) & vbLf
    Next
    Run = s
End Function
