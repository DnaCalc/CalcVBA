Attribute VB_Name = "P_UBound"
Option Explicit
' UBound (language function; fixed, dynamic, unallocated and multi-dimensional arrays, errors 9 and 13).
Private Const N As Long = 46

Private Function UBH(a() As Long, ByVal d As Long) As Long
    UBH = UBound(a, d)
End Function
Private Function UBV(a As Variant) As Long
    UBV = UBound(a)
End Function

Private Function T(ByVal k As Long) As String
    Dim v As Variant
    Dim fa(1 To 5) As Long, g(5) As Long, m(1 To 3, 0 To 4, 2 To 6) As Long, da() As Long, va As Variant, vn(-5 To 5) As Double
    Dim nn As Long, bt() As Byte
    On Error GoTo fail
    Select Case k
        Case 1: v = UBound(fa)                        ' fixed 1 To 5
        Case 2: v = UBound(g)                         ' Dim g(5): 0 To 5
        Case 3: v = UBound(m)                         ' default dimension 1
        Case 4: v = UBound(m, 1)
        Case 5: v = UBound(m, 2)
        Case 6: v = UBound(m, 3)
        Case 7: nn = 4: v = UBound(m, nn)                      ' dimension past the last: error 9
        Case 8: nn = 0: v = UBound(m, nn)                      ' dimension 0: error 9
        Case 9: nn = -1: v = UBound(m, nn)                     ' negative dimension: error 9
        Case 10: v = UBound(m, Null)                   ' Null dimension: error 94
        Case 11: v = UBound(m, "2")                    ' numeric string dimension
        Case 12: v = UBound(m, 1.5)                    ' fractional dimension rounds to 2
        Case 13: v = UBound(m, 2.5)                    ' banker's rounding to 2 (3 would give 6)
        Case 14: v = UBound(m, True)                   ' -1: error 9
        Case 15: v = UBound(m, Empty)                  ' 0: error 9
        Case 16: v = UBound(da)                        ' unallocated dynamic array: error 9
        Case 17: ReDim da(3 To 7): v = UBound(da)      ' allocated
        Case 18: ReDim da(2): Erase da: v = UBound(da) ' after Erase: error 9
        Case 19: ReDim da(1 To 3): ReDim Preserve da(1 To 6): v = UBound(da)   ' ReDim Preserve
        Case 20: ReDim da(0 To 0): v = UBound(da)      ' single element
        Case 21: nn = -1: ReDim da(0 To nn): v = UBound(da)   ' upper below lower: error 9
        Case 22: nn = 0: ReDim da(1 To nn): v = UBound(da)    ' 1 To 0: error 9
        Case 23: v = UBound(vn)                        ' negative lower bound
        Case 24: v = UBound(Array(1, 2, 3))            ' Array()
        Case 25: v = UBound(Array())                   ' empty Array(): -1
        Case 26: v = UBound(Split("a b c", " "))       ' Split
        Case 27: v = UBound(Split(""))                 ' empty Split: -1
        Case 28: va = 5: v = UBound(va)                ' non-array number: error 13
        Case 29: va = "abc": v = UBound(va)            ' string: error 13
        Case 30: va = Null: v = UBound(va)             ' Null: error 13
        Case 31: v = UBound(va)                        ' Empty Variant: error 13
        Case 32: Set va = New Collection: v = UBound(va)    ' object: error 13
        Case 33: va = Array(1, 2): v = UBound(va)      ' Variant holding an array
        Case 34: ReDim va(1 To 2, 1 To 3): v = UBound(va, 2)   ' two-dimensional Variant array
        Case 35: ReDim va(1 To 2, 1 To 3): v = UBound(va)      ' default dimension of a Variant array
        Case 36: v = UBound(Array(Array(1, 2), 3)(0))  ' nested array
        Case 37: v = TypeName(UBound(fa))              ' result type Long
        Case 38: v = UBound(fa, 1)                     ' explicit dimension 1
        Case 39: va = fa: v = UBound(va)               ' Variant copy of a fixed array
        Case 40: v = UBH(fa, 1)                        ' array passed to a procedure keeps its bounds
        Case 41: v = UBH(fa, 2)                        ' no second dimension: error 9
        Case 42: v = UBV(Array(1, 2, 3, 4))            ' Variant array passed on
        Case 43: v = UBound(StrConv("abc", vbFromUnicode))   ' Byte array
        Case 44: bt = StrConv("abcd", vbFromUnicode): v = UBound(bt)   ' Byte array in a variable
        Case 45: ReDim da(1 To 3): v = UBound(da, 1)   ' dimension on a 1-D dynamic array
        Case 46: v = "compile-refused"
    End Select
    T = "UBound." & k & "=" & Sh(v)
    Exit Function
fail:
    T = "UBound." & k & "=ERR" & Err.Number
End Function

Public Function Run() As String
    Dim k As Long, s As String
    For k = 1 To N
        s = s & T(k) & vbLf
    Next
    Run = s
End Function
