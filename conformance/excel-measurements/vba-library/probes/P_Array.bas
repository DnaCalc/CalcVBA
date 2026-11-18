Attribute VB_Name = "P_Array"
Option Explicit
' Array() (default Option Base 0): bounds, element types, nesting, result type.
Private Const N As Long = 26

Private Function T(ByVal k As Long) As String
    Dim v As Variant, a() As Variant, x As Variant, n As Long, i As Long, src As Variant, s As String
    On Error GoTo fail
    Select Case k
        Case 1: v = Array(1, 2, 3)                                         ' typical
        Case 2: v = Array()                                                ' empty: bounds 0,-1
        Case 3: v = Array(): v = LBound(v) & "," & UBound(v)               ' bounds of the empty array
        Case 4: v = Array("a")                                             ' one element
        Case 5: v = Array(1, "a", 2.5, Null, Empty, True, 1&, 1.5!, 1@, #1/2/2000#)    ' element types
        Case 6: v = Array(Array(1, 2), 3)                                  ' nested
        Case 7: v = Array(Array(1, 2), 3): v = UBound(v(0)) & "," & v(0)(1) & "," & v(1)    ' nested access
        Case 8: v = VarType(Array(1, 2))                                   ' 8204 = vbArray + vbVariant
        Case 9: v = TypeName(Array(1, 2))                                  ' Variant()
        Case 10: v = TypeName(Array())                                     ' Variant() when empty
        Case 11: v = VarType(Array())
        Case 12: v = Array(1, 2)(1)                                        ' index the result directly
        Case 13: v = Array(Array())                                        ' nested empty
        Case 14: v = Array(Nothing, New Collection)                        ' objects as elements
        Case 15: v = Array(1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20): v = UBound(v) & "," & v(19)    ' 20 elements
        Case 16: src = 1: v = Array(src): src = 5: v = v(0)                 ' elements are copies
        Case 17: src = Array(1, 2): v = Array(src): v = LBound(v) & "," & UBound(v) & "," & TypeName(v(0))    ' an array variable becomes one element
        Case 18: v = Array(CVErr(5), CVErr(2042))                          ' error values
        Case 19: v = Array(DateSerial(2000, 1, 2), CCur(1.5), CDec(2), CByte(3), CSng(0.1))    ' more element types
        Case 20: v = IsArray(Array())                                      ' IsArray on the empty array
        Case 21: a = Array(1, 2, 3): v = LBound(a) & "," & UBound(a) & "," & TypeName(a)    ' assigned to a Variant dynamic array
        Case 22: a = Array(1, 2): ReDim Preserve a(0 To 3): a(3) = 9: v = a    ' grow after Array
        Case 23: n = 0: For Each x In Array(1, 2, 3): n = n + x: Next: v = n    ' For Each over Array()
        Case 24: v = VBA.Array(7, 8)                                       ' qualified name
        Case 25: v = Array(1, 2, 3)(5)                                     ' out of range index: 9
        Case 26: v = Array(-2147483648#, 2147483647, 32767, -32768, 255)   ' literal types at the edges
    End Select
    T = "Array." & k & "=" & Sh(v)
    Exit Function
fail:
    T = "Array." & k & "=ERR" & Err.Number
End Function

Public Function Run() As String
    Dim k As Long, s As String
    For k = 1 To N
        s = s & T(k) & vbLf
    Next
    Run = s
End Function
