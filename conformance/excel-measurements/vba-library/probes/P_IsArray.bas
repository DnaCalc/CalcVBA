Attribute VB_Name = "P_IsArray"
Option Explicit
' IsArray (Information).
Private Const N As Long = 18

Private Function T(ByVal k As Long) As String
    Dim v As Variant
    Dim va As Variant, fa(1 To 3) As Long, da() As Long, ua() As String
    Dim m(1 To 2, 1 To 2) As Double, c As Collection, s As String
    On Error GoTo fail
    Select Case k
        Case 1: v = IsArray(fa)                       ' fixed-size typed array
        Case 2: ReDim da(2): v = IsArray(da)          ' allocated dynamic array
        Case 3: v = IsArray(ua)                       ' unallocated dynamic array is still an array
        Case 4: v = IsArray(Array(1, 2))              ' Variant array from Array
        Case 5: v = IsArray(Array())                  ' empty Array()
        Case 6: v = IsArray(Split("a b", " "))        ' String array
        Case 7: va = 5: v = IsArray(va)               ' Variant holding a scalar
        Case 8: s = "abc": v = IsArray(s)             ' String
        Case 9: v = IsArray(Null)                     ' Null
        Case 10: v = IsArray(Empty)                    ' Empty
        Case 11: va = fa: v = IsArray(va)              ' Variant holding an array
        Case 12: v = IsArray(m)                        ' two-dimensional
        Case 13: Set c = New Collection: v = IsArray(c)   ' object
        Case 14: v = IsArray(Array(Array(1))(0))       ' array element that is an array
        Case 15: v = IsArray(Array(1, 2)(0))           ' array element that is a scalar
        Case 16: v = IsArray("(1,2)")                  ' string that looks like an array
        Case 17: v = IsArray(CVErr(5))                 ' error value
        Case 18: ReDim da(2): Erase da: v = IsArray(da)   ' after Erase of a dynamic array
    End Select
    T = "IsArray." & k & "=" & Sh(v)
    Exit Function
fail:
    T = "IsArray." & k & "=ERR" & Err.Number
End Function

Public Function Run() As String
    Dim k As Long, s As String
    For k = 1 To N
        s = s & T(k) & vbLf
    Next
    Run = s
End Function
