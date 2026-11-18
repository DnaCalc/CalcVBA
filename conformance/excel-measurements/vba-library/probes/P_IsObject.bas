Attribute VB_Name = "P_IsObject"
Option Explicit
' IsObject (Information).
Private Const N As Long = 14

Private Function T(ByVal k As Long) As String
    Dim v As Variant
    Dim va As Variant, o As Object, c As Collection, d As Object
    On Error GoTo fail
    Select Case k
        Case 1: Set c = New Collection: v = IsObject(c)   ' Collection
        Case 2: v = IsObject(o)                           ' unset Object variable (Nothing)
        Case 3: Set va = New Collection: v = IsObject(va) ' Variant holding an object
        Case 4: Set va = Nothing: v = IsObject(va)        ' Variant holding Nothing
        Case 5: va = Empty: v = IsObject(va)              ' Empty
        Case 6: v = IsObject(Null)                        ' Null
        Case 7: v = IsObject(5)                           ' number
        Case 8: v = IsObject("abc")                       ' String
        Case 9: Set d = CreateObject("Scripting.Dictionary"): v = IsObject(d)   ' Dictionary
        Case 10: v = IsObject(Err)                         ' ErrObject
        Case 11: Set c = New Collection: v = IsObject(Array(c))  ' array of objects is not an object
        Case 12: Set c = New Collection: v = IsObject(c.Count)   ' property result
        Case 13: v = IsObject(CVErr(5))                    ' error value
        Case 14: Set c = New Collection: c.Add 1: v = IsObject(c.Item(1))  ' element value
    End Select
    T = "IsObject." & k & "=" & Sh(v)
    Exit Function
fail:
    T = "IsObject." & k & "=ERR" & Err.Number
End Function

Public Function Run() As String
    Dim k As Long, s As String
    For k = 1 To N
        s = s & T(k) & vbLf
    Next
    Run = s
End Function
