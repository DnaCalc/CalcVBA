Attribute VB_Name = "P_Collection"
Option Explicit
' Collection class: Add (Key/Before/After), Item, Remove, Count, For Each, default member, bang.
Private Const N As Long = 52

Private Function L(ByVal c As Collection) As String
    Dim x As Variant, s As String
    For Each x In c
        s = s & Sh(x) & ";"
    Next
    L = s
End Function

Private Function DupThenCount() As String
    Dim c As New Collection
    On Error Resume Next
    c.Add 1, "k"
    c.Add 2, "k"
    DupThenCount = "err=" & Err.Number & " count=" & c.Count & " items=" & L(c)
End Function

Private Function Visit(ByVal mode As Long) As String
    Dim c As New Collection, x As Variant, n As Long, s As String
    c.Add 1: c.Add 2: c.Add 3
    For Each x In c
        n = n + 1
        s = s & Sh(x) & ";"
        If mode = 1 And n = 1 Then c.Add 99
        If mode = 2 And n = 1 Then c.Remove 1
        If n > 10 Then Exit For
    Next
    Visit = "n=" & n & " " & s & "count=" & c.Count
End Function

Private Function T(ByVal k As Long) As String
    Dim v As Variant, c As Collection, o1 As Object, o2 As Object, x As Variant, n As Long, d As Double, i As Long
    On Error GoTo fail
    Set c = New Collection
    Select Case k
        Case 1: c.Add "a": c.Add "b": c.Add "c": v = L(c)               ' items in order
        Case 2: c.Add 10, "k": v = c.Item("k")                           ' by key
        Case 3: c.Add 1, "k": c.Add 2, "k": v = c.Count                  ' duplicate key: 457
        Case 4: v = DupThenCount()                                       ' state after the 457
        Case 5: c.Add 1, "a": c.Add 2, "A": v = c.Count                  ' keys differ in case: 457
        Case 6: c.Add 1, "a": c.Add 2, "a ": v = c.Count                 ' trailing space is a different key
        Case 7: c.Add 1, ChrW(233): c.Add 2, "e": v = c.Count            ' e-acute vs e: different keys
        Case 8: c.Add 1, ChrW(233): c.Add 2, ChrW(201): v = c.Count      ' e-acute vs E-acute: case only
        Case 9: c.Add 1, "": v = c.Count                                 ' empty-string key
        Case 10: c.Add 1, "": v = c.Item("")                             ' look up the empty key
        Case 11: c.Add 1, 5: v = c.Count                                 ' numeric key: 13
        Case 12: c.Add 1, CStr(5): v = c.Item("5")                       ' string digits as key
        Case 13: c.Add 1, "1": c.Add 2, "2": v = c(1) & "/" & c("1")     ' index 1 vs key "1"
        Case 14: c.Add 1, Null: v = c.Count                              ' Null key
        Case 15: c.Add 1, Empty: v = c.Count                             ' Empty key
        Case 16: c.Add 1, True: v = c.Count                              ' Boolean key
        Case 17: c.Add "a": c.Add "b": c.Add "c": c.Add "X", , 2: v = L(c)          ' before index
        Case 18: c.Add "a": c.Add "b": c.Add "c": c.Add "X", , , 2: v = L(c)        ' after index
        Case 19: c.Add "a", "ka": c.Add "b", "kb": c.Add "X", , "kb": v = L(c)      ' before key
        Case 20: c.Add "a", "ka": c.Add "b", "kb": c.Add "X", , , "ka": v = L(c)    ' after key
        Case 21: c.Add "a": c.Add "X", , 1, 1: v = L(c)                  ' Before and After both: 5
        Case 22: c.Add "a": c.Add "b": c.Add "X", , 3: v = L(c)          ' before = Count+1
        Case 23: c.Add "a": c.Add "b": c.Add "X", , , 2: v = L(c)        ' after = Count: append
        Case 24: c.Add "a": c.Add "b": c.Add "X", , , 3: v = L(c)        ' after = Count+1
        Case 25: c.Add "a": c.Add "X", , , 0: v = L(c)                   ' after = 0
        Case 26: c.Add "X", , 1: v = L(c)                                ' before 1 on empty collection
        Case 27: c.Add "a": c.Add "X", , "zz": v = L(c)                  ' before missing key
        Case 28: c.Add "a": c.Add "X", , "1": v = L(c)                   ' before string digits: a key
        Case 29: c.Add "a": c.Add "b": c.Add "X", , 1.5: v = L(c)        ' before Double rounds
        Case 30: c.Add "a": c.Add "b": c.Add "c": v = c.Item(2.5)        ' Double index rounds half to even
        Case 31: c.Add "a": c.Add "b": v = c.Item(1.5)                   ' 1.5 rounds to 2
        Case 32: c.Add "a": c.Add "b": v = c.Item(0.5)                   ' rounds to 0
        Case 33: c.Add "a": c.Add "b": v = c.Item("2")                   ' string digits are a key lookup
        Case 34: c.Add "a": c.Add "b": v = c.Item(0)                     ' index 0
        Case 35: c.Add "a": c.Add "b": v = c.Item(3)                     ' index past Count
        Case 36: c.Add "a": v = c.Item("zz")                             ' missing key
        Case 37: c.Add "a": v = c.Item(Null)                             ' Null index
        Case 38: c.Add "a": v = c.Item(Empty)                            ' Empty index
        Case 39: c.Add "a": c.Add "b": c.Add "c": c.Remove 2: v = L(c)   ' remove by index
        Case 40: c.Add "a", "ka": c.Add "b", "kb": c.Remove "ka": v = L(c) & c.Count    ' remove by key
        Case 41: c.Add "a": c.Remove 2: v = c.Count                      ' remove past Count
        Case 42: c.Add "a", "k": c.Remove "zz": v = c.Count              ' remove missing key
        Case 43: c.Add "a": c.Remove 1: c.Remove 1: v = c.Count         ' remove from empty collection
        Case 44: v = Visit(0) & "|" & Visit(1) & "|" & Visit(2)          ' For Each; add and remove during enumeration
        Case 45: c.Add 1, "k": v = c(1) & "," & c("k") & "," & c!k       ' default member and bang
        Case 46: c.Add 7, "k": v = c.Item(Index:="k"): c.Add Item:=8, Key:="m": c.Remove Index:=1: v = v & "/" & c.Count & c.Item(Index:="m")    ' named arguments
        Case 47: Set o1 = New Collection: c.Add o1: Set o2 = c(1): v = (o1 Is o2) & TypeName(o2)    ' object item identity
        Case 48: c.Add Nothing, "n": Set v = c("n"): v = TypeName(v)    ' Nothing as an item
        Case 49: c.Add Null: c.Add Empty: c.Add Array(1, 2): v = L(c)   ' Variant items
        Case 50: v = TypeName(c) & "," & TypeName(c.Count) & "," & c.Count    ' type of Count
        Case 51: For i = 1 To 5: c.Add i * 10: Next: n = 0: For Each x In c: n = n + x: Next: v = n & "/" & TypeName(x)   ' sum and loop-variable type
        Case 52: v = Sh(c.Item(1))                                       ' Item on an empty collection
    End Select
    T = "Collection." & k & "=" & Sh(v)
    Exit Function
fail:
    T = "Collection." & k & "=ERR" & Err.Number
End Function

Public Function Run() As String
    Dim k As Long, s As String
    For k = 1 To N
        s = s & T(k) & vbLf
    Next
    Run = s
End Function
