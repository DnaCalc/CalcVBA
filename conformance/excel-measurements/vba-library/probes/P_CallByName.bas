Attribute VB_Name = "P_CallByName"
Option Explicit
' CallByName (Interaction; Collection and Scripting.Dictionary members by name).
Private Const N As Long = 33

Private Function T(ByVal k As Long) As String
    Dim v As Variant
    Dim c As Collection, d As Object, o As Object
    On Error GoTo fail
    Select Case k
        Case 1: Set c = New Collection: v = CallByName(c, "Count", VbGet)                 ' property get on an empty Collection
        Case 2: Set c = New Collection: CallByName c, "Add", VbMethod, "x": v = CallByName(c, "Count", VbGet)   ' method call, then Count
        Case 3: Set c = New Collection: c.Add "x": v = CallByName(c, "Item", VbMethod, 1)  ' Item as a method
        Case 4: Set c = New Collection: c.Add "x": v = CallByName(c, "Item", VbGet, 1)     ' Item as a property get
        Case 5: Set c = New Collection: CallByName c, "Add", VbMethod, "x", "k": v = CallByName(c, "Item", VbGet, "k")   ' Add with a key, Item by key
        Case 6: Set c = New Collection: c.Add "x": CallByName c, "Remove", VbMethod, 1: v = c.Count   ' Remove
        Case 7: Set c = New Collection: c.Add "x": v = TypeName(CallByName(c, "Count", VbGet))   ' result type Long
        Case 8: Set c = New Collection: v = CallByName(c, "NoSuchMember", VbGet)          ' missing member: 438
        Case 9: Set c = New Collection: v = CallByName(c, "NoSuchMember", VbMethod)       ' missing method: 438
        Case 10: Set c = New Collection: v = CallByName(c, "NoSuchMember", VbLet, 1)       ' missing property let
        Case 11: Set c = New Collection: v = CallByName(c, "count", VbGet)                 ' name is case-insensitive
        Case 12: Set c = New Collection: v = CallByName(c, "", VbGet)                      ' empty member name
        Case 13: Set c = New Collection: v = CallByName(c, "Count", VbLet, 5)              ' read-only property assignment
        Case 14: Set c = New Collection: v = CallByName(c, "Count", VbMethod)              ' property used as a method
        Case 15: Set c = New Collection: v = CallByName(c, "Add", VbGet, "x")              ' method used as a property get
        Case 16: Set c = New Collection: v = CallByName(c, "Count", VbSet, c)              ' VbSet on a non-object property
        Case 17: Set c = New Collection: v = CallByName(c, "Count", 0)                     ' invalid call type 0
        Case 18: Set c = New Collection: v = CallByName(c, "Count", 99)                    ' invalid call type 99
        Case 19: Set c = New Collection: v = CallByName(c, "Add", VbMethod)                ' missing required argument
        Case 20: Set c = New Collection: v = CallByName(c, "Item", VbGet, 5)               ' unsupported property get: error 438
        Case 21: v = "compile-refused"
        Case 22: v = CallByName(o, "Count", VbGet)                                         ' Nothing object: error 5
        Case 23: Set d = CreateObject("Scripting.Dictionary"): CallByName d, "Add", VbMethod, "k", 1: v = CallByName(d, "Count", VbGet)   ' Dictionary Add and Count
        Case 24: Set d = CreateObject("Scripting.Dictionary"): d.Add "k", 1: v = CallByName(d, "Item", VbGet, "k")      ' Dictionary Item get
        Case 25: Set d = CreateObject("Scripting.Dictionary"): d.Add "k", 1: v = CallByName(d, "Exists", VbMethod, "k")  ' Exists true
        Case 26: Set d = CreateObject("Scripting.Dictionary"): v = CallByName(d, "Exists", VbMethod, "k")                ' Exists false
        Case 27: Set d = CreateObject("Scripting.Dictionary"): d.Add "k", 1: CallByName d, "Item", VbLet, "k", 5: v = d("k")   ' Item property let
        Case 28: Set d = CreateObject("Scripting.Dictionary"): CallByName d, "CompareMode", VbLet, 1: v = CallByName(d, "CompareMode", VbGet)   ' CompareMode let and get
        Case 29: v = CallByName(Err, "Number", VbGet)                                      ' ErrObject property
        Case 30: Err.Clear: v = CallByName(Err, "Description", VbGet)                      ' empty description
        Case 31: v = CallByName(Err, "Raise", VbMethod, 5)                                 ' Err.Raise through CallByName: error 5
        Case 32: v = VbMethod & "," & VbGet & "," & VbLet & "," & VbSet                    ' constants
        Case 33: v = TypeName(VbGet)                                                       ' enum constant type
    End Select
    T = "CallByName." & k & "=" & Sh(v)
    Exit Function
fail:
    T = "CallByName." & k & "=ERR" & Err.Number
End Function

Public Function Run() As String
    Dim k As Long, s As String
    For k = 1 To N
        s = s & T(k) & vbLf
    Next
    Run = s
End Function
