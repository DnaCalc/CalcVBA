Attribute VB_Name = "P_LibraryNamesEdges"
Option Explicit
' Calendar, MacID and Load/Unload cases captured in Excel 16.0.20527.
Private Const N As Long = 23
Private Function T(ByVal k As Long) As String
    Dim v As Variant, x As Variant, o As Object
    On Error GoTo fail
    Calendar = 0
    Select Case k
        Case 1: Calendar = 1: v = CDbl(DateSerial(1445, 1, 1))
        Case 2: Calendar = 1: v = Year(#7/19/2023#)
        Case 3: Calendar = 1: v = Month(#7/19/2023#)
        Case 4: Calendar = 1: v = Day(#7/19/2023#)
        Case 5: Calendar = 1: v = CDbl(DateAdd("m", 1, #7/19/2023#))
        Case 6: Calendar = 1: v = DateDiff("m", #7/19/2023#, #8/18/2023#)
        Case 7: Calendar = 1: v = DatePart("y", #7/19/2023#)
        Case 8: Calendar = 1: v = Format$(#7/19/2023#, "yyyy-mm-dd")
        Case 9: Calendar = 1: v = CDbl(DateValue("1445/01/01"))
        Case 10: v = MacID("abcde")
        Case 11: v = MacID("a")
        Case 12: v = MacID("ab")
        Case 13: v = MacID("ab" & Chr$(0) & "d")
        Case 14: v = MacID(ChrW$(256) & "abc") ' locale ANSI substitution
        Case 15: Load o: v = "loaded Nothing"
        Case 16: Unload o: v = "unloaded Nothing"
        Case 17: Set o = New Collection: Load o: v = "loaded Collection"
        Case 18: Set o = New Collection: Unload o: v = "unloaded Collection"
        Case 19: x = 5: Load x: v = "loaded number"
        Case 20: x = 5: Unload x: v = "unloaded number"
        Case 21: x = Null: Load x: v = "loaded Null"
        Case 22: x = Empty: Unload x: v = "unloaded Empty"
        Case 23: v = (UserForms Is UserForms)
    End Select
    Calendar = 0
    T = "LibraryNamesEdges." & k & "=" & Sh(v)
    Exit Function
fail:
    Dim en As Long
    en = Err.Number
    On Error Resume Next
    Calendar = 0
    T = "LibraryNamesEdges." & k & "=ERR" & en
End Function
Public Function Run() As String
    Dim k As Long, s As String
    For k = 1 To N
        s = s & T(k) & vbLf
    Next
    Run = s
End Function
