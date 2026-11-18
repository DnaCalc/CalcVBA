Attribute VB_Name = "P_FloatingConversionEdges"
Option Explicit
' Excel-backed floating conversion overflow, embedded spaces and Single rounding.
Private Const N As Long = 35
Private Function T(ByVal k As Long) As String
    Dim v As Variant
    On Error GoTo fail
    Select Case k
        Case 1: v = CDbl("1 000")
        Case 2: v = CDbl("1 0 0 0")
        Case 3: v = CDbl(" 1 . 5 ")
        Case 4: v = CDbl("1" & vbTab & "000")
        Case 5: v = CDbl("1" & ChrW(160) & "000")
        Case 6: v = CDbl("- 1 000")
        Case 7: v = CDbl("1 E 3")
        Case 8: v = CDbl("1 2.3 4")
        Case 9: v = CDbl("1e3 0")
        Case 10: v = CDbl("1e309")
        Case 11: v = CDbl("-1e309")
        Case 12: v = CDbl("1e-400")
        Case 13: v = CDbl("1e309x")
        Case 14: v = CSng("1 000")
        Case 15: v = CSng("1 0 0 0")
        Case 16: v = CSng(" 1 . 5 ")
        Case 17: v = CSng("1" & vbTab & "000")
        Case 18: v = CSng("1" & ChrW(160) & "000")
        Case 19: v = CSng("- 1 000")
        Case 20: v = CSng("1 E 3")
        Case 21: v = CSng("1 2.3 4")
        Case 22: v = CSng("1e3 0")
        Case 23: v = CSng("1e309")
        Case 24: v = CSng("-1e309")
        Case 25: v = CSng("1e-400")
        Case 26: v = CSng("1e309x")
        Case 27: v = CSng(3.40282346E38)
        Case 28: v = CSng(3.40282347E38)
        Case 29: v = CSng(3.40282356779733E38)
        Case 30: v = CSng(3.40282356779734E38)
        Case 31: v = CSng(-3.40282347E38)
        Case 32: v = CSng(-3.40282356779733E38)
        Case 33: v = CSng(-3.40282356779734E38)
        Case 34: v = CSng("3.40282347E38")
        Case 35: v = CSng("3.5E38")
    End Select
    T = "FloatingConversionEdges." & k & "=" & Sh(v)
    Exit Function
fail:
    T = "FloatingConversionEdges." & k & "=ERR" & Err.Number
End Function
Public Function Run() As String
    Dim k As Long, s As String
    For k = 1 To N
        s = s & T(k) & vbLf
    Next
    Run = s
End Function
