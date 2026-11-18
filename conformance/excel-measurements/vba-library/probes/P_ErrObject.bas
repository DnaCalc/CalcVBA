Attribute VB_Name = "P_ErrObject"
Option Explicit
' Err object: Raise, Number, Description, Source, HelpFile, HelpContext, Clear, LastDllError, Err lifetime.
Private Const N As Long = 54

Private Function Dump() As String
    Dump = "N=" & Err.Number & ";S=" & Esc(Err.Source) & ";D=" & Esc(Err.Description) & ";HF=" & Esc(Err.HelpFile) & ";HC=" & Err.HelpContext
End Function

Private Sub Thrower()
    Err.Raise 5, , "thrown"
End Sub

Private Sub DivZero()
    Dim z As Long, x As Long
    x = 1 \ z
End Sub

Private Function RaiseThenExit() As Long
    On Error Resume Next
    Err.Raise 5, "inner", "inner desc"
    RaiseThenExit = Err.Number
    Exit Function
End Function

Private Sub HandlerThenExitSub()
    On Error GoTo h
    Err.Raise 5, "hs", "hd"
    Exit Sub
h:
    Exit Sub
End Sub

Private Function H1() As String
    Dim s As String
    On Error GoTo h
    Err.Raise 5, "s1", "d1"
    s = s & "after=" & Dump()
    H1 = s
    Exit Function
h:
    s = "inh=" & Dump() & "|"
    Resume Next
End Function

Private Function H2() As String
    Dim n As Long
    n = RaiseThenExit()
    H2 = "ret=" & n & " after=" & Dump()
End Function

Private Function H3(ByVal which As Long) As String
    On Error GoTo h
    If which = 1 Then Thrower Else DivZero
    H3 = "noerr"
    Exit Function
h:
    H3 = "caught " & Dump()
End Function

Private Function H4() As String
    Dim s As String
    On Error GoTo h
    Err.Raise 5, "s1", "d1"
    s = s & "after=" & Dump()
    H4 = s
    Exit Function
h:
    s = "inh=" & Err.Number & ";"
    Err.Clear
    s = s & "cleared=" & Err.Number & ";"
    Resume Next
End Function

Private Function H5() As String
    HandlerThenExitSub
    H5 = "after exit sub in handler: " & Dump()
End Function

Private Function H6() As String
    Dim s As String
    On Error GoTo h
    Err.Raise 5
    H6 = "noerr"
    Exit Function
h:
    s = Dump()
    On Error GoTo h2
    Err.Raise 11
    Exit Function
h2:
    H6 = s & "|second=" & Dump()
End Function

Private Function G(ByVal k As Long) As String
    Dim s As String, v As Variant, x As Double, z As Long
    On Error Resume Next
    Err.Clear
    Select Case k
        Case 1: Err.Raise 5: s = Dump()                                   ' Number only
        Case 2: Err.Raise 5, "src": s = Dump()                            ' with Source
        Case 3: Err.Raise 5, "src", "desc": s = Dump()                    ' Source and Description
        Case 4: Err.Raise 5, , "desc": s = Dump()                         ' Description only
        Case 5: Err.Raise 5, "s", "d", "hf.hlp", 123: s = Dump()          ' every argument
        Case 6: Err.Raise 6: s = Dump()                                   ' default description: Overflow
        Case 7: Err.Raise 9: s = Dump()                                   ' Subscript out of range
        Case 8: Err.Raise 11: s = Dump()                                  ' Division by zero
        Case 9: Err.Raise 13: s = Dump()                                  ' Type mismatch
        Case 10: Err.Raise 91: s = Dump()                                 ' Object variable not set
        Case 11: Err.Raise vbObjectError + 1000: s = Dump()               ' user number
        Case 12: Err.Raise vbObjectError + 1000, "S", "D": s = Dump()     ' user number with text
        Case 13: Err.Raise 1000: s = Dump()                               ' plain user number
        Case 14: Err.Raise 0: s = Dump()                                  ' number 0: error 5
        Case 15: Err.Raise 65535: s = Dump()                              ' top of the 16-bit range
        Case 16: Err.Raise 65536: s = Dump()                              ' above 65535
        Case 17: Err.Raise 70000: s = Dump()
        Case 18: Err.Raise &H80004005: s = Dump()                         ' negative Long (E_FAIL)
        Case 19: Err.Raise -5: s = Dump()                                 ' small negative
        Case 20: Err.Raise vbObjectError: s = Dump()                      ' vbObjectError itself
        Case 21: Err.Raise 5, "s", "d": Err.Clear: s = Dump()             ' Clear resets everything
        Case 22: Err.Number = 11: s = Dump()                              ' write Number, read Description
        Case 23: Err.Number = 1000: s = Dump()
        Case 24: Err.Description = "abc": s = Dump()                      ' write Description without Raise
        Case 25: Err.Source = "x": Err.HelpFile = "h.hlp": Err.HelpContext = 7: s = Dump()
        Case 26: Err.Number = 11: Err.Description = "mine": s = Dump()    ' own description wins
        Case 27: Err.Raise 5, "s", "d": x = 1: s = Dump()                 ' a later good statement keeps Err
        Case 28: Err.Raise 5, "s", "d": On Error GoTo 0: s = Dump()       ' On Error GoTo 0 after an error
        Case 29: Err.Raise 5, "s", "d": On Error Resume Next: s = Dump()  ' On Error Resume Next after an error
        Case 30: z = 0: x = 1 / z: s = Dump()                             ' real runtime error: Source and Description
        Case 31: x = CInt(40000): s = Dump()                              ' Overflow from CInt
        Case 32: x = CLng("abc"): s = Dump()                              ' Type mismatch from CLng
        Case 33: s = TypeName(Err.LastDllError)                           ' LastDllError type
        Case 34: s = TypeName(Err) & "," & TypeName(Err.Number) & "," & TypeName(Err.Description) & "," & TypeName(Err.HelpContext)
        Case 35: v = "abc": Err.Raise v: s = Dump()                       ' non-numeric Number
        Case 36: Err.Raise 5, Null: s = Dump()                            ' Null Source
        Case 37: Err.Raise 5, 123: s = Dump()                             ' numeric Source
        Case 38: Err.Raise 5, , Null: s = Dump()                          ' Null Description
        Case 39: Err.Raise 5, "s1", "d1": Err.Raise 5, "s2": s = Dump()   ' second Raise keeps old description?
        Case 40: Err.Raise 1000, , "d": Err.Clear: s = Dump()
        Case 41: Err.Raise 5, , "d", , 55: s = Dump()                     ' HelpContext without HelpFile
        Case 42: Err.Raise 429: s = Dump()
        Case 43: Err.Raise 53: s = Dump()
        Case 44: Err.Raise 438: s = Dump()
        Case 45: Err.Raise 424: s = Dump()
        Case 46: Err.Raise 7: s = Dump()
        Case 47: Err.Raise 457: s = Dump()
        Case 48: Err.Raise 5, "s", "d", "hf.hlp", 123: Err.Raise 6: s = Dump()    ' Raise after Raise: help fields reset?
        Case 49: s = Dump()                                               ' fresh state after Clear
        Case 50: s = H1()                                                 ' Resume Next in a handler
        Case 51: s = H2()                                                 ' Exit Function after an error
        Case 52: s = H3(1) & "|" & H3(2)                                  ' error propagated from a callee
        Case 53: s = H4() & "|" & H5()                                    ' Clear in a handler; Exit Sub in a handler
        Case 54: s = H6()                                                 ' error inside a handler
    End Select
    G = s
End Function

Private Function T(ByVal k As Long) As String
    T = "ErrObject." & k & "=" & G(k)
End Function

Public Function Run() As String
    Dim k As Long, s As String
    For k = 1 To N
        s = s & T(k) & vbLf
    Next
    Run = s
End Function
