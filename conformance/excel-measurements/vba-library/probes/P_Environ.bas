Attribute VB_Name = "P_Environ"
Option Explicit
' Environ / Environ$ (Interaction; values never printed, only lengths and relations).
Private Const N As Long = 30

Private Function T(ByVal k As Long) As String
    Dim v As Variant
    On Error GoTo fail
    Select Case k
        Case 1: v = Len(Environ("TEMP")) > 0                      ' TEMP exists
        Case 2: v = Len(Environ$("TEMP")) > 0                     ' $ form
        Case 3: v = TypeName(Environ("TEMP"))                     ' String
        Case 4: v = TypeName(Environ$("TEMP"))                    ' String
        Case 5: v = InStr(Environ(1), "=") > 0                    ' by number: NAME=value
        Case 6: v = InStr(Environ$(1), "=") > 0                   ' $ form by number
        Case 7: v = TypeName(Environ(1))                          ' String
        Case 8: v = Environ(0)                                    ' error 5
        Case 9: v = Environ(-1)                                   ' error 5
        Case 10: v = Environ$(0)                                   ' error 5
        Case 11: v = Environ("NO_SUCH_VAR_CALCVBA")                ' missing name: ""
        Case 12: v = Environ$("NO_SUCH_VAR_CALCVBA")
        Case 13: v = Environ("temp") = Environ("TEMP")             ' name is case-insensitive
        Case 14: v = Environ("")                                   ' empty name
        Case 15: v = Environ(" ")                                  ' blank name
        Case 16: v = Environ("TEMP=")                              ' name with equals sign
        Case 17: v = Environ(255) = Environ(255.0)                 ' same entry
        Case 18: v = Environ(32000)                                ' number past the last entry: ""
        Case 19: v = Environ(1.5) = Environ(2)                     ' fractional index rounds
        Case 20: v = Environ(2.5) = Environ(2)                     ' banker's rounding
        Case 21: v = Environ("1")                                  ' numeric string is a name
        Case 22: v = Environ(Null)                                 ' Null: error 94 or 13
        Case 23: v = Environ(Empty)                                ' Empty: 0
        Case 24: v = Environ(True)                                 ' -1: error 5
        Case 25: v = Environ("x") = Environ$("x")                  ' both forms agree on a missing name
        Case 26: v = Len(Environ("windir")) > 0                    ' another standard variable
        Case 27: v = Environ(Expression:="NO_SUCH_VAR_CALCVBA")    ' named argument
        Case 28: v = Environ(CByte(1)) = Environ(1)                ' Byte index
        Case 29: v = Environ("TEMP") = Environ(Chr(84) & "EMP")    ' computed name
        Case 30: v = Environ(Array(1))                             ' array: error 13
    End Select
    T = "Environ." & k & "=" & Sh(v)
    Exit Function
fail:
    T = "Environ." & k & "=ERR" & Err.Number
End Function

Public Function Run() As String
    Dim k As Long, s As String
    For k = 1 To N
        s = s & T(k) & vbLf
    Next
    Run = s
End Function
