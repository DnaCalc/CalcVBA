Attribute VB_Name = "P_Date"
Option Explicit
' Date / Date$ (VBA.DateTime; MS-VBAL 6.1.2.x).
Private Const N As Long = 14

Private Function T(ByVal k As Long) As String
    Dim v As Variant, a As Variant, b As Variant, c As Variant
    On Error GoTo fail
    Select Case k
        Case 1: v = TypeName(Date)   ' type is Date
        Case 2: v = VarType(Date) = vbDate   ' VarType
        Case 3: v = TypeName(Date$)   ' Date$ returns String
        Case 4: v = Len(Date$)   ' Date$ is mm-dd-yyyy, 10 chars
        Case 5: v = Mid$(Date$, 3, 1) & Mid$(Date$, 6, 1)   ' separators of Date$ are dashes
        Case 6: v = CDbl(Date) = Int(CDbl(Date))   ' no time part
        Case 7: v = CDbl(Date) > 43831   ' later than 2020-01-01
        Case 8: v = Date$ = Format(Date, "mm-dd-yyyy")   ' Date$ matches the mm-dd-yyyy rendering (fails only across midnight)
        Case 9: v = Month(Date) = Val(Left$(Date$, 2))   ' month of Date agrees with Date$ (fails only across midnight)
        Case 10: v = Year(Date) = Val(Right$(Date$, 4))   ' year of Date agrees with Date$
        Case 11: v = Date = Int(Now)   ' equals Int(Now) (fails only across midnight)
        Case 12: v = IsDate(Date$)   ' IsDate on Date$
        Case 13: v = TypeName(Date + 1)   ' Date + Double stays Date
        Case 14: v = TypeName(VBA.Date)   ' qualified call
    End Select
    T = "Date." & k & "=" & Sh(v)
    Exit Function
fail:
    Dim en As Long
    en = Err.Number
    T = "Date." & k & "=ERR" & en
End Function

Public Function Run() As String
    Dim k As Long, s As String
    For k = 1 To N
        s = s & T(k) & vbLf
    Next
    Run = s
End Function
