Attribute VB_Name = "P_Calendar"
Option Explicit
' Calendar property (VBA.DateTime; MS-VBAL 6.1.2.x).
Private Const N As Long = 15

Private Function T(ByVal k As Long) As String
    Dim v As Variant, a As Variant, b As Variant, c As Variant
    On Error GoTo fail
    Select Case k
        Case 1: v = Calendar   ' default is Gregorian (0)
        Case 2: v = TypeName(Calendar)   ' type of the property
        Case 3: Calendar = 0: v = Calendar   ' set Gregorian and read back
        Case 4: Calendar = vbCalGreg: v = Calendar   ' named constant
        Case 5: Calendar = "0": v = Calendar   ' numeric string assigned
        Case 6: Calendar = 2: v = Calendar: Calendar = 0   ' 2 is out of range: error 5
        Case 7: Calendar = -1: v = Calendar: Calendar = 0   ' negative: error 5
        Case 8: Calendar = 1: v = Calendar: Calendar = 0   ' Hijri accepted; reset to 0 in the same case
        Case 9: Calendar = 1: v = Year(DateSerial(2020, 1, 1)): Calendar = 0   ' Year under Hijri; reset to 0 in the same case
        Case 10: Calendar = 1: v = TypeName(Month(DateSerial(2020, 6, 15))): Calendar = 0   ' Month type under Hijri; reset to 0
        Case 11: Calendar = 1.4: v = Calendar: Calendar = 0   ' fractional value rounds to 1; reset to 0
        Case 12: a = Null: Calendar = a: v = Calendar: Calendar = 0   ' Null assigned: error 94
        Case 13: Calendar = "abc": v = Calendar: Calendar = 0   ' non-numeric string: error 13
        Case 14: v = VBA.Calendar   ' qualified read
        Case 15: Calendar = 0: v = Year(DateSerial(2020, 1, 1))   ' Gregorian year after reset
    End Select
    T = "Calendar." & k & "=" & Sh(v)
    Exit Function
fail:
    Dim en As Long
    en = Err.Number
    On Error Resume Next
    Calendar = 0
    T = "Calendar." & k & "=ERR" & en
End Function

Public Function Run() As String
    Dim k As Long, s As String
    For k = 1 To N
        s = s & T(k) & vbLf
    Next
    Run = s
End Function
