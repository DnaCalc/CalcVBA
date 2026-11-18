Attribute VB_Name = "P_Partition"
Option Explicit
' Partition (Interaction; range text padded to the width of Stop).
Private Const N As Long = 44

Private Function T(ByVal k As Long) As String
    Dim v As Variant
    Dim va As Variant, i As Integer
    On Error GoTo fail
    Select Case k
        Case 1: v = Partition(5, 0, 99, 10)           ' standard: " 0: 9"
        Case 2: v = Partition(100, 0, 200, 10)        ' " 100: 109"
        Case 3: v = Partition(0, 0, 99, 10)           ' first range
        Case 4: v = Partition(99, 0, 99, 10)          ' last range
        Case 5: v = Partition(55, 0, 99, 10)
        Case 6: v = Partition(-1, 0, 99, 10)          ' below Start: " :-1"
        Case 7: v = Partition(100, 0, 99, 10)         ' above Stop: "100: "
        Case 8: v = Partition(1000, 0, 99, 10)
        Case 9: v = Partition(5, 0, 99, 1)            ' interval 1
        Case 10: v = Partition(5, 0, 10, 1)
        Case 11: v = Partition(10, 0, 10, 1)           ' Stop equals Number
        Case 12: v = Partition(95, 0, 100, 10)         ' final range clipped to Stop
        Case 13: v = Partition(100, 0, 100, 10)
        Case 14: v = Partition(0, 0, 100, 10)
        Case 15: v = Partition(5, 0, 99, 200)          ' interval larger than the range
        Case 16: v = Partition(5, 0, 1000, 100)
        Case 17: v = Partition(550, 0, 1000, 100)
        Case 18: v = Partition(5, 1, 99, 10)           ' Start not 0
        Case 19: v = Partition(15, 1, 99, 10)
        Case 20: v = Partition(0, 1, 99, 10)           ' below Start
        Case 21: v = Partition(5, 0, 99, 0)            ' Interval 0: error 5
        Case 22: v = Partition(5, 0, 99, -1)           ' negative interval: error 5
        Case 23: v = Partition(5, -1, 99, 10)          ' negative Start: error 5
        Case 24: v = Partition(5, 99, 0, 10)           ' Stop below Start: error 5
        Case 25: v = Partition(5, 10, 10, 1)           ' Stop equal to Start: error 5
        Case 26: v = Partition(Null, 0, 99, 10)        ' Null number: Null
        Case 27: v = Partition(5, Null, 99, 10)        ' Null Start: error 94
        Case 28: v = Partition(5, 0, Null, 10)
        Case 29: v = Partition(5, 0, 99, Null)
        Case 30: v = Partition(Empty, 0, 99, 10)       ' Empty number: 0
        Case 31: v = Partition("5", 0, 99, 10)         ' numeric string
        Case 32: v = Partition("x", 0, 99, 10)         ' error 13
        Case 33: v = Partition(5.5, 0, 99, 10)         ' fractional number
        Case 34: v = Partition(5, 0, 99.5, 10)         ' fractional Stop
        Case 35: v = Partition(5, 0.5, 99, 10)         ' fractional Start
        Case 36: v = Partition(5, 0, 99, 10.5)         ' fractional Interval
        Case 37: v = Partition(2147483647, 0, 2147483647, 1000)   ' Long extremes
        Case 38: v = Partition(2147483647, 0, 2147483647, 1)
        Case 39: v = Partition(5, 0, 2147483648#, 10)  ' Stop beyond Long: overflow or error 5
        Case 40: v = Partition(1000000, 0, 1000000, 100000)
        Case 41: v = Partition(-5, -10, 10, 5)         ' negative Start: error 5
        Case 42: v = Partition(5, 0, 9, 3)             ' Stop not on a boundary
        Case 43: v = Partition(9, 0, 9, 3)
        Case 44: v = TypeName(Partition(5, 0, 99, 10)) ' String
    End Select
    T = "Partition." & k & "=" & Sh(v)
    Exit Function
fail:
    T = "Partition." & k & "=ERR" & Err.Number
End Function

Public Function Run() As String
    Dim k As Long, s As String
    For k = 1 To N
        s = s & T(k) & vbLf
    Next
    Run = s
End Function
