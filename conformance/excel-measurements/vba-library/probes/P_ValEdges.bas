Attribute VB_Name = "P_ValEdges"
Option Explicit
' Excel-backed Val overflow, radix-width and character handling neighbors.
Private Const N As Long = 75
Private Function T(ByVal k As Long) As String
    Dim v As Variant
    On Error GoTo fail
    Select Case k
        Case 1: v = Val("&HFFFFFFFF")
        Case 2: v = Val("&H100000000")
        Case 3: v = Val("&H100000001")
        Case 4: v = Val("&H1FFFFFFFF")
        Case 5: v = Val("&HFFFFFFFFF")
        Case 6: v = Val("&H100008000")
        Case 7: v = Val("&H10000FFFF")
        Case 8: v = Val("&HFFFFFFFFFFFFFFFF")
        Case 9: v = Val("&H10000000000000000")
        Case 10: v = Val("&O40000000000")
        Case 11: v = Val("&O40000000001")
        Case 12: v = Val("&H100000000&")
        Case 13: v = Val("&H100000000^")
        Case 14: v = Val("&H100000001^")
        Case 15: v = Val("&H8000%")
        Case 16: v = Val("&H10000%")
        Case 17: v = Val("&HFFFFFFFF&")
        Case 18: v = Val("-&HFFFFFFFF")
        Case 19: v = Val("-&H100000001")
        Case 20: v = Val("&H100000000!")
        Case 21: v = Val("1e400")
        Case 22: v = Val("-1e400")
        Case 23: v = Val("1e-400")
        Case 24: v = Val("1.7976931348623157E308")
        Case 25: v = Val("1.7976931348623159E308")
        Case 26: v = Val("1" & ChrW(9) & "2")
        Case 27: v = Val("1" & ChrW(10) & "2")
        Case 28: v = Val("1" & ChrW(11) & "2")
        Case 29: v = Val("1" & ChrW(12) & "2")
        Case 30: v = Val("1" & ChrW(13) & "2")
        Case 31: v = Val("1" & ChrW(32) & "2")
        Case 32: v = Val("1" & ChrW(160) & "2")
        Case 33: v = Val("1" & ChrW(5760) & "2")
        Case 34: v = Val("1" & ChrW(8192) & "2")
        Case 35: v = Val("1" & ChrW(8199) & "2")
        Case 36: v = Val("1" & ChrW(8203) & "2")
        Case 37: v = Val("1" & ChrW(8239) & "2")
        Case 38: v = Val("1" & ChrW(8287) & "2")
        Case 39: v = Val("1" & ChrW(12288) & "2")
        Case 40: v = Val(CVErr(5))
        Case 41: v = Val(CVErr(65535))
        Case 42: v = Val(True)
        Case 43: v = Val(Empty)
        Case 44: v = Val(Null)
        Case 45: v = Val(Array(1))
        Case 46: v = Val("&17")
        Case 47: v = Val("&H00000000000000001")
        Case 48: v = Val("-&O10")
        Case 49: v = Val("+&H7F")
        Case 50: v = Val("&HFFFFFFFF!")
        Case 51: v = Val("&HFFFFFFFF#")
        Case 52: v = Val("&HFFFFFFFF@")
        Case 53: v = Val("&HFFFFFFFF^")
        Case 54: v = Val("&HFFFFFFFF%")
        Case 55: v = Val("&H100000000%")
        Case 56: v = Val("&H100000000#")
        Case 57: v = Val("&H100000000@")
        Case 58: v = Val("&H8000000000000000^")
        Case 59: v = Val("&HFFFFFFFFFFFFFFFF^")
        Case 60: v = Val("&HFFFFFFFFFFFFFFFF!")
        Case 61: v = Val("&H10000000000000000^")
        Case 62: v = Val("&H10000000000000000!")
        Case 63: v = Val("&HFFFF%")
        Case 64: v = Val("&H80000000")
        Case 65: v = Val("&H00008000")
        Case 66: v = Val("&H00010000")
        Case 67: v = Val("&H000000000001FFFF")
        Case 68: v = Val("&O2000000000000000000000")
        Case 69: v = Val("&HFF$")
        Case 70: v = Val("&HFFZ")
        Case 71: v = Val("&H0.5")
        Case 72: v = Val("12.5%")
        Case 73: v = Val("32768%")
        Case 74: v = Val("1.5!")
        Case 75: v = Val("1e40!")
    End Select
    T = "ValEdges." & k & "=" & Sh(v)
    Exit Function
fail:
    T = "ValEdges." & k & "=ERR" & Err.Number
End Function
Public Function Run() As String
    Dim k As Long, s As String
    For k = 1 To N
        s = s & T(k) & vbLf
    Next
    Run = s
End Function
