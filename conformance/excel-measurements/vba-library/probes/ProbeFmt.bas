Attribute VB_Name = "ProbeFmt"
Option Explicit
' Shared value formatter for the VBA library probes (conformance/excel-measurements/vba-library).
' Sh(v) gives the type and the exact value: IEEE bits for Double, Single and Date, the full text
' for Currency and Decimal, escaped text for String, elements for small arrays.
Private Type DblT
    d As Double
End Type
Private Type Bits8
    lo As Long
    hi As Long
End Type
Private Type SngT
    s As Single
End Type
Private Type Bits4
    b As Long
End Type

Public Function Hex8(ByVal d As Double) As String
    Dim a As DblT, b As Bits8
    a.d = d
    LSet b = a
    Hex8 = Right$("00000000" & Hex$(b.hi), 8) & Right$("00000000" & Hex$(b.lo), 8)
End Function

Public Function Hex4(ByVal s As Single) As String
    Dim a As SngT, b As Bits4
    a.s = s
    LSet b = a
    Hex4 = Right$("00000000" & Hex$(b.b), 8)
End Function

Public Function Esc(ByVal s As String) As String
    Dim i As Long, c As Long, r As String
    For i = 1 To Len(s)
        c = AscW(Mid$(s, i, 1)) And &HFFFF&
        If c < 32 Or c > 126 Or c = 123 Or c = 125 Then r = r & "{" & c & "}" Else r = r & Mid$(s, i, 1)
    Next
    Esc = r
End Function

Public Function Sh(v As Variant) As String
    Dim r As String, i As Long, n As Long
    If IsObject(v) Then
        If v Is Nothing Then Sh = "Nothing" Else Sh = "obj " & TypeName(v)
        Exit Function
    End If
    If IsArray(v) Then
        On Error Resume Next
        n = -2
        n = UBound(v)
        If n = -2 Then
            Sh = TypeName(v) & "[unallocated]"
            Exit Function
        End If
        On Error GoTo 0
        r = TypeName(v) & "[" & LBound(v) & "," & UBound(v) & "]{"
        For i = LBound(v) To UBound(v)
            If i - LBound(v) >= 12 Then r = r & "...": Exit For
            r = r & Sh(v(i))
            If i < UBound(v) Then r = r & ","
        Next
        Sh = r & "}"
        Exit Function
    End If
    Select Case VarType(v)
        Case vbEmpty: Sh = "Empty"
        Case vbNull: Sh = "Null"
        Case vbString: Sh = "String:[" & Esc(v) & "]"
        Case vbDouble: Sh = "Double:" & Hex8(v) & "|" & CStr(v)
        Case vbSingle: Sh = "Single:" & Hex4(v) & "|" & CStr(v)
        Case vbDate: Sh = "Date:" & Hex8(CDbl(v))
        Case vbError: Sh = "Error:" & CLng(v)
        Case Else: Sh = TypeName(v) & ":" & CStr(v)
    End Select
End Function
