Attribute VB_Name = "Breadth"
Option Explicit

Public Function StringSuite() As String
    Dim text As String, data As Variant
    text = "  a" & ChrW(&H3BB) & "b  "
    text = UCase$(Trim$(text))
    text = Replace(text, "B", "xy")
    data = Left(text, 2) & ":" & Mid$(text, 3) & ":" & Right$(text, 2)
    StringSuite = data & ":" & CStr(InStr(1, text, "xy")) & ":" & StrReverse("abc")
End Function

Public Function VariantSuite() As String
    Dim a As Variant, b As Variant, value As Variant, result As String
    a = CInt(7): b = CByte(2)
    value = a + b
    result = CStr(VarType(value)) & ":" & CStr(value)
    a = 7.5: b = 2.5
    value = a \ b
    result = result & ":" & CStr(value)
    value = a Mod b
    result = result & ":" & CStr(value)
    a = 2: b = 3
    value = a ^ b
    result = result & ":" & CStr(value)
    a = Null
    value = a & "x"
    result = result & ":" & CStr(IsNull(a + 1)) & ":" & CStr(value)
    value = "null-branch"
    If a Then value = "BAD"
    result = result & ":" & CStr(value) & ":" & CStr(IsNull(Not a))
    value = CVErr(2042)
    result = result & ":" & CStr(IsError(value)) & ":" & CStr(value)
    a = Empty
    value = a + 3
    VariantSuite = result & ":" & CStr(value) & ":" & CStr(IsEmpty(a))
End Function

Public Function ArraySuite() As String
    Dim numbers(1 To 3) As Long, item As Variant, data As Variant
    Dim total As Long, result As String
    numbers(1) = 1: numbers(2) = 2: numbers(3) = 3
    For Each item In numbers
        total = total + item
        If item = 1 Then numbers(2) = 99
    Next
    result = CStr(total)
    data = Array(CInt(2), "x", True)
    For Each item In data
        result = result & ":" & CStr(VarType(item)) & "=" & CStr(item)
    Next
    For Each item In data
        result = result & ":early"
        Exit For
    Next
    ArraySuite = result
End Function

Public Function LibrarySuite() As String
    Dim text As String, parts As Variant
    parts = Split("a,b,c", ",")
    text = Join(parts, "/")
    text = text & ":" & Format$(DateSerial(2024, 2, 29), "yyyy-mm-dd")
    text = text & ":" & Hex$(255) & ":" & String$(3, "x")
    text = text & ":" & CStr(Round(2.5)) & ":" & CStr(Sgn(-8))
    LibrarySuite = text & ":" & CStr(Int(-1.25)) & ":" & CStr(Fix(-1.25))
End Function

Public Function ErrorSuite() As String
    Dim x As Double, result As String
    On Error Resume Next
    x = 1# / 0#
    result = CStr(Err.Number) & "|" & CStr(Err.Source <> "")
    Err.Clear
    result = result & "|" & CStr(Err.Number)
    Err.Raise 513, "Native.Source", "detail", "help.chm", 17
    result = result & "|" & CStr(Err.Number) & "|" & Err.Source & "|" & Err.Description & "|" & Err.HelpFile & "|" & CStr(Err.HelpContext)
    Err.Raise 514
    result = result & "|" & CStr(Err.Number) & "|" & Err.Description
    On Error GoTo 0
    ErrorSuite = result & "|" & CStr(Err.Number)
End Function

Public Function ResumeSuite() As String
    Dim x As Double, divisor As Double, phase As Long, result As String
    On Error GoTo failed
    phase = 1
10  x = 10# / divisor
    phase = 2
20  Error 6
    phase = 3
30  Error 13
done:
    ResumeSuite = result & CStr(Err.Number) & "|" & CStr(x)
    Exit Function
failed:
    result = result & CStr(Err.Number) & ":" & CStr(Erl) & "|"
    If phase = 1 Then
        divisor = 2#
        Resume
    ElseIf phase = 2 Then
        Resume Next
    Else
        Resume done
    End If
End Function

Private Sub NestedFault()
    Err.Raise 519, "Nested.Source", "nested fault"
End Sub

Private Sub HandlerFault()
    On Error GoTo failed
    Error 6
    Exit Sub
failed:
    NestedFault
End Sub

Public Function NestedErrorSuite() As String
    Dim result As String
    On Error GoTo failed
    HandlerFault
    NestedErrorSuite = result & "|" & CStr(Err.Number)
    Exit Function
failed:
    result = CStr(Err.Number) & "|" & Err.Source & "|" & Err.Description
    Resume Next
End Function

Public Function MixedSuite() As String
    Dim text As String, number As Long, value As Variant, result As String
    text = "2": number = 2
    result = CStr(text = number) & "|" & CStr(number = text)
    value = "2"
    result = result & "|" & CStr(value = number)
    value = 2&
    result = result & "|" & CStr(text = value)
    text = "abcde"
    Mid$(text, 2, 2) = "XY"
    result = result & "|" & text
    result = result & "|" & CStr(StrComp(ChrW$(&HC5), ChrW$(&HE5), vbTextCompare))
    On Error Resume Next
    text = "two"
    result = result & CStr(text = number)
    result = result & "|err=" & CStr(Err.Number)
    On Error GoTo 0
    MixedSuite = result
End Function

Public Function ArrayAllocationSuite() As String
Dim a() As Long, item As Variant, result As String, v As Variant
ReDim a(1 To 3): a(1)=1: a(2)=2: a(3)=3
On Error Resume Next
For Each item In a
    If item=1 Then
        ReDim Preserve a(1 To 5)
        result=result & "resize=" & CStr(Err.Number) & ";"
        Err.Clear
    End If
    result=result & CStr(item) & ","
Next
ReDim a(1 To 4)
result=result & "after=" & CStr(Err.Number) & ";"
v=Array(1&,2&,3&)
For Each item In v
    If item=1 Then
        v=Array(8&,9&,10&)
        result=result & "replace=" & CStr(Err.Number) & ";"
        Err.Clear
    End If
    result=result & CStr(item) & ","
Next
ArrayAllocationSuite=result
End Function

Public Function EraseSuite() As String
    Dim fixed(1 To 3) As Long, dynamic() As String, item As Variant, result As String
    fixed(1)=1: fixed(2)=2: fixed(3)=3
    On Error Resume Next
    For Each item In fixed
        If item=1 Then Erase fixed
        result=result & CStr(item) & ":" & CStr(Err.Number) & ","
        Err.Clear
    Next
    Erase fixed
    result=result & "|reset=" & CStr(fixed(2))
    ReDim dynamic(2 To 4)
    dynamic(2)="owned"
    Erase dynamic
    On Error Resume Next
    result=result & "bound=" & CStr(UBound(dynamic))
    result=result & "error=" & CStr(Err.Number)
    On Error GoTo 0
    ReDim dynamic(3 To 5)
    EraseSuite=result & "|" & CStr(LBound(dynamic)) & ":" & CStr(UBound(dynamic))
End Function

Public Function ErrorBoundsSuite() As String
Dim result As String
On Error Resume Next
Err.Raise 0, "Explicit", "zero"
result=CStr(Err.Number) & ":" & CStr(Err.Source <> "Explicit") & ":" & Err.Description
Err.Clear
Error 0
result=result & "|" & CStr(Err.Number) & ":" & Err.Description
Err.Clear
Err.Raise 65536, "Explicit", "large"
result=result & "|" & CStr(Err.Number) & ":" & CStr(Err.Source <> "Explicit") & ":" & Err.Description
Err.Clear
Err.Raise -1, "Explicit", "negative"
ErrorBoundsSuite=result & "|" & CStr(Err.Number) & ":" & Err.Source & ":" & Err.Description
End Function

Public Function ActivationSuite() As String
    Dim dictionary As Object
    Set dictionary = CreateObject("Scripting.Dictionary")
    ActivationSuite = CStr(VarType(dictionary)) & ":" & CStr(Not (dictionary Is Nothing))
End Function

Public Function ForeignObjectSuite(ByVal foreign As Object) As String
    ForeignObjectSuite = CStr(VarType(foreign)) & ":" & CStr(foreign Is Nothing) & "|" & ActivationSuite()
End Function
