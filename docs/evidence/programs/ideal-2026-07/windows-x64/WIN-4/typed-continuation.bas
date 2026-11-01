Option Explicit
Private journal As String
Private failed As Boolean
Private entries As Long

Private Function HostRead(ByVal member As Long) As Double
    journal = journal & CStr(member) & ","
    If member = 7 And Not failed Then
        failed = True
        Err.Raise 11
    End If
    If member = 7 Then
        HostRead = 42#
    Else
        HostRead = CDbl(member + 1)
    End If
End Function

Private Sub AliasLeaf(ByRef a As Double, ByRef b As Double)
    a = a + HostRead(1)
    b = b + a
End Sub
Private Sub AliasMiddle(ByRef value As Double)
    AliasLeaf value, value
End Sub
Private Function NestedAliases() As Double
    Dim value As Double
    entries = entries + 1
    value = HostRead(0)
    AliasMiddle value
    NestedAliases = value + HostRead(2) + entries
End Function
Private Sub FailingLeaf(ByRef value As Double)
    value = value + 3#
    value = value + HostRead(7)
    value = value + 1000#
End Sub
Private Function CatchNested() As Double
    Dim value As Double
    On Error GoTo Handler
    value = HostRead(0)
    FailingLeaf value
    CatchNested = value + HostRead(2)
    Exit Function
Handler:
    value = value + Err.Number
    value = value + HostRead(8)
    Resume Next
End Function
Private Function RetryStatement() As Double
    Dim value As Double
    On Error GoTo Handler
    entries = entries + 1
    value = HostRead(7)
    RetryStatement = value + entries
    Exit Function
Handler:
    value = HostRead(8)
    Resume
End Function
Private Function SkipStatement() As Double
    Dim value As Double
    On Error Resume Next
    value = HostRead(7)
    value = Err.Number
    SkipStatement = value + HostRead(2)
End Function
Private Sub RebindObject(ByRef first As Range, ByRef second As Range, ByVal replacement As Range)
    Set first = replacement
    Set second = first
End Sub

Public Function RunProbe() As String
    On Error GoTo FailedProbe
    Dim text As String
    Dim value As Double
    Dim original As Range
    Dim replacement As Range
    journal = ""
    failed = False
    entries = 0
    value = NestedAliases()
    text = "alias=" & CStr(value) & ":" & journal
    journal = ""
    failed = False
    value = CatchNested()
    text = text & "|catch=" & CStr(value) & ":" & journal
    journal = ""
    failed = False
    entries = 0
    value = RetryStatement()
    text = text & "|retry=" & CStr(value) & ":" & journal
    journal = ""
    failed = False
    value = SkipStatement()
    text = text & "|skip=" & CStr(value) & ":" & journal
    Set original = ThisWorkbook.Worksheets(1).Range("A1")
    Set replacement = ThisWorkbook.Worksheets(1).Range("B1")
    original.Value2 = 1#
    replacement.Value2 = 42#
    RebindObject original, original, replacement
    text = text & "|object=" & CStr(original.Value2) & ":" & original.Address
    RunProbe = text
    Exit Function
FailedProbe:
    RunProbe = "FAILED:" & CStr(Err.Number) & ":" & Err.Description
End Function
