Option Explicit
Option Base 1
Public Type Sample
    Count As Long
    Title As String
    Slots(1 To 2) As Long
End Type
Private entered As Long
Public Function EntryCount() As Long
    EntryCount = entered
End Function
Private Sub BumpRecord(ByRef first As Sample, ByRef second As Sample)
    first.Count = first.Count + 1
    second.Count = second.Count + 1
End Sub
Private Sub TouchTail(ParamArray items() As Variant)
    items(0) = 9
End Sub
Public Function SumTail(ParamArray items() As Variant) As Double
    Dim i As Long
    For i = LBound(items) To UBound(items)
        SumTail = SumTail + items(i)
    Next
End Function
Public Function BoundsTail(ParamArray items() As Variant) As String
    BoundsTail = LBound(items) & ":" & UBound(items)
End Function
Public Function SumFrom(ByVal first As Long, ParamArray items() As Variant) As Double
    Dim i As Long
    SumFrom = first
    For i = LBound(items) To UBound(items)
        SumFrom = SumFrom + items(i)
    Next
End Function
Public Function MakeSample() As Sample
    MakeSample.Count = 41
    MakeSample.Title = "owned"
    MakeSample.Slots(2) = 7
End Function
Public Function AdvanceSample(ByRef value As Sample) As Sample
    entered = entered + 1
    value.Count = value.Count + 1
    AdvanceSample = value
End Function
Public Function AdvanceArray(ByRef items() As Variant) As Variant
    entered = entered + 1
    items(0) = 9
    AdvanceArray = items
End Function
Public Function RecordReturns() As String
    Dim original As Sample, result As Sample
    original = MakeSample()
    result = AdvanceSample(original)
    RecordReturns = original.Count & ":" & result.Count & ":" & result.Title & ":" & result.Slots(2)
End Function
Public Function RunProbe() As String
    Dim rec As Sample, other As Sample, x As Long
    rec.Count = 41
    rec.Title = "owned"
    rec.Slots(2) = 7
    other = rec
    rec.Count = 42
    RunProbe = "record=" & rec.Count & ":" & rec.Title & ":" & rec.Slots(2) & "|copy=" & other.Count
    BumpRecord rec, rec
    RunProbe = RunProbe & "|aliases=" & rec.Count
    x = 5
    TouchTail x
    RunProbe = RunProbe & "|param=" & x & ":" & SumTail(2, 3, 4) & ":" & SumTail() & "|empty=" & BoundsTail()
End Function
