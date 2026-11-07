Attribute VB_Name = "ExecutionRegressions"
Option Explicit
Public observed As String
Public Sub Main()
    observed = ExecutionSuite()
End Sub
Public Function ExecutionSuite() As String
    Dim small As Variant, wide As Variant, a(0 To 1) As Variant, first As Variant, second As Variant
    small = 3
    wide = 3&
    ExecutionSuite = CStr(VarType(small)) & ":" & CStr(VarType(wide))
    a(0) = 5
    a(1) = 9
    Mutate a(0), a(1)
    ExecutionSuite = ExecutionSuite & ":" & CStr(a(0)) & ":" & CStr(a(1))
    Dim unset As Object, number As Long
    On Error Resume Next
    Dim plain As Variant
    plain = 7
    Set small = plain
    ExecutionSuite = ExecutionSuite & ":" & CStr(Err.Number)
    Err.Clear
    number = CInt(unset)
    ExecutionSuite = ExecutionSuite & ":" & CStr(Err.Number)
    Err.Clear
    Let small = unset
    ExecutionSuite = ExecutionSuite & ":" & CStr(Err.Number)
    Err.Clear
    number = unset
    ExecutionSuite = ExecutionSuite & ":" & CStr(Err.Number)
    Err.Clear
    UseLong unset
    ExecutionSuite = ExecutionSuite & ":" & CStr(Err.Number)
End Function
Private Sub UseLong(ByVal value As Long)
End Sub
Private Sub Mutate(ParamArray values() As Variant)
    values(0) = 11
    values(1) = 13
End Sub
Public Function ArrayOrderSuite() As String
    Dim a(1 To 2, 3 To 4) As Long, b(1 To 2, 3 To 4, 5 To 6) As String
    Dim i As Long, j As Long, k As Long, item As Variant, result As String
    For i = 1 To 2
        For j = 3 To 4
            a(i,j) = i * 10 + j
            For k = 5 To 6
                b(i,j,k) = CStr(i) & CStr(j) & CStr(k)
            Next
        Next
    Next
    For Each item In a
        result = result & CStr(item) & ":"
    Next
    result = result & "|"
    For Each item In b
        result = result & CStr(item) & ":"
    Next
    ArrayOrderSuite = result
End Function
Public Function RandomAritySuite() As String
    Dim discarded As Single, a As Single, b As Single
    discarded = Rnd(-1)
    Randomize 1
    a = Rnd()
    b = Rnd()
    discarded = Rnd(-1)
    Randomize 1
    RandomAritySuite = CStr(a = Rnd()) & ":" & CStr(b = Rnd()) & ":" & CStr(a <> b) & ":" & CStr(Rnd(0) = b)
End Function
Public Function CounterSuite() As String
    Dim i As Variant, wide As Variant
    For i = 1 To 3
        i = i + 1
    Next
    For wide = 1& To 3&
    Next
    CounterSuite = CStr(VarType(i)) & ":" & CStr(i) & ":" & CStr(VarType(wide)) & ":" & CStr(wide)
End Function
Public Function NestedCallSuite() As String
    Dim small As Variant, wide As Variant
    small = AddThree(AddThree(AddThree(1)))
    wide = AddThree(AddThree(AddThree(1&)))
    NestedCallSuite = CStr(VarType(small)) & ":" & CStr(small) & ":" & CStr(VarType(wide)) & ":" & CStr(wide)
End Function
Private Function AddThree(ByRef value As Variant) As Variant
    AddThree = value + 3
End Function
