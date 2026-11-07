Attribute VB_Name = "ExecutionRegressions"
Option Explicit
Public observed As String
Private Declare PtrSafe Function lstrlenA Lib "kernel32" (ByVal lpString As LongPtr) As Long
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
Public Function DisagreementSuite() As String
    Dim emptyVar As Variant, nothingObj As Object, dyn() As Long, elemArr As Variant
    Dim big As Double, inf As Double, nan As Double
    Dim multiplyError As Long, divideError As Long, a As Boolean, b As Boolean, compareError As Long
    On Error Resume Next
    With emptyVar
        .Foo = 1
    End With
    DisagreementSuite = CStr(Err.Number)
    Err.Clear
    With emptyVar
        With .inner
            .Foo = 1
        End With
    End With
    DisagreementSuite = DisagreementSuite & ":" & CStr(Err.Number)
    Err.Clear
    With nothingObj
        .Foo = 1
    End With
    DisagreementSuite = DisagreementSuite & ":" & CStr(Err.Number)
    Err.Clear
    ReDim dyn(3 To 2)
    DisagreementSuite = DisagreementSuite & ":" & CStr(Err.Number)
    Err.Clear
    ReDim dyn(0 To 0)
    dyn(0) = 5
    ReDim Preserve dyn(0 To -1)
    DisagreementSuite = DisagreementSuite & ":" & CStr(Err.Number) & ":" & CStr(dyn(0))
    Err.Clear
    elemArr = Array(5)
    MutateFirst elemArr
    DisagreementSuite = DisagreementSuite & ":" & CStr(elemArr(0)) & ":" & CStr(VarType(elemArr(0)))
    Err.Clear
    big = 1E+308
    inf = big * big
    multiplyError = Err.Number
    Err.Clear
    nan = inf / inf
    divideError = Err.Number
    Err.Clear
    a = (inf > 0)
    b = (nan <> nan)
    compareError = Err.Number
    DisagreementSuite = DisagreementSuite & ":" & CStr(multiplyError) & ":" & CStr(divideError) & ":" & CStr(compareError) & ":" & CStr(a) & ":" & CStr(b)
End Function
Private Sub MutateFirst(ParamArray items() As Variant)
    items(0)(0) = 99
End Sub
Public Function ObjectLetSuite() As String
    Dim o As Object, a% As Object
    On Error Resume Next
    o = vbNullString
    ObjectLetSuite = CStr(Err.Number)
    Err.Clear
    o = 1
    ObjectLetSuite = ObjectLetSuite & ":" & CStr(Err.Number)
    Err.Clear
    a = 1
    ObjectLetSuite = ObjectLetSuite & ":" & CStr(Err.Number)
End Function
Public Function ArrayPtrSuite() As String
    Dim buf(0 To 2) As Byte
    buf(0) = 65
    buf(1) = 66
    buf(2) = 0
    ArrayPtrSuite = CStr(lstrlenA(VarPtr(buf(0)))) & ":" & CStr(buf(1))
End Function
Public Function ArrayPtrShapeSuite() As String
    Dim buf(1 To 3) As Byte
    buf(1) = 65
    buf(2) = 66
    buf(3) = 0
    ArrayPtrShapeSuite = CStr(lstrlenA(VarPtr(buf(1)))) & ":" & CStr(LBound(buf)) & ":" & CStr(UBound(buf)) & ":" & CStr(VarType(buf)) & ":" & TypeName(buf)
End Function
Public Function ObjectCompareSuite() As String
    Dim o As Object
    On Error Resume Next
    If o = 1 Then ObjectCompareSuite = "t"
    ObjectCompareSuite = ObjectCompareSuite & CStr(Err.Number)
End Function
