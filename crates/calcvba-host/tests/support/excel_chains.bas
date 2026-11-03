Public Function ObjectChains(ByVal ws As Object) As String
    Dim cell As Object, boxed As Variant
    ws.Range("A1").Value2 = 7#
    ws.Range("A1").Offset(1).Value = 12#
    Set cell = ws.Range("A1").Offset(1)
    Set boxed = cell
    boxed.Offset(0, 1).Value2 = boxed.Value + ws.Range("A1").Value2
    ws.Range("B3").Formula = "=A1+A2+B2"
    ObjectChains = ws.Range("A1").Offset(1).Address(False, False)
    ObjectChains = ObjectChains & ":" & ws.Range("A2").Value & ":" & boxed.Offset(0, 1).Value2
    ObjectChains = ObjectChains & ":" & ws.Range("B3").Value2
    ObjectChains = ObjectChains & ":" & (ws.Range("A1") + 2#)
    ObjectChains = ObjectChains & ":" & ws.Range("A1").Offset(RowOffset:=1).Address(False, False)
End Function
Public Function VariantChains(ByVal ws As Variant) As String
    Dim target As Variant
    Set target = ws.Range("D1")
    target.Value2 = 3#
    target.Offset(1, 0).Value2 = target.Value2 * 4#
    ws.Cells(2, 5).Value2 = target.Offset(1).Value2 + 2#
    VariantChains = ws.Cells(2, 5).Address(False, False) & ":" & ws.Cells(2, 5).Value2
End Function
Public Function ChainErrors(ByVal ws As Object) As String
    Dim uninitialized As Object, plain As Variant, answer As Variant
    On Error Resume Next
    answer = uninitialized.Value
    ChainErrors = Err.Number
    Err.Clear
    plain = 12
    answer = plain.Value
    ChainErrors = ChainErrors & ":" & Err.Number
    Err.Clear
    answer = ws.NoSuchMember
    ChainErrors = ChainErrors & ":" & Err.Number
End Function
Public Function ChainSnapshot(ByVal ws As Object) As String
    ChainSnapshot = ws.Range("A1").Value2 & ":" & ws.Range("A2").Value2
    ChainSnapshot = ChainSnapshot & ":" & ws.Range("B2").Value2 & ":" & ws.Range("B3").Value2
    ChainSnapshot = ChainSnapshot & ":" & ws.Range("D1").Value2 & ":" & ws.Range("D2").Value2 & ":" & ws.Range("E2").Value2
End Function
