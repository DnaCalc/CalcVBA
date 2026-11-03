Public Function RunProbe() As String
    Dim cell As Excel.Range, block As Excel.Range, functions As Excel.WorksheetFunction
    Dim value As Variant, packet(1 To 2, 1 To 2) As Variant
    Dim route As Long, parts As String, code As Long
    Set cell = ThisWorkbook.Worksheets(1).Range("A1")
    Set block = ThisWorkbook.Worksheets(1).Range("A1:B2")
    Set functions = Application.WorksheetFunction
    cell.NumberFormat = "General"
    WriteRange cell, 42
    RunProbe = "scalar=" & ReadRange(cell) & ":"
    WriteValue2 cell, 42
    RunProbe = RunProbe & ReadValue2(cell) & ":"
    WriteDefault cell, 42
    RunProbe = RunProbe & ReadDefault(cell) & ":" & RuntimeDefault(cell)
    WriteRange block, 42
    value = ReadRange(block)
    RunProbe = RunProbe & "|fill=" & value(1, 1) & ":"
    WriteValue2 block, 42
    value = ReadValue2(block)
    RunProbe = RunProbe & value(2, 1) & ":"
    WriteDefault block, 42
    value = ReadDefault(block)
    RunProbe = RunProbe & value(2, 2)
    packet(1, 1) = 1.25
    packet(2, 1) = True
    packet(1, 2) = "alpha"
    packet(2, 2) = CVErr(2042)
    For route = 0 To 2
        Select Case route
        Case 0
            WriteRange block, packet
            value = ReadRange(block)
        Case 1
            WriteValue2 block, packet
            value = ReadValue2(block)
        Case 2
            WriteDefault block, packet
            value = ReadDefault(block)
        End Select
        If LBound(value, 1) <> 1 Or UBound(value, 1) <> 2 Or _
           LBound(value, 2) <> 1 Or UBound(value, 2) <> 2 Then Err.Raise 5
        If route > 0 Then parts = parts & ";"
        parts = parts & value(1, 1) & ":" & value(2, 1) & ":" & _
            value(1, 2) & ":" & CStr(value(2, 2))
    Next
    RunProbe = RunProbe & "|arrays=" & parts
    packet(1, 1) = 1: packet(2, 1) = 2
    packet(1, 2) = 3: packet(2, 2) = 4
    WriteValue2 block, packet
    RunProbe = RunProbe & "|functions=" & SumArray(functions, 7) & ":" & _
        SumArray(functions, Array(2, 3)) & ":" & SumDifferentArities(functions, 2, 3) & ":" & SumRange(Application, block)
    RunProbe = RunProbe & "|more=" & AverageArray(functions, Array(2, 3)) & ":" & _
        SumTyped(functions, 2, 3) & ":" & MissingMatch(functions, Array("a", "b"))
    On Error Resume Next
    Err.Clear
    value = functions.Match("missing", Array("a", "b"), 0)
    code = Err.Number
    On Error GoTo 0
    RunProbe = RunProbe & "|bad=" & code
    cell.NumberFormat = "yyyy-mm-dd"
    cell.Value2 = 45000
    RunProbe = RunProbe & "|date=" & VarType(ReadRange(cell)) & ":" & VarType(ReadValue2(cell))
    cell.NumberFormat = "$0.00"
    cell.Value2 = 1.234567
    RunProbe = RunProbe & "|currency=" & VarType(ReadRange(cell)) & ":" & VarType(ReadValue2(cell))
    RunProbe = RunProbe & "|kinds=" & ScalarKinds(cell)
    WriteRange cell, True
    value = ReadRange(cell)
    RunProbe = RunProbe & "|scalars=" & VarType(value) & ":" & CStr(value)
    WriteValue2 cell, "text"
    value = ReadValue2(cell)
    RunProbe = RunProbe & ":" & VarType(value) & ":" & CStr(value)
    WriteDefault cell, CVErr(2042)
    value = ReadDefault(cell)
    RunProbe = RunProbe & ":" & VarType(value) & ":" & CStr(value)
    WriteDefault cell, Empty
    RunProbe = RunProbe & ":" & VarType(ReadDefault(cell))
    Set cell = ThisWorkbook.Worksheets(1).Range("C1")
    cell.NumberFormat = "yyyy-mm-dd"
    cell.Value2 = 45000
    RunProbe = RunProbe & "|prepared=" & VarType(ReadRange(cell))
    Set cell = ThisWorkbook.Worksheets(1).Range("D1")
    cell.NumberFormat = "$0.00"
    cell.Value2 = 1.234567
    RunProbe = RunProbe & ":" & VarType(ReadRange(cell))
End Function
