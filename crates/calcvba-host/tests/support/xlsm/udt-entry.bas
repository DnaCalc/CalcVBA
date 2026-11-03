Private Type RowData
    Title As String
    Amount As Double
    Counts(1 To 2) As Long
    Data() As Double
End Type
Public Function RunProbe() As String
    Dim ws As Object, row As RowData, copied As RowData
    Set ws = ThisWorkbook.Worksheets(1)
    row.Title = "summary"
    row.Amount = 12.5
    row.Counts(1) = 3
    row.Counts(2) = 7
    ReDim row.Data(1 To 2, 1 To 2)
    row.Data(1, 1) = 1.25
    row.Data(1, 2) = 2.5
    row.Data(2, 1) = 3.75
    row.Data(2, 2) = 5#
    ReDim Preserve row.Data(1 To 2, 1 To 3)
    copied = row
    copied.Counts(1) = 9
    copied.Data(1, 1) = 99#
    ws.Range("A1").Value2 = copied.Title
    ws.Range("A2").Value2 = copied.Amount
    ws.Range("B1").Value2 = row.Counts(1)
    ws.Range("B2").Value2 = copied.Counts(1)
    ws.Range("D1:E1").Value2 = row.Counts
    ws.Range("G1:I2").Value = row.Data
    ws.Range("K1").Value2 = copied.Data(1, 1)
    ws.Range("K2").Value2 = row.Data(1, 1)
    ws.Range("K3").Value2 = VarType(row.Counts)
    ws.Range("K4").Value2 = VarType(row.Data)
    RunProbe = "ok"
End Function
