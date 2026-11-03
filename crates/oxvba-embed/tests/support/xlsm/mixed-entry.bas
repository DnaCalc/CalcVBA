Public Function RunProbe() As String
    Dim ws As Object
    Set ws = ThisWorkbook.Worksheets(1)
    RunProbe = Api.RunProbe() & ";chains=" & ChainOps.ObjectChains(ws)
    RunProbe = RunProbe & ";" & ChainOps.VariantChains(ws)
    RunProbe = RunProbe & ";" & ChainOps.ChainErrors(ws) & ";" & ChainOps.ChainSnapshot(ws)
    ws.Range("A1").NumberFormat = "0.00"
    ws.Range("B2").Interior.Color = 255
    ws.Range("B2").Font.Bold = True
    ws.Range("H1").Formula = "=1/0"
End Function
