Public Function RunProbe() As String
    Dim ws As Object
    Set ws = ThisWorkbook.Worksheets(1)
    RunProbe = ObjectChains(ws) & ";" & VariantChains(ws) & ";" & ChainErrors(ws) & ";" & ChainSnapshot(ws)
End Function
