Public Function RunProbe() As String
    Dim target As Excel.Range
    Set target = ThisWorkbook.Worksheets(1).Range("F1")
    RunProbe = FormatKinds(target)
    Set target = ThisWorkbook.Worksheets(1).Range("F1:G2")
    RunProbe = RunProbe & ";shape=" & RangeShape(target)
    RunProbe = RunProbe & ";formula=" & RangeFormula(target)
    RunProbe = RunProbe & ";app=" & ApplicationOptions(Application)
End Function
