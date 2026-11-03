Public Function RunProbe() As String
    Dim target As Excel.Range
    Set target = ThisWorkbook.Worksheets(1).Range("F1")
    RunProbe = FormatKinds(target)
    Set target = ThisWorkbook.Worksheets(1).Range("F1:G2")
    RunProbe = RunProbe & ";shape=" & RangeShape(target)
    RunProbe = RunProbe & ";formula=" & RangeFormula(target)
    RunProbe = RunProbe & ";app=" & ApplicationOptions(Application)
    RunProbe = RunProbe & ";range=" & RangeIteration(target)
    RunProbe = RunProbe & ";first=" & RangeFirst(target)
    RunProbe = RunProbe & ";ensure=" & EnsureSheets(ThisWorkbook)
    RunProbe = RunProbe & ";sheets=" & SheetIteration(ThisWorkbook)
    RunProbe = RunProbe & ";books=" & BookIteration(Application)
    RunProbe = RunProbe & ";items=" & BookItem(Application, ThisWorkbook) & ":" & SheetItem(ThisWorkbook)
    RunProbe = RunProbe & ";nothing=" & NothingIteration()
End Function
