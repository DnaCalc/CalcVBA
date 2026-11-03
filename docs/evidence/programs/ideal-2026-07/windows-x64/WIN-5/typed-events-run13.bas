Public Function RunProbe() As String
    Dim changedRange As Excel.Range
    Set changedRange = ThisWorkbook.Worksheets(1).Range("F10")
    oracleEventsActive = True
    RunProbe = "off=" & StartChange(changedRange, True)
    RunProbe = RunProbe & ";on=" & StartChange(changedRange, False)
    RunProbe = RunProbe & ";close=" & TryClose(ThisWorkbook)
    oracleEventsActive = False
End Function
