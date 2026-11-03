Public Function RunProbe() As String
    Dim book As Object, added As Object, removed As Object
    Set book = ThisWorkbook
    Set added = book.Worksheets.Add(After:=book.Worksheets(book.Worksheets.Count))
    added.Name = "Report"
    added.Range("A1").Value2 = "new sheet"
    Set removed = book.Worksheets.Add
    removed.Name = "Temporary"
    book.Application.DisplayAlerts = False
    removed.Delete
    book.Worksheets(1).Range("A1").Value2 = book.Worksheets.Count
    RunProbe = "ok"
End Function
