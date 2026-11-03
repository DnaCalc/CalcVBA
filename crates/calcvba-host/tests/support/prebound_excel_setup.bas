Public Function App(ByVal window As Object) As Object
    Set App = window.Application
End Function
Public Function WorkbookName(ByVal app As Object) As String
    WorkbookName = app.ActiveWorkbook.FullName
End Function
Public Function Block(ByVal app As Object) As Object
    Set Block = app.ActiveWorkbook.Worksheets(1).Range("A1:B2")
End Function
Public Function Cell(ByVal app As Object) As Object
    Set Cell = app.ActiveWorkbook.Worksheets(1).Range("A1")
End Function
Public Function CellA2(ByVal app As Object) As Object
    Set CellA2 = app.ActiveWorkbook.Worksheets(1).Range("A2")
End Function
Public Function CellB1(ByVal app As Object) As Object
    Set CellB1 = app.ActiveWorkbook.Worksheets(1).Range("B1")
End Function
Public Function DateCell(ByVal app As Object) As Object
    Set DateCell = app.ActiveWorkbook.Worksheets(1).Range("C1")
End Function
Public Function CurrencyCell(ByVal app As Object) As Object
    Set CurrencyCell = app.ActiveWorkbook.Worksheets(1).Range("D1")
End Function
