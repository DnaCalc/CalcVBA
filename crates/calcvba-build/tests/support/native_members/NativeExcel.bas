Attribute VB_Name = "NativeExcel"
Option Explicit
Public Function ExcelSuite(ByVal app As Excel.Application) As String
    Dim book As Excel.Workbook, sheet As Excel.Worksheet
    Set book = app.Workbooks.Add()
    Set sheet = book.Worksheets.Item(1)
    sheet.Range("A1:B2").Value2 = 7
    ExcelSuite = app.Name & ":" & CStr(app.WorksheetFunction.Sum(sheet.Range("A1:B2"))) & ":" & CStr(TypeOf app Is Excel.Application)
    book.Close False
End Function
