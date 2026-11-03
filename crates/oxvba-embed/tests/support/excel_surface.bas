Public Function WindowApplication(ByVal win As Excel.Window) As Excel.Application
    Set WindowApplication = win.Application
End Function
Public Function ActiveBook(ByVal app As Excel.Application) As Excel.Workbook
    Set ActiveBook = app.ActiveWorkbook
End Function
Public Function BookPath(ByVal book As Excel.Workbook) As String
    BookPath = book.FullName
End Function
Public Function FirstSheet(ByVal book As Excel.Workbook) As Excel.Worksheet
    Set FirstSheet = book.Worksheets.Item(1)
End Function
Public Function SheetRange(ByVal sheet As Excel.Worksheet) As Excel.Range
    Set SheetRange = sheet.Range("F1:G2")
End Function
Public Function FormatKinds(ByVal target As Excel.Range) As String
    Dim cell As Excel.Range
    Set cell = target.Cells(1, 1)
    cell.NumberFormat = "yyyy-mm-dd"
    cell.Value2 = 45000#
    FormatKinds = VarType(cell.Value) & ":" & VarType(cell.Value2)
    cell.NumberFormat = "$0.00"
    cell.Value2 = 1.234567
    FormatKinds = FormatKinds & "|" & VarType(cell.Value) & ":" & VarType(cell.Value2)
    FormatKinds = FormatKinds & "|" & cell.NumberFormat & "|" & CStr(cell.Value)
    cell.NumberFormat = "General"
End Function
Public Function RangeShape(ByVal target As Excel.Range) As String
    Dim cell As Excel.Range
    Set cell = target.Cells(2, 1)
    RangeShape = cell.Address & "|" & target.Offset(1, 2).Address
    RangeShape = RangeShape & "|" & target.Resize(3, 1).Address
    RangeShape = RangeShape & "|" & target.Rows.Count & ":" & target.Columns.Count & ":" & target.Count
    RangeShape = RangeShape & "|" & target.Address(False, False)
End Function
Public Function RangeFormula(ByVal target As Excel.Range) As String
    Dim cell As Excel.Range
    Set cell = target.Cells(1, 1)
    cell.NumberFormat = "0.00"
    cell.Formula = "=2+3"
    RangeFormula = cell.Formula & ":" & cell.Value2 & ":" & cell.Text
    cell.FormulaR1C1 = "=RC[1]*2"
    cell.Offset(0, 1).Value2 = 4#
    RangeFormula = RangeFormula & "|" & cell.Formula & ":" & cell.FormulaR1C1 & ":" & cell.Value2
    cell.NumberFormat = "General"
End Function
Public Function ApplicationOptions(ByVal app As Excel.Application) As String
    Dim screen As Boolean, events As Boolean
    screen = app.ScreenUpdating
    events = app.EnableEvents
    app.ScreenUpdating = False
    app.EnableEvents = False
    ApplicationOptions = app.ScreenUpdating & ":" & app.EnableEvents
    app.Calculate
    ApplicationOptions = ApplicationOptions & "|" & app.Evaluate("SUM(2,3)")
    app.EnableEvents = events
    app.ScreenUpdating = screen
End Function
Public Function SetCurrencyFormat(ByVal target As Excel.Range) As Long
    target.NumberFormat = "General"
    target.NumberFormat = "$0.00"
    target.Value2 = 1.234567
    SetCurrencyFormat = 0
End Function
Public Function ReadCurrencyFormat(ByVal target As Excel.Range) As String
    Dim cell As Excel.Range
    Set cell = target.Cells(1, 1)
    ReadCurrencyFormat = VarType(cell.Value) & ":" & VarType(cell.Value2)
    ReadCurrencyFormat = ReadCurrencyFormat & "|" & cell.NumberFormat & "|" & CStr(cell.Value)
End Function
