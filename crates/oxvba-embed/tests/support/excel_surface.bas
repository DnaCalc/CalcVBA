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
Public Function RangeIteration(ByVal target As Excel.Range) As String
    Dim cell As Excel.Range, listing As String
    target.Value2 = 1#
    For Each cell In target
        listing = listing & cell.Address(False, False) & ":" & cell.Value2 & ";"
        cell.Offset(0, 1).Value2 = cell.Value2 + 10#
    Next
    RangeIteration = listing & "nothing=" & (cell Is Nothing) & ":" & VarType(cell)
End Function
Public Function RangeFirst(ByVal target As Excel.Range) As String
    Dim cell As Excel.Range
    For Each cell In target
        RangeFirst = cell.Address(False, False)
        Exit For
    Next
    RangeFirst = RangeFirst & ":" & cell.Address(False, False)
End Function
Public Function EnsureSheets(ByVal book As Excel.Workbook) As Long
    Dim first As Excel.Worksheet, added As Excel.Worksheet
    Set first = book.Worksheets.Item(1)
    first.Name = "Run13First"
    If book.Worksheets.Count = 1 Then
        Set added = book.Worksheets.Add(After:=first)
        added.Name = "Run13Second"
    End If
    EnsureSheets = book.Worksheets.Count
End Function
Public Function SheetIteration(ByVal book As Excel.Workbook) As String
    Dim sh As Excel.Worksheet, listing As String, seen As Long
    For Each sh In book.Worksheets
        seen = seen + 1
        listing = listing & sh.Name & ";"
    Next
    SheetIteration = seen & ":" & listing & "nothing=" & (sh Is Nothing)
End Function
Public Function BookIteration(ByVal app As Excel.Application) As String
    Dim book As Excel.Workbook, seen As Long
    For Each book In app.Workbooks
        seen = seen + 1
    Next
    BookIteration = seen & ":" & (book Is Nothing)
End Function
Public Function BookItem(ByVal app As Excel.Application, ByVal book As Excel.Workbook) As String
    Dim selected As Excel.Workbook
    Set selected = app.Workbooks.Item(book.Name)
    BookItem = selected.Name
End Function
Public Function SheetItem(ByVal book As Excel.Workbook) As String
    Dim selected As Excel.Worksheet
    Set selected = book.Worksheets.Item("Run13Second")
    SheetItem = selected.Name
End Function
Public Function NothingIteration() As Long
    On Error GoTo Failed
    Dim target As Excel.Range, cell As Excel.Range
    For Each cell In target
        NothingIteration = 1
    Next
    Exit Function
Failed:
    NothingIteration = Err.Number
End Function
