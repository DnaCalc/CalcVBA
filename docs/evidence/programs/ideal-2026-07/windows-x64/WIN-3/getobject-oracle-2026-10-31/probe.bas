Option Explicit
Public Function RunProbe() As String
    Dim cell As Object
    Dim matrix As Variant
    Dim started As Date
    Dim readyNumber As Integer
    Set cell = ThisWorkbook.Worksheets(1).Range("A1")
    cell.Value2 = 42
    ThisWorkbook.Worksheets(1).Range("B1").Value2 = "alpha"
    ThisWorkbook.Worksheets(1).Range("A2").Value2 = True
    ThisWorkbook.Worksheets(1).Range("B2").Value2 = CVErr(2042)
    readyNumber = FreeFile
    Open "C:\Work\DnaCalc\codex-runs\calcvba-excel-object-probe\ready.txt" For Output As #readyNumber
    Print #readyNumber, ThisWorkbook.FullName
    Print #readyNumber, Application.Hwnd
    Close #readyNumber
    started = Now
    Do While Dir("C:\Work\DnaCalc\codex-runs\calcvba-excel-object-probe\release.txt") = ""
        DoEvents
        If DateDiff("s", started, Now) > 150 Then Err.Raise 5, "OwnedProbe", "External reader did not finish"
    Loop
    matrix = ThisWorkbook.Worksheets(1).Range("A1:B2").Value2
    RunProbe = CStr(cell.Value2) & "|" & cell.Address(False, False) & "|" & CStr(ThisWorkbook.Worksheets.Count) & "|" & CStr(matrix(1, 2)) & "|" & CStr(matrix(2, 1)) & "|" & CStr(matrix(2, 2)) & "|" & CStr(LBound(matrix, 1)) & "|" & CStr(UBound(matrix, 1)) & "|" & CStr(LBound(matrix, 2)) & "|" & CStr(UBound(matrix, 2))
End Function
