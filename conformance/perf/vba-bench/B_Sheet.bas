Attribute VB_Name = "B_Sheet"
Option Explicit

Public Function RangeBlock(ByVal n As Long) As Double
    Dim cells() As Variant, area As Range, readback As Variant, row As Long
    Set area = ThisWorkbook.Worksheets(1).Range("A1").Resize(n, 1)
    ReDim cells(1 To n, 1 To 1)
    For row = 1 To n
        cells(row, 1) = row
    Next row
    area.Value2 = cells
    readback = area.Value2
    For row = 1 To n
        RangeBlock = RangeBlock + readback(row, 1)
    Next row
End Function

Public Function RangeCells(ByVal n As Long) As Double
    Dim row As Long, area As Range
    Set area = ThisWorkbook.Worksheets(1).Range("A1").Resize(n, 1)
    For row = 1 To n
        area.Cells(row, 1).Value2 = row
    Next row
    For row = 1 To n
        RangeCells = RangeCells + area.Cells(row, 1).Value2
    Next row
End Function

Public Function InvoiceUdf(ByVal amount As Double) As Double
    InvoiceUdf = amount * 2# + 1#
End Function

Public Function UdfRecalculate(ByVal n As Long) As Double
    Dim area As Range
    Set area = ThisWorkbook.Worksheets(1).Range("B1").Resize(n, 1)
    area.Dirty
    area.Calculate
    UdfRecalculate = Application.WorksheetFunction.Sum(area)
End Function
