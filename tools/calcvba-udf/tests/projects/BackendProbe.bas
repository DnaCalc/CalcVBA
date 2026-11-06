Attribute VB_Name = "BackendProbe"
Option Explicit
Private sharedApplication As Object
Public BackendResult As String
Public Property Get Application() As Object
    Set Application=sharedApplication
End Property
Public Sub Main()
    Dim owner As Object, cells As Excel.Range
    Set owner=GetObject("@RUNNING_OBJECT_PATH@")
    Set sharedApplication=owner.Application
    Set cells=sharedApplication.ActiveWorkbook.Worksheets(1).Range("A1:A2")
    cells.Cells(1,1).Value2=2#
    cells.Cells(2,1).Value2=3#
    BackendResult="owned-range=" & CStr(Functions.RangeSum(cells)) & ":" & Functions.ReferenceKind(cells)
    If BackendResult<>"owned-range=5:Range:2" Then Err.Raise 513&,"BackendProbe",BackendResult
End Sub
