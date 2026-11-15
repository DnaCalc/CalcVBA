Attribute VB_Name = "CalcVBALoader"
' The VBA loader for CalcVBA in Excel (docs/design/excel-inprocess-loaders-2026-11-15.md).
' Copies this workbook's other standard, class and document modules (each document with its
' workbook, sheet or chart) into a CalcVBA project, references
' Excel's type library, registers Application as the instance of Excel.Global, and runs one
' entry on VM3 (backend 0) or the JIT (backend 1). The harness replaces <DLL> with the path of
' a fresh copy of calcvba_excel.dll.
Option Explicit

Private Declare PtrSafe Function calcvba_new Lib "<DLL>" (ByRef handle As LongPtr) As Long
Private Declare PtrSafe Function calcvba_add_module Lib "<DLL>" (ByVal handle As LongPtr, ByVal name As LongPtr, ByVal kind As Long, ByVal source As LongPtr) As Long
Private Declare PtrSafe Function calcvba_add_document Lib "<DLL>" (ByVal handle As LongPtr, ByVal name As LongPtr, ByVal hostType As LongPtr, ByVal source As LongPtr, ByVal document As Object) As Long
Private Declare PtrSafe Function calcvba_add_reference Lib "<DLL>" (ByVal handle As LongPtr, ByVal name As LongPtr, ByVal libid As LongPtr, ByVal major As Long, ByVal minor As Long) As Long
Private Declare PtrSafe Function calcvba_set_app_object Lib "<DLL>" (ByVal handle As LongPtr, ByVal cls As LongPtr, ByVal obj As Object) As Long
Private Declare PtrSafe Function calcvba_compile Lib "<DLL>" (ByVal handle As LongPtr, ByVal backend As Long) As Long
Private Declare PtrSafe Function calcvba_run Lib "<DLL>" (ByVal handle As LongPtr, ByVal moduleName As LongPtr, ByVal entry As LongPtr, ByRef result As Variant) As Long
Private Declare PtrSafe Function calcvba_last_error Lib "<DLL>" (ByVal handle As LongPtr, ByRef number As Long, ByRef description As Variant) As Long
Private Declare PtrSafe Function calcvba_close Lib "<DLL>" (ByVal handle As LongPtr) As Long

Private Const EXCEL_GLOBAL As String = "{00020812-0000-0000-C000-000000000046}"
Private Const EXCEL_LIBID As String = "{00020813-0000-0000-C000-000000000046}"

' The workbook, sheet or chart whose document module is codeName.
Private Function DocumentObject(ByVal codeName As String) As Object
    Dim sheet As Object
    If ThisWorkbook.CodeName = codeName Then Set DocumentObject = ThisWorkbook: Exit Function
    For Each sheet In ThisWorkbook.Sheets
        If sheet.CodeName = codeName Then Set DocumentObject = sheet: Exit Function
    Next
End Function

Private Function Failure(ByVal handle As LongPtr, ByVal hr As Long, ByVal what As String) As String
    Dim number As Long, description As Variant
    If hr = 0 Then Exit Function
    calcvba_last_error handle, number, description
    Failure = what & " failed (hr=&H" & Hex(hr) & ", error " & number & "): " & description
End Function

' Runs moduleName.entry (a Function) in CalcVBA; returns its value as text, or the failure.
Public Function RunCalcVBA(ByVal moduleName As String, ByVal entry As String, ByVal backend As Long) As String
    Dim handle As LongPtr, comp As Object, code As String, message As String, result As Variant
    Dim document As Object
    If calcvba_new(handle) <> 0 Then RunCalcVBA = "calcvba_new failed": Exit Function
    For Each comp In ThisWorkbook.VBProject.VBComponents
        If comp.Name <> "CalcVBALoader" And comp.Name <> "CalcVBACompare" Then
            code = ""
            If comp.CodeModule.CountOfLines > 0 Then code = comp.CodeModule.Lines(1, comp.CodeModule.CountOfLines)
            If comp.Type = 1 Or comp.Type = 2 Then
                message = Failure(handle, calcvba_add_module(handle, StrPtr(comp.Name), comp.Type, StrPtr(code)), "add_module " & comp.Name)
            ElseIf comp.Type = 100 Then
                Set document = DocumentObject(comp.Name)
                If document Is Nothing Then
                    message = "no workbook, sheet or chart has the code name " & comp.Name
                Else
                    message = Failure(handle, calcvba_add_document(handle, StrPtr(comp.Name), StrPtr("Excel." & TypeName(document)), StrPtr(code), document), "add_document " & comp.Name)
                End If
            End If
            If message <> "" Then GoTo Finish
        End If
    Next
    message = Failure(handle, calcvba_add_reference(handle, StrPtr("Excel"), StrPtr(EXCEL_LIBID), 1, 9), "add_reference")
    If message <> "" Then GoTo Finish
    message = Failure(handle, calcvba_set_app_object(handle, StrPtr(EXCEL_GLOBAL), Application), "set_app_object")
    If message <> "" Then GoTo Finish
    message = Failure(handle, calcvba_compile(handle, backend), "compile")
    If message <> "" Then GoTo Finish
    message = Failure(handle, calcvba_run(handle, StrPtr(moduleName), StrPtr(entry), result), "run")
    If message = "" Then message = CStr(result)
Finish:
    calcvba_close handle
    RunCalcVBA = message
End Function
