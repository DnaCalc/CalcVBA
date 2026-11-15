Attribute VB_Name = "CalcVBACompare"
' Runs one probe entry in real VBA, then in CalcVBA on VM3 and on the JIT, in this workbook.
' The harness replaces <MODULE> and <ENTRY>.
Option Explicit

Public Function CompareAll() As String
    Dim vba As String
    vba = Application.Run("<MODULE>.<ENTRY>")
    CompareAll = "== VBA" & vbLf & vba & "== VM3" & vbLf & CalcVBALoader.RunCalcVBA("<MODULE>", "<ENTRY>", 0) & vbLf _
        & "== JIT" & vbLf & CalcVBALoader.RunCalcVBA("<MODULE>", "<ENTRY>", 1) & vbLf
End Function
