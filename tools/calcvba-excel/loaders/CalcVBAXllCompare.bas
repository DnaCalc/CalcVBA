Attribute VB_Name = "CalcVBACompare"
' Runs one probe entry in real VBA, then through the XLL loader (CALCVBA.RUN) on VM3 and on the
' JIT. The XLL compiles the probe module's source on its own, with no VBA project around it.
' The harness replaces <DLL>, <MODULE> and <ENTRY>.
Option Explicit

Public Function CompareAll() As String
    Dim code As String, comp As Object, vba As String
    If Not Application.RegisterXLL("<DLL>") Then CompareAll = "RegisterXLL refused <DLL>": Exit Function
    Set comp = ThisWorkbook.VBProject.VBComponents("<MODULE>")
    code = comp.CodeModule.Lines(1, comp.CodeModule.CountOfLines)
    vba = Application.Run("<MODULE>.<ENTRY>")
    CompareAll = "== VBA" & vbLf & vba & "== VM3" & vbLf & Application.Run("CALCVBA.RUN", code, "<MODULE>", "<ENTRY>", 0) & vbLf _
        & "== JIT" & vbLf & Application.Run("CALCVBA.RUN", code, "<MODULE>", "<ENTRY>", 1) & vbLf
End Function
