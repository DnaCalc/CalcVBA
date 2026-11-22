Attribute VB_Name = "BenchExcel"
Option Explicit
' One warm-up followed by five timed calls. The runner appends RunAll from cases.tsv.
Private Declare PtrSafe Function QueryPerformanceCounter Lib "kernel32" (ByRef c As LongLong) As Long
Private Declare PtrSafe Function QueryPerformanceFrequency Lib "kernel32" (ByRef f As LongLong) As Long
Private Const REPS As Long = 5

Private Function RunWorkload(ByVal name As String, ByVal n As Long, ByVal path As String, ByVal functions As Object) As Double
    Select Case name
        Case "LoopLong": RunWorkload = B_Work.LoopLong(n)
        Case "LoopDouble": RunWorkload = B_Work.LoopDouble(n)
        Case "StrBuild": RunWorkload = B_Work.StrBuild(n)
        Case "StrFuncs": RunWorkload = B_Work.StrFuncs(n)
        Case "ArrayScan": RunWorkload = B_Work.ArrayScan(n)
        Case "ArraySort": RunWorkload = B_Work.ArraySort(n)
        Case "Coll": RunWorkload = B_Work.Coll(n)
        Case "Dict": RunWorkload = B_Work.Dict(n)
        Case "ClassCalls": RunWorkload = B_Work.ClassCalls(n)
        Case "LateClassCalls": RunWorkload = B_Work.LateClassCalls(n)
        Case "VariantMix": RunWorkload = B_Work.VariantMix(n)
        Case "Amort": RunWorkload = B_Work.Amort(n)
        Case "UdtArray": RunWorkload = B_Work.UdtArray(n)
        Case "WorksheetSum": RunWorkload = B_Work.WorksheetSum(n, functions)
        Case "FileLines": RunWorkload = B_Work.FileLines(n, path)
        Case "Recursion": RunWorkload = B_Work.Recursion(n)
        Case "RangeBlock": RunWorkload = B_Sheet.RangeBlock(n)
        Case "RangeCells": RunWorkload = B_Sheet.RangeCells(n)
        Case "UdfRecalculate": RunWorkload = B_Sheet.UdfRecalculate(n)
        Case Else: Err.Raise 5, , "Unknown workload " & name
    End Select
End Function

Private Function Bench(ByVal name As String, ByVal n As Long, ByVal path As String) As String
    Dim f As LongLong, t0 As LongLong, t1 As LongLong, r As Long, k As Long, j As Long
    Dim ms(1 To REPS) As Double, tmp As Double, sum As Double, runs As String, checksum As String
    Dim functions As Object
    Set functions = Application.WorksheetFunction
    If name = "UdfRecalculate" Then
        ThisWorkbook.Worksheets(1).Range("B1").Resize(n, 1).FormulaR1C1 = "=InvoiceUdf(ROW())"
    End If
    If QueryPerformanceFrequency(f) = 0 Then Err.Raise 5, , "QueryPerformanceFrequency failed"
    sum = RunWorkload(name, n, path, functions)
    checksum = ProbeFmt.Hex8(sum)
    For r = 1 To REPS
        QueryPerformanceCounter t0
        sum = RunWorkload(name, n, path, functions)
        QueryPerformanceCounter t1
        If ProbeFmt.Hex8(sum) <> checksum Then Err.Raise 5, , "Unstable result for " & name
        ms(r) = (t1 - t0) * 1000# / f
        runs = runs & vbTab & Milliseconds(ms(r))
    Next r
    For k = 1 To REPS - 1
        For j = k + 1 To REPS
            If ms(j) < ms(k) Then tmp = ms(k): ms(k) = ms(j): ms(j) = tmp
        Next j
    Next k
    Bench = name & vbTab & CStr(n) & vbTab & checksum & vbTab & Milliseconds(ms((REPS + 1) \ 2)) & runs & vbLf
End Function

Private Function Milliseconds(ByVal ms As Double) As String
    Milliseconds = Trim$(Str$(ms))
End Function
