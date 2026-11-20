Attribute VB_Name = "BenchExcel"
Option Explicit
' Times each B_Work workload in Excel VBA (dc-tey4.7): QueryPerformanceCounter around the call,
' REPS runs, median. One line per workload: name=checksum|median ms|all runs ms.
' The sizes must match SIZES in crates/calcvba-differential/tests/vba_bench.rs.
Private Declare PtrSafe Function QueryPerformanceCounter Lib "kernel32" (ByRef c As LongLong) As Long
Private Declare PtrSafe Function QueryPerformanceFrequency Lib "kernel32" (ByRef f As LongLong) As Long
Private Const REPS As Long = 5

Private Function Call1(ByVal name As String, ByVal n As Long) As Double
    Select Case name
        Case "LoopLong": Call1 = B_Work.LoopLong(n)
        Case "LoopDouble": Call1 = B_Work.LoopDouble(n)
        Case "StrBuild": Call1 = B_Work.StrBuild(n)
        Case "StrFuncs": Call1 = B_Work.StrFuncs(n)
        Case "ArraySort": Call1 = B_Work.ArraySort(n)
        Case "Coll": Call1 = B_Work.Coll(n)
        Case "Dict": Call1 = B_Work.Dict(n)
        Case "ClassCalls": Call1 = B_Work.ClassCalls(n)
        Case "VariantMix": Call1 = B_Work.VariantMix(n)
        Case "Amort": Call1 = B_Work.Amort(n)
        Case "UdtArray": Call1 = B_Work.UdtArray(n)
    End Select
End Function

Private Function Bench(ByVal name As String, ByVal n As Long) As String
    Dim f As LongLong, t0 As LongLong, t1 As LongLong, r As Long, k As Long, j As Long
    Dim ms(1 To REPS) As Double, tmp As Double, sum As Double, runs As String
    QueryPerformanceFrequency f
    For r = 1 To REPS
        QueryPerformanceCounter t0
        sum = Call1(name, n)
        QueryPerformanceCounter t1
        ms(r) = (t1 - t0) * 1000# / f
        runs = runs & IIf(r > 1, ",", "") & Format$(ms(r), "0.000")
    Next r
    For k = 1 To REPS - 1
        For j = k + 1 To REPS
            If ms(j) < ms(k) Then tmp = ms(k): ms(k) = ms(j): ms(j) = tmp
        Next j
    Next k
    Bench = name & "=" & CStr(sum) & "|" & Format$(ms((REPS + 1) \ 2), "0.000") & "|" & runs & vbLf
End Function

Public Function RunAll() As String
    Dim s As String
    s = s & Bench("LoopLong", 5000000)
    s = s & Bench("LoopDouble", 3000000)
    s = s & Bench("StrBuild", 50000)
    s = s & Bench("StrFuncs", 200000)
    s = s & Bench("ArraySort", 200000)
    s = s & Bench("Coll", 100000)
    s = s & Bench("Dict", 100000)
    s = s & Bench("ClassCalls", 1000000)
    s = s & Bench("VariantMix", 1000000)
    s = s & Bench("Amort", 20000)
    s = s & Bench("UdtArray", 1000000)
    RunAll = s
End Function
