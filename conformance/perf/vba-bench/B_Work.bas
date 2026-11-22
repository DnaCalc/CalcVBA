Attribute VB_Name = "B_Work"
Option Explicit
' VBA performance benchmark workloads (dc-tey4.7). Each function does n units of work and returns
' a Double checksum, so Excel VBA and CalcVBA can be compared for correctness as well as time.
' Both runners read the workload sizes from cases.tsv.

Private Type Rec
    Id As Long
    Amount As Double
    Name As String
End Type

Public Function LoopLong(ByVal n As Long) As Double
    Dim i As Long, acc As Long
    For i = 1 To n
        acc = (acc + (i Mod 17) * 3) Mod 1000003
    Next i
    LoopLong = acc
End Function

Public Function LoopDouble(ByVal n As Long) As Double
    Dim i As Long, x As Double, acc As Double
    For i = 1 To n
        x = i * 0.5
        acc = acc + Sqr(x) / (1# + x)
    Next i
    LoopDouble = acc
End Function

Public Function StrBuild(ByVal n As Long) As Double
    Dim i As Long, s As String
    For i = 1 To n
        s = s & Chr$(65 + i Mod 26)
    Next i
    StrBuild = Len(s) + Asc(Mid$(s, n \ 2, 1))
End Function

Public Function StrFuncs(ByVal n As Long) As Double
    Dim i As Long, s As String, t As String, acc As Double
    s = "The quick brown fox jumps over the lazy dog"
    For i = 1 To n
        t = UCase$(Mid$(s, 1 + i Mod 20, 10))
        acc = acc + InStr(1, s, "o", vbBinaryCompare) + Len(Replace(t, "O", "00")) + Len(Trim$(" " & t & " "))
    Next i
    StrFuncs = acc
End Function

Private Sub QSort(a() As Double, ByVal lo As Long, ByVal hi As Long)
    Dim i As Long, j As Long, p As Double, t As Double
    i = lo: j = hi: p = a((lo + hi) \ 2)
    Do While i <= j
        Do While a(i) < p: i = i + 1: Loop
        Do While a(j) > p: j = j - 1: Loop
        If i <= j Then
            t = a(i): a(i) = a(j): a(j) = t
            i = i + 1: j = j - 1
        End If
    Loop
    If lo < j Then QSort a, lo, j
    If i < hi Then QSort a, i, hi
End Sub

Public Function ArraySort(ByVal n As Long) As Double
    Dim a() As Double, i As Long, seed As Long, acc As Double
    ReDim a(1 To n)
    seed = 12345
    For i = 1 To n
        seed = (seed * 1103 + 12345) Mod 65521
        a(i) = seed / 7#
    Next i
    QSort a, 1, n
    For i = 1 To n Step 97
        acc = acc + a(i) * i
    Next i
    ArraySort = acc
End Function

Public Function Coll(ByVal n As Long) As Double
    Dim c As New Collection, i As Long, v As Variant, acc As Double
    For i = 1 To n
        c.Add i * 2, "k" & i
    Next i
    For i = 1 To n Step 3
        acc = acc + c("k" & i)
    Next i
    For Each v In c
        acc = acc + v
    Next v
    Coll = acc + c.Count
End Function

Public Function Dict(ByVal n As Long) As Double
    Dim d As Object, i As Long, acc As Double
    Set d = CreateObject("Scripting.Dictionary")
    For i = 1 To n
        d("k" & i) = i
    Next i
    For i = 1 To n * 2 Step 2
        If d.Exists("k" & i) Then acc = acc + d("k" & i)
    Next i
    Dict = acc + d.Count
End Function

Public Function ClassCalls(ByVal n As Long) As Double
    Dim a As B_Acc, i As Long
    Set a = New B_Acc
    For i = 1 To n
        a.AddValue i Mod 7
    Next i
    ClassCalls = a.Total + a.Calls
End Function

Public Function VariantMix(ByVal n As Long) As Double
    Dim i As Long, v As Variant, w As Variant, acc As Variant
    acc = 0
    For i = 1 To n
        v = i
        w = "2"
        If i Mod 2 = 0 Then v = CDbl(i) / 4
        acc = acc + v * w - (i Mod 3)
    Next i
    VariantMix = CDbl(acc)
End Function

Public Function Amort(ByVal n As Long) As Double
    ' Corpus-style: monthly loan schedules with the financial and date functions.
    Dim i As Long, m As Long, rate As Double, pv As Double, pay As Double, acc As Double
    Dim d As Date
    For i = 1 To n
        rate = (5 + i Mod 7) / 1200#
        pv = 10000 + (i Mod 50) * 1000
        pay = Pmt(rate, 12, -pv)
        d = DateSerial(2024, 1 + i Mod 12, 1)
        For m = 1 To 12
            acc = acc + Round(IPmt(rate, m, 12, -pv), 2)
            d = DateAdd("m", 1, d)
        Next m
        acc = acc + pay + Month(d)
    Next i
    Amort = acc
End Function

Public Function UdtArray(ByVal n As Long) As Double
    Dim r() As Rec, i As Long, k As Long, acc As Double
    ReDim r(1 To 1000)
    For i = 1 To 1000
        r(i).Id = i
        r(i).Name = "R" & i
    Next i
    For i = 1 To n
        k = 1 + i Mod 1000
        r(k).Amount = r(k).Amount + i * 0.25
    Next i
    For i = 1 To 1000
        acc = acc + r(i).Amount + Len(r(i).Name)
    Next i
    UdtArray = acc
End Function

Public Function ArrayScan(ByVal n As Long) As Double
    Dim a() As Long, i As Long, acc As Double
    ReDim a(1 To n)
    For i = 1 To n
        a(i) = i Mod 101
    Next i
    For i = n To 1 Step -1
        acc = acc + a(i) * (i Mod 13)
    Next i
    ArrayScan = acc
End Function

Public Function LateClassCalls(ByVal n As Long) As Double
    Dim a As Object, i As Long
    Set a = New B_Acc
    For i = 1 To n
        a.AddValue CDbl(i Mod 7)
    Next i
    LateClassCalls = a.Total + a.Calls
End Function

Public Function WorksheetSum(ByVal n As Long, ByVal functions As Object) As Double
    Dim i As Long, acc As Double
    For i = 1 To n
        acc = acc + functions.Sum(i, 1 + i Mod 7)
    Next i
    WorksheetSum = acc
End Function

Public Function FileLines(ByVal n As Long, ByVal path As String) As Double
    Dim f As Integer, i As Long, line As String, acc As Double
    f = FreeFile
    Open path For Output As #f
    For i = 1 To n
        Print #f, "invoice-" & CStr(i)
    Next i
    Close #f
    Open path For Input As #f
    For i = 1 To n
        Line Input #f, line
        acc = acc + Len(line)
    Next i
    Close #f
    Kill path
    FileLines = acc
End Function

Private Function Fib(ByVal n As Long) As Long
    If n < 2 Then
        Fib = n
    Else
        Fib = Fib(n - 1) + Fib(n - 2)
    End If
End Function

Public Function Recursion(ByVal n As Long) As Double
    Recursion = Fib(n)
End Function
