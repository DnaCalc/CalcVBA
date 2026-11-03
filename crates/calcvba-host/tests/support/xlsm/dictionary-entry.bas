Public Function RunProbe() As String
    Dim ws As Object, dict As Object, entries As Variant
    Set ws = ThisWorkbook.Worksheets(1)
    Set dict = CreateObject("Scripting.Dictionary")
    dict.Add "alpha", 10&
    dict.Add "beta", 20&
    dict("alpha") = dict("alpha") + 3&
    entries = dict.Keys
    ws.Range("A1").Value2 = dict.Count
    ws.Range("A2").Value2 = dict.Exists("beta")
    ws.Range("A3").Value2 = dict("alpha")
    ws.Range("A4").Value2 = entries(0)
    dict.Remove "beta"
    ws.Range("A5").Value2 = dict.Count
    RunProbe = "ok"
End Function
