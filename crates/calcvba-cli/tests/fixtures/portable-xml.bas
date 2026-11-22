Option Explicit

Public Sub Main()
    Debug.Print Probe()
End Sub

Public Function Probe() As String
    Dim doc As Object, other As Object, held As Object, again As Object, missing As Object
    Dim loaded As Boolean, bad As Boolean
    Set doc = CreateObject("mSxMl2.dOmDoCuMeNt")
    Set other = GetObject("", "MSXML2.DOMDocument.6.0")
    doc.async = False
    loaded = doc.loadXML("<root>hello &amp; world</root>")
    Set held = doc.documentElement
    Set again = doc.documentElement
    Probe = CStr(held Is again) & "|" & CStr(doc.async) & "|" & CStr(loaded) & "|" & held.nodeName & "|" & held.Text
    loaded = other.loadXML("<independent/>")
    bad = doc.loadXML("<broken>")
    Set missing = doc.documentElement
    Probe = Probe & "|" & CStr(loaded) & "|" & CStr(bad) & "|" & _
            CStr(missing Is Nothing) & "|" & held.nodeName & "|" & other.documentElement.nodeName
End Function
