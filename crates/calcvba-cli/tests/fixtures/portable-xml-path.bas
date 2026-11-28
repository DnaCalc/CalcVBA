Sub Main()
    Dim doc As Object
    Dim node As Object
    Dim child As Object
    Dim missing As Object
    Set doc = CreateObject("Msxml2.DOMDocument.6.0")
    doc.async = False
    If Not doc.LoadXML("<root><group/><group><item>first</item><item>second</item></group></root>") Then Exit Sub
    Set node = doc.SelectSingleNode("/root/group/item")
    Set child = doc.DocumentElement.SelectSingleNode(".//item")
    Set missing = doc.SelectSingleNode("//missing")
    Debug.Print node.nodeName & "|" & node.Text & "|" & child.Text & "|" & CStr(missing Is Nothing) & "|" & CStr(node Is child)
End Sub
