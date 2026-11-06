Attribute VB_Name = "Indexing"
Option Explicit
Public observed As String
Public Sub Main()
    observed = IndexSuite()
End Sub
Public Function IndexSuite() As String
    Dim value As Variant, item As Indexed, other As Indexed, result As Object, indexed As Variant
    value = Split("red/blue", "/")
    IndexSuite = value(1)
    value(0) = "green"
    IndexSuite = IndexSuite & ":" & value(0)
    Set item = New Indexed
    Set value = item
    value(1) = "changed"
    indexed = value(1)
    IndexSuite = IndexSuite & "|" & indexed
    Set other = New Indexed
    Set value(2) = other
    Set result = value(2)
    IndexSuite = IndexSuite & "|" & CStr(result Is other)
    Set value = New Collection
    value.Add "entry"
    IndexSuite = IndexSuite & "|" & value(1)
    Set value = CreateObject("Scripting.Dictionary")
    value(1) = "dictionary"
    IndexSuite = IndexSuite & "|" & value(1)
    value = 3
    On Error Resume Next
    IndexSuite = IndexSuite & value(0)
    IndexSuite = IndexSuite & "|" & CStr(Err.Number)
End Function
