Attribute VB_Name = "Expressions"
Option Explicit
Private Type FieldRecord
    number As Long
    text As String
    pointer As LongPtr
End Type
Private calls As Long
Private trace As String
Private Function MakeItems() As Variant
    calls=calls+1: trace=trace & "F"
    MakeItems=Array("left","right")
End Function
Private Function NextIndex() As Long
    trace=trace & "I": NextIndex=0
End Function
Private Function NextDimension() As Long
    trace=trace & "D": NextDimension=1
End Function
Public Function ArraySuite() As String
    calls=0: trace=""
    ArraySuite=CStr(UBound(Split("a,b,c",","))) & ":" & CStr(LBound(Split("a,b",","))) & ":" & Split("a,b",",")(1)
    ArraySuite=ArraySuite & ":" & MakeItems()(NextIndex()) & ":" & CStr(calls) & ":" & trace
    trace=""
    ArraySuite=ArraySuite & ":" & CStr(UBound(MakeItems(),NextDimension())) & ":" & trace
End Function
Private Function CharacterTrace(ByVal binaryMode As Boolean) As String
    Dim f As Long, text As String, i As Long, n As Long
    f=FreeFile
    If binaryMode Then
        Open "@INPUT_PATH@" For Binary As #f
    Else
        Open "@INPUT_PATH@" For Input As #f
    End If
    On Error Resume Next
    For i=1 To 12
        Err.Clear
        text=Input$(1,#f)
        n=Err.Number
        If n<>0 Then
            CharacterTrace=CharacterTrace & "E" & CStr(n) & "@" & CStr(Seek(f))
            Exit For
        End If
        CharacterTrace=CharacterTrace & CStr(AscW(text)) & "@" & CStr(Seek(f)) & ","
    Next i
    Err.Clear
    text=Input(0,#f)
    CharacterTrace=CharacterTrace & "|zero=" & CStr(Err.Number) & ":" & CStr(Len(text))
    Err.Clear
    text=Input(-1,#f)
    CharacterTrace=CharacterTrace & "|negative=" & CStr(Err.Number)
    Close #f
    Err.Clear
    text=Input(1,#511)
    CharacterTrace=CharacterTrace & "|handle=" & CStr(Err.Number)
    Err.Clear
    If binaryMode Then
        Open "@INPUT_PATH@" For Binary As #f
    Else
        Open "@INPUT_PATH@" For Input As #f
    End If
    text=Input(99,#f)
    CharacterTrace=CharacterTrace & "|over=" & CStr(Err.Number) & ":" & CStr(Seek(f)) & ":" & CStr(Len(text))
    Close #f
End Function
Public Function InputSuite() As String
    InputSuite=CharacterTrace(False) & "#" & CharacterTrace(True)
End Function
Public Function ObjectIndexSuite() As String
    ObjectIndexSuite=MakeCollection()(1) & ":" & CStr(MakeDictionary()("x"))
End Function
Private Function MakeCollection() As Variant
    Dim c As New Collection
    c.Add "collection"
    Set MakeCollection=c
End Function
Private Function MakeDictionary() As Variant
    Dim d As Object
    Set d=CreateObject("Scripting.Dictionary")
    d.Add "x",8
    Set MakeDictionary=d
End Function
Public Function ScopeSuite() As String
    Dim item As Nullable, opaque As Object, n As Long
    Set item=New Nullable
    ScopeSuite=item.Defaults() & "|" & CStr(item.Internal(7)) & ":" & CStr(item.ReadInternal())
    Set opaque=item
    On Error Resume Next
    n=opaque.Internal(7)
    ScopeSuite=ScopeSuite & "|" & CStr(Err.Number)
    Err.Clear
    n=CallByName(opaque,"Hidden",vbMethod)
    ScopeSuite=ScopeSuite & ":" & CStr(Err.Number)
    Err.Clear
End Function
Public Function ControlSuite() As String
    ControlSuite=ControlTrace("@CONTROL_PATH@")
End Function
Public Function BufferedControlSuite() As String
    BufferedControlSuite=ControlTrace("@BUFFERED_PATH@")
End Function
Private Function ControlTrace(ByVal path As String) As String
    Dim f As Long, i As Long, part As String, number As Long
    f=FreeFile
    Open path For Input As #f
    On Error Resume Next
    For i=1 To 4
        part=Input(1,#f)
        number=Err.Number
        Err.Clear
        If number=0 Then
            ControlTrace=ControlTrace & CStr(AscW(part))
        Else
            ControlTrace=ControlTrace & "E" & CStr(number)
        End If
        ControlTrace=ControlTrace & "@" & CStr(Seek(f)) & ":" & CStr(EOF(f)) & ","
    Next i
    Close #f
End Function
Public Function ByValFieldSuite() As String
    Dim value As FieldRecord
    value.number=7&
    value.text="field"
    value.pointer=4294967297^
    ByValFieldSuite=NumberField(value.number) & ":" & TextField(value.text) & ":" & PointerField(value.pointer)
End Function
Private Function NumberField(ByVal value As Long) As String
    NumberField=CStr(value+1&)
End Function
Private Function TextField(ByVal value As String) As String
    TextField=value & "!"
End Function
Private Function PointerField(ByVal value As LongPtr) As String
    PointerField=CStr(value)
End Function
Public Sub Main()
    Dim observed As String
    observed=ArraySuite() & "#" & InputSuite() & "#" & ObjectIndexSuite() & "#" & ScopeSuite() & "#" & ByValFieldSuite() & "#" & ControlSuite() & "#" & BufferedControlSuite()
End Sub
