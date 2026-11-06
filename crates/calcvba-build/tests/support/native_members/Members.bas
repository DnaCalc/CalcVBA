Attribute VB_Name = "Members"
Option Explicit
Public created As Long, destroyed As Long
Public Function MemberSuite() As String
    Dim items As New Collection, iface As ICounter
    Dim receiver As Object, dictionary As Object, value As Variant, number As Long
    On Error GoTo Failed
10 Set iface = New Worker
20 number = 3
30 MemberSuite = CStr(iface.Apply(number)) & ":" & CStr(number)
40 Set iface = New OtherWorker
50 number = 3
60 MemberSuite = MemberSuite & "|" & CStr(iface.Apply(number)) & ":" & CStr(number) & ":" & CStr(TypeOf iface Is ICounter)
70 items.Add "a", "first"
80 items.Add 2, Key:="second"
90 Set receiver = items
100 receiver.Add 3, "third"
110 MemberSuite = MemberSuite & "|" & TypeName(items) & ":" & CStr(items.Count) & ":" & CStr(items("first")) & ":"
120 For Each value In items
130 MemberSuite = MemberSuite & CStr(value)
140 Next value
150 items.Remove "second"
160 MemberSuite = MemberSuite & ":" & CStr(items.Count)
170 Set items = Nothing
180 MemberSuite = MemberSuite & ":" & CStr(items.Count)
190 CallByName receiver, "Add", vbMethod, "z", "last"
200 MemberSuite = MemberSuite & "|" & CStr(CallByName(receiver, "Item", vbMethod, "last")) & ":" & CStr(CallByName(receiver, "Count", vbMethod))
210 Set dictionary = CreateObject("Scripting.Dictionary")
220 CallByName dictionary, "Add", vbMethod, "x", 4
230 CallByName dictionary, "Item", vbLet, "x", 9
240 MemberSuite = MemberSuite & "|" & CStr(dictionary.Count) & ":" & CStr(CallByName(dictionary, "Item", vbGet, "x"))
    On Error Resume Next
    value = CallByName(receiver, "Count", vbGet)
    MemberSuite = MemberSuite & "|" & CStr(Err.Number)
    Err.Clear
    On Error GoTo Failed
    MemberSuite = MemberSuite & ":" & CStr(CallByName(receiver, "Count", vbMethod))
    Exit Function
Failed:
    MemberSuite=MemberSuite & "|fault:" & CStr(Erl) & ":" & CStr(Err.Number) & ":" & Err.Description
End Function
Public Function ConstructorFaultSuite() As String
    Dim failed As BadCtor
    created = 0
    destroyed = 0
    On Error Resume Next
    Set failed = New BadCtor
    ConstructorFaultSuite = CStr(Err.Number) & ":" & TypeName(failed) & ":" & CStr(created) & ":" & CStr(destroyed)
End Function
Public Sub Main()
    Dim observed As String
    observed = MemberSuite() & "#" & ConstructorFaultSuite() & "#" & NothingSuite() & "#" & ProjectCallByNameSuite()
End Sub

Public Function NothingSuite() As String
    Dim receiver As Object, value As Long
    On Error Resume Next
    value = receiver.Count
    NothingSuite = CStr(Err.Number)
    Err.Clear
End Function

Public Function ProjectCallByNameSuite() As String
    Dim receiver As Object
    Set receiver=New Worker
    CallByName receiver,"Name",vbLet,"named"
    ProjectCallByNameSuite=CStr(CallByName(receiver,"Name",vbGet)) & ":" & CStr(CallByName(receiver,"Echo",vbMethod))
End Function
