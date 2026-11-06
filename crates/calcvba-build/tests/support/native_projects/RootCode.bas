Attribute VB_Name = "RootCode"
Option Explicit
Public observed As String
Private numbers(0 To 1) As Long
Public Sub Main()
    observed = Probe()
End Sub
Public Function Probe() As String
    Leaf.Common.Reset
    numbers(0) = 9
    Probe = CStr(UBound(numbers)) & ":" & Leaf.Common.InitialState() & ":" & CStr(numbers(0)) & "|"
    Probe = Probe & ProjectSuite() & "#" & ProjectSuite()
End Function
Private Function ProjectSuite() As String
    Dim count As Long, importedObj As Leaf.Token, ownToken As Token, value As Long
    count = LeftLib.LeftCode.UseLeaf() + RightLib.RightCode.UseLeaf()
    Set importedObj = Leaf.Common.MakeToken()
    importedObj.PutValue 11
    value = importedObj.GetValue()
    Set ownToken = New Token
    ProjectSuite = CStr(count) & ":" & CStr(Leaf.Common.TotalValue()) & ":" & CStr(value) & ":" & TypeName(importedObj)
    ProjectSuite = ProjectSuite & ":" & CStr(TypeOf ownToken Is Leaf.Token) & ":" & CStr(TypeOf importedObj Is Leaf.Token)
    Set importedObj = Nothing
    ProjectSuite = ProjectSuite & ":" & CStr(Leaf.Common.EndedValue())
    On Error Resume Next
    Leaf.Common.Fail
    ProjectSuite = ProjectSuite & ":" & CStr(Err.Number) & ":" & Err.Source
    Err.Clear
End Function
