Attribute VB_Name = "Chains"
Option Explicit
Public observed As String
Public Sub Main()
    observed = ChainSuite()
End Sub
Public Function ControlByName(ByVal caption As String) As IChain
    Set ControlByName = New ChainControl
End Function
Public Function ChainSuite() As String
    Dim control As IChain, core As ChainCore, direct As Long, chained As Long
    Set control = ControlByName("test")
    Set core = control.Core
    direct = core.MeasureTextCount + control.OnEnter()
    chained = ControlByName("test").Core.MeasureTextCount + ControlByName("test").OnEnter()
    ChainSuite = CStr(direct) & ":" & CStr(chained)
    Dim obj As Object
    Set obj = core
    On Error Resume Next
    direct = obj.MeasureTextCount
    ChainSuite = ChainSuite & ":" & CStr(Err.Number)
    Err.Clear
    direct = CallByName(core, "MeasureTextCount", VbGet)
    ChainSuite = ChainSuite & ":" & CStr(Err.Number)
    Err.Clear
    Set obj = control
    direct = obj.IChain_OnEnter()
    ChainSuite = ChainSuite & ":" & CStr(Err.Number)
End Function
Public Function CaseSuite() As String
    Dim key As Long, value As Long, result As String
    For key = 0 To 4
        value = 0
        Select Case key
        Case 0: value = 3: value = value + 1
        Case 1 To 2: value = 8
        Case Is > 3: value = 16
        Case Else: value = 12
        End Select
        result = result & CStr(value) & ":"
    Next
    CaseSuite = result
End Function
