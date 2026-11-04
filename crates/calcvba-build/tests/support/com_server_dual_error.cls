Attribute VB_Name = "ScalarErrors"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Function Ping() As Long
    Ping = 42
End Function
Public Function AddPair(ByVal a As Long, ByVal b As Long) As Long
    AddPair = a + b
End Function
Public Function Divide(ByVal a As Double, ByVal b As Double) As Double
    If a = 457 Then Err.Raise 457, "ScalarErrors", "dual failure", "dual.chm", 81
    Divide = a / b
End Function
