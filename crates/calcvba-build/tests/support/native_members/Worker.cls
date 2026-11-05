VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "Worker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements ICounter
Private Function ICounter_Apply(ByRef value As Long) As Long
    value = value + 2
    ICounter_Apply = value * 3
End Function
