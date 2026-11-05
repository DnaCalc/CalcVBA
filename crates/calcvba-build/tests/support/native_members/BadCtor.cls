VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "BadCtor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Sub Class_Initialize()
    Members.created = Members.created + 1
    Err.Raise 513, "Ctor", "failure"
End Sub
Private Sub Class_Terminate()
    Members.destroyed = Members.destroyed + 1
End Sub
