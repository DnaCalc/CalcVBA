Attribute VB_Name = "BadInitialize"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private Sub Class_Initialize()
    Err.Raise 457, "BadInitialize", "initialization failed", "init.chm", 73
End Sub
Public Function Ping() As Long
    Ping = 42
End Function
