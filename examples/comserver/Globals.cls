VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "Globals"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private mValue As Long
Public Function NativeGlobalAnswer() As Long
    NativeGlobalAnswer = 42
End Function
Public Function NextGlobalValue() As Long
    mValue = mValue + 1
    NextGlobalValue = mValue
End Function
