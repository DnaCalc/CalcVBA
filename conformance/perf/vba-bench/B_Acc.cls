VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "B_Acc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' Benchmark class: a property and method call per unit of work (dc-tey4.7).
Private mTotal As Double
Private mCalls As Long

Public Sub AddValue(ByVal x As Double)
    mTotal = mTotal + x
    mCalls = mCalls + 1
End Sub

Public Property Get Total() As Double
    Total = mTotal
End Property

Public Property Get Calls() As Long
    Calls = mCalls
End Property
