Attribute VB_Name = "ChainControl"
Option Explicit
Implements IChain
Private Property Get IChain_Core() As ChainCore
    Set IChain_Core = New ChainCore
End Property
Private Function IChain_OnEnter() As Long
    IChain_OnEnter = 19
End Function
