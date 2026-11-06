Attribute VB_Name = "Indexed"
Option Explicit
Private text As String
Private held As Object
Public Property Get Item(ByVal key As Long) As Variant
    If key = 2 Then
        Set Item = held
    Else
        Item = CStr(key) & ":" & text
    End If
End Property
Attribute Item.VB_UserMemId = 0
Public Property Let Item(ByVal key As Long, ByVal value As Variant)
    text = CStr(value)
End Property
Attribute Item.VB_UserMemId = 0
Public Property Set Item(ByVal key As Long, ByVal value As Object)
    Set held = value
End Property
Attribute Item.VB_UserMemId = 0
