Attribute VB_Name = "Keyed"
Option Explicit
Private stored As String
Public Property Get Item(ByVal key As String) As String
    Item = key & ":" & stored
End Property
Attribute Item.VB_UserMemId = 0
Public Property Let Item(ByVal key As String, ByVal value As String)
    stored = value
End Property
Attribute Item.VB_UserMemId = 0
