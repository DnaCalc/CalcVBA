Attribute VB_Name = "SpecK"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' A class with no default member, used by SpecExpr.
Public n As Long

Public Function Arr() As Variant
    Arr = Array(10, 20)
End Function

Public Property Get Items() As Collection
    Dim c As New Collection
    c.Add "first": c.Add "second"
    Set Items = c
End Property

Public Sub Bump(n As Long)
    n = n + 1
End Sub

Public Property Get Left() As Single
    Left = 1.5
End Property

Public Function UseLeft() As String
    UseLeft = Left("abcd", 2)
End Function

Public Function PassMe() As String
    PassMe = SpecExpr.TakeAny(Me)
End Function
