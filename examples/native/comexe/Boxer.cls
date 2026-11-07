VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "Boxer"
Attribute VB_Creatable = True
Attribute VB_Exposed = True
Option Explicit
Private held As Variant
Public Function EchoVariant(ByVal v As Variant) As Variant
    EchoVariant = v
End Function
Public Function Describe(ByVal v As Variant) As String
    Describe = TypeName(v) & ":" & CStr(v)
End Function
Public Sub Fill(ByRef v As Variant)
    v = "filled:" & TypeName(v)
End Sub
Public Property Get Stored() As Variant
    Stored = held
End Property
Public Property Let Stored(ByVal v As Variant)
    held = v
End Property
Public Function CountOf(ByVal o As Object) As Long
    CountOf = o.Count
End Function
Public Function Pass(ByVal o As Object) As Object
    Set Pass = o
End Function
Public Function Pick(ByVal a As Long, Optional ByVal b As Variant, Optional ByVal c As Variant) As String
    Dim t As String
    t = CStr(a)
    If IsMissing(b) Then t = t & ":missing" Else t = t & ":" & CStr(b)
    If IsMissing(c) Then t = t & ":missing" Else t = t & ":" & CStr(c)
    Pick = t
End Function
