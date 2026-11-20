Attribute VB_Name = "VariantIs"
Option Explicit
' dc-yqq0.15: the native lane's Is takes each Variant operand's object (Object coercion), not the cell header.
Public Function SameObject() As Boolean
    Dim a As Variant, b As Variant
    Set a = New Collection: Set b = a
    SameObject = (a Is b) And Not (a Is New Collection)
End Function
