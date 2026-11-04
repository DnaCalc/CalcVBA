Attribute VB_Name = "TextCompare"
Option Explicit
Option Compare Text

Public Function TextSame(ByVal first As String, ByVal second As String) As Boolean
    TextSame = first = second
End Function
