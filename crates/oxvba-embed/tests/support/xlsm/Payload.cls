Private Items() As Long
Public Values As Variant
Public Sub Fill()
    ReDim Items(1 To 2, 1 To 2)
    Items(1, 1) = 11
    Items(1, 2) = 12
    Items(2, 1) = 21
    Items(2, 2) = 22
    ReDim Preserve Items(1 To 2, 1 To 3)
End Sub
Public Property Get Matrix() As Variant
    Matrix = Items
End Property
Public Sub PutItem(ByVal row As Long, ByVal column As Long, ByVal value As Long)
    Items(row, column) = value
End Sub
Public Function GetItem(ByVal row As Long, ByVal column As Long) As Long
    GetItem = Items(row, column)
End Function
