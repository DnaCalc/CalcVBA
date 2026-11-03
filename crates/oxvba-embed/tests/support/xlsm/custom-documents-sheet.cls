Private SavedCaption As String
Public Total As Long
Public Link As Object
Private Function Secret() As Long
    Secret = 99
End Function
Public Function Bump(ByRef value As Long) As Long
    value = value + 4
    Bump = value * 2
End Function
Public Property Get Caption() As String
    Caption = SavedCaption
End Property
Public Property Let Caption(ByVal value As String)
    SavedCaption = value
End Property
