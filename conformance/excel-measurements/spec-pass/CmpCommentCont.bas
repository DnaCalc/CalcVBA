Attribute VB_Name = "CmpCommentCont"
Public Function Run() As String
    Dim s As String
    s = "a" ' note _
    s = "b"
    Run = "commentCont=" & s & vbLf
End Function
