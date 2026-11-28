Public Sub Main()
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Debug.Print fso.GetFileName("C:\Data\Missing.tar.GZ") & "|" & fso.GetExtensionName("C:\Data\Missing.tar.GZ") & "|" & fso.GetFileName("C:") & "|" & fso.GetExtensionName("C:\Data.v1\NoExt") & "|" & fso.GetFileName("D:Leaf.log")
End Sub
