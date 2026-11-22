Option Explicit

Public Sub Main()
    Debug.Print Probe()
End Sub

Public Function Probe() As String
    Dim files As Object
    Set files = CreateObject("sCrIpTiNg.fIlEsYsTeMoBjEcT")
    Probe = CStr(files.FileExists("c:\dAtA\NoTe.TxT")) & "|" & _
            CStr(files.FolderExists("C:\DATA")) & "|" & _
            CStr(files.FileExists("C:\DATA")) & "|" & _
            CStr(files.FolderExists("C:\DATA\NOTE.TXT")) & "|" & _
            CStr(files.FileExists("C:\DATA\Missing.txt"))
End Function
