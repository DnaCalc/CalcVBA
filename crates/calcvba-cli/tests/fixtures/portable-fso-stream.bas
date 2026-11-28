Sub Main()
    Dim fso As Object
    Dim output As Object
    Dim inputFile As Object
    Dim data As String
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set output = fso.CreateTextFile("C:\Data\Made.txt", True, False)
    output.Write "first"
    output.WriteLine " line"
    output.Close
    Set inputFile = fso.OpenTextFile("c:\data\MADE.txt")
    data = inputFile.ReadAll
    inputFile.Close
    Debug.Print CStr(FileLen("C:\data\Made.txt")) & "|" & CStr(data = "first line" & vbCrLf)
End Sub
