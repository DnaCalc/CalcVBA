Sub Main()
    Dim output As Object
    Dim inputFile As Object
    Dim bytes As Variant
    Set output = CreateObject("ADODB.Stream")
    output.Open
    output.Charset = "utf-8"
    output.WriteText "caf" & ChrW(233)
    output.SaveToFile "C:\Data\Utf8.txt", 2
    Debug.Print CStr(output.Position) & "|" & CStr(output.Size)
    output.Close
    Set inputFile = CreateObject("ADODB.Stream")
    inputFile.Open
    inputFile.Type = 1
    inputFile.LoadFromFile "c:\data\utf8.TXT"
    bytes = inputFile.Read(3)
    Debug.Print CStr(bytes(0)) & "|" & CStr(bytes(1)) & "|" & CStr(bytes(2))
    inputFile.Position = 0
    inputFile.Type = 2
    inputFile.Charset = "utf-8"
    Debug.Print CStr(inputFile.ReadText = "caf" & ChrW(233))
    inputFile.Close
End Sub
