Attribute VB_Name = "NativeCom"
Option Explicit
Public Function EarlyDictionarySuite(ByVal dictionary As Scripting.Dictionary) As String
    dictionary.RemoveAll
    dictionary.CompareMode = 1
    dictionary.Add "first", 7
    dictionary.Item("second") = 11
    EarlyDictionarySuite = CStr(dictionary.Count) & ":" & CStr(dictionary.Item("first")) & ":" & CStr(dictionary.Item("second")) & ":" & CStr(dictionary.Exists("FIRST")) & ":" & CStr(TypeOf dictionary Is Scripting.Dictionary) & ":" & CStr(TypeOf dictionary Is Collection)
    On Error Resume Next
    dictionary.CompareMode = 0
    EarlyDictionarySuite = EarlyDictionarySuite & "|" & CStr(Err.Number) & ":" & Err.Description
    Err.Clear
End Function
Public Sub Main()
    Dim observed As String
    observed = EarlyOwnDictionarySuite()
End Sub
Public Function EarlyOwnDictionarySuite() As String
    Dim dictionary As New Scripting.Dictionary
    EarlyOwnDictionarySuite = EarlyDictionarySuite(dictionary)
    Set dictionary = Nothing
    If dictionary.Count <> 0 Then Err.Raise 513
End Function
