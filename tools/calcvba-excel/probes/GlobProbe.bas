Attribute VB_Name = "GlobProbe"
Option Explicit

Private Sub Setup()
    ThisWorkbook.Worksheets(1).Range("A1").Value = 42
    ThisWorkbook.Names.Add Name:="MyDog", RefersTo:="=" & ThisWorkbook.Worksheets(1).Name & "!$A$1"
    ThisWorkbook.Names.Add Name:="MyConst", RefersTo:="=7*6"
End Sub

Private Function Show(ByVal v As Variant) As String
    If IsObject(v) Then
        Show = "obj:" & TypeName(v)
    ElseIf IsError(v) Then
        Show = "err:" & CStr(v)
    ElseIf IsArray(v) Then
        Show = "array:" & TypeName(v)
    Else
        Show = TypeName(v) & ":" & CStr(v)
    End If
End Function

Function RunAll() As String
    Dim out As String
    Setup
    out = out & "TW_IS_APP_TW=" & (ThisWorkbook Is Application.ThisWorkbook) & vbLf
    out = out & "TW_IS_EXCEL_TW=" & (ThisWorkbook Is Excel.ThisWorkbook) & vbLf
    out = out & "TYPENAME_TW=" & TypeName(ThisWorkbook) & vbLf
    out = out & "TYPENAME_APP=" & TypeName(Application) & vbLf
    out = out & "APP_IS_APP_APP=" & (Application Is Application.Application) & vbLf
    out = out & "APP_IS_EXCEL_APP=" & (Application Is Excel.Application) & vbLf
    out = out & "BRACKET_MYDOG=" & Show([MyDog]) & " VALUE=" & [MyDog].Value & " ADDR=" & [MyDog].Address & vbLf
    out = out & "BRACKET_A1=" & Show([A1]) & vbLf
    out = out & "BRACKET_EXPR=" & Show([A1+1]) & vbLf
    out = out & "BRACKET_CONST=" & Show([MyConst]) & vbLf
    out = out & "BRACKET_ARRAY=" & Show([{1,2,3}]) & vbLf
    out = out & "BRACKET_UNKNOWN=" & Show([NoSuchName]) & vbLf
    out = out & "EVALUATE=" & Show(Evaluate("MyDog")) & vbLf
    out = out & "EXCEL_EVALUATE=" & Show(Excel.Evaluate("A1+1")) & vbLf
    out = out & "APP_SUM=" & Show(Application.Sum(1, 2)) & vbLf
    out = out & "BARE_RANGE=" & Show(Range("A1").Value) & vbLf
    out = out & "EXCEL_RANGE=" & Show(Excel.Range("A1").Value) & vbLf
    out = out & "SHEET_BRACKET=" & Show(ActiveSheet.[A1]) & vbLf
    out = out & "SHEET_BRACKET_EXPR=" & Show(ThisWorkbook.Worksheets(1).[A1*2]) & vbLf
    out = out & "APP_BRACKET=" & AppBracket() & vbLf
    out = out & "APP_MYDOG=" & AppMyDog() & vbLf
    out = out & "APP_MYDOG_OBJ=" & AppMyDogObj() & vbLf
    out = out & "BRACKET_LOCAL=" & BracketLocal() & vbLf
    [A1] = 5
    out = out & "BRACKET_ASSIGN=" & Show(Range("A1").Value) & vbLf
    [MyDog] = 6
    out = out & "BRACKET_NAME_ASSIGN=" & Show(Range("A1").Value) & vbLf
    RunAll = out
End Function

Private Function AppBracket() As String
    On Error GoTo fail
    AppBracket = Show(Application.[A1])
    Exit Function
fail:
    AppBracket = "raised " & Err.Number & " " & Err.Description
End Function

Private Function AppMyDog() As String
    On Error GoTo fail
    AppMyDog = Show(Application.MyDog)
    Exit Function
fail:
    AppMyDog = "raised " & Err.Number & " " & Err.Description
End Function

Private Function AppMyDogObj() As String
    Dim o As Object
    On Error GoTo fail
    Set o = Application
    AppMyDogObj = Show(o.MyDog)
    Exit Function
fail:
    AppMyDogObj = "raised " & Err.Number & " " & Err.Description
End Function

Private Function BracketLocal() As String
    Dim abc As Long
    abc = 9
    On Error GoTo fail
    BracketLocal = Show([abc])
    Exit Function
fail:
    BracketLocal = "raised " & Err.Number & " " & Err.Description
End Function
