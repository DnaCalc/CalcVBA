Option Explicit
Public Function ProbeNativeExports() As String
    Dim first As __LIB__.Calculator, second As __LIB__.Calculator
    Dim counter As __LIB__.Counter, item As __LIB__.PublicItem
    Dim anyObject As Object, stage As String
    Dim sink As EventClient, number As Long, text As String, data As Variant
    Dim dictionary As Object, saved As Object
    Dim scalar As __LIB__.ScalarErrors, bad As __LIB__.BadInitialize, quotient As Double
    Dim code As Long, description As String, errorSource As String, helpFile As String, helpContext As Long
    On Error GoTo Failed
    stage = "create"
    Set first = New __LIB__.Calculator
    Set second = New __LIB__.Calculator
    If first.Value <> 5 Or second.Value <> 5 Then Err.Raise 5, , "Class_Initialize"
    first.Value = 11
    If first.Value <> 11 Or second.Value <> 5 Then Err.Raise 5, , "instance state"
    If first.Add(20, 22) <> 42 Then Err.Raise 5, , "typed method"
    number = 8: text = ChrW(&H3A9)
    stage = "ByRef scalar"
    first.Bump number, text
    If number <> 9 Or text <> (ChrW(&H3A9) & ChrW(&H3BB)) Then Err.Raise 5, , "ByRef writeback: " & CStr(number) & ":" & text
    If Not first.MissingValue() Or first.MissingValue(Empty) Then Err.Raise 5, , "optional Variant"
    number = 2
    first.AliasPair number, number
    If number <> 13 Then Err.Raise 5, , "ByRef alias"
    stage = "matrix"
    data = first.Values()
    If LBound(data, 1) <> 1 Or UBound(data, 1) <> 2 Or LBound(data, 2) <> -1 Or UBound(data, 2) <> 1 Then Err.Raise 5, , "array bounds"
    If data(1, -1) <> 7 Or data(1, 0) <> ChrW(&H3BB) Or data(1, 1) <> True Or data(2, 1) <> 9 Then Err.Raise 5, , "array contents"
    If CStr(data(2, -1)) <> CStr(CVErr(2042)) Or Not IsEmpty(data(2, 0)) Then Err.Raise 5, , "array error/empty"
    If first.Echo((ChrW(&H3BB) & ChrW(&H3A9))) <> (ChrW(&H3BB) & ChrW(&H3A9)) Then Err.Raise 5, , "Variant string"
    If first.Echo(True) <> True Then Err.Raise 5, , "Variant Boolean"
    If CStr(first.Echo(CVErr(2042))) <> CStr(CVErr(2042)) Then Err.Raise 5, , "Variant error"
    stage = "public item"
    Set item = first.MakeItem()
    If item.Ping() <> 42 Or first.PrivateValue() <> 77 Then Err.Raise 5, , "class visibility"
    stage = "ByRef object"
    first.ReplaceItem anyObject
    Set item = anyObject
    stage = "ByRef object"
    first.ReplaceItem anyObject
    If anyObject Is item Then Err.Raise 5, , "ByRef object identity"
    Set item = anyObject
    If item.Ping() <> 42 Then Err.Raise 5, , "ByRef object result"
    Set anyObject = Nothing
    stage = "foreign objects"
    Set dictionary = CreateObject("Scripting.Dictionary")
    dictionary.Add "first", 1
    If first.ObjectCount(dictionary) <> 11 Or first.ObjectCount(dictionary, Empty) <> 1 Then Err.Raise 5, , "foreign optional Object"
    Set saved = dictionary
    first.KeepObject dictionary
    If Not dictionary Is saved Then Err.Raise 5, , "foreign ByRef identity"
    If dictionary.Count <> 2 Or dictionary("kept") <> 42 Then Err.Raise 5, , "foreign member call"
    Set anyObject = first.EchoObject(dictionary)
    If Not anyObject Is dictionary Then Err.Raise 5, , "foreign Object return"
    Set first.Foreign = dictionary
    Set anyObject = first.Foreign
    If Not anyObject Is dictionary Then Err.Raise 5, , "foreign property"
    Set first.Foreign = Nothing
    If first.ObjectName(Application) <> Application.Name Then Err.Raise 5, , "foreign Excel Application"
    first.ReplaceItem dictionary
    If dictionary Is saved Or dictionary.Ping() <> 42 Or saved.Count <> 2 Then Err.Raise 5, , "foreign ByRef replacement"
    Set dictionary = Nothing: Set saved = Nothing: Set anyObject = Nothing
    stage = "dual"
    Set counter = New __LIB__.Counter
    counter.Value = 123
    If counter.Value <> 123 Then Err.Raise 5, , "dual property"
    stage = "global"
    If NativeGlobalAnswer() <> 42 Then Err.Raise 5, , "GlobalMultiUse"
    If NextGlobalValue() <> 1 Then Err.Raise 5, , "global initial state"
    If NextGlobalValue() <> 2 Then Err.Raise 5, , "global retained state"
    stage = "events"
    Set sink = New EventClient
    Set sink.Source = first
    first.Fire 99
    If sink.Seen <> 100 Or sink.ReadBack <> 12 Or first.Value <> 13 Then Err.Raise 5, , "WithEvents re-entry"
    Set sink.Source = Nothing
    Set sink = Nothing
    stage = "factory error"
    On Error Resume Next
    Set bad = New __LIB__.BadInitialize
    code = Err.Number: description = Err.Description: errorSource = Err.Source
    helpFile = Err.HelpFile: helpContext = Err.HelpContext
    Err.Clear
    On Error GoTo Failed
    If code <> 457 Or description <> "initialization failed" Or errorSource <> "BadInitialize" Or Right(helpFile, 8) <> "init.chm" Or helpContext <> 73 Then Err.Raise 5, , "factory details: " & CStr(code) & ":" & description & ":" & errorSource
    stage = "dual error"
    Set scalar = New __LIB__.ScalarErrors
    On Error Resume Next
    quotient = scalar.Divide(1, 0)
    code = Err.Number: description = Err.Description
    Err.Clear
    On Error GoTo Failed
    If code <> 11 Or description <> "Division by zero" Then Err.Raise 5, , "dual details: " & CStr(code) & ":" & description
    On Error Resume Next
    quotient = scalar.Divide(457, 1)
    code = Err.Number: description = Err.Description: errorSource = Err.Source
    helpFile = Err.HelpFile: helpContext = Err.HelpContext
    Err.Clear
    On Error GoTo Failed
    If code <> 457 Or description <> "dual failure" Or errorSource <> "ScalarErrors" Or Right(helpFile, 8) <> "dual.chm" Or helpContext <> 81 Then Err.Raise 5, , "dual raised details"
    If scalar.Divide(84, 2) <> 42 Then Err.Raise 5, , "dual recovery"
    Set scalar = Nothing
    On Error Resume Next
    stage = "fault"
    first.MutateThenRaise number, text
    code = Err.Number: description = Err.Description
    errorSource = Err.Source: helpFile = Err.HelpFile: helpContext = Err.HelpContext
    Err.Clear
    On Error GoTo Failed
    If code <> 457 Or description <> "private COM test" Then Err.Raise 5, , "COM error: " & CStr(code) & ":" & description
    If number <> 67 Or text <> ("fault" & ChrW(&H3BB)) Then Err.Raise 5, , "ByRef before fault"
    If errorSource <> "Calculator" Or Right(helpFile, 8) <> "test.chm" Or helpContext <> 91 Then Err.Raise 5, , "COM error details"
    Set item = Nothing
    Set counter = Nothing
    Set second = Nothing
    Set first = Nothing
    ProbeNativeExports = "instances|ByRef|Variant|matrix|visibility|dual|global|events|error=457"
    Exit Function
Failed:
    ProbeNativeExports = "FAIL|" & stage & "|" & CStr(Err.Number) & "|" & Err.Description & "|" & Err.Source
End Function
