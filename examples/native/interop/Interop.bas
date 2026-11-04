Attribute VB_Name = "Interop"
Option Explicit
Private counter As Long
Private cached As String
Private heldObject As Object
Public Type NativeRecord
    tag As Byte
    code As Integer
    count As Long
    amount As Double
    enabled As Boolean
    text As String
End Type

Public Type SmallNativeRecord
    value As Long
End Type

Public Function SmallRecord(ByVal value As Long) As SmallNativeRecord
    SmallRecord.value = value
End Function

Public Function StoreString(ByVal value As String) As String
    cached = cached & value
    StoreString = cached
End Function

Public Function HoldObject(ByVal value As Object) As Boolean
    Set heldObject = value
    HoldObject = Not (heldObject Is Nothing)
End Function

Public Function Bump(ByRef first As Long, ByRef second As Long) As Long
    first = first + 1&
    second = second + first
    Bump = first + second
End Function

Public Function BumpDouble(ByRef value As Double) As Double
    value = value + 0.5#
    BumpDouble = value * 2#
End Function

Public Function Flip(ByRef value As Boolean) As Boolean
    value = Not value
    Flip = value
End Function

Public Function NextValue(ByVal stepValue As Long) As Long
    counter = counter + stepValue
    NextValue = counter
End Function

Public Function StaticValue(ByVal stepValue As Long) As Long
    Static total As Long
    total = total + stepValue
    StaticValue = total
End Function

Public Function Divide(ByVal numerator As Double, ByVal denominator As Double) As Double
    Divide = numerator / denominator
End Function

Public Sub Mutate(ByRef value As Long)
    value = value + 2&
End Sub

Public Function Nested(ByRef value As Long) As Long
    Mutate value
    Nested = Bump(value, value)
End Function

Public Function EchoString(ByVal value As String) As String
    EchoString = "<" & value & ">"
End Function

Public Function MutateString(ByRef value As String) As Long
    value = value & "!"
    MutateString = Len(value)
End Function

Public Function StringLength(ByVal value As String) As Long
    StringLength = Len(value)
End Function

Public Function AddVariant(ByRef value As Variant, ByVal increment As Long) As Variant
    value = value + increment
    AddVariant = value
End Function

Public Function VariantText(ByVal value As Variant) As String
    VariantText = CStr(value)
End Function

Public Function VariantTruth(ByVal value As Variant) As Long
    If value Then
        VariantTruth = 7&
    Else
        VariantTruth = 9&
    End If
End Function

Public Function EchoObject(ByVal value As Object) As Object
    Set EchoObject = value
End Function

Public Function SameObject(ByVal first As Object, ByVal second As Object) As Boolean
    SameObject = first Is second
End Function

Public Sub ClearObject(ByRef value As Object)
    Set value = Nothing
End Sub

Public Function ChangeRecord(ByRef value As NativeRecord) As Double
    value.count = value.count + 2&
    value.amount = value.amount + 0.5#
    value.enabled = Not value.enabled
    value.text = value.text & "!"
    ChangeRecord = value.amount + value.count
End Function

Public Function EchoRecord(ByRef value As NativeRecord) As NativeRecord
    EchoRecord = value
End Function

Public Function SumArray(ByRef values() As Long) As Long
    Dim i As Long
    For i = LBound(values) To UBound(values)
        SumArray = SumArray + values(i)
    Next
End Function

Public Function ChangeArray(ByRef values() As Long) As Long
    values(LBound(values)) = values(LBound(values)) + 7&
    ChangeArray = SumArray(values)
End Function

Public Function ReadMatrix(ByRef values() As Double, ByVal row As Long, ByVal column As Long) As Double
    ReadMatrix = values(row,column)
    values(row,column) = ReadMatrix + 0.25#
End Function

Public Function GrowMatrix(ByRef values() As Double) As Double
    ReDim Preserve values(-1 To 0, 3 To 6)
    GrowMatrix = values(0, 5)
    values(0, 6) = 7.5#
End Function

Public Sub AllocateArray(ByRef values() As Long)
    ReDim values(-2 To 1)
    values(-2) = 11&
    values(1) = 31&
End Sub

Public Function RaiseCustom(ByVal number As Long) As Long
    Err.Raise number, "NativeCase", "custom failure"
End Function

Public Function ReadLastDllError() As Long
    ReadLastDllError = Err.LastDllError
End Function

Public Function SameString(ByVal first As String, ByVal second As String) As Boolean
    SameString = first = second
End Function

Public Function JoinArray(ByRef values() As String) As String
    JoinArray = values(LBound(values)) & "|" & values(UBound(values))
    values(UBound(values)) = values(UBound(values)) & "!"
End Function

' These exact bodies are also executed directly by the VM3/JIT conformance test.
Public Function ValueSuite() As String
    Dim a As Long, b As Long, d As Double, flag As Boolean, text As String
    Dim value As Variant, output As Variant, record As NativeRecord, copied As NativeRecord
    Dim values() As Long, matrix() As Double, strings() As String, object As Object
    Dim result As String
    a = 5&: b = 7&
    result = CStr(Bump(a, b)) & ":" & CStr(a) & ":" & CStr(b)
    a = 5&: result = result & ":" & CStr(Bump(a, a)) & ":" & CStr(a)
    d = 1.25#: result = result & ":" & CStr(BumpDouble(d)) & ":" & CStr(d)
    flag = True: result = result & ":" & CStr(Flip(flag))
    text = "omega": result = result & ":" & EchoString(text)
    a = MutateString(text): result = result & ":" & text & ":" & CStr(a)
    value = "37": output = AddVariant(value, 5&)
    result = result & ":" & VariantText(value) & ":" & VariantText(output)
    record.count = 5&: record.amount = 1.25#: record.enabled = True: record.text = "ab"
    d = ChangeRecord(record): copied = EchoRecord(record)
    result = result & ":" & CStr(d) & ":" & copied.text
    AllocateArray values
    result = result & ":" & CStr(ChangeArray(values)) & ":" & CStr(LBound(values))
    ReDim matrix(-1 To 0, 3 To 5): matrix(0, 5) = 3.5#
    d = ReadMatrix(matrix, 0&, 5&): result = result & ":" & CStr(d) & ":" & CStr(matrix(0, 5))
    d = GrowMatrix(matrix): result = result & ":" & CStr(d) & ":" & CStr(UBound(matrix, 2)) & ":" & CStr(matrix(0, 6))
    ReDim strings(-2 To -1): strings(-2) = "left": strings(-1) = "right"
    result = result & ":" & JoinArray(strings) & ":" & strings(-1)
    Set object = EchoObject(Nothing)
    result = result & ":" & CStr(SameObject(object, Nothing))
    value = Null: result = result & ":" & CStr(VariantTruth(value))
    value = 0&: result = result & ":" & CStr(VariantTruth(value))
    value = 1&: result = result & ":" & CStr(VariantTruth(value))
    ValueSuite = result
End Function

Public Function RaiseDynamic(ByVal number As Long, ByVal text As String) As Long
    Dim source As Variant, description As Variant
    source = "Native." & text
    description = text & " failure"
    Err.Raise number, source, description, text & ".chm", number + 7&
End Function
