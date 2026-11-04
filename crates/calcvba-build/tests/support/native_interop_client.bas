Option Explicit
Private Declare PtrSafe Function NRaiseDynamic Lib "__DLL__" Alias "RaiseDynamic" (ByVal number As Long, ByVal text As LongPtr) As Long
Private Declare PtrSafe Function NHelp Lib "__DLL__" Alias "CalcVbaGetLastErrorHelp" (ByRef helpFile As LongPtr, ByRef helpContext As Long) As Long
Private Declare PtrSafe Function NGrowMatrix Lib "__DLL__" Alias "GrowMatrix" (ByRef values() As Double) As Double
Private Declare PtrSafe Function NShutdown Lib "__DLL__" Alias "CalcVbaShutdown" () As Long
Private Declare PtrSafe Function NStoreString Lib "__DLL__" Alias "StoreString" (ByVal value As LongPtr) As LongPtr
Private Declare PtrSafe Function NHoldObject Lib "__DLL__" Alias "HoldObject" (ByVal value As LongPtr) As Boolean
Private Declare PtrSafe Function NVariantTruth Lib "__DLL__" Alias "VariantTruth" (ByVal value As LongPtr) As Long
Private Declare PtrSafe Function NTextSame Lib "__DLL__" Alias "TextSame" (ByVal first As LongPtr, ByVal second As LongPtr) As Boolean
Private Declare PtrSafe Function NSmallRecord Lib "__DLL__" Alias "SmallRecord" (ByVal value As Long) As Long
Private Declare PtrSafe Function NValueSuite Lib "__DLL__" Alias "ValueSuite" () As LongPtr
Private Declare PtrSafe Function NRaiseCustom Lib "__DLL__" Alias "RaiseCustom" (ByVal number As Long) As Long
Private Declare PtrSafe Function NReadLastError Lib "__DLL__" Alias "ReadLastDllError" () As Long
Private Declare PtrSafe Function NSameString Lib "__DLL__" Alias "SameString" (ByVal first As LongPtr, ByVal second As LongPtr) As Boolean
Private Declare PtrSafe Function NJoinArray Lib "__DLL__" Alias "JoinArray" (ByRef values() As String) As LongPtr
Private Declare PtrSafe Sub SetWinError Lib "kernel32" Alias "SetLastError" (ByVal number As Long)
Private Declare PtrSafe Function NBump Lib "__DLL__" Alias "Bump" (ByRef a As Long, ByRef b As Long) As Long
Private Declare PtrSafe Function NBumpDouble Lib "__DLL__" Alias "BumpDouble" (ByRef value As Double) As Double
Private Declare PtrSafe Function NFlip Lib "__DLL__" Alias "Flip" (ByRef value As Boolean) As Boolean
Private Declare PtrSafe Function NNext Lib "__DLL__" Alias "NextValue" (ByVal value As Long) As Long
Private Declare PtrSafe Function NStatic Lib "__DLL__" Alias "StaticValue" (ByVal value As Long) As Long
Private Declare PtrSafe Sub NMutate Lib "__DLL__" Alias "Mutate" (ByRef value As Long)
Private Declare PtrSafe Function NNested Lib "__DLL__" Alias "Nested" (ByRef value As Long) As Long
Private Declare PtrSafe Function NDivide Lib "__DLL__" Alias "Divide" (ByVal a As Double, ByVal b As Double) As Double
Private Declare PtrSafe Function NDivideHR Lib "__DLL__" Alias "DivideHResult" (ByVal a As Double, ByVal b As Double, ByRef value As Double) As Long
Private Declare PtrSafe Function NError Lib "__DLL__" Alias "CalcVbaGetLastError" (ByRef source As LongPtr, ByRef description As LongPtr) As Long
Private Declare PtrSafe Function NEchoString Lib "__DLL__" Alias "EchoString" (ByVal value As LongPtr) As LongPtr
Private Declare PtrSafe Function NMutateString Lib "__DLL__" Alias "MutateString" (ByVal value As LongPtr) As Long
Private Declare PtrSafe Function NStringLength Lib "__DLL__" Alias "StringLength" (ByVal value As LongPtr) As Long
Private Declare PtrSafe Function NLengthAnsi Lib "__DLL__" Alias "LengthAnsi" (ByVal value As String) As Long
Private Declare PtrSafe Function NLengthUtf16 Lib "__DLL__" Alias "LengthUtf16" (ByVal value As LongPtr) As Long
Private Declare PtrSafe Sub NAddVariant Lib "__DLL__" Alias "AddVariant" (ByRef result As Variant, ByRef value As Variant, ByVal increment As Long)
Private Declare PtrSafe Function NVariantText Lib "__DLL__" Alias "VariantText" (ByVal value As LongPtr) As LongPtr
Private Declare PtrSafe Function NEchoObject Lib "__DLL__" Alias "EchoObject" (ByVal value As LongPtr) As LongPtr
Private Declare PtrSafe Function NSameObject Lib "__DLL__" Alias "SameObject" (ByVal a As LongPtr, ByVal b As LongPtr) As Boolean
Private Declare PtrSafe Sub NClearObject Lib "__DLL__" Alias "ClearObject" (ByRef value As LongPtr)
Private Declare PtrSafe Function NChangeRecord Lib "__DLL__" Alias "ChangeRecord" (ByVal value As LongPtr) As Double
Private Declare PtrSafe Sub NEchoRecord Lib "__DLL__" Alias "EchoRecord" (ByVal result As LongPtr, ByVal value As LongPtr)
Private Declare PtrSafe Function NSumArray Lib "__DLL__" Alias "SumArray" (ByRef values() As Long) As Long
Private Declare PtrSafe Function NChangeArray Lib "__DLL__" Alias "ChangeArray" (ByRef values() As Long) As Long
Private Declare PtrSafe Sub NAllocateArray Lib "__DLL__" Alias "AllocateArray" (ByRef values() As Long)
Private Declare PtrSafe Function NReadMatrix Lib "__DLL__" Alias "ReadMatrix" (ByRef values() As Double, ByVal row As Long, ByVal column As Long) As Double
Private Declare PtrSafe Function BStrLength Lib "oleaut32" Alias "SysStringLen" (ByVal value As LongPtr) As Long
Private Declare PtrSafe Sub FreeBStr Lib "oleaut32" Alias "SysFreeString" (ByVal value As LongPtr)
Private Declare PtrSafe Sub CopyBytes Lib "kernel32" Alias "RtlMoveMemory" (ByVal destination As LongPtr, ByVal source As LongPtr, ByVal count As LongPtr)

Private Type NativeRecord
    tag As Byte
    code As Integer
    count As Long
    amount As Double
    enabled As Boolean
    text As String
End Type

Private Function TakeString(ByVal pointer As LongPtr) As String
    Dim length As Long
    length = BStrLength(pointer)
    TakeString = String$(length, ChrW$(0))
    If length Then CopyBytes StrPtr(TakeString), pointer, length * 2&
    FreeBStr pointer
End Function

Private Sub AppendValue(ByRef result As String, ByVal value As Variant)
    If Len(result) Then result = result & "|"
    result = result & CStr(value)
End Sub

Public Function ProbeNativeExports() As String
    Dim result As String, a As Long, b As Long, d As Double, flag As Boolean
    Dim source As LongPtr, description As LongPtr, text As String, returned As String, pointer As LongPtr
    Dim value As Variant, output As Variant, object As Object
    Dim record As NativeRecord, copied As NativeRecord
    Dim values() As Long, matrix() As Double, strings() As String, other As String
    a = 5&: b = 7&
    AppendValue result, NBump(a, b): AppendValue result, a: AppendValue result, b
    a = 5&: AppendValue result, NBump(a, a): AppendValue result, a
    d = 1.25#: AppendValue result, NBumpDouble(d): AppendValue result, d
    flag = True: AppendValue result, CLng(NFlip(flag)): AppendValue result, CLng(NFlip(flag))
    AppendValue result, NNext(3&): AppendValue result, NNext(7&)
    AppendValue result, NStatic(4&): AppendValue result, NStatic(5&)
    a = 5&: NMutate a: AppendValue result, a
    a = 4&: AppendValue result, NNested(a): AppendValue result, a
    AppendValue result, NDivide(9#, 2#): AppendValue result, NDivide(9#, 0#)
    AppendValue result, NError(source, description)
    returned = TakeString(source): text = TakeString(description)
    If returned <> "NativeInterop" Or text <> "Division by zero" Then Err.Raise 513
    d = 99#: AppendValue result, NDivideHR(9#, 0#, d): AppendValue result, d
    AppendValue result, NDivideHR(9#, 2#, d): AppendValue result, d
    text = "A" & ChrW$(0) & ChrW$(&H3A9)
    returned = TakeString(NEchoString(StrPtr(text)))
    AppendValue result, Len(returned)
    If returned <> "<" & text & ">" Then Err.Raise 514
    AppendValue result, NStringLength(StrPtr(text))
    AppendValue result, NMutateString(VarPtr(text))
    If text <> "A" & ChrW$(0) & ChrW$(&H3A9) & "!" Then Err.Raise 515
    AppendValue result, NLengthAnsi("Hello")
    text = ChrW$(&H6F22) & ChrW$(&H5B57) & "x"
    AppendValue result, NLengthUtf16(StrPtr(text))
    value = 37&: NAddVariant output, value, 5&
    AppendValue result, value: AppendValue result, output
    returned = TakeString(NVariantText(VarPtr(value)))
    If returned <> "42" Then Err.Raise 516
    AppendValue result, Len(returned)
    Set object = CreateObject("Scripting.Dictionary")
    pointer = NEchoObject(ObjPtr(object))
    AppendValue result, CLng(NSameObject(ObjPtr(object), pointer))
    AppendValue result, CLng(NSameObject(ObjPtr(object), 0))
    NClearObject pointer
    If pointer <> 0 Then Err.Raise 517
    Set object = Nothing
    record.tag = 127: record.code = -123: record.count = 5&: record.amount = 1.25#: record.enabled = True: record.text = "ab"
    AppendValue result, NChangeRecord(VarPtr(record))
    AppendValue result, record.count: AppendValue result, record.amount: AppendValue result, CLng(record.enabled): AppendValue result, Len(record.text)
    If record.tag <> 127 Or record.code <> -123 Or record.text <> "ab!" Then Err.Raise 518
    NEchoRecord VarPtr(copied), VarPtr(record)
    If copied.text <> record.text Then Err.Raise 519
    AppendValue result, copied.count
    NAllocateArray values
    AppendValue result, NSumArray(values): AppendValue result, NChangeArray(values)
    If LBound(values) <> -2 Or UBound(values) <> 1 Then Err.Raise 520
    ReDim matrix(-1 To 0, 3 To 5)
    matrix(0, 5) = 3.5#
    AppendValue result, NReadMatrix(matrix, 0&, 5&): AppendValue result, matrix(0, 5)
    AppendValue result, NGrowMatrix(matrix): AppendValue result, matrix(0, 6): AppendValue result, UBound(matrix, 2)
    AppendValue result, NRaiseCustom(513&)
    AppendValue result, NError(source, description)
    returned = TakeString(source): text = TakeString(description)
    If returned <> "NativeCase" Or text <> "custom failure" Then Err.Raise 521
    SetWinError 12345&: AppendValue result, NReadLastError()
    text = "A" & ChrW$(0) & ChrW$(&H3A9): other = "A" & ChrW$(0) & ChrW$(&H3AA)
    AppendValue result, CLng(NSameString(StrPtr(text), StrPtr(text)))
    AppendValue result, CLng(NSameString(StrPtr(text), StrPtr(other)))
    ReDim strings(-2 To -1)
    strings(-2) = "left": strings(-1) = "right"
    returned = TakeString(NJoinArray(strings)): AppendValue result, Len(returned)
    AppendValue result, Len(strings(-1))
    If returned <> "left|right" Or strings(-1) <> "right!" Then Err.Raise 522
    value = "37": NAddVariant output, value, 5&
    AppendValue result, VarType(value): AppendValue result, VarType(output): AppendValue result, output
    value = True: returned = TakeString(NVariantText(VarPtr(value)))
    If returned <> "True" Then Err.Raise 523
    AppendValue result, Len(returned)
    value = Null: NAddVariant output, value, 5&
    AppendValue result, VarType(value): AppendValue result, VarType(output)
    pointer = NVariantText(VarPtr(value))
    AppendValue result, CLng(pointer <> 0)
    AppendValue result, NError(source, description)
    returned = TakeString(source): text = TakeString(description)
    text = ChrW$(&HD800): returned = TakeString(NEchoString(StrPtr(text)))
    AppendValue result, Len(returned): AppendValue result, AscW(Mid$(returned, 2, 1))
    AppendValue result, NSmallRecord(37&)
    returned = TakeString(NValueSuite())
    If returned <> "19:6:13:24:12:3.5:1.75:False:<omega>:omega!:6:42:42:8.75:ab!:49:-2:3.5:3.75:3.75:6:7.5:left|right:right!:True:9:9:7" Then Err.Raise 524
    AppendValue result, Len(returned)
    value = Empty: AppendValue result, NVariantTruth(VarPtr(value))
    value = Null: AppendValue result, NVariantTruth(VarPtr(value))
    value = 0&: AppendValue result, NVariantTruth(VarPtr(value))
    value = 1&: AppendValue result, NVariantTruth(VarPtr(value))
    text = ChrW$(&HC5): other = ChrW$(&HE5)
    AppendValue result, CLng(NTextSame(StrPtr(text), StrPtr(other)))
    text = "a": returned = TakeString(NStoreString(StrPtr(text))): AppendValue result, Len(returned)
    If returned <> "a" Then Err.Raise 525
    text = "b": returned = TakeString(NStoreString(StrPtr(text))): AppendValue result, Len(returned)
    If returned <> "ab" Then Err.Raise 526
    Set object = CreateObject("Scripting.Dictionary")
    AppendValue result, CLng(NHoldObject(ObjPtr(object)))
    Set object = Nothing
    Dim helpFile As LongPtr, helpContext As Long
    text = "dynamic"
    AppendValue result, NRaiseDynamic(513&, StrPtr(text))
    AppendValue result, NError(source, description)
    returned = TakeString(source): text = TakeString(description)
    If returned <> "Native.dynamic" Or text <> "dynamic failure" Then Err.Raise 527
    AppendValue result, Len(returned): AppendValue result, Len(text)
    AppendValue result, NHelp(helpFile, helpContext)
    returned = TakeString(helpFile)
    If returned <> "dynamic.chm" Or helpContext <> 520& Then Err.Raise 528
    AppendValue result, Len(returned): AppendValue result, helpContext
    If NNext(0&) <> 10& Or NHelp(helpFile, helpContext) <> 0& Then Err.Raise 529
    returned = TakeString(helpFile)
    If returned <> "" Or helpContext <> 0& Then Err.Raise 530
    AppendValue result, NShutdown()
    AppendValue result, NNext(0&)
    AppendValue result, NError(source, description)
    returned = TakeString(source): text = TakeString(description)
    ProbeNativeExports = result
End Function

Public Sub Main()
    Dim result As String
    result = ProbeNativeExports()
End Sub
