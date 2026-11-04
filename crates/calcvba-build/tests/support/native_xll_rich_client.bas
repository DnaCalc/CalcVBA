Option Explicit
Private Declare PtrSafe Function NEcho Lib "__DLL__" Alias "NativeEcho" (ByVal value As LongPtr) As LongPtr
Private Declare PtrSafe Function NVariantEcho Lib "__DLL__" Alias "NativeVariantEcho" (ByRef value As Variant) As LongPtr
Private Declare PtrSafe Sub NFree Lib "__DLL__" Alias "xlAutoFree12" (ByVal value As LongPtr)
Private Declare PtrSafe Function NOwned Lib "__DLL__" Alias "OwnedBuffers" () As Long
Private Declare PtrSafe Function NCount Lib "__DLL__" Alias "OutstandingResults" () As Long
Private Declare PtrSafe Function NShutdown Lib "__DLL__" Alias "CalcVbaShutdown" () As Long
Private Declare PtrSafe Function Allocate Lib "ole32" Alias "CoTaskMemAlloc" (ByVal bytes As LongPtr) As LongPtr
Private Declare PtrSafe Sub Release Lib "ole32" Alias "CoTaskMemFree" (ByVal value As LongPtr)
Private Declare PtrSafe Sub MoveBytes Lib "kernel32" Alias "RtlMoveMemory" (ByVal destination As LongPtr, ByVal source As LongPtr, ByVal bytes As LongPtr)
Private Type Oper
    payload0 As LongLong
    payload1 As LongLong
    payload2 As LongLong
    kind As Long
    padding As Long
End Type
Private Type MatrixSize
    rows As Long
    columns As Long
End Type
Private Type Rect
    firstRow As Long
    lastRow As Long
    firstColumn As Long
    lastColumn As Long
End Type
Private Function CallEcho(ByRef inputValue As Oper, ByRef result As Oper, ByVal expectedKind As Long) As LongPtr
    CallEcho = NEcho(VarPtr(inputValue))
    If CallEcho = 0^ Then Err.Raise 513&, "RichClient", "null native result"
    MoveBytes VarPtr(result), CallEcho, 32^
    If result.kind <> (expectedKind Or 16384&) Then Err.Raise 514&, "RichClient", "result type: input=" & CStr(inputValue.kind) & " expected=" & CStr(expectedKind) & " observed=" & CStr(result.kind)
End Function
Private Sub FreeEcho(ByVal pointer As LongPtr)
    NFree pointer
    If NOwned() <> 0& Or NCount() <> 0& Then Err.Raise 515&, "RichClient", "unreleased return buffer"
End Sub
Private Function CountedString(ByVal value As String) As LongPtr
    Dim length As Long
    length = Len(value)
    CountedString = Allocate(CLngPtr(length) * 2^ + 2^)
    MoveBytes CountedString, VarPtr(length), 2^
    MoveBytes CountedString + 2^, StrPtr(value), CLngPtr(length) * 2^
End Function
Private Function TextAt(ByVal pointer As LongPtr) As String
    Dim length As Long
    MoveBytes VarPtr(length), pointer, 2^
    TextAt = String$(length, ChrW$(0))
    If length <> 0& Then MoveBytes StrPtr(TextAt), pointer + 2^, CLngPtr(length) * 2^
End Function
Public Function ProbeNativeExports() As String
    Dim inputValue As Oper, result As Oper, child As Oper, pointer As LongPtr
    Dim number As Double, observed As Double, text As String, encoded As LongPtr
    Dim code As Long, codes As Variant, i As Long, cells As LongPtr, size As MatrixSize
    Dim variants As Variant, value As Variant
    variants = Array(CByte(7), CInt(-123), CLng(456), CLngLng(9007199254740993^), CSng(1.25), CDbl(2.75), CCur(12.3456), CDate(45123.25), CDec(1.125))
    For i = 0& To 8&
        value = variants(i)
        pointer = NVariantEcho(value)
        If pointer = 0^ Then Err.Raise 534&
        MoveBytes VarPtr(result), pointer, 32^
        MoveBytes VarPtr(observed), VarPtr(result), 8^
        If result.kind <> 16385& Or observed <> CDbl(value) Then Err.Raise 535&
        FreeEcho pointer
    Next i
    number = 123.5
    MoveBytes VarPtr(inputValue), VarPtr(number), 8^
    inputValue.kind = 1&
    pointer = CallEcho(inputValue, result, 1&)
    MoveBytes VarPtr(observed), VarPtr(result), 8^
    If observed <> number Then Err.Raise 516&
    FreeEcho pointer
    text = "A" & ChrW$(0) & ChrW$(955) & ChrW$(937)
    encoded = CountedString(text)
    inputValue.kind = 2&: inputValue.payload0 = encoded
    pointer = CallEcho(inputValue, result, 2&)
    If TextAt(result.payload0) <> text Or result.payload0 = encoded Then Err.Raise 517&
    FreeEcho pointer
    inputValue.kind = 4&: inputValue.payload0 = -1^
    pointer = CallEcho(inputValue, result, 4&)
    If result.payload0 <> 1^ Then Err.Raise 518&
    FreeEcho pointer
    ' Deliberately leave the inactive half of the union nonzero.
    inputValue.payload0 = &H7FF0000000000000^
    inputValue.kind = 4&: code = 0&
    MoveBytes VarPtr(inputValue), VarPtr(code), 4^
    pointer = CallEcho(inputValue, result, 4&)
    If result.payload0 <> 0^ Then Err.Raise 519&
    FreeEcho pointer
    codes = Array(0&, 7&, 15&, 23&, 29&, 36&, 42&, 43&)
    For i = 0& To 7&
        code = codes(i): inputValue.kind = 16&
        MoveBytes VarPtr(inputValue), VarPtr(code), 4^
        pointer = CallEcho(inputValue, result, 16&)
        If result.payload0 <> CLngLng(code) Then Err.Raise 520&
        FreeEcho pointer
    Next i
    inputValue.kind = 128&: pointer = CallEcho(inputValue, result, 128&): FreeEcho pointer
    inputValue.kind = 256&: pointer = CallEcho(inputValue, result, 256&): FreeEcho pointer
    inputValue.kind = 2048&: code = -3&
    MoveBytes VarPtr(inputValue), VarPtr(code), 4^
    pointer = CallEcho(inputValue, result, 1&)
    MoveBytes VarPtr(observed), VarPtr(result), 8^
    If observed <> -3# Then Err.Raise 521&
    FreeEcho pointer
    cells = Allocate(192^)
    For i = 0& To 5&
        If i = 0& Then
            child.kind = 1&: number = 2.5: MoveBytes VarPtr(child), VarPtr(number), 8^
        ElseIf i = 1& Then
            child.kind = 2&: child.payload0 = encoded
        ElseIf i = 2& Then
            child.kind = 4&: child.payload0 = 1^
        ElseIf i = 3& Then
            child.kind = 16&: child.payload0 = 42^
        ElseIf i = 4& Then
            child.kind = 128&
        Else
            child.kind = 256&
        End If
        MoveBytes cells + CLngPtr(i) * 32^, VarPtr(child), 32^
    Next i
    inputValue.kind = 64&: inputValue.payload0 = cells
    size.rows = 2&: size.columns = 3&
    MoveBytes VarPtr(inputValue) + 8^, VarPtr(size), 8^
    pointer = CallEcho(inputValue, result, 64&)
    MoveBytes VarPtr(size), VarPtr(result) + 8^, 8^
    If size.rows <> 2& Or size.columns <> 3& Then Err.Raise 522&
    For i = 0& To 5&
        MoveBytes VarPtr(child), result.payload0 + CLngPtr(i) * 32^, 32^
        If i = 0& Then
            MoveBytes VarPtr(observed), VarPtr(child), 8^
            If child.kind <> 1& Or observed <> 2.5 Then Err.Raise 523&
        ElseIf i = 1& Then
            If child.kind <> 2& Or TextAt(child.payload0) <> text Then Err.Raise 524&
        ElseIf i = 2& Then
            If child.kind <> 4& Or child.payload0 <> 1^ Then Err.Raise 525&
        ElseIf i = 3& Then
            If child.kind <> 16& Or child.payload0 <> 42^ Then Err.Raise 526&
        ElseIf i = 4& Then
            If child.kind <> 128& Then Err.Raise 527&
        Else
            If child.kind <> 256& Then Err.Raise 528&
        End If
    Next i
    FreeEcho pointer
    Release cells: Release encoded
    Dim rectangle As Rect, copied As Rect
    inputValue.kind = 1024&: inputValue.payload0 = 1^
    rectangle.firstRow = 2&: rectangle.lastRow = 4&: rectangle.firstColumn = 1&: rectangle.lastColumn = 3&
    MoveBytes VarPtr(inputValue) + 4^, VarPtr(rectangle), 16^
    pointer = CallEcho(inputValue, result, 1024&)
    MoveBytes VarPtr(copied), VarPtr(result) + 4^, 16^
    If copied.firstRow <> 2& Or copied.lastRow <> 4& Or copied.firstColumn <> 1& Or copied.lastColumn <> 3& Then Err.Raise 529&
    FreeEcho pointer
    Dim areas As LongPtr
    areas = Allocate(36^): code = 2&
    MoveBytes areas, VarPtr(code), 4^
    rectangle.firstRow = 0&: rectangle.lastRow = 1&: rectangle.firstColumn = 2&: rectangle.lastColumn = 3&
    MoveBytes areas + 4^, VarPtr(rectangle), 16^
    rectangle.firstRow = 5&: rectangle.lastRow = 6&: rectangle.firstColumn = 7&: rectangle.lastColumn = 8&
    MoveBytes areas + 20^, VarPtr(rectangle), 16^
    inputValue.kind = 8&: inputValue.payload0 = areas: inputValue.payload1 = &H123456789ABCDEF^
    pointer = CallEcho(inputValue, result, 8&)
    If result.payload0 = areas Or result.payload1 <> inputValue.payload1 Then Err.Raise 530&
    MoveBytes VarPtr(copied), result.payload0 + 20^, 16^
    If copied.firstRow <> 5& Or copied.lastRow <> 6& Or copied.firstColumn <> 7& Or copied.lastColumn <> 8& Then Err.Raise 531&
    FreeEcho pointer: Release areas
    text = String$(32767&, ChrW$(955))
    encoded = CountedString(text)
    inputValue.kind = 2&: inputValue.payload0 = encoded
    pointer = CallEcho(inputValue, result, 2&)
    If TextAt(result.payload0) <> text Then Err.Raise 532&
    FreeEcho pointer: Release encoded
    If NShutdown() <> 0& Then Err.Raise 533&
    ProbeNativeExports = "numeric-types|number|string|boolean|errors|missing|nil|integer|matrix|sref|mref|long-string|owned=0"
End Function
Public Sub Main()
    Dim observed As String
    observed = ProbeNativeExports()
End Sub
