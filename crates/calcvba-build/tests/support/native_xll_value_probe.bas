Option Explicit
Private Declare PtrSafe Sub PeekHeader Lib "kernel32" Alias "RtlMoveMemory" (ByVal destination As LongPtr, ByVal source As LongPtr, ByVal bytes As LongPtr)
Private Sub Require(ByVal condition As Boolean, ByVal message As String)
    If Not condition Then Err.Raise 513&, "XlValueProbe", message
End Sub
Public Function ProbeNativeExports() As String
    Dim value As Variant, returned As Variant, oper As XLOPER12, echoed As XLOPER12
    Dim values As Variant, i As Long, text As String
    Dim subtype As Integer, scode As Long
    value = CInt(7)
    PeekHeader VarPtr(subtype), VarPtr(value), 2^
    Call Require(subtype = 2, "native Integer Variant tag")
    value = CVErr(2042&)
    PeekHeader VarPtr(scode), VarPtr(value) + 8^, 4^
    Call Require(scode = &H800A07FA, "native error SCODE")
    values = Array(CByte(7), CInt(-123), CLng(456), CLngLng(9007199254740993^), CSng(1.25), CDbl(2.75), CCur(12.3456), CDate(45123.25), CDec(1.125))
    For i = 0& To 8&
        value = values(i)
        oper = XlFromVariant(value)
        returned = XlToVariant(oper)
        Call Require(oper.xltype = 1&, "numeric type")
        Call Require(CDbl(returned) = CDbl(value), "numeric value")
        XlFreeValue oper
    Next i
    text = "A" & ChrW$(0) & ChrW$(955) & ChrW$(937)
    value = text: oper = XlFromVariant(value): returned = XlToVariant(oper)
    Call Require(VarType(returned) = 8& And CStr(returned) = text, "Unicode string")
    XlFreeValue oper
    value = True: oper = XlFromVariant(value): returned = XlToVariant(oper)
    Call Require(VarType(returned) = 11& And CBool(returned), "Boolean")
    XlFreeValue oper
    value = CVErr(2042&): oper = XlFromVariant(value): returned = XlToVariant(oper)
    echoed = XlFromVariant(returned)
    Call Require(oper.xltype = 16& And oper.payload0 = 42^ And echoed.payload0 = 42^, "error value")
    XlFreeValue oper: XlFreeValue echoed
    oper.xltype = 128&: returned = XlToVariant(oper): echoed = XlFromVariant(returned)
    Call Require(IsMissing(returned) And echoed.xltype = 128&, "missing value")
    XlFreeValue oper: XlFreeValue echoed
    value = Empty: oper = XlFromVariant(value): returned = XlToVariant(oper)
    Call Require(oper.xltype = 256& And IsEmpty(returned), "nil value")
    XlFreeValue oper
    value = Null: oper = XlFromVariant(value)
    Call Require(oper.xltype = 16& And oper.payload0 = 15^, "Null maps to VALUE")
    XlFreeValue oper
    Dim matrix(-1 To 0, 3 To 5) As Variant
    matrix(-1, 3) = 2.5: matrix(-1, 4) = text: matrix(-1, 5) = True
    matrix(0, 3) = CVErr(2042&): matrix(0, 4) = Empty: matrix(0, 5) = "tail"
    value = matrix
    oper = XlFromVariant(value): returned = XlToVariant(oper)
    Call Require(LBound(returned, 1) = 1& And UBound(returned, 1) = 2&, "matrix row bounds")
    Call Require(LBound(returned, 2) = 1& And UBound(returned, 2) = 3&, "matrix column bounds")
    Call Require(returned(1, 1) = 2.5 And returned(1, 2) = text And returned(1, 3) = True, "matrix first row")
    Call Require(IsError(returned(2, 1)) And IsEmpty(returned(2, 2)) And returned(2, 3) = "tail", "matrix second row")
    XlFreeValue oper
    Dim typed(-2 To -1) As Double
    typed(-2) = 3.25: typed(-1) = 7.5: value = typed
    oper = XlFromVariant(value): returned = XlToVariant(oper)
    Call Require(UBound(returned, 1) = 1& And UBound(returned, 2) = 2&, "typed vector shape")
    Call Require(returned(1, 1) = 3.25 And returned(1, 2) = 7.5, "typed vector values")
    XlFreeValue oper
    value = Array("#XlReference12", 8&, &H123456789ABCDEF^, 2&, 0&, 1&, 2&, 3&, 5&, 6&, 7&, 8&)
    oper = XlFromVariant(value): returned = XlToVariant(oper): echoed = XlFromVariant(returned)
    Call Require(oper.xltype = 8& And echoed.xltype = 8&, "reference kind")
    Call Require(oper.payload1 = &H123456789ABCDEF^ And echoed.payload1 = oper.payload1, "reference sheet ID")
    Call Require(returned(3) = 2& And returned(8) = 5& And returned(11) = 8&, "reference areas")
    XlFreeValue oper: XlFreeValue echoed
    value = Array("#XlReference12", 1024&, 0^, 1&, 2&, 4&, 1&, 3&)
    oper = XlFromVariant(value): returned = XlToVariant(oper): echoed = XlFromVariant(returned)
    Call Require(oper.xltype = 1024& And echoed.xltype = 1024& And returned(4) = 2& And returned(7) = 3&, "simple reference")
    XlFreeValue oper: XlFreeValue echoed
    Call Require(XlOwnedBuffers() = 0&, "owned buffer balance")
    ProbeNativeExports = "numeric-types|Unicode|Boolean|error|missing|nil|Null|matrix|typed-vector|references|owned=0"
End Function
Public Sub Main()
    Dim observed As String
    observed = ProbeNativeExports()
End Sub
