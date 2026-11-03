Option Explicit
Private Declare PtrSafe Function NativeAddLong Lib "__DLL__" Alias "AddLong" (ByVal a As Long, ByVal b As Long) As Long
Private Declare PtrSafe Function NativeDirectLong Lib "__DLL__" Alias "DirectLong" (ByVal a As Long, ByVal b As Long) As Long
Private Declare PtrSafe Function NativeAddDouble Lib "__DLL__" Alias "AddDouble" (ByVal a As Double, ByVal b As Double) As Double
Private Declare PtrSafe Function NativeGreater Lib "__DLL__" Alias "Greater" (ByVal a As Long, ByVal b As Long) As Boolean
Private Declare PtrSafe Function NativeEchoBoolean Lib "__DLL__" Alias "EchoBoolean" (ByVal value As Boolean) As Boolean
Private Declare PtrSafe Function NativeMixed5 Lib "__DLL__" Alias "Mixed5" (ByVal a As Double, ByVal b As Long, ByVal flag As Boolean, ByVal d As Double, ByVal e As Long) As Double
Private Declare PtrSafe Function NativeDivide Lib "__DLL__" Alias "Divide" (ByVal a As Double, ByVal b As Double) As Double
Private Declare PtrSafe Function NativeLibraryAbs Lib "__DLL__" Alias "LibraryAbs" (ByVal value As Double) As Double
Private Declare PtrSafe Function NativeLibrarySqr Lib "__DLL__" Alias "LibrarySqr" (ByVal value As Double) As Double

Private Declare PtrSafe Function NativeRoundLong Lib "__DLL__" Alias "RoundLong" (ByVal value As Double) As Long
Private Declare PtrSafe Function NativeRoundThenAdd Lib "__DLL__" Alias "RoundThenAdd" (ByVal value As Double) As Long
Private Declare PtrSafe Function NativeBooleanDouble Lib "__DLL__" Alias "BooleanDouble" (ByVal value As Boolean) As Double
Private Declare PtrSafe Function NativeBooleanLong Lib "__DLL__" Alias "BooleanLong" (ByVal value As Boolean) As Long
Private Declare PtrSafe Function NativeDoubleBoolean Lib "__DLL__" Alias "DoubleBoolean" (ByVal value As Double) As Boolean
Private Declare PtrSafe Function NativeLongBoolean Lib "__DLL__" Alias "LongBoolean" (ByVal value As Long) As Boolean
Private Declare PtrSafe Function NativeSumTo Lib "__DLL__" Alias "SumTo" (ByVal value As Long) As Long
Private Declare PtrSafe Function NativeDepth Lib "__DLL__" Alias "Depth" (ByVal value As Long) As Long
Private Declare PtrSafe Function NativeWeighted32 Lib "__DLL__" Alias "Weighted32" (ByVal a01 As Long, ByVal a02 As Long, ByVal a03 As Long, ByVal a04 As Long, ByVal a05 As Long, ByVal a06 As Long, ByVal a07 As Long, ByVal a08 As Long, ByVal a09 As Long, ByVal a10 As Long, ByVal a11 As Long, ByVal a12 As Long, ByVal a13 As Long, ByVal a14 As Long, ByVal a15 As Long, ByVal a16 As Long, ByVal a17 As Long, ByVal a18 As Long, ByVal a19 As Long, ByVal a20 As Long, ByVal a21 As Long, ByVal a22 As Long, ByVal a23 As Long, ByVal a24 As Long, ByVal a25 As Long, ByVal a26 As Long, ByVal a27 As Long, ByVal a28 As Long, ByVal a29 As Long, ByVal a30 As Long, ByVal a31 As Long, ByVal a32 As Long) As Long

Private Declare PtrSafe Function NativeIntegerDivide Lib "__DLL__" Alias "IntegerDivide" (ByVal a As Long, ByVal b As Long) As Long
Private Declare PtrSafe Function NativeIntegerMod Lib "__DLL__" Alias "IntegerMod" (ByVal a As Long, ByVal b As Long) As Long
Private Declare PtrSafe Function NativeBooleanAnd Lib "__DLL__" Alias "BooleanAnd" (ByVal a As Boolean, ByVal b As Boolean) As Boolean
Private Declare PtrSafe Function NativeBooleanOr Lib "__DLL__" Alias "BooleanOr" (ByVal a As Boolean, ByVal b As Boolean) As Boolean
Private Declare PtrSafe Function NativeBooleanXor Lib "__DLL__" Alias "BooleanXor" (ByVal a As Boolean, ByVal b As Boolean) As Boolean
Private Declare PtrSafe Function NativeBooleanEqv Lib "__DLL__" Alias "BooleanEqv" (ByVal a As Boolean, ByVal b As Boolean) As Boolean
Private Declare PtrSafe Function NativeBooleanImp Lib "__DLL__" Alias "BooleanImp" (ByVal a As Boolean, ByVal b As Boolean) As Boolean
Private Declare PtrSafe Function NativeBooleanNot Lib "__DLL__" Alias "BooleanNot" (ByVal a As Boolean) As Boolean
Private Declare PtrSafe Function NativeLongAnd Lib "__DLL__" Alias "LongAnd" (ByVal a As Long, ByVal b As Long) As Long
Private Declare PtrSafe Function NativeLongOr Lib "__DLL__" Alias "LongOr" (ByVal a As Long, ByVal b As Long) As Long
Private Declare PtrSafe Function NativeLongXor Lib "__DLL__" Alias "LongXor" (ByVal a As Long, ByVal b As Long) As Long
Private Declare PtrSafe Function NativeLongEqv Lib "__DLL__" Alias "LongEqv" (ByVal a As Long, ByVal b As Long) As Long
Private Declare PtrSafe Function NativeLongImp Lib "__DLL__" Alias "LongImp" (ByVal a As Long, ByVal b As Long) As Long
Private Declare PtrSafe Function NativeLongNot Lib "__DLL__" Alias "LongNot" (ByVal a As Long) As Long

Public Function ProbeNativeExports() As String
    ProbeNativeExports = CStr(NativeAddLong(-7&, 49&)) & "|" & CStr(NativeDirectLong(5&, 32&))
    ProbeNativeExports = ProbeNativeExports & "|" & CStr(NativeAddDouble(1.25, 2.5))
    ProbeNativeExports = ProbeNativeExports & "|" & CStr(CLng(NativeGreater(5&, -2&))) & "|" & CStr(CLng(NativeGreater(-2&, 5&)))
    ProbeNativeExports = ProbeNativeExports & "|" & CStr(CLng(NativeEchoBoolean(True))) & "|" & CStr(CLng(NativeEchoBoolean(False)))
    ProbeNativeExports = ProbeNativeExports & "|" & CStr(NativeMixed5(1.25, 2&, True, 3.5, 4&)) & "|" & CStr(NativeMixed5(1.25, 2&, False, 3.5, 4&))
    ProbeNativeExports = ProbeNativeExports & "|" & CStr(NativeDivide(9#, 2#)) & "|" & CStr(NativeDivide(9#, 0#))
    ProbeNativeExports = ProbeNativeExports & "|" & CStr(NativeAddLong(2147483647, 1&))
    ProbeNativeExports = ProbeNativeExports & "|" & CStr(NativeLibraryAbs(-12.5)) & "|" & CStr(NativeLibrarySqr(25#)) & "|" & CStr(NativeLibrarySqr(-1#))
    ProbeNativeExports = ProbeNativeExports & "|" & CStr(NativeRoundLong(-3.5))
    ProbeNativeExports = ProbeNativeExports & "|" & CStr(NativeRoundLong(-2.5))
    ProbeNativeExports = ProbeNativeExports & "|" & CStr(NativeRoundLong(-1.5))
    ProbeNativeExports = ProbeNativeExports & "|" & CStr(NativeRoundLong(-0.5))
    ProbeNativeExports = ProbeNativeExports & "|" & CStr(NativeRoundLong(0.5))
    ProbeNativeExports = ProbeNativeExports & "|" & CStr(NativeRoundLong(1.5))
    ProbeNativeExports = ProbeNativeExports & "|" & CStr(NativeRoundLong(2.5))
    ProbeNativeExports = ProbeNativeExports & "|" & CStr(NativeRoundLong(3.5))
    ProbeNativeExports = ProbeNativeExports & "|" & CStr(NativeRoundLong(2147483647#))
    ProbeNativeExports = ProbeNativeExports & "|" & CStr(NativeRoundLong(2147483647.5))
    ProbeNativeExports = ProbeNativeExports & "|" & CStr(NativeRoundLong(-2147483648.5))
    ProbeNativeExports = ProbeNativeExports & "|" & CStr(NativeRoundLong(-2147483648.6))
    ProbeNativeExports = ProbeNativeExports & "|" & CStr(NativeRoundThenAdd(2.5))
    ProbeNativeExports = ProbeNativeExports & "|" & CStr(NativeRoundThenAdd(1E+30))
    ProbeNativeExports = ProbeNativeExports & "|" & CStr(NativeBooleanDouble(True))
    ProbeNativeExports = ProbeNativeExports & "|" & CStr(NativeBooleanDouble(False))
    ProbeNativeExports = ProbeNativeExports & "|" & CStr(NativeBooleanLong(True))
    ProbeNativeExports = ProbeNativeExports & "|" & CStr(NativeBooleanLong(False))
    ProbeNativeExports = ProbeNativeExports & "|" & CStr(CLng(NativeDoubleBoolean(-0.5)))
    ProbeNativeExports = ProbeNativeExports & "|" & CStr(CLng(NativeDoubleBoolean(0#)))
    ProbeNativeExports = ProbeNativeExports & "|" & CStr(CLng(NativeDoubleBoolean(1E-100)))
    ProbeNativeExports = ProbeNativeExports & "|" & CStr(CLng(NativeLongBoolean(-2147483647)))
    ProbeNativeExports = ProbeNativeExports & "|" & CStr(CLng(NativeLongBoolean(0&)))
    ProbeNativeExports = ProbeNativeExports & "|" & CStr(CLng(NativeLongBoolean(1&)))
    ProbeNativeExports = ProbeNativeExports & "|" & CStr(NativeSumTo(10&))
    ProbeNativeExports = ProbeNativeExports & "|" & CStr(NativeSumTo(65535))
    ProbeNativeExports = ProbeNativeExports & "|" & CStr(NativeSumTo(65536))
    ProbeNativeExports = ProbeNativeExports & "|" & CStr(NativeDepth(32&))
    ProbeNativeExports = ProbeNativeExports & "|" & CStr(NativeDepth(127&))
    ProbeNativeExports = ProbeNativeExports & "|" & CStr(NativeDepth(128&))
    ProbeNativeExports = ProbeNativeExports & "|" & CStr(NativeDepth(10000&))
    ProbeNativeExports = ProbeNativeExports & "|" & CStr(NativeDepth(32&))
    ProbeNativeExports = ProbeNativeExports & "|" & CStr(NativeWeighted32(1&, 2&, 3&, 4&, 5&, 6&, 7&, 8&, 9&, 10&, 11&, 12&, 13&, 14&, 15&, 16&, 17&, 18&, 19&, 20&, 21&, 22&, 23&, 24&, 25&, 26&, 27&, 28&, 29&, 30&, 31&, 32&))
    ProbeNativeExports = ProbeNativeExports & "|" & CStr(NativeWeighted32(0&, 0&, 0&, 0&, 0&, 0&, 0&, 0&, 0&, 0&, 0&, 0&, 0&, 0&, 0&, 0&, 0&, 0&, 0&, 0&, 0&, 0&, 0&, 0&, 0&, 0&, 0&, 0&, 0&, 0&, 0&, -7&))
    If NativeIntegerDivide(-7&, 3&) <> ReferenceIntegerDivide(-7&, 3&) Then Err.Raise 5
    ProbeNativeExports = ProbeNativeExports & "|" & CStr(NativeIntegerDivide(-7&, 3&))
    If NativeIntegerDivide(7&, -3&) <> ReferenceIntegerDivide(7&, -3&) Then Err.Raise 5
    ProbeNativeExports = ProbeNativeExports & "|" & CStr(NativeIntegerDivide(7&, -3&))
    If NativeIntegerDivide((-2147483647 - 1&), -1&) <> ReferenceIntegerDivide((-2147483647 - 1&), -1&) Then Err.Raise 5
    ProbeNativeExports = ProbeNativeExports & "|" & CStr(NativeIntegerDivide((-2147483647 - 1&), -1&))
    If NativeIntegerDivide(9&, 0&) <> ReferenceIntegerDivide(9&, 0&) Then Err.Raise 5
    ProbeNativeExports = ProbeNativeExports & "|" & CStr(NativeIntegerDivide(9&, 0&))
    If NativeIntegerDivide((-2147483647 - 1&), 1&) <> ReferenceIntegerDivide((-2147483647 - 1&), 1&) Then Err.Raise 5
    ProbeNativeExports = ProbeNativeExports & "|" & CStr(NativeIntegerDivide((-2147483647 - 1&), 1&))
    If NativeIntegerMod(-7&, 3&) <> ReferenceIntegerMod(-7&, 3&) Then Err.Raise 5
    ProbeNativeExports = ProbeNativeExports & "|" & CStr(NativeIntegerMod(-7&, 3&))
    If NativeIntegerMod(7&, -3&) <> ReferenceIntegerMod(7&, -3&) Then Err.Raise 5
    ProbeNativeExports = ProbeNativeExports & "|" & CStr(NativeIntegerMod(7&, -3&))
    If NativeIntegerMod((-2147483647 - 1&), -1&) <> ReferenceIntegerMod((-2147483647 - 1&), -1&) Then Err.Raise 5
    ProbeNativeExports = ProbeNativeExports & "|" & CStr(NativeIntegerMod((-2147483647 - 1&), -1&))
    If NativeIntegerMod(9&, 0&) <> ReferenceIntegerMod(9&, 0&) Then Err.Raise 5
    ProbeNativeExports = ProbeNativeExports & "|" & CStr(NativeIntegerMod(9&, 0&))
    If NativeIntegerMod((-2147483647 - 1&), 1&) <> ReferenceIntegerMod((-2147483647 - 1&), 1&) Then Err.Raise 5
    ProbeNativeExports = ProbeNativeExports & "|" & CStr(NativeIntegerMod((-2147483647 - 1&), 1&))
    If NativeBooleanAnd(False, False) <> (False And False) Then Err.Raise 5
    ProbeNativeExports = ProbeNativeExports & "|" & CStr(CLng(NativeBooleanAnd(False, False)))
    If NativeBooleanAnd(False, True) <> (False And True) Then Err.Raise 5
    ProbeNativeExports = ProbeNativeExports & "|" & CStr(CLng(NativeBooleanAnd(False, True)))
    If NativeBooleanAnd(True, False) <> (True And False) Then Err.Raise 5
    ProbeNativeExports = ProbeNativeExports & "|" & CStr(CLng(NativeBooleanAnd(True, False)))
    If NativeBooleanAnd(True, True) <> (True And True) Then Err.Raise 5
    ProbeNativeExports = ProbeNativeExports & "|" & CStr(CLng(NativeBooleanAnd(True, True)))
    If NativeBooleanOr(False, False) <> (False Or False) Then Err.Raise 5
    ProbeNativeExports = ProbeNativeExports & "|" & CStr(CLng(NativeBooleanOr(False, False)))
    If NativeBooleanOr(False, True) <> (False Or True) Then Err.Raise 5
    ProbeNativeExports = ProbeNativeExports & "|" & CStr(CLng(NativeBooleanOr(False, True)))
    If NativeBooleanOr(True, False) <> (True Or False) Then Err.Raise 5
    ProbeNativeExports = ProbeNativeExports & "|" & CStr(CLng(NativeBooleanOr(True, False)))
    If NativeBooleanOr(True, True) <> (True Or True) Then Err.Raise 5
    ProbeNativeExports = ProbeNativeExports & "|" & CStr(CLng(NativeBooleanOr(True, True)))
    If NativeBooleanXor(False, False) <> (False Xor False) Then Err.Raise 5
    ProbeNativeExports = ProbeNativeExports & "|" & CStr(CLng(NativeBooleanXor(False, False)))
    If NativeBooleanXor(False, True) <> (False Xor True) Then Err.Raise 5
    ProbeNativeExports = ProbeNativeExports & "|" & CStr(CLng(NativeBooleanXor(False, True)))
    If NativeBooleanXor(True, False) <> (True Xor False) Then Err.Raise 5
    ProbeNativeExports = ProbeNativeExports & "|" & CStr(CLng(NativeBooleanXor(True, False)))
    If NativeBooleanXor(True, True) <> (True Xor True) Then Err.Raise 5
    ProbeNativeExports = ProbeNativeExports & "|" & CStr(CLng(NativeBooleanXor(True, True)))
    If NativeBooleanEqv(False, False) <> (False Eqv False) Then Err.Raise 5
    ProbeNativeExports = ProbeNativeExports & "|" & CStr(CLng(NativeBooleanEqv(False, False)))
    If NativeBooleanEqv(False, True) <> (False Eqv True) Then Err.Raise 5
    ProbeNativeExports = ProbeNativeExports & "|" & CStr(CLng(NativeBooleanEqv(False, True)))
    If NativeBooleanEqv(True, False) <> (True Eqv False) Then Err.Raise 5
    ProbeNativeExports = ProbeNativeExports & "|" & CStr(CLng(NativeBooleanEqv(True, False)))
    If NativeBooleanEqv(True, True) <> (True Eqv True) Then Err.Raise 5
    ProbeNativeExports = ProbeNativeExports & "|" & CStr(CLng(NativeBooleanEqv(True, True)))
    If NativeBooleanImp(False, False) <> (False Imp False) Then Err.Raise 5
    ProbeNativeExports = ProbeNativeExports & "|" & CStr(CLng(NativeBooleanImp(False, False)))
    If NativeBooleanImp(False, True) <> (False Imp True) Then Err.Raise 5
    ProbeNativeExports = ProbeNativeExports & "|" & CStr(CLng(NativeBooleanImp(False, True)))
    If NativeBooleanImp(True, False) <> (True Imp False) Then Err.Raise 5
    ProbeNativeExports = ProbeNativeExports & "|" & CStr(CLng(NativeBooleanImp(True, False)))
    If NativeBooleanImp(True, True) <> (True Imp True) Then Err.Raise 5
    ProbeNativeExports = ProbeNativeExports & "|" & CStr(CLng(NativeBooleanImp(True, True)))
    If NativeBooleanNot(False) <> (Not False) Then Err.Raise 5
    ProbeNativeExports = ProbeNativeExports & "|" & CStr(CLng(NativeBooleanNot(False)))
    If NativeBooleanNot(True) <> (Not True) Then Err.Raise 5
    ProbeNativeExports = ProbeNativeExports & "|" & CStr(CLng(NativeBooleanNot(True)))
    If NativeLongAnd(-123&, 456&) <> (-123& And 456&) Then Err.Raise 5
    ProbeNativeExports = ProbeNativeExports & "|" & CStr(NativeLongAnd(-123&, 456&))
    If NativeLongAnd((-2147483647 - 1&), 65535) <> ((-2147483647 - 1&) And 65535) Then Err.Raise 5
    ProbeNativeExports = ProbeNativeExports & "|" & CStr(NativeLongAnd((-2147483647 - 1&), 65535))
    If NativeLongOr(-123&, 456&) <> (-123& Or 456&) Then Err.Raise 5
    ProbeNativeExports = ProbeNativeExports & "|" & CStr(NativeLongOr(-123&, 456&))
    If NativeLongOr((-2147483647 - 1&), 65535) <> ((-2147483647 - 1&) Or 65535) Then Err.Raise 5
    ProbeNativeExports = ProbeNativeExports & "|" & CStr(NativeLongOr((-2147483647 - 1&), 65535))
    If NativeLongXor(-123&, 456&) <> (-123& Xor 456&) Then Err.Raise 5
    ProbeNativeExports = ProbeNativeExports & "|" & CStr(NativeLongXor(-123&, 456&))
    If NativeLongXor((-2147483647 - 1&), 65535) <> ((-2147483647 - 1&) Xor 65535) Then Err.Raise 5
    ProbeNativeExports = ProbeNativeExports & "|" & CStr(NativeLongXor((-2147483647 - 1&), 65535))
    If NativeLongEqv(-123&, 456&) <> (-123& Eqv 456&) Then Err.Raise 5
    ProbeNativeExports = ProbeNativeExports & "|" & CStr(NativeLongEqv(-123&, 456&))
    If NativeLongEqv((-2147483647 - 1&), 65535) <> ((-2147483647 - 1&) Eqv 65535) Then Err.Raise 5
    ProbeNativeExports = ProbeNativeExports & "|" & CStr(NativeLongEqv((-2147483647 - 1&), 65535))
    If NativeLongImp(-123&, 456&) <> (-123& Imp 456&) Then Err.Raise 5
    ProbeNativeExports = ProbeNativeExports & "|" & CStr(NativeLongImp(-123&, 456&))
    If NativeLongImp((-2147483647 - 1&), 65535) <> ((-2147483647 - 1&) Imp 65535) Then Err.Raise 5
    ProbeNativeExports = ProbeNativeExports & "|" & CStr(NativeLongImp((-2147483647 - 1&), 65535))
    If NativeLongNot((-2147483647 - 1&)) <> (Not (-2147483647 - 1&)) Then Err.Raise 5
    ProbeNativeExports = ProbeNativeExports & "|" & CStr(NativeLongNot((-2147483647 - 1&)))
    If NativeLongNot(0&) <> (Not 0&) Then Err.Raise 5
    ProbeNativeExports = ProbeNativeExports & "|" & CStr(NativeLongNot(0&))
    If NativeLongNot(2147483647) <> (Not 2147483647) Then Err.Raise 5
    ProbeNativeExports = ProbeNativeExports & "|" & CStr(NativeLongNot(2147483647))
End Function

Public Sub Main()
    Dim result As String
    result = ProbeNativeExports()
End Sub

Private Function ReferenceIntegerDivide(ByVal a As Long, ByVal b As Long) As Long
    On Error Resume Next
    ReferenceIntegerDivide = a \ b
    Err.Clear
End Function

Private Function ReferenceIntegerMod(ByVal a As Long, ByVal b As Long) As Long
    On Error Resume Next
    ReferenceIntegerMod = a Mod b
    Err.Clear
End Function
