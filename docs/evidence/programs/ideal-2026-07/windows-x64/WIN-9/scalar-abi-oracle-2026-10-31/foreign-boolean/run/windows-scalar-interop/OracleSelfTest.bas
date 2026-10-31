Private Declare PtrSafe Function CurrencyValue Lib "C:\Work\DnaCalc\codex-runs\calcvba-native-scalar-probe\scalar.dll" () As Currency
Private Declare PtrSafe Function DateValue Lib "C:\Work\DnaCalc\codex-runs\calcvba-native-scalar-probe\scalar.dll" () As Date
Private Declare PtrSafe Function ByteValue Lib "C:\Work\DnaCalc\codex-runs\calcvba-native-scalar-probe\scalar.dll" () As Byte
Private Declare PtrSafe Function BoolOne Lib "C:\Work\DnaCalc\codex-runs\calcvba-native-scalar-probe\scalar.dll" () As Boolean
Private Declare PtrSafe Function Mixed Lib "C:\Work\DnaCalc\codex-runs\calcvba-native-scalar-probe\scalar.dll" (ByVal a As Long, ByVal b As Double, ByVal c As Single, ByVal d As Currency, ByVal e As Integer, ByVal f As Date, ByVal g As LongPtr, ByVal h As Boolean, ByVal i As Byte, ByVal j As LongLong) As Long
Public Function RunProbe() As String
    Dim d As Date, c As Currency, p As LongPtr, ll As LongLong
    d = CDate(45123.75): c = 1234.5@: p = CLngPtr("4294967338"): ll = CLngLng("-9876543210987")
    RunProbe = CStr(CurrencyValue() = 123456789.0123@) & "|" & CStr(CDbl(DateValue()) = 45123.75) & "|" & TypeName(ByteValue()) & "|" & CStr(CLng(BoolOne())) & "|" & CStr(Mixed(19, 23.5, 12.5!, c, -2, d, p, True, 251, ll))
End Function
