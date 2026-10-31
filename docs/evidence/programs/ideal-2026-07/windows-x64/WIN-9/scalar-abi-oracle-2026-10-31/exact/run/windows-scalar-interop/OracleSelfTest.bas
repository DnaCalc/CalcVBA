Private g As Long
Private Declare PtrSafe Function NativeAlias Lib "C:\Work\DnaCalc\codex-runs\calcvba-exact-native-probe\scalar.dll" Alias "AliasLong" (ByRef a As Long, ByRef b As Long) As Long
Private Declare PtrSafe Function NativeCount Lib "C:\Work\DnaCalc\codex-runs\calcvba-exact-native-probe\scalar.dll" Alias "AliasCallCount" () As Long
Private Declare PtrSafe Function AI Lib "C:\Work\DnaCalc\codex-runs\calcvba-exact-native-probe\scalar.dll" Alias "AliasInteger" (ByRef a As Integer, ByRef b As Integer) As Integer
Private Declare PtrSafe Function AB Lib "C:\Work\DnaCalc\codex-runs\calcvba-exact-native-probe\scalar.dll" Alias "AliasByte" (ByRef a As Byte, ByRef b As Byte) As Byte
Private Declare PtrSafe Function AT Lib "C:\Work\DnaCalc\codex-runs\calcvba-exact-native-probe\scalar.dll" Alias "AliasBoolean" (ByRef a As Boolean, ByRef b As Boolean) As Boolean
Private Declare PtrSafe Function ADbl Lib "C:\Work\DnaCalc\codex-runs\calcvba-exact-native-probe\scalar.dll" Alias "AliasDouble" (ByRef a As Double, ByRef b As Double) As Double
Private Declare PtrSafe Function ASng Lib "C:\Work\DnaCalc\codex-runs\calcvba-exact-native-probe\scalar.dll" Alias "AliasSingle" (ByRef a As Single, ByRef b As Single) As Single
Private Declare PtrSafe Function ALL Lib "C:\Work\DnaCalc\codex-runs\calcvba-exact-native-probe\scalar.dll" Alias "AliasI64" (ByRef a As LongLong, ByRef b As LongLong) As LongLong
Private Declare PtrSafe Function AP Lib "C:\Work\DnaCalc\codex-runs\calcvba-exact-native-probe\scalar.dll" Alias "AliasI64" (ByRef a As LongPtr, ByRef b As LongPtr) As LongPtr
Private Declare PtrSafe Function AC Lib "C:\Work\DnaCalc\codex-runs\calcvba-exact-native-probe\scalar.dll" Alias "AliasCurrency" (ByRef a As Currency, ByRef b As Currency) As Currency
Private Declare PtrSafe Function AD Lib "C:\Work\DnaCalc\codex-runs\calcvba-exact-native-probe\scalar.dll" Alias "AliasDate" (ByRef a As Date, ByRef b As Date) As Date
Private Declare PtrSafe Function NativeCurrency Lib "C:\Work\DnaCalc\codex-runs\calcvba-exact-native-probe\scalar.dll" Alias "CurrencyValue" () As Currency
Private Declare PtrSafe Function NativeDate Lib "C:\Work\DnaCalc\codex-runs\calcvba-exact-native-probe\scalar.dll" Alias "DateValue" () As Date
Private Declare PtrSafe Function NativeByte Lib "C:\Work\DnaCalc\codex-runs\calcvba-exact-native-probe\scalar.dll" Alias "ByteValue" () As Byte
Private Declare PtrSafe Function NativeMixed Lib "C:\Work\DnaCalc\codex-runs\calcvba-exact-native-probe\scalar.dll" Alias "Mixed" (ByVal a As Long, ByVal b As Double, ByVal c As Single, ByVal d As Currency, ByVal e As Integer, ByVal f As Date, ByVal g As LongPtr, ByVal h As Boolean, ByVal i As Byte, ByVal j As LongLong) As Long
Private Function Forward(ByRef a As Long, ByRef b As Long) As Long
Forward = NativeAlias(a, b)
End Function
Private Function GlobalMix(ByRef a As Long) As Long
GlobalMix = NativeAlias(a, g)
End Function
Public Function RunProbe() As String
Dim x As Long, y As Long, r As Long, text As String
Dim i As Integer
Dim b As Byte
Dim truth As Boolean
Dim dbl As Double
Dim s As Single
Dim ll As LongLong
Dim ptr As LongPtr
Dim cy As Currency
Dim dt As Date
x = 1
r = NativeAlias(x, x)
text = CStr(r) & "|" & CStr(x)
x = 1: y = 1
r = NativeAlias(x, y)
text = text & "|" & CStr(r) & "|" & CStr(x) & "|" & CStr(y)
x = 1
r = Forward(x, x)
text = text & "|" & CStr(r) & "|" & CStr(x)
x = 1
r = NativeAlias((x), x)
text = text & "|" & CStr(r) & "|" & CStr(x)
g = 1
r = GlobalMix(g)
text = text & "|" & CStr(r) & "|" & CStr(g)
i = AI(i, i)
text = text & "|" & CStr(i = -32766)
b = AB(b, b)
text = text & "|" & CStr(b = 251)
truth = AT(truth, truth)
text = text & "|" & CStr(truth = False)
dbl = ADbl(dbl, dbl)
text = text & "|" & CStr(dbl = 123456789.75)
s = ASng(s, s)
text = text & "|" & CStr(s = 123.75)
ll = ALL(ll, ll)
text = text & "|" & CStr(ll = CLngLng(4294967338#))
ptr = AP(ptr, ptr)
text = text & "|" & CStr(ptr = CLngPtr(4294967338#))
cy = AC(cy, cy)
text = text & "|" & CStr(cy = 123456790@)
dt = AD(dt, dt)
text = text & "|" & CStr(dt = CDate(45123.75))
text = text & "|" & CStr(NativeCount())
cy = 1234.5@: dt = CDate(45123.75): ptr = CLngPtr(4294967338#): ll = CLngLng(-9876543210987#)
text = text & "|" & CStr(NativeCurrency() = 123456789.0123@) & "|" & CStr(CDbl(NativeDate()) = 45123.75) & "|" & TypeName(NativeByte()) & "|" & CStr(NativeMixed(19, 23.5, 12.5!, cy, -2, dt, ptr, True, 251, ll))
RunProbe = text & "|" & CStr(Err.Number) & "|" & CStr(Err.LastDllError) & "|" & Err.Source & "|" & Err.Description
End Function
