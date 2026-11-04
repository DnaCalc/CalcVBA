Attribute VB_Name = "Wide"
Option Explicit
Public Function IncByte(ByRef value As Byte) As Byte
value=value+1&
IncByte=value
End Function
Public Function IncInteger(ByRef value As Integer) As Integer
value=value+1&
IncInteger=value
End Function
Public Function IncLongLong(ByRef value As LongLong) As LongLong
value=value+1^
IncLongLong=value
End Function
Public Function IncPointer(ByRef value As LongPtr) As LongPtr
value=value+8^
IncPointer=value
End Function
Public Function IncSingle(ByRef value As Single) As Single
value=value+0.5!
IncSingle=value
End Function
Public Function IncCurrency(ByRef value As Currency) As Currency
value=value+0.0001@
IncCurrency=value
End Function
Public Function IncDate(ByRef value As Date) As Date
value=value+0.5#
IncDate=value
End Function
Public Function SumWidths(ByVal b As Byte, ByVal i As Integer, ByVal ll As LongLong, ByVal s As Single, ByVal c As Currency, ByVal d As Date) As Double
SumWidths=CDbl(b)+CDbl(i)+CDbl(ll)+CDbl(s)+CDbl(c)+CDbl(d)
End Function
Public Function WideSuite() As String
Dim b As Byte, i As Integer, ll As LongLong, p As LongPtr, s As Single, c As Currency, d As Date
Dim result As String
b=CByte(254):result=CStr(IncByte(b)) & ":" & CStr(b)
i=CInt(-2):result=result & ":" & CStr(IncInteger(i)) & ":" & CStr(i)
ll=CLngLng(&H20000000000001^):result=result & ":" & CStr(IncLongLong(ll)) & ":" & CStr(ll)
p=CLngPtr(&H123456789ABCDEF^):result=result & ":" & CStr(IncPointer(p)) & ":" & CStr(p)
s=CSng(1.25!):result=result & ":" & CStr(IncSingle(s)) & ":" & CStr(s)
c=CCur(1.2345@):result=result & ":" & CStr(CLng(IncCurrency(c)*10000@)) & ":" & CStr(CLng(c*10000@))
d=CDate(45000.25#):result=result & ":" & CStr(CDbl(IncDate(d))) & ":" & CStr(CDbl(d))
result=result & ":" & CStr(SumWidths(5,-3,100^,1.25!,2.5@,45000.75#))
WideSuite=result
End Function
