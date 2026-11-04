Option Explicit
Private Declare PtrSafe Function NByte Lib "__DLL__" Alias "IncByte" (ByRef value As Byte) As Byte
Private Declare PtrSafe Function NInteger Lib "__DLL__" Alias "IncInteger" (ByRef value As Integer) As Integer
Private Declare PtrSafe Function NLongLong Lib "__DLL__" Alias "IncLongLong" (ByRef value As LongLong) As LongLong
Private Declare PtrSafe Function NPointer Lib "__DLL__" Alias "IncPointer" (ByRef value As LongPtr) As LongPtr
Private Declare PtrSafe Function NSingle Lib "__DLL__" Alias "IncSingle" (ByRef value As Single) As Single
Private Declare PtrSafe Function NCurrency Lib "__DLL__" Alias "IncCurrency" (ByRef value As Currency) As Currency
Private Declare PtrSafe Function NDate Lib "__DLL__" Alias "IncDate" (ByRef value As Date) As Date
Private Declare PtrSafe Function NSum Lib "__DLL__" Alias "SumWidths" (ByVal b As Byte, ByVal i As Integer, ByVal ll As LongLong, ByVal s As Single, ByVal c As Currency, ByVal d As Date) As Double
Private Declare PtrSafe Function NSuite Lib "__DLL__" Alias "WideSuite" () As LongPtr
Private Declare PtrSafe Function NError Lib "__DLL__" Alias "CalcVbaGetLastError" (ByVal source As LongPtr, ByVal description As LongPtr) As Long
Private Declare PtrSafe Function BStrLength Lib "oleaut32" Alias "SysStringLen" (ByVal text As LongPtr) As Long
Private Declare PtrSafe Sub FreeBStr Lib "oleaut32" Alias "SysFreeString" (ByVal text As LongPtr)
Private Declare PtrSafe Sub CopyBytes Lib "kernel32" Alias "RtlMoveMemory" (ByVal destination As LongPtr, ByVal source As LongPtr, ByVal count As LongPtr)
Private Function TakeString(ByVal pointer As LongPtr) As String
    Dim length As Long
    length=BStrLength(pointer)
    TakeString=String$(length,ChrW$(0))
    If length Then CopyBytes StrPtr(TakeString),pointer,length*2&
    FreeBStr pointer
End Function
Private Sub AppendValue(ByRef result As String, ByVal value As Variant)
    If Len(result) Then result=result & "|"
    result=result & CStr(value)
End Sub
Public Function ProbeNativeExports() As String
    Dim b As Byte, i As Integer, ll As LongLong, p As LongPtr, s As Single, c As Currency, d As Date
    Dim result As String, text As String
    b=254: AppendValue result,NByte(b): AppendValue result,b
    i=-2: AppendValue result,NInteger(i): AppendValue result,i
    ll=&H20000000000001^: AppendValue result,NLongLong(ll): AppendValue result,ll
    p=&H123456789ABCDEF^: AppendValue result,NPointer(p): AppendValue result,p
    s=1.25!: AppendValue result,NSingle(s): AppendValue result,s
    c=1.2345@: AppendValue result,CLng(NCurrency(c)*10000@): AppendValue result,CLng(c*10000@)
    d=45000.25#: AppendValue result,CDbl(NDate(d)): AppendValue result,CDbl(d)
    AppendValue result,NSum(5,-3,100^,1.25!,2.5@,45000.75#)
    text=TakeString(NSuite())
    If text <> "255:255:-1:-1:9007199254740994:9007199254740994:81985529216486903:81985529216486903:1.75:1.75:12346:12346:45000.75:45000.75:45106.5" Then Err.Raise 527
    AppendValue result,Len(text)
    AppendValue result,NByte(b): AppendValue result,b: AppendValue result,NError(0,0)
    i=32767: AppendValue result,NInteger(i): AppendValue result,i: AppendValue result,NError(0,0)
    ll=&H7FFFFFFFFFFFFFFF^: AppendValue result,NLongLong(ll): AppendValue result,ll: AppendValue result,NError(0,0)
    ProbeNativeExports=result
End Function
Public Sub Main()
    Dim result As String
    result=ProbeNativeExports()
End Sub
