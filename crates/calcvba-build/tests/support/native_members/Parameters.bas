Attribute VB_Name = "Parameters"
Option Explicit
Option Base 1
Private trace As String
Private stored As String * 4
Public Function ParameterSuite() As String
    Dim x As Variant, y As Variant, text As Variant
    x=5: y=9: text="a"
    Change x,x,y,text,(x)
    ParameterSuite=CStr(x) & ":" & CStr(y) & ":" & text & ":" & trace & "|" & EmptyBounds()
End Function
Private Sub Change(ParamArray items() As Variant)
    trace=CStr(LBound(items)) & ":" & CStr(UBound(items))
    items(0)=11
    trace=trace & ":" & CStr(items(1))
    items(1)=13
    trace=trace & ":" & CStr(items(0))
    items(2)=15
    items(3)="b"
    items(4)=19
End Sub
Private Function EmptyBounds(ParamArray items() As Variant) As String
    EmptyBounds=CStr(LBound(items)) & ":" & CStr(UBound(items))
End Function
Public Function FixedSuite() As String
    Dim value As String * 3
    FixedSuite=CStr(Len(value)) & ":[" & value & "]:"
    value="abcde"
    FixedSuite=FixedSuite & value & ":"
    value="x"
    stored="xy"
    FixedSuite=FixedSuite & "[" & value & "]:" & CStr(Len(stored)) & ":[" & stored & "]"
End Function
Public Function GoSubSuite() As String
    trace=""
    GoSub Outer
    GoSub Outer
    GoSubSuite=trace
    On Error Resume Next
    Return
    GoSubSuite=GoSubSuite & "|" & CStr(Err.Number)
    Err.Clear
    Exit Function
Outer:
    trace=trace & "O"
    GoSub Inner
    trace=trace & "R"
    Return
Inner:
    trace=trace & "I"
    Return
End Function
Private Function OptionalCallee(Optional ByVal x As Long=7, Optional value As Variant, Optional ByVal text As String="s") As String
    OptionalCallee=CStr(x) & ":" & CStr(IsMissing(value)) & ":" & text
End Function
Public Function OptionalSuite() As String
    OptionalSuite=OptionalCallee() & "|" & OptionalCallee(value:=9,text:="t") & "|" & OptionalCallee(2, , "q")
End Function
Public Sub Main()
    Dim observed As String
    observed=ParameterSuite() & "#" & FixedSuite() & "#" & GoSubSuite() & "#" & OptionalSuite() & "#" & NumericWidthSuite() & "#" & TypedParameterSuite()
End Sub
Public Function NumericWidthSuite() As String
    Dim fraction As Double, first As Long, second As Long, nested As Variant
    fraction=-12.5
    first=CLng(CInt(fraction))
    second=CInt(-3)
    nested=CLng(CInt(7))
    NumericWidthSuite=CStr(first) & ":" & CStr(second) & ":" & CStr(nested) & ":" & TypeName(nested)
End Function
Public Function TypedParameterSuite() As String
    Dim number As Byte, text As String * 3
    number = 1: text = "a"
    TypedChange number, text, number, (number)
    TypedParameterSuite = CStr(number) & ":[" & text & "]:" & trace
End Function
Private Sub TypedChange(ParamArray items() As Variant)
    On Error Resume Next
    items(0) = 300
    trace = CStr(Err.Number) & ":" & TypeName(items(0)) & ":" & CStr(items(0))
    Err.Clear
    items(1) = "zzzz"
    trace = trace & ":[" & CStr(items(1)) & "]"
    items(2) = 9
    trace = trace & ":" & TypeName(items(0)) & ":" & CStr(items(0))
    items(3) = 7
    trace = trace & ":" & CStr(items(3))
End Sub
