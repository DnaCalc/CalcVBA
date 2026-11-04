Attribute VB_Name = "Declares"
Option Explicit
Public Type Payload
    tag As Byte
    fixed As String * 4
    number As Long
    text As String
    value As Variant
End Type
Public Type Pair
    first As Long
    second As Long
End Type
Public Type Returned
    first As Long
    second As Long
    text As String
    value As Variant
End Type
Private Declare PtrSafe Function Change Lib "__FIXTURE__" Alias "FixtureRecord" (ByRef value As Any) As Long
Private Declare PtrSafe Function GetPair Lib "__FIXTURE__" Alias "FixturePair" (ByVal value As Long) As Pair
Private Declare PtrSafe Function GetReturned Lib "__FIXTURE__" Alias "FixtureReturned" () As Returned
Private Declare PtrSafe Function GetVariant Lib "__FIXTURE__" Alias "FixtureVariant" (ByVal value As Long) As Variant
Private Declare PtrSafe Function GetVariantText Lib "__FIXTURE__" Alias "FixtureVariantText" () As Variant
Private Declare PtrSafe Function AliasAnsi Lib "__FIXTURE__" Alias "FixtureAnsiAlias" (ByRef first As String, ByRef second As String) As Long
Public Function ForeignSuite() As String
    Dim payload As Payload, pair As Pair, returned As Returned, value As Variant, result As String, text As String, status As Long
    payload.tag = 7: payload.fixed = "ABCD": payload.number = 37&: payload.text = "abc": payload.value = "wide"
    status = Change(payload)
    result = CStr(status) & ":" & payload.fixed & ":" & CStr(payload.number) & ":" & payload.text & ":" & CStr(payload.value)
    value = GetVariant(37&): result = result & "|" & CStr(value)
    value = GetVariantText(): result = result & "|" & CStr(value)
    pair = GetPair(37&): result = result & "|" & CStr(pair.first) & ":" & CStr(pair.second)
    returned = GetReturned(): result = result & "|" & CStr(returned.first) & ":" & returned.text & ":" & CStr(returned.value)
    text = "seed": status = AliasAnsi(text, text)
    ForeignSuite = result & "|" & CStr(status) & ":" & text
End Function
Public Function ProbeNativeExports() As String
    ProbeNativeExports = ForeignSuite()
End Function
Public Sub Main()
    Dim result As String
    result = ForeignSuite()
End Sub
