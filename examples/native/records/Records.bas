Attribute VB_Name = "Records"
Option Explicit
Private Declare PtrSafe Function ReallocateAnsi Lib "oleaut32" Alias "SysReAllocStringLen" (ByRef text As String, ByVal source As LongPtr, ByVal units As Long) As Long
Private Declare PtrSafe Function AllocateAnsi Lib "oleaut32" Alias "SysAllocStringByteLen" (ByVal text As String, ByVal bytes As Long) As String
Public Type RichRecord
    tag As Byte
    value As Variant
    tail As Long
    text As String * 4
End Type
Public Type FixedFields
    tag As Byte
    numbers(-1 To 1, 3 To 4) As Long
    texts(2 To 3) As String
    variants(0 To 1) As Variant
    fixedText(-2 To -1) As String * 3
    records(5 To 6) As RichRecord
End Type

Public Function FixedFieldSuite() As String
    Dim value As FixedFields, copied As FixedFields, row As Long, column As Long
    For column = 3& To 4&
        For row = -1& To 1&
            value.numbers(row, column) = row + column * 10&
        Next row
    Next column
    value.texts(2) = "left": value.texts(3) = "right"
    value.variants(0) = "42": value.variants(1) = 9&
    value.fixedText(-2) = "abcdef": value.fixedText(-1) = "q"
    Dim record As RichRecord
    record.value = "record": record.text = "xy"
    value.records(5) = record
    copied = value
    value.texts(2) = "changed": value.variants(0) = 0&
    record = copied.records(5)
    FixedFieldSuite = CStr(copied.numbers(-1, 3)) & ":" & CStr(copied.numbers(1, 4)) & ":" & copied.texts(2) & ":" & CStr(copied.variants(0)) & ":" & copied.fixedText(-2) & ":" & copied.fixedText(-1) & ":" & CStr(record.value)
End Function

Public Function FixedFieldIndex(ByVal index As Long) As Long
    Dim value As FixedFields
    value.numbers(-1, 3) = 29&
    FixedFieldIndex = value.numbers(index, 3)
End Function

Public Function ChangeRecord(ByRef record As RichRecord) As String
    record.value = record.value + 5&
    record.tail = record.tail + 1&
    record.text = "abcdef"
    ChangeRecord = CStr(record.value) & ":" & record.text & ":" & CStr(record.tail)
End Function

Public Function EchoRecord(ByRef record As RichRecord) As RichRecord
    EchoRecord = record
End Function

Public Function ValueSuite() As String
    Dim record As RichRecord, copied As RichRecord, description As String
    record.tag = 7
    record.value = "7"
    record.tail = 9&
    record.text = "q"
    description = ChangeRecord(record)
    copied = EchoRecord(record)
    record.value = "changed"
    ValueSuite = description & ":" & CStr(copied.value) & ":" & copied.text & ":" & CStr(copied.tail)
End Function

Public Function GrowStrings(ByRef values() As String) As Long
    Dim last As Long, first As Long
    first = LBound(values): last = UBound(values) + 1&
    ReDim Preserve values(first To last)
    values(last) = "new"
    GrowStrings = Len(values(first)) + Len(values(first + 1&)) + Len(values(last))
End Function

Public Sub AllocateStrings(ByRef values() As String)
    ReDim values(-2 To -1)
    values(-2) = "left"
    values(-1) = "right"
End Sub

Public Sub Main()
End Sub

Public Function ClearAnsi(ByRef text As String) As Long
    ClearAnsi = ReallocateAnsi(text, 0^, 0&)
End Function
Public Function AnsiText() As String
    AnsiText = AllocateAnsi("reply", 5&)
End Function

Public Sub AllocateRecords(ByRef values() As RichRecord)
    ReDim values(-2 To -1)
    values(-2).value = "7"
    values(-2).text = "q"
    values(-2).tail = 9&
End Sub
Public Function RecordArrays(ByRef values() As RichRecord) As String
    ReDim Preserve values(-2 To 0)
    values(0) = values(-2)
    values(-2).value = "changed"
    RecordArrays = CStr(values(0).value) & ":" & values(0).text & ":" & CStr(LBound(values)) & ":" & CStr(UBound(values))
End Function
Public Function ArraySuite() As String
    Dim values() As RichRecord
    AllocateRecords values
    ArraySuite = RecordArrays(values)
End Function
