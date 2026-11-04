Option Explicit
Private Type GridRecord
    values(-1 To 1, 3 To 4) As Long
End Type
Private Declare PtrSafe Function NGrid Lib "__DLL__" Alias "GridBoundary" (ByRef value As GridRecord) As Long
Private Declare PtrSafe Function NFixedSuite Lib "__DLL__" Alias "FixedFieldSuite" () As LongPtr
Private Declare PtrSafe Function NFixedIndex Lib "__DLL__" Alias "FixedFieldIndex" (ByVal index As Long) As Long
Private Declare PtrSafe Function NLastError Lib "__DLL__" Alias "CalcVbaGetLastError" (ByVal source As LongPtr, ByVal description As LongPtr) As Long
Private Declare PtrSafe Function NClearAnsi Lib "__DLL__" Alias "ClearAnsi" (ByVal text As LongPtr) As Long
Private Declare PtrSafe Function NAnsiText Lib "__DLL__" Alias "AnsiText" () As LongPtr
Private Type RichRecord
    tag As Byte
    value As Variant
    tail As Long
    text As String * 4
End Type
Private Declare PtrSafe Sub NAllocateRecords Lib "__DLL__" Alias "AllocateRecords" (ByRef values() As RichRecord)
Private Declare PtrSafe Function NRecordArrays Lib "__DLL__" Alias "RecordArrays" (ByRef values() As RichRecord) As LongPtr
Private Declare PtrSafe Function NArraySuite Lib "__DLL__" Alias "ArraySuite" () As LongPtr
Private Declare PtrSafe Function NChange Lib "__DLL__" Alias "ChangeRecord" (ByVal record As LongPtr) As LongPtr
Private Declare PtrSafe Sub NEcho Lib "__DLL__" Alias "EchoRecord" (ByVal output As LongPtr, ByVal record As LongPtr)
Private Declare PtrSafe Function NSuite Lib "__DLL__" Alias "ValueSuite" () As LongPtr
Private Declare PtrSafe Function NGrow Lib "__DLL__" Alias "GrowStrings" (ByRef values() As String) As Long
Private Declare PtrSafe Sub NAllocate Lib "__DLL__" Alias "AllocateAnsiStrings" (ByRef values() As String)
Private Declare PtrSafe Function NShutdown Lib "__DLL__" Alias "CalcVbaShutdown" () As Long
Private Declare PtrSafe Function BStrLength Lib "oleaut32" Alias "SysStringLen" (ByVal value As LongPtr) As Long
Private Declare PtrSafe Sub FreeBStr Lib "oleaut32" Alias "SysFreeString" (ByVal value As LongPtr)
Private Declare PtrSafe Sub CopyBytes Lib "kernel32" Alias "RtlMoveMemory" (ByVal destination As LongPtr, ByVal source As LongPtr, ByVal count As LongPtr)
Private Function TakeString(ByVal pointer As LongPtr) As String
    Dim length As Long
    length = BStrLength(pointer)
    TakeString = String$(length, ChrW$(0))
    If length Then CopyBytes StrPtr(TakeString), pointer, length * 2&
    FreeBStr pointer
End Function
Public Function ProbeNativeExports() As String
    Dim grid As GridRecord, row As Long, column As Long
    For column = 3& To 4&
        For row = -1& To 1&
            grid.values(row, column) = row + column * 10&
        Next row
    Next column
    If NGrid(grid) <> 39& Or grid.values(0, 3) <> 555& Or grid.values(-1, 4) <> 39& Then Err.Raise 525&
    Dim record As RichRecord, copied As RichRecord, values() As String
    Dim description As String, suite As String, result As String, count As Long
    record.tag = 7: record.value = "7": record.tail = 9&: record.text = "q"
    NEcho VarPtr(copied), VarPtr(record)
    description = TakeString(NChange(VarPtr(record)))
    If description <> "12:abcd:10" Or record.value <> 12# Or record.tail <> 10& Or record.text <> "abcd" Or record.tag <> 7 Then Err.Raise 513&
    If copied.value <> "7" Or copied.text <> "q   " Then Err.Raise 514&
    suite = TakeString(NSuite())
    If suite <> "12:abcd:10:12:abcd:10" Then Err.Raise 515&
    result = description & "|" & CStr(copied.value) & "|" & CStr(Len(copied.text)) & "|" & CStr(Len(suite))
    NAllocate values
    count = NGrow(values)
    If values(-2) <> "left" Or values(-1) <> "right" Or values(0) <> "new" Then Err.Raise 516&
    result = result & "|" & CStr(count) & "|" & CStr(LBound(values)) & "|" & CStr(UBound(values)) & "|" & CStr(Len(values(0)))
    ReDim values(-2 To -1): values(-2) = "left": values(-1) = "right"
    result = result & "|" & CStr(NGrow(values))
    Dim ansi As String
    ansi = "initial"
    result = result & "|" & CStr(NClearAnsi(VarPtr(ansi))) & "|" & CStr(Len(ansi))
    ansi = TakeString(NAnsiText())
    If ansi <> "reply" Then Err.Raise 517&
    result = result & "|" & CStr(Len(ansi))
    Dim records() As RichRecord
    NAllocateRecords records
    description = TakeString(NRecordArrays(records))
    If description <> "7:q   :-2:0" Or records(0).value <> "7" Or records(-2).value <> "changed" Then Err.Raise 518&
    result = result & "|" & description
    Erase records
    NAllocateRecords records
    If records(-2).value <> "7" Or records(-2).text <> "q   " Then Err.Raise 520&
    description = TakeString(NArraySuite())
    If description <> "7:q   :-2:0" Then Err.Raise 519&
    If TakeString(NFixedSuite()) <> "29:41:left:42:abc:q  :record" Then Err.Raise 521&
    If NFixedIndex(-1&) <> 29& Then Err.Raise 522&
    If NFixedIndex(-2&) <> 0& Or NLastError(0^, 0^) <> 9& Then Err.Raise 523&
    If NFixedIndex(2&) <> 0& Or NLastError(0^, 0^) <> 9& Then Err.Raise 524&
    result = result & "|" & description & "|" & CStr(NShutdown())
    ProbeNativeExports = result
End Function
Public Sub Main()
    Dim result As String
    result = ProbeNativeExports()
End Sub
