Attribute VB_Name = "XlValues"
Option Explicit

Public Type XLOPER12
    payload0 As LongLong
    payload1 As LongLong
    payload2 As LongLong
    xltype As Long
    padding As Long
End Type


' Ordinary VBA, using the documented x64 Excel and Automation layouts.
Private Declare PtrSafe Function Allocate Lib "ole32" Alias "CoTaskMemAlloc" (ByVal bytes As LongPtr) As LongPtr
Private Declare PtrSafe Sub Release Lib "ole32" Alias "CoTaskMemFree" (ByVal address As LongPtr)
Private Declare PtrSafe Sub MoveBytes Lib "kernel32" Alias "RtlMoveMemory" (ByVal destination As LongPtr, ByVal source As LongPtr, ByVal bytes As LongPtr)
Private Declare PtrSafe Function CopyFromVariant Lib "oleaut32" Alias "VariantCopyInd" (ByVal destination As LongPtr, ByRef source As Variant) As Long
Private Declare PtrSafe Function ClearNativeVariant Lib "oleaut32" Alias "VariantClear" (ByVal value As LongPtr) As Long
Private Declare PtrSafe Function NewBstr Lib "oleaut32" Alias "SysAllocStringLen" (ByVal source As LongPtr, ByVal length As Long) As LongPtr
Private Declare PtrSafe Function ArrayCreate Lib "oleaut32" Alias "SafeArrayCreate" (ByVal kind As Integer, ByVal rank As Long, ByVal bounds As LongPtr) As LongPtr
Private Declare PtrSafe Function ArrayRank Lib "oleaut32" Alias "SafeArrayGetDim" (ByVal value As LongPtr) As Long
Private Declare PtrSafe Function ArrayLower Lib "oleaut32" Alias "SafeArrayGetLBound" (ByVal value As LongPtr, ByVal dimension As Long, ByRef bound As Long) As Long
Private Declare PtrSafe Function ArrayUpper Lib "oleaut32" Alias "SafeArrayGetUBound" (ByVal value As LongPtr, ByVal dimension As Long, ByRef bound As Long) As Long
Private Declare PtrSafe Function ArrayRead Lib "oleaut32" Alias "SafeArrayGetElement" (ByVal value As LongPtr, ByVal indices As LongPtr, ByVal output As LongPtr) As Long
Private Declare PtrSafe Function ArrayWrite Lib "oleaut32" Alias "SafeArrayPutElement" (ByVal value As LongPtr, ByVal indices As LongPtr, ByRef inputValue As Variant) As Long

Private Type NativeVariant
    kind As Integer
    reserved0 As Integer
    reserved1 As Integer
    reserved2 As Integer
    payload As LongLong
    tail As LongLong
End Type
Private Type MatrixSize
    rows As Long
    columns As Long
End Type
Private Type MatrixBounds
    rows As Long
    rowLower As Long
    columns As Long
    columnLower As Long
End Type
Private Type MatrixIndex
    row As Long
    column As Long
End Type
Public Type XlRectangle
    firstRow As Long
    lastRow As Long
    firstColumn As Long
    lastColumn As Long
End Type

Private buffers As Long

Public Function XlOwnedBuffers() As Long
    XlOwnedBuffers = buffers
End Function
Private Function NewBuffer(ByVal bytes As LongPtr) As LongPtr
    NewBuffer = Allocate(bytes)
    If NewBuffer <> 0^ Then buffers = buffers + 1&
End Function
Private Sub FreeBuffer(ByVal address As LongPtr)
    If address <> 0^ Then
        Release address
        buffers = buffers - 1&
    End If
End Sub
Private Function VariantKind(ByRef value As Variant) As Long
    VariantKind = VarType(value)
End Function
Private Function FromNative(ByRef value As NativeVariant) As Variant
    ' Transfer this private owner's complete native VARIANT into the fresh result.
    ' The source has numeric fields, so disabling its tag prevents a second clear.
    MoveBytes VarPtr(FromNative), VarPtr(value), 24^
    value.kind = 0
End Function
Private Function NativeError(ByVal code As Long) As Variant
    Dim value As NativeVariant
    value.kind = 10
    value.payload = CLngLng(code)
    NativeError = FromNative(value)
End Function

Public Function XlNumber(ByVal value As Double) As XLOPER12
    MoveBytes VarPtr(XlNumber), VarPtr(value), 8^
    XlNumber.xltype = 1&
End Function
Public Function XlError(ByVal code As Long) As XLOPER12
    XlError.payload0 = CLngLng(code)
    XlError.xltype = 16&
End Function
Public Function XlString(ByVal value As String) As XLOPER12
    Dim length As Long, pointer As LongPtr
    length = Len(value)
    If length > 32767& Then
        XlString = XlError(15&)
        Exit Function
    End If
    pointer = NewBuffer(CLngPtr(length) * 2^ + 2^)
    If pointer = 0^ Then
        XlString = XlError(15&)
        Exit Function
    End If
    MoveBytes pointer, VarPtr(length), 2^
    If length <> 0& Then MoveBytes pointer + 2^, StrPtr(value), CLngPtr(length) * 2^
    XlString.payload0 = pointer
    XlString.xltype = 2&
End Function

' Empty maps to nil; a missing optional Variant keeps DISP_E_PARAMNOTFOUND.
' Null and values without an Excel representation map to #VALUE!.
Public Function XlFromVariant(ByRef value As Variant) As XLOPER12
    Dim kind As Long, code As Long
    kind = VariantKind(value)
    If (kind And 8192&) <> 0& Then
        XlFromVariant = ArrayToOper(value)
    ElseIf kind = 8& Then
        XlFromVariant = XlString(CStr(value))
    ElseIf kind = 0& Then
        XlFromVariant.xltype = 256&
    ElseIf kind = 11& Then
        XlFromVariant.xltype = 4&
        If CBool(value) Then XlFromVariant.payload0 = 1^
    ElseIf kind = 10& Then
        MoveBytes VarPtr(code), VarPtr(value) + 8^, 4^
        If code = &H80020004 Then
            XlFromVariant.xltype = 128&
        Else
            code = code And 65535&
            If code >= 2000& Then
                XlFromVariant = XlError(code - 2000&)
            Else
                XlFromVariant = XlError(15&)
            End If
        End If
    ElseIf kind = 2& Or kind = 3& Or kind = 4& Or kind = 5& Or kind = 6& Or kind = 7& Or kind = 14& Or kind = 17& Or kind = 20& Then
        XlFromVariant = XlNumber(CDbl(value))
    Else
        XlFromVariant = XlError(15&)
    End If
End Function

Public Function XlToVariant(ByRef value As XLOPER12) As Variant
    Dim raw As NativeVariant, kind As Long, length As Long, status As Long, scalar As Long
    kind = value.xltype And 4095&
    If kind = 1& Then
        raw.kind = 5
        raw.payload = value.payload0
    ElseIf kind = 2048& Then
        MoveBytes VarPtr(scalar), VarPtr(value), 4^
        raw.kind = 3
        raw.payload = CLngLng(scalar)
    ElseIf kind = 4& Then
        MoveBytes VarPtr(scalar), VarPtr(value), 4^
        raw.kind = 11
        If scalar <> 0& Then raw.payload = -1^
    ElseIf kind = 16& Then
        MoveBytes VarPtr(scalar), VarPtr(value), 4^
        raw.kind = 10
        raw.payload = CLngLng(&H800A0000 Or (2000& + scalar))
    ElseIf kind = 128& Then
        raw.kind = 10
        raw.payload = CLngLng(&H80020004)
    ElseIf kind = 256& Then
        raw.kind = 0
    ElseIf kind = 2& Then
        If value.payload0 = 0^ Then
            XlToVariant = NativeError(&H800A07DF)
            Exit Function
        End If
        MoveBytes VarPtr(length), value.payload0, 2^
        If length > 32767& Then
            XlToVariant = NativeError(&H800A07DF)
            Exit Function
        End If
        raw.payload = NewBstr(value.payload0 + 2^, length)
        raw.kind = 8
        If raw.payload = 0^ And length <> 0& Then
            XlToVariant = NativeError(&H800A07DF)
            Exit Function
        End If
    ElseIf kind = 64& Then
        XlToVariant = OperToArray(value)
        Exit Function
    ElseIf kind = 8& Or kind = 1024& Then
        XlToVariant = ReferenceToVariant(value)
        Exit Function
    Else
        XlToVariant = NativeError(&H800A07DF)
        Exit Function
    End If
    XlToVariant = FromNative(raw)
End Function

Private Function ReadArrayValue(ByRef raw As NativeVariant, ByRef index As MatrixIndex) As Variant
    Dim item As NativeVariant, status As Long, element As Long
    element = raw.kind And 4095&
    If element = 12& Then
        status = ArrayRead(raw.payload, VarPtr(index), VarPtr(item))
    Else
        item.kind = CInt(element)
        status = ArrayRead(raw.payload, VarPtr(index), VarPtr(item) + 8^)
    End If
    If status <> 0& Then
        ReadArrayValue = NativeError(&H800A07DF)
        status = ClearNativeVariant(VarPtr(item))
    Else
        ReadArrayValue = FromNative(item)
    End If
End Function

Private Function ArrayToOper(ByRef value As Variant) As XLOPER12
    Dim raw As NativeVariant, rank As Long, firstRow As Long, lastRow As Long
    Dim firstColumn As Long, lastColumn As Long, size As MatrixSize, index As MatrixIndex
    Dim row As Long, column As Long, item As Variant, child As XLOPER12, status As Long, pointer As LongPtr
    status = CopyFromVariant(VarPtr(raw), value)
    rank = ArrayRank(raw.payload)
    If rank < 1& Or rank > 2& Then
        ArrayToOper = XlError(15&)
        status = ClearNativeVariant(VarPtr(raw))
        Exit Function
    End If
    status = ArrayLower(raw.payload, 1&, firstRow)
    status = ArrayUpper(raw.payload, 1&, lastRow)
    If rank = 1& Then
        index.row = firstRow
        item = ReadArrayValue(raw, index)
        If VariantKind(item) = 8& Then
            If CStr(item) = "#XlReference12" Then
                ArrayToOper = ReferenceFromArray(raw, firstRow, lastRow)
                status = ClearNativeVariant(VarPtr(raw))
                Exit Function
            End If
        End If
        size.rows = 1&: size.columns = lastRow - firstRow + 1&
    Else
        status = ArrayLower(raw.payload, 2&, firstColumn)
        status = ArrayUpper(raw.payload, 2&, lastColumn)
        size.rows = lastRow - firstRow + 1&: size.columns = lastColumn - firstColumn + 1&
    End If
    If size.rows < 1& Or size.columns < 1& Or size.rows > 1048576& Or size.columns > 16384& Then
        ArrayToOper = XlError(15&)
        status = ClearNativeVariant(VarPtr(raw))
        Exit Function
    End If
    pointer = NewBuffer(CLngPtr(size.rows) * CLngPtr(size.columns) * 32^)
    If pointer = 0^ Then
        ArrayToOper = XlError(15&)
        status = ClearNativeVariant(VarPtr(raw))
        Exit Function
    End If
    For row = 0& To size.rows - 1&
        For column = 0& To size.columns - 1&
            If rank = 1& Then
                index.row = firstRow + column
            Else
                index.row = firstRow + row: index.column = firstColumn + column
            End If
            item = ReadArrayValue(raw, index)
            If (VariantKind(item) And 8192&) <> 0& Then
                child = XlError(15&)
            Else
                child = XlFromVariant(item)
            End If
            MoveBytes pointer + (CLngPtr(row) * CLngPtr(size.columns) + CLngPtr(column)) * 32^, VarPtr(child), 32^
        Next column
    Next row
    ArrayToOper.payload0 = pointer
    MoveBytes VarPtr(ArrayToOper) + 8^, VarPtr(size), 8^
    ArrayToOper.xltype = 64&
    status = ClearNativeVariant(VarPtr(raw))
End Function

Private Function OperToArray(ByRef value As XLOPER12) As Variant
    Dim bounds As MatrixBounds, index As MatrixIndex, size As MatrixSize
    Dim raw As NativeVariant, child As XLOPER12, item As Variant, status As Long, row As Long, column As Long
    MoveBytes VarPtr(size), VarPtr(value) + 8^, 8^
    If size.rows < 1& Or size.columns < 1& Or size.rows > 1048576& Or size.columns > 16384& Or value.payload0 = 0^ Then
        OperToArray = NativeError(&H800A07DF)
        Exit Function
    End If
    bounds.rows = size.rows: bounds.columns = size.columns
    bounds.rowLower = 1&: bounds.columnLower = 1&
    raw.kind = 8204
    raw.payload = ArrayCreate(12, 2&, VarPtr(bounds))
    If raw.payload = 0^ Then
        OperToArray = NativeError(&H800A07DF)
        Exit Function
    End If
    For row = 1& To size.rows
        For column = 1& To size.columns
            index.row = row: index.column = column
            MoveBytes VarPtr(child), value.payload0 + (CLngPtr(index.row - 1&) * CLngPtr(size.columns) + CLngPtr(index.column - 1&)) * 32^, 32^
            item = XlToVariant(child)
            status = ArrayWrite(raw.payload, VarPtr(index), item)
            If status <> 0& Then
                status = ClearNativeVariant(VarPtr(raw))
                OperToArray = NativeError(&H800A07DF)
                Exit Function
            End If
        Next column
    Next row
    OperToArray = FromNative(raw)
End Function

Public Function XlReferenceCount(ByRef value As XLOPER12) As Long
    Dim kind As Long
    kind = value.xltype And 4095&
    If kind = 1024& Then
        XlReferenceCount = 1&
    ElseIf kind = 8& And value.payload0 <> 0^ Then
        MoveBytes VarPtr(XlReferenceCount), value.payload0, 2^
    End If
End Function
Public Function XlReferenceArea(ByRef value As XLOPER12, ByVal index As Long) As XlRectangle
    Dim kind As Long
    If index < 0& Or index >= XlReferenceCount(value) Then Err.Raise 9&
    kind = value.xltype And 4095&
    If kind = 1024& Then
        MoveBytes VarPtr(XlReferenceArea), VarPtr(value) + 4^, 16^
    Else
        MoveBytes VarPtr(XlReferenceArea), value.payload0 + 4^ + CLngPtr(index) * 16^, 16^
    End If
End Function

' Reference values use a reserved tagged Variant vector:
' marker, type, sheet ID, area count, then four zero-based coordinates per area.
' Keeping the sheet ID as LongLong preserves its exact pointer-sized value.
Private Function ReferenceToVariant(ByRef value As XLOPER12) As Variant
    Dim fields() As Variant, count As Long, area As Long, rectangle As XlRectangle, kind As Long
    kind = value.xltype And 4095&
    If kind = 1024& Then
        count = 1&
    ElseIf value.payload0 <> 0^ Then
        MoveBytes VarPtr(count), value.payload0, 2^
    End If
    If count < 1& Then
        ReferenceToVariant = NativeError(&H800A07DF)
        Exit Function
    End If
    ReDim fields(0 To 3& + count * 4&)
    fields(0) = "#XlReference12": fields(1) = kind
    If kind = 8& Then fields(2) = value.payload1 Else fields(2) = 0^
    fields(3) = count
    For area = 0& To count - 1&
        If kind = 1024& Then
            MoveBytes VarPtr(rectangle), VarPtr(value) + 4^, 16^
        Else
            MoveBytes VarPtr(rectangle), value.payload0 + 4^ + CLngPtr(area) * 16^, 16^
        End If
        fields(4& + area * 4&) = rectangle.firstRow
        fields(5& + area * 4&) = rectangle.lastRow
        fields(6& + area * 4&) = rectangle.firstColumn
        fields(7& + area * 4&) = rectangle.lastColumn
    Next area
    ReferenceToVariant = fields
End Function

Private Function ReferenceFromArray(ByRef raw As NativeVariant, ByVal first As Long, ByVal last As Long) As XLOPER12
    Dim index As MatrixIndex, item As Variant, kind As Long, count As Long, area As Long
    Dim rectangle As XlRectangle, pointer As LongPtr, sheet As LongLong, singleCount As Long
    ReferenceFromArray = XlError(15&)
    If last - first < 7& Then Exit Function
    index.row = first + 1&: item = ReadArrayValue(raw, index): kind = CLng(item)
    index.row = first + 2&: item = ReadArrayValue(raw, index): sheet = CLngLng(item)
    index.row = first + 3&: item = ReadArrayValue(raw, index): count = CLng(item)
    If count < 1& Or count > 65535& Or last - first <> 3& + count * 4& Then Exit Function
    If kind <> 8& And kind <> 1024& Then Exit Function
    If kind = 1024& And count <> 1& Then Exit Function
    If kind = 8& Then
        pointer = NewBuffer(4^ + CLngPtr(count) * 16^)
        If pointer = 0^ Then Exit Function
        MoveBytes pointer, VarPtr(count), 4^
    End If
    For area = 0& To count - 1&
        index.row = first + 4& + area * 4&: item = ReadArrayValue(raw, index): rectangle.firstRow = CLng(item)
        index.row = index.row + 1&: item = ReadArrayValue(raw, index): rectangle.lastRow = CLng(item)
        index.row = index.row + 1&: item = ReadArrayValue(raw, index): rectangle.firstColumn = CLng(item)
        index.row = index.row + 1&: item = ReadArrayValue(raw, index): rectangle.lastColumn = CLng(item)
        If rectangle.firstRow < 0& Or rectangle.lastRow < rectangle.firstRow Or rectangle.lastRow > 1048575& Or rectangle.firstColumn < 0& Or rectangle.lastColumn < rectangle.firstColumn Or rectangle.lastColumn > 16383& Then
            FreeBuffer pointer
            Exit Function
        End If
        If kind = 8& Then
            MoveBytes pointer + 4^ + CLngPtr(area) * 16^, VarPtr(rectangle), 16^
        Else
            MoveBytes VarPtr(ReferenceFromArray) + 4^, VarPtr(rectangle), 16^
        End If
    Next area
    If kind = 8& Then
        ReferenceFromArray.payload0 = pointer
        ReferenceFromArray.payload1 = sheet
    Else
        singleCount = 1&
        MoveBytes VarPtr(ReferenceFromArray), VarPtr(singleCount), 4^
    End If
    ReferenceFromArray.xltype = kind
End Function

' Only pass values allocated by this library; Excel-owned callback values use xlFree.
Public Sub XlFreeValue(ByRef value As XLOPER12)
    Dim kind As Long, size As MatrixSize, index As LongPtr, child As XLOPER12, cleared As XLOPER12
    kind = value.xltype And 4095&
    If kind = 64& Then
        MoveBytes VarPtr(size), VarPtr(value) + 8^, 8^
        For index = 0^ To CLngPtr(size.rows) * CLngPtr(size.columns) - 1^
            MoveBytes VarPtr(child), value.payload0 + index * 32^, 32^
            XlFreeValue child
        Next index
        FreeBuffer value.payload0
    ElseIf kind = 2& Or kind = 8& Then
        FreeBuffer value.payload0
    End If
    value = cleared
End Sub
Public Function XlReturn(ByRef value As XLOPER12) As LongPtr
    XlReturn = NewBuffer(32^)
    If XlReturn = 0^ Then
        XlFreeValue value
        Exit Function
    End If
    value.xltype = value.xltype Or 16384&
    MoveBytes XlReturn, VarPtr(value), 32^
End Function
Public Sub XlReleaseReturn(ByVal address As LongPtr)
    Dim value As XLOPER12
    If address = 0^ Then Exit Sub
    MoveBytes VarPtr(value), address, 32^
    XlFreeValue value
    FreeBuffer address
End Sub
