' Automation VARIANT layout and ByRef Variant binding probes (dc-yqq0.1).
' One `id=value` line per case; each case traps its own error as `ERR<n>`.
' Values are address-independent so Excel, VM3 and the JIT can be compared directly.
Option Explicit

Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal dst As LongPtr, ByVal src As LongPtr, ByVal n As LongPtr)
Private Declare PtrSafe Function VarPtrArray Lib "VBE7" Alias "VarPtr" (ByRef a() As Any) As LongPtr

Private Type TSimple
    n As Long
    d As Double
End Type

Private Type TWithVar
    a As Long
    v As Variant
    b As Integer
End Type

Private Type TNested
    x As Byte
    t As TWithVar
    arr(1 To 2) As Variant
    s As String
    f As String * 3
End Type

Private Type TRec
    a As Long
    s As String
End Type

Private Type TObjField
    a As Byte
    o As Object
    b As Byte
End Type

Private Type TDynField
    a As Byte
    d() As Long
    b As Byte
End Type

Private Type TFixField
    a As Byte
    f(0 To 1, 0 To 2) As Long
End Type

Private Type TFixedStr
    a As Byte
    s As String * 3
    b As Byte
    l As Long
End Type

Private Type TClassField
    a As Byte
    c As Collection
End Type

Private Function HexAt(ByVal p As LongPtr, ByVal n As Long) As String
    Dim b As Byte, i As Long, out As String
    For i = 0 To n - 1
        CopyMemory VarPtr(b), p + i, 1
        out = out & Right$("0" & Hex$(b), 2)
    Next i
    HexAt = out
End Function

Private Function VtAt(ByVal p As LongPtr) As String
    Dim vt As Integer
    CopyMemory VarPtr(vt), p, 2
    VtAt = "&H" & Hex$(vt)
End Function

Private Function PtrAt(ByVal p As LongPtr) As LongPtr
    Dim x As LongPtr
    CopyMemory VarPtr(x), p, LenB(x)
    PtrAt = x
End Function

Private Function VtOf(v As Variant) As String
    VtOf = VtAt(VarPtr(v))
End Function

Private Function ArrayInfo(v As Variant) As String
    Dim psa As LongPtr, cDims As Integer, fFeatures As Integer, cbElements As Long, vtPrefix As Long
    psa = PtrAt(VarPtr(v) + 8)
    If psa = 0 Then ArrayInfo = "vt=" & VtAt(VarPtr(v)) & " null": Exit Function
    CopyMemory VarPtr(cDims), psa, 2
    CopyMemory VarPtr(fFeatures), psa + 2, 2
    CopyMemory VarPtr(cbElements), psa + 4, 4
    CopyMemory VarPtr(vtPrefix), psa - 4, 4
    ArrayInfo = "vt=" & VtAt(VarPtr(v)) & " cDims=" & cDims & " fFeatures=&H" & Hex$(fFeatures) & " cbElements=" & cbElements & " vtPrefix=&H" & Hex$(vtPrefix)
End Function

Private Function Binding(v As Variant, ByVal target As LongPtr) As String
    Binding = VtAt(VarPtr(v)) & " refIsTarget=" & (PtrAt(VarPtr(v) + 8) = target) & " cellIsTarget=" & (VarPtr(v) = target) & " TypeName=" & TypeName(v) & " VarType=" & VarType(v)
End Function

Private Function RefHolds(v As Variant, ByVal expected As LongPtr) As Boolean
    RefHolds = (PtrAt(PtrAt(VarPtr(v) + 8)) = expected)
End Function

Private Sub AssignTo(v As Variant, ByVal value As Variant)
    v = value
End Sub

Private Function InnerCell(v As Variant) As LongPtr
    InnerCell = VarPtr(v)
End Function

Private Function OuterCell(v As Variant, ByRef inner As LongPtr) As LongPtr
    inner = InnerCell(v)
    OuterCell = VarPtr(v)
End Function

Private Function MissingInfo(Optional x As Variant) As String
    MissingInfo = VtAt(VarPtr(x)) & " scode=" & HexAt(VarPtr(x) + 8, 4) & " IsMissing=" & IsMissing(x)
End Function

Private Function ByValVt(ByVal v As Variant) As String
    ByValVt = VtAt(VarPtr(v))
End Function

Private Sub SetNothing(v As Variant)
    Set v = Nothing
End Sub

Private Sub SetNew(v As Variant)
    Set v = New Collection
End Sub

' ---- strides ----
Private Function CStrideVariant() As String
    On Error GoTo E
    Dim a(0 To 3) As Variant
    CStrideVariant = CStr(VarPtr(a(1)) - VarPtr(a(0)))
    Exit Function
E:  CStrideVariant = "ERR" & Err.Number
End Function

Private Function CStrideLong() As String
    On Error GoTo E
    Dim a(0 To 3) As Long
    CStrideLong = CStr(VarPtr(a(1)) - VarPtr(a(0))) & " a3-a0=" & (VarPtr(a(3)) - VarPtr(a(0)))
    Exit Function
E:  CStrideLong = "ERR" & Err.Number
End Function

Private Function CStrideString() As String
    On Error GoTo E
    Dim a(0 To 1) As String
    CStrideString = CStr(VarPtr(a(1)) - VarPtr(a(0)))
    Exit Function
E:  CStrideString = "ERR" & Err.Number
End Function

Private Function CStrideUdt() As String
    On Error GoTo E
    Dim a(0 To 1) As TWithVar
    CStrideUdt = CStr(VarPtr(a(1)) - VarPtr(a(0)))
    Exit Function
E:  CStrideUdt = "ERR" & Err.Number
End Function

' ---- cell bytes and tags ----
Private Function CBytes(ByVal which As String) As String
    On Error GoTo E
    Dim v As Variant
    Select Case which
        Case "Long": v = CLng(5)
        Case "Double": v = CDbl(1.5)
        Case "Decimal": v = CDec("12345678901234567890.5")
        Case "Currency": v = CCur(1.5)
        Case "Date": v = CDate(1)
        Case "LongLong": v = CLngLng(7)
    End Select
    CBytes = HexAt(VarPtr(v), 8) & " " & HexAt(VarPtr(v) + 8, 8)
    Exit Function
E:  CBytes = "ERR" & Err.Number
End Function

Private Function CVt(ByVal which As String) As String
    On Error GoTo E
    Dim v As Variant
    Select Case which
        Case "Empty": v = Empty
        Case "Null": v = Null
        Case "Integer": v = CInt(5)
        Case "Long": v = CLng(5)
        Case "Single": v = CSng(1.5)
        Case "Double": v = CDbl(1.5)
        Case "Currency": v = CCur(1.5)
        Case "Date": v = CDate(1)
        Case "String": v = "abc"
        Case "Boolean": v = True
        Case "Byte": v = CByte(7)
        Case "LongLong": v = CLngLng(7)
        Case "Decimal": v = CDec(1)
        Case "Error": v = CVErr(2042)
        Case "Nothing": Set v = Nothing
        Case "Collection": Set v = New Collection
    End Select
    CVt = VtOf(v)
    Exit Function
E:  CVt = "ERR" & Err.Number
End Function

Private Function CPayload(ByVal which As String) As String
    On Error GoTo E
    Dim v As Variant
    Select Case which
        Case "Boolean": v = True: CPayload = HexAt(VarPtr(v) + 8, 2)
        Case "Error": v = CVErr(2042): CPayload = HexAt(VarPtr(v) + 8, 4)
        Case "Nothing": Set v = Nothing: CPayload = HexAt(VarPtr(v) + 8, 8)
        Case "Integer": v = CInt(-2): CPayload = HexAt(VarPtr(v) + 8, 2)
    End Select
    Exit Function
E:  CPayload = "ERR" & Err.Number
End Function

Private Function CMissing() As String
    On Error GoTo E
    CMissing = MissingInfo()
    Exit Function
E:  CMissing = "ERR" & Err.Number
End Function

' ---- SAFEARRAYs ----
Private Function CArray(ByVal which As String) As String
    On Error GoTo E
    Dim v As Variant, la(0 To 3) As Long, va(0 To 3) As Variant, sa(0 To 1) As String
    Dim dla() As Long, dva() As Variant, da(0 To 1) As Double
    Select Case which
        Case "ArrayFn": v = Array(1, 2)
        Case "Long": v = la
        Case "Variant": v = va
        Case "String": v = sa
        Case "Double": v = da
        Case "UnallocatedLong": v = dla
        Case "DynamicVariant": ReDim dva(0 To 2): v = dva
    End Select
    CArray = ArrayInfo(v)
    Exit Function
E:  CArray = "ERR" & Err.Number
End Function

' ---- ByRef Variant parameter binding ----
Private Function CBind(ByVal which As String) As String
    On Error GoTo E
    Dim n As Long, s As String, d As Double, b As Boolean, o As Collection, u As Object, v As Variant
    Dim la(0 To 1) As Long, va(0 To 1) As Variant, t As TRec
    n = 1: s = "a": d = 1: b = True: v = 5
    Select Case which
        Case "Long": CBind = Binding(n, VarPtr(n))
        Case "String": CBind = Binding(s, VarPtr(s))
        Case "Double": CBind = Binding(d, VarPtr(d))
        Case "Boolean": CBind = Binding(b, VarPtr(b))
        Case "Variant": CBind = Binding(v, VarPtr(v))
        Case "LongElement": CBind = Binding(la(1), VarPtr(la(1)))
        Case "VariantElement": CBind = Binding(va(1), VarPtr(va(1)))
        Case "UdtLongField": CBind = Binding(t.a, VarPtr(t.a))
        Case "UdtStringField": CBind = Binding(t.s, VarPtr(t.s))
        Case "Expression": CBind = Binding(n + 0, VarPtr(n))
        Case "Parenthesised": CBind = Binding((n), VarPtr(n))
        Case "LongArray": CBind = Binding(la, 0)
        Case "ObjectVar": Set u = New Collection: CBind = Binding(u, VarPtr(u))
        Case "TypedObjectVar": Set o = New Collection: CBind = Binding(o, VarPtr(o)) & " *ref=ObjPtr:" & RefHolds(o, ObjPtr(o))
        Case "ByValLong": CBind = ByValVt(n)
    End Select
    Exit Function
E:  CBind = "ERR" & Err.Number
End Function

Private Function CAssign(ByVal which As String) As String
    On Error GoTo E
    Dim n As Long, s As String, v As Variant, t As TRec, la(0 To 1) As Long
    n = 1
    Select Case which
        Case "LongFromString": AssignTo n, "12": CAssign = CStr(n)
        Case "LongFrom2_5": AssignTo n, 2.5: CAssign = CStr(n)
        Case "LongFrom3_5": AssignTo n, 3.5: CAssign = CStr(n)
        Case "LongFromText": AssignTo n, "abc": CAssign = CStr(n)
        Case "LongFromNull": AssignTo n, Null: CAssign = CStr(n)
        Case "LongOverflow": AssignTo n, 1E+20: CAssign = CStr(n)
        Case "StringFrom42": s = "a": AssignTo s, 42: CAssign = s
        Case "VariantFromText": v = 1: AssignTo v, "text": CAssign = TypeName(v) & ":" & v
        Case "UdtField": t.a = 1: AssignTo t.a, 9: CAssign = CStr(t.a)
        Case "LongElement": la(1) = 1: AssignTo la(1), 8: CAssign = CStr(la(1))
    End Select
    Exit Function
E:  CAssign = "ERR" & Err.Number & " n=" & n
End Function

Private Function CNested(ByVal which As String) As String
    On Error GoTo E
    Dim n As Long, v As Variant, oc As LongPtr, ic As LongPtr
    Select Case which
        Case "Long": n = 1: oc = OuterCell(n, ic): CNested = "outer=inner:" & (oc = ic)
        Case "Variant": v = 1: oc = OuterCell(v, ic): CNested = "outer=inner:" & (oc = ic) & " outer=caller:" & (oc = VarPtr(v))
    End Select
    Exit Function
E:  CNested = "ERR" & Err.Number
End Function

Private Function CObjWriteBack(ByVal which As String) As String
    On Error GoTo E
    Dim o As Collection, first As LongPtr
    Set o = New Collection
    Select Case which
        Case "SetNothing": SetNothing o: CObjWriteBack = CStr(o Is Nothing)
        Case "SetNew": first = ObjPtr(o): SetNew o: CObjWriteBack = CStr(ObjPtr(o) <> first)
    End Select
    Exit Function
E:  CObjWriteBack = "ERR" & Err.Number
End Function

' ---- memory copies and writes through VarPtr ----
Private Function CCopy(ByVal which As String) As String
    On Error GoTo E
    Dim v As Variant, w As Variant, z As Variant
    Select Case which
        Case "Double24": v = 12345.678#: CopyMemory VarPtr(w), VarPtr(v), 24: CCopy = TypeName(w) & ":" & CStr(w)
        Case "Decimal24": v = CDec("1.25"): CopyMemory VarPtr(w), VarPtr(v), 24: CCopy = TypeName(w) & ":" & CStr(w)
        Case "Decimal16": v = CDec("1.25"): CopyMemory VarPtr(w), VarPtr(v), 16: CCopy = TypeName(w) & ":" & CStr(w)
        Case "String24"
            v = "shared": CopyMemory VarPtr(w), VarPtr(v), 24
            CCopy = TypeName(w) & ":" & w & " sameBSTR=" & (StrPtr(w) = StrPtr(v))
            CopyMemory VarPtr(w), VarPtr(z), 24
    End Select
    Exit Function
E:  CCopy = "ERR" & Err.Number
End Function

Private Function CWrite(ByVal which As String) As String
    On Error GoTo E
    Dim n As Long, src As Long, v As Variant, a(0 To 3) As Variant, t As TWithVar, la(0 To 3) As Long, p As LongPtr
    Dim four(0 To 3) As Long
    src = 77: n = 1
    Select Case which
        Case "Long": CopyMemory VarPtr(n), VarPtr(src), 4: CWrite = CStr(n)
        Case "StoredPointerTwice"
            p = VarPtr(n)
            CopyMemory p, VarPtr(src), 4
            src = 78: CopyMemory p, VarPtr(src), 4
            CWrite = CStr(n)
        Case "VariantPayload": v = CLng(1): CopyMemory VarPtr(v) + 8, VarPtr(src), 4: CWrite = CStr(v)
        Case "VariantElement": a(2) = CLng(1): CopyMemory VarPtr(a(2)) + 8, VarPtr(src), 4: CWrite = CStr(a(2))
        Case "UdtField": t.a = 1: CopyMemory VarPtr(t.a), VarPtr(src), 4: CWrite = CStr(t.a)
        Case "LongArrayBlock"
            four(0) = 1: four(1) = 2: four(2) = 3: four(3) = 4
            CopyMemory VarPtr(la(0)), VarPtr(four(0)), 16
            CWrite = la(0) & "," & la(1) & "," & la(2) & "," & la(3)
    End Select
    Exit Function
E:  CWrite = "ERR" & Err.Number
End Function

Private Function CBstrPrefix() As String
    On Error GoTo E
    Dim s As String, plen As Long
    s = "hello": CopyMemory VarPtr(plen), StrPtr(s) - 4, 4
    CBstrPrefix = plen & " LenB=" & LenB(s) & " nullStrPtr=" & StrPtr(vbNullString)
    Exit Function
E:  CBstrPrefix = "ERR" & Err.Number
End Function

' ---- UDT sizes and offsets ----
Private Function CUdt(ByVal which As String) As String
    On Error GoTo E
    Dim ts As TSimple, t As TWithVar, tn As TNested
    Select Case which
        Case "TSimple": CUdt = "Len=" & Len(ts) & " LenB=" & LenB(ts)
        Case "TWithVar": CUdt = "Len=" & Len(t) & " LenB=" & LenB(t) & " v=" & (VarPtr(t.v) - VarPtr(t)) & " b=" & (VarPtr(t.b) - VarPtr(t))
        Case "TNested": CUdt = "Len=" & Len(tn) & " LenB=" & LenB(tn) & " t=" & (VarPtr(tn.t) - VarPtr(tn)) & " arr1=" & (VarPtr(tn.arr(1)) - VarPtr(tn)) & " arr2=" & (VarPtr(tn.arr(2)) - VarPtr(tn)) & " s=" & (VarPtr(tn.s) - VarPtr(tn))
        Case "FieldVariantVt": t.v = "x": CUdt = VtAt(VarPtr(t.v))
    End Select
    Exit Function
E:  CUdt = "ERR" & Err.Number
End Function

Private Function UdtArrayInfo(ByVal pp As LongPtr) As String
    Dim psa As LongPtr, fFeatures As Integer, cbElements As Long, vtPrefix As Long, recInfo As LongPtr
    psa = PtrAt(pp)
    If psa = 0 Then UdtArrayInfo = "null": Exit Function
    CopyMemory VarPtr(fFeatures), psa + 2, 2
    CopyMemory VarPtr(cbElements), psa + 4, 4
    CopyMemory VarPtr(vtPrefix), psa - 4, 4
    CopyMemory VarPtr(recInfo), psa - 8, LenB(recInfo)
    UdtArrayInfo = "fFeatures=&H" & Hex$(fFeatures) & " cbElements=" & cbElements & " vtPrefix=&H" & Hex$(vtPrefix) & " recInfo=" & (recInfo <> 0)
End Function

Private Function CUdtMem(ByVal which As String) As String
    On Error GoTo E
    Dim t1 As TObjField, t2 As TDynField, t3 As TFixField, t4 As TFixedStr, t5 As TClassField
    Dim ra() As TRec, fa(0 To 1) As TRec
    Select Case which
        Case "ObjectField": CUdtMem = "Len=" & Len(t1) & " LenB=" & LenB(t1) & " b=" & (VarPtr(t1.b) - VarPtr(t1))
        Case "ObjectFieldHoldsPtr": Set t1.o = New Collection: CUdtMem = CStr(PtrAt(VarPtr(t1) + 8) = ObjPtr(t1.o))
        Case "ClassField": CUdtMem = "Len=" & Len(t5) & " LenB=" & LenB(t5)
        Case "DynArrayField": CUdtMem = "Len=" & Len(t2) & " LenB=" & LenB(t2) & " b=" & (VarPtr(t2.b) - VarPtr(t2))
        Case "DynArrayFieldHoldsPtr": ReDim t2.d(0 To 3): CUdtMem = CStr(PtrAt(VarPtr(t2) + 8) <> 0)
        Case "FixedArrayOrder": CUdtMem = "Len=" & Len(t3) & " LenB=" & LenB(t3) & " f00=" & (VarPtr(t3.f(0, 0)) - VarPtr(t3)) & " f10=" & (VarPtr(t3.f(1, 0)) - VarPtr(t3)) & " f01=" & (VarPtr(t3.f(0, 1)) - VarPtr(t3))
        Case "FixedString": t4.s = "xyz": CUdtMem = "Len=" & Len(t4) & " LenB=" & LenB(t4) & " b=" & (VarPtr(t4.b) - VarPtr(t4)) & " l=" & (VarPtr(t4.l) - VarPtr(t4)) & " bytes=" & HexAt(VarPtr(t4), 8)
        Case "DynUdtArray": ReDim ra(0 To 2): CUdtMem = UdtArrayInfo(VarPtrArray(ra)) & " stride=" & (VarPtr(ra(1)) - VarPtr(ra(0)))
        Case "FixedUdtArray": CUdtMem = UdtArrayInfo(VarPtrArray(fa))
    End Select
    Exit Function
E:  CUdtMem = "ERR" & Err.Number
End Function

Private Function CForEach() As String
    On Error GoTo E
    Dim la(0 To 1) As Long, item As Variant
    For Each item In la
        CForEach = VtOf(item)
        Exit For
    Next item
    Exit Function
E:  CForEach = "ERR" & Err.Number
End Function

Private Sub Add(ByRef out As String, ByVal id As String, ByVal value As String)
    out = out & id & "=" & value & vbLf
End Sub

Public Function RunAll() As String
    Dim out As String, k As Variant
    Add out, "STRIDE.Variant", CStrideVariant()
    Add out, "STRIDE.Long", CStrideLong()
    Add out, "STRIDE.String", CStrideString()
    Add out, "STRIDE.UdtWithVariant", CStrideUdt()
    For Each k In Array("Long", "Double", "Decimal", "Currency", "Date", "LongLong")
        Add out, "BYTES." & k, CBytes(k)
    Next k
    For Each k In Array("Empty", "Null", "Integer", "Long", "Single", "Double", "Currency", "Date", "String", "Boolean", "Byte", "LongLong", "Decimal", "Error", "Nothing", "Collection")
        Add out, "VT." & k, CVt(k)
    Next k
    For Each k In Array("Boolean", "Error", "Nothing", "Integer")
        Add out, "PAYLOAD." & k, CPayload(k)
    Next k
    Add out, "MISSING", CMissing()
    For Each k In Array("ArrayFn", "Long", "Variant", "String", "Double", "UnallocatedLong", "DynamicVariant")
        Add out, "ARRAY." & k, CArray(k)
    Next k
    For Each k In Array("Long", "String", "Double", "Boolean", "Variant", "LongElement", "VariantElement", "UdtLongField", "UdtStringField", "Expression", "Parenthesised", "LongArray", "ObjectVar", "TypedObjectVar", "ByValLong")
        Add out, "BIND." & k, CBind(k)
    Next k
    For Each k In Array("LongFromString", "LongFrom2_5", "LongFrom3_5", "LongFromText", "LongFromNull", "LongOverflow", "StringFrom42", "VariantFromText", "UdtField", "LongElement")
        Add out, "ASSIGN." & k, CAssign(k)
    Next k
    Add out, "NESTED.Long", CNested("Long")
    Add out, "NESTED.Variant", CNested("Variant")
    Add out, "OBJBACK.SetNothing", CObjWriteBack("SetNothing")
    Add out, "OBJBACK.SetNew", CObjWriteBack("SetNew")
    For Each k In Array("Double24", "Decimal24", "Decimal16", "String24")
        Add out, "COPY." & k, CCopy(k)
    Next k
    For Each k In Array("Long", "StoredPointerTwice", "VariantPayload", "VariantElement", "UdtField", "LongArrayBlock")
        Add out, "WRITE." & k, CWrite(k)
    Next k
    Add out, "BSTR.Prefix", CBstrPrefix()
    For Each k In Array("TSimple", "TWithVar", "TNested", "FieldVariantVt")
        Add out, "UDT." & k, CUdt(k)
    Next k
    Add out, "FOREACH.LongItem", CForEach()
    For Each k In Array("ObjectField", "ObjectFieldHoldsPtr", "ClassField", "DynArrayField", "DynArrayFieldHoldsPtr", "FixedArrayOrder", "FixedString", "DynUdtArray", "FixedUdtArray")
        Add out, "UDTMEM." & k, CUdtMem(k)
    Next k
    RunAll = out
End Function
