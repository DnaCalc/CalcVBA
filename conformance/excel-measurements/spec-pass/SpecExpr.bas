Attribute VB_Name = "SpecExpr"
Option Explicit
' MS-VBAL §5.6 expressions (names, member access, index, !, With, arguments) and §5.5.2
' Set-coercion: open points from the dc-ez3rd code inspection. Needs the class SpecK.
Private Const n As Long = 5
Private Type Pt
    x As Long
End Type

Private Function Sh(ByVal v As Variant) As String
    If IsObject(v) Then Sh = "obj " & TypeName(v): Exit Function
    If IsArray(v) Then Sh = TypeName(v) & "[" & LBound(v) & "," & UBound(v) & "]": Exit Function
    Select Case VarType(v)
        Case vbNull: Sh = "Null"
        Case vbEmpty: Sh = "Empty"
        Case vbString: Sh = "String:[" & v & "]"
        Case Else: Sh = TypeName(v) & ":" & CStr(v)
    End Select
End Function

Private Function ArrF() As Variant
    ArrF = Array(7, 8, 9)
End Function

Private Function Shadow(n As Long) As Long
    Shadow = n
End Function

Public Function TakeAny(o As Object) As String
    TakeAny = TypeName(o)
End Function

Private Function TakeK(o As SpecK) As String
    TakeK = TypeName(o)
End Function

Private Function SumArr(a() As Long) As Long
    SumArr = UBound(a) + 100
End Function

Private Sub IncL(x As Long)
    x = x + 1
End Sub

Private Function PA(ParamArray r() As Variant) As String
    PA = VarType(r(1)) & " " & IsMissing(r(1))
End Function

Private Function T(ByVal k As String) As String
    Dim v As Variant, vv As Variant, c As Collection, kk As SpecK, o As Object, lk As Long
    Dim cc As Collection, pts(1) As Pt, i As Long, la(2) As Long
    On Error GoTo fail
    Select Case k
        Case "bangColl": Set c = New Collection: c.Add 7, "k": v = c!k
        Case "bangWith": Set c = New Collection: c.Add 8, "k": With c: v = !k: End With
        Case "idxZeroParamFn": Set kk = New SpecK: v = kk.Arr(1)
        Case "idxZeroParamProp": Set kk = New SpecK: v = kk.Items(2)
        Case "idxZeroParamMod": v = ArrF(1)
        Case "paramVsConst": v = Shadow(9)
        Case "parenObjArg": Set kk = New SpecK: v = TakeK((kk))
        Case "typeOfObject": Set kk = New SpecK: v = TypeOf kk Is Object
        Case "typeOfObjectColl": Set c = New Collection: v = TypeOf c Is Object
        Case "withUdtIndex": i = 0: With pts(i): i = 1: .x = 5: End With: v = pts(0).x & "," & pts(1).x
        Case "leftException": Set kk = New SpecK: v = kk.UseLeft
        Case "emptyIndexArr": v = SumArr(la())
        Case "lateNamedByRef": Set o = New SpecK: lk = 1: o.Bump n:=lk: v = lk
        Case "latePosByRef": Set o = New SpecK: lk = 1: o.Bump lk: v = lk
        Case "paOmitted": v = PA(1, , 3)
        Case "byRefField": Set kk = New SpecK: kk.n = 1: IncL kk.n: v = kk.n
        Case "typeOfEmpty": vv = Empty: v = TypeOf vv Is Collection
        Case "typeOfNothing": Set o = Nothing: v = TypeOf o Is Collection
        Case "passMe": Set kk = New SpecK: v = kk.PassMe
        Case "setCollToClass": Set vv = New Collection: Set kk = vv: v = TypeName(kk)
        Case "setClassToColl": Set vv = New SpecK: Set cc = vv: v = TypeName(cc)
        Case "setClassToObj": Set vv = New SpecK: Set o = vv: v = TypeName(o)
        Case Else: v = "unknown case"
    End Select
    T = k & "=" & Sh(v): Exit Function
fail:
    T = k & "=ERR" & Err.Number
End Function

Public Function Run() As String
    Dim keys As Variant, i As Long, s As String
    keys = Split("bangColl bangWith idxZeroParamFn idxZeroParamProp idxZeroParamMod paramVsConst parenObjArg typeOfObject typeOfObjectColl withUdtIndex leftException emptyIndexArr lateNamedByRef latePosByRef paOmitted byRefField typeOfEmpty typeOfNothing passMe setCollToClass setClassToColl setClassToObj", " ")
    For i = 0 To UBound(keys)
        s = s & T(keys(i)) & vbLf
    Next
    Run = s
End Function
