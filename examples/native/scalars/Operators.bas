Attribute VB_Name = "Operators"
Option Explicit

Public Function IntegerDivide(ByVal a As Long, ByVal b As Long) As Long
    IntegerDivide = a \ b
End Function

Public Function IntegerMod(ByVal a As Long, ByVal b As Long) As Long
    IntegerMod = a Mod b
End Function

Public Function BooleanAnd(ByVal a As Boolean, ByVal b As Boolean) As Boolean
    BooleanAnd = a And b
End Function

Public Function BooleanOr(ByVal a As Boolean, ByVal b As Boolean) As Boolean
    BooleanOr = a Or b
End Function

Public Function BooleanXor(ByVal a As Boolean, ByVal b As Boolean) As Boolean
    BooleanXor = a Xor b
End Function

Public Function BooleanEqv(ByVal a As Boolean, ByVal b As Boolean) As Boolean
    BooleanEqv = a Eqv b
End Function

Public Function BooleanImp(ByVal a As Boolean, ByVal b As Boolean) As Boolean
    BooleanImp = a Imp b
End Function

Public Function BooleanNot(ByVal a As Boolean) As Boolean
    BooleanNot = Not a
End Function

Public Function LongAnd(ByVal a As Long, ByVal b As Long) As Long
    LongAnd = a And b
End Function

Public Function LongOr(ByVal a As Long, ByVal b As Long) As Long
    LongOr = a Or b
End Function

Public Function LongXor(ByVal a As Long, ByVal b As Long) As Long
    LongXor = a Xor b
End Function

Public Function LongEqv(ByVal a As Long, ByVal b As Long) As Long
    LongEqv = a Eqv b
End Function

Public Function LongImp(ByVal a As Long, ByVal b As Long) As Long
    LongImp = a Imp b
End Function

Public Function LongNot(ByVal a As Long) As Long
    LongNot = Not a
End Function
