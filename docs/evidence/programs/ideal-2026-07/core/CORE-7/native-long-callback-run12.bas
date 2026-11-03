Option Explicit
Private Declare PtrSafe Function MutateThenCall Lib "__LIBRARY__" (ByRef value As Long, ByVal callback As LongPtr) As Long
Private aliasValue As Long
Public Function CallbackValue() As Long
    aliasValue = aliasValue + 1
    CallbackValue = aliasValue
End Function
Public Function RunProbe() As String
    Dim observed As Long
    aliasValue = 2
    observed = MutateThenCall(aliasValue, AddressOf CallbackValue)
    RunProbe = observed & ":" & aliasValue
End Function
