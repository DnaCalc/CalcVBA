Option Explicit
Private Declare PtrSafe Function MutateDoubleThenCall Lib "__LIBRARY__" (ByRef value As Double, ByRef sameValue As Double, ByVal callback As LongPtr) As Double
Private Declare PtrSafe Function MutateBooleanThenCall Lib "__LIBRARY__" (ByRef value As Boolean, ByRef sameValue As Boolean, ByVal callback As LongPtr) As Boolean
__CALLBACK_BODIES__
Public Function RunProbe() As String
    Dim observed As Double, observedFlag As Boolean
    aliasValue = 2#
    observed = MutateDoubleThenCall(aliasValue, aliasValue, AddressOf CallbackValue)
    RunProbe = "double=" & observed & ":" & aliasValue
    aliasFlag = False
    observedFlag = MutateBooleanThenCall(aliasFlag, aliasFlag, AddressOf BooleanCallbackValue)
    RunProbe = RunProbe & ";boolean=" & observedFlag & ":" & aliasFlag
    RunProbe = RunProbe & ";not=" & BooleanCallbackValue() & ":" & BooleanCallbackValue()
End Function
