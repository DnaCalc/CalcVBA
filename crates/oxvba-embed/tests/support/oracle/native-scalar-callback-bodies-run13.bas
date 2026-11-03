Private aliasValue As Double
Private aliasFlag As Boolean
Public Function CallbackValue() As Double
    aliasValue = aliasValue + 0.5
    CallbackValue = aliasValue
End Function
Public Function BooleanCallbackValue() As Boolean
    aliasFlag = Not aliasFlag
    BooleanCallbackValue = aliasFlag
End Function
