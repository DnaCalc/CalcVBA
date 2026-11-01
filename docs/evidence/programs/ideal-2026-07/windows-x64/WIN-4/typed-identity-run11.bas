Option Explicit
Public Function Same(ByVal left As Variant, ByVal right As Variant) As Boolean
    Same = left Is right
End Function
Public Function Catch(ByVal left As Variant, ByVal right As Variant) As Long
    On Error GoTo Failed
    Dim result As Boolean
    result = left Is right
    Catch = 1
    Exit Function
Failed:
    Catch = Err.Number
End Function
Public Function RunProbe() As String
    Dim a As Object, b As Object, v As Variant, n As Object
    Set a = CreateObject("Scripting.Dictionary")
    Set b = CreateObject("Scripting.Dictionary")
    Set v = a
    RunProbe = "same=" & Same(a, a) & "|different=" & Same(a, b) & "|boxed=" & Same(a, v)
    RunProbe = RunProbe & "|nothing=" & Same(n, Nothing) & "|one=" & Same(a, Nothing)
    RunProbe = RunProbe & "|empty=" & Catch(Empty, a) & "|null=" & Catch(Null, a)
    RunProbe = RunProbe & "|long=" & Catch(0, a) & "|text=" & Catch("x", a) & "|error=" & Catch(CVErr(2042), a)
End Function
