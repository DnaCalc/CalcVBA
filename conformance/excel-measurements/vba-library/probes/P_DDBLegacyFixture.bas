Attribute VB_Name = "P_DDBLegacyFixture"
Option Explicit
' UNMEASURED exact bits of an old test's formerly approximate integer expectation.
Public Function Run() As String
    Dim v As Variant
    On Error GoTo fail
    v = DDB(10000, 1000, 5, 5)
    Run = "DDBLegacyFixture.1=" & Sh(v) & vbLf
    Exit Function
fail:
    Run = "DDBLegacyFixture.1=ERR" & Err.Number & vbLf
End Function
