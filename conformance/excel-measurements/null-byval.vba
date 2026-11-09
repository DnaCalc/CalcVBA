Function Run() As String
  Dim h As Target, o As Object, out As String, v As Variant
  Set h = New Target: Set o = h: v = Null
  On Error Resume Next
  Err.Clear: h.MString Null: out = out & "early.lit.String=" & Err.Number & vbLf
  Err.Clear: h.MString v: out = out & "early.var.String=" & Err.Number & vbLf
  Err.Clear: o.MString Null: out = out & "late.lit.String=" & Err.Number & vbLf
  Err.Clear: o.MString v: out = out & "late.var.String=" & Err.Number & vbLf
  Err.Clear: CallByName o, "MString", VbMethod, v: out = out & "cbn.String=" & Err.Number & vbLf
  Err.Clear: h.MCurrency Null: out = out & "early.lit.Currency=" & Err.Number & vbLf
  Err.Clear: h.MCurrency v: out = out & "early.var.Currency=" & Err.Number & vbLf
  Err.Clear: o.MCurrency Null: out = out & "late.lit.Currency=" & Err.Number & vbLf
  Err.Clear: o.MCurrency v: out = out & "late.var.Currency=" & Err.Number & vbLf
  Err.Clear: CallByName o, "MCurrency", VbMethod, v: out = out & "cbn.Currency=" & Err.Number & vbLf
  Err.Clear: h.MSingle Null: out = out & "early.lit.Single=" & Err.Number & vbLf
  Err.Clear: h.MSingle v: out = out & "early.var.Single=" & Err.Number & vbLf
  Err.Clear: o.MSingle Null: out = out & "late.lit.Single=" & Err.Number & vbLf
  Err.Clear: o.MSingle v: out = out & "late.var.Single=" & Err.Number & vbLf
  Err.Clear: CallByName o, "MSingle", VbMethod, v: out = out & "cbn.Single=" & Err.Number & vbLf
  Err.Clear: h.MLong Null: out = out & "early.lit.Long=" & Err.Number & vbLf
  Err.Clear: h.MLong v: out = out & "early.var.Long=" & Err.Number & vbLf
  Err.Clear: o.MLong Null: out = out & "late.lit.Long=" & Err.Number & vbLf
  Err.Clear: o.MLong v: out = out & "late.var.Long=" & Err.Number & vbLf
  Err.Clear: CallByName o, "MLong", VbMethod, v: out = out & "cbn.Long=" & Err.Number & vbLf
  Err.Clear: h.MBoolean Null: out = out & "early.lit.Boolean=" & Err.Number & vbLf
  Err.Clear: h.MBoolean v: out = out & "early.var.Boolean=" & Err.Number & vbLf
  Err.Clear: o.MBoolean Null: out = out & "late.lit.Boolean=" & Err.Number & vbLf
  Err.Clear: o.MBoolean v: out = out & "late.var.Boolean=" & Err.Number & vbLf
  Err.Clear: CallByName o, "MBoolean", VbMethod, v: out = out & "cbn.Boolean=" & Err.Number & vbLf
  Err.Clear: h.MInteger Null: out = out & "early.lit.Integer=" & Err.Number & vbLf
  Err.Clear: h.MInteger v: out = out & "early.var.Integer=" & Err.Number & vbLf
  Err.Clear: o.MInteger Null: out = out & "late.lit.Integer=" & Err.Number & vbLf
  Err.Clear: o.MInteger v: out = out & "late.var.Integer=" & Err.Number & vbLf
  Err.Clear: CallByName o, "MInteger", VbMethod, v: out = out & "cbn.Integer=" & Err.Number & vbLf
  Run = out
End Function
