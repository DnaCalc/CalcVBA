VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "Calculator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Event Changed(ByVal value As Variant)
Private mValue As Long
Private mForeign As Object
Private Sub Class_Initialize()
    mValue = 5
End Sub
Public Property Get Value() As Long
    Value = mValue
End Property
Public Property Let Value(ByVal value As Long)
    mValue = value
End Property
Public Function Add(ByVal a As Long, ByVal b As Long) As Long
    Add = a + b
End Function
Public Sub Bump(ByRef number As Long, ByRef text As String)
    number = number + 1
    text = text & ChrW(&H3BB)
End Sub
Public Function Echo(ByVal value As Variant) As Variant
    Echo = value
End Function
Public Function Values() As Variant
    Dim data(1 To 2, -1 To 1) As Variant
    data(1, -1) = 7
    data(1, 0) = ChrW(&H3BB)
    data(1, 1) = True
    data(2, -1) = CVErr(2042)
    data(2, 0) = Empty
    data(2, 1) = 9
    Values = data
End Function
Public Function MakeItem() As Object
    Set MakeItem = New PublicItem
End Function
Public Function PrivateValue() As Long
    Dim item As PrivateItem
    Set item = New PrivateItem
    PrivateValue = item.Ping()
End Function
Public Sub Fire(ByVal value As Variant)
    RaiseEvent Changed(value)
End Sub
Public Sub RaiseError()
    Err.Raise 457, "Calculator", "private COM test", "test.chm", 91
End Sub
Public Function MissingValue(Optional ByVal value As Variant) As Boolean
    MissingValue = IsMissing(value)
End Function
Public Sub AliasPair(ByRef first As Long, ByRef second As Long)
    first = first + 1
    second = second + 10
End Sub
Public Sub ReplaceItem(ByRef value As Object)
    Set value = New PublicItem
End Sub
Public Sub MutateThenRaise(ByRef number As Long, ByRef text As String)
    number = 67
    text = "fault" & ChrW(&H3BB)
    RaiseError
End Sub
Public Function ObjectCount(ByVal value As Object, Optional ByVal extra As Variant) As Long
    ObjectCount = value.Count
    If IsMissing(extra) Then ObjectCount = ObjectCount + 10
End Function
Public Sub KeepObject(ByRef value As Object)
    value.Add "kept", 42
End Sub
Public Function EchoObject(ByVal value As Object) As Object
    Set EchoObject = value
End Function
Public Function ObjectName(ByVal value As Object) As String
    ObjectName = value.Name
End Function
Public Property Set Foreign(ByVal value As Object)
    Set mForeign = value
End Property
Public Property Get Foreign() As Object
    Set Foreign = mForeign
End Property
