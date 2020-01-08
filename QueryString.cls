VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "QueryString"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim m_Query As String
Dim Coll As New Collection
Dim m_Return As String


Public Function Init()
Dim x, y, ub, j
x = Split(m_Query, ";", , vbTextCompare)
ub = UBound(x)
Set Coll = New Collection
For j = 0 To ub
    y = Split(x(j), "=")
    If UBound(y) = 1 Then
        Coll.Add y(1), y(0)
    End If
    DoEvents
Next
End Function

Public Property Get Query() As String
    Query = m_Query
End Property
Public Property Let Query(ByVal vNewValue As String)
    m_Query = vNewValue
    Init
End Property

Public Function Add(Key As String, value As String)
m_Query = m_Query & Key & "=" & value & ";"
End Function

Public Property Get Item(Key As String) As String
Attribute Item.VB_UserMemId = 0
Attribute Item.VB_MemberFlags = "200"
On Error Resume Next
Item = Coll(Key)
End Property
