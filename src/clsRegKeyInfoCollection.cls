VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRegKeyInfoCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_Col As New Collection

Public Sub Add( _
    Key As String, _
    ValueName As String, _
    Optional DefaultData As Variant = vbNullString, _
    Optional FixAction As ENUM_REG_ACTION_BASED = RESTORE_VALUE, _
    Optional ValueType As REG_VALUE_TYPE = REG_VALUE_TYPE_UNKNOWN, _
    Optional Description As String = "")
    
    Dim RKI As clsRegKeyInfo
    Set RKI = New clsRegKeyInfo
    RKI.Init Key, ValueName, DefaultData, FixAction, ValueType, Description
    m_Col.Add RKI
    
End Sub

Public Sub AddRKI(RKI As clsRegKeyInfo)
    m_Col.Add RKI
End Sub

Public Function Enumerator() As IEnumVARIANT
Attribute Enumerator.VB_UserMemId = -4
Attribute Enumerator.VB_MemberFlags = "40"
   Set Enumerator = m_Col.[_NewEnum]
End Function

Public Function Item(Index As Long) As clsRegKeyInfo
    Set Item = m_Col.Item(Index)
End Function

Public Function GetCollection() As Collection
    Set GetCollection = m_Col
End Function

Public Property Get Count() As Long
    Count = m_Col.Count
End Property

Public Function KeysToArray() As String()
    If Count <> 0 Then
        Dim RKI As clsRegKeyInfo
        Dim Index As Long
        ReDim arr(1 To Count) As String
        For Each RKI In m_Col
            Index = Index + 1
            arr(Index) = RKI.Key
        Next
        KeysToArray = arr
    End If
End Function

Private Sub Class_Terminate()
    Set m_Col = Nothing
End Sub
