VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRegKeyInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_Key As String
Private m_DataChecker As clsDataChecker
Private m_Description As String
Private m_ValueType As REG_VALUE_TYPE
Private m_FixAction As ENUM_REG_ACTION_BASED

Public Sub Init( _
    Key As String, _
    ValueName As String, _
    Optional DefaultData As Variant = vbNullString, _
    Optional FixAction As ENUM_REG_ACTION_BASED = RESTORE_VALUE, _
    Optional ValueType As REG_VALUE_TYPE = REG_VALUE_TYPE_UNKNOWN, _
    Optional Description As String = "")
    
    m_Key = Key
    Set m_DataChecker = New clsDataChecker
    m_DataChecker.AddValueData ValueName, DefaultData
    m_DataChecker.MoveNext
    m_FixAction = FixAction
    m_Description = Description
End Sub

Public Property Get Key() As String
    Key = m_Key
End Property

Public Property Get ValueName() As String
    ValueName = m_DataChecker.ValueNameStr
End Property

Public Property Get DefaultData() As String
    DefaultData = m_DataChecker.DataStr
End Property

Public Property Get DefaultDataLong() As Long
    DefaultDataLong = m_DataChecker.DataLong
End Property

Public Property Get DataChecker() As clsDataChecker
    DataChecker = m_DataChecker
End Property

Public Property Get Description() As String
    Description = m_Description
End Property

Public Property Get ValueType() As REG_VALUE_TYPE
    ValueType = m_ValueType
End Property

Public Property Let ValueType(ByVal valType As REG_VALUE_TYPE)
    m_ValueType = valType
End Property

Public Property Get FixAction() As ENUM_REG_ACTION_BASED
    FixAction = m_FixAction
End Property

Private Sub Class_Terminate()
    Set m_DataChecker = Nothing
End Sub
