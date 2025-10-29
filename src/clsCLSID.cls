VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCLSID"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public FileName As String
Public FriendlyName As String
Public ProgID As String
Public ClassName As String
Public RegRedirected As Boolean
Public UserData As New Collection

Public Sub InitByProgID(sProgId As String, x32RegRedirected As Boolean)
    ProgID = sProgId
    RegRedirected = x32RegRedirected
    GetFileByCLSID sProgId, FileName, , x32RegRedirected
    FriendlyName = Reg.GetString(HKCR, ProgID, "", x32RegRedirected)
End Sub

Public Sub InitByClassName(sClassName As String, x32RegRedirected As Boolean)
    ClassName = sClassName
    RegRedirected = x32RegRedirected
    ProgID = Reg.GetString(HKCR, sClassName & "\Clsid", "", x32RegRedirected)
    InitByProgID ProgID, x32RegRedirected
End Sub

Public Sub AddUserData(data As String)
    UserData.Add data
End Sub

