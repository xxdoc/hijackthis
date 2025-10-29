VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2340
   ClientLeft      =   105
   ClientTop       =   435
   ClientWidth     =   3630
   LinkTopic       =   "Form1"
   ScaleHeight     =   2340
   ScaleWidth      =   3630
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private oAll As Object

Private Declare Function GetFileAttributes Lib "kernel32.dll" Alias "GetFileAttributesW" (ByVal lpFileName As Long) As Long
Private Declare Function GetFullPathName Lib "kernel32.dll" Alias "GetFullPathNameW" (ByVal lpFileName As Long, ByVal nBufferLength As Long, ByVal lpBuffer As Long, ByVal lpFilePart As Long) As Long

Private Const INVALID_FILE_ATTRIBUTES   As Long = -1&
Private Const FILE_ATTRIBUTE_DIRECTORY  As Long = &H10&
Private Const MAX_PATH_W As Long = 32767

Private Sub Form_Load()
    
    Set oAll = CreateObject("Scripting.Dictionary")
    
    'get all hashes
    Dim ff%, s$
    ff = FreeFile()
    Open App.Path & "\Hashes.csv" For Input As #ff
    Do While Not EOF(ff)
        Line Input #ff, s
        SaveHash s
    Loop
    Close #ff
    
    Dim sCurrentDB As String
    sCurrentDB = GetAbsolutePath("..\..\..\..\..\database\DisallowedCert.txt")
    
    If Not FileExists(sCurrentDB) Then
        sCurrentDB = GetAbsolutePath("..\..\..\..\database\DisallowedCert.txt")
        If Not FileExists(sCurrentDB) Then
            MsgBox "Failed to find DB file: " & sCurrentDB
            Unload Me
            Exit Sub
        End If
    End If
    
    Open sCurrentDB For Input As #ff
    
    Do While Not EOF(ff)
        Line Input #ff, s
        RemoveHash s
    Loop
    Close #ff
    
    Dim key As Variant
    Dim iCount As Long
    Open App.Path & "\Hashes_new.txt" For Output As #ff
    For Each key In oAll.keys
        'Print #ff, "        .Add """ & oAll(key) & """, """ & key & """"
        Print #ff, oAll(key) & ";" & key
        iCount = iCount + 1
    Next
    Close #ff
    
    Set oAll = Nothing
    
    MsgBox "Hashes_new.txt file is generated with " & iCount & " records."
    
    Unload Me
End Sub

Sub RemoveHash(s$)
    Dim sName As String
    Dim sHash As String
    Dim arr
    If Len(s) = 0 Then Exit Sub
    
    arr = Split(s, ";")
    If UBound(arr) = 1 Then
        sName = arr(0)
        sHash = arr(1)
        If oAll.Exists(sHash) Then
            oAll.Remove sHash
            Debug.Print "[REMOVED] " & sHash & " - " & sName
        End If
    End If
End Sub

Sub SaveHash(s$)
    If Len(s) = 0 Then Exit Sub
    Dim pos&
    Dim sName As String
    Dim sHash As String
    pos = InStr(s, ";")
    If pos <> 0 Then
        sName = Left$(s, pos - 1)
        If sName <> "Certificate name" Then
            s = Mid$(s, pos + 1)
            pos = InStr(s, ";")
            If pos <> 0 Then
                sHash = Left$(s, pos - 1)
                If Not oAll.Exists(sHash) Then
                    oAll.Add sHash, sName
                    Debug.Print "[ADDED] " & sHash & " - " & sName
                End If
            End If
        End If
    End If
End Sub

Public Function FileExists(ByVal sFile As String, Optional bUseWow64 As Boolean, Optional bAllowNetwork As Boolean) As Boolean
    
    Dim ret As Long
    ret = GetFileAttributes(StrPtr(sFile))
    If ret <> INVALID_FILE_ATTRIBUTES And (0 = (ret And FILE_ATTRIBUTE_DIRECTORY)) Then
        FileExists = True
    End If
    
End Function

Public Function GetAbsolutePath(sFilename As String) As String
    Dim cnt        As Long
    Dim sFullName  As String
    sFullName = String$(MAX_PATH_W, 0)
    cnt = GetFullPathName(StrPtr(sFilename), MAX_PATH_W, StrPtr(sFullName), 0&)
    If cnt Then
        GetAbsolutePath = Left$(sFullName, cnt)
    Else
        GetAbsolutePath = sFilename
    End If
    If Right$(GetAbsolutePath, 1) = "\" Then GetAbsolutePath = Left$(GetAbsolutePath, Len(GetAbsolutePath) - 1)
End Function
