Attribute VB_Name = "modProcess"
Option Explicit

Private Type SHELLEXECUTEINFO
    cbSize          As Long
    fMask           As Long
    hWnd            As Long
    lpVerb          As Long
    lpFile          As Long
    lpParameters    As Long
    lpDirectory     As Long
    nShow           As Long
    hInstApp        As Long
    lpIDList        As Long
    lpClass         As Long
    hkeyClass       As Long
    dwHotKey        As Long
    hIcon           As Long
    hProcess        As Long
End Type

Private Declare Function ShellExecuteEx Lib "shell32.dll" Alias "ShellExecuteExW" (SEI As SHELLEXECUTEINFO) As Long
Private Declare Function GetExitCodeProcess Lib "kernel32.dll" (ByVal hProcess As Long, lpExitCode As Long) As Long
Private Declare Function CloseHandle Lib "kernel32.dll" (ByVal hObject As Long) As Long


Public Function RunAsAndWait(sFile As String, sArgs As String) As Long
    'returns exit code; 1 - if failed to launch
    
    Const SEE_MASK_NOCLOSEPROCESS As Long = &H40&
    Const SEE_MASK_NOASYNC As Long = &H100&
    Const SEE_MASK_NO_CONSOLE As Long = &H8000&
    
    Dim uSEI As SHELLEXECUTEINFO
    With uSEI
        .cbSize = Len(uSEI)
        .fMask = SEE_MASK_NOCLOSEPROCESS Or SEE_MASK_NOASYNC Or SEE_MASK_NO_CONSOLE
        .lpFile = StrPtr(sFile)
        .lpParameters = StrPtr(sArgs)
        .lpDirectory = StrPtr(App.Path)
        .lpVerb = StrPtr("runas")
        .nShow = 1
    End With
    ShellExecuteEx uSEI
    
    If uSEI.hInstApp <= 32 Then
        WriteC "ShellExecuteEx failed with error: " & uSEI.hInstApp, cErr
    End If
    
    If 0 = uSEI.hProcess Then
'        If uSEI.hInstApp > 0 Then
'            RunAsAndWait = uSEI.hInstApp
'        Else
'            RunAsAndWait = 1
'        End If
    Else
        GetExitCodeProcess uSEI.hProcess, RunAsAndWait
        CloseHandle uSEI.hProcess
    End If
End Function


Public Function ParseCommandLine(Line As String, argc As Long, argv() As String) As Boolean
  On Error GoTo ErrorHandler
  Dim Lex$(), nL&, nA&, Unit$, St$
  St = Line
  If Len(St) > 0 Then ParseCommandLine = True
  Lex = Split(St) '��������� �� �������� �� ������� ��� ������� ������
  ReDim argv(0 To UBound(Lex) + 1) As String '���������� �������� ������ �� ����������� ���������� ����� ����������
  argv(0) = App.Path
  If Len(St) <> 0 Then
    Do While nL <= UBound(Lex)
      Unit = Lex(nL) '���������� ������� ������� ��� ������ ������ ���������
      If Len(Unit) <> 0 Then '������ �� ������� �������� ����� �����������
        '���� � ������� ������� ������� ��� �������� �� �����, �� �������� ������� "������������"
        If (Len(Lex(nL)) - Len(Replace$(Lex(nL), """", ""))) Mod 2 = 1 Then
          Do
            nL = nL + 1
            If nL > UBound(Lex) Then Exit Do '���� �� ��������� ����������� �������, � ������ ������ ���
            Unit = Unit & " " & Lex(nL) '��������� �������� ��������
          ' �������� ������ ����������� 1 ��� �������� ������ ������� ������� �� ����� ������������ � ��� ������ ��������� (����� ����� �������)
          Loop Until (Len(Lex(nL)) - Len(Replace$(Lex(nL), """", ""))) Mod 2 = 1
        End If
        Unit = Replace$(Unit, """", "") '������� �������
        nA = nA + 1 '������� ���-�� �������� ����������
        argv(nA) = Unit
      End If
      nL = nL + 1 '������� ������� �������
    Loop
  End If
  ReDim Preserve argv(0 To nA) ' ������� ������ �� ��������� ����� ����������
  argc = nA
  Exit Function
ErrorHandler:
  WriteC "Parser.ParseCommandLine", cErr
End Function
