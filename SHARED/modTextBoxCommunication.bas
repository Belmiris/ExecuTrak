Attribute VB_Name = "modTextBoxCommunication"
Option Explicit

Private Type PROCESS_INFORMATION
   hProcess As Long
   hThread As Long
   dwProcessId As Long
   dwThreadId As Long
End Type

Private Type STARTUPINFO
   cb As Long
   lpReserved As String
   lpDesktop As String
   lpTitle As String
   dwX As Long
   dwY As Long
   dwXSize As Long
   dwYSize As Long
   dwXCountChars As Long
   dwYCountChars As Long
   dwFillAttribute As Long
   dwFlags As Long
   wShowWindow As Integer
   cbReserved2 As Integer
   lpReserved2 As Long
   hStdInput As Long
   hStdOutput As Long
   hStdError As Long
End Type

Private Declare Function CreateProcess Lib "kernel32" _
   Alias "CreateProcessA" _
   (ByVal lpApplicationName As String, _
   ByVal lpCommandLine As String, _
   lpProcessAttributes As Any, _
   lpThreadAttributes As Any, _
   ByVal bInheritHandles As Long, _
   ByVal dwCreationFlags As Long, _
   lpEnvironment As Any, _
   ByVal lpCurrentDriectory As String, _
   lpStartupInfo As STARTUPINFO, _
   lpProcessInformation As PROCESS_INFORMATION) As Long

Private Declare Function OpenProcess Lib "kernel32.dll" _
   (ByVal dwAccess As Long, _
   ByVal fInherit As Integer, _
   ByVal hObject As Long) As Long

Private Declare Function TerminateProcess Lib "kernel32" _
   (ByVal hProcess As Long, _
   ByVal uExitCode As Long) As Long

Private Declare Function CloseHandle Lib "kernel32" _
   (ByVal hObject As Long) As Long

Declare Function IsWindow Lib "user32" (ByVal hWnd As Long) As Long

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" ( _
    ByVal hWnd As Long, _
    ByVal wMsg As Long, _
    ByVal wParam As Integer, _
    ByVal lParam As String) As Long

Const SYNCHRONIZE = 1048576
Const NORMAL_PRIORITY_CLASS = &H20&
Const WM_SETTEXT = &HC&

Public Function RunTextBoxApp(szExePath As String, szParams As String) As Long
    Dim pInfo As PROCESS_INFORMATION
    Dim sInfo As STARTUPINFO
    Dim sNull As String
    Dim baseFolder As String
    Dim lSuccess As Long
    Dim lRetValue As Long
    Dim cmd As String
    
    cmd = IIf(Len(szParams) > 0, szExePath & " " & szParams, szExePath)
    
    sInfo.cb = Len(sInfo)
    lSuccess = CreateProcess(sNull, _
                            cmd, _
                            ByVal 0&, _
                            ByVal 0&, _
                            1&, _
                            NORMAL_PRIORITY_CLASS, _
                            ByVal 0&, _
                            sNull, _
                            sInfo, _
                            pInfo)
    
    RunTextBoxApp = pInfo.hProcess
    
End Function

Public Sub SendTextBoxMessage(hWnd As Long, message As String)
    
    Call SendMessage(hWnd, WM_SETTEXT, 0&, message)
    
End Sub

Public Function IsWindowValid(hWnd As Long) As Boolean
    
    If hWnd <> 0 Then
        IsWindowValid = IsWindow(hWnd) <> 0
    Else
        IsWindowValid = False
    End If
    
End Function
