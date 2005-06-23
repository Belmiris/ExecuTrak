Attribute VB_Name = "modTask"
'THIS MODULE IS FOR VB 32-BIT DEVELOPMENT ONLY

Option Explicit

Private Const TH32CS_SNAPPROCESS = 2

Private Type PROCESSENTRY32
    dwSize As Long
    cntUsage As Long
    th32ProcessID As Long
    th32DefaultHeapID As Long
    th32ModuleID As Long
    cntThreads As Long
    th32ParentProcessID As Long
    pcPriClassBase As Long
    dwFlags As Long
    szExeFile As String * 512
End Type

Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function fnGetWindow Lib "user32" Alias "GetWindow" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function IsWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    
Private Declare Function CreateToolhelp32Snapshot Lib "kernel32.dll" _
            (ByVal dwFlags As Long, _
             ByVal th32ProcessID As Long) As Long

Private Declare Function CloseHandle Lib "kernel32.dll" _
        (ByVal hSnapshot As Long) As Boolean
             
Private Declare Function Process32First Lib "kernel32.dll" _
            (ByVal hSnapshot As Long, _
             lppe As PROCESSENTRY32) As Long

Private Declare Function Process32Next Lib "kernel32.dll" _
            (ByVal hSnapshot As Long, _
             lppe As PROCESSENTRY32) As Long

Global Const SHELL_OK As Integer = 32

Private Function GetWindow(ByVal hwnd As Integer, _
                    ByVal wCmd As Integer) As Integer
                    
    GetWindow = fnUINT2INT(CLng(fnGetWindow(hwnd, wCmd)))
    
End Function
'
'Function        : EndTask - terminate a running windows application.
'Passed Variables: the applications main window handle, form that called the function
'Returns         : true if application terminated successfully, false if it failed
'
Public Function EndTask(ByVal hTargetWnd As Long, _
                        Optional frmCurrent As Variant) As Boolean

    Dim nReturnValue As Long

    On Error Resume Next 'turn off error trapping for this function

    If Not IsMissing(frmCurrent) Then
        If GetWindow(hTargetWnd, GW_OWNER) = frmCurrent.hwnd Then
            EndTask = False
            Exit Function
        End If
    End If

    If IsWindow(hTargetWnd) = False Or (GetWindowLong(hTargetWnd, GWL_STYLE) And WS_DISABLED) Then
        nReturnValue = False
    Else
        
        nReturnValue = PostMessage(hTargetWnd, WM_CANCELMODE, 0, 0&)
        nReturnValue = PostMessage(hTargetWnd, WM_CLOSE, 0, 0&)
        
        nReturnValue = True
        
        DoEvents
    
    End If

    EndTask = nReturnValue

End Function
'
'Function        : IsWndRunning - returns an hWnd for the hInstance handle passed in.
'Passed Variables: sWindowTitle, the windows title.
'                : hInstance, instance handle returned from the Shell function (optional).
'                : hInstance will be set if the program is running.
'Returns         : hWnd, a window handle for the hInstance passed.
'Comments        : The hWnd will be the current active window in the application, not only the main window handle.
Public Function IsWndRunning(sWindowTitle As String) As Long
    
    Const MAX_LENGTH = 128
    
    Dim hTempWnd As Long 'temp window handle
    Dim lProcID As Long
    Dim sTemp As String
    Dim lTemp As Long
    Dim bCheck As Boolean
    
    Dim sName As String
    Dim sWIN_NAME As String
    Dim sWIN_CLASS As String
    
    sName = UCase(Trim(sWindowTitle))
    
    IsWndRunning = 0
    hTempWnd = GetDesktopWindow
    hTempWnd = fnGetWindow(hTempWnd, GW_CHILD)
    
    
    Do While hTempWnd <> 0
        lTemp = GetWindowThreadProcessId(hTempWnd, lProcID)
        
        sWIN_NAME = ""
        sWIN_CLASS = ""
        
        sTemp = Space(MAX_LENGTH)
        lTemp = GetWindowText(hTempWnd, sTemp, MAX_LENGTH)
        sWIN_NAME = UCase(Trim(Left(sTemp, lTemp)))
        lTemp = GetClassName(hTempWnd, sTemp, MAX_LENGTH)
        sWIN_CLASS = UCase(Trim(Left(sTemp, lTemp)))
                
        If InStr(sWIN_NAME, sName) > 0 Then
            IsWndRunning = hTempWnd
            Exit Do
        Else
            If InStr(sWIN_CLASS, sName) > 0 Then
                IsWndRunning = hTempWnd
                Exit Do
            End If
        End If
        
        hTempWnd = fnGetWindow(hTempWnd, GW_HWNDNEXT)
    Loop
End Function

Private Function fnUINT2INT(lValue As Long) As Integer

    If lValue > 32767 Then
        fnUINT2INT = CInt(lValue - 65536)
    Else
        fnUINT2INT = CInt(lValue)
    End If

End Function

Public Function fnExeIsRunning(ByVal sExe As String, _
                               Optional ByVal lProcID As Long = -1) As Boolean
                               '#lProcID added by wj 06/23/05
    Dim lRet As Long
    Dim hSnap As Long
    Dim proc As PROCESSENTRY32
    Dim sRunningExe As String
    Dim nPosi As Integer
    Dim sTemp As String
    
    sExe = UCase(fnExtractFileName(sExe))
    
    ' Windows 95 uses ToolHelp32 functions
    ' Take a picture of current process list
    hSnap = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0)
    If hSnap = -1 Then
        Exit Function
    End If
    proc.dwSize = Len(proc)
     ' Iterate through the processes
    lRet = Process32First(hSnap, proc)
    Do While lRet
        'get the running exe before the zero characters
        nPosi = InStr(proc.szExeFile, Chr(0))
        If nPosi > 0 Then
            sRunningExe = Left(proc.szExeFile, nPosi - 1)
        Else
            sRunningExe = proc.szExeFile
        End If
        sTemp = fnExtractFileName(UCase(Trim(sRunningExe)))
        Debug.Print sTemp
        
        If lProcID < 0 Then
            '#Check EXE only - this is always the case before 06/23/05
            If sTemp = sExe Then
                fnExeIsRunning = True
                CloseHandle hSnap
                Exit Function
            End If
        Else
            '#Check EXE and Process ID
            If sTemp = sExe And lProcID = proc.th32ProcessID Then
                fnExeIsRunning = True
                CloseHandle hSnap
                Exit Function
            End If
        End If
        lRet = Process32Next(hSnap, proc)
    Loop
    
    CloseHandle hSnap
    
    fnExeIsRunning = False
End Function

Public Function fnRunExe(sExe As String, _
                         Optional nMode As Integer = vbNormalFocus, _
                         Optional bForcedRun As Boolean = False, _
                         Optional bCheckRun As Boolean = True, _
                         Optional bShowMsg As Boolean = True, _
                         Optional szErrorMessage As String = "") As Boolean
    
    If Not bForcedRun Then
        If fnExeIsRunning(sExe) Then
            fnRunExe = True
            Exit Function
        End If
    End If
    
    fnRunExe = False
    
    On Error GoTo errLaunching
    'LockWin frmCall  'trap user unput during application load
     
    Dim lTempInstance As Long 'store the instnace handle in a long first
    
    lTempInstance = Shell(sExe, nMode) 'launch the application
    szErrorMessage = sExe & " " & Err.Description
             
    If (lTempInstance < 0) Or (lTempInstance > SHELL_OK) Then
        If bCheckRun Then
            If fnExeIsRunning(sExe) Then
                fnRunExe = True
            Else
                If bShowMsg Then
                    MsgBox "Unable to launch program " & sExe, vbOKOnly + vbCritical, "Error"
                End If
            End If
        Else
            fnRunExe = True
        End If
    Else 'error occured clear handles and display error message
        If bShowMsg Then
            MsgBox szErrorMessage, vbOKOnly + vbCritical, "Error"
        End If
    End If
    
    Exit Function

errLaunching:
    szErrorMessage = "Unable to launch program " & sExe & " (" & Err.Description & ")"
    
    If bShowMsg Then
        MsgBox szErrorMessage, vbOKOnly + vbCritical
    End If
End Function

Private Function fnExtractFileName(ByVal sPath As String) As String

    Dim i As Integer
    Dim sTemp As String
    Dim sChar As String * 1
    
    'david 01/03/2001
    sPath = UCase(sPath)
    i = InStr(sPath, ".EXE")
    If i > 0 Then
        sTemp = Left(sPath, i - 1)
        i = InStrRev(sTemp, "\")
        
        If i > 0 Then
            fnExtractFileName = Mid(sTemp, i + 1)
        Else
            fnExtractFileName = sTemp
        End If
    Else
        i = Len(sPath)
        Do
            sChar = Mid(sPath, i, 1)
            i = i - 1
            If sChar = "." Then
                Exit Do
            End If
        Loop Until i = 0
        If i = 0 Then
            fnExtractFileName = sPath
        Else
            sTemp = ""
            Do While i > 0
                sChar = Mid(sPath, i, 1)
                If sChar = "\" Then
                    Exit Do
                End If
                sTemp = sChar & sTemp
                i = i - 1
            Loop
            fnExtractFileName = sTemp
        End If
    End If
End Function

Public Function fnKillProgram(hProgram As Long) As Integer
    fnKillProgram = EndTask(hProgram)
    
    If fnKillProgram <> 0 Then
        #If DEVELOP Then
            MsgBox "Scheduler terminated"
        #End If
    Else
        #If DEVELOP Then
            MsgBox "Error - Cannot terminate Scheduler"
        #End If
    End If
End Function


