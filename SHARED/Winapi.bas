Attribute VB_Name = "modWinApi"
'***********************************************************'
'
' Copyright (c) 1996 FACTOR, A Division of W.R.Hess Company
'
' Module name   : WINAPI.BAS
'
' This module implements win api constants and declarations
'
' Functions:
'
Option Explicit

'================================
'Windows Data Types and Constants
'================================

Public Const SW_HIDE = 0
Public Const SW_MINIMIZE = 6
Public Const SW_RESTORE = 9
Public Const SW_SHOW = 5
Public Const SW_SHOWMAXIMIZED = 3
Public Const SW_SHOWMINIMIZED = 2
Public Const SW_SHOWMINNOACTIVE = 7
Public Const SW_SHOWNA = 8
Public Const SW_SHOWNOACTIVATE = 4
Public Const SW_SHOWNORMAL = 1

' Global Const RIGHT_BUTTON = 2 'Mouse Button Flags defined in constant.txt
Global Const TWIPS = 1 'ScaleMode Constant

Public Const GW_HWNDFIRST = 0
Public Const GW_HWNDLAST = 1
Public Const GW_HWNDNEXT = 2
Public Const GW_HWNDPREV = 3
Public Const GW_OWNER = 4
Public Const GW_CHILD = 5
Public Const GWL_STYLE = (-16)
Public Const GWW_HINSTANCE = (-6)
Public Const GWW_HWNDPARENT = (-8)

Public Const WM_CLOSE = &H10
Public Const WM_CANCELMODE = &H1F

Public Const WS_DISABLED = &H8000000

Public Const WM_SYSCOMMAND = &H112

Public Const HWND_BROADCAST = &HFFFF
Public Const HWND_DESKTOP = 0

Public Const HWND_TOP = 0
Public Const HWND_BOTTOM = 1
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2

'Show window position constants
Public Const SWP_DRAWFRAME = &H20
Public Const SWP_HIDEWINDOW = &H80
Public Const SWP_NOACTIVATE = &H10
Public Const SWP_NOCOPYBITS = &H100
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOREDRAW = &H8
Public Const SWP_NOREPOSITION = &H200
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOZORDER = &H4
Public Const SWP_SHOWWINDOW = &H40

Public Const SC_SIZE = &HF000
Public Const SC_MOVE = &HF010
Public Const SC_MINIMIZE = &HF020
Public Const SC_MAXIMIZE = &HF030
Public Const SC_NEXTWINDOW = &HF040
Public Const SC_PREVWINDOW = &HF050
Public Const SC_CLOSE = &HF060
Public Const SC_VSCROLL = &HF070
Public Const SC_HSCROLL = &HF080
Public Const SC_MOUSEMENU = &HF090
Public Const SC_KEYMENU = &HF100
Public Const SC_ARRANGE = &HF110
Public Const SC_RESTORE = &HF120
Public Const SC_TASKLIST = &HF130
Public Const SC_SCREENSAVE = &HF140
Public Const SC_HOTKEY = &HF150

#If Win32 Then
    Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
    End Type
    Type POINTAPI
        x As Long
        y As Long
    End Type
    Type WINDOWPLACEMENT
        Length As Long
        flags As Long
        showCmd As Long
        ptMinPosition As POINTAPI
        ptMaxPosition As POINTAPI
        rcNormalPosition As RECT
    End Type
    Declare Function GetDC Lib "user32" ( _
        ByVal hwnd As Long) As Long
        
    Declare Function ReleaseDC Lib "user32" ( _
        ByVal hwnd As Long, _
        ByVal hdc As Long) As Long
        
    Declare Function ExtractIcon Lib "shell32.dll" Alias "ExtractIconA" ( _
        ByVal hInst As Long, _
        ByVal lpszExeFileName As String, _
        ByVal nIconIndex As Long) As Long
        
    Declare Function DrawIcon Lib "user32" ( _
        ByVal hdc As Long, _
        ByVal x As Long, _
        ByVal y As Long, _
        ByVal hIcon As Long) As Long
        
    Declare Function GetCursorPos Lib "user32" ( _
        lpPoint As POINTAPI) As Long
        
    Declare Function GetKeyState Lib "user32" ( _
        ByVal nVirtKey As Long) As Integer
        
    Declare Function GetWindowRect Lib "user32" ( _
        ByVal hwnd As Long, _
        lpRect As RECT) As Long
        
    Declare Function lstrcpy Lib "kernel32" Alias "lstrcpyA" ( _
        ByVal lpString1 As String, _
        ByVal lpString2 As String) As Long
        
    Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" ( _
        ByVal lpApplicationName As String, _
        ByVal lpKeyName As Any, _
        ByVal lpDefault As String, _
        ByVal lpReturnedString As String, _
        ByVal nSize As Long, _
        ByVal lpFileName As String) As Long
        
    Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" ( _
        ByVal lpApplicationName As String, _
        ByVal lpKeyName As Any, _
        ByVal lpString As Any, _
        ByVal lpFileName As String) As Long
        
    Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" ( _
        ByVal lpBuffer As String, _
        ByVal nSize As Long) As Long
        
    Declare Function GetWindowPlacement Lib "user32" ( _
        ByVal hwnd As Long, _
        lpwndpl As WINDOWPLACEMENT) As Long
        
    Declare Function GetDesktopWindow Lib "user32" () As Long
    
    Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" ( _
        ByVal lpBuffer As String, _
        ByVal nSize As Long) As Long
        
    Declare Function SetWindowPos Lib "user32" ( _
        ByVal hwnd As Long, _
        ByVal hWndInsertAfter As Long, _
        ByVal x As Long, _
        ByVal y As Long, _
        ByVal cx As Long, _
        ByVal cy As Long, _
        ByVal wFlags As Long) As Long
        
    Declare Function ShowWindow Lib "user32" ( _
        ByVal hwnd As Long, _
        ByVal nCmdShow As Long) As Long
        
    Declare Function SetFocus Lib "user32" ( _
        ByVal hwnd As Long) As Long
        
    Declare Function GetMenu Lib "user32" ( _
        ByVal hwnd As Long) As Long
        
    Declare Function GetSubMenu Lib "user32" ( _
        ByVal hMenu As Long, _
        ByVal nPos As Long) As Long
        
    Declare Function GetSystemMenu Lib "user32" ( _
        ByVal hwnd As Long, _
        ByVal bRevert As Long) As Long
        
    Declare Function TrackPopupMenu Lib "user32" ( _
        ByVal hMenu As Long, _
        ByVal wFlags As Long, _
        ByVal x As Long, _
        ByVal y As Long, _
        ByVal nReserved As Long, _
        ByVal hwnd As Long, _
        lprc As RECT) As Long
        
    Declare Function InsertMenu Lib "user32" Alias "InsertMenuA" ( _
        ByVal hMenu As Long, _
        ByVal nPosition As Long, _
        ByVal wFlags As Long, _
        ByVal wIDNewItem As Long, _
        ByVal lpNewItem As String) As Long
        
    Declare Function EnableWindow Lib "user32" ( _
        ByVal hwnd As Long, _
        ByVal fEnable As Long) As Long
        
    Declare Function ModifyMenu Lib "user32" Alias "ModifyMenuA" ( _
        ByVal hMenu As Long, _
        ByVal nPosition As Long, _
        ByVal wFlags As Long, _
        ByVal wIDNewItem As Long, _
        ByVal lpString As String) As Long
        
    Declare Function WinHelp Lib "user32" Alias "WinHelpA" ( _
        ByVal hwnd As Long, _
        ByVal lpHelpFile As String, _
        ByVal wCommand As Long, _
        ByVal dwData As Long) As Long
        
    Declare Function GetParent Lib "user32" ( _
        ByVal hwnd As Long) As Long
        
    Declare Function GetLastActivePopup Lib "user32" ( _
        ByVal hwndOwnder As Long) As Long
#Else
    Type POINTAPI 'Point structure
        x As Integer
        y As Integer
    End Type
    Type RECT 'used for tooltip and context menus
        Left As Integer
        Top As Integer
        Right As Integer
        Bottom As Integer
    End Type

    Type WINDOWPLACEMENT 'Used for Context menus
        Length As Integer
        flags As Integer
        showCmd As Integer
        PtMinPos As POINTAPI
        ptMaxPos As POINTAPI
        rcNormalPos As RECT
    End Type
    
    '=================
    'Windows API Calls
    '=================
    Declare Function GetDC Lib "USER" ( _
        ByVal hwnd As Integer _
        ) As Integer
    
    Declare Sub ReleaseDC Lib "USER" ( _
        ByVal hwnd As Integer, _
        ByVal hdc As Integer)
    
    Declare Function ExtractIcon Lib "shell" ( _
        ByVal hInstance As Integer, _
        ByVal lpszAppName As String, _
        ByVal Icon As Integer) _
        As Integer
    
    Declare Function DrawIcon Lib "USER" ( _
        ByVal hdc As Integer, _
        ByVal x As Integer, _
        ByVal y As Integer, _
        ByVal hIcon As Integer) _
        As Integer
    
    Declare Sub GetCursorPos Lib "USER" ( _
        lpPoint As POINTAPI)
    
    Declare Function GetKeyState Lib "USER" ( _
        ByVal Vkey As Integer _
        ) As Integer
    
    Declare Sub GetWindowRect Lib "USER" ( _
        ByVal hwnd As Integer, _
        lpRect As RECT)
    
    Declare Function lstrcpy Lib "kernel" ( _
        ByVal lpString1 As Any, _
        ByVal lpString2 As Any _
        ) As Long
    
    Declare Function GetPrivateProfileString Lib "kernel" ( _
        ByVal lpApplicationName As String, _
        ByVal lpKeyName As Any, _
        ByVal lpDefault As String, _
        ByVal lpReturnedString As String, _
        ByVal nSize As Integer, _
        ByVal lpFileName As String _
        ) As Integer
    
    Declare Function WritePrivateProfileString Lib "kernel" ( _
        ByVal lpApplicationName As String, _
        ByVal lpKeyName As Any, _
        ByVal lpString As Any, _
        ByVal lplFileName As String _
        ) As Integer
    
    Declare Function GetWindowsDirectory Lib "kernel" ( _
        ByVal lpBuffer As String, _
        ByVal nSize As Integer _
        ) As Integer
    
    Declare Function SetWindowsPlacement Lib "USER" ( _
        ByVal hwnd As Integer, _
        lpwndpl As WINDOWPLACEMENT _
        ) As Integer
    
    Declare Function GetDesktopWindow Lib "USER" Alias "GetDeskTopWindow" () As Integer
    
    Declare Sub SetWindowPos Lib "USER" ( _
        ByVal h1 As Integer, _
        ByVal h2 As Integer, _
        ByVal x As Integer, _
        ByVal y As Integer, _
        ByVal cx As Integer, _
        ByVal cy As Integer, _
        ByVal f As Integer)
    
    
    Declare Function GetSystemDirectory Lib "kernel" ( _
        ByVal lpBuffer As String, _
        ByVal nSize As Integer _
        ) As Integer
    
    Declare Function ShowWindow Lib "USER" ( _
        ByVal hwnd As Integer, _
        ByVal nCmdShow As Integer _
        ) As Integer
    
    Declare Function SetFocusAPI Lib "USER" Alias "SetFocus" ( _
        ByVal hwnd As Integer _
        ) As Integer
    
    Declare Function GetMenu Lib "USER" ( _
        ByVal hwnd As Integer _
        ) As Integer
    
    Declare Function GetSubMenu Lib "USER" ( _
        ByVal hMenu As Integer, _
        ByVal nPos As Integer _
        ) As Integer
    
    Declare Function GetSystemMenu Lib "USER" ( _
        ByVal hwnd As Integer, _
        ByVal bRevert As Integer _
        ) As Integer
    
    Declare Function TrackPopupMenu Lib "USER" ( _
        ByVal hMenu As Integer, _
        ByVal wFlags As Integer, _
        ByVal x As Integer, _
        ByVal y As Integer, _
        ByVal nReserved As Integer, _
        ByVal hwnd As Integer, _
        lpReserved As Any _
        ) As Integer
    
    Declare Function InsertMenu Lib "USER" ( _
        ByVal hMenu As Integer, _
        ByVal nPosition As Integer, _
        ByVal wFlags As Integer, _
        ByVal wIDNewItem As Integer, _
        ByVal lpNewItem As Any _
        ) As Integer
    
    Declare Function EnableWindow Lib "USER" ( _
        ByVal hwnd As Integer, _
        ByVal aBOOL As Integer _
        ) As Integer
    
    Declare Function ModifyMenu Lib "USER" ( _
        ByVal hMenu As Integer, _
        ByVal nPosition As Integer, _
        ByVal wFlags As Integer, _
        ByVal wIDNewItem As Integer, _
        ByVal lpString As Any _
        ) As Integer
    
    Declare Function WinHelp Lib "USER" ( _
        ByVal hwnd As Integer, _
        ByVal lpHelpFile As String, _
        ByVal wCommand As Integer, _
        ByVal dwData As Any _
        ) As Integer
    
    Declare Function GetParent Lib "USER" ( _
        ByVal hwnd As Integer _
        ) As Integer
    
    Declare Function GetLastActivePopup Lib "USER" (ByVal hwndOwnder As Integer) As Integer
    Declare Function AnyPopup Lib "USER" () As Integer
#End If
    
