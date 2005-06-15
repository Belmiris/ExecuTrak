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
Public Const xSC_CLOSE  As Long = -10
Public Const SC_VSCROLL = &HF070
Public Const SC_HSCROLL = &HF080
Public Const SC_MOUSEMENU = &HF090
Public Const SC_KEYMENU = &HF100
Public Const SC_ARRANGE = &HF110
Public Const SC_RESTORE = &HF120
Public Const SC_TASKLIST = &HF130
Public Const SC_SCREENSAVE = &HF140
Public Const SC_HOTKEY = &HF150

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
    length As Long
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
    ByVal hDC As Long) As Long
    
Declare Function ExtractIcon Lib "shell32.dll" Alias "ExtractIconA" ( _
    ByVal hInst As Long, _
    ByVal lpszExeFileName As String, _
    ByVal nIconIndex As Long) As Long
    
Declare Function DrawIcon Lib "user32" ( _
    ByVal hDC As Long, _
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

'rajneesh & david 10/30/00
Public Const MF_BYPOSITION = &H400&
Public Const MF_REMOVE = &H1000&

Public Declare Function GetMenuItemCount Lib "user32" _
    (ByVal hMenu As Long) As Long

Public Declare Function DrawMenuBar Lib "user32" _
    (ByVal hMenu As Long) As Long

Public Declare Function RemoveMenu Lib "user32" _
    (ByVal hMenu As Long, ByVal nPosition As Long, _
    ByVal wFlags As Long) As Long


Public Declare Function CopyFile Lib "kernel32.dll" Alias "CopyFileA" _
    (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, _
    ByVal bFailIfExists As Long) As Long

'david 11/15/00
Public Declare Function SendMessageByNum Lib "user32" Alias _
"SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal _
wParam As Long, ByVal lParam As Long) As Long

Public Const LB_SETHORIZONTALEXTENT = &H194

'david 04/01/2002
'''''''''''''' change screen resolution API
Public Const EWX_LOGOFF = 0
Public Const EWX_SHUTDOWN = 1
Public Const EWX_REBOOT = 2
Public Const EWX_FORCE = 4
Public Const CCDEVICENAME = 32
Public Const CCFORMNAME = 32
Public Const DM_BITSPERPEL = &H40000
Public Const DM_PELSWIDTH = &H80000
Public Const DM_PELSHEIGHT = &H100000
Public Const CDS_UPDATEREGISTRY = &H1
Public Const CDS_TEST = &H4
Public Const DISP_CHANGE_SUCCESSFUL = 0
Public Const DISP_CHANGE_RESTART = 1

'Public Const HWND_BROADCAST = &HFFFF&
Public Const WM_DISPLAYCHANGE = &H7E&
Public Const SPI_SETNONCLIENTMETRICS = 42

Public Const SM_CXSCREEN = 0
Public Const SM_CYSCREEN = 1

Type typDevMODE
    dmDeviceName       As String * CCDEVICENAME
    dmSpecVersion      As Integer
    dmDriverVersion    As Integer
    dmSize             As Integer
    dmDriverExtra      As Integer
    dmFields           As Long
    dmOrientation      As Integer
    dmPaperSize        As Integer
    dmPaperLength      As Integer
    dmPaperWidth       As Integer
    dmScale            As Integer
    dmCopies           As Integer
    dmDefaultSource    As Integer
    dmPrintQuality     As Integer
    dmColor            As Integer
    dmDuplex           As Integer
    dmYResolution      As Integer
    dmTTOption         As Integer
    dmCollate          As Integer
    dmFormName         As String * CCFORMNAME
    dmUnusedPadding    As Integer
    dmBitsPerPel       As Integer
    dmPelsWidth        As Long
    dmPelsHeight       As Long
    dmDisplayFlags     As Long
    dmDisplayFrequency As Long
End Type

Declare Function EnumDisplaySettings Lib "user32" Alias "EnumDisplaySettingsA" _
    (ByVal lpszDeviceName As Long, ByVal iModeNum As Long, lptypDevMode As Any) As Boolean
Declare Function ChangeDisplaySettings Lib "user32" Alias "ChangeDisplaySettingsA" _
    (lptypDevMode As Any, ByVal dwFlags As Long) As Long
Declare Function ExitWindowsEx Lib "user32" _
    (ByVal uFlags As Long, ByVal dwReserved As Long) As Long
Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
    (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Declare Function GetSystemMetrics Lib "user32.dll" (ByVal nIndex As Long) As Long

'#API Call to get username
Declare Function W32GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, _
    nSize As Long) As Long
'#API to get host name
Declare Function W32GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
'# Structure, API call to get free memory (for resource logging)

Type LOG_MEMORY_STATUS
    dwLength As Long
    dwMemoryLoad As Long
    dwTotalPhys As Long
    dwAvailPhys As Long
    dwTotalPageFile As Long
    dwAvailPageFile As Long
    dwTotalVirtual As Long
    dwAvailVirtual As Long
End Type

Declare Sub GlobalMemoryStatus Lib "kernel32" (lpBuffer As LOG_MEMORY_STATUS)
Declare Function GetCurrentProcessId Lib "kernel32" () As Long


