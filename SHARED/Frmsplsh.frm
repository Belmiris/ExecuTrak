VERSION 5.00
Object = "{C75015E0-2232-11D3-B440-0060971E99AF}#1.0#0"; "FACTFRM.OCX"
Begin VB.Form frmSplash 
   Appearance      =   0  'Flat
   BackColor       =   &H00800000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Select Data Sources"
   ClientHeight    =   2835
   ClientLeft      =   1545
   ClientTop       =   2625
   ClientWidth     =   7545
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   10.5
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00800000&
   Icon            =   "Frmsplsh.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2835
   ScaleWidth      =   7545
   Begin FACTFRMLib.FactorFrame FactorFrame1 
      Height          =   2844
      Left            =   0
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   0
      Width           =   7548
      _Version        =   65536
      _ExtentX        =   13314
      _ExtentY        =   5016
      _StockProps     =   77
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   10.48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelOuter      =   0
      BorderWidth     =   0
      TitleBarHeight  =   24
      BeginProperty PanelFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.PictureBox picMain 
         Height          =   2604
         Left            =   132
         ScaleHeight     =   2550
         ScaleWidth      =   5670
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   120
         Width           =   5724
         Begin VB.ComboBox cmbDataSet 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   10.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   324
            Left            =   1860
            TabIndex        =   0
            Top             =   180
            Width           =   1992
         End
         Begin VB.TextBox txtDatabase 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   10.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   336
            Left            =   1860
            TabIndex        =   1
            Top             =   636
            Width           =   1995
         End
         Begin VB.TextBox txtHost 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   10.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   336
            Left            =   1860
            TabIndex        =   2
            Top             =   1056
            Width           =   1995
         End
         Begin VB.TextBox txtUserName 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   10.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   336
            Left            =   1860
            TabIndex        =   3
            Top             =   1476
            Width           =   1995
         End
         Begin VB.TextBox txtPassword 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   10.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   336
            IMEMode         =   3  'DISABLE
            Left            =   1860
            PasswordChar    =   "*"
            TabIndex        =   4
            Top             =   1896
            Width           =   1995
         End
         Begin VB.Image Image1 
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Height          =   1932
            Left            =   4188
            Picture         =   "Frmsplsh.frx":030A
            Stretch         =   -1  'True
            Top             =   252
            Width           =   1404
         End
         Begin VB.Label lblStatic 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Caption         =   "Data Set Name :"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   10.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   312
            Index           =   0
            Left            =   144
            TabIndex        =   14
            Top             =   228
            Width           =   1776
         End
         Begin VB.Label lblStatic 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Caption         =   "Password :"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   10.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   312
            Index           =   1
            Left            =   156
            TabIndex        =   13
            Top             =   1956
            Width           =   1308
         End
         Begin VB.Label lblStatic 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Caption         =   "User Name :"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   10.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   312
            Index           =   4
            Left            =   144
            TabIndex        =   12
            Top             =   1524
            Width           =   1428
         End
         Begin VB.Label lblStatic 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Caption         =   "Host Name :"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   10.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   312
            Index           =   3
            Left            =   144
            TabIndex        =   11
            Top             =   1092
            Width           =   1416
         End
         Begin VB.Label lblStatic 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Caption         =   "Database Name :"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   10.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   312
            Index           =   2
            Left            =   144
            TabIndex        =   10
            Top             =   672
            Width           =   1704
         End
      End
      Begin VB.CommandButton btnOK 
         Caption         =   "O&K"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   6084
         TabIndex        =   5
         Top             =   516
         Width           =   1308
      End
      Begin VB.CommandButton btnCancel 
         Caption         =   "&Cancel"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   6084
         TabIndex        =   6
         Top             =   1236
         Width           =   1308
      End
      Begin VB.CommandButton btnHelp 
         Caption         =   "&Help"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   6084
         TabIndex        =   7
         Top             =   1980
         Width           =   1308
      End
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Public Sub main()
'    If Command = t_szHandShake Then
'        Load frmMain "Name of the Main Form"
'    Else
'        frmSplash.Caption = "Select Data Sources"
'        frmSplash.Show vbModal
'    End If
'End Sub
'Public Sub subShowMainForm()
'    frmMain.Show
'End Sub

Option Explicit

Private Const SECTION_INI_DS = "ODBC 32 bit Data Sources"
Private Const SECTION_REG_DS = "ODBC Data Sources"
Private Const PARM_HOST = "HostName"
Private Const PARM_SERVERNAME = "ServerName"
Private Const PARM_DATABASE = "Database"
Private Const PARM_USERID = "LogonID"
Private Const PARM_USERID2 = "UID"
Private Const PARM_SERVICE = "Service"
Private Const PARM_PROTOCOL = "Protocol"
Private Const PARM_YIELDPROC = "YieldProc"
Private Const PARM_CB = "CursorBehavior"

Private Const FACTOR_REGISTER = "Software\Factor\ExecTrak\"
Private Const szODBC_REG_KEY1 = ".Default\Software\ODBC\ODBC.INI\"
Private Const szODBC_REG_KEY2 = "Software\ODBC\ODBC.INI\"
Private Const HKEY_CLASSES_ROOT = &H80000000
Private Const HKEY_CURRENT_USER = &H80000001
Private Const HKEY_LOCAL_MACHINE = &H80000002
Private Const HKEY_USERS = &H80000003
Private Const KEY_QUERY_VALUE = &H1
Private Const KEY_ENUMERATE_SUB_KEYS = &H8
Private Const KEY_NOTIFY = &H10
Private Const READ_CONTROL = &H20000
Private Const STANDARD_RIGHTS_READ = (READ_CONTROL)
Private Const SYNCHRONIZE = &H100000
Private Const KEY_ALL_ACCESS = &H3F
Private Const KEY_READ = ((STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) And (Not SYNCHRONIZE))

Private Const REG_OPTION_NON_VOLATILE = 0
Private Const REG_SZ As Long = 1
Private Const REG_DWORD As Long = 4
Private Const REG_OPTION_VOLATILE = 1           ' Key is not preserved when system is rebooted

Private Const INFORMIX_DATA_SOURCE = "INFORMIX"
Private Const SECURITY_DATABASE = "XTRACKSECURITY"
Private Const szSECURITY As String = "XTrackSecurity"

Private Const HELP_CONTENTS = &H3

Private Const LINE_COLOR1 = &HFFFFFF
Private Const LINE_COLOR2 = &H808080
Private Const LINE_COLOR3 = 0

Private Const ERROR_NONE = 0
Private Const ERROR_BADDB = 1
Private Const ERROR_BADKEY = 2
Private Const ERROR_CANTOPEN = 3
Private Const ERROR_CANTREAD = 4
Private Const ERROR_CANTWRITE = 5
Private Const ERROR_OUTOFMEMORY = 6
Private Const ERROR_INVALID_PARAMETER = 7
Private Const ERROR_ACCESS_DENIED = 8
Private Const ERROR_INVALID_PARAMETERS = 87
Private Const ERROR_NO_MORE_ITEMS = 259

Private Const PROCESS_ALL_ACCESS = &H1F0FFF
Private Const STILL_ACTIVE = &H103    'Not sure

'ODBC API function declarations and constants
'
Private Const SQL_ERROR = -1
Private Const SQL_INVALID_HANDLE = -2
Private Const SQL_NO_DATA_FOUND = 100
Private Const SQL_SUCCESS = 0
Private Const SQL_SUCCESS_WITH_INFO = 1

Private Const SQL_FETCH_NEXT = 1
Private Const SQL_FETCH_FIRST = 2
Private Const SQL_FETCH_LAST = 3
Private Const SQL_FETCH_PRIOR = 4
Private Const SQL_FETCH_ABSOLUTE = 5
Private Const SQL_FETCH_RELATIVE = 6
Private Const SQL_FETCH_BOOKMARK = 8

Private Declare Function RegCloseKey Lib "advapi32.dll" _
    (ByVal hKey As Long) As Long

Private Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" _
    (ByVal hKey As Long, _
     ByVal lpSubKey As String, _
     ByVal Reserved As Long, _
     ByVal lpClass As String, _
     ByVal dwOptions As Long, _
     ByVal samDesired As Long, _
     lpSecurityAttributes As Long, _
     phkResult As Long, lpdwDisposition As Long) As Long

Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" _
    (ByVal hKey As Long, _
     ByVal lpSubKey As String, _
     ByVal ulOptions As Long, _
     ByVal samDesired As Long, _
     phkResult As Long) As Long

Private Declare Function RegQueryValueExString Lib "advapi32.dll" Alias "RegQueryValueExA" _
    (ByVal hKey As Long, _
     ByVal lpValueName As String, _
     ByVal lpReserved As Long, _
     lpType As Long, _
     ByVal lpData As String, _
     lpcbData As Long) As Long

Private Declare Function RegQueryValueExLong Lib "advapi32.dll" Alias "RegQueryValueExA" _
    (ByVal hKey As Long, _
     ByVal lpValueName As String, _
     ByVal lpReserved As Long, _
     lpType As Long, _
     lpData As Long, _
     lpcbData As Long) As Long

Private Declare Function RegQueryValueExNULL Lib "advapi32.dll" Alias "RegQueryValueExA" _
    (ByVal hKey As Long, _
     ByVal lpValueName As String, _
     ByVal lpReserved As Long, _
     lpType As Long, _
     ByVal lpData As Long, _
     lpcbData As Long) As Long

Private Declare Function RegEnumKey Lib "advapi32.dll" Alias "RegEnumKeyA" ( _
    ByVal hKey As Long, _
    ByVal dwIndex As Long, _
    ByVal lpName As String, _
    ByVal cbName As Long) As Long
    
Private Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" ( _
    ByVal hKey As Long, _
    ByVal dwIndex As Long, _
    ByVal lpValueName As String, _
    lpcbValueName As Long, _
    lpReserved As Long, _
    lpType As Long, _
    lpData As Any, _
    lpcbData As Long) As Long
    
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" _
    (ByVal hKey As Long, _
     ByVal lpValueName As String, _
     ByVal Reserved As Long, _
     ByVal dwType As Long, _
     ByVal lpData As String, _
     ByVal cbData As Long) As Long         ' Note that if you declare the lpData parameter as String, you must pass it By Value.

Private Declare Function GetPrivateProfileSection Lib "kernel32" Alias "GetPrivateProfileStringA" ( _
    ByVal lpApplicationName As String, _
    ByVal lpKeyName As Long, _
    ByVal lpDefault As String, _
    ByVal lpReturnedString As String, _
    ByVal nSize As Long, _
    ByVal lpFileName As String) As Long

Private Declare Function fnRegisterDll Lib "REGDLL.DLL" _
    Alias "RegisterDLL" (ByVal sPathName As String) As Long

Private Declare Function SQLAllocConnect Lib "odbc32.dll" (ByVal henv&, phdbc&) As Integer
Private Declare Function SQLAllocEnv Lib "odbc32.dll" (phenv&) As Integer
Private Declare Function SQLAllocStmt Lib "odbc32.dll" (ByVal hdbc&, phstmt&) As Integer
Private Declare Function SQLFetch Lib "odbc32.dll" (ByVal hstmt&) As Integer
Private Declare Function SQLFreeConnect Lib "odbc32.dll" (ByVal hdbc&) As Integer
Private Declare Function SQLFreeEnv Lib "odbc32.dll" (ByVal henv&) As Integer
Private Declare Function SQLFreeStmt Lib "odbc32.dll" (ByVal hstmt&, ByVal fOption%) As Integer
Private Declare Function SQLDrivers Lib "odbc32.dll" (ByVal henv&, ByVal fDirection%, ByVal szDriverDesc$, ByVal cbDriverDescMax%, pcbDriverDesc%, ByVal szDriverAttr$, ByVal cbDrvrAttrMax%, pcbDrvrAttr%) As Integer
Private Declare Function SQLDataSources Lib "odbc32.dll" (ByVal henv&, ByVal fDirection%, ByVal szDSN$, ByVal cbDSNMax%, pcbDSN%, ByVal szDescription$, ByVal cbDescriptionMax%, pcbDescription%) As Integer

'david 10/30/00
Private Const szODBC_DATABASE = "Database"
Private Const szODBC_HOST = "HostName"
Private Const szODBC_SERVERNAME = "ServerName"
Private Const szODBC_SERVICE = "Service"
Private Const szODBC_YIELDPROC = "YieldProc"
Private Const szODBC_CB = "CursorBehavior"
Private Const szODBC_PROTOCOL = "Protocol"
Private Const szODBC_DRIVER = "Driver"

'key name for Informix Driver
Private Const szODBC_SERVERNAME2 = "Server"
Private Const szODBC_CLIENT_LOCALE = "CLIENT_LOCALE"
Private Const szODBC_DB_LOCALE = "DB_LOCALE"
Private Const szODBC_VMBCHARLENEXACT = "VMBCHARLENEXACT"
Private Const szODBC_ENABLESCROLLABLECURSORS = "ENABLESCROLLABLECURSORS"
Private Const szODBC_ENABLEINSERTCURSORS = "ENABLEINSERTCURSORS"
Private Const szODBC_OPTIMIZEAUTOCOMMIT = "OPTIMIZEAUTOCOMMIT"
Private Const szODBC_OPTOFC = "OPTOFC"
Private Const szODBC_REPORTKEYSETCURSORS = "REPORTKEYSETCURSORS"
Private Const szODBC_NEEDODBCTYPESONLY = "NEEDODBCTYPESONLY"
Private Const szODBC_FETCHBUFFERSIZE = "FETCHBUFFERSIZE"

Private Const gszSPACE As String = " "
Private Const szDRIVER_DESCRIPTION As String = "INFORMIX"
Private Const szHENV_ERROR As String = "Cannot Allocation Environment Handle"
Private Const gszCOMMA As String = ","
Private Const szFORM_NAME As String = "FRMSPLASH.FRM"
Private Const gszMODULE_ERROR As String = "Module Error"

Private colDrivers As Collection

Private m_sDSN As String
Private m_sUID As String
Private m_sPWD As String
Private m_sHost As String
Private m_sDriver As String

Private m_sODBC_INI_Path As String
Private m_lODBC_INI_Key As Long
'
'david 04/04/2001
Private m_bAutoConnect As Boolean
Private m_sConnectionError As String
'

Private Function fnAllBoxFilled() As Boolean

    fnAllBoxFilled = True
    If fnNeedFocus(txtDatabase) Then
        fnAllBoxFilled = False
    Else
        If fnNeedFocus(txtHost) Then
            fnAllBoxFilled = False
        Else
            If fnNeedFocus(txtUserName) Then
                fnAllBoxFilled = False
            Else
                If fnNeedFocus(txtPassword) Then
                    fnAllBoxFilled = False
                End If
            End If
        End If
    End If
End Function

Private Function fnParentDir(sCurr As String) As String

    Dim i As Integer
    
    i = Len(sCurr)
    Do While i > 0
        If Mid(sCurr, i, 1) = "\" Then
            Exit Do
        End If
        i = i - 1
    Loop
    If i > 0 Then
        fnParentDir = Left(sCurr, i)
    Else
        fnParentDir = ""
    End If
End Function

Private Sub subGetDSN_INFO(sDSN As String, sDatabase As String, _
                           sHost As String, _
                           sUserID As String)
    
    Dim sODBCKey As String
    
    If Not fnSetODBCINIPath(sDSN) Then
        Exit Sub
    End If
    
    sODBCKey = m_sODBC_INI_Path & sDSN
    
    sDatabase = QueryValue(m_lODBC_INI_Key, sODBCKey, szODBC_DATABASE)
    
    sHost = QueryValue(m_lODBC_INI_Key, sODBCKey, szODBC_SERVERNAME)
    If Trim(sHost) = "" Then
        sHost = QueryValue(m_lODBC_INI_Key, sODBCKey, szODBC_HOST)
    End If
    If Trim(sHost) = "" Then
        sHost = QueryValue(m_lODBC_INI_Key, sODBCKey, szODBC_SERVERNAME2)
    End If
    
    sUserID = QueryValue(m_lODBC_INI_Key, sODBCKey, PARM_USERID)
    If Trim(sUserID) = "" Then
        sUserID = QueryValue(m_lODBC_INI_Key, sODBCKey, PARM_USERID2)
    End If
End Sub

Private Function QueryValue(ByVal lKey As Long, _
                           sKeyName As String, _
                           sValueName As String) As String
    Dim lRetVal As Long         'result of the API functions
    Dim hKey As Long         'handle of opened key
    Dim vValue As Variant      'setting of queried value

    QueryValue = ""
    lRetVal = RegOpenKeyEx(lKey, sKeyName, 0, KEY_ALL_ACCESS, hKey)
    If lRetVal = 0 Then
        lRetVal = QueryValueEx(hKey, sValueName, vValue)
        If lRetVal = 0 Then
            QueryValue = vValue
        End If
        RegCloseKey (hKey)
    End If
End Function


Public Function RegisterDll(sPathName As String, bCheck As Boolean) As Boolean
    Const LOCAL_PATH = "C:\FACTOR\OLE\"

    Dim sFileDateTime1 As String
    Dim lFileSize1 As Long
    Dim sFileDateTime2 As String
    Dim lFileSize2 As Long
    Dim sLoaclPathName As String
    
    On Error GoTo errRegDll
    If fnPreparePath(LOCAL_PATH) Then
        If bCheck Then
            sLoaclPathName = LOCAL_PATH & fnExtractName(sPathName, True)
            lFileSize1 = FileLen(sPathName)
            sFileDateTime1 = Trim(FileDateTime(sPathName))
            If fnIsFile(sLoaclPathName) Then
                lFileSize2 = FileLen(sLoaclPathName)
                sFileDateTime2 = Trim(FileDateTime(sLoaclPathName))
                If lFileSize1 = lFileSize2 Then
                    If sFileDateTime1 = sFileDateTime2 Then
                        RegisterDll = True
                        Exit Function
                    End If
                End If
            End If
        End If
        FileCopy sPathName, sLoaclPathName
        If fnRegisterDll(sLoaclPathName) = 0 Then
            RegisterDll = True
        Else
            RegisterDll = False
        End If
    End If
    Exit Function
errRegDll:
End Function

Private Function fnIsPath(sPath As String) As Boolean

    On Error Resume Next
    ChDir sPath
    If Err.Number > 0 Then
        fnIsPath = False
    Else
        fnIsPath = True
    End If
End Function

Private Function fnIsFile(ByVal szFilename As String) As Boolean
    
    On Error GoTo errNotFile

    fnIsFile = False
    If InStr(szFilename, "?") > 0 Then
        Exit Function
    End If
    If InStr(szFilename, "*") > 0 Then
        Exit Function
    End If
    If Trim(szFilename) <> "" Then
        Open szFilename For Input As #29
        Close #29
        fnIsFile = True
    End If
    Exit Function
errNotFile:
    #If DEVELOP Then
        MsgBox "Error # " & Err.Number & vbCrLf & "Error Message: " & Err.Description & " - " & szFilename
    #End If
End Function

Private Function fnPreparePath(sOrigPath As String) As Boolean

    Dim sDirs() As String
    Dim sPath As String
    Dim i As Integer
    Dim i1 As Integer
    
    subParseString sDirs, sOrigPath, "\"
    If Right(sDirs(0), 1) = ":" Then
        i1 = 1
        sDirs(1) = sDirs(0) & "\" & sDirs(1)
    Else
        i1 = 0
    End If
    fnPreparePath = False
    On Error Resume Next
    sPath = ""
    For i = i1 To UBound(sDirs)
        If i = i1 Then
            sPath = sDirs(i)
        Else
            sPath = sPath & "\" & sDirs(i)
        End If
        If Not fnIsPath(sPath) Then
            Err.Clear
            MkDir sPath
            If Err.Number <> 0 Then
                Exit Function
            End If
        End If
    Next i
    fnPreparePath = True
End Function

Private Function RegSetValue(ByVal lKey As Long, _
                            sKeyName As String, _
                            sValueName As String, _
                            ByVal lType As Long, _
                            ByVal sValue As String) As Boolean

    Dim lRetVal As Long         'result of the API functions
    Dim hKey As Long         'handle of opened key
    Dim vValue As String      'setting of queried value
    Dim cbData As Long
    
    RegSetValue = False
'    lRetVal = RegOpenKeyEx(lKey, sKeyName, 0, KEY_ALL_ACCESS, hKey)
    lRetVal = RegCreateKeyEx(lKey, sKeyName, 0, 0, REG_OPTION_VOLATILE, KEY_ALL_ACCESS, 0, hKey, cbData)
    If lRetVal = 0 Then
        If lType >= 0 Then
            If lType = REG_DWORD Then
                vValue = Chr(sValue)
                cbData = 4
            Else
                vValue = sValue & Chr(0)
                cbData = Len(vValue)
            End If
            lRetVal = RegSetValueEx(hKey, sValueName, 0, lType, vValue, cbData)
            If lRetVal = 0 Then
                RegSetValue = True
            End If
        End If
        RegCloseKey (hKey)
    End If
End Function

Private Function fnExtractName(sFile As String, _
                               bIncludeExt As Boolean) As String
    
    Dim nPos As Integer
    Dim sTemp As String
    
    nPos = Len(sFile)
    Do While nPos > 0
        If Mid(sFile, nPos, 1) = "\" Then
            Exit Do
        End If
        nPos = nPos - 1
    Loop
    nPos = Len(sFile) - nPos
    If nPos > 0 Then
        sTemp = Right(sFile, nPos)
    End If
    If bIncludeExt Then
        fnExtractName = sTemp
    Else
        nPos = Len(sTemp)
        Do While nPos > 0
            If Mid(sTemp, nPos, 1) = "." Then
                Exit Do
            End If
            nPos = nPos - 1
        Loop
        If nPos > 1 Then
            fnExtractName = Left(sTemp, nPos - 1)
        Else
            fnExtractName = sTemp
        End If
    End If
End Function

Private Function fnNeedFocus(txtBox As Textbox) As Boolean
    If Trim(txtBox.Text) = "" Then
        subSetFocus txtBox
        fnNeedFocus = True
    Else
        fnNeedFocus = False
    End If
End Function

Private Sub subSelectText(txtBox As Textbox)

    txtBox.SelStart = 0
    txtBox.SelLength = Len(txtBox.Text)
    
End Sub

Private Sub subSetToNextBox(txtBox As Control)
 
    Select Case txtBox.TabIndex
        Case cmbDataSet.TabIndex
            If Not fnNeedFocus(txtDatabase) Then
                If Not fnNeedFocus(txtHost) Then
                    If Not fnNeedFocus(txtUserName) Then
                        subSetFocus txtPassword
                    End If
                End If
            End If
            'subSetFocus txtDatabase
        Case txtDatabase.TabIndex
            subSetFocus txtHost
        Case txtHost.TabIndex
            subSetFocus txtUserName
        Case txtUserName.TabIndex
            subSetFocus txtPassword
        Case txtPassword.TabIndex
            If fnAllBoxFilled Then
                btnOK_Click
            Else
                subSetFocus btnOK
            End If
    End Select
    
End Sub

Private Sub subSetFocus(cntlTemp As Control, ParamArray arryControls() As Variant)
    'Set focus to a textbox or a command button control

    Const nTrialNumber As Integer = 1
    Dim nCount As Integer

    nCount = 0
    On Error GoTo errSetFocus
    cntlTemp.SetFocus
    Exit Sub
tryNext:
    On Error GoTo errNext
    arryControls(nCount).SetFocus
extSetFocus:
    On Error GoTo 0
    Exit Sub
    
errSetFocus:
    If nCount < nTrialNumber Then
        nCount = nCount + 1
        DoEvents
        Resume
    Else
        nCount = 0
        Resume tryNext
    End If
errNext:
    If nCount < UBound(arryControls) Then
        nCount = nCount + 1
        Resume tryNext
    Else
        Resume extSetFocus
    End If
End Sub

Private Sub subParseString(sParam() As String, _
                           sSrc As String, _
                           sDelim As String, _
                           Optional vStart As Variant, _
                           Optional vEnd As Variant)
                          
    If Trim(sSrc) = "" Then
        Exit Sub
    End If

    Const nArrayInc As Integer = 5
    Dim i1 As Integer
    Dim i2 As Integer
    Dim k As Integer
    Dim nEnd As Integer
    Dim sTemp As String
    
    If IsMissing(vStart) Then
        i1 = 1
    Else
        i1 = vStart
    End If
    If IsMissing(vEnd) Then
        nEnd = Len(sSrc)
    Else
        nEnd = vEnd
    End If
    If i1 < 1 Then i1 = 1
    i2 = 1
    k = 0
    ReDim sParam(nArrayInc)
    While i1 <= nEnd And i2 > 0 And i2 <= nEnd
        i2 = InStr(i1, sSrc, sDelim)
        If i2 >= i1 And i2 <= nEnd Then
            If k > UBound(sParam) Then
                ReDim Preserve sParam(k + nArrayInc)
            End If
            sTemp = Mid$(sSrc, i1, i2 - i1)
            If sTemp <> "" Or sDelim <> " " Then
                sParam(k) = sTemp
                k = k + 1
            End If
            i1 = i2 + 1
        End If
    Wend
    If i2 <= nEnd Then
        If k > UBound(sParam) Then
            ReDim Preserve sParam(k + nArrayInc)
        End If
        sParam(k) = Trim$(Mid$(sSrc, i1, nEnd - i1 + 1))
        ReDim Preserve sParam(k)
    Else
        If k > 0 Then
            sParam(k - 1) = Trim$(Mid$(sSrc, i1, nEnd - i1 + 1))
            ReDim Preserve sParam(k - 1)
        End If
    End If
End Sub

Private Function QueryValueEx(ByVal lhKey As Long, _
                      ByVal szValueName As String, _
                      vValue As Variant) As Long
    Dim cch As Long
    Dim lrc As Long
    Dim lType As Long
    Dim lValue As Long
    Dim sValue As String

    On Error GoTo QueryValueExError

    ' Determine the size and type of data to be read
    lrc = RegQueryValueExNULL(lhKey, szValueName, 0&, lType, 0&, cch)
    If lrc = ERROR_NONE Then
        Select Case lType
            ' For strings
            Case REG_SZ:
                sValue = String(cch, 0)
                lrc = RegQueryValueExString(lhKey, szValueName, 0&, lType, sValue, cch)
                If lrc = ERROR_NONE Then
                    vValue = fnSZStr2Str(sValue)
                Else
                    vValue = Empty
                End If
            ' For DWORDS
            Case REG_DWORD:
                lrc = RegQueryValueExLong(lhKey, szValueName, 0&, lType, lValue, cch)
                If lrc = ERROR_NONE Then
                    vValue = lValue
                End If
            Case Else
                'all other data types not supported
                lrc = -1
        End Select
    End If
QueryValueExExit:
    QueryValueEx = lrc
    Exit Function
QueryValueExError:
    Resume QueryValueExExit
End Function

Private Function fnSZStr2Str(szStr As String) As String

    Dim nPos As Integer
    
    nPos = InStr(szStr, Chr(0))
    If nPos > 0 Then
        fnSZStr2Str = Left(szStr, nPos - 1)
    Else
        fnSZStr2Str = szStr
    End If
End Function

Private Sub btnCancel_Click()
    End
End Sub

Private Sub btnHelp_Click()
    WinHelp Me.hwnd, szHelpFileName, HELP_CONTENTS, CLng(0)
End Sub

Private Sub btnOK_Click()
    Dim sPWD As String
    
    Screen.MousePointer = vbHourglass
    
    sPWD = txtPassword.Text
    If sPWD = "" Then
        sPWD = "fakePPP"
    End If
    
    On Error GoTo errTrap
    
    m_sDSN = cmbDataSet.Text
    m_sDriver = colDrivers(m_sDSN)
    m_sUID = txtUserName.Text
    m_sPWD = sPWD
    
    If Not fnSetODBCINIPath(m_sDSN) Then
        m_sConnectionError = "Cannot find ODBC.INI in the Windows registry. Please report this message to support."
        
        If Not m_bAutoConnect Then
            subCriticalMsg m_sConnectionError, szFORM_NAME
        End If
        
        Exit Sub
    End If
    
    t_szConnect = fnConnectString(m_sDSN)
    
    Me.Hide
    
    subShowMainForm
    
    Exit Sub
    
errTrap:
    Screen.MousePointer = vbDefault
    
    If Err.Number = 5 Then
        m_sConnectionError = "Data Source Name is not valid."
        
        If Not m_bAutoConnect Then
            subCriticalMsg m_sConnectionError, szFORM_NAME
            subSetFocus cmbDataSet
        End If
    Else
        m_sConnectionError = "An error has occurred." + vbCrLf + vbCrLf + "Error Code: " & _
            Err.Number & vbCrLf & "Error Desc: " + Err.Description + "."
        
        If Not m_bAutoConnect Then
            subCriticalMsg m_sConnectionError + vbCrLf + vbCrLf + _
                "Please report this message to support.", szFORM_NAME
            subSetFocus txtPassword
        End If
    End If
End Sub

Private Sub cmbDataSet_Click()

    Dim sDSN As String
    Dim sDatabase As String
    Dim sHost As String
    Dim sUserID As String
    
    sDSN = Trim(cmbDataSet.Text)
    
    subGetDSN_INFO sDSN, sDatabase, sHost, sUserID
    
    txtDatabase.Text = Trim(sDatabase)
    txtHost.Text = Trim(sHost)
    txtUserName.Text = Trim(sUserID)
    
    If Not ActiveControl Is cmbDataSet Then Exit Sub
    
End Sub

Private Sub cmbDataSet_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        subSetToNextBox cmbDataSet
        KeyAscii = 0
    End If
End Sub

Private Sub Form_Activate()
    On Error Resume Next
    cmbDataSet.SetFocus
End Sub

Private Sub Form_Load()
    tfnDisableFormSystemClose Me
    tfnCenterForm Me
    If fnGetDataSources(cmbDataSet) = 0 Then
        MsgBox "At least one Data Source Name needs to be created to run the program.", vbExclamation
        End
    End If
End Sub

Private Sub picMain_Paint()
    subMakeVSLookFrame picMain
End Sub


Private Sub txtDatabase_GotFocus()
    subSelectText txtDatabase
End Sub

Private Sub txtDatabase_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        subSetToNextBox txtDatabase
        KeyAscii = 0
    End If
End Sub

Private Sub txtHost_GotFocus()
    subSelectText txtHost
End Sub

Private Sub txtHost_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        subSetToNextBox txtHost
        KeyAscii = 0
    End If
End Sub


Private Sub txtPassword_GotFocus()
    subSelectText txtPassword
End Sub

Private Sub txtPassword_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        subSetToNextBox txtPassword
        KeyAscii = 0
    End If

End Sub

Private Sub txtUserName_GotFocus()
    subSelectText txtUserName
End Sub

Private Sub txtUserName_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        subSetToNextBox txtUserName
        KeyAscii = 0
    End If
End Sub

Private Sub subMakeVSLookFrame(picFrame As PictureBox)
    
    Dim X1 As Integer
    Dim Y1 As Integer
    Dim X2 As Integer
    Dim Y2 As Integer
    
    picFrame.BorderStyle = 0
    picFrame.DrawStyle = vbSolid
    X1 = 0  ' picFrame.Left
    X2 = picFrame.ScaleWidth - Screen.TwipsPerPixelX
    Y1 = 0  'picFrame.Top
    Y2 = picFrame.ScaleHeight - Screen.TwipsPerPixelY
    picFrame.Line (X1, Y1)-(X2, Y1), LINE_COLOR1
    picFrame.Line -(X2, Y2), LINE_COLOR2
    picFrame.Line -(X1, Y2), LINE_COLOR2
    picFrame.Line (X1, Y2 - Screen.TwipsPerPixelY)-(X1, Y1), LINE_COLOR1
    X1 = X1 + Screen.TwipsPerPixelX
    X2 = X2 - Screen.TwipsPerPixelX
    Y1 = Y1 + Screen.TwipsPerPixelY
    Y2 = Y2 - Screen.TwipsPerPixelY
    picFrame.Line (X1, Y1)-(X2, Y1), LINE_COLOR2
    picFrame.Line -(X2, Y2), LINE_COLOR1
    picFrame.Line -(X1 - 2 * Screen.TwipsPerPixelX, Y2), LINE_COLOR1
    picFrame.Line (X1, Y2 - 2 * Screen.TwipsPerPixelY)-(X1, Y1), LINE_COLOR2
End Sub

Public Function Connect(sDSN As String, _
                        sUID As String, _
                        sPWD As String, _
                        Optional sErrMsg As String = "") As Boolean
    
    m_bAutoConnect = True
    m_sConnectionError = ""
    
    cmbDataSet = sDSN
    cmbDataSet_Click
    txtUserName = sUID
    txtPassword = sPWD
    
    btnOK_Click
    
    If m_sConnectionError <> "" Then
        sErrMsg = m_sConnectionError
        Connect = False
    Else
        Connect = True
    End If
End Function

'
'This routine fills a list or combo box with all available
'ODBC Data Source Names (DSN's) found in ODBC.INI matching
'the szDRIVER_DESCRIPTION string defined above.
'Example.
'[ODBC Data Sources]
'factor=INTERSOLV INFORMIX5
'
'factor would be the DSN
'INTERSOLV INFORMIX5 would be the Driver Description
'
Private Function fnGetDataSources(plstObject As ComboBox) As Integer

    Dim szDataSourceName As String       'Data Source Name returned from SQL function call
    Dim szDriverDescription As String    'Driver Description returned from SQL function call
    Dim nDataSourceNameLen As Integer    'DSN string length returned from SQL function call
    Dim nDriverDescriptionLen As Integer 'Driver Description string length returned from SQL function call
    Dim nReturn As Integer               'return code from SQL call
    Dim lHenv As Long                    'handle to the environment

    Const DSN_LENGTH = 32          'Data Source Name (fixed) length passed to SQL function
    Const DRIVER_DESC_LENGTH = 255 'Driver Description (fixed) length passed to SQL function

    plstObject.Clear         'clear old list box entries
    Set colDrivers = New Collection 'clear drivers collection
    
    If SQLAllocEnv(lHenv) <> SQL_ERROR Then 'get valid environment handle

        szDataSourceName = String(DSN_LENGTH, gszSPACE)
        szDriverDescription = String(DRIVER_DESC_LENGTH, gszSPACE) 'set fixed length strings, pad with spaces

        nReturn = SQLDataSources(lHenv, SQL_FETCH_FIRST, szDataSourceName, DSN_LENGTH, nDataSourceNameLen, _
            szDriverDescription, DRIVER_DESC_LENGTH, nDriverDescriptionLen) 'get first DSN/Driver Description values

        While nReturn = SQL_SUCCESS Or nReturn = SQL_SUCCESS_WITH_INFO 'process if SQL function call OK

            szDriverDescription = Left(szDriverDescription, nDriverDescriptionLen) 'strip any spaces and --> terminating NULL
            szDataSourceName = Left(szDataSourceName, nDataSourceNameLen)

            If InStr(1, szDriverDescription, szDRIVER_DESCRIPTION) > 0 Then 'check for application Driver
                If Not szDataSourceName = szSECURITY Then 'don'y display security entry it its exists
                    plstObject.AddItem szDataSourceName   'add to DataSourceName ListBox if true
                    colDrivers.Add Item:=szDriverDescription, Key:=szDataSourceName 'save driver using DataSourceName as Key
                End If
            End If

            szDataSourceName = String(DSN_LENGTH, gszSPACE)
            szDriverDescription = String(DRIVER_DESC_LENGTH, gszSPACE) 're-initialized fixed length strings for next fetch

            nReturn = SQLDataSources(lHenv, SQL_FETCH_NEXT, szDataSourceName, DSN_LENGTH, nDataSourceNameLen, _
                szDriverDescription, DRIVER_DESC_LENGTH, nDriverDescriptionLen) 'get next DSN/Driver Description values

        Wend

        SQLFreeEnv (lHenv) 'free the environment handle

    Else
        subCriticalMsg szHENV_ERROR & gszCOMMA & szFORM_NAME, gszMODULE_ERROR
    End If
    If plstObject.ListCount > 0 Then
        plstObject.ListIndex = 0
    End If

    fnGetDataSources = plstObject.ListCount 'return the number of valid DSN's configured in ODBC.INI

End Function

Private Sub subCriticalMsg(sMsg As String, _
                          sCaption As String)

    MsgBox sMsg, vbOKOnly + vbCritical, sCaption
    
End Sub

'david 10/30/00
Private Function fnConnectString(sDSN As String) As String

    Dim sTemp As String
    Dim sODBCKey As String
    Dim sDatabase As String
    
    sODBCKey = m_sODBC_INI_Path & sDSN
    m_sHost = QueryValue(m_lODBC_INI_Key, sODBCKey, szODBC_SERVERNAME)
    If Trim(m_sHost) = "" Then
        m_sHost = QueryValue(m_lODBC_INI_Key, sODBCKey, szODBC_HOST)
    End If
    If Trim(m_sHost) = "" Then
        m_sHost = QueryValue(m_lODBC_INI_Key, sODBCKey, szODBC_SERVERNAME2)
    End If
    fnConnectString = "ODBC;DSN=" & sDSN & ";UID=" & m_sUID _
            & ";PWD=" & m_sPWD
    sDatabase = QueryValue(m_lODBC_INI_Key, sODBCKey, szODBC_DATABASE)
    fnConnectString = fnConnectString & ";DB=" & sDatabase & ";HOST=" & m_sHost
    sTemp = QueryValue(m_lODBC_INI_Key, sODBCKey, szODBC_SERVICE)
    fnConnectString = fnConnectString & ";SERV=" & sTemp
    sTemp = QueryValue(m_lODBC_INI_Key, sODBCKey, szODBC_YIELDPROC)
    fnConnectString = fnConnectString & ";YLD=" & sTemp
    sTemp = QueryValue(m_lODBC_INI_Key, sODBCKey, szODBC_CB)
    fnConnectString = fnConnectString & ";CB=" & sTemp
    sTemp = QueryValue(m_lODBC_INI_Key, sODBCKey, szODBC_PROTOCOL)
    fnConnectString = fnConnectString & ";PRO=" & sTemp
    
End Function

Private Function fnSetODBCINIPath(sDSN As String) As Boolean
    Dim sTemp As String
    
    m_sODBC_INI_Path = szODBC_REG_KEY2
    m_lODBC_INI_Key = HKEY_CURRENT_USER
    sTemp = QueryValue(m_lODBC_INI_Key, m_sODBC_INI_Path & sDSN, szODBC_DATABASE)
    If sTemp = "" Then
        m_sODBC_INI_Path = szODBC_REG_KEY2
        m_lODBC_INI_Key = HKEY_LOCAL_MACHINE
        sTemp = QueryValue(m_lODBC_INI_Key, m_sODBC_INI_Path & sDSN, szODBC_DATABASE)
    End If
    If sTemp = "" Then
        m_sODBC_INI_Path = szODBC_REG_KEY1
        m_lODBC_INI_Key = HKEY_USERS
        sTemp = QueryValue(m_lODBC_INI_Key, m_sODBC_INI_Path & sDSN, szODBC_DATABASE)
    End If
    If sTemp = "" Then
        fnSetODBCINIPath = False
    Else
        fnSetODBCINIPath = True
    End If
End Function

Public Sub GetODBCINIPath(lODBCKey As Long, _
                           sODBCPath As String, _
                           sDSN As String)
    Dim sTemp As String
    
    sODBCPath = szODBC_REG_KEY2
    lODBCKey = HKEY_CURRENT_USER
    sTemp = QueryValue(lODBCKey, sODBCPath & sDSN, PARM_DATABASE)
    If sTemp = "" Then
        sODBCPath = szODBC_REG_KEY1
        lODBCKey = HKEY_USERS
        sTemp = QueryValue(lODBCKey, sODBCPath & sDSN, PARM_DATABASE)
        If sTemp = "" Then
            sODBCPath = szODBC_REG_KEY2
            lODBCKey = HKEY_LOCAL_MACHINE
            sTemp = QueryValue(lODBCKey, sODBCPath & sDSN, PARM_DATABASE)
        End If
    End If
End Sub

Public Function DBConnect(sDSN As String, sUID As String, sPWD As String, Optional sHost As String = "") As String
    
    m_sUID = sUID
    m_sPWD = sPWD
    
    DBConnect = fnConnectString(sDSN)
End Function

