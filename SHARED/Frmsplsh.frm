VERSION 5.00
Begin VB.Form frmSplash 
   Appearance      =   0  'Flat
   BackColor       =   &H00800000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Select Data Sources"
   ClientHeight    =   2832
   ClientLeft      =   1548
   ClientTop       =   2628
   ClientWidth     =   7536
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   10.8
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
   ScaleHeight     =   2832
   ScaleWidth      =   7536
   Begin VB.CommandButton btnHelp 
      Caption         =   "&Help"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   396
      Left            =   6096
      TabIndex        =   13
      Top             =   1968
      Width           =   1308
   End
   Begin VB.CommandButton btnCancel 
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   396
      Left            =   6096
      TabIndex        =   12
      Top             =   1224
      Width           =   1308
   End
   Begin VB.CommandButton btnOK 
      Caption         =   "O&K"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   396
      Left            =   6096
      TabIndex        =   11
      Top             =   504
      Width           =   1308
   End
   Begin VB.PictureBox picMain 
      Height          =   2604
      Left            =   144
      ScaleHeight     =   2556
      ScaleWidth      =   5676
      TabIndex        =   0
      Top             =   108
      Width           =   5724
      Begin VB.TextBox txtPassword 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.2
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
         TabIndex        =   5
         Top             =   1896
         Width           =   1995
      End
      Begin VB.TextBox txtUserName 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   336
         Left            =   1860
         TabIndex        =   4
         Top             =   1476
         Width           =   1995
      End
      Begin VB.TextBox txtHost 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   336
         Left            =   1860
         TabIndex        =   3
         Top             =   1056
         Width           =   1995
      End
      Begin VB.TextBox txtDatabase 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   336
         Left            =   1860
         TabIndex        =   2
         Top             =   636
         Width           =   1995
      End
      Begin VB.ComboBox cmbDataSet 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   324
         Left            =   1860
         TabIndex        =   1
         Top             =   180
         Width           =   1992
      End
      Begin VB.Label lblStatic 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Database Name :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.2
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
      Begin VB.Label lblStatic 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Host Name :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.2
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
         TabIndex        =   9
         Top             =   1092
         Width           =   1416
      End
      Begin VB.Label lblStatic 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "User Name :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.2
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
         TabIndex        =   8
         Top             =   1524
         Width           =   1428
      End
      Begin VB.Label lblStatic 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Password :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.2
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
         TabIndex        =   7
         Top             =   1956
         Width           =   1308
      End
      Begin VB.Label lblStatic 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Data Set Name :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.2
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
         TabIndex        =   6
         Top             =   228
         Width           =   1776
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
    Private Const PARM_SERVICE = "Service"
    Private Const PARM_PROTOCOL = "Protocol"
    Private Const PARM_YIELDPROC = "YieldProc"
    Private Const PARM_CB = "CursorBehavior"
    
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
    
    Private Declare Function RegCloseKey Lib "advapi32.dll" _
        (ByVal hKey As Long) As Long
    
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
        
    Private Declare Function GetPrivateProfileSection Lib "kernel32" Alias "GetPrivateProfileStringA" ( _
        ByVal lpApplicationName As String, _
        ByVal lpKeyName As Long, _
        ByVal lpDefault As String, _
        ByVal lpReturnedString As String, _
        ByVal nSize As Long, _
        ByVal lpFileName As String) As Long

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

Private Function fnDatabase(lODBCKey As Long, _
                            sODBCPath As String) As String
    fnDatabase = QueryValue(lODBCKey, sODBCPath, PARM_DATABASE)
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

Private Function fnUserID(lODBCKey As Long, _
                          sODBCPath As String) As String
    fnUserID = QueryValue(lODBCKey, sODBCPath, PARM_USERID)
End Function


Private Function fnHost(lODBCKey As Long, _
                        sODBCPath As String) As String
    
    fnHost = QueryValue(lODBCKey, sODBCPath, PARM_SERVERNAME)
    If Trim(fnHost) = "" Then
        fnHost = QueryValue(lODBCKey, sODBCPath, PARM_HOST)
    End If

End Function

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


Private Function fnNeedFocus(txtBox As TextBox) As Boolean
    If Trim(txtBox.Text) = "" Then
        subSetFocus txtBox
        fnNeedFocus = True
    Else
        fnNeedFocus = False
    End If
End Function

Private Sub subLoadDSSection(aryDS() As String, _
                             nCount As Integer, _
                             lKey As Long, _
                             sSubKey As String)
    Const ARY_INC = 5
    Const BUFFER_SIZE = 1024
    
    Dim lRet As Long
    Dim lKeyHandle As Long
    Dim lIndx As Long
    Dim sKeyName As String
    Dim sKeyValue As String
    Dim lDataLength As Long
    Dim lNameLength As Long
    Dim sData As String
    Dim sTemp As String
    
    lRet = RegOpenKeyEx(lKey, sSubKey, 0&, KEY_READ, lKeyHandle)
    If lRet = ERROR_NONE Then
        ReDim aryDS(ARY_INC)
        nCount = 0
        lIndx = 0
        Do
            lNameLength = BUFFER_SIZE
            lDataLength = BUFFER_SIZE
            sKeyName = Space(lNameLength)
            sKeyValue = Space(lDataLength)
            lRet = RegEnumValue(lKeyHandle, lIndx, sKeyName, lNameLength, 0&, REG_DWORD, ByVal sKeyValue, lDataLength)
            If lRet = ERROR_NONE Then
                sData = Left(sKeyValue, lDataLength)
                If InStr(sData, INFORMIX_DATA_SOURCE) > 0 Then
                    If UBound(aryDS) < nCount Then
                        ReDim Preserve aryDS(nCount + ARY_INC)
                    End If
                    sTemp = Left(sKeyName, lNameLength)
                    sData = QueryValue(lKey, fnParentDir(sSubKey) & sTemp, PARM_DATABASE)
                    If Trim(sData) <> "" Then
                        aryDS(nCount) = sTemp
                        nCount = nCount + 1
                    End If
                End If
            End If
            lIndx = lIndx + 1
        Loop While lRet = ERROR_NONE
        RegCloseKey lKeyHandle
    End If
End Sub

Private Sub subSelectText(txtBox As TextBox)

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

Private Sub subLoadDataSources()

    Dim aryBuffer() As String
    Dim nCount As Integer
    Dim i As Integer
    Dim j As Integer
    Dim bInList As Boolean
    
    subLoadDSSection aryBuffer, nCount, HKEY_LOCAL_MACHINE, szODBC_REG_KEY2 & SECTION_REG_DS
    For i = 0 To nCount - 1
        If UCase(aryBuffer(i)) <> SECURITY_DATABASE Then
            cmbDataSet.AddItem aryBuffer(i)
        End If
    Next i
    subLoadDSSection aryBuffer, nCount, HKEY_CURRENT_USER, szODBC_REG_KEY2 & SECTION_REG_DS
    For i = 0 To nCount - 1
        bInList = False
        For j = 0 To cmbDataSet.ListCount - 1
            If aryBuffer(i) = cmbDataSet.List(j) Then
                bInList = True
                Exit For
            End If
        Next j
        If Not bInList Then
            If UCase(aryBuffer(i)) <> SECURITY_DATABASE Then
                cmbDataSet.AddItem aryBuffer(i)
            End If
        End If
    Next i
    subLoadDSSection aryBuffer, nCount, HKEY_USERS, szODBC_REG_KEY1 & SECTION_REG_DS
    For i = 0 To nCount - 1
        bInList = False
        For j = 0 To cmbDataSet.ListCount - 1
            If aryBuffer(i) = cmbDataSet.List(j) Then
                bInList = True
                Exit For
            End If
        Next j
        If Not bInList Then
            If UCase(aryBuffer(i)) <> SECURITY_DATABASE Then
                cmbDataSet.AddItem aryBuffer(i)
            End If
        End If
    Next i
    If cmbDataSet.ListCount > 0 Then
        cmbDataSet.ListIndex = 0
    End If

    If Trim(cmbDataSet.Text) <> "" And cmbDataSet.ListCount = 1 Then
        subSetFocus txtPassword
    Else
        subSetFocus cmbDataSet
    End If
End Sub


Public Function DBConnect(sDsn As String, _
                             sUID As String, _
                             sPWD As String, _
                             Optional vHost As Variant) As String

    Dim sFile As String
    Dim sConnect As String
    Dim sTemp As String
    
    Dim sODBCPath As String
    Dim sODBCRoot As String
    Dim lODBCKey As Long
    Dim sDatabase As String
    Dim m_sHost As String
    
    GetODBCINIPath lODBCKey, sODBCRoot, sDsn
    sODBCPath = sODBCRoot & sDsn
    If IsMissing(vHost) Then
        m_sHost = fnHost(lODBCKey, sODBCPath)
    Else
        m_sHost = vHost
    End If
    If Trim(m_sHost) = "" Then
        m_sHost = txtHost.Text
    End If
    sConnect = "ODBC;DSN=" & sDsn & ";UID=" & sUID _
            & ";PWD=" & sPWD
    If Trim(sDsn) = "" Then
        sDatabase = txtDatabase.Text
    Else
        sDatabase = fnDatabase(lODBCKey, sODBCPath)
    End If
    sConnect = sConnect & ";DB=" & sDatabase & ";HOST=" & m_sHost
    sTemp = QueryValue(lODBCKey, sODBCPath, PARM_SERVICE)
    sConnect = sConnect & ";SERV=" & sTemp
    sTemp = QueryValue(lODBCKey, sODBCPath, PARM_YIELDPROC)
    If sTemp <> "" Then
        sConnect = sConnect & ";YLD=" & sTemp
    End If
    sTemp = QueryValue(lODBCKey, sODBCPath, PARM_CB)
    If sTemp <> "" Then
        sConnect = sConnect & ";CB=" & sTemp
    End If
    sTemp = QueryValue(lODBCKey, sODBCPath, PARM_PROTOCOL)
    If sTemp <> "" Then
        sConnect = sConnect & ";PRO=" & sTemp
    End If
    
    DBConnect = sConnect
End Function


Public Sub GetODBCINIPath(lODBCKey As Long, _
                           sODBCPath As String, _
                           sDsn As String)
    Dim sTemp As String
    
    sODBCPath = szODBC_REG_KEY2
    lODBCKey = HKEY_CURRENT_USER
    sTemp = QueryValue(lODBCKey, sODBCPath & sDsn, PARM_DATABASE)
    If sTemp = "" Then
        sODBCPath = szODBC_REG_KEY1
        lODBCKey = HKEY_USERS
        sTemp = QueryValue(lODBCKey, sODBCPath & sDsn, PARM_DATABASE)
        If sTemp = "" Then
            sODBCPath = szODBC_REG_KEY2
            lODBCKey = HKEY_LOCAL_MACHINE
            sTemp = QueryValue(lODBCKey, sODBCPath & sDsn, PARM_DATABASE)
        End If
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
    Me.Hide
    sPWD = txtPassword.Text
    If sPWD = "" Then
        sPWD = "fakePPP"
    End If
    t_szConnect = DBConnect(cmbDataSet.Text, txtUserName.Text, sPWD, txtHost.Text)
    subShowMainForm
End Sub

Private Sub cmbDataSet_Click()

    Dim lODBCKey As Long
    Dim sODBCPath As String
    Dim sODBCRoot As String
    Dim sDsn As String
    
    sDsn = Trim(cmbDataSet.Text)
    GetODBCINIPath lODBCKey, sODBCRoot, sDsn
    sODBCPath = sODBCRoot & sDsn
    
    txtDatabase.Text = fnDatabase(lODBCKey, sODBCPath)
    txtHost.Text = fnHost(lODBCKey, sODBCPath)
    txtUserName.Text = fnUserID(lODBCKey, sODBCPath)
    
    If Not ActiveControl Is cmbDataSet Then Exit Sub
        
'    If Not fnNeedFocus(txtDatabase) Then
'        If Not fnNeedFocus(txtHost) Then
'            If Not fnNeedFocus(txtUserName) Then
'                subSetFocus txtPassword
'            End If
'        End If
'    End If
    
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
    tfnCenterForm Me
    subLoadDataSources
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
