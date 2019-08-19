VERSION 5.00
Begin VB.Form frmSYFTPAPP 
   Caption         =   "FTP Launcher"
   ClientHeight    =   8310
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7230
   LinkTopic       =   "Form3"
   ScaleHeight     =   8310
   ScaleWidth      =   7230
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtHost 
      Height          =   375
      Left            =   1920
      TabIndex        =   9
      Text            =   "ether"
      Top             =   240
      Width           =   5055
   End
   Begin VB.TextBox txtFolder 
      Height          =   375
      Left            =   1920
      TabIndex        =   8
      Text            =   "/home/rick/tst"
      Top             =   720
      Width           =   5055
   End
   Begin VB.TextBox txtUser 
      Height          =   375
      Left            =   1920
      TabIndex        =   7
      Text            =   "ssfactor"
      Top             =   1200
      Width           =   5055
   End
   Begin VB.TextBox txtPassword 
      Height          =   375
      Left            =   1920
      TabIndex        =   6
      Text            =   "menus"
      Top             =   1680
      Width           =   5055
   End
   Begin VB.CheckBox chkSFTP 
      Caption         =   "Use SFTP"
      Height          =   495
      Left            =   1920
      TabIndex        =   5
      Top             =   2040
      Width           =   1215
   End
   Begin VB.TextBox txtResult 
      Height          =   3255
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   4
      Top             =   4200
      Width           =   6735
   End
   Begin VB.TextBox txtFile 
      Height          =   375
      Left            =   1920
      TabIndex        =   3
      Text            =   "C:\Temp\FTPTHIS.TXT"
      Top             =   2640
      Width           =   5055
   End
   Begin VB.TextBox txtWindow 
      Height          =   375
      Left            =   1920
      TabIndex        =   2
      Top             =   3600
      Width           =   5055
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "Send"
      Height          =   495
      Left            =   5760
      TabIndex        =   1
      Top             =   7680
      Width           =   1215
   End
   Begin VB.CheckBox chkResultWindow 
      Caption         =   "Show Results"
      Height          =   375
      Left            =   1920
      TabIndex        =   0
      Top             =   3120
      Value           =   1  'Checked
      Width           =   2175
   End
   Begin VB.Label Label2 
      Caption         =   "FTP Host Name:"
      Height          =   375
      Left            =   240
      TabIndex        =   15
      Top             =   240
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "FTP Folder:"
      Height          =   375
      Left            =   240
      TabIndex        =   14
      Top             =   720
      Width           =   1455
   End
   Begin VB.Label Label4 
      Caption         =   "FTP User:"
      Height          =   375
      Left            =   240
      TabIndex        =   13
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Label Label5 
      Caption         =   "FTP Password:"
      Height          =   375
      Left            =   240
      TabIndex        =   12
      Top             =   1680
      Width           =   1455
   End
   Begin VB.Label Label7 
      Caption         =   "File to FTP:"
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   2640
      Width           =   1455
   End
   Begin VB.Label Label9 
      Caption         =   "Result Window:"
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   3600
      Width           =   1455
   End
End
Attribute VB_Name = "frmSYFTPAPP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function OpenProcess Lib "kernel32" _
    (ByVal dwDesiredAccess As Long, _
    ByVal bInheritHandle As Long, _
    ByVal dwProcessId As Long) As Long
    
Private Declare Function WaitForSingleObject Lib "kernel32" _
    (ByVal hHandle As Long, _
    ByVal dwMilliseconds As Long) As Long
    
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Private Const SYNCHRONIZE = &H100000
Private Const INFINITE = &HFFFFFFFF

Private Const INSERT_INI As String = "insert into sys_ini " & vbCrLf & _
                                     "(ini_file_name, ini_user_id, ini_section, ini_field_name, ini_value) " & vbCrLf & _
                                     "values " & vbCrLf & _
                                     "( 'SYFTPAPP', " & vbCrLf & _
                                     "  '{USER}', " & vbCrLf & _
                                     "  '{SECTION}', " & vbCrLf & _
                                     "  '{FTP-ID}', " & vbCrLf & _
                                     "  '{VALUES}' )" & vbCrLf

Private Const CLEAR_INI As String = "delete from sys_ini " & vbCrLf & _
                                    " where ini_file_name = 'SYFTPAPP' " & vbCrLf & _
                                    "   and ini_field_name < '{FTP-ID}' " & vbCrLf

Private Const GET_RESULT As String = "select ini_value " & vbCrLf & _
                                    "   from sys_ini " & vbCrLf & _
                                    "  where ini_file_name = 'SYFTPAPP' " & vbCrLf & _
                                    "    and ini_user_id = '{USER}' " & vbCrLf & _
                                    "    and ini_section = 'FTP RESULT' " & vbCrLf & _
                                    "    and ini_field_name = '{FTP-ID}' " & vbCrLf

Private Const CLEANUP_ID As String = "delete from sys_ini " & vbCrLf & _
                                     " where ini_file_name = 'SYFTPAPP' " & vbCrLf & _
                                     "   and ini_user_id = '{USER}' " & vbCrLf & _
                                     "   and ini_field_name = '{FTP-ID}' " & vbCrLf
'

Private Sub cmdSend_Click()
    If ExecuteFTP() Then
        MsgBox "FTP Successful"
    Else
        MsgBox "FTP Failed"
    End If
End Sub

Private Sub txtWindow_Change()
    On Error Resume Next
    If Trim(Me.txtWindow.Text) <> "" Then
        If (Trim(Me.txtResult.Text) <> "") Then
            Me.txtResult.Text = Me.txtResult.Text & vbCrLf & txtWindow.Text
        Else
            Me.txtResult.Text = txtWindow.Text
        End If
    End If
    tfnLog Me.txtWindow.Text, io.StandardLogFile
End Sub

Public Property Get FTPHost() As String
    FTPHost = Me.txtHost
End Property

Public Property Let FTPHost(ByVal vNewValue As String)
    Me.txtHost = Trim(vNewValue)
End Property

Public Property Get FTPFolder() As String
    FTPFolder = Me.txtFolder
End Property

Public Property Let FTPFolder(ByVal vNewValue As String)
    Me.txtFolder = Trim(vNewValue)
End Property

Public Property Get ftpUser() As String
    ftpUser = Me.txtUser
End Property

Public Property Let ftpUser(ByVal vNewValue As String)
    Me.txtUser = Trim(vNewValue)
End Property

Public Property Get FTPPassword() As String
    FTPPassword = Me.txtPassword
End Property

Public Property Let FTPPassword(ByVal vNewValue As String)
    Me.txtPassword = Trim(vNewValue)
End Property

Public Property Get FTPType() As String
    FTPType = IIf(chkSFTP.value = vbChecked, "SFTP", "FTP")
End Property

Public Property Let FTPType(ByVal vNewValue As String)
    chkSFTP.value = IIf(UCase(Trim(vNewValue)) = "SFTP", vbChecked, vbUnchecked)
End Property

Public Property Get FTPFile() As String
    FTPFile = Me.txtFile
End Property

Public Property Let FTPFile(ByVal vNewValue As String)
    Me.txtFile = Trim(vNewValue)
End Property

Public Property Get FTPWindow() As Long
    Dim lng As Long
    If IsNumeric(Me.txtWindow.Text) Then
        lng = CLng(Me.txtWindow.Text)
    Else
        lng = Me.txtWindow.hWnd
    End If
    FTPWindow = lng
End Property

Public Property Let FTPWindow(ByVal vNewValue As Long)
    Me.txtWindow.Text = vNewValue
End Property

Public Property Get FTPResults() As String
    FTPResults = Replace(Me.txtResult.Text, "|", vbCrLf)
End Property

'********************************************************************
' FUNCTIONS
'********************************************************************

Public Function ExecuteFTP() As String
    On Error GoTo FINISHED
    Dim sId As String
    Dim sExe As String
    
    sId = Format(Now, "yyyyMMddHHmmss")
    
    subClear
    
    sExe = IIf(Right(App.Path, 1) = "\", App.Path + "SYFTPAPP.EXE", App.Path + "\SYFTPAPP.EXE")
    If Dir(sExe) = "" Then Err.Raise -1111, "", sExe + " was not found"
    
    If FTPHost = "" Then Err.Raise -2222, "ExecuteFTL", "FTP Host was not set!"
    'If m_folder = "" Then m_folder = ""
    If ftpUser = "" Then Err.Raise -2222, "ExecuteFTL", "FTP User was not set!"
    If FTPPassword = "" Then Err.Raise -2222, "ExecuteFTL", "FTP Password was not set!"
    'If FTPType = "" Then FTPType = "FTP"
    If FTPFile = "" Then Err.Raise -2222, "", "The file to FTP was not set!"
    If Dir(FTPFile) = "" Then Err.Raise -2222, "", "The file to FTP does not exist!"
    
    subPutRequest sId
    
    subLaunchExe sId
    
    subGetResult sId
    
    ExecuteFTP = ""
    Err.Clear
FINISHED:
    If Err.number <> 0 Then
        ExecuteFTP = "Error in ExecuteFTP: " & Replace(Err.Description, "|", vbCrLf) & vbCrLf
        Err.Clear
    End If
    
    On Error Resume Next
    subCleanup sId
    Err.Clear
End Function

Private Sub subPutRequest(sId As String)
    Dim sSql As String
    Dim sValues As String
    Dim cnt As Long
    
    sValues = "ID=" & sId & "|" & _
              "HOST = " & Replace(Me.FTPHost, " '", "''") & "|" & _
              "FOLDER=" & Replace(Me.FTPFolder, "'", "''") & "|" & _
              "USER=" & Replace(Me.ftpUser, "'", "''") & "|" & _
              "PWD=" & Replace(Me.FTPPassword, "'", "''") & "|" & _
              "TYPE=" & Replace(Me.FTPType, "'", "''") & "|" & _
              "FILE=" & Replace(Me.FTPFile, "'", "''")
    
    If Len(sValues) > 512 Then
        Err.Raise -3333, "subPutRequest", "The values string is too long for the database table! " & vbCrLf & "Try putting the file to FTP in a folder with a shorter path."
    End If
    
    sSql = Replace(INSERT_INI, "{USER}", tfnGetUserName())
    sSql = Replace(sSql, "{SECTION}", "FTP SETTINGS")
    sSql = Replace(sSql, "{FTP-ID}", sId)
    sSql = Replace(sSql, "{VALUES}", sValues)
    tfnLog fnHidePassword(sSql)
    cnt = fnExecuteSQL(sSql)
    
End Sub

Private Sub subClear()
    Dim sId As String
    Dim dt As Date
    Dim sSql As String
    Dim cnt As Long
    
    dt = DateAdd("d", -1, Now)
    sId = Format(dt, "yyyyMMdd000000")
   'sId = Format(dt, "yyyyMMddHHmmss")
    
    sSql = Replace(CLEAR_INI, "{FTP-ID}", sId)
    cnt = fnExecuteSQL(sSql)
    
End Sub

Private Sub subLaunchExe(sId As String)
    On Error GoTo FINISHED
    Dim sExe As String
    Dim sShell As String
    Dim ProcessID As Long
    Dim ProcessHandle As Long
    Dim wfsoReply As Long
    Dim nWaitStart As Long
    
    sExe = IIf(Right(App.Path, 1) = "\", App.Path + "SYFTPAPP.EXE", App.Path + "\SYFTPAPP.EXE")
    
    sShell = Chr(34) & sExe & Chr(34) & " " & _
             Chr(34) & "CONNECTION=" & t_dbMainDatabase.Connect & Chr(34) & " " & _
             Chr(34) & "ID=" & sId & Chr(34) & " " & _
             Chr(34) & "MODE=SILENT" & Chr(34) & " " & _
             Chr(34) & "HWND=" & Me.FTPWindow & Chr(34)
    
    tfnLog sShell
    
    nWaitStart = Timer
    ProcessID = Shell(sShell, VbAppWinStyle.vbHide)
    ProcessHandle = OpenProcess(SYNCHRONIZE, True, ProcessID)
    wfsoReply = WaitForSingleObject(ProcessHandle, 500)
    Do While wfsoReply = 258
        DoEvents
        If (Timer - nWaitStart) > 900 Then
            Err.Raise -8888, "", "Timed out waiting for SYFTPAPP.EXE to finish!"
        End If
        wfsoReply = WaitForSingleObject(ProcessHandle, 500)
    Loop
    
    Err.Clear
FINISHED:
    If Err.number <> 0 Then
        MsgBox "Error in subLaunchExe: " & Err.Description & vbCrLf & sShell
        Err.Raise Err.number, "subLaunchExe", Err.Description
    End If
End Sub

Private Sub subGetResult(sId As String)
    Dim sSql As String
    Dim rsTemp As Recordset
    Dim sResult As String
    
    sSql = Replace(GET_RESULT, "{USER}", tfnGetUserName())
    sSql = Replace(sSql, "{FTP-ID}", sId)
    If fnExecuteQuery(sSql, rsTemp) < 1 Then
        Err.Raise -7777, "subGetResult", "No result for the FTP process was found in the database. Please check the log file and or Event Viewer."
    End If
        
    sResult = Trim(rsTemp!ini_value & "")
    If Len(sResult) >= 7 Then
        If UCase(Left(sResult, 7)) = "SUCCESS" Then
            Exit Sub
        End If
    End If
    
    Err.Raise -7777, "subGetResult", "FTP Failed: " & sResult
    
End Sub

Private Sub subCleanup(sId As String)
    On Error GoTo FINISHED
    Dim sSql As String
    
    sSql = Replace(CLEANUP_ID, "{USER}", tfnGetUserName())
    sSql = Replace(sSql, "{FTP-ID}", sId)
    fnExecuteSQL sSql
    
    Err.Clear
FINISHED:
    If Err.number <> 0 Then
        MsgBox "Error Cleaning up the FTP database info: " & Err.Description & vbCrLf & sSql
        Err.Clear
    End If
End Sub

Private Function fnExecuteQuery(strSQL As String, rsTemp As Recordset) As Long
    On Error GoTo FINISHED
    
    Set rsTemp = t_dbMainDatabase.OpenRecordset(strSQL, dbOpenSnapshot, dbSQLPassThrough)
    If rsTemp.RecordCount > 0 Then
       rsTemp.MoveLast
       rsTemp.MoveFirst
    End If
    fnExecuteQuery = rsTemp.RecordCount
    
    Err.Clear
FINISHED:
    If Err.number <> 0 Then
        MsgBox "Error in fnExecuteQuery: " & Err.Description & vbCrLf & strSQL
        Err.Raise Err.number, "fnExecuteQuery", Err.Description
    End If
End Function

Private Function fnExecuteSQL(strSQL As String) As Long
    On Error GoTo FINISHED
    
    fnExecuteSQL = t_dbMainDatabase.ExecuteSQL(strSQL)
    
    Err.Clear
FINISHED:
    If Err.number <> 0 Then
        MsgBox "Error in fnExecuteSQL: " & Err.Description & vbCrLf & strSQL
        Err.Raise Err.number, "fnExecuteSQL", Err.Description
    End If
End Function

Private Function fnHidePassword(ByVal sString, Optional sDelim = ";") As String
    On Error GoTo FINISHED
    Dim aryParts() As String
    Dim sName As String
    Dim i As Long
    Dim eq As Long
    Dim bFoundString As Boolean
    Dim bFoundPart As Boolean
    Dim sNewString As String
    
    fnHidePassword = sString
    
    aryParts = Split(sString, sDelim)
    For i = 0 To UBound(aryParts)
        bFoundPart = False
        eq = InStr(1, aryParts(i), "=")
        If eq > 1 Then
            sName = Left(aryParts(i), eq - 1)
            If UCase(Trim(sName)) = "PWD" Or sName = "PASSWORD" Then
                bFoundString = True
                bFoundPart = True
            End If
            If Not bFoundPart Then
                If sNewString = "" Then
                    sNewString = aryParts(i)
                Else
                    sNewString = sNewString & sDelim & aryParts(i)
                End If
            Else
                If sNewString = "" Then
                    sNewString = sName & "=XXXX"
                Else
                    sNewString = sNewString & sDelim & sName & "=XXXX"
                End If
            End If
        Else
            If sNewString = "" Then
                sNewString = aryParts(i)
            Else
                sNewString = sNewString & sDelim & aryParts(i)
            End If
        End If
    Next
    
    If bFoundString Then
        fnHidePassword = sNewString
    End If
    
FINISHED:
    Err.Clear
End Function

