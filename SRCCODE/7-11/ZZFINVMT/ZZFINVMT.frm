VERSION 5.00
Object = "{C75015E0-2232-11D3-B440-0060971E99AF}#1.0#0"; "FACTFRM.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{478E45E0-5745-11CF-8918-00A02416C765}#1.0#0"; "SQAOTE32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmZZFINVMT 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Process Retail Sales Purchase Invoices"
   ClientHeight    =   6060
   ClientLeft      =   1056
   ClientTop       =   1956
   ClientWidth     =   8880
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.6
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   1
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6060
   ScaleWidth      =   8880
   Begin MSComctlLib.ProgressBar PbProgressBar 
      Height          =   264
      Left            =   6924
      TabIndex        =   8
      Top             =   5748
      Width           =   1896
      _ExtentX        =   3344
      _ExtentY        =   466
      _Version        =   393216
      Appearance      =   1
   End
   Begin FACTFRMLib.FactorFrame ffraStatusbar 
      Height          =   360
      Left            =   0
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   5700
      Width           =   8880
      _Version        =   65536
      _ExtentX        =   15663
      _ExtentY        =   635
      _StockProps     =   77
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Align           =   2
      CaptionPos      =   1
      Style           =   5
      BeginProperty PanelFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin FACTFRMLib.FactorFrame efraBackground 
      Height          =   5184
      Left            =   0
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   492
      Width           =   8892
      _Version        =   65536
      _ExtentX        =   15684
      _ExtentY        =   9144
      _StockProps     =   77
      BackColor       =   8388608
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelOuter      =   0
      PicturePos      =   0
      BeginProperty PanelFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin FACTFRMLib.FactorFrame efraProcess 
         Height          =   4512
         Left            =   84
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   84
         Width           =   8712
         _Version        =   65536
         _ExtentX        =   15367
         _ExtentY        =   7959
         _StockProps     =   77
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   5
         Caption         =   " "
         BeginProperty PanelFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin MSComDlg.CommonDialog dlgFileNames 
            Left            =   2820
            Top             =   -12
            _ExtentX        =   677
            _ExtentY        =   677
            _Version        =   393216
         End
         Begin VB.ListBox lstStatus 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   4128
            HelpContextID   =   802
            ItemData        =   "ZZFINVMT.frx":0000
            Left            =   48
            List            =   "ZZFINVMT.frx":0002
            TabIndex        =   5
            TabStop         =   0   'False
            Top             =   336
            Width           =   8604
         End
         Begin VB.Label lblProcess 
            Caption         =   "Invoice Processing Status:"
            Height          =   264
            Left            =   84
            TabIndex        =   1
            Top             =   60
            Width           =   2412
         End
      End
      Begin FACTFRMLib.FactorFrame cmdProcess 
         Height          =   396
         HelpContextID   =   800
         Left            =   4548
         TabIndex        =   9
         Top             =   4692
         Width           =   1308
         _Version        =   65536
         _ExtentX        =   2307
         _ExtentY        =   698
         _StockProps     =   77
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "&Process"
         CaptionPos      =   4
         PicturePos      =   3
         ShowFocusRect   =   -1  'True
         Style           =   3
         BorderWidth     =   4
         BeginProperty PanelFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin FACTFRMLib.FactorFrame cmdExitCancelBtn 
         Height          =   396
         HelpContextID   =   15
         Left            =   7500
         TabIndex        =   11
         Top             =   4692
         Width           =   1308
         _Version        =   65536
         _ExtentX        =   2307
         _ExtentY        =   698
         _StockProps     =   77
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "E&xit"
         CaptionPos      =   4
         PicturePos      =   3
         ShowFocusRect   =   -1  'True
         Style           =   3
         BorderWidth     =   4
         BeginProperty PanelFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin FACTFRMLib.FactorFrame cmdPrintReport 
         Height          =   396
         HelpContextID   =   801
         Left            =   6036
         TabIndex        =   10
         Top             =   4692
         Width           =   1308
         _Version        =   65536
         _ExtentX        =   2307
         _ExtentY        =   698
         _StockProps     =   77
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Print &Report"
         CaptionPos      =   4
         ShowFocusRect   =   -1  'True
         Style           =   3
         BorderWidth     =   4
         BeginProperty PanelFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin FACTFRMLib.FactorFrame efraToolBar 
      Height          =   468
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   8880
      _Version        =   65536
      _ExtentX        =   15663
      _ExtentY        =   825
      _StockProps     =   77
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Align           =   1
      FMName          =   "ZZFINVMT"
      CaptionPos      =   4
      Style           =   6
      BeginProperty PanelFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.Timer tmrKeyBoard 
         Enabled         =   0   'False
         Interval        =   250
         Left            =   6624
         Top             =   96
      End
      Begin MSComctlLib.Toolbar tbToolbar 
         Height          =   372
         Left            =   60
         TabIndex        =   2
         Top             =   84
         Width           =   6876
         _ExtentX        =   12129
         _ExtentY        =   656
         ButtonWidth     =   614
         ButtonHeight    =   572
         _Version        =   393216
      End
   End
   Begin VB.Data datComboDropDown 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   384
      Left            =   2940
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4752
      Visible         =   0   'False
      Width           =   1920
   End
   Begin SQAOTestObjectsCtl.SQAOTest SQAOTest1 
      Height          =   456
      Left            =   10896
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   96
      Width           =   456
      _ExtentX        =   804
      _ExtentY        =   804
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
         HelpContextID   =   1
      End
      Begin VB.Menu mnuPrint 
         Caption         =   "&Print"
         Enabled         =   0   'False
         HelpContextID   =   7
      End
   End
   Begin VB.Menu mnuMainEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuCancel 
         Caption         =   "Ca&ncel"
         Enabled         =   0   'False
         HelpContextID   =   1
      End
      Begin VB.Menu mnuEditSep10 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCut 
         Caption         =   "Cu&t"
         Enabled         =   0   'False
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuCopy 
         Caption         =   "&Copy"
         Enabled         =   0   'False
         HelpContextID   =   8
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuPaste 
         Caption         =   "&Paste"
         Enabled         =   0   'False
         HelpContextID   =   9
         Shortcut        =   ^V
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Options"
      Begin VB.Menu mnuCopyFrom 
         Caption         =   "&Copy From"
         Enabled         =   0   'False
         HelpContextID   =   10
      End
      Begin VB.Menu mnuProcess 
         Caption         =   "&Process"
      End
      Begin VB.Menu mnuEmpty 
         Caption         =   "-"
      End
      Begin VB.Menu mnuModules 
         Caption         =   "mnuModules"
         Index           =   0
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuContents 
         Caption         =   "&Contents"
      End
      Begin VB.Menu mnuHelpSep10 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
         HelpContextID   =   11
      End
   End
End
Attribute VB_Name = "frmZZFINVMT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private t_bStartupFlag As Boolean 'optional startup flag
Private t_bDataChanged As Boolean 'data changed flag
Private t_bUpdateTable As Boolean 'update data flag

Private t_nFormMode As Integer         'global used to track the current form operating mode
Private Const IDLE_MODE As Integer = 0 'idle mode activates the NoDrop Cursor

'========================
'Standard Button Captions
'========================
Private Const t_szCAPTION_CANCEL As String = "&Cancel"
Private Const t_szCAPTION_EXIT As String = "E&xit"

'==========================
'Status Bar Default Strings
'==========================
Private Const t_szEXIT As String = "Exit"
Private Const t_szCANCEL As String = "Cancel"

Private Const t_szPRINT As String = "Print"
Private Const t_szHELP As String = "Help"
'

Const INI_BUFFRER_SIZE As Integer = 1024
Const INISECTION As String = "Incoming Invoice File Location"
Const INISUBSECTION As String = "filePath"
Private g_bProcessOk As Boolean

Private Sub cmdPrintReport_Click()
    
    If g_bProcessOk Then
        frmReports.chkPrintErr.Value = vbUnchecked
        frmReports.chkPrintErr.Enabled = False
        frmReports.chkPrintProcLog.Value = vbChecked
    Else
        frmReports.chkPrintErr.Value = vbChecked
        frmReports.chkPrintErr.Enabled = True
        frmReports.chkPrintProcLog.Value = vbUnchecked
    End If
    
    frmReports.Show vbModal
    subSetFocus cmdExitCancelBtn
End Sub

Private Sub cmdPrintReport_GotFocus()
    tfnSetStatusBarMessage "Print Report"
End Sub

Private Sub cmdProcess_Click()
    Dim sFileName As String
    
    #If PROTOTYPE Then
        Exit Sub
    #End If
    
    subSetExitCancelBtn "Cancel"
    subEnableCancelBtn False
    subEnablePrintBtn False
    subEnableProcessBtn False
    t_bDataChanged = True
    lstStatus.Clear
    g_bProcessOk = False
    
    If fnFileName(sFileName) Then
        subShowHourGlass
        tfnSetStatusBarMessage "Processing, Please Wait..."
        
        If fnIsFile(sFileName) Then
            subDisplayMsg "START PROCESSING FILE " & sFileName & " AT " & Time & " " & Date
            
            If fnProcessRSInvFile(sFileName) Then
                subDisplayMsg "Processing successfully."
                g_bProcessOk = True
            Else
                subDisplayMsg "Processing Failed."
            End If
            
            subSetProgress 0
            subDisplayMsg "END PROCESSING FILE " & sFileName & " AT " & Time & " " & Date
        Else
            subDisplayMsg "Input file not found"
        End If
        
    Else
        subDisplayMsg "Action cancelled"
    End If
    
    subHideHourGlass
    subEnableCancelBtn True
    subEnablePrintBtn True
    DoEvents
    subSetFocus cmdPrintReport
    
End Sub

Private Sub cmdProcess_GotFocus()
    tfnSetStatusBarMessage "Process"
End Sub

'===========
'Form Events
'===========
Private Sub Form_Initialize() 'called before Form_Load
    t_bStartupFlag = True
    t_bDataChanged = False
    t_bUpdateTable = False
    
    t_nFormMode = IDLE_MODE
    
    CRLF = Chr(10) + Chr(13)

    ' ** change the help file for the application
    App.HelpFile = szHelp7_11
End Sub

Private Sub Form_Unload(CANCEL As Integer)

    On Error Resume Next
    
    Set objErrHandler = Nothing
    
    Set objCurrTabControl = Nothing
    
    Unload frmContext
    Unload frmAbout
    
    'project local database object variables
    If Not dbLocal Is Nothing Then
        dbLocal.Close
    End If
    
    Set dbLocal = Nothing
    
    If Not t_dbMainDatabase Is Nothing Then
        t_dbMainDatabase.Close
    End If
    
    Set t_dbMainDatabase = Nothing
    Set t_wsWorkSpace = Nothing
    Set t_engFactor = Nothing
    Set t_oleObject = Nothing
    End
End Sub

Private Sub Form_Load()

#If Not PROTOTYPE Then
        If tfnAuthorizeExecute(Command) = False Then 'Check for handshake if not in the development mode
            Unload Me
            Exit Sub
        End If
        
        If tfnOpenDatabase = False Then 'open the database, ODBC Dialog Box during developemnt, oleObject Connection String when not
            Unload Me
            Exit Sub
        End If
        
        'connect to local database
        Set dbLocal = tfnOpenLocalDatabase()
        
        If dbLocal Is Nothing Then
            MsgBox "System Error: Unable to open Local Database, Program will be terminated", vbExclamation
            Unload Me
            Exit Sub
        End If
        
        subInitErrorHandler   ' Setup Error Control
        tfnUpdateVersion
#End If
        
    tfnDisableFormSystemClose Me
    subSetupToolBar
    tmrKeyBoard.Enabled = False
    tfnCenterForm Me
    t_bStartupFlag = False
    Me.Show
    subShowHourGlass
    Me.Enabled = False
    DoEvents
    Me.Enabled = True
    tfnResetScreen
    tmrKeyBoard.Enabled = True
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_GotFocus()
    tmrKeyBoard.Enabled = True
End Sub

Private Sub Form_LostFocus()
    tmrKeyBoard.Enabled = False
End Sub

Private Sub cmdExitCancelBtn_GotFocus()
    
    If cmdExitCancelBtn.Caption = t_szCAPTION_EXIT Then
        tfnSetStatusBarMessage t_szEXIT
    Else
        tfnSetStatusBarMessage t_szCANCEL
    End If

End Sub

Private Sub cmdExitCancelBtn_Click()
    
    If cmdExitCancelBtn.Caption = t_szCAPTION_CANCEL Then
        subCancel
    Else
        subExit
    End If

End Sub


'=====================
'Toolbar Button Events
'=====================
Private Sub subCancel()
    
    If t_bDataChanged Then
        
        If Not tfnCancelExit(t_szCANCEL_MESSAGE) Then
            Exit Sub
        End If
    
    End If
    
    subEnableProcessBtn True
    DoEvents
    subSetFocus cmdProcess
    t_nFormMode = IDLE_MODE
    t_bDataChanged = False
    lstStatus.Clear
    tfnResetScreen 'reset all the buttons
End Sub

Private Sub subExit()
    
    If t_bDataChanged Then
        
        If Not tfnCancelExit(t_szEXIT_MESSAGE) Then
            Exit Sub
        End If
    
    End If
    
    Unload Me
End Sub

Private Sub lstStatus_GotFocus()
    tfnSetStatusBarMessage "Process Status"
End Sub

'============
'Menu  Events
'============
Private Sub mnuExit_Click()
    subExit
End Sub

Private Sub mnuContents_Click()
    frmContext.RunItem HELP_UP
End Sub

Private Sub mnuAbout_Click()
    tfnCenterForm frmAbout, Me
    frmAbout.Show vbModal
End Sub

'========================================
'Main Edit Menu Events Cancel, Copy/Paste
'========================================
Private Sub mnuMainEdit_Click()
  
    If t_nFormMode <> IDLE_MODE And TypeOf Screen.ActiveControl Is TextBox Then
        mnuCopy.Enabled = (Screen.ActiveControl.SelLength > 0)
        mnuCut.Enabled = (Screen.ActiveControl.SelLength > 0)
        mnuPaste.Enabled = (Clipboard.GetFormat(vbCFText))
    Else
        mnuCopy.Enabled = False
        mnuCut.Enabled = False
        mnuPaste.Enabled = False
    End If
    
End Sub

Private Sub mnuCancel_Click()
    subCancel
End Sub

Private Sub mnuCopy_Click()
    Clipboard.SetText Screen.ActiveControl.Text, vbCFText
End Sub

Private Sub mnuCut_Click()
    Clipboard.SetText Screen.ActiveControl.Text, vbCFText
    Screen.ActiveControl.Text = ""
End Sub

Private Sub mnuPaste_Click()
  
  If TypeOf Screen.ActiveControl Is TextBox Then
    Screen.ActiveControl.Text = Clipboard.GetText(vbCFText)
  End If
  
End Sub

'======================
'form support functions
'======================

'
'Function        : tfnShowStatusBarMessage - show status bar messages
'Passed Variables: message string
'Returns         : none
'
Public Sub tfnSetStatusBarMessage(szMessage As String)
    
    If t_bStartupFlag Then
        Exit Sub
    End If
    
    ffraStatusbar.ForeColor = STANDARD_TEXT_COLOR
    ffraStatusbar.Font.Bold = False
    ffraStatusbar.Caption = szMessage
    ffraStatusbar.Refresh

End Sub

Private Sub tfnSetInitializingMessage()

    ffraStatusbar.ForeColor = STANDARD_TEXT_COLOR
    ffraStatusbar.Font.Bold = False
    ffraStatusbar.Caption = "Initializing program.  Please wait..."
    ffraStatusbar.Refresh

End Sub

'
'Function        : tfnSetStatusBarError - show status bar error message in red
'Passed Variables: error message string
'Returns         : none
'
Private Sub tfnSetStatusBarError(szErrorMessage As String, Optional vNoBeep As Variant)
    
    ffraStatusbar.ForeColor = ERROR_TEXT_COLOR
    ffraStatusbar.Font.Bold = True
    ffraStatusbar.Caption = szErrorMessage
    
    If IsMissing(vNoBeep) Then
        Beep
    End If
    
    ffraStatusbar.Refresh

End Sub
'
'Function        : tfnSetStatusBarCorrect - entry ok status bar message
'Passed Variables: entry message string
'Returns         : none
'
Private Sub tfnSetStatusBarCorrect(szCorrectMessage As String)
    ffraStatusbar.ForeColor = CORRECT_TEXT_COLOR
    ffraStatusbar.Font.Bold = True
    ffraStatusbar.Caption = szCorrectMessage
    ffraStatusbar.Refresh
End Sub
'
'Function        : tfnResetScreenButtons - sets all the form buttons to the startup condition
'Passed Variables: none
'Returns         : none
'
Private Sub tfnResetScreen()
    
    On Error Resume Next
    
    subSetExitCancelBtn "EXIT"
    frmContext.ButtonEnabled(COPY_UP) = False
    mnuExit.Enabled = True
    mnuCopy.Enabled = False
    mnuPaste.Enabled = False
#If PROTOTYPE Then
    cmdProcess.Enabled = False
    mnuProcess.Enabled = False
    cmdPrintReport.Enabled = False
    mnuPrint.Enabled = False
    frmContext.ButtonEnabled(PRINT_UP) = False
#Else
    cmdProcess.Enabled = True
    mnuProcess.Enabled = True
#End If
    PbProgressBar.Visible = False
End Sub

Private Sub mnuPrint_Click()
    cmdPrintReport_Click
End Sub

Private Sub mnuProcess_Click()
    cmdProcess_Click
End Sub

Private Sub tmrKeyboard_Timer() 'status bar timer - 250ms
    tfnUpdateStatusBar Me 'process the status bar
End Sub

Private Sub subInitErrorHandler()
    
    If objErrHandler Is Nothing Then
        Set objErrHandler = New clsErrorHandler
        
        With objErrHandler
            Set .FormParent = Me
            Set .DatabaseEngine = t_engFactor
            Set .LocalDatabase = dbLocal
        End With
    
    End If

End Sub

Private Sub subSetExitCancelBtn(sExitCancel As String)
    
    If sExitCancel = "EXIT" Then
        cmdExitCancelBtn.Caption = t_szCAPTION_EXIT
        frmContext.ButtonEnabled(CANCEL_UP) = False
        mnuCancel.Enabled = False
    Else
        cmdExitCancelBtn.Caption = t_szCAPTION_CANCEL
        frmContext.ButtonEnabled(CANCEL_UP) = True
        mnuCancel.Enabled = True
    End If
    
End Sub

Private Sub Form_Resize()
    frmContext.FormResize
End Sub

Private Sub subSetupToolBar()
    
    With frmContext
        .BeginSetupToolbar Me
        .EndSetupToolbar
        .HelpFile = szHelpFileName
    End With
    
End Sub

Public Sub TBButtonCallBack(ByVal nID As Integer)
    
    Select Case nID
        Case CANCEL_UP
            subCancel
        Case EXIT_UP
            subExit
        Case PRINT_UP
            cmdPrintReport_Click
    End Select
    
End Sub

Private Sub mnuModules_Click(Index As Integer)
    frmContext.MenuClick Index
End Sub

Private Sub tbToolbar_ButtonClick(ByVal Button As Button)
    frmContext.ButtonClick Button
End Sub

Private Sub tbToolbar_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    frmContext.TBMouseMove
End Sub

Public Sub subShowHourGlass()
    tmrKeyBoard.Enabled = False
    Screen.MousePointer = vbHourglass
End Sub

Public Sub subHideHourGlass()
    Screen.MousePointer = vbDefault
    tmrKeyBoard.Enabled = True
End Sub

Private Sub subSetFocus(v As Control)
    On Error Resume Next
    v.SetFocus
End Sub

Private Sub subEnableCancelBtn(bOnOff As Boolean)
    cmdExitCancelBtn.Enabled = bOnOff
    frmContext.ButtonEnabled(CANCEL_UP) = bOnOff
    mnuCancel.Enabled = bOnOff
End Sub

Private Sub subEnablePrintBtn(bOnOff As Boolean)
    cmdPrintReport.Enabled = bOnOff
    frmContext.ButtonEnabled(PRINT_UP) = bOnOff
    mnuPrint.Enabled = bOnOff
End Sub

Private Sub subEnableProcessBtn(bOnOff As Boolean)
    cmdProcess.Enabled = bOnOff
    mnuProcess.Enabled = bOnOff
End Sub

Private Function fnFileName(sFileName As String) As Boolean
    Dim sInitDirPath As String
    Dim sDirPath As String
    Dim nDirPos As Integer
    
    fnFileName = False
    sInitDirPath = fnGetInitDirPath()
    
    With dlgFileNames
        .InitDir = sInitDirPath
        .DialogTitle = "Invoice File Location"
        .FileName = "RSP*.*"
        .flags = cdlOFNHideReadOnly Or cdlOFNFileMustExist
        .Filter = "Invoice File|RSP*.*|All Files(*.*)|*.*"
        .CancelError = True
        On Error Resume Next
        .ShowOpen
        
        'Click cancel
        If Err.Number = 32755 Then
            sFileName = ""
            sDirPath = sInitDirPath
        Else
            sFileName = .FileName
            nDirPos = InStrRev(sFileName, "\")
            
            If nDirPos > 0 Then
                sDirPath = Left$(sFileName, nDirPos - 1)
            Else
                sDirPath = .InitDir
            End If
            
            fnFileName = True
        End If

    End With
    
    If sDirPath <> "" Then
        subWriteInitDirPath sDirPath
    End If
    
End Function

Public Function fnIsFile(ByVal szFilename As String) As Boolean
    
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

Private Function fnGetInitDirPath() As String
    Dim szINI As String
    Dim sDir As String
    Dim sINIFile As String
    Dim lTotal As Long
    
    szINI = Space(INI_BUFFRER_SIZE)
    sINIFile = App.Path & "\zzfinvmt.ini"
    
    lTotal = GetPrivateProfileString(INISECTION, INISUBSECTION, szEMPTY, szINI, INI_BUFFRER_SIZE, sINIFile)
    
    If lTotal <> 0 Then
        szINI = Left$(szINI, lTotal)
    Else
        szINI = App.Path
    End If
    
    fnGetInitDirPath = szINI
    
End Function

Private Sub subWriteInitDirPath(sDirPath As String)
    Dim sINIFile As String
    
    sINIFile = App.Path & "\zzfinvmt.ini"
    WritePrivateProfileString INISECTION, INISUBSECTION, sDirPath, sINIFile
End Sub

