VERSION 5.00
Begin VB.Form AuthorityCode 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Authorization Code Input"
   ClientHeight    =   2040
   ClientLeft      =   735
   ClientTop       =   3555
   ClientWidth     =   5880
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2040
   ScaleWidth      =   5880
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picButtons 
      BorderStyle     =   0  'None
      Height          =   444
      Left            =   690
      ScaleHeight     =   450
      ScaleWidth      =   4410
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   1416
      Width           =   4410
      Begin VB.CommandButton cmdView 
         Caption         =   "&View Only"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   396
         Left            =   3060
         TabIndex        =   8
         Top             =   30
         Visible         =   0   'False
         Width           =   1308
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "O&K"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   396
         Left            =   0
         TabIndex        =   6
         Top             =   24
         Width           =   1308
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   396
         Left            =   1530
         TabIndex        =   7
         Top             =   24
         Width           =   1308
      End
   End
   Begin VB.PictureBox picAuthCode 
      BorderStyle     =   0  'None
      Height          =   816
      Left            =   756
      ScaleHeight     =   810
      ScaleWidth      =   4470
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   504
      Width           =   4476
      Begin VB.TextBox txtAuthCode 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   720
         PasswordChar    =   "*"
         TabIndex        =   5
         Top             =   348
         Width           =   2808
      End
      Begin VB.Label lblCaption 
         Caption         =   "Need Manager Authorization Code to Override"
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
         Left            =   96
         TabIndex        =   9
         Top             =   12
         Width           =   4392
      End
   End
   Begin VB.PictureBox picCustomerInfo 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   348
      Left            =   780
      ScaleHeight     =   345
      ScaleWidth      =   4935
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   120
      Width           =   4932
      Begin VB.Label lblCustInfo 
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
         Left            =   36
         TabIndex        =   10
         Top             =   12
         Width           =   4740
      End
   End
   Begin VB.PictureBox picMessage 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   960
      Left            =   72
      ScaleHeight     =   960
      ScaleWidth      =   5580
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   264
      Width           =   5580
      Begin VB.PictureBox picIcon 
         BorderStyle     =   0  'None
         Height          =   480
         Left            =   72
         Picture         =   "AuthCode.frx":0000
         ScaleHeight     =   480
         ScaleWidth      =   495
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   288
         Width           =   492
      End
      Begin VB.Label lblMsg 
         Caption         =   "Do you want to continue?"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   264
         Left            =   744
         TabIndex        =   1
         Top             =   552
         Width           =   4764
      End
   End
End
Attribute VB_Name = "AuthorityCode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Usage:
'   1. Initialize:
'       After open database, call
'          AuthorityCode.SetAuthCode
'       in form_load
'   2. To popup, do the following
'       AuthorityCode.CustomerNumber = 12345678
'       AuthorityCode.Show vbModal
'   3. To check whether can continue processing
'    If AuthorityCode.Authorized Then
'        It is authorized.you can continue
'    Else
'        It is not authorized. Stop here
'    End If
'   4. To check whether an authorization code was set. Use
'        AuthorityCode.AuthCodeSet
'       This property returns true if it is set
'       otherwise it returns false
Option Explicit
    
Private Const SYSPARM4007 = 4007
Private sAuthCode As String
Private bAuthCodeSet As Boolean
Private nAction As Integer
'david 07/31/2011  #3013-682745
'tracks date, customer #, user ID, program overriden, delivery ticket #/work order # if available
Private m_lCustomer As Long
Private m_sUserID As String
Private m_sProgramID As String
Private m_lTicketWoNbr As Long
Private m_sAddEditMode As String 'ADD or EDIT
Private m_bViewOnly As Boolean
Private m_sSysParm4086 As String
Private m_sSysParm4087 As String

Private m_sAuthType As String  'sta_auth_type - (M)anager Code, (P)osting Code
'''''''''''''''''''''''''''''''
'

Property Get AuthCodeSet() As Boolean
    AuthCodeSet = bAuthCodeSet
End Property

Property Get Authorized() As Boolean
    If nAction = vbOK Then
        Authorized = True
    Else
        Authorized = False
    End If
End Property

Property Get ViewOnly() As Boolean
    ViewOnly = m_bViewOnly
End Property

'david 07/31/2011  #3013-682745
'tracks date, customer #, user ID, program overriden, delivery ticket #/work order # if available
Property Let CustomerNumber(ByVal lCust As Long)
    lblCustInfo.Caption = "Customer (" & CStr(lCust) & ") is on credit hold."
    m_lCustomer = lCust
End Property

Private Function fnControlRight(ctrlTemp As Control) As Integer
    fnControlRight = ctrlTemp.Left + ctrlTemp.Width
End Function

'Editing mode = (A)dd or (E)dit
Public Function SetAuthCode(Optional sUserID As String = "", _
                            Optional sProgramID As String = "", _
                            Optional lTicketWoNbr As Long = 0, _
                            Optional sAddEditMode As String = "") As Boolean
    Const SUB_NAME = "SetAuthCode"
    
    Dim strSQL As String
    Dim rsTemp As Recordset
    
    'david 07/31/2011  #3013-682745
    'tracks date, customer #, user ID, program overriden, delivery ticket #/work order # if available
    m_sUserID = sUserID
    m_sProgramID = sProgramID
    m_lTicketWoNbr = lTicketWoNbr
    m_sAddEditMode = UCase(Left(sAddEditMode, 1))
    '''''''''''''''''''''''''''''''
    
    strSQL = "SELECT parm_field FROM sys_parm" _
         & " WHERE parm_nbr = " & SYSPARM4007
    
    bAuthCodeSet = False
    sAuthCode = ""
    
    On Error GoTo errQuery
    Set rsTemp = t_dbMainDatabase.OpenRecordset(strSQL, dbOpenSnapshot, dbSQLPassThrough)
    On Error GoTo 0
    If rsTemp.RecordCount > 0 Then
        If Not IsNull(rsTemp!parm_field) Then
            sAuthCode = Trim(rsTemp!parm_field)
            If sAuthCode <> "" Then
                bAuthCodeSet = True
            End If
        End If
    End If
    
    'david 07/31/2011  #3013-682745
    m_sSysParm4086 = ""
    m_sSysParm4087 = ""
    
    strSQL = "SELECT parm_field FROM sys_parm" _
         & " WHERE parm_nbr = " & 4086
    
    On Error GoTo errQuery
    Set rsTemp = t_dbMainDatabase.OpenRecordset(strSQL, dbOpenSnapshot, dbSQLPassThrough)
    On Error GoTo 0
    If rsTemp.RecordCount > 0 Then
        m_sSysParm4086 = UCase(Trim(rsTemp!parm_field & ""))
    End If
    
    If m_sSysParm4086 = "Y" Then
        strSQL = "SELECT parm_field FROM sys_parm" _
             & " WHERE parm_nbr = " & 4087
        
        On Error GoTo errQuery
        Set rsTemp = t_dbMainDatabase.OpenRecordset(strSQL, dbOpenSnapshot, dbSQLPassThrough)
        On Error GoTo 0
        If rsTemp.RecordCount > 0 Then
            m_sSysParm4087 = Trim(rsTemp!parm_field & "")
        End If
    End If
    '''''''''''''''''''''''''''''''
    
    subSetMyState
    SetAuthCode = True
    Exit Function

errQuery:
    tfnErrHandler SUB_NAME, strSQL
    SetAuthCode = False
End Function

Private Function fnValidCode() As Boolean
    'david 07/31/2011  #3013-682745
'    If Trim(txtAuthCode.Text) = sAuthCode Or (m_sSysParm4086 = "Y" And Trim(txtAuthCode.Text) = m_sSysParm4087) Then
'        fnValidCode = True
'    End If
    If m_sAddEditMode = "E" Then
        If m_sSysParm4086 = "Y" And m_sSysParm4087 <> "" Then
            If Trim(txtAuthCode.Text) = m_sSysParm4087 Then
                m_sAuthType = "P"
                fnValidCode = True
            End If
        Else
            If Trim(txtAuthCode.Text) = sAuthCode Then
                m_sAuthType = "M"
                fnValidCode = True
            End If
        End If
    Else
        If Trim(txtAuthCode.Text) = sAuthCode Then
            m_sAuthType = "M"
            fnValidCode = True
        End If
    End If
    '''''''''''''''''''''''''''''''
End Function

Private Sub subEnableOK(bFlag As Boolean)
'    cmdOK.Enabled = bFlag
    cmdOk.Enabled = True  'Always enabled
    
End Sub

Private Sub subSetMyState()
    Me.Caption = App.Title
    If bAuthCodeSet Then
        'david 07/31/2011  #3013-682745
        If m_sAddEditMode = "E" Then
            'three buttons
            cmdView.Visible = True
            picButtons.Left = 690
            cmdCancel.Left = 1530
            
            If m_sSysParm4086 = "Y" And m_sSysParm4087 <> "" Then
                lblCaption.Caption = "Need Posting Authorization Code to Override"
            Else
                lblCaption.Caption = "Need Manager Authorization Code to Override"
            End If
        Else
            'two buttons
            cmdView.Visible = False
            picButtons.Left = 1110
            cmdCancel.Left = 2076
            lblCaption.Caption = "Need Manager Authorization Code to Override"
        End If
        '''''''''''''''''''''''''''''''
        
        cmdOk.Caption = "O&K"
        cmdCancel.Caption = "&Cancel"
        picCustomerInfo.Top = 156
        picMessage.Visible = False
        picAuthCode.Visible = True
        subEnableOK False
        txtAuthCode.Text = ""
    Else
        'david 07/31/2011  #3013-682745
        'two buttons
        cmdView.Visible = False
        picButtons.Left = 1110
        cmdCancel.Left = 2076
        '''''''''''''''''''''''''''''''
        
        cmdOk.Caption = "&Yes"
        cmdCancel.Caption = "&No"
        picCustomerInfo.Top = 384
        picMessage.Visible = True
        picAuthCode.Visible = False
    End If
End Sub

Private Sub cmdView_Click()
    Me.Reset
    Me.Hide
    m_bViewOnly = True
End Sub

Private Sub cmdCancel_Click()
    Me.Reset
    Me.Hide
    nAction = vbCancel
End Sub

Private Sub cmdOK_Click()
    Static bDonotShowError As Boolean
    
    If bAuthCodeSet Then
        If fnValidCode() Then
            InsertAuthCodeTracking m_lCustomer, m_sUserID, m_sProgramID, m_lTicketWoNbr, (Not bDonotShowError)
            bDonotShowError = True
            
            Me.Reset
            Me.Hide
            nAction = vbOK
        Else
            MsgBox "Invalid authorization code.", vbOKOnly + vbExclamation
            subSelectText txtAuthCode
            txtAuthCode.SetFocus
        End If
    Else
        Me.Reset
        Me.Hide
        nAction = vbOK
    End If
End Sub

Private Sub Form_Activate()
    DoEvents
    Screen.MousePointer = vbDefault
    If bAuthCodeSet Then
        subEnableOK False
        txtAuthCode.Text = ""
        On Error Resume Next
        txtAuthCode.SetFocus
    Else
        On Error Resume Next
        cmdCancel.SetFocus
    End If
End Sub

Private Sub Form_Initialize()
    Me.Reset
End Sub

Private Sub Form_Load()
    tfnCenterForm Me
End Sub

Private Sub txtAuthCode_Change()
    If fnValidCode Then
        subEnableOK True
    Else
        subEnableOK False
    End If
End Sub

Private Sub subSelectText(txtBox As Textbox)
    txtBox.SelStart = 0
    txtBox.SelLength = Len(txtBox.Text)
End Sub

Private Sub txtAuthCode_GotFocus()
    subSelectText txtAuthCode
End Sub

Private Sub txtAuthCode_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        cmdOK_Click
    End If
End Sub

'david 07/31/2011  #3013-682745
'tracks date, customer #, user ID, program overriden, delivery ticket #/work order # if available
Public Sub Reset()
    m_lCustomer = 0
    m_sUserID = ""
    m_sProgramID = ""
    m_lTicketWoNbr = 0
    m_sAddEditMode = ""
    m_bViewOnly = False
    'Important!!!
    'DO NOT clear m_sAuthType
End Sub

Public Sub InsertAuthCodeTracking(lCustomer As Long, sUserID As String, sProgramID As String, lTicketWoNbr As Long, bShowError As Boolean)
    Const SUB_NAME As String = "InsertAuthCodeTracking"
    
    Dim strSQL As String
    
        
    If lCustomer <= 0 Or sUserID = "" Or sProgramID = "" Then
        Exit Sub
    End If
    
    If Not fnColumnExists("sys_track_auth", "sta_date", bShowError) Then
        Exit Sub
    End If
    
    strSQL = "insert into sys_track_auth" _
        & " (sta_date, sta_customer, sta_user," _
        & " sta_program, sta_ticket_wo, sta_auth_type) values ("
    strSQL = strSQL + tfnDateString(Date, True) + ", "
    strSQL = strSQL & lCustomer & ", "
    strSQL = strSQL + tfnSQLString(sUserID) + ", "
    strSQL = strSQL + tfnSQLString(sProgramID) + ", "
    strSQL = strSQL & IIf(lTicketWoNbr > 0, lTicketWoNbr, "null") & ","
    strSQL = strSQL & tfnSQLString(m_sAuthType) & ");"
    
    fnExecuteSQL strSQL, SUB_NAME, bShowError
End Sub

Private Function fnColumnExists(TableName As String, ColumnName As String, Optional bShowError As Boolean = True) As Boolean
    Const SUB_NAME As String = "fnColumnExists"
    
    Dim strSQL As String
    Dim rsTemp As Recordset
    
    strSQL = "select tabName, colName" & _
        " from systables, syscolumns" & _
        " where systables.tabid = syscolumns.tabid " & _
        " and tabName = " & tfnSQLString(TableName) & _
        " and colName = " & tfnSQLString(ColumnName)

    fnColumnExists = fnRecordset(rsTemp, strSQL, SUB_NAME, bShowError) > 0
End Function

Private Function fnRecordset(rsTemp As Recordset, strSQL As String, _
                 Optional sCalledFrom As String = "", _
                 Optional bShowErrow As Boolean = True) As Long
    On Error GoTo SQLError
        
    Set rsTemp = t_dbMainDatabase.OpenRecordset(strSQL, dbOpenSnapshot, dbSQLPassThrough)
    
    If rsTemp.RecordCount > 0 Then
       rsTemp.MoveLast
       rsTemp.MoveFirst
    End If
    
    fnRecordset = rsTemp.RecordCount
    Exit Function
    
SQLError:
    tfnErrHandler "fnRecordset" + IIf(sCalledFrom <> "", "," & sCalledFrom, ""), strSQL, bShowErrow
    fnRecordset = -1
    
    On Error GoTo 0
End Function

Private Function fnExecuteSQL(strSQL As String, sCalledFrom As String, Optional bShowError As Boolean = True) As Boolean
    On Error GoTo SQLError
    
    t_dbMainDatabase.ExecuteSQL strSQL
    
    fnExecuteSQL = True
    Exit Function

SQLError:
    tfnErrHandler "fnExecuteSQL" + IIf(sCalledFrom <> "", "," & sCalledFrom, ""), strSQL, bShowError
    fnExecuteSQL = False
End Function
'''''''''''''''''''''''''''''''
