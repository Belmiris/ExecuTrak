VERSION 5.00
Begin VB.Form AuthorityCode 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Authorization Code Input"
   ClientHeight    =   2040
   ClientLeft      =   732
   ClientTop       =   3552
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
      Left            =   1152
      ScaleHeight     =   444
      ScaleWidth      =   3432
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   1416
      Width           =   3432
      Begin VB.CommandButton cmdOK 
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
         Left            =   36
         TabIndex        =   6
         Top             =   24
         Width           =   1308
      End
      Begin VB.CommandButton cmdCancel 
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
         Left            =   2076
         TabIndex        =   7
         Top             =   24
         Width           =   1308
      End
   End
   Begin VB.PictureBox picAuthCode 
      BorderStyle     =   0  'None
      Height          =   816
      Left            =   756
      ScaleHeight     =   816
      ScaleWidth      =   4476
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   504
      Width           =   4476
      Begin VB.TextBox txtAuthCode 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.2
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
            Size            =   10.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   324
         Left            =   96
         TabIndex        =   8
         Top             =   12
         Width           =   4392
      End
   End
   Begin VB.PictureBox picCustomerInfo 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   348
      Left            =   780
      ScaleHeight     =   348
      ScaleWidth      =   4932
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   120
      Width           =   4932
      Begin VB.Label lblCustInfo 
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
         Left            =   36
         TabIndex        =   9
         Top             =   12
         Width           =   4740
      End
   End
   Begin VB.PictureBox picMessage 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.2
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
         ScaleWidth      =   492
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   288
         Width           =   492
      End
      Begin VB.Label lblMsg 
         Caption         =   "Do you want to continue?"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.2
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

Property Let CustomerNumber(ByVal lCust As Long)
    lblCustInfo.Caption = "Customer (" & CStr(lCust) & ") is on credit hold."
End Property

Private Function fnControlRight(ctrlTemp As Control) As Integer
    fnControlRight = ctrlTemp.Left + ctrlTemp.Width
End Function


Public Function SetAuthCode() As Boolean
    Const SUB_NAME = "SetAuthCode"
    
    Dim strSQL As String
    Dim rsTemp As Recordset
    
    strSQL = "SELECT parm_field FROM sys_parm" _
         & " WHERE parm_nbr = " & SYSPARM4007
    bAuthCodeSet = False
    On Error GoTo errQuery
    Set rsTemp = t_dbMainDatabase.OpenRecordset(strSQL, dbOpenSnapshot, dbSQLPassThrough)
    If rsTemp.RecordCount > 0 Then
        If Not IsNull(rsTemp!parm_field) Then
            sAuthCode = Trim(rsTemp!parm_field)
            If sAuthCode <> "" Then
                bAuthCodeSet = True
            End If
        End If
    End If
    subSetMyState
    SetAuthCode = True
    Exit Function
errQuery:
    tfnErrHandler SUB_NAME, strSQL
    SetAuthCode = False
End Function

Private Function fnValidCode() As Boolean
    fnValidCode = False
    If Trim(txtAuthCode.Text) = sAuthCode Then
        fnValidCode = True
    End If
End Function

Private Sub subEnableOK(bFlag As Boolean)
'    cmdOK.Enabled = bFlag
    cmdOK.Enabled = True  'Always enabled
    
End Sub


Private Sub subSetMyState()
    Me.Caption = App.Title
    If bAuthCodeSet Then
        cmdOK.Caption = "O&K"
        cmdCancel.Caption = "&Cancel"
        picCustomerInfo.Top = 156
        picMessage.Visible = False
        picAuthCode.Visible = True
        subEnableOK False
        txtAuthCode.Text = ""
    Else
        cmdOK.Caption = "&Yes"
        cmdCancel.Caption = "&No"
        picCustomerInfo.Top = 384
        picMessage.Visible = True
        picAuthCode.Visible = False
    End If

End Sub

Private Sub cmdCancel_Click()
    Me.Hide
    nAction = vbCancel
End Sub


Private Sub cmdOK_Click()
    If fnValidCode Then
        Me.Hide
        nAction = vbOK
    Else
        MsgBox "Invalid authorization code.", vbOKOnly + vbExclamation
        subSelectText txtAuthCode
        txtAuthCode.SetFocus
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


