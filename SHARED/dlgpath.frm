VERSION 5.00
Begin VB.Form frmDlgPath 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Select Output Path"
   ClientHeight    =   3090
   ClientLeft      =   3150
   ClientTop       =   4050
   ClientWidth     =   5205
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   10.5
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3090
   ScaleWidth      =   5205
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdCancel 
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
      HelpContextID   =   15
      Left            =   3540
      TabIndex        =   6
      Top             =   2424
      Width           =   1476
   End
   Begin VB.CommandButton cmdOk 
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
      HelpContextID   =   16
      Left            =   3540
      TabIndex        =   5
      Top             =   1884
      Width           =   1476
   End
   Begin VB.DriveListBox drvDrive 
      Height          =   324
      HelpContextID   =   2101
      Left            =   3384
      TabIndex        =   3
      Top             =   1080
      Width           =   1740
   End
   Begin VB.DirListBox dirPath 
      Height          =   2292
      HelpContextID   =   2100
      Left            =   48
      TabIndex        =   0
      Top             =   696
      Width           =   3216
   End
   Begin VB.Label lblCaptions 
      Caption         =   "Drive"
      Height          =   300
      Index           =   0
      Left            =   3384
      TabIndex        =   4
      Top             =   756
      Width           =   1644
   End
   Begin VB.Label lblPath 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   324
      Left            =   60
      TabIndex        =   1
      Top             =   336
      Width           =   5076
   End
   Begin VB.Label lblCaptions 
      Caption         =   "Path "
      Height          =   276
      Index           =   1
      Left            =   72
      TabIndex        =   2
      Top             =   72
      Width           =   5064
   End
End
Attribute VB_Name = "frmDlgPath"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
    Private m_nAction As Integer

Property Get Action() As Integer
    Action = m_nAction
End Property

Private Sub cmdCancel_Click()
    m_nAction = vbCancel
    Me.Hide
End Sub

Private Sub cmdOK_Click()
    m_nAction = vbOK
    Me.Hide
End Sub


Private Sub dirPath_Change()
    lblPath.Caption = dirPath.Path
End Sub

Private Sub drvDrive_Change()
    On Error Resume Next
    dirPath.Path = drvDrive.Drive
End Sub


Private Sub Form_Activate()
    
    'dirPath.d = drvDrive.Drive
    lblPath.Caption = dirPath.Path

End Sub

Private Sub Form_Load()
    subCenterForm Me
End Sub


