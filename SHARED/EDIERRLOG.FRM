VERSION 5.00
Object = "{C75015E0-2232-11D3-B440-0060971E99AF}#1.0#0"; "FACTFRM.OCX"
Begin VB.Form EDIErrLog 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Error message"
   ClientHeight    =   3744
   ClientLeft      =   36
   ClientTop       =   264
   ClientWidth     =   5604
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3744
   ScaleWidth      =   5604
   StartUpPosition =   3  'Windows Default
   Begin FACTFRMLib.FactorFrame cmdExit 
      Height          =   396
      Left            =   4068
      TabIndex        =   3
      Top             =   3312
      Width           =   1512
      _Version        =   65536
      _ExtentX        =   2667
      _ExtentY        =   698
      _StockProps     =   77
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   10.19
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "E&xit"
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
   Begin FACTFRMLib.FactorFrame cmdIgnore 
      Height          =   396
      Left            =   2352
      TabIndex        =   2
      Top             =   3312
      Width           =   1512
      _Version        =   65536
      _ExtentX        =   2660
      _ExtentY        =   688
      _StockProps     =   77
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   10.05
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "&Ignore"
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
   Begin VB.TextBox txtMsg 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2904
      Left            =   84
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   348
      Width           =   5472
   End
   Begin VB.Label Label1 
      Caption         =   "Error Message"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.2
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   96
      TabIndex        =   1
      Top             =   72
      Width           =   1536
   End
End
Attribute VB_Name = "EDIErrLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_nAction As Integer

Private Sub cmdExit_Click()
    m_nAction = vbCancel
    Me.Hide
End Sub

Private Sub cmdIgnore_Click()
    m_nAction = vbIgnore
    Me.Hide
End Sub

Private Sub Form_Load()
    tfnDisableFormSystemClose Me
    tfnCenterForm Me, frmECF810SU
End Sub

Property Get Action() As Integer
    Action = m_nAction
End Property
