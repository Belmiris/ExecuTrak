VERSION 5.00
Object = "{C75015E0-2232-11D3-B440-0060971E99AF}#1.0#0"; "factfrm.ocx"
Begin VB.Form objDrivePath 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4680
   ClientLeft      =   4365
   ClientTop       =   3045
   ClientWidth     =   5325
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4680
   ScaleWidth      =   5325
   ShowInTaskbar   =   0   'False
   Begin VB.DirListBox Dir1 
      Height          =   2544
      Left            =   216
      TabIndex        =   2
      Top             =   912
      Width           =   2772
   End
   Begin VB.DriveListBox Drive1 
      Height          =   324
      Left            =   216
      TabIndex        =   4
      Top             =   4092
      Width           =   2772
   End
   Begin FACTFRMLib.FactorFrame cmdOK 
      Height          =   396
      HelpContextID   =   4
      Left            =   3810
      TabIndex        =   5
      Top             =   3540
      Width           =   1305
      _Version        =   65536
      _ExtentX        =   2302
      _ExtentY        =   688
      _StockProps     =   77
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "&OK"
      CaptionPos      =   4
      PicturePos      =   3
      ShowFocusRect   =   -1  'True
      Style           =   3
      BorderWidth     =   4
      BeginProperty PanelFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin FACTFRMLib.FactorFrame cmdCancel 
      Height          =   396
      HelpContextID   =   5
      Left            =   3810
      TabIndex        =   6
      Top             =   4095
      Width           =   1305
      _Version        =   65536
      _ExtentX        =   2302
      _ExtentY        =   688
      _StockProps     =   77
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "&Cancel"
      CaptionPos      =   4
      PicturePos      =   3
      ShowFocusRect   =   -1  'True
      Style           =   3
      BorderWidth     =   4
      BeginProperty PanelFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblDrive 
      Caption         =   "Dri&ves:"
      Height          =   252
      Left            =   216
      TabIndex        =   3
      Top             =   3732
      Width           =   2172
   End
   Begin VB.Label lblFolderName 
      Caption         =   "D:\HARB\TLC\EXCHANGE\INBOUND\"
      Height          =   252
      Left            =   216
      TabIndex        =   1
      Top             =   552
      Width           =   4620
   End
   Begin VB.Label lblFolder 
      Caption         =   "&Folders:"
      Height          =   252
      Left            =   216
      TabIndex        =   0
      Top             =   180
      Width           =   2172
   End
End
Attribute VB_Name = "objDrivePath"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public nStatus As Integer  'either vbOK or vbCANCEL

Private Sub Form_Activate()
    lblFolderName = UCase(Dir1.Path)
    Dir1.SetFocus
End Sub

Private Sub Drive1_Change()
    On Error Resume Next
    Dir1.Path = Drive1.Drive
End Sub

Private Sub Dir1_Change()
    lblFolderName = UCase(Dir1.Path)
End Sub

Private Sub Form_Load()
    tfnCenterForm Me
    tfnDisableFormSystemClose Me
End Sub

Private Sub lblDrive_Click()
    Drive1.SetFocus
End Sub

Private Sub lblFolder_Click()
    Dir1.SetFocus
End Sub

Property Let DialogTitle(sTitle As String)
    Me.Caption = sTitle
End Property

Property Let InitDir(sInitDir As String)
    If Trim(sInitDir) = "" Then Exit Property
    
    On Error Resume Next
    Drive1.Drive = Left(sInitDir, InStr(sInitDir, ":"))
    Dir1.Path = sInitDir
End Property

Property Get InitDir() As String
    On Error Resume Next
    InitDir = Dir1.Path
End Property

Property Get Drive() As String
    On Error Resume Next
    Drive = Drive1.Drive
End Property

Property Get Cancel() As Boolean
    Cancel = nStatus = vbCancel
End Property

Property Get OK() As Boolean
    OK = nStatus = vbOK
End Property

Property Get Path() As String
    Path = lblFolderName
End Property

Public Sub ShowOpen()
    Me.Show vbModal
End Sub

Private Sub cmdOK_Click()
    nStatus = vbOK
    Me.Hide
End Sub

Private Sub cmdCancel_Click()
    nStatus = vbCancel
    Me.Hide
End Sub
