VERSION 5.00
Object = "{C75015E0-2232-11D3-B440-0060971E99AF}#1.0#0"; "FACTFRM.OCX"
Object = "{01028C21-0000-0000-0000-000000000046}#4.0#0"; "TG32OV.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmZZSESLIP 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Day off Slip Days Entry"
   ClientHeight    =   6348
   ClientLeft      =   1056
   ClientTop       =   1956
   ClientWidth     =   8964
   DrawMode        =   14  'Copy Pen
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
   ScaleHeight     =   6348
   ScaleWidth      =   8964
   Begin FACTFRMLib.FactorFrame ffraStatusbar 
      Height          =   360
      Left            =   0
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   5988
      Width           =   8964
      _Version        =   65536
      _ExtentX        =   15811
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
      FMName          =   "POFSELGP"
      CaptionPos      =   1
      FloodColor      =   16744576
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
      Height          =   5472
      Left            =   0
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   504
      Width           =   8940
      _Version        =   65536
      _ExtentX        =   15769
      _ExtentY        =   9652
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
      FMName          =   "POFSELGP"
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
      Begin FACTFRMLib.FactorFrame cmdAddBtn 
         Height          =   396
         HelpContextID   =   10
         Left            =   132
         TabIndex        =   1
         Top             =   4920
         Width           =   1308
         _Version        =   65536
         _ExtentX        =   2302
         _ExtentY        =   688
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
         Caption         =   "&Add"
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
      Begin FACTFRMLib.FactorFrame cmdEditBtn 
         Height          =   396
         HelpContextID   =   11
         Left            =   1608
         TabIndex        =   2
         Top             =   4920
         Width           =   1308
         _Version        =   65536
         _ExtentX        =   2302
         _ExtentY        =   688
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
         Caption         =   "&Edit"
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
      Begin FACTFRMLib.FactorFrame cmdDeleteBtn 
         Height          =   396
         HelpContextID   =   12
         Left            =   3084
         TabIndex        =   12
         Top             =   4920
         Width           =   1308
         _Version        =   65536
         _ExtentX        =   2302
         _ExtentY        =   688
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
         Caption         =   "&Delete"
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
      Begin FACTFRMLib.FactorFrame cmdUpdateInsertBtn 
         Height          =   396
         HelpContextID   =   13
         Left            =   4560
         TabIndex        =   9
         Top             =   4920
         Width           =   1308
         _Version        =   65536
         _ExtentX        =   2302
         _ExtentY        =   688
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
         Caption         =   "&Update"
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
      Begin FACTFRMLib.FactorFrame cmdRefreshSelectBtn 
         Height          =   396
         HelpContextID   =   14
         Left            =   6036
         TabIndex        =   10
         Top             =   4920
         Width           =   1308
         _Version        =   65536
         _ExtentX        =   2302
         _ExtentY        =   688
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
         Caption         =   "&Refresh"
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
         Top             =   4920
         Width           =   1308
         _Version        =   65536
         _ExtentX        =   2302
         _ExtentY        =   688
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
      Begin FACTFRMLib.FactorFrame efraOuterFrame 
         Height          =   4608
         Left            =   132
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   144
         Width           =   8676
         _Version        =   65536
         _ExtentX        =   15303
         _ExtentY        =   8128
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
         FMName          =   "POFSELGP"
         BeginProperty PanelFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin DBTrueGrid.TDBGrid tblSlip 
            Height          =   3360
            HelpContextID   =   604
            Left            =   84
            OleObjectBlob   =   "ZZSESLIP.frx":0000
            TabIndex        =   8
            Top             =   1140
            Width           =   8520
         End
         Begin FACTFRMLib.FactorFrame efraEmployee 
            Height          =   900
            Left            =   84
            TabIndex        =   18
            TabStop         =   0   'False
            Top             =   120
            Width           =   8520
            _Version        =   65536
            _ExtentX        =   15028
            _ExtentY        =   1587
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
            BevelOuter      =   6
            BeginProperty PanelFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Begin VB.TextBox txtEmpName 
               BackColor       =   &H00FFFFFF&
               DataSource      =   "datVendor"
               ForeColor       =   &H00000000&
               Height          =   360
               HelpContextID   =   602
               Left            =   2604
               TabIndex        =   5
               Tag             =   "pn_name"
               Top             =   372
               Width           =   5448
            End
            Begin VB.TextBox txtEmployee 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00000000&
               Height          =   360
               HelpContextID   =   600
               Left            =   84
               TabIndex        =   3
               Tag             =   "pn_alt"
               Top             =   360
               Width           =   2100
            End
            Begin FACTFRMLib.FactorFrame cmdEmpName 
               Height          =   360
               HelpContextID   =   603
               Left            =   8064
               TabIndex        =   6
               TabStop         =   0   'False
               Tag             =   "Run #3"
               Top             =   372
               Width           =   360
               _Version        =   65536
               _ExtentX        =   635
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
               CaptionPos      =   4
               Picture         =   "ZZSESLIP.frx":12EB
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
            Begin FACTFRMLib.FactorFrame cmdEmployee 
               Height          =   360
               HelpContextID   =   601
               Left            =   2196
               TabIndex        =   4
               TabStop         =   0   'False
               Tag             =   "Run #3"
               Top             =   372
               Width           =   360
               _Version        =   65536
               _ExtentX        =   635
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
               CaptionPos      =   4
               Picture         =   "ZZSESLIP.frx":13FD
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
            Begin VB.Label lblEmpName 
               Caption         =   "Employee Name"
               Height          =   252
               Left            =   2616
               TabIndex        =   20
               Top             =   108
               Width           =   1956
            End
            Begin VB.Label lblEmpNumber 
               Caption         =   "Employee Number"
               Height          =   252
               Left            =   96
               TabIndex        =   19
               Top             =   108
               Width           =   1632
            End
         End
         Begin FACTFRMLib.FactorFrame cmdSearch 
            Height          =   336
            HelpContextID   =   605
            Left            =   2424
            TabIndex        =   21
            Top             =   1260
            Width           =   984
            _Version        =   65536
            _ExtentX        =   1736
            _ExtentY        =   593
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
            CaptionPos      =   4
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
      End
   End
   Begin FACTFRMLib.FactorFrame efraToolBar 
      Height          =   468
      Left            =   0
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   0
      Width           =   8964
      _Version        =   65536
      _ExtentX        =   15811
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
      FMName          =   "ZZSESLIP"
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
         Height          =   312
         Left            =   60
         TabIndex        =   0
         Top             =   84
         Width           =   3396
         _ExtentX        =   5990
         _ExtentY        =   550
         ButtonWidth     =   508
         ButtonHeight    =   466
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
   Begin DBTrueGrid.TDBGrid tblComboDropdown 
      Bindings        =   "ZZSESLIP.frx":150F
      Height          =   2484
      Left            =   36
      OleObjectBlob   =   "ZZSESLIP.frx":152E
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   660
      Width           =   3756
   End
   Begin DBTrueGrid.TDBGrid tblDropDown 
      Bindings        =   "ZZSESLIP.frx":2812
      Height          =   1008
      Left            =   324
      OleObjectBlob   =   "ZZSESLIP.frx":282C
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   708
      Width           =   2604
   End
   Begin VB.Data datDropDown 
      Caption         =   "DropDown"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   324
      Left            =   564
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1824
      Visible         =   0   'False
      Width           =   1908
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
      Begin VB.Menu mnuOptionsSep10 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAdd 
         Caption         =   "&Add"
         HelpContextID   =   4
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "&Edit"
         HelpContextID   =   5
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "&Delete"
         Enabled         =   0   'False
         HelpContextID   =   6
      End
      Begin VB.Menu mnuUpdateInsert 
         Caption         =   "&Update"
         Enabled         =   0   'False
         HelpContextID   =   2
      End
      Begin VB.Menu mnuRefreshSelect 
         Caption         =   "&Refresh"
         Enabled         =   0   'False
         HelpContextID   =   3
      End
      Begin VB.Menu mnuOptionsSep20 
         Caption         =   "-"
         Visible         =   0   'False
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
Attribute VB_Name = "frmZZSESLIP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit


Private t_bStartupFlag As Boolean 'optional startup flag

Private t_bUpdateTable As Boolean 'update data flag

Private t_nFormMode As Integer         'global used to track the current form operating mode

Private Const IDLE_MODE As Integer = 0 'idle mode activates the NoDrop Cursor
Private Const ADD_MODE As Integer = 1  'flag set when in the add mode
Private Const EDIT_MODE As Integer = 2 'flag set when in the edit mode
Private Const DELETE_MODE As Integer = 3 'flag set when in the delete mode

'========================
'Standard Button Captions
'========================
Private Const t_szCAPTION_INSERT As String = "&Insert"
Private Const t_szCAPTION_UPDATE As String = "&Update"
Private Const t_szCAPTION_REFRESH As String = "&Refresh"
Private Const t_szCAPTION_SELECT As String = "&Select"
Private Const t_szCAPTION_CANCEL As String = "&Cancel"
Private Const t_szCAPTION_EXIT As String = "E&xit"

'==========================
'Status Bar Default Strings
'==========================
Private Const t_szCLEAR As String = ""
Private Const t_szADD As String = "Select Add, Edit or Exit"
Private Const t_szEDIT As String = "Select Add, Edit or Exit"
Private Const t_szDELETE As String = "Delete"
Private Const t_szINSERT As String = "Insert"
Private Const t_szUPDATE As String = "Update"
Private Const t_szREFRESH As String = "Refresh"
Private Const t_szSELECT As String = "Select"
Private Const t_szEXIT As String = "Exit"
Private Const t_szCANCEL As String = "Cancel"
Private Const t_szADDEDIT As String = "Select Add, Edit or Exit"

Private Const t_szPRINT As String = "Print"
Private Const t_szCOPYFROM As String = "Copy From"
Private Const t_szHELP As String = "Help"

Private Const DATA_INIT = 0
Private Const DATA_LOADED = 1
Private Const DATA_CHANGED = 2

Private nDataStatus As Integer
Private tgmSlip As clsTGSpreadSheet
Private cvalidate As cValidateInput
Private tgfDropdown As clsFloatingDropDown

Private Const ColPrftCtr As Integer = 0
Private Const colPrftDesc As Integer = 1
Private Const ColSlipDate As Integer = 2
Private Const colShiftNum As Integer = 3
Private Const nMaxCol As Integer = 3


Private Sub cmdDeleteBtn_Click()
    Dim sBCode As String
    Dim lRow As Long
    Dim strSQL As String
    
    If tgmSlip.RowCount < 1 Then
        subEnableDeleteBtn False
        Exit Sub
    End If
    
    lRow = tgmSlip.GetCurrentRowNumber
    
    If t_nFormMode = EDIT_MODE Then
        
        If Not tfnCancelExit("Are you sure you want to delete the current record?") Then
            Exit Sub
        End If
            
        strSQL = "DELETE FROM bonus_day_off_slip "
        strSQL = strSQL & " WHERE bd_empno = " & tfnRound(txtEmployee)
        strSQL = strSQL & " AND bd_slip_date = " & tfnDateString(tgmSlip.CellValue(ColSlipDate, lRow), True)
        strSQL = strSQL & " AND bd_prft_ctr = " & tfnRound(tgmSlip.CellValue(ColPrftCtr, lRow))
        strSQL = strSQL & " AND bd_shift = " & tfnRound(tgmSlip.CellValue(colShiftNum, lRow))
        
        If Not fnExecuteSQL(strSQL) Then
            Exit Sub
        End If
        
        tgmSlip.DeleteRow
    End If
    
    If tgmSlip.RowCount < 1 Then
        subEnableDeleteBtn False
        
        If t_nFormMode = EDIT_MODE Then
            cmdExitCancelBtn_Click
        End If
        
    End If
    
End Sub

Private Sub cmdSearch_Click()
    tgfDropdown.ButtonClick
End Sub

Private Sub cmdSearch_LostFocus()
    tgfDropdown.LostFocus cmdSearch
End Sub

Private Sub efraBackground_GotFocus()
    cvalidate.GotFocus efraBackground
End Sub

'===========
'Form Events
'===========
Private Sub Form_Initialize() 'called before Form_Load
    t_bStartupFlag = True
    t_bUpdateTable = False
    
    t_nFormMode = IDLE_MODE
    
    CRLF = Chr(10) + Chr(13)

    ' ** change the help file for the application
    App.HelpFile = szHelp7_11
End Sub

Private Sub Form_Unload(CANCEL As Integer)
    
    On Error Resume Next
    
    Set tgfDropdown = Nothing
    Set objErrHandler = Nothing
    
    Set objCurrTabControl = Nothing
    
    Unload frmContext
    Unload frmAbout
    
    'project local database object variables
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
        
        subInitErrorHandler   ' Setup Error Control
         
        tfnUpdateVersion
        
        If Not fnCreateSearchTable("prm_empno", "prm_empname") Then
            MsgBox "Failed to create temporary employee table", vbCritical
            Unload Me
            End
        End If
        
    #End If
    
    ' disable system menu Close and Close icon on form
    tfnDisableFormSystemClose Me
    subSetupToolBar
    
    tmrKeyBoard.Enabled = False
    tfnCenterForm Me
    
    subSetupCombos
    subInitValidation
    subInitializeSpreadSheet
    subSetupDropDown
    
    Me.Show
    subShowHourGlass
    
    tfnSetInitializingMessage
    Me.Enabled = False
    
    t_bStartupFlag = False
    
    Me.Enabled = True
    tfnResetScreen 'set the default screen
    tmrKeyBoard.Enabled = True
    tfnSetStatusBarMessage t_szADDEDIT
    subHideHourGlass
    
End Sub

Private Sub Form_GotFocus()
    tmrKeyBoard.Enabled = True
End Sub

Private Sub Form_LostFocus()
    tmrKeyBoard.Enabled = False
End Sub


'====================
'Screen Button Events
'====================
Private Sub cmdAddBtn_GotFocus()
    tfnSetStatusBarMessage t_szADDEDIT
End Sub

Private Sub cmdEditBtn_GotFocus()
    tfnSetStatusBarMessage t_szADDEDIT
End Sub

Private Sub cmdDeleteBtn_GotFocus()
    tfnSetStatusBarMessage t_szDELETE
End Sub

Private Sub cmdUpdateInsertBtn_GotFocus()
    
    If cmdUpdateInsertBtn.Caption = t_szCAPTION_UPDATE Then
        tfnSetStatusBarMessage t_szUPDATE
    Else
        tfnSetStatusBarMessage t_szINSERT
    End If
    
End Sub

Private Sub cmdRefreshSelectBtn_GotFocus()
    
    If cmdRefreshSelectBtn.Caption = t_szCAPTION_REFRESH Then
        tfnSetStatusBarMessage t_szREFRESH
    Else
        tfnSetStatusBarMessage t_szSELECT
    End If
    
End Sub

Private Sub cmdExitCancelBtn_GotFocus()
    
    If cmdExitCancelBtn.Caption = t_szCAPTION_EXIT Then
        tfnSetStatusBarMessage t_szEXIT
    Else
        tfnSetStatusBarMessage t_szCANCEL
    End If
    
End Sub

Private Sub cmdAddBtn_Click()
    
    #If PROTOTYPE Then
        Exit Sub
    #End If
    
    t_nFormMode = ADD_MODE
    't_bStartupFlag = False
    subEnableOutFrame True
    subEnableAddBtn False
    subEnableEditBtn False
    subEnableUpdateBtn False
    subEnableRefreshBtn False
    subEnableCancel True
    subSetupSlipGrid
    subSetFocus txtEmployee
    
    cmdUpdateInsertBtn.Caption = t_szCAPTION_INSERT
    mnuUpdateInsert.Caption = t_szCAPTION_INSERT
    
End Sub

Private Sub subSetFocus(v As Control)
    On Error Resume Next
    v.SetFocus
End Sub

Private Sub cmdEditBtn_Click()
    #If PROTOTYPE Then
        Exit Sub
    #End If
    
    t_nFormMode = EDIT_MODE
    subEnableOutFrame True
    subEnableAddBtn False
    subEnableEditBtn False
    subEnableUpdateBtn False
    subEnableRefreshBtn False
    subEnableCancel True
    subSetupSlipGrid
    subSetFocus txtEmployee
    
    cmdUpdateInsertBtn.Caption = t_szCAPTION_UPDATE
    mnuUpdateInsert.Caption = t_szCAPTION_UPDATE
End Sub

Private Sub cmdUpdateInsertBtn_Click()
    subShowHourGlass
    tfnSetStatusBarCorrect "Inserting data, please wait..."
    
    If fnInsertData() Then
        tfnUnlockRow
        tfnResetScreen
    End If
    
    subHideHourGlass
End Sub

Private Sub cmdExitCancelBtn_Click()

    If cmdExitCancelBtn.Caption = t_szCAPTION_CANCEL Then
        cmdTBCancel_Click
    Else
        cmdTBExit_Click
    End If
    
End Sub

Private Sub cmdRefreshSelectBtn_Click()
    
    If Not tfnCancelExit(t_szREFRESH_MESSAGE) Then
        Exit Sub
    End If
    
    subShowHourGlass
    fnLoadData
    nDataStatus = DATA_INIT
    subEnableRefreshBtn False
    subEnableUpdateBtn False
    subHideHourGlass
End Sub

'=====================
'Toolbar Button Events
'=====================
Private Sub cmdTBPrint_Click()
    MsgBox "Print Function"
End Sub

Private Sub cmdTBCopyFrom_Click()
    MsgBox "Copy From Function"
End Sub

Private Sub cmdTBCancel_Click()
    
    If nDataStatus = DATA_CHANGED Then
        If Not tfnCancelExit(t_szCANCEL_MESSAGE) Then
            Exit Sub
        End If
    End If
    
    tfnResetScreen
    
End Sub

Private Sub cmdTBExit_Click()
    
    If nDataStatus = DATA_CHANGED Then
        If Not tfnCancelExit(t_szEXIT_MESSAGE) Then
            Exit Sub
        End If
    End If
    
    tfnUnlockRow
    
    Unload Me
    
End Sub

'============
'Menu  Events
'============
Private Sub mnuExit_Click()
    cmdTBExit_Click
End Sub

Private Sub mnuPrint_Click()
    cmdTBPrint_Click
End Sub

Private Sub mnuAdd_Click()
    cmdAddBtn_Click
End Sub

Private Sub mnuEdit_Click()
    cmdEditBtn_Click
End Sub

Private Sub mnuDelete_Click()
    cmdDeleteBtn_Click
End Sub

Private Sub mnuRefreshSelect_Click()
    cmdRefreshSelectBtn_Click
End Sub

Private Sub mnuUpdateInsert_Click()
    cmdUpdateInsertBtn_Click
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
  
  If t_nFormMode <> IDLE_MODE And TypeOf Screen.ActiveControl Is Textbox Then
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
    cmdTBCancel_Click
End Sub

Private Sub mnuCopy_Click()
    Clipboard.SetText Screen.ActiveControl.Text, vbCFText
End Sub

Private Sub mnuCut_Click()
    Clipboard.SetText Screen.ActiveControl.Text, vbCFText
    Screen.ActiveControl.Text = ""
End Sub

Private Sub mnuPaste_Click()
  If TypeOf Screen.ActiveControl Is Textbox Then
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
Private Sub tfnSetStatusBarMessage(szMessage As String)
    
    If t_bStartupFlag Then
        Exit Sub
    End If
    
    ffraStatusbar.ForeColor = STANDARD_TEXT_COLOR
    ffraStatusbar.Font.Bold = False
    ffraStatusbar.Caption = szMessage
    ffraStatusbar.Refresh

End Sub

Public Sub tfnSetInitializingMessage()

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
    
    t_nFormMode = IDLE_MODE
    subShowHourGlass
    subEnableEmployee True
    subEnableOutFrame False
    subEnableAddBtn True
    subEnableEditBtn True
    subEnableUpdateBtn False
    subEnableDeleteBtn False
    subEnableRefreshBtn False
    subEnableCancel False
    tfnUnlockRow
    
    cmdUpdateInsertBtn.Caption = t_szCAPTION_UPDATE
    mnuUpdateInsert.Caption = t_szCAPTION_UPDATE
    
    cmdRefreshSelectBtn.Caption = t_szCAPTION_REFRESH
    mnuRefreshSelect.Caption = t_szCAPTION_REFRESH
    
    frmContext.ButtonEnabled(PRINT_UP) = False
    frmContext.ButtonEnabled(COPY_UP) = False
    
    mnuExit.Enabled = True
    mnuPrint.Enabled = False
    
    mnuCopy.Enabled = False
    mnuPaste.Enabled = False
    
    txtEmployee = ""
    txtEmpName = ""
    cvalidate.ResetFlags
    tgmSlip.ClearData
    nDataStatus = DATA_INIT
    subSetFocus cmdAddBtn
    subHideHourGlass
    DoEvents
    
End Sub

Private Sub tblDropDown_SelChange(CANCEL As Integer)
    CANCEL = True
End Sub

Private Sub tblSlip_AfterColEdit(ByVal ColIndex As Integer)
    tgmSlip.AfterColEdit ColIndex
    nDataStatus = DATA_CHANGED
    subCheckDataStatus
End Sub

Private Sub tblSlip_BeforeColEdit(ByVal ColIndex As Integer, ByVal KeyAscii As Integer, CANCEL As Integer)
    Dim lCurRow As Long
    
    tgmSlip.BeforeColEdit ColIndex, KeyAscii, CANCEL
    
    subEnableUpdateBtn False
    
    If t_nFormMode = ADD_MODE Then
        lCurRow = tgmSlip.GetCurrentRowNumber
        
        If tgmSlip.CellValue(colShiftNum, lCurRow) = "" Then
            tgmSlip.CellValue(colShiftNum, lCurRow) = 1
        End If
        
    End If
    
End Sub

Private Sub tblSlip_Change()
    tgmSlip.Change
    
    If tblSlip.col = ColPrftCtr Then
        tblSlip.Columns(colPrftDesc).Text = ""
    End If
    
End Sub

Private Sub tblSlip_Click()
    tgfDropdown.TableClick tblSlip
End Sub

Private Sub tblSlip_FirstRowChange()
    tgfDropdown.FirstRowChange
End Sub

Private Sub tblSlip_GotFocus()
    tgfDropdown.GotFocus
    tgmSlip.GotFocus
    
    If t_nFormMode = EDIT_MODE Then
        tfnSetStatusBarMessage "Day Off Slip Days"
    End If
    
    If tgfDropdown.ValidSelection Then
        nDataStatus = DATA_CHANGED
    End If
    
End Sub

Private Sub tblSlip_KeyDown(KeyCode As Integer, Shift As Integer)
    tgfDropdown.KeyDown tblSlip, KeyCode
    tgmSlip.KeyDown KeyCode, Shift
End Sub

Private Sub tblSlip_KeyPress(KeyAscii As Integer)
    
    If tgfDropdown.ValidSelection Then
        nDataStatus = DATA_CHANGED
    End If
    
    If Not tgmSlip.Keypress(KeyAscii) Then
        KeyAscii = 0
    End If
    
End Sub

Private Sub tblSlip_LostFocus()
    tgmSlip.LostFocus
    tgfDropdown.LostFocus tblSlip
End Sub

Private Sub tblSlip_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    Dim nCurCol As Integer
    Dim lCurRow As Long
    
    If tgfDropdown.RowColChange(tblSlip) Then
        tgmSlip.RowColChange LastRow, LastCol
    End If
    
    If t_nFormMode = ADD_MODE Then
        nCurCol = tblSlip.col
        lCurRow = tgmSlip.GetCurrentRowNumber
        
        If tgmSlip.ValidCell(ColPrftCtr, lCurRow) Then
            
            If nCurCol = colPrftDesc Then
                tgmSlip.GotoNextColumn LastCol
            End If
            
        End If
        
    End If
End Sub

Private Sub tblSlip_Scroll(CANCEL As Integer)
    tgfDropdown.Scroll
End Sub

Private Sub tblSlip_SelChange(CANCEL As Integer)
    CANCEL = True
End Sub

Private Sub tblSlip_UnboundReadData(ByVal RowBuf As DBTrueGrid.RowBuffer, StartLocation As Variant, ByVal ReadPriorRows As Boolean)
    tgmSlip.ReadData RowBuf, StartLocation, ReadPriorRows
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
            Set .LocalDatabase = tfnOpenLocalDatabase()
        End With
        
    End If
    
End Sub

Private Sub subEnableAddBtn(bOnOff As Boolean)
    cmdAddBtn.Enabled = bOnOff
    mnuAdd.Enabled = bOnOff
End Sub

Private Sub subEnableEditBtn(bOnOff As Boolean)
    cmdEditBtn.Enabled = bOnOff
    mnuEdit.Enabled = bOnOff
End Sub

Private Sub subEnableDeleteBtn(bOnOff As Boolean)
    cmdDeleteBtn.Enabled = bOnOff
    mnuDelete.Enabled = bOnOff
End Sub

Private Sub subEnableRefreshBtn(bOnOff As Boolean)
    cmdRefreshSelectBtn.Enabled = bOnOff
    mnuRefreshSelect.Enabled = bOnOff
End Sub

Private Sub subEnableUpdateBtn(bOnOff As Boolean)
    cmdUpdateInsertBtn.Enabled = bOnOff
    mnuUpdateInsert.Enabled = bOnOff
End Sub

Private Sub subEnableCancel(bStatus As Boolean)
    mnuCancel.Enabled = bStatus
    frmContext.ButtonEnabled(CANCEL_UP) = bStatus
    
    If bStatus Then
        cmdExitCancelBtn.Caption = t_szCAPTION_CANCEL
    Else
        cmdExitCancelBtn.Caption = t_szCAPTION_EXIT
    End If
    
End Sub

Private Sub Form_Resize()
    frmContext.FormResize
End Sub

Private Sub subSetupToolBar()
    
    With frmContext
        .BeginSetupToolbar Me
        '.AddButton "Add &Employee", EMP_MST_UP
        .EndSetupToolbar
    
        .HelpFile = szHelpFileName
    End With
    
End Sub

Public Sub TBButtonCallBack(ByVal nID As Integer)
    
    Select Case nID
        Case CANCEL_UP
            cmdTBCancel_Click
        Case EXIT_UP
            cmdTBExit_Click
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

Private Sub subInitializeSpreadSheet()
    Dim myWidth As Variant
    Dim myField As Variant
    Dim i As Integer
    
    myWidth = Array(0.2, 0.4, 0.2, 0.2)
    myField = Array("bd_prft_ctr", "prft_name", "bd_slip_date", "bd_shift")
    
    While tblSlip.Columns.Count > 0
        tblSlip.Columns.Remove 0
    Wend
    
    For i = 0 To UBound(myWidth)
        tblSlip.Columns.Add i
        
        With tblSlip.Columns(i)
            .Width = myWidth(i) * tblSlip.Width - 50
            .DataField = myField(i)
            .Visible = True
        End With
        
    Next
    
    tblSlip.Columns(ColPrftCtr).Caption = "Profit Center"
    tblSlip.Columns(colPrftDesc).Caption = "Profit Center Description"
    tblSlip.Columns(ColSlipDate).Caption = "Slip Date"
    tblSlip.Columns(colShiftNum).Caption = "Shift Number"

    Set tgmSlip = New clsTGSpreadSheet
    
    With tgmSlip
        Set .Form = Me
        Set .StatusBar = ffraStatusbar
        Set .Table = tblSlip
        
        .AddEditColumn ColPrftCtr, "Enter a Profit center", szIntegerPattern
        .AddESCColumn colPrftDesc, ColPrftCtr
        .AddEditColumn colPrftDesc, "Enter Profit center description ", "^P{1,40}$"
        .AddEditColumn ColSlipDate, "Enter Slip Date", szDatePattern
        .AddEditColumn colShiftNum, "Enter Shift Number", szIntegerPattern
        .AllowAddNew = True
    End With
    
End Sub

Private Sub tblDropDown_Click()
    tgfDropdown.TableClick tblDropDown
End Sub

Private Sub tblDropDown_KeyPress(KeyAscii As Integer)
    tgfDropdown.Keypress tblDropDown, KeyAscii
End Sub

Private Sub tblDropDown_LostFocus()
    tgfDropdown.LostFocus tblDropDown
End Sub

Private Sub tblDropDown_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    tgfDropdown.MouseUp y
End Sub

Private Sub tblDropDown_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    tgfDropdown.RowColChange tblDropDown
End Sub

Private Function fnLoadData() As Boolean
    Const SUB_NAME As String = "fnLoadData"
    Dim strSQL As String
    Dim rsTemp As Recordset
    
    fnLoadData = False
    
    strSQL = "SELECT bd_prft_ctr, prft_name, bd_slip_date, bd_shift "
    strSQL = strSQL & " FROM sys_prft_ctr, bonus_day_off_slip "
    strSQL = strSQL & " WHERE prft_ctr = bd_prft_ctr AND bd_empno = " & tfnRound(txtEmployee)
    strSQL = strSQL & " ORDER BY bd_prft_ctr, bd_slip_date, bd_shift"
        
    If GetRecordSet(rsTemp, strSQL, , SUB_NAME) < 0 Then
        MsgBox "Failed to access database to get records", vbCritical
        Exit Function
    End If
    
    If rsTemp.RecordCount = 0 Then
        MsgBox "No bonus record found for the employee", vbExclamation
        Exit Function
    End If
    
    tgmSlip.FillWithRecordset rsTemp, True
    subEnableDeleteBtn True
    
    fnLoadData = True
        
End Function

Private Sub txtEmployee_Change()
    cvalidate.Change txtEmployee
    tgcDropdown.Change txtEmployee
End Sub

Private Sub txtEmployee_GotFocus()
    
    tgcDropdown.GotFocus txtEmployee
    cvalidate.GotFocus txtEmployee
    
    If tgcDropdown.SingleRecordSelected Then
        
        If fnSwitchPhase() Then
            subSetFocus tblSlip
        End If
        
    End If
    
End Sub

Private Sub txtEmployee_KeyPress(KeyAscii As Integer)
    Dim bKeyCode As Boolean
    
    If KeyAscii = vbKeyReturn Then
        tgcDropdown.ComboSQL(txtEmployee) = fnSetEmpComboSQL
        subShowHourGlass
    End If
    
    bKeyCode = tgcDropdown.Keypress(txtEmployee, KeyAscii)
    
    If Not bKeyCode Then
    
        If KeyAscii = vbKeyReturn Then
          
            If tgcDropdown.SingleRecordSelected Then
                
                If fnSwitchPhase() Then
                    subSetFocus tblSlip
                End If
                
            End If
          
          KeyAscii = 0
          
       End If
       
    Else
        cvalidate.Keypress txtEmployee, KeyAscii
    End If
    
    subHideHourGlass
End Sub

Private Sub txtEmployee_LostFocus()
    tgcDropdown.LostFocus txtEmployee
    cvalidate.LostFocus txtEmployee, cmdEmployee, txtEmpName, cmdEmpName, tblComboDropdown
End Sub

Private Sub txtEmployee_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    'frmContext.MouseDown Button, EMP_MST_UP
End Sub

Private Sub txtEmpName_Change()
    tgcDropdown.Change txtEmpName
End Sub

Private Sub txtEmpName_GotFocus()
    tfnSetStatusBarMessage "Enter Employee Name"
    tgcDropdown.GotFocus txtEmpName

    If tgcDropdown.SingleRecordSelected Then
        
        If fnSwitchPhase() Then
            subSetFocus tblSlip
        End If
    
    End If
    
End Sub

Private Sub txtEmpName_KeyPress(KeyAscii As Integer)
    Dim bKeyCode As Boolean
    
    If KeyAscii = vbKeyReturn Then
        tgcDropdown.ComboSQL(txtEmpName) = fnSetEmpComboSQL
        subShowHourGlass
    End If
    
    bKeyCode = tgcDropdown.Keypress(txtEmpName, KeyAscii)
    
    If Not bKeyCode Then
        
        If KeyAscii = vbKeyReturn Then
            
            If tgcDropdown.SingleRecordSelected Then
                
                If fnSwitchPhase() Then
                    subSetFocus tblSlip
                End If
            
            End If
            
            KeyAscii = 0
        End If
        
    End If
    
    subHideHourGlass
End Sub

Private Sub cmdEmployee_Click()
    subShowHourGlass
    tgcDropdown.ComboSQL(txtEmployee) = fnSetEmpComboSQL
    tgcDropdown.Click cmdEmployee
    subHideHourGlass
End Sub

Private Sub cmdEmpName_Click()
    subShowHourGlass
    tgcDropdown.ComboSQL(txtEmpName) = fnSetEmpComboSQL
    tgcDropdown.Click cmdEmpName
    subHideHourGlass
End Sub

Private Sub subInitValidation()
    Set cvalidate = New cValidateInput
    Set cvalidate.Form = Me
    Set cvalidate.StatusBar = ffraStatusbar
    
    cvalidate.AddEditBox txtEmployee, "Enter Employee Number"
    cvalidate.MinTabIndex = tbToolbar.TabIndex - 1
    cvalidate.MaxTabIndex = txtEmpName.TabIndex + 1
    cvalidate.ESCControl = cmdExitCancelBtn
    Set cvalidate.ControlForFocus = efraBackground
    Set cvalidate.LastBox = txtEmployee
End Sub

Private Function fnSetEmpComboSQL() As String
    Dim strSQL As String
    
    strSQL = "SELECT DISTINCT prm_empno, prm_empname FROM sTmpCustTable"
    
    If t_nFormMode = EDIT_MODE Then
        strSQL = strSQL & ", bonus_day_off_slip WHERE prm_empno = bd_empno "
    End If

    fnSetEmpComboSQL = strSQL
End Function

Public Function fnInvalidData(txtBox As Textbox) As Boolean
    #If PROTOTYPE Then
        Exit Function
    #End If

    Select Case txtBox.TabIndex
        Case txtEmployee.TabIndex
            fnInvalidData = Not fnValidEmployee(txtBox)
    End Select
    
End Function

Private Function fnValidEmployee(Box As Textbox) As Boolean
    Const SUB_NAME As String = "fnValidEmployee"
    Dim strSQL As String
    Dim rsTemp As Recordset
    Dim sMsg As String
    
    fnValidEmployee = False
    
    If Trim(Box.Text) = "" Then
        cvalidate.SetErrorMessage Box, "You Must Enter a Employee Number"
        Exit Function
    End If
    
    If Not IsNumeric(Trim(Box.Text)) Then
        cvalidate.SetErrorMessage Box, "Employee Number does not exist"
        Exit Function
    End If
    
    If t_nFormMode = ADD_MODE Then
        strSQL = "SELECT prm_empno, prm_empname FROM sTmpCustTable WHERE prm_empno = " & tfnRound(Box.Text)
    Else
        strSQL = "SELECT prm_empno, prm_empname FROM sTmpCustTable, bonus_day_off_slip  WHERE"
        strSQL = strSQL & " prm_empno = bd_empno AND prm_empno = " & Box.Text
    End If
    
    If GetRecordSet(rsTemp, strSQL, , SUB_NAME) < 0 Then
        cvalidate.SetErrorMessage Box, "Failed to access Database"
        Exit Function
    End If
            
    If rsTemp.RecordCount = 0 Then
        cvalidate.SetErrorMessage Box, "Employee Number does not exist"
        tgcDropdown.SetText txtEmployee, ""
        Exit Function
    Else
        tgcDropdown.SetText txtEmployee, rsTemp!prm_empname
    End If
    
    fnValidEmployee = True
    
End Function

Private Sub tblComboDropdown_Click()
    tgcDropdown.Click tblComboDropdown
End Sub

Private Sub tblComboDropdown_GotFocus()
    tgcDropdown.GotFocus tblComboDropdown
End Sub

Private Sub tblComboDropdown_LostFocus()
    tgcDropdown.LostFocus tblComboDropdown
End Sub

Private Sub tblComboDropdown_KeyPress(KeyAscii As Integer)
    tgcDropdown.Keypress tblComboDropdown, KeyAscii
End Sub

Private Sub tblComboDropdown_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    tgcDropdown.RowColChange
End Sub

Private Sub tblComboDropdown_SelChange(CANCEL As Integer)
    tgcDropdown.SelChange CANCEL
End Sub

Private Sub tblComboDropDown_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    tgcDropdown.TableMouseUp y
End Sub

Private Sub subSetupCombos()
    Set tgcDropdown = CreateObject(t_szOLECOMBO)
    
    #If PROTOTYPE Then
        Exit Sub
    #End If
    
    With tgcDropdown
        Set .Form = Me
        Set .DBEngine = t_engFactor
        Set .DataBase = t_dbMainDatabase
        Set .DataLink = datComboDropDown
        Set .Table = tblComboDropdown
        
        .AddCombo
        .AddComboBox txtEmployee, cmdEmployee, "prm_empno", .SQL_LONG_TYPE
        .AddComboBox txtEmpName, cmdEmpName, "prm_empname", .SQL_STRING_TYPE(60)
    End With
     
End Sub

Private Function fnSwitchPhase() As Boolean
    Dim bYesNo As Boolean
    
    If Not tfnLockRow(efraToolBar.FMName, "bonus_day_off_slip", txtEmployee) Then
        Exit Function
    End If
    
    
    If t_nFormMode = ADD_MODE Then
        bYesNo = True
    Else
        
        If Not fnLoadData Then
            Exit Function
        End If
        
        bYesNo = True
        
        nDataStatus = DATA_LOADED
    End If
    
    subEnableEmployee Not bYesNo
    fnSwitchPhase = bYesNo
    
End Function

Private Sub subEnableEmployee(bOnOff As Boolean)
    txtEmployee.Enabled = bOnOff
    cmdEmployee.Enabled = bOnOff
    txtEmpName.Enabled = bOnOff
    cmdEmpName.Enabled = bOnOff
    subEnableSearchbtn cmdEmployee, bOnOff
    subEnableSearchbtn cmdEmpName, bOnOff
    tblSlip.Enabled = Not bOnOff
End Sub

Private Sub subEnableOutFrame(bOnOff As Boolean)
    efraOuterFrame.Enabled = bOnOff
End Sub

Private Function fnCreateSearchTable(szNumber As String, szName As String) As Boolean
    Dim sDropSQL As String
    Dim strSQL As String
    
    On Error Resume Next
    sDropSQL = " DROP TABLE sTmpCustTable"
    t_dbMainDatabase.ExecuteSQL sDropSQL
    On Error GoTo 0

    On Error GoTo EXITHERE
    strSQL = " CREATE TEMP  TABLE sTmpCustTable " & " ( " & szNumber & " INTEGER," & szName & " CHAR(60)) "
    t_dbMainDatabase.ExecuteSQL strSQL
    
    strSQL = " INSERT INTO  sTmpCustTable " & " SELECT prm_empno, TRIM(prm_last_name) || ', '"
    strSQL = strSQL & " || TRIM(prm_first_name) || ' '  || TRIM(prm_middle_name) "
    strSQL = strSQL & " FROM pr_master WHERE  TRIM(prm_middle_name)<> ''"
    t_dbMainDatabase.ExecuteSQL strSQL
    
    strSQL = " INSERT INTO  sTmpCustTable " & " SELECT prm_empno, TRIM(prm_last_name) || ', '"
    strSQL = strSQL & " || TRIM(prm_first_name)  "
    strSQL = strSQL & " FROM pr_master WHERE  TRIM(prm_middle_name) = '' OR prm_middle_name IS NULL"
    t_dbMainDatabase.ExecuteSQL strSQL
    
    fnCreateSearchTable = True
    Exit Function
    
EXITHERE:
    tfnErrHandler "fnCreateSearchTable", strSQL
    Err.Clear
    fnCreateSearchTable = False
End Function

Private Function fnValidSlipDate(sDate As String, lCurRow As Long, nCurCol As Integer) As Boolean
    Dim nCode As Integer
    Dim sTemp As String
    
    fnValidSlipDate = False
    
    If sDate = "" Then
        tgmSlip.ErrorMessage(nCurCol) = "The Slip Date can't be empty"
        Exit Function
    End If
    
    nCode = SRegExpMatch(szDatePattern, sDate)
    
    If nCode = 0 Then
        sTemp = tfnDateString(tfnFormatDate(sDate))
        
        If IsDate(sTemp) Then
            fnValidSlipDate = True
            tgmSlip.CellValue(nCurCol, lCurRow) = sTemp
        End If
        
    Else
        tgmSlip.ErrorMessage(nCurCol) = "Invalid Date entered"
    End If

End Function

Private Function fnValidPrftCtr(szText As String, lRow As Long, nCol As Integer) As Boolean
    Const SUB_NAME As String = "fnValidPrftCtr"
    Dim strSQL As String
    Dim rsTemp As Recordset
    Dim sPrftName As String
    
    fnValidPrftCtr = False
    
    If Trim(szText) = "" Then
        tgmSlip.ErrorMessage(ColPrftCtr) = "You must enter a valid profit center"
        Exit Function
    End If
    
    strSQL = "SELECT * FROM sys_prft_ctr WHERE prft_ctr = " & Trim(szText)
    
    If GetRecordSet(rsTemp, strSQL, , SUB_NAME) <> 1 Then
        tgmSlip.ErrorMessage(ColPrftCtr) = "Invalid profit center specified"
        Exit Function
    Else
        tgmSlip.CellValue(colPrftDesc, lRow) = Trim(rsTemp!prft_name & "")
    End If
    
    fnValidPrftCtr = True

End Function

Private Function fnValidShiftNum(szText As String, lCurRow As Long, nCurCol As Integer) As Boolean
    
    If szText = "" Then
        tgmSlip.ErrorMessage(nCurCol) = "The shift Number can't be empty"
        Exit Function
    Else
        fnValidShiftNum = fnValidEntireRowValue(szText, lCurRow, nCurCol)
    End If
    
End Function

Private Function fnValidEntireRowValue(szText As String, lCurRow As Long, nCurCol As Integer) As Boolean
    Dim strSQL As String
    Dim rsTemp As Recordset
    Dim i As Long
    
    If tgmSlip.ValidCell(ColPrftCtr, lCurRow) And tgmSlip.ValidCell(ColSlipDate, lCurRow) Then
        strSQL = "SELECT * FROM bonus_day_off_slip WHERE "
        strSQL = strSQL & " bd_empno = " & tfnRound(txtEmployee)
        strSQL = strSQL & " AND bd_prft_ctr = " & tgmSlip.CellValue(ColPrftCtr, lCurRow)
        strSQL = strSQL & " AND bd_slip_date = " & tfnDateString(tgmSlip.CellValue(ColSlipDate, lCurRow), True)
        strSQL = strSQL & " AND bd_shift = " & szText
    
        If GetRecordSet(rsTemp, strSQL) < 0 Then
            tgmSlip.ErrorMessage(nCurCol) = "Failed to access Database"
            Exit Function
        ElseIf rsTemp.RecordCount > 0 Then
            tgmSlip.ErrorMessage(nCurCol) = "The record is already in the system"
            Exit Function
        End If
        
        For i = 0 To tgmSlip.RowCount - 1
            
            If i <> lCurRow Then
                
                If CStr(tgmSlip.CellValue(ColPrftCtr, i)) = CStr(tgmSlip.CellValue(ColPrftCtr, lCurRow)) _
                    And CStr(tgmSlip.CellValue(ColSlipDate, i)) = CStr(tgmSlip.CellValue(ColSlipDate, lCurRow)) _
                    And CStr(tgmSlip.CellValue(colShiftNum, i)) = CStr(szText) Then
                    
                    tgmSlip.ErrorMessage(nCurCol) = "Duplicate data in the grid"
                    Exit Function
                End If
            
            End If
            
        Next
        
    End If
    
    fnValidEntireRowValue = True
End Function
Private Sub subSetupDropDown()
    
    subEnableSearchbtn cmdSearch, True
    
    Set tgfDropdown = New clsFloatingDropDown
    
    With tgfDropdown
        Set .SearchButton = cmdSearch
        Set .MainTable = tblSlip
        Set .DropDownTable = tblDropDown
        Set .DataBase = t_dbMainDatabase
        Set .DataLink = datDropDown
        Set .Form = Me
        Set .EditClass = tgmSlip
        .DefaultCursorControl = True
    
        .AddDropDown 1
        .AddColumn ColPrftCtr, "prft_ctr"
        .ColumnCaption(ColPrftCtr) = "Profit Center"
        .ColumnType(ColPrftCtr) = .COLUMN_TYPE_INTEGER
        .AddColumn colPrftDesc, "prft_name"
        .ColumnCaption(colPrftDesc) = "Profit Center description"
        .ColumnType(colPrftDesc) = .COLUMN_TYPE_STRING
        .DropDownSQL(1) = "SELECT prft_ctr, prft_name FROM sys_prft_ctr"
    End With
    
End Sub


Public Function fnValidCellValue(tgTable As TDBGrid, _
                                 ByVal nCurrCol As Integer, _
                                 ByVal lCurrRow As Long, _
                                 szText As String) As Boolean
                                 
    Dim bValid As Boolean
    
    Select Case nCurrCol
        Case ColPrftCtr
            bValid = fnValidPrftCtr(szText, lCurrRow, nCurrCol)
            
            If bValid Then
                    tgmSlip.Revalidate colShiftNum, lCurrRow
            End If
            
        Case colPrftDesc
            bValid = True
        Case ColSlipDate
            bValid = fnValidSlipDate(szText, lCurrRow, nCurrCol)
                
            If bValid Then
                tgmSlip.Revalidate colShiftNum, lCurrRow
            End If
            
        Case colShiftNum
            bValid = fnValidShiftNum(szText, lCurrRow, nCurrCol)
    End Select
    
    fnValidCellValue = bValid
End Function

Private Sub txtEmpName_LostFocus()
    
    If cvalidate.LostFocus(txtEmployee, cmdEmployee, txtEmpName, cmdEmpName, tblComboDropdown) Then
        tgcDropdown.LostFocus txtEmpName
        
        If cvalidate.ValidInput(txtEmployee) And Not tblSlip.Enabled Then
            
            If fnSwitchPhase() Then
                DoEvents
                subSetFocus tblSlip
            End If
            
        End If
        
    End If
    
End Sub

Private Function fnGetShiftLink(lProfitCenter As Long, lShift As Long, dDate As Date) As Long
    Dim strSQL As String
    Dim rsShiftLink As Recordset
    Dim strInsSQL As String
    
    strSQL = "SELECT rssl_shl FROM rs_shiftlink WHERE rssl_prft_ctr = " & CStr(lProfitCenter) & _
                " AND rssl_shift = " & CStr(lShift) & " AND rssl_date = '" & CStr(Format(dDate, "mm/dd/yyyy")) & "'"
                
    
    If GetRecordSet(rsShiftLink, strSQL) < 0 Then
        Exit Function
    ElseIf rsShiftLink.RecordCount > 0 Then
        fnGetShiftLink = rsShiftLink!rssl_shl
    Else
    
        strInsSQL = "INSERT INTO rs_shiftlink VALUES(" & CStr(lProfitCenter) & ", " & CStr(lShift) & ", '" & CStr(Format(dDate, "mm/dd/yyyy")) & "', 0)"
        
        If fnExecuteSQL(strInsSQL) Then
            
            If GetRecordSet(rsShiftLink, strSQL) > 0 Then
                fnGetShiftLink = rsShiftLink!rssl_shl
            End If
            
        End If
        
    End If
    
End Function

Private Sub subCheckDataStatus()

    If t_nFormMode = ADD_MODE Then
    
        If tgmSlip.ValidData Then
            subEnableUpdateBtn True
        End If
    
    End If
    
End Sub

Private Function fnInsertData() As Boolean
    Dim strSQL As String
    Dim vData() As Variant
    Dim i As Long
    Dim nCount As Integer
    
    If t_nFormMode = ADD_MODE Then
        
        For i = 0 To tgmSlip.RowCount - 1
            tgmSlip.GetRow vData, nCount, i
            
            If fnGetShiftLink(CLng(vData(ColPrftCtr)), CLng(vData(colShiftNum)), CDate(vData(ColSlipDate))) = 0 Then
                Exit Function
            End If
            
            strSQL = "INSERT INTO bonus_day_off_slip(bd_empno, bd_slip_date, bd_prft_ctr, bd_shift) VALUES("
            strSQL = strSQL & tfnRound(txtEmployee) & "," & tfnDateTimeString(vData(ColSlipDate), True) & ","
            strSQL = strSQL & tfnRound(vData(ColPrftCtr)) & "," & tfnRound(vData(colShiftNum)) & ")"
            
            If Not fnExecuteSQL(strSQL) Then
                Exit Function
            End If
            
        Next
        
    End If
    
    fnInsertData = True
    
End Function

Private Sub subSetupSlipGrid()

    With tgmSlip
    
        If t_nFormMode = ADD_MODE Then
            .AddEditColumn ColPrftCtr, "Enter a Profit center", szIntegerPattern
            .AddESCColumn colPrftDesc, ColPrftCtr
            .AddEditColumn colPrftDesc, "Enter Profit center description ", "^P{1,40}$"
            .AddEditColumn ColSlipDate, "Enter Slip Date", szDatePattern
            .AddEditColumn colShiftNum, "Enter Shift Number", szIntegerPattern
            .AllowAddNew = True
        Else
            .RemoveEditColumn ColPrftCtr
            .RemoveEditColumn colPrftDesc
            .RemoveEditColumn ColSlipDate
            .RemoveEditColumn colShiftNum
            .AllowAddNew = False
        End If
        
    End With
    
End Sub
