VERSION 5.00
Object = "{C75015E0-2232-11D3-B440-0060971E99AF}#1.0#0"; "FACTFRM.OCX"
Object = "{01028C21-0000-0000-0000-000000000046}#4.0#0"; "TG32OV.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmZZFITEMT 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Item Cost Maintenance"
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
      TabIndex        =   18
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
      TabIndex        =   12
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
         TabIndex        =   17
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
         TabIndex        =   14
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
         TabIndex        =   15
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
         TabIndex        =   16
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
         TabIndex        =   20
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
         Begin DBTrueGrid.TDBGrid tblDetail 
            Height          =   2880
            HelpContextID   =   708
            Left            =   84
            OleObjectBlob   =   "ZZFITEMT.frx":0000
            TabIndex        =   13
            Top             =   1644
            Width           =   8520
         End
         Begin FACTFRMLib.FactorFrame efraHeader 
            Height          =   1512
            Left            =   84
            TabIndex        =   11
            Top             =   120
            Width           =   8520
            _Version        =   65536
            _ExtentX        =   15028
            _ExtentY        =   2667
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
            Begin VB.TextBox txtItemDesc 
               Height          =   360
               HelpContextID   =   706
               Left            =   2616
               TabIndex        =   9
               Top             =   1056
               Width           =   5448
            End
            Begin VB.TextBox txtVendorDesc 
               BackColor       =   &H00FFFFFF&
               DataSource      =   "datVendor"
               ForeColor       =   &H00000000&
               Height          =   360
               HelpContextID   =   702
               Left            =   2616
               TabIndex        =   5
               Tag             =   "pn_name"
               Top             =   372
               Width           =   5448
            End
            Begin FACTFRMLib.FactorFrame cmdItemDesc 
               Height          =   360
               HelpContextID   =   707
               Left            =   8064
               TabIndex        =   10
               TabStop         =   0   'False
               Top             =   1056
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
            Begin FACTFRMLib.FactorFrame cmdItemCode 
               Height          =   360
               HelpContextID   =   705
               Left            =   2196
               TabIndex        =   8
               TabStop         =   0   'False
               Top             =   1056
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
            Begin VB.TextBox txtItemCode 
               Height          =   360
               HelpContextID   =   704
               Left            =   84
               TabIndex        =   7
               Top             =   1056
               Width           =   2100
            End
            Begin VB.TextBox txtVendor 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00000000&
               Height          =   360
               HelpContextID   =   700
               Left            =   84
               TabIndex        =   3
               Tag             =   "pn_alt"
               Top             =   372
               Width           =   2100
            End
            Begin FACTFRMLib.FactorFrame cmdVendorDesc 
               Height          =   360
               HelpContextID   =   703
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
               Picture         =   "ZZFITEMT.frx":12F1
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
            Begin FACTFRMLib.FactorFrame cmdVendor 
               Height          =   360
               HelpContextID   =   701
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
               Picture         =   "ZZFITEMT.frx":1403
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
            Begin VB.Label lblItemDesc 
               Caption         =   "Item Description"
               Height          =   228
               Left            =   2616
               TabIndex        =   25
               Top             =   756
               Width           =   1488
            End
            Begin VB.Label lblItemCode 
               Caption         =   "Item No."
               Height          =   264
               Left            =   96
               TabIndex        =   24
               Top             =   756
               Width           =   756
            End
            Begin VB.Label lblVendorName 
               Caption         =   "Vendor Description"
               Height          =   252
               Left            =   2616
               TabIndex        =   23
               Top             =   108
               Width           =   1956
            End
            Begin VB.Label lblVendorNo 
               Caption         =   "Vendor No."
               Height          =   252
               Left            =   96
               TabIndex        =   22
               Top             =   108
               Width           =   1632
            End
         End
      End
   End
   Begin FACTFRMLib.FactorFrame efraToolBar 
      Height          =   468
      Left            =   0
      TabIndex        =   19
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
      FMName          =   "ZZFITEMT"
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
      Bindings        =   "ZZFITEMT.frx":1515
      Height          =   2484
      Left            =   36
      OleObjectBlob   =   "ZZFITEMT.frx":1534
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   660
      Width           =   3756
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
Attribute VB_Name = "frmZZFITEMT"
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
Private Const szHelp7_ELEVEN As String = "7-Eleven.HLP"
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
Private tgmDetail As clsTGSpreadSheet
Private cValidate As cValidateInput

Private Const ColUOM As Integer = 0
Private Const ColEffDate As Integer = 1
Private Const ColCost As Integer = 2
Private Const ColRetail As Integer = 3
Private Const nMaxCol As Integer = 3


Private Sub cmdDeleteBtn_Click()
    Dim sBCode As String
    Dim lRow As Long
    Dim strSQL As String
    
    If tgmDetail.RowCount < 1 Then
        subEnableDeleteBtn False
        Exit Sub
    End If
    
    lRow = tgmDetail.GetCurrentRowNumber
    
    If t_nFormMode = EDIT_MODE Then
        
        If Not tfnCancelExit("Are you sure you want to delete the current record?") Then
            Exit Sub
        End If
            
            
        strSQL = "DELETE FROM item_cost_maint "
        strSQL = strSQL & " WHERE icm_vendor = " & tfnRound(txtVendor)
        strSQL = strSQL & " AND icm_code = " & tfnRound(txtItemCode)
        strSQL = strSQL & " AND icm_eff_date = " & tfnDateString(tgmDetail.CellValue(ColEffDate, lRow), True)
        
        If Not fnExecuteSQL(strSQL) Then
            Exit Sub
        End If
        
        tgmDetail.DeleteRow
    End If
    
    If tgmDetail.RowCount < 1 Then
        subEnableDeleteBtn False
        
        If t_nFormMode = EDIT_MODE Then
            cmdExitCancelBtn_Click
        End If
        
    End If
    
End Sub

Private Sub cmdItemCode_Click()
    subShowHourGlass
    tgcDropdown.ComboSQL(txtItemCode) = fnSetItemCodeComboSQL
    tgcDropdown.Click cmdItemCode
    subHideHourGlass

End Sub

Private Sub cmdItemDesc_Click()
    subShowHourGlass
    tgcDropdown.ComboSQL(txtItemDesc) = fnSetItemCodeComboSQL
    tgcDropdown.Click cmdItemDesc
    subHideHourGlass
End Sub

Private Sub efraHeader_Gotfocus()
    cValidate.GotFocus efraBackground
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
    App.HelpFile = szHelp7_ELEVEN
End Sub

Private Sub Form_Unload(CANCEL As Integer)
    
    On Error Resume Next
    
    Set objErrHandler = Nothing
    
    Set objCurrTabControl = Nothing
    
    Unload frmContext
    Unload frmAbout
    
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
        
    #End If
    
    ' disable system menu Close and Close icon on form
    tfnDisableFormSystemClose Me
    subSetupToolBar
    
    tmrKeyBoard.Enabled = False
    tfnCenterForm Me
    
    subSetupCombos
    subInitValidation
    subInitializeSpreadSheet
    
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
    
    With tgmDetail
        .AddEditColumn ColUOM, "Enter Unit of Measure", "^P{1,5}$"
        .AddEditColumn ColEffDate, "Enter effective Date", szDatePattern
        .AllowAddNew = True
    End With
    
    subEnableOutFrame True
    subEnableHeader True
    subEnableAddBtn False
    subEnableEditBtn False
    subEnableUpdateBtn False
    subEnableRefreshBtn False
    subEnableCancel True
    subSetFocus txtVendor
    
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
    
    With tgmDetail
        .RemoveEditColumn ColUOM
        .RemoveEditColumn ColEffDate
        .AllowAddNew = False
    End With
    
    subEnableOutFrame True
    subEnableHeader True
    subEnableAddBtn False
    subEnableEditBtn False
    subEnableUpdateBtn False
    subEnableRefreshBtn False
    subEnableCancel True
    subSetFocus txtVendor
    
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
    subEnableOutFrame False
    subEnableHeader False
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
    
    txtVendor = ""
    txtVendorDesc = ""
    txtItemCode = ""
    txtItemDesc = ""
    cValidate.ResetFlags
    tgmDetail.ClearData
    nDataStatus = DATA_INIT
    subSetFocus cmdAddBtn
    subHideHourGlass
    DoEvents
    
End Sub

Private Sub tblDetail_AfterColEdit(ByVal ColIndex As Integer)
    tgmDetail.AfterColEdit ColIndex
    nDataStatus = DATA_CHANGED
    subCheckDataStatus
End Sub

Private Sub tblDetail_BeforeColEdit(ByVal ColIndex As Integer, ByVal KeyAscii As Integer, CANCEL As Integer)
    tgmDetail.BeforeColEdit ColIndex, KeyAscii, CANCEL
    subEnableUpdateBtn False
End Sub

Private Sub tblDetail_Change()
    tgmDetail.Change
End Sub


Private Sub tblDetail_GotFocus()
    tgmDetail.GotFocus
End Sub

Private Sub tblDetail_KeyDown(KeyCode As Integer, Shift As Integer)
    tgmDetail.KeyDown KeyCode, Shift
End Sub

Private Sub tblDetail_KeyPress(KeyAscii As Integer)
    
    If Not tgmDetail.Keypress(KeyAscii) Then
        KeyAscii = 0
    End If
    
    If tblDetail.col = ColUOM Then
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If
    
End Sub

Private Sub tblDetail_LostFocus()
    tgmDetail.LostFocus
End Sub

Private Sub tblDetail_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    tgmDetail.RowColChange LastRow, LastCol
End Sub

Private Sub tblDetail_SelChange(CANCEL As Integer)
    CANCEL = True
End Sub

Private Sub tblDetail_UnboundReadData(ByVal RowBuf As DBTrueGrid.RowBuffer, StartLocation As Variant, ByVal ReadPriorRows As Boolean)
    tgmDetail.ReadData RowBuf, StartLocation, ReadPriorRows
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
    Dim sDecimal10Pattern As String
    
    myWidth = Array(0.25, 0.25, 0.25, 0.25)
    myField = Array("icm_uom", "icm_eff_date", "icm_cost", "icm_retail")
    
    While tblDetail.Columns.Count > 0
        tblDetail.Columns.Remove 0
    Wend
    
    For i = 0 To UBound(myWidth)
        tblDetail.Columns.Add i
        
        With tblDetail.Columns(i)
            .Width = myWidth(i) * tblDetail.Width - 50
            .DataField = myField(i)
            .Visible = True
        End With
        
    Next
    
    tblDetail.Columns(ColUOM).Caption = "Unit Of Measure"
    tblDetail.Columns(ColEffDate).Caption = "Effective Date"
    tblDetail.Columns(ColCost).Caption = "Cost"
    tblDetail.Columns(ColRetail).Caption = "Retail"

    Set tgmDetail = New clsTGSpreadSheet
    
    With tgmDetail
        Set .Form = Me
        Set .StatusBar = ffraStatusbar
        Set .Table = tblDetail
        
        sDecimal10Pattern = tfnDecimalPattern(10, 3, False, True)
        .AddEditColumn ColUOM, "Enter Unit of Measure", "^P{1,5}$"
        .AddEditColumn ColEffDate, "Enter Effective Date ", szDatePattern
        .AddEditColumn ColCost, "Enter Cost", sDecimal10Pattern
        .AddEditColumn ColRetail, "Enter Retail Price", sDecimal10Pattern
        '.AllowAddNew = True
    End With
    
End Sub

Private Function fnLoadData() As Boolean
    Const SUB_NAME As String = "fnLoadData"
    Dim strSQL As String
    Dim rsTemp As Recordset
    
    fnLoadData = False
    
    strSQL = "SELECT icm_uom, icm_eff_date, icm_cost, icm_retail"
    strSQL = strSQL & " FROM item_cost_maint "
    strSQL = strSQL & " WHERE icm_vendor = " & tfnRound(txtVendor)
    strSQL = strSQL & " AND icm_code = " & tfnSQLString(txtItemCode)
    strSQL = strSQL & " ORDER BY icm_uom, icm_eff_date, icm_cost"
        
    If GetRecordSet(rsTemp, strSQL, , SUB_NAME) < 0 Then
        MsgBox "Failed to access database to get records", vbCritical
        Exit Function
    End If
    
    If rsTemp.RecordCount = 0 Then
        MsgBox "No detail available for this vendor and item", vbExclamation
        Exit Function
    End If
    
    tgmDetail.FillWithRecordset rsTemp, True
    subEnableDeleteBtn True
    
    fnLoadData = True
        
End Function

Private Sub txtItemCode_Change()
    cValidate.Change txtItemCode
    tgcDropdown.Change txtItemCode
End Sub

Private Sub txtItemCode_GotFocus()
    tgcDropdown.GotFocus txtItemCode
    cValidate.GotFocus txtItemCode

    If tgcDropdown.SingleRecordSelected Then
        
        If fnSwitchPhase() Then
            subSetFocus tblDetail
        End If
        
    End If
    
End Sub

Private Sub txtItemCode_KeyPress(KeyAscii As Integer)
    Dim bKeyCode As Boolean
    
    If KeyAscii = vbKeyReturn Then
        tgcDropdown.ComboSQL(txtItemCode) = fnSetItemCodeComboSQL
        subShowHourGlass
    End If
    
    bKeyCode = tgcDropdown.Keypress(txtItemCode, KeyAscii)
    
    If Not bKeyCode Then
    
        If KeyAscii = vbKeyReturn Then
          
            If tgcDropdown.SingleRecordSelected Then
                
                If fnSwitchPhase() Then
                    subSetFocus tblDetail
                End If
                
            End If
          
          KeyAscii = 0
          
       End If
       
    Else
        cValidate.Keypress txtItemCode, KeyAscii
    End If
    
    subHideHourGlass
End Sub

Private Sub txtItemCode_LostFocus()
    tgcDropdown.LostFocus txtItemCode
    cValidate.LostFocus txtItemCode, cmdItemCode, txtItemDesc, cmdItemDesc, tblComboDropdown
End Sub

Private Sub txtItemDesc_Change()
    tgcDropdown.Change txtItemDesc
End Sub

Private Sub txtItemDesc_GotFocus()
    tfnSetStatusBarMessage "Enter Item Description"
    tgcDropdown.GotFocus txtItemDesc

    If tgcDropdown.SingleRecordSelected Then
        
        If fnSwitchPhase() Then
            subSetFocus tblDetail
        End If
    
    End If
    
End Sub

Private Sub txtItemDesc_KeyPress(KeyAscii As Integer)
    Dim bKeyCode As Boolean
    
    If KeyAscii = vbKeyReturn Then
        tgcDropdown.ComboSQL(txtItemDesc) = fnSetItemCodeComboSQL
        subShowHourGlass
    End If
    
    bKeyCode = tgcDropdown.Keypress(txtItemDesc, KeyAscii)
    
    If Not bKeyCode Then
        
        If KeyAscii = vbKeyReturn Then
            
            If tgcDropdown.SingleRecordSelected Then
                
                If fnSwitchPhase() Then
                    subSetFocus tblDetail
                End If
            
            End If
            
            KeyAscii = 0
        End If
        
    End If
    
    subHideHourGlass
End Sub

Private Sub txtItemDesc_LostFocus()
    
    If cValidate.LostFocus(txtItemCode, cmdItemCode, txtItemDesc, cmdItemDesc, tblComboDropdown) Then
        tgcDropdown.LostFocus txtItemDesc
        
        If cValidate.FirstInvalidInput < 0 And Not tblDetail.Enabled Then
            
            If fnSwitchPhase() Then
                DoEvents
                subSetFocus tblDetail
            End If
            
        End If
        
    End If
    
End Sub

Private Sub txtVendor_Change()
    cValidate.Change txtVendor
    tgcDropdown.Change txtVendor
End Sub

Private Sub txtVendor_GotFocus()
    
    tgcDropdown.GotFocus txtVendor
    cValidate.GotFocus txtVendor
    
    If tgcDropdown.SingleRecordSelected Then
        subSetFocus txtItemCode
    End If
    
End Sub

Private Sub txtVendor_KeyPress(KeyAscii As Integer)
    Dim bKeyCode As Boolean
    
    If KeyAscii = vbKeyReturn Then
        tgcDropdown.ComboSQL(txtVendor) = fnSetVendorComboSQL
        subShowHourGlass
    End If
    
    bKeyCode = tgcDropdown.Keypress(txtVendor, KeyAscii)
    
    If Not bKeyCode Then
    
        If KeyAscii = vbKeyReturn Then
          
            If tgcDropdown.SingleRecordSelected Then
                subSetFocus txtItemCode
            End If
          
          KeyAscii = 0
          
       End If
       
    Else
        cValidate.Keypress txtVendor, KeyAscii
    End If
    
    subHideHourGlass
End Sub

Private Sub txtVendor_LostFocus()
    tgcDropdown.LostFocus txtVendor
    cValidate.LostFocus txtVendor, cmdVendor, txtVendorDesc, cmdVendorDesc, tblComboDropdown
End Sub

Private Sub txtVendorDesc_Change()
    tgcDropdown.Change txtVendorDesc
End Sub

Private Sub txtVendorDesc_GotFocus()
    tfnSetStatusBarMessage "Enter Vendor Description"
    tgcDropdown.GotFocus txtVendorDesc

    If tgcDropdown.SingleRecordSelected Then
        subSetFocus txtItemCode
    End If
    
End Sub

Private Sub txtVendorDesc_KeyPress(KeyAscii As Integer)
    Dim bKeyCode As Boolean
    
    If KeyAscii = vbKeyReturn Then
        tgcDropdown.ComboSQL(txtVendorDesc) = fnSetVendorComboSQL
        subShowHourGlass
    End If
    
    bKeyCode = tgcDropdown.Keypress(txtVendorDesc, KeyAscii)
    
    If Not bKeyCode Then
        
        If KeyAscii = vbKeyReturn Then
            
            If tgcDropdown.SingleRecordSelected Then
                subSetFocus txtItemCode
            End If
            
            KeyAscii = 0
        End If
        
    End If
    
    subHideHourGlass
End Sub

Private Sub cmdvendor_Click()
    subShowHourGlass
    tgcDropdown.ComboSQL(txtVendor) = fnSetVendorComboSQL
    tgcDropdown.Click cmdVendor
    subHideHourGlass
End Sub

Private Sub cmdvendordesc_Click()
    subShowHourGlass
    tgcDropdown.ComboSQL(txtVendorDesc) = fnSetVendorComboSQL
    tgcDropdown.Click cmdVendorDesc
    subHideHourGlass
End Sub

Private Sub subInitValidation()
    Set cValidate = New cValidateInput
    Set cValidate.Form = Me
    Set cValidate.StatusBar = ffraStatusbar
    
    cValidate.AddEditBox txtVendor, "Enter Vendor Number"
    cValidate.AddEditBox txtItemCode, "Enter Item Code"
    cValidate.MinTabIndex = tbToolbar.TabIndex - 1
    cValidate.MaxTabIndex = efraHeader.TabIndex + 1
    Set cValidate.ControlForFocus = efraHeader
    Set cValidate.LastBox = txtItemCode
    cValidate.ESCControl = tblDetail
    cValidate.ESCControl = cmdExitCancelBtn
End Sub

Private Function fnSetVendorComboSQL() As String
    Dim strSQL As String
    
    If t_nFormMode = ADD_MODE Then
        strSQL = "SELECT DISTINCT pn_alt, pn_name FROM p_altName "
        
        If cValidate.ValidInput(txtItemCode) Then
            strSQL = strSQL & ", rs_b_item WHERE pn_alt = rsbi_vendor "
            strSQL = strSQL & " AND rsbi_code = " & tfnSQLString(txtItemCode)
        End If
        
    ElseIf t_nFormMode = EDIT_MODE Then
        strSQL = "SELECT DISTINCT pn_alt, pn_name FROM item_cost_maint, p_altname WHERE icm_vendor = pn_alt"
        
        If cValidate.ValidInput(txtItemCode) Then
            strSQL = strSQL & " AND icm_code = " & tfnSQLString(txtItemCode)
        End If
        
    End If

    fnSetVendorComboSQL = strSQL
End Function

Private Function fnSetItemCodeComboSQL() As String
    Dim strSQL As String
    
    If t_nFormMode = ADD_MODE Then
        strSQL = "SELECT Distinct rsbi_code, rsbi_desc FROM rs_b_item"
        
        If cValidate.ValidInput(txtVendor) Then
            strSQL = strSQL & " WHERE rsbi_vendor = " & txtVendor
        End If
    
    ElseIf t_nFormMode = EDIT_MODE Then
            
        strSQL = "SELECT DISTINCT rsbi_code, rsbi_desc FROM rs_b_item, item_cost_maint WHERE rsbi_code = icm_code "
        strSQL = strSQL & " AND icm_vendor = rsbi_vendor "
        
        If cValidate.ValidInput(txtVendor) Then
            strSQL = strSQL & " And icm_vendor = " & txtVendor
        End If
    
    End If
    
    fnSetItemCodeComboSQL = strSQL
    
End Function

Public Function fnInvalidData(txtBox As Textbox) As Boolean
    #If PROTOTYPE Then
        Exit Function
    #End If

    Select Case txtBox.TabIndex
        Case txtVendor.TabIndex
            fnInvalidData = Not fnValidVendor(txtBox)
        Case txtItemCode.TabIndex
            fnInvalidData = Not fnValidItemCode(txtBox)
    End Select
    
End Function

Private Function fnValidVendor(Box As Textbox) As Boolean
    Const SUB_NAME As String = "fnValidVendor"
    Dim strSQL As String
    Dim rsTemp As Recordset
    Dim sMsg As String
    
    fnValidVendor = False
    
    If Trim(Box.Text) = "" Then
        cValidate.SetErrorMessage Box, "You Must Enter a Vendor Number"
        Exit Function
    End If
    
    If Not IsNumeric(Trim(Box.Text)) Then
        cValidate.SetErrorMessage Box, "Vendor Number Number does not exist"
        Exit Function
    End If
    
    If t_nFormMode = ADD_MODE Then
        strSQL = "SELECT pn_alt, pn_name FROM p_altName WHERE pn_alt = " & tfnRound(Box.Text)
    Else
        strSQL = "SELECT pn_alt, pn_name FROM p_altname, item_cost_maint WHERE icm_vendor = pn_alt"
        strSQL = strSQL & " AND icm_vendor = " & tfnSQLString(txtVendor)
    End If
    
    If GetRecordSet(rsTemp, strSQL, , SUB_NAME) < 0 Then
        cValidate.SetErrorMessage Box, "Failed to access Database"
        Exit Function
    End If
            
    If rsTemp.RecordCount = 0 Then
        cValidate.SetErrorMessage Box, "Vendor Number does not exist"
        tgcDropdown.SetText txtVendor, ""
        Exit Function
    Else
        tgcDropdown.SetText txtVendor, rsTemp!pn_name
    End If
    
    fnValidVendor = True
    
End Function

Private Function fnValidItemCode(Box As Textbox) As Boolean
    Const SUB_NAME As String = "fnValidItemCode"
    Dim strSQL As String
    Dim rsTemp As Recordset
    Dim sMsg As String
    
    fnValidItemCode = False
    
    If Trim(Box.Text) = "" Then
        cValidate.SetErrorMessage Box, "You Must Enter a item Code"
        Exit Function
    End If
    
    If t_nFormMode = ADD_MODE Then
        strSQL = "SELECT rsbi_code, rsbi_desc FROM rs_b_item WHERE  rsbi_code = " & tfnSQLString(Box.Text)
        
        If cValidate.ValidInput(txtVendor) Then
            strSQL = strSQL & " AND rsbi_vendor = " & txtVendor
        End If
        
    Else
    
        strSQL = "SELECT rsbi_code, rsbi_desc FROM rs_b_item, item_cost_maint WHERE icm_vendor = rsbi_vendor"
        strSQL = strSQL & " AND rsbi_code = icm_code  AND rsbi_code = " & tfnSQLString(txtItemCode)
        
        If cValidate.ValidInput(txtVendor) Then
            strSQL = strSQL & " AND icm_vendor = " & tfnSQLString(txtVendor)
        End If
        
    End If
    
    If GetRecordSet(rsTemp, strSQL, , SUB_NAME) < 0 Then
        cValidate.SetErrorMessage Box, "Failed to access Database"
        Exit Function
    End If
            
    If rsTemp.RecordCount = 0 Then
        cValidate.SetErrorMessage Box, "Item Code does not exist"
        tgcDropdown.SetText txtItemCode, ""
        Exit Function
    Else
        tgcDropdown.SetText txtItemCode, rsTemp!rsbi_desc
    End If
    
    fnValidItemCode = True
    

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
        Set .Database = t_dbMainDatabase
        Set .DataLink = datComboDropDown
        Set .Table = tblComboDropdown
        
        .addcombo
        .addComboBox txtVendor, cmdVendor, "pn_alt", .SQL_LONG_TYPE
        .addComboBox txtVendorDesc, cmdVendorDesc, "pn_name", .SQL_STRING_TYPE(30)
        
        .addcombo
        .addComboBox txtItemCode, cmdItemCode, "rsbi_code", .SQL_STRING_TYPE(10)
        .addComboBox txtItemDesc, cmdItemDesc, "rsbi_desc", .SQL_STRING_TYPE(18)
    End With
     
End Sub

Private Function fnSwitchPhase() As Boolean
    Dim bYesNo As Boolean
    
    If Not tfnLockRow(efraToolBar.FMName, "item_cost_maint", txtVendor & "," & txtItemCode) Then
        Exit Function
    End If
    
    If Not cValidate.ValidInput(txtVendor) Then
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
    
    subEnableHeader Not bYesNo
    fnSwitchPhase = bYesNo
    
End Function

Private Sub subEnableHeader(bOnOff As Boolean)
    efraHeader.Enabled = bOnOff
    subEnableSearchbtn cmdVendor, bOnOff
    subEnableSearchbtn cmdVendorDesc, bOnOff
    subEnableSearchbtn cmdItemCode, bOnOff
    subEnableSearchbtn cmdItemDesc, bOnOff
    tblDetail.Enabled = Not bOnOff
End Sub

Private Sub subEnableOutFrame(bOnOff As Boolean)
    efraOuterFrame.Enabled = bOnOff
End Sub

Private Function fnValidEffectiveDate(sDate As String, lCurRow As Long, nCurCol As Integer) As Boolean
    Dim nCode As Integer
    Dim sTemp As String
    Dim strSQL As String
    Dim rsTemp As Recordset
    Dim i As Long
    
    fnValidEffectiveDate = False
    
    If sDate = "" Then
        tgmDetail.ErrorMessage(nCurCol) = "The effective Date can't be empty"
        Exit Function
    End If
    
    nCode = SRegExpMatch(szDatePattern, sDate)
    
    If nCode = 0 Then
        sTemp = tfnDateString(tfnFormatDate(sDate))
        tgmDetail.CellValue(nCurCol, lCurRow) = sTemp
        
        If IsDate(sTemp) Then
        
            If tgmDetail.ValidCell(ColUOM, lCurRow) Then
                strSQL = "SELECT * FROM item_cost_maint WHERE icm_vendor = " & tfnRound(txtVendor)
                strSQL = strSQL & " AND icm_code = " & tfnSQLString(txtItemCode)
                strSQL = strSQL & " AND icm_eff_date = " & tfnDateString(sDate, True)
                strSQL = strSQL & " AND icm_uom = " & tfnSQLString(tgmDetail.CellValue(ColUOM, lCurRow))
                
                If GetRecordSet(rsTemp, strSQL, nDB_REMOTE, "fnValidEffectiveDate") > 0 Then
                    tgmDetail.ErrorMessage(nCurCol) = "The record is already in the system"
                    Exit Function
                End If
                
            End If
        Else
            tgmDetail.ErrorMessage(nCurCol) = "Invalid Date entered"
            Exit Function
        End If
        
        For i = 0 To tgmDetail.RowCount - 1
        
            If i <> lCurRow Then
                
                If CStr(tgmDetail.CellValue(ColEffDate, i)) = sTemp _
                     And CStr(tgmDetail.CellValue(ColUOM, i)) = CStr(tgmDetail.CellValue(ColUOM, lCurRow)) Then
                     
                    tgmDetail.ErrorMessage(nCurCol) = "The record is duplicate in the grid"
                    Exit Function
                End If
        
            End If
        
        Next

        fnValidEffectiveDate = True
    Else
        tgmDetail.ErrorMessage(nCurCol) = "Invalid Date entered"
    End If
    
End Function

Private Function fnValidUOM(szText As String, lCurRow As Long, nCurCol As Integer) As Boolean

    If Trim(szText) = "" Then
        tgmDetail.ErrorMessage(nCurCol) = "The Unit of Measure can't be empty."
        Exit Function
    End If
    
    fnValidUOM = True
    
End Function

Private Function fnValidCost(szText As String, lCurRow As Long, nCurCol As Integer) As Boolean

    If Trim(szText) = "" Then
        tgmDetail.ErrorMessage(nCurCol) = "The Cost or retail can't be empty"
        Exit Function
    ElseIf Not IsNumeric(szText) Then
        tgmDetail.ErrorMessage(nCurCol) = "Invalid Cost entered"
        Exit Function
    ElseIf CDbl(szText) < 0 Then
        tgmDetail.ErrorMessage(nCurCol) = "The cost or retail can't be less than 0"
        Exit Function
    End If
    
    fnValidCost = True
    
End Function

Public Function fnValidCellValue(tgTable As TDBGrid, _
                                 ByVal nCurrCol As Integer, _
                                 ByVal lCurrRow As Long, _
                                 szText As String) As Boolean
                                 
    Dim bValid As Boolean
    
    Select Case nCurrCol
        Case ColUOM
            bValid = fnValidUOM(szText, lCurrRow, nCurrCol)
            
            If bValid And tgmDetail.CellValue(ColEffDate, lCurrRow) <> "" Then
                tgmDetail.Revalidate ColEffDate, lCurrRow
            End If
            
        Case ColEffDate
            bValid = fnValidEffectiveDate(szText, lCurrRow, nCurrCol)
        Case ColCost
            bValid = fnValidCost(szText, lCurrRow, nCurrCol)
        Case ColRetail
            bValid = fnValidCost(szText, lCurrRow, nCurrCol)
    End Select
    
    fnValidCellValue = bValid
    
End Function

Private Sub txtVendorDesc_LostFocus()
    
    If cValidate.LostFocus(txtVendor, cmdVendor, txtVendorDesc, cmdVendorDesc, tblComboDropdown) Then
        tgcDropdown.LostFocus txtVendorDesc
    End If
    
End Sub

Private Sub subCheckDataStatus()

    If t_nFormMode = ADD_MODE Then
    
        If tgmDetail.ValidData Then
            subEnableUpdateBtn True
        End If
    
    ElseIf t_nFormMode = EDIT_MODE Then
        
        If cValidate.FirstInvalidInput < 0 And tgmDetail.ValidData Then
            subEnableUpdateBtn True
        End If
        
        If nDataStatus = DATA_CHANGED Then
            subEnableRefreshBtn True
        Else
            subEnableRefreshBtn False
        End If
        
    End If
    
End Sub

Private Function fnInsertData() As Boolean
    Dim strSQL As String
    Dim vData() As Variant
    Dim i As Long
    Dim nCount As Integer
    
    If t_nFormMode = ADD_MODE Then
        
        For i = 0 To tgmDetail.RowCount - 1
            tgmDetail.GetRow vData, nCount, i
        
            strSQL = "INSERT INTO item_cost_maint(icm_vendor, icm_code, icm_uom, icm_eff_date, icm_cost, icm_retail) VALUES("
            strSQL = strSQL & tfnRound(txtVendor) & "," & tfnSQLString(txtItemCode) & "," & tfnSQLString(vData(ColUOM)) & ","
            strSQL = strSQL & tfnDateString(vData(ColEffDate), True) & "," & tfnRound(vData(ColCost), 3) & "," & tfnRound(vData(ColRetail), 3) & ")"
            
            If Not fnExecuteSQL(strSQL) Then
                Exit Function
            End If
            
        Next
        
    ElseIf t_nFormMode = EDIT_MODE Then
        
        For i = 0 To tgmDetail.RowCount - 1
            tgmDetail.GetRow vData, nCount, i
            
            strSQL = "UPDATE item_cost_maint SET icm_uom = " & tfnSQLString(vData(ColUOM))
            strSQL = strSQL & ", icm_cost = " & tfnRound(vData(ColCost), 3) & ", icm_retail = " & tfnRound(vData(ColRetail), 3)
            strSQL = strSQL & " WHERE icm_vendor = " & txtVendor
            strSQL = strSQL & " AND icm_code = " & tfnSQLString(txtItemCode)
            strSQL = strSQL & " and icm_eff_date = " & tfnDateString(vData(ColEffDate), True)
            
            If Not fnExecuteSQL(strSQL) Then
                Exit Function
            End If
            
        Next
        
    End If
    
    fnInsertData = True
    
End Function
