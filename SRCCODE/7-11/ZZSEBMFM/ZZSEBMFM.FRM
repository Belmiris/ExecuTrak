VERSION 5.00
Object = "{C75015E0-2232-11D3-B440-0060971E99AF}#1.0#0"; "FACTFRM.OCX"
Object = "{01028C21-0000-0000-0000-000000000046}#4.0#0"; "TG32OV.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmZZSEBMFM 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Commission Master File Maintenance"
   ClientHeight    =   6348
   ClientLeft      =   1056
   ClientTop       =   1956
   ClientWidth     =   8964
   BeginProperty Font 
      Name            =   "Arial"
      Size            = 9.75 
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
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   5988
      Width           =   8964
      _Version        =   65536
      _ExtentX        =   15811
      _ExtentY        =   635
      _StockProps     =   77
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            = 9.75 
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
         Size            = 9.75 
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
      TabIndex        =   13
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
         Size        = 8.75   
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
         Size            = 9.75 
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
         Top             =   4932
         Width           =   1308
         _Version        =   65536
         _ExtentX        =   2302
         _ExtentY        =   688
         _StockProps     =   77
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            = 9.75 
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
            Size            = 9.75 
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
         Top             =   4932
         Width           =   1308
         _Version        =   65536
         _ExtentX        =   2302
         _ExtentY        =   688
         _StockProps     =   77
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            = 9.75 
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
            Size            = 9.75 
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
         TabIndex        =   6
         Top             =   4932
         Width           =   1308
         _Version        =   65536
         _ExtentX        =   2302
         _ExtentY        =   688
         _StockProps     =   77
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            = 9.75 
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
            Size            = 9.75 
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
         TabIndex        =   3
         Top             =   4932
         Width           =   1308
         _Version        =   65536
         _ExtentX        =   2302
         _ExtentY        =   688
         _StockProps     =   77
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            = 9.75 
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
            Size            = 9.75 
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
         TabIndex        =   4
         Top             =   4932
         Width           =   1308
         _Version        =   65536
         _ExtentX        =   2302
         _ExtentY        =   688
         _StockProps     =   77
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            = 9.75 
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
            Size            = 9.75 
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
         TabIndex        =   5
         Top             =   4932
         Width           =   1308
         _Version        =   65536
         _ExtentX        =   2302
         _ExtentY        =   688
         _StockProps     =   77
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            = 9.75 
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
            Size            = 9.75 
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin FACTFRMLib.FactorFrame efraOuterFrame 
         Height          =   4656
         Left            =   132
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   132
         Width           =   8676
         _Version        =   65536
         _ExtentX        =   15303
         _ExtentY        =   8213
         _StockProps     =   77
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            = 9.75 
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
            Size            = 9.75 
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin DBTrueGrid.TDBGrid tblBonus 
            Height          =   3660
            HelpContextID   =   404
            Left            =   84
            OleObjectBlob   =   "ZZSEBMFM.frx":0000
            TabIndex        =   11
            Top             =   924
            Width           =   8508
         End
         Begin FACTFRMLib.FactorFrame efraEmployee 
            Height          =   780
            Left            =   84
            TabIndex        =   20
            TabStop         =   0   'False
            Top             =   84
            Width           =   8520
            _Version        =   65536
            _ExtentX        =   15028
            _ExtentY        =   1376
            _StockProps     =   77
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            = 9.75 
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   6
            BeginProperty PanelFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            = 9.75 
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Begin VB.TextBox txtGrade 
               BackColor       =   &H00FFFFFF&
               Enabled         =   0   'False
               ForeColor       =   &H00000000&
               Height          =   360
               HelpContextID   =   400
               Left            =   6852
               TabIndex        =   23
               Tag             =   "pn_alt"
               Top             =   324
               Width           =   1560
            End
            Begin VB.TextBox txtEmpName 
               BackColor       =   &H00FFFFFF&
               DataSource      =   "datVendor"
               ForeColor       =   &H00000000&
               Height          =   360
               HelpContextID   =   402
               Left            =   2148
               TabIndex        =   9
               Tag             =   "pn_name"
               Top             =   324
               Width           =   4248
            End
            Begin VB.TextBox txtEmployee 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00000000&
               Height          =   360
               HelpContextID   =   400
               Left            =   84
               TabIndex        =   7
               Tag             =   "pn_alt"
               Top             =   324
               Width           =   1608
            End
            Begin FACTFRMLib.FactorFrame cmdEmployee 
               Height          =   360
               HelpContextID   =   401
               Left            =   1704
               TabIndex        =   8
               TabStop         =   0   'False
               Tag             =   "Run #3"
               Top             =   324
               Width           =   360
               _Version        =   65536
               _ExtentX        =   635
               _ExtentY        =   635
               _StockProps     =   77
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            = 9.75 
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               CaptionPos      =   4
               Picture         =   "ZZSEBMFM.frx":12DC
               Style           =   3
               BorderWidth     =   4
               BeginProperty PanelFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            = 9.75 
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin FACTFRMLib.FactorFrame cmdEmpName 
               Height          =   360
               HelpContextID   =   403
               Left            =   6408
               TabIndex        =   10
               TabStop         =   0   'False
               Tag             =   "Run #3"
               Top             =   324
               Width           =   360
               _Version        =   65536
               _ExtentX        =   635
               _ExtentY        =   635
               _StockProps     =   77
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            = 9.75 
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               CaptionPos      =   4
               Picture         =   "ZZSEBMFM.frx":13EE
               Style           =   3
               BorderWidth     =   4
               BeginProperty PanelFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            = 9.75 
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin VB.Label lblGrade 
               Caption         =   "Grade"
               Height          =   252
               Left            =   6852
               TabIndex        =   24
               Top             =   60
               Width           =   1632
            End
            Begin VB.Label lblEmpName 
               Caption         =   "Employee Name"
               Height          =   252
               Left            =   2148
               TabIndex        =   22
               Top             =   60
               Width           =   1956
            End
            Begin VB.Label lblEmpNumber 
               Caption         =   "Employee Number"
               Height          =   252
               Left            =   84
               TabIndex        =   21
               Top             =   60
               Width           =   1632
            End
         End
         Begin VB.PictureBox cmdDropdown 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            ForeColor       =   &H80000008&
            Height          =   240
            HelpContextID   =   22
            Left            =   108
            ScaleHeight     =   216
            ScaleWidth      =   228
            TabIndex        =   18
            TabStop         =   0   'False
            Top             =   948
            Visible         =   0   'False
            Width           =   255
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
         Size            = 9.75 
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Align           =   1
      FMName          =   "ZZSEBMFM"
      CaptionPos      =   4
      Style           =   6
      BeginProperty PanelFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            = 9.75 
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
      Bindings        =   "ZZSEBMFM.frx":1500
      Height          =   2484
      Left            =   36
      OleObjectBlob   =   "ZZSEBMFM.frx":151F
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   660
      Width           =   3756
   End
   Begin MSComctlLib.ProgressBar pbLock 
      Height          =   264
      Left            =   7044
      TabIndex        =   17
      Top             =   6036
      Visible         =   0   'False
      Width           =   1836
      _ExtentX        =   3239
      _ExtentY        =   466
      _Version        =   393216
      Appearance      =   1
   End
   Begin DBTrueGrid.TDBGrid tblDropDown 
      Bindings        =   "ZZSEBMFM.frx":2803
      Height          =   1008
      Left            =   324
      OleObjectBlob   =   "ZZSEBMFM.frx":281D
      TabIndex        =   19
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
Attribute VB_Name = "frmZZSEBMFM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***********************************************************
'
' Copyright (c) 2000 FACTOR, A Division of W.R.Hess Company
'
' Module name:   ZZSEBMFM
' Programmer :   Rajneesh Aggarwal
' Date       :   12/08/2000
'
'***********************************************************

Option Explicit

Private objCursor As clsCursor
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
Private Const szHelpEDI As String = "EDI888.HLP"
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
Private tgmBonus As clsTGSpreadSheet
Private cvalidate As cValidateInput
Private tgfDropdown As clsFloatingDropDown

Private Const ColPrftCtr As Integer = 0
Private Const ColBCode As Integer = 1
Private Const ColBCDesc As Integer = 2
Private Const colSequence As Integer = 3
Private Const ColElgDate As Integer = 4
Private Const ColStopDate As Integer = 5
Private Const ColElgClass As Integer = 6

Private sEmpFirstPrftCtr As String
Private sEmpPrftCtrList As String

Private nHdnBCLnk As Integer
'

Public Function fnValidCellValue(tgTable As TDBGrid, ByVal nCol As Integer, _
                                 ByVal lRow As Long, sText As String) As Boolean
    Dim strSQL As String
    
    Select Case nCol
        Case ColBCode, ColBCDesc
            fnValidCellValue = fnValidBonusCode(sText, lRow, nCol)
        Case ColPrftCtr
            fnValidCellValue = fnValidPrftCtr(sText, lRow, nCol)
        Case ColElgDate
            fnValidCellValue = fnValidDate(sText, lRow, nCol)
        Case ColElgClass
            fnValidCellValue = True
        Case colSequence
            fnValidCellValue = fnValidSequence(sText, lRow, nCol)
        Case ColStopDate
            fnValidCellValue = fnValidDate(sText, lRow, nCol)
    End Select
    
End Function

Private Sub subSetupMousePointer()
    Set objCursor = New clsCursor
    Set objCursor.MainForm = Me
    
    objCursor.AddControls tblComboDropdown, frmAbout
    
    objCursor.AddScreen
    
    'objCursor.AddArea efraInnerFrame
    objCursor.AddArea efraOuterFrame
    'objCursor.AddArea txtUserId
    'objCursor.SetLeftTop txtUserId
    'objCursor.SetRightBottom cmdUserIdSearchBtn
End Sub

Private Sub cmdDeleteBtn_Click()
    Dim sBCode As String
    Dim lRow As Long
    
    'Just a check
    If tgmBonus.RowCount < 1 Then
        Exit Sub
    End If
    
    lRow = tgmBonus.GetCurrentRowNumber
    sBCode = tgmBonus.CellValue(ColBCode, lRow)
    
    If t_nFormMode = EDIT_MODE Then
        If Not tfnCancelExit("Are you sure you want to delete the current record?") Then
            Exit Sub
        End If
        If Not fnDeleteBonusCode(sBCode, lRow) Then
            MsgBox "Failed to delete the Bonus code", vbCritical
            Exit Sub
        End If
    End If
    
    tgmBonus.DeleteRow
    
    If t_nFormMode = EDIT_MODE And tgmBonus.RowCount = 0 Then
        cmdTBCancel_Click
    End If
    
End Sub

Private Sub cmdDropDown_Click()
    subSetFloatingSQL
    tgfDropdown.ButtonClick
End Sub

Private Sub cmdDropDown_GotFocus()
    tgfDropdown.GotFocus cmdDropdown
End Sub

Private Sub cmdDropDown_LostFocus()
    tgfDropdown.LostFocus cmdDropdown
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
    App.HelpFile = szHelpEDI
End Sub

Private Sub Form_Unload(CANCEL As Integer)
    
    On Error Resume Next
    
    Set tgfDropdown = Nothing
    Set objCursor = Nothing
    Set objErrHandler = Nothing
    
    Set objCurrTabControl = Nothing
    
    Unload frmContext
    Unload frmAbout
    
    'project local database object variables
    
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
    #If FACTOR_MENU Then
        #If PROTOTYPE Then
            MsgBox "This is a prototype screen only.", vbExclamation
        #End If
    #End If
    
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
    
    subSetGridWidth tblBonus
    subInitSpreadsheets
    subSetFloatingDropDown
    subSetupCombos
    subInitValidation
    
    Me.Show
    Screen.MousePointer = vbHourglass
    tfnSetInitializingMessage
    Me.Enabled = False
    
    subSetupMousePointer
    t_bStartupFlag = False
    
    Me.Enabled = True
    tfnResetScreen 'set the default screen
    tmrKeyBoard.Enabled = True
    tfnSetStatusBarMessage t_szADDEDIT
    Screen.MousePointer = vbDefault
    
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
    t_bStartupFlag = False
    subEnableAddBtn False
    subEnableEditBtn False
    subEnableUpdateBtn False
    subEnableRefreshBtn False
    subEnableCancel True
    
    efraEmployee.Enabled = True
    subEnableSearchbtn cmdEmployee, True
    subEnableSearchbtn cmdEmpName, True

    cmdUpdateInsertBtn.Caption = t_szCAPTION_INSERT
    mnuUpdateInsert.Caption = t_szCAPTION_INSERT
    cmdRefreshSelectBtn.Caption = t_szCAPTION_REFRESH
    mnuRefreshSelect.Caption = t_szCAPTION_REFRESH
    
    subSetFocus txtEmployee
    
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
    subSetFocus efraBackground
    
    subEnableAddBtn False
    subEnableEditBtn False
    subEnableUpdateBtn False
    subEnableRefreshBtn False
    subEnableCancel True
    
    efraEmployee.Enabled = True
    subEnableSearchbtn cmdEmployee, True
    subEnableSearchbtn cmdEmpName, True
    
    cmdUpdateInsertBtn.Caption = t_szCAPTION_UPDATE
    mnuUpdateInsert.Caption = t_szCAPTION_UPDATE
    cmdRefreshSelectBtn.Caption = t_szCAPTION_REFRESH
    mnuRefreshSelect.Caption = t_szCAPTION_REFRESH
    
    subSetFocus txtEmployee
    
End Sub

Private Sub cmdUpdateInsertBtn_Click()
    Dim i As Long
    
    For i = 0 To tgmBonus.RowCount - 1
        If cmdUpdateInsertBtn.Caption = t_szCAPTION_INSERT Then
            If Not fnInsertBonusCode(i) Then
                MsgBox "Failed to insert the record", vbCritical
                Exit Sub
            End If
        Else
            If Not fnUpdateBonusCode(i) Then
                MsgBox "Failed to update the record", vbCritical
                Exit Sub
            End If
        End If
    Next i
        
    tfnResetScreen
    
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
    
    Screen.MousePointer = vbHourglass
    fnLoadData
    Screen.MousePointer = vbDefault
    
    nDataStatus = DATA_INIT
    subEnableRefreshBtn False
    subEnableUpdateBtn False
    
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
    
    t_nFormMode = IDLE_MODE
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
    
    Screen.MousePointer = vbHourglass
    subEnableEmployee True
    
    subEnableAddBtn True
    subEnableEditBtn True
    subEnableUpdateBtn False
    subEnableDeleteBtn False
    subEnableRefreshBtn False
    tblBonus.Enabled = False
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
    
    sEmpFirstPrftCtr = ""
    txtEmployee = ""
    txtEmpName = ""
    txtGrade = ""
    
    cvalidate.ResetFlags
    tgmBonus.ClearData
    tblBonus.col = ColPrftCtr
    
    nDataStatus = DATA_INIT
    subEnableCancel False
    
    efraEmployee.Enabled = False
    subEnableSearchbtn cmdEmployee, False
    subEnableSearchbtn cmdEmpName, False
    
    Screen.MousePointer = vbDefault
    subSetFocus cmdAddBtn
    DoEvents
    
End Sub

Private Sub tblDropDown_SelChange(CANCEL As Integer)
    CANCEL = True
End Sub

Private Sub tblBonus_AfterColEdit(ByVal ColIndex As Integer)
    Dim lRow As Long
    
    lRow = tgmBonus.GetCurrentRowNumber
    tgmBonus.AfterColEdit ColIndex
    
    If t_nFormMode = EDIT_MODE Then
        If nDataStatus = DATA_CHANGED Then
            subEnableRefreshBtn True
        Else
            subEnableRefreshBtn False
        End If
    End If

End Sub

Private Sub tblBonus_BeforeColEdit(ByVal ColIndex As Integer, ByVal KeyAscii As Integer, CANCEL As Integer)
    tgmBonus.BeforeColEdit ColIndex, KeyAscii, CANCEL
End Sub

Private Sub tblBonus_Change()
    tgmBonus.Change
    If tblBonus.col = ColBCode Then
        tgmBonus.CellValue(ColBCDesc, tgmBonus.GetCurrentRowNumber) = ""
    End If
    
    subEnableUpdateBtn False
    nDataStatus = DATA_CHANGED
    
End Sub

Private Sub tblBonus_FirstRowChange()
    tgmBonus.FirstRowChange
    tgfDropdown.FirstRowChange
    
    If tblBonus.Row = -1 Then
        subEnableDeleteBtn False
    End If
End Sub

Private Sub tblBonus_GotFocus()
    tfnSetStatusBarMessage "Commission Master"
    tgmBonus.GotFocus
    tgfDropdown.GotFocus
End Sub

Private Sub tblBonus_KeyDown(KeyCode As Integer, Shift As Integer)
    tgmBonus.KeyDown KeyCode, Shift
    tgfDropdown.KeyDown tblBonus, KeyCode
End Sub

Private Sub tblBonus_KeyPress(KeyAscii As Integer)
    Dim lRow As Long
    
    If tblBonus.col = ColBCode Then
        KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
    End If
    
    tgfDropdown.Keypress tblBonus, KeyAscii
    lRow = tgmBonus.GetCurrentRowNumber
    
    If Not tgmBonus.Keypress(KeyAscii) Then
        KeyAscii = 0
    End If
    
    If t_nFormMode = EDIT_MODE Then
        If KeyAscii = vbKeyReturn And tgmBonus.RowCount - 1 = lRow And tblBonus.col = ColStopDate Then
            If cmdUpdateInsertBtn.Enabled Then
                subSetFocus cmdUpdateInsertBtn
            Else
                subSetFocus cmdExitCancelBtn
            End If
        End If
    End If
    
End Sub

Private Sub tblBonus_LeftColChange()
    tgfDropdown.LeftColChange
End Sub

Private Sub tblBonus_LostFocus()
    tgmBonus.LostFocus
    tgfDropdown.LostFocus tblBonus
    subSetStdBtn
End Sub

Private Sub tblBonus_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    frmContext.MouseDown Button, EMP_MST_UP
End Sub

Private Sub tblBonus_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    Dim lRow As Long
    
    If tgfDropdown.RowColChange(tblBonus) Then
        tgmBonus.RowColChange LastRow, LastCol
    End If
    
    lRow = tgmBonus.GetCurrentRowNumber
    
    If t_nFormMode = ADD_MODE Then
        If tgmBonus.CellValue(ColBCode, lRow) <> "" Then
            subEnableDeleteBtn True
        Else
            subEnableDeleteBtn False
        End If
    Else
        If tgmBonus.RowCount > 0 Then
            subEnableDeleteBtn True
        Else
            subEnableDeleteBtn False
        End If
    End If

    subSetStdBtn
    
End Sub

Private Sub tblBonus_SelChange(CANCEL As Integer)
    CANCEL = True
End Sub

Private Sub tblBonus_UnboundReadData(ByVal RowBuf As DBTrueGrid.RowBuffer, StartLocation As Variant, ByVal ReadPriorRows As Boolean)
    tgmBonus.ReadData RowBuf, StartLocation, ReadPriorRows
End Sub

Private Sub tblBonus_HeadClick(ByVal ColIndex As Integer)
    tgmBonus.HeadClick ColIndex, nDataStatus = DATA_CHANGED
End Sub

Private Sub tmrKeyboard_Timer() 'status bar timer - 250ms
    tfnUpdateStatusBar Me 'process the status bar
    objCursor.SetMousePointer
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
        .AddButton "Add &Employee", EMP_MST_UP
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

Private Sub subSetGridWidth(tbl As TDBGrid)
    Dim myWidth As Variant
    Dim myField As Variant
    
    Dim i As Integer
    
    myWidth = Array(0.11, 0.11, 0.36, 0.08, 0.12, 0.12, 0.1)
    myField = Array("bm_eligible_pc", "bm_bonus_code", "bc_code_desc", _
        "bm_sequence", "bm_eligible_date", "bm_stop_date", "bm_eligible_class")
    
    While tbl.Columns.Count > 0
        tbl.Columns.Remove 0
    Wend
    
    For i = 0 To UBound(myWidth)
        tbl.Columns.Add i
        With tbl.Columns(i)
            .Width = myWidth(i) * tbl.Width - 50
            .DataField = myField(i)
            .Visible = True
            .HeadAlignment = vbCenter
        End With
    Next
    
    'tbl.Caption = "Commission Master"
    tbl.Columns(ColPrftCtr).Caption = "Profit Center"
    tbl.Columns(ColBCode).Caption = "Bonus Code"
    tbl.Columns(ColBCDesc).Caption = "Bonus Code Description"
    tbl.Columns(colSequence).Caption = "Seq."
    tbl.Columns(ColElgDate).Caption = "Eligible Date"
    tbl.Columns(ColStopDate).Caption = "Stop Date"
    tbl.Columns(ColElgClass).Caption = "Eligible Class"

End Sub
Private Sub subInitSpreadsheets()

    If Not tgmBonus Is Nothing Then
        Set tgmBonus = Nothing
    End If
    
    Set tgmBonus = New clsTGSpreadSheet
    Set tgmBonus.Table = tblBonus
    Set tgmBonus.StatusBar = ffraStatusbar ' message bar name
    Set tgmBonus.Form = Me
    Set tgmBonus.engFactor = t_engFactor
    
    With tgmBonus
        .AddEditColumn ColPrftCtr, "Enter Profit Center", szIntegerPattern
        .AddEditColumn ColBCode, "Enter Bonus Code", "^P{1,4}$"
        .AddEditColumn ColBCDesc, "Enter Bonus Code Description", "^P{1,40}"
        .AddEditColumn colSequence, "Enter Sequence", szIntegerPattern
        .AddEditColumn ColElgDate, "Enter Eligible Date", szDatePattern
        .AddEditColumn ColStopDate, "Enter Stop Date", szDatePattern
        .AddEditColumn ColElgClass, "Enter Eligible Class", "^P{1,2}$"
        .ColumnForNewRow = 0
        nHdnBCLnk = .AddHiddenField("bm_bonus_code")
        
        .AddSortColumn ColPrftCtr, ColPrftCtr, .NUMERIC_TYPE, .ASCENDING, .CASE_SENSITIVE, _
            colSequence, .NUMERIC_TYPE, .ASCENDING, .CASE_SENSITIVE
        
        .AddSortColumn ColBCode, ColBCode, .STRING_TYPE, .ASCENDING, .CASE_SENSITIVE, _
            ColPrftCtr, .NUMERIC_TYPE, .ASCENDING, .CASE_SENSITIVE
            
        .SortByColumn = True
    End With
    
    tgmBonus.AllowAddNew = False
    
End Sub

Private Sub subSetFloatingDropDown()

    If Not tgfDropdown Is Nothing Then
        Set tgfDropdown = Nothing
    End If

    Set tgfDropdown = New clsFloatingDropDown
    
    With tgfDropdown
        Set .DataBase = t_dbMainDatabase
        Set .SearchButton = cmdDropdown
        Set .MainTable = tblBonus
        Set .DropDownTable = tblDropDown
        Set .DataLink = datDropDown
        Set .Form = Me
        Set .EditClass = tgmBonus
        
        .SearchOnReturn = False
        .DefaultCursorControl = True
        
        .AddDropDown 1
        .AddColumn ColBCode, "bc_bonus_code", .COLUMN_TYPE_STRING
        .AddColumn ColBCDesc, "bc_code_desc", .COLUMN_TYPE_STRING
        
        .AddDropDown 2
        .AddColumn ColPrftCtr, "prft_ctr", .COLUMN_TYPE_INTEGER
        .AddExtraColumn "Description", "prft_name", 3000
        
    End With
    
End Sub

Private Sub subSetFloatingSQL()
    Dim sSql As String
    Dim lRow As Long
    Dim i As Long
    Dim sCheckDubInGrid
    
    lRow = tgmBonus.GetCurrentRowNumber
    
    If tblBonus.col = ColBCode Or tblBonus.col = ColBCDesc Then
        sCheckDubInGrid = fnCheckDubInGrid(ColBCode, lRow)
        sSql = "SELECT bc_bonus_code, bc_code_desc FROM bonus_codes WHERE 1 = 1"
        If t_nFormMode = ADD_MODE Then
            sSql = sSql & " AND bc_bonus_code NOT IN(SELECT bm_bonus_code FROM bonus_master"
            sSql = sSql & " WHERE bm_empno = " & tfnRound(txtEmployee) & ")"
        End If
        If sCheckDubInGrid <> "" Then
            sSql = sSql & " AND bc_bonus_code NOT IN(" & sCheckDubInGrid & ")"
        End If
    ElseIf tblBonus.col = ColPrftCtr Then
        sSql = "SELECT prft_ctr, prft_name FROM sys_prft_ctr WHERE prft_ctr IN ("
        sSql = sSql + sEmpPrftCtrList + ")"
    End If
    
    tgfDropdown.SetSQL tblBonus.col, sSql

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

Private Function fnDeleteBonusCode(sCode As String, lRow As Long) As Boolean
    Const SUB_NAME As String = "fnDeleteBonusCode"
    Dim strSQL As String
    
    fnDeleteBonusCode = False
    
    strSQL = "DELETE FROM bonus_master WHERE bm_bonus_code = " & tfnSQLString(Trim(sCode))
    strSQL = strSQL & " AND bm_empno = " & tfnRound(txtEmployee)
    
    If fnExecuteSQL(strSQL, , SUB_NAME) Then
        fnDeleteBonusCode = True
    End If

End Function

Private Function fnInsertBonusCode(lRow As Long) As Boolean
    Const SUB_NAME As String = "fnInsertBonusCode"
    Dim strSQL As String
    Dim rsTemp As Recordset
    
    fnInsertBonusCode = False
    
    If Trim(tgmBonus.CellValue(ColBCode, lRow)) = "" Then
        fnInsertBonusCode = True
        Exit Function
    End If
    
    strSQL = "INSERT INTO bonus_master(bm_empno, bm_bonus_code, bm_sequence, bm_eligible_date,"
    strSQL = strSQL & " bm_eligible_class, bm_eligible_pc, bm_stop_date, bm_entry_id, bm_entry_date)"
    strSQL = strSQL & "VALUES(" & tfnRound(txtEmployee) & ", "
    strSQL = strSQL & tfnSQLString(Trim(tgmBonus.CellValue(ColBCode, lRow))) & ", "
    strSQL = strSQL & tfnRound(tgmBonus.CellValue(colSequence, lRow)) & ", "
    strSQL = strSQL & tfnDateString(Trim(tgmBonus.CellValue(ColElgDate, lRow)), True) & ", "
    strSQL = strSQL & tfnSQLString(Trim(tgmBonus.CellValue(ColElgClass, lRow))) & ", "
    strSQL = strSQL & tfnRound(tgmBonus.CellValue(ColPrftCtr, lRow)) & ", "
    strSQL = strSQL & tfnDateString(Trim(tgmBonus.CellValue(ColStopDate, lRow)), True) & ", "
    strSQL = strSQL & tfnSQLString(Trim(tfnGetUserName)) & ", "
    strSQL = strSQL & tfnDateString(Trim(Date), True) & ")"
    
    If fnExecuteSQL(strSQL, , SUB_NAME) Then
        fnInsertBonusCode = True
    End If

End Function

Private Function fnLoadData() As Boolean
    Const SUB_NAME As String = "fnLoadData"
    Dim strSQL As String
    Dim rsTemp As Recordset
    Dim sLockedUser As String
    
    fnLoadData = False
    
    tfnSetStatusBarMessage "Loading Data. Please wait ..."
    
    'Lock the whole table.
    strSQL = "bm_empno = " & tfnRound(txtEmployee)
    If Not tfnLockRow("ZZSEBMFM", "bonus_master", strSQL, , sLockedUser) Then
        tfnSetStatusBarError "Commission record for the employee is locked by " + sLockedUser
        Exit Function
    End If
    
    strSQL = "SELECT bm_eligible_pc, bm_bonus_code, bc_code_desc, bm_sequence, "
    strSQL = strSQL & " bm_eligible_date, bm_stop_date, bm_eligible_class"
    strSQL = strSQL & " FROM bonus_master, bonus_codes"
    strSQL = strSQL & " WHERE bc_bonus_code = bm_bonus_code"
    strSQL = strSQL & " AND bm_empno = " & tfnRound(txtEmployee)
    strSQL = strSQL & " ORDER BY bm_eligible_pc, bm_bonus_code, bm_sequence"
    
    If GetRecordSet(rsTemp, strSQL, , SUB_NAME) < 0 Then
        MsgBox "Failed to access database to get records", vbCritical
        tfnSetStatusBarError "Failed to access database to get records"
        Exit Function
    End If
    
    If rsTemp.RecordCount = 0 Then
        MsgBox "No Commission record found for the employee", vbExclamation
        tfnSetStatusBarError "No record available for Edit"
        Exit Function
    End If

    tgmBonus.FillWithRecordset rsTemp
    
    fnLoadData = True
        
End Function

Private Function fnUpdateBonusCode(lRow As Long) As Boolean
    Const SUB_NAME As String = "fnUpdateBonusCode"
    Dim strSQL As String
    Dim rsTemp As Recordset
  
    fnUpdateBonusCode = False
    
    strSQL = "UPDATE bonus_master SET bm_bonus_code = " & tfnSQLString(Trim(tgmBonus.CellValue(ColBCode, lRow))) & ","
    strSQL = strSQL & " bm_sequence = " & tfnRound(tgmBonus.CellValue(colSequence, lRow)) & ","
    strSQL = strSQL & " bm_eligible_date = " & tfnDateString(Trim(tgmBonus.CellValue(ColElgDate, lRow)), True) & ","
    strSQL = strSQL & " bm_eligible_class = " & tfnSQLString(Trim(tgmBonus.CellValue(ColElgClass, lRow))) & ","
    strSQL = strSQL & " bm_eligible_pc = " & Trim(tgmBonus.CellValue(ColPrftCtr, lRow)) & ","
    strSQL = strSQL & " bm_stop_date = " & tfnDateString(Trim(tgmBonus.CellValue(ColStopDate, lRow)), True) & ","
    strSQL = strSQL & " bm_entry_id = " & tfnSQLString(Trim(tfnGetUserName)) & ","
    strSQL = strSQL & " bm_entry_date = " & tfnDateString(Trim(Date), True)
    strSQL = strSQL & " WHERE bm_bonus_code = " & tfnSQLString(Trim(tgmBonus.CellValue(nHdnBCLnk, lRow)))
    strSQL = strSQL & " AND bm_empno = " & tfnRound(txtEmployee)
    
    If fnExecuteSQL(strSQL, , SUB_NAME) Then
        fnUpdateBonusCode = True
    End If

End Function

Private Function fnCheckDubInGrid(nCol_Nbr As Integer, lCurrRow As Long) As String
'This function will return all the values already used in the grid.
    Dim lRowCount As Long
    Dim k As Integer
    Dim sTemp As String
    
    fnCheckDubInGrid = ""
    
    lRowCount = tgmBonus.RowCount
    sTemp = ""
    
    If lRowCount > 0 Then
        For k = 0 To lRowCount - 1
            If k <> lCurrRow Then
                If sTemp = "" Then
                    sTemp = tfnSQLString(tgmBonus.CellValue(nCol_Nbr, k))
                Else
                    sTemp = sTemp & "," & tfnSQLString(tgmBonus.CellValue(nCol_Nbr, k))
                End If
            End If
        Next
    End If
    
    fnCheckDubInGrid = sTemp

End Function

Private Sub subSetStdBtn()
    
    If tgmBonus.RowCount < 1 Then
        subEnableUpdateBtn False
        subEnableDeleteBtn False
    Else
        If tgmBonus.ValidData Then
            If t_nFormMode = ADD_MODE Then
                subEnableUpdateBtn True
            Else
                If nDataStatus = DATA_CHANGED Then
                    subEnableUpdateBtn True
                Else
                    subEnableUpdateBtn False
                End If
            End If
        Else
            subEnableUpdateBtn False
        End If
    End If
End Sub

Private Sub txtEmployee_Change()
    cvalidate.Change txtEmployee
    tgcDropdown.Change txtEmployee
    cmdExitCancelBtn.Caption = t_szCAPTION_CANCEL
    txtEmpName = ""
End Sub

Private Sub txtEmployee_GotFocus()
    
    If t_bStartupFlag Then
        Exit Sub
    End If
    
    tgcDropdown.GotFocus txtEmployee
    cvalidate.GotFocus txtEmployee
    
    If tgcDropdown.SingleRecordSelected Then
        subEnterPhaseTwo
        Screen.MousePointer = vbDefault
    End If
    
End Sub

Private Sub txtEmployee_KeyPress(KeyAscii As Integer)
    Dim bKeyCode As Boolean
    
    If KeyAscii = vbKeyReturn Then
        tgcDropdown.ComboSQL(txtEmployee) = fnSetComboSQL(txtEmployee.TabIndex)
        Screen.MousePointer = vbHourglass
    End If
        
    bKeyCode = tgcDropdown.Keypress(txtEmployee, KeyAscii)
    Screen.MousePointer = vbDefault
    
    If Not bKeyCode Then
       If KeyAscii = vbKeyReturn Then
          If tgcDropdown.SingleRecordSelected Then
                subEnterPhaseTwo
                Screen.MousePointer = vbDefault
          End If
          KeyAscii = 0
       End If
    Else
        cvalidate.Keypress txtEmployee, KeyAscii
    End If

End Sub

Private Sub txtEmployee_LostFocus()
    tgcDropdown.LostFocus txtEmployee
    cvalidate.LostFocus txtEmployee, cmdEmployee
    
    If txtEmployee <> "" Then
        't_bDataChanged = True
    End If
    Screen.MousePointer = vbDefault

End Sub

Private Sub txtEmployee_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    'frmContext.MouseDown Button, EMP_MST_UP
End Sub

Private Sub txtEmpName_Change()
    tgcDropdown.Change txtEmpName
    cmdExitCancelBtn.Caption = t_szCAPTION_CANCEL
End Sub

Private Sub txtEmpName_GotFocus()
    tfnSetStatusBarMessage "Enter Employee Name"
    tgcDropdown.GotFocus txtEmpName
    Screen.MousePointer = vbDefault
    
    If tgcDropdown.SingleRecordSelected Then
        subEnterPhaseTwo
        Screen.MousePointer = vbDefault
    End If
End Sub

Private Sub txtEmpName_KeyPress(KeyAscii As Integer)
    Dim bKeyCode As Boolean
    
    If KeyAscii = vbKeyReturn Then
        tgcDropdown.ComboSQL(txtEmpName) = fnSetComboSQL(txtEmpName.TabIndex)
        Screen.MousePointer = vbHourglass
    End If
    
    bKeyCode = tgcDropdown.Keypress(txtEmpName, KeyAscii)
    Screen.MousePointer = vbDefault
    
    If Not bKeyCode Then
        If KeyAscii = vbKeyReturn Then
            If tgcDropdown.SingleRecordSelected Then
                subEnterPhaseTwo
                Screen.MousePointer = vbDefault
            End If
        KeyAscii = 0
        End If
    End If

End Sub

Private Sub cmdEmployee_Click()
    Screen.MousePointer = vbHourglass
    tgcDropdown.ComboSQL(txtEmployee) = fnSetComboSQL(txtEmployee.TabIndex)
    tgcDropdown.Click cmdEmployee
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdEmpName_Click()
    tgcDropdown.ComboSQL(txtEmpName) = fnSetComboSQL(txtEmpName.TabIndex)
    tgcDropdown.Click cmdEmpName
End Sub

Private Sub subInitValidation()
    Set cvalidate = New cValidateInput
    Set cvalidate.Form = Me
    Set cvalidate.StatusBar = ffraStatusbar
    
    cvalidate.AddEditBox txtEmployee, "Enter Employee Number"
    cvalidate.MinTabIndex = txtEmployee.TabIndex
    cvalidate.MaxTabIndex = txtEmpName.TabIndex
    
End Sub

Private Function fnSetComboSQL(nTabIndex) As String
    Dim strSQL As String
    Select Case nTabIndex
        Case txtEmployee.TabIndex, txtEmpName.TabIndex
            strSQL = "SELECT prm_empno, prm_empname FROM sTmpCustTable"
            If t_nFormMode = EDIT_MODE Then
                strSQL = strSQL & " WHERE prm_empno IN (SELECT bm_empno FROM bonus_master)"
            End If
    End Select
    fnSetComboSQL = strSQL
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
    Dim i As Integer
    Dim sTemp As String
    
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
        strSQL = "SELECT * FROM pr_master WHERE prm_empno = " & tfnRound(Box.Text)
    Else
        strSQL = "SELECT * FROM pr_master WHERE"
        strSQL = strSQL & " prm_empno IN (SELECT DISTINCT bm_empno FROM bonus_master)"
        strSQL = strSQL & " AND prm_empno = " & Box.Text
    End If
    
    If GetRecordSet(rsTemp, strSQL, , SUB_NAME) < 0 Then
        cvalidate.SetErrorMessage Box, "Failed to access Database"
        Exit Function
    End If
            
    If rsTemp.RecordCount = 0 Then
        cvalidate.SetErrorMessage Box, "Employee Number does not exist"
        Exit Function
    End If
    
    If t_nFormMode = ADD_MODE Then
        sEmpFirstPrftCtr = fnGetField(rsTemp!prm_prft_ctr1)
    End If
    
    'make prft_ctr list
    sEmpPrftCtrList = fnGetField(rsTemp!prm_prft_ctr1)
    
    For i = 2 To 5
        sTemp = fnGetField(rsTemp.Fields("prm_prft_ctr" & i))
        If sTemp <> "" Then
            sEmpPrftCtrList = sEmpPrftCtrList + "," + sTemp
        End If
    Next i
    
    If sEmpPrftCtrList <> "" Then
        If Left(sEmpPrftCtrList, 1) = "," Then
            sEmpPrftCtrList = Mid(sEmpPrftCtrList, 2)
        End If
    End If
    
    txtGrade = tfnRound(rsTemp!prm_emp_level)
    
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
    Exit Sub
    Dim bCode As Boolean
    
    bCode = tgcDropdown.Keypress(tblComboDropdown, KeyAscii)
    
    If Not bCode Then
        KeyAscii = 0
    End If
    
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
    Set tgcDropdown.Form = Me
    Set tgcDropdown.DBEngine = t_engFactor
    Set tgcDropdown.DataBase = t_dbMainDatabase
    Set tgcDropdown.DataLink = datComboDropDown
    Set tgcDropdown.Table = tblComboDropdown
    
    #If PROTOTYPE Then
        Exit Sub
    #End If
    
     With tgcDropdown
        .AddCombo
        .AddComboBox txtEmployee, cmdEmployee, "prm_empno", .SQL_LONG_TYPE
        .AddComboBox txtEmpName, cmdEmpName, "prm_empname", .SQL_STRING_TYPE(60)
     End With
End Sub

Private Sub subEnterPhaseTwo()

    Screen.MousePointer = vbHourglass

    If Not tfnLockRow(efraToolBar.FMName, "bm_master", txtEmployee) Then
        Exit Sub
    End If
    
    Screen.MousePointer = vbHourglass
    subEnableEmployee False
    
    tblBonus.Enabled = True
    
    If t_nFormMode = ADD_MODE Then
        tgmBonus.AllowAddNew = True
    Else
        tgmBonus.AllowAddNew = False
        If Not fnLoadData Then
            cmdTBCancel_Click
        End If
        nDataStatus = DATA_LOADED
    End If
    
    subSetFocus tblBonus
    Screen.MousePointer = vbDefault
    
End Sub

Private Sub subEnableEmployee(bOnOff As Boolean)
    efraEmployee.Enabled = bOnOff
    txtEmployee.Enabled = bOnOff
    cmdEmployee.Enabled = bOnOff
    txtEmpName.Enabled = bOnOff
    cmdEmpName.Enabled = bOnOff
    subEnableSearchbtn cmdEmployee, bOnOff
    subEnableSearchbtn cmdEmpName, bOnOff
End Sub

Private Function fnValidBonusCode(szText As String, lRow As Long, nCol As Integer) As Boolean
    Const SUB_NAME As String = "fnValidBonusCode"
    Dim strSQL As String
    Dim rsTemp As Recordset
    Dim sBonusCode As String
    Dim sBCDesc As String
    Dim sCheckDubInGrid As String
    Dim arrBonus() As String
    Dim i As Long
    
    fnValidBonusCode = False
    
    sBonusCode = tgmBonus.CellValue(ColBCode, lRow)
    If nCol = ColBCDesc Then
        If Trim(sBonusCode) <> "" And tgmBonus.ValidCell(ColBCode, lRow) Then
            sBCDesc = fnGetBonusCodeDesc(sBonusCode)
            tgmBonus.CellValue(nCol, lRow) = sBCDesc
        End If
        fnValidBonusCode = True
        Exit Function
    End If

    sCheckDubInGrid = fnCheckDubInGrid(ColBCode, lRow)
    
    If Trim(szText) = "" Then
        tgmBonus.ErrorMessage(ColBCode) = "You must enter a valid Bonus code"
        Exit Function
    End If
    
    strSQL = "SELECT * FROM bonus_codes WHERE bc_bonus_code = " & tfnSQLString(Trim(szText))
    If GetRecordSet(rsTemp, strSQL, , SUB_NAME) < 1 Then
        tgmBonus.ErrorMessage(ColBCode) = "Bonus code does not exist"
        Exit Function
    End If
    
    'Check Dublication in the Table.
    If t_nFormMode = ADD_MODE Then
        strSQL = "SELECT * FROM bonus_master WHERE bm_bonus_code = " & tfnSQLString(szText)
        strSQL = strSQL & " AND bm_empno = " & tfnRound(txtEmployee)
        If GetRecordCount(strSQL, , SUB_NAME) > 0 Then
            tgmBonus.ErrorMessage(ColBCode) = "Bonus code already exists in the table"
            Exit Function
        End If
    End If
    
    'Check Dublication in the Grid.
    If sCheckDubInGrid <> "" Then
        arrBonus = Split(sCheckDubInGrid, ",")
        For i = 0 To UBound(arrBonus)
            If tfnSQLString(szText) = arrBonus(i) Then
                tgmBonus.ErrorMessage(ColBCode) = "Bonus code already exists in the grid"
                Exit Function
            End If
        Next i
    End If
    
    If sEmpFirstPrftCtr <> "" Then
        tgmBonus.CellValue(ColPrftCtr, lRow) = sEmpFirstPrftCtr
    End If
    tgmBonus.CellValue(ColElgDate, lRow) = Date
    
    fnValidBonusCode = True

End Function

Private Function fnCreateSearchTable(szNumber As String, szName As String) As Boolean
    Dim sSql As String
    Dim sSQL1 As String
    Dim sSQL2 As String
    Dim sSQLL  As String
    Dim sSqDrop As String

    'Drop temp table first
    sSqDrop = " DROP TABLE sTmpCustTable"
    On Error GoTo ErrorDropTable
    t_dbMainDatabase.ExecuteSQL sSqDrop
    'Create a temp table
    sSql = " CREATE TEMP  TABLE sTmpCustTable " & " ( " & szNumber & " INTEGER," & szName & " CHAR(60)) "
    On Error GoTo errCreateTable
    t_dbMainDatabase.ExecuteSQL sSql
    On Error GoTo errInsertRecords
    sSQL1 = " INSERT INTO  sTmpCustTable " & " SELECT prm_empno, TRIM(prm_last_name) || ', '"
    sSQL1 = sSQL1 & " || TRIM(prm_first_name) || ' '  || TRIM(prm_middle_name) "
    sSQL1 = sSQL1 & " FROM pr_master WHERE  TRIM(prm_middle_name)<> ''"
    'sSQL1 = sSQL1 & " AND prm_security_code <= " & tfnSQLString(Security_Code)
    t_dbMainDatabase.ExecuteSQL sSQL1
    sSQL2 = " INSERT INTO  sTmpCustTable " & " SELECT prm_empno, TRIM(prm_last_name) || ', '"
    sSQL2 = sSQL2 & " || TRIM(prm_first_name)  "
    sSQL2 = sSQL2 & " FROM pr_master WHERE  TRIM(prm_middle_name) = '' OR prm_middle_name IS NULL"
    'sSQL2 = sSQL2 & " AND prm_security_code <= " & tfnSQLString(Security_Code)
    t_dbMainDatabase.ExecuteSQL sSQL2
    fnCreateSearchTable = True
    Exit Function
ErrorDropTable:
    Resume Next
errCreateTable:
  '  MsgBox "Cannot create temporary table for searching.", vbOKOnly + vbCritical, App.Title
    tfnErrHandler "fnCreateSearchTable", sSql
    Err.Clear
    fnCreateSearchTable = False
    Exit Function
errInsertRecords:
  '  MsgBox "Cannot insert records into temmporary table.", vbOKOnly + vbCritical, App.Title
  tfnErrHandler "fnCreateSearchTable", sSQL1 & "or " & sSQL2
    Err.Clear
    On Error Resume Next
    fnCreateSearchTable = False
End Function

Private Function fnValidPrftCtr(szText As String, lRow As Long, nCol As Integer) As Boolean
    Const SUB_NAME As String = "fnValidPrftCtr"
    Dim strSQL As String
    Dim rsTemp As Recordset
    
    fnValidPrftCtr = False
    
    If Trim(szText) = "" Then
        tgmBonus.ErrorMessage(ColPrftCtr) = "You must enter a valid profit center"
        Exit Function
    End If
    
    strSQL = "SELECT * FROM sys_prft_ctr WHERE prft_ctr = " & Trim(szText)
    If GetRecordSet(rsTemp, strSQL, , SUB_NAME) < 1 Then
        tgmBonus.ErrorMessage(ColPrftCtr) = "Profit center does not exist"
        Exit Function
    End If
    
    fnValidPrftCtr = True

End Function

Private Function fnGetBonusCodeDesc(sCode As String) As String
    Const SUB_NAME As String = "fnGetBonusCodeDesc"
    Dim strSQL As String
    Dim rsTemp As Recordset
    
    fnGetBonusCodeDesc = ""
    
    strSQL = "SELECT bc_code_desc FROM bonus_codes WHERE bc_bonus_code = "
    strSQL = strSQL & tfnSQLString(Trim(sCode))
    
    If GetRecordSet(rsTemp, strSQL, , SUB_NAME) = 1 Then
        fnGetBonusCodeDesc = fnGetField(rsTemp!bc_code_desc)
    End If
    
End Function

Private Function fnValidDate(szText As String, lRow As Long, nCol As Integer) As Boolean
    Const SUB_NAME As String = "fnValidDate"
    Dim sTemp As String
    
    fnValidDate = False
    
    sTemp = "Eligible Date "
    If nCol = ColStopDate Then
        sTemp = "Stop Date "
    End If

    If szText = "" Then
        tgmBonus.ErrorMessage(nCol) = sTemp & "must be entered"
        Exit Function
    End If
    
    tgmBonus.CellValue(nCol, lRow) = CDate(tfnFormatDate(szText))
    
    If nCol = ColElgDate Then
        fnValidDate = True
        Exit Function
    End If
    
    If nCol = ColStopDate Then
        If CDate(tfnFormatDate(szText)) < CDate(tgmBonus.CellValue(ColElgDate, lRow)) Then
            tgmBonus.ErrorMessage(nCol) = sTemp + " must be earlier than Stop Date"
            tgmBonus.ErrorMessage(ColElgDate) = sTemp + " must be later than Stop Date"
            tgmBonus.ValidCell(ColElgDate, lRow) = False
            Exit Function
        End If
    Else
        If CDate(tfnFormatDate(szText)) < CDate(tgmBonus.CellValue(ColElgDate, lRow)) Then
            tgmBonus.ErrorMessage(nCol) = sTemp + " must be earlier than Stop Date"
            tgmBonus.ErrorMessage(ColStopDate) = sTemp + " must be later than Stop Date"
            tgmBonus.ValidCell(ColStopDate, lRow) = False
            Exit Function
        End If
    End If
    
    If nCol = ColStopDate Then
        tgmBonus.ValidCell(ColElgDate, lRow) = True
    Else
        tgmBonus.ValidCell(ColStopDate, lRow) = True
    End If
    
    fnValidDate = True

End Function

Private Function fnValidSequence(szText As String, lRow As Long, nCol As Integer) As Boolean
    Const SUB_NAME As String = "fnValidSequence"
    Dim strSQL As String
    Dim sCheckDubInGrid As String
    Dim sArSeq() As String
    Dim i As Integer
    
    fnValidSequence = False
    sCheckDubInGrid = fnCheckDubInGrid(colSequence, lRow)
    
    If Trim(szText) = "" Then
        tgmBonus.ErrorMessage(nCol) = "You must enter a sequence"
        Exit Function
    End If
    
    strSQL = "SELECT * FROM bonus_master WHERE bm_empno = " & tfnRound(txtEmployee)
    strSQL = strSQL & " AND bm_sequence = " & tfnRound(szText)
    
    If GetRecordCount(strSQL, , SUB_NAME) > 0 Then
        tgmBonus.ErrorMessage(nCol) = "Sequence already exists in the table"
        Exit Function
    End If
    
    'Check Dublication in the Grid...
    If sCheckDubInGrid <> "" Then
        sArSeq = Split(sCheckDubInGrid, ",")
        For i = 0 To UBound(sArSeq)
            If tfnSQLString(szText) = sArSeq(i) Then
                tgmBonus.ErrorMessage(nCol) = "Sequence already exists in the grid"
                Exit Function
            End If
        Next i
    End If
    
    fnValidSequence = True

End Function
