VERSION 5.00
Object = "{01028C21-0000-0000-0000-000000000046}#4.0#0"; "TG32OV.OCX"
Object = "{C75015E0-2232-11D3-B440-0060971E99AF}#1.0#0"; "FACTFRM.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmZZMPNSMT 
   BackColor       =   &H00C0C0C0&
   Caption         =   "MPNS Store Data Maintenance"
   ClientHeight    =   6045
   ClientLeft      =   1050
   ClientTop       =   1950
   ClientWidth     =   8865
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
   ScaleHeight     =   6045
   ScaleWidth      =   8865
   Begin VB.Data datDropDown 
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   6390
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "ssfactor.sp_table"
      Top             =   4050
      Visible         =   0   'False
      Width           =   1695
   End
   Begin FACTFRMLib.FactorFrame efraBackground 
      Height          =   5205
      Left            =   0
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   480
      Width           =   8850
      _Version        =   65536
      _ExtentX        =   15610
      _ExtentY        =   9181
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
      PicturePos      =   0
      TitleBarHeight  =   24
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
         Height          =   390
         HelpContextID   =   10
         Left            =   60
         TabIndex        =   14
         Top             =   4710
         Width           =   1305
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
         TitleBarHeight  =   24
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
         Height          =   390
         HelpContextID   =   11
         Left            =   1530
         TabIndex        =   15
         Top             =   4710
         Width           =   1305
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
         TitleBarHeight  =   24
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
         Height          =   390
         HelpContextID   =   12
         Left            =   3000
         TabIndex        =   16
         Top             =   4710
         Width           =   1305
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
         Enabled         =   0   'False
         Caption         =   "&Delete"
         CaptionPos      =   4
         PicturePos      =   3
         ShowFocusRect   =   -1  'True
         Style           =   3
         BorderWidth     =   4
         TitleBarHeight  =   24
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
         Height          =   390
         HelpContextID   =   13
         Left            =   4470
         TabIndex        =   17
         Top             =   4710
         Width           =   1305
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
         TitleBarHeight  =   24
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
         Height          =   390
         HelpContextID   =   15
         Left            =   7410
         TabIndex        =   19
         Top             =   4710
         Width           =   1350
         _Version        =   65536
         _ExtentX        =   2381
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
         TitleBarHeight  =   24
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
         Height          =   390
         HelpContextID   =   14
         Left            =   5940
         TabIndex        =   18
         Top             =   4710
         Width           =   1305
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
         TitleBarHeight  =   24
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
      Begin FACTFRMLib.FactorFrame efraInnerFrame1 
         Height          =   585
         Left            =   60
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   990
         Width           =   8715
         _Version        =   65536
         _ExtentX        =   15372
         _ExtentY        =   1032
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
         TitleBarHeight  =   24
         BeginProperty PanelFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            = 9.75  
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin VB.TextBox txtSeries 
            Height          =   396
            HelpContextID   =   1802
            Left            =   1710
            MaxLength       =   5
            TabIndex        =   9
            Top             =   60
            Width           =   1488
         End
         Begin FACTFRMLib.FactorFrame cmdSeries 
            Height          =   396
            HelpContextID   =   1803
            Left            =   3225
            TabIndex        =   11
            TabStop         =   0   'False
            Top             =   60
            Width           =   390
            _Version        =   65536
            _ExtentX        =   688
            _ExtentY        =   698
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
            Enabled         =   0   'False
            CaptionPos      =   4
            Picture         =   "zzmpnsmt.frx":0000
            Style           =   3
            BorderWidth     =   4
            TitleBarHeight  =   24
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
         Begin VB.Label lblSeries 
            Caption         =   "Series:"
            Height          =   255
            Left            =   840
            TabIndex        =   20
            Top             =   150
            Width           =   855
         End
      End
      Begin FACTFRMLib.FactorFrame efraOuterFrame 
         Height          =   2985
         Left            =   60
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   1650
         Width           =   8715
         _Version        =   65536
         _ExtentX        =   15372
         _ExtentY        =   5265
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
         TitleBarHeight  =   24
         BeginProperty PanelFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            = 9.75  
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin DBTrueGrid.TDBGrid tgTable 
            Height          =   2835
            HelpContextID   =   120
            Left            =   0
            OleObjectBlob   =   "zzmpnsmt.frx":0112
            TabIndex        =   13
            Top             =   90
            Width           =   8640
         End
      End
      Begin FACTFRMLib.FactorFrame efraInnerFrame 
         Height          =   855
         Left            =   60
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   60
         Width           =   8715
         _Version        =   65536
         _ExtentX        =   15372
         _ExtentY        =   1508
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
         TitleBarHeight  =   24
         BeginProperty PanelFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            = 9.75  
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin VB.OptionButton optRptBtn 
            Caption         =   "South"
            Height          =   240
            HelpContextID   =   103
            Index           =   4
            Left            =   6270
            TabIndex        =   7
            Top             =   330
            Width           =   1245
         End
         Begin VB.OptionButton optRptBtn 
            Caption         =   "Moviequik"
            Height          =   240
            HelpContextID   =   100
            Index           =   1
            Left            =   810
            TabIndex        =   1
            Top             =   330
            Value           =   -1  'True
            Width           =   1395
         End
         Begin VB.OptionButton optRptBtn 
            Caption         =   "Payphone"
            Height          =   240
            HelpContextID   =   101
            Index           =   2
            Left            =   2610
            TabIndex        =   3
            Top             =   330
            Width           =   1410
         End
         Begin VB.OptionButton optRptBtn 
            Caption         =   "North"
            Height          =   240
            HelpContextID   =   102
            Index           =   3
            Left            =   4620
            TabIndex        =   5
            Top             =   330
            Width           =   1320
         End
      End
   End
   Begin FACTFRMLib.FactorFrame ffraStatusbar 
      Height          =   360
      Left            =   0
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   5685
      Width           =   8865
      _Version        =   65536
      _ExtentX        =   15637
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
      CaptionPos      =   1
      Style           =   5
      TitleBarHeight  =   24
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
   Begin FACTFRMLib.FactorFrame efraToolBar 
      Height          =   465
      Left            =   0
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   0
      Width           =   8865
      _Version        =   65536
      _ExtentX        =   15637
      _ExtentY        =   820
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
      FMName          =   "ZZMPNSMT"
      CaptionPos      =   4
      Style           =   6
      TitleBarHeight  =   24
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
         Height          =   372
         Left            =   60
         TabIndex        =   0
         Top             =   84
         Width           =   6876
         _ExtentX        =   12118
         _ExtentY        =   635
         ButtonWidth     =   635
         ButtonHeight    =   582
         _Version        =   393216
      End
   End
   Begin DBTrueGrid.TDBGrid tblComboDropdown 
      Bindings        =   "zzmpnsmt.frx":1745
      Height          =   2484
      Left            =   0
      OleObjectBlob   =   "zzmpnsmt.frx":175F
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   0
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
Attribute VB_Name = "frmZZMPNSMT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'=======================================================================
'
'PROGRAMER:         HEHUA(SAM) ZHENG (hzheng2000@hotmail.com)
'
'========================================================================



Option Explicit

Private t_bStartupFlag As Boolean 'optional startup flag
Private t_bDataChanged As Boolean 'data changed flag
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
'

'************************************************
Private Const t_totalOptionBtn As Integer = 4
Private Const COL_PERIOD = 0
Private Const COL_NBR = 1
Private Const COL_YTD = 2

Private Const C_M = 1
Private Const C_P = 2
Private Const C_N = 3
Private Const C_S = 4

Private nDataStatus As Integer
Private Const DATA_INITIAL = 0
Private Const DATA_LOADING = 1
Private Const DATA_LOADED = 2
Private Const DATA_CHANGED = 3

Private t_nCurrentOptBtn As Integer
Private t_nCurrentRpt As String

Public dbLocal As DataBase
Public tgmDetail As clsTGSpreadSheet
Public cValidate As cValidateInput

Private bInHere As Boolean
    
Public tgcDropdown As Object
Private Const t_szoptRptBtnGotFocus As String = "Select a report type for Add or Edit, then press Enter key"
Private Const szLongPattern As String = "^(#{0,9}|[0-1]#{0,9}|\2(\0#{0,8}|\1([0-3]#{0,7}|\4([0-6]#{0,6}|\7([0-3]#{0,5}|\4([0-7]#{0,4}|\8([0-2]#{0,3}|\3([0-5]#?#?|\6([0-3]#?|\4[0-7]?)?)?)?)?)?)?)?)?)$"



Private Sub cmdSeries_Click()
    Screen.MousePointer = vbHourglass
    tgcDropdown.Click cmdSeries
    Screen.MousePointer = vbDefault
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
    App.HelpFile = "711MPNS.HLP"
End Sub

Private Sub Form_Unload(CANCEL As Integer)
    tfnUnlockRow
    
    On Error Resume Next
    
    Set tgcDropdown = Nothing
    'Set cValidate = Nothing
    'Set objCursor = Nothing
    Set objErrHandler = Nothing
    
    Set objCurrTabControl = Nothing
    
    Unload frmContext
    Unload frmAbout
    
    'project local database object variables
  '  If Not dbLocal Is Nothing Then
  '      dbLocal.Close
  '  End If
   ' Set dbLocal = Nothing
    
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

    #If FACTOR_MENU = 0 Then
        't_szConnect = "ODBC;DSN=Gasup;DB=/factor/gasup/factor;HOST=ether;SERV=sqlexec;SRVR=ether;PRO=sesoctcp;UID=ma;PWD=menus"
        t_szConnect = "ODBC;DSN=dktest;DB=/factor/dktest/factor;HOST=ether;SERV=sqlexec;SRVR=ether;PRO=sesoctcp;UID=hzheng;PWD=shangyu"
    #End If
    If tfnAuthorizeExecute(Command) = False Then 'Check for handshake if not in the development mode
        Unload Me
        Exit Sub
    End If
     
    If tfnOpenDatabase = False Then 'open the database, ODBC Dialog Box during developemnt, oleObject Connection String when not
        Unload Me
        Exit Sub
    End If
  
    'connect to local database
    #If FACTOR_MENU <> 1 Then
        Set dbLocal = tfnOpenLocalDatabase()
        If dbLocal Is Nothing Then
            MsgBox "System Error: Unable to open Local Database, Program will be terminated", vbExclamation
            Unload Me
            Exit Sub
        End If
    #End If
   
    subInitErrorHandler   ' Setup Error Control
     
    tfnUpdateVersion

    ' disable system menu Close and Close icon on form
    tfnDisableFormSystemClose Me
    
    subSetupToolBar
    
    tmrKeyBoard.Enabled = False
    tfnCenterForm Me
    Me.Show
    Screen.MousePointer = vbHourglass
    tfnSetInitializingMessage
    Me.Enabled = False
    DoEvents

    '***************************************************
    ' INSERT YOUR FORM LOAD CODE HERE
    ' | | | | | | |
    ' v v v v v v v
    
    Screen.MousePointer = vbHourglass
    
    subSetupCombos
    subSetupTDTable
    subInitValidation  'set up validation object
    Screen.MousePointer = vbHourglass
    
    ' ^ ^ ^ ^ ^ ^ ^
    ' | | | | | | |
    '***************************************************
    
    Me.Enabled = True
    
    '***************************************************
    ' SET YOUR FIRST STATUSBAR MESSAGE HERE
    ' | | | | | | |
    ' v v v v v v v
        
    ' ^ ^ ^ ^ ^ ^ ^
    ' | | | | | | |
    '***************************************************
    t_bStartupFlag = False
    tmrKeyBoard.Enabled = True
    tfnResetScreen 'set the default screen

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
    tfnSetStatusBarMessage t_szADD
End Sub

Private Sub cmdEditBtn_GotFocus()
    tfnSetStatusBarMessage t_szEDIT
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
     t_nFormMode = ADD_MODE
     subAfterAddEditClick

End Sub

Private Sub cmdEditBtn_Click()
    
    t_nFormMode = EDIT_MODE
    subAfterAddEditClick

End Sub

Private Sub subAfterAddEditClick()
    If Not tfnLockRow("ZZMPNSMT", "gf_store_data", "gf_store_data") Then
        Exit Sub
    End If
    
    tgTable.Columns(COL_YTD).Locked = True
    If t_nFormMode = ADD_MODE Then
        cmdUpdateInsertBtn.Caption = t_szCAPTION_INSERT
        mnuUpdateInsert.Caption = t_szCAPTION_INSERT
        tgmDetail.AllowAddNew = True
        tgTable.Columns(COL_PERIOD).Locked = False
    End If
    
    If t_nFormMode = EDIT_MODE Then
        cmdUpdateInsertBtn.Caption = t_szCAPTION_UPDATE
        mnuUpdateInsert.Caption = t_szCAPTION_UPDATE
        tgmDetail.AllowAddNew = False
        tgTable.Columns(COL_PERIOD).Locked = True
    End If
       
    cmdRefreshSelectBtn.Caption = t_szCAPTION_REFRESH
    mnuRefreshSelect.Caption = t_szCAPTION_REFRESH
    subSetExitCancelBtn "CANCEL"

    subEnableAddBtn False
    subEnableEditBtn False

    subEnableOptionBtns t_nCurrentOptBtn, False
    
    tgcDropdown.SQL = "select unique gfsd_series from gf_store_data where gfsd_flag = " & tfnSQLString(t_nCurrentRpt)
    subEnableSeries True
    txtSeries.SetFocus

 
End Sub
Private Sub cmdDeleteBtn_Click()
    Dim lRow As Long
    Dim sMsg As String
    Dim sCode As String
    
    sMsg = "Are you sure you want to delete this record?"
    If tfnCancelExit(sMsg) Then
        subSetBusyState True
        lRow = tgmDetail.GetCurrentRowNumber
        sCode = tgmDetail.CellValue(COL_PERIOD, lRow)
        'If fnCodeNotExist(sCode) Then ' ok, can delete
           If fnDeleteAPeriod(sCode) Then
             '''Good Job!
             If tfnUnlockRow("gf_store_data") Then
             ''good job
             End If
           End If
        'Else
        '  tfnSetStatusBarError "The current code(row) cannot be deleted!"
        'End If
        subSetBusyState False
    End If

End Sub

Private Sub cmdUpdateInsertBtn_Click()
       
    Dim bRet As Boolean
    
    Screen.MousePointer = vbHourglass
    bRet = fnInsertUpdateData
    If bRet Then
        If tfnUnlockRow("gf_store_data") Then
             ''good job
        End If

        tfnResetScreen
    End If
    Screen.MousePointer = vbDefault
    
End Sub

Private Sub cmdExitCancelBtn_Click()
    If tfnUnlockRow("gf_store_data") Then
    ''good job
    End If

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
    
    nDataStatus = DATA_INITIAL
    subLoadScreen
    subEnableUpdateBtn False
    nDataStatus = DATA_LOADED

    subEnableRefreshBtn False
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
    If t_bDataChanged Then
        If Not tfnCancelExit(t_szCANCEL_MESSAGE) Then
            Exit Sub
        End If
    End If
    
    tfnUnlockRow
    
    t_nFormMode = IDLE_MODE
    t_bDataChanged = False
    
    tfnResetScreen 'reset all the buttons
End Sub

Private Sub cmdTBExit_Click()
    If t_bDataChanged Then
        If Not tfnCancelExit(t_szEXIT_MESSAGE) Then
            Exit Sub
        End If
    End If
    
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
        
    cmdUpdateInsertBtn.Caption = t_szCAPTION_UPDATE
    mnuUpdateInsert.Caption = t_szCAPTION_UPDATE
    
    cmdRefreshSelectBtn.Caption = t_szCAPTION_REFRESH
    mnuRefreshSelect.Caption = t_szCAPTION_REFRESH
    
 '   tfnSetStatusBarMessage t_szADD
    
    frmContext.ButtonEnabled(PRINT_UP) = False
    frmContext.ButtonEnabled(COPY_UP) = False
    
    mnuExit.Enabled = True
    mnuPrint.Enabled = False
    
    mnuCopy.Enabled = False
    mnuPaste.Enabled = False
    
    subEnableAddBtn True
    subEnableEditBtn True
    subEnableUpdateBtn False
    subEnableRefreshBtn False
    subSetExitCancelBtn "EXIT"

    txtSeries.Text = ""
    tgmDetail.ClearData
   'enable all opt btns,
    subEnableOptionBtns 0, True

    subEnableSeries False
    tgTable.Enabled = False
    t_nCurrentOptBtn = C_M
    
    
    nDataStatus = DATA_INITIAL
    cValidate.ResetFlags
    tgcDropdown.ResetFlags
    'set focus
    optRptBtn(t_nCurrentOptBtn).SetFocus
     tfnSetStatusBarMessage t_szoptRptBtnGotFocus
   
End Sub


Private Sub subEnableSeries(bFlag As Boolean)
   txtSeries.Enabled = bFlag
   cmdSeries.Enabled = bFlag
   subEnableSearchButton cmdSeries, bFlag
End Sub

Private Sub subEnableSearchButton(ByRef ctrlButton As Control, _
                                 ByVal bStatus As Boolean)
    'Eanble a search button

    ctrlButton.Enabled = bStatus
    If bStatus Then
        ctrlButton.Picture = frmContext.LoadPicture(SEARCH_UP)
    Else
        ctrlButton.Picture = frmContext.LoadPicture(SEARCH_DOWN)
    End If
End Sub

Private Sub optRptBtn_GotFocus(Index As Integer)
    t_nCurrentOptBtn = Index
    subGetRptFlag Index
    tfnSetStatusBarMessage t_szoptRptBtnGotFocus
 
End Sub

Private Sub optRptBtn_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        subSetFocus cmdAddBtn
        KeyAscii = 0
    End If
End Sub

Private Sub tblComboDropDown_Click()
'nm
    tgcDropdown.Click tblComboDropdown
End Sub

Private Sub tblComboDropDown_GotFocus()
'nm
    tgcDropdown.GotFocus tblComboDropdown
    subSetBusyState False
End Sub

Private Sub tblComboDropDown_KeyPress(KeyAscii As Integer)
'nm
    tgcDropdown.Keypress tblComboDropdown, KeyAscii
End Sub

Private Sub tblComboDropDown_LostFocus()
'nm
    tgcDropdown.LostFocus tblComboDropdown
End Sub

Private Sub tblComboDropDown_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
'nm
    tgcDropdown.TableMouseUp y
End Sub

Private Sub tblComboDropDown_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
'nm
    tgcDropdown.RowColChange
End Sub

Private Sub tblComboDropDown_SelChange(CANCEL As Integer)
'nm
    tgcDropdown.SelChange CANCEL
End Sub



Private Sub tgTable_AfterColEdit(ByVal ColIndex As Integer)
    
    Dim lRow As Long
    Dim ntemp As Long
    Dim i    As Long
    
    tgmDetail.AfterColEdit ColIndex
    
    lRow = tgmDetail.GetCurrentRowNumber
    If tgmDetail.CellValue(COL_NBR, lRow) <> "" Then
        ntemp = tgmDetail.CellValue(COL_NBR, 0)
        tgmDetail.CellValue(COL_YTD, 0) = ntemp
        For i = 1 To tgmDetail.RowCount - 1
            ntemp = ntemp + tgmDetail.CellValue(COL_NBR, i)
            tgmDetail.CellValue(COL_YTD, i) = ntemp
        Next i
        
        tgmDetail.Rebind
    End If
    subSetButtonStatus
End Sub

Private Sub tgTable_BeforeColEdit(ByVal ColIndex As Integer, ByVal KeyAscii As Integer, CANCEL As Integer)
    Dim lRow As Long
    
    tgmDetail.BeforeColEdit ColIndex, KeyAscii, CANCEL
    
     subEnableUpdateBtn False
          
End Sub

Private Sub tgTable_Change()
    tgmDetail.Change
 
End Sub


Private Sub tgTable_FirstRowChange()
    tgmDetail.FirstRowChange
    
End Sub

Private Sub tgTable_GotFocus()
    subEnableSeries False
    tgmDetail.GotFocus
    Screen.MousePointer = vbDefault
    
    subCheckDelete
End Sub

Private Sub tgTable_KeyDown(KeyCode As Integer, Shift As Integer)
    tgmDetail.KeyDown KeyCode, Shift
End Sub

Private Sub tgTable_KeyPress(KeyAscii As Integer)
 
   Dim lRow As Long
    
    lRow = tgmDetail.GetCurrentRowNumber
    
    If Not tgmDetail.Keypress(KeyAscii) Then
        KeyAscii = 0
     End If
    
   subSetButtonStatus
End Sub


Private Sub tgTable_LostFocus()
    
    tgmDetail.LostFocus
       
End Sub

Private Sub tgTable_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    Dim lRow As Long
    
    lRow = tgmDetail.GetCurrentRowNumber
    tgmDetail.RowColChange LastRow, LastCol
    
    If tgTable.col = COL_YTD Then
       tgTable.col = COL_PERIOD
       If t_nFormMode = EDIT_MODE Then
          If lRow < tgmDetail.RowCount - 1 Then
             tgTable.Row = lRow + 1
          End If
       End If
       If t_nFormMode = ADD_MODE Then
          If lRow < tgmDetail.RowCount Then
             tgTable.Row = lRow + 1
          End If
       End If
    End If
    If t_nFormMode = EDIT_MODE And tgTable.col = 0 Then
        tgTable.col = 1
    End If
    
    
    
    subCheckDelete
End Sub

Private Sub tgTable_SelChange(CANCEL As Integer)
    CANCEL = True
End Sub

Private Sub tgTable_UnboundReadData(ByVal RowBuf As DBTrueGrid.RowBuffer, StartLocation As Variant, ByVal ReadPriorRows As Boolean)
    tgmDetail.ReadData RowBuf, StartLocation, ReadPriorRows
End Sub

Private Sub tmrKeyboard_Timer() 'status bar timer - 250ms
    tfnUpdateStatusBar Me 'process the status bar
    'objCursor.SetMousePointer
End Sub

Private Sub subInitErrorHandler()
    If objErrHandler Is Nothing Then
        Set objErrHandler = New clsErrorHandler
        With objErrHandler
            Set .FormParent = Me
            Set .DatabaseEngine = t_engFactor
           ' Set .LocalDatabase = dbLocal
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
        .AddButton "Add Pro&fit Center", PRFTCNTR_UP
        .AddButton "Add &G/L Period/Series", GLPD_UP
        .AddButton "Add Custo&mer", CUSTOMER_UP
        .AddButton "Add &Product", PRODUCT_UP
        .AddButton "Add Ta&x Use Group", TAXUSE_UP
        .AddButton "Add G/L Accoun&t", GL_UP
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

'Click event subroutine for all search buttons
Private Sub subSearchButtonClick(cmdSearchButton As Control)
    subShowHourGlass
    tgcDropdown.Click cmdSearchButton
    subHideHourGlass
End Sub

Public Sub subShowHourGlass()
    tmrKeyBoard.Enabled = False
    Screen.MousePointer = vbHourglass
End Sub

Public Sub subHideHourGlass()
    Screen.MousePointer = vbDefault
    tmrKeyBoard.Enabled = True
End Sub

'************************************************************
'Now I start my code!



Private Sub subEnableOptionBtns(ByVal nIdx As Integer, ByVal bEnable As Boolean)

    Dim i As Integer
     
    efraInnerFrame.Enabled = bEnable
    
End Sub

Private Sub optRptBtn_Click(Index As Integer)
    
    t_nCurrentOptBtn = Index
    subGetRptFlag Index
    tfnSetStatusBarMessage t_szoptRptBtnGotFocus
End Sub


Private Sub subSetupCombos()

    'Setup the combos
    'Set the general properties of the combos
    Set tgcDropdown = CreateObject(t_szOLECOMBO)
    Set tgcDropdown.DBEngine = t_engFactor
    Set tgcDropdown.Form = Me
    Set tgcDropdown.DataBase = t_dbMainDatabase
    Set tgcDropdown.DataLink = datDropDown
    Set tgcDropdown.Table = tblComboDropdown

    'The Parameter Combo
    tgcDropdown.AddCombo
            
    tgcDropdown.AddComboBox txtSeries, cmdSeries, _
                "gfsd_series", tgcDropdown.SQL_STRING_TYPE(5)

End Sub


Private Sub subSetupTDTable()
    Dim i As Integer
    Dim nWidth As Integer
    
    nWidth = tgTable.Width
    
    With tgTable.Columns(COL_PERIOD)
        .DataField = ""
        .Width = 0.2 * nWidth
    End With
    With tgTable.Columns(COL_NBR)
        .DataField = ""
        .Width = 0.4 * nWidth
    End With
    With tgTable.Columns(COL_YTD)
        .DataField = ""
        .Width = 0.4 * nWidth
    End With

    Set tgmDetail = New clsTGSpreadSheet
    With tgmDetail
        Set .Form = Me
        Set .StatusBar = ffraStatusbar
        Set .Table = tgTable
        
        .AddEditColumn COL_PERIOD, "Enter a period", szLongPattern
        .AddEditColumn COL_NBR, "Enter the number of stores", szLongPattern
        .AddEditColumn COL_YTD, "Year to date total number of stores (display only)", szLongPattern
        .ClearData
    End With
    
End Sub

Private Sub subSetFocus(cntlTemp As Control, ParamArray arryControls() As Variant)
'nm
    'Set focus to a textbox or a command button control

    Const nTrialNumber As Integer = 1
    Dim nCount As Integer

    nCount = 0
    On Error GoTo errSetFocus
    cntlTemp.SetFocus
    Exit Sub
tryNext:
    On Error GoTo errNext
    arryControls(nCount).SetFocus
extSetFocus:
    On Error GoTo 0
    Exit Sub
    
errSetFocus:
    If nCount < nTrialNumber Then
        nCount = nCount + 1
        DoEvents
        Resume
    Else
        nCount = 0
        Resume tryNext
    End If
errNext:
    If nCount < UBound(arryControls) Then
        nCount = nCount + 1
        Resume tryNext
    Else
        Resume extSetFocus
    End If
End Sub

Private Sub subInitValidation()
    Set cValidate = New cValidateInput
    With cValidate
         Set .Form = Me
         Set .StatusBar = ffraStatusbar
    
         .AddEditBox txtSeries, "Enter series here"
         .MinTabIndex = tbToolbar.TabIndex
         .MaxTabIndex = tgTable.TabIndex
        .SetFirstControls cmdUpdateInsertBtn, cmdExitCancelBtn
        .ESCControl = tgTable
        .ESCControl = cmdExitCancelBtn
      End With
 End Sub

Public Function fnInvalidData(txtBox As Textbox) As Boolean
    Select Case txtBox.TabIndex
        Case txtSeries.TabIndex
            fnInvalidData = Not fnValidSeries(txtBox)
    End Select
End Function

Private Function fnValidSeries(txtBox As Textbox) As Boolean
    If Trim(txtBox.Text) = "" Then
       fnValidSeries = False
    Else
       fnValidSeries = True
    End If
End Function
Public Function fnValidCellValue(tgTDTable As TDBGrid, ByVal nCol As Integer, ByVal nRow As Long, sText As String) As Boolean
    Dim strSQL As String
    Dim rsTemp As Recordset
    Dim nRet As Integer
    Dim i As Long
    Dim sLockedUser As String
    
    sText = Trim(sText)
 
    Select Case nCol
        Case COL_PERIOD
            fnValidCellValue = fnValidPeriod(sText, nRow)
        Case COL_NBR
            fnValidCellValue = fnValidNumber(sText, nRow)
        Case Else
            fnValidCellValue = True
    End Select
    
End Function

Private Function fnValidPeriod(ByVal sText As String, ByVal nRow As Long)
    Dim nRowCount As Long
    Dim i As Long
    
    fnValidPeriod = False
    
    '1.check  empty
    If sText = "" Then
        tgmDetail.ErrorMessage(COL_PERIOD) = "The period is required"
        Exit Function
    End If
    
    '2. check duplicated in the filled table on screen
    nRowCount = tgmDetail.RowCount
    For i = 0 To nRowCount - 1
        If Trim(tgmDetail.CellValue(COL_PERIOD, i)) = sText And i <> nRow Then
            tgmDetail.ErrorMessage(COL_PERIOD) = "The duplicated period encounted"
            Exit Function
        End If
    Next i
    
    '3. check already in use in the db table
    If fnPeriodInUse(sText) Then
        tgmDetail.ErrorMessage(COL_PERIOD) = "The entered period is already in use"
        Exit Function
    End If
               
    fnValidPeriod = True
End Function

Private Function fnPeriodInUse(ByVal sText As String) As Boolean
    Const SUB_NAME = "fnPeriodInUse"
    Dim strSQL As String
    Dim rsTemp As Recordset

    fnPeriodInUse = False

    strSQL = "SELECT gfsd_period FROM gf_store_data " _
            & " WHERE gfsd_flag =" & tfnSQLString(t_nCurrentRpt) _
            & " and gfsd_series =" & tfnSQLString(txtSeries.Text) _
            & " and gfsd_period =" & tfnSQLString(sText)
    Set rsTemp = fnOpenRecord(strSQL, SUB_NAME, "")
    If Not rsTemp Is Nothing Then
        If rsTemp.RecordCount > 0 Then
            fnPeriodInUse = True
        End If
    End If
End Function

Private Function fnValidNumber(ByVal sText As String, ByVal nRow As Long)
     fnValidNumber = False
    
    '1.check  empty
    If sText = "" Then
        tgmDetail.ErrorMessage(COL_NBR) = "The number of stores is required"
        Exit Function
    End If
   
   '2
    If Mid(sText, 1, 1) = "-" Then
       tgmDetail.ErrorMessage(COL_NBR) = "The number of stores must be positive"
       Exit Function
    End If
    
    fnValidNumber = True
End Function

Private Sub subDataChanged()
'nm
    If nDataStatus = DATA_LOADED Then
        nDataStatus = DATA_CHANGED
        If t_nFormMode <> ADD_MODE Then
            subEnableRefreshBtn True
        End If
    End If
    subEnableUpdateBtn False
    
End Sub


Private Sub subSetButtonStatus()
'nm
    If fnDataChanged Then
        subEnableDeleteBtn False
        If t_nFormMode = EDIT_MODE Then
            subEnableRefreshBtn True
        Else
            subEnableRefreshBtn False
        End If
        Dim nCurrRow As Integer
        nCurrRow = tgmDetail.GetCurrentRowNumber
      
        
         If tgmDetail.ValidCell(COL_PERIOD, nCurrRow) And tgmDetail.ValidCell(COL_NBR, nCurrRow) Then
             subEnableUpdateBtn True
         Else
             subEnableUpdateBtn False
        End If
       
    End If
 
End Sub

'copied from Ma's
Public Function fnOpenRecord(strSQL As String, _
                              Optional vCaller As Variant, _
                              Optional vMsg As Variant, _
                              Optional vDB As Variant) As Recordset
    Const SUB_NAME = "fnOpenRecord"
    ' Get records from the given SQL statement
    Dim objDB As DataBase
    Dim rsTemp As Recordset

    If IsMissing(vDB) Then
        Set objDB = t_dbMainDatabase
    Else
        Set objDB = vDB
    End If
    On Error GoTo SQLError
    If objDB Is t_dbMainDatabase Then
        Set rsTemp = objDB.OpenRecordset(strSQL, dbOpenSnapshot, dbSQLPassThrough)
    Else
        Set rsTemp = objDB.OpenRecordset(strSQL, dbOpenSnapshot)
    End If
    If rsTemp.RecordCount > 0 Then
        rsTemp.MoveLast
        rsTemp.MoveFirst
    End If
    Set fnOpenRecord = rsTemp

    On Error GoTo 0
    Exit Function
SQLError:
    Set fnOpenRecord = Nothing
    Dim bShow As Boolean
    bShow = Not IsMissing(vMsg)
    If IsMissing(vCaller) Then
        tfnErrHandler SUB_NAME, strSQL, bShow
    Else
        tfnErrHandler SUB_NAME & "," & CStr(vCaller), strSQL, bShow
    End If
End Function

'Copied from Ma's
Public Function fnExecuteSQL(strSQL As String, _
                             Optional vCaller As Variant, _
                             Optional vMsg As Variant, _
                             Optional vDB As Variant) As Boolean
    Const SUB_NAME = "fnExecuteSQL"
    Dim objDB As DataBase
    
    If IsMissing(vDB) Then
        Set objDB = t_dbMainDatabase
    Else
        Set objDB = vDB
    End If
    On Error GoTo errExecute
    If objDB Is t_dbMainDatabase Then
        objDB.ExecuteSQL strSQL
    Else
        objDB.Execute strSQL
    End If
    fnExecuteSQL = True

    On Error GoTo 0
    Exit Function

errExecute:
    fnExecuteSQL = False
    Dim bShow As Boolean
    bShow = Not IsMissing(vMsg)
    If IsMissing(vCaller) Then
        tfnErrHandler SUB_NAME, strSQL, bShow
    Else
        tfnErrHandler SUB_NAME & "," & CStr(vCaller), strSQL, bShow
    End If
End Function


Private Function fnDataChanged() As Boolean

    If nDataStatus = DATA_CHANGED Or tgmDetail.GetChangedRowCount > 0 Then
        fnDataChanged = True
    Else
        fnDataChanged = False
    End If
    
End Function

Private Sub subCheckDelete()
'OK
    Dim bFlag As Boolean
    Dim lRow As Long
    
    bFlag = False
    If nDataStatus >= DATA_LOADED Then
        If tgTable.Row >= 0 Then
            lRow = tgmDetail.GetCurrentRowNumber
            If lRow < tgmDetail.RowCount Then
                If tgmDetail.CellValue(COL_PERIOD, lRow) <> "" Then
                        bFlag = True
                End If
            End If
        End If
    End If
    
    subEnableDeleteBtn bFlag
    
End Sub

Private Sub subLoadScreen()
'nm
    Const SUB_NAME = "subLoadScreen"

    'If nDataStatus <> DATA_INITIAL Then
    '    Exit Sub
    'End If
    subSetBusyState True
    nDataStatus = DATA_LOADING
    If fnLoadData Then
        nDataStatus = DATA_LOADED
        subEnableDeleteBtn True
        subSetFocus tgTable
    Else
        nDataStatus = DATA_INITIAL
        
    End If
    subSetBusyState False
End Sub

Public Sub subSetBusyState(bFlag As Boolean)
 'nm
    If bFlag Then
        Screen.MousePointer = vbHourglass
    Else
        Screen.MousePointer = vbDefault
    End If
End Sub


Private Function fnLoadData() As Boolean

    Const SUB_NAME = "fnLoadData"
    
    Dim strSQL As String
    Dim rsTemp As Recordset
    Dim vData() As Variant
    Dim lRow As Long
    Dim lytd As Long
    
    fnLoadData = False
    
    lytd = 0
    strSQL = "SELECT gfsd_period, gfsd_nbr from gf_store_data " _
             & " where gfsd_flag = " & tfnSQLString(t_nCurrentRpt) _
             & " and gfsd_series = " & tfnSQLString(txtSeries.Text) _
             & " order by gfsd_period "
    
    Set rsTemp = fnOpenRecord(strSQL, SUB_NAME, "")
    If Not rsTemp Is Nothing Then
        With rsTemp
             If .RecordCount > 0 Then
                ReDim vData(tgTable.Columns.Count - 1, .RecordCount - 1)
                lRow = 0
                While Not .EOF
                    vData(COL_PERIOD, lRow) = fnCStr(!gfsd_period)
                    vData(COL_NBR, lRow) = fnCStr(!gfsd_nbr)
                    lytd = lytd + !gfsd_nbr
                    vData(COL_YTD, lRow) = fnCStr(lytd)
                    
                    .MoveNext
                    lRow = lRow + 1
                Wend
            Else
               MsgBox " No data in the gf_store_data table", vbCritical + vbOKOnly
               Exit Function
            End If
        End With
        tgmDetail.FillWithArray vData
    End If
    
    fnLoadData = True
End Function

Public Function fnCStr(vTemp As Variant) As String
'nm
    If IsNull(vTemp) Then
        fnCStr = ""
    Else
        fnCStr = Trim(vTemp)
    End If
End Function

Private Function fnInsertUpdateData() As Boolean
   Const SUB_NAME = "fnInsertUpdateData"
    
    Dim strSQL As String
    Dim i As Long
    Dim nLine As Integer
    Dim vData() As Variant
    Dim nCols As Integer
    Dim bDone As Boolean
    
    fnInsertUpdateData = False
       
    For i = 0 To tgmDetail.RowCount - 1
        nLine = i + 1
        If Not fnNullValue(tgmDetail.CellValue(COL_PERIOD, i)) Then
            tgmDetail.GetRow vData, nCols, i
            If fnLineNotExist(vData) Then
               bDone = fnInsertLine(vData)
            Else
               bDone = fnUpdateLine(vData)
            End If
            
            If Not bDone Then
                Exit Function
            End If
        End If
    Next i
    
    fnInsertUpdateData = True
    
End Function

Public Function fnNullValue(vTemp As Variant) As Boolean
'nm
    If IsNull(vTemp) Then
        fnNullValue = True
    Else
        If Trim(vTemp) = "" Then
            fnNullValue = True
        Else
            fnNullValue = False
        End If
    End If
End Function

Private Function fnInsertLine(vData() As Variant) As Boolean

    Const SUB_NAME = "fnInsertLine"
    Dim strSQL As String
    
    fnInsertLine = False
    
    strSQL = "INSERT INTO gf_store_data(gfsd_flag, gfsd_series,gfsd_period,gfsd_nbr) VALUES(" _
            & tfnSQLString(t_nCurrentRpt) & "," & tfnSQLString(txtSeries.Text) _
            & "," & tfnSQLString(vData(COL_PERIOD)) _
            & "," & tfnSQLString(vData(COL_NBR)) & ")"
    'execute the SQL statement:
    If Not fnExecuteSQL(strSQL, SUB_NAME, "") Then
        MsgBox "Failed to insert records into the gf_store_data table", vbCritical + vbOKOnly
        Exit Function
    End If
    
    fnInsertLine = True
    
End Function

Private Function fnUpdateLine(vData() As Variant) As Boolean

    Const SUB_NAME = "fnUpdateLine"
    Dim strSQL As String
    
    fnUpdateLine = False
    
    strSQL = "UPDATE gf_store_data SET  gfsd_nbr=" & tfnSQLString(vData(COL_NBR)) _
             & " WHERE gfsd_flag=" & tfnSQLString(t_nCurrentRpt) _
             & " and gfsd_series = " & tfnSQLString(txtSeries.Text) _
             & " and gfsd_period = " & tfnSQLString(vData(COL_PERIOD))
 
    If Not fnExecuteSQL(strSQL, SUB_NAME, "") Then
        MsgBox "Failed to update records to the gf_store_data table", vbCritical + vbOKOnly
        Exit Function
    End If
    
    fnUpdateLine = True
    
End Function




Private Function fnDeleteAPeriod(ByVal sPeriod As String) As Boolean
    Const SUB_NAME = "fnDelectAPeriod"
    Dim strSQL As String
    
    fnDeleteAPeriod = False
    
    strSQL = "DELETE FROM gf_store_data WHERE gfsd_flag = " & tfnSQLString(t_nCurrentRpt) _
            & " and gfsd_series = " & tfnSQLString(txtSeries.Text) _
            & " and gfsd_period = " & tfnSQLString(sPeriod)
   
    If Not fnExecuteSQL(strSQL, SUB_NAME, "") Then
        MsgBox "Failed to delete records from the gf_store_data table", vbCritical + vbOKOnly
        Exit Function
    End If
    
    tgmDetail.DeleteRow
    fnDeleteAPeriod = True
   
End Function

Private Function fnLineNotExist(vData() As Variant) As Boolean
    Const SUB_NAME = "fnLineNotExist"
    Dim strSQL As String
    Dim rsTemp As Recordset

    fnLineNotExist = False

    strSQL = "SELECT * FROM gf_store_data " _
            & " WHERE gfsd_flag = " & tfnSQLString(t_nCurrentRpt) _
            & "  AND gfsd_series = " & tfnSQLString(txtSeries.Text) _
            & " AND gfsd_period = " & tfnSQLString(vData(COL_PERIOD))
 
    Set rsTemp = fnOpenRecord(strSQL, SUB_NAME, "")
    If Not rsTemp Is Nothing Then
        With rsTemp
            If .RecordCount <= 0 Then
               fnLineNotExist = True
            End If
        End With
    End If
End Function


Private Sub subGetRptFlag(Index As Integer)
   Select Case Index
        Case 1
            t_nCurrentRpt = "M"
        Case 2
            t_nCurrentRpt = "P"
        Case 3
            t_nCurrentRpt = "N"
        Case 4
            t_nCurrentRpt = "S"
    End Select
End Sub

Private Sub txtSeries_Change()

    tgcDropdown.Change txtSeries
    cValidate.Change txtSeries
    
    If txtSeries.Text = "" Then
       tgmDetail.ClearData
    End If
End Sub

Private Sub txtSeries_Click()
    tgcDropdown.Click txtSeries
End Sub

Private Sub txtSeries_GotFocus()
    tgcDropdown.GotFocus txtSeries
    cValidate.GotFocus txtSeries
    Screen.MousePointer = vbDefault
    
    If tgcDropdown.SingleRecordSelected Then
      If cValidate.ValidInput(txtSeries) Then
         subEnableTgTable True
         'If t_nFormMode = EDIT_MODE Then
            subLoadScreen
         'End If
      End If
    End If
End Sub

Private Sub txtSeries_KeyPress(KeyAscii As Integer)

    Dim bCode As Boolean
    
    If KeyAscii = vbKeyReturn Then
        Screen.MousePointer = vbHourglass
        If txtSeries.Text <> "" And fnSeriesIsNew() Then
           subEnableTgTable True
           Screen.MousePointer = vbDefault
           Exit Sub
        End If
    End If
    
    bCode = tgcDropdown.Keypress(txtSeries, KeyAscii)
    Screen.MousePointer = vbDefault

    If Not bCode Then
        If KeyAscii = vbKeyReturn Then
            If tgcDropdown.SingleRecordSelected Then
                If cValidate.ValidInput(txtSeries) Then
                   subEnableTgTable True
                   subLoadScreen
                End If
            End If
        End If
        KeyAscii = 0
    Else
        cValidate.Keypress txtSeries, KeyAscii
    End If
End Sub

Private Sub txtSeries_LostFocus()
   If cValidate.LostFocus(txtSeries, cmdSeries, tblComboDropdown) Then
        If cValidate.ValidInput(txtSeries) Then
            'do nothing
        End If
        subSetButtonStatus
        If txtSeries.Text <> "" Then
            subEnableTgTable True
        End If
        tgcDropdown.LostFocus txtSeries
    End If
End Sub

Private Sub subEnableTgTable(bFlag As Boolean)
    tgTable.Enabled = bFlag
    If bFlag Then
       tgTable.col = 0  'may change
       tgTable.SetFocus
    End If
End Sub

Private Function fnSeriesIsNew() As Boolean
    Const SUB_NAME = "fnSeriesIsNew"
    Dim strSQL As String
    Dim rsTemp As Recordset

    fnSeriesIsNew = False

    strSQL = "SELECT * FROM gf_store_data " _
            & " WHERE gfsd_flag = " & tfnSQLString(t_nCurrentRpt) _
            & "  AND gfsd_series = " & tfnSQLString(txtSeries.Text)
 
    Set rsTemp = fnOpenRecord(strSQL, SUB_NAME, "")
    If Not rsTemp Is Nothing Then
        With rsTemp
            If .RecordCount <= 0 Then
               fnSeriesIsNew = True
            End If
        End With
    End If
End Function

