VERSION 5.00
Object = "{C75015E0-2232-11D3-B440-0060971E99AF}#1.0#0"; "FACTFRM.OCX"
Object = "{3D388220-1F4E-11D3-B440-0060971E99AF}#1.0#0"; "FACTTAB.OCX"
Object = "{01028C21-0000-0000-0000-000000000046}#4.0#0"; "TG32OV.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmZZSEBCMT 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Commission Code File Maintenance"
   ClientHeight    =   6360
   ClientLeft      =   1056
   ClientTop       =   1956
   ClientWidth     =   8964
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
   ScaleHeight     =   6360
   ScaleWidth      =   8964
   Begin VB.Data datDropDown 
      Caption         =   "datDropDown"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   384
      Left            =   5148
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4440
      Visible         =   0   'False
      Width           =   1920
   End
   Begin FACTFRMLib.FactorFrame ffraStatusbar 
      Height          =   360
      Left            =   0
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   6000
      Width           =   8970
      _Version        =   65536
      _ExtentX        =   15822
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
      Left            =   24
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   492
      Width           =   8928
      _Version        =   65536
      _ExtentX        =   15748
      _ExtentY        =   9652
      _StockProps     =   77
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
      Begin FACTTABLib.FactorTab eTabMain 
         Height          =   5412
         Left            =   12
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   12
         Width           =   8916
         _Version        =   65536
         _ExtentX        =   15727
         _ExtentY        =   9546
         _StockProps     =   68
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabsPerPage     =   4
         Caption         =   "Co&mmission Code|Comm &Type / Freq|Commission &Grades"
         Begin FACTFRMLib.FactorFrame efraBaseGrades 
            Height          =   5085
            Left            =   18150
            TabIndex        =   29
            TabStop         =   0   'False
            Top             =   0
            Width           =   8910
            _Version        =   65536
            _ExtentX        =   15716
            _ExtentY        =   8969
            _StockProps     =   77
            BackColor       =   8388608
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9.6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   0
            BorderWidth     =   0
            BeginProperty PanelFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Begin FACTFRMLib.FactorFrame efraBaseIIGrades 
               Height          =   4440
               Left            =   96
               TabIndex        =   30
               TabStop         =   0   'False
               Top             =   84
               Width           =   8712
               _Version        =   65536
               _ExtentX        =   15367
               _ExtentY        =   7832
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
               Begin VB.PictureBox cmdDropdown 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00C0C0C0&
                  ForeColor       =   &H80000008&
                  Height          =   240
                  HelpContextID   =   22
                  Index           =   1
                  Left            =   0
                  ScaleHeight     =   216
                  ScaleWidth      =   228
                  TabIndex        =   34
                  TabStop         =   0   'False
                  Top             =   0
                  Visible         =   0   'False
                  Width           =   255
               End
               Begin DBTrueGrid.TDBGrid tblBonus 
                  Height          =   4320
                  HelpContextID   =   250
                  Index           =   1
                  Left            =   72
                  OleObjectBlob   =   "ZZSEBCMT.frx":0000
                  TabIndex        =   8
                  Top             =   60
                  Width           =   8580
               End
            End
            Begin FACTFRMLib.FactorFrame cmdAddBtn 
               Height          =   396
               HelpContextID   =   10
               Index           =   1
               Left            =   84
               TabIndex        =   9
               Top             =   4620
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
               Index           =   1
               Left            =   1560
               TabIndex        =   10
               Top             =   4620
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
               Index           =   1
               Left            =   3036
               TabIndex        =   14
               Top             =   4620
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
               Index           =   1
               Left            =   4524
               TabIndex        =   11
               Top             =   4620
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
               Index           =   1
               Left            =   6012
               TabIndex        =   12
               Top             =   4620
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
               Index           =   1
               Left            =   7500
               TabIndex        =   13
               Top             =   4620
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
         End
         Begin FACTFRMLib.FactorFrame efraBaseType 
            Height          =   5085
            Left            =   18075
            TabIndex        =   28
            TabStop         =   0   'False
            Top             =   0
            Width           =   8910
            _Version        =   65536
            _ExtentX        =   15716
            _ExtentY        =   8969
            _StockProps     =   77
            BackColor       =   8388608
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9.6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   0
            BorderWidth     =   0
            BeginProperty PanelFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Begin FACTFRMLib.FactorFrame efraBaseView 
               Height          =   4356
               Left            =   96
               TabIndex        =   32
               TabStop         =   0   'False
               Top             =   660
               Width           =   8712
               _Version        =   65536
               _ExtentX        =   15367
               _ExtentY        =   7683
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
               BeginProperty PanelFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   9.6
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Begin DBTrueGrid.TDBGrid tblView 
                  Height          =   4236
                  HelpContextID   =   227
                  Index           =   0
                  Left            =   60
                  OleObjectBlob   =   "ZZSEBCMT.frx":12DF
                  TabIndex        =   17
                  Top             =   60
                  Width           =   8592
               End
               Begin DBTrueGrid.TDBGrid tblView 
                  Height          =   4236
                  HelpContextID   =   228
                  Index           =   1
                  Left            =   60
                  OleObjectBlob   =   "ZZSEBCMT.frx":25BD
                  TabIndex        =   18
                  Top             =   60
                  Width           =   8592
               End
            End
            Begin FACTFRMLib.FactorFrame efraView 
               Height          =   480
               Left            =   96
               TabIndex        =   31
               TabStop         =   0   'False
               Top             =   84
               Width           =   8712
               _Version        =   65536
               _ExtentX        =   15367
               _ExtentY        =   847
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
               BeginProperty PanelFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   9.6
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Begin VB.OptionButton optView 
                  Caption         =   "Commission Types"
                  Height          =   272
                  HelpContextID   =   226
                  Index           =   1
                  Left            =   4956
                  TabIndex        =   16
                  Top             =   100
                  Width           =   2256
               End
               Begin VB.OptionButton optView 
                  Caption         =   "Commission Frequency"
                  Height          =   272
                  HelpContextID   =   225
                  Index           =   0
                  Left            =   936
                  TabIndex        =   15
                  Top             =   100
                  Value           =   -1  'True
                  Width           =   2500
               End
            End
         End
         Begin FACTFRMLib.FactorFrame efraBaseBCode 
            Height          =   5055
            Left            =   15
            TabIndex        =   26
            TabStop         =   0   'False
            Top             =   15
            Width           =   8880
            _Version        =   65536
            _ExtentX        =   15663
            _ExtentY        =   8916
            _StockProps     =   77
            BackColor       =   8388608
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9.6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   0
            BorderWidth     =   0
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
               Index           =   0
               Left            =   84
               TabIndex        =   2
               Top             =   4620
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
               Index           =   0
               Left            =   1560
               TabIndex        =   3
               Top             =   4620
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
               Index           =   0
               Left            =   3036
               TabIndex        =   7
               Top             =   4620
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
            Begin FACTFRMLib.FactorFrame cmdRefreshSelectBtn 
               Height          =   396
               HelpContextID   =   14
               Index           =   0
               Left            =   6012
               TabIndex        =   5
               Top             =   4620
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
               Index           =   0
               Left            =   7500
               TabIndex        =   6
               Top             =   4620
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
            Begin FACTFRMLib.FactorFrame efraBaseIICode 
               Height          =   4440
               Left            =   96
               TabIndex        =   27
               TabStop         =   0   'False
               Top             =   84
               Width           =   8712
               _Version        =   65536
               _ExtentX        =   15367
               _ExtentY        =   7832
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
               Begin VB.PictureBox cmdDropdown 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00C0C0C0&
                  ForeColor       =   &H80000008&
                  Height          =   240
                  HelpContextID   =   22
                  Index           =   0
                  Left            =   72
                  ScaleHeight     =   216
                  ScaleWidth      =   228
                  TabIndex        =   33
                  TabStop         =   0   'False
                  Top             =   60
                  Visible         =   0   'False
                  Width           =   255
               End
               Begin DBTrueGrid.TDBGrid tblBonus 
                  Height          =   4320
                  HelpContextID   =   200
                  Index           =   0
                  Left            =   72
                  OleObjectBlob   =   "ZZSEBCMT.frx":389B
                  TabIndex        =   1
                  Top             =   60
                  Width           =   8580
               End
            End
            Begin FACTFRMLib.FactorFrame cmdUpdateInsertBtn 
               Height          =   396
               HelpContextID   =   13
               Index           =   0
               Left            =   4524
               TabIndex        =   4
               Top             =   4620
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
         End
      End
   End
   Begin FACTFRMLib.FactorFrame efraToolBar 
      Height          =   468
      Left            =   0
      TabIndex        =   21
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
      FMName          =   "ZZSEBCMT"
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
         TabIndex        =   0
         Top             =   84
         Width           =   3396
         _ExtentX        =   5990
         _ExtentY        =   656
         ButtonWidth     =   609
         ButtonHeight    =   582
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
      Bindings        =   "ZZSEBCMT.frx":4B7A
      Height          =   2484
      Left            =   36
      OleObjectBlob   =   "ZZSEBCMT.frx":4B99
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   660
      Width           =   3756
   End
   Begin MSComctlLib.ProgressBar pbLock 
      Height          =   264
      Left            =   7044
      TabIndex        =   23
      Top             =   6036
      Visible         =   0   'False
      Width           =   1836
      _ExtentX        =   3239
      _ExtentY        =   466
      _Version        =   393216
      Appearance      =   1
   End
   Begin DBTrueGrid.TDBGrid tblDropDown 
      Bindings        =   "ZZSEBCMT.frx":5E7D
      Height          =   1008
      Left            =   324
      OleObjectBlob   =   "ZZSEBCMT.frx":5E97
      TabIndex        =   24
      Top             =   708
      Width           =   2604
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
Attribute VB_Name = "frmZZSEBCMT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***********************************************************
'
' Copyright (c) 2000 FACTOR, A Division of W.R.Hess Company
'
' Module name:   ZZSEBCMT
' Programmer :   Rajneesh Aggarwal
' Date       :   12/28/2000
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
Private t_bTabSwitch As Boolean

Private Const DATA_INIT = 0
Private Const DATA_LOADED = 1
Private Const DATA_CHANGED = 2

Private Const TabBonusCode As Integer = 0
Private Const TabBonusMaint As Integer = 1
Private Const TabBonusGrade As Integer = 2

Private nDataStatus As Integer
Private tgmBonus(1) As clsTGSpreadSheet
Private tgmView(1) As clsTGSpreadSheet
Private tgfDropdown(1) As clsFloatingDropDown

Private Const nColCode As Integer = 0
Private Const nColDesc As Integer = 1

Private Const nColBCGrade As Integer = 2
Private Const nColBCType As Integer = 3
Private Const nColBCFreq As Integer = 4
Private Const nColBCOverride As Integer = 5

Private Const nColBGLevel As Integer = 2

Private nxColBCDuplicates As Integer

Private nxColBGOldLevel As Integer
Private nxColBGDuplicates As Integer
Private nxColBGDescEditable As Integer

Private Const nFrequency As Integer = 0
Private Const nBonusCode As Integer = 0
Private Const nGrade As Integer = 1
Private Const nType As Integer = 1

Private szCommFreqPattern As String
'

Public Function fnValidCellValue(tgTable As TDBGrid, ByVal nCol As Integer, _
                                 ByVal lRow As Long, sText As String) As Boolean
    Dim strSQL As String
    Dim Index As Integer
    
    If eTabMain.CurrTab = TabBonusCode Then
        Index = nBonusCode
    ElseIf eTabMain.CurrTab = TabBonusGrade Then
        Index = nGrade
    End If
    
    If Index = nBonusCode Then
        Select Case nCol
        Case nColCode, nColDesc
            fnValidCellValue = fnValidBonusCode(sText, lRow, nCol)
        Case nColBCGrade
            fnValidCellValue = fnValidBonusGrade(sText, lRow)
        Case nColBCType
                fnValidCellValue = fnValidBonusType(sText, lRow)
        Case nColBCFreq
            fnValidCellValue = fnValidBonusFreq(sText, lRow)
        Case nColBCOverride
            If sText <> "" And InStr("YN", UCase(sText)) > 0 Then
                fnValidCellValue = True
            Else
                tgmBonus(nBonusCode).ErrorMessage(nColBCOverride) = "Override amount for pay code must be Y or N"
            End If
        End Select
    Else  'nGrade
        Select Case nCol
        Case nColCode, nColDesc
            fnValidCellValue = fnValidGradeCode(sText, lRow, nCol)
        Case nColBGLevel
            fnValidCellValue = fnValidGradeLevel(sText, lRow)
        End Select
    End If
End Function

Private Sub subSetupMousePointer()
    Set objCursor = New clsCursor
    Set objCursor.MainForm = Me
    
    objCursor.AddControls tblComboDropdown, frmAbout
    
    objCursor.AddScreen
    
    'objCursor.AddArea efraInnerFrame
    'objCursor.AddArea efraOuterFrame
    'objCursor.AddArea txtUserId
    'objCursor.SetLeftTop txtUserId
    'objCursor.SetRightBottom cmdUserIdSearchBtn
End Sub

Private Sub cmdDeleteBtn_Click(Index As Integer)
    Dim sCode As String
    
    'Just a check
    If tgmBonus(Index).RowCount < 1 Then
        Exit Sub
    End If
    
    sCode = tgmBonus(Index).CellValue(nColCode, tgmBonus(Index).GetCurrentRowNumber)
    
    If t_nFormMode = EDIT_MODE Then
        If Not fnCheckDelete(sCode, Index) Then
            Exit Sub
        End If
    End If
        
    If Not tfnCancelExit("Are you sure you want to delete the current record?") Then
        Exit Sub
    End If
    
    If t_nFormMode = EDIT_MODE Then
        If Not fnDeleteCode(sCode, Index) Then
            MsgBox "Failed to delete the record", vbCritical
            Exit Sub
        End If
    End If
    
    tgmBonus(Index).DeleteRow
    
    If t_nFormMode = EDIT_MODE And tgmBonus(Index).RowCount = 0 Then
        cmdTBCancel_Click Index
    End If
    
End Sub

Private Sub cmdDropDown_Click(Index As Integer)
'    Dim Index As Integer
'    If eTabMain.CurrTab = TabBonusCode Then
'        Index = nBonusCode
'    ElseIf eTabMain.CurrTab = TabBonusGrade Then
'        Index = nGrade
'    End If
    subSetFloatingSQL Index
    tgfDropdown(Index).ButtonClick
End Sub

Private Sub cmdDropDown_GotFocus(Index As Integer)
    tgfDropdown(Index).GotFocus cmdDropdown
End Sub

Private Sub cmdDropDown_LostFocus(Index As Integer)
    tgfDropdown(Index).LostFocus cmdDropdown
End Sub

Private Sub efraBaseIICode_GotFocus()
    subSetFocus tblBonus(nBonusCode)
End Sub

Private Sub efraBaseIIGrades_GotFocus()
    subSetFocus tblBonus(nGrade)
End Sub

Private Sub eTabMain_Click()
    Dim i As Integer

    If eTabMain.CurrTab = TabBonusCode Then
        If cmdAddBtn(nBonusCode).Enabled Then
            subSetFocus cmdAddBtn(nBonusCode)
        Else
            subSetFocus efraBaseIICode
        End If
    ElseIf eTabMain.CurrTab = TabBonusMaint Then
        For i = 0 To 1
            If optView(i).Value Then
                subSetFocus optView(i)
            End If
        Next
    Else
        If cmdAddBtn(nGrade).Enabled Then
            subSetFocus cmdAddBtn(nGrade)
        Else
            subSetFocus efraBaseIIGrades
        End If
    End If
    
    If t_bTabSwitch Then
        t_bTabSwitch = False
    End If

End Sub

Private Sub eTabMain_Switch(OldTab As Integer, NewTab As Integer, CANCEL As Integer)
    If Not eTabMain.TabEnabled(NewTab) Then
        CANCEL = True
        Exit Sub
    End If
    
    If OldTab = TabBonusCode And NewTab = TabBonusMaint Then
        If tblBonus(nBonusCode).col = nColBCType Then
            optView(1).Value = True
        ElseIf tblBonus(nBonusCode).col = nColBCFreq Then
            optView(0).Value = True
        End If
    End If
    
    If Not CANCEL Then t_bTabSwitch = True
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
    
    Set tgfDropdown(0) = Nothing
    Set tgfDropdown(1) = Nothing
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
        
        If Not fnCreateTempP_ALTNAME("pn_vendor", "pn_name") Then
            MsgBox "failed to create vendor temp table", vbCritical
            End
        End If
        
        If Not fnCreateTbl_tmpGradeLevel Then
            Unload Me
            Exit Sub
        End If
    #End If
    
    ' disable system menu Close and Close icon on form
    tfnDisableFormSystemClose Me
    subSetupToolBar
    
    tmrKeyBoard.Enabled = False
    tfnCenterForm Me
    
    subSetGridWidth tblBonus(nBonusCode)
    subSetGridWidth tblBonus(nGrade)
    subSetGridWidth tblView(nFrequency)
    subSetGridWidth tblView(nType)
    
    subMakeCommFrequencyPattern
    
    subInitSpreadsheets nBonusCode
    subInitSpreadsheets nGrade
    subInitViewGrid
    subSetFloatingDropDown nBonusCode
    subSetFloatingDropDown nGrade
    
    subLoadViewData
    
    Me.Show
    Screen.MousePointer = vbHourglass
    tfnSetInitializingMessage
    Me.Enabled = False
    
    subSetupMousePointer
    t_bStartupFlag = False
    
    Me.Enabled = True
    tfnResetScreen nBonusCode
    tfnResetScreen nGrade
    
    tmrKeyBoard.Enabled = True
    tfnSetStatusBarMessage t_szADDEDIT
    Screen.MousePointer = vbDefault
    
    #If PROTOTYPE Then
        frmContext.ButtonEnabled(VENDOR_UP) = False
    #End If
    
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
Private Sub cmdAddBtn_GotFocus(Index As Integer)
    tfnSetStatusBarMessage t_szADDEDIT
End Sub

Private Sub cmdEditBtn_GotFocus(Index As Integer)
    tfnSetStatusBarMessage t_szADDEDIT
End Sub

Private Sub cmdDeleteBtn_GotFocus(Index As Integer)
    tfnSetStatusBarMessage t_szDELETE
End Sub

Private Sub cmdUpdateInsertBtn_GotFocus(Index As Integer)
    If cmdUpdateInsertBtn(Index).Caption = t_szCAPTION_UPDATE Then
        tfnSetStatusBarMessage t_szUPDATE
    Else
        tfnSetStatusBarMessage t_szINSERT
    End If
End Sub

Private Sub cmdRefreshSelectBtn_GotFocus(Index As Integer)
    If cmdRefreshSelectBtn(Index).Caption = t_szCAPTION_REFRESH Then
        tfnSetStatusBarMessage t_szREFRESH
    Else
        tfnSetStatusBarMessage t_szSELECT
    End If
End Sub

Private Sub cmdExitCancelBtn_GotFocus(Index As Integer)
    If cmdExitCancelBtn(Index).Caption = t_szCAPTION_EXIT Then
        tfnSetStatusBarMessage t_szEXIT
    Else
        tfnSetStatusBarMessage t_szCANCEL
    End If
End Sub

Private Sub cmdAddBtn_Click(Index As Integer)
    
    #If PROTOTYPE Then
        Exit Sub
    #End If
    
    t_nFormMode = ADD_MODE
    t_bStartupFlag = False
    subEnableAddBtn False, Index
    subEnableEditBtn False, Index
    subEnableUpdateBtn False, Index
    subEnableRefreshBtn False, Index
    subEnableCancel True, Index
    If Index = nBonusCode Then
        eTabMain.TabEnabled(TabBonusGrade) = False
    ElseIf Index = nGrade Then
        eTabMain.TabEnabled(TabBonusCode) = False
    End If
    
'    subInitSpreadsheets Index
'    subSetFloatingDropDown Index
    With tgmBonus(Index)
        If Index = nBonusCode Then
            .AddEditColumn nColCode, "Enter Commission Code", "^P{1,4}$"
        Else
            .AddEditColumn nColCode, "Enter Grade Code", "^([A-Z])$"
        End If
    End With
    tgmBonus(Index).AllowAddNew = True
    
    tblBonus(Index).Enabled = True
    
    cmdUpdateInsertBtn(Index).Caption = t_szCAPTION_INSERT
    mnuUpdateInsert.Caption = t_szCAPTION_INSERT
    
    cmdRefreshSelectBtn(Index).Caption = t_szCAPTION_REFRESH
    mnuRefreshSelect.Caption = t_szCAPTION_REFRESH

    subSetFocus tblBonus(Index)
    
End Sub

Private Sub subSetFocus(v As Control)
    On Error Resume Next
    v.SetFocus
End Sub

Private Sub cmdEditBtn_Click(Index As Integer)
    #If PROTOTYPE Then
        Exit Sub
    #End If
    
    t_nFormMode = EDIT_MODE
    subSetFocus efraBackground
    
    subShowHourGlass
    subEnableAddBtn False, Index
    subEnableEditBtn False, Index
    subEnableUpdateBtn False, Index
    subEnableRefreshBtn False, Index
    subEnableCancel True, Index
    If Index = nBonusCode Then
        eTabMain.TabEnabled(TabBonusGrade) = False
    ElseIf Index = nGrade Then
        eTabMain.TabEnabled(TabBonusCode) = False
    End If
    
'    subInitSpreadsheets Index
'    subSetFloatingDropDown Index
    With tgmBonus(Index)
        If Index = nBonusCode Then
            .RemoveEditColumn nColCode
        Else
            .RemoveEditColumn nColCode
        End If
    End With
    tgmBonus(Index).AllowAddNew = False
    
    tblBonus(Index).Enabled = True
    cmdUpdateInsertBtn(Index).Caption = t_szCAPTION_UPDATE
    mnuUpdateInsert.Caption = t_szCAPTION_UPDATE
    
    cmdRefreshSelectBtn(Index).Caption = t_szCAPTION_REFRESH
    mnuRefreshSelect.Caption = t_szCAPTION_REFRESH
    
    subSetFocus tblBonus(Index)
    If Not fnLoadData(Index) Then
        'cmdTBCancel_Click Index
        tblBonus(Index).Enabled = False
    End If
    
    subHideHourGlass
    nDataStatus = DATA_LOADED

End Sub

Private Sub cmdUpdateInsertBtn_Click(Index As Integer)
    Dim i As Long
    
    If cmdUpdateInsertBtn(Index).Caption = t_szCAPTION_INSERT Then
        For i = 0 To tgmBonus(Index).RowCount - 1
            If Not fnInsertGrdOrBonusCode(i, Index) Then
                MsgBox "Failed to insert the record", vbCritical
                Exit Sub
            End If
        Next i
    Else
        If Index = nBonusCode Then
            For i = 0 To tgmBonus(Index).RowCount - 1
                If Not fnUpdateBonusCode(i) Then
                    MsgBox "Failed to update the record", vbCritical
                    Exit Sub
                End If
            Next i
        Else
            If Not fnUpdateBonusGrade() Then
                MsgBox "Failed to update the record", vbCritical
                Exit Sub
            End If
        End If
    End If
        
    tfnResetScreen Index
    
End Sub

Private Sub cmdExitCancelBtn_Click(Index As Integer)
    If cmdExitCancelBtn(Index).Caption = t_szCAPTION_CANCEL Then
        cmdTBCancel_Click Index
    Else
        cmdTBExit_Click
    End If
End Sub

Private Sub cmdRefreshSelectBtn_Click(Index As Integer)
    
    If Not tfnCancelExit(t_szREFRESH_MESSAGE) Then
        Exit Sub
    End If
    
    fnLoadData Index, True
    
    nDataStatus = DATA_INIT
    subEnableRefreshBtn False, Index
    subEnableUpdateBtn False, Index
    
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

Private Sub cmdTBCancel_Click(Index As Integer)
    
    If nDataStatus = DATA_CHANGED Then
        If Not tfnCancelExit(t_szCANCEL_MESSAGE) Then
            Exit Sub
        End If
    End If
    
    t_nFormMode = IDLE_MODE
    tfnResetScreen Index
    
End Sub

Private Sub cmdTBExit_Click()
    
    If nDataStatus = DATA_CHANGED Then
        If Not tfnCancelExit(t_szEXIT_MESSAGE) Then
            Exit Sub
        End If
    End If
    
    subShowHourGlass
    
    tfnUnlockRow
        
    subHideHourGlass
    
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
    cmdAddBtn_Click eTabMain.CurrTab / 2  'david 01/12/2001
End Sub

Private Sub mnuEdit_Click()
    cmdEditBtn_Click eTabMain.CurrTab / 2  'david 01/12/2001
End Sub

Private Sub mnuDelete_Click()
    cmdDeleteBtn_Click eTabMain.CurrTab / 2  'david 01/12/2001
End Sub

Private Sub mnuRefreshSelect_Click()
    cmdRefreshSelectBtn_Click eTabMain.CurrTab / 2  'david 01/12/2001
End Sub

Private Sub mnuUpdateInsert_Click()
    cmdUpdateInsertBtn_Click eTabMain.CurrTab / 2  'david 01/12/2001
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
    cmdTBCancel_Click eTabMain.CurrTab / 2  'david 01/12/2001
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
Private Sub tfnResetScreen(Index As Integer)
    
    On Error Resume Next
    
    subShowHourGlass
    
    subEnableAddBtn True, Index
    subEnableEditBtn True, Index
    subEnableUpdateBtn False, Index
    subEnableDeleteBtn False, Index
    subEnableRefreshBtn False, Index
    subEnableCancel False, Index
    eTabMain.TabEnabled(TabBonusCode) = True
    eTabMain.TabEnabled(TabBonusGrade) = True
    
    tblBonus(Index).col = nColCode
    tblBonus(Index).Enabled = False
    tfnUnlockRow
    
    cmdUpdateInsertBtn(Index).Caption = t_szCAPTION_UPDATE
    mnuUpdateInsert.Caption = t_szCAPTION_UPDATE
    cmdRefreshSelectBtn(Index).Caption = t_szCAPTION_REFRESH
    mnuRefreshSelect.Caption = t_szCAPTION_REFRESH
    frmContext.ButtonEnabled(PRINT_UP) = False
    frmContext.ButtonEnabled(COPY_UP) = False
    
    mnuExit.Enabled = True
    mnuPrint.Enabled = False
    mnuCopy.Enabled = False
    mnuPaste.Enabled = False
    
    tgmBonus(Index).ClearData
    tblBonus(Index).col = nColCode
    
    nDataStatus = DATA_INIT
    subSetFocus cmdAddBtn(Index)
    DoEvents
    
    subHideHourGlass
End Sub

Private Sub optView_Click(Index As Integer)
    Dim i As Integer
    
    For i = 0 To 1
        tblView(i).TabStop = False
    Next i
    
    tblView(Index).ZOrder vbBringToFront
    tblView(Index).TabStop = True
End Sub

Private Sub optView_GotFocus(Index As Integer)
    
    If Index = nFrequency Then
        tfnSetStatusBarMessage "View Commission Frequency Codes"
    Else
        tfnSetStatusBarMessage "View Commission Types"
    End If
    
End Sub

Private Sub tblBonus_HeadClick(Index As Integer, ByVal ColIndex As Integer)
    If Index = nGrade Then
        tgmBonus(Index).HeadClick ColIndex, nDataStatus = DATA_CHANGED
        subResetDescEditableFlag
    End If
End Sub

Private Sub tblDropDown_SelChange(CANCEL As Integer)
    CANCEL = True
End Sub

Private Sub tblView_SelChange(Index As Integer, CANCEL As Integer)
    CANCEL = True
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

Private Sub subEnableAddBtn(bOnOff As Boolean, Index As Integer)
    cmdAddBtn(Index).Enabled = bOnOff
    mnuAdd.Enabled = bOnOff
End Sub

Private Sub subEnableEditBtn(bOnOff As Boolean, Index As Integer)
    cmdEditBtn(Index).Enabled = bOnOff
    mnuEdit.Enabled = bOnOff
End Sub

Private Sub subEnableDeleteBtn(bOnOff As Boolean, Index As Integer)
    cmdDeleteBtn(Index).Enabled = bOnOff
    mnuDelete.Enabled = bOnOff
End Sub

Private Sub subEnableRefreshBtn(bOnOff As Boolean, Index As Integer)
    cmdRefreshSelectBtn(Index).Enabled = bOnOff
    mnuRefreshSelect.Enabled = bOnOff
End Sub

Private Sub subEnableUpdateBtn(bOnOff As Boolean, Index As Integer)
    cmdUpdateInsertBtn(Index).Enabled = bOnOff
    mnuUpdateInsert.Enabled = bOnOff
End Sub

Private Sub subEnableCancel(bStatus As Boolean, Index As Integer)
    mnuCancel.Enabled = bStatus
    frmContext.ButtonEnabled(CANCEL_UP) = bStatus
    If bStatus Then
        cmdExitCancelBtn(Index).Caption = t_szCAPTION_CANCEL
    Else
        cmdExitCancelBtn(Index).Caption = t_szCAPTION_EXIT
    End If
End Sub

Private Sub Form_Resize()
    frmContext.FormResize
End Sub

Private Sub subSetupToolBar()
    With frmContext
        .BeginSetupToolbar Me
        '.AddButton "Add &Vendor", VENDOR_UP
        .EndSetupToolbar
    
        .HelpFile = szHelpFileName
    End With
End Sub

Public Sub TBButtonCallBack(ByVal nID As Integer)
    Select Case nID
        Case CANCEL_UP
            cmdTBCancel_Click eTabMain.CurrTab / 2  'david 01/12/2001
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
    tmrKeyBoard.Enabled = True
    Screen.MousePointer = vbDefault
End Sub

Private Sub subSetGridWidth(tbl As TDBGrid)
    Dim myWidth As Variant
    Dim myField As Variant
    
    Dim i As Integer
    
    If tbl.Name = "tblBonus" And tbl.Index = nGrade Then
        myWidth = Array(0.25, 0.5, 0.25)
        myField = Array("bg_grade", "bg_grade_desc", "bg_emp_level")
    ElseIf tbl.Name = "tblBonus" And tbl.Index = nBonusCode Then
        myWidth = Array(0.12, 0.42, 0.12, 0.12, 0.12, 0.1)
        myField = Array("bc_bonus_code", "bc_code_desc", "bc_grade", "bc_type", "bc_frequency", "bc_override")
    Else ' tblView
        myWidth = Array(0.25, 0.75)
        If tbl.Index = nFrequency Then
            myField = Array("bfq_frequency", "bfq_freq_desc")
        ElseIf tbl.Index = nType Then
            myField = Array("bt_type", "bt_type_desc")
        End If
    End If
    
    While tbl.Columns.Count > 0
        tbl.Columns.Remove 0
    Wend
    
    tbl.ExtendRightColumn = True
    
    For i = 0 To UBound(myWidth)
        tbl.Columns.Add i
        With tbl.Columns(i)
            .Width = myWidth(i) * (tbl.Width - 300)
            .DataField = myField(i)
            .Visible = True
            .HeadAlignment = vbCenter
        End With
    Next
    
    If tbl.Name = "tblBonus" And tbl.Index = nGrade Then
        tbl.Caption = "Commission Grades"
        tbl.Columns(nColCode).Caption = "Grade Code"
        tbl.Columns(nColCode).Alignment = vbCenter
        tbl.Columns(nColDesc).Caption = "Grade Description"
        tbl.Columns(nColBGLevel).Caption = "Grade Level"
    ElseIf tbl.Name = "tblBonus" And tbl.Index = nBonusCode Then
        tbl.Caption = "Commission Codes"
        tbl.Columns(nColCode).Caption = "Comm. Code"
        tbl.Columns(nColDesc).Caption = "Commission Code Description"
        tbl.Columns(nColBCGrade).Caption = "Comm. Grade"
        tbl.Columns(nColBCGrade).Alignment = vbCenter
        tbl.Columns(nColBCType).Caption = "Comm. Type"
        tbl.Columns(nColBCType).Alignment = vbCenter
        tbl.Columns(nColBCFreq).Caption = "Comm. Frequency"
        tbl.Columns(nColBCFreq).Alignment = vbCenter
        tbl.Columns(nColBCOverride).Caption = "Override"
        tbl.Columns(nColBCOverride).Alignment = vbCenter
    Else ' tblView
        If tbl.Index = nFrequency Then
            tbl.Caption = "Commission Frequency"
            tbl.Columns(nColCode).Caption = "Commission Frequency"
            tbl.Columns(nColDesc).Caption = "Frequency Description"
        ElseIf tbl.Index = nType Then
            tbl.Caption = "Commission Type"
            tbl.Columns(nColCode).Caption = "Commission Type"
            tbl.Columns(nColDesc).Caption = "Type Description"
        End If
        tbl.Columns(nColCode).Alignment = vbCenter
    End If
    
End Sub

Private Sub subInitSpreadsheets(Index As Integer)
    
    If Not tgmBonus(Index) Is Nothing Then
        Set tgmBonus(Index) = Nothing
    End If
    
    Set tgmBonus(Index) = New clsTGSpreadSheet
    Set tgmBonus(Index).Table = tblBonus(Index)
    Set tgmBonus(Index).StatusBar = ffraStatusbar ' message bar name
    Set tgmBonus(Index).Form = Me
    Set tgmBonus(Index).engFactor = t_engFactor
    
    With tgmBonus(Index)
        If Index = nBonusCode Then
            .AddEditColumn nColCode, "Enter Commission Code", "^P{1,4}$"
            .AddEditColumn nColDesc, "Enter Code Description", "^P{1,40}$"
            .AddEditColumn nColBCGrade, "Enter Commission Grade", "^([A-Z])$"
            .AddEditColumn nColBCType, "Enter Commission Type", "^(([T])([123])([CEX]))$"
            .AddEditColumn nColBCFreq, "Enter Commission Frequency", szCommFreqPattern
            .AddEditColumn nColBCOverride, "Override the amount of this pay code in payroll (Y or N)", "^[YNyn]$"
            nxColBCDuplicates = .AddHiddenField("duplicates")
            .ColumnForNewRow = 0
            .RemoveEditColumn nColCode
        Else  'nGrade
            .AddEditColumn nColCode, "Enter Grade Code", "^([A-Z])$"
            .AddEditColumn nColDesc, "Enter Grade Description", "^P{1,20}$"
            .AddEditColumn nColBGLevel, "Enter Grade Level", szIntegerPattern
            nxColBGOldLevel = .AddHiddenField("old_level")
            nxColBGDuplicates = .AddHiddenField("duplicates")
            nxColBGDescEditable = .AddHiddenField("desc_editable")
            .ColumnForNewRow = 0
            .RemoveEditColumn nColCode
        
            .AddSortColumn nColCode, nColCode, .STRING_TYPE, .ASCENDING, .CASE_SENSITIVE
            
            .AddSortColumn nColBGLevel, nColBGLevel, .NUMERIC_TYPE, .ASCENDING, .CASE_SENSITIVE
                
            .SortByColumn = True
        End If
    End With
    'add by Ning
    With tblBonus(nGrade).Columns(nColBGLevel)
        .DataField = "bg_emp_level"
    End With
End Sub

Private Sub subSetFloatingDropDown(Index As Integer)

    If Not tgfDropdown(Index) Is Nothing Then
        Set tgfDropdown(Index) = Nothing
    End If

    Set tgfDropdown(Index) = New clsFloatingDropDown
    
    With tgfDropdown(Index)
        Set .DataBase = t_dbMainDatabase
        Set .SearchButton = cmdDropdown(Index)
        Set .MainTable = tblBonus(Index)
        Set .DropDownTable = tblDropDown
        Set .DataLink = datDropDown
        Set .Form = Me
        Set .EditClass = tgmBonus(Index)
        
        .SearchOnReturn = False
        .DefaultCursorControl = True
        
        If Index = nBonusCode Then
            .AddDropDown 1
            .AddColumn nColCode, "prpa_pay_code", .COLUMN_TYPE_STRING
            .AddExtraColumn "Pay Code Description", "prpa_desc", 3000
            .VisibleRows = 10
            
            .AddDropDown 2
            .AddColumn nColBCGrade, "bg_grade", .COLUMN_TYPE_STRING
            .AddExtraColumn "Description", "bg_grade_desc", 2000
            
            .AddDropDown 3
            .AddColumn nColBCType, "bt_type", .COLUMN_TYPE_STRING
            .AddExtraColumn "Description", "bt_type_desc", 5000
            .SetExtendLeft
            
            .AddDropDown 4
            .AddColumn nColBCFreq, "bfq_frequency", .COLUMN_TYPE_STRING
            .AddExtraColumn "Description", "bfq_freq_desc", 2100
            .SetExtendLeft
        ElseIf Index = nGrade Then
            .AddDropDown 1
            .AddColumn nColBGLevel, "grade_level", .COLUMN_TYPE_INTEGER
            .VisibleRows = 10
        End If
    End With
    
End Sub

Private Sub subSetFloatingSQL(Index As Integer)
    Dim sSql As String
    Dim sCheckDubInGrid As String
    
    
    If Index = nBonusCode Then
        sCheckDubInGrid = fnCheckDubInGrid(nColCode, tgmBonus(Index).GetCurrentRowNumber, Index)
        Select Case tblBonus(Index).col
            Case nColCode
                sSql = "SELECT prpa_pay_code, prpa_desc FROM pr_pay WHERE prpa_type = 'P'"
                sSql = sSql + " AND prpa_pay_code NOT IN (SELECT bc_bonus_code FROM"
                sSql = sSql + " bonus_codes)"
                If sCheckDubInGrid <> "" Then
                    sSql = sSql + " AND prpa_pay_code NOT IN (" + sCheckDubInGrid + ")"
                End If
                tgfDropdown(Index).SetSQL nColCode, sSql
            Case nColBCGrade
                sSql = "SELECT DISTINCT bg_grade, bg_grade_desc FROM bonus_grades"
                tgfDropdown(Index).SetSQL nColBCGrade, sSql
            Case nColBCType
                sSql = "SELECT bt_type, bt_type_desc FROM bonus_types"
                tgfDropdown(Index).SetSQL nColBCType, sSql
            Case nColBCFreq
                sSql = "SELECT bfq_frequency, bfq_freq_desc FROM bonus_frequency"
                tgfDropdown(Index).SetSQL nColBCFreq, sSql
        End Select
        
    ElseIf Index = nGrade Then
        sCheckDubInGrid = fnCheckDubInGrid(nColBGLevel, tgmBonus(Index).GetCurrentRowNumber, Index)
        sSql = "SELECT grade_level FROM tmpGradeLevel"
        If sCheckDubInGrid <> "" Then
            sSql = sSql & " WHERE grade_level NOT IN ( "
            sSql = sSql & sCheckDubInGrid & " )"
            If t_nFormMode = ADD_MODE Then
                sSql = sSql & " AND grade_level NOT IN ( SELECT bg_emp_level FROM bonus_grades ) "
            End If
        Else
            If t_nFormMode = ADD_MODE Then
                sSql = sSql & " WHERE grade_level NOT IN ( SELECT bg_emp_level FROM bonus_grades ) "
            End If
        End If
        tgfDropdown(Index).SetSQL nColBGLevel, sSql
    End If

End Sub

Private Sub tblDropDown_Click()
    Dim Index As Integer
    
    If eTabMain.CurrTab = TabBonusCode Then
        Index = nBonusCode
    ElseIf eTabMain.CurrTab = TabBonusGrade Then
        Index = nGrade
    End If
    
    tgfDropdown(Index).TableClick tblDropDown
    
End Sub

Private Sub tblDropDown_KeyPress(KeyAscii As Integer)
    Dim Index As Integer
    
    If eTabMain.CurrTab = TabBonusCode Then
        Index = nBonusCode
    ElseIf eTabMain.CurrTab = TabBonusGrade Then
        Index = nGrade
    End If
    
    tgfDropdown(Index).Keypress tblDropDown, KeyAscii
End Sub

Private Sub tblDropDown_LostFocus()
    Dim Index As Integer
    
    If eTabMain.CurrTab = TabBonusCode Then
        Index = nBonusCode
    ElseIf eTabMain.CurrTab = TabBonusGrade Then
        Index = nGrade
    End If
    
    tgfDropdown(Index).LostFocus tblDropDown
End Sub

Private Sub tblDropDown_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim Index As Integer
    
    If eTabMain.CurrTab = TabBonusCode Then
        Index = nBonusCode
    ElseIf eTabMain.CurrTab = TabBonusGrade Then
        Index = nGrade
    End If
    
    tgfDropdown(Index).MouseUp y
End Sub

Private Sub tblDropDown_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    Dim Index As Integer
    
    If eTabMain.CurrTab = TabBonusCode Then
        Index = nBonusCode
    ElseIf eTabMain.CurrTab = TabBonusGrade Then
        Index = nGrade
    End If
    
    tgfDropdown(Index).RowColChange tblDropDown
End Sub

Private Function fnValidBonusCode(szText As String, ByVal lRow As Long, _
                                       nCol As Integer) As Boolean
    Const SUB_NAME As String = "fnValidBonusCode"
    Dim strSQL As String
    Dim sCode As String
    Dim sCodeDesc As String
    Dim bFoundDuplicate  As Boolean
    Dim i As Long

    fnValidBonusCode = False
    
    If nCol = nColDesc Then
        If Trim(szText) = "" Then
            tgmBonus(nBonusCode).ErrorMessage(nCol) = "You must enter a description"
        Else
            fnValidBonusCode = True
        End If
        Exit Function
    End If
    
    'bonus code validation
    If t_nFormMode = EDIT_MODE Then
        fnValidBonusCode = True
        Exit Function
    End If
    
    If Trim(szText) = "" Then
        tgmBonus(nBonusCode).ErrorMessage(nColCode) = "You must enter a code"
        Exit Function
    End If
    
    'Check Dublication in the Table.
    strSQL = "SELECT * FROM bonus_codes WHERE bc_bonus_code = " & tfnSQLString(szText)
    If GetRecordCount(strSQL, , SUB_NAME) > 0 Then
        tgmBonus(nBonusCode).ErrorMessage(nColCode) = "Code already exists in the table"
        Exit Function
    End If
    
    'check for duplication
    bFoundDuplicate = False
    For i = 0 To tgmBonus(nBonusCode).RowCount - 1
        If i <> lRow Then
            If fnGetField(tgmBonus(nBonusCode).CellValue(nColCode, i)) = szText Then
                tgmBonus(nBonusCode).CellValue(nxColBCDuplicates, i) = "Y"
                tgmBonus(nBonusCode).ValidCell(nColCode, i) = False
                bFoundDuplicate = True
            Else
                If fnGetField(tgmBonus(nBonusCode).CellValue(nxColBCDuplicates, i)) = "Y" Then
                    If fnGetField(tgmBonus(nBonusCode).CellValue(nColCode, i)) = fnGetField(tgmBonus(nBonusCode).OldValue) Then
                        'reset the duplication flag
                        tgmBonus(nBonusCode).CellValue(nxColBCDuplicates, i) = ""
                        tgmBonus(nBonusCode).ValidCell(nColCode, i) = True
                    End If
                End If
            End If
        End If
    Next i
    
    If bFoundDuplicate Then
        tgmBonus(nBonusCode).CellValue(nxColBCDuplicates, lRow) = "Y"
        tgmBonus(nBonusCode).ErrorMessage(nColCode) = "Duplicate Bonus Code encountered"
        Exit Function
    End If
    
    If Not fnValidBonusPayCode(szText, lRow) Then
        Exit Function
    End If
    
    tgmBonus(nBonusCode).CellValue(nxColBCDuplicates, lRow) = ""
    
    If fnGetField(tgmBonus(nBonusCode).CellValue(nColBCOverride, lRow)) = "" Then
        tgmBonus(nBonusCode).CellValue(nColBCOverride, lRow) = "N"
    End If
    
    fnValidBonusCode = True

End Function

Private Function fnValidGradeCode(szText As String, ByVal lRow As Long, _
                                       nCol As Integer) As Boolean
    
    Const SUB_NAME As String = "fnValidGradeCode"
    Dim strSQL As String
    Dim rsTemp As Recordset
    Dim sCode As String
    Dim sCodeDesc As String
    Dim i As Long

    fnValidGradeCode = False
    
    If nCol = nColDesc Then
        If Trim(szText) = "" Then
            tgmBonus(nGrade).ErrorMessage(nCol) = "You must enter a description"
        Else
            fnValidGradeCode = True
        End If
    
        Exit Function
    End If
    
    'grade code validation
    If t_nFormMode = EDIT_MODE Then
        fnValidGradeCode = True
        Exit Function
    End If
    
    If Trim(szText) = "" Then
        tgmBonus(nGrade).ErrorMessage(nColCode) = "You must enter a code"
        Exit Function
    End If
    
    strSQL = "SELECT bg_grade_desc FROM bonus_grades WHERE bg_grade = " & tfnSQLString(szText)
    If GetRecordSet(rsTemp, strSQL, , SUB_NAME) > 0 Then
        tgmBonus(nGrade).CellValue(nColDesc, lRow) = fnGetField(rsTemp!bg_grade_desc)
        tgmBonus(nGrade).CellValue(nxColBGDescEditable, lRow) = "N"
        tblBonus(nGrade).col = nColDesc
    Else
        tgmBonus(nGrade).CellValue(nxColBGDescEditable, lRow) = "Y"
    End If
    
    'check the grid
    If tgmBonus(nGrade).CellValue(nxColBGDescEditable, lRow) = "Y" Then
        For i = 0 To lRow - 1
            If fnGetField(tgmBonus(nGrade).CellValue(nColCode, i)) = szText Then
                If tgmBonus(nGrade).ValidCell(nColDesc, i) Then
                    tgmBonus(nGrade).CellValue(nxColBGDescEditable, lRow) = "N"
                    tgmBonus(nGrade).CellValue(nColDesc, lRow) = fnGetField(tgmBonus(nGrade).CellValue(nColDesc, i))
                    tblBonus(nGrade).col = nColDesc
                    Exit For
                End If
            End If
        Next i
    End If
    
    fnValidGradeCode = True

End Function

Private Function fnDeleteCode(sCode As String, Index As Integer) As Boolean
    Const SUB_NAME As String = "fnDeleteCode"
    Dim strSQL As String
    
    fnDeleteCode = False
    
    If Index = nBonusCode Then
        strSQL = "DELETE FROM bonus_codes WHERE bc_bonus_code = "
    Else
        strSQL = "DELETE FROM bonus_grades WHERE bg_grade = "
    End If
    
    strSQL = strSQL & tfnSQLString(Trim(sCode))
    
    If Index = nGrade Then
        strSQL = strSQL & " AND bg_emp_level = "
        strSQL = strSQL & tfnRound(tgmBonus(Index).CellValue(nxColBGOldLevel, _
            tgmBonus(Index).GetCurrentRowNumber))
    End If
    
    If fnExecuteSQL(strSQL, , SUB_NAME) Then
        fnDeleteCode = True
    End If

End Function

Private Function fnInsertGrdOrBonusCode(lRow As Long, Index As Integer) As Boolean
    Const SUB_NAME As String = "fnInsertGrdOrBonusCode"
    Dim strSQL As String
    Dim rsTemp As Recordset
    
    fnInsertGrdOrBonusCode = False
    
    If tgmBonus(Index).CellValue(nColCode, lRow) = "" Then
        fnInsertGrdOrBonusCode = True
        Exit Function
    End If
    
    If Index = nBonusCode Then
        strSQL = "INSERT INTO bonus_codes (bc_bonus_code, bc_code_desc,"
        strSQL = strSQL & " bc_grade, bc_type, bc_frequency, bc_override) VALUES("
        strSQL = strSQL & tfnSQLString(tgmBonus(Index).CellValue(nColCode, lRow)) & ","
        strSQL = strSQL & tfnSQLString(Trim(tgmBonus(Index).CellValue(nColDesc, lRow))) & ","
        strSQL = strSQL & tfnSQLString(Trim(tgmBonus(Index).CellValue(nColBCGrade, lRow))) & ","
        strSQL = strSQL & tfnSQLString(Trim(tgmBonus(Index).CellValue(nColBCType, lRow))) & ","
        strSQL = strSQL & tfnSQLString(Trim(tgmBonus(Index).CellValue(nColBCFreq, lRow))) & ","
        strSQL = strSQL & tfnSQLString(Trim(tgmBonus(Index).CellValue(nColBCOverride, lRow)))
        strSQL = strSQL & ")"
    Else 'nGrade
        strSQL = "INSERT INTO bonus_grades(bg_grade, bg_grade_desc, bg_emp_level)"
        strSQL = strSQL & " VALUES(" & tfnSQLString(tgmBonus(Index).CellValue(nColCode, lRow)) & ","
        strSQL = strSQL & tfnSQLString(Trim(tgmBonus(Index).CellValue(nColDesc, lRow))) & ","
        strSQL = strSQL & tfnRound(tgmBonus(Index).CellValue(nColBGLevel, lRow))
        strSQL = strSQL & ")"
    End If
    
    If fnExecuteSQL(strSQL, , SUB_NAME) Then
        fnInsertGrdOrBonusCode = True
    End If

End Function

Private Function fnLoadData(Index As Integer, Optional bRefresh As Boolean = False) As Boolean
    Const SUB_NAME As String = "fnLoadData"
    Dim strSQL As String
    Dim rsTemp As Recordset
    Dim i As Integer
    Dim lRow As Long
    Dim sLockedUser As String
    
    lRow = tgmBonus(Index).GetCurrentRowNumber
    fnLoadData = False
    
    subShowHourGlass
    tfnSetStatusBarMessage "Loading data. Please wait ..."
    
    If Not bRefresh Then
        'lock the records for editing
        If Index = nBonusCode Then
            strSQL = "bc_bonus_code"
            If Not tfnLockRow("ZZSEBCMT", "bonus_codes", strSQL, , sLockedUser) Then
                tfnSetStatusBarError "Bonus code record(s) is locked by " + sLockedUser
                subHideHourGlass
                Exit Function
            End If
        Else 'nGrade
            strSQL = "bg_grade"
            If Not tfnLockRow("ZZSEBCMT", "bonus_grades", strSQL, , sLockedUser) Then
                tfnSetStatusBarError "Bonus grade record(s) is locked by " + sLockedUser
                subHideHourGlass
                Exit Function
            End If
        End If
    End If
    
    If Index = nBonusCode Then
        strSQL = "SELECT * FROM bonus_codes ORDER BY bc_bonus_code"
    Else 'nGrade
        strSQL = "SELECT * FROM bonus_grades ORDER BY bg_grade, bg_emp_level"
    End If
    
    If GetRecordSet(rsTemp, strSQL, , SUB_NAME) < 0 Then
        MsgBox "Failed to access database to get records.", vbCritical
        tfnSetStatusBarError "Failed to access database to get records"
        subHideHourGlass
        Exit Function
    End If
    
    If rsTemp.RecordCount = 0 Then
        MsgBox "No record available for Edit.", vbExclamation
        tfnSetStatusBarError "No record available for Edit"
        subHideHourGlass
        Exit Function
    End If
    
    tgmBonus(Index).FillWithRecordset rsTemp
    
    If Index = nGrade Then
        If tgmBonus(nGrade).RowCount > 0 Then
            'save the old grade level
            For lRow = 1 To tgmBonus(nGrade).RowCount - 1
                tgmBonus(nGrade).CellValue(nxColBGOldLevel, lRow) = _
                    tgmBonus(nGrade).CellValue(nColBGLevel, lRow)
            Next lRow
        End If
        
        subResetDescEditableFlag True
    End If
    
    fnLoadData = True
    
    tblBonus(Index).col = nColDesc
    subHideHourGlass
        
End Function

Private Function fnUpdateBonusCode(lRow As Long) As Boolean
    Const SUB_NAME As String = "fnUpdateBonusCode"
    Dim strSQL As String
    Dim rsTemp As Recordset
    
    fnUpdateBonusCode = False
    
    strSQL = "UPDATE bonus_codes SET"
    strSQL = strSQL & " bc_code_desc = " & tfnSQLString(Trim(tgmBonus(nBonusCode).CellValue(nColDesc, lRow)))
    strSQL = strSQL & ", bc_type = " & tfnSQLString(Trim(tgmBonus(nBonusCode).CellValue(nColBCType, lRow)))
    strSQL = strSQL & ", bc_grade = " & tfnSQLString(Trim(tgmBonus(nBonusCode).CellValue(nColBCGrade, lRow)))
    strSQL = strSQL & ", bc_frequency = " & tfnSQLString(Trim(tgmBonus(nBonusCode).CellValue(nColBCFreq, lRow)))
    strSQL = strSQL & ", bc_override = " & tfnSQLString(Trim(tgmBonus(nBonusCode).CellValue(nColBCOverride, lRow)))
    strSQL = strSQL & " WHERE bc_bonus_code = " & tfnSQLString(Trim(tgmBonus(nBonusCode).CellValue(nColCode, lRow)))
    
    fnUpdateBonusCode = fnExecuteSQL(strSQL, , SUB_NAME)
End Function

Private Function fnUpdateBonusGrade() As Boolean
    Const SUB_NAME As String = "fnUpdateBonusGrade"
    Dim strSQL As String
    Dim i As Long
    
    fnUpdateBonusGrade = False
    
    strSQL = "DELETE FROM bonus_grades"
    
    If Not fnExecuteSQL(strSQL, , SUB_NAME) Then
        Exit Function
    End If
    
    For i = 0 To tgmBonus(nGrade).RowCount - 1
        If Not fnInsertGrdOrBonusCode(i, nGrade) Then
            Exit Function
        End If
    Next i
    
    fnUpdateBonusGrade = True
End Function

Private Function fnCheckDubInGrid(nCol_Nbr As Integer, lCurrRow As Long, _
                                  Index As Integer) As String
'This function will return all the values already used in the grid.
    Dim lRowCount As Long
    Dim k As Integer
    Dim sTemp As String
    
    fnCheckDubInGrid = ""
    
    lRowCount = tgmBonus(Index).RowCount
    sTemp = ""
    
    If lRowCount > 0 Then
        For k = 0 To lRowCount - 1
            If k <> lCurrRow Then
                If sTemp = "" Then
                    If nCol_Nbr = nColBGLevel Then
                        sTemp = tgmBonus(Index).CellValue(nColBGLevel, k)
                    Else
                        sTemp = tfnSQLString(tgmBonus(Index).CellValue(nColCode, k))
                    End If
                Else
                    If nCol_Nbr = nColBGLevel Then
                        sTemp = sTemp & "," & tgmBonus(Index).CellValue(nColBGLevel, k)
                    Else
                        sTemp = sTemp & "," & tfnSQLString(tgmBonus(Index).CellValue(nColCode, k))
                    End If
                End If
            End If
        Next
    End If
    
    fnCheckDubInGrid = sTemp

End Function

Private Sub subParseString(sParam() As String, _
                           sSrc As String, _
                           sDelim As String, _
                           Optional vStart As Variant, _
                           Optional vEnd As Variant)
                          
    If Trim(sSrc) = "" Then
        Exit Sub
    End If

    Const nArrayInc As Integer = 5
    Dim i1 As Integer
    Dim i2 As Integer
    Dim k As Integer
    Dim nEnd As Integer
    Dim sTemp As String
    
    If IsMissing(vStart) Then
        i1 = 1
    Else
        i1 = vStart
    End If
    If IsMissing(vEnd) Then
        nEnd = Len(sSrc)
    Else
        nEnd = vEnd
    End If
    If i1 < 1 Then i1 = 1
    i2 = 1
    k = 0
    ReDim sParam(nArrayInc)
    While i1 <= nEnd And i2 > 0 And i2 <= nEnd
        i2 = InStr(i1, sSrc, sDelim)
        If i2 >= i1 And i2 <= nEnd Then
            If k > UBound(sParam) Then
                ReDim Preserve sParam(k + nArrayInc)
            End If
            sTemp = Mid$(sSrc, i1, i2 - i1)
            If sTemp <> "" Or sDelim <> " " Then
                sParam(k) = sTemp
                k = k + 1
            End If
            i1 = i2 + 1
        End If
    Wend
    If i2 <= nEnd Then
        If k > UBound(sParam) Then
            ReDim Preserve sParam(k + nArrayInc)
        End If
        sParam(k) = Trim$(Mid$(sSrc, i1, nEnd - i1 + 1))
        ReDim Preserve sParam(k)
    Else
        If k > 0 Then
            sParam(k - 1) = Trim$(Mid$(sSrc, i1, nEnd - i1 + 1))
            ReDim Preserve sParam(k - 1)
        End If
    End If
End Sub

Private Sub subSetStdBtn(Index As Integer)
    
    If tgmBonus(Index).RowCount < 1 Then
        subEnableUpdateBtn False, Index
        subEnableDeleteBtn False, Index
    Else
        If tgmBonus(Index).ValidData Then
            If t_nFormMode = ADD_MODE Then
                subEnableUpdateBtn True, Index
            Else
                If nDataStatus = DATA_CHANGED Then
                    subEnableUpdateBtn True, Index
                Else
                    subEnableUpdateBtn False, Index
                End If
            End If
        Else
            subEnableUpdateBtn False, Index
        End If
    End If
End Sub

Private Sub tblView_GotFocus(Index As Integer)
    If Index = nFrequency Then
        tfnSetStatusBarMessage "Commission Frequency"
    Else
        tfnSetStatusBarMessage "Commission Type"
    End If
End Sub

Private Sub tblView_UnboundReadData(Index As Integer, ByVal RowBuf As DBTrueGrid.RowBuffer, StartLocation As Variant, ByVal ReadPriorRows As Boolean)
    tgmView(Index).ReadData RowBuf, StartLocation, ReadPriorRows
End Sub

Private Sub subLoadViewData()
    Const SUB_NAME As String = "subLoadViewData"
    Dim strSQL As String
    
    #If PROTOTYPE Then
        Exit Sub
    #End If
    
    'load frequency...
    strSQL = "SELECT * FROM bonus_frequency ORDER BY bfq_frequency"
    tgmView(nFrequency).FillWithSQL t_dbMainDatabase, strSQL, , True
    
    'Load Types...
    strSQL = "SELECT * FROM bonus_types ORDER BY bt_type"
    tgmView(nType).FillWithSQL t_dbMainDatabase, strSQL, , True
    
End Sub

Private Sub tblBonus_AfterColEdit(Index As Integer, ByVal ColIndex As Integer)
    Dim lRow As Long
    Dim i As Long
    
    lRow = tgmBonus(Index).GetCurrentRowNumber
    tgmBonus(Index).AfterColEdit ColIndex
    
    'when a description for a Grade changed
    'change all description for the same Grade
    If Index = nGrade Then
        If ColIndex = nColDesc Then
            If fnGetField(tgmBonus(Index).CellValue(ColIndex, lRow)) <> _
               fnGetField(tgmBonus(Index).OldValue) Then
                If tgmBonus(Index).ValidCell(ColIndex, lRow) Then
                    For i = 0 To tgmBonus(Index).RowCount - 1
                        If i <> lRow Then
                            If fnGetField(tgmBonus(Index).CellValue(nColCode, lRow)) = _
                               fnGetField(tgmBonus(Index).CellValue(nColCode, i)) Then
                                tgmBonus(Index).CellValue(ColIndex, i) = _
                                    tgmBonus(Index).CellValue(ColIndex, lRow)
                            End If
                        End If
                    Next i
                    tgmBonus(Index).Rebind
                End If
            End If
        End If
    End If
    
    If t_nFormMode = EDIT_MODE Then
        If nDataStatus = DATA_CHANGED Then
            subEnableRefreshBtn True, Index
        Else
            subEnableRefreshBtn False, Index
        End If
    End If
End Sub

Private Sub tblBonus_BeforeColEdit(Index As Integer, ByVal ColIndex As Integer, ByVal KeyAscii As Integer, CANCEL As Integer)
    If Index = nGrade And ColIndex = nColDesc Then
        If tgmBonus(Index).CellValue(nxColBGDescEditable, tgmBonus(Index).GetCurrentRowNumber) = "N" Then
            CANCEL = True
            Exit Sub
        End If
    End If
    tgmBonus(Index).BeforeColEdit ColIndex, KeyAscii, CANCEL
    
    If t_nFormMode = ADD_MODE And Index = nBonusCode Then
        If Trim(tgmBonus(Index).CellValue(nColBCOverride, tgmBonus(Index).GetCurrentRowNumber)) = "" Then
            tgmBonus(Index).CellValue(nColBCOverride, tgmBonus(Index).GetCurrentRowNumber) = "N"
        End If
    End If
    


End Sub

Private Sub tblBonus_Change(Index As Integer)
    tgmBonus(Index).Change
    
    subEnableUpdateBtn False, Index
    nDataStatus = DATA_CHANGED
End Sub

Private Sub tblBonus_FirstRowChange(Index As Integer)
    tgmBonus(Index).FirstRowChange
    tgfDropdown(Index).FirstRowChange
    
    If tblBonus(Index).Row = -1 Then
        subEnableDeleteBtn False, Index
    End If
End Sub

Private Sub tblBonus_GotFocus(Index As Integer)
    If Index = nBonusCode Then
        tfnSetStatusBarMessage "Commission Codes"
    Else
        tfnSetStatusBarMessage "Commission Grades"
    End If

    tgmBonus(Index).GotFocus
    tgfDropdown(Index).GotFocus
End Sub

Private Sub tblBonus_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    tgmBonus(Index).KeyDown KeyCode, Shift
    tgfDropdown(Index).KeyDown tblBonus(Index), KeyCode
    
    If KeyCode = vbKeyDelete Then
        If Not tblBonus(Index).EditActive Then
            tgmBonus(Index).CellValue(tblBonus(Index).col, tgmBonus(Index).GetCurrentRowNumber) = ""
            tgmBonus(Index).ResetFlags tgmBonus(Index).GetCurrentRowNumber, tblBonus(Index).col
            If t_nFormMode = EDIT_MODE Then
                subEnableRefreshBtn True, Index
            End If
        End If
    End If
End Sub

Private Sub tblBonus_KeyPress(Index As Integer, KeyAscii As Integer)
    Dim lRow As Long
    
    If tblBonus(Index).col = nColCode Or tblBonus(Index).col = nColBCOverride Or tblBonus(Index).col = nColBCType _
       Or tblBonus(Index).col = nColBCGrade Or tblBonus(Index).col = nColBCFreq Then
        KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
    End If
    
    tgfDropdown(Index).Keypress tblBonus(Index), KeyAscii
    
    lRow = tgmBonus(Index).GetCurrentRowNumber
    
    If Not tgmBonus(Index).Keypress(KeyAscii) Then
        KeyAscii = 0
    End If
    
    If t_nFormMode = EDIT_MODE Then
        If KeyAscii = vbKeyReturn And tgmBonus(Index).RowCount - 1 = lRow Then
            If Index = nBonusCode Then
                If tblBonus(Index).col = nColBCFreq Then
                    'subSetStdBtn Index
                    If cmdUpdateInsertBtn(Index).Enabled Then
                        subSetFocus cmdUpdateInsertBtn(Index)
                    Else
                        subSetFocus cmdExitCancelBtn(Index)
                    End If
                End If
            Else
                If tblBonus(Index).col = nColBGLevel Then
                    subSetStdBtn Index
                    'If cmdUpdateInsertBtn(Index).Enabled Then
                    '    subSetFocus cmdUpdateInsertBtn(Index)
                    'Else
                    '    subSetFocus cmdExitCancelBtn(Index)
                    'End If
                End If
            End If
        End If
    End If
    
End Sub

Private Sub tblBonus_LeftColChange(Index As Integer)
    tgfDropdown(Index).LeftColChange
End Sub

Private Sub tblBonus_LostFocus(Index As Integer)
    Dim lRow As Long
    lRow = tgmBonus(Index).GetCurrentRowNumber
    tgmBonus(Index).LostFocus
    tgfDropdown(Index).LostFocus tblBonus(Index)
    If Not tgmBonus(Index).ValidData(lRow) And lRow = 0 Then
        Exit Sub
    End If
    subSetStdBtn Index
    If ActiveControl Is cmdExitCancelBtn Then
        subSetFocus cmdUpdateInsertBtn(Index)
    End If
End Sub

Private Sub tblBonus_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
'    If Index = nTabUOM Then
'        frmContext.MouseDown Button, UOM_UP
'    End If
End Sub

Private Sub tblBonus_RowColChange(Index As Integer, LastRow As Variant, ByVal LastCol As Integer)
    Dim lRow As Long
    Dim sValue As String
    
    If tgfDropdown(Index).RowColChange(tblBonus(Index)) Then
        tgmBonus(Index).RowColChange LastRow, LastCol
    End If
    
    lRow = tgmBonus(Index).GetCurrentRowNumber
    
    If tgmBonus(Index).RowCount > 0 Then
        subEnableDeleteBtn True, Index
    Else
        subEnableDeleteBtn False, Index
    End If

    If Not tgmBonus(Index).ValidData(lRow) And lRow = 0 Then
        Exit Sub
    End If
    
    
    subSetStdBtn Index
    
End Sub

Private Sub tblBonus_SelChange(Index As Integer, CANCEL As Integer)
    CANCEL = True
End Sub

Private Sub tblBonus_UnboundReadData(Index As Integer, ByVal RowBuf As DBTrueGrid.RowBuffer, StartLocation As Variant, ByVal ReadPriorRows As Boolean)
    tgmBonus(Index).ReadData RowBuf, StartLocation, ReadPriorRows
End Sub

Private Sub subInitViewGrid()
    Dim Index As Integer
    
    For Index = 0 To 1
        If Not tgmView(Index) Is Nothing Then Set tgmView(Index) = Nothing
        Set tgmView(Index) = New clsTGSpreadSheet
        Set tgmView(Index).Table = tblView(Index)
        Set tgmView(Index).StatusBar = ffraStatusbar ' Message bar name
        Set tgmView(Index).Form = Me
        Set tgmView(Index).engFactor = t_engFactor
        tgmView(Index).SetupTable True
        tgmView(Index).AllowAddNew = False
    Next Index

End Sub

Private Function fnValidBonusPayCode(szText As String, ByVal lRow As Long) As Boolean

    Const SUB_NAME As String = "fnValidBonusPayCode"
    Dim strSQL As String
    Dim rsTemp As Recordset
    Dim sErrMsg As String

    fnValidBonusPayCode = False

    If Trim(szText) = "" Then
        tgmBonus(nBonusCode).ErrorMessage(nBonusCode) = "You must enter a pay code"
        Exit Function
    End If

    strSQL = "SELECT prpa_desc, prpa_type FROM pr_pay WHERE prpa_pay_code = " & tfnSQLString(szText)

    If GetRecordSet(rsTemp, strSQL, , SUB_NAME) <= 0 Then
        tgmBonus(nBonusCode).ErrorMessage(nBonusCode) = "Pay code does not exist"
        Exit Function
    End If

    If fnGetField(rsTemp!prpa_type) <> "P" Then
        tgmBonus(nBonusCode).ErrorMessage(nBonusCode) = "Pay code Type is not 'P'"
        Exit Function
    End If

    If fnGetField(tgmBonus(nBonusCode).CellValue(nColDesc, lRow)) = "" Then
        tgmBonus(nBonusCode).CellValue(nColDesc, lRow) = fnGetField(rsTemp!prpa_desc)
    End If

    fnValidBonusPayCode = True
End Function

Private Function fnValidBonusGrade(szText As String, ByVal lRow As Long) As Boolean
    Const SUB_NAME As String = "fnValidBonusGrade"
    Dim strSQL As String

    fnValidBonusGrade = False
    
    If Trim(szText) = "" Then
        tgmBonus(nBonusCode).ErrorMessage(nColBCGrade) = "You must enter a commission grade"
        Exit Function
    End If
    
    strSQL = "SELECT bg_grade FROM bonus_grades WHERE bg_grade = " & tfnSQLString(szText)
    
    If GetRecordCount(strSQL, , SUB_NAME) = 0 Then
        tgmBonus(nBonusCode).ErrorMessage(nColBCGrade) = "Commission grade does not exist"
        Exit Function
    End If
    
    fnValidBonusGrade = True

End Function

Private Function fnValidGradeLevel(szText As String, ByVal lRow As Long) As Boolean
    
    Const SUB_NAME As String = "fnValidGradeLevel"
    Dim strSQL As String
    Dim rsTemp As Recordset
    Dim i As Long
    Dim bFoundDuplicate  As Boolean
    
    fnValidGradeLevel = False
    
    If Trim(szText) = "" Then
        tgmBonus(nGrade).ErrorMessage(nColBGLevel) = "You must enter a Grade Level"
        Exit Function
    End If
    
    If tfnRound(szText) < 0 Then
        tgmBonus(nGrade).ErrorMessage(nColBGLevel) = "Grade Level must be greater than 0"
        Exit Function
    End If
    
    'david 01/02/2001
    'PER BOB, DO NOT VALIDATE AGAINST THE PR_MASTER.prm_emp_level
    #If False Then
        strSQL = "SELECT prm_emp_level FROM pr_master WHERE prm_emp_level = " & tfnRound(szText)
        strSQL = strSQL + " GROUP BY prm_emp_level"
        If GetRecordCount(strSQL, , SUB_NAME) = 0 Then
            tgmBonus(nGrade).ErrorMessage(nColBGLevel) = "Grade Level does not exist"
            Exit Function
        End If
    #End If
    
    'check bonus_grades table for unique bg_emp_level
    If t_nFormMode = ADD_MODE Then
        strSQL = "SELECT bg_grade FROM bonus_grades WHERE bg_emp_level = " & tfnRound(szText)
        If GetRecordSet(rsTemp, strSQL, , SUB_NAME) > 0 Then
            tgmBonus(nGrade).ErrorMessage(nColBGLevel) = "Grade Level already been assigned to Grade " + tfnSQLString(fnGetField(rsTemp!bg_grade))
            Exit Function
        End If
    End If
    'check pr_master table for valid values
    strSQL = "SELECT grade_level FROM tmpGradeLevel WHERE grade_level = " & tfnRound(szText)
    If GetRecordSet(rsTemp, strSQL, , SUB_NAME) = 0 Then
        tgmBonus(nGrade).ErrorMessage(nColBGLevel) = "Grade Level is not a valid value. "
        Exit Function
    End If
    
    'check for duplication
    bFoundDuplicate = False
    For i = 0 To tgmBonus(nGrade).RowCount - 1
        If i <> lRow Then
            If tfnRound(tgmBonus(nGrade).CellValue(nColBGLevel, i)) = tfnRound(szText) Then
                tgmBonus(nGrade).CellValue(nxColBGDuplicates, i) = "Y"
                tgmBonus(nGrade).ValidCell(nColBGLevel, i) = False
                bFoundDuplicate = True
            Else
                If fnGetField(tgmBonus(nGrade).CellValue(nxColBGDuplicates, i)) = "Y" Then
                    If tfnRound(tgmBonus(nGrade).CellValue(nColBGLevel, i)) = tfnRound(tgmBonus(nGrade).OldValue) Then
                        'reset the duplication flag
                        tgmBonus(nGrade).CellValue(nxColBGDuplicates, i) = ""
                        tgmBonus(nGrade).ValidCell(nColBGLevel, i) = True
                    End If
                End If
            End If
        End If
    Next i

    If bFoundDuplicate Then
        tgmBonus(nGrade).CellValue(nxColBGDuplicates, lRow) = "Y"
        tgmBonus(nGrade).ErrorMessage(nColBGLevel) = "Duplicate Grade level encountered"
        Exit Function
    End If
    
    subRevalidate nColBGLevel

    tgmBonus(nGrade).CellValue(nxColBGDuplicates, lRow) = ""
    fnValidGradeLevel = True

End Function

Private Sub subRevalidate(ByVal nCol As Integer)
    Dim i As Long
    
    For i = 0 To tgmBonus(nGrade).RowCount - 1
        If Not tgmBonus(nGrade).Revalidate(nCol, i) Then
            Exit For
        End If
    Next
End Sub

Private Function fnValidBonusType(szText As String, ByVal lRow As Long) As Boolean
    Const SUB_NAME As String = "fnValidBonusType"
    Dim strSQL As String

    fnValidBonusType = False
    
    If Trim(szText) = "" Then
        tgmBonus(nBonusCode).ErrorMessage(nColBCType) = "You must enter a Commission type"
        Exit Function
    End If
    
    strSQL = "SELECT bt_type FROM bonus_types WHERE bt_type = " & tfnSQLString(szText)
    
    If GetRecordCount(strSQL, , SUB_NAME) = 0 Then
        tgmBonus(nBonusCode).ErrorMessage(nColBCType) = "Commission type does not exist"
        Exit Function
    End If
    
    fnValidBonusType = True

End Function

Private Function fnValidBonusFreq(szText As String, ByVal lRow As Long) As Boolean
    Const SUB_NAME As String = "fnValidBonusFreq"
    Dim strSQL As String

    fnValidBonusFreq = False
    
    If Trim(szText) = "" Then
        tgmBonus(nBonusCode).ErrorMessage(nColBCFreq) = "You must enter a commission frequency"
        Exit Function
    End If
    
    strSQL = "SELECT * FROM bonus_frequency WHERE bfq_frequency = " & tfnSQLString(szText)
    
    If GetRecordCount(strSQL, , SUB_NAME) <= 0 Then
        tgmBonus(nBonusCode).ErrorMessage(nColBCFreq) = "Commission frequency does not exist"
        Exit Function
    End If
    
    fnValidBonusFreq = True

End Function

Private Function fnGetPayCodeDesc(sCode As String, Optional sErrMsg As String = "") As String
    Const SUB_NAME As String = "fnGetPayCodeDesc"
    Dim strSQL As String
    Dim rsTemp As Recordset
    
    fnGetPayCodeDesc = ""
    
    strSQL = "SELECT prpa_desc FROM pr_pay WHERE prpa_pay_code = " & tfnSQLString(Trim(sCode))
    
    If GetRecordSet(rsTemp, strSQL, , SUB_NAME) <= 0 Then
        fnGetPayCodeDesc = "Invalid pay code"
        sErrMsg = "Pay code does not exist"
    Else
        fnGetPayCodeDesc = fnGetField(rsTemp!prpa_desc)
    End If

End Function

Private Function fnCheckDelete(sCode As String, Index As Integer) As Boolean
    Const SUB_NAME As String = "fnCheckDelete"
    Dim strSQL As String
    Dim sTitle As String
    Dim sByWhom As String
    Dim i As Long
    Dim nCount As Integer
    
    If Index = nBonusCode Then
        sTitle = "Bonus code"
        sByWhom = "Bonus Formula"
        strSQL = "SELECT bf_bonus_code FROM bonus_formula WHERE bf_bonus_code = "
    Else
        'count GradeCode
        For i = 0 To tgmBonus(nGrade).RowCount - 1
            If fnGetField(tgmBonus(nGrade).CellValue(nColCode, i)) = sCode Then
                nCount = nCount + 1
                If nCount > 1 Then
                    Exit For
                End If
            End If
        Next i
        
        If nCount > 1 Then
            fnCheckDelete = True
            Exit Function
        End If
        
        sTitle = "Bonus grade"
        sByWhom = "Bonus Code"
        strSQL = "SELECT bc_bonus_code FROM bonus_codes WHERE bc_grade = "
    End If
    
    strSQL = strSQL & tfnSQLString(Trim(sCode))
    
    If GetRecordCount(strSQL, , SUB_NAME) > 0 Then
        MsgBox sTitle + " is in used by " + sByWhom + ", and cannot be deleted.", vbExclamation
        Exit Function
    End If

    fnCheckDelete = True
End Function

Private Sub subMakeCommFrequencyPattern()
    Const SUB_NAME As String = "subMakeCommFrequencyPattern"
    Dim strSQL As String
    Dim rsTemp As Recordset
    Dim i As Integer
    
    strSQL = "SELECT bfq_frequency FROM bonus_frequency"
    
    If GetRecordSet(rsTemp, strSQL, , SUB_NAME) > 0 Then
        szCommFreqPattern = "^(["
        For i = 1 To rsTemp.RecordCount
            szCommFreqPattern = szCommFreqPattern + fnGetField(rsTemp!bfq_frequency)
            rsTemp.MoveNext
        Next i
        szCommFreqPattern = szCommFreqPattern + "])$"
    End If
    
    If szCommFreqPattern = "" Then
        szCommFreqPattern = "^([MPA])$"
    End If
End Sub

'this function also stores the grade level into a hidden column
Private Sub subResetDescEditableFlag(Optional bChangeDesc As Boolean = False)
    Dim lRow As Long
    Dim i As Long
    Dim sGrade As String
    Dim sDesc As String
    
    If tgmBonus(nGrade).RowCount = 0 Then
        Exit Sub
    End If
    
    
    For lRow = 0 To tgmBonus(nGrade).RowCount - 1
        sDesc = fnGetField(tgmBonus(nGrade).CellValue(nColDesc, lRow))
        
        For i = 0 To lRow - 1
            If fnGetField(tgmBonus(nGrade).CellValue(nColCode, lRow)) = _
               fnGetField(tgmBonus(nGrade).CellValue(nColCode, i)) Then
                tgmBonus(nGrade).CellValue(nxColBGDescEditable, lRow) = "N"
                If bChangeDesc Then
                    tgmBonus(nGrade).CellValue(nColDesc, lRow) = sDesc
                End If
                Exit For
            End If
        Next i
            
        'this is the first grade code
        If i > lRow - 1 Then
            tgmBonus(nGrade).CellValue(nxColBGDescEditable, lRow) = "Y"
        End If
    Next lRow
    
    If bChangeDesc Then
        tgmBonus(nGrade).Rebind
    End If
End Sub

Private Function fnCreateTbl_tmpGradeLevel() As Boolean
    Const SUB_NAME = "fnCreateTbl_tmpGradeLevel"
    Dim strSQL As String
    fnCreateTbl_tmpGradeLevel = False
    On Error Resume Next
    strSQL = "DROP TABLE tmpGradeLevel"
    t_dbMainDatabase.ExecuteSQL strSQL
    
    On Error GoTo err_handler
    strSQL = "SELECT DISTINCT prm_emp_level grade_level FROM pr_master WHERE prm_emp_level is Not Null INTO TEMP tmpGradeLevel "
    t_dbMainDatabase.ExecuteSQL strSQL
    
    fnCreateTbl_tmpGradeLevel = True
    Exit Function
err_handler:
    tfnErrHandler SUB_NAME, strSQL
End Function

'Private Function fnUpdateTbl_tmpGradeLevel(tblIndex As Integer) As Boolean
'    Const SUB_NAME = "fnUpdateTbl_tmpGradeLevel"
'    Dim strSQL As String
'    Dim vTemp() As Variant
'    Dim rsTemp As Recordset
'    Dim lRows As Long
'
'    fnCreateTbl_tmpGradeLevel = False
'
'    On Error GoTo err_handler
'    If Not fnCreateTbl_tmpGradeLevel Then
'        Exit Function
'    End If
'
'    tgmBonus(tblIndex).GetRecords vTemp, lRows, "bg_emp_level"
'
'    If lRows <= 0 Then
'        fnUpdateTbl_tmpGradeLevel = True
'        Exit Function
'    End If
'
'    Dim i As Integer
'    Dim nLevel As Integer
'    For i = 0 To lRows - 1
'        nLevel = vTemp(0, i)
'        strSQL = "DELETE FROM tmpGradeLevel WHERE grade_level=" & nLevel
'        t_dbMainDatabase.ExecuteSQL strSQL
'
'    Next
'    fnUpdateTbl_tmpGradeLevel = True
'    Exit Function
'err_handler:
'    tfnErrHandler SUB_NAME, strSQL
'End Function
'
