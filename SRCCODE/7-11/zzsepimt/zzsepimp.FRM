VERSION 5.00
Object = "{C75015E0-2232-11D3-B440-0060971E99AF}#1.0#0"; "FACTFRM.OCX"
Object = "{3D388220-1F4E-11D3-B440-0060971E99AF}#1.0#0"; "FACTTAB.OCX"
Object = "{01028C21-0000-0000-0000-000000000046}#4.0#0"; "TG32OV.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{478E45E0-5745-11CF-8918-00A02416C765}#1.0#0"; "SQAOTE32.OCX"
Begin VB.Form frmZZSEPIMT 
   BackColor       =   &H00C0C0C0&
   Caption         =   "ZZSEPIMT"
   ClientHeight    =   6060
   ClientLeft      =   1890
   ClientTop       =   2175
   ClientWidth     =   8880
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
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6060
   ScaleWidth      =   8880
   Begin FACTFRMLib.FactorFrame efraBackground 
      Height          =   5184
      Left            =   0
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   516
      Width           =   8880
      _Version        =   65536
      _ExtentX        =   15663
      _ExtentY        =   9144
      _StockProps     =   77
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.99
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
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin FACTTABLib.FactorTab etabMain 
         Height          =   4905
         HelpContextID   =   201
         Left            =   0
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   0
         Width           =   8895
         _Version        =   65536
         _ExtentX        =   15690
         _ExtentY        =   8652
         _StockProps     =   68
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabsPerPage     =   4
         Caption         =   "&Pay Increase maint.|Pay &Increase Report"
         Begin FACTFRMLib.FactorFrame efraTabBack1 
            Height          =   4545
            Left            =   15
            TabIndex        =   7
            TabStop         =   0   'False
            Top             =   15
            Width           =   8865
            _Version        =   65536
            _ExtentX        =   15637
            _ExtentY        =   8017
            _StockProps     =   77
            BackColor       =   8388608
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   0
            TitleBarHeight  =   24
            BeginProperty PanelFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Begin FACTFRMLib.FactorFrame efraTabFrame1 
               Height          =   3852
               Left            =   96
               TabIndex        =   8
               TabStop         =   0   'False
               Top             =   96
               Width           =   8652
               _Version        =   65536
               _ExtentX        =   15266
               _ExtentY        =   6800
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
               BevelOuter      =   5
               TitleBarHeight  =   24
               BeginProperty PanelFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Begin DBTrueGrid.TDBGrid tblSpreadsheet 
                  Height          =   1755
                  Left            =   510
                  OleObjectBlob   =   "zzsepimp.frx":0000
                  TabIndex        =   9
                  Top             =   1620
                  Width           =   5145
               End
               Begin FACTFRMLib.FactorFrame FactorFrame1 
                  Height          =   1065
                  Left            =   135
                  TabIndex        =   10
                  Top             =   150
                  Width           =   8430
                  _Version        =   65536
                  _ExtentX        =   14870
                  _ExtentY        =   1879
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
                  BevelOuter      =   6
                  TitleBarHeight  =   24
                  BeginProperty PanelFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   12
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Begin VB.TextBox txtDesc 
                     Height          =   420
                     Left            =   3330
                     TabIndex        =   12
                     Text            =   "Text2"
                     Top             =   480
                     Width           =   3495
                  End
                  Begin VB.TextBox txtLevel 
                     Height          =   420
                     Left            =   390
                     TabIndex        =   11
                     Top             =   480
                     Width           =   2205
                  End
                  Begin FACTFRMLib.FactorFrame cmdLevel 
                     Height          =   390
                     HelpContextID   =   2901
                     Left            =   2610
                     TabIndex        =   23
                     TabStop         =   0   'False
                     Tag             =   "Run #3"
                     Top             =   495
                     Width           =   390
                     _Version        =   65536
                     _ExtentX        =   688
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
                     CaptionPos      =   4
                     Picture         =   "zzsepimp.frx":1326
                     Style           =   3
                     BorderWidth     =   4
                     TitleBarHeight  =   24
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
                  Begin FACTFRMLib.FactorFrame cmdCustNoSrch4 
                     Height          =   390
                     HelpContextID   =   2901
                     Left            =   6840
                     TabIndex        =   24
                     TabStop         =   0   'False
                     Tag             =   "Run #3"
                     Top             =   495
                     Width           =   390
                     _Version        =   65536
                     _ExtentX        =   688
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
                     CaptionPos      =   4
                     Picture         =   "zzsepimp.frx":1438
                     Style           =   3
                     BorderWidth     =   4
                     TitleBarHeight  =   24
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
                  Begin VB.Label Label3 
                     Caption         =   "Employee Level Description"
                     Height          =   345
                     Left            =   3345
                     TabIndex        =   26
                     Top             =   135
                     Width           =   3420
                  End
                  Begin VB.Label Label2 
                     Caption         =   "Employee Level"
                     Height          =   270
                     Left            =   420
                     TabIndex        =   25
                     Top             =   165
                     Width           =   2085
                  End
               End
            End
            Begin FACTFRMLib.FactorFrame cmdDeleteBtn 
               Height          =   390
               HelpContextID   =   12
               Index           =   0
               Left            =   3420
               TabIndex        =   13
               Top             =   4080
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
               Caption         =   "&Delete"
               CaptionPos      =   4
               PicturePos      =   3
               ShowFocusRect   =   -1  'True
               Style           =   3
               BorderWidth     =   4
               TitleBarHeight  =   24
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
            Begin FACTFRMLib.FactorFrame cmdEditBtn 
               Height          =   390
               HelpContextID   =   11
               Index           =   0
               Left            =   2085
               TabIndex        =   14
               Top             =   4080
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
               Caption         =   "&Edit"
               CaptionPos      =   4
               PicturePos      =   3
               ShowFocusRect   =   -1  'True
               Style           =   3
               BorderWidth     =   4
               TitleBarHeight  =   24
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
            Begin FACTFRMLib.FactorFrame cmdAddBtn 
               Height          =   390
               HelpContextID   =   10
               Index           =   0
               Left            =   750
               TabIndex        =   15
               Top             =   4080
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
               Caption         =   "&Add"
               CaptionPos      =   4
               PicturePos      =   3
               ShowFocusRect   =   -1  'True
               Style           =   3
               BorderWidth     =   4
               TitleBarHeight  =   24
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
            Begin FACTFRMLib.FactorFrame cmdRefreshSelectBtn 
               Height          =   390
               HelpContextID   =   14
               Index           =   0
               Left            =   4755
               TabIndex        =   16
               Top             =   4080
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
               Caption         =   "&Refresh"
               CaptionPos      =   4
               PicturePos      =   3
               ShowFocusRect   =   -1  'True
               Style           =   3
               BorderWidth     =   4
               TitleBarHeight  =   24
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
            Begin FACTFRMLib.FactorFrame cmdCancelBtn 
               Height          =   390
               HelpContextID   =   15
               Index           =   0
               Left            =   7440
               TabIndex        =   17
               Top             =   4080
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
               Caption         =   "Ca&ncel"
               CaptionPos      =   4
               PicturePos      =   3
               ShowFocusRect   =   -1  'True
               Style           =   3
               BorderWidth     =   4
               TitleBarHeight  =   24
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
            Begin FACTFRMLib.FactorFrame cmdOKBtn 
               Height          =   390
               HelpContextID   =   16
               Index           =   0
               Left            =   6090
               TabIndex        =   18
               Top             =   4080
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
               Caption         =   "&Update"
               CaptionPos      =   4
               PicturePos      =   3
               ShowFocusRect   =   -1  'True
               Style           =   3
               BorderWidth     =   4
               TitleBarHeight  =   24
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
         End
         Begin FACTFRMLib.FactorFrame efraTabBack2 
            Height          =   4575
            Left            =   18075
            TabIndex        =   19
            TabStop         =   0   'False
            Top             =   0
            Width           =   8895
            _Version        =   65536
            _ExtentX        =   15690
            _ExtentY        =   8070
            _StockProps     =   77
            BackColor       =   8388608
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   0
            TitleBarHeight  =   24
            BeginProperty PanelFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Begin VB.CommandButton cmdProcess 
               Caption         =   "&Process"
               Height          =   390
               Left            =   3030
               TabIndex        =   30
               Top             =   4050
               Width           =   1335
            End
            Begin VB.CommandButton cmdPrint 
               Caption         =   "Print &Report"
               Height          =   390
               Left            =   4485
               TabIndex        =   29
               Top             =   4050
               Width           =   1455
            End
            Begin FACTFRMLib.FactorFrame FactorFrame3 
               Height          =   3855
               Left            =   90
               TabIndex        =   20
               TabStop         =   0   'False
               Top             =   75
               Width           =   8655
               _Version        =   65536
               _ExtentX        =   15266
               _ExtentY        =   6800
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
               BevelOuter      =   5
               TitleBarHeight  =   24
               BeginProperty PanelFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Begin VB.TextBox txtReport 
                  Height          =   3000
                  Left            =   135
                  TabIndex        =   22
                  Top             =   690
                  Width           =   8400
               End
               Begin VB.Label Label1 
                  Caption         =   "AutomaticPay Increase Processing"
                  Height          =   465
                  Left            =   225
                  TabIndex        =   21
                  Top             =   285
                  Width           =   3465
               End
            End
            Begin FACTFRMLib.FactorFrame cmdCancelBtn 
               Height          =   390
               HelpContextID   =   15
               Index           =   1
               Left            =   7440
               TabIndex        =   27
               Top             =   4050
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
               Caption         =   "Ca&ncel"
               CaptionPos      =   4
               PicturePos      =   3
               ShowFocusRect   =   -1  'True
               Style           =   3
               BorderWidth     =   4
               TitleBarHeight  =   24
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
            Begin FACTFRMLib.FactorFrame cmdOKBtn 
               Height          =   390
               HelpContextID   =   16
               Index           =   1
               Left            =   6060
               TabIndex        =   28
               Top             =   4050
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
               Caption         =   "&Update"
               CaptionPos      =   4
               PicturePos      =   3
               ShowFocusRect   =   -1  'True
               Style           =   3
               BorderWidth     =   4
               TitleBarHeight  =   24
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
         End
      End
   End
   Begin FACTFRMLib.FactorFrame ffraStatusbar 
      Height          =   360
      Left            =   0
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   5700
      Width           =   8880
      _Version        =   65536
      _ExtentX        =   15663
      _ExtentY        =   635
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
      Align           =   2
      CaptionPos      =   1
      Style           =   5
      TitleBarHeight  =   24
      BeginProperty PanelFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin FACTFRMLib.FactorFrame efraToolBar 
      Height          =   468
      Left            =   0
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   0
      Width           =   8880
      _Version        =   65536
      _ExtentX        =   15663
      _ExtentY        =   825
      _StockProps     =   77
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.81
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Align           =   1
      FMName          =   "ZZSEPIMT"
      CaptionPos      =   4
      Style           =   6
      TitleBarHeight  =   24
      BeginProperty PanelFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.81
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
         Width           =   6876
         _ExtentX        =   12118
         _ExtentY        =   556
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
      Left            =   2088
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4632
      Visible         =   0   'False
      Width           =   1920
   End
   Begin DBTrueGrid.TDBGrid tblComboDropdown 
      Bindings        =   "zzsepimp.frx":154A
      Height          =   2484
      Left            =   408
      OleObjectBlob   =   "zzsepimp.frx":1569
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   684
      Width           =   3756
   End
   Begin SQAOTestObjectsCtl.SQAOTest SQAOTest1 
      Height          =   570
      Left            =   10890
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   90
      Width           =   570
      _ExtentX        =   1005
      _ExtentY        =   1005
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
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "&Edit"
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "&Delete"
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
Attribute VB_Name = "frmZZSEPIMT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private t_bStartupFlag As Boolean 'optional startup flag
Private t_bDataChanged As Boolean 'data changed flag
Private t_bUpdateTable As Boolean 'update data flag

Private t_nFormMode As Integer         'global used to track the current form operating mode

Private Const IDLE_MODE As Integer = 0   'idle mode activates the NoDrop Cursor
Private Const ADD_MODE As Integer = 1    'flag set when in the add mode
Private Const EDIT_MODE As Integer = 2   'flag set when in the edit mode
Private Const DELETE_MODE As Integer = 3 'flag set when in the delete mode

'=================
'Tabbed Dialog Constants
'=================
Private t_bTabSwitch As Boolean   'true if the tab was switched - used by Click event
Private Const TAB_1 As Integer = 0
Private Const TAB_2 As Integer = 1
Private Const TAB_3 As Integer = 2
Private Const TAB_4 As Integer = 3
Private Const TAB_COUNT As Integer = 4

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

Private Sub subSetupCombos()

    Set tgcDropdown = CreateObject(t_szOLECOMBO)
    Set tgcDropdown.DBEngine = t_engFactor
    Set tgcDropdown.Form = Me
    Set tgcDropdown.Database = t_dbMainDatabase
    Set tgcDropdown.DataLink = datComboDropDown
    Set tgcDropdown.Table = tblComboDropdown
    'Setup the combos
    ' Set the general properties of the combos
    Set tgcDropdown = CreateObject(t_szOLECOMBO)
    
    nComboForAdd = tgcDropdown.AddCombo
    
    tgcDropdown.SQL = "SELECT prm_emp_level FROM pr_master"
    tgcDropdown.AddComboBox txtLevel, cmdLevel, "prm_emp_level", .SQL_INT_TYPE
    
    
    nComboForEdit = tgcDropdown.AddCombo
    tgcDropdown.SQL = "SELECT pi_emp_level, pi_emp_level_desc FROM pay_increase"
    tgcDropdown.AddComboBox txtLevel, cmdLevel, "pi_emp_level", .SQL_INT_TYPE
    tgcDropdown.AddComboBox txtDesc, cmdDesc, "pi_emp_level_desc", .SQL_String_TYPE(20)
        

End Sub
Private Sub subInitializeSpreadSheet()
    Dim sDecPatter As String
    szDecPattern = tfnDecimalPattern(8, 2, False, False)

    Set tgmSpreadsheet = New clsTGSpreadSheet
    Set tgmSpreadsheet.Table = tblSpreadsheet
    Set tgmSpreadsheet.StatusBar = ffraStatusbar
    Set tgmSpreadsheet.Form = Me

    'tgmSpreadsheet.AllowRevalidate = True
    'Add editable column and column value validation
    tgmSpreadsheet.AddEditColumn COL_DAYS, "Enter Days Employed", szIntegerPattern
    tgmSpreadsheet.AddEditColumn COL_RATE, "Enter Hourly Pay Rate", szDecPattern, "#,###,##0.00"

    tgmSpreadsheet.ColumnForNewRow = 0


    With tblSpreadsheet.Columns(COL_DAYS)
        .DataField = "pi_days"
    End With
    With tblSpreadsheet.Columns(COL_RATE)
        .DataField = "pi_pay"
    End With

    tblSpreadsheet.EditDropDown = True

End Sub
Private Sub subSetupMousePointer()
    Set objCursor = New clsCursor
    Set objCursor.MainForm = Me
    
    objCursor.AddControls tblComboDropdown, frmAbout
    
    objCursor.AddScreen TAB_1
    objCursor.AddArea efraTabFrame1
    objCursor.AddScreen TAB_2
    objCursor.AddArea efraTabFrame2
    objCursor.AddScreen TAB_3
    objCursor.AddArea efraTabFrame3
    objCursor.AddScreen TAB_4
    objCursor.AddArea efraTabFrame4
    'objCursor.AddArea txtUserId
    'objCursor.SetLeftTop txtUserId
    'objCursor.SetRightBottom cmdUserIdSearchBtn
End Sub

Private Sub subInitValidation()
    Set cValidate = New cValidateInput
    Set cValidate.Form = Me
    Set cValidate.StatusBar = ffraStatusbar
    
    cValidate.AddEditBox txtLevel, "Enter Employee Level"
    cValidate.AddEditBox txtDesc, "Enter Employee Level Description"

    cValidate.ESCControl = tblComboDropdown
    cValidate.ESCControl = cmdExitCancelBtn
    cValidate.MinTabIndex = 0
    Set cValidate.ControlForFocus = efraTabFrame1
    Set cValidate.LastBox = txtLevel
    cValidate.SetFirstControls cmdAddBtn, cmdEditBtn, cmdDeleteBtn, cmdRefreshSelectBtn, cmdUpdateInsertBtn, cmdExitCancelBtn
End Sub

Public Function fnInvalidData(txtBox As Textbox) As Boolean
End Function

Private Sub Form_Unload(CANCEL As Integer)
    tfnUnlockRow
    
    On Error Resume Next
    
    Set tgcDropdown = Nothing
    Set cValidate = Nothing
    Set objCursor = Nothing
    Set objErrHandler = Nothing
    
    Set objCurrTabControl = Nothing
    
    Unload frmContext
    Unload frmAbout
    
    'project local database object variables
    If Not dbLocal Is Nothing Then
        dbLocal.Close
    End If
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



Private Sub tblComboDropDown_Click()
    tgcDropdown.Click tblComboDropdown
End Sub

Private Sub tblComboDropDown_GotFocus()
    tgcDropdown.GotFocus tblComboDropdown
End Sub


Private Sub tblComboDropDown_KeyPress(KeyAscii As Integer)
    tgcDropdown.Keypress tblComboDropdown, KeyAscii
End Sub


Private Sub tblComboDropDown_LostFocus()
    tgcDropdown.LostFocus tblComboDropdown
End Sub

Private Sub tblComboDropDown_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    tgcDropdown.TableMouseUp y
End Sub


Private Sub tblComboDropDown_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    tgcDropdown.RowColChange
End Sub


Private Sub tblComboDropDown_SelChange(CANCEL As Integer)
    tgcDropdown.SelChange CANCEL
End Sub

Private Sub etabMain_Click()
    DoEvents
    
    If t_bTabSwitch Then
    
        ' if you need to set focus to a control on the new tab,
        ' do it in here - refer to etabMain.CurrTab to get the
        ' current tab.
    
    
        t_bTabSwitch = False
    
    End If
    
    subSetMenu
End Sub

Private Sub etabMain_Switch(OldTab As Integer, NewTab As Integer, CANCEL As Integer)
    DoEvents
    
    If Not etabMain.TabEnabled(NewTab) Then
        CANCEL = True
        Exit Sub
    End If

    ' Add any code to prohibit the switch here - set Cancel to True to stop switching.
    ' Also, don't rely on the value of OldTab.  Only NewTab is valid.
    
    If Not CANCEL Then
        t_bTabSwitch = True
    End If
End Sub


'===========
'Form Events
'===========
Private Sub Form_Initialize() 'called before Form_Load
    t_bStartupFlag = True
    t_bDataChanged = False
    t_bUpdateTable = False
    t_bTabSwitch = False
    
    t_nFormMode = IDLE_MODE
    
    ' disable tab preview
            
    CRLF = Chr(10) + Chr(13)

    ' ** change the help file for this application
    App.HelpFile = szHelp7_11
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
  
    'connect to local database
    Set dbLocal = tfnOpenLocalDatabase()
    If dbLocal Is Nothing Then
        MsgBox "System Error: Unable to open Local Database, Program will be terminated", vbExclamation
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
    
    subSetupCombos  'Set up the combo dropdowns
    subInitializeSpreadSheet 'set up true grid
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
    
    subSetExitCancelBtn "EXIT"
    
    tfnResetScreen "Initail"  'set the default screen
    Screen.MousePointer = vbHourglass
    
    subSetupMousePointer
    t_bStartupFlag = False
    Screen.MousePointer = vbHourglass
    
    ' ^ ^ ^ ^ ^ ^ ^
    ' | | | | | | |
    '***************************************************
    
    tmrKeyBoard.Enabled = True

    Screen.MousePointer = vbDefault
    
    cmdAddBtn(etabMain.CurrTab).SetFocus
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
    tfnSetStatusBarMessage t_szADD
End Sub

Private Sub cmdEditBtn_GotFocus(Index As Integer)
    tfnSetStatusBarMessage t_szEDIT
End Sub

Private Sub cmdDeleteBtn_GotFocus(Index As Integer)
    tfnSetStatusBarMessage t_szDELETE
End Sub

Private Sub cmdUpdateInsertBtn_GotFocus()
    If cmdUpdateInsertBtn.Caption = t_szCAPTION_INSERT Then
        tfnSetStatusBarMessage t_szINSERT
    Else
        tfnSetStatusBarMessage t_szUPDATE
    End If
End Sub

Private Sub cmdRefreshSelectBtn_GotFocus(Index As Integer)
    tfnSetStatusBarMessage t_szREFRESH
End Sub

Private Sub cmdExitCancelBtn_GotFocus()
    If cmdExitCancelBtn.Caption = t_szCAPTION_EXIT Then
        tfnSetStatusBarMessage t_szEXIT
    Else
        tfnSetStatusBarMessage t_szCANCEL
    End If
End Sub

Private Sub cmdAddBtn_Click(Index As Integer)
    Select Case Index
    Case TAB_1
        efraTabFrame1.Enabled = True
    Case TAB_2
        efraTabFrame2.Enabled = True
    Case TAB_3
        efraTabFrame3.Enabled = True
    Case TAB_4
        efraTabFrame4.Enabled = True
    End Select
    
    subEnableAddBtn False
    subEnableEditBtn False
    subEnableCancelBtn True
    
    cmdUpdateInsertBtn.Caption = t_szCAPTION_INSERT
    mnuUpdateInsert.Caption = t_szCAPTION_INSERT
    
    t_nFormMode = ADD_MODE
End Sub

Private Sub cmdEditBtn_Click(Index As Integer)
    Select Case Index
    Case TAB_1
    Case TAB_2
    Case TAB_3
    Case TAB_4
    End Select
    
    subEnableAddBtn False
    subEnableEditBtn False
    subEnableCancelBtn True
    cmdUpdateInsertBtn.Caption = t_szCAPTION_UPDATE
    mnuUpdateInsert.Caption = t_szCAPTION_UPDATE
    
    t_nFormMode = EDIT_MODE
End Sub

Private Sub cmdDeleteBtn_Click(Index As Integer)
    MsgBox "Delete Button Code Here"
    
    Select Case Index
    Case TAB_1
    Case TAB_2
    Case TAB_3
    Case TAB_4
    End Select
    
    t_nFormMode = DELETE_MODE
    
    subEnableDeleteBtn False
End Sub

Private Sub cmdOKBtn_Click(Index As Integer)
    Select Case Index
    Case TAB_1
    Case TAB_2
    Case TAB_3
    Case TAB_4
    End Select
    
    subEnableAddBtn True
    subEnableEditBtn True
    subEnableCancelBtn False
    subEnableOKBtn False
    
    cmdUpdateInsertBtn.Caption = t_szCAPTION_UPDATE
    mnuUpdateInsert.Caption = t_szCAPTION_UPDATE
    
    t_nFormMode = IDLE_MODE
End Sub

Private Sub cmdCancelBtn_Click(Index As Integer)
    cmdTBCancel_Click
End Sub

Private Sub cmdUpdateInsertBtn_Click()
    If t_bUpdateTable Then
        cmdUpdateInsertBtn.Enabled = False
        t_bDataChanged = False
        tfnResetScreen
    End If
End Sub

Private Sub cmdRefreshSelectBtn_Click(Index As Integer)
    If Not tfnCancelExit(t_szREFRESH_MESSAGE) Then
        Exit Sub
    End If
    
    MsgBox "Refresh Button Code Here"
    
    Select Case Index
    Case TAB_1
    Case TAB_2
    Case TAB_3
    Case TAB_4
    End Select
    
    subEnableRefreshBtn False
End Sub

Private Sub cmdExitCancelBtn_Click()
    If cmdExitCancelBtn.Caption = t_szCAPTION_EXIT Then
        cmdTBExit_Click
    Else
        cmdTBCancel_Click
    End If
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
    
    Select Case etabMain.CurrTab
    Case TAB_1
    Case TAB_2
    Case TAB_3
    Case TAB_4
    End Select
    
    tfnUnlockRow
    t_bDataChanged = False
    t_nFormMode = IDLE_MODE
    
    tfnResetScreen
    
    cmdAddBtn(etabMain.CurrTab).SetFocus
End Sub

Private Sub cmdTBExit_Click()
    If t_bDataChanged Then
        If Not tfnCancelExit(t_szEXIT_MESSAGE) Then
            Exit Sub
        End If
    End If
    
    Select Case etabMain.CurrTab
    Case TAB_1
    Case TAB_2
    Case TAB_3
    Case TAB_4
    End Select
    
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
    cmdAddBtn_Click etabMain.CurrTab
End Sub

Private Sub mnuEdit_Click()
    cmdEditBtn_Click etabMain.CurrTab
End Sub

Private Sub mnuDelete_Click()
    cmdDeleteBtn_Click etabMain.CurrTab
End Sub

Private Sub mnuUpdateInsert_Click()
    cmdUpdateInsertBtn_Click
End Sub

Private Sub mnuRefreshSelect_Click()
    cmdRefreshSelectBtn_Click etabMain.CurrTab
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
        If Not Screen.ActiveControl.Locked Then
            Screen.ActiveControl.Text = Clipboard.GetText(vbCFText)
        End If
    End If
End Sub

'======================
'form support functions
'======================

'
'Function        : tfnShowStatusBarMessage - show status bar messages
'Passed Variables: message index, Optional general purpose message
'Returns         : none
'
Public Sub tfnSetStatusBarMessage(szStatusMessage As String)
    
    If t_bStartupFlag Then
        Exit Sub
    End If
    
    ffraStatusbar.ForeColor = STANDARD_TEXT_COLOR
    ffraStatusbar.Font.Bold = False
    ffraStatusbar.Caption = szStatusMessage
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
Private Sub tfnResetScreen(Optional vInitial)
    Dim nTab As Integer
    
    On Error Resume Next
    
    If IsMissing(vInitial) Then
        'codes to clear screen
        Select Case etabMain.CurrTab
        Case TAB_1
        Case TAB_2
        Case TAB_3
        Case TAB_4
        End Select
        
        subEnableAddBtn True
        subEnableEditBtn True
        subEnableDeleteBtn False
        subEnableRefreshBtn False
        subEnableUpdateBtn False
        subEnableCancelBtn False
        subEnableOKBtn False
    Else
        For nTab = 0 To TAB_COUNT - 1
            subEnableAddBtn True, nTab
            subEnableEditBtn True, nTab
            subEnableDeleteBtn False, nTab
            subEnableRefreshBtn False, nTab
            subEnableUpdateBtn False
            subEnableCancelBtn False, nTab
            subEnableOKBtn False, nTab
        Next nTab
    End If
    
    cmdUpdateInsertBtn = t_szCAPTION_UPDATE
    mnuUpdateInsert.Caption = t_szCAPTION_UPDATE
    
    frmContext.ButtonEnabled(PRINT_UP) = False
    frmContext.ButtonEnabled(COPY_UP) = False
    
    
    tfnSetStatusBarMessage t_szADD
End Sub

'=========================
'tooltip support functions
'=========================

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
            Set .LocalDatabase = dbLocal
        End With
    End If
End Sub

Private Sub subEnableAddBtn(bOnOff As Boolean, Optional vTab)
    Dim nTab As Integer
    
    If IsMissing(vTab) Then
        nTab = etabMain.CurrTab
    Else
        nTab = vTab
    End If
    
    cmdAddBtn(nTab).Enabled = bOnOff
    mnuAdd.Enabled = bOnOff
End Sub

Private Sub subEnableEditBtn(bOnOff As Boolean, Optional vTab)
    Dim nTab As Integer
    
    If IsMissing(vTab) Then
        nTab = etabMain.CurrTab
    Else
        nTab = vTab
    End If
    
    cmdEditBtn(nTab).Enabled = bOnOff
    mnuEdit.Enabled = bOnOff
End Sub

Private Sub subEnableDeleteBtn(bOnOff As Boolean, Optional vTab)
    Dim nTab As Integer
    
    If IsMissing(vTab) Then
        nTab = etabMain.CurrTab
    Else
        nTab = vTab
    End If
    
    cmdDeleteBtn(nTab).Enabled = bOnOff
    mnuDelete.Enabled = bOnOff
End Sub

Private Sub subEnableRefreshBtn(bOnOff As Boolean, Optional vTab)
    Dim nTab As Integer
    
    If IsMissing(vTab) Then
        nTab = etabMain.CurrTab
    Else
        nTab = vTab
    End If
    
    cmdRefreshSelectBtn(nTab).Enabled = bOnOff
    mnuRefreshSelect.Enabled = bOnOff
End Sub

Private Sub subEnableUpdateBtn(bOnOff As Boolean)
    cmdUpdateInsertBtn.Enabled = bOnOff
    mnuUpdateInsert.Enabled = bOnOff
End Sub

Private Sub subEnableCancelBtn(bOnOff As Boolean, Optional vTab)
    Dim nTab As Integer
    
    If IsMissing(vTab) Then
        nTab = etabMain.CurrTab
    Else
        nTab = vTab
    End If
    
    cmdCancelBtn(nTab).Enabled = bOnOff
    mnuCancel.Enabled = bOnOff
    frmContext.ButtonEnabled(CANCEL_UP) = bOnOff
End Sub

Private Sub subEnableOKBtn(bOnOff As Boolean, Optional vTab)
    Dim nTab As Integer
    
    If IsMissing(vTab) Then
        nTab = etabMain.CurrTab
    Else
        nTab = vTab
    End If
    
    cmdOKBtn(nTab).Enabled = bOnOff
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

Private Sub subSetMenu()
    mnuAdd.Enabled = cmdAddBtn(etabMain.CurrTab).Enabled
    mnuEdit.Enabled = cmdEditBtn(etabMain.CurrTab).Enabled
    mnuDelete.Enabled = cmdDeleteBtn(etabMain.CurrTab).Enabled
    mnuRefreshSelect = cmdRefreshSelectBtn(etabMain.CurrTab).Enabled
    subEnableCancelBtn cmdCancelBtn(etabMain.CurrTab).Enabled
End Sub

Private Sub Form_Resize()
    frmContext.FormResize
End Sub

Private Sub subSetupToolBar()
    With frmContext
        .BeginSetupToolbar Me
'        .AddButton "Add Pro&fit Center", PRFTCNTR_UP
'        .AddButton "Add &G/L Period/Series", GLPD_UP
'        .AddButton "Add Custome&r", CUSTOMER_UP
'        .AddButton "Add &Product", PRODUCT_UP
'        .AddButton "Add Ta&x Use Group", TAXUSE_UP
'        .AddButton "Add G/L &Account", GL_UP
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

Private Sub subEnableSearchButton(ByRef ctrlButton As FactorFrame, ByVal bStatus As Boolean)
    'Eanble a search button

    ctrlButton.Style = 3    ' command button
    ctrlButton.ShowFocusRect = True
    ctrlButton.Enabled = bStatus
    If bStatus Then
        ctrlButton.Picture = frmContext.LoadPicture(SEARCH_UP)
    Else
        ctrlButton.Picture = frmContext.LoadPicture(SEARCH_DOWN)
    End If
End Sub

