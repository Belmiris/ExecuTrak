VERSION 5.00
Object = "{C75015E0-2232-11D3-B440-0060971E99AF}#1.0#0"; "FACTFRM.OCX"
Object = "{3D388220-1F4E-11D3-B440-0060971E99AF}#1.0#0"; "FACTTAB.OCX"
Object = "{01028C21-0000-0000-0000-000000000046}#4.0#0"; "TG32OV.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{478E45E0-5745-11CF-8918-00A02416C765}#1.0#0"; "SQAOTE32.OCX"
Begin VB.Form frmZZSEPIMT 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Automatic Pay Increase Maintenance"
   ClientHeight    =   6060
   ClientLeft      =   1896
   ClientTop       =   2172
   ClientWidth     =   8880
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.6
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
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   516
      Width           =   8880
      _Version        =   65536
      _ExtentX        =   15663
      _ExtentY        =   9144
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
      Begin FACTFRMLib.FactorFrame cmdExit 
         Height          =   396
         HelpContextID   =   15
         Left            =   7470
         TabIndex        =   8
         Top             =   4620
         Width           =   1260
         _Version        =   65536
         _ExtentX        =   2222
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
      Begin FACTTABLib.FactorTab etabMain 
         Height          =   4860
         HelpContextID   =   201
         Left            =   0
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   0
         Width           =   8895
         _Version        =   65536
         _ExtentX        =   15690
         _ExtentY        =   8572
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
         Caption         =   "Pay Increase &Maint.|Pay &Increase Report"
         Begin FACTFRMLib.FactorFrame efraTabBack1 
            Height          =   4548
            Left            =   12
            TabIndex        =   10
            TabStop         =   0   'False
            Top             =   12
            Width           =   8868
            _Version        =   65536
            _ExtentX        =   15642
            _ExtentY        =   8022
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
            BeginProperty PanelFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Begin FACTFRMLib.FactorFrame efraTabFrame1 
               Height          =   3852
               Left            =   96
               TabIndex        =   11
               TabStop         =   0   'False
               Top             =   96
               Width           =   8652
               _Version        =   65536
               _ExtentX        =   15266
               _ExtentY        =   6800
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
               Begin DBTrueGrid.TDBGrid tblSpreadsheet 
                  Height          =   2028
                  HelpContextID   =   1402
                  Left            =   636
                  OleObjectBlob   =   "zzsepimt.frx":0000
                  TabIndex        =   16
                  Top             =   1584
                  Width           =   7380
               End
               Begin FACTFRMLib.FactorFrame efraTop 
                  Height          =   1068
                  Left            =   132
                  TabIndex        =   12
                  TabStop         =   0   'False
                  Top             =   156
                  Width           =   8412
                  _Version        =   65536
                  _ExtentX        =   14838
                  _ExtentY        =   1884
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
                  Begin VB.TextBox txtDesc 
                     Height          =   360
                     HelpContextID   =   1401
                     Left            =   3972
                     MaxLength       =   20
                     TabIndex        =   14
                     Top             =   480
                     Width           =   3495
                  End
                  Begin VB.TextBox txtLevel 
                     Height          =   360
                     HelpContextID   =   1400
                     Left            =   504
                     MaxLength       =   1
                     TabIndex        =   13
                     Top             =   480
                     Width           =   2205
                  End
                  Begin FACTFRMLib.FactorFrame cmdLevel 
                     CausesValidation=   0   'False
                     Height          =   360
                     HelpContextID   =   2901
                     Left            =   2724
                     TabIndex        =   18
                     TabStop         =   0   'False
                     Tag             =   "Run #3"
                     Top             =   480
                     Width           =   396
                     _Version        =   65536
                     _ExtentX        =   698
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
                     Picture         =   "zzsepimt.frx":1326
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
                  Begin FACTFRMLib.FactorFrame cmdDesc 
                     CausesValidation=   0   'False
                     Height          =   360
                     HelpContextID   =   2901
                     Left            =   7476
                     TabIndex        =   19
                     TabStop         =   0   'False
                     Tag             =   "Run #3"
                     Top             =   480
                     Width           =   396
                     _Version        =   65536
                     _ExtentX        =   698
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
                     Picture         =   "zzsepimt.frx":1438
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
                  Begin VB.Label Label3 
                     Caption         =   "Employee Level Description"
                     Height          =   348
                     Left            =   3984
                     TabIndex        =   21
                     Top             =   132
                     Width           =   3420
                  End
                  Begin VB.Label Label2 
                     Caption         =   "Employee Level"
                     Height          =   276
                     Left            =   528
                     TabIndex        =   20
                     Top             =   168
                     Width           =   2088
                  End
               End
            End
            Begin FACTFRMLib.FactorFrame cmdDeleteBtn 
               Height          =   396
               HelpContextID   =   12
               Index           =   0
               Left            =   3036
               TabIndex        =   9
               Top             =   4080
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
            Begin FACTFRMLib.FactorFrame cmdEditBtn 
               Height          =   396
               HelpContextID   =   11
               Index           =   0
               Left            =   1560
               TabIndex        =   4
               Top             =   4080
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
            Begin FACTFRMLib.FactorFrame cmdAddBtn 
               Height          =   396
               HelpContextID   =   10
               Index           =   0
               Left            =   96
               TabIndex        =   3
               Top             =   4080
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
            Begin FACTFRMLib.FactorFrame cmdRefreshSelectBtn 
               Height          =   396
               HelpContextID   =   14
               Index           =   0
               Left            =   4500
               TabIndex        =   6
               Top             =   4080
               Width           =   1308
               _Version        =   65536
               _ExtentX        =   2307
               _ExtentY        =   698
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
               Left            =   7428
               TabIndex        =   7
               Top             =   4080
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
               Caption         =   "Ca&ncel"
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
               HelpContextID   =   16
               Index           =   0
               Left            =   5964
               TabIndex        =   5
               Top             =   4080
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
         Begin FACTFRMLib.FactorFrame efraTabBack2 
            Height          =   4572
            Left            =   14460
            TabIndex        =   15
            TabStop         =   0   'False
            Top             =   0
            Width           =   8892
            _Version        =   65536
            _ExtentX        =   15684
            _ExtentY        =   8064
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
            BeginProperty PanelFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Begin FACTFRMLib.FactorFrame efraTabFrame2 
               Height          =   3855
               Left            =   90
               TabIndex        =   17
               TabStop         =   0   'False
               Top             =   75
               Width           =   8655
               _Version        =   65536
               _ExtentX        =   15266
               _ExtentY        =   6800
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
               Begin DBTrueGrid.TDBGrid tblGrid 
                  Height          =   3045
                  HelpContextID   =   1403
                  Left            =   90
                  OleObjectBlob   =   "zzsepimt.frx":154A
                  TabIndex        =   24
                  TabStop         =   0   'False
                  Top             =   165
                  Width           =   8475
               End
            End
            Begin FACTFRMLib.FactorFrame cmdExitCancelBtn 
               Height          =   396
               HelpContextID   =   15
               Index           =   1
               Left            =   7440
               TabIndex        =   30
               Top             =   4080
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
               Caption         =   "Ca&ncel"
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
               HelpContextID   =   16
               Index           =   1
               Left            =   6024
               TabIndex        =   29
               Top             =   4080
               Width           =   1308
               _Version        =   65536
               _ExtentX        =   2307
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
            Begin FACTFRMLib.FactorFrame cmdPrint 
               Height          =   396
               HelpContextID   =   1405
               Index           =   1
               Left            =   4464
               TabIndex        =   27
               Top             =   4080
               Width           =   1428
               _Version        =   65536
               _ExtentX        =   2514
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
               Caption         =   "Print &Report"
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
            Begin FACTFRMLib.FactorFrame cmdProcess 
               Height          =   396
               HelpContextID   =   1404
               Index           =   1
               Left            =   3036
               TabIndex        =   25
               Top             =   4080
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
               Caption         =   "&Process"
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
   Begin FACTFRMLib.FactorFrame ffraStatusbar 
      Height          =   360
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   5700
      Width           =   8880
      _Version        =   65536
      _ExtentX        =   15663
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
      CaptionPos      =   1
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
      Bindings        =   "zzsepimt.frx":42BD
      Height          =   2484
      Left            =   408
      OleObjectBlob   =   "zzsepimt.frx":42DC
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   684
      Width           =   3756
   End
   Begin SQAOTestObjectsCtl.SQAOTest SQAOTest1 
      Height          =   456
      Left            =   10896
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   96
      Width           =   456
      _ExtentX        =   804
      _ExtentY        =   804
   End
   Begin FACTFRMLib.FactorFrame efraToolBar 
      Height          =   465
      Left            =   0
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   0
      Width           =   8880
      _Version        =   65536
      _ExtentX        =   15663
      _ExtentY        =   820
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
      FMName          =   "ZZSEPIMT"
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
         TabIndex        =   23
         Top             =   84
         Width           =   6876
         _ExtentX        =   12129
         _ExtentY        =   656
         ButtonWidth     =   635
         ButtonHeight    =   572
         _Version        =   393216
      End
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
'===============================
'track status of data processing
'===============================
Private nDataStatus As Integer
Private Const DATA_INIT = 0
Private Const DATA_LOADED = 1
Private Const DATA_CHANGED = 2
'
'true grid class
Private cValidate As cValidateInput
Private tgmSpreadsheet As clsTGSpreadSheet
Public tgmGrid As clsTGSpreadSheet
Private strHourCode As String
Private strOverCode As String
'===============================
'spreadsheet field
'===============================
Private Const COL_DAYS = 0
Private Const COL_RATE = 1
Private nComboForAdd As Integer
Private nComboForEdit As Integer
'tblGrid
Private Const COL_EMPNO = 0
Private Const COL_NAME = 1
Private Const COL_LEVEL = 2
Private Const COL_DESC = 3
Private Const COL_HIRE_DATE = 4
Private Const COL_DAYS_WORKED = 5
Private Const COL_OLD_RATE = 6
Private Const COL_NEW_RATE = 7
Private Const COL_DIFF = 9
Public rsReport As Recordset

Private Const DB_ERR_READ = "Read database error. Contact Factor."


Private Sub subSetupCombos()

    Set tgcDropdown = CreateObject(t_szOLECOMBO)
    Set tgcDropdown.DBEngine = t_engFactor
    Set tgcDropdown.Form = Me
    Set tgcDropdown.Database = t_dbMainDatabase
    Set tgcDropdown.DataLink = datComboDropDown
    Set tgcDropdown.Table = tblComboDropdown
    'Setup the combos
    ' Set the general properties of the combos
    nComboForAdd = tgcDropdown.AddCombo
    
    tgcDropdown.SQL = "SELECT DISTINCT bg_emp_level,bg_grade_desc FROM bonus_grades"
    tgcDropdown.AddComboBox txtLevel, cmdLevel, "bg_emp_level", tgcDropdown.SQL_INT_TYPE
    tgcDropdown.AddComboBox txtDesc, cmdDesc, "bg_grade_desc", tgcDropdown.SQL_String_TYPE(20)
    
    nComboForEdit = tgcDropdown.AddCombo
    tgcDropdown.SQL = "SELECT DISTINCT zzsepi_emp_level, zzsepi_lvl_desc FROM zzse_pay_increase"
    tgcDropdown.AddComboBox txtLevel, cmdLevel, "zzsepi_emp_level", tgcDropdown.SQL_INT_TYPE
    tgcDropdown.AddComboBox txtDesc, cmdDesc, "zzsepi_lvl_desc", tgcDropdown.SQL_String_TYPE(20)
        

End Sub
Private Sub subInitializeSpreadSheet()
    Dim szDecPattern As String
    szDecPattern = tfnDecimalPattern(8, 2, False, False)

    Set tgmSpreadsheet = New clsTGSpreadSheet
    Set tgmSpreadsheet.Table = tblSpreadsheet
    Set tgmSpreadsheet.StatusBar = ffraStatusbar
    Set tgmSpreadsheet.Form = Me

    'tgmSpreadsheet.AllowRevalidate = True
    'Add editable column and column value validation
    tgmSpreadsheet.AddEditColumn COL_DAYS, "Enter Days Employed", szIntegerPattern
    tgmSpreadsheet.AddEditColumn COL_RATE, "Enter Hourly Pay Rate", szDecPattern, "#,###,##0.00"

'    tgmSpreadsheet.ColumnForNewRow = 0


    With tblSpreadsheet.Columns(COL_DAYS)
        .DataField = "zzsepi_days"
    End With
    With tblSpreadsheet.Columns(COL_RATE)
        .DataField = "zzsepi_pay"
    End With

    'tblSpreadsheet.EditDropDown = True

End Sub
Private Sub subSetupMousePointer()
    Set objCursor = New clsCursor
    Set objCursor.MainForm = Me
    
    objCursor.AddControls tblComboDropdown, frmAbout
    
    objCursor.AddScreen TAB_1
    objCursor.AddArea efraTabFrame1
    objCursor.AddScreen TAB_2
    objCursor.AddArea efraTabFrame2

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
    cValidate.ESCControl = cmdExit
    cValidate.ESCControl = cmdExitCancelBtn(0)
    
    cValidate.MinTabIndex = 0
    Set cValidate.ControlForFocus = efraTabFrame1
    Set cValidate.LastBox = txtLevel
    efraTabFrame1.TabIndex = txtDesc.TabIndex + 1
    efraTabFrame1.TabStop = True
    cValidate.SetFirstControls cmdAddBtn(0), cmdEditBtn(0), cmdDeleteBtn(0), cmdRefreshSelectBtn(0), cmdUpdateInsertBtn(0), cmdExitCancelBtn(0)
End Sub

Public Function fnInvalidData(txtBox As Textbox) As Boolean
    Select Case txtBox.TabIndex
        Case txtLevel.TabIndex
            fnInvalidData = Not fnValidNumer
        Case txtDesc.TabIndex
            fnInvalidData = False 'Not fnValidNumer
    End Select
End Function



Private Sub cmdDesc_Click()
    tgcDropdown.Click cmdDesc
End Sub

Private Sub cmdExit_Click()
    cmdTBExit_Click
End Sub

Private Sub cmdExit_GotFocus()
    tfnSetStatusBarMessage "Exit program"
End Sub

Private Sub cmdLevel_Click()
    tgcDropdown.Click cmdLevel
End Sub



Private Sub cmdPrint_Click(Index As Integer)
   subPrintRpt
    
End Sub

Private Sub cmdPrint_GotFocus(Index As Integer)
    tfnSetStatusBarMessage "Print a report of proposed pay increases "
End Sub

Private Sub cmdProcess_Click(Index As Integer)
    Dim rsTemp As Recordset
    tblGrid.Enabled = True
    #If PROTOTYPE Then
        Exit Sub
    #End If
    etabMain.TabEnabled(TAB_1) = False
    subEnableCancelBtn True
    cmdProcess(1).Enabled = False
    subEnableUpdateBtn False
    
    If fnProcess(rsReport) Then
        subEnableUpdateBtn True
        cmdPrint(1).Enabled = True
        cmdPrint(1).SetFocus
    End If
    
End Sub

Private Sub cmdProcess_GotFocus(Index As Integer)
    tfnSetStatusBarMessage "Find employees who need pay raises"
End Sub

Private Sub efraTabFrame1_GotFocus()
    cValidate.GotFocus efraTabFrame1
End Sub

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



Private Sub mnuOptions_Click()
    If etabMain.CurrTab = TAB_2 Then
        mnuAdd.Enabled = False
        mnuEdit.Enabled = False
    End If
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
    If etabMain.CurrTab = TAB_1 Then
        subSetFocus cmdAddBtn(0)
    Else
        cmdProcess(1).Enabled = True
        cmdProcess(1).SetFocus
    End If
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
    cmdPrint(1).Enabled = False
    subEnableUpdateBtn False
    subEnableCancelBtn True, 0
    subEnableCancelBtn True, 1
    Select Case NewTab
         Case TAB_1
            subEnableDeleteBtn False, 0
            
            subEnableRefreshBtn False, 0
            
         Case TAB_2
            cmdProcess(1).Enabled = True
            subEnableCancelBtn True
            
    End Select
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
#If Not PROTOTYPE Then
    If Not fnGetPayCode Then
        MsgBox "System Error: Unable to access system parameter, Program will be terminated", vbExclamation
        Unload Me
        Exit Sub
    End If
#End If
    subSetupCombos  'Set up the combo dropdowns
    subInitializeSpreadSheet 'set up true grid
    subInitGrid
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
    
    'subSetExitCancelBtn "EXIT"
    
    
    Screen.MousePointer = vbHourglass
    
    subSetupMousePointer
    t_bStartupFlag = False
    Screen.MousePointer = vbHourglass
    
    ' ^ ^ ^ ^ ^ ^ ^
    ' | | | | | | |
    '***************************************************
    
    tmrKeyBoard.Enabled = True
    tfnResetScreen "Initail"  'set the default screen
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
Private Sub cmdAddBtn_GotFocus(Index As Integer)
    tfnSetStatusBarMessage t_szADD
End Sub

Private Sub cmdEditBtn_GotFocus(Index As Integer)
    tfnSetStatusBarMessage t_szEDIT
End Sub

Private Sub cmdDeleteBtn_GotFocus(Index As Integer)
    tfnSetStatusBarMessage t_szDELETE
End Sub

Private Sub cmdUpdateInsertBtn_GotFocus(Index As Integer)
    If cmdUpdateInsertBtn(Index).Caption = t_szCAPTION_INSERT Then
        tfnSetStatusBarMessage t_szINSERT
    Else
        tfnSetStatusBarMessage t_szUPDATE
    End If
End Sub

Private Sub cmdRefreshSelectBtn_GotFocus(Index As Integer)
    tfnSetStatusBarMessage t_szREFRESH
End Sub

Private Sub cmdExitCancelBtn_GotFocus(Index As Integer)
    If cmdExitCancelBtn(etabMain.CurrTab).Caption = t_szCAPTION_EXIT Then
        tfnSetStatusBarMessage t_szEXIT
    Else
        tfnSetStatusBarMessage t_szCANCEL
    End If
End Sub

Private Sub cmdAddBtn_Click(Index As Integer)
    #If PROTOTYPE Then
        Exit Sub
    #End If
    
    etabMain.TabEnabled(TAB_2) = False
    efraTabFrame1.Enabled = True
    
    subEnableAddBtn False
    subEnableEditBtn False
    subEnableCancelBtn True
    
    cmdUpdateInsertBtn(0).Caption = t_szCAPTION_INSERT
    mnuUpdateInsert.Caption = t_szCAPTION_INSERT
    
    t_nFormMode = ADD_MODE
    'enable combo dropdown
    tgcDropdown.ComboActive(nComboForAdd) = True
    tgcDropdown.ComboActive(nComboForEdit) = False
    tgmSpreadsheet.AllowAddNew = True
    tgmSpreadsheet.ClearData
    
    subEnableTop True
    
    cValidate.ResetFlags
    txtLevel.SetFocus

End Sub

Private Sub cmdEditBtn_Click(Index As Integer)
    #If PROTOTYPE Then
        Exit Sub
    #End If

    etabMain.TabEnabled(TAB_2) = False
    subEnableAddBtn False
    subEnableEditBtn False
    subEnableCancelBtn True
    cmdUpdateInsertBtn(etabMain.CurrTab).Caption = t_szCAPTION_UPDATE
    mnuUpdateInsert.Caption = t_szCAPTION_UPDATE
    tgmSpreadsheet.AllowAddNew = False
    tgmSpreadsheet.ClearData
    t_nFormMode = EDIT_MODE
    'enable combo dropdown
    tgcDropdown.ComboActive(nComboForAdd) = False
    tgcDropdown.ComboActive(nComboForEdit) = True
    
    subEnableTop True
    cValidate.ResetFlags
    txtLevel.SetFocus
    
End Sub

Private Sub cmdDeleteBtn_Click(Index As Integer)
    Dim strSQL As String
    Dim sMsg As String
    Dim sCode As String
    Dim lRow As Long
    Dim strType As String
       
    sMsg = "Are you sure you want to delete the current record?"
    
    If Not tfnConfirm(sMsg, vbDefaultButton2) Then
        Exit Sub
    End If
    
    lRow = tgmSpreadsheet.GetCurrentRowNumber
    sCode = fnGetField(tgmSpreadsheet.CellValue(COL_DAYS, lRow))
    
    If sCode = "" Then
        Exit Sub
    End If
    If t_nFormMode = EDIT_MODE Then
        
        strSQL = "DELETE FROM zzse_pay_increase "
        strSQL = strSQL & " WHERE zzsepi_emp_level = " & Val(txtLevel.Text) _
                       & " AND zzsepi_days= " & sCode & ""
        If Not fnExecuteSQL(strSQL, "Delete_Click") Then
            Exit Sub
        End If
        If tgmSpreadsheet.RowCount = 1 Then
            tfnResetScreen
        Else
            tgmSpreadsheet.DeleteRow tgmSpreadsheet.Bookmark(lRow)
        End If
    Else
        tgmSpreadsheet.DeleteRow tgmSpreadsheet.Bookmark(lRow)
    End If
    
    subEnableDeleteBtn False
End Sub

'Private Sub cmdOKBtn_Click(Index As Integer)
'    Select Case Index
'    Case TAB_1
'    Case TAB_2
'    Case TAB_3
'    Case TAB_4
'    End Select
'
'    subEnableAddBtn True
'    subEnableEditBtn True
'    subEnableCancelBtn False
'    subEnableOKBtn False
'
'    cmdUpdateInsertBtn.Caption = t_szCAPTION_UPDATE
'    mnuUpdateInsert.Caption = t_szCAPTION_UPDATE
'
'    t_nFormMode = IDLE_MODE
'End Sub

Private Sub cmdCancelBtn_Click(Index As Integer)
    cmdTBCancel_Click
End Sub

Private Sub cmdUpdateInsertBtn_Click(Index As Integer)
    Select Case Index
        Case TAB_1
            If fnUpdateInsert1 Then
                tfnResetScreen
            End If
        
        Case TAB_2
            If fnUpdateInsert2 Then
                tfnResetScreen
            End If
    End Select
    cmdProcess(1).Enabled = True
    subEnableCancelBtn True, Index
    cmdUpdateInsertBtn(Index).Enabled = False
End Sub

Private Sub cmdRefreshSelectBtn_Click(Index As Integer)
    If Not tfnCancelExit(t_szREFRESH_MESSAGE) Then
        Exit Sub
    End If
    
    subEnableUpdateBtn False
    subEnableDeleteBtn True
    subEnableRefreshBtn False
    
    nDataStatus = DATA_INIT
    subLoadData
End Sub

Private Sub cmdExitCancelBtn_Click(Index As Integer)
    If cmdExitCancelBtn(etabMain.CurrTab).Caption = t_szCAPTION_EXIT Then
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
    
    tfnUnlockRow
    t_bDataChanged = False
    t_nFormMode = IDLE_MODE
    
    tfnResetScreen
    Select Case etabMain
     Case TAB_1
        cmdAddBtn(etabMain.CurrTab).SetFocus
     Case TAB_2
        cmdProcess(1).Enabled = True
        cmdProcess(1).SetFocus
    End Select
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
    cmdAddBtn_Click etabMain.CurrTab
End Sub

Private Sub mnuEdit_Click()
    cmdEditBtn_Click etabMain.CurrTab
End Sub

Private Sub mnuDelete_Click()
    cmdDeleteBtn_Click etabMain.CurrTab
End Sub

Private Sub mnuUpdateInsert_Click()
    cmdUpdateInsertBtn_Click etabMain.CurrTab
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
    
    If Not etabMain.TabEnabled(TAB_1) Then
        etabMain.TabEnabled(TAB_1) = True
    End If
    etabMain.Visible = True
    subEnableAllControls True

    t_bDataChanged = False
    subEnableAddBtn True
    subEnableEditBtn True
    subEnableDeleteBtn False
    subEnableRefreshBtn False
    subEnableUpdateBtn False
    subEnableCancelBtn False
    subEnableOKBtn False
    cmdPrint(1).Enabled = False
  
    tblGrid.Enabled = False
    tblSpreadsheet.Enabled = False
    txtLevel.Enabled = False
    txtDesc.Enabled = False
    subEnableSearchButton cmdLevel, False
    subEnableSearchButton cmdDesc, False
    
    cmdUpdateInsertBtn(0) = t_szCAPTION_UPDATE
    mnuUpdateInsert.Caption = t_szCAPTION_UPDATE
    
    frmContext.ButtonEnabled(PRINT_UP) = False
    frmContext.ButtonEnabled(COPY_UP) = False
    
    'clear data
    subClearAll
    
    nDataStatus = DATA_INIT
    If etabMain.CurrTab = TAB_1 Then
        subSetFocus cmdAddBtn(0)
        tfnSetStatusBarMessage t_szADD
    Else
        cmdProcess(1).Enabled = True
        cmdProcess(1).SetFocus
        tfnSetStatusBarMessage "Find employees who need pay raises"
    End If
End Sub


Private Sub tblGrid_FirstRowChange()
    tgmGrid.FirstRowChange
End Sub

Private Sub tblGrid_GotFocus()
tfnSetStatusBarMessage "View Grid Data. Total # of Rows: " & tgmGrid.RowCount
End Sub

Private Sub tblGrid_SelChange(CANCEL As Integer)
    CANCEL = True
End Sub

Private Sub tblGrid_UnboundReadData(ByVal RowBuf As DBTrueGrid.RowBuffer, StartLocation As Variant, ByVal ReadPriorRows As Boolean)
    tgmGrid.ReadData RowBuf, StartLocation, ReadPriorRows
End Sub

Private Sub tblSpreadsheet_AfterColEdit(ByVal ColIndex As Integer)
    tgmSpreadsheet.AfterColEdit ColIndex
    nDataStatus = DATA_CHANGED
    t_bDataChanged = True
    subSetButtonStatus
End Sub

Private Sub tblSpreadsheet_BeforeColEdit(ByVal ColIndex As Integer, ByVal KeyAscii As Integer, CANCEL As Integer)
    subEnableUpdateBtn False
    tgmSpreadsheet.BeforeColEdit ColIndex, KeyAscii, CANCEL
End Sub

Private Sub tblSpreadsheet_Change()
    tgmSpreadsheet.Change
End Sub



Private Sub tblSpreadsheet_FirstRowChange()
    tgmSpreadsheet.FirstRowChange
    subCheckDelete
End Sub

Private Sub tblSpreadsheet_GotFocus()
    tgmSpreadsheet.GotFocus
End Sub

Private Sub tblSpreadsheet_KeyDown(KeyCode As Integer, Shift As Integer)
    tgmSpreadsheet.KeyDown KeyCode, Shift
End Sub

Private Sub tblSpreadsheet_KeyPress(KeyAscii As Integer)
    If Not tgmSpreadsheet.Keypress(KeyAscii) Then
        KeyAscii = 0
    End If
'    If KeyAscii <> vbKeyReturn Then
'        If tblSpreadsheet.col = COL_DAYS Then
'            tgmSpreadsheet.CellValue(COL_RATE, tgmSpreadsheet.GetCurrentRowNumber) = ""
'
'        End If
'    End If
End Sub

Private Sub tblSpreadsheet_LostFocus()
    tgmSpreadsheet.LostFocus
End Sub

Private Sub tblSpreadsheet_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    'Dim lRow As Long
    'lRow = tgmSpreadsheet.GetCurrentRowNumber
    tgmSpreadsheet.RowColChange LastRow, LastCol
   ' If tblSpreadsheet.col = COL_DAYS Then
   '     If tgmSpreadsheet.ValidCell(COL_DAYS, lRow) Then
            'tgmSpreadsheet.GotoNextColumn LastCol
   '     End If

    'End If
    
    subCheckDelete
End Sub

Private Sub tblSpreadsheet_SelChange(CANCEL As Integer)
     CANCEL = True
End Sub

Private Sub tblSpreadsheet_UnboundReadData(ByVal RowBuf As DBTrueGrid.RowBuffer, StartLocation As Variant, ByVal ReadPriorRows As Boolean)
    tgmSpreadsheet.ReadData RowBuf, StartLocation, ReadPriorRows
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
    
    cmdDeleteBtn(0).Enabled = bOnOff
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
    cmdUpdateInsertBtn(0).Enabled = bOnOff
    cmdUpdateInsertBtn(1).Enabled = bOnOff
    mnuUpdateInsert.Enabled = bOnOff
End Sub

Private Sub subEnableCancelBtn(bOnOff As Boolean, Optional vTab)
    Dim nTab As Integer
    
    If IsMissing(vTab) Then
        nTab = etabMain.CurrTab
    Else
        nTab = vTab
    End If
    
    cmdExitCancelBtn(nTab).Enabled = bOnOff
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
    
    cmdUpdateInsertBtn(nTab).Enabled = bOnOff
End Sub

'Private Sub subSetExitCancelBtn(sExitCancel As String)
'    If sExitCancel = "EXIT" Then
'        cmdExitCancelBtn(Index).Caption = t_szCAPTION_EXIT
'        frmContext.ButtonEnabled(CANCEL_UP) = False
'        mnuCancel.Enabled = False
'    Else
'        cmdExitCancelBtn(Index).Caption = t_szCAPTION_CANCEL
'        frmContext.ButtonEnabled(CANCEL_UP) = True
'        mnuCancel.Enabled = True
'    End If
'End Sub

Private Sub subSetMenu()
    mnuAdd.Enabled = cmdAddBtn(0).Enabled
    mnuEdit.Enabled = cmdEditBtn(0).Enabled
    mnuDelete.Enabled = cmdDeleteBtn(0).Enabled
    mnuRefreshSelect = cmdRefreshSelectBtn(0).Enabled
    subEnableCancelBtn cmdExitCancelBtn(etabMain.CurrTab).Enabled
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

Private Sub txtDesc_Change()
    tgcDropdown.Change txtDesc
    cValidate.Change txtDesc

End Sub

Private Sub txtDesc_Click()
    tgcDropdown.Click txtDesc
End Sub

Private Sub txtDesc_GotFocus()
    tfnSetStatusBarMessage "Enter Employ Level Description"
    tgcDropdown.GotFocus txtDesc
    cValidate.GotFocus txtDesc

    If tgcDropdown.SingleRecordSelected Then
        subLoadData
    End If
End Sub

Private Sub txtDesc_KeyPress(KeyAscii As Integer)
    tgcDropdown.Keypress txtDesc, KeyAscii
'    If t_nFormMode = EDIT_MODE Then
        If KeyAscii = vbKeyReturn Then
            If tgcDropdown.SingleRecordSelected Then
                subLoadData
            End If
        End If
'    Else
'        If KeyAscii = vbKeyReturn Then
'            tblSpreadsheet.Enabled = True
'            subSetFocus tblSpreadsheet
'        End If
'    End If
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtDesc_LostFocus()
'    If t_nFormMode = EDIT_MODE Then
        If cValidate.LostFocus(txtLevel, cmdLevel, txtDesc, cmdDesc, tblComboDropdown) Then
            If cValidate.ValidInput(txtLevel) Then
                tgcDropdown.LostFocus txtLevel
                subLoadData
            End If
        End If

'    End If
    
End Sub



Private Sub txtLevel_Change()
    tfnRegExpControlChange txtLevel, szIntegerPattern

    tgcDropdown.Change txtLevel
    cValidate.Change txtLevel
    
    If nDataStatus = DATA_LOADED Then
        nDataStatus = DATA_CHANGED
        t_bDataChanged = True
    End If
    
End Sub

Private Sub txtLevel_GotFocus()
'    tfnSetStatusBarMessage "Enter From Location Code "
    tfnRegExpControlGotFocus txtLevel, szIntegerPattern
    
    tgcDropdown.GotFocus txtLevel
    cValidate.GotFocus txtLevel

    If tgcDropdown.SingleRecordSelected Then
        subLoadData
    End If

End Sub

Private Sub txtLevel_KeyPress(KeyAscii As Integer)
    Dim bCode As Boolean
    
'    If KeyAscii = vbKeyReturn Then
'        subSetBusyState True
'    End If
'
    If Not tfnRegExpControlKeyPress(txtLevel, KeyAscii, szIntegerPattern) Then
        KeyAscii = 0
    End If
 
    bCode = tgcDropdown.Keypress(txtLevel, KeyAscii)

    If Not bCode Then
        If KeyAscii = vbKeyReturn Then
            If tgcDropdown.SingleRecordSelected Then
                subLoadData
            End If
        End If
        KeyAscii = 0
    Else
        cValidate.Keypress txtLevel, KeyAscii
        
    End If
'    subSetBusyState False
End Sub

Private Sub txtLevel_LostFocus()
    If cValidate.LostFocus(txtLevel, cmdLevel, txtDesc, cmdDesc, tblComboDropdown) Then
        If cValidate.ValidInput(txtLevel) Then
            tgcDropdown.LostFocus txtLevel
            subLoadData
        End If
    End If
End Sub
Private Sub subLoadData()
    Dim strSQL As String
    Dim rsTemp As Recordset
    If nDataStatus >= DATA_LOADED Then
        Exit Sub
    End If
    
    If cValidate.FirstInvalidInput >= 0 Then
        Exit Sub
    End If
    
    
    If t_nFormMode = EDIT_MODE Then
        tfnSetStatusBarCorrect "Loading Data ... Please wait ..."
        strSQL = "SELECT * FROM zzse_pay_increase WHERE zzsepi_emp_level=" & Val(txtLevel.Text)
        'lock table to do edit
        If Not tfnLockRow("ZZSEPIMT", "zzse_pay_increase", strSQL) Then
            tfnSetStatusBarCorrect ""
            Exit Sub
        End If
    End If
    
    nDataStatus = DATA_LOADED
    
    subEnableTop False

    tblSpreadsheet.Enabled = True
    
    If t_nFormMode = EDIT_MODE Then
        'fill main spreadsheet with data
        strSQL = "SELECT zzsepi_days, zzsepi_pay " & _
        " FROM zzse_pay_increase WHERE zzsepi_emp_level = " & Val(txtLevel.Text) & " ORDER BY zzsepi_days "
        tgmSpreadsheet.FillWithSQL t_dbMainDatabase, strSQL
    End If
    
    DoEvents
    subSetFocus tblSpreadsheet

End Sub
Private Sub subEnableTop(bTrueFalse As Boolean)
    txtLevel.Enabled = bTrueFalse
    txtDesc.Enabled = bTrueFalse
    subEnableSearchButton cmdLevel, bTrueFalse
    subEnableSearchButton cmdDesc, bTrueFalse

End Sub
Private Sub subSetFocus(v As Control)
    On Error Resume Next
    v.SetFocus
End Sub
Private Sub subClearAll()

    txtLevel.Text = ""
    txtDesc.Text = ""
    
    tgmGrid.ClearData
    tgmGrid.ResetFlags
    tgmSpreadsheet.ClearData
    tgmSpreadsheet.ResetFlags
    
End Sub
Private Sub subSetButtonStatus()
    Dim bStatus As Boolean
    
    bStatus = (tgmSpreadsheet.ValidData And (tgmSpreadsheet.RowCount > 0))
    
    If t_nFormMode = EDIT_MODE Then
        subEnableRefreshBtn True
        'subEnableDelete False
    End If
        
    subEnableUpdateBtn bStatus
End Sub
Private Sub subCheckDelete()
    Dim lRow As Long
    
    If tblSpreadsheet.Row < 0 Then
        subEnableDeleteBtn False
    Else
        lRow = tgmSpreadsheet.GetCurrentRowNumber
        If lRow < 0 Then
            subEnableDeleteBtn False
        Else
            If fnGetField(tgmSpreadsheet.CellValue(COL_DAYS, lRow)) <> "" Then
                subEnableDeleteBtn True
            Else
                subEnableDeleteBtn False
            End If
        End If
    End If
End Sub
Public Sub subEnableAllControls(bStatus As Boolean)

    efraTabFrame1.Enabled = bStatus
    efraTabFrame2.Enabled = bStatus
    efraTabBack1.Enabled = bStatus
    efraTabBack2.Enabled = bStatus
    etabMain.TabEnabled(TAB_2) = bStatus
    etabMain.TabEnabled(TAB_1) = bStatus
    efraTop.Enabled = bStatus
    tgmSpreadsheet.Locked = Not bStatus
End Sub
Private Function fnGetField(x As Variant) As String
     If IsNull(x) Then
        fnGetField = ""
     Else
        fnGetField = Trim(x)
     End If
End Function
Public Function fnValidCellValue(tgTable As Object, ByVal nCurrCol As Integer, ByVal lCurrRow As Long, szText As String) As Boolean

    fnValidCellValue = True
    Select Case nCurrCol
        Case COL_DAYS
            fnValidCellValue = fnValidNumberForAdd(szText, lCurrRow)
        Case COL_RATE
            If IsNumeric(szText) Then
                fnValidCellValue = True
            ElseIf Trim(szText) = "" Then
                tgmSpreadsheet.ErrorMessage(nCurrCol) = "Hourly Par Rate cannot be blank"
                fnValidCellValue = False
            ElseIf Trim(szText) = "" Then
                tgmSpreadsheet.ErrorMessage(nCurrCol) = "Hourly Par Rate must be numeric"
                fnValidCellValue = False
            End If
    End Select
End Function
Private Function fnValidNumberForAdd(sText As String, ByVal lCurrRow As Long) As Boolean
    Const SUB_NAME = "fnValidNumberForAdd"
    Dim sCode As String
    Dim sLogin As String
    Dim sName As String
    Dim vTemp() As Variant
    Dim lRows As Long
    Dim strSQL As String
    Dim rsTemp As Recordset

    fnValidNumberForAdd = False
    If Trim$(sText) = "" Then
        tgmSpreadsheet.ErrorMessage(COL_DAYS) = "Days Employed cannot be blank."
        Exit Function
    Else
        sCode = tgmSpreadsheet.CellValue(COL_DAYS, lCurrRow)

    'check database if duplicated data in database
        strSQL = "SELECT zzsepi_days FROM zzse_pay_increase WHERE zzsepi_days =" & sCode & " AND zzsepi_emp_level=" & Val(txtLevel.Text)
        If GetRecordSet(rsTemp, strSQL) < 0 Then
            tgmSpreadsheet.ErrorMessage(COL_DAYS) = DB_ERR_READ
            Exit Function
        Else
            If rsTemp.RecordCount > 0 Then
                tgmSpreadsheet.ErrorMessage(COL_DAYS) = "Same Days Employed already exists in database."
                Exit Function
            End If
        End If
     'check screen to find duplicated data
        Dim I As Integer
        For I = 0 To tgmSpreadsheet.RowCount - 1
            If Trim(tgmSpreadsheet.CellValue(COL_DAYS, I)) = sCode And I <> lCurrRow Then
                tgmSpreadsheet.ErrorMessage(COL_DAYS) = "Duplicated Days Employed encountered."
                Exit Function
            End If
        Next I
        'tgmSpreadsheet.BeforeColEdit COL_DAYS, 0, False
        tgmSpreadsheet.BeforeColEdit COL_RATE, 0, False
        fnValidNumberForAdd = True
    End If
End Function
' this function also returns record count
Public Function GetRecordSet(rsTemp As Recordset, szSQL As String, _
                   Optional szCalledFrom As Variant, _
                   Optional bShowErrow As Variant) As Long
    On Error GoTo SQLError
    Set rsTemp = t_dbMainDatabase.OpenRecordset(szSQL, dbOpenSnapshot, dbSQLPassThrough)
    
    
    If rsTemp.RecordCount > 0 Then
       rsTemp.MoveLast
       rsTemp.MoveFirst
    End If
    GetRecordSet = rsTemp.RecordCount
    Exit Function
SQLError:
    If IsMissing(szCalledFrom) Then
       szCalledFrom = ""
    End If
    If IsMissing(bShowErrow) Then
       bShowErrow = True
    End If
    GetRecordSet = -1
    tfnErrHandler "GetRecordSet," & szCalledFrom, szSQL, bShowErrow
    On Error GoTo 0
End Function

Private Function fnUpdateInsert1() As Boolean
    Const SUB_NAME = "fnUpdateInsert1"
    Dim strSQL As String
    Dim vData() As Variant
    Dim lRowCount As Long
    Dim nColCount As Integer
    Dim I As Long
    
    Screen.MousePointer = vbHourglass
    'delete the data of same account number
    If t_nFormMode = EDIT_MODE Then
        strSQL = "DELETE FROM zzse_pay_increase WHERE zzsepi_emp_level=" & Val(Trim(txtLevel.Text))
        If Not fnExecuteSQL(strSQL, SUB_NAME) Then
            Screen.MousePointer = vbDefault
            Exit Function
        End If
        lRowCount = tgmSpreadsheet.RowCount
    Else
    
    lRowCount = tgmSpreadsheet.GetChangedRowCount
    End If
    
    'insert true grid sceen data into database
    For I = 0 To lRowCount - 1
        If t_nFormMode = EDIT_MODE Then
            tgmSpreadsheet.GetRow vData, nColCount, I
        Else
            tgmSpreadsheet.GetChangedRow vData, nColCount, I
        End If
        If fnGetField(vData(COL_DAYS)) <> "" Then

            strSQL = "INSERT INTO zzse_pay_increase (zzsepi_emp_level,zzsepi_lvl_desc, zzsepi_days, zzsepi_pay ) VALUES(" _
                & Val(txtLevel.Text) & ",'" & txtDesc.Text & "'," _
               & vData(COL_DAYS) & "," & vData(COL_RATE) & " )"
            If Not fnExecuteSQL(strSQL, SUB_NAME) Then
                Screen.MousePointer = vbDefault
                Exit Function
            End If
        End If
    Next
        
    fnUpdateInsert1 = True
    Screen.MousePointer = vbDefault
End Function
Private Function fnUpdateInsert2() As Boolean
    Const SUB_NAME = "fnUpdateInsert2"
    Dim strSQL As String
    Dim vData() As Variant
    Dim lRowCount As Long
    Dim nColCount As Integer
    Dim I As Long
    
    Screen.MousePointer = vbHourglass

    'lRowCount = tgmGrid.GetChangedRowCount
    
    lRowCount = tgmGrid.RowCount
    'insert true grid sceen data into database
    For I = 0 To lRowCount - 1
        'tgmGrid.GetChangedRow vData, nColCount, i
        tgmGrid.GetRow vData, nColCount, I
        If fnGetField(vData(COL_EMPNO)) <> "" Then

            strSQL = "UPDATE pr_emp_pay SET prep_amount= " _
               & vData(COL_NEW_RATE)
            strSQL = strSQL & " WHERE prep_empno=" & vData(COL_EMPNO)
            strSQL = strSQL & " AND prep_pay_code='" & strHourCode & "'"
            If Not fnExecuteSQL(strSQL, SUB_NAME) Then
                Screen.MousePointer = vbDefault
                Exit Function
            End If
            strSQL = "UPDATE pr_emp_pay SET prep_amount= " _
               & vData(COL_NEW_RATE) * 1.5
            strSQL = strSQL & " WHERE prep_empno=" & vData(COL_EMPNO)
            strSQL = strSQL & " AND prep_pay_code='" & strOverCode & "'"
            If Not fnExecuteSQL(strSQL, SUB_NAME) Then
                Screen.MousePointer = vbDefault
                Exit Function
            End If
        End If
    Next
        
    fnUpdateInsert2 = True
    Screen.MousePointer = vbDefault
End Function
Private Function fnExecuteSQL(szSQL As String, _
                Optional szCalledFrom As Variant, Optional bShowError As Variant) As Boolean
                
      Dim szMsg As String
      
      On Error GoTo SQLError
      
      t_dbMainDatabase.ExecuteSQL szSQL
      
      fnExecuteSQL = True
      Exit Function
      
SQLError:
      fnExecuteSQL = False
      If IsMissing(szCalledFrom) Then
         szCalledFrom = ""
      End If
      If IsMissing(bShowError) Then
         bShowError = True
      End If
      tfnErrHandler "fnExecuteSQL, " & szCalledFrom, szSQL, bShowError
      On Error GoTo 0
End Function

Private Function fnProcess(rsTemp As Recordset) As Boolean
    Dim strSQL As String
    Dim vData() As Variant
    Const SUB_NAME As String = "fnProcess"
    fnProcess = False
    On Error Resume Next
    t_dbMainDatabase.ExecuteSQL "Drop table tmpTableReport"
'    On Error GoTo err_handler
    
    
    strSQL = "SELECT DISTINCT prm_empno empno, prm_last_name lname,prm_first_name fname,prm_middle_name mname, prm_emp_level emp_level,zzsepi_lvl_desc emp_level_desc,prm_date_hired date_hired, "
    strSQL = strSQL & " prep_amount old_rate, zzsepi_pay new_rate, zzsepi_days zzsepi_days"
    strSQL = strSQL & " FROM pr_master,pr_emp_pay,zzse_pay_increase "
    strSQL = strSQL & " WHERE prm_job_class='CS' AND (prm_date_termed IS NULL OR prm_date_termed ='')  AND prm_emp_level=zzsepi_emp_level "
    strSQL = strSQL & " AND prm_empno=prep_empno AND prep_pay_code='" & strHourCode & "'"
    strSQL = strSQL & " AND Today-prm_date_hired >0 "
    strSQL = strSQL & " AND zzsepi_pay-prep_amount>0 "
    strSQL = strSQL & " INTO TEMP tmpTableReport"
    
    t_dbMainDatabase.ExecuteSQL strSQL
        
    strSQL = "SELECT DISTINCT empno, lname,fname, mname, emp_level,emp_level_desc, date_hired,  "
    strSQL = strSQL & " Today - date_hired days_worked "
    strSQL = strSQL & ", old_rate, new_rate"
    strSQL = strSQL & ", new_rate-old_rate rate_diff "
    strSQL = strSQL & " FROM tmpTableReport "
    strSQL = strSQL & " WHERE zzsepi_days = (SELECT MAX(b.zzsepi_days) FROM tmpTableReport AS b WHERE emp_level = b.emp_level AND empno = b.empno  AND b.zzsepi_days <= (Today - date_hired))"
    strSQL = strSQL & " ORDER BY empno,emp_level"
    
    If GetRecordSet(rsTemp, strSQL, , SUB_NAME) < 0 Then
        MsgBox "Failed to access database to get records", vbCritical
        Exit Function
    End If
    If rsTemp.RecordCount = 0 Then
        MsgBox "No employee needs to increase pay.", vbExclamation
        Exit Function
    End If
            
    tgmGrid.FillWithRecordset rsTemp
    
    fnProcess = True
End Function
Private Function fnWriteLog(vData As Variant, sProgName As String, sRptTitle As String, sLineHeader As String, Optional sFileName As String = "ZZSEPILG", Optional bAppend As Boolean = False) As Boolean
    Const LINES As Integer = 30
    Dim nFilePtr As Integer

    Dim nPage As Integer, nLine As Integer, nPages As Integer
    fnWriteLog = False
    
    tfnSetStatusBarMessage "Writing Log File: " + sFileName + ". Please wait..."
    
    nFilePtr = FreeFile
    On Error GoTo ErrorWriteFile
    If bAppend Then
        Open sFileName For Append As #nFilePtr
    Else
        Open sFileName For Output As #nFilePtr
    End If
    Dim I As Integer
    I = 0
    'Print #nFilePtr, "Program ID: ZZSEPILG"
    Print #nFilePtr, sProgName
    Print #nFilePtr, sRptTitle
    Print #nFilePtr, "Run Date: ", Format(Date, "mm/dd/yyyy")
    Print #nFilePtr, "Run Time: ", Format(Time, "hh:mm:ss AMPM")
    Print #nFilePtr, String(CHAR_PER_LINE, "-")
    Print #nFilePtr, sLineHeader
    
    Do While I <= INDEX_REPORT
        'print title and time
        Print #nFilePtr, vData(I)
        I = I + 1
    Loop
    
    Close #nFilePtr
    
    tfnSetStatusBarCorrect "Finished Writing Log File: " & sFileName
    fnWriteLog = True
    Exit Function

ErrorWriteFile:

    MsgBox "An error has occurred while writing Log File: " & sFileName & _
        ". Err Code: " & Err.Number & ", Err Desc: " & Err.Description, vbCritical
    
    tfnSetStatusBarError "An error has occurred while writing the Log File: " & sFileName
    If nFilePtr <> 0 Then
        Close #nFilePtr
    End If
End Function
Private Function fnGetPayCode() As Boolean
    Dim strSQL As String
    Dim rsTemp As Recordset
    Const SUB_NAME As String = "fnGetPayCode"
    fnGetPayCode = False
    
    strSQL = "SELECT parm_field FROM sys_parm WHERE parm_nbr=30854"
    If GetRecordSet(rsTemp, strSQL, , SUB_NAME) < 0 Then
        MsgBox "Failed to access database to get records", vbCritical
        Exit Function
    End If
    If rsTemp.RecordCount = 0 Then
        MsgBox "No hourly/overtime pay code can be found.", vbExclamation
        Exit Function
    End If
    If IsNull(rsTemp!parm_field) Then
        Exit Function
    End If
    strHourCode = Mid(Trim(rsTemp!parm_field), 1, 4)
    strOverCode = Mid(Trim(rsTemp!parm_field), 6, 4)
    
    fnGetPayCode = True
End Function
Private Sub subInitGrid()
    Dim sPtnAmount As String
    Dim sFmtAmount As String
    Dim k As Integer
    Dim nWidth As Integer
    
    
    sPtnAmount = tfnDecimalPattern(16, 2, True)
    sFmtAmount = "###,###,##0.00"
    
    Set tgmGrid = New clsTGSpreadSheet
    Set tgmGrid.Table = tblGrid
    Set tgmGrid.StatusBar = ffraStatusbar
    Set tgmGrid.Form = Me
    

    nWidth = tblGrid.Width
    With tblGrid.Columns(COL_EMPNO)
        .Width = 0.08 * nWidth
        .DataField = "empno"
    End With
    With tblGrid.Columns(COL_NAME)
        .Width = 0.12 * nWidth
        .DataField = "lname"
    End With
    With tblGrid.Columns(COL_LEVEL)
        .Width = 0.06 * nWidth
        .DataField = "emp_level"
    End With
    With tblGrid.Columns(COL_DESC)
        .Width = 0.14 * nWidth
        .DataField = "emp_level_desc"
    End With
    With tblGrid.Columns(COL_HIRE_DATE)
        .Width = 0.13 * nWidth
        .DataField = "date_hired"
    End With
    With tblGrid.Columns(COL_DAYS_WORKED)
        .Width = 0.16 * nWidth
        .DataField = "days_worked"
    End With
    With tblGrid.Columns(COL_OLD_RATE)
        .Width = 0.11 * nWidth
        .DataField = "old_rate"
    End With
    With tblGrid.Columns(COL_NEW_RATE)
        .Width = 0.11 * nWidth
        .DataField = "new_rate"
    End With
    With tblGrid.Columns(COL_DIFF)
        .Width = 0.09 * nWidth
        .DataField = "rate_diff"
    End With
    tgmGrid.DisplayFormat(COL_OLD_RATE) = sFmtAmount
    tgmGrid.DisplayFormat(COL_NEW_RATE) = sFmtAmount
    tgmGrid.DisplayFormat(COL_DIFF) = sFmtAmount
    
    tgmGrid.SetupTable True
    tgmGrid.AllowAddNew = False
    tgmGrid.ClearData

End Sub
Public Sub subPrintRpt()
    

    Dim sProgName As String
    Dim sRptTitle As String
    Dim sLineHeader As String
    Dim sExtraLine As String
    
    Dim lCount As Long
    Dim I As Long
    Dim sLine As String
    
    sProgName = "Program ID: ZZSEPIMT"
    sRptTitle = "7-Eleven Automatic Pay Increase Processing Log"
    
    sLineHeader = fnTranc("Emp#", 10, LEFT_JUSTIFY) & _
                  fnTranc("Name", 15, LEFT_JUSTIFY) & _
                  fnTranc("Level", 6, LEFT_JUSTIFY) & _
                  fnTranc("Description", 21, LEFT_JUSTIFY) & _
                  fnTranc("Hire Date", 15, LEFT_JUSTIFY) & _
                  fnTranc("Days Worked", 12, LEFT_JUSTIFY) & _
                  fnTranc("Old Rate", 15, LEFT_JUSTIFY) & _
                  fnTranc("New Rate", 15, LEFT_JUSTIFY) & _
                  fnTranc("Diff", 15, LEFT_JUSTIFY)
        

    INDEX_REPORT = -1
    
    lCount = frmZZSEPIMT.tgmGrid.RowCount
    
    If lCount <= 0 Then
        Exit Sub
    End If
    
    
    For I = 0 To lCount - 1
        With tgmGrid
           sLine = fnTranc(.CellValue(COL_EMPNO, I), 10, LEFT_JUSTIFY) & _
              fnTranc(.CellValue(COL_NAME, I), 15, LEFT_JUSTIFY) & _
              fnTranc(.CellValue(COL_LEVEL, I), 6, LEFT_JUSTIFY) & _
              fnTranc(.CellValue(COL_DESC, I), 21, LEFT_JUSTIFY) & _
              fnTranc(.CellValue(COL_HIRE_DATE, I), 15, LEFT_JUSTIFY) & _
              fnTranc(.CellValue(COL_DAYS_WORKED, I), 12, LEFT_JUSTIFY) & _
              fnTranc(.CellValue(COL_OLD_RATE, I), 15, LEFT_JUSTIFY) & _
              fnTranc(.CellValue(COL_NEW_RATE, I), 15, LEFT_JUSTIFY) & _
              fnTranc(.CellValue(COL_DIFF, I), 15, LEFT_JUSTIFY)
          'put in the array report
          subLineFeed sLine
       End With
    Next
    

    'write log file
    fnWriteLog ary_REPORT, sProgName, sRptTitle, sLineHeader
    
    
    tfnSetStatusBarMessage "Sending Data to Printer -- " & Printer.DeviceName & " ..."
    fnSendToPrinter sProgName, sRptTitle, sExtraLine, sLineHeader
    
    tfnSetStatusBarMessage "Report has been sent to Printer -- " & Printer.DeviceName
    
End Sub
Private Function fnValidNumer() As Boolean
    Const SUB_NAME = "fnValidNumer"
    Dim strSQL As String
    Dim strType As String
    Dim strLoc As String
    Dim rsTemp As Recordset
    
    fnValidNumer = False

    strLoc = Trim(txtLevel.Text)
    If strLoc = "" Then
        cValidate.SetErrorMessage txtLevel, "Employee Level cannot be blank!"
        Exit Function

    End If
        
    fnValidNumer = True
End Function
