VERSION 5.00
Object = "{C75015E0-2232-11D3-B440-0060971E99AF}#1.0#0"; "FACTFRM.OCX"
Object = "{3D388220-1F4E-11D3-B440-0060971E99AF}#1.0#0"; "FACTTAB.OCX"
Object = "{01028C21-0000-0000-0000-000000000046}#4.0#0"; "TG32OV.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmZZSEBPRC 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Process Commission Checks"
   ClientHeight    =   6060
   ClientLeft      =   1056
   ClientTop       =   1956
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
   HelpContextID   =   1
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6060
   ScaleWidth      =   8880
   Begin FACTFRMLib.FactorFrame cmdExitCancelBtn 
      Height          =   396
      HelpContextID   =   15
      Left            =   7500
      TabIndex        =   61
      Top             =   5268
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
   Begin FACTFRMLib.FactorFrame efraBackground 
      Height          =   5184
      Left            =   0
      TabIndex        =   63
      TabStop         =   0   'False
      Top             =   492
      Width           =   8880
      _Version        =   65536
      _ExtentX        =   15663
      _ExtentY        =   9144
      _StockProps     =   77
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.4
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
      Begin FACTTABLib.FactorTab eTabMain 
         Height          =   5016
         Left            =   48
         TabIndex        =   65
         TabStop         =   0   'False
         Top             =   12
         Width           =   8796
         _Version        =   65536
         _ExtentX        =   15515
         _ExtentY        =   8848
         _StockProps     =   68
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   10.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Pay En&try|Process Chec&ks|&View/Approve Checks|View &Details"
         Begin FACTFRMLib.FactorFrame efraBaseDetail 
            Height          =   4716
            Left            =   14580
            TabIndex        =   70
            TabStop         =   0   'False
            Top             =   0
            Width           =   8796
            _Version        =   65536
            _ExtentX        =   15515
            _ExtentY        =   8318
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
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Begin FACTFRMLib.FactorFrame efraBaseIIDetail 
               Height          =   4152
               Left            =   60
               TabIndex        =   71
               TabStop         =   0   'False
               Top             =   60
               Width           =   8664
               _Version        =   65536
               _ExtentX        =   15282
               _ExtentY        =   7324
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
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Begin VB.TextBox txtEmployee 
                  BackColor       =   &H00FFFFFF&
                  ForeColor       =   &H00000000&
                  Height          =   360
                  HelpContextID   =   575
                  Left            =   60
                  TabIndex        =   54
                  Tag             =   "pn_alt"
                  Top             =   276
                  Width           =   1872
               End
               Begin VB.TextBox txtEmpName 
                  BackColor       =   &H00FFFFFF&
                  DataSource      =   "datVendor"
                  ForeColor       =   &H00000000&
                  Height          =   360
                  HelpContextID   =   577
                  Left            =   2364
                  TabIndex        =   56
                  Tag             =   "pn_name"
                  Top             =   276
                  Width           =   5868
               End
               Begin DBTrueGrid.TDBGrid tblDetails 
                  Height          =   3408
                  HelpContextID   =   579
                  Left            =   60
                  OleObjectBlob   =   "ZZSEBPRC.frx":0000
                  TabIndex        =   58
                  Top             =   696
                  Width           =   8556
               End
               Begin FACTFRMLib.FactorFrame cmdEmployee 
                  Height          =   360
                  HelpContextID   =   576
                  Left            =   1944
                  TabIndex        =   55
                  TabStop         =   0   'False
                  Tag             =   "Run #3"
                  Top             =   276
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
                  Picture         =   "ZZSEBPRC.frx":12DE
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
               Begin FACTFRMLib.FactorFrame cmdEmpName 
                  Height          =   360
                  HelpContextID   =   578
                  Left            =   8244
                  TabIndex        =   57
                  TabStop         =   0   'False
                  Tag             =   "Run #3"
                  Top             =   276
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
                  Picture         =   "ZZSEBPRC.frx":13F0
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
               Begin VB.Label lblEmployee 
                  Caption         =   "Employee Number"
                  Height          =   252
                  Left            =   60
                  TabIndex        =   73
                  Top             =   36
                  Width           =   1836
               End
               Begin VB.Label lblEmpName 
                  Caption         =   "Employee Name"
                  Height          =   252
                  Left            =   2364
                  TabIndex        =   72
                  Top             =   36
                  Width           =   1968
               End
            End
            Begin FACTFRMLib.FactorFrame cmdPrint 
               Height          =   396
               HelpContextID   =   32
               Index           =   3
               Left            =   5988
               TabIndex        =   59
               Top             =   4260
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
               Caption         =   "&Print"
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
            Begin FACTFRMLib.FactorFrame cmdCancel 
               Height          =   396
               HelpContextID   =   15
               Index           =   3
               Left            =   7428
               TabIndex        =   60
               Top             =   4260
               Width           =   1308
               _Version        =   65536
               _ExtentX        =   2302
               _ExtentY        =   688
               _StockProps     =   77
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   10.2
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
                  Size            =   9.6
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
         End
         Begin FACTFRMLib.FactorFrame efraBaseProcess 
            Height          =   4716
            Left            =   14460
            TabIndex        =   67
            TabStop         =   0   'False
            Top             =   0
            Width           =   8796
            _Version        =   65536
            _ExtentX        =   15515
            _ExtentY        =   8318
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
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Begin FACTFRMLib.FactorFrame efraBaseIIProcess 
               Height          =   4152
               Left            =   60
               TabIndex        =   75
               TabStop         =   0   'False
               Top             =   60
               Width           =   8664
               _Version        =   65536
               _ExtentX        =   15282
               _ExtentY        =   7324
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
               Caption         =   " "
               BeginProperty PanelFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   9.6
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Begin FACTFRMLib.FactorFrame efraProcessDate 
                  Height          =   1332
                  Left            =   72
                  TabIndex        =   98
                  TabStop         =   0   'False
                  Top             =   48
                  Width           =   2400
                  _Version        =   65536
                  _ExtentX        =   4233
                  _ExtentY        =   2350
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
                     Size            =   12
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Begin VB.TextBox txtFrequency 
                     BackColor       =   &H00FFFFFF&
                     ForeColor       =   &H00000000&
                     Height          =   360
                     HelpContextID   =   530
                     Left            =   1404
                     TabIndex        =   34
                     Tag             =   "pn_alt"
                     Top             =   252
                     Width           =   552
                  End
                  Begin VB.TextBox txtStartDate 
                     BackColor       =   &H00FFFFFF&
                     ForeColor       =   &H00000000&
                     Height          =   360
                     HelpContextID   =   529
                     Left            =   48
                     TabIndex        =   32
                     Tag             =   "pn_alt"
                     Top             =   252
                     Width           =   1224
                  End
                  Begin VB.TextBox txtEndDate 
                     BackColor       =   &H00FFFFFF&
                     ForeColor       =   &H00000000&
                     Height          =   360
                     HelpContextID   =   529
                     Left            =   48
                     TabIndex        =   33
                     Tag             =   "pn_alt"
                     Top             =   900
                     Width           =   1224
                  End
                  Begin FACTFRMLib.FactorFrame cmdFrequency 
                     Height          =   360
                     HelpContextID   =   531
                     Left            =   1968
                     TabIndex        =   35
                     TabStop         =   0   'False
                     Tag             =   "Run #3"
                     Top             =   252
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
                     Picture         =   "ZZSEBPRC.frx":1502
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
                  Begin VB.Label lblFrequency 
                     Caption         =   "Frequency"
                     Height          =   252
                     Left            =   1404
                     TabIndex        =   101
                     Top             =   12
                     Width           =   960
                  End
                  Begin VB.Label lblDate 
                     Caption         =   "Starting  Date"
                     Height          =   252
                     Left            =   48
                     TabIndex        =   100
                     Top             =   12
                     Width           =   1488
                  End
                  Begin VB.Label Label1 
                     Caption         =   "Ending Date"
                     Height          =   252
                     Left            =   48
                     TabIndex        =   99
                     Top             =   660
                     Width           =   1236
                  End
               End
               Begin VB.TextBox txtEmpProcess 
                  BackColor       =   &H00FFFFFF&
                  ForeColor       =   &H00000000&
                  Height          =   360
                  HelpContextID   =   532
                  Left            =   2508
                  TabIndex        =   40
                  Tag             =   "pn_alt"
                  Top             =   960
                  Width           =   1104
               End
               Begin VB.TextBox txtEmpNameProcess 
                  BackColor       =   &H00FFFFFF&
                  DataSource      =   "datVendor"
                  ForeColor       =   &H00000000&
                  Height          =   360
                  HelpContextID   =   534
                  Left            =   4044
                  TabIndex        =   42
                  Tag             =   "pn_name"
                  Top             =   960
                  Width           =   4176
               End
               Begin VB.ListBox lstProcess 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  ForeColor       =   &H00000000&
                  Height          =   2352
                  HelpContextID   =   536
                  IntegralHeight  =   0   'False
                  ItemData        =   "ZZSEBPRC.frx":1614
                  Left            =   72
                  List            =   "ZZSEBPRC.frx":1616
                  TabIndex        =   44
                  Top             =   1404
                  Width           =   8532
               End
               Begin VB.TextBox txtPrftCtrName 
                  BackColor       =   &H00FFFFFF&
                  DataSource      =   "datVendor"
                  ForeColor       =   &H00000000&
                  Height          =   360
                  HelpContextID   =   527
                  Left            =   4044
                  TabIndex        =   38
                  Tag             =   "pn_name"
                  Top             =   300
                  Width           =   4176
               End
               Begin VB.TextBox txtPrftCtr 
                  BackColor       =   &H00FFFFFF&
                  ForeColor       =   &H00000000&
                  Height          =   360
                  HelpContextID   =   525
                  Left            =   2508
                  TabIndex        =   36
                  Tag             =   "pn_alt"
                  Top             =   300
                  Width           =   1104
               End
               Begin MSComctlLib.ProgressBar pbBarMain 
                  Height          =   312
                  Left            =   72
                  TabIndex        =   76
                  Top             =   3780
                  Width           =   8532
                  _ExtentX        =   15050
                  _ExtentY        =   550
                  _Version        =   393216
                  Appearance      =   1
               End
               Begin FACTFRMLib.FactorFrame cmdPrftCtr 
                  Height          =   360
                  HelpContextID   =   526
                  Left            =   3624
                  TabIndex        =   37
                  TabStop         =   0   'False
                  Tag             =   "Run #3"
                  Top             =   300
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
                  Picture         =   "ZZSEBPRC.frx":1618
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
               Begin FACTFRMLib.FactorFrame cmdPrftCtrName 
                  Height          =   360
                  HelpContextID   =   528
                  Left            =   8232
                  TabIndex        =   39
                  TabStop         =   0   'False
                  Tag             =   "Run #3"
                  Top             =   300
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
                  Picture         =   "ZZSEBPRC.frx":172A
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
               Begin FACTFRMLib.FactorFrame cmdEmpProcess 
                  Height          =   360
                  HelpContextID   =   533
                  Left            =   3624
                  TabIndex        =   41
                  TabStop         =   0   'False
                  Tag             =   "Run #3"
                  Top             =   960
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
                  Picture         =   "ZZSEBPRC.frx":183C
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
               Begin FACTFRMLib.FactorFrame cmdEmpNameProcess 
                  Height          =   360
                  HelpContextID   =   535
                  Left            =   8232
                  TabIndex        =   43
                  TabStop         =   0   'False
                  Tag             =   "Run #3"
                  Top             =   960
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
                  Picture         =   "ZZSEBPRC.frx":194E
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
               Begin VB.Label lblEmpProcess 
                  Caption         =   "Employee Nbr"
                  Height          =   252
                  Left            =   2508
                  TabIndex        =   95
                  Top             =   720
                  Width           =   1296
               End
               Begin VB.Label lblEmpNameProcess 
                  Caption         =   "Employee Name"
                  Height          =   252
                  Left            =   4044
                  TabIndex        =   94
                  Top             =   720
                  Width           =   1956
               End
               Begin VB.Label lblPrftCtrName 
                  Caption         =   "Profit Center Name"
                  Height          =   252
                  Left            =   4044
                  TabIndex        =   78
                  Top             =   60
                  Width           =   1956
               End
               Begin VB.Label lblPrftCtr 
                  Caption         =   "Profit Center"
                  Height          =   252
                  Left            =   2508
                  TabIndex        =   77
                  Top             =   60
                  Width           =   1296
               End
            End
            Begin FACTFRMLib.FactorFrame cmdProcess 
               Height          =   396
               HelpContextID   =   537
               Left            =   5988
               TabIndex        =   45
               Top             =   4260
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
               Caption         =   "P&rocess"
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
            Begin FACTFRMLib.FactorFrame cmdCancel 
               Height          =   396
               HelpContextID   =   15
               Index           =   1
               Left            =   7428
               TabIndex        =   47
               Top             =   4260
               Width           =   1308
               _Version        =   65536
               _ExtentX        =   2302
               _ExtentY        =   688
               _StockProps     =   77
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   10.2
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
               HelpContextID   =   32
               Index           =   1
               Left            =   48
               TabIndex        =   46
               Top             =   4260
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
               Caption         =   "&Print"
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
         Begin FACTFRMLib.FactorFrame efraBasePayEntry 
            Height          =   4692
            Left            =   12
            TabIndex        =   66
            TabStop         =   0   'False
            Top             =   12
            Width           =   8772
            _Version        =   65536
            _ExtentX        =   15473
            _ExtentY        =   8276
            _StockProps     =   77
            BackColor       =   -2147483633
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
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Begin FACTTABLib.FactorTab eTabSub 
               Height          =   4716
               Left            =   0
               TabIndex        =   79
               TabStop         =   0   'False
               Top             =   0
               Width           =   8784
               _Version        =   65536
               _ExtentX        =   15494
               _ExtentY        =   8318
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
               Position        =   3
               TabsPerPage     =   2
               Caption         =   "Store &Sales|Employee &Hours"
               Begin FACTFRMLib.FactorFrame efraBaseHours 
                  Height          =   4428
                  Left            =   14460
                  TabIndex        =   81
                  TabStop         =   0   'False
                  Top             =   0
                  Width           =   8784
                  _Version        =   65536
                  _ExtentX        =   15494
                  _ExtentY        =   7810
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
                     Size            =   12
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Begin FACTFRMLib.FactorFrame efraBaseIIHours 
                     Height          =   4152
                     Left            =   48
                     TabIndex        =   27
                     Top             =   48
                     Width           =   8376
                     _Version        =   65536
                     _ExtentX        =   14774
                     _ExtentY        =   7324
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
                        Size            =   12
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Begin VB.TextBox txtTotal 
                        Alignment       =   1  'Right Justify
                        Enabled         =   0   'False
                        Height          =   324
                        Left            =   7032
                        MultiLine       =   -1  'True
                        TabIndex        =   97
                        Top             =   3780
                        Width           =   1272
                     End
                     Begin VB.TextBox txtTotalDollars 
                        Alignment       =   1  'Right Justify
                        Enabled         =   0   'False
                        Height          =   324
                        Left            =   4632
                        MultiLine       =   -1  'True
                        TabIndex        =   96
                        Top             =   3780
                        Width           =   1332
                     End
                     Begin VB.PictureBox cmdFloatingBtn 
                        Appearance      =   0  'Flat
                        BackColor       =   &H00C0C0C0&
                        ForeColor       =   &H80000008&
                        Height          =   240
                        HelpContextID   =   22
                        Left            =   96
                        ScaleHeight     =   216
                        ScaleWidth      =   228
                        TabIndex        =   91
                        Top             =   720
                        Visible         =   0   'False
                        Width           =   255
                     End
                     Begin VB.TextBox txtEmployeeName 
                        BackColor       =   &H00FFFFFF&
                        DataSource      =   "datVendor"
                        ForeColor       =   &H00000000&
                        Height          =   360
                        HelpContextID   =   515
                        Left            =   2028
                        TabIndex        =   21
                        Tag             =   "pn_name"
                        Top             =   276
                        Width           =   3576
                     End
                     Begin VB.TextBox txtEmployeeNumber 
                        BackColor       =   &H00FFFFFF&
                        ForeColor       =   &H00000000&
                        Height          =   360
                        HelpContextID   =   513
                        Left            =   72
                        TabIndex        =   17
                        Tag             =   "pn_alt"
                        Top             =   276
                        Width           =   1524
                     End
                     Begin VB.TextBox txtSSN 
                        Height          =   360
                        HelpContextID   =   517
                        Left            =   6036
                        TabIndex        =   23
                        Top             =   276
                        Width           =   1896
                     End
                     Begin FACTFRMLib.FactorFrame cmdEmployeeNumber 
                        Height          =   360
                        HelpContextID   =   514
                        Left            =   1608
                        TabIndex        =   19
                        TabStop         =   0   'False
                        Tag             =   "Run #3"
                        Top             =   276
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
                        Picture         =   "ZZSEBPRC.frx":1A60
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
                     Begin FACTFRMLib.FactorFrame cmdEmployeeName 
                        Height          =   360
                        HelpContextID   =   516
                        Left            =   5616
                        TabIndex        =   22
                        TabStop         =   0   'False
                        Tag             =   "Run #3"
                        Top             =   276
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
                        Picture         =   "ZZSEBPRC.frx":1B72
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
                     Begin FACTFRMLib.FactorFrame cmdSSN 
                        Height          =   360
                        HelpContextID   =   518
                        Left            =   7944
                        TabIndex        =   24
                        TabStop         =   0   'False
                        Tag             =   "Run #3"
                        Top             =   276
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
                        Picture         =   "ZZSEBPRC.frx":1C84
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
                     Begin DBTrueGrid.TDBGrid tblTimeCard 
                        Height          =   3036
                        HelpContextID   =   519
                        Left            =   72
                        OleObjectBlob   =   "ZZSEBPRC.frx":1D96
                        TabIndex        =   25
                        Top             =   696
                        Width           =   5892
                     End
                     Begin DBTrueGrid.TDBGrid tblProfitCenter 
                        Height          =   3036
                        HelpContextID   =   520
                        Left            =   6036
                        OleObjectBlob   =   "ZZSEBPRC.frx":29D5
                        TabIndex        =   26
                        Top             =   696
                        Width           =   2268
                     End
                     Begin VB.Label lblTotalHr 
                        Alignment       =   1  'Right Justify
                        Caption         =   "Total Hours/Dollars:"
                        Height          =   252
                        Left            =   2892
                        TabIndex        =   88
                        Top             =   3828
                        Width           =   1692
                     End
                     Begin VB.Label lblSSN 
                        Caption         =   "Social Security Number"
                        Height          =   252
                        Left            =   6036
                        TabIndex        =   87
                        Top             =   24
                        Width           =   2268
                     End
                     Begin VB.Label lblHEmpName 
                        Caption         =   "Employee Name"
                        Height          =   252
                        Left            =   2028
                        TabIndex        =   86
                        Top             =   24
                        Width           =   1968
                     End
                     Begin VB.Label lblEmpNo 
                        Caption         =   "Employee Number"
                        Height          =   252
                        Left            =   72
                        TabIndex        =   85
                        Top             =   24
                        Width           =   1848
                     End
                     Begin VB.Label lblTotalPC 
                        Alignment       =   1  'Right Justify
                        Caption         =   "PC Total:"
                        Height          =   252
                        Left            =   6048
                        TabIndex        =   84
                        Top             =   3828
                        Width           =   888
                     End
                  End
                  Begin FACTFRMLib.FactorFrame cmdAddBtn 
                     Height          =   396
                     HelpContextID   =   10
                     Index           =   4
                     Left            =   36
                     TabIndex        =   13
                     Top             =   4248
                     Width           =   1308
                     _Version        =   65536
                     _ExtentX        =   2307
                     _ExtentY        =   698
                     _StockProps     =   77
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "Arial"
                        Size            =   10.2
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
                     Index           =   4
                     Left            =   1452
                     TabIndex        =   15
                     Top             =   4248
                     Width           =   1308
                     _Version        =   65536
                     _ExtentX        =   2307
                     _ExtentY        =   698
                     _StockProps     =   77
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "Arial"
                        Size            =   10.2
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
                  Begin FACTFRMLib.FactorFrame cmdCancel 
                     Height          =   396
                     HelpContextID   =   15
                     Index           =   4
                     Left            =   7128
                     TabIndex        =   29
                     Top             =   4248
                     Width           =   1308
                     _Version        =   65536
                     _ExtentX        =   2307
                     _ExtentY        =   698
                     _StockProps     =   77
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "Arial"
                        Size            =   10.2
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
                     Index           =   4
                     Left            =   5712
                     TabIndex        =   28
                     Top             =   4248
                     Width           =   1308
                     _Version        =   65536
                     _ExtentX        =   2307
                     _ExtentY        =   698
                     _StockProps     =   77
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "Arial"
                        Size            =   10.2
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
                  Begin FACTFRMLib.FactorFrame cmdDelete 
                     Height          =   396
                     HelpContextID   =   12
                     Index           =   4
                     Left            =   2868
                     TabIndex        =   31
                     Top             =   4248
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
                  Begin FACTFRMLib.FactorFrame cmdRefresh 
                     Height          =   396
                     HelpContextID   =   14
                     Index           =   4
                     Left            =   4296
                     TabIndex        =   30
                     Top             =   4248
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
               End
               Begin FACTFRMLib.FactorFrame efraBaseSales 
                  Height          =   4692
                  Left            =   12
                  TabIndex        =   80
                  TabStop         =   0   'False
                  Top             =   12
                  Width           =   8472
                  _Version        =   65536
                  _ExtentX        =   14944
                  _ExtentY        =   8276
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
                     Size            =   12
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Begin FACTFRMLib.FactorFrame efraBaseIISales 
                     Height          =   4152
                     Left            =   48
                     TabIndex        =   12
                     Top             =   48
                     Width           =   8376
                     _Version        =   65536
                     _ExtentX        =   14774
                     _ExtentY        =   7324
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
                        Size            =   12
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
                        TabIndex        =   90
                        Top             =   768
                        Visible         =   0   'False
                        Width           =   255
                     End
                     Begin FACTFRMLib.FactorFrame efraOptSales 
                        Height          =   648
                        Left            =   72
                        TabIndex        =   89
                        TabStop         =   0   'False
                        Top             =   72
                        Width           =   4392
                        _Version        =   65536
                        _ExtentX        =   7747
                        _ExtentY        =   1143
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
                           Size            =   12
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Begin VB.OptionButton optType 
                           Caption         =   "Week Sales"
                           Height          =   272
                           HelpContextID   =   500
                           Index           =   0
                           Left            =   84
                           TabIndex        =   3
                           Top             =   36
                           Value           =   -1  'True
                           Width           =   2016
                        End
                        Begin VB.OptionButton optType 
                           Caption         =   "One Month Sales"
                           Height          =   272
                           HelpContextID   =   501
                           Index           =   1
                           Left            =   84
                           TabIndex        =   5
                           Top             =   324
                           Width           =   2016
                        End
                        Begin VB.OptionButton optType 
                           Caption         =   "Gas Sales"
                           Height          =   272
                           HelpContextID   =   503
                           Index           =   2
                           Left            =   2244
                           TabIndex        =   6
                           Top             =   324
                           Width           =   2076
                        End
                        Begin VB.OptionButton optType 
                           Caption         =   "Three Month Sales"
                           Height          =   272
                           HelpContextID   =   502
                           Index           =   3
                           Left            =   2244
                           TabIndex        =   4
                           Top             =   36
                           Width           =   2076
                        End
                     End
                     Begin VB.TextBox txtFromDate 
                        BackColor       =   &H00FFFFFF&
                        ForeColor       =   &H00000000&
                        Height          =   360
                        HelpContextID   =   504
                        Left            =   4524
                        TabIndex        =   7
                        Tag             =   "pn_alt"
                        Top             =   348
                        Width           =   1488
                     End
                     Begin VB.TextBox txtToDate 
                        BackColor       =   &H00FFFFFF&
                        DataSource      =   "datVendor"
                        ForeColor       =   &H00000000&
                        Height          =   360
                        HelpContextID   =   506
                        Left            =   6444
                        TabIndex        =   9
                        Tag             =   "pn_name"
                        Top             =   348
                        Width           =   1488
                     End
                     Begin DBTrueGrid.TDBGrid tblSales 
                        Height          =   3324
                        HelpContextID   =   508
                        Left            =   72
                        OleObjectBlob   =   "ZZSEBPRC.frx":3618
                        TabIndex        =   11
                        Top             =   768
                        Width           =   8244
                     End
                     Begin FACTFRMLib.FactorFrame cmdFromDate 
                        Height          =   360
                        HelpContextID   =   505
                        Left            =   6024
                        TabIndex        =   8
                        TabStop         =   0   'False
                        Tag             =   "Run #3"
                        Top             =   348
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
                        Picture         =   "ZZSEBPRC.frx":48F4
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
                     Begin FACTFRMLib.FactorFrame cmdToDate 
                        Height          =   360
                        HelpContextID   =   507
                        Left            =   7944
                        TabIndex        =   10
                        TabStop         =   0   'False
                        Tag             =   "Run #3"
                        Top             =   348
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
                        Picture         =   "ZZSEBPRC.frx":4A06
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
                     Begin VB.Label lblFromDate 
                        Caption         =   "From Date"
                        Height          =   252
                        Left            =   4536
                        TabIndex        =   83
                        Top             =   96
                        Width           =   1380
                     End
                     Begin VB.Label lblToDate 
                        Caption         =   "To Date"
                        Height          =   252
                        Left            =   6444
                        TabIndex        =   82
                        Top             =   108
                        Width           =   1500
                     End
                  End
                  Begin FACTFRMLib.FactorFrame cmdUpdateInsertBtn 
                     Height          =   396
                     HelpContextID   =   13
                     Index           =   0
                     Left            =   5712
                     TabIndex        =   14
                     Top             =   4248
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
                  Begin FACTFRMLib.FactorFrame cmdCancel 
                     Height          =   396
                     HelpContextID   =   15
                     Index           =   0
                     Left            =   7128
                     TabIndex        =   16
                     Top             =   4248
                     Width           =   1308
                     _Version        =   65536
                     _ExtentX        =   2302
                     _ExtentY        =   688
                     _StockProps     =   77
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "Arial"
                        Size            =   10.2
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
                     Left            =   36
                     TabIndex        =   1
                     Top             =   4248
                     Width           =   1308
                     _Version        =   65536
                     _ExtentX        =   2302
                     _ExtentY        =   688
                     _StockProps     =   77
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "Arial"
                        Size            =   10.2
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
                     Left            =   1452
                     TabIndex        =   2
                     Top             =   4248
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
                  Begin FACTFRMLib.FactorFrame cmdDelete 
                     Height          =   396
                     HelpContextID   =   12
                     Index           =   0
                     Left            =   2868
                     TabIndex        =   20
                     Top             =   4248
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
                  Begin FACTFRMLib.FactorFrame cmdRefresh 
                     Height          =   396
                     HelpContextID   =   14
                     Index           =   0
                     Left            =   4296
                     TabIndex        =   18
                     Top             =   4248
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
               End
            End
         End
         Begin FACTFRMLib.FactorFrame efraBaseView 
            Height          =   4716
            Left            =   14520
            TabIndex        =   69
            TabStop         =   0   'False
            Top             =   0
            Width           =   8796
            _Version        =   65536
            _ExtentX        =   15515
            _ExtentY        =   8318
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
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Begin FACTFRMLib.FactorFrame efraBaseIIView 
               Height          =   4152
               Left            =   60
               TabIndex        =   74
               TabStop         =   0   'False
               Top             =   60
               Width           =   8664
               _Version        =   65536
               _ExtentX        =   15282
               _ExtentY        =   7324
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
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Begin DBTrueGrid.TDBGrid tblApprove 
                  Height          =   4044
                  HelpContextID   =   550
                  Left            =   60
                  OleObjectBlob   =   "ZZSEBPRC.frx":4B18
                  TabIndex        =   48
                  Top             =   60
                  Width           =   8556
               End
            End
            Begin FACTFRMLib.FactorFrame cmdOk 
               Height          =   396
               HelpContextID   =   16
               Left            =   5988
               TabIndex        =   51
               Top             =   4260
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
               Caption         =   "O&K"
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
            Begin FACTFRMLib.FactorFrame cmdCancel 
               Height          =   396
               HelpContextID   =   15
               Index           =   2
               Left            =   7428
               TabIndex        =   53
               Top             =   4260
               Width           =   1308
               _Version        =   65536
               _ExtentX        =   2302
               _ExtentY        =   688
               _StockProps     =   77
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   10.2
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
                  Size            =   9.6
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin FACTFRMLib.FactorFrame cmdSelectAll 
               Height          =   396
               HelpContextID   =   551
               Left            =   48
               TabIndex        =   49
               Top             =   4260
               Width           =   1308
               _Version        =   65536
               _ExtentX        =   2302
               _ExtentY        =   688
               _StockProps     =   77
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   10.2
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "&Select All"
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
            Begin FACTFRMLib.FactorFrame cmdApprove 
               Height          =   396
               HelpContextID   =   552
               Left            =   1476
               TabIndex        =   50
               Top             =   4260
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
               Caption         =   "&Approve"
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
               HelpContextID   =   32
               Index           =   2
               Left            =   2904
               TabIndex        =   52
               Top             =   4260
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
               Caption         =   "&Print"
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
      Begin DBTrueGrid.TDBGrid tblComboDropdown 
         Bindings        =   "ZZSEBPRC.frx":5DF6
         Height          =   2484
         Left            =   108
         OleObjectBlob   =   "ZZSEBPRC.frx":5E15
         TabIndex        =   68
         TabStop         =   0   'False
         Top             =   12
         Width           =   3756
      End
      Begin VB.Data datDropDown 
         Caption         =   "DropDown"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   324
         Left            =   336
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   1608
         Visible         =   0   'False
         Width           =   1908
      End
   End
   Begin FACTFRMLib.FactorFrame ffraStatusbar 
      Height          =   360
      Left            =   0
      TabIndex        =   62
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
   Begin FACTFRMLib.FactorFrame efraToolBar 
      Height          =   468
      Left            =   0
      TabIndex        =   64
      TabStop         =   0   'False
      Top             =   0
      Width           =   8880
      _Version        =   65536
      _ExtentX        =   15663
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
      FMName          =   "ZZSEBPRC"
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
         Width           =   6876
         _ExtentX        =   12129
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
   Begin DBTrueGrid.TDBGrid tblDropDown 
      Bindings        =   "ZZSEBPRC.frx":70F9
      Height          =   1008
      Index           =   0
      Left            =   168
      OleObjectBlob   =   "ZZSEBPRC.frx":7113
      TabIndex        =   92
      Top             =   684
      Width           =   2604
   End
   Begin DBTrueGrid.TDBGrid tblDropDown 
      Bindings        =   "ZZSEBPRC.frx":83F5
      Height          =   1008
      Index           =   4
      Left            =   144
      OleObjectBlob   =   "ZZSEBPRC.frx":840F
      TabIndex        =   93
      Top             =   624
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
Attribute VB_Name = "frmZZSEBPRC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***********************************************************************
'Copyright (c) 1999 FACTOR, A Division of W.R.Hess Company             *
'Program ID     : ZZSEBPRC                                             *
'Date Compiled  : 28 Dec 00                                            *
'Programmer     : Rajneesh Aggarwal                                    *
'***********************************************************************
Option Explicit

Private t_bStartupFlag As Boolean 'optional startup flag
Private t_bDataChanged As Boolean 'data changed flag
Private t_bUpdateTable As Boolean 'update data flag
Private t_bTabSwitch As Boolean

'========================
'Standard Button Captions
'========================
Private Const t_szCAPTION_INSERT As String = "&Insert"
Private Const t_szCAPTION_UPDATE As String = "&Update"
Private Const t_szCAPTION_REFRESH As String = "&Refresh"
Private Const t_szCAPTION_CANCEL As String = "&Cancel"
Private Const t_szCAPTION_EXIT As String = "E&xit"

'==========================
'Status Bar Default Strings
'==========================
Private Const t_szEXIT As String = "Exit"
Private Const t_szCANCEL As String = "Cancel"

Private Const t_szPRINT As String = "Print"
Private Const t_szHELP As String = "Help"

Private Const nWeek As Integer = 0
Private Const sWeek As String = "W"
Private Const nOneMth As Integer = 1
Private Const sOneMth As String = "M"
Private Const nGas As Integer = 2
Private Const sGas As String = "G"
Private Const nThreeMth As Integer = 3
Private Const sThreeMth As String = "Q"

Private tgfDropdown(4) As clsFloatingDropDown

Private bCancelProcess As Boolean
Private cValidate As cValidateInput
Private cValidSls As cValidateInput

Private bSalesRecordExist As Boolean
Private sFreqRegExp As String

Private objHours As clsPRFHOURS
'

Private Sub cmdApprove_Click()
    Dim nRow As Integer
    
    nRow = tgmApprove.GetCurrentRowNumber
    
    If tgsApprove.Count = 0 Then
        tgmApprove.CellValue(colAApprove, nRow) = colAppYes
    Else
        subSetAction colAppYes
    End If
    
    tgmApprove.Rebind
    nDataStatus = DATA_CHANGED
    
End Sub

Private Sub cmdApprove_GotFocus()
    tfnSetStatusBarMessage "Approve"
End Sub

Private Sub cmdCancel_Click(Index As Integer)
    subCancel Index
End Sub

Private Sub cmdCancel_GotFocus(Index As Integer)
    tfnSetStatusBarMessage "Cancel"
End Sub

Private Sub cmdDropDown_Click(Index As Integer)
    subSetFloatingSQL (Index)
    tgfDropdown(Index).ButtonClick
End Sub

Private Sub cmdDropDown_GotFocus(Index As Integer)
    tgfDropdown(Index).GotFocus cmdDropdown(Index)
End Sub

Private Sub cmdDropDown_LostFocus(Index As Integer)
    tgfDropdown(Index).LostFocus cmdDropdown(Index)
End Sub

Private Sub cmdOK_Click()
    Dim strSQL As String
    Dim vArrDetail() As Variant
    Dim sArBCodes() As String, sArCodeLvls() As String
    Dim rsTemp As Recordset
    Dim lEmpNo As Long

    #If PROTOTYPE Then
        Exit Sub
    #End If

    Dim i As Integer, j As Integer, k As Integer
    
    For i = 0 To tgmApprove.RowCount - 1
        If tgmApprove.CellValue(colAApprove, i) = colAppYes Then
            lEmpNo = tgmApprove.CellValue(colAEmpNo, i)
            'Get the amount for each pay code from the hidden column...
            sArBCodes = Split(tgmApprove.CellValue(colAHdnBAmtLvls, i), ",")
            'Get the SQL for loading the details...
            strSQL = "ABC"
            If fnLoadBonusDetails(lEmpNo, strSQL) Then
                If GetRecordSet(rsTemp, strSQL, , "cmdOK_Click") <= 0 Then
                    MsgBox "Failed to insert the record", vbExclamation
                    Exit Sub
                End If
                rsTemp.MoveFirst
                For j = 0 To rsTemp.RecordCount - 1
                    ReDim Preserve vArrDetail(4, j)
                    vArrDetail(0, j) = lEmpNo
                    vArrDetail(1, j) = fnGetField(rsTemp!bm_bonus_code)
                    If InStr(1, sArBCodes(j), "~") > 0 Then
                        'Add all the levels...
                        sArCodeLvls = Split(sArBCodes(j), "~")
                        For k = 0 To UBound(sArCodeLvls)
                            vArrDetail(2, j) = vArrDetail(2, j) + tfnRound(sArCodeLvls(k), 6)
                        Next k
                    Else
                        vArrDetail(2, j) = sArBCodes(j)
                    End If
                    vArrDetail(3, j) = 0
                    vArrDetail(4, j) = fnGetField(rsTemp!bm_eligible_date)
                    rsTemp.MoveNext
                Next j
                
                'Now Insert the values from the array created above into the hold table...
                For j = 0 To UBound(vArrDetail, 2)
                    If Not fnInsertHoldBonus(CLng(vArrDetail(0, j)), CStr(vArrDetail(1, j)), _
                            tfnRound(vArrDetail(2, j), 6), CLng(vArrDetail(3, j)), CStr(vArrDetail(4, j))) Then
                        Exit Sub
                    End If
                Next j
            End If
        End If
    Next i
    
    tfnResetScreen TabApprove
    tfnResetScreen TabProcess
    tfnResetScreen TabDetails
    eTabMain.CurrTab = TabProcess
    
End Sub

Private Sub cmdPrint_Click(Index As Integer)
    subPrint Index
End Sub

Private Sub cmdPrint_GotFocus(Index As Integer)
    tfnSetStatusBarMessage "Print"
End Sub

Private Sub cmdSelectAll_Click()
    If tgmApprove.RowCount > 0 Then
        tgsApprove.SelectAll
    End If
End Sub

Private Sub cmdSelectAll_GotFocus()
    tfnSetStatusBarMessage "Select All"
End Sub

Private Sub efraBaseHours_GotFocus()
    subSetFocus cmdAddBtn(nTabHours)
End Sub

Private Sub efraBaseIIDetail_GotFocus()
    subSetFocus txtEmployee
End Sub

Private Sub efraBaseIIHours_GotFocus()
    If objHours Is Nothing Then Exit Sub
    
    objHours.efraBridge_GotFocus
End Sub

Private Sub efraBaseIIProcess_GotFocus()
    subFillStartEndDateFreq
    subBuildFrequencyRegExp
    
    If txtStartDate.Enabled Then
        subSetFocus txtStartDate
    Else
        subSetFocus cmdProcess
    End If
End Sub

Private Sub efraBaseIISales_GotFocus()
    cValidSls.GotFocus efraBaseIISales
End Sub

Private Sub efraBaseIIView_GotFocus()
    subSetFocus tblApprove
End Sub

Private Sub efraBaseSales_GotFocus()
    subSetFocus cmdAddBtn(TabSales)
End Sub

Private Sub eTabMain_Click()
    Select Case eTabMain.CurrTab
        Case TabSales
            If eTabSub.CurrTab = TabSales Then
                subSetFocus efraBaseSales
            Else
                subSetFocus efraBaseHours
            End If
        Case TabProcess
            frmContext.ButtonEnabled(FO_HOLD_UP) = False
            subSetFocus efraBaseIIProcess
        Case TabApprove
            #If PROTOTYPE Then
                tblApprove.Enabled = False
                Exit Sub
            #End If
            frmContext.ButtonEnabled(FO_HOLD_UP) = False
            subSetFocus efraBaseIIView
        Case TabDetails
            frmContext.ButtonEnabled(FO_HOLD_UP) = False
            #If PROTOTYPE Then
                tblDetails.Enabled = False
                Exit Sub
            #End If
            subSetFocus efraBaseIIDetail
    End Select
    
    If eTabMain.CurrTab <> TabSales Then
        subEnablePrint eTabMain.CurrTab, cmdPrint(eTabMain.CurrTab).Enabled
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
    If Not CANCEL Then t_bTabSwitch = True
End Sub

Private Sub eTabSub_Click()
    Select Case eTabSub.CurrTab
        Case TabSales
            subSetFocus efraBaseSales
        Case TabHours
            subSetFocus efraBaseHours
    End Select
End Sub

Private Sub eTabSub_Switch(OldTab As Integer, NewTab As Integer, CANCEL As Integer)
    If Not eTabSub.TabEnabled(NewTab) Then
        CANCEL = True
        Exit Sub
    End If

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
    
    t_nFormMode = IDLE_MODE
    nDataStatus = DATA_INIT
    
    CRLF = Chr(10) + Chr(13)

    ' ** change the help file for the application
    App.HelpFile = szHelp7_11
End Sub

Private Sub Form_Unload(CANCEL As Integer)
    tfnUnlockRow
    
    On Error Resume Next
    
    Set objErrHandler = Nothing
    
    Set objCurrTabControl = Nothing
    
    Set objMath = Nothing
    Set objCond = Nothing
    
    Unload frmFORMULA
    Unload frmSplash
    Unload frmContext
    Unload frmAbout
    
    Set objHours = Nothing
    
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

Private Sub Form_Load()
    Dim sErrorMessage As String
    
    #If Not PROTOTYPE Then
        'If tfnAuthorizeExecute(Command) = False Then 'Check for handshake if not in the development mode
        '    Unload Me
        '    Exit Sub
        'End If
        
        'open the database, ODBC Dialog Box during developemnt, oleObject Connection String when not
        Screen.MousePointer = vbHourglass
        
        If Not tfnOpenDatabase(False, sErrorMessage) Then
            sErrorMessage = "Unable to open Database, Program terminates"
            subLogErrMsg "**System Error: " + sErrorMessage
            MsgBox sErrorMessage + ".", vbCritical
            Unload Me
            Exit Sub
        End If
        
        'connect to local database
        Set dbLocal = tfnOpenLocalDatabase(False, sErrorMessage)
        If dbLocal Is Nothing Then
            sErrorMessage = "Unable to open Local Database, Program terminates"
            subLogErrMsg "**System Error: " + sErrorMessage
            MsgBox sErrorMessage + ".", vbCritical
            Unload Me
            Exit Sub
        End If
    
        Screen.MousePointer = vbHourglass
    
        If Not fnCreateSearchTable("prm_empno", "prm_empname") Then
            sErrorMessage = "Failed to create temporary Employee Table. Program terminates"
            subLogErrMsg "**System Error: " + sErrorMessage
            MsgBox sErrorMessage + ".", vbCritical
            Unload Me
            Exit Sub
        End If
    
        If Not fnCreateTempTableVar() Then
            sErrorMessage = "Failed to create temporary Variable Table. Program terminates"
            subLogErrMsg "**System Error: " + sErrorMessage
            MsgBox sErrorMessage + ".", vbCritical
            Unload Me
            Exit Sub
        End If
    #End If
    
    tfnSetInitializingMessage
    Screen.MousePointer = vbHourglass
    
    subSetExitCancelBtn "EXIT"
    Screen.MousePointer = vbHourglass
    frmContext.ButtonEnabled(CANCEL_UP) = True
    mnuCancel.Enabled = True
    eTabMain.CurrTab = TabSales
    Me.Enabled = False
    
    subInitErrorHandler   ' Setup Error Control
    subInitSpreadsheets
    subSetFloatingDropDown TabSales
    subSetupCombos
    subInitValidation
    
    tfnDisableFormSystemClose Me
    subSetupToolBar
    
    tmrKeyBoard.Enabled = False
    tfnCenterForm Me
    
    Me.Show
    DoEvents
    
    tfnSetInitializingMessage
    Screen.MousePointer = vbHourglass
    
    'initialize the PRFHOURS class
    If Not fnInitialPRFHOURSclass() Then
        Unload Me
        Exit Sub
    End If
    
    Screen.MousePointer = vbHourglass
    
    Set objMath = New clsEquation
    Set objCond = New clsCondition
    
    #If Not PROTOTYPE Then
        tfnUpdateVersion
    #End If

    Me.Enabled = True
    
    t_bStartupFlag = False
    tmrKeyBoard.Enabled = True
    
    subEnablePrint 1, False, True
    tfnResetScreen TabProcess
    tfnResetScreen nTabHours
    tfnResetScreen TabSales
    
    Screen.MousePointer = vbDefault
    
    #If PROTOTYPE Then
        subSetProgress 100
        eTabMain.TabEnabled(TabApprove) = True
        eTabMain.TabEnabled(TabDetails) = True
    #End If
    
    subSetFocus cmdAddBtn(eTabMain.CurrTab)
    
End Sub

Private Sub Form_GotFocus()
    tmrKeyBoard.Enabled = True
End Sub

Private Sub Form_LostFocus()
    tmrKeyBoard.Enabled = False
End Sub

Private Sub cmdExitCancelBtn_GotFocus()
    If cmdExitCancelBtn.Caption = t_szCAPTION_EXIT Then
        tfnSetStatusBarMessage t_szEXIT
    Else
        tfnSetStatusBarMessage t_szCANCEL
    End If
End Sub

Private Sub cmdExitCancelBtn_Click()

    Screen.MousePointer = vbHourglass
    
    If cmdExitCancelBtn.Caption = t_szCAPTION_CANCEL Then
        subCancel eTabMain.CurrTab
    Else
        subExit
    End If
    
End Sub

'=====================
'Toolbar Button Events
'=====================
Private Sub subCancel(Index As Integer)
    
    #If PROTOTYPE Then
        Exit Sub
    #End If
    
    If Index = TabProcess Then
        If cmdProcess.Enabled = False And cValidate.FirstInvalidInput < 0 Then
            bCancelProcess = True
            Exit Sub
        End If
    End If

    tfnResetScreen Index
    Screen.MousePointer = vbDefault

End Sub

Private Sub subExit()
    If t_bDataChanged Then
        If Not tfnCancelExit(t_szEXIT_MESSAGE) Then
            Exit Sub
        End If
    End If
    
    Unload Me
End Sub

Private Sub lstProcess_GotFocus()
    tfnSetStatusBarMessage "Process Status"
End Sub

'============
'Menu  Events
'============
Private Sub mnuExit_Click()
    subExit
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
    subCancel eTabMain.CurrTab
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
Public Sub tfnSetStatusBarMessage(szMessage As String)
    
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
Private Sub tfnResetScreen(Index As Integer)
    Dim i As Integer
    
    On Error Resume Next
    
    Screen.MousePointer = vbHourglass
    subSetFocus efraBackground
    
    Select Case Index
        Case TabSales
            If nDataStatus = DATA_CHANGED Then
                If Not tfnCancelExit(t_szCANCEL_MESSAGE) Then
                    Exit Sub
                End If
            End If
            
            nDataStatus = DATA_INIT
            txtFromDate = ""
            txtToDate = ""
            cValidSls.ResetFlags
            tgmSales.ClearData
            subEnableFirstLineSlsOrHrs Index, False
            tblSales.Enabled = False
            
            frmContext.ButtonEnabled(CANCEL_UP) = False
            cmdCancel(TabSales).Enabled = False
            mnuCancel.Enabled = False
            cmdAddBtn(Index).Enabled = True
            cmdEditBtn(Index).Enabled = True
            eTabMain.TabEnabled(TabProcess) = True
            eTabSub.TabEnabled(TabHours) = True
            subSetFocus cmdAddBtn(Index)
            cmdUpdateInsertBtn(Index).Caption = t_szCAPTION_UPDATE
        Case nTabHours
            objHours.mnuCancel_Click
            eTabMain.TabEnabled(TabProcess) = True
            eTabSub.TabEnabled(TabSales) = True
        Case TabProcess
            If nDataStatus = DATA_CHANGED Then
                If Not tfnCancelExit(t_szCANCEL_MESSAGE) Then
                    Screen.MousePointer = vbDefault
                    Exit Sub
                End If
            End If
            nDataStatus = DATA_INIT
            txtStartDate = ""
            txtEndDate = ""
            txtFrequency = ""
            txtPrftCtr = ""
            txtPrftCtrName = ""
            txtEmpProcess = ""
            txtEmpNameProcess = ""
            lstProcess.Clear
    
            cValidate.ResetFlags
            
            eTabMain.TabEnabled(TabSales) = True
            eTabMain.TabEnabled(TabDetails) = False
            eTabMain.TabEnabled(TabApprove) = False
            subEnablePrint Index, False
            subEnableFirstLineProcess True
            bCancelProcess = False
            'cmdProcess.Enabled = False
            subSetProgress 0
            
            If eTabMain.CurrTab = TabProcess Then
                subFillStartEndDateFreq
                subSetFocus txtStartDate
            End If
        Case TabApprove
            If nDataStatus = DATA_CHANGED Then
                If tfnCancelExit(t_szCANCEL_MESSAGE) Then
                    tgmApprove.ClearData
                    nDataStatus = DATA_INIT
                    fnFillApproveTab vArrBonus
                    If eTabMain.CurrTab = TabApprove Then
                        subSetFocus tblApprove
                    End If
                End If
            End If
        Case TabDetails
            txtEmployee = ""
            txtEmpName = ""
            tgmDetail.ClearData
            subEnableEmployee True
            tblDetails.Enabled = False
            subEnablePrint Index, False
            If eTabMain.CurrTab = TabDetails Then
                subSetFocus txtEmployee
            End If
    End Select
    
    frmContext.ButtonEnabled(COPY_UP) = False
    frmContext.ButtonEnabled(FO_HOLD_UP) = False
    mnuExit.Enabled = True
    mnuCopy.Enabled = False
    mnuPaste.Enabled = False
    cmdRefresh(Index).Enabled = False
    cmdUpdateInsertBtn(Index).Enabled = False
    cmdDelete(Index).Enabled = False
    
    Screen.MousePointer = vbDefault
    
End Sub

Private Sub mnuPrint_Click()
    subPrint eTabMain.CurrTab
End Sub

Private Sub optType_GotFocus(Index As Integer)
    Select Case Index
        Case nWeek
            tfnSetStatusBarMessage "Weekly Sales"
        Case nOneMth
            tfnSetStatusBarMessage "One Monthly Sales"
        Case nThreeMth
            tfnSetStatusBarMessage "Quaterly Sales"
        Case nGas
            tfnSetStatusBarMessage "Gas Sales"
    End Select
End Sub

Private Sub optType_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        subSetFocus txtFromDate
    End If
End Sub

Private Sub tblApprove_AfterColEdit(ByVal ColIndex As Integer)
    tgmApprove.AfterColEdit ColIndex
End Sub

Private Sub tblApprove_BeforeColEdit(ByVal ColIndex As Integer, ByVal KeyAscii As Integer, CANCEL As Integer)
    tgmApprove.BeforeColEdit ColIndex, KeyAscii, CANCEL
End Sub

Private Sub tblApprove_Change()
    tgmApprove.Change
End Sub

Private Sub tblApprove_Click()
    tgsApprove.Click
End Sub

Private Sub tblApprove_FirstRowChange()
    tgmApprove.FirstRowChange
End Sub

Private Sub tblApprove_GotFocus()
    tfnSetStatusBarMessage "Press enter key to see commission details"
    tgsApprove.GotFocus
    tgmApprove.GotFocus
End Sub

Private Sub tblApprove_KeyDown(KeyCode As Integer, Shift As Integer)
    tgsApprove.KeyDown KeyCode, Shift
    tgmApprove.KeyDown KeyCode, Shift
End Sub

Private Sub tblApprove_KeyPress(KeyAscii As Integer)
    Dim nRow As Integer
    
    If KeyAscii <> vbKeyReturn Then
        Exit Sub
    End If
    
    If tgsApprove.Count > 1 Then
        MsgBox "Only one detail can be viewed at a time", vbInformation
        Exit Sub
    End If
    
    nRow = tgmApprove.GetCurrentRowNumber
    txtEmployee = tgmApprove.CellValue(colAEmpNo, nRow)
    txtEmpName = tgmApprove.CellValue(colAEmpName, nRow)
    subEnterBonusPhaseII
    
End Sub

Private Sub tblApprove_LostFocus()
    tgmApprove.LostFocus
End Sub

Private Sub tblApprove_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    tgsApprove.MouseUp Button, Shift, Y
End Sub

Private Sub tblApprove_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    tgmApprove.RowColChange LastRow, LastCol
    tgsApprove.RowColChange LastRow, LastCol
End Sub

Private Sub tblApprove_SelChange(CANCEL As Integer)
    tgsApprove.SelChange CANCEL
    CANCEL = True
End Sub

Private Sub tblApprove_UnboundReadData(ByVal RowBuf As DBTrueGrid.RowBuffer, StartLocation As Variant, ByVal ReadPriorRows As Boolean)
    tgmApprove.ReadData RowBuf, StartLocation, ReadPriorRows
End Sub

Private Sub tblDetails_GotFocus()
    tfnSetStatusBarMessage "Press enter key to see formula details"
End Sub

Private Sub tblDetails_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        subShowFormulaDetails
    End If
End Sub

Private Sub tblDetails_UnboundReadData(ByVal RowBuf As DBTrueGrid.RowBuffer, StartLocation As Variant, ByVal ReadPriorRows As Boolean)
    tgmDetail.ReadData RowBuf, StartLocation, ReadPriorRows
End Sub

Private Sub tblSales_Click()
    tgsSales.Click
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
            Set .LocalDatabase = dbLocal
        End With
    End If
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
            'for PRFHOURS
            .AddButton "Add &Profit Center", PROFITCENTER_UP
            .AddButton "Add E&mployee", EMP_MST_UP, , True
            
            .AddButton "Add Commission &Code", PRDCLS_UP
            .AddButton "Add Co&mmission Formula", SYS_LOCKS_UP
            .AddButton "&Launch Commssion Master", POFAPLVL_UP
            .AddButton "&Export to Payroll", PRAPPROV_UP
            .AddButton "View Formula Details", FO_HOLD_UP, , True
        .EndSetupToolbar
    
        .HelpFile = szHelpFileName
    End With
End Sub

Public Sub TBButtonCallBack(ByVal nID As Integer)
    Select Case nID
        Case CANCEL_UP
            If eTabMain.CurrTab = 0 Then
                If eTabSub.CurrTab = TabSales Then
                    subCancel eTabSub.CurrTab
                Else
                    subCancel nTabHours
                End If
            Else
                subCancel eTabMain.CurrTab
            End If
        Case EXIT_UP
            If eTabMain.CurrTab = 0 And eTabSub.CurrTab = TabHours Then
                If objHours.fnExit() Then
                    subExit
                End If
            Else
                subExit
            End If
        Case FO_HOLD_UP  'Approve PR
            subShowFormulaDetails
        Case PRINT_UP
            subPrint eTabMain.CurrTab
        Case Else
            objHours.TBButtonCallBack nID
    End Select
End Sub

Private Sub mnuModules_Click(Index As Integer)
    frmContext.MenuClick Index
End Sub

Private Sub tbToolbar_ButtonClick(ByVal Button As Button)
    frmContext.ButtonClick Button
End Sub

Private Sub tbToolbar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
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

Private Sub subSetProgress(sngPercent As Single)
    If sngPercent > 100 Then sngPercent = 100
    If sngPercent > 0# Then
        If Not pbBarMain.Visible Then
            pbBarMain.ZOrder 0
            pbBarMain.Visible = True
        End If
    Else
        'pbBarMain.Visible = False
        pbBarMain.Value = 0
        If pbBarMain.ToolTipText = "" Then
            pbBarMain.ToolTipText = "Process Checks progress bar"
        End If
    End If
    
    pbBarMain.Value = sngPercent
    pbBarMain.Refresh
End Sub

Private Function fnCheckCancel() As Boolean
    DoEvents
    fnCheckCancel = False
    
    If bCancelProcess Then
        If MsgBox("Are you sure you want to cancel the process?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
            bCancelProcess = False
            Exit Function
        End If
    End If
     
    fnCheckCancel = bCancelProcess
    
End Function

Private Sub cmdProcess_Click()
    Const SUB_NAME As String = "cmdProcess_Click"
    Dim strSQL As String
    Dim rsTemp As Recordset
    Dim nCount As Integer
    Dim sEmpNo As String, sCode As String
    Dim dBLvlAmt As Double
    Dim dTotalBonus As Double
    Dim sAmtAllLevels As String, sAmtAllBCodes As String
    Dim nSize As Integer: nSize = -1
    Dim i As Integer
    
    #If PROTOTYPE Then
        Exit Sub
    #End If
    
    If Not tfnCancelExit("Processing may take several minutes, are you sure you want to continue?") Then
        Exit Sub
    End If
    
    ReDim vArrBonus(colAHdnBAmtLvls, 0)
    eTabMain.TabEnabled(TabSales) = False
    bCancelProcess = False
    cmdProcess.Enabled = False
    eTabMain.TabEnabled(TabApprove) = False
    eTabMain.TabEnabled(TabDetails) = False
    subEnablePrint TabProcess, False
    subEnableFirstLineProcess False
    
    lstProcess.Clear
    subLogErrMsg "Started processing commission formulas..."
    subLogErrMsg " "
    
    strSQL = "SELECT bm_empno, bc_type, bc_grade, bc_bonus_code, bm_sequence, bf_level"
    strSQL = strSQL & " FROM bonus_master, bonus_codes, bonus_formula"
    strSQL = strSQL & " WHERE bm_bonus_code = bc_bonus_code"
    strSQL = strSQL & " AND bm_bonus_code = bf_bonus_code"
    If cValidate.ValidInput(txtPrftCtr) And txtPrftCtr <> "" Then
        strSQL = strSQL & " AND bm_eligible_pc = " & Trim(txtPrftCtr)
    End If
    If cValidate.ValidInput(txtFrequency) And txtFrequency <> "" Then
        strSQL = strSQL & " AND bc_frequency = " & tfnSQLString(Trim(txtFrequency))
    End If
    strSQL = strSQL & " AND " & tfnDateString(txtEndDate, True)
    strSQL = strSQL & " BETWEEN bm_eligible_date AND bm_stop_date"
'    strSQL = strSQL & " AND " & tfnDateString(Date, True)
'    strSQL = strSQL & " BETWEEN bm_eligible_date AND bm_stop_date"
    strSQL = strSQL & " ORDER BY bm_empno, bc_bonus_code, bm_sequence, bf_level"
    
    Screen.MousePointer = vbHourglass
    nCount = GetRecordSet(rsTemp, strSQL, , SUB_NAME)
    If nCount < 0 Then
        subLogErrMsg "Failed to access the database"
        GoTo TERMINATE_PROCESS
    End If
    
    If nCount = 0 Then
        subLogErrMsg "No record found to process"
        GoTo TERMINATE_PROCESS
    End If
    
    rsTemp.MoveFirst
    For i = 1 To nCount
        DoEvents
        If bCancelProcess Then GoTo TERMINATE_PROCESS
        subSetProgress i * (100 / nCount)
        Screen.MousePointer = vbHourglass
        subLogErrMsg "Calculating commission for employee " & tfnSQLString(rsTemp!bm_empno) _
                   & ", commission code " & tfnSQLString(rsTemp!bc_bonus_code) _
                   & " and level " & tfnSQLString(Trim(rsTemp!bf_level))
        If sEmpNo <> CStr(rsTemp!bm_empno) Then
            If i > 1 Then
                nSize = nSize + 1
                ReDim Preserve vArrBonus(colAHdnBAmtLvls, nSize)
                vArrBonus(colAEmpNo, nSize) = Trim(sEmpNo)
                vArrBonus(colAEmpName, nSize) = fnGetEmployeeName(sEmpNo)
                vArrBonus(colADate, nSize) = txtEndDate
                vArrBonus(colABonusAmt, nSize) = Format(dTotalBonus, "##,##0.00")
                vArrBonus(colAHdnBAmtLvls, nSize) = Trim(sAmtAllBCodes) 'Hidden Column
            End If
            dBLvlAmt = fnGetBonusAmount(rsTemp)
            dTotalBonus = dBLvlAmt
            sAmtAllBCodes = CStr(dBLvlAmt)
        Else
            dBLvlAmt = fnGetBonusAmount(rsTemp)
            If sCode <> CStr(rsTemp!bc_bonus_code) Then
                sAmtAllBCodes = sAmtAllBCodes & "," & CStr(dBLvlAmt) & ""
            Else
                sAmtAllBCodes = sAmtAllBCodes & "~" & CStr(dBLvlAmt) & ""
            End If
            dTotalBonus = dTotalBonus + dBLvlAmt
        End If
        sCode = fnGetField(rsTemp!bc_bonus_code)
        sEmpNo = CStr(rsTemp!bm_empno)
        'last record...
        If i = nCount Then
            nSize = nSize + 1
            ReDim Preserve vArrBonus(colAHdnBAmtLvls, nSize)
            vArrBonus(colAEmpNo, nSize) = Trim(sEmpNo)
            vArrBonus(colAEmpName, nSize) = fnGetEmployeeName(sEmpNo)
            vArrBonus(colADate, nSize) = txtEndDate
            vArrBonus(colABonusAmt, nSize) = Format(dTotalBonus, "##,##0.00")
            vArrBonus(colAHdnBAmtLvls, nSize) = Trim(sAmtAllBCodes) 'Hidden Column
        End If
        rsTemp.MoveNext
    Next i
    
    If fnFillApproveTab(vArrBonus) Then
        eTabMain.TabEnabled(TabApprove) = True
        'eTabMain.CurrTab = TabApprove
        subSetFocus tblApprove
    End If
    
TERMINATE_PROCESS:
    subLogErrMsg " "
    If bCancelProcess Then subLogErrMsg "Processing terminated on user's request"
    subLogErrMsg "*Finished Processing*"
    Screen.MousePointer = vbDefault
    'tfnSetStatusBarMessage "Finished Processing"
    subSetProgress 0
    cmdProcess.Enabled = True
    subEnablePrint TabProcess, True
    subSetFocus cmdProcess
End Sub

Private Sub cmdProcess_GotFocus()
    tfnSetStatusBarMessage "Process"
End Sub

Private Sub subSetGridWidth(tbl As TDBGrid)
    Dim myWidth As Variant
    Dim myField As Variant
    Dim VItem As New ValueItem
    Dim vitems As ValueItems
    Dim i As Integer
    
    Select Case tbl.Name
        Case "tblSales"
            myWidth = Array(0.12, 0.43, 0.15, 0.15, 0.15)
            myField = Array("prft_ctr", "prft_name", "amount", "from_date", "to_date")
        Case "tblTimeCard"
            myWidth = Array(0.21, 0.21, 0.16, 0.16, 0.26)
            myField = Array("prh_date", "prh_prft_ctr", "prh_pay_code", "prh_pay_type", "prh_hours")
        Case "tblProfitCenter"
            myWidth = Array(0.5, 0.5)
            myField = Array("", "")
        Case "tblApprove"
            'myWidth = Array(0.1, 0.13, 0.32, 0.1, 0.11, 0.12, 0.12)
            'myField = Array("", "", "", "", "", "", "")
            myWidth = Array(0.1, 0.1, 0.13, 0.22, 0.1, 0.11, 0.12, 0.12)
            myField = Array("", "", "", "", "", "", "", "")
        Case "tblDetails"
            myWidth = Array(0.07, 0.4, 0.07, 0.07, 0.11, 0.13, 0.15)
            myField = Array("bm_bonus_code", "bc_code_desc", "bf_level", "bc_type", "bc_frequency", "bm_eligible_date", "")
    End Select
    
    While tbl.Columns.Count > 0
        tbl.Columns.Remove 0
    Wend
    
    tbl.ExtendRightColumn = True
    
    For i = 0 To UBound(myWidth)
        tbl.Columns.Add i
        With tbl.Columns(i)
            .Width = myWidth(i) * (tbl.Width - 255)
            .DataField = myField(i)
            .Visible = True
            .HeadAlignment = vbCenter
        End With
    Next
    
    Select Case tbl.Name
        Case "tblSales"
            tbl.Caption = "Store Sales"
            tbl.Columns(colSPrftCtr).Caption = "Profit Ctr"
            tbl.Columns(colSPrftName).Caption = "Profit Center Name"
            tbl.Columns(colSAmount).Caption = "Amount"
            tbl.Columns(colSAmount).Alignment = vbRightJustify
            tbl.Columns(colSFromDate).Caption = "From Date"
            tbl.Columns(colSToDate).Caption = "To Date"
        Case "tblTimeCard"
            tbl.Caption = "Time Card Entry"
            tbl.Columns(colHClockIn).Caption = "Clock-In Date"
            tbl.Columns(colHPrftCtr).Caption = "Profit Center"
            tbl.Columns(colHPayCode).Caption = "Pay Code"
            tbl.Columns(colHPayType).Caption = "Pay Type"
            tbl.Columns(colHHrsDol).Caption = "Hours/Dollars"
            tbl.Columns(colHHrsDol).Alignment = vbRightJustify
        Case "tblProfitCenter"
            tbl.Caption = "Profit Center Total"
            tbl.Columns(colPProfit).Caption = "Profit"
            tbl.Columns(colPTotal).Caption = "Total"
            tbl.Columns(colPTotal).Alignment = vbRightJustify
        Case "tblApprove"
            Set vitems = tbl.Columns(colAApprove).ValueItems
            VItem.Value = colAppYes: VItem.DisplayValue = "Y": vitems.Add VItem
            VItem.Value = colAppNo: VItem.DisplayValue = "N": vitems.Add VItem
            vitems.Presentation = 1
            vitems.CycleOnClick = False
            vitems.Translate = True
            vitems.DefaultItem = colAppNo
            tbl.Caption = "Commission Approval"
            tbl.Columns(colAApprove).Caption = "Approve"
            tbl.Columns(colAPrftCtr).Caption = "Profit Ctr"
            tbl.Columns(colAEmpNo).Caption = "Employee No."
            tbl.Columns(colAEmpName).Caption = "Employee Name"
            tbl.Columns(colAPayCode).Caption = "Pay Code"
            tbl.Columns(colAPayHours).Caption = "Pay Hours"
            tbl.Columns(colADate).Caption = "Date"
            tbl.Columns(colABonusAmt).Caption = "Amount"
            tbl.Columns(colABonusAmt).Alignment = vbRightJustify
        Case "tblDetails"
            tbl.Caption = "Commission Details"
            tbl.Columns(colDBCode).Caption = "Code"
            tbl.Columns(colDBCDesc).Caption = "Code Description"
            tbl.Columns(colDBLevel).Caption = "Level"
            tbl.Columns(colDBLevel).Alignment = vbCenter
            tbl.Columns(colDBType).Caption = "Type"
            tbl.Columns(colDBType).Alignment = vbCenter
            tbl.Columns(colDBFreq).Caption = "Frequency"
            tbl.Columns(colDBFreq).Alignment = vbCenter
            tbl.Columns(colDElgDate).Caption = "Eligible Date"
            tbl.Columns(colDBAmt).Caption = "Amount"
            tbl.Columns(colDBAmt).Alignment = vbRightJustify
    End Select
End Sub

Private Sub subInitSpreadsheets()
    Dim sDecimalString As String
    
    'Table Sales Class Implementation
    sDecimalString = tfnDecimalPattern(10, 2)
    subSetGridWidth tblSales
    Set tgmSales = New clsTGSpreadSheet
    Set tgmSales.Table = tblSales
    Set tgmSales.StatusBar = ffraStatusbar ' message bar name
    Set tgmSales.Form = Me
    Set tgmSales.engFactor = t_engFactor
    tgmSales.AddEditColumn colSPrftCtr, "Enter Profit Center", szIntegerPattern
    tgmSales.AddEditColumn colSPrftName, "Enter Profit Center Name", "^P{1,40}"
    tgmSales.AddEditColumn colSAmount, "Enter Amount", sDecimalString
    'tgmSales.AddEditColumn colSFromDate, "Enter From Date", szDatePattern
    'tgmSales.AddEditColumn colSToDate, "Enter To Date", szDatePattern
    tgmSales.DisplayFormat(colSAmount) = "###,###,##0.00"
    ColxSOldPrftCtr = tgmSales.AddHiddenField("old_prft_ctr")
    tgmSales.AllowAddNew = True
    
    'Implement the selector class
    Set tgsSales = New clsTGSelector
    tgsSales.AvoidBeep = False
    Set tgsSales.EditorClass = tgmSales
    tgsSales.SelectCurrRow = False
    tgsSales.AllowMultipleSelect = True
    tgsSales.RowHighLighted = True
    
    'setup Time Card and Profit Center Grid
    'the class implementation will be done in clsPRFHOURS
    subSetGridWidth tblTimeCard
    subSetGridWidth tblProfitCenter
    
    'Table Approve Class Implementation
    subSetGridWidth tblApprove
    Set tgmApprove = New clsTGSpreadSheet
    Set tgmApprove.Table = tblApprove
    Set tgmApprove.StatusBar = ffraStatusbar ' message bar name
    Set tgmApprove.Form = Me
    Set tgmApprove.engFactor = t_engFactor
    tgmApprove.AddEditColumn colAApprove, "Select Yes, No"
    tgmApprove.AllowAddNew = False
    colAHdnBAmtLvls = tgmApprove.AddHiddenField("HiddenLevels")
    
    'Implement the selector class
    Set tgsApprove = New clsTGSelector
    tgsApprove.AvoidBeep = False
    Set tgsApprove.EditorClass = tgmApprove
    tgsApprove.SelectCurrRow = False
    tgsApprove.RowHighLighted = True
    
    'Table Detail Class Implementation
    subSetGridWidth tblDetails
    Set tgmDetail = New clsTGSpreadSheet
    Set tgmDetail.Table = tblDetails
    Set tgmDetail.StatusBar = ffraStatusbar ' message bar name
    Set tgmDetail.Form = Me
    Set tgmDetail.engFactor = t_engFactor
    tgmDetail.SetupTable True
    tgmDetail.ClearData
    tgmDetail.AllowAddNew = False

End Sub

Private Sub txtEmployee_Change()
    cValidate.Change txtEmployee
    tgcDropdown.Change txtEmployee
End Sub

Private Sub txtEmployee_GotFocus()
    tgcDropdown.GotFocus txtEmployee
    cValidate.GotFocus txtEmployee
    
    If tgcDropdown.SingleRecordSelected Then
        subEnterBonusPhaseII
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
                subEnterBonusPhaseII
                Screen.MousePointer = vbDefault
          End If
          KeyAscii = 0
       End If
    Else
        cValidate.Keypress txtEmployee, KeyAscii
    End If

End Sub

Private Sub txtEmployee_LostFocus()
    tgcDropdown.LostFocus txtEmployee
    cValidate.LostFocus txtEmployee, cmdEmployee
    Screen.MousePointer = vbDefault

End Sub

Private Sub txtEmpName_Change()
    tgcDropdown.Change txtEmpName
End Sub

Private Sub txtEmpName_GotFocus()
    tfnSetStatusBarMessage "Enter Employee Name"
    tgcDropdown.GotFocus txtEmpName
    Screen.MousePointer = vbDefault
    
    If tgcDropdown.SingleRecordSelected Then
        subEnterBonusPhaseII
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
                subEnterBonusPhaseII
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

    'Class implementation for Process Tab
    Set cValidate = New cValidateInput
    Set cValidate.Form = Me
    Set cValidate.StatusBar = ffraStatusbar
    cValidate.AddEditBox txtStartDate, "Enter Starting Date"
    cValidate.AddEditBox txtEndDate, "Enter Ending Date"
    cValidate.AddEditBox txtFrequency, "Enter Frequency"
    cValidate.AddEditBox txtPrftCtr, "Enter Profit Center Number"
    cValidate.AddEditBox txtEmpProcess, "Enter Employee Number"
    cValidate.AddEditBox txtEmployee, "Enter Employee Number"
    cValidate.MinTabIndex = txtStartDate.TabIndex
    cValidate.MaxTabIndex = txtEmpName.TabIndex
    cValidate.ESCControl = cmdCancel(TabProcess)
    
    'Class implementation for Sales Tab
    Set cValidSls = New cValidateInput
    Set cValidSls.Form = Me
    Set cValidSls.StatusBar = ffraStatusbar
    cValidSls.AddEditBox txtFromDate, "Enter From Date"
    cValidSls.AddEditBox txtToDate, "Enter To Date"
    cValidSls.MinTabIndex = optType(nWeek).TabIndex
    cValidSls.MaxTabIndex = tblSales.TabIndex
    Set cValidSls.ControlForFocus = efraBaseIISales
    Set cValidSls.LastBox = txtToDate
    cValidSls.SetFirstControls cmdDelete(TabSales), cmdRefresh(TabSales), cmdCancel(TabSales), cmdUpdateInsertBtn(TabSales), cmdExitCancelBtn
    
End Sub

Private Function fnSetComboSQL(nTabIndex As Integer) As String
    Dim strSQL As String
    
    Select Case nTabIndex
        Case txtPrftCtr.TabIndex, txtPrftCtrName.TabIndex
            strSQL = "SELECT prft_ctr, prft_name FROM sys_prft_ctr WHERE prft_ctr IN "
            strSQL = strSQL & "(SELECT DISTINCT bm_eligible_pc FROM bonus_master)"
        Case txtFrequency.TabIndex
            strSQL = "SELECT bfq_frequency, bfq_freq_desc FROM bonus_frequency"
        Case txtEmployee.TabIndex, txtEmpName.TabIndex, txtEmpProcess.TabIndex, txtEmpNameProcess.TabIndex
            strSQL = "SELECT prm_empno, prm_empname FROM sTmpEmpTable"
            strSQL = strSQL & " WHERE prm_empno IN (SELECT bm_empno FROM bonus_master)"
        Case txtFromDate.TabIndex
            strSQL = fnGetSalesSQL(txtFromDate)
        Case txtToDate.TabIndex
            strSQL = fnGetSalesSQL(txtToDate)
    End Select
    fnSetComboSQL = strSQL
End Function

Public Function fnInValidData(txtBox As Textbox) As Boolean
    #If PROTOTYPE Then
        Exit Function
    #End If

    fnInValidData = True

    Select Case txtBox.TabIndex
        Case txtStartDate.TabIndex, txtEndDate.TabIndex
            fnInValidData = Not fnValidProcessDate(txtBox)
        Case txtFrequency.TabIndex
            fnInValidData = Not fnValidBonusFreq(txtBox)
        Case txtPrftCtr.TabIndex
            fnInValidData = Not fnValidPrftCtr(txtBox)
        Case txtEmployee.TabIndex, txtEmpProcess.TabIndex
            fnInValidData = Not fnValidEmployee(txtBox)
        Case txtEmployeeNumber.TabIndex, txtEmployeeName.TabIndex, txtSSN.TabIndex
            fnInValidData = objHours.fnInValidData(txtBox)
        Case txtFromDate.TabIndex, txtToDate.TabIndex
            fnInValidData = Not fnValidSalesDate(txtBox)
    End Select
End Function

Private Function fnValidBonusFreq(txtBox As Textbox) As Boolean
    Const SUB_NAME As String = "fnValidBonusFreq"
    Dim strSQL As String

    fnValidBonusFreq = False
    
    If Trim(txtBox) = "" Then
        cValidate.SetErrorMessage txtBox, "You must enter a Commission frequency"
        Exit Function
    End If
    
    strSQL = "SELECT * FROM bonus_frequency WHERE bfq_frequency = " & tfnSQLString(txtBox)
    
    If GetRecordCount(strSQL, , SUB_NAME) <= 0 Then
        cValidate.SetErrorMessage txtBox, "Commission frequency does not exist"
        Exit Function
    End If
    
    fnValidBonusFreq = True

End Function

Private Sub subBuildFrequencyRegExp()
    Const SUB_NAME As String = "subBuildFrequencyRegExp"
    
    Dim strSQL As String
    Dim rsTemp As Recordset
    
    If sFreqRegExp <> "" Then
        Exit Sub
    End If
    
    strSQL = "SELECT * FROM bonus_frequency"
    
    If GetRecordSet(rsTemp, strSQL, , SUB_NAME) < 0 Then
        MsgBox "Failed to access Database.", vbExclamation
        Exit Sub
    End If
    
    If rsTemp.RecordCount = 0 Then
        MsgBox "Employee Number does not exist", vbExclamation
        Exit Sub
    End If
    
    On Error GoTo errTrap
    
    sFreqRegExp = "^([" + fnCstr(rsTemp!bfq_frequency)
    rsTemp.MoveNext
    
    While Not rsTemp.EOF
        sFreqRegExp = sFreqRegExp + fnCstr(rsTemp!bfq_frequency)
        rsTemp.MoveNext
    Wend
    
    sFreqRegExp = sFreqRegExp + "])$"
    
    Exit Sub
    
errTrap:
    sFreqRegExp = "^P$"
End Sub

Private Function fnValidEmployee(txtBox As Textbox) As Boolean
    Const SUB_NAME As String = "fnValidEmployee"
    Dim strSQL As String
    Dim rsTemp As Recordset
    Dim sEmpName As String
    
    fnValidEmployee = False
    
    If Trim(txtBox.Text) = "" Then
        If eTabMain.CurrTab = TabDetails Then
            cValidate.SetErrorMessage txtBox, "You must enter an Employee Number"
        Else
            fnValidEmployee = True
        End If
        Exit Function
    End If
    
    If Not IsNumeric(Trim(txtBox.Text)) Then
        cValidate.SetErrorMessage txtBox, "Employee Number does not exist"
        Exit Function
    End If
    
    strSQL = "SELECT * FROM sTmpEmpTable WHERE prm_empno = " & tfnRound(txtBox)
    
    If GetRecordSet(rsTemp, strSQL, , SUB_NAME) < 0 Then
        cValidate.SetErrorMessage txtBox, "Failed to access Database"
        Exit Function
    End If
    
    If rsTemp.RecordCount = 0 Then
        cValidate.SetErrorMessage txtBox, "Employee Number does not exist"
        Exit Function
    End If
    
    sEmpName = fnCstr(rsTemp!prm_empname)
    
    strSQL = "SELECT bm_empno FROM bonus_master"
    strSQL = strSQL & " WHERE bm_empno = " & tfnRound(txtBox)
    strSQL = strSQL & " GROUP BY bm_empno"
    
    If GetRecordSet(rsTemp, strSQL, , SUB_NAME) < 0 Then
        cValidate.SetErrorMessage txtBox, "Failed to access Database"
        Exit Function
    End If
            
    If rsTemp.RecordCount = 0 Then
        cValidate.SetErrorMessage txtBox, "Employee is not set up in Bonus Master"
        Exit Function
    End If
    
    fnValidEmployee = True
    
End Function

Private Sub tblComboDropDown_Click()
    tgcDropdown.Click tblComboDropdown
End Sub

Private Sub tblComboDropDown_GotFocus()
    tgcDropdown.GotFocus tblComboDropdown
End Sub

Private Sub tblComboDropDown_LostFocus()
    tgcDropdown.LostFocus tblComboDropdown
End Sub

Private Sub tblComboDropDown_KeyPress(KeyAscii As Integer)
    tgcDropdown.Keypress tblComboDropdown, KeyAscii
    Exit Sub
    Dim bCode As Boolean
    
    bCode = tgcDropdown.Keypress(tblComboDropdown, KeyAscii)
    
    If Not bCode Then
        KeyAscii = 0
    End If
    
End Sub

Private Sub tblComboDropDown_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    tgcDropdown.RowColChange
End Sub

Private Sub tblComboDropDown_SelChange(CANCEL As Integer)
    tgcDropdown.SelChange CANCEL
End Sub

Private Sub tblComboDropDown_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    tgcDropdown.TableMouseUp Y
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
        .AddComboBox txtPrftCtr, cmdPrftCtr, "prft_ctr", .SQL_INT_TYPE
        .AddComboBox txtPrftCtrName, cmdPrftCtrName, "prft_name", .SQL_STRING_TYPE(40)
     
        .AddCombo
        .AddComboBox txtEmpProcess, cmdEmpProcess, "prm_empno", .SQL_LONG_TYPE
        .AddComboBox txtEmpNameProcess, cmdEmpNameProcess, "prm_empname", .SQL_STRING_TYPE(60)
     
        .AddCombo
        .AddComboBox txtFrequency, cmdFrequency, "bfq_frequency", .SQL_STRING_TYPE(1)
        .AddExtraColumn "bfq_freq_desc", 1300
        .SetExtend txtFrequency, 2
        
        .AddCombo
        .AddComboBox txtEmployee, cmdEmployee, "prm_empno", .SQL_LONG_TYPE
        .AddComboBox txtEmpName, cmdEmpName, "prm_empname", .SQL_STRING_TYPE(60)
        
        .AddCombo
        .AddComboBox txtFromDate, cmdFromDate, "bs_from_date", .SQL_DATE_TYPE
        .SetOrderingDescent txtFromDate

        .AddCombo
        .AddComboBox txtToDate, cmdToDate, "bs_to_date", .SQL_DATE_TYPE
        .SetOrderingDescent txtToDate
     End With
End Sub

Private Sub subEnterBonusPhaseII()
    Dim sArBCodes() As String
    Dim sArCodeLvls() As String
    Dim nRow As Integer
    Dim i As Integer, j As Integer, k As Integer

    nRow = tgmApprove.GetCurrentRowNumber
    Screen.MousePointer = vbHourglass
    subEnableEmployee False
    tblDetails.Enabled = True
    
    If Not fnLoadBonusDetails(txtEmployee) Then
        cmdCancel_Click (TabDetails)
        Exit Sub
    End If
    
    sArBCodes = Split(tgmApprove.CellValue(colAHdnBAmtLvls, nRow), ",")
    For i = 0 To UBound(sArBCodes)
        sArCodeLvls = Split(sArBCodes(i), "~")
        For j = 0 To UBound(sArCodeLvls)
            tgmDetail.CellValue(colDBAmt, k) = Format(sArCodeLvls(j), "#,#0.00")
            k = k + 1
        Next j
    Next i
    tgmDetail.Rebind
    
    If eTabMain.CurrTab = TabApprove Then
        eTabMain.TabEnabled(TabDetails) = True
        eTabMain.CurrTab = TabDetails
    End If
    
    frmContext.ButtonEnabled(FO_HOLD_UP) = True
    subEnablePrint TabDetails, True
    subSetFocus tblDetails
    Screen.MousePointer = vbDefault
    
End Sub

Private Function fnLoadBonusDetails(lEmpNbr As Long, Optional sSql As String) As Boolean
    Const SUB_NAME As String = "fnLoadBonusDetails"
    Dim strSQL As String
    Dim rsTemp As Recordset
    
    strSQL = "SELECT bm_bonus_code, bm_eligible_date, bc_type, bc_frequency,"
    strSQL = strSQL & " bc_code_desc, bm_sequence"
    If sSql = "" Then
        strSQL = strSQL & ", bf_level"
    End If
    strSQL = strSQL & " FROM bonus_master, bonus_codes"
    If sSql = "" Then
        strSQL = strSQL & ", bonus_formula"
    End If
    strSQL = strSQL & " WHERE bm_bonus_code = bc_bonus_code"
    If sSql = "" Then
        strSQL = strSQL & " AND bm_bonus_code = bf_bonus_code"
    End If
    strSQL = strSQL & " AND bm_empno = " & tfnRound(lEmpNbr)
    If txtPrftCtr <> "" Then
        strSQL = strSQL & " AND bm_eligible_pc = " & Trim(txtPrftCtr)
    End If
    If txtFrequency <> "" Then
        strSQL = strSQL & " AND bc_frequency = " & tfnSQLString(Trim(txtFrequency))
    End If
    If txtStartDate <> "" Then
        strSQL = strSQL & " AND bm_eligible_date <= " & tfnDateString(txtStartDate, True)
        strSQL = strSQL & " AND bm_stop_date >= " & tfnDateString(txtStartDate, True)
    Else
        strSQL = strSQL & " AND bm_eligible_date <= " & tfnDateString(Date, True)
        strSQL = strSQL & " AND bm_stop_date >= " & tfnDateString(Date, True)
    End If
    strSQL = strSQL & " ORDER BY bm_bonus_code, bm_sequence"
    If sSql = "" Then
        strSQL = strSQL & ", bf_level"
    End If
    
    If sSql = "" Then
        tgmDetail.FillWithSQL t_dbMainDatabase, strSQL
        If tgmDetail.RowCount <= 0 Then
            MsgBox "No record found for the selection criteria", vbExclamation
            Exit Function
        End If
    Else
        sSql = strSQL
    End If
    
    fnLoadBonusDetails = True
    
End Function

Private Sub subEnableEmployee(bOnOff As Boolean)
    txtEmployee.Enabled = bOnOff
    cmdEmployee.Enabled = bOnOff
    txtEmpName.Enabled = bOnOff
    cmdEmpName.Enabled = bOnOff
    subEnableSearchbtn cmdEmployee, bOnOff
    subEnableSearchbtn cmdEmpName, bOnOff
End Sub

Private Sub subShowFormulaDetails()
    Dim sCode As String
    Dim nLevel As Integer
    Dim nRow As Integer
    
    If tgmDetail.RowCount = 0 Then Exit Sub
    
    Screen.MousePointer = vbHourglass
    nRow = tgmDetail.GetCurrentRowNumber
    sCode = tgmDetail.CellValue(colDBCode, nRow)
    nLevel = tgmDetail.CellValue(colDBLevel, nRow)
    
    frmFORMULA.fnLoadBonusFormula sCode, nLevel
    Screen.MousePointer = vbDefault
    frmFORMULA.Show vbModal

End Sub

Private Sub txtStartDate_Change()
    cmdProcess.Enabled = False
    cValidate.Change txtStartDate
End Sub

Private Sub txtStartDate_GotFocus()
    cValidate.GotFocus txtStartDate
    SelectIt txtStartDate
End Sub

Private Sub txtStartDate_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        subSetFocus txtEndDate
        KeyAscii = 0
    Else
        tfnRegExpControlDateKeyPress txtStartDate, KeyAscii
        cValidate.Keypress txtStartDate, KeyAscii
    End If
End Sub

Private Sub txtStartDate_LostFocus()
    cValidate.LostFocus txtStartDate
    cmdProcess.Enabled = cValidate.FirstInvalidInput < 0
End Sub

Private Sub txtEndDate_Change()
    cmdProcess.Enabled = False
    cValidate.Change txtEndDate
End Sub

Private Sub txtEndDate_GotFocus()
    cValidate.GotFocus txtEndDate
    SelectIt txtEndDate
    If cValidate.ValidInput(txtEndDate) Then
        cValidate.GotFocus txtEndDate
    End If
End Sub

Private Sub txtEndDate_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        subSetFocus txtFrequency
        KeyAscii = 0
    Else
        tfnRegExpControlDateKeyPress txtEndDate, KeyAscii
        cValidate.Keypress txtEndDate, KeyAscii
    End If
End Sub

Private Sub txtEndDate_LostFocus()
    cValidate.LostFocus txtEndDate
    cmdProcess.Enabled = cValidate.FirstInvalidInput < 0
End Sub

Private Sub txtFrequency_Change()
    cValidate.Change txtFrequency
    tgcDropdown.Change txtFrequency
    cmdProcess.Enabled = False
End Sub

Private Sub txtFrequency_GotFocus()
    tgcDropdown.GotFocus txtFrequency
    cValidate.GotFocus txtFrequency
    
    If tgcDropdown.SingleRecordSelected Then
        subSetFocus txtPrftCtr
    End If
    
End Sub

Private Sub txtFrequency_KeyPress(KeyAscii As Integer)
    Dim bKeyCode As Boolean
    
    KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
    
    If KeyAscii = vbKeyReturn Then
        tgcDropdown.ComboSQL(txtFrequency) = fnSetComboSQL(txtFrequency.TabIndex)
        Screen.MousePointer = vbHourglass
    Else
        tfnRegExpControlKeyPress txtFrequency, KeyAscii, sFreqRegExp
    End If
        
    bKeyCode = tgcDropdown.Keypress(txtFrequency, KeyAscii)
    Screen.MousePointer = vbDefault
    
    If Not bKeyCode Then
        If KeyAscii = vbKeyReturn Then
            If tgcDropdown.SingleRecordSelected Then
                  subSetFocus txtPrftCtr
            End If
            KeyAscii = 0
            Screen.MousePointer = vbDefault
       End If
    Else
        cValidate.Keypress txtFrequency, KeyAscii
    End If

End Sub

Private Sub txtFrequency_LostFocus()
    tgcDropdown.LostFocus txtFrequency
    If cValidate.LostFocus(txtFrequency, cmdFrequency) Then
        cmdProcess.Enabled = cValidate.FirstInvalidInput < 0
    End If
End Sub

Private Sub cmdFrequency_Click()
    tgcDropdown.ComboSQL(txtFrequency) = fnSetComboSQL(txtFrequency.TabIndex)
    tgcDropdown.Click cmdFrequency
End Sub

Private Sub txtPrftCtr_Change()
    cValidate.Change txtPrftCtr
    tgcDropdown.Change txtPrftCtr
    cmdProcess.Enabled = False
    lstProcess.Clear
    txtPrftCtrName = ""
End Sub

Private Sub txtPrftCtr_GotFocus()
    tgcDropdown.GotFocus txtPrftCtr
    cValidate.GotFocus txtPrftCtr
    
    If tgcDropdown.SingleRecordSelected Then
        subSetFocus txtEmpProcess
    End If
    
End Sub

Private Sub txtPrftCtr_KeyPress(KeyAscii As Integer)
    Dim bKeyCode As Boolean
    
    If KeyAscii = vbKeyReturn Then
        tgcDropdown.ComboSQL(txtPrftCtr) = fnSetComboSQL(txtPrftCtr.TabIndex)
        Screen.MousePointer = vbHourglass
    End If
        
    bKeyCode = tgcDropdown.Keypress(txtPrftCtr, KeyAscii)
    Screen.MousePointer = vbDefault
    
    If Not bKeyCode Then
        If KeyAscii = vbKeyReturn Then
            If tgcDropdown.SingleRecordSelected Then
                  subSetFocus txtEmpProcess
            End If
            KeyAscii = 0
            Screen.MousePointer = vbDefault
       End If
    Else
        cValidate.Keypress txtPrftCtr, KeyAscii
    End If

End Sub

Private Sub txtPrftCtr_LostFocus()
    tgcDropdown.LostFocus txtPrftCtr
    If cValidate.LostFocus(txtPrftCtr, cmdPrftCtr, txtPrftCtrName, cmdPrftCtrName) Then
        cmdProcess.Enabled = cValidate.FirstInvalidInput < 0
    End If
End Sub

Private Sub txtPrftCtrName_Change()
    tgcDropdown.Change txtPrftCtrName
End Sub

Private Sub txtPrftCtrName_GotFocus()
    tfnSetStatusBarMessage "Enter Profit Center Name"
    tgcDropdown.GotFocus txtPrftCtrName
    Screen.MousePointer = vbDefault
    
    If tgcDropdown.SingleRecordSelected Then
        subSetFocus txtEmpProcess
        Screen.MousePointer = vbDefault
    End If
End Sub

Private Sub txtPrftCtrName_KeyPress(KeyAscii As Integer)
    Dim bKeyCode As Boolean
    
    If KeyAscii = vbKeyReturn Then
        tgcDropdown.ComboSQL(txtPrftCtrName) = fnSetComboSQL(txtPrftCtrName.TabIndex)
        Screen.MousePointer = vbHourglass
    End If
    
    bKeyCode = tgcDropdown.Keypress(txtPrftCtrName, KeyAscii)
    Screen.MousePointer = vbDefault
    
    If Not bKeyCode Then
        If KeyAscii = vbKeyReturn Then
            If tgcDropdown.SingleRecordSelected Then
                subSetFocus txtEmpProcess
            End If
            KeyAscii = 0
            Screen.MousePointer = vbDefault
        End If
    End If

End Sub

Private Sub txtPrftCtrName_LostFocus()
    tgcDropdown.LostFocus txtPrftCtrName
    If cValidate.LostFocus(txtPrftCtr, cmdPrftCtr, txtPrftCtrName, cmdPrftCtrName) Then
        cmdProcess.Enabled = cValidate.FirstInvalidInput < 0
        If ActiveControl Is cmdCancel(TabProcess) Then
            subSetFocus cmdProcess
        End If
    End If
End Sub

Private Sub cmdPrftCtr_Click()
    Screen.MousePointer = vbHourglass
    tgcDropdown.ComboSQL(txtPrftCtr) = fnSetComboSQL(txtPrftCtr.TabIndex)
    tgcDropdown.Click cmdPrftCtr
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdPrftCtrName_Click()
    tgcDropdown.ComboSQL(txtPrftCtrName) = fnSetComboSQL(txtPrftCtrName.TabIndex)
    tgcDropdown.Click cmdPrftCtrName
End Sub

Private Sub txtEmpProcess_Change()
    cmdProcess.Enabled = False
    cValidate.Change txtEmpProcess
    tgcDropdown.Change txtEmpProcess
End Sub

Private Sub txtEmpProcess_GotFocus()
    tgcDropdown.GotFocus txtEmpProcess
    cValidate.GotFocus txtEmpProcess
    
    If tgcDropdown.SingleRecordSelected Then
        cmdProcess.Enabled = cValidate.FirstInvalidInput < 0
        If cmdProcess.Enabled Then
            subSetFocus cmdProcess
        End If
        Screen.MousePointer = vbDefault
    End If
    
End Sub

Private Sub txtEmpProcess_KeyPress(KeyAscii As Integer)
    Dim bKeyCode As Boolean
    
    If KeyAscii = vbKeyReturn Then
        tgcDropdown.ComboSQL(txtEmpProcess) = fnSetComboSQL(txtEmpProcess.TabIndex)
        Screen.MousePointer = vbHourglass
    End If
        
    bKeyCode = tgcDropdown.Keypress(txtEmpProcess, KeyAscii)
    Screen.MousePointer = vbDefault
    
    If Not bKeyCode Then
       If KeyAscii = vbKeyReturn Then
          If tgcDropdown.SingleRecordSelected Then
                cmdProcess.Enabled = cValidate.FirstInvalidInput < 0
                If cmdProcess.Enabled Then
                    subSetFocus cmdProcess
                End If
                Screen.MousePointer = vbDefault
          End If
          KeyAscii = 0
       End If
    Else
        cValidate.Keypress txtEmpProcess, KeyAscii
    End If

End Sub

Private Sub txtEmpProcess_LostFocus()
    tgcDropdown.LostFocus txtEmpProcess
    If cValidate.LostFocus(txtEmpProcess, cmdEmpProcess, txtEmpNameProcess, cmdEmpName) Then
        cmdProcess.Enabled = cValidate.FirstInvalidInput < 0
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub txtEmpNameProcess_Change()
    tgcDropdown.Change txtEmpNameProcess
End Sub

Private Sub txtEmpNameProcess_GotFocus()
    tfnSetStatusBarMessage "Enter EmpProcess Name"
    tgcDropdown.GotFocus txtEmpNameProcess
    Screen.MousePointer = vbDefault
    
    If tgcDropdown.SingleRecordSelected Then
        cmdProcess.Enabled = cValidate.FirstInvalidInput < 0
        If cmdProcess.Enabled Then
            subSetFocus cmdProcess
        End If
        Screen.MousePointer = vbDefault
    End If
End Sub

Private Sub txtEmpNameProcess_KeyPress(KeyAscii As Integer)
    Dim bKeyCode As Boolean
    
    If KeyAscii = vbKeyReturn Then
        tgcDropdown.ComboSQL(txtEmpNameProcess) = fnSetComboSQL(txtEmpNameProcess.TabIndex)
        Screen.MousePointer = vbHourglass
    End If
    
    bKeyCode = tgcDropdown.Keypress(txtEmpNameProcess, KeyAscii)
    Screen.MousePointer = vbDefault
    
    If Not bKeyCode Then
        If KeyAscii = vbKeyReturn Then
            If tgcDropdown.SingleRecordSelected Then
                cmdProcess.Enabled = cValidate.FirstInvalidInput < 0
                If cmdProcess.Enabled Then
                    subSetFocus cmdProcess
                End If
                Screen.MousePointer = vbDefault
            End If
        KeyAscii = 0
        End If
    End If
End Sub

Private Sub txtEmpNameProcess_LostFocus()
    tgcDropdown.LostFocus txtEmpNameProcess
    If cValidate.LostFocus(txtEmpProcess, cmdEmpProcess, txtEmpNameProcess, cmdEmpName) Then
        cmdProcess.Enabled = cValidate.FirstInvalidInput < 0
        If lstProcess.ListCount = 0 Then
            If cmdProcess.Enabled Then
                subSetFocus cmdProcess
            Else
                subSetFocus cmdCancel(TabProcess)
            End If
        End If
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdEmpProcess_Click()
    Screen.MousePointer = vbHourglass
    tgcDropdown.ComboSQL(txtEmpProcess) = fnSetComboSQL(txtEmpProcess.TabIndex)
    tgcDropdown.Click cmdEmpProcess
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdEmpNameProcess_Click()
    tgcDropdown.ComboSQL(txtEmpNameProcess) = fnSetComboSQL(txtEmpNameProcess.TabIndex)
    tgcDropdown.Click cmdEmpNameProcess
End Sub

Private Function fnValidPrftCtr(Box As Textbox) As Boolean
    Const SUB_NAME As String = "fnValidPrftCtr"
    Dim strSQL As String
    Dim rsTemp As Recordset
    
    fnValidPrftCtr = False
    
    If Trim(Box.Text) = "" Then
        fnValidPrftCtr = True
        Exit Function
    End If
    
    If Not IsNumeric(Trim(Box.Text)) Then
        cValidate.SetErrorMessage Box, "Profit Center Number does not exist"
        Exit Function
    End If
    
    strSQL = "SELECT prft_ctr FROM sys_prft_ctr"
    strSQL = strSQL & " WHERE prft_ctr = " & Box.Text
    
    If GetRecordSet(rsTemp, strSQL, , SUB_NAME) < 0 Then
        cValidate.SetErrorMessage Box, "Failed to access Database"
        Exit Function
    End If
            
    If rsTemp.RecordCount = 0 Then
        cValidate.SetErrorMessage Box, "Profit Center Number does not exist"
        Exit Function
    End If
    
    fnValidPrftCtr = True
    
End Function

Private Sub subEnablePrint(Index As Integer, bYesNo As Boolean, _
                           Optional bAllBtns As Boolean = False)
    Dim i As Integer
    
    If bAllBtns Then
        For i = 1 To 3
            cmdPrint(i).Enabled = bYesNo
        Next i
    Else
        cmdPrint(Index).Enabled = bYesNo
    End If
    
    frmContext.ButtonEnabled(PRINT_UP) = bYesNo
    mnuPrint.Enabled = bYesNo
    
End Sub

Private Sub subPrint(Index As Integer)
    If Index = TabProcess Then
        Screen.MousePointer = vbHourglass
        tfnSetStatusBarMessage "Printing report..."
        subPrintProcess lstProcess
        Screen.MousePointer = vbDefault
        tfnSetStatusBarMessage " "
        Exit Sub
    End If
    
    If Not fnCreateReport(Index) Then
        Screen.MousePointer = vbDefault
        tfnSetStatusBarMessage "Failed to print the report"
        Exit Sub
    End If
    
    If Index = TabApprove Then
        subSetFocus tblApprove
    ElseIf Index = TabDetails Then
        subSetFocus tblDetails
    End If
    

End Sub

Private Sub subEnableFirstLineProcess(bYesNo As Boolean)
    txtStartDate.Enabled = bYesNo
    txtEndDate.Enabled = bYesNo
    txtFrequency.Enabled = bYesNo
    txtPrftCtr.Enabled = bYesNo
    txtPrftCtrName.Enabled = bYesNo
    txtEmpProcess.Enabled = bYesNo
    txtEmpNameProcess.Enabled = bYesNo
    subEnableSearchbtn cmdPrftCtr, bYesNo
    subEnableSearchbtn cmdPrftCtrName, bYesNo
    subEnableSearchbtn cmdFrequency, bYesNo
    subEnableSearchbtn cmdEmpProcess, bYesNo
    subEnableSearchbtn cmdEmpNameProcess, bYesNo
    
End Sub

Private Function fnGetEmployeeName(sEmpNo As String) As String
    Const SUB_NAME As String = ""
    Dim strSQL As String
    Dim rsTemp As Recordset
    
    fnGetEmployeeName = ""
    
    strSQL = "SELECT prm_empname FROM sTmpEmpTable WHERE"
    strSQL = strSQL & " prm_empno = " & tfnRound(sEmpNo)
    
    If GetRecordSet(rsTemp, strSQL, , SUB_NAME) < 0 Then
        subLogErrMsg "Failed to access database to get employee name"
        Exit Function
    End If
    
    If rsTemp.RecordCount > 0 Then
        fnGetEmployeeName = fnGetField(rsTemp!prm_empname)
    End If
    
End Function

Private Sub subSetAction(nCol As Integer)
    Dim i As Long
    Dim lCount As Long
    Dim lTemp() As Long
    
    tgsApprove.GetSelected lTemp, lCount
    
    For i = 0 To lCount - 1
        Screen.MousePointer = vbHourglass
        tgmApprove.CellValue(colAApprove, lTemp(i)) = nCol
    Next i
    
    Screen.MousePointer = vbDefault
    
End Sub

Private Function fnFillApproveTab(sArray As Variant) As Boolean
    Dim i As Integer
    
    fnFillApproveTab = False
    
    If UBound(vArrBonus) <= 0 Then
        Exit Function
    End If
    
    tgmApprove.FillWithArray vArrBonus
    For i = 0 To UBound(vArrBonus)
        ''
    Next i
    
    subEnablePrint TabApprove, True
    fnFillApproveTab = True

End Function

Public Function fnValidCellValue(Table As TDBGrid, ByVal nCol As Integer, _
                                 ByVal lRow As Long, sText As String) As Boolean
    #If PROTOTYPE Then
        fnValidCellValue = True
        Exit Function
    #End If
    
    Select Case Table.TabIndex
        Case tblSales.TabIndex
            Select Case nCol
                Case colSPrftCtr
                    fnValidCellValue = fnValidGridPrftCtr(sText, nCol, lRow)
                Case colSPrftName
                    fnValidCellValue = True
                Case colSAmount
                    If Trim(sText) = "" Then
                        tgmSales.ErrorMessage(nCol) = "You must enter the amount"
                        Exit Function
                    Else
                        fnValidCellValue = True
                    End If
                Case colSFromDate, colSToDate
                    fnValidCellValue = fnValidGridFromToDate(sText, nCol, lRow)
            End Select
        Case tblTimeCard.TabIndex
            fnValidCellValue = objHours.fnValidCellValue(Table, nCol, lRow, sText)
        Case tblProfitCenter.TabIndex
            fnValidCellValue = True
        Case tblApprove.TabIndex
            fnValidCellValue = True
    End Select
    
End Function

Private Sub cmdAddBtn_Click(Index As Integer)
    #If PROTOTYPE Then
        Exit Sub
    #End If
    
    If Index = TabSales Then
        t_nFormMode = ADD_MODE
        subEnableFirstLineSlsOrHrs Index, True
        
        frmContext.ButtonEnabled(CANCEL_UP) = True
        cmdCancel(TabSales).Enabled = True
        mnuCancel.Enabled = True
        
        cmdEditBtn(Index).Enabled = False
        cmdAddBtn(Index).Enabled = False
        cmdUpdateInsertBtn(Index).Caption = t_szCAPTION_INSERT
        
        tgcDropdown.ComboOn(txtFromDate) = False
        tgcDropdown.ComboOn(txtToDate) = False
        
        eTabSub.TabEnabled(TabHours) = False
        eTabMain.TabEnabled(TabProcess) = False
        
        subSetFocusoptType
    Else 'Index is Hours...
        eTabMain.TabEnabled(TabProcess) = False
        eTabSub.TabEnabled(TabSales) = False
        objHours.cmdAddBtn_Click
    End If
    
End Sub

Private Sub cmdDelete_Click(Index As Integer)
    Dim i As Long
    Dim lCount As Long
    Dim lTemp() As Long
    Dim sMsg As String
    Dim sPrftCtr As String
    Dim sAryPrftCtr() As String
    
    Select Case Index
        Case TabSales
            If tgmSales.RowCount < 1 Then
                Exit Sub
            End If
            
            If tgsSales.Count > 0 Then
                If tgsSales.Count = tgmSales.RowCount Then
                    If t_nFormMode = EDIT_MODE Then
                        sMsg = "Are you sure you want to delete all the records for the From Date and To Date?"
                    Else
                        sMsg = "Are you sure you want to delete all the rows from the Grid"
                    End If
                Else
                    If t_nFormMode = EDIT_MODE Then
                        sMsg = "Are you sure you want to delete the " & IIf(tgsSales.Count > 1, tgsSales.Count & " ", "") & "selected records for the From Date and To Date?"
                    Else
                        sMsg = "Are you sure you want to delete the " & IIf(tgsSales.Count > 1, tgsSales.Count & " ", "") & "selected rows from the Grid"
                    End If
                End If
                If Not tfnCancelExit(sMsg) Then
                    Exit Sub
                End If
                
                tgsSales.GetSelected lTemp, lCount
                
                If lCount > 0 Then
                    ReDim sAryPrftCtr(lCount - 1)
                    For i = 0 To lCount - 1
                        sAryPrftCtr(i) = fnCstr(tgmSales.CellValue(ColxSOldPrftCtr, _
                            lTemp(i)))
                    Next i
                    
                    For i = lCount - 1 To 0 Step -1
                        Screen.MousePointer = vbHourglass
                        If Not fnDeleteSales(fnGetSalesType(), sAryPrftCtr(i), txtToDate, txtFromDate) Then
                            tfnSetStatusBarError "Failed to delete the sales record"
                            Screen.MousePointer = vbDefault
                            Exit Sub
                        End If
                    Next i
                    
                    tgsSales.Delete
                    
                    If tgsSales.Count = tgmSales.RowCount Then
                        If t_nFormMode = EDIT_MODE Then
                            tfnResetScreen Index
                        Else
                            tgmSales.ClearData
                            tblSales.SetFocus
                        End If
                    Else
                        'tblSales.SetFocus
                    End If
                    Screen.MousePointer = vbDefault
                End If
                
                Exit Sub
            End If
            
            If t_nFormMode = EDIT_MODE Then
                If Not tfnCancelExit("Are you sure you want to delete the current record?") Then
                    Exit Sub
                End If
                sPrftCtr = fnCstr(tgmSales.CellValue(colSPrftCtr, tgmSales.GetCurrentRowNumber))
                If Not fnDeleteSales(fnGetSalesType(), sPrftCtr, txtToDate, txtFromDate) Then
                    tfnSetStatusBarError "Failed to delete the sales record"
                    Exit Sub
                End If
            End If
            
            tgmSales.DeleteRow
            
            If t_nFormMode = EDIT_MODE And tgmSales.RowCount = 0 Then
                tfnResetScreen Index
            End If
        Case nTabHours
            objHours.cmdDeleteBtn_Click
    End Select
End Sub

Private Sub cmdEditBtn_Click(Index As Integer)
    #If PROTOTYPE Then
        Exit Sub
    #End If
    
    If Index = TabSales Then
        t_nFormMode = EDIT_MODE
        subEnableFirstLineSlsOrHrs Index, True
        
        frmContext.ButtonEnabled(CANCEL_UP) = True
        cmdCancel(TabSales).Enabled = True
        mnuCancel.Enabled = True
        
        cmdEditBtn(Index).Enabled = False
        cmdAddBtn(Index).Enabled = False
        
        cmdUpdateInsertBtn(Index).Caption = t_szCAPTION_UPDATE
    
        tgcDropdown.ComboOn(txtFromDate) = True
        tgcDropdown.ComboOn(txtToDate) = True
        
        eTabSub.TabEnabled(TabHours) = False
        eTabMain.TabEnabled(TabProcess) = False
        
        subSetFocusoptType
    Else 'Index is Hours...
        eTabMain.TabEnabled(TabProcess) = False
        eTabSub.TabEnabled(TabSales) = False
        objHours.cmdEditBtn_Click
    End If
    
End Sub

Private Sub cmdRefresh_Click(Index As Integer)
    Dim sErrMsg As String
    
    Select Case Index
        Case TabSales
            cmdUpdateInsertBtn(Index).Enabled = False
            
            sErrMsg = fnLoadSales()
            If sErrMsg <> "" Then
                tfnSetStatusBarError sErrMsg
                Exit Sub
            End If
        Case nTabHours
            objHours.cmdRefreshSelectBtn_Click
    End Select
End Sub

Private Sub cmdUpdateInsertBtn_Click(Index As Integer)
    Select Case Index
        Case TabSales
            If Not fnCheckDuplicateInGrid() Then
                Exit Sub
            End If
            
            If bSalesRecordExist Then
                If MsgBox("Sales record(s) already exists for From Date/To Date and will be replaced." _
                   + " Are you sure you want to replace the existing sales record with the sales record on the Grid?", _
                   vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
                    Exit Sub
                End If
                
                If Not fnDeleteSalesRecord() Then
                    tfnSetStatusBarError "Failed to delete the existing record before insert"
                    Exit Sub
                End If
            End If
            
            If Not fnInsertUpdateSales() Then
                If t_nFormMode = ADD_MODE Then
                    tfnSetStatusBarError "Failed to insert sales record"
                Else
                    tfnSetStatusBarError "Failed to update sales record"
                End If
                Exit Sub
            End If
    
            nDataStatus = DATA_INIT
            tfnResetScreen Index
        Case nTabHours
            objHours.cmdUpdateInsertBtn_Click
    End Select
End Sub

Private Sub cmdAddBtn_GotFocus(Index As Integer)
    tfnSetStatusBarMessage ADD_EDIT_MSG
End Sub

Private Sub cmdEditBtn_GotFocus(Index As Integer)
    tfnSetStatusBarMessage ADD_EDIT_MSG
End Sub

Private Sub cmdDelete_GotFocus(Index As Integer)
    tfnSetStatusBarMessage "Delete"
End Sub

Private Sub cmdRefresh_GotFocus(Index As Integer)
    tfnSetStatusBarMessage t_szCAPTION_REFRESH
End Sub

Private Sub cmdUpdateInsertBtn_GotFocus(Index As Integer)
    If t_nFormMode = ADD_MODE Then
        tfnSetStatusBarMessage ("Insert")
    Else
        tfnSetStatusBarMessage ("Update")
    End If
End Sub

Public Function fnGetSalesType() As String
    Dim i As Integer
    Dim sType As String
    
    For i = 0 To 3
        If optType(i).Value Then
            Select Case i
                Case nWeek
                    sType = sWeek
                Case nOneMth
                    sType = sOneMth
                Case nThreeMth
                    sType = sThreeMth
                Case nGas
                    sType = sGas
            End Select
            Exit For
        End If
    Next i
    
    fnGetSalesType = sType
    
End Function

Private Sub subSetFloatingDropDown(Index As Integer)

    If Not tgfDropdown(Index) Is Nothing Then
        Set tgfDropdown(Index) = Nothing
    End If
    
    Set tgfDropdown(Index) = New clsFloatingDropDown
    With tgfDropdown(Index)
        Set .DataBase = t_dbMainDatabase
        Set .SearchButton = cmdDropdown(Index)
        Set .DropDownTable = tblDropDown(Index)
        Set .DataLink = datDropDown
        Set .Form = Me
            .SearchOnReturn = False
            .DefaultCursorControl = True
        
        Select Case Index
            Case TabSales
                Set .MainTable = tblSales
                Set .EditClass = tgmSales
                .AddDropDown 1
                .AddColumn colSPrftCtr, "prft_ctr", .COLUMN_TYPE_INTEGER
                .AddColumn colSPrftName, "prft_name", .COLUMN_TYPE_STRING
        End Select
    End With
    
End Sub

Private Sub subSetFloatingSQL(Index As Integer)
    Dim strSQL As String
    Dim sPrftCtrList As String
    
    Select Case Index
        Case TabSales
            strSQL = "SELECT prft_ctr, prft_name FROM sys_prft_ctr"
            strSQL = strSQL + " WHERE prft_type IN ('R', 'B')"
            sPrftCtrList = fnBuildPrftCtrList()
            If sPrftCtrList <> "" Then
                strSQL = strSQL + " AND prft_ctr NOT IN (" + sPrftCtrList + ")"
            End If
            tgfDropdown(Index).SetSQL colSPrftCtr, strSQL
    End Select
    
End Sub

Private Sub tblDropDown_Click(Index As Integer)
    If Index = nTabHours Then
        objHours.tblFloating_Click
    Else
        tgfDropdown(Index).TableClick tblDropDown(Index)
    End If
End Sub

Private Sub tblDropDown_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index = nTabHours Then
        objHours.tblFloating_KeyPress KeyAscii
    Else
        tgfDropdown(Index).Keypress tblDropDown(Index), KeyAscii
    End If
End Sub

Private Sub tblDropDown_LostFocus(Index As Integer)
    If Index = nTabHours Then
        objHours.tblFloating_LostFocus
    Else
        tgfDropdown(Index).LostFocus tblDropDown(Index)
    End If
End Sub

Private Sub tblDropDown_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Index = nTabHours Then
        objHours.tblFloating_MouseUp Button, Shift, X, Y
    Else
        tgfDropdown(Index).MouseUp Y
    End If
End Sub

Private Sub tblDropDown_RowColChange(Index As Integer, LastRow As Variant, ByVal LastCol As Integer)
    If Index = nTabHours Then
        objHours.tblFloating_RowColChange LastRow, LastCol
    Else
        tgfDropdown(Index).RowColChange tblDropDown(Index)
    End If
End Sub

Private Sub subEnterPhaseIISlsOrHrs(Index As Integer)
    Dim sErrMsg As String
    
    If Index = TabSales Then
        If cValidSls.FirstInvalidInput >= 0 Then
            'subSetFocus cmdCancel(Index)
            Exit Sub
        End If
    End If

    subSetFocus efraBackground
    
    Screen.MousePointer = vbHourglass
    
    If Not txtFromDate.Enabled Then
        Exit Sub
    End If
    
    subEnableFirstLineSlsOrHrs Index, False
    
    If Index = TabSales Then
        sErrMsg = fnLoadSales()
        If sErrMsg <> "" Then
            subSetFocus cmdCancel(TabSales)
            DoEvents
            tfnSetStatusBarError sErrMsg
            'cmdCancel_Click (TabSales)
            subEnableFirstLineSlsOrHrs Index, True
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
    End If
        
    If Index = TabSales Then
        tblSales.Enabled = True
        DoEvents
        subSetFocus tblSales
    End If
    
    Screen.MousePointer = vbDefault
End Sub

Private Sub subEnableFirstLineSlsOrHrs(Index As Integer, bYesNo As Boolean)
    Select Case Index
        Case TabSales
            efraOptSales.Enabled = bYesNo
            txtFromDate.Enabled = bYesNo
            txtToDate.Enabled = bYesNo
            If t_nFormMode = ADD_MODE Then
                bYesNo = False
            End If
            subEnableSearchbtn cmdFromDate, bYesNo
            subEnableSearchbtn cmdToDate, bYesNo
    End Select
End Sub

Private Sub SubEnableDeleteBtn(bOnOff As Boolean, Index As Integer)
    cmdDelete(Index).Enabled = bOnOff
End Sub

Private Sub subEnableRefreshBtn(bOnOff As Boolean, Index As Integer)
    cmdRefresh(Index).Enabled = bOnOff
End Sub

Private Sub subEnableUpdateBtn(bOnOff As Boolean, Index As Integer)
    cmdUpdateInsertBtn(Index).Enabled = bOnOff
End Sub

Private Function fnValidGridPrftCtr(sText As String, nCol As Integer, lRow As Long) As Boolean
    Const SUB_NAME As String = "fnValidGridPrftCtr"
    Dim strSQL As String
    Dim rsTemp As Recordset
    
    fnValidGridPrftCtr = False
            
    strSQL = "SELECT prft_ctr, prft_name, prft_type FROM sys_prft_ctr"
    strSQL = strSQL + " WHERE prft_ctr = " & tfnRound(sText)
    
    If GetRecordSet(rsTemp, strSQL, , SUB_NAME) < 0 Then
        tgmSales.ErrorMessage(nCol) = "Failed to access Database"
        Exit Function
    End If
    
    If rsTemp.RecordCount = 0 Then
        tgmSales.ErrorMessage(nCol) = "Profit Center does not exist"
        Exit Function
    End If
    
    If Not (fnGetField(rsTemp!prft_type) = "R" Or fnGetField(rsTemp!prft_type) = "B") Then
        tgmSales.ErrorMessage(nCol) = "Profit Center type must be 'R' or 'B'"
        Exit Function
    End If
    
    tgmSales.CellValue(colSPrftName, lRow) = fnGetField(rsTemp!prft_name)
    
    If fnCstr(tgmSales.CellValue(colSFromDate, lRow)) = "" Then
        tgmSales.CellValue(colSFromDate, lRow) = txtFromDate
        tgmSales.CellValue(colSToDate, lRow) = txtToDate
    End If
    
Debug.Print tblSales.col
    Select Case tblSales.col
    Case colSPrftCtr
        tblSales.col = colSPrftName
    Case colSPrftName
        tblSales.col = colSAmount
    End Select
    
    fnValidGridPrftCtr = True
    
End Function

Private Function fnValidSalesDate(Box As Textbox) As Boolean
    Const SUB_NAME As String = "fnValidSalesDate"
    Dim strSQL As String
    Dim sTemp As String
    Dim sErrMsg As String
    
    sTemp = "From Date"
    If Box Is txtToDate Then
        sTemp = "To Date"
    End If
    
    fnValidSalesDate = False
    
    If Trim(Box) = "" Then
        'fnValidSalesDate = True
        cValidSls.SetErrorMessage Box, "You must enter a " & sTemp
        Exit Function
    End If
    
    If Not IsValidDate(Box) Then
        cValidSls.SetErrorMessage Box, sTemp & " is not valid"
        Exit Function
    End If
    
    Box = tfnFormatDate(Box)
    
    If Box Is txtFromDate Then
        If Not IsValidDate(txtToDate) Then
            fnValidSalesDate = True
            Exit Function
        End If
    Else
        If Not IsValidDate(txtFromDate) Then
            fnValidSalesDate = True
            Exit Function
        End If
    End If
    
    If CDate(tfnDateString(txtFromDate)) > CDate(tfnDateString(txtToDate)) Then
        cValidSls.SetErrorMessage txtFromDate, "From Date must be earlier than To Date"
        cValidSls.SetErrorMessage txtToDate, "To Date must be later than From Date"
        
        If Box Is txtFromDate Then
            cValidSls.ValidInput(txtToDate) = False
        Else
            cValidSls.ValidInput(txtFromDate) = False
        End If
    
        Exit Function
    End If
    
    If t_nFormMode = ADD_MODE Then
        sErrMsg = fnCheckSales()
        If sErrMsg <> "" Then
            cValidSls.SetErrorMessage txtFromDate, sErrMsg
            cValidSls.SetErrorMessage txtToDate, sErrMsg
            Exit Function
        End If
    End If
    
    If Box Is txtFromDate Then
        cValidSls.ValidInput(txtToDate) = True
        If ActiveControl Is txtToDate Then
            tfnSetStatusBarMessage "Enter To Date"
        End If
    Else
        cValidSls.ValidInput(txtFromDate) = True
        If ActiveControl Is txtFromDate Then
            tfnSetStatusBarMessage "Enter From Date"
        End If
    End If
    
    fnValidSalesDate = True
    
End Function

Private Function fnValidProcessDate(Box As Textbox) As Boolean
    Dim strSQL As String
    Dim sTemp As String
    Dim sErrMsg As String
    
    sTemp = "Starting Date"
    If Box Is txtEndDate Then
        sTemp = "Ending Date"
    End If
    
    fnValidProcessDate = False
    
    If Trim(Box) = "" Then
        'fnValidProcessDate = True
        cValidate.SetErrorMessage Box, "You must enter a" + _
            IIf(Left(sTemp, 1) = "E", "n", "") + " " & sTemp
        Exit Function
    End If
    
    If Not IsValidDate(Box) Then
        cValidate.SetErrorMessage Box, sTemp & " is not valid"
        Exit Function
    End If
    
    Box = tfnFormatDate(Box)
    
    If Box Is txtStartDate Then
        If Not IsValidDate(txtEndDate) Then
            fnValidProcessDate = True
            Exit Function
        End If
    Else
        If Not IsValidDate(txtStartDate) Then
            fnValidProcessDate = True
            Exit Function
        End If
    End If
    
    If CDate(tfnDateString(txtStartDate)) > CDate(tfnDateString(txtEndDate)) Then
        cValidate.SetErrorMessage txtStartDate, "Starting Date must be earlier than Ending Date"
        cValidate.SetErrorMessage txtEndDate, "Ending Date must be later than Starting Date"
        
        If Box Is txtStartDate Then
            cValidate.ValidInput(txtEndDate) = False
        Else
            cValidate.ValidInput(txtStartDate) = False
        End If
    
        Exit Function
    End If
    
    sErrMsg = fnCheckFrequency(txtStartDate, txtEndDate, txtFrequency)
    If sErrMsg <> "" Then
        cValidate.SetErrorMessage txtStartDate, sErrMsg
        cValidate.SetErrorMessage txtEndDate, sErrMsg
        Exit Function
    End If
    
    If Box Is txtStartDate Then
        cValidate.ValidInput(txtEndDate) = True
        If ActiveControl Is txtEndDate Then
            tfnSetStatusBarMessage "Enter Ending Date"
        End If
    Else
        cValidate.ValidInput(txtStartDate) = True
        If ActiveControl Is txtStartDate Then
            tfnSetStatusBarMessage "Enter Starting Date"
        End If
    End If
    
    fnValidProcessDate = True
    
End Function

Private Sub txtFromDate_Change()
    cValidSls.Change txtFromDate
    tgcDropdown.Change txtFromDate
End Sub

Private Sub txtFromDate_GotFocus()
    cValidSls.GotFocus txtFromDate
    tgcDropdown.GotFocus txtFromDate
    If tgcDropdown.SingleRecordSelected Then
        subSetFocus txtToDate
        Screen.MousePointer = vbDefault
    End If
End Sub

Private Sub txtFromDate_KeyPress(KeyAscii As Integer)
    Dim bKeyCode As Boolean
    
    If KeyAscii = vbKeyReturn Then
        If tgcDropdown.ComboOn(txtFromDate) Then
            tgcDropdown.ComboSQL(txtFromDate) = fnSetComboSQL(txtFromDate.TabIndex)
            Screen.MousePointer = vbHourglass
        End If
    End If
        
    bKeyCode = tgcDropdown.Keypress(txtFromDate, KeyAscii)
    Screen.MousePointer = vbDefault
    
    If Not bKeyCode Then
        If KeyAscii = vbKeyReturn Then
            If tgcDropdown.ComboOn(txtFromDate) Then
                If tgcDropdown.SingleRecordSelected Then
                      subSetFocus txtToDate
                      Screen.MousePointer = vbDefault
                End If
            Else
'                If cValidSls.FirstInvalidInput < 0 Then
'                    subEnterPhaseIISlsOrHrs TabSales
'                Else
                    subSetFocus txtToDate
'                End If
            End If
          KeyAscii = 0
       End If
    Else
        cValidSls.Keypress txtFromDate, KeyAscii
    End If
End Sub

Private Sub txtFromDate_LostFocus()
    txtFromDate = tfnFormatDate(txtFromDate)
    
    If cValidSls.LostFocus(txtFromDate, cmdFromDate, tblComboDropdown) Then
'        subEnterPhaseIISlsOrHrs TabSales
    End If
    tgcDropdown.LostFocus txtFromDate
End Sub

Private Sub cmdFromDate_Click()
    Screen.MousePointer = vbHourglass
    tgcDropdown.ComboSQL(txtFromDate) = fnSetComboSQL(txtFromDate.TabIndex)
    tgcDropdown.Click cmdFromDate
    Screen.MousePointer = vbDefault
End Sub

Private Sub txtToDate_Change()
    cValidSls.Change txtToDate
    tgcDropdown.Change txtToDate
End Sub

Private Sub txtToDate_GotFocus()
    cValidSls.GotFocus txtToDate
    tgcDropdown.GotFocus txtToDate
    If tgcDropdown.SingleRecordSelected Then
        subEnterPhaseIISlsOrHrs TabSales
        Screen.MousePointer = vbDefault
    End If
End Sub

Private Sub txtToDate_KeyPress(KeyAscii As Integer)
    Dim bKeyCode As Boolean
    
    If KeyAscii = vbKeyReturn Then
        If tgcDropdown.ComboOn(txtToDate) Then
            tgcDropdown.ComboSQL(txtToDate) = fnSetComboSQL(txtToDate.TabIndex)
            Screen.MousePointer = vbHourglass
        End If
    End If
        
    bKeyCode = tgcDropdown.Keypress(txtToDate, KeyAscii)
    Screen.MousePointer = vbDefault
    
    If Not bKeyCode Then
       If KeyAscii = vbKeyReturn Then
            If tgcDropdown.ComboOn(txtToDate) Then
                If tgcDropdown.SingleRecordSelected Then
                      subEnterPhaseIISlsOrHrs TabSales
                      Screen.MousePointer = vbDefault
                End If
            Else
                If cValidSls.FirstInvalidInput < 0 Then
                    subEnterPhaseIISlsOrHrs TabSales
                Else
                    SendKeys "{TAB}", True
                End If
            End If
          KeyAscii = 0
       End If
    Else
        cValidSls.Keypress txtToDate, KeyAscii
    End If
End Sub

Private Sub txtToDate_LostFocus()
    txtToDate = tfnFormatDate(txtToDate)
    
    If cValidSls.LostFocus(txtToDate, cmdToDate, tblComboDropdown) Then
        If Not (ActiveControl Is cmdCancel(TabSales) Or ActiveControl Is cmdExitCancelBtn) Then
            subEnterPhaseIISlsOrHrs TabSales
        End If
    End If
    tgcDropdown.LostFocus txtToDate
End Sub

Private Sub cmdToDate_Click()
    Screen.MousePointer = vbHourglass
    tgcDropdown.ComboSQL(txtToDate) = fnSetComboSQL(txtToDate.TabIndex)
    tgcDropdown.Click cmdToDate
    Screen.MousePointer = vbDefault
End Sub

Private Sub tblSales_AfterColEdit(ByVal ColIndex As Integer)
    Dim lRow As Long
    
    lRow = tgmSales.GetCurrentRowNumber
    tgmSales.AfterColEdit ColIndex
    
    If t_nFormMode = EDIT_MODE Then
        If nDataStatus = DATA_CHANGED Then
            subEnableRefreshBtn True, TabSales
        Else
            subEnableRefreshBtn False, TabSales
        End If
    End If

End Sub

Private Sub tblSales_BeforeColEdit(ByVal ColIndex As Integer, ByVal KeyAscii As Integer, CANCEL As Integer)
    If ColIndex = colSPrftCtr Then
        tgmSales.CellValue(colSPrftName, tgmSales.GetCurrentRowNumber) = ""
    End If
    tgmSales.BeforeColEdit ColIndex, KeyAscii, CANCEL
End Sub

Private Sub tblSales_Change()
    tgmSales.Change
    subEnableUpdateBtn False, TabSales
    nDataStatus = DATA_CHANGED
End Sub

Private Sub tblSales_FirstRowChange()
    tgmSales.FirstRowChange
    tgfDropdown(TabSales).FirstRowChange
    
    If tblSales.Row = -1 And tgsSales.Count = 0 Then
        SubEnableDeleteBtn False, TabSales
    End If
    
End Sub

Private Sub tblSales_GotFocus()
    tfnSetStatusBarMessage "Store Sales"
    tgsSales.GotFocus
    tgmSales.GotFocus
    tgfDropdown(TabSales).GotFocus
End Sub

Private Sub tblSales_KeyDown(KeyCode As Integer, Shift As Integer)
    If tblSales.SelBookmarks.Count > 0 Then
        If KeyCode = vbKeyDelete And Shift = 0 Then
            KeyCode = 0
            cmdDelete_Click TabSales
            Exit Sub
        End If
    End If
    tgsSales.KeyDown KeyCode, Shift
    tgmSales.KeyDown KeyCode, Shift
    tgfDropdown(TabSales).KeyDown tblSales, KeyCode
End Sub

Private Sub tblSales_KeyPress(KeyAscii As Integer)
    Dim lRow As Long
    
    tgfDropdown(TabSales).Keypress tblSales, KeyAscii
    lRow = tgmSales.GetCurrentRowNumber
    
    If Not tgmSales.Keypress(KeyAscii) Then
        KeyAscii = 0
    End If
End Sub

Private Sub tblSales_LeftColChange()
    tgfDropdown(TabSales).LeftColChange
End Sub

Private Sub tblSales_LostFocus()
    tgmSales.LostFocus
    tgfDropdown(TabSales).LostFocus tblSales
    subSetStdBtn TabSales, tgmSales
End Sub

Private Sub tblSales_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    tgsSales.MouseUp Button, Shift, Y
End Sub

Private Sub tblSales_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    Dim lRow As Long
    
    If tgfDropdown(TabSales).RowColChange(tblSales) Then
        tgmSales.RowColChange LastRow, LastCol
    End If
    
    lRow = tgmSales.GetCurrentRowNumber
    
    If t_nFormMode = ADD_MODE Then
        If tgmSales.CellValue(colSPrftCtr, lRow) <> "" Then
            SubEnableDeleteBtn True, TabSales
        Else
            SubEnableDeleteBtn False, TabSales
        End If
    Else
        If tgmSales.RowCount > 0 Then
            SubEnableDeleteBtn True, TabSales
        Else
            SubEnableDeleteBtn False, TabSales
        End If
    End If
    
    tgsSales.RowColChange LastRow, LastCol
    
    subSetStdBtn TabSales, tgmSales

End Sub

Private Sub tblSales_SelChange(CANCEL As Integer)
    tgsSales.SelChange CANCEL
    CANCEL = True
End Sub

Private Sub tblSales_UnboundReadData(ByVal RowBuf As DBTrueGrid.RowBuffer, StartLocation As Variant, ByVal ReadPriorRows As Boolean)
    tgmSales.ReadData RowBuf, StartLocation, ReadPriorRows
End Sub

Private Sub subSetStdBtn(Index As Integer, tgmEditor As clsTGSpreadSheet)
    
    If tgmEditor.RowCount < 1 Then
        subEnableUpdateBtn False, Index
        SubEnableDeleteBtn False, Index
        Exit Sub
    End If
    
    If tgmEditor.ValidData Then
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
    
End Sub

Private Function fnValidGridFromToDate(sText As String, nCol As Integer, lRow As Long) As Boolean
    Const SUB_NAME As String = "fnValidGridFromToDate"
    Dim strSQL As String
    Dim rsTemp As Recordset
    Dim sTemp As String

    fnValidGridFromToDate = True
    Exit Function

    If nCol = colSFromDate Then
        sTemp = "From Date"
    Else
        sTemp = "To Date"
    End If

    If Trim(sText) = "" Then
        tgmSales.ErrorMessage(nCol) = "You must enter a " & sTemp
        Exit Function
    End If

    If Len(Trim(sText)) < 6 Then
        tgmSales.ErrorMessage(nCol) = "Invalid date format"
        Exit Function
    End If

    If nCol = colSFromDate Then
        tgmSales.CellValue(colSFromDate, lRow) = CDate(tfnFormatDate(sText))
    Else
        tgmSales.CellValue(colSToDate, lRow) = CDate(tfnFormatDate(sText))
    End If

    If Not IsDate(Trim(sText)) Then
        tgmSales.ErrorMessage(nCol) = "Invalid date format"
        Exit Function
    End If

    fnValidGridFromToDate = True

End Function

Private Function fnLoadSales() As String
    Const SUB_NAME As String = "fnLoadSales"
    Dim strSQL As String
    Dim rsTemp As Recordset
    Dim i As Long
    
    strSQL = fnGetSalesSQL()
    
    If GetRecordSet(rsTemp, strSQL, , SUB_NAME) < 0 Then
        fnLoadSales = "Failed to access database to load the sales record"
        Exit Function
    End If
    
    If rsTemp.RecordCount = 0 Then
        If t_nFormMode = ADD_MODE Then
            If MsgBox("Sales record not available for the From Date and To Date. Do you want to continue?", vbQuestion + vbYesNo) = vbNo Then
                fnLoadSales = "No Sales record available to Add"
            End If
        Else
            fnLoadSales = "No Sales record available to Edit"
        End If
        Exit Function
    End If
    
    tgmSales.FillWithRecordset rsTemp
    
    For i = 0 To tgmSales.RowCount - 1
        'fill the from/to date in the grid
        If t_nFormMode = ADD_MODE Then
            tgmSales.CellValue(colSFromDate, i) = txtFromDate
            tgmSales.CellValue(colSToDate, i) = txtToDate
        Else
            tgmSales.CellValue(colSFromDate, i) = tfnFormatDate(tgmSales.CellValue(colSFromDate, i))
            tgmSales.CellValue(colSToDate, i) = tfnFormatDate(tgmSales.CellValue(colSToDate, i))
            
            'store the profit center
            tgmSales.CellValue(ColxSOldPrftCtr, i) = tgmSales.CellValue(colSPrftCtr, i)
        End If
    Next i

    tgmSales.Rebind
    DoEvents
    
    fnLoadSales = ""

End Function

Private Function fnGetSalesSQL(Optional txtBox As Textbox = Nothing) As String
    Dim strSQL As String
    Dim sSalesType As String
    
    If t_nFormMode = EDIT_MODE Then
        sSalesType = "EDIT_MODE"
    Else
        sSalesType = fnGetSalesType
    End If
    
    Select Case sSalesType
        Case sWeek, sOneMth, sThreeMth
            strSQL = "SELECT rssl_prft_ctr AS prft_ctr, prft_name,"
            strSQL = strSQL & " SUM(rsc_retail) as amount"
            strSQL = strSQL & " FROM rs_shiftlink, sys_prft_ctr, rs_scat, rs_cat"
            strSQL = strSQL & " WHERE rssl_shl = rsc_shl"
            strSQL = strSQL & " AND rssl_prft_ctr = prft_ctr"
            strSQL = strSQL & " AND rsc_catagory = rsct_catagory"
            strSQL = strSQL & " AND rsct_type IN ('M', 'N', 'D')"
            strSQL = strSQL & " AND rssl_date BETWEEN " & tfnDateString(txtFromDate, True)
            strSQL = strSQL & " AND " & tfnDateString(txtToDate, True)
            strSQL = strSQL & " GROUP BY rssl_prft_ctr, prft_name"
            strSQL = strSQL & " ORDER BY rssl_prft_ctr"
        Case sGas
            strSQL = "SELECT rsd_prft_ctr AS prft_ctr, prft_name,"
            strSQL = strSQL & " SUM(rsd_gal) as amount"
            strSQL = strSQL & " FROM rs_daily, sys_prft_ctr"
            strSQL = strSQL & " WHERE prft_ctr = rsd_prft_ctr"
            strSQL = strSQL & " AND rsd_date BETWEEN " & tfnDateString(txtFromDate, True)
            strSQL = strSQL & " AND " & tfnDateString(txtToDate, True)
            strSQL = strSQL & " GROUP BY rsd_prft_ctr, prft_name"
            strSQL = strSQL & " ORDER BY rsd_prft_ctr"
        Case "EDIT_MODE"
            If txtBox Is Nothing Then
                strSQL = "SELECT prft_ctr, prft_name, bs_sales_amount as amount,"
                strSQL = strSQL & " bs_to_date as to_date, bs_from_date as from_date"
                strSQL = strSQL & " FROM bonus_sales, sys_prft_ctr "
                strSQL = strSQL & " WHERE bs_prft_ctr = prft_ctr"
                strSQL = strSQL & " AND bs_sales_type = " & tfnSQLString(fnGetSalesType())
                strSQL = strSQL & " AND bs_from_date = " & tfnDateString(txtFromDate, True)
                strSQL = strSQL & " AND bs_to_date = " & tfnDateString(txtToDate, True)
                strSQL = strSQL & " ORDER BY prft_ctr"
            Else
                If txtBox Is txtFromDate Then
                    strSQL = "SELECT bs_from_date"
                    strSQL = strSQL & " FROM bonus_sales"
                    strSQL = strSQL & " WHERE bs_sales_type = " & tfnSQLString(fnGetSalesType())
                    If IsValidDate(txtToDate) Then
                        strSQL = strSQL & " AND bs_to_date = " & tfnDateString(txtToDate, True)
                    End If
                    strSQL = strSQL & " GROUP BY bs_from_date"
                Else
                    strSQL = "SELECT bs_to_date"
                    strSQL = strSQL & " FROM bonus_sales"
                    strSQL = strSQL & " WHERE bs_sales_type = " & tfnSQLString(fnGetSalesType())
                    If IsValidDate(txtFromDate) Then
                        strSQL = strSQL & " AND bs_from_date = " & tfnDateString(txtFromDate, True)
                    End If
                    strSQL = strSQL & " GROUP BY bs_to_date"
                End If
            End If
    End Select
    fnGetSalesSQL = strSQL
    
End Function

Private Function IsValidDate(ByVal sDate As String) As Boolean
    If sDate = "" Then
        Exit Function
    End If
    
    sDate = tfnFormatDate(sDate)
    
    If SRegExpMatch(szDatePattern, sDate) <> 0 Then
        Exit Function
    End If
    
    If Not IsDate(tfnDateString(sDate)) Then
        Exit Function
    End If
    
    IsValidDate = True
End Function

Private Sub subSetFocusoptType()
    Dim i As Integer
    
    For i = 0 To 3
        If optType(i).Value Then
            subSetFocus optType(i)
            Exit Sub
        End If
    Next i
    
    subSetFocus optType(0)
End Sub

Private Function fnCheckSales() As String
    Const SUB_NAME As String = "fnCheckSales"
    Dim strSQL As String
    Dim rsTemp As Recordset
    
    bSalesRecordExist = False
    
    'check from date
    strSQL = "SELECT COUNT(bs_from_date) AS cnt_date"
    strSQL = strSQL & " FROM bonus_sales"
    strSQL = strSQL & " WHERE bs_sales_type = " & tfnSQLString(fnGetSalesType())
    strSQL = strSQL & " AND " & tfnDateString(txtFromDate, True)
    strSQL = strSQL & " BETWEEN bs_from_date AND bs_to_date"
    
    If GetRecordSet(rsTemp, strSQL, , SUB_NAME) <= 0 Then
        fnCheckSales = "Failed to access database"
        Exit Function
    End If
    
    If tfnRound(rsTemp!cnt_date) > 0 Then
        bSalesRecordExist = True
    End If
    
    If Not bSalesRecordExist Then
        strSQL = "SELECT COUNT(bs_from_date) AS cnt_date"
        strSQL = strSQL & " FROM bonus_sales"
        strSQL = strSQL & " WHERE bs_sales_type = " & tfnSQLString(fnGetSalesType())
        strSQL = strSQL & " AND bs_from_date BETWEEN " & tfnDateString(txtFromDate, True)
        strSQL = strSQL & " AND " & tfnDateString(txtToDate, True)
        
        If GetRecordSet(rsTemp, strSQL, , SUB_NAME) <= 0 Then
            fnCheckSales = "Failed to access database"
            Exit Function
        End If
        
        If tfnRound(rsTemp!cnt_date) > 0 Then
            bSalesRecordExist = True
        End If
    End If

    'check to date
    If Not bSalesRecordExist Then
        strSQL = "SELECT COUNT(bs_to_date) AS cnt_date"
        strSQL = strSQL & " FROM bonus_sales"
        strSQL = strSQL & " WHERE bs_sales_type = " & tfnSQLString(fnGetSalesType())
        strSQL = strSQL & " AND " & tfnDateString(txtToDate, True)
        strSQL = strSQL & " BETWEEN bs_from_date AND bs_to_date"
        
        If GetRecordSet(rsTemp, strSQL, , SUB_NAME) <= 0 Then
            fnCheckSales = "Failed to access database"
            Exit Function
        End If
        
        If tfnRound(rsTemp!cnt_date) > 0 Then
            bSalesRecordExist = True
        End If
    End If
    
    If Not bSalesRecordExist Then
        strSQL = "SELECT COUNT(bs_to_date) AS cnt_date"
        strSQL = strSQL & " FROM bonus_sales"
        strSQL = strSQL & " WHERE bs_sales_type = " & tfnSQLString(fnGetSalesType())
        strSQL = strSQL & " AND bs_to_date BETWEEN " & tfnDateString(txtFromDate, True)
        strSQL = strSQL & " AND " & tfnDateString(txtToDate, True)
        
        If GetRecordSet(rsTemp, strSQL, , SUB_NAME) <= 0 Then
            fnCheckSales = "Failed to access database"
            Exit Function
        End If
        
        If tfnRound(rsTemp!cnt_date) > 0 Then
            bSalesRecordExist = True
        End If
    End If
    
    If bSalesRecordExist Then
        If MsgBox("Sales record(s) already exists for From Date/To Date and will be replaced." _
           + " Are you sure you want to replaced the existing sales record?", _
           vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
            bSalesRecordExist = False
            fnCheckSales = "Sales record(s) already exists for From Date/To Date"
            Exit Function
        End If
    End If
    
    fnCheckSales = ""
End Function

Private Function fnCheckDuplicateInGrid() As Boolean
    Dim lRow As Long
    Dim i As Long
    
    If tgmSales.RowCount = 1 Then
        fnCheckDuplicateInGrid = True
        Exit Function
    End If
    
    For lRow = 0 To tgmSales.RowCount - 1
        For i = lRow + 1 To tgmSales.RowCount - 1
            If tfnRound(tgmSales.CellValue(colSPrftCtr, lRow)) = _
               tfnRound(tgmSales.CellValue(colSPrftCtr, i)) Then
                On Error Resume Next
                tblSales.Bookmark = tgmSales.Bookmark(i)
                subSetFocus tblSales
                DoEvents
                tfnSetStatusBarError "Duplicate Profit Center encountered"
                Exit Function
            End If
        Next i
    Next lRow

    fnCheckDuplicateInGrid = True
End Function

Private Function fnCstr(v) As String
    If Not IsNull(v) Then
        fnCstr = Trim(v)
    End If
End Function

Private Function fnBuildPrftCtrList() As String
    Dim sTemp As String
    Dim i As Long
    
    If tgmSales.RowCount <= 1 Then
        Exit Function
    End If
    
    For i = 0 To tgmSales.RowCount - 1
        If i <> tgmSales.GetCurrentRowNumber Then
            If tgmSales.ValidCell(colSPrftCtr, i) Then
                sTemp = sTemp & tfnRound(tgmSales.CellValue(colSPrftCtr, i)) & ","
            End If
        End If
    Next i
    
    If sTemp <> "" Then
        fnBuildPrftCtrList = Left(sTemp, Len(sTemp) - 1)
    End If
End Function


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Sub txtEmployeeNumber_Change()
    objHours.txtEmployeeNumber_Change
End Sub

Private Sub txtEmployeeNumber_GotFocus()
    objHours.txtEmployeeNumber_GotFocus
End Sub

Private Sub txtEmployeeNumber_KeyPress(KeyAscii As Integer)
    objHours.txtEmployeeNumber_KeyPress KeyAscii
End Sub

Private Sub txtEmployeeNumber_LostFocus()
    objHours.txtEmployeeNumber_LostFocus
End Sub

Private Sub txtEmployeeName_Change()
    objHours.txtEmployeeName_Change
End Sub

Private Sub txtEmployeeName_GotFocus()
    objHours.txtEmployeeName_GotFocus
End Sub

Private Sub txtEmployeeName_KeyPress(KeyAscii As Integer)
    objHours.txtEmployeeName_KeyPress KeyAscii
End Sub

Private Sub cmdEmployeeNumber_Click()
    objHours.cmdEmployeeNumber_Click
End Sub

Private Sub cmdEmployeeName_Click()
    objHours.cmdEmployeeName_Click
End Sub

Private Sub txtSSN_Change()
    objHours.txtSSN_Change
End Sub

Private Sub txtSSN_GotFocus()
    objHours.txtSSN_GotFocus
End Sub

Private Sub txtSSN_KeyPress(KeyAscii As Integer)
    objHours.txtSSN_KeyPress KeyAscii
End Sub

Private Sub txtSSN_LostFocus()
    objHours.txtSSN_LostFocus
End Sub

Private Sub cmdSSN_Click()
    objHours.cmdSSN_Click
End Sub

Private Sub tblTimeCard_AfterColEdit(ByVal ColIndex As Integer)
    objHours.tblTimeCard_AfterColEdit ColIndex
End Sub

Private Sub tblTimeCard_BeforeColEdit(ByVal ColIndex As Integer, ByVal KeyAscii As Integer, CANCEL As Integer)
    objHours.tblTimeCard_BeforeColEdit ColIndex, KeyAscii, CANCEL
End Sub

Private Sub tblTimeCard_Change()
    objHours.tblTimeCard_Change
End Sub

Private Sub tblTimeCard_FirstRowChange()
    objHours.tblTimeCard_FirstRowChange
End Sub

Private Sub tblTimeCard_GotFocus()
    objHours.tblTimeCard_GotFocus
End Sub

Private Sub tblTimeCard_KeyDown(KeyCode As Integer, Shift As Integer)
    objHours.tblTimeCard_KeyDown KeyCode, Shift
End Sub

Private Sub tblTimeCard_KeyPress(KeyAscii As Integer)
    objHours.tblTimeCard_KeyPress KeyAscii
End Sub

Private Sub tblTimeCard_LeftColChange()
    'objHours.tblTimeCard_LeftColChange
End Sub

Private Sub tblTimeCard_LostFocus()
    objHours.tblTimeCard_LostFocus
End Sub

Private Sub tblTimeCard_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    objHours.tblTimeCard_MouseDown Button, Shift, X, Y
End Sub

Private Sub tblTimeCard_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    objHours.tblTimeCard_RowColChange LastRow, LastCol
End Sub

Private Sub tblTimeCard_SelChange(CANCEL As Integer)
    CANCEL = True
End Sub

Private Sub tblTimeCard_UnboundReadData(ByVal RowBuf As DBTrueGrid.RowBuffer, StartLocation As Variant, ByVal ReadPriorRows As Boolean)
    objHours.tblTimeCard_UnboundReadData RowBuf, StartLocation, ReadPriorRows
End Sub

Private Sub tblProfitCenter_UnboundReadData(ByVal RowBuf As DBTrueGrid.RowBuffer, StartLocation As Variant, ByVal ReadPriorRows As Boolean)
    objHours.tblProfitCenter_UnboundReadData RowBuf, StartLocation, ReadPriorRows
End Sub

Private Sub cmdFloatingBtn_Click()
    objHours.cmdFloatingBtn_Click
End Sub

Private Sub cmdFloatingBtn_GotFocus()
    objHours.cmdFloatingBtn_GotFocus
End Sub

Private Sub cmdFloatingBtn_LostFocus()
    objHours.cmdFloatingBtn_LostFocus
End Sub

Private Function fnInitialPRFHOURSclass() As Boolean
    
    On Error GoTo errTrap
    
    Set objHours = New clsPRFHOURS
    
    Set objHours.MainForm = Me
    Set objHours.FormToolBar = tbToolbar
    Set objHours.StatusBar = ffraStatusbar
    Set objHours.EmpNumTexBox = txtEmployeeNumber
    Set objHours.EmpNumButton = cmdEmployeeNumber
    Set objHours.EmpNameTexBox = txtEmployeeName
    Set objHours.EmpNameButton = cmdEmployeeName
    Set objHours.SSNTexBox = txtSSN
    Set objHours.SSNButton = cmdSSN
    Set objHours.ComboDropDownData = datComboDropDown
    Set objHours.FloatingData = datDropDown
    Set objHours.ComboDropdownGrid = tblComboDropdown
    Set objHours.TimeCardGrid = tblTimeCard
    Set objHours.ProfitCenterGrid = tblProfitCenter
    Set objHours.FloatingGrid = tblDropDown(nTabHours)
    Set objHours.FloatingButton = cmdFloatingBtn
    Set objHours.TotalDollarsTexBox = txtTotalDollars
    Set objHours.TotalTexBox = txtTotal
    Set objHours.AddButton = cmdAddBtn(nTabHours)
    Set objHours.EditButton = cmdEditBtn(nTabHours)
    Set objHours.DeleteButton = cmdDelete(nTabHours)
    Set objHours.UpdateInsertButton = cmdUpdateInsertBtn(nTabHours)
    Set objHours.RefreshSelectButton = cmdRefresh(nTabHours)
    Set objHours.CancelButton = cmdCancel(nTabHours)
    Set objHours.CancelMenuButton = mnuCancel
    Set objHours.ExitButton = cmdExitCancelBtn
    Set objHours.Bridge = efraBaseIIHours
    
    objHours.Form_Initialize
    objHours.Form_Load
    
    fnInitialPRFHOURSclass = True
    
    Exit Function
    
errTrap:
    tfnErrHandler "fnInitialPRFHOURSclass"
End Function

Private Sub subFillStartEndDateFreq()
    
    Dim nMM As Integer
    Dim nDD As Integer
    Dim nYY As Integer
    
    If cValidate.ValidInput(txtStartDate) Then
        Exit Sub
    End If
    
    nMM = Month(tfnDateString(Date))
    nYY = Year(tfnDateString(Date))
    
    'set start date to first day of the month
    nDD = 1
    
    txtStartDate = tfnFormatDate(Format(nMM, "00") + "/" + _
        Format(nDD, "00") + "/" + Format(nYY, "0000"))
    
    nDD = fnLastDayOfMonth(nMM, nYY)
    
    txtEndDate = tfnFormatDate(Format(nMM, "00") + "/" + _
        Format(nDD, "00") + "/" + Format(nYY, "0000"))
        
    txtFrequency = "M"
    
    cValidate.ResetFlags
    cmdProcess.Enabled = cValidate.FirstInvalidInput < 0
End Sub

Private Function fnLastDayOfMonth(nMM As Integer, nYY As Integer) As Integer
    Dim sTemp As String
    Dim nDD As Integer
    
    Select Case nMM
    Case 1, 3, 5, 7, 8, 10, 12
        nDD = 31
    Case 2
        nDD = 29
        sTemp = tfnFormatDate(Format(nMM, "00") + "/" + _
            Format(nDD, "00") + "/" + Format(nYY, "0000"))
        If Not IsDate(nDD) Then
            nDD = 28
        End If
    Case Else
        nDD = 30
    End Select
    
    fnLastDayOfMonth = nDD
End Function

Private Function fnCheckFrequency(sStartDate, sEndDate, sFrequency) As String
    Select Case sFrequency
    Case "W"
    Case "M"
    Case "Q"
    Case "Y"
    End Select
    
End Function
