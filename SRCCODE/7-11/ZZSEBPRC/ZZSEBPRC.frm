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
   ScaleHeight     =   6060
   ScaleWidth      =   8880
   Begin FACTFRMLib.FactorFrame cmdExitCancelBtn 
      Height          =   396
      HelpContextID   =   15
      Left            =   7500
      TabIndex        =   63
      Top             =   5268
      Width           =   1308
      _Version        =   65536
      _ExtentX        =   2307
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
   Begin FACTFRMLib.FactorFrame efraBackground 
      Height          =   5184
      Left            =   0
      TabIndex        =   65
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
         Size            = 8.25 
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
         Size            = 9.75 
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin FACTTABLib.FactorTab eTabMain 
         Height          =   5016
         Left            =   48
         TabIndex        =   67
         TabStop         =   0   'False
         Top             =   12
         Width           =   8796
         _Version        =   65536
         _ExtentX        =   15515
         _ExtentY        =   8848
         _StockProps     =   68
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            = 9.75 
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Pay En&try|Proce&ss Checks|&View/Approve Checks|View &Details"
         Begin FACTFRMLib.FactorFrame efraBaseDetail 
            Height          =   4728
            HelpContextID   =   550
            Left            =   14580
            TabIndex        =   72
            TabStop         =   0   'False
            Top             =   0
            Width           =   8796
            _Version        =   65536
            _ExtentX        =   15515
            _ExtentY        =   8340
            _StockProps     =   77
            BackColor       =   8388608
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            = 9.75 
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
               Size            = 9.75 
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Begin FACTFRMLib.FactorFrame efraBaseIIDetail 
               Height          =   4152
               Left            =   60
               TabIndex        =   73
               TabStop         =   0   'False
               Top             =   60
               Width           =   8664
               _Version        =   65536
               _ExtentX        =   15282
               _ExtentY        =   7324
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
               BeginProperty PanelFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            = 9.75 
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Begin VB.TextBox txtDPrftCtr 
                  BackColor       =   &H00FFFFFF&
                  ForeColor       =   &H00000000&
                  Height          =   360
                  HelpContextID   =   555
                  Left            =   72
                  TabIndex        =   56
                  Tag             =   "pn_alt"
                  Top             =   924
                  Width           =   1104
               End
               Begin VB.TextBox txtDPrftCtrName 
                  BackColor       =   &H00FFFFFF&
                  DataSource      =   "datVendor"
                  ForeColor       =   &H00000000&
                  Height          =   360
                  HelpContextID   =   557
                  Left            =   1608
                  TabIndex        =   58
                  Tag             =   "pn_name"
                  Top             =   924
                  Width           =   4176
               End
               Begin VB.TextBox txtEmployee 
                  BackColor       =   &H00FFFFFF&
                  ForeColor       =   &H00000000&
                  Height          =   360
                  HelpContextID   =   551
                  Left            =   60
                  TabIndex        =   52
                  Tag             =   "pn_alt"
                  Top             =   276
                  Width           =   1872
               End
               Begin VB.TextBox txtEmpName 
                  BackColor       =   &H00FFFFFF&
                  DataSource      =   "datVendor"
                  ForeColor       =   &H00000000&
                  Height          =   360
                  HelpContextID   =   553
                  Left            =   2364
                  TabIndex        =   54
                  Tag             =   "pn_name"
                  Top             =   276
                  Width           =   5868
               End
               Begin DBTrueGrid.TDBGrid tblDetails 
                  Height          =   2724
                  HelpContextID   =   559
                  Left            =   60
                  OleObjectBlob   =   "ZZSEBPRC.frx":0000
                  TabIndex        =   60
                  Top             =   1356
                  Width           =   8556
               End
               Begin FACTFRMLib.FactorFrame cmdEmployee 
                  Height          =   360
                  HelpContextID   =   552
                  Left            =   1944
                  TabIndex        =   53
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
                     Size            = 9.75 
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
                  HelpContextID   =   554
                  Left            =   8244
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
                     Size            = 9.75 
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
                     Size            = 9.75 
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
               End
               Begin FACTFRMLib.FactorFrame cmdDPrftCtr 
                  Height          =   360
                  HelpContextID   =   556
                  Left            =   1188
                  TabIndex        =   57
                  TabStop         =   0   'False
                  Tag             =   "Run #3"
                  Top             =   924
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
                  Picture         =   "ZZSEBPRC.frx":1502
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
               Begin FACTFRMLib.FactorFrame cmdDPrftCtrName 
                  Height          =   360
                  HelpContextID   =   558
                  Left            =   5796
                  TabIndex        =   59
                  TabStop         =   0   'False
                  Tag             =   "Run #3"
                  Top             =   924
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
                  Picture         =   "ZZSEBPRC.frx":1614
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
               Begin VB.Label Label3 
                  Caption         =   "Profit Center"
                  Height          =   252
                  Left            =   72
                  TabIndex        =   104
                  Top             =   684
                  Width           =   1296
               End
               Begin VB.Label Label2 
                  Caption         =   "Profit Center Name"
                  Height          =   252
                  Left            =   1608
                  TabIndex        =   103
                  Top             =   684
                  Width           =   1956
               End
               Begin VB.Label lblEmployee 
                  Caption         =   "Employee Number"
                  Height          =   252
                  Left            =   60
                  TabIndex        =   75
                  Top             =   36
                  Width           =   1836
               End
               Begin VB.Label lblEmpName 
                  Caption         =   "Employee Name"
                  Height          =   252
                  Left            =   2364
                  TabIndex        =   74
                  Top             =   36
                  Width           =   1968
               End
            End
            Begin FACTFRMLib.FactorFrame cmdPrint 
               Height          =   396
               HelpContextID   =   32
               Index           =   3
               Left            =   5988
               TabIndex        =   61
               Top             =   4260
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
               Caption         =   "&Print"
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
            Begin FACTFRMLib.FactorFrame cmdCancel 
               Height          =   396
               HelpContextID   =   15
               Index           =   3
               Left            =   7428
               TabIndex        =   62
               Top             =   4260
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
               Caption         =   "&Cancel"
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
         End
         Begin FACTFRMLib.FactorFrame efraBaseProcess 
            Height          =   4728
            Left            =   14460
            TabIndex        =   69
            TabStop         =   0   'False
            Top             =   0
            Width           =   8796
            _Version        =   65536
            _ExtentX        =   15515
            _ExtentY        =   8340
            _StockProps     =   77
            BackColor       =   8388608
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            = 9.75 
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
               Size            = 9.75 
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Begin FACTFRMLib.FactorFrame efraBaseIIProcess 
               Height          =   4152
               HelpContextID   =   520
               Left            =   60
               TabIndex        =   77
               TabStop         =   0   'False
               Top             =   60
               Width           =   8664
               _Version        =   65536
               _ExtentX        =   15282
               _ExtentY        =   7324
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
               Caption         =   " "
               BeginProperty PanelFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            = 9.75 
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Begin FACTFRMLib.FactorFrame efraProcessDate 
                  Height          =   1332
                  Left            =   72
                  TabIndex        =   99
                  TabStop         =   0   'False
                  Top             =   48
                  Width           =   2400
                  _Version        =   65536
                  _ExtentX        =   4233
                  _ExtentY        =   2350
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
                  Begin VB.TextBox txtFrequency 
                     BackColor       =   &H00FFFFFF&
                     ForeColor       =   &H00000000&
                     Height          =   360
                     HelpContextID   =   523
                     Left            =   1404
                     TabIndex        =   31
                     Tag             =   "pn_alt"
                     Top             =   252
                     Width           =   552
                  End
                  Begin VB.TextBox txtStartDate 
                     BackColor       =   &H00FFFFFF&
                     ForeColor       =   &H00000000&
                     Height          =   360
                     HelpContextID   =   521
                     Left            =   48
                     TabIndex        =   30
                     Tag             =   "pn_alt"
                     Top             =   252
                     Width           =   1224
                  End
                  Begin VB.TextBox txtEndDate 
                     BackColor       =   &H00FFFFFF&
                     ForeColor       =   &H00000000&
                     Height          =   360
                     HelpContextID   =   522
                     Left            =   48
                     TabIndex        =   32
                     Tag             =   "pn_alt"
                     Top             =   900
                     Width           =   1224
                  End
                  Begin FACTFRMLib.FactorFrame cmdFrequency 
                     Height          =   360
                     HelpContextID   =   524
                     Left            =   1968
                     TabIndex        =   33
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
                        Size            = 9.75 
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     CaptionPos      =   4
                     Picture         =   "ZZSEBPRC.frx":1726
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
                  Begin VB.Label lblFrequency 
                     Caption         =   "Frequency"
                     Height          =   252
                     Left            =   1404
                     TabIndex        =   102
                     Top             =   12
                     Width           =   960
                  End
                  Begin VB.Label lblDate 
                     Caption         =   "Starting  Date"
                     Height          =   252
                     Left            =   48
                     TabIndex        =   101
                     Top             =   12
                     Width           =   1488
                  End
                  Begin VB.Label Label1 
                     Caption         =   "Ending Date"
                     Height          =   252
                     Left            =   48
                     TabIndex        =   100
                     Top             =   660
                     Width           =   1236
                  End
               End
               Begin VB.TextBox txtEmpProcess 
                  BackColor       =   &H00FFFFFF&
                  ForeColor       =   &H00000000&
                  Height          =   360
                  HelpContextID   =   529
                  Left            =   2508
                  TabIndex        =   38
                  Tag             =   "pn_alt"
                  Top             =   960
                  Width           =   1104
               End
               Begin VB.TextBox txtEmpNameProcess 
                  BackColor       =   &H00FFFFFF&
                  DataSource      =   "datVendor"
                  ForeColor       =   &H00000000&
                  Height          =   360
                  HelpContextID   =   531
                  Left            =   4044
                  TabIndex        =   40
                  Tag             =   "pn_name"
                  Top             =   960
                  Width           =   4176
               End
               Begin VB.ListBox lstProcess 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  ForeColor       =   &H00000000&
                  Height          =   2352
                  HelpContextID   =   533
                  IntegralHeight  =   0   'False
                  ItemData        =   "ZZSEBPRC.frx":1838
                  Left            =   72
                  List            =   "ZZSEBPRC.frx":183A
                  TabIndex        =   42
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
                  TabIndex        =   36
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
                  TabIndex        =   34
                  Tag             =   "pn_alt"
                  Top             =   300
                  Width           =   1104
               End
               Begin MSComctlLib.ProgressBar pbBarMain 
                  Height          =   312
                  Left            =   72
                  TabIndex        =   78
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
                  TabIndex        =   35
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
                     Size            = 9.75 
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
                     Size            = 9.75 
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
                     Size            = 9.75 
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
                     Size            = 9.75 
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
               End
               Begin FACTFRMLib.FactorFrame cmdEmpProcess 
                  Height          =   360
                  HelpContextID   =   530
                  Left            =   3624
                  TabIndex        =   39
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
                     Size            = 9.75 
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
                     Size            = 9.75 
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
               End
               Begin FACTFRMLib.FactorFrame cmdEmpNameProcess 
                  Height          =   360
                  HelpContextID   =   532
                  Left            =   8232
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
                     Size            = 9.75 
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
                     Size            = 9.75 
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
                  TabIndex        =   96
                  Top             =   720
                  Width           =   1296
               End
               Begin VB.Label lblEmpNameProcess 
                  Caption         =   "Employee Name"
                  Height          =   252
                  Left            =   4044
                  TabIndex        =   95
                  Top             =   720
                  Width           =   1956
               End
               Begin VB.Label lblPrftCtrName 
                  Caption         =   "Profit Center Name"
                  Height          =   252
                  Left            =   4044
                  TabIndex        =   80
                  Top             =   60
                  Width           =   1956
               End
               Begin VB.Label lblPrftCtr 
                  Caption         =   "Profit Center"
                  Height          =   252
                  Left            =   2508
                  TabIndex        =   79
                  Top             =   60
                  Width           =   1296
               End
            End
            Begin FACTFRMLib.FactorFrame cmdProcess 
               Height          =   396
               HelpContextID   =   534
               Left            =   5988
               TabIndex        =   43
               Top             =   4260
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
               Caption         =   "P&rocess"
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
            Begin FACTFRMLib.FactorFrame cmdCancel 
               Height          =   396
               HelpContextID   =   15
               Index           =   1
               Left            =   7428
               TabIndex        =   45
               Top             =   4260
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
               Caption         =   "&Cancel"
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
            Begin FACTFRMLib.FactorFrame cmdPrint 
               Height          =   396
               HelpContextID   =   32
               Index           =   1
               Left            =   48
               TabIndex        =   44
               Top             =   4260
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
               Caption         =   "&Print"
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
         End
         Begin FACTFRMLib.FactorFrame efraBasePayEntry 
            Height          =   4704
            HelpContextID   =   500
            Left            =   12
            TabIndex        =   68
            TabStop         =   0   'False
            Top             =   12
            Width           =   8772
            _Version        =   65536
            _ExtentX        =   15473
            _ExtentY        =   8297
            _StockProps     =   77
            BackColor       =   -2147483633
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            = 9.75 
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
               Size            = 9.75 
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Begin FACTTABLib.FactorTab eTabSub 
               Height          =   4716
               Left            =   0
               TabIndex        =   81
               TabStop         =   0   'False
               Top             =   0
               Width           =   8784
               _Version        =   65536
               _ExtentX        =   15494
               _ExtentY        =   8318
               _StockProps     =   68
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            = 9.75 
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Position        =   3
               TabsPerPage     =   2
               Caption         =   "Store Sa&les|Employee &Hours"
               Begin FACTFRMLib.FactorFrame efraBaseHours 
                  Height          =   4428
                  Left            =   14460
                  TabIndex        =   83
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
                     Size            = 9.75 
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
                     Size            = 9.75 
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Begin FACTFRMLib.FactorFrame efraBaseIIHours 
                     Height          =   4152
                     HelpContextID   =   510
                     Left            =   48
                     TabIndex        =   25
                     Top             =   48
                     Width           =   8376
                     _Version        =   65536
                     _ExtentX        =   14774
                     _ExtentY        =   7324
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
                     BeginProperty PanelFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "MS Sans Serif"
                        Size            = 9.75 
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
                        TabIndex        =   98
                        Top             =   3780
                        Width           =   1272
                     End
                     Begin VB.TextBox txtTotalDollars 
                        Alignment       =   1  'Right Justify
                        Enabled         =   0   'False
                        Height          =   324
                        Left            =   4632
                        MultiLine       =   -1  'True
                        TabIndex        =   97
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
                        TabIndex        =   92
                        Top             =   720
                        Visible         =   0   'False
                        Width           =   255
                     End
                     Begin VB.TextBox txtEmployeeName 
                        BackColor       =   &H00FFFFFF&
                        DataSource      =   "datVendor"
                        ForeColor       =   &H00000000&
                        Height          =   360
                        HelpContextID   =   513
                        Left            =   2028
                        TabIndex        =   19
                        Tag             =   "pn_name"
                        Top             =   276
                        Width           =   3576
                     End
                     Begin VB.TextBox txtEmployeeNumber 
                        BackColor       =   &H00FFFFFF&
                        ForeColor       =   &H00000000&
                        Height          =   360
                        HelpContextID   =   511
                        Left            =   72
                        TabIndex        =   15
                        Tag             =   "pn_alt"
                        Top             =   276
                        Width           =   1524
                     End
                     Begin VB.TextBox txtSSN 
                        Height          =   360
                        HelpContextID   =   515
                        Left            =   6036
                        TabIndex        =   21
                        Top             =   276
                        Width           =   1896
                     End
                     Begin FACTFRMLib.FactorFrame cmdEmployeeNumber 
                        Height          =   360
                        HelpContextID   =   512
                        Left            =   1608
                        TabIndex        =   17
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
                           Size            = 9.75 
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
                           Size            = 9.75 
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                     End
                     Begin FACTFRMLib.FactorFrame cmdEmployeeName 
                        Height          =   360
                        HelpContextID   =   514
                        Left            =   5616
                        TabIndex        =   20
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
                           Size            = 9.75 
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        CaptionPos      =   4
                        Picture         =   "ZZSEBPRC.frx":1D96
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
                     Begin FACTFRMLib.FactorFrame cmdSSN 
                        Height          =   360
                        HelpContextID   =   516
                        Left            =   7944
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
                           Size            = 9.75 
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        CaptionPos      =   4
                        Picture         =   "ZZSEBPRC.frx":1EA8
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
                     Begin DBTrueGrid.TDBGrid tblTimeCard 
                        Height          =   3036
                        HelpContextID   =   517
                        Left            =   72
                        OleObjectBlob   =   "ZZSEBPRC.frx":1FBA
                        TabIndex        =   23
                        Top             =   696
                        Width           =   5892
                     End
                     Begin DBTrueGrid.TDBGrid tblProfitCenter 
                        Height          =   3036
                        HelpContextID   =   518
                        Left            =   6036
                        OleObjectBlob   =   "ZZSEBPRC.frx":2BF9
                        TabIndex        =   24
                        TabStop         =   0   'False
                        Top             =   696
                        Width           =   2268
                     End
                     Begin VB.Label lblTotalHr 
                        Alignment       =   1  'Right Justify
                        Caption         =   "Total Hours/Dollars:"
                        Height          =   252
                        Left            =   2892
                        TabIndex        =   90
                        Top             =   3828
                        Width           =   1692
                     End
                     Begin VB.Label lblSSN 
                        Caption         =   "Social Security Number"
                        Height          =   252
                        Left            =   6036
                        TabIndex        =   89
                        Top             =   24
                        Width           =   2268
                     End
                     Begin VB.Label lblHEmpName 
                        Caption         =   "Employee Name"
                        Height          =   252
                        Left            =   2028
                        TabIndex        =   88
                        Top             =   24
                        Width           =   1968
                     End
                     Begin VB.Label lblEmpNo 
                        Caption         =   "Employee Number"
                        Height          =   252
                        Left            =   72
                        TabIndex        =   87
                        Top             =   24
                        Width           =   1848
                     End
                     Begin VB.Label lblTotalPC 
                        Alignment       =   1  'Right Justify
                        Caption         =   "PC Total:"
                        Height          =   252
                        Left            =   6048
                        TabIndex        =   86
                        Top             =   3828
                        Width           =   888
                     End
                  End
                  Begin FACTFRMLib.FactorFrame cmdAddBtn 
                     Height          =   396
                     HelpContextID   =   10
                     Index           =   4
                     Left            =   36
                     TabIndex        =   11
                     Top             =   4248
                     Width           =   1308
                     _Version        =   65536
                     _ExtentX        =   2307
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
                     Index           =   4
                     Left            =   1452
                     TabIndex        =   13
                     Top             =   4248
                     Width           =   1308
                     _Version        =   65536
                     _ExtentX        =   2307
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
                  Begin FACTFRMLib.FactorFrame cmdCancel 
                     Height          =   396
                     HelpContextID   =   15
                     Index           =   4
                     Left            =   7128
                     TabIndex        =   27
                     Top             =   4248
                     Width           =   1308
                     _Version        =   65536
                     _ExtentX        =   2307
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
                     Caption         =   "&Cancel"
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
                     Index           =   4
                     Left            =   5712
                     TabIndex        =   26
                     Top             =   4248
                     Width           =   1308
                     _Version        =   65536
                     _ExtentX        =   2307
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
                  Begin FACTFRMLib.FactorFrame cmdDelete 
                     Height          =   396
                     HelpContextID   =   12
                     Index           =   4
                     Left            =   2868
                     TabIndex        =   29
                     Top             =   4248
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
                  Begin FACTFRMLib.FactorFrame cmdRefresh 
                     Height          =   396
                     HelpContextID   =   14
                     Index           =   4
                     Left            =   4296
                     TabIndex        =   28
                     Top             =   4248
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
               End
               Begin FACTFRMLib.FactorFrame efraBaseSales 
                  Height          =   4692
                  HelpContextID   =   501
                  Left            =   12
                  TabIndex        =   82
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
                     Size            = 9.75 
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
                     Size            = 9.75 
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Begin FACTFRMLib.FactorFrame efraBaseIISales 
                     Height          =   4152
                     Left            =   48
                     TabIndex        =   10
                     Top             =   48
                     Width           =   8376
                     _Version        =   65536
                     _ExtentX        =   14774
                     _ExtentY        =   7324
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
                     BeginProperty PanelFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "MS Sans Serif"
                        Size            = 9.75 
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Begin VB.TextBox txtSalesType 
                        BackColor       =   &H00FFFFFF&
                        ForeColor       =   &H00000000&
                        Height          =   360
                        HelpContextID   =   502
                        Left            =   96
                        TabIndex        =   3
                        Tag             =   "pn_alt"
                        Top             =   336
                        Width           =   2220
                     End
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
                        TabIndex        =   91
                        Top             =   768
                        Visible         =   0   'False
                        Width           =   255
                     End
                     Begin VB.TextBox txtFromDate 
                        BackColor       =   &H00FFFFFF&
                        ForeColor       =   &H00000000&
                        Height          =   360
                        HelpContextID   =   504
                        Left            =   4392
                        TabIndex        =   5
                        Tag             =   "pn_alt"
                        Top             =   336
                        Width           =   1488
                     End
                     Begin VB.TextBox txtToDate 
                        BackColor       =   &H00FFFFFF&
                        DataSource      =   "datVendor"
                        ForeColor       =   &H00000000&
                        Height          =   360
                        HelpContextID   =   506
                        Left            =   6444
                        TabIndex        =   7
                        Tag             =   "pn_name"
                        Top             =   336
                        Width           =   1488
                     End
                     Begin DBTrueGrid.TDBGrid tblSales 
                        Height          =   3348
                        HelpContextID   =   508
                        Left            =   72
                        OleObjectBlob   =   "ZZSEBPRC.frx":383C
                        TabIndex        =   9
                        Top             =   768
                        Width           =   8244
                     End
                     Begin FACTFRMLib.FactorFrame cmdFromDate 
                        Height          =   360
                        HelpContextID   =   505
                        Left            =   5892
                        TabIndex        =   6
                        TabStop         =   0   'False
                        Tag             =   "Run #3"
                        Top             =   336
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
                        Picture         =   "ZZSEBPRC.frx":4B18
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
                     Begin FACTFRMLib.FactorFrame cmdToDate 
                        Height          =   360
                        HelpContextID   =   507
                        Left            =   7944
                        TabIndex        =   8
                        TabStop         =   0   'False
                        Tag             =   "Run #3"
                        Top             =   336
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
                        Picture         =   "ZZSEBPRC.frx":4C2A
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
                     Begin FACTFRMLib.FactorFrame cmdSalesType 
                        Height          =   360
                        HelpContextID   =   503
                        Left            =   2340
                        TabIndex        =   4
                        TabStop         =   0   'False
                        Tag             =   "Run #3"
                        Top             =   336
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
                        Picture         =   "ZZSEBPRC.frx":4D3C
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
                     Begin VB.Label Label4 
                        Caption         =   "Sales Type"
                        Height          =   252
                        Left            =   108
                        TabIndex        =   106
                        Top             =   84
                        Width           =   1380
                     End
                     Begin VB.Label lblFromDate 
                        Caption         =   "From Date"
                        Height          =   252
                        Left            =   4404
                        TabIndex        =   85
                        Top             =   84
                        Width           =   1380
                     End
                     Begin VB.Label lblToDate 
                        Caption         =   "To Date"
                        Height          =   252
                        Left            =   6444
                        TabIndex        =   84
                        Top             =   84
                        Width           =   1500
                     End
                  End
                  Begin FACTFRMLib.FactorFrame cmdUpdateInsertBtn 
                     Height          =   396
                     HelpContextID   =   13
                     Index           =   0
                     Left            =   5712
                     TabIndex        =   12
                     Top             =   4248
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
                  Begin FACTFRMLib.FactorFrame cmdCancel 
                     Height          =   396
                     HelpContextID   =   15
                     Index           =   0
                     Left            =   7128
                     TabIndex        =   14
                     Top             =   4248
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
                     Caption         =   "&Cancel"
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
                  Begin FACTFRMLib.FactorFrame cmdDelete 
                     Height          =   396
                     HelpContextID   =   12
                     Index           =   0
                     Left            =   2868
                     TabIndex        =   18
                     Top             =   4248
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
                  Begin FACTFRMLib.FactorFrame cmdRefresh 
                     Height          =   396
                     HelpContextID   =   14
                     Index           =   0
                     Left            =   4296
                     TabIndex        =   16
                     Top             =   4248
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
               End
            End
         End
         Begin FACTFRMLib.FactorFrame efraBaseView 
            Height          =   4728
            HelpContextID   =   540
            Left            =   14520
            TabIndex        =   71
            TabStop         =   0   'False
            Top             =   0
            Width           =   8796
            _Version        =   65536
            _ExtentX        =   15515
            _ExtentY        =   8340
            _StockProps     =   77
            BackColor       =   8388608
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            = 9.75 
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
               Size            = 9.75 
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Begin FACTFRMLib.FactorFrame efraBaseIIView 
               Height          =   4152
               Left            =   60
               TabIndex        =   76
               TabStop         =   0   'False
               Top             =   60
               Width           =   8664
               _Version        =   65536
               _ExtentX        =   15282
               _ExtentY        =   7324
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
               BeginProperty PanelFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            = 9.75 
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Begin VB.PictureBox picTextExtension 
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000005&
                  ForeColor       =   &H80000008&
                  Height          =   288
                  Left            =   4116
                  ScaleHeight     =   264
                  ScaleWidth      =   1488
                  TabIndex        =   105
                  TabStop         =   0   'False
                  Top             =   384
                  Width           =   1512
               End
               Begin DBTrueGrid.TDBGrid tblApprove 
                  Height          =   4044
                  HelpContextID   =   541
                  Left            =   60
                  OleObjectBlob   =   "ZZSEBPRC.frx":4E4E
                  TabIndex        =   46
                  Top             =   60
                  Width           =   8556
               End
            End
            Begin FACTFRMLib.FactorFrame cmdOk 
               Height          =   396
               HelpContextID   =   16
               Left            =   5988
               TabIndex        =   49
               Top             =   4260
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
               Caption         =   "O&K"
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
            Begin FACTFRMLib.FactorFrame cmdCancel 
               Height          =   396
               HelpContextID   =   15
               Index           =   2
               Left            =   7428
               TabIndex        =   51
               Top             =   4260
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
               Caption         =   "&Cancel"
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
            Begin FACTFRMLib.FactorFrame cmdSelectAll 
               Height          =   396
               HelpContextID   =   542
               Left            =   48
               TabIndex        =   47
               Top             =   4260
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
               Caption         =   "&Select All"
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
            Begin FACTFRMLib.FactorFrame cmdApprove 
               Height          =   396
               HelpContextID   =   543
               Left            =   1476
               TabIndex        =   48
               Top             =   4260
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
               Caption         =   "&Approve"
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
            Begin FACTFRMLib.FactorFrame cmdPrint 
               Height          =   396
               HelpContextID   =   32
               Index           =   2
               Left            =   2904
               TabIndex        =   50
               Top             =   4260
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
               Caption         =   "&Print"
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
         End
      End
      Begin DBTrueGrid.TDBGrid tblComboDropdown 
         Bindings        =   "ZZSEBPRC.frx":612C
         Height          =   2484
         Left            =   108
         OleObjectBlob   =   "ZZSEBPRC.frx":614B
         TabIndex        =   70
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
      TabIndex        =   64
      TabStop         =   0   'False
      Top             =   5700
      Width           =   8880
      _Version        =   65536
      _ExtentX        =   15663
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
      Height          =   468
      Left            =   0
      TabIndex        =   66
      TabStop         =   0   'False
      Top             =   0
      Width           =   8880
      _Version        =   65536
      _ExtentX        =   15663
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
      FMName          =   "ZZSEBPRC"
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
      Bindings        =   "ZZSEBPRC.frx":742F
      Height          =   1008
      Index           =   0
      Left            =   168
      OleObjectBlob   =   "ZZSEBPRC.frx":7449
      TabIndex        =   93
      Top             =   684
      Width           =   2604
   End
   Begin DBTrueGrid.TDBGrid tblDropDown 
      Bindings        =   "ZZSEBPRC.frx":872B
      Height          =   1008
      Index           =   4
      Left            =   144
      OleObjectBlob   =   "ZZSEBPRC.frx":8745
      TabIndex        =   94
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
      Begin VB.Menu sep02 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDetailLog 
         Caption         =   "Show &Detail Log"
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
'Programmer     : David Chai, Junsong Qiu                                           *
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

Private tgfDropdown(4) As clsFloatingDropDown

Private cValidate As cValidateInput
Private cValidSls As cValidateInput
Private cValidDetail As cValidateInput

Private bSalesRecordExist As Boolean
Private sFreqRegExp As String

Private objHours As clsPRFHOURS

Private bLoadingBonusDetail As Boolean
Private bProcessing As Boolean
Private bCancelProcess As Boolean
'

Private Sub cmdApprove_Click()
    Dim lRow As Long
    
    lRow = tgmApprove.GetCurrentRowNumber
    
    If tgsApprove.Count = 0 Then
        tgmApprove.CellValue(colAApprove, tgmApprove.GetCurrentRowNumber) = sColAppYes
        tblApprove.col = 1
        tblApprove.col = 0
    Else
        subSetApproveAll
        tblApprove_LostFocus
        tblApprove_GotFocus
    End If
    
    cmdOk.Enabled = True
    cmdCancel(TabApprove).Enabled = True
    
    tgmApprove.Rebind
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
    #If PROTOTYPE Then
        Exit Sub
    #End If
    
    If Not fnCheckApprove() Then
        Exit Sub
    End If
    
    cmdOk.Enabled = False
    Me.Enabled = False
    
    Dim sErrMsg As String
    
    sErrMsg = fnInsertHoldBonus()
    
    If sErrMsg <> "" Then
        Me.Enabled = True
        cmdOk.Enabled = True
        subSetFocus cmdOk
        DoEvents
        tfnSetStatusBarError sErrMsg
        Exit Sub
    End If
    
    Me.Enabled = True
    
    'nDataStatus = DATA_INIT
    'tgmApprove.ClearData
    
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
    If txtEmployee.Enabled Then
        subSetFocus txtEmployee
    Else
        subSetFocus tblDetails
    End If
End Sub

Private Sub efraBaseIIHours_GotFocus()
    If objHours Is Nothing Then Exit Sub
    
    objHours.efraBridge_GotFocus
End Sub

Private Sub efraBaseIIProcess_GotFocus()
    'subFillStartEndDateFreq
    subBuildFrequencyRegExp
    
    If txtStartDate.Enabled Then
        subSetFocus txtStartDate
    ElseIf cmdPrint(TabProcess).Enabled Then
        subSetFocus cmdPrint(TabProcess)
    ElseIf cmdProcess.Enabled Then
        subSetFocus cmdProcess
    Else
        subSetFocus cmdCancel(TabProcess)
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
            frmContext.ButtonEnabled(FO_HOLD_UP) = False
            #If PROTOTYPE Then
                tblApprove.Enabled = False
                Exit Sub
            #End If
            subSetFocus efraBaseIIView
                
            If fnHasApprove() Then
                cmdOk.Enabled = True
                cmdCancel(TabApprove).Enabled = True
            Else
                cmdOk.Enabled = False
                cmdCancel(TabApprove).Enabled = False
            End If
            
        Case TabDetails
            frmContext.ButtonEnabled(FO_HOLD_UP) = (tgmDetail.RowCount > 0) 'True
            
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

Private Sub eTabMain_Switch(OldTab As Integer, NewTab As Integer, Cancel As Integer)
    If Not eTabMain.TabEnabled(NewTab) Then
        Cancel = True
        Exit Sub
    End If
    If Not Cancel Then t_bTabSwitch = True
End Sub

Private Sub eTabSub_Click()
    Select Case eTabSub.CurrTab
        Case TabSales
            subSetFocus efraBaseSales
        Case TabHours
            subSetFocus efraBaseHours
    End Select
End Sub

Private Sub eTabSub_Switch(OldTab As Integer, NewTab As Integer, Cancel As Integer)
    If Not eTabSub.TabEnabled(NewTab) Then
        Cancel = True
        Exit Sub
    End If

    If Not Cancel Then
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

Private Sub Form_Unload(Cancel As Integer)
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

        If Not fnCreateSalesTable() Then
            sErrorMessage = "Failed to create temporary Sales Type Table. Program terminates"
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
        If bProcessing Then
            bCancelProcess = True
            Exit Sub
        End If
        
        If nDataStatus = DATA_CHANGED Then
            If Not tfnCancelExit(t_szCANCEL_MESSAGE) Then
                Screen.MousePointer = vbDefault
                Exit Sub
            End If
        End If
        
    End If
    
    If Index = TabApprove Then
        
        If fnHasApprove() Then
            
            If Not tfnCancelExit(t_szCANCEL_MESSAGE) Then
                Screen.MousePointer = vbDefault
                Exit Sub
            End If
            
        End If
        
    End If
    
    tfnResetScreen Index
    Screen.MousePointer = vbDefault

End Sub

Private Sub subExit()
    Screen.MousePointer = vbDefault
    
    If nDataStatus = DATA_CHANGED Then
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
Private Sub mnuDetailLog_Click()
    If mnuDetailLog.CHECKED Then
        mnuDetailLog.CHECKED = False
    Else
        mnuDetailLog.CHECKED = True
    End If

    bShowDetail = mnuDetailLog.CHECKED
End Sub

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
Public Sub tfnSetStatusBarError(szErrorMessage As String, Optional vNoBeep As Variant)
    
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
            txtSalesType = ""
            txtFromDate = ""
            txtToDate = ""
            sSalesTypeCode = ""
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
            nDataStatus = DATA_INIT
            txtStartDate = ""
            txtEndDate = ""
            txtFrequency = ""
            txtPrftCtr = ""
            txtPrftCtrName = ""
            txtEmpProcess = ""
            txtEmpNameProcess = ""
            bLoadingBonusDetail = False

            bShowDetail = mnuDetailLog.CHECKED

            cValidate.ResetFlags
            
            eTabMain.TabEnabled(TabSales) = True
            eTabMain.TabEnabled(TabDetails) = False
            eTabMain.TabEnabled(TabApprove) = False
            subEnablePrint Index, False
            subEnableFirstLineProcess True
            bProcessing = False
            bCancelProcess = False
            cmdProcess.Enabled = False
            subSetProgress 0
            
            tgmApprove.ClearData
            
            txtEmployee = ""
            txtEmpName = ""
            txtDPrftCtr = ""
            txtDPrftCtrName = ""
            bLoadingBonusDetail = False
            tgmDetail.ClearData
            subEnableEmployee True
            subEnableDPrftCtr True
            
            If eTabMain.CurrTab = TabProcess Then
                subLogErrMsg "", True
                'subFillStartEndDateFreq
                subSetFocus txtStartDate
            End If
        Case TabApprove
           ' If fnHasApprove() Then
            '    If tfnCancelExit(t_szCANCEL_MESSAGE) Then
                    tgmApprove.ClearData
                    tgmApprove.FillWithArray vArrBonus
                    If eTabMain.CurrTab = TabApprove Then
                        subSetFocus tblApprove
                        
                        On Error Resume Next
                        If tgmApprove.RowCount > 1 Then
                            tblApprove.Row = 1
                            tblApprove.Row = 0
                        Else
                            tblApprove.col = 2
                            tblApprove.col = 0
                        End If
                        
                        If fnHasApprove() Then
                            cmdOk.Enabled = True
                            cmdCancel(Index).Enabled = True
                        Else
                            cmdOk.Enabled = False
                            cmdCancel(Index).Enabled = False
                        End If
                        
                    End If
                    
             '   End If
                
            'End If
            
            
            bLoadingBonusDetail = False
        Case TabDetails
            txtEmployee = ""
            txtEmpName = ""
            txtDPrftCtr = ""
            txtDPrftCtrName = ""
            bLoadingBonusDetail = False
            tgmDetail.ClearData
            subEnableEmployee True
            subEnableDPrftCtr True
            'tblDetails.Enabled = False
            tblDetails.Enabled = True
            subEnablePrint Index, False
            cValidDetail.ResetFlags
           
            If eTabMain.CurrTab = TabDetails Then
                subSetFocus txtEmployee
            End If
    
    End Select
    
    frmContext.ButtonEnabled(COPY_UP) = False
    frmContext.ButtonEnabled(FO_HOLD_UP) = False
    mnuExit.Enabled = True
    mnuCopy.Enabled = False
    mnuPaste.Enabled = False
    
    If (Index <> nTabHours) Then
        cmdRefresh(Index).Enabled = False
        cmdUpdateInsertBtn(Index).Enabled = False
        cmdDelete(Index).Enabled = False
    End If
    
    Screen.MousePointer = vbDefault
    
End Sub

Private Sub mnuPrint_Click()
    subPrint eTabMain.CurrTab
End Sub

Private Sub tblApprove_AfterColEdit(ByVal ColIndex As Integer)
    tgmApprove.AfterColEdit ColIndex
End Sub

Private Sub tblApprove_BeforeColEdit(ByVal ColIndex As Integer, ByVal KeyAscii As Integer, Cancel As Integer)
    tgmApprove.BeforeColEdit ColIndex, KeyAscii, Cancel
    
    If ColIndex = colAApprove Then
        
        If fnHasApprove() Then
            cmdOk.Enabled = True
            cmdCancel(TabApprove).Enabled = True
        Else
            cmdOk.Enabled = False
            cmdCancel(TabApprove).Enabled = False
        End If
        
    End If
    
End Sub

Private Sub tblApprove_Change()
    tgmApprove.Change
End Sub

Private Sub tblApprove_Click()
'    If tblApprove.col = colAApprove Then
'        tblApprove.col = 1
'        tblApprove.col = 0
'    End If
'
    tgsApprove.Click
End Sub

Private Sub tblApprove_DblClick()
    bLoadingBonusDetail = False
    subEnterBonusPhaseII
End Sub

Private Sub tblApprove_FirstRowChange()
    tgmApprove.FirstRowChange
End Sub

Private Sub tblApprove_GotFocus()
    tfnSetStatusBarMessage "Press enter key to see commission details"
    tgsApprove.GotFocus
    tgmApprove.GotFocus
End Sub

Private Sub tblApprove_HeadClick(ByVal ColIndex As Integer)
    tgmApprove.HeadClick ColIndex
End Sub

Private Sub tblApprove_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 0 And KeyCode = vbKeyReturn Then
        KeyCode = 0
        
        If tgsApprove.Count > 1 Then
            tfnSetStatusBarError "Only one detail can be viewed at a time"
            Exit Sub
        End If
        
        bLoadingBonusDetail = False
        subEnterBonusPhaseII
    Else
        tgsApprove.KeyDown KeyCode, Shift
        tgmApprove.KeyDown KeyCode, Shift
    End If
End Sub

Private Sub tblApprove_KeyPress(KeyAscii As Integer)
    If KeyAscii <> vbKeyReturn Then
        Exit Sub
    End If
End Sub

Private Sub tblApprove_LostFocus()
    tgmApprove.LostFocus
End Sub

Private Sub tblApprove_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    tgsApprove.MouseUp Button, Shift, y
End Sub

Private Sub tblApprove_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    tgmApprove.RowColChange LastRow, LastCol
    tgsApprove.RowColChange LastRow, LastCol
End Sub

Private Sub tblApprove_SelChange(Cancel As Integer)
    tgsApprove.SelChange Cancel
    Cancel = True
End Sub

Private Sub tblApprove_UnboundReadData(ByVal RowBuf As DBTrueGrid.RowBuffer, StartLocation As Variant, ByVal ReadPriorRows As Boolean)
    tgmApprove.ReadData RowBuf, StartLocation, ReadPriorRows
End Sub

Private Sub tblDetails_DblClick()
    subShowFormulaDetails
End Sub

Private Sub tblDetails_GotFocus()
    tfnSetStatusBarMessage "Press enter key to see formula details"
End Sub

Private Sub tblDetails_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        subShowFormulaDetails
    End If
End Sub

Private Sub tblDetails_SelChange(Cancel As Integer)
    Cancel = True
End Sub

Private Sub tblDetails_UnboundReadData(ByVal RowBuf As DBTrueGrid.RowBuffer, StartLocation As Variant, ByVal ReadPriorRows As Boolean)
    tgmDetail.ReadData RowBuf, StartLocation, ReadPriorRows
End Sub

Private Sub tblSales_Click()
    tgsSales.Click
End Sub

Private Sub tmrKeyboard_Timer() 'status bar timer - 250ms
    tfnUpdateStatusBar Me 'process the status bar

    If Not tgcExtension Is Nothing Then
        tgcExtension.ShowColumnExt
    End If
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
            If t_oleObject Is Nothing Then
                .AddButton "Add &Profit Center", PROFITCENTER_UP, , True
                .AddButton "Add Emplo&yee", EMP_MST_UP, , True
            Else
                .AddButton "Add &Profit Center", PROFITCENTER_UP
                .AddButton "Add Emplo&yee", EMP_MST_UP
            End If
            
            .AddButton "Add Commission Co&de", PRDCLS_UP, , True
            .AddButton "Add Commission &Formula", SYS_LOCKS_UP, , True
            .AddButton "Launch Commssion &Master", POFAPLVL_UP, , True
            .AddButton "&View Formula Details", FO_HOLD_UP, , True
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
        Case FO_HOLD_UP
            subShowFormulaDetails
        Case PRINT_UP
            subPrint eTabMain.CurrTab
        Case PROFITCENTER_UP
            tfnRun "syfprftc"
        Case EMP_MST_UP
            tfnRun "prfmstmn"
        Case PRDCLS_UP
            tfnRun "zzsebcmt"
        Case SYS_LOCKS_UP
            tfnRun "zzsebfmt"
        Case POFAPLVL_UP
            tfnRun "zzsebmfm"
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
        If MsgBox("Are you sure you want to cancel the Commission calculation process?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
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
    Dim sEmpNo As String
    Dim sPrftCtr As String
    Dim sCode As String
    Dim nLevel As Integer
    Dim dBLvlAmt As Double
    Dim dTotalBonus As Double
    Dim nSize As Integer
    Dim i As Integer
    Dim bError As Boolean: bError = False
    
    Dim sSysParm30854 As String
    
    #If PROTOTYPE Then
        Exit Sub
    #End If
    
    If Not tfnCancelExit("Processing may take several minutes. Are you sure you want to continue?") Then
        Exit Sub
    End If
    
    bNoRecordFound = False
    
    subLogErrMsg "", True
    
    subLogErrMsg "Commission Processing" + IIf(bShowDetail, " Detail", "") + " Log"
    
    If bShowDetail Then
        subLogErrMsg "Log will be saved in " + sLogFilePath
    End If
    
    subLogErrMsg " "
    
    subLogErrMsg "Started processing commission formulas..."
    subLogErrMsg " "
    
    i = fnCheckBonusHold()
    
    If i = vbCancel Then
        subLogErrMsg "Processing terminates."
        Exit Sub
    End If
    
    If i = vbNo Then
        subLogErrMsg "Processing terminated on user's request."
        Exit Sub
    End If
    
    'get sysparm#30854
    strSQL = "SELECT parm_field FROM sys_parm WHERE parm_nbr = 30854"
    
    If GetRecordSet(rsTemp, strSQL, , SUB_NAME) < 0 Then
        subLogErrMsg "Failed to access the database."
        subLogErrMsg "Processing terminates."
        Exit Sub
    End If
    
    If rsTemp.RecordCount > 0 Then
        
        If IsNull(rsTemp!parm_field) Then
            
            If bShowDetail Then
                subLogErrMsg "SysParm#30854 is NULL"
            End If
        
        Else
            sSysParm30854 = UCase(rsTemp!parm_field)
            
            If bShowDetail Then
                subLogErrMsg "SysParm#30854 = " + tfnSQLString(sSysParm30854)
            End If
            
            If Len(sSysParm30854) >= 4 Then
                sPayCode_RegHrs = Left(sSysParm30854, 4)
            End If
            
            If Len(sSysParm30854) >= 9 Then
                sPayCode_OtHrs = Trim(Mid(sSysParm30854, 6, 4))
            End If
        
            If bShowDetail Then
                subLogErrMsg "Pay Code for Regular Hour = " + tfnSQLString(sPayCode_RegHrs)
                subLogErrMsg "Pay Code for Overtime Hour = " + tfnSQLString(sPayCode_OtHrs)
            End If
            
        End If
        
    Else
        subLogErrMsg "SysParm#30854 not found"
    End If
    
    subLogErrMsg " "
    
    ReDim vArrBonus(colAHdnBAmtLvls, 0)
    eTabMain.TabEnabled(TabSales) = False
    bProcessing = True
    bCancelProcess = False
    cmdProcess.Enabled = False
    eTabMain.TabEnabled(TabApprove) = False
    eTabMain.TabEnabled(TabDetails) = False
    subEnablePrint TabProcess, False
    subEnableFirstLineProcess False
    
    strSQL = "SELECT bm_empno, bc_type, bc_grade, bc_bonus_code, bc_code_desc, bf_level, "
    strSQL = strSQL & " bm_eligible_pc, bm_sequence, bm_override, prft_name"
    strSQL = strSQL & " FROM bonus_master, bonus_codes, bonus_formula, sys_prft_ctr"
    strSQL = strSQL & " WHERE bm_bonus_code = bc_bonus_code"
    strSQL = strSQL & " AND bm_bonus_code = bf_bonus_code"
    strSQL = strSQL & " AND bm_eligible_pc = prft_ctr"
    
    If cValidate.ValidInput(txtPrftCtr) And txtPrftCtr <> "" Then
        strSQL = strSQL & " AND bm_eligible_pc = " & tfnRound(txtPrftCtr)
    End If
    
    If cValidate.ValidInput(txtEmpProcess) And txtEmpProcess <> "" Then
        strSQL = strSQL & " AND bm_empno = " & tfnRound(txtEmpProcess)
    End If
    
    If cValidate.ValidInput(txtFrequency) And txtFrequency <> "" Then
        strSQL = strSQL & " AND bc_frequency = " & tfnSQLString(Trim(txtFrequency))
    End If
    
    strSQL = strSQL & " AND " & tfnDateString(txtStartDate, True)
    strSQL = strSQL & " BETWEEN bm_eligible_date AND bm_stop_date"

'    strSQL = strSQL & " AND " & tfnDateString(txtEndDate, True)
'    strSQL = strSQL & " BETWEEN bm_eligible_date AND bm_stop_date"
    strSQL = strSQL & " ORDER BY bm_empno, bm_eligible_pc, bm_sequence, bc_bonus_code, bf_level"
    
    Screen.MousePointer = vbHourglass
    nCount = GetRecordSet(rsTemp, strSQL, , SUB_NAME)
    
    If nCount < 0 Then
        subLogErrMsg "Failed to access the database."
        bError = True
        GoTo TERMINATE_PROCESS
    End If
    
    If nCount = 0 Then
        subLogErrMsg "No record found to process."
        'bNoRecordFound = True
        bError = True
        GoTo TERMINATE_PROCESS
    End If
    
    rsTemp.MoveFirst
    nSize = -1
    
    For i = 1 To nCount
        DoEvents
        If bCancelProcess Then
            GoTo TERMINATE_PROCESS
        End If
        
        subSetProgress i * (100 / nCount)
        
        Screen.MousePointer = vbHourglass
        
        nLevel = tfnRound(rsTemp!bf_level)
        
        subLogErrMsg "Calculating commission for employee " & tfnRound(rsTemp!bm_empno) _
            & ", profit center " & tfnRound(rsTemp!bm_eligible_pc) & ", commission code " _
            & tfnSQLString(rsTemp!bc_bonus_code) & ", level " & nLevel
        
        If sEmpNo <> fnGetField(rsTemp!bm_empno) Or _
           sPrftCtr <> fnGetField(rsTemp!bm_eligible_pc) Or _
           sCode <> fnGetField(rsTemp!bc_bonus_code) Then
            
            If nSize >= 0 Then
                vArrBonus(colABonusAmt, nSize) = Format(dTotalBonus, "##,##0.00")
            
                If bShowDetail Then
                    subLogErrMsg "Commission Code " + tfnSQLString(rsTemp!bc_bonus_code) _
                        + " calculation result: "
                    subLogErrMsg "Total = " & vArrBonus(colABonusAmt, nSize) _
                        & "(" & vArrBonus(colAHdnBAmtLvls, nSize) & ")"
                End If
            End If
            
            nSize = nSize + 1
            ReDim Preserve vArrBonus(colAHdnBAmtLvls, nSize)
            vArrBonus(colAApprove, nSize) = sColAppNo
            vArrBonus(colAEmpNo, nSize) = fnGetField(rsTemp!bm_empno)
            vArrBonus(colAEmpName, nSize) = fnGetEmployeeName(fnGetField(rsTemp!bm_empno))
            vArrBonus(colADate, nSize) = txtEndDate
            vArrBonus(colAPrftCtr, nSize) = fnGetField(rsTemp!bm_eligible_pc)
            vArrBonus(colAPayCode, nSize) = fnGetField(rsTemp!bc_bonus_code)
            vArrBonus(colABonusAmt, nSize) = Format(dTotalBonus, "##,##0.00")
            vArrBonus(colAHdsOverride, nSize) = fnGetField(rsTemp!bm_override)
            vArrBonus(colAHdnPrftName, nSize) = fnGetField(rsTemp!prft_name) 'Hidden Column
            vArrBonus(colAHdsBonusDesc, nSize) = fnGetField(rsTemp!bc_code_desc) 'Hidden Column
            vArrBonus(colAHdnSeq, nSize) = tfnRound(rsTemp!bm_sequence) 'Hidden Column
            vArrBonus(colAHdnBAmtLvls, nSize) = "" 'Hidden Column
            
            dBLvlAmt = fnGetBonusAmount(rsTemp)
            dTotalBonus = dBLvlAmt
            
            vArrBonus(colAHdnBAmtLvls, nSize) = nLevel & "*" & fnGetField(dBLvlAmt)
        Else  'everything (empno, prftctr, paycode) is the same, except the bonus code level
            dBLvlAmt = fnGetBonusAmount(rsTemp)
            dTotalBonus = dTotalBonus + dBLvlAmt
            vArrBonus(colAHdnBAmtLvls, nSize) = vArrBonus(colAHdnBAmtLvls, nSize) _
                + "~" & nLevel & "*" & fnGetField(dBLvlAmt)
        End If
        
        sCode = fnGetField(rsTemp!bc_bonus_code)
        sEmpNo = fnGetField(rsTemp!bm_empno)
        sPrftCtr = fnGetField(rsTemp!bm_eligible_pc)
        
        'last record...
        If i = nCount Then
            vArrBonus(colABonusAmt, nSize) = Format(dTotalBonus, "##,##0.00")
        
            If bShowDetail Then
                subLogErrMsg "Commission Code " + tfnSQLString(rsTemp!bc_bonus_code) _
                    + " calculation result: "
                subLogErrMsg "Total = " & vArrBonus(colABonusAmt, nSize) _
                    & "(" & vArrBonus(colAHdnBAmtLvls, nSize) & ")"
            End If
            
        End If
        
        rsTemp.MoveNext
    Next i
    
    tgmApprove.FillWithArray vArrBonus
    
    nDataStatus = DATA_CHANGED
    
    cmdPrint(TabApprove).Enabled = True
    eTabMain.TabEnabled(TabDetails) = True
    eTabMain.TabEnabled(TabApprove) = True
    eTabMain.CurrTab = TabApprove
    subSetFocus tblApprove
    
TERMINATE_PROCESS:
    
    bProcessing = False
    
    subLogErrMsg " "
    
    If bCancelProcess Then
        subLogErrMsg "Processing terminated on user's request."
        cmdProcess.Enabled = True
    End If
    
    If bError Then
        cmdProcess.Enabled = True
    End If
    
    subLogErrMsg "*Finished Processing*"
    
    If bNoRecordFound Then
        MsgBox "Data was found to be missing while processing the commissions. This " _
             & "could cause the comminsions to be miscalculated. Please review the " _
             & "Process Checks Log, and re-process if neccessary.", vbInformation + vbOKOnly
    End If
    
    Screen.MousePointer = vbDefault
    
    subSetProgress 0
    subEnablePrint TabProcess, (tgmApprove.RowCount > 0)  'True
    
    If bError Then
        subSetFocus cmdProcess
    End If
    
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
'            myWidth = Array(0.21, 0.21, 0.16, 0.16, 0.26)
'            myField = Array("prh_date", "prh_prft_ctr", "prh_pay_code", "prh_pay_type", "prh_hours")
            myWidth = Array(0.25, 0.25, 0.2, 0.3)
            myField = Array("bh_date", "bh_prft_ctr", "bh_pay_code", "bh_hours")
        Case "tblProfitCenter"
            myWidth = Array(0.5, 0.5)
            myField = Array("", "")
        Case "tblApprove"
            myWidth = Array(0.1, 0.13, 0.22, 0.12, 0.1, 0.1, 0.11, 0.12)
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
            tbl.Columns(colHHrsDol).Caption = "Hours/Dollars"
            tbl.Columns(colHHrsDol).Alignment = vbRightJustify
        Case "tblProfitCenter"
            tbl.Caption = "Profit Center Total"
            tbl.Columns(colPProfit).Caption = "Profit"
            tbl.Columns(colPTotal).Caption = "Total"
            tbl.Columns(colPTotal).Alignment = vbRightJustify
        Case "tblApprove"
            tbl.Columns(colAApprove).ValueItems.MaxComboItems = 2
            Set vitems = tbl.Columns(colAApprove).ValueItems
            VItem.Value = sColAppYes: VItem.DisplayValue = "Y": vitems.Add VItem
            VItem.Value = sColAppNo: VItem.DisplayValue = "N": vitems.Add VItem
            vitems.Presentation = 1
            vitems.CycleOnClick = True
            vitems.Translate = True
            'vitems.DefaultItem = sColAppNo
            tbl.Caption = "Commission Approval"
            tbl.HeadLines = 2
            tbl.Columns(colAApprove).Caption = "Approve"
            tbl.Columns(colAEmpNo).Caption = "Employee Number"
            tbl.Columns(colAEmpName).Caption = "Employee Name"
            tbl.Columns(colADate).Caption = "Process Date"
            tbl.Columns(colAPrftCtr).Caption = "Profit Center"
            tbl.Columns(colAPayCode).Caption = "Pay Code"
            tbl.Columns(colAPayHours).Caption = "Pay Hours"
            tbl.Columns(colABonusAmt).Caption = "Comm. Amount"
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
            tbl.Columns(colDBAmt).Caption = "Comm. Amount"
            tbl.Columns(colDBAmt).Alignment = vbRightJustify
    End Select
End Sub

Private Sub subInitSpreadsheets()
    Dim sDecimalString As String
    
    On Error GoTo errTrap
    
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
    With tgmApprove
        Set .Table = tblApprove
        Set .StatusBar = ffraStatusbar ' message bar name
        Set .Form = Me
        Set .engFactor = t_engFactor
        .AddEditColumn colAApprove, "Select Yes, No"
        .AllowAddNew = False
                    
        colAHdsOverride = .AddHiddenField("HiddenOverride")
        colAHdnPrftName = .AddHiddenField("HiddenPrftName")
        colAHdsBonusDesc = .AddHiddenField("HiddenBonusDesc")
        colAHdnSeq = .AddHiddenField("HiddenSeq")
        colAHdnBAmtLvls = .AddHiddenField("HiddenLevels")
        
        .SortByColumn = True
        
        .AddSortColumn colAEmpNo, colAEmpNo, .NUMERIC_TYPE, .ASCENDING, .CASE_SENSITIVE, _
            colAPrftCtr, .NUMERIC_TYPE, .ASCENDING, .CASE_SENSITIVE, _
            colAHdnSeq, .NUMERIC_TYPE, .ASCENDING, .CASE_SENSITIVE, _
            colAPayCode, .STRING_TYPE, .ASCENDING, .CASE_SENSITIVE
    
        .AddSortColumn colAPrftCtr, colAPrftCtr, .NUMERIC_TYPE, .ASCENDING, .CASE_SENSITIVE, _
            colAEmpNo, .NUMERIC_TYPE, .ASCENDING, .CASE_SENSITIVE, _
            colAHdnSeq, .NUMERIC_TYPE, .ASCENDING, .CASE_SENSITIVE, _
            colAPayCode, .STRING_TYPE, .ASCENDING, .CASE_SENSITIVE
    
        .AddSortColumn colAPayCode, colAPrftCtr, .NUMERIC_TYPE, .ASCENDING, .CASE_SENSITIVE, _
            colAEmpNo, .NUMERIC_TYPE, .ASCENDING, .CASE_SENSITIVE, _
            colAPayCode, .STRING_TYPE, .ASCENDING, .CASE_SENSITIVE
    
    End With
    
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

    colDHdnEmpNo = tgmDetail.AddHiddenField("bm_empno")
    colDHdnPrftCtr = tgmDetail.AddHiddenField("bm_eligible_pc")
    tgmDetail.DisplayFormat(colDBAmt) = "###,###,##0.00"

    'setup the text extension
    subSetupExetension
    
    Exit Sub

errTrap:
    tfnErrHandler "subInitSpreadsheets"
    'add by junsong 03/16/01
    Unload Me
    
End Sub

Private Sub subSetupExetension()
    Set tgcExtension = New clsColumnExtension
    Set tgcExtension.Form = Me
    Set tgcExtension.PictureBox = picTextExtension
    Set tgcExtension.Table = tblApprove
    
    'changed by junsong 03/20/01
    'tgcExtension.PositionDefault = tgcExtension.PositionTop
    tgcExtension.UseEditorRow = True
    'end changed
    
    tgcExtension.AddColumn colAPrftCtr
    tgcExtension.Style(colAPrftCtr) = 3
    
    tgcExtension.AddColumn colAPayCode
    tgcExtension.Style(colAPayCode) = 3
End Sub

Public Function fnGetText(tgTable As TDBGrid, ByVal nCol As Integer, ByVal nRow As Integer) As String
    Select Case nCol
    Case colAPrftCtr
        fnGetText = tgmApprove.CellValue(colAHdnPrftName, nRow)
    Case colAPayCode
        fnGetText = tgmApprove.CellValue(colAHdsBonusDesc, nRow)
    End Select
End Function

'put all codes in tmrKeyboard_Timer()
'Private Sub tmrExtension_Timer()
'    If Not tgcExtension Is Nothing Then
'        tgcExtension.ShowColumnExt
'    End If
'End Sub

Private Sub subInitValidation()
    
    'Class implementation for Sales Tab
    Set cValidSls = New cValidateInput
    Set cValidSls.Form = Me
    Set cValidSls.StatusBar = ffraStatusbar
    cValidSls.AddEditBox txtSalesType, "Enter or select Sales Type"
    cValidSls.AddEditBox txtFromDate, "Enter From Date"
    cValidSls.AddEditBox txtToDate, "Enter To Date"
    cValidSls.MinTabIndex = txtSalesType.TabIndex
    cValidSls.MaxTabIndex = tblSales.TabIndex
    Set cValidSls.ControlForFocus = efraBaseIISales
    Set cValidSls.LastBox = txtToDate
    cValidSls.SetFirstControls cmdDelete(TabSales), cmdRefresh(TabSales), cmdCancel(TabSales), cmdUpdateInsertBtn(TabSales), cmdExitCancelBtn

    'Class implementation for Process Tab
    Set cValidate = New cValidateInput
    Set cValidate.Form = Me
    Set cValidate.StatusBar = ffraStatusbar
    cValidate.AddEditBox txtStartDate, "Enter Starting Date"
    cValidate.AddEditBox txtEndDate, "Enter Ending Date"
    cValidate.AddEditBox txtFrequency, "Enter Frequency"
    cValidate.AddEditBox txtPrftCtr, "Enter Profit Center Number"
    cValidate.AddEditBox txtEmpProcess, "Enter Employee Number"
    cValidate.MinTabIndex = txtStartDate.TabIndex
    cValidate.MaxTabIndex = txtEmpName.TabIndex
    cValidate.ESCControl = cmdCancel(TabProcess)
    cValidate.ESCControl = cmdExitCancelBtn
    
    'Class implementation for Details Tab
    Set cValidDetail = New cValidateInput
    Set cValidDetail.Form = Me
    Set cValidDetail.StatusBar = ffraStatusbar
    cValidDetail.AddEditBox txtEmployee, "Enter Employee Number"
    cValidDetail.AddEditBox txtEmpName, "Enter Employee Name"
    cValidDetail.GreenMessage(txtEmpName) = False
    cValidDetail.AddEditBox txtDPrftCtr, "Enter Profit Center Number"
    cValidDetail.AddEditBox txtDPrftCtrName, "Enter Profit Center Name"
    cValidDetail.GreenMessage(txtDPrftCtrName) = False
    cValidDetail.MinTabIndex = txtEmployee.TabIndex
    cValidDetail.MaxTabIndex = tblDetails.TabIndex
    cValidDetail.ESCControl = cmdPrint(TabDetails)
    cValidDetail.ESCControl = cmdCancel(TabDetails)
    cValidDetail.ESCControl = cmdExitCancelBtn
    
End Sub

Private Function fnSetComboSQL(nTabIndex As Integer) As String
    Dim strSQL As String
    Dim sApproveEmpList As String
    Dim sApprovePrftCtrList As String
    
    Select Case nTabIndex
        Case txtPrftCtr.TabIndex, txtPrftCtrName.TabIndex
            strSQL = "SELECT prft_ctr, prft_name FROM sys_prft_ctr WHERE prft_ctr IN "
            strSQL = strSQL & "(SELECT DISTINCT bm_eligible_pc FROM bonus_master)"
        Case txtFrequency.TabIndex
            strSQL = "SELECT bfq_frequency, bfq_freq_desc FROM bonus_frequency"
        Case txtEmpProcess.TabIndex, txtEmpNameProcess.TabIndex
            strSQL = "SELECT prm_empno, prm_empname FROM sTmpEmpTable"
            strSQL = strSQL & " WHERE prm_empno IN (SELECT bm_empno FROM bonus_master)"
        Case txtEmployee.TabIndex, txtEmpName.TabIndex
            strSQL = "SELECT prm_empno, prm_empname FROM sTmpEmpTable"
            strSQL = strSQL & " WHERE prm_empno IN (SELECT bm_empno FROM bonus_master)"
            
            If cValidDetail.ValidInput(txtDPrftCtr) Then
                sApproveEmpList = fnBuildList(tgmApprove, colAEmpNo, 1, False, True, True, colAPrftCtr, txtDPrftCtr)
            Else
                sApproveEmpList = fnBuildList(tgmApprove, colAEmpNo, 1, False, True, True)
            End If
            
            If sApproveEmpList <> "" Then
                strSQL = strSQL & " AND prm_empno IN (" + sApproveEmpList + ")"
            End If
        Case txtDPrftCtr.TabIndex, txtDPrftCtrName.TabIndex
            strSQL = "SELECT prft_ctr, prft_name FROM sys_prft_ctr"
            strSQL = strSQL & " WHERE prft_ctr IN (SELECT DISTINCT bm_eligible_pc FROM bonus_master)"
            
            If cValidDetail.ValidInput(txtEmployee) Then
                sApprovePrftCtrList = fnBuildList(tgmApprove, colAPrftCtr, 1, False, True, True, colAEmpNo, txtEmployee)
            Else
                sApprovePrftCtrList = fnBuildList(tgmApprove, colAPrftCtr, 1, False, True, True)
            End If
            
            If sApprovePrftCtrList <> "" Then
                strSQL = strSQL & " AND prft_ctr IN (" + sApprovePrftCtrList + ")"
            End If
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
        Case txtSalesType.TabIndex
            fnInValidData = Not fnValidSalesType(txtBox)
        Case txtFromDate.TabIndex, txtToDate.TabIndex
            fnInValidData = Not fnValidSalesDate(txtBox)
        Case txtStartDate.TabIndex, txtEndDate.TabIndex
            fnInValidData = Not fnValidProcessDate(txtBox)
        Case txtFrequency.TabIndex
            fnInValidData = Not fnValidFrequency(txtBox)
        Case txtPrftCtr.TabIndex
            fnInValidData = Not fnValidPrftCtr(txtBox)
        Case txtEmpProcess.TabIndex
            fnInValidData = Not fnValidEmpProcess(txtBox)
        Case txtEmployee.TabIndex, txtDPrftCtr.TabIndex
            fnInValidData = Not fnValidDetailEmpPrftCtr(txtBox)
        Case txtEmployeeNumber.TabIndex, txtEmployeeName.TabIndex, txtSSN.TabIndex
            fnInValidData = objHours.fnInValidData(txtBox)
        Case Else
            fnInValidData = False
    End Select
End Function

Private Function fnGetPrftName(nPrftCtr As Integer) As String
    Const SUB_NAME As String = "fnGetPrftName"
    
    Dim strSQL As String
    Dim rsTemp As Recordset
    
    strSQL = "SELECT prft_name FROM sys_prft_ctr"
    strSQL = strSQL + " WHERE prft_ctr = " & nPrftCtr
    
    If GetRecordSet(rsTemp, strSQL, , SUB_NAME) <= 0 Then
        Exit Function
    End If

    fnGetPrftName = fnGetField(rsTemp!prft_name)
End Function

Private Function fnValidFrequency(txtBox As Textbox) As Boolean
    Const SUB_NAME As String = "fnValidFrequency"
    Dim strSQL As String
    Dim sErrMsg As String

    fnValidFrequency = False
    
    If Trim(txtBox) = "" Then
        cValidate.SetErrorMessage txtBox, "You must enter a Commission frequency"
        Exit Function
    End If
    
    strSQL = "SELECT * FROM bonus_frequency WHERE bfq_frequency = " & tfnSQLString(txtBox)
    
    If GetRecordCount(strSQL, , SUB_NAME) <= 0 Then
        cValidate.SetErrorMessage txtBox, "Commission frequency does not exist"
        Exit Function
    End If
    
    If IsValidDate(txtStartDate) And txtEndDate = "" Then
        txtEndDate = fnGetProposedEndDate(txtStartDate, txtFrequency)
        cValidate.ResetFlags txtEndDate
    End If
    
    If cValidate.ValidInput(txtStartDate) And cValidate.ValidInput(txtEndDate) Then
        sErrMsg = fnCheckFrequency(txtStartDate, txtEndDate, txtFrequency, False)
        
        If sErrMsg <> "" Then
            cValidate.SetErrorMessage txtFrequency, sErrMsg
            cValidate.SetErrorMessage txtStartDate, sErrMsg
            cValidate.SetErrorMessage txtEndDate, sErrMsg
            cValidate.ValidInput(txtStartDate) = False
            cValidate.ValidInput(txtEndDate) = False
            Exit Function
        End If
        
    End If
    
    fnValidFrequency = True
End Function

Private Sub subBuildFrequencyRegExp()
    Const SUB_NAME As String = "subBuildFrequencyRegExp"
    
    Dim strSQL As String
    Dim rsTemp As Recordset
    
    #If PROTOTYPE Then
        Exit Sub
    #End If
    
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

Private Function fnValidEmpProcess(txtBox As Textbox) As Boolean
    Const SUB_NAME As String = "fnValidEmpProcess"
    Dim strSQL As String
    Dim rsTemp As Recordset
    Dim sEmpName As String
    
    fnValidEmpProcess = False
    
    If Trim(txtBox.Text) = "" Then
        If eTabMain.CurrTab = TabDetails Then
            cValidate.SetErrorMessage txtBox, "You must enter an Employee Number"
        Else
            fnValidEmpProcess = True
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
        cValidate.SetErrorMessage txtBox, "Employee is not set up in Commission Master"
        Exit Function
    End If
    
    fnValidEmpProcess = True
    
End Function

Private Function fnValidDetailEmpPrftCtr(txtBox As Textbox) As Boolean
    Dim sTemp As String
    Dim lEmpNo As Long
    Dim nPrftCtr As Integer
    Dim i As Long
    
    If txtBox Is txtEmployee Then
        sTemp = "Employee Number"
    Else
        sTemp = "Profit Center"
    End If
    
    If txtBox = "" Then
        cValidDetail.SetErrorMessage txtBox, "You must enter a" + IIf(Left(sTemp, 1) = "E", "n ", " ") + sTemp
        Exit Function
    End If
    
    If txtBox Is txtEmployee Then
        If Not cValidDetail.ValidInput(txtDPrftCtr) Then
            fnValidDetailEmpPrftCtr = True
            Exit Function
        End If
    Else
        If Not cValidDetail.ValidInput(txtEmployee) Then
            fnValidDetailEmpPrftCtr = True
            Exit Function
        End If
    End If
    
    If txtBox Is txtEmployee Then
        lEmpNo = tfnRound(txtBox)
        nPrftCtr = tfnRound(txtDPrftCtr)
    Else
        nPrftCtr = tfnRound(txtBox)
        lEmpNo = tfnRound(txtEmployee)
    End If
    
    If txtDPrftCtr <> "" Then
        txtDPrftCtrName = fnGetPrftName(nPrftCtr)
    End If
    
    If txtDPrftCtr <> "" And txtEmployee <> "" Then
        For i = 0 To tgmApprove.RowCount - 1
            If tfnRound(tgmApprove.CellValue(colAEmpNo, i)) = lEmpNo And _
               tfnRound(tgmApprove.CellValue(colAPrftCtr, i)) = nPrftCtr Then
                If txtBox Is txtEmployee Then
                    cValidDetail.ValidInput(txtDPrftCtr) = True
                Else
                    cValidDetail.ValidInput(txtEmployee) = True
                End If
                
                fnValidDetailEmpPrftCtr = True
                Exit Function
            End If
        Next i
    
        If txtBox Is txtEmployee Then
            cValidDetail.ValidInput(txtDPrftCtr) = False
            cValidDetail.SetErrorMessage txtDPrftCtr, "Employee/Profit Center is not in the Comm. Approval Grid"
        Else
            cValidDetail.ValidInput(txtEmployee) = False
            cValidDetail.SetErrorMessage txtEmployee, "Employee/Profit Center is not in the Comm. Approval Grid"
        End If
        
        cValidDetail.SetErrorMessage txtBox, "Employee/Profit Center is not in the Comm. Approval Grid"
    Else
        fnValidDetailEmpPrftCtr = True
    End If
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

Private Sub tblComboDropDown_SelChange(Cancel As Integer)
    tgcDropdown.SelChange Cancel
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
        .AddComboBox txtPrftCtr, cmdPrftCtr, "prft_ctr", .SQL_INT_TYPE
        .AddComboBox txtPrftCtrName, cmdPrftCtrName, "prft_name", .SQL_STRING_TYPE(40)
     
        .AddCombo
        .AddComboBox txtEmpProcess, cmdEmpProcess, "prm_empno", .SQL_LONG_TYPE
        .AddComboBox txtEmpNameProcess, cmdEmpNameProcess, "prm_empname", .SQL_STRING_TYPE(60)
     
        .AddCombo
        .AddComboBox txtFrequency, cmdFrequency, "bfq_frequency", .SQL_STRING_TYPE(1)
        .AddExtraColumn "bfq_freq_desc", 1300
        .SetExtend txtFrequency, 2
        
        .AddCombo "SELECT * FROM tmp_sales_type"
        .AddComboBox txtSalesType, cmdSalesType, "tst_desc", .SQL_STRING_TYPE(20)
        .AddExtraColumn "tst_type"
        'ole combo not working???
        .SetColumnCaptions "Sales Type Description", "Type"
        
        .AddCombo
        .AddComboBox txtFromDate, cmdFromDate, "bs_from_date", .SQL_DATE_TYPE
        .SetOrderingDescent txtFromDate
        .AddComboBox txtToDate, cmdToDate, "bs_to_date", .SQL_DATE_TYPE
        .SetOrderingDescent txtToDate
             
        .AddCombo
        .AddComboBox txtEmployee, cmdEmployee, "prm_empno", .SQL_LONG_TYPE
        .AddComboBox txtEmpName, cmdEmpName, "prm_empname", .SQL_STRING_TYPE(60)
     
        .AddCombo
        .AddComboBox txtDPrftCtr, cmdDPrftCtr, "prft_ctr", .SQL_INT_TYPE
        .AddComboBox txtDPrftCtrName, cmdDPrftCtrName, "prft_name", .SQL_STRING_TYPE(40)
     End With
End Sub

Private Sub subEnterBonusPhaseII()
    Dim aryLevels() As String
    Dim aryLvlAmt() As String
    Dim aryLevel() As Integer
    Dim aryAmount() As Double
    Dim lEmpNo As Long
    Dim nPrftCtr As Integer
    Dim sBonusCode As String
    Dim lRow As Long
    Dim i As Long
    Dim j As Integer
    
    If bLoadingBonusDetail Or tgmApprove.RowCount <= 0 Then
        Exit Sub
    End If
    
    bLoadingBonusDetail = True
    
    lRow = tgmApprove.GetCurrentRowNumber
    
    If eTabMain.CurrTab = TabApprove Then
        txtEmployee = tgmApprove.CellValue(colAEmpNo, lRow)
        txtEmpName = tgmApprove.CellValue(colAEmpName, lRow)
        txtDPrftCtr = tgmApprove.CellValue(colAPrftCtr, lRow)
        txtDPrftCtrName = fnGetPrftName(txtDPrftCtr)
        lEmpNo = tfnRound(txtEmployee)
        nPrftCtr = tfnRound(txtDPrftCtr)
        sBonusCode = tgmApprove.CellValue(colAPayCode, lRow)
    Else
        If txtEmployee = "" Then
            lEmpNo = -1
        Else
            lEmpNo = tfnRound(txtEmployee)
        End If
        If txtDPrftCtr = "" Then
            nPrftCtr = -1
        Else
            nPrftCtr = tfnRound(txtDPrftCtr)
        End If
        sBonusCode = ""
    End If
    
    Screen.MousePointer = vbHourglass
    subEnableEmployee False
    subEnableDPrftCtr False
    tblDetails.Enabled = True
    
    If Not fnLoadBonusDetails(lEmpNo, nPrftCtr, sBonusCode) Then
        cmdCancel_Click TabDetails
        Exit Sub
    End If
    
    'fill more data in detail grid
    i = 0
    Do
        lEmpNo = tfnRound(tgmDetail.CellValue(colDHdnEmpNo, i))
        nPrftCtr = tfnRound(tgmDetail.CellValue(colDHdnPrftCtr, i))
        sBonusCode = fnGetField(tgmDetail.CellValue(colDBCode, i))
        
        'find the row in tgmApprove
        For lRow = 0 To tgmApprove.RowCount - 1
            If lEmpNo = fnGetField(tgmApprove.CellValue(colAEmpNo, lRow)) And _
               nPrftCtr = fnGetField(tgmApprove.CellValue(colAPrftCtr, lRow)) And _
               sBonusCode = fnGetField(tgmApprove.CellValue(colAPayCode, lRow)) Then
                Exit For
            End If
        Next lRow
        
        If lRow < tgmApprove.RowCount Then
            'get level amount
            aryLevels = Split(tgmApprove.CellValue(colAHdnBAmtLvls, lRow), "~")
            ReDim aryLevel(UBound(aryLevels))
            ReDim aryAmount(UBound(aryLevels))
            For j = 0 To UBound(aryLevels)
                aryLvlAmt = Split(aryLevels(j), "*")
                aryLevel(j) = tfnRound(aryLvlAmt(0))
                aryAmount(j) = tfnRound(aryLvlAmt(1), 2)
            Next j
            
            j = 0
            
            While lEmpNo = tfnRound(tgmDetail.CellValue(colDHdnEmpNo, i)) And _
               nPrftCtr = tfnRound(tgmDetail.CellValue(colDHdnPrftCtr, i)) And _
               sBonusCode = fnGetField(tgmDetail.CellValue(colDBCode, i))
                tgmDetail.CellValue(colDBCDesc, i) = tgmApprove.CellValue(colAHdsBonusDesc, lRow)
                If j <= UBound(aryLevels) Then
                    tgmDetail.CellValue(colDBAmt, i) = aryAmount(j)
                Else
                    tgmDetail.CellValue(colDBAmt, i) = 0
                End If
                i = i + 1
                j = j + 1
            Wend
        Else
            i = i + 1
        End If
    Loop Until i > tgmDetail.RowCount - 1
    
    
    tgmDetail.Rebind
    
    If eTabMain.CurrTab = TabApprove Then
        eTabMain.TabEnabled(TabDetails) = True
        eTabMain.CurrTab = TabDetails
    End If
    
    frmContext.ButtonEnabled(FO_HOLD_UP) = True
    subEnablePrint TabDetails, (tgmDetail.RowCount > 0)  'True
    subSetFocus tblDetails
    Screen.MousePointer = vbDefault
    
End Sub

Private Sub subEnableEmployee(bOnOff As Boolean)
    txtEmployee.Enabled = bOnOff
    cmdEmployee.Enabled = bOnOff
    txtEmpName.Enabled = bOnOff
    cmdEmpName.Enabled = bOnOff
    subEnableSearchbtn cmdEmployee, bOnOff
    subEnableSearchbtn cmdEmpName, bOnOff
End Sub

Private Sub subEnableDPrftCtr(bOnOff As Boolean)
    txtDPrftCtr.Enabled = bOnOff
    cmdDPrftCtr.Enabled = bOnOff
    txtDPrftCtrName.Enabled = bOnOff
    cmdDPrftCtrName.Enabled = bOnOff
    subEnableSearchbtn cmdDPrftCtr, bOnOff
    subEnableSearchbtn cmdDPrftCtrName, bOnOff
End Sub

Private Sub subShowFormulaDetails()
    Dim sCode As String
    Dim nLevel As Integer
    Dim nRow As Integer
    
    If tgmDetail.RowCount <= 0 Then
        Exit Sub
    End If
    
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
        subSetFocus txtFrequency
        KeyAscii = 0
    Else
        tfnRegExpControlDateKeyPress txtStartDate, KeyAscii
        cValidate.Keypress txtStartDate, KeyAscii
    End If
    
End Sub

Private Sub txtStartDate_LostFocus()
    cValidate.LostFocus txtStartDate
    subFillStartEndDateFreq
    cmdProcess.Enabled = cValidate.FirstInvalidInput < 0
    
    ' add following statements to overcome the problem
    ' when txtfrequency got focus, but validate start date is not done
    ' and the default value is not set, after default text set, we need
    ' highlight it. junsong 03/13/01
    
    On Error Resume Next
    
    If ActiveControl Is txtFrequency Then
        txtFrequency_GotFocus
    End If
    
End Sub

Private Sub txtEndDate_Change()
    cmdProcess.Enabled = False
    cValidate.Change txtEndDate
End Sub

Private Sub txtEndDate_GotFocus()
    cValidate.GotFocus txtEndDate
    SelectIt txtEndDate
End Sub

Private Sub txtEndDate_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        subSetFocus txtPrftCtr
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
    SelectIt txtFrequency
    
    If tgcDropdown.SingleRecordSelected Then
        subSetFocus txtEndDate
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
                  subSetFocus txtEndDate
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
    If ActiveControl Is txtPrftCtr Then
        subLogErrMsg "", True
    End If
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
        tfnSetStatusBarMessage "Report was printed successfully"
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

Private Sub subSetApproveAll()
    Dim i As Long
    Dim lCount As Long
    Dim lTemp() As Long
    
    tgsApprove.GetSelected lTemp, lCount
    
    For i = 0 To lCount - 1
        Screen.MousePointer = vbHourglass
        tgmApprove.CellValue(colAApprove, lTemp(i)) = sColAppYes
    Next i
    
    Screen.MousePointer = vbDefault
    
End Sub

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
        
        subSetFocus txtSalesType
        
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
                        If Not fnDeleteSales(sSalesTypeCode, sAryPrftCtr(i), txtToDate, txtFromDate) Then
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
            
            If Not tfnCancelExit("Are you sure you want to delete the current record?") Then
                Exit Sub
            End If
                
            If t_nFormMode = EDIT_MODE Then
                sPrftCtr = fnCstr(tgmSales.CellValue(colSPrftCtr, tgmSales.GetCurrentRowNumber))
                If Not fnDeleteSales(sSalesTypeCode, sPrftCtr, txtToDate, txtFromDate) Then
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
        
        subSetFocus txtSalesType
        
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
                If MsgBox("Sales record(s) already exist for From Date/To Date and will be replaced." _
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
            eTabMain.TabEnabled(TabProcess) = True
            eTabSub.TabEnabled(TabSales) = True
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
    
    Select Case Index
        Case TabSales
            strSQL = "SELECT prft_ctr, prft_name FROM sys_prft_ctr"
            strSQL = strSQL + " WHERE prft_type IN ('R', 'B')"
            
            Dim sPrftCtrList As String
            sPrftCtrList = fnBuildList(tgmSales, colSPrftCtr, 1, True, False, True)
            
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

Private Sub tblDropDown_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    If Index = nTabHours Then
        objHours.tblFloating_MouseUp Button, Shift, x, y
    Else
        tgfDropdown(Index).MouseUp y
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
        
        If t_nFormMode = ADD_MODE Then
            subSetStdBtn TabSales, tgmSales
        End If
        
        tblSales.Enabled = True
        DoEvents
        subSetFocus tblSales
    End If
    
    Screen.MousePointer = vbDefault
End Sub

Private Sub subEnableFirstLineSlsOrHrs(Index As Integer, bYesNo As Boolean)
    Select Case Index
        Case TabSales
            subEnableSalesType bYesNo
            txtFromDate.Enabled = bYesNo
            txtToDate.Enabled = bYesNo
            If t_nFormMode = ADD_MODE Then
                bYesNo = False
            End If
            subEnableSearchbtn cmdFromDate, bYesNo
            subEnableSearchbtn cmdToDate, bYesNo
    End Select
End Sub

Private Sub subEnableSalesType(bYesNo As Boolean)
    txtSalesType.Enabled = bYesNo
    subEnableSearchbtn cmdSalesType, bYesNo
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

Private Function fnValidSalesType(Box As Textbox) As Boolean
    Dim i As Integer
    
    If Trim(Box) = "" Then
        cValidSls.SetErrorMessage Box, "You must enter a Sales Type"
        Exit Function
    End If

    For i = 0 To UBound(arySalesDesc)
        If arySalesDesc(i) = Box.Text Then
            sSalesTypeCode = arySalesType(i)
            fnValidSalesType = True
            Exit Function
        End If
    Next i
    
    sSalesTypeCode = ""
    cValidSls.SetErrorMessage Box, "Sales Type is not valid"
End Function

Private Function fnValidSalesDate(Box As Textbox) As Boolean
    Const SUB_NAME As String = "fnValidSalesDate"
    Dim strSQL As String
    Dim sTemp As String
    Dim sErrMsg As String
    Dim sFreq As String
    
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
    
    sFreq = sSalesTypeCode
    
    If Box Is txtFromDate Then
        If Not IsValidDate(txtToDate) Then
            If txtToDate <> "" Or t_nFormMode = EDIT_MODE Then 'Or sFreq = sGas Then
                fnValidSalesDate = True
                Exit Function
            End If
            
            txtToDate = fnGetProposedEndDate(txtFromDate, sFreq)
            cValidSls.ValidInput(txtToDate) = True
            SelectIt txtToDate
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
        sErrMsg = fnCheckFrequency(txtFromDate, txtToDate, sFreq)
        If sErrMsg <> "" Then
            cValidSls.SetErrorMessage txtFromDate, sErrMsg
            cValidSls.SetErrorMessage txtToDate, sErrMsg
            If Box Is txtFromDate Then
                cValidSls.ValidInput(txtToDate) = False
            Else
                cValidSls.ValidInput(txtFromDate) = False
            End If
            
            Exit Function
        End If
        
        sErrMsg = fnCheckSales(sFreq)
        If sErrMsg <> "" Then
            cValidSls.SetErrorMessage txtFromDate, sErrMsg
            cValidSls.SetErrorMessage txtToDate, sErrMsg
            If Box Is txtFromDate Then
                cValidSls.ValidInput(txtToDate) = False
            Else
                cValidSls.ValidInput(txtFromDate) = False
            End If
            
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
            If cValidate.ValidInput(txtFrequency) And txtEndDate = "" Then
    
                txtEndDate = fnGetProposedEndDate(txtStartDate, txtFrequency)
                cValidate.ResetFlags txtEndDate
            Else
                fnValidProcessDate = True
                Exit Function
            End If
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
    
    If cValidate.ValidInput(txtFrequency) Then
        sErrMsg = fnCheckFrequency(txtStartDate, txtEndDate, txtFrequency)
        If sErrMsg <> "" Then
            cValidate.SetErrorMessage txtStartDate, sErrMsg
            cValidate.SetErrorMessage txtEndDate, sErrMsg
            cValidate.SetErrorMessage txtFrequency, sErrMsg
            cValidate.ValidInput(txtFrequency) = False
            If Box Is txtStartDate Then
                cValidate.ValidInput(txtEndDate) = False
            Else
                cValidate.ValidInput(txtStartDate) = False
            End If
            Exit Function
        End If
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

Private Sub txtSalesType_Change()
    cValidSls.Change txtSalesType
    tgcDropdown.Change txtSalesType
End Sub

Private Sub txtSalesType_GotFocus()
    cValidSls.GotFocus txtSalesType
    tgcDropdown.GotFocus txtSalesType
    If tgcDropdown.SingleRecordSelected Then
        subSetFocus txtFromDate
        Screen.MousePointer = vbDefault
    End If
End Sub

Private Sub txtSalesType_KeyPress(KeyAscii As Integer)
    Dim bKeyCode As Boolean
    
    If KeyAscii = vbKeyReturn Then
        Screen.MousePointer = vbHourglass
    End If
        
    bKeyCode = tgcDropdown.Keypress(txtSalesType, KeyAscii)
    Screen.MousePointer = vbDefault
    
    If Not bKeyCode Then
        If KeyAscii = vbKeyReturn Then
            If tgcDropdown.SingleRecordSelected Then
                  subSetFocus txtFromDate
            End If
            KeyAscii = 0
       End If
    Else
        cValidSls.Keypress txtSalesType, KeyAscii
    End If
End Sub

Private Sub txtSalesType_LostFocus()
    cValidSls.LostFocus txtSalesType, cmdSalesType, tblComboDropdown
    tgcDropdown.LostFocus txtSalesType
End Sub

Private Sub cmdSalesType_Click()
    Screen.MousePointer = vbHourglass
    tgcDropdown.Click cmdSalesType
    Screen.MousePointer = vbDefault
End Sub

Private Sub txtFromDate_Change()
    cValidSls.Change txtFromDate
    tgcDropdown.Change txtFromDate
    If ActiveControl Is txtFromDate Then
        subEnableSalesType False
    End If
End Sub

Private Sub txtFromDate_GotFocus()
    cValidSls.GotFocus txtFromDate
    tgcDropdown.GotFocus txtFromDate
    If tgcDropdown.SingleRecordSelected Then
        subEnableSalesType False
        
        If t_nFormMode = EDIT_MODE Then
            subEnterPhaseIISlsOrHrs TabSales
        Else
            subSetFocus txtToDate
            Screen.MousePointer = vbDefault
        End If
        
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
                    
                    If t_nFormMode = EDIT_MODE Then
                        subEnterPhaseIISlsOrHrs TabSales
                    Else
                        subSetFocus txtToDate
                        Screen.MousePointer = vbDefault
                    End If
                    
                End If
            Else
                subSetFocus txtToDate
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
    If ActiveControl Is txtToDate Then
        subEnableSalesType False
    End If
End Sub

Private Sub txtToDate_GotFocus()
    cValidSls.GotFocus txtToDate
    tgcDropdown.GotFocus txtToDate
    If tgcDropdown.SingleRecordSelected Then
        subEnableSalesType False
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
        
        If Not (ActiveControl Is cmdCancel(TabSales) Or ActiveControl Is cmdExitCancelBtn Or ActiveControl Is txtFromDate) Then
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

    subSetStdBtn TabSales, tgmSales
End Sub

Private Sub tblSales_BeforeColEdit(ByVal ColIndex As Integer, ByVal KeyAscii As Integer, Cancel As Integer)
    If ColIndex = colSPrftCtr Then
        tgmSales.CellValue(colSPrftName, tgmSales.GetCurrentRowNumber) = ""
    End If
    tgmSales.BeforeColEdit ColIndex, KeyAscii, Cancel
End Sub

Private Sub tblSales_Change()
    tgmSales.Change
    subEnableUpdateBtn False, TabSales
    nDataStatus = DATA_CHANGED
End Sub

Private Sub tblSales_FirstRowChange()
    tgmSales.FirstRowChange
    tgfDropdown(TabSales).FirstRowChange
    
    If tblSales.Row = -1 Then
        If tgsSales.Count = 0 Then
            SubEnableDeleteBtn False, TabSales
        Else
            SubEnableDeleteBtn True, TabSales
        End If
    Else
        SubEnableDeleteBtn True, TabSales
    End If
End Sub

Private Sub tblSales_GotFocus()
    tfnSetStatusBarMessage "Store Sales"
    tgsSales.GotFocus
    tgmSales.GotFocus
    tgfDropdown(TabSales).GotFocus
    
    If tgfDropdown(TabSales).ValidSelection Then
        tblSales_AfterColEdit tblSales.col
    End If
End Sub

Private Sub tblSales_KeyDown(KeyCode As Integer, Shift As Integer)
    If (Not tblSales.EditActive) And tblSales.SelBookmarks.Count > 0 Then
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
End Sub

Private Sub tblSales_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    tgsSales.MouseUp Button, Shift, y
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
    
    If tblSales.col = colSPrftName Then
        If tgmSales.ValidCell(colSPrftCtr, lRow) Then
            If tfnRound(LastCol) <> colSAmount Then
                tblSales.col = colSAmount
            End If
        End If
    End If
End Sub

Private Sub tblSales_SelChange(Cancel As Integer)
    tgsSales.SelChange Cancel
    Cancel = True
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
    Dim sAmountPattern
    Dim i As Long
    
    If sSalesTypeCode = sRatio Then
        sAmountPattern = tfnDecimalPattern(10, 2, True)
    Else
        sAmountPattern = tfnDecimalPattern(10, 2)
    End If
    
    tgmSales.SetPattern colSAmount, sAmountPattern
    
    strSQL = fnGetSalesSQL()
    
    If GetRecordSet(rsTemp, strSQL, , SUB_NAME) < 0 Then
        fnLoadSales = "Failed to access database to load the sales record"
        Exit Function
    End If
    
    If rsTemp.RecordCount = 0 Then
        
        If t_nFormMode = ADD_MODE Then
            If MsgBox("Sales record not available for the From Date and To Date. " _
               + "Do you want to continue?", vbQuestion + vbYesNo) = vbNo Then
                fnLoadSales = "No Sales record available to Add"
            End If
        Else
            fnLoadSales = "No Sales record available to Edit"
        End If
        
        Exit Function
    End If
    
    tgmSales.FillWithRecordset rsTemp, , True
    
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
        sSalesType = sSalesTypeCode
    End If
    
    Select Case sSalesType
        Case sBiWeek, sOneMth
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
        Case sRatio
            strSQL = "SELECT prft_ctr, prft_name " ',  0.00 as amount"
            strSQL = strSQL & " FROM sys_prft_ctr"
            strSQL = strSQL & " WHERE prft_type IN ('R', 'B')"
            strSQL = strSQL & " ORDER BY prft_ctr"
        Case "EDIT_MODE"
            If txtBox Is Nothing Then  'SQL for populating the sales grid in edit
                strSQL = "SELECT prft_ctr, prft_name, bs_sales_amount as amount,"
                strSQL = strSQL & " bs_to_date as to_date, bs_from_date as from_date"
                strSQL = strSQL & " FROM bonus_sales, sys_prft_ctr "
                strSQL = strSQL & " WHERE bs_prft_ctr = prft_ctr"
                strSQL = strSQL & " AND bs_sales_type = " & tfnSQLString(sSalesTypeCode)
                strSQL = strSQL & " AND bs_from_date = " & tfnDateString(txtFromDate, True)
                strSQL = strSQL & " AND bs_to_date = " & tfnDateString(txtToDate, True)
                strSQL = strSQL & " ORDER BY prft_ctr"
            Else
                If txtBox Is txtFromDate Then  'From Date dropdown SQL
                    strSQL = "SELECT bs_from_date, bs_to_date"
                    strSQL = strSQL & " FROM bonus_sales"
                    strSQL = strSQL & " WHERE bs_sales_type = " & tfnSQLString(sSalesTypeCode)
                    If IsValidDate(txtToDate) Then
                        strSQL = strSQL & " AND bs_to_date = " & tfnDateString(txtToDate, True)
                    End If
                    strSQL = strSQL & " GROUP BY bs_from_date, bs_to_date"
                Else  'To Date dropdown SQL
                    strSQL = "SELECT bs_from_date, bs_to_date"
                    strSQL = strSQL & " FROM bonus_sales"
                    strSQL = strSQL & " WHERE bs_sales_type = " & tfnSQLString(sSalesTypeCode)
                    
                    If IsValidDate(txtFromDate) Then
                        strSQL = strSQL & " AND bs_from_date = " & tfnDateString(txtFromDate, True)
                    End If
                    
                    strSQL = strSQL & " GROUP BY bs_from_date, bs_to_date"
                End If
            End If
    End Select
    fnGetSalesSQL = strSQL
    
End Function

Private Function fnCheckSales(sFreq As String) As String
    Const SUB_NAME As String = "fnCheckSales"
    Dim strSQL As String
    Dim rsTemp As Recordset
    
    bSalesRecordExist = False
    
    'check from date
    strSQL = "SELECT COUNT(bs_from_date) AS cnt_date"
    strSQL = strSQL & " FROM bonus_sales"
    strSQL = strSQL & " WHERE bs_sales_type = " & tfnSQLString(sFreq)
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
        strSQL = strSQL & " WHERE bs_sales_type = " & tfnSQLString(sFreq)
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
        strSQL = strSQL & " WHERE bs_sales_type = " & tfnSQLString(sFreq)
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
        strSQL = strSQL & " WHERE bs_sales_type = " & tfnSQLString(sFreq)
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
        If MsgBox("Sales record(s) already exist for From Date/To Date and will be replaced." _
           + " Are you sure you want to replace the existing sales record?", _
           vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
            bSalesRecordExist = False
            fnCheckSales = "Sales record(s) already exist for From Date/To Date"
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

Private Sub txtEmployee_Change()
    cValidDetail.Change txtEmployee
    tgcDropdown.Change txtEmployee
End Sub

Private Sub txtEmployee_GotFocus()
    tgcDropdown.GotFocus txtEmployee
    cValidDetail.GotFocus txtEmployee
    
    If tgcDropdown.SingleRecordSelected Then
        If cValidDetail.FirstInvalidInput < 0 Then
            subEnterBonusPhaseII
        Else
            subSetFocus txtDPrftCtr
        End If
    End If
    
    Screen.MousePointer = vbDefault
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
                If cValidDetail.FirstInvalidInput < 0 Then
                    subEnterBonusPhaseII
                Else
                    subSetFocus txtDPrftCtr
                End If
                Screen.MousePointer = vbDefault
          End If
          KeyAscii = 0
       End If
    Else
        cValidDetail.Keypress txtEmployee, KeyAscii
    End If

End Sub

Private Sub txtEmployee_LostFocus()
    tgcDropdown.LostFocus txtEmployee
    If cValidDetail.LostFocus(txtEmployee, cmdEmployee, txtEmpName, cmdEmpName) Then
        
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub txtEmpName_Change()
    cValidDetail.Change txtEmpName
    tgcDropdown.Change txtEmpName
End Sub

Private Sub txtEmpName_GotFocus()
    tgcDropdown.GotFocus txtEmpName
    cValidDetail.GotFocus txtEmpName
    Screen.MousePointer = vbDefault
    
    If tgcDropdown.SingleRecordSelected Then
        If cValidDetail.FirstInvalidInput < 0 Then
            subEnterBonusPhaseII
        Else
            subSetFocus txtDPrftCtr
        End If
    End If

    Screen.MousePointer = vbDefault
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
                If cValidDetail.FirstInvalidInput < 0 Then
                    subEnterBonusPhaseII
                Else
                    subSetFocus txtDPrftCtr
                End If
                
                Screen.MousePointer = vbDefault
            End If
            KeyAscii = 0
        End If
    Else
        cValidDetail.Keypress txtEmpName, KeyAscii
    End If

End Sub

Private Sub txtEmpName_LostFocus()
    tgcDropdown.LostFocus txtEmpName
    If cValidDetail.LostFocus(txtEmployee, cmdEmployee, txtEmpName, cmdEmpName) Then
        'MsgBox ""
    End If
    Screen.MousePointer = vbDefault
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
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Sub txtDPrftCtr_Change()
    cValidDetail.Change txtDPrftCtr
    tgcDropdown.Change txtDPrftCtr
End Sub

Private Sub txtDPrftCtr_GotFocus()
    tgcDropdown.GotFocus txtDPrftCtr
    cValidDetail.GotFocus txtDPrftCtr
    
    If tgcDropdown.SingleRecordSelected Then
        If cValidDetail.FirstInvalidInput < 0 Then
            subEnterBonusPhaseII
        End If
    End If
    
    Screen.MousePointer = vbDefault
End Sub

Private Sub txtDPrftCtr_KeyPress(KeyAscii As Integer)
    Dim bKeyCode As Boolean
    
    If KeyAscii = vbKeyReturn Then
        tgcDropdown.ComboSQL(txtDPrftCtr) = fnSetComboSQL(txtDPrftCtr.TabIndex)
        Screen.MousePointer = vbHourglass
    End If
        
    bKeyCode = tgcDropdown.Keypress(txtDPrftCtr, KeyAscii)
    Screen.MousePointer = vbDefault
    
    If Not bKeyCode Then
       If KeyAscii = vbKeyReturn Then
          If tgcDropdown.SingleRecordSelected Then
                If cValidDetail.FirstInvalidInput < 0 Then
                    subEnterBonusPhaseII
                End If
                Screen.MousePointer = vbDefault
          End If
          KeyAscii = 0
       End If
    Else
        cValidDetail.Keypress txtDPrftCtr, KeyAscii
    End If

End Sub

Private Sub txtDPrftCtr_LostFocus()
    tgcDropdown.LostFocus txtDPrftCtr
    If cValidDetail.LostFocus(txtDPrftCtr, cmdDPrftCtr, txtDPrftCtrName, cmdDPrftCtrName) Then
        If cValidDetail.FirstInvalidInput < 0 Then
            If ActiveControl Is tblDetails Then
                subEnterBonusPhaseII
            End If
        End If
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub txtDPrftCtrName_Change()
    cValidDetail.Change txtDPrftCtrName
    tgcDropdown.Change txtDPrftCtrName
End Sub

Private Sub txtDPrftCtrName_GotFocus()
    tgcDropdown.GotFocus txtDPrftCtrName
    cValidDetail.GotFocus txtDPrftCtrName
    Screen.MousePointer = vbDefault
    
    If tgcDropdown.SingleRecordSelected Then
        If cValidDetail.FirstInvalidInput < 0 Then
            subEnterBonusPhaseII
        End If
    End If

    Screen.MousePointer = vbDefault
End Sub

Private Sub txtDPrftCtrName_KeyPress(KeyAscii As Integer)
    Dim bKeyCode As Boolean
    
    If KeyAscii = vbKeyReturn Then
        tgcDropdown.ComboSQL(txtDPrftCtrName) = fnSetComboSQL(txtDPrftCtrName.TabIndex)
        Screen.MousePointer = vbHourglass
    End If
    
    bKeyCode = tgcDropdown.Keypress(txtDPrftCtrName, KeyAscii)
    Screen.MousePointer = vbDefault
    
    If Not bKeyCode Then
        If KeyAscii = vbKeyReturn Then
            If tgcDropdown.SingleRecordSelected Then
                If cValidDetail.FirstInvalidInput < 0 Then
                    subEnterBonusPhaseII
                End If
                
                Screen.MousePointer = vbDefault
            End If
            KeyAscii = 0
        End If
    Else
        cValidDetail.Keypress txtDPrftCtrName, KeyAscii
    End If

End Sub

Private Sub txtDPrftCtrName_LostFocus()
    tgcDropdown.LostFocus txtDPrftCtrName
    If cValidDetail.LostFocus(txtDPrftCtr, cmdDPrftCtr, txtDPrftCtrName, cmdDPrftCtrName) Then
        If cValidDetail.FirstInvalidInput < 0 Then
            If ActiveControl Is tblDetails Then
                subEnterBonusPhaseII
            End If
        End If
    End If
    
    If ActiveControl Is tblDetails Then
        If tgmDetail.RowCount <= 0 Then
            SendKeys "{TAB}", True
        End If
    End If
    
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdDPrftCtr_Click()
    Screen.MousePointer = vbHourglass
    tgcDropdown.ComboSQL(txtDPrftCtr) = fnSetComboSQL(txtDPrftCtr.TabIndex)
    tgcDropdown.Click cmdDPrftCtr
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdDPrftCtrName_Click()
    tgcDropdown.ComboSQL(txtDPrftCtrName) = fnSetComboSQL(txtDPrftCtrName.TabIndex)
    tgcDropdown.Click cmdDPrftCtrName
End Sub

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

Private Sub txtEmployeeName_LostFocus()
    objHours.txtEmployeeName_LostFocus
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

Private Sub tblTimeCard_BeforeColEdit(ByVal ColIndex As Integer, ByVal KeyAscii As Integer, Cancel As Integer)
    objHours.tblTimeCard_BeforeColEdit ColIndex, KeyAscii, Cancel
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

Private Sub tblTimeCard_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    objHours.tblTimeCard_MouseDown Button, Shift, x, y
End Sub

Private Sub tblTimeCard_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    objHours.tblTimeCard_RowColChange LastRow, LastCol
End Sub

Private Sub tblTimeCard_SelChange(Cancel As Integer)
    Cancel = True
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
    
'    Dim nMM As Integer
'    Dim nDD As Integer
'    Dim nYY As Integer
'
    If nDataStatus = DATA_CHANGED Then
        Exit Sub
    End If
    
    If Not cValidate.ValidInput(txtStartDate) Then
        Exit Sub
    End If
    
'    nDD = 1
'    nMM = Month(tfnDateString(Date))
'    nYY = Year(tfnDateString(Date))
'
'    'set start date to first day of the month
'    txtStartDate = tfnFormatDate(Format(nMM, "00") + "/" + _
'        Format(nDD, "00") + "/" + Format(nYY, "0000"))
'
    If Trim(txtFrequency) = "" Then
        txtFrequency = "P"
    End If
    
    If Trim(txtEndDate) = "" Then
        txtEndDate = fnGetProposedEndDate(txtStartDate, txtFrequency)
    End If
    
    cValidate.ResetFlags
    cmdProcess.Enabled = cValidate.FirstInvalidInput < 0
End Sub

Private Function fnCheckFrequency(sStartDate As String, _
                                  sEndDate As String, _
                                  sFrequency As String, Optional bShowMsg As Boolean = True) As String
    Dim sDate As String
    
    If sFrequency = sGas Then
        Exit Function
    End If
    
    sDate = fnGetProposedEndDate(sStartDate, sFrequency)
    
    If CDate(sEndDate) <> CDate(sDate) Then
        
        If bShowMsg Then
            
            If MsgBox("For Frequency " + tfnSQLString(sFrequency) + ", the Ending Date " _
               + tfnDateString(sEndDate, True) + " is different from the system proposed Ending Date " _
               + tfnDateString(sDate, True) + ". Are you sure you want to override the system " _
               + "Ending Date?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
                txtEndDate = sDate
                'fnCheckFrequency = "Ending Date entered is not same as system proposed Ending Date"
            End If
            
        Else
            txtEndDate = sDate
        End If
        
    End If
    
End Function

