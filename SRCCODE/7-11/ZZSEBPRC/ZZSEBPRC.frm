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
      TabIndex        =   58
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
      TabIndex        =   60
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
         TabIndex        =   62
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
            Begin FACTFRMLib.FactorFrame efraBaseIIDetail 
               Height          =   4152
               Left            =   60
               TabIndex        =   68
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
                  TabIndex        =   51
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
                  TabIndex        =   53
                  Tag             =   "pn_name"
                  Top             =   276
                  Width           =   5868
               End
               Begin DBTrueGrid.TDBGrid tblDetails 
                  Height          =   3408
                  HelpContextID   =   579
                  Left            =   60
                  OleObjectBlob   =   "ZZSEBPRC.frx":0000
                  TabIndex        =   55
                  Top             =   696
                  Width           =   8556
               End
               Begin FACTFRMLib.FactorFrame cmdEmployee 
                  Height          =   360
                  HelpContextID   =   576
                  Left            =   1944
                  TabIndex        =   52
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
                  TabIndex        =   54
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
                  TabIndex        =   70
                  Top             =   36
                  Width           =   1836
               End
               Begin VB.Label lblEmpName 
                  Caption         =   "Employee Name"
                  Height          =   252
                  Left            =   2364
                  TabIndex        =   69
                  Top             =   36
                  Width           =   1968
               End
            End
            Begin FACTFRMLib.FactorFrame cmdPrint 
               Height          =   396
               HelpContextID   =   32
               Index           =   3
               Left            =   5988
               TabIndex        =   56
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
               TabIndex        =   57
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
            TabIndex        =   64
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
               TabIndex        =   72
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
               Begin VB.TextBox txtEmpProcess 
                  BackColor       =   &H00FFFFFF&
                  ForeColor       =   &H00000000&
                  Height          =   360
                  HelpContextID   =   532
                  Left            =   60
                  TabIndex        =   37
                  Tag             =   "pn_alt"
                  Top             =   924
                  Width           =   1104
               End
               Begin VB.TextBox txtEmpNameProcess 
                  BackColor       =   &H00FFFFFF&
                  DataSource      =   "datVendor"
                  ForeColor       =   &H00000000&
                  Height          =   360
                  HelpContextID   =   534
                  Left            =   1596
                  TabIndex        =   39
                  Tag             =   "pn_name"
                  Top             =   924
                  Width           =   6624
               End
               Begin VB.ListBox lstProcess 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  ForeColor       =   &H00000000&
                  Height          =   2400
                  HelpContextID   =   536
                  ItemData        =   "ZZSEBPRC.frx":1502
                  Left            =   72
                  List            =   "ZZSEBPRC.frx":1504
                  TabIndex        =   41
                  Top             =   1344
                  Width           =   8532
               End
               Begin VB.TextBox txtPrftCtrName 
                  BackColor       =   &H00FFFFFF&
                  DataSource      =   "datVendor"
                  ForeColor       =   &H00000000&
                  Height          =   360
                  HelpContextID   =   527
                  Left            =   1596
                  TabIndex        =   32
                  Tag             =   "pn_name"
                  Top             =   276
                  Width           =   4176
               End
               Begin VB.TextBox txtPrftCtr 
                  BackColor       =   &H00FFFFFF&
                  ForeColor       =   &H00000000&
                  Height          =   360
                  HelpContextID   =   525
                  Left            =   60
                  TabIndex        =   30
                  Tag             =   "pn_alt"
                  Top             =   276
                  Width           =   1104
               End
               Begin VB.TextBox txtDate 
                  BackColor       =   &H00FFFFFF&
                  ForeColor       =   &H00000000&
                  Height          =   360
                  HelpContextID   =   529
                  Left            =   6204
                  TabIndex        =   34
                  Tag             =   "pn_alt"
                  Top             =   276
                  Width           =   1212
               End
               Begin VB.TextBox txtFrequency 
                  BackColor       =   &H00FFFFFF&
                  ForeColor       =   &H00000000&
                  Height          =   360
                  HelpContextID   =   530
                  Left            =   7488
                  TabIndex        =   35
                  Tag             =   "pn_alt"
                  Top             =   276
                  Width           =   732
               End
               Begin MSComctlLib.ProgressBar pbBarMain 
                  Height          =   348
                  Left            =   72
                  TabIndex        =   73
                  Top             =   3756
                  Width           =   8532
                  _ExtentX        =   15050
                  _ExtentY        =   614
                  _Version        =   393216
                  Appearance      =   1
               End
               Begin FACTFRMLib.FactorFrame cmdPrftCtr 
                  Height          =   360
                  HelpContextID   =   526
                  Left            =   1176
                  TabIndex        =   31
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
                  Picture         =   "ZZSEBPRC.frx":1506
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
                  Left            =   5784
                  TabIndex        =   33
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
               Begin FACTFRMLib.FactorFrame cmdFrequency 
                  Height          =   360
                  HelpContextID   =   531
                  Left            =   8232
                  TabIndex        =   36
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
                  Left            =   1176
                  TabIndex        =   38
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
                  TabIndex        =   40
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
                  Left            =   60
                  TabIndex        =   98
                  Top             =   684
                  Width           =   1296
               End
               Begin VB.Label lblEmpNameProcess 
                  Caption         =   "Employee Name"
                  Height          =   252
                  Left            =   1596
                  TabIndex        =   97
                  Top             =   684
                  Width           =   1956
               End
               Begin VB.Label lblPrftCtrName 
                  Caption         =   "Profit Center Name"
                  Height          =   252
                  Left            =   1596
                  TabIndex        =   77
                  Top             =   36
                  Width           =   1956
               End
               Begin VB.Label lblPrftCtr 
                  Caption         =   "Profit Center"
                  Height          =   252
                  Left            =   60
                  TabIndex        =   76
                  Top             =   36
                  Width           =   1296
               End
               Begin VB.Label lblDate 
                  Caption         =   "Date"
                  Height          =   252
                  Left            =   6204
                  TabIndex        =   75
                  Top             =   36
                  Width           =   1032
               End
               Begin VB.Label lblFrequency 
                  Caption         =   "Frequency"
                  Height          =   252
                  Left            =   7488
                  TabIndex        =   74
                  Top             =   36
                  Width           =   1032
               End
            End
            Begin FACTFRMLib.FactorFrame cmdProcess 
               Height          =   396
               HelpContextID   =   537
               Left            =   5988
               TabIndex        =   42
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
               TabIndex        =   44
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
               TabIndex        =   43
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
            TabIndex        =   63
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
               TabIndex        =   78
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
                  TabIndex        =   80
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
                     TabIndex        =   81
                     TabStop         =   0   'False
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
                        Index           =   4
                        Left            =   96
                        ScaleHeight     =   216
                        ScaleWidth      =   228
                        TabIndex        =   94
                        Top             =   720
                        Visible         =   0   'False
                        Width           =   255
                     End
                     Begin VB.TextBox txtHEmpName 
                        BackColor       =   &H00FFFFFF&
                        DataSource      =   "datVendor"
                        ForeColor       =   &H00000000&
                        Height          =   360
                        HelpContextID   =   515
                        Left            =   2028
                        TabIndex        =   24
                        Tag             =   "pn_name"
                        Top             =   276
                        Width           =   3576
                     End
                     Begin VB.TextBox txtHEmployee 
                        BackColor       =   &H00FFFFFF&
                        ForeColor       =   &H00000000&
                        Height          =   360
                        HelpContextID   =   513
                        Left            =   72
                        TabIndex        =   22
                        Tag             =   "pn_alt"
                        Top             =   276
                        Width           =   1524
                     End
                     Begin VB.TextBox txtSSN 
                        Height          =   360
                        HelpContextID   =   517
                        Left            =   6036
                        TabIndex        =   26
                        Top             =   276
                        Width           =   1896
                     End
                     Begin FACTFRMLib.FactorFrame cmdHEmployee 
                        Height          =   360
                        HelpContextID   =   514
                        Left            =   1608
                        TabIndex        =   23
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
                     Begin FACTFRMLib.FactorFrame cmdHEmpName 
                        Height          =   360
                        HelpContextID   =   516
                        Left            =   5616
                        TabIndex        =   25
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
                        TabIndex        =   27
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
                     Begin DBTrueGrid.TDBGrid tblHours 
                        Height          =   3036
                        HelpContextID   =   519
                        Left            =   72
                        OleObjectBlob   =   "ZZSEBPRC.frx":1D96
                        TabIndex        =   28
                        Top             =   696
                        Width           =   5892
                     End
                     Begin DBTrueGrid.TDBGrid tblPrftCtr 
                        Height          =   3036
                        HelpContextID   =   520
                        Left            =   6036
                        OleObjectBlob   =   "ZZSEBPRC.frx":29D2
                        TabIndex        =   29
                        Top             =   696
                        Width           =   2268
                     End
                     Begin VB.Label lblPCTotal 
                        Alignment       =   1  'Right Justify
                        BackColor       =   &H00FFFFFF&
                        BorderStyle     =   1  'Fixed Single
                        Height          =   324
                        Left            =   7032
                        TabIndex        =   93
                        Top             =   3780
                        Width           =   1272
                     End
                     Begin VB.Label lblTotalHrs 
                        Alignment       =   1  'Right Justify
                        BackColor       =   &H00FFFFFF&
                        BorderStyle     =   1  'Fixed Single
                        Height          =   324
                        Left            =   4632
                        TabIndex        =   92
                        Top             =   3780
                        Width           =   1332
                     End
                     Begin VB.Label lblTotalHr 
                        Alignment       =   1  'Right Justify
                        Caption         =   "Total Hours/Dollars:"
                        Height          =   252
                        Left            =   2892
                        TabIndex        =   89
                        Top             =   3828
                        Width           =   1692
                     End
                     Begin VB.Label lblSSN 
                        Caption         =   "Social Security Number"
                        Height          =   252
                        Left            =   6036
                        TabIndex        =   88
                        Top             =   24
                        Width           =   2268
                     End
                     Begin VB.Label lblHEmpName 
                        Caption         =   "Employee Name"
                        Height          =   252
                        Left            =   2028
                        TabIndex        =   87
                        Top             =   24
                        Width           =   1968
                     End
                     Begin VB.Label lblEmpNo 
                        Caption         =   "Employee Number"
                        Height          =   252
                        Left            =   72
                        TabIndex        =   86
                        Top             =   24
                        Width           =   1848
                     End
                     Begin VB.Label lblTotalPC 
                        Alignment       =   1  'Right Justify
                        Caption         =   "PC Total:"
                        Height          =   252
                        Left            =   6048
                        TabIndex        =   85
                        Top             =   3828
                        Width           =   888
                     End
                  End
                  Begin FACTFRMLib.FactorFrame cmdAddBtn 
                     Height          =   396
                     HelpContextID   =   10
                     Index           =   4
                     Left            =   36
                     TabIndex        =   16
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
                     TabIndex        =   17
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
                     TabIndex        =   21
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
                     TabIndex        =   18
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
                     Index           =   4
                     Left            =   4296
                     TabIndex        =   19
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
                  TabIndex        =   79
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
                     TabIndex        =   82
                     TabStop         =   0   'False
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
                        TabIndex        =   91
                        Top             =   768
                        Visible         =   0   'False
                        Width           =   255
                     End
                     Begin FACTFRMLib.FactorFrame efraOptSales 
                        Height          =   648
                        Left            =   72
                        TabIndex        =   90
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
                           TabIndex        =   7
                           Top             =   36
                           Width           =   2016
                        End
                        Begin VB.OptionButton optType 
                           Caption         =   "One Month Sales"
                           Height          =   272
                           HelpContextID   =   501
                           Index           =   1
                           Left            =   84
                           TabIndex        =   9
                           Top             =   324
                           Width           =   2016
                        End
                        Begin VB.OptionButton optType 
                           Caption         =   "Gas Sales"
                           Height          =   272
                           HelpContextID   =   503
                           Index           =   2
                           Left            =   2244
                           TabIndex        =   10
                           Top             =   324
                           Width           =   2076
                        End
                        Begin VB.OptionButton optType 
                           Caption         =   "Three Month Sales"
                           Height          =   272
                           HelpContextID   =   502
                           Index           =   3
                           Left            =   2244
                           TabIndex        =   8
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
                        TabIndex        =   11
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
                        TabIndex        =   13
                        Tag             =   "pn_name"
                        Top             =   348
                        Width           =   1488
                     End
                     Begin DBTrueGrid.TDBGrid tblSales 
                        Height          =   3324
                        HelpContextID   =   508
                        Left            =   72
                        OleObjectBlob   =   "ZZSEBPRC.frx":3610
                        TabIndex        =   15
                        Top             =   768
                        Width           =   8244
                     End
                     Begin FACTFRMLib.FactorFrame cmdFromDate 
                        Height          =   360
                        HelpContextID   =   505
                        Left            =   6024
                        TabIndex        =   12
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
                        Picture         =   "ZZSEBPRC.frx":48EC
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
                        TabIndex        =   14
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
                        Picture         =   "ZZSEBPRC.frx":49FE
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
                        TabIndex        =   84
                        Top             =   96
                        Width           =   1380
                     End
                     Begin VB.Label lblToDate 
                        Caption         =   "To Date"
                        Height          =   252
                        Left            =   6444
                        TabIndex        =   83
                        Top             =   108
                        Width           =   1500
                     End
                  End
                  Begin FACTFRMLib.FactorFrame cmdUpdateInsertBtn 
                     Height          =   396
                     HelpContextID   =   13
                     Index           =   0
                     Left            =   5712
                     TabIndex        =   3
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
                     TabIndex        =   6
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
                     TabIndex        =   5
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
                     TabIndex        =   4
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
            TabIndex        =   66
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
               Begin DBTrueGrid.TDBGrid tblApprove 
                  Height          =   4044
                  HelpContextID   =   550
                  Left            =   60
                  OleObjectBlob   =   "ZZSEBPRC.frx":4B10
                  TabIndex        =   45
                  Top             =   60
                  Width           =   8556
               End
            End
            Begin FACTFRMLib.FactorFrame cmdOk 
               Height          =   396
               HelpContextID   =   16
               Left            =   5988
               TabIndex        =   48
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
               TabIndex        =   46
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
               TabIndex        =   47
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
               TabIndex        =   49
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
         Bindings        =   "ZZSEBPRC.frx":5DEE
         Height          =   2484
         Left            =   108
         OleObjectBlob   =   "ZZSEBPRC.frx":5E0D
         TabIndex        =   65
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
      TabIndex        =   59
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
      TabIndex        =   61
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
      Bindings        =   "ZZSEBPRC.frx":70F1
      Height          =   1008
      Index           =   0
      Left            =   168
      OleObjectBlob   =   "ZZSEBPRC.frx":710B
      TabIndex        =   95
      Top             =   684
      Width           =   2604
   End
   Begin DBTrueGrid.TDBGrid tblDropDown 
      Bindings        =   "ZZSEBPRC.frx":83ED
      Height          =   1008
      Index           =   4
      Left            =   144
      OleObjectBlob   =   "ZZSEBPRC.frx":8407
      TabIndex        =   96
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
Private Const sWeek As String = "1"
Private Const nOneMth As Integer = 1
Private Const sOneMth As String = "2"
Private Const nGas As Integer = 2
Private Const sGas As String = "G"
Private Const nThreeMth As Integer = 3
Private Const sThreeMth As String = "3"

Private tgfDropdown(4) As clsFloatingDropDown

Private bCancelProcess As Boolean
Private cValidate As cValidateInput
Private cValidSls As cValidateInput
Private cValidHrs As cValidateInput
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

Private Sub efraBaseIIProcess_GotFocus()
    If txtPrftCtr.Enabled Then
        subSetFocus txtPrftCtr
    Else
        subSetFocus cmdProcess
    End If
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
    App.HelpFile = "WHOLSALE.HLP"
End Sub

Private Sub Form_Unload(CANCEL As Integer)
    tfnUnlockRow
    
    On Error Resume Next
    
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

Private Sub Form_Load()
    Dim sErrorMessage As String
    
    #If Not PROTOTYPE Then
        'If tfnAuthorizeExecute(Command) = False Then 'Check for handshake if not in the development mode
        '    Unload Me
        '    Exit Sub
        'End If
        
        'open the database, ODBC Dialog Box during developemnt, oleObject Connection String when not
        If Not tfnOpenDatabase(False, sErrorMessage) Then
            subLogErrMsg sErrorMessage
            subLogErrMsg "**System Error: Unable to open Database, Program was terminated"
            Unload Me
            Exit Sub
        End If
        
        'connect to local database
        Set dbLocal = tfnOpenLocalDatabase(False, sErrorMessage)
        If dbLocal Is Nothing Then
            subLogErrMsg sErrorMessage
            subLogErrMsg "**System Error: Unable to open Local Database, Program was terminated"
            Unload Me
            Exit Sub
        End If
    
        If Not fnCreateSearchTable("prm_empno", "prm_empname") Then
            MsgBox "Failed to create temporary employee table", vbCritical
            Unload Me
            End
        End If
    #End If
    
    subSetExitCancelBtn "EXIT"
    Screen.MousePointer = vbHourglass
    frmContext.ButtonEnabled(CANCEL_UP) = True
    mnuCancel.Enabled = True
    eTabMain.CurrTab = TabSales
    Me.Enabled = False
    
    subInitErrorHandler   ' Setup Error Control
    subInitSpreadsheets
    subSetFloatingDropDown TabSales
    subSetFloatingDropDown nTabHours
    subSetupCombos
    subInitValidation
    Set clsMath = New clsEquation
    
    #If Not PROTOTYPE Then
        tfnUpdateVersion
    #End If
    
    tfnDisableFormSystemClose Me
    subSetupToolBar
    
    tmrKeyBoard.Enabled = False
    tfnCenterForm Me

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
    
    frmContext.ButtonEnabled(COPY_UP) = False
    frmContext.ButtonEnabled(FO_HOLD_UP) = False
    mnuExit.Enabled = True
    mnuCopy.Enabled = False
    mnuPaste.Enabled = False
    For i = 0 To 3
        optType(i).Value = False
    Next i
    cmdRefresh(Index).Enabled = False
    cmdUpdateInsertBtn(Index).Enabled = False
    cmdDelete(Index).Enabled = False
    cValidate.ResetFlags
    
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
            subEnableControls Index, False
            cmdAddBtn(Index).Enabled = True
            cmdEditBtn(Index).Enabled = True
            eTabMain.TabEnabled(TabProcess) = True
            eTabSub.TabEnabled(TabHours) = True
            subSetFocus cmdAddBtn(Index)
        Case nTabHours
            If nDataStatus = DATA_CHANGED Then
                If Not tfnCancelExit(t_szCANCEL_MESSAGE) Then
                    Screen.MousePointer = vbDefault
                    Exit Sub
                End If
            End If
            nDataStatus = DATA_INIT
            txtHEmployee = ""
            txtHEmpName = ""
            txtSSN = ""
            cValidHrs.ResetFlags
            tgmHours.ClearData
            tgmPrftCtr.ClearData
            lblPCTotal = ""
            lblTotalHrs = ""
            subEnableControls Index, False
            cmdAddBtn(Index).Enabled = True
            cmdEditBtn(Index).Enabled = True
            eTabSub.TabEnabled(TabSales) = True
            eTabMain.TabEnabled(TabProcess) = True
            subSetFocus cmdAddBtn(Index)
        Case TabProcess
            If nDataStatus = DATA_CHANGED Then
                If Not tfnCancelExit(t_szCANCEL_MESSAGE) Then
                    Screen.MousePointer = vbDefault
                    Exit Sub
                End If
            End If
            nDataStatus = DATA_INIT
            txtPrftCtr = ""
            txtPrftCtrName = ""
            txtDate = CDate(Date)
            txtFrequency = ""
            lstProcess.Clear
            eTabMain.TabEnabled(TabSales) = True
            eTabMain.TabEnabled(TabDetails) = False
            eTabMain.TabEnabled(TabApprove) = False
            subEnablePrint Index, False
            subEnableFirstLineProcess True
            bCancelProcess = False
            cmdProcess.Enabled = True
            subSetProgress 0
            subSetFocus txtPrftCtr
        Case TabApprove
            If nDataStatus = DATA_CHANGED Then
                If tfnCancelExit(t_szCANCEL_MESSAGE) Then
                    tgmApprove.ClearData
                    nDataStatus = DATA_INIT
                    fnFillApproveTab vArrBonus
                    subSetFocus tblApprove
                End If
            End If
        Case TabDetails
            txtEmployee = ""
            txtEmpName = ""
            tgmDetail.ClearData
            subEnableEmployee True
            tblDetails.Enabled = False
            subEnablePrint Index, False
            subSetFocus txtEmployee
    End Select
    
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

Private Sub tblApprove_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
    tgsApprove.MouseUp Button, Shift, y
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

Private Sub tblPrftCtr_UnboundReadData(ByVal RowBuf As DBTrueGrid.RowBuffer, StartLocation As Variant, ByVal ReadPriorRows As Boolean)
    tgmPrftCtr.ReadData RowBuf, StartLocation, ReadPriorRows
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
            .AddButton "Add Commission &Code", PRDCLS_UP
            .AddButton "Add Co&mmission Formula", SYS_LOCKS_UP
            .AddButton "&Launch Commssion Master", PROFITCENTER_UP
            .AddButton "&Export to Payroll", PRAPPROV_UP
            .AddButton "View Formula Details", FO_HOLD_UP, , True
        .EndSetupToolbar
    
        .HelpFile = szHelpFileName
    End With
End Sub

Public Sub TBButtonCallBack(ByVal nID As Integer)
    Select Case nID
        Case CANCEL_UP
            subCancel eTabMain.CurrTab
        Case EXIT_UP
            subExit
        Case FO_HOLD_UP  'Approve PR
            subShowFormulaDetails
        Case PRINT_UP
            subPrint eTabMain.CurrTab
    End Select
End Sub

Private Sub mnuModules_Click(Index As Integer)
    frmContext.MenuClick Index
End Sub

Private Sub tbToolbar_ButtonClick(ByVal Button As Button)
    frmContext.ButtonClick Button
End Sub

Private Sub tbToolbar_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
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
        pbBarMain.Visible = False
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
    
    ReDim vArrBonus(7, 0)
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
    
    strSQL = "SELECT bm_empno, bc_type, bc_bonus_code, bm_sequence, bf_level"
    strSQL = strSQL & " FROM bonus_master, bonus_codes, bonus_formula"
    strSQL = strSQL & " WHERE bm_bonus_code = bc_bonus_code"
    strSQL = strSQL & " AND bm_bonus_code = bf_bonus_code"
    If txtPrftCtr <> "" Then
        strSQL = strSQL & " AND bm_eligible_pc = " & Trim(txtPrftCtr)
    End If
    If txtFrequency <> "" Then
        strSQL = strSQL & " AND bc_frequency = " & tfnSQLString(Trim(txtFrequency))
    End If
    If txtDate <> "" Then
        strSQL = strSQL & " AND bm_eligible_date <= " & tfnDateString(txtDate, True)
        strSQL = strSQL & " AND bm_stop_date >= " & tfnDateString(txtDate, True)
    Else
        strSQL = strSQL & " AND bm_eligible_date <= " & tfnDateString(Date, True)
        strSQL = strSQL & " AND bm_stop_date >= " & tfnDateString(Date, True)
    End If
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
                ReDim Preserve vArrBonus(7, nSize)
                vArrBonus(colAEmpNo, nSize) = Trim(sEmpNo)
                vArrBonus(colAEmpName, nSize) = fnGetEmployeeName(sEmpNo)
                vArrBonus(colADate, nSize) = txtDate
                vArrBonus(colABonusAmt, nSize) = Format(dTotalBonus, "##,##0.00")
                vArrBonus(colAHdnBAmtLvls, nSize) = Trim(sAmtAllBCodes) 'Hidden Column
            End If
            dBLvlAmt = fnGetBonusAmount(rsTemp!bm_empno, rsTemp!bc_type, rsTemp!bc_bonus_code, rsTemp!bf_level)
            dTotalBonus = dBLvlAmt
            sAmtAllBCodes = CStr(dBLvlAmt)
        Else
            dBLvlAmt = fnGetBonusAmount(rsTemp!bm_empno, rsTemp!bc_type, rsTemp!bc_bonus_code, rsTemp!bf_level)
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
            ReDim Preserve vArrBonus(7, nSize)
            vArrBonus(colAEmpNo, nSize) = Trim(sEmpNo)
            vArrBonus(colAEmpName, nSize) = fnGetEmployeeName(sEmpNo)
            vArrBonus(colADate, nSize) = txtDate
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
        Case "tblHours"
            myWidth = Array(0.21, 0.21, 0.16, 0.16, 0.26)
            myField = Array("prh_date", "prh_prft_ctr", "prh_pay_code", "prh_pay_type", "prh_hours")
        Case "tblPrftCtr"
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
            .Width = myWidth(i) * (tbl.Width - 50)
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
        Case "tblHours"
            tbl.Caption = "Time Card Entry"
            tbl.Columns(colHClockIn).Caption = "Clock-In Date"
            tbl.Columns(colHPrftCtr).Caption = "Profit Center"
            tbl.Columns(colHPayCode).Caption = "Pay Code"
            tbl.Columns(colHPayType).Caption = "Pay Type"
            tbl.Columns(colHHrsDol).Caption = "Hours/Dollars"
            tbl.Columns(colHHrsDol).Alignment = vbRightJustify
        Case "tblPrftCtr"
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
    With tgmSales
        .AddEditColumn colSPrftCtr, "Enter Profit Center", szIntegerPattern
        .AddEditColumn colSPrftName, "Enter Profit Center Name", "^P{1,40}"
        .AddEditColumn colSAmount, "Enter Amount", sDecimalString
        .AddEditColumn colSFromDate, "Enter From Date", szDatePattern
        .AddEditColumn colSToDate, "Enter To Date", szDatePattern
        .DisplayFormat(colSAmount) = "###,###,##0.00"
    End With
    
    'Table Time Card Class Implementation
    sDecimalString = tfnDecimalPattern(8, 2)
    subSetGridWidth tblHours
    Set tgmHours = New clsTGSpreadSheet
    Set tgmHours.Table = tblHours
    Set tgmHours.StatusBar = ffraStatusbar ' message bar name
    Set tgmHours.Form = Me
    Set tgmHours.engFactor = t_engFactor
    With tgmHours
        .AddEditColumn colHClockIn, "Enter Clock In Date", szDatePattern
        .AddEditColumn colHPrftCtr, "Enter Profit Center", szIntegerPattern
        .AddEditColumn colHPayCode, "Enter Pay Code", "^P{1,4}$"
        '.AddEditColumn colHPayType, "Enter Pay Type", ""
        .AddEditColumn colHHrsDol, "Enter Hours/Dollars", sDecimalString
        .DisplayFormat(colHHrsDol) = "###,###,##0.00"
        .ColumnForNewRow = 0
        ColHHdnSource = .AddHiddenField("prh_source")
    End With
    
    'Table Sales Class Implementation
    subSetGridWidth tblPrftCtr
    Set tgmPrftCtr = New clsTGSpreadSheet
    Set tgmPrftCtr.Table = tblPrftCtr
    Set tgmPrftCtr.StatusBar = ffraStatusbar ' message bar name
    Set tgmPrftCtr.Form = Me
    Set tgmPrftCtr.engFactor = t_engFactor
    tgmPrftCtr.SetupTable True
    tgmPrftCtr.ClearData
    tgmPrftCtr.DisplayFormat(colPTotal) = "###,###,##0.00"
    tgmPrftCtr.AllowAddNew = False
    
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
    txtEmpName = ""
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
    cValidate.AddEditBox txtPrftCtr, "Enter Profit Center Number"
    cValidate.AddEditBox txtDate, "Enter Eligible Date"
    cValidate.AddEditBox txtFrequency, "Enter Frequency"
    cValidate.AddEditBox txtEmployee, "Enter Employee Number"
    cValidate.MinTabIndex = txtPrftCtr.TabIndex
    cValidate.MaxTabIndex = txtEmpName.TabIndex
    
    'Class implementation for Hours Tab
    Set cValidHrs = New cValidateInput
    Set cValidHrs.Form = Me
    Set cValidHrs.StatusBar = ffraStatusbar
    cValidHrs.AddEditBox txtHEmployee, "Enter Employee Number"
    cValidHrs.AddEditBox txtHEmpName, "Enter Employee Name"
    cValidHrs.AddEditBox txtSSN, "Enter Social Security Number"
    cValidHrs.MinTabIndex = txtHEmployee.TabIndex
    cValidHrs.MaxTabIndex = tblHours.TabIndex
    
    'Class implementation for Sales Tab
    Set cValidSls = New cValidateInput
    Set cValidSls.Form = Me
    Set cValidSls.StatusBar = ffraStatusbar
    cValidSls.AddEditBox txtFromDate, "Enter From Date"
    cValidSls.AddEditBox txtToDate, "Enter To Date"
    cValidSls.MinTabIndex = optType(nWeek).TabIndex
    cValidSls.MaxTabIndex = tblSales.TabIndex
    
End Sub

Private Function fnSetComboSQL(nTabIndex As Integer) As String
    Dim strSQL As String
    
    Select Case nTabIndex
        Case txtPrftCtr.TabIndex, txtPrftCtrName.TabIndex
            strSQL = "SELECT prft_ctr, prft_name FROM sys_prft_ctr WHERE prft_ctr IN "
            strSQL = strSQL & "(SELECT DISTINCT bm_eligible_pc FROM bonus_master)"
        Case txtFrequency.TabIndex
            strSQL = "SELECT bf_frequency, bf_freq_desc FROM bonus_frequency"
        Case txtEmployee.TabIndex, txtEmpName.TabIndex
            strSQL = "SELECT prm_empno, prm_empname FROM sTmpCustTable"
            strSQL = strSQL & " WHERE prm_empno IN (SELECT bm_empno FROM bonus_master)"
        Case txtHEmployee.TabIndex, txtHEmpName.TabIndex, txtSSN.TabIndex
            strSQL = "SELECT prm_empno, prm_empname, prm_ssn FROM sTmpCustTable "
            strSQL = strSQL & " WHERE prm_empno IN("
            If t_nFormMode = ADD_MODE Then
                strSQL = strSQL & "SELECT DISTINCT prep_empno FROM pr_emp_pay)"
            Else
                strSQL = strSQL & "SELECT DISTINCT prh_empno FROM pr_hours)"
            End If
        Case txtFromDate.TabIndex, txtToDate.TabIndex
            strSQL = fnGetSalesSQL(True, nTabIndex)
    End Select
    fnSetComboSQL = strSQL
End Function

Public Function fnInvalidData(txtBox As Textbox) As Boolean
    #If PROTOTYPE Then
        Exit Function
    #End If

    Select Case txtBox.TabIndex
        Case txtPrftCtr.TabIndex
            fnInvalidData = Not fnValidPrftCtr(txtBox)
        Case txtDate.TabIndex
            fnInvalidData = Not fnValidDate(txtBox)
        Case txtEmployee.TabIndex, txtHEmployee.TabIndex
            fnInvalidData = Not fnValidEmployee(txtBox)
        Case txtFromDate.TabIndex, txtToDate.TabIndex
            fnInvalidData = Not fnValidSalesDate(txtBox)
        Case txtSSN.TabIndex
            fnInvalidData = False
    End Select
    
End Function

Private Function fnValidEmployee(Box As Textbox) As Boolean
    Const SUB_NAME As String = "fnValidEmployee"
    Dim strSQL As String
    Dim rsTemp As Recordset
    Dim sMsg As String
    Dim clsError As cValidateInput
    
    If eTabMain.CurrTab = TabSales Then
        Set clsError = cValidHrs
    Else 'Tab Details...
        Set clsError = cValidate
    End If
    
    fnValidEmployee = False
    
    If eTabMain.CurrTab = TabProcess Then
        fnValidEmployee = True
        Exit Function
    End If
    
    If Trim(Box.Text) = "" Then
        clsError.SetErrorMessage Box, "You Must Enter a Employee Number"
        Exit Function
    End If
    
    If Not IsNumeric(Trim(Box.Text)) Then
        clsError.SetErrorMessage Box, "Employee Number does not exist"
        Exit Function
    End If
    
    strSQL = "SELECT prm_empno FROM pr_master WHERE prm_empno IN("
    If clsError Is cValidHrs Then
        If t_nFormMode = ADD_MODE Then
            strSQL = strSQL & " SELECT prep_empno FROM pr_emp_pay GROUP BY prep_empno"
        Else
            strSQL = strSQL & " SELECT prh_empno FROM pr_hours GROUP BY prh_empno"
        End If
    Else
        strSQL = strSQL & " SELECT DISTINCT bm_empno FROM bonus_master"
    End If
    strSQL = strSQL & " ) AND prm_empno = " & Box.Text
    
    If GetRecordSet(rsTemp, strSQL, , SUB_NAME) < 0 Then
        clsError.SetErrorMessage Box, "Failed to access Database"
        Exit Function
    End If
            
    If rsTemp.RecordCount = 0 Then
        clsError.SetErrorMessage Box, "Employee Number does not exist"
        Exit Function
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

Private Sub tblComboDropDown_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
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
        .AddComboBox txtFrequency, cmdFrequency, "bf_frequency", .SQL_STRING_TYPE(1)
        .AddExtraColumn "bf_freq_desc", 1300
        .SetExtend txtFrequency, 2
        
        .AddCombo
        .AddComboBox txtEmployee, cmdEmployee, "prm_empno", .SQL_LONG_TYPE
        .AddComboBox txtEmpName, cmdEmpName, "prm_empname", .SQL_STRING_TYPE(60)
     
        .AddCombo
        .AddComboBox txtHEmployee, cmdHEmployee, "prm_empno", .SQL_LONG_TYPE
        .AddComboBox txtHEmpName, cmdHEmpName, "prm_empname", .SQL_STRING_TYPE(60)
        .AddComboBox txtSSN, cmdSSN, "prm_ssn", .SQL_STRING_TYPE(11)
        
        .AddCombo
        .AddComboBox txtFromDate, cmdFromDate, "bs_from_date", .SQL_STRING_TYPE(10)
        
        .AddCombo
        .AddComboBox txtToDate, cmdToDate, "bs_to_date", .SQL_STRING_TYPE(10)
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
    If txtDate <> "" Then
        strSQL = strSQL & " AND bm_eligible_date <= " & tfnDateString(txtDate, True)
        strSQL = strSQL & " AND bm_stop_date >= " & tfnDateString(txtDate, True)
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
        subSetFocus txtDate
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
                subSetFocus txtDate
                Screen.MousePointer = vbDefault
          End If
          KeyAscii = 0
       End If
    Else
        cValidate.Keypress txtPrftCtr, KeyAscii
    End If

End Sub

Private Sub txtPrftCtr_LostFocus()
    tgcDropdown.LostFocus txtPrftCtr
    cValidate.LostFocus txtPrftCtr, cmdPrftCtr
    
    If cValidate.FirstInvalidInput < 0 Then
        cmdProcess.Enabled = True
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
        subSetFocus txtDate
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
                subSetFocus txtDate
                Screen.MousePointer = vbDefault
            End If
        KeyAscii = 0
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

Private Sub txtFrequency_Change()
    cValidate.Change txtFrequency
    tgcDropdown.Change txtFrequency
    cmdProcess.Enabled = False
End Sub

Private Sub txtFrequency_GotFocus()
    tgcDropdown.GotFocus txtFrequency
    cValidate.GotFocus txtFrequency
    
    If tgcDropdown.SingleRecordSelected Then
        If cmdProcess.Enabled Then
            subSetFocus cmdProcess
        Else
            subSetFocus cmdCancel(TabProcess)
        End If
    End If
    
End Sub

Private Sub txtFrequency_KeyPress(KeyAscii As Integer)
    Dim bKeyCode As Boolean
    
    KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
    
    If KeyAscii = vbKeyReturn Then
        tgcDropdown.ComboSQL(txtFrequency) = fnSetComboSQL(txtFrequency.TabIndex)
        Screen.MousePointer = vbHourglass
    End If
        
    bKeyCode = tgcDropdown.Keypress(txtFrequency, KeyAscii)
    Screen.MousePointer = vbDefault
    
    If Not bKeyCode Then
       If KeyAscii = vbKeyReturn Then
          If tgcDropdown.SingleRecordSelected Then
                If cmdProcess.Enabled Then
                    subSetFocus cmdProcess
                Else
                    subSetFocus cmdCancel(TabProcess)
                End If
                Screen.MousePointer = vbDefault
          End If
          KeyAscii = 0
       End If
    Else
        cValidate.Keypress txtFrequency, KeyAscii
    End If

End Sub

Private Sub txtFrequency_LostFocus()
    tgcDropdown.LostFocus txtFrequency
    cValidate.LostFocus txtFrequency, cmdFrequency
    
    If cValidate.FirstInvalidInput < 0 Then
        cmdProcess.Enabled = True
    End If
    
End Sub

Private Sub txtDate_Change()
    cmdProcess.Enabled = False
    cValidate.Change txtDate
End Sub

Private Sub txtDate_GotFocus()
    cValidate.GotFocus txtDate
    SelectIt txtDate
End Sub

Private Sub txtDate_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        subSetFocus txtFrequency
        KeyAscii = 0
    Else
        tfnRegExpControlKeyPress txtDate, KeyAscii, szDatePattern
        cValidate.Keypress txtDate, KeyAscii
    End If
End Sub

Private Sub txtDate_LostFocus()
    cValidate.LostFocus txtDate
    If cValidate.FirstInvalidInput < 0 Then
        cmdProcess.Enabled = True
    End If
End Sub

Private Sub cmdFrequency_Click()
    tgcDropdown.ComboSQL(txtFrequency) = fnSetComboSQL(txtFrequency.TabIndex)
    tgcDropdown.Click cmdFrequency
End Sub

Private Function fnValidPrftCtr(Box As Textbox) As Boolean
    Const SUB_NAME As String = "fnValidPrftCtr"
    Dim strSQL As String
    Dim rsTemp As Recordset
    Dim sMsg As String
    
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

Private Function fnValidDate(Box As Textbox) As Boolean
    Const SUB_NAME As String = "fnValidDate"
    Dim strSQL As String
    Dim rsTemp As Recordset
    
    fnValidDate = False
    
    If Trim(Box.Text) = "" Then
        fnValidDate = True
        Exit Function
    End If
    
    If Len(Trim(Box)) < 6 Then
        cValidate.SetErrorMessage Box, "Invalid date format"
        Exit Function
    End If
    
    txtDate = CDate(tfnFormatDate(txtDate))
    
    If Not IsDate(Trim(Box.Text)) Then
        cValidate.SetErrorMessage Box, "Invalid date format"
        Exit Function
    End If
    
    fnValidDate = True
    
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
    txtPrftCtr.Enabled = bYesNo
    txtPrftCtrName.Enabled = bYesNo
    txtDate.Enabled = bYesNo
    txtFrequency.Enabled = bYesNo
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
    
    strSQL = "SELECT prm_empname FROM sTmpCustTable WHERE"
    strSQL = strSQL & " prm_empno = " & tfnRound(sEmpNo)
    
    If GetRecordSet(rsTemp, strSQL, , SUB_NAME) < 0 Then
        subLogErrMsg "Failed to access database to get employee name"
        Exit Function
    End If
    
    If rsTemp.RecordCount = 1 Then
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
                Case colSPrftCtr, colSPrftName
                    fnValidCellValue = fnValidGridPrftCtr(sText, nCol, lRow)
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
        Case tblHours.TabIndex
            Select Case nCol
                Case colHClockIn
                    fnValidCellValue = fnValidClockInDate(sText, lRow)
                Case colHPrftCtr
                    fnValidCellValue = fnValidGridPrftCtr(sText, nCol, lRow)
                Case colHPayCode
                    fnValidCellValue = fnValidPayCode(sText, nCol, lRow)
                Case colHPayType
                    fnValidCellValue = True
                Case colHHrsDol
                    If Trim(sText) = "" Then
                        tgmHours.ErrorMessage(nCol) = "You must enter an Hour/Dollar Number"
                        Exit Function
                    Else
                        fnValidCellValue = True
                    End If
            End Select
        Case tblPrftCtr.TabIndex
            fnValidCellValue = True
        Case tblApprove.TabIndex
            fnValidCellValue = True
    End Select
    
End Function

Private Sub cmdAddBtn_Click(Index As Integer)
    #If PROTOTYPE Then
        Exit Sub
    #End If
    
    cmdEditBtn(Index).Enabled = False
    cmdAddBtn(Index).Enabled = False
    cmdUpdateInsertBtn(Index).Caption = t_szCAPTION_INSERT
    t_nFormMode = ADD_MODE
    subEnableFirstLineSlsOrHrs Index, True
    
    If Index = TabSales Then
        tgmSales.AllowAddNew = True
        eTabSub.TabEnabled(TabHours) = False
        eTabMain.TabEnabled(TabProcess) = False
        subSetFocus optType(nWeek)
    Else 'Index is Hours...
        tgmHours.AllowAddNew = True
        eTabMain.TabEnabled(TabProcess) = False
        eTabSub.TabEnabled(TabSales) = False
        subSetFocus txtHEmployee
    End If
    
End Sub

Private Sub cmdDelete_Click(Index As Integer)
    Dim nRow As Integer
    
    Select Case Index
        Case TabSales
            If tgmSales.RowCount < 1 Then
                Exit Sub
            End If
            If t_nFormMode = EDIT_MODE Then
                If Not tfnCancelExit("Are you sure you want to delete the current record?") Then
                    Exit Sub
                End If
                If Not fnDeleteSales(fnGetSalesType, txtToDate, txtFromDate) Then
                    MsgBox "Failed to delete the sales record", vbExclamation
                    Exit Sub
                End If
            End If
            tgmSales.DeleteRow
            If t_nFormMode = EDIT_MODE And tgmSales.RowCount = 0 Then
                tfnResetScreen Index
            End If
        Case nTabHours
            If tgmHours.RowCount < 1 Then
                Exit Sub
            End If
            nRow = tgmHours.GetCurrentRowNumber
            If t_nFormMode = EDIT_MODE Then
                If Not tfnCancelExit("Are you sure you want to delete the current record?") Then
                    Exit Sub
                End If
                If Not fnDeleteHours(txtHEmployee, txtSSN, tgmHours.CellValue(colHPayCode, nRow)) Then
                    MsgBox "Failed to delete the hours record", vbExclamation
                    Exit Sub
                End If
            End If
            tgmHours.DeleteRow
            If t_nFormMode = EDIT_MODE And tgmHours.RowCount = 0 Then
                tfnResetScreen Index
            End If
    End Select
End Sub

Private Sub cmdEditBtn_Click(Index As Integer)
    #If PROTOTYPE Then
        Exit Sub
    #End If
    
    cmdEditBtn(Index).Enabled = False
    cmdAddBtn(Index).Enabled = False
    cmdUpdateInsertBtn(Index).Caption = t_szCAPTION_UPDATE
    t_nFormMode = EDIT_MODE
    subEnableFirstLineSlsOrHrs Index, True
    
    If Index = TabSales Then
        tgmSales.AllowAddNew = False
        eTabSub.TabEnabled(TabHours) = False
        eTabMain.TabEnabled(TabProcess) = False
        subSetFocus optType(nWeek)
    Else 'Index is Hours...
        tgmHours.AllowAddNew = False
        eTabMain.TabEnabled(TabProcess) = False
        eTabSub.TabEnabled(TabSales) = False
        subSetFocus txtHEmployee
    End If
    
End Sub

Private Sub cmdRefresh_Click(Index As Integer)
    Select Case Index
        Case TabSales
            If Not fnLoadSales() Then
                Exit Sub
            End If
        Case nTabHours
            If Not fnLoadHours(txtHEmployee, txtSSN) Then
                Exit Sub
            End If
    End Select
    cmdUpdateInsertBtn(Index).Enabled = False
End Sub

Private Sub cmdUpdateInsertBtn_Click(Index As Integer)
    Dim i As Integer
    Dim lChkLnk As Long
    
    Select Case Index
        Case TabSales
            For i = 0 To tgmSales.RowCount - 1
                If Not fnInsertUpdateSales(tgmSales.CellValue(colSPrftCtr, i), _
                         tgmSales.CellValue(colSFromDate, i), _
                         tgmSales.CellValue(colSToDate, i), _
                         tgmSales.CellValue(colSAmount, i), fnGetSalesType) Then
                    Exit Sub
                End If
            Next i
        Case nTabHours
            For i = 0 To tgmHours.RowCount - 1
                lChkLnk = -1
                If tgmHours.CellValue(colHHrsDol, i) <> 0 Then
                    lChkLnk = 0
                End If
                If t_nFormMode = ADD_MODE Then
                    If Not fnInsertHours(txtHEmployee, txtSSN, _
                                         tgmHours.CellValue(colHClockIn, i), _
                                         tgmHours.CellValue(colHPrftCtr, i), _
                                         tgmHours.CellValue(colHPayCode, i), _
                                         tgmHours.CellValue(colHPayType, i), _
                                         tgmHours.CellValue(colHHrsDol, i), _
                                         0, lChkLnk, "P") Then
                        MsgBox "Failed to insert the employee hours", vbExclamation
                        Exit Sub
                    End If
                Else
                    'Delete the record first and then insert it...
                    If Not fnDeleteHours(txtHEmployee, txtSSN) Then
                        MsgBox "Failed to update employee hours", vbExclamation
                        Exit Sub
                    End If
                    If Not fnInsertHours(txtHEmployee, txtSSN, _
                                         tgmHours.CellValue(colHClockIn, i), _
                                         tgmHours.CellValue(colHPrftCtr, i), _
                                         tgmHours.CellValue(colHPayCode, i), _
                                         tgmHours.CellValue(colHPayType, i), _
                                         tgmHours.CellValue(colHHrsDol, i), 0, lChkLnk, _
                                         tgmHours.CellValue(ColHHdnSource, i)) Then
                        MsgBox "Failed to update employee hours", vbExclamation
                        Exit Sub
                    End If
                End If
            Next i
    End Select
    nDataStatus = DATA_INIT
    tfnResetScreen Index
End Sub

Private Sub subEnableControls(Index As Integer, bYesNo As Boolean)
    subEnableFirstLineSlsOrHrs Index, bYesNo
    If Index = TabSales Then
        tblSales.Enabled = bYesNo
    Else
        tblHours.Enabled = bYesNo
        tblPrftCtr.Enabled = bYesNo
    End If
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

Private Function fnLoadHours(lEmpNo As Long, sSSN As String) As Boolean
    Const SUB_NAME As String = "fnLoadHours"
    Dim strSQL As String
    
    fnLoadHours = False
    If t_nFormMode = ADD_MODE Then
        fnLoadHours = True
        Exit Function
    End If
    
    strSQL = "SELECT * FROM pr_hours WHERE 1=1"
    If tfnRound(lEmpNo) <> 0 Then
        strSQL = strSQL & " AND prh_empno = " & tfnRound(lEmpNo)
    End If
    If Trim(sSSN) <> "" Then
        strSQL = strSQL & " AND prh_ssn = " & tfnSQLString(Trim(sSSN))
    End If
    strSQL = strSQL & " ORDER BY prh_prft_ctr"
    
    tgmHours.FillWithSQL t_dbMainDatabase, strSQL
    
    If tgmHours.RowCount = 0 Then
        MsgBox "No entry found for the selection criteria", vbExclamation
        Exit Function
    End If
    
    fnLoadPrftCtr
    fnLoadHours = True

End Function

Private Function fnGetSalesType() As String
    Dim i As Integer
    Dim sType As String
    
    For i = 0 To 3
        If optType(i).Value Then
            Select Case i
                Case nWeek
                    sType = "1"
                Case nOneMth
                    sType = "2"
                Case nThreeMth
                    sType = "3"
                Case nGas
                    sType = "G"
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
            Case nTabHours
                Set .MainTable = tblHours
                Set .EditClass = tgmHours
                .AddDropDown 1
                .AddColumn colHPrftCtr, "prft_ctr", .COLUMN_TYPE_INTEGER
                .AddExtraColumn "Profit Center Name", "prft_name", 2500
                
                .AddDropDown 2
                .AddColumn colHPayCode, "prep_pay_code", .COLUMN_TYPE_STRING
                .AddExtraColumn "Pay Description", "prpa_desc", 2500
        End Select
    End With
    
End Sub

Private Sub subSetFloatingSQL(Index As Integer)
    Dim strSQL As String
    Dim nYear As Integer
    
    Select Case Index
        Case TabSales
            strSQL = "SELECT prft_ctr, prft_name FROM sys_prft_ctr"
            tgfDropdown(Index).SetSQL colSPrftCtr, strSQL
        Case nTabHours
            If tblHours.col = colHPrftCtr Then
                strSQL = "SELECT prft_ctr, prft_name FROM sys_prft_ctr"
                tgfDropdown(Index).SetSQL colHPrftCtr, strSQL
            Else 'colHPayCode
                On Error Resume Next
                nYear = Year(CDate(tgmHours.CellValue(colHClockIn, tgmHours.GetCurrentRowNumber)))
                strSQL = "SELECT distinct prep_pay_code, prpa_desc FROM pr_emp_pay,pr_pay "
                If nYear > 0 Then
                    strSQL = strSQL & " WHERE prep_year = " & tfnSQLString(nYear)
                End If
                strSQL = strSQL & " AND prpa_pay_code = prep_pay_code"
                strSQL = strSQL & " AND (prep_active IS NULL OR prep_active <> 'N')"
                strSQL = strSQL & " AND ((prpa_type='N' AND prpa_calc_method='D')"
                strSQL = strSQL & " OR   (prpa_type='P' AND prpa_calc_method='H')) "
                strSQL = strSQL & " AND prep_empno = " & tfnRound(txtHEmployee)
                tgfDropdown(Index).SetSQL colHPayCode, strSQL
            End If
    End Select
    
End Sub

Private Sub tblDropDown_Click(Index As Integer)
    tgfDropdown(Index).TableClick tblDropDown(Index)
End Sub

Private Sub tblDropDown_KeyPress(Index As Integer, KeyAscii As Integer)
    tgfDropdown(Index).Keypress tblDropDown(Index), KeyAscii
End Sub

Private Sub tblDropDown_LostFocus(Index As Integer)
    tgfDropdown(Index).LostFocus tblDropDown(Index)
End Sub

Private Sub tblDropDown_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, y As Single)
    tgfDropdown(Index).MouseUp y
End Sub

Private Sub tblDropDown_RowColChange(Index As Integer, LastRow As Variant, ByVal LastCol As Integer)
    tgfDropdown(Index).RowColChange tblDropDown(Index)
End Sub

Private Sub txtHEmployee_Change()
    cValidHrs.Change txtHEmployee
    tgcDropdown.Change txtHEmployee
    txtHEmpName = ""
End Sub

Private Sub txtHEmployee_GotFocus()
    tgcDropdown.GotFocus txtHEmployee
    cValidHrs.GotFocus txtHEmployee
    
    If tgcDropdown.SingleRecordSelected Then
        subEnterPhaseIISlsOrHrs nTabHours
        Screen.MousePointer = vbDefault
    End If
    
End Sub

Private Sub txtHEmployee_KeyPress(KeyAscii As Integer)
    Dim bKeyCode As Boolean
    
    If KeyAscii = vbKeyReturn Then
        tgcDropdown.ComboSQL(txtHEmployee) = fnSetComboSQL(txtHEmployee.TabIndex)
        Screen.MousePointer = vbHourglass
    End If
        
    bKeyCode = tgcDropdown.Keypress(txtHEmployee, KeyAscii)
    Screen.MousePointer = vbDefault
    
    If Not bKeyCode Then
       If KeyAscii = vbKeyReturn Then
          If tgcDropdown.SingleRecordSelected Then
                subEnterPhaseIISlsOrHrs nTabHours
                Screen.MousePointer = vbDefault
          End If
          KeyAscii = 0
       End If
    Else
        cValidHrs.Keypress txtHEmployee, KeyAscii
    End If

End Sub

Private Sub txtHEmployee_LostFocus()
    tgcDropdown.LostFocus txtHEmployee
    If cValidHrs.LostFocus(txtSSN, cmdEmployee, txtHEmployee, cmdHEmployee, txtHEmpName, cmdHEmpName, tblComboDropdown) Then
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub txtHEmpName_Change()
    tgcDropdown.Change txtHEmpName
End Sub

Private Sub txtHEmpName_GotFocus()
    tfnSetStatusBarMessage "Enter Employee Name"
    tgcDropdown.GotFocus txtHEmpName
    Screen.MousePointer = vbDefault
    
    If tgcDropdown.SingleRecordSelected Then
        subEnterPhaseIISlsOrHrs nTabHours
        Screen.MousePointer = vbDefault
    End If
End Sub

Private Sub txtHEmpName_KeyPress(KeyAscii As Integer)
    Dim bKeyCode As Boolean
    
    If KeyAscii = vbKeyReturn Then
        tgcDropdown.ComboSQL(txtHEmpName) = fnSetComboSQL(txtHEmpName.TabIndex)
        Screen.MousePointer = vbHourglass
    End If
    
    bKeyCode = tgcDropdown.Keypress(txtHEmpName, KeyAscii)
    Screen.MousePointer = vbDefault
    
    If Not bKeyCode Then
        If KeyAscii = vbKeyReturn Then
            If tgcDropdown.SingleRecordSelected Then
                subEnterPhaseIISlsOrHrs nTabHours
                Screen.MousePointer = vbDefault
            End If
        KeyAscii = 0
        End If
    End If

End Sub
Private Sub cmdHEmployee_Click()
    Screen.MousePointer = vbHourglass
    tgcDropdown.ComboSQL(txtHEmployee) = fnSetComboSQL(txtHEmployee.TabIndex)
    tgcDropdown.Click cmdHEmployee
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdHEmpName_Click()
    tgcDropdown.ComboSQL(txtHEmpName) = fnSetComboSQL(txtHEmpName.TabIndex)
    tgcDropdown.Click cmdHEmpName
End Sub

Private Sub subEnterPhaseIISlsOrHrs(Index As Integer)
    
    If Index = nTabHours Then
        If cValidHrs.FirstInvalidInput >= 0 Then
            subSetFocus cmdCancel(Index)
            Exit Sub
        End If
    ElseIf Index = TabSales Then
        If cValidSls.FirstInvalidInput >= 0 Then
            subSetFocus cmdCancel(Index)
            Exit Sub
        End If
    End If

    Screen.MousePointer = vbHourglass
    subEnableControls Index, True
    subEnableFirstLineSlsOrHrs Index, False
    
    If Index = TabSales Then
        If Not fnLoadSales() Then
            cmdCancel_Click (TabSales)
            Exit Sub
        End If
    Else ' nTabHours
        If Not fnLoadHours(txtHEmployee, txtSSN) Then
            cmdCancel_Click (nTabHours)
            Exit Sub
        End If
    End If
        
    If Index = TabSales Then
        subSetFocus tblSales
    Else
        subSetFocus tblHours
    End If
    Screen.MousePointer = vbDefault
    
End Sub

Private Sub subEnableFirstLineSlsOrHrs(Index As Integer, bYesNo As Boolean)
    Select Case Index
        Case TabSales
            txtFromDate.Enabled = bYesNo
            subEnableSearchbtn cmdFromDate, bYesNo
            txtToDate.Enabled = bYesNo
            subEnableSearchbtn cmdToDate, bYesNo
            efraOptSales.Enabled = bYesNo
            'Dim i As Integer
            'For i = 0 To 3
            '    optType(i).Enabled = bYesNo
            'Next i
        Case TabHours, nTabHours
            txtHEmployee.Enabled = bYesNo
            txtHEmpName.Enabled = bYesNo
            subEnableSearchbtn cmdHEmployee, bYesNo
            subEnableSearchbtn cmdHEmpName, bYesNo
            txtSSN.Enabled = bYesNo
            subEnableSearchbtn cmdSSN, bYesNo
    End Select
End Sub

Private Sub tblHours_AfterColEdit(ByVal ColIndex As Integer)
    tgmHours.AfterColEdit ColIndex
    Dim lRow As Long: lRow = tgmHours.GetCurrentRowNumber
    
    If tgmHours.RowCount = 0 Then Exit Sub
    
    If tblHours.col = colHClockIn Then
        tgmHours.CellValue(colHClockIn, lRow) = CDate(tfnFormatDate(tgmHours.CellValue(colHClockIn, lRow)))
    End If
    
    If t_nFormMode = EDIT_MODE Then
        If nDataStatus = DATA_CHANGED Then
            subEnableRefreshBtn True, nTabHours
        Else
            subEnableRefreshBtn False, nTabHours
        End If
    End If

End Sub

Private Sub tblHours_BeforeColEdit(ByVal ColIndex As Integer, ByVal KeyAscii As Integer, CANCEL As Integer)
    tgmHours.BeforeColEdit ColIndex, KeyAscii, CANCEL
End Sub

Private Sub tblHours_Change()
    tgmHours.Change
    subEnableUpdateBtn False, nTabHours
    nDataStatus = DATA_CHANGED
End Sub

Private Sub tblHours_FirstRowChange()
    tgmHours.FirstRowChange
    tgfDropdown(nTabHours).FirstRowChange
    
    If tblHours.Row = -1 Then
        subEnableDeleteBtn False, nTabHours
    End If
    
End Sub

Private Sub tblHours_GotFocus()
    tfnSetStatusBarMessage "Time Card Entry"
    tgmHours.GotFocus
    tgfDropdown(nTabHours).GotFocus
End Sub

Private Sub tblHours_KeyDown(KeyCode As Integer, Shift As Integer)
    tgmHours.KeyDown KeyCode, Shift
    tgfDropdown(nTabHours).KeyDown tblHours, KeyCode
End Sub

Private Sub tblHours_KeyPress(KeyAscii As Integer)
    Dim lRow As Long
    
    tgfDropdown(nTabHours).Keypress tblHours, KeyAscii
    lRow = tgmHours.GetCurrentRowNumber
    
    If tblHours.col = colHPayCode Then
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If
    
    If Not tgmHours.Keypress(KeyAscii) Then
        KeyAscii = 0
    End If
    
    If t_nFormMode = EDIT_MODE Then
        If KeyAscii = vbKeyReturn And tgmHours.RowCount - 1 = lRow And tblHours.col = colHHrsDol Then
            If cmdUpdateInsertBtn(nTabHours).Enabled Then
                subSetFocus cmdUpdateInsertBtn(nTabHours)
            Else
                subSetFocus cmdCancel(nTabHours)
            End If
        End If
    End If
    
End Sub

Private Sub tblHours_LeftColChange()
    tgfDropdown(nTabHours).LeftColChange
End Sub

Private Sub tblHours_LostFocus()
    tgmHours.LostFocus
    tgfDropdown(nTabHours).LostFocus tblHours
    subSetStdBtn nTabHours, tgmHours
End Sub

Private Sub tblHours_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
    'frmContext.MouseDown Button, EMP_MST_UP
End Sub

Private Sub tblHours_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    Dim lRow As Long
    
    If tgfDropdown(nTabHours).RowColChange(tblHours) Then
        tgmHours.RowColChange LastRow, LastCol
    End If
    
    lRow = tgmHours.GetCurrentRowNumber
    
    If t_nFormMode = ADD_MODE Then
        If tgmHours.CellValue(colHClockIn, lRow) <> "" Then
            subEnableDeleteBtn True, nTabHours
        Else
            subEnableDeleteBtn False, nTabHours
        End If
    Else
        If tgmHours.RowCount > 0 Then
            subEnableDeleteBtn True, nTabHours
        Else
            subEnableDeleteBtn False, nTabHours
        End If
    End If
    
    fnLoadPrftCtr
    subSetStdBtn nTabHours, tgmHours

End Sub

Private Sub tblHours_SelChange(CANCEL As Integer)
    CANCEL = True
End Sub

Private Sub tblHours_UnboundReadData(ByVal RowBuf As DBTrueGrid.RowBuffer, StartLocation As Variant, ByVal ReadPriorRows As Boolean)
    tgmHours.ReadData RowBuf, StartLocation, ReadPriorRows
End Sub

Private Sub subEnableDeleteBtn(bOnOff As Boolean, Index As Integer)
    cmdDelete(Index).Enabled = bOnOff
  '  mnuDelete.Enabled = bOnOff
End Sub

Private Sub subEnableRefreshBtn(bOnOff As Boolean, Index As Integer)
    cmdRefresh(Index).Enabled = bOnOff
   ' mnuRefreshSelect.Enabled = bOnOff
End Sub

Private Sub subEnableUpdateBtn(bOnOff As Boolean, Index As Integer)
    cmdUpdateInsertBtn(Index).Enabled = bOnOff
    'mnuUpdateInsert.Enabled = bOnOff
End Sub

Private Function fnValidGridPrftCtr(sText As String, nCol As Integer, lRow As Long) As Boolean
    Const SUB_NAME As String = "fnValidGridPrftCtr"
    Dim strSQL As String
    Dim rsTemp As Recordset
    
    fnValidGridPrftCtr = False
    
    
    
    fnValidGridPrftCtr = True
    
End Function

Private Function fnValidPayCode(sText As String, nCol As Integer, lRow As Long) As Boolean
    Const SUB_NAME As String = "fnValidPayCode"
    Dim sChkDubInGrid
    Dim strSQL As String
    Dim rsTemp As Recordset
    Dim nYear As Integer
    
    fnValidPayCode = False
    
    If Trim(sText) = "" Then
        tgmHours.ErrorMessage(nCol) = "You must enter a Pay Code"
        Exit Function
    End If
    
    strSQL = "SELECT prpa_pay_code FROM pr_pay WHERE prpa_pay_code = " & tfnSQLString(sText)
    If GetRecordCount(strSQL, , SUB_NAME) <= 0 Then
        tgmHours.ErrorMessage(nCol) = "Pay Code does not exist"
        Exit Function
    End If
    
    nYear = Year(CDate(tgmHours.CellValue(colHClockIn, lRow)))
    If nYear <= 0 Then Exit Function
    
    strSQL = "SELECT DISTINCT prep_pay_code FROM pr_emp_pay,pr_pay "
    strSQL = strSQL & " WHERE prep_year = " & tfnSQLString(nYear)
    strSQL = strSQL & " AND (prep_active IS NULL OR prep_active <> 'N')"
    strSQL = strSQL & " AND prpa_pay_code = prep_pay_code"
    strSQL = strSQL & " AND ((prpa_type='N' AND prpa_calc_method='D')"
    strSQL = strSQL & " OR   (prpa_type='P' AND prpa_calc_method='H')) "
    strSQL = strSQL & " AND prep_empno = " & tfnRound(txtHEmployee)
    
    If GetRecordCount(strSQL, , SUB_NAME) <= 0 Then
        tgmHours.ErrorMessage(nCol) = "This Pay Code is invalid "
        Exit Function
    End If
    
    If Not IsDate(tgmHours.CellValue(colHClockIn, lRow)) Then
        tblHours.col = colHClockIn
        Exit Function
    End If
    
    'Check duplication in table
    Dim nPrftCtr As Integer: nPrftCtr = tgmHours.CellValue(colHPrftCtr, lRow)
    Dim sDate As String:     sDate = tgmHours.CellValue(colHClockIn, lRow)
    
    strSQL = "SELECT prh_empno FROM pr_hours "
    strSQL = strSQL & " WHERE prh_empno = " & tfnRound(txtHEmployee)
    strSQL = strSQL & " AND prh_prft_ctr = " & nPrftCtr
    strSQL = strSQL & " AND prh_date = " & tfnDateString(sDate, True)
    strSQL = strSQL & " AND prh_pay_code = " & tfnSQLString(sText)
    strSQL = strSQL & " AND prh_pay_type = " & tfnSQLString(tgmHours.CellValue(colHPayType, lRow))
    
    If GetRecordCount(strSQL) > 0 And t_nFormMode = ADD_MODE Then
        tgmHours.ErrorMessage(nCol) = "Date-PrftCtr-PayCode already exists in the table"
        Exit Function
    End If

    'Check duplication in the Grid
    Dim Range() As String, i As Integer
    sChkDubInGrid = fnCheckDubInHrsGrid(lRow)
    If sChkDubInGrid <> "" Then
        Range = Split(sChkDubInGrid, ",")
        For i = 0 To UBound(Range)
            If sDate & "~" & CStr(nPrftCtr) & "~" & Trim(sText) = Range(i) Then
                tgmHours.ErrorMessage(nCol) = "Date-PrftCtr-PayCode already exists in the grid"
                Exit Function
            End If
        Next i
    End If
    
    fnValidPayCode = True
    
End Function

Private Sub txtSSN_Change()
    cValidHrs.Change txtSSN
    tgcDropdown.Change txtSSN
End Sub

Private Sub txtSSN_GotFocus()
    tgcDropdown.GotFocus txtSSN
    cValidHrs.GotFocus txtSSN
    
    If tgcDropdown.SingleRecordSelected Then
        subEnterPhaseIISlsOrHrs nTabHours
        Screen.MousePointer = vbDefault
    End If
    
End Sub

Private Sub txtSSN_KeyPress(KeyAscii As Integer)
    Dim bKeyCode As Boolean
    
    If KeyAscii = vbKeyReturn Then
        tgcDropdown.ComboSQL(txtSSN) = fnSetComboSQL(txtSSN.TabIndex)
        Screen.MousePointer = vbHourglass
    End If
        
    bKeyCode = tgcDropdown.Keypress(txtSSN, KeyAscii)
    Screen.MousePointer = vbDefault
    
    If Not bKeyCode Then
       If KeyAscii = vbKeyReturn Then
          If tgcDropdown.SingleRecordSelected Then
                subEnterPhaseIISlsOrHrs nTabHours
                Screen.MousePointer = vbDefault
          End If
          KeyAscii = 0
       End If
    Else
        cValidHrs.Keypress txtSSN, KeyAscii
    End If

End Sub

Private Sub txtSSN_LostFocus()
    tgcDropdown.LostFocus txtSSN
    If cValidHrs.LostFocus(txtSSN, cmdHEmployee, txtHEmployee, cmdHEmployee, txtHEmpName, cmdHEmpName, tblComboDropdown) Then
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdSSN_Click()
    Screen.MousePointer = vbHourglass
    tgcDropdown.ComboSQL(txtSSN) = fnSetComboSQL(txtSSN.TabIndex)
    tgcDropdown.Click cmdSSN
    Screen.MousePointer = vbDefault
End Sub

Private Function fnValidSalesDate(Box As Textbox) As Boolean
    Const SUB_NAME As String = "fnValidSalesDate"
    Dim strSQL As String
    Dim rsTemp As Recordset
    Dim sMsg As String
    
    sMsg = " From Date "
    If Box.Name = "txtToDate" Then
        sMsg = " To Date "
    End If
    
    fnValidSalesDate = False
    
    If Trim(Box.Text) = "" Then
        fnValidSalesDate = True
        Exit Function
    End If
    
    If Not IsDate(tfnFormatDate(Box)) Then
        cValidSls.SetErrorMessage Box, "Invalid Date Format"
        Exit Function
    End If
    
    Box.Text = CDate(tfnFormatDate(Box))
    
    If Box.Name = "txtFromDate" And cValidSls.ValidInput(txtToDate) And txtToDate <> "" Then
        If Box.Text > txtToDate Then
            fnValidSalesDate = False
            cValidate.SetErrorMessage Box, "From Date must be less than To Date"
            Exit Function
        End If
    End If
    
    If Box.Name = "txtToDate" And cValidSls.ValidInput(txtFromDate) And txtFromDate <> "" Then
        If Box.Text < txtFromDate Then
            fnValidSalesDate = False
            cValidSls.SetErrorMessage Box, "To Date must be greater than From Date"
            Exit Function
        End If
    End If
    
    fnValidSalesDate = True
    
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
        tgcDropdown.ComboSQL(txtFromDate) = fnSetComboSQL(txtFromDate.TabIndex)
        Screen.MousePointer = vbHourglass
    End If
        
    bKeyCode = tgcDropdown.Keypress(txtFromDate, KeyAscii)
    Screen.MousePointer = vbDefault
    
    If Not bKeyCode Then
       If KeyAscii = vbKeyReturn Then
          If tgcDropdown.SingleRecordSelected Then
                subSetFocus txtToDate
                Screen.MousePointer = vbDefault
          End If
          KeyAscii = 0
       End If
    Else
        cValidSls.Keypress txtFromDate, KeyAscii
    End If
End Sub

Private Sub txtFromDate_LostFocus()
    cValidSls.LostFocus txtFromDate, cmdFromDate, tblComboDropdown
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
        tgcDropdown.ComboSQL(txtToDate) = fnSetComboSQL(txtToDate.TabIndex)
        Screen.MousePointer = vbHourglass
    End If
        
    bKeyCode = tgcDropdown.Keypress(txtToDate, KeyAscii)
    Screen.MousePointer = vbDefault
    
    If Not bKeyCode Then
       If KeyAscii = vbKeyReturn Then
          If tgcDropdown.SingleRecordSelected Then
                subEnterPhaseIISlsOrHrs TabSales
                Screen.MousePointer = vbDefault
          End If
          KeyAscii = 0
       End If
    Else
        cValidSls.Keypress txtToDate, KeyAscii
    End If
End Sub

Private Sub txtToDate_LostFocus()
    cValidSls.LostFocus txtToDate, cmdToDate, tblComboDropdown
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
    
    If tblSales.Row = -1 Then
        subEnableDeleteBtn False, TabSales
    End If
    
End Sub

Private Sub tblSales_GotFocus()
    tfnSetStatusBarMessage "Store Sales"
    tgmSales.GotFocus
    tgfDropdown(TabSales).GotFocus
End Sub

Private Sub tblSales_KeyDown(KeyCode As Integer, Shift As Integer)
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
    
    If t_nFormMode = EDIT_MODE Then
        If KeyAscii = vbKeyReturn And tgmSales.RowCount - 1 = lRow And tblSales.col = colSToDate Then
            If cmdUpdateInsertBtn(TabSales).Enabled Then
                subSetFocus cmdUpdateInsertBtn(TabSales)
            Else
                subSetFocus cmdCancel(TabSales)
            End If
        End If
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

Private Sub tblSales_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
    'frmContext.MouseDown Button, EMP_MST_UP
End Sub

Private Sub tblSales_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    Dim lRow As Long
    
    If tgfDropdown(TabSales).RowColChange(tblSales) Then
        tgmSales.RowColChange LastRow, LastCol
    End If
    
    lRow = tgmSales.GetCurrentRowNumber
    
    If t_nFormMode = ADD_MODE Then
        If tgmSales.CellValue(colSPrftCtr, lRow) <> "" Then
            subEnableDeleteBtn True, TabSales
        Else
            subEnableDeleteBtn False, TabSales
        End If
    Else
        If tgmSales.RowCount > 0 Then
            subEnableDeleteBtn True, TabSales
        Else
            subEnableDeleteBtn False, TabSales
        End If
    End If
    
    subSetStdBtn TabSales, tgmSales

End Sub

Private Sub tblSales_SelChange(CANCEL As Integer)
    CANCEL = True
End Sub

Private Sub tblSales_UnboundReadData(ByVal RowBuf As DBTrueGrid.RowBuffer, StartLocation As Variant, ByVal ReadPriorRows As Boolean)
    tgmSales.ReadData RowBuf, StartLocation, ReadPriorRows
End Sub

Private Function fnValidClockInDate(ByVal szText As String, ByVal lRow As Long) As Boolean
    Const SUB_NAME As String = "fnValidClockInDate"
    Dim strSQL As String
    Dim rsTemp As Recordset
    Dim nError As Integer
    
    fnValidClockInDate = False
    
    If Trim(szText) = "" Then
        tgmHours.ErrorMessage(colHClockIn) = "You must enter a Clock-in Date"
        Exit Function
    End If
    
    If Not IsFactorDate(szText, nError) Then 'the format of sztext has been changed here
        If nError = 1 Then
            tgmHours.ErrorMessage(colHClockIn) = "Invalid date format"
        Else
            tgmHours.ErrorMessage(colHClockIn) = "Invalid date"
        End If
        Exit Function
    End If
    
    'check in table gl_period
    strSQL = "SELECT glp_status FROM gl_period "
    strSQL = strSQL & " WHERE glp_end_dt >= " & tfnSQLString(CDate(szText))
    strSQL = strSQL & " AND glp_beg_dt <= " & tfnSQLString(CDate(szText))
    
    If GetRecordSet(rsTemp, strSQL, , SUB_NAME) <= 0 Then
        tgmHours.ErrorMessage(colHClockIn) = "There is no valid G/L Period containing this Date"
        Exit Function
    End If
    
    Select Case fnGetField(rsTemp!glp_status)
        Case "O"
        Case "W"
            tgmHours.ErrorMessage(colHClockIn) = "This G/L Period is about to close!"
        Case "C"
            tgmHours.ErrorMessage(colHClockIn) = "This G/L Period is already closed"
            Exit Function
        Case Else
            tgmHours.ErrorMessage(colHClockIn) = "Invalid G/L Period Status"
            Exit Function
    End Select
    
    fnValidClockInDate = True
    
End Function

'This function will return all the values already used in the grid.
Private Function fnCheckDubInHrsGrid(lCurrRow As Long) As String
    Dim lRowCount As Long
    Dim k As Integer
    Dim sTemp As String
    
    fnCheckDubInHrsGrid = ""
    
    lRowCount = tgmHours.RowCount
    sTemp = ""
    
    If lRowCount > 0 Then
        For k = 0 To lRowCount - 1
            If k <> lCurrRow Then
                If sTemp = "" Then
                    sTemp = tgmHours.CellValue(colHClockIn, k) & "~" & tgmHours.CellValue(colHPrftCtr, k) & "~" & tgmHours.CellValue(colHPayCode, k)
                Else
                    sTemp = sTemp & "," & tgmHours.CellValue(colHClockIn, k) & "~" & tgmHours.CellValue(colHPrftCtr, k) & "~" & tgmHours.CellValue(colHPayCode, k)
                End If
            End If
        Next
    End If
    
    fnCheckDubInHrsGrid = sTemp

End Function

Private Sub subSetStdBtn(Index As Integer, tgmEditor As clsTGSpreadSheet)
    
    If tgmEditor.RowCount < 1 Then
        subEnableUpdateBtn False, Index
        subEnableDeleteBtn False, Index
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
    Dim sMsg As String
    
    fnValidGridFromToDate = False
    
    If nCol = colSFromDate Then
        sMsg = " From Date"
    Else
        sMsg = " To Date"
    End If
    
    If Trim(sText) = "" Then
        tgmSales.ErrorMessage(nCol) = "You must enter a " & sMsg
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

Private Function fnLoadSales() As Boolean
    Const SUB_NAME As String = "fnLoadSales"
    Dim strSQL As String
    
    fnLoadSales = False
    strSQL = fnGetSalesSQL()
    
    If GetRecordCount(strSQL, , SUB_NAME) <= 0 Then
        MsgBox "Failed to load the sales", vbExclamation
        Exit Function
    End If
    
    tgmSales.FillWithSQL t_dbMainDatabase, strSQL
    fnLoadSales = True

End Function

Private Function fnGetSalesSQL(Optional bCombo As Boolean = False, _
                               Optional nTabIndex As Integer) As String
    Dim strSQL As String
    Dim sSalesType As String
    
    If t_nFormMode = EDIT_MODE Then
        sSalesType = "EDIT_MODE"
    Else
        sSalesType = fnGetSalesType
    End If
    
    Select Case sSalesType
        Case sWeek
            strSQL = ""
            strSQL = strSQL & ""
        Case sOneMth
            strSQL = ""
            strSQL = strSQL & ""
        Case sThreeMth
            strSQL = ""
            strSQL = strSQL & ""
        Case sGas
            strSQL = "SELECT rsd_prft_ctr AS prft_ctr, prft_name, "
            strSQL = strSQL & " SUM(rsd_gal) as amount, rsd_date as from_date"
            strSQL = strSQL & " FROM rs_daily, sys_prft_ctr "
            strSQL = strSQL & " WHERE prft_ctr = rsd_prft_ctr"
            If Trim(txtFromDate) <> "" Then
                strSQL = strSQL & " AND rsd_date >= " & tfnDateString(Trim(txtFromDate), True)
            End If
            If Trim(txtToDate) <> "" Then
                strSQL = strSQL & " AND rsd_date <= " & tfnDateString(Trim(txtToDate), True)
            End If
            strSQL = strSQL & " GROUP BY rsd_date, rsd_prft_ctr, prft_name"
            strSQL = strSQL & " ORDER BY rsd_prft_ctr"
            If bCombo Then
                If nTabIndex = txtFromDate.TabIndex Then
                    strSQL = "SELECT rsd_date AS bs_from_date FROM rs_daily WHERE 1=1"
                    If txtToDate <> "" Then
                        strSQL = strSQL & " AND rsd_date <= " & tfnDateString(Trim(txtToDate), True)
                    End If
                ElseIf nTabIndex = txtToDate.TabIndex Then
                    strSQL = "SELECT rsd_date AS bs_to_date FROM rs_daily WHERE 1=1"
                    If txtFromDate <> "" Then
                        strSQL = strSQL & " AND rsd_date >= " & tfnDateString(Trim(txtFromDate), True)
                    End If
                End If
                strSQL = strSQL & " GROUP BY rsd_date"
            End If
        Case "EDIT_MODE"
            strSQL = "SELECT bs_to_date as to_date, bs_from_date as from_date,"
            strSQL = strSQL & " bs_sales_amount as amount, prft_ctr, prft_name"
            strSQL = strSQL & " FROM bonus_sales, sys_prft_ctr "
            strSQL = strSQL & " WHERE bs_prft_ctr = prft_ctr"
            strSQL = strSQL & " AND bs_sales_type = " & tfnSQLString(Trim(fnGetSalesType))
            If Trim(txtFromDate) <> "" Then
                strSQL = strSQL & " AND bs_from_date >= " & tfnDateString(Trim(txtFromDate), True)
            End If
            If Trim(txtToDate) <> "" Then
                strSQL = strSQL & " AND bs_to_date <= " & tfnDateString(Trim(txtToDate), True)
            End If
            If bCombo Then
                strSQL = "SELECT bs_from_date, bs_to_date FROM bonus_sales"
                If nTabIndex = txtFromDate.TabIndex Then
                    If txtToDate <> "" Then
                        strSQL = strSQL & " AND bs_from_date <= " & tfnDateString(Trim(txtToDate), True)
                    End If
                ElseIf nTabIndex = txtToDate.TabIndex Then
                    If txtFromDate <> "" Then
                        strSQL = strSQL & " WHERE bs_to_date >= " & tfnDateString(Trim(txtFromDate), True)
                    End If
                End If
            End If
    End Select
    fnGetSalesSQL = strSQL
    
End Function

Private Function fnLoadPrftCtr() As Boolean
    Dim i As Integer, j As Integer
    Dim sArPrftCtr() As Variant
    
    tgmPrftCtr.ClearData
    For i = 0 To tgmHours.RowCount
        If i = tgmHours.RowCount Then
            Exit For
        End If
        If i = 0 Then
            ReDim Preserve sArPrftCtr(1, j)
            sArPrftCtr(0, j) = tgmHours.CellValue(colHPrftCtr, i)
            sArPrftCtr(1, j) = tfnRound(tgmHours.CellValue(colHHrsDol, i), DEFAULT_DECIMALS)
        Else
            If tgmHours.CellValue(colHPrftCtr, i) = sArPrftCtr(0, j) Then
                sArPrftCtr(1, j) = tfnRound(sArPrftCtr(1, j), 6) + tfnRound(tgmHours.CellValue(colHHrsDol, i), 6)
            Else
                j = j + 1
                ReDim Preserve sArPrftCtr(1, j)
                sArPrftCtr(0, j) = tgmHours.CellValue(colHPrftCtr, i)
                sArPrftCtr(1, j) = tfnRound(tgmHours.CellValue(colHHrsDol, i), DEFAULT_DECIMALS)
            End If
        End If
    Next i
    
    tgmPrftCtr.FillWithArray sArPrftCtr
    subUpdateTotals

End Function

Private Sub subUpdateTotals()
    Dim i As Integer
    
    lblTotalHrs = ""
    lblPCTotal = ""

    For i = 0 To tgmHours.RowCount - 1
        lblTotalHrs = Val(lblTotalHrs) + tfnRound(tgmHours.CellValue(colHHrsDol, i), DEFAULT_DECIMALS)
    Next i

    For i = 0 To tgmPrftCtr.RowCount - 1
        lblPCTotal = Val(lblPCTotal) + tfnRound(tgmPrftCtr.CellValue(colPTotal, i), DEFAULT_DECIMALS)
    Next i

End Sub
