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
      Size            = 9.75 
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
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   516
      Width           =   8880
      _Version        =   65536
      _ExtentX        =   15663
      _ExtentY        =   9144
      _StockProps     =   77
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
      BeginProperty PanelFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            = 9.75 
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin FACTFRMLib.FactorFrame cmdExit 
         Height          =   390
         HelpContextID   =   15
         Left            =   7524
         TabIndex        =   50
         Top             =   4716
         Width           =   1260
         _Version        =   65536
         _ExtentX        =   2222
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
      Begin FACTTABLib.FactorTab etabMain 
         Height          =   4860
         HelpContextID   =   201
         Left            =   96
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   84
         Width           =   8892
         _Version        =   65536
         _ExtentX        =   15690
         _ExtentY        =   8572
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
         TabsPerPage     =   5
         Caption         =   "&Job Class|Pay &Code|Pay Incr &Maint.|Pay &Incr Report"
         Begin FACTFRMLib.FactorFrame efraTabBack4 
            Height          =   4572
            Left            =   14580
            TabIndex        =   46
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
            Begin FACTFRMLib.FactorFrame cmdUpdateInsertBtn 
               Height          =   390
               HelpContextID   =   16
               Index           =   3
               Left            =   5916
               TabIndex        =   32
               Top             =   4008
               Width           =   1308
               _Version        =   65536
               _ExtentX        =   2307
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
            Begin FACTFRMLib.FactorFrame efraTabFrame4 
               Height          =   3852
               Left            =   72
               TabIndex        =   47
               TabStop         =   0   'False
               Top             =   96
               Width           =   8652
               _Version        =   65536
               _ExtentX        =   15266
               _ExtentY        =   6800
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
               Begin VB.TextBox txtProcessDate 
                  Height          =   360
                  Left            =   4188
                  TabIndex        =   28
                  Top             =   120
                  Width           =   2136
               End
               Begin DBTrueGrid.TDBGrid tblGrid 
                  Height          =   3264
                  HelpContextID   =   1408
                  Left            =   36
                  OleObjectBlob   =   "zzsepimt.frx":0000
                  TabIndex        =   30
                  Top             =   528
                  Width           =   8544
               End
               Begin VB.PictureBox picText 
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000005&
                  ForeColor       =   &H80000008&
                  Height          =   324
                  Left            =   108
                  ScaleHeight     =   300
                  ScaleWidth      =   1020
                  TabIndex        =   48
                  Top             =   456
                  Width           =   1044
               End
               Begin VB.Label Label1 
                  Caption         =   "Enter the date for calculating the pay increase"
                  Height          =   264
                  Left            =   84
                  TabIndex        =   53
                  Top             =   156
                  Width           =   4176
               End
            End
            Begin FACTFRMLib.FactorFrame cmdExitCancelBtn 
               Height          =   390
               HelpContextID   =   15
               Index           =   3
               Left            =   7380
               TabIndex        =   33
               Top             =   4008
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
               Caption         =   "Ca&ncel"
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
               Height          =   390
               HelpContextID   =   1410
               Index           =   3
               Left            =   4452
               TabIndex        =   31
               Top             =   4008
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
               Caption         =   "Print &Report"
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
            Begin FACTFRMLib.FactorFrame cmdProcess 
               Height          =   390
               HelpContextID   =   1409
               Index           =   3
               Left            =   2988
               TabIndex        =   29
               Top             =   4008
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
               Caption         =   "&Process"
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
         Begin FACTFRMLib.FactorFrame efraTabBack3 
            Height          =   4572
            Left            =   14520
            TabIndex        =   41
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
            Begin FACTFRMLib.FactorFrame cmdRefreshSelectBtn 
               Height          =   390
               HelpContextID   =   14
               Index           =   2
               Left            =   5916
               TabIndex        =   25
               Top             =   4008
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
            Begin FACTFRMLib.FactorFrame efraTabFrame3 
               Height          =   3852
               Left            =   72
               TabIndex        =   42
               TabStop         =   0   'False
               Top             =   96
               Width           =   8652
               _Version        =   65536
               _ExtentX        =   15266
               _ExtentY        =   6800
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
               Begin DBTrueGrid.TDBGrid tblSpreadsheet 
                  Height          =   2028
                  HelpContextID   =   1407
                  Left            =   636
                  OleObjectBlob   =   "zzsepimt.frx":30CF
                  TabIndex        =   23
                  Top             =   1584
                  Width           =   7380
               End
               Begin FACTFRMLib.FactorFrame efraTop 
                  Height          =   1068
                  Left            =   132
                  TabIndex        =   43
                  TabStop         =   0   'False
                  Top             =   156
                  Width           =   8412
                  _Version        =   65536
                  _ExtentX        =   14838
                  _ExtentY        =   1884
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
                  Begin VB.TextBox txtLevel 
                     Height          =   360
                     HelpContextID   =   1403
                     Left            =   504
                     MaxLength       =   1
                     TabIndex        =   19
                     Top             =   480
                     Width           =   2205
                  End
                  Begin VB.TextBox txtDesc 
                     Height          =   360
                     HelpContextID   =   1405
                     Left            =   3972
                     MaxLength       =   20
                     TabIndex        =   21
                     Top             =   480
                     Width           =   3495
                  End
                  Begin FACTFRMLib.FactorFrame cmdLevel 
                     CausesValidation=   0   'False
                     Height          =   360
                     HelpContextID   =   1404
                     Left            =   2724
                     TabIndex        =   20
                     TabStop         =   0   'False
                     Tag             =   "Run #3"
                     Top             =   480
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
                     Picture         =   "zzsepimt.frx":43F5
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
                  Begin FACTFRMLib.FactorFrame cmdDesc 
                     CausesValidation=   0   'False
                     Height          =   360
                     HelpContextID   =   1406
                     Left            =   7482
                     TabIndex        =   22
                     TabStop         =   0   'False
                     Tag             =   "Run #3"
                     Top             =   480
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
                     Picture         =   "zzsepimt.frx":4507
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
                  Begin VB.Label Label2 
                     Caption         =   "Employee Level"
                     Height          =   276
                     Left            =   528
                     TabIndex        =   45
                     Top             =   168
                     Width           =   2088
                  End
                  Begin VB.Label Label3 
                     Caption         =   "Employee Level Description"
                     Height          =   348
                     Left            =   3984
                     TabIndex        =   44
                     Top             =   132
                     Width           =   3420
                  End
               End
            End
            Begin FACTFRMLib.FactorFrame cmdDeleteBtn 
               Height          =   390
               HelpContextID   =   12
               Index           =   2
               Left            =   2988
               TabIndex        =   27
               Top             =   4008
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
            Begin FACTFRMLib.FactorFrame cmdEditBtn 
               Height          =   390
               HelpContextID   =   11
               Index           =   2
               Left            =   1512
               TabIndex        =   18
               Top             =   4008
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
            Begin FACTFRMLib.FactorFrame cmdAddBtn 
               Height          =   390
               HelpContextID   =   10
               Index           =   2
               Left            =   72
               TabIndex        =   17
               Top             =   4008
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
            Begin FACTFRMLib.FactorFrame cmdExitCancelBtn 
               Height          =   390
               HelpContextID   =   15
               Index           =   2
               Left            =   7380
               TabIndex        =   26
               Top             =   4008
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
               Caption         =   "Ca&ncel"
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
               Height          =   390
               HelpContextID   =   16
               Index           =   2
               Left            =   4452
               TabIndex        =   24
               Top             =   4008
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
         End
         Begin FACTFRMLib.FactorFrame efraTabBack1 
            Height          =   4548
            Left            =   12
            TabIndex        =   35
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
               Size            = 9.75 
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   0
            BeginProperty PanelFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            = 9.75 
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Begin FACTFRMLib.FactorFrame efraTabFrame1 
               Height          =   3852
               Left            =   72
               TabIndex        =   7
               TabStop         =   0   'False
               Top             =   96
               Width           =   8652
               _Version        =   65536
               _ExtentX        =   15266
               _ExtentY        =   6800
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
               Begin DBTrueGrid.TDBGrid tblJobClass 
                  Height          =   3456
                  HelpContextID   =   1401
                  Left            =   636
                  OleObjectBlob   =   "zzsepimt.frx":4619
                  TabIndex        =   2
                  Top             =   192
                  Width           =   7380
               End
            End
            Begin FACTFRMLib.FactorFrame cmdDeleteBtn 
               Height          =   390
               HelpContextID   =   12
               Index           =   0
               Left            =   2988
               TabIndex        =   6
               Top             =   4008
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
            Begin FACTFRMLib.FactorFrame cmdEditBtn 
               Height          =   390
               HelpContextID   =   11
               Index           =   0
               Left            =   1512
               TabIndex        =   1
               Top             =   4008
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
            Begin FACTFRMLib.FactorFrame cmdAddBtn 
               Height          =   390
               HelpContextID   =   10
               Index           =   0
               Left            =   72
               TabIndex        =   0
               Top             =   4008
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
            Begin FACTFRMLib.FactorFrame cmdUpdateInsertBtn 
               Height          =   390
               HelpContextID   =   14
               Index           =   0
               Left            =   4452
               TabIndex        =   3
               Top             =   4008
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
            Begin FACTFRMLib.FactorFrame cmdExitCancelBtn 
               Height          =   390
               HelpContextID   =   15
               Index           =   0
               Left            =   7380
               TabIndex        =   5
               Top             =   4008
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
               Caption         =   "Ca&ncel"
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
               Height          =   390
               HelpContextID   =   16
               Index           =   0
               Left            =   5916
               TabIndex        =   4
               Top             =   4008
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
         Begin FACTFRMLib.FactorFrame efraTabBack2 
            Height          =   4572
            Left            =   14460
            TabIndex        =   36
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
               Size            = 9.75 
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   0
            BeginProperty PanelFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            = 9.75 
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Begin FACTFRMLib.FactorFrame cmdRefreshSelectBtn 
               Height          =   390
               HelpContextID   =   14
               Index           =   1
               Left            =   5916
               TabIndex        =   12
               Top             =   4008
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
            Begin FACTFRMLib.FactorFrame efraTabFrame2 
               Height          =   3852
               Left            =   72
               TabIndex        =   49
               TabStop         =   0   'False
               Top             =   96
               Width           =   8652
               _Version        =   65536
               _ExtentX        =   15266
               _ExtentY        =   6800
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
               Begin FACTFRMLib.FactorFrame cmdFDropDown2 
                  Height          =   360
                  HelpContextID   =   1602
                  Left            =   3552
                  TabIndex        =   52
                  TabStop         =   0   'False
                  Top             =   2052
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
                  Picture         =   "zzsepimt.frx":5940
                  Style           =   3
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
               Begin VB.Data datFDropDown2 
                  Caption         =   "Data1"
                  Connect         =   "Access"
                  DatabaseName    =   ""
                  DefaultCursorType=   0  'DefaultCursor
                  DefaultType     =   2  'UseODBC
                  Exclusive       =   0   'False
                  Height          =   588
                  Left            =   1260
                  Options         =   0
                  ReadOnly        =   0   'False
                  RecordsetType   =   1  'Dynaset
                  RecordSource    =   ""
                  Top             =   2100
                  Visible         =   0   'False
                  Width           =   1404
               End
               Begin DBTrueGrid.TDBGrid tblPayCode 
                  Height          =   3456
                  HelpContextID   =   1402
                  Left            =   636
                  OleObjectBlob   =   "zzsepimt.frx":5A52
                  TabIndex        =   10
                  Top             =   192
                  Width           =   7380
               End
            End
            Begin FACTFRMLib.FactorFrame cmdDeleteBtn 
               Height          =   390
               HelpContextID   =   12
               Index           =   1
               Left            =   2988
               TabIndex        =   14
               Top             =   4008
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
            Begin FACTFRMLib.FactorFrame cmdEditBtn 
               Height          =   390
               HelpContextID   =   11
               Index           =   1
               Left            =   1512
               TabIndex        =   9
               Top             =   4008
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
            Begin FACTFRMLib.FactorFrame cmdAddBtn 
               Height          =   390
               HelpContextID   =   10
               Index           =   1
               Left            =   72
               TabIndex        =   8
               Top             =   4008
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
            Begin FACTFRMLib.FactorFrame cmdExitCancelBtn 
               Height          =   390
               HelpContextID   =   15
               Index           =   1
               Left            =   7380
               TabIndex        =   13
               Top             =   4008
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
               Caption         =   "Ca&ncel"
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
               Height          =   390
               HelpContextID   =   16
               Index           =   1
               Left            =   4452
               TabIndex        =   11
               Top             =   4008
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
         End
      End
   End
   Begin FACTFRMLib.FactorFrame ffraStatusbar 
      Height          =   360
      Left            =   0
      TabIndex        =   15
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
      Bindings        =   "zzsepimt.frx":6D68
      Height          =   2484
      Left            =   408
      OleObjectBlob   =   "zzsepimt.frx":6D87
      TabIndex        =   40
      TabStop         =   0   'False
      Top             =   684
      Width           =   3756
   End
   Begin SQAOTestObjectsCtl.SQAOTest SQAOTest1 
      Height          =   456
      Left            =   10896
      TabIndex        =   39
      TabStop         =   0   'False
      Top             =   96
      Width           =   456
      _ExtentX        =   804
      _ExtentY        =   804
   End
   Begin FACTFRMLib.FactorFrame efraToolBar 
      Height          =   465
      Left            =   0
      TabIndex        =   37
      TabStop         =   0   'False
      Top             =   0
      Width           =   8880
      _Version        =   65536
      _ExtentX        =   15663
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
      FMName          =   "ZZSEPIMT"
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
         Height          =   372
         Left            =   60
         TabIndex        =   38
         Top             =   84
         Width           =   6876
         _ExtentX        =   12129
         _ExtentY        =   656
         ButtonWidth     =   635
         ButtonHeight    =   572
         _Version        =   393216
      End
   End
   Begin DBTrueGrid.TDBGrid tblFDropdown2 
      Bindings        =   "zzsepimt.frx":806B
      Height          =   3180
      Left            =   0
      OleObjectBlob   =   "zzsepimt.frx":8087
      TabIndex        =   51
      Top             =   0
      Width           =   4812
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
Private tgmJobClass As clsTGSpreadSheet
Private tgmPayCode As clsTGSpreadSheet
Private strHourCode As String
Private strOverCode As String
Private objText As clsColumnExtension
'===============================
'spreadsheet field
'===============================
Private Const COL_DAYS = 0
Private Const COL_RATE = 1
Private nComboForAdd As Integer
Private nComboForEdit As Integer
'===============================
'Job Class Grid
'===============================
Private Const COL_JOBCLASS = 0
Private Const COL_JOBDESC = 1
'===============================
'Pay Code Grid
'===============================
Private Const COL_PAYCODE = 0
Private Const COL_CONVFACTOR = 1
'tblGrid
Private Const COL_EMPNO = 0
Private Const COL_NAME = 1
Private Const COL_JCLASS = 2
Private Const COL_LEVEL = 3
Private Const COL_DESC = 4
Private Const COL_HIRE_DATE = 5
Private Const COL_DAYS_WORKED = 6
Private Const COL_PCODE = 7
Private Const COL_OLD_RATE = 8
Private Const COL_NEW_RATE = 9
Private Const COL_DIFF = 10
Private nCOL_CONVT_FACTOR As Integer
Private nCOL_PREP_YEAR As Integer

Public rsReport As Recordset

Private Const DB_ERR_READ = "Read database error. Contact Factor."
Private tgfDropDown1 As clsFloatingDropDown
Private tgfDropDown2 As clsFloatingDropDown
Private nFDropDown1 As Integer
Private nFDropDown2 As Integer

Public Sub subEnableSearchbtn(cmdButton As Control, bYesNo As Boolean)
    #If Win16 Then
        If bYesNo Then
            cmdButton.Picture = LoadResPicture(SEARCH_UP, vbResBitmap)
        Else
            cmdButton.Picture = LoadResPicture(SEARCH_DOWN, vbResBitmap)
        End If
        
     #Else
        If bYesNo Then
            cmdButton.Picture = frmContext.LoadPicture(SEARCH_UP)
        Else
            cmdButton.Picture = frmContext.LoadPicture(SEARCH_DOWN)
        End If
     #End If
     
     cmdButton.Enabled = bYesNo
End Sub
Private Sub subInitFloatingDropdowns()
    
    Dim strSQL As String
    
    ' floating dropdown for Tab_1
'    subEnableSearchbtn cmdFDropDown1, True
'
'    Set tgfDropDown1 = New clsFloatingDropDown
'    Set tgfDropDown1.DataBase = t_dbMainDatabase
'    Set tgfDropDown1.Form = Me
'    Set tgfDropDown1.SearchButton = cmdFDropDown1 ' a new search button
'    Set tgfDropDown1.DropDownTable = tblFDropdown1 ' the true Grid f dropdown
'    Set tgfDropDown1.MainTable = tblJobClass   ' the main true grid
'    Set tgfDropDown1.DataLink = datFDropDown1    ' name of data control for F-dropdown
'    Set tgfDropDown1.EditClass = tgmJobClass
'    tgfDropDown1.DefaultCursorControl = True
'
'    'nFDropDown1 = tgfDropDown1.AddDropDown(, strSQL)
'    tgfDropDown1.AddDropDown 1, "SELECT DISTINCT prm_job_class FROM pr_master "
'    tgfDropDown1.AddColumn 0, "prm_job_class", tgfDropDown1.COLUMN_TYPE_STRING
'
'    tgfDropDown1.SetUnique 0, True
    
    ' floating dropdown for Tab_2
    subEnableSearchbtn cmdFDropDown2, True
    
    Set tgfDropDown2 = New clsFloatingDropDown
    Set tgfDropDown2.DataBase = t_dbMainDatabase
    Set tgfDropDown2.Form = Me
    Set tgfDropDown2.SearchButton = cmdFDropDown2 ' a new search button
    Set tgfDropDown2.DropDownTable = tblFDropdown2 ' the true Grid f dropdown
    Set tgfDropDown2.MainTable = tblPayCode  ' the main true grid
    Set tgfDropDown2.DataLink = datFDropDown2    ' name of data control for F-dropdown
    Set tgfDropDown2.EditClass = tgmPayCode
    tgfDropDown2.DefaultCursorControl = True
    
    'nFDropDown2 = tgfDropDown2.AddDropDown(, strSQL)
    tgfDropDown2.AddDropDown 1, "SELECT prpa_pay_code,prpa_desc FROM pr_pay "
    tgfDropDown2.AddColumn COL_PAYCODE, "prpa_pay_code", tgfDropDown2.COLUMN_TYPE_STRING
    tgfDropDown2.AddExtraColumn "Pay Description", "prpa_desc", 3000, tgfDropDown2.COLUMN_TYPE_STRING
    tgfDropDown2.SetUnique COL_PAYCODE, True
End Sub

Private Sub subSetupCombos()

    Set tgcDropdown = CreateObject(t_szOLECOMBO)
    Set tgcDropdown.DBEngine = t_engFactor
    Set tgcDropdown.Form = Me
    Set tgcDropdown.DataBase = t_dbMainDatabase
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
    Dim szDecPatternConvtFactor As String
    szDecPattern = tfnDecimalPattern(8, 2, False, False)
    szDecPatternConvtFactor = tfnDecimalPattern(10, 6, False, True)
    Dim myWidth As Long
    
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

    Set tgmJobClass = New clsTGSpreadSheet
    Set tgmJobClass.Table = tblJobClass
    Set tgmJobClass.StatusBar = ffraStatusbar
    Set tgmJobClass.Form = Me

    tgmJobClass.AddEditColumn COL_JOBCLASS, "Enter a Job Class", "^P{0,2}$"
    tgmJobClass.AddEditColumn COL_JOBDESC, "Enter a Job Description", "^P{0,30}$"
    myWidth = tblJobClass.Width - 250
    With tblJobClass
        .Columns(COL_JOBCLASS).DataField = "zzsej_job_class"
        .Columns(COL_JOBDESC).DataField = "zzsej_job_desc"
        .Columns(COL_JOBCLASS).Width = 0.2 * myWidth
        .Columns(COL_JOBDESC).Width = 0.1 * myWidth
    End With
    
    Set tgmPayCode = New clsTGSpreadSheet
    Set tgmPayCode.Table = tblPayCode
    Set tgmPayCode.StatusBar = ffraStatusbar
    Set tgmPayCode.Form = Me

    tgmPayCode.AddEditColumn COL_PAYCODE, "Enter a Pay Code", "^P{0,4}$"
    tgmPayCode.AddEditColumn COL_CONVFACTOR, "Enter a Pay Code Conversion Factor", szDecPatternConvtFactor
    myWidth = tblPayCode.Width - 250
    With tblPayCode
        .Columns(COL_PAYCODE).DataField = "prpa_pay_code"
        .Columns(COL_CONVFACTOR).DataField = "zzsep_factor"
        .Columns(COL_PAYCODE).Width = 0.2 * myWidth
        .Columns(COL_CONVFACTOR).Width = 0.1 * myWidth
    End With
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
    cValidate.SetFirstControls cmdAddBtn(0), cmdEditBtn(0), cmdDeleteBtn(0), cmdRefreshSelectBtn(0), cmdUpdateInsertBtn(2), cmdExitCancelBtn(0)
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

Private Sub cmdFDropDown2_Click()
    tgfDropDown2.ButtonClick
End Sub

Private Sub cmdFDropDown2_LostFocus()
    tgfDropDown2.LostFocus cmdFDropDown2
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
    
    If Not IsFactorDate(txtProcessDate) Then
        MsgBox "Please enter the date for calculating the pay increase", vbExclamation + vbOKOnly
        subSetFocus txtProcessDate
        Exit Sub
    End If
    
    tblGrid.Enabled = True
    #If PROTOTYPE Then
        Exit Sub
    #End If
    etabMain.TabEnabled(TAB_1) = False
    etabMain.TabEnabled(TAB_2) = False
    etabMain.TabEnabled(TAB_3) = False
    subEnableCancelBtn True
    cmdProcess(3).Enabled = False
    txtProcessDate.Enabled = False
    subEnableUpdateBtn False
    
    If Not fnCheckPayCode Then
        Exit Sub
    End If
    
    If Not fnCheckJobClass Then
        Exit Sub
    End If
    tfnSetStatusBarMessage "Processing...Please wait..."
    subShowHourGlass
    If fnProcess(rsReport) Then
        subEnableUpdateBtn True
        cmdPrint(3).Enabled = True
        cmdPrint(3).SetFocus
        subEnablePrint True
    End If
    subHideHourGlass
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
    ElseIf etabMain.CurrTab = TAB_2 Then
        subSetFocus cmdAddBtn(1)
    ElseIf etabMain.CurrTab = TAB_3 Then
        subSetFocus cmdAddBtn(2)
    Else
        txtProcessDate.Enabled = True
        cmdProcess(3).Enabled = True
        txtProcessDate.SetFocus
        'cmdProcess(3).SetFocus
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
    cmdPrint(3).Enabled = False
    subEnablePrint False
    subEnableUpdateBtn False
    subEnableCancelBtn False, 0
    subEnableCancelBtn False, 1
    subEnableCancelBtn False, 2
    Select Case NewTab
         Case TAB_1
            subEnableDeleteBtn False, 0
            
            subEnableRefreshBtn False, 0
         Case TAB_2
            subEnableDeleteBtn False, 1
            
            subEnableRefreshBtn False, 1
         Case TAB_3
            subEnableDeleteBtn False, 2
            
            subEnableRefreshBtn False, 2
            
         Case TAB_4
            txtProcessDate.Enabled = True
            cmdProcess(3).Enabled = True
            subEnableCancelBtn False, 3
            
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
'#If Not PROTOTYPE Then
'    If Not fnGetPayCode Then
'        MsgBox "System Error: Unable to access system parameter, Program will be terminated", vbExclamation
'        Unload Me
'        Exit Sub
'    End If
'#End If
    subSetupCombos  'Set up the combo dropdowns
    subInitializeSpreadSheet 'set up true grid
    subInitGrid
    subInitFloatingDropdowns
    subInitText     'tooltips
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
    Select Case etabMain.CurrTab
    Case TAB_1
        efraTabFrame1.Enabled = True
        etabMain.TabEnabled(TAB_2) = False
        etabMain.TabEnabled(TAB_3) = False
        etabMain.TabEnabled(TAB_4) = False
        cmdUpdateInsertBtn(0).Caption = t_szCAPTION_INSERT
        tgmJobClass.AddEditColumn COL_JOBCLASS, "Enter a Job Class", "^P{0,2}$"
        tgmJobClass.AllowAddNew = True
        tgmJobClass.ClearData
        tblJobClass.Enabled = True
        subSetFocus tblJobClass
    Case TAB_2
        efraTabFrame2.Enabled = True
        etabMain.TabEnabled(TAB_1) = False
        etabMain.TabEnabled(TAB_3) = False
        etabMain.TabEnabled(TAB_4) = False
        cmdUpdateInsertBtn(1).Caption = t_szCAPTION_INSERT
        tgmPayCode.AddEditColumn COL_PAYCODE, "Enter a Pay Code", "^P{0,4}$"
        'tgmPayCode.AddEditColumn COL_CONVFACTOR
        tgmPayCode.AllowAddNew = True
        tgmPayCode.ClearData
        tblPayCode.Enabled = True
        subSetFocus tblPayCode
    Case TAB_3
        efraTabFrame3.Enabled = True
        etabMain.TabEnabled(TAB_1) = False
        etabMain.TabEnabled(TAB_2) = False
        etabMain.TabEnabled(TAB_4) = False
        tgcDropdown.ComboActive(nComboForAdd) = True
        tgcDropdown.ComboActive(nComboForEdit) = False
        tgmSpreadsheet.AllowAddNew = True
        tgmSpreadsheet.ClearData
        subEnableTop True
        cmdUpdateInsertBtn(2).Caption = t_szCAPTION_INSERT
        cValidate.ResetFlags
        txtLevel.SetFocus
    Case TAB_4
        efraTabFrame4.Enabled = True
        etabMain.TabEnabled(TAB_1) = False
        etabMain.TabEnabled(TAB_2) = False
        etabMain.TabEnabled(TAB_3) = False
        tgmGrid.ClearData
    End Select
    subEnableAddBtn False
    subEnableEditBtn False
    subEnableCancelBtn True
    mnuUpdateInsert.Caption = t_szCAPTION_INSERT
    t_nFormMode = ADD_MODE
End Sub

Private Sub cmdEditBtn_Click(Index As Integer)
    #If PROTOTYPE Then
        Exit Sub
    #End If
    t_nFormMode = EDIT_MODE
    Select Case etabMain.CurrTab
    Case TAB_1
        etabMain.TabEnabled(TAB_2) = False
        etabMain.TabEnabled(TAB_3) = False
        etabMain.TabEnabled(TAB_4) = False
        tgmJobClass.RemoveEditColumn COL_JOBCLASS
        tgmJobClass.ClearData
        tgmJobClass.AllowAddNew = False
        tblJobClass.Enabled = True
        subLoadJobClass
        subSetFocus tblJobClass
    Case TAB_2
        etabMain.TabEnabled(TAB_1) = False
        etabMain.TabEnabled(TAB_3) = False
        etabMain.TabEnabled(TAB_4) = False
        tgmPayCode.RemoveEditColumn COL_PAYCODE
        tgmPayCode.ClearData
        tblPayCode.Enabled = True
        subLoadPayCode
        tgmPayCode.AllowAddNew = False
        subSetFocus tblPayCode
    Case TAB_3
        etabMain.TabEnabled(TAB_1) = False
        etabMain.TabEnabled(TAB_2) = False
        etabMain.TabEnabled(TAB_4) = False
        tgmSpreadsheet.AllowAddNew = False
        tgmSpreadsheet.ClearData
        tgcDropdown.ComboActive(nComboForAdd) = False
        tgcDropdown.ComboActive(nComboForEdit) = True
        subEnableTop True
        cValidate.ResetFlags
        txtLevel.SetFocus
    Case TAB_4
        etabMain.TabEnabled(TAB_1) = False
        etabMain.TabEnabled(TAB_2) = False
        etabMain.TabEnabled(TAB_3) = False
        tgmGrid.ClearData
    End Select
    subEnableAddBtn False
    subEnableEditBtn False
    subEnableCancelBtn True
    cmdUpdateInsertBtn(etabMain.CurrTab).Caption = t_szCAPTION_UPDATE
    mnuUpdateInsert.Caption = t_szCAPTION_UPDATE
End Sub
Private Sub subLoadJobClass()
    Dim strSQL As String
    Dim nCount As Long
    Dim rsTemp As Recordset
    
    If nDataStatus = DATA_LOADED Then
        Exit Sub
    End If
    strSQL = "SELECT * FROM zzse_job_class ORDER BY zzsej_job_class"
    nCount = GetRecordSet(rsTemp, strSQL)
        If nCount <= 0 Then
            MsgBox "No record(s) found for editing for the given criteria"
            Exit Sub
        End If
    tgmJobClass.FillWithSQL t_dbMainDatabase, strSQL
    
    nDataStatus = DATA_LOADED
    subSetFocus tblJobClass
End Sub
Private Sub subLoadPayCode()
    Dim strSQL As String
    Dim nCount As Long
    Dim rsTemp As Recordset
    Dim aryPayCode() As Variant
    Dim i As Long
    
    If nDataStatus = DATA_LOADED Then
        Exit Sub
    End If
    strSQL = "SELECT * FROM zzse_pay_code ORDER BY zzsep_code"
    nCount = GetRecordSet(rsTemp, strSQL)
    If nCount <= 0 Then
        MsgBox "No record(s) found for editing for the given criteria"
        Exit Sub
    End If
    ReDim aryPayCode(1, nCount - 1)
    i = 0
    While Not rsTemp.EOF
        aryPayCode(COL_PAYCODE, i) = fnGetField(rsTemp!zzsep_code)
        aryPayCode(COL_CONVFACTOR, i) = fnGetField(rsTemp!zzsep_factor)
        i = i + 1
    rsTemp.MoveNext
    Wend
    tgmPayCode.FillWithArray aryPayCode
    'tgmPayCode.FillWithSQL t_dbMainDatabase, strSQL
    
    nDataStatus = DATA_LOADED
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
    Select Case etabMain.CurrTab
    Case TAB_1
        lRow = tgmJobClass.GetCurrentRowNumber
        If t_nFormMode = EDIT_MODE Then
            If fnGetField(tgmJobClass.CellValue(COL_JOBCLASS, lRow)) <> szEMPTY Then
                strSQL = " DELETE FROM zzse_job_class WHERE "
                strSQL = strSQL & " zzsej_job_class = " & tfnSQLString(fnGetField(tgmJobClass.CellValue(COL_JOBCLASS, lRow)))
                If Not fnExecuteSQL(strSQL, "cmdDeleteBtn_Click") Then
                    tfnSetStatusBarError "Error while deleting record."
                    Exit Sub
                End If
            End If
        End If
        tgmJobClass.DeleteRow
        tgmJobClass.Bookmark (lRow)
        subSetFocus tblJobClass

        If t_nFormMode = EDIT_MODE Then
            If tgmJobClass.RowCount < 1 Then
                tfnResetScreen
                Exit Sub
            End If
        Else
           subSetButtonStatus
        End If
    Case TAB_2
        lRow = tgmPayCode.GetCurrentRowNumber
        If t_nFormMode = EDIT_MODE Then
            If fnGetField(tgmPayCode.CellValue(COL_PAYCODE, lRow)) <> szEMPTY Then
                strSQL = " DELETE FROM zzse_pay_code WHERE "
                strSQL = strSQL & " zzsep_code = " & tfnSQLString(fnGetField(tgmPayCode.CellValue(COL_PAYCODE, lRow)))
                If Not fnExecuteSQL(strSQL, "cmdDeleteBtn_Click") Then
                    tfnSetStatusBarError "Error while deleting record."
                    Exit Sub
                End If
            End If
        End If
        tgmPayCode.DeleteRow
        tgmPayCode.Bookmark (lRow)
        subSetFocus tblPayCode

        If t_nFormMode = EDIT_MODE Then
            If tgmPayCode.RowCount < 1 Then
                tfnResetScreen
                Exit Sub
            End If
        Else
           subSetButtonStatus
        End If
    Case TAB_3
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
    End Select
    subEnableDeleteBtn False
End Sub

'Private Sub cmdOKBtn_Click(Index As Integer)
'    Select Case Index
'    Case TAB_3
'    Case TAB_4
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
        Case TAB_3
            If fnUpdateInsert3 Then
                tfnResetScreen
            End If
        
        Case TAB_4
            If fnUpdateInsert4 Then
                tfnResetScreen
            End If
    End Select
    
    txtProcessDate.Enabled = True
    cmdProcess(3).Enabled = True
    subEnableCancelBtn False, Index
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
    Select Case etabMain.CurrTab
    Case TAB_1
        subLoadJobClass
    Case TAB_2
        subLoadPayCode
    Case TAB_3
        subLoadPayIncrease
    End Select
    
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
        cmdAddBtn(etabMain.CurrTab).SetFocus
     Case TAB_3
        cmdAddBtn(etabMain.CurrTab).SetFocus
     Case TAB_4
        txtProcessDate.Enabled = True
        cmdProcess(3).Enabled = True
        subSetFocus txtProcessDate
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
'    cmdTBPrint_Click
    subPrintRpt
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
    cmdPrint(3).Enabled = False
    subEnablePrint False
  
    tblJobClass.Enabled = False
    tblPayCode.Enabled = False
    tblGrid.Enabled = False
    tblSpreadsheet.Enabled = False
    txtLevel.Enabled = False
    txtDesc.Enabled = False
    subEnableSearchButton cmdLevel, False
    subEnableSearchButton cmdDesc, False
    
    cmdUpdateInsertBtn(etabMain.CurrTab) = t_szCAPTION_UPDATE
    mnuUpdateInsert.Caption = t_szCAPTION_UPDATE
    
    frmContext.ButtonEnabled(PRINT_UP) = False
    frmContext.ButtonEnabled(COPY_UP) = False
    
    'clear data
    subClearAll
    
    nDataStatus = DATA_INIT
    If etabMain.CurrTab = TAB_1 Then
        subSetFocus cmdAddBtn(0)
    ElseIf etabMain.CurrTab = TAB_2 Then
        subSetFocus cmdAddBtn(1)
    ElseIf etabMain.CurrTab = TAB_2 Then
        subSetFocus cmdAddBtn(2)
    Else
        txtProcessDate.Enabled = True
        cmdProcess(3).Enabled = True
        'cmdProcess(3).SetFocus
        subSetFocus txtProcessDate
        tfnSetStatusBarMessage "Find employees who need pay raises"
        Exit Sub
    End If
    tfnSetStatusBarMessage t_szADD

End Sub


Private Sub tblFDropdown1_Click()

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

Private Sub tblPayCode_Click()
    tgfDropDown2.TableClick tblPayCode
End Sub

Private Sub tblPayCode_Scroll(CANCEL As Integer)
    tgfDropDown2.Scroll
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
    objText.ShowColumnExt
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
    cmdDeleteBtn(1).Enabled = bOnOff
    cmdDeleteBtn(2).Enabled = bOnOff
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
    cmdUpdateInsertBtn(2).Enabled = bOnOff
    cmdUpdateInsertBtn(3).Enabled = bOnOff
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
'        .AddButton "Print Report", PRFTCNTR_UP
'        .AddButton "Add &G/L Period/Series", GLPD_UP
'        .AddButton "Add Custome&r", CUSTOMER_UP
'        .AddButton "Add &Product", PRODUCT_UP
'        .AddButton "Add Ta&x Use Group", TAXUSE_UP
'        .AddButton "Add G/L &Account", GL_UP
        .EndSetupToolbar
    
        .HelpFile = szHelpFileName
    End With
End Sub
Private Sub subEnablePrint(bFlag As Boolean)
    frmContext.ButtonEnabled(PRINT_UP) = bFlag
    mnuPrint.Enabled = bFlag
End Sub
Public Sub TBButtonCallBack(ByVal nID As Integer)
    Select Case nID
        Case PRINT_UP
            subPrintRpt
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
        subLoadPayIncrease
    End If
End Sub

Private Sub txtDesc_KeyPress(KeyAscii As Integer)
    tgcDropdown.Keypress txtDesc, KeyAscii
'    If t_nFormMode = EDIT_MODE Then
        If KeyAscii = vbKeyReturn Then
            If tgcDropdown.SingleRecordSelected Then
                subLoadPayIncrease
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
                subLoadPayIncrease
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
        subLoadPayIncrease
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
                subLoadPayIncrease
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
            subLoadPayIncrease
        End If
    End If
End Sub
Private Sub subLoadPayIncrease()
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
    txtProcessDate = ""
    tgmJobClass.ClearData
    tgmJobClass.ResetFlags
    tgmPayCode.ClearData
    tgmPayCode.ResetFlags
    tgmGrid.ClearData
    tgmGrid.ResetFlags
    tgmSpreadsheet.ClearData
    tgmSpreadsheet.ResetFlags
    
End Sub
Private Sub subSetButtonStatus()
    Dim bStatus As Boolean
    
    Select Case etabMain.CurrTab
    Case TAB_1
        bStatus = (tgmJobClass.ValidData And (tgmJobClass.RowCount > 0))
    Case TAB_2
        bStatus = (tgmPayCode.ValidData And (tgmPayCode.RowCount > 0))
    Case TAB_3
        bStatus = (tgmSpreadsheet.ValidData And (tgmSpreadsheet.RowCount > 0))
    End Select
    
    
    If t_nFormMode = EDIT_MODE Then
        subEnableRefreshBtn True
    End If
        
    subEnableUpdateBtn bStatus
End Sub
Private Sub subCheckDelete()
    Dim lRow As Long
    Select Case etabMain.CurrTab
    Case TAB_1
        If tblJobClass.Row < 0 Then
            subEnableDeleteBtn False
        Else
            lRow = tgmJobClass.GetCurrentRowNumber
            If lRow < 0 Then
                subEnableDeleteBtn False
            Else
                If fnGetField(tgmJobClass.CellValue(COL_JOBCLASS, lRow)) <> "" Then
                    subEnableDeleteBtn True
                Else
                    subEnableDeleteBtn False
                End If
            End If
        End If
    Case TAB_2
        If tblPayCode.Row < 0 Then
            subEnableDeleteBtn False
        Else
            lRow = tgmPayCode.GetCurrentRowNumber
            If lRow < 0 Then
                subEnableDeleteBtn False
            Else
                If fnGetField(tgmPayCode.CellValue(COL_PAYCODE, lRow)) <> "" Then
                    subEnableDeleteBtn True
                Else
                    subEnableDeleteBtn False
                End If
            End If
        End If
    Case TAB_3
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
    End Select
End Sub
Public Sub subEnableAllControls(bStatus As Boolean)

    efraTabFrame1.Enabled = bStatus
    efraTabFrame2.Enabled = bStatus
    efraTabFrame3.Enabled = bStatus
    efraTabFrame4.Enabled = bStatus
    efraTabBack1.Enabled = bStatus
    efraTabBack2.Enabled = bStatus
    efraTabBack3.Enabled = bStatus
    efraTabBack4.Enabled = bStatus
    etabMain.TabEnabled(TAB_1) = bStatus
    etabMain.TabEnabled(TAB_2) = bStatus
    etabMain.TabEnabled(TAB_3) = bStatus
    etabMain.TabEnabled(TAB_4) = bStatus
    efraTop.Enabled = bStatus
    tgmJobClass.Locked = Not bStatus
    tgmPayCode.Locked = Not bStatus
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


Select Case tgTable.TabIndex
    Case tblJobClass.TabIndex
        fnValidCellValue = fnVerifyingJobClass(tgTable, nCurrCol, lCurrRow, szText)
    Case tblPayCode.TabIndex
        fnValidCellValue = fnVerifyingPayCode(tgTable, nCurrCol, lCurrRow, szText)
    Case tblSpreadsheet.TabIndex
        fnValidCellValue = fnVerifyingSpreadsheet(tgTable, nCurrCol, lCurrRow, szText)
End Select
End Function
Public Function fnVerifyingSpreadsheet(Table As TDBGrid, _
                                 ByVal nCol As Integer, _
                                 ByVal lRow As Long, _
                                 sText As String) As Boolean
    
    Const SUB_NAME As String = "fnVerifyingSpreadsheet"
    Dim strSQL As String
    Dim rsTemp As Recordset
    Dim sProdDesc As String
    Dim nRet As Integer
    Dim i As Long

    fnVerifyingSpreadsheet = True
    Select Case nCol
        Case COL_DAYS
            fnVerifyingSpreadsheet = fnValidNumberForAdd(sText, lRow)
        Case COL_RATE
            If IsNumeric(sText) Then
                fnVerifyingSpreadsheet = True
            ElseIf Trim(sText) = "" Then
                tgmSpreadsheet.ErrorMessage(nCol) = "Hourly Par Rate cannot be blank"
                fnVerifyingSpreadsheet = False
            ElseIf Trim(sText) = "" Then
                tgmSpreadsheet.ErrorMessage(nCol) = "Hourly Par Rate must be numeric"
                fnVerifyingSpreadsheet = False
            End If
    End Select
End Function
Public Function fnVerifyingJobClass(Table As TDBGrid, _
                                 ByVal nCol As Integer, _
                                 ByVal lRow As Long, _
                                 sText As String) As Boolean
    
    Const SUB_NAME As String = "fnVerifyingJobClass"
    Dim strSQL As String
    Dim rsTemp As Recordset
    Dim sProdDesc As String
    Dim nRet As Integer
    Dim i As Long

    fnVerifyingJobClass = True
    If t_nFormMode = EDIT_MODE Then
        Exit Function
    End If
    If nCol = COL_JOBCLASS Then
        If sText = "" Then
            tgmJobClass.ErrorMessage(nCol) = "You must enter a Job Class"
            fnVerifyingJobClass = False
        Else
            For i = 0 To tgmJobClass.RowCount - 1
                If i <> lRow Then
                    If fnGetField(tgmJobClass.CellValue(COL_JOBCLASS, i)) = fnGetField(sText) Then
                             tgmJobClass.ErrorMessage(nCol) = "Duplicate row(s) found on the screen"
                             fnVerifyingJobClass = False
                             Exit Function
                    End If
                End If
            Next
            strSQL = " SELECT prm_job_class FROM pr_master WHERE prm_job_class=" & tfnSQLString(sText)
            If GetRecordSet(rsTemp, strSQL, , SUB_NAME) <= 0 Then
                fnVerifyingJobClass = False
                tgmJobClass.ErrorMessage(0) = "This Job Class does not exist in the Payroll Master!"
                Exit Function
            End If
                
                strSQL = " SELECT prm_job_class FROM pr_master WHERE prm_job_class IN (SELECT zzsej_job_class FROM zzse_job_class"
                strSQL = strSQL & " WHERE  zzsej_job_class = " & tfnSQLString(sText) & ")"
                
                If GetRecordSet(rsTemp, strSQL, , SUB_NAME) > 0 Then
                    fnVerifyingJobClass = False
                    tgmJobClass.ErrorMessage(0) = "This Job Class has already been set up!"
                    Exit Function
                End If
        End If
    End If
End Function
Public Function fnVerifyingPayCode(Table As TDBGrid, _
                                 ByVal nCol As Integer, _
                                 ByVal lRow As Long, _
                                 sText As String) As Boolean
    
    Const SUB_NAME As String = "fnVerifyingPayCode"
    Dim strSQL As String
    Dim rsTemp As Recordset
    Dim sProdDesc As String
    Dim nRet As Integer
    Dim i As Long

    fnVerifyingPayCode = True
    If nCol = COL_PAYCODE Then
        If sText = "" Then
            tgmPayCode.ErrorMessage(nCol) = "You must enter a Pay Code"
            fnVerifyingPayCode = False
        Else
            strSQL = " SELECT prpa_pay_code FROM pr_pay"
                strSQL = strSQL & " WHERE  prpa_pay_code = " & tfnSQLString(sText)
                If GetRecordSet(rsTemp, strSQL, , SUB_NAME) <= 0 Then
                    fnVerifyingPayCode = False
                    tgmPayCode.ErrorMessage(0) = "This Pay Code is not set up in the Payroll System!"
                    Exit Function
                End If
            If t_nFormMode = ADD_MODE Then
                strSQL = " SELECT prpa_pay_code FROM pr_pay WHERE prpa_pay_code IN (SELECT zzsep_code FROM zzse_pay_code"
                strSQL = strSQL & " WHERE  zzsep_code = " & tfnSQLString(sText) & ")"
                If GetRecordSet(rsTemp, strSQL, , SUB_NAME) > 0 Then
                    fnVerifyingPayCode = False
                    tgmPayCode.ErrorMessage(0) = "This Pay Code has already been set up!"
                    Exit Function
                End If
            End If
            For i = 0 To tgmPayCode.RowCount - 1
                If i <> lRow Then
                    If fnGetField(tgmPayCode.CellValue(COL_PAYCODE, i)) = fnGetField(sText) Then
                             tgmPayCode.ErrorMessage(nCol) = "Duplicate row(s) found on the screen"
                             fnVerifyingPayCode = False
                             Exit Function
                    End If
                End If
            Next
        End If
    ElseIf nCol = COL_CONVFACTOR Then
        If sText = "" Then
            tgmPayCode.ErrorMessage(nCol) = "You must enter a Pay Code Conversion Factor."
            fnVerifyingPayCode = False
        End If
    End If
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
        Dim i As Integer
        For i = 0 To tgmSpreadsheet.RowCount - 1
            If Trim(tgmSpreadsheet.CellValue(COL_DAYS, i)) = sCode And i <> lCurrRow Then
                tgmSpreadsheet.ErrorMessage(COL_DAYS) = "Duplicated Days Employed encountered."
                Exit Function
            End If
        Next i
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
     Dim strSQL As String
     Dim lRowCount As Long
     Dim i As Long
     Dim sType As String
     
     lRowCount = tgmJobClass.RowCount
     
     For i = 0 To lRowCount - 1
        If t_nFormMode = ADD_MODE Then
            If fnGetField(tgmJobClass.CellValue(0, i)) <> "" Then
               strSQL = "INSERT INTO zzse_job_class(zzsej_job_class,zzsej_job_desc) VALUES "
               strSQL = strSQL & "(" & tfnSQLString(fnGetField(tgmJobClass.CellValue(COL_JOBCLASS, i)))
               strSQL = strSQL & "," & tfnSQLString(fnGetField(tgmJobClass.CellValue(COL_CONVFACTOR, i))) & ")"
                If Not fnExecuteSQL(strSQL, "fnUpdateInsert1") Then
                    Exit Function
                End If
            End If
        ElseIf t_nFormMode = EDIT_MODE Then
            If fnGetField(tgmJobClass.CellValue(COL_JOBCLASS, i)) <> "" Then 'unneccesary
                strSQL = "UPDATE zzse_job_class SET zzsej_job_desc = " & tfnSQLString(fnGetField(tgmJobClass.CellValue(COL_CONVFACTOR, i)))
                strSQL = strSQL & " WHERE zzsej_job_class = " & tfnSQLString(fnGetField(tgmJobClass.CellValue(COL_JOBCLASS, i)))
                If Not fnExecuteSQL(strSQL, "fnUpdateInsert1") Then
                    Exit Function
                End If
            End If
        End If
     Next
     fnUpdateInsert1 = True
End Function
Private Function fnUpdateInsert2() As Boolean
     Dim strSQL As String
     Dim lRowCount As Long
     Dim i As Long
     Dim sType As String
     
     lRowCount = tgmPayCode.RowCount
     
     For i = 0 To lRowCount - 1
        If t_nFormMode = ADD_MODE Then
            If fnGetField(tgmPayCode.CellValue(0, i)) <> "" Then
               strSQL = "INSERT INTO zzse_pay_code(zzsep_code,zzsep_factor) VALUES "
               strSQL = strSQL & "(" & tfnSQLString(fnGetField(tgmPayCode.CellValue(COL_PAYCODE, i)))
               strSQL = strSQL & "," & tfnRound(tgmPayCode.CellValue(COL_CONVFACTOR, i), DEFAULT_DECIMALS) & ")"
                If Not fnExecuteSQL(strSQL, , "fnUpdateInsert2") Then
                    Exit Function
                End If
            End If
        ElseIf t_nFormMode = EDIT_MODE Then
            If fnGetField(tgmPayCode.CellValue(COL_PAYCODE, i)) <> "" Then 'unneccesary
                strSQL = "UPDATE zzse_pay_code SET zzsep_factor = " & tfnRound(tgmPayCode.CellValue(COL_CONVFACTOR, i), DEFAULT_DECIMALS)
                strSQL = strSQL & " WHERE zzsep_code = " & tfnSQLString(fnGetField(tgmPayCode.CellValue(COL_PAYCODE, i)))
                If Not fnExecuteSQL(strSQL, , "fnUpdateInsert2") Then
                    Exit Function
                End If
            End If
        End If
     Next
     fnUpdateInsert2 = True
End Function
Private Function fnUpdateInsert3() As Boolean
    Const SUB_NAME = "fnUpdateInsert3"
    Dim strSQL As String
    Dim vData() As Variant
    Dim lRowCount As Long
    Dim nColCount As Integer
    Dim i As Long
    
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
    For i = 0 To lRowCount - 1
        If t_nFormMode = EDIT_MODE Then
            tgmSpreadsheet.GetRow vData, nColCount, i
        Else
            tgmSpreadsheet.GetChangedRow vData, nColCount, i
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
        
    fnUpdateInsert3 = True
    Screen.MousePointer = vbDefault
End Function
Private Function fnUpdateInsert4() As Boolean
    Const SUB_NAME = "fnUpdateInsert4"
    Dim strSQL As String
    Dim vData() As Variant
    Dim lRowCount As Long
    Dim nColCount As Integer
    Dim sOldEmp As String
    Dim i As Long
    Dim bRateChanged As Boolean
    
    Screen.MousePointer = vbHourglass

    'lRowCount = tgmGrid.GetChangedRowCount
    
    lRowCount = tgmGrid.RowCount
    
    sOldEmp = fnGetField(tgmGrid.CellValue(COL_EMPNO, 0))
    
    'insert true grid sceen data into database
    For i = 0 To lRowCount - 1
        'tgmGrid.GetChangedRow vData, nColCount, i
        tgmGrid.GetRow vData, nColCount, i
        
        If fnGetField(vData(COL_EMPNO)) <> "" Then
            strSQL = "UPDATE pr_emp_pay SET prep_amount= " _
              & tfnFormatDecimal(vData(COL_NEW_RATE), 2) ' * vData(nCOL_CONVT_FACTOR), 2)
            strSQL = strSQL & " WHERE prep_empno=" & tfnRound(vData(COL_EMPNO))
            strSQL = strSQL & " AND prep_pay_code= " & tfnSQLString(vData(COL_PCODE))
            strSQL = strSQL & " AND prep_year = " & tfnRound(vData(nCOL_PREP_YEAR))
            
            If Not fnExecuteSQL(strSQL, SUB_NAME) Then
                Screen.MousePointer = vbDefault
                Exit Function
            End If
        
        End If
        
        'Added by junsong 12/05/2002
        'This program is for 7_Eleven only, if Not, Check the sys_parm 30853
        'After updating the pr_emp_pay, we need to update pr_master also, to the
        'Pr_history and pr_pay_history can link together
        
        If fnGetField(vData(COL_EMPNO)) <> sOldEmp And bRateChanged Then
            strSQL = "UPDATE pr_master SET prm_emp_level = prm_emp_level"
            strSQL = strSQL & " WHERE prm_empno = " & tfnRound(sOldEmp)
            
            fnExecuteSQL strSQL, SUB_NAME
            sOldEmp = fnGetField(vData(COL_EMPNO))
            bRateChanged = False
        Else
        
            If tfnRound(vData(COL_NEW_RATE), 2) <> tfnRound(vData(COL_OLD_RATE), 2) Then
                bRateChanged = True
            End If
            
        End If
        
    Next
        
    If sOldEmp <> "" And bRateChanged Then
        strSQL = "UPDATE pr_master SET prm_emp_level = prm_emp_level"
        strSQL = strSQL & " WHERE prm_empno = " & tfnRound(sOldEmp)
        fnExecuteSQL strSQL, SUB_NAME
    End If
    
    fnUpdateInsert4 = True
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
    
    
'    strSQL = "SELECT DISTINCT prm_empno empno, prm_last_name lname,prm_first_name fname,prm_middle_name mname, prm_emp_level emp_level,zzsepi_lvl_desc emp_level_desc,prm_date_hired date_hired, "
'    strSQL = strSQL & " prep_amount old_rate, zzsepi_pay new_rate, zzsepi_days zzsepi_days"
'    strSQL = strSQL & " FROM pr_master,pr_emp_pay,zzse_pay_increase "
'    strSQL = strSQL & " WHERE prm_job_class='CS' AND (prm_date_termed IS NULL OR prm_date_termed ='')  AND prm_emp_level=zzsepi_emp_level "
'    strSQL = strSQL & " AND prm_empno=prep_empno AND prep_pay_code='" & strHourCode & "'"
'    strSQL = strSQL & " AND Today-prm_date_hired >0 "
'    strSQL = strSQL & " AND zzsepi_pay-prep_amount>0 "
'    strSQL = strSQL & " INTO TEMP tmpTableReport"
'
'    t_dbMainDatabase.ExecuteSQL strSQL
'
'    strSQL = "SELECT DISTINCT empno, lname,fname, mname, emp_level,emp_level_desc, date_hired,  "
'    strSQL = strSQL & " Today - date_hired days_worked "
'    strSQL = strSQL & ", old_rate, new_rate"
'    strSQL = strSQL & ", new_rate-old_rate rate_diff "
'    strSQL = strSQL & " FROM tmpTableReport "
'    strSQL = strSQL & " WHERE zzsepi_days = (SELECT MAX(b.zzsepi_days) FROM tmpTableReport AS b WHERE emp_level = b.emp_level AND empno = b.empno  AND b.zzsepi_days <= (Today - date_hired))"
'    strSQL = strSQL & " ORDER BY empno,emp_level"

    strSQL = "SELECT DISTINCT prm_empno empno, prm_last_name lname,prm_first_name fname,prm_middle_name mname,"
    strSQL = strSQL & " prm_job_class job_class,prm_emp_level emp_level,"
    strSQL = strSQL & " zzsepi_lvl_desc emp_level_desc,prm_date_hired date_hired, "
    strSQL = strSQL & " prep_amount old_rate, zzsepi_pay zzsepi_pay, zzsepi_days zzsepi_days,"
    strSQL = strSQL & " prep_pay_code pay_code, zzsep_factor conv_factor, prep_year prep_year,"
    strSQL = strSQL & " prep_amount prep_amount"
    strSQL = strSQL & " FROM pr_master,pr_emp_pay,zzse_pay_increase, "
    strSQL = strSQL & " zzse_job_class, zzse_pay_code"
    strSQL = strSQL & " WHERE prm_job_class= zzsej_job_class AND (prm_date_termed IS NULL OR prm_date_termed ='')  AND prm_emp_level=zzsepi_emp_level "
    strSQL = strSQL & " AND prm_empno=prep_empno AND zzsep_code = prep_pay_code"
'    strSQL = strSQL & " AND Today-prm_date_hired >0 "
    strSQL = strSQL & " AND prm_date_hired < " & tfnDateString(txtProcessDate, True)
    strSQL = strSQL & " AND (zzsepi_pay * zzsep_factor)-prep_amount>0 "
    strSQL = strSQL & " AND prep_year = ("
    strSQL = strSQL & " SELECT MAX(t2.prep_year) FROM pr_emp_pay t2 "
    strSQL = strSQL & " WHERE prep_empno = t2.prep_empno AND prep_pay_code = t2.prep_pay_code)"
    strSQL = strSQL & " INTO TEMP tmpTableReport"
    t_dbMainDatabase.ExecuteSQL strSQL
        
    strSQL = "SELECT DISTINCT empno, lname,fname, mname, job_class,emp_level,emp_level_desc, date_hired,  "
    'strSQL = strSQL & " (Today - date_hired) days_worked "
    'strSQL = strSQL & ", pay_code,old_rate"
    strSQL = strSQL & "  pay_code,old_rate"
    strSQL = strSQL & ", zzsepi_pay, conv_factor,prep_year "
    strSQL = strSQL & " FROM tmpTableReport AS a, zzse_pay_code"
    strSQL = strSQL & " WHERE a.zzsepi_days = (SELECT MAX(b.zzsepi_days) "
    strSQL = strSQL & " FROM tmpTableReport AS b WHERE a.emp_level = b.emp_level "
    'strSQL = strSQL & " AND a.empno = b.empno  AND b.zzsepi_days <= (Today - b.date_hired))"
    strSQL = strSQL & " AND a.empno = b.empno  AND (b.zzsepi_days + b.date_hired) <= " & tfnDateString(txtProcessDate, True) & ")"
    'AND (b.zzsepi_pay * b.conv_factor)-b.prep_amount>0)"
    strSQL = strSQL & " ORDER BY a.empno,a.emp_level"
    

    If GetRecordSet(rsTemp, strSQL, , SUB_NAME) < 0 Then
        MsgBox "Failed to access database to get records", vbCritical
        Exit Function
    End If
    If rsTemp.RecordCount = 0 Then
        MsgBox "No employee needs to increase pay.", vbExclamation
        Exit Function
    End If
    Dim strTemp As String
    'Dim vData() As Variant
    Dim nRowCount As Long
    Const GRID_COL_COUNT = 12
    Dim i As Long, j As Long
    Dim prev_empno As Long
    Dim new_empno As Long
    
    i = -1
    prev_empno = -999
    If rsTemp.RecordCount > 0 Then
        nRowCount = rsTemp.RecordCount
        'new_empno = fnGetField(rsTemp!empno)
        ReDim vData(GRID_COL_COUNT, nRowCount - 1)
        rsTemp.MoveFirst
        Do
            'If new_empno <> prev_empno Then
                i = i + 1
                vData(COL_EMPNO, i) = fnGetField(rsTemp!empno)
                vData(COL_JCLASS, i) = fnGetField(rsTemp!job_class)
                vData(COL_LEVEL, i) = fnGetField(rsTemp!emp_level)
                vData(COL_DESC, i) = fnGetField(rsTemp!emp_level_desc)
                vData(COL_HIRE_DATE, i) = rsTemp!date_hired
                'vData(COL_DAYS_WORKED, i) = fnGetField(rsTemp!days_worked)
                vData(COL_DAYS_WORKED, i) = CDate(txtProcessDate) - CDate(fnGetField(rsTemp!date_hired))
                vData(COL_PCODE, i) = fnGetField(rsTemp!pay_code)
                vData(COL_OLD_RATE, i) = tfnFormatDecimal(rsTemp!old_rate, 2)
               ' vData(COL_NEW_RATE, i) = tfnFormatDecimal(rsTemp!new_rate, 2)
               ' vData(COL_DIFF, i) = tfnFormatDecimal(rsTemp!rate_diff, 2)
                
                If (Trim(fnRSToString(rsTemp!fname)) = "") And (Trim(fnRSToString(rsTemp!mname)) = "") Then
                    strTemp = Trim(fnRSToString(rsTemp!lname))
                    vData(COL_NAME, i) = strTemp
                Else
                    strTemp = Trim(fnRSToString(rsTemp!lname)) & ", " & Trim(fnRSToString(rsTemp!fname)) & " " & Trim(fnRSToString(rsTemp!mname))
                    vData(COL_NAME, i) = strTemp
                End If
                vData(nCOL_CONVT_FACTOR, i) = tfnFormatDecimal(rsTemp!conv_factor, 2)
                vData(nCOL_PREP_YEAR, i) = tfnRound(rsTemp!prep_year)
                vData(COL_NEW_RATE, i) = tfnFormatDecimal(rsTemp!zzsepi_pay, 2) * tfnFormatDecimal(rsTemp!conv_factor, 2)
                vData(COL_DIFF, i) = vData(COL_NEW_RATE, i) - vData(COL_OLD_RATE, i)
'            Else
'                vData(COL_PCODE, i) = fnGetField(rsTemp!pay_code)
'                vData(COL_OLD_RATE, i) = tfnFormatDecimal(rsTemp!old_rate, 2)
'                vData(COL_NEW_RATE, i) = tfnFormatDecimal(rsTemp!new_rate, 2)
'                vData(COL_DIFF, i) = tfnFormatDecimal(rsTemp!rate_diff, 2)
'            End If
            rsTemp.MoveNext
            
        Loop Until rsTemp.EOF
    
    
    End If
    tgmGrid.FillWithArray vData
    
    fnProcess = True
End Function
Private Function fnRSToString(ByVal vTemp As Variant) As String
    If IsNull(vTemp) Then
        fnRSToString = ""
    Else
        fnRSToString = vTemp
    End If
End Function
Private Function fnBuildEmpTrailer() As String
    Dim i As Long
    Dim j As Long
    Dim prev_emp As String
    Dim curr_emp As String
    
    prev_emp = "-9999"
    j = 0
    For i = 0 To tgmGrid.RowCount - 1
        curr_emp = fnGetField(tgmGrid.CellValue(COL_EMPNO, i))
        If prev_emp <> curr_emp Then
            j = j + 1
        End If
        prev_emp = curr_emp
    Next
    fnBuildEmpTrailer = j & " Employee(s) Changed"
End Function
Private Function fnWriteLog(vData As Variant, sProgName As String, sRptTitle As String, szTrailer As String, sLineHeader As String, Optional sFileName As String = "ZZSEPILG", Optional bAppend As Boolean = False) As Boolean
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
    Dim i As Integer
    i = 0
    'Print #nFilePtr, "Program ID: ZZSEPILG"
    Print #nFilePtr, sProgName
    Print #nFilePtr, sRptTitle
    Print #nFilePtr, "Run Date: ", Format(Date, "mm/dd/yyyy")
    Print #nFilePtr, "Run Time: ", Format(Time, "hh:mm:ss AMPM")
    Print #nFilePtr, String(CHAR_PER_LINE, "-")
    Print #nFilePtr, sLineHeader
    
    Do While i <= INDEX_REPORT
        'print title and time
        Print #nFilePtr, vData(i)
        i = i + 1
    Loop
    szTrailer = fnBuildEmpTrailer
    Print #nFilePtr, szTrailer
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
Private Function fnCheckPayCode() As Boolean
    Dim strSQL As String
    Dim rsTemp As Recordset
    Const SUB_NAME As String = "fnCheckPayCode"
    
    fnCheckPayCode = False
    
    strSQL = "SELECT * FROM zzse_pay_code"
    If GetRecordSet(rsTemp, strSQL, , SUB_NAME) < 0 Then
        MsgBox "Failed to access database to get records", vbCritical
        Exit Function
    End If
    If rsTemp.RecordCount = 0 Then
        MsgBox "No Pay Codes have been set up in the Automatic Pay Increase program.", vbExclamation
        Exit Function
    End If
    
    fnCheckPayCode = True
End Function
Private Function fnCheckJobClass() As Boolean
    Dim strSQL As String
    Dim rsTemp As Recordset
    Const SUB_NAME As String = "fnCheckJobClass"
    
    fnCheckJobClass = False
    
    strSQL = "SELECT * FROM zzse_job_class"
    If GetRecordSet(rsTemp, strSQL, , SUB_NAME) < 0 Then
        MsgBox "Failed to access database to get records", vbCritical
        Exit Function
    End If
    If rsTemp.RecordCount = 0 Then
        MsgBox "No Job Class have been set up in the Automatic Pay Increase program.", vbExclamation
        Exit Function
    End If
    
    fnCheckJobClass = True
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
    
    nCOL_CONVT_FACTOR = tgmGrid.AddHiddenField("zzsep_factor")
    nCOL_PREP_YEAR = tgmGrid.AddHiddenField("prep_year")
    nWidth = tblGrid.Width
'    With tblGrid.Columns(COL_EMPNO)
'        .Width = 0.12 * nWidth
'        .DataField = "empno"
'    End With
'    With tblGrid.Columns(COL_NAME)
'        .Width = 0.12 * nWidth
'        .DataField = "lname"
'    End With
'    With tblGrid.Columns(COL_JOBCLASS)
'        .Width = 0.03 * nWidth
'        .DataField = "job_class"
'    End With
'    With tblGrid.Columns(COL_LEVEL)
'        .Width = 0.05 * nWidth
'        .DataField = "emp_level"
'    End With
'    With tblGrid.Columns(COL_DESC)
'        .Width = 0.14 * nWidth
'        .DataField = "emp_level_desc"
'    End With
'    With tblGrid.Columns(COL_HIRE_DATE)
'        .Width = 0.13 * nWidth
'        .DataField = "date_hired"
'    End With
'    With tblGrid.Columns(COL_DAYS_WORKED)
'        .Width = 0.08 * nWidth
'        .DataField = "days_worked"
'    End With
'    With tblGrid.Columns(COL_PAYCODE)
'        .Width = 0.06 * nWidth
'        .DataField = "pay_code"
'    End With
'    With tblGrid.Columns(COL_OLD_RATE)
'        .Width = 0.07 * nWidth
'        .DataField = "old_rate"
'    End With
'    With tblGrid.Columns(COL_NEW_RATE)
'        .Width = 0.07 * nWidth
'        .DataField = "new_rate"
'    End With
'    With tblGrid.Columns(COL_DIFF)
'        .Width = 0.01 * nWidth
'        .DataField = "rate_diff"
'    End With
    tgmGrid.DisplayFormat(COL_OLD_RATE) = sFmtAmount
    tgmGrid.DisplayFormat(COL_NEW_RATE) = sFmtAmount
    tgmGrid.DisplayFormat(COL_DIFF) = sFmtAmount
    
    tgmGrid.SetupTable True
    tgmGrid.AllowAddNew = False
    tgmGrid.ClearData

End Sub
'initialization of Extension for grid tool tips
Private Sub subInitText()

   Set objText = New clsColumnExtension
   With objText
       Set .Form = Me
       Set .PictureBox = picText
       Set .StatusBar = ffraStatusbar
       Set .Table = tblGrid
   End With
   objText.AddColumn COL_EMPNO, "Employee Number"
   objText.AddColumn COL_NAME, "Employee Name"
   objText.AddColumn COL_DESC, "Employee Level Description"
'   objText.AddColumn COL_AcctDesc, "Gl Account Description"
'   objText.AddColumn COL_Amount, "Amount"
End Sub
Public Sub subPrintRpt()
    Dim sProgName As String
    Dim sRptTitle As String
    Dim sLineHeader As String
    Dim sExtraLine As String
    
    Dim lCount As Long
    Dim i As Long
    Dim sLine As String
    
    sProgName = "Program ID: ZZSEPIMT"
    sRptTitle = "7-Eleven Automatic Pay Increase Processing Log"
    
    sLineHeader = fnTranc("Emp#", 10, LEFT_JUSTIFY) & _
                  fnTranc("Name", 26, LEFT_JUSTIFY) & _
                  fnTranc("JCls", 5, LEFT_JUSTIFY) & _
                  fnTranc("Level", 6, LEFT_JUSTIFY) & _
                  fnTranc("Description", 21, LEFT_JUSTIFY) & _
                  fnTranc("Hire Date", 15, LEFT_JUSTIFY) & _
                  fnTranc("Days Wkd", 10, RIGHT_JUSTIFY) & _
                  fnTranc("PCode", 6, RIGHT_JUSTIFY) & _
                  fnTranc("Old Rate", 12, RIGHT_JUSTIFY) & _
                  fnTranc("New Rate", 12, RIGHT_JUSTIFY) & _
                  fnTranc("Rate Diff", 12, RIGHT_JUSTIFY)
        

    INDEX_REPORT = -1
    
    lCount = frmZZSEPIMT.tgmGrid.RowCount
    
    If lCount <= 0 Then
        Exit Sub
    End If
    
    
    For i = 0 To lCount - 1
        With tgmGrid
           sLine = fnTranc(.CellValue(COL_EMPNO, i), 10, LEFT_JUSTIFY) & _
              fnTranc(.CellValue(COL_NAME, i), 26, LEFT_JUSTIFY) & _
              fnTranc(.CellValue(COL_JCLASS, i), 5, LEFT_JUSTIFY) & _
              fnTranc(.CellValue(COL_LEVEL, i), 6, LEFT_JUSTIFY) & _
              fnTranc(.CellValue(COL_DESC, i), 21, LEFT_JUSTIFY) & _
              fnTranc(.CellValue(COL_HIRE_DATE, i), 15, LEFT_JUSTIFY) & _
              fnTranc(.CellValue(COL_DAYS_WORKED, i), 10, RIGHT_JUSTIFY) & _
              fnTranc(.CellValue(COL_PCODE, i), 6, RIGHT_JUSTIFY) & _
              fnTranc(.CellValue(COL_OLD_RATE, i), 12, RIGHT_JUSTIFY) & _
              fnTranc(tfnFormatDecimal(.CellValue(COL_NEW_RATE, i), 2), 12, RIGHT_JUSTIFY) & _
              fnTranc(tfnFormatDecimal(.CellValue(COL_DIFF, i), 2), 12, RIGHT_JUSTIFY)
          'put in the array report
          subLineFeed sLine
       End With
    Next

    'write log file
    fnWriteLog ary_REPORT, sProgName, sRptTitle, sExtraLine, sLineHeader
    If lCount > 0 Then
        subLineFeed String(CHAR_PER_LINE, "-")
        subLineFeed sExtraLine
    End If
    sExtraLine = szEMPTY
    
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
Private Sub tbljobclass_AfterColEdit(ByVal ColIndex As Integer)
    tgmJobClass.AfterColEdit ColIndex
    nDataStatus = DATA_CHANGED
    t_bDataChanged = True
    subSetButtonStatus
End Sub

Private Sub tbljobclass_BeforeColEdit(ByVal ColIndex As Integer, ByVal KeyAscii As Integer, CANCEL As Integer)
    subEnableUpdateBtn False
    tgmJobClass.BeforeColEdit ColIndex, KeyAscii, CANCEL
End Sub

Private Sub tbljobclass_Change()
    tgmJobClass.Change
End Sub

Private Sub tbljobclass_FirstRowChange()
    tgmJobClass.FirstRowChange
    'tgfDropDown1.FirstRowChange
    subCheckDelete
End Sub

Private Sub tbljobclass_GotFocus()
    tgmJobClass.GotFocus
    'tgfDropDown1.GotFocus
End Sub

Private Sub tbljobclass_KeyDown(KeyCode As Integer, Shift As Integer)
    tgmJobClass.KeyDown KeyCode, Shift
End Sub

Private Sub tbljobclass_KeyPress(KeyAscii As Integer)
    If Not tgmJobClass.Keypress(KeyAscii) Then
        KeyAscii = 0
    End If
    If tblJobClass.col = COL_JOBCLASS Then
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If
'    If KeyAscii <> vbKeyReturn Then
'        If tbljobclass.col = COL_DAYS Then
'            tgmjobclass.CellValue(COL_RATE, tgmjobclass.GetCurrentRowNumber) = ""
'
'        End If
'    End If
End Sub

Private Sub tbljobclass_LostFocus()
    tgmJobClass.LostFocus
     'tgfDropDown1.LostFocus tblJobClass
End Sub

Private Sub tbljobclass_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    'If tgfDropDown1 Is Nothing Then
     '   Exit Sub
    'End If
    'If tgfDropDown1.RowColChange(tblJobClass) Then
        tgmJobClass.RowColChange LastRow, LastCol
    'End If
    
    subCheckDelete
End Sub

Private Sub tbljobclass_SelChange(CANCEL As Integer)
     CANCEL = True
End Sub

Private Sub tbljobclass_UnboundReadData(ByVal RowBuf As DBTrueGrid.RowBuffer, StartLocation As Variant, ByVal ReadPriorRows As Boolean)
    tgmJobClass.ReadData RowBuf, StartLocation, ReadPriorRows
End Sub

Private Sub tblpaycode_AfterColEdit(ByVal ColIndex As Integer)
    tgmPayCode.AfterColEdit ColIndex
    nDataStatus = DATA_CHANGED
    t_bDataChanged = True
    subSetButtonStatus
End Sub

Private Sub tblpaycode_BeforeColEdit(ByVal ColIndex As Integer, ByVal KeyAscii As Integer, CANCEL As Integer)
    subEnableUpdateBtn False
    tgmPayCode.BeforeColEdit ColIndex, KeyAscii, CANCEL
End Sub

Private Sub tblpaycode_Change()
    tgmPayCode.Change
End Sub

Private Sub tblpaycode_FirstRowChange()
    tgmPayCode.FirstRowChange
    tgfDropDown2.FirstRowChange
    subCheckDelete
End Sub

Private Sub tblpaycode_GotFocus()
    tgfDropDown2.GotFocus
    tgmPayCode.GotFocus
    
    If tgfDropDown2.ValidSelection Then
         subSetButtonStatus
    End If
    
End Sub

Private Sub tblpaycode_KeyDown(KeyCode As Integer, Shift As Integer)
    tgmPayCode.KeyDown KeyCode, Shift
End Sub

Private Sub tblpaycode_KeyPress(KeyAscii As Integer)
    If Not tgmPayCode.Keypress(KeyAscii) Then
        KeyAscii = 0
    End If
    If tblPayCode.col = COL_PAYCODE Then
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If
    tgfDropDown2.Keypress tblPayCode, KeyAscii
'    If KeyAscii <> vbKeyReturn Then
'        If tblpaycode.col = COL_DAYS Then
'            tgmpaycode.CellValue(COL_RATE, tgmpaycode.GetCurrentRowNumber) = ""
'
'        End If
'    End If
End Sub

Private Sub tblpaycode_LostFocus()
    tgmPayCode.LostFocus
    'tblpaycode_AfterColEdit COL_PAYCODE
     tgfDropDown2.LostFocus tblPayCode
     
End Sub

Private Sub tblpaycode_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If tgfDropDown2 Is Nothing Then
        Exit Sub
    End If
    If tgfDropDown2.RowColChange(tblPayCode) Then
        tgmPayCode.RowColChange LastRow, LastCol
    End If
'    If tgmPayCode.ValidCell(COL_PAYCODE, tgmGrid.GetCurrentRowNumber) Then
'    End If
    subCheckDelete
End Sub

Private Sub tblpaycode_SelChange(CANCEL As Integer)
     CANCEL = True
End Sub

Private Sub tblpaycode_UnboundReadData(ByVal RowBuf As DBTrueGrid.RowBuffer, StartLocation As Variant, ByVal ReadPriorRows As Boolean)
    tgmPayCode.ReadData RowBuf, StartLocation, ReadPriorRows
End Sub


Private Sub tblFDropDown2_Click()
    tgfDropDown2.TableClick tblFDropdown2
End Sub

Private Sub tblFDropDown2_KeyPress(KeyAscii As Integer)
    tgfDropDown2.Keypress tblFDropdown2, KeyAscii
End Sub

'***********************************************************************
Private Sub tblFDropDown2_Keydown(KeyCode As Integer, Shift As Integer)
    tgfDropDown2.KeyDown tblFDropdown2, KeyCode
End Sub
'***********************************************************************
Private Sub tblFDropDown2_LostFocus()
    tgfDropDown2.LostFocus tblFDropdown2
End Sub

Private Sub tblFDropDown2_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    tgfDropDown2.MouseUp y
End Sub

Private Sub tblFDropDown2_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    tgfDropDown2.RowColChange tblFDropdown2
End Sub

Private Sub txtProcessDate_GotFocus()
    tfnRegExpControlGotFocus txtProcessDate, szDatePattern
    tfnSetStatusBarMessage "Enter the date to use for calculating the pay increase"
    txtProcessDate.SelStart = 0
    txtProcessDate.SelLength = Len(txtProcessDate)
End Sub

Private Sub txtProcessDate_KeyPress(KeyAscii As Integer)
    tfnRegExpControlDateKeyPress txtProcessDate, KeyAscii
    
    If KeyAscii = vbKeyReturn Then
        subSetFocus cmdProcess(3)
        KeyAscii = 0
    End If
    
End Sub

Private Sub txtProcessDate_LostFocus()
    
    If IsFactorDate(txtProcessDate) Then
        txtProcessDate = tfnFormatDate(txtProcessDate)
    End If
    
End Sub

Private Function IsFactorDate(ByVal sDate As String) As Boolean
   
    If SRegExpMatch(szDatePattern, sDate) = 0 Then
        sDate = tfnDateString(tfnFormatDate(sDate))
      
        If IsDate(sDate) Then
            IsFactorDate = True
        Else
            IsFactorDate = False
        End If
        
    Else
        IsFactorDate = False
    End If
   
End Function

