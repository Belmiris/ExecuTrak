VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C75015E0-2232-11D3-B440-0060971E99AF}#1.0#0"; "FACTFRM.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form LogForm 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Flat File Processor"
   ClientHeight    =   6036
   ClientLeft      =   876
   ClientTop       =   2076
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
   ScaleHeight     =   6036
   ScaleWidth      =   8880
   Begin FACTFRMLib.FactorFrame ffraStatusbar 
      Height          =   360
      Left            =   0
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   5676
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
      TabIndex        =   3
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
      FMName          =   "FFPROC"
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
         TabIndex        =   5
         Top             =   84
         Width           =   6876
         _ExtentX        =   12129
         _ExtentY        =   656
         ButtonWidth     =   614
         ButtonHeight    =   572
         _Version        =   393216
      End
   End
   Begin FACTFRMLib.FactorFrame ffBackground2 
      Height          =   5196
      Left            =   12
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   480
      Width           =   8880
      _Version        =   65536
      _ExtentX        =   15663
      _ExtentY        =   9165
      _StockProps     =   77
      BackColor       =   8388608
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   7.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelOuter      =   0
      Align           =   5
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
      Begin FACTFRMLib.FactorFrame cmdClose 
         Height          =   396
         HelpContextID   =   15
         Left            =   7464
         TabIndex        =   16
         Top             =   4716
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
         Caption         =   "&Close"
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
      Begin FACTFRMLib.FactorFrame cmdOK 
         Height          =   396
         HelpContextID   =   3005
         Left            =   6048
         TabIndex        =   15
         Top             =   4716
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
      Begin FACTFRMLib.FactorFrame ffSetupInOut 
         Height          =   4536
         Left            =   108
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   96
         Width           =   8652
         _Version        =   65536
         _ExtentX        =   15261
         _ExtentY        =   8001
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
         Begin FACTFRMLib.FactorFrame ffBacks 
            Height          =   2196
            Index           =   6
            Left            =   84
            TabIndex        =   34
            TabStop         =   0   'False
            Top             =   2280
            Width           =   8460
            _Version        =   65536
            _ExtentX        =   14922
            _ExtentY        =   3873
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
            Begin FACTFRMLib.FactorFrame ffIOInput 
               Height          =   840
               Left            =   84
               TabIndex        =   47
               TabStop         =   0   'False
               Top             =   60
               Width           =   732
               _Version        =   65536
               _ExtentX        =   1291
               _ExtentY        =   1482
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
               BevelOuter      =   0
               BeginProperty PanelFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Begin VB.Label lblCaptions 
                  Caption         =   "File"
                  Height          =   372
                  Index           =   17
                  Left            =   24
                  TabIndex        =   49
                  Top             =   312
                  Width           =   696
               End
               Begin VB.Label lblCaptions 
                  Caption         =   "Input"
                  Height          =   288
                  Index           =   13
                  Left            =   24
                  TabIndex        =   48
                  Top             =   48
                  Width           =   696
               End
            End
            Begin FACTFRMLib.FactorFrame FactorFrame3 
               Height          =   876
               Left            =   36
               TabIndex        =   42
               TabStop         =   0   'False
               Top             =   1224
               Width           =   8328
               _Version        =   65536
               _ExtentX        =   14690
               _ExtentY        =   1545
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
               Begin VB.TextBox txtBackupPath 
                  Height          =   372
                  HelpContextID   =   2001
                  Left            =   2052
                  TabIndex        =   31
                  Top             =   72
                  Width           =   5112
               End
               Begin VB.TextBox txtBackupName 
                  Height          =   372
                  Left            =   2052
                  TabIndex        =   32
                  Top             =   492
                  Width           =   3000
               End
               Begin FACTFRMLib.FactorFrame cmdBackupPath 
                  Height          =   372
                  HelpContextID   =   2002
                  Left            =   7200
                  TabIndex        =   43
                  TabStop         =   0   'False
                  Top             =   72
                  Width           =   1068
                  _Version        =   65536
                  _ExtentX        =   1884
                  _ExtentY        =   656
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
                  Caption         =   "Browse ..."
                  CaptionPos      =   4
                  PicturePos      =   3
                  ShowFocusRect   =   -1  'True
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
               Begin VB.Line Line1 
                  Index           =   1
                  X1              =   0
                  X2              =   8364
                  Y1              =   0
                  Y2              =   0
               End
               Begin VB.Line Line1 
                  BorderColor     =   &H80000005&
                  Index           =   0
                  X1              =   -12
                  X2              =   8352
                  Y1              =   12
                  Y2              =   12
               End
               Begin VB.Label lblCaptions 
                  Caption         =   "Backup"
                  Height          =   288
                  Index           =   22
                  Left            =   72
                  TabIndex        =   46
                  Top             =   48
                  Width           =   732
               End
               Begin VB.Label lblCaptions 
                  Caption         =   "File Path"
                  Height          =   288
                  Index           =   21
                  Left            =   912
                  TabIndex        =   45
                  Top             =   120
                  Width           =   1200
               End
               Begin VB.Label lblCaptions 
                  Caption         =   "File Name"
                  Height          =   336
                  Index           =   16
                  Left            =   912
                  TabIndex        =   44
                  Top             =   552
                  Width           =   1200
               End
            End
            Begin VB.TextBox txtFilePath 
               Height          =   372
               HelpContextID   =   2001
               Left            =   2076
               TabIndex        =   25
               Top             =   60
               Width           =   5112
            End
            Begin VB.TextBox txtFileName 
               Height          =   372
               Left            =   2076
               TabIndex        =   26
               Top             =   480
               Width           =   3000
            End
            Begin FACTFRMLib.FactorFrame ffBacks 
               Height          =   312
               Index           =   7
               Left            =   792
               TabIndex        =   35
               TabStop         =   0   'False
               Top             =   864
               Width           =   3900
               _Version        =   65536
               _ExtentX        =   6879
               _ExtentY        =   550
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
               BevelOuter      =   0
               BeginProperty PanelFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Begin VB.OptionButton optBKFInput 
                  Caption         =   "No"
                  Height          =   240
                  Index           =   1
                  Left            =   2424
                  TabIndex        =   28
                  Top             =   72
                  Width           =   768
               End
               Begin VB.OptionButton optBKFInput 
                  Caption         =   "Yes"
                  Height          =   240
                  Index           =   0
                  Left            =   1476
                  TabIndex        =   27
                  Top             =   72
                  Width           =   768
               End
               Begin VB.Label lblCaptions 
                  Caption         =   "Backup File"
                  Height          =   264
                  Index           =   9
                  Left            =   96
                  TabIndex        =   36
                  Top             =   60
                  Width           =   1392
               End
            End
            Begin FACTFRMLib.FactorFrame cmdInputPath 
               Height          =   372
               HelpContextID   =   2002
               Left            =   7236
               TabIndex        =   37
               TabStop         =   0   'False
               Top             =   60
               Width           =   1068
               _Version        =   65536
               _ExtentX        =   1884
               _ExtentY        =   656
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
               Caption         =   "Browse ..."
               CaptionPos      =   4
               PicturePos      =   3
               ShowFocusRect   =   -1  'True
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
            Begin FACTFRMLib.FactorFrame ffBacks 
               Height          =   312
               Index           =   8
               Left            =   4692
               TabIndex        =   38
               TabStop         =   0   'False
               Top             =   864
               Width           =   3648
               _Version        =   65536
               _ExtentX        =   6435
               _ExtentY        =   550
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
               BevelOuter      =   0
               BeginProperty PanelFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Begin VB.OptionButton optRMFInput 
                  Caption         =   "No"
                  Height          =   240
                  Index           =   1
                  Left            =   2616
                  TabIndex        =   30
                  Top             =   72
                  Width           =   768
               End
               Begin VB.OptionButton optRMFInput 
                  Caption         =   "Yes"
                  Height          =   240
                  Index           =   0
                  Left            =   1536
                  TabIndex        =   29
                  Top             =   72
                  Width           =   768
               End
               Begin VB.Label lblCaptions 
                  Caption         =   "Remove File"
                  Height          =   264
                  Index           =   10
                  Left            =   84
                  TabIndex        =   39
                  Top             =   60
                  Width           =   1776
               End
            End
            Begin VB.Label lblCaptions 
               Caption         =   "File Path"
               Height          =   288
               Index           =   12
               Left            =   900
               TabIndex        =   41
               Top             =   108
               Width           =   1200
            End
            Begin VB.Label lblCaptions 
               Caption         =   "File Name"
               Height          =   336
               Index           =   11
               Left            =   900
               TabIndex        =   40
               Top             =   540
               Width           =   1200
            End
         End
         Begin FACTFRMLib.FactorFrame ffBacks 
            Height          =   780
            Index           =   12
            Left            =   84
            TabIndex        =   50
            TabStop         =   0   'False
            Top             =   48
            Width           =   2340
            _Version        =   65536
            _ExtentX        =   4128
            _ExtentY        =   1376
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
            Begin VB.OptionButton optRunModeIO 
               Caption         =   "Auto"
               Height          =   240
               Index           =   1
               Left            =   1236
               TabIndex        =   19
               Top             =   360
               Width           =   1044
            End
            Begin VB.OptionButton optRunModeIO 
               Caption         =   "Manual"
               Height          =   240
               Index           =   0
               Left            =   72
               TabIndex        =   18
               Top             =   360
               Width           =   1224
            End
            Begin VB.Label lblCaptions 
               Caption         =   "Start Mode"
               Height          =   264
               Index           =   24
               Left            =   84
               TabIndex        =   51
               Top             =   72
               Width           =   1572
            End
         End
         Begin FACTFRMLib.FactorFrame ffBacks 
            Height          =   1332
            Index           =   5
            Left            =   84
            TabIndex        =   52
            TabStop         =   0   'False
            Top             =   888
            Width           =   8448
            _Version        =   65536
            _ExtentX        =   14901
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
            Begin VB.TextBox txtLogFile 
               Height          =   384
               Left            =   2280
               TabIndex        =   24
               Top             =   852
               Width           =   3156
            End
            Begin VB.TextBox txtLogPath 
               Height          =   384
               HelpContextID   =   2001
               Left            =   2280
               TabIndex        =   23
               Top             =   384
               Width           =   4908
            End
            Begin FACTFRMLib.FactorFrame ffBacks 
               Height          =   672
               Index           =   0
               Left            =   72
               TabIndex        =   53
               TabStop         =   0   'False
               Top             =   84
               Width           =   1836
               _Version        =   65536
               _ExtentX        =   3238
               _ExtentY        =   1185
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
               Begin VB.OptionButton optWriteLog 
                  Caption         =   "No"
                  Height          =   240
                  Index           =   1
                  Left            =   912
                  TabIndex        =   22
                  Top             =   348
                  Width           =   888
               End
               Begin VB.OptionButton optWriteLog 
                  Caption         =   "Yes"
                  Height          =   240
                  Index           =   0
                  Left            =   60
                  TabIndex        =   21
                  Top             =   348
                  Width           =   888
               End
               Begin VB.Label lblCaptions 
                  Caption         =   "Write Log"
                  Height          =   264
                  Index           =   1
                  Left            =   84
                  TabIndex        =   54
                  Top             =   36
                  Width           =   1212
               End
            End
            Begin FACTFRMLib.FactorFrame cmdLogPath 
               Height          =   384
               HelpContextID   =   2002
               Left            =   7236
               TabIndex        =   55
               TabStop         =   0   'False
               Top             =   384
               Width           =   1068
               _Version        =   65536
               _ExtentX        =   1884
               _ExtentY        =   677
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
               Caption         =   "Browse ..."
               CaptionPos      =   4
               PicturePos      =   3
               ShowFocusRect   =   -1  'True
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
            Begin VB.Label lblCaptions 
               Caption         =   "Log File Name"
               Height          =   336
               Index           =   2
               Left            =   384
               TabIndex        =   57
               Top             =   912
               Width           =   1704
            End
            Begin VB.Label lblCaptions 
               Caption         =   "Log File Path"
               Height          =   288
               Index           =   3
               Left            =   2280
               TabIndex        =   56
               Top             =   96
               Width           =   1608
            End
         End
         Begin FACTFRMLib.FactorFrame ffBacks 
            Height          =   780
            Index           =   3
            Left            =   2532
            TabIndex        =   58
            TabStop         =   0   'False
            Top             =   48
            Width           =   6000
            _Version        =   65536
            _ExtentX        =   10583
            _ExtentY        =   1376
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
            Begin VB.TextBox txtWorkPath 
               Height          =   384
               HelpContextID   =   2001
               Left            =   120
               TabIndex        =   20
               Top             =   324
               Width           =   4656
            End
            Begin FACTFRMLib.FactorFrame cmdWorkPath 
               Height          =   384
               HelpContextID   =   2002
               Left            =   4824
               TabIndex        =   59
               TabStop         =   0   'False
               Top             =   324
               Width           =   1068
               _Version        =   65536
               _ExtentX        =   1884
               _ExtentY        =   677
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
               Caption         =   "Browse ..."
               CaptionPos      =   4
               PicturePos      =   3
               ShowFocusRect   =   -1  'True
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
            Begin VB.Label lblCaptions 
               Caption         =   "Working Path"
               Height          =   288
               Index           =   0
               Left            =   132
               TabIndex        =   60
               Top             =   48
               Width           =   1608
            End
         End
      End
   End
   Begin FACTFRMLib.FactorFrame ffBackground1 
      Height          =   5184
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   480
      Width           =   8880
      _Version        =   65536
      _ExtentX        =   15663
      _ExtentY        =   9144
      _StockProps     =   77
      BackColor       =   8388608
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   7.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelOuter      =   0
      Align           =   5
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
      Begin FACTFRMLib.FactorFrame cmdProcess 
         Height          =   396
         HelpContextID   =   3005
         Left            =   6036
         TabIndex        =   1
         Top             =   4716
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
      Begin FACTFRMLib.FactorFrame cmdExitCancelBtn 
         Height          =   396
         HelpContextID   =   15
         Left            =   7464
         TabIndex        =   2
         Top             =   4716
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
         Caption         =   "&Exit"
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
      Begin MSComDlg.CommonDialog dlgFilenames 
         Left            =   3780
         Top             =   4692
         _ExtentX        =   699
         _ExtentY        =   699
         _Version        =   393216
      End
      Begin VB.FileListBox lstFile 
         Height          =   288
         Left            =   2100
         TabIndex        =   13
         Top             =   4800
         Visible         =   0   'False
         Width           =   1572
      End
      Begin FACTFRMLib.FactorFrame efraTabFrame1 
         Height          =   4536
         Left            =   108
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   84
         Width           =   8652
         _Version        =   65536
         _ExtentX        =   15261
         _ExtentY        =   8001
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
         Begin VB.ListBox lstOutput 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   10.2
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3924
            HelpContextID   =   3001
            Left            =   48
            TabIndex        =   7
            TabStop         =   0   'False
            Top             =   72
            Width           =   8556
         End
         Begin FACTFRMLib.FactorFrame efraProgressBar 
            Height          =   408
            Left            =   48
            TabIndex        =   9
            Top             =   4068
            Visible         =   0   'False
            Width           =   8568
            _Version        =   65536
            _ExtentX        =   15113
            _ExtentY        =   720
            _StockProps     =   77
            ForeColor       =   -2147483634
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
            CaptionPos      =   4
            FloodColor      =   8388608
            FloodDirection  =   1
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
         Begin FACTFRMLib.FactorFrame efraProgressBar1 
            Height          =   408
            Left            =   48
            TabIndex        =   10
            Top             =   4068
            Visible         =   0   'False
            Width           =   8568
            _Version        =   65536
            _ExtentX        =   15113
            _ExtentY        =   720
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
            FloodColor      =   12582912
            FloodDirection  =   1
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
         Begin FACTFRMLib.FactorFrame efraSearch 
            Height          =   468
            Left            =   48
            TabIndex        =   11
            Top             =   4032
            Visible         =   0   'False
            Width           =   8568
            _Version        =   65536
            _ExtentX        =   15113
            _ExtentY        =   826
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
            FloodColor      =   8388608
            FloodDirection  =   1
            BeginProperty PanelFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Begin VB.TextBox txtSearch 
               Height          =   372
               HelpContextID   =   3002
               Left            =   1344
               TabIndex        =   8
               Top             =   36
               Width           =   6588
            End
            Begin FACTFRMLib.FactorFrame cmdSearch 
               Height          =   372
               HelpContextID   =   3003
               Left            =   7968
               TabIndex        =   4
               TabStop         =   0   'False
               Tag             =   "Copy From"
               Top             =   36
               Width           =   372
               _Version        =   65536
               _ExtentX        =   656
               _ExtentY        =   656
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
                  Size            =   9.6
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin VB.Label lblSearch 
               Alignment       =   1  'Right Justify
               Caption         =   "Search for"
               Height          =   300
               Left            =   60
               TabIndex        =   12
               Top             =   96
               Width           =   1152
            End
         End
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuViewLog 
         Caption         =   "&View Log"
         Visible         =   0   'False
         Begin VB.Menu mnuLastLog 
            Caption         =   "&Last File"
         End
         Begin VB.Menu mnuSelectLog 
            Caption         =   "&Select"
         End
      End
      Begin VB.Menu mnuFileSep 
         Caption         =   "-"
         Visible         =   0   'False
      End
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
      Visible         =   0   'False
      Begin VB.Menu mnuCancel 
         Caption         =   "Ca&ncel"
         Enabled         =   0   'False
         HelpContextID   =   1
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Options"
      Begin VB.Menu mnuProcess 
         Caption         =   "&Process"
      End
      Begin VB.Menu optSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuModules 
         Caption         =   ""
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
Attribute VB_Name = "LogForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'***********************************************************'
'
' Copyright (c) 1996 FACTOR, A Division of W.R.Hess Company
'
' Module name   : TABTEMP.FRM
'
Option Explicit

'=================
'Tabbed Dialog Constants
'=================
Private Const t_szCAPTION_CANCEL As String = "&Cancel"
Private Const t_szCAPTION_EXIT As String = "E&xit"
Private Const t_szEXIT_MESSAGE = "All changes will be lost! Do you want to exit anyway ?"
Private Const t_szCANCEL_MESSAGE = "All changes will be lost! Do you want to cancel anyway ?"
Private Const t_szREFRESH_MESSAGE = "All changes will be lost! Do you want to refresh anyway ?"

Private Const szEMPTY = ""
'==========================
'Status Bar Default Strings
'==========================

Private Const t_szCLEAR As String = ""
Private Const t_szEXIT As String = "Exit"
Private Const t_szCANCEL As String = "Clear the screen"
Private Const t_szADDEDIT As String = "Add or Edit"
Private Const EXTNSN_DATA_FILE = "TXT"
Private Const LOG_FILE_EXTN = "Log"

Private Const QUERY_FILE_DATA = 1
Private Const QUERY_FILE_LOG = 2

'keyBoard constants for the CAPS, NUM and SCROLL LOCK keys
Private Const VK_NUMLOCK As Integer = &H90
Private Const VK_SCROLL As Integer = &H91
Private Const VK_CAPITAL As Integer = &H14

Private Const DATA_INITIAL = 0
Private Const DATA_LOADING = 1
Private Const DATA_LOADING_ERROR = 2
Private Const DATA_LOADED = 3
Private Const DATA_CHANGED = 4

Private Const EDI_SETUP_UP = 9050
Private Const OPT_YES = 0
Private Const OPT_NO = 1
Private Const OPT_YNDFT = 2
Private Const OPT_MANUAL = 0
Private Const OPT_AUTO = 1
Private Const OPT_MADFT = 2
Private Const OPT_CREATE = 0
Private Const OPT_APPEND = 1
Private Const OPT_CADFT = 2

Private Const csPathPattern = "P"
Private Const csFilePattern = "P"

Private Const PROM_BACKUP = "Select 'Yes' to create a backup file"
Private Const PROM_REMOVE = "Select 'Yes' to remove the original data file after processing"
Private Const PROM_WRITEMODE = "Create a new file or append to existing one for output"
Private Const PROM_RUNMODE = "Select 'Auto' to start the program in Auto mode"

Private sInitAppPath As String
Private sInputPath As String
Private sIFileType As String
Private nDataStatus As Integer

Private cValidateIO As cValidateInput

Private Declare Function GetKeyState Lib "user32" ( _
        ByVal nVirtKey As Long) As Integer

Public Sub EnableButtons(ByVal bFlag As Boolean)

    subEnableProcess bFlag
    cmdExitCancelBtn.Enabled = bFlag
    frmContext.ButtonEnabled(EXIT_UP) = bFlag
    frmContext.ButtonEnabled(CANCEL_UP) = bFlag
    frmContext.ButtonEnabled(HELP_UP) = bFlag
    mnuFile.Enabled = bFlag
    mnuOptions.Enabled = bFlag
    mnuHelp.Enabled = bFlag
    
End Sub

Private Function fnDataChanged() As Boolean

    fnDataChanged = (nDataStatus = DATA_CHANGED)
    
End Function

Private Function fnValidWorkPath(txtBox As Textbox) As Boolean

    Dim sPath As String
    Dim sTemp As String
    Dim nTemp As Integer
    Dim bCreated As Boolean
    
    sPath = Trim(txtBox.Text)
    bCreated = False
    If sPath = "" Then
        fnValidWorkPath = True
        Exit Function
    Else
        If fnIsFilePath(sPath) Then
            fnValidWorkPath = True
        Else
            If vbYes = fnConfirmed("The work path entered does not exist. Do you want to create?") Then
                fnValidWorkPath = fnCreatePath(sPath)
                bCreated = True
            Else
                fnValidWorkPath = False
            End If
        End If
    End If
    If fnValidWorkPath Then
        If UCase(txtBox.Tag) <> UCase(sPath) Then
            subChangePath udtLogInfo.m_sPath, txtBox.Tag, sPath
            If UCase(txtLogPath.Text) <> UCase(udtLogInfo.m_sPath) Then
                txtLogPath.Text = udtLogInfo.m_sPath
            End If
        End If
        subCheckWorkPath bCreated
    End If
End Function


Public Sub ShowLog(sLog As String)
    Dim aryLogs() As String
    Dim i As Integer
    
    If Trim(sLog) = "" Then
        lstOutput.AddItem sLog
    Else
        subParseString aryLogs, sLog, vbCrLf
        For i = 0 To UBound(aryLogs)
            lstOutput.AddItem aryLogs(i)
        Next i
        lstOutput.ListIndex = lstOutput.ListCount - 1
    End If
End Sub

Private Sub subCheckWorkPath(ByVal bCreated As Boolean)
    
    Dim nConfirmed As Integer
    Dim sTemp As String
    Dim sPath As String
    
    sTemp = UCase(txtWorkPath.Tag)
    sPath = UCase(Trim(txtWorkPath.Text))
    If sPath <> sTemp Or bCreated Then
        subAddSlash sPath
        nConfirmed = vbDefault
        subFixPath udtLogInfo.m_sPath, nConfirmed, sTemp, sPath
        subFixPath udtInputInfo.m_sPath, nConfirmed, sTemp, sPath
        subFixPath udtBackupInfo.m_sPath, nConfirmed, sTemp, sPath
    End If

End Sub

Private Sub subClearIO()
    
    optRunModeIO(OPT_AUTO).value = False
    optRunModeIO(OPT_MANUAL).value = False
    txtFilePath.Text = ""
    txtFileName.Text = ""
    optBKFInput(OPT_YES).value = False
    optBKFInput(OPT_NO).value = False
    optRMFInput(OPT_YES).value = False
    optRMFInput(OPT_NO).value = False
    txtBackupPath.Text = ""
    txtBackupName.Text = ""
    
    txtWorkPath.Text = ""
    txtWorkPath.Tag = ""
    optWriteLog(OPT_YES).value = False
    optWriteLog(OPT_NO).value = False
    txtLogFile.Text = ""
    txtLogPath.Text = ""
    cValidateIO.ResetFlags
    nDataStatus = DATA_INITIAL

End Sub

Private Sub subEnableIO(ByVal bFlag As Boolean)
    
    txtFilePath.Enabled = bFlag
    txtFileName.Enabled = bFlag
    cmdInputPath.Enabled = bFlag
    optBKFInput(OPT_YES).Enabled = bFlag
    optBKFInput(OPT_NO).Enabled = bFlag
    optRMFInput(OPT_YES).Enabled = bFlag
    optRMFInput(OPT_NO).Enabled = bFlag
    subEnableIOBackup bFlag
    optRunModeIO(OPT_AUTO).Enabled = bFlag
    optRunModeIO(OPT_MANUAL).Enabled = bFlag

    txtWorkPath.Enabled = bFlag
    cmdWorkPath.Enabled = bFlag
    optWriteLog(OPT_YES).Enabled = bFlag
    optWriteLog(OPT_NO).Enabled = bFlag
    subEnableLog bFlag

End Sub


Private Sub subDataChanged()
    If nDataStatus = DATA_LOADED Then
        nDataStatus = DATA_CHANGED
    End If
End Sub

Private Sub subEnableIOBackup(ByVal bFlag As Boolean)
    txtBackupPath.Enabled = bFlag
    txtBackupName.Enabled = bFlag
    cmdBackupPath.Enabled = bFlag
End Sub

Private Sub subEnableLog(ByVal bFlag As Boolean)
    txtLogFile.Enabled = bFlag
    txtLogPath.Enabled = bFlag
    cmdLogPath.Enabled = bFlag
End Sub


Private Sub subEnablePrint(bFlag As Boolean)
    frmContext.ButtonEnabled(PRINT_UP) = bFlag
    mnuPrint.Enabled = bFlag
End Sub

Public Sub ShowProgress(fPercent As Single)
    Dim nPercent As Integer
    
    nPercent = fPercent * 100
    efraProgressBar.FloodPercent = nPercent
    If nPercent < 50 Then
        efraProgressBar.ForeColor = &H80000012
    Else
        efraProgressBar.ForeColor = &H8000000E
    End If
    efraProgressBar.Caption = Format(nPercent, "#0") & "%"
    efraProgressBar.Refresh
    
End Sub

Public Sub ShowProgressbar(bFlag As Boolean)
    efraProgressBar.Visible = bFlag
    If lstOutput.ListCount > 0 Then
        efraSearch.Visible = Not bFlag
    Else
        efraSearch.Visible = False
    End If
End Sub


Private Sub subEnableSUOK(ByVal bFlag As Boolean)
    cmdOk.Enabled = bFlag
End Sub

Private Sub subFixPath(sIPath As String, _
                       nConfirmed As Integer, _
                       sOldPath As String, _
                       sNewPath As String)

    Dim nTemp As Integer
    Dim bFlag As Boolean
    Dim sPath As String
    
    sPath = sIPath
    nTemp = Len(sOldPath)
    If UCase(Left(sPath, nTemp)) <> sOldPath Then
        If Not fnIsFilePath(sPath) Then
            subProcessPath sPath
            If vbYes = fnConfirmed("Path: '" & sPath & "' does not exist. Do you want to create it?.") Then
                fnCreatePath sPath
            End If
        End If
    Else
        sPath = sNewPath & fnGetSubPath(sOldPath, sPath)
        If sPath <> "" Then
            If Not fnIsFilePath(sPath) Then
                If nConfirmed = vbDefault Then
                    nConfirmed = fnConfirmed("Do you also want to create the associate subdirectories?")
                ElseIf nConfirmed = vbNo Or nConfirmed = vbYes Then
                    nConfirmed = fnConfirmed("Do you want to create path " & sPath & "?")
                End If
            End If
        End If
    End If

End Sub

Private Sub subChangePath(sPath As String, _
                          sOldPath As String, _
                          sNewPath As String)

    Dim nTemp As Integer
    Dim sTemp As String
    
    nTemp = Len(sOldPath)
    If nTemp > 0 Then
        If UCase(Left(sPath, nTemp)) = sOldPath Then
            sTemp = fnGetSubPath(sOldPath, sPath)
            If sTemp <> "" Then
                subAddSlash sNewPath
                sPath = sNewPath & sTemp
            End If
        End If
        If Not fnIsFilePath(sPath) Then
            If Not fnIsWholePath(sPath) Then
                subProcessPath sPath
            End If
        End If
    End If
End Sub


Private Sub subLoadInOut()
    Dim nConfirmed As Integer

    If fnTestFlag(m_nRunParm, RP_MASK_AUTO) Then
        optRunModeIO(OPT_AUTO).value = True
    Else
        optRunModeIO(OPT_MANUAL).value = True
    End If
    txtWorkPath.Text = m_sWorkPath
    txtWorkPath.Tag = m_sWorkPath
    If fnTestFlag(m_nRunParm, RP_MASK_WTLOG) Then
        optWriteLog(OPT_YES).value = True
    Else
        optWriteLog(OPT_NO).value = True
    End If
    txtLogPath.Text = udtLogInfo.m_sPath
    If Not cValidateIO.ValidInput(txtLogPath) Then
    End If
    txtLogFile.Text = fnMakeFName(udtLogInfo.m_sFile, udtLogInfo.m_sType)
    
    txtFilePath.Text = udtInputInfo.m_sPath
    If Not cValidateIO.ValidInput(txtFilePath) Then
    End If
    txtFileName.Text = fnMakeFName(udtInputInfo.m_sFile, udtInputInfo.m_sType)
    If fnTestFlag(m_nRunParm, RP_MASK_BKFILE) Then
        optBKFInput(OPT_YES).value = True
    Else
        optBKFInput(OPT_NO).value = True
    End If
    If fnTestFlag(m_nRunParm, RP_MASK_RMFILE) Then
        optRMFInput(OPT_YES).value = True
    Else
        optRMFInput(OPT_NO).value = True
    End If
    txtBackupPath.Text = udtBackupInfo.m_sPath
    If Not cValidateIO.ValidInput(txtBackupPath) Then
    End If
    txtBackupName.Text = fnMakeFName(udtBackupInfo.m_sFile, udtBackupInfo.m_sType)
End Sub

Private Sub subResetSetup()

    nDataStatus = DATA_INITIAL
    subLoadInOut
    subEnableSUCancel False
    subEnableSUOK False
    subSetOptFocus optRunModeIO(OPT_YES), optRunModeIO(OPT_NO)
    nDataStatus = DATA_LOADED
    
End Sub

Private Sub subSetButtonStatus()
    Dim bFlag As Boolean
    
    If fnDataChanged Then
        bFlag = (cValidateIO.FirstInvalidInput < 0)
    Else
        bFlag = False
    End If
    If bFlag Then
        subEnableSUOK True
    End If
    
End Sub

Public Sub subSetFileInfo()
    sInputPath = udtInputInfo.m_sPath
    subProcessPath sInputPath
    sIFileType = udtInputInfo.m_sType
End Sub

Private Sub subSetToNextBox(txtBox As Control)

    Select Case txtBox.TabIndex
        Case optRunModeIO(OPT_YES).TabIndex, optRunModeIO(OPT_NO).TabIndex
            subSetFocus txtWorkPath
        Case txtWorkPath.TabIndex
            subSetOptFocus optWriteLog(OPT_YES), optWriteLog(OPT_NO)
        Case optWriteLog(OPT_YES).TabIndex, optWriteLog(OPT_NO).TabIndex
            If txtLogPath.Enabled Then
                subSetFocus txtLogPath
            Else
                subSetFocus txtFilePath
            End If
        Case txtLogPath.TabIndex
            subSetFocus txtLogFile
        Case txtLogFile.TabIndex
            subSetFocus txtFilePath
        Case txtFilePath.TabIndex
            subSetFocus txtFileName
        Case txtFileName.TabIndex
            subSetOptFocus optBKFInput(OPT_YES), optBKFInput(OPT_NO)
        Case optBKFInput(OPT_YES).TabIndex, optBKFInput(OPT_NO).TabIndex
            subSetOptFocus optRMFInput(OPT_YES), optRMFInput(OPT_NO)
        Case optRMFInput(OPT_YES).TabIndex, optRMFInput(OPT_NO).TabIndex
            If txtBackupPath.Enabled Then
                subSetFocus txtBackupPath
            Else
                subSetFocus cmdOk, cmdClose
            End If
        Case txtBackupPath.TabIndex
            subSetFocus txtBackupName
        Case txtBackupName.TabIndex
            subSetFocus cmdOk, cmdClose
    End Select
End Sub

Private Sub subSetupValidation()
    Set cValidateIO = New cValidateInput
    With cValidateIO
        Set .Form = Me
        Set .StatusBar = ffraStatusbar
        .AddEditBox txtWorkPath, "Enter the working path"
        .AddEditBox txtLogPath, "Enter the path where the log file will be written to"
        .AddEditBox txtLogFile, "Enter the log file name"
        .AddEditBox txtFilePath, "Enter the input data file path"
        .AddEditBox txtFileName, "Enter the input data file name"
        .AddEditBox txtBackupPath, "Enter the backup path for input data"
        .AddEditBox txtBackupName, "Enter the backup name for input data"
    End With
End Sub

Private Sub subShowSetup(ByVal bShow As Boolean)

    ffBackground2.Visible = bShow
    ffBackground1.Visible = Not bShow
    subEnableProcess Not bShow
    frmContext.ButtonEnabled(EDI_SETUP_UP) = Not bShow
    If bShow Then
        ffBackground2.ZOrder 0
    Else
        ffBackground2.ZOrder 1
    End If
    
End Sub

Private Sub subStoreIOInfo()
    
    m_sWorkPath = txtWorkPath.Text
    udtLogInfo.m_sPath = txtLogPath

    udtLogInfo.m_sFile = fnStripFileName(txtLogFile.Text)
    udtLogInfo.m_sType = fnFileType(txtLogFile.Text)
    If optWriteLog(OPT_YES).value Then
        subSetFlag m_nRunParm, RP_MASK_WTLOG, True
    ElseIf optWriteLog(OPT_NO).value Then
        subSetFlag m_nRunParm, RP_MASK_WTLOG, False
    End If
    If optRunModeIO(OPT_AUTO).value Then
        subSetFlag m_nRunParm, RP_MASK_AUTO, True
    ElseIf optRunModeIO(OPT_MANUAL).value Then
        subSetFlag m_nRunParm, RP_MASK_AUTO, False
    End If
    udtInputInfo.m_sPath = txtFilePath.Text
    udtInputInfo.m_sFile = fnStripFileName(txtFileName.Text)
    udtInputInfo.m_sType = fnFileType(txtFileName.Text)
    If optBKFInput(OPT_YES).value Then
        subSetFlag m_nRunParm, RP_MASK_BKFILE, True
    ElseIf optBKFInput(OPT_NO).value Then
        subSetFlag m_nRunParm, RP_MASK_BKFILE, False
    End If
    If optRMFInput(OPT_YES).value Then
        subSetFlag m_nRunParm, RP_MASK_RMFILE, True
    ElseIf optRMFInput(OPT_NO).value Then
        subSetFlag m_nRunParm, RP_MASK_RMFILE, False
    End If
    udtBackupInfo.m_sPath = txtBackupPath.Text
    udtBackupInfo.m_sFile = fnStripFileName(txtBackupName.Text)
    udtBackupInfo.m_sType = fnFileType(txtBackupName.Text)

End Sub


Public Sub TBButtonCallBack(ByVal nID As Integer)
    Select Case nID
        Case PRINT_UP
            subPrintClick
        Case COPY_UP
        Case EDI_SETUP_UP
            subShowSetup True
            subResetSetup
        Case CANCEL_UP
            subCancel
        Case EXIT_UP
            subExit
    End Select
End Sub

Private Sub subSetupToolbar()
    With frmContext
        .BeginSetupToolbar Me
        .AddButton "Setup INI File", EDI_SETUP_UP, , True
        .EndSetupToolbar
    End With
End Sub

Private Sub cmdClose_Click()

    If cmdClose.Caption = t_szCAPTION_CANCEL Then
        If fnDataChanged Then
            If vbYes <> fnConfirmed(t_szCANCEL_MESSAGE) Then
                Exit Sub
            End If
        End If
        subResetSetup
    Else
        If fnDataChanged Then
            If vbYes <> fnConfirmed(t_szEXIT_MESSAGE) Then
                Exit Sub
            End If
        End If
        'subResetSetup
        subShowSetup False
    End If
End Sub

Private Sub cmdClose_GotFocus()
    tfnSetStatusBarMessage "Go back to main screen"
End Sub

Private Sub cmdBackupPath_Click()
    cmdBackupPath.Enabled = False
    cmdBackupPath.Enabled = True

    subBrowsePath txtBackupPath, "Select Backup File Path for input Files"
    cValidateIO.LostFocus txtBackupPath
    If Not cValidateIO.ValidInput(txtBackupPath) Then
        subSetFocus txtBackupPath
    End If
End Sub

Private Sub cmdInputPath_Click()
    cmdInputPath.Enabled = False
    cmdInputPath.Enabled = True
    subBrowsePath txtFilePath, "Select Path for input data files"

    cValidateIO.LostFocus txtFilePath
    If Not cValidateIO.ValidInput(txtFilePath) Then
        subSetFocus txtFilePath
    End If
End Sub

Private Sub cmdLogPath_Click()
    cmdLogPath.Enabled = False
    cmdLogPath.Enabled = True
    If txtLogPath.Text = "" Then
        subBrowsePath txtWorkPath, "Select Log File Path"
    Else
        subBrowsePath txtLogPath, "Select Log File Path"
    End If
    cValidateIO.LostFocus txtLogPath
    If Not cValidateIO.ValidInput(txtLogPath) Then
        subSetFocus txtLogPath
    End If
End Sub

Private Sub cmdOK_Click()
    'Click
    subStoreIOInfo
    subWriteInOut
    sInputPath = udtInputInfo.m_sPath
    sIFileType = udtInputInfo.m_sType
    subShowSetup False
    
End Sub

Private Function fnValidPath(txtBox As Textbox) As Boolean

    Dim sPath As String
    Dim sTemp As String
    Dim bFlag As Boolean

    sPath = Trim(txtBox.Text)
    If sPath = "" Then
        fnValidPath = True
    Else
        If fnIsFilePath(sPath) Then
            fnValidPath = True
        Else
            bFlag = False
            If Not fnIsWholePath(sPath) Then
                sTemp = Trim(txtWorkPath.Text)
                subAddSlash sTemp
                sPath = sTemp & sPath
                bFlag = True
            End If
            If fnIsFilePath(sPath) Then
                fnValidPath = True
            Else
                If vbYes = fnConfirmed("The path( " & sPath & ") entered does not exist. Do you want to create?") Then
                    fnValidPath = fnCreatePath(sPath)
                Else
                    fnValidPath = False
                End If
            End If
            If bFlag And fnValidPath Then
                txtBox.Text = sPath
            End If
        End If
    End If
End Function

Private Sub cmdOK_GotFocus()
    tfnSetStatusBarMessage "Store the changes"
End Sub


Private Sub cmdProcess_Click()
    subEnableProcess False
    subEnableCancel True
    frmContext.ButtonEnabled(EDI_SETUP_UP) = False
    
    Dim sFile As String
    
    sFile = fnFileName(QUERY_FILE_DATA)

    If fnIsFile(sFile) Then
        subPrepareLog
        'Do process
        bBatchMode = False
        subProcessFile sFile
    Else
        subEnableCancel False
        subEnableProcess True
        frmContext.ButtonEnabled(EDI_SETUP_UP) = True
    End If

End Sub

Private Sub subSetFocus(cntlTemp As Control, ParamArray arryControls() As Variant)
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

Public Sub subEnableProcess(bFlag As Boolean)
    cmdProcess.Enabled = bFlag
    mnuProcess.Enabled = bFlag
End Sub

Private Sub cmdProcess_GotFocus()
    tfnSetStatusBarMessage "Click to process data file"
End Sub

Private Sub cmdWorkPath_Click()
    cmdWorkPath.Enabled = False
    cmdWorkPath.Enabled = True
    subBrowsePath txtWorkPath, "Select the Work Path"
    cValidateIO.LostFocus txtWorkPath
    If Not cValidateIO.ValidInput(txtWorkPath) Then
        subSetFocus txtWorkPath
    End If
End Sub

Private Sub subBrowsePath(txtBox As Textbox, _
                          sTitle As String)
    
    frmDlgPath.Caption = sTitle
    On Error Resume Next
    If Mid(txtBox.Text, 2, 1) = ":" Then
        frmDlgPath.drvDrive.Drive = Left(txtBox.Text, 1)
    End If
    frmDlgPath.dirPath.Path = txtBox.Text
'    frmDlgPath.lblPath.Caption = txtBox.Text
    frmDlgPath.Show vbModal
    If frmDlgPath.Action = vbOK Then
        txtBox.Text = frmDlgPath.lblPath.Caption
        subSetButtonStatus
        subSetToNextBox txtBox
    End If
End Sub

Private Sub Form_Resize()
    frmContext.FormResize
End Sub

'===========
'Form Events
'===========

Private Sub Form_Initialize() 'called before Form_Load
    
    ' ** change the help file for this application
'    App.HelpFile = szHelpElecCommerce
    sInitAppPath = App.Path
    subGoLevelUp sInitAppPath
    subAddSlash sInitAppPath
    
End Sub

Private Sub Form_Load()

    subSetupToolbar
    
    ' disable system menu Close and Close icon on form
'    tfnDisableFormSystemClose Me

    tmrKeyBoard.Enabled = False
    subCenterForm Me
'    Me.Show
    Screen.MousePointer = vbHourglass
    tfnSetInitializingMessage
    Me.Enabled = False
    DoEvents

    '***************************************************
    ' INSERT YOUR FORM LOAD CODE HERE
    ' | | | | | | |
    ' v v v v v v v
    subSetupValidation
    
    ' ^ ^ ^ ^ ^ ^ ^
    ' | | | | | | |
    '***************************************************
    Screen.MousePointer = vbDefault
    Me.Enabled = True
    
    '***************************************************
    ' SET YOUR FIRST STATUSBAR MESSAGE HERE
    ' | | | | | | |
    ' v v v v v v v
    
    
    ' ^ ^ ^ ^ ^ ^ ^
    ' | | | | | | |
    '***************************************************
    tfnResetScreen 'set the default screen
    
    tmrKeyBoard.Enabled = True

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
    
    If cmdExitCancelBtn.Caption = t_szCAPTION_EXIT Then
        subExit
    Else
        subCancel
    End If

End Sub


Private Sub subPrintClick()
    
    Screen.MousePointer = vbHourglass
    
    subPrint lstOutput
    subSetFocus cmdExitCancelBtn
    Screen.MousePointer = vbDefault
End Sub

Private Sub subCancel()

    tfnResetScreen
    
End Sub

Private Sub subExit()

    Unload Me
    End
    
End Sub

Private Sub mnuExit_Click()
    subExit
End Sub

Private Sub mnuModules_Click(Index As Integer)
    frmContext.MenuClick Index
End Sub

Private Sub mnuPrint_Click()
    subPrintClick
End Sub

Private Sub mnuContents_Click()
    frmContext.RunItem HELP_UP
End Sub

Private Sub mnuAbout_Click()
    subCenterForm frmAbout, Me
    frmAbout.Show vbModal
End Sub


Private Sub optBKFInput_Click(Index As Integer)
    subDataChanged
    subSetButtonStatus
    If Index = OPT_YES Then
        subEnableIOBackup True
    Else
        subEnableIOBackup False
    End If
End Sub


Private Sub optBKFInput_GotFocus(Index As Integer)
    tfnSetStatusBarMessage PROM_BACKUP
End Sub


Private Sub optBKFInput_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = 0
        subSetToNextBox optBKFInput(Index)
    End If
End Sub

Private Sub optBKFInput_LostFocus(Index As Integer)
    subSetButtonStatus
End Sub

Private Sub optRMFInput_Click(Index As Integer)
    subDataChanged
    subSetButtonStatus
End Sub


Private Sub optRMFInput_GotFocus(Index As Integer)
    tfnSetStatusBarMessage PROM_REMOVE
End Sub


Private Sub optRMFInput_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = 0
        subSetToNextBox optRMFInput(Index)
    End If
End Sub


Private Sub optRMFInput_LostFocus(Index As Integer)
    subSetButtonStatus
End Sub

Private Sub optRunModeIO_Click(Index As Integer)
    subDataChanged
    subSetButtonStatus
End Sub

Private Sub optRunModeIO_GotFocus(Index As Integer)
    tfnSetStatusBarMessage PROM_RUNMODE
End Sub


Private Sub optRunModeIO_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = 0
        subSetToNextBox optRunModeIO(Index)
    End If
End Sub


Private Sub optWriteLog_Click(Index As Integer)
    subDataChanged
    subSetButtonStatus
    If Index = OPT_YES Then
        subEnableLog True
    Else
        subEnableLog False
    End If
End Sub

Private Sub optWriteLog_GotFocus(Index As Integer)
    tfnSetStatusBarMessage "Select 'Yes' to create a log file"
End Sub


Private Sub optWriteLog_LostFocus(Index As Integer)
    subSetButtonStatus
End Sub


Private Sub txtWorkPath_Change()
    cValidateIO.Change txtWorkPath
    tfnRegExpControlChange txtWorkPath, csPathPattern
    subDataChanged
End Sub

Private Sub txtFilePath_Change()
    cValidateIO.Change txtFilePath
    tfnRegExpControlChange txtFilePath, csPathPattern
    subDataChanged
End Sub

Private Sub txtFilePath_GotFocus()
    cValidateIO.GotFocus txtFilePath
    subSelectText txtFilePath
    tfnRegExpControlGotFocus txtFilePath, csPathPattern
End Sub

Private Sub txtFilePath_KeyPress(KeyAscii As Integer)

    Dim bCode As Boolean
    
    bCode = tfnRegExpControlKeyPress(txtFilePath, KeyAscii, csPathPattern)

    If bCode Then
        If KeyAscii = vbKeyReturn Then
            subSetToNextBox txtFilePath
            KeyAscii = 0
        Else
            cValidateIO.Keypress txtFilePath, KeyAscii
        End If
    Else
        KeyAscii = 0
    End If

End Sub

Private Sub txtFilePath_LostFocus()
    If cValidateIO.LostFocus(txtFilePath) Then
        subSetButtonStatus
    End If
End Sub

Private Sub txtFileName_Change()
    cValidateIO.Change txtFileName
    tfnRegExpControlChange txtFileName, csFilePattern
    subDataChanged
End Sub

Private Sub txtFileName_GotFocus()
    cValidateIO.GotFocus txtFileName
    subSelectText txtFileName
    tfnRegExpControlGotFocus txtFileName, csFilePattern
End Sub

Private Sub txtFileName_KeyPress(KeyAscii As Integer)

    Dim bCode As Boolean
    
    bCode = tfnRegExpControlKeyPress(txtFileName, KeyAscii, csFilePattern)

    If bCode Then
        If KeyAscii = vbKeyReturn Then
            subSetToNextBox txtFileName
            KeyAscii = 0
        Else
            cValidateIO.Keypress txtFileName, KeyAscii
        End If
    Else
        KeyAscii = 0
    End If

End Sub

Private Sub txtFileName_LostFocus()
    If cValidateIO.LostFocus(txtFileName) Then
        subSetButtonStatus
    End If
End Sub

Private Sub txtBackupPath_Change()
    cValidateIO.Change txtBackupPath
    tfnRegExpControlChange txtBackupPath, csPathPattern
    subDataChanged
End Sub

Private Sub txtBackupPath_GotFocus()
    cValidateIO.GotFocus txtBackupPath
    subSelectText txtBackupPath
    tfnRegExpControlGotFocus txtBackupPath, csPathPattern
End Sub

Private Sub txtBackupPath_KeyPress(KeyAscii As Integer)

    Dim bCode As Boolean
    
    bCode = tfnRegExpControlKeyPress(txtBackupPath, KeyAscii, csPathPattern)

    If bCode Then
        If KeyAscii = vbKeyReturn Then
            subSetToNextBox txtBackupPath
            KeyAscii = 0
        Else
            cValidateIO.Keypress txtBackupPath, KeyAscii
        End If
    Else
        KeyAscii = 0
    End If

End Sub

Private Sub txtBackupPath_LostFocus()
    If cValidateIO.LostFocus(txtBackupPath) Then
        subSetButtonStatus
    End If
End Sub

Private Sub txtBackupName_Change()
    cValidateIO.Change txtBackupName
    tfnRegExpControlChange txtBackupName, csFilePattern
    subDataChanged
End Sub

Private Sub txtBackupName_GotFocus()
    cValidateIO.GotFocus txtBackupName
    subSelectText txtBackupName
    tfnRegExpControlGotFocus txtBackupName, csFilePattern
End Sub

Private Sub txtBackupName_KeyPress(KeyAscii As Integer)

    Dim bCode As Boolean
    
    bCode = tfnRegExpControlKeyPress(txtBackupName, KeyAscii, csFilePattern)

    If bCode Then
        If KeyAscii = vbKeyReturn Then
            subSetToNextBox txtBackupName
            KeyAscii = 0
        Else
            cValidateIO.Keypress txtBackupName, KeyAscii
        End If
    Else
        KeyAscii = 0
    End If

End Sub

Private Sub txtBackupName_LostFocus()
    If cValidateIO.LostFocus(txtBackupName) Then
        subSetButtonStatus
    End If
End Sub

Private Sub txtWorkPath_GotFocus()
    cValidateIO.GotFocus txtWorkPath
    subSelectText txtWorkPath
    tfnRegExpControlGotFocus txtWorkPath, csPathPattern
End Sub

Private Sub subSelectText(txtBox As Textbox)
    With txtBox
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtWorkPath_KeyPress(KeyAscii As Integer)

    Dim bCode As Boolean
    
    bCode = tfnRegExpControlKeyPress(txtWorkPath, KeyAscii, csPathPattern)

    If bCode Then
        If KeyAscii = vbKeyReturn Then
            subSetToNextBox txtWorkPath
            KeyAscii = 0
        Else
            cValidateIO.Keypress txtWorkPath, KeyAscii
        End If
    Else
        KeyAscii = 0
    End If

End Sub

Private Sub txtLogPath_Change()
    cValidateIO.Change txtLogPath
    tfnRegExpControlChange txtLogPath, csPathPattern
    subDataChanged
End Sub

Private Sub txtLogPath_GotFocus()
    cValidateIO.GotFocus txtLogPath
    subSelectText txtLogPath
    tfnRegExpControlGotFocus txtLogPath, csPathPattern
End Sub

Private Sub txtLogPath_KeyPress(KeyAscii As Integer)

    Dim bCode As Boolean
    
    bCode = tfnRegExpControlKeyPress(txtLogPath, KeyAscii, csPathPattern)

    If bCode Then
        If KeyAscii = vbKeyReturn Then
            subSetToNextBox txtLogPath
            KeyAscii = 0
        Else
            cValidateIO.Keypress txtLogPath, KeyAscii
        End If
    Else
        KeyAscii = 0
    End If

End Sub

Private Sub txtLogPath_LostFocus()
    If cValidateIO.LostFocus(txtLogPath) Then
        subSetButtonStatus
    End If
End Sub

Private Sub txtLogFile_Change()
    cValidateIO.Change txtLogFile
    tfnRegExpControlChange txtLogFile, csFilePattern
    subDataChanged
End Sub

Private Sub txtLogFile_GotFocus()
    cValidateIO.GotFocus txtLogFile
    subSelectText txtLogFile
    tfnRegExpControlGotFocus txtLogFile, csFilePattern
End Sub

Private Sub txtLogFile_KeyPress(KeyAscii As Integer)

    Dim bCode As Boolean
    
    bCode = tfnRegExpControlKeyPress(txtLogFile, KeyAscii, csFilePattern)

    If bCode Then
        If KeyAscii = vbKeyReturn Then
            subSetToNextBox txtLogFile
            KeyAscii = 0
        Else
            cValidateIO.Keypress txtLogFile, KeyAscii
        End If
    Else
        KeyAscii = 0
    End If

End Sub

Private Sub txtLogFile_LostFocus()
    If cValidateIO.LostFocus(txtLogFile) Then
        subSetButtonStatus
    End If
End Sub

Private Sub txtWorkPath_LostFocus()
    If cValidateIO.LostFocus(txtWorkPath) Then
        subSetButtonStatus
    End If
End Sub

Private Sub mnuCancel_Click()
    subCancel
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
Private Sub tfnResetScreen()
    
    On Error Resume Next
    
    subShowSetup False
    DoEvents
    lstOutput.SetFocus
    
    subEnablePrint False
    subEnableProcess True
    
    lstOutput.Clear
    txtSearch.Text = ""
    
    subEnableCancel False
    
    DoEvents
    subSetFocus cmdProcess
End Sub

Public Function fnInvalidData(txtBox As Textbox) As Boolean

    Select Case txtBox.TabIndex
        Case txtWorkPath.TabIndex
            fnInvalidData = Not fnValidWorkPath(txtBox)
        Case txtLogPath.TabIndex
            fnInvalidData = Not fnValidPath(txtBox)
        Case txtFilePath.TabIndex
            fnInvalidData = Not fnValidPath(txtBox)
        Case txtBackupPath.TabIndex
            fnInvalidData = Not fnValidPath(txtBox)
    End Select
End Function

Private Sub subSetOptFocus(ParamArray optControls())
    'Set the focus to a set of option buttons, such that the focus
    ' goes to the one with value property being true (selected)

    Dim i As Integer
    
    For i = 0 To UBound(optControls)
        If optControls(i).value Then
            Exit For
        End If
    Next
    If i > UBound(optControls) Then
        i = 0
    End If
    
    On Error Resume Next
    optControls(i).SetFocus
    On Error GoTo 0
End Sub

Public Sub subEnableCancel(bStatus As Boolean)
    mnuCancel.Enabled = bStatus
    frmContext.ButtonEnabled(CANCEL_UP) = bStatus
    If bStatus Then
        cmdExitCancelBtn.Caption = t_szCAPTION_CANCEL
    Else
        cmdExitCancelBtn.Caption = t_szCAPTION_EXIT
    End If
End Sub

Private Sub subEnableSUCancel(bStatus As Boolean)
    mnuCancel.Enabled = bStatus
    frmContext.ButtonEnabled(CANCEL_UP) = bStatus
    If bStatus Then
        cmdClose.Caption = t_szCAPTION_CANCEL
    Else
        cmdClose.Caption = "&Close"
    End If
End Sub

Private Sub mnuProcess_Click()
    cmdProcess_Click
End Sub

Private Function fnFileName(nType As Integer) As String
    Dim sExtn As String
    Dim sFilter As String
    Dim sPath As String
    Dim sName As String
    
    With dlgFilenames
        Select Case nType
            Case QUERY_FILE_DATA
                sExtn = "*." & sIFileType    'fnDefaultFiles(sSectionParms, "*." & EXTNSN_DATA_FILE)
                .InitDir = sInputPath
                sFilter = "Text (*.txt)|*.txt|All Files (*.*)|*.*"
                If UCase(udtInputInfo.m_sType) = "*.TXT" Or udtInputInfo.m_sType = "*.*" Then
                    .Filter = sFilter
                    .FileName = "*.txt"
                Else
                    .Filter = "Flat Text File (*." & udtInputInfo.m_sType & ")|*." & udtInputInfo.m_sType & "|" & sFilter
                    .FileName = sExtn
                End If
                .DialogTitle = "Open EFT Transaction File For Processing"
            Case QUERY_FILE_LOG
                .InitDir = "*.LOG"      'fnDefaultParm(sSectionParms, KEY_LOG_FILE_PATH, "")
                .DialogTitle = "Open Log File For Viewing"
                .FileName = "P*." & LOG_FILE_EXTN
                .Filter = "SYDTNTPP Log File|P*." & LOG_FILE_EXTN & "|All Files (*.*)|*.*"
        End Select
        .CancelError = True
        On Error Resume Next
        .ShowOpen
        If Err.Number = 32755 Then
            fnFileName = ""
            'Canceled
        Else
            fnFileName = .FileName
            subParseFile sPath, sName, .FileName
            sIFileType = fnFileType(sName)
            sInputPath = sPath
        End If
    End With
End Function

Private Sub optWriteLog_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = 0
        subSetToNextBox optWriteLog(Index)
    End If
End Sub


Private Sub tbToolbar_ButtonClick(ByVal Button As Button)
    frmContext.ButtonClick Button
End Sub

Private Sub tbToolbar_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    frmContext.TBMouseMove
End Sub


'=========================
'tooltip support functions
'=========================

Private Sub tmrKeyboard_Timer() 'status bar timer - 250ms

    tfnUpdateStatusBar Me 'process the status bar

End Sub

Public Sub tfnUpdateStatusBar(frmForm As Form)
    Dim intKeyStatus As Integer

    On Error Resume Next

    intKeyStatus = GetKeyState(VK_CAPITAL)

    If intKeyStatus = 1 Then
        frmForm.ffraStatusbar.PanelCaption(2) = "CAPS"
    Else
        frmForm.ffraStatusbar.PanelCaption(2) = szEMPTY
        End If

    DoEvents

    intKeyStatus = GetKeyState(VK_SCROLL)

    If intKeyStatus = 1 Then
        frmForm.ffraStatusbar.PanelCaption(0) = "SCRL"
    Else
        frmForm.ffraStatusbar.PanelCaption(0) = szEMPTY
    End If

    DoEvents

    intKeyStatus = GetKeyState(VK_NUMLOCK)

    If intKeyStatus = 1 Then
        frmForm.ffraStatusbar.PanelCaption(1) = "NUM"
    Else
        frmForm.ffraStatusbar.PanelCaption(1) = szEMPTY
    End If
End Sub


