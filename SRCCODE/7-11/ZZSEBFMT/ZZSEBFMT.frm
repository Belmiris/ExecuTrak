VERSION 5.00
Object = "{C75015E0-2232-11D3-B440-0060971E99AF}#1.0#0"; "FACTFRM.OCX"
Object = "{01028C21-0000-0000-0000-000000000046}#4.0#0"; "TG32OV.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{478E45E0-5745-11CF-8918-00A02416C765}#1.0#0"; "SQAOTE32.OCX"
Begin VB.Form frmZZSEBFMT 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Commission Formula Maintenance"
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
   Begin FACTFRMLib.FactorFrame efraBackground 
      Height          =   5184
      Left            =   0
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   492
      Width           =   8880
      _Version        =   65536
      _ExtentX        =   15663
      _ExtentY        =   9144
      _StockProps     =   77
      BackColor       =   8388608
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
      Begin FACTFRMLib.FactorFrame efraBase 
         Height          =   4632
         Left            =   60
         TabIndex        =   29
         Top             =   60
         Width           =   8760
         _Version        =   65536
         _ExtentX        =   15452
         _ExtentY        =   8170
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
         Begin VB.TextBox txtLevel 
            Height          =   360
            HelpContextID   =   304
            Left            =   7488
            TabIndex        =   7
            Top             =   288
            Width           =   792
         End
         Begin VB.TextBox txtMaxTotal 
            Height          =   360
            HelpContextID   =   310
            Left            =   7032
            TabIndex        =   13
            Top             =   948
            Width           =   1620
         End
         Begin VB.TextBox txtAmount2 
            Height          =   360
            HelpContextID   =   309
            Left            =   5304
            TabIndex        =   12
            Top             =   948
            Width           =   1596
         End
         Begin VB.TextBox txtVariable2 
            Height          =   360
            HelpContextID   =   313
            Left            =   2604
            TabIndex        =   16
            Top             =   1608
            Width           =   2064
         End
         Begin VB.TextBox txtVariable3 
            Height          =   360
            HelpContextID   =   315
            Left            =   5100
            TabIndex        =   18
            Top             =   1608
            Width           =   2064
         End
         Begin VB.TextBox txtPercent 
            Height          =   360
            HelpContextID   =   306
            Left            =   108
            TabIndex        =   9
            Top             =   948
            Width           =   1596
         End
         Begin VB.TextBox txtAmount1 
            Height          =   360
            HelpContextID   =   308
            Left            =   3576
            TabIndex        =   11
            Top             =   948
            Width           =   1600
         End
         Begin VB.TextBox txtAdjFormula 
            Height          =   360
            HelpContextID   =   319
            Left            =   108
            TabIndex        =   23
            Top             =   3552
            Width           =   8544
         End
         Begin VB.TextBox txtAdjCond 
            Height          =   360
            HelpContextID   =   320
            Left            =   108
            TabIndex        =   24
            Top             =   4188
            Width           =   8544
         End
         Begin VB.TextBox txtFormula 
            Height          =   360
            HelpContextID   =   317
            Left            =   108
            TabIndex        =   21
            Top             =   2256
            Width           =   8544
         End
         Begin VB.TextBox txtBonusCodeDesc 
            Height          =   360
            HelpContextID   =   302
            Left            =   1560
            TabIndex        =   5
            Top             =   288
            Width           =   5496
         End
         Begin VB.TextBox txtBonusCode 
            Height          =   360
            HelpContextID   =   300
            Left            =   108
            TabIndex        =   3
            Top             =   288
            Width           =   1020
         End
         Begin VB.TextBox txtVariable1 
            Height          =   360
            HelpContextID   =   311
            Left            =   108
            TabIndex        =   14
            Top             =   1608
            Width           =   2064
         End
         Begin VB.TextBox txtDollar 
            Height          =   360
            HelpContextID   =   307
            Left            =   1836
            TabIndex        =   10
            Top             =   948
            Width           =   1608
         End
         Begin VB.TextBox txtCondition 
            Height          =   360
            HelpContextID   =   318
            Left            =   108
            TabIndex        =   22
            Top             =   2904
            Width           =   8544
         End
         Begin FACTFRMLib.FactorFrame cmdBonusCode 
            Height          =   360
            HelpContextID   =   301
            Left            =   1140
            TabIndex        =   4
            TabStop         =   0   'False
            Tag             =   "Run #3"
            Top             =   288
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
            Picture         =   "ZZSEBFMT.frx":0000
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
         Begin FACTFRMLib.FactorFrame cmdBonusCodeDesc 
            Height          =   360
            HelpContextID   =   303
            Left            =   7068
            TabIndex        =   6
            TabStop         =   0   'False
            Tag             =   "Run #3"
            Top             =   288
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
            Picture         =   "ZZSEBFMT.frx":0112
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
         Begin FACTFRMLib.FactorFrame cmdVariable1 
            Height          =   360
            HelpContextID   =   312
            Left            =   2184
            TabIndex        =   15
            TabStop         =   0   'False
            Tag             =   "Run #3"
            Top             =   1608
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
            Picture         =   "ZZSEBFMT.frx":0224
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
         Begin FACTFRMLib.FactorFrame cmdVariable2 
            Height          =   360
            HelpContextID   =   314
            Left            =   4680
            TabIndex        =   17
            TabStop         =   0   'False
            Tag             =   "Run #3"
            Top             =   1608
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
            Picture         =   "ZZSEBFMT.frx":0336
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
         Begin FACTFRMLib.FactorFrame cmdVariable3 
            Height          =   360
            HelpContextID   =   316
            Left            =   7176
            TabIndex        =   19
            TabStop         =   0   'False
            Tag             =   "Run #3"
            Top             =   1608
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
            Picture         =   "ZZSEBFMT.frx":0448
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
         Begin FACTFRMLib.FactorFrame cmdLevel 
            Height          =   360
            HelpContextID   =   305
            Left            =   8292
            TabIndex        =   8
            TabStop         =   0   'False
            Tag             =   "Run #3"
            Top             =   288
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
            Picture         =   "ZZSEBFMT.frx":055A
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
         Begin VB.Label txtBonusType 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   360
            Left            =   7596
            TabIndex        =   20
            Top             =   1608
            Width           =   1056
         End
         Begin VB.Label lblType 
            Caption         =   "Comm Type"
            Height          =   276
            Left            =   7596
            TabIndex        =   50
            Top             =   1356
            Width           =   1140
         End
         Begin VB.Label lblLevel 
            Caption         =   "Level"
            Height          =   276
            Left            =   7500
            TabIndex        =   49
            Top             =   36
            Width           =   1152
         End
         Begin VB.Label lblMaxAmount 
            Caption         =   "Max. Total (mxt)"
            Height          =   276
            Left            =   7032
            TabIndex        =   47
            Top             =   708
            Width           =   1512
         End
         Begin VB.Label lblAmount2 
            Caption         =   "Amount2 (amt2)"
            Height          =   276
            Left            =   5328
            TabIndex        =   46
            Top             =   708
            Width           =   1500
         End
         Begin VB.Label lblAmount1 
            Caption         =   "Amount1 (amt1)"
            Height          =   276
            Left            =   3588
            TabIndex        =   45
            Top             =   708
            Width           =   1500
         End
         Begin VB.Label lblVariable1 
            Caption         =   "Variable1 (v1)"
            Height          =   276
            Left            =   120
            TabIndex        =   44
            Top             =   1368
            Width           =   2004
         End
         Begin VB.Label lblVariable2 
            Caption         =   "Variable2 (v2)"
            Height          =   276
            Left            =   2628
            TabIndex        =   43
            Top             =   1368
            Width           =   2004
         End
         Begin VB.Label lblVariable3 
            Caption         =   "Variable3 (v3)"
            Height          =   276
            Left            =   5124
            TabIndex        =   42
            Top             =   1368
            Width           =   2004
         End
         Begin VB.Label lblFormula 
            Caption         =   "Formula"
            Height          =   276
            Left            =   120
            TabIndex        =   41
            Top             =   2028
            Width           =   2004
         End
         Begin VB.Label lblCondition 
            Caption         =   "Condition"
            Height          =   276
            Left            =   120
            TabIndex        =   40
            Top             =   2676
            Width           =   2004
         End
         Begin VB.Label lblAdjustment 
            Caption         =   "Adjustment Formula"
            Height          =   276
            Left            =   120
            TabIndex        =   39
            Top             =   3312
            Width           =   2004
         End
         Begin VB.Label lblAdjCond 
            Caption         =   "Adjustment Condition"
            Height          =   276
            Left            =   120
            TabIndex        =   38
            Top             =   3936
            Width           =   2004
         End
         Begin VB.Label lblBonusCode 
            Caption         =   "Comm Code"
            Height          =   276
            Left            =   120
            TabIndex        =   36
            Top             =   36
            Width           =   1332
         End
         Begin VB.Label lblDollar 
            Caption         =   "Dollar (dol)"
            Height          =   276
            Left            =   1848
            TabIndex        =   34
            Top             =   708
            Width           =   1500
         End
         Begin VB.Label lblPercent 
            Caption         =   "Percent (pct)"
            Height          =   276
            Left            =   120
            TabIndex        =   32
            Top             =   708
            Width           =   1500
         End
         Begin VB.Label lblBonusDesc 
            Caption         =   "Commission Code Description"
            Height          =   276
            Left            =   1572
            TabIndex        =   30
            Top             =   36
            Width           =   3060
         End
      End
      Begin FACTFRMLib.FactorFrame cmdExitCancelBtn 
         Height          =   396
         HelpContextID   =   15
         Left            =   7524
         TabIndex        =   35
         Top             =   4740
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
      Begin FACTFRMLib.FactorFrame cmdDelete 
         Height          =   396
         HelpContextID   =   12
         Left            =   3024
         TabIndex        =   37
         Top             =   4740
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
         Left            =   1536
         TabIndex        =   2
         Top             =   4740
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
         Left            =   48
         TabIndex        =   1
         Top             =   4740
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
      Begin FACTFRMLib.FactorFrame cmdRefresh 
         Height          =   396
         HelpContextID   =   14
         Left            =   6024
         TabIndex        =   33
         Top             =   4740
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
      Begin FACTFRMLib.FactorFrame cmdUpdateInsertBtn 
         Height          =   396
         HelpContextID   =   13
         Left            =   4524
         TabIndex        =   31
         Top             =   4740
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
      Begin DBTrueGrid.TDBGrid tblComboDropdown 
         Bindings        =   "ZZSEBFMT.frx":066C
         Height          =   2124
         Left            =   216
         OleObjectBlob   =   "ZZSEBFMT.frx":068B
         TabIndex        =   48
         TabStop         =   0   'False
         Top             =   144
         Visible         =   0   'False
         Width           =   4896
      End
      Begin FACTFRMLib.FactorFrame efraEditSelect 
         Height          =   4632
         Left            =   60
         TabIndex        =   51
         TabStop         =   0   'False
         Top             =   60
         Width           =   8760
         _Version        =   65536
         _ExtentX        =   15452
         _ExtentY        =   8170
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
         Begin DBTrueGrid.TDBGrid tblEditSelect 
            Height          =   4452
            HelpContextID   =   227
            Left            =   84
            OleObjectBlob   =   "ZZSEBFMT.frx":197F
            TabIndex        =   52
            Top             =   84
            Width           =   8592
         End
      End
   End
   Begin FACTFRMLib.FactorFrame ffraStatusbar 
      Height          =   360
      Left            =   0
      TabIndex        =   25
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
      TabIndex        =   27
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
      FMName          =   "ZZSEBFMT"
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
   Begin SQAOTestObjectsCtl.SQAOTest SQAOTest1 
      Height          =   456
      Left            =   10884
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   96
      Width           =   456
      _ExtentX        =   804
      _ExtentY        =   804
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
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "&Edit"
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "&Delete"
      End
      Begin VB.Menu mnuUpdateInsert 
         Caption         =   "&Update"
      End
      Begin VB.Menu mnuRefreshSelect 
         Caption         =   "&Refresh"
      End
      Begin VB.Menu mnuOptionsSep15 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditSelect 
         Caption         =   "Show Selection Grid"
         Checked         =   -1  'True
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
Attribute VB_Name = "frmZZSEBFMT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***********************************************************************
'Copyright (c) 1999 FACTOR, A Division of W.R.Hess Company             *
'Program ID     : ZZSEBFMT                                             *
'Programmer     : Rajneesh Aggarwal                                    *
'***********************************************************************
Option Explicit

Private t_bStartupFlag As Boolean 'optional startup flag
Private t_bDataChanged As Boolean 'data changed flag
Private t_bUpdateTable As Boolean 'update data flag

Private t_nFormMode As Integer         'global used to track the current form operating mode
Private Const IDLE_MODE As Integer = 0 'idle mode activates the NoDrop Cursor
Private Const ADD_MODE As Integer = 1    'flag set when in the add mode
Private Const EDIT_MODE As Integer = 2   'flag set when in the edit mode

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
Private Const t_szEXIT As String = "Exit"
Private Const t_szCANCEL As String = "Cancel"

Private Const t_szPRINT As String = "Print"
Private Const t_szHELP As String = "Help"
Private cValidate As cValidateInput

Private tgmEditSelect As clsTGSpreadSheet
Private tgsEditSelect  As clsTGSelector

Private Const nColCode As Integer = 0
Private Const nColDesc As Integer = 1
Private Const nColLevel As Integer = 2
Private Const nColFormula As Integer = 3
Private Const nColCondition As Integer = 4
Private Const nColType As Integer = 5
Private Const nColPercent As Integer = 6
Private Const nColDollar As Integer = 7
Private Const nColAmt1 As Integer = 8
Private Const nColAmt2 As Integer = 9
Private Const nColVar1 As Integer = 10
Private Const nColVar2 As Integer = 11
Private Const nColVar3 As Integer = 12
Private Const nColMaxTotal As Integer = 13
Private Const nColAdjFormula As Integer = 14
Private Const nColAdjCondition As Integer = 15

Private Const sRegExp80 As String = "^(P{1,80})$"

Private bDataLoaded As Boolean
'

Private Sub cmdAddBtn_Click()
    #If PROTOTYPE Then
        Exit Sub
    #End If
    
    subEnableEditBtn False
    subEnableAddBtn False
    cmdUpdateInsertBtn.Caption = t_szCAPTION_INSERT
    subSetExitCancelBtn "CANCEL"
    t_nFormMode = ADD_MODE
    subEnableFirstLine True
    'txtPercent.Enabled = True
    subEnableSearchbtn cmdLevel, False
    DoEvents
    subSetFocus txtBonusCode
    
End Sub

Private Sub cmdDelete_Click()
    If Not tfnCancelExit("Are you sure you want to delete the current record?") Then
        Exit Sub
    End If
    
    If Not fnDeleteBonusFormula(txtBonusCode, txtLevel.Tag) Then
        tfnSetStatusBarError "Failed to delete the Commission formula"
        Exit Sub
    End If
    
    tfnResetScreen
    
End Sub

Private Sub cmdEditBtn_Click()
    #If PROTOTYPE Then
        Exit Sub
    #End If
    
    Dim bGoEdit As Boolean
    
    If mnuEditSelect.Enabled And mnuEditSelect.CHECKED = True Then
        mnuEditSelect.Enabled = False
        If Not fnLoadEditSelectGrid() Then
            Exit Sub
        End If
    Else
        bGoEdit = True
    End If
    
    If bGoEdit Then
        efraEditSelect.Visible = False
        efraBase.Visible = True
        efraBase.ZOrder 0
        subEnableEditBtn False
        subEnableAddBtn False
        cmdUpdateInsertBtn.Caption = t_szCAPTION_UPDATE
        cmdRefresh.Caption = t_szCAPTION_REFRESH
        mnuRefreshSelect.Caption = t_szCAPTION_REFRESH
        subSetExitCancelBtn "CANCEL"
        t_nFormMode = EDIT_MODE
        subEnableFirstLine True
        DoEvents
        subSetFocus txtBonusCode
    Else  'show edit selection popup grid
        If mnuEditSelect.CHECKED = True Then
            efraBase.Visible = False
            efraEditSelect.Visible = True
            efraEditSelect.ZOrder 0
            
            subEnableEditBtn False
            subEnableAddBtn False
            cmdUpdateInsertBtn.Caption = t_szCAPTION_UPDATE
            cmdRefresh.Caption = t_szCAPTION_SELECT
            mnuRefreshSelect.Caption = t_szCAPTION_SELECT
            subSetExitCancelBtn "CANCEL"
            subSetFocus tblEditSelect
            subEnableRefreshBtn True
        End If
    End If
End Sub

Private Sub cmdRefresh_Click()
    If cmdRefresh.Caption = t_szCAPTION_REFRESH Then
        If Not tfnCancelExit(t_szREFRESH_MESSAGE) Then
            Exit Sub
        End If
        
        fnLoadBonusFormula txtBonusCode, txtLevel.Tag
        subEnableUpdateBtn False
        subEnableDeleteBtn True
        
        txtLevel.Enabled = True
        subSetFocus txtLevel
    Else  'Select
        Dim lCount As Long
        Dim lTemp() As Long
        
        tgsEditSelect.GetSelected lTemp, lCount
        
        If lCount < 0 Then
            Exit Sub
        End If
        
        t_nFormMode = EDIT_MODE

        'fill the key fields textboxes
        txtBonusCode = tgmEditSelect.CellValue(nColCode, lTemp(0))
        txtBonusCodeDesc = tgmEditSelect.CellValue(nColDesc, lTemp(0))
        txtLevel = tgmEditSelect.CellValue(nColLevel, lTemp(0))
        
        cmdRefresh.Caption = t_szCAPTION_REFRESH
        mnuRefreshSelect.Caption = t_szCAPTION_REFRESH
        
        efraEditSelect.Visible = False
        efraBase.Visible = True
        efraBase.ZOrder 0
        
        cValidate.ResetFlags txtBonusCode, True
        cValidate.ResetFlags txtBonusCodeDesc, True
        cValidate.ResetFlags txtLevel, True
        
        subEnterStageII
        subEnableRefreshBtn False
        
        DoEvents
        subSetFocus txtLevel
    End If
End Sub

Private Sub cmdRefresh_GotFocus()
    If cmdRefresh.Caption = t_szCAPTION_REFRESH Then
        tfnSetStatusBarMessage "Refresh"
    Else
        tfnSetStatusBarMessage "Select"
    End If
End Sub

Private Sub cmdUpdateInsertBtn_Click()
    
    If t_nFormMode = ADD_MODE Then
        If Not fnInsertBonusFormula() Then
            tfnSetStatusBarError "Failed to insert the Commission formula"
            Exit Sub
        End If
    Else
        If Not fnUpdateBonusFormula() Then
            tfnSetStatusBarError "Failed to update the Commission formula"
            Exit Sub
        End If
    End If
    
    tfnResetScreen

End Sub

Private Sub cmdVariable1_Click()
    tgcDropdown.ComboSQL(txtVariable1) = fnSetComboSQL(txtVariable1.TabIndex)
    tgcDropdown.Click cmdVariable1
End Sub

Private Sub cmdVariable2_Click()
    tgcDropdown.ComboSQL(txtVariable2) = fnSetComboSQL(txtVariable2.TabIndex)
    tgcDropdown.Click cmdVariable2
End Sub

Private Sub cmdVariable3_Click()
    tgcDropdown.ComboSQL(txtVariable3) = fnSetComboSQL(txtVariable3.TabIndex)
    tgcDropdown.Click cmdVariable3
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
        If tfnAuthorizeExecute(Command) = False Then 'Check for handshake if not in the development mode
            Unload Me
            Exit Sub
        End If
        
        'open the database, ODBC Dialog Box during developemnt, oleObject Connection String when not
        If Not tfnOpenDatabase(False, sErrorMessage) Then
            MsgBox "Unable to open Database. Program terminates.", vbCritical
            Unload Me
            Exit Sub
        End If
        
        'connect to local database
        Set dbLocal = tfnOpenLocalDatabase(False, sErrorMessage)
        If dbLocal Is Nothing Then
            MsgBox "Unable to open Local Database. Program terminates.", vbCritical
            Unload Me
            Exit Sub
        End If
    #End If
    
    subInitErrorHandler   ' Setup Error Control
     
    #If Not PROTOTYPE Then
        tfnUpdateVersion
        If Not fnCreateTempTableVar() Then
            MsgBox "Failed to create temporary Variable Table. Program terminates.", vbCritical
            Unload Me
            Exit Sub
        End If
    #End If
    
    tfnDisableFormSystemClose Me
    subSetupToolBar
    subSetupCombos
    subInitValidation
    subInitSpreadsheets
    
    tmrKeyBoard.Enabled = False
    tfnCenterForm Me
    
    Screen.MousePointer = vbHourglass
    Me.Enabled = False
    DoEvents

    Screen.MousePointer = vbHourglass
    Me.Enabled = True
    
    t_bStartupFlag = False
    
    tfnResetScreen
    tmrKeyBoard.Enabled = True
    Screen.MousePointer = vbDefault
    
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

    If cmdExitCancelBtn.Caption = t_szCAPTION_CANCEL Then
        subCancel
    Else
        subExit
    End If
    
End Sub


'=====================
'Toolbar Button Events
'=====================
Private Sub subCancel()
    Dim bConfirm As Boolean
    
    If t_nFormMode = ADD_MODE Then
        bConfirm = cValidate.ValidInput(txtBonusCode) And _
            cValidate.ValidInput(txtLevel) And t_bDataChanged
    Else
        If efraEditSelect.Visible Then
            bConfirm = False
        Else
            bConfirm = bDataLoaded And t_bDataChanged
        End If
    End If
    
    If bConfirm Then
        If Not tfnCancelExit(t_szEXIT_MESSAGE) Then
            Exit Sub
        End If
    End If
    
    subSetFocus efraBackground
    
    If efraEditSelect.Visible Then
        efraEditSelect.Visible = False
        efraBase.Visible = True
        efraBase.ZOrder 0
        #If CANCEL_EDIT Then
            Screen.MousePointer = vbHourglass
            tfnResetScreen
        #Else
            cmdEditBtn_Click
        #End If
    Else
        Screen.MousePointer = vbHourglass
        tfnResetScreen
    End If
End Sub

Private Sub subExit()
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
Private Sub mnuAdd_Click()
    cmdAddBtn_Click
End Sub

Private Sub mnuEdit_Click()
    cmdEditBtn_Click
End Sub

Private Sub mnuDelete_Click()
    cmdDelete_Click
End Sub

Private Sub mnuUpdateInsert_Click()
    cmdUpdateInsertBtn_Click
End Sub

Private Sub mnuRefreshSelect_Click()
    cmdRefresh_Click
End Sub

Private Sub mnuEditSelect_Click()
    If mnuEditSelect.CHECKED Then
        mnuEditSelect.CHECKED = False
    Else
        mnuEditSelect.CHECKED = True
    End If
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
    subCancel
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
Private Sub tfnResetScreen()
    
    On Error Resume Next
    
    frmContext.ButtonEnabled(PRINT_UP) = False
    frmContext.ButtonEnabled(COPY_UP) = False
    t_nFormMode = IDLE_MODE
    
    txtBonusCode = ""
    txtBonusCodeDesc = ""
    txtLevel = ""
    txtPercent = ""
    txtDollar = ""
    txtAmount1 = ""
    txtAmount2 = ""
    txtMaxTotal = ""
    txtVariable1 = ""
    txtVariable2 = ""
    txtVariable3 = ""
    txtBonusType = ""
    txtFormula = ""
    txtCondition = ""
    txtAdjFormula = ""
    txtAdjCond = ""
    
    'for storing the old value
    txtLevel.Tag = ""
    txtPercent.Tag = ""
    txtDollar.Tag = ""
    txtAmount1.Tag = ""
    txtAmount2.Tag = ""
    txtMaxTotal.Tag = ""
    txtVariable1.Tag = ""
    txtVariable2.Tag = ""
    txtVariable3.Tag = ""
    txtFormula.Tag = ""
    txtCondition.Tag = ""
    txtAdjFormula.Tag = ""
    txtAdjCond.Tag = ""
    
    tgcDropdown.ComboOn(txtBonusCode) = True
    tgcDropdown.ComboOn(txtLevel) = True
    
    tblEditSelect.Enabled = False
    
    t_bDataChanged = False
    bDataLoaded = False
    
    subEnableRefreshBtn False
    subEnableUpdateBtn False
    subEnableDeleteBtn False
    
    cValidate.ResetFlags
    subEnableControls False
    
    mnuEditSelect.Enabled = True
    mnuExit.Enabled = True
    mnuPrint.Enabled = False
    mnuCopy.Enabled = False
    mnuPaste.Enabled = False
    subEnableAddBtn True
    subEnableEditBtn True
    subSetExitCancelBtn "EXIT"
    
    cmdRefresh.Caption = t_szCAPTION_REFRESH
    mnuRefreshSelect.Caption = t_szCAPTION_REFRESH
    
    Screen.MousePointer = vbDefault
    
    subSetFocus cmdAddBtn

    Set cValidate.LastBox = txtBonusCode
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

Private Sub tblComboDropDown_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    tgcDropdown.TableMouseUp Y
End Sub

Private Sub tblComboDropDown_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    tgcDropdown.RowColChange
End Sub

Private Sub tblComboDropDown_SelChange(CANCEL As Integer)
    tgcDropdown.SelChange CANCEL
End Sub

Private Sub tblEditSelect_DblClick()
    If cmdRefresh.Enabled Then
        cmdRefresh_Click
    End If
End Sub

Private Sub tblEditSelect_HeadClick(ByVal ColIndex As Integer)
    tgmEditSelect.HeadClick ColIndex
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

Public Sub subSetExitCancelBtn(sExitCancel As String)
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
            '.AddButton "Add &Vendor Cross Reference", EDIVNDXR_UP
            '.AddButton "Add &UOM Cross Reference", EDIUOMXR_UP
        .EndSetupToolbar
    
        .HelpFile = szHelpFileName
    End With
End Sub

Public Sub TBButtonCallBack(ByVal nID As Integer)
    Select Case nID
        Case CANCEL_UP
            subCancel
        Case EXIT_UP
            subExit
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

Private Sub subSetupCombos()
    Set tgcDropdown = CreateObject(t_szOLECOMBO)
    Set tgcDropdown.Form = Me
    Set tgcDropdown.DBEngine = t_engFactor
    Set tgcDropdown.Database = t_dbMainDatabase
    Set tgcDropdown.DataLink = datComboDropDown
    Set tgcDropdown.Table = tblComboDropdown
    
    #If PROTOTYPE Then
        Exit Sub
    #End If
    
     With tgcDropdown
        .AddCombo
        .AddComboBox txtBonusCode, cmdBonusCode, "bc_bonus_code", .SQL_STRING_TYPE(4)
        .AddComboBox txtBonusCodeDesc, cmdBonusCodeDesc, "bc_code_desc", .SQL_STRING_TYPE(40)
        .AddExtraColumn "bc_type"
        'david 02/12/2001
        'ole combo not working???
        .SetColumnCaptions "Bonus Code", "Bonus Code Description", "Type"
        
        .AddCombo
        .AddComboBox txtLevel, cmdLevel, "bf_level", .SQL_INT_TYPE
     
        .AddCombo , 10
        .AddComboBox txtVariable1, cmdVariable1, "variable", .SQL_STRING_TYPE(18)
        
        .AddCombo
        .AddComboBox txtVariable2, cmdVariable2, "variable", .SQL_STRING_TYPE(18)
        
        .AddCombo
        .AddComboBox txtVariable3, cmdVariable3, "variable", .SQL_STRING_TYPE(18)
     End With
     
End Sub

Private Sub txtBonusCode_Change()
    subEnableUpdateBtn False
    
    cValidate.Change txtBonusCode
    tgcDropdown.Change txtBonusCode
    txtBonusType.Caption = ""
    
    If Not ActiveControl Is txtBonusCode Then Exit Sub
    
    t_bDataChanged = True
    subEnableRefreshBtn True
End Sub

Private Sub txtBonusCode_GotFocus()
    If txtAdjCond.Enabled Then
        Set cValidate.LastBox = txtAdjCond
    ElseIf txtLevel.Enabled Then
        Set cValidate.LastBox = txtLevel
    Else
        Set cValidate.LastBox = txtBonusCode
    End If
    
    tgcDropdown.GotFocus txtBonusCode
    cValidate.GotFocus txtBonusCode
    
    If tgcDropdown.SingleRecordSelected Then
        subSetFocus txtLevel
    End If
    
End Sub

Private Sub txtBonusCode_KeyPress(KeyAscii As Integer)
    Dim bKeyCode As Boolean
    
    KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
    
    If KeyAscii = vbKeyReturn Then
        tgcDropdown.ComboSQL(txtBonusCode) = fnSetComboSQL(txtBonusCode.TabIndex)
        Screen.MousePointer = vbHourglass
    End If
        
    bKeyCode = tgcDropdown.Keypress(txtBonusCode, KeyAscii)
    Screen.MousePointer = vbDefault
    
    If Not bKeyCode Then
       If KeyAscii = vbKeyReturn Then
          If tgcDropdown.SingleRecordSelected Then
                subSetFocus txtLevel
          End If
          KeyAscii = 0
       End If
    Else
        cValidate.Keypress txtBonusCode, KeyAscii
    End If

End Sub

Private Sub txtBonusCode_LostFocus()
    tgcDropdown.LostFocus txtBonusCode
    If cValidate.LostFocus(txtBonusCode, cmdBonusCode, txtBonusCodeDesc, cmdBonusCodeDesc, tblComboDropdown) Then
        If txtBonusCode <> "" Then
            If t_nFormMode = ADD_MODE Then
                t_bDataChanged = True
            End If
        End If
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdBonusCode_Click()
    tgcDropdown.ComboSQL(txtBonusCode) = fnSetComboSQL(txtBonusCode.TabIndex)
    tgcDropdown.Click cmdBonusCode
End Sub

Private Sub txtBonusCodeDesc_Change()
    tgcDropdown.Change txtBonusCodeDesc
    
    If Not ActiveControl Is txtBonusCodeDesc Then Exit Sub
    
    t_bDataChanged = True
    subEnableRefreshBtn True
End Sub

Private Sub txtBonusCodeDesc_GotFocus()
    If txtAdjCond.Enabled Then
        Set cValidate.LastBox = txtAdjCond
    ElseIf txtLevel.Enabled Then
        Set cValidate.LastBox = txtLevel
    Else
        Set cValidate.LastBox = txtBonusCode
    End If
    
'    tfnSetStatusBarMessage "Enter BonusCode Description"
    tgcDropdown.GotFocus txtBonusCodeDesc
    Screen.MousePointer = vbDefault
    
    If tgcDropdown.SingleRecordSelected Then
        subSetFocus txtLevel
    End If
End Sub

Private Sub txtBonusCodeDesc_KeyPress(KeyAscii As Integer)
    Dim bKeyCode As Boolean
    
    If KeyAscii = vbKeyReturn Then
        tgcDropdown.ComboSQL(txtBonusCodeDesc) = fnSetComboSQL(txtBonusCodeDesc.TabIndex)
        Screen.MousePointer = vbHourglass
    End If
    
    bKeyCode = tgcDropdown.Keypress(txtBonusCodeDesc, KeyAscii)
    Screen.MousePointer = vbDefault
    
    If Not bKeyCode Then
        If KeyAscii = vbKeyReturn Then
            If tgcDropdown.SingleRecordSelected Then
                subSetFocus txtLevel
            End If
        KeyAscii = 0
        End If
    End If

End Sub

Private Sub txtBonusCodeDesc_LostFocus()
    tgcDropdown.LostFocus txtBonusCodeDesc
    If cValidate.LostFocus(txtBonusCode, cmdBonusCode, txtBonusCodeDesc, cmdBonusCodeDesc, tblComboDropdown) Then
        If cValidate.FirstInvalidInput < 0 Then
            'cmd
        End If
    End If
End Sub

Private Sub cmdBonusCodeDesc_Click()
    tgcDropdown.ComboSQL(txtBonusCodeDesc) = fnSetComboSQL(txtBonusCodeDesc.TabIndex)
    tgcDropdown.Click cmdBonusCodeDesc
End Sub

Private Function fnSetComboSQL(nTabIndex) As String
    Dim strSQL As String
    
    Select Case nTabIndex
        Case txtBonusCode.TabIndex, txtBonusCodeDesc.TabIndex
            If t_nFormMode = ADD_MODE Then
                strSQL = "SELECT DISTINCT bc_bonus_code, bc_code_desc, bc_type FROM bonus_codes"
                If tfnRound(txtLevel) > 0 Then
                    strSQL = strSQL + " WHERE bc_bonus_code NOT IN (SELECT bf_bonus_code FROM"
                    strSQL = strSQL + " bonus_formula WHERE bf_level = " & tfnRound(txtLevel) & ")"
                End If
            Else
                strSQL = "SELECT DISTINCT bc_bonus_code, bc_code_desc, bc_type FROM bonus_codes, bonus_formula"
                strSQL = strSQL + " WHERE bc_bonus_code = bf_bonus_code"
            End If
        Case txtLevel.TabIndex
            strSQL = "SELECT bf_level FROM bonus_formula"
            If txtBonusCode <> "" Then
                strSQL = strSQL & " WHERE bf_bonus_code = " & tfnSQLString(Trim(txtBonusCode))
            End If
        Case txtVariable1.TabIndex, txtVariable2.TabIndex, txtVariable3.TabIndex
            strSQL = "SELECT variable FROM tmpVariable"
    End Select
    
    fnSetComboSQL = strSQL
End Function

Private Sub subInitValidation()
    Set cValidate = New cValidateInput
    Set cValidate.Form = Me
    Set cValidate.StatusBar = ffraStatusbar
    
    cValidate.AddEditBox txtBonusCode, "Enter Commission Code"
    cValidate.AddEditBox txtBonusCodeDesc, "Enter Commission Code Description"
    cValidate.AddEditBox txtLevel, "Enter Level"
    cValidate.AddEditBox txtPercent, "Enter Percentage"
    cValidate.AddEditBox txtDollar, "Enter Dollar Amount"
    cValidate.AddEditBox txtAmount1, "Enter Amount 1"
    cValidate.AddEditBox txtAmount2, "Enter Amount 2"
    cValidate.AddEditBox txtMaxTotal, "Enter Maxiumum Total"
    cValidate.AddEditBox txtVariable1, "Enter Variable 1"
    cValidate.AddEditBox txtVariable2, "Enter Variable 2"
    cValidate.AddEditBox txtVariable3, "Enter Variable 3"
    cValidate.AddEditBox txtFormula, "Enter Formula"
    cValidate.AddEditBox txtCondition, "Enter Condition"
    cValidate.AddEditBox txtAdjFormula, "Enter Adjustment Formula"
    cValidate.AddEditBox txtAdjCond, "Enter Adjustment Condition"
    cValidate.MinTabIndex = tbToolbar.TabIndex
    
    Set cValidate.ControlForFocus = efraBase
    Set cValidate.LastBox = txtBonusCode
    cValidate.SetFirstControls cmdUpdateInsertBtn, cmdRefresh, cmdDelete, cmdExitCancelBtn
    cValidate.MaxTabIndex = efraBase.TabIndex + 1
    
End Sub

Public Function fnInvalidData(txtBox As Textbox) As Boolean
    #If PROTOTYPE Then
        Exit Function
    #End If

    Select Case txtBox.TabIndex
        Case txtBonusCode.TabIndex, txtBonusCodeDesc.TabIndex
            fnInvalidData = Not fnValidBonusCode(txtBox)
        Case txtLevel.TabIndex
            fnInvalidData = Not fnValidLevel(txtBox)
        Case txtVariable1.TabIndex, txtVariable2.TabIndex, txtVariable3.TabIndex
            fnInvalidData = Not fnValidVariables(txtBox)
        Case txtFormula.TabIndex, txtAdjFormula.TabIndex
            fnInvalidData = Not fnValidFormula(txtBox)
        Case txtCondition.TabIndex, txtAdjCond.TabIndex
            fnInvalidData = Not fnValidCondition(txtBox)
        Case Else
            fnInvalidData = False
    End Select
    
    If Not fnInvalidData Then
        If Not (txtBox Is txtCondition Or txtBox Is txtFormula Or txtBox Is txtAdjCond Or txtBox Is txtAdjFormula) Then
            cValidate.ResetFlags txtCondition
            cValidate.ResetFlags txtFormula
            cValidate.ResetFlags txtAdjCond
            cValidate.ResetFlags txtAdjFormula
            cValidate.ResetFlags txtCondition
        End If
    End If
End Function

Private Function fnValidBonusCode(Box As Textbox) As Boolean
    Const SUB_NAME As String = "fnValidBonusCode"
    Dim strSQL As String
    Dim rsTemp As Recordset
    
    If Box.Name = "txtBonusCodeDesc" Then
        fnValidBonusCode = True
        Exit Function
    End If
    
    fnValidBonusCode = False
    
    If Trim(Box.Text) = "" Then
        cValidate.SetErrorMessage Box, "Commission Code cannot be blank."
        Exit Function
    End If
   
    strSQL = "SELECT bc_type FROM bonus_codes WHERE bc_bonus_code = "
    strSQL = strSQL & tfnSQLString(Box.Text)
    
    If GetRecordSet(rsTemp, strSQL, , SUB_NAME) < 0 Then
        cValidate.SetErrorMessage Box, "Failed to access Database"
        Exit Function
    End If
    
    If rsTemp.RecordCount = 0 Then
        cValidate.SetErrorMessage Box, "Commission Code does not exist"
        Exit Function
    End If
    
    If rsTemp.RecordCount = 1 Then
        txtBonusType.Caption = fnGetField(rsTemp!bc_type)
    End If
    
    If tfnRound(txtLevel) <> 0 Then
        strSQL = "SELECT * FROM bonus_formula WHERE bf_bonus_code = " & tfnSQLString(Trim(Box))
        strSQL = strSQL & " AND bf_level = " & tfnRound(txtLevel)
        
        If GetRecordCount(strSQL, , SUB_NAME) = 1 Then
            If t_nFormMode = ADD_MODE Then
                cValidate.SetErrorMessage Box, "Comm. formula, Level " & txtLevel & " already exists"
                Exit Function
            End If
        End If
    End If
    
    fnValidBonusCode = True

End Function

Private Function fnInsertBonusFormula() As Boolean
    
    Const SUB_NAME As String = "fnInsertBonusFormula"
    Dim strSQL As String
    
    fnInsertBonusFormula = False
    
    strSQL = "INSERT INTO bonus_formula(bf_bonus_code, bf_level, bf_percent, bf_dollar,"
    strSQL = strSQL & " bf_amount1, bf_amount2, bf_formula, bf_condition, bf_variable1,"
    strSQL = strSQL & " bf_variable2, bf_variable3, bf_max_total, bf_adj_formula, "
    strSQL = strSQL & " bf_adj_condition) VALUES(" & tfnSQLString(Trim(txtBonusCode)) & ","
    strSQL = strSQL & tfnRound(txtLevel) & ", "
    strSQL = strSQL & tfnRound(txtPercent, DEFAULT_DECIMALS) & ", "
    strSQL = strSQL & tfnRound(txtDollar, 2) & ", "
    strSQL = strSQL & tfnRound(txtAmount1, DEFAULT_DECIMALS) & ", "
    strSQL = strSQL & tfnRound(txtAmount2, DEFAULT_DECIMALS) & ", "
    strSQL = strSQL & tfnSQLString(Trim(txtFormula)) & ", "
    strSQL = strSQL & tfnSQLString(Trim(txtCondition)) & ", "
    strSQL = strSQL & tfnSQLString(Trim(txtVariable1)) & ", "
    strSQL = strSQL & tfnSQLString(Trim(txtVariable2)) & ", "
    strSQL = strSQL & tfnSQLString(Trim(txtVariable3)) & ", "
    strSQL = strSQL & tfnRound(txtMaxTotal, 2) & ", "
    strSQL = strSQL & tfnSQLString(Trim(txtAdjFormula)) & ", "
    strSQL = strSQL & tfnSQLString(Trim(txtAdjCond)) & ")"
    
    If fnExecuteSQL(strSQL, , SUB_NAME) Then
        fnInsertBonusFormula = True
    End If
    
End Function

Private Function fnUpdateBonusFormula() As Boolean
    
    Const SUB_NAME As String = "fnUpdateBonusFormula"
    Dim strSQL As String
    
    fnUpdateBonusFormula = False
    
    strSQL = "UPDATE bonus_formula SET"
    strSQL = strSQL & " bf_level = " & tfnRound(txtLevel, DEFAULT_DECIMALS) & ","
    strSQL = strSQL & " bf_percent = " & tfnRound(txtPercent, DEFAULT_DECIMALS) & ","
    strSQL = strSQL & " bf_dollar = " & tfnRound(txtDollar, 2) & ","
    strSQL = strSQL & " bf_amount1 = " & tfnRound(txtAmount1, DEFAULT_DECIMALS) & ","
    strSQL = strSQL & " bf_amount2 = " & tfnRound(txtAmount2, DEFAULT_DECIMALS) & ","
    strSQL = strSQL & " bf_formula = " & tfnSQLString(Trim(txtFormula)) & ","
    strSQL = strSQL & " bf_condition = " & tfnSQLString(Trim(txtCondition)) & ","
    strSQL = strSQL & " bf_variable1 = " & tfnSQLString(Trim(txtVariable1)) & ","
    strSQL = strSQL & " bf_variable2 = " & tfnSQLString(Trim(txtVariable2)) & ","
    strSQL = strSQL & " bf_variable3 = " & tfnSQLString(Trim(txtVariable3)) & ","
    strSQL = strSQL & " bf_max_total = " & tfnRound(txtMaxTotal, 2) & ","
    strSQL = strSQL & " bf_adj_formula = " & tfnSQLString(Trim(txtAdjFormula)) & ","
    strSQL = strSQL & " bf_adj_condition = " & tfnSQLString(Trim(txtAdjCond))
    strSQL = strSQL & " WHERE bf_bonus_code = " & tfnSQLString(txtBonusCode.Tag)
    strSQL = strSQL & " AND bf_level = " & tfnRound(txtLevel.Tag)
    
    If fnExecuteSQL(strSQL, , SUB_NAME) Then
        fnUpdateBonusFormula = True
    End If
    
End Function

Private Function fnLoadBonusFormula(sCode As String, nLevel As Integer) As Boolean
    Const SUB_NAME As String = "fnLoadBonusFormula"
    Dim strSQL As String
    Dim rsTemp As Recordset
    
    fnLoadBonusFormula = False
    
    If Trim(sCode) = "" Or nLevel = 0 Then
        Exit Function
    End If
    
    strSQL = "SELECT * FROM bonus_formula WHERE bf_bonus_code = " & tfnSQLString(Trim(sCode))
    strSQL = strSQL & " AND bf_level = " & tfnRound(nLevel)
    
    If GetRecordSet(rsTemp, strSQL, , SUB_NAME) < 0 Then
        MsgBox "Failed to access database to get Commission formula", vbExclamation
        Exit Function
    End If
    
    If rsTemp.RecordCount = 0 Then
        MsgBox "No record found for the selection criteria.", vbExclamation
        subSetFocus txtBonusCode
        Exit Function
    End If
    
    subEnableControls True
    subEnableFirstLine False
    
    If rsTemp.RecordCount = 1 Then
        txtBonusCodeDesc = fnGetBonusCodeDesc(sCode)
        txtLevel = fnGetField(rsTemp!bf_level)
        txtPercent = tfnRound(rsTemp!bf_percent, DEFAULT_DECIMALS)
        txtDollar = tfnFormatDecimal(rsTemp!bf_dollar, 2)
        txtAmount1 = tfnFormatDecimal(rsTemp!bf_amount1, 2, DEFAULT_DECIMALS)
        txtAmount2 = tfnFormatDecimal(rsTemp!bf_amount2, 2, DEFAULT_DECIMALS)
        txtMaxTotal = tfnRound(rsTemp!bf_max_total, 2)
        txtVariable1 = fnGetField(rsTemp!bf_variable1)
        txtVariable2 = fnGetField(rsTemp!bf_variable2)
        txtVariable3 = fnGetField(rsTemp!bf_variable3)
        txtFormula = fnGetField(rsTemp!bf_formula)
        txtCondition = fnGetField(rsTemp!bf_condition)
        txtAdjFormula = fnGetField(rsTemp!bf_adj_formula)
        txtAdjCond = fnGetField(rsTemp!bf_adj_condition)
        
        'store the old value into Tag property of the textbox
        txtBonusCode.Tag = txtBonusCode
        txtBonusCodeDesc.Tag = txtBonusCodeDesc
        txtLevel.Tag = txtLevel
        txtPercent.Tag = txtPercent
        txtDollar.Tag = txtDollar
        txtAmount1.Tag = txtAmount1
        txtAmount2.Tag = txtAmount2
        txtMaxTotal.Tag = txtMaxTotal
        txtVariable1.Tag = txtVariable1
        txtVariable2.Tag = txtVariable2
        txtVariable3.Tag = txtVariable3
        txtFormula.Tag = txtFormula
        txtCondition.Tag = txtCondition
        txtAdjFormula.Tag = txtAdjFormula
        txtAdjCond.Tag = txtAdjCond
    End If
    
    If txtBonusType.Caption <> "" Then
        subEnableVariables Trim(txtBonusType)
    End If
    
    t_bDataChanged = False
    bDataLoaded = True
    
    subEnableRefreshBtn False
    fnLoadBonusFormula = True
End Function

Private Function fnGetBonusCodeDesc(sCode As String) As String
    Const SUB_NAME As String = ""
    Dim strSQL As String
    Dim rsTemp As Recordset
    
    fnGetBonusCodeDesc = ""
    
    strSQL = "SELECT bc_code_desc FROM bonus_codes WHERE bc_bonus_code = "
    strSQL = strSQL & tfnSQLString(Trim(sCode))
    
    If GetRecordSet(rsTemp, strSQL, , SUB_NAME) < 0 Then
        Exit Function
    End If
    
    If rsTemp.RecordCount = 0 Then
        Exit Function
    End If

    fnGetBonusCodeDesc = fnGetField(rsTemp!bc_code_desc)
End Function

Private Sub cmdUpdateInsertBtn_GotFocus()
    
    If t_nFormMode = ADD_MODE Then
        tfnSetStatusBarMessage ("Insert")
    Else
        tfnSetStatusBarMessage ("Update")
    End If

End Sub

Private Sub subEnableControls(bYesNo As Boolean)

    txtBonusCode.Enabled = bYesNo
    subEnableSearchbtn cmdBonusCode, bYesNo
    txtBonusCodeDesc.Enabled = bYesNo
    subEnableSearchbtn cmdBonusCodeDesc, bYesNo
    txtLevel.Enabled = bYesNo
    subEnableSearchbtn cmdLevel, bYesNo
    txtPercent.Enabled = bYesNo
    txtDollar.Enabled = bYesNo
    txtAmount1.Enabled = bYesNo
    txtAmount2.Enabled = bYesNo
    txtMaxTotal.Enabled = bYesNo
    txtVariable1.Enabled = bYesNo
    subEnableSearchbtn cmdVariable1, bYesNo
    txtVariable2.Enabled = bYesNo
    subEnableSearchbtn cmdVariable2, bYesNo
    txtVariable3.Enabled = bYesNo
    subEnableSearchbtn cmdVariable3, bYesNo
    txtFormula.Enabled = bYesNo
    txtCondition.Enabled = bYesNo
    txtAdjFormula.Enabled = bYesNo
    txtAdjCond.Enabled = bYesNo
    
End Sub

Private Sub cmdAddBtn_GotFocus()
    tfnSetStatusBarMessage ADD_EDIT_MSG
End Sub

Private Sub cmdEditBtn_GotFocus()
    tfnSetStatusBarMessage ADD_EDIT_MSG
End Sub

Private Sub subEnableFirstLine(bYesNo As Boolean)
    txtBonusCode.Enabled = bYesNo
    subEnableSearchbtn cmdBonusCode, bYesNo
    txtBonusCodeDesc.Enabled = bYesNo
    subEnableSearchbtn cmdBonusCodeDesc, bYesNo
    txtLevel.Enabled = bYesNo
    subEnableSearchbtn cmdLevel, bYesNo
End Sub

Private Sub txtlevel_Change()
    subEnableUpdateBtn False
    
    cValidate.Change txtLevel
    tgcDropdown.Change txtLevel

    If Not ActiveControl Is txtLevel Then Exit Sub
    
    t_bDataChanged = True
    subEnableRefreshBtn True
End Sub

Private Sub txtlevel_GotFocus()
    If txtAdjCond.Enabled Then
        Set cValidate.LastBox = txtAdjCond
    ElseIf txtLevel.Enabled Then
        Set cValidate.LastBox = txtLevel
    Else
        Set cValidate.LastBox = txtBonusCode
    End If
    
    tgcDropdown.GotFocus txtLevel
    cValidate.GotFocus txtLevel
    
    If tgcDropdown.SingleRecordSelected Then
        subEnterStageII
    End If
    
End Sub

Private Sub txtlevel_KeyPress(KeyAscii As Integer)
    Dim bKeyCode As Boolean
    
    If t_nFormMode = EDIT_MODE Then
        If tgcDropdown.ComboOn(txtLevel) Then
            If KeyAscii = vbKeyReturn Then
                tgcDropdown.ComboSQL(txtLevel) = fnSetComboSQL(txtLevel.TabIndex)
                Screen.MousePointer = vbHourglass
            End If
        End If
        
        bKeyCode = tgcDropdown.Keypress(txtLevel, KeyAscii)
        Screen.MousePointer = vbDefault
        
        If Not bKeyCode Then
            If KeyAscii = vbKeyReturn Then
                If tgcDropdown.ComboOn(txtLevel) Then
                    If tgcDropdown.SingleRecordSelected Then
                          subEnterStageII
                    End If
                Else
                    subSetFocus txtPercent
                End If
                KeyAscii = 0
           End If
        Else
            cValidate.Keypress txtLevel, KeyAscii
        End If
    Else
        If KeyAscii = vbKeyReturn Then
            subEnterStageII
            subSetFocus txtPercent
            KeyAscii = 0
            If Not txtPercent.Enabled Then
                SendKeys "{TAB}", True
            End If
        Else
            tfnRegExpControlKeyPress txtLevel, KeyAscii, szIntegerPattern
            cValidate.Keypress txtLevel, KeyAscii
        End If
    End If

End Sub

Private Sub txtlevel_LostFocus()
    tgcDropdown.LostFocus txtLevel
    cValidate.LostFocus txtLevel, cmdLevel

    If cValidate.FirstInvalidInput < 0 Then
        subEnableUpdateBtn True
    End If
End Sub

Private Sub cmdlevel_Click()
    tgcDropdown.ComboSQL(txtLevel) = fnSetComboSQL(txtLevel.TabIndex)
    tgcDropdown.Click cmdLevel
End Sub

Private Sub subEnterStageII()

    If Not (cValidate.ValidInput(txtBonusCode) And cValidate.ValidInput(txtBonusCodeDesc) And cValidate.ValidInput(txtLevel)) Then
        Exit Sub
    End If

    If t_nFormMode = ADD_MODE Then
        subEnableControls True
        subEnableFirstLine False
        If txtBonusType.Caption <> "" Then
            subEnableVariables Trim(txtBonusType)
        End If
    Else
        If fnLoadBonusFormula(txtBonusCode, txtLevel) Then
            subEnableDeleteBtn True
        End If
    End If
    
    txtLevel.Enabled = True
    
    tgcDropdown.ComboOn(txtBonusCode) = False
    tgcDropdown.ComboOn(txtLevel) = False
    
    If t_nFormMode = ADD_MODE Then
        subSetFocus txtPercent
        If txtPercent.Enabled Then
            SelectIt txtPercent
        End If
    Else
        subSetFocus txtLevel
    End If
End Sub

Private Function fnValidLevel(Box As Textbox) As Boolean
    Const SUB_NAME As String = ""
    Dim strSQL As String
    Dim rsTemp As Recordset
    
    fnValidLevel = False
    
    If tfnRound(txtLevel) = 0 Then
        cValidate.SetErrorMessage Box, "Level is a required field"
        Exit Function
    End If
    
    If tfnRound(txtLevel) = tfnRound(txtLevel.Tag) Then
        fnValidLevel = True
        Exit Function
    End If
    
    If cValidate.ValidInput(txtBonusCode) Then
        strSQL = "SELECT * FROM bonus_formula WHERE bf_bonus_code = " & tfnSQLString(Trim(txtBonusCode))
        strSQL = strSQL & " AND bf_level = " & tfnRound(Box)
        
        If GetRecordCount(strSQL, , SUB_NAME) > 0 Then
            If t_nFormMode = ADD_MODE Or txtLevel.Tag <> "" Then
                cValidate.SetErrorMessage Box, "Comm. formula for this Level already exists"
                Exit Function
            End If
        End If
    End If
    
    fnValidLevel = True

End Function

Private Sub txtPercent_Change()
    subEnableUpdateBtn False
    cValidate.Change txtPercent
    
    If Not ActiveControl Is txtPercent Then Exit Sub
    
    t_bDataChanged = True
    subEnableRefreshBtn True
End Sub

Private Sub txtPercent_GotFocus()
    If txtAdjCond.Enabled Then
        Set cValidate.LastBox = txtAdjCond
    ElseIf txtLevel.Enabled Then
        Set cValidate.LastBox = txtLevel
    Else
        Set cValidate.LastBox = txtBonusCode
    End If
    
    SelectIt txtPercent
    cValidate.GotFocus txtPercent
End Sub

Private Sub txtPercent_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        subSetFocus txtDollar
        KeyAscii = 0
    Else
        tfnRegExpControlKeyPress txtPercent, KeyAscii, tfnDecimalPattern(8, 6, , True)
        cValidate.Keypress txtPercent, KeyAscii
    End If
End Sub

Private Sub txtPercent_LostFocus()
    cValidate.LostFocus txtPercent
    subFormatTo2Dec txtPercent
    
    If cValidate.FirstInvalidInput < 0 Then
        subEnableUpdateBtn True
    End If
End Sub

Private Sub txtDollar_Change()
    subEnableUpdateBtn False
    cValidate.Change txtDollar
    
    If Not ActiveControl Is txtDollar Then Exit Sub
    
    t_bDataChanged = True
    subEnableRefreshBtn True
End Sub

Private Sub txtDollar_GotFocus()
    SelectIt txtDollar
    cValidate.GotFocus txtDollar
End Sub

Private Sub txtDollar_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        subSetFocus txtAmount1
        KeyAscii = 0
    Else
        tfnRegExpControlKeyPress txtDollar, KeyAscii, tfnDecimalPattern(8, 2)
        cValidate.Keypress txtDollar, KeyAscii
    End If
End Sub

Private Sub txtDollar_LostFocus()
    cValidate.LostFocus txtDollar
    subFormatTo2Dec txtDollar
    
    If cValidate.FirstInvalidInput < 0 Then
        subEnableUpdateBtn True
    End If
End Sub

Private Sub txtAmount1_Change()
    subEnableUpdateBtn False
    cValidate.Change txtAmount1
    
    If Not ActiveControl Is txtAmount1 Then Exit Sub
    
    t_bDataChanged = True
    subEnableRefreshBtn True
End Sub

Private Sub txtAmount1_GotFocus()
    SelectIt txtAmount1
    cValidate.GotFocus txtAmount1
End Sub

Private Sub txtAmount1_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        subSetFocus txtAmount2
        KeyAscii = 0
    Else
        tfnRegExpControlKeyPress txtAmount1, KeyAscii, tfnDecimalPattern(10, 6, , True)
        cValidate.Keypress txtAmount1, KeyAscii
    End If
End Sub

Private Sub txtAmount1_LostFocus()
    cValidate.LostFocus txtAmount1
    subFormatTo2Dec txtAmount1
    
    If cValidate.FirstInvalidInput < 0 Then
        subEnableUpdateBtn True
    End If
End Sub

Private Sub txtAmount2_Change()
    subEnableUpdateBtn False
    cValidate.Change txtAmount2
    
    If Not ActiveControl Is txtAmount2 Then Exit Sub
    
    t_bDataChanged = True
    subEnableRefreshBtn True
End Sub

Private Sub txtAmount2_GotFocus()
    SelectIt txtAmount2
    cValidate.GotFocus txtAmount2
End Sub

Private Sub txtAmount2_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If txtMaxTotal.Enabled Then
            subSetFocus txtMaxTotal
        Else
            subSetFocus txtVariable1
        End If
        KeyAscii = 0
    Else
        tfnRegExpControlKeyPress txtAmount2, KeyAscii, tfnDecimalPattern(10, 6, , True)
        cValidate.Keypress txtAmount2, KeyAscii
    End If
End Sub

Private Sub txtAmount2_LostFocus()
    cValidate.LostFocus txtAmount2
    subFormatTo2Dec txtAmount2
    
    If cValidate.FirstInvalidInput < 0 Then
        subEnableUpdateBtn True
    End If
End Sub

Private Sub txtMaxTotal_Change()
    subEnableUpdateBtn False
    cValidate.Change txtMaxTotal
    
    If Not ActiveControl Is txtMaxTotal Then Exit Sub
    
    t_bDataChanged = True
    subEnableRefreshBtn True
End Sub

Private Sub txtMaxTotal_GotFocus()
    SelectIt txtMaxTotal
    cValidate.GotFocus txtMaxTotal
End Sub

Private Sub txtMaxTotal_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        subSetFocus txtVariable1
        KeyAscii = 0
    Else
        tfnRegExpControlKeyPress txtMaxTotal, KeyAscii, tfnDecimalPattern(8, 2)
        cValidate.Keypress txtMaxTotal, KeyAscii
    End If
End Sub

Private Sub txtMaxTotal_LostFocus()
    cValidate.LostFocus txtMaxTotal
    If cValidate.FirstInvalidInput < 0 Then
        subEnableUpdateBtn True
    End If
End Sub

Private Sub txtVariable1_Change()
    cValidate.Change txtVariable1
    tgcDropdown.Change txtVariable1
    
    If Not ActiveControl Is txtVariable1 Then Exit Sub
    
    subEnableUpdateBtn False
    t_bDataChanged = True
    subEnableRefreshBtn True
End Sub

Private Sub txtVariable1_GotFocus()
    tgcDropdown.GotFocus txtVariable1
    cValidate.GotFocus txtVariable1
    
    If tgcDropdown.SingleRecordSelected Then
        If txtVariable2.Enabled Then
            subSetFocus txtVariable2
        Else
            subSetFocus txtFormula
        End If
    End If
    
End Sub

Private Sub txtVariable1_KeyPress(KeyAscii As Integer)
    Dim bKeyCode As Boolean
    
    If KeyAscii = vbKeyReturn Then
        tgcDropdown.ComboSQL(txtVariable1) = fnSetComboSQL(txtVariable1.TabIndex)
        Screen.MousePointer = vbHourglass
    End If
        
    bKeyCode = tgcDropdown.Keypress(txtVariable1, KeyAscii)
    Screen.MousePointer = vbDefault
    
    If Not bKeyCode Then
       If KeyAscii = vbKeyReturn Then
          If tgcDropdown.SingleRecordSelected Then
                If txtVariable2.Enabled Then
                    subSetFocus txtVariable2
                Else
                    subSetFocus txtFormula
                End If
          End If
          KeyAscii = 0
       End If
    Else
        cValidate.Keypress txtVariable1, KeyAscii
    End If

End Sub

Private Sub txtVariable1_LostFocus()
    tgcDropdown.LostFocus txtVariable1
    cValidate.LostFocus txtVariable1, cmdVariable1
    If cValidate.FirstInvalidInput < 0 Then
        subEnableUpdateBtn True
    End If
End Sub

Private Sub txtVariable2_Change()
    cValidate.Change txtVariable2
    tgcDropdown.Change txtVariable2
    
    If Not ActiveControl Is txtVariable2 Then Exit Sub

    subEnableUpdateBtn False
    t_bDataChanged = True
    subEnableRefreshBtn True
End Sub

Private Sub txtVariable2_GotFocus()
    tgcDropdown.GotFocus txtVariable2
    cValidate.GotFocus txtVariable2
    
    If tgcDropdown.SingleRecordSelected Then
        If txtVariable3.Enabled Then
            subSetFocus txtVariable3
        Else
            subSetFocus txtFormula
        End If
    End If
    
End Sub

Private Sub txtVariable2_KeyPress(KeyAscii As Integer)
    Dim bKeyCode As Boolean
    
    If KeyAscii = vbKeyReturn Then
        tgcDropdown.ComboSQL(txtVariable2) = fnSetComboSQL(txtVariable2.TabIndex)
        Screen.MousePointer = vbHourglass
    End If
        
    bKeyCode = tgcDropdown.Keypress(txtVariable2, KeyAscii)
    Screen.MousePointer = vbDefault
    
    If Not bKeyCode Then
       If KeyAscii = vbKeyReturn Then
          If tgcDropdown.SingleRecordSelected Then
                If txtVariable3.Enabled Then
                    subSetFocus txtVariable3
                Else
                    subSetFocus txtFormula
                End If
          End If
          KeyAscii = 0
       End If
    Else
        cValidate.Keypress txtVariable2, KeyAscii
    End If

End Sub

Private Sub txtVariable2_LostFocus()
    tgcDropdown.LostFocus txtVariable2
    cValidate.LostFocus txtVariable2, cmdVariable2
    If cValidate.FirstInvalidInput < 0 Then
        subEnableUpdateBtn True
    End If
End Sub

Private Sub txtVariable3_Change()
    cValidate.Change txtVariable3
    tgcDropdown.Change txtVariable3
    
    If Not ActiveControl Is txtVariable3 Then Exit Sub
    
    subEnableUpdateBtn False
    t_bDataChanged = True
    subEnableRefreshBtn True
End Sub

Private Sub txtVariable3_GotFocus()
    tgcDropdown.GotFocus txtVariable3
    cValidate.GotFocus txtVariable3
    
    If tgcDropdown.SingleRecordSelected Then
        subSetFocus txtFormula
    End If
End Sub

Private Sub txtVariable3_KeyPress(KeyAscii As Integer)
    Dim bKeyCode As Boolean
    
    If KeyAscii = vbKeyReturn Then
        tgcDropdown.ComboSQL(txtVariable3) = fnSetComboSQL(txtVariable3.TabIndex)
        Screen.MousePointer = vbHourglass
    End If
        
    bKeyCode = tgcDropdown.Keypress(txtVariable3, KeyAscii)
    Screen.MousePointer = vbDefault
    
    If Not bKeyCode Then
       If KeyAscii = vbKeyReturn Then
          If tgcDropdown.SingleRecordSelected Then
                subSetFocus txtFormula
          End If
          KeyAscii = 0
       End If
    Else
        cValidate.Keypress txtVariable3, KeyAscii
    End If
End Sub

Private Sub txtVariable3_LostFocus()
    tgcDropdown.LostFocus txtVariable3
    cValidate.LostFocus txtVariable3, cmdVariable3
    If cValidate.FirstInvalidInput < 0 Then
        subEnableUpdateBtn True
    End If
End Sub

Private Sub txtFormula_Change()
    subEnableUpdateBtn False
    cValidate.Change txtFormula
    
    If Not ActiveControl Is txtFormula Then Exit Sub
    
    t_bDataChanged = True
    subEnableRefreshBtn True
End Sub

Private Sub txtFormula_GotFocus()
    If txtAdjCond.Enabled Then
        Set cValidate.LastBox = txtAdjCond
    ElseIf txtLevel.Enabled Then
        Set cValidate.LastBox = txtLevel
    Else
        Set cValidate.LastBox = txtBonusCode
    End If
    SelectIt txtFormula
    cValidate.GotFocus txtFormula
End Sub

Private Sub txtFormula_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        subSetFocus txtCondition
        KeyAscii = 0
    Else
        KeyAscii = Asc(LCase(Chr(KeyAscii)))
        tfnRegExpControlKeyPress txtFormula, KeyAscii, sRegExp80
        cValidate.Keypress txtFormula, KeyAscii
    End If
End Sub

Private Sub txtFormula_LostFocus()
    cValidate.LostFocus txtFormula
    If cValidate.FirstInvalidInput < 0 Then
        subEnableUpdateBtn True
    End If
End Sub

Private Sub txtCondition_Change()
    subEnableUpdateBtn False
    cValidate.Change txtCondition
    
    If Not ActiveControl Is txtCondition Then Exit Sub
    
    t_bDataChanged = True
    subEnableRefreshBtn True
End Sub

Private Sub txtCondition_GotFocus()
    SelectIt txtCondition
    cValidate.GotFocus txtCondition
End Sub

Private Sub txtCondition_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        subSetFocus txtAdjFormula
        KeyAscii = 0
    Else
        KeyAscii = Asc(LCase(Chr(KeyAscii)))
        tfnRegExpControlKeyPress txtCondition, KeyAscii, sRegExp80
        cValidate.Keypress txtCondition, KeyAscii
    End If
End Sub

Private Sub txtCondition_LostFocus()
    cValidate.LostFocus txtCondition
    If cValidate.FirstInvalidInput < 0 Then
        subEnableUpdateBtn True
    End If
End Sub

Private Sub txtAdjFormula_Change()
    subEnableUpdateBtn False
    cValidate.Change txtAdjFormula
    
    If Not ActiveControl Is txtAdjFormula Then Exit Sub
    
    t_bDataChanged = True
    subEnableRefreshBtn True
End Sub

Private Sub txtAdjFormula_GotFocus()
    SelectIt txtAdjFormula
    cValidate.GotFocus txtAdjFormula
End Sub

Private Sub txtAdjFormula_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        subSetFocus txtAdjCond
        KeyAscii = 0
    Else
        KeyAscii = Asc(LCase(Chr(KeyAscii)))
        tfnRegExpControlKeyPress txtAdjFormula, KeyAscii, sRegExp80
        cValidate.Keypress txtAdjFormula, KeyAscii
    End If
End Sub

Private Sub txtAdjFormula_LostFocus()
    cValidate.LostFocus txtAdjFormula
    If cValidate.FirstInvalidInput < 0 Then
        subEnableUpdateBtn True
    End If
End Sub

Private Sub txtAdjCond_Change()
    subEnableUpdateBtn False
    cValidate.Change txtAdjCond
    
    If Not ActiveControl Is txtAdjCond Then Exit Sub
    
    t_bDataChanged = True
    subEnableRefreshBtn True
End Sub

Private Sub txtAdjCond_GotFocus()
    SelectIt txtAdjCond
    cValidate.GotFocus txtAdjCond
End Sub

Private Sub txtAdjCond_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If cValidate.FirstInvalidInput < 0 Then
            subEnableUpdateBtn True
        End If
        
        If cmdUpdateInsertBtn.Enabled Then
            subSetFocus cmdUpdateInsertBtn
        Else
            'subSetFocus cmdExitCancelBtn
            SendKeys "{TAB}", True
        End If
        KeyAscii = 0
    Else
        KeyAscii = Asc(LCase(Chr(KeyAscii)))
        tfnRegExpControlKeyPress txtAdjCond, KeyAscii, sRegExp80
        cValidate.Keypress txtAdjCond, KeyAscii
    End If
End Sub

Private Sub txtAdjCond_LostFocus()
    cValidate.LostFocus txtAdjCond
    If cValidate.FirstInvalidInput < 0 Then
        subEnableUpdateBtn True
        If ActiveControl Is cmdRefresh Then
            subSetFocus cmdUpdateInsertBtn
        End If
    End If
End Sub

Private Sub efraBase_GotFocus()
    cValidate.GotFocus efraBase
End Sub

Private Function fnCreateTempTableVar() As Boolean
    Const SUB_NAME As String = "fnCreateTempTableVar"
    
    Dim strSQL As String
    Dim sarrVariable() As Variant
    Dim i As Integer
    
    'predefined variables
    sarrVariable = Array("inside_sales", "gallons_sold", "day_off_slip_day", "total_pay", _
        "months_in_grade", "months_as_manager", "years_as_manager", "months_employed", _
        "shortage_amount", "check_amount", "pay_hours")
    
    On Error GoTo Continue
    strSQL = "DROP TABLE tmpvariable"
    t_dbMainDatabase.ExecuteSQL strSQL
    
Continue:
    strSQL = "CREATE TEMP TABLE tmpVariable (Variable char(18))"
    If Not fnExecuteSQL(strSQL, , SUB_NAME) Then
        Exit Function
    End If
    
    For i = 0 To UBound(sarrVariable)
        strSQL = "INSERT INTO tmpvariable VALUES(" & tfnSQLString(sarrVariable(i)) & ")"
        If Not fnExecuteSQL(strSQL, , SUB_NAME) Then
            Exit Function
        End If
    Next
    
    fnCreateTempTableVar = True
End Function

Private Function fnValidVariables(Box As Textbox) As Boolean
    Const SUB_NAME As String = "fnValidVariables"
    Dim strSQL As String
    
    fnValidVariables = False
    
    If txtBonusType.Caption <> "" Then
        If Mid(txtBonusType, 2, 1) = "1" Then
            If Box.Name = "txtVariable2" Or Box.Name = "txtVariable3" Then
                fnValidVariables = True
                Exit Function
            End If
        ElseIf Mid(txtBonusType, 2, 1) = "2" Then
            If Box.Name = "txtVariable3" Then
                fnValidVariables = True
                Exit Function
            End If
        End If
    End If
    
    Box = Trim(Box)
    
    If Trim(Box) = "" Then
        fnValidVariables = True
        Exit Function
    End If
    
    strSQL = "SELECT * FROM tmpVariable WHERE variable = " & tfnSQLString(Trim(Box))
    
    If GetRecordCount(strSQL, , SUB_NAME) <> 1 Then
        cValidate.SetErrorMessage Box, "Variable does not exist, select one from the dropdown"
        Exit Function
    End If
    
    fnValidVariables = True

End Function

Private Function fnDeleteBonusFormula(sCode As String, nLevel As Integer) As Boolean
    Const SUB_NAME As String = "fnDeleteBonusFormula"
    Dim strSQL As String
    
    fnDeleteBonusFormula = False
    
    If Trim(sCode) = "" Or tfnRound(nLevel) = 0 Then
        Exit Function
    End If
    
    strSQL = "DELETE FROM bonus_formula WHERE bf_bonus_code = " & tfnSQLString(Trim(sCode))
    strSQL = strSQL & " AND bf_level = " & tfnRound(nLevel)
    
    If fnExecuteSQL(strSQL, , SUB_NAME) Then
        fnDeleteBonusFormula = True
    End If
    
End Function

Private Function fnValidFormula(Box As Textbox) As Boolean
    Const SUB_NAME As String = "fnValidFormula"
    
    Dim sErrMsg As String
    
    fnValidFormula = False
    
    If Box.Name = "txtAdjFormula" And Trim(Box.Text) = "" Then
        fnValidFormula = True
        Exit Function
    End If
    
    If Trim(Box) = "" Then
        cValidate.SetErrorMessage Box, "Commission formula cannot be blank"
        Exit Function
    End If
    
    sErrMsg = fnCheckFormula(Box, txtBonusType)
    
    If sErrMsg <> "" Then
        cValidate.SetErrorMessage Box, sErrMsg
        Exit Function
    End If
    
    fnValidFormula = True

End Function

Private Function fnValidCondition(Box As Textbox) As Boolean
    Const SUB_NAME As String = "fnValidCondition"
    
    Dim sErrMsg As String
    
    fnValidCondition = False
    
    If Box.Text = "" Then
        'cValidate.SetErrorMessage Box, "You must specify a condition"
        fnValidCondition = True
        Exit Function
    End If
    
    sErrMsg = fnCheckCondition(Box, txtBonusType)
    
    If sErrMsg <> "" Then
        cValidate.SetErrorMessage Box, sErrMsg
        Exit Function
    End If
    
    fnValidCondition = True

End Function

Private Sub subEnableVariables(sBonusType As String, Optional lRow As Long = -1)
    If sBonusType = "" Then
        Exit Sub
    End If
    
    On Error GoTo errTrap
    
    If Len(sBonusType) = 3 Then
        'number of variables allowed
        Select Case tfnRound(Mid(sBonusType, 2, 1))
            Case 1
                If lRow < 0 Then
                    txtVariable2 = "Not Used"
                    txtVariable2.Enabled = False
                    subEnableSearchbtn cmdVariable2, False
                    txtVariable3 = "Not Used"
                    txtVariable3.Enabled = False
                    subEnableSearchbtn cmdVariable3, False
                Else
                    tgmEditSelect.CellValue(nColVar2, lRow) = "Not Used"
                    tgmEditSelect.CellValue(nColVar3, lRow) = "Not Used"
                End If
            Case 2
                If lRow < 0 Then
                    txtVariable3 = "Not Used"
                    txtVariable3.Enabled = False
                    subEnableSearchbtn cmdVariable3, False
                Else
                    tgmEditSelect.CellValue(nColVar3, lRow) = "Not Used"
                End If
        End Select
        
        'max total allowed
        If UCase(Right(sBonusType, 1)) <> "E" Then
            If lRow < 0 Then
                txtMaxTotal = "Not Used"
                txtMaxTotal.Enabled = False
            Else
                tgmEditSelect.CellValue(nColMaxTotal, lRow) = "Not Used"
            End If
        End If
    End If

    Exit Sub
    
errTrap:
    tfnErrHandler "subEnableVariables"
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
    cmdDelete.Enabled = bOnOff
    mnuDelete.Enabled = bOnOff
End Sub

Private Sub subEnableRefreshBtn(bOnOff As Boolean)
    If bOnOff Then
        If t_nFormMode = ADD_MODE Then
            bOnOff = False
        End If
    End If
    
    cmdRefresh.Enabled = bOnOff
    mnuRefreshSelect.Enabled = bOnOff
End Sub

Private Sub subEnableUpdateBtn(bOnOff As Boolean)
    If bOnOff Then
        bOnOff = t_bDataChanged
    End If
    
    cmdUpdateInsertBtn.Enabled = bOnOff
    mnuUpdateInsert.Enabled = bOnOff
End Sub

Private Sub subSetGridWidth()
    Dim myWidth As Variant
    Dim myField As Variant
    
    Dim i As Integer
    
    myWidth = Array(600, 2500, 560, 2000, 3000, 630, 630, 850, 850, 850, 920, _
        920, 920, 920, 2000, 3000)
    myField = Array("bf_bonus_code", "bc_code_desc", "bf_level", "bf_formula", "bf_condition", _
        "bc_type", "bf_percent", "bf_dollar", "bf_amount1", "bf_amount2", "bf_variable1", _
         "bf_variable2", "bf_variable3", "bf_max_total", "bf_adj_formula", "bf_adj_condition")
    
    While tblEditSelect.Columns.Count > 0
        tblEditSelect.Columns.Remove 0
    Wend
    
    'tblEditSelect.ExtendRightColumn = True
    
    For i = 0 To UBound(myWidth)
        tblEditSelect.Columns.Add i
        With tblEditSelect.Columns(i)
            .Width = myWidth(i)
            .DataField = myField(i)
            .Visible = True
            .HeadAlignment = vbCenter
        End With
    Next
    
    tblEditSelect.Caption = "Commission Formula"
    tblEditSelect.Columns(nColCode).Caption = "Bonus Code"
    tblEditSelect.Columns(nColDesc).Caption = "Bonus Code Description"
    tblEditSelect.Columns(nColLevel).Caption = "Level"
    tblEditSelect.Columns(nColFormula).Caption = "Formula"
    tblEditSelect.Columns(nColCondition).Caption = "Condition"
    tblEditSelect.Columns(nColType).Caption = "Formula Type"
    tblEditSelect.Columns(nColPercent).Caption = "Percent (pct)"
    tblEditSelect.Columns(nColDollar).Caption = "Dollar (dol)"
    tblEditSelect.Columns(nColAmt1).Caption = "Amount1 (amt1)"
    tblEditSelect.Columns(nColAmt2).Caption = "Amount2 (amt2)"
    tblEditSelect.Columns(nColVar1).Caption = "Variable1 (v1)"
    tblEditSelect.Columns(nColVar2).Caption = "Variable2 (v2)"
    tblEditSelect.Columns(nColVar3).Caption = "Variable3 (v3)"
    tblEditSelect.Columns(nColMaxTotal).Caption = "MaxTotal (mxt)"
    tblEditSelect.Columns(nColAdjFormula).Caption = "Adj. Formula"
    tblEditSelect.Columns(nColAdjCondition).Caption = "Adj. Condition"
End Sub

Private Sub subInitSpreadsheets()
    
    subSetGridWidth
    
    Set tgmEditSelect = New clsTGSpreadSheet
    
    With tgmEditSelect
        Set .Table = tblEditSelect
        Set .StatusBar = ffraStatusbar ' message bar name
        Set .Form = Me
        Set .engFactor = t_engFactor
        .AllowAddNew = False
        .SetupTable True
        
        .SortByColumn = True
        
        .AddSortColumn nColCode, nColCode, .STRING_TYPE, .ASCENDING, .CASE_SENSITIVE, _
            nColLevel, .NUMERIC_TYPE, .ASCENDING, .CASE_SENSITIVE
            
        .AddSortColumn nColType, nColType, .STRING_TYPE, .ASCENDING, .CASE_SENSITIVE, _
            nColCode, .STRING_TYPE, .ASCENDING, .CASE_SENSITIVE, _
            nColLevel, .NUMERIC_TYPE, .ASCENDING, .CASE_SENSITIVE
        
    End With
    
    'Implement the selector class
    Set tgsEditSelect = New clsTGSelector
    tgsEditSelect.AvoidBeep = False
    Set tgsEditSelect.EditorClass = tgmEditSelect
    tgsEditSelect.SelectCurrRow = True
    tgsEditSelect.RowHighLighted = True
    tgsEditSelect.AllowMultipleSelect = False
End Sub

Private Function fnLoadEditSelectGrid()
    Const SUB_NAME As String = "fnLoadEditSelectGrid"
    Dim strSQL As String
    Dim i As Long
    
    strSQL = "SELECT * FROM bonus_formula, bonus_codes"
    strSQL = strSQL & " WHERE bf_bonus_code = bc_bonus_code"
    strSQL = strSQL & " ORDER BY bf_bonus_code, bf_level"
        
    tblEditSelect.Enabled = True
    tblEditSelect.col = 0
    tgmEditSelect.AllowAddNew = False
    tgmEditSelect.FillWithSQL t_dbMainDatabase, strSQL
    tgmEditSelect.AllowAddNew = False
    
    If tgmEditSelect.RowCount <= 0 Then
        MsgBox "No record available for Edit.", vbInformation
        Exit Function
    End If
    
    'fill NOT USED
    For i = 0 To tgmEditSelect.RowCount - 1
        subEnableVariables tgmEditSelect.CellValue(nColType, i), i
    Next i
    
    tgmEditSelect.Rebind
    
    tgsEditSelect.Click
    
    fnLoadEditSelectGrid = True
End Function

'*************************
'tblEditSelect grid Events
'*************************
'Private Sub tblEditSelect_AfterColEdit(ByVal ColIndex As Integer)
'    tgmEditSelect.AfterColEdit ColIndex
'End Sub
'
'Private Sub tblEditSelect_BeforeColEdit(ByVal ColIndex As Integer, ByVal KeyAscii As Integer, CANCEL As Integer)
'    tgmEditSelect.BeforeColEdit ColIndex, KeyAscii, CANCEL
'End Sub
'
Private Sub tblEditSelect_Change()
    tgmEditSelect.Change
End Sub

Private Sub tblEditSelect_Click()
    tgsEditSelect.Click
End Sub

Private Sub tblEditSelect_FirstRowChange()
    tgmEditSelect.FirstRowChange
End Sub

Private Sub tblEditSelect_GotFocus()
    tfnSetStatusBarMessage "Select a Commission Formula row"
    tgsEditSelect.GotFocus
    tgmEditSelect.GotFocus
End Sub

Private Sub tblEditSelect_KeyDown(KeyCode As Integer, Shift As Integer)
    #If SELECT_ON_ENTER Then
        If KeyCode = vbKeyReturn Then
            If cmdRefresh.Enabled Then
                KeyCode = 0
                cmdRefresh_Click
                Exit Sub
            End If
        End If
    #End If
    tgsEditSelect.KeyDown KeyCode, Shift
    tgmEditSelect.KeyDown KeyCode, Shift
End Sub

Private Sub tblEditSelect_KeyPress(KeyAscii As Integer)
    Dim nRow As Integer
    
    If KeyAscii <> vbKeyReturn Then
        Exit Sub
    End If
    
    If tgsEditSelect.Count > 1 Then
        MsgBox "Only one record can be edited at a time", vbInformation
        Exit Sub
    End If
    
    subSetFocus cmdRefresh
End Sub

Private Sub tblEditSelect_LostFocus()
    tgmEditSelect.LostFocus
End Sub

Private Sub tblEditSelect_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    tgsEditSelect.MouseUp Button, Shift, Y
End Sub

Private Sub tblEditSelect_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    tgmEditSelect.RowColChange LastRow, LastCol
    tgsEditSelect.RowColChange LastRow, LastCol
End Sub

Private Sub tblEditSelect_SelChange(CANCEL As Integer)
    tgsEditSelect.SelChange CANCEL
    CANCEL = True
End Sub

Private Sub tblEditSelect_UnboundReadData(ByVal RowBuf As DBTrueGrid.RowBuffer, StartLocation As Variant, ByVal ReadPriorRows As Boolean)
    tgmEditSelect.ReadData RowBuf, StartLocation, ReadPriorRows
End Sub

'return error message if any
Private Function fnCheckFormula(ByVal sFormula As String, ByVal sBonusType As String) As String
    Dim sErrMsg As String
    Dim objEvaluate As clsEquation
    Dim V1 As Double
    Dim V2 As Double
    Dim V3 As Double
    
    On Error GoTo errTrap
    
    sFormula = LCase(Trim(sFormula))
    
    'check formula using bonus type
    sErrMsg = fnCheckVarAllowed(sFormula, sBonusType)
    
    If sErrMsg <> "" Then
        fnCheckFormula = sErrMsg
        Exit Function
    End If
    
    'start formula evaluation
    Set objEvaluate = New clsEquation
    
    objEvaluate.Var("pct") = tfnRound(txtPercent, DEFAULT_DECIMALS)
    objEvaluate.Var("dol") = tfnRound(txtDollar, 2)
    objEvaluate.Var("amt1") = tfnRound(txtAmount1, DEFAULT_DECIMALS)
    objEvaluate.Var("amt2") = tfnRound(txtAmount2, DEFAULT_DECIMALS)
    objEvaluate.Var("mxt") = tfnRound(txtMaxTotal, DEFAULT_DECIMALS)
    If Trim(txtVariable1) <> "" Then
        V1 = fnGetVarValue("v1", txtVariable1, sErrMsg)
        If sErrMsg = "" Then
            objEvaluate.Var("v1") = V1
        End If
    End If
    If Trim(txtVariable2) <> "" Then
        V2 = fnGetVarValue("v2", txtVariable2, sErrMsg)
        If sErrMsg = "" Then
            objEvaluate.Var("v2") = V2
        End If
    End If
    If Trim(txtVariable3) <> "" Then
        V3 = fnGetVarValue("v3", txtVariable3, sErrMsg)
        If sErrMsg = "" Then
            objEvaluate.Var("v3") = V3
        End If
    End If
    
    objEvaluate.Equation = sFormula
    
    fnCheckFormula = tfnFixAmpersand(objEvaluate.Solve())
    
    Set objEvaluate = Nothing
    
    Exit Function
    
errTrap:
    tfnErrHandler "fnCheckFormula"
    fnCheckFormula = "Failed to validate Formula"

End Function

'return error message if any
Private Function fnCheckCondition(ByVal sCond As String, ByVal sBonusType As String) As String
    Dim sErrMsg As String
    Dim objCondition As clsCondition
    Dim V1 As Double
    Dim V2 As Double
    Dim V3 As Double
    
    On Error GoTo errTrap
    
    sCond = LCase(Trim(sCond))
    
    If Left(sCond, 2) = "if" Then
        sCond = Trim(Mid(sCond, 3))
    End If
    
    'check condition using bonus type
    sErrMsg = fnCheckVarAllowed(sCond, sBonusType)
    
    If sErrMsg <> "" Then
        fnCheckCondition = sErrMsg
        Exit Function
    End If
    
    'start formula evaluation
    Set objCondition = New clsCondition
    
    objCondition.Var("pct") = tfnRound(txtPercent, DEFAULT_DECIMALS)
    objCondition.Var("dol") = tfnRound(txtDollar, 2)
    objCondition.Var("amt1") = tfnRound(txtAmount1, DEFAULT_DECIMALS)
    objCondition.Var("amt2") = tfnRound(txtAmount2, DEFAULT_DECIMALS)
    objCondition.Var("mxt") = tfnRound(txtMaxTotal, DEFAULT_DECIMALS)
    If Trim(txtVariable1) <> "" Then
        V1 = fnGetVarValue("v1", txtVariable1, sErrMsg)
        If sErrMsg = "" Then
            objCondition.Var("v1") = V1
        End If
    End If
    If Trim(txtVariable2) <> "" Then
        V2 = fnGetVarValue("v2", txtVariable2, sErrMsg)
        If sErrMsg = "" Then
            objCondition.Var("v2") = V2
        End If
    End If
    If Trim(txtVariable3) <> "" Then
        V3 = fnGetVarValue("v3", txtVariable3, sErrMsg)
        If sErrMsg = "" Then
            objCondition.Var("v3") = V3
        End If
    End If
    
    objCondition.Equation = sCond
    
    fnCheckCondition = tfnFixAmpersand(objCondition.Solve())
    
    Set objCondition = Nothing
    
    Exit Function
    
errTrap:
    tfnErrHandler "fnCheckCondition"
    fnCheckCondition = "Failed to validate Condition"

End Function

Private Function fnCheckVarAllowed(sFormula As String, sBonusType As String) As String
    Dim sInvalidVar As String
    Dim aryInvalidVar() As String
    Dim i As Integer
    
    'check formula using bonus type
    'vaid bonus type format: T[123][ECX]
    If Len(sBonusType) = 3 Then
        Select Case tfnRound(Mid(sBonusType, 2, 1))
        Case 1
            sInvalidVar = sInvalidVar + "v2,v3"
        Case 2
            sInvalidVar = sInvalidVar + "v3"
        End Select
    
        If UCase(Right(sBonusType, 1)) <> "E" Then
            sInvalidVar = sInvalidVar + ",mxt"
        End If
    End If
    
    aryInvalidVar = Split(sInvalidVar, ",")
    
    For i = 0 To UBound(aryInvalidVar)
        If aryInvalidVar(i) <> "" Then
            If InStr(sFormula, aryInvalidVar(i)) > 0 Then
                fnCheckVarAllowed = tfnSQLString(aryInvalidVar(i)) + _
                    " is not valid for Bonus Type " + tfnSQLString(sBonusType)
                Exit Function
            End If
        End If
    Next i

    fnCheckVarAllowed = ""
End Function

Private Function fnGetVarValue(sV As String, _
                               ByVal sVariable As String, _
                               sErrMsg As String) As Double
                               
    Const SUB_NAME As String = "fnGetVarValue"
    
    Dim i As Long
    Dim sarrVariable()
    
    'predefined variables - SHOULD BE THE SAME AS THE DEFINITION IN ZZSEBFMT
    sarrVariable = Array("inside_sales", _
                         "gallons_sold", _
                         "day_off_slip_day", _
                         "total_pay", _
                         "months_in_grade", _
                         "months_as_manager", _
                         "years_as_manager", _
                         "months_employed", _
                         "shortage_amount", _
                         "check_amount", _
                         "pay_hours", _
                         "not used")
    
    fnGetVarValue = 0#
    sErrMsg = ""
    
    sVariable = LCase(Trim(sVariable))
    
    If sVariable = "" Then
        sErrMsg = "Variable is not defined in " + sV
        Exit Function
    End If
    
    Select Case sVariable
        Case sarrVariable(0)  'inside sales
            fnGetVarValue = 1.1
        Case sarrVariable(1)  'gallons sold
            fnGetVarValue = 2.2
        Case sarrVariable(2)  'day off slip days
            fnGetVarValue = 3.3
        
        Case sarrVariable(3)  'total pay
            fnGetVarValue = 4.4
            
        Case sarrVariable(4)  'months in grade
            fnGetVarValue = 5.5
        
        Case sarrVariable(5)  'months as manager
            fnGetVarValue = 6.6
        
        Case sarrVariable(6)  'years as manager
            fnGetVarValue = 7.7
        
        Case sarrVariable(7)  'months employed
            fnGetVarValue = 8.8
            
        Case sarrVariable(8)  'shortage amount
            fnGetVarValue = 9.9
            
        Case sarrVariable(9)  'check_amount
            fnGetVarValue = 10.12
            
        Case sarrVariable(10)  'pay hours
            fnGetVarValue = 11.21
            
        Case sarrVariable(11)  'not used
            Exit Function
        
        Case Else
            sErrMsg = "Variable '" + sVariable + "' in " + sV + " is not valid"
            Exit Function
    
    End Select
End Function

'will be call in textbox lostfocus event ONLY
Private Sub subFormatTo2Dec(txtBox As Textbox)
    If ActiveControl Is cmdAddBtn Then
        Exit Sub
    End If
    
    If txtBox = "" Then
        txtBox = "0.00"
    Else
        txtBox = Format(txtBox, "###,###,###,##0.00########")
    End If
End Sub
