VERSION 5.00
Object = "{C75015E0-2232-11D3-B440-0060971E99AF}#1.0#0"; "FACTFRM.OCX"
Object = "{3D388220-1F4E-11D3-B440-0060971E99AF}#1.0#0"; "FACTTAB.OCX"
Object = "{01028C21-0000-0000-0000-000000000046}#4.0#0"; "TG32OV.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{478E45E0-5745-11CF-8918-00A02416C765}#1.0#0"; "SQAOTE32.OCX"
Begin VB.Form frmZZSEIMPT 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Import Invoices"
   ClientHeight    =   6060
   ClientLeft      =   1050
   ClientTop       =   1950
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
   HelpContextID   =   1
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6060
   ScaleWidth      =   8880
   Begin FACTFRMLib.FactorFrame efraBackground 
      Height          =   5148
      Left            =   0
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   552
      Width           =   8880
      _Version        =   65536
      _ExtentX        =   15663
      _ExtentY        =   9080
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
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin FACTTABLib.FactorTab etabMain 
         Height          =   4644
         Left            =   0
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   0
         Width           =   8880
         _Version        =   65536
         _ExtentX        =   15663
         _ExtentY        =   8191
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
         TabsPerPage     =   2
         Caption         =   "I&mport Invoices|Impor&t Invoice Detail"
         Begin FACTFRMLib.FactorFrame efraTab2Back 
            Height          =   4356
            Left            =   14460
            TabIndex        =   10
            TabStop         =   0   'False
            Top             =   0
            Width           =   8880
            _Version        =   65536
            _ExtentX        =   15663
            _ExtentY        =   7683
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
            BeginProperty PanelFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Begin FACTFRMLib.FactorFrame eLblStatus 
               Height          =   372
               Left            =   60
               TabIndex        =   56
               TabStop         =   0   'False
               Top             =   3888
               Width           =   2772
               _Version        =   65536
               _ExtentX        =   4890
               _ExtentY        =   656
               _StockProps     =   77
               ForeColor       =   32768
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
               Caption         =   "Out of Balance By 999999.999"
               BorderWidth     =   0
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
            Begin FACTFRMLib.FactorFrame cmdDetailCancel 
               Height          =   396
               HelpContextID   =   15
               Left            =   7488
               TabIndex        =   51
               Top             =   3864
               Width           =   1308
               _Version        =   65536
               _ExtentX        =   2307
               _ExtentY        =   698
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
               Caption         =   "Cance&l"
               CaptionPos      =   4
               ShowFocusRect   =   -1  'True
               Style           =   3
               BorderWidth     =   4
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
            Begin FACTFRMLib.FactorFrame cmdOk 
               Height          =   396
               HelpContextID   =   16
               Left            =   6000
               TabIndex        =   50
               Top             =   3864
               Width           =   1308
               _Version        =   65536
               _ExtentX        =   2307
               _ExtentY        =   698
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
               Caption         =   "O&K"
               CaptionPos      =   4
               ShowFocusRect   =   -1  'True
               Style           =   3
               BorderWidth     =   4
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
            Begin FACTFRMLib.FactorFrame cmdDetailDelete 
               Height          =   396
               HelpContextID   =   12
               Left            =   4524
               TabIndex        =   52
               Top             =   3864
               Width           =   1308
               _Version        =   65536
               _ExtentX        =   2307
               _ExtentY        =   698
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
               ShowFocusRect   =   -1  'True
               Style           =   3
               BorderWidth     =   4
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
            Begin FACTFRMLib.FactorFrame cmdExceptions 
               Height          =   396
               HelpContextID   =   1311
               Left            =   3036
               TabIndex        =   49
               Top             =   3864
               Width           =   1308
               _Version        =   65536
               _ExtentX        =   2307
               _ExtentY        =   698
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
               Caption         =   "Exceptio&ns"
               CaptionPos      =   4
               ShowFocusRect   =   -1  'True
               Style           =   3
               BorderWidth     =   4
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
            Begin FACTFRMLib.FactorFrame efraTab2 
               Height          =   3756
               Left            =   48
               TabIndex        =   12
               TabStop         =   0   'False
               Top             =   36
               Width           =   8748
               _Version        =   65536
               _ExtentX        =   15430
               _ExtentY        =   6625
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
               BeginProperty PanelFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Begin FACTFRMLib.FactorFrame efraDetail 
                  Height          =   2724
                  Left            =   72
                  TabIndex        =   47
                  TabStop         =   0   'False
                  Top             =   984
                  Width           =   8604
                  _Version        =   65536
                  _ExtentX        =   15176
                  _ExtentY        =   4805
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
                  BeginProperty PanelFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Begin VB.PictureBox picFullDetailText 
                     Appearance      =   0  'Flat
                     BackColor       =   &H80000005&
                     ForeColor       =   &H80000008&
                     Height          =   444
                     Left            =   336
                     ScaleHeight     =   420
                     ScaleWidth      =   1440
                     TabIndex        =   58
                     Top             =   1644
                     Visible         =   0   'False
                     Width           =   1476
                  End
                  Begin DBTrueGrid.TDBGrid tblDetail 
                     Height          =   2616
                     HelpContextID   =   1310
                     Left            =   48
                     OleObjectBlob   =   "ZZSEIMPT.frx":0000
                     TabIndex        =   48
                     Top             =   48
                     Width           =   8520
                  End
                  Begin FACTFRMLib.FactorFrame cmdSearch 
                     Height          =   336
                     Left            =   396
                     TabIndex        =   53
                     Top             =   252
                     Width           =   1140
                     _Version        =   65536
                     _ExtentX        =   2011
                     _ExtentY        =   593
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
                     Style           =   3
                     BorderWidth     =   4
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
                  Begin VB.Data datDropDown 
                     Caption         =   "DropDown"
                     Connect         =   "Access"
                     DatabaseName    =   ""
                     DefaultCursorType=   0  'DefaultCursor
                     DefaultType     =   2  'UseODBC
                     Exclusive       =   0   'False
                     Height          =   420
                     Left            =   276
                     Options         =   0
                     ReadOnly        =   0   'False
                     RecordsetType   =   1  'Dynaset
                     RecordSource    =   ""
                     Top             =   2052
                     Visible         =   0   'False
                     Width           =   1428
                  End
               End
               Begin FACTFRMLib.FactorFrame efraDetailHeader 
                  Height          =   840
                  Left            =   72
                  TabIndex        =   34
                  TabStop         =   0   'False
                  Top             =   132
                  Width           =   8604
                  _Version        =   65536
                  _ExtentX        =   15176
                  _ExtentY        =   1482
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
                  BeginProperty PanelFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Begin VB.Label Label2 
                     Caption         =   "Amount"
                     Height          =   216
                     Left            =   6684
                     TabIndex        =   55
                     Top             =   468
                     Width           =   696
                  End
                  Begin VB.Label lblAmount 
                     Caption         =   "00000.00"
                     BeginProperty Font 
                        Name            =   "Arial"
                        Size            =   9.75
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   216
                     Left            =   7440
                     TabIndex        =   46
                     Top             =   468
                     Width           =   1080
                  End
                  Begin VB.Label lblInvoice 
                     Caption         =   "000000"
                     BeginProperty Font 
                        Name            =   "Arial"
                        Size            =   9.75
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   216
                     Left            =   5688
                     TabIndex        =   45
                     Top             =   468
                     Width           =   984
                  End
                  Begin VB.Label Label12 
                     Caption         =   "Invoice #"
                     Height          =   216
                     Left            =   4872
                     TabIndex        =   44
                     Top             =   468
                     Width           =   756
                  End
                  Begin VB.Label lblShift 
                     Caption         =   "00000"
                     BeginProperty Font 
                        Name            =   "Arial"
                        Size            =   9.75
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   240
                     Left            =   7440
                     TabIndex        =   43
                     Top             =   120
                     Width           =   732
                  End
                  Begin VB.Label Label10 
                     Caption         =   "Shift"
                     Height          =   240
                     Left            =   6684
                     TabIndex        =   42
                     Top             =   120
                     Width           =   468
                  End
                  Begin VB.Label lblVendorName 
                     Caption         =   "WWWWWWWWWW"
                     BeginProperty Font 
                        Name            =   "Arial"
                        Size            =   9.75
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   216
                     Left            =   2220
                     TabIndex        =   41
                     Top             =   468
                     Width           =   2772
                  End
                  Begin VB.Label lblVendor 
                     Caption         =   "0000000"
                     BeginProperty Font 
                        Name            =   "Arial"
                        Size            =   9.75
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   216
                     Left            =   1152
                     TabIndex        =   40
                     Top             =   468
                     Width           =   912
                  End
                  Begin VB.Label Label6 
                     Caption         =   "Vendor No."
                     Height          =   216
                     Left            =   72
                     TabIndex        =   39
                     Top             =   468
                     Width           =   1056
                  End
                  Begin VB.Label lblReportDate 
                     Caption         =   "01/01/00"
                     BeginProperty Font 
                        Name            =   "Arial"
                        Size            =   9.75
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   240
                     Left            =   4884
                     TabIndex        =   38
                     Top             =   120
                     Width           =   1056
                  End
                  Begin VB.Label lblPrftCtrName 
                     Caption         =   "WWWWWWWW"
                     BeginProperty Font 
                        Name            =   "Arial"
                        Size            =   9.75
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   240
                     Left            =   2220
                     TabIndex        =   37
                     Top             =   120
                     Width           =   2700
                  End
                  Begin VB.Label lblPrftCtr 
                     Caption         =   "00000"
                     BeginProperty Font 
                        Name            =   "Arial"
                        Size            =   9.75
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   240
                     Left            =   1320
                     TabIndex        =   36
                     Top             =   120
                     Width           =   696
                  End
                  Begin VB.Label Label1 
                     Caption         =   "Profit Center:"
                     Height          =   240
                     Left            =   72
                     TabIndex        =   35
                     Top             =   120
                     Width           =   1176
                  End
               End
            End
         End
         Begin FACTFRMLib.FactorFrame efraTabBack1 
            Height          =   4332
            Left            =   12
            TabIndex        =   9
            TabStop         =   0   'False
            Top             =   12
            Width           =   8856
            _Version        =   65536
            _ExtentX        =   15621
            _ExtentY        =   7641
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
            BeginProperty PanelFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Begin FACTFRMLib.FactorFrame efraTab1 
               Height          =   3756
               Left            =   60
               TabIndex        =   11
               TabStop         =   0   'False
               Top             =   36
               Width           =   8748
               _Version        =   65536
               _ExtentX        =   15430
               _ExtentY        =   6625
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
               BeginProperty PanelFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Begin FACTFRMLib.FactorFrame efraHeader 
                  Height          =   1848
                  Left            =   108
                  TabIndex        =   29
                  TabStop         =   0   'False
                  Top             =   1872
                  Width           =   8544
                  _Version        =   65536
                  _ExtentX        =   15071
                  _ExtentY        =   3260
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
                  BeginProperty PanelFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Begin VB.PictureBox picFullHeaderText 
                     Appearance      =   0  'Flat
                     BackColor       =   &H80000005&
                     ForeColor       =   &H80000008&
                     Height          =   396
                     Left            =   840
                     ScaleHeight     =   360
                     ScaleWidth      =   1275
                     TabIndex        =   57
                     Top             =   1068
                     Visible         =   0   'False
                     Width           =   1308
                  End
                  Begin DBTrueGrid.TDBGrid tblHeader 
                     Height          =   1752
                     HelpContextID   =   1309
                     Left            =   0
                     OleObjectBlob   =   "ZZSEIMPT.frx":12DD
                     TabIndex        =   30
                     Top             =   36
                     Width           =   8520
                  End
               End
               Begin FACTFRMLib.FactorFrame efraProfitCtr 
                  Height          =   1800
                  Left            =   108
                  TabIndex        =   13
                  TabStop         =   0   'False
                  Top             =   72
                  Width           =   8544
                  _Version        =   65536
                  _ExtentX        =   15071
                  _ExtentY        =   3175
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
                  BeginProperty PanelFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Begin VB.CheckBox chkAllInvoice 
                     Caption         =   "Include Processed Invoices"
                     Height          =   300
                     HelpContextID   =   1308
                     Left            =   336
                     TabIndex        =   28
                     Top             =   1404
                     Width           =   6888
                  End
                  Begin FACTFRMLib.FactorFrame cmdToPrftName 
                     Height          =   360
                     HelpContextID   =   1307
                     Left            =   7668
                     TabIndex        =   27
                     TabStop         =   0   'False
                     Top             =   948
                     Width           =   360
                     _Version        =   65536
                     _ExtentX        =   635
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
                     Style           =   3
                     BorderWidth     =   4
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
                  Begin FACTFRMLib.FactorFrame cmdFromPrftName 
                     Height          =   360
                     HelpContextID   =   1303
                     Left            =   7668
                     TabIndex        =   26
                     TabStop         =   0   'False
                     Top             =   300
                     Width           =   360
                     _Version        =   65536
                     _ExtentX        =   635
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
                     Style           =   3
                     BorderWidth     =   4
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
                  Begin FACTFRMLib.FactorFrame cmdToPrftCtr 
                     Height          =   360
                     HelpContextID   =   1305
                     Left            =   3288
                     TabIndex        =   25
                     TabStop         =   0   'False
                     Top             =   948
                     Width           =   360
                     _Version        =   65536
                     _ExtentX        =   635
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
                     CaptionPos      =   4
                     Style           =   3
                     BorderWidth     =   4
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
                  Begin FACTFRMLib.FactorFrame cmdFromPrftCtr 
                     Height          =   360
                     HelpContextID   =   1301
                     Left            =   3288
                     TabIndex        =   24
                     TabStop         =   0   'False
                     Top             =   300
                     Width           =   360
                     _Version        =   65536
                     _ExtentX        =   635
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
                     Style           =   3
                     BorderWidth     =   4
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
                  Begin VB.TextBox txtToPrftName 
                     Height          =   360
                     HelpContextID   =   1306
                     Left            =   3693
                     TabIndex        =   23
                     Top             =   948
                     Width           =   3960
                  End
                  Begin VB.TextBox txtToPrftCtr 
                     Height          =   360
                     HelpContextID   =   1304
                     Left            =   1212
                     TabIndex        =   22
                     Top             =   948
                     Width           =   2062
                  End
                  Begin VB.TextBox txtFromPrftName 
                     Height          =   360
                     HelpContextID   =   1302
                     Left            =   3693
                     TabIndex        =   21
                     Top             =   300
                     Width           =   3960
                  End
                  Begin VB.TextBox txtFromPrftCtr 
                     Height          =   360
                     HelpContextID   =   1300
                     Left            =   1212
                     TabIndex        =   20
                     Top             =   300
                     Width           =   2062
                  End
                  Begin VB.Label lblToPrftName 
                     Caption         =   "Profit Center Name"
                     Height          =   336
                     Left            =   3696
                     TabIndex        =   19
                     Top             =   696
                     Width           =   1860
                  End
                  Begin VB.Label lblToProfitCtr 
                     Caption         =   "Profit Center No."
                     Height          =   336
                     Left            =   1212
                     TabIndex        =   18
                     Top             =   696
                     Width           =   1860
                  End
                  Begin VB.Label lblFromPrftName 
                     Caption         =   "Profit Center Name"
                     Height          =   336
                     Left            =   3696
                     TabIndex        =   17
                     Top             =   24
                     Width           =   1860
                  End
                  Begin VB.Label lblFromPrftCtr 
                     Caption         =   "Profit Center No."
                     Height          =   336
                     Left            =   1212
                     TabIndex        =   16
                     Top             =   24
                     Width           =   1860
                  End
                  Begin VB.Label lblTo 
                     Caption         =   "To:"
                     Height          =   300
                     Left            =   336
                     TabIndex        =   15
                     Top             =   1008
                     Width           =   840
                  End
                  Begin VB.Label lblFrom 
                     Caption         =   "From:"
                     Height          =   300
                     Left            =   336
                     TabIndex        =   14
                     Top             =   348
                     Width           =   1056
                  End
               End
            End
            Begin FACTFRMLib.FactorFrame cmdQuery 
               Height          =   396
               HelpContextID   =   18
               Left            =   4584
               TabIndex        =   31
               Top             =   3864
               Width           =   1308
               _Version        =   65536
               _ExtentX        =   2307
               _ExtentY        =   698
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
               Caption         =   "&Query"
               CaptionPos      =   4
               ShowFocusRect   =   -1  'True
               Style           =   3
               BorderWidth     =   4
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
            Begin FACTFRMLib.FactorFrame cmdHeaderSelect 
               Height          =   396
               HelpContextID   =   17
               Left            =   6048
               TabIndex        =   32
               Top             =   3864
               Width           =   1308
               _Version        =   65536
               _ExtentX        =   2307
               _ExtentY        =   698
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
               Caption         =   "&Select"
               CaptionPos      =   4
               ShowFocusRect   =   -1  'True
               Style           =   3
               BorderWidth     =   4
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
            Begin FACTFRMLib.FactorFrame cmdHeaderDelete 
               Height          =   396
               HelpContextID   =   12
               Left            =   7512
               TabIndex        =   33
               Top             =   3864
               Width           =   1308
               _Version        =   65536
               _ExtentX        =   2307
               _ExtentY        =   698
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
               ShowFocusRect   =   -1  'True
               Style           =   3
               BorderWidth     =   4
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
         End
      End
      Begin FACTFRMLib.FactorFrame cmdUpdateInsertBtn 
         Height          =   396
         HelpContextID   =   13
         Left            =   5976
         TabIndex        =   1
         Top             =   4704
         Width           =   1308
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
         Caption         =   "&Insert"
         CaptionPos      =   4
         PicturePos      =   3
         ShowFocusRect   =   -1  'True
         Style           =   3
         BorderWidth     =   4
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
      Begin FACTFRMLib.FactorFrame cmdExitCancelBtn 
         Height          =   396
         HelpContextID   =   15
         Left            =   7440
         TabIndex        =   2
         Top             =   4704
         Width           =   1308
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
         Caption         =   "E&xit"
         CaptionPos      =   4
         PicturePos      =   3
         ShowFocusRect   =   -1  'True
         Style           =   3
         BorderWidth     =   4
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
   End
   Begin FACTFRMLib.FactorFrame ffraStatusbar 
      Height          =   360
      Left            =   0
      TabIndex        =   3
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
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   0
      Width           =   8880
      _Version        =   65536
      _ExtentX        =   15663
      _ExtentY        =   825
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
      Align           =   1
      FMName          =   "ZZSEIMPT"
      CaptionPos      =   4
      Style           =   6
      BeginProperty PanelFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
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
         Left            =   0
         TabIndex        =   0
         Top             =   36
         Width           =   3396
         _ExtentX        =   5980
         _ExtentY        =   556
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
   Begin DBTrueGrid.TDBGrid tblComboDropdown 
      Bindings        =   "ZZSEIMPT.frx":25BA
      Height          =   2484
      Left            =   72
      OleObjectBlob   =   "ZZSEIMPT.frx":25D9
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   576
      Width           =   3756
   End
   Begin SQAOTestObjectsCtl.SQAOTest SQAOTest1 
      Height          =   570
      Left            =   10890
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   90
      Width           =   570
      _ExtentX        =   1005
      _ExtentY        =   1005
   End
   Begin DBTrueGrid.TDBGrid tblDropDown 
      Bindings        =   "ZZSEIMPT.frx":38BD
      Height          =   648
      Left            =   0
      OleObjectBlob   =   "ZZSEIMPT.frx":38D7
      TabIndex        =   54
      TabStop         =   0   'False
      Top             =   600
      Width           =   1752
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
Attribute VB_Name = "frmZZSEIMPT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***********************************************************'
'
' Copyright (c) 1996 FACTOR, A Division of W.R.Hess Company
'
' Module name   : ZZEININV.FRM
' Date          : February 1, 2001
' Programmer(s) : Junsong Qiu
' This module is used to process the data in the holding invoice information table.


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
Private cValidate As cValidateInput
Private tgmHeader As clsTGSpreadSheet
Private tgmDetail As clsTGSpreadSheet
Private tgfDropdown As clsFloatingDropDown
Private objFullHeaderText As clsColumnExtension
Private objFullDetailText As clsColumnExtension

Private g_bDetailDataChanged As Boolean
Private g_bExceptionsClicked As Boolean
Private g_vHoldDetailData() As Variant
Private g_nHoldSelectedPrftCtr As Integer
Private g_lHoldSelectedVendor As Long
Private g_lHoldSelectedInvoice As Long
Private g_dtHoldRptDate As Date
Private g_lHoldShiftNum As Long
Private g_lHoldShiftLink As Long
Private g_lNewLineNo As Long
Private g_lSelectedHeaderRow As Long

Private Sub subSetupCombos()
    Dim sSql As String

    'Setup the combos
    ' Set the general properties of the combos
    Set tgcDropdown = CreateObject(t_szOLECOMBO)
    
    If Not tgcDropdown Is Nothing Then
        With tgcDropdown
            Set .DBEngine = t_engFactor
            Set .Form = Me
            Set .DataBase = t_dbMainDatabase
            Set .DataLink = datComboDropDown
            Set .Table = tblComboDropdown
            
            'The profit center combo
            .AddCombo
            .AddComboBox txtFromPrftCtr, cmdFromPrftCtr, "prft_ctr", .SQL_INT_TYPE
            .AddComboBox txtFromPrftName, cmdFromPrftName, "prft_name", .SQL_STRING_TYPE(40)
            
            .AddCombo
            .AddComboBox txtToPrftCtr, cmdToPrftCtr, "prft_ctr", .SQL_INT_TYPE
            .AddComboBox txtToPrftName, cmdToPrftName, "prft_name", .SQL_STRING_TYPE(40)
        End With
        
    End If
    
End Sub


Private Sub subInitValidation()
    Set cValidate = New cValidateInput
    Set cValidate.Form = Me
    Set cValidate.StatusBar = ffraStatusbar
    
    With cValidate
        .AddEditBox txtFromPrftCtr, "Enter From Profit center"
        .AddEditBox txtToPrftCtr, "Enter To Profit Center"
        .ESCControl = cmdExitCancelBtn
        .MinTabIndex = tbToolbar.TabIndex
        .MaxTabIndex = chkAllInvoice.TabIndex
        Set .LastBox = txtToPrftCtr
    End With

End Sub

Public Function fnInvalidData(txtBox As Textbox) As Boolean
    fnInvalidData = Not fnValidProfitCtr(txtBox)
End Function

Public Function fnValidCellValue(tgTable As TDBGrid, _
                                 ByVal nCurrCol As Integer, _
                                 ByVal lCurrRow As Long, _
                                 szText As String) As Boolean
    fnValidCellValue = True
    
    If tgTable Is tblHeader Then
        fnValidCellValue = True
        Exit Function
    End If
    
    Select Case nCurrCol
        Case ColDetailItemCode
            fnValidCellValue = fnValidItemCode(szText, nCurrCol, lCurrRow)
        Case ColDetailUOM
            fnValidCellValue = fnValidUOM(szText, nCurrCol, lCurrRow)
        Case ColDetailQty
            fnValidCellValue = fnValidDetailQty(szText, nCurrCol, lCurrRow)
        Case ColDetailCost
            fnValidCellValue = fnValidDetailCost(szText, nCurrCol, lCurrRow)
    End Select
    

    
End Function

Private Function fnValidItemCode(szText As String, nCurrCol As Integer, lCurrDetailRow As Long) As Boolean
    Dim strSQL As String
    Dim rsTemp As Recordset
    Dim lVendor As Long
    
    lVendor = tfnRound(tgmHeader.CellValue(ColHeaderVendor, tgmHeader.GetCurrentRowNumber))
    
    If Trim(szText) = "" Then
        tgmDetail.ErrorMessage(nCurrCol) = "The item code can't be empty"
        Exit Function
    End If
    
    strSQL = "SELECT rsbi_code, rsbi_desc, rsbi_stock_unit FROM rs_b_item WHERE rsbi_vendor = " & lVendor
    strSQL = strSQL & " AND rsbi_code = " & tfnSQLString(szText)
    
    If fnGetRecord(rsTemp, strSQL, nDB_REMOTE, "fnValidItemCode") < 0 Then
        tgmDetail.ErrorMessage(nCurrCol) = "Database Access Error"
    ElseIf rsTemp.RecordCount = 0 Then
        tgmDetail.ErrorMessage(nCurrCol) = "Invalid Item Code entered"
    Else
        fnValidItemCode = True
        subLoadInfoForThisRow lVendor, Trim(szText), Trim(rsTemp!rsbi_stock_unit & ""), Trim(rsTemp!rsbi_desc & "")
        'tgmDetail.CellValue(ColDetailItemDesc, lCurrDetailRow) = Trim(rsTemp!rsbi_desc & "")
    End If
    
End Function

Private Function fnValidUOM(szText As String, nCurrCol As Integer, lCurrDetailRow As Long) As Boolean
    Dim strSQL As String
    Dim rsTemp As Recordset
    Dim lVendor As Long
    Dim lInvoice As Long
    Dim nPrftctr As Integer
    Dim dtInvDate As Date
    Dim lCurHeaderRow As Long
    Dim dCost As Double
    Dim dRetailPrice As Double
    Dim sItemCode As String

    lCurHeaderRow = tgmHeader.GetCurrentRowNumber
    lVendor = tfnRound(tgmHeader.CellValue(ColHeaderVendor, lCurHeaderRow))
    lInvoice = tfnRound(tgmHeader.CellValue(ColHeaderInvoice, lCurHeaderRow))
    nPrftctr = tfnRound(tgmHeader.CellValue(ColHeaderPrftctr, lCurHeaderRow))
    dtInvDate = tfnFormatDate(tgmHeader.CellValue(colHdnHeaderInvDate, lCurHeaderRow))
    
    If Trim(szText) = "" Then
        tgmDetail.ErrorMessage(ColDetailUOM) = " The Unit of Measure can't be empty."
        Exit Function
    End If
    
    If tgmDetail.ValidCell(ColDetailItemCode, lCurrDetailRow) Then
        sItemCode = tgmDetail.CellValue(ColDetailItemCode, lCurrDetailRow)
        
        strSQL = "SELECT rsbi_stock_unit FROM rs_b_item WHERE "
        strSQL = strSQL & " rsbi_vendor = " & lVendor
        strSQL = strSQL & " AND rsbi_code = " & tfnSQLString(sItemCode)
        
        If fnGetRecord(rsTemp, strSQL, nDB_REMOTE, "fnValidItemCode") < 0 Then
            tgmDetail.ErrorMessage(nCurrCol) = "Database Access Error"
        ElseIf rsTemp.RecordCount = 0 Then
            tgmDetail.ErrorMessage(nCurrCol) = "Invalid Item entered"
        Else
        
            If UCase(Trim(rsTemp!rsbi_stock_unit & "")) = UCase(Trim(szText)) Then
                
                If Not fnGetPriceBookCost(sItemCode, lVendor, nPrftctr, dtInvDate, dCost, dRetailPrice) Then
                    fnValidUOM = False
                    tgmDetail.ErrorMessage(nCurrCol) = "No Price book available for this item"
                    Exit Function
                End If
                
            Else
            
                If Not fnGetPriceFromItemCostMaint(lVendor, sItemCode, szText, dtInvDate, dCost, dRetailPrice) Then
                    fnValidUOM = False
                    tgmDetail.ErrorMessage(nCurrCol) = "No Price for this Unit of Measure."
                    Exit Function
                End If
                
            End If
            
            tgmDetail.CellValue(ColDetailCost, lCurrDetailRow) = tfnRound(dCost, 3)
            tgmDetail.CellValue(ColDetailPBCost, lCurrDetailRow) = tfnRound(dCost, 3)
            tgmDetail.CellValue(ColDetailRetail, lCurrDetailRow) = tfnRound(dRetailPrice, 3)
            tgmDetail.CellValue(colHdnDetailFlag, lCurrDetailRow) = "Y"
            fnValidUOM = True
            
        End If
        
    Else
        fnValidUOM = True
    End If
    
End Function


Private Function fnValidDetailQty(szText As String, nCurrCol As Integer, lCurrDetailRow As Long) As Boolean
    
    If Trim(szText) = "" Then
        tgmDetail.ErrorMessage(nCurrCol) = "The Quantity can't be empty"
        Exit Function
    End If
    
    fnValidDetailQty = True
End Function

Private Function fnValidDetailCost(szText As String, nCurrCol As Integer, lCurrDetailRow As Long) As Boolean
    
    If Trim(szText) = "" Then
        tgmDetail.ErrorMessage(nCurrCol) = "The cost can't be empty"
        Exit Function
    End If
    
    If Not IsNumeric(szText) Then
        tgmDetail.ErrorMessage(nCurrCol) = "Invalid Cost entered"
    ElseIf CDbl(szText) < 0 Then
        tgmDetail.ErrorMessage(nCurrCol) = "The cost can't be less than 0"
    Else
    
        If CDbl(szText) <> tgmDetail.CellValue(ColDetailPBCost, lCurrDetailRow) Then
            tgmDetail.CellValue(colHdnDetailFlag, lCurrDetailRow) = "N"
        Else
            tgmDetail.CellValue(colHdnDetailFlag, lCurrDetailRow) = "Y"
        End If
        
        fnValidDetailCost = True
    End If
    
End Function

Private Function fnValidProfitCtr(txtBox As Textbox) As Boolean
    Dim strSQL As String
    Dim rsTemp As Recordset
    
    If Trim(txtBox.Text) = "" Then
        fnValidProfitCtr = True
        Exit Function
    ElseIf Not IsNumeric(txtBox.Text) Then
        fnValidProfitCtr = False
        cValidate.SetErrorMessage txtBox, "Invalid Profit center entered"
    Else
        strSQL = "SELECT prft_ctr, prft_name FROM sys_prft_ctr WHERE prft_ctr = " & txtBox.Text
        
        If fnGetRecord(rsTemp, strSQL, nDB_REMOTE, "fnValidProfitCtr") < 0 Then
            fnValidProfitCtr = False
            cValidate.SetErrorMessage txtBox, "Database access error"
        ElseIf rsTemp.RecordCount = 0 Then
            fnValidProfitCtr = False
            cValidate.SetErrorMessage txtBox, "The profit center doesn't exist."
            tgcDropdown.SetText txtBox, ""
        Else
            fnValidProfitCtr = True
            tgcDropdown.SetText txtBox, rsTemp!prft_name & ""
        End If
        
    End If
        
    If fnValidProfitCtr Then
        Select Case txtBox.TabIndex
            Case txtFromPrftCtr.TabIndex
            
                If IsNumeric(txtToPrftCtr.Text) Then
                     
                     If CInt(txtBox.Text) > CInt(txtToPrftCtr.Text) Then
                        fnValidProfitCtr = False
                        cValidate.SetErrorMessage txtBox, "From Profit Center is greater than To Profit center"
                     End If
                
                End If
            
            Case txtToPrftCtr.TabIndex
                
                If IsNumeric(txtFromPrftCtr.Text) Then
                     
                     If CInt(txtBox.Text) < CInt(txtFromPrftCtr.Text) Then
                        fnValidProfitCtr = False
                        cValidate.SetErrorMessage txtBox, "To Profit Center is less than From Profit center"
                     End If
                
                End If
            
        End Select
        
    End If
                
            
End Function

Private Sub subInitializeSpreadSheet()
    Dim vHeaderWidth As Variant
    Dim vHeaderNames As Variant
    Dim vHeaderFields As Variant
    Dim vDetailWidth As Variant
    Dim vDetailNames As Variant
    Dim vDetailFields As Variant
    Dim sDecimal12Pattern As String
    Dim sQuantityPattern As String
    Dim sTotalAmountPattern  As String
    Dim i As Integer
    
    vHeaderWidth = Array(0.12, 0.17, 0.12, 0.11, 0.16, 0.12, 0.14, 0.08)
    vHeaderNames = Array("Profit Cntr.", "Profit Cntr Name", "Date", "Vendor #", "Vendor Name", "Invoice #", "Invoice Total", "Posted")
    vHeaderFields = Array("rsphh_prft_ctr", "prft_name", "rsphh_rpt_date", "rsphh_vendor", "pn_name", "rsphh_invoice" _
                         , "rsphh_invoice_amt", "rsphh_status")
    vDetailWidth = Array(0.08, 0.12, 0.16, 0.08, 0.08, 0.08, 0.08, 0.08, 0.08, 0.08, 0.08)
    vDetailNames = Array("Line", "Item Code", "Description", "UOM", "Qty.", "Cost", "Ext. Cost", "PB Cost", "Ext. PB Cost", "Retail", "Ext. Retail")
    vDetailFields = Array("rsphd_line_nbr", "rsphd_code", "rsbi_desc", "rsphd_stock_unit", "rsphd_qty", "rsphd_cost", "Ext. cost", "rsbp_cost", "Ext. PB Cost", _
                        "rsphd_retail", "Ext. Retail")

    
    sTotalAmountPattern = tfnDecimalPattern(12, 2, False, True)
    sDecimal12Pattern = tfnDecimalPattern(12, 3, False, True)
    sQuantityPattern = tfnDecimalPattern(12, 3, True, True)
    
    While tblHeader.Columns.Count > 0
        tblHeader.Columns.Remove 0
    Wend
    
    For i = 0 To UBound(vHeaderNames)
        tblHeader.Columns.Add i
        
        With tblHeader.Columns(i)
            .Caption = vHeaderNames(i)
            .Width = vHeaderWidth(i) * tblHeader.Width - 50
            .DataField = vHeaderFields(i)
            .Visible = True
        End With
        
    Next
    
    Set tgmHeader = New clsTGSpreadSheet
    
    With tgmHeader
        Set .Form = Me
        Set .StatusBar = ffraStatusbar
        Set .Table = tblHeader
        
        .AddEditColumn ColHeaderInvAmount, "Enter total Invoice Amount (can't edit if processed)", sTotalAmountPattern
        ColHdnHeaderShift = .AddHiddenField("rsphh_shift")
        ColHdnHeaderTerm = .AddHiddenField("rsphh_std_term")
        ColHdnHeaderType = .AddHiddenField("rsphh_type")
        ColHdnHeaderDraft = .AddHiddenField("rsphh_draft_nbr")
        colHdnHeaderInvDate = .AddHiddenField("rsphh_inv_date")
        .AllowAddNew = False
    End With
    
    While tblDetail.Columns.Count > 0
        tblDetail.Columns.Remove 0
    Wend
    
    For i = 0 To UBound(vDetailNames)
        tblDetail.Columns.Add i
        
        With tblDetail.Columns(i)
            .Caption = vDetailNames(i)
            .Width = vDetailWidth(i) * tblDetail.Width
            .DataField = vDetailFields(i)
            .Visible = True
        End With
        
    Next
    
    Set tgmDetail = New clsTGSpreadSheet
    
    With tgmDetail
        Set .Form = Me
        Set .StatusBar = ffraStatusbar
        Set .Table = tblDetail
        .AllowAddNew = True
    End With
    
    With tgmDetail
        .AddEditColumn ColDetailItemCode, "Enter Item Code", "^P{1,10}$"
        .AddESCColumn ColDetailItemDesc, ColDetailItemCode
        .AddEditColumn ColDetailItemDesc, "Enter Item Description", "^P{0,18}$"
        .AddEditColumn ColDetailUOM, "Enter Unit of Measure", "^P{1,5}$"
        .AddEditColumn ColDetailQty, "Enter Quantity", sQuantityPattern
        .AddEditColumn ColDetailCost, "Enter Cost", sDecimal12Pattern
        colHdnDetailFlag = .AddHiddenField("cost flag")
    End With
    
End Sub

Private Sub subSetupDropDown()
    
    subEnableSearchButton cmdSearch, True
    
    Set tgfDropdown = New clsFloatingDropDown
    
    With tgfDropdown
        Set .SearchButton = cmdSearch
        Set .MainTable = tblDetail
        Set .DropDownTable = tblDropDown
        Set .DataBase = t_dbMainDatabase
        Set .DataLink = datDropDown
        Set .Form = Me
        Set .EditClass = tgmDetail
        .DefaultCursorControl = True
    
        .AddDropDown 1
        .AddColumn ColDetailItemCode, "rsbi_code"
        .ColumnCaption(ColDetailItemCode) = "Code"
        .ColumnType(ColDetailItemCode) = .COLUMN_TYPE_STRING
        .AddColumn ColDetailItemDesc, "rsbi_desc"
        .ColumnCaption(ColDetailItemDesc) = "Description"
        .ColumnType(ColDetailItemDesc) = .COLUMN_TYPE_STRING
    End With
    
End Sub

Private Sub subInitFullText()
   Set objFullHeaderText = New clsColumnExtension
   
   With objFullHeaderText
        Set .Form = Me
        Set .PictureBox = picFullHeaderText
        Set .StatusBar = ffraStatusbar
        Set .Table = tblHeader
        .AddColumn ColHeaderPrftDesc, "Profit Cntr. Name"
        .AddColumn ColHeaderVendor, "Vendor Number"
        .AddColumn ColHeaderVendorName, "Vendor Name"
        .AddColumn ColHeaderInvoice, "Invoice Number"
    End With
    
    Set objFullDetailText = New clsColumnExtension
    
    With objFullDetailText
        Set .Form = Me
        Set .PictureBox = picFullDetailText
        Set .StatusBar = ffraStatusbar
        Set .Table = tblDetail
        .AddColumn ColDetailItemDesc, "Item Description"
    End With
    
End Sub

Private Sub chkAllInvoice_GotFocus()
    tfnSetStatusBarMessage "Check to select all Invoice including processed."
End Sub

Private Sub chkAllInvoice_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Then
        subSetFocus cmdQuery
    End If
    
End Sub

Private Sub cmdDetailCancel_Click()
    
    If g_bDetailDataChanged Then
        
        If Not tfnCancelExit(t_szCANCEL_MESSAGE) Then
            Exit Sub
        End If
                
    End If
    
    g_bDetailDataChanged = False
    subResetExceptions
    
    If fnLoadDetailData(tgmHeader.GetCurrentRowNumber) Then
    End If
    
End Sub

Private Sub cmdDetailCancel_GotFocus()
    tfnSetStatusBarMessage "Cancel"
End Sub

Private Sub cmdDetailDelete_Click()
    Dim lCurHeaderRow As Long
    Dim lCurDetailRow As Long
    Dim lVendor As Long
    Dim lInvoice As Long
    Dim lLineNo As Long
    Dim sProcessed As String
    
    lCurHeaderRow = tgmHeader.GetCurrentRowNumber
    lVendor = tfnRound(tgmHeader.CellValue(ColHeaderVendor, lCurHeaderRow))
    lInvoice = tfnRound(tgmHeader.CellValue(ColHeaderInvoice, lCurHeaderRow))
    sProcessed = tgmHeader.CellValue(ColHeaderStatus, lCurHeaderRow)
    lCurDetailRow = tgmDetail.GetCurrentRowNumber
    lLineNo = tfnRound(tgmDetail.CellValue(ColDetailLine, lCurDetailRow))
        
    If lLineNo = 0 Then
        Exit Sub
    End If
    
    If tgmDetail.RowCount > 0 Then
        
        If Not tfnCancelExit("Are you sure you want to delete the current record?") Then
            Exit Sub
        End If
        
        If fnDeleteDetailData(lVendor, lInvoice, lLineNo) Then
            tgmDetail.DeleteRow
        End If
        
        If tgmDetail.RowCount > 0 Then
            
            If fnLoadDetailData(lCurHeaderRow) Then
                '
            End If
            
        End If
                                  
        subResetExceptions
    End If

    If tgmDetail.RowCount < 1 Then
        subEnableDetailCancelBtn False
        subEnableDetailDeleteBtn False
        subEnableDetailOKBtn False
        
        If Not tfnCancelExit("Do you want to delete the header since it has no more details?") Then
            Exit Sub
        End If
        
        If fnDeleteData(lVendor, lInvoice, sProcessed) Then
            subUnlockHeader False
            etabMain.CurrTab = 0
            etabMain.TabEnabled(1) = False
            tgmHeader.DeleteRow
        End If
                
    End If
         
    If tgmHeader.RowCount < 1 Then
        subEnableHeaderButtons False
        t_bDataChanged = False
        cmdExitCancelBtn_Click
    End If

End Sub

Private Sub cmdDetailDelete_GotFocus()
    tfnSetStatusBarMessage "Delete"
End Sub

Private Sub cmdExceptions_Click()

    subLoadDataNoException g_bExceptionsClicked
    
    If tgmDetail.RowCount < 1 Then
        subEnableDetailDeleteBtn False
        subEnableDetailOKBtn False
    Else
        subEnableDetailDeleteBtn True
        subEnableDetailOKBtn cmdOk.Enabled
    End If
    
    subCheckButtonStatus
End Sub

Private Sub cmdExceptions_GotFocus()
    
    If cmdExceptions.Caption = "Exceptions" Then
        tfnSetStatusBarMessage "Exceptions"
    Else
        tfnSetStatusBarMessage "All Data"
    End If
    
End Sub

Private Sub cmdFromPrftCtr_Click()
    subShowHourGlass
    tgcDropdown.ComboSQL(txtFromPrftCtr) = fnSetPrftComboSQL(txtFromPrftCtr)
    tgcDropdown.Click cmdFromPrftCtr
    subHideHourGlass
End Sub

Private Sub cmdFromPrftName_Click()
    subShowHourGlass
    tgcDropdown.ComboSQL(txtFromPrftName) = fnSetPrftComboSQL(txtFromPrftCtr)
    tgcDropdown.Click cmdFromPrftName
    subHideHourGlass
End Sub

Private Sub cmdHeaderDelete_Click()
    Dim lCurHeaderRow As Long
    Dim lVendor As Long
    Dim lInvoice As Long
    Dim nPrftctr As Integer
    Dim dtRptDate As Date
    Dim lShiftNum As Long
    Dim sProcessed As String
    
    lCurHeaderRow = tgmHeader.GetCurrentRowNumber
    lVendor = tfnRound(tgmHeader.CellValue(ColHeaderVendor, lCurHeaderRow))
    lInvoice = tfnRound(tgmHeader.CellValue(ColHeaderInvoice, lCurHeaderRow))
    nPrftctr = tfnRound(tgmHeader.CellValue(ColHeaderPrftctr, lCurHeaderRow))
    dtRptDate = tfnFormatDate(tgmHeader.CellValue(ColHeaderRptDate, lCurHeaderRow))
    lShiftNum = tfnRound(tgmHeader.CellValue(ColHdnHeaderShift, lCurHeaderRow))
    sProcessed = tgmHeader.CellValue(ColHeaderStatus, lCurHeaderRow)
    
    If tgmHeader.RowCount > 0 Then
        
        If Not tfnCancelExit("Are you sure you want to delete the current record?") Then
            Exit Sub
        End If
            
        If fnDeleteData(lVendor, lInvoice, sProcessed) Then
            
            'If nPrftctr = g_nHoldSelectedPrftCtr And dtRptDate = g_dtHoldRptDate And lShiftNum = g_lHoldShiftNum Then
            If lVendor = g_lHoldSelectedVendor And lInvoice = g_lHoldSelectedInvoice Then
                subUnlockHeader False
            End If
            
            etabMain.TabEnabled(1) = False
            tgmHeader.DeleteRow
        End If
        
    End If
    
    If tgmHeader.RowCount < 1 Then
        subEnableHeaderButtons False
        t_bDataChanged = False
        cmdExitCancelBtn_Click
    End If
    
    g_lSelectedHeaderRow = -99

End Sub

Private Sub cmdHeaderDelete_GotFocus()
    tfnSetStatusBarMessage "Delete"
End Sub

Private Sub cmdHeaderSelect_Click()
    Dim lCurHeaderRow As Long
    Dim nPrftctr As Integer
    Dim dtRptDate As Date
    Dim lShiftNum As Long
    Dim lVendor As Long
    Dim lInvoice As Long

    lCurHeaderRow = tgmHeader.GetCurrentRowNumber
    g_lSelectedHeaderRow = lCurHeaderRow
    
    nPrftctr = tfnRound(tgmHeader.CellValue(ColHeaderPrftctr, lCurHeaderRow))
    dtRptDate = tfnFormatDate(tgmHeader.CellValue(ColHeaderRptDate, lCurHeaderRow))
    lShiftNum = tfnRound(tgmHeader.CellValue(ColHdnHeaderShift, lCurHeaderRow))
    lVendor = tfnRound(tgmHeader.CellValue(ColHeaderVendor, lCurHeaderRow))
    lInvoice = tfnRound(tgmHeader.CellValue(ColHeaderInvoice, lCurHeaderRow))
    
    subShowHourGlass
    tfnSetStatusBarCorrect "Populating Data. Please wait..."
    
    'If nPrftctr <> g_nHoldSelectedPrftCtr Or dtRptDate <> g_dtHoldRptDate Or lShiftNum <> g_lHoldShiftNum Then
    If lVendor <> g_lHoldSelectedVendor Or lInvoice <> g_lHoldSelectedInvoice Then
        subUnlockHeader False
    End If
    
    
    subResetExceptions
    
    If fnLoadDetailData(lCurHeaderRow) Then
        
        If tgmHeader.CellValue(ColHeaderStatus, lCurHeaderRow) = "N" Then
            'tgmDetail.Locked = False
            
            If Not fnlockHeader(lVendor, lInvoice, nPrftctr, dtRptDate, lShiftNum) Then
                etabMain.TabEnabled(1) = False
                etabMain.CurrTab = 0
                tgmDetail.ClearData
                tfnSetStatusBarError "Unable to load details"
                GoTo EXITHERE
            End If
            
        Else
            subEnableDetailOKBtn False
            MsgBox "The data has already been posted, if you want to change something, go to Retail Sales Purchases Entry.", vbInformation
            'tgmDetail.Locked = True
        End If
        
        etabMain.TabEnabled(1) = True
        etabMain.CurrTab = 1
        subSetTabDetailCaption lCurHeaderRow
        subEnableDetailCancelBtn False
        subEnableDetailDeleteBtn True
        subSetFocus tblDetail
    Else
        subSetFocus cmdExitCancelBtn
    End If
    
EXITHERE:
    g_nHoldSelectedPrftCtr = nPrftctr
    g_dtHoldRptDate = dtRptDate
    g_lHoldShiftNum = lShiftNum
    g_lHoldSelectedVendor = lVendor
    g_lHoldSelectedInvoice = lInvoice
    subHideHourGlass
    
End Sub

Private Sub cmdHeaderSelect_GotFocus()
    tfnSetStatusBarMessage "Select"
End Sub

Private Sub cmdOK_Click()
    
    If fnUpdateDetail() Then
        subEnableUpdateBtn True
        subSetFocus cmdUpdateInsertBtn
'        tgmDetail.ClearData
'        etabMain.CurrTab = 0
'        etabMain.TabEnabled(1) = False
        subEnableDetailOKBtn False
    Else
        subEnableUpdateBtn False
    End If
    
    subResetExceptions
End Sub

Private Sub cmdOk_GotFocus()
    tfnSetStatusBarMessage "OK"
End Sub

Private Sub cmdQuery_Click()
        
    subShowHourGlass
    g_lSelectedHeaderRow = -99
    
    If fnLoadHeaderData() Then
        subEnableHeaderButtons True
        subEnableProfitctr False
        
        subSetFocus tblHeader
    End If
    
    subUnlockHeader False
    
    subEnableCancel True
    subHideHourGlass
End Sub


Private Sub cmdQuery_GotFocus()
    tfnSetStatusBarMessage "Query"
End Sub

Private Sub cmdSearch_Click()
    Dim lCurHeaderRow As Long
    Dim lVendor As Long
    
    lCurHeaderRow = tgmHeader.GetCurrentRowNumber
    lVendor = tfnRound(tgmHeader.CellValue(ColHeaderVendor, lCurHeaderRow))
    tgfDropdown.DropDownSQL(1) = "SELECT rsbi_code, rsbi_desc FROM rs_b_item Where rsbi_vendor = " & lVendor
    tgfDropdown.ButtonClick
End Sub

Private Sub cmdSearch_LostFocus()
    tgfDropdown.LostFocus cmdSearch
End Sub

Private Sub cmdToPrftCtr_Click()
    subShowHourGlass
    tgcDropdown.ComboSQL(txtToPrftCtr) = fnSetPrftComboSQL(txtToPrftCtr)
    tgcDropdown.Click cmdToPrftCtr
    subHideHourGlass
End Sub

Private Sub cmdToPrftName_Click()
    subShowHourGlass
    tgcDropdown.ComboSQL(txtToPrftName) = fnSetPrftComboSQL(txtToPrftCtr)
    tgcDropdown.Click cmdToPrftName
    subHideHourGlass
End Sub

Private Sub etabMain_GotFocus()
    
    Select Case etabMain.CurrTab
        Case 0
            
            If efraProfitCtr.Enabled Then
                subSetFocus txtFromPrftCtr
            ElseIf tblHeader.Enabled Then
                subSetFocus tblHeader
            End If
            
        Case 1
            If tblDetail.Enabled Then
                subSetFocus tblDetail
            End If
            
    End Select

End Sub

Private Sub etabMain_Switch(OldTab As Integer, NewTab As Integer, CANCEL As Integer)
    
    If Not etabMain.TabEnabled(NewTab) Then
        CANCEL = True
        Exit Sub
    End If
    
    Select Case NewTab
        Case 0
            
            If efraProfitCtr.Enabled Then
                subSetFocus txtFromPrftCtr
            ElseIf tblHeader.Enabled Then
                subSetFocus tblHeader
            End If
            
            subEnableUpdateBtn False
        Case 1
            If tblDetail.Enabled Then
                subSetFocus tblDetail
            End If
            
            Dim lCurHeaderRow As Long
            
            lCurHeaderRow = tgmHeader.GetCurrentRowNumber
            
            If tgmHeader.CellValue(ColHeaderStatus, lCurHeaderRow) = "N" Then
                tgmDetail.Locked = False
            Else
                tgmDetail.Locked = True
            End If
            
            subCheckButtonStatus
    End Select

End Sub

Private Sub tblDetail_Click()
    tgfDropdown.TableClick tblDetail
End Sub

Private Sub tblDetail_FirstRowChange()
    tgfDropdown.FirstRowChange
End Sub

Private Sub tblDetail_Scroll(CANCEL As Integer)
    tgfDropdown.Scroll
End Sub

Private Sub tblDropDown_Click()
    tgfDropdown.TableClick tblDropDown
End Sub

Private Sub tblDropDown_KeyPress(KeyAscii As Integer)
    tgfDropdown.Keypress tblDropDown, KeyAscii
End Sub

Private Sub tblDropDown_LostFocus()
    tgfDropdown.LostFocus tblDropDown
End Sub

Private Sub tblDropDown_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    tgfDropdown.MouseUp y
End Sub

Private Sub tblDropDown_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    tgfDropdown.RowColChange tblDropDown
End Sub

Private Sub tblDropDown_SelChange(CANCEL As Integer)
    CANCEL = True
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
    App.HelpFile = szHelp7_11
End Sub

Private Sub Form_Unload(CANCEL As Integer)
    tfnUnlockRow
    
    On Error Resume Next
    
    Set tgcDropdown = Nothing
    Set cValidate = Nothing
    Set objErrHandler = Nothing
    Set tgfDropdown = Nothing
    Set objCurrTabControl = Nothing
    Set tgmDetail = Nothing
    Set tgmHeader = Nothing
    Set objFullDetailText = Nothing
    Set objFullHeaderText = Nothing
    
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
    subInitValidation  'set up validation object
    subInitializeSpreadSheet
    subSetupDropDown
    subInitFullText
    Screen.MousePointer = vbHourglass
    
    ' ^ ^ ^ ^ ^ ^ ^
    ' | | | | | | |
    '***************************************************
    
    Me.Enabled = True
    
    '***************************************************
    ' SET YOUR FIRST STATUSBAR MESSAGE HERE
    ' | | | | | | |
    ' v v v v v v v
    
    tfnResetScreen 'set the default screen
    Screen.MousePointer = vbHourglass
    
    t_bStartupFlag = False
    Screen.MousePointer = vbHourglass
    
    ' ^ ^ ^ ^ ^ ^ ^
    ' | | | | | | |
    '***************************************************
    
    tmrKeyBoard.Enabled = True

    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_GotFocus()
    tmrKeyBoard.Enabled = True
End Sub

Private Sub Form_LostFocus()
    tmrKeyBoard.Enabled = False
End Sub

Private Sub cmdDeleteBtn_GotFocus()
    tfnSetStatusBarMessage t_szDELETE
End Sub

Private Sub cmdUpdateInsertBtn_GotFocus()
    tfnSetStatusBarMessage t_szINSERT
End Sub

Private Sub cmdExitCancelBtn_GotFocus()
    
    If cmdExitCancelBtn.Caption = t_szCAPTION_EXIT Then
        tfnSetStatusBarMessage t_szEXIT
    Else
        tfnSetStatusBarMessage t_szCANCEL
    End If
    
End Sub

Private Sub cmdUpdateInsertBtn_Click()
    
    If fnProcessData() Then
        subUnlockHeader True
        cmdUpdateInsertBtn.Enabled = False
        etabMain.CurrTab = 0
        'etabMain.TabEnabled(1) = False
        tfnSetStatusBarCorrect "Data have been inserted successfully"
    Else
        tfnSetStatusBarError "Unable to insert data"
    End If
    
End Sub

Private Sub cmdExitCancelBtn_Click()
    
    If cmdExitCancelBtn.Caption = t_szCAPTION_CANCEL Then
        cmdTBCancel_Click
    Else
        cmdTBExit_Click
    End If
    
End Sub

'=====================
'Toolbar Button Events
'============
'Menu  Events
'============

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
    
    subEnableUpdateBtn False
    subSetExitCancelBtn "EXIT"
    subEnableProfitctr True
    subEnableHeaderButtons False
    cmdUpdateInsertBtn.Caption = t_szCAPTION_INSERT
    mnuUpdateInsert.Caption = t_szCAPTION_INSERT
    
    frmContext.ButtonEnabled(PRINT_UP) = False
    frmContext.ButtonEnabled(COPY_UP) = False
    
    mnuExit.Enabled = True
    mnuPrint.Enabled = False
    
    mnuCopy.Enabled = False
    mnuPaste.Enabled = False
    
    etabMain.CurrTab = 0
#If Not PROTOTYPE Then
    etabMain.TabEnabled(1) = False
#Else
    efraProfitCtr.Enabled = False
    efraHeader.Enabled = False
    efraDetail.Enabled = False
    subEnableHeaderButtons False
    cmdQuery.Enabled = False
    subEnableDetailExceptions False
    subEnableDetailCancelBtn False
    subEnableDetailDeleteBtn False
    subEnableDetailOKBtn False
#End If
    txtFromPrftCtr.Text = ""
    txtFromPrftName.Text = ""
    txtToPrftCtr.Text = ""
    txtToPrftName.Text = ""
    chkAllInvoice.value = vbUnchecked
    tgmHeader.ClearData
    tgmDetail.ClearData
    cValidate.ResetFlags
    g_bDetailDataChanged = False
    subSetFocus txtFromPrftCtr
End Sub

Private Sub subSetFocus(v As Control)
    On Error Resume Next
    v.SetFocus
End Sub

Private Sub subEnableHeaderButtons(bOnOff As Boolean)
    cmdHeaderDelete.Enabled = bOnOff
    cmdHeaderSelect.Enabled = bOnOff
    cmdQuery.Enabled = Not bOnOff
End Sub

Private Sub subEnableProfitctr(bOnOff As Boolean)
    efraProfitCtr.Enabled = bOnOff
    subEnableSearchButton cmdFromPrftCtr, bOnOff
    subEnableSearchButton cmdFromPrftName, bOnOff
    subEnableSearchButton cmdToPrftCtr, bOnOff
    subEnableSearchButton cmdToPrftName, bOnOff
    chkAllInvoice.Enabled = bOnOff
    tblHeader.Enabled = Not bOnOff
End Sub

Private Sub tblHeader_AfterColEdit(ByVal ColIndex As Integer)
    Dim lCurHeaderRow As Long
    lCurHeaderRow = tgmHeader.GetCurrentRowNumber
    
    tgmHeader.AfterColEdit ColIndex
    
    If ColIndex = ColHeaderInvAmount Then
        tgmHeader.CellValue(ColHeaderInvAmount, lCurHeaderRow) = tfnRound(tgmHeader.CellValue(ColHeaderInvAmount, lCurHeaderRow), 2)
        
        If CStr(tgmHeader.CellValue(ColHeaderInvoice, lCurHeaderRow)) = lblInvoice.Caption _
            And CStr(tgmHeader.CellValue(ColHeaderVendor, lCurHeaderRow)) = lblVendor.Caption Then
            lblAmount.Caption = tgmHeader.CellValue(ColHeaderInvAmount, lCurHeaderRow)
        End If
        
    End If
    
End Sub

Private Sub tblHeader_BeforeColEdit(ByVal ColIndex As Integer, ByVal KeyAscii As Integer, CANCEL As Integer)
    Dim lCurHeaderRow As Long
    
    lCurHeaderRow = tgmHeader.GetCurrentRowNumber
    
    If tgmHeader.CellValue(ColHeaderStatus, lCurHeaderRow) = "Y" Then
        CANCEL = True
    End If
    
    tgmHeader.BeforeColEdit ColIndex, KeyAscii, CANCEL
    
End Sub

Private Sub tblHeader_Change()
    tgmHeader.Change
End Sub

Private Sub tblHeader_GotFocus()
    tgmHeader.GotFocus
End Sub

Private Sub tblHeader_KeyDown(KeyCode As Integer, Shift As Integer)
    tgmHeader.KeyDown KeyCode, Shift
End Sub

Private Sub tblHeader_KeyPress(KeyAscii As Integer)
    
    If Not tgmHeader.Keypress(KeyAscii) Then
        KeyAscii = 0
    End If
    
End Sub

Private Sub tblHeader_LostFocus()
    tgmHeader.LostFocus
End Sub

Private Sub tblHeader_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    Dim lRow As Long
    
    'add this one, otherwise when mouse click on the other col instead of Invoice amount, it will not work.
    'lRow = tblHeader.Row
    tgmHeader.RowColChange LastRow, LastCol
    'tblHeader.Row = lRow '- tfnRound(tblHeader.FirstRow)
    ''''''''
    
    If tgmHeader.RowCount > 0 Then
        
        If tgmHeader.GetCurrentRowNumber <> g_lSelectedHeaderRow Then
            etabMain.TabEnabled(1) = False
        Else
            etabMain.TabEnabled(1) = True
        End If
        
    End If
    
End Sub

Private Sub tblHeader_SelChange(CANCEL As Integer)
    CANCEL = True
End Sub

Private Sub tblHeader_UnboundReadData(ByVal RowBuf As DBTrueGrid.RowBuffer, StartLocation As Variant, ByVal ReadPriorRows As Boolean)
    tgmHeader.ReadData RowBuf, StartLocation, ReadPriorRows
End Sub

Private Sub tblDetail_AfterColEdit(ByVal ColIndex As Integer)
    Dim lCurDetailRow As Long
    
    tgmDetail.AfterColEdit ColIndex
    g_bDetailDataChanged = True
    lCurDetailRow = tgmDetail.GetCurrentRowNumber
    
    If ColIndex = ColDetailQty Or ColIndex = ColDetailUOM Then
        tgmDetail.CellValue(ColDetailExtCost, lCurDetailRow) = tfnRound(tfnRound(tgmDetail.CellValue(ColDetailCost, lCurDetailRow), 3) _
                        * tfnRound(tgmDetail.CellValue(ColDetailQty, lCurDetailRow), 3), 3)
        tgmDetail.CellValue(ColDetailExtPBCost, lCurDetailRow) = tfnRound(tfnRound(tgmDetail.CellValue(ColDetailPBCost, lCurDetailRow), 3) _
                        * tfnRound(tgmDetail.CellValue(ColDetailQty, lCurDetailRow), 3), 3)
        tgmDetail.CellValue(ColDetailExtRetail, lCurDetailRow) = tfnRound(tfnRound(tgmDetail.CellValue(ColDetailRetail, lCurDetailRow), 3) _
                        * tfnRound(tgmDetail.CellValue(ColDetailQty, lCurDetailRow), 3), 3)
    ElseIf ColIndex = ColDetailCost Then
        tgmDetail.CellValue(ColDetailExtCost, lCurDetailRow) = tfnRound(tfnRound(tgmDetail.CellValue(ColDetailCost, lCurDetailRow), 3) _
                        * tfnRound(tgmDetail.CellValue(ColDetailQty, lCurDetailRow), 3), 3)
    End If
    
    subEnableDetailCancelBtn True
    subCheckButtonStatus
    
End Sub

Private Sub tblDetail_BeforeColEdit(ByVal ColIndex As Integer, ByVal KeyAscii As Integer, CANCEL As Integer)
    tgmDetail.BeforeColEdit ColIndex, KeyAscii, CANCEL
    subEnableDetailOKBtn False
End Sub

Private Sub tblDetail_Change()
    tgmDetail.Change
    
    If tblDetail.col = ColDetailItemCode Then
        tblDetail.Columns(ColDetailItemDesc).Text = ""
    End If
    
End Sub

Private Sub tblDetail_GotFocus()
    tgmDetail.GotFocus
    tgfDropdown.GotFocus
    
    If tgfDropdown.ValidSelection Then
        subCheckButtonStatus
        subEnableDetailCancelBtn True
    End If
    
    If tgmHeader.CellValue(ColHeaderStatus, tgmHeader.GetCurrentRowNumber) = "Y" Then
        tfnSetStatusBarMessage "Invoice Detail"
    End If
    
End Sub

Private Sub tblDetail_KeyDown(KeyCode As Integer, Shift As Integer)
    tgfDropdown.KeyDown tblDetail, KeyCode
    tgmDetail.KeyDown KeyCode, Shift
End Sub

Private Sub tblDetail_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Then
        
        If tblDetail.col = ColDetailItemCode Or tblDetail.col = ColDetailItemDesc Then
            Dim lCurHeaderRow As Long
            Dim lVendor As Long
            
            lCurHeaderRow = tgmHeader.GetCurrentRowNumber
            lVendor = tfnRound(tgmHeader.CellValue(ColHeaderVendor, lCurHeaderRow))
            tgfDropdown.DropDownSQL(1) = "SELECT rsbi_code, rsbi_desc FROM rs_b_item Where rsbi_vendor = " & lVendor
        End If
        
    End If
    
    If tgfDropdown.ValidSelection Then
        subCheckButtonStatus
        subEnableDetailCancelBtn True
    End If
    
    If Not tgmDetail.Keypress(KeyAscii) Then
        KeyAscii = 0
    End If
    
End Sub

Private Sub tblDetail_LostFocus()
    tgfDropdown.LostFocus tblDetail
    tgmDetail.LostFocus
End Sub

Private Sub tblDetail_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    
    If tgfDropdown.RowColChange(tblDetail) Then
        tgmDetail.RowColChange LastRow, LastCol
    End If
    
End Sub

Private Sub tblDetail_SelChange(CANCEL As Integer)
    CANCEL = True
End Sub

Private Sub tblDetail_UnboundReadData(ByVal RowBuf As DBTrueGrid.RowBuffer, StartLocation As Variant, ByVal ReadPriorRows As Boolean)
    tgmDetail.ReadData RowBuf, StartLocation, ReadPriorRows
End Sub

Private Sub tmrKeyboard_Timer() 'status bar timer - 250ms
    tfnUpdateStatusBar Me 'process the status bar
    
    If etabMain.CurrTab = 0 Then
        objFullHeaderText.ShowColumnExt
    ElseIf etabMain.CurrTab = 1 Then
        objFullDetailText.ShowColumnExt
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

Private Sub subEnableDetailOKBtn(bOnOff As Boolean)
    cmdOk.Enabled = bOnOff
End Sub

Private Sub subEnableDetailCancelBtn(bOnOff As Boolean)
    cmdDetailCancel.Enabled = bOnOff
End Sub

Private Sub subEnableDetailDeleteBtn(bOnOff As Boolean)
    Dim lCurHeaderRow As Long
            
    lCurHeaderRow = tgmHeader.GetCurrentRowNumber
            
    If tgmDetail.RowCount > 0 And tgmHeader.CellValue(ColHeaderStatus, lCurHeaderRow) = "Y" Then
        cmdDetailDelete.Enabled = False
    Else
        cmdDetailDelete.Enabled = bOnOff
    End If
    
End Sub

Private Sub subEnableDetailExceptions(bOnOff As Boolean)
    cmdExceptions.Enabled = bOnOff
End Sub

Private Sub Form_Resize()
    frmContext.FormResize
End Sub

Private Sub subSetupToolBar()
    
    With frmContext
        .BeginSetupToolbar Me
        .EndSetupToolbar
        .HelpFile = szHelpFileName
    End With
    
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
    
    subUnlockHeader False
    
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
    
    subUnlockHeader False
    Unload Me
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

Private Sub subEnableCancel(bStatus As Boolean)
    mnuCancel.Enabled = bStatus
    frmContext.ButtonEnabled(CANCEL_UP) = bStatus
    
    If bStatus Then
        cmdExitCancelBtn.Caption = t_szCAPTION_CANCEL
    Else
        cmdExitCancelBtn.Caption = t_szCAPTION_EXIT
    End If
    
End Sub

Private Sub txtFromPrftCtr_Change()
    cValidate.Change txtFromPrftCtr
    tgcDropdown.Change txtFromPrftCtr
End Sub

Private Sub txtFromPrftCtr_GotFocus()
    tgcDropdown.GotFocus txtFromPrftCtr
    cValidate.GotFocus txtFromPrftCtr
    
    If tgcDropdown.SingleRecordSelected Then
        subSetFocus txtToPrftCtr
    End If
    
End Sub

Private Sub txtFromPrftCtr_KeyPress(KeyAscii As Integer)
    Dim bKeyCode As Boolean
    
    If KeyAscii = vbKeyReturn Then
        tgcDropdown.ComboSQL(txtFromPrftCtr) = fnSetPrftComboSQL(txtFromPrftCtr)
        subShowHourGlass
    End If
    
    bKeyCode = tgcDropdown.Keypress(txtFromPrftCtr, KeyAscii)
    
    If Not bKeyCode Then
    
        If KeyAscii = vbKeyReturn Then
          
            If tgcDropdown.SingleRecordSelected Then
                subSetFocus txtToPrftCtr
            End If
          
          KeyAscii = 0
          
       End If
       
    Else
        cValidate.Keypress txtFromPrftCtr, KeyAscii
    End If
    
    subHideHourGlass
End Sub

Private Sub txtFromPrftCtr_LostFocus()
    tgcDropdown.LostFocus txtFromPrftCtr
    cValidate.LostFocus txtFromPrftCtr, cmdFromPrftCtr, txtFromPrftName, cmdFromPrftName, tblComboDropdown
End Sub

Private Sub txtTOPrftCtr_Change()
    cValidate.Change txtToPrftCtr
    tgcDropdown.Change txtToPrftCtr
End Sub

Private Sub txtToPrftCtr_GotFocus()
    tgcDropdown.GotFocus txtToPrftCtr
    cValidate.GotFocus txtToPrftCtr
    
    If tgcDropdown.SingleRecordSelected Then
        subSetFocus chkAllInvoice
    End If
    
End Sub

Private Sub txtToPrftCtr_KeyPress(KeyAscii As Integer)
    Dim bKeyCode As Boolean
    
    If KeyAscii = vbKeyReturn Then
        tgcDropdown.ComboSQL(txtToPrftCtr) = fnSetPrftComboSQL(txtToPrftCtr)
        subShowHourGlass
    End If
    
    bKeyCode = tgcDropdown.Keypress(txtToPrftCtr, KeyAscii)
    
    If Not bKeyCode Then
    
        If KeyAscii = vbKeyReturn Then
          
            If tgcDropdown.SingleRecordSelected Then
                subSetFocus chkAllInvoice
            End If
          
          KeyAscii = 0
          
       End If
       
    Else
        cValidate.Keypress txtToPrftCtr, KeyAscii
    End If
    
    subHideHourGlass
End Sub

Private Sub txtToPrftCtr_LostFocus()
    tgcDropdown.LostFocus txtToPrftCtr
    cValidate.LostFocus txtToPrftCtr, cmdToPrftCtr, txtToPrftName, cmdToPrftName, tblComboDropdown
End Sub

Private Sub txtFromPrftName_Change()
    tgcDropdown.Change txtFromPrftName
End Sub

Private Sub txtFromPrftName_GotFocus()
    tfnSetStatusBarMessage "Enter Profit Center Name"
    tgcDropdown.GotFocus txtFromPrftName

    If tgcDropdown.SingleRecordSelected Then
        subSetFocus txtToPrftCtr
    End If
    
End Sub

Private Sub txtFromPrftName_KeyPress(KeyAscii As Integer)
    Dim bKeyCode As Boolean
    
    If KeyAscii = vbKeyReturn Then
        tgcDropdown.ComboSQL(txtFromPrftName) = fnSetPrftComboSQL(txtFromPrftCtr)
        subShowHourGlass
    End If
    
    bKeyCode = tgcDropdown.Keypress(txtFromPrftName, KeyAscii)
    
    If Not bKeyCode Then
        
        If KeyAscii = vbKeyReturn Then
            
            If tgcDropdown.SingleRecordSelected Then
                subSetFocus txtToPrftCtr
            End If
            
            KeyAscii = 0
        End If
        
    End If
    
End Sub

Private Sub txtFromPrftName_LostFocus()

    If cValidate.LostFocus(txtFromPrftCtr, cmdFromPrftCtr, txtFromPrftName, cmdFromPrftName, tblComboDropdown) Then
        tgcDropdown.LostFocus txtFromPrftName
    End If
    
End Sub

Private Sub txtToPrftName_Change()
    tgcDropdown.Change txtToPrftName
End Sub

Private Sub txtToPrftName_GotFocus()
    tfnSetStatusBarMessage "Enter Profit Center Name"
    tgcDropdown.GotFocus txtToPrftName

    If tgcDropdown.SingleRecordSelected Then
        subSetFocus chkAllInvoice
    End If
    
End Sub

Private Sub txtToPrftName_KeyPress(KeyAscii As Integer)
    Dim bKeyCode As Boolean
    
    If KeyAscii = vbKeyReturn Then
        tgcDropdown.ComboSQL(txtToPrftName) = fnSetPrftComboSQL(txtToPrftCtr)
        subShowHourGlass
    End If
    
    bKeyCode = tgcDropdown.Keypress(txtToPrftName, KeyAscii)
    
    If Not bKeyCode Then
        
        If KeyAscii = vbKeyReturn Then
            
            If tgcDropdown.SingleRecordSelected Then
                subSetFocus chkAllInvoice
            End If
            
            KeyAscii = 0
        End If
        
    End If
    
End Sub

Private Sub txtToPrftName_LostFocus()

    If cValidate.LostFocus(txtToPrftCtr, cmdToPrftCtr, txtToPrftName, cmdToPrftName, tblComboDropdown) Then
        tgcDropdown.LostFocus txtToPrftName
    End If
    
End Sub

Private Function fnSetPrftComboSQL(txtBox As Textbox) As String
    Dim strSQL As String
    
    strSQL = "SELECT prft_ctr, prft_name FROM sys_prft_ctr WHERE "
    strSQL = strSQL & " (prft_type = 'R' OR prft_type = 'B')"
          
    Select Case txtBox.TabIndex
        Case txtFromPrftCtr.TabIndex, txtFromPrftName.TabIndex
            
            If cValidate.ValidInput(txtToPrftCtr) And Trim(txtToPrftCtr) <> "" Then
                strSQL = strSQL & " AND prft_ctr <= " & txtToPrftCtr
            End If
            
        Case txtToPrftCtr.TabIndex, txtToPrftName.TabIndex
            
            If cValidate.ValidInput(txtFromPrftCtr) And Trim(txtFromPrftCtr) <> "" Then
                strSQL = strSQL & " AND prft_ctr >= " & txtFromPrftCtr
            End If
            
    End Select
    
    fnSetPrftComboSQL = strSQL
    
End Function

Private Function fnLoadHeaderData() As Boolean
    Dim strSQL As String
    
    If cValidate.FirstInvalidInput >= 0 Then
        MsgBox "Invalid Profit center entered, Try again", vbInformation
        Exit Function
    End If
    
    If Trim(txtFromPrftCtr) <> "" And Trim(txtToPrftCtr) = "" Then
        txtToPrftCtr.Text = txtFromPrftCtr.Text
        txtToPrftName.Text = txtFromPrftName.Text
    End If
    
    strSQL = "SELECT rsphh_prft_ctr, prft_name, rsphh_rpt_date, rsphh_vendor, pn_name, rsphh_invoice, rsphh_invoice_amt"
    strSQL = strSQL & " ,rsphh_status,rsphh_shift, rsphh_std_term, rsphh_type, rsphh_draft_nbr, rsphh_inv_date "
    strSQL = strSQL & " FROM rs_p_hold_header, sys_prft_ctr, p_altname "
    strSQL = strSQL & " WHERE rsphh_prft_ctr = prft_ctr AND rsphh_vendor = pn_alt "
    
    If chkAllInvoice.value <> vbChecked Then
        strSQL = strSQL & " AND rsphh_status = 'N' "
    End If
    
    If Trim(txtFromPrftCtr) <> "" Then
        strSQL = strSQL & " AND rsphh_prft_ctr >= " & txtFromPrftCtr
        strSQL = strSQL & " AND rsphh_prft_ctr <= " & txtToPrftCtr
    End If
    
    strSQL = strSQL & " ORDER BY rsphh_prft_ctr, rsphh_vendor, rsphh_invoice"
    
    tgmHeader.FillWithSQL t_dbMainDatabase, strSQL, nDB_REMOTE, True
    
    If tgmHeader.RowCount > 0 Then
        t_bDataChanged = True
        fnLoadHeaderData = True
    Else
        MsgBox "No Records found for the given criteria", vbInformation
        fnLoadHeaderData = False
    End If
    
End Function

Private Function fnLoadDetailData(lCurHeaderRow As Long) As Boolean
    Dim strSQL As String
    Dim rsTemp As Recordset
    Dim rsDetail As Recordset
    Dim lCount As Long
    Dim lVendor As Long
    Dim lInvoice As Long
    Dim sItemCode As String
    Dim nPrftctr As Integer
    Dim dtInvDate As Date
    Dim sItemUOM As String
    Dim vArray() As Variant
    Dim dPBCost As Double
    Dim dPBRetail As Double
    Dim i As Long
    
    lVendor = tfnRound(tgmHeader.CellValue(ColHeaderVendor, lCurHeaderRow))
    lInvoice = tfnRound(tgmHeader.CellValue(ColHeaderInvoice, lCurHeaderRow))
    nPrftctr = tfnRound(tgmHeader.CellValue(ColHeaderPrftctr, lCurHeaderRow))
    dtInvDate = CDate(tfnFormatDate(tgmHeader.CellValue(colHdnHeaderInvDate, lCurHeaderRow)))
    
    strSQL = "SELECT rsphd_line_nbr, rsphd_code, rsbi_desc, rsbi_stock_unit, rsphd_qty, rsphd_cost"
    strSQL = strSQL & ", rsphd_retail, rsphd_stock_unit FROM rs_p_hold_detail, rs_b_item "
    strSQL = strSQL & " WHERE rsphd_code = rsbi_code AND rsphd_vendor = rsbi_vendor "
    strSQL = strSQL & " AND rsphd_vendor = " & lVendor
    strSQL = strSQL & " AND rsphd_invoice = " & lInvoice
    
    lCount = fnGetRecord(rsDetail, strSQL, nDB_REMOTE, "fnLoadDetailData")
    g_lNewLineNo = lCount + 1
    
    If lCount <= 0 Then
        fnLoadDetailData = False
        Exit Function
    End If
    
    ReDim vArray(nMaxDetailCol, lCount - 1)
    rsDetail.MoveFirst
    
    i = 0
    
    Do Until rsDetail.EOF
        vArray(ColDetailLine, i) = rsDetail!rsphd_line_nbr
        vArray(ColDetailItemCode, i) = rsDetail!rsphd_code
        sItemCode = CStr(rsDetail!rsphd_code & "")
        vArray(ColDetailItemDesc, i) = rsDetail!rsbi_desc
        vArray(ColDetailQty, i) = tfnRound(rsDetail!rsphd_qty, 3)
        vArray(ColDetailCost, i) = tfnRound(rsDetail!rsphd_cost, 3)
        vArray(ColDetailExtCost, i) = vArray(ColDetailQty, i) * vArray(ColDetailCost, i)
        vArray(ColDetailRetail, i) = tfnRound(rsDetail!rsphd_retail, 3)
        vArray(ColDetailExtRetail, i) = vArray(ColDetailRetail, i) * vArray(ColDetailQty, i)
        vArray(ColDetailUOM, i) = Trim(rsDetail!rsphd_stock_unit & "")
        sItemUOM = Trim$(rsDetail!rsbi_stock_unit & "")
        
        If sItemUOM = CStr(vArray(ColDetailUOM, i)) Then
           
            If fnGetPriceBookCost(sItemCode, lVendor, nPrftctr, dtInvDate, dPBCost, dPBRetail) Then
                vArray(ColDetailPBCost, i) = dPBCost
            Else
                vArray(ColDetailPBCost, i) = vArray(ColDetailCost, i)
            End If
            
        Else
            
            If fnGetPriceFromItemCostMaint(lVendor, sItemCode, CStr(vArray(ColDetailUOM, i)), dtInvDate, dPBCost, dPBRetail) Then
                vArray(ColDetailPBCost, i) = dPBCost
            Else
                vArray(ColDetailPBCost, i) = vArray(ColDetailCost, i)
            End If
            
        End If
        
        If vArray(ColDetailCost, i) <> vArray(ColDetailPBCost, i) Then
            vArray(colHdnDetailFlag, i) = "N"
        Else
            vArray(colHdnDetailFlag, i) = "Y"
        End If
        
        vArray(ColDetailExtPBCost, i) = vArray(ColDetailPBCost, i) * vArray(ColDetailQty, i)
        i = i + 1
        rsDetail.MoveNext
    Loop

    tgmDetail.FillWithArray vArray(), True, True
    
    subCheckButtonStatus
    
    If tgmDetail.RowCount > 0 Then
        fnLoadDetailData = True
    End If
    
End Function

Private Function fnIsAmountBalance(dTotalCost As Double) As Boolean
    Dim fAmount As Double
    Dim dAmount  As Double
    
    dAmount = tgmHeader.CellValue(ColHeaderInvAmount, tgmHeader.GetCurrentRowNumber)
    
    fAmount = dAmount - dTotalCost
    
    If Abs(fAmount) < 0.005 Then
        eLblStatus.ForeColor = &HFFFFFF
        eLblStatus.Caption = "The invoice is balanced"
        eLblStatus.Refresh
        fnIsAmountBalance = True
    Else
        eLblStatus.ForeColor = &HFF
        eLblStatus.Caption = "Out of balance by " & Format(fAmount, "#0.00")
        eLblStatus.Refresh
        fnIsAmountBalance = False
    End If
    
End Function

Private Function fnGetPriceBookCost(sItemCode As String, lVendor As Long, _
            nPrftctr As Integer, dtInvDate As Date, dPBCost As Double, dRetailPrice As Double) As Boolean
    Dim lBook As Long
    Dim lSubBook As Long
    Dim strSQL As String
    Dim rsTemp As Recordset
    Dim lIcLink As Long
    
    Const sFUNC_NAME As String = "fnGetPriceBookCost"
    
    strSQL = "SELECT rsbi_ic_lnk FROM rs_b_item WHERE rsbi_vendor = " & lVendor
    strSQL = strSQL & " AND rsbi_code = " & tfnSQLString(sItemCode)
    
    If fnGetRecord(rsTemp, strSQL, nDB_REMOTE, sFUNC_NAME) > 0 Then
        lIcLink = IIf(IsNull(rsTemp!rsbi_ic_lnk), -1, rsTemp!rsbi_ic_lnk)
    Else
        Exit Function
    End If
    
    strSQL = "SELECT rsbs_book, rsbs_subbook FROM rs_b_store"
    strSQL = strSQL & " WHERE rsbs_vendor = " & lVendor
    strSQL = strSQL & " AND  rsbs_prft_ctr = " & nPrftctr

    If fnGetRecord(rsTemp, strSQL, nDB_REMOTE, sFUNC_NAME) > 0 Then
        lBook = IIf(IsNull(rsTemp!rsbs_book), -1, rsTemp!rsbs_book)
        lSubBook = IIf(IsNull(rsTemp!rsbs_subbook), -1, rsTemp!rsbs_subbook)
    Else
        Exit Function
    End If
    
    strSQL = "SELECT rsbp_cost, rsbp_retail FROM rs_b_price WHERE rsbp_promo = 'Y'"
    strSQL = strSQL & " AND " & tfnDateString(dtInvDate, True) & " BETWEEN rsbp_date and rsbp_ending_date "
    strSQL = strSQL & " AND rsbp_bk_lnk = "
    strSQL = strSQL & " (SELECT rsbb_bk_lnk FROM rs_b_book WHERE rsbb_vendor = " & lVendor
    strSQL = strSQL & " AND rsbb_book = " & lBook
    strSQL = strSQL & " AND rsbb_subbook = " & lSubBook
    strSQL = strSQL & " AND rsbb_ic_lnk = " & lIcLink
    strSQL = strSQL & ")"

    If fnGetRecord(rsTemp, strSQL, nDB_REMOTE, sFUNC_NAME) > 0 Then
        dPBCost = tfnRound(rsTemp!rsbp_cost, 3)
        dRetailPrice = tfnRound(rsTemp!rsbp_retail, 3)
    Else
        strSQL = "SELECT rsbp_cost, rsbp_retail FROM rs_b_price WHERE rsbp_promo = 'N'"
        strSQL = strSQL & " AND " & tfnDateString(dtInvDate, True) & " BETWEEN rsbp_date and rsbp_ending_date "
        strSQL = strSQL & " AND rsbp_bk_lnk = "
        strSQL = strSQL & " (SELECT rsbb_bk_lnk FROM rs_b_book WHERE rsbb_vendor = " & lVendor
        strSQL = strSQL & " AND rsbb_book = " & lBook
        strSQL = strSQL & " AND rsbb_subbook = " & lSubBook
        strSQL = strSQL & " AND rsbb_ic_lnk = " & lIcLink
        strSQL = strSQL & ")"

        If fnGetRecord(rsTemp, strSQL, nDB_REMOTE, sFUNC_NAME) > 0 Then
            dPBCost = tfnRound(rsTemp!rsbp_cost, 3)
            dRetailPrice = tfnRound(rsTemp!rsbp_retail, 3)
        Else
            dPBCost = 0
            dRetailPrice = 0
            fnGetPriceBookCost = False
            Exit Function
        End If
            
    End If
    
    fnGetPriceBookCost = True

End Function

Private Function fnGetPriceFromItemCostMaint(lVendor As Long, sItemCode As String, sUOM As String, dtInvDate As Date, dCost As Double, dRetail As Double) As Boolean
    Dim strSQL As String
    Dim rsTemp As Recordset
    
    strSQL = "SELECT icm1.icm_cost, icm1.icm_retail FROM item_cost_maint icm1 "
    strSQL = strSQL & " WHERE icm1.icm_vendor = " & tfnRound(lVendor)
    strSQL = strSQL & " AND icm1.icm_code = " & tfnSQLString(sItemCode)
    strSQL = strSQL & " AND icm1.icm_uom = " & tfnSQLString(UCase(sUOM))
    strSQL = strSQL & " AND icm1.icm_eff_date = "
    strSQL = strSQL & " (SELECT MAX(icm_eff_date) FROM item_cost_maint "
    strSQL = strSQL & " WHERE icm_vendor = " & tfnRound(lVendor)
    strSQL = strSQL & " AND icm_code = " & tfnSQLString(sItemCode)
    strSQL = strSQL & " AND icm_eff_date <= " & tfnDateString(dtInvDate, True)
    strSQL = strSQL & ")"
    
    If fnGetRecord(rsTemp, strSQL, nDB_REMOTE, "fnGetPriceFromItemCostMaint") <= 0 Then
        fnGetPriceFromItemCostMaint = False
    Else
        dCost = IIf(IsNull(rsTemp!icm_cost), 0, rsTemp!icm_cost)
        dRetail = IIf(IsNull(rsTemp!icm_retail), 0, rsTemp!icm_retail)
        fnGetPriceFromItemCostMaint = True
    End If
        
End Function

Private Function fnDeleteData(lVendor As Long, lInvoice As Long, sProcessed As String) As Boolean
    Dim strSQL As String
    
    'delete in rs_b_hold_detail for this vendor and invoice
    strSQL = "DELETE FROM rs_p_hold_detail WHERE rsphd_vendor = " & lVendor
    strSQL = strSQL & " AND rsphd_invoice = " & lInvoice
    
    If Not fnExecuteSQL(strSQL, nDB_REMOTE, "fnDeleteData") Then
        Exit Function
    End If

    'delete data in rs_b_hold_header for this vendor and invoice
    strSQL = "DELETE FROM rs_p_hold_header WHERE rsphh_vendor = " & lVendor
    strSQL = strSQL & " AND rsphh_invoice = " & lInvoice
    
    If Not fnExecuteSQL(strSQL, nDB_REMOTE, "fnDeleteData") Then
        Exit Function
    End If
    
    'not inserted, you can delete invoice link
    If sProcessed = "N" Then
        strSQL = "DELETE FROM p_nbr WHERE pno_vendor = " & lVendor
        strSQL = strSQL & " AND pno_invoice = " & lInvoice
    
        If Not fnExecuteSQL(strSQL, nDB_REMOTE, "fnDeleteData") Then
            Exit Function
        End If
    
    End If
    
    fnDeleteData = True
End Function

Private Sub subLoadInfoForThisRow(lVendor As Long, sItemCode As String, sUOM As String, sItemDesc As String)
    Dim vData(nMaxDetailCol) As Variant
    Dim lInvoice As Long
    Dim nPrftctr As Integer
    Dim dtInvDate As Date
    Dim lCurHeaderRow As Long
    Dim lCurDetailRow As Long
    Dim dCost As Double
    Dim dRetailPrice As Double
    
    lCurHeaderRow = tgmHeader.GetCurrentRowNumber
    lInvoice = tfnRound(tgmHeader.CellValue(ColHeaderInvoice, lCurHeaderRow))
    nPrftctr = tfnRound(tgmHeader.CellValue(ColHeaderPrftctr, lCurHeaderRow))
    dtInvDate = CDate(tfnFormatDate(tgmHeader.CellValue(colHdnHeaderInvDate, lCurHeaderRow)))
    
    If Not fnGetPriceBookCost(sItemCode, lVendor, nPrftctr, dtInvDate, dCost, dRetailPrice) Then
        '
    End If
    
    lCurDetailRow = tgmDetail.GetCurrentRowNumber
    
    If CStr(tgmDetail.CellValue(ColDetailLine, lCurDetailRow)) = "" Then
        vData(ColDetailLine) = g_lNewLineNo
        g_lNewLineNo = g_lNewLineNo + 1
    Else
        vData(ColDetailLine) = tgmDetail.CellValue(ColDetailLine, lCurDetailRow)
    End If
    
    vData(ColDetailItemCode) = sItemCode
    vData(ColDetailItemDesc) = sItemDesc
    vData(ColDetailQty) = 0#
    vData(ColDetailCost) = tfnRound(dCost, 3)
    vData(ColDetailExtCost) = 0#
    vData(ColDetailPBCost) = tfnRound(dCost, 3)
    vData(ColDetailExtPBCost) = 0#
    vData(ColDetailRetail) = tfnRound(dRetailPrice, 3)
    vData(ColDetailExtRetail) = 0#
    vData(ColDetailUOM) = sUOM
    vData(colHdnDetailFlag) = "Y"

    tgmDetail.SetCurrentRowData vData

End Sub

Private Sub subLoadDataNoException(bYesNo As Boolean)
    Dim vData() As Variant
    Dim vAllData() As Variant
    Dim nCol As Integer
    Dim lFilter As Long
    Dim lNoFilter As Long
    Dim i As Long
    Dim j As Integer
    
    If Not bYesNo Then
        
        For i = 0 To tgmDetail.RowCount - 1
            tgmDetail.GetRow vData, nCol, i
            
            If vData(colHdnDetailFlag) = "N" Then
                ReDim Preserve vAllData(nMaxDetailCol, lFilter)
                
                For j = 0 To nMaxDetailCol
                    vAllData(j, lFilter) = vData(j)
                Next
                
                lFilter = lFilter + 1
            Else
                ReDim Preserve g_vHoldDetailData(nMaxDetailCol, lNoFilter)
                
                For j = 0 To nMaxDetailCol
                    g_vHoldDetailData(j, lNoFilter) = vData(j)
                Next
                
                lNoFilter = lNoFilter + 1
            End If
                        
        Next
        
        tgmDetail.FillWithArray vAllData(), True, True
        g_bExceptionsClicked = True
        cmdExceptions.Caption = "&All Data"
    Else
            
        For i = 0 To tgmDetail.RowCount - 1
            tgmDetail.GetRow vData, nCol, i
            subReOrderDetailData vData
        Next
        
        tgmDetail.FillWithArray g_vHoldDetailData(), True, True
        subResetExceptions
        
    End If
    
End Sub

Private Sub subReOrderDetailData(vData As Variant)
    Dim i As Long
    Dim j As Long
    Dim m As Long
    Dim bInsert As Boolean
    
    'This array may be empty first time
    If CStr(g_vHoldDetailData(0, 0)) <> "" Then
        ReDim Preserve g_vHoldDetailData(nMaxDetailCol, UBound(g_vHoldDetailData, 2) + 1)
    End If
    
    For i = 0 To UBound(g_vHoldDetailData, 2) - 1
        
        If g_vHoldDetailData(0, i) > vData(0) Then
            
            For j = UBound(g_vHoldDetailData, 2) To i + 1 Step -1
                
                For m = 0 To nMaxDetailCol
                    g_vHoldDetailData(m, j) = g_vHoldDetailData(m, j - 1)
                Next
                    
            Next
            
            For m = 0 To nMaxDetailCol
                g_vHoldDetailData(m, i) = vData(m)
            Next
                    
            bInsert = True
            Exit For
        End If
        
    Next
    
    If Not bInsert Then
        
        For m = 0 To nMaxDetailCol
            g_vHoldDetailData(m, UBound(g_vHoldDetailData, 2)) = vData(m)
        Next
    
    End If
    
End Sub

Private Function fnDeleteDetailData(lVendor As Long, lInvoice As Long, lLineNo As Long) As Boolean
    Dim strSQL As String
    
    'update first because we may add more than one line, if you delete and repopulate data, all lines will gone.
    If Not fnUpdateDetail() Then
        Exit Function
    End If
    
    strSQL = "DELETE FROM rs_p_hold_detail WHERE rsphd_vendor = " & lVendor
    strSQL = strSQL & " AND rsphd_invoice = " & lInvoice
    strSQL = strSQL & " AND rsphd_line_nbr = " & lLineNo
    
    If fnExecuteSQL(strSQL, nDB_REMOTE, "fnDeleteDetailData") Then
        
        strSQL = "UPDATE rs_p_hold_detail SET rsphd_line_nbr = rsphd_line_nbr -1 "
        strSQL = strSQL & " WHERE rsphd_vendor = " & lVendor
        strSQL = strSQL & " AND rsphd_invoice = " & lInvoice
        strSQL = strSQL & " AND rsphd_line_nbr > " & lLineNo
        
        If fnExecuteSQL(strSQL, nDB_REMOTE, "fnDeleteDetailData") Then
            fnDeleteDetailData = True
        End If
        
    End If
    
End Function

Private Sub subCheckButtonStatus()
    Dim dTotalCost As Double
    Dim sStatus As String
    
    sStatus = tgmHeader.CellValue(ColHeaderStatus, tgmHeader.GetCurrentRowNumber)

    dTotalCost = fnGetToTalCost

    If fnIsAmountBalance(dTotalCost) Then
        
        If tgmDetail.ValidData And sStatus = "N" Then
            subEnableDetailOKBtn True
        Else
            subEnableDetailOKBtn False
            subEnableUpdateBtn False
        End If
        
    Else
        subEnableDetailOKBtn False
        subEnableUpdateBtn False
    End If
    
End Sub

Private Function fnGetToTalCost() As Double
    Dim i As Long
    Dim dTotalCost As Double
    
    For i = 0 To tgmDetail.RowCount - 1
        dTotalCost = dTotalCost + tfnRound(tgmDetail.CellValue(ColDetailExtCost, i), 3)
    Next
    
    fnGetToTalCost = dTotalCost
    
End Function

Private Function fnUpdateDetail() As Boolean
    Dim strSQL As String
    Dim lVendor As Long
    Dim lInvoice As Long
    Dim lCurHeaderRow As Long
    Dim i As Long
    Const sFUNC_NAME As String = "fnUpdateDetail"
    
    lCurHeaderRow = tgmHeader.GetCurrentRowNumber
    lVendor = tfnRound(tgmHeader.CellValue(ColHeaderVendor, lCurHeaderRow))
    lInvoice = tfnRound(tgmHeader.CellValue(ColHeaderInvoice, lCurHeaderRow))
    
    strSQL = "DELETE FROM rs_p_hold_detail WHERE rsphd_vendor = " & lVendor
    strSQL = strSQL & " AND rsphd_invoice = " & lInvoice
    
    If fnExecuteSQL(strSQL, nDB_REMOTE, sFUNC_NAME) Then
        
        For i = 0 To tgmDetail.RowCount - 1
            strSQL = "INSERT INTO rs_p_hold_detail (rsphd_vendor, rsphd_invoice, rsphd_line_nbr, "
            strSQL = strSQL & " rsphd_type, rsphd_code, rsphd_qty, rsphd_stock_unit, rsphd_cost, "
            strSQL = strSQL & " rsphd_retail) VALUES(" & lVendor & "," & lInvoice & "," & i + 1 & ","
            strSQL = strSQL & "'P', " & tfnSQLString(tgmDetail.CellValue(ColDetailItemCode, i)) & ","
            strSQL = strSQL & tfnRound(tgmDetail.CellValue(ColDetailQty, i), 3) & ","
            strSQL = strSQL & tfnSQLString(tgmDetail.CellValue(ColDetailUOM, i)) & ","
            strSQL = strSQL & tfnRound(tgmDetail.CellValue(ColDetailCost, i), 3) & ","
            strSQL = strSQL & tfnRound(tgmDetail.CellValue(ColDetailRetail, i), 3) & ")"
            
            If Not fnExecuteSQL(strSQL, nDB_REMOTE, sFUNC_NAME) Then
                Exit Function
            End If
            
        Next
        
        fnUpdateDetail = True
    End If
    
    'right now invoice amount can be changed, so update header
    strSQL = "UPDATE rs_p_hold_header SET rsphh_invoice_amt = " & tfnRound(tgmHeader.CellValue(ColHeaderInvAmount, lCurHeaderRow), 2)
    strSQL = strSQL & " WHERE rsphh_vendor = " & lVendor & " AND rsphh_invoice = " & lInvoice
    
    If Not fnExecuteSQL(strSQL, nDB_REMOTE, sFUNC_NAME) Then
        Exit Function
    End If
    
End Function

Private Sub subSetTabDetailCaption(lCurHeaderRow As Long)
    lblPrftCtr.Caption = tgmHeader.CellValue(ColHeaderPrftctr, lCurHeaderRow)
    lblPrftCtrName.Caption = tgmHeader.CellValue(ColHeaderPrftDesc, lCurHeaderRow)
    lblReportDate.Caption = tgmHeader.CellValue(ColHeaderRptDate, lCurHeaderRow)
    lblShift.Caption = tgmHeader.CellValue(ColHdnHeaderShift, lCurHeaderRow)
    lblVendor.Caption = tgmHeader.CellValue(ColHeaderVendor, lCurHeaderRow)
    lblVendorName.Caption = tgmHeader.CellValue(ColHeaderVendorName, lCurHeaderRow)
    lblInvoice.Caption = tgmHeader.CellValue(ColHeaderInvoice, lCurHeaderRow)
    lblAmount.Caption = tfnRound(tgmHeader.CellValue(ColHeaderInvAmount, lCurHeaderRow), 2)
End Sub

Private Sub subUnlockHeader(bAlreadyInsert As Boolean)
    
    If g_nHoldSelectedPrftCtr > 0 Then
    
        If g_lHoldShiftLink > 0 Then
            fnUnlockShiftReport g_nHoldSelectedPrftCtr, g_lHoldShiftLink, bAlreadyInsert
            g_lHoldShiftLink = 0
        End If
        
        g_nHoldSelectedPrftCtr = 0
        g_dtHoldRptDate = #12/31/1899#
        g_lHoldShiftNum = 0
        g_lHoldSelectedInvoice = 0
        g_lHoldSelectedVendor = 0
    End If
    
End Sub

Private Function fnlockHeader(lVendor As Long, lInvoice As Long, nPrftctr As Integer, dtRptDate As Date, lShiftNum As Long)
    Dim sMsg As String
    Dim sUser As String
    
    fnlockHeader = True
    
    'If nPrftctr <> g_nHoldSelectedPrftCtr Or dtRptDate <> g_dtHoldRptDate Or lShiftNum <> g_lHoldShiftNum Then
    If lVendor <> g_lHoldSelectedVendor Or lInvoice <> g_lHoldSelectedInvoice Then
        g_lHoldShiftLink = fnLockShiftReport(nPrftctr, lShiftNum, dtRptDate, sUser)
    End If
    
    If g_lHoldShiftLink < 0 Then
        
        If g_lHoldShiftLink = -3 Then
            sMsg = "Purchase entry no longer allowed on this shift because sales have been entered"
        ElseIf g_lHoldShiftLink = -2 Then
            sMsg = "Date and shift not expected"
        ElseIf g_lHoldShiftLink = -1 Then
            
            If sUser <> "" Then
                sMsg = "Profit center " & CStr(nPrftctr) & " is being used by " & sUser
            Else
                sMsg = "Profit center " & CStr(nPrftctr) & " is locked."
            End If
            
        End If
        
        MsgBox sMsg, vbInformation
        fnlockHeader = False
    End If
        
End Function

Private Sub subResetExceptions()
    ReDim g_vHoldDetailData(nMaxDetailCol, 0) As Variant
    g_bExceptionsClicked = False
    cmdExceptions.Caption = "Exceptio&ns"
End Sub
'I think no data already in the the rs_shifthold for this shift link.
Private Function fnProcessData() As Boolean
    Dim lCurHeaderRow As Long
    Dim dtRptDate As Date
    Dim sPayType As String
    Dim sMsg As String
    Dim lGlAccount As Long
    Dim bOk As Boolean
    
    lCurHeaderRow = tgmHeader.GetCurrentRowNumber
    
    dtRptDate = tfnFormatDate(tgmHeader.CellValue(ColHeaderRptDate, lCurHeaderRow))
    sPayType = tgmHeader.CellValue(ColHdnHeaderType, lCurHeaderRow)
    
    subShowHourGlass
    
    If g_lHoldShiftLink <= 0 Then
        MsgBox "Invalid Shift link", vbInformation
        GoTo EXITHERE
    End If
    
    tfnSetStatusBarCorrect "Processing data, please wait..."
    sMsg = fnCheckGLStatus(dtRptDate)
    
    If sMsg <> "" Then
        MsgBox sMsg, vbInformation
        GoTo EXITHERE
    End If
    
    lGlAccount = fnGetPayGLAccount(sPayType)
    
    If lGlAccount <= 0 Then
        MsgBox "GL account is not set up.", vbInformation
        GoTo EXITHERE
    End If
    
    'bOk = fnCleanShiftHoldData
    'do here 04/05...........
    bOk = True
    
    If bOk Then
        bOk = fnInsertVendorInvoice()
    End If
    
    If bOk Then
        bOk = fnInsertPriceBookInvDetail()
    End If
    
    If bOk Then
        bOk = fnInsertCategoryAndGLInfo(lGlAccount)
    End If
    
    If Not bOk Then
        bOk = fnCleanShiftHoldData()
    Else
        bOk = fnUpdateInvHeader
    End If
    
    fnProcessData = bOk
    subHideHourGlass
    Exit Function
    
EXITHERE:
    fnProcessData = False
    subHideHourGlass
End Function

Private Function fnCleanShiftHoldData() As Boolean
    Dim strSQL As String
    Const FUNC_NAME As String = "fnCleanShiftHoldData"
    
    If MsgBox("There some something wrong during inserting, do you want to delete all " _
        & " information for this shift link " & g_lHoldShiftLink & "?", vbQuestion + vbYesNo) = vbNo Then
        fnCleanShiftHoldData = True
        Exit Function
    End If
    
    strSQL = "DELETE FROM rs_shifthold WHERE rssh_shl = " & g_lHoldShiftLink
    strSQL = strSQL & " AND rssh_type IN ('V', 'G', 'C', 'Y') "
    
    
    If fnExecuteSQL(strSQL, nDB_REMOTE, FUNC_NAME) Then
        fnCleanShiftHoldData = True
    End If
    
End Function

Private Function fnInsertVendorInvoice() As Boolean
    Dim strSQL As String
    Dim lVendor As Long
    Dim sTerm As String
    Dim lInvoice As Long
    Dim dAmount As Double
    Dim sPayType As String
    Dim lDraftNbr As Long
    Dim lCurHeaderRow As Long
    Dim nPayType As Integer
    Dim dtInvDate As Date
    Dim lMaxSeqNum As Long
    Const FUNC_NAME As String = "fnInsertVendorInvoice"
    
    lCurHeaderRow = tgmHeader.GetCurrentRowNumber
    lVendor = tfnRound(tgmHeader.CellValue(ColHeaderVendor, lCurHeaderRow))
    lInvoice = tfnRound(tgmHeader.CellValue(ColHeaderInvoice, lCurHeaderRow))
    sTerm = tgmHeader.CellValue(ColHdnHeaderTerm, lCurHeaderRow)
    dAmount = tfnRound(tgmHeader.CellValue(ColHeaderInvAmount, lCurHeaderRow), 2)
    sPayType = tgmHeader.CellValue(ColHdnHeaderType, lCurHeaderRow)
    lDraftNbr = tfnRound(tgmHeader.CellValue(ColHdnHeaderDraft, lCurHeaderRow))
    dtInvDate = tfnFormatDate(tgmHeader.CellValue(colHdnHeaderInvDate, lCurHeaderRow))
    
    If sPayType = "M" Or sPayType = "D" Then
    
        If Not fnValidDraftNbr(lDraftNbr) Then
            Exit Function
        End If
    
    End If
    
    Select Case sPayType
        Case "C":
            nPayType = 1
        Case "D"
            nPayType = 2
        Case "P"
            nPayType = 3
        Case "T"
            nPayType = 4
        Case "M"
            nPayType = 5
    End Select
    
    lMaxSeqNum = fnGetMaxSeqNum("V")
    
    strSQL = "INSERT INTO rs_shifthold(rssh_shl, rssh_type, rssh_code, "
    strSQL = strSQL & " rssh_1, rssh_2, rssh_3, rssh_4, rssh_5, rssh_6, rssh_7, rssh_8,"
    strSQL = strSQL & " rssh_seq)VALUES(" & g_lHoldShiftLink & ",'V', " & tfnSQLString(sTerm) & ","
    strSQL = strSQL & lVendor & "," & lInvoice & "," & dAmount & "," & nPayType & ","
    strSQL = strSQL & IIf((sPayType = "M" Or sPayType = "D"), lDraftNbr, "NULL") & ","
    strSQL = strSQL & fnDateToSQLIntegerString(dtInvDate) & ", NULL, NULL," & lMaxSeqNum + 1 & ")"
    
    If Not fnExecuteSQL(strSQL, nDB_REMOTE, FUNC_NAME) Then
        Exit Function
    End If
    
    fnInsertVendorInvoice = True
End Function

Private Function fnValidDraftNbr(lDraftNbr As Long) As Boolean
    Dim strSQL As String
    Dim rsTemp As Recordset
    Const FUNC_NAME As String = "fnValidDraftNbr"
    
    'Check p_draft
    strSQL = "SELECT pdr_draft_nbr FROM p_draft WHERE pdr_draft_nbr = " & lDraftNbr
    
    If fnGetRecord(rsTemp, strSQL, nDB_REMOTE, FUNC_NAME) > 0 Then
        MsgBox "Draft number has been used by other file", vbInformation
        Exit Function
    End If

    'Check rs_shifthold
    strSQL = "SELECT rssh_5 FROM rs_shifthold WHERE rssh_type = 'V' AND rssh_4 = 2 AND rssh_5 = " & lDraftNbr
    
    If fnGetRecord(rsTemp, strSQL, nDB_REMOTE, FUNC_NAME) > 0 Then
        MsgBox "Draft number has been used by holding file", vbInformation
        Exit Function
    End If
    
    fnValidDraftNbr = True

End Function


Private Function fnInsertPriceBookInvDetail() As Boolean
    Dim strSQL As String
    Dim lVendor As Long
    Dim lInvoice As Long
    Dim nPrftctr As Integer
    Dim dtInvDate As Date
    Dim lCurHeaderRow As Long
    Dim sItemCode As String
    Dim dCost As Double
    Dim dRetail As Double
    Dim dExtCost As Double
    Dim dExtRetail As Double
    Dim dQty As Double
    Dim lPriceBookLnk As Long
    Dim lMaxSeq As Long
    Dim i As Long
    
    Const FUNC_NAME As String = "fnInsertPriceBookInvDetail"
    
    lCurHeaderRow = tgmHeader.GetCurrentRowNumber
    lVendor = tfnRound(tgmHeader.CellValue(ColHeaderVendor, lCurHeaderRow))
    nPrftctr = tfnRound(tgmHeader.CellValue(ColHeaderPrftctr, lCurHeaderRow))
    lInvoice = tfnRound(tgmHeader.CellValue(ColHeaderInvoice, lCurHeaderRow))
    dtInvDate = tfnFormatDate(tgmHeader.CellValue(colHdnHeaderInvDate, lCurHeaderRow))
    lMaxSeq = fnGetMaxSeqNum("Y")
    
    For i = 0 To tgmDetail.RowCount - 1
        sItemCode = tgmDetail.CellValue(ColDetailItemCode, i)
        dCost = tfnRound(tgmDetail.CellValue(ColDetailCost, i), 2)
        dRetail = tfnRound(tgmDetail.CellValue(ColDetailRetail, i), 2)
        dExtCost = tfnRound(tgmDetail.CellValue(ColDetailExtCost, i), 2)
        dExtRetail = tfnRound(tgmDetail.CellValue(ColDetailExtRetail, i), 2)
        dQty = tfnRound(tgmDetail.CellValue(ColDetailQty, i), 3)
        lPriceBookLnk = fnGetPriceBookLink(sItemCode, lVendor, nPrftctr)
        lMaxSeq = lMaxSeq + 1
        
        If lPriceBookLnk = 0 Then
            MsgBox "No Price book link found", vbInformation
            Exit Function
        End If
        
        'making price book link negative when the cost is an exception
        If tgmDetail.CellValue(colHdnDetailFlag, i) = "N" Then
            lPriceBookLnk = -1 * lPriceBookLnk
        End If
        
        strSQL = "INSERT INTO rs_shifthold(rssh_shl, rssh_type, rssh_code, "
        strSQL = strSQL & " rssh_1, rssh_2, rssh_3, rssh_4, rssh_5, rssh_6, rssh_7, rssh_8,"
        strSQL = strSQL & " rssh_seq)VALUES(" & g_lHoldShiftLink & ",'Y', " & tfnSQLString(sItemCode) & ","
        strSQL = strSQL & lVendor & "," & lInvoice & "," & lPriceBookLnk & "," & dExtCost & ","
        strSQL = strSQL & dExtRetail & "," & dCost & "," & dRetail & "," & dQty & "," & lMaxSeq & ")"
        
        If Not fnExecuteSQL(strSQL, nDB_REMOTE, FUNC_NAME) Then
            Exit Function
        End If
        
    Next
    
    fnInsertPriceBookInvDetail = True
End Function

Private Function fnGetPriceBookLink(sItemCode As String, lVendor As Long, nPrftctr As Integer) As Long
    Dim lBook As Long
    Dim lSubBook As Long
    Dim strSQL As String
    Dim rsTemp As Recordset
    Dim lIcLink As Long
    
    Const sFUNC_NAME As String = "fnGetPriceBookLink"
    
    strSQL = "SELECT rsbi_ic_lnk FROM rs_b_item WHERE rsbi_vendor = " & lVendor
    strSQL = strSQL & " AND rsbi_code = " & tfnSQLString(sItemCode)
    
    If fnGetRecord(rsTemp, strSQL, nDB_REMOTE, sFUNC_NAME) > 0 Then
        lIcLink = tfnRound(rsTemp!rsbi_ic_lnk)
    Else
        Exit Function
    End If
    
    strSQL = "SELECT rsbb_bk_lnk FROM rs_b_store, rs_b_book"
    strSQL = strSQL & " WHERE rsbs_vendor = " & lVendor
    strSQL = strSQL & " AND  rsbs_prft_ctr = " & nPrftctr
    strSQL = strSQL & " AND rsbb_book = rsbs_book AND rsbb_subbook = rsbs_subbook "
    strSQL = strSQL & " AND rsbb_vendor = rsbs_vendor AND rsbb_ic_lnk = " & lIcLink

    If fnGetRecord(rsTemp, strSQL, nDB_REMOTE, sFUNC_NAME) > 0 Then
        fnGetPriceBookLink = tfnRound(rsTemp!rsbb_bk_lnk)
    Else
        Exit Function
    End If
    
End Function

Private Function fnInsertCategoryAndGLInfo(lGlAccount As Long) As Boolean
    Dim strSQL As String
    Dim rsTemp As Recordset
    Dim lVendor As Long
    Dim lCurHeaderRow As Long
    Dim sItemCode As String
    Dim dExtCost As Double
    Dim dExtRetail As Double
    Dim sCategory As String
    Dim sCat() As String
    Dim dCatTotalCost() As Double
    Dim dCatTotalRetail() As Double
    Dim sCostAccount As String
    Dim dTotalCredit As Double
    Dim dTotalDebit As Double
    Dim lMaxCatSeq As Long
    Dim lCurCatSeq As Long
    Dim lMaxGLSeq As Long
    Dim bUpdate As Boolean
    Dim i As Long
    Const FUNC_NAME As String = "fnInsertCategoryAndGLInfo"
    
    lCurHeaderRow = tgmHeader.GetCurrentRowNumber
    lVendor = tfnRound(tgmHeader.CellValue(ColHeaderVendor, lCurHeaderRow))
    dTotalCredit = tfnRound(tgmHeader.CellValue(ColHeaderInvAmount, lCurHeaderRow), 2)
    
    ReDim sCat(0) As String
    ReDim dCatTotalCost(0) As Double
    ReDim dCatTotalRetail(0) As Double
    
    For i = 0 To tgmDetail.RowCount - 1
        sItemCode = tgmDetail.CellValue(ColDetailItemCode, i)
        dExtCost = tfnRound(tgmDetail.CellValue(ColDetailExtCost, i), 3)
        dExtRetail = tfnRound(tgmDetail.CellValue(ColDetailExtRetail, i), 3)
        sCategory = fnGetItemCategory(lVendor, sItemCode)
        subFiterCategory sCategory, dExtCost, dExtRetail, sCat(), dCatTotalCost(), dCatTotalRetail(), i
    Next
    
    If sCat(0) <> "" Then
        
        lMaxCatSeq = fnGetMaxSeqNum("C")
        lMaxGLSeq = fnGetMaxSeqNum("G")

        For i = 0 To UBound(sCat)
            strSQL = "SELECT rssh_seq FROM rs_shifthold WHERE rssh_shl = " & g_lHoldShiftLink
            strSQL = strSQL & " AND rssh_type = 'C' AND rssh_code = " & tfnSQLString(sCat(i))
            
            If fnGetRecord(rsTemp, strSQL, nDB_REMOTE, FUNC_NAME) < 0 Then
                Exit Function
            ElseIf rsTemp.RecordCount = 0 Then
                
                strSQL = "INSERT INTO rs_shifthold(rssh_shl, rssh_type, rssh_code, "
                strSQL = strSQL & " rssh_1, rssh_2, rssh_3, rssh_4, rssh_5, rssh_6, rssh_7, rssh_8,"
                strSQL = strSQL & " rssh_seq)VALUES(" & g_lHoldShiftLink & ",'C', " & tfnSQLString(sCat(i)) & ","
                strSQL = strSQL & "0, 0, 0," & tfnRound(dCatTotalCost(i), 2) & "," & tfnRound(dCatTotalRetail(i), 2) & ","
                strSQL = strSQL & "0, 0, NULL, " & lMaxCatSeq + 1 & ")"
            
                If Not fnExecuteSQL(strSQL, nDB_REMOTE, FUNC_NAME) Then
                    Exit Function
                End If
            
                sCostAccount = fnGetCostGLAccount(sCat(i))
            
                If sCostAccount = "" Then
                    Exit Function
                ElseIf Not fnInsertGlInfo(sCostAccount, dCatTotalCost(i), lMaxGLSeq, False) Then
                    Exit Function
                End If
               
                lMaxCatSeq = lMaxCatSeq + 1
                lMaxGLSeq = lMaxGLSeq + 1
                dTotalDebit = dTotalDebit + tfnRound(dCatTotalCost(i), 2)
                
            Else
                lCurCatSeq = tfnRound(rsTemp!rssh_seq)
                
                If Not fnUpdateCatory(dCatTotalCost(i), dCatTotalRetail(i), lCurCatSeq) Then
                    Exit Function
                End If
                
                sCostAccount = fnGetCostGLAccount(sCat(i))
            
                If sCostAccount = "" Then
                    Exit Function
                ElseIf Not fnUpdateGlInfo(sCostAccount, dCatTotalCost(i), lMaxGLSeq, False) Then
                    Exit Function
                End If
               
                bUpdate = True
                dTotalDebit = dTotalDebit + tfnRound(dCatTotalCost(i), 2)
             
            End If
            
        Next
        
        If dTotalDebit <> dTotalCredit Then
            'MsgBox " The account is not balance", vbInformation
            Exit Function
        Else
            If bUpdate Then
                
                If Not fnUpdateGlInfo(CStr(lGlAccount), dTotalCredit, lMaxGLSeq, True) Then
                    Exit Function
                End If
                
            Else
            
                If Not fnInsertGlInfo(CStr(lGlAccount), dTotalCredit, UBound(sCat) + 1, True) Then
                    Exit Function
                End If
            
            End If
            
            
        End If
            
        fnInsertCategoryAndGLInfo = True
    Else
        fnInsertCategoryAndGLInfo = False
    End If
        
End Function

Private Function fnUpdateCatory(ByVal dCatTotalCost As Double, ByVal dCatTotalRetail As Double, ByVal lSeq As Long) As Boolean
    Dim dOrgCost As Double
    Dim dOrgRetail As Double
    Dim strSQL As String
    Dim rsTemp As Recordset
    Const FUNC_NAME As String = "fnUpdateCatory"
    
    strSQL = "SELECT rssh_4, rssh_5 FROM rs_shifthold WHERE rssh_shl = " & g_lHoldShiftLink
    strSQL = strSQL & " AND rssh_type = 'C' and rssh_seq = " & lSeq
    
    If fnGetRecord(rsTemp, strSQL, nDB_REMOTE, FUNC_NAME) <= 0 Then
        Exit Function
    Else
        dOrgCost = tfnRound(rsTemp!rssh_4, 2)
        dOrgRetail = tfnRound(rsTemp!rssh_5, 2)
    End If
    
    dCatTotalCost = dCatTotalCost + dOrgCost
    dCatTotalRetail = dCatTotalRetail + dOrgRetail

    strSQL = "UPDATE rs_shifthold set rssh_4 = " & tfnRound(dCatTotalCost, 2)
    strSQL = strSQL & " ,rssh_5 = " & tfnRound(dCatTotalRetail, 2)
    strSQL = strSQL & "  WHERE rssh_shl = " & g_lHoldShiftLink
    strSQL = strSQL & " AND rssh_type = 'C' AND rssh_seq = " & lSeq
    
    If Not fnExecuteSQL(strSQL, nDB_REMOTE, FUNC_NAME) Then
        Exit Function
    End If

    fnUpdateCatory = True
End Function

Private Function fnGetItemCategory(lVendor As Long, sItemCode As String) As String
    Dim strSQL As String
    Dim rsTemp As Recordset
    
    strSQL = "SELECT rsbi_cat FROM rs_b_item WHERE rsbi_vendor = " & lVendor
    strSQL = strSQL & " AND rsbi_code = " & tfnSQLString(sItemCode)
    
    If fnGetRecord(rsTemp, strSQL, nDB_REMOTE, "fnGetItemCategory") > 0 Then
        fnGetItemCategory = Trim(rsTemp!rsbi_cat & "")
    End If
    
End Function

Private Sub subFiterCategory(sCategory As String, dExtCost As Double, dExtRetail As Double, _
                            sCat() As String, dCatTotalCost() As Double, dCatTotalRetail() As Double, lTimes As Long)
    Dim i As Long
    Dim bFound As Boolean
    
    bFound = False
    
    If sCategory <> "" Then
        'the first time
        If lTimes = 0 Then
            sCat(0) = sCategory
            dCatTotalCost(0) = dExtCost
            dCatTotalRetail(0) = dExtRetail
        Else
                
            For i = 0 To UBound(sCat)
             
                If sCategory = sCat(i) Then
                    dCatTotalCost(i) = dCatTotalCost(i) + dExtCost
                    dCatTotalRetail(i) = dCatTotalRetail(i) + dExtRetail
                    bFound = True
                End If
                
            Next
            
            If Not bFound Then
                ReDim Preserve sCat(UBound(sCat) + 1)
                ReDim Preserve dCatTotalCost(UBound(dCatTotalCost) + 1)
                ReDim Preserve dCatTotalRetail(UBound(dCatTotalRetail) + 1)
                sCat(UBound(sCat)) = sCategory
                dCatTotalCost(UBound(dCatTotalRetail)) = dExtCost
                dCatTotalRetail(UBound(dCatTotalRetail)) = dExtRetail
            End If
    
        
        End If
        
    End If
    
End Sub

Private Function fnGetCostGLAccount(sCat As String) As String
    Dim strSQL As String
    Dim rsTemp As Recordset
    
    strSQL = "SELECT rsct_gl_cogs FROM rs_cat WHERE rsct_catagory = " & tfnSQLString(sCat)
    
    If fnGetRecord(rsTemp, strSQL, nDB_REMOTE, "fnGetCostGLAccount") > 0 Then
        fnGetCostGLAccount = rsTemp!rsct_gl_cogs & ""
    End If
End Function

Private Function fnInsertGlInfo(sCostAccount As String, dExtCost As Double, lSeq As Long, bTotal As Boolean) As Boolean
    Dim strSQL As String
    Dim strHeader As String
    
    If lSeq Mod 2 = 0 Then
        strHeader = "INSERT INTO rs_shifthold(rssh_shl, rssh_type, rssh_code, "
        strHeader = strHeader & " rssh_1, rssh_2, rssh_3, rssh_4, rssh_5, rssh_6, rssh_7, rssh_8,"
        strHeader = strHeader & " rssh_seq)VALUES"
            
        If bTotal Then
            strSQL = strHeader & "(" & g_lHoldShiftLink & ",'G', NULL, " & tfnRound(sCostAccount) & ","
            strSQL = strSQL & " 0, " & tfnRound(dExtCost, 2) & ", 0, 0, 0, NULL, NULL, " & lSeq \ 2 + 1 & ")"
        Else
            strSQL = strHeader & "(" & g_lHoldShiftLink & ",'G', NULL, " & tfnRound(sCostAccount) & ","
            strSQL = strSQL & tfnRound(dExtCost, 2) & ", 0, 0, 0, 0, NULL, NULL, " & lSeq \ 2 + 1 & ")"
        End If
    Else
    
        If bTotal Then
            strSQL = "UPDATE rs_shifthold SET(rssh_4, rssh_6) = (" & tfnRound(sCostAccount) & "," & tfnRound(dExtCost, 2)
            strSQL = strSQL & ") WHERE rssh_shl = " & g_lHoldShiftLink & " AND rssh_type = 'G' AND rssh_seq = " & lSeq \ 2 + 1
        Else
            strSQL = "UPDATE rs_shifthold SET(rssh_4, rssh_5) = (" & tfnRound(sCostAccount) & "," & tfnRound(dExtCost, 2)
            strSQL = strSQL & ") WHERE rssh_shl = " & g_lHoldShiftLink & " AND rssh_type = 'G' AND rssh_seq = " & lSeq \ 2 + 1
        End If
    
    End If
    
    If fnExecuteSQL(strSQL, nDB_REMOTE, "fnInsertGlInfo") Then
        fnInsertGlInfo = True
    End If

End Function

Private Function fnUpdateGlInfo(ByVal sCostAccount As String, ByVal dExtCost As Double, _
                                ByVal lMaxSeq As Long, ByVal bTotal As Boolean) As Boolean
    Dim strSQL As String
    Dim strSQL2 As String
    Dim rsTemp As Recordset
    Dim lCurGLSeq As Long
    Dim dOrgCredit1 As Double
    Dim dorgDebit1 As Double
    Dim dOrgCredit2 As Double
    Dim dOrgDebit2 As Double
    
    strSQL = "SELECT rssh_1, rssh_2, rssh_3, rssh_4, rssh_5, rssh_6, rssh_seq FROM rs_shifthold  "
    strSQL = strSQL & " WHERE rssh_shl = " & g_lHoldShiftLink
    strSQL = strSQL & " AND rssh_type = 'G' AND rssh_1 = " & tfnSQLString(sCostAccount)
    
    If fnGetRecord(rsTemp, strSQL, nDB_REMOTE, "fnUpdateGlInfo") > 0 Then
        dorgDebit1 = tfnRound(rsTemp!rssh_2, 2)
        dOrgCredit1 = tfnRound(rsTemp!rssh_3, 2)
        lCurGLSeq = tfnRound(rsTemp!rssh_seq)
        
        If Not bTotal Then
            dExtCost = dExtCost + dorgDebit1
            strSQL2 = "UPDATE rs_shifthold set rssh_2 = " & tfnRound(dExtCost, 2)
        Else
            dExtCost = dExtCost + dOrgCredit1
            strSQL2 = "UPDATE rs_shifthold set rssh_3 =  " & tfnRound(dExtCost, 2)
        End If
        
    Else
        
        strSQL = "SELECT rssh_1, rssh_2, rssh_3, rssh_4, rssh_5, rssh_6, rssh_seq FROM rs_shifthold  "
        strSQL = strSQL & " WHERE rssh_shl = " & g_lHoldShiftLink
        strSQL = strSQL & " AND rssh_type = 'G' AND rssh_4 = " & tfnSQLString(sCostAccount)
    
        If fnGetRecord(rsTemp, strSQL, nDB_REMOTE, "fnUpdateGlInfo") > 0 Then
            dOrgDebit2 = tfnRound(rsTemp!rssh_5, 2)
            dOrgCredit2 = tfnRound(rsTemp!rssh_6, 2)
            lCurGLSeq = tfnRound(rsTemp!rssh_seq)
            
            If Not bTotal Then
                dExtCost = dExtCost + dOrgDebit2
                strSQL2 = "UPDATE rs_shifthold set rssh_5 = " & tfnRound(dExtCost, 2)
            Else
                dExtCost = dExtCost + dOrgCredit2
                strSQL2 = "UPDATE rs_shifthold set rssh_6 =  " & tfnRound(dExtCost, 2)
            End If
        
        Else
            
            If fnInsertGlInfo(sCostAccount, dExtCost, lMaxSeq, True) Then
               fnUpdateGlInfo = True
            End If
            
            Exit Function
    
        End If

    End If
    
            
    strSQL2 = strSQL2 & " WHERE rssh_shl = " & g_lHoldShiftLink
    strSQL2 = strSQL2 & " AND rssh_type = 'G' AND rssh_seq = " & lCurGLSeq
    
    If Not fnExecuteSQL(strSQL2, nDB_REMOTE, "fnUpdateGlInfo") Then
        Exit Function
    End If

    fnUpdateGlInfo = True
End Function

Private Function fnUpdateInvHeader() As Boolean
    Dim strSQL As String
    Dim lCurHeaderRow As Long
    Dim lVendor As Long
    Dim lInvoice As Long

    lCurHeaderRow = tgmHeader.GetCurrentRowNumber
    
    lVendor = tfnRound(tgmHeader.CellValue(ColHeaderVendor, lCurHeaderRow))
    lInvoice = tfnFormatDate(tgmHeader.CellValue(ColHeaderInvoice, lCurHeaderRow))
    
    strSQL = "UPDATE rs_p_hold_header SET rsphh_status = 'Y' WHERE "
    strSQL = strSQL & " rsphh_vendor = " & lVendor
    strSQL = strSQL & " AND rsphh_invoice = " & lInvoice
    
    If fnExecuteSQL(strSQL, nDB_REMOTE, "fnUpdateInvHeader") Then
        tgmHeader.CellValue(ColHeaderStatus, lCurHeaderRow) = "Y"
        fnUpdateInvHeader = True
    End If
    
End Function

Private Function fnGetMaxSeqNum(sType As String) As Long
    Dim strSQL As String
    Dim rsTemp As Recordset
    
    strSQL = "SELECT rssh_seq, rssh_4 FROM rs_shifthold WHERE "
    strSQL = strSQL & " rssh_shl = " & tfnRound(g_lHoldShiftLink)
    strSQL = strSQL & " AND rssh_type = " & tfnSQLString(sType)
    strSQL = strSQL & " ORDER BY rssh_seq DESC"
    
    If fnGetRecord(rsTemp, strSQL, nDB_REMOTE, "fnGetMaxSeqNum") <= 0 Then
        fnGetMaxSeqNum = 0
    Else
        If sType = "G" Then
            
            If tfnRound(rsTemp!rssh_4) = 0 Then
                fnGetMaxSeqNum = 2 * tfnRound(rsTemp!rssh_seq) - 1
            Else
                fnGetMaxSeqNum = 2 * tfnRound(rsTemp!rssh_seq)
            End If
            
        Else
            fnGetMaxSeqNum = tfnRound(rsTemp!rssh_seq)
        End If
        
    End If
    
End Function
