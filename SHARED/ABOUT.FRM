VERSION 5.00
Object = "{C75015E0-2232-11D3-B440-0060971E99AF}#1.0#0"; "FACTFRM.OCX"
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About"
   ClientHeight    =   1428
   ClientLeft      =   4536
   ClientTop       =   3000
   ClientWidth     =   5664
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1428
   ScaleWidth      =   5664
   ShowInTaskbar   =   0   'False
   Begin FACTFRMLib.FactorFrame cmdOK 
      Height          =   396
      Left            =   4200
      TabIndex        =   3
      Top             =   936
      Width           =   1308
      _Version        =   65536
      _ExtentX        =   2302
      _ExtentY        =   688
      _StockProps     =   77
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.68
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "&OK"
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
   Begin VB.Label lblProduct 
      Height          =   192
      Left            =   768
      TabIndex        =   2
      Top             =   336
      Width           =   4740
   End
   Begin VB.Label lblCompany 
      Height          =   192
      Left            =   768
      TabIndex        =   1
      Top             =   600
      Width           =   4740
   End
   Begin VB.Label lblModuleName 
      Height          =   192
      Left            =   768
      TabIndex        =   0
      Top             =   60
      Width           =   4740
   End
   Begin VB.Image imgFactor 
      Height          =   672
      Left            =   96
      Picture         =   "ABOUT.frx":0000
      Top             =   60
      Width           =   504
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***********************************************************'
'
' Copyright (c) 1996 FACTOR, A Division of W.R.Hess Company
'
' Module name   : ABOUT.FRM
' Date          : Feb 21, 1995
' Programmer(s) : Jeffrey Wedekind
'
' Standard FACTOR About Box
'
'
Option Explicit

Private Sub cmdOK_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()

'    lblCompany.Caption = App.CompanyName
    lblCompany.Caption = App.LegalCopyright  ' Change to copy right for VB6. Ma 07/30/99
    lblModuleName.Caption = "Module Name: " & App.FileDescription
    lblProduct.Caption = "Version: " & CStr(App.Major) & "." & CStr(App.Minor) & "." & CStr(App.Revision)
    Screen.MousePointer = vbDefault
End Sub

