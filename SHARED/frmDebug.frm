VERSION 5.00
Begin VB.Form frmDebug 
   ClientHeight    =   1215
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4305
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1215
   ScaleWidth      =   4305
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtPassword 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1710
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   540
      Width           =   1125
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   345
      Left            =   3090
      TabIndex        =   0
      Top             =   600
      Width           =   945
   End
   Begin VB.Label lblPassword 
      Caption         =   "Password"
      Height          =   285
      Left            =   630
      TabIndex        =   3
      Top             =   570
      Width           =   1005
   End
   Begin VB.Image imgFactor 
      Height          =   840
      Left            =   0
      Picture         =   "frmDebug.frx":0000
      Top             =   0
      Width           =   630
   End
   Begin VB.Label lblModuleName 
      Caption         =   "Please Call Factor! Enter Password to continue."
      Height          =   405
      Left            =   675
      TabIndex        =   1
      Top             =   60
      Width           =   3480
   End
End
Attribute VB_Name = "frmDebug"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'#Please doe not change this function!!!!!!
'#Here is a copy

''Public Function fnGeneratePassWord() As String
''    Const nSeed = 7
''    fnGeneratePassWord = Format(CLng(Format(Date + nSeed, "yyyyddmm")) Mod CLng(Format(Date, "mmdd")), "0000")
''End Function

'#Please doe not change this function!!!!!!
Public Function fnGeneratePassWord() As String
    Const nSeed = 7
    fnGeneratePassWord = Format(CLng(Format(Date + nSeed, "yyyyddmm")) Mod CLng(Format(Date, "mmdd")), "0000")
End Function

Private Sub Form_Load()
    On Error Resume Next
    Screen.MousePointer = vbDefault
    tfnDisableFormSystemClose Me
    txtPassword.SetFocus
    
    cmdOK.Enabled = False
End Sub

Private Sub txtPassword_Change()
    cmdOK.Enabled = (Trim(txtPassword) = fnGeneratePassWord)
End Sub
