VERSION 5.00
Begin VB.Form frmDebug 
   Caption         =   "Please Call Factor!"
   ClientHeight    =   975
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3030
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   975
   ScaleWidth      =   3030
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtPassword 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   690
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   480
      Width           =   1125
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   345
      Left            =   2010
      TabIndex        =   0
      Top             =   540
      Width           =   945
   End
   Begin VB.Label lblPassword 
      Caption         =   "Enter Password"
      Height          =   285
      Left            =   720
      TabIndex        =   2
      Top             =   120
      Width           =   2205
   End
   Begin VB.Image imgFactor 
      Height          =   840
      Left            =   0
      Picture         =   "frmDebug.frx":0000
      Top             =   0
      Width           =   630
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

Private Sub cmdOK_Click()
    txtPassword = ""
    Me.Hide
End Sub

Private Sub Form_Load()
    On Error Resume Next
    tfnCenterForm Me
    Screen.MousePointer = vbDefault
    tfnDisableFormSystemClose Me
    txtPassword.SetFocus
    
    cmdOK.Enabled = False
End Sub

Private Sub txtPassword_Change()
    cmdOK.Enabled = (Trim(txtPassword) = fnGeneratePassWord)
End Sub

Private Sub txtPassword_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
       KeyAscii = 0
       SendKeys "{TAB}"
    End If
End Sub
