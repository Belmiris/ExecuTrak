VERSION 5.00
Begin VB.Form frmDebug 
   ClientHeight    =   1425
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3690
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1425
   ScaleWidth      =   3690
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtPassword 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1230
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   960
      Width           =   1245
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   405
      Left            =   2580
      TabIndex        =   0
      Top             =   960
      Width           =   1065
   End
   Begin VB.Label lblErr 
      Caption         =   "Call Factor!"
      Height          =   315
      Index           =   1
      Left            =   690
      TabIndex        =   4
      Top             =   420
      Width           =   2895
   End
   Begin VB.Label lblErr 
      Caption         =   "Some data error(s) occurred. Please"
      Height          =   285
      Index           =   0
      Left            =   690
      TabIndex        =   3
      Top             =   90
      Width           =   2895
   End
   Begin VB.Label lblPassword 
      Caption         =   "Enter Password"
      Height          =   285
      Left            =   0
      TabIndex        =   2
      Top             =   1020
      Width           =   1275
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

Public Sub SetFormTitleCaption(sFormTitle As String, sTitle1 As String, sTitle2 As String)
    Me.Caption = sFormTitle
    lblErr(0).Caption = sTitle1
    lblErr(1).Caption = sTitle2
End Sub

Private Sub Form_Activate()
    On Error Resume Next
    txtPassword_Change
    If cmdOK.Enabled Then
        cmdOK.SetFocus
    Else
        txtPassword.SetFocus
    End If
End Sub

Private Sub Form_Load()
    On Error Resume Next
    Me.Caption = App.EXEName & " - Error Alert!"
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
