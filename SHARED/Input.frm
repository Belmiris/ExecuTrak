VERSION 4.00
Begin VB.Form frmSave 
   ClientHeight    =   1164
   ClientLeft      =   972
   ClientTop       =   2724
   ClientWidth     =   7248
   Height          =   1548
   Left            =   924
   LinkTopic       =   "Form2"
   ScaleHeight     =   1164
   ScaleWidth      =   7248
   Top             =   2388
   Width           =   7344
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   372
      Left            =   5724
      TabIndex        =   3
      Top             =   720
      Width           =   1332
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   372
      Left            =   4296
      TabIndex        =   2
      Top             =   720
      Width           =   1332
   End
   Begin VB.TextBox txtInput 
      Height          =   288
      Left            =   192
      TabIndex        =   0
      Top             =   312
      Width           =   6864
   End
   Begin VB.Label lblCaption 
      Height          =   276
      Left            =   204
      TabIndex        =   1
      Top             =   48
      Width           =   7128
   End
End
Attribute VB_Name = "frmSave"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Option Explicit

    Private Const SAVE_CASE = 1
    Private Const SAVE_SOLUTION = 2
    
    Private m_nCaseFlag As Integer
    Private m_lCase As Long
    
Public Property Let PCase(lTemp As Long)
    m_lCase = lTemp
End Property

Property Get SaveCase() As Integer
    SaveCase = SAVE_CASE
End Property

Property Get SaveSolution() As Integer
    SaveSolution = SAVE_SOLUTION
End Property

Property Let SaveFlag(nTemp As Integer)
    m_nCaseFlag = nTemp
End Property

Private Sub cmdCancel_Click()
    Me.Hide
End Sub


Private Sub cmdOK_Click()
    Select Case m_nCaseFlag
        Case SAVE_CASE
            subSaveCases txtInput.Text
        Case SAVE_SOLUTION
            subSaveSolutions m_lCase, txtInput.Text
    End Select
    Me.Hide
End Sub

Private Sub Form_Activate()
    Select Case m_nCaseFlag
        Case SAVE_CASE
            lblCaption.Caption = "Enter case description"
        Case SAVE_SOLUTION
            lblCaption.Caption = "Enter solution description"
    End Select
End Sub

