VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmBlobExample 
   Caption         =   "ModBlob Example"
   ClientHeight    =   2400
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3540
   LinkTopic       =   "Form1"
   ScaleHeight     =   2400
   ScaleWidth      =   3540
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdAction 
      Caption         =   "Store"
      Height          =   405
      Left            =   1050
      TabIndex        =   7
      Top             =   1920
      Width           =   1200
   End
   Begin VB.TextBox txtID 
      Height          =   330
      Left            =   0
      TabIndex        =   6
      Top             =   1500
      Width           =   2355
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Retrieve"
      Height          =   435
      Left            =   855
      TabIndex        =   4
      Top             =   45
      Width           =   975
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Store"
      Height          =   315
      Left            =   75
      TabIndex        =   3
      Top             =   90
      Value           =   -1  'True
      Width           =   705
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Browse"
      Height          =   330
      Left            =   2430
      TabIndex        =   2
      Top             =   855
      Width           =   1065
   End
   Begin VB.TextBox txtFileName 
      Height          =   330
      Left            =   15
      TabIndex        =   0
      Top             =   840
      Width           =   2370
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   2955
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Select a file to store"
      Filter          =   "*.bmp"
   End
   Begin VB.Label Label2 
      Caption         =   "ID"
      Height          =   270
      Left            =   45
      TabIndex        =   5
      Top             =   1245
      Width           =   1725
   End
   Begin VB.Label Label1 
      Caption         =   "Blob File"
      Height          =   270
      Left            =   15
      TabIndex        =   1
      Top             =   510
      Width           =   1725
   End
End
Attribute VB_Name = "frmBlobExample"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const nDB_REMOTE As Integer = 1

Private Sub cmdAction_Click()
    frmBlobExample.MousePointer = vbHourglass
    
    If Option1.value = True And Dir(txtFileName) <> "" _
    And txtID.Text = Int(txtID.Text) Then
        SaveBlob txtFileName.Text, "blob_det", Int(txtID.Text), True
    ElseIf Option1.value = False And txtID.Text = Int(txtID.Text) Then
        GetBlob txtFileName.Text, "blob_det", Int(txtID.Text)
    Else
        MsgBox "Error"
    End If
    frmBlobExample.MousePointer = vbDefault
    
End Sub

Private Sub Command1_Click()
    CommonDialog1.ShowOpen
    If CommonDialog1.FileName <> "" Then
        txtFileName.Text = CommonDialog1.FileName
    End If
    
End Sub

Private Sub Form_Load()
    If tfnOpenDatabase = False Then 'open the database, ODBC Dialog Box during developemnt, oleObject Connection String when not
        Unload Me
        Exit Sub
    End If

End Sub

Private Sub Option1_Click()
Option2_LostFocus
End Sub

Private Sub Option1_KeyUp(KeyCode As Integer, Shift As Integer)
Option2_LostFocus
End Sub

Private Sub Option1_LostFocus()
Option2_LostFocus
End Sub

Private Sub Option2_Click()
Option2_LostFocus
End Sub

Private Sub Option2_KeyUp(KeyCode As Integer, Shift As Integer)
Option2_LostFocus
End Sub

Private Sub Option2_LostFocus()
    If Option2.value Then
        Label1.Caption = "Store as file:"
        Command1.Enabled = False
        cmdAction.Caption = "Get File"
    Else
        Label1.Caption = "File to store:"
        Command1.Enabled = True
        cmdAction.Caption = "Store File"
        
    End If
End Sub
