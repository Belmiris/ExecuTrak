VERSION 5.00
Begin VB.Form frmReports 
   Caption         =   "Print Reports"
   ClientHeight    =   1608
   ClientLeft      =   48
   ClientTop       =   288
   ClientWidth     =   3672
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   1608
   ScaleWidth      =   3672
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2208
      TabIndex        =   3
      Top             =   1116
      Width           =   1308
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "O&K"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   288
      TabIndex        =   2
      Top             =   1104
      Width           =   1308
   End
   Begin VB.CheckBox chkPrintProcLog 
      Caption         =   "Print Processing Log"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   348
      Left            =   828
      TabIndex        =   1
      Top             =   576
      Width           =   2076
   End
   Begin VB.CheckBox chkPrintErr 
      Caption         =   "Print Error Log"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   348
      Left            =   828
      TabIndex        =   0
      Top             =   108
      Value           =   1  'Checked
      Width           =   2076
   End
End
Attribute VB_Name = "frmReports"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim sProcessLogFile As String
    Dim sErrorLogFile As String
    
    sProcessLogFile = App.Path & "\zzrinvpl.log"
    sErrorLogFile = App.Path & "\zzrinver.log"
    
    If chkPrintErr.Value = vbChecked Then
        
        If Not fnFileExist(sErrorLogFile) Then
            MsgBox "No Errors log file available.", vbExclamation
            Exit Sub
        End If
    
        subSentErrorLogToPrinter
    End If
    
    DoEvents
    
    If chkPrintProcLog.Value = vbChecked Then
        
        If Not fnFileExist(sProcessLogFile) Then
            MsgBox "No processing log file available.", vbExclamation
            Exit Sub
        End If

        subSentProcLogToPrinter
    End If
    
    Unload Me
End Sub

Private Sub Form_Load()
    tfnCenterForm Me, frmZZFINVMT
End Sub
 _
 _
 _
 _
 _
 _
 _
 _
 _
 _
 _
 _

