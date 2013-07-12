VERSION 5.00
Begin VB.Form frmReports 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Print Reports"
   ClientHeight    =   1620
   ClientLeft      =   30
   ClientTop       =   270
   ClientWidth     =   3675
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1620
   ScaleWidth      =   3675
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      HelpContextID   =   15
      Left            =   2208
      TabIndex        =   3
      Top             =   1116
      Width           =   1308
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "O&K"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      HelpContextID   =   16
      Left            =   288
      TabIndex        =   2
      Top             =   1104
      Width           =   1308
   End
   Begin VB.CheckBox chkPrintProcLog 
      Caption         =   "Print Pre-Processing Log"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   348
      HelpContextID   =   804
      Left            =   672
      TabIndex        =   1
      Top             =   576
      Width           =   2628
   End
   Begin VB.CheckBox chkPrintErr 
      Caption         =   "Print Error Log"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   348
      HelpContextID   =   803
      Left            =   672
      TabIndex        =   0
      Top             =   108
      Value           =   1  'Checked
      Width           =   2532
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
    
    sProcessLogFile = io.ApplicationPath & "zzseplog.log"
    sErrorLogFile = io.ApplicationPath & "zzseelog.log"
    
    If chkPrintErr.value = vbChecked Then
        
        If Not fnFileExist(sErrorLogFile) Then
            MsgBox "No Errors log file available.", vbExclamation
            Exit Sub
        End If
    
        subSentErrorLogToPrinter
    End If
    
    DoEvents
    
    If chkPrintProcLog.value = vbChecked Then
        
        If Not fnFileExist(sProcessLogFile) Then
            MsgBox "No processing log file available.", vbExclamation
            Exit Sub
        End If

        subSentProcLogToPrinter
    End If
    
    Unload Me
End Sub

Private Sub Form_Load()
    tfnCenterForm Me, frmzzseinvp
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

