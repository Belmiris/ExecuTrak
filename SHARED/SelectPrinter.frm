VERSION 5.00
Begin VB.Form SelectPrinter 
   Caption         =   "Select a Printer"
   ClientHeight    =   5145
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6315
   LinkTopic       =   "Form1"
   ScaleHeight     =   5145
   ScaleWidth      =   6315
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   495
      Left            =   4920
      TabIndex        =   2
      Top             =   4560
      Width           =   1215
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Default         =   -1  'True
      Height          =   495
      Left            =   3600
      TabIndex        =   1
      Top             =   4560
      Width           =   1215
   End
   Begin VB.ListBox lstPrinters 
      Height          =   4140
      IntegralHeight  =   0   'False
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   6015
   End
   Begin VB.Label lblDefault 
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   6015
   End
End
Attribute VB_Name = "SelectPrinter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_bCanceled As Boolean

Private Sub Form_Load()
    On Error GoTo FINISHED
    
    m_bCanceled = True
    
    DisplayPrinters
    
FINISHED:
    If Err.Number <> 0 Then
        MsgBox "Error loading select printer form: " + Err.Description
        Err.Clear
    End If
End Sub

Public Property Get Canceled() As Boolean
    Canceled = m_bCanceled
End Property

Private Sub cmdCancel_Click()
    On Error GoTo FINISHED
    
    m_bCanceled = True
    
    Unload Me
    
FINISHED:
    If Err.Number <> 0 Then
        MsgBox "Error unloading select printer form: " + Err.Description
        Err.Clear
    End If
End Sub

Private Sub cmdOk_Click()
    On Error GoTo FINISHED
    Dim selName As String
    Dim pr As Printer
    
    If lstPrinters.ListIndex > -1 Then
        selName = lstPrinters.List(lstPrinters.ListIndex)
        
        If selName <> "" Then
            For Each pr In Printers
                If pr.DeviceName = selName Then
                    Set Printer = pr
                    m_bCanceled = False
                    Unload Me
                    Exit Sub
                End If
            Next pr
        End If
    Else
        MsgBox "Please select a printer from the list"
    End If
    
    Err.Clear
FINISHED:
    If Err.Number <> 0 Then
        MsgBox "Error selecting default printer: " + Err.Description
        Err.Clear
    End If
End Sub

Public Function fnGetDefaultPrinterName()
    On Error GoTo FINISHED
    Dim objWMIService, colInstalledPrinters, objPrinter
        
    fnGetDefaultPrinterName = Printer.DeviceName
    
    Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
    Set colInstalledPrinters = objWMIService.ExecQuery("Select * from Win32_Printer where Default = True")
    
    If Not colInstalledPrinters Is Nothing Then
        For Each objPrinter In colInstalledPrinters
            If Not objPrinter Is Nothing Then
                fnGetDefaultPrinterName = Trim(objPrinter.Name)
                'MsgBox "Default Printer = " & sName
                Exit For
            End If
        Next
    End If
    
    Set objWMIService = Nothing
    
    Err.Clear
FINISHED:
    If Err.Number <> 0 Then
        MsgBox "Error retrieving default printer: " & Err.Description
        Err.Clear
    End If
End Function

Sub DisplayPrinters()
    On Error GoTo FINISHED
    Dim objWMIService, colInstalledPrinters, objPrinter, oOption
    Dim sName As String
    Dim sDefault As String
        
    lstPrinters.Clear
    
    sDefault = fnGetDefaultPrinterName()
    
    Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
    Set colInstalledPrinters = objWMIService.ExecQuery("Select * from Win32_Printer")
    
    For Each objPrinter In colInstalledPrinters
        If Not objPrinter Is Nothing Then
            sName = Trim(objPrinter.Name)
            lstPrinters.AddItem Trim(sName)
            If sName = sDefault Then
                lstPrinters.ListIndex = lstPrinters.ListCount - 1
            End If
        End If
    Next

    Set objWMIService = Nothing
    
    Err.Clear
FINISHED:
    If Err.Number <> 0 Then
        MsgBox "Error loading printers: " & Err.Description
        Err.Clear
    End If
End Sub

Public Sub subAssignDefaultPrinter()
    Dim sName As String
    Dim pr As Printer
    
    sName = fnGetDefaultPrinterName()
    
    For Each pr In Printers
        If pr.DeviceName = sName Then
            Set Printer = pr
            m_bCanceled = False
            Exit Sub
        End If
    Next pr
    
End Sub

