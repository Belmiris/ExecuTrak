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
   Begin VB.CommandButton cmdMakeDefault 
      Caption         =   "Make Default"
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   4560
      Width           =   1215
   End
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

Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, lpSecurityAttributes As SECURITY_ATTRIBUTES, phkResult As Long, lpdwDisposition As Long) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegQueryValueExLong Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Long, lpcbData As Long) As Long
Private Declare Function RegQueryValueExNULL Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As Long, lpcbData As Long) As Long
Private Declare Function RegQueryValueExString Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpData As String, ByVal cbData As Long) As Long        ' Note that if you declare the lpData parameter as String, you must pass it By Value.

Private Const ERROR_NONE = 0
Private Const ERROR_BADDB = 1
Private Const ERROR_BADKEY = 2
Private Const ERROR_CANTOPEN = 3
Private Const ERROR_CANTREAD = 4
Private Const ERROR_CANTWRITE = 5
Private Const ERROR_OUTOFMEMORY = 6
Private Const ERROR_INVALID_PARAMETER = 7
Private Const ERROR_ACCESS_DENIED = 8
Private Const ERROR_INVALID_PARAMETERS = 87
Private Const ERROR_NO_MORE_ITEMS = 259

Private Const HKEY_CLASSES_ROOT = &H80000000
Private Const HKEY_CURRENT_USER = &H80000001
Private Const HKEY_LOCAL_MACHINE = &H80000002
Private Const HKEY_USERS = &H80000003

Private Const KEY_ALL_ACCESS = &H3F
Private Const KEY_QUERY_VALUE = &H1
Private Const KEY_ENUMERATE_SUB_KEYS = &H8
Private Const KEY_NOTIFY = &H10
Private Const READ_CONTROL = &H20000
Private Const STANDARD_RIGHTS_READ = (READ_CONTROL)
Private Const SYNCHRONIZE = &H100000
Private Const KEY_READ = ((STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) And (Not SYNCHRONIZE))

Private Const REG_OPTION_NON_VOLATILE = 0
Private Const REG_SZ As Long = 1
Private Const REG_DWORD As Long = 4
Private Const REG_OPTION_VOLATILE = 1           ' Key is not preserved when system is rebooted

Private Type SECURITY_ATTRIBUTES
        nLength As Long
        lpSecurityDescriptor As Long
        bInheritHandle As Long
End Type

Private m_bCanceled As Boolean
Private m_bForceDefault As Boolean
'

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

Public Property Get ForceDefault() As Boolean
    ForceDefault = m_bForceDefault
End Property

Public Property Let ForceDefault(ByVal vNewValue As Boolean)
    m_bForceDefault = vNewValue
End Property

'Exit
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

'Set the VB Printer by Selection in List then Exit
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
                    
                    If Me.ForceDefault Then
                        fnForceDefaultPrinter selName
                    End If
                    
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
        MsgBox "Error selecting printer: " + Err.Description
        Err.Clear
    End If
End Sub

'Set the Selected Printer to be the WMI Default Printer
Private Sub cmdMakeDefault_Click()
    On Error GoTo FINISHED
    Dim selName As String
    Dim pr As Printer
    
    If lstPrinters.ListIndex > -1 Then
        selName = lstPrinters.List(lstPrinters.ListIndex)
        
        If selName <> "" Then
            If fnSetDefaultPrinter(selName) Then
                Me.lblDefault.Caption = "Default Print is: " & selName
            Else
                Err.Raise Number:=-1, Description:="Could Not Set the Default Printer."
            End If
        End If
    Else
        MsgBox "Please select a printer from the list"
    End If
    
    Err.Clear
FINISHED:
    If Err.Number <> 0 Then
        MsgBox "Error setting default printer: " + Err.Description
        Err.Clear
    End If
End Sub

'Retrieve the WMI Default Printer and Return its Name
Public Function fnGetDefaultPrinterName()
    On Error GoTo FINISHED
    Dim objWMIService, colInstalledPrinters, objPrinter
        
    fnGetDefaultPrinterName = Printer.DeviceName
    
    Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
    Set colInstalledPrinters = objWMIService.ExecQuery("Select * from Win32_Printer where Default = True")
    
    If Not colInstalledPrinters Is Nothing Then
        For Each objPrinter In colInstalledPrinters
            If Not objPrinter Is Nothing Then
                fnGetDefaultPrinterName = Trim(objPrinter.name)
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

'Set the WMI Default Printer by Name
Public Function fnSetDefaultPrinter(ByRef strPrinterName As String) As Boolean
    On Error GoTo FINISHED
    Dim objWMIService, colInstalledPrinters, objPrinter
        
    fnSetDefaultPrinter = False
    
    Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
    Set colInstalledPrinters = objWMIService.ExecQuery("Select * from Win32_Printer Where Name = '" & strPrinterName & "'")
 
    For Each objPrinter In colInstalledPrinters
        objPrinter.SetDefaultPrinter
    Next
        
    fnSetDefaultPrinter = True
    
    Set objWMIService = Nothing
    
    Err.Clear
FINISHED:
    If Err.Number <> 0 Then
        fnSetDefaultPrinter = False
        Err.Clear
    End If
End Function

'Retrieve a list of WMI Printers and add them to a Listbox
Sub DisplayPrinters()
    On Error GoTo FINISHED
    Dim objWMIService, colInstalledPrinters, objPrinter, oOption
    Dim sName As String
    Dim sDefault As String
        
    lstPrinters.Clear
    
    sDefault = fnGetDefaultPrinterName()
    
    Me.lblDefault.Caption = "Default Print is: " & sDefault
    
    Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
    Set colInstalledPrinters = objWMIService.ExecQuery("Select * from Win32_Printer")
    
    For Each objPrinter In colInstalledPrinters
        If Not objPrinter Is Nothing Then
            sName = Trim(objPrinter.name)
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

' Sets the VB Printer to the Default Windows Printer
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

Private Function fnForceDefaultPrinter(sName As String) As Boolean
    Const METH_NAME = "fnForceDefaultPrinter"
    On Error GoTo FINSISHED
    Dim sData As String
    Dim sValue As String
    
    sData = Trim(QueryValue(HKEY_CURRENT_USER, "Software\Microsoft\Windows NT\CurrentVersion\Devices", sName))
    
    If sData = "" Then
        MsgBox "No data found in registry devices for printer '" & sName & "'"
        Exit Function
    End If
        
    sValue = sName & "," & sData
    
    If Not RegSetValue(HKEY_CURRENT_USER, "Software\Microsoft\Windows NT\CurrentVersion\Windows", "Device", 1, sValue) Then
        MsgBox "Failed to set the printer '" & sName & "' as the default printer in the registry"
        Exit Function
    End If
    
    fnForceDefaultPrinter = True
    
    Err.Clear
FINSISHED:
    If Err.Number <> 0 Then
        MsgBox "Error in " & METH_NAME & ": " & Err.Description
        Err.Clear
    End If
End Function

Private Function QueryValue(ByVal lKey As Long, _
                           sKeyName As String, _
                           sValueName As String) As String
    Dim lRetVal As Long         'result of the API functions
    Dim hKey As Long            'handle of opened key
    Dim vValue As Variant       'setting of queried value

    QueryValue = ""
    lRetVal = RegOpenKeyEx(lKey, sKeyName, 0, KEY_READ, hKey)
    If lRetVal = 0 Then
        lRetVal = QueryValueEx(hKey, sValueName, vValue)
        If lRetVal = 0 Then
            QueryValue = vValue
        End If
        RegCloseKey (hKey)
    End If

End Function

Private Function QueryValueEx(ByVal lhKey As Long, _
                      ByVal szValueName As String, _
                      vValue As Variant) As Long
    Dim cch As Long
    Dim lrc As Long
    Dim lType As Long
    Dim lValue As Long
    Dim sValue As String

    On Error GoTo QueryValueExError

    ' Determine the size and type of data to be read
    lrc = RegQueryValueExNULL(lhKey, szValueName, 0&, lType, 0&, cch)
    If lrc = ERROR_NONE Then
        Select Case lType
            ' For strings
            Case REG_SZ:
                sValue = String(cch, 0)
                lrc = RegQueryValueExString(lhKey, szValueName, 0&, lType, sValue, cch)
                If lrc = ERROR_NONE Then
                    vValue = RemoveNull(sValue)
                Else
                    vValue = Empty
                End If
            ' For DWORDS
            Case REG_DWORD:
                lrc = RegQueryValueExLong(lhKey, szValueName, 0&, lType, lValue, cch)
                If lrc = ERROR_NONE Then
                    vValue = lValue
                End If
            Case Else
                'all other data types not supported
                lrc = -1
        End Select
    End If
QueryValueExExit:
    QueryValueEx = lrc
    Exit Function
QueryValueExError:
    Resume QueryValueExExit

End Function

Public Function RegSetValue(ByVal lKey As Long, _
                            sKeyName As String, _
                            sValueName As String, _
                            ByVal lType As Long, _
                            ByVal sValue As String) As Boolean

    Dim lRetVal As Long         'result of the API functions
    Dim hKey As Long         'handle of opened key
    Dim vValue As String      'setting of queried value
    Dim cbData As Long
    
    Dim Sec_Att As SECURITY_ATTRIBUTES
    Sec_Att.nLength = 12&
    Sec_Att.lpSecurityDescriptor = 0&
    Sec_Att.bInheritHandle = False
    
    RegSetValue = False

    'lRetVal = RegCreateKeyEx(lKey, sKeyName, 0, 0, REG_OPTION_VOLATILE, KEY_ALL_ACCESS, 0, hKey, cbData)
    lRetVal = RegCreateKeyEx(lKey, sKeyName, 0&, "", REG_OPTION_VOLATILE, KEY_ALL_ACCESS, Sec_Att, hKey, cbData)
    If lRetVal = 0 Then
        If lType >= 0 Then
            If lType = REG_DWORD Then
                vValue = Chr(sValue)
                cbData = 4
            Else
                vValue = sValue & Chr(0)
                cbData = Len(vValue)
            End If
            lRetVal = RegSetValueEx(hKey, sValueName, 0, lType, vValue, cbData)
            If lRetVal = 0 Then
                RegSetValue = True
            End If
        End If
        RegCloseKey (hKey)
    End If

End Function

Private Function RemoveNull(szStr As String) As String
    Dim nPos As Integer
    
    nPos = InStr(szStr, Chr(0))
    If nPos > 0 Then
        RemoveNull = Left(szStr, nPos - 1)
    Else
        RemoveNull = szStr
    End If

End Function

