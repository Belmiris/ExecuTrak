VERSION 5.00
Begin VB.Form clsPipelineInterface 
   Caption         =   "clsPipelineInterface"
   ClientHeight    =   1335
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3180
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   1335
   ScaleWidth      =   3180
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "clsPipelineInterface"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const nERROR_FREE As Integer = 0
Private Const nERROR_FORM_NOTSET As Integer = 1
Private Const nERROR_UNKOWN_CLIENT As Integer = 2
Private Const nERROR_NO_PARAM As Integer = 3
Private Const nERROR_VNAME_MISSING As Integer = 4
Private Const nERROR_SQL_QUERY As Integer = 5
Private Const nERROR_NO_REQUEST As Integer = 6
Private Const nERROR_NO_ID As Integer = 7
Private Const nERROR_CALL_BACK As Integer = 8

Private nErrorNumber As Integer

Private m_objPipeLine As Object

Private m_sPipelineName As String
Private m_frmParent As Form
Private m_MyID As String
Private m_TheirID As String
'

Property Let PipelineName(v As String)
    m_sPipelineName = UCase(v)
End Property

Property Set MainForm(frm As Form)
    Set m_frmParent = frm
End Property

Property Get PipelineName() As String
    PipelineName = m_sPipelineName
End Property

Property Let MyID(v As String)
    m_MyID = UCase(v)
End Property

Property Get MyID() As String
    MyID = m_MyID
End Property

Property Let TheirID(v As String)
    m_TheirID = UCase(v)
End Property

Property Get TheirID() As String
    TheirID = m_TheirID
End Property

Public Function SetParamValue(v, sFieldID As String, Optional bShowError As Boolean = True) As String
    Dim nMousePointer As Integer
    Dim sErrMsg As String
    
    If m_frmParent Is Nothing Or m_MyID = "" Or m_TheirID = "" Then
        sErrMsg = "MainForm, MyID or TheirID not set!"
        MsgBox sErrMsg, vbCritical
        SetParamValue = sErrMsg
        Exit Function
    End If
    
    nMousePointer = Screen.MousePointer
    Screen.MousePointer = vbHourglass
    
    If m_objPipeLine Is Nothing Then
        On Error Resume Next
        Set m_objPipeLine = CreateObject(t_szOLEPIPELINE)
        On Error GoTo 0
    End If

    If m_objPipeLine Is Nothing Then
        tfnWaitSeconds 1
        On Error Resume Next
        Set m_objPipeLine = CreateObject(t_szOLEPIPELINE)
        On Error GoTo 0
    End If
    
    If m_objPipeLine Is Nothing Then
        sErrMsg = "Cannot create OLE PipeLine object in " + m_sPipelineName + "!"
        
        If bShowError Then
            MsgBox sErrMsg, vbCritical
        End If
        
        SetParamValue = sErrMsg
        Screen.MousePointer = nMousePointer
        Exit Function
    Else
        Set m_objPipeLine.Form = Me
        m_objPipeLine.MyID = m_MyID
    End If
    
    m_objPipeLine.SetParamValue v, sFieldID, m_TheirID
End Function

'return the instance/exe handler or 0 if failed to run the program
Public Function RunApp(sProgramPath As String, sProgramName As String, sArgs As String, _
                       Optional bHandShake As Boolean = True, _
                       Optional bUseFactMenu As Boolean = True, _
                       Optional vWindowStyle As Integer = SW_SHOWNORMAL, _
                       Optional ByVal lProcID As Long = 0, _
                       Optional ByVal bShowError As Boolean = True, _
                       Optional ByRef sErrMsg As String) As Long
    Dim bCreatePipeline As Boolean
    Dim nMousePointer As Integer
    Dim lExeHandler As Long
    
    sErrMsg = ""
    
    If m_frmParent Is Nothing Or m_MyID = "" Or m_TheirID = "" Then
        sErrMsg = "MainForm, MyID or TheirID not set!"
        MsgBox sErrMsg, vbCritical
        Exit Function
    End If
    
    nMousePointer = Screen.MousePointer
    Screen.MousePointer = vbHourglass
    
    If m_objPipeLine Is Nothing Then
        On Error Resume Next
        Set m_objPipeLine = CreateObject(t_szOLEPIPELINE)
        bCreatePipeline = True
        On Error GoTo 0
    End If

    If m_objPipeLine Is Nothing Then
        tfnWaitSeconds 1
        On Error Resume Next
        Set m_objPipeLine = CreateObject(t_szOLEPIPELINE)
        bCreatePipeline = True
        On Error GoTo 0
    End If
    
    If m_objPipeLine Is Nothing Then
        sErrMsg = "Unable to create OLE PipeLine object in " + m_sPipelineName + "!"
        
        If bShowError Then
            MsgBox sErrMsg, vbCritical
        End If
        
        Screen.MousePointer = nMousePointer
        Exit Function
    Else
        If bCreatePipeline Then
            Set m_objPipeLine.Form = Me
            m_objPipeLine.MyID = m_MyID
        End If
    End If
    
    m_objPipeLine.PutRequest m_TheirID
    
    If bUseFactMenu Then
        If t_oleObject Is Nothing Then
            sErrMsg = "Unable to run '" + sProgramName + "' in " + m_sPipelineName + "!"
            
            If bShowError Then
                MsgBox "Factor Menu is not running!", vbCritical
                MsgBox sErrMsg, vbCritical
            End If
        Else
            'note: sProgramName, ID only and should not have .EXE
            tfnExecuteProgram t_oleObject, UCase(sProgramName)
        End If
    
        Screen.MousePointer = nMousePointer
        Exit Function
    End If
    
    'lunch app that does not need factor menu
    Dim sProgram As String
    
    If bHandShake Then
        sProgram = App.Path + "\" + Trim(sProgramName) + ".EXE"
        
        If Not fnExeIsRunning(sProgram, lProcID) Then
            If Not tfnRun(sProgramName, vWindowStyle, bHandShake, sArgs, lExeHandler, False) Then
                sErrMsg = "Unable to run '" + sProgram + "' in " + m_sPipelineName + "!"
                
                If bShowError Then
                    MsgBox sErrMsg, vbCritical
                End If
                
                Screen.MousePointer = nMousePointer
                Exit Function
            End If
        End If
    Else
        sProgram = Trim(sProgramPath)
        If sProgram <> "" Then
            If Right(sProgram, 1) <> "\" Then
                sProgram = sProgram + "\"
            End If
        End If
    
        sProgram = sProgram + Trim(sProgramName)
    
        If Not fnRunExe(sProgram, vWindowStyle, , , bShowError) Then
            sErrMsg = "Unable to run '" + sProgram + "' in " + m_sPipelineName + "!"
            
            If bShowError Then
                MsgBox sErrMsg, vbCritical
            End If
            
            Screen.MousePointer = nMousePointer
            Exit Function
        End If
    End If
    
    RunApp = lExeHandler
    Screen.MousePointer = nMousePointer
End Function

Public Sub Reset()
    On Error Resume Next

    If Not m_objPipeLine Is Nothing Then
        m_objPipeLine.DeleteRequest
        Set m_objPipeLine = Nothing
    End If
End Sub

Public Sub SendSignal(ByVal v As String)
    Dim nMousePointer As Integer
    Dim bCreatePipeline As Boolean
    Dim sErrMsg As String
    
    If m_frmParent Is Nothing Or m_MyID = "" Or m_TheirID = "" Then
        'sErrMsg = "MainForm, MyID or TheirID not set!"
        'MsgBox sErrMsg, vbCritical
        Exit Sub
    End If
    
    v = Trim(v)
    If v = "" Then
        Exit Sub
    End If
    
    nMousePointer = Screen.MousePointer
    Screen.MousePointer = vbHourglass
    
    If m_objPipeLine Is Nothing Then
        On Error Resume Next
        Set m_objPipeLine = CreateObject(t_szOLEPIPELINE)
        bCreatePipeline = True
        On Error GoTo 0
    End If

    If m_objPipeLine Is Nothing Then
        tfnWaitSeconds 1
        On Error Resume Next
        Set m_objPipeLine = CreateObject(t_szOLEPIPELINE)
        bCreatePipeline = True
        On Error GoTo 0
    End If
    
    If m_objPipeLine Is Nothing Then
        sErrMsg = "Unable to create OLE PipeLine object in " + m_sPipelineName + "!"
        MsgBox sErrMsg, vbCritical
        Screen.MousePointer = nMousePointer
        Exit Sub
    Else
        If bCreatePipeline Then
            Set m_objPipeLine.Form = Me
            m_objPipeLine.MyID = m_MyID
        End If
    End If
    
    m_objPipeLine.SetParamValue v, "signal", m_TheirID
    m_objPipeLine.PutRequest m_TheirID
    Screen.MousePointer = nMousePointer
End Sub

Property Get ErrorNumber() As Integer
    ErrorNumber = nErrorNumber
End Property

Property Get ErrorDescription() As String
    Select Case nErrorNumber
        Case nERROR_FORM_NOTSET
            ErrorDescription = "The call back form is not set in PipeLine OLE"
        Case nERROR_UNKOWN_CLIENT
            ErrorDescription = "Unkown client who receives parameters"
        Case nERROR_NO_PARAM
            ErrorDescription = "No parameters for the module"
        Case nERROR_VNAME_MISSING
            ErrorDescription = "Parameter name missing for multi-parameter assignment"
        Case nERROR_SQL_QUERY
            ErrorDescription = "SQL query error in PipeLine OLE" '& CRLF & CRLF & strSQL
        Case nERROR_NO_REQUEST
            ErrorDescription = "No request is send to wait"
        Case nERROR_NO_ID
            ErrorDescription = "Module ID is not set"
        Case nERROR_CALL_BACK
            ErrorDescription = "Call back function not defined in the call back form"
    End Select
End Property

'need to be public
Public Sub BackFromServer()
    If Not m_objPipeLine Is Nothing Then
        m_objPipeLine.DeleteRequest
    
        If Not m_frmParent Is Nothing Then
            On Error Resume Next
            m_frmParent.BackFromServer m_sPipelineName
            If Err.num > 0 Then
                nErrorNumber = nERROR_CALL_BACK
            End If
            On Error GoTo 0
        End If
    End If
End Sub

Public Sub CleanUp()
    Unload Me
End Sub

Private Sub Form_Initialize()
    m_sPipelineName = "Pipeline Object"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Reset
End Sub
