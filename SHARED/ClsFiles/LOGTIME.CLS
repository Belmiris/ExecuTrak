VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLOGTIME"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type typTask
    sSubName As String
    sDesc As String
    sngStart As Single
End Type

Const MAX_LOGTIME As Long = 5000

Private m_Task() As typTask
Private m_bLogActive As Boolean
Private m_nLastLog As Long
Private m_sFileName As String
Private m_nFilePtr As Integer
'

Private Sub Class_Initialize()
    m_bLogActive = False
    ReDim m_Task(MAX_LOGTIME)
    m_nLastLog = -1
    m_sFileName = ""
    m_nFilePtr = -999
End Sub

Private Sub Class_Terminate()
    ReDim m_Task(0)
    If m_nFilePtr <> -999 Then
        On Error Resume Next
        Close #m_nFilePtr
    End If
End Sub

Public Sub TurnOn()
    If m_bLogActive Then
        ShowMessage "Log Time already Turn On."
        Exit Sub
    End If
    
    If m_sFileName = "" Then
        m_sFileName = InputBox("Enter a path and filename for TimeLog (if any):", "FileName", App.Path + "\" + "TIMELOG.TXT")
    End If
    
    m_bLogActive = True
    m_nLastLog = -1
    
    On Error GoTo errTrap
    
    If m_sFileName <> "" Then
        m_nFilePtr = FreeFile
        Open m_sFileName For Append As #m_nFilePtr
    End If
    
    Exit Sub

errTrap:
    m_nFilePtr = -999
    ShowMessage "Error(s) has occurred." + vbCrLf + vbCrLf + "ErrCode=" & Err.Number & _
        vbCrLf & "ErrDesc=" & Err.Description
End Sub

Public Sub TurnOff()
    m_bLogActive = False
    If m_nFilePtr <> -999 Then
        On Error Resume Next
        Close #m_nFilePtr
    End If
End Sub

Public Sub Start(sSubName As String, sTaskDesc As String)
    If Not m_bLogActive Then
        Exit Sub
    End If
    
    m_nLastLog = m_nLastLog + 1
    
    If m_nLastLog >= MAX_LOGTIME Then
        ShowMessage "Log Time Counts reach the Max Count (" & MAX_LOGTIME & ").", True
        Exit Sub
    End If
    
    m_Task(m_nLastLog).sSubName = sSubName
    m_Task(m_nLastLog).sDesc = sTaskDesc
    m_Task(m_nLastLog).sngStart = Timer
End Sub

Public Sub Finish(sSubName As String, Optional sTaskDesc As String = "")
    Dim l As Long
    Dim sMsg As String
    
    If Not m_bLogActive Then
        Exit Sub
    End If

    If m_nLastLog < 0 Then
        ShowMessage "No Time Log available!"
        Exit Sub
    End If
    
    If m_Task(m_nLastLog).sSubName <> sSubName Then
        If MsgBox("Time Log for '" + sSubName + "' not found. Continue search?", vbYesNo + vbQuestion) = vbNo Then
            Exit Sub
        End If
        
        For l = m_nLastLog To 0 Step -1
            If m_Task(l).sSubName = sSubName Then
                Exit For
            End If
        Next l
        
        m_nLastLog = l
    End If
    
    If m_nLastLog < 0 Then
        ShowMessage "No Time Log available!"
        Exit Sub
    End If
    
    sMsg = IIf(m_Task(m_nLastLog).sSubName = "", "", "FUNC=" _
        + m_Task(m_nLastLog).sSubName + ", ") _
        + IIf(m_Task(m_nLastLog).sDesc = "", "", "DESC=" _
        + m_Task(m_nLastLog).sDesc + ", ") _
        + "Time=" & Timer - m_Task(m_nLastLog).sngStart
    
    Debug.Print sMsg
    ShowMessage sMsg, False
    
    m_nLastLog = m_nLastLog - 1
End Sub

Property Let FileName(sFilename As String)
    m_sFileName = sFilename
End Property

Private Sub ShowMessage(sMsg As String, Optional bShowMsgBox)
    If m_nFilePtr <> -999 Then
        Print #m_nFilePtr, sMsg
        
        If Not IsMissing(bShowMsgBox) Then
            If bShowMsgBox Then
                MsgBox sMsg, vbInformation
            End If
        End If
    Else
        If Not IsMissing(bShowMsgBox) Then
            If bShowMsgBox Then
                MsgBox sMsg, vbInformation
            End If
        Else
            MsgBox sMsg, vbInformation
        End If
    End If
End Sub
