VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLockFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Private m_sLockFilename As String
Private m_bShowMsgBox As Boolean
Private m_nMinutesPerSession As Integer
Private m_nDeleteLockAfterMinutes As Integer
'

Property Let ShowMessageBox(bYesNo As Boolean)
    m_bShowMsgBox = bYesNo
End Property

Property Get ShowMessageBox() As Boolean
    ShowMessageBox = m_bShowMsgBox
End Property

Property Let MinutesPerSession(nMinutes As Integer)
    m_nMinutesPerSession = nMinutes
End Property

Property Get MinutesPerSession() As Integer
    MinutesPerSession = m_nMinutesPerSession
End Property

Property Let DeleteLockAfterMinutes(nMinutes As Integer)
    m_nDeleteLockAfterMinutes = nMinutes
End Property

Property Get DeleteLockAfterMinutes() As Integer
    DeleteLockAfterMinutes = m_nDeleteLockAfterMinutes
End Property

Public Function Initialize(sID As String) As String
    If sID = "" Then
        MsgBox "Main Program ID (LockFilename) is required.", vbExclamation
    Else
        m_sLockFilename = sID
    End If
    
    Initialize = fnDeleteOldLocks()
End Function

Private Sub Class_Initialize()
    m_bShowMsgBox = True
    m_nMinutesPerSession = 5
    m_nDeleteLockAfterMinutes = 10
End Sub

Private Sub Class_Terminate()
    m_sLockFilename = ""
End Sub

'return error message if anay
Public Function LockIt(Optional sLockFileID As String = "", _
                       Optional sErrMsg As String) As Boolean
    
    If sLockFileID = "" Then
        sLockFileID = m_sLockFilename
    End If
    
    If sLockFileID = "" Then
        MsgBox "Program ID (LockFilename) is required.", vbExclamation
        Exit Function
    End If
    
    sErrMsg = ""
    
    If Not IsLocked(sLockFileID, sErrMsg) Then
        If sErrMsg = "" Then
            sErrMsg = fnWriteLockEntry(ComputerName(), UserName(), Now)
            
            If sErrMsg = "" Then
                'got the lock
                LockIt = True
                Exit Function
            End If
        End If
    End If
    
    LockIt = False
        
End Function

Public Sub UnlockIt(Optional sLockFileID As String = "", _
                    Optional sErrMsg As String = "")
    
    If sLockFileID = "" Then
        sLockFileID = m_sLockFilename
    End If
    
    If sLockFileID = "" Then
        MsgBox "Program ID (LockFilename) is required.", vbExclamation
        Exit Sub
    End If
    
    sErrMsg = ""
    
    If Not IsLocked(sLockFileID, "") Then
        If sErrMsg = "" And fnExists(m_sLockFilename) Then
            If fnDeleteFile(m_sLockFilename, sErrMsg) Then
                sErrMsg = ""
            Else
                If m_bShowMsgBox Then
                    MsgBox sErrMsg, vbCritical
                End If
            End If
        End If
    End If

End Sub

Public Function IsLocked(sLockFileID As String, _
                         Optional sErrMsg As String = "") As Boolean
    
    If sLockFileID = "" Then
        sLockFileID = m_sLockFilename
    End If
    
    If sLockFileID = "" Then
        MsgBox "Program ID (LockFilename) is required.", vbExclamation
        Exit Function
    End If
    
    sErrMsg = ""
    
    If Not fnExists(sLockFileID) Then
        Exit Function
    End If
    
    Dim sLockedComputer As String
    Dim sLockedUser As String
    
    sLockedComputer = ""
    sLockedUser = ""
    
    sErrMsg = fnReadLockEntry(sLockedComputer, sLockedUser)
    
    If sErrMsg <> "" Then
        If m_bShowMsgBox Then
            MsgBox sErrMsg, vbCritical
            Exit Function
        End If
    End If
    
    If sLockedComputer <> ComputerName() Or sLockedUser <> UserName() Then
        If sLockedUser = "" Then
            sLockedUser = " another user"
        Else
            sLockedUser = " user " + sLockedUser
        End If
        
        sErrMsg = "Process was locked by " + sLockedUser + "."
        IsLocked = True
    End If
    
End Function

'return error message if anay
Private Function fnDeleteOldLocks() As String
    
    Dim sLockedComputer As String
    Dim sLockedUser As String
    Dim sLockedTime As String
    Dim bDelete As Boolean
    Dim sErrMsg As String
    
    If m_sLockFilename = "" Then
        MsgBox "Program ID (LockFilename) is required.", vbExclamation
        Exit Function
    End If
    
    fnDeleteOldLocks = fnReadLockEntry(sLockedComputer, sLockedUser, sLockedTime)
    
    If fnDeleteOldLocks <> "" Then
        Exit Function
    End If
    
    bDelete = False
    
    'delete the lock file if the locktime is 5 minuties or older
    If sLockedComputer = "" Or sLockedUser = "" Then
        bDelete = True
    ElseIf sLockedComputer = ComputerName() And sLockedUser = UserName() Then
        bDelete = True
        
        If sLockedTime <> "" Then
            If DateDiff("n", CDate(sLockedTime), Now) < m_nMinutesPerSession Then
                bDelete = False
            End If
        End If
    'delete lock file that is more than 30 minutes old
    ElseIf DateDiff("n", CDate(sLockedTime), Now) > m_nDeleteLockAfterMinutes Then
        bDelete = True
    End If
    
    If bDelete And fnExists(m_sLockFilename) Then
        If Not fnDeleteFile(m_sLockFilename, sErrMsg) Then
            fnDeleteOldLocks = sErrMsg
        End If
    End If
    
End Function

Private Function fnReadLockEntry(sLockedComputer As String, _
                                 sLockedUser As String, _
                                 Optional sLockedTime As String = "") As String

    On Error GoTo errTrap
    
    sLockedComputer = ""
    sLockedUser = ""
    sLockedTime = ""
    
    If Not fnExists(m_sLockFilename) Then
        Exit Function
    End If
        
    Dim nFileNumber As Integer
    Dim sTemp As String
    Dim ary() As String
    
    nFileNumber = FreeFile
    
    Open m_sLockFilename For Input As #nFileNumber
    Line Input #nFileNumber, sTemp
    Close nFileNumber
    
    ary = Split(sTemp, "|")
    
    On Error Resume Next
    sLockedComputer = ary(0)
    sLockedUser = ary(1)
    sLockedTime = ary(2)
    
    Exit Function

errTrap:
    fnReadLockEntry = "Failed to read lock entry." + vbCrLf _
        + "Err Code: " & Err.Number & ", Err Desc: " & Err.Description
End Function

Private Function fnWriteLockEntry(sComputer As String, _
                                  sUser As String, _
                                  sTime As String) As String

    On Error GoTo errTrap
    
    Dim sLockedComputer As String
    Dim sLockedUser As String
    Dim sErrMsg As String
    
    If IsLocked(m_sLockFilename, sErrMsg) Then
        fnWriteLockEntry = sErrMsg
        Exit Function
    ElseIf sErrMsg <> "" Then
        fnWriteLockEntry = sErrMsg
        Exit Function
    End If
        
    Dim nFileNumber As Integer
    Dim sTemp As String
    
    sTemp = sComputer + "|" + sUser + "|" + sTime
    nFileNumber = FreeFile
    
    Open m_sLockFilename For Output As #nFileNumber
    Print #nFileNumber, sTemp
    Close nFileNumber
    
    Exit Function

errTrap:
    fnWriteLockEntry = "Failed to write lock entry." + vbCrLf _
        + "Err Code: " & Err.Number & ", Err Desc: " & Err.Description
End Function

Private Function fnExists(sFile1 As String, Optional Attributes As Integer = vbNormal) As Boolean

    Dim sFile As String
    
    On Error Resume Next
    sFile = ""
    sFile = Dir(sFile1, Attributes)
    
    fnExists = sFile <> ""

End Function

Private Function fnDeleteFile(sFile1 As String, sRetMsg As String) As Boolean

    On Error GoTo errTrap
    Kill sFile1
    sRetMsg = "Successfully deleted " + sFile1 + "."
    fnDeleteFile = True
    Exit Function
    
errTrap:
    
    sRetMsg = "Failed to delete " + sFile1 + "." + vbCrLf
    sRetMsg = sRetMsg + "Err Code: " & Err.Number & ", Err Desc: " & Err.Description

End Function

Public Function ComputerName() As String
    Dim lngReturnCode As Long
    Dim strHostname As String
    strHostname = Space$(50)
    lngReturnCode = GetComputerName(strHostname, 50)
    ComputerName = Left$(strHostname, InStr(strHostname, Chr$(0)) - 1)
End Function

Public Function UserName() As String
    Dim nSize As Long
    Dim sName As String
    
    nSize = 128
    sName = Space(nSize)
    GetUserName sName, nSize
    
    sName = Trim(UCase(Trim(Left(sName, nSize - 1))))
    
    If Len(sName) > 12 Then
        sName = Left(sName, 12)
    End If
    
    UserName = Trim(sName)
End Function
