Attribute VB_Name = "modRegDll"
Option Explicit

    Private Const FACTOR_REGISTER = "Software\Factor\ExecTrak\"
    
    Private Declare Function RegisterDll Lib "REGDLL.DLL" _
        Alias "RegisterDLL" (ByVal sPathName As String) As Long


Private Function fnExtractName(sFile As String) As String
    
    Dim nPos As Integer
    Dim sTemp As String
    
    nPos = Len(sFile)
    Do While nPos > 0
        If Mid(sFile, nPos, 1) = "\" Then
            Exit Do
        End If
        nPos = nPos - 1
    Loop
    nPos = Len(sFile) - nPos
    If nPos > 0 Then
        sTemp = Right(sFile, nPos)
    End If
    nPos = Len(sTemp)
    Do While nPos > 0
        If Mid(sTemp, nPos, 1) = "." Then
            Exit Do
        End If
        nPos = nPos - 1
    Loop
    If nPos > 1 Then
        fnExtractName = Left(sTemp, nPos - 1)
    Else
        fnExtractName = sTemp
    End If
    
End Function

Public Function fnRegisterDll(sPathName As String, bCheck As Boolean) As Boolean

    Dim sRegDateTime As String
    Dim lRegSize As Long
    Dim sFileDateTime As String
    Dim lFileSize As Long
    Dim sKeySize As String
    Dim sKeyDate As String
    
    On Error GoTo errRegDll
    sKeySize = fnExtractName(sPathName) & "Size"
    sKeyDate = fnExtractName(sPathName) & "DateTime"
    lFileSize = FileLen(sPathName)
    sFileDateTime = Trim(FileDateTime(sPathName))
    If bCheck Then
        lRegSize = Val(QueryValue(HKEY_LOCAL_MACHINE, FACTOR_REGISTER, sKeySize))
        If lRegSize = lFileSize Then
            sRegDateTime = Trim(QueryValue(HKEY_LOCAL_MACHINE, FACTOR_REGISTER, sKeyDate))
            If sRegDateTime = sFileDateTime Then
                fnRegisterDll = True
                Exit Function
            End If
        End If
    End If
    If RegisterDll(sPathName) = 0 Then
        fnRegisterDll = True
        RegSetValue HKEY_LOCAL_MACHINE, FACTOR_REGISTER, sKeySize, REG_SZ, lFileSize
        RegSetValue HKEY_LOCAL_MACHINE, FACTOR_REGISTER, sKeyDate, REG_SZ, sFileDateTime
    Else
        fnRegisterDll = False
    End If
    Exit Function
errRegDll:
'    If Err.Number = 53 Then
'        MsgBox "OLE server file: '" & sPathName & "' not found. Please contact Factor."
'    Else
'        MsgBox "Can not register OLE server (" & sPathName & "). Please contact Factor."
'    End If
End Function

