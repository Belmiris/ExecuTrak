VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FileServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

    #Const UPDATE_RIGHT_AWAY = False
    
    Private Const BUFFER_SIZE As Long = 111 '16384   '32768        '&H8000
    
    Private sFileBuffer As String * BUFFER_SIZE
    Private nInputFile As Integer
    
    Private sInputBuffer As String
    Private nInputPointer As Long
    Private nInputPointerOld As Long
    Private bFileEnd As Boolean
    Private lCurrentRecord As Long
    Private bBufferModefied As Boolean
    Private lResidualLength As Integer
    Private m_sFileName As String
    
Public Sub CloseFile()
    If bBufferModefied Then
        Put #nInputFile, lCurrentRecord, sFileBuffer
        bBufferModefied = False
    End If
    Close #nInputFile
End Sub

Property Get EndOfFile() As Boolean
    EndOfFile = False
    If bFileEnd Then
        If nInputPointer >= Len(sInputBuffer) Then
            EndOfFile = True
        End If
    End If
End Property

Public Sub Replace(ByVal nOffset As Integer, _
                   sOld As String, _
                   sNew As String)

    Dim nPos As Integer
    Dim nLen As Integer
    Dim sTemp As String
    Dim nTemp1 As Integer
    Dim lPtr As Long
    
    nPos = Len(sNew)
    nLen = Len(sOld)
    If nLen > nPos Then
        sTemp = sNew & Space(nLen - nPos)
    Else
        sTemp = Left(sNew, nLen)
    End If
    nPos = nInputPointerOld + nOffset - lResidualLength - 1
    lPtr = lCurrentRecord
    If nPos <= 0 Then
        While nPos <= 0
            lPtr = lPtr - 1
            nPos = nPos + BUFFER_SIZE
        Wend
        Get #nInputFile, lPtr, sFileBuffer
    End If
    If (nPos + nLen) > BUFFER_SIZE Then
        nTemp1 = BUFFER_SIZE - nPos + 1
        Mid(sFileBuffer, nPos, nTemp1) = Left(sTemp, nTemp1)
        nLen = nLen - nTemp1
        Do
            Put #nInputFile, lPtr, sFileBuffer
            lPtr = lPtr + 1
            sTemp = Right(sTemp, nLen)
            sFileBuffer = Left(sTemp, BUFFER_SIZE)
            nLen = nLen - BUFFER_SIZE
        Loop Until nLen < 0
        nLen = nLen + BUFFER_SIZE
        If nLen > 0 Then
            Get #nInputFile, lPtr, sFileBuffer
            Mid(sFileBuffer, 1, nLen) = sTemp
            Put #nInputFile, lPtr, sFileBuffer
        End If
        bBufferModefied = False
        Get #nInputFile, lCurrentRecord, sFileBuffer
    Else
        Mid(sFileBuffer, nPos, nLen) = sTemp
        #If UPDATE_RIGHT_AWAY Then
            Put #nInputFile, lCurrentRecord, sFileBuffer
            Close #nInputFile
            Open m_sFileName For Random As #nInputFile Len = BUFFER_SIZE
            bBufferModefied = False
        #Else
            bBufferModefied = True
        #End If
    End If
End Sub

Private Function subGetFromFile() As String
    If bBufferModefied Then
        Put #nInputFile, lCurrentRecord, sFileBuffer
    End If
    sFileBuffer = Space(BUFFER_SIZE)
    lCurrentRecord = lCurrentRecord + 1
    Get #nInputFile, lCurrentRecord, sFileBuffer
    bBufferModefied = False
    
    If lResidualLength > 0 Then
        sInputBuffer = Right(sInputBuffer, lResidualLength)
    Else
        sInputBuffer = ""
    End If
    
    If EOF(nInputFile) Then
        sInputBuffer = sInputBuffer & fnTrim(sFileBuffer)
        bFileEnd = True
    Else
        sInputBuffer = sInputBuffer & sFileBuffer
    End If
    If Asc(Left(sFileBuffer, 1)) = &HA Then
        nInputPointer = nInputPointer + 1
    End If
End Function

Private Function fnTrim(sSrc As String) As String
    Dim i As Integer
    
    i = Len(sSrc)
    Do While i > 0
        If Asc(Mid(sSrc, i, 1)) <> 0 Then
            Exit Do
        End If
        i = i - 1
    Loop
    fnTrim = (Left(sSrc, i))
End Function

Public Function GetLine(Optional fixed As Long) As String
    ' Added optional fixed to get fixed length lines
    Dim nPos As Integer
    Dim nLen As Integer
    
    GetLine = ""
    If fixed > 0 Then
        If (Len(sInputBuffer) - nInputPointer) + 1 >= fixed Then
            nPos = nInputPointer + fixed
        End If
    Else
        If nPos = 0 Then
            nPos = InStr(nInputPointer, sInputBuffer, Chr(13))
        End If
        If nPos = 0 Then
            nPos = InStr(nInputPointer, sInputBuffer, Chr(10))
        End If
        If nPos = 0 Then
            nPos = InStr(nInputPointer, sInputBuffer, Chr(12))  'Page break
        End If
    End If
    If nPos = 0 Then
        nLen = Len(sInputBuffer)
        If EOF(nInputFile) Then
            nPos = nLen - nInputPointer
            If nPos > 0 Then
                GetLine = Right(sInputBuffer, nPos)
                nInputPointerOld = nInputPointer
                nInputPointer = nLen
            End If
            Exit Function
        Else
            While nPos = 0 And Not EOF(nInputFile)
                lResidualLength = Len(sInputBuffer) - nInputPointer + 1
                subGetFromFile
                If Len(sInputBuffer) > 0 Then
                    If Asc(Left(sInputBuffer, 1)) = &HA And fixed = 0 Then
                        nInputPointer = 2
                    Else
                        nInputPointer = 1
                    End If
                    nInputPointerOld = nInputPointer
                    '
                    If fixed > 0 Then
                        If (Len(sInputBuffer) - nInputPointer) + 1 >= fixed Then
                            nPos = nInputPointer + fixed
                        End If
                    Else
                        If nPos = 0 Then
                            nPos = InStr(nInputPointer, sInputBuffer, Chr(13))
                        End If
                        If nPos = 0 Then
                            nPos = InStr(nInputPointer, sInputBuffer, Chr(10))
                        End If
                        If nPos = 0 Then
                            nPos = InStr(nInputPointer, sInputBuffer, Chr(12))
                        End If
                    End If
                End If
            Wend
        End If
    End If
    nLen = nPos - nInputPointer
    If nLen > 0 Then
        GetLine = Mid(sInputBuffer, nInputPointer, nLen)
    End If
    nInputPointerOld = nInputPointer
    
    If fixed > 0 Then
        nInputPointer = nPos
    Else
        'Add by JQ to check the nPOS must not equl to 0
        If nPos <> 0 And nPos + 1 <= Len(sInputBuffer) Then
            
            If Asc(Mid(sInputBuffer, nPos, 1)) = &HD And Asc(Mid(sInputBuffer, nPos + 1, 1)) = &HA Then
                nInputPointer = nPos + 2
            Else
                nInputPointer = nPos + 1
            End If
        Else
            nInputPointer = nPos + 1
        End If
    End If
End Function

'Public Function GetLine() As String
'    Dim nPos As Integer
'    Dim nLen As Integer
'
'    GetLine = ""
'    nPos = InStr(nInputPointer, sInputBuffer, Chr(13))
'    If nPos = 0 Then
'        nPos = InStr(nInputPointer, sInputBuffer, Chr(10))
'    End If
'    If nPos = 0 Then
'        nPos = InStr(nInputPointer, sInputBuffer, Chr(12))  'Page break
'    End If
'    If nPos = 0 Then
'        nLen = Len(sInputBuffer)
'        If EOF(nInputFile) Then
'            nPos = nLen - nInputPointer
'            If nPos > 0 Then
'                GetLine = Right(sInputBuffer, nPos)
'                nInputPointerOld = nInputPointer
'                nInputPointer = nLen
'            End If
'            Exit Function
'        Else
'            While nPos = 0 And Not EOF(nInputFile)
'                lResidualLength = Len(sInputBuffer) - nInputPointer + 1
'                subGetFromFile
'                If Len(sInputBuffer) > 0 Then
'                    If Asc(Left(sInputBuffer, 1)) = &HA Then
'                        nInputPointer = 2
'                    Else
'                        nInputPointer = 1
'                    End If
'                    nInputPointerOld = nInputPointer
'                    nPos = InStr(nInputPointer, sInputBuffer, Chr(13))
'                    If nPos = 0 Then
'                        nPos = InStr(nInputPointer, sInputBuffer, Chr(10))
'                    End If
'                End If
'            Wend
'        End If
'    End If
'    nLen = nPos - nInputPointer
'    If nLen > 0 Then
'        GetLine = Mid(sInputBuffer, nInputPointer, nLen)
'    End If
'    nInputPointerOld = nInputPointer
'
'    'Add by JQ to check the nPOS must not equl to 0
'    If nPos <> 0 And nPos + 1 <= Len(sInputBuffer) Then
'
'        If Asc(Mid(sInputBuffer, nPos, 1)) = &HD And Asc(Mid(sInputBuffer, nPos + 1, 1)) = &HA Then
'            nInputPointer = nPos + 2
'        Else
'            nInputPointer = nPos + 1
'        End If
'    Else
'        nInputPointer = nPos + 1
'    End If
'End Function

Public Function PrepareFile(sName As String) As Boolean

    nInputFile = FreeFile
    On Error GoTo errOpen
    Open sName For Input As #nInputFile
    Close #nInputFile
    Open sName For Random As #nInputFile Len = BUFFER_SIZE
    bFileEnd = False
    lCurrentRecord = 0
    lResidualLength = 0
    subGetFromFile
    nInputPointer = 1
    m_sFileName = sName
    PrepareFile = True
    Exit Function
errOpen:
    PrepareFile = False
End Function


Private Sub Class_Terminate()
    CloseFile
End Sub

