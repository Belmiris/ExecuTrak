VERSION 4.00
Begin VB.Form frmMain 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Copy Product Codes"
   ClientHeight    =   6348
   ClientLeft      =   912
   ClientTop       =   1572
   ClientWidth     =   8880
   BeginProperty Font 
      name            =   "Arial"
      charset         =   1
      weight          =   400
      size            =   9.6
      underline       =   0   'False
      italic          =   0   'False
      strikethrough   =   0   'False
   EndProperty
   Height          =   6672
   HelpContextID   =   1
   KeyPreview      =   -1  'True
   Left            =   864
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   6348
   ScaleWidth      =   8880
   Top             =   1296
   Width           =   8976
   Begin VB.CommandButton cmdTest 
      Caption         =   "Test 2"
      Height          =   372
      Left            =   7308
      TabIndex        =   4
      Top             =   5316
      Width           =   1104
   End
   Begin VB.CommandButton cmdDisplay 
      Caption         =   "Display"
      Height          =   408
      Left            =   4704
      TabIndex        =   3
      Top             =   6456
      Width           =   1296
   End
   Begin VB.ListBox lstOutput 
      Height          =   5016
      Left            =   48
      TabIndex        =   2
      Top             =   72
      Width           =   8760
   End
   Begin VB.CommandButton cmdEnd 
      Caption         =   "End"
      Height          =   420
      Left            =   7284
      TabIndex        =   1
      Top             =   5748
      Width           =   1140
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "Start"
      Height          =   420
      Left            =   6024
      TabIndex        =   0
      Top             =   5748
      Width           =   1212
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Option Explicit

    Private colFMTaxRecords As Collection

    Private EDIEngine As New EDIToolEngine

Private Function fnAppenCriteria(sField As String, _
                                 sCriteria As String) As String

    Dim sTemp() As String
    Dim nFlags() As Integer
    Dim sResult1 As String
    Dim sResult2 As String
    Dim nSize As Integer
    Dim i As Integer
    Dim nPos As Integer
    Dim sOR As String
    
    subParseString sTemp, sCriteria
    nSize = UBound(sTemp)
    ReDim nFlags(nSize)
    sResult1 = ""
    sOR = ""
    For i = 0 To nSize
        nPos = InStr(sTemp(i), "*")
        If nPos > 0 Then
            Mid(sTemp(i), nPos, 1) = "%"
        Else
            nPos = InStr(sTemp(i), "%")
        End If
        If nPos > 0 Then
            nFlags(i) = 1
            sResult1 = sResult1 & sOR & "(" & sField & " LIKE " & tfnSQLString(sTemp(i)) & ")"
            sOR = " OR "
        End If
        nPos = InStr(sTemp(i), "?")
        If nPos > 0 Then
            nFlags(i) = 1
            sResult1 = sResult1 & sOR & "(" & sField & " MATCHES " & tfnSQLString(sTemp(i)) & ")"
            sOR = " OR "
        End If
        nPos = InStr(sTemp(i), ":")
        If nPos > 0 Then
            nFlags(i) = 1
            If nPos > 1 Then
                sResult1 = sResult1 & sOR & "((" & sField & " >= " & tfnSQLString(Left(sTemp(i), nPos - 1)) & ")"
                nPos = Len(sTemp(i)) - nPos
                If nPos > 0 Then
                    sResult1 = sResult1 & " AND (" & sField & " <= " & tfnSQLString(Right(sTemp(i), nPos)) & "))"
                Else
                    sResult1 = sResult1 & ")"
                End If
            Else
                nPos = Len(sTemp(i)) - nPos
                If nPos > 0 Then
                    sResult1 = sResult1 & "(" & sField & " <= " & tfnSQLString(Right(sTemp(i), nPos)) & ")"
                End If
            End If
            sOR = " OR "
        End If
    Next i
    sResult2 = sField & " IN ("
    sOR = ""
    For i = 0 To nSize
        If nFlags(i) = 0 Then
            sResult2 = sResult2 & sOR & tfnSQLString(sTemp(i))
            sOR = ", "
        End If
    Next i

    If sOR = "" Then
        fnAppenCriteria = "(" & sResult1 & ")"
    Else
        If sResult1 = "" Then
            fnAppenCriteria = sResult2 & ")"
        Else
            fnAppenCriteria = "(" & sResult1 & " OR " & sResult2 & "))"
        End If
    End If
    
End Function

Private Sub subParseString(sParam() As String, _
                           sSrc As String, _
                           Optional vStart As Variant, _
                           Optional vEnd As Variant)
                          
    If Trim(sSrc) = "" Then
        Exit Sub
    End If

    Const nArrayInc As Integer = 5
    Dim i1 As Integer
    Dim i2 As Integer
    Dim k As Integer
    Dim nEnd As Integer
    If IsMissing(vStart) Then
        i1 = 1
    Else
        i1 = vStart
    End If
    If IsMissing(vEnd) Then
        nEnd = Len(sSrc)
    Else
        nEnd = vEnd
    End If
    If i1 < 1 Then i1 = 1
    i2 = 1
    k = 0
    ReDim sParam(nArrayInc)
    While i1 <= nEnd And i2 > 0 And i2 <= nEnd
        i2 = InStr(i1, sSrc, ",")
        If i2 >= i1 And i2 <= nEnd Then
            If k > UBound(sParam) Then
                ReDim Preserve sParam(k + nArrayInc)
            End If
            sParam(k) = Trim$(Mid$(sSrc, i1, i2 - i1))
            k = k + 1
            i1 = i2 + 1
        End If
    Wend
    If i2 <= nEnd Then
        If k > UBound(sParam) Then
            ReDim Preserve sParam(k + nArrayInc)
        End If
        sParam(k) = Trim$(Mid$(sSrc, i1, nEnd - i1 + 1))
        ReDim Preserve sParam(k)
    Else
        If k > 0 Then
            sParam(k - 1) = Trim$(Mid$(sSrc, i1, nEnd - i1 + 1))
            ReDim Preserve sParam(k - 1)
        End If
    End If
End Sub

Private Sub cmdDisplay_Click()
    Const Sep = ",    "
    Dim ciTemp As ETFMTaxRecord
    Dim sTemp As String
    
    lstOutput.Clear
    
    If Not colFMTaxRecords Is Nothing Then
    For Each ciTemp In colFMTaxRecords
        With ciTemp
            sTemp = "Sched  = " & .TaxScheduleCode
            sTemp = sTemp & Sep & "PCode = " & .ProductCode
            sTemp = sTemp & Sep & "CarName = " & .CarrierName
            sTemp = sTemp & Sep & "CarID = " & .CarrierFedID
            lstOutput.AddItem sTemp
            sTemp = ""
            sTemp = sTemp & Sep & "OCity = " & .OriginCity
            sTemp = sTemp & Sep & "OState = " & .OriginState
            sTemp = sTemp & Sep & "DCity = " & .DestCity
            sTemp = sTemp & Sep & "DState = " & .DestState
            lstOutput.AddItem sTemp
            sTemp = ""
            sTemp = sTemp & Sep & "BSName = " & .BuySellName
            sTemp = sTemp & Sep & "BSID = " & .BuySellFedID
            sTemp = sTemp & Sep & "TYPE = " & .DetailType
            sTemp = sTemp & Sep & "BOL = " & .LadingNumber
            lstOutput.AddItem sTemp
            sTemp = ""
            sTemp = sTemp & Sep & "Date = " & .LadingDate
            sTemp = sTemp & Sep & "Qty = " & .Quantity
            sTemp = sTemp & Sep & "Movement = " & .Movement
        End With
        lstOutput.AddItem sTemp
        lstOutput.AddItem ""
    Next
    End If
End Sub

Private Sub cmdEnd_Click()
    Dim sTemp() As String

    End
End Sub

Private Sub cmdOpen_Click()
    
    Dim sStartDate As String
    Dim sEndDate As String
    Dim sSchedule As String
    Dim sClass As String
    Dim sType As String
    Dim nTypeCode As String
    
    Screen.MousePointer = vbHourglass
    sType = EDIEngine.EDITypeD
    sType = EDIEngine.EDITypeR
    If sType = EDIEngine.EDITypeR Then
        'For Sample database
        sStartDate = "3/1/98"
        sEndDate = "3/31/98"
'        sSchedule = "O????"
        sSchedule = "AR??"
'        sSchedule = "720DY"
        sClass = "GAS, DIESL"
'        sClass = "DIESL"
        nTypeCode = "1"
    Else
        'For gasup database
        sStartDate = "2/1/97"
        sEndDate = "2/28/97"
        sSchedule = "X*, O*"
        sClass = "GAS"
        nTypeCode = "5I "
    End If
    
    With EDIEngine
        .StartDate = sStartDate
        .EndDate = sEndDate
        .EDIType = sType
        .TaxpayerLicense = "NoLicense"
        .TheirID = "THEIRID"
        .Mode = .ModeTest
        .FileName = "TEST.TXT"  '"C:\VBDEV\EDITOOL\TEST.TXT"
        .BeginTrans
        .AddEDIRecords sSchedule, nTypeCode, sClass, "124"
        .EndTrans
    End With
    
    Screen.MousePointer = vbDefault

        MsgBox EDIEngine.LastMessage

End Sub


Private Sub cmdTest_Click()
    'Test sitatuation with different tax type code
    
    Dim sStartDate As String
    Dim sEndDate As String
    Dim sSchedule As String
    Dim sClass As String
    Dim sType As String
    Dim nTypeCode As String
    
    Screen.MousePointer = vbHourglass
    sType = EDIEngine.EDITypeD
    sType = EDIEngine.EDITypeR
    If sType = EDIEngine.EDITypeR Then
        'For Sample database
        sStartDate = "3/1/98"
        sEndDate = "3/31/98"
'        sSchedule = "AR??"
        sSchedule = "AMB"
        nTypeCode = "1"
    End If
    
    With EDIEngine
        .StartDate = sStartDate
        .EndDate = sEndDate
        .EDIType = sType
        .TaxpayerLicense = "NoLicense"
        .TheirID = "THEIRID"
        .Mode = .ModeTest
        .FileName = "TEST.TXT"  '"C:\VBDEV\EDITOOL\TEST.TXT"
        
        .BeginTrans
        
        'First tax type for gas
        sSchedule = "AMB"
        sClass = "GAS"
        .TaxpayerLicense = "License1"
        .TaxTypeCode = "TXTP1"
        .BeginDetail
        .AddEDIRecords sSchedule, nTypeCode, sClass, "124"
        .EndDetail
        
        'second tax type for diesl
        sSchedule = "AR??"
        sClass = "DIESL"
        .TaxpayerLicense = "License2"
        .TaxTypeCode = "TXTP2"
        .BeginDetail
        .AddEDIRecords sSchedule, nTypeCode, sClass, "124"
        .EndDetail
        
        .EndTrans
    End With
    
    Screen.MousePointer = vbDefault

    MsgBox EDIEngine.LastMessage

End Sub


Private Sub Form_Load()

    If tfnAuthorizeExecute(Command) = False Then 'Check for handshake if not in the development mode
        End 'this check makes sure the application can only be run from the FACTMENU program
    End If
    
    If tfnOpenDatabase = False Then 'open the database, ODBC Dialog Box during developemnt, oleObject Connection String when not
        End 'exit if database fails to connect - message box is in the tfnOpenDatabase function
    End If
    subInitErrorHandler
End Sub


Private Sub subInitErrorHandler()
    Set objErrHandler = New clsErrorHandler
    With objErrHandler
        Set .FormParent = Me
        Set .DatabaseEngine = t_engFactor
        Set .LocalDatabase = tfnOpenLocalDatabase
    End With
End Sub






Private Function fnCentury(sDate As String) As String

    Dim dDate As Date
    If IsDate(sDate) Then
        dDate = CDate(sDate)
        fnCentury = CStr(Int(Year(dDate) / 100))
    End If
    
End Function

Private Function fnFixLength(sSrc As String, _
                             nLen As Integer) As String

    Dim nTemp As Integer
    
    nTemp = Len(sSrc)
    If nTemp < nLen Then
        fnFixLength = sSrc & Space(nLen - nTemp)
    Else
        fnFixLength = Left(sSrc, nLen)
    End If

End Function



