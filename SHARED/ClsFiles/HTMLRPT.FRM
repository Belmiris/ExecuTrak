VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form clsHtmlReport 
   Caption         =   "Report Browser"
   ClientHeight    =   8355
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11715
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8355
   ScaleWidth      =   11715
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   8340
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11685
      ExtentX         =   20611
      ExtentY         =   14711
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
End
Attribute VB_Name = "clsHtmlReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**************************************************************'
' Call:         447166
' Description:  This form is used to preview/print the html style
'               report
' Programmer:   David Chai,
'               Clone the Form (frmBrowser) that was done by Junsong Qiu
' Date:         05/26/2004
'**************************************************************'

Option Explicit
    
Private m_sReportID As String
Private m_sURL As String
Private m_nFileNum As Long
Private m_sCompanyName As String
Private m_sRunDate As String
Private m_sRunTime As String
Private m_sRptTitle1 As String
Private m_sRptTitle2 As String
Private m_sColHeading1 As String
Private m_sColHeading2 As String
'

Public Function FormatRptCol(sData As String, nLength As Integer, Optional bRightAdjustment As Boolean = False) As String
    If Len(Trim(sData)) >= nLength Then
        FormatRptCol = Mid(Trim(sData), 1, nLength)
    Else
        If bRightAdjustment Then
            FormatRptCol = Space(nLength - Len(Trim(sData))) & Trim(sData)
        Else
            FormatRptCol = Trim(sData) & Space(nLength - Len(Trim(sData)))
        End If
        
    End If
End Function

'functions to build the html report
Public Sub BuildInternetReport(frmMain As Form, _
                                    sReportTitle As String, _
                                    sColumnHeading As String, _
                                    aryBody() As String, _
                                    aryCriteria() As String, _
                                    Optional bBrowseIt As Boolean = True, _
                                    Optional bPrintIt As Boolean = True, _
                                    Optional sReportID As String = "", _
                                    Optional nOrientation As Integer = vbPRORPortrait, _
                                    Optional bCompressed As Boolean = False, _
                                    Optional bDoubleSpace As Boolean = False)
    
    Dim sHeaderLine As String
    
    Dim i As Long
    Dim bFileOpen As Boolean
    Dim sReportLine As String
    Dim lLineNum As Long
    
    On Error GoTo EXITHERE

    Dim nPageWidth As Integer
    Dim nLinesPerPage As Integer
    
    Screen.MousePointer = vbHourglass
    
    If nOrientation = vbPRORPortrait Then
        If bCompressed Then
            nPageWidth = 132
            nLinesPerPage = 90
        Else
            nPageWidth = 104
            nLinesPerPage = 72
        End If
    Else
        If fnVB_IDE() Then
            MsgBox "The landscape style report is not working properly because the program cannot " _
                + "set the Print Orientation automatically when Preview.", vbInformation
        End If
        
        If bCompressed Then
            nPageWidth = 232
            nLinesPerPage = 66
        Else
            nPageWidth = 132
            nLinesPerPage = 52
        End If
    End If
    
    'initialize member variables
    m_sURL = ""
    
    m_sReportID = sReportID
    
    If m_sReportID = "" Then
        m_sReportID = UCase(App.Title)
    End If
    
    m_sURL = LCase(App.Path & "\" & m_sReportID) & ".html"
    m_nFileNum = FreeFile

    Open m_sURL For Output As #m_nFileNum
    
    bFileOpen = True
    
    sHeaderLine = "<HTML><STYLE> @media print {DIV.PAGEBREAK {page-break-before: always}}</STYLE>"
    
    If bCompressed Then
        sHeaderLine = sHeaderLine + "<BODY style=""font-family:'Lucida Console';mso-bidi-font-family:'Courier New';font-size:6.3pt""><PRE>"
    Else
        sHeaderLine = sHeaderLine + "<BODY style=""font-family: Courier New; font-size: 8pt""><PRE>"
    End If
    
    Print #m_nFileNum, sHeaderLine
    
    subInitReportHeader sReportTitle, sColumnHeading
    
    'write report criteria page
    If aryCriteria(0) <> "" Then
        subReportHeader nPageWidth
        lLineNum = 4

        For i = 0 To UBound(aryCriteria)
            Screen.MousePointer = vbHourglass
            
            If lLineNum > nLinesPerPage Then
                Print #m_nFileNum, "<DIV CLASS=""PAGEBREAK""></DIV>"
                subReportHeader nPageWidth
                lLineNum = 4
            End If
    
            sHeaderLine = aryCriteria(i)
            Print #m_nFileNum, sHeaderLine
            lLineNum = lLineNum + 1
            Screen.MousePointer = vbHourglass
        Next
    
        Print #m_nFileNum, "<DIV CLASS=""PAGEBREAK""></DIV>"
    End If
    
    'write report details
    Screen.MousePointer = vbHourglass
        
    subReportHeader nPageWidth
    subColumnHeader
    lLineNum = 7
    
    For i = 0 To UBound(aryBody)
        Screen.MousePointer = vbHourglass
        
        If lLineNum > nLinesPerPage Then
            Print #m_nFileNum, "<DIV CLASS=""PAGEBREAK""></DIV>"
            subColumnHeader
            lLineNum = 3
        End If

        sHeaderLine = aryBody(i)
        Print #m_nFileNum, sHeaderLine
        lLineNum = lLineNum + 1
        
        'double spacing
        If bDoubleSpace Then
            'skip a blank line
            sHeaderLine = ""
            Print #m_nFileNum, sHeaderLine
            lLineNum = lLineNum + 1
        End If
        
        Screen.MousePointer = vbHourglass
    Next

    sHeaderLine = "</PRE></body></html>"
    Print #m_nFileNum, sHeaderLine

    Close #m_nFileNum
    bFileOpen = False

    If bBrowseIt Then
        Screen.MousePointer = vbHourglass
        BrowseInternetReport m_sURL
        
        If bPrintIt Then
            Screen.MousePointer = vbHourglass
            PrintInternetReport
        End If
    End If
    
EXITHERE:
    
    If bFileOpen Then
        Close #m_nFileNum
    End If

    Screen.MousePointer = vbDefault
End Sub

Public Sub BrowseInternetReport(sURL As String)
    Const navNoHistory As Integer = 2
    Const navNoReadFromCache As Integer = 4
    Const navNoWriteToCache As Integer = 8

    If sURL = "" Then
        Exit Sub
    End If
    
    'WebBrowser1.Navigate "http://ether4/cgi-bin/reportsrv.cgi?USERNAME=ssfactor&PASSWORD=menus&DATABASE=/factor/vbdev&CATEGORY=R&REPORTNAME=smrworpt&REPORTTITLE=Work%20Order%20Maintenance&REPORTEXE=smrrun01&REPORT=Work%20Order%20Maintenance|smrworpt|smmenu|&COLCOUNT=10&PRINTSEL=1&ENTRYTYPE=STANDARD|DATE|STANDARD|STANDARD|STANDARD|STANDARD|STANDARD|STANDARD|STANDARD|STANDARD|&ENTRYORDER=PARMa|PARMb|PARMc|PARMd|PARMe|PARMf|PARMg|PARMh|PARMi|PARMj|&PARMa=101&COMPRESS=1&ACTION=GO"
    WebBrowser1.Navigate sURL, navNoHistory + navNoReadFromCache + navNoWriteToCache, "_self"
End Sub

Public Sub PrintInternetReport(Optional bPreview As Boolean = True)
    Dim oleTemp As OLECMDF
    Dim i As Integer
    Dim bPrintIt As Boolean
    Const TIME_OUT_SECND = 120
    Const CHK_INTERVAL = 2
    
    On Error Resume Next
       
    bPrintIt = False
    i = 0
    
    Do While True
        oleTemp = WebBrowser1.QueryStatusWB(OLECMDID_PRINTPREVIEW)
        
        If oleTemp And OLECMDF_ENABLED Then
            bPrintIt = True
            Exit Do
        End If
        
        If i > TIME_OUT_SECND / CHK_INTERVAL Then
            '# time out
            Exit Do
        End If
        
        i = i + 1
        tfnWaitSeconds CHK_INTERVAL
    Loop
    
    If bPrintIt Then
        'WebBrowser1.ExecWB OLECMDID_PAGESETUP, OLECMDEXECOPT_DONTPROMPTUSER, Null, Null
        If bPreview Then
            WebBrowser1.ExecWB OLECMDID_PRINTPREVIEW, OLECMDEXECOPT_PROMPTUSER, 0, 0  'Null, Null
        Else
            WebBrowser1.ExecWB OLECMDID_PRINT, OLECMDEXECOPT_DONTPROMPTUSER, 0, 0
        End If
    Else
        '"The internet Browser Print command is currently disabled."
        MsgBox "Can't print report right now."
    End If
    
End Sub

Private Function fnRightJustified(ByVal sIn As String, sFormatString As String) As String
    fnRightJustified = Format(Format(sIn, sFormatString), String(Len(sFormatString), "@"))
End Function

Private Sub subInitReportHeader(sReportTitle As String, sColumnHeading As String)
    Dim nPosi As Integer
    
    m_sRunDate = Format(Date, "MM/DD/YYYY")
    m_sRunTime = Format(Time, "HH:MM AMPM")
    m_sCompanyName = fnGetCompanyName()

    nPosi = InStr(sReportTitle, vbCrLf)
    If nPosi > 0 Then
        m_sRptTitle1 = Left(sReportTitle, nPosi - 1)
        m_sRptTitle2 = Mid(sReportTitle, nPosi + 2)
    Else
        m_sRptTitle1 = sReportTitle
        m_sRptTitle2 = ""
    End If

    nPosi = InStr(sColumnHeading, vbCrLf)
    If nPosi > 0 Then
        m_sColHeading1 = Left(sColumnHeading, nPosi - 1)
        m_sColHeading2 = Mid(sColumnHeading, nPosi + 2)
    Else
        m_sColHeading1 = sReportTitle
        m_sColHeading2 = ""
    End If
End Sub

Private Sub subReportHeader(nPageWidth As Integer)
    Dim nSpc As Integer
    Dim sHeaderLine As String
    
    On Error Resume Next
    nSpc = (nPageWidth - Len(m_sCompanyName)) / 2 - Len("Program ID: " & m_sReportID)
    sHeaderLine = "Program ID: " + m_sReportID + Space(nSpc) + m_sCompanyName
    sHeaderLine = sHeaderLine & Space(nPageWidth - Len(sHeaderLine))
    sHeaderLine = "<B>" & sHeaderLine & "</B>"
    Print #m_nFileNum, sHeaderLine
    
    nSpc = (nPageWidth - Len(m_sRptTitle1)) / 2 - Len(Trim(m_sRunDate & "Run Date: "))
    sHeaderLine = "Run Date: " + m_sRunDate + Space(nSpc) + m_sRptTitle1
    sHeaderLine = "<B>" & sHeaderLine & "</B>"
    Print #m_nFileNum, sHeaderLine

    nSpc = (nPageWidth - Len(m_sRptTitle2)) / 2 - Len(Trim(m_sRunTime & "Run Time: "))
    sHeaderLine = "Run Time: " + m_sRunTime + Space(nSpc) + m_sRptTitle2
    sHeaderLine = "<B>" & sHeaderLine & "</B>"
    Print #m_nFileNum, sHeaderLine

    Print #m_nFileNum, ""
End Sub

Private Sub subColumnHeader()
    Dim sHeaderLine As String
    
    sHeaderLine = "<B>" & m_sColHeading1 & "</B>"
    Print #m_nFileNum, sHeaderLine

    sHeaderLine = "<B>" & m_sColHeading2 & "</B>"
    Print #m_nFileNum, sHeaderLine
    
    Print #m_nFileNum, ""
End Sub

Private Function fnGetCompanyName() As String
    Dim strSQL As String
    Dim rsTemp As Recordset
    
    strSQL = "select con_name from co_company_name"
    Set rsTemp = fnOpenRecord(strSQL)
    
    If Not rsTemp Is Nothing Then
        If rsTemp.RecordCount > 0 Then
            fnGetCompanyName = Trim(rsTemp!con_name & "")
        End If
    End If
End Function

Private Function fnOpenRecord(strSQL As String, Optional vDB As Variant) As Recordset
    Const SUB_NAME = "fnOpenRecord"
    
    ' Get records from the given SQL statement
    Dim objDB As Database
    Dim rsTemp As Recordset

    If IsMissing(vDB) Then
        Set objDB = t_dbMainDatabase
    Else
        Set objDB = vDB
    End If
    
    On Error GoTo SQLError
    If objDB Is t_dbMainDatabase Then
        Set rsTemp = objDB.OpenRecordset(strSQL, dbOpenSnapshot, dbSQLPassThrough)
    Else
        Set rsTemp = objDB.OpenRecordset(strSQL, dbOpenSnapshot)
    End If
    
    If rsTemp.RecordCount > 0 Then
        rsTemp.MoveLast
        rsTemp.MoveFirst
    End If
    
    Set fnOpenRecord = rsTemp

    On Error GoTo 0
    
    Exit Function

SQLError:
    Set fnOpenRecord = Nothing
End Function

Private Sub Form_Load()
    tfnCenterForm Me
End Sub

Private Function fnVB_IDE() As Boolean
    On Error GoTo errTrap
    Debug.Print 1 / 0
    Exit Function
    
errTrap:
    fnVB_IDE = True

End Function

