VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form HtmlReport 
   Caption         =   "Report Browser"
   ClientHeight    =   7620
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11715
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7620
   ScaleWidth      =   11715
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   7644
      Left            =   12
      TabIndex        =   0
      Top             =   -36
      Width           =   11712
      ExtentX         =   20659
      ExtentY         =   13483
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
Attribute VB_Name = "HtmlReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**************************************************************'
' Call:        373319-1
' Description: This is the form that will laungh the report server for
'              report syrtbchg
' Programmer:  Junsong Qiu
' Date:        02/25/2003
'**************************************************************'

Option Explicit

Private m_sURL As String
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
                                    Optional bBrowseIt As Boolean = True, _
                                    Optional bPrintIt As Boolean = True, _
                                    Optional sReportID As String = "", _
                                    Optional nOrientation As Integer = vbPRORPortrait)
    
    Dim sStatusbarMsg As String
    
    Dim sHeaderLine As String
    Dim nFileNum As Long
    Dim i As Long
    Dim bFileOpen As Boolean
    Dim sCompanyName As String
    Dim sRunDate As String
    Dim sRunTime As String
    Dim nSpc As Integer
    Dim sReportLine As String
    Dim sPageNum As String
    Dim lLineNum As Long
    Dim sHeader1 As String
    Dim sHeader2 As String
    Dim nPosi As Integer
    
    On Error Resume Next
    sStatusbarMsg = frmMain.ffraStatusbar.Caption
    frmMain.tfnSetStatusBarCorrect "Creating Report. Please wait..."
    On Error GoTo 0
    
    On Error GoTo EXITHERE

    Dim nPageWidth As Integer
    Dim nLinesPerPage As Integer
    
    If nOrientation = vbPRORPortrait Then
        nPageWidth = 104
        nLinesPerPage = 72
    Else
        nPageWidth = 132
        nLinesPerPage = 51
    End If
    
    'initialize member variables
    m_sURL = ""
    
    If sReportID = "" Then
        sReportID = UCase(App.Title)
    End If
    
    m_sURL = LCase(App.Path & "\" & sReportID) & ".html"
    nFileNum = FreeFile

    Open m_sURL For Output As #nFileNum
    
    bFileOpen = True
    
    sHeaderLine = "<HTML><STYLE> @media print { DIV.PAGEBREAK {page-break-before: always}}</STYLE><BODY style=""font-family: Courier New; font-size: 8pt""><PRE>"
    Print #nFileNum, sHeaderLine

    sRunDate = Format(Date, "MM/DD/YYYY")
    sRunTime = Format(Time, "HH:MM AMPM")
    sCompanyName = fnGetCompanyName()
    
    Dim sRptTitle1 As String
    Dim sRptTitle2 As String

    nPosi = InStr(sReportTitle, vbCrLf)
    If nPosi > 0 Then
        sRptTitle1 = Left(sReportTitle, nPosi - 1)
        sRptTitle2 = Mid(sReportTitle, nPosi + 2)
    Else
        sRptTitle1 = sReportTitle
        sRptTitle2 = ""
    End If
    
    Dim sColHeading1 As String
    Dim sColHeading2 As String

    nPosi = InStr(sColumnHeading, vbCrLf)
    If nPosi > 0 Then
        sColHeading1 = Left(sColumnHeading, nPosi - 1)
        sColHeading2 = Mid(sColumnHeading, nPosi + 2)
    Else
        sColHeading1 = sReportTitle
        sColHeading2 = ""
    End If

    'sPageNum = "Page No.  " + fnRightJustified(nPageNumber, "####")
    nSpc = (nPageWidth - Len(sCompanyName)) / 2 - Len("Program ID: " & sReportID)
    sHeaderLine = "Program ID: " + sReportID + Space(nSpc) + sCompanyName
    sHeaderLine = sHeaderLine & Space(nPageWidth - Len(sHeaderLine)) ' - Len(sPageNum)) & sPageNum
    sHeaderLine = "<B>" & sHeaderLine & "</B>"
    Print #nFileNum, sHeaderLine
    
    nSpc = (nPageWidth - Len(sRptTitle1)) / 2 - Len(Trim(sRunDate & "Run Date: "))
    sHeaderLine = "Run Date: " + sRunDate + Space(nSpc) + sRptTitle1
    sHeaderLine = "<B>" & sHeaderLine & "</B>"
    Print #nFileNum, sHeaderLine

    nSpc = (nPageWidth - Len(sRptTitle2)) / 2 - Len(Trim(sRunTime & "Run Time: "))
    sHeaderLine = "Run Time: " + sRunTime + Space(nSpc) + sRptTitle2
    sHeaderLine = "<B>" & sHeaderLine & "</B>"
    Print #nFileNum, sHeaderLine

    Print #nFileNum, ""

    sHeaderLine = sColHeading1
    sHeaderLine = "<B>" & sHeaderLine & "</B>"
    sHeader1 = sHeaderLine
    Print #nFileNum, sHeaderLine

    sHeaderLine = sColHeading2
    sHeaderLine = "<B>" & sHeaderLine & "</B>"
    sHeader2 = sHeaderLine
    Print #nFileNum, sHeaderLine
    lLineNum = 6
    
    For i = 0 To UBound(aryBody)
        If lLineNum > nLinesPerPage Then
            lLineNum = 3
            Print #nFileNum, "<DIV CLASS=""PAGEBREAK""></DIV>"
            Print #nFileNum, sHeader1
            Print #nFileNum, sHeader2
            Print #nFileNum, ""
        End If

        sHeaderLine = aryBody(i)
        Print #nFileNum, sHeaderLine
        lLineNum = lLineNum + 1
    Next

    sHeaderLine = "</PRE></body></html>"
    Print #nFileNum, sHeaderLine

    Close #nFileNum
    bFileOpen = False

    If bBrowseIt Then
        BrowseInternetReport m_sURL
        
        If bPrintIt Then
            PrintInternetReport
        End If
    End If
    
EXITHERE:
    If bFileOpen Then
        Close #nFileNum
    End If
End Sub

Public Sub BrowseInternetReport(sURL As String)
    Const navNoHistory As Integer = 2
    Const navNoReadFromCache As Integer = 4
    Const navNoWriteToCache As Integer = 8

    Dim byteData() As Byte
    
    If sURL = "" Then
        Exit Sub
    End If
    
    'WebBrowser1.Navigate "http://ether4/cgi-bin/reportsrv.cgi?USERNAME=ssfactor&PASSWORD=menus&DATABASE=/factor/vbdev&CATEGORY=R&REPORTNAME=smrworpt&REPORTTITLE=Work%20Order%20Maintenance&REPORTEXE=smrrun01&REPORT=Work%20Order%20Maintenance|smrworpt|smmenu|&COLCOUNT=10&PRINTSEL=1&ENTRYTYPE=STANDARD|DATE|STANDARD|STANDARD|STANDARD|STANDARD|STANDARD|STANDARD|STANDARD|STANDARD|&ENTRYORDER=PARMa|PARMb|PARMc|PARMd|PARMe|PARMf|PARMg|PARMh|PARMi|PARMj|&PARMa=101&COMPRESS=1&ACTION=GO"
    WebBrowser1.Navigate sURL, navNoHistory + navNoReadFromCache + navNoWriteToCache, "_self", byteData
End Sub

Public Sub PrintInternetReport()
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
        WebBrowser1.ExecWB OLECMDID_PRINTPREVIEW, OLECMDEXECOPT_PROMPTUSER, Null, Null
    Else
        '"The internet Browser Print command is currently disabled."
        MsgBox "Can't print report right now."
    End If
    
End Sub

Private Function fnRightJustified(ByVal sIn As String, sFormatString As String) As String
    fnRightJustified = Format(Format(sIn, sFormatString), String(Len(sFormatString), "@"))
End Function

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
    Dim objDB As DataBase
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
