VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsVBReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit
Private m_nBottomMargin As Integer
Private m_sStubFont As String
Private m_nPrintFontName As String
Private m_nPrintFontSize As Integer
Private m_nPageLength As Integer
Private sStoreFont As String
Private nLeftMargin As Integer
Private nTextHeight As Integer
Private nPrevPage As Integer
Private nPageNumber As Integer
Private m_sRptHeader As String
Private m_sProgID As String
Private m_sSubHeader As String
Private m_sColHeader As String
Private m_sCompanyName As String
Private m_bPrintHeaderEachPage As Boolean

Enum eNumPrintDirection
    PRORPortrait = vbPRORPortrait
    PRORLandscape = vbPRORLandscape
End Enum

Private m_nDirection As eNumPrintDirection

Property Let CompanyName(sValue As String)
    m_sCompanyName = sValue
End Property

Property Get CompanyName() As String
    CompanyName = m_sCompanyName
End Property

Property Let Direction(eNumValue As eNumPrintDirection)
    m_nDirection = eNumValue
End Property

Property Get Direction() As eNumPrintDirection
    Direction = m_nDirection
End Property

Property Let BOTTOM_MARGIN(nValue As Integer)
    m_nBottomMargin = nValue
End Property

Property Get BOTTOM_MARGIN() As Integer
    BOTTOM_MARGIN = m_nBottomMargin
End Property

Property Let STUB_FONT(sValue As String)
    m_sStubFont = sValue
End Property

Property Get STUB_FONT() As String
    STUB_FONT = m_sStubFont
End Property

Property Let PRINT_FONTNAME(sValue As String)
    m_nPrintFontName = sValue
End Property

Property Get PRINT_FONTNAME() As String
    PRINT_FONTNAME = m_nPrintFontName
End Property

Property Let PRINT_FONTSIZE(nValue As Integer)
    m_nPrintFontSize = nValue
End Property

Property Get PRINT_FONTSIZE() As Integer
    PRINT_FONTSIZE = m_nPrintFontSize
End Property

Property Let PageLength(nValue As Integer)
    m_nPageLength = nValue
End Property

Property Get PageLength() As Integer
    PageLength = m_nPageLength
End Property

Property Let ReportHeader(sValue As String)
    m_sRptHeader = sValue
End Property

Property Get ReportHeader() As String
    ReportHeader = m_sRptHeader
End Property

Property Let ProgramID(sValue As String)
    m_sProgID = sValue
End Property

Property Get ProgramID() As String
    ProgramID = m_sProgID
End Property

Property Let ReportSubHeader(sValue As String)
    m_sSubHeader = sValue
End Property

Property Get ReportSubHeader() As String
     ReportSubHeader = m_sSubHeader
End Property

Property Let ColumnHeader(sValue As String)
    m_sColHeader = sValue
End Property

Property Get ColumnHeader() As String
    ColumnHeader = m_sColHeader
End Property

Public Sub PrintDefaultRptHeader()
    Dim sCompanyName As String
    Dim sRundate As String
    Dim sRuntime As String
    Dim nSpc As Integer
    Dim sReportLine As String, sPageNum As String
    
    sRundate = Format(Date, "MM/DD/YYYY")
    sRuntime = Format(Time, "HH:MM AMPM")

    sPageNum = "Page No.  " + fnRightJustified(nPageNumber, "####")
    nSpc = (PageLength - Len(Trim(CompanyName))) / 2 - Len("Program ID: " & ProgramID)
    sReportLine = "Program ID: " + ProgramID + Space(nSpc) + CompanyName
    sReportLine = sReportLine & Space(PageLength - Len(sReportLine) - Len(sPageNum)) & sPageNum
    
    SendTextToPrinter sReportLine
    
    nSpc = (PageLength - Len(Trim(ReportHeader))) / 2 - Len(Trim(sRundate & "Run Date: "))
    sReportLine = "Run Date: " + sRundate + Space(nSpc) + ReportHeader
    SendTextToPrinter sReportLine

    nSpc = (PageLength - Len(ReportSubHeader)) / 2 - Len(Trim(sRuntime & "Run Time: "))
    sReportLine = "Run Time: " + sRuntime + Space(nSpc) + ReportSubHeader
    SendTextToPrinter sReportLine
    
    SendTextToPrinter Space(PageLength)
    SendTextToPrinter ColumnHeader
End Sub

Public Function InitPrinter(Optional vNextPage As Variant) As Boolean
    Dim sErrMsg As String
    
    On Error GoTo ErrInitPrinter
    
    Printer.Orientation = Direction
    Printer.FontName = STUB_FONT
    Printer.Print " "
    Printer.FontName = PRINT_FONTNAME
    Printer.FontSize = PRINT_FONTSIZE

    
    nLeftMargin = (Printer.ScaleWidth - Printer.TextWidth(Space(PageLength))) / 2
    nTextHeight = Printer.ScaleHeight - BOTTOM_MARGIN
    
    InitPrinter = True
    
    Exit Function
    
ErrInitPrinter:
    sErrMsg = "An error has occurred while initializing the Printer. Err Code: " & _
        Err.Number & ", Err Desc: " & Err.Description
    MsgBox "Called by InitPrinter, " & sErrMsg, vbExclamation
    Printer.EndDoc
End Function

Public Sub SendTextToPrinter(ByVal sOut As String)

    Printer.CurrentX = nLeftMargin
    
    If sOut <> "\page\" Then
        Printer.Print sOut
    End If
    
    If Printer.CurrentY >= nTextHeight Or sOut = "\page\" Then
        Printer.Print Space(100) & "< continued >"
        Printer.NewPage
        nPageNumber = nPageNumber + 1
        
        If PrintHeaderEachPage Then
            PrintDefaultRptHeader
        End If
        
    End If
    
End Sub

Public Sub EndPrintDocument()
    On Error Resume Next
    Printer.NewPage
    Printer.EndDoc
    Printer.FontName = sStoreFont
    Printer.Orientation = vbPRORPortrait
End Sub

Private Function fnRightJustified(ByVal sIn As String, sFormatString As String) As String
    fnRightJustified = Format(Format(sIn, sFormatString), String(Len(sFormatString), "@"))
End Function

Private Function fnCenter(sIn As String, nMaxLen As Integer) As String
    Dim nDiff As String, nSpcLeft As Integer
    
    nDiff = nMaxLen - Len(sIn)
    If nDiff >= nMaxLen Then
        fnCenter = sIn
        Exit Function
    End If
    
    nSpcLeft = Int(nDiff / 2)
    fnCenter = Space(nSpcLeft) + sIn + Space(nDiff - nSpcLeft)
End Function

Private Sub Class_Initialize()
    m_nBottomMargin = 600
    m_sStubFont = "Arial"
    m_nPrintFontName = "Courier New"     '"Courier 17*1"
    m_nPrintFontSize = 10
    m_nPageLength = 120
    m_nDirection = PRORLandscape
    sStoreFont = Printer.FontName
    nPageNumber = 1
End Sub

Private Sub Class_Terminate()
    Printer.FontName = sStoreFont
    Printer.Orientation = vbPRORPortrait
End Sub

Property Let PrintHeaderEachPage(bValue As Boolean)
    m_bPrintHeaderEachPage = bValue
End Property

Property Get PrintHeaderEachPage() As Boolean
    PrintHeaderEachPage = m_bPrintHeaderEachPage
End Property

Public Sub PrintEmptyLine()
    SendTextToPrinter Space(PageLength)
End Sub
