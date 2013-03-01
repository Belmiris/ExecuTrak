Attribute VB_Name = "modPrintStuff"
Option Explicit

Private m_charHeight As Long
Private m_charWidth As Long
Private m_columnHeader1 As String
Private m_columnHeader2 As String
Private m_columns As Long
Private m_columnSpacing As Long
Private m_columnWidths() As Long
Private m_gridArray() As Variant
Private m_header1 As String
Private m_header1Left As Long
Private m_header2 As String
Private m_header2Left As Long
Private m_headers() As String
Private m_marginBottom As Long
Private m_marginLeft As Long
Private m_marginRight As Long
Private m_marginTop As Long
Private m_pageCount As Long
Private m_rows As Long
Private m_spaceWidth As Long
Private m_title As String
Private m_totalColumnWidth As Long

Public Function PrintGridArray(GridArray() As Variant, headers() As String, Title As String) As Boolean
    Dim message$
    
    On Error GoTo FINISHED
        
    Initialize
    
    m_headers = headers
    m_gridArray = GridArray
    m_title = Title
        
    m_rows = GetRowCount()
    m_columns = GetColumnCount()
    If m_rows < 1 Or m_columns < 1 Then
        MsgBox "No Data To Print"
        Exit Function
    End If
    
    If UBound(m_headers) <> m_columns Then
        If (UBound(m_headers) < m_columns) Then
            m_columns = UBound(m_headers)
        Else
            MsgBox "Number of columns and headers do not match"
            Exit Function
        End If
    End If
    
    SetColumnWidths
    
    If Not SetOrientation Then
        MsgBox "Data is too wide to print"
        Exit Function
    End If
    
    SetMargins
        
    m_pageCount = EstimatePageCount()
    
    If m_pageCount = 1 Then
        message = "Is the printer '" & Printer.DeviceName & "' ready to print " & CStr(m_pageCount) & " page?"
    Else
        message = "Is the printer '" & Printer.DeviceName & "' ready to print " & CStr(m_pageCount) & " pages?"
    End If
    
    If vbYes = MsgBox(message, vbYesNo, "Print " & Title) Then
        SendToPrinter
        PrintGridArray = True
    End If
    
    Err.Clear
FINISHED:
    If Err.Number <> 0 Then
        MsgBox "Print Error: " & Err.Description
        Err.Clear
    End If
End Function

Private Sub Initialize()

    m_charHeight = 0
    m_charWidth = 0
    m_columnHeader1 = ""
    m_columnHeader2 = ""
    m_columns = 0
    m_columnSpacing = 0
    ReDim m_columnWidths(0) As Long
    m_header1 = ""
    m_header1Left = 0
    m_header2 = ""
    m_header2Left = 0
    ReDim m_headers(0) As String
    m_marginBottom = 0
    m_marginLeft = 0
    m_marginRight = 0
    m_marginTop = 0
    m_pageCount = 0
    m_rows = 0
    m_spaceWidth = 0
    m_title = ""
    m_totalColumnWidth = 0
    
    Printer.ScaleMode = vbTwips
    Printer.FontSize = 8
    Printer.FontName = "Tahoma"
    
    m_charWidth = Printer.TextWidth("W")
    m_charHeight = Printer.TextHeight("W")
    m_spaceWidth = Printer.TextWidth(" ")
    m_columnSpacing = m_charWidth * 2

End Sub

Private Function GetRowCount() As Long
    GetRowCount = -1
    
    On Error GoTo FINISHED
    GetRowCount = UBound(m_gridArray, 2)
    
FINISHED:
    Err.Clear
End Function

Private Function GetColumnCount() As Long
    GetColumnCount = -1
    
    On Error GoTo FINISHED
    GetColumnCount = UBound(m_gridArray, 1)
    
FINISHED:
    Err.Clear
End Function

Private Sub SetColumnWidths()
    Dim row As Long
    Dim col As Long
    Dim val As String
    Dim w As Long
    
    m_totalColumnWidth = 0
    ReDim m_columnWidths(m_columns)
    
    For col = 0 To m_columns - 1
        val = m_headers(col) & ""
        m_columnWidths(col) = Printer.TextWidth(val)
    Next col
    
    For row = 0 To m_rows - 1
        For col = 0 To m_columns
            val = m_gridArray(col, row) & ""
            w = Printer.TextWidth(val)
            
            If w > m_columnWidths(col) Then
                m_columnWidths(col) = w
            End If
        Next col
    Next row
    
    For col = 0 To m_columns - 1
        m_totalColumnWidth = m_totalColumnWidth + m_columnWidths(col) + m_columnSpacing
    Next
    
    m_totalColumnWidth = m_totalColumnWidth - m_columnSpacing
    
End Sub

Private Sub SetMargins()
        
    m_marginLeft = 0
    m_marginTop = 0
    m_marginRight = 0
    m_marginBottom = 0
    
    If m_totalColumnWidth < Printer.ScaleWidth Then
        m_marginLeft = (Printer.ScaleWidth - m_totalColumnWidth) / 2
    End If
    
End Sub

Private Function SetOrientation() As Boolean
    Dim col As Long
    Dim max As Long
       
    Printer.Orientation = vbPRORPortrait
    max = Printer.ScaleWidth - (m_marginLeft + m_marginRight + m_columnSpacing * m_columns)
    
    If m_totalColumnWidth <= max Then
        SetOrientation = True
        Exit Function
    End If
        
    Printer.Orientation = vbPRORLandscape
    max = Printer.ScaleWidth - (m_marginLeft + m_marginRight + m_columnSpacing * m_columns)
    
    If m_totalColumnWidth <= max Then
        SetOrientation = True
        Exit Function
    End If
    
End Function

Private Function EstimatePageCount() As Long
    Dim linesPerPage As Long
    Dim max As Double
    Dim ph As Double
    Dim th As Double
    Dim pc As Double
    
    If m_charHeight = 0 Then Exit Function
    If m_rows = 0 Or m_columns = 0 Then Exit Function
    
    ' Allow for 3 lines in Title, 2 lines in column headers, 3 lines in footer
    th = m_charHeight
    ph = Printer.ScaleHeight - (m_marginTop + m_marginBottom + (m_charHeight * 8))
    
    linesPerPage = ph / th - 0.5
    pc = m_rows / linesPerPage + 0.5
    If pc < 1 Then pc = 1
    
    EstimatePageCount = Format(pc, "0")
    
End Function

'********************************************************************
' PRINTING
'********************************************************************

Private Sub SendToPrinter()
    Dim row As Long
    Dim done As Boolean
    
    row = 0
    
    While Not done
        row = PrintPage(row)
        If (row < m_rows - 1) Then
            Printer.NewPage
        Else
            done = True
        End If
    Wend
    
    Printer.EndDoc
    
End Sub

Private Sub PrintHeader()
    Dim col As Long
    Dim Top, Left As Long
    Dim drawWidth As Long
    
    drawWidth = Printer.drawWidth
    Printer.drawWidth = 3
    
    If Len(m_header1) < 1 Then
        m_header1 = m_title
        m_header1Left = (Printer.ScaleWidth - Printer.TextWidth(m_header1)) / 2
    End If
    
    If Len(m_header2) < 1 Then
        m_header2 = "Report Date: " & CStr(Now)
        m_header2Left = (Printer.ScaleWidth - Printer.TextWidth(m_header2)) / 2
    End If
    
    Printer.CurrentX = m_header1Left
    Printer.CurrentY = m_marginTop
    Printer.Print m_header1
    
    Printer.CurrentX = m_header2Left
    Printer.Print m_header2
    
    Top = Printer.CurrentY + m_charHeight
    Left = m_marginLeft
    
    For col = 0 To m_columns - 1
        Printer.CurrentX = Left
        Printer.CurrentY = Top
        Printer.Print m_headers(col)
        Printer.Line (Left, Top + m_charHeight)-(Left + m_columnWidths(col), Top + m_charHeight)
        
        Left = Left + m_columnWidths(col) + m_columnSpacing
    Next col
    
End Sub

Private Function PrintPage(startRow As Long) As Long
    Dim col As Long
    Dim row As Long
    Dim max As Long
    Dim Left As Long
    Dim Top As Long
    Dim val As String
    
    PrintHeader
    
    If startRow < 0 Or startRow >= m_rows Then
        PrintPage = m_rows
        Exit Function
    End If
    
    max = Printer.ScaleHeight - m_charHeight * 3
    Top = m_marginTop + m_charHeight * 4
    
    For row = startRow To m_rows - 1
        Left = m_marginLeft
        
        For col = 0 To m_columns - 1
            Printer.CurrentX = Left
            Printer.CurrentY = Top
                    
            val = m_gridArray(col, row) & ""
            Printer.Print val
            Left = Left + m_columnWidths(col) + m_columnSpacing
        Next col
        
        Top = Top + m_charHeight
        If Top >= max Then
            Exit For
        End If
    Next row
        
    PrintFooter
    PrintPage = row + 1
    
End Function

Private Sub PrintFooter()
    Dim Left As Long
    Dim val As String
    
    val = "Page " + CStr(Printer.Page)
    Left = (Printer.ScaleWidth - Printer.TextWidth(val)) / 2
    
    Printer.CurrentY = Printer.ScaleHeight - (m_marginBottom + m_charHeight)
    Printer.CurrentX = Left
    Printer.Print val
    
End Sub
