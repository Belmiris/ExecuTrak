Attribute VB_Name = "modmyBASIC"
'*******************************************************
' Most of the following codes are from Weigong,
' some are added by Xitai
'
' Set myform = me in form_load
'
'*******************************************************

Option Explicit
' my messages ( Need to be updated to standard ones)
Public Const ADD_EDIT_MSG = "Select Add, Edit Or Exit."
Public Const DEL_WARNING_MSG = " Are you sure you want to delete the current record? "
Public Const DELALL_WARNING_MSG = "This will delete records in several screens." _
                & " Are you sure you want to delete them all?"
Public Const DELROW_WARNING_MSG = " Are you sure you want to delete the current row? "
Public Const INSERT_ERR_MSG = " Can not insert, Please try again later."
Public Const DELETE_ERR_MSG = " Can not delete, Please try again later."
Public Const UPDATE_ERR_MSG = " Can not update, Please try again later."
Public Const Edit_ERR_MSG = " No Record Available To Be Edited"
Public Const CAN_NOT_SWITCH_MSG = "Can not Switch Tab !Data is not valid in current screen."

'patterns
' positive smallint=short integer  (0 --32767)
Public Const PAT_SMALL_INT As String = "^(#{0,4}|[0-2]?#{0,4}|\3?([0-1]?#{0,3}|\2?([0-6]?#{0,2}|\7?([0-5]?#?|\6?[0-7]?))))$"
'positive integer=long (0 --- 2,147,483,647)
Public Const PAT_INTEGER As String = "^#{0,9}$ "
 '"^#{0,9}|[0-1]?#{0,9}|\2?(0?#{0,8}|\1?([0-3]?#{0,7}|\4?([0-6]?#{0,6}|(\7?([0-3]?#{0,5}|\4?([0-7]?#{0,4}|"

'varibles and constants
Public myForm As Form  ' "set myform=me" in form_load
Public Const DATA_INIT As Integer = 0
Public Const DATA_LOADED As Integer = 1
Public Const DATA_CHANGED As Integer = 2
Public nDataStatus As Integer 'data loaded,inti,changed flag
Public bUpdateTable As Boolean  'almost never used
Public Const SCROLL_BAR_WIDTH As Integer = 250
Public dbLocal As DataBase
Public Const nDB_LOCAL As Integer = 0
Public Const nDB_REMOTE As Integer = 1
Public Const DB_REMOTE = 1
Public Const DB_LOCAL = 0
Global t_engFactor2nd As DBEngine
Global t_wsWorkSpace2nd As Workspace
Global t_dbMainDatabase2nd As DataBase
'======================================================


'===for printer below====
Public hPrevFont As Integer
Public hCurrFont As Integer
Public Type myColumn
    columnID As Integer
    columnName As String
    ColumnType As String
    columnLength As Integer
End Type

Public Type myIndex
    indexName As String
    indexType As String
    indexFields(7) As String
    nFields As Integer
End Type

Public Type sysParm
    parm_field As String
    parm_desc As String
End Type
'==========================below from Ma' printer project==============

Public Declare Function GetDeviceCaps Lib "GDI32" ( _
   ByVal hDC As Long, ByVal iCapabilitiy As Long) As Long

'Private Declare Function GetDeviceCaps Lib "GDI" _
'   (ByVal hDC As Integer, ByVal nIndex As Integer) As Integer '2 integer

Public Type LOGFONT
    lfHeight As Long
    lfWidth As Long
    lfEscapement As Long
    lfOrientation As Long
    lfWeight As Long
    lfItalic As String * 1
    lfUnderline As String * 1
    lfStrikeOut As String * 1
    lfCharSet As String * 1
    lfOutPrecision As String * 1
    lfClipPrecision As String * 1
    lfQuality As String * 1
    lfPitchAndFamily As String * 1
    lfFaceName As String * 32 'LF_FACESIZE
End Type '5 integer

'Private Declare Function CreateFontIndirect Lib "gdi32" _
'        (lpLogFont As LOGFONT) As Long  'one integer

Public Declare Function DeleteObject Lib "GDI32" _
        (ByVal hObject As Long) As Long  '2 integer

'Private Declare Function SelectObject Lib "GDI" _
'        (ByVal hDC As Integer, _
'         ByVal hObject As Integer) As Integer
         
Public Declare Function SelectObject Lib "GDI32" ( _
   ByVal hDC As Long, ByVal hObject As Long) As Long
   
Public Declare Function CreateFontIndirect Lib "GDI32" _
    Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long
         
'================================above from Ma printer=============


'======================================================

'    Set dbLocal = tfnOpenLocalDatabase
'    If dbLocal Is Nothing Then
'        subEndThisProgram False
'    End If

Public Function fnGetField(x As Variant) As String
     If IsNull(x) Then
        fnGetField = ""
     Else
        fnGetField = Trim(x)
     End If
     
End Function
'=====================================================
'warning: any non-numeric will be zero if isNum is true
'it will turn nothing(but not NULL, not "") to 0 or ""
Public Function fnGetValue(x As Variant, isNum As Boolean) As String
    
    x = "" & x
    
    If Not isNum Then
        fnGetValue = x
    Else
        If Len(x) = 0 Then
            fnGetValue = 0
        Else
            If IsNumeric(x) Then
                fnGetValue = x
            Else
                fnGetValue = 0
            End If
        End If
    End If
    
    If isNum Then
        'fnGetValue = "0" & X
    Else
        'fnGetValue = "" & X
    End If
    
End Function
'=============================

Public Sub subInitERRHandler()

   If objErrHandler Is Nothing Then
       Set dbLocal = tfnOpenLocalDatabase()
       Set objErrHandler = New clsErrorHandler
       With objErrHandler
          Set .FormParent = myForm
          Set .DatabaseEngine = t_engFactor
          Set .LocalDatabase = dbLocal
       End With
   End If
End Sub
'=====================================================

Public Function MyStr(ByVal szParameter As Variant, Optional vNoQuotes) As String
   
' Properly quotes and formats an SQL string.  If vNoQuotes is present, the result WILL NOT BE QUOTED
' for each ' character found, insert a double ''.  Leave "%* alone
    
    Dim nIdx As Integer
    Dim nPos As Integer
    
    If IsNull(szParameter) Then
       szParameter = ""
    Else
       szParameter = Trim(szParameter)
    End If
    
    nIdx = 1
    nPos = InStr(nIdx, szParameter, "'")
    
    While nPos <> 0
        szParameter = Left(szParameter, nPos) & "'" & Right(szParameter, Len(szParameter) - nPos)
        nIdx = nPos + 2
        nPos = InStr(nIdx, szParameter, "'")
    Wend
    
    ' quote the whole string - optional
    If IsMissing(vNoQuotes) Then
        MyStr = "'" & szParameter & "'"
    Else
        MyStr = szParameter
    End If
End Function
'======================================================

Public Function GetFieldData(rs As Recordset, Optional szField As Variant) As String
   If IsMissing(szField) Then
        If Not IsNull(rs.Fields(0)) Then
            GetFieldData = Trim(rs.Fields(0))
        Else
            GetFieldData = ""
        End If
   Else
        If Not IsNull(rs.Fields(szField)) Then
           GetFieldData = Trim(rs.Fields(szField))
        Else
           GetFieldData = ""
        End If
   End If
End Function

' this function also returns record count
Public Function GetRecordSet(rsTemp As Recordset, szSql As String, _
                   Optional nDB As Variant, Optional szCalledFrom As Variant, _
                   Optional bShowErrow As Variant) As Long
    On Error GoTo SQLError
    
    If IsMissing(nDB) Then
       nDB = nDB_REMOTE
    End If
    
    Select Case nDB
       Case nDB_LOCAL
       
         Set rsTemp = dbLocal.OpenRecordset(szSql, dbOpenSnapshot)
         
        
       Case nDB_REMOTE
       
         Set rsTemp = t_dbMainDatabase.OpenRecordset(szSql, dbOpenSnapshot, dbSQLPassThrough)
         
    End Select
    If rsTemp.RecordCount > 0 Then
            rsTemp.MoveLast
            rsTemp.MoveFirst
    End If
    GetRecordSet = rsTemp.RecordCount
    Exit Function
SQLError:
    If IsMissing(szCalledFrom) Then
       szCalledFrom = ""
    End If
    If IsMissing(bShowErrow) Then
       bShowErrow = True
    End If
    GetRecordSet = -1
    tfnErrHandler "GetRecordSet," & szCalledFrom, szSql, bShowErrow
    On Error GoTo 0
End Function

' this function also returns record count
Public Function GetRecordCount(szSql As String, _
                   Optional nDB As Variant, Optional szCalledFrom As Variant, _
                   Optional bShowErrow As Variant) As Long
    
    Dim rsTemp As Recordset
    
    On Error GoTo SQLError
    
    If IsMissing(nDB) Then
       nDB = nDB_REMOTE
    End If
    
    Select Case nDB
       Case nDB_LOCAL
       
         Set rsTemp = dbLocal.OpenRecordset(szSql, dbOpenSnapshot)
         
         If rsTemp.RecordCount > 0 Then
            rsTemp.MoveLast
            rsTemp.MoveFirst
         End If
       Case nDB_REMOTE
       
         Set rsTemp = t_dbMainDatabase.OpenRecordset(szSql, dbOpenSnapshot, dbSQLPassThrough)
         If rsTemp.RecordCount > 0 Then
            rsTemp.MoveLast
            rsTemp.MoveFirst
         End If
    End Select
    GetRecordCount = rsTemp.RecordCount
    Exit Function
SQLError:
    If IsMissing(szCalledFrom) Then
       szCalledFrom = ""
    End If
    If IsMissing(bShowErrow) Then
       bShowErrow = True
    End If
    GetRecordCount = -1
    tfnErrHandler "GetRecordCount," & szCalledFrom, szSql, bShowErrow
    On Error GoTo 0
End Function
'=================================================

Public Sub subSetFocus(cntlTemp As Control)
    'Set focus to a textbox or a command button control
    Const nNumberOFTry  As Integer = 1
    Dim nCount As Integer
    nCount = 0
    On Error GoTo errSetFocus
    cntlTemp.SetFocus
extSetFocus:
    On Error GoTo 0
    Exit Sub
errSetFocus:
    If nCount < nNumberOFTry Then
        nCount = nCount + 1
        DoEvents
        Resume
    Else
        Resume extSetFocus
    End If
End Sub
'=====================================================

' the following subroutines depend on template form
Public Sub subEnableAdd(bYesNo As Boolean)
    With myForm
        .cmdAddBtn.Enabled = bYesNo
        .mnuAdd.Enabled = bYesNo
    End With
End Sub
'=====================================================

Public Sub subEnableAddTab(bYesNo As Boolean)
    With myForm
        .cmdAddBtn.Enabled = bYesNo
        '.mnuEdit.Enabled = bYesNo
    End With
End Sub
'=====================================================

Public Sub subEnableDeleteTAB(bYesNo As Boolean)
    With myForm
        .cmdDeleteBtn.Enabled = bYesNo
        '.mnudelete.Enabled = bYesNo
    End With
End Sub
'=====================================================

Public Sub subEnableDelete(bYesNo As Boolean)
    With myForm
        .cmdDeleteBtn.Enabled = bYesNo
        .mnuDelete.Enabled = bYesNo
    End With
End Sub
'=====================================================

Public Sub subEnableSearchBtn(cmdButton As Control, bYesNo As Boolean)
  
    If bYesNo Then
        cmdButton.Picture = frmContext.LoadPicture(SEARCH_UP)
    Else
        cmdButton.Picture = frmContext.LoadPicture(SEARCH_DOWN)
    End If
    cmdButton.Enabled = bYesNo
    
End Sub
'=====================================================

Public Sub subEnableEditTab(bYesNo As Boolean)
   With myForm
        .cmdEditBtn.Enabled = bYesNo
        .mnuEdit.Enabled = bYesNo
   End With
End Sub
'=====================================================

Public Sub subEnableEdit(bYesNo As Boolean)
   With myForm
        .cmdEditBtn.Enabled = bYesNo
        .mnuEdit.Enabled = bYesNo
   End With
End Sub
'=====================================================

Public Sub subEnableUpdateInsert(bYesNo As Boolean)
    myForm.cmdUpdateInsertBtn.Enabled = bYesNo
    myForm.mnuUpdateInsert.Enabled = bYesNo
    ' set the mnucaption and a flag
    myForm.mnuUpdateInsert.Caption = myForm.cmdUpdateInsertBtn.Caption
    bUpdateTable = bYesNo
End Sub
'=====================================================

' this  function is usful when we add key fields in TRUE GRID.
' It returns the linked key field for SQL
' Example: vdata=(a,b,c,d) nThisrow=2 nTotalrows=4
'          LinkListedData = "('a','b','d')"
Public Function LinkListedData(vData(), nThisRow As Long, _
               nTotalRows As Long) As String
     Dim sLinks As String
     Dim i As Long
     sLinks = ""
     For i = 0 To nTotalRows - 1
        If i <> nThisRow Then
            If sLinks = "" Then
                If Trim(vData(i)) <> "" Then
                    sLinks = MyStr(vData(i))
                End If
            Else
                If Trim(vData(i)) <> "" Then
                    sLinks = sLinks & ", " & MyStr(vData(i))
                End If
            End If
        End If
     Next i
     If sLinks = "" Then
        sLinks = Chr(34) & Chr(34)
     End If
     LinkListedData = "(" & sLinks & ")"
End Function
'=====================================================

'test the first nterms in vdata to see if szTest is already in vdata
Public Function IsAlreadyListed(vData(), ByVal nTerms As Long, _
                         ByVal nThisRow As Long, ByVal szTest As String) As Boolean
                          
     Dim i As Long
     
     IsAlreadyListed = False
     For i = 0 To nTerms - 1
        If i <> nThisRow Then
            If Trim(vData(i)) = Trim(szTest) Then
               IsAlreadyListed = True
               Exit For
            End If
        End If
     Next i
End Function
'=====================================================

Public Function IsFactorDate(sDate As String, Optional VErrorCode) As Boolean
   Dim nError As Integer
   Const IS_DATE As Integer = 0
   Const NON_DATE_FORMAT As Integer = 1
   Const NOT_A_DATE As Integer = 2
   
   If SRegExpMatch(szDatePattern, sDate) = 0 Then  'pass regular expression
      sDate = tfnDateString(tfnFormatDate(sDate))  'formate it (change 010197 to 01/01/97)
      If IsDate(sDate) Then        'just for possible leap year
        IsFactorDate = True
      Else
        IsFactorDate = False
        nError = NOT_A_DATE
      End If
   Else
      IsFactorDate = False
      nError = NON_DATE_FORMAT
   End If
   nError = IS_DATE
      
   If Not IsMissing(VErrorCode) Then
        VErrorCode = nError
   End If
End Function
'=====================================================

Public Function fnCreateTemp_Small(szTableName As String, _
            ByVal szFieldName, ParamArray ArrayValues() As Variant) As Boolean
   
   Dim szSql As String
   Dim k As Integer
   
   fnCreateTemp_Small = False
   
   On Error GoTo errCreateTable
   
   szSql = "SELECT tabname FROM systables WHERE tabname = " & szTableName
   If GetRecordCount(szSql) > 0 Then
      szSql = "DROP TABLE " & szTableName
      t_dbMainDatabase.ExecuteSQL szSql
   End If
   
   szSql = "CREATE TEMP TABLE " & szTableName & "(" & szFieldName & " CHAR(1))"
   t_dbMainDatabase.ExecuteSQL szSql

   For k = 0 To UBound(ArrayValues)
      szSql = "INSERT INTO " & szTableName & "(" & szFieldName & ") VALUES (" _
           & MyStr(ArrayValues(k)) & ")"
      t_dbMainDatabase.ExecuteSQL szSql
   Next
   fnCreateTemp_Small = True
   
    Exit Function
errCreateTable:
    MsgBox "Can not create temporary table for searching.", vbOKOnly + vbCritical, App.Title
    err.Clear
    tfnErrHandler "fnCreateTemp_Small", szSql
   
End Function
'=====================================================

' this function is same as VB function: cdbl()
Public Function fnDeFormatDecimal(ByVal sFormated As String) As Double
 
    Dim sTemp As String
    Dim nPosi As Integer

    On Error Resume Next
    sTemp = ""
    Do
        nPosi = InStr(sFormated, ",")
        If nPosi = 0 Then
            fnDeFormatDecimal = Val(sFormated)
            Exit Function
        End If
        sTemp = Left(sFormated, nPosi - 1)
        sFormated = sTemp & Right(sFormated, Len(sFormated) - nPosi)
    Loop
End Function

'this function will calculate the total working days between two entered days
'but it will still count holidays as working days

Public Function fnWorkingDays(ByVal dStartDate, ByVal dEndDate) As Integer
    
    Dim d As Date
    Dim nDays As Integer
    
    nDays = 0
    For d = dStartDate To dEndDate
        If Weekday(d) <> vbSaturday And Weekday(d) <> vbSunday Then
            nDays = nDays + 1
        End If
    Next
    fnWorkingDays = nDays
    
End Function
'=====================================================

Private Function fnIsHoliday(ByVal dDate) As Boolean
    
    fnIsHoliday = False
    
    'new year day
    If Month(dDate) = 1 Then
        If Day(dDate) = 1 Then
            fnIsHoliday = True
        End If
    End If
    
    'labor day
    If Month(dDate) = 9 Then
        If Day(dDate) = 1 Then
            fnIsHoliday = True
        End If
    End If
    
    If Month(dDate) = 12 Then
        If Day(dDate) = 25 Then
            fnIsHoliday = True
        End If
    End If
    
    If Month(dDate) = 12 Then
        If Day(dDate) = 24 Then
            fnIsHoliday = True
        End If
    End If
    
    If Month(dDate) = 1 Then
        If Day(dDate) = 1 Then
            fnIsHoliday = True
        End If
    End If
    
    If Month(dDate) = 1 Then
        If Day(dDate) = 1 Then
            fnIsHoliday = True
        End If
    End If
    
    If Month(dDate) = 7 Then
        If Day(dDate) = 4 Then
            fnIsHoliday = True
        End If
    End If
    
    If Month(dDate) = 1 Then
        If Day(dDate) = 1 Then
            fnIsHoliday = True
        End If
    End If
    
    If Month(dDate) = 1 Then
        If Day(dDate) = 1 Then
            fnIsHoliday = True
        End If
    End If
    
    If Month(dDate) = 1 Then
        If Day(dDate) = 1 Then
            fnIsHoliday = True
        End If
    End If
    
    If Month(dDate) = 1 Then
        If Day(dDate) = 1 Then
            fnIsHoliday = True
        End If
    End If
    
End Function

'vbuppercase = 1, vblowercase = 2, vbpropercase = 3
Public Function fnCaseConvert(ByVal sInput As String, ByVal vConvertTo As Integer) As String

    If Not IsMissing(vConvertTo) And sInput <> "" Then 'if conversion constants passed
        If vConvertTo = vbUpperCase Or vConvertTo = vbLowerCase Or vConvertTo = vbProperCase Then
            fnCaseConvert = StrConv(sInput, vConvertTo)
        End If
    End If

End Function
'=====================================================

Public Function fnWriteINI(szSection As String, _
                            szKey As String, _
                            szValue As String, _
                            szINIFile As String) As Integer

    Dim bStatus As Boolean 'status returned from api call
    
    'write the [value] for the [section], [key], and ini file sent
    bStatus = WritePrivateProfileString(szSection, szKey, szValue, szINIFile)
    
    fnWriteINI = bStatus

End Function
'=====================================================

Public Function fnReadINI(szSection As String, _
                          szKey As String, _
                          szINIFile As String) As String

    Dim nLength As Integer 'length of the value returned for api call
    Dim szINI As String    'string to hold the value retrieved
    Dim MAX_BUFFER_SIZE As Integer
    
    MAX_BUFFER_SIZE = 50
    
    szINI = Space(MAX_BUFFER_SIZE) 'clear and make the string fixed length
    
    'get the [value] for the [section], [key], and ini file sent
    nLength = GetPrivateProfileString(szSection, szKey, "", szINI, MAX_BUFFER_SIZE, szINIFile)
    
    If nLength <> 0 Then 'if length positive [value] has been found
        szINI = Left(szINI, nLength) 'make it a basic string
    Else
        szINI = ""
    End If
    
    fnReadINI = szINI 'return the value

End Function
'=====================================================

'
'Function : fnStripNULL - strips off the NULL terminator on C strings
'Variables: NULL terminated string
'Return   : original string with the null removed
'
Public Function fnStripNULL(ByRef szString As String) As String
  
    Dim nPos As Integer 'position of the NULL terminator
   
    If Len(szString) = Null Or Len(szString) = 0 Then 'make sure string is valid
        szString = "" 'gszEMPTY 'if not set the string to an empty string
    Else
    
        nPos = InStr(szString, Chr(0)) 'get the position of the NULL terminator
        
        If nPos > 0 Then 'if nPos is greater than 0 then a NULL was found
           szString = Left(szString, nPos - 1) 'strip off the NULL terminator
        End If 'if string did not have a NULL do not change it
    
    End If
    
    fnStripNULL = szString 'return the string

End Function
'=====================================================

'Function : fnParseString - parses a delimited string, default is a space
'Variables: string to parse, optional delimiter, NOTE: original string is destroyed in the process, converion constant
'Return   : first deliminted substring in main string
'
Public Function fnParseString(ByRef szMainString As String, Optional vDelimiter As Variant, Optional vConvertTo As Variant) As String

    Dim nPos As Integer        'position of the delimiter
    Dim szDelimiter As String  'delimiter to search for in the main string
    Dim szBuffer As String     'string buffer
    
    If IsMissing(vDelimiter) Then 'set the delimiter to as space if none was passed
        szDelimiter = " " 'gszSPACE
    Else
        szDelimiter = vDelimiter
    End If
    
    szBuffer = Left(szMainString, 1)
    
    Do While szBuffer = szDelimiter
        szMainString = Mid(szMainString, 1)
        szBuffer = Left(szMainString, 1)
    Loop
    
    nPos = InStr(szMainString, szDelimiter) 'search for a delimiter

    If nPos > 0 Then 'if delimiter found
        szBuffer = Left(szMainString, nPos - 1) 'parse of the substring
        szMainString = Mid(szMainString, nPos + 1)   'remove the substring and delimiter from the main string
    Else
        szBuffer = szMainString 'return the last substring
        szMainString = "" 'gszEMPTY 'empty the string
    End If
    
    If Not IsMissing(vConvertTo) Then 'if conversion constants passed
        If vConvertTo = vbUpperCase Or vConvertTo = vbLowerCase Or vConvertTo = vbProperCase Then
            szBuffer = StrConv(szBuffer, vConvertTo) 'convert to the case constant sent, if its valid
        End If
    End If
    
    fnParseString = szBuffer

End Function
'=====================================================
'
'Function : fnGetAppDir - returns the application directory path
'Variables: optional variable to add a slash to the end of the path
'Return   : directory path
'
Public Function fnGetAppDir(Optional vAddSlash As Variant) As String
    
    Dim szTemp As String 'temp to hold the path
    Dim gszSLASH As String
    gszSLASH = "/"
        
    szTemp = App.Path 'use the App object to retrieve the path
        
    If Not IsMissing(vAddSlash) Then
        If Right(szTemp, 1) <> gszSLASH And vAddSlash = True Then 'add a slash if it needs one
            szTemp = szTemp + gszSLASH
        End If
    End If
    
    fnGetAppDir = szTemp 'return the path

End Function
'=====================================================

'same as fnWriteINI()
Public Function tfnWriteINI(szSection As String, szKey As String, szValue As String, szINIFile As String) As Integer

    Dim bStatus As Boolean 'status returned from api call
    
    'write the [value] for the [section], [key], and ini file sent
    bStatus = WritePrivateProfileString(szSection, szKey, szValue, szINIFile)
    
    tfnWriteINI = bStatus

End Function
'=====================================================
'
'Function : max - returns the maximum of the 2 values passed
'Variables: two variables types of any kind
'Return   : the max of the 2
'
Public Function max(a As Variant, b As Variant) As Variant
    max = -a * (a >= b) - b * (a < b)
End Function
'=====================================================
'
'Function : min - returns the minimum of the 2 values passed
'Variables: two variable types of any kind
'Return   : the min of the 2
'
Public Function min(a As Variant, b As Variant) As Variant
    min = -a * (a <= b) - b * (a > b)
End Function
'=====================================================
'
'Function : LOWORD - lower 2 bytes of a long
'Variables: long variable
'Return   : integer value of lower 2 bytes
'
Public Function LOWORD(lVal As Long) As Integer
    LOWORD = lVal And MAX_INT
End Function
'=====================================================
'
'Function : HIWORD - gets the upper 2 bytes of a long
'Variables: long variable
'Return   : integer value of upper 2 bytes
'
Public Function HIWORD(lVal As Long) As Integer
    HIWORD = lVal& \ MAX_INT
End Function
'=====================================================
'
'Function : fnFixAmpersand - adds ampersand to a string with an ampersand - override default button behavior
'Variables: string to check for ampersand
'Return   : text with any single ampersands replaced with double ampersands
'
Public Function fnFixAmpersand(ByVal szTextIn As String) As String
    
    Dim szTemp As String 'temp string to hold converted string
    Dim nPos As Integer  'holds the position of the ampersand
    Dim gszAMPERSAND As String
    gszAMPERSAND = "&"
    Dim gszEMPTY As String
    gszEMPTY = ""

    nPos = InStr(szTextIn, gszAMPERSAND) 'search for an ampersand
    
    If nPos <> 0 Then 'if no ampersand found just return the original string
        
        szTemp = gszEMPTY 'clear the temp string
        
        Do While nPos <> 0 'search for all the ampersnads in the string
            szTemp = szTemp + Left(szTextIn, nPos) + gszAMPERSAND 'add another ampersand next to the other
            szTextIn = Mid(szTextIn, nPos + 1) 'strip off substring saved in szTemp
            nPos = InStr(szTextIn, gszAMPERSAND) 'search for next ampersand
        Loop
        
        szTemp = szTemp + szTextIn 'save the last part of the original string
        
        fnFixAmpersand = szTemp 'return the modified string
        Exit Function
        
    End If
    
    fnFixAmpersand = szTextIn 'no ampersand found return the original string

End Function
'=====================================================
'
'Function : fnCenterForm - centers a form in the screen
'Variables: pointer to the form, optional pointer to parent form
'Return   : none
'
Sub fnCenterForm(frmCurrent As Form, Optional vParentForm As Variant)
  
    If IsMissing(vParentForm) Then
        frmCurrent.Left = (Screen.Width - frmCurrent.Width) / 2
        frmCurrent.Top = (Screen.Height - frmCurrent.Height) / 2
    Else
        
        If vParentForm.Width > frmCurrent.Width Then
            frmCurrent.Left = vParentForm.Left + (vParentForm.Width - frmCurrent.Width) / 2
        Else
            frmCurrent.Left = (Screen.Width - frmCurrent.Width) / 2
        End If

        If vParentForm.Height > frmCurrent.Height Then
            frmCurrent.Top = vParentForm.Top + (vParentForm.Height - frmCurrent.Height) / 2
        Else
            frmCurrent.Top = (Screen.Height - frmCurrent.Height) / 2
        End If
    End If
    
End Sub
'=====================================================
'
'Function : fnDisableFormSystemClose
'Variables: pointer to the form
'Return   : none
'
Public Sub fnDisableFormSystemClose(ByRef frmForm As Form)
    
    Dim nCode As Integer
    
    nCode = GetSystemMenu(myForm.hwnd, False)
    Call ModifyMenu(nCode, SC_CLOSE, 1, 0, "&Close")
    Call ModifyMenu(nCode, SC_SIZE, 1, 0, "&Size")

End Sub
'=====================================================

Function fnUINT2INT(lValue As Long) As Integer

    If lValue > 32767 Then
        fnUINT2INT = CInt(lValue - 65536)
    Else
        fnUINT2INT = CInt(lValue)
    End If

End Function
'=================================

' A generally used utility function, execute SQL and take care of errors
Public Function fnGetExeSQLCount(strSQL As String, _
                             Optional vCaller As Variant, _
                             Optional vMsg As Variant, _
                             Optional vDB As Variant) As Integer

    Dim objDB As DataBase
    
    If IsMissing(vDB) Then
        Set objDB = t_dbMainDatabase
    Else
        Set objDB = vDB
    End If
    
    On Error GoTo errExecute
    If objDB Is t_dbMainDatabase Then
        fnGetExeSQLCount = objDB.ExecuteSQL(strSQL)
    Else
        objDB.Execute strSQL
        fnGetExeSQLCount = 0
    End If

    On Error GoTo 0
    Exit Function

errExecute:
    Dim bShow As Boolean
    
    #If DEVELOP Then
        'subShowODBCError vMsg, strSQL
    #Else
        bShow = Not IsMissing(vMsg)
        If IsMissing(vCaller) Then
            tfnErrHandler "fngetExeSQLcount", strSQL, , bShow
        Else
            tfnErrHandler "fngetExesqlcount\vCaller", strSQL, , bShow
        End If
    #End If
    fnGetExeSQLCount = -1
    
End Function
'==========================================================================

' A generally used utility function, execute SQL and take care of errors
Public Function fnExecuteSQL(strSQL As String, _
                             Optional vCaller As Variant, _
                             Optional vMsg As Variant, _
                             Optional vDB As Variant) As Integer

    'david 01/27/2005  #465134
    'if the ExecuteSQL returns more than 32767 rows,
    'the function return type Integer
    'will cause RUN-TIME ERROR 6 Overflow
    'but I don't want to change the return type to Long
    'because it may crash other programs that assigning
    'the return value of this function to a Integer type variable.
    Dim lRet As Long
    
    Dim objDB As DataBase
    
    If IsMissing(vDB) Then
        Set objDB = t_dbMainDatabase
    Else
        Set objDB = vDB
    End If
    
    On Error GoTo errExecute
    If objDB Is t_dbMainDatabase Then
        'david 01/27/2005  #465134
        lRet = objDB.ExecuteSQL(strSQL)
        
        If lRet > 32000 Then
            lRet = 32000
        End If
        
        fnExecuteSQL = CInt(lRet)
    Else
        objDB.Execute strSQL
        fnExecuteSQL = 0
    End If

    On Error GoTo 0
    Exit Function

errExecute:
    Dim bShow As Boolean
    
    #If DEVELOP Then
        subShowODBCError vMsg, strSQL
    #Else
        bShow = Not IsMissing(vMsg)
        If IsMissing(vCaller) Then
            tfnErrHandler "fnExecuteSQL", strSQL, , bShow
        Else
            tfnErrHandler "fnExecuteSQL\vCaller", strSQL, , bShow
        End If
    #End If
    fnExecuteSQL = -1
    
End Function
'============================================================

'Weigong's
Public Function fnExeSQL(szSql As String, Optional nDB As Variant, _
                Optional szCalledFrom As Variant, Optional bShowError As Variant) As Boolean
                
    Dim szMsg As String
    
    On Error GoTo SQLError
    
    If IsMissing(nDB) Then
        nDB = DB_REMOTE
    End If
    
    Select Case nDB
        Case DB_LOCAL
            dbLocal.Execute szSql
        Case DB_REMOTE
            t_dbMainDatabase.ExecuteSQL szSql
    End Select
    
    fnExeSQL = True
    Exit Function
      
SQLError:
      fnExeSQL = False
      If IsMissing(szCalledFrom) Then
         szCalledFrom = ""
      End If
      If IsMissing(bShowError) Then
         bShowError = True
      End If
      tfnErrHandler "fnExeSQL, " & szCalledFrom, szSql, bShowError
      On Error GoTo 0
End Function
'=====================

Public Sub subTextSelected()

On Error Resume Next
    Dim txtBox As Control
    Set txtBox = myForm.ActiveControl
    
    If TypeOf txtBox Is Textbox Then
        txtBox.SelStart = 0
        txtBox.SelLength = Len(Trim(txtBox.Text))
    End If
    
End Sub
'==========================================

Private Sub subShowODBCError(Optional vMsg As Variant, Optional vSQL As Variant)

    Dim i As Integer
    Dim sMsgs As String
    Dim sNumbers As String
    Dim sODBCErrors As String
    
    #If DEVELOP Then
        Dim strSQL As String
        If IsMissing(vSQL) Then
            strSQL = ""
        Else
            strSQL = vSQL
        End If
    #End If
    
    If err.Number = 3146 Or t_engFactor.Errors.Count > 2 Then
        With t_engFactor.Errors
            If .Count > 0 Then
                For i = 0 To .Count - 2
                    sMsgs = sMsgs & "Number: " & .Item(i).Number & Space(5) _
                        & .Item(i).Description & vbCrLf
                Next
            End If
            If .Count <= 2 Then
                sNumbers = ""
            Else
                sNumbers = "s"
            End If
        End With
        sODBCErrors = "The following error" & sNumbers _
            & " occurred while doing an ODBC query:" _
            & vbCrLf & vbCrLf & vbCrLf & sMsgs
    Else
        sODBCErrors = err.Description
    End If
    
    Dim sMsg As String
    If IsMissing(vMsg) Then
        #If DEVELOP Then
            sMsg = "An error occurred while doing a SQL query" & vbCrLf & vbCrLf _
                & "Error# " & CStr(err.Number) & vbCrLf & err.Description
            sMsg = sMsg & vbCrLf & vbCrLf & strSQL & vbCrLf & vbCrLf & sODBCErrors
            Clipboard.SetText strSQL
        #Else
            sMsg = ""
        #End If
    Else
        sMsg = vMsg
        #If DEVELOP Then
            If Trim(sMsg) = "" Then
                sMsg = "SQL: " & strSQL & vbCrLf & vbCrLf & sODBCErrors
            Else
                sMsg = sMsg & vbCrLf & vbCrLf & "SQL: " & strSQL & vbCrLf & vbCrLf _
                    & sODBCErrors
            End If
            Clipboard.SetText strSQL
        #Else
            If Trim(sMsg) = "" Then
                sMsg = sODBCErrors
            Else
            
                sMsg = sMsg & vbCrLf & vbCrLf & sODBCErrors
            End If
        #End If
    End If
    If sMsg <> "" Then
        MsgBox sMsg, vbOKOnly + vbCritical, App.Title
    End If
    err.Clear

End Sub
'==========================================================================

Public Sub subEnableUpdate(bStatus As Boolean)

    myForm.cmdUpdateInsertBtn.Enabled = bStatus
    myForm.mnuUpdateInsert.Enabled = bStatus
    
End Sub
'=====================================================================

Public Sub subEnableRefresh(bStatus As Boolean)

    On Error GoTo tryAnother
    If bStatus Then
        subEnableDelete False
    End If
    
    myForm.cmdRefreshSelectBtn.Enabled = bStatus
    myForm.mnuRefreshSelect.Enabled = bStatus

    Exit Sub
tryAnother:
    On Error Resume Next
    myForm.cmdRefresh.Enabled = bStatus
    myForm.mnuRefresh.Enabled = bStatus
End Sub
'=====================================================================

Public Sub subEnableSearchButton(ByRef ctrlButton As FactorFrame, _
                                 ByVal bStatus As Boolean)

    ctrlButton.Style = 3  'command button
    ctrlButton.ShowFocusRect = True 'show a rectangular if focused
    ctrlButton.Enabled = bStatus
    If bStatus Then
        ctrlButton.Picture = frmContext.LoadPicture(SEARCH_UP)
    Else
        ctrlButton.Picture = frmContext.LoadPicture(SEARCH_DOWN)
    End If
    
End Sub
'=====================================================================

'take care of transfer '.' and ""
Public Function fncDbl(ByVal sInput As String) As Double
    
    If Trim(sInput) = "" Then
        fncDbl = 0
        Exit Function
    End If
    
    If Trim(sInput) = "." Then
        fncDbl = 0
        Exit Function
    End If
     
    fncDbl = CDbl(sInput)
    
End Function
'================================================

'If a box allow at most 'Digit_Limit' digits, how to convert input to a decimal?
'the following function to avoid beeping after converting
'a big integer of (Digit_Limit -1) or more digits to a double with 2 decimals.
Private Function ShowDecimalOld(ByVal dInput As Double, Digit_Limit As Integer) As String
    
    Dim lMax As Double
    Dim k As Integer
    
    lMax = 1
    For k = 1 To Digit_Limit - 2
         lMax = lMax * 10
    Next k
    If dInput < lMax Then
        ShowDecimalOld = tfnFormatDecimal(dInput, 2)
    ElseIf dInput >= lMax And dInput < lMax * 10 Then
        ShowDecimalOld = tfnFormatDecimal(dInput, 1)
    Else
        ShowDecimalOld = tfnFormatDecimal(dInput, 0)
    End If
    
End Function
'=====================================================================

'If a box allow at most 'Digit_Limit' digits, how to convert input to a decimal?
'the following function to avoid beeping after converting
'a big integer of (Digit_Limit -1) or more digits to a double with 2 decimals.
'dInput must be "", ".", double, integer, long
Public Function ShowDecimal(ByVal dInput As Variant, Digit_Limit As Integer, _
    Optional Decimal_Limit As Variant) As String
    
    Dim lAbs As Double
    Dim lMax As Double
    Dim k As Integer
    Dim n As Double
    Dim mySign As Integer
    Dim bDone As Boolean
    Dim Compensation As Double
    Dim lVal As Double
    
    bDone = False
    
    If IsMissing(Decimal_Limit) Then
        If Digit_Limit > 2 Then
            Decimal_Limit = 2
        Else
            Decimal_Limit = Digit_Limit - 1
        End If
    End If
    
    'the least is 0 for decimal_limit
    If Decimal_Limit < 0 Then
        Decimal_Limit = 0
    End If
    
    'the least is 1 for digit_limit
    If Digit_Limit < 1 Then
        Digit_Limit = 1
    End If
    
    'at most 15 digits
    If Digit_Limit > 15 Then
        Digit_Limit = 15
    End If
    
    'decimal_limit at least one less than digit_limit
    If Digit_Limit - Decimal_Limit < 1 Then
        Decimal_Limit = Digit_Limit - 1
    End If
            
    'for the case NOthing
    dInput = "" & dInput
            
    'dInput = "0" & dInput 'prefix '0' cant take care of negative
    If dInput = "" Or dInput = "." Then
        dInput = 0
    End If
        
    If Not IsNumeric(dInput) Then
        Exit Function
        'dInput = 0
    End If
        
    If dInput < 0 Then
        mySign = -1
        lAbs = -1 * dInput
    Else
        mySign = 1
        lAbs = dInput
    End If
        
    '1Max -1 is the largest possible integer
    'with decimal_limit number of decimals
    lMax = 1
    For k = 1 To Digit_Limit - Decimal_Limit
        lMax = lMax * 10
    Next k
            
    If lAbs < lMax Then
        ShowDecimal = tfnFormatDecimal(lAbs * mySign, Decimal_Limit)
        bDone = True
    Else
        k = 1
        n = 1
        While k <= Decimal_Limit
            If lAbs >= lMax * n And lAbs < lMax * n * 10 Then
                ShowDecimal = tfnFormatDecimal(lAbs * mySign, Decimal_Limit - k)
                bDone = True
            End If
            k = k + 1
            n = n * 10
        Wend
    End If
    
    'only give digit_limit effective digits, i.e. 1234 = 1200 if digit_limit = 2
    If Not bDone Then
        n = 0
        k = 1
        lMax = 1
        While n = 0
            If lMax <= lAbs And lAbs < 10 * lMax Then
                n = k
            End If
            k = k + 1
            lMax = lMax * 10
        Wend ' n=5 and lmx = 100,000 if labs = 12,345
        
        Compensation = 1
        
        'figue out the integer needed for add later
        If n - Digit_Limit <= 0 Then
            Exit Function 'impossible since ndone is false
        Else
            For k = 1 To n - Digit_Limit
                Compensation = Compensation * 10
            Next
        End If 'if labs = 12345, digit_limit = 2, then compensation = 1000
            
        'figue out the integer with only digit_limit effective digits
        'if labs = 12500, digit_limit = 2 then lval = 12000
        
        lVal = 0
        
        'if use "\", will overflow, so have to use 'val(left(val ...'
        For k = 1 To Digit_Limit
            lVal = lVal + Val(Left(Val((lAbs - lVal) / (lMax / 10)), 1)) * (lMax / 10)
            lMax = lMax / 10
        Next k
                
        'now add compensation if necessary, dont forget put back the sign
        '1254 = 1300, 1249 = 1200 if digit_limit = 2
        If lAbs - lVal >= 5 * (Compensation / 10) Then
            ShowDecimal = mySign * lVal + mySign * Compensation
        Else
            ShowDecimal = mySign * lVal
        End If
    End If
        
End Function
'=====================================================================

Public Sub subSetFont(sFontName As String, _
                       ByVal FONTSIZE As Integer, _
                       ByVal bBold As Boolean)
    
    Dim font As LOGFONT
    
    ''font.lfEscapement = nRotate * 10   ' 180-degree rotation
    font.lfFaceName = sFontName & Chr$(0) 'Null character at end
    font.lfHeight = -FONTSIZE * GetDeviceCaps(Printer.hDC, 90) / 72 'LOGPIXELSY) / 72 ' one inch contains 72 points.
    'LOGPIXELSY
    If bBold Then
        font.lfWeight = 700
    Else
        font.lfWeight = 400
    End If
    
    If hCurrFont <> 0 Then
        DeleteObject hCurrFont
    End If
    hCurrFont = CreateFontIndirect(font)
    
    If hPrevFont = 0 Then
        hPrevFont = SelectObject(Printer.hDC, hCurrFont)
    Else
        SelectObject Printer.hDC, hCurrFont
    End If

End Sub
'===========================================

Public Sub SubGetDefaultPrinter()

    Dim myPrinter As Printer
    
    For Each myPrinter In Printers
        If myPrinter.Orientation = vbPRORPortrait Then
            ' Set printer as system default.
            Set Printer = myPrinter
            ' Stop looking for a printer.
            Exit For
        End If
    Next
End Sub
'===========================================

'open local database, need to reset sInifilename
Public Function fnOpenLocalDatabase() As DataBase
    
    Dim sDataBasePath As String
    Dim sIniFileName As String
    
    sIniFileName = App.Path & "\duty.ini"
    sDataBasePath = fnReadINI("DATABASE", "path", sIniFileName)
        
    On Error GoTo ERROR_CONNECTING 'set the runtime error handler for database connection

    If t_engFactor Is Nothing Then
        Set t_engFactor = New DBEngine 'create a new dDBEngine
        t_engFactor.IniPath = tfnGetSystemDir 'put the path in engine ini variable
    End If
    
    If t_wsWorkSpace Is Nothing Then
        Set t_wsWorkSpace = t_engFactor.Workspaces(0) 'set the default workspace handle
    End If

    Set fnOpenLocalDatabase = t_wsWorkSpace.OpenDatabase(sDataBasePath) '(t_engFactor.IniPath & "\factor.mdb")
    On Error GoTo 0
    Exit Function

ERROR_CONNECTING:
    MsgBox err.Description, vbOKOnly + vbCritical, "Local Access " & szCONNECTION_ERROR

    Set fnOpenLocalDatabase = Nothing

    On Error GoTo 0
End Function
'==========================

Public Function fnGetSysParm(ByVal nNum As Integer) As String 'sysParm

    Dim sSql As String
    Dim rsTemp As Recordset
    
    sSql = "SELECT parm_field, parm_desc FROM sys_parm WHERE" _
        & " parm_nbr = " & nNum
        
    If GetRecordSet(rsTemp, sSql) > 0 Then
        fnGetSysParm = fnGetField(rsTemp!parm_field)
        'fnGetSysParm.parm_desc = fnGetField(rsTemp!parm_desc)
    Else
        fnGetSysParm = ""
        'fnGetSysParm.parm_desc = ""
    End If
    
End Function
'===============================
Public Function fnGetDecimalDef(ByVal nLength As Integer) As String
    Select Case nLength
        Case 511
            fnGetDecimalDef = "decimal(1)"
        Case 512
            fnGetDecimalDef = "decimal(2,0)"
        Case 513
            fnGetDecimalDef = "decimal(2,1)"
        Case 514
            fnGetDecimalDef = "decimal(2,2)"
        
        Case 767
            fnGetDecimalDef = "decimal(2)" '253
        Case 768
            fnGetDecimalDef = "decimal(3,0)"
        Case 769
            fnGetDecimalDef = "decimal(3,1)"
        Case 770
            fnGetDecimalDef = "decimal(3,2)"
        Case 771
            fnGetDecimalDef = "decimal(3,3)"
        
        Case 1023
            fnGetDecimalDef = "decimal(3)" '252
        Case 1024
            fnGetDecimalDef = "decimal(4,0)"
        Case 1025
            fnGetDecimalDef = "decimal(4,1)"
        Case 1026
            fnGetDecimalDef = "decimal(4,2)"
        Case 1027
            fnGetDecimalDef = "decimal(4,3)"
        Case 1028
            fnGetDecimalDef = "decimal(4,4)"
        
        Case 1279
            fnGetDecimalDef = "decimal(4)" '251
        Case 1280
            fnGetDecimalDef = "decimal(5,0)"
        Case 1281
            fnGetDecimalDef = "decimal(5,1)"
        Case 1282
            fnGetDecimalDef = "decimal(5,2)"
        Case 1283
            fnGetDecimalDef = "decimal(5,3)"
        Case 1284
            fnGetDecimalDef = "decimal(5,4)"
        Case 1285
            fnGetDecimalDef = "decimal(5,5)"
        
        Case 1535
            fnGetDecimalDef = "decimal(5)" '250
        Case 1536
            fnGetDecimalDef = "decimal(6,0)"
        Case 1537
            fnGetDecimalDef = "decimal(6,1)"
        Case 1538
            fnGetDecimalDef = "decimal(6,2)"
        Case 1539
            fnGetDecimalDef = "decimal(6,3)"
        Case 1540
            fnGetDecimalDef = "decimal(6,4)"
        Case 1541
            fnGetDecimalDef = "decimal(6,5)"
        Case 1542
            fnGetDecimalDef = "decimal(6,6)"
        
        Case 1791
            fnGetDecimalDef = "decimal(6)" '249
        Case 1792
            fnGetDecimalDef = "decimal(7,0)"
        Case 1793
            fnGetDecimalDef = "decimal(7,1)"
        Case 1794
            fnGetDecimalDef = "decimal(7,2)"
        Case 1795
            fnGetDecimalDef = "decimal(7,3)"
        Case 1796
            fnGetDecimalDef = "decimal(7,4)"
        Case 1797
            fnGetDecimalDef = "decimal(7,5)"
        Case 1798
            fnGetDecimalDef = "decimal(7,6)"
        Case 1799
            fnGetDecimalDef = "decimal(7,7)"
        
        Case 2047
            fnGetDecimalDef = "decimal(7)" '248
        Case 2048
            fnGetDecimalDef = "decimal(8,0)"
        Case 2049
            fnGetDecimalDef = "decimal(8,1)"
        Case 2050
            fnGetDecimalDef = "decimal(8,2)"
        Case 2051
            fnGetDecimalDef = "decimal(8,3)"
        Case 2052
            fnGetDecimalDef = "decimal(8,4)"
        Case 2053
            fnGetDecimalDef = "decimal(8,5)"
        Case 2054
            fnGetDecimalDef = "decimal(8,6)"
        Case 2055
            fnGetDecimalDef = "decimal(8,7)"
        Case 2056
            fnGetDecimalDef = "decimal(8,8)"
        
        Case 2303
            fnGetDecimalDef = "decimal(8)" '247
        Case 2304
            fnGetDecimalDef = "decimal(9,0)"
        Case 2305
            fnGetDecimalDef = "decimal(9,1)"
        Case 2306
            fnGetDecimalDef = "decimal(9,2)"
        Case 2307
            fnGetDecimalDef = "decimal(9,3)"
        Case 2308
            fnGetDecimalDef = "decimal(9,4)"
        Case 2309
            fnGetDecimalDef = "decimal(9,5)"
        Case 2310
            fnGetDecimalDef = "decimal(9,6)"
        Case 2311
            fnGetDecimalDef = "decimal(9,7)"
        Case 2312
            fnGetDecimalDef = "decimal(9,8)"
        Case 2313
            fnGetDecimalDef = "decimal(9,9)"
        
        Case 2559
            fnGetDecimalDef = "decimal(9)" '246
        Case 2560
            fnGetDecimalDef = "decimal(10,0)"
        Case 2561
            fnGetDecimalDef = "decimal(10,1)"
        Case 2562
            fnGetDecimalDef = "decimal(10,2)"
        Case 2563
            fnGetDecimalDef = "decimal(10,3)"
        Case 2564
            fnGetDecimalDef = "decimal(10,4)"
        Case 2565
            fnGetDecimalDef = "decimal(10,5)"
        Case 2566
            fnGetDecimalDef = "decimal(10,6)"
        Case 2567
            fnGetDecimalDef = "decimal(10,7)"
        Case 2568
            fnGetDecimalDef = "decimal(10,8)"
        Case 2569
            fnGetDecimalDef = "decimal(10,9)"
        Case 2570
            fnGetDecimalDef = "decimal(10,10)"
        
        Case 2815
            fnGetDecimalDef = "decimal(10)" '245
        Case 2816
            fnGetDecimalDef = "decimal(11,0)"
        Case 2817
            fnGetDecimalDef = "decimal(11,1)"
        Case 2818
            fnGetDecimalDef = "decimal(11,2)"
        Case 2819
            fnGetDecimalDef = "decimal(11,3)"
        Case 2820
            fnGetDecimalDef = "decimal(11,4)"
        Case 2821
            fnGetDecimalDef = "decimal(11,5)"
        Case 2822
            fnGetDecimalDef = "decimal(11,6)"
        Case 2823
            fnGetDecimalDef = "decimal(11,7)"
        Case 2824
            fnGetDecimalDef = "decimal(11,8)"
        Case 2825
            fnGetDecimalDef = "decimal(11,9)"
        Case 2826
            fnGetDecimalDef = "decimal(11,10)"
        Case 2827
            fnGetDecimalDef = "decimal(11,11)"
        
        Case 3071
            fnGetDecimalDef = "decimal(11)" '244
        Case 3072
            fnGetDecimalDef = "decimal(12,0)"
        Case 3073
            fnGetDecimalDef = "decimal(12,1)"
        Case 3074
            fnGetDecimalDef = "decimal(12,2)"
        Case 3075
            fnGetDecimalDef = "decimal(12,3)"
        Case 3076
            fnGetDecimalDef = "decimal(12,4)"
        Case 3077
            fnGetDecimalDef = "decimal(12,5)"
        Case 3078
            fnGetDecimalDef = "decimal(12,6)"
        Case 3079
            fnGetDecimalDef = "decimal(12,7)"
        Case 3080
            fnGetDecimalDef = "decimal(12,8)"
        Case 3081
            fnGetDecimalDef = "decimal(12,9)"
        Case 3082
            fnGetDecimalDef = "decimal(12,10)"
        Case 3083
            fnGetDecimalDef = "decimal(12,11)"
        Case 3084
            fnGetDecimalDef = "decimal(12,12)"
        
        Case 3327
            fnGetDecimalDef = "decimal(12)" '243
        Case 3328
            fnGetDecimalDef = "decimal(13,0)"
        Case 3329
            fnGetDecimalDef = "decimal(13,1)"
        Case 3330
            fnGetDecimalDef = "decimal(13,2)"
        Case 3331
            fnGetDecimalDef = "decimal(13,3)"
        Case 3332
            fnGetDecimalDef = "decimal(13,4)"
        Case 3333
            fnGetDecimalDef = "decimal(13,5)"
        Case 3334
            fnGetDecimalDef = "decimal(13,6)"
        Case 3335
            fnGetDecimalDef = "decimal(13,7)"
        Case 3336
            fnGetDecimalDef = "decimal(13,8)"
        Case 3337
            fnGetDecimalDef = "decimal(13,9)"
        Case 3338
            fnGetDecimalDef = "decimal(13,10)"
        Case 3339
            fnGetDecimalDef = "decimal(13,11)"
        Case 3340
            fnGetDecimalDef = "decimal(13,12)"
        Case 3341
            fnGetDecimalDef = "decimal(13,13)"
        
        Case 3583
            fnGetDecimalDef = "decimal(13)" '242
        Case 3584
            fnGetDecimalDef = "decimal(14,0)"
        Case 3585
            fnGetDecimalDef = "decimal(14,1)"
        Case 3586
            fnGetDecimalDef = "decimal(14,2)"
        Case 3587
            fnGetDecimalDef = "decimal(14,3)"
        Case 3588
            fnGetDecimalDef = "decimal(14,4)"
        Case 3589
            fnGetDecimalDef = "decimal(14,5)"
        Case 3590
            fnGetDecimalDef = "decimal(14,6)"
        Case 3591
            fnGetDecimalDef = "decimal(14,7)"
        Case 3592
            fnGetDecimalDef = "decimal(14,8)"
        Case 3593
            fnGetDecimalDef = "decimal(14,9)"
        Case 3594
            fnGetDecimalDef = "decimal(14,10)"
        Case 3595
            fnGetDecimalDef = "decimal(14,11)"
        Case 3596
            fnGetDecimalDef = "decimal(14,12)"
        Case 3597
            fnGetDecimalDef = "decimal(14,13)"
        Case 3598
            fnGetDecimalDef = "decimal(14,14)"
        
        Case 3839
            fnGetDecimalDef = "decimal(14)" '241
        Case 3840
            fnGetDecimalDef = "decimal(15,0)"
        Case 3841
            fnGetDecimalDef = "decimal(15,1)"
        Case 3842
            fnGetDecimalDef = "decimal(15,2)"
        Case 3843
            fnGetDecimalDef = "decimal(15,3)"
        Case 3844
            fnGetDecimalDef = "decimal(15,4)"
        Case 3845
            fnGetDecimalDef = "decimal(15,5)"
        Case 3846
            fnGetDecimalDef = "decimal(15,6)"
        Case 3847
            fnGetDecimalDef = "decimal(15,7)"
        Case 3848
            fnGetDecimalDef = "decimal(15,8)"
        Case 3849
            fnGetDecimalDef = "decimal(15,9)"
        Case 3850
            fnGetDecimalDef = "decimal(15,10)"
        Case 3851
            fnGetDecimalDef = "decimal(15,11)"
        Case 3852
            fnGetDecimalDef = "decimal(15,12)"
        Case 3853
            fnGetDecimalDef = "decimal(15,13)"
        Case 3854
            fnGetDecimalDef = "decimal(15,14)"
        Case 3855
            fnGetDecimalDef = "decimal(15,15)"
        
        Case 4095
            fnGetDecimalDef = "decimal(15)"     '240
        Case 4096
            fnGetDecimalDef = "decimal(16,0)"
        Case 4097
            fnGetDecimalDef = "decimal(16,1)"
        Case 4098
            fnGetDecimalDef = "decimal(16,2)"
        Case 4099
            fnGetDecimalDef = "decimal(16,3)"
        Case 4100
            fnGetDecimalDef = "decimal(16,4)"
        Case 4101
            fnGetDecimalDef = "decimal(16,5)"
        Case 4102
            fnGetDecimalDef = "decimal(16,6)"
        Case 4103
            fnGetDecimalDef = "decimal(16,7)"
        Case 4104
            fnGetDecimalDef = "decimal(16,8)"
        Case 4105
            fnGetDecimalDef = "decimal(16,9)"
        Case 4106
            fnGetDecimalDef = "decimal(16,10)"
        Case 4107
            fnGetDecimalDef = "decimal(16,11)"
        Case 4108
            fnGetDecimalDef = "decimal(16,12)"
        Case 4109
            fnGetDecimalDef = "decimal(16,13)"
        Case 4110
            fnGetDecimalDef = "decimal(16,14)"
        Case 4111
            fnGetDecimalDef = "decimal(16,15)"
        Case 4112
            fnGetDecimalDef = "decimal(16,16)"
        
        Case 4351
            fnGetDecimalDef = "decimal(16)"  '239
        Case Else
            fnGetDecimalDef = "Deciaml"
    End Select
End Function
'===============================

Public Function fnClearLastComma(sInput As String) As String
    If fnGetField(sInput) = "" Then
        fnClearLastComma = ""
        Exit Function
    End If
    Dim sTemp As String
    sTemp = Trim(sInput)
    
    'take care of sInput = "  ,, "
    While Mid(sTemp, Len(sTemp)) = "," And Trim(sTemp) <> ""
        sTemp = " " & fnGetField(Left(sTemp, Len(sTemp) - 1))
    Wend
    fnClearLastComma = fnGetField(sTemp)
    
End Function
'===========================================

Public Function fnOpen2ndDatabase() As Boolean
    
    On Error GoTo ERROR_CONNECTING 'set the runtime error handler for database connection
                                                
    Set t_dbMainDatabase2nd = _
        t_wsWorkSpace.OpenDatabase _
        ("", False, False, "ODBC;DSN=Interim32;DB=/factor/interim/factor;HOST=ether5;SERV=sqlexec;SRVR=ether5;PRO=sesoctcp;" _
            & "UID=" & tfnGetUserName & ";PWD=" & tfnGetUserPassward)
        
        '("", False, False, "ODBC;DSN=interim32;DB=/factor/interim/factor;HOST=ether5;SERV=sqlexec;PRO=tcp-ip;" _
        '    & "UID=" & tfnGetUserName & ";PWD=" & tfnGetUserPassward)
    'check database.connect = ODBC;DSN=Interim32;DB=/factor/interim/factor;HOST=ether5;SERV=sqlexec;SRVR=ether5;PRO=sesoctcp;UID=ssfactor;PWD=menus

    fnOpen2ndDatabase = True
    Exit Function

ERROR_CONNECTING:
    MsgBox err.Description, vbOKOnly + vbCritical, szCONNECTION_ERROR
    fnOpen2ndDatabase = False

End Function
'===============================================

Public Function tfnGetUserPassward() As String
' return the current user passward as was logged into factmenu
    
    Dim sTemp As String, sUser As String, nPosi As Integer
    Const sKeyWord As String = "PWD="
    
   ' #If DEVELOP Or (Not FACTOR_MENU) Then
   
    If t_dbMainDatabase Is Nothing Then Exit Function
    sTemp = t_dbMainDatabase.Connect
    
    nPosi = InStr(sTemp, sKeyWord)
    If nPosi = 0 Then
        nPosi = InStr(sTemp, LCase(sKeyWord))
    End If
        
    If nPosi = 0 Then Exit Function
    tfnGetUserPassward = Mid(sTemp, nPosi + Len(sKeyWord))
        
   ' #Else
   '     tfnGetUserPassward = t_oleObject.UserName
   ' #End If
    
End Function

Public Function GetRecordSet2nd(rsTemp As Recordset, szSql As String, _
                   Optional nDB As Variant, Optional szCalledFrom As Variant, _
                   Optional bShowErrow As Variant) As Long
    On Error GoTo SQLError
    
    If IsMissing(nDB) Then
       nDB = nDB_REMOTE
    End If
    
    Select Case nDB
       'Case nDB_LOCAL
       
       '  Set rsTemp = dbLocal.OpenRecordset(szSql, dbOpenSnapshot)
         
       '  If rsTemp.RecordCount > 0 Then
       '     rsTemp.MoveLast
       '     rsTemp.MoveFirst
       '  End If
       Case nDB_REMOTE
       
         Set rsTemp = t_dbMainDatabase2nd.OpenRecordset(szSql, dbOpenSnapshot, dbSQLPassThrough)
         
    End Select
    GetRecordSet2nd = rsTemp.RecordCount
    Exit Function
SQLError:
    If IsMissing(szCalledFrom) Then
       szCalledFrom = ""
    End If
    If IsMissing(bShowErrow) Then
       bShowErrow = True
    End If
    GetRecordSet2nd = -1
    tfnErrHandler "GetRecordSet2nd," & szCalledFrom, szSql, bShowErrow
    On Error GoTo 0
End Function

