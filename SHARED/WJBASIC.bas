Attribute VB_Name = "modWJBASIC"
' Weigong    02/03/97
' (1)The following codes are varibles,constants and subroutines used
' in most of the Weigong's projects(Mostly in SM modules and not
' used in later projects since this kind of Mod is easy to be messed up)
' (2)Since it is already messed up, I am changing the name of the module
' Now and if a problem encounter with the old "MYBASIC", it may be someone
' else. Use this one. 10/28/98
'
' Set myform=me in form_load
'*************************************************************************************
Option Explicit
'*****************************************
#If Win32 Then
    'Try to avoid GPF on exit tests (take out these codes if ..)
    'WJ 12/18/98
Public nMainFormLoaded_GPF As Integer
Public Const Loaded_Flag = 99
#End If
'*****************************************
Public Const ADD_EDIT_MSG = "Select Add, Edit or Exit."
Public Const DEL_WARNING_MSG = " Are you sure to delete the current record? "
Public Const DELALL_WARNING_MSG = "This will delete records in several screens." _
                         & " Are you sure to delete them all?"
Public Const DELROW_WARNING_MSG = " Are you sure you want to delete the current row? "
Public Const Edit_ERR_MSG = " No Record Available To Be Edited"
Public Const REFRESH_WARNING_MSG = "Data has not been saved. Are you sure you want to refresh?"


'varibles and constants
Public myForm As Form  ' "set myform=me" in form_load
Public Const DATA_INIT As Integer = 0
Public Const DATA_LOADED As Integer = 1
Public Const DATA_CHANGED As Integer = 2
Public nDataStatus As Integer 'data loaded,inti,changed flag
'Public bUpdateTable As Boolean  'almost never use

Public Const TEMP_ar_customer As String = "tmp_ar_customer" ' temp tables
Public Const TEMP_ar_altname As String = "tmp_ar_altname"
Public Const TEMP_sys_prft_ctr As String = "tmp_sys_prft_ctr" 'sometimes for Union customer
Public Const TEMP_p_altname As String = "tmp_p_altname"

Public Const SCROLL_BAR_WIDTH As Integer = 250

Public dbLocal As DataBase
Public Const nDB_LOCAL As Integer = 0
Public Const nDB_REMOTE As Integer = 1

'This function is used for testing if a line text contains
'any useful characters. Sometimes in Editbox, it contains
'lots of returns and we do not want them.
Public Function fnIsBlank(ByVal sTest As String) As Boolean
    Dim i As Integer
    Dim nCount As Integer
    Dim c As String
    
    fnIsBlank = True
    sTest = Trim(sTest)
    
    nCount = Len(sTest)
    If nCount >= 1 Then
        For i = 1 To nCount
            c = Mid(sTest, i, 1)
            If Asc(c) <> 10 And Asc(c) <> 13 Then
                fnIsBlank = False
                Exit Function
            End If
        Next
    End If
End Function
Public Function fnUcase(ByVal kCode As Integer) As Integer
    fnUcase = Asc(UCase(Chr(kCode)))
End Function
Public Function fnGetField(x As Variant) As String
     If IsNull(x) Then
        fnGetField = ""
     Else
        fnGetField = Trim(x)
     End If
End Function
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

'This is almost exactly the same as tfnSQLstring except
'szParameter is trimed
Public Function MyStr(ByVal szParameter As Variant, Optional vNoQuotes) As String
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

'This is not used very often. Used as compatible with old codes
'Note also use fnGetfiled(rs!...) might be faster!
Public Function GetFieldData(rs As Recordset, Optional szField As Variant) As String
   If IsMissing(szField) Then
        GetFieldData = fnGetField(rs.Fields(0))
   Else
        GetFieldData = fnGetField(rs.Fields(szField))
   End If
End Function
Public Function fnIsTableOutThere(ByVal szTableName As String) As Boolean
    Dim szSQL As String
    Dim rsTemp As Recordset
    
    On Error GoTo ErrorTrap:
    szSQL = "SELECT * FROM " & szTableName & " WHERE 1<>1"
    Set rsTemp = t_dbMainDatabase.OpenRecordset(szSQL, dbOpenSnapshot, dbSQLPassThrough)
    fnIsTableOutThere = True
    Exit Function
ErrorTrap:
    fnIsTableOutThere = False
    Err.Clear
    On Error GoTo 0
End Function
' this function also returns record count
Public Function GetRecordSet(rsTemp As Recordset, szSQL As String, _
                   Optional nDB As Variant, Optional szCalledFrom As Variant, _
                   Optional bShowErrow As Variant) As Long
    On Error GoTo SQLError
    
    If IsMissing(nDB) Then
       nDB = nDB_REMOTE
    End If
    
    Select Case nDB
       Case nDB_LOCAL
       
         Set rsTemp = dbLocal.OpenRecordset(szSQL, dbOpenSnapshot)
         If rsTemp.RecordCount > 0 Then
            rsTemp.MoveLast
            rsTemp.MoveFirst
         End If
       Case nDB_REMOTE
       
         Set rsTemp = t_dbMainDatabase.OpenRecordset(szSQL, dbOpenSnapshot, dbSQLPassThrough)
         #If Win32 Then
            If rsTemp.RecordCount > 0 Then
               rsTemp.MoveLast
               rsTemp.MoveFirst
            End If
         #End If
    End Select
    
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
    tfnErrHandler "GetRecordSet," & szCalledFrom, szSQL, bShowErrow
    On Error GoTo 0
End Function

' this function also returns record count
Public Function GetRecordCount(szSQL As String, _
                   Optional nDB As Variant, Optional szCalledFrom As Variant, _
                   Optional bShowErrow As Variant) As Long
    
    Dim rsTemp As Recordset
    
    On Error GoTo SQLError
    
    If IsMissing(nDB) Then
       nDB = nDB_REMOTE
    End If
    
    Select Case nDB
       Case nDB_LOCAL
       
         Set rsTemp = dbLocal.OpenRecordset(szSQL, dbOpenSnapshot)
         If rsTemp.RecordCount > 0 Then
            rsTemp.MoveLast
            rsTemp.MoveFirst
         End If
       Case nDB_REMOTE
       
         Set rsTemp = t_dbMainDatabase.OpenRecordset(szSQL, dbOpenSnapshot, dbSQLPassThrough)
         #If Win32 Then
            If rsTemp.RecordCount > 0 Then
               rsTemp.MoveLast
               rsTemp.MoveFirst
            End If
         #End If
      
    End Select
    GetRecordCount = rsTemp.RecordCount
    rsTemp.Close
    Exit Function
SQLError:
    If IsMissing(szCalledFrom) Then
       szCalledFrom = ""
    End If
    If IsMissing(bShowErrow) Then
       bShowErrow = True
    End If
    GetRecordCount = -1
    tfnErrHandler "GetRecordCount," & szCalledFrom, szSQL, bShowErrow
    On Error GoTo 0
End Function
Public Function fnExecuteSQL(szSQL As String, Optional nDB As Variant, _
                Optional szCalledFrom As Variant, Optional bShowError As Variant) As Boolean
                
      Dim szMsg As String
      
      On Error GoTo SQLError
      
      If IsMissing(nDB) Then
       nDB = nDB_REMOTE
      End If
    
      Select Case nDB
        
        Case nDB_LOCAL
           dbLocal.Execute szSQL
        Case nDB_REMOTE
           t_dbMainDatabase.ExecuteSQL szSQL
      End Select
      
      fnExecuteSQL = True
      Exit Function
      
SQLError:
      fnExecuteSQL = False
      If IsMissing(szCalledFrom) Then
         szCalledFrom = ""
      End If
      If IsMissing(bShowError) Then
         bShowError = True
      End If
      tfnErrHandler "fnExecuteSQL, " & szCalledFrom, szSQL, bShowError
      On Error GoTo 0
End Function
Public Sub SelectIt(Box As Textbox)
    On Error GoTo errTrap
    
    Box.SelStart = 0
    Box.SelLength = Len(RTrim(Box.Text))
    Exit Sub
    
errTrap:
    On Error GoTo 0
End Sub
Public Sub subSetFirstFocus(ParamArray arryControls())
    Dim i As Integer
    
    On Error Resume Next
    
    For i = 0 To UBound(arryControls)
        If arryControls(i).Enabled Then
            arryControls(i).SetFocus
            Exit For
        End If
    Next
End Sub

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
Public Sub subEnableDropBtn(ctlDropDown As Control, bYesNo As Boolean)
    #If Win16 Then
        If bYesNo Then
            ctlDropDown.Picture = LoadResPicture(DROPDOWN_UP, vbResBitmap)
        Else
            ctlDropDown.Picture = LoadResPicture(DROPDOWN_DOWN, vbResBitmap)
        End If
        ctlDropDown.Enabled = bYesNo
    #Else
        If bYesNo Then
            ctlDropDown.Picture = frmContext.LoadPicture(DROPDOWN_UP)
        Else
            ctlDropDown.Picture = frmContext.LoadPicture(DROPDOWN_DOWN)
        End If
    
    #End If
End Sub
Public Sub subEnableSearchbtn(cmdButton As Control, bYesNo As Boolean)
    #If Win16 Then
        If bYesNo Then
            cmdButton.Picture = LoadResPicture(SEARCH_UP, vbResBitmap)
        Else
            cmdButton.Picture = LoadResPicture(SEARCH_DOWN, vbResBitmap)
        End If
        
     #Else
        If bYesNo Then
            cmdButton.Picture = frmContext.LoadPicture(SEARCH_UP)
        Else
            cmdButton.Picture = frmContext.LoadPicture(SEARCH_DOWN)
        End If
     #End If
     
     cmdButton.Enabled = bYesNo
End Sub
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
'test the first nterms in vdata to see if szTest is already in vdata
Public Function IsAlreadyListed(vData(), ByVal nTerms As Long, _
                         ByVal nThisRow As Long, ByVal szTest As String) As Boolean
                          
     Dim i As Long
     
     IsAlreadyListed = False
     For i = 0 To nTerms - 1
        If i <> nThisRow Then
            If fnGetField(vData(i)) = Trim(szTest) Then
               IsAlreadyListed = True
               Exit For
            End If
        End If
     Next i
End Function

' This is for Master customer only (smaller than ar_altname)
' create a temp table for ar_altname
' two fields only:szNumber(an_customer(unique)),szName (first + last Name)
Public Function fnCreateTempAR_CUSTOMER(szNumber As String, szName As String) As Boolean
    
    Dim strSQL As String
    
    On Error GoTo errCreateTable
    strSQL = "CREATE TEMP TABLE " & TEMP_ar_customer & " (" & szNumber & " INTEGER, " & szName & " CHAR(60))"
    t_dbMainDatabase.ExecuteSQL strSQL
    
    On Error GoTo errInsertRecords
    strSQL = "INSERT INTO " & TEMP_ar_customer & " SELECT an_customer, TRIM(an_name) || ', ' || TRIM(an_first_name) FROM ar_altname, ar_customer WHERE (an_customer = cust_customer) AND (TRIM (an_name) <> '' AND TRIM(an_first_name) <>'')"
    t_dbMainDatabase.ExecuteSQL strSQL
    strSQL = "INSERT INTO " & TEMP_ar_customer & " SELECT an_customer, an_name FROM ar_altname, ar_customer WHERE (an_customer = cust_customer) AND (an_first_name IS NULL OR TRIM(an_first_name) ='')"
    t_dbMainDatabase.ExecuteSQL strSQL
    strSQL = "INSERT INTO " & TEMP_ar_customer & " SELECT an_customer, an_first_name FROM ar_altname, ar_customer WHERE (an_customer = cust_customer) AND (an_name IS NULL OR TRIM(an_name) ='') AND (TRIM(an_first_name) <>'')"
    t_dbMainDatabase.ExecuteSQL strSQL
    
    'Sam Zheng on 07/29/2004: Only get active customers.
    'But for those master customers, they should be always active according to
    'the current design. The user has no way to set them inactive in
    'VB program ARFALTNM.
    'I put this one line of code here in case later we change the mind!
    'tfnGetActiveAltCustomers TEMP_ar_customer, szNumber  '<<<---
    'end of Sam's message
    
    fnCreateTempAR_CUSTOMER = True
extCreateSearchTable:
    On Error GoTo 0
    Exit Function
errCreateTable:
    'MsgBox "Can not create temporary table for searching.", vbOKOnly + vbCritical, App.Title
    'Err.Clear
    tfnErrHandler "fnCreateTempAR_CUSTOMER", strSQL
    fnCreateTempAR_CUSTOMER = False
    Resume extCreateSearchTable
errInsertRecords:
    'MsgBox "Can not insert records into temporary table.", vbOKOnly + vbCritical, App.Title
    'Err.Clear
    tfnErrHandler "fnCreateTempAR_CUSTOMER", strSQL
    On Error Resume Next
    strSQL = "DROP TABLE " & TEMP_ar_customer
    t_dbMainDatabase.ExecuteSQL strSQL
    fnCreateTempAR_CUSTOMER = False
    Resume extCreateSearchTable
End Function

'create a temp table for ar_altname
' two fields only:szNumber(an_customer(unique)),szName (first + last Name)
Public Function fnCreateTempAR_ALTNAME(szNumber As String, _
                                       szName As String, _
                                       Optional szExcludeInactiveFlag As String = "N") As Boolean
    
    Dim szSQL As String
    On Error GoTo LetUsGo:
    szSQL = "DROP TABLE " & TEMP_ar_altname
    t_dbMainDatabase.ExecuteSQL szSQL
LetUsGo:
    On Error GoTo errCreateTable
    szSQL = "CREATE TEMP TABLE " & TEMP_ar_altname & " (" & szNumber & " INTEGER, " & szName & " CHAR(60))"
    t_dbMainDatabase.ExecuteSQL szSQL
    
    On Error GoTo errInsertRecords
    ' 3 cases
    szSQL = " SELECT an_customer, TRIM(an_name) || ', ' || TRIM(an_first_name) FROM " _
           & "ar_altname WHERE (TRIM (an_name) <> '' AND TRIM(an_first_name) <>'')"
    szSQL = "INSERT INTO " & TEMP_ar_altname & szSQL
    t_dbMainDatabase.ExecuteSQL szSQL
     
    szSQL = " SELECT an_customer, an_name FROM ar_altname WHERE " _
          & "(an_first_name IS NULL OR TRIM(an_first_name) ='')"
    szSQL = "INSERT INTO " & TEMP_ar_altname & szSQL
    t_dbMainDatabase.ExecuteSQL szSQL
    
    szSQL = " SELECT an_customer, an_first_name FROM ar_altname WHERE " _
        & "(an_name IS NULL OR TRIM(an_name) ='') AND (TRIM(an_first_name) <>'')"
    szSQL = "INSERT INTO " & TEMP_ar_altname & szSQL
    t_dbMainDatabase.ExecuteSQL szSQL
    
    'Sam Zheng on 07/29/2004: Exclude inactive customers
    'the normal value is 'N', 'O', 'B'.
    ' 'I' simply means not exclude--> 'I'nclude every number!
    If szExcludeInactiveFlag = "I" Then
        tfnGetActiveAltCustomers
    Else
        tfnGetActiveAltCustomers TEMP_ar_altname, szNumber, szExcludeInactiveFlag
    End If
    '''''
    
    fnCreateTempAR_ALTNAME = True
extCreateSearchTable:
    On Error GoTo 0
    Exit Function
errCreateTable:
    tfnErrHandler "fnCreateTempAR_ALTNAME", szSQL
    fnCreateTempAR_ALTNAME = False
    Resume extCreateSearchTable
errInsertRecords:
    tfnErrHandler "fnCreateTempAR_ALTNAME", szSQL
    On Error Resume Next
    szSQL = "DROP TABLE " & TEMP_ar_altname
    t_dbMainDatabase.ExecuteSQL szSQL
    fnCreateTempAR_ALTNAME = False
    Resume extCreateSearchTable
End Function
'create a temp table for ar_altname
' two fields only:szNumber(an_customer(unique)),szName (first + last Name)
Public Function fnCreateTempP_ALTNAME(szNumber As String, szName As String) As Boolean
    
    Dim szSQL As String
    
    On Error GoTo LetUsGo:
    szSQL = "DROP TABLE " & TEMP_p_altname
    t_dbMainDatabase.ExecuteSQL szSQL
LetUsGo:
    On Error GoTo errCreateTable
    szSQL = "CREATE TEMP TABLE " & TEMP_p_altname & " (" & szNumber & " INTEGER, " & szName & " CHAR(60))"
    t_dbMainDatabase.ExecuteSQL szSQL
    
    On Error GoTo errInsertRecords
    ' 3 cases
    szSQL = " SELECT pn_alt, TRIM(pn_name) || ', ' || TRIM(pn_first_name) FROM " _
           & "p_altname WHERE (TRIM (pn_name) <> '' AND TRIM(pn_first_name) <>'')"
    szSQL = "INSERT INTO " & TEMP_p_altname & szSQL
    t_dbMainDatabase.ExecuteSQL szSQL
     
    szSQL = " SELECT pn_alt, pn_name FROM p_altname WHERE " _
          & "(pn_first_name IS NULL OR TRIM(pn_first_name) ='')"
    szSQL = "INSERT INTO " & TEMP_p_altname & szSQL
    t_dbMainDatabase.ExecuteSQL szSQL
    
    szSQL = " SELECT pn_alt, pn_first_name FROM p_altname WHERE " _
        & "(pn_name IS NULL OR TRIM(pn_name) ='') AND (TRIM(pn_first_name) <>'')"
    szSQL = "INSERT INTO " & TEMP_p_altname & szSQL
    t_dbMainDatabase.ExecuteSQL szSQL
    
    fnCreateTempP_ALTNAME = True
extCreateSearchTable:
    On Error GoTo 0
    Exit Function
errCreateTable:
    tfnErrHandler "fnCreateTempP_ALTNAME", szSQL
    fnCreateTempP_ALTNAME = False
    Resume extCreateSearchTable
errInsertRecords:
    tfnErrHandler "fnCreateTempP_ALTNAME", szSQL
    On Error Resume Next
    szSQL = "DROP TABLE " & TEMP_p_altname
    t_dbMainDatabase.ExecuteSQL szSQL
    fnCreateTempP_ALTNAME = False
    Resume extCreateSearchTable
End Function
Public Function fnCreateTempSYS_PRFT_CTR(szNumber As String, szName As String) As Boolean
    On Error GoTo ErrorHandler
    Dim szSQL As String
    On Error GoTo LetUsGo:
    szSQL = "DROP TABLE " & TEMP_sys_prft_ctr
    t_dbMainDatabase.ExecuteSQL szSQL
LetUsGo:
    
    szSQL = "SELECT prft_ctr " & szNumber & ",prft_name " & szName & " FROM " _
         & "sys_prft_ctr INTO TEMP " & TEMP_sys_prft_ctr
    t_dbMainDatabase.ExecuteSQL szSQL
    fnCreateTempSYS_PRFT_CTR = True
    Exit Function
ErrorHandler:
    tfnErrHandler "fnCreateTempSYS_PRFT_CTR", szSQL
    fnCreateTempSYS_PRFT_CTR = False
    On Error Resume Next
    szSQL = "DROP TABLE " & TEMP_sys_prft_ctr
    t_dbMainDatabase.ExecuteSQL szSQL
    On Error GoTo 0
End Function
Public Function fnGetSysParm(ByVal nbr As String) As String
    Dim strSQL As String
    Dim rsTemp As Recordset
    
    fnGetSysParm = ""
    strSQL = "SELECT parm_field FROM sys_parm WHERE parm_nbr=" & tfnSQLString(nbr)
    
    If GetRecordSet(rsTemp, strSQL, , "fnGetSysParm") > 0 Then
        fnGetSysParm = fnGetField(rsTemp!parm_field)
    End If
End Function
Public Function IsFactorDate(ByVal sDate As String, Optional VErrorCode) As Boolean
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

Public Function fnCreateTemp_Small(szTableName As String, _
            ByVal szFieldName, ParamArray ArrayValues() As Variant) As Boolean
   Dim szSQL As String
   Dim k As Integer
   
   fnCreateTemp_Small = False
   On Error GoTo errCreateTable
   
   szSQL = "CREATE TEMP TABLE " & szTableName & "(" & szFieldName & " CHAR(1))"
   t_dbMainDatabase.ExecuteSQL szSQL

   On Error GoTo errInsertRecords
   For k = 0 To UBound(ArrayValues)
      szSQL = "INSERT INTO " & szTableName & "(" & szFieldName & ") VALUES (" _
           & MyStr(ArrayValues(k)) & ")"
      t_dbMainDatabase.ExecuteSQL szSQL
   Next
   fnCreateTemp_Small = True
extCreateSearchTable:
    On Error GoTo 0
    Exit Function
errCreateTable:
    tfnErrHandler "fnCreateTemp_Small", szSQL
    Resume extCreateSearchTable
errInsertRecords:
    tfnErrHandler "fnCreateTemp_Small", szSQL
    On Error Resume Next
    szSQL = "DROP TABLE " & szTableName
    t_dbMainDatabase.ExecuteSQL szSQL
    
    Resume extCreateSearchTable
   
End Function

#If Win16 Then ' we no longer use these subs later
    ' the following subroutines depend on template form
    Public Sub subEnableAdd(bYesNo As Boolean)
        With myForm
            .cmdAddBtn.Enabled = bYesNo
            .mnuAdd.Enabled = bYesNo
        End With
    End Sub
    
    Public Sub subEnableDelete(bYesNo As Boolean)
        With myForm
            .cmdDeleteBtn.Enabled = bYesNo
            .mnuDelete.Enabled = bYesNo
        End With
    End Sub
    Public Sub subEnableEdit(bYesNo As Boolean)
       With myForm
            .cmdEditBtn.Enabled = bYesNo
            .mnuEdit.Enabled = bYesNo
       End With
    End Sub
    Public Sub subEnableUpdateInsert(bYesNo As Boolean)
        myForm.cmdUpdateInsertBtn.Enabled = bYesNo
        myForm.mnuUpdateInsert.Enabled = bYesNo
    End Sub

#End If



