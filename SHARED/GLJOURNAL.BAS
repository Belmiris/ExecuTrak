Attribute VB_Name = "modGLJOURNAL"
'*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=
'   *** *** *** *** *** *** *** *** ***  N O T I C E ! ! *** *** *** *** *** *** *** *** ***
'*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=
' This module is intended to be shared/used among multiple various projects. Avoid making any
' special application-specific enhancements to this module that may break other programs that
' are already using this module.
'*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=

Option Explicit

Public NewGLEntryNum As Long

Public frmMain As Form
Private Const LocalDB As Integer = 0
Private Const RemoteDB As Integer = 1

'Private iIndex As Integer                            'index used for everything
Private lNewBatchNbr As Long                         'number determined to be the next batch number available
Public iHoldRefPd As Integer
Public iPostingPeriod As Integer        'holds period for posting
Public szHoldSeries As String        'holds series for reversals
Private szErrorType As String
Private vPostingPC As Variant       'holds profit center being posted
Private vPostingAccount As Variant       'holds Account being posted

Private dNewPeriodAmt0 As Double
Private dNewPeriodAmt1 As Double
Private dNewPeriodAmt2 As Double
Private dNewPeriodAmt3 As Double
Private dNewPeriodAmt4 As Double
Private dNewPeriodAmt5 As Double
Private dNewPeriodAmt6 As Double
Private dNewPeriodAmt7 As Double
Private dNewPeriodAmt8 As Double
Private dNewPeriodAmt9 As Double
Private dNewPeriodAmt10 As Double
Private dNewPeriodAmt11 As Double
Private dNewPeriodAmt12 As Double
Private dNewPeriodAmt13 As Double
Private dNewPeriodAmt14 As Double

Private Const nPC_COL As Integer = 0
Private Const nACCOUNT_COL As Integer = 1
Private Const nAMOUNT_COL As Integer = 2
Private Const nTEXT_COL As Integer = 3
Private Const nMAX_COL = 3

Private aryJournal()

Private sPARM_2203 As String
Public sPARM_3101 As String

Public txtTransDate As String
Public txtTransNbr As String
Private rsTempLocalJE As Recordset
Private rsTempLocalText As Recordset
Private rsMyRecordset As Recordset

'AccumNPost module
Private rsGLMaster As Recordset
Private iPeriodAccum As Integer
Private rsTempLocalJEAccum As Recordset
'

'return error message if any
Public Function GLJOURNAL_ResetVariables() As String
    ReDim aryJournal(nMAX_COL, 0)
    fnSetUpTempDatabase
    subSetsPARM_2203_3101
End Function

'return error message if any
Public Sub GLJOURNAL_AddJournalEntry(ByVal nPrftCtr As Integer, ByVal lAcct As Long, ByVal dfAmt As Double, sText As String)
    Dim nIndex As Integer
    
    On Error Resume Next
    nIndex = -1
    nIndex = UBound(aryJournal, 2)
    On Error GoTo errTrap
    
    If nIndex < 0 Then
        nIndex = 0
        ReDim aryJournal(nMAX_COL, nIndex)
    ElseIf nIndex >= 0 And fnCStr(aryJournal(nPC_COL, nIndex)) <> "" Then
        nIndex = nIndex + 1
        ReDim Preserve aryJournal(nMAX_COL, nIndex)
    End If
    
    aryJournal(nPC_COL, nIndex) = nPrftCtr
    aryJournal(nACCOUNT_COL, nIndex) = lAcct
    aryJournal(nAMOUNT_COL, nIndex) = dfAmt
    aryJournal(nTEXT_COL, nIndex) = sText
    
    Exit Sub

errTrap:
    tfnErrHandler "AddJournalEntry"
End Sub

Public Function GLJOURNAL_CheckTransDate(sTransDate As String) As String
    Dim sErrMsg1 As String
    Dim sErrMsg2 As String
    
    txtTransDate = sTransDate
    
    sErrMsg1 = fnCheckTransDate()
    
    If sErrMsg1 <> "" Then
        sErrMsg1 = "Work Order Date (" + txtTransDate + "): " + sErrMsg1
        
        'try today's date
        txtTransDate = tfnFormatDate(Date)
        sErrMsg2 = fnCheckTransDate()
    
        If sErrMsg2 <> "" Then
            sErrMsg2 = "Today's Date (" + txtTransDate + "): " + sErrMsg2
        End If
    End If
    
    If sErrMsg2 <> "" Then
        GLJOURNAL_CheckTransDate = sErrMsg1 + vbCrLf + sErrMsg2
    End If
End Function

'return error message if any
Public Function GLJOURNAL_PostGLJOURNAL(sTransDate As String) As String
    Screen.MousePointer = vbHourglass
    GLJOURNAL_PostGLJOURNAL = GLJOURNAL_CheckTransDate(sTransDate)
    
    If GLJOURNAL_PostGLJOURNAL <> "" Then
        Screen.MousePointer = vbDefault
        Exit Function
    End If
    
    Screen.MousePointer = vbHourglass
    If Not fnValidateGLs() Then
        GLJOURNAL_PostGLJOURNAL = "GL Master Record(s) not found."
        Screen.MousePointer = vbDefault
        Exit Function
    End If
    
    Screen.MousePointer = vbHourglass
    If Not fnGetTransNumber() Then
        GLJOURNAL_PostGLJOURNAL = "Failed to generate Transaction Number."
        Screen.MousePointer = vbDefault
        Exit Function
    End If
    
    txtTransNbr = lNewBatchNbr
    
    Screen.MousePointer = vbHourglass
    If Not fnWriteJournalRec() Then
        GLJOURNAL_PostGLJOURNAL = "Failed to store G/L and Journal Entry."
        Screen.MousePointer = vbDefault
        Exit Function
    End If
    
    Screen.MousePointer = vbHourglass
    If Not fnProcessJournalEntries() Then
        GLJOURNAL_PostGLJOURNAL = "Failed to process G/L and Journal Entry."
        Screen.MousePointer = vbDefault
        Exit Function
    End If
    Screen.MousePointer = vbDefault
    
End Function

Private Function fnWriteJournalRec() As Boolean
    Const SUB_NAME = "fnWriteJournalRec"
    
    Dim szSQLStmt As String
    Dim nColCount As Integer
    Dim lRowCount As Long
    Dim i As Long
    Dim szMsg As String
    Dim aryText() As String
    Dim intText As Integer
    
    fnWriteJournalRec = False
    
    lRowCount = -1
    
    On Error Resume Next
    lRowCount = UBound(aryJournal, 2)
    
    If lRowCount < 0 Then
        Exit Function
    End If
    
    If fnCStr(aryJournal(nPC_COL, 0)) = "" Then
        Exit Function
    End If
    
    On Error GoTo INSERT_LOCAL_ERROR
    
    For i = 0 To lRowCount
        
        If fnCStr(aryJournal(nPC_COL, i)) <> "" Then
            szSQLStmt = "INSERT INTO GLJournal (iEntryNumber,iAccount,iProfitCenter,"
            szSQLStmt = szSQLStmt & "iPostingPeriod,dPostingAmount,cFlag,dTransactionDate,"
            szSQLStmt = szSQLStmt & "iTransactionShift,iBatchNumber,cAcctDesc,btext)"
            szSQLStmt = szSQLStmt & " VALUES (" & i + 1 & ","
            szSQLStmt = szSQLStmt & aryJournal(nACCOUNT_COL, i) & ","
            szSQLStmt = szSQLStmt & aryJournal(nPC_COL, i) & ","
            szSQLStmt = szSQLStmt & iHoldRefPd & ","
            szSQLStmt = szSQLStmt & tfnRound(aryJournal(nAMOUNT_COL, i), 2) & ", 'D',"
            szSQLStmt = szSQLStmt & tfnDateString(txtTransDate, True) & ", 1,"
            szSQLStmt = szSQLStmt & txtTransNbr & ","
            szSQLStmt = szSQLStmt & tfnSQLString("G/L Description") & ","
            szSQLStmt = szSQLStmt & IIf(fnCStr(aryJournal(nTEXT_COL, i)) <> "", "'Y'", "'N'") & ")"
                
            If Not fnInsertNewRecordLoc(szSQLStmt, "GLJournal") Then
                Exit Function
            End If
            
            If aryJournal(nTEXT_COL, i) <> "" Then
                If InStr(1, aryJournal(nTEXT_COL, i), vbCrLf) > 0 Then
                    aryText = Split(aryJournal(nTEXT_COL, i), vbCrLf)
                    For intText = 0 To UBound(aryText)
                        szSQLStmt = "INSERT INTO GLText (iLink,iSequenceNumber,cText,cSrc) VALUES ("
                        szSQLStmt = szSQLStmt & i + 1 & ", 1, " & tfnSQLString(aryText(intText)) & ", 'GJ')"
                        If Not fnInsertNewRecordLoc(szSQLStmt, "GLText") Then
                            Exit Function
                        End If
                    Next
                Else
                    szSQLStmt = "INSERT INTO GLText (iLink,iSequenceNumber,cText,cSrc) VALUES ("
                    szSQLStmt = szSQLStmt & i + 1 & ", 1, " & tfnSQLString(aryJournal(nTEXT_COL, i)) & ", 'GJ')"
                    If Not fnInsertNewRecordLoc(szSQLStmt, "GLText") Then
                        Exit Function
                    End If
                End If
            End If
            
        End If
    Next i

    fnWriteJournalRec = True
        
RETURN_HERE:
    Exit Function

INSERT_LOCAL_ERROR:
    tfnErrHandler SUB_NAME, szSQLStmt
        Resume RETURN_HERE
End Function

Private Function fnProcessJournalEntries() As Boolean
    Const SUB_NAME = "fnProcessJournalEntries"
    
    Dim szSQLStmt As String
    Dim lSerialNumber As Long
    Dim dfNewAmount As Double
    Dim szNewFlag As String
    
    Set rsTempLocalJE = Nothing
    Set rsTempLocalText = Nothing
    dfNewAmount = 0
    szNewFlag = ""
    
    Dim lEntryNumber As Long
    
    If Not fnAccumNPost() Then
        fnProcessJournalEntries = False
        Exit Function
    End If
    
    'set up sql and read the general ledger
    'journal entries on the local database
    szSQLStmt = "SELECT * FROM GLJournal ORDER BY iEntryNumber"
    
    On Error GoTo READ_ERROR
    Set rsTempLocalJE = dbLocal.OpenRecordset(szSQLStmt, dbOpenDynaset)
    On Error GoTo 0
                                                    'check the record count
    If rsTempLocalJE.EOF Then
        MsgBox "fnProcessJournalEntries - " + Err.Description + vbCrLf + vbCrLf _
                 & "No records were found to insert. " + vbCrLf _
                 & "Please try your transactions again.  " + vbCrLf + vbCrLf _
                 & "If this error continues contact your System Administrator!", vbCritical
        fnProcessJournalEntries = False
        GoTo GET_OUT
    End If
 
    '*************************************************************
    'records found continue with adding records
    While Not rsTempLocalJE.EOF             'entry number is stored as 0 in informix
        If rsTempLocalJE!iAccount <> -4578 Then    'do not process this number it is a dumby entry
            If rsTempLocalJE!dPostingAmount < 0 Then            'if neg. entry
                dfNewAmount = rsTempLocalJE!dPostingAmount * -1 'reverse to positive
                szNewFlag = "C"                                 'and store them as a credit
            Else                                                'OR
                dfNewAmount = rsTempLocalJE!dPostingAmount
                szNewFlag = "D"                                 'store them as a debit
            End If                                              'Posting function
                                                                'will check the accounts normal
            'write the Journal Entry and retrieve serial number needed for
            'writing the text lines
            lSerialNumber = fnWriteJournalEntry(rsTempLocalJE!iPostingPeriod, rsTempLocalJE!dTransactionDate, dfNewAmount, szNewFlag)
            NewGLEntryNum = lSerialNumber
            
            If lSerialNumber = -1 Then
                GoTo GET_OUT
            End If
            
            'write journal entry text records
            lEntryNumber = tfnRound(rsTempLocalJE!iEntryNumber)
            
            If rsTempLocalJE!btext = "Y" Then         'verify that text exists
                fnWriteTextEntry lSerialNumber, lEntryNumber 'write the text records remote database
            Else
                If sPARM_2203 = "Y" Then
                    fnWriteTextEntry lSerialNumber, lEntryNumber
                End If
            End If
        End If

        rsTempLocalJE.MoveNext
    Wend
    
    fnProcessJournalEntries = True
    
RETURN_HERE:
    Screen.MousePointer = vbDefault
    Exit Function

READ_ERROR:                       'Error occurred during a read
    tfnErrHandler SUB_NAME, szSQLStmt

GET_OUT:
    fnProcessJournalEntries = False
    Resume RETURN_HERE
End Function

Private Function fnWriteJournalEntry(iReferencePeriod As Integer, _
                                      iNextDate As Date, _
                                      dAmount As Double, _
                                      szFlag As String) As Long
        
    Const SUB_NAME = "fnWriteJournalEntry"
    
    'add new record to the gl journal table
    'and then retrieve the serial number of that record
    Dim szMsg As String
    Dim szSQLStmt As String
    Dim szSQLstmtLocal As String
    
    fnWriteJournalEntry = -1
            
    'build SQL to store record
    'note the 0 in the first field is always needed
    'by informis to store the serial number correctly
    szSQLStmt = "INSERT INTO gl_journal VALUES (0, " _
                & rsTempLocalJE!iAccount & ", " _
                & rsTempLocalJE!iProfitCenter & ", " _
                & iReferencePeriod & ", " _
                & dAmount & ", " _
                & tfnSQLString(szFlag) & ", " _
                & tfnSQLString(iNextDate) & ", " _
                & rsTempLocalJE!iTransactionShift & ", " _
                & rsTempLocalJE!iBatchNumber & ")"
    
    '*******       'store the record          ****
    On Error GoTo INSERT_ERROR
    t_dbMainDatabase.ExecuteSQL szSQLStmt
    On Error GoTo 0
    
    'now you have to go back to the table and retrieve the serieal number of the
    'record just stored.  It should aways be the highest entry number with the
    'current batch #
    szSQLStmt = "SELECT Max(glj_entry_nbr) AS gljEntryNbr FROM gl_journal " _
              & "WHERE  glj_batch = " & rsTempLocalJE!iBatchNumber
    
    If fnGotOneRecord(szSQLStmt, "GLText", False) Then
        fnWriteJournalEntry = rsMyRecordset!gljEntryNbr   'pass back serial #
    Else
        MsgBox szErrorType + vbCrLf + vbCrLf _
            & "Unable to retrieve the serial number from" + vbCrLf _
            & "the last Journal entry written." + vbCrLf _
            & "Print may not have all entries.  Please" + vbCrLf _
            & "Contact your System Administrator!", vbCritical
    
        Exit Function
    End If
        
RETURN_HERE:
    Exit Function

INSERT_ERROR:
    tfnErrHandler SUB_NAME, szSQLStmt
    Resume RETURN_HERE
End Function

Private Function fnWriteTextEntry(ByVal lSerialNumber As Long, ByVal iEntryNumber As Long) As Boolean
    'find all records on local text table where
    'the link is the same as the entry number on the
    'local GlJournal table
    Const SUB_NAME = "fnWriteTextEntry"
    
    Dim szActiveTable As String                     'used in error messages
    Dim szSQLStmt As String
    fnWriteTextEntry = False
    
    szSQLStmt = "SELECT * FROM GLText WHERE " _
              & "iLink = " & iEntryNumber 'rsTempLocalJE!iEntryNumber
    szSQLStmt = szSQLStmt & " ORDER BY iSequenceNumber" 'WJ
    
    'The seq no can start with any number (it depends on on how many times
    'the after col edit is fired. Ha,ha! The easy way to fix it is here) WJ 05/04/99
    Dim nNewSeqNumber As Integer
    
    nNewSeqNumber = 0
    
    Set rsTempLocalText = Nothing
    szActiveTable = "GLText - local database"
    
    On Error GoTo READ_ERROR
    Set rsTempLocalText = dbLocal.OpenRecordset(szSQLStmt, dbOpenDynaset)
    On Error GoTo 0
                            
    'check to make sure there are records
    If rsTempLocalText.RecordCount > 0 Then
        While Not rsTempLocalText.EOF
            nNewSeqNumber = nNewSeqNumber + 1
            szSQLStmt = "INSERT INTO gl_jrnl_text VALUES (" _
                      & lSerialNumber & ", " _
                      & nNewSeqNumber & ", " _
                      & tfnSQLString(rsTempLocalText!cText) & ", ""GJ"")"
            
            szActiveTable = "gl_jrnl_text - remote database"
            
            On Error GoTo INSERT_ERROR
            t_dbMainDatabase.ExecuteSQL szSQLStmt
            On Error GoTo 0
            
            rsTempLocalText.MoveNext
        Wend
    End If
     
    fnWriteTextEntry = True
    
RETURN_HERE:
    Exit Function
    
INSERT_ERROR:
    tfnErrHandler SUB_NAME, szSQLStmt
    fnSetUpTempDatabase
    Resume RETURN_HERE
        
READ_ERROR:                       'Error occurred during a read
    tfnErrHandler SUB_NAME, szSQLStmt
    fnSetUpTempDatabase
    Resume RETURN_HERE
End Function

Private Function fnGetTransNumber() As Boolean
    'reads current gl_journal file for highest batch number
    'add 1 to it then check the sys_data_lock to see it already
    'exists, if so adds 1 to it and reads again and returns the good number.
    Const SUB_NAME As String = "fnGetTransNumber"
    
    Dim szSQLStmt As String
    Dim szSQLStmtStatic As String
    Dim bCheck As Boolean
    Dim bError As Boolean
    
    bError = False
    
    'set up sql and read the last batch number on gl_journal and up it by 1
    szSQLStmt = "SELECT Max(glj_batch) HighBatch " _
              & "FROM gl_journal"
              
    Screen.MousePointer = vbHourglass
    lNewBatchNbr = fnReadForTransNbr(szSQLStmt) + 1
    
    If lNewBatchNbr = 0 Then
        Exit Function
    End If
    
    bCheck = True
    szSQLStmtStatic = "SELECT count (*) nCount " _
              & "FROM sys_data_lock WHERE sdl_prog = 'glegjrnl'"
              
    Do      'read the sys_data_lock table for a match, upping the number each time
            'until no record is found
        szSQLStmt = szSQLStmtStatic & " and sdl_data = '" & Trim(Str(lNewBatchNbr)) & "'"

        If fnCheckTblForOneOrMoreRecords(szSQLStmt, "sys_data_lock", True, bError) Then
           bCheck = True                            'Set value of flag to False.
           lNewBatchNbr = lNewBatchNbr + 1          'up the new journal number by one each time through
        Else
            If bError Then         'there was a problem reading table
                Screen.MousePointer = vbDefault
                Exit Function
            End If
                    
            bCheck = False      'good number found
        End If
    Loop Until bCheck = False                'Exit loop immediately
    
    If fnWriteNewBatchNbr() Then
        fnGetTransNumber = True
    End If
    
    Screen.MousePointer = vbDefault
End Function

Private Function fnReadForTransNbr(szSQLStmt As String) As Long
    'this function reads a table using input SQL and returns
    'true if only one record is found
    Const SUB_NAME = "fnReadForTransNbr"
    
    Dim rsTemp As Recordset
    
    fnReadForTransNbr = -1
    
    On Error GoTo READ_ERROR
    Set rsTemp = t_dbMainDatabase.OpenRecordset(szSQLStmt, DBOPEN_SNAPSHOT, SQL_PASSTHROUGH)
    On Error GoTo 0
    
    If rsTemp.RecordCount = 1 Then
        fnReadForTransNbr = tfnRound(rsTemp!HighBatch)
    End If
    
    Exit Function
READ_ERROR:                       'Error occurred during the delete
    tfnErrHandler SUB_NAME, szSQLStmt
End Function

Private Function fnCheckTblForOneOrMoreRecords(szSQLStmt As String, _
                                               szTableName As String, _
                                               bShowMsg As Boolean, _
                                               bError As Boolean) As Boolean
    'this function reads a table using input SQL looking for a count of records
    'it returns true if one or more records are found
    Dim rsTemp As Recordset
    
    szErrorType = ""
    bError = False
    
    On Error GoTo READ_ERROR
    Set rsTemp = t_dbMainDatabase.OpenRecordset(szSQLStmt, DBOPEN_SNAPSHOT, SQL_PASSTHROUGH)
    On Error GoTo 0
    
    If rsTemp!nCount = 0 Then
        fnCheckTblForOneOrMoreRecords = False
    Else
        fnCheckTblForOneOrMoreRecords = True
    End If

RETURN_HERE:
    Exit Function

READ_ERROR:                       'Error occurred during the read
    szErrorType = "fnGotOneRecord - " & Err.Description
    If bShowMsg Then
        MsgBox "fnCheckTblForOneOrMoreRecords - " + Err.Description + vbCrLf + vbCrLf _
                 & "An error has occurred while reading the " & szTableName + vbCrLf & "table.  " _
                 & "Please exit the screen and try it again.  " + vbCrLf _
                 & "If this error continues contact your System Administrator!", vbCritical
       
    End If
        
    fnCheckTblForOneOrMoreRecords = False
    bError = True
    Resume RETURN_HERE
End Function

Private Function fnWriteNewBatchNbr() As Boolean
    Dim szSQLStmt As String
    fnWriteNewBatchNbr = False
    
    'build insert sql
    szSQLStmt = "INSERT INTO sys_data_lock VALUES "
    szSQLStmt = szSQLStmt & "(" & tfnSQLString(Trim$(tfnGetUserName)) & ", ""glegjrnl"", "
    szSQLStmt = szSQLStmt & """" & Trim(Str(lNewBatchNbr)) & """)"
    
    'now go insert the record
    If fnInsertNewRecord(szSQLStmt, "sys_data_lock") Then                              'Add a new record to sys_prft_ctr
        fnWriteNewBatchNbr = True
    End If
End Function

Private Function fnInsertNewRecord(szSQLStmt As String, szTblName As String) As Boolean
    'This function will process the query to insert a record
    Const SUB_NAME = "fnInsertNewRecord"
    
    On Error GoTo ADD_ERROR                                'Add a new record to sys_prft_ctr
    t_dbMainDatabase.ExecuteSQL szSQLStmt
    fnInsertNewRecord = True
    On Error GoTo 0

RETURN_HERE:
    Exit Function

ADD_ERROR:      'Error during the add
    tfnErrHandler SUB_NAME, szSQLStmt
    Resume RETURN_HERE
End Function

Private Function fnInsertNewRecordLoc(szSQLStmt As String, szTblName As String) As Boolean
    Const SUB_NAME = "fnInsertNewRecordLoc"
    
    On Error GoTo INSERT_LOCAL_ERROR
    dbLocal.Execute szSQLStmt
    On Error GoTo 0
    fnInsertNewRecordLoc = True
    
RETURN_HERE:
    Exit Function

INSERT_LOCAL_ERROR:      'Error during the insert
    tfnErrHandler SUB_NAME, szSQLStmt
    Resume RETURN_HERE
End Function

Private Function fnCStr(vTemp, Optional bRTrim As Boolean = False) As String
    If IsNull(vTemp) Then
        fnCStr = ""
    Else
        fnCStr = IIf(bRTrim, RTrim(vTemp), Trim(vTemp))
    End If
End Function

Private Function fnSetUpTempDatabase() As Boolean
        'deletes all the old records that may have been left hanging around
    fnDeleteTable "GLJournal"      'delete any old journal records
    fnDeleteTable "GLText"      'delete any old text records
    
End Function

Private Function fnDeleteTable(szTableName As String) As Boolean
    Const SUB_NAME = "fnDeleteTable"
            'delete local table
    Dim szDeleteSQL As String
    fnDeleteTable = False
    
    'delete all old local records
    On Error GoTo DELETE_ERROR
    szDeleteSQL = "DELETE * FROM " & szTableName
    dbLocal.Execute szDeleteSQL   'make sure all previous records are gone
    On Error GoTo 0
    
    fnDeleteTable = True
    Exit Function
    
DELETE_ERROR:                       'Error occurred during the delete
    tfnErrHandler SUB_NAME, szDeleteSQL
End Function

Private Sub subSetsPARM_2203_3101()
    'looks for parameter 2203 if not there or field value is not "Y" then
    'text is not repeated so sPARM_2203 is left as false
    Dim szSQLStmt As String
    
    szSQLStmt = "SELECT parm_field FROM sys_parm WHERE parm_nbr = 2203"
    If fnGotOneRecord(szSQLStmt, "sys_parm", False) Then
        sPARM_2203 = UCase(fnCStr(rsMyRecordset!parm_field))
    End If
    
    szSQLStmt = "SELECT parm_field FROM sys_parm WHERE parm_nbr = 3101"
    If fnGotOneRecord(szSQLStmt, "sys_parm", False) Then
        sPARM_3101 = fnCStr(rsMyRecordset!parm_field)
    End If
End Sub

Private Function fnCheckTransDate() As String
    Dim szSQLStmt As String
    Dim sStatus As String

    
    szSQLStmt = "SELECT glp_status, glp_series, glp_glpd, glp_refpd " _
        & "FROM gl_period WHERE " & tfnDateString(txtTransDate, True) & " BETWEEN glp_beg_dt AND glp_end_dt "
    
    If fnGotOneRecord(szSQLStmt, "gl_journal", False) Then
        sStatus = rsMyRecordset!glp_status
        
        If sStatus = "C" Then
            fnCheckTransDate = "Accounting period has been closed for this date."
        Else
            If sStatus = "O" Or sStatus = "G" Or sStatus = "W" Then
                iHoldRefPd = rsMyRecordset!glp_refpd
                szHoldSeries = rsMyRecordset!glp_series
                iPostingPeriod = rsMyRecordset!glp_glpd
            Else
                fnCheckTransDate = "Unknown G/L status (" & sStatus & ") for the accounting period."
            End If
        End If
    Else
        fnCheckTransDate = "Accounting period has not been created for this date."
    End If
End Function

Private Function fnGotOneRecord(szSQLStmt As String, szTableName As String, bShowMsg As Boolean) As Boolean
    'this function reads a table using input SQL and returns
    'true if only one record is found
    Dim rsTemp As Recordset
    
    szErrorType = ""            'passes back the error message is there is one
    
    Set rsMyRecordset = Nothing
    
    On Error GoTo READ_ERROR
    Set rsTemp = t_dbMainDatabase.OpenRecordset(szSQLStmt, DBOPEN_SNAPSHOT, SQL_PASSTHROUGH)
    On Error GoTo 0
    
    If rsTemp.EOF Then
        Set rsMyRecordset = Nothing
        fnGotOneRecord = False
    Else
        Set rsMyRecordset = rsTemp
        fnGotOneRecord = True
    End If

RETURN_HERE:
    Screen.MousePointer = vbDefault
    Exit Function

READ_ERROR:                       'Error occurred during the read
    szErrorType = "fnGotOneRecord - " & Err.Description
    If bShowMsg Then
        MsgBox "fnGotOneRecord - " + Err.Description + vbCrLf + vbCrLf _
             & "An error has occurred while reading the " & szTableName & "" + vbCrLf _
             & "table.  Please exit the screen and try again.  " + vbCrLf _
             & "If this error continues contact your System Administrator!", vbOKOnly + vbCritical
    End If
    
    fnGotOneRecord = False
    Resume RETURN_HERE
End Function

Private Function rsOneRecord(szSQLStmt As String, szTableName As String) As Recordset
        'this function reads a table using input SQL and returns
        'a recordset in rsOneRecord
    Set rsOneRecord = Nothing
    
    On Error GoTo READ_ERROR
    Set rsOneRecord = t_dbMainDatabase.OpenRecordset(szSQLStmt, DBOPEN_SNAPSHOT, SQL_PASSTHROUGH)
    On Error GoTo 0
    
RIGHT_HERE:
    Exit Function
    
READ_ERROR:                       'Error occurred during the read
    szErrorType = "rsOneRecord - " + Err.Description
    tfnErrHandler "rsOneRecord", szSQLStmt
    Resume RIGHT_HERE
End Function

'''''''''''''''''''''''''''  FUNCTIONS IN AccumNPost module  '''''''''''''''''''''''''''''''''''''''
Private Function fnUpdateGLMaster(ByVal iPostingPC As Integer, ByVal lPostingAccount As Long, szSeries As String) As Boolean
    Const SUB_NAME = "fnUpdateGLMaster"
        'updated the master
    Dim ncounter As Integer
    Dim szMsg As String
    Dim rsProc As Recordset
    Dim szProcParms As String
    Dim szProcSQL As String
    Dim szSQLStmt As String
    Dim szSQLstmtLocal As String
    
    ncounter = 0
    fnUpdateGLMaster = False
    Screen.MousePointer = vbHourglass     'set mouse

REBUILD_IT:
    szProcParms = tfnSQLString(Trim$(szSeries)) & ", " & lPostingAccount _
    & ", " & iPostingPC _
    & ", " & dNewPeriodAmt0 & ", " & dNewPeriodAmt1 & ", " & dNewPeriodAmt2 _
    & ", " & dNewPeriodAmt3 & ", " & dNewPeriodAmt4 & ", " & dNewPeriodAmt5 _
    & ", " & dNewPeriodAmt6 & ", " & dNewPeriodAmt7 & ", " & dNewPeriodAmt8 _
    & ", " & dNewPeriodAmt9 & ", " & dNewPeriodAmt10 & ", " & dNewPeriodAmt11 _
    & ", " & dNewPeriodAmt12 & ", " & dNewPeriodAmt13 & ", " & dNewPeriodAmt14
   
SUBMIT_IT:
    Set rsProc = Nothing
    szProcSQL = "EXECUTE PROCEDURE upd_gl_master(" & szProcParms & ");"
    On Error GoTo PROC_PROBLEM
    Set rsProc = t_dbMainDatabase.OpenRecordset(szProcSQL, dbOpenSnapshot, dbSQLPassThrough)
    
    If rsProc.Fields(0) = 0 Then    'record is locked by someone else
        If ncounter < 1000 Then
            ncounter = ncounter + 1
            If Not frmMain Is Nothing Then
                frmMain.tfnSetStatusBarMessage "Waiting for locked record. Please wait..."
            End If
            GoTo SUBMIT_IT
        Else
            tfnErrHandler SUB_NAME, szProcSQL, rsProc.Fields(rsProc.Fields.Count - 1), False
            If MsgBox("GLEGJRNL - fnUpdateGLMaster" + vbCrLf + vbCrLf _
                 & "Series " & szSeries + vbCrLf _
                 & "Profit Center " & iPostingPC + vbCrLf _
                 & "Account " & lPostingAccount + vbCrLf _
                 & "This record is locked by another user....." + vbCrLf _
                 & "I can continue trying or backout all " + vbCrLf _
                 & "Journal Entries for this transaction" + vbCrLf + vbCrLf _
                 & "         CONTINUE? ", vbYesNo + vbQuestion) = vbYes Then
                 
                ncounter = 0
                GoTo SUBMIT_IT
            Else
                fnUpdateGLMaster = False
                GoTo RETURN_HERE
            End If
        End If
    Else
        If Not frmMain Is Nothing Then
            frmMain.tfnSetStatusBarMessage "Processing Journal Entries. Please wait..."
        End If
        fnUpdateGLMaster = True
    End If
    
RETURN_HERE:
    Screen.MousePointer = vbDefault
    Exit Function

PROC_PROBLEM:
    tfnErrHandler SUB_NAME, szProcSQL
    fnUpdateGLMaster = False
    Resume RETURN_HERE
End Function

Private Function fnBuildTheAmount(dfNewAmount As Double) As String
    Dim szSQLStmt As String
    
    'case statement builds the middle of SQL depending
    'on what period the transaction is for
    Select Case iPeriodAccum
    Case 1
        dNewPeriodAmt1 = dfNewAmount
    Case 2
        dNewPeriodAmt2 = dfNewAmount
    Case 3
        dNewPeriodAmt3 = dfNewAmount
    Case 4
        dNewPeriodAmt4 = dfNewAmount
    Case 5
        dNewPeriodAmt5 = dfNewAmount
    Case 6
        dNewPeriodAmt6 = dfNewAmount
    Case 7
        dNewPeriodAmt7 = dfNewAmount
    Case 8
        dNewPeriodAmt8 = dfNewAmount
    Case 9
        dNewPeriodAmt9 = dfNewAmount
    Case 10
        dNewPeriodAmt10 = dfNewAmount
    Case 11
        dNewPeriodAmt11 = dfNewAmount
    Case 12
        dNewPeriodAmt12 = dfNewAmount
    Case 13
        dNewPeriodAmt13 = dfNewAmount
    End Select
End Function

Private Sub subLoadNewVariablies()
    dNewPeriodAmt0 = 0
    dNewPeriodAmt1 = 0
    dNewPeriodAmt2 = 0
    dNewPeriodAmt3 = 0
    dNewPeriodAmt4 = 0
    dNewPeriodAmt5 = 0
    dNewPeriodAmt6 = 0
    dNewPeriodAmt7 = 0
    dNewPeriodAmt8 = 0
    dNewPeriodAmt9 = 0
    dNewPeriodAmt10 = 0
    dNewPeriodAmt11 = 0
    dNewPeriodAmt12 = 0
    dNewPeriodAmt13 = 0
    dNewPeriodAmt14 = 0
End Sub

Private Function fnAccumNPost() As Boolean
    Const SUB_NAME = "fnAccumNPost"
    
    Dim dEntryAccum As Double
    Dim szSQLStmt As String
    
    fnAccumNPost = False
    
    fnDeleteTable "GLMasterBackOutParms"      'delete old backout records
    
    subLoadNewVariablies    'zero out amounts
    dEntryAccum = 0
    Set rsTempLocalJEAccum = Nothing
    vPostingPC = Null
    vPostingAccount = Null
    
    szSQLStmt = "SELECT * FROM GLJournal " _
        & "ORDER BY iProfitCenter, iAccount"
    
    On Error GoTo READ_ERROR
    Set rsTempLocalJEAccum = dbLocal.OpenRecordset(szSQLStmt, dbOpenDynaset)
    On Error GoTo 0
                                                    
    If rsTempLocalJEAccum.RecordCount = 0 Then
        MsgBox "fnAccumNPost - " + vbCrLf + vbCrLf _
                 & "No records were found to insert. " + vbCrLf _
                 & "This screen will terminate.  Please " + vbCrLf _
                 & "try your transactions again.  " + vbCrLf + vbCrLf _
                 & "If this error continues contact your System Administrator!", vbCritical
        Exit Function
    End If
    
    If rsTempLocalJEAccum.EOF Then
        MsgBox "fnAccumNPost - " + vbCrLf + vbCrLf _
            & "No records were found to insert. " + vbCrLf _
            & "This screen will terminate.  Please " + vbCrLf _
            & "try your transactions again.  " + vbCrLf + vbCrLf _
            & "If this error continues contact your System Administrator!", vbCritical
        
        Exit Function
    End If
 
    '*************************************************************
    'records found continue with posting
    iPeriodAccum = iPostingPeriod
    rsTempLocalJEAccum.MoveLast
    rsTempLocalJEAccum.MoveFirst
    
    While Not rsTempLocalJEAccum.EOF             'entry number is stored as 0 in informix
        If rsTempLocalJEAccum!iAccount <> -4578 Then    'do not process this number it is a dumby entry
            If IsNull(vPostingPC) Or IsNull(vPostingAccount) Then
                dEntryAccum = 0
                vPostingPC = rsTempLocalJEAccum!iProfitCenter
                vPostingAccount = rsTempLocalJEAccum!iAccount
            End If
            
            If vPostingPC <> rsTempLocalJEAccum!iProfitCenter _
              Or vPostingAccount <> rsTempLocalJEAccum!iAccount Then
                iPeriodAccum = iPostingPeriod
                subLoadNewVariablies    'zero out amounts
                'check for normal balance on master
                szSQLStmt = "SELECT * FROM gl_master " _
                    & "WHERE glm_series = " & tfnSQLString(szHoldSeries) _
                    & " AND glm_account = " & vPostingAccount _
                    & " AND glm_prft_ctr = " & vPostingPC
                          
                Set rsGLMaster = rsOneRecord(szSQLStmt, "gl_master")   'go get the gl_master record
                
                If rsGLMaster Is Nothing Then
                    GoTo ERROR_ROUTINE
                End If
                If rsGLMaster!glm_norm_bal = "C" Then
                    dEntryAccum = dEntryAccum * -1      'if credit reverse sign
                End If
                
                fnBuildTheAmount dEntryAccum           'load the period bucket being posted
                If Not frmMain Is Nothing Then
                    frmMain.tfnSetStatusBarMessage "Processing Journal Entries. Please wait..."
                End If

                If fnUpdateGLMaster(vPostingPC, vPostingAccount, szHoldSeries) Then
                   fnWriteBackoutEntry vPostingPC, vPostingAccount, szHoldSeries
                Else
                   GoTo ERROR_ROUTINE
                End If
                
                'put in new values
                vPostingPC = rsTempLocalJEAccum!iProfitCenter
                vPostingAccount = rsTempLocalJEAccum!iAccount
                dEntryAccum = 0  'zero out accumulator
            End If
                
            'accumulate the period value
            dEntryAccum = dEntryAccum + rsTempLocalJEAccum!dPostingAmount
        End If
        
        rsTempLocalJEAccum.MoveNext
        
        If rsTempLocalJEAccum.EOF Then
            iPeriodAccum = iPostingPeriod
            subLoadNewVariablies    'zero out amounts
                    
            'check for normal balance on master
            szSQLStmt = "SELECT * FROM gl_master " _
                      & "WHERE glm_series = " & tfnSQLString(szHoldSeries) _
                      & " AND glm_account = " & vPostingAccount _
                      & " AND glm_prft_ctr = " & vPostingPC
                      
            Set rsGLMaster = rsOneRecord(szSQLStmt, "gl_master")   'go get the gl_master record
            
            If rsGLMaster Is Nothing Then
                GoTo ERROR_ROUTINE
            End If
            If rsGLMaster!glm_norm_bal = "C" Then
                dEntryAccum = dEntryAccum * -1      'if credit reverse sign
            End If
            
            fnBuildTheAmount dEntryAccum           'load the period bucket being posted
            If Not frmMain Is Nothing Then
                frmMain.tfnSetStatusBarMessage "Processing Journal Entries. Please wait..."
            End If

            If fnUpdateGLMaster(vPostingPC, vPostingAccount, szHoldSeries) Then
               fnWriteBackoutEntry vPostingPC, vPostingAccount, szHoldSeries
            Else
               GoTo ERROR_ROUTINE
            End If
       End If
    Wend

    fnAccumNPost = True

RETURN_HERE:
    Exit Function
    
ERROR_ROUTINE:
    fnAccumNPost = False
    If Err.Number <> 0 Then
        tfnErrHandler SUB_NAME, szSQLStmt, False
    End If
    
    MsgBox "Entries will now be backed out", vbCritical
    
    If fnReverseEntries Then
        If Not frmMain Is Nothing Then
            frmMain.tfnSetStatusBarError "Entries have been backed out. Advise you cancel and try again"
        End If
        GoTo RETURN_HERE
    Else
       subCrashIt
       GoTo RETURN_HERE
    End If

READ_ERROR:
    MsgBox "fnAccumNPost - " + Err.Description + vbCrLf + vbCrLf _
        & "A Read error has occured on table GLJournal. " + vbCrLf _
        & "No records have been posted. " + vbCrLf _
        & "This screen will terminate.  Please " + vbCrLf _
        & "try your transactions again.  " + vbCrLf + vbCrLf _
        & "If this error continues contact your System Administrator!", vbCritical
        
        fnAccumNPost = False
End Function

Private Function fnWriteBackoutEntry(ByVal iPostingPC As Integer, ByVal lPostingAccount As Long, szSeries As String) As Boolean
    Dim szSQLStmt As String
    Dim szProcParms As String
    
    szProcParms = tfnSQLString(Trim$(szSeries)) & ", " & lPostingAccount _
        & ", " & iPostingPC _
        & ", " & (dNewPeriodAmt0 * -1) & ", " & (dNewPeriodAmt1 * -1) & ", " & (dNewPeriodAmt2 * -1) _
        & ", " & (dNewPeriodAmt3 * -1) & ", " & (dNewPeriodAmt4 * -1) & ", " & (dNewPeriodAmt5 * -1) _
        & ", " & (dNewPeriodAmt6 * -1) & ", " & (dNewPeriodAmt7 * -1) & ", " & (dNewPeriodAmt8 * -1) _
        & ", " & (dNewPeriodAmt9 * -1) & ", " & (dNewPeriodAmt10 * -1) & ", " & (dNewPeriodAmt11 * -1) _
        & ", " & (dNewPeriodAmt12 * -1) & ", " & (dNewPeriodAmt13 * -1) & ", " & (dNewPeriodAmt14 * -1)
    
    szSQLStmt = "INSERT INTO GLMasterBackOutParms VALUES (""" & Trim$(szProcParms) & """,0)"
            
    On Error GoTo ERROR_ROUTINE
    dbLocal.Execute szSQLStmt                'executes the SQL
    On Error GoTo 0
    Exit Function
    
ERROR_ROUTINE:
    MsgBox "error writing to backout table"
    Resume Next
End Function

Private Function fnReverseEntries() As Boolean
    Const SUB_NAME = "fnReverseEntries"
    
    Dim szSQLStmt As String
    Dim rsTemp As Recordset
    Dim rsProc As Recordset
    Dim szProcSQL As String
    Dim ncounter As Long
    fnReverseEntries = True
    
    szSQLStmt = "SELECT * FROM GLMasterBackOutParms " _
        & "ORDER BY EntryNo"
    
    On Error GoTo READ_ERROR
    Set rsTemp = dbLocal.OpenRecordset(szSQLStmt, dbOpenDynaset)
    On Error GoTo 0
    If rsTemp.RecordCount = 0 Then
        Exit Function
    End If
    
    rsTemp.MoveFirst
    If rsTemp.EOF Then
        Exit Function
    End If
    While Not rsTemp.EOF
        
SUBMIT_IT:
        Set rsProc = Nothing
        szProcSQL = "EXECUTE PROCEDURE upd_gl_master(" & rsTemp!BackOutProcParm & ");"
        szSQLStmt = szProcSQL
        On Error GoTo PROC_PROBLEM
        
        Set rsProc = t_dbMainDatabase.OpenRecordset(szProcSQL, dbOpenSnapshot, dbSQLPassThrough)
        If rsProc.Fields(0) = 0 Then    'record is locked by someone else
            If ncounter < 4000 Then
                DoEvents
                ncounter = ncounter + 1
                GoTo SUBMIT_IT
            Else
                tfnErrHandler SUB_NAME, szSQLStmt, rsProc.Fields(rsProc.Fields.Count - 1)
                fnReverseEntries = False
            End If
        End If
        
        rsTemp.MoveNext
    Wend
   
RETURN_HERE:
    Exit Function

PROC_PROBLEM:
    tfnErrHandler SUB_NAME, szSQLStmt
    fnReverseEntries = False
    Resume RETURN_HERE

READ_ERROR:
    subCriticalMsgOK "fnUpdateGLMaster - " + Err.Description + vbCrLf + vbCrLf _
        & "An error has occurred while reading the backout entries on  " + vbCrLf _
        & "the GLMasterBackOutParms table." + vbCrLf _
        & "Please call support and inform them that " + vbCrLf _
        & "NO backout entries were processed" & fnCStr(rsTemp!EntryNo) & " had a problem." + vbCrLf _
        & "They will need to back out all entries"
    
    fnReverseEntries = False
    Resume RETURN_HERE
End Function

Private Sub subCrashIt()
    Dim szMsg As String
    
    szMsg = "This program has encountered a problem that" + vbCrLf _
        & "requires support or your System Administrator to" + vbCrLf _
        & "access tables that were created during this run." + vbCrLf + vbCrLf _
        & "Please DO NOT RUN THIS PROGRAM AGAIN until you are given" + vbCrLf _
        & "the OK.  Doing so may cause you to loose information" + vbCrLf _
        & "needed to bring your accounts back in balance!"
    
    subCriticalMsg szMsg
    
    Unload frmMain
End Sub

Private Sub subCriticalMsg(szMessageText As String)
    MsgBox szMessageText & vbCrLf & "If this error continues contact your System Administrator!", vbCritical
End Sub

Private Sub subCriticalMsgOK(szMessageText As String)
    'displays a critical message to the user no reply
    MsgBox szMessageText & vbCrLf & "If this error continues contact your " _
        & "System Administrator!", vbCritical
End Sub

Private Function fnValidGL(ByVal nPrftCtr As Integer, ByVal lAcct As Long, Optional sDesc As String = "") As Boolean
    Const SUB_NAME As String = "fnValidGL"
    
    Dim strSQL As String
    Dim rsTemp As Recordset
    
    strSQL = "SELECT glm_description FROM gl_master WHERE glm_series = " + tfnSQLString(szHoldSeries)
    strSQL = strSQL + " AND glm_prft_ctr = " & nPrftCtr
    strSQL = strSQL + " AND glm_account = " & lAcct
    
    If fnRecordset(rsTemp, strSQL, RemoteDB, SUB_NAME) <= 0 Then
        Exit Function
    End If
    
    sDesc = fnCStr(rsTemp!glm_description)
    
    fnValidGL = True
End Function

Private Function fnValidateGLs() As Boolean
    
    Dim szSQLStmt As String
    Dim lRowCount As Long
    Dim i As Long
    Dim sInvalidGLs As String
    
    fnValidateGLs = False
    
    lRowCount = UBound(aryJournal, 2) + 1
    
    If lRowCount < 1 Then
        Exit Function
    End If
    
    sInvalidGLs = ""
    
    For i = 0 To lRowCount - 1
        If fnCStr(aryJournal(nPC_COL, i)) <> "" And fnCStr(aryJournal(nACCOUNT_COL, i)) <> "" Then
            If fnCStr(aryJournal(nACCOUNT_COL, i)) <> "-4578" Then ' non posting acct
                If Not fnValidGL(aryJournal(nPC_COL, i), aryJournal(nACCOUNT_COL, i)) Then
                    If sInvalidGLs <> "" Then
                        sInvalidGLs = sInvalidGLs + vbCrLf
                    End If
                    
                    sInvalidGLs = sInvalidGLs + "Profit Center - " + fnCStr(aryJournal(nPC_COL, i)) _
                        + ", GL Account - " + fnCStr(aryJournal(nACCOUNT_COL, i))
                End If
           End If
        End If
    Next i
    
    If sInvalidGLs <> "" Then
        MsgBox "Unable to locate a GL Master record for the following Profit Center and GL Account:" _
            + vbCrLf + vbCrLf + sInvalidGLs + vbCrLf + vbCrLf _
            + "Add the records to GL Master and try again.", vbExclamation
        Exit Function
    End If
    
    fnValidateGLs = True
End Function

Private Function fnRecordset(rsTemp As Recordset, szSQL As String, _
                   Optional nDB As Variant, Optional szCalledFrom As Variant, _
                   Optional bShowErrow As Variant) As Long
    On Error GoTo SQLError
    If IsMissing(nDB) Then
       nDB = RemoteDB
    End If
    Select Case nDB
        Case LocalDB
            Set rsTemp = dbLocal.OpenRecordset(szSQL, dbOpenSnapshot)
        Case RemoteDB
            Set rsTemp = t_dbMainDatabase.OpenRecordset(szSQL, dbOpenSnapshot, dbSQLPassThrough)
    End Select
    If rsTemp.RecordCount > 0 Then
       rsTemp.MoveLast
       rsTemp.MoveFirst
    End If
    fnRecordset = rsTemp.RecordCount
    Exit Function
SQLError:
    If IsMissing(szCalledFrom) Then
        szCalledFrom = ""
    End If
    If IsMissing(bShowErrow) Then
        bShowErrow = False
    End If
    fnRecordset = -1
    tfnErrHandler "GetRecordSet," & szCalledFrom, szSQL, bShowErrow
    On Error GoTo 0
End Function


