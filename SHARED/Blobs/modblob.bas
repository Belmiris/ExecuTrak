Attribute VB_Name = "modBlob"
'##############################################################################
'# blob.bas
'# Written to allow storage of BLOB objects in a non-blob ISAM db.
'#
'# uses b64c.exe (from b64c.c)
'#
'# Requires table(s) in the following format:
'#create table blob_dunning_hdr
'#(
'#id serial,
'#Desc Char(20)
'#);
'#
'#create table blob_dunning_det
'#(
'#ser_lnk integer,
'#seq integer,
'#data char(72)
'#);
'#
'# Per Howard, we need a uniform table prefix since field names will be
'# the same (no unique field prefix).  The table prefix for blobs will be
'# "blob", so all tables will be as the following for dunning letter templates:
'# blob_dunning_hdr
'# blob_dunning_det
'#
'#The only table this library cares about is blob det.  ser_lnk links to a
'# record somewhere used to describe the blob.  seq keeps the lines in sequence,
'# and data is the actual mime64 encoded data.
'#
'# Also requires standard ExecuTrak Template environment (and DAO et al)
'##############################################################################
Option Explicit
'##############################################################################
'# Function: SaveBlob
'# Author: Robert Atwood
'# Date: 11/05/04
'# Arguments:
'#  sFilename (string), full path and name of file to store as blob
'#  sTableName (string), name of table to store file in
'#              NOTE: MUST be compatible with following structure:
'#              SER_LNK integer,
'#              SEQ integer,
'#              DATA char (74)
'#  lID (double), ID number to store to
'#  bOverWrite  (boolean), If TRUE, overwrites existing record.  If false,
'#                          causes an error if it detects an existing record.
'#  Returns: Boolean (True for success, false for failure)
'#  NOTE: There should be a unique index on ser_lnk and seq
'##############################################################################
Function SaveBlob(sFileName As String, sTableName As String, lID As Long, _
                  bOverWrite As Boolean)
    Dim sCommand As String
    Dim bSuccess As Boolean
    Dim sFileNameTarg As String
    Dim nFilePointer As Integer
    Dim lSequence As Long
    Dim strSQL As String
    Dim sTempLine As String
    Dim rsTemp As Recordset
    On Error GoTo ErrorOut
    
    bSuccess = False
    If Dir(sFileName) <> "" Then
        sFileNameTarg = RTrim(sFileName) + Trim(Str(App.ThreadID))
        bSuccess = True
    End If
    
    If bSuccess = True Then
        sCommand = """" + App.Path + "\b64.exe"" -e """ + sFileName + """ """ + sFileNameTarg + """"
        If Shell(sCommand) < 1 Then
            '#Failed to execute
            bSuccess = False
        Else
            bSuccess = True
        End If
    End If

    If bSuccess Then
        '#Delete existing entries, if any, if in overwrite mode
        If bOverWrite Then
            strSQL = "delete from " & sTableName & " where ser_lnk=" & _
                              Str(lID)
                              
            ModBlobExecuteSQL strSQL, "SaveBlob", False
            bSuccess = True
        Else
            strSQL = "select count(*) from " & sTableName & " where ser_lnk=" & _
                              Str(lID)
            If ModBlobGetRecordSet(rsTemp, strSQL, "ModBlob::SaveBlob", False) _
                                   > 0 Then
                If Not (rsTemp Is Nothing) Then
                    If rsTemp.Fields(0).value > 0 Then bSuccess = False
                Else
                    bSuccess = True
                End If
            End If
        End If
    End If
    If bSuccess Then
        '#Open file as text
        nFilePointer = FreeFile
        Open sFileNameTarg For Input As #nFilePointer
        lSequence = 0
        'Loop through each line of file, storing in tablename
        Do While Not (EOF(nFilePointer) And bSuccess)
            Line Input #nFilePointer, sTempLine
            'Remove CRLF from EOL
            sTempLine = Replace(sTempLine, vbCrLf, "")
            strSQL = "insert into " + sTableName + _
                   " values (" + Str(lID) + "," + Str(lSequence) + _
                   ",'" + sTempLine + "')"
            bSuccess = ModBlobExecuteSQL(strSQL, "ModBlob::SaveBlob", False)
            lSequence = lSequence + 1
        Loop
        Close #nFilePointer
    End If
    SaveBlob = bSuccess
    Kill sFileNameTarg
    CleanUp (rsTemp)
Exit Function
ErrorOut:
    bSuccess = False
    Kill sFileNameTarg
    CleanUp rsTemp
    On Error GoTo 0
End Function

'##############################################################################
'# Function: GetBlob
'# Author: Robert Atwood
'# Date: 11/05/04
'# Arguments:
'#  sFilename (string), full path and name of file to store as blob
'#  sTableName (string), name of table to store file in
'#              NOTE: MUST be compatible with following structure:
'#              ID integer,
'#              SEQ integer,
'#              DATA char (74)
'#  lID (double), ID number to store to
'#  Returns: Boolean (True for success, false for failure)
'##############################################################################
Function GetBlob(sFileName As String, sTableName As String, lID As Long)
    Dim sCommand As String
    Dim bSuccess As Boolean
    Dim sSourceFileName As String
    Dim rsTemp As Recordset
    Dim nFilePointer As Integer
    Dim lRowCount As Integer
    Dim strSQL As String
    
    bSuccess = False
On Error GoTo ErrorOut

    '# Create temp file to hold data from table in
    sSourceFileName = RTrim(sFileName) + Trim(Str(App.ThreadID))
    '# get recordset from table here...
    strSQL = "Select * from " + sTableName + " where ser_lnk = " & Str(lID) & _
           " order by seq"
    lRowCount = ModBlobGetRecordSet(rsTemp, strSQL, "GetBlob", False)
    If lRowCount > 0 Then
        nFilePointer = FreeFile
        Open sSourceFileName For Output As #nFilePointer
        While Not rsTemp.EOF
            Print #nFilePointer, rsTemp!Data
            rsTemp.MoveNext
        Wend
        Close #nFilePointer
        bSuccess = True
    Else
        bSuccess = False
    End If

    If bSuccess Then
        '# Eyebrowse are for VB quote excape.
        sCommand = """" + App.Path + "\b64.exe"" -d """ + sSourceFileName + """ """ + sFileName + """"
        '# Run b64c.exe from appdir
        If Shell(sCommand) < 1 Then
            '#Failed to execute
            bSuccess = False
        Else
            bSuccess = True
        End If
    '# Open file and start inserting...
    End If
    GetBlob = bSuccess
    Kill sSourceFileName
    CleanUp (rsTemp)
    Exit Function
ErrorOut:
    bSuccess = False
    Kill sSourceFileName
    CleanUp rsTemp
    On Error GoTo 0

End Function
'##############################################################################
'# Function: DeleteBlob
'# Author: Robert Atwood
'# Date: 11/05/04
'# Arguments:
'#  sTableName (string), name of table to delete blob from
'#  lID (double), ID number of blob to delete
'#  Returns: Boolean (True for success, false for failure)
'##############################################################################
Function DeleteBlob(sFileName As String, sTableName As String, lID As Long)
    Dim strSQL As String
    Dim bSuccess As Boolean
    bSuccess = False
On Error GoTo ErrorOut
    strSQL = "delete from " + sTableName + " where ser_lnk=" + Str(lID)
    bSuccess = ModBlobExecuteSQL(strSQL, "ModBlob::DeleteBlob", False)
    DeleteBlob = bSuccess
    Exit Function
ErrorOut:
    bSuccess = False
End Function

'###############################################################################
'# ModBlobExecuteSQL- included because we can't #$@! count on standard execution
'#                  function in these @#!$ programs
'###############################################################################
Private Function ModBlobExecuteSQL(strSQL As String, sCalledFrom As String, _
                             Optional bShowError As Boolean = True) As Boolean
    
    Const SUB_NAME As String = "fnExecuteSQL"
    
    On Error GoTo ErrorAccessRecords
    t_dbMainDatabase.ExecuteSQL strSQL
    
    ModBlobExecuteSQL = True
    Exit Function

ErrorAccessRecords:
    tfnErrHandler SUB_NAME + "," + sCalledFrom, strSQL, bShowError
    ModBlobExecuteSQL = False
End Function

'###############################################################################
'# ModBlobGetRecordSet- included because we can't #$@! count on standard execution
'#                  function in these @#!$ programs
'###############################################################################
Private Function ModBlobGetRecordSet(rsTemp As Recordset, szSQL As String, _
                    Optional szCalledFrom As Variant, _
                   Optional bShowErrow As Variant) As Long
    On Error GoTo SQLError
    
    Set rsTemp = t_dbMainDatabase.OpenRecordset(szSQL, dbOpenSnapshot, dbSQLPassThrough)
    If rsTemp.RecordCount > 0 Then
       rsTemp.MoveLast
       rsTemp.MoveFirst
    End If
    ModBlobGetRecordSet = rsTemp.RecordCount
    Exit Function
SQLError:
    If IsMissing(szCalledFrom) Then
        szCalledFrom = ""
    End If
    If IsMissing(bShowErrow) Then
        bShowErrow = True
    End If
    ModBlobGetRecordSet = -1
    tfnErrHandler "GetRecordSet," & szCalledFrom, szSQL, bShowErrow
    On Error GoTo 0
End Function

