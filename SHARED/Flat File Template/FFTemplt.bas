Attribute VB_Name = "FFTemplate"
'Avaliable useful public functions:
'1. For data file
'   1a. Open file:  fnPrepareFile(FILENAME)    Return true if successful
'   1b. Read Line:  fnGetLine   return a line
'   1c. Test EOF:   fnEOF       return true if End Of File
'   1d. CloseFile   subCloseFile    does not return a value

'2. For log:
    'subWriteLog
    
'3. For Output:
'   3a. Setup:      fnSetupPrinter (Orientation)    return true if successful
'   3b. Set Title:  subSetTitle (Title) does not return a value
'   3c. Add a Line: subAddLine (sLine)  does not return a value
'   3d. Print:      fnStartPrint (Info)     return true if successful
'       When 3c is used to buffer lines, 3d must be called to print out the contents
'   3e. Print:      fnSendToPrinter(AryLines(), Info)   return true if successful
'       When 3e is used, you must pass in the array where the data is held

'4  For Database Operation:
'   4a. Execute a SQL:   fnExecuteSQL strSQL, caller, DB     return true if successful
'   4b. Open a record:   fnOpenRecord strSQL, caller, Msg, DB    return the recordset

'5  String Parsing:
'   subParseString sParms(), Source, Delimiter


'Modification to this module:
'1. Sub:        subGetInfo to supply the correct info(Module ID and Caption)
'2. Function    fnAllowStandalone   Return true if allow standalone mode
'                                   Otherwise return false
'3. Sub         subProcessFile      Put codes and function calls here for processing

Option Explicit

Public Function fnAllowStandalone() As Boolean
    fnAllowStandalone = False
End Function

Public Sub subGetInfo(aryInfo() As String)

    ReDim aryInfo(1)
    
    aryInfo(0) = "ZZFMBNK"
    aryInfo(1) = "Murphy Bank File Import"
    
End Sub

Public Sub subProcessFile(sFile As String)

    Dim sLine As String
    
    If fnPrepareFile(sFile) Then
        While Not fnEOF
            sLine = fnGetLine
            'Process this line here
            '... ...
'            subWriteLog "... ..."
        Wend
        subCloseFile
    End If
    
End Sub

