VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSVParse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'VERSION 1.0 CLASS
'BEGIN
'  MultiUse = -1  'True
'  Persistable = 0  'NotPersistable
'  DataBindingBehavior = 0  'vbNone
'  DataSourceBehavior = 0   'vbNone
'  MTSTransactionMode = 0   'NotAnMTSObject
'End
'Attribute VB_Name = "CSVParse"
'Attribute VB_GlobalNameSpace = False
'Attribute VB_Creatable = True
'Attribute VB_PredeclaredId = False
'Attribute VB_Exposed = False
Option Explicit

Dim FSO                                         ' A File Scripting object
Dim TS                                          ' A Text Stream
Private mvarFileName        As String           ' Path to the target file
Private mvarFieldCount      As Integer          ' Count of fields in this record
Private mvarStatus          As Boolean          ' Are we able to provide data?
Private strRecord           As String           ' Current record
Private Fields()            As String           ' Current record field array
Private strErrMsg           As String           ' Last error message
Private mvarFieldSeperator  As String           ' User defined field seperator

'------------------------------------------------------------------------------
'   CSV Parser
'   This class handles retrieving elements from a CSV (C_omma S_eperated V_alues)
'   file. In the CSV file each line is a record and each field in the record is
'   seperated from its neighbor by a delimiter character. The character is usually
'   a comma (,) but can be any character.
'
'   This class requires a reference to the MS Scripting Runtime.
'
'   Create an instance of the class (Dim CSVP as New CSVParse)
'   Set the FieldSeperator property if it is not comma.
'   Set the FileName property using the full path to the target file.
'      a. Read the Status property. If it is false, the file was not
'         accessed so call the GetErrorMessage function to retrieve the
'         descripition of the problem
'   Process the file as follows:
'
'       While CSVP.LoadNextLine = True
'           MyString = CSVP.GetField(n) <- for each field you want to read
'                                          where n is the field number where
'           .                              1 is the first field.
'           .
'           .
'       Wend
'----------------------------------------------------------------------------
'   Return the message string
Public Function GetErrorMessage() As String
    GetErrorMessage = strErrMsg
End Function

'   Set the field delimiter character. Default is the comma.
Public Property Let FieldSeperator(ByVal vData As String)
    mvarFieldSeperator = Trim(vData)
End Property

Public Property Get FieldSeperator() As String
    FieldSeperator = mvarFieldSeperator
End Property

'   Internal status set
Private Property Let Status(ByVal vData As Boolean)
    mvarStatus = vData
End Property

'   Tell caller the status
Public Property Get Status() As Boolean
    Status = mvarStatus
End Property

'   Give out the number of fields in this record
Public Property Get FieldCount() As Integer
    FieldCount = mvarFieldCount
End Property

'   Set the target file name
Public Property Let FileName(ByVal vData As String)
    mvarFileName = vData                            'Set the file path
    mvarStatus = AccessTargetFile()                 'Open it
End Property

'   Reads the next line of text and parses it into fields array
Public Function LoadNextLine() As Boolean
    On Error GoTo LoadNextLine_Err
    If TS.AtEndOfStream Then
        LoadNextLine = False
        Exit Function
    End If
    strRecord = TS.ReadLine
    
'============================================================================
'===   code added to handle quotes and commas which are intended as data  ===
'===
'==
    Const DATA_QUOTE_CHAR = "_dataquote_"
    Const DATA_DELIM = "_datadelim_"
    
    
    Dim bAbortOnError As Boolean
    Dim iFirstQuote As Long
    Dim iSecondQuote As Long
    Dim i As Long
    
    Dim sField As String
    Dim sNewField As String
    
    strRecord = Replace(strRecord, FieldSeperator & """""""", FieldSeperator & DATA_QUOTE_CHAR) 'data quote as first char in quoted field
    strRecord = Replace(strRecord, """""""" & FieldSeperator, DATA_QUOTE_CHAR & FieldSeperator) 'data quote as last char in quoted field
    strRecord = Replace(strRecord, """""", DATA_QUOTE_CHAR) 'other data quote within field
    
    bAbortOnError = False
    Do Until InStr(1, strRecord, """") = 0 Or bAbortOnError
        iFirstQuote = InStr(1, strRecord, """")
        iSecondQuote = InStr(iFirstQuote + 1, strRecord, """")
        
        If iFirstQuote = 0 Or iSecondQuote = 0 Then
            bAbortOnError = True
        Else
            sField = Mid(strRecord, iFirstQuote, (iSecondQuote - iFirstQuote) + 1)
            sNewField = Replace(sField, FieldSeperator, DATA_DELIM) 'replace commas with placeholders
            sNewField = Replace(sNewField, """", "") 'remove quotes
            strRecord = Replace(strRecord, sField, sNewField, , 1) 'replace first instance with placeholder version
        End If
    Loop
    

'--- original ---
    ReDim Fields(0)
    Fields = Split(strRecord, FieldSeperator)       'Break out the string of fields
    mvarFieldCount = UBound(Fields) + 1             '# of fields available
'--- original ---

    
    'now replace placeholders with the real characters again
    For i = 0 To mvarFieldCount - 1
        Fields(i) = Replace(Fields(i), DATA_QUOTE_CHAR, """")
        Fields(i) = Replace(Fields(i), DATA_DELIM, FieldSeperator)
    Next
    
'==
'===
'============================================================================
    
    LoadNextLine = True
    Exit Function
LoadNextLine_Err:
    LoadNextLine = False                        'Should happen at EOF...
End Function

'   Pass back the specified field
Public Function GetField(FieldNum As Integer) As String
    If FieldNum < 1 Or FieldNum > FieldCount Then
        GetField = ""
    Else
        GetField = Trim(Fields(FieldNum - 1))
    End If
End Function

'   Open the target file...
Private Function AccessTargetFile() As Boolean
    Const ForReading = 1, ForWriting = 2, ForAppending = 8

    On Error Resume Next
    TS.Close                    'Close if open.
    On Error GoTo AccessTargetFile_Err
    Status = True
    strErrMsg = ""
    Set TS = FSO.OpenTextFile(mvarFileName, ForReading)
    AccessTargetFile = True
    Status = True
    Exit Function
AccessTargetFile_Err:
    strErrMsg = CStr(Err.number) & " " & Err.Description & " in AccessTargetFile."
    AccessTargetFile = False
End Function

'   Normal VB class initialization
Private Sub Class_Initialize()
    Status = False                          'Not open for business yet
    FieldSeperator = ","                    'Default to comma
    mvarFileName = ""                       'No file yet
    Set FSO = CreateObject("Scripting.FileSystemObject")
End Sub

'   Normal VB class termination
Private Sub Class_Terminate()
    Set FSO = Nothing                       'Clean up - destroy objects
    Set TS = Nothing
End Sub


