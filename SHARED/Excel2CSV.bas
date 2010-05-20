Attribute VB_Name = "modExcel2CSV"
Option Explicit
'***********************************************************'
'
' Copyright (c) 2010 FACTOR, A Division of W.R.Hess Company
'
' Module name       : modExcel2CSV (excel2csv.bas)
'
' Revision          :
'
' Date              : 05/18/2010
'
' Programmer(s)     : Paul Jeanquart
'
' Specification     : Generically convert an Excel to CSV
'                       Uses OLE automation and reads individual cells
'                       to avoid pitfalls with "save as csv"
'                       and to give greater control to programmer
'
'***********************************************************'

' Main call   Excel2CSV
'               returns true/false true if all files processed, false if any errors
'               argument 1 path of file(s) to process
'               argument 2 file spec, can be a single file, wildcards or multiples seperated by a ;
'               argument 3 True/False to recurse sub folders
'               argument 4 True/False to process only the first sheet in each work book
'               argument 5 ArchiveDir a sub folder to put the xls files after processing
'               argument 6 sErrorMsg textual error message if any

Public Function Excel2CSV(ByVal path As String, ByVal filespec As String, Optional RecurseDirs As Boolean, Optional FirstSheetOnly As Boolean = False, Optional ArchiveDir As String = "bak", Optional sErrorMsg As String) As Boolean
    Dim cFiles As Collection
    Dim vFile As Variant
    
    Set cFiles = GetAllFiles(path, filespec, RecurseDirs)
    
    
    If Not DirExists(path & "\" & ArchiveDir) Then
        If tfnConfirm("Backup directory: " & path & "\" & ArchiveDir & vbCrLf & "Does not exist, do you want to create it?") Then
            If Not MakeDirPath(path & "\" & ArchiveDir) Then
                sErrorMsg = "Could not create backup directory: " & path & "\" & ArchiveDir
                Exit Function
            End If
        Else
            sErrorMsg = "Backup Directory does not exist: " & path & "\" & ArchiveDir
            Exit Function
        End If
    End If
    
    For Each vFile In cFiles
        If fnExcelConvertOneFile(vFile, FirstSheetOnly, ArchiveDir, sErrorMsg) Then
             Exit Function
        End If
    Next

    Excel2CSV = True

End Function


Private Function fnExcelConvertOneFile(vFileSpec As Variant, bFirstSheetOnly As Boolean, sArchiveDir As String, Optional sErrorMsg As String) As Boolean
    Dim oXL As Excel.Application, oBook As Excel.Workbook, oSheet As Excel.Worksheet, vValue As Variant
    Dim oCell As Variant
    Dim Sheet As Excel.Worksheet
    Dim j As Long
    Dim i As Long
    Dim prevRow As Long
    Dim sTemp As String
    Dim sOutputLine As String
    Dim bCSVFileOpen As Boolean
    Dim sFileSpec As String
    Dim sBackupfile As String
    
    sFileSpec = CStr(vFileSpec)

      Set oXL = New Excel.Application
      oXL.DisplayAlerts = False

      Set oBook = oXL.Workbooks.Open(vFileSpec, , ReadOnly:=True)
        For Each oSheet In oBook.Sheets
         
            prevRow = 1
            sOutputLine = ""
            
            For Each oCell In oSheet.UsedRange
                sTemp = ""
                If Not IsNull(oCell.value) Then
                    If Not Trim(oCell.value) = "" Then
                        If Not Build_CSV_Char(oCell.value, sTemp) Then
                            Exit Function
                        End If
                    End If
                End If
                If oCell.row <> prevRow Then
                    sOutputLine = Trim_CSV(sOutputLine)
                    If sOutputLine <> "" Then
                        If Not bCSVFileOpen Then
                            If Open_CSV(vFileSpec & ".sheet_" & oSheet.Index & ".csv") Then
                                bCSVFileOpen = True
                            Else
                                Exit Function
                            End If
                        End If
                        Write_CSV sOutputLine
                    End If
                    sOutputLine = ""
                    prevRow = oCell.row
                End If
                sOutputLine = sOutputLine & sTemp & ","
            Next
            
            Set oCell = Nothing
            
            sOutputLine = Trim_CSV(sOutputLine)
            If sOutputLine <> "" Then
                If Not bCSVFileOpen Then
                    If Open_CSV(vFileSpec & ".sheet_" & oSheet.Index & ".csv") Then
                        bCSVFileOpen = True
                    Else
                        Exit Function
                    End If
                End If
                Write_CSV sOutputLine
            End If

            Set oSheet = Nothing
            
            If bCSVFileOpen Then
                bCSVFileOpen = False
                Close_CSV
            End If
            
            If bFirstSheetOnly Then GoTo exit_for_each
            
        Next
   
exit_for_each:

      oBook.Close
  
      Set oBook = Nothing
      oXL.Quit
      DoEvents
      
      Set oXL = Nothing
      
      'rename file to backup folder
      sBackupfile = BackupFileName(GetFileName(sFileSpec), GetFilePath(sFileSpec) & "\" & sArchiveDir, True)
      Name sFileSpec As sBackupfile

End Function


' Returns a collection with the names of all the files
' that match a file specification
'
' The file specification can include wildcards; multiple
' specifications can be provided, using a semicolon-delimited
' list, as in "*.tmp;*.bat"
' If RECURSEDIR is True the search is extended to all subdirectories
'
' It raises no error if path is invalid
'

Private Function GetAllFiles(ByVal path As String, ByVal filespec As String, _
    Optional RecurseDirs As Boolean) As Collection
    Dim spec As Variant
    Dim file As Variant
    Dim subdir As Variant
    Dim subdirs As New Collection
    Dim specs() As String
    
    ' initialize the result
    Set GetAllFiles = New Collection
    
    ' ensure that path has a trailing backslash
    If Right$(path, 1) <> "\" Then path = path & "\"
    
    ' get the list of provided file specifications
    specs() = Split(filespec, ";")
    
    ' this is necessary to ignore duplicates in result
    ' caused by overlapping file specifications
    On Error Resume Next
                
    ' at each iteration search for a different filespec
    For Each spec In specs
        ' start the search
        file = Dir$(path & spec)
        Do While Len(file)
            ' we've found a new file
            file = path & file
            GetAllFiles.Add file, file
            ' get ready for the next iteration
            file = Dir$
        Loop
    Next
    
    ' first, build the list of subdirectories to be searched
    If RecurseDirs Then
        ' get the collection of subdirectories
        ' start the search
        file = Dir$(path & "*.*", vbDirectory)
        Do While Len(file)
            ' we've found a new directory
            If file = "." Or file = ".." Then
                ' exclude the "." and ".." entries
            ElseIf (GetAttr(path & file) And vbDirectory) = 0 Then
                ' ignore regular files
            Else
                ' this is a directory, include the path in the collection
                file = path & file
                subdirs.Add file, file
            End If
            ' get next directory
            file = Dir$
        Loop
        
        ' parse each subdirectory
        For Each subdir In subdirs
            ' use GetAllFiles recursively
            For Each file In GetAllFiles(subdir, filespec, True)
                GetAllFiles.Add file, file
            Next
        Next
    End If
    
End Function

' Retrieve a file's path
'
' Note: trailing backslashes are never included in the result

Public Function GetFilePath(FileName As String) As String
    Dim i As Long
    For i = Len(FileName) To 1 Step -1
        Select Case Mid$(FileName, i, 1)
            Case ":"
                ' colons are always included in the result
                GetFilePath = Left$(FileName, i)
                Exit For
            Case "\"
                ' backslash aren't included in the result
                GetFilePath = Left$(FileName, i - 1)
                Exit For
        End Select
    Next
End Function

' Return the extension of a file name

Public Function GetFileExtension(ByVal FileName As String) As String
    Dim i As Long
    For i = Len(FileName) To 1 Step -1
        Select Case Mid$(FileName, i, 1)
            Case "."
                GetFileExtension = Mid$(FileName, i + 1)
                Exit For
            Case ":", "\"
                Exit For
        End Select
    Next
End Function

' Return the extension of a file name

Public Function GetFileName(ByVal FileName As String) As String
    Dim i As Long
    For i = Len(FileName) To 1 Step -1
        Select Case Mid$(FileName, i, 1)
            Case "\"
                GetFileName = Mid$(FileName, i + 1)
                Exit For
            Case ":"
                Exit For
        End Select
    Next
End Function


'========================================================================================
'   The CSV File Format
'
'   Each record is one line.
'
'   A record separator may consist of a line feed (ASCII/LF=0x0A), or a carriage
'   return and line feed pair (ASCII/CRLF=0x0D 0x0A).
'
'   Fields may contain embedded line-breaks so a record may span more than one line.
'
'   Leading and trailing space-characters adjacent to comma field separators are ignored.
'
'   Fields with embedded commas must be delimited with double-quote characters.
'
'   Fields that contain double quote characters must be surounded by double-quotes, and
'   the embedded double-quotes must each be represented by a pair of consecutive double
'   quotes.
'
'   A field that contains embedded line-breaks must be surounded by double-quotes
'
'   Fields with leading or trailing spaces must be delimited with double-quote characters.
'
'   Fields may always be delimited with double quotes.
'   The delimiters will always be discarded.
'
'   The first record in a CSV file may be a header record containing column (field) names
'   There is no mechanism for automatically discerning if the first record is a header row,
'   so in the general case, this will have to be provided by an outside process (such as
'   prompting the user). The header row is encoded just like any other CSV record in
'   accordance with the rules above.
'========================================================================================
  
Public Function Build_CSV_Char(p_valuein As String, ByRef p_valueout, Optional bQuote As Boolean = False) As Boolean
    On Error GoTo errTrap
    
    Const FunctionName As String = "Build_CSV_Char"
    Dim Quote          As String
    
    Build_CSV_Char = False

    ' set our field delimiter to 0 characters and trim the target field
    If bQuote Then Quote = """"
    p_valueout = Trim(p_valuein)
    
    ' If the field length is greater than 0 then we have some checking to do
    If Len(p_valueout) > 0 Then
    
        ' First look to see if there is a comma in the field
        If InStr(1, p_valueout, ",") > 0 Then
            Quote = """"
        End If
        
        ' Now look to see if we have any quotes in the file
        If InStr(1, p_valueout, """") > 0 Then
            Quote = """"
            p_valueout = Replace(p_valueout, """", """""")
        End If
        
        ' And finally check for the vbLF or vbCRLF
        If InStr(1, p_valueout, vbLf) > 0 _
        Or InStr(1, p_valueout, vbCrLf) > 0 Then
            Quote = """"
        End If
        
    End If
    
    p_valueout = Quote & p_valueout & Quote
    
    Build_CSV_Char = True
    Exit Function
        
errTrap:
    'ErrorTrapMessage FunctionName, Err
End Function

Public Function Open_CSV(OutputFile As String) As Boolean
    On Error GoTo errTrap
    
    Const FunctionName As String = "Open_CSV"
    
    Open_CSV = False
    
    Open OutputFile For Output As #1
    
    Open_CSV = True
    Exit Function
 
errTrap:
    'ErrorTrapMessage FunctionName, Err
End Function

Public Function Write_CSV(CSV_Data As String) As Boolean
    On Error GoTo errTrap
    
    Const FunctionName As String = "Write_CSV"
    
    Write_CSV = False
    
    Print #1, CSV_Data
    
    Write_CSV = True
    Exit Function
 
errTrap:
    'ErrorTrapMessage FunctionName, Err
End Function

Public Function Close_CSV() As Boolean
    On Error GoTo errTrap
    
    Const FunctionName As String = "Close_CSV"
    
    Close_CSV = False
    
    Close #1
    
    Close_CSV = True
    Exit Function
 
errTrap:
    'ErrorTrapMessage FunctionName, Err
End Function

' Remove Trailing commas
Public Function Trim_CSV(CSV_Data As String) As String
    Dim i As Long
    CSV_Data = Trim(CSV_Data)
    
    For i = Len(CSV_Data) To 1 Step -1
        Select Case Mid$(CSV_Data, i, 1)
            Case Is <> ","
                Trim_CSV = Left$(CSV_Data, i)
                Exit For
        End Select
    Next
End Function




' Return True if a directory exists
' (the directory name can also include a trailing backslash)

Private Function DirExists(DirName As String) As Boolean
    On Error GoTo ErrorHandler
    ' test the directory attribute
    DirExists = GetAttr(DirName) And vbDirectory
ErrorHandler:
    ' if an error occurs, this function returns False
End Function

' Create a nested directory
'
' This is similar to the MkDir command, but it creates any
' intermediate directory if necessary

Public Function MakeDirPath(DirName As String) As Boolean
    On Error GoTo ErrorHandler
    
    Dim i As Long, path As String
    
    Do
        i = InStr(i + 1, DirName & "\", "\")
        path = Left$(DirName, i - 1)
        ' don't try to create a root directory
        If Right$(path, 1) <> ":" And Dir$(path, vbDirectory) = "" Then
            ' make this subdirectory if it doesn't exist
            ' (exits if any error)
            MkDir path
        End If
    Loop Until i >= Len(DirName)
    MakeDirPath = True
    Exit Function
ErrorHandler:
    ' if an error occurs, this function returns False
End Function

