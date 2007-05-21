VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsComboControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***********************************************************'
'
' Copyright (c) 1996 FACTOR, A Division of W.R.Hess Company
'
' Module name   : COMBMAIN.CLS
' Date          : August 19, 1996
' Programmer(s) : Qinggang Ma
'
' This is the main class for the floating drop down.

'Useage of this dropdown class:
'
'   1. Files to be included in the project
'           i. COMBOSET.Cls
'          ii. DROPDOWN.CLS
'         iii. COMBMAIN.CLS (This file)
'
'   2. Initilization:
'       a. General:
'           i. Declare an object as clsComboControl (Example: objDropDown)
'          ii. Initialize the class (Example: Set objDropDown = New clsComboControl)
'         iii. Set the form property (Example:  Set objDropDown.Form = Me)
'          iv. Set the database property (Example: Set objDropDown.DataBase = t_dbMainDatabase)
'           v. Set the datalink property (Example: Set objDropDown.DataLink = data1)
'          vi. Set the table property (Example: Set objDropDown.Table = TDBGrid1)
'       b. Add a combo dropdown:
'           i. Call AddCombo function
'               Example:   objDropDown.AddCombo
'               or         n_hCombo = objDropDown.AddCombo
'               This function returns a handle to the combo. In case you need
'                   the set-up of the combo, you can use this handle.
'          ii. Set the SQL statement for the combo:
'               Use the SQL property for current combo:
'                   Example:  objDropDown.SQL = "SELECT ..."
'               or use the SetComboSQL statement:
'                   Example:  objDropDown.SetComboSQL "SELECT ...", ComboHandle (optional)
'       c. Add a Textbox and Command button comboset to a combo
'           Call AddComboBox sub:
'               Example: objDropDown.AddComboBox Text1, Command1, "prft_ctr", objDropDown.SQL_LONG_TYPE (optional)
'
'   3. Event calls:
'       a. In True Grid table, make the following calls in the events:
'
'                Events         Function Call to the class
'           i.   Click          objDropDown.Click TDBGrid1
'          ii.   GotFOcus       objDropDown.GotFocus TDBGrid1
'         iii.   Keypress       objDropDown.Keypress TDBGrid1, KeyAscii
'          iv.   LostFocus      objDropDown.LostFocus TDBGrid1
'           v.   MouseUp        objDropDown.TableMouseUp Y
'          vi.   RowColChange   objDropDown.RowColChange
'         vii.   SelChange      objDropDown.SelChange CANCEL
'
'       b. In each text box, make the following calls
'
'                Events         Function Call to the class
'           i.   Change         objDropDown.Change Text1
'          ii.   Click          objDropDown.Click Text1
'         iii.   GotFocus       objDropDown.GotFocus Text1
'          iv.   Keypress:
'                               If Not objDropDown.Keypress(Text1, KeyAscii) Then
'                                   KeyAscii = 0
'                               End If
'       c. In each command button, make the following calls
'
'                Events         Function Call to the class
'                Click          objDropDown.Click cmdButton1
'
'   4. Auxillary functions and properties
'       a. AddExtraColumn(sField As String, _
'                          nWidth As Integer, _
'                          Optional vCombo As Variant)
'          This funciton is useful when a column in the combo is needed, but
'            it is not populated with any textbox and command button.
'       b. CurrentCombo
'           Property: get or set the current combo handle.
'       c. SingleRecordSelected
'           Property: Return true if a single record is selected.
'       d. VisibleRows
'           Property: Set or get the visiblerows of the current combo.
'       e. SQL_DATE_TYPE, SQL_INT_TYPE, SQL_LONG_TYPE, and SQL_STRING_TYPE
'           Properties: Read-only. Return the constant to determine the type of
'               a textbox.
'          SQL_PHONE_TYPE is ADDED on 01/03/2001 WJ

'          ******************************************************************
'          *david 10/30/2001 for 3.26.xxxx release
'          *The SQL_STRING_TYPE SQL will be case INSENSITIVE (default)
'          *This means that the search (such as name, desc and etc.) will NOT BE case sensitive
'          *
'          *new property
'          *CaseSensitive() - set or return the value of the SQL_STRING_TYPE to be treated
'          *                  True - Query will be case sensitive,
'          *                  False - Query will NOT be case sensitive
'          ******************************************************************
'
'       f. DropDown, DropUp, and Position
'           Properties: DropDown and DropUp are read-only, Position is write-only
'           This three properties is used to set the position where the table appears.
'               Example:  objDropDown.Position = objDropDown.DropDown
'                           Table appears below the textboxes
'               Example:  objDropDown.Position = objDropDown.DropUp
'                           Table appears above the textboxes
'       g. LocalDatabase
'           Property: write-only set local database.
'               When a local access database is needed. This property is required.
'       h. SetColumnCaptions
'               Sub: Set the captions of the columns in the dropdown table
'               Example: SetColumnCaptions "ProfitCenter","Description"
'       i. SetOrderingAscent, SetOrderingDesscent
'               Sub: Sets to ascent order when the list is drop.
'                    The variable (objControl) can be any of the controls in the dropdown combo.
'       j. SetPositionAbove, SetPositionBelow
'               Sub: Set the position of the dropdown table
'               Example: objDropDown.SetPositionAbove txtVendor   (Control)
'                    or   objDropDown.SetPositionAbove nDDVendor  (Combo handle)
'       k. SetToAccess, SetToInformix
'               Sub: Set the database type of a combo. If no handle is passed,
'                    it set the currentcombo. You can pass in the combo handle
'                    to set the specific one.
'
'       l. RegExpPattern(objControl As Object, sPattern As String)
'           Property: write-only
'               This property is needed only if you want overwrite the default one.
'       m. GetComboSQL(Optional vCombo As Variant)
'           Function: Return the SQL of current combo. If vCombo is passed, return
'               the SQL of the specified combo.
'       n. ComboCount:
'           Property: read-only. Return the number of the combos in the control.

Option Explicit

Private Const dbgBlackLine = 1
Private Const dbgFixedSingle = 1
Private Const dbgNone = 0
Private Const dbgHorizontal = 1
Private Const dbgVertical = 2
Private Const dbgBoth = 3
Private Const dbgAutomatic = 4
Private Const dbgHighlightRow = 3
Private Const dbg3D = 1
Private Const dbgScrollOnSelect = 0
Private Const dbgControlNavigation = 0
Private Const dbgMarqueeRow = 2
Private Const dbgCurrentCell = 1
Private Const dbgUnbound = 1

Private Const TYPE_STRING As Integer = -1
Private Const TYPE_LONG As Integer = -2
Private Const TYPE_INT As Integer = -3
Private Const TYPE_DECIMAL As Integer = -4
Private Const TYPE_DATE As Integer = -5
Private Const TYPE_OBJECT As Integer = -6
Private Const TYPE_PHONE As Integer = -7

Private Const MAX_TYPES As Integer = -8

Private Const TYPE_DCM_INT_MASK = 19 ' a prime
Private Const DEFAULT_COLWIDTH As Integer = 1500

Private Const FACTORSEARCH = 1
Private Const COMMONDROPDOWN = 2

Private Const DB_INFORMIX As Integer = 0
Private Const DB_ACCESS As Integer = 1

Private Const nDROP_DOWN As Integer = 1
Private Const nDROP_UP As Integer = 2

Private Const nORDERING_DESCENT As Integer = -1
Private Const nORDERING_ASSCENT As Integer = 1
Private Const nORDERING_NONE As Integer = 2

Private Const CONTROL_TYPE_TEXTBOX As Integer = 1
Private Const CONTROL_TYPE_BUTTON As Integer = 2

Private Const nVISIBLEROWS_DEFAULT As Integer = 7
Private Const SCROLLBAR_WIDTH As Integer = 264

Private Const EXTRACOLUMNOFFSET = 1000
Private Const EXTR_COL_EXT_RIGHT = 1
Private Const EXTR_COL_EXT_LEFT = 2
Private Const KEY_CTRL_Y = &H19
Private Const sINITIAL As Integer = -1

Private Const SQL_INT_MAX As String = "32767"
Private Const SQL_LONG_MAX As String = "2147483647"

Private Const nHEIGHTOFFSET = 3
Private Const nERROR_DLLFILE = 53
Private Const nERROR_SQL = -100

Private Const sSTRING_PATTERN As String = "^P*$"
Private Const sLONG_PATTERN As String = szLongPattern
Private Const sINT_PATTERN As String = szIntegerPattern
Private Const sFieldSep As String * 1 = ","

Private sPercentOrAster(1) As String * 1
Private nCurrentCombo As Integer
Private colComboControls As Collection

Private tgDropDown As Object
Private frmParent As Form
Private datLink As Data
Private dbRemote As DataBase
Private t_dbLocal As DataBase
Private engDatabase As DBEngine
Private sGridData() As String
Private lRowCount As Long
Private m_nOringalFormWidth As Integer
Private m_nOringalFormHeight As Integer

Private clsCurrentDropdown As clsComboDropDown

Private bTableClicked As Boolean
Private nClickPointY As Single
Private bValidSelection As Boolean
Private lRecordsReturned As Long
Private lComboIndex As Integer
Private nErrorNumber As Integer
Private nLastRow As Integer
Private m_bDebug As Boolean
Private m_bFillingText As Boolean
Private m_bProtoType As Boolean

'david 10/30/2001
Private m_bCaseSensitive As Boolean

'david 01/28/2002
Private lHeadClickDropDown As Long
Private bHeaderClick As Boolean

'david 01/30/2002
Private m_bSortOnColumnClick As Boolean

#If Win32 Then
    Private Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
    Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
    Private Declare Function ScreenToClient Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
    Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
#Else
    Private Declare Function SetParent Lib "USER" (ByVal hWndChild As Integer, ByVal hWndNewParent As Integer) As Integer
    Private Declare Function GetParent Lib "USER" (ByVal hwnd As Integer) As Integer
    Private Declare Sub GetWindowRect Lib "USER" (ByVal hwnd As Integer, lpRect As RECT)
    Private Declare Function ScreenToClient Lib "USER" (ByVal hwnd As Integer, lpPoint As POINTAPI) As Boolean
#End If

'#04/22/2003 Vals added by
'#WJ Magic #407922. Automaticall Detect Data Type
'#And Try to Format and Align the Grid Column Display
Private Const DB_CHAR_TYPE = 0
Private Const DB_INT_TYPE = 1
Private Const DB_NUM_TYPE = 2
Private Const DB_DATE_TYPE = 3
'

Property Set Form(frmMain As Object)
    Set frmParent = frmMain
    m_nOringalFormWidth = frmMain.Width
    m_nOringalFormHeight = frmMain.Height
End Property

Property Set DataBase(dbSource As Object)
    'Set the database property of the dropdown control
    Set dbRemote = dbSource
End Property

Property Set DataLink(dLink As Object)
    'Set the datalink property of the dropdown control
    Set datLink = dLink
End Property

Property Set DBEngine(objTemp As Object)
    Set engDatabase = objTemp
End Property

Property Set LocalDatabase(dbSource As Object)
    Set t_dbLocal = dbSource
End Property

'#04/22/2003 Function added by
'#WJ Magic #407922. Automaticall Detect Data Type
'#And Try to Format and Align the Grid Column Display

Private Function fnFieldType(ByVal rsTemp As Recordset, _
                             ByVal sFieldName As String, _
                             ByVal nDB As Integer, _
                             ByRef right_nbr As Integer) As Integer
    Dim strSQL As String
    Dim rsCol As Recordset
    Dim dbField As field
    Dim bFound As Boolean
    Dim nType As Integer
    
    On Error GoTo errTrap
    
    fnFieldType = DB_CHAR_TYPE
    
    right_nbr = 0
    bFound = False
    For Each dbField In rsTemp.Fields
        If Trim(dbField.Name) = Trim(sFieldName) Then
             bFound = True
             Exit For
        End If
    Next
    If bFound = False Then
        Exit Function
    End If
    
    Select Case dbField.Type
        Case dbInteger, dbLong
            fnFieldType = DB_INT_TYPE
        Case dbCurrency
            right_nbr = 2
            fnFieldType = DB_NUM_TYPE
        Case dbDate
            If dbField.Size = 8 Then
                fnFieldType = DB_DATE_TYPE
            End If
        Case dbSingle, dbDouble, dbNumeric, dbDecimal, dbFloat
            fnFieldType = DB_NUM_TYPE
            If nDB = DB_INFORMIX Then
                strSQL = "SELECT * FROM syscolumns WHERE colName='" & dbField.Name & "'"
                If fnGetRecord(rsCol, strSQL, nDB) Then
                    If rsCol.RecordCount > 0 Then
                        fnInformixColTypeDesc rsCol!ColType, rsCol!colLength, right_nbr
                    End If
                End If
            End If
    End Select
    
    Exit Function
errTrap:
    On Error GoTo 0
End Function

'#04/22/2003 Function added by
'#WJ Magic #407922. Automaticall Detect Data Type
'#And Try to Format and Align the Grid Column Display
Private Function fnInformixColTypeDesc(ByVal nColType As Integer, _
                               ByVal nColLength As Integer, _
                               ByRef right_nbr As Integer) As String
     
     Dim left_nbr As Integer
     Dim coltypecal As Integer
     
     On Error GoTo errTrap
     
     right_nbr = 0
     
     If nColType >= 256 Then
        coltypecal = nColType - 256
     Else
        coltypecal = nColType
     End If
     Select Case coltypecal
        Case 0
            fnInformixColTypeDesc = "char(" & nColLength & ")"
        Case 1
            fnInformixColTypeDesc = "smallint"
        Case 2
            fnInformixColTypeDesc = "integer"
        Case 3
            fnInformixColTypeDesc = "float"
        Case 4
            fnInformixColTypeDesc = "smallfloat"
        Case 5
            left_nbr = nColLength \ 256
            right_nbr = nColLength - (left_nbr * 256)
            Select Case right_nbr
                Case 0
                    fnInformixColTypeDesc = "decimal(" & left_nbr & ",0)"
                Case 255
                    fnInformixColTypeDesc = "decimal(" & left_nbr & ")"
                Case Else
                    fnInformixColTypeDesc = "decimal(" & left_nbr & "," & right_nbr & ")"
            End Select
        Case 6
            fnInformixColTypeDesc = "serial(1)"
        Case 7
            fnInformixColTypeDesc = "date"
        Case 8
            left_nbr = nColLength \ 256
            right_nbr = nColLength - (left_nbr * 256)
            Select Case right_nbr
                Case 0
                    fnInformixColTypeDesc = "money(" & left_nbr & ",0)"
                Case 255
                    fnInformixColTypeDesc = "money(" & left_nbr & ")"
                Case Else
                    fnInformixColTypeDesc = "money(" & left_nbr & "," & right_nbr & ")"
            End Select
        Case 10
            fnInformixColTypeDesc = "datetime"
        Case 14
            fnInformixColTypeDesc = "interval"
     End Select
     If right_nbr >= 8 Then
        right_nbr = 0
     End If
     Exit Function
errTrap:
    On Error GoTo 0

End Function

Property Let AppendCriteria(objControl As Object, bTemp As Boolean)
    Dim lDropDown As Integer
    Dim nctrlType As Integer
    Dim clsTemp As clsComboSet

    fnGetKeys lDropDown, lComboIndex, nctrlType, objControl
    If lDropDown < 0 Then
        Exit Property
    End If
    
    Set clsTemp = colComboControls.Item(lDropDown).Combos.Item(lComboIndex)
    clsTemp.AppendCriteria = bTemp
End Property

Public Function CellValue(ByVal nCol As Integer, _
                          ByVal lRow As Long) As String

    On Error Resume Next
    
    Dim nColCount As Integer
        
    nColCount = colComboControls(nCurrentCombo).Count
    CellValue = sGridData(nCol, lRow)
End Function

Property Let ProtoType(bTemp As Boolean)
    m_bProtoType = bTemp
End Property

Private Sub subColWidth(nW As Integer, _
                        bAddTo As Boolean, _
                        colTemp As Collection, _
                        ByVal nCol As Integer)
    
    Dim txtBox As Object
    Dim cmdButton As Object
    Dim objTextBox As Object
    Dim nTop As Long
    
    Set txtBox = colTemp(nCol).Textbox
    Set cmdButton = colTemp(nCol).Button
    bAddTo = True
    If fnLeftMostCol(colTemp, nCol) Then
        nW = txtBox.Width + cmdButton.Left + cmdButton.Width - (txtBox.Left + txtBox.Width)
    Else
        Set objTextBox = colTemp.Item(nCol - 1).Button
        If objTextBox Is Nothing Then
            Set objTextBox = colTemp.Item(nCol - 1).Textbox
        End If
        If objTextBox Is Nothing Then
            nW = DEFAULT_COLWIDTH
        Else
            If cmdButton Is Nothing Then
                nW = txtBox.Left + txtBox.Width - (objTextBox.Left + objTextBox.Width)
            Else
                If colTemp.Item(nCol - 1).Textbox Is Nothing Then
                    nTop = objTextBox.Top
                Else
                    nTop = colTemp.Item(nCol - 1).Textbox.Top
                End If
                If Abs(txtBox.Top - nTop) < 100 Then
                    nW = cmdButton.Left + cmdButton.Width - (objTextBox.Left + objTextBox.Width)
                Else
                    nW = txtBox.Width
                    bAddTo = False
                End If
            End If
        End If
    End If
    If nW < 0 Then
        nW = DEFAULT_COLWIDTH
    End If
End Sub

Private Function fnLeftMostCol(colTemp As Collection, _
                               nCol As Integer) As Boolean

    Dim i As Integer
    
    If nCol > 1 Then
        For i = 1 To nCol - 1
            If Not colTemp(i).Button Is Nothing Then
                fnLeftMostCol = False
                Exit Function
            End If
        Next i
    End If
    fnLeftMostCol = True

End Function

Property Get KeyChanged(ctrl As Object) As Boolean
    Dim lDropDown As Integer
    Dim nctrlType As Integer
    Dim clsTemp As clsComboSet

    fnGetKeys lDropDown, lComboIndex, nctrlType, ctrl
    
    KeyChanged = False
    If lDropDown > 0 Then
    For Each clsTemp In colComboControls.Item(lDropDown).Combos
        If clsTemp.KeyField Then
            If clsTemp.Changed Then
                KeyChanged = True
                Exit For
            End If
        End If
    Next
    End If
End Property

Property Let KeyField(objControl As Object, bTemp As Boolean)
    Dim lDropDown As Integer
    Dim nctrlType As Integer
    Dim clsTemp As clsComboSet

    fnGetKeys lDropDown, lComboIndex, nctrlType, objControl
    If lDropDown < 0 Then
        Exit Property
    End If
    
    Set clsTemp = colComboControls.Item(lDropDown).Combos.Item(lComboIndex)
    clsTemp.KeyField = bTemp
End Property

Private Function fnGetTop() As Integer
    Dim clsTemp As clsComboSet
    Dim nHeight As Integer
    
    If clsCurrentDropdown.Position = nDROP_UP Then
        For Each clsTemp In clsCurrentDropdown.Combos
            If Not clsTemp.Textbox Is Nothing Then
                nHeight = clsTemp.Textbox.Height
                Exit For
            End If
        Next clsTemp
        fnGetTop = clsCurrentDropdown.TableTop - nHeight - tgDropDown.Height
    Else
        fnGetTop = clsCurrentDropdown.TableTop
    End If

End Function

Private Function fnComboSet(ByVal nDropDown As Integer, _
                            ByVal nCombo As Integer, _
                            Optional vField As Variant) As clsComboSet

    Set fnComboSet = Nothing
    Dim ciTemp As clsComboSet
    If IsMissing(vField) Then
        For Each ciTemp In colComboControls.Item(nDropDown).Combos
            If Not ciTemp.KeyField Then
                Set fnComboSet = ciTemp
                Exit For
            End If
        Next ciTemp
    Else
        Dim clsTemp As clsComboSet
        Dim sKey As String
        sKey = UCase(Trim(vField))
        For Each clsTemp In colComboControls.Item(nDropDown).Combos
            If sKey = UCase(clsTemp.AliasName) Then
                Set fnComboSet = clsTemp
                Exit For
            End If
        Next clsTemp
    End If
End Function

Private Function fnSQLValidation(ByVal lDropDown As Integer, _
                                 ByVal bTakeEmpty As Boolean) As String
    Const szUNION As String = " UNION "

    Dim szQuery As String
    Dim szCode As String
    Dim sColName As String
    Dim nPos As Integer
    Dim lCol As Integer
    Dim nColCount As Integer
    Dim nTop As Integer
    Dim nLeft As Integer
    Dim colCombo As Collection
    
    'Get the dropdown class
    Set clsCurrentDropdown = colComboControls.Item(lDropDown)
    Set colCombo = clsCurrentDropdown.Combos

    szQuery = clsCurrentDropdown.SQL

    nPos = InStr(UCase(szQuery), szUNION)
    If nPos = 0 Then
        szQuery = fnAppendCriteria1(szQuery, clsCurrentDropdown.dbType, 0, bTakeEmpty, False)
        'david 09/21/2004
        sColName = fnGetFieldName(colCombo.Item(lComboIndex).DataField)
        '''''''''''''''''
    Else
        Dim szQuery1 As String
        Dim szQuery2 As String
        Dim sRPart As String
        Dim nLen As Integer
        
        sRPart = szQuery
        szQuery = ""
        lCol = 1
        While sRPart <> ""
            If nPos > 1 Then
                szQuery1 = Left(sRPart, nPos - 1)
                nLen = Len(sRPart) - nPos - Len(szUNION) + 1
                If nLen > 0 Then
                    sRPart = Right(sRPart, nLen)
                Else
                    sRPart = ""
                End If
            Else
                szQuery1 = sRPart
                sRPart = ""
            End If

            szQuery1 = fnAppendCriteria1(szQuery1, clsCurrentDropdown.dbType, lCol, bTakeEmpty, False)
            If szQuery = "" Then
                szQuery = szQuery1
            Else
                szQuery = szQuery & szUNION & szQuery1
            End If
            nPos = InStr(sRPart, szUNION)
            If nPos = 0 Then
                nPos = Len(sRPart) + 1
            End If
            lCol = lCol + 1
        Wend
        sColName = CStr(lComboIndex)
    End If
    'Set clsCurrentDropdown = Nothing
    fnSQLValidation = szQuery

End Function

Property Let Populate(objControl As Object, bTemp As Boolean)
    Dim lDropDown As Integer
    Dim nctrlType As Integer
    Dim clsTemp As clsComboSet

    fnGetKeys lDropDown, lComboIndex, nctrlType, objControl
    If lDropDown < 0 Then
        Exit Property
    End If
    
    Set clsTemp = colComboControls.Item(lDropDown).Combos.Item(lComboIndex)
    clsTemp.Populate = bTemp
End Property

Property Let FieldType(objControl As Object, nTemp As Integer)
    Dim lDropDown As Integer
    Dim nctrlType As Integer
    Dim clsTemp As clsComboSet

    fnGetKeys lDropDown, lComboIndex, nctrlType, objControl
    If lDropDown < 0 Then
        Exit Property
    End If
    
    Set clsTemp = colComboControls.Item(lDropDown).Combos.Item(lComboIndex)
    subSetType objControl, clsTemp, nTemp
End Property

Public Sub Resize()
    Dim clsDD As clsComboDropDown
    Dim clsCombo As New clsComboSet
    Dim i As Integer
    Dim bLeftSet As Boolean
    Dim nLeft As Integer
    Dim nTop As Integer
    Dim nTW As Integer
    Dim nWCol As Integer
    Dim bAdd As Boolean
    
    tgDropDown.Visible = False
    For Each clsDD In colComboControls
        nTW = 0
        bLeftSet = False
        For i = 1 To clsDD.Combos.Count
            Set clsCombo = clsDD.Combos(i)
            If Not bLeftSet Then
                If Not (clsCombo.Textbox Is Nothing Or clsCombo.Button Is Nothing) Then
                    bLeftSet = True
                    fnGetOffsets nLeft, nTop, clsCombo.Textbox
                    clsDD.TableLeft = nLeft - nTW
                    clsDD.TableTop = nTop
                End If
            End If
            If clsCombo.AutoWidth Then
                subColWidth nWCol, bAdd, clsDD.Combos, i
                clsCombo.ColWidth = nWCol
            End If
            If bAdd Then
                nTW = nTW + clsCombo.ColWidth
            End If
        Next i
        clsDD.TableWidth = nTW
    Next clsDD
End Sub

Property Let SearchOnReturn(objControl As Object, bTemp As Boolean)
    Dim lDropDown As Integer
    Dim nctrlType As Integer
    Dim clsTemp As clsComboDropDown

    fnGetKeys lDropDown, lComboIndex, nctrlType, objControl
    If lDropDown < 0 Then
        Exit Property
    End If
    
    Set clsTemp = colComboControls.Item(lDropDown)
    clsTemp.SearchOnReturn = bTemp
End Property

Property Get SearchOnReturn(objControl As Object) As Boolean
    Dim lDropDown As Integer
    Dim nctrlType As Integer
    Dim clsTemp As clsComboDropDown

    fnGetKeys lDropDown, lComboIndex, nctrlType, objControl
    If lDropDown < 0 Then
        Exit Property
    End If
    
    Set clsTemp = colComboControls.Item(lDropDown)
    SearchOnReturn = clsTemp.SearchOnReturn
End Property

Property Get RecordsReturned() As Long
    RecordsReturned = lRecordsReturned
    lRecordsReturned = -1
End Property

Public Sub ResetFlags()

    Dim clsTemp1 As clsComboDropDown
    Dim clsTemp2 As clsComboSet
    
    For Each clsTemp1 In colComboControls
        For Each clsTemp2 In clsTemp1.Combos
            clsTemp2.Changed = True
            clsTemp2.FieldValue = ""
        Next clsTemp2
        clsTemp1.Cleared = True
    Next clsTemp1
End Sub

Public Function SecondaryBox(objCtrl As Object, sField As String) As Object
    Dim lDropDown As Integer
    Dim nctrlType As Integer
    Dim clsTemp As clsComboSet
    Dim sKey As String
    
    fnGetKeys lDropDown, lComboIndex, nctrlType, objCtrl
    If lDropDown < 0 Or lComboIndex < 0 Then
        Exit Function
    End If
    Set SecondaryBox = Nothing
    sKey = UCase(Trim(sField))
    For Each clsTemp In colComboControls.Item(lDropDown).Combos
        If sKey = UCase(clsTemp.AliasName) Then
            Set SecondaryBox = clsTemp.Textbox
        End If
    Next clsTemp

End Function

Property Let ComboLeft(txtBox As Object, _
                       ByVal nLeft As Integer)
    Dim nTop As Integer
    Dim lDropDown As Integer
    Dim nctrlType As Integer
    Dim clsTemp As clsComboDropDown

    fnGetKeys lDropDown, lComboIndex, nctrlType, txtBox
    If lDropDown < 0 Then
        Exit Property
    End If
    Set clsTemp = colComboControls.Item(lDropDown)
    
    fnGetOffsets nLeft, nTop, txtBox
    clsTemp.TableLeft = nLeft
    clsTemp.ExtendReset = False
    
End Property

Property Get ComboLeft(txtBox As Object) As Integer
    Dim nTop As Integer
    Dim nLeft As Integer
    Dim lDropDown As Integer
    Dim nctrlType As Integer
    Dim clsTemp As clsComboDropDown

    fnGetKeys lDropDown, lComboIndex, nctrlType, txtBox
    If lDropDown < 0 Then
        Exit Property
    End If
    Set clsTemp = colComboControls.Item(lDropDown)
    fnGetOffsets nLeft, nTop, txtBox
    ComboLeft = clsTemp.TableLeft - nLeft
    'Always return 0, because it is zero relative to the left of the textbox.
End Property

Property Get ComboWidth(txtBox As Object) As Integer
    Dim lDropDown As Integer
    Dim nctrlType As Integer
    Dim clsTemp As clsComboDropDown

    fnGetKeys lDropDown, lComboIndex, nctrlType, txtBox
    If lDropDown < 0 Then
        Exit Property
    End If
    ComboWidth = colComboControls.Item(lDropDown).TableWidth + SCROLLBAR_WIDTH
                        
End Property

Public Sub SetExtend(sField As String, _
                     ByVal nDir As Integer)
    Dim lDropDown As Integer
    Dim clsTemp As clsComboDropDown
    Dim clsCombo As clsComboSet
    Dim colTemp As Collection

    subIndexFromField lDropDown, lComboIndex, sField
    If lDropDown < 0 Or lComboIndex < 0 Then
        Exit Sub
    End If
    Set clsTemp = colComboControls.Item(lDropDown)
    Set colTemp = clsTemp.Combos
    Set clsCombo = colTemp(lComboIndex)
    
    If Not ((lComboIndex = 1 And nDir = EXTR_COL_EXT_LEFT) Or (lComboIndex = colTemp.Count And nDir = EXTR_COL_EXT_RIGHT)) Then
        colTemp.Remove lComboIndex
        If nDir = EXTR_COL_EXT_LEFT Then
            colTemp.Add clsCombo, , 1
        Else
            colTemp.Add clsCombo
        End If
    End If
    If nDir = EXTR_COL_EXT_LEFT Then
        clsTemp.TableLeft = clsTemp.TableLeft - clsCombo.ColWidth
        clsTemp.ExtendReset = True
    Else
        If clsTemp.ExtendReset Then
            clsTemp.TableLeft = clsTemp.TableLeft + clsCombo.ColWidth
            clsTemp.ExtendReset = False
        End If
    End If

End Sub

Public Sub SetText(objCtrl As Object, _
                   vText As Variant, _
                   Optional vField As Variant)

    Dim lDropDown As Integer
    Dim nctrlType As Integer
    Dim clsTemp As clsComboSet
    
    fnGetKeys lDropDown, lComboIndex, nctrlType, objCtrl
    If lDropDown < 0 Or lComboIndex < 0 Then
        Exit Sub
    End If
    Set clsTemp = fnComboSet(lDropDown, lComboIndex, vField)
    If Not clsTemp Is Nothing Then
        With clsTemp
            If IsNull(vText) Then
                .FieldValue = ""
            Else
                .FieldValue = Trim(vText)
            End If
            .Valid = True
            .Changed = False
        End With
        For Each clsTemp In colComboControls.Item(lDropDown).Combos
            If clsTemp.KeyField Then
                clsTemp.Changed = False
            End If
        Next
        With colComboControls.Item(lDropDown)
            .Cleared = False
        End With
    End If
End Sub

Public Function GetText(objCtrl As Object) As String

    Dim lDropDown As Integer
    Dim nctrlType As Integer
    Dim clsTemp As clsComboSet
    
    fnGetKeys lDropDown, lComboIndex, nctrlType, objCtrl
    If lDropDown < 0 Or lComboIndex < 0 Then
        Exit Function
    End If
    Set clsTemp = colComboControls.Item(lDropDown).Combos.Item(lComboIndex)
    If Not clsTemp Is Nothing Then
        With clsTemp
            If .Valid Then
                GetText = .FieldValue
            End If
        End With
    End If
End Function

Public Sub FillText(objCtrl As Object)

    Dim lDropDown As Integer
    Dim nctrlType As Integer
    Dim clsTemp As clsComboSet
    Dim clsCycle As clsComboSet
    
    fnGetKeys lDropDown, lComboIndex, nctrlType, objCtrl
    If lDropDown < 0 Or lComboIndex < 0 Then
        Exit Sub
    End If
    Set clsTemp = colComboControls.Item(lDropDown).Combos.Item(lComboIndex)
    If Not clsTemp Is Nothing Then
        With clsTemp
            If .Valid Then
                If .KeyField Then
                    m_bFillingText = True
                    For Each clsCycle In colComboControls.Item(lDropDown).Combos
                        If Not clsCycle Is clsTemp Then
                            clsCycle.Textbox = clsCycle.FieldValue
                        End If
                    Next
                    DoEvents
                    m_bFillingText = False
                Else
                    m_bFillingText = True
                    objCtrl = .FieldValue
                    DoEvents
                    m_bFillingText = False
                End If
            End If
        End With
    End If
End Sub

Private Sub subSetType(objControl As Object, _
                       clsTemp As clsComboSet, _
                       ByVal nType As Integer)
    With clsTemp
        .ComboType = nType
        .RegExpPattern = fnGetPCode(nType)
        If nType < MAX_TYPES Then
            If TypeOf objControl Is Textbox Then
                objControl.MaxLength = -nType \ TYPE_DCM_INT_MASK
            End If
        End If
    End With
End Sub

Property Let TableWidth(objControl As Object, _
                        ByVal nTemp As Long)

    Dim lDropDown As Integer
    Dim nctrlType As Integer
    Dim clsTemp As clsComboDropDown
    Dim clsCombo As clsComboSet
    Dim nWidth As Long

    fnGetKeys lDropDown, lComboIndex, nctrlType, objControl
    If lDropDown < 0 Then
        Exit Property
    End If
    
    Set clsTemp = colComboControls.Item(lDropDown)
    nWidth = 0
    For Each clsCombo In clsTemp.Combos
        nWidth = nWidth + clsCombo.ColWidth
    Next
    clsTemp.TableWidth = nTemp
    If nTemp <= nWidth Then
        clsTemp.HorScroll = True
    Else
        clsTemp.HorScroll = False
    End If
    
End Property

Public Property Let TextSet(objCtrl As Object, bFlag As Boolean)
    'If bFlag is true. The text in the boxes are flaged as SET, so that
    ' all the text boxes will be cleared when a key is pressed in
    ' the key field (The first text box added to the dropdown)
    
    Dim i As Integer
    Dim nDropDown As Integer
    Dim nCombo As Integer
    
    If objCtrl Is Nothing Then
        For i = 1 To colComboControls.Count
            colComboControls.Item(i).Cleared = Not bFlag
        Next
    Else
        fnGetKeys nDropDown, nCombo, i, objCtrl
        If nDropDown > 0 Then
            colComboControls.Item(nDropDown).Cleared = Not bFlag
        End If
    End If

End Property

Public Property Get TextSet(objCtrl As Object) As Boolean
    'If bFlag is true. The text boxes will be cleared when a key
    ' is pressed in the key field (The first text box added to the dropdown)
    
    Dim i As Integer
    Dim nDropDown As Integer
    Dim nCombo As Integer
    
    If Not objCtrl Is Nothing Then
        fnGetKeys nDropDown, nCombo, i, objCtrl
        If nDropDown > 0 Then
            TextSet = Not colComboControls.Item(nDropDown).Cleared
        End If
    End If

End Property

Property Let ShowNoDataMessage(objControl As Object, bStatus As Boolean)
    'Set the SQL of the specific combo

    Dim lDropDown As Integer
    Dim nctrlType As Integer
    Dim lCombo As Integer
    
    'Get the index of the current combo
    fnGetKeys lDropDown, lCombo, nctrlType, objControl
    If lDropDown < 0 Then
        Exit Property
    End If
    
    colComboControls.Item(lDropDown).ShowNoDataMessage = bStatus

End Property

Public Sub SetDebug(bStatus As Boolean)
    m_bDebug = bStatus
End Sub

Public Sub AddExtraField(txtBox As Object, _
                         sField As String, _
                         Optional vType As Variant, _
                         Optional vCombo As Variant)
                         
    'Add an extra field.
    ' Contrast to AddExtraColumn, this field is not shown in the dropdown table,
    ' but it will be populated with the given textbox.

    Dim nCombo As Integer
    Dim clsCombo As clsComboSet
    Dim colTemp As Collection
    Dim i As Integer
    Dim sTemp As String
    
    If IsMissing(vCombo) Then
        nCombo = nCurrentCombo
    Else
        nCombo = vCombo
    End If
    
    If nCombo <= 0 Then
        Exit Sub
    End If

    Set colTemp = colComboControls.Item(nCombo).Combos
    sTemp = UCase(Trim(sField))
    For i = 1 To colTemp.Count
        With colTemp.Item(i)
            If .Button Is Nothing Then
                If sTemp = UCase(Trim(.AliasName)) Then
                    Set clsCombo = colTemp.Item(i)
                    Exit For
                End If
            End If
        End With
    Next
    If clsCombo Is Nothing Then
        Set clsCombo = New clsComboSet
        'Add to the dropdown control
        colTemp.Add clsCombo
        clsCombo.ColWidth = -1
    End If
    With clsCombo
        Set .Textbox = txtBox
        Set .Button = Nothing
        .AliasName = Trim(sField)
        .FieldValue = ""
        .AppendCriteria = False
        .Populate = True
        .Changed = True
        .Valid = False
        .Visible = False
    End With

    fnSetupAlias nCombo      'For population of the text

    'Setup the regular expression
    If IsMissing(vType) Then
        clsCombo.ComboType = TYPE_STRING
        clsCombo.RegExpPattern = fnGetPCode(TYPE_STRING)
    Else
'        If vType > 0 Then
'            Dim nLen As Integer
'            clsCombo.RegExpPattern = fnRegExpPCode("^P{0," & CStr(vType) & "}$")
'        Else
            clsCombo.ComboType = vType
            clsCombo.RegExpPattern = fnGetPCode(vType)
            If vType < MAX_TYPES Then
                txtBox.MaxLength = -vType \ TYPE_DCM_INT_MASK + 1
            End If
'        End If
    End If
    
End Sub

Property Let ComboSQL(objControl As Object, sSql As String)
    'Set the SQL of the specific combo

    Dim lDropDown As Integer
    Dim nctrlType As Integer
    Dim lCombo As Integer
    
    'Get the index of the current combo
    fnGetKeys lDropDown, lCombo, nctrlType, objControl
    If lDropDown < 0 Then
        Exit Property
    End If
    
    SetComboSQL sSql, lDropDown

End Property

Property Let ComboErrorMessage(objControl As Object, sMsg As String)
    
    'Set the SQL of the specific combo

    Dim nIndex As Integer
    
    'Get the index of the current combo
    nIndex = fnGetComboIndex(objControl)
    
    If nIndex > 0 Then
        colComboControls.Item(nIndex).ErrorMessage = sMsg
    End If

End Property

Property Get ComboSQL(objControl As Object) As String
    'Returns the SQL of the specific combo

    Dim lDropDown As Integer
    Dim nctrlType As Integer
    Dim lCombo As Integer
    
    'Get the index of the combo
    fnGetKeys lDropDown, lCombo, nctrlType, objControl

    If lDropDown < 0 Then
        Exit Property
    End If
    
    ComboSQL = GetComboSQL(lDropDown)

End Property

Property Let ErrorMessage(sMsg As String)
    If colComboControls Is Nothing Then
        Exit Property
    End If
    If colComboControls.Count = 0 Then
        Exit Property
    End If
    
    colComboControls.Item(nCurrentCombo).ErrorMessage = sMsg

End Property

Property Get ExtLeft() As Integer
    ExtLeft = EXTR_COL_EXT_LEFT
End Property

Property Get ExtRight() As Integer
    ExtRight = EXTR_COL_EXT_RIGHT
End Property

Private Function fnAppendCriteria(ByVal szQuery As String, _
                                  ByVal nDBType As Integer, _
                                  nCol As Integer) As String
    Dim nPos As Integer
    Dim lCol As Integer
    Dim szLink As String
    Dim szCode As String
    Dim sTemp As String
    Dim sDataField As String
    Dim sGroupByClause As String
    Dim bAppendCriteria As Boolean
    Dim colCombo As Collection
    
    sTemp = UCase$(szQuery)
    nPos = InStr(sTemp, " WHERE ")
    If nPos = 0 Then
        szLink = " WHERE ("
    Else
        szLink = " AND ("
    End If
    nPos = InStr(sTemp, " GROUP BY ")
    If nPos > 1 Then
        sGroupByClause = Right(szQuery, Len(sTemp) - nPos + 1)
        szQuery = Left(szQuery, nPos - 1)
    Else
        sGroupByClause = ""
    End If
    Set colCombo = clsCurrentDropdown.Combos
    For lCol = 1 To colCombo.Count
        If Not (colCombo.Item(lCol).Textbox Is Nothing) Then
            bAppendCriteria = False
            If colCombo.Item(lCol).AppendCriteria Then
                If TypeOf colCombo.Item(lCol).Textbox Is Textbox Then
                    If Not colCombo.Item(lCol).Textbox.Locked Then
                        bAppendCriteria = True
                    End If
                End If
            End If
            If bAppendCriteria Then
                szCode = Trim$(colCombo.Item(lCol).Textbox.Text)
                ' if a criteria has been entered, add to the SQL
'                If szCode <> "" And Not (colCombo.Item(lCol).Button Is Nothing) Then    'Old
                If szCode <> "" Then
                    'david 09/21/2004
                    sDataField = fnGetFieldName(fnUnscrambleDataField(colCombo.Item(lCol).DataField, nCol))
                    '''''''''''''''''
                    Select Case colCombo.Item(lCol).ComboType
                        Case TYPE_DATE
                            If tfnRegExpControlChange(colCombo.Item(lCol).Textbox, colCombo.Item(lCol).RegExpPattern) Then
                                sTemp = tfnDateString(tfnFormatDate(szCode))
                                If IsDate(sTemp) Then
                                    szQuery = szQuery & szLink & sDataField & " >= '" & sTemp & "')"
                                    szLink = " AND ("
                                End If
                            End If
                        Case Is < MAX_TYPES
                            'Decimal type
                            If szCode = "-" Then
                                szQuery = szQuery & szLink & sDataField & " < 0)"
                            ElseIf CDbl(szCode) < 0 Then
                                szQuery = szQuery & szLink & sDataField & " <= '" & szCode & "')"
                            Else
                                szQuery = szQuery & szLink & sDataField & " >= '" & szCode & "')"
                            End If
                            szLink = " AND ("
                        Case TYPE_LONG
                            szQuery = szQuery & szLink & fnGenetateSearchString(sDataField, szCode, 10) & ")"
                            szLink = " AND ("
                        Case TYPE_INT
                            szQuery = szQuery & szLink & fnGenetateSearchString(sDataField, szCode, 5) & ")"
                            szLink = " AND ("
                        Case TYPE_PHONE
                            szQuery = szQuery & szLink & sDataField & " LIKE " & fnSQLString(fnPhoneCriteria(szCode, nDBType) & sPercentOrAster(nDBType)) & ")"
                            szLink = " AND ("
                        Case TYPE_STRING, Is > 0
                            
                            If nDBType = DB_ACCESS Then
                                szQuery = szQuery & szLink & sDataField & " LIKE " & fnSQLString(szCode & sPercentOrAster(nDBType)) & ")"
                            Else
                                
                                If m_bCaseSensitive Then
                                    szQuery = szQuery & szLink & sDataField & " LIKE " & fnSQLString(szCode & sPercentOrAster(nDBType)) & ")"
                                Else
                                    szQuery = szQuery & szLink & sDataField & " MATCHES " & fnSQLString(CaseInSensitiveString(szCode)) & ")"
                                End If
                                
                            End If
                            
                            szLink = " AND ("
                    End Select
                End If
            End If   'If bAppendCriteria Then
        End If  'If Not (colCombo.Item(lCol).Textbox Is Nothing) Then
    Next
    fnAppendCriteria = szQuery & sGroupByClause
End Function

Private Function fnPhoneCriteria(ByVal S As String, _
                        ByVal nDBType As Integer) As String
     Dim sTemp As String
     Dim i As Integer
     Dim sChar As String
     Dim sWild As String
     
     fnPhoneCriteria = ""
     
     S = Trim(S)
     sWild = sPercentOrAster(nDBType)
     
     If S <> "" Then
        sTemp = ""
        
        For i = 1 To Len(S)
           sChar = Mid(S, i, 1)
           If IsNumeric(sChar) Then
                sTemp = sTemp & sChar
           ElseIf sChar = "(" Or sChar = ")" Or sChar = "-" Or sChar = " " Then
                sTemp = sTemp & sWild
           ElseIf (sChar = "?" Or sChar = "%" Or sChar = "*") Then
                sTemp = sTemp & sWild
           End If
        Next
        If Len(sTemp) = 10 Then
            fnPhoneCriteria = sWild & Mid(sTemp, 1, 3) & sWild & sWild & Mid(sTemp, 4, 3) & "-" & Mid(sTemp, 7, 4)
        ElseIf Len(sTemp) >= 7 Then ' use the first 7
            fnPhoneCriteria = sWild & sWild & sWild & Mid(sTemp, 1, 3) & sWild & Mid(sTemp, 4, 4)
        ElseIf Len(sTemp) >= 3 Then
            fnPhoneCriteria = sWild & Mid(sTemp, 1, 3) & sWild & Mid(sTemp, 4, 3)
        Else
            fnPhoneCriteria = S
        End If
     End If
End Function

'david 01/30/2002
'take care of '%', '_' and '% %'
'this function will rebuild the criteria string so that it can be used for case insensitive WHERE clause
'NOTE: need to use the MATCHES keyword in the WHERE clause
Public Function CaseInSensitiveString(ByVal S As String) As String
    Dim i As Integer
    Dim sRet As String
    Dim sChar As String
    Dim bStartInserted As Boolean
    
    S = Trim(S)
    
    If Trim(S) = "" Then
        sRet = "*"
    Else
        For i = 1 To Len(S)
            sChar = Mid(S, i, 1)
            Select Case sChar
            Case " "
                sRet = sRet + " "
                bStartInserted = False
            Case "\", "?", "*"
                sRet = sRet + "\" + sChar
                bStartInserted = False
            Case Else
                If IsAphabet(sChar) Then
                    sRet = sRet + "[" + UCase(sChar) + LCase(sChar) + "]"
                ElseIf sChar = "_" Then
                    sRet = sRet + "?"
                ElseIf sChar = "%" Then
                    sRet = sRet + "*"
                    bStartInserted = True
                Else
                    sRet = sRet + sChar
                End If
                
                bStartInserted = False
            End Select
        Next i
            
        sRet = sRet + "*"
    End If
    
    CaseInSensitiveString = sRet
End Function

Private Function IsAphabet(ByVal sChar As String) As Boolean
    sChar = UCase(sChar)
    
    If sChar >= "A" And sChar <= "Z" Then
        IsAphabet = True
    End If
End Function

Private Function fnAppendCriteria1(ByVal szQuery As String, _
                                   ByVal nDBType As Integer, _
                                   nCol As Integer, _
                                   bTakeEmpty As Boolean, _
                                   bAllFields As Boolean) As String
    Dim nPos As Integer
    Dim lCol As Integer
    Dim szLink As String
    Dim szCode As String
    Dim sTemp As String
    Dim sDataField As String
    Dim sGroupByClause As String
    Dim bAppendCriteria As Boolean
    Dim colCombo As Collection
    
    sTemp = UCase$(szQuery)
    nPos = InStr(sTemp, " WHERE ")
    If nPos = 0 Then
        szLink = " WHERE ("
    Else
        szLink = " AND ("
    End If
    nPos = InStr(sTemp, " GROUP BY ")
    If nPos > 1 Then
        sGroupByClause = Right(szQuery, Len(sTemp) - nPos + 1)
        szQuery = Left(szQuery, nPos - 1)
    Else
        sGroupByClause = ""
    End If
    Set colCombo = clsCurrentDropdown.Combos
    For lCol = 1 To colCombo.Count
        If Not (colCombo.Item(lCol).Textbox Is Nothing) Then
            bAppendCriteria = False
            If colCombo.Item(lCol).AppendCriteria Then
                If bAllFields Or colCombo.Item(lCol).KeyField Then
                    If TypeOf colCombo.Item(lCol).Textbox Is Textbox Then
                        If Not colCombo.Item(lCol).Textbox.Locked Then
                            bAppendCriteria = True
                        End If
                    End If
                End If
            End If
            If bAppendCriteria Then
                szCode = Trim$(colCombo.Item(lCol).Textbox.Text)
                ' if a criteria has been entered, add to the SQL
                'david 09/21/2004
                sDataField = fnGetFieldName(fnUnscrambleDataField(colCombo.Item(lCol).DataField, nCol))
                '''''''''''''''''
                If Not bTakeEmpty Or szCode <> "" Then
                    Select Case colCombo.Item(lCol).ComboType
                        Case TYPE_DATE
                            If tfnRegExpControlChange(colCombo.Item(lCol).Textbox, colCombo.Item(lCol).RegExpPattern) Then
                                szQuery = szQuery & szLink & sDataField & " = '" & tfnDateString(tfnFormatDate(szCode)) & "')"
                                szLink = " AND ("
                            End If
                        Case Is < MAX_TYPES
                            szQuery = szQuery & szLink & sDataField & " = '" & szCode & "')"
                            szLink = " AND ("
                        Case TYPE_LONG, TYPE_INT
                            szQuery = szQuery & szLink & sDataField & " = " & szCode & ")"
                            szLink = " AND ("
                        Case TYPE_PHONE
                            szQuery = szQuery & szLink & sDataField & " = " & fnSQLString(szCode) & ")"
                            szLink = " AND ("
                        Case TYPE_STRING, Is > 0
                            szQuery = szQuery & szLink & sDataField & " = " & fnSQLString(szCode) & ")"
                            szLink = " AND ("
                    End Select
                End If   'If Not bTakeEmpty Or szCode <> "" Then
            End If   'If bAppendCriteria Then
        End If  'If Not (colCombo.Item(lCol).Textbox Is Nothing) Then
    Next
    fnAppendCriteria1 = szQuery & sGroupByClause
End Function


Private Sub fnExtractFields(sSql As String, colTemp As Collection)
    'Extract fields from the SQL statement (Not used now)
    Dim sTemp As String
    Dim sLPart As String
    Dim sRPart As String
    Dim nPos As Integer
    Dim nLen As Integer
    Dim sSQLFromClause As String
    Dim i As Integer
    
    sTemp = UCase(sSql)
    nPos = InStr(sTemp, " FROM")
    If nPos = 0 Then
        Exit Sub
    End If

    nLen = Len(sTemp) - nPos
    If nLen < 0 Then
        Exit Sub
    End If
    
    sSQLFromClause = Right(sTemp, nLen)
    sLPart = Trim(Left(sTemp, nPos))
    
    With colTemp
        For i = 1 To .Count
            If .Item(i).ColWidth < 0 Then
                sTemp = UCase(Trim(.Item(i).AliasName))
                nPos = InStr(sLPart, sTemp)
                If nPos > 1 Then
                    nLen = Len(sLPart) - nPos - Len(sTemp)
                    If nLen > 0 Then
                        sTemp = Right(sLPart, nLen)
                    Else
                        sTemp = ""
                    End If
                    sLPart = Trim(Left(sLPart, nPos - 1))
                    If Right(sLPart, 1) = "," Then
                        nLen = Len(sLPart) - 1
                        If nLen > 0 Then
                            sLPart = Left(sLPart, nLen)
                        End If
                    End If
                End If
            End If
        Next
    End With
    sSql = sLPart & " " & sSQLFromClause
End Sub

Private Function fnGetComboIndex(vTemp As Variant) As Integer
    'Return the index to a combo, depends on the type of the passed variable
    Dim nTemp As Integer
    nTemp = fnTypeOf(vTemp)
    If nTemp = TYPE_OBJECT Then
        Dim nCmb As Integer
        Dim nType As Integer
        Dim obj As Object
        Set obj = vTemp
        fnGetKeys nTemp, nCmb, nType, obj
    ElseIf nTemp = TYPE_INT Then
        nTemp = vTemp
    Else
        MsgBox "Type mismatch when calling functions in clsComboControl"
        nTemp = -1
        Exit Function
    End If
    fnGetComboIndex = nTemp
End Function



Private Function fnRegExpPCode(sPattern As String)
    Dim sBuffer As String * 4096
    Dim nCode As Integer
    
    sBuffer = Space(4096)

    nCode = GetRegExpPCode(sPattern, sBuffer)
    
    If nCode <= 0 Then
        fnRegExpPCode = ""
    Else
        Mid(sBuffer, nCode + 1, 1) = Chr(0)
        fnRegExpPCode = Left(sBuffer, nCode + 1)
    End If

End Function

Private Function fnTypeOf(v As Variant) As Integer
    If TypeOf v Is Object  Then
        fnTypeOf = TYPE_OBJECT
        Exit Function
    End If

    Dim sTemp As String
    Dim nLen As Integer
    Dim nChar As Integer
    Dim i As Integer
    
    fnTypeOf = TYPE_INT
    sTemp = CStr(v)
    nLen = Len(sTemp)
    If nLen = 0 Then
        fnTypeOf = TYPE_STRING
    Else
        i = 1
        Do
            nChar = Asc(Mid$(sTemp, i, 1))
            If nChar < &H30 Or nChar > &H39 Then
                fnTypeOf = TYPE_STRING
                Exit Do
            End If
            i = i + 1
        Loop Until i > nLen
    End If
End Function

Private Function fnUnscrambleDataField(sDataField As String, nCol As Integer) As String
    Dim sArray() As String
    fnParseString sArray, sDataField
    If nCol > 1 And nCol - 1 <= UBound(sArray) Then
        fnUnscrambleDataField = sArray(nCol - 1)
    Else
        fnUnscrambleDataField = sArray(0)
    End If
End Function
'david 09/21/2004
'taking care of field with alias
Private Function fnGetFieldName(ByVal sField As String, Optional bReturnAlias As Boolean = False) As String
    Dim sDataField As String
    Dim sAlias As String
    
    Dim nPosi As Integer
    nPosi = InStrRev(sField, " ")
    If nPosi > 0 Then
        'field alias found
        sAlias = Trim(Mid(sField, nPosi + 1))
        sField = Trim(Left(sField, nPosi - 1))
        nPosi = InStrRev(LCase(sField), " as")
        If nPosi > 0 Then
            sField = Trim(Left(sField, nPosi - 1))
        End If
        If InStr(sField, "||") > 0 Then
            If Left(sField, 1) <> "(" Then
                sField = "(" + sField + ")"
            End If
        End If
        sDataField = sField
    Else
        sDataField = sField
        sAlias = sField
    End If
    
    If sAlias = "" Then
        sAlias = sDataField
    End If

    If bReturnAlias Then
        fnGetFieldName = sAlias
    Else
        fnGetFieldName = sDataField
    End If
End Function

Property Let Position(nType As Integer)
    colComboControls.Item(nCurrentCombo).Position = nType
End Property

Property Get DropUp() As Integer
    'Returns the value of dropup flag
    DropUp = nDROP_UP
End Property

Public Sub ReadData(ByVal RowBuf As Object, _
                    StartLocation As Variant, _
                    ByVal ReadPriorRows As Boolean)
    Dim Bookmk As Variant
    Dim Index As Long
    Dim RelPos As Integer
    Dim RowsFetched As Integer
    Dim i As Integer
    Dim j As Integer
    
    Bookmk = StartLocation
    If ReadPriorRows Then
        RelPos = -1
    Else
        RelPos = 1
    End If
    RowsFetched = 0
    
    For i = 0 To RowBuf.RowCount - 1
        If IsNull(Bookmk) Then
            If RelPos < 0 Then  ' the grid is asking for rows prior to Bookmk
                Index = lRowCount
            Else  ' the grid is asking for rows after Bookmk
                Index = -1
            End If
        Else
            Index = Val(Bookmk)
        End If
        Index = Index + RelPos
        If Index < 0 Or Index >= lRowCount Then
            Exit For
        End If
        Bookmk = Str(Index)
        For j = 0 To RowBuf.ColumnCount - 1
            RowBuf.value(i, j) = sGridData(j, Index)
        Next j
        RowBuf.Bookmark(i) = Bookmk
        RowsFetched = RowsFetched + 1
    Next i
    RowBuf.RowCount = RowsFetched
    
End Sub


Private Sub fnSetupAlias(nIndex As Integer)
    Dim sFields() As String
    Dim sTemp As String
    Dim nPos As Integer
    Dim nLen As Integer
    Dim nKeyLen As Integer
    Dim sAlias As String
    Dim colTemp As Collection
    Dim clsTemp As clsComboSet
    Dim i As Integer
    Dim nCount As Integer
    Dim SQL As String
    Dim sRPart As String
    Dim nFieldPos() As Integer
    Dim nTemp As Integer

    Set colTemp = colComboControls.Item(nIndex).Combos
    sRPart = LCase(Trim(colComboControls.Item(nIndex).SQL))
    ReDim nFieldPos(colTemp.Count)
    
    If sRPart = "" Or colTemp Is Nothing Then
        Exit Sub
    End If
    If colTemp.Count = 0 Then
        Exit Sub
    End If
    For Each clsTemp In colTemp
        If clsTemp.DataField <> "" Then
            clsTemp.DataField = fnUnscrambleDataField(clsTemp.DataField, 0)
        End If
    Next clsTemp

    nCount = 0
    While sRPart <> ""
        nPos = InStr(sRPart, " union ")
        If nPos > 1 Then
            SQL = Left(sRPart, nPos - 1)
            nLen = Len(sRPart) - nPos - 6
            If nLen > 0 Then
                sRPart = Right(sRPart, nLen)
            Else
                sRPart = ""
            End If
        Else
            SQL = Trim(sRPart)
            sRPart = ""
        End If

        nPos = InStr(SQL, " from ")
        If nPos = 0 Or Left(SQL, 7) <> "select " Then
            If nErrorNumber <> nERROR_SQL Then
                MsgBox "Check your SQL statement"
                nErrorNumber = nERROR_SQL
            End If
            Exit Sub
        End If
        sTemp = Left(SQL, nPos - 1)
    
        nPos = InStr(sTemp, " distinct ")
        nKeyLen = 9
        If nPos = 0 Then
            nPos = InStr(sTemp, " unique ")
            nKeyLen = 7
        End If
        If nPos > 0 Then
            nLen = Len(sTemp) - nPos - nKeyLen
        Else
            nLen = Len(sTemp) - 7
        End If

        sTemp = Right$(sTemp, nLen)
        fnParseString sFields, sTemp
        
        For i = 0 To UBound(sFields)
            nPos = InStr(sFields(i), " as ")
            nKeyLen = 3
            If nPos = 0 Then
                nPos = InStr(sFields(i), " ")
                nKeyLen = 0
            End If
            If nPos > 0 Then
                nLen = Len(sFields(i)) - nPos - nKeyLen
                If nLen <= 0 Then
                    If nErrorNumber <> nERROR_SQL Then
                        MsgBox "Error in setting up search combos"
                        nErrorNumber = nERROR_SQL
                    End If
                    Exit Sub
                End If
                sAlias = Trim(Right(sFields(i), nLen))
            Else
                sAlias = Trim(sFields(i))
            End If
            nTemp = 1
            For Each clsTemp In colTemp
                If LCase(clsTemp.AliasName) = sAlias Then
                    If nPos > 0 Then
                        If nCount = 0 Then
                            clsTemp.DataField = Left(sFields(i), nPos)
                        Else
                            clsTemp.DataField = clsTemp.DataField & sFieldSep & Left(sFields(i), nPos)
                        End If
                        nFieldPos(nTemp) = i
                    Else
                        If nCount = 0 Then
                            clsTemp.DataField = sFields(i)
                        Else
                            clsTemp.DataField = clsTemp.DataField & sFieldSep & sFields(i)
                        End If
                        nFieldPos(nTemp) = i
                    End If
                    Exit For
                ElseIf LCase(clsTemp.DataField) = sAlias Then
                    If nCount = 0 Then
                        clsTemp.DataField = sFields(i)
                    Else
                        clsTemp.DataField = clsTemp.DataField & sFieldSep & sFields(i)
                    End If
                    nFieldPos(nTemp) = i
                    Exit For
                ElseIf nCount > 0 Then
                    If nFieldPos(nTemp) = i Then
                        clsTemp.DataField = clsTemp.DataField & sFieldSep & sFields(i)
                        Exit For
                    End If
                End If
                nTemp = nTemp + 1
            Next clsTemp
        Next
        nCount = nCount + 1
    Wend
    nErrorNumber = 0
End Sub

Property Let RegExpPattern(objControl As Object, sPattern As String)
    Dim lDropDown As Integer
    Dim nctrlType As Integer
    Dim clsTemp As clsComboSet

    fnGetKeys lDropDown, lComboIndex, nctrlType, objControl
    If lDropDown < 0 Then
        Exit Property
    End If
    
    Set clsTemp = colComboControls.Item(lDropDown).Combos.Item(lComboIndex)
    clsTemp.RegExpPattern = fnRegExpPCode(sPattern)
    
    If clsTemp.RegExpPattern = "" Then
        MsgBox "Regular expression error for TextBox: " & objControl.Name
    End If
    
End Property

Public Sub ReLocateTable(Optional vTemp)
    Dim lDropDown As Integer
    Dim i As Integer
    Dim colTemp As Collection
    Dim nTW As Integer

    
    If IsMissing(vTemp) Then
        lDropDown = CurrentCombo
    Else
        lDropDown = fnGetComboIndex(vTemp)
    End If

    If lDropDown <= 0 Then
        Exit Sub
    End If

    Set colTemp = colComboControls.Item(lDropDown).Combos
    nTW = 0
    
    For i = 1 To colTemp.Count
        With colTemp.Item(i)
            If i = 1 Then
                Dim nLeft As Integer
                Dim nTop As Integer
                fnGetOffsets nLeft, nTop, .Textbox
                colComboControls.Item(lDropDown).TableLeft = nLeft  ' + txtBox.Left
                colComboControls.Item(lDropDown).TableTop = nTop    '+ txtBox.Height     '+ txtBox.Top
'                .ColWidth = .TextBox.Width
                If Not .Button Is Nothing Then
                    .ColWidth = .Button.Left + .Button.Width - .Textbox.Left
                End If
            Else
'                .ColWidth = .TextBox.Left + .TextBox.Width - (colTemp.Item(i - 1).TextBox.Left + colTemp.Item(i - 1).TextBox.Width)
                If Not .Button Is Nothing Then
                    .ColWidth = .Button.Left + .Button.Width - (colTemp.Item(i - 1).Button.Left + colTemp.Item(i - 1).Button.Width)
                End If
            End If
            nTW = nTW + .ColWidth  'colTemp.Item(nBoxCount).TextBox.Width
        End With
    Next
    colComboControls.Item(lDropDown).TableWidth = nTW

End Sub

Public Sub SetColumnCaptions(ParamArray sCaps())
    Dim nCol As Integer
    
    nCol = UBound(sCaps)
    If nCol < 0 Then
        Exit Sub
    End If

    Dim colTemp As Collection
    Dim i As Integer

    Set colTemp = colComboControls.Item(nCurrentCombo).Combos
    i = 0
    For i = 0 To nCol
        If i < colTemp.Count Then
            colTemp.Item(i + 1).Caption = sCaps(i)
        End If
    Next
    
    colComboControls.Item(nCurrentCombo).HeadLines = 1
End Sub

Public Sub SetOrderingDescent(objControl As Object)
    Dim nDropDown As Integer
    Dim nCmb As Integer
    Dim nType As Integer
    
    fnGetKeys nDropDown, nCmb, nType, objControl

    If nDropDown >= 0 Then
        colComboControls.Item(nDropDown).Combos.Item(nCmb).OrderFlag = nORDERING_DESCENT
    End If
End Sub

Public Sub SetOrderingAscent(objControl As Object)
    
    Dim nDropDown As Integer
    Dim nCmb As Integer
    Dim nType As Integer
    
    fnGetKeys nDropDown, nCmb, nType, objControl

    If nDropDown > 0 Then
        colComboControls.Item(nDropDown).Combos.Item(nCmb).OrderFlag = nORDERING_ASSCENT
    End If


End Sub

Public Sub SetOrderingNone(objControl As Object)
    
    Dim nDropDown As Integer
    Dim nCmb As Integer
    Dim nType As Integer
    
    fnGetKeys nDropDown, nCmb, nType, objControl

    If nDropDown > 0 Then
        colComboControls.Item(nDropDown).Combos.Item(nCmb).OrderFlag = nORDERING_NONE
    End If

End Sub

Public Sub SetPositionAbove(Optional vTemp)
    Dim nTemp As Integer
    If IsMissing(vTemp) Then
        nTemp = CurrentCombo
    Else
        nTemp = fnGetComboIndex(vTemp)
    End If

    If nTemp > 0 Then
        colComboControls.Item(nTemp).Position = nDROP_UP
    End If

End Sub

Public Sub SetPositionBelow(Optional vTemp)
    
    Dim nTemp As Integer
    If IsMissing(vTemp) Then
        nTemp = CurrentCombo
    Else
        nTemp = fnGetComboIndex(vTemp)
    End If

    If nTemp > 0 Then
        colComboControls.Item(nTemp).Position = nDROP_DOWN
    End If
    
End Sub

Property Get SQL_DATE_TYPE() As Integer
    SQL_DATE_TYPE = TYPE_DATE
End Property

Public Function AddCombo(Optional sSql As Variant, _
                         Optional nRows As Variant, _
                         Optional nPos As Variant, _
                         Optional nDBType As Variant) As Integer

    'Add a combo to the control

    Dim colCombo As New Collection
    Dim clsDD As New clsComboDropDown

    clsDD.Cleared = False
    clsDD.HorScroll = False
    If Not IsMissing(sSql) Then
        clsDD.SQL = sSql
    End If
    If IsMissing(nRows) Then
        clsDD.VisibleRows = nVISIBLEROWS_DEFAULT
    Else
        clsDD.VisibleRows = nRows
    End If
    If IsMissing(nPos) Then
        clsDD.Position = nDROP_DOWN
    Else
        clsDD.Position = nPos
    End If
    If IsMissing(nDBType) Then
        clsDD.dbType = DB_INFORMIX
    Else
        clsDD.dbType = nDBType
    End If
    On Error GoTo errAddDropDown

    'Initialize the properties of the dropdown
    Set clsDD.Combos = New Collection
    With clsDD
        .HeadLines = 0
        .ShowNoDataMessage = True
        .ComboOn = True
        .Active = True
        .SearchOnReturn = True
        .ExtendReset = False
        .ReQuery = True
        .DataEntered = False
    End With
    
    colComboControls.Add clsDD
    nCurrentCombo = colComboControls.Count
    AddCombo = nCurrentCombo
    
extAddDropDown:
    On Error GoTo 0
    Exit Function
errAddDropDown:
    MsgBox Err.Description
    AddCombo = -1
    Resume extAddDropDown
End Function

Property Get dbAccess() As Integer
    'Returns the value of Access database flag
    dbAccess = DB_ACCESS
End Property
Property Get dbInformix() As Integer
    'Returns the value of Informix database flag
    dbInformix = DB_INFORMIX
End Property

Property Get DropDown() As Integer
    'Returns the value of dropdown flag
    DropDown = nDROP_DOWN
End Property
Public Sub SetComboSQL(SQL As String, Optional vTemp As Variant)
    Dim nIndex As Integer
    
    If IsMissing(vTemp) Then
        nIndex = nCurrentCombo
    Else
        nIndex = fnGetComboIndex(vTemp)
    End If
    If nIndex < 0 Then
        Exit Sub
    End If

    If colComboControls Is Nothing Then
        Set colComboControls = New Collection
    End If
    If colComboControls.Count = 0 Then
        Dim clsDD As New clsComboDropDown
        colComboControls.Add clsDD
    End If
    
    If nIndex > 0 And nIndex <= colComboControls.Count Then
        colComboControls.Item(nIndex).SQL = SQL
    End If
    
    fnSetupAlias nIndex
End Sub

Public Function GetComboSQL(Optional vIndex As Variant) As String
    Dim nIndex As Integer
    
    If IsMissing(vIndex) Then
        nIndex = nCurrentCombo
    Else
        nIndex = vIndex
    End If
    
    If nIndex > 0 And nIndex <= colComboControls.Count Then
        GetComboSQL = colComboControls.Item(nIndex).SQL
    End If
End Function

Public Sub Change(objControl As Object, _
                  Optional vClearCombo As Variant)
    'Event call for text change in a textbox
    If m_bFillingText Then
        Exit Sub
    End If
    If objControl Is Nothing Then
        Exit Sub
    End If
    
    If objControl.TabIndex = tgDropDown.TabIndex Then
        Exit Sub
    End If

    Dim lDropDown As Integer
    Dim nctrlType As Integer
    Dim lCombo As Integer
    Dim bClearCombo As Boolean
    
    'Index for the dropdown
    fnGetKeys lDropDown, lCombo, nctrlType, objControl
    
    If lDropDown < 0 Then
        Exit Sub
    End If
    
    If nctrlType = CONTROL_TYPE_TEXTBOX Then
        
        Dim clsTemp As clsComboSet
        
        Set clsTemp = colComboControls.Item(lDropDown).Combos.Item(lCombo)
        clsTemp.Changed = True
        If clsTemp.AppendCriteria Then
            colComboControls.Item(lDropDown).DataEntered = True
        End If
        If nErrorNumber <> nERROR_DLLFILE Then
        
            'Check the input
            If tfnRegExpControlChange(objControl, clsTemp.RegExpPattern) Then
            End If
            'Clear the text boxes if it is the key field
            If IsMissing(vClearCombo) Then
                bClearCombo = True
            Else
                bClearCombo = vClearCombo
            End If
            If clsTemp.KeyField Then
                If bClearCombo Then
                    If Not colComboControls.Item(lDropDown).Cleared Then
                        subClearCombo clsTemp, lDropDown, True
                    End If
                End If
            End If
        End If
    End If

End Sub


Private Sub fnGetOffsets(nLeft As Integer, _
                         nTop As Integer, _
                         txtBox As Object)

    'Find out the offsets of the table for a combo

    #If Win32 Then
        Dim nTbl_hWnd As Long
        Dim n_hWnd As Long
        Dim n_ParenthWnd As Long
    #Else
        Dim nTbl_hWnd As Integer
        Dim n_hWnd As Integer
        Dim n_ParenthWnd As Integer
    #End If
    Dim nClientLeft As Single
    Dim nClientTop As Single
    
    nTbl_hWnd = frmParent.hwnd
    n_hWnd = txtBox.hwnd
    n_ParenthWnd = GetParent(n_hWnd)
    nLeft = 0
    nTop = txtBox.Height
    While n_hWnd <> nTbl_hWnd And n_ParenthWnd > 0
        fnGetLeftTop nClientLeft, nClientTop, n_ParenthWnd, n_hWnd
        nLeft = nLeft + nClientLeft
        nTop = nTop + nClientTop
        n_hWnd = n_ParenthWnd
        n_ParenthWnd = GetParent(n_hWnd)
    Wend
End Sub

#If Win32 Then
    Private Function fnGetLeftTop(nLeft As Single, nTop As Single, ByVal nParenthWnd As Long, ByVal n_hWnd As Long)
        'Returns the Left and Top of a control
        Dim myRect As RECT
        Dim myPoint As POINTAPI
        Dim nLimit As Integer
        
        nLimit = frmParent.Width
    
        GetWindowRect n_hWnd, myRect
    
        myPoint.x = myRect.Left
        myPoint.y = myRect.Top
        
        If ScreenToClient(nParenthWnd, myPoint) Then
    '        nLeft = (myPoint.X - 1) * Screen.TwipsPerPixelX    'For True Gridtable
            nLeft = myPoint.x * Screen.TwipsPerPixelX     'For textbox
            If nLeft < -nLimit Or nLeft > nLimit Then
                nLeft = Screen.TwipsPerPixelX
            End If
    '        nTop = (myPoint.Y - 1) * Screen.TwipsPerPixelY    'For True Gridtable
            nTop = myPoint.y * Screen.TwipsPerPixelY       'For textbox
            If nTop < 0 Or nTop > frmParent.Height Then
                nTop = Screen.TwipsPerPixelX
            End If
        End If
    
    End Function
#Else
    Private Function fnGetLeftTop(nLeft As Single, nTop As Single, ByVal nParenthWnd As Integer, ByVal n_hWnd As Integer)
        'Returns the Left and Top of a control
        Dim myRect As RECT
        Dim myPoint As POINTAPI
        Dim nLimit As Integer
        
        nLimit = frmParent.Width
    
        GetWindowRect n_hWnd, myRect
    
        myPoint.x = myRect.Left
        myPoint.y = myRect.Top
        
        If ScreenToClient(nParenthWnd, myPoint) Then
    '        nLeft = (myPoint.X - 1) * Screen.TwipsPerPixelX    'For True Gridtable
            nLeft = myPoint.x * Screen.TwipsPerPixelX     'For textbox
            If nLeft < -nLimit Or nLeft > nLimit Then
                nLeft = Screen.TwipsPerPixelX
            End If
    '        nTop = (myPoint.Y - 1) * Screen.TwipsPerPixelY    'For True Gridtable
            nTop = myPoint.y * Screen.TwipsPerPixelY       'For textbox
            If nTop < 0 Or nTop > frmParent.Height Then
                nTop = Screen.TwipsPerPixelX
            End If
        End If
    
    End Function
#End If

Property Let CurrentCombo(c As Integer)
    'Sets the index of the current SQL
    nCurrentCombo = c
End Property
Property Get CurrentCombo() As Integer
    'Returns the index of the current SQL
    CurrentCombo = nCurrentCombo
End Property


Public Sub AddExtraColumn(sField As String, _
                          Optional nWidth As Variant, _
                          Optional vBox As Variant, _
                          Optional vExt As Variant, _
                          Optional vCombo As Variant, _
                          Optional sOrderByFields As String = "")
    
    'Add an extra column to the dropdown table.
    ' This column can not be used to enter search criteria.
    ' If vBox is missed, this column will not be populated with any textbox
    
    Dim nCurr As Integer
    Dim clsDD As clsComboDropDown
    Dim clsCombo As clsComboSet
    Dim colTemp As Collection
    Dim nExtn As Integer
    Dim sTemp As String
    Dim i As Integer

    If IsMissing(vCombo) Then
        nCurr = nCurrentCombo
    Else
        nCurr = vCombo
    End If

    If IsMissing(vExt) Then
        nExtn = EXTR_COL_EXT_RIGHT
    Else
        nExtn = vExt
    End If
    
    If nCurr > colComboControls.Count Then
        Set clsDD = New clsComboDropDown
        Set clsDD.Combos = New Collection
        colComboControls.Add clsDD
        nCurr = colComboControls.Count
    Else
        Set clsDD = colComboControls.Item(nCurr)
    End If
    
    Set colTemp = clsDD.Combos
    
    sTemp = UCase(Trim(sField))
    For i = 1 To colTemp.Count
        With colTemp.Item(i)
            If .Button Is Nothing Then
                If sTemp = UCase(Trim(.AliasName)) Then
                    Set clsCombo = colTemp.Item(i)
                    Exit For
                End If
            End If
        End With
    Next
    If clsCombo Is Nothing Then
        Set clsCombo = New clsComboSet
        'Add to the dropdown control
        If nExtn = EXTR_COL_EXT_LEFT Then
            colTemp.Add clsCombo, , 1
        Else
            colTemp.Add clsCombo
        End If
    End If
    
    With clsCombo
        If IsMissing(vBox) Then
            .Populate = False
        Else
            Set .Textbox = vBox
            .Populate = True
        End If
        Set .Button = Nothing
        .AliasName = Trim(sField)
        .FieldValue = ""
        .AppendCriteria = False
        .ExtDir = nExtn
        .Changed = True
        .Valid = False
        .Visible = True
    End With

    fnSetupAlias nCurr   'For population of the text
    
    If IsMissing(nWidth) Then
        If IsMissing(vBox) Then
            clsCombo.ColWidth = DEFAULT_COLWIDTH
        Else
            clsCombo.ColWidth = vBox.Width
        End If
    Else
        clsCombo.ColWidth = nWidth
    End If
    
    'david 01/28/2002
    If sOrderByFields <> "" Then
        clsCombo.OrderByFields = sOrderByFields
    End If
    '''''''''''''''''
    
    clsDD.TableWidth = clsDD.TableWidth + clsCombo.ColWidth
    If clsCombo.ExtDir = EXTR_COL_EXT_LEFT Then
        clsDD.TableLeft = clsDD.TableLeft - clsCombo.ColWidth
    End If
    
End Sub
Public Sub SetToAccess(Optional vTemp)
' Set the database type to local access database
    Dim nCombo As Integer
    If IsMissing(vTemp) Then
        nCombo = nCurrentCombo
    Else
        nCombo = fnGetComboIndex(vTemp)
    End If
    
    If nCombo > 0 Then
        colComboControls.Item(nCombo).dbType = DB_ACCESS
    End If
End Sub

Public Sub SetToInformix(Optional vTemp)
' Set the database type to informix
    Dim nCombo As Integer
    If IsMissing(vTemp) Then
        nCombo = nCurrentCombo
    Else
        nCombo = fnGetComboIndex(vTemp)
    End If
    
    If nCombo > 0 Then
        colComboControls.Item(nCombo).dbType = DB_INFORMIX
    End If
End Sub

Property Get ComboCount() As Integer
    'Returns the number of combos in the control
    ComboCount = colComboControls.Count
End Property

Private Sub fnGetKeys(nDropDown As Integer, _
                      nCombo As Integer, _
                      nType As Integer, _
                      objKey As Object, _
                      Optional vQueryActive As Variant)
    'Return the index of combo, comboset, and type of the control
    ' for the input control

    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim colTemp As Collection
    Dim bQueryActive As Boolean
    
    nDropDown = -1
    nCombo = -1
    If objKey Is Nothing Then
        Exit Sub
    End If
    If IsMissing(vQueryActive) Then
        bQueryActive = False
    Else
        bQueryActive = vQueryActive
    End If
    
    For i = 1 To colComboControls.Count
        With colComboControls.Item(i)
            If .Active Or bQueryActive Then
                Set colTemp = .Combos
                For j = 1 To colTemp.Count
                    If Not (colTemp.Item(j).Textbox Is Nothing) Then
                        If colTemp.Item(j).Textbox.TabIndex = objKey.TabIndex Then
                            nDropDown = i
                            nCombo = j
                            nType = CONTROL_TYPE_TEXTBOX
                            If Not (colTemp.Item(j).Button Is Nothing) Then
                                Exit Sub
                            End If
                        End If
                    End If
                    If Not (colTemp.Item(j).Button Is Nothing) Then
                        If colTemp.Item(j).Button.TabIndex = objKey.TabIndex Then
                            nDropDown = i
                            nCombo = j
                            nType = CONTROL_TYPE_BUTTON
                            Exit Sub
                        End If
                    End If
                Next j
            End If
        End With
    Next i
End Sub

Private Sub subIndexFromField(nDropDown As Integer, _
                              nCombo As Integer, _
                              sField As String)
    'Return the index of combo, comboset, and type of the control
    ' for the input control

    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim colTemp As Collection
    
    nDropDown = -1
    nCombo = -1
    For i = 1 To colComboControls.Count
        With colComboControls.Item(i)
            If .Active Then
                Set colTemp = .Combos
                For j = 1 To colTemp.Count
                    'david 09/21/2004
                    If fnGetFieldName(colTemp.Item(j).DataField) = sField Or _
                       fnGetFieldName(colTemp.Item(j).DataField, True) = sField Then
                        nDropDown = i
                        nCombo = j
                        Exit Sub
                    End If
                Next j
            End If
        End With
    Next i
End Sub


Private Sub fnSetFocus(cntlTemp As Object)
    Const nTrialNumber As Integer = 1
    Dim nCount As Integer

    nCount = 0

    On Error GoTo errSetFocus
    cntlTemp.SetFocus
    
extSetFocus:
    On Error GoTo 0
    Exit Sub
    
errSetFocus:
    If nCount < nTrialNumber Then
        nCount = nCount + 1
        DoEvents
        Resume
    Else
        Resume extSetFocus
    End If
End Sub

Private Function fnLoadTable(ByVal lDropDown As Integer, _
                             Optional bCalledByHeaderClick As Boolean = False, _
                             Optional ColIndex As Integer = -1) As Boolean
    'Load the dropdown table with the given SQL and criteria
    
    Const szUNION As String = " UNION "

    Dim szCode As String
    Dim szQuery As String
    Dim sColName As String
    Dim rsTemp As Recordset
    Dim nPos As Integer
    Dim lCol As Integer
    Dim nColCount As Integer
    Dim nTop As Integer
    Dim nLeft As Integer
    Dim colCombo As Collection
    
    If m_bProtoType Then
        Exit Function
    End If
    
    'david 01/28/2002
    If bCalledByHeaderClick Then
        If ColIndex < 0 Then
            Exit Function
        End If
    End If
    ''''''''''''''''''
    
    'Make our mouse pointer an hourglass
    Screen.MousePointer = vbHourglass
    bValidSelection = False
    
    'david 01/28/2002
    If Not bCalledByHeaderClick Then
        tgDropDown.Visible = False
    End If
    '''''''''''''''''
    frmParent.Cls

    lRecordsReturned = -1
    If lDropDown < 0 Then
        Screen.MousePointer = vbDefault
        Exit Function
    End If

    'david 01/28/2002
    If Not bCalledByHeaderClick Then
        If clsCurrentDropdown Is colComboControls.Item(lDropDown) Then
            If Not clsCurrentDropdown.ReQuery Then
                If Not clsCurrentDropdown.DataEntered Then
                    If datLink.Recordset Is Nothing Then
                        Exit Function
                    End If
                    If datLink.Recordset.RecordCount > 1 Then
                        tgDropDown.Visible = True
                        tgDropDown.ZOrder 0
                        bValidSelection = False
                    End If
                    fnSetFocus tgDropDown
                    fnLoadTable = True
                    Screen.MousePointer = vbDefault
                    Exit Function
                End If
            End If
        Else
            'Get the dropdown class
            Set clsCurrentDropdown = colComboControls.Item(lDropDown)
        End If
    
        'david 01/28/2002
        lHeadClickDropDown = lDropDown
        '''''''''''''''''
    Else
        Set clsCurrentDropdown = colComboControls.Item(lDropDown)
    End If
    '''''''''''''''
    
    Set colCombo = clsCurrentDropdown.Combos
    szQuery = clsCurrentDropdown.SQL
    'If there is a 'UNION statement, don't do anything
    nPos = InStr(UCase(szQuery), szUNION)
    
    If nPos = 0 Then
        szQuery = fnAppendCriteria(szQuery, clsCurrentDropdown.dbType, 0)
        'david 09/21/2004
        sColName = fnGetFieldName(colCombo.Item(lComboIndex).DataField, True)
        '''''''''''''''''
    Else
        Dim szQuery1 As String
        Dim szQuery2 As String
        Dim sRPart As String
        Dim nLen As Integer
        
        sRPart = szQuery
        szQuery = ""
        lCol = 1
        
        While sRPart <> ""
            If nPos > 1 Then
                szQuery1 = Left(sRPart, nPos - 1)
                nLen = Len(sRPart) - nPos - Len(szUNION) + 1
                If nLen > 0 Then
                    sRPart = Right(sRPart, nLen)
                Else
                    sRPart = ""
                End If
            Else
                szQuery1 = sRPart
                sRPart = ""
            End If

            szQuery1 = fnAppendCriteria(szQuery1, clsCurrentDropdown.dbType, lCol)
            
            If szQuery = "" Then
                szQuery = szQuery1
            Else
                szQuery = szQuery & szUNION & szQuery1
            End If
            nPos = InStr(sRPart, szUNION)
            If nPos = 0 Then
                nPos = Len(sRPart) + 1
            End If
            lCol = lCol + 1
        Wend
        
        sColName = CStr(lComboIndex)
    End If
    
    'david 01/28/2002
    Screen.MousePointer = vbHourglass
    
    If bCalledByHeaderClick Then
        Dim sTemp As String
        
        bHeaderClick = True
        
        If colCombo.Item(lComboIndex).OrderByFields <> "" Then
            sTemp = colCombo.Item(lComboIndex).OrderByFields
        Else
            'david 09/21/2004
            sTemp = fnGetFieldName(tgDropDown.Columns(ColIndex).DataField, True)
            '''''''''''''''''
        End If
        
        sTemp = fnCnvtOrderByFieldsToColIndex(szQuery, sTemp)
        
        If colCombo.Item(lComboIndex).PrevOrderIsAscending() Then
            colCombo.Item(lComboIndex).PrevOrderIsAscending() = False
            'order by descending
            sTemp = fnFixFirstSortOrder(sTemp, True)
            
        Else
            colCombo.Item(lComboIndex).PrevOrderIsAscending() = True
            'order by ascending
            sTemp = fnFixFirstSortOrder(sTemp, False)
        End If
    
        If sTemp <> "" Then
            szQuery = szQuery & " ORDER BY " & sTemp
        End If
    Else
        If colCombo.Item(lComboIndex).OrderFlag <> nORDERING_NONE Then
            szQuery = szQuery & " ORDER BY " & sColName
            
            If colCombo.Item(lComboIndex).OrderFlag = nORDERING_DESCENT Then
                szQuery = szQuery & " DESC"
                colCombo.Item(lComboIndex).PrevOrderIsAscending() = False
            Else
                colCombo.Item(lComboIndex).PrevOrderIsAscending() = True
            End If
        End If
    End If
    ''''''''''''''''''

    'Get the records
    If Not fnGetRecord(rsTemp, szQuery, clsCurrentDropdown.dbType) Then
        Screen.MousePointer = vbDefault
        ' Set the return value false
        tgDropDown.Visible = False
        frmParent.Cls
        fnSetFocus colCombo.Item(lComboIndex).Textbox
        Exit Function
    End If
    
    Screen.MousePointer = vbHourglass
    If tgDropDown.DataMode = dbgUnbound Then
        If rsTemp.RecordCount > 0 Then
            ReDim sGridData(rsTemp.Fields.Count - 1, rsTemp.RecordCount - 1)
            lRowCount = 0
            While Not rsTemp.EOF
                For lCol = 1 To colCombo.Count
                    If IsNull(rsTemp.Fields(colCombo.Item(lCol).AliasName)) Then
                        sGridData(lCol - 1, lRowCount) = ""
                    Else
                        sGridData(lCol - 1, lRowCount) = Trim(rsTemp.Fields(colCombo.Item(lCol).AliasName))
                    End If
                Next
                lRowCount = lRowCount + 1
                rsTemp.MoveNext
            Wend
        End If
    End If
    
    'Check there is something for us to display before we show the results
    Screen.MousePointer = vbHourglass
    If rsTemp.RecordCount > 0 Then
        ' Set the return values
        If rsTemp.RecordCount = 1 Then
            If tgDropDown.DataMode = dbgUnbound Then
                tgDropDown.FirstRow = 0
                tgDropDown.Row = 0
            Else
                Set datLink.Recordset = rsTemp
            End If
            fnGetText
            tgDropDown.Visible = False
            bValidSelection = True
            frmParent.Cls
            fnSetFocus colCombo.Item(lComboIndex).Textbox
        Else
            If tgDropDown.DataMode <> dbgUnbound Then
                Set datLink.Recordset = rsTemp
            End If
            
            While tgDropDown.Columns.Count > 0
                tgDropDown.Columns.Remove 0
            Wend
            
            nColCount = 0
            nPos = 1
            nLen = clsCurrentDropdown.TableWidth
            
            '#Add two vals WJ 04/22/2003
            Dim nDataType As Integer
            Dim right_nbr As Integer
            For lCol = 1 To colCombo.Count
                If colCombo.Item(lCol).ColWidth > 0 Then
                    If colCombo.Item(lCol).Visible Then
                        nPos = lCol
                        tgDropDown.Columns.Add nColCount
                        With tgDropDown.Columns(nColCount)
                            .Width = colCombo.Item(lCol).ColWidth
                            'david 09/21/2004
                            .DataField = fnGetFieldName(colCombo.Item(lCol).AliasName, True)
                            '''''''''''''''''
                            .Caption = colCombo.Item(lCol).Caption
                            '-----------Start for #407922 ----------
                            '#Add the following by WJ on 04/22/2003
                            '#for #407922. Detect the Data Type and
                            '#format/Align the display fields
                            nDataType = fnFieldType(rsTemp, .DataField, clsCurrentDropdown.dbType, right_nbr)
                            '0 - left; 1 - right, 2 - mid
                            Select Case nDataType
                                Case DB_NUM_TYPE
                                    .Alignment = 1
                                    If right_nbr >= 1 And right_nbr <= 8 Then
                                        .NumberFormat = "0." & String(right_nbr, "0")
                                    End If
                                Case DB_INT_TYPE
                                    .Alignment = 0
                                Case DB_DATE_TYPE
                                   .Alignment = 0
                                   .NumberFormat = "mm/dd/yyyy"
                                Case Else
                                   .Alignment = 0
                            End Select
                            '-----------End for #407922 ----------
                            .Visible = True
                        End With
                        nColCount = nColCount + 1
                    Else
                        nLen = nLen - colCombo.Item(lCol).ColWidth
                    End If
                End If
            Next
            
            tgDropDown.HeadLines = clsCurrentDropdown.HeadLines
            
            'david 01/28/2002
            If tgDropDown.HeadLines > 0 Then
                tgDropDown.ColumnHeaders = True
            Else
                tgDropDown.ColumnHeaders = False
            End If
            
            nLeft = clsCurrentDropdown.TableLeft
            
            If colCombo.Item(nPos).Button Is Nothing Then
                tgDropDown.Width = nLen + SCROLLBAR_WIDTH
            Else
                lCol = tgDropDown.Columns.Count - 1
                tgDropDown.Columns(lCol).Width = tgDropDown.Columns(lCol).Width - SCROLLBAR_WIDTH
                tgDropDown.Width = nLen
            End If

            Dim nH As Integer
            
            nH = clsCurrentDropdown.VisibleRows
            
            If clsCurrentDropdown.HorScroll Then
                nH = nH - 1
            End If
            
            If rsTemp.RecordCount <= nH Then
                nH = rsTemp.RecordCount
                If clsCurrentDropdown.HorScroll Then
                    nH = nH + 1
                End If
                If clsCurrentDropdown.HorScroll Then
                    tgDropDown.ScrollBars = dbgHorizontal
                Else
                    tgDropDown.ScrollBars = dbgNone
                End If
            Else
                nH = clsCurrentDropdown.VisibleRows
                If clsCurrentDropdown.HorScroll Then
                    tgDropDown.ScrollBars = dbgBoth
                Else
                    tgDropDown.ScrollBars = dbgVertical
                End If
            End If
            
            tgDropDown.Height = (tgDropDown.HeadLines + nH) * tgDropDown.RowHeight + nHEIGHTOFFSET * Screen.TwipsPerPixelY
            
            nTop = fnGetTop
            If tgDropDown.DataMode = dbgUnbound Then
                tgDropDown.Refresh
            End If
            
            tgDropDown.Move nLeft, nTop
            tgDropDown.Visible = True
            tgDropDown.ZOrder 0
            bValidSelection = False
            fnSetFocus tgDropDown
            fnLoadTable = True
        End If
    Else
        If clsCurrentDropdown.ShowNoDataMessage Then
            ' Set the return value false
            Dim sMsg As String
            Screen.MousePointer = vbDefault
            If clsCurrentDropdown.ErrorMessage = "" Then
                sMsg = "No record found for the given criteria"
            Else
                sMsg = clsCurrentDropdown.ErrorMessage
            End If
            MsgBox sMsg, vbOKOnly + vbInformation
        End If
        fnSetFocus colCombo.Item(lComboIndex).Textbox
        fnLoadTable = False
    End If
    
    lRecordsReturned = rsTemp.RecordCount
    
    'Reset our mouse pointer to the default
    Screen.MousePointer = vbDefault
End Function

Private Sub fnParseString(sParam() As String, _
                          sSrc As String, _
                          Optional vStart As Variant, _
                          Optional vEnd As Variant)
                          
    If Trim(sSrc) = "" Then
        Exit Sub
    End If

    Const nArrayInc As Integer = 5
    Dim i1 As Integer
    Dim i2 As Integer
    Dim k As Integer
    Dim nEnd As Integer
    If IsMissing(vStart) Then
        i1 = 1
    Else
        i1 = vStart
    End If
    If IsMissing(vEnd) Then
        nEnd = Len(sSrc)
    Else
        nEnd = vEnd
    End If
    If i1 < 1 Then i1 = 1
    i2 = 1
    k = 0
    ReDim sParam(nArrayInc)
    While i1 <= nEnd And i2 > 0 And i2 <= nEnd And k < UBound(sParam)
        i2 = InStr(i1, sSrc, ",")
        If i2 > i1 And i2 <= nEnd Then
            sParam(k) = Trim$(Mid$(sSrc, i1, i2 - i1))
            k = k + 1
            If k > UBound(sParam) Then
                ReDim Preserve sParam(k + nArrayInc)
            End If
            i1 = i2 + 1
        End If
    Wend
    If i2 <= nEnd Then
        sParam(k) = Trim$(Mid$(sSrc, i1, nEnd - i1 + 1))
        ReDim Preserve sParam(k)
    Else
        If k > 0 Then
            sParam(k - 1) = Trim$(Mid$(sSrc, i1, nEnd - i1 + 1))
            ReDim Preserve sParam(k - 1)
        End If
    End If
End Sub

Public Sub SelChange(Cancel As Integer)
    Cancel = True
End Sub

Public Function SQL_DECIMAL_TYPE(Optional ByVal vWhole As Variant, _
                                 Optional ByVal vDecimal As Variant) As Integer
    Const DEFAULT_LENGTH = 8
    Dim nWhole As Integer
    Dim nDecimal As Integer
    
    If IsMissing(vWhole) Then
        nWhole = DEFAULT_LENGTH
    Else
        nWhole = Val(vWhole)
        If nWhole <= 0 Then
            nWhole = DEFAULT_LENGTH
        End If
    End If
    If IsMissing(vDecimal) Then
        nDecimal = 0
    Else
        nDecimal = Val(vDecimal)
    End If
    SQL_DECIMAL_TYPE = -(TYPE_DCM_INT_MASK * nWhole + nDecimal)
End Function

Property Get SQL_LONG_TYPE() As Integer
    SQL_LONG_TYPE = TYPE_LONG
End Property


Property Get SQL_INT_TYPE() As Integer
    SQL_INT_TYPE = TYPE_INT
End Property

Property Get SQL_PHONE_TYPE() As Integer 'WJ
    SQL_PHONE_TYPE = TYPE_PHONE
End Property

Public Function Keypress(objControl As Object, KeyAscii As Integer) As Boolean
    If objControl.TabIndex = tgDropDown.TabIndex Then
        If KeyAscii = vbKeyReturn Then
            fnGetText
            fnSetFocus clsCurrentDropdown.Combos.Item(lComboIndex).Textbox
            tgDropDown.Visible = False
        ElseIf KeyAscii = vbKeyEscape Then
            fnSetFocus clsCurrentDropdown.Combos.Item(lComboIndex).Textbox
            tgDropDown.Visible = False
        End If
    Else
        Dim lDropDown As Integer
        Dim nctrlType As Integer
        Dim clsTemp As clsComboSet

        fnGetKeys lDropDown, lComboIndex, nctrlType, objControl
        If lDropDown < 0 Then
            Keypress = True
            Exit Function
        End If
        
        Set clsTemp = colComboControls.Item(lDropDown).Combos.Item(lComboIndex)
        If nctrlType = CONTROL_TYPE_BUTTON Then
            If colComboControls.Item(lDropDown).ComboOn Then
                fnLoadTable lDropDown
            End If
        Else
            If KeyAscii = vbKeyReturn Then
                Keypress = False
                If colComboControls.Item(lDropDown).ComboOn Then
                    If m_bProtoType Then
                        Keypress = False
                        bValidSelection = True
                        Exit Function
                    End If
                    If Not clsTemp.AppendCriteria Then
                        SingleRecordSelected = True
                        Exit Function
                    End If
                    If Not (clsTemp.Button Is Nothing) Then
                        If colComboControls.Item(lDropDown).SearchOnReturn Then
                            fnLoadTable lDropDown
                        End If
                    End If
                End If
            ElseIf KeyAscii = KEY_CTRL_Y Then
                KeyAscii = 0
                Keypress = True
                If clsTemp.Button Is Nothing Then
                    objControl.Text = ""
                    Exit Function
                End If
                subClearCombo Nothing, lDropDown, True
            ElseIf KeyAscii = vbKeyEscape Then
                ' wipe the current text box's text
                objControl.Text = ""
                KeyAscii = 0
                Keypress = True
            Else
                ' validate text + key against the regular expression
                On Error GoTo errKeyPress
                If nErrorNumber <> nERROR_DLLFILE Then
                    Keypress = tfnRegExpControlKeyPress(objControl, KeyAscii, clsTemp.RegExpPattern)
                Else
                    Keypress = True
                End If
            End If
        End If
    End If
extKeyPress:
    On Error GoTo 0
    Exit Function
errKeyPress:
    If Err.Number = 53 Then
        Keypress = True
        Err.Clear
        Resume extKeyPress
    End If
End Function

Public Function AddComboBox(txtBox As Object, _
                            cmdButton As Object, _
                            sField As String, _
                            Optional nType As Variant, _
                            Optional nWidth As Variant, _
                            Optional sOrderByFields As String = "") As Boolean

    'Add a comboset(a text box and a command button) to a combo
    If frmParent Is Nothing Then
        'The parent form is required for the setup
        MsgBox "You need to set the form property first"
        Exit Function
    End If

    Dim clsDD As clsComboDropDown
    Dim clsCombo As New clsComboSet
    Dim colTemp As Collection
    Dim nWCol As Integer
    Dim bAdd As Boolean
    
    AddComboBox = False

'    On Error GoTo errAddComboBox
    
    'Use the current combo
    Set clsDD = colComboControls.Item(nCurrentCombo)
    If clsDD.Combos Is Nothing Then
        Set clsDD.Combos = New Collection
    End If
    
    Set colTemp = clsDD.Combos
    With clsCombo
        If colTemp.Count = 0 Then
            .KeyField = True
        Else
            .KeyField = False
        End If
        Set .Textbox = txtBox
        Set .Button = cmdButton
        .AliasName = Trim(sField)
        .DataField = Trim(sField)
        .FieldValue = ""
        .AppendCriteria = True
        .Populate = True
        .Changed = True
        .Valid = False
        .Visible = True
    End With
    colTemp.Add clsCombo

    fnSetupAlias nCurrentCombo     'For population of the text

    Dim nW As Integer
    Dim nBoxCount As Integer
    nBoxCount = colTemp.Count
    'Determine the column width of the dropdown table
    If IsMissing(nWidth) Then
        subColWidth nWCol, bAdd, colTemp, nBoxCount
        clsCombo.ColWidth = nWCol
        clsCombo.AutoWidth = True
    Else
        clsCombo.ColWidth = nWidth
        bAdd = True
        clsCombo.AutoWidth = False
    End If
'    clsCombo.ColWidth = nW
    
    'Set the left and top properties of the dropdown class
    If nBoxCount = 1 Then
        Dim nLeft As Integer
        Dim nTop As Integer
        fnGetOffsets nLeft, nTop, txtBox
        clsDD.TableLeft = nLeft
        clsDD.TableTop = nTop
    End If
    If bAdd Then
        clsDD.TableWidth = clsDD.TableWidth + clsCombo.ColWidth   'nW   'colTemp.Item(nBoxCount).TextBox.Width
    Else
        clsDD.HorScroll = True
    End If
    
    'Set the regular expression pattern
    If IsMissing(nType) Then
        clsCombo.ComboType = TYPE_STRING
        clsCombo.RegExpPattern = fnGetPCode(TYPE_STRING)
    Else
        subSetType txtBox, clsCombo, nType
    End If
    
    If clsCombo.RegExpPattern = "" Then
        MsgBox "Regular expression error for TextBox: " & txtBox.Name
    Else
        AddComboBox = True
        nErrorNumber = 0
    End If
    
    'david 01/28/2002
    If sOrderByFields <> "" Then
        clsCombo.OrderByFields = sOrderByFields
    End If
    '''''''''''''''''
    
extAddComboBox:
    On Error GoTo 0
    Exit Function
errAddComboBox:
    AddComboBox = False
    If nErrorNumber = 0 Then
        If Err.Number = nERROR_DLLFILE Then
            'File not found (DLL file for the regexp)
            MsgBox "Regular expression library is not found.", vbExclamation + vbOKOnly
            nErrorNumber = nERROR_DLLFILE
        Else
            MsgBox "Can not add text box to drop down combo." & vbCrLf & vbCrLf & "Error number :  " & CStr(Err.Number) & vbCrLf & "Error Details:  " & Err.Description, vbOKOnly + vbExclamation
        End If
    End If
    Err.Clear
    Exit Function
End Function

Public Function AddCriteriaField(txtBox As Object, _
                                 sField As String, _
                                 Optional nType As Variant, _
                                 Optional vPopulate As Variant) As Boolean

    'Add a comboset(a text box and a command button) to a combo
    If frmParent Is Nothing Then
        'The parent form is required for the setup
        MsgBox "You need to set the form property first"
        Exit Function
    End If

    Dim clsDD As clsComboDropDown
    Dim clsCombo As New clsComboSet
    Dim colTemp As Collection
    Dim bPopulate As Boolean
    
    AddCriteriaField = False

    If IsMissing(vPopulate) Then
        bPopulate = False
    Else
        bPopulate = vPopulate
    End If
'    On Error GoTo errAddComboBox
    
    'Use the current combo
    Set clsDD = colComboControls.Item(nCurrentCombo)
    If clsDD.Combos Is Nothing Then
        Set clsDD.Combos = New Collection
    End If
    
    Set colTemp = clsDD.Combos
    With clsCombo
        If colTemp.Count = 0 Then
            .KeyField = True
        Else
            .KeyField = False
        End If
        Set .Textbox = txtBox
        .AliasName = Trim(sField)
        .DataField = Trim(sField)
        .FieldValue = ""
        .AppendCriteria = True
        .Populate = bPopulate
        .Changed = True
        .Valid = False
        .Visible = True
    End With
    colTemp.Add clsCombo

    fnSetupAlias nCurrentCombo     'For population of the text

    'Set the regular expression pattern
    If IsMissing(nType) Then
        clsCombo.ComboType = TYPE_STRING
        clsCombo.RegExpPattern = fnGetPCode(TYPE_STRING)
    Else
        clsCombo.ComboType = nType
        clsCombo.RegExpPattern = fnGetPCode(nType)
        If nType < MAX_TYPES Then
            txtBox.MaxLength = -nType \ TYPE_DCM_INT_MASK
        End If
    End If
    
    If clsCombo.RegExpPattern = "" Then
        MsgBox "Regular expression error for TextBox: " & txtBox.Name
    Else
        AddCriteriaField = True
        nErrorNumber = 0
    End If
    
extAddComboBox:
    On Error GoTo 0
    Exit Function
errAddComboBox:
    AddCriteriaField = False
    If nErrorNumber = 0 Then
        If Err.Number = nERROR_DLLFILE Then
            'File not found (DLL file for the regexp)
            MsgBox "Regular expression library is not found.", vbExclamation + vbOKOnly
            nErrorNumber = nERROR_DLLFILE
        Else
            MsgBox "Can not add text box to drop down combo." & vbCrLf & vbCrLf & "Error number :  " & CStr(Err.Number) & vbCrLf & "Error Details:  " & Err.Description, vbOKOnly + vbExclamation
        End If
    End If
    Err.Clear
    Exit Function
End Function

Private Sub subClearCombo(clsTemp As clsComboSet, _
                          ByVal lDropDown As Integer, _
                          ByVal bClearKey As Boolean)

    Dim ciTemp As clsComboSet
    ' wipe out all text boxes' text except for the key fields
    If colComboControls.Item(lDropDown).ComboOn Then
        colComboControls.Item(lDropDown).Cleared = True
        For Each ciTemp In colComboControls.Item(lDropDown).Combos
            If Not ciTemp Is clsTemp Then
                With ciTemp
                    If Not .Textbox Is Nothing Then
                        If TypeOf .Textbox Is Textbox Then
                            If Not .KeyField Or bClearKey Then
                                If .Textbox.Enabled And Not .Textbox.Locked Then
                                    .Textbox.Text = ""
                                End If
                            End If
                        Else
                            .Textbox.Caption = ""
                        End If
                    End If
                End With
            End If
        Next ciTemp
    End If
End Sub

Property Set Table(tblGrid As Object)
    
    Const szROW_STYLE As String = "RowStyle"
    Const szCURRENT_CELL_STYLE As String = "CurrCellStyle"
    
    Dim szBuf As String
    Dim n As Integer
    Dim m_styRow As Object            ' row style
    Dim m_styCurrentCell As Object    ' current cell style
    
    On Error GoTo AddToCollectionError
    
    Set tgDropDown = tblGrid
    
    tblGrid.Top = 1200
    
    SetParent tgDropDown.hwnd, frmParent.hwnd
    
    ' set some properties for the table
    Set m_styRow = tblGrid.Styles.Add(szROW_STYLE)
    Set m_styCurrentCell = tblGrid.Styles.Add(szCURRENT_CELL_STYLE)
    
    ' create current cell style
    With m_styCurrentCell
        .BackColor = &HFF0000
        .ForeColor = &HFFFFFF
    End With
    
    ' create highlight row style
    With m_styRow
        .BackColor = &HFF0000
        .ForeColor = &HFFFFFF
    End With
    
    
    For n = 0 To tblGrid.Columns.Count - 1
        With tblGrid.Columns(n)
            .DividerStyle = dbgBlackLine
            .Locked = False
            .AllowSizing = True
            .Visible = True
            .WrapText = False
            .Button = False
            .FetchStyle = False
        End With
    Next
  
  ' set default properties
   On Error Resume Next
    With tblGrid
        .AllowAddNew = False
        .AllowDelete = False
        .AllowUpdate = False
        .Enabled = True
        .MarqueeUnique = True
        .EditDropDown = False
        .RowDividerStyle = dbgBlackLine
        .BackColor = &HFFFFFF
        .ForeColor = &H0
        .HeadLines = 0
        .Appearance = dbg3D
        .BorderStyle = dbgFixedSingle
        .AllowArrows = True
        .TabAcrossSplits = False
        .WrapCellPointer = False
        .ExposeCellMode = dbgScrollOnSelect
        .TabAction = dbgControlNavigation
        .Splits(0).Locked = True
        .Splits(0).AllowFocus = True
        .Splits(0).AllowSizing = False
        .AllowRowSizing = False
        .ExtendRightColumn = True
        .ScrollBars = dbgAutomatic
        .MarqueeStyle = dbgHighlightRow
        .RecordSelectors = False
        .Visible = False
    End With
    
    ' add the new styles to the table
    tblGrid.AddCellStyle dbgMarqueeRow, m_styRow
    tblGrid.AddCellStyle dbgMarqueeRow + dbgCurrentCell, m_styCurrentCell
    
    On Error GoTo 0
    
    Set m_styRow = Nothing
    Set m_styCurrentCell = Nothing

    Exit Property
    
AddToCollectionError:
    
    On Error GoTo 0

    Set m_styRow = Nothing
    Set m_styCurrentCell = Nothing
End Property

Private Function fnGenetateSearchString(sField As String, sPrefix As String, ByVal nFieldLen As Integer) As String
    'Returns the criteria part of a SQL statement for a numerical fields

    Dim sTemp As String
    Dim sTrimedPrefix As String
    Dim nLenNumber As Integer
    Dim nloop As Integer
    Dim nLoopEnd As Integer
    Dim nxloop As Integer
    Dim nPos As Integer
    
    sTrimedPrefix = CStr(Val(sPrefix))
    nLenNumber = Len(sTrimedPrefix)
    
    'Check whether there is a ':' or '-'
    nPos = InStr(sTrimedPrefix, ":")
    If nPos = 0 Then
        nPos = InStr(sTrimedPrefix, "-")
    End If
    If nPos > 0 Then
        'If there is, limit the field to a range
        If nPos > 1 Then
            sTemp = "(" & sField & " >= " & Left(sTrimedPrefix, nPos - 1)
        Else
            sTemp = ""
        End If
        nPos = nLenNumber - nPos
        If nPos > 0 Then
            If sTemp = "" Then
                sTemp = "(" & sField & " <= " & Right(sTrimedPrefix, nPos)
            Else
                sTemp = sTemp & " AND " & sField & " <= " & Right(sTrimedPrefix, nPos)
            End If
        End If
    Else
        'Otherwise, generate the criteria for concantenate
        sTemp = "(" & sField & " = " & sTrimedPrefix & ")"
        If Val(sTrimedPrefix) > 0 Then
            nLoopEnd = nFieldLen - nLenNumber
            For nloop = 1 To nLoopEnd
                sTemp = sTemp & " OR (" & sField & " >= " & sTrimedPrefix
                For nxloop = 1 To nloop
                     sTemp = sTemp + "0"
                Next
    
                sTemp = sTemp & " AND " & sField & " <= " & sTrimedPrefix
                For nxloop = 1 To nloop
                     sTemp = sTemp & "9"
                Next
    
                sTemp = sTemp & ")"
            Next
        End If
    End If
    fnGenetateSearchString = sTemp
End Function

Public Sub Click(objControl As Object)
    'Event call for all the controls
    If objControl.TabIndex = tgDropDown.TabIndex Then
        If m_bSortOnColumnClick Then
            'david 02/05/2002
            Dim nRowTop As Integer
    
            If tgDropDown.Row >= 0 Then
                nRowTop = tgDropDown.RowTop(0)
    
                If nClickPointY < nRowTop Then
                    bTableClicked = True
                Else
                    If tgDropDown.Row = tgDropDown.RowContaining(nClickPointY) Then
                        bTableClicked = True
                        RowColChange
                    Else
                        'Otherwise, set the flag
                        bTableClicked = True
                    End If
                End If
            Else
                'Otherwise, set the flag
                bTableClicked = True
            End If
        Else
            If tgDropDown.Row = tgDropDown.RowContaining(nClickPointY) Then
                bTableClicked = True
                RowColChange
            Else
                'Otherwise, set the flag
                bTableClicked = True
            End If
        End If
        '''''''''''''''
    Else
        Dim lDropDown As Integer
        Dim nctrlType As Integer
        fnGetKeys lDropDown, lComboIndex, nctrlType, objControl
        
        If lDropDown <= 0 Then
            Exit Sub
        End If
        If nctrlType = CONTROL_TYPE_TEXTBOX Then
            Exit Sub
        End If
    
        'If the control is a command button, load the table
        If colComboControls.Item(lDropDown).ComboOn Then
            If m_bProtoType Then
                If lComboIndex > 0 Then
                    bValidSelection = True
                    On Error Resume Next
                    colComboControls(lDropDown).Combos(lComboIndex).Textbox.SetFocus
                End If
                Exit Sub
            End If
            fnLoadTable lDropDown
        End If
    End If
End Sub

Public Function fnSQLString(ByRef szParameter As String) As String
'
' Properly quotes and formats an SQL string.  If vNoQuotes is present, the result WILL NOT BE QUOTED
' for each ' character found, insert a double ''.  Leave "%* alone
    
    Dim nIdx As Integer
    Dim nPos As Integer
    
    nIdx = 1
    nPos = InStr(nIdx, szParameter, "'")
    
    While nPos <> 0
        szParameter = Left(szParameter, nPos) & "'" & Right(szParameter, Len(szParameter) - nPos)
        nIdx = nPos + 2
        nPos = InStr(nIdx, szParameter, "'")
    Wend
    
    fnSQLString = "'" & szParameter & "'"

End Function

Public Function fnGetRecord(rsTemp As Object, strSQL As String, nDB As Integer)
    ' Get records from the given SQL statement
    ' nDB = 1 ---> Informax Database (remote)
    '     = 2 ---> Access Database (local)

    Select Case nDB
    Case DB_INFORMIX
        'david 09/12/2002  #383118
        'the dbSQLPassThrough parameter causes some of the selected records
        'NOT being shown in the dropdown list (when scrolling down),
        'so remove the dbSQLPassThrough parameter
        'but WITHOUT dbSQLPassThrough parameter, the Query that involves with
        'the Temp table will fail.
        
        On Error Resume Next
        Set rsTemp = dbRemote.OpenRecordset(strSQL, dbOpenSnapshot)
        
        'david 05/16/2007  #5533331
        'STRANGE in papco database the following SQL returns 0 record!
        'SELECT an_phone,an_customer,an_name,an_first_name
        'FROM ar_altname WHERE an_customer = an_cust
        'AND (an_phone LIKE '%000%%') ORDER BY an_phone
        'Added 'Or rsTemp.RecordCount = 0'
        'to try again use dbSQLPassThrough and it works...
        If Err.Number <> 0 Or rsTemp.RecordCount = 0 Then
            'subShowDebugInfo "OpenRecordset WITHOUT dbSQLPassThrough failed. Try next..."
            On Error GoTo SQLError
            Set rsTemp = dbRemote.OpenRecordset(strSQL, dbOpenSnapshot, dbSQLPassThrough)
        End If
        '''''''''''''''''''''''''''
    Case DB_ACCESS
        On Error GoTo SQLError
        Set rsTemp = t_dbLocal.OpenRecordset(strSQL, dbOpenSnapshot)
    Case Else
        fnGetRecord = False
        MsgBox "Drop down database not setup", vbOKOnly + vbCritical
        Exit Function
    End Select
    
    Select Case nDB
    Case DB_INFORMIX
        If rsTemp.RecordCount > 0 Then
            'david 09/12/2002  #383118
            'speed up the dropdown list to show
            'rsTemp.MoveLast
            'rsTemp.MoveNext
            rsTemp.Move 12
            '''''''''''''''''''''''''''
            rsTemp.MoveFirst
            
            'david 09/12/2002  #383118
            'with all the possible work around above (when OpenRecordset),
            'some of the selected records are still NOT being shown in the dropdown list (when scrolling down)
            'THIS IS A WORK AROUND!!!
            If rsTemp.RecordCount = 13 Then
                'subShowDebugInfo "RecordCount = 13"
                rsTemp.MoveLast
                rsTemp.MoveFirst
            End If
            '''''''''''''''''''''''''''
        End If
    Case DB_ACCESS
        If rsTemp.RecordCount > 0 Then
            rsTemp.MoveLast
            rsTemp.MoveFirst
        End If
    End Select
    
    fnGetRecord = True

quitsub:
    On Error GoTo 0
    Exit Function
SQLError:
    subShowODBCError strSQL
    Screen.MousePointer = vbDefault
    fnGetRecord = False
    Resume quitsub
End Function

Private Sub subShowODBCError(Optional vSQL As Variant)
    Dim i As Integer
    Dim sMsgs As String
    Dim sNumbers As String
    Dim sODBCErrors As String
    
    If m_bDebug Then
        Dim strSQL As String
        If IsMissing(vSQL) Then
            strSQL = ""
        Else
            strSQL = vSQL
        End If
    End If
    
    If Err.Number = 3146 Then
        With engDatabase.Errors
            If .Count > 0 Then
                For i = 0 To .Count - 2
                    sMsgs = sMsgs & "Number: " & .Item(i).Number & Space(5) & .Item(i).Description & vbCrLf
                Next
            End If
            If .Count <= 2 Then
                sNumbers = ""
            Else
                sNumbers = "s"
            End If
        End With
        sODBCErrors = "The following error" & sNumbers & " occurred while doing an ODBC query:" & vbCrLf & vbCrLf _
                       & vbCrLf & sMsgs
    Else
        sODBCErrors = Err.Description
    End If

    Dim sMsg As String
    If m_bDebug Then
        sMsg = "SQL: " & strSQL & vbCrLf & vbCrLf & sODBCErrors
        Clipboard.SetText strSQL
    Else
        sMsg = sODBCErrors
    End If
    MsgBox sMsg, vbOKOnly + vbCritical
    Err.Clear

End Sub

Private Sub fnGetText()
    'Populate the text box with the data from the recordset

    Dim i As Integer
    Dim lDropDown As Long
    Dim lColumn As Long
    Dim colColumn As Collection
    Dim colCombo As Collection
    
    On Error GoTo errGetText
    m_bFillingText = True
    If tgDropDown.DataMode = dbgUnbound Then
        Dim lRow As Long
        If Not clsCurrentDropdown Is Nothing Then
            Set colCombo = clsCurrentDropdown.Combos
            lRow = fnCurrentRow
            For i = 1 To colCombo.Count
                With colCombo.Item(i)
                    If colCombo.Item(i).Populate Then
                        If Not .Textbox Is Nothing Then
                            If .ComboType = TYPE_DATE Then
                                If TypeOf .Textbox Is Textbox Then
                                    .Textbox.Text = tfnFormatDate(sGridData(i - 1, lRow))
                                Else
                                    .Textbox.Caption = tfnFormatDate(sGridData(i - 1, lRow))
                                End If
                            Else
                                If TypeOf .Textbox Is Textbox Then
                                    .Textbox.Text = Trim$(sGridData(i - 1, lRow))
                                Else
                                    .Textbox.Caption = Trim$(sGridData(i - 1, lRow))
                                End If
                            End If
                            .FieldValue = .Textbox
                        End If      'If Not .Textbox Is Nothing Then
                        .Changed = False
                        .Valid = True
                    End If   'If colCombo.Item(lCol).Populate Then
                End With
            Next i
    '        tgDropDown.Visible = False
            bValidSelection = True
            clsCurrentDropdown.Cleared = False
            clsCurrentDropdown.DataEntered = False
    '        tgDropDown.ZOrder 1
        End If
    Else
        If Not datLink.Recordset Is Nothing Then
            If Not clsCurrentDropdown Is Nothing Then
                Set colCombo = clsCurrentDropdown.Combos
                For i = 1 To colCombo.Count
                    With colCombo.Item(i)
                        If colCombo.Item(i).Populate Then
                            If Not .Textbox Is Nothing Then
                                'david 09/21/2004
                                If IsNull(datLink.Recordset(fnGetFieldName(.AliasName, True))) Then
                                    If TypeOf .Textbox Is Textbox Then
                                        .Textbox.Text = ""
                                    Else
                                        .Textbox.Caption = ""
                                    End If
                                Else
                                    If .ComboType = TYPE_DATE Then
                                        If TypeOf .Textbox Is Textbox Then
                                            .Textbox.Text = tfnFormatDate(datLink.Recordset(fnGetFieldName(.AliasName, True)))
                                        Else
                                            .Textbox.Caption = tfnFormatDate(datLink.Recordset(fnGetFieldName(.AliasName, True)))
                                        End If
                                    Else
                                        If TypeOf .Textbox Is Textbox Then
                                            .Textbox.Text = Trim$(datLink.Recordset(fnGetFieldName(.AliasName, True)))
                                        Else
                                            .Textbox.Caption = Trim$(datLink.Recordset(fnGetFieldName(.AliasName, True)))
                                        End If
                                    End If
                                End If
                                '''''''''''''''''
                                .FieldValue = .Textbox
                            End If      'If Not .Textbox Is Nothing Then
                            .Changed = False
                            .Valid = True
                        End If   'If colCombo.Item(lCol).Populate Then
                    End With
                Next i
    '            tgDropDown.Visible = False
                bValidSelection = True
                clsCurrentDropdown.Cleared = False
    '            tgDropDown.ZOrder 1
                clsCurrentDropdown.DataEntered = False
            End If
        End If
    End If
    m_bFillingText = False
    On Error GoTo 0
    Exit Sub
errGetText:
    If i <= colCombo.Count Then
        If TypeOf colCombo.Item(i).Textbox Is Textbox Then
            colCombo.Item(i).Textbox.Text = ""
        Else
            colCombo.Item(i).Textbox.Caption = ""
        End If
    End If
    m_bFillingText = False
    Err.Clear
    Resume Next
End Sub
Public Sub GotFocus(objControl As Object)
    
    If objControl Is Nothing Then
        Exit Sub
    End If
    
    If objControl.TabIndex = tgDropDown.TabIndex Then
        Exit Sub
    End If

    Dim lDropDown As Integer
    Dim nctrlType As Integer
    Dim lCombo As Integer

    fnGetKeys lDropDown, lCombo, nctrlType, objControl
    
    If lDropDown < 0 Then
        Exit Sub
    End If
    
    On Error Resume Next
    
    If nctrlType = CONTROL_TYPE_TEXTBOX Then
        
        Dim clsTemp As clsComboSet
        
        Set clsTemp = colComboControls.Item(lDropDown).Combos.Item(lCombo)
        
        tfnRegExpControlChange objControl, clsTemp.RegExpPattern
        
        clsTemp.Textbox.SelStart = 0

        clsTemp.Textbox.SelLength = Len(clsTemp.Textbox.Text)

    End If

    On Error GoTo 0
    
    nCurrentCombo = lDropDown
    
End Sub


Private Function fnGetPCode(ByVal nType As Integer)
    
    'get the PCode for the regular expression
    Dim nCode As Integer
    Dim nWhole As Integer
    Dim nDecimal As Integer
    
    Dim sBuffer As String * 4096
    sBuffer = Space(4096)

    Select Case nType
        Case TYPE_STRING, Is > 0
            If nType < 0 Then
                nCode = GetRegExpPCode(sSTRING_PATTERN, sBuffer)
            Else
                nCode = GetRegExpPCode("^P{0," & CStr(nType) & "}$", sBuffer)
            End If
        Case TYPE_PHONE
             nCode = GetRegExpPCode("^P{0,14}$", sBuffer)
        Case TYPE_LONG
            nCode = GetRegExpPCode(sLONG_PATTERN, sBuffer)
        Case TYPE_INT
            nCode = GetRegExpPCode(sINT_PATTERN, sBuffer)
        Case TYPE_DATE
            nCode = GetRegExpPCode(szDatePattern, sBuffer)
        Case Is < MAX_TYPES
            nWhole = -nType \ TYPE_DCM_INT_MASK
            nDecimal = -nType Mod TYPE_DCM_INT_MASK
            If nDecimal = 0 Then
                nCode = GetRegExpPCode("^#{0," & CStr(nWhole) & "}\.?#{0," & CStr(nWhole) & "}$", sBuffer)
            Else
                nCode = GetRegExpPCode("^#{0," & CStr(nWhole) & "}\.?#{0," & CStr(nDecimal) & "}$", sBuffer)
            End If
        Case Else
            MsgBox "Invalid data type"
    End Select

    If nCode <= 0 Then
        fnGetPCode = ""
    Else
        fnGetPCode = Left(sBuffer, nCode)
    End If

End Function


Public Sub LostFocus(objControl As Object)

    Dim nTemp As Integer
    Dim lDropDown As Integer
    Dim nctrlType As Integer
    Dim strSQL As String
    Dim rsTemp As Recordset
    Dim clsTemp As clsComboSet
    
    If frmParent Is Nothing Then
        Exit Sub
    End If

    If objControl.TabIndex = tgDropDown.TabIndex Then
        If lComboIndex > 0 Then
            On Error Resume Next
            nTemp = frmParent.ActiveControl.TabIndex
            On Error GoTo 0
            If nTemp = tgDropDown.TabIndex Then
                fnSetFocus clsCurrentDropdown.Combos.Item(lComboIndex).Textbox
            End If
        End If
        tgDropDown.Visible = False
'        If nTab <> cmdSearch.TabIndex And nTab <> tgDropDown.TabIndex Then
'            cmdSearch.Visible = False
'            cmdSearch.ZOrder 1
'        End If
'    ElseIf objControl.TabIndex = tgDropDown.TabIndex Then
'    ElseIf objControl.TabIndex = cmdSearch.TabIndex Then
'        If nTab <> tgMain.TabIndex And nTab <> tgDropDown.TabIndex Then
'            cmdSearch.Visible = False
'            tgDropDown.Visible = False
'            cmdSearch.ZOrder 1
'        End If
    Else
        fnGetKeys lDropDown, lComboIndex, nctrlType, objControl
        If nctrlType = CONTROL_TYPE_TEXTBOX Then
            If lDropDown > 0 Then
                For Each clsTemp In colComboControls.Item(lDropDown).Combos
                    If frmParent.ActiveControl Is clsTemp.Textbox Or frmParent.ActiveControl Is clsTemp.Button Then
                        Exit Sub
                    End If
                Next
'                If Validate(objControl) Then
                If Not colComboControls.Item(lDropDown).Cleared Then
                    If colComboControls.Item(lDropDown).ComboOn Then
                        For Each clsTemp In colComboControls.Item(lDropDown).Combos
                            If clsTemp.KeyField Then
                                If clsTemp.Changed Then
                                    Exit Sub
                                End If
                            Else
                                If clsTemp.Populate Then
                                    If Not clsTemp.Textbox Is Nothing Then
                                        m_bFillingText = True
                                        clsTemp.Textbox = clsTemp.FieldValue
                                        DoEvents
                                        m_bFillingText = False
                                    End If
                                End If
                            End If
                        Next
                        colComboControls.Item(lDropDown).Cleared = False
                    End If
                Else
'                    strSQL = fnSQLValidation(lDropDown, False)
'                    If fnGetRecord(rsTemp, strSQL, clsCurrentDropdown.dbType) Then
'                        If rsTemp.RecordCount = 1 Then
'                            For Each clsTemp In colComboControls.Item(lDropDown).Combos
'                                If Not clsTemp.Textbox Is Nothing Then
'                                    clsTemp.Valid = True
'                                    clsTemp.Changed = False
'                                    If IsNull(rsTemp.Fields(clsTemp.AliasName)) Then
'                                        clsTemp.Textbox = ""
'                                        clsTemp.FieldValue = ""
'                                    Else
'                                        clsTemp.FieldValue = Trim(rsTemp.Fields(clsTemp.AliasName))
'                                        clsTemp.Textbox = clsTemp.FieldValue
'                                    End If
'                                End If
'                            Next
'                            colComboControls.Item(lDropDown).Cleared = False
'                        End If
'                    End If
                End If
            End If
        End If 'If nctrlType = CONTROL_TYPE_TEXTBOX Then
    End If
End Sub
Public Sub MouseUp(y As Single)
    nClickPointY = y
End Sub
Public Sub TableMouseUp(y As Single)
    nClickPointY = y
End Sub

Private Function fnCurrentRow() As Long
    If IsNull(tgDropDown.FirstRow) Then
        fnCurrentRow = tgDropDown.Row
    Else
        fnCurrentRow = Val(tgDropDown.FirstRow) + tgDropDown.Row
    End If
End Function

Public Sub RowColChange()
    Dim lCol As Long
    
    'david 01/28/2002
    If bHeaderClick Then
        bHeaderClick = False
        bTableClicked = False
        Exit Sub
    End If
    ''''''''''''''''''
    
    If bTableClicked Then
        fnGetText
        fnSetFocus clsCurrentDropdown.Combos.Item(lComboIndex).Textbox
        tgDropDown.Visible = False
        tgDropDown.ZOrder 1
        bTableClicked = False
    Else
        If tgDropDown.Row < 0 Then
            tgDropDown.Row = nLastRow
        End If
        nLastRow = tgDropDown.Row
    End If
End Sub
Property Let SQL(SQL As String)
    
    If colComboControls Is Nothing Then
        Set colComboControls = New Collection
    End If
    If colComboControls.Count = 0 Then
        Dim clsDD As New clsComboDropDown
        colComboControls.Add clsDD
    End If
    
    colComboControls.Item(nCurrentCombo).SQL = SQL
    
    fnSetupAlias nCurrentCombo
End Property

Property Get SQL() As String
    If colComboControls Is Nothing Then
        Exit Property
    End If
    If colComboControls.Count = 0 Then
        Exit Property
    End If
    
    SQL = colComboControls.Item(nCurrentCombo).SQL
End Property




Public Function SQL_STRING_TYPE(Optional vLen As Variant) As Integer
    If IsMissing(vLen) Then
        SQL_STRING_TYPE = TYPE_STRING
    Else
        Dim nLen As Integer
        nLen = CInt(vLen)
        If nLen > 0 Then
            SQL_STRING_TYPE = nLen
        Else
            SQL_STRING_TYPE = TYPE_STRING
        End If
    End If
End Function

Private Sub fnTableClick(tgTable As Object)
    If tgTable.TabIndex = tgDropDown.TabIndex Then
    End If
End Sub

Property Get SingleRecordSelected() As Boolean
    SingleRecordSelected = bValidSelection
    bValidSelection = False
End Property
Property Let SingleRecordSelected(bFlag As Boolean)
    bValidSelection = bFlag
End Property

Property Let ComboOn(objControl As Object, bStatus As Boolean)
    Dim lDropDown As Integer
    Dim nctrlType As Integer
    Dim clsTemp As clsComboDropDown

    fnGetKeys lDropDown, lComboIndex, nctrlType, objControl
    If lDropDown < 0 Then
        Exit Property
    End If
    
    Set clsTemp = colComboControls.Item(lDropDown)
    clsTemp.ComboOn = bStatus

End Property

Property Let ReQuery(objControl As Object, bStatus As Boolean)
    Dim lDropDown As Integer
    Dim nctrlType As Integer
    Dim clsTemp As clsComboDropDown

    fnGetKeys lDropDown, lComboIndex, nctrlType, objControl
    If lDropDown < 0 Then
        Exit Property
    End If
    
    Set clsTemp = colComboControls.Item(lDropDown)
    clsTemp.ReQuery = bStatus

End Property


Property Let ComboActive(vTemp As Variant, bStatus As Boolean)
    Dim lDropDown As Integer
    Dim nctrlType As Integer
    Dim objControl As Object
    Dim clsTemp As clsComboDropDown

    If IsObject(vTemp) Then
        Set objControl = vTemp
        fnGetKeys lDropDown, lComboIndex, nctrlType, objControl, True
        If lDropDown < 0 Then
            Exit Property
        End If
    ElseIf IsNumeric(vTemp) Then
        lDropDown = Val(vTemp)
    End If
    If lDropDown > 0 And lDropDown <= colComboControls.Count Then
        Set clsTemp = colComboControls.Item(lDropDown)
        clsTemp.Active = bStatus
    End If
End Property

Property Get ComboActive(vTemp As Variant) As Boolean
    
    Dim lDropDown As Integer
    Dim nctrlType As Integer
    Dim objControl As Object
    Dim clsTemp As clsComboDropDown

    ComboActive = True
    If IsObject(vTemp) Then
        Set objControl = vTemp
        fnGetKeys lDropDown, lComboIndex, nctrlType, objControl, True
        If lDropDown < 0 Then
            Exit Property
        End If
    ElseIf IsNumeric(vTemp) Then
        lDropDown = Val(vTemp)
    End If
    If lDropDown > 0 And lDropDown <= colComboControls.Count Then
        ComboActive = colComboControls.Item(lDropDown).Active
    End If
End Property

Property Get ColVisible(vTemp As Variant) As Boolean
    
    Dim lDropDown As Integer
    Dim nctrlType As Integer
    Dim objControl As Object
    Dim clsTemp As clsComboDropDown
    Dim colTemp As Collection
    
    ColVisible = False
    If IsObject(vTemp) Then
        Set objControl = vTemp
        fnGetKeys lDropDown, lComboIndex, nctrlType, objControl, False
        If lDropDown < 0 Then
            Exit Property
        End If
        If lDropDown > 0 And lDropDown <= colComboControls.Count Then
            Set colTemp = colComboControls.Item(lDropDown).Combos
            If lComboIndex > 0 And lComboIndex <= colTemp.Count Then
                ColVisible = colTemp(lComboIndex).Visible
            End If
        End If
    End If
End Property

Property Let ColVisible(vTemp As Variant, bTemp As Boolean)
    
    Dim lDropDown As Integer
    Dim nctrlType As Integer
    Dim objControl As Object
    Dim clsTemp As clsComboDropDown
    Dim colTemp As Collection
    
    If IsObject(vTemp) Then
        Set objControl = vTemp
        fnGetKeys lDropDown, lComboIndex, nctrlType, objControl, False
        If lDropDown < 0 Then
            Exit Property
        End If
        If lDropDown > 0 And lDropDown <= colComboControls.Count Then
            Set colTemp = colComboControls.Item(lDropDown).Combos
            If lComboIndex > 0 And lComboIndex <= colTemp.Count Then
                colTemp(lComboIndex).Visible = bTemp
            End If
        End If
    End If
End Property

Property Get ComboOn(objControl As Object) As Boolean
    Dim lDropDown As Integer
    Dim nctrlType As Integer
    Dim clsTemp As clsComboDropDown

    fnGetKeys lDropDown, lComboIndex, nctrlType, objControl
    If lDropDown < 0 Then
        Exit Property
    End If
    
    Set clsTemp = colComboControls.Item(lDropDown)
    ComboOn = clsTemp.ComboOn

End Property

Property Let Validate(txtBox As Object, bFlag As Boolean)
    Dim lDropDown As Integer
    Dim i As Integer
    Dim nCombo As Integer
    Dim clsTemp As clsComboSet
    
    If Not txtBox Is Nothing Then
        fnGetKeys lDropDown, nCombo, i, txtBox
        If lDropDown > 0 Then
            'Get the dropdown class
            For Each clsTemp In colComboControls.Item(lDropDown).Combos
                clsTemp.Valid = bFlag
                clsTemp.Changed = False
            Next
        End If
    End If
End Property

Property Get Validate(txtBox As Object) As Boolean
    'Validate against the SQL for the dropdown
    
    Dim szQuery As String
    Dim rsTemp As Recordset
    Dim lDropDown As Integer
    Dim i As Integer
    Dim nCombo As Integer
    Dim clsTemp As clsComboSet
    
    If m_bProtoType Then
        Validate = True
        Exit Property
    End If
    If txtBox Is Nothing Then
        Validate = True
    Else
        fnGetKeys lDropDown, nCombo, i, txtBox
        If lDropDown > 0 Then
            If nCombo > 0 Then
                Set clsTemp = colComboControls.Item(lDropDown).Combos.Item(nCombo)
                If Not clsTemp.Changed Then
                    Validate = clsTemp.Valid
                    Screen.MousePointer = vbDefault
                    Exit Property
                End If
            End If
            If Trim(txtBox) = "" Then
                Exit Property
            End If
            
            Screen.MousePointer = vbHourglass
            szQuery = fnSQLValidation(lDropDown, True)
            If szQuery = "" Then
                Screen.MousePointer = vbDefault
                Exit Property
            End If
            'Get the records
            If Not fnGetRecord(rsTemp, szQuery, clsCurrentDropdown.dbType) Then
                Screen.MousePointer = vbDefault
                Exit Property
            End If
        
            'Check there is something for us to display before we show the results
            If rsTemp.RecordCount = 1 Then
                Validate = True
                For Each clsTemp In colComboControls.Item(lDropDown).Combos
                    If clsTemp.Populate Then
                        clsTemp.Valid = True
                        clsTemp.Changed = False
                        If IsNull(rsTemp.Fields(clsTemp.AliasName)) Then
                            clsTemp.FieldValue = ""
                        Else
                            clsTemp.FieldValue = Trim(rsTemp.Fields(clsTemp.AliasName))
                        End If
                    End If
                Next
                colComboControls.Item(lDropDown).Cleared = False
            Else
                Validate = False
                For Each clsTemp In colComboControls.Item(lDropDown).Combos
                    clsTemp.Valid = False
                    clsTemp.Changed = False
                Next
            End If
            
            'Reset our mouse pointer to the default
            Screen.MousePointer = vbDefault
        Else
            Validate = True
        End If   'If nDropDown > 0 Then
    End If
End Property

Property Let VisibleRows(nRows As Integer)
    
    colComboControls.Item(nCurrentCombo).VisibleRows = nRows

End Property

Property Get VisibleRows() As Integer
    
    VisibleRows = colComboControls.Item(nCurrentCombo).VisibleRows

End Property

'david 10/30/2001
Property Let CaseSensitive(bStatus As Boolean)
     m_bCaseSensitive = bStatus
End Property

Property Get CaseSensitive() As Boolean
     CaseSensitive = m_bCaseSensitive
End Property

Property Let SortOnColumnClick(bStatus As Boolean)
    m_bSortOnColumnClick = bStatus
End Property

Property Get SortOnColumnClick() As Boolean
     SortOnColumnClick = m_bSortOnColumnClick
End Property

Private Sub Class_Initialize()
    Set colComboControls = New Collection
    sPercentOrAster(0) = "%"
    sPercentOrAster(1) = "*"
    nCurrentCombo = 1
    lRecordsReturned = -1
    m_bFillingText = False
    m_bProtoType = False
    m_bCaseSensitive = False
    lHeadClickDropDown = -1
    m_bSortOnColumnClick = False
End Sub

Private Sub Class_Terminate()

    Dim colTemp As clsComboDropDown
    'Clear all the contents in the collections
    While colComboControls.Count > 0
        While colComboControls.Item(1).Combos.Count > 0
            Set colComboControls.Item(1).Combos.Item(1).Textbox = Nothing
            Set colComboControls.Item(1).Combos.Item(1).Button = Nothing
            colComboControls.Item(1).Combos.Remove 1
        Wend
        Set colComboControls.Item(1).Combos = Nothing
        colComboControls.Remove 1
    Wend
    Set colComboControls = Nothing
    Set clsCurrentDropdown = Nothing
    Set tgDropDown = Nothing
    Set frmParent = Nothing
    Set datLink = Nothing
    Set dbRemote = Nothing
    Set t_dbLocal = Nothing
    Set engDatabase = Nothing
    Set colTemp = Nothing
End Sub

'david 01/28/2002
Public Sub HeadClick(ByVal ColIndex As Integer)
    If Not m_bSortOnColumnClick Then
        Exit Sub
    End If
    
    If lHeadClickDropDown < 0 Then
        Exit Sub
    End If
    
    fnLoadTable lHeadClickDropDown, True, ColIndex
End Sub

'david 01/28/2002
Private Function fnFixFirstSortOrder(ByVal sOrder As String, bDescending As Boolean) As String
    Dim nPosi As Integer
    Dim sTemp As String
    Dim sOthers As String
    If sOrder = "" Then
        Exit Function
    End If
    
    If Trim(sOrder) = "" Then
        Exit Function
    End If
    
    nPosi = InStr(sOrder, ",")
    
    If nPosi > 0 Then
        sTemp = Trim(Left(sOrder, nPosi - 1))
        sOthers = Mid(sOrder, nPosi)
    Else
        sTemp = Trim(sOrder)
        sOthers = ""
    End If

    nPosi = InStr(sTemp, " ")
    
    If nPosi > 0 Then
        sTemp = Trim(Left(sTemp, nPosi - 1))
    End If
    
    If bDescending Then
        sTemp = sTemp + " DESC"
    End If

    fnFixFirstSortOrder = sTemp + sOthers
End Function

Private Function fnCnvtOrderByFieldsToColIndex(strSQL As String, sOrderByList As String) As String
    Const szUNION As String = " UNION "
    Const szSELECT As String = "SELECT "
    Const szFROM As String = " FROM "
    
    Dim arySQL() As String
    Dim aryField() As String
    Dim aryOrderBy() As String
    Dim aryByColIndex() As String
    Dim nCountOrderBy As Integer
    
    Dim sField As String
    Dim sOrderBy As String
    
    Dim sTemp As String
    Dim nPosi1 As Integer
    Dim nPosi2 As Integer
    
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    
    On Error GoTo errTrap
    
    If InStr(strSQL, szUNION) <= 0 Then
        fnCnvtOrderByFieldsToColIndex = sOrderByList
        Exit Function
    End If
    
    arySQL() = Split(strSQL, szUNION)
    aryOrderBy() = Split(sOrderByList, ",")
    
    nCountOrderBy = UBound(aryOrderBy)
    ReDim aryByColIndex(nCountOrderBy)
    
    For i = 0 To UBound(arySQL)
        'extract fields
        nPosi1 = InStr(UCase(arySQL(i)), szSELECT)
        nPosi2 = InStr(UCase(arySQL(i)), szFROM)
        
        If nPosi1 > 0 And nPosi2 > 0 And nPosi1 < nPosi2 Then
            sTemp = Trim(Mid(arySQL(i), nPosi1 + Len(szSELECT), nPosi2 - (Len(szSELECT) + nPosi1)))
            aryField() = Split(sTemp, ",")
            
            For j = 0 To UBound(aryOrderBy)
                sOrderBy = UCase(Trim(aryOrderBy(j)))
                'check asc/desc
                sTemp = ""
                nPosi1 = InStr(sOrderBy, " ")
                
                If nPosi1 > 0 Then
                    sTemp = Trim(Mid(sOrderBy, nPosi1 + 1))
                    sOrderBy = Left(sOrderBy, nPosi1 - 1)
                End If
                
                If aryByColIndex(j) = "" Then
                    If IsNumeric(sOrderBy) Then
                        aryByColIndex(j) = sOrderBy & Trim(" " + sTemp)
                        nCountOrderBy = nCountOrderBy - 1
                    Else
                        For k = 0 To UBound(aryField)
                            sField = UCase(Trim(aryField(k)))
                            
                            If InStr(sField, sOrderBy) > 0 Then
                                aryByColIndex(j) = (k + 1) & Trim(" " + sTemp)
                                nCountOrderBy = nCountOrderBy - 1
                            End If
                            
                            If nCountOrderBy < 0 Then
                                Exit For
                            End If
                        Next k
                    End If
                End If
                
                If nCountOrderBy < 0 Then
                    Exit For
                End If
            Next j
        End If
    
        If nCountOrderBy < 0 Then
            Exit For
        End If
    Next i
    
    sTemp = ""
    
    For j = 0 To UBound(aryByColIndex)
        If aryByColIndex(j) <> 0 Then
            If sTemp <> "" Then
                sTemp = sTemp + ", "
            End If
            
            sTemp = sTemp + aryByColIndex(j)
        End If
    Next j
    
    fnCnvtOrderByFieldsToColIndex = sTemp
    
    Exit Function
    
errTrap:
    'error, return empty order by
    fnCnvtOrderByFieldsToColIndex = ""
End Function

'##############################################################################
' subShowDebugInfo:
' Author: David Chai
' Date: 09/13/2002
' Project Number: N/A
' Program Version: N/A
' ARGS:
'   sMsg: string, message to show on the main form's status bar
' Returns:
'   none
' Description:
'   show the message on the main form' status bar in red color.
'   NOTE: THIS FUNCTION SHOULD BE CALLED FOR DEBUGGING PURPOSELY!!!
'-
'##############################################################################
Public Sub subShowDebugInfo(sMsg As String)
    Dim lColor As Long
    Dim bBold As Boolean
    Dim sStatusMsg As String
    
    On Error Resume Next
    
    'save the old status bar message
    With frmParent
        
        lColor = .ffraStatusbar.ForeColor
        bBold = .ffraStatusbar.Font.Bold
        sStatusMsg = .ffraStatusbar
        
        .ffraStatusbar.ForeColor = &HFF&
        .ffraStatusbar.Font.Bold = True
        .ffraStatusbar = sMsg
        .ffraStatusbar.Refresh
        
        'pause for 1 second...
        Dim sngTimer As Single
        sngTimer = Timer
        While Timer - sngTimer < 0.5
            'do nothing
        Wend
        
        'restore the old status bar message
        .ffraStatusbar.ForeColor = lColor
        .ffraStatusbar.Font.Bold = bBold
        .ffraStatusbar = sStatusMsg
        .ffraStatusbar.Refresh
    End With
End Sub



