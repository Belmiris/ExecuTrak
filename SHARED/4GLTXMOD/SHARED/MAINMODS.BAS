Attribute VB_Name = "ModMainmods"
'***********************************************************************************
'
' Copyright (c) 1996 FACTOR, A Division of W.R.Hess Company
'
' Module name   : MAINMODS.BAS
' Date          : Mar 12, 1996
' Programmer(s) :
'
' This module is a 4GL-to-BAS translation of:
' @(#) mainmods!4gl 1!39@(#)    Last Delta:5/1/95
' Copyright (c) 1988,89,90,91 FACTOR, A Division of WR Hess Company
'   note: original 4GL comments (extremely rare!) have been preserved with consecutive single quotes ('')
'
' Public Functions & Subs:
'
'                   abs_value
'                   data_ok
'                   db_path
'                   dround
'                   dserror                 dsmsg
'                   end_of_program
'                   fastinit
'                   fsysinit
'                   get_company_name
'                   gl_period_ok
'                   lderror
'                   main_error              main_warning
'                   makef
'                   mst_pc
'                   next_program
'                   packit                  unpackit
'                   prep_logn               prep_next           prep_parent
'                   print_scrn
'                   process_locked          unlock_process
'                   pserror
'                   psmsg
'                   rlocation_rsm
'                   rprft_name
'                   roundf
'                   set_mst_master          set_mst_pc
'                   set_term_units
'                   sql_execute
'                   sverrors
'                   term_units
'                   top_bar
'                   top_company
'                   translate_text
'                   trunct
'                   start_of_program
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                   fsys_parm           - moved to REFNFAST.BAS
'                   fsys_prft_ctr       -   "    "      "
'                   stdconv             -   "    "      "
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Private:          fnMainmodsExecute
'                   fnMainmodsOpenRecordset
'
'***********************************************************************************

Option Base 1
Option Explicit

#Const cDEBUG = 0
#Const DEVELOPMENT = 0
Const conErrmsgDimension = 20
 

''  **** NOTE ****
'' Uncomment the function acdealloc at the end of this program
'' for compilation under Informix 6!0 4GL or higher! Earlier
'' versions of 4GL had this function built in to deallocate
'' string space! It is not necessary for 6!0, but we need to
'' support both old and new versions for a time! When all customers
'' are at 6!0 or higher, the call may be removed FROM all 4gl's!
'' TH 05/01/95
'' database(factor)

Private pr_company_con_name As String 'As Recordset
Private pa_serr_active(conErrmsgDimension) As Integer
Private pr_errmsg As Recordset

Private mstpc As Variant 'Integer
Private termsunits As Variant 'Double
Private init_mods As Boolean 'Long
Private init_fast As Boolean 'Long
Private init_fsys As Boolean 'Long
Private init_wndw As Boolean 'Long

Public bDontShowError As Boolean

' the following are included here, until a decision is made regarding how these
' variables will be "visible" FROM the template:
'Dim t_szConnect As String            'holds ODBC connect string
'Dim t_engFactor As DBEngine          'pointer to database engine
'Dim t_wsWorkSpace As Workspace       'pointer to the default workspace
'Dim t_dbMainDatabase As DataBase     'main database handle


''  $ start_of_program    The startup caller function
'' jrs 4-17-87
'' jrs 4-26-87
'' function start   main
'' This function is called FROM every seperate module to handle
'' the program loads in a uniform manner! It does the following:
'' calls lderror() - load error messages
'' loads pr_company - The company name for the standard screen
'' CHANGE LOG:
'' 04-17-87 jrs added the security check
'' 04-26-87 jrs allow "logn" to allways pass security
'' 05-01-87 JRS added dsmsg(), psmsg()

Public Function fnMainmodsExecute(strSQL As String, _
                             Optional vCaller As Variant, _
                             Optional vDB As Variant) As Boolean
    Const SUB_NAME = "fnMainmodsExecute"
    Dim objDB As Database
    
    If IsMissing(vDB) Then
        Set objDB = t_dbMainDatabase
    Else
        Set objDB = vDB
    End If
    On Error GoTo errExecute
    If objDB Is t_dbMainDatabase Then
        objDB.ExecuteSQL strSQL
    Else
        objDB.Execute strSQL
    End If
    
    fnMainmodsExecute = True

    On Error GoTo 0
    Exit Function

errExecute:
    fnMainmodsExecute = False
    If IsMissing(vCaller) Then
        tfnErrHandler SUB_NAME, strSQL, Not bDontShowError
    Else
        tfnErrHandler SUB_NAME & "," & CStr(vCaller), strSQL, Not bDontShowError
    End If
End Function


Public Function fnMainmodsOpenRecordset(strSQL As String, _
                              Optional vCaller As Variant, _
                              Optional vDB As Variant) As Recordset
    Const SUB_NAME = "fnMainmodsOpenRecordset"
    ' Get records from the given SQL statement
    Dim objDB As Database
    Dim rsTemp As Recordset

    If IsMissing(vDB) Then
        Set objDB = t_dbMainDatabase
    Else
        Set objDB = vDB
    End If
    On Error GoTo SQLError
    If objDB Is t_dbMainDatabase Then
        Set rsTemp = objDB.OpenRecordset(strSQL, dbOpenSnapshot, dbSQLPassThrough)
    Else
        Set rsTemp = objDB.OpenRecordset(strSQL, dbOpenSnapshot)
    End If
    If rsTemp.RecordCount > 0 Then
        rsTemp.MoveLast
        rsTemp.MoveFirst
    End If
    Set fnMainmodsOpenRecordset = rsTemp

    On Error GoTo 0
    Exit Function
SQLError:
    Set fnMainmodsOpenRecordset = Nothing
    If IsMissing(vCaller) Then
        tfnErrHandler SUB_NAME, strSQL, Not bDontShowError
    Else
        tfnErrHandler SUB_NAME & "," & CStr(vCaller), strSQL, Not bDontShowError
    End If
End Function

Public Function start_of_program(un As String)
    Const SUB_NAME = "start_of_program"
    
    Dim tstit As String, error_name As String, strt_msg As String
    Dim rs As Recordset
    
    mstpc = Null
    termsunits = Null
    Call lderror
    '' declare c_co_comp cursor for:    '' open c_co_comp:  '' fetch c_co_comp into pr_company
    Set rs = fnMainmodsOpenRecordset("SELECT * FROM co_company_name", SUB_NAME)
    If rs Is Nothing Then
        pr_company_con_name = "NO COMPANY NAME FOUND"
    Else
        pr_company_con_name = rs!con_name
        Set rs = Nothing
    End If
    '' close c_co_comp
    
    tstit = Environ("FUNIT")
    If Len(tstit) = 0 Then tstit = "NULL"
    '' comment out this if statement if rds
    If un <> "logn" Then
        If tstit <> un Then
            If Not bDontShowError Then
                'Beep
                MsgBox "You just FAILED your security check", vbCritical, "start_of_program"
                End 'Exit Program(2734)
            End If
        End If
    End If
    
    error_name = "/dev/null"
    If un <> "logn" Then
        '' declare c_error_name cursor for
        Dim c_error_name As String
        c_error_name = _
                            " SELECT parm_field FROM sys_parm" & _
                            " WHERE parm_nbr = 98"
        '' open c_error_name    '' fetch c_error_name into error_name
        Set rs = fnMainmodsOpenRecordset(c_error_name, SUB_NAME)
        If Not (rs Is Nothing) Then error_name = rs!parm_field
        '' close c_error_name
    End If
    Call startlog(error_name)
    
    '' let strt_msg = "Starting ",un, " ", Environ("LOGNAME")
    '' call errorlog(strt_msg)
    '' options help file 'fullhelp' ''
    
    init_mods = False
    init_fast = False
    init_fsys = False
    init_wndw = False
    CleanUp rs
    ''set lock mode to wait
End Function    ''start_of_program


Public Sub end_of_program()
    Call sverrors
End Sub ''end_of_program

'' uncomment this function if rds
'' function acdealloc()
'' end function

Public Function fastinit() As Boolean
    If init_fast Then
        fastinit = False: Exit Function
    End If
    init_fast = True
    fastinit = True
End Function ''fastinit


Public Function fsysinit() As Boolean
    If init_fsys Then
        fsysinit = False: Exit Function
    End If
    init_fsys = True
    fsysinit = True
End Function ''fsysinit


Public Function db_path() As String
    Dim r_path As String
    Dim r_pos As Long
    r_path = Environ("DBPATH")
    r_pos = InStr(r_path, ":")
    If r_pos = 0 Then r_pos = Len(r_path) + 1
    r_pos = r_pos - 1
    db_path = Mid(r_path, 1, r_pos)
End Function ''db_path


Public Sub top_bar(f As String, k As String)
#If False Then 'probably will not be used
    Dim i As String * 1
    Dim lLen As Long, strt As Long
    
    j = 79 '' spaces char*
    j(66, 78) = "Help = Ctrl-w"
    j(2, 9) = f
    lLen = length(k)
    strt = (79 - lLen) / 2
    j(strt, strt + lLen) = k ' /* clipped */
    '' comment out this next call for rds
    Call cubox("l", 4, 1, 22, 79)
    '' display j at 3,1 attribute(reverse)
    Call top_company("")
#End If
End Sub         ''top_bar


Public Sub top_company(k As String)
#If False Then 'probably will not be used
    Dim lLen As Long, strt As Long
    '' display "" at 1,1
    '' display "" at 2,1
    lLen = length(pr_company_con_name)
    strt = (80 - lLen) / 2
    '' display pr_company_con_name at 1,strt
    If k = Null Then
            k = Date
    End If
    lLen = length(k)
    strt = (80 - lLen) / 2
    '' display k at 2,strt
#End If
End Sub         ''top_company

Public Function sql_execute(sql_str As String) As Boolean
    Const SUB_NAME = "sql_execute"
    
    If fnMainmodsExecute(sql_str, SUB_NAME) Then
        sql_execute = True  'show that an error did occur
    Else
        sql_execute = False 'show that no error will occur ????
    End If
End Function    ''sql_execute

Public Sub lderror()
    Const SUB_NAME = "lderror"
    
    Dim i As Integer
    Dim cut_off As Double
    Dim rs As Recordset
    
    '' SELECT avg(serr_active) into cut_off FROM sys_errors
    Set rs = fnMainmodsOpenRecordset("SELECT avg(serr_active) dAvg FROM sys_errors", SUB_NAME)
    If Not (rs Is Nothing) Then cut_off = rs!dAvg
    
    Set pr_errmsg = fnMainmodsOpenRecordset("SELECT * FROM sys_errors WHERE serr_active > " & QStr(cut_off * 1.1), SUB_NAME)

    For i = LBound(pa_serr_active) To conErrmsgDimension
        pa_serr_active(i) = 0   'set current error-counting array to zeroes
    Next
    CleanUp rs
End Sub ''lderror

Public Function pserror(i As Long)
        Call dserror(i)
        '' sleep(5)
End Function    ''pserror

Public Function psmsg(i As Long)
        Call dsmsg(i)
        '' sleep(5)
End Function    ''psmsg

Private Function fnGetMessage(i As Long) As Recordset
    Const SUB_NAME = "fnGetMessage"
    
    Dim j As Integer, k As Integer
    
    If pr_errmsg Is Nothing Then
        Exit Function
    End If
    With pr_errmsg
        If .RecordCount > 0 Then
            .MoveFirst
            j = LBound(pa_serr_active)
            While Not .EOF
                If !serr_nbr = i Then
                    pa_serr_active(j) = pa_serr_active(j) + 1
                    Set fnGetMessage = pr_errmsg    'current record
                    Exit Function
                End If
                .MoveNext
                j = j + 1
            Wend
        End If
    End With
    
    'first check whether this error number exists
    Set fnGetMessage = fnMainmodsOpenRecordset("SELECT * FROM sys_errors WHERE serr_nbr = " & QStr(i), SUB_NAME)
    If Not (fnGetMessage Is Nothing) Then
        Call fnMainmodsExecute("UPDATE sys_errors SET serr_active = serr_active + 1 WHERE serr_nbr = " & QStr(i), SUB_NAME)
    End If
End Function    'fnGetMessage

Public Sub dserror(i As Long)
    Dim errmsg As String
    Dim rs_errmsg As Recordset
    
    Set rs_errmsg = fnGetMessage(i)
    If rs_errmsg Is Nothing Then
        errmsg = "Error message " & i & " is not defined"
    Else
        errmsg = rs_errmsg!serr_mesage
    End If
    '' options      ''
    '' message line 24,     ''              '' error line 24,       ''
    '' comment line 23
    '' error errmsg
    #If DEVELOPMENT Then
        If Not bDontShowError Then
            MsgBox errmsg, vbExclamation, "sub dserror (development only message?)"
        End If
    #End If
    CleanUp rs_errmsg
End Sub         ''dserror

Public Sub dsmsg(i As Long)
    Dim errmsg As String
    Dim rs_errmsg As Recordset
    
    Set rs_errmsg = fnGetMessage(i)
    If rs_errmsg Is Nothing Then
        errmsg = "Message " & i & " is not defined"
    Else
        errmsg = rs_errmsg!serr_mesage
    End If
    '' options
    '' message line 24,     ''              '' error line 24,       ''
    '' comment line 23
    '' error errmsg
    #If DEVELOPMENT Then
        If Not bDontShowError Then
            MsgBox errmsg, vbExclamation, "sub dsmsg (development only message?)"
        End If
    #End If
    CleanUp rs_errmsg
End Sub         ''dsmsg

Public Sub sverrors()
    Const SUB_NAME = "sverrors"
    
    Dim j As Integer
    Dim sSql As String
    
    If pr_errmsg Is Nothing Then
        Exit Sub
    End If
    With pr_errmsg
        If .RecordCount > 0 Then
            .MoveFirst
            j = LBound(pa_serr_active)
            While Not .EOF
                If pa_serr_active(j) > 0 Then
                    sSql = "UPDATE sys_errors" & _
                           " SET serr_active = serr_active + " & QStr(pa_serr_active(j)) & _
                           " WHERE serr_nbr = " & QStr(!serr_nbr)
                    Call fnMainmodsExecute(sSql, SUB_NAME)
                End If
                .MoveNext
                j = j + 1
            Wend
        End If
    End With
End Sub ''sverrors

''  $ data_ok()  common function to ask operator if data is correct
'' rwr 4-20-87
'' rwr 4-20-87
'' function start
'' This function is called any time you need the operator to validate
'' the data they have entered!  Destroys the top 2 lines on the screen!
'' returns true if data is ok
'' false if data is not ok
'' calls print_scrn() - the screen print utility
'' CHANGE LOG:

Public Function data_ok() As Boolean
    Dim dok As Long, answer As Long
    '' options help file 'fullhelp'
    '' options help key control-w
    If Not bDontShowError Then
        Do While True
            answer = MsgBox("DATA OK?", vbYesNo)
            Select Case answer
                Case vbYes: '' command( "Yes" ) '' "Accept this data" help 29010
                    dok = True
                    Exit Do 'while
                Case vbNo:  '' command( "No" )  '' "Change this data" help 29010
                    dok = False
                    Exit Do 'while
                Case Else
                    '' command( "Print" )
                    '' "Print this screen on hardcopy" help 29010
                    '' Call print_scrn
                    '' command( key(interrupt) )
                    '' dok = False
                    '' Exit Menu
            End Select
        Loop 'While menu
        data_ok = dok
    Else
        data_ok = vbYes
    End If

End Function ''data_ok

''  $ print_scrn() common function print the screen to a printer
'' jrs 4-20-87
'' jrs 4-20-87
'' function print screen
'' This function is called any time you need to print the screen
'' returns nothing
'' CHANGE LOG:

Public Sub print_scrn()
    'run "prntscrn"
End Sub ''print_scrn

''  $ packit  Pack 5 Different alpha fields into 1
'' jrs 04-13-87
'' jrs 04-13-87
'' process name address
'' This process packs the last name, first name, attention line, and
'' two address lines together to form a single alpha field!
'' Parameters Passed: last name          CHAR(30)
'' first name         CHAR(30)
'' attention line     CHAR(30)
'' address line 1     CHAR(30)
'' address line 2     CHAR(30)
'' Parameters Return: 1 alpha field CHAR(90)
'' Error Conditions: Input field contains a ^
'' CHANGE LOG:

Public Function packit(ln As String, fn As String, al As String, _
                        a1 As String, a2 As String) As String
    Dim line_out As String
    Dim i As Long, j As Long, k As Long
    line_out = RTrim(ln) & " ^" & RTrim(fn) & " ^"
    i = Len(line_out)
    j = Len(al)
    k = i + j
    If k < 84 Then
            line_out = RTrim(line_out) & RTrim(al) & " ^"
    Else
            i = 84 - i
            line_out = RTrim(line_out) & Mid(al, 1, i) & " ^ ^ ^"
            packit = line_out: Exit Function
    End If
    i = Len(line_out)
    j = Len(a1)
    k = i + j
    If k < 86 Then
            line_out = RTrim(line_out) & RTrim(a1) & " ^"
    Else
            i = 86 - i
            line_out = RTrim(line_out) & Mid(a1, 1, i) & " ^ ^"
            packit = line_out: Exit Function
    End If
    i = Len(line_out)
    j = Len(a2)
    k = i + j
    If k < 88 Then
            line_out = RTrim(line_out) & RTrim(a2) & " ^"
    Else
            i = 88 - i
            line_out = RTrim(line_out) & Mid(a2, 1, i) & " ^"
            packit = line_out: Exit Function
    End If
    packit = line_out
End Function    ''packit

''  $ unpackit    The unpacking of a packed alpha field
'' jrs 04-13-87
'' jrs 04-13-87
'' process name address
'' This process breaks a packed routine into the five associated
'' parts!
'' Passed parameter: alpha field       CHAR(90)
'' Returned parameters: last name          CHAR(30)
'' first name         CHAR(30)
'' attention line     CHAR(30)
'' address line 1     CHAR(30)
'' address line 2     CHAR(30)
'' Error Conditions: passed alpha not properly packed
'' returns a2 filled with @
'' CHANGE LOG:

Public Sub unpackit(line_in As String, fn() As String) 'array "fn" is output
        
    Dim line_part As String
    Dim pos As Integer, i As Integer
    
    line_part = line_in
    i = LBound(fn)
    While True
        pos = (InStr(line_part, "^") - 1)
        If pos < 1 Then Exit Sub
        fn(i) = RTrim(Mid(line_part, 1, pos))
        i = i + 1
        pos = pos + 2
        line_part = Mid(line_part, pos)
    Wend
End Sub         ''unpackit

''  $ translate_text  Translation of one string to another within a string
'' jrs 05-18-87
'' jrs 05-18-87
'' function string convertion
'' This function changes one string to another within a longer string! It
'' is used to replace general parameters in WHERE clauses built in the
'' reports SELECTion criteria!
'' Passed Parameters:      w   the WHERE clause
'' p the parm to be replaced
'' c the column name to use instead
'' Returned Parameters:    w   The WHERE clause after the change

Public Function translate_text(w As String, p As String, c As String) As String
        
    Dim pos As Long
    pos = InStr(w, p)
    Select Case pos
        Case 1
                w = RTrim(c) & Mid(w, Len(p) + 1)
                translate_text = w: Exit Function
        Case Is > 1
                w = RTrim(Mid(w, 1, pos - 1)) & c & RTrim(Mid(w, pos + Len(p)))
                translate_text = w: Exit Function
    End Select 'Case
    translate_text = w
End Function    ''translate_text

''  $ get_company_name        Returns company name FROM memory
'' jrs 5-18-87
'' jrs 5-18-87
'' function company
'' This function return the name and address loaded during
'' start_of_program()!

Public Function get_company_name() As String
    get_company_name = pr_company_con_name
End Function ''get_company_name

Public Function gl_period_ok(d As Date) As Recordset
    Const SUB_NAME = "gl_period_ok"
    
    Dim pr_gl_period As Recordset   'Cgl_period
    Dim c_gpo As String
    '' declare c_gpo cursor for
    c_gpo = _
                " SELECT * FROM gl_period" & _
                " WHERE " & QStr(d) & " between glp_beg_dt and glp_end_dt"
    '' open c_gpo
    '' fetch c_gpo into pr_gl_period
    Set pr_gl_period = fnMainmodsOpenRecordset(c_gpo, SUB_NAME)
    If pr_gl_period Is Nothing Then
        Set gl_period_ok = Nothing
        Exit Function
    End If
    '' close c_gpo
    
    If Left(pr_gl_period!glp_status, 1) = "O" Then
            Set gl_period_ok = pr_gl_period: Exit Function
    End If
    If Left(pr_gl_period!glp_status, 1) = "W" Then
            Call dserror(2001)
            Set gl_period_ok = pr_gl_period: Exit Function
    End If
    Call dserror(2002)  'G/L Period is closed.
    Set gl_period_ok = Nothing
End Function    ''gl_period_ok

Public Function mst_pc() As Integer
    mst_pc = CInt(mstpc)            'note: mstpc defined as variant so Null can be used
End Function ''mst_pc

Public Sub set_mst_pc(c As Integer)
    mstpc = c
End Sub         ''set_mst_pc

Public Function term_units() As Double
    term_units = CDbl(termsunits)   'note: termsunits defined as variant so Null can be used
End Function ''term_units

Public Sub set_term_units(c As Double)
    termsunits = c
End Sub     ''set_term_units

Public Function set_mst_master() As Integer
    Dim pr_parm As Recordset    'Csys_parm
    Set pr_parm = fsys_parm(2, t_dbMainDatabase)
    If Not (pr_parm Is Nothing) Then mstpc = pr_parm!parm_field
    set_mst_master = mstpc
End Function ''set_mst_master


Public Sub rlocation_rsm(ByVal c As Variant, _
                                v_pc_int As Long, _
                                v_pump_int As Long, _
                                v_prft_name As String) 'returns the 3 " v_ ..." parms
        
    Dim b As Long, d As Long, e As Long
    
    v_prft_name = ""    'Null
    v_pc_int = 0
    v_pump_int = 0
    If c = 0 Then
        v_prft_name = rprft_name(v_pc_int)  'get the profit center name string
    ElseIf Not IsNull(c) Then
        b = c / 100000
        d = b * 100000
        v_pc_int = c - d
        e = c - v_pc_int
        v_pump_int = e / 100000
        v_prft_name = rprft_name(v_pc_int)
    End If
End Sub         ''rlocation_rsm

Public Function trunct(amt As Double, rnd As Long) As Double
        
    Dim i As Long
    Dim rndf As Double
    Dim rstg As String
    
    rndf = 0.5
    rstg = "##########."
    For i = 1 To rnd
        rndf = rndf / 10
        rstg = RTrim(rstg) & "#"
    Next
    Select Case amt
        Case Is < 0
                amt = amt + rndf
        Case Is > 0
                amt = amt - rndf
    End Select 'Case
    trunct = Format(amt, rstg)
End Function    ''trunct

Public Function dround(v As Double) As Double
    dround = roundf(v, 2)
    ' the algorithm below, implemented in 4GL, does not work in VB on a pc. I've left it
    ' in place should it become necessary to "tinker" with it; until then, the old way
    ' (which is also used for very large numbers) is employed.
    #If False Then
        Dim chop_it As Long
        Dim int_part As Long
        Dim dec_part As Double
        '' This is done to avoid the "Value too large to fit in an INTEGER" error!
        '' The function roundf() is used only when the limit is exceeded because
        '' it is slower!  The use of roundf() here will probably be rare!
        '' Technically, the cutoff should be at about 21,474,830 because it is
        '' then multiplied by 100!  But I tested this on the 6000, 280, and RT
        '' and it works! DH 09/18/91
        '' 2,147,483,000
        If (v > 2147483000) Then
            v = roundf(v, 2)
            dround = v: Exit Function
        End If
        int_part = v
        dec_part = v - int_part
        If (dec_part < 0) Then
            chop_it = (dec_part * 100) - 0.5
        Else
            chop_it = (dec_part * 100) + 0.5
        End If
        dec_part = chop_it / 100
        v = int_part + dec_part
        dround = v: Exit Function
    #End If
End Function    ''dround

Public Function roundf(amt As Double, rnd As Integer) As Double
    Dim i As Integer
    Dim rstg As String
    
    'rndf = 0.5
    rstg = "############0."     '"-------------."
    For i = 1 To rnd
        rstg = RTrim(rstg) & "#"
    Next
    roundf = Format(amt, rstg)
End Function    ''roundf

Public Function abs_value(v As Double) As Double
    'note: even tho I'm hard pressed for time, I refuse to do this
    'If v < 0 Then v = v * -1
    abs_value = Abs(v)
End Function    ''abs_value

Private Sub startlog(sMsg As String)
    #If DEVELOPMENT Then
        MsgBox sMsg, vbInformation, "startlog"
    #End If
End Sub 'startlog
Private Sub errorlog(sMsg As String)
    #If DEVELOPMENT Then
        MsgBox sMsg, vbInformation, "errorlog"
    #End If
End Sub 'errorlog

Public Sub next_program(prnt As String, munit As String, mprog As String)
    Dim aix_unit As String
    aix_unit = "DBPATH=" & RTrim(Environ("DBPATH")) & _
                    ";export DBPATH " & _
                    "PARENT=" & RTrim(prnt) & ";export PARENT " & _
                    "FUNIT=" & RTrim(munit) & ";export FUNIT " & _
                    RTrim(mprog) & ".4ge " & RTrim(munit)
    Call prep_next(aix_unit)
End Sub ''next_program


Public Sub prep_next(run_str As String)
    Dim batch_name As String
    batch_name = Environ("TTY")
    If batch_name = "" Then Exit Sub
    Call startlog(batch_name)
    Call errorlog(Mid(run_str, 1, 100))
    Call errorlog(Mid(run_str, 101))
End Sub ''prep_next

Public Sub prep_logn()
    Dim aix_unit As String
    aix_unit = "DBPATH=" & RTrim(Environ("DATADIR")) & _
                    ":" & RTrim(Environ("PROGPATH")) & _
                    " export DBPATH logn"
    Call prep_next(aix_unit)
End Sub ''prep_logn

Public Sub prep_parent()
    Dim aix_unit As String
    Dim parent As String
    parent = Environ("PARENT")
    If parent = "logn" Then
        Call prep_logn
        Exit Sub
    End If
    aix_unit = "DBPATH=" & RTrim(Environ("DBPATH")) & _
                    " export DBPATH " & _
                    "FUNIT=" & RTrim(parent) & _
                    " export FUNIT " & _
                    RTrim(parent) & ".4ge " & _
                    RTrim(parent)
    Call prep_next(aix_unit)
End Sub ''prep_parent

Public Function makef(v As Variant, lSize As Long) As String
    ' right align 'lSize' digits (including decimal) after stripping trailing zeroes;
    ' ensure this stripped number does not "overflow"; if it does return "splat" string.
    Dim long_text As String
    Dim txt As String
    Dim c As String * 1
    Dim i As Long
    
    If IsNull(v) Then
        txt = ""              '0        1         2
    Else                      '1234567890123456789012345    'original 4GL string:
        long_text = Format(v, "###############0.00000000")  '"---------------&.########")
        i = Len(long_text) '25
        c = Mid(long_text, i, 1)
        Do While c = "0"
            i = i - 1
            c = Mid(long_text, i, 1)
            If c = "." Then
                i = i - 1
                Exit Do 'While
            End If
        Loop 'While
        txt = Mid(long_text, 1, i)
        If Len(txt) > lSize Then txt = "*****************"
    End If
    makef = txt

End Function    ''makef

Public Function process_locked(p As String, Optional vLockUser) As Boolean
    Const SUB_NAME = "process_locked"
    
    Dim u As String
    Dim rs As Recordset
    Dim c_sys_lock As String
    
    If Not IsMissing(vLockUser) Then
        vLockUser = ""
    End If
    
    c_sys_lock = _
                    " SELECT lock_user FROM sys_lock" & _
                    " WHERE lock_name = " & QStr(p)
    Set rs = fnMainmodsOpenRecordset(c_sys_lock, SUB_NAME)
    
    If rs Is Nothing Then
        process_locked = True
        Exit Function
    End If
    
    If rs.RecordCount = 0 Then
        u = tfnGetUserName
        If fnMainmodsExecute("insert into sys_lock values (" & QStr(p) & "," & QStr(u) & ")", SUB_NAME) Then
            process_locked = False  ' Thats good, we got a lock on the process
        Else
            process_locked = True
        End If
    Else
        process_locked = True   ' Thats bad, someone else has a lock on the process
        
        If Not IsMissing(vLockUser) Then
            If Not IsNull(rs!lock_user) Then
                vLockUser = Trim(rs!lock_user)
            End If
        End If
    End If
    CleanUp rs
End Function    ''process_locked

Public Function unlock_process(p As String) As Boolean
    Const SUB_NAME = "unlock_process"
    unlock_process = fnMainmodsExecute("delete FROM sys_lock WHERE lock_name = " & QStr(p), SUB_NAME)
End Function    ''unlock_process

Public Function rprft_name(code As Long) As String
    Dim pr_sys_prft_ctr As Recordset    'Csys_prft_ctr
    Set pr_sys_prft_ctr = fsys_prft_ctr(code, t_dbMainDatabase)
    If Not (pr_sys_prft_ctr Is Nothing) Then rprft_name = pr_sys_prft_ctr!prft_name
    CleanUp pr_sys_prft_ctr
End Function    ''rprft_name

Public Function main_warning()
    Dim stat As Long
    Dim err_text As String
    ''stat = status
    ''err_text = err_get(stat)
    '' error err_text
    #If DEVELOPMENT Then
        MsgBox "Unfinished function: emulation of 4GL 'err_get' needed", vbExclamation, "sub main_warning"
    #End If
    main_warning = err_text
End Function ''main_warning

Public Function main_error()
    Dim stat As Long
    Dim err_text As String
    ''stat = status
    ''err_text = err_get(stat)
    '' error err_text
    #If DEVELOPMENT Then
        MsgBox "Unfinished function: emulation of 4GL 'err_get' needed", vbExclamation, "sub main_error"
    #End If
    main_error = err_text
End Function ''main_error
