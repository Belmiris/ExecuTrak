Attribute VB_Name = "ModRefnfast"
Attribute VB_Description = "REFNFAST.4GL translation 1.19 - Last Delta 8/16/95"
'***********************************************************************************
'
' Copyright (c) 1996 FACTOR, A Division of W.R.Hess Company
'
' Module name   : REFNFAST.BAS
' Date          : Feb 29, 1996
' Programmer(s) :
'
' This module is a 4GL-to-BAS translation of:
' @(#) refnfast.4gl 1.19@(#) Last Delta: 8/16/95 }
' Copyright (c) 1988,89,90,91 FACTOR, A Division of WR Hess Company
'
' Functions:    :   create_temp_fmtank      create_temp_secloc
'                   far_altname
'                   far_class
'                   far_customer
'                   far_fnc_table
'                   far_type                far_type_ctr
'                   fco_inv_codes
'                   ffm_frgt_cd             ffm_frgt_cd2
'                   ffm_purch_term
'                   ffm_tank                ffm_tank_hist           ffm_tank_lnk
'                   ffm_terminal
'                   ffo_class
'                   fgf_header
'                   fgl_master
'                   fgl_period
'                   fp_altname
'                   fp_class
'                   fprodlnk                fproduct
'                   fp_vendor
'                   frs_b_period
'                   frs_cat                 frs_alt_cat
'                   finv_classes
'                   finv_header
'                   finv_master
'                   finv_uom
'                   finv_whlsale
'                   frs_b_item
'                   frs_meter
'                   frs_shiftlink           fshift_link
'                   fsp_excpt_clas          fsp_excpt_cust
'                   fsp_group
'                   fsys_parm
'                   fsys_prft_ctr
'                   ftx_table
'                   ftx_use_g               ftx_use_gl              ftx_use_g_link
'                   fvendor
'                   it_is_a_fuel
'                   program_time_stamp
'                   mastr_pc
'                   rgfh_path_code
'                   set_mastr_master
'                   stdconv
'                   uround
'
'***********************************************************************************

Option Base 1
Option Explicit

#Const cDEBUG = 0
#Const DEVELOPMENT = 1
Const conDimension = 10

' Dim p_ln As String, p_fn As String, p_al As String, p_a1 As String, p_a2 As String
' Dim ar_ln As String, ar_fn As String, ar_al As String, ar_a1 As String, ar_a2 As String
' Dim class_type As String * 1
' Dim pc_active As Long
' Dim pc_number As Long
' Dim non_fuel As Boolean 'Long
Dim master_prft_ctr As Integer
Dim dbPrivateDB As DataBase
'

#If False Then
' prototype -- instructions for use:
' (1) copy this code;
' (2) replace all QQ with the appropriate table name;
' (3) replace all ZZ with appropriate "WHERE" criterion;
' (4) replace all TT with appropriate parameter type(s).
Public Function fQQ( _
                                sParm1 As TT, _
                                db As DataBase _
                                ) As Recordset
    Const SUB_NAME = "fQQ"
    'returning pr_QQ As Recordset
        
    Dim sSql As String
    Dim iIndex As Integer
    Dim rs As Recordset
    Static rsArray() As Recordset

    On Error GoTo fQQRedim
    iIndex = 1
    Do While Not (rsArray(iIndex) Is Nothing)
        If rsArray(iIndex)!ZZ = sParm1 Then
            Set fQQ = rsArray(iIndex)
            Exit Function   'return now with this "buffered" recordset
        End If
        iIndex = iIndex + 1
        If iIndex > conDimension Then Exit Do 'While  'Note: iIndex will be used below
    Loop 'While
    
    sSql = "SELECT * FROM QQ WHERE ZZ = " & QStr(sParm1)
    Set rs = fnMainmodsOpenRecordset(sSql, SUB_NAME, db)
    If rs Is Nothing Then
        Exit Function
    End If

    If rs.RecordCount > 0 Then
        rs.MoveFirst    'populate with first record
        #If cDEBUG Then
            On Error Resume Next
            While Not rs.EOF
                Debug.Print "fQQ(" & QStr(sParm1) & ") = " & rs.Fields(0)
                rs.MoveNext    ' Locate next record
            Wend
            On Error GoTo 0
            rs.MoveFirst
        #End If
        If iIndex > conDimension Then
            ' shuffle array elements up (lose oldest -- highest -- element)
            For iIndex = conDimension - 1 To 1 Step -1
                Set rsArray(iIndex + 1) = rsArray(iIndex)
            Next
            iIndex = 1 'the current recordset will be placed as the first element
        End If
        Set rsArray(iIndex) = rs        'buffer this recordset (which also gets returned)
    Else
        iIndex = conDimension
        ReDim rsArray(iIndex)           'this will cause "Nothing" to be returned
    End If
    
    Set fQQ = rsArray(iIndex)
    Exit Function
    
fQQRedim:
    If Err.Number = 9 Then 'Subscript out of range
        ' this is the expected error on the FIRST call to this function; we
        ' must dimension the static array and retry the first reference to it.
        Err.Clear
        ReDim rsArray(conDimension)
        On Error GoTo 0
        Resume
    Else
        Err.Raise Err.Number
    End If
    
End Function    ''fQQ
#End If

Public Function QStr(numORstr As Variant) As String
    Dim nIdx As Integer
    Dim nPos As Integer
    Dim szParameter As String
    
    If IsNull(numORstr) Then
        szParameter = ""
    Else
        szParameter = numORstr
    End If

    nIdx = 1
    nPos = InStr(nIdx, szParameter, "'")
    
    While nPos <> 0
        szParameter = Left(szParameter, nPos) & "'" & Right(szParameter, Len(szParameter) - nPos)
        nIdx = nPos + 2
        nPos = InStr(nIdx, szParameter, "'")
    Wend
    
    QStr = "'" & szParameter & "'"
End Function 'QStr


Public Function create_temp_fmtank(strt_dt As Date, end_dt As Date, db As DataBase) As Long
    Const SUB_NAME = "create_temp_fmtank"
    
    Dim ret_cnt As Long
    Dim rs As Recordset
    Dim sSql As String
    
    Set dbPrivateDB = db    'use global database var to avoid passing as parm
     
    fnMainmodsExecute "SELECT * FROM fm_tank WHERE tk_strt_date <= " & QStr(end_dt) & _
                " INTO TEMP t1_fmtank", SUB_NAME, db
    fnMainmodsExecute "insert into t1_fmtank " & _
                " select * from fm_tank_hist where tkh_strt_date <= " & QStr(end_dt), SUB_NAME, db
    fnMainmodsExecute "create index t1_ctf_fmtank on t1_fmtank(tk_prft_ctr,tk_num)", SUB_NAME, db
    fnMainmodsExecute "select tk_prft_ctr t2_1, tk_num t2_2, max (tk_strt_date) max_dt from t1_fmtank" & _
                " where tk_strt_date < " & QStr(strt_dt) & _
                " group by tk_prft_ctr,tk_num into temp t2_fmtank", SUB_NAME, db
    fnMainmodsExecute "create index t2_ctf_fmtank on t2_fmtank(t2_1,t2_2)", SUB_NAME, db
    fnMainmodsExecute "select * from t1_fmtank,t2_fmtank" & _
                " where tk_prft_ctr = t2_1" & _
                " and tk_num = t2_2" & _
                " and max_dt <= tk_strt_date" & _
                " into temp t3_fmtank", SUB_NAME, db
    fnMainmodsExecute "create index t3_ctf_fmtank on t3_fmtank(tk_prft_ctr,tk_num)", SUB_NAME, db
    fnMainmodsExecute "select tk_prft_ctr t4_1, tk_num t4_2, tk_prodlnk t4_3, min (tk_strt_date) min_dt from t3_fmtank" & _
                " where 1 = 1" & _
                " group by tk_prft_ctr,tk_num,tk_prodlnk into temp t4_fmtank", SUB_NAME, db
    fnMainmodsExecute " create index t4_ctf_fmtank on t4_fmtank(t4_1,t4_2,t4_3)", SUB_NAME, db
   
    fnMainmodsExecute "select * from t3_fmtank,t4_fmtank" & _
                " where tk_prft_ctr = t4_1" & _
                " and tk_num = t4_2" & _
                " and tk_prodlnk = t4_3" & _
                " and tk_strt_date = min_dt" & _
                " into temp temp_fmtank", SUB_NAME, db
    fnMainmodsExecute "create index temp_ctf_fmtank on temp_fmtank(tk_prft_ctr,tk_num)", SUB_NAME, db
    
        sSql = "SELECT COUNT(*) nCount FROM temp_fmtank"
        Set rs = fnMainmodsOpenRecordset(sSql, SUB_NAME, db)
        If rs Is Nothing Then
            Exit Function
        End If
        
        If Not (rs Is Nothing) Then ret_cnt = rs!nCount
    On Error GoTo 0
    
    fnMainmodsExecute "drop table t1_fmtank", SUB_NAME, db
    fnMainmodsExecute "drop table t2_fmtank", SUB_NAME, db
    fnMainmodsExecute "drop table t3_fmtank", SUB_NAME, db
    fnMainmodsExecute "drop table t4_fmtank", SUB_NAME, db
    create_temp_fmtank = ret_cnt

End Function    ''create_temp_fmtank

Public Function create_temp_secloc(strt_dt As Date, end_dt As Date, _
                                    incld As String, db As DataBase) As Long
    Const SUB_NAME = "create_temp_secloc"
        
    ' Dim incld As String * 1 ''' like F-Fuel only,  A-All
    Dim ret_cnt As Long
    Dim sSql As String
    Dim rs As Recordset
    
    Set dbPrivateDB = db    'use global database var to avoid passing as parm
    
    ret_cnt = create_temp_fmtank(strt_dt, end_dt, db)
    fnMainmodsExecute "create index temp_cts_fmtank on temp_fmtank (tk_prft_ctr,tk_num,tk_prodlnk)", SUB_NAME, db
    If Left(incld, 1) = "F" Then
            fnMainmodsExecute "select * from inv_secndloc,temp_fmtank" & _
                        " where sl_prft_ctr = tk_prft_ctr" & _
                        " and sl_loc = tk_num" & _
                        " and sl_prodlnk = tk_prodlnk" & _
                        " into temp temp_secloc", SUB_NAME, db
    ElseIf Left(incld, 1) = "A" Then
            fnMainmodsExecute "select * from inv_secndloc" & _
                        " where 1=1" & _
                        " into temp temp_secloc", SUB_NAME, db
    End If
    fnMainmodsExecute "create index temp_ctf_secloc on temp_secloc(sl_prft_ctr,sl_prodlnk,sl_loc)", SUB_NAME, db
       
    On Error GoTo create_temp_seclocERROR
        '' " select count(*) into ret_cnt from temp_secloc":
        sSql = "SELECT COUNT(*) nCount FROM temp_secloc"
        Set rs = fnMainmodsOpenRecordset(sSql, SUB_NAME, db)
        If rs Is Nothing Then
            Exit Function
        End If
        ret_cnt = rs!nCount
    On Error GoTo 0
    
    create_temp_secloc = ret_cnt
    Exit Function
    
create_temp_seclocERROR:
    If Err.Number > 3100 Then
        #If DEVELOPMENT Then
            Beep
            MsgBox "DB call failed on table temp_secloc; SQL =" & sSql, vbCritical, "create_temp_secloc"
        #End If
        Err.Clear
        Resume Next
    Else
        Err.Raise Err.Number
    End If

End Function    ''create_temp_secloc

Public Function far_altname( _
                                sParm1 As Long, _
                                db As DataBase _
                                ) As Recordset
    Const SUB_NAME = "far_altname"
    'returning pr_ar_altname As Recordset
        
    Dim sSql As String
    Dim iIndex As Integer
    Dim rs As Recordset
    Static rsArray() As Recordset

    On Error GoTo far_altnameRedim
    iIndex = 1
    Do While Not (rsArray(iIndex) Is Nothing)
        If rsArray(iIndex)!an_customer = sParm1 Then
            Set far_altname = rsArray(iIndex)
            Exit Function   'return now with this "buffered" recordset
        End If
        iIndex = iIndex + 1
        If iIndex > conDimension Then Exit Do 'While  'Note: iIndex will be used below
    Loop 'While
    
    sSql = "SELECT * FROM ar_altname WHERE an_customer = " & QStr(sParm1)
    Set rs = fnMainmodsOpenRecordset(sSql, SUB_NAME, db)
    If rs.RecordCount > 0 Then
        rs.MoveFirst    'populate with first record
        #If cDEBUG Then
            On Error Resume Next
            While Not rs.EOF
                Debug.Print "far_altname(" & QStr(sParm1) & ") = " & rs.Fields(0)
                rs.MoveNext    ' Locate next record
            Wend
            On Error GoTo 0
            rs.MoveFirst
        #End If
        If iIndex > conDimension Then
            ' shuffle array elements up (lose oldest -- highest -- element)
            For iIndex = conDimension - 1 To 1 Step -1
                Set rsArray(iIndex + 1) = rsArray(iIndex)
            Next
            iIndex = 1 'the current recordset will be placed as the first element
        End If
        Set rsArray(iIndex) = rs        'buffer this recordset (which also gets returned)
    Else
        iIndex = conDimension
        ReDim rsArray(iIndex)           'this will cause "Nothing" to be returned
    End If
    
    Set far_altname = rsArray(iIndex)
    Exit Function
    
far_altnameRedim:
    If Err.Number = 9 Then 'Subscript out of range
        ' this is the expected error on the FIRST call to this function; we
        ' must dimension the static array and retry the first reference to it.
        Err.Clear
        ReDim rsArray(conDimension)
        On Error GoTo 0
        Resume
    Else
        Err.Raise Err.Number
    End If
    
End Function    ''far_altname

Public Function far_class( _
                                sParm1 As String, _
                                db As DataBase _
                                ) As Recordset
    Const SUB_NAME = "far_class"
    
    'returning pr_ar_class As Recordset
        
    Dim sSql As String
    Dim iIndex As Integer
    Dim rs As Recordset
    Static rsArray() As Recordset

    On Error GoTo far_classRedim
    iIndex = 1
    Do While Not (rsArray(iIndex) Is Nothing)
        If rsArray(iIndex)!arc_code = sParm1 Then
            Set far_class = rsArray(iIndex)
            Exit Function   'return now with this "buffered" recordset
        End If
        iIndex = iIndex + 1
        If iIndex > conDimension Then Exit Do 'While  'Note: iIndex will be used below
    Loop 'While
    
    sSql = "SELECT * FROM ar_class WHERE arc_code = " & QStr(sParm1)
    Set rs = fnMainmodsOpenRecordset(sSql, SUB_NAME, db)
    If rs Is Nothing Then
        Exit Function
    End If
    If rs.RecordCount > 0 Then
        rs.MoveFirst    'populate with first record
        #If cDEBUG Then
            On Error Resume Next
            While Not rs.EOF
                Debug.Print "far_class(" & QStr(sParm1) & ") = " & rs.Fields(0)
                rs.MoveNext    ' Locate next record
            Wend
            On Error GoTo 0
            rs.MoveFirst
        #End If
        If iIndex > conDimension Then
            ' shuffle array elements up (lose oldest -- highest -- element)
            For iIndex = conDimension - 1 To 1 Step -1
                Set rsArray(iIndex + 1) = rsArray(iIndex)
            Next
            iIndex = 1 'the current recordset will be placed as the first element
        End If
        Set rsArray(iIndex) = rs        'buffer this recordset (which also gets returned)
    Else
        iIndex = conDimension
        ReDim rsArray(iIndex)           'this will cause "Nothing" to be returned
    End If
    
    Set far_class = rsArray(iIndex)
    Exit Function
    
far_classRedim:
    If Err.Number = 9 Then 'Subscript out of range
        ' this is the expected error on the FIRST call to this function; we
        ' must dimension the static array and retry the first reference to it.
        Err.Clear
        ReDim rsArray(conDimension)
        On Error GoTo 0
        Resume
    Else
        Err.Raise Err.Number
    End If
    
End Function    ''far_class


Public Function far_customer( _
                                sParm1 As Long, _
                                db As DataBase _
                                ) As Recordset
    Const SUB_NAME = "far_customer"
    'returning pr_ar_customer As Recordset
        
    Dim sSql As String
    Dim iIndex As Integer
    Dim rs As Recordset
    Static rsArray() As Recordset

    On Error GoTo far_customerRedim
    iIndex = 1
    Do While Not (rsArray(iIndex) Is Nothing)
        If rsArray(iIndex)!cust_customer = sParm1 Then
            Set far_customer = rsArray(iIndex)
            Exit Function   'return now with this "buffered" recordset
        End If
        iIndex = iIndex + 1
        If iIndex > conDimension Then Exit Do 'While  'Note: iIndex will be used below
    Loop 'While
    
    sSql = "SELECT * FROM ar_customer WHERE cust_customer = " & QStr(sParm1)
    Set rs = fnMainmodsOpenRecordset(sSql, SUB_NAME, db)
    If rs Is Nothing Then
        Exit Function
    End If
    If rs.RecordCount > 0 Then
        rs.MoveFirst    'populate with first record
        #If cDEBUG Then
            On Error Resume Next
            While Not rs.EOF
                Debug.Print "far_customer(" & QStr(sParm1) & ") = " & rs.Fields(0)
                rs.MoveNext    ' Locate next record
            Wend
            On Error GoTo 0
            rs.MoveFirst
        #End If
        If iIndex > conDimension Then
            ' shuffle array elements up (lose oldest -- highest -- element)
            For iIndex = conDimension - 1 To 1 Step -1
                Set rsArray(iIndex + 1) = rsArray(iIndex)
            Next
            iIndex = 1 'the current recordset will be placed as the first element
        End If
        Set rsArray(iIndex) = rs        'buffer this recordset (which also gets returned)
    Else
        iIndex = conDimension
        ReDim rsArray(iIndex)           'this will cause "Nothing" to be returned
    End If
    
    Set far_customer = rsArray(iIndex)
    Exit Function
    
far_customerRedim:
    If Err.Number = 9 Then 'Subscript out of range
        ' this is the expected error on the FIRST call to this function; we
        ' must dimension the static array and retry the first reference to it.
        Err.Clear
        ReDim rsArray(conDimension)
        On Error GoTo 0
        Resume
    Else
        Err.Raise Err.Number
    End If
    
End Function    ''far_customer

Public Function far_fnc_table( _
                                sParm1 As String, _
                                db As DataBase _
                                ) As Recordset
    Const SUB_NAME = "far_fnc_table"
    
    'returning pr_ar_fnc_table As Recordset
        
    Dim sSql As String
    Dim iIndex As Integer
    Dim rs As Recordset
    Static rsArray() As Recordset

    On Error GoTo far_fnc_tableRedim
    iIndex = 1
    Do While Not (rsArray(iIndex) Is Nothing)
        If rsArray(iIndex)!fc_fnc_code = sParm1 Then
            Set far_fnc_table = rsArray(iIndex)
            Exit Function   'return now with this "buffered" recordset
        End If
        iIndex = iIndex + 1
        If iIndex > conDimension Then Exit Do 'While  'Note: iIndex will be used below
    Loop 'While
    
    sSql = "SELECT * FROM ar_fnc_table WHERE fc_fnc_code = " & QStr(sParm1)
    Set rs = fnMainmodsOpenRecordset(sSql, SUB_NAME, db)
    If rs Is Nothing Then
        Exit Function
    End If
    If rs.RecordCount > 0 Then
        rs.MoveFirst    'populate with first record
        #If cDEBUG Then
            On Error Resume Next
            While Not rs.EOF
                Debug.Print "far_fnc_table(" & QStr(sParm1) & ") = " & rs.Fields(0)
                rs.MoveNext    ' Locate next record
            Wend
            On Error GoTo 0
            rs.MoveFirst
        #End If
        If iIndex > conDimension Then
            ' shuffle array elements up (lose oldest -- highest -- element)
            For iIndex = conDimension - 1 To 1 Step -1
                Set rsArray(iIndex + 1) = rsArray(iIndex)
            Next
            iIndex = 1 'the current recordset will be placed as the first element
        End If
        Set rsArray(iIndex) = rs        'buffer this recordset (which also gets returned)
    Else
        iIndex = conDimension
        ReDim rsArray(iIndex)           'this will cause "Nothing" to be returned
    End If
    
    Set far_fnc_table = rsArray(iIndex)
    Exit Function
    
far_fnc_tableRedim:
    If Err.Number = 9 Then 'Subscript out of range
        ' this is the expected error on the FIRST call to this function; we
        ' must dimension the static array and retry the first reference to it.
        Err.Clear
        ReDim rsArray(conDimension)
        On Error GoTo 0
        Resume
    Else
        Err.Raise Err.Number
    End If
    
End Function    ''far_fnc_table

Public Function far_type( _
                                sParm1 As String, _
                                db As DataBase _
                                ) As Recordset
    Const SUB_NAME = "far_type"
    'returning pr_ar_type As Recordset
        
    Dim sSql As String
    Dim iIndex As Integer
    Dim rs As Recordset
    Static rsArray() As Recordset

    On Error GoTo far_typeRedim
    iIndex = 1
    Do While Not (rsArray(iIndex) Is Nothing)
        If rsArray(iIndex)!at_type_code = sParm1 Then
            Set far_type = rsArray(iIndex)
            Exit Function   'return now with this "buffered" recordset
        End If
        iIndex = iIndex + 1
        If iIndex > conDimension Then Exit Do 'While  'Note: iIndex will be used below
    Loop 'While
    
    sSql = "SELECT * FROM ar_type WHERE at_type_code = " & QStr(sParm1)
    Set rs = fnMainmodsOpenRecordset(sSql, SUB_NAME, db)
    If rs Is Nothing Then
        Exit Function
    End If
    If rs.RecordCount > 0 Then
        rs.MoveFirst    'populate with first record
        #If cDEBUG Then
            On Error Resume Next
            While Not rs.EOF
                Debug.Print "far_type(" & QStr(sParm1) & ") = " & rs.Fields(0)
                rs.MoveNext    ' Locate next record
            Wend
            On Error GoTo 0
            rs.MoveFirst
        #End If
        If iIndex > conDimension Then
            ' shuffle array elements up (lose oldest -- highest -- element)
            For iIndex = conDimension - 1 To 1 Step -1
                Set rsArray(iIndex + 1) = rsArray(iIndex)
            Next
            iIndex = 1 'the current recordset will be placed as the first element
        End If
        Set rsArray(iIndex) = rs        'buffer this recordset (which also gets returned)
    Else
        iIndex = conDimension
        ReDim rsArray(iIndex)           'this will cause "Nothing" to be returned
    End If
    
    Set far_type = rsArray(iIndex)
    Exit Function
    
far_typeRedim:
    If Err.Number = 9 Then 'Subscript out of range
        ' this is the expected error on the FIRST call to this function; we
        ' must dimension the static array and retry the first reference to it.
        Err.Clear
        ReDim rsArray(conDimension)
        On Error GoTo 0
        Resume
    Else
        Err.Raise Err.Number
    End If
    
End Function    ''far_type

Public Function far_type_ctr( _
                                sParm1 As String, _
                                sParm2 As Integer, _
                                db As DataBase _
                                ) As Recordset
    Const SUB_NAME = "far_type_ctr"
    
    'returning pr_ar_type_ctr As Recordset
        
    Dim sSql As String
    Dim iIndex As Integer
    Dim rs As Recordset
    Static rsArray() As Recordset

    On Error GoTo far_type_ctrRedim
    iIndex = 1
    Do While Not (rsArray(iIndex) Is Nothing)
        If rsArray(iIndex)!at_type_code = sParm1 And _
           rsArray(iIndex)!at_prft_ctr = sParm2 Then
            Set far_type_ctr = rsArray(iIndex)
            Exit Function   'return now with this "buffered" recordset
        End If
        iIndex = iIndex + 1
        If iIndex > conDimension Then Exit Do 'While  'Note: iIndex will be used below
    Loop 'While
    
    sSql = "SELECT * FROM ar_type WHERE at_type_code = " & QStr(sParm1) & _
           " AND at_prft_ctr = " & QStr(sParm2)
    Set rs = fnMainmodsOpenRecordset(sSql, SUB_NAME, db)
    If rs Is Nothing Then
        Exit Function
    End If
    If rs.RecordCount > 0 Then
        rs.MoveFirst    'populate with first record
        #If cDEBUG Then
            On Error Resume Next
            While Not rs.EOF
                Debug.Print "far_type_ctr(" & QStr(sParm1) & ") = " & rs.Fields(0)
                rs.MoveNext    ' Locate next record
            Wend
            On Error GoTo 0
            rs.MoveFirst
        #End If
        If iIndex > conDimension Then
            ' shuffle array elements up (lose oldest -- highest -- element)
            For iIndex = conDimension - 1 To 1 Step -1
                Set rsArray(iIndex + 1) = rsArray(iIndex)
            Next
            iIndex = 1 'the current recordset will be placed as the first element
        End If
        Set rsArray(iIndex) = rs        'buffer this recordset (which also gets returned)
    Else
        iIndex = conDimension
        ReDim rsArray(iIndex)           'this will cause "Nothing" to be returned
    End If
    
    Set far_type_ctr = rsArray(iIndex)
    Exit Function
    
far_type_ctrRedim:
    If Err.Number = 9 Then 'Subscript out of range
        ' this is the expected error on the FIRST call to this function; we
        ' must dimension the static array and retry the first reference to it.
        Err.Clear
        ReDim rsArray(conDimension)
        On Error GoTo 0
        Resume
    Else
        Err.Raise Err.Number
    End If
    
End Function    ''far_type_ctr

Public Function fco_inv_codes( _
                                sParm1 As String, _
                                db As DataBase _
                                ) As Recordset
    Const SUB_NAME = "fco_inv_codes"
    
    'returning pr_co_inv_codes As Recordset
        
    Dim sSql As String
    Dim iIndex As Integer
    Dim rs As Recordset
    Static rsArray() As Recordset

    On Error GoTo fco_inv_codesRedim
    iIndex = 1
    Do While Not (rsArray(iIndex) Is Nothing)
        If rsArray(iIndex)!coic_user_class = sParm1 Then
            Set fco_inv_codes = rsArray(iIndex)
            Exit Function   'return now with this "buffered" recordset
        End If
        iIndex = iIndex + 1
        If iIndex > conDimension Then Exit Do 'While  'Note: iIndex will be used below
    Loop 'While
    
    sSql = "SELECT * FROM co_inv_codes WHERE coic_user_class = " & QStr(sParm1)
    Set rs = fnMainmodsOpenRecordset(sSql, SUB_NAME, db)
    If rs Is Nothing Then
        Exit Function
    End If
    If rs.RecordCount > 0 Then
        rs.MoveFirst    'populate with first record
        #If cDEBUG Then
            On Error Resume Next
            While Not rs.EOF
                Debug.Print "fco_inv_codes(" & QStr(sParm1) & ") = " & rs.Fields(0)
                rs.MoveNext    ' Locate next record
            Wend
            On Error GoTo 0
            rs.MoveFirst
        #End If
        If iIndex > conDimension Then
            ' shuffle array elements up (lose oldest -- highest -- element)
            For iIndex = conDimension - 1 To 1 Step -1
                Set rsArray(iIndex + 1) = rsArray(iIndex)
            Next
            iIndex = 1 'the current recordset will be placed as the first element
        End If
        Set rsArray(iIndex) = rs        'buffer this recordset (which also gets returned)
    Else
        iIndex = conDimension
        ReDim rsArray(iIndex)           'this will cause "Nothing" to be returned
    End If
    
    Set fco_inv_codes = rsArray(iIndex)
    Exit Function
    
fco_inv_codesRedim:
    If Err.Number = 9 Then 'Subscript out of range
        ' this is the expected error on the FIRST call to this function; we
        ' must dimension the static array and retry the first reference to it.
        Err.Clear
        ReDim rsArray(conDimension)
        On Error GoTo 0
        Resume
    Else
        Err.Raise Err.Number
    End If
    
End Function    ''fco_inv_codes

Public Function ffm_frgt_cd( _
                                sParm1 As String, _
                                db As DataBase _
                                ) As Recordset
    Const SUB_NAME = "ffm_frgt_cd"
    
    'returning pr_fm_frgt_cd As Recordset
        
    Dim sSql As String
    Dim iIndex As Integer
    Dim rs As Recordset
    Static rsArray() As Recordset

    On Error GoTo ffm_frgt_cdRedim
    iIndex = 1
    Do While Not (rsArray(iIndex) Is Nothing)
        If rsArray(iIndex)!fmfc_code = sParm1 Then
            Set ffm_frgt_cd = rsArray(iIndex)
            Exit Function   'return now with this "buffered" recordset
        End If
        iIndex = iIndex + 1
        If iIndex > conDimension Then Exit Do 'While  'Note: iIndex will be used below
    Loop 'While
    
    sSql = "SELECT * FROM fm_frgt_cd WHERE fmfc_code = " & QStr(sParm1) & _
           " ORDER BY fmfc_prod_class,fmfc_date_end desc"
    Set rs = fnMainmodsOpenRecordset(sSql, SUB_NAME, db)
    If rs Is Nothing Then
        Exit Function
    End If
    If rs.RecordCount > 0 Then
        rs.MoveFirst    'populate with first record
        #If cDEBUG Then
            On Error Resume Next
            While Not rs.EOF
                Debug.Print "ffm_frgt_cd(" & QStr(sParm1) & ") = " & rs.Fields(0)
                rs.MoveNext    ' Locate next record
            Wend
            On Error GoTo 0
            rs.MoveFirst
        #End If
        If iIndex > conDimension Then
            ' shuffle array elements up (lose oldest -- highest -- element)
            For iIndex = conDimension - 1 To 1 Step -1
                Set rsArray(iIndex + 1) = rsArray(iIndex)
            Next
            iIndex = 1 'the current recordset will be placed as the first element
        End If
        Set rsArray(iIndex) = rs        'buffer this recordset (which also gets returned)
    Else
        iIndex = conDimension
        ReDim rsArray(iIndex)           'this will cause "Nothing" to be returned
    End If
    
    Set ffm_frgt_cd = rsArray(iIndex)
    Exit Function
    
ffm_frgt_cdRedim:
    If Err.Number = 9 Then 'Subscript out of range
        ' this is the expected error on the FIRST call to this function; we
        ' must dimension the static array and retry the first reference to it.
        Err.Clear
        ReDim rsArray(conDimension)
        On Error GoTo 0
        Resume
    Else
        Err.Raise Err.Number
    End If
    
End Function    ''ffm_frgt_cd

'Vijaya on 01/07/04 Magic#420344 Make same as 4gl Fucntion
'added Optional sParm4 As String = "" That is Carrier
'Changed the SQL added fmfc_vendor
Public Function ffm_frgt_cd2( _
                                sParm1 As String, _
                                sParm2 As String, _
                                sParm3 As Date, _
                                db As DataBase, _
                                Optional sParm4 As String = "") As Recordset
    Const SUB_NAME = "ffm_frgt_cd2"
    
    'returning pr_fm_frgt_cd As Recordset
        
    Dim sSql As String
    Dim iIndex As Integer
    Dim rs As Recordset
    Static rsArray() As Recordset

    On Error GoTo ffm_frgt_cd2Redim
    iIndex = 1
    Do While Not (rsArray(iIndex) Is Nothing)
        If Trim(rsArray(iIndex)!fmfc_code & "") = sParm1 And _
           Trim(rsArray(iIndex)!fmfc_prod_class & "") = sParm2 And _
           CDate(rsArray(iIndex)!fmfc_date_start) <= CDate(sParm3) And _
           CDate(rsArray(iIndex)!fmfc_date_end) >= CDate(sParm3) And _
           Trim(rsArray(iIndex)!fmfc_vendor & "") = sParm4 Then
           Set ffm_frgt_cd2 = rsArray(iIndex)
           Exit Function   'return now with this "buffered" recordset
        End If
        iIndex = iIndex + 1
        If iIndex > conDimension Then Exit Do 'While  'Note: iIndex will be used below
    Loop 'While
    
    sSql = " SELECT * from fm_frgt_cd "
    sSql = sSql & " WHERE fmfc_code = " & QStr(sParm1)
    sSql = sSql & " AND fmfc_prod_class = " & QStr(sParm2)
    sSql = sSql & " AND " & QStr(sParm3) & " BETWEEN fmfc_date_start AND fmfc_date_end"
    If sParm4 = "" Then
        sSql = sSql & " AND fmfc_vendor IS NULL"
    Else
        sSql = sSql & " AND fmfc_vendor = " & QStr(sParm4)
    End If

    Set rs = fnMainmodsOpenRecordset(sSql, SUB_NAME, db)
    If rs Is Nothing Then
        Exit Function
    End If
    If rs.RecordCount > 0 Then
        rs.MoveFirst    'populate with first record
        #If cDEBUG Then
            On Error Resume Next
            While Not rs.EOF
                Debug.Print "ffm_frgt_cd2(" & QStr(sParm1) & ") = " & rs.Fields(0)
                rs.MoveNext    ' Locate next record
            Wend
            On Error GoTo 0
            rs.MoveFirst
        #End If
        If iIndex > conDimension Then
            ' shuffle array elements up (lose oldest -- highest -- element)
            For iIndex = conDimension - 1 To 1 Step -1
                Set rsArray(iIndex + 1) = rsArray(iIndex)
            Next
            iIndex = 1 'the current recordset will be placed as the first element
        End If
        Set rsArray(iIndex) = rs        'buffer this recordset (which also gets returned)
    Else
        iIndex = conDimension
        ReDim rsArray(iIndex)           'this will cause "Nothing" to be returned
    End If
    
    Set ffm_frgt_cd2 = rsArray(iIndex)
    Exit Function
    
ffm_frgt_cd2Redim:
    If Err.Number = 9 Then 'Subscript out of range
        ' this is the expected error on the FIRST call to this function; we
        ' must dimension the static array and retry the first reference to it.
        Err.Clear
        ReDim rsArray(conDimension)
        On Error GoTo 0
        Resume
    Else
        Err.Raise Err.Number
    End If
    
End Function    ''ffm_frgt_cd2

Public Function ffm_purch_term( _
                                sParm1 As Integer, _
                                sParm2 As Long, _
                                sParm3 As Long, _
                                db As DataBase _
                                ) As Recordset
    Const SUB_NAME = "ffm_purch_term"
    
    'returning pr_fm_purch_term As Recordset
        
    Dim sSql As String
    Dim iIndex As Integer
    Dim rs As Recordset
    Static rsArray() As Recordset

    On Error GoTo ffm_purch_termRedim
    iIndex = 1
    Do While Not (rsArray(iIndex) Is Nothing)
        If rsArray(iIndex)!fmpt_terminal = sParm1 And _
           rsArray(iIndex)!fmpt_vendor = sParm2 And _
           rsArray(iIndex)!fmpt_prodlnk = sParm3 Then
            Set ffm_purch_term = rsArray(iIndex)
            Exit Function   'return now with this "buffered" recordset
        End If
        iIndex = iIndex + 1
        If iIndex > conDimension Then Exit Do 'While  'Note: iIndex will be used below
    Loop 'While
    
    sSql = "SELECT * FROM fm_purch_term WHERE fmpt_terminal = " & QStr(sParm1) & _
           " AND fmpt_vendor = " & QStr(sParm2) & _
           " AND fmpt_prodlnk = " & QStr(sParm3)
    Set rs = fnMainmodsOpenRecordset(sSql, SUB_NAME, db)
    If rs Is Nothing Then
        Exit Function
    End If
    If rs.RecordCount > 0 Then
        rs.MoveFirst    'populate with first record
        #If cDEBUG Then
            On Error Resume Next
            While Not rs.EOF
                Debug.Print "ffm_purch_term(" & QStr(sParm1) & ") = " & rs.Fields(0)
                rs.MoveNext    ' Locate next record
            Wend
            On Error GoTo 0
            rs.MoveFirst
        #End If
        If iIndex > conDimension Then
            ' shuffle array elements up (lose oldest -- highest -- element)
            For iIndex = conDimension - 1 To 1 Step -1
                Set rsArray(iIndex + 1) = rsArray(iIndex)
            Next
            iIndex = 1 'the current recordset will be placed as the first element
        End If
        Set rsArray(iIndex) = rs        'buffer this recordset (which also gets returned)
    Else
        iIndex = conDimension
        ReDim rsArray(iIndex)           'this will cause "Nothing" to be returned
    End If
    
    Set ffm_purch_term = rsArray(iIndex)
    Exit Function
    
ffm_purch_termRedim:
    If Err.Number = 9 Then 'Subscript out of range
        ' this is the expected error on the FIRST call to this function; we
        ' must dimension the static array and retry the first reference to it.
        Err.Clear
        ReDim rsArray(conDimension)
        On Error GoTo 0
        Resume
    Else
        Err.Raise Err.Number
    End If
    
End Function    ''ffm_purch_term

Public Function ffm_tank( _
                                sParm1 As Integer, _
                                sParm2 As String, _
                                sParm3 As Date, _
                                db As DataBase _
                                ) As Recordset
    Const SUB_NAME = "ffm_tank"
    
    'returning pr_fm_tank As Recordset
        
    Dim sSql As String
    Dim iIndex As Integer
    Dim rs As Recordset
    Static rsArray() As Recordset

    On Error GoTo ffm_tankRedim
    iIndex = 1
    Do While Not (rsArray(iIndex) Is Nothing)
        If rsArray(iIndex)!tk_prft_ctr = sParm1 And _
           rsArray(iIndex)!tk_num = sParm2 Then
            If rsArray(iIndex)!tk_strt_date <= sParm3 Then
                Set ffm_tank = rsArray(iIndex)
                Exit Function   'return now with this "buffered" recordset
            Else
                Set ffm_tank = ffm_tank_hist(sParm1, sParm2, sParm3, db)
                Exit Function   'return now with this "buffered" recordset
            End If
        End If
        iIndex = iIndex + 1
        If iIndex > conDimension Then Exit Do 'While  'Note: iIndex will be used below
    Loop 'While
    
    sSql = " SELECT * from fm_tank " & _
           " WHERE tk_prft_ctr = " & QStr(sParm1) & _
           " AND tk_num = " & QStr(sParm2)
    Set rs = fnMainmodsOpenRecordset(sSql, SUB_NAME, db)
    If rs Is Nothing Then
        Exit Function
    End If
    If rs.RecordCount > 0 Then
        rs.MoveFirst    'populate with first record
        #If cDEBUG Then
            On Error Resume Next
            While Not rs.EOF
                Debug.Print "ffm_tank(" & QStr(sParm1) & ") = " & rs.Fields(0)
                rs.MoveNext    ' Locate next record
            Wend
            On Error GoTo 0
            rs.MoveFirst
        #End If
        If iIndex > conDimension Then
            ' shuffle array elements up (lose oldest -- highest -- element)
            For iIndex = conDimension - 1 To 1 Step -1
                Set rsArray(iIndex + 1) = rsArray(iIndex)
            Next
            iIndex = 1 'the current recordset will be placed as the first element
        End If
        Set rsArray(iIndex) = rs        'buffer this recordset (which also gets returned)
    Else
        iIndex = conDimension
        ReDim rsArray(iIndex)           'this will cause "Nothing" to be returned
    End If
    
    Set ffm_tank = rsArray(iIndex)
    Exit Function
    
ffm_tankRedim:
    If Err.Number = 9 Then 'Subscript out of range
        ' this is the expected error on the FIRST call to this function; we
        ' must dimension the static array and retry the first reference to it.
        Err.Clear
        ReDim rsArray(conDimension)
        On Error GoTo 0
        Resume
    Else
        Err.Raise Err.Number
    End If
    
End Function    ''ffm_tank

Public Function ffm_tank_hist( _
                                sParm1 As Integer, _
                                sParm2 As String, _
                                sParm3 As Date, _
                                db As DataBase _
                                ) As Recordset
    Const SUB_NAME = "ffm_tank_hist"
    
    'returning pr_fm_tank_hist As Recordset
        
    Dim sSql As String
    Dim iIndex As Integer
    Dim rs As Recordset
    Static rsArray() As Recordset

    On Error GoTo ffm_tank_histRedim
    iIndex = 1
    Do While Not (rsArray(iIndex) Is Nothing)
        If rsArray(iIndex)!tkh_prft_ctr = sParm1 And _
           rsArray(iIndex)!tkh_num = sParm2 Then
                Set ffm_tank_hist = rsArray(iIndex)
                Exit Function   'return now with this "buffered" recordset
        End If
        iIndex = iIndex + 1
        If iIndex > conDimension Then Exit Do 'While  'Note: iIndex will be used below
    Loop 'While
    
    sSql = " SELECT * from fm_tank_hist " & _
           " WHERE tkh_prft_ctr = " & QStr(sParm1) & _
           " AND tkh_num = " & QStr(sParm2) & _
           " AND tkh_strt_date <= " & QStr(sParm3) & _
           " ORDER BY tkh_prft_ctr, tkh_num, tkh_strt_date"
    Set rs = fnMainmodsOpenRecordset(sSql, SUB_NAME, db)
    If rs Is Nothing Then
        Exit Function
    End If
    If rs.RecordCount > 0 Then
        rs.MoveFirst    'populate with first record
        #If cDEBUG Then
            On Error Resume Next
            While Not rs.EOF
                Debug.Print "ffm_tank_hist(" & QStr(sParm1) & ") = " & rs.Fields(0)
                rs.MoveNext    ' Locate next record
            Wend
            On Error GoTo 0
            rs.MoveFirst
        #End If
        If iIndex > conDimension Then
            ' shuffle array elements up (lose oldest -- highest -- element)
            For iIndex = conDimension - 1 To 1 Step -1
                Set rsArray(iIndex + 1) = rsArray(iIndex)
            Next
            iIndex = 1 'the current recordset will be placed as the first element
        End If
        Set rsArray(iIndex) = rs        'buffer this recordset (which also gets returned)
    Else
        iIndex = conDimension
        ReDim rsArray(iIndex)           'this will cause "Nothing" to be returned
    End If
    
    Set ffm_tank_hist = rsArray(iIndex)
    Exit Function
    
ffm_tank_histRedim:
    If Err.Number = 9 Then 'Subscript out of range
        ' this is the expected error on the FIRST call to this function; we
        ' must dimension the static array and retry the first reference to it.
        Err.Clear
        ReDim rsArray(conDimension)
        On Error GoTo 0
        Resume
    Else
        Err.Raise Err.Number
    End If
    
End Function    ''ffm_tank_hist

Public Function ffm_tank_lnk( _
                                sParm1 As Integer, _
                                sParm2 As String, _
                                sParm3 As Date, _
                                sParm4 As Long, _
                                db As DataBase _
                                ) As Recordset
    Const SUB_NAME = "ffm_tank_lnk"
    'returning pr_fm_tank_hist As Recordset
        
    Dim sSql As String, c_fmm_hist_lnk As String
    Dim iIndex As Integer
    Dim rs As Recordset
    Static rsArray() As Recordset
    Dim low_dte As Date

    On Error GoTo ffm_tank_lnkRedim
    iIndex = 1
    Do While Not (rsArray(iIndex) Is Nothing)
        If rsArray(iIndex)!tk_prft_ctr = sParm1 And _
           rsArray(iIndex)!tk_num = sParm2 Then
            If (rsArray(iIndex)!tk_strt_date <= sParm3) And _
               (rsArray(iIndex)!tk_prodlnk = sParm4) Then
                    Set ffm_tank_lnk = rsArray(iIndex)
                    Exit Function   'return now with this "buffered" recordset
            Else
                ' select max(tkh_strt_date) into " & QStr(low_dte)
                sSql = _
                        "select max(tkh_strt_date)" & _
                        " from fm_tank_hist" & _
                        " where tkh_strt_date < " & QStr(sParm3) & _
                        " and tkh_prft_ctr = " & QStr(sParm1) & _
                        " and tkh_num = " & QStr(sParm2)
                Dim xx As Recordset
                Set xx = fnMainmodsOpenRecordset(sSql, SUB_NAME, db)
                If xx Is Nothing Then
                    Exit Function
                End If
                If xx.RecordCount > 0 Then
                    xx.MoveFirst 'populate
                    low_dte = xx!tkh_strt_date
                    c_fmm_hist_lnk = "SELECT * from fm_tank_hist" & _
                                        " WHERE tkh_prft_ctr = " & QStr(sParm1) & _
                                        " AND tkh_num = " & QStr(sParm2) & _
                                        " AND tkh_strt_date between " & QStr(low_dte) & _
                                        " AND " & QStr(sParm3)
                    'For Each ( c_fmm_hist_lnk into pr_fm_tank   )
                    Dim pr_fm_tank As Recordset
                    Set pr_fm_tank = fnMainmodsOpenRecordset(c_fmm_hist_lnk, SUB_NAME, db)
                    If pr_fm_tank Is Nothing Then
                        Exit Function
                    End If
                    If pr_fm_tank.RecordCount > 0 Then
                        If (pr_fm_tank!tk_prodlnk = sParm4) Then
                                Set ffm_tank_lnk = pr_fm_tank: Exit Function
                        End If
                    End If
                End If
                Set ffm_tank_lnk = Nothing: Exit Function
            End If
        End If
        iIndex = iIndex + 1
        If iIndex > conDimension Then Exit Do 'While  'Note: iIndex will be used below
    Loop 'While
    
    sSql = " SELECT * from fm_tank " & _
           " WHERE tk_prft_ctr = " & QStr(sParm1) & _
           " AND tk_num = " & QStr(sParm2)
    
    Set rs = fnMainmodsOpenRecordset(sSql, SUB_NAME, db)
    If rs Is Nothing Then
        Exit Function
    End If
    If rs.RecordCount > 0 Then
        rs.MoveFirst    'populate with first record
        #If cDEBUG Then
            On Error Resume Next
            While Not rs.EOF
                Debug.Print "ffm_tank_lnk(" & QStr(sParm1) & ") = " & rs.Fields(0)
                rs.MoveNext    ' Locate next record
            Wend
            On Error GoTo 0
            rs.MoveFirst
        #End If
        If iIndex > conDimension Then
            ' shuffle array elements up (lose oldest -- highest -- element)
            For iIndex = conDimension - 1 To 1 Step -1
                Set rsArray(iIndex + 1) = rsArray(iIndex)
            Next
            iIndex = 1 'the current recordset will be placed as the first element
        End If
        Set rsArray(iIndex) = rs        'buffer this recordset (which also gets returned)
        
        If (rsArray(iIndex)!tk_strt_date <= sParm3) And _
           (rsArray(iIndex)!tk_prodlnk = sParm4) Then
                Set ffm_tank_lnk = rsArray(iIndex)
                Exit Function   'return now with this "buffered" recordset
        Else
            ' select max(tkh_strt_date) into " & QStr(low_dte)
            sSql = _
                    "select max(tkh_strt_date)" & _
                    " from fm_tank_hist" & _
                    " where tkh_strt_date < " & QStr(sParm3) & _
                    " and tkh_prft_ctr = " & QStr(sParm1) & _
                    " and tkh_num = " & QStr(sParm2)
            'Dim xx As Recordset
            Set xx = fnMainmodsOpenRecordset(sSql, SUB_NAME, db)
            If xx Is Nothing Then
                Exit Function
            End If
            If xx.RecordCount > 0 Then
                xx.MoveFirst 'populate
                low_dte = xx!tkh_strt_date
                c_fmm_hist_lnk = "SELECT * from fm_tank_hist" & _
                                    " WHERE tkh_prft_ctr = " & QStr(sParm1) & _
                                    " AND tkh_num = " & QStr(sParm2) & _
                                    " AND tkh_strt_date between " & QStr(low_dte) & _
                                    " AND " & QStr(sParm3)
                'For Each ( c_fmm_hist_lnk into pr_fm_tank   )
                'Dim pr_fm_tank As Recordset
                Set pr_fm_tank = fnMainmodsOpenRecordset(c_fmm_hist_lnk, SUB_NAME, db)
                If pr_fm_tank Is Nothing Then
                    Exit Function
                End If
                If pr_fm_tank.RecordCount > 0 Then
                    If (pr_fm_tank!tk_prodlnk = sParm4) Then
                            Set ffm_tank_lnk = pr_fm_tank: Exit Function
                    End If
                End If
            End If
            Set ffm_tank_lnk = Nothing: Exit Function
        End If
    
    Else
        iIndex = conDimension
        ReDim rsArray(iIndex)           'this will cause "Nothing" to be returned
    End If
    
    Set ffm_tank_lnk = rsArray(iIndex)
    Exit Function
    
ffm_tank_lnkRedim:
    If Err.Number = 9 Then 'Subscript out of range
        ' this is the expected error on the FIRST call to this function; we
        ' must dimension the static array and retry the first reference to it.
        Err.Clear
        ReDim rsArray(conDimension)
        On Error GoTo 0
        Resume
    Else
        Err.Raise Err.Number
    End If
    
End Function    ''ffm_tank_lnk


Public Function ffm_terminal( _
                                sParm1 As Long, _
                                db As DataBase _
                                ) As Recordset
    Const SUB_NAME = "ffm_terminal"
    'returning pr_fm_terminal As Recordset
        
    Dim sSql As String
    Dim iIndex As Integer
    Dim rs As Recordset
    Static rsArray() As Recordset

    On Error GoTo ffm_terminalRedim
    iIndex = 1
    Do While Not (rsArray(iIndex) Is Nothing)
        If rsArray(iIndex)!fmt_code = sParm1 Then
            Set ffm_terminal = rsArray(iIndex)
            Exit Function   'return now with this "buffered" recordset
        End If
        iIndex = iIndex + 1
        If iIndex > conDimension Then Exit Do 'While  'Note: iIndex will be used below
    Loop 'While
    
    sSql = "SELECT * FROM fm_terminal WHERE fmt_code = " & QStr(sParm1)
    Set rs = fnMainmodsOpenRecordset(sSql, SUB_NAME, db)
    If rs Is Nothing Then
        Exit Function
    End If
    If rs.RecordCount > 0 Then
        rs.MoveFirst    'populate with first record
        #If cDEBUG Then
            On Error Resume Next
            While Not rs.EOF
                Debug.Print "ffm_terminal(" & QStr(sParm1) & ") = " & rs.Fields(0)
                rs.MoveNext    ' Locate next record
            Wend
            On Error GoTo 0
            rs.MoveFirst
        #End If
        If iIndex > conDimension Then
            ' shuffle array elements up (lose oldest -- highest -- element)
            For iIndex = conDimension - 1 To 1 Step -1
                Set rsArray(iIndex + 1) = rsArray(iIndex)
            Next
            iIndex = 1 'the current recordset will be placed as the first element
        End If
        Set rsArray(iIndex) = rs        'buffer this recordset (which also gets returned)
    Else
        iIndex = conDimension
        ReDim rsArray(iIndex)           'this will cause "Nothing" to be returned
    End If
    
    Set ffm_terminal = rsArray(iIndex)
    Exit Function
    
ffm_terminalRedim:
    If Err.Number = 9 Then 'Subscript out of range
        ' this is the expected error on the FIRST call to this function; we
        ' must dimension the static array and retry the first reference to it.
        Err.Clear
        ReDim rsArray(conDimension)
        On Error GoTo 0
        Resume
    Else
        Err.Raise Err.Number
    End If
    
End Function    ''ffm_terminal

Public Function ffo_class( _
                                sParm1 As String, _
                                db As DataBase _
                                ) As Recordset
    Const SUB_NAME = "ffo_class"

    'returning pr_fo_class As Recordset
        
    Dim sSql As String
    Dim iIndex As Integer
    Dim rs As Recordset
    Static rsArray() As Recordset

    On Error GoTo ffo_classRedim
    iIndex = 1
    Do While Not (rsArray(iIndex) Is Nothing)
        If rsArray(iIndex)!focl_class_code = sParm1 Then
            Set ffo_class = rsArray(iIndex)
            Exit Function   'return now with this "buffered" recordset
        End If
        iIndex = iIndex + 1
        If iIndex > conDimension Then Exit Do 'While  'Note: iIndex will be used below
    Loop 'While
    
    sSql = "SELECT * FROM fo_class WHERE focl_class_code = " & QStr(sParm1)
    Set rs = fnMainmodsOpenRecordset(sSql, SUB_NAME, db)
    If rs Is Nothing Then
        Exit Function
    End If
    If rs.RecordCount > 0 Then
        rs.MoveFirst    'populate with first record
        #If cDEBUG Then
            On Error Resume Next
            While Not rs.EOF
                Debug.Print "ffo_class(" & QStr(sParm1) & ") = " & rs.Fields(0)
                rs.MoveNext    ' Locate next record
            Wend
            On Error GoTo 0
            rs.MoveFirst
        #End If
        If iIndex > conDimension Then
            ' shuffle array elements up (lose oldest -- highest -- element)
            For iIndex = conDimension - 1 To 1 Step -1
                Set rsArray(iIndex + 1) = rsArray(iIndex)
            Next
            iIndex = 1 'the current recordset will be placed as the first element
        End If
        Set rsArray(iIndex) = rs        'buffer this recordset (which also gets returned)
    Else
        iIndex = conDimension
        ReDim rsArray(iIndex)           'this will cause "Nothing" to be returned
    End If
    
    Set ffo_class = rsArray(iIndex)
    Exit Function
    
ffo_classRedim:
    If Err.Number = 9 Then 'Subscript out of range
        ' this is the expected error on the FIRST call to this function; we
        ' must dimension the static array and retry the first reference to it.
        Err.Clear
        ReDim rsArray(conDimension)
        On Error GoTo 0
        Resume
    Else
        Err.Raise Err.Number
    End If
    
End Function    ''ffo_class

Public Function fgf_header( _
                                sParm1 As String, _
                                db As DataBase _
                                ) As Recordset
    Const SUB_NAME = "fgf_header"
    
    'returning pr_gf_header As Recordset
        
    Dim sSql As String
    Dim iIndex As Integer
    Dim rs As Recordset
    Static rsArray() As Recordset

    On Error GoTo fgf_headerRedim
    iIndex = 1
    Do While Not (rsArray(iIndex) Is Nothing)
        If rsArray(iIndex)!gfh_path_code = sParm1 Then
            Set fgf_header = rsArray(iIndex)
            Exit Function   'return now with this "buffered" recordset
        End If
        iIndex = iIndex + 1
        If iIndex > conDimension Then Exit Do 'While  'Note: iIndex will be used below
    Loop 'While
    
    sSql = "SELECT * FROM gf_header WHERE gfh_path_code = " & QStr(sParm1)
    Set rs = fnMainmodsOpenRecordset(sSql, SUB_NAME, db)
    If rs Is Nothing Then
        Exit Function
    End If
    If rs.RecordCount > 0 Then
        rs.MoveFirst    'populate with first record
        #If cDEBUG Then
            On Error Resume Next
            While Not rs.EOF
                Debug.Print "fgf_header(" & QStr(sParm1) & ") = " & rs.Fields(0)
                rs.MoveNext    ' Locate next record
            Wend
            On Error GoTo 0
            rs.MoveFirst
        #End If
        If iIndex > conDimension Then
            ' shuffle array elements up (lose oldest -- highest -- element)
            For iIndex = conDimension - 1 To 1 Step -1
                Set rsArray(iIndex + 1) = rsArray(iIndex)
            Next
            iIndex = 1 'the current recordset will be placed as the first element
        End If
        Set rsArray(iIndex) = rs        'buffer this recordset (which also gets returned)
    Else
        iIndex = conDimension
        ReDim rsArray(iIndex)           'this will cause "Nothing" to be returned
    End If
    
    Set fgf_header = rsArray(iIndex)
    Exit Function
    
fgf_headerRedim:
    If Err.Number = 9 Then 'Subscript out of range
        ' this is the expected error on the FIRST call to this function; we
        ' must dimension the static array and retry the first reference to it.
        Err.Clear
        ReDim rsArray(conDimension)
        On Error GoTo 0
        Resume
    Else
        Err.Raise Err.Number
    End If
    
End Function    ''fgf_header

Public Function fgl_master( _
                                sParm1 As Variant, _
                                sParm2 As Long, _
                                sParm3 As Variant, _
                                db As DataBase _
                                ) As Recordset
    Const SUB_NAME = "fgl_master"
    
    'returning pr_gl_master As Recordset
        
    Dim sSql As String
    Dim iIndex As Integer
    Dim rs As Recordset
    Dim pprd As Recordset
    Static rsArray() As Recordset
    
    If IsNull(sParm3) Then
            sParm3 = mastr_pc()
            If IsNull(sParm3) Then sParm3 = set_mastr_master(db)
    End If
    If IsNull(sParm1) Then
            'wrk_date = Date
            Set pprd = fgl_period(Date, db) 'wrk_date
            If Not (pprd Is Nothing) Then sParm1 = pprd!glp_series
    End If

    On Error GoTo fgl_masterRedim
    iIndex = 1
    Do While Not (rsArray(iIndex) Is Nothing)
        If rsArray(iIndex)!glm_series = sParm1 And _
           rsArray(iIndex)!glm_account = sParm2 And _
           rsArray(iIndex)!glm_prft_ctr = sParm3 Then
            Set fgl_master = rsArray(iIndex)
            Exit Function   'return now with this "buffered" recordset
        End If
        iIndex = iIndex + 1
        If iIndex > conDimension Then Exit Do 'While  'Note: iIndex will be used below
    Loop 'While
    
    sSql = "SELECT * FROM gl_master WHERE glm_series = " & QStr(sParm1) & _
           " AND  glm_account = " & QStr(sParm2) & _
           " AND glm_prft_ctr = " & QStr(sParm3)
    Set rs = fnMainmodsOpenRecordset(sSql, SUB_NAME, db)
    If rs Is Nothing Then
        Exit Function
    End If
    If rs.RecordCount > 0 Then
        rs.MoveFirst    'populate with first record
        #If cDEBUG Then
            On Error Resume Next
            While Not rs.EOF
                Debug.Print "fgl_master(" & QStr(sParm1) & ") = " & rs.Fields(0)
                rs.MoveNext    ' Locate next record
            Wend
            On Error GoTo 0
            rs.MoveFirst
        #End If
        If iIndex > conDimension Then
            ' shuffle array elements up (lose oldest -- highest -- element)
            For iIndex = conDimension - 1 To 1 Step -1
                Set rsArray(iIndex + 1) = rsArray(iIndex)
            Next
            iIndex = 1 'the current recordset will be placed as the first element
        End If
        Set rsArray(iIndex) = rs        'buffer this recordset (which also gets returned)
    Else
        iIndex = conDimension
        ReDim rsArray(iIndex)           'this will cause "Nothing" to be returned
    End If
    
    Set fgl_master = rsArray(iIndex)
    Exit Function
    
fgl_masterRedim:
    If Err.Number = 9 Then 'Subscript out of range
        ' this is the expected error on the FIRST call to this function; we
        ' must dimension the static array and retry the first reference to it.
        Err.Clear
        ReDim rsArray(conDimension)
        On Error GoTo 0
        Resume
    Else
        Err.Raise Err.Number
    End If
    
End Function    ''fgl_master

Public Function fgl_period( _
                                sParm1 As Date, _
                                db As DataBase _
                                ) As Recordset
    Const SUB_NAME = "fgl_period"
    
    'returning pr_gl_period As Recordset
        
    Dim sSql As String
    Dim iIndex As Integer
    Dim rs As Recordset
    Static rsArray() As Recordset

    On Error GoTo fgl_periodRedim
    iIndex = 1
    Do While Not (rsArray(iIndex) Is Nothing)
        If rsArray(iIndex)!glp_beg_dt <= sParm1 And _
           rsArray(iIndex)!glp_end_dt >= sParm1 Then
            Set fgl_period = rsArray(iIndex)
            Exit Function   'return now with this "buffered" recordset
        End If
        iIndex = iIndex + 1
        If iIndex > conDimension Then Exit Do 'While  'Note: iIndex will be used below
    Loop 'While
    
    sSql = "SELECT * FROM gl_period WHERE " & QStr(sParm1) & " between glp_beg_dt and glp_end_dt"
    Set rs = fnMainmodsOpenRecordset(sSql, SUB_NAME, db)
    If rs Is Nothing Then
        Exit Function
    End If
    If rs.RecordCount > 0 Then
        rs.MoveFirst    'populate with first record
        #If cDEBUG Then
            On Error Resume Next
            While Not rs.EOF
                Debug.Print "fgl_period(" & QStr(sParm1) & ") = " & rs.Fields(0)
                rs.MoveNext    ' Locate next record
            Wend
            On Error GoTo 0
            rs.MoveFirst
        #End If
        If iIndex > conDimension Then
            ' shuffle array elements up (lose oldest -- highest -- element)
            For iIndex = conDimension - 1 To 1 Step -1
                Set rsArray(iIndex + 1) = rsArray(iIndex)
            Next
            iIndex = 1 'the current recordset will be placed as the first element
        End If
        Set rsArray(iIndex) = rs        'buffer this recordset (which also gets returned)
    Else
        iIndex = conDimension
        ReDim rsArray(iIndex)           'this will cause "Nothing" to be returned
    End If
    
    Set fgl_period = rsArray(iIndex)
    Exit Function
    
fgl_periodRedim:
    If Err.Number = 9 Then 'Subscript out of range
        ' this is the expected error on the FIRST call to this function; we
        ' must dimension the static array and retry the first reference to it.
        Err.Clear
        ReDim rsArray(conDimension)
        On Error GoTo 0
        Resume
    Else
        Err.Raise Err.Number
    End If
    
End Function    ''fgl_period


Public Function fp_altname( _
                                sParm1 As Long, _
                                db As DataBase _
                                ) As Recordset
    Const SUB_NAME = "fp_altname"
    'returning pr_p_altname As Recordset
        
    Dim sSql As String
    Dim iIndex As Integer
    Dim rs As Recordset
    Static rsArray() As Recordset

    On Error GoTo fp_altnameRedim
    iIndex = 1
    Do While Not (rsArray(iIndex) Is Nothing)
        If rsArray(iIndex)!pn_alt = sParm1 Then
            Set fp_altname = rsArray(iIndex)
            Exit Function   'return now with this "buffered" recordset
        End If
        iIndex = iIndex + 1
        If iIndex > conDimension Then Exit Do 'While  'Note: iIndex will be used below
    Loop 'While
    
    sSql = "SELECT * FROM p_altname WHERE pn_alt = " & QStr(sParm1)
    Set rs = fnMainmodsOpenRecordset(sSql, SUB_NAME, db)
    If rs Is Nothing Then
        Exit Function
    End If
    If rs.RecordCount > 0 Then
        rs.MoveFirst    'populate with first record
        #If cDEBUG Then
            On Error Resume Next
            While Not rs.EOF
                Debug.Print "fp_altname(" & QStr(sParm1) & ") = " & rs.Fields(0)
                rs.MoveNext    ' Locate next record
            Wend
            On Error GoTo 0
            rs.MoveFirst
        #End If
        If iIndex > conDimension Then
            ' shuffle array elements up (lose oldest -- highest -- element)
            For iIndex = conDimension - 1 To 1 Step -1
                Set rsArray(iIndex + 1) = rsArray(iIndex)
            Next
            iIndex = 1 'the current recordset will be placed as the first element
        End If
        Set rsArray(iIndex) = rs        'buffer this recordset (which also gets returned)
    Else
        iIndex = conDimension
        ReDim rsArray(iIndex)           'this will cause "Nothing" to be returned
    End If
    
    Set fp_altname = rsArray(iIndex)
    Exit Function
    
fp_altnameRedim:
    If Err.Number = 9 Then 'Subscript out of range
        ' this is the expected error on the FIRST call to this function; we
        ' must dimension the static array and retry the first reference to it.
        Err.Clear
        ReDim rsArray(conDimension)
        On Error GoTo 0
        Resume
    Else
        Err.Raise Err.Number
    End If
    
End Function    ''fp_altname

Public Function fp_class( _
                                sParm1 As String, _
                                db As DataBase _
                                ) As Recordset
    Const SUB_NAME = "fp_class"
    'returning pr_p_class As Recordset
        
    Dim sSql As String
    Dim iIndex As Integer
    Dim rs As Recordset
    Static rsArray() As Recordset

    On Error GoTo fp_classRedim
    iIndex = 1
    Do While Not (rsArray(iIndex) Is Nothing)
        If rsArray(iIndex)!pcl_code = sParm1 Then
            Set fp_class = rsArray(iIndex)
            Exit Function   'return now with this "buffered" recordset
        End If
        iIndex = iIndex + 1
        If iIndex > conDimension Then Exit Do 'While  'Note: iIndex will be used below
    Loop 'While
    
    sSql = "SELECT * FROM p_class WHERE pcl_code = " & QStr(sParm1)
    Set rs = fnMainmodsOpenRecordset(sSql, SUB_NAME, db)
    If rs Is Nothing Then
        Exit Function
    End If
    If rs.RecordCount > 0 Then
        rs.MoveFirst    'populate with first record
        #If cDEBUG Then
            On Error Resume Next
            While Not rs.EOF
                Debug.Print "fp_class(" & QStr(sParm1) & ") = " & rs.Fields(0)
                rs.MoveNext    ' Locate next record
            Wend
            On Error GoTo 0
            rs.MoveFirst
        #End If
        If iIndex > conDimension Then
            ' shuffle array elements up (lose oldest -- highest -- element)
            For iIndex = conDimension - 1 To 1 Step -1
                Set rsArray(iIndex + 1) = rsArray(iIndex)
            Next
            iIndex = 1 'the current recordset will be placed as the first element
        End If
        Set rsArray(iIndex) = rs        'buffer this recordset (which also gets returned)
    Else
        iIndex = conDimension
        ReDim rsArray(iIndex)           'this will cause "Nothing" to be returned
    End If
    
    Set fp_class = rsArray(iIndex)
    Exit Function
    
fp_classRedim:
    If Err.Number = 9 Then 'Subscript out of range
        ' this is the expected error on the FIRST call to this function; we
        ' must dimension the static array and retry the first reference to it.
        Err.Clear
        ReDim rsArray(conDimension)
        On Error GoTo 0
        Resume
    Else
        Err.Raise Err.Number
    End If
    
End Function    ''fp_class

Public Function fprodlnk( _
                                sParm1 As String, _
                                db As DataBase _
                                ) As Recordset
    Const SUB_NAME = "fprodlnk"
    'returning pr_inv_header As Recordset
        
    Dim sSql As String
    Dim iIndex As Integer
    Dim rs As Recordset
    Static rsArray() As Recordset

    On Error GoTo fprodlnkRedim
    iIndex = 1
    Do While Not (rsArray(iIndex) Is Nothing)
        If rsArray(iIndex)!ivh_product = sParm1 Then
            Set fprodlnk = rsArray(iIndex)
            Exit Function   'return now with this "buffered" recordset
        End If
        iIndex = iIndex + 1
        If iIndex > conDimension Then Exit Do 'While  'Note: iIndex will be used below
    Loop 'While
    
    sSql = "SELECT * FROM inv_header WHERE ivh_product = " & QStr(sParm1)
    Set rs = fnMainmodsOpenRecordset(sSql, SUB_NAME, db)
    If rs Is Nothing Then
        Exit Function
    End If
    If rs.RecordCount > 0 Then
        rs.MoveFirst    'populate with first record
        #If cDEBUG Then
            On Error Resume Next
            While Not rs.EOF
                Debug.Print "fprodlnk(" & QStr(sParm1) & ") = " & rs.Fields(0)
                rs.MoveNext    ' Locate next record
            Wend
            On Error GoTo 0
            rs.MoveFirst
        #End If
        If iIndex > conDimension Then
            ' shuffle array elements up (lose oldest -- highest -- element)
            For iIndex = conDimension - 1 To 1 Step -1
                Set rsArray(iIndex + 1) = rsArray(iIndex)
            Next
            iIndex = 1 'the current recordset will be placed as the first element
        End If
        Set rsArray(iIndex) = rs        'buffer this recordset (which also gets returned)
    Else
        iIndex = conDimension
        ReDim rsArray(iIndex)           'this will cause "Nothing" to be returned
    End If
    
    Set fprodlnk = rsArray(iIndex)
    Exit Function
    
fprodlnkRedim:
    If Err.Number = 9 Then 'Subscript out of range
        ' this is the expected error on the FIRST call to this function; we
        ' must dimension the static array and retry the first reference to it.
        Err.Clear
        ReDim rsArray(conDimension)
        On Error GoTo 0
        Resume
    Else
        Err.Raise Err.Number
    End If
    
End Function    ''fprodlnk

Public Function fproduct( _
                                sParm1 As Long, _
                                db As DataBase _
                                ) As Recordset
    Const SUB_NAME = "fproduct"
    
    'returning pr_inv_header As Recordset
        
    Dim sSql As String
    Dim iIndex As Integer
    Dim rs As Recordset
    Static rsArray() As Recordset

AGAIN:

    On Error GoTo fproductRedim
    iIndex = 1
    Do While Not (rsArray(iIndex) Is Nothing)
        If rsArray(iIndex)!ivh_link = sParm1 Then
            Set fproduct = rsArray(iIndex)
            Exit Function   'return now with this "buffered" recordset
        End If
        iIndex = iIndex + 1
        If iIndex > conDimension Then Exit Do 'While  'Note: iIndex will be used below
    Loop 'While
    
    sSql = "SELECT * FROM inv_header WHERE ivh_link = " & QStr(sParm1)
    Set rs = fnMainmodsOpenRecordset(sSql, SUB_NAME, db)
    If rs Is Nothing Then
        Exit Function
    End If
    If rs.RecordCount > 0 Then
        rs.MoveFirst    'populate with first record
        #If cDEBUG Then
            On Error Resume Next
            While Not rs.EOF
                Debug.Print "fproduct(" & QStr(sParm1) & ") = " & rs.Fields(0)
                rs.MoveNext    ' Locate next record
            Wend
            On Error GoTo 0
            rs.MoveFirst
        #End If
        If iIndex > conDimension Then
            ' shuffle array elements up (lose oldest -- highest -- element)
            For iIndex = conDimension - 1 To 1 Step -1
                Set rsArray(iIndex + 1) = rsArray(iIndex)
            Next
            iIndex = 1 'the current recordset will be placed as the first element
        End If
        Set rsArray(iIndex) = rs        'buffer this recordset (which also gets returned)
    Else
        iIndex = conDimension
        ReDim rsArray(iIndex)           'this will cause "Nothing" to be returned
    End If
    
    Set fproduct = rsArray(iIndex)
    Exit Function
    
fproductRedim:
    If Err.Number = 9 Then 'Subscript out of range
        ' this is the expected error on the FIRST call to this function; we
        ' must dimension the static array and retry the first reference to it.
        Err.Clear
        ReDim rsArray(conDimension)
        On Error GoTo 0
        Resume
    Else
        If Err.Number = 3420 Then 'Object invalid or no longer set.
        Err.Clear
        ReDim rsArray(conDimension)
        On Error GoTo 0
        Resume AGAIN
        Else
            Err.Raise Err.Number
        End If
    End If
    
End Function    ''fproduct

Public Function fp_vendor( _
                                sParm1 As Long, _
                                db As DataBase _
                                ) As Recordset
    Const SUB_NAME = "fp_vendor"
    'returning pr_p_vendor As Recordset
        
    Dim sSql As String
    Dim iIndex As Integer
    Dim rs As Recordset
    Static rsArray() As Recordset

    On Error GoTo fp_vendorRedim
    iIndex = 1
    Do While Not (rsArray(iIndex) Is Nothing)
        If rsArray(iIndex)!pm_vendor = sParm1 Then
            Set fp_vendor = rsArray(iIndex)
            Exit Function   'return now with this "buffered" recordset
        End If
        iIndex = iIndex + 1
        If iIndex > conDimension Then Exit Do 'While  'Note: iIndex will be used below
    Loop 'While
    
    sSql = "SELECT * FROM p_vendor WHERE pm_vendor = " & QStr(sParm1)
    Set rs = fnMainmodsOpenRecordset(sSql, SUB_NAME, db)
    If rs Is Nothing Then
        Exit Function
    End If
    If rs.RecordCount > 0 Then
        rs.MoveFirst    'populate with first record
        #If cDEBUG Then
            On Error Resume Next
            While Not rs.EOF
                Debug.Print "fp_vendor(" & QStr(sParm1) & ") = " & rs.Fields(0)
                rs.MoveNext    ' Locate next record
            Wend
            On Error GoTo 0
            rs.MoveFirst
        #End If
        If iIndex > conDimension Then
            ' shuffle array elements up (lose oldest -- highest -- element)
            For iIndex = conDimension - 1 To 1 Step -1
                Set rsArray(iIndex + 1) = rsArray(iIndex)
            Next
            iIndex = 1 'the current recordset will be placed as the first element
        End If
        Set rsArray(iIndex) = rs        'buffer this recordset (which also gets returned)
    Else
        iIndex = conDimension
        ReDim rsArray(iIndex)           'this will cause "Nothing" to be returned
    End If
    
    Set fp_vendor = rsArray(iIndex)
    Exit Function
    
fp_vendorRedim:
    If Err.Number = 9 Then 'Subscript out of range
        ' this is the expected error on the FIRST call to this function; we
        ' must dimension the static array and retry the first reference to it.
        Err.Clear
        ReDim rsArray(conDimension)
        On Error GoTo 0
        Resume
    Else
        Err.Raise Err.Number
    End If
    
End Function    ''fp_vendor

Public Function frs_b_period( _
                                sParm1 As Date, _
                                db As DataBase _
                                ) As Recordset
    Const SUB_NAME = "frs_b_period"
    'returning pr_rs_b_period As Recordset
        
    Dim sSql As String
    Dim iIndex As Integer
    Dim rs As Recordset
    Static rsArray() As Recordset

    On Error GoTo frs_b_periodRedim
    iIndex = 1
    Do While Not (rsArray(iIndex) Is Nothing)
        If rsArray(iIndex)!rsbw_start_date <= sParm1 And _
           rsArray(iIndex)!rsbw_end_date >= sParm1 Then
            Set frs_b_period = rsArray(iIndex)
            Exit Function   'return now with this "buffered" recordset
        End If
        iIndex = iIndex + 1
        If iIndex > conDimension Then Exit Do 'While  'Note: iIndex will be used below
    Loop 'While
    
    sSql = "SELECT * FROM rs_b_period WHERE " & QStr(sParm1) & " between rsbw_start_date and rsbw_end_date"
    Set rs = fnMainmodsOpenRecordset(sSql, SUB_NAME, db)
    If rs Is Nothing Then
        Exit Function
    End If
    If rs.RecordCount > 0 Then
        rs.MoveFirst    'populate with first record
        #If cDEBUG Then
            On Error Resume Next
            While Not rs.EOF
                Debug.Print "frs_b_period(" & QStr(sParm1) & ") = " & rs.Fields(0)
                rs.MoveNext    ' Locate next record
            Wend
            On Error GoTo 0
            rs.MoveFirst
        #End If
        If iIndex > conDimension Then
            ' shuffle array elements up (lose oldest -- highest -- element)
            For iIndex = conDimension - 1 To 1 Step -1
                Set rsArray(iIndex + 1) = rsArray(iIndex)
            Next
            iIndex = 1 'the current recordset will be placed as the first element
        End If
        Set rsArray(iIndex) = rs        'buffer this recordset (which also gets returned)
    Else
        iIndex = conDimension
        ReDim rsArray(iIndex)           'this will cause "Nothing" to be returned
    End If
    
    Set frs_b_period = rsArray(iIndex)
    Exit Function
    
frs_b_periodRedim:
    If Err.Number = 9 Then 'Subscript out of range
        ' this is the expected error on the FIRST call to this function; we
        ' must dimension the static array and retry the first reference to it.
        Err.Clear
        ReDim rsArray(conDimension)
        On Error GoTo 0
        Resume
    Else
        Err.Raise Err.Number
    End If
    
End Function    ''frs_b_period

Public Function frs_cat( _
                                sParm1 As String, _
                                db As DataBase _
                                ) As Recordset
    Const SUB_NAME = "frs_cat"
    
    'returning pr_rs_cat As Recordset
        
    Dim sSql As String
    Dim iIndex As Integer
    Dim rs As Recordset
    Static rsArray() As Recordset

    On Error GoTo frs_catRedim
    iIndex = 1
    Do While Not (rsArray(iIndex) Is Nothing)
        If rsArray(iIndex)!rscta_catagory = sParm1 Then
            Set frs_cat = rsArray(iIndex)
            Exit Function   'return now with this "buffered" recordset
        End If
        iIndex = iIndex + 1
        If iIndex > conDimension Then Exit Do 'While  'Note: iIndex will be used below
    Loop 'While
    
    sSql = "SELECT * FROM rs_cat WHERE rscta_catagory = " & QStr(sParm1)
    Set rs = fnMainmodsOpenRecordset(sSql, SUB_NAME, db)
    If rs Is Nothing Then
        Exit Function
    End If
    If rs.RecordCount > 0 Then
        rs.MoveFirst    'populate with first record
        #If cDEBUG Then
            On Error Resume Next
            While Not rs.EOF
                Debug.Print "frs_cat(" & QStr(sParm1) & ") = " & rs.Fields(0)
                rs.MoveNext    ' Locate next record
            Wend
            On Error GoTo 0
            rs.MoveFirst
        #End If
        If iIndex > conDimension Then
            ' shuffle array elements up (lose oldest -- highest -- element)
            For iIndex = conDimension - 1 To 1 Step -1
                Set rsArray(iIndex + 1) = rsArray(iIndex)
            Next
            iIndex = 1 'the current recordset will be placed as the first element
        End If
        Set rsArray(iIndex) = rs        'buffer this recordset (which also gets returned)
    Else
        iIndex = conDimension
        ReDim rsArray(iIndex)           'this will cause "Nothing" to be returned
    End If
    
    Set frs_cat = rsArray(iIndex)
    Exit Function
    
frs_catRedim:
    If Err.Number = 9 Then 'Subscript out of range
        ' this is the expected error on the FIRST call to this function; we
        ' must dimension the static array and retry the first reference to it.
        Err.Clear
        ReDim rsArray(conDimension)
        On Error GoTo 0
        Resume
    Else
        Err.Raise Err.Number
    End If
    
End Function    ''frs_cat

Public Function frs_alt_cat( _
                                sParm1 As String, _
                                db As DataBase _
                                ) As Recordset
    Const SUB_NAME = "frs_alt_cat"
    'returning pr_rs_alt_cat As Recordset
        
    Dim sSql As String
    Dim iIndex As Integer
    Dim rs As Recordset
    Static rsArray() As Recordset

    On Error GoTo frs_alt_catRedim
    iIndex = 1
    Do While Not (rsArray(iIndex) Is Nothing)
        If rsArray(iIndex)!rsct_catagory = sParm1 Then
            Set frs_alt_cat = rsArray(iIndex)
            Exit Function   'return now with this "buffered" recordset
        End If
        iIndex = iIndex + 1
        If iIndex > conDimension Then Exit Do 'While  'Note: iIndex will be used below
    Loop 'While
    
    sSql = "SELECT * FROM rs_alt_cat WHERE rsct_catagory = " & QStr(sParm1)
    Set rs = fnMainmodsOpenRecordset(sSql, SUB_NAME, db)
    If rs Is Nothing Then
        Exit Function
    End If
    If rs.RecordCount > 0 Then
        rs.MoveFirst    'populate with first record
        #If cDEBUG Then
            On Error Resume Next
            While Not rs.EOF
                Debug.Print "frs_alt_cat(" & QStr(sParm1) & ") = " & rs.Fields(0)
                rs.MoveNext    ' Locate next record
            Wend
            On Error GoTo 0
            rs.MoveFirst
        #End If
        If iIndex > conDimension Then
            ' shuffle array elements up (lose oldest -- highest -- element)
            For iIndex = conDimension - 1 To 1 Step -1
                Set rsArray(iIndex + 1) = rsArray(iIndex)
            Next
            iIndex = 1 'the current recordset will be placed as the first element
        End If
        Set rsArray(iIndex) = rs        'buffer this recordset (which also gets returned)
    Else
        iIndex = conDimension
        ReDim rsArray(iIndex)           'this will cause "Nothing" to be returned
    End If
    
    Set frs_alt_cat = rsArray(iIndex)
    Exit Function
    
frs_alt_catRedim:
    If Err.Number = 9 Then 'Subscript out of range
        ' this is the expected error on the FIRST call to this function; we
        ' must dimension the static array and retry the first reference to it.
        Err.Clear
        ReDim rsArray(conDimension)
        On Error GoTo 0
        Resume
    Else
        Err.Raise Err.Number
    End If
    
End Function    ''frs_alt_cat

Public Function finv_classes( _
                                sParm1 As String, _
                                sParm2 As String, _
                                db As DataBase _
                                ) As Recordset
    Const SUB_NAME = "finv_classes"
    
    'returning pr_inv_classes As Recordset
        
    Dim sSql As String
    Dim iIndex As Integer
    Dim rs As Recordset
    Static rsArray() As Recordset

    On Error GoTo finv_classesRedim
    iIndex = 1
    Do While Not (rsArray(iIndex) Is Nothing)
        If rsArray(iIndex)!icl_code = sParm1 And _
           rsArray(iIndex)!icl_type = sParm2 Then
            Set finv_classes = rsArray(iIndex)
            Exit Function   'return now with this "buffered" recordset
        End If
        iIndex = iIndex + 1
        If iIndex > conDimension Then Exit Do 'While  'Note: iIndex will be used below
    Loop 'While
    
    sSql = "SELECT * FROM inv_classes WHERE icl_code = " & QStr(sParm1) & _
            " AND icl_type = " & QStr(sParm2)
    Set rs = fnMainmodsOpenRecordset(sSql, SUB_NAME, db)
    If rs Is Nothing Then
        Exit Function
    End If
    If rs.RecordCount > 0 Then
        rs.MoveFirst    'populate with first record
        #If cDEBUG Then
            On Error Resume Next
            While Not rs.EOF
                Debug.Print "finv_classes(" & QStr(sParm1) & ") = " & rs.Fields(0)
                rs.MoveNext    ' Locate next record
            Wend
            On Error GoTo 0
            rs.MoveFirst
        #End If
        If iIndex > conDimension Then
            ' shuffle array elements up (lose oldest -- highest -- element)
            For iIndex = conDimension - 1 To 1 Step -1
                Set rsArray(iIndex + 1) = rsArray(iIndex)
            Next
            iIndex = 1 'the current recordset will be placed as the first element
        End If
        Set rsArray(iIndex) = rs        'buffer this recordset (which also gets returned)
    Else
        iIndex = conDimension
        ReDim rsArray(iIndex)           'this will cause "Nothing" to be returned
    End If
    
    Set finv_classes = rsArray(iIndex)
    Exit Function
    
finv_classesRedim:
    If Err.Number = 9 Then 'Subscript out of range
        ' this is the expected error on the FIRST call to this function; we
        ' must dimension the static array and retry the first reference to it.
        Err.Clear
        ReDim rsArray(conDimension)
        On Error GoTo 0
        Resume
    Else
        Err.Raise Err.Number
    End If
    
End Function    ''finv_classes

Public Function finv_header( _
                                sParm1 As String, _
                                db As DataBase _
                                ) As Recordset
    Const SUB_NAME = "finv_header"
    
    'returning pr_inv_header As Recordset
        
    Dim sSql As String
    Dim iIndex As Integer
    Dim rs As Recordset
    Static rsArray() As Recordset

    On Error GoTo finv_headerRedim
    iIndex = 1
    Do While Not (rsArray(iIndex) Is Nothing)
        If rsArray(iIndex)!ivh_product = sParm1 Then
            Set finv_header = rsArray(iIndex)
            Exit Function   'return now with this "buffered" recordset
        End If
        iIndex = iIndex + 1
        If iIndex > conDimension Then Exit Do 'While  'Note: iIndex will be used below
    Loop 'While
    
    sSql = "SELECT * FROM inv_header WHERE ivh_product = " & QStr(sParm1)
    Set rs = fnMainmodsOpenRecordset(sSql, SUB_NAME, db)
    If rs Is Nothing Then
        Exit Function
    End If
    If rs.RecordCount > 0 Then
        rs.MoveFirst    'populate with first record
        #If cDEBUG Then
            On Error Resume Next
            While Not rs.EOF
                Debug.Print "finv_header(" & QStr(sParm1) & ") = " & rs.Fields(0)
                rs.MoveNext    ' Locate next record
            Wend
            On Error GoTo 0
            rs.MoveFirst
        #End If
        If iIndex > conDimension Then
            ' shuffle array elements up (lose oldest -- highest -- element)
            For iIndex = conDimension - 1 To 1 Step -1
                Set rsArray(iIndex + 1) = rsArray(iIndex)
            Next
            iIndex = 1 'the current recordset will be placed as the first element
        End If
        Set rsArray(iIndex) = rs        'buffer this recordset (which also gets returned)
    Else
        iIndex = conDimension
        ReDim rsArray(iIndex)           'this will cause "Nothing" to be returned
    End If
    
    Set finv_header = rsArray(iIndex)
    Exit Function
    
finv_headerRedim:
    If Err.Number = 9 Then 'Subscript out of range
        ' this is the expected error on the FIRST call to this function; we
        ' must dimension the static array and retry the first reference to it.
        Err.Clear
        ReDim rsArray(conDimension)
        On Error GoTo 0
        Resume
    Else
        Err.Raise Err.Number
    End If
    
End Function    ''finv_header

Public Function finv_master( _
                                sParm1 As Long, _
                                sParm2 As Long, _
                                db As DataBase _
                                ) As Recordset
    Const SUB_NAME = "finv_master"
    'returning pr_inv_master As Recordset
        
    Dim sSql As String
    Dim iIndex As Integer
    Dim rs As Recordset
    Static rsArray() As Recordset

    On Error GoTo finv_masterRedim
    iIndex = 1
    Do While Not (rsArray(iIndex) Is Nothing)
        If rsArray(iIndex)!invt_prodlnk = sParm1 And _
           rsArray(iIndex)!invt_prft_ctr = sParm2 Then
            Set finv_master = rsArray(iIndex)
            Exit Function   'return now with this "buffered" recordset
        End If
        iIndex = iIndex + 1
        If iIndex > conDimension Then Exit Do 'While  'Note: iIndex will be used below
    Loop 'While
    
    sSql = "SELECT * FROM inv_master WHERE invt_prodlnk = " & QStr(sParm1) & _
            " AND invt_prft_ctr = " & QStr(sParm2)
    Set rs = fnMainmodsOpenRecordset(sSql, SUB_NAME, db)
    If rs Is Nothing Then
        Exit Function
    End If
    If rs.RecordCount > 0 Then
        rs.MoveFirst    'populate with first record
        #If cDEBUG Then
            On Error Resume Next
            While Not rs.EOF
                Debug.Print "finv_master(" & QStr(sParm1) & ") = " & rs.Fields(0)
                rs.MoveNext    ' Locate next record
            Wend
            On Error GoTo 0
            rs.MoveFirst
        #End If
        If iIndex > conDimension Then
            ' shuffle array elements up (lose oldest -- highest -- element)
            For iIndex = conDimension - 1 To 1 Step -1
                Set rsArray(iIndex + 1) = rsArray(iIndex)
            Next
            iIndex = 1 'the current recordset will be placed as the first element
        End If
        Set rsArray(iIndex) = rs        'buffer this recordset (which also gets returned)
    Else
        iIndex = conDimension
        ReDim rsArray(iIndex)           'this will cause "Nothing" to be returned
    End If
    
    Set finv_master = rsArray(iIndex)
    Exit Function
    
finv_masterRedim:
    If Err.Number = 9 Then 'Subscript out of range
        ' this is the expected error on the FIRST call to this function; we
        ' must dimension the static array and retry the first reference to it.
        Err.Clear
        ReDim rsArray(conDimension)
        On Error GoTo 0
        Resume
    Else
        Err.Raise Err.Number
    End If
    
End Function    ''finv_master

Public Function finv_uom( _
                                sParm1 As String, _
                                db As DataBase _
                                ) As Recordset
    Const SUB_NAME = "finv_uom"
    
    'returning pr_inv_uom As Recordset
        
    Dim sSql As String
    Dim iIndex As Integer
    Dim rs As Recordset
    Static rsArray() As Recordset

    On Error GoTo finv_uomRedim
    iIndex = 1
    Do While Not (rsArray(iIndex) Is Nothing)
        If rsArray(iIndex)!inu_code = sParm1 Then
            Set finv_uom = rsArray(iIndex)
            Exit Function   'return now with this "buffered" recordset
        End If
        iIndex = iIndex + 1
        If iIndex > conDimension Then Exit Do 'While  'Note: iIndex will be used below
    Loop 'While
    
    sSql = "SELECT * FROM inv_uom WHERE inu_code = " & QStr(sParm1)
    Set rs = fnMainmodsOpenRecordset(sSql, SUB_NAME, db)
    If rs Is Nothing Then
        Exit Function
    End If
    If rs.RecordCount > 0 Then
        rs.MoveFirst    'populate with first record
        #If cDEBUG Then
            On Error Resume Next
            While Not rs.EOF
                Debug.Print "finv_uom(" & QStr(sParm1) & ") = " & rs.Fields(0)
                rs.MoveNext    ' Locate next record
            Wend
            On Error GoTo 0
            rs.MoveFirst
        #End If
        If iIndex > conDimension Then
            ' shuffle array elements up (lose oldest -- highest -- element)
            For iIndex = conDimension - 1 To 1 Step -1
                Set rsArray(iIndex + 1) = rsArray(iIndex)
            Next
            iIndex = 1 'the current recordset will be placed as the first element
        End If
        Set rsArray(iIndex) = rs        'buffer this recordset (which also gets returned)
    Else
        iIndex = conDimension
        ReDim rsArray(iIndex)           'this will cause "Nothing" to be returned
    End If
    
    Set finv_uom = rsArray(iIndex)
    Exit Function
    
finv_uomRedim:
    If Err.Number = 9 Then 'Subscript out of range
        ' this is the expected error on the FIRST call to this function; we
        ' must dimension the static array and retry the first reference to it.
        Err.Clear
        ReDim rsArray(conDimension)
        On Error GoTo 0
        Resume
    Else
        Err.Raise Err.Number
    End If
    
End Function    ''finv_uom

Public Function finv_whlsale( _
                                sParm1 As Long, _
                                sParm2 As Integer, _
                                db As DataBase _
                                ) As Recordset
    Const SUB_NAME = "finv_whlsale"
    
    'returning pr_inv_whlsale As Recordset
        
    Dim sSql As String
    Dim iIndex As Integer
    Dim rs As Recordset
    Static rsArray() As Recordset

    On Error GoTo finv_whlsaleRedim
    iIndex = 1
    Do While Not (rsArray(iIndex) Is Nothing)
        If rsArray(iIndex)!inw_prodlnk = sParm1 And _
           rsArray(iIndex)!inw_prft_ctr = sParm2 Then
            Set finv_whlsale = rsArray(iIndex)
            Exit Function   'return now with this "buffered" recordset
        End If
        iIndex = iIndex + 1
        If iIndex > conDimension Then Exit Do 'While  'Note: iIndex will be used below
    Loop 'While
    
    sSql = "SELECT * FROM inv_whlsale WHERE inw_prodlnk = " & QStr(sParm1) & _
           " AND inw_prft_ctr = " & QStr(sParm2)
    Set rs = fnMainmodsOpenRecordset(sSql, SUB_NAME, db)
    If rs Is Nothing Then
        Exit Function
    End If
    If rs.RecordCount > 0 Then
        rs.MoveFirst    'populate with first record
        #If cDEBUG Then
            On Error Resume Next
            While Not rs.EOF
                Debug.Print "finv_whlsale(" & QStr(sParm1) & ") = " & rs.Fields(0)
                rs.MoveNext    ' Locate next record
            Wend
            On Error GoTo 0
            rs.MoveFirst
        #End If
        If iIndex > conDimension Then
            ' shuffle array elements up (lose oldest -- highest -- element)
            For iIndex = conDimension - 1 To 1 Step -1
                Set rsArray(iIndex + 1) = rsArray(iIndex)
            Next
            iIndex = 1 'the current recordset will be placed as the first element
        End If
        Set rsArray(iIndex) = rs        'buffer this recordset (which also gets returned)
    Else
        iIndex = conDimension
        ReDim rsArray(iIndex)           'this will cause "Nothing" to be returned
    End If
    
    Set finv_whlsale = rsArray(iIndex)
    Exit Function
    
finv_whlsaleRedim:
    If Err.Number = 9 Then 'Subscript out of range
        ' this is the expected error on the FIRST call to this function; we
        ' must dimension the static array and retry the first reference to it.
        Err.Clear
        ReDim rsArray(conDimension)
        On Error GoTo 0
        Resume
    Else
        Err.Raise Err.Number
    End If
    
End Function    ''finv_whlsale

Public Function frs_b_item( _
                                sParm1 As String, _
                                sParm2 As Long, _
                                db As DataBase _
                                ) As Recordset
    Const SUB_NAME = "frs_b_item"
    'returning pr_rs_b_item As Recordset
        
    Dim sSql As String
    Dim iIndex As Integer
    Dim rs As Recordset
    Static rsArray() As Recordset

    On Error GoTo frs_b_itemRedim
    iIndex = 1
    Do While Not (rsArray(iIndex) Is Nothing)
        If rsArray(iIndex)!rsbi_code = sParm1 And _
           rsArray(iIndex)!rsbi_vendor = sParm2 Then
            Set frs_b_item = rsArray(iIndex)
            Exit Function   'return now with this "buffered" recordset
        End If
        iIndex = iIndex + 1
        If iIndex > conDimension Then Exit Do 'While  'Note: iIndex will be used below
    Loop 'While
    
    sSql = "SELECT * FROM rs_b_item WHERE rsbi_code = " & QStr(sParm1) & _
           " AND rsbi_vendor = " & QStr(sParm2)
    Set rs = fnMainmodsOpenRecordset(sSql, SUB_NAME, db)
    If rs Is Nothing Then
        Exit Function
    End If
    If rs.RecordCount > 0 Then
        rs.MoveFirst    'populate with first record
        #If cDEBUG Then
            On Error Resume Next
            While Not rs.EOF
                Debug.Print "frs_b_item(" & QStr(sParm1) & ") = " & rs.Fields(0)
                rs.MoveNext    ' Locate next record
            Wend
            On Error GoTo 0
            rs.MoveFirst
        #End If
        If iIndex > conDimension Then
            ' shuffle array elements up (lose oldest -- highest -- element)
            For iIndex = conDimension - 1 To 1 Step -1
                Set rsArray(iIndex + 1) = rsArray(iIndex)
            Next
            iIndex = 1 'the current recordset will be placed as the first element
        End If
        Set rsArray(iIndex) = rs        'buffer this recordset (which also gets returned)
    Else
        iIndex = conDimension
        ReDim rsArray(iIndex)           'this will cause "Nothing" to be returned
    End If
    
    Set frs_b_item = rsArray(iIndex)
    Exit Function
    
frs_b_itemRedim:
    If Err.Number = 9 Then 'Subscript out of range
        ' this is the expected error on the FIRST call to this function; we
        ' must dimension the static array and retry the first reference to it.
        Err.Clear
        ReDim rsArray(conDimension)
        On Error GoTo 0
        Resume
    Else
        Err.Raise Err.Number
    End If
    
End Function    ''frs_b_item

Public Function frs_meter( _
                                sParm1 As Long, _
                                db As DataBase _
                                ) As Recordset
    Const SUB_NAME = "frs_meter"
    
    'returning pr_rs_meter As Recordset
        
    Dim sSql As String
    Dim iIndex As Integer
    Dim rs As Recordset
    Static rsArray() As Recordset

    On Error GoTo frs_meterRedim
    iIndex = 1
    Do While Not (rsArray(iIndex) Is Nothing)
        If rsArray(iIndex)!rsm_meter_nbr = sParm1 Then
            Set frs_meter = rsArray(iIndex)
            Exit Function   'return now with this "buffered" recordset
        End If
        iIndex = iIndex + 1
        If iIndex > conDimension Then Exit Do 'While  'Note: iIndex will be used below
    Loop 'While
    
    sSql = "SELECT * FROM rs_meter WHERE rsm_meter_nbr = " & QStr(sParm1)
    Set rs = fnMainmodsOpenRecordset(sSql, SUB_NAME, db)
    If rs Is Nothing Then
        Exit Function
    End If
    If rs.RecordCount > 0 Then
        rs.MoveFirst    'populate with first record
        #If cDEBUG Then
            On Error Resume Next
            While Not rs.EOF
                Debug.Print "frs_meter(" & QStr(sParm1) & ") = " & rs.Fields(0)
                rs.MoveNext    ' Locate next record
            Wend
            On Error GoTo 0
            rs.MoveFirst
        #End If
        If iIndex > conDimension Then
            ' shuffle array elements up (lose oldest -- highest -- element)
            For iIndex = conDimension - 1 To 1 Step -1
                Set rsArray(iIndex + 1) = rsArray(iIndex)
            Next
            iIndex = 1 'the current recordset will be placed as the first element
        End If
        Set rsArray(iIndex) = rs        'buffer this recordset (which also gets returned)
    Else
        iIndex = conDimension
        ReDim rsArray(iIndex)           'this will cause "Nothing" to be returned
    End If
    
    Set frs_meter = rsArray(iIndex)
    Exit Function
    
frs_meterRedim:
    If Err.Number = 9 Then 'Subscript out of range
        ' this is the expected error on the FIRST call to this function; we
        ' must dimension the static array and retry the first reference to it.
        Err.Clear
        ReDim rsArray(conDimension)
        On Error GoTo 0
        Resume
    Else
        Err.Raise Err.Number
    End If
    
End Function    ''frs_meter

Public Function frs_shiftlink( _
                                sParm1 As Long, _
                                db As DataBase _
                                ) As Recordset
    Const SUB_NAME = "frs_shiftlink"
    'returning pr_rs_shiftlink As Recordset
        
    Dim sSql As String
    Dim iIndex As Integer
    Dim rs As Recordset
    Static rsArray() As Recordset

    On Error GoTo frs_shiftlinkRedim
    iIndex = 1
    Do While Not (rsArray(iIndex) Is Nothing)
        If rsArray(iIndex)!rssl_shl = sParm1 Then
            Set frs_shiftlink = rsArray(iIndex)
            Exit Function   'return now with this "buffered" recordset
        End If
        iIndex = iIndex + 1
        If iIndex > conDimension Then Exit Do 'While  'Note: iIndex will be used below
    Loop 'While
    
    sSql = "SELECT * FROM rs_shiftlink WHERE rssl_shl = " & QStr(sParm1)
    Set rs = fnMainmodsOpenRecordset(sSql, SUB_NAME, db)
    If rs Is Nothing Then
        Exit Function
    End If
    If rs.RecordCount > 0 Then
        rs.MoveFirst    'populate with first record
        #If cDEBUG Then
            On Error Resume Next
            While Not rs.EOF
                Debug.Print "frs_shiftlink(" & QStr(sParm1) & ") = " & rs.Fields(0)
                rs.MoveNext    ' Locate next record
            Wend
            On Error GoTo 0
            rs.MoveFirst
        #End If
        If iIndex > conDimension Then
            ' shuffle array elements up (lose oldest -- highest -- element)
            For iIndex = conDimension - 1 To 1 Step -1
                Set rsArray(iIndex + 1) = rsArray(iIndex)
            Next
            iIndex = 1 'the current recordset will be placed as the first element
        End If
        Set rsArray(iIndex) = rs        'buffer this recordset (which also gets returned)
    Else
        iIndex = conDimension
        ReDim rsArray(iIndex)           'this will cause "Nothing" to be returned
    End If
    
    Set frs_shiftlink = rsArray(iIndex)
    Exit Function
    
frs_shiftlinkRedim:
    If Err.Number = 9 Then 'Subscript out of range
        ' this is the expected error on the FIRST call to this function; we
        ' must dimension the static array and retry the first reference to it.
        Err.Clear
        ReDim rsArray(conDimension)
        On Error GoTo 0
        Resume
    Else
        Err.Raise Err.Number
    End If
    
End Function    ''frs_shiftlink

Public Function fshift_link( _
                                sParm1 As Integer, _
                                sParm2 As Date, _
                                sParm3 As Integer, _
                                db As DataBase _
                                ) As Recordset
    Const SUB_NAME = "fshift_link"
    
    'returning pr_rs_shiftlink As Recordset
        
    Dim sSql As String
    Dim iIndex As Integer
    Dim rs As Recordset
    Static rsArray() As Recordset

    On Error GoTo fshift_linkRedim
    iIndex = 1
    Do While Not (rsArray(iIndex) Is Nothing)
        If rsArray(iIndex)!rssl_prft_ctr = sParm1 And _
           rsArray(iIndex)!rssl_date = sParm2 And _
           rsArray(iIndex)!rssl_shift = sParm3 Then
            Set fshift_link = rsArray(iIndex)
            Exit Function   'return now with this "buffered" recordset
        End If
        iIndex = iIndex + 1
        If iIndex > conDimension Then Exit Do 'While  'Note: iIndex will be used below
    Loop 'While
    
    sSql = "SELECT * FROM rs_shiftlink WHERE rssl_prft_ctr = " & QStr(sParm1) & _
           " AND  rssl_date = " & QStr(sParm2) & _
           " AND rssl_shift = " & QStr(sParm3)
    Set rs = fnMainmodsOpenRecordset(sSql, SUB_NAME, db)
    If rs Is Nothing Then
        Exit Function
    End If
    If rs.RecordCount > 0 Then
        rs.MoveFirst    'populate with first record
        #If cDEBUG Then
            On Error Resume Next
            While Not rs.EOF
                Debug.Print "fshift_link(" & QStr(sParm1) & ") = " & rs.Fields(0)
                rs.MoveNext    ' Locate next record
            Wend
            On Error GoTo 0
            rs.MoveFirst
        #End If
        If iIndex > conDimension Then
            ' shuffle array elements up (lose oldest -- highest -- element)
            For iIndex = conDimension - 1 To 1 Step -1
                Set rsArray(iIndex + 1) = rsArray(iIndex)
            Next
            iIndex = 1 'the current recordset will be placed as the first element
        End If
        Set rsArray(iIndex) = rs        'buffer this recordset (which also gets returned)
    Else
        iIndex = conDimension
        ReDim rsArray(iIndex)           'this will cause "Nothing" to be returned
    End If
    
    Set fshift_link = rsArray(iIndex)
    Exit Function
    
fshift_linkRedim:
    If Err.Number = 9 Then 'Subscript out of range
        ' this is the expected error on the FIRST call to this function; we
        ' must dimension the static array and retry the first reference to it.
        Err.Clear
        ReDim rsArray(conDimension)
        On Error GoTo 0
        Resume
    Else
        Err.Raise Err.Number
    End If
    
End Function    ''fshift_link

Public Function fsp_excpt_clas( _
                                sParm1 As Long, _
                                sParm2 As String, _
                                db As DataBase _
                                ) As Recordset
    Const SUB_NAME = "fsp_excpt_clas"
    'returning pr_sp_excpt_clas As Recordset
        
    Dim sSql As String
    Dim iIndex As Integer
    Dim rs As Recordset
    Static rsArray() As Recordset

    On Error GoTo fsp_excpt_clasRedim
    iIndex = 1
    Do While Not (rsArray(iIndex) Is Nothing)
        If rsArray(iIndex)!cce_cust = sParm1 And _
           rsArray(iIndex)!cce_product_cl = sParm2 Then
            Set fsp_excpt_clas = rsArray(iIndex)
            Exit Function   'return now with this "buffered" recordset
        End If
        iIndex = iIndex + 1
        If iIndex > conDimension Then Exit Do 'While  'Note: iIndex will be used below
    Loop 'While
    
    sSql = "SELECT * FROM sp_excpt_clas WHERE cce_cust = " & QStr(sParm1) & _
            " AND cce_product_cl = " & QStr(sParm2)
    Set rs = fnMainmodsOpenRecordset(sSql, SUB_NAME, db)
    If rs Is Nothing Then
        Exit Function
    End If
    If rs.RecordCount > 0 Then
        rs.MoveFirst    'populate with first record
        #If cDEBUG Then
            On Error Resume Next
            While Not rs.EOF
                Debug.Print "fsp_excpt_clas(" & QStr(sParm1) & ") = " & rs.Fields(0)
                rs.MoveNext    ' Locate next record
            Wend
            On Error GoTo 0
            rs.MoveFirst
        #End If
        If iIndex > conDimension Then
            ' shuffle array elements up (lose oldest -- highest -- element)
            For iIndex = conDimension - 1 To 1 Step -1
                Set rsArray(iIndex + 1) = rsArray(iIndex)
            Next
            iIndex = 1 'the current recordset will be placed as the first element
        End If
        Set rsArray(iIndex) = rs        'buffer this recordset (which also gets returned)
    Else
        iIndex = conDimension
        ReDim rsArray(iIndex)           'this will cause "Nothing" to be returned
    End If
    
    Set fsp_excpt_clas = rsArray(iIndex)
    Exit Function
    
fsp_excpt_clasRedim:
    If Err.Number = 9 Then 'Subscript out of range
        ' this is the expected error on the FIRST call to this function; we
        ' must dimension the static array and retry the first reference to it.
        Err.Clear
        ReDim rsArray(conDimension)
        On Error GoTo 0
        Resume
    Else
        Err.Raise Err.Number
    End If
    
End Function    ''fsp_excpt_clas
    

Public Function fsp_excpt_cust( _
                                sParm1 As Long, _
                                sParm2 As Long, _
                                db As DataBase _
                                ) As Recordset
    Const SUB_NAME = "fsp_excpt_cust"
    
    'returning pr_sp_excpt_cust As Recordset
        
    Dim sSql As String
    Dim iIndex As Integer
    Dim rs As Recordset
    Static rsArray() As Recordset

    On Error GoTo fsp_excpt_custRedim
    iIndex = 1
    Do While Not (rsArray(iIndex) Is Nothing)
        If rsArray(iIndex)!ce_cust = sParm1 And _
           rsArray(iIndex)!ce_prodlnk = sParm1 Then
            Set fsp_excpt_cust = rsArray(iIndex)
            Exit Function   'return now with this "buffered" recordset
        End If
        iIndex = iIndex + 1
        If iIndex > conDimension Then Exit Do 'While  'Note: iIndex will be used below
    Loop 'While
    
    sSql = "SELECT * FROM sp_excpt_cust WHERE ce_cust = " & QStr(sParm1) & _
            " AND ce_prodlnk = " & QStr(sParm2)
    Set rs = fnMainmodsOpenRecordset(sSql, SUB_NAME, db)
    If rs Is Nothing Then
        Exit Function
    End If
    If rs.RecordCount > 0 Then
        rs.MoveFirst    'populate with first record
        #If cDEBUG Then
            On Error Resume Next
            While Not rs.EOF
                Debug.Print "fsp_excpt_cust(" & QStr(sParm1) & ") = " & rs.Fields(0)
                rs.MoveNext    ' Locate next record
            Wend
            On Error GoTo 0
            rs.MoveFirst
        #End If
        If iIndex > conDimension Then
            ' shuffle array elements up (lose oldest -- highest -- element)
            For iIndex = conDimension - 1 To 1 Step -1
                Set rsArray(iIndex + 1) = rsArray(iIndex)
            Next
            iIndex = 1 'the current recordset will be placed as the first element
        End If
        Set rsArray(iIndex) = rs        'buffer this recordset (which also gets returned)
    Else
        iIndex = conDimension
        ReDim rsArray(iIndex)           'this will cause "Nothing" to be returned
    End If
    
    Set fsp_excpt_cust = rsArray(iIndex)
    Exit Function
    
fsp_excpt_custRedim:
    If Err.Number = 9 Then 'Subscript out of range
        ' this is the expected error on the FIRST call to this function; we
        ' must dimension the static array and retry the first reference to it.
        Err.Clear
        ReDim rsArray(conDimension)
        On Error GoTo 0
        Resume
    Else
        Err.Raise Err.Number
    End If
    
End Function    ''fsp_excpt_cust

Public Function fsp_group( _
                                sParm1 As String, _
                                db As DataBase _
                                ) As Recordset
    Const SUB_NAME = "fsp_group"
    'returning pr_sp_group As Recordset
        
    Dim sSql As String
    Dim iIndex As Integer
    Dim rs As Recordset
    Static rsArray() As Recordset

    On Error GoTo fsp_groupRedim
    iIndex = 1
    Do While Not (rsArray(iIndex) Is Nothing)
        If rsArray(iIndex)!spg_code = sParm1 Then
            Set fsp_group = rsArray(iIndex)
            Exit Function   'return now with this "buffered" recordset
        End If
        iIndex = iIndex + 1
        If iIndex > conDimension Then Exit Do 'While  'Note: iIndex will be used below
    Loop 'While
    
    sSql = "SELECT * FROM sp_group WHERE spg_code = " & QStr(sParm1)
    Set rs = fnMainmodsOpenRecordset(sSql, SUB_NAME, db)
    If rs Is Nothing Then
        Exit Function
    End If
    If rs.RecordCount > 0 Then
        rs.MoveFirst    'populate with first record
        #If cDEBUG Then
            On Error Resume Next
            While Not rs.EOF
                Debug.Print "fsp_group(" & QStr(sParm1) & ") = " & rs.Fields(0)
                rs.MoveNext    ' Locate next record
            Wend
            On Error GoTo 0
            rs.MoveFirst
        #End If
        If iIndex > conDimension Then
            ' shuffle array elements up (lose oldest -- highest -- element)
            For iIndex = conDimension - 1 To 1 Step -1
                Set rsArray(iIndex + 1) = rsArray(iIndex)
            Next
            iIndex = 1 'the current recordset will be placed as the first element
        End If
        Set rsArray(iIndex) = rs        'buffer this recordset (which also gets returned)
    Else
        iIndex = conDimension
        ReDim rsArray(iIndex)           'this will cause "Nothing" to be returned
    End If
    
    Set fsp_group = rsArray(iIndex)
    Exit Function
    
fsp_groupRedim:
    If Err.Number = 9 Then 'Subscript out of range
        ' this is the expected error on the FIRST call to this function; we
        ' must dimension the static array and retry the first reference to it.
        Err.Clear
        ReDim rsArray(conDimension)
        On Error GoTo 0
        Resume
    Else
        Err.Raise Err.Number
    End If
    
End Function    ''fsp_group

Public Function fsys_parm( _
                                sParm1 As Integer, _
                                db As DataBase _
                                ) As Recordset
    Const SUB_NAME = "fsys_parm"
    'returning pr_sys_parm As Recordset
        
    Dim sSql As String
    Dim iIndex As Integer
    Dim rs As Recordset
    Static pa_sys_parm() As Recordset

    On Error GoTo fsys_parmRedim
    iIndex = 1
    Do While Not (pa_sys_parm(iIndex) Is Nothing)
        If pa_sys_parm(iIndex)!parm_nbr = sParm1 Then
            Set fsys_parm = pa_sys_parm(iIndex)
            Exit Function   'return now with this "buffered" recordset
        End If
        iIndex = iIndex + 1
        If iIndex > conDimension Then Exit Do 'While  'Note: iIndex will be used below
    Loop 'While
    
    sSql = "SELECT * FROM sys_parm WHERE parm_nbr = " & QStr(sParm1)
    Set rs = fnMainmodsOpenRecordset(sSql, SUB_NAME, db)
    If rs Is Nothing Then
        Exit Function
    End If
    If rs.RecordCount > 0 Then
        rs.MoveFirst    'populate with first record
        #If cDEBUG Then
            On Error Resume Next
            While Not rs.EOF
                Debug.Print "fsys_parm(" & QStr(sParm1) & ") = " & rs!parm_desc
                Debug.Print "...  parm_field =" & rs!parm_field
                rs.MoveNext    ' Locate next record
            Wend
            On Error GoTo 0
            rs.MoveFirst
        #End If
        If iIndex > conDimension Then
            ' shuffle array elements up (lose oldest element)
            For iIndex = conDimension - 1 To 1 Step -1
                Set pa_sys_parm(iIndex + 1) = pa_sys_parm(iIndex)
            Next
            iIndex = 1 'the current recordset will be placed as the first element
        End If
        Set pa_sys_parm(iIndex) = rs        'buffer this recordset (which also gets returned)
    Else
        iIndex = conDimension
        ReDim pa_sys_parm(iIndex)           'this will cause "Nothing" to be returned
    End If
    
    Set fsys_parm = pa_sys_parm(iIndex)
    Exit Function
    
fsys_parmRedim:
    If Err.Number = 9 Then 'Subscript out of range
        ' this is the expected error on the FIRST call to this function; we
        ' must dimension the static array and retry the first reference to it.
        Err.Clear
        ReDim pa_sys_parm(conDimension)
        On Error GoTo 0
        Resume
    Else
        Err.Raise Err.Number
    End If
    
End Function    ''fsys_parm

Public Function fsys_prft_ctr( _
                                sParm1 As Long, _
                                db As DataBase _
                                ) As Recordset
    Const SUB_NAME = "fsys_prft_ctr"
    
    'returning pr_sys_prft_ctr As Recordset
        
    Dim sSql As String
    Dim iIndex As Integer
    Dim rs As Recordset
    Static rsArray() As Recordset
    
    On Error GoTo fsys_prft_ctrRedim
    iIndex = 1
    Do While Not (rsArray(iIndex) Is Nothing)
        If rsArray(iIndex)!prft_ctr = sParm1 Then
            Set fsys_prft_ctr = rsArray(iIndex)
            Exit Function   'return now with this "buffered" recordset
        End If
        iIndex = iIndex + 1
        If iIndex > conDimension Then Exit Do 'While  'Note: iIndex will be used below
    Loop 'While
    
    sSql = "SELECT * FROM sys_prft_ctr WHERE prft_ctr = " & QStr(sParm1)
    Set rs = fnMainmodsOpenRecordset(sSql, SUB_NAME, db)
    If rs Is Nothing Then
        Exit Function
    End If
    If rs.RecordCount > 0 Then
        rs.MoveFirst    'populate with first record
        #If cDEBUG Then
            On Error Resume Next
            While Not rs.EOF
                Debug.Print "fsys_prft_ctr(" & QStr(sParm1) & ") = " & rs.Fields(0)
                rs.MoveNext    ' Locate next record
            Wend
            On Error GoTo 0
            rs.MoveFirst
        #End If
        If iIndex > conDimension Then
            ' shuffle array elements up (lose oldest -- highest -- element)
            For iIndex = conDimension - 1 To 1 Step -1
                Set rsArray(iIndex + 1) = rsArray(iIndex)
            Next
            iIndex = 1 'the current recordset will be placed as the first element
        End If
        Set rsArray(iIndex) = rs        'buffer this recordset (which also gets returned)
    Else
        iIndex = conDimension
        ReDim rsArray(iIndex)           'this will cause "Nothing" to be returned
    End If
    
    Set fsys_prft_ctr = rsArray(iIndex)
    Exit Function
    
fsys_prft_ctrRedim:
    If Err.Number = 9 Then 'Subscript out of range
        ' this is the expected error on the FIRST call to this function; we
        ' must dimension the static array and retry the first reference to it.
        Err.Clear
        ReDim rsArray(conDimension)
'        On Error GoTo 0
        Resume
    ElseIf Err.Number <> 3146 Then
        Err.Raise Err.Number
    End If

End Function    ''fsys_prft_ctr

Public Function ftx_use_g( _
                                sParm1 As String, _
                                sParm2 As String, _
                                db As DataBase _
                                ) As Recordset
    Const SUB_NAME = "ftx_use_g"
    'returning pr_tx_use_g As Recordset
        
    Dim sSql As String
    Dim iIndex As Integer
    Dim rs As Recordset
    Static rsArray() As Recordset

    On Error GoTo ftx_use_gRedim
    iIndex = 1
    Do While Not (rsArray(iIndex) Is Nothing)
        If rsArray(iIndex)!tu_use_g = sParm1 And _
           rsArray(iIndex)!tu_prodtcl = sParm2 Then
            Set ftx_use_g = rsArray(iIndex)
            Exit Function   'return now with this "buffered" recordset
        End If
        iIndex = iIndex + 1
        If iIndex > conDimension Then Exit Do 'While  'Note: iIndex will be used below
    Loop 'While
    
    sSql = "SELECT * FROM tx_use_g WHERE tu_use_g = " & QStr(sParm1) & _
           " AND tu_prodtcl = " & QStr(sParm2)
    Set rs = fnMainmodsOpenRecordset(sSql, SUB_NAME, db)
    If rs Is Nothing Then
        Exit Function
    End If
    If rs.RecordCount > 0 Then
        rs.MoveFirst    'populate with first record
        #If cDEBUG Then
            On Error Resume Next
            While Not rs.EOF
                Debug.Print "ftx_use_g(" & QStr(sParm1) & ") = " & rs.Fields(0)
                rs.MoveNext    ' Locate next record
            Wend
            On Error GoTo 0
            rs.MoveFirst
        #End If
        If iIndex > conDimension Then
            ' shuffle array elements up (lose oldest -- highest -- element)
            For iIndex = conDimension - 1 To 1 Step -1
                Set rsArray(iIndex + 1) = rsArray(iIndex)
            Next
            iIndex = 1 'the current recordset will be placed as the first element
        End If
        Set rsArray(iIndex) = rs        'buffer this recordset (which also gets returned)
    Else
        iIndex = conDimension
        ReDim rsArray(iIndex)           'this will cause "Nothing" to be returned
    End If
    
    Set ftx_use_g = rsArray(iIndex)
    Exit Function
    
ftx_use_gRedim:
    If Err.Number = 9 Then 'Subscript out of range
        ' this is the expected error on the FIRST call to this function; we
        ' must dimension the static array and retry the first reference to it.
        Err.Clear
        ReDim rsArray(conDimension)
        On Error GoTo 0
        Resume
    Else
        Err.Raise Err.Number
    End If
    
End Function    ''ftx_use_g

Public Function ftx_use_gl( _
                                sParm1 As String, _
                                db As DataBase _
                                ) As Recordset
    Const SUB_NAME = "ftx_use_gl"
    'returning pr_tx_use_g As Recordset
        
    Dim sSql As String
    Dim iIndex As Integer
    Dim rs As Recordset
    Static rsArray() As Recordset

    On Error GoTo ftx_use_glRedim
    iIndex = 1
    Do While Not (rsArray(iIndex) Is Nothing)
        If rsArray(iIndex)!tu_use_g = sParm1 Then
            Set ftx_use_gl = rsArray(iIndex)
            Exit Function   'return now with this "buffered" recordset
        End If
        iIndex = iIndex + 1
        If iIndex > conDimension Then Exit Do 'While  'Note: iIndex will be used below
    Loop 'While
    
    sSql = "SELECT * FROM tx_use_g WHERE tu_use_g = " & QStr(sParm1)
    Set rs = fnMainmodsOpenRecordset(sSql, SUB_NAME, db)
    If rs Is Nothing Then
        Exit Function
    End If
    If rs.RecordCount > 0 Then
        rs.MoveFirst    'populate with first record
        #If cDEBUG Then
            On Error Resume Next
            While Not rs.EOF
                Debug.Print "ftx_use_gl(" & QStr(sParm1) & ") = " & rs.Fields(0)
                rs.MoveNext    ' Locate next record
            Wend
            On Error GoTo 0
            rs.MoveFirst
        #End If
        If iIndex > conDimension Then
            ' shuffle array elements up (lose oldest -- highest -- element)
            For iIndex = conDimension - 1 To 1 Step -1
                Set rsArray(iIndex + 1) = rsArray(iIndex)
            Next
            iIndex = 1 'the current recordset will be placed as the first element
        End If
        Set rsArray(iIndex) = rs        'buffer this recordset (which also gets returned)
    Else
        iIndex = conDimension
        ReDim rsArray(iIndex)           'this will cause "Nothing" to be returned
    End If
    
    Set ftx_use_gl = rsArray(iIndex)
    Exit Function
    
ftx_use_glRedim:
    If Err.Number = 9 Then 'Subscript out of range
        ' this is the expected error on the FIRST call to this function; we
        ' must dimension the static array and retry the first reference to it.
        Err.Clear
        ReDim rsArray(conDimension)
        On Error GoTo 0
        Resume
    Else
        Err.Raise Err.Number
    End If
    
End Function    ''ftx_use_gl

Public Function ftx_use_g_link( _
                                sParm1 As Long, _
                                db As DataBase _
                                ) As Recordset
    Const SUB_NAME = "ftx_use_g_link"
    'returning pr_tx_use_g As Recordset
        
    Dim sSql As String
    Dim iIndex As Integer
    Dim rs As Recordset
    Static rsArray() As Recordset

    On Error GoTo ftx_use_g_linkRedim
    iIndex = 1
    Do While Not (rsArray(iIndex) Is Nothing)
        If rsArray(iIndex)!tu_link = sParm1 Then
            Set ftx_use_g_link = rsArray(iIndex)
            Exit Function   'return now with this "buffered" recordset
        End If
        iIndex = iIndex + 1
        If iIndex > conDimension Then Exit Do 'While  'Note: iIndex will be used below
    Loop 'While
    
    sSql = "SELECT * FROM tx_use_g WHERE tu_link = " & QStr(sParm1)
    Set rs = fnMainmodsOpenRecordset(sSql, SUB_NAME, db)
    If rs Is Nothing Then
        Exit Function
    End If
    If rs.RecordCount > 0 Then
        rs.MoveFirst    'populate with first record
        #If cDEBUG Then
            On Error Resume Next
            While Not rs.EOF
                Debug.Print "ftx_use_g_link(" & QStr(sParm1) & ") = " & rs.Fields(0)
                rs.MoveNext    ' Locate next record
            Wend
            On Error GoTo 0
            rs.MoveFirst
        #End If
        If iIndex > conDimension Then
            ' shuffle array elements up (lose oldest -- highest -- element)
            For iIndex = conDimension - 1 To 1 Step -1
                Set rsArray(iIndex + 1) = rsArray(iIndex)
            Next
            iIndex = 1 'the current recordset will be placed as the first element
        End If
        Set rsArray(iIndex) = rs        'buffer this recordset (which also gets returned)
    Else
        iIndex = conDimension
        ReDim rsArray(iIndex)           'this will cause "Nothing" to be returned
    End If
    
    Set ftx_use_g_link = rsArray(iIndex)
    Exit Function
    
ftx_use_g_linkRedim:
    If Err.Number = 9 Then 'Subscript out of range
        ' this is the expected error on the FIRST call to this function; we
        ' must dimension the static array and retry the first reference to it.
        Err.Clear
        ReDim rsArray(conDimension)
        On Error GoTo 0
        Resume
    Else
        Err.Raise Err.Number
    End If
    
End Function    ''ftx_use_g_link

Public Function fvendor(code As Long, db As DataBase) As Recordset
    Dim pr_p_altname As Recordset
    Dim pr_p_vendor As Recordset
    
    Set pr_p_altname = fp_altname(code, db)
    If pr_p_altname Is Nothing Then
        Set pr_p_vendor = Nothing
    Else
        Set pr_p_vendor = fp_vendor(pr_p_altname!pn_vendor, db)
    End If
    Set fvendor = pr_p_vendor   ''it_is, pr_p_altname, pr_p_vendor

End Function    ''fvendor

Public Function it_is_a_fuel(code As String, db As DataBase) As Boolean
    Dim pr_inv_header As Recordset
    Dim pr_co_inv_codes As Recordset
    
    it_is_a_fuel = False
    
    Set pr_inv_header = finv_header(code, db)
    If Not (pr_inv_header Is Nothing) Then
        Set pr_co_inv_codes = fco_inv_codes(pr_inv_header!ivh_class, db)
        If pr_co_inv_codes Is Nothing Then
            Exit Function
        End If
        If IsNull(pr_co_inv_codes!coic_prog_class) Then
            Exit Function
        End If
        
        'trim the pr_co_inv_codes!coic_prog_class david 11/11/99
        If Trim(pr_co_inv_codes!coic_prog_class) = "fuel" Then
            it_is_a_fuel = True
            Exit Function
        End If
    End If
End Function    ''it_is_a_fuel


Public Sub program_time_stamp(ts_prog As String, ts_msg As String, ts_flag As String)
        
        Dim run_stmt As String
        run_stmt = "echo `date` " & RTrim(ts_msg)
        If Left(ts_flag, 1) = "N" Then
                run_stmt = RTrim(run_stmt) & " >"
        Else
                run_stmt = RTrim(run_stmt) & " >>"
        End If
        run_stmt = RTrim(run_stmt) & " " & RTrim(ts_prog) & ".time"
        'run run_stmt  ''???? what to do here
        #If DEVELOPMENT Then
            Beep
            MsgBox run_stmt, vbInformation, "program_time_stamp"
        #End If
End Sub     ''program_time_stamp

Public Function mastr_pc()
    mastr_pc = master_prft_ctr
End Function

Public Function rgfh_path_code(c As String, db As DataBase) As String
       
    Dim lr_gf_header As Recordset
    Set lr_gf_header = fgf_header(c, db)
    If Not (lr_gf_header Is Nothing) Then
        rgfh_path_code = lr_gf_header!gfh_desc
    Else
        rgfh_path_code = " "
    End If

End Function    ''rgfh_path_code

Public Function set_mastr_master(db As DataBase) As Integer
    Dim lc_parm As Recordset
    
    Set lc_parm = fsys_parm(2, db)
    If Not (lc_parm Is Nothing) Then master_prft_ctr = lc_parm!parm_field
    set_mastr_master = master_prft_ctr
End Function ''set_mastr_master


Public Function stdconv(from_uom_stock As String, _
                        to_uom_pricing As String, _
                        ByRef conv_fact As Double, _
                        db As DataBase _
                        ) As Byte
    Const SUB_NAME = "stdconv"
    
    'returning flag (stdconv), conv_fact as Double
    If from_uom_stock = to_uom_pricing Then
        conv_fact = 1#
        stdconv = 0 'flag
        Exit Function
    End If
    
    Dim sSql As String
    Dim rs As Recordset
    
    sSql = "SELECT trns_factor FROM sys_convert" & _
                " WHERE trns_unit_1 = " & QStr(from_uom_stock) & _
                " AND trns_unit_2 = " & QStr(to_uom_pricing)
    Set rs = fnMainmodsOpenRecordset(sSql, SUB_NAME, db)
    If rs Is Nothing Then
        Exit Function
    End If
    If rs.RecordCount > 0 Then
        rs.MoveFirst 'populate
        conv_fact = CDbl(rs!trns_factor)
        stdconv = 0 'flag
        Exit Function
    End If
    
    sSql = "SELECT inu_code FROM inv_uom" & _
                  " WHERE inu_code = " & QStr(from_uom_stock)
    Set rs = fnMainmodsOpenRecordset(sSql, SUB_NAME, db)
    If rs Is Nothing Then
        Exit Function
    End If
    
    If rs.RecordCount > 0 Then
        Set rs = Nothing
        sSql = "SELECT inu_code FROM inv_uom" & _
                    " WHERE inu_code = " & QStr(to_uom_pricing)
        Set rs = fnMainmodsOpenRecordset(sSql, SUB_NAME, db)
        If rs Is Nothing Then
            Exit Function
        End If
        If rs.RecordCount > 0 Then
            conv_fact = 0#
            stdconv = 2 'flag
            Exit Function
        End If
    End If
    conv_fact = 0#
    stdconv = 1 'flag
End Function 'stdconv

Public Function uround(val As Double, uom As String, db As DataBase) As Double
    
    Dim round As Integer    'number decimal places (defaults to zero)
    Dim rs As Recordset
    
    Set rs = finv_uom(uom, db)
    If Not (rs Is Nothing) Then round = CInt(rs!inu_rnd)
    
    uround = roundf(val, round)
End Function    ''uround

Private Function roundf(amt As Double, rnd As Integer) As Double
    Dim i As Integer
    Dim rstg As String

    'rndf = 0.5
    rstg = "############0."     '"-------------."
    For i = 1 To rnd
        rstg = RTrim(rstg) & "#"
    Next
    roundf = Format(amt, rstg)
End Function    ''roundf

'Vijaya on 02/05/04 Magic#395302
Public Function ftx_Header_detail( _
                            sParm1 As Long, _
                            sParm2 As Date, _
                            db As DataBase _
                            ) As Recordset
    Const SUB_NAME = "ftx_Header_detail"
    'returning pa_tx_detail As Recordset
        
    Dim strSQL As String
    Dim iIndex As Integer
    Dim rs As Recordset
    Static rsArray() As Recordset

    On Error GoTo ftx_Header_detailRedim
    iIndex = 1
    Do While Not (rsArray(iIndex) Is Nothing)
        If rsArray(iIndex)!txh_trn = sParm1 Then
            Set ftx_Header_detail = rsArray(iIndex)
            Exit Function   'return now with this "buffered" recordset
        End If
        iIndex = iIndex + 1
        If iIndex > conDimension Then Exit Do 'While  'Note: iIndex will be used below
    Loop 'While
    
    strSQL = "SELECT * from tx_header,tx_detail"
    strSQL = strSQL & " WHERE txd_trn = txh_trn"
    strSQL = strSQL & " AND txd_trn = " & QStr(sParm1)
    strSQL = strSQL & " AND " & tfnDateString(sParm2, True)
    strSQL = strSQL & " BETWEEN txd_beg_date AND txd_end_date"
    
    Set rs = fnMainmodsOpenRecordset(strSQL, SUB_NAME, db)
    If rs Is Nothing Then
        Exit Function
    End If
    If rs.RecordCount > 0 Then
        rs.MoveFirst    'populate with first record
        #If cDEBUG Then
            On Error Resume Next
            While Not rs.EOF
                Debug.Print "ftx_Header_detail(" & QStr(sParm1) & ") = " & rs.Fields(0)
                rs.MoveNext    ' Locate next record
            Wend
            On Error GoTo 0
            rs.MoveFirst
        #End If
        If iIndex > conDimension Then
            ' shuffle array elements up (lose oldest -- highest -- element)
            For iIndex = conDimension - 1 To 1 Step -1
                Set rsArray(iIndex + 1) = rsArray(iIndex)
            Next
            iIndex = 1 'the current recordset will be placed as the first element
        End If
        Set rsArray(iIndex) = rs        'buffer this recordset (which also gets returned)
    Else
        iIndex = conDimension
        ReDim rsArray(iIndex)           'this will cause "Nothing" to be returned
    End If
    
    Set ftx_Header_detail = rsArray(iIndex)
    Exit Function
    
ftx_Header_detailRedim:
    If Err.Number = 9 Then 'Subscript out of range
        ' this is the expected error on the FIRST call to this function; we
        ' must dimension the static array and retry the first reference to it.
        Err.Clear
        ReDim rsArray(conDimension)
        On Error GoTo 0
        Resume
    Else
        Err.Raise Err.Number
    End If
End Function    ''ftx_Header_detail
''''''''''''''''''''''''''''''''
