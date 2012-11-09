Attribute VB_Name = "gltxmods"
Option Explicit

Public pr_gl_master_glm_account As Long  'used to store invalid gl account
Public pr_gl_master_glm_prft_ctr As Long  ' and profit center for error message
Public sErr_Use_grp As String
Public sErr_tClass As String
Public sErr_ProductCode As String
Public sErr_ProductClass As String

Global stat_flag As Integer
Global buckt As Long
Global tb_buckt As Long
Global td_buckt As Long
Global del_cost As Double  'Like inv_master.invt_gl_cost
Global clearing_acct As Long
Global ap_cntl As Long
Global trk_rev As Long
Global trk_pc As Long
Global frgt_acct As Long
Global frgt_rev As Long
Global freight_account As String * 1
Global freight_revenue As String * 1
Global i As Long, j As Long, k As Long, L As Long, m As Long, n As Long
Global gpc_num As Long
Global gproduct As String * 16
Global gtax_use As String * 8
Global gtunits As Double
Global gcost As Double
Global gcust As Long
Global gsale As Double
Global gorigin As Long
Global gdest As Long
Global go_type As String * 1
Global charge_flag As String * 1
Global post_amt As Double
Global gl_num As Long
Global ar_cntl As Long
Global freight_amount As Double
Global gross_limit As Single 'Long  'Like tx_table.tt_limit
Private de_tax_amount As Single
Global re_tax_amount As Single

Global pr_cust_ctr As Recordset  '             record like sys_prft_ctr.*
Global pr_sys_parm As Recordset  '              record like sys_parm.*
Global pr_use_g As Recordset  '                 record like tx_use_g.*
Global pr_use_link As Recordset  '              record like tx_use_link.*
Global pr_table As Recordset  '                 record like tx_table.*
Global pr_header_detail As Recordset  '         record like tx_header.*,tx_detail.*
Global pr_inv_classes As Recordset  '           record like inv_classes.*
Global pr_gl_master As Recordset  '             record like gl_master.*
Global pr_gl_period As Recordset  '             record like gl_period.*
Global pr_fm_frgt_cd As Recordset  '            record like fm_frgt_cd.*
Global pr_altname As Recordset  '               record like ar_altname.*
Global pr_customer As Recordset  '              record like ar_customer.*
Global pr_ar_class As Recordset  '              record like ar_class.*

Global tul_max As Integer
Global ulink_cnt As Integer
Type tx_use_link_type
    tul_link As Long
    tul_trn As Long
    tul_sched As String * 5
End Type
Global pa_ulink() As tx_use_link_type  '         (800) of record like tx_use_link.*
Global tt_max As Integer
Global tt_cnt As Integer
Type tx_table_type
    tt_code As String * 5
    tt_level As String * 1
    tt_where As String * 4
    tt_order As Integer
    tt_gl_acct As Long
    tt_rate As Single
    tt_charge As String * 1
    tt_basis As String * 1
    tt_limit As Single 'Integer
    tt_date As Date
    tt_trn As Long
    tt_descr As String * 30
End Type

Global pa_table() As tx_table_type  '        (100) of record like tx_table.*

Type smm_tax_type
    smm_class As String * 5
    smm_d_amt As Double
    smm_bunits As Double
    smm_gunits As Double
    smm_nunits As Double
End Type
Global smm_tax() As smm_tax_type
Global smm_cnt As Integer
Global smm_ptr As Integer

Type grp_tax_type
    grp_class As String * 5
    grp_bunits As Double
    grp_tax As Double
End Type
Global grp_tax() As grp_tax_type
Global grp_cnt As Integer
Global grp_ptr As Integer
Global dsl_fuel_group As String * 5
Global by_line_or_inv As String * 1   '# L=limits by line, I=limits by invoice

Global trn_max As Integer
Global trn_arry_cnt As Integer
Type trn_type
    trn As Long  '               like fm_lad_tax.fmldt_trn,
    ulink As Long  '             like fm_lad_tax.fmldt_ulink,
    ttype As String * 1  '             like fm_lad_tax.fmldt_type,
    amount As Single  '            like fm_lad_tax.fmldt_amount,
    tunits As Double  '            like fm_lad_tax.fmldt_gal,
    charge As String * 1  '            like fm_lad_tax.fmldt_charge,
    tt_gl_acct  As Long  '         like tx_table.tt_gl_acct,
    extn As Double
    ''limit As Integer  '             like tx_table.tt_limit,
    limit As Single 'WJ Changed to signle on 1/11/99
    order_group As Integer  '        like tx_table.tt_order
End Type
Global trn_arry() As trn_type

Global fltr_arry_cnt As Integer
Global fltr_arry() As trn_type

Global sls_tax_cnt As Integer
Type sls_tax_arry_type
    trn As Long
    amount As Double
End Type
Global sls_tax_arry(50) As sls_tax_arry_type

Global dtx_max As Integer
Global dtx_cnt As Integer
Type dtx_type
    Description As String * 30
    amount As Double
End Type
Global dtx() As dtx_type

'# this is the loading fee
Type ldf_ary_type
    dv_nbr As Long
    dv_line As Integer
    dv_prodlnk As Long
    dv_prodtcl As String * 5
    dv_vol As Double
    dv_atax As Double
End Type
Global ldf_ary(50) As ldf_ary_type
Global ldf_cnt As Integer
Global ldf_ptr As Integer

Global sum_fuels As Double
Global effective_sum_fuels As Double
Global gas_gals As Double
Global dsl_gals As Double
Global a_tax_total(25) As Double             '# keeps a total of all "A"
                                                   '# taxes paid

Global base_amount As Double
Global last_order_group As Integer
Global last_limit As Single 'As Integer  'Like tx_table.tt_limit
Global group_amount As Double
Global tax_amount As Single
Global true_sign As Double

Private bUseNonStkMstr As Boolean

Private Const nDBInformix As Integer = 1
'Vijaya on 02/05/04 Magic#395302
Public bExists_tx_header As Boolean
''''''''''''''''''''''''''''''''
'

Public bUseGLAdjust As Boolean
Public lAdjGLAccount As Long
Public sTaxUseGroup As String

Private origin_state As String      'Vijaya on 11/13/09 #647512
'

Public Sub initialize_tax_buffers()
    error_found = False
    
    'Vijaya on 12/09/04 Magic#465142
    'Don't Delete trn_max because some times it won't call setup_for_gl
    'It should be in Two Places
    'Change from 100 to 500 by JQ, ticket 490248
    trn_max = 500
    tul_max = 800
    tt_max = 500

    freight_amount = 0
    ReDim pa_table(tt_max)
    tt_cnt = 0

    ReDim pa_ulink(tul_max)
    ulink_cnt = 0
    
    init_detax
    de_tax_amount = 0#
    
    Set pr_sys_parm = fsys_parm(1055, t_dbMainDatabase)
    If pr_sys_parm Is Nothing Then
        dsl_fuel_group = "NONE"    '#value that can not occur
    Else
        dsl_fuel_group = fnCStr(pr_sys_parm!parm_field)
    End If
    
    start_ldfee

    origin_state = "^^"     'Vijaya on 11/13/09 #647512
End Sub


Public Sub setup_for_gl(prod As String, this_pc As Long, tax_acct As Long)
    trn_max = 500
    
    ReDim trn_arry(trn_max)
    trn_arry_cnt = 0
    fltr_arry_cnt = 0

    clear_gl_lines

    Set pr_sys_parm = fsys_parm(1051, t_dbMainDatabase)
    If IsNull(pr_sys_parm!parm_field) Then
        found_error = True: error_found = 1051
        Exit Sub
    End If
    
    buckt = tfnRound(pr_sys_parm!parm_field)

    Set pr_gl_master = fgl_master(sSeries, buckt, this_pc, t_dbMainDatabase)
    If pr_gl_master Is Nothing Then
        pr_gl_master_glm_account = buckt
        pr_gl_master_glm_prft_ctr = this_pc
        found_error = True: error_found = 10
        Exit Sub
    End If

    Set pr_sys_parm = fsys_parm(1052, t_dbMainDatabase)
    If IsNull(pr_sys_parm!parm_field) Then
        found_error = True: error_found = 1052
        Exit Sub
    End If
    clearing_acct = tfnRound(pr_sys_parm!parm_field)

    Set pr_gl_master = fgl_master(sSeries, clearing_acct, this_pc, t_dbMainDatabase)
    If pr_gl_master Is Nothing Then
        pr_gl_master_glm_account = clearing_acct
        pr_gl_master_glm_prft_ctr = this_pc
        found_error = True: error_found = 10
        Exit Sub
    End If

    set_this_pc this_pc

    If tax_acct > 0 Then '# taxes in cost layers (sysparm 1006 = Y)
        tb_buckt = tax_acct
        td_buckt = tax_acct
    Else                 '# taxes in cost layers (sysparm 1006 = N)
        Set pr_sys_parm = fsys_parm(1020, t_dbMainDatabase)
        If IsNull(pr_sys_parm!parm_field) Then
            tb_buckt = 0
        Else
            tb_buckt = tfnRound(pr_sys_parm!parm_field)
        End If
        td_buckt = tb_buckt
        If tb_buckt <> 0 Then
            Set pr_gl_master = fgl_master(sSeries, tb_buckt, this_pc, t_dbMainDatabase)
            If pr_gl_master Is Nothing Then
                '# if sysparm 1020 does not contain a valid acct'# treat
                '#  it as if the parm were null, i.e. turn of the tax in cost
                '#  feature
                tb_buckt = 0
                td_buckt = 0
            End If
        End If
    End If

    Set pr_sys_parm = fsys_parm(5001, t_dbMainDatabase)
    If IsNull(pr_sys_parm!parm_field) Then
        found_error = True: error_found = 5001
        Exit Sub
    End If
    ap_cntl = tfnRound(pr_sys_parm!parm_field)
    Set pr_gl_master = fgl_master(sSeries, ap_cntl, this_pc, t_dbMainDatabase)
    If pr_gl_master Is Nothing Then
        pr_gl_master_glm_account = ap_cntl
        pr_gl_master_glm_prft_ctr = this_pc
        found_error = True: error_found = 10
        Exit Sub
    End If

    If prod = "" Then
        If tb_buckt = 0 Then
'            return true, 1020
            found_error = True: error_found = 1020
            Exit Sub
        End If
        found_error = True: error_found = 1
        Exit Sub
    End If

    bUseNonStkMstr = False

    Set pr_inv_header = finv_header(prod, t_dbMainDatabase)
    Set pr_inv_master = finv_master(pr_inv_header!ivh_link, this_pc, t_dbMainDatabase)
    
    If pr_inv_master Is Nothing Then
        Dim rsTemp As Recordset
        'changed by junsong
        If fnCheckNonStkMstr(rsTemp, pr_inv_header!ivh_link, this_pc) Then
            bUseNonStkMstr = True
            Set pr_inv_master = rsTemp
        Else
            found_error = True: error_found = 1
            Exit Sub
        End If
    End If
      '  found_error = True: error_found = 1
      '  Exit Sub
    'Else
    '    If pr_inv_master.RecordCount = 0 Then
    '        Dim rsTemp As Recordset
            'changed by junsong
    '        If fnCheckNonStkMstr(rsTemp, pr_inv_header!ivh_link, this_pc) Then
    '            bUseNonStkMstr = True
    '            Set pr_inv_master = rsTemp
    '        Else
    '            found_error = True: error_found = 1
    '            Exit Sub
    '        End If
    '    End If
    ' End If
    
    If Not bUseNonStkMstr Then
        Set pr_gl_master = fgl_master(sSeries, pr_inv_master!invt_gl_sales, this_pc, t_dbMainDatabase)
        If pr_gl_master Is Nothing Then
            pr_gl_master_glm_account = pr_inv_master!invt_gl_sales
            pr_gl_master_glm_prft_ctr = this_pc
            found_error = True: error_found = 10
            Exit Sub
        End If
    End If
    
    Set pr_gl_master = fgl_master(sSeries, pr_inv_master!invt_gl_cost, this_pc, t_dbMainDatabase)
    If pr_gl_master Is Nothing Then
        pr_gl_master_glm_account = pr_inv_master!invt_gl_cost
        pr_gl_master_glm_prft_ctr = this_pc
        found_error = True: error_found = 10
        Exit Sub
    End If

    If Not bUseNonStkMstr Then
        Set pr_gl_master = fgl_master(sSeries, pr_inv_master!invt_gl_invt, this_pc, t_dbMainDatabase)
        If pr_gl_master Is Nothing Then
            pr_gl_master_glm_account = pr_inv_master!invt_gl_invt
            pr_gl_master_glm_prft_ctr = this_pc
            found_error = True: error_found = 10
            Exit Sub
        End If
    End If
    
    If tb_buckt = 0 Then
        tb_buckt = pr_inv_master!invt_gl_cost
    End If
    If td_buckt = 0 Then
        td_buckt = pr_inv_master!invt_gl_sales '# was invt_gl_sales
    End If                                         '# before this change

    Set pr_inv_classes = finv_classes(pr_inv_header!ivh_prodtcl, "T", t_dbMainDatabase)
    If pr_inv_classes Is Nothing Then
        found_error = True: error_found = 12
        Exit Sub
    End If

    'Added by JQ 02/16/2006 for ticket 393360
    If bUseGLAdjust Then
         If pr_inv_classes!icl_fuel_tax = "N" Then
            Set pr_use_g = ftx_use_g(sTaxUseGroup, pr_inv_header!ivh_prodtcl, t_dbMainDatabase)
            
            If Not pr_use_g Is Nothing And lAdjGLAccount > 0 Then
                Set pr_gl_master = fgl_master(sSeries, lAdjGLAccount, this_pc, t_dbMainDatabase)
                If pr_gl_master Is Nothing Then
                    pr_gl_master_glm_account = lAdjGLAccount
                    pr_gl_master_glm_prft_ctr = this_pc
                    found_error = True: error_found = 10
                    Exit Sub
                End If
                
                td_buckt = lAdjGLAccount
            End If
         End If
    End If
    
    found_error = False: error_found = 0
End Sub

'This function is called in WSERCVER program only.
Public Sub set_gl_adjust(bValue As Boolean, lGLAccount As Long)
    bUseGLAdjust = bValue
    lAdjGLAccount = lGLAccount
End Sub

Public Sub set_this_pc(pc As Long)
    gpc_num = pc
End Sub

Public Function gthis_pc(c As Long)
    If c > -1 Then
        gpc_num = c
    Else
        gthis_pc = gpc_num
    End If
End Function


Public Function get_a_trn(which_one As Integer) As trn_type
    get_a_trn = trn_arry(which_one)
End Function

Public Function trn_count()
    trn_count = trn_arry_cnt
End Function

Public Function get_a_fltr(which_one As Integer) As trn_type
    get_a_fltr = fltr_arry(which_one)
End Function

Public Function fltr_count()
    fltr_count = fltr_arry_cnt
End Function

Public Sub filter_trn()
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer

    fltr_arry_cnt = 0
    ReDim fltr_arry(trn_max)

    If trn_arry_cnt < 1 Then
        Exit Sub
    End If

    i = 1
    While i <= trn_arry_cnt
        If (trn_arry(i).charge = "Y" Or trn_arry(i).charge = "N") And (trn_arry(i).limit = 0 Or trn_arry(i).extn <> 0) Then
            i = i + 1
            GoTo continue_while
        End If
        If (trn_arry(i).limit = 0) Or (trn_arry(i).extn <> 0) Then
            fltr_arry_cnt = fltr_arry_cnt + 1
            fltr_arry(fltr_arry_cnt) = trn_arry(i)
        End If
        For j = i To trn_arry_cnt
            k = j + 1
            If k <= trn_arry_cnt Then
                trn_arry(j) = trn_arry(k)
            End If
        Next
        
        trn_arry_cnt = trn_arry_cnt - 1

continue_while:
    Wend
End Sub
            

Public Sub extrn_bl_gl(this_pc As Long, _
                     Product As String, _
                     tax_use As String, _
                     ByVal tunits As Double, _
                     ByVal nunits As Double, _
                     ByVal gunits As Double, _
                     ByVal tdolrs As Double, _
                     ByVal cost As Double, _
                     tax_acct As Long)
    
    '# if the 'include taxes in cost layer' option is being used
    '# sysparm 1006 = Y, then tax_acct will be non-zero
    setup_for_gl Product, this_pc, tax_acct
    sTaxUseGroup = tax_use
    If found_error Then
        Exit Sub
    End If

    If tunits < 0 Or cost < 0 Then
        true_sign = -1
        tunits = tunits * -1
        nunits = nunits * -1
        gunits = gunits * -1
        tdolrs = tdolrs * -1
        cost = cost * -1
    Else
        true_sign = 1
    End If


    If pr_inv_classes!icl_fuel_tax = "Y" And tunits <> 0 Then      '{  No taxes anyway }

        Set pr_use_g = ftx_use_g(tax_use, pr_inv_header!ivh_prodtcl, t_dbMainDatabase)
    
        If pr_use_g Is Nothing Then
            sErr_Use_grp = tax_use
            sErr_tClass = pr_inv_header!ivh_prodtcl
            found_error = True
            error_found = 29
            Exit Sub
        End If

        trn_arry_cnt = 0
        load_trn_arry found_error, error_found
        If found_error Then
            Exit Sub
        End If

        init_tax_amt tdolrs
        fnd_smmry Product
        fnd_grp_tax Product

        For i = 1 To trn_arry_cnt
            Select Case trn_arry(i).ttype
            Case "N"
                trn_arry(i).tunits = nunits
            Case "G"
                trn_arry(i).tunits = gunits
            Case "A"
                trn_arry(i).tunits = nunits
            Case Else
                trn_arry(i).tunits = tunits
            End Select
            Select Case trn_arry(i).ttype
            Case "$"
                gross_limit = smm_tax(smm_ptr).smm_d_amt
            Case Else
                gross_limit = smm_tax(smm_ptr).smm_gunits
            End Select
            trn_arry(i).extn = compute_tax_amt(trn_arry(i).tunits, trn_arry(i).amount, _
                            trn_arry(i).order_group, trn_arry(i).limit, trn_arry(i).ttype, _
                            trn_arry(i).trn) * true_sign
            trn_arry(i).tunits = trn_arry(i).tunits * true_sign
        Next

    End If  '{end of tax check }

'{
'        post the product entries
'}
    post_amt = dround(cost) * true_sign
    post_debit clearing_acct, post_amt
    post_credit buckt, post_amt

    If found_error Then
        Exit Sub
    End If

'{
'        use the rules for taxes
'}
    If pr_inv_classes!icl_fuel_tax = "Y" And tunits <> 0 Then      '{  No taxes anyway }
        post_also "T"
        For L = 1 To trn_arry_cnt
            post_tax "B", trn_arry(L).tt_gl_acct, trn_arry(L).charge, trn_arry(L).extn, buckt
        Next

    End If   '{end of tax check }

    filter_trn
    found_error = False: error_found = 0

End Sub

Public Sub extrn_ap_gl(this_pc As Long, Product As String, tax_use As String, tunits As Double, _
                            tdolrs As Double, cost As Double)
    
    setup_for_gl Product, this_pc, 0
    
    If found_error Then
        If Product <> "" Then
            Exit Sub
        End If
    End If

    post_amt = dround(cost)
    post_debit buckt, post_amt
    post_credit ap_cntl, post_amt
    
    If found_error Then
        Exit Sub
    End If

    filter_trn
    found_error = False: error_found = 0

End Sub

Public Sub intrn_ap_gl(this_pc As Long, Product As String, tax_use As String, tunits As Double, _
                            tdolrs As Double, cost As Double)
    
    setup_for_gl Product, this_pc, 0
    
    If found_error Then
        If Product <> "" Then
            Exit Sub
        End If
    End If

    post_amt = dround(cost)
    post_debit pr_inv_master!invt_gl_cost, post_amt
    post_credit buckt, post_amt

    If found_error Then
        Exit Sub
    End If

    filter_trn
    found_error = False: error_found = 0

End Sub

Public Sub intrn_bl_gl(this_pc As Long, _
                       Product As String, _
                       tax_use As String, _
                       ByVal tunits As Double, _
                       ByVal nunits As Double, _
                       ByVal gunits As Double, _
                       ByVal tdolrs As Double, _
                       ByVal cost As Double, _
                       tax_acct As Long)
    
    On Error GoTo errTrap
    
    setup_for_gl Product, this_pc, tax_acct
    
    If found_error Then
        Exit Sub
    End If

    If tunits < 0 Or cost < 0 Then
        true_sign = -1
        tunits = tunits * -1
        nunits = nunits * -1
        gunits = gunits * -1
        tdolrs = tdolrs * -1
        cost = cost * -1
    Else
        true_sign = 1
    End If

    If pr_inv_classes!icl_fuel_tax = "Y" And tunits <> 0 Then       '{  No taxes anyway }

        Set pr_use_g = ftx_use_g(tax_use, pr_inv_header!ivh_prodtcl, t_dbMainDatabase)

        If pr_use_g Is Nothing Then
            sErr_Use_grp = tax_use
            sErr_tClass = pr_inv_header!ivh_prodtcl
            found_error = True
            error_found = 29
            Exit Sub
        End If

        trn_arry_cnt = 0
        load_trn_arry found_error, error_found
        If found_error Then
            Exit Sub
        End If

        init_tax_amt tdolrs
        fnd_smmry Product
        fnd_grp_tax Product

        For i = 1 To trn_arry_cnt
            Select Case trn_arry(i).ttype
            Case "N"
                trn_arry(i).tunits = nunits
            Case "G"
                trn_arry(i).tunits = gunits
            Case "A"
                trn_arry(i).tunits = nunits
            Case Else
                trn_arry(i).tunits = tunits
            End Select
            Select Case trn_arry(i).ttype
            Case "$"
                gross_limit = smm_tax(smm_ptr).smm_d_amt
            Case Else
                gross_limit = smm_tax(smm_ptr).smm_gunits
            End Select
            trn_arry(i).extn = compute_tax_amt(trn_arry(i).tunits, trn_arry(i).amount, _
                                trn_arry(i).order_group, trn_arry(i).limit, _
                                trn_arry(i).ttype, trn_arry(i).trn) * true_sign
            trn_arry(i).tunits = trn_arry(i).tunits * true_sign

        Next

    End If    '          {end of tax check }


'{
'        post the product entries
'}
    post_amt = dround(cost) * true_sign
    post_debit clearing_acct, post_amt
    post_credit pr_inv_master!invt_gl_cost, post_amt

    If found_error Then
        Exit Sub
    End If

'{
'                use the rules for taxes
'}
    If pr_inv_classes!icl_fuel_tax = "Y" And tunits <> 0 Then   '{  No taxes anyway }
        post_also ("T")
        For L = 1 To trn_arry_cnt
            If trn_arry(L).charge = "Y" Then
            'Vijaya on 02/05/04 Magic#395302
            'Uncommented Becaue we need to be same as 4GL code
'               return true,11
                found_error = True: error_found = 11
                Exit Sub
            End If
            post_tax "B", trn_arry(L).tt_gl_acct, trn_arry(L).charge, trn_arry(L).extn, buckt
        Next

    End If              '{end of tax check }

    filter_trn
    found_error = False: error_found = 0
    
    Exit Sub

errTrap:
    #If PROCESSING Then
        tfnErrHandler "intrn_bl_gl", False
    #Else
        tfnErrHandler "intrn_bl_gl"
    #End If
End Sub

Public Sub intrn_del_gl(this_pc As Long, _
                        Product As String, _
                        tax_use As String, _
                        ByVal tunits As Double, _
                        ByVal nunits As Double, _
                        ByVal gunits As Double, _
                        ByVal tdolrs As Double, _
                        ByVal cost As Double)
    
    setup_for_gl Product, this_pc, 0
    If found_error Then
        Exit Sub
    End If


    If tunits < 0 Or cost < 0 Then
        true_sign = -1
        tunits = tunits * -1
        nunits = nunits * -1
        gunits = gunits * -1
        tdolrs = tdolrs * -1
        cost = cost * -1
    Else
        true_sign = 1
    End If

    If pr_inv_classes!icl_fuel_tax = "Y" And tunits <> 0 Then     '{  No taxes anyway }

        Set pr_use_g = ftx_use_g(tax_use, pr_inv_header!ivh_prodtcl, t_dbMainDatabase)

        If pr_use_g Is Nothing Then
            sErr_Use_grp = tax_use
            sErr_tClass = pr_inv_header!ivh_prodtcl
            found_error = True
            error_found = 29
            Exit Sub
        End If

        trn_arry_cnt = 0
        load_trn_arry found_error, error_found
        If found_error Then
            Exit Sub
        End If

        init_tax_amt tdolrs
        fnd_smmry Product
        fnd_grp_tax Product

        For i = 1 To trn_arry_cnt
            Select Case trn_arry(i).ttype
            Case "N"
                trn_arry(i).tunits = nunits
            Case "G"
                trn_arry(i).tunits = gunits
            Case "A"
                trn_arry(i).tunits = nunits
            Case Else
                trn_arry(i).tunits = tunits
            End Select
            Select Case trn_arry(i).ttype
            Case "$"
                gross_limit = smm_tax(smm_ptr).smm_d_amt
            Case Else
                gross_limit = smm_tax(smm_ptr).smm_gunits
            End Select
            trn_arry(i).extn = compute_tax_amt(trn_arry(i).tunits, trn_arry(i).amount, _
                            trn_arry(i).order_group, trn_arry(i).limit, trn_arry(i).ttype, _
                            trn_arry(i).trn) * true_sign
            trn_arry(i).tunits = trn_arry(i).tunits * true_sign
        Next
    End If    '{end of tax check }


'{
'        post the product entries
'}
    post_amt = dround(cost) * true_sign
    post_debit pr_inv_master!invt_gl_cost, post_amt
    post_credit clearing_acct, post_amt

    If found_error Then
        Exit Sub
    End If

'{
'        use the rules for taxes
'}
    If pr_inv_classes!icl_fuel_tax = "Y" And tunits <> 0 Then  '{  No taxes anyway }
        For L = 1 To trn_arry_cnt
            If trn_arry(L).charge = "Y" Then
                found_error = True: error_found = 11
                Exit Sub
            End If
            post_tax "D", trn_arry(L).tt_gl_acct, trn_arry(L).charge, trn_arry(L).extn, ar_cntl
        Next

    End If   '{end of tax check }

    filter_trn
    found_error = False: error_found = 0

End Sub

Public Sub extrn_del_gl(this_pc As Long, Product As String, _
                        tax_use As String, ByVal tunits As Double, _
                        ByVal nunits As Double, ByVal gunits As Double, _
                        ByVal tdolrs As Double, ByVal cost As Double, _
                        cust As Long, ByVal Sale As Double)
    
    On Error GoTo errTrap
    
    setup_for_gl Product, this_pc, 0
    sTaxUseGroup = tax_use
    If found_error Then
        Exit Sub
    End If

    If tunits < 0 Or cost < 0 Then
        true_sign = -1
        tunits = tunits * -1
        nunits = nunits * -1
        gunits = gunits * -1
        tdolrs = tdolrs * -1
        cost = cost * -1
        Sale = Sale * -1
    Else
        true_sign = 1
    End If

    Set pr_cust_ctr = fsys_prft_ctr(cust, t_dbMainDatabase)
    If pr_cust_ctr Is Nothing Then
        Set pr_altname = far_altname(cust, t_dbMainDatabase)
        If pr_altname Is Nothing Then
            found_error = True: error_found = 5
            Exit Sub
        End If
        Set pr_customer = far_customer(pr_altname!an_cust, t_dbMainDatabase)
        If pr_customer Is Nothing Then
            found_error = True: error_found = 5
            Exit Sub
        End If
        Set pr_ar_class = far_class(pr_customer!cust_cust_class, t_dbMainDatabase)
    Else
        Set pr_ar_class = far_class("PRFTC", t_dbMainDatabase)
    End If
    If pr_ar_class Is Nothing Then
            found_error = True: error_found = 6
            Exit Sub
    End If
    ar_cntl = pr_ar_class!arc_glm


    If tunits <> 0 Then      '{  No taxes anyway }

        Set pr_use_g = ftx_use_g(tax_use, pr_inv_header!ivh_prodtcl, t_dbMainDatabase)

        If pr_use_g Is Nothing Then
            sErr_Use_grp = tax_use
            sErr_tClass = pr_inv_header!ivh_prodtcl
            found_error = True
            error_found = 29
            Exit Sub
        End If

        'Set Origin State if necessary
        If origin_state = "^^" Then
            set_dflt_origin this_pc
        End If
        
        trn_arry_cnt = 0
        load_trn_arry found_error, error_found
        If found_error Then
            Exit Sub
        End If

        init_tax_amt tdolrs
        fnd_smmry Product
        fnd_grp_tax Product

        For i = 1 To trn_arry_cnt
            Select Case trn_arry(i).ttype
            Case "N"
                trn_arry(i).tunits = nunits
            Case "G"
                trn_arry(i).tunits = gunits
            Case "A"
                trn_arry(i).tunits = nunits
            Case Else
                trn_arry(i).tunits = tunits
            End Select
            Select Case trn_arry(i).ttype
            Case Is = "$"
                gross_limit = smm_tax(smm_ptr).smm_d_amt
            Case Else
                gross_limit = smm_tax(smm_ptr).smm_gunits
            End Select
            trn_arry(i).extn = compute_tax_amt(trn_arry(i).tunits, trn_arry(i).amount, _
                                trn_arry(i).order_group, trn_arry(i).limit, trn_arry(i).ttype, _
                                trn_arry(i).trn) * true_sign
            trn_arry(i).tunits = trn_arry(i).tunits * true_sign
            
            '-----------------------------------------------------------------
            '#Add the following codes by WJ on 05/28/03
            '#Magic #400088.
            '# Calculate re_tax_amount - used to balance detax routine
            If trn_arry(i).charge = "Y" Then
                 re_tax_amount = re_tax_amount + trn_arry(i).extn
            End If
            '------------------------------------------------------------------
        Next
        '-----------------------------------------------------------------
        '#Add the following codes by WJ on 05/28/03
        '#Magic #400088.
        '# If de_tax_amount has a value other than null, it means that
        '# this function has been called after calling the detax function.
        '# There are situations where the sale will be off small amounts
        '# due to rounding errors.  This 'if' statement will correct
        '# the rounding errors.

        If tfnRound(de_tax_amount, 2) <> 0 Then
            Sale = Sale + tfnRound((de_tax_amount - re_tax_amount) * true_sign, 2)
        End If
        '----------------------------------------------------------------
'{
'        deal with detax rounding error
'}

' in informix, de_tax_amount is always NULL  ???????????????????????
' so bypass the following statements

'        If de_tax_amount <> re_tax_amount Then
'            sale = sale + ((de_tax_amount - re_tax_amount) * true_sign)
'        End If

    End If  '{end of tax check }


'{
'        post the product entries
'}

    post_amt = dround(cost) * true_sign
    post_debit pr_inv_master!invt_gl_cost, post_amt
    post_credit clearing_acct, post_amt

    If found_error Then
        Exit Sub
    End If

    post_amt = dround(Sale) * true_sign
    post_debit ar_cntl, post_amt
    post_credit pr_inv_master!invt_gl_sales, post_amt

    If found_error Then
        Exit Sub
    End If


'{
'        use the rules for taxes
'}

    If tunits <> 0 Then      '{  No taxes anyway }
        For L = 1 To trn_arry_cnt
            post_tax "D", trn_arry(L).tt_gl_acct, trn_arry(L).charge, trn_arry(L).extn, ar_cntl
        Next

    End If   '{end of tax check }

    filter_trn
    found_error = False: error_found = 0
    
    Exit Sub

errTrap:
    #If PROCESSING Then
        tfnErrHandler "extrn_del_gl", False
    #Else
        tfnErrHandler "extrn_del_gl"
    #End If
End Sub

Public Sub post_tax(bod As String, tax_gl As Long, charge_flag As String, post_amt As Double, ext_gl As Long)
    Dim t_buck As Long

    If post_amt = 0# Then
        Exit Sub
    End If

    post_also "T"

    If pr_inv_classes!icl_fuel_tax = "N" Then
        bod = "S"
    End If

    Select Case bod
        Case "B"
            post_amt = dround(post_amt)
            Select Case charge_flag
              Case "Y"
                post_credit ext_gl, post_amt
                post_debit tb_buckt, post_amt
              Case "N"
                post_credit tax_gl, post_amt
                post_debit tb_buckt, post_amt
              Case "D"
                post_debit tax_gl, post_amt
                post_credit tb_buckt, post_amt
              Case "C"
                post_debit tb_buckt, post_amt
                post_credit tax_gl, post_amt
            End Select
        Case "S"
            post_amt = dround(post_amt)
            Select Case charge_flag
              Case "Y"
                post_debit ext_gl, post_amt
                post_credit tax_gl, post_amt
            End Select
        Case "D"
            post_amt = dround(post_amt)
            Select Case charge_flag
              Case "Y"
                post_debit ext_gl, post_amt
                post_credit td_buckt, post_amt
              Case "N"
                post_debit tax_gl, post_amt
                post_credit tb_buckt, post_amt
              Case "D"
                post_debit tax_gl, post_amt
                post_credit tb_buckt, post_amt
              Case "C"
                post_debit tb_buckt, post_amt
                post_credit tax_gl, post_amt
            End Select
        Case Else
            t_buck = buckt
    End Select
End Sub

Public Sub load_trn_arry(bParam1 As Integer, nParam2 As Integer)
    Dim si As Integer, sj As Integer, Changed As Integer
    Dim skeyi As String * 22, skeyj As String * 22
    Dim found_sw As Integer, sSql As String, pr_use_link As Recordset
    Dim buseTax As Boolean
    

    On Error GoTo errTrap
    
    'Vijaya on 02/05/04 Magic#395302
    If fnCStr(t_tax_date) = "" Then
        set_tax_date Date
    End If
    If bExists_tx_header Then
        'declare c_taxes cursor
        sSql = "SELECT * FROM tx_use_link, tx_header, tx_detail"
        sSql = sSql & " WHERE tul_link =" & pr_use_g!tu_link
        sSql = sSql & " AND tul_trn  = txh_trn"
        sSql = sSql & " AND tul_trn  = txd_trn"
        sSql = sSql & " AND " & tfnDateString(t_tax_date, True) & " between txd_beg_date "
        sSql = sSql & " AND txd_end_date"
        fnGetRecord pr_use_link, sSql, nDBInformix, "load_trn_arry"
        pr_use_link.MoveLast
        pr_use_link.MoveFirst
        While Not pr_use_link.EOF
            'Check Origin State if tax is restricted on the Origin
            buseTax = True
            If fnCheck_txh_origin_st Then
                If fnCStr(pr_use_link.txh_origin_st) <> "" Then
                    If fnCStr(pr_use_link.txh_origin_st) <> origin_state Then
                        buseTax = False
                    End If
                End If
            End If
            If buseTax Then
                trn_arry_cnt = trn_arry_cnt + 1
                ReDim Preserve trn_arry(trn_arry_cnt)
                j = trn_arry_cnt
                If j <= trn_max Then
                    trn_arry(j).trn = pr_use_link.tul_trn
                    trn_arry(j).ulink = pr_use_link.tul_link
                    trn_arry(j).ttype = pr_use_link.txd_basis
                    trn_arry(j).amount = pr_use_link.txd_rate
                    trn_arry(j).charge = pr_use_link.txd_charge
                    trn_arry(j).tt_gl_acct = pr_use_link.txd_gl_acct
                    trn_arry(j).limit = pr_use_link.txd_limit
                    'we don't use limit_type
                    trn_arry(j).order_group = pr_use_link.txd_order
                    Set pr_gl_master = fgl_master(sSeries, trn_arry(j).tt_gl_acct, gpc_num, t_dbMainDatabase)
                    
                    If pr_gl_master Is Nothing Then
                        pr_gl_master_glm_account = trn_arry(j).tt_gl_acct
                        pr_gl_master_glm_prft_ctr = gpc_num
                        bParam1 = True
                        nParam2 = 10
                        Exit Sub
                    End If
                End If
            End If
            pr_use_link.MoveNext
        Wend
    Else
    ''''''''''''''''''''''''''''''''
    'old style
        found_sw = 0
        For i = 1 To ulink_cnt
            If pa_ulink(i).tul_link = pr_use_g!tu_link Then
                trn_arry_cnt = trn_arry_cnt + 1
                'david 06/18/2001
                ReDim Preserve trn_arry(trn_arry_cnt)
                j = trn_arry_cnt
                trn_arry(j).trn = pa_ulink(i).tul_trn
                trn_arry(j).ulink = pa_ulink(i).tul_link
                found_sw = 1
            End If
        Next
    
    '    declare c_tx_link cursor for
            sSql = "select * from tx_use_link Where tul_link = " & pr_use_g!tu_link
            fnGetRecord pr_use_link, sSql, nDBInformix, "load_trn_arry"
    
        If found_sw = 0 Then
            pr_use_link.MoveLast
            pr_use_link.MoveFirst
            While Not pr_use_link.EOF
                ulink_cnt = ulink_cnt + 1
                ReDim Preserve pa_ulink(ulink_cnt)
                With pa_ulink(ulink_cnt)
                    .tul_link = pr_use_link!tul_link
                    .tul_trn = pr_use_link!tul_trn
                    .tul_sched = pr_use_link!tul_sched
                End With
                trn_arry_cnt = trn_arry_cnt + 1
                ReDim Preserve trn_arry(trn_arry_cnt)
                i = ulink_cnt
                j = trn_arry_cnt
                trn_arry(j).trn = pa_ulink(i).tul_trn
                trn_arry(j).ulink = pa_ulink(i).tul_link
                pr_use_link.MoveNext
           Wend
        End If
    
        If trn_arry_cnt = 0 Then
            bParam1 = True
            nParam2 = 2
            Exit Sub
        End If
    
        Dim k As Long
        For i = 1 To trn_arry_cnt
            For j = 1 To tt_max
                If pa_table(j).tt_trn = 0 Then
                    k = trn_arry(i).trn
                    Set pr_table = ftx_table(k, t_dbMainDatabase)
                    If pr_table Is Nothing Then
                        bParam1 = True
                        nParam2 = 3
                        Exit Sub
                    End If
                    'david 02/20/2003  #391336
                    'the basis 'A' tax should be excluded!!!
                    If fnCStr(pr_table!tt_basis) <> "A" Then
                        If tt_cnt < tt_max Then
                            tt_cnt = tt_cnt + 1
                        End If
                        For k = tt_cnt To 2 Step -1
                            j = k - 1
                            pa_table(k) = pa_table(j)
                        Next
                        With pa_table(1)
                            .tt_code = pr_table!tt_code
                            .tt_level = pr_table!tt_level
                            .tt_where = pr_table!tt_where
                            .tt_order = pr_table!tt_order
                            .tt_gl_acct = pr_table!tt_gl_acct
                            .tt_rate = pr_table!tt_rate
                            .tt_charge = pr_table!tt_charge
                            .tt_basis = pr_table!tt_basis
                            .tt_limit = pr_table!tt_limit
                            .tt_date = pr_table!tt_date
                            .tt_trn = pr_table!tt_trn
                            .tt_descr = pr_table!tt_descr
                        End With
                    End If
                    ''''''''''''''''''''''''''
                End If
                If trn_arry(i).trn = pa_table(j).tt_trn Then
                    trn_arry(i).ttype = pa_table(j).tt_basis
                    trn_arry(i).amount = pa_table(j).tt_rate
                    trn_arry(i).charge = pa_table(j).tt_charge
                    trn_arry(i).limit = pa_table(j).tt_limit
                    trn_arry(i).order_group = pa_table(j).tt_order
                    trn_arry(i).tt_gl_acct = pa_table(j).tt_gl_acct
                    Set pr_gl_master = fgl_master(sSeries, trn_arry(i).tt_gl_acct, gpc_num, t_dbMainDatabase)
                    
                    If pr_gl_master Is Nothing Then
                        pr_gl_master_glm_account = trn_arry(i).tt_gl_acct
                        pr_gl_master_glm_prft_ctr = gpc_num
                        bParam1 = True
                        nParam2 = 10
                        Exit Sub
                    End If
                    Exit For
                End If
            Next
        Next
    End If
'{
'         Simple bubble sort to force order then abs(limit)
'}
'!!!!!!
'!!!!!!
'NOTE: 01/15/99 DAVID CHAI --- USE THE ELEMENT 0 AS A TEMPORARY STORAGE

    Changed = trn_arry_cnt
    While Changed > 1
        Changed = 0
        For si = 2 To trn_arry_cnt
            sj = si - 1
            skeyi = Format(trn_arry(si).order_group, "######") & Format(trn_arry(si).limit, "###########")
            skeyj = Format(trn_arry(sj).order_group, "######") & Format(trn_arry(sj).limit, "###########")
            If skeyj > skeyi Then
                trn_arry(0) = trn_arry(si)
                trn_arry(si) = trn_arry(sj)
                trn_arry(sj) = trn_arry(0)
                Changed = si
            End If
        Next
    Wend
    bParam1 = False
    nParam2 = 0

Exit Sub

errTrap:
    #If PROCESSING Then
        tfnErrHandler "load_trn_arry", False
    #Else
        tfnErrHandler "load_trn_arry"
    #End If
'
'For si = 1 To trn_arry_cnt
'With trn_arry(si)
'Debug.Print .ttype; " ";
'Debug.Print Format(.amount, "####0.000"); " ";
'Debug.Print .charge; " ";
'Debug.Print .limit; " ";
'Debug.Print .order_group; " ";
'Debug.Print .tt_gl_acct
'End With
'Next
End Sub

Public Sub set_tax_freight(l_amount As Double)
    freight_amount = tfnRound(l_amount, 2)
End Sub

Public Sub init_tax_amt(tdolrs As Double)

    base_amount = Abs(tdolrs)
    last_order_group = 0
    last_limit = 0
    tax_amount = 0
    group_amount = 0
    re_tax_amount = 0

End Sub

Public Sub fueltax_errors(e_num As Integer)
    Dim fe_msg As String

    fe_msg = tax_batch(e_num)

    #If PROCESSING Then
        #If DEVELOP Then
            MsgBox fe_msg, vbCritical
        #End If
        subPostError "fueltax_errors", fe_msg, nCurrIndex
    #Else
        MsgBox fe_msg, vbExclamation
    #End If

End Sub

Public Function tax_batch(e_num As Integer) As String
 Dim fe_msg As String
    Select Case e_num
        Case 1
             fe_msg = "Inventory record not found."
        Case 2
             fe_msg = "No Tax Records in use group."
        Case 3
             fe_msg = "No tax record for trn in link."
        Case 4
             fe_msg = "Not enough room(maybe) in passing array."
        Case 5
             fe_msg = "Customer Not Found."
        Case 6
             fe_msg = "Customer Class Not found."
        Case 7
             fe_msg = "System Parms 1016 & 1017 do not match."
        Case 8
             fe_msg = "Freight code not found."
        Case 9
             fe_msg = "Freight Rate not found."
        Case 10
             fe_msg = "G/L Account" & Str$(pr_gl_master_glm_account) & " not found for Profit Center" & Str$(pr_gl_master_glm_prft_ctr) & "."
        Case 11
             fe_msg = "'Y' flag is not allowed on Internal Transactions."
        Case 12
             fe_msg = "Tax Class Not Found."
        Case 13
             fe_msg = "Interstate Origin Terminal Not Found."
        Case 14
             fe_msg = "Interstate Origin Terminal City Not Found."
        Case 15
             fe_msg = "Interstate Origin Profit Center Not Found."
        Case 16
             fe_msg = "Interstate Origin Profit Center City Not Found."
        Case 17
             fe_msg = "Interstate Dest. Customer Not Found."
        Case 18
             fe_msg = "Interstate Dest. Customer City Not Found."
        Case 19
             fe_msg = "Interstate Dest. Profit Center Not Found."
        Case 20
             fe_msg = "Interstate Dest. Profit Center City Not Found."
        Case 21
             fe_msg = "Interstate Origin State Use Group Not Found."
        Case 22
             fe_msg = "Interstate Dest. State Use Group Not Found."
        Case 23
             fe_msg = "Interstate Origin State Not Found."
        Case 24
             fe_msg = "Interstate Origin County Use Group Not Found."
        Case 25
             fe_msg = "Interstate Dest. State Not Found."
        Case 26
             fe_msg = "Interstate Dest. County Use Group Not Found."
        Case 27
             fe_msg = "Interstate Origin City Use Group Not Found."
        Case 28
             fe_msg = "Interstate Dest. City Use Group Not Found."
        Case 29
             fe_msg = "Tax Use Group: " & sErr_Use_grp _
                & IIf(Trim(sErr_tClass) <> "", " and/or Product Class: " _
                & sErr_tClass, "") & " Not Found in tx_use_g Table."
        Case 1051
             fe_msg = "System Parm 1051(Expected Liabilities) is missing."
        Case 1052
             fe_msg = "System Parm 1052(Clearing Account) is missing."
        Case 1016
             fe_msg = "System Parm 1016(Freight Account) is missing."
        Case 1017
             fe_msg = "System Parm 1017(Freight Revenue) is missing."
        Case 1018
             fe_msg = "System Parm 1018(Trucking Revenue) is missing."
        Case 1019
             fe_msg = "System Parm 1019(Trucking Profit Center) is missing."
        Case 1020
             fe_msg = "System Parm 1020(Taxes In Inventory) is missing."
        Case 5001
             fe_msg = "System Parm 5001(Accounts Payable Control) is missing."
        Case 51  'added by david 04/19/2002
             fe_msg = "Product Class " + tfnSQLString(sErr_ProductClass) + " Not found in inv_classes"
             
             If sErr_ProductCode <> "" Then
                fe_msg = fe_msg + " for product " + tfnSQLString(sErr_ProductCode)
             End If
             
             fe_msg = fe_msg + "."
        Case Else
             fe_msg = "Error Detected in G/L & Tax Routines."
    End Select

    tax_batch = fe_msg

End Function

Public Function compute_tax_amt(tunits As Double, rate As Single, order_group As Integer, limit As Single, basis As String, trn As Long)
    Dim sign_it As Integer
    Dim wrk_limit As Single
    Dim wrk_units As Double
    Dim wrk_amount As Double
    Dim basis_amount As Double
    Dim freight_tmp As Double
    Dim hold_basis As String * 1
    Dim delta As Double
    Dim sWHERE As String
    
    'Vijaya on 02/05/04 Magic#395302
    Dim stx_charge As String
     
    If fnCStr(t_tax_date) = "" Then
        set_tax_date Date
    End If
    ''''''''''''''''''''''''''''''''
        
    If tunits < 0 Then
        tunits = Abs(tunits)
        sign_it = -1
    Else
        sign_it = 1
    End If

'#   made changes 12/21/89 to only put in freight on the record that has
'#   an "F" on it.  Back it out after this record is calculated   rwr
'#
    hold_basis = basis
    Select Case Trim$(basis)
      Case "$"
        basis_amount = base_amount
      Case "F"
        freight_tmp = freight_amount
        If (freight_tmp < 0) Then
            freight_tmp = freight_tmp * -1
        End If
        basis_amount = freight_tmp + base_amount
        base_amount = freight_tmp + base_amount
        basis = "$"
      Case "A"
        basis_amount = 1
        base_amount = 1
      Case Else
        basis_amount = tunits
    End Select
    'Vijaya on 02/05/04 Magic#395302
    'Moved the code down same as 4GL
    'group_amount = group_amount + tax_amount
    If order_group <> last_order_group Then
        If last_order_group <> 0 Then
            base_amount = base_amount + group_amount
        End If
        group_amount = 0
        last_limit = 0
    End If
    last_order_group = order_group

    gross_limit = Abs(gross_limit)
    
    'Vijaya on 02/05/04 Magic#395302
    If bExists_tx_header Then
        Set pr_header_detail = ftx_Header_detail(trn, CDate(t_tax_date), t_dbMainDatabase)
        If Not pr_header_detail Is Nothing Then
            If pr_header_detail.RecordCount > 0 Then
               If Not IsNull(pr_header_detail!txh_where) Then
                    sWHERE = Trim(pr_header_detail!txh_where)
               End If
               stx_charge = pr_header_detail!txd_charge
            End If
        End If
    Else
        Set pr_table = ftx_table(trn, t_dbMainDatabase)
        If Not pr_table Is Nothing Then
            If pr_table.RecordCount > 0 Then
               If Not IsNull(pr_table!tt_where) Then
                    sWHERE = Trim(pr_table!tt_where)
               End If
               stx_charge = pr_table!tt_charge
            End If
        End If
    End If

    Select Case True
    Case limit < 0
        wrk_limit = 0 - limit
        If (basis = "A") Then   '# texas loading fee will be "A"
            If (ldf_ptr = 0) Then        '# we are doing the bol side
                If (Abs(effective_sum_fuels) >= last_limit And Abs(effective_sum_fuels) < wrk_limit) Then
                    wrk_units = tunits
                    wrk_amount = base_amount
                Else
                    wrk_units = 0
                    wrk_amount = 0
                End If
            Else
                If Abs(tunits) >= last_limit And Abs(tunits) < wrk_limit Then
                    wrk_units = tunits
                    wrk_amount = base_amount
                Else
                    wrk_units = 0
                wrk_amount = 0
                End If
            End If
        Else   '# the other flaky texas tax will not be "A"
            If (gross_limit >= last_limit And gross_limit < wrk_limit) Then
                wrk_units = tunits
                wrk_amount = base_amount
            Else
                wrk_units = 0
                wrk_amount = 0
            End If
        End If
        last_limit = wrk_limit
    Case limit > 0
        If basis = "$" Then
          wrk_units = tunits
          wrk_amount = base_amount
        Else
          wrk_limit = limit
          wrk_units = 0
          wrk_amount = 0
          If basis_amount >= last_limit And basis_amount <= wrk_limit Then
              wrk_units = tunits - last_limit
              wrk_amount = base_amount - last_limit
          Else
            If basis_amount > wrk_limit Then
              wrk_units = wrk_limit - last_limit
              wrk_amount = wrk_limit - last_limit
            End If
          End If
          last_limit = wrk_limit
        End If
    Case Else
        wrk_units = tunits
        wrk_amount = base_amount
    End Select

    Select Case Trim$(basis)
    Case "$"
        
        'Vijaya on 02/05/04 Moved the code up
        '#Florida has a very strange way to calculate taxes.
        '#So if Level="S"--state and Where="FL"--Florida,
        '#we have a new routine to calculate the taxes.
        '#Magic #381858. WJ 09/03/2002
        'Set pr_table = ftx_table(trn_arry(I).trn, t_dbMainDatabase)
        '#This was wrong. Should use trn. WJ, 4/01/2003,Magic 404131
        
        '#If it is Florida State Taxes
        If sWHERE = "FL" Then
            'tax_amount = florida_bracket_tax(wrk_amount, rate)
            tax_amount = round_up_sls_tx(wrk_amount * rate)
        Else
            '#For all other taxes, it is simple - Dol * Rate and
            '#Round it.
            tax_amount = tfnRound(wrk_amount * rate, 2)
        End If
        
        If limit > 0 Then
            by_line_or_inv = fnGetTT_LIMIT_TYPE(trn) 'get it when we need it
            If by_line_or_inv = "I" Then
                tax_amount = tfnRound(sales_tax_limit(tax_amount, limit, trn), 2)
            Else
                If tax_amount > limit Then
                    tax_amount = tfnRound(limit, 2)
                End If
            End If
        End If
    Case "A"
        If (wrk_units <> 0) Then
            If (ldf_ptr <> 0) Then
                tax_amount = tfnRound(Abs(ldf_ary(ldf_ptr).dv_atax), 2)
            Else
                If (sum_fuels = 0) Then
                    tax_amount = 0#
                Else
                    tax_amount = tfnRound((wrk_units / Abs(sum_fuels)) * rate, 2)
                End If
                'tax_amount = dround(tax_amount)
                delta = rate - (tax_amount + a_tax_total(order_group))
                delta = Abs(delta)
                If (delta < 0.5) Then
                    tax_amount = rate - a_tax_total(order_group)
                    tax_amount = Abs(tax_amount)
                End If
                a_tax_total(order_group) = a_tax_total(order_group) + tax_amount
            End If
        Else
            tax_amount = 0#
        End If
    Case Else
        tax_amount = tfnRound(wrk_units * rate, 2)
    End Select

    'tax_amount = dround(tax_amount) * sign_it
    tax_amount = tax_amount * sign_it
    If (hold_basis = "F") Then
        base_amount = base_amount - freight_tmp
    End If

    'Vijaya on 02/05/04 Magic#395302
    'move the code down Because it be the same as 4GL
    If stx_charge = "Y" Then
        group_amount = group_amount + tax_amount
    End If
    ''''''''''''''''''''''''''''''''
    
    compute_tax_amt = tax_amount

End Function


'#This routine is for Florida State Tax only
'#WJ 08/29/2002         Magic#: 376809
Private Function florida_bracket_tax(ByVal wk_amt As Double, ByVal rate As Double) As Double
     Dim dol_part      As Double
     Dim cent_part     As Double
     Dim tax_amt       As Double
     Dim sign_it       As Integer
     Dim str_amt       As String
     
     '#make the wk_amt non-negative and compute the tax
     '#apply the sign to tax amt after it is calculated
     If wk_amt < 0 Then
        sign_it = -1
        wk_amt = wk_amt * sign_it
     Else
        sign_it = 1
     End If
     
     wk_amt = tfnRound(wk_amt, 2)
     str_amt = tfnFormatDecimal(wk_amt, 2)
     dol_part = Left(str_amt, Len(str_amt) - 3)
     
     tax_amt = 0
     Select Case rate
        Case 0.06
           If wk_amt > 1.09 Then
              tax_amt = dol_part * rate
              cent_part = wk_amt - dol_part
           Else
              cent_part = wk_amt
           End If
           
           If cent_part >= 0.1 And cent_part < 0.17 Then
                tax_amt = tax_amt + 0.01
           ElseIf cent_part >= 0.17 And cent_part < 0.34 Then
                tax_amt = tax_amt + 0.02
           ElseIf cent_part >= 0.34 And cent_part < 0.51 Then
                tax_amt = tax_amt + 0.03
           ElseIf cent_part >= 0.51 And cent_part < 0.66 Then
                tax_amt = tax_amt + 0.04
           ElseIf cent_part >= 0.67 And cent_part < 0.84 Then
                tax_amt = tax_amt + 0.05
           ElseIf cent_part >= 0.84 Then
                tax_amt = tax_amt + 0.06
           End If

        Case 0.07
           If wk_amt > 1.09 Then
              tax_amt = dol_part * rate
              cent_part = wk_amt - dol_part
           Else
              cent_part = wk_amt
           End If
         
           If cent_part >= 0.1 And cent_part < 0.15 Then
                tax_amt = tax_amt + 0.01
           ElseIf cent_part >= 0.15 And cent_part < 0.29 Then
                tax_amt = tax_amt + 0.02
           ElseIf cent_part >= 0.29 And cent_part < 0.43 Then
                tax_amt = tax_amt + 0.03
           ElseIf cent_part >= 0.43 And cent_part < 0.58 Then
                tax_amt = tax_amt + 0.04
           ElseIf cent_part >= 0.58 And cent_part < 0.72 Then
                tax_amt = tax_amt + 0.05
           ElseIf cent_part >= 0.72 And cent_part < 0.86 Then
                tax_amt = tax_amt + 0.06
           ElseIf cent_part >= 0.86 Then
                tax_amt = tax_amt + 0.07
           End If

        Case 0.025
           If wk_amt > 2.09 Then
              '#calculate dol_part. it is quotient of wk_amt mod 2
              '#example wk_amt = 9.99, dol_part = 2
              str_amt = tfnFormatDecimal(dol_part / 2, 2)
              dol_part = Left(str_amt, Len(str_amt) - 3)

              tax_amt = dol_part * rate * 2
              cent_part = wk_amt - dol_part * 2
           Else
              Let cent_part = wk_amt
           End If

           If cent_part >= 0.1 And cent_part < 0.41 Then
               tax_amt = tax_amt + 0.01
           ElseIf cent_part >= 0.41 And cent_part < 0.81 Then
               tax_amt = tax_amt + 0.02
           ElseIf cent_part >= 0.81 And cent_part < 1.21 Then
               tax_amt = tax_amt + 0.03
           ElseIf cent_part >= 1.21 And cent_part < 1.61 Then
               tax_amt = tax_amt + 0.04
           ElseIf cent_part >= 1.61 Then
               tax_amt = tax_amt + 0.05
           End If

        Case Else
           tax_amt = wk_amt * rate
     End Select

     '##apply the sign back and return
     florida_bracket_tax = tfnRound(tax_amt * sign_it, 2)
End Function

Public Sub init_detax()
    dtx_max = 100
    dtx_cnt = 0
    
    ReDim dtx(dtx_max)

    de_tax_amount = 0
End Sub

Public Sub detax(use_group As String, tclass As String, _
                 tdolrs As Double, tunits As Double)

    Dim level_rate As Single
    Dim last_level As Integer
    
    'Vijaya on 02/05/04 Magic#395302
    Dim stx_charge As String
    ''''''''''''''''''''''''''''''''
    
    On Error GoTo errTrap

    'Vijaya on 02/05/04 Magic#395302
    If fnCStr(t_tax_date) = "" Then
        set_tax_date Date
    End If
    ''''''''''''''''''''''''''''''''
    
    Set pr_use_g = ftx_use_g(use_group, tclass, t_dbMainDatabase)

    If pr_use_g Is Nothing Then
        sErr_Use_grp = use_group
        sErr_tClass = tclass
        found_error = True
        error_found = 29
        Exit Sub
    End If
    
    trn_arry_cnt = 0
    load_trn_arry found_error, error_found
    If found_error Then
        Exit Sub
    End If

    base_amount = Abs(tdolrs)
    last_order_group = -32767
    last_limit = 0
    tax_amount = 0
    group_amount = 0
    dtx_cnt = trn_arry_cnt
    last_level = 0

    If trn_arry_cnt > 0 Then
        For i = 1 To trn_arry_cnt
            If trn_arry(i).order_group = 0 Then
                'Vijaya on 02/05/04 Magic#395302
                If bExists_tx_header Then
                    Set pr_header_detail = ftx_Header_detail(trn_arry(i).trn, CDate(t_tax_date), t_dbMainDatabase)
                    If Not pr_header_detail Is Nothing Then
                        If pr_header_detail.RecordCount > 0 Then
                            dtx(i).Description = pr_header_detail!txh_descr
                        End If
                    End If
                Else
                    Set pr_table = ftx_table(trn_arry(i).trn, t_dbMainDatabase)
                    'ReDim Preserve dtx(i)
                    dtx(i).Description = pr_table!tt_descr
                End If
                ''''''''''''''''''''''''''''''''
                dtx(i).amount = de_tax_amt(i, _
                                tunits, _
                                trn_arry(i).amount, _
                                -32767, _
                                trn_arry(i).limit, _
                                trn_arry(i).ttype, _
                                trn_arry(i).charge)

                de_tax_amount = de_tax_amount + dtx(i).amount
            End If
        Next
    End If

    level_rate = 0
    If trn_arry_cnt > 0 Then
        j = 0 'WJ 12/04/2001 Magic 339720
        For i = trn_arry_cnt To 1 Step -1
            If trn_arry(i).order_group <> 0 Then
                'Vijaya on 02/05/04 Magic#395302
                If bExists_tx_header Then
                    Set pr_header_detail = ftx_Header_detail(trn_arry(i).trn, CDate(t_tax_date), t_dbMainDatabase)
                    dtx(i).Description = pr_header_detail!txh_descr
                    stx_charge = pr_header_detail!txd_charge
                Else
                    Set pr_table = ftx_table(trn_arry(i).trn, t_dbMainDatabase)
                    dtx(i).Description = pr_table!tt_descr
                    stx_charge = pr_table!tt_charge
                End If
                ''''''''''''''''''''''''''''''''
                If last_level = 0 Then
                    last_level = trn_arry(i).order_group
                End If
                If trn_arry(i).order_group = last_level Then
                    'Vijaya on 02/05/04 Magic#395302 Changed from
                    'If pr_table!tt_charge = "Y" Then to stx_charge
                    If stx_charge = "Y" Then
                        level_rate = level_rate + trn_arry(i).amount
                        j = i
                    End If
                    trn_arry(i).amount = 0
                    dtx(i).amount = 0
                Else
                    'Vijaya on 02/05/04 Magic#395302 Changed
                    'Because it be same as 4GL Code
                    If j > 0 Then
                        trn_arry(j).amount = level_rate
                        dtx(j).amount = de_tax_amt(j, tunits, trn_arry(j).amount, trn_arry(j).order_group, trn_arry(j).limit, trn_arry(j).ttype, trn_arry(j).charge)
                        de_tax_amount = de_tax_amount + dtx(j).amount
                    End If
                    'Vijaya on 02/05/04 Magic#395302 Changed
                    'Because it be same as 4GL Code
                    If stx_charge = "Y" Then
                        level_rate = trn_arry(i).amount
                        j = i
                    Else
                        level_rate = 0
                        j = 0
                    End If
                    trn_arry(i).amount = 0
                    dtx(i).amount = 0
                    last_level = trn_arry(i).order_group
                End If
            End If
        Next
        'last_level is not null
        If last_level <> 0 Then
            If j > 0 Then 'WJ 12/04/2001 Magic 339720
                trn_arry(j).amount = level_rate
                dtx(j).amount = de_tax_amt(j, tunits, trn_arry(j).amount, trn_arry(j).order_group, trn_arry(j).limit, trn_arry(j).ttype, trn_arry(j).charge)
                de_tax_amount = de_tax_amount + dtx(j).amount
            End If
        End If
    End If

    found_error = False: error_found = 0

    Exit Sub
    
errTrap:
    #If PROCESSING Then
        tfnErrHandler "detax", False
    #Else
        tfnErrHandler "detax"
    #End If
End Sub

Public Function dtx_count() As Integer
    dtx_count = trn_arry_cnt
End Function

Public Function get_a_dtx(w1 As Integer) As dtx_type
    get_a_dtx = dtx(w1)
End Function

Public Function de_tax_amt(trn_i As Long, tunits As Double, rate As Single, order_group As Integer, limit As Single, basis As String, charge_flag As String) As Double
    Dim i1 As Integer
    Dim wrk_limit As Single
    Dim wrk_units As Double
    Dim wrk_amount As Double
    Dim basis_amount As Double
    Dim freight_tmp As Double
    Dim sign_it As Integer

    If tunits < 0 Then
         tunits = Abs(tunits)
         sign_it = -1
    Else
        sign_it = 1
    End If

    If charge_flag <> "Y" Then
        de_tax_amt = 0#
        Exit Function
    End If

    Select Case Trim$(basis)
      Case "$"
        basis_amount = base_amount
      Case "F"
        freight_tmp = freight_amount
        If (freight_tmp < 0) Then
            freight_tmp = freight_tmp * -1
        End If
        basis_amount = freight_tmp + base_amount
        base_amount = freight_tmp + base_amount
        basis = "$"
      Case "A"
        basis_amount = 1
        base_amount = 1
        basis = "$"
      Case Else
        basis_amount = tunits
    End Select

    group_amount = group_amount + tax_amount
    If order_group <> last_order_group Then
        If last_order_group <> 0 Then
            base_amount = base_amount - group_amount
        End If
        group_amount = 0
    End If
    last_order_group = order_group

    If trn_i > 1 Then
        i1 = trn_i - 1
        If limit < 0 Then
            last_limit = 0 - trn_arry(i1).limit
        Else
            last_limit = trn_arry(i1).limit
        End If
    Else
        last_limit = 0
    End If

    gross_limit = Abs(gross_limit)

    Select Case True
    Case limit < 0
        wrk_limit = 0 - limit
        If gross_limit >= last_limit And gross_limit <= wrk_limit Then
            wrk_units = tunits
            wrk_amount = base_amount
        Else
            wrk_units = 0
            wrk_amount = 0
        End If
    Case limit > 0
        wrk_limit = limit
        wrk_units = 0
        wrk_amount = 0
        If basis_amount >= last_limit And basis_amount <= wrk_limit Then
            wrk_units = tunits - last_limit
            wrk_amount = base_amount - last_limit
        Else
            If basis_amount > wrk_limit Then
                wrk_units = wrk_limit - last_limit
                wrk_amount = wrk_limit - last_limit
            End If
        End If
    Case Else
        wrk_units = tunits
        wrk_amount = base_amount
    End Select

    If Trim$(basis) <> "$" Then
        tax_amount = wrk_units * rate
    Else
        If (rate + 1 = 0) Then
            tax_amount = 0
        Else
            tax_amount = wrk_amount - (wrk_amount / (rate + 1))
        End If
    End If

    tax_amount = dround(tax_amount) * sign_it

    de_tax_amt = tax_amount

End Function

Public Sub start_smmry()
    Dim i As Integer

    'Change array size from 5 to 50 t avoid runtime error 9. ticket 490248
    ReDim smm_tax(50)
    
    smm_cnt = 0

    ReDim grp_tax(50)
    grp_cnt = 0

    sum_fuels = 0
    For i = 1 To 25
        a_tax_total(i) = 0
    Next

    For i = 1 To 50
        sls_tax_arry(i).trn = 0
        sls_tax_arry(i).amount = 0
    Next
    sls_tax_cnt = 0
End Sub

Public Sub bld_smmry(prod As String, d_amt As Double, bunits As Double, gunits As Double, nunits As Double)

    Dim lr_header As Recordset  '   record like inv_header.*
    Dim lr_classes As Recordset  '   record like inv_classes.*
    Dim rec_fnd As Boolean

    Set lr_header = finv_header(prod, t_dbMainDatabase)
    rec_fnd = False
    If smm_cnt > 0 Then
        For i = 1 To smm_cnt
            If smm_tax(i).smm_class = lr_header!ivh_prodtcl Then
                smm_tax(i).smm_d_amt = smm_tax(i).smm_d_amt + d_amt
                smm_tax(i).smm_bunits = smm_tax(i).smm_bunits + bunits
                smm_tax(i).smm_gunits = smm_tax(i).smm_gunits + gunits
                smm_tax(i).smm_nunits = smm_tax(i).smm_nunits + nunits
                rec_fnd = True
                Exit For
            End If
        Next
    End If
    If (rec_fnd = False) Then
        smm_cnt = smm_cnt + 1
        i = smm_cnt
        smm_tax(i).smm_class = lr_header!ivh_prodtcl
        smm_tax(i).smm_d_amt = d_amt
        smm_tax(i).smm_bunits = bunits
        smm_tax(i).smm_gunits = gunits
        smm_tax(i).smm_nunits = nunits
    End If

    If (it_is_a_fuel(prod, t_dbMainDatabase)) Then
        Set lr_classes = finv_classes(lr_header!ivh_prodtcl, "T", t_dbMainDatabase)
        'david 04/19/2002
        If lr_classes Is Nothing Then
            sErr_ProductCode = Trim(prod)
            sErr_ProductClass = fnCStr(lr_header!ivh_prodtcl)
            found_error = True: error_found = 51
            Exit Sub
        End If
        If (IsNull(lr_classes!icl_fuel_group) Or lr_classes!icl_fuel_group = " ") Then
            Exit Sub
        End If
        sum_fuels = sum_fuels + nunits
        If grp_cnt > 0 Then
            For i = 1 To grp_cnt
                If grp_tax(i).grp_class = lr_classes!icl_fuel_group Then
                    grp_tax(i).grp_bunits = grp_tax(i).grp_bunits + nunits
                    Exit Sub
                End If
            Next
        End If
        grp_cnt = grp_cnt + 1
        i = grp_cnt
        grp_tax(i).grp_class = lr_classes!icl_fuel_group
        grp_tax(i).grp_bunits = nunits
        grp_tax(i).grp_tax = 0
    End If
End Sub

Private Sub fnd_smmry(prod As String)
    Dim d_amt As Double, bunits As Double, gunits As Double, nunits As Double

    Dim lr_header As Recordset  '    record like inv_header.*
    Set lr_header = finv_header(prod, t_dbMainDatabase)
    If smm_cnt > 0 Then
        For i = 1 To smm_cnt
            If smm_tax(i).smm_class = lr_header!ivh_prodtcl Then
                smm_ptr = i
                Exit Sub
            End If
        Next
    End If
    smm_ptr = smm_cnt + 1
End Sub

Public Sub fnd_grp_tax(prod As String)

    Dim lr_header As Recordset  '   record like inv_header.*
    Dim lr_classes As Recordset  '   record like inv_classes.*

    gas_gals = 0
    dsl_gals = 0
    Set lr_header = finv_header(prod, t_dbMainDatabase)
    Set lr_classes = finv_classes(lr_header!ivh_prodtcl, "T", t_dbMainDatabase)
    For i = 1 To grp_cnt
        If grp_tax(i).grp_class = dsl_fuel_group Then
            dsl_gals = dsl_gals + grp_tax(i).grp_bunits
        Else
            gas_gals = gas_gals + grp_tax(i).grp_bunits
        End If
    Next
    effective_sum_fuels = sum_fuels

    '# Return if not Texas routine
    If dsl_fuel_group = "NONE" Then
        Exit Sub
    End If

    If Abs(sum_fuels) < 7000 Or Abs(sum_fuels) >= 8000 Then
        Exit Sub
    End If
    If Abs(gas_gals) >= 7000 Then
        effective_sum_fuels = 8000
        Exit Sub
    End If
End Sub

Public Sub reset_grp_tax()
    Dim i As Integer

    For i = 1 To grp_cnt
        grp_tax(i).grp_tax = 0
    Next
End Sub

'# these are the loading fee routines
Public Sub start_ldfee()
    ldf_cnt = 0
    ldf_ptr = 0
End Sub

'# This function calculates the tax amt where the limit applies to
'# the total invoice and not just a line item.
Public Function sales_tax_limit(calc_tax As Single, limit As Single, trn As Long)
    Dim X As Integer
    Dim not_found As Boolean

    If calc_tax = 0 Then
        sales_tax_limit = calc_tax
        Exit Function
    End If
    not_found = True
    For X = 1 To sls_tax_cnt
        If sls_tax_arry(X).trn = trn Then
            not_found = False
            Exit For
        End If
    Next
    If not_found Then
        sls_tax_cnt = sls_tax_cnt + 1
        X = sls_tax_cnt
        sls_tax_arry(X).trn = trn
    End If
    sls_tax_arry(X).amount = sls_tax_arry(X).amount + calc_tax
    If sls_tax_arry(X).amount > limit Then
        calc_tax = calc_tax - sls_tax_arry(X).amount + limit
        sls_tax_arry(X).amount = limit
    End If
    sales_tax_limit = calc_tax
End Function

Private Function dround(v As Variant) As Variant
'    Debug.Print "v = "; v
    If v <> 0! Then
        dround = tfnRound(v, 2)
    Else
        dround = 0!
    End If
'    Debug.Print "dround = "; dround
End Function    ''dround

'Next function WJ added on 1/11/99
'Returns: "I" or "L" ----Cal Tax by Invoice or Line by Line
'Originally this is determined by parm110031 but now it related to trn number
Private Function fnGetTT_LIMIT_TYPE(ByVal lTrn As Long) As String
    Dim rsTemp As Recordset
    Dim strSQL As String
    Static nFieldTester As Integer
    Static nHoldSystemParm As String
    
    If nFieldTester = -99 Then ' do try to read the new field from the table
                               'since we already failed to do so
        fnGetTT_LIMIT_TYPE = nHoldSystemParm
        Exit Function
    End If
    
    On Error GoTo FieldNotExit
    'Vijaya on 02/06/04 Magic#395302
    If bExists_tx_header Then
        strSQL = "SELECT txd_lmt_type FROM tx_header, tx_detail"
        strSQL = strSQL & " WHERE txh_trn = txd_trn"
        strSQL = strSQL & " AND txh_trn = " & val(lTrn)
        strSQL = strSQL & " AND " & tfnDateString(t_tax_date, True) & " between txd_beg_date "
        strSQL = strSQL & " AND txd_end_date"
    Else
        strSQL = "SELECT tt_lmt_type FROM tx_table WHERE tt_trn=" & val(lTrn)
    End If
    ''''''''''''''''''''''''''''''''
    Set rsTemp = t_dbMainDatabase.OpenRecordset(strSQL, dbOpenSnapshot, dbSQLPassThrough)
    If rsTemp Is Nothing Then
        GoTo FieldNotExit
        Exit Function
    End If
    If rsTemp.RecordCount < 1 Then
        fnGetTT_LIMIT_TYPE = "I"
    Else
        'Vijaya on 02/06/04 Magic#395302
        If bExists_tx_header Then
            If Not IsNull(rsTemp!txd_lmt_type) Then
                fnGetTT_LIMIT_TYPE = Trim(rsTemp!txd_lmt_type)
            Else
                fnGetTT_LIMIT_TYPE = "I"
            End If
        Else
            If Not IsNull(rsTemp!tt_lmt_type) Then
                fnGetTT_LIMIT_TYPE = Trim(rsTemp!tt_lmt_type)
            Else
                fnGetTT_LIMIT_TYPE = "I"
            End If
        End If
    End If
    Exit Function
FieldNotExit:
    'Table not updated (new field tt-lmt_type) not exsit
    Dim pr_sys_parm As Recordset
    Set pr_sys_parm = fsys_parm(10031, t_dbMainDatabase)
    
    If pr_sys_parm Is Nothing Then
        nHoldSystemParm = "I"
    ElseIf pr_sys_parm.RecordCount < 1 Then
        nHoldSystemParm = "I"
    Else
        If IsNull(pr_sys_parm!parm_field) Then
            nHoldSystemParm = "I"
        ElseIf Trim(pr_sys_parm!parm_field) = "L" Then
            nHoldSystemParm = "L"
        Else
            nHoldSystemParm = "I"
        End If
    End If
    fnGetTT_LIMIT_TYPE = nHoldSystemParm
    nFieldTester = -99
    
    Err.Clear
    On Error GoTo 0
End Function

Private Function fnGetRecord(rsTemp As Recordset, strSQL As String, _
           nDB As Integer, szCallingFunction As String, Optional vShowMsg) As Long

    Dim strErrorString As String, bShowMsg As Boolean
    
    #If PROCESSING Then
        bShowMsg = False
    #Else
        bShowMsg = True
    #End If
    
    If Not IsMissing(vShowMsg) Then
        bShowMsg = vShowMsg
    End If
    
    On Error GoTo SQLError
    Select Case nDB
        Case 1  'REMOTEdb
            Set rsTemp = t_dbMainDatabase.OpenRecordset(strSQL, dbOpenSnapshot, dbSQLPassThrough)
        Case 2  'LOCALdb
            Set rsTemp = dbLocal.OpenRecordset(strSQL, dbOpenSnapshot)
    End Select

    If rsTemp.RecordCount > 0 Then
        rsTemp.MoveLast
        rsTemp.MoveFirst
    End If

    fnGetRecord = rsTemp.RecordCount
quitsub:
    On Error GoTo 0
    Exit Function
SQLError:
    Dim strMsg As String
    strMsg = "fnGetRecord," & szCallingFunction
    tfnErrHandler strMsg, strSQL, bShowMsg
    fnGetRecord = -1
    Resume quitsub
End Function

Private Function fnCStr(vTemp As Variant) As String

    If IsNull(vTemp) Then
        fnCStr = ""
    Else
        fnCStr = Trim(vTemp)
    End If
    
End Function

Private Function fnCheckNonStkMstr(rsIn As Recordset, _
                                   lProdlnk As Long, _
                                   lPrftCtr As Long) As Boolean
    
    Const SUB_NAME As String = "fnCheckNonStkMstr"
    Dim sSql As String
    
    sSql = "SELECT ponsm_gl_cost AS invt_gl_cost, 0 AS invt_gl_sales, " _
        + "0 AS invt_gl_invt FROM po_non_stock_mstr WHERE ponsm_link = " & lProdlnk _
        & " AND ponsm_prft_ctr = " & lPrftCtr
        
    fnCheckNonStkMstr = fnGetRecord(rsIn, sSql, nDBInformix, SUB_NAME, False) > 0
End Function

'Vijaya on 02/04/04 Magic#395302
'We don't need this function but we need to make
'sure if some Program doesn't have date and use this module
'call in form_load this function and pass today date
Public Sub set_tax_date(txdt As String, Optional bCheckGlSeries As Boolean = True)
    
    Static bAlready_check As Boolean
    
    #If Not NO_DST Then
        t_tax_date = tfnFormatDate(txdt)
        If bCheckGlSeries Then
            Set pr_gl_period = fgl_period(CDate(t_tax_date), t_dbMainDatabase)
            If Not pr_gl_period Is Nothing Then
                If pr_gl_period.RecordCount > 0 Then
                    sSeries = fnCStr(pr_gl_period!glp_series)
                Else
                    sSeries = "^^^^^"
                End If
            Else
                sSeries = "^^^^^"
            End If
        End If
        
        If Not bAlready_check Then
            'Changed the following code by JQ 06/11/2004
            'If the txd_order not in the tx_detail (database verison 3.28.2)
            'we need to use data the tx_table
            'Otherwise, customer will get problem if they don't update database!
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            'bExists_tx_header = fnCheckTableFieldExist("tx_header", "txh_code")
            bExists_tx_header = fnCheckTableFieldExist("tx_detail", "txd_order")
            bAlready_check = True
        End If
    #Else
        'IF Non-Date Sensitive Taxes we need to flag set to false
        'so it won't look tx_header and tx_detail.
        bExists_tx_header = False
    #End If
End Sub

Public Function fnCheckTableFieldExist(sTable As String, sField As String) As Boolean
    Const SUB_NAME As String = "fnCheckTableFieldExist"
    Dim strSQL As String
    Dim rsTemp As Recordset
    
    fnCheckTableFieldExist = False
    
    strSQL = "SELECT colname FROM systables, syscolumns WHERE systables.tabid = syscolumns.tabid"
    strSQL = strSQL + " AND tabname = " + tfnSQLString(sTable)
    If sField <> "" Then
        strSQL = strSQL + " AND colname = " + tfnSQLString(sField)
    End If
    
    If fnGetRecord(rsTemp, strSQL, nDBInformix, SUB_NAME) > 0 Then
        fnCheckTableFieldExist = rsTemp.RecordCount > 0
    End If
    Set rsTemp = Nothing
End Function

Public Function round_up_sls_tx(wk_amt As Double) As Single

    Dim frac_cents As Double
    Dim tax_amt As Double
    Dim sign_it As Integer
    Dim str_amt As String
    Dim nPos As Integer

    'make the wk_amt non-negative and compute the tax
    'apply the sign to tax amt after it is calculated
    If wk_amt < 0 Then
       sign_it = -1
       wk_amt = wk_amt * sign_it
    Else
        sign_it = 1
    End If

    'format the amt using 10 whole #, 6 decimal
    str_amt = Format(wk_amt, "0000000000.000000")
    tax_amt = Mid(str_amt, 1, 13)
    frac_cents = Mid(str_amt, 14, 4)
    ' Always add a penny if there is a fraction!
    If frac_cents > 0 Then
        tax_amt = tax_amt + 0.01
    End If
    round_up_sls_tx = tfnRound(tax_amt * sign_it, 2)
End Function

'#424531 02/23/2004 I change this back to public so that we can
'#compile version <= 3.28.00. We can change it to private after
'#later --      WJ
'Private Function ftx_table( _
                                sParm1 As Long, _
                                db As DataBase _
                                ) As Recordset
 Public Function ftx_table( _
                                sParm1 As Long, _
                                db As DataBase _
                                )
    Const SUB_NAME = "ftx_table"
    'returning pr_tx_table As Recordset
        
    Dim sSql As String
    Dim iIndex As Integer
    Dim rs As Recordset
    Const conDimension = 10
    Static rsArray() As Recordset

    On Error GoTo ftx_tableRedim
    iIndex = 1
    Do While Not (rsArray(iIndex) Is Nothing)
        If rsArray(iIndex)!tt_trn = sParm1 Then
            Set ftx_table = rsArray(iIndex)
            Exit Function   'return now with this "buffered" recordset
        End If
        iIndex = iIndex + 1
        If iIndex > conDimension Then Exit Do 'While  'Note: iIndex will be used below
    Loop 'While
    
    sSql = "SELECT * FROM tx_table WHERE tt_trn = " & QStr(sParm1)
    Set rs = fnMainmodsOpenRecordset(sSql, SUB_NAME, db)
    If rs Is Nothing Then
        Exit Function
    End If
    If rs.RecordCount > 0 Then
        rs.MoveFirst    'populate with first record
        #If cDEBUG Then
            On Error Resume Next
            While Not rs.EOF
                Debug.Print "ftx_table(" & QStr(sParm1) & ") = " & rs.Fields(0)
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
    
    Set ftx_table = rsArray(iIndex)
    Exit Function
    
ftx_tableRedim:
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
    
End Function    ''ftx_table
''''''''''''''''''''''''''''''''

'Vijaya on 11/13/09 #647512
Private Sub set_dflt_origin(ByVal iPrftCtr As Integer)
    Const SUB_NAME As String = "set_dflt_origin"
    
    Dim rsTemp  As Recordset
    Dim strSQL  As String
    
    If iPrftCtr < 0 Then
        Exit Sub
    End If
    
    strSQL = "SELECT city_state FROM sys_prft_ctr,sys_city "
    strSQL = strSQL & "  WHERE prft_city = city_code "
    strSQL = strSQL & "    AND prft_ctr  = " & tfnRound(iPrftCtr)
    If fnGetRecord(rsTemp, strSQL, nDBInformix, SUB_NAME) < 1 Then
        Exit Sub
    End If
    origin_state = Trim(rsTemp!city_state & "")
End Sub
'''''''''''''''''''''''''''
