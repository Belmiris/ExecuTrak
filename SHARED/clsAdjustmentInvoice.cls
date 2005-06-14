VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAdjustmentInvoice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum AdjustmentPayMethod
    HoldPayment = 1
    PayShort = 2
End Enum

Private adjPayMethod As AdjustmentPayMethod
Private bPayAsInvoiced As Boolean
Private iProfitCtr As Integer
Private lVendor As Long
Private sInvoiceNbr As String
Private lInvoiceLink As Long
Private lShiftLink As Long
Private sShiftDate As String
Private sAdjustmentCategory As String
Private lAdjustmentGLAcct As Long
Private lAcctPayableAcct As Long
Private lPeriod As Long
Private dCreditAmt As Double
Private dAdjustmentAmt As Double
Private dOpenAmt As Double
Private dPendingAmt As Double
Private dClearedAmt As Double
Private dAppliedAmt As Double
Private bAmtsCalculated As Boolean
Private sSQL As String
Private colAdjustments As clsAdjustments

Private Const OPEN_STATUS As String = "O"
Private Const SQL_NULL As String = "null"
Private Const ADJ_OPEN As String = "O"
Private Const ADJ_PENDING As String = "P"
Private Const ADJ_CLEARED As String = "C"
Private Const ADJ_APPLIED As String = "A"

Private Const SQL_SELECT_CATEGORY_PURCHASE_ACCT As String = _
    "select rsct_gl_cogs from rs_cat where rsct_catagory = '@category'"
    
Private Const SQL_SELECT_ALT_CATEGORY_PURCHASE_ACCT As String = _
    "select rscta_gl_cogs from rs_alt_cat where rscta_catagory = '@category'"

Private Const SQL_SELECT_ADJUSTMENTS As String = _
    " select pia_inv_lnk,  pia_prft_ctr,  pia_merch_date, " & _
    " pia_vendor, pia_inv_nbr,  pia_seq, pia_adj_type, " & _
    " pia_adj_status, pia_itm_code, pia_catagory, pia_qty_inv, " & _
    " pia_qty_recv,  pia_cost_inv,  pia_cost_pb, " & _
    " pia_ext_cost_inv,  pia_ext_cost_pb,  pia_credit_amt,    " & _
    " pia_new_inv_nbr, ph_shl, rsbb_bk_lnk, rsbi_grp_nbr " & _
    " from   p_inv_adj, p_item,  outer (rs_b_item, rs_b_store, rs_b_book) " & _
    " where pia_inv_lnk   = ph_lnk " & _
    " and   pia_vendor    = rsbi_vendor " & _
    " and   pia_itm_code  = rsbi_code " & _
    " and   rsbi_ic_lnk   = rsbb_ic_lnk " & _
    " and   rsbi_vendor   = rsbb_vendor " & _
    " and   rsbs_subbook  = rsbb_subbook " & _
    " and   rsbs_book     = rsbb_book " & _
    " and   rsbs_prft_ctr = pia_prft_ctr " & _
    " and   rsbs_vendor   = rsbi_vendor " & _
    " and   pia_inv_lnk  = @inv_lnk " & _
    " order by pia_seq "
    
Private Const SQL_SELECT_PAY_AS_INVOICE_FLAG As String = _
    " select stmh_pay_proc from st_merch_hdr " & _
    " where stmh_store = @profit_ctr " & _
    " and stmh_date = '@date' " & _
    " and stmh_vendor = @vendor " & _
    " and stmh_inv_nbr = '@inv_nbr' "
    
Private Const SQL_SELECT_GL_PERIOD As String = _
    " select glp_status, gljap_period, glp_end_dt " & _
    " from gl_jrnl_ap, p_item, gl_period " & _
    " where gljap_shl = ph_shl " & _
    " and gljap_period = glp_refpd " & _
    " and ph_vendor = @vendor " & _
    " and ph_invoice = '@inv_nbr' "

Public Function BuildInvoice(lInvoiceLnk As Long) As Boolean
    Dim rsTemp As Recordset
    Dim oAdjustment As clsAdjustment
    Dim bProcessedFirstRecord As Boolean
    
    sSQL = SQLParm(SQL_SELECT_ADJUSTMENTS, "@inv_lnk", lInvoiceLnk)
    
    If fnRecordset(rsTemp, sSQL) > 0 Then
        iProfitCtr = tfnRound(rsTemp("pia_prft_ctr"))
        lVendor = tfnRound(rsTemp("pia_vendor"))
        sInvoiceNbr = fnGetField(rsTemp("pia_inv_nbr"))
        lInvoiceLink = lInvoiceLnk
        lShiftLink = tfnRound(rsTemp("ph_shl"))
        sShiftDate = fnGetField(rsTemp("pia_merch_date"))
        
        Set colAdjustments = New clsAdjustments
        
        Do Until rsTemp.EOF
            
            Set oAdjustment = New clsAdjustment
            oAdjustment.Build Me, fnGetField(rsTemp("pia_catagory")), tfnRound(rsTemp("pia_qty_inv")), _
                                       tfnRound(rsTemp("pia_qty_recv")), tfnRound(rsTemp("pia_cost_inv"), 4), _
                                       tfnRound(rsTemp("pia_cost_pb"), 4), _
                                       tfnRound(rsTemp("pia_ext_cost_inv"), 4) - tfnRound(rsTemp("pia_ext_cost_pb"), 4), _
                                       fnGetField(rsTemp("pia_adj_status")), fnGetField(rsTemp("pia_seq")), fnGetField(rsTemp("pia_new_inv_nbr")), _
                                       tfnRound(rsTemp("pia_credit_amt"), 4), fnGetField(rsTemp("pia_itm_code")), _
                                       tfnRound(rsTemp("rsbb_bk_lnk")), tfnRound(rsTemp("rsbi_grp_nbr"))
                                       
            colAdjustments.Add oAdjustment
            
            rsTemp.MoveNext
        Loop
    End If
    
    Set rsTemp = Nothing
    
End Function

Public Function ClearAdjustments(sUserID As String) As Boolean
    Dim oAdjustment As clsAdjustment
    
    For Each oAdjustment In colAdjustments
        If Not oAdjustment.Clear(sUserID) Then
            Exit Function
        End If
    Next
    
    ClearAdjustments = True
    
End Function

Private Function BuildCreditInvoice() As Boolean
    Dim oAdjustment As clsAdjustment
    Dim sInvoiceNbr As String
        
    'Only open adjustments are selected.  We will take the first adjustment with
    'an credit invoice number and begin creating the credit invoice
    For Each oAdjustment In colAdjustments
        With oAdjustment
            If .CreditInvoiceNbr <> vbNullString Then
                sInvoiceNbr = .CreditInvoiceNbr
            Else
            
            End If
        End With
    Next
    
End Function

Public Function GetCategoryPurchaseAcct(sCategoryCode) As Long
    Dim sAcct As String
    
    sSQL = SQLParm(SQL_SELECT_CATEGORY_PURCHASE_ACCT, _
                          "@category", sCategoryCode)
    
    sAcct = fnQueryForField(sSQL, "rsct_gl_cogs")
    
    'If not in the RS_CAT table, check the RS_ALT_CAT table
    If sAcct = vbNullString Then
        sSQL = SQLParm(SQL_SELECT_ALT_CATEGORY_PURCHASE_ACCT, _
                          "@category", sCategoryCode)
    
        sAcct = fnQueryForField(sSQL, "rscta_gl_cogs")
    End If
    
    GetCategoryPurchaseAcct = CLng(sAcct)
    
End Function

Public Property Get PayAsInvoiced() As Boolean
    Static bRetrievedFromDB As Boolean
    Dim sPayAsInvoiced As String
    
    If Not bRetrievedFromDB Then
        sSQL = SQLParm(SQL_SELECT_PAY_AS_INVOICE_FLAG, _
                                 "@profit_ctr", iProfitCtr, _
                                 "@date", sShiftDate, _
                                 "@vendor", lVendor, _
                                 "@inv_nbr", sInvoiceNbr)
                                 
        sPayAsInvoiced = fnQueryForField(sSQL, "stmh_pay_proc")
        
        If sPayAsInvoiced = "Y" Then
            bPayAsInvoiced = True
        End If
        
        bRetrievedFromDB = True
    End If
    
    PayAsInvoiced = bPayAsInvoiced
End Property

Public Property Get PayMethod() As AdjustmentPayMethod
    Static bRetrievedFromDB As Boolean
    Dim sPayMethod As String
    
    If Not bRetrievedFromDB Then
        sSQL = SQLParm(SQL_SELECT_VENDOR_PAY_METHOD, _
                                "@vendor", lVendor)
                                 
        sPayMethod = fnQueryForField(sSQL, "pm_discrep_meth")
        
        If sPayMethod = "H" Then
            adjPayMethod = HoldPayment
        Else
            adjPayMethod = PayShort
        End If
        
    End If
    
    PayMethod = adjPayMethod
End Property

Public Property Get ProfitCenter() As Integer
    ProfitCenter = iProfitCtr
End Property

Public Property Get Vendor() As Long
    Vendor = lVendor
End Property

Public Property Get InvoiceNumber() As String
    InvoiceNumber = sInvoiceNbr
End Property

Public Property Get InvoiceLink() As Long
    InvoiceLink = lInvoiceLink
End Property

Public Property Get ShiftLink() As Long
    ShiftLink = lShiftLink
End Property

Public Property Get AdjustmentCategory() As String
    If sAdjustmentCategory = vbNullString Then
        sAdjustmentCategory = GetSysParm(3025)
    End If
    
    AdjustmentCategory = sAdjustmentCategory
End Property

Public Property Get GLAdjustmentAccount() As Long
    If lAdjustmentGLAcct = 0 Then
        lAdjustmentGLAcct = GetCategoryPurchaseAcct(Me.AdjustmentCategory)
    End If
    
    GLAdjustmentAccount = lAdjustmentGLAcct
End Property

Public Property Get GLPayableAccount() As Long
    If lAdjustmentGLAcct = 0 Then
       lAcctPayableAcct = GetSysParm(3008)
    End If
    
    GLPayableAccount = lAcctPayableAcct
End Property

Public Property Get GLPeriod() As Long
    
    sSQL = SQLParm(SQL_SELECT_GL_PERIOD, _
                          "@vendor", lVendor, _
                          "@inv_nbr", sInvoiceNbr)
              
    GLPeriod = fnQueryForField(sSQL, "gljap_period")
    
    If GLPeriod = Null Then
        GLPeriod = 0
    End If
                
End Property

Public Property Get Adjustments() As clsAdjustments
    Set Adjustments = colAdjustments
End Property

Public Property Get CreditAmt() As Double
    If dCreditAmt = 0 And Not bAmtsCalculated Then
        CalcAmts
    End If
    
    CreditAmt = dCreditAmt
End Property

Public Property Get AdjustmentAmt() As Double
    If dAdjustmentAmt = 0 And Not bAmtsCalculated Then
        CalcAmts
    End If
    
    AdjustmentAmt = dAdjustmentAmt
End Property

Public Property Get OpenAmt() As Double
    If dOpenAmt = 0 And Not bAmtsCalculated Then
        CalcAmts
    End If
    
    OpenAmt = dOpenAmt
End Property

Public Property Get PendingAmt() As Double
    If dPendingAmt = 0 And Not bAmtsCalculated Then
        CalcAmts
    End If
    
    PendingAmt = dPendingAmt
End Property

Public Property Get ClearedAmt() As Double
    If dClearedAmt = 0 And Not bAmtsCalculated Then
        CalcAmts
    End If
    
    ClearedAmt = dClearedAmt
    
End Property

Public Property Get AppliedAmt() As Double
    If dAppliedAmt = 0 And Not bAmtsCalculated Then
        CalcAmts
    End If
    
    AppliedAmt = dAppliedAmt

End Property

Private Sub CalcAmts()
    Dim oAdj As clsAdjustment
    
    For Each oAdj In colAdjustments
        dAdjustmentAmt = dAdjustmentAmt + oAdj.AdjustmentAmt
        dCreditAmt = dCreditAmt + oAdj.CreditAmt
        Select Case oAdj.AdjustmentStatus
            Case ADJ_OPEN
                dOpenAmt = dOpenAmt + IIf(oAdj.CreditAmt <> 0, oAdj.CreditAmt, oAdj.AdjustmentAmt)
            Case ADJ_PENDING
                dPendingAmt = dPendingAmt + IIf(oAdj.CreditAmt <> 0, oAdj.CreditAmt, oAdj.AdjustmentAmt)
            Case ADJ_CLEARED
                dClearedAmt = dClearedAmt + IIf(oAdj.CreditAmt <> 0, oAdj.CreditAmt, oAdj.AdjustmentAmt)
            Case ADJ_APPLIED
                dAppliedAmt = dAppliedAmt + IIf(oAdj.CreditAmt <> 0, oAdj.CreditAmt, oAdj.AdjustmentAmt)
        End Select
    Next
    
    bAmtsCalculated = True
    
End Sub

Private Sub Class_Terminate()
    Dim oAdjustment As clsAdjustment
    
    If Not colAdjustments Is Nothing Then
        For Each oAdjustment In colAdjustments
            Set oAdjustment = Nothing
        Next
        
        Set colAdjustments = Nothing
    End If
    
End Sub
