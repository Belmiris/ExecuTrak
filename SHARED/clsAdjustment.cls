VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAdjustment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module     : clsAdjustment
' DateTime  : 6/2/2005 08:57
' Author      : Chris Albrecht
' Purpose    : Class used to clear and apply adjustment records
'                  for StoreTrak invoices
'Notes        : modcommon must be included in any projects
'                   this class is copied to
'---------------------------------------------------------------------------------------

Option Explicit

Private oInvoice As clsAdjustmentInvoice
Private sItemCode As String
Private lPriceBookLink As Long
Private lGroupLink As Long
Private sItemCategory As String
Private lInvoiceQty As Long
Private lPriceBookQty As Long
Private lQtyDiff As Long
Private dInvoiceCost As Double
Private dPriceBookCost As Double
Private dCostDiff As Double
Private sCreditInvoiceNbr As String
Private dCreditAmt As Double
Private dAdjustmentAmt As Double
Private sAdjustmentStatus As String
Private dUnitCost As Double
Private dCreditUnitCost As Double
Private bAmtsCalculated As Boolean
Private iInvSeq As Integer
Private sSQL As String

Private Const SQL_NULL As String = "null"
Private Const ADJ_OPEN As String = "O"
Private Const ADJ_PENDING As String = "P"
Private Const ADJ_CLEARED As String = "C"
Private Const ADJ_APPLIED As String = "A"

Public Enum AdjStatus
    OpenAdjustment = 0
    PendingAdjustment = 1
    ClearedAdjustment = 2
    AppliedAdjustment = 3
End Enum

Private Enum GLEntryType
    Debit = 0
    Credit = 1
End Enum

Private Const SQL_INSERT_RS_PURCLINE As String = _
    " insert into rs_purcline " & _
    " values (@shift_link, @vendor, '@inv_nbr', @line_num, " & _
    " @count, @cost, 0, @ext_cost, 0, '@category', null, @pb_link, @group_lnk) "

Private Const SQL_INSERT_GL_JRNL_AP As String = _
    " insert into gl_jrnl_ap " & _
    " values (0, @account, @period, @shift_link, @amount, '@flag', @inv_lnk, " & _
    " 0, @vendor, @profit_ctr, '@inv_nbr', today) "
    
Private Const SQL_UPDATE_ADJUSTMENT_STATUS As String = _
    " update p_inv_adj " & _
    " set pia_new_inv_nbr = '@credit_inv_nbr' " & _
    " ,pia_adj_status = '@status' " & _
    " ,pia_credit_amt = @credit_amt " & _
    " ,pia_user_id = '@user_id' " & _
    " where pia_inv_lnk = @inv_lnk " & _
    " and pia_seq = @seq "
    
Private Const SQL_SELECT_NEXT_LINE_NBR As String = _
    " select max(rsl_line_num) + 1 as new_line_nbr " & _
    " from rs_purcline " & _
    " where rsl_shl = @shift_link "
    
Public Property Get ItemCode() As String
    ItemCode = sItemCode
End Property
Public Property Get PriceBookLink() As Long
    PriceBookLink = lPriceBookLink
End Property
Public Property Get GroupLink() As Long
    GroupLink = lGroupLink
End Property
Public Property Get Category() As String
    Category = sItemCategory
End Property
Public Property Get InvoiceQty() As Long
    InvoiceQty = lInvoiceQty
End Property
Public Property Get PriceBookQty() As Long
    PriceBookQty = lPriceBookQty
End Property
Public Property Get InvoiceCost() As Double
    InvoiceCost = dInvoiceCost
End Property
Public Property Get PriceBookCost() As Double
    PriceBookCost = dPriceBookCost
End Property
Public Property Get AdjustmentAmt() As Double
    AdjustmentAmt = dAdjustmentAmt
End Property
Public Property Get AdjustmentStatus() As AdjStatus
    Select Case sAdjustmentStatus
        Case ADJ_OPEN: AdjustmentStatus = OpenAdjustment
        Case ADJ_PENDING: AdjustmentStatus = PendingAdjustment
        Case ADJ_CLEARED: AdjustmentStatus = ClearedAdjustment
        Case ADJ_APPLIED: AdjustmentStatus = AppliedAdjustment
    End Select
        
End Property
Public Property Get CreditInvoiceNbr() As String
    CreditInvoiceNbr = sCreditInvoiceNbr
End Property
Public Property Let CreditInvoiceNbr(sInvoiceNbr As String)
    sCreditInvoiceNbr = sInvoiceNbr
End Property
Public Property Get CreditAmt() As Double
    CreditAmt = dCreditAmt
End Property
Public Property Let CreditAmt(dAmt As Double)
    bAmtsCalculated = False
    dCreditAmt = CreditAmt
End Property
Public Property Get QtyDiff() As Long
    If lQtyDiff = 0 And Not bAmtsCalculated Then
        CalcAmts
    End If
    QtyDiff = lQtyDiff
End Property
Public Property Get CostDiff() As Double
    If dCostDiff = 0 And Not bAmtsCalculated Then
        CalcAmts
    End If
    CostDiff = dCostDiff
End Property
Public Property Get CreditUnitCost() As Double
    If dCreditUnitCost = 0 And Not bAmtsCalculated Then
        CalcAmts
    End If
    CreditUnitCost = dCreditUnitCost
End Property
Public Property Get UnitCost() As Double
    If dUnitCost = 0 And Not bAmtsCalculated Then
        CalcAmts
    End If
    UnitCost = dUnitCost
End Property
Public Property Get Invoice() As clsAdjustmentInvoice
    Set Invoice = oInvoice
End Property

Public Property Get Sequence() As Integer
    Sequence = iInvSeq
End Property

Private Sub CalcAmts()
    dCostDiff = dInvoiceCost - dPriceBookCost
    lQtyDiff = lInvoiceQty - lPriceBookQty
    
    If lQtyDiff <> 0 Then
        dCreditUnitCost = dCreditAmt / lQtyDiff
        dUnitCost = dAdjustmentAmt / lQtyDiff
    Else
        dCreditUnitCost = dCreditAmt / lInvoiceQty
        dUnitCost = dAdjustmentAmt / lInvoiceQty
    End If
    
End Sub

Public Sub Build(oAdjInvoice As clsAdjustmentInvoice, sCategory As String, lInvQty As Long, _
                                lPBQty As Long, dInvCost As Double, dPBCost As Double, dAdjAmt As Double, _
                                sAdjStatus As String, iSeq As Integer, Optional sCreditInv As String = vbNullString, _
                                Optional dCreditInvAmt As Double = 0#, Optional sItmCode As String = vbNullString, _
                                Optional lPBLnk As Long = 0, Optional lGroupLnk As Long = 0)
                                
    Set oInvoice = oAdjInvoice
    sItemCode = sItmCode
    sItemCategory = sCategory
    lInvoiceQty = lInvQty
    lPriceBookQty = lPBQty
    dInvoiceCost = dInvCost
    dPriceBookCost = dPBCost
    dAdjustmentAmt = dAdjAmt
    sAdjustmentStatus = sAdjStatus
    sCreditInvoiceNbr = sCreditInv
    dCreditAmt = dCreditInvAmt
    iInvSeq = iSeq
    lPriceBookLink = lPBLnk
    lGroupLink = lGroupLnk
    
   CalcAmts
        
End Sub
    
Public Function Clear(sUserID As String) As Boolean
    
    On Error GoTo ErrHandler
    
        'If Pay as Invoiced is No, then we need to adjust the invoice and GL tables
        If oInvoice.PayAsInvoiced = False Then
            ' Add the Item Code / Category to the RS_PURCLINE
            InsertInvoice IIf(lQtyDiff <> 0, lQtyDiff, lInvoiceQty), dUnitCost, sItemCategory, lPriceBookLink, lGroupLink
            
            'Remove the Adjustment Category from the RS_PURCLINE
            InsertInvoice 1, dAdjustmentAmt * -1, oInvoice.AdjustmentCategory, SQL_NULL, SQL_NULL
            
            'Debit the Item Category from the GL_JRNL_AP
            InsertGL oInvoice.GLAdjustmentAccount, oInvoice.GLPeriod, dAdjustmentAmt, Debit
            
            'Credit the Adjustment Category from the GL_JRNL_AP
            InsertGL oInvoice.GLPayableAccount, oInvoice.GLPeriod, dAdjustmentAmt, Credit
        End If
        
        SaveAdjustment sUserID, ClearedAdjustment
        
        Clear = True
        
        Exit Function
    
ErrHandler:
    Clear = False
    
End Function

Private Function InsertInvoice(lQty As Long, dCost As Double, _
                                            sCategory, lPBLnk, lGroupLnk) As Boolean
    
    sSQL = SQLParm(SQL_INSERT_RS_PURCLINE, _
                                "@shift_link", oInvoice.ShiftLink, _
                                "@vendor", oInvoice.Vendor, _
                                "@inv_nbr", oInvoice.InvoiceNumber, _
                                "@line_num", GetNewPurclineLineNbr(), _
                                "@count", lQty, _
                                "@cost", dCost, _
                                "@ext_cost", dCost * lQty, _
                                "@category", sCategory, _
                                "@pb_link", lPBLnk, _
                                "@group_lnk", lGroupLnk)
                          
    InsertInvoice = fnExecSQL(sSQL)
    
End Function

Private Function GetNewPurclineLineNbr() As Integer
    Dim sLineNbr As String
    
    sSQL = SQLParm(SQL_SELECT_NEXT_LINE_NBR, _
                          "@shift_link", oInvoice.ShiftLink)

    sLineNbr = fnQueryForField(sSQL, "new_line_nbr")
    
    If sLineNbr = vbNullString Then
        GetNewPurclineLineNbr = 1
    Else
        GetNewPurclineLineNbr = CInt(sLineNbr)
    End If
    
End Function

Private Function InsertGL(lacct As Long, lPeriod As Long, dAmt As Double, entryType As GLEntryType) As Boolean
    Dim sType As String
    
    If entryType = Credit Then
        sType = "C"
    Else
        sType = "D"
    End If
    
    sSQL = SQLParm(SQL_INSERT_GL_JRNL_AP, _
                                "@account", lacct, _
                                "@period", lPeriod, _
                                "@shift_link", oInvoice.ShiftLink, _
                                "@amount", dAmt, _
                                "@flag", sType, _
                                "@inv_lnk", oInvoice.InvoiceLink, _
                                "@vendor", oInvoice.Vendor, _
                                "@profit_ctr", oInvoice.ProfitCenter, _
                                "@inv_nbr", oInvoice.InvoiceNumber)
                                
    InsertGL = fnExecSQL(sSQL)
    
End Function

Public Function SaveAdjustment(sUserID As String, Status As AdjStatus) As Boolean
    Dim sStatus As String
    
    If IsMissing(sStatus) Then
        sStatus = sAdjustmentStatus
    Else
        Select Case Status
            Case OpenAdjustment: sStatus = ADJ_OPEN
            Case PendingAdjustment: sStatus = ADJ_PENDING
            Case ClearedAdjustment: sStatus = ADJ_CLEARED
            Case AppliedAdjustment: sStatus = ADJ_APPLIED
        End Select
    End If
    
    sSQL = SQLParm(SQL_UPDATE_ADJUSTMENT_STATUS, _
                          "@credit_inv_nbr", sCreditInvoiceNbr, _
                          "@credit_amt", dCreditAmt, _
                          "@status", sStatus, _
                          "@user_id", sUserID, _
                          "@inv_lnk", oInvoice.InvoiceLink, _
                          "@seq", iInvSeq)

    SaveAdjustment = fnExecSQL(sSQL)
    
End Function

Private Sub Class_Terminate()
    Set oInvoice = Nothing
End Sub
