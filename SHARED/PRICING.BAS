Attribute VB_Name = "ModPricing"
Attribute VB_Description = "PRICING.4GL translation 1.23@(#) Last Delta: 11/6/91"
'***********************************************************************************
'
' Copyright (c) 1996 FACTOR, A Division of W.R.Hess Company
'
' Module name   : PRICING.BAS
' Date          : Feb 23, 1996
' Date(last mod): Mar 26, 1996
' Programmer(s) :
'
' This module is a 4GL-to-BAS translation of:
' @(#) pricing.4gl 1.23@(#) Last Delta: 11/6/91 }
' Copyright (c) 1988,1989 FACTOR, A Division of WR Hess Company   }
'
' Functions:    :   product_price   - external
'                   prod_price      - external (dest_flag always I)
'
'                   sp_price        - internal to product_price & prod_price
'
'                   set_wrk_date    (NOTE!!! At least one call to "set_wrk_date" must precede the
'                                    first call to functions (1) product_price, and/or (2) prod_price)
'
'                   book_pricing    - internal to sp_price
'
'                   main            - for testing if cMain = 1
'                   fnFirstCharU    - internal
'***********************************************************************************

Option Explicit
                        ' 1: Main sub is included FOR TESTING ONLY
#Const cMAIN = 0        ' 0:   "   "   " skipped by #If directive

                        ' 1: debug output is enabled
#Const cDEBUG = 0       ' 0: debug output is disabled

Public Const PRC_ERR_NO_TABLEREC = -1
'
' define "globals" set (as output) by routine "set_wrk_date" :
Private bProfitCenterActive As Boolean      'was Long
Private byGroupOnly As Byte                 'was Long
Private wrk_date As Date
Private cust_adv_action As String * 1
Private mstpc As Integer 'temporary until mainmods is coded

'
' define "globals" set (as output) by routine "book_pricing":
' note: "book_pricing is an internal function used by sp_price
Private bBook_active As Boolean
Private bBook_origin As Boolean
Private bBook_dest As Boolean
Private bBook_city As Boolean
Private bBook_term As Boolean

Private m_OverRideFlag_YN As String 'Add per magic 312713-48
'Added by junsong per magic 355816'
Private Const PRICE_BOOK_NOT_FOUND = "Price book was not set for customer"
Private m_sSysParm6203 As String

Public Function product_price( _
                                DEST_FLAG As String, _
                                sProductLink As String, _
                                ByVal lProfitCenter As Long, _
                                ByVal lCustomer As Long, _
                                dfQuantity As Double, _
                                dfCost As Double, _
                                lOrigin As Long, _
                                t_dbMainDatabase As Database, _
                                Optional lSite As Long = -1, _
                                Optional sInvoiceDate As String = "", _
                                Optional sErrMsg As String = "", _
                                Optional zzerc_contract_nbr As Long = -1 _
                             ) As Double
    ' dest_flag is I or E:
    'Modification for Schmitt Sales Net Pricing. Ma, 11/3/99
    product_price = sp_price(sProductLink, lProfitCenter, lCustomer, _
                                dfQuantity, dfCost, dfCost, Null, lOrigin, DEST_FLAG, _
                                t_dbMainDatabase, lSite, sInvoiceDate, sErrMsg, zzerc_contract_nbr)
    
    'add by junsong per magic 355816
    'If we can't find the price book for altnate customer, we try to find
    'the price book for its master customer
    
    If sErrMsg = PRICE_BOOK_NOT_FOUND And m_sSysParm6203 = "Y" Then
        Dim lMasterCust As Long
        Dim rsCust As Recordset
        
        Set rsCust = far_altname(lCustomer, t_dbMainDatabase)
        
        If Not (rsCust Is Nothing) Then
            lMasterCust = tfnRound(rsCust!an_cust)
            
            If lMasterCust <> lCustomer Then
                ' I changed from 'Y' to 'N' if we can not find the book for master vendor too,
                ' then I don't want it to return error, so I use the master customer level
                ' to calculate the price Junsong 02/26/01
                m_sSysParm6203 = "N"
                    
                product_price = sp_price(sProductLink, lProfitCenter, lMasterCust, _
                                dfQuantity, dfCost, dfCost, Null, lOrigin, DEST_FLAG, _
                                t_dbMainDatabase, lSite, sInvoiceDate, sErrMsg, zzerc_contract_nbr)
                m_sSysParm6203 = "Y"
    
            End If
            
        End If
        
    End If
    '
End Function ' product_price

'Modification for Schmitt Sales Net Pricing. Ma, 11/3/99
Public Function Fuel_price( _
                                DEST_FLAG As String, _
                                sProductLink As String, _
                                ByVal lProfitCenter As Long, _
                                ByVal lCustomer As Long, _
                                dfQuantity As Double, _
                                dfCost As Double, _
                                dfDiscountedCost As Double, _
                                lOrigin As Long, _
                                t_dbMainDatabase As Database, _
                                Optional lSite As Long = -1, _
                                Optional sInvoiceDate As String = "", _
                                Optional sErrMsg As String = "", _
                                Optional zzerc_contract_nbr As Long = -1 _
                             ) As Double
                        
    ' dest_flag is I or E:
    Fuel_price = sp_price(sProductLink, lProfitCenter, lCustomer, _
                                dfQuantity, dfCost, dfDiscountedCost, Null, lOrigin, DEST_FLAG, _
                                t_dbMainDatabase, lSite, sInvoiceDate, sErrMsg, zzerc_contract_nbr)
    
    'add by junsong per magic 355816
    'If we can't find the price book for altnate customer, we try to find
    'the price book for its master customer
    
    If sErrMsg = PRICE_BOOK_NOT_FOUND And m_sSysParm6203 = "Y" Then
        Dim lMasterCust As Long
        Dim rsCust As Recordset
        
        Set rsCust = far_altname(lCustomer, t_dbMainDatabase)
        
        If Not (rsCust Is Nothing) Then
            lMasterCust = tfnRound(rsCust!an_cust)
            
            If lMasterCust <> lCustomer Then
                ' I changed from 'Y' to 'N' if we can not find the book for master vendor too,
                ' then I don't want it to return error, so I use the master customer level
                ' to calculate the price Junsong 02/26/01
                m_sSysParm6203 = "N"
        
                Fuel_price = sp_price(sProductLink, lProfitCenter, lMasterCust, _
                                dfQuantity, dfCost, dfDiscountedCost, Null, lOrigin, DEST_FLAG, _
                                t_dbMainDatabase, lSite, sInvoiceDate, sErrMsg, zzerc_contract_nbr)
    
                m_sSysParm6203 = "Y"
            End If
            
        End If
        
    End If
    
End Function ' product_price

Public Function prod_price( _
                                sProductLink As String, _
                                lProfitCenter As Long, _
                                lCustomer As Long, _
                                dfQuantity As Double, _
                                dfCost As Double, _
                                lOrigin As Long, _
                                t_dbMainDatabase As Database, _
                                Optional lSite As Long = -1, _
                                Optional sInvoiceDate As String = "", _
                                Optional sErrMsg As String = "", _
                                Optional zzerc_contract_nbr As Long = -1 _
                            ) As Double
    ' dest_flag is always I:
    'Modification for Schmitt Sales Net Pricing. Ma, 11/3/99
    prod_price = sp_price(sProductLink, lProfitCenter, lCustomer, _
                                dfQuantity, dfCost, dfCost, Null, lOrigin, "I", _
                                t_dbMainDatabase, lSite, sInvoiceDate, sErrMsg, zzerc_contract_nbr)

    'add by junsong per magic 355816
    'If we can't find the price book for altnate customer, we try to find
    'the price book for its master customer
    
    If sErrMsg = PRICE_BOOK_NOT_FOUND And m_sSysParm6203 = "Y" Then
        Dim lMasterCust As Long
        Dim rsCust As Recordset
        
        Set rsCust = far_altname(lCustomer, t_dbMainDatabase)
        
        If Not (rsCust Is Nothing) Then
            lMasterCust = tfnRound(rsCust!an_cust)
            
            If lMasterCust <> lCustomer Then
                ' I changed from 'Y' to 'N' if we can not find the book for master vendor
                ' then I don't want it to return error, so I use the master customer level
                ' to calculate the price Junsong 02/26/01
                m_sSysParm6203 = "N"
                
                prod_price = sp_price(sProductLink, lProfitCenter, lMasterCust, _
                                dfQuantity, dfCost, dfCost, Null, lOrigin, "I", _
                                t_dbMainDatabase, lSite, sInvoiceDate, sErrMsg, zzerc_contract_nbr)
                m_sSysParm6203 = "Y"
                'change back to yes for other customer
            End If
            
        End If
        
    End If
    
End Function ' prod_price

Private Function sp_price(ByVal sProductLink As String, _
                            ByVal lProfitCenter As Long, _
                            ByVal lCustomer As Long, _
                            ByVal dfQuantity As Double, _
                            ByVal dfCost As Double, _
                            ByVal dfDiscountedCost As Double, _
                            ByVal vLevel As Variant, _
                            ByVal lOrigin As Long, _
                            ByVal DEST_FLAG As String, _
                            t_dbMainDatabase As Database, _
                            Optional lSite As Long = -1, _
                            Optional sInvoiceDate As String = "", _
                            Optional sErrMsg As String = "", _
                            Optional zzerc_contract_nbr As Long = -1 _
                            ) As Double
    

    Const SUB_NAME = "sp_price"

    'Double dfQuantity      ' was decimal[10] in the 4GL code
    Dim dfPrice As Double   ' dfPrice is the major container for this routine's output.
    
    ' define "general-purpose" recordsets, that is, those that contain fields which
    ' may be used throughout (in multiple places) the function:
    Dim pr_ar_altname As Recordset
    Dim pr_inv_header As Recordset
    Dim pr_inv_master As Recordset
    Dim pr_sys_prft_ctr As Recordset
        
    #If cDEBUG Then
        Debug.Print "function sp_price called"
    #End If
    
    m_OverRideFlag_YN = "N"
    
    dfPrice = 0#   'set DEFAULT value (floating zero)
    sErrMsg = ""
    
    ' for the parm sProductLink (product code), retrieve the inv_header record
    Set pr_inv_header = finv_header(sProductLink, t_dbMainDatabase)
    If pr_inv_header Is Nothing Then
        sp_price = dfPrice   '1, dfPrice=0
        sErrMsg = "Failed to validate product " + tfnSQLString(sProductLink)
        Exit Function       'cannot procede with invalid product
    End If
    
    '--------- start magic 312713-48 ----------------------------
    Dim strSQL As String
    Dim rsTemp As Recordset
    
    Static sParm_7041 As String
    
    Set pr_inv_master = Nothing
    If sParm_7041 = szEMPTY Then
        strSQL = "SELECT * FROM sys_parm WHERE parm_nbr= 7041"
        Set rsTemp = fnMainmodsOpenRecordset(strSQL, SUB_NAME, t_dbMainDatabase)
        If Not rsTemp Is Nothing Then
            If rsTemp.RecordCount > 0 Then
                If Not IsNull(rsTemp!parm_field) Then
                     sParm_7041 = UCase(Trim(rsTemp!parm_field))
                End If
            End If
        End If
        
        If sParm_7041 <> "Y" Then
            sParm_7041 = "N"
        End If
        
    End If
    If sParm_7041 = "Y" Then
        If Not IsMissing(sInvoiceDate) Then
            If IsDate(sInvoiceDate) Then
                strSQL = "SELECT invt_prft_ctr , invt_prodlnk, invt_gl_sales, invt_gl_cost,"
                strSQL = strSQL & "invt_secndloc, invt_rop, invt_goal, invt_formula, invt_sales_mtdd, "
                strSQL = strSQL & "invt_sales_mtdq, invt_sales_ytdd, invt_sales_ytdq, invt_cost_mtdd, invt_n_min, "
                strSQL = strSQL & "invt_n_max, invt_n_min_cost, invt_n_max_cost, invt_gl_invt, invt_cust_sernbr, "
                strSQL = strSQL & "invt_trk_cust_qoh, fodp_price AS invt_price, fodp_cost AS invt_cost,fodp_date"
                strSQL = strSQL & " FROM inv_master,fo_date_pricing WHERE invt_prft_ctr=fodp_prft_ctr AND invt_prodlnk=fodp_prodlnk"
                strSQL = strSQL & " AND invt_prodlnk=" & pr_inv_header!ivh_link & " AND invt_prft_ctr=" & CInt(lProfitCenter)
                strSQL = strSQL & " AND fodp_date < " & tfnDateString(sInvoiceDate, True) & " ORDER BY fodp_date DESC"
            
                Set pr_inv_master = fnMainmodsOpenRecordset(strSQL, SUB_NAME, t_dbMainDatabase)
                
                If Not pr_inv_master Is Nothing Then
                    If pr_inv_master.RecordCount = 0 Then
                        Set pr_inv_master = Nothing
                    End If
                End If
            End If
        End If
    End If
    
    ' for the parm lProfitCenter (profit center #) AND
    ' the product link from the header rec, retrieve the inv_master record:
    If pr_inv_master Is Nothing Then
        Set pr_inv_master = finv_master(pr_inv_header!ivh_link, CInt(lProfitCenter), t_dbMainDatabase)
    Else
        m_OverRideFlag_YN = "Y"
    End If
    '--------- end magic 312713-48 ------------------------------
    
    If pr_inv_master Is Nothing Then
        sp_price = dfPrice   '2, dfPrice=0
        sErrMsg = "Failed to validate product " + tfnSQLString(sProductLink) + ", " & lProfitCenter & " in inv_master table"
        Exit Function       'cannot procede with no master inventory record
    End If
    If dfCost = 0# Then dfCost = pr_inv_master!invt_cost 'use master rec "cost" if cost parm zero
    
    ''function stdconv was located in mainmods for the 4GL code -- now in refnfast;
    ''it returns a flag AND a value( dfConversionFactor )
    Dim byFlag As Byte, dfConversionFactor As Double
    byFlag = stdconv(pr_inv_header!ivh_uom_stock, pr_inv_header!ivh_uom_pricing, dfConversionFactor, t_dbMainDatabase)
    Select Case byFlag
        Case 0:
                    If dfConversionFactor <> 0# Then
                        dfCost = dfCost / dfConversionFactor
                    Else
                        dfCost = 0#
                    End If
        Case Else:  dfCost = pr_inv_master!invt_cost
    End Select
    
    '--------------------------------------------------------------------------------
    Call book_pricing(t_dbMainDatabase)   'get bBook_active boolean (and others if bBook_active)
    
    Do While bBook_active
        Dim sWorkLocation As String
        
        If bBook_origin Then
            If bBook_city Then
                If fnFirstCharU(DEST_FLAG) = "I" Then
                    Set pr_sys_prft_ctr = fsys_prft_ctr(lOrigin, t_dbMainDatabase)
                    If pr_sys_prft_ctr Is Nothing Then Exit Do 'while
                    sWorkLocation = pr_sys_prft_ctr!prft_city
                Else
                    Dim pr_terminal As Recordset
                    Set pr_terminal = ffm_terminal(lOrigin, t_dbMainDatabase)
                    If pr_terminal Is Nothing Then Exit Do 'while
                    sWorkLocation = pr_terminal!fmt_city
                    Set pr_terminal = Nothing
                End If
            End If                            'Format("5", 'using "<<<<<"
            If bBook_term Then sWorkLocation = Format(lOrigin, "####0")
        End If  'bBook_origin
        
        If bBook_dest Then
            If bBook_city Then
                Set pr_sys_prft_ctr = fsys_prft_ctr(lCustomer, t_dbMainDatabase)
                If pr_sys_prft_ctr Is Nothing Then
                    Set pr_ar_altname = far_altname(lCustomer, t_dbMainDatabase)
                    If pr_ar_altname Is Nothing Then
                        sp_price = dfPrice   '3, dfPrice
                        sErrMsg = "Failed to validate profit center/customer " + tfnSQLString(lCustomer)
                        Exit Function
                    End If
                    sWorkLocation = pr_ar_altname!an_city
                Else
                    sWorkLocation = pr_sys_prft_ctr!prft_city
                End If
            End If  'bBook_city               'Format("9", 'using "<<<<<<<<<"
            If bBook_term Then sWorkLocation = Format(lCustomer, "########0")
        End If  'bBook_dest
        
        ' define cursor strings:
        Dim c_last_term As String, c_book_price As String, c_purch_term As String
        Dim pr_sp_book As Recordset
        
        
        '////////////////////// START CHANGE FOR  #340618 WJ 02/14/02 East River //////////////////////
        Set pr_sp_book = Nothing
        
        If zzerc_contract_nbr >= 0 Then
            c_book_price = "SELECT zzerc_customer spb_customer, zzerc_loc spb_loc," & _
                           " zzerc_prodlnk spb_prodlnk,zzerc_date spb_date" & _
                           " zzerc_pterminal spb_pterminal,zzerc_pvendor spb_pvendor" & _
                           " zzerc_oper spb_oper,zzerc_price spb_price" & _
                           " FROM zzer_contract " & _
                           " WHERE zzerc_contract_nbr=" & tfnRound(zzerc_contract_nbr) & _
                           " AND zzerc_date <= " & QStr(wrk_date) & _
                           " ORDER BY spb_customer, spb_loc, spb_prodlnk, spb_date desc"
            'Ohhh, Order By is really not useful!
            Set pr_sp_book = fnMainmodsOpenRecordset(c_book_price, SUB_NAME)
            
            If Not pr_sp_book Is Nothing Then
                If pr_sp_book.RecordCount <= 0 Then
                    Set pr_sp_book = Nothing
                End If
            End If
        End If
        

        If pr_sp_book Is Nothing Then
            'Codes inside this if are the privious
            c_book_price = "SELECT * FROM sp_book" & _
                           " WHERE spb_customer = " & QStr(lCustomer) & _
                           " AND spb_loc = " & QStr(sWorkLocation) & _
                           " AND spb_prodlnk = " & QStr(pr_inv_master!invt_prodlnk) & _
                           " AND spb_date <= " & QStr(wrk_date) & _
                           " ORDER BY spb_customer, spb_loc, spb_prodlnk, spb_date desc"
                            
            ' open c_book_price
            ' fetch c_book_price into pr_sp_book
            
            Set pr_sp_book = fnMainmodsOpenRecordset(c_book_price, SUB_NAME)
        End If
        
        '////////////////////// END CHANGE FOR  #340618 WJ 02/14/02 East River //////////////////////
        
        If pr_sp_book Is Nothing Then
            sErrMsg = "Failed to get sp_book info"
            Exit Function
        End If
        
        If pr_sp_book.RecordCount > 0 Then
            pr_sp_book.MoveFirst 'populate with first record
        Else
            'changed by junsong 02/26/02
            If m_sSysParm6203 = "Y" Then
                
                If fnIsException(lCustomer, tfnRound(pr_inv_header!ivh_link), fnGetField(pr_inv_header!ivh_class)) Then
                    Exit Do
                Else
                    sErrMsg = PRICE_BOOK_NOT_FOUND
                    Exit Function
                End If
                
            Else
                Exit Do 'while bBook_active
            End If
            
        End If
        ' close c_book_price
        
        If fnFirstCharU(pr_sp_book!spb_oper) = "B" Then
            sp_price = pr_sp_book!spb_price + pr_inv_master!invt_price
            'david 07/11/2001
            'Exit Function
            GoTo CALC_CUST_PRICING_DISC
        End If
        If fnFirstCharU(pr_sp_book!spb_oper) = "=" Then
            sp_price = pr_sp_book!spb_price
            'david 07/11/2001
            'Exit Function
            GoTo CALC_CUST_PRICING_DISC
        End If
        If CVar(pr_sp_book!spb_pterminal) = Null Then
            sp_price = dfCost + pr_sp_book!spb_price
            'david 07/11/2001
            'Exit Function
            GoTo CALC_CUST_PRICING_DISC
        End If
        ' ///////////////////////////////
        
        '--------------------------------
        ' declare c_purch_term cursor for
        c_purch_term = "SELECT * FROM fm_purch_term" & _
                       " WHERE fmpt_terminal = " & QStr(pr_sp_book!spb_pterminal) & _
                       " AND fmpt_vendor = " & QStr(pr_sp_book!spb_pvendor) & _
                       " AND fmpt_prodlnk = " & QStr(pr_sp_book!spb_prodlnk)
                        
        ' open c_purch_term
        ' fetch c_purch_term into pr_purch_term
        Dim pr_purch_term As Recordset
        Set pr_purch_term = fnMainmodsOpenRecordset(c_purch_term, SUB_NAME)
        If pr_purch_term Is Nothing Then
            sErrMsg = "Failed to get fm_purch_term info"
            Exit Function
        End If
        If pr_purch_term.RecordCount > 0 Then
            pr_purch_term.MoveFirst 'populate with first record
        Else
            Exit Do 'while bBook_active
        End If
        ' close c_purch_term
                
        If pr_purch_term!fmpt_date <= wrk_date Then
            sp_price = pr_purch_term!fmpt_cost + pr_sp_book!spb_price
            'david 07/11/2001
            'Exit Function
            GoTo CALC_CUST_PRICING_DISC
        End If
        '--------------------------------
        
        '**************************************
        ' declare c_last_term scroll cursor for
        c_last_term = "SELECT * FROM fm_cost" & _
                      " WHERE fmc_terminal = " & QStr(pr_sp_book!spb_pterminal) & _
                      " AND fmc_vendor = " & QStr(pr_sp_book!spb_pvendor) & _
                      " AND fmc_prodlnk = " & QStr(pr_sp_book!spb_prodlnk) & _
                      " ORDER BY fmc_terminal,fmc_vendor,fmc_date,fmc_time,fmc_prodlnk"
        ' open c_last_term
        ' fetch last c_last_term into pr_cost
        Dim pr_cost As Recordset
        Dim VALIDRecordSet As Boolean
        Set pr_cost = fnMainmodsOpenRecordset(c_last_term, SUB_NAME)
        If pr_cost Is Nothing Then
            sErrMsg = "Failed to get fm_cost info"
            Exit Function
        End If
        If pr_cost.RecordCount > 0 Then
            pr_cost.MoveLast 'populate with last record
            VALIDRecordSet = True
        Else
            VALIDRecordSet = False
        End If
      
        While VALIDRecordSet    'this mneumonic replaces "<> NOTFOUND" in the 4GL
            'If pr_cost!fmc_date <= wrk_date Then
            If DateDiff("d", pr_cost!fmc_date, wrk_date) >= 0 Then
                sp_price = pr_cost!fmc_cost + pr_sp_book!spb_price
                'david 07/11/2001
                'Exit Function
                GoTo CALC_CUST_PRICING_DISC
            End If
            ' fetch previous c_last_term into pr_cost.*
            pr_cost.MovePrevious 'populate with previous record
            If pr_cost.BOF Then VALIDRecordSet = False
        Wend
        Exit Do 'while bBook_active
        '**************************************
        
    Loop 'While bBook_active
    '--------------------------------------------------------------------------------
    '
    ' either (1) book pricing is not active, or (2) a valid price was not determined;
    ' first, the pricing level must be determined:
    '
    If IsNull(vLevel) Then  'as currently implemented, parm vLevel is ALWAYS null;
        ' if lCustomer is a profit center, then we'll use it; otherwise we'll attempt to
        ' locate a valid customer in the customer master table.
        Set pr_sys_prft_ctr = fsys_prft_ctr(lCustomer, t_dbMainDatabase)
        If Not (pr_sys_prft_ctr Is Nothing) Then
            vLevel = pr_sys_prft_ctr!prft_lvl       '** use valid profit center's pricing level
        Else
            Set pr_ar_altname = far_altname(lCustomer, t_dbMainDatabase)
            If pr_ar_altname Is Nothing Then
                sp_price = dfPrice
                sErrMsg = "Failed to validate customer " + tfnSQLString(lCustomer)
                Exit Function   '3, dfPrice=0
            End If
            Dim pr_ar_customer As Recordset
            Set pr_ar_customer = far_customer(pr_ar_altname!an_cust, t_dbMainDatabase)
            If pr_ar_customer Is Nothing Then
                sp_price = dfPrice
                sErrMsg = "Failed to validate customer " + tfnSQLString(lCustomer) + " in ar_customer table"
                Exit Function   '4, dfPrice=0
            Else
                vLevel = pr_ar_customer!cust_lvl    '** use valid customer's pricing level
            End If
            
            ' now check the "exceptions":
            Dim rsException As Recordset
            Set rsException = fsp_excpt_cust(lCustomer, pr_inv_header!ivh_link, t_dbMainDatabase)
            If Not (rsException Is Nothing) Then
                vLevel = rsException!ce_lvl         '** use valid "product override" pricing level
                Set rsException = Nothing
            Else
                Set rsException = fsp_excpt_clas(lCustomer, pr_inv_header!ivh_class, t_dbMainDatabase)
                If Not (rsException Is Nothing) Then
                    vLevel = rsException!cce_lvl    '** use valid "class override" pricing level
                    'pr_ar_customer!cust_terms = rsException!cce_terms  'never used
                    Set rsException = Nothing
                End If
            End If
            Set pr_ar_customer = Nothing
        End If
    End If 'IsNull(vLevel)
    
    '
    ' Group calculations:
    '
    Dim pr_cust_grp As Recordset
    Dim pr_grp_price As Recordset
    Dim wrk_prft_ctr As Long
    Dim dfLowestPrice As Double
    Dim cLowestType As String * 1
    Dim c_groups As String, c_grp_price As String, c_1l As String   'cursor strings

    dfLowestPrice = 99999999999.99
    
    If byGroupOnly <> 3 Then
        
        ' declare c_groups cursor for
        c_groups = "SELECT * FROM sp_cust_grp" & _
                   " WHERE spcg_customer = " & QStr(lCustomer)
        
        If bProfitCenterActive Then
            wrk_prft_ctr = lProfitCenter
        Else
            wrk_prft_ctr = mstpc  ''mst_pc() ???? in mainmods
        End If
        
        Set pr_cust_grp = fnMainmodsOpenRecordset(c_groups, SUB_NAME)
        If pr_cust_grp Is Nothing Then
            sErrMsg = "Failed to get sp_cust_grp info"
            Exit Function
        End If
        With pr_cust_grp
            If .RecordCount > 0 Then
                .MoveFirst   'populate with first pr_cust_grp record
                While Not .EOF
                    'For Each ( c_groups into pr_cust_grp.* )
                        ' declare c_grp_price cursor for
                        c_grp_price = "SELECT * FROM sp_grp_price" & _
                                      " WHERE spgp_prodlnk = " & QStr(pr_inv_header!ivh_link) & _
                                      " AND spgp_group = " & QStr(!spcg_group) & _
                                      " AND spgp_prft_ctr = " & QStr(wrk_prft_ctr) & _
                                      " AND " & QStr(wrk_date) & " between spgp_start AND spgp_end"
                        
                        Set pr_grp_price = fnMainmodsOpenRecordset(c_grp_price, SUB_NAME)
                        If pr_grp_price Is Nothing Then
                            sErrMsg = "Failed to get sp_grp_price info"
                            Exit Function
                        End If
                        If pr_grp_price.RecordCount > 0 Then
                            pr_grp_price.MoveFirst 'populate with first pr_grp_price
                            While Not pr_grp_price.EOF
                                'For Each ( c_grp_price into pr_grp_price.* )
                                    If fnFirstCharU(pr_grp_price!spgp_flag) = "B" Then
                                        ' B = base inventory price + group price
                                        sp_price = pr_inv_master!invt_price + pr_grp_price!spgp_price
                                        'david 07/11/2001
                                        'Exit Function
                                        GoTo CALC_CUST_PRICING_DISC
                                    End If
                                    If pr_grp_price!spgp_price < dfLowestPrice Then
                                        dfLowestPrice = pr_grp_price!spgp_price
                                        cLowestType = UCase(pr_grp_price!spgp_flag)
                                    End If  ' V = value, P = percentage, or B = base+group
                                'Next
                                pr_grp_price.MoveNext   ' Locate next pr_grp_price record
                            Wend
                            Set pr_grp_price = Nothing
                        End If
                    'Next
                    .MoveNext    ' Locate next pr_cust_grp record
                Wend
                Set pr_cust_grp = Nothing
            End If
        End With 'pr_cust_grp
    End If  'byGroupOnly <> 3
    
    ' formula / group calculations
    If byGroupOnly = 1 Then
        ' group levels (fixed pricing) calculations
        'pr_cust_grp!spcg_group = Format(vLevel, "####0")
        'dfPrice = Null
        
        ' a slight modification (to spgp_group) is made to c_grp_price cursor:
        c_grp_price = "SELECT * FROM sp_grp_price" & _
                        " WHERE spgp_prodlnk = " & QStr(pr_inv_header!ivh_link) & _
                        " AND spgp_group = " & QStr(Format(vLevel, "####0")) & _
                        " AND spgp_prft_ctr = " & QStr(wrk_prft_ctr) & _
                        " AND " & QStr(wrk_date) & " BETWEEN spgp_start AND spgp_end"
        
        Set pr_grp_price = fnMainmodsOpenRecordset(c_grp_price, SUB_NAME)
        If pr_grp_price Is Nothing Then
            sErrMsg = "Failed to get sp_grp_price info"
            Exit Function
        End If
        
        With pr_grp_price
            If .RecordCount > 0 Then
                .MoveLast 'populate with last pr_grp_price
                While Not .EOF
                    dfPrice = !spgp_price
                    .MoveNext   'so that While/Wend terminates
                Wend
            End If
        End With 'pr_grp_price
    Else
        ' formula calculations
        ' declare c_1l cursor for
        c_1l = "SELECT * FROM sp_table" & _
                " WHERE pt_formula = " & QStr(pr_inv_master!invt_formula) & _
                " AND pt_lvl = " & QStr(vLevel) & _
                " ORDER BY pt_quantity ASC"
        Dim pr_sp_table As Recordset    'was an array of (10)
        Set pr_sp_table = fnMainmodsOpenRecordset(c_1l, SUB_NAME)
        If pr_sp_table Is Nothing Then
            sErrMsg = "Failed to get sp_table info"
            Exit Function
        End If
        If pr_sp_table.RecordCount > 0 Then
            pr_sp_table.MoveFirst 'populate
            Do While Not pr_sp_table.EOF
                If pr_sp_table!pt_quantity > dfQuantity Then Exit Do
                pr_sp_table.MoveNext
            Loop
            pr_sp_table.MovePrevious
            If pr_sp_table.BOF Then
                sp_price = dfPrice  '5, dfPrice
                'david 07/11/2001
                'Exit Function
                GoTo CALC_CUST_PRICING_DISC
            End If
        Else
            If dfPrice = 0 Then
                sp_price = PRC_ERR_NO_TABLEREC
            Else
                sp_price = dfPrice      '5, dfPrice
            End If
            'david 07/11/2001
            'Exit Function
            GoTo CALC_CUST_PRICING_DISC
        End If
        
        Select Case fnFirstCharU(pr_sp_table!pt_type)
            Case "F"
                dfPrice = pr_inv_master!invt_price   'inventory base
            Case "C"
                dfPrice = dfCost                     'cost
            'Modification for Schmitt Sales Net Pricing. Ma, 11/3/99
            Case "D"
                dfPrice = dfDiscountedCost
            Case Else
                sp_price = pr_sp_table!pt_factor    'use factor
                'david 07/11/2001
                'Exit Function
                GoTo CALC_CUST_PRICING_DISC
        End Select
        
        Select Case Left(LTrim(pr_sp_table!pt_operator), 1)
            Case "+"
                dfPrice = dfPrice + pr_sp_table!pt_factor
            Case "/"
                If pr_sp_table!pt_factor <> 0 Then
                    dfPrice = dfPrice / pr_sp_table!pt_factor
                Else
                    dfPrice = 0#
                End If
            Case "-"
                dfPrice = dfPrice - pr_sp_table!pt_factor
            Case "*"
                dfPrice = dfPrice * pr_sp_table!pt_factor
            Case "="
                dfPrice = pr_sp_table!pt_factor
            Case Else
                dfPrice = 0#   'stmt added during translation to VB
        End Select
    ' end of fixed or formula
    End If
    
    If dfLowestPrice = 99999999999.99 Then
        dfLowestPrice = dfPrice
    Else
        If cLowestType = "P" Then
            dfPrice = dfPrice * dfLowestPrice
            dfLowestPrice = dfPrice
        End If
    End If
    If (dfPrice < dfLowestPrice) And (cLowestType = "V") Then
        If cust_adv_action <> "A" Then dfLowestPrice = dfPrice
    End If
    
    sp_price = dfLowestPrice


'david 07/11/2001
'new feature!!!
'customer pricing discount and
'Ceiling Price consideration
CALC_CUST_PRICING_DISC:
    
    Dim nPriceFactor As Integer
    Dim pr_fo_product As Recordset
    
    Dim i As Long
    
    Dim nYYYY1 As Integer
    Dim nYYYY2 As Integer
    Dim sMMDDYYYY0 As String
    Dim sMMDDYYYY1 As String
    Dim sMMDDYYYY2 As String
    Dim dfCeilingPrice As Double
    
    If pr_ar_altname Is Nothing Then
        Set pr_ar_altname = far_altname(lCustomer, t_dbMainDatabase)
        If pr_ar_altname Is Nothing Then
            sErrMsg = "Failed to validate customer " + tfnSQLString(lCustomer)
            Exit Function
        End If
    End If
    
    If pr_ar_customer Is Nothing Then
        Set pr_ar_customer = far_customer(pr_ar_altname!an_cust, t_dbMainDatabase)
        If pr_ar_customer Is Nothing Then
            sErrMsg = "Failed to validate customer " + tfnSQLString(lCustomer) + " in ar_customer table"
            Exit Function
        End If
    End If
    
    nPriceFactor = tfnRound(pr_ar_customer!cust_pricefactor)
    
    If nPriceFactor > 0 Then
        sp_price = sp_price - tfnRound(sp_price * nPriceFactor / 100, DEFAULT_DECIMALS)
    End If
    
    'Ceiling Price consideration
    'Compare the final price with the ceiling price and
    'return the lowest price
    If lSite < 0 Then
        Exit Function
    End If
    
    If Not IsDate(tfnDateString(sInvoiceDate)) Then
        sErrMsg = "Invoice Date " + sInvoiceDate + " is not valid"
        Exit Function
    End If
    
    'checking fo_product table
    bDontShowError = True
    c_1l = "SELECT * FROM fo_product WHERE 1 <> 1"
    Set pr_fo_product = fnMainmodsOpenRecordset(c_1l, SUB_NAME)
    
    If pr_fo_product Is Nothing Then
        Exit Function
    End If
    
    bDontShowError = False
    
    'getting fo_product info
    c_1l = "SELECT * FROM fo_product" & _
        " WHERE fop_customer = " & lCustomer & _
        " AND fop_tank = " & lSite & _
        " AND fop_prodlnk = " & pr_inv_header!ivh_link & _
        " ORDER BY fop_start_date"
        '"AND " & tfnDateString(sInvoiceDate, True) & _
        " BETWEEN fop_start_date AND fop_end_date"
    Set pr_fo_product = fnMainmodsOpenRecordset(c_1l, SUB_NAME)
    
    If pr_fo_product Is Nothing Then
        sErrMsg = "Failed to get fo_product info"
        Exit Function
    End If
    
    If pr_fo_product.RecordCount > 0 Then
        sMMDDYYYY0 = tfnDateString(sInvoiceDate)
        nYYYY1 = tfnRound(Right(sMMDDYYYY0, 4))
        nYYYY2 = nYYYY1 + 1
        
        'locate the right period
        For i = 1 To pr_fo_product.RecordCount
            If Not (IsNull(pr_fo_product!fop_start_date) Or IsNull(pr_fo_product!fop_end_date)) Then
                If Len(pr_fo_product!fop_start_date) = 5 And Len(pr_fo_product!fop_end_date) = 5 Then
                    sMMDDYYYY1 = pr_fo_product!fop_start_date + "/" + Format(nYYYY1, "0000")
                    
                    If Left(pr_fo_product!fop_start_date, 2) > Left(pr_fo_product!fop_end_date, 2) Then
                        sMMDDYYYY2 = pr_fo_product!fop_end_date + "/" + Format(nYYYY2, "0000")
                    Else
                        sMMDDYYYY2 = pr_fo_product!fop_end_date + "/" + Format(nYYYY1, "0000")
                    End If
                    
                    'checking
                    If CDate(sMMDDYYYY0) >= CDate(sMMDDYYYY1) And CDate(sMMDDYYYY0) <= CDate(sMMDDYYYY2) Then
                        Exit For
                    End If
                End If
            End If
        Next i
        
        If Not pr_fo_product.EOF Then
            dfCeilingPrice = tfnRound(pr_fo_product!fop_ceiling_price, DEFAULT_DECIMALS)
            
            If sp_price > dfCeilingPrice Then
                sp_price = dfCeilingPrice
            End If
        End If
    End If
End Function ' sp_price pppppppppppppppppppppppppppppppppppppppppppppppppppppppppppp

'note: no internal callers; one external caller: "input_del_prod" in fmemvdel
'   this routine sets the following "global" vars, generally for use in
'   the "sp_price" function:
'       Public bProfitCenterActive As Boolean
'       Public byGroupOnly As Long
'       Public wrk_date As Date
'       Public cust_adv_action As String * 1
'       Public mstpc As Integer

Public Sub set_wrk_date(dDate As Date, t_dbMainDatabase As Database)
    Dim pr_sys_parm As Recordset
   
    #If cDEBUG Then
        Debug.Print "sub set_wrk_date called with parm: " & dDate
    #End If
    
    wrk_date = dDate
    
    Set pr_sys_parm = fsys_parm(2, t_dbMainDatabase) 'set_mst_master() 'from mainmods
    If Not (pr_sys_parm Is Nothing) Then mstpc = pr_sys_parm!parm_field
     
    Set pr_sys_parm = fsys_parm(6101, t_dbMainDatabase)
    If pr_sys_parm Is Nothing Then
        cust_adv_action = "L"
    Else
        cust_adv_action = fnFirstCharU(pr_sys_parm!parm_field)
    End If
            
    bProfitCenterActive = False
    Set pr_sys_parm = fsys_parm(6102, t_dbMainDatabase)
    If Not (pr_sys_parm Is Nothing) Then
        If fnFirstCharU(pr_sys_parm!parm_field) = "Y" Then bProfitCenterActive = True
    End If
    
    Set pr_sys_parm = fsys_parm(6103, t_dbMainDatabase) 'get STD pricing option
    If pr_sys_parm Is Nothing Then
        byGroupOnly = 3  'pr_sys_parm!parm_field = "P"
    Else
        Select Case fnFirstCharU(pr_sys_parm!parm_field)
            Case "Y":   byGroupOnly = 1
            Case "N":   byGroupOnly = 2
            Case Else:  byGroupOnly = 3
        End Select
    End If
    
    'added by junsong 02/26/02
    m_sSysParm6203 = "N"
    Set pr_sys_parm = fsys_parm(6203, t_dbMainDatabase)
    
    If Not (pr_sys_parm Is Nothing) Then
        
        If fnFirstCharU(pr_sys_parm!parm_field) = "Y" Then
            m_sSysParm6203 = "Y"
        End If
        
    End If
    'end added
    
End Sub ' set_wrk_date

Private Sub book_pricing(t_dbMainDatabase As Database)
    Const SUB_NAME = "book_pricing"
    
    Dim pr_sys_parm As Recordset
    Dim sSql As String
    
    #If cDEBUG Then
        Debug.Print "sub book_pricing called"
    #End If
    
    bBook_active = False 'this boolean is the primary return value
                         'if True, then one each of the following pairs is set True
    bBook_origin = False '  pair A
    bBook_dest = False   '  pair A
    bBook_city = False   '  pair B
    bBook_term = False   '  pair B
        
    Set pr_sys_parm = fsys_parm(6200, t_dbMainDatabase)
    If pr_sys_parm Is Nothing Then
        sSql = "INSERT INTO sys_parm VALUES(6200,'N','Automatic Setup Default N')" 'pr_sys_parm)
        fnMainmodsExecute sSql, SUB_NAME
        'Set pr_sys_parm = fnMainmodsOpenRecordset("sys_parm", dbOpenDynaset)
        'If pr_sys_parm Is Nothing Then
        '    Exit Function
        'End If
        'pr_sys_parm.Edit   ' Enable editing.
        'pr_sys_parm.Fields("parm_nbr") = 6200
        'pr_sys_parm.Fields("parm_field") = "N"
        'pr_sys_parm.Fields("parm_desc") = "Automatic Setup Default N"
        'On Error Resume Next    'in case we don't have priviledges
        'pr_sys_parm.Update ' Save changes.
    End If
    If fnFirstCharU(pr_sys_parm!parm_field) = "Y" Then
        bBook_active = True
        
        Set pr_sys_parm = fsys_parm(6202, t_dbMainDatabase)
        If pr_sys_parm Is Nothing Then
            sSql = "INSERT INTO sys_parm VALUES(6202,'D','Automatic Setup Default D')"
            fnMainmodsExecute sSql, SUB_NAME
        End If
        If fnFirstCharU(pr_sys_parm!parm_field) = "D" Then bBook_dest = True
        If fnFirstCharU(pr_sys_parm!parm_field) = "O" Then bBook_origin = True
                        
        Set pr_sys_parm = fsys_parm(6201, t_dbMainDatabase)
        If pr_sys_parm Is Nothing Then
            sSql = "INSERT INTO sys_parm VALUES(6201,'C','Automatic Setup Default C')"
            fnMainmodsExecute sSql, SUB_NAME
        End If
        If fnFirstCharU(pr_sys_parm!parm_field) = "C" Then bBook_city = True
        If fnFirstCharU(pr_sys_parm!parm_field) = "T" Then bBook_term = True
    
    End If 'pr_sys_parm!parm_field = "Y" (bBook_active = True)

End Sub ' book_pricing

Private Function fnFirstCharU(sParm As String) As String
    fnFirstCharU = UCase(Left(LTrim(sParm), 1)) 'return uppercase of leftmost char of a string
End Function    'fnFirstCharU

Public Function fnGetOverRideFlag() As String
    If m_OverRideFlag_YN = "Y" Then
        fnGetOverRideFlag = "Y"
    Else
        fnGetOverRideFlag = "N"
    End If
End Function

Private Function fnIsException(lCustomer As Long, lInvLink As Long, sClass As String) As Boolean
    Dim rsException As Recordset
    
    fnIsException = False
    Set rsException = fsp_excpt_cust(lCustomer, lInvLink, t_dbMainDatabase)
    
    If Not (rsException Is Nothing) Then
        fnIsException = True
    Else
        Set rsException = fsp_excpt_clas(lCustomer, sClass, t_dbMainDatabase)
        
        If Not (rsException Is Nothing) Then
            fnIsException = True
        End If
        
    End If

    
End Function

#If cMAIN Then
Sub Main()          'include for testing

    Dim t_szConnect As String            'holds ODBC connect string
    Dim t_engFactor As DBEngine          'pointer to database engine
    Dim t_wsWorkSpace As Workspace       'pointer to the default workspace
    Dim t_dbMainDatabase As Database     'main database handle

    #If cDEBUG Then
        Debug.Print " "
        Debug.Print "sub Main called @ " & Time
    #End If
    
     'first, open the database:
    t_szConnect = "ODBC;"
    Set t_engFactor = New DBEngine
    Set t_wsWorkSpace = t_engFactor.Workspaces(0)
    Set t_dbMainDatabase = t_wsWorkSpace.OpenDatabase("", False, False, t_szConnect)
    
    Dim DEST_FLAG As String * 1
    Dim sProductLink As String
    Dim lProfitCenter As Long
    Dim lCustomer As Long
    Dim dfQuantity As Double
    Dim dfCost As Double
    Dim lOrigin As Long
    
    #If True Then
        'miscellaneous testing
        #If True Then
            ' calculate price for one item:
            DEST_FLAG = "E"
            sProductLink = "09" '"07"
            lProfitCenter = 1300 '135
            lCustomer = 23040014
            dfQuantity = 3500 '1
            dfCost = 0.564
            lOrigin = 44
            'the following call is required in order that "globals" get set:
            Call set_wrk_date(DateAdd("yyyy", -1, Date), t_dbMainDatabase)     '1 year prior to current date
            
            Dim P3 As Double, P4 As Double
            P3 = product_price(DEST_FLAG, sProductLink, lProfitCenter, lCustomer, dfQuantity, dfCost, lOrigin, t_dbMainDatabase)
            P4 = prod_price(sProductLink, lProfitCenter, lCustomer, dfQuantity, dfCost, lOrigin, t_dbMainDatabase)
            '#If cDEBUG Then
            If Not bDontShowError Then
                Beep
                MsgBox "product_price(" & QStr(sProductLink) & ")=" & QStr(P3) & Chr(10) & _
                       "prod---_price(" & QStr(sProductLink) & ")=" & QStr(P4), _
                    vbInformation, "Test 2 callers of sp_price:"
            End If
            '#End If
        #End If
        
        #If False Then
            ret_cnt = create_temp_fmtank(DateAdd("yyyy", -1, Date), _
                                DateAdd("yyyy", -0, Date), t_dbMainDatabase)
                '#If cDEBUG Then
                If Not bDontShowError Then
                    Beep
                    MsgBox "value = " & QStr(ret_cnt), vbInformation, "Test of create_temp_fmtank"
                End If
                '#End If
        
            ret_cnt = create_temp_secloc(DateAdd("yyyy", -1, Date), _
                                DateAdd("yyyy", -0, Date), "F", t_dbMainDatabase)
                '#If cDEBUG Then
                    Beep
                    MsgBox "value = " & QStr(ret_cnt), vbInformation, "Test of create_temp_secloc (F)"
                '#End If
            Dim ret_cnt As Long
            ret_cnt = create_temp_secloc(DateAdd("yyyy", -1, Date), _
                                DateAdd("yyyy", -0, Date), "A", t_dbMainDatabase)
                '#If cDEBUG Then
                If Not bDontShowError Then
                    Beep
                    MsgBox "value = " & QStr(ret_cnt), vbInformation, "Test of create_temp_secloc (A)"
                End If
                '#End If
            start_of_program ("NULL") 'test mainmods module
            Call end_of_program
            gl_period_ok (DateAdd("yyyy", -1, Date))
            Stop
        #End If
    #End If
    
    #If False Then
        ' calculate prices for all items in inv_header: NOTE: this is a MAJOR test
        DEST_FLAG = "E"
        dfCost = 0
        dfQuantity = 1
        
        Const xsql As String = "SELECT * FROM inv_header"
        Dim pr_prod As Recordset, pr_mast As Recordset
        Set pr_prod = t_dbMainDatabase.OpenRecordset(xsql, dbOpenDynaset, dbSQLPassThrough)
        If pr_prod.RecordCount > 0 Then
            'the following call is required in order that "globals" get set:
            Call set_wrk_date(DateAdd("yyyy", -1, Date), t_dbMainDatabase)     '1 year prior to current date
            pr_prod.MoveFirst 'populate
            While Not pr_prod.EOF
                'For Each
                    sProductLink = pr_prod!ivh_product
                    Dim ysql As String
                    ysql = "SELECT * FROM inv_master" & _
                            " WHERE invt_prodlnk = " & QStr(pr_prod!ivh_link)
                    Set pr_mast = t_dbMainDatabase.OpenRecordset(ysql, dbOpenDynaset, dbSQLPassThrough)
                    If pr_mast.RecordCount > 0 Then
                        pr_mast.MoveFirst 'populate
                        'While Not pr_mast.EOF
                            lProfitCenter = pr_mast!invt_prft_ctr
                            Const zsql As String = "SELECT * FROM sys_prft_ctr"
                            Dim pr_prft As Recordset
                            Set pr_prft = t_dbMainDatabase.OpenRecordset(zsql, dbOpenDynaset, dbSQLPassThrough)
                            If pr_prft.RecordCount > 0 Then
                                pr_prft.MoveFirst 'populate
                                'While Not pr_prft.EOF
                                    lCustomer = pr_prft!prft_ctr
                                    Dim p1 As Double, P2 As Double
                                    p1 = product_price(DEST_FLAG, sProductLink, lProfitCenter, lCustomer, dfQuantity, dfCost, lOrigin, t_dbMainDatabase)
                                    P2 = prod_price(sProductLink, lProfitCenter, lCustomer, dfQuantity, dfCost, lOrigin, t_dbMainDatabase)
                                    #If True Then 'cDEBUG Then
                                        Debug.Print "product_price(" & sProductLink & lProfitCenter & " " & lCustomer & ")=" & Str(p1)
                                        Debug.Print "prod   _price(" & sProductLink & lProfitCenter & " " & lCustomer & ")=" & Str(P2)
                                        'Stop
                                    #End If
                                    pr_prft.MoveNext
                                'Wend
                            End If
                            'Stop
                            pr_mast.MoveNext 'populate
                        'Wend
                    End If
                'Next
                pr_prod.MoveNext    ' Locate next record.
            Wend
            'Stop
        End If
        Stop    'need this to view all the prices
    #End If
    
    t_dbMainDatabase.Close
    
End Sub 'Main mmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmm
#End If 'cMAIN




