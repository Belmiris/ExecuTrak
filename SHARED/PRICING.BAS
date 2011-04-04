Attribute VB_Name = "ModPricing"
Attribute VB_Description = "PRICING.4GL translation 1.23@(#) Last Delta: 11/6/91"
Option Explicit
'***************************************************************************************
'
' Copyright (c) 2002 FACTOR, A Division of W.R.Hess Company
'
' Module name   : PRICING.BAS
'
' Here are the public functions:
' Public functions:
'  (1) set_wrk_date       -- If you do not set this date, it will be today's date
'                         (Inherited from preious VB routine)
'
'  (2) product_price      -- Entry point (Inherited from preious VB routine)
'  (3) Fuel_price         -- Entry point (Inherited from preious VB routine)
'  (4) prod_price         -- Entry point (Inherited from preious VB routine)
'
'  (5) fnGetOverRideFlag  -- Return a Y/N flag (Inherited from preious VB routine)
'  (6) fnGetPricingStatus -- New. Return 0 if successul; positive int if not.
'                            You may also get error messages.
'   Note: sProductCode is actually product code.
'
'  Date: 06/12/2002
'  Programmer: Weigong Jiang
'
'  Date: 03/01/2004 #399465 Weigong
'  (1)Remove all the commented out (06/12/2002,convert VB Pricing routine to 4GL
'  pricing routine)codes that are no longer used. You can alyways find them
'  from Source Safe.
'  (2)Implement the Time Sensitive Fuel cost routine
'**************************************************************************************

Private wrk_date As Date
Public Const PRC_ERR_NO_TABLEREC = -1  'inherited from old public variable
Private m_status As Integer
Private m_rtn_msg As String
Private m_OverRideFlag_YN As String

'------------ #WJ 399465 03/01/04 -----------------
'#Varibles for Time Sensitive Fuel cost routine
Public Const PRICE_FM_LAD_PROD_DELIMITER = "|"
Private m_fuel_load_datetime As String '#will be calered after each price call
Private Const FUEL_COST_CALCULATED = "COST_CAL_DONE" '#feature for possible use in the future
Private Const NEED_FM_LOAD_TIME_STATUS = 901 '#Work with syvprice.4ge. Indicating
                                             '#fuel loading time needed but not provided
Private Type TP_LOAD_TIME
    mv_nbr As Long
    lad_nbr As Long
    prod As String 'Lading Product
    dt   As String
    tm As String
    L_D As String '#Lading or Delivery
    cust As Long '#Used when L_D = "D" when cancel delivery add
End Type
Private m_aryLoadTime() As TP_LOAD_TIME  '#Hold the loading date/time
Private m_LoadTimeCnt As Long         '#Array counter
Private m_frmQueryTime As Form           '#Need to be set to frmQueryTime
                                         '#If not set, the popup feature is not
                                         '#Available. FM entry only
Private m_local As DataBase              '#Local Table. Used to query load date
                                         '#if not set, the feature is not available
Private m_bQueryTime_Active As Boolean           '#If false, the popup feature is not
'

'david 08/08/2006  #527819 NDR
#Const DEBUG_LOG = False
''''''''''''''''''''''''''''''
'
                                         '#Available.
'------------ #WJ 399465 03/01/04 -----------------

' #448713 - Chris Albrecht - 06/13/2005
' Add status flag parameter.  When status flag = "1", the ceiling price
' and factor pricing functions will not be used when calculating the
' market price
' 612568- Chris Miller - 6/23/2009
' Add sOverrideOD so that we can call syvprice.4gl and pricing.4gl using
' an override to Origin/Destination pricing as set in sys parm 6202
Public Function product_price( _
                                DEST_FLAG As String, _
                                sProductCode As String, _
                                ByVal lProfitCenter As Long, _
                                ByVal lCustomer As Long, _
                                dfQuantity As Double, _
                                dfCost As Double, _
                                lOrigin As Long, _
                                t_dbMainDatabase As DataBase, _
                                Optional lSite As Long = -1, _
                                Optional sInvoiceDate As String = "", _
                                Optional sErrMsg As String = "", _
                                Optional zzerc_contract_nbr As String = "", _
                                Optional sStatusFlag As String, _
                                Optional sOverrideOD As String = "" _
                             ) As Double
    ' dest_flag is I or E:
    'Modification for Schmitt Sales Net Pricing. Ma, 11/3/99
    product_price = sp_price(sProductCode, lProfitCenter, lCustomer, _
                                dfQuantity, dfCost, dfCost, Null, lOrigin, DEST_FLAG, _
                                t_dbMainDatabase, lSite, sInvoiceDate, sErrMsg, zzerc_contract_nbr, _
                                sStatusFlag, sOverrideOD)
    
End Function ' product_price

'Modification for Schmitt Sales Net Pricing. Ma, 11/3/99
'david 01/13/2010  #654066, #638753-NDR
'change DEST_FLAG to ORIGIN_DEST_FLAG
'The value pass into ORIGIN_DEST_FLAG will be LadingIEflag and DeliveryIEflag combine.
'i.e. II, IE, EI, and EE, if only one flag is passing in then both flag will be set to the same flag
'david 03/22/2011  #3276
'Added DeliveryDate and DeliveryTime
'david 03/22/2011  #3375-696170
'Added Vendor for Quote Pricing
Public Function Fuel_price( _
                                ByVal ORIGIN_DEST_FLAG As String, _
                                ByVal sProductCode As String, _
                                ByVal lProfitCenter As Long, _
                                ByVal lCustomer As Long, _
                                ByVal dfQuantity As Double, _
                                ByVal dfCost As Double, _
                                ByVal dfDiscountedCost As Double, _
                                ByVal lOrigin As Long, _
                                t_dbMainDatabase As DataBase, _
                                Optional ByVal lSite As Long = -1, _
                                Optional ByVal sInvoiceDate As String = "", _
                                Optional ByRef sErrMsg As String = "", _
                                Optional ByVal zzerc_contract_nbr As String = "", _
                                Optional ByVal lMvNbr As Long = 0, _
                                Optional ByVal sLadingProducts As String = "", _
                                Optional ByVal sDeliveryDate As String = "", _
                                Optional ByVal sDeliveryTime As String = "", _
                                Optional ByVal lVendor As Long = -1, _
                                Optional ByVal lDelivery As Long = 0 _
                             ) As Double
    Dim sLoadingDate As String
    Dim lMaxLading As Long
    Dim ORIGIN_FLAG As String
    Dim DEST_FLAG As String
    Dim bUseDeliveryTimeForPricing As Boolean
    
    If Len(ORIGIN_DEST_FLAG) = 2 Then
        ORIGIN_FLAG = Left(ORIGIN_DEST_FLAG, 1)
        DEST_FLAG = Right(ORIGIN_DEST_FLAG, 1)
    Else
        ORIGIN_FLAG = ORIGIN_DEST_FLAG
        DEST_FLAG = ORIGIN_DEST_FLAG
    End If
    
    If DEST_FLAG = "E" And lMvNbr > 0 And m_bQueryTime_Active Then
        '#If Fuel Loading date/time has been entered, use it to cal price
        sub_Check_and_Set_DateTime lMvNbr, sLadingProducts, sProductCode, lMaxLading, sLoadingDate
        
        'david 03/22/2011  #3276
        If sDeliveryDate <> "" And sDeliveryTime <> "" Then
            Dim strSQL As String
            Dim rsTemp As Recordset
            
            strSQL = "select parm_field from sys_parm where parm_nbr = 1545"
            
            If Price_fnOpenRecord(rsTemp, strSQL, t_dbMainDatabase) > 0 Then
                If Trim(rsTemp!parm_field & "") = "Y" Then
                    bUseDeliveryTimeForPricing = True
                    sLoadingDate = sDeliveryDate
                    set_fuel_load_datetime sDeliveryDate, sDeliveryTime
                End If
            End If
        End If
        ''''''''''''''''''''''''
    End If
    
    If Not bUseDeliveryTimeForPricing Then
        lDelivery = 0
    End If
    
    ' dest_flag is I or E:
    Fuel_price = sp_price(sProductCode, lProfitCenter, lCustomer, _
                          dfQuantity, dfCost, dfDiscountedCost, Null, lOrigin, ORIGIN_FLAG, _
                          t_dbMainDatabase, lSite, sInvoiceDate, sErrMsg, zzerc_contract_nbr, _
                          , , , ORIGIN_FLAG, lVendor)
    
    If DEST_FLAG = "E" And lMvNbr > 0 And m_bQueryTime_Active Then
        If fnNeed_ReCal_Price_With_Loading_Time(sProductCode, lMvNbr, lMaxLading, _
           sLoadingDate, lCustomer, lDelivery, sDeliveryTime) Then
           '#The following conditions must be satisfied in order to recalcualte price:
           '#(1)We want to do this recalcualtion. ie m_bQueryTime_Active must be true
           '#(2)Loading Time is NOT provided
           '#(3)We need time to do the calculation. The previous 4GE call will tell us this
            Fuel_price = sp_price(sProductCode, lProfitCenter, lCustomer, _
                                  dfQuantity, dfCost, dfDiscountedCost, Null, lOrigin, ORIGIN_FLAG, _
                                  t_dbMainDatabase, lSite, sInvoiceDate, sErrMsg, zzerc_contract_nbr, _
                                  , , , ORIGIN_FLAG, lVendor)
        End If
        '#Clear m_fuel_load_datetime
        set_fuel_load_datetime szEMPTY, szEMPTY, False
    End If
    
End Function ' product_price

Public Function prod_price( _
                                sProductCode As String, _
                                lProfitCenter As Long, _
                                lCustomer As Long, _
                                dfQuantity As Double, _
                                dfCost As Double, _
                                lOrigin As Long, _
                                t_dbMainDatabase As DataBase, _
                                Optional lSite As Long = -1, _
                                Optional sInvoiceDate As String = "", _
                                Optional sErrMsg As String = "", _
                                Optional zzerc_contract_nbr As String = "", _
                                Optional bTieredPricing As Boolean = False _
                            ) As Double
    ' dest_flag is always I:
    'Modification for Schmitt Sales Net Pricing. Ma, 11/3/99
    prod_price = sp_price(sProductCode, lProfitCenter, lCustomer, _
                                dfQuantity, dfCost, dfCost, Null, lOrigin, "I", _
                                t_dbMainDatabase, lSite, sInvoiceDate, sErrMsg, _
                                zzerc_contract_nbr, "", "", bTieredPricing)
End Function ' prod_price

Public Sub set_wrk_date(dDate As Date, t_dbMainDatabase As DataBase)
    wrk_date = dDate
End Sub

' #448713 - Chris Albrecht - 06/13/2005
' Add status flag parameter.  When status flag = "1", the ceiling price
' and factor pricing functions will not be used when calculating the
' market price
' 612568- Chris Miller - 6/23/2009
' Add sOverrideOD so that we can call syvprice.4gl and pricing.4gl using
' an override to Origin/Destination pricing as set in sys parm 6202
' david 11/30/2010  #3000
' Add sQuotePriceOriginFlag so that the quote pricing can handle E-Terminal and I-Profit Center
'david 03/22/2011  #3375-696170
'Added Vendor for Quote Pricing
Private Function sp_price(ByVal sProductCode As String, _
                            ByVal lProfitCenter As Long, _
                            ByVal lCustomer As Long, _
                            ByVal dfQuantity As Double, _
                            ByVal dfCost As Double, _
                            ByVal dfDiscountedCost As Double, _
                            ByVal vLevel As Variant, _
                            ByVal lOrigin As Long, _
                            ByVal DEST_FLAG As String, _
                            t_dbMainDatabase As DataBase, _
                            Optional lSite As Long = -1, _
                            Optional sInvoiceDate As String = "", _
                            Optional sErrMsg As String = "", _
                            Optional zzerc_contract_nbr As String = "", _
                            Optional sStatusFlag As String = "", _
                            Optional sOverrideOD As String = "", _
                            Optional bTieredPricing As Boolean = False, _
                            Optional sQuotePriceOriginFlag As String = "", _
                            Optional lVendor As Long = -1 _
                            ) As Double
                            
        Dim i As Integer
        'david 10/20/2010  #2478
        '#448713 - Increase from 15 to 16  '612568- increase from 16 to 17
        'david 03/22/2011  #3375-696170
        Const MAX_ARGS As Integer = 20
        Dim arg_val(MAX_ARGS) As String
        '''''''''''''''''''''''''''''''
        ''''''''''''''''''''''''
        Dim sRunCmd As String
        Dim sOutPutstr As String
        Static sta_cnt As Long
        
        arg_val(0) = "syvprice.4ge"
        arg_val(1) = tfnGetUserName
        arg_val(2) = "'" & Trim(sProductCode) & "'"
        arg_val(3) = CStr(lProfitCenter)
        arg_val(4) = CStr(lCustomer)
        arg_val(5) = CStr(dfQuantity)
        arg_val(6) = CStr(dfCost)
        arg_val(7) = CStr(dfDiscountedCost)
        arg_val(8) = tfnSQLString(vLevel)
        arg_val(9) = CStr(lOrigin)
        arg_val(10) = tfnSQLString(DEST_FLAG)
        
        If wrk_date = CDate("12:00:00 AM") Then
           wrk_date = Date
        End If
        arg_val(11) = tfnFormatDate(wrk_date)
        arg_val(12) = tfnRound(lSite)
        If Not IsDate(sInvoiceDate) Then
            sInvoiceDate = CStr(Date)
        End If
        arg_val(13) = tfnFormatDate(sInvoiceDate)
        If Trim(zzerc_contract_nbr) = szEMPTY Then
            arg_val(14) = "' '"
        Else
            arg_val(14) = tfnSQLString(Trim(zzerc_contract_nbr))
        End If
        
        '#Time sensitive Fuel Cost - used in FM Entry Programs
        If Trim(m_fuel_load_datetime) = szEMPTY Then
            arg_val(15) = "' '"
        Else
            arg_val(15) = tfnSQLString(Trim(m_fuel_load_datetime))
        End If
        
        ' #448713 - Chris Albrecht - 06/13/2005
        ' Add status flag parameter.  When status flag = "1", the ceiling price
        ' and factor pricing functions will not be used when calculating the
        ' market price
        If Trim(sStatusFlag) = szEMPTY Then
            arg_val(16) = "' '"
        Else
            arg_val(16) = tfnSQLString(Trim(sStatusFlag))
        End If
        
        ' 612568- Chris Miller - 6/23/2009
        ' Add sOverrideOD so that we can call syvprice.4gl and pricing.4gl using
        ' an override to Origin/Destination pricing as set in sys parm 6202
        If Trim(sOverrideOD) = szEMPTY Then
            arg_val(17) = "' '"
        Else
            arg_val(17) = tfnSQLString(Trim(sOverrideOD))
        End If
        
        'david 10/20/2010  #2478
        If bTieredPricing Then
            arg_val(18) = "'Y'"
        Else
            arg_val(18) = "'N'"
        End If
        ''''''''''''''''''''''''
        
        'david 11/30/2010  #3000
        If Trim(sQuotePriceOriginFlag) = szEMPTY Then
            arg_val(19) = "' '"
        Else
            arg_val(19) = tfnSQLString(Trim(sQuotePriceOriginFlag))
        End If
        ''''''''''''''''''''''''
        
        'david 03/22/2011  #3375-696170
        arg_val(20) = lVendor
        '''''''''''''''''''''''''''''''
        
        sRunCmd = arg_val(0)
        
        'david 11/30/2010  #3000
        'david 10/20/2010  #2478
        'For i = 1 To 17 '#448713 - Increase from 15 to 16  '612568- increase from 16 to 17
        For i = 1 To MAX_ARGS
        ''''''''''''''''''''''''
        ''''''''''''''''''''''''
           sRunCmd = sRunCmd & Space(1) & arg_val(i)
        Next
        
        sOutPutstr = fnRun4GLPricing(sRunCmd)
        
        Dim ary_out As Variant
        Dim nPos As Integer
        
        nPos = InStr(sOutPutstr, "syvprice~")
        
        'david 08/08/2006  #527819 NDR
        #If DEBUG_LOG Then
            tfnLog sRunCmd, DEBUG_LOG_PATH + "PRICING.LOG"
            tfnLog sOutPutstr, DEBUG_LOG_PATH + "PRICING.LOG"
        #End If
        ''''''''''''''''''''''''''''''
        
        If nPos > 0 Then
            If nPos > 1 Then 'should be ob 1st pos
                 sOutPutstr = Right(sOutPutstr, Len(sOutPutstr) - nPos + 1)
            End If
            ary_out = Split(sOutPutstr, "~")
            If UBound(ary_out) >= 7 Then
                m_status = tfnRound(ary_out(2))
                sp_price = tfnRound(ary_out(3), 8)
                m_OverRideFlag_YN = ary_out(4)
                m_rtn_msg = ary_out(7)
                If InStr(m_rtn_msg, "Unknown") > 0 Then
                    m_rtn_msg = "Cannot retrieve data from file " & m_rtn_msg & ". Error Code = " & m_status
                End If
            Else
                GoTo Handle_Error
            End If
        Else
             GoTo Handle_Error
        End If
        Exit Function
        
Handle_Error:
        sta_cnt = sta_cnt + 1
        m_status = -1
        m_OverRideFlag_YN = "N"
        m_rtn_msg = "Error in launaching 4GE"
        If sta_cnt < 10 Then
            sOutPutstr = "A message has been returned from the server:" & vbCrLf & sOutPutstr & vbCrLf & vbCrLf & "Command sent to server:" & vbCrLf & sRunCmd
            tfnErrHandler "sp_price", 401, sOutPutstr
        End If
        sp_price = 0
End Function

Public Function fnGetOverRideFlag() As String
    fnGetOverRideFlag = m_OverRideFlag_YN
End Function
Public Function fnGetPricingStatus(Optional sMsg As String) As Integer
    If Not IsMissing(sMsg) Then
        sMsg = m_rtn_msg
    End If
    fnGetPricingStatus = m_status
    
End Function

''' ============ >>>>>>>> 399465 wj 03/01/2004 Windows time sensitive
''                                             fuel cost
'#
'#!!!!!This part require a Form frmQueryTime and m_local database !!!!!!
'#If m_frmQueryTime or m_local not set, simply ignore this routine and do
'#fuel price as before.
'#
'#In what follows, we implement Date sensitive fuel cost pricing. It
'#will be used in fuel movement entry and its editor. syvprice.4ge has
'#also been modified for this project.
'#Entry Point:
'#(1)Public Sub Pricing_FuelCostSetQueryTimeForm
'#(2)Public Sub Pricing_AddToFuelLoadTimeAry
'#Call this once in main form load of FM Program
Public Sub Pricing_FuelCostSetQueryTimeForm(frm As Form, _
                                       vLocalDB As DataBase, _
                                       bQueryTime As Boolean)
    '#if m_bQueryTime_Active=true, m_frmQueryTime
    '# and m_local must be set
    '#You may want to set m_bQueryTime_Active = false if fuel price is
    '#called by a processor
    Set m_frmQueryTime = frm
    Set m_local = vLocalDB
    If m_frmQueryTime Is Nothing Then
       m_bQueryTime_Active = False
    Else
       m_bQueryTime_Active = bQueryTime
    End If
End Sub

'david 03/22/2011  #3276
Public Sub Pricing_ResetFuelLoadTimeAry()
    m_LoadTimeCnt = 0
    ReDim m_aryLoadTime(0)
End Sub

'#Call this function when you know the loading time. It is typically
'#after Query Time
Public Sub Pricing_AddToFuelLoadTimeAry(ByVal lMv As Long, _
                                        ByVal lLad As Long, _
                                        ByVal sPCode As String, _
                                        ByVal sDate As String, _
                                        ByVal sTime As String, _
                                        ByVal sLDflag As String, _
                                        ByVal lCustomer As Long)
    Dim i As Long

    For i = 0 To m_LoadTimeCnt - 1
        With m_aryLoadTime(i)
            If .mv_nbr = lMv And .lad_nbr = lLad And .prod = sPCode And .L_D = sLDflag Then
                .dt = Format(sDate, "mm/dd/yyyy")
                .tm = Format(sTime, "hh:mm")
                Exit Sub
            End If
        End With
    Next
    
    ReDim Preserve m_aryLoadTime(m_LoadTimeCnt) As TP_LOAD_TIME
    With m_aryLoadTime(m_LoadTimeCnt)
        .mv_nbr = lMv
        .lad_nbr = lLad
        .dt = Format(sDate, "mm/dd/yyyy")
        .tm = Format(sTime, "hh:mm")
        .prod = sPCode
        .L_D = sLDflag
        .cust = lCustomer
    End With
    m_LoadTimeCnt = m_LoadTimeCnt + 1
End Sub

'#Call this in Add Cancel (Lading/Delivey)
Public Sub Pricing_RemoveFromFuelLoadTimeAry(ByVal lMv As Long, _
                                        ByVal lLad As Long, _
                                        ByVal sLDflag As String, _
                                        ByVal lCustomer As Long)
    Dim i As Long

    For i = 0 To m_LoadTimeCnt - 1
        With m_aryLoadTime(i)
            If .mv_nbr = lMv And .L_D = sLDflag Then
                If sLDflag = "L" Then
                    If .lad_nbr = lLad Then
                        .mv_nbr = -1
                    End If
                Else
                    If .cust = lCustomer Then
                        .mv_nbr = -1
                    End If
                End If
            End If
        End With
    Next
End Sub

'#Set the loading Date and Time if already entered.
'#If blending, you will see more than one products.
'#Take the Max Lading Date/Time
Private Sub sub_Check_and_Set_DateTime(ByVal lMv As Long, _
                                       ByVal sLadingProducts As String, _
                                       ByVal sDelProd As String, _
                                       ByRef lMaxLading As Long, _
                                       ByRef sLoadingDate As String)
    Dim i As Long
    Dim j As Integer
    
    Dim strSQL As String
    Dim rsTemp As Recordset
    Dim vProd As Variant
    Dim sTemp As String
    Dim dDt As Date
    Dim sDT As String
    Dim sTm As String
    Dim bFound As Boolean
    
    Const AN_OLD_DATE = "01/01/1900"

    dDt = CDate(AN_OLD_DATE)
    
    If Trim(sLadingProducts) = szEMPTY Then
        Exit Sub
    End If
    sTm = szEMPTY
    sDT = szEMPTY
    bFound = False
    
    vProd = Split(sLadingProducts, PRICE_FM_LAD_PROD_DELIMITER)
    
    '#Time already entered? Delivery
    For i = 0 To m_LoadTimeCnt - 1
        With m_aryLoadTime(i)
            If .mv_nbr = lMv Then
                 If Trim(sDelProd) = Trim(.prod) And .L_D = "D" Then
                     If IsDate(.dt) Then
                        dDt = CDate(.dt)
                        sDT = .dt
                        sTm = .tm
                        lMaxLading = .lad_nbr
                        bFound = True
                        Exit For
                     End If
                 End If
            End If
        End With
    Next
    If Not bFound Then
        '#Time already entered? Lading
        For i = 0 To m_LoadTimeCnt - 1
            With m_aryLoadTime(i)
                If .mv_nbr = lMv Then
                    For j = 0 To UBound(vProd)
                        If Trim(vProd(j)) = Trim(.prod) Then
                            'Take the max one if blending
                            sTemp = Trim(.dt) & " " & Trim(.tm)
                            If IsDate(sTemp) Then
                               If dDt < CDate(sTemp) Then
                                    dDt = CDate(sTemp)
                                    sDT = .dt
                                    sTm = .tm
                                    lMaxLading = .lad_nbr
                                    bFound = True
                               End If
                            End If
                        End If
                    Next
                End If
            End With
        Next
    End If
    
    '#If already entered, use the date, time
    If dDt > CDate(AN_OLD_DATE) And sTm <> szEMPTY And bFound Then
        sLoadingDate = Format(dDt, "mm/dd/yyyy")
        set_fuel_load_datetime sLoadingDate, sTm
        Exit Sub
    End If
    
    lMaxLading = 0
    sLoadingDate = szEMPTY
    If Not m_local Is Nothing Then
        '#Figure out the loading Date. Also, put the Lading number
        sTemp = "('dummy product code'"
        For j = 0 To UBound(vProd)
            If Trim(vProd(j)) <> szEMPTY Then
                sTemp = sTemp & "," & tfnSQLString(vProd(j))
            End If
        Next
        sTemp = sTemp & ")"
        
        'david 06/03/2010  #662437-1838
        strSQL = "SELECT MLading.dPickupDate AS dt,MLading.lLading AS lad, szTime"
        '''''''''''''''''''''''''''''''
        strSQL = strSQL & " FROM MLading, MLadingDetail"
        strSQL = strSQL & " WHERE MLading.lMovement = MLadingDetail.lMovement"
        strSQL = strSQL & " AND MLading.lLading = MLadingDetail.lLading"
        strSQL = strSQL & " AND MLading.lMovement=" & lMv
        strSQL = strSQL & " AND szProductCode IN " & sTemp
        'david 06/03/2010  #662437-1838
        strSQL = strSQL & " ORDER BY MLading.dPickupDate DESC, szTime DESC"
        '''''''''''''''''''''''''''''''
        If Price_fnOpenRecord(rsTemp, strSQL, m_local) > 0 Then
            sLoadingDate = Format(rsTemp!dt & szEMPTY, "mm/dd/yyyy")
            If Not IsDate(sLoadingDate) Then
                sLoadingDate = szEMPTY
            'david 06/03/2010  #662437-1838
            Else
                sTm = Trim(rsTemp!szTime & "")
            '''''''''''''''''''''''''''''''
            End If
            lMaxLading = tfnRound(rsTemp!lad)
            set_fuel_load_datetime sLoadingDate, sTm
        End If
    End If
End Sub

Private Function fnNeed_ReCal_Price_With_Loading_Time(ByVal sDelvProd As String, _
                                                      ByVal lMv As Long, _
                                                      ByVal lLad As Long, _
                                                      ByVal sLoadingDate As String, _
                                                      ByVal lCustomer As Long, _
                                                      ByVal lDelivery As Long, _
                                                      ByRef sTime As String) As Boolean
     Dim strSQL As String
     Dim rsTemp As Recordset
     Dim sPDesc As String
     
     fnNeed_ReCal_Price_With_Loading_Time = False
     
     If Not IsDate(sLoadingDate) Then
        Exit Function
     End If
     
     If fnGetPricingStatus = NEED_FM_LOAD_TIME_STATUS Then
        strSQL = "SELECT ivh_desc FROM inv_header WHERE ivh_product = " & tfnSQLString(sDelvProd)
        If Price_fnOpenRecord(rsTemp, strSQL, t_dbMainDatabase) > 0 Then
            sPDesc = Trim(rsTemp!ivh_desc & szEMPTY)
            sTime = fnFMQueryTime(sDelvProd, sPDesc, sLoadingDate, lMv, lLad, lCustomer, lDelivery)
            set_fuel_load_datetime sLoadingDate, sTime
            
            fnNeed_ReCal_Price_With_Loading_Time = True
        End If
    End If
End Function

Private Function fnFMQueryTime(sPCode As String, _
                             sPDesc As String, _
                             sDate As String, _
                             ByVal lMv As Long, _
                             ByVal lLad As Long, _
                             ByVal lCustomer As Long, _
                             ByVal lDelivery As Long) As String
    If m_frmQueryTime Is Nothing Then
        Exit Function
    End If
    
    With m_frmQueryTime
        'set the form and labels caption
        If lDelivery > 0 Then
            .LadingOrDelivery = "D"
            .lblLading = lDelivery
        Else
            .LadingOrDelivery = "L"
            .lblLading = lLad
        End If
        
        .lblMovement = lMv
        .lblProductCode = sPCode
        .lblProductDesc = sPDesc
        .lblDate = sDate
        .txtTime = ""
        .Show vbModal
        fnFMQueryTime = .Time
        
        Pricing_AddToFuelLoadTimeAry lMv, lLad, sPCode, sDate, .Time, "D", lCustomer
    End With
End Function

''#The price routine will react to the variable fuel_load_datetime
''#when it calcualts the fuel cost.
''#There are three possible values for fuel_load_datetime char(16)
''#(1) null --  use wrk_date ti figure out cost as before
''#(2) mm/dd/yyyy hh:mm format
''#    if bytes 1-10 is not a valid date, wrk_date will be used
''#    if bytes 12-16 is not of format hh:mm, loading time will be null
''#(3) 'COST_CAL_DONE' - the cost passed in will be used.
''#                      this is for future use when a complicated
''#                      fuel cost already calcualted (such as blending)
'#Change it to public for Murphy FM Upload WJ 04/26/2004
Public Sub set_fuel_load_datetime(ByVal sFuelLoadDate As String, _
                                  ByVal sFuelLoadTime As String, _
                                  Optional bUseCalcualtedFmCost As Boolean = False)
    If bUseCalcualtedFmCost Then
        m_fuel_load_datetime = FUEL_COST_CALCULATED
    Else
        m_fuel_load_datetime = szEMPTY
        If sFuelLoadDate <> szEMPTY Then
            sFuelLoadDate = Format(sFuelLoadDate, "mm/dd/yyyy")
            If IsDate(sFuelLoadDate) Then
                 m_fuel_load_datetime = sFuelLoadDate & Space(1) & Format(sFuelLoadTime, "hh:mm")
            End If
        End If
    End If
End Sub

Private Function Price_fnOpenRecord(ByRef rsTemp As Recordset, _
                                    strSQL As String, objDB As DataBase) As Long
    On Error GoTo SQLError
    
    If objDB Is t_dbMainDatabase Then
        Set rsTemp = objDB.OpenRecordset(strSQL, dbOpenSnapshot, dbSQLPassThrough)
        If rsTemp.RecordCount > 0 Then
            rsTemp.MoveLast
            rsTemp.MoveFirst
        End If
    Else
        Set rsTemp = objDB.OpenRecordset(strSQL, dbOpenSnapshot)
    End If
    Price_fnOpenRecord = rsTemp.RecordCount
    Exit Function
SQLError:
    Price_fnOpenRecord = -1
End Function

'''<<<<============= 399465 wj  02/24/2004 Windows time sensitive fuel cost



