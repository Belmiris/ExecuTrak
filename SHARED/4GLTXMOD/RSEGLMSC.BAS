Attribute VB_Name = "rseglmsc"
Option Explicit

Public Sub fuel_gl()
    Dim pc As Integer
    Dim acct As Long
    Dim dc_flag As String * 1
    Dim amt As Double
    Dim i As Integer, j As Integer, k As Integer
    Dim ti As Integer, tj As Integer, tk As Integer
    Dim arry_cnt As Integer
    Dim pr_ar_class As Recordset  'record like ar_class.*
    Dim meter_total As Double
    Dim meter_gl As Double
    Dim meter_max As Double
    Dim meter_acct As Long
    Dim fuel_only As Double
    Dim tax_only As Double
    Dim pr_trn As trn_type
    
    On Error GoTo errTrap
    
    initialize_tax_buffers
    
    Debug.Print "Preparing G/L & Tax Information"
    
    'david 11/05/2001
    #If Not ZZEMMASU Then
        'david 06/15/2001
        subShowStatusMsg "Preparing G/L & Tax Information"
    #End If
    ''''''''''''''''''''
    
    Set pr_ar_class = far_class("PRFTC", t_dbMainDatabase)
    initialize_gl
    
    ReDim pa_ftax(60)
    
    pa_ftax_cnt = 0

    start_smmry
    
    meter_total = 0
    
    #If PROCESSING Then
        nCurrIndex = 0
    #End If
    
    'got everything in aryComp, so do this
    Dim ivh_product As String
    
    'david 11/05/2001
    #If Not ZZEMMASU Then
        'david 06/15/2001
        subShowStatusMsg "Collecting Blend Info"
    #End If
    ''''''''''''''''''''
        
    For j = 0 To UBound(aryComp)
        ivh_product = aryComp(j).sProduct
        ivh_product = Left$(ivh_product, InStr(ivh_product, ",") - 1)
        bld_smmry ivh_product, aryComp(j).dfCompDollars, aryComp(j).dfCompUnits, _
            aryComp(j).dfCompUnits, aryComp(j).dfCompUnits
    Next j
    
    For j = 0 To UBound(aryComp)
        If Not (tfnRound(aryComp(j).dfCompDollars, 2) = 0 And aryComp(j).dfCompUnits = 0) Then
Debug.Print "PumpNo: "; aryComp(j).sPump; "Seq: "; aryComp(j).sSeq; "Product: "; aryComp(j).sProduct
            ivh_product = aryComp(j).sProduct
            ivh_product = Left$(ivh_product, InStr(ivh_product, ",") - 1)
            meter_total = meter_total + aryComp(j).dfCompDollars
            Set pr_inv_master = finv_master(aryComp(j).lProductLink, lProfitCenter, t_dbMainDatabase)
            Set pr_inv_header = fproduct(aryComp(j).lProductLink, t_dbMainDatabase)
            
            'david 11/05/2001
            #If Not ZZEMMASU Then
                'david 06/15/2001
                subShowStatusMsg "Calculating Fuel Taxes - " & ivh_product
            #End If
            ''''''''''''''''''''
            
            tax_only = compute_fuel_taxes(j) 'pass
            
            If found_error Then
                If error_found = 29 Then
                    #If PROCESSING Then
                        #If DEVELOP Then
                            MsgBox "Tax Use Group: " & aryComp(j).sUseGroup & " and Tax Class: " & aryComp(j).sTaxClass & _
                            " for Pump Number: " & aryComp(j).sPump & ", Pump Seq: " & aryComp(j).sSeq & _
                            ", Product: " & aryComp(j).sProduct & " Not Found in tx_use_g Table", vbCritical
                        #End If
                        subPostError "fuel_gl", "TaxUseG: " & aryComp(j).sUseGroup & _
                            ",Tax Class: " & aryComp(j).sTaxClass & " for Pump: " & _
                            aryComp(j).sPump & ", Seq: " & aryComp(j).sSeq & ", Product: " & _
                            aryComp(j).sProduct & " Not Found in tx_use_g Table.", nCurrIndex
                    #Else
                        MsgBox "Tax Use Group: " & aryComp(j).sUseGroup & " and Tax Class: " & aryComp(j).sTaxClass & _
                        " for Pump Number: " & aryComp(j).sPump & ", Pump Seq: " & aryComp(j).sSeq & _
                        ", Product: " & aryComp(j).sProduct & " Not Found in tx_use_g Table", vbCritical
                    #End If
                Else
                    fueltax_errors error_found
                End If
                Exit Sub
            End If
            
            tax_only = dround(tax_only)
            fuel_only = aryComp(j).dfCompDollars
            fuel_only = fuel_only - tax_only
    
            'david 11/05/2001
            #If Not ZZEMMASU Then
                'david 06/15/2001
                subShowStatusMsg "Preparing G/L (1) - " & lProfitCenter & ", " & _
                    ivh_product & ", " & sProfitUseGroup
            #End If
            ''''''''''''''''''''
    
            intrn_bl_gl lProfitCenter, ivh_product, sProfitUseGroup, _
                                     aryComp(j).dfCompUnits, _
                                     aryComp(j).dfCompUnits, _
                                     aryComp(j).dfCompUnits, _
                                     fuel_only, 0, 0

    
            If found_error Then
                fueltax_errors error_found
                Exit Sub
            End If
    
            arry_cnt = trn_count()
            For ti = 1 To arry_cnt
                pr_trn = get_a_trn(ti)
                it_is = False
                If pa_ftax_cnt > 0 Then
                For tj = 1 To pa_ftax_cnt
                   If pa_ftax(tj).ftax_prodlnk = aryComp(j).lProductLink And pa_ftax(tj).ftax_trn = pr_trn.trn And pa_ftax(tj).ftax_ulink = pr_trn.ulink Then
                         pa_ftax(tj).ftax_amount = pa_ftax(tj).ftax_amount + pr_trn.extn
                         pa_ftax(tj).ftax_tunits = pa_ftax(tj).ftax_tunits + pr_trn.tunits
                         it_is = True
                         Exit For
                   End If
                Next
                End If
                If it_is <> True Then
                   pa_ftax_cnt = pa_ftax_cnt + 1
                   tk = pa_ftax_cnt
                   pa_ftax(tk).ftax_prodlnk = aryComp(j).lProductLink
                   pa_ftax(tk).ftax_trn = pr_trn.trn
                   pa_ftax(tk).ftax_ulink = pr_trn.ulink
                   pa_ftax(tk).ftax_amount = pr_trn.extn
                   pa_ftax(tk).ftax_tunits = pr_trn.tunits
                   pa_ftax(tk).ftax_tax_rate = pr_trn.amount
                End If
            Next
    
    
            arry_cnt = fltr_count()
            For ti = 1 To arry_cnt
                pr_trn = get_a_fltr(ti)
                it_is = False
                If pa_ftax_cnt > 0 Then
                    For tj = 1 To pa_ftax_cnt
                       If pa_ftax(tj).ftax_prodlnk = aryComp(j).lProductLink And _
                          pa_ftax(tj).ftax_trn = pr_trn.trn And _
                          pa_ftax(tj).ftax_ulink = pr_trn.ulink Then
                             pa_ftax(tj).ftax_amount = pa_ftax(tj).ftax_amount + pr_trn.extn
                             pa_ftax(tj).ftax_tunits = pa_ftax(tj).ftax_tunits + pr_trn.tunits
                             it_is = True
                             Exit For
                       End If
                    Next
                End If
                If it_is <> True Then
                   pa_ftax_cnt = pa_ftax_cnt + 1
                   tk = pa_ftax_cnt
                   pa_ftax(tk).ftax_prodlnk = aryComp(j).lProductLink
                   pa_ftax(tk).ftax_trn = pr_trn.trn
                   pa_ftax(tk).ftax_ulink = pr_trn.ulink
                   pa_ftax(tk).ftax_amount = pr_trn.extn
                   pa_ftax(tk).ftax_tunits = pr_trn.tunits
                   pa_ftax(tk).ftax_tax_rate = pr_trn.amount
                End If
            Next
    
            'david 11/05/2001
            #If Not ZZEMMASU Then
                'david 06/15/2001
                subShowStatusMsg "Preparing G/L (2) - " & lProfitCenter & ", " & _
                    ivh_product & ", " & aryComp(j).sUseGroup
            #End If
            ''''''''''''''''''''
                
            extrn_del_gl lProfitCenter, ivh_product, _
                             aryComp(j).sUseGroup, aryComp(j).dfCompUnits, _
                             aryComp(j).dfCompUnits, aryComp(j).dfCompUnits, fuel_only, _
                             0, lProfitCenter, fuel_only
            
            If found_error Then
                fueltax_errors error_found
                Exit Sub
            End If
    
            arry_cnt = trn_count()
            For ti = 1 To arry_cnt
                pr_trn = get_a_trn(ti)
                it_is = False
                If pa_ftax_cnt > 0 Then
                For tj = 1 To pa_ftax_cnt
                   If pa_ftax(tj).ftax_prodlnk = aryComp(j).lProductLink And pa_ftax(tj).ftax_trn = pr_trn.trn And pa_ftax(tj).ftax_ulink = pr_trn.ulink Then
                         pa_ftax(tj).ftax_amount = pa_ftax(tj).ftax_amount + pr_trn.extn
                         pa_ftax(tj).ftax_tunits = pa_ftax(tj).ftax_tunits + pr_trn.tunits
                         it_is = True
                         Exit For
                   End If
                Next
                End If
                If it_is <> True Then
                   pa_ftax_cnt = pa_ftax_cnt + 1
                   tk = pa_ftax_cnt
                   pa_ftax(tk).ftax_prodlnk = aryComp(j).lProductLink
                   pa_ftax(tk).ftax_trn = pr_trn.trn
                   pa_ftax(tk).ftax_ulink = pr_trn.ulink
                   pa_ftax(tk).ftax_amount = pr_trn.extn
                   pa_ftax(tk).ftax_tunits = pr_trn.tunits
                   pa_ftax(tk).ftax_tax_rate = pr_trn.amount
                End If
            Next
    
            arry_cnt = fltr_count()
            For ti = 1 To arry_cnt
                pr_trn = get_a_fltr(ti)
                it_is = False
                If pa_ftax_cnt > 0 Then
                For tj = 1 To pa_ftax_cnt
                   If pa_ftax(tj).ftax_prodlnk = aryComp(j).lProductLink And pa_ftax(tj).ftax_trn = pr_trn.trn And pa_ftax(tj).ftax_ulink = pr_trn.ulink Then
                         pa_ftax(tj).ftax_amount = pa_ftax(tj).ftax_amount + pr_trn.extn
                         pa_ftax(tj).ftax_tunits = pa_ftax(tj).ftax_tunits + pr_trn.tunits
                         it_is = True
                         Exit For
                   End If
                Next
                End If
                If it_is <> True Then
                   pa_ftax_cnt = pa_ftax_cnt + 1
                   tk = pa_ftax_cnt
                   pa_ftax(tk).ftax_prodlnk = aryComp(j).lProductLink
                   pa_ftax(tk).ftax_trn = pr_trn.trn
                   pa_ftax(tk).ftax_ulink = pr_trn.ulink
                   pa_ftax(tk).ftax_amount = pr_trn.extn
                   pa_ftax(tk).ftax_tunits = pr_trn.tunits
                   pa_ftax(tk).ftax_tax_rate = pr_trn.amount
                End If
            Next
        End If 'aryComp(j).dfCompDollars > 0
        
        #If PROCESSING Then
            nCurrIndex = j
        #End If
    Next
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Debug.Print "Collecting Summary"
    Dim gl_tmp As gl_type, meter_gl_amt As Double
    
    'david 11/05/2001
    #If Not ZZEMMASU Then
        'david 06/15/2001
        subShowStatusMsg "Collecting Summary"
    #End If
    ''''''''''''''''''''
    
    j = gl_count()
    
    meter_acct = 0
    meter_gl_amt = -1#
    
    meter_gl = 0#
    
    For i = 1 To j
'        get_a_gl i, pc, acct, dc_flag, amt
        gl_tmp = get_a_gl(i)
        With gl_tmp
            If .dc_flag = "C" Then
                meter_gl = meter_gl + .gl_amount
                If .gl_amount >= 0 Then
                    If .gl_amount > meter_gl_amt Then
                        meter_acct = .account
                        meter_gl_amt = .gl_amount
                    End If
                End If
                
                .gl_amount = .gl_amount * -1
            End If
            If .account <> pr_ar_class!arc_glm Then
                If .dc_flag <> "C" Then
                    meter_gl = meter_gl - .gl_amount
                End If
                
                For k = 1 To 500
                    If pa_gldy(k).gldy_acct = 0 Then
                        pa_gldy(k).gldy_acct = .account
                        pa_gldy(k).gldy_debit = .gl_amount
                        pa_gldy_cnt = k
                        Exit For
                    End If
                Next
            End If
        End With
    Next
    If meter_total <> meter_gl Then
        For k = 1 To 500
            If pa_gldy(k).gldy_acct = meter_acct Then
                pa_gldy(k).gldy_debit = pa_gldy(k).gldy_debit + (meter_gl - meter_total)
                Exit For
            End If
        Next
    End If
    
    subListGL
    
    Exit Sub

errTrap:
    tfnErrHandler "fuel_gl"
End Sub

Public Sub subListGL()
    Dim k As Integer
    
    For k = 1 To 500
        With pa_gldy(k)
            If .gldy_acct > 0 Then
                Debug.Print .gldy_acct; " "; .gldy_debit
            End If
    End With
    Next
End Sub

Public Function compute_fuel_taxes(w As Integer) As Double
    
    Dim w1 As Integer, dtx As dtx_type
    Dim lcnt As Integer, ret_val As Double

    On Error GoTo errTrap
    
    Set pr_inv_header = fproduct(aryComp(w).lProductLink, t_dbMainDatabase)

    set_this_pc lProfitCenter
'pass
    detax aryComp(w).sUseGroup, aryComp(w).sTaxClass, _
        aryComp(w).dfCompDollars, aryComp(w).dfCompUnits

    
    If found_error Then Exit Function
    
    ret_val = 0
    lcnt = dtx_count()
    For w1 = 1 To lcnt
        dtx = get_a_dtx(w1)
        ret_val = ret_val + dtx.amount
    Next

    compute_fuel_taxes = ret_val

    Exit Function
    
errTrap:
    #If PROCESSING Then
        tfnErrHandler "compute_fuel_taxes", False
    #Else
        tfnErrHandler "compute_fuel_taxes"
    #End If
End Function

