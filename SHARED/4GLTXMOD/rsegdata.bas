Attribute VB_Name = "rsegdata"
'{   @(#) rsegdata.4gl 1.22@(#)  Last Delta:1/16/96   }

'{ Copyright (c) 1988,89,90,91 FACTOR, A Division of WR Hess Company   }

'{ rsegdata }

Global it_is As Integer

Global pr_inv_master As Recordset  '            record like inv_master.*
Global pr_inv_header  As Recordset  '            record like inv_header.*

Global pa_tax As Recordset  '(10) of      record like tx_table.*

Type pa_ftax_type
    ftax_prodlnk As Long
    ftax_trn As Long
    ftax_ulink As Long
    ftax_amount As Double
    ftax_tunits As Double
    ftax_tax_rate As Double
End Type

Global pa_ftax() As pa_ftax_type
Global pa_ftax_cnt As Integer

Type pa_gldy_type
    gldy_vendor As Long
    gldy_invoice As Long
    gldy_acct As Long
    gldy_debit As Double
End Type

Global pa_gldy() As pa_gldy_type
Global pa_gldy_cnt As Integer
Global pa_gldywrk As pa_gldy_type

Global found_error As Integer
Global error_found As Integer
