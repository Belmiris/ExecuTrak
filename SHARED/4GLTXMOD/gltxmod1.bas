Attribute VB_Name = "gltxmod1"
'{ @('#) gltxmod1.4gl 1.7@('#) Last Delta:10/22/93  Tax & GL & Freight Routines }

'{   Copyright (c) 1988, FACTOR, A Division of WR Hess Company   }
'Database Factor

'{     This module contains the tax and product gl routines

'    An error condition causes a return of true and the error number
'    A normal completion causes a return of false and a zero, in most
'    cases. The others are returning actual information.


'}


''# ********************************************************************
''# NOTE: if you want to see the detailed journal entries being made
''#       while skipping all of the complex intervening steps
''#       just set a break point at the beginning of the following
''#       routines: debit_line, credit_line, debit_tax, credit_tax.
''#       There is already a dummy line to break on in these routines.
''#       Below is a sample 4db file to read in to the dubugger to
''#       get a perfect output for analyzing the journal entries.
''#       The line numbers have to be adjusted for current lines, and
''#       this break configuration will write results to a file.
''#       Search for 'debugging' to find the lines to break on.
''#
''#    Sample 4db file to read into the debugger to set break points.
''#    --------------------------------------------------------------
''# break gltxmod1.284 {pr 'debit line'>>line_tax.out}
''# break gltxmod1.284 {pr gll_acct>>line_tax.out}
''# break gltxmod1.284 {pr gll_amount>>line_tax.out}
''# break gltxmod1.318 {pr 'credit line'>>line_tax.out}
''# break gltxmod1.318 {pr gll_acct>>line_tax.out}
''# break gltxmod1.318 {pr gll_amount>>line_tax.out}
''# break gltxmod1.353 {pr 'debit tax'>>line_tax.out}
''# break gltxmod1.353 {pr txs_acct>>line_tax.out}
''# break gltxmod1.353 {pr txs_amount>>line_tax.out}
'# break gltxmod1.387 {pr 'credit tax'>>line_tax.out}
'# break gltxmod1.387 {pr txs_acct>>line_tax.out}
'# break gltxmod1.387 {pr txs_amount>>line_tax.out}
'#
'# ********************************************************************

Option Explicit

Global i As Integer, j As Integer, k As Integer, L As Integer, m As Integer, n As Integer
Global post_amt As Double
Global gl_num As Long
Global post_also_flag As String * 1

Global pr_gl_master As Recordset  '            record like gl_master.*
Global pr_tmp_master As Recordset  '             record like gl_master.*
Global this_pc As Long

Global tul_max As Integer
Global ulink_cnt As Integer
Global tt_max As Integer
Global tt_cnt As Integer

Global gl_max As Integer
Global gl_arry_cnt As Integer
Type gl_type
    prft_ctr As Long
    account As Long
    dc_flag As String * 1
    gl_amount As Double
End Type
Global gl_arry() As gl_type

Global txs_max As Integer
Global txs_arry_cnt As Integer
Global txs_arry() As gl_type

Global gll_max As Integer
Global gll_arry_cnt As Integer
Global gll_arry() As gl_type

Public Sub initialize_gl()
    tul_max = 800
    tt_max = 100
    gl_max = 100
    txs_max = 100
    gll_max = 100
    dtx_max = 100

    gl_arry_cnt = 0
    ReDim gl_arry(gl_max)

    txs_arry_cnt = 0
    ReDim txs_arry(txs_max)
    
    ReDim pa_gldy(500)
End Sub

Public Sub clear_gl_lines()


    gll_arry_cnt = 0
    ReDim gll_arry(gll_max)

    post_also "G"

End Sub

Public Function post_also(c As String)
    If c = "" Then
        post_also = post_also_flag
    Else
        post_also_flag = c
    End If
End Function


Public Function get_a_gl(which_one As Integer) As gl_type  'pc, acct, dc_flag, amt
    get_a_gl = gl_arry(which_one)
End Function

Public Function gl_count()
    gl_count = gl_arry_cnt
End Function

Public Function get_a_gll(which_one As Integer) As gl_type
    get_a_gll = gll_arry(which_one)
End Function

Public Function gll_count()
    gll_count = gll_arry_cnt
End Function

Public Function get_a_txs(which_one As Integer) As gl_type
    get_a_txs = txs_arry(which_one)
End Function

Public Function txs_count()
    txs_count = txs_arry_cnt
End Function

Public Sub post_debit(gl_acct As Long, gl_amount As Double)
    Dim i As Integer

    If gl_amount = 0# Then
        Exit Sub
    End If

Debug.Print "post_debit-" & gl_acct & ", " & gl_amount

    this_pc = gthis_pc(-1)
    Set pr_tmp_master = fgl_master(sSeries, gl_acct, this_pc, t_dbMainDatabase)
    If pr_tmp_master Is Nothing Then
        #If PROCESSING Then
            #If DEVELOP Then
                MsgBox "GL error in post_debit Acct:" & gl_acct & _
                    ", Prft Ctr:" & this_pc, vbCritical
            #End If
            subPostError "post_debit", "GL error in post_debit Acct:" & gl_acct & _
                ", Prft Ctr:" & this_pc, nCurrIndex
        #End If
'''''added 07-10-1997
        pr_gl_master_glm_account = gl_acct
        pr_gl_master_glm_prft_ctr = this_pc
        found_error = True: error_found = 10
        Exit Sub
    End If

    If post_also("") = "T" Then
        debit_tax gl_acct, gl_amount
    Else
        debit_line gl_acct, gl_amount
    End If

    For i = 1 To gl_arry_cnt
        If (gl_arry(i).prft_ctr = gthis_pc(-1)) And (gl_arry(i).account = gl_acct) Then
            If gl_arry(i).dc_flag = "D" Then
                gl_arry(i).gl_amount = gl_arry(i).gl_amount + gl_amount
            Else
                gl_arry(i).gl_amount = gl_arry(i).gl_amount - gl_amount
                If gl_arry(i).gl_amount < 0 Then
                    gl_arry(i).dc_flag = "D"
                    gl_arry(i).gl_amount = 0 - gl_arry(i).gl_amount
                End If
            End If
            Exit Sub
        End If
    Next
    
    gl_arry_cnt = gl_arry_cnt + 1
    i = gl_arry_cnt
    gl_arry(i).prft_ctr = gthis_pc(-1)
    gl_arry(i).account = gl_acct
    gl_arry(i).dc_flag = "D"
    gl_arry(i).gl_amount = gl_amount
End Sub

Public Sub post_credit(gl_acct As Long, gl_amount As Double)
    Dim i As Integer
    If gl_amount = 0# Then
        Exit Sub
    End If

Debug.Print "post_credit-" & gl_acct & ", " & gl_amount

    this_pc = gthis_pc(-1)
    Set pr_tmp_master = fgl_master(sSeries, gl_acct, this_pc, t_dbMainDatabase)
    If pr_tmp_master Is Nothing Then
        #If PROCESSING Then
            #If DEVELOP Then
                MsgBox "GL error in post_credit Acct:" & gl_acct & _
                    ",Prft Ctr:" & this_pc, vbCritical
            #End If
            subPostError "post_credit", "GL error in post_credit Acct:" & gl_acct & _
                ",Prft Ctr:" & this_pc, nCurrIndex
        #End If
        pr_gl_master_glm_account = gl_acct
        pr_gl_master_glm_prft_ctr = this_pc
        found_error = True: error_found = 10
        Exit Sub
    End If

    If post_also("") = "T" Then
        credit_tax gl_acct, gl_amount
    Else
        credit_line gl_acct, gl_amount
    End If

    For i = 1 To gl_arry_cnt
        If (gl_arry(i).prft_ctr = gthis_pc(-1)) And (gl_arry(i).account = gl_acct) Then
            If gl_arry(i).dc_flag = "C" Then
                gl_arry(i).gl_amount = gl_arry(i).gl_amount + gl_amount
            Else
                gl_arry(i).gl_amount = gl_arry(i).gl_amount - gl_amount
                If gl_arry(i).gl_amount < 0 Then
                    gl_arry(i).dc_flag = "C"
                    gl_arry(i).gl_amount = 0 - gl_arry(i).gl_amount
                End If
            End If
            Exit Sub
        End If
    Next
    
    gl_arry_cnt = gl_arry_cnt + 1
    i = gl_arry_cnt
    gl_arry(i).prft_ctr = gthis_pc(-1)
    gl_arry(i).account = gl_acct
    gl_arry(i).dc_flag = "C"
    gl_arry(i).gl_amount = gl_amount
End Sub


Public Sub debit_line(gll_acct As Long, gll_amount As Double)
    Dim i As Integer

    i = 1 '# this line is here for debugging purposes only, put a
              '# break here to see product(non-tax) debits
    For i = 1 To gll_arry_cnt
        If (gll_arry(i).prft_ctr = gthis_pc(-1)) And (gll_arry(i).account = gll_acct) Then
            If gll_arry(i).dc_flag = "D" Then
                gll_arry(i).gl_amount = gll_arry(i).gl_amount + gll_amount
            Else
                gll_arry(i).gl_amount = gll_arry(i).gl_amount - gll_amount
                If gll_arry(i).gl_amount < 0 Then
                    gll_arry(i).dc_flag = "D"
                    gll_arry(i).gl_amount = 0 - gll_arry(i).gl_amount
                End If
            End If
            Exit Sub
        End If
    Next
    
    gll_arry_cnt = gll_arry_cnt + 1
    i = gll_arry_cnt
    gll_arry(i).prft_ctr = gthis_pc(-1)
    gll_arry(i).account = gll_acct
    gll_arry(i).dc_flag = "D"
    gll_arry(i).gl_amount = gll_amount
End Sub

Public Sub credit_line(gll_acct As Long, gll_amount As Double)
    Dim i As Integer

    i = 1 '# this line is here for debugging purposes only, put a
              '# break here to see product(non-tax) credits
    For i = 1 To gll_arry_cnt
        If (gll_arry(i).prft_ctr = gthis_pc(-1)) And (gll_arry(i).account = gll_acct) Then
            If gll_arry(i).dc_flag = "C" Then
                gll_arry(i).gl_amount = gll_arry(i).gl_amount + gll_amount
            Else
                gll_arry(i).gl_amount = gll_arry(i).gl_amount - gll_amount
                If gll_arry(i).gl_amount < 0 Then
                    gll_arry(i).dc_flag = "C"
                    gll_arry(i).gl_amount = 0 - gll_arry(i).gl_amount
                End If
            End If
            Exit Sub
        End If
    Next
    
    gll_arry_cnt = gll_arry_cnt + 1
    i = gll_arry_cnt
    gll_arry(i).prft_ctr = gthis_pc(-1)
    gll_arry(i).account = gll_acct
    gll_arry(i).dc_flag = "C"
    gll_arry(i).gl_amount = gll_amount
End Sub


Public Sub debit_tax(txs_acct As Long, txs_amount As Double)
    Dim i As Integer

    i = 1 '# this line is here for debugging purposes only, put a
              '# break here to see tax debits
    For i = 1 To txs_arry_cnt
        If (txs_arry(i).prft_ctr = gthis_pc(-1)) And (txs_arry(i).account = txs_acct) Then
            If txs_arry(i).dc_flag = "D" Then
                txs_arry(i).gl_amount = txs_arry(i).gl_amount + txs_amount
            Else
                txs_arry(i).gl_amount = txs_arry(i).gl_amount - txs_amount
                If txs_arry(i).gl_amount < 0 Then
                    txs_arry(i).dc_flag = "D"
                    txs_arry(i).gl_amount = 0 - txs_arry(i).gl_amount
                End If
            End If
            Exit Sub
        End If
    Next
    
    txs_arry_cnt = txs_arry_cnt + 1
    i = txs_arry_cnt
    txs_arry(i).prft_ctr = gthis_pc(-1)
    txs_arry(i).account = txs_acct
    txs_arry(i).dc_flag = "D"
    txs_arry(i).gl_amount = txs_amount
End Sub

Public Sub credit_tax(txs_acct As Long, txs_amount As Double)
    Dim i As Integer

    i = 1 '# this line is here for debugging purposes only, put a
              '# break here to see tax credits
    For i = 1 To txs_arry_cnt
        If (txs_arry(i).prft_ctr = gthis_pc(-1)) And (txs_arry(i).account = txs_acct) Then
            If txs_arry(i).dc_flag = "C" Then
                txs_arry(i).gl_amount = txs_arry(i).gl_amount + txs_amount
            Else
                txs_arry(i).gl_amount = txs_arry(i).gl_amount - txs_amount
                If txs_arry(i).gl_amount < 0 Then
                    txs_arry(i).dc_flag = "C"
                    txs_arry(i).gl_amount = 0 - txs_arry(i).gl_amount
                End If
            End If
            Exit Sub
        End If
    Next
    
    txs_arry_cnt = txs_arry_cnt + 1
    i = txs_arry_cnt
    txs_arry(i).prft_ctr = gthis_pc(-1)
    txs_arry(i).account = txs_acct
    txs_arry(i).dc_flag = "C"
    txs_arry(i).gl_amount = txs_amount
End Sub



