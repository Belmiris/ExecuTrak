Attribute VB_Name = "modDiscount"
Option Explicit
'/////////////////////////////////////////////////////////////////////////////
'//  modDiscount.bas        :      Implementation file.
'//
'//  Initial programmer      :       Shawn Scotzin
'//
'//  Date completed          :       5/13/96
'//
'//  Purpose                      :       Compute's invoice and discount due dates as well as discount amounts
'//
'//  Functions                   :       fnComputeDiscount
'//
'//  Inputs                         :      Invoice Date, Invoice Dollar Amount, Invoice Quantity (Gross ?),
'//                                            Terms Code
'//
'//
'/////////////////////////////////////////////////////////////////////////////



' This is what the fnComputeDiscount will return

Type udtDiscount
    dInvoiceDate                    As Date      '// The date that the invoice in due
    dDiscountDueDate           As Date      '// The date of the discount
    fDiscountAmount              As Double  '// The amount of the discount
End Type


Public bDontShowError1 As Boolean

Public Function fnComputeDiscount(dInvDate As Date, _
                                  fDollarAmount As Double, _
                                  fInvQuantity As Double, _
                                  strTerms As String) As udtDiscount

Dim rsTerms               As Recordset        '// The Recordset for the terms


Dim udtDisc                As udtDiscount     '//This will be the record that we return

' No dollars then get out
If fDollarAmount < 0 Then
    With udtDisc
        .dInvoiceDate = dInvDate
        .dDiscountDueDate = dInvDate
        .fDiscountAmount = 0
    End With
    
    fnComputeDiscount = udtDisc
    Exit Function
End If
    
Set rsTerms = fnGetTerms(strTerms)

With udtDisc
    .dInvoiceDate = dInvDate
    .dDiscountDueDate = dInvDate
    .fDiscountAmount = 0
End With

'If no terms then get out
If Not rsTerms Is Nothing Then
    If rsTerms.RecordCount > 0 Then
        With udtDisc
            .dInvoiceDate = fnComputeDueDate(dInvDate, rsTerms!term_due_operator, rsTerms!term_due_date_days)
            .dDiscountDueDate = fnComputeDueDate(dInvDate, rsTerms!term_disc_operator, rsTerms!term_disc_date_day)
            .fDiscountAmount = fnComputeDiscountAmount(rsTerms, fDollarAmount, fInvQuantity)
        End With
    End If
End If

fnComputeDiscount = udtDisc

End Function


'// Get the terms records for the terms code
Private Function fnGetTerms(strTerms As String) As Recordset
    Const SUB_NAME = "fnGetTerms"
    
    Dim strSQL As String
    Dim rsTerms As Recordset
    
    strSQL = "SELECT * FROM sys_terms WHERE term_code = " & tfnSQLString(strTerms)
    On Error GoTo GET_TERMS_ERROR
        Set rsTerms = t_dbMainDatabase.OpenRecordset(strSQL, dbOpenSnapshot, dbSQLPassThrough)
    On Error GoTo 0
    
    If rsTerms.RecordCount > 0 Then
        Set fnGetTerms = rsTerms
    Else
        Set fnGetTerms = Nothing
    End If
    
    Exit Function
GET_TERMS_ERROR:
    tfnErrHandler SUB_NAME, strSQL, Not bDontShowError1
End Function


'// Computes the discount due date
'david 04/28/2005  #479149
Public Function fnComputeDueDate(dInvDate As Date, strOperator As Variant, vDays As Variant) As Date

    Dim sMonth As String
    Dim sYear As String
    Dim nDays As Integer
    Dim dDueDate As Date
    
    dDueDate = dInvDate
    
'    strOperator = rsTerms!term_disc_operator
    
    If IsNull(strOperator) Or IsNull(nDays) Then
        fnComputeDueDate = dInvDate
        Exit Function
    End If
    
    nDays = Val(vDays)

    strOperator = Left(strOperator, 1)
    Select Case strOperator
        
        Case "+":
            dDueDate = DateAdd("d", nDays, dDueDate)
        Case "=":
            If nDays < 1 Then
                nDays = 1
            End If
            If nDays > 31 Then
                nDays = 31
            End If
            dDueDate = DateAdd("m", 1, dDueDate)
            sMonth = CStr(Month(dDueDate)) & "/"
            sYear = "/" & CStr(Year(dDueDate))
            On Error GoTo errTooManyDays
            dDueDate = CDate(sMonth & CStr(nDays) & sYear)
    End Select

    fnComputeDueDate = dDueDate

extDueDate:
    On Error GoTo 0
    Exit Function
    
errTooManyDays:
    If nDays > 28 Then
        If Err.Number = 13 Then
            nDays = nDays - 1
            Resume
        End If
    End If
    MsgBox "Error occurred while calculating due days." & vbOKOnly + vbExclamation
    Err.Clear
    Resume extDueDate
End Function

Private Function fnGetUnitsDiscount(rsTerms As Recordset, fInvQuantity As Double) As Recordset
    Const SUB_NAME = "fnGetUnitsDiscount"
    
    Dim strSQL As String
    Dim rsSysDiscUnits   As Recordset        '// Recordset for unit based discounts
    
    strSQL = "SELECT * FROM sys_disc_units WHERE sdu_term_code = " & tfnSQLString(rsTerms!term_code) & _
                    " AND " & CStr(fInvQuantity) & " <=  sdu_max_units ORDER BY sdu_max_units"
                    
    On Error GoTo GET_UNITS_DISC_ERROR
    
        Set rsSysDiscUnits = t_dbMainDatabase.OpenRecordset(strSQL, dbOpenSnapshot, dbSQLPassThrough)
    On Error GoTo 0
    
    If rsSysDiscUnits.RecordCount > 0 Then
        Set fnGetUnitsDiscount = rsSysDiscUnits
    Else
        Set fnGetUnitsDiscount = Nothing
    End If
    
    Exit Function

GET_UNITS_DISC_ERROR:
    tfnErrHandler SUB_NAME, strSQL, Not bDontShowError1
End Function

'// Compute the discount amount based on the terms and the
'// dollar amount or the unit quantity
Private Function fnComputeDiscountAmount(rsTerms As Recordset, _
                                         fDollarAmount As Double, _
                                         fInvQuantity As Double) As Double

Dim rsSysDiscUnits   As Recordset        '// Recordset for unit based discounts
Dim dDiscount  As Double

dDiscount = fDollarAmount * (tfnRound(rsTerms!term_disc_percent, DEFAULT_DECIMALS) / 100)

If fInvQuantity <> 0 Then
    Set rsSysDiscUnits = fnGetUnitsDiscount(rsTerms, Abs(fInvQuantity))
Else
    fnComputeDiscountAmount = dDiscount
    Exit Function
End If

'If rsSysDiscUnits Is Nothing Then
'    fnComputeDiscountAmount = dDiscount
'    Exit Function
'Else
If Not rsSysDiscUnits Is Nothing Then
    '// If there is a unit based discount for these terms then it will overide the
    '// the default discount for these terms
    rsSysDiscUnits.MoveFirst
    dDiscount = rsSysDiscUnits!sdu_cents_unit * fInvQuantity
End If
    
fnComputeDiscountAmount = dDiscount

End Function



