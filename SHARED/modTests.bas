Attribute VB_Name = "modTests"
' This module was created to identify scenarios which need testing
'
' To make these Test functions available:
'    Under Project Menu > Properties > Make Tab
'        Add/Set the "Conditional Compilation Arguments" for ALLOW_TESTS = 1
'
'    PLEASE return ALLOW_TESTS = 0 before release

#If ALLOW_TESTS Then ' these functions should not be compiled into the final product


    'Print G9126Test1(1.0825) ' returns additional garbage
    ' 1.08249998092651
    '
    'Print G9126Test1(1.0825, 1)
    ' 1.0825
    '
    'Print G9126Test1(1.0825, 2)
    ' 1.0825
    '
    'Print G9126Test1(0.00000005) ' returns scientific/exponential notation
    ' 5.00000005843049E-08
    '
    'Print G9126Test1(0.00000005, 1)
    ' 0.00000005
    '
    'Print G9126Test1(0.00000005, 2)
    ' 0.00000005
    '
    'analyze single to double issue
    Public Function G9126Test1(sng As Single, Optional Scenario As Integer = 0) As Double
        
        Select Case Scenario
            Case 1 ' Cstr()
                G9126Test1 = CDbl(CStr(sng))
            Case 2 ' tfnRound()
                G9126Test1 = CDbl(tfnRound(sng, 8))
            Case Else ' raw
                G9126Test1 = CDbl(sng)
        End Select
    End Function
    
    
    
    '
    '
    'analyze single to double issues:
    '
    '    Scenario 1
    '        If get_tax_free_price(aryComp(i).dfPrice, _
                              curr_date, _
                              prft_ctr, _
                              aryComp(i).sUseGroup, _
                              aryComp(i).sTaxClass, _
                              aryComp(i).dfTaxFreePrice) = False Then
    '      Print G8967Test1(1)
    '
    '    Scenario 2
    '        dPrice = zzf_Calc_Price(.lProductLink, .dfPrice, dRentRate, , "DO NOT INCLUDE Rent")
    '            .dfPrice = tfnRound(dPrice + dRentRate, DEFAULT_DECIMALS)
    '      Print G8967Test1(2)
    '
    Public Function G8967Test1(Scenario As Integer)
        Dim sInput As Single
        Dim d As Double
        Dim dZero As Double
        Dim v As Variant
        Dim s As Single
        Dim dOutput1 As Double
        Dim dOutput2 As Double
        Dim sOutput1 As Single
        Dim sOutput2 As Single
        Dim lPC As Integer
        Dim lGP As Integer
        Dim strDate As String
        
        Select Case Scenario
            Case 1 ' get_tax_free_price()
                Debug.Print "G8967Test1(1) will first pass the price (0 to 10 step .01) as a single into get_tax_free_price()"
                Debug.Print " The second step will pass the price as a single into tfnRound() first before passing into get_tax_free_price()"
                Debug.Print " A warning is displayed if any difference is found..."
                Debug.Print ""
                For d = 0 To 10 Step 0.01 ' looping a single caused the single to go scientific after .05
                    
                    s = d
                    get_tax_free_price s, _
                          Date, _
                          1, _
                          "OKSUPP", _
                          "GAS", _
                          dOutput1
                              
                    get_tax_free_price tfnRound(s, 8), _
                          Date, _
                          1, _
                          "OKSUPP", _
                          "GAS", _
                          dOutput2
                          
                    If dOutput1 <> dOutput2 Then
                        Debug.Print "    WARNING: difference found when " & s & " is passed to get_tax_free_price() : " & dOutput1 & " vs " & dOutput2
                    End If
                Next
                Debug.Print "G8967Test1(1) COMPLETE"
            Case 2 ' dPrice = zzf_Calc_Price(.lProductLink, .dfPrice, dRentRate, , "DO NOT INCLUDE Rent")
                    '.dfPrice = tfnRound(dPrice + dRentRate, DEFAULT_DECIMALS)
                Debug.Print "G8967Test1(2) will first pass the price (0 to 10 step .01) as a single into zzf_Calc_Price()"
                Debug.Print " The second step will pass the price as a single into tfnRound() first before passing into zzf_Calc_Price()"
                Debug.Print " A warning is displayed if any difference is found..."
                Debug.Print ""
                
                lPC = gn_ZZ_Current_Prft_Ctr
                lGP = gn_ZZ_Current_Group
                strDate = frmRSENTRY!txtReportDate
                gn_ZZ_Current_Prft_Ctr = 123
                gn_ZZ_Current_Group = 1
                frmRSENTRY!txtReportDate = Date
                
                For s = 0 To 10 Step 0.01
                    
                    dOutput1 = zzf_Calc_Price(1, s, , , "DO NOT INCLUDE Rent")
                    sOutput1 = tfnRound(dOutput1 + dZero, 8)
                              
                    dOutput2 = zzf_Calc_Price(1, tfnRound(s, 8), , , "DO NOT INCLUDE Rent")
                    sOutput2 = tfnRound(dOutput2 + dZero, 8)
                          
                    If dOutput1 <> dOutput2 Then
                        Debug.Print "    WARNING: difference found when " & s & " is passed to zzf_Calc_Price() : " & dOutput1 & " vs " & dOutput2
                    End If
                    If sOutput1 <> sOutput2 Then
                        Debug.Print "    WARNING: difference found when " & s & " is passed to zzf_Calc_Price() : " & sOutput1 & " vs " & sOutput2
                    End If
                Next
                
                gn_ZZ_Current_Prft_Ctr = lPC
                gn_ZZ_Current_Group = lGP
                frmRSENTRY!txtReportDate = strDate
                
                Debug.Print "G8967Test1(2) COMPLETE"
        End Select
        
    End Function
    
#End If
