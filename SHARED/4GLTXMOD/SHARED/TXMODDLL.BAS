Attribute VB_Name = "modTXMODDLL"
Option Explicit
#Const DEBUGGING = False
#Const HARD_CODED_TX_LD_FEE = False
'# 2/26/04 Robert Atwood
'Added compile-time flag NO_DST to turn off or on Date Sensitive Taxes.  If it
'exists, will compile without date sensitivity, otherwise compiles with.

Private Const ERROR_LOADING1 = 1201
Private Const ERROR_LOADING2 = 1202
Private Const MOD_DLL_NAME = "TxMod32.dll"

#If DEBUGGING Then
    Declare Function TXMOD_Attach Lib "TxMod32.dll" _
        (ByVal hInst As Long, ByVal hwnd As Long, _
        ByVal szConnect As String) As Long
    
    Declare Function TXMOD_Detach Lib "TxMod32.dll" _
        (ByVal hInst As Long, ByVal hwnd As Long) As Long
        
    Declare Function TXMOD_VB_CalculateLineTax Lib "TxMod32.dll" _
        (ByVal hInst As Long, ByVal hwnd As Long, _
        ByRef bpLineItem As CMovementLineItem, _
        ByRef nRecords As Long) As Long
    
    Declare Function TXMOD_GetLineTaxes Lib "TxMod32.dll" _
        (ByVal hInst As Long, ByVal hwnd As Long, _
        ByRef bpTaxRecord() As CApplicableTaxes) As Long
    
    Declare Function TXMOD_GetLineTax Lib "TxMod32.dll" _
        (ByVal hInst As Long, ByVal hwnd As Long, ByVal nIdx As Long, _
        ByRef bpTaxRecord As CApplicableTaxes) As Long
    
    Declare Function TXMOD_DetermineAccounts Lib "TxMod32.dll" _
        (ByVal hInst As Long, ByVal hwnd As Long, _
        ByRef bpGlItem As CGLTaxEffectItem, ByRef lDebitAccount As Long, _
        ByRef lCreditAccount As Long) As Long

    'david 08/29/2003  #393678-4
    Declare Function TXMOD_gl_sd_gals Lib "TxMod32.dll" _
        (ByVal hInst As Long, ByVal hwnd As Long, ByVal nVessel As Integer, _
        ByVal dGallons As Double, ByVal dMiles As Double, ByVal dTotGals As Double) As Long
    ''''''''''''''''''''''''''''
    'Vijaya 02/10/2004  #395302
    Declare Function TXMOD_tx_date Lib "TxMod32.dll" _
        (ByVal hInst As Long, ByVal hwnd As Long, ByVal sDate As String) As Long
    ''''''''''''''''''''''''''''
    '#411523wj042505 -- Set it after TXMOD_tx_date -- Used in Limit Type "I" tax calculations
    Declare Function TXMOD_set_accm_txble_amt Lib "TxMod32.dll" _
        (ByVal hInst As Long, ByVal hwnd As Long, ByVal dAccmTxbleAmt As Double) As Long
    
#Else
    Declare Function TXMOD_Attach Lib "TxMod32.dll" _
        (ByVal hInst As Long, ByVal hwnd As Long, _
        ByVal szConnect As String) As Long
    
    Declare Function TXMOD_Detach Lib "TxMod32.dll" _
        (ByVal hInst As Long, ByVal hwnd As Long) As Long
        
    Declare Function TXMOD_VB_CalculateLineTax Lib "TxMod32.dll" _
        (ByVal hInst As Long, ByVal hwnd As Long, _
        ByRef bpLineItem As CMovementLineItem, _
        ByRef nRecords As Long) As Long
    
    Declare Function TXMOD_GetLineTaxes Lib "TxMod32.dll" _
        (ByVal hInst As Long, ByVal hwnd As Long, _
        ByRef bpTaxRecord() As CApplicableTaxes) As Long
    
    Declare Function TXMOD_GetLineTax Lib "TxMod32.dll" _
        (ByVal hInst As Long, ByVal hwnd As Long, ByVal nIdx As Long, _
        ByRef bpTaxRecord As CApplicableTaxes) As Long
    
    Declare Function TXMOD_DetermineAccounts Lib "TxMod32.dll" _
        (ByVal hInst As Long, ByVal hwnd As Long, _
        ByRef bpGlItem As CGLTaxEffectItem, ByRef lDebitAccount As Long, _
        ByRef lCreditAccount As Long) As Long
    
    'david 08/29/2003  #393678-4
    Declare Function TXMOD_gl_sd_gals Lib "TxMod32.dll" _
        (ByVal hInst As Long, ByVal hwnd As Long, ByVal nVessel As Integer, _
        ByVal dGallons As Double, ByVal dMiles As Double, ByVal dTotGals As Double) As Long
    ''''''''''''''''''''''''''''
    'Vijaya 02/10/2004  #395302
    Declare Function TXMOD_tx_date Lib "TxMod32.dll" _
        (ByVal hInst As Long, ByVal hwnd As Long, ByVal sDate As String) As Long
    ''''''''''''''''''''''''''''
    '#411523wj042505 -- Set it after TXMOD_tx_date -- Used in Limit Type "I" tax calculations
    Declare Function TXMOD_set_accm_txble_amt Lib "TxMod32.dll" _
        (ByVal hInst As Long, ByVal hwnd As Long, ByVal dAccmTxbleAmt As Double) As Long
    
#End If

' Error codes
Public Const TXMOD_OK = &H0                 ' All's well
Public Const TXMOD_NOMEM = &H7FFF           ' Out of memory
Public Const TXMOD_BADPARMS = &H7FFE        ' Bad parameters passed
Public Const TXMOD_NODBASE = &H7FFD         ' No database available
Public Const TXMOD_NOAPPNODE = &H7FFC       ' Process not registered
Public Const TXMOD_OUTOFRANGE = &H7FFB      ' Index out of range
Public Const TXMOD_NOLINETAX = &H7FFA       ' Line tax hasn't been calculated
Public Const TXMOD_TASKBLOCKED = &H7FF9     ' DLL crtical section in use

Public Const TXMOD_LINETAX_OK = 0           ' All's well
Public Const TXMOD_LINETAX_NOMEM = 1        ' Out of memory
Public Const TXMOD_LINETAX_NODBASE = 2      ' No database available
Public Const TXMOD_LINETAX_NOTAXRECS = 3    ' No tax records in the dbase
Public Const TXMOD_LINETAX_NOINVHEAD = 4    ' No inventory header record
Public Const TXMOD_LINETAX_NOINVCLASS = 5   ' No inventory class record
Public Const TXMOD_LINETAX_BADSQL = 6       ' SQL statement error
Public Const TXMOD_LINETAX_BADBASIS = 7     ' Bad basis flag
Public Const TXMOD_LINETAX_BADBILLED = 8    ' Bad billed units flag
Public Const TXMOD_LINETAX_BADLTYPE = 9     ' Bad line type (int/ext bl/del)
Public Const TXMOD_LINETAX_BADPARMS = 10    ' Bad parameters passed

Public Const TXMOD_INTBOL = 0               ' Internal bill of lading
Public Const TXMOD_EXTBOL = 1               ' External bill of lading
Public Const TXMOD_INTDEL = 2               ' Internal delivery
Public Const TXMOD_EXTDEL = 3               ' External delivery

Public Const TXMOD_GL_OK = 0                ' All's well
Public Const TXMOD_GL_NOMEM = 1             ' Out of memory
Public Const TXMOD_GL_NODBASE = 2           ' No database available
Public Const TXMOD_GL_CANTOPENDBASE = 3     ' Can't open the database
Public Const TXMOD_GL_BADLTYPE = 4          ' Bad line type (int/ext bl/del intstate)
Public Const TXMOD_GL_NOINVMASTER = 5       ' No inventory master record
Public Const TXMOD_GL_FUELTAXERR = 6        ' Can't find "is a fuel" flag
Public Const TXMOD_GL_BADFUELFLAG = 7       ' Bad fuel flag
Public Const TXMOD_GL_BADCHARGEFLAG = 8     ' Bad charge flag
Public Const TXMOD_GL_NOINVHEAD = 9         ' No inventory header record
Public Const TXMOD_GL_NOINVCLASS = 10       ' No inventory class record

Public Const TXMOD_GL_INTBOL = 0            ' Internal bill of lading
Public Const TXMOD_GL_EXTBOL = 1            ' External bill of lading
Public Const TXMOD_GL_INTDEL = 2            ' Internal delivery
Public Const TXMOD_GL_EXTDEL = 3            ' External delivery
Public Const TXMOD_GL_ISTATE = 4            ' Inter/Intra state

Public Type CApplicableTaxes
    m_nOrder As Long
    m_lGlAcct As Long
    m_lLimit As Double
    m_lTrn As Long
    m_cLevel As Byte
    m_cCharge As Byte
    m_cBasis As Byte
    m_cFuelFlag As Byte
    m_fRate As Single
    m_dAmount As Double
    m_sCode As String * 5
    m_sWhere As String * 4
    m_sDate As String * 10
    m_sDescr As String * 30
    m_cLimitType As Byte
End Type

Public Type CMovementLineItem
    m_sProduct As String * 16
    m_sUseGroup As String * 8
    m_dDollarBase As Double
    m_dFreight As Double
    m_dNetUnits As Double
    m_dGrossUnits As Double
    m_tLineType As Long
    m_cBilledUnitsFlag As Byte
End Type

Public Type CGLTaxEffectItem
    m_sProduct As String * 16
    m_nProfitCenter As Integer
    m_lTaxGLAccount As Long
    m_cChargeFlag As Byte
    m_cFuelFlag As Byte
    m_tLineType As Integer
End Type

Private Const TAXMOD_LOADED = 1
Private Const TAXMOD_LOADING_FAILED = 2
Private Const TAXMOD_TRIES = 1000

Private Const TAX_BASIS_G = "G"
Private Const TAX_BASIS_N = "N"
Private Const TAX_BASIS_A = "A"
Private Const TAX_BASIS_B = "B"
Private Const TAX_BASIS_F = "F"
Private Const TAX_BASIS_D = "$"

Private Const ARRAY_SIZE_INC = 5
Private Type tpTaxSums
    m_lTrn As Long
    m_lLimit As Double
    m_cLimitType As Byte
    m_lGallons As Long
    m_cBasis As Byte
    m_fAmount As Double
    m_fRate As Double
End Type
Private arryTaxSums() As tpTaxSums
Private nTaxItems As Integer
Private nArraySize As Integer
Private nTaxModLoaded As Integer

Public Taxes() As CApplicableTaxes

'Ma, 2/26/99 Temporary patch for Taxes Loading Fee
Private sSysParm1055 As String
Private Type tpTaxGroup
    m_sProdCode As String
    m_sGroup As String
End Type
Private aryTaxGroup() As tpTaxGroup
Private nTaxGroupCount As Integer
Private Type tpTaxInfo
    m_lInvoice As Long
    m_nLine As Integer
    m_InputParms As CMovementLineItem
    m_Taxes() As CApplicableTaxes
    m_nRecord As Integer
End Type
Private aryTaxInfo() As tpTaxInfo
Private nInfoCount As Integer
Private fLadingTypeATax As Double
Private lInternalFuel As Long
Private m_Disable_A_Taxes As Boolean

''#452155wj101304 New LOading Fees Routine VVVVVVVVVVVV
Public Type BOL_AtAx
     fmld_bl_nbr As Long
     fmld_seq_nbr As Long
     sum_bunits As Long
     bl_ngu As String
     a_tax_tot As Double
     a_tax_used As Double
End Type
    
Public Type Loading_Fees
    dv_nbr As Long
    dv_line As Long
    dv_prodlnk As Long
    dv_prodtcl  As String
    dv_vol As Long
    dv_atax As Double
    dv_IE As String
End Type
Public aryLDF() As Loading_Fees
Public ldf_cnt As Integer '#this is the Ubound of the Array

Private Sub subDealWithATaxes(Taxes() As CApplicableTaxes, ByRef lRecords As Long)
    
    Dim i As Long
    Dim tmpTaxes() As CApplicableTaxes
    Dim lCount As Long
    Dim bFindTexasLoadingFee As Boolean
    
    'This function is added by WJ on 10/05/2001 per Magic #338055
    
    #If FACTOR_MENU Then
    On Error Resume Next
    #End If
    
    If Not m_Disable_A_Taxes Then
        'If not disable A Taxes, typically called from fuel movement
        'we do not need to do anything.
        Exit Sub
    End If
    
    'This routine will be used in all but fule movement programs
    'It will disable the taxes with Basis= "A" (loading fees)
    
    lCount = -1
    For i = 0 To lRecords - 1
        If Taxes(i).m_cBasis <> vbKeyA Then
            lCount = lCount + 1
            ReDim Preserve tmpTaxes(lCount) As CApplicableTaxes
            tmpTaxes(lCount) = Taxes(i)
        End If
    Next
    
    If lCount >= 0 Then
        lRecords = lCount + 1
        ReDim Taxes(lRecords - 1) As CApplicableTaxes
        For i = 0 To lRecords - 1
          Taxes(i) = tmpTaxes(i)
        Next
    End If
End Sub

'david 08/29/2003  #393678-4
'add two more optional parameters
Public Function fnCalculateTaxes(sProduct As String, _
                                  sTaxUseGroup As String, _
                                  dGross As Double, _
                                  dQuantity As Double, _
                                  ByRef dTaxAmount As Double, _
                                  Optional bShowMsgBox As Boolean = True, _
                                  Optional sErrMsg As String = "", _
                                  Optional bApplyLimit As Boolean = True, _
                                  Optional hld_vessel As Integer = -1, _
                                  Optional hld_gals As Double = 0#, _
                                  Optional hld_miles As Double = 0#, _
                                  Optional hld_tot_gals As Double = 0#, _
                                  Optional dAccmTxbleAmt As Double = 0#) As Boolean

    'Calculate taxes by calling the tax mods
    Dim myStruct As CMovementLineItem
    Dim nRet As Integer
    Dim nRecords As Long
    Dim dSum As Double
    Dim i As Integer
    Dim nTries As Integer
    
    'if the tax mods are not loaded, load them
    If Not fnLoadTaxMods Then
        Exit Function
    End If

    ' assign initial values to the input structure
    With myStruct
        .m_sProduct = sProduct
        .m_sUseGroup = sTaxUseGroup
        .m_dDollarBase = dGross
        .m_dFreight = 0
        .m_dNetUnits = dQuantity
        .m_dGrossUnits = dQuantity
        .m_cBilledUnitsFlag = Asc("N")
        .m_tLineType = TXMOD_EXTDEL
    End With
    
    'Determine the number of items in the record
    nTries = TAXMOD_TRIES
    'david 08/29/2003  #393678-4
    subSet_gl_sd_gals hld_vessel, hld_gals, hld_miles, hld_tot_gals
    ''''''''''''''''''''''''''''
    'Vijaya on 02/10/04 Magic#395302
    subSet_tx_date t_tax_date
    ''''''''''''''''''''''''''''''''
    
    '#411523wj042505
    If Abs(dAccmTxbleAmt) >= 0.01 Then
        subSet_AccmTxbleAmt dAccmTxbleAmt
    End If
    Do
        nRet = TXMOD_VB_CalculateLineTax(App.hInstance, Forms(0).hwnd, myStruct, nRecords)
        If i = TXMOD_TASKBLOCKED Then
            DoEvents
        End If
        nTries = nTries - 1
    Loop Until nRet <> TXMOD_TASKBLOCKED Or nTries <= 0
    
    If nRet <> TXMOD_OK Then
        subTaxModErrors nRet, bShowMsgBox, sErrMsg
        fnCalculateTaxes = False
        
        Exit Function
    End If

    ' Calculate the taxes
    ReDim Taxes(nRecords - 1)
    nTries = TAXMOD_TRIES
    Do
        nRet = TXMOD_GetLineTaxes(App.hInstance, Forms(0).hwnd, Taxes())
        If i = TXMOD_TASKBLOCKED Then
            DoEvents
        End If
        nTries = nTries - 1
    Loop Until nRet <> TXMOD_TASKBLOCKED Or nTries <= 0
    If nRet <> TXMOD_OK Then
        subTaxModErrors nRet, bShowMsgBox, sErrMsg
        fnCalculateTaxes = False
        Exit Function
    End If
    
    subDealWithATaxes Taxes, nRecords
    
    If bApplyLimit Then
        subApplyLimit Taxes, nRecords
    End If

    'Calculate the sum of the taxes
    dSum = 0
    For i = 0 To nRecords - 1
        If Taxes(i).m_cCharge = vbKeyY Then
            dSum = dSum + tfnRound(Taxes(i).m_dAmount, 2)
        End If
    Next i
    dTaxAmount = dSum
    fnCalculateTaxes = True
End Function

Private Function fnGetTaxGroup(sProdCode As String) As String
    'Ma, 2/26/99 Temporary patch for Taxes Loading Fee
    Const SUB_NAME = "fnGetTaxGroup"
    
    Dim strSQL As String
    Dim rsTemp As Recordset
    Dim i As Integer
    
    For i = 0 To nTaxGroupCount - 1
        If sProdCode = aryTaxGroup(i).m_sProdCode Then
            fnGetTaxGroup = aryTaxGroup(i).m_sGroup
            Exit Function
        End If
    Next i
    
    If nTaxGroupCount = 0 Then
        ReDim aryTaxGroup(ARRAY_SIZE_INC)
    Else
        If nTaxGroupCount > UBound(aryTaxGroup) Then
            ReDim Preserve aryTaxGroup(UBound(aryTaxGroup) + ARRAY_SIZE_INC)
        End If
    End If
    strSQL = "SELECT icl_fuel_group FROM inv_classes, inv_header" _
           & " WHERE icl_code = ivh_prodtcl" _
           & " AND icl_type = 'T'" _
           & " AND ivh_product = " & tfnSQLString(sProdCode)
    On Error GoTo errGetCode
    Set rsTemp = t_dbMainDatabase.OpenRecordset(strSQL, dbOpenSnapshot, dbSQLPassThrough)
    If rsTemp.RecordCount > 0 Then
        aryTaxGroup(i).m_sProdCode = sProdCode
        If IsNull(rsTemp!icl_fuel_group) Then
            aryTaxGroup(i).m_sGroup = ""
        Else
            aryTaxGroup(i).m_sGroup = Trim(rsTemp!icl_fuel_group)
        End If
    End If
    fnGetTaxGroup = aryTaxGroup(i).m_sGroup
    Exit Function
errGetCode:
    tfnErrHandler strSQL, SUB_NAME
End Function

Public Function fnGetTypeATax() As Double
    'Ma, 3/3/99 Temporary patch for Taxes Loading Fee
    Dim i As Integer
    Dim j As Integer
    
    fnGetTypeATax = 0
    For i = 0 To nInfoCount - 1
        For j = 0 To aryTaxInfo(i).m_nRecord - 1
            If aryTaxInfo(i).m_Taxes(j).m_cCharge <> vbKeyN Then
                If aryTaxInfo(i).m_Taxes(j).m_cBasis = vbKeyA Then
                    fnGetTypeATax = fnGetTypeATax + aryTaxInfo(i).m_Taxes(j).m_dAmount
                End If
            End If
        Next j
    Next i
End Function

Public Function fnRetrieveLineTax(ByVal nIdx As Integer) As Double
    'Ma, 2/26/99 Temporary patch for Taxes Loading Fee
    Dim i As Integer
    fnRetrieveLineTax = 0
    If nIdx >= 0 And nIdx < nInfoCount Then
        With aryTaxInfo(nIdx)
            For i = 0 To .m_nRecord - 1
                If .m_Taxes(i).m_cCharge = vbKeyY Then
                    fnRetrieveLineTax = fnRetrieveLineTax + .m_Taxes(i).m_dAmount
                End If
            Next i
        End With
    End If
End Function

Public Sub subAddInternalAmount(ByVal lAmount As Long)
    'Ma, 4/5/99 Temporary patch for Taxes Loading Fee
    lInternalFuel = lInternalFuel + lAmount
End Sub

Public Sub subFixTypeATax()
    'Ma, 2/26/99 Temporary patch for Taxes Loading Fee
    subFixTypeATaxForInfo
End Sub

Public Sub subFixTypeATaxForInfo()
    'Ma, 3/1/99 Temporary patch for Taxes Loading Fee

    Dim i As Integer
    Dim j As Integer
    Dim nDel As Integer
    Dim lFuelAmount As Long
    Dim lTaxableAmount As Long
    Dim lGasAmount As Long
    Dim lTotalFuel As Long
    Dim sParm1055 As String
    Dim fTaxes As Double
    Dim fTaxRate As Double
    Dim fTotal As Double
    Dim lDevNumber As Long
    Dim sTaxGroup As String
    Dim nIdx As Integer
    Dim bTexasLDFee As Boolean
    
    Dim nArrayIndex(50) As Integer
    
    'david 12/04/2001
    Dim nSign As Integer
    '''''''''''''''''
    
    If nInfoCount <= 0 Then
        Exit Sub
    End If
    
    ReDim ldelnumbers(nInfoCount - 1) As Long
    ReDim lDelFuelAmount(nInfoCount - 1) As Long
    ReDim fTypeATaxes(nInfoCount - 1) As Double
    Dim nDelCount As Integer
    
    nDelCount = 0
    For i = 0 To nInfoCount - 1
        For j = 0 To nDelCount - 1
            If ldelnumbers(j) = aryTaxInfo(i).m_lInvoice Then
                Exit For
            End If
        Next j
        If j = nDelCount Then
            ldelnumbers(j) = aryTaxInfo(i).m_lInvoice
            nDelCount = nDelCount + 1
        End If
    Next i
    
    sParm1055 = "NONE"
    lTotalFuel = lInternalFuel
    
    For nDel = 0 To nDelCount - 1
        nTaxItems = 0
        lFuelAmount = 0
        lGasAmount = 0
        
        For nIdx = 0 To nInfoCount - 1
            With aryTaxInfo(nIdx)
                If .m_lInvoice = ldelnumbers(nDel) Then
                    For i = 0 To .m_nRecord - 1
                        If .m_Taxes(i).m_cCharge = vbKeyY Then
                            If .m_Taxes(i).m_cBasis = vbKeyA Then
                                For j = 0 To nTaxItems - 1
                                    If arryTaxSums(j).m_lTrn = .m_Taxes(i).m_lTrn Then
                                        Exit For
                                    End If
                                Next j
                                If j = nTaxItems Then
                                    If UBound(arryTaxSums) < nTaxItems Then
                                        ReDim Preserve arryTaxSums(UBound(arryTaxSums) + ARRAY_SIZE_INC)
                                    End If
                                    arryTaxSums(j).m_lTrn = .m_Taxes(i).m_lTrn
                                    arryTaxSums(j).m_lLimit = .m_Taxes(i).m_lLimit
                                    arryTaxSums(j).m_cBasis = .m_Taxes(i).m_cBasis
                                    arryTaxSums(j).m_fAmount = .m_Taxes(i).m_dAmount
                                    arryTaxSums(j).m_fRate = .m_Taxes(i).m_fRate
                                    arryTaxSums(j).m_lGallons = .m_InputParms.m_dGrossUnits
                                    nTaxItems = nTaxItems + 1
                                Else
                                    arryTaxSums(j).m_lGallons = arryTaxSums(j).m_lGallons + .m_InputParms.m_dGrossUnits
                                End If
                                If Trim(.m_Taxes(i).m_sWhere) = "TX" And .m_Taxes(i).m_cLevel = vbKeyS Then
                                    sParm1055 = sSysParm1055
                                    bTexasLDFee = True
                                End If
                            End If
                        End If
                    Next i
                    
                    sTaxGroup = fnGetTaxGroup(.m_InputParms.m_sProduct)
                    
                    If sTaxGroup <> sParm1055 Then
                        lGasAmount = lGasAmount + .m_InputParms.m_dGrossUnits
                    End If
                    
                    lFuelAmount = lFuelAmount + .m_InputParms.m_dGrossUnits
                End If
            End With
        Next nIdx
        
        'When we are here. We should have the same amount for all arryTaxSums.m_lGallons
        ' And they should be equal to lFuelAmount.
        ' If they are not equal. There is a problem and need to know how to calculate.
        ' Rgiht now, we assume that they are equal
        
        #If Not HARD_CODED_TX_LD_FEE Then
            If nTaxItems > 0 Then
                'david 12/03/2001
                nSign = 1
                If lFuelAmount < 0 Then
                    nSign = -1
                    lFuelAmount = Abs(lFuelAmount)
                    lGasAmount = Abs(lGasAmount)
                End If
                '''''''''''''''''
                
                lTaxableAmount = lFuelAmount
                
                If bTexasLDFee Then
                    'For Texas only
                    If lGasAmount = lFuelAmount Then
                        If lGasAmount >= 7000 And lGasAmount < 8000 Then
                            lTaxableAmount = 8000
                        End If
                    End If
                End If
                
                fTypeATaxes(nDel) = arryTaxSums(nTaxItems - 1).m_lLimit
                
                For i = 0 To nTaxItems - 1
                    If lTaxableAmount < Abs(arryTaxSums(i).m_lLimit) Then
                        fTypeATaxes(nDel) = arryTaxSums(i).m_fRate * nSign
                        Exit For
                    End If
                Next i
                
                lDelFuelAmount(nDel) = lFuelAmount * nSign
                lTotalFuel = lTotalFuel + lFuelAmount * nSign
            Else
                fTypeATaxes(nDel) = 0
            End If
        #Else
        '---------------------Start of Special Codes for texas loading fees  WJ 10/08/01 ------------------
            'The codes (above) commented out are the old ones
            Dim m As Integer
            
            If nTaxItems > 0 Then
                'david 12/03/2001
                nSign = 1
                If lFuelAmount < 0 Then
                    nSign = -1
                    lFuelAmount = Abs(lFuelAmount)
                    lGasAmount = Abs(lGasAmount)
                End If
                '''''''''''''''''
                
                lTaxableAmount = lFuelAmount
                
                If bTexasLDFee Then
                    'For Texas only
                    If lGasAmount = lFuelAmount Then
                        If lGasAmount >= 7000 And lGasAmount < 8000 Then
                            lTaxableAmount = 8000
                        End If
                    End If
                    
                    'Vijaya on 08/18/03 Magic#417233
                    'Coming sept 1st2003, The Tx Loading Fee rates are changing.
                    
                    If lTaxableAmount = 0 Then
                       fTypeATaxes(nDel) = 0#
                    ElseIf lTaxableAmount < 2500 Then
                       fTypeATaxes(nDel) = 10# * nSign
                    ElseIf lTaxableAmount < 5000 Then
                       fTypeATaxes(nDel) = 20# * nSign
                    ElseIf lTaxableAmount < 8000 Then
                       fTypeATaxes(nDel) = 30# * nSign
                    ElseIf lTaxableAmount < 10000 Then
                       fTypeATaxes(nDel) = 40# * nSign
                    Else
                       m = 3
                       
                       Do While True
                          If lTaxableAmount < m * 5000 Then
                             fTypeATaxes(nDel) = m * 20# * nSign
                             Exit Do
                          Else
                             m = m + 1
                          End If
                       Loop
                    End If
                Else
                    fTypeATaxes(nDel) = arryTaxSums(nTaxItems - 1).m_lLimit
                    
                    For i = 0 To nTaxItems - 1
                        If lTaxableAmount < Abs(arryTaxSums(i).m_lLimit) Then
                            fTypeATaxes(nDel) = arryTaxSums(i).m_fRate * nSign
                            Exit For
                        End If
                    Next i
                End If
            
                lDelFuelAmount(nDel) = lFuelAmount * nSign
                lTotalFuel = lTotalFuel + lFuelAmount * nSign
            Else
                fTypeATaxes(nDel) = 0
            End If
            '--------------------- End of Special Codes for texas loading fees WJ 10/08/01 ------------------
        #End If
    Next nDel
    
    fTaxes = 0
    
    For nDel = 0 To nDelCount - 1
        fTaxes = fTaxes + tfnRound(fTypeATaxes(nDel), 2)
    Next nDel
    '#Magic #405291 WJ 04/24/2003
    'If fTaxes > fLadingTypeATax Then
    If Abs(fTaxes) > Abs(fLadingTypeATax) Then
        fTaxes = fLadingTypeATax
        'Need to redistribute among the delivery items
        If lTotalFuel = 0 Then
            fTaxRate = 0
        Else
            fTaxRate = fTaxes / lTotalFuel
        End If
        fTotal = 0
        For i = 0 To nDelCount - 1
            If fTypeATaxes(i) <> 0 Then
                fTypeATaxes(i) = tfnRound(fTaxRate * lDelFuelAmount(i), 2)
            End If
            fTotal = fTotal + fTypeATaxes(i)
        Next i
        fTaxRate = tfnRound(fTotal, 2) - fTaxes
        If Abs(fTaxRate) > 0.01 And Abs(Abs(fTaxRate)) < 0.05 Then
            'Deal with rounding error
            fTypeATaxes(0) = fTypeATaxes(0) - fTaxRate
        End If
    End If
    
    Dim bApplied As Boolean
    
    If fTaxes <> 0 Then
        'Now Need to redistribute to the individual line amount
        For nDel = 0 To nDelCount - 1
            If fTypeATaxes(nDel) <> 0 Then
                If lDelFuelAmount(nDel) = 0 Then
                    fTaxRate = 0
                Else
                    fTaxRate = fTypeATaxes(nDel) / lDelFuelAmount(nDel)
                End If
                For nIdx = 0 To nInfoCount - 1
                    With aryTaxInfo(nIdx)
                        If .m_lInvoice = ldelnumbers(nDel) Then
                            
                            bApplied = False
                            
                            For i = 0 To .m_nRecord - 1
                                If .m_Taxes(i).m_cCharge = vbKeyY Then
                                    If .m_Taxes(i).m_cBasis = vbKeyA Then
                                        #If HARD_CODED_TX_LD_FEE Then
        '--------------------- Start of Special Codes for texas loading fees WJ 10/08/01 ------------------
                                            If Trim(.m_Taxes(i).m_sWhere) = "TX" And .m_Taxes(i).m_cLevel = vbKeyS Then
                                                If bApplied Then
                                                  .m_Taxes(i).m_dAmount = 0
                                                Else
                                                  .m_Taxes(i).m_dAmount = tfnRound(fTaxRate * .m_InputParms.m_dGrossUnits, 2)
                                                  '.m_Taxes(i).m_fRate = .m_Taxes(i).m_dAmount
                                                  bApplied = True
                                                End If
                                            Else
                                                 If .m_Taxes(i).m_dAmount <> 0 Then
                                                   .m_Taxes(i).m_dAmount = tfnRound(fTaxRate * .m_InputParms.m_dGrossUnits, 2)
                                                 End If
                                            End If
                                        'The codes (below) commented out are the old ones
         '--------------------- End of Special Codes for texas loading fees WJ 10/08/01 ------------------
                                        #Else
                                            If .m_Taxes(i).m_dAmount <> 0 Then
                                                .m_Taxes(i).m_dAmount = tfnRound(fTaxRate * .m_InputParms.m_dGrossUnits, 2)
                                            End If
                                        #End If
                                    End If
                                End If
                            Next i
                        End If
                    End With
                Next nIdx
            End If
        Next nDel
    End If
End Sub

Public Sub subRetrieveTaxInfo(Taxes() As CApplicableTaxes, _
                              ByRef nRecords As Long, _
                              ByVal lInvoice As Long, _
                              ByVal nLine As Integer, _
                              Optional vSumTypeA As Variant)
    'Ma, 3/1/99 Temporary patch for Taxes Loading Fee
    'When lInvoice < 0, nLine is the Sequence of adding in,
    'When lInvoice >= 0, nLine is the Delivery/Lading line number
    
    Dim i As Integer
    Dim j As Integer
    Dim bSum As Boolean
    Dim bFlag As Boolean
    Dim nIdxLine As Integer
    Dim lBaseAmount As Long
    
    If IsMissing(vSumTypeA) Then
        bSum = True
    Else
        bSum = vSumTypeA
    End If
    
    If lInvoice >= 0 Then
        For i = 0 To nInfoCount - 1
            If lInvoice = aryTaxInfo(i).m_lInvoice Then
                If nLine = aryTaxInfo(i).m_nLine Then
                    nLine = i
                    Exit For
                End If
            End If
        Next i
        
        If i >= nInfoCount Then
            nLine = -1
        End If
    End If
    
    If nLine >= 0 And nLine < nInfoCount Then
        nRecords = aryTaxInfo(nLine).m_nRecord
        
        If nRecords > 0 Then
            ReDim Taxes(nRecords - 1)
            For i = 0 To nRecords - 1
                Taxes(i) = aryTaxInfo(nLine).m_Taxes(i)
            Next i
        End If
        
        If bSum Then
            nIdxLine = -1
            
            For i = 0 To nInfoCount - 1
                If aryTaxInfo(i).m_lInvoice = lInvoice Then
                    For j = 0 To aryTaxInfo(i).m_nRecord - 1
                        If aryTaxInfo(i).m_Taxes(j).m_cCharge = vbKeyY Then
                            If aryTaxInfo(i).m_Taxes(j).m_cBasis = vbKeyA Then
                                If aryTaxInfo(i).m_Taxes(j).m_dAmount <> 0 Then
                                    'The first line with type A taxes
                                    nIdxLine = i
                                End If
                            End If
                        End If
                    Next j
                End If
                If nIdxLine >= 0 Then
                    Exit For
                End If
            Next i
            
            If nIdxLine = nLine Then
                'Sum Type A Amount
                For i = 0 To nRecords - 1
                    If Taxes(i).m_cBasis = vbKeyA Then
                        If Taxes(i).m_dAmount <> 0 Then
                            nIdxLine = i
                            Exit For
                        End If
                    End If
                Next i
                
                lBaseAmount = aryTaxInfo(nLine).m_InputParms.m_dGrossUnits
                
                For i = 0 To nInfoCount - 1
                    If aryTaxInfo(i).m_lInvoice = lInvoice Then
                        If i <> nLine Then
                            bFlag = False
                            
                            For j = 0 To aryTaxInfo(i).m_nRecord - 1
                                If aryTaxInfo(i).m_Taxes(j).m_cCharge = vbKeyY Then
                                    If aryTaxInfo(i).m_Taxes(j).m_cBasis = vbKeyA Then
                                        If aryTaxInfo(i).m_Taxes(j).m_dAmount <> 0 Then
                                            Taxes(nIdxLine).m_dAmount = Taxes(nIdxLine).m_dAmount + aryTaxInfo(i).m_Taxes(j).m_dAmount
                                            bFlag = True
                                        End If
                                    End If
                                End If
                            Next j
                            
                            If bFlag Then
                                lBaseAmount = lBaseAmount + aryTaxInfo(i).m_InputParms.m_dGrossUnits
                            End If
                        End If
                    End If
                Next i
                Taxes(nIdxLine).m_lGlAcct = lBaseAmount
            Else
                'Set Type A tax amount to 0
                For i = 0 To nRecords - 1
                    If Taxes(i).m_cBasis = vbKeyA Then
                        Taxes(i).m_dAmount = 0
                    End If
                Next i
            End If
        End If
    Else
        nRecords = 0
    End If
End Sub

Public Sub subSetLadingTypeATax(fTemp As Double)
    fLadingTypeATax = fTemp
End Sub

Public Sub subSetSysParm1055(sTemp As String)
    'Ma, 2/26/99 Temporary patch for Taxes Loading Fee
    If sTemp = "" Then
        sSysParm1055 = "NONE"
    Else
        sSysParm1055 = sTemp
    End If
    ReDim aryTaxGroup(ARRAY_SIZE_INC)
    nTaxGroupCount = 0
End Sub

Public Sub subUnloadTaxMods(Optional bShowMsgBox As Boolean = True, _
                            Optional sErrMsg As String = "")
    'Unload the tax mods
    If nTaxModLoaded = TAXMOD_LOADED Then
        Dim i As Integer
        Do
            i = TXMOD_Detach(App.hInstance, Forms(0).hwnd)
            DoEvents
        Loop Until i <> TXMOD_TASKBLOCKED
        If i <> TXMOD_OK Then
            If bShowMsgBox Then
                MsgBox "There is a problem detaching from tax module", vbCritical
            Else
                sErrMsg = "There is a problem detaching from tax module"
            End If
        End If
    End If
End Sub

Private Sub subTaxModErrors(ByVal nMsgNo As Long, _
                            Optional bShowMsgBox As Boolean = True, _
                            Optional sErrMsg As String = "")
    
    Static nLastError As Integer
    Dim sMsg As String
        
    If nLastError = nMsgNo Then
        Exit Sub
    End If
    nLastError = nMsgNo

    'Display errors during the tax mod call
    Const nButtons = vbOKOnly + vbCritical
    
    Select Case nMsgNo
        Case TXMOD_LINETAX_OK = 0           ' All's well
        Case TXMOD_LINETAX_NOMEM         ' Out of memory
            sMsg = "Out of memory for tax calculation"
        Case TXMOD_LINETAX_NODBASE       ' No database available
            sMsg = "No database availble for tax calculation"
        Case TXMOD_LINETAX_NOTAXRECS     ' No tax records in the dbase
            sMsg = "No tax records for the customer/product"
        Case TXMOD_LINETAX_NOINVHEAD     ' No inventory header record
            sMsg = "No inventory header record available"
        Case TXMOD_LINETAX_NOINVCLASS    ' No inventory class record
            sMsg = "No inventory class record available"
        Case TXMOD_LINETAX_BADSQL        ' SQL statement error
            sMsg = "An error exists in the SQL statement for tax calculations"
        Case TXMOD_LINETAX_BADBASIS      ' Bad basis flag
            sMsg = "Bad basis flag from tax mods"
        Case TXMOD_LINETAX_BADBILLED     ' Bad billed units flag
            sMsg = "Bad billed units flag"
        Case TXMOD_LINETAX_BADLTYPE      ' Bad line type (int/ext bl/del)
            sMsg = "Bad line type"
        Case TXMOD_LINETAX_BADPARMS     ' Bad parameters passed
            sMsg = "Bad parameters for tax calculation"
    End Select

    If bShowMsgBox Then
        MsgBox sMsg + ".", nButtons
    Else
        sErrMsg = sMsg + "."
    End If

End Sub

'// Calculates and returns the taxes
'david 08/29/2003  #393678-4
'add two more optional parameters
Public Function fnGetTaxAmount(szProduct As String, _
                               ByVal lNet As Long, _
                               ByVal lGross As Long, _
                               ByVal fCurrentCost As Double, _
                               ByVal fFreight As Double, _
                               szTaxUseGroup As String, _
                               szBillingFlag As String, _
                               ByVal bExternal As Boolean, _
                               ByVal bLading As Boolean, _
                               Optional bShowMsgBox As Boolean = True, _
                               Optional sErrMsg As String = "", _
                               Optional bApplyLimit As Boolean = True, _
                               Optional hld_vessel As Integer = -1, _
                               Optional hld_gals As Double = 0#, _
                               Optional hld_miles As Double = 0#, _
                               Optional hld_tot_gals As Double = 0#) As Double

    Dim fDollarBase As Double
    Dim fQuantity As Double
    'Calculate taxes by calling the tax mods
    Dim myStruct As CMovementLineItem
    Dim Taxes() As CApplicableTaxes
    Dim nRet As Long
    Dim nRecords As Long
    Dim dSum As Double
    Dim i As Integer
    Dim nLineType As Long
    Dim nTries As Integer

    fnGetTaxAmount = 0
    'if the tax mods are not loaded, load them
    If Not fnLoadTaxMods Then
        Exit Function
    End If

    If (szBillingFlag = "N") Then
        fDollarBase = fCurrentCost * lNet
    Else
        fDollarBase = fCurrentCost * lGross
    End If
    If bLading Then
        If bExternal Then
            nLineType = TXMOD_EXTBOL   ' External bill of lading
        Else
            nLineType = TXMOD_INTBOL  ' Internal bill of lading
        End If
    Else
        If bExternal Then
            nLineType = TXMOD_EXTDEL   ' External delivery
        Else
            nLineType = TXMOD_INTDEL   ' Internal delivery
        End If
    End If

    ' assign initial values to the input structure
    With myStruct
        .m_sProduct = szProduct
        .m_sUseGroup = szTaxUseGroup
        .m_dDollarBase = fDollarBase
        .m_dFreight = fFreight                    'Freight amount
        .m_dNetUnits = lNet
        .m_dGrossUnits = lGross
        .m_cBilledUnitsFlag = Asc(szBillingFlag)
        .m_tLineType = nLineType
    End With
    
    'Determine the number of items in the record
    nTries = TAXMOD_TRIES
    'david 08/29/2003  #393678-4
    subSet_gl_sd_gals hld_vessel, hld_gals, hld_miles, hld_tot_gals
    ''''''''''''''''''''''''''''
    'Vijaya on 02/10/04 Magic#395302
    subSet_tx_date t_tax_date
    ''''''''''''''''''''''''''''''''
    
    Do
        nRet = TXMOD_VB_CalculateLineTax(App.hInstance, Forms(0).hwnd, myStruct, nRecords)
        If i = TXMOD_TASKBLOCKED Then
            DoEvents
        End If
        nTries = nTries - 1
    Loop Until nRet <> TXMOD_TASKBLOCKED Or nTries <= 0
    
    If nRet <> TXMOD_OK Then
        subTaxModErrors nRet, bShowMsgBox, sErrMsg
        Exit Function
    End If

    ' Calculate the taxes
    ReDim Taxes(nRecords - 1)
    
    Do
        nRet = TXMOD_GetLineTaxes(App.hInstance, Forms(0).hwnd, Taxes)
        If i = TXMOD_TASKBLOCKED Then
            DoEvents
        End If
        nTries = nTries - 1
    Loop Until nRet <> TXMOD_TASKBLOCKED Or nTries <= 0
    
    If nRet <> TXMOD_OK Then
        subTaxModErrors nRet, bShowMsgBox, sErrMsg
        Exit Function
    End If
    
    subDealWithATaxes Taxes, nRecords
    
    If bApplyLimit Then
        subApplyLimit Taxes, nRecords
    End If
    
    'Calculate the sum of the taxes
    dSum = 0
    For i = 0 To nRecords - 1
        If Taxes(i).m_cCharge = vbKeyY Then
            dSum = dSum + tfnRound(Taxes(i).m_dAmount, 2)
        End If
    Next i
    fnGetTaxAmount = dSum
    
End Function

'// Calculates and returns the taxes
'david 08/29/2003  #393678-4
'add two more optional parameters
Public Function fnGetTaxAmount1(szProduct As String, _
                               ByVal lNet As Double, _
                               ByVal lGross As Double, _
                               ByVal fDollarBase As Double, _
                               ByVal fFreight As Double, _
                               szTaxUseGroup As String, _
                               szBillingFlag As String, _
                               ByVal bExternal As Boolean, _
                               ByVal bLading As Boolean, _
                               Optional bShowMsgBox As Boolean = True, _
                               Optional sErrMsg As String = "", _
                               Optional bApplyLimit As Boolean = True, _
                               Optional hld_vessel As Integer = -1, _
                               Optional hld_gals As Double = 0#, _
                               Optional hld_miles As Double = 0#, _
                               Optional hld_tot_gals As Double = 0#) As Double

    Dim fQuantity As Double
    'Calculate taxes by calling the tax mods
    Dim myStruct As CMovementLineItem
    Dim Taxes() As CApplicableTaxes
    Dim nRet As Long
    Dim nRecords As Long
    Dim dSum As Double
    Dim i As Integer
    Dim nLineType As Long
    Dim nTries As Integer

    fnGetTaxAmount1 = 0
    'if the tax mods are not loaded, load them
    If Not fnLoadTaxMods Then
        Exit Function
    End If

    If bLading Then
        If bExternal Then
            nLineType = TXMOD_EXTBOL   ' External bill of lading
        Else
            nLineType = TXMOD_INTBOL  ' Internal bill of lading
        End If
    Else
        If bExternal Then
            nLineType = TXMOD_EXTDEL   ' External delivery
        Else
            nLineType = TXMOD_INTDEL   ' Internal delivery
        End If
    End If

    ' assign initial values to the input structure
    With myStruct
        .m_sProduct = szProduct
        .m_sUseGroup = szTaxUseGroup
        .m_dDollarBase = fDollarBase
        .m_dFreight = fFreight                    'Freight amount
        .m_dNetUnits = lNet
        .m_dGrossUnits = lGross
        .m_cBilledUnitsFlag = Asc(szBillingFlag)
        .m_tLineType = nLineType
    End With
    
    'Determine the number of items in the record
    nTries = TAXMOD_TRIES
    'david 08/29/2003  #393678-4
    subSet_gl_sd_gals hld_vessel, hld_gals, hld_miles, hld_tot_gals
    ''''''''''''''''''''''''''''
    'Vijaya on 02/10/04 Magic#395302
    subSet_tx_date t_tax_date
    ''''''''''''''''''''''''''''''''
    
    Do
        nRet = TXMOD_VB_CalculateLineTax(App.hInstance, Forms(0).hwnd, myStruct, nRecords)
        If i = TXMOD_TASKBLOCKED Then
            DoEvents
        End If
        nTries = nTries - 1
    Loop Until nRet <> TXMOD_TASKBLOCKED Or nTries <= 0
    
    If nRet <> TXMOD_OK Then
        subTaxModErrors nRet, bShowMsgBox, sErrMsg
        Exit Function
    End If

    ' Calculate the taxes
    ReDim Taxes(nRecords - 1)
    
    Do
        nRet = TXMOD_GetLineTaxes(App.hInstance, Forms(0).hwnd, Taxes)
        If i = TXMOD_TASKBLOCKED Then
            DoEvents
        End If
        nTries = nTries - 1
    Loop Until nRet <> TXMOD_TASKBLOCKED Or nTries <= 0
    
    If nRet <> TXMOD_OK Then
        subTaxModErrors nRet, bShowMsgBox, sErrMsg
        Exit Function
    End If
    
    subDealWithATaxes Taxes, nRecords
    
    If bApplyLimit Then
        subApplyLimit Taxes, nRecords
    End If
    
    'Calculate the sum of the taxes
    dSum = 0
    For i = 0 To nRecords - 1
        If Taxes(i).m_cCharge = vbKeyY Then
            dSum = dSum + tfnRound(Taxes(i).m_dAmount, 2)
        End If
    Next i
    fnGetTaxAmount1 = dSum
    
End Function

'david 08/29/2003  #393678-4
'add two more optional parameters
Public Sub subAddTaxItem(szProduct As String, _
                          ByVal lNet As Long, _
                          ByVal lGross As Long, _
                          ByVal fCurrentCost As Double, _
                          ByVal fFreight As Double, _
                          szTaxUseGroup As String, _
                          szBillingFlag As String, _
                          ByVal bExternal As Boolean, _
                          ByVal bLading As Boolean, _
                          Optional hld_vessel As Integer = -1, _
                          Optional hld_gals As Double = 0#, _
                          Optional hld_miles As Double = 0#, _
                          Optional hld_tot_gals As Double = 0#)
                          
    'Ma, 2/26/99 Temporary patch for Taxes Loading Fee
    
    Dim fDollarBase As Double
    Dim fQuantity As Double
    Dim nLineType As Integer

    If (szBillingFlag = "N") Then
        fDollarBase = fCurrentCost * lNet
    Else
        fDollarBase = fCurrentCost * lGross
    End If
    If bLading Then
        If bExternal Then
            nLineType = TXMOD_EXTBOL   ' External bill of lading
        Else
            nLineType = TXMOD_INTBOL  ' Internal bill of lading
            lInternalFuel = lInternalFuel + lGross
        End If
    Else
        If bExternal Then
            nLineType = TXMOD_EXTDEL   ' External delivery
        Else
            nLineType = TXMOD_INTDEL   ' Internal delivery
            lInternalFuel = lInternalFuel + lGross
        End If
    End If
    
    'david 08/29/2003  #393678-4
    fnAddTaxInfoItem -1, nInfoCount, szProduct, szTaxUseGroup, fDollarBase, _
        fFreight, CDbl(lGross), CDbl(lNet), szBillingFlag, nLineType, , , _
        hld_vessel, hld_gals, hld_miles, hld_tot_gals
    ''''''''''''''''''''''''''''
End Sub

'Add By junsong 06/04/2002 to fix deciaml gallon
'david 08/29/2003  #393678-4
'add two more optional parameters
Public Sub subAddTaxItem1(szProduct As String, _
                          ByVal fNet As Double, _
                          ByVal fGross As Double, _
                          ByVal fCurrentCost As Double, _
                          ByVal fFreight As Double, _
                          szTaxUseGroup As String, _
                          szBillingFlag As String, _
                          ByVal bExternal As Boolean, _
                          ByVal bLading As Boolean, _
                          Optional hld_vessel As Integer = -1, _
                          Optional hld_gals As Double = 0#, _
                          Optional hld_miles As Double = 0#, _
                          Optional hld_tot_gals As Double = 0#)
                          
    'Ma, 2/26/99 Temporary patch for Taxes Loading Fee
    
    Dim fDollarBase As Double
    Dim fQuantity As Double
    Dim nLineType As Integer

    If (szBillingFlag = "N") Then
        fDollarBase = fCurrentCost * fNet
    Else
        fDollarBase = fCurrentCost * fGross
    End If
    If bLading Then
        If bExternal Then
            nLineType = TXMOD_EXTBOL   ' External bill of lading
        Else
            nLineType = TXMOD_INTBOL  ' Internal bill of lading
            lInternalFuel = lInternalFuel + round(fGross)
        End If
    Else
        If bExternal Then
            nLineType = TXMOD_EXTDEL   ' External delivery
        Else
            nLineType = TXMOD_INTDEL   ' Internal delivery
            lInternalFuel = lInternalFuel + round(fGross)
        End If
    End If
    
    'david 08/29/2003  #393678-4
    fnAddTaxInfoItem -1, nInfoCount, szProduct, szTaxUseGroup, fDollarBase, _
        fFreight, fGross, fNet, szBillingFlag, nLineType, , , _
        hld_vessel, hld_gals, hld_miles, hld_tot_gals
    ''''''''''''''''''''''''''''

End Sub

Private Sub subApplyLimit(Taxes() As CApplicableTaxes, _
                          ByVal nCount As Integer)

    Dim i As Integer
    Dim j As Integer
    Dim bFind As Boolean
    
    For i = 0 To nCount - 1
        bFind = False
        If Taxes(i).m_cLimitType = vbKeyI Then
            For j = 0 To nTaxItems - 1
                If Taxes(i).m_lTrn = arryTaxSums(j).m_lTrn Then
                    bFind = True
                    Exit For
                End If
            Next j
            If Not bFind Then
                If nTaxItems > nArraySize Then
                    nArraySize = nArraySize + ARRAY_SIZE_INC
                    ReDim Preserve arryTaxSums(nArraySize)
                End If
                With arryTaxSums(nTaxItems)
                    .m_lTrn = Taxes(i).m_lTrn
                    .m_lLimit = Taxes(i).m_lLimit
                    .m_cLimitType = Taxes(i).m_cLimitType
                    .m_cBasis = Taxes(i).m_cBasis
                    .m_fAmount = .m_lLimit
                End With
                j = nTaxItems
                nTaxItems = nTaxItems + 1
            End If
            
            'Consider the total tax for the specific tran number
            With arryTaxSums(j)
                If .m_cBasis <> vbKeyA Then
                    If .m_lLimit > 0 Then
                        If Abs(tfnRound(Taxes(i).m_dAmount, 2)) > .m_fAmount Then
                            Taxes(i).m_dAmount = tfnRound(Sgn(Taxes(i).m_dAmount) * .m_fAmount, 2)
                            .m_fAmount = 0#
                        Else
                            .m_fAmount = .m_fAmount - tfnRound(Abs(Taxes(i).m_dAmount), 2)
                        End If
                    End If
                End If
            End With
        End If
    Next i
End Sub


Public Sub subResetTaxes(Optional vTwoD As Variant, _
                         Optional vResetTypeATax As Variant)
    'Ma, 2/26/99 Temporary patch for Taxes Loading Fee
    nTaxItems = 0
    nArraySize = ARRAY_SIZE_INC
    ReDim arryTaxSums(nArraySize)

    With arryTaxSums(nTaxItems)
        .m_lTrn = -1
        .m_lLimit = 0
        .m_cBasis = vbKeySpace
        .m_fAmount = 0
    End With
    If Not IsMissing(vResetTypeATax) Then
        If vResetTypeATax Then
            fLadingTypeATax = 1000000
        End If
    End If
    If Not IsMissing(vTwoD) Then
        If vTwoD Then
            nInfoCount = 0
            ReDim aryTaxInfo(ARRAY_SIZE_INC)
        End If
    End If
    lInternalFuel = 0
End Sub
Public Function fnLoadTaxMods(Optional bShowMsgBox As Boolean = True, _
                              Optional sErrMsg As String = "", _
                              Optional bDisableATax As Boolean = False)
    Const SUB_NAME = "fnLoadTaxMods"
    
    Static nTaxModLoaded As Integer
    
    On Error GoTo errLoadTaxMods
    If nTaxModLoaded = TAXMOD_LOADED Then
        fnLoadTaxMods = True
        Exit Function
    End If
    fnLoadTaxMods = False
    If nTaxModLoaded = TAXMOD_LOADING_FAILED Then
        Exit Function
    End If

    'Load the tax mods
    Dim i As Long
    Dim szConnect As String

    nTaxModLoaded = TAXMOD_LOADING_FAILED
    If t_dbMainDatabase Is Nothing Then
        'szConnect = "ODBC;DSN=gasup32;DB=/factor/gasup/factor;HOST=;SRVR=ether5;UID=ma;PWD=menus"
        Exit Function
    Else
        szConnect = t_dbMainDatabase.Connect
    End If
    
    Do
        i = TXMOD_Attach(App.hInstance, Forms(0).hwnd, szConnect)
        If i = TXMOD_TASKBLOCKED Then
            DoEvents
        End If
    Loop Until (i <> TXMOD_TASKBLOCKED)
    If i <> TXMOD_OK Then
        If bShowMsgBox Then
            MsgBox "Cannot load tax module. Taxes will not be displayed.", vbCritical
        Else
            sErrMsg = "Cannot load tax module. Taxes will not be displayed."
        End If
    Else
        nTaxModLoaded = TAXMOD_LOADED
        m_Disable_A_Taxes = bDisableATax
        fnLoadTaxMods = True
    End If
    subResetTaxes
    Exit Function
    
errLoadTaxMods:
    If Err.Number = 53 Then
        tfnErrHandler SUB_NAME, ERROR_LOADING1, MOD_DLL_NAME & " is not found in Windows System directory"
    Else
        tfnErrHandler SUB_NAME, ERROR_LOADING2, Err.Description
    End If
    nTaxModLoaded = TAXMOD_LOADING_FAILED

End Function

'david 08/29/2003  #393678-4
'add two more optional parameters
Public Function fnTaxInformation(sProduct As String, _
                                 sTaxUseGroup As String, _
                                 fDollarBase As Double, _
                                 fFreight As Double, _
                                 fGross As Double, _
                                 fNet As Double, _
                                 sBillFlag As String, _
                                 ByRef Taxes() As CApplicableTaxes, _
                                 ByRef nRecords As Long, _
                                 Optional bShowMsgBox As Boolean = True, _
                                 Optional sErrMsg As String = "", _
                                 Optional bApplyLimit As Boolean = True, _
                                 Optional hld_vessel As Integer = -1, _
                                 Optional hld_gals As Double = 0#, _
                                 Optional hld_miles As Double = 0#, _
                                 Optional hld_tot_gals As Double = 0#) As Boolean
    
    'Get the tax information array
    Dim myStruct As CMovementLineItem
    Dim nRet As Long
    Dim i As Integer
    Dim nTries As Integer

    fnTaxInformation = False
    If Not fnLoadTaxMods Then
        Exit Function
    End If

    With myStruct
        .m_sProduct = sProduct
        .m_sUseGroup = sTaxUseGroup
        .m_dDollarBase = fDollarBase
        .m_dFreight = 0
        .m_dNetUnits = fNet
        .m_dGrossUnits = fGross
        .m_dFreight = fFreight
        .m_cBilledUnitsFlag = Asc(sBillFlag)
        .m_tLineType = TXMOD_EXTDEL
    End With
    nTries = TAXMOD_TRIES
    'david 08/29/2003  #393678-4
    subSet_gl_sd_gals hld_vessel, hld_gals, hld_miles, hld_tot_gals
    ''''''''''''''''''''''''''''
    'Vijaya on 02/10/04 Magic#395302
    subSet_tx_date t_tax_date
    ''''''''''''''''''''''''''''''''
    
    Do
        nRet = TXMOD_VB_CalculateLineTax(App.hInstance, Forms(0).hwnd, myStruct, nRecords)
        If i = TXMOD_TASKBLOCKED Then
            DoEvents
        End If
        nTries = nTries - 1
    Loop Until nRet <> TXMOD_TASKBLOCKED Or nTries <= 0
    If nRet <> TXMOD_OK Then
        subTaxModErrors nRet, bShowMsgBox, sErrMsg
        fnTaxInformation = False
        Exit Function
    End If
        
    ReDim Taxes(nRecords - 1)
    nTries = TAXMOD_TRIES
    Do
        nRet = TXMOD_GetLineTaxes(App.hInstance, Forms(0).hwnd, Taxes())
        If i = TXMOD_TASKBLOCKED Then
            DoEvents
        End If
        nTries = nTries - 1
    Loop Until nRet <> TXMOD_TASKBLOCKED Or nTries <= 0
    
    If nRet <> TXMOD_OK Then
        subTaxModErrors nRet, bShowMsgBox, sErrMsg
        fnTaxInformation = False
        Exit Function
    End If
    
    subDealWithATaxes Taxes, nRecords
    
    If bApplyLimit Then
        subApplyLimit Taxes, nRecords
    End If
    fnTaxInformation = True
End Function

'david 08/29/2003  #393678-4
'add two more optional parameters
Public Function fnTaxInformation1(sProduct As String, _
                                 sTaxUseGroup As String, _
                                 fDollarBase As Double, _
                                 fFreight As Double, _
                                 fGross As Double, _
                                 fNet As Double, _
                                 sBillFlag As String, _
                                 bExternal As Boolean, _
                                 bBOLFlag As Boolean, _
                                 ByRef Taxes() As CApplicableTaxes, _
                                 ByRef nRecords As Long, _
                                 Optional bShowMsgBox As Boolean = True, _
                                 Optional sErrMsg As String = "", _
                                 Optional bApplyLimit As Boolean = True, _
                                 Optional hld_vessel As Integer = -1, _
                                 Optional hld_gals As Double = 0#, _
                                 Optional hld_miles As Double = 0#, _
                                 Optional hld_tot_gals As Double = 0#) As Boolean
    
    'Get the tax information array
    Dim myStruct As CMovementLineItem
    Dim nRet As Long
    Dim i As Integer
    Dim nTries As Integer

    fnTaxInformation1 = False
    If Not fnLoadTaxMods Then
        Exit Function
    End If

    With myStruct
        .m_sProduct = sProduct
        .m_sUseGroup = sTaxUseGroup
        .m_dDollarBase = fDollarBase
        .m_dFreight = 0
        .m_dNetUnits = fNet
        .m_dGrossUnits = fGross
        .m_dFreight = fFreight
        .m_cBilledUnitsFlag = Asc(sBillFlag)
        If bExternal Then
            If bBOLFlag Then
                .m_tLineType = TXMOD_EXTBOL
            Else
                .m_tLineType = TXMOD_EXTDEL
            End If
        Else
            If bBOLFlag Then
                .m_tLineType = TXMOD_INTBOL
            Else
                .m_tLineType = TXMOD_INTDEL
            End If
        End If
    End With
    nTries = TAXMOD_TRIES
    'david 08/29/2003  #393678-4
    subSet_gl_sd_gals hld_vessel, hld_gals, hld_miles, hld_tot_gals
    ''''''''''''''''''''''''''''
    'Vijaya on 02/10/04 Magic#395302
    subSet_tx_date t_tax_date
    ''''''''''''''''''''''''''''''''
    
    Do
        nRet = TXMOD_VB_CalculateLineTax(App.hInstance, Forms(0).hwnd, myStruct, nRecords)
        If i = TXMOD_TASKBLOCKED Then
            DoEvents
        End If
        nTries = nTries - 1
    Loop Until nRet <> TXMOD_TASKBLOCKED Or nTries <= 0
    If nRet <> TXMOD_OK Then
        subTaxModErrors nRet, bShowMsgBox, sErrMsg
        fnTaxInformation1 = False
        Exit Function
    End If
        
    ReDim Taxes(nRecords - 1)
    nTries = TAXMOD_TRIES
    Do
        nRet = TXMOD_GetLineTaxes(App.hInstance, Forms(0).hwnd, Taxes())
        If i = TXMOD_TASKBLOCKED Then
            DoEvents
        End If
        nTries = nTries - 1
    Loop Until nRet <> TXMOD_TASKBLOCKED Or nTries <= 0
    
    If nRet <> TXMOD_OK Then
        subTaxModErrors nRet, bShowMsgBox, sErrMsg
        fnTaxInformation1 = False
        Exit Function
    End If
    
    subDealWithATaxes Taxes, nRecords
    
    If bApplyLimit Then
        subApplyLimit Taxes, nRecords
    End If
    fnTaxInformation1 = True
End Function

'david 08/29/2003  #393678-4
'add two more optional parameters
Public Function fnAddTaxInfoItem(ByVal lInv As Long, _
                                 ByVal nLine As Integer, _
                                 sProduct As String, _
                                 sTaxUseGroup As String, _
                                 fDollarBase As Double, _
                                 fFreight As Double, _
                                 fGross As Double, _
                                 fNet As Double, _
                                 sBillFlag As String, _
                                 Optional vLineType As Variant, _
                                 Optional bShowMsgBox As Boolean = True, _
                                 Optional sErrMsg As String = "", _
                                 Optional hld_vessel As Integer = -1, _
                                 Optional hld_gals As Double = 0#, _
                                 Optional hld_miles As Double = 0#, _
                                 Optional hld_tot_gals As Double = 0#) As Boolean
    
    'Ma, 3/1/99 Temporary patch for Taxes Loading Fee
    'Get the tax information array
    Dim Taxes() As CApplicableTaxes
    Dim myStruct As CMovementLineItem
    Dim nRet As Long
    Dim i As Integer
    Dim j As Integer
    Dim nIdx As Integer
    Dim nTries As Integer
    Dim nRecords As Long
    
    fnAddTaxInfoItem = False
    If Not fnLoadTaxMods Then
        Exit Function
    End If

    With myStruct
        .m_sProduct = sProduct
        .m_sUseGroup = sTaxUseGroup
        .m_dDollarBase = fDollarBase
        .m_dFreight = 0
        .m_dNetUnits = fNet
        .m_dGrossUnits = fGross
        .m_dFreight = fFreight
        If sBillFlag <> "" Then
        .m_cBilledUnitsFlag = Asc(sBillFlag)
        Else
        .m_cBilledUnitsFlag = vbKeyN
        End If
        If IsMissing(vLineType) Then
            .m_tLineType = TXMOD_EXTDEL
        Else
            .m_tLineType = vLineType
        End If
    End With
    nTries = TAXMOD_TRIES
    'david 08/29/2003  #393678-4
    subSet_gl_sd_gals hld_vessel, hld_gals, hld_miles, hld_tot_gals
    ''''''''''''''''''''''''''''
    'Vijaya on 02/10/04 Magic#395302
    subSet_tx_date t_tax_date
    ''''''''''''''''''''''''''''''''
    
    Do
        nRet = TXMOD_VB_CalculateLineTax(App.hInstance, Forms(0).hwnd, myStruct, nRecords)
        
        If i = TXMOD_TASKBLOCKED Then
            DoEvents
        End If
        nTries = nTries - 1
    Loop Until nRet <> TXMOD_TASKBLOCKED Or nTries <= 0
    If nRet <> TXMOD_OK Then
        subTaxModErrors nRet, bShowMsgBox, sErrMsg
        Exit Function
    End If
        
    ReDim Taxes(nRecords - 1)
    nTries = TAXMOD_TRIES
    Do
        nRet = TXMOD_GetLineTaxes(App.hInstance, Forms(0).hwnd, Taxes())
        If i = TXMOD_TASKBLOCKED Then
            DoEvents
        End If
        nTries = nTries - 1
    Loop Until nRet <> TXMOD_TASKBLOCKED Or nTries <= 0
    
    If nRet <> TXMOD_OK Then
        subTaxModErrors nRet, bShowMsgBox, sErrMsg
        Exit Function
    End If
    
    subDealWithATaxes Taxes, nRecords
    
    For nIdx = 0 To nInfoCount - 1
        If aryTaxInfo(nIdx).m_lInvoice = lInv Then
            If aryTaxInfo(nIdx).m_nLine = nLine Then
                Exit For
            End If
        End If
    Next nIdx
    If nIdx >= nInfoCount Then
        If UBound(aryTaxInfo) < nInfoCount Then
            ReDim Preserve aryTaxInfo(UBound(aryTaxInfo) + ARRAY_SIZE_INC)
        End If
        nInfoCount = nInfoCount + 1
    End If
    
    aryTaxInfo(nIdx).m_nRecord = nRecords
    aryTaxInfo(nIdx).m_lInvoice = lInv
    aryTaxInfo(nIdx).m_nLine = nLine
    aryTaxInfo(nIdx).m_InputParms = myStruct
    ReDim aryTaxInfo(nIdx).m_Taxes(nRecords - 1)
    For i = 0 To nRecords - 1
        aryTaxInfo(nIdx).m_Taxes(i) = Taxes(i)
    Next i
    fnAddTaxInfoItem = True
End Function

'david 08/29/2003  #393678-4
'add two more optional parameters
Public Function fnChangeTaxInfoItem(ByVal lInv As Long, _
                                    ByVal nLine As Integer, _
                                    Optional sProduct As Variant, _
                                    Optional sTaxUseGroup As Variant, _
                                    Optional fDollarBase As Variant, _
                                    Optional fFreight As Variant, _
                                    Optional fGross As Variant, _
                                    Optional fNet As Variant, _
                                    Optional sBillFlag As Variant, _
                                    Optional bShowMsgBox As Boolean = True, _
                                    Optional sErrMsg As String = "", _
                                    Optional hld_vessel As Integer = -1, _
                                    Optional hld_gals As Double = 0#, _
                                    Optional hld_miles As Double = 0#, _
                                    Optional hld_tot_gals As Double = 0#) As Boolean
    
    'Ma, 3/1/99 Temporary patch for Taxes Loading Fee
    'Get the tax information array
    Dim Taxes() As CApplicableTaxes
    Dim myStruct As CMovementLineItem
    Dim nRet As Long
    Dim i As Integer
    Dim j As Integer
    Dim nIdx As Integer
    Dim nTries As Integer
    Dim nRecords As Long
    Dim nSumIdx As Integer
    
    fnChangeTaxInfoItem = False

    For nIdx = 0 To nInfoCount - 1
        If aryTaxInfo(nIdx).m_lInvoice = lInv Then
            If aryTaxInfo(nIdx).m_nLine = nLine Then
                Exit For
            End If
        End If
    Next nIdx
    If nIdx >= nInfoCount Then
        Exit Function
    End If
    With aryTaxInfo(nIdx).m_InputParms
        If Not IsMissing(sProduct) Then
            .m_sProduct = sProduct
        End If
        If Not IsMissing(sTaxUseGroup) Then
            .m_sUseGroup = sTaxUseGroup
        End If
        If Not IsMissing(fDollarBase) Then
            .m_dDollarBase = fDollarBase
        End If
        If Not IsMissing(fNet) Then
            .m_dNetUnits = fNet
        End If
        If Not IsMissing(fGross) Then
            .m_dGrossUnits = fGross
        End If
        If Not IsMissing(fFreight) Then
            .m_dFreight = fFreight
        End If
        If Not IsMissing(sBillFlag) Then
            .m_cBilledUnitsFlag = Asc(sBillFlag)
        End If
    End With
    myStruct = aryTaxInfo(nIdx).m_InputParms
    
    nTries = TAXMOD_TRIES
    'david 08/29/2003  #393678-4
    subSet_gl_sd_gals hld_vessel, hld_gals, hld_miles, hld_tot_gals, bShowMsgBox, sErrMsg
    ''''''''''''''''''''''''''''
    'Vijaya on 02/10/04 Magic#395302
    subSet_tx_date t_tax_date
    ''''''''''''''''''''''''''''''''
    
    Do
        nRet = TXMOD_VB_CalculateLineTax(App.hInstance, Forms(0).hwnd, myStruct, nRecords)
        If i = TXMOD_TASKBLOCKED Then
            DoEvents
        End If
        nTries = nTries - 1
    Loop Until nRet <> TXMOD_TASKBLOCKED Or nTries <= 0
    If nRet <> TXMOD_OK Then
        subTaxModErrors nRet, bShowMsgBox, sErrMsg
        Exit Function
    End If
        
    ReDim Taxes(nRecords - 1)
    nTries = TAXMOD_TRIES
    Do
        nRet = TXMOD_GetLineTaxes(App.hInstance, Forms(0).hwnd, Taxes())
        If i = TXMOD_TASKBLOCKED Then
            DoEvents
        End If
        nTries = nTries - 1
    Loop Until nRet <> TXMOD_TASKBLOCKED Or nTries <= 0
    
    If nRet <> TXMOD_OK Then
        subTaxModErrors nRet, bShowMsgBox, sErrMsg
        Exit Function
    End If
    
    subDealWithATaxes Taxes, nRecords
    
    aryTaxInfo(nIdx).m_nRecord = nRecords
    aryTaxInfo(nIdx).m_lInvoice = lInv
    aryTaxInfo(nIdx).m_nLine = nLine
    ReDim aryTaxInfo(nIdx).m_Taxes(nRecords - 1)
    For i = 0 To nRecords - 1
        aryTaxInfo(nIdx).m_Taxes(i) = Taxes(i)
    Next i
    
    fnChangeTaxInfoItem = True
End Function

'when Deaccumulating tax limit and basis are optional
Public Sub subAccumulateTax(ByVal lTrn As Long, _
                            ByVal fAmount As Double, _
                            Optional lLimit As Variant, _
                            Optional cBasis As Variant, _
                            Optional vLimitType As Variant)
    Dim j As Integer
    Dim bFind As Boolean
    
    bFind = False
    For j = 0 To nTaxItems - 1
        With arryTaxSums(j)
            If lTrn = .m_lTrn Then
                bFind = True
                Exit For
            End If
        End With
    Next j
    If Not bFind Then
            If IsMissing(lLimit) Then
                #If DEVELOP Then
                    MsgBox "lLimit parm is missing, set to 0"
                #End If
                lLimit = 0
            End If
            If IsMissing(cBasis) Then
                #If DEVELOP Then
                    MsgBox "cBasis parm is missing, set to ' '"
                #End If
                cBasis = " "
            End If
        If nTaxItems > nArraySize Then
            nArraySize = nArraySize + ARRAY_SIZE_INC
            ReDim Preserve arryTaxSums(nArraySize)
        End If
        With arryTaxSums(nTaxItems)
            .m_lTrn = lTrn
            If IsMissing(lLimit) Then
                .m_lLimit = 0
            Else
                .m_lLimit = val(lLimit)
            End If
            If IsMissing(lLimit) Then
                .m_cBasis = vbKeySpace
            Else
                .m_cBasis = Asc(cBasis)
            End If
            If IsMissing(vLimitType) Then
                .m_cLimitType = vbKeyI   'Default to "I"
            Else
                .m_cLimitType = Asc(vLimitType)
            End If
            .m_fAmount = tfnRound(fAmount, 2)
        End With
        nTaxItems = nTaxItems + 1
    Else
        With arryTaxSums(j)
            .m_fAmount = .m_fAmount + tfnRound(fAmount, 2)
        End With
    End If
End Sub

Public Sub subAccumulateInvoiceTax(ByVal lTrn As Long, _
                            ByVal fAmount As Double, _
                            Optional lLimit As Variant, _
                            Optional cBasis As Variant, _
                            Optional vLimitType As Variant)
    Dim j As Integer
    Dim bFind As Boolean
    
    bFind = False
    For j = 0 To nTaxItems - 1
        With arryTaxSums(j)
            If lTrn = .m_lTrn Then
                bFind = True
                Exit For
            End If
        End With
    Next j
    If Not bFind Then
            If IsMissing(lLimit) Then
                #If DEVELOP Then
                    MsgBox "lLimit parm is missing, set to 0"
                #End If
                lLimit = 0
            End If
            If IsMissing(cBasis) Then
                #If DEVELOP Then
                    MsgBox "cBasis parm is missing, set to ' '"
                #End If
                cBasis = " "
            End If
        If nTaxItems > nArraySize Then
            nArraySize = nArraySize + ARRAY_SIZE_INC
            ReDim Preserve arryTaxSums(nArraySize)
        End If
        With arryTaxSums(nTaxItems)
            .m_lTrn = lTrn
            If IsMissing(lLimit) Then
                .m_lLimit = 0
            Else
                .m_lLimit = tfnRound(lLimit, 2)
            End If
            If IsMissing(lLimit) Then
                .m_cBasis = vbKeySpace
            Else
                .m_cBasis = Asc(cBasis)
            End If
            If IsMissing(vLimitType) Then
                .m_cLimitType = vbKeyI   'Default to "I"
            Else
                .m_cLimitType = Asc(vLimitType)
            End If
            .m_fAmount = tfnRound(lLimit, 2)
        End With
        j = nTaxItems
        
        nTaxItems = nTaxItems + 1
    End If
    
    With arryTaxSums(j)
        .m_fAmount = .m_fAmount - tfnRound(fAmount, 2)
    End With
End Sub

'##############################################################################
'   subSet_gl_sd_gals:
'   Author: David Chai
'   Date: 08/29/2003
'   Project Number: 393678
'   Description:
'   To activate the SD imputed tax on freight feature
'   INPUT: carrier vessel number, gallons that are subjected to the SD imputed tax on freight
'   OUTPUT: None.
'##############################################################################
Private Sub subSet_gl_sd_gals(ByVal hld_vessel As Integer, _
                              ByVal hld_gals As Double, _
                              ByVal hld_miles As Double, _
                              ByVal hld_tot_gals As Double, _
                              Optional bShowMsgBox As Boolean = True, _
                              Optional sErrMsg As String = "")
    
    Static bAlreadyShowMsg As Boolean
    Static bAlreadyShowErr As Boolean
    
    Dim nRet As Long
    Dim nTries As Integer
    Dim i As Integer

    If hld_vessel < 0 Then
        Exit Sub
    End If
    
    Dim sTxMod32 As String
    
    On Error GoTo errTrap
    
    sTxMod32 = fnGetSystemDir(True) + "txmod32.dll"
    
    If frmContext.GetProductVersion(sTxMod32) > "3.27.1.0" Then
        'new dll
        Do
            hld_miles = round(hld_miles)
            hld_tot_gals = round(hld_tot_gals, 1)
            nRet = TXMOD_gl_sd_gals(App.hInstance, Forms(0).hwnd, hld_vessel, hld_gals, hld_miles, hld_tot_gals)
            If i = TXMOD_TASKBLOCKED Then
                DoEvents
            End If
            nTries = nTries - 1
        Loop Until nRet <> TXMOD_TASKBLOCKED Or nTries <= 0
        
        If nRet <> TXMOD_OK Then
            subTaxModErrors nRet, bShowMsgBox, sErrMsg
            Exit Sub
        End If
    Else
        'old dll
        If Not bAlreadyShowMsg Then
            bAlreadyShowMsg = True
            MsgBox "In order to use the Tax on Freight feature, the latest version of TXMOD32.DLL is required.", vbExclamation
        End If
    End If
    
    Exit Sub
    
errTrap:
    If Not bAlreadyShowErr Then
        bAlreadyShowErr = True
        
        If bShowMsgBox Then
            MsgBox "An error has occurred while calling TXMOD_gl_sd_gals() function in TXMOD32.DLL." _
                + vbCrLf + vbCrLf + "Err Code:" & Err.Number & ", Err Desc:" & Err.Description & "."
        Else
            sErrMsg = "An error has occurred while calling TXMOD_gl_sd_gals() function in TXMOD32.DLL"
        End If
    End If
End Sub
'
'Function : tfnGetSystemDir - gets the windows system directory
'Variables: optional variable to add a slash to the end of the path
'Return   : the path to the windows system directory
'
Private Function fnGetSystemDir(Optional vAddSlash As Variant) As String
    Const MAX_STRING_LENGTH As Integer = 255
    
    Dim nLength As Integer     'length returned from API call
    Dim szSystemDir As String  'temp string to hold system directory
    
    szSystemDir = Space(MAX_STRING_LENGTH) 'set the string to a fixed length for API call, pad with spaces
    nLength = GetSystemDirectory(szSystemDir, MAX_STRING_LENGTH) 'call the API function to get the system directory
    szSystemDir = Left(szSystemDir, nLength) 'trim off the excess spaces
    
    If Not IsMissing(vAddSlash) Then
        If Right(szSystemDir, 1) <> szSLASH And vAddSlash = True Then 'add a slash if it needs one
            szSystemDir = szSystemDir + szSLASH
        End If
    End If
    
    fnGetSystemDir = szSystemDir 'return system directory back to the caller
End Function


'##############################################################################
'   subSet_tx_date:
'   Author: Vijaya Alla
'   Date: 02/10/04
'   Project Number: 395302-Factor-Date Sensitive Taxes
'   Description: if tx_date is empty then use todays date
'   INPUT: t_tax_date
'   OUTPUT: None.
'##############################################################################
Private Sub subSet_tx_date(ByVal sDate As String, _
                              Optional bShowMsgBox As Boolean = True, _
                              Optional sErrMsg As String = "")
#If Not NO_DST Then
    Static bAlreadyShowMsg As Boolean
    Static bAlreadyShowErr As Boolean
    
    Dim nRet As Long
    Dim nTries As Integer
    Dim i As Integer

    Dim sTxMod32 As String
    
    On Error GoTo errTrap
        
    sTxMod32 = fnGetSystemDir(True) + "txmod32.dll"
    'if the date is empty we don't need to initialize the date
    'Taxmod32.dll consider is todays date.
    If Trim(sDate & "") = "" Then
        Exit Sub
    End If
    sDate = tfnFormatDate(sDate)
    If frmContext.GetProductVersion(sTxMod32) > "3.28.1.0" Then
        Do
            nRet = TXMOD_tx_date(App.hInstance, Forms(0).hwnd, sDate)
            If i = TXMOD_TASKBLOCKED Then
                DoEvents
            End If
            nTries = nTries - 1
        Loop Until nRet <> TXMOD_TASKBLOCKED Or nTries <= 0
        
        If nRet <> TXMOD_OK Then
            subTaxModErrors nRet, bShowMsgBox, sErrMsg
            Exit Sub
        End If
    Else
        'old dll
        If Not bAlreadyShowMsg Then
            bAlreadyShowMsg = True
            MsgBox "In order to use the Date Sensitive Taxes feature, the latest version of TXMOD32.DLL is required.", vbExclamation
        End If
    End If
    
    Exit Sub
    
errTrap:
    If Not bAlreadyShowErr Then
        bAlreadyShowErr = True
        
        If bShowMsgBox Then
            MsgBox "An error has occurred while calling TXMOD_tx_date() function in TXMOD32.DLL." _
                + vbCrLf + vbCrLf + "Err Code:" & Err.Number & ", Err Desc:" & Err.Description & "."
        Else
            sErrMsg = "An error has occurred while calling TXMOD_tx_date() function in TXMOD32.DLL"
        End If
    End If
#Else
    'Do nothing to avoid date sensitivity
#End If

End Sub

'#411523wj042505
Private Function fnVersionOK(ByVal sFileName As String, ByVal sVerToChk As String) As Boolean
    Dim vVer As Variant
    Dim i As Integer
    Static sVersion As String
    
    On Error GoTo errExit
    If sVersion = "" Then
        sVersion = frmContext.GetProductVersion(sFileName)
        vVer = Split(sVersion, ".")
        For i = 0 To UBound(vVer)
            sVersion = Format(vVer(i), "00") & "."
        Next
    End If
    
    fnVersionOK = (Mid(sVersion, 1, 11) > Mid(sVerToChk, 1, 11))
    Exit Function
errExit:
    fnVersionOK = False
End Function
'#411523wj042505
Private Sub subSet_AccmTxbleAmt(ByVal dAccmTxbleAmt As Double, _
                              Optional bShowMsgBox As Boolean = True, _
                              Optional sErrMsg As String = "")
    Static bAlreadyShowMsg As Boolean
    Static bAlreadyShowErr As Boolean
    
    Dim nRet As Long
    Dim nTries As Integer
    Dim i As Integer
    Dim sTxMod32 As String
    
    On Error GoTo errTrap
        
    If Abs(dAccmTxbleAmt) < 0.01 Then
        '#Do not need to reset it. The variable get cleared
        Exit Sub
    End If
    sTxMod32 = fnGetSystemDir(True) + "txmod32.dll"
    
    If fnVersionOK(sTxMod32, "03.28.10.04") Then
        Do
            nRet = TXMOD_set_accm_txble_amt(App.hInstance, Forms(0).hwnd, dAccmTxbleAmt)
            If i = TXMOD_TASKBLOCKED Then
                DoEvents
            End If
            nTries = nTries - 1
        Loop Until nRet <> TXMOD_TASKBLOCKED Or nTries <= 0
        
        If nRet <> TXMOD_OK Then
            subTaxModErrors nRet, bShowMsgBox, sErrMsg
            Exit Sub
        End If
    Else
        'old dll
        If Not bAlreadyShowMsg Then
            bAlreadyShowMsg = True
            'david 08/11/2005  per Tom, we don't need to show this message
            'MsgBox "In order to use Single Article Taxes, the latest version of TXMOD32.DLL is required.", vbExclamation
        End If
    End If
    
    Exit Sub
    
errTrap:
    If Not bAlreadyShowErr Then
        bAlreadyShowErr = True
        
        If bShowMsgBox Then
            MsgBox "An error has occurred while calling subSet_AccmTxbleAmt() function in TXMOD32.DLL." _
                + vbCrLf + vbCrLf + "Err Code:" & Err.Number & ", Err Desc:" & Err.Description & "."
        Else
            sErrMsg = "An error has occurred while calling subSet_AccmTxbleAmt() function in TXMOD32.DLL"
        End If
    End If
End Sub
''''''''''''''''''''''''''''''''
'#------------------------------------------------------------
'#"Start ATax Region" - 452155wj101304
'# We intend to fix the Loading Fees display issue. The temp patch made
'# in 1999 did not reflect the real Atax charged in the Back Ends. The backends
'# routine starts with BOL side loading Fees and prorate based delivery. Then
'# monkey with the numbers ... We now will follow the same logic as 4GL.
'# The Next public function is the only entry point of this new ATax routine.
'#(1)This routine will be used only by Fuel Movement and Its Editor
'#   Thus, m_Disable_A_Taxes must be false
'#(2)Call this sub at the end of "Delivery and Lading OK click" code
'#(3)This sub is intended to replace all the Temp A Tax fixes atarted from 1999
'#   For now, we still keep the codes in place. When field cTaxBasis (local
'#   MTaxDetail) is in every user's PC, we can remove the "temp patch".
Public Sub txMod_ReBuildATax_MTaxDetail(m_local As Database, lMovement As Long)
    Const SUB_NAME = "txMod_ReBuildATax_MTaxDetail"
    Dim strSQL As String
    Dim rsTemp As Recordset
    Dim rsLadHeader As Recordset
    Dim fTaxes As Double
    Dim fBase As Double
    Dim sNGU As String
    Dim nLine As Integer
    Dim lATaxGal As Long
    
    Dim aryATax() As BOL_AtAx
    Dim BOL_atx_ptr As Integer
    
    Dim rsLadD As Recordset
    Dim rsMvBal As Recordset
    Dim rsDelD As Recordset
    Dim lTotATaxGal As Long
         
    Dim bTexasTax As Boolean
    Dim bTexasDiverted As Boolean
    Dim bFound As Boolean
    Dim sum_atax As Double
    Dim this_tax As Double
    Dim i As Integer
    Dim k As Integer
    Dim this_units As Long
    Dim dsl_fuel_group As String
    
    Dim sTaxCode As String
    Dim sTaxDesc As String
    Dim sTaxState As String
    Dim fTaxRate As Double
    Dim delta As Double
    
    Dim lLadGal As Long
    Dim lDelGal As Long
    
    Dim sIEFlag As String
    Dim bBOL_External As Boolean
    Dim lDest As Long
    
    '#Make sure it is FM or Its Editor
    If m_Disable_A_Taxes Then
        Exit Sub
    End If
    
    '#Test if "cTaxBasis" is a column in local table
    '#MTaxDetail. If not, this routine will exist. In other
    '#words, the new routine needes this field. This is why
    '#we did not remove the old "temp patch " codes -- does not
    '#hurt if new routine is used
    strSQL = "SELECT cTaxBasis FROM MTaxDetail WHERE lMovement = " & lMovement 'Vijaya on 01/28/05 Magic#458098
    Set rsTemp = txMod_OpenRecord(strSQL, SUB_NAME, , m_local)
    If rsTemp Is Nothing Then
         Exit Sub
    End If
    
    '#lLading May be -1. Skip this routine if -1 since it is not ready to do loading fee
    strSQL = "SELECT * FROM MBalance WHERE lMovement = " & lMovement & " AND lLading <= 0 "
    Set rsMvBal = txMod_OpenRecord(strSQL, SUB_NAME, , m_local)
    If rsMvBal Is Nothing Then
         Exit Sub
    ElseIf rsMvBal.RecordCount > 0 Then
         Exit Sub
    End If
    
    '#If not balanced, do not need to do it
    strSQL = "SELECT SUM(lGrossGals) AS GAL FROM MLading WHERE lMovement = " & lMovement
    Set rsTemp = txMod_OpenRecord(strSQL, , , m_local)
    If rsTemp Is Nothing Then
        Exit Sub
    ElseIf rsTemp.RecordCount = 0 Then
        Exit Sub
    End If
    lLadGal = tfnRound(rsTemp!gal)
    strSQL = "SELECT SUM(lGrossGals) AS GAL FROM MDelivery WHERE lMovement = " & lMovement
    Set rsTemp = txMod_OpenRecord(strSQL, SUB_NAME, , m_local)
    If rsTemp Is Nothing Then
        Exit Sub
    ElseIf rsTemp.RecordCount = 0 Then
        Exit Sub
    End If
    
    lDelGal = tfnRound(rsTemp!gal)
    'david 02/04/2005  #470421
    'commented out the Or ...
    'because if the user zero out the delivery gallon
    'we need to set the TX loading fee to 0
    If lLadGal <> lDelGal Then  'Or lDelGal = 0 Then
        Exit Sub
    End If
    
    sTaxState = "TX" '#Defalt to Texas
    
    strSQL = "SELECT * FROM MBalance WHERE lMovement = " & lMovement
    Set rsMvBal = txMod_OpenRecord(strSQL, SUB_NAME, , m_local)
    If rsMvBal Is Nothing Then
         Exit Sub
    ElseIf rsMvBal.RecordCount = 0 Then
         Exit Sub
    End If

    strSQL = "SELECT * FROM MLading WHERE lMovement = " & lMovement
    Set rsLadHeader = txMod_OpenRecord(strSQL, SUB_NAME, , m_local)
    If rsLadHeader Is Nothing Then
        Exit Sub
    ElseIf rsLadHeader.RecordCount = 0 Then
        Exit Sub
    End If
    
    strSQL = "SELECT MDeliveryDetail.*,MDelivery.lDestination,MDelivery.bExternal FROM MDelivery,MDeliveryDetail " _
             & " WHERE MDelivery.lMovement=MDeliveryDetail.lMovement " _
             & " AND MDelivery.lDelivery=MDeliveryDetail.lDelivery " _
             & " AND MDelivery.lMovement = " & lMovement
             
    Set rsDelD = txMod_OpenRecord(strSQL, SUB_NAME, , m_local)
    If rsDelD Is Nothing Then
         Exit Sub
    ElseIf rsDelD.RecordCount = 0 Then
         Exit Sub
    End If
    
    '#Build Lading side A-Taxes array --> aryATax
    BOL_atx_ptr = -1
    While Not rsLadHeader.EOF
        strSQL = "SELECT * " _
               & " FROM MLadingDetail " _
               & " WHERE lLading = " & CStr(rsLadHeader!lLading) _
               & " AND lMovement = " & lMovement _
               & " ORDER BY lLineNumber"
               
        Set rsLadD = txMod_OpenRecord(strSQL, SUB_NAME, , m_local)
        If rsLadD Is Nothing Then
            Exit Sub
        End If
        subResetTaxes True, True
        lTotATaxGal = 0
        nLine = 0
        
        strSQL = "SELECT dPickupDate FROM mLading"
        strSQL = strSQL & " WHERE lMovement=" & lMovement
        strSQL = strSQL & " AND lLading=" & tfnRound(rsLadHeader!lLading)
        Set rsTemp = txMod_OpenRecord(strSQL, SUB_NAME, , m_local)
        If rsTemp Is Nothing Then
            Exit Sub
        End If
        If rsTemp.RecordCount > 0 Then
            t_tax_date = rsTemp!dPickupDate & ""
        Else
            t_tax_date = Date
        End If
        With rsLadD
            If .RecordCount > 0 Then
                While Not .EOF
                    
                    strSQL = "SELECT SUM(MBalance.lGross) AS G_GAL, SUM(MBalance.lNet) AS N_GAL,MDelivery.lDestination"
                    strSQL = strSQL & " FROM MBalance, MDelivery WHERE MBalance.lDelivery= MDelivery.lDelivery"
                    strSQL = strSQL & " AND MBalance.lLading =" & rsLadD!lLading
                    strSQL = strSQL & " AND MBalance.lLadingLine =" & rsLadD!lLineNumber
                    strSQL = strSQL & " AND MBalance.lMovement =" & lMovement
                    strSQL = strSQL & " GROUP BY MDelivery.lDestination"
                    Set rsTemp = txMod_OpenRecord(strSQL, SUB_NAME, , m_local)
                    If rsTemp Is Nothing Then
                        Exit Sub
                    End If
                            
                    Do While Not rsTemp.EOF
                        If !bGross Then
                            sNGU = "G"
                            fBase = rsTemp!G_GAL * !fCost
                            lTotATaxGal = lTotATaxGal + rsTemp!G_GAL
                        Else
                            sNGU = "N"
                            fBase = rsTemp!N_GAL * !fCost
                            lTotATaxGal = lTotATaxGal + rsTemp!N_GAL
                        End If
                        lDest = tfnRound(rsTemp!lDestination)
                        If Not fnATaxInfoItem(rsLadHeader!bExternal, rsLadHeader!lTerminal, lDest, rsLadHeader!lLading, nLine, _
                                !szProductCode, !szTaxUseGroup, _
                                fBase, 0, _
                                CDbl(rsTemp!G_GAL), CDbl(rsTemp!N_GAL), _
                                sNGU) Then
                            Exit Sub
                        End If
                        nLine = nLine + 1
                        rsTemp.MoveNext
                    Loop
                    .MoveNext
                Wend '#Lading D
                subProrateBOLATax rsLadHeader!lLading, _
                                   nLine, _
                                   bTexasTax, sTaxState
                
                fTaxes = fnGetTypeATax
                .MoveFirst
                While Not .EOF
                    If !bGross Then
                        sNGU = "G"
                        lATaxGal = !lGross
                    Else
                        sNGU = "N"
                        lATaxGal = !lNet
                    End If
                    BOL_atx_ptr = BOL_atx_ptr + 1
                    ReDim Preserve aryATax(BOL_atx_ptr) As BOL_AtAx
                    aryATax(BOL_atx_ptr).fmld_bl_nbr = !lLading
                    aryATax(BOL_atx_ptr).fmld_seq_nbr = !lLineNumber
                    aryATax(BOL_atx_ptr).sum_bunits = lATaxGal
                    aryATax(BOL_atx_ptr).bl_ngu = sNGU
                    If lTotATaxGal = 0 Then
                        aryATax(BOL_atx_ptr).a_tax_tot = 0
                    Else
                        aryATax(BOL_atx_ptr).a_tax_tot = fTaxes * lATaxGal / lTotATaxGal
                    End If
                    aryATax(BOL_atx_ptr).a_tax_used = 0
                    .MoveNext
                Wend '#Lading D
            End If
        End With
        rsLadHeader.MoveNext
    Wend '#Lading
    If BOL_atx_ptr < 0 Then
        Exit Sub
    End If
    
    '#(2) Build Loading Fees in Delivery Side --> put_ldfee_rec
    
    start_ldfee_rec
    Do While Not rsDelD.EOF
        'sum_bunits = 0
        sum_atax = 0
        rsMvBal.MoveFirst
        Do While Not rsMvBal.EOF
            this_tax = 0
            If rsMvBal!lDelivery = rsDelD!lDelivery And rsMvBal!lDeliveryLine = rsDelD!lLineNumber Then
                bFound = False
                For i = 0 To BOL_atx_ptr
                    If aryATax(i).fmld_bl_nbr = rsMvBal!lLading And aryATax(i).fmld_seq_nbr = rsMvBal!lLadingLine Then
                        bFound = True
                        Exit For
                    End If
                Next i
                this_tax = 0
                this_units = 0
                If bFound Then
                    sNGU = aryATax(i).bl_ngu
                    If (sNGU = "N") Then
                        this_units = rsMvBal!lNet
                    Else
                        this_units = rsMvBal!lGross
                    End If
                    If aryATax(i).sum_bunits <> 0 Then
                        this_tax = (Abs(this_units) / Abs(aryATax(i).sum_bunits)) * Abs(aryATax(i).a_tax_tot)
                    End If
                    this_tax = tfnRound(this_tax, 2)
                Else
                    '#Not enough information to cal A Tax.
                    Exit Do
                End If
                  
                this_tax = tfnRound(Abs(this_tax), 2)
                delta = Abs(aryATax(i).a_tax_tot) - (this_tax + Abs(aryATax(i).a_tax_used))
                delta = tfnRound(Abs(delta), 2)
                If (delta < 0.05) Then
                    this_tax = Abs(aryATax(i).a_tax_tot) - Abs(aryATax(i).a_tax_used)
                End If
                
                If Not bTexasTax Then
                     '#determine if Internal or extLad
                     bBOL_External = True
                     rsLadHeader.MoveFirst
                     Do While Not rsLadHeader.EOF
                        If rsLadHeader!lLading = rsMvBal!lLading Then
                            bBOL_External = rsLadHeader!bExternal
                            Exit Do
                        End If
                        rsLadHeader.MoveNext
                     Loop
                     If bBOL_External = False Then
                         bTexasTax = True
                         bTexasDiverted = False
                     ElseIf rsDelD!bExternal = True Then
                         strSQL = "SELECT * FROM ar_altname WHERE an_customer =" & tfnRound(rsDelD!lDestination)
                         Set rsTemp = txMod_OpenRecord(strSQL)
                         If Not rsTemp Is Nothing Then
                            If rsTemp.RecordCount > 0 Then
                                strSQL = "SELECT * FROM sys_city WHERE city_code = " & tfnSQLString(rsTemp!an_city)
                                Set rsTemp = txMod_OpenRecord(strSQL)
                                If Not rsTemp Is Nothing Then
                                    If rsTemp.RecordCount > 0 Then
                                        sTaxState = Trim(rsTemp!city_state & "")
                                        If sTaxState = "TX" Then
                                            bTexasDiverted = True
                                        End If
                                    End If
                                End If
                            End If
                         End If
                    End If
               End If
                
                'david 03/23/2005  #478014
                'fixed RTE 9 subscript out of range
                sum_atax = sum_atax + tfnRound(this_tax, 2)
                aryATax(i).a_tax_used = aryATax(i).a_tax_used + this_tax
            End If
            
            rsMvBal.MoveNext
        Loop
        If rsDelD!bExternal Then
            sIEFlag = "E"
        Else
            sIEFlag = "I"
        End If
        put_ldfee_rec rsDelD!lDelivery, _
                          rsDelD!lLineNumber, _
                          rsDelD!lNet, _
                          sum_atax, _
                          rsDelD!lProductLink, _
                          sIEFlag
                          
        rsDelD.MoveNext
    Loop
    
    If bTexasTax Or bTexasDiverted Then
        dsl_fuel_group = sSysParm1055
    Else
        dsl_fuel_group = "NONE"
    End If
    monkey_the_numbers Trim(dsl_fuel_group), bTexasDiverted
    
    strSQL = "DELETE FROM MTaxDetail WHERE lMovement = " & lMovement & " AND cTaxBasis = 'A'"
    txMod_ExeCuteSQL strSQL, , , m_local
    
    For i = 0 To ldf_cnt
        '#We do not need to populate Internal Loading Fees
        'david 02/04/2005  #470421
        'modify the IF statement below to insert zero loading fee
        'and is to be used to reset the fSalesTax in invoice header
        'and the fTotalTax in invoice line to zero
        'when the delivery gallons is zero and the loading fee is zero
        'ZERO loading fee records will be deleted later...
        If aryLDF(i).dv_nbr <> 0 And aryLDF(i).dv_IE = "E" _
           And (tfnRound(aryLDF(i).dv_atax, 2) <> 0 _
           Or tfnRound(aryLDF(i).dv_vol, 1) = 0) Then
            '#Try to Get the same desc as 4GL does.
            If fnGetATaxCodeDesc(rsDelD, aryLDF(i).dv_nbr, _
                                 aryLDF(i).dv_vol, _
                                 sTaxCode, sTaxDesc, fTaxRate) Then
                If sTaxDesc = "**Loading Fees**" Then
                    '#Exceeds limit
                    sTaxDesc = "Loading Fees"
                    fTaxRate = 0
                    For k = 0 To ldf_cnt
                        If aryLDF(k).dv_nbr = aryLDF(i).dv_nbr And aryLDF(i).dv_IE = "E" Then
                            fTaxRate = fTaxRate + aryLDF(k).dv_atax
                        End If
                    Next
                    fTaxRate = tfnRound(Abs(fTaxRate))
                End If
                
                If aryLDF(i).dv_vol * aryLDF(i).dv_atax < 0 Then
                    aryLDF(i).dv_atax = aryLDF(i).dv_atax * (-1)
                End If
                strSQL = "INSERT INTO MTaxDetail (lMovement,lDelivery,lLineNumber,szTaxCode," _
                         & "szTaxDescription,fBaseAmount,fTaxRate,fTaxAmount,cTaxBasis) VALUES ("
                strSQL = strSQL & lMovement & ","
                strSQL = strSQL & CStr(aryLDF(i).dv_nbr) & ","
                strSQL = strSQL & CStr(aryLDF(i).dv_line) & ","
                strSQL = strSQL & tfnSQLString(sTaxCode) & ","
                strSQL = strSQL & tfnSQLString(sTaxDesc) & ","
                strSQL = strSQL & aryLDF(i).dv_vol & ","
                strSQL = strSQL & CStr(fTaxRate) & ","
                strSQL = strSQL & tfnRound(aryLDF(i).dv_atax, 2) & ",'A')"
                
                txMod_ExeCuteSQL strSQL, , , m_local
            End If
        End If
    Next
    
    'Adjusting Tax Amount in MInvoice local table
    'Vijaya on 12/06/04 Magic#465810
    'Update the Field in MInvoice.fTotalTax (Local)
    strSQL = "SELECT lMovement,lDelivery,SUM(fTaxAmount) as dTxAmt FROM MTaxDetail"
    strSQL = strSQL & " WHERE lMovement =" & lMovement 'Vijaya on 01/28/05 Magic#458098
    strSQL = strSQL & " GROUP BY lMovement,lDelivery"
    strSQL = strSQL & " ORDER BY lMovement,lDelivery"
    Set rsTemp = txMod_OpenRecord(strSQL, SUB_NAME, , m_local)
    If rsTemp Is Nothing Then
        Exit Sub
    End If
    If rsTemp.RecordCount > 0 Then
        While Not rsTemp.EOF
            strSQL = "UPDATE MInvoice SET fTotalTax =" & tfnRound(rsTemp!dTxAmt, 2)
            strSQL = strSQL & " WHERE lMovement =" & tfnRound(rsTemp!lMovement)
            strSQL = strSQL & " AND lDelivery =" & tfnRound(rsTemp!lDelivery)
            txMod_ExeCuteSQL strSQL, , , m_local
            rsTemp.MoveNext
        Wend
    End If
    
    'Adjusting Tax Amount in MInvoiceLine local table
    'Update the Field in MInvoiceLine.fSalesTax (Local)
    strSQL = "SELECT lMovement,lDelivery,lLineNumber,SUM(fTaxAmount) as dTxAmt"
    strSQL = strSQL & " FROM MTaxDetail"
    strSQL = strSQL & " WHERE lMovement =" & lMovement 'Vijaya on 01/28/05 Magic#458098
    strSQL = strSQL & " GROUP BY lMovement,lDelivery,lLineNumber"
    strSQL = strSQL & " ORDER BY lMovement,lDelivery,lLineNumber"
    Set rsTemp = txMod_OpenRecord(strSQL, SUB_NAME, , m_local)
    If rsTemp Is Nothing Then
        Exit Sub
    End If
    If rsTemp.RecordCount > 0 Then
        While Not rsTemp.EOF
            strSQL = "UPDATE MInvoiceLine SET fSalesTax =" & tfnRound(rsTemp!dTxAmt, 2)
            strSQL = strSQL & " WHERE lMovement =" & tfnRound(rsTemp!lMovement)
            strSQL = strSQL & " AND lDelivery =" & tfnRound(rsTemp!lDelivery)
            strSQL = strSQL & " AND lLineNumber =" & tfnRound(rsTemp!lLineNumber)
            txMod_ExeCuteSQL strSQL, , , m_local
            rsTemp.MoveNext
        Wend
    End If
    ''''''''''''''''''''''''''''''''

    'david 02/04/2005  #470421
    'delete the ZERO loading fee records
    strSQL = "DELETE FROM MTaxDetail"
    strSQL = strSQL & " WHERE lMovement =" & lMovement
    strSQL = strSQL & " AND fBaseAmount = 0"
    strSQL = strSQL & " AND fTaxAmount = 0"
    strSQL = strSQL & " AND cTaxBasis = 'A'"
    txMod_ExeCuteSQL strSQL, , , m_local
    ''''''''''''''''''''''''''
End Sub

Private Function fnGetATaxCodeDesc(ByVal rsDD As Recordset, _
                               ByVal lDelNbr As Long, _
                               ByVal lGal As Long, ByRef sTxCode As String, _
                               ByRef sTaxDesc As String, ByRef fTaxRate As Double) As Boolean
    
    Const SUB_NAME As String = "subGetATaxCodeDesc"
    
    Dim strSQL As String
    Dim rsTemp As Recordset
    Dim rsTrnCode As Recordset
    Dim sTrnCode As String
    
    Dim i As Integer
    Dim j As Integer
    
    fnGetATaxCodeDesc = False
    
    sTxCode = "LDF"
    sTaxDesc = "**Loading Fees**"
    rsDD.MoveFirst
    sTrnCode = szEMPTY
    Do While Not rsDD.EOF
        If lDelNbr = rsDD!lDelivery Then
            strSQL = "SELECT tul_trn FROM tx_use_g,tx_use_link,inv_header"
            strSQL = strSQL & " WHERE ivh_link =" & tfnRound(rsDD!lProductLink)
            strSQL = strSQL & " AND tu_use_g = " & tfnSQLString(rsDD!szTaxUseGroup)
            strSQL = strSQL & " AND tu_link = tul_link"
            strSQL = strSQL & " AND tu_prodtcl = ivh_prodtcl"
            Set rsTrnCode = txMod_OpenRecord(strSQL, SUB_NAME)
            If rsTrnCode Is Nothing Then
                Exit Function
            End If
            If rsTrnCode.RecordCount > 0 Then
                While Not rsTrnCode.EOF
                    If sTrnCode = "" Then
                        sTrnCode = tfnRound(rsTrnCode!tul_trn)
                    Else
                        sTrnCode = sTrnCode & "," & tfnRound(rsTrnCode!tul_trn)
                    End If
                    rsTrnCode.MoveNext
                Wend
            End If
        End If
        rsDD.MoveNext
    Loop
    If sTrnCode = szEMPTY Then
        Exit Function
    End If
    If lGal > 0 Then
        lGal = lGal * (-1)
    End If
    strSQL = "SELECT tt_code,tt_descr,tt_rate,tt_limit FROM tx_table"
    strSQL = strSQL & " WHERE tt_level = 'S' "
    strSQL = strSQL & " AND tt_charge = 'Y' AND tt_basis = 'A'"
    strSQL = strSQL & " AND tt_trn IN (" & sTrnCode & ")"
    strSQL = strSQL & " AND tt_limit < " & CStr(lGal)
    strSQL = strSQL & " ORDER BY tt_limit DESC"
    strSQL = tfnFix_tx_table(strSQL)
    
    Set rsTemp = txMod_OpenRecord(strSQL)
    If rsTemp Is Nothing Then
        Exit Function
    ElseIf rsTemp.RecordCount = 0 Then
        Exit Function
    End If
    sTxCode = rsTemp!tt_code & ""
    sTaxDesc = rsTemp!tt_descr & ""
    fTaxRate = tfnRound(rsTemp!tt_rate, 2)
    
    fnGetATaxCodeDesc = True
End Function

Private Sub start_ldfee_rec()
    ldf_cnt = -1
    ReDim aryLDF(0)
End Sub
Private Sub put_ldfee_rec(ByVal dv_nbr As Long, _
                          ByVal dv_line As Long, _
                          ByVal bunits As Long, _
                          ByVal atax As Double, _
                          ByVal ProdLnk As Long, _
                          ByVal sIE As String)
    Dim rsTemp As Recordset
    Dim sClass As String
    Dim strSQL As String
    
    strSQL = "SELECT * FROM inv_header WHERE ivh_link = " & ProdLnk
    Set rsTemp = txMod_OpenRecord(strSQL)
    If rsTemp Is Nothing Then
        sClass = "*****"
    ElseIf rsTemp.RecordCount = 0 Then
        sClass = "*****"
    Else
        sClass = Trim(rsTemp!ivh_prodtcl & "")
    End If
    ldf_cnt = ldf_cnt + 1
    ReDim Preserve aryLDF(ldf_cnt) As Loading_Fees
    
    aryLDF(ldf_cnt).dv_nbr = dv_nbr
    aryLDF(ldf_cnt).dv_line = dv_line
    aryLDF(ldf_cnt).dv_prodlnk = ProdLnk
    aryLDF(ldf_cnt).dv_prodtcl = sClass
    aryLDF(ldf_cnt).dv_vol = bunits
    aryLDF(ldf_cnt).dv_atax = atax
    aryLDF(ldf_cnt).dv_IE = sIE

End Sub

'#Monkey around just like 4GL does so that we can match the
'#T-Tax 4GL gives
Private Sub monkey_the_numbers(ByVal dsl_fuel_group As String, _
                               ByVal texas_diverted As Boolean)
    Dim atax_tot        As Double
    Dim atax_new        As Double
    Dim atax_inv        As Double
    Dim atax_wrk        As Double
    Dim wrk_nbr         As Long
    Dim gal_inv         As Double
    Dim gal_wrk         As Double
    Dim x               As Integer
    Dim y               As Integer
    'david 08/04/2005  #489634
    Dim txlf1           As Double
    Dim txlf2           As Double
    Dim txlf3           As Double
    Dim txlf4           As Double
    Dim txlf_increment  As Double
    ''''''''''''''''''''''''''

    '# New variables to loads of 10,000+
    Dim low_end         As Long
    Dim high_end        As Long
    Dim incr_qty        As Long
    Dim incr_amt        As Double

    If ldf_cnt < 0 Then
        Exit Sub
    End If
    '# Return if not Texas routine
    If dsl_fuel_group = "NONE" Then
        Exit Sub
    End If
     
    Dim dsl_gals As Long
    Dim gas_gals As Long
    
    On Error GoTo Stop_Monkey_Around
    
    dsl_gals = 0
    gas_gals = 0
    
    For x = 0 To ldf_cnt
        If Trim(aryLDF(x).dv_prodtcl) = dsl_fuel_group Then
            dsl_gals = dsl_gals + aryLDF(x).dv_vol
        Else
            gas_gals = gas_gals + aryLDF(x).dv_vol
        End If
    Next
    '# Return if LPG only
    If gas_gals = 0 And dsl_gals = 0 Then
        Exit Sub
    End If

    '# Add all of the 'A' tax amounts
    atax_tot = 0
    For x = 0 To ldf_cnt
        atax_tot = atax_tot + tfnRound(aryLDF(x).dv_atax, 2)
    Next
    
    If tfnRound(atax_tot, 2) = 0 And texas_diverted = False Then
        Exit Sub
    End If

    '# Calculate Texas loading fees without prorating
    wrk_nbr = 0
    atax_new = 0
    atax_inv = 0
    gal_inv = 0
    gas_gals = 0
    dsl_gals = 0
    x = -1
    'david 08/04/2005  #489634
    txlf1 = 0#
    txlf2 = 0#
    txlf3 = 0#
    txlf4 = 0#
    txlf_increment = 0#
    
    If t_tax_date < CDate("9/1/2005") Then
        txlf1 = 10#
        txlf2 = 20#
        txlf3 = 30#
        txlf4 = 40#
        txlf_increment = 20#
    ElseIf t_tax_date >= CDate("9/1/2005") And t_tax_date < CDate("9/1/2006") Then
        txlf1 = 5#
        txlf2 = 10#
        txlf3 = 15#
        txlf4 = 20#
        txlf_increment = 10#
    ElseIf t_tax_date >= CDate("9/1/2006") And t_tax_date < CDate("9/1/2007") Then
        txlf1 = 2#
        txlf2 = 4#
        txlf3 = 6#
        txlf4 = 8#
        txlf_increment = 4#
    Else  'if t_tax_date >= CDate("9/1/2007")
        'set to zero
    End If
    ''''''''''''''''''''''''''
    
    ReDim Preserve aryLDF(ldf_cnt + 10)
    
    Do While True
        If tfnRound(atax_tot, 2) = 0 And texas_diverted Then
            Exit Do
        End If

        x = x + 1
        If x > ldf_cnt Then
            aryLDF(x).dv_nbr = 0
        End If
        If wrk_nbr = 0 Then
            wrk_nbr = aryLDF(x).dv_nbr
        End If
        If wrk_nbr <> aryLDF(x).dv_nbr Then
            gal_wrk = Abs(gal_inv)
            If Abs(gas_gals) >= 7000 And Abs(gas_gals) < 8000 And Abs(gal_inv) < 10000 Then
                gal_wrk = 8000
            End If
            If gal_wrk < 2500 Then
                atax_inv = txlf1
            End If
            If gal_wrk >= 2500 And gal_wrk < 5000 Then
                atax_inv = txlf2
            End If
            If gal_wrk >= 5000 And gal_wrk < 8000 Then
                atax_inv = txlf3
            End If
            If gal_wrk >= 8000 And gal_wrk < 10000 Then
                atax_inv = txlf4
            End If
            If gal_wrk >= 10000 Then
                low_end = 10000
                incr_qty = 5000
                incr_amt = txlf_increment
                atax_inv = txlf4
                Do While True
                    high_end = low_end + incr_qty
                    atax_inv = atax_inv + incr_amt
                    If gal_wrk >= low_end And gal_wrk < high_end Then
                        Exit Do
                    End If
                    low_end = high_end
                Loop
            End If
            atax_new = atax_new + atax_inv
            wrk_nbr = aryLDF(x).dv_nbr
            If wrk_nbr = 0 Then
                Exit Do
            End If
            gal_inv = 0
            gas_gals = 0
            dsl_gals = 0
        End If
        gal_inv = gal_inv + aryLDF(x).dv_vol
        If aryLDF(x).dv_prodtcl = dsl_fuel_group Then
            dsl_gals = dsl_gals + aryLDF(x).dv_vol
        Else
            gas_gals = gas_gals + aryLDF(x).dv_vol
        End If
    Loop

    '# Screw with the numbers only if new $ <= old $!
    If atax_new > atax_tot Then
        Exit Sub
    End If

    '# So, do it again, but this time save the numbers!
    wrk_nbr = 0
    atax_inv = 0
    gal_inv = 0
    gas_gals = 0
    dsl_gals = 0
    x = -1

    Do While True
        x = x + 1
        If x > ldf_cnt Then
            aryLDF(x).dv_nbr = 0
        End If
        If wrk_nbr = 0 Then
            wrk_nbr = aryLDF(x).dv_nbr
        End If
        If wrk_nbr <> aryLDF(x).dv_nbr Then
            gal_wrk = Abs(gal_inv)
            If Abs(gas_gals) >= 7000 And Abs(gas_gals) < 8000 And Abs(gal_inv) < 10000 Then
                gal_wrk = 8000
            End If
            If gal_wrk < 2500 Then
                atax_inv = txlf1
            End If
            If gal_wrk >= 2500 And gal_wrk < 5000 Then
                atax_inv = txlf2
            End If
            If gal_wrk >= 5000 And gal_wrk < 8000 Then
                atax_inv = txlf3
            End If
            If gal_wrk >= 8000 And gal_wrk < 10000 Then
                atax_inv = txlf4
            End If
            If gal_wrk >= 10000 Then
                low_end = 10000
                incr_qty = 5000
                incr_amt = txlf_increment
                atax_inv = txlf4
                Do While True
                    high_end = low_end + incr_qty
                    atax_inv = atax_inv + incr_amt
                    If gal_wrk >= low_end And gal_wrk < high_end Then
                        Exit Do
                    End If
                    low_end = high_end
                Loop
            End If
            atax_new = 0
            If ldf_cnt > 0 Then
                For y = 0 To ldf_cnt - 1
                    If aryLDF(y).dv_nbr = wrk_nbr Then
                        If gal_inv <> 0 Then
                            atax_wrk = (aryLDF(y).dv_vol / gal_inv) * atax_inv
                        Else
                            atax_wrk = 0
                        End If
                        atax_wrk = tfnRound(atax_wrk, 2)
                        aryLDF(y).dv_atax = atax_wrk
                        atax_new = atax_new + atax_wrk
                    End If
                Next
            End If
            atax_wrk = atax_inv - atax_new
            aryLDF(ldf_cnt).dv_atax = tfnRound(atax_wrk, 2)
            wrk_nbr = aryLDF(x).dv_nbr
            If wrk_nbr = 0 Then
                Exit Do
            End If
            gal_inv = 0
            gas_gals = 0
            dsl_gals = 0
        End If
        gal_inv = gal_inv + aryLDF(x).dv_vol
        If aryLDF(x).dv_prodtcl = dsl_fuel_group Then
            dsl_gals = dsl_gals + aryLDF(x).dv_vol
        Else
            gas_gals = gas_gals + aryLDF(x).dv_vol
        End If
    Loop
    Exit Sub
Stop_Monkey_Around:
    MsgBox "An Error Occurred In Calculating Loading Fees. Error Description: " & Err.Description
End Sub

'#based on the Total Gallons of the BOL, find the A Tax
'#Then prorate it
Private Sub subProrateBOLATax(ByVal lBOL As Long, _
                              ByVal nLadDCnt As Integer, _
                              ByRef Texas_Tax As Boolean, _
                              ByRef sTaxState As String)
    Dim i As Integer
    Dim j As Integer
    Dim nLadDIdx As Integer
    Dim lFuelAmount As Long
    Dim lTaxableAmount As Long
    Dim lGasAmount As Long
    Dim sParm1055 As String
    Dim fTaxes As Double
    Dim fTaxRate As Double
    Dim sTaxGroup As String
    Dim nIdx As Integer
    Dim lUnits() As Long
    Dim fTypeATaxes() As Double
    
    Dim nSign As Integer
    
    If nInfoCount <= 0 Or nLadDCnt <= 0 Then
        Exit Sub
    End If
    ReDim fTypeATaxes(nLadDCnt - 1) As Double
    ReDim lUnits(nLadDCnt - 1) As Long
    
    sParm1055 = "NONE"
    
    For nLadDIdx = 0 To nLadDCnt - 1
        nTaxItems = 0
        lFuelAmount = 0
        lGasAmount = 0
        For nIdx = 0 To nInfoCount - 1
            With aryTaxInfo(nIdx)
                If .m_lInvoice = lBOL Then
                    For i = 0 To .m_nRecord - 1
                        If .m_Taxes(i).m_cCharge = vbKeyY Or .m_Taxes(i).m_cCharge = vbKeyC Then
                            If .m_Taxes(i).m_cBasis = vbKeyA Then
                                For j = 0 To nTaxItems - 1
                                    If arryTaxSums(j).m_lTrn = .m_Taxes(i).m_lTrn Then
                                        Exit For
                                    End If
                                Next j
                                If j = nTaxItems Then
                                    If UBound(arryTaxSums) < nTaxItems Then
                                        ReDim Preserve arryTaxSums(UBound(arryTaxSums) + ARRAY_SIZE_INC)
                                    End If
                                    arryTaxSums(j).m_lTrn = .m_Taxes(i).m_lTrn
                                    arryTaxSums(j).m_lLimit = .m_Taxes(i).m_lLimit
                                    arryTaxSums(j).m_cBasis = .m_Taxes(i).m_cBasis
                                    arryTaxSums(j).m_fAmount = .m_Taxes(i).m_dAmount
                                    arryTaxSums(j).m_fRate = .m_Taxes(i).m_fRate
                                    arryTaxSums(j).m_lGallons = .m_InputParms.m_dGrossUnits
                                    nTaxItems = nTaxItems + 1
                                Else
                                    arryTaxSums(j).m_lGallons = arryTaxSums(j).m_lGallons + .m_InputParms.m_dGrossUnits
                                End If
                                If Trim(.m_Taxes(i).m_sWhere) = "TX" And .m_Taxes(i).m_cLevel = vbKeyS Then
                                    sParm1055 = sSysParm1055
                                End If
                            End If
                        End If
                        If Trim(.m_Taxes(i).m_sWhere) = "TX" And .m_Taxes(i).m_cLevel = vbKeyS And .m_Taxes(i).m_cCharge <> vbKeyN Then
                            Texas_Tax = True
                        End If
                    Next i
                    
                    sTaxGroup = fnGetTaxGroup(.m_InputParms.m_sProduct)
                    
                    If sTaxGroup <> sParm1055 Then
                        lGasAmount = lGasAmount + .m_InputParms.m_dGrossUnits
                    End If
                    
                    lFuelAmount = lFuelAmount + .m_InputParms.m_dGrossUnits
                End If
            End With
        Next nIdx
        If nTaxItems > 0 Then
            nSign = 1
            If lFuelAmount < 0 Then
                nSign = -1
                lFuelAmount = Abs(lFuelAmount)
                lGasAmount = Abs(lGasAmount)
            End If
            lTaxableAmount = lFuelAmount
            
            If Texas_Tax Then
                If lGasAmount = lFuelAmount Then
                    If lGasAmount >= 7000 And lGasAmount < 8000 Then
                        lTaxableAmount = 8000
                    End If
                End If
            End If
            fTypeATaxes(nLadDIdx) = arryTaxSums(nTaxItems - 1).m_lLimit
            For i = 0 To nTaxItems - 1
                If lTaxableAmount < Abs(arryTaxSums(i).m_lLimit) Then
                    fTypeATaxes(nLadDIdx) = arryTaxSums(i).m_fRate * nSign
                    Exit For
                End If
            Next i
            
            lUnits(nLadDIdx) = lFuelAmount * nSign
        Else
            fTypeATaxes(nLadDIdx) = 0
        End If
        
    Next nLadDIdx
    
    fTaxes = 0
    For nLadDIdx = 0 To UBound(fTypeATaxes)
        fTaxes = fTaxes + tfnRound(fTypeATaxes(nLadDIdx), 2)
    Next
    
    If fTaxes <> 0 Then
        'Now Need to redistribute to the individual line amount
        For nLadDIdx = 0 To nLadDCnt - 1
            If fTypeATaxes(nLadDIdx) <> 0 Then
                If lUnits(nLadDIdx) = 0 Then
                    fTaxRate = 0
                Else
                    fTaxRate = fTypeATaxes(nLadDIdx) / lUnits(nLadDIdx)
                End If
                For nIdx = 0 To nInfoCount - 1
                    With aryTaxInfo(nIdx)
                        If .m_lInvoice = lBOL Then
                            For i = 0 To .m_nRecord - 1
                                If .m_Taxes(i).m_cCharge <> vbKeyN Then
                                    If .m_Taxes(i).m_cBasis = vbKeyA Then
                                            If .m_Taxes(i).m_dAmount <> 0 Then
                                                'david 08/08/2005  #489634
                                                'fixed rounding problem
                                                '.m_Taxes(i).m_dAmount = tfnRound(fTaxRate * .m_InputParms.m_dGrossUnits, 2)
                                                .m_Taxes(i).m_dAmount = tfnRound(fTaxRate * .m_InputParms.m_dGrossUnits, DEFAULT_DECIMALS)
                                                sTaxState = .m_Taxes(i).m_sWhere
                                            End If
                                    End If
                                End If
                            Next i
                        End If
                    End With
                Next nIdx
            End If
        Next nLadDIdx
    End If
End Sub

Private Function fnATaxInfoItem(ByVal bExt As Boolean, _
                                 ByVal nOrig As Integer, _
                                 ByVal lDest As Long, _
                                 ByVal lInv As Long, _
                                 ByVal nLine As Integer, _
                                 sProduct As String, _
                                 sTaxUseGroup As String, _
                                 fDollarBase As Double, _
                                 fFreight As Double, _
                                 fGross As Double, _
                                 fNet As Double, _
                                 sBillFlag As String, _
                                 Optional vLineType As Variant, _
                                 Optional bShowMsgBox As Boolean = True, _
                                 Optional sErrMsg As String = "", _
                                 Optional hld_vessel As Integer = -1, _
                                 Optional hld_gals As Double = 0#, _
                                 Optional hld_miles As Double = 0#, _
                                 Optional hld_tot_gals As Double = 0#) As Boolean
    
    Dim Taxes() As CApplicableTaxes
    Dim myStruct As CMovementLineItem
    Dim nRet As Long
    Dim i As Integer
    Dim j As Integer
    Dim nIdx As Integer
    Dim nTries As Integer
    Dim nRecords As Long
    Dim k As Integer
    
    fnATaxInfoItem = False
    If Not fnLoadTaxMods Then
        Exit Function
    End If

    With myStruct
        .m_sProduct = sProduct
        .m_sUseGroup = sTaxUseGroup
        .m_dDollarBase = fDollarBase
        .m_dFreight = 0
        .m_dNetUnits = fNet
        .m_dGrossUnits = fGross
        .m_dFreight = fFreight
        If sBillFlag <> "" Then
        .m_cBilledUnitsFlag = Asc(sBillFlag)
        Else
        .m_cBilledUnitsFlag = vbKeyN
        End If
        If IsMissing(vLineType) Then
            .m_tLineType = TXMOD_EXTDEL
        Else
            .m_tLineType = vLineType
        End If
    End With
    nTries = TAXMOD_TRIES
    subSet_gl_sd_gals hld_vessel, hld_gals, hld_miles, hld_tot_gals
    subSet_tx_date t_tax_date
    
    Do
        nRet = TXMOD_VB_CalculateLineTax(App.hInstance, Forms(0).hwnd, myStruct, nRecords)
        
        If i = TXMOD_TASKBLOCKED Then
            DoEvents
        End If
        nTries = nTries - 1
    Loop Until nRet <> TXMOD_TASKBLOCKED Or nTries <= 0
    If nRet <> TXMOD_OK Then
        subTaxModErrors nRet, bShowMsgBox, sErrMsg
        Exit Function
    End If
        
    ReDim Taxes(nRecords - 1)
    nTries = TAXMOD_TRIES
    Do
        nRet = TXMOD_GetLineTaxes(App.hInstance, Forms(0).hwnd, Taxes())
        If i = TXMOD_TASKBLOCKED Then
            DoEvents
        End If
        nTries = nTries - 1
    Loop Until nRet <> TXMOD_TASKBLOCKED Or nTries <= 0
    
    If nRet <> TXMOD_OK Then
        subTaxModErrors nRet, bShowMsgBox, sErrMsg
        Exit Function
    End If
        
    For nIdx = 0 To nInfoCount - 1
        If aryTaxInfo(nIdx).m_lInvoice = lInv Then
            If aryTaxInfo(nIdx).m_nLine = nLine Then
                Exit For
            End If
        End If
    Next nIdx
    If nIdx >= nInfoCount Then
        If UBound(aryTaxInfo) < nInfoCount Then
            ReDim Preserve aryTaxInfo(UBound(aryTaxInfo) + ARRAY_SIZE_INC)
        End If
        nInfoCount = nInfoCount + 1
    End If
    
    aryTaxInfo(nIdx).m_nRecord = nRecords
    aryTaxInfo(nIdx).m_lInvoice = lInv
    aryTaxInfo(nIdx).m_nLine = nLine
    aryTaxInfo(nIdx).m_InputParms = myStruct
    ReDim aryTaxInfo(nIdx).m_Taxes(nRecords - 1)
    For i = 0 To nRecords - 1
        aryTaxInfo(nIdx).m_Taxes(i) = Taxes(i)
    Next i
    
    Dim sIE As String
    Dim vTaxUseG As Variant
    Dim FoundIdx As Integer

    If bExt Then
        sIE = "E"
    Else
        sIE = "I"
    End If
    
    vTaxUseG = interstate_use_groups(sIE, nOrig, lDest, sProduct, sTaxUseGroup)
    If Trim(vTaxUseG) <> szEMPTY Then
        vTaxUseG = Split(vTaxUseG, "|")
        For i = 0 To UBound(vTaxUseG)
            myStruct.m_sUseGroup = vTaxUseG(i)
            Do
                nRet = TXMOD_VB_CalculateLineTax(App.hInstance, Forms(0).hwnd, myStruct, nRecords)
                
                If i = TXMOD_TASKBLOCKED Then
                    DoEvents
                End If
                nTries = nTries - 1
            Loop Until nRet <> TXMOD_TASKBLOCKED Or nTries <= 0
            If nRet <> TXMOD_OK Then
                subTaxModErrors nRet, bShowMsgBox, sErrMsg
                Exit Function
            End If
                
            ReDim Taxes(nRecords - 1)
            nTries = TAXMOD_TRIES
            Do
                nRet = TXMOD_GetLineTaxes(App.hInstance, Forms(0).hwnd, Taxes())
                If i = TXMOD_TASKBLOCKED Then
                    DoEvents
                End If
                nTries = nTries - 1
            Loop Until nRet <> TXMOD_TASKBLOCKED Or nTries <= 0
            
            If nRet <> TXMOD_OK Then
                subTaxModErrors nRet, bShowMsgBox, sErrMsg
                Exit Function
            End If
                
            For nIdx = 0 To nInfoCount - 1
                If aryTaxInfo(nIdx).m_lInvoice = lInv Then
                    If aryTaxInfo(nIdx).m_nLine = nLine Then
                        Exit For
                    End If
                End If
            Next nIdx
            If nIdx >= nInfoCount Then
                If UBound(aryTaxInfo) < nInfoCount Then
                    ReDim Preserve aryTaxInfo(UBound(aryTaxInfo) + ARRAY_SIZE_INC)
                End If
                nInfoCount = nInfoCount + 1
            End If
            
            aryTaxInfo(nIdx).m_lInvoice = lInv
            aryTaxInfo(nIdx).m_nLine = nLine
            aryTaxInfo(nIdx).m_InputParms = myStruct
                        
            For j = 0 To nRecords - 1
                FoundIdx = -1
                For k = 0 To aryTaxInfo(nIdx).m_nRecord - 1
                    If aryTaxInfo(nIdx).m_Taxes(k).m_lTrn = Taxes(j).m_lTrn Then
                        FoundIdx = k
                        Exit For
                    End If
                Next
                If FoundIdx < 0 Then
                    aryTaxInfo(nIdx).m_nRecord = aryTaxInfo(nIdx).m_nRecord + 1
                    ReDim Preserve aryTaxInfo(nIdx).m_Taxes(aryTaxInfo(nIdx).m_nRecord - 1)
                    aryTaxInfo(nIdx).m_Taxes(aryTaxInfo(nIdx).m_nRecord - 1) = Taxes(j)
                Else
                    aryTaxInfo(nIdx).m_Taxes(FoundIdx) = Taxes(j)
                End If
                
            Next
        Next
    End If
    
    fnATaxInfoItem = True
End Function

Private Function interstate_use_groups(ByVal sIE As String, _
                                     ByVal nOrig As Integer, _
                                     ByVal lDest As Long, _
                                     ByVal sProd As String, _
                                     ByVal sTaxUseG) As String
        Dim strSQL As String
        Dim or_state As String
        Dim ds_state As String
        Dim rsTemp As Recordset
        Dim n_flag As Integer
        Static staParm10032 As String
        Dim inter_use As String
        Dim sTaxUseOut As String
         
         
        interstate_use_groups = szEMPTY
        
        strSQL = "SELECT icl_fuel_tax FROM inv_classes,inv_header WHERE ivh_prodtcl=icl_code AND icl_type = 'T'"
        Set rsTemp = txMod_OpenRecord(strSQL)
        If rsTemp Is Nothing Then
            Exit Function
        ElseIf rsTemp.RecordCount = 0 Then
            Exit Function
        End If
        If Trim(rsTemp!icl_fuel_tax & szEMPTY) <> "Y" Then
            Exit Function
        End If
        
        If sIE = "E" Then
            strSQL = "SELECT city_state FROM fm_terminal,sys_city WHERE fmt_city = city_code AND fmt_code=" & CStr(nOrig)
        Else
            strSQL = "SELECT city_state FROM sys_prft_ctr,sys_city WHERE prft_city = city_code AND prft_ctr=" & CStr(nOrig)
        End If
        Set rsTemp = txMod_OpenRecord(strSQL)
        If rsTemp Is Nothing Then
            Exit Function
        ElseIf rsTemp.RecordCount = 0 Then
            Exit Function
        End If
        or_state = Trim(rsTemp!city_state & szEMPTY)
        
        strSQL = "SELECT city_state FROM sys_prft_ctr,sys_city WHERE prft_city = city_code AND prft_ctr=" & CStr(lDest)
        Set rsTemp = txMod_OpenRecord(strSQL)
        If rsTemp Is Nothing Then
            Exit Function
        ElseIf rsTemp.RecordCount = 0 Then
            strSQL = "SELECT city_state FROM ar_altname,sys_city WHERE an_customer=" & CStr(lDest) & " AND an_city = city_code"
            Set rsTemp = txMod_OpenRecord(strSQL)
            If rsTemp Is Nothing Then
                Exit Function
            ElseIf rsTemp.RecordCount = 0 Then
                Exit Function
            End If
        End If
        ds_state = Trim(rsTemp!city_state & szEMPTY)
            
        
        n_flag = 0
        If staParm10032 = szEMPTY Then
            staParm10032 = "N"
            strSQL = "SELECT parm_field FROM sys_parm WHERE parm_nbr = 10032"
            Set rsTemp = txMod_OpenRecord(strSQL)
            If rsTemp Is Nothing Then
                Exit Function
            End If
            If rsTemp.RecordCount > 0 Then
                If Trim(rsTemp!parm_field & szEMPTY) = "A" Or Trim(rsTemp!parm_field & szEMPTY) = "M" Then
                    staParm10032 = Trim(rsTemp!parm_field & szEMPTY)
                End If
            End If
        End If
        If staParm10032 <> "N" Then
            If sTaxUseG <> szEMPTY And sIE = "E" Then
                n_flag = fnTaxOrigin_is_PhysicalDest(sTaxUseG, sProd, ds_state)
            End If
        End If
        If or_state = ds_state Then
            Exit Function
        End If
                
        sTaxUseOut = szEMPTY
        
        inter_use = or_state & Space(2) & "SSSS"
        If n_flag = 1 Then
            inter_use = or_state & "RRSSSS"
        End If
        sTaxUseOut = inter_use
        
        inter_use = "SSSS" & Space(2) & ds_state
        If n_flag = 1 Then
            inter_use = "SSSSRR" & ds_state
        End If
        sTaxUseOut = sTaxUseOut & "|" & inter_use
        
        interstate_use_groups = sTaxUseOut
End Function

Private Function fnTaxOrigin_is_PhysicalDest(ByVal c_use_g As String, _
                                     ByVal prodcode As String, _
                                     ByVal ds_state As String) As Integer
   Dim strSQL As String
   Dim rsTemp As Recordset
   Dim tax_origin As String
   Dim n_flag As Integer
  
   strSQL = "SELECT * FROM tx_header" _
        & " WHERE txh_level = 'S' and txh_trn IN " _
        & "(SELECT tul_trn FROM tx_use_link, tx_use_g, inv_header" _
        & " WHERE tul_link = tu_link AND tu_use_g = " & tfnSQLString(c_use_g) _
        & " AND tu_prodtcl = ivh_prodtcl AND ivh_product = " & tfnSQLString(prodcode) & ")"

   tax_origin = szEMPTY
   Set rsTemp = txMod_OpenRecord(strSQL)
   If rsTemp Is Nothing Then
        Exit Function
   ElseIf rsTemp.RecordCount = 0 Then
        Exit Function
   End If
   n_flag = 0
   Do While Not rsTemp.EOF
      If tax_origin = szEMPTY Then
         tax_origin = Trim(rsTemp.txh_where & szEMPTY)
         n_flag = 1
      Else
         If tax_origin <> Trim(rsTemp.txh_where & szEMPTY) Then
            n_flag = 0
            Exit Do
         End If
      End If
      rsTemp.MoveNext
   Loop
   
   If n_flag = 1 And tax_origin = ds_state Then
      fnTaxOrigin_is_PhysicalDest = 1
   Else
      fnTaxOrigin_is_PhysicalDest = 0
   End If

End Function
Private Function txMod_ExeCuteSQL(strSQL As String, _
                             Optional vCaller As Variant, _
                             Optional vMsg As Variant, _
                             Optional vDB As Variant) As Boolean

    Dim objDB As Database
    
    If IsMissing(vDB) Then
        Set objDB = t_dbMainDatabase
    Else
        Set objDB = vDB
    End If

    On Error GoTo errExecute
    If objDB Is t_dbMainDatabase Then
        txMod_ExeCuteSQL = objDB.ExecuteSQL(strSQL)
    Else
        objDB.Execute strSQL
    End If
    txMod_ExeCuteSQL = True
extExecute:
    On Error GoTo 0
    Exit Function

errExecute:
    Dim bShow As Boolean
    bShow = Not IsMissing(vMsg)
    If IsMissing(vCaller) Then
        tfnErrHandler "", strSQL, bShow
    Else
        tfnErrHandler CStr(vCaller), strSQL, bShow
    End If
    Err.Clear
    txMod_ExeCuteSQL = False
    Resume extExecute

End Function

Private Function txMod_OpenRecord(strSQL As String, _
                              Optional vCaller As Variant, _
                              Optional vMsg As Variant, _
                              Optional vDB As Variant, _
                              Optional bMoveLast As Boolean = True) As Recordset
    ' Get records from the given SQL statement
    Dim objDB As Database
    Dim rsTemp As Recordset

    If IsMissing(vDB) Then
        Set objDB = t_dbMainDatabase
    Else
        Set objDB = vDB
    End If
    On Error GoTo SQLError
    If objDB Is t_dbMainDatabase Then
        Set rsTemp = objDB.OpenRecordset(strSQL, dbOpenSnapshot, dbSQLPassThrough)
    Else
        Set rsTemp = objDB.OpenRecordset(strSQL, dbOpenSnapshot)
    End If
    
    If bMoveLast Then
        If rsTemp.RecordCount > 0 Then
            rsTemp.MoveLast
            rsTemp.MoveFirst
        End If
    End If
    
    Set txMod_OpenRecord = rsTemp

quitsub:
    On Error GoTo 0
    Exit Function

SQLError:
    Dim bShow As Boolean
    bShow = Not IsMissing(vMsg)
    If IsMissing(vCaller) Then
        tfnErrHandler "", strSQL, bShow
    Else
        tfnErrHandler CStr(vCaller), strSQL, bShow
    End If
    Set txMod_OpenRecord = Nothing
    Err.Clear
    Resume quitsub
End Function

'#"End ATax Code Region" - 452155wj101304
'#------------------------------------------------------------


