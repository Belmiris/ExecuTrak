Attribute VB_Name = "modTXMODDLL"
Option Explicit
#Const DEBUGGING = False
#Const HARD_CODED_TX_LD_FEE = False

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

'Public Type CApplicableTaxes
'    m_sCode As String * 5
'    m_cLevel As Byte
'    m_sWhere As String * 4
'    m_nOrder As Integer
'    m_lGlAcct As Long
'    m_fRate As Single
'    m_cCharge As Byte
'    m_cBasis As Byte
'    m_lLimit As Long
'    m_sDate As String * 10
'    m_lTrn As Long
'    m_sDescr As String * 30
'    m_cFuelFlag As Byte
'    m_dAmount As Double
'End Type

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
Public Function fnCalculateTaxes(sProduct As String, _
                                  sTaxUseGroup As String, _
                                  dGross As Double, _
                                  dQuantity As Double, _
                                  ByRef dTaxAmount As Double, _
                                  Optional bShowMsgBox As Boolean = True, _
                                  Optional sErrMsg As String = "", _
                                  Optional bApplyLimit As Boolean = True) As Boolean

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
            If aryTaxInfo(i).m_Taxes(j).m_cCharge = vbKeyY Then
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
                    
                    If lTaxableAmount = 0 Then
                       fTypeATaxes(nDel) = 0#
                    ElseIf lTaxableAmount < 2500 Then
                       fTypeATaxes(nDel) = 12.5 * nSign
                    ElseIf lTaxableAmount < 5000 Then
                       fTypeATaxes(nDel) = 25# * nSign
                    ElseIf lTaxableAmount < 8000 Then
                       fTypeATaxes(nDel) = 37.5 * nSign
                    ElseIf lTaxableAmount < 10000 Then
                       fTypeATaxes(nDel) = 50# * nSign
                    Else
                       m = 3
                       
                       Do While True
                          If lTaxableAmount < m * 5000 Then
                             fTypeATaxes(nDel) = m * 25# * nSign
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
                               Optional bApplyLimit As Boolean = True) As Double

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
        .m_dDollarBase = Abs(fDollarBase)
        .m_dFreight = Abs(fFreight)                     'Freight amount
        .m_dNetUnits = Abs(lNet)
        .m_dGrossUnits = Abs(lGross)
        .m_cBilledUnitsFlag = Asc(szBillingFlag)
        .m_tLineType = nLineType
    End With
    
    'Determine the number of items in the record
    nTries = TAXMOD_TRIES
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
                               Optional bApplyLimit As Boolean = True) As Double

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
        .m_dDollarBase = Abs(fDollarBase)
        .m_dFreight = Abs(fFreight)                     'Freight amount
        .m_dNetUnits = Abs(lNet)
        .m_dGrossUnits = Abs(lGross)
        .m_cBilledUnitsFlag = Asc(szBillingFlag)
        .m_tLineType = nLineType
    End With
    
    'Determine the number of items in the record
    nTries = TAXMOD_TRIES
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

Public Sub subAddTaxItem(szProduct As String, _
                          ByVal lNet As Long, _
                          ByVal lGross As Long, _
                          ByVal fCurrentCost As Double, _
                          ByVal fFreight As Double, _
                          szTaxUseGroup As String, _
                          szBillingFlag As String, _
                          ByVal bExternal As Boolean, _
                          ByVal bLading As Boolean)
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
    fnAddTaxInfoItem -1, nInfoCount, szProduct, szTaxUseGroup, fDollarBase, fFreight, CDbl(lGross), CDbl(lNet), szBillingFlag, nLineType
End Sub

'Add By junsong 06/04/2002 to fix deciaml gallon
Public Sub subAddTaxItem1(szProduct As String, _
                          ByVal fNet As Double, _
                          ByVal fGross As Double, _
                          ByVal fCurrentCost As Double, _
                          ByVal fFreight As Double, _
                          szTaxUseGroup As String, _
                          szBillingFlag As String, _
                          ByVal bExternal As Boolean, _
                          ByVal bLading As Boolean)
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
    fnAddTaxInfoItem -1, nInfoCount, szProduct, szTaxUseGroup, fDollarBase, fFreight, fGross, fNet, szBillingFlag, nLineType
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
                                 Optional bApplyLimit As Boolean = True) As Boolean
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
                                 Optional bApplyLimit As Boolean = True) As Boolean
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
                                 Optional sErrMsg As String = "") As Boolean
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
                                    Optional sErrMsg As String = "") As Boolean
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
