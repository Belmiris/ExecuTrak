VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EDIToolEngine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***********************************************************'
'
' Copyright (c) 1996 FACTOR, A Division of W.R.Hess Company
'
' Module name   : ETENGINE.CLS
' Date          : Apr. 14, 1998
' Programmer    : Qinggang Ma
'
' Defines a tool to write flat text file for EDI fuel movement tax information
'   More types can be put into this module to write other types of EDI flat file
'
'***********************************************************'

'Properties:
'Get       EDITypeR:   returns a flag for R type record
'Get       EDITypeD:   returns a flag for D type record
'Get       ModeTest:   Return a flag for test mode
'Get       ModeProduct:   Return a flag for product mode
'Let       StartDate:   Set the start date
'Let       EndDate:     Set the end date
'Let       EDIType:     Set the EDI type
'Let       TaxpayerLicense: Set the taxpayer's license number
'Let       TheirID:     Set their ID
'Let       Mode:        Set the mode --- test or product
'Let       FileName     Set the EDI flat text file name
'Get       LastError:   Return the last error code (0 means no error)
'Get       LastMessage: Return the last error message

'Public method to call
'       BeginTrans: Reset all variables, need to call before generate an EDI flat text file
'       AddEDIRecords(sSchedule As String, _
'                     nTypeCode As String, _
'                     sTaxClass As String, _
'                     nProdRefCode As String)
'        : Parameters for this function are schedule code,
'                                       tax type code,
'                                       product tax class
'                                       product reference number
'       EndTrans:   Write EDI header record and close the file

Option Explicit
    
Private Const EDI_TYPE_R = "R"
Private Const EDI_TYPE_D = "D"
Private Const MODE_PRODUCT = "P"
Private Const MODE_TEST = "T"
Private Const SYSPARM_FED_ID = 9008
Private Const ARRAY_INC = 10
Private Const FILE_HANDLE = 21
Private Const REPORT_STATE = "AR"

Private Const EDI_FMTAX_RECORD_LENGGH = 232

Private Const EC_INITIAL = 0
Private Const EC_INVALID_START_DATE = -1001
Private Const EC_INVALID_END_DATE = -1002
Private Const EC_NULL_SCHD_CODE = -2001
Private Const EC_NULL_TAX_CLASS = -3001
Private Const EC_INVALID_EDI_TYPE = -4001
Private Const EC_INVALID_TYPE_CODE = -4002
Private Const EC_NULL_FEIN = -4003
Private Const EC_CREATE_TT1_R = -5001
Private Const EC_CREATE_TT2_R = -5002
Private Const EC_CREATE_TT3_R = -5003
Private Const EC_GET_RECORD_R = -5004
Private Const EC_CREATE_TT1_D = -5101
Private Const EC_CREATE_TT2_D = -5102
Private Const EC_CREATE_TT3_D = -5103
Private Const EC_GET_RECORD_D = -5104
Private Const EC_INVALID_THIER_ID = -6001
Private Const EC_INVALID_TAX_LICENSE = -7001
Private Const EC_INVALID_MODE = -8001
Private Const EC_INIT_PARM1_1 = -9101
Private Const EC_INIT_PARM1_2 = -9101
Private Const EC_INIT_PARM2_1 = -9201
Private Const EC_INIT_PARM2_2 = -9202

Private Const EC_CANCEL_PROCESS = -9203
Private Const EC_EXIT_PROGRAM = -9204

Private Const EDI_ITM_RECORD_TYPE = "ITM"
Private Const EDI_ITM_TAX_SCHEDULE_CODE = "T3"
Private Const EDI_ITM_PRODUCT_TYPE = "PG"
Private Const EDI_ITM_RECEIVER_TRANS_SET = "94"
Private Const EDI_ITM_SHIP_METHOD_R = "J "
Private Const EDI_ITM_SHIP_METHOD_D = "J "
Private Const EDI_ITM_CARRIER_ID_CODE = "CA"
Private Const EDI_ITM_CARRIER_FEIN_TYPE = "24"
Private Const EDI_ITM_ORIGIN_TYPE = "N"
Private Const EDI_ITM_ORIGIN_ENTITY_ID = "SF"
Private Const EDI_ITM_ORIGIN_IRS_T_ID = "TC"
Private Const EDI_ITM_DEST_TYPE = "N"
Private Const EDI_ITM_DEST_ENTITY_ID = "ST"
Private Const EDI_ITM_DEST_IRS_T_ID = "TC"
Private Const EDI_ITM_BUY_SELL_ENTITY_ID_R = "SE"
Private Const EDI_ITM_BUY_SELL_ENTITY_ID_D = "BY"
Private Const EDI_ITM_BUY_SELL_FEIN_TYPE = "24"
Private Const EDI_ITM_REF_TYPE_CODE = "BM"
Private Const EDI_ITM_DATE_TYPE_CODE = "095"
Private Const EDI_ITM_TAX_INFO_ID = "5010"
Private Const EDI_ITM_UOM = "GA"

Private Const EDI_ITM_ORIGIN_IRS_T_CODE = "         "
Private Const EDI_ITM_DEST_IRS_T_CODE = "         "

Private Const EDI_HDR_RECORD_TYPE = "HDR"
Private Const EDI_HDR_REF_NBR_CODE = "T6"
Private Const EDI_HDR_TAX_AUTHORITY_CODE = "47"
Private Const EDI_HDR_STATE_CODE = "AR"
Private Const EDI_HDR_TAXPAYER_FEIN_TYPE = "24"
Private Const EDI_HDR_STATE_ID_NBR = "49"
Private Const EDI_HDR_FILING_PERIOD_CODE = "194"
Private Const EDI_HDR_RANGE_OF_DATE = "RD6"
Private Const EDI_HDR_TAX_INFO_ID_NBR_1 = "5010"
Private Const EDI_HDR_UOM = "GA"
Private Const EDI_HDR_TAX_INFO_ID_NBR_2 = "5067"
Private Const EDI_HDR_VERSION_NBR = "1.00"
Private Const EDI_HDR_PRIMARY_TAXPAYER_COD = "TP"
Private Const EDI_HDR_TEST_PRODUCTION = "T"

Private Type tpLinkName
    m_sName As String
    m_lLink As Long
End Type
Private aryProduct() As tpLinkName
Private nProductCount As Integer
Private aryTerminal() As tpLinkName
Private nTerminalCount As Integer
Private aryProfitCenter() As tpLinkName
Private nProfitCenterCount As Integer

Private aryCustomerCity() As tpLinkName
Private nCustomerCityCount As Integer

Private Type tpLinkInfo
    m_sLink As String
    m_sInfo1 As String
    m_sInfo2 As String
    m_sInfo3 As String
End Type

Private aryCity() As tpLinkInfo
Private nCityCount As Integer
Private aryCarrier() As tpLinkInfo
Private nCarrierCount As Integer
Private aryCustomer() As tpLinkInfo
Private nCustomerCount As Integer
Private aryVendor() As tpLinkInfo
Private nVendorCount As Integer

Private Type tpVendorTaxID
    m_lVendor As Long
    m_sState As String
    m_sTaxClass As String
    m_sTaxID As String
End Type
Private aryVendorTaxID() As tpVendorTaxID
Private nVTaxIDCount As Integer

Private Type tpCustomerTaxID
    m_lCustomer As Long
    m_lTrnLink As Long
    m_sTaxID As String
End Type
Private aryCustomerTaxID() As tpCustomerTaxID
Private nCTaxIDCount As Integer

'Variables(properties) need to set by the program
Private m_sStartDate As String
Private m_sEndDate As String
Private m_sEDIType As String
Private m_sTheirID As String
Private m_sTaxTypeCode As String
Private m_sTaxpayerLicense As String
Private m_sMode As String
Private m_sFileName As String

Private m_sCarrier_Name As String
Private m_sCarrier_FEIN As String

'Variables need to set inside the program
Private m_nTypeCode As String
Private m_nProductRefCode As String
Private m_lTotalGallons As Long
Private m_sCompany As String
Private m_sTaxpayerID As String
Private m_nLineLength As Integer
Private m_nErrorNumber As Integer
Private m_sLastVBError As String
Private m_lHeaderPosition As Long

Private m_lTaxPaidTrn(20) As Long
Private m_nTPTCount As Integer

Private sFMTaxLineBuffer As String * EDI_FMTAX_RECORD_LENGGH
'

'david 01/25/00
'this function will return #of tax record for the schedule
Public Function AddEDIRecords(sSchedule As String, _
                         nTypeCode As String, _
                         sTaxClass As String, _
                         nProdRefCode As String, _
                         Optional vExtractFrom) As Long

    Dim ciTemp As ETFMTaxRecord
    Dim colItems As Collection
    Dim sEDIItem As String
    Dim lTaxRecords As Long
    
    If m_nErrorNumber <> ErrorFree Then
        Exit Function
    End If
    
    m_nTypeCode = nTypeCode
    m_nProductRefCode = nProdRefCode
    
    lTaxRecords = 0
    
    Set colItems = fnOpenFMTaxRecord(sSchedule, nTypeCode, sTaxClass, nProdRefCode, vExtractFrom)
    If m_nErrorNumber <> EC_INITIAL Then
        Exit Function
    End If
    
    If fnCheckExitCancel Then
        Exit Function
    End If
    
    If Not colItems Is Nothing Then
        subCheckFEIN colItems
        If m_nErrorNumber <> EC_INITIAL Then
            Exit Function
        End If
    End If
    
    If Not colItems Is Nothing Then
        For Each ciTemp In colItems
            m_lTotalGallons = m_lTotalGallons + ciTemp.Quantity
            sEDIItem = fnBuildEDIFMTaxItem(ciTemp)
            If Not fnWriteLine(sEDIItem) Then
                Exit Function
            End If
            lTaxRecords = lTaxRecords + 1
        Next ciTemp
    End If
    
    AddEDIRecords = lTaxRecords
End Function

Public Sub BeginDetail()
    m_lTotalGallons = 0
    Seek #FILE_HANDLE, m_lHeaderPosition + 1
End Sub

Public Sub BeginTrans()

    Dim ciTemp As New ETFMTaxRecord
    Dim sTemp As String
    
    Close #FILE_HANDLE
    m_nErrorNumber = ErrorFree
    m_sLastVBError = ""
    
    If m_nErrorNumber <> ErrorFree Then
        Exit Sub
    End If
    
    'Get a template item line
    sTemp = fnBuildEDIFMTaxItem(ciTemp)
    m_nLineLength = Len(sTemp)
    
    On Error Resume Next
    Kill m_sFileName
    On Error GoTo errOpenFile
    Open m_sFileName For Random As #FILE_HANDLE Len = EDI_FMTAX_RECORD_LENGGH
    If fnWriteLine(sTemp) Then
        subCheckRequiredFields
        m_lHeaderPosition = 1
        m_lTotalGallons = 0
    End If
    
    Exit Sub
errOpenFile:
    m_nErrorNumber = Err.Number
    m_sLastVBError = fnGetVBError
End Sub

Property Let EDIType(sTemp As String)
    m_sEDIType = sTemp
End Property

Property Get EDITypeR() As String
    EDITypeR = EDI_TYPE_R
End Property

Property Get EDITypeD() As String
    EDITypeD = EDI_TYPE_D
End Property

Property Let EndDate(sTemp As String)
    m_sEndDate = tfnFormatDate(sTemp)
End Property

Property Let Carrier_Name(sTemp As String)
    m_sCarrier_Name = sTemp
End Property

Property Let Carrier_FEIN(sTemp As String)
    m_sCarrier_FEIN = sTemp
End Property

Public Sub EndTrans()
    If m_nErrorNumber <> ErrorFree Then
        Exit Sub
    End If
    
'    If m_lTotalGallons >= 0 Then
'        If fnWriteLine(fnBuildEDIFMTaxHeader, m_lHeaderPosition) Then
'            Close #FILE_HANDLE
'        End If
'    Else
        Close #FILE_HANDLE
'    End If
End Sub

Public Sub EndDetail()
    Dim lCurr As Long
    If m_nErrorNumber <> ErrorFree Then
        Exit Sub
    End If
    
    lCurr = Seek(FILE_HANDLE)
'    If m_lTotalGallons >= 0 Then
        If fnWriteLine(fnBuildEDIFMTaxHeader, m_lHeaderPosition) Then
            m_lHeaderPosition = lCurr
            m_lTotalGallons = 0
        End If
'    Else
'        If fnWriteLine(Space(230), m_lHeaderPosition) Then  'david
'            m_lHeaderPosition = lCurr  'david
'            m_lTotalGallons = 0
'        End If  'david
'    End If
End Sub


Property Get ErrorFree() As Integer
    ErrorFree = EC_INITIAL
End Property

Property Let FileName(sTemp As String)
    m_sFileName = sTemp
End Property

Private Function fnConvertFEDID(sID As String) As String
    Dim sTemp As String
    Dim I As Integer
    Dim sChar As String
    
    sTemp = ""
    For I = 1 To Len(sID)
        sChar = Mid(sID, I, 1)
        If sChar <> "-" And sChar <> " " Then
            sTemp = sTemp & sChar
        End If
    Next
    fnConvertFEDID = sTemp
    
End Function

Private Function fnBuildEDIFMTaxItem(ciTemp As ETFMTaxRecord) As String

    Dim sEDIItem As String
    
    
    With ciTemp
        sEDIItem = EDI_ITM_RECORD_TYPE & EDI_ITM_TAX_SCHEDULE_CODE & fnFixLength(m_nTypeCode, 3)
        sEDIItem = sEDIItem & EDI_ITM_PRODUCT_TYPE & fnFixLength(m_nProductRefCode, 3) _
                 & EDI_ITM_RECEIVER_TRANS_SET
        
        If m_sEDIType = EDI_TYPE_R Then
            sEDIItem = sEDIItem & EDI_ITM_SHIP_METHOD_R
        ElseIf m_sEDIType = EDI_TYPE_D Then
            sEDIItem = sEDIItem & EDI_ITM_SHIP_METHOD_D
        Else
            sEDIItem = sEDIItem & "  "
            'Error
        End If

        sEDIItem = sEDIItem & EDI_ITM_CARRIER_ID_CODE & fnFixLength(.CarrierName, 35) _
                 & EDI_ITM_CARRIER_FEIN_TYPE & fnFixLength(.CarrierFedID, 9) & EDI_ITM_ORIGIN_TYPE & EDI_ITM_ORIGIN_ENTITY_ID _
                 & fnFixLength(.OriginState, 2) & EDI_ITM_ORIGIN_IRS_T_ID & EDI_ITM_ORIGIN_IRS_T_CODE & fnFixLength(.OriginCity, 18) _
                 & fnFixLength(.OriginState, 2) & EDI_ITM_DEST_TYPE & EDI_ITM_DEST_ENTITY_ID & fnFixLength(.DestState, 2) _
                 & EDI_ITM_DEST_IRS_T_ID & EDI_ITM_DEST_IRS_T_CODE & fnFixLength(.DestCity, 18) _
                 & fnFixLength(.DestState, 2)
        
        If m_sEDIType = EDI_TYPE_R Then
            sEDIItem = sEDIItem & EDI_ITM_BUY_SELL_ENTITY_ID_R
        ElseIf m_sEDIType = EDI_TYPE_D Then
            sEDIItem = sEDIItem & EDI_ITM_BUY_SELL_ENTITY_ID_D
        Else
            'Error
            sEDIItem = sEDIItem & "  "
        End If
        sEDIItem = sEDIItem & fnFixLength(.BuySellName, 35) & EDI_ITM_BUY_SELL_FEIN_TYPE & fnFixLength(.BuySellFedID, 9) & fnFixLength(m_sEDIType, 1) _
                 & EDI_ITM_REF_TYPE_CODE & fnFixLength(.LadingNumber, 15) & EDI_ITM_DATE_TYPE_CODE & fnFormatDate(.LadingDate) & fnFixLength(fnCentury(.LadingDate), 2) _
                 & EDI_ITM_TAX_INFO_ID & fnFixLength(CStr(.Quantity), 10) & EDI_ITM_UOM
    End With
    
    fnBuildEDIFMTaxItem = sEDIItem

End Function

Private Function fnBuildEDIFMTaxHeader() As String

    Dim sHDRLine As String

'    sHDRLine = EDI_HDR_RECORD_TYPE & fnFixLength(m_sTheirID, 35) & EDI_HDR_REF_NBR_CODE _
             & m_sTaxTypeCode & EDI_HDR_TAX_AUTHORITY_CODE & EDI_HDR_STATE_CODE _
             & fnFormatDate(Date) & EDI_HDR_TAXPAYER_FEIN_TYPE & fnFixLength(m_sTaxpayerID, 9) _
             & EDI_HDR_STATE_ID_NBR & fnFixLength(m_sTaxpayerLicense, 11) & EDI_HDR_FILING_PERIOD_CODE _
             & fnFixLength(fnCentury(Date), 2) _
             & EDI_HDR_RANGE_OF_DATE & fnFixLength(fnFormatDate(m_sStartDate), 6) & "-" & fnFixLength(fnFormatDate(m_sEndDate), 6) _
             & EDI_HDR_TAX_INFO_ID_NBR_1 & fnFixLength(CStr(m_lTotalGallons), 10) & EDI_HDR_UOM _
             & EDI_HDR_TAX_INFO_ID_NBR_2 & EDI_HDR_VERSION_NBR _
             & EDI_HDR_PRIMARY_TAXPAYER_COD & fnFixLength(m_sCompany, 35) & m_sMode

    sHDRLine = EDI_HDR_RECORD_TYPE & fnFixLength(m_sTheirID, 35) & EDI_HDR_REF_NBR_CODE _
             & m_sTaxTypeCode & EDI_HDR_TAX_AUTHORITY_CODE & EDI_HDR_STATE_CODE _
             & fnFormatDate(Date) & EDI_HDR_TAXPAYER_FEIN_TYPE & fnFixLength(m_sTaxpayerID, 9) _
             & EDI_HDR_STATE_ID_NBR & fnFixLength(m_sTaxpayerLicense, 11) & EDI_HDR_FILING_PERIOD_CODE _
             & fnFixLength(fnCentury(m_sEndDate), 2) _
             & EDI_HDR_RANGE_OF_DATE & fnFixLength(fnFormatDate(m_sEndDate), 6) & Space(7) _
             & EDI_HDR_TAX_INFO_ID_NBR_1 & fnFixLength(CStr(m_lTotalGallons), 10) & EDI_HDR_UOM _
             & EDI_HDR_TAX_INFO_ID_NBR_2 & EDI_HDR_VERSION_NBR _
             & EDI_HDR_PRIMARY_TAXPAYER_COD & fnFixLength(m_sCompany, 35) & m_sMode

    fnBuildEDIFMTaxHeader = fnFixLength(sHDRLine, m_nLineLength)
End Function

Private Function fnCentury(sDate As String) As String

    Dim dDate As Date
    If IsDate(tfnDateString(sDate)) Then
        dDate = CDate(tfnDateString(sDate))
        fnCentury = CStr(Int(Year(dDate) / 100))
    End If
    
End Function

Private Function fnCStr(vTemp As Variant) As String
    If IsNull(vTemp) Then
        fnCStr = ""
    Else
        fnCStr = Trim(vTemp)
    End If
End Function

Private Function fnFixLength(sSrc As String, _
                             nLen As Integer) As String

    Dim nTemp As Integer
    
    nTemp = Len(sSrc)
    If nTemp < nLen Then
        fnFixLength = sSrc & Space(nLen - nTemp)
    Else
        fnFixLength = Left(sSrc, nLen)
    End If

End Function

Private Function fnFormatDate(sDate As String) As String
    If IsDate(sDate) Then
        fnFormatDate = fnFixLength(Format(sDate, "YYMMDD"), 6)
    Else
        fnFormatDate = Space(6)
    End If
End Function

Private Function fnGetVBError() As String
    fnGetVBError = "VB run-time error# " & Err.Number & ": " & Err.Description
End Function

Private Function fnWriteLine(sTemp As String, _
                             Optional lLoc As Variant) As Boolean
    sFMTaxLineBuffer = sTemp & vbCrLf
    If IsMissing(lLoc) Then
        Put #FILE_HANDLE, , sFMTaxLineBuffer
    Else
        Put #FILE_HANDLE, Val(lLoc), sFMTaxLineBuffer
    End If
    fnWriteLine = True
    Exit Function
errWirte:
    m_nErrorNumber = Err.Number
    m_sLastVBError = fnGetVBError
    fnWriteLine = False
End Function

Property Get LastError() As Integer
    LastError = m_nErrorNumber
End Property

Property Get LastMessage() As String
    Select Case m_nErrorNumber
        Case EC_INITIAL
            LastMessage = "EDI flat file(" & m_sFileName & ") is successfully created"
        Case EC_INVALID_START_DATE
            LastMessage = "Invalid start date or start date not set"
        Case EC_INVALID_END_DATE
            LastMessage = "Invalid end date or end date not set"
        Case EC_NULL_SCHD_CODE
            LastMessage = "Schedule code not set"
        Case EC_NULL_TAX_CLASS
            LastMessage = "Product tax class code not set"
        Case EC_INVALID_EDI_TYPE
            LastMessage = "Invalid EDI type or EDI type not set"
        Case EC_INVALID_TYPE_CODE
            LastMessage = "Invalid EDI type code"
        Case EC_CREATE_TT1_R, EC_CREATE_TT2_R, EC_GET_RECORD_R, EC_CREATE_TT1_D, EC_CREATE_TT2_D, EC_GET_RECORD_D
            LastMessage = "ODBC call failed while trying to gether data"
        Case EC_INVALID_THIER_ID
            LastMessage = "Their ID not set"
        Case EC_INVALID_TAX_LICENSE
            LastMessage = "Tax license number not set"
        Case EC_INVALID_MODE
            LastMessage = "Running mode not valid or not set"
        Case EC_INIT_PARM1_1, EC_INIT_PARM1_2, EC_INIT_PARM2_1, EC_INIT_PARM2_2
            LastMessage = "Failed to retrieve data from database"
        Case EC_CANCEL_PROCESS, EC_EXIT_PROGRAM
            LastMessage = "Process has been cancelled"
        Case Else
            LastMessage = m_sLastVBError
    End Select
End Property


Property Let Mode(sTemp As String)
    m_sMode = sTemp
End Property

Property Get ModeTest() As String
    ModeTest = MODE_TEST
End Property

Property Get ModeProduct() As String
    ModeProduct = MODE_PRODUCT
End Property

Private Function fnOpenFMTaxRecord(sSchedule As String, _
                         ByVal nTypeCode As String, _
                         sTaxClass As String, _
                         nProdRefCode As String, _
                         Optional vExtractFrom) As Collection

    Const SUB_NAME = "fnOpenFMTaxRecord"

    Dim strSQL As String
    Dim rsTemp As Recordset
    Dim ciTemp As ETFMTaxRecord
    Dim sInfo1 As String
    Dim sInfo2 As String
    Dim sInfo3 As String
    Dim sTemp As String
    Dim colTemp As Collection
    Dim sType As String

    strSQL = "DROP TABLE fm_temp1"
    fnExecuteSQL strSQL, SUB_NAME
    strSQL = "DROP TABLE fm_temp2"
    fnExecuteSQL strSQL, SUB_NAME
    
    If m_sEDIType = EDI_TYPE_R Then
        sType = "B"
        
        strSQL = "SELECT fmm_movement, fmm_carrier, fml_bl_nbr, fml_vendor, fml_terminal," _
            & " fml_int_ext, fml_date FROM fm_lading, fm_movement WHERE fml_date" _
            & " BETWEEN " & tfnDateString(m_sStartDate, True) & " AND " _
            & tfnDateString(m_sEndDate, True) & " AND fmm_dispatch_flag = 'C'" _
            & " AND fml_movement = fmm_movement" & " INTO TEMP fm_temp1"
        
        If fnExecuteSQL(strSQL, SUB_NAME, "") < 0 Then
            m_nErrorNumber = EC_CREATE_TT1_R
            Exit Function
        End If
        
        If fnCheckExitCancel Then
            Exit Function
        End If
        
        strSQL = "SELECT fmm_movement, fmm_carrier, fmldt_bl_nbr, fmldt_line, fmldt_trn," _
            & " fml_int_ext, fml_vendor, fml_terminal, fml_date" _
            & " FROM fm_lad_tax, tx_use_link, fm_temp1" _
            & " WHERE fmldt_movement = fmm_movement" _
            & " AND fmldt_bl_nbr = fml_bl_nbr" _
            & " AND fmldt_trn = tul_trn" _
            & " AND fmldt_ulink = tul_link" _
            & " AND (" & fnAppenCriteria("tul_sched", sSchedule) _
            & ") AND fmldt_type = '" & sType & "'"
        
        If m_nTPTCount >= 0 Then
            Select Case nTypeCode
            Case "1"
                If Not IsMissing(vExtractFrom) Then
                    strSQL = strSQL & " AND (fmldt_trn IN " & _
                        fnTrnNumbers(m_lTaxPaidTrn, m_nTPTCount) & ")"
                End If
            Case "3"
                strSQL = strSQL & " AND (fmldt_trn NOT IN " & _
                    fnTrnNumbers(m_lTaxPaidTrn, m_nTPTCount) & ")"
            End Select
        End If
        
        strSQL = strSQL & " INTO TEMP fm_temp2"
        
        If fnExecuteSQL(strSQL, SUB_NAME, "") < 0 Then
            m_nErrorNumber = EC_CREATE_TT2_R
            Exit Function
        End If
    
        If fnCheckExitCancel Then
            Exit Function
        End If
    
        strSQL = "SELECT fmm_movement, fmldt_bl_nbr, fmm_carrier, fml_int_ext, fml_vendor," _
            & " fml_terminal, fml_date, fmmb_bl_line, ivh_link, ivh_prodtcl, fmldt_trn," _
            & " fmd_dest_code, fmd_prft_ctr, fmd_dest_flag, fmmb_dv_vol namount" _
            & " FROM fm_temp2, fm_mv_bal, fm_delivery, inv_header"
        strSQL = strSQL & " WHERE fmldt_line = fmmb_bl_line" _
            & " AND fmldt_bl_nbr = fmmb_bl_nbr" _
            & " AND fmm_movement = fmmb_movement" _
            & " AND fmd_invoice= fmmb_dv_nbr" _
            & " AND fmd_movement = fmmb_movement" _
            & " AND fmmb_bl_prodlnk = ivh_link" _
            & " AND (" & fnAppenCriteria("ivh_prodtcl", sTaxClass) & ")"

         strSQL = strSQL & " ORDER BY fmldt_bl_nbr"

        m_nErrorNumber = EC_GET_RECORD_R
    Else
        sType = "I"
        
        strSQL = "SELECT fmm_movement, fmm_carrier, fmd_invoice, fmd_dest_code," _
            & " fmd_dest_flag, fmd_prft_ctr, fmd_date" _
            & " FROM fm_delivery, fm_movement WHERE fmd_date" _
            & " BETWEEN " & tfnDateString(m_sStartDate, True) & " AND " _
            & tfnDateString(m_sEndDate, True) _
            & " AND fmm_dispatch_flag = 'C'" _
            & " AND fmd_movement = fmm_movement"
        
        'for export, consider both Internal and External Delivery!
        'david 08/27/99
        If nTypeCode <> "7" Then
            strSQL = strSQL & " AND fmd_dest_flag = 'E'"
        End If
        
        strSQL = strSQL & " INTO TEMP fm_temp1"
        
        If fnExecuteSQL(strSQL, SUB_NAME, "") < 0 Then
            m_nErrorNumber = EC_CREATE_TT1_D
            Exit Function
        End If

        If fnCheckExitCancel Then
            Exit Function
        End If

        'for Export the criteria fmldt_bl_nbr = fmd_invoice MAY NOT VALID!!!
        'david 08/18/99
        'for Others Schedule DO IT, becuase it work before
        'david 08/27/99
        If nTypeCode <> "7" Then
            strSQL = "SELECT fmm_movement, fmm_carrier, fmldt_bl_nbr, fmldt_line," _
                & " fmldt_trn, fmd_dest_code, fmd_dest_flag, fmd_prft_ctr, fmd_date" _
                & " FROM fm_lad_tax, tx_use_link, fm_temp1" _
                & " WHERE fmldt_movement = fmm_movement" _
                & " AND fmldt_bl_nbr = fmd_invoice" _
                & " AND fmldt_trn = tul_trn" _
                & " AND fmldt_ulink = tul_link" _
                & " AND (" & fnAppenCriteria("tul_sched", sSchedule) _
                & ") AND fmldt_type = '" & sType & "'" _
                & " INTO TEMP fm_temp2"
        Else
            strSQL = "SELECT fmm_movement, fmm_carrier, fmldt_bl_nbr, fmldt_line, fmldt_trn, " _
                & " fmd_invoice, fmd_dest_code, fmd_dest_flag, fmd_prft_ctr, fmd_date" _
                & " FROM fm_lad_tax, tx_use_link, fm_temp1" _
                & " WHERE fmldt_movement = fmm_movement" _
                & " AND fmldt_trn = tul_trn" _
                & " AND fmldt_ulink = tul_link"
                strSQL = strSQL & " AND (" & fnAppenCriteria("tul_sched", sSchedule) _
                & ") INTO TEMP fm_temp2"
        End If
        
        If fnExecuteSQL(strSQL, SUB_NAME, "") < 0 Then
            m_nErrorNumber = EC_CREATE_TT2_D
            Exit Function
        End If
        
        If fnCheckExitCancel Then
            Exit Function
        End If
        
        If nTypeCode <> "7" Then
            strSQL = "SELECT fmm_movement, fmldt_bl_nbr, fmm_carrier, fmd_dest_code," _
                & " fmd_dest_flag, fmd_prft_ctr, fmd_date, fmldt_trn, fmmb_dv_line, ivh_link," _
                & " ivh_prodtcl, fml_int_ext, fml_vendor, fml_terminal, fmmb_dv_vol namount" _
                & " FROM fm_temp2, fm_mv_bal, fm_lading, inv_header" _
                & " WHERE fmm_movement = fmmb_movement" _
                & " AND fmldt_bl_nbr = fmmb_dv_nbr" _
                & " AND fmldt_line = fmmb_dv_line"
            strSQL = strSQL & " AND fml_bl_nbr = fmmb_bl_nbr" _
                & " AND fml_movement = fmmb_movement" _
                & " AND fmmb_dv_prodlnk = ivh_link" _
                & " AND (" & fnAppenCriteria("ivh_prodtcl", sTaxClass) & ")"
        Else
            strSQL = "SELECT DISTINCT fmm_movement, fmldt_bl_nbr, fmm_carrier, fmd_dest_code," _
                & " fmd_dest_flag, fmd_prft_ctr, fmd_date, fmldt_trn, fmmb_dv_line, ivh_link," _
                & " ivh_prodtcl, fml_int_ext, fml_vendor, fml_terminal, fmmb_dv_vol namount" _
                & " FROM fm_temp2, fm_mv_bal, fm_lading, inv_header" _
                & " WHERE fmm_movement = fmmb_movement" _
                & " AND fmldt_bl_nbr = fmmb_bl_nbr" _
                & " AND fmd_invoice = fmmb_dv_nbr"
            strSQL = strSQL & " AND fml_bl_nbr = fmmb_bl_nbr" _
                & " AND fml_movement = fmmb_movement" _
                & " AND fmmb_dv_prodlnk = ivh_link" _
                & " AND (" & fnAppenCriteria("ivh_prodtcl", sTaxClass) & ")"
        End If
        
        strSQL = strSQL & " ORDER BY fmldt_bl_nbr"
        m_nErrorNumber = EC_GET_RECORD_R
    End If
    
    Set rsTemp = fnOpenRecord(strSQL, SUB_NAME, "", , False)
    If rsTemp Is Nothing Then
        Exit Function
    End If
    
    If fnCheckExitCancel Then
        Exit Function
    End If
    
    m_nErrorNumber = EC_INITIAL
    
    With rsTemp
        If .RecordCount > 0 Then
            Set colTemp = New Collection
            
            '.MoveFirst
            While Not .EOF
                Set ciTemp = New ETFMTaxRecord
                ciTemp.TaxScheduleCode = sSchedule
                ciTemp.ProductCode = fnProductCode(tfnRound(!ivh_link))
                sTemp = fnCStr(!fmm_carrier)
                subCarrierInfo sInfo1, sInfo2, sTemp
                ciTemp.CarrierCode = sTemp
                ciTemp.CarrierName = sInfo1
                ciTemp.CarrierFedID = fnConvertFEDID(sInfo2)
                ciTemp.TTTrnNumber = tfnRound(!fmldt_trn)
                sTemp = fnTerminalCity(tfnRound(!fml_terminal), rsTemp!fml_int_ext)
                subGetCityInfo sInfo1, sInfo2, sTemp
                ciTemp.OriginCity = sInfo1
                ciTemp.OriginState = sInfo2
                sTemp = fnGetDestCity(tfnRound(!fmd_dest_code), !fmd_dest_flag)
                If sTemp = "" Then
                    sTemp = fnProfitCenterCity(tfnRound(!fmd_prft_ctr))
                End If
                subGetCityInfo sInfo1, sInfo2, sTemp
                ciTemp.DestCity = sInfo1
                ciTemp.DestState = sInfo2
                If m_sEDIType = EDI_TYPE_R Then
                    'Lading, get vendor info
                    sTemp = fnCStr(!fml_vendor)
                    If Val(sTemp) = 43040004 Then
'                        Stop
                    End If
                    
                    If sTemp <> "" Then
                        'get vendor info
                        subVendorInfo sInfo1, sInfo2, sInfo3, sTemp
                        ciTemp.BuySellCode = sTemp
                        ciTemp.BuySellName = sInfo1
                        ciTemp.BuySellFedID = fnConvertFEDID(sInfo3)
                    
                        'subGetCityInfo sInfo1, sInfo3, sInfo2
                        sInfo2 = fnDestState(fnCStr(!fmd_dest_code), fnCStr(!fmd_dest_flag))
                        sInfo3 = fnCStr(!ivh_prodtcl)
                        sTemp = fnVendorTaxID(tfnRound(!fml_vendor), sInfo2, sInfo3)
                        If sTemp <> "" Then
                            ciTemp.BuySellFedID = sTemp
                        End If
                    Else
                        'get customer info
                        sTemp = fnCStr(!fmd_prft_ctr)
                        subCustomerInfo sInfo1, sInfo2, sInfo3, sTemp, "I"
                        ciTemp.BuySellCode = sTemp
                        ciTemp.BuySellName = sInfo1
                        ciTemp.BuySellFedID = fnCustomerTaxID(Val(sTemp), tfnRound(!fmldt_trn), _
                            fnCStr(!fmd_dest_flag))
                        If ciTemp.BuySellFedID = "" Then
                            ciTemp.BuySellFedID = m_sCarrier_FEIN  '"000000000"
                        End If
                    End If
                    
                    ciTemp.LadingDate = tfnFormatDate(!fml_date)
                Else  'EDI_TYPE_D
                    'Delivery, get customer info
                    sTemp = fnCStr(!fmd_dest_code)
                    subCustomerInfo sInfo1, sInfo2, sInfo3, sTemp, fnCStr(!fmd_dest_flag)
                    ciTemp.BuySellCode = sTemp
                    ciTemp.BuySellName = sInfo1
                    ciTemp.BuySellFedID = fnCustomerTaxID(Val(sTemp), tfnRound(!fmldt_trn), _
                        fnCStr(!fmd_dest_flag))
                    ciTemp.LadingDate = tfnFormatDate(!fmd_date)
                End If
                ciTemp.DetailType = m_sEDIType
                ciTemp.LadingNumber = tfnRound(!fmldt_bl_nbr)
                ciTemp.Quantity = tfnRound(!namount)
                ciTemp.Movement = tfnRound(!fmm_movement)
                
                colTemp.Add ciTemp
                
                .MoveNext
            
                If fnCheckExitCancel Then
                    Exit Function
                End If
            Wend
        End If
    End With
    
    'retrieve wholesale fuel taxes
    If m_sEDIType = EDI_TYPE_D Then
        strSQL = "DROP TABLE fm_temp1"
        fnExecuteSQL strSQL, SUB_NAME
        strSQL = "DROP TABLE fm_temp2"
        fnExecuteSQL strSQL, SUB_NAME
        
        strSQL = "SELECT ih_nbr, ih_cust, ih_prft_ctr, ih_shipto, ih_ship_date, idp_seq_nbr," _
               & " idp_prodlnk FROM si_invoice, si_inv_line, inv_header WHERE" _
               & " ih_nbr = idp_nbr AND idp_prodlnk = ivh_link" _
               & " AND (" & fnAppenCriteria("ivh_prodtcl", sTaxClass) & ")" _
               & " AND ih_ship_date BETWEEN " & tfnDateString(m_sStartDate, True) _
               & " AND " & tfnDateString(m_sEndDate, True) & " INTO TEMP fm_temp1"
        
        If fnExecuteSQL(strSQL, SUB_NAME, "") < 0 Then
            m_nErrorNumber = EC_CREATE_TT1_D
            Exit Function
        End If
        
        If fnCheckExitCancel Then
            Exit Function
        End If
        
        strSQL = "SELECT fm_temp1.*, ilt_trn, ilt_units" _
               & " FROM fm_temp1, tx_use_link, si_inv_tax" _
               & " WHERE ih_nbr = ilt_nbr" _
               & " AND idp_seq_nbr = ilt_line" _
               & " AND ilt_trn = tul_trn" _
               & " AND ilt_ulink = tul_link" _
               & " AND (" & fnAppenCriteria("tul_sched", sSchedule) _
               & ") INTO TEMP fm_temp2"
        If fnExecuteSQL(strSQL, SUB_NAME, "") < 0 Then
            m_nErrorNumber = EC_CREATE_TT2_D
            Exit Function
        End If
        
        If fnCheckExitCancel Then
            Exit Function
        End If
        
        strSQL = "SELECT * FROM fm_temp2 ORDER BY ih_nbr"
        m_nErrorNumber = EC_GET_RECORD_R
        
        Set rsTemp = fnOpenRecord(strSQL, SUB_NAME, "", , False)
        If rsTemp Is Nothing Then
            Exit Function
        End If
        
        m_nErrorNumber = EC_INITIAL
        
        With rsTemp
            If .RecordCount > 0 Then
                If colTemp Is Nothing Then
                    Set colTemp = New Collection
                End If
                
                '.MoveFirst
                While Not .EOF
                    Set ciTemp = New ETFMTaxRecord
                    ciTemp.TaxScheduleCode = sSchedule
                    ciTemp.ProductCode = fnProductCode(tfnRound(!idp_prodlnk))
                    'sTemp = fnCStr(!fmm_carrier)
                    'subCarrierInfo sInfo1, sInfo2, sTemp
                    ciTemp.CarrierCode = ""  'sTemp
                    'ciTemp.CarrierName = ""  'sInfo1
                    'ciTemp.CarrierFedID = "000000000"  'fnConvertFEDID(sInfo2)
                    ciTemp.CarrierName = m_sCarrier_Name
                    ciTemp.CarrierFedID = m_sCarrier_FEIN
                    'sTemp = fnTerminalCity(tfnRound(!fml_terminal))
                    sTemp = fnProfitCenterCity(tfnRound(!ih_prft_ctr))
                    subGetCityInfo sInfo1, sInfo2, sTemp
                    ciTemp.OriginCity = sInfo1
                    ciTemp.OriginState = sInfo2
                    sTemp = fnGetDestCity(tfnRound(!ih_cust), "E")
                    subGetCityInfo sInfo1, sInfo2, sTemp
                    ciTemp.DestCity = sInfo1
                    ciTemp.DestState = sInfo2
                    
                    'wholesale, get customer info
                    sTemp = fnCStr(!ih_cust)
                    subCustomerInfo sInfo1, sInfo2, sInfo3, sTemp, "E"
                    ciTemp.BuySellCode = sTemp
                    ciTemp.BuySellName = sInfo1
                    ciTemp.BuySellFedID = fnCustomerTaxID(Val(sTemp), tfnRound(!ilt_trn), "E")
                    If ciTemp.BuySellFedID = "" Then
                        ciTemp.BuySellFedID = m_sCarrier_FEIN  '"000000000"
                    End If
                    ciTemp.LadingDate = tfnFormatDate(!ih_ship_date)
                    
                    ciTemp.DetailType = m_sEDIType
                    ciTemp.LadingNumber = tfnRound(!ih_nbr)
                    ciTemp.Quantity = tfnRound(!ilt_units)
                    ciTemp.Movement = tfnRound(!ih_nbr)
                    colTemp.Add ciTemp
                    .MoveNext
                
                    If fnCheckExitCancel Then
                        Exit Function
                    End If
                Wend
            End If
        End With
    End If
    
    Set fnOpenFMTaxRecord = colTemp
    
End Function

Private Function fnDestState(sDest As String, _
                             sEI As String) As String

'    fnDestState = REPORT_STATE
'    Exit Function
    
    Dim sInfo1 As String
    Dim sInfo2 As String
    Dim sInfo3 As String
    
    subCustomerInfo sInfo1, sInfo2, sInfo3, sDest, sEI
    subGetCityInfo sInfo1, sInfo3, sInfo2
    fnDestState = sInfo3
End Function

Private Function fnProductCode(ByVal lLink As Long) As String

    Const SUB_NAME = "fnProductCode"
    
    Dim strSQL As String
    
    strSQL = "SELECT ivh_product msname FROM inv_header" _
           & " WHERE ivh_link = " & CStr(lLink)
    fnProductCode = fnGetName(aryProduct, nProductCount, strSQL, lLink)
    
End Function

Public Sub Reset()

End Sub

Property Let StartDate(sTemp As String)
    m_sStartDate = tfnFormatDate(sTemp)
End Property

Private Sub subCheckRequiredFields()

    If Not IsDate(m_sStartDate) Then
        m_nErrorNumber = EC_INVALID_START_DATE
    End If
    If Not IsDate(m_sEndDate) Then
        m_nErrorNumber = EC_INVALID_END_DATE
    End If
    If Trim(m_sTheirID) = "" Then
        m_nErrorNumber = EC_INVALID_THIER_ID
    End If
    If Trim(m_sTaxpayerLicense) = "" Then
        m_nErrorNumber = EC_INVALID_TAX_LICENSE
    End If
    If Trim(m_sMode) <> MODE_TEST And Trim(m_sMode) <> MODE_PRODUCT Then
        m_nErrorNumber = EC_INVALID_MODE
    End If
End Sub

Private Sub subInitialParms()
    Const SUB_NAME = "subInitialParms"
    
    Dim strSQL As String
    Dim rsTemp As Recordset
    
    strSQL = "SELECT parm_field FROM sys_parm WHERE parm_nbr = " & SYSPARM_FED_ID
    Set rsTemp = fnOpenRecord(strSQL, SUB_NAME, "", , False)
    If rsTemp Is Nothing Then
        m_nErrorNumber = EC_INIT_PARM1_1
        Exit Sub
    End If
    If rsTemp.RecordCount = 0 Then
        m_nErrorNumber = EC_INIT_PARM1_2
        Exit Sub
    End If
    m_sTaxpayerID = fnConvertFEDID(fnCStr(rsTemp!parm_field))
    rsTemp.Close
    
    strSQL = "SELECT con_name FROM co_company_name"
    Set rsTemp = fnOpenRecord(strSQL, SUB_NAME, "", , False)
    If rsTemp Is Nothing Then
        m_nErrorNumber = EC_INIT_PARM2_1
        Exit Sub
    End If
    If rsTemp.RecordCount = 0 Then
        m_nErrorNumber = EC_INIT_PARM2_2
        Exit Sub
    End If
    m_sCompany = fnCStr(rsTemp!con_name)
    rsTemp.Close
    
End Sub

Private Sub subParseString(sParam() As String, _
                           sSrc As String, _
                           Optional vStart As Variant, _
                           Optional vEnd As Variant)
                          
    If Trim(sSrc) = "" Then
        Exit Sub
    End If

    Const nArrayInc As Integer = 5
    Dim i1 As Integer
    Dim i2 As Integer
    Dim k As Integer
    Dim nEnd As Integer
    If IsMissing(vStart) Then
        i1 = 1
    Else
        i1 = vStart
    End If
    If IsMissing(vEnd) Then
        nEnd = Len(sSrc)
    Else
        nEnd = vEnd
    End If
    If i1 < 1 Then i1 = 1
    i2 = 1
    k = 0
    ReDim sParam(nArrayInc)
    While i1 <= nEnd And i2 > 0 And i2 <= nEnd
        i2 = InStr(i1, sSrc, ",")
        If i2 >= i1 And i2 <= nEnd Then
            If k > UBound(sParam) Then
                ReDim Preserve sParam(k + nArrayInc)
            End If
            sParam(k) = Trim$(Mid$(sSrc, i1, i2 - i1))
            k = k + 1
            i1 = i2 + 1
        End If
    Wend
    If i2 <= nEnd Then
        If k > UBound(sParam) Then
            ReDim Preserve sParam(k + nArrayInc)
        End If
        sParam(k) = Trim$(Mid$(sSrc, i1, nEnd - i1 + 1))
        ReDim Preserve sParam(k)
    Else
        If k > 0 Then
            sParam(k - 1) = Trim$(Mid$(sSrc, i1, nEnd - i1 + 1))
            ReDim Preserve sParam(k - 1)
        End If
    End If
End Sub

Private Function fnAppenCriteria(sField As String, _
                                 sCriteria As String) As String
    
    Const sERR_CODE  As String = "ERR"
    
    Dim sTemp As String, sPosi As Integer, aryCodes() As String, aryOp() As String
    Dim nPosi As Integer, I As Integer, j As Integer, aryCriteria() As String
    
    'parse the code
    sTemp = sCriteria
    subParseString aryCriteria, sTemp
    
    fnAppenCriteria = ""
    
    If UBound(aryCriteria) = 0 Then
        If aryCriteria(0) = "" Then
            Exit Function
        End If
    End If
    
    For j = 0 To UBound(aryCriteria)
        If aryCriteria(j) <> "" Then
            nPosi = InStr(aryCriteria(j), sRANGE)
            If nPosi > 0 Then
                If nPosi >= Len(aryCriteria(j)) Then
                    MsgBox "Syntax error in Code " + tfnSQLString(aryCriteria(j)) + "."
                    fnAppenCriteria = ""
                    Exit Function
                End If
                
                ReDim aryCodes(1)
                ReDim aryCodes(1)
                
                aryCodes(0) = Left(aryCriteria(j), nPosi - 1)
                aryOp(0) = fnCheckSyntax(aryCodes(0), sRANGE)
                If aryOp(0) = sERR_CODE Then
                    MsgBox "Syntax error in Code " + _
                        tfnSQLString(aryCodes(0)) + " of " + tfnSQLString(aryCriteria(j)) + "."
                    fnAppenCriteria = ""
                    Exit Function
                End If
                aryCodes(1) = Mid(aryCriteria(j), nPosi + 1)
                aryOp(1) = fnCheckSyntax(aryCodes(1), sRANGE)
                If aryOp(1) = sERR_CODE Then
                    MsgBox "Syntax error in Code " + _
                        tfnSQLString(aryCodes(1)) + " of " + tfnSQLString(aryCriteria(j)) + "."
                    fnAppenCriteria = ""
                    Exit Function
                End If
                
                fnAppenCriteria = fnAppenCriteria + "(" + sField + " BETWEEN " + _
                    tfnSQLString(aryCodes(0)) + " AND " + tfnSQLString(aryCodes(1)) + ") OR "
            Else
                ReDim aryCodes(0)
                ReDim aryOp(0)
                
                aryCodes(0) = aryCriteria(j)
                aryOp(0) = fnCheckSyntax(aryCriteria(j), "")
                If aryOp(0) = sERR_CODE Then
                    MsgBox "Syntax error in Code " + _
                        tfnSQLString(aryCodes(0)) + "."
                    fnAppenCriteria = ""
                    Exit Function
                End If
                
                fnAppenCriteria = fnAppenCriteria + "(" + sField + " " + aryOp(I) + _
                    " " + tfnSQLString(aryCodes(0)) + ") OR "
            End If
        End If
    Next j

    If Right(fnAppenCriteria, 3) = "OR " Then
        fnAppenCriteria = Mid(fnAppenCriteria, 1, Len(fnAppenCriteria) - 3)
    End If
    
End Function

'return operator in where clause or 'ERR' string
Private Function fnCheckSyntax(sCode As String, sOp As String) As String
    Const sERR_CODE  As String = "ERR"
    
    Dim nPosi As Integer
    
    fnCheckSyntax = ""
    
    sCode = Trim(sCode)
    
    If sCode = "" Then Exit Function
    
    fnCheckSyntax = sERR_CODE
    
    If InStr(sCode, sRANGE) > 0 Or InStr(sCode, sOR) > 0 Then Exit Function
    
    If InStr(sCode, "?") > 0 Or InStr(sCode, "*") > 0 Then
        If sOp = sRANGE Then
            Exit Function
        End If
        
        If InStr(sCode, "%") > 0 Then
            Exit Function
        End If
        
        fnCheckSyntax = "MATCHES"
    Else
        If Right(sCode, 1) = "%" Then
            If InStr(sCode, "?") > 0 Or InStr(sCode, "*") > 0 Then
                Exit Function
            End If
        
            fnCheckSyntax = "LIKE"
        Else
            
            If InStr(sCode, "%") > 0 Then
                fnCheckSyntax = "LIKE"
            Else
                fnCheckSyntax = "="
            End If
        End If
    End If
End Function

Private Function fnGetName(aryStore() As tpLinkName, _
                           nCount As Integer, _
                           strSQL As String, _
                           ByVal lLink As Long) As String

    Const SUB_NAME = "fnGetName"
    Dim I As Integer
    
    For I = 0 To nCount - 1
        If lLink = aryStore(I).m_lLink Then
            fnGetName = aryStore(I).m_sName
            Exit Function
        End If
    Next I
    
    Dim rsTemp As Recordset
    
    Set rsTemp = fnOpenRecord(strSQL, SUB_NAME, "", , False)
    If Not rsTemp Is Nothing Then
        If rsTemp.RecordCount > 0 Then
            If nCount = 0 Then
                ReDim aryStore(ARRAY_INC)
            Else
                If UBound(aryStore) < nCount Then
                    ReDim Preserve aryStore(nCount + ARRAY_INC)
                End If
            End If
            aryStore(nCount).m_lLink = lLink
            aryStore(nCount).m_sName = fnCStr(rsTemp!msname)
            fnGetName = aryStore(nCount).m_sName
            nCount = nCount + 1
        End If
    End If
    
End Function


Private Function fnTerminalCity(ByVal nTerminal As Integer, sIntExtFlag As String) As String
    
    Dim strSQL As String
    
    If sIntExtFlag = "I" Then
        fnTerminalCity = fnProfitCenterCity(nTerminal)
    Else
        strSQL = "SELECT fmt_city msname FROM fm_terminal" _
               & " WHERE fmt_code = " & CStr(nTerminal)
    
        fnTerminalCity = fnGetName(aryTerminal, nTerminalCount, strSQL, nTerminal)
    End If
End Function

Private Sub subVendorInfo(sName As String, _
                          sCity As String, _
                          sFedId As String, _
                          sVendor As String)

    Dim strSQL As String
    Dim nIdx As Integer
    
    strSQL = "SELECT pn_name info1, pn_city info2, pm_fed_id info3 FROM p_altname, p_vendor" _
           & " WHERE pn_alt = " & sVendor _
           & " AND pm_vendor = pn_vendor "
    nIdx = fnGetInfoLink(aryVendor, nVendorCount, strSQL, sVendor)
    If nIdx >= 0 Then
        sName = aryVendor(nIdx).m_sInfo1
        sCity = aryVendor(nIdx).m_sInfo2
        sFedId = aryVendor(nIdx).m_sInfo3
    End If
    
End Sub

Private Sub subCarrierInfo(sName As String, _
                           sFedId As String, _
                           sCarrier As String)
    
    Dim strSQL As String
    Dim nIdx As Integer
    
    strSQL = "SELECT pn_name info1, pm_fed_id info2, ' ' info3 FROM p_altname, p_vendor" _
           & " WHERE pn_vendor = " & sCarrier _
           & " AND pm_vendor = pn_vendor "

    nIdx = fnGetInfoLink(aryCarrier, nCarrierCount, strSQL, sCarrier)
    If nIdx >= 0 Then
        sName = aryCarrier(nIdx).m_sInfo1
        sFedId = aryCarrier(nIdx).m_sInfo2
    End If
    
End Sub


Private Sub subCustomerInfo(sName As String, _
                            sCity As String, _
                            sFedId As String, _
                            sCustomer As String, _
                            sEI As String)
    
    Dim strSQL As String
    Dim nIdx As Integer
    
    If sEI = "E" Then
        strSQL = "SELECT an_name info1, an_city info2, ' ' info3 FROM ar_altname" _
               & " WHERE an_customer = " & sCustomer
    Else
        strSQL = "SELECT prft_name info1, prft_city info2, ' ' info3 FROM sys_prft_ctr" _
               & " WHERE prft_ctr = " & sCustomer
    End If
    nIdx = fnGetInfoLink(aryCustomer, nCustomerCount, strSQL, sCustomer)
    If nIdx >= 0 Then
        sName = aryCustomer(nIdx).m_sInfo1
        sCity = aryCustomer(nIdx).m_sInfo2
        sFedId = aryCustomer(nIdx).m_sInfo3
    End If
End Sub

Private Function fnProfitCenterCity(ByVal nProfitCenter As Integer) As String
    
    Dim strSQL As String
    
    strSQL = "SELECT prft_city msname FROM sys_prft_ctr" _
           & " WHERE prft_ctr = " & CStr(nProfitCenter)

    fnProfitCenterCity = fnGetName(aryProfitCenter, nProfitCenterCount, strSQL, nProfitCenter)
End Function

Private Function fnGetDestCity(ByVal fmd_dest_code As Long, fmd_dest_flag As Variant) As String
    Dim sDestFlag As String
    Dim strSQL As String
    
    fnGetDestCity = ""
    
    On Error Resume Next
    sDestFlag = Trim(fmd_dest_flag)
    
    If sDestFlag = "I" Then
        fnGetDestCity = fnProfitCenterCity(fmd_dest_code)
    Else
        strSQL = "SELECT an_city msname FROM ar_altname" _
               & " WHERE an_customer = " & fmd_dest_code
    
        fnGetDestCity = fnGetName(aryCustomerCity, nCustomerCityCount, strSQL, fmd_dest_code)
    End If
End Function


Private Sub subGetCityInfo(sName As String, _
                           sState As String, _
                           sCode As String)

    Dim strSQL As String
    Dim nIdx As Integer
    
    strSQL = "SELECT city_name info1, city_state info2, ' ' info3 FROM sys_city" _
           & " WHERE city_code = " & tfnSQLString(sCode)

    nIdx = fnGetInfoLink(aryCity, nCityCount, strSQL, sCode)
    If nIdx >= 0 Then
        sName = aryCity(nIdx).m_sInfo1
        sState = aryCity(nIdx).m_sInfo2
    End If
    
End Sub

Private Function fnGetInfoLink(aryInfo() As tpLinkInfo, _
                               nCount As Integer, _
                               strSQL As String, _
                               sCode As String) As Integer

    Const SUB_NAME = "fnGetInfoLink"
    Dim I As Integer
    
    fnGetInfoLink = -1
    For I = 0 To nCount - 1
        If sCode = aryInfo(I).m_sLink Then
            fnGetInfoLink = I
            Exit Function
        End If
    Next I

    Dim rsTemp As Recordset
    
    Set rsTemp = fnOpenRecord(strSQL, SUB_NAME, "", , False)
    If Not rsTemp Is Nothing Then
        If rsTemp.RecordCount > 0 Then
            If nCount = 0 Then
                ReDim aryInfo(ARRAY_INC)
            Else
                If UBound(aryInfo) < nCount Then
                    ReDim Preserve aryInfo(nCount + ARRAY_INC)
                End If
            End If
            aryInfo(nCount).m_sLink = sCode
            aryInfo(nCount).m_sInfo1 = fnCStr(rsTemp!info1)
            aryInfo(nCount).m_sInfo2 = fnCStr(rsTemp!info2)
            aryInfo(nCount).m_sInfo3 = fnCStr(rsTemp!info3)
            fnGetInfoLink = nCount
            nCount = nCount + 1
        End If
    End If
    
End Function


Private Function fnVendorTaxID(ByVal lVendor As Long, _
                               sState As String, _
                               sTaxClass As String) As String

    Const SUB_NAME = "fnVendorTaxID"
    Dim I As Integer
    
    fnVendorTaxID = ""
    For I = 0 To nVTaxIDCount - 1
        If lVendor = aryVendorTaxID(I).m_lVendor Then
            If sState = aryVendorTaxID(I).m_sState Then
                If sTaxClass = aryVendorTaxID(I).m_sTaxClass Then
                    fnVendorTaxID = aryVendorTaxID(I).m_sTaxID
                    Exit Function
                End If
            End If
        End If
    Next I

    Dim strSQL As String
    Dim rsTemp As Recordset
    
    If nVTaxIDCount = 0 Then
        ReDim aryVendorTaxID(ARRAY_INC)
    Else
        If UBound(aryVendorTaxID) < nVTaxIDCount Then
            ReDim Preserve aryVendorTaxID(nVTaxIDCount + ARRAY_INC)
        End If
    End If
    aryVendorTaxID(nVTaxIDCount).m_lVendor = lVendor
    aryVendorTaxID(nVTaxIDCount).m_sState = sState
    aryVendorTaxID(nVTaxIDCount).m_sTaxClass = sTaxClass
    aryVendorTaxID(nVTaxIDCount).m_sTaxID = ""
    strSQL = "SELECT pvt_id_nbr FROM p_vend_txid" _
           & " WHERE pvt_vendor = " & CStr(lVendor) _
           & " AND pvt_state = '" & sState & "'" _
           & " AND pvt_tcl = '" & sTaxClass & "'"
    Set rsTemp = fnOpenRecord(strSQL, SUB_NAME, "", , False)
    If Not rsTemp Is Nothing Then
        If rsTemp.RecordCount > 0 Then
            aryVendorTaxID(nVTaxIDCount).m_sTaxID = fnConvertFEDID(fnCStr(rsTemp!pvt_id_nbr))
        End If
    End If
    fnVendorTaxID = aryVendorTaxID(nVTaxIDCount).m_sTaxID
    nVTaxIDCount = nVTaxIDCount + 1
    
End Function

Private Function fnCustomerTaxID(ByVal lCustomer As Long, _
                                 ByVal lTrn As Long, _
                                 sEI As String) As String

    Const SUB_NAME = "fnCustomerTaxID"
    Dim I As Integer
    
    fnCustomerTaxID = ""
    For I = 0 To nCTaxIDCount - 1
        If lCustomer = aryCustomerTaxID(I).m_lCustomer Then
            If lTrn = aryCustomerTaxID(I).m_lTrnLink Then
                fnCustomerTaxID = aryCustomerTaxID(I).m_sTaxID
                Exit Function
            End If
        End If
    Next I

    Dim strSQL As String
    Dim rsTemp As Recordset
    
    If nCTaxIDCount = 0 Then
        ReDim aryCustomerTaxID(ARRAY_INC)
    Else
        If UBound(aryCustomerTaxID) < nCTaxIDCount Then
            ReDim Preserve aryCustomerTaxID(nCTaxIDCount + ARRAY_INC)
        End If
    End If
    
    aryCustomerTaxID(nCTaxIDCount).m_lCustomer = lCustomer
    aryCustomerTaxID(nCTaxIDCount).m_lTrnLink = lTrn
    aryCustomerTaxID(nCTaxIDCount).m_sTaxID = ""
    
    If sEI = "E" Then
        strSQL = "SELECT txcp_permit FROM tx_cust_prmt" _
               & " WHERE txcp_customer = " & CStr(lCustomer) _
               & " AND txcp_trn = " & CStr(lTrn)
        
        Set rsTemp = fnOpenRecord(strSQL, SUB_NAME, "", , False)
        If Not rsTemp Is Nothing Then
            If rsTemp.RecordCount > 0 Then
                aryCustomerTaxID(nCTaxIDCount).m_sTaxID = fnConvertFEDID(fnCStr(rsTemp!txcp_permit))
            End If
        End If
    Else
        aryCustomerTaxID(nCTaxIDCount).m_sTaxID = fnConvertFEDID(fnGetSysParm(9008))
    End If
    
    fnCustomerTaxID = aryCustomerTaxID(nCTaxIDCount).m_sTaxID
    nCTaxIDCount = nCTaxIDCount + 1
    
End Function

Private Function fnExecuteSQL(strSQL As String, _
                             Optional vCaller As Variant, _
                             Optional vMsg As Variant, _
                             Optional vDB As Variant) As Integer
    Const SUB_NAME = "fnExecuteSQL"
    Dim objDB As Database
    
    If IsMissing(vDB) Then
        Set objDB = t_dbMainDatabase
    Else
        Set objDB = vDB
    End If
    On Error GoTo errExecute
    If objDB Is t_dbMainDatabase Then
        objDB.ExecuteSQL strSQL
    Else
        objDB.Execute strSQL
    End If
    fnExecuteSQL = objDB.RecordsAffected

    On Error GoTo 0
    Exit Function

errExecute:
    fnExecuteSQL = -1
    Dim bShow As Boolean
    bShow = Not IsMissing(vMsg)
    If IsMissing(vCaller) Then
        tfnErrHandler SUB_NAME, strSQL, bShow
    Else
        tfnErrHandler SUB_NAME & "," & CStr(vCaller), strSQL, bShow
    End If
    m_sLastVBError = fnGetVBError
End Function


Private Function fnOpenRecord(strSQL As String, _
                              Optional vCaller As Variant, _
                              Optional vMsg As Variant, _
                              Optional vDB As Variant, _
                              Optional vMoveLast) As Recordset
    Const SUB_NAME = "fnOpenRecord"
    ' Get records from the given SQL statement
    Dim objDB As Database
    Dim rsTemp As Recordset
    Dim bMoveLast As Boolean

    If IsMissing(vDB) Then
        Set objDB = t_dbMainDatabase
    Else
        Set objDB = vDB
    End If
    
    If IsMissing(vMoveLast) Then
        bMoveLast = True
    Else
        bMoveLast = vMoveLast
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
    
    Set fnOpenRecord = rsTemp

    On Error GoTo 0
    
    Exit Function

SQLError:
    Set fnOpenRecord = Nothing
    Dim bShow As Boolean
    bShow = Not IsMissing(vMsg)
    If IsMissing(vCaller) Then
        tfnErrHandler SUB_NAME, strSQL, bShow
    Else
        tfnErrHandler SUB_NAME & "," & CStr(vCaller), strSQL, bShow
    End If
    m_sLastVBError = fnGetVBError
End Function

Property Let TaxpayerLicense(sTemp As String)
    m_sTaxpayerLicense = sTemp
End Property

Property Let TaxTypeCode(sTemp As String)
    m_sTaxTypeCode = sTemp
End Property

Property Let TheirID(sTemp As String)
    m_sTheirID = sTemp
End Property


Private Sub Class_Initialize()
    subInitialParms
    m_nErrorNumber = EC_INITIAL
End Sub

Private Sub subCheckFEIN(colItems As Collection)
    Dim ciTemp As ETFMTaxRecord, sErrMsg As String
    Dim I As Integer
    
    sErrMsg = ""
    I = -1
    
    For Each ciTemp In colItems
        If Trim(ciTemp.CarrierFedID) = "" Then
            subAddErrFEIN sErrMsg, ciTemp.CarrierName, ciTemp.CarrierCode
        End If
        If Trim(ciTemp.BuySellFedID) = "" Then
            subAddErrFEIN sErrMsg, ciTemp.BuySellName, ciTemp.BuySellCode
        End If
    Next ciTemp
    
    If sErrMsg <> "" Then
        m_nErrorNumber = EC_NULL_FEIN
        m_sLastVBError = "FEIN for the following Carrier,Vendor or Customer is Blank:\" + _
            Left(sErrMsg, Len(sErrMsg) - 1)
    End If
End Sub

Private Sub subAddErrFEIN(sErrMsg As String, sName As String, sCode As String)
    On Error Resume Next
    If sErrMsg = "" Then
        sErrMsg = sErrMsg + sCode + IIf(sCode = "", "", "-") + sName + "|"
    Else
        If InStr(sErrMsg, sName) <= 0 Then
            sErrMsg = sErrMsg + sCode + IIf(sCode = "", "", "-") + sName + "|"
        End If
    End If
End Sub

Private Function fnCheckExitCancel() As Boolean
    DoEvents
    fnCheckExitCancel = True
    If bExitProgram Then
        m_nErrorNumber = EC_CANCEL_PROCESS
        Exit Function
    End If
    If bCancelProcess Then
        m_nErrorNumber = EC_EXIT_PROGRAM
        Exit Function
    End If
    fnCheckExitCancel = False
End Function

Property Let TaxPaidTaxCode(sCode As String)
    Const SUB_NAME = "TaxPaidTaxCode"
    
    Dim strSQL As String
    Dim rsTemp As Recordset
    
    m_nTPTCount = -1
    
    If sCode = "" Then
        Exit Property
    End If
    
    strSQL = "SELECT DISTINCT tt_trn FROM tx_table" _
           & " WHERE tt_level = 'S'" _
           & " AND (" & fnAppenCriteria("tt_code", sCode) & ")"

    Set rsTemp = fnOpenRecord(strSQL, SUB_NAME, "", , False)
    If rsTemp Is Nothing Then
        Exit Property
    End If
    
    While Not rsTemp.EOF
        m_nTPTCount = m_nTPTCount + 1
        If m_nTPTCount <= 20 Then
            m_lTaxPaidTrn(m_nTPTCount) = tfnRound(rsTemp!tt_trn)
        End If
        rsTemp.MoveNext
    Wend
End Property

Private Function fnTrnNumbers(lSource() As Long, _
                             ByVal nCount As Integer) As String

    Dim I As Integer
    Dim sTemp As String
    
    sTemp = lSource(0)
    For I = 1 To nCount
        sTemp = sTemp & "," & CStr(lSource(I))
    Next I
    fnTrnNumbers = "(" & sTemp & ")"
End Function



