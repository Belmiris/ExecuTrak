Attribute VB_Name = "modTemplate"
'***********************************************************'
'
' Copyright (c) 1996 FACTOR, A Division of W.R.Hess Company
'
' Module name   : TEMPLATE.BAS
'
' This module implements global template functions
'
' Functions:
'
Option Explicit

'=======================
'Global System Variables
'=======================

Global t_oleObject As Object         'pointer to the FACTOR Main Menu oleObject
Global t_szConnect As String         'This holds the ODBC connect string passed from oleObject
Global t_engFactor As DBEngine       'pointer to database engine
Global t_wsWorkSpace As Workspace    'pointer to the default workspace
Global t_dbMainDatabase As DataBase  'main database handle

Global CRLF As String 'carriage return linefeed string

Public Const DEBUG_LOG_PATH = "C:\FACTOR\TEMP\"
'##################################################
'# Added 10-30-01 Robert C Atwood
'##################################################
Public Const LOCAL_FACTOR_PATH = "C:\FACTOR\"

'**************************************************
'Constant for Help File name and Help Error message
'**************************************************
Global Const szHelpFileName As String = "FACTOR.HLP"
Global Const szHelpAdvCStore As String = "ADVCSTOR.HLP"
Global Const szHelpSysMgt As String = "SYSADMIN.HLP"
Global Const szHelpWhlSale As String = "WHOLSALE.HLP"
Global Const szHelpRetail As String = "RETAIL.HLP"
Global Const szHelpAdvFinancial As String = "ADVFIN.HLP"
Global Const szHelpAcctRec As String = "AR.HLP"
Global Const szHelpFuelMgt As String = "FUELMGT.HLP"
Global Const szHelpAcctPay As String = "AP.HLP"
Global Const szHelpOrderEntry As String = "ORDERENT.HLP"
Global Const szHelpGenLdgr As String = "GL.HLP"
Global Const szHelpService As String = "SERVICE.HLP"
Global Const szHelpPayroll As String = "PAYROLL.HLP"
Global Const szHelpDispatch As String = "FD.HLP"
Global Const szHelpFuelOil As String = "FO.HLP"
Global Const szHelpTax As String = "TAX.HLP"
Global Const szHelpEdiTaxFiling As String = "ET.HLP"
Global Const szHelpElecCommerce As String = "EC.HLP"
Global Const szHelpCMSystem As String = "CMS.HLP"
Global Const szHelpPO As String = "PO.HLP" ' Wenstrong, For Purchase Order
Global Const szHelpProfitTrak As String = "ProfitTrak.HLP" ' Wenstrong, For ProfitTrak
Global Const szHelpAPPALACH As String = "APPALACH.HLP" 'help file name for APPALACHIAN
Global Const szHelpReadyMix As String = "RM.HLP" 'help file name for Ready Mix
Global Const szHelp7_11 As String = "7-ELEVEN.HLP" 'help file name for 7-11 Commission Check custom
Global Const szHelpICUSTINQ As String = "ICUSTINQ.HLP" 'Internet project Customer inquiry
Global Const szHelpGasCheck As String = "ZZEFOGCK.HLP" 'Gas Check Data Entry
Global Const szHelpDrakeOil As String = "ZZEDPCLU.HLP" 'Card Lock Processing
Global Const szHelpSalesMark As String = "SaleMark.HLP" 'Sales and Marketing
Global Const szHelpDrakeOilFile As String = "ZZFDPEDT.HLP"  'Card Lock File Maintenance     'Vijaya 12/06/01...
Global Const szHelpZZFMURMT As String = "ZZFMURMT.HLP"  'Retail Sales Export File Maintenance     'Vijaya 12/12/01...
Global Const szHelpZZFPCAFM As String = "ZZFPCAFM.HLP"  'Retail Sales Export File Maintenance     'Vijaya 12/12/01...
Global Const szHelpTouchStar As String = "TCHSTAR.HLP"  'TocuhStar     'Vijaya 02/24/01...
Global Const szHelpMgntRpt As String = "MRPT.HLP"  'Management Reports File Maintenance     'Vijaya on 07/12/02

Public Const t_szEXIT_MESSAGE = "All changes will be lost! Do you want to exit anyway ?"
Public Const t_szCANCEL_MESSAGE = "All changes will be lost! Do you want to cancel anyway ?"
Public Const t_szREFRESH_MESSAGE = "All changes will be lost! Do you want to refresh anyway ?"

'=======================
'Global System Constants
'=======================

'string constants
Global Const t_szHandShake As String = "\factmenu" 'used to prevent application from launching without FACTMENU
Global Const t_szOLEObjectName As String = "FactMenu.clsRequest" 'registration database identifier
Public Const t_szOLECOMBO As String = "OleCombo.clsComboControl"
Public Const t_szOLEPIPELINE As String = "PPLINE32.Pipeline32"


'single character constants
Global Const szEMPTY As String = ""
Global Const szAMPERSAND As String = "&"
Global Const szDASH As String = "-"
Global Const szEQUAL As String = "="
Global Const szSEMICOLON As String = ";"
Global Const szCOLON As String = ":"
Global Const szSPACE As String = " "
Global Const szCOMMA As String = ","
Global Const szSLASH As String = "\"
Global Const szBANG As String = "!"

'keyBoard constants for the CAPS, NUM and SCROLL LOCK keys
Global Const VK_NUMLOCK As Integer = &H90
Global Const VK_SCROLL As Integer = &H91
Global Const VK_CAPITAL As Integer = &H14

'string constants
Global Const szCAPS As String = "CAPS" 'caps lock
Global Const szNUM As String = "NUM"   'num lock
Global Const szSCRL As String = "SCRL" 'scroll lock
Global Const szSKIP As String = "SKIP" 'TAG field used to skip resize
Global Const szODBC As String = "ODBC;" 'used to trigger ODBC Dialog Box during development
Global Const szCONNECTION_ERROR As String = "Database Connection Error" 'connect error message
Global Const szRUNEXE_ERROR As String = "Unable to Execute :" 'context menu message
Global Const szRUNEXE_TITLE As String = "Application Error"   'context menu error message box title
Global Const szRUN_ERROR As String = "This Application Must Be Run From The FACTMENU Program." 'invalid handshake message
Global Const szLOG_FILE_NAME As String = "FACTOR.LOG" 'log file for the application - used with tfnLog function
Global Const szFACTOR_INI As String = "FACTOR.INI" 'application INI filename

'resize event constants
Global Const WINDOW_STATE_MINIMIZED As Integer = vbMinimized
Global Const WINDOW_STATE_MAXIMIZED As Integer = vbMaximized

Global Const ELASTIC_COMMAND_BUTTON As Integer = 3
Global Const ELASTIC_CLASSIC As Integer = 0

Global Const MAX_STRING_LENGTH As Integer = 255 'used with fixed length strings - normally with windows api calls
Global Const MAX_INT As Long = 65535

Global Const RESIZE_PROPORTIONAL As Integer = 7 'used to lock the elastics in place - turned off during design time

Global Const ASC_RETURN As Integer = 13

'mouse cursor constants
Global Const DEFAULT_CURSOR As Integer = vbDefault      'Sets the default cursor
Global Const ARROW_CURSOR As Integer = vbArrow          'Arrow cursor index
Global Const HOURGLASS_CURSOR As Integer = vbHourglass  'Hourglass cursor index
Global Const NODROP_CURSOR As Integer = vbNoDrop        'nodrop cursor index

'status bar colors
Global Const ERROR_TEXT_COLOR As Long = &HFF&
Global Const CORRECT_TEXT_COLOR As Long = &H8000&
Global Const STANDARD_TEXT_COLOR As Long = &H0&

'DB access constants
Global Const DB_INCONSISTENT = dbInconsistent
Global Const SQL_PASSTHROUGH = dbSQLPassThrough
Global Const DBOPEN_SNAPSHOT = dbOpenSnapshot
Global Const DBOPEN_DYNASET = dbOpenDynaset
Global Const DBOPEN_TABLE = dbOpenTable
Global Const DBOPEN_READONLLY = dbReadOnly

'icon resource constants
Public Const WIN95_CRITICAL = 100
Public Const WIN95_STOP = 101
Public Const WIN31_INFORMATION = 102
Public Const WIN95_INFORMATION = 103
Public Const WIN31_QUESTION = 104
Public Const WIN95_QUESTION = 105
Public Const WIN31_EXCLAMATION = 106
Public Const WIN95_EXCLAMATION = 107

'toolbar resource constants
Public Const CANCEL_DOWN = 100
Public Const CANCEL_UP = 150
Public Const COPY_DOWN = 200
Public Const COPY_UP = 250
Public Const HELP_DOWN = 300
Public Const HELP_UP = 350
Public Const PREV_MENU_DOWN = 400
Public Const PREV_MENU_UP = 450
Public Const OK_DOWN = 500
Public Const OK_UP = 550
Public Const PRINT_DOWN = 600
Public Const PRINT_UP = 650
Public Const SEARCH_DOWN = 700
Public Const SEARCH_UP = 750
Public Const MAIN_MENU_DOWN = 800
Public Const MAIN_MENU_UP = 850
Public Const EXIT_DOWN = 900
Public Const EXIT_UP = 950
Public Const DROPDOWN_DOWN = 1000
Public Const DROPDOWN_UP = 1050
Public Const TAXCLASS_DOWN = 1100
Public Const TAXCLASS_UP = 1150
Public Const TAXTABLE_DOWN = 1200
Public Const TAXTABLE_UP = 1250
Public Const CITY_DOWN = 1300
Public Const CITY_UP = 1350
Public Const GL_DOWN = 1400
Public Const GL_UP = 1450
Public Const VENDOR_DOWN = 1500
Public Const VENDOR_UP = 1550
Public Const PRODUCT_DOWN = 1600
Public Const PRODUCT_UP = 1650
Public Const TAXUSE_DOWN = 1700
Public Const TAXUSE_UP = 1750
Public Const TERMS_DOWN = 1800
Public Const TERMS_UP = 1850
Public Const PROFITCENTER_DOWN = 1900
Public Const PROFITCENTER_UP = 1950
Public Const CATEGORY_DOWN = 2000
Public Const CATEGORY_UP = 2050
Public Const GROUP_DOWN = 2100
Public Const GROUP_UP = 2150
Public Const ITEM_DOWN = 2200
Public Const ITEM_UP = 2250
Public Const PRICECNG_DWN = 2300
Public Const PRICECNG_UP = 2350
Public Const FRMULA_DOWN = 2400
Public Const FRMULA_UP = 2450
Public Const GASPMP_DWN = 2500
Public Const GASPMP_UP = 2550
Public Const PRFTCNTR_DWN = PROFITCENTER_DOWN
Public Const PRFTCNTR_UP = PROFITCENTER_UP
Public Const RSTBRDNG_DWN = 2700
Public Const RSTBRDNG__UP = 2750
Public Const TANK_DOWN = 2800
Public Const TANK_UP = 2850
Public Const UOM_DWN = 2900
Public Const UOM_UP = 2950
Public Const GLPD_DWN = 3000
Public Const GLPD_UP = 3050
Public Const BOOK_DOWN = 3100
Public Const BOOK_UP = 3150
Public Const CHANGEPRODUCT_DOWN = 3200
Public Const CHANGEPRODUCT_UP = 3250
Public Const FUELTERMINAL_DOWN = 3300
Public Const FUELTERMINAL_UP = 3350
Public Const FREIGHTCODE_DOWN = 3400
Public Const FREIGHTCODE_UP = 3450
Public Const CUSTOMER_DOWN = 3500
Public Const CUSTOMER_UP = 3550
Public Const INVHDR_DOWN = 3600
Public Const INVHDR_UP = 3650
Public Const INVMSTR_DOWN = 3700
Public Const INVMSTR_UP = 3750
Public Const PRICECHNG_DOWN = 3800
Public Const PRICECHNG_UP = 3850
Public Const RESETREADING_DOWN = 3900
Public Const RESETREADING_UP = 3950
Public Const SHIP_ADDRESS_DOWN = 4000
Public Const SHIP_ADDRESS_UP = 4050
Public Const CUSTOMER_INFO_DOWN = 4100
Public Const CUSTOMER_INFO_UP = 4150
Public Const BEST_BUY_DOWN = 4200
Public Const BEST_BUY_UP = 4250
Public Const CONTROL_DOWN = 4300
Public Const CONTROL_UP = 4350
Public Const UOM_CNVRT_DOWN = 4400
Public Const UOM_CNVRT_UP = 4450
Public Const SLIDE_RIGHT_DOWN = 4500
Public Const SLIDE_RIGHT_UP = 4550
Public Const SLIDE_LEFT_DOWN = 4575
Public Const SLIDE_LEFT_UP = 4600
Public Const INVOICE_DATA_DOWN = 4650
Public Const INVOICE_DATA_UP = 4700
Public Const CLEAR_PRINTGP_DOWN = 4750
Public Const CLEAR_PRINTGP_UP = 4800
Public Const CHECK_SELECT_DOWN = 4850
Public Const CHECK_SELECT_UP = 4900
Public Const CUT_DOWN = 4950
Public Const CUT_UP = 5000
Public Const PASTE_DOWN = 5050
Public Const PASTE_UP = 5100
Public Const PRIORITY_DOWN = 5150
Public Const PRIORITY_UP = 5200
Public Const PROBLEM_DOWN = 5250
Public Const PROBLEM_UP = 5300
Public Const DEVICE_DOWN = 5350
Public Const DEVICE_UP = 5400
Public Const RESOURCE_DOWN = 5450
Public Const RESOURCE_UP = 5500
Public Const USER_TEST_DOWN = 5550
Public Const USER_TEST_UP = 5600
Public Const COMMENT_DOWN = 5650
Public Const COMMENT_UP = 5700
Public Const CARRIER_DOWN = 5750
Public Const CARRIER_UP = 5800
Public Const CALCULATOR_DOWN = 5850
Public Const CALCULATOR_UP = 5900
Public Const LOCATION_DOWN = 5950
Public Const LOCATION_UP = 6000
Public Const TRAILER_DOWN = 6050
Public Const TRAILER_UP = 6100
Public Const TRANS_TYPE_DOWN = 6150
Public Const TRANS_TYPE_UP = 6200
Public Const CUSTOMER_CLASS_DOWN = 6250
Public Const CUSTOMER_CLASS_UP = 6300
Public Const FINANCIAL_CHARGE_DOWN = 6350
Public Const FINANCIAL_CHARGE_UP = 6400
Public Const STATEMENT_CYCLE_DOWN = 6450
Public Const STATEMENT_CYCLE_UP = 6500
Public Const WO_CLASS_DOWN = 6550
Public Const WO_CLASS_UP = 6600
Public Const RESOURCE_TYPE_DOWN = 6650
Public Const RESOURCE_TYPE_UP = 6700
Public Const SKILL_CODE_DOWN = 6750
Public Const SKILL_CODE_UP = 6800
Public Const WO_CODE_DOWN = 6850
Public Const WO_CODE_UP = 6900
Public Const MFGNA_DOWN = 6950
Public Const MFGNA_UP = 7000
Public Const ADVANCED_DISPATCH_DOWN = 7050
Public Const ADVANCED_DISPATCH_UP = 7100
Public Const HEATING_CLASS_DOWN = 7150
Public Const HEATING_CLASS_UP = 7200
Public Const FO_TANK_DOWN = 7250
Public Const FO_TANK_UP = 7300
Public Const ADD_TEXT_DOWN = 7350
Public Const ADD_TEXT_UP = 7400
Public Const INSPECT_DOWN = 7450
Public Const INSPECT_UP = 7500
Public Const PRDCLS_DOWN = 7550
Public Const PRDCLS_UP = 7600
Public Const SNDBATCH_DOWN = 7650
Public Const SNDBATCH_UP = 7700
Public Const SNDSQL_DOWN = 7750
Public Const SNDSQL_UP = 7800
Public Const ORDER_DOWN = 7850 'WJ
Public Const ORDER_UP = 7900   'WJ
Public Const DSPTCH_DOWN = 7950 'WJ
Public Const DSPTCH_UP = 8000   'WJ
Public Const NOTES_DOWN = 8050 'WJ
Public Const NOTES_UP = 8100   'WJ
Public Const COMNT_DOWN = COMMENT_DOWN  'WJ
Public Const COMNT_UP = COMMENT_UP   'WJ
Public Const SOURCE_DEST_DOWN = 8250
Public Const SOURCE_DEST_UP = 8300
Public Const CSTOREPRD_DOWN = 8500
Public Const CSTOREPRD_UP = 8550
Public Const INVSETUP_DOWN = 8600
Public Const INVSETUP_UP = 8650
Public Const PRICEGRP_DOWN = 8700
Public Const PRICEGRP_UP = 8750
Public Const MARKETING_COMNT_UP = 8200
Public Const ORDER_INSTR_UP = 8400
Public Const ARTYPE_UP = 8850
Public Const EDI_SETUP_UP = 9050
Public Const REGION_UP = 9150
Public Const VW4GL_UP = 9250
Public Const EMP_MST_UP = 9350
Public Const VIEW_RELEASE_UP = 9450
Public Const PAY_CODE_UP = 9600
Public Const MOVEMENT_LOOKUP_UP = 9700
Public Const RSPURCH_UP = 9800
Public Const CHANGE_BOL_UP = 9900
Public Const AP_INVCEN_UP = 10050
Public Const DEGREE_DAY_UP = 10150
Public Const F_MOVEMENT_UP = 10250
Public Const COMPANY_UP = 10350
Public Const EFT_VEN_CRXF_UP = 10450
Public Const FO_FORCE_TICKET_UP = 10500
Public Const FO_QUEUE_TICKET_UP = 10600
Public Const PRINT_GROUP_UP = 10700
Public Const FO_SCH_DEL_UP = 10850
Public Const FO_SITE_UP = 10900
Public Const SM_CONTRACT_UP = 11000
Public Const PURCH_GROUP_UP = 11100
Public Const BUILD_MATERIAL = 11150
Public Const IQFACT_STREAM_UP = 11200
Public Const FO_HOLD_UP = 11250
Public Const CDPLAYER_UP = 11350
Public Const AREPAYCC_UP = 11400
Public Const SECURITY_UP = 11450
Public Const APFVOIDR_UP = 11500
Public Const SYS_LOCKS_UP = 11550
Public Const DTN_ENT_UP = 11600
Public Const ADJUST_UP = 11700
Public Const CC_MASTER_UP = 11750
Public Const WHL_RECVER_UP = 11800
Public Const POFBRSPO_UP = 11850
Public Const POEOENTR_UP = 11900
Public Const POFNOSTK_UP = 11950
Public Const POSORDER_UP = 12000
Public Const POFSELGP_UP = 12150
Public Const POFBRSPR_UP = 12200
Public Const POERENTR_UP = 12250
Public Const SMEWKORD_UP = 12300
Public Const POFVENDR_UP = 12350
Public Const POFAPLVL_UP = 12400
Public Const POFVNPRI_UP = 12450
Public Const CLOSEPO_UP = 12500
Public Const POAPPROV_UP = 12550
Public Const PRAPPROV_UP = 12600
Public Const PRPRINT_UP = 12650

'The following bitmaps need to be loaded into the form
'And define 2 public functions to supply the bitmaps and module names:
'1. Public Function GetPicture(ByVal nID As Integer) As Picture
'2. Public Function GetModuleName(ByVal nID As Integer) As String
Public Const IMPORT_UP = 12700
Public Const IMPORT_DOWN = 12750
Public Const EXPORT_UP = 12800
Public Const EXPORT_DOWN = 12850
Public Const TSTPRINT_UP = 12900
Public Const EDIVNDXR_UP = 13000
Public Const EDIUOMXR_UP = 13100
Public Const RPFZONMT_UP = 13200
Public Const REEXPORT_UP = 13300

Public Const RECURRING_AP_UP = 13350
Public Const RECURRING_GRP_UP = 13400

'new toolbar buttons id in FOFSITE
Public Const ROUTE_CODE_UP = 13450
Public Const FO_PRODUCT_UP = 13500
Public Const RELATED_SITE_UP = 13550
Public Const SITE_METER_UP = 13600
Public Const DEVICE_LOC_CODE_UP = 13650
Public Const DELVRY_FREQ_UP = 13700

'new toolbar buttons id in ARFMASTR
Public Const CUST_ACCESS_UP = 13750
Public Const SALESMAN_UP = 13800
Public Const DISPATCH_ZONE_UP = 13850

'note: these button does not launch EXE program
'require callback when add button
Public Const DRIVER_DIRECTION_UP = 13900
Public Const CUST_NOTE_UP = 13950
Public Const ORDER_STAT_UP = 14000
Public Const VIEW_DETAIL_UP = 14050
Public Const TOGGLE_TRUCK_UP = 14100
'''

Public Const TRUCK_UP = 14150
Public Const DRIVER_UP = 14200
Public Const DELV_REASON_UP = 14250
Public Const CONTRACT_TYPE_UP = 14300

'note: these button does not launch EXE program
'require callback when add button
Public Const RENEW_CONTRACT_UP = 14350
'''

'Robert Atwood 09-19-01 For TBKit Reportserver mod
Public Const RPTSRV_UP = 14400
Public Const RPTSRVSEC_UP = 14450
'Robert Atwood 10-03-01 for WOENTRY toolbar
Public Const ADDNOTES_UP = 14500
Public Const VIEWWOHIST_UP = 14550
Public Const VIEWCUSTINFO_UP = 14600
Public Const POERCVER_UP = 14650
Public Const WSERCVER_UP = 14700

Public Const RPT_SECURITY_UP = 14750
Public Const MASS_BESTBUY_UP = 14800
Public Const PRINT_SINGLE_ORDER_UP = 14850
Public Const VENDOR_CATEGORY_UP = 14900
Public Const CSTORE_TIER_MAINT_UP = 14950
Public Const NO_DELIVERY_REASON_UP = 15000
Public Const NEW_COMPANY_UP = 15050
Public Const QUALIFICATION_UP = 15100

'david 06/26/2002
Public Const CUSTOMER_CONTRACT_UP = 15150
Public Const CONTRACT_SELECTION_UP = 15200
Public Const FM_FO_INTERFACE_UP = 15250
'''''''''''''''''

'generic buttons for toolbar button that requires new bitmap
'note: these button does not launch EXE program
'require callback when add button
Public Const GENERIC1_UP = 32701
Public Const GENERIC2_UP = 32702
Public Const GENERIC3_UP = 32703
Public Const GENERIC4_UP = 32704
Public Const GENERIC5_UP = 32705
Public Const GENERIC6_UP = 32706
Public Const GENERIC7_UP = 32707
'''

Public Const TEXT_HEIGHT As Integer = 390
Public Const CURSOR_RESET As Integer = -1   'used to set cursor back to the default condition

Public Type CursorMode

    nFrameCount As Integer
    arryFrames() As Object
    narryFrameCursor() As Integer
    narryInitialFrameCursor() As Integer
    szarryFrameMessage() As String

End Type

Private Type tpLockHandles
    m_lHandle As Long
    m_sTable As String
End Type

Private arryLockHandles() As tpLockHandles
Private nHandleCount As Integer

'global variable for mouse pointer
Public objCurrTabControl As Object

Public tgcDropdown As Object
Public Const SYSTEM_AR_TRAN_CODES = " ('BB','BC','BD','BM','CC','CF','CO','DD','FC','FD','HC','OB','OC','PR','PY','RP','SA','XC','XF') " 'Hard coded sys ar tran codes WJ 4/14/99
Private Const Log10 = 2.30258509299405
Private SYS_PARM_14000 As String
Private SYS_PARM_6005 As String

Public Const CUST_ON_HOLD_STATS = " ('BH','OH') " 'WJ 04/18/2001

'david 04/01/2002
Public Const t_lBigFormWidth As Long = 11835
Public Const t_lBigFormHeight As Long = 8760
'

Public Function tfnIs_ON_HOLD(ByVal vStatus) As Boolean
    Dim sCustStatus As String * 2
    
    If IsNull(vStatus) Then
        vStatus = ""
    Else
        vStatus = Trim(vStatus)
    End If
    
    sCustStatus = vStatus
    
    tfnIs_ON_HOLD = (Right(sCustStatus, 1) = "H")
End Function

Public Function tfnIS_RM(Optional sRetSysParm14000 As String = "") As Boolean
    Dim strSQL As String
    Dim rsTemp As Recordset
    On Error GoTo ErrTrap
    If Not (SYS_PARM_14000 = "Y" Or SYS_PARM_14000 = "N") Then
        SYS_PARM_14000 = "N"
        strSQL = "SELECT parm_field FROM sys_parm WHERE parm_nbr = 14000"
        Set rsTemp = t_dbMainDatabase.OpenRecordset(strSQL, dbOpenSnapshot, SQL_PASSTHROUGH)
        If Not rsTemp.EOF Then
            If Not IsNull(rsTemp!parm_field) Then
                SYS_PARM_14000 = UCase(Trim$(rsTemp!parm_field))
            End If
        End If
        rsTemp.Close
    End If
    
    If SYS_PARM_14000 = "Y" Then
        tfnIS_RM = True
    Else
        tfnIS_RM = False
    End If
    
    sRetSysParm14000 = SYS_PARM_14000
    
    Exit Function
    
ErrTrap:
    tfnIS_RM = False
    ''tfnErrHandler "tfnIS_RM", strSQl

End Function

'======================
'Form Support Functions
'======================
Public Function tfnIsSysTranCode(ByVal sARTranCode As String) As Boolean 'WJ 4/14/99
    sARTranCode = "'" & UCase(sARTranCode) & "'"
    tfnIsSysTranCode = (InStr(SYSTEM_AR_TRAN_CODES, sARTranCode) > 0)
End Function
Public Sub tfnClearLog(szFilename)

    On Error Resume Next
    Kill szFilename
    
End Sub

Public Function tfnGetNamedString(sSource As String, sName As String) As String
    tfnGetNamedString = ""
    If sSource = "" Or sName = "" Then
        Exit Function
    End If
    
    Dim nPos1 As Integer
    Dim nPos2 As Integer
    Dim sUcaseSource As String
    Dim sUcaseName As String
    
    sUcaseSource = UCase(sSource)
    sUcaseName = UCase(sName)
    nPos1 = InStr(sUcaseSource, sUcaseName)

    If nPos1 > 0 Then
        nPos1 = InStr(nPos1, sUcaseSource, "=")
        If nPos1 > 0 Then
            nPos2 = InStr(nPos1, sUcaseSource, ";")
            If nPos2 = 0 Then
                nPos2 = Len(sUcaseSource) + 1
            End If
            nPos1 = nPos1 + 1
            If nPos2 > nPos1 Then
                tfnGetNamedString = Trim(Mid(sSource, nPos1, nPos2 - nPos1))
            End If
        End If
    End If
End Function

Public Function tfnGetUserName() As String
    'return the current username as was logged into factmenu
    
    #If DEVELOP Or (FACTOR_MENU >= 0) Then
        tfnGetUserName = "ssfactor"
        If t_dbMainDatabase Is Nothing Then Exit Function
            
        tfnGetUserName = tfnGetNamedString(t_dbMainDatabase.Connect, "UID")
    #Else
        If t_oleObject Is Nothing Then
            If Not t_dbMainDatabase Is Nothing Then
                tfnGetUserName = tfnGetNamedString(t_dbMainDatabase.Connect, "UID")
            End If
        Else
            tfnGetUserName = t_oleObject.UserName
        End If
    #End If
    
End Function

'Function : tfnGet_AR_Access_Flag
'Variables: Cust #, User(optional)
'Return   : (1)szEmpty --- no access at all
'           (2)E       --- Editable
'           (3)V       --- View Only

Public Function tfnGet_AR_Access_Flag(ByVal sCust As String, _
                                Optional vUser As Variant) As String
        
    Dim strSQL As String
    Dim rsTemp As Recordset
    Dim sAccess As String
    Dim sUser As String
    Dim sZone As String
    
    Static sSys_Parm_8 As String
    
    On Error GoTo ErrorTrap
    
    If sSys_Parm_8 = szEMPTY Then
        strSQL = "SELECT parm_field FROM sys_parm WHERE parm_nbr = 8 "
        
        Set rsTemp = t_dbMainDatabase.OpenRecordset(strSQL, dbOpenSnapshot, dbSQLPassThrough)
   
        If Not rsTemp Is Nothing Then
            If rsTemp.RecordCount > 0 Then
                If Not IsNull(rsTemp!parm_field) Then
                    sSys_Parm_8 = UCase(Trim(rsTemp!parm_field))
                End If
            End If
        End If
        If sSys_Parm_8 <> "Y" Then
            sSys_Parm_8 = "N"
        End If
    End If
    
    If sSys_Parm_8 = "Y" Then
        If IsMissing(vUser) Then
            sUser = tfnGetUserName
        Else
            sUser = vUser
        End If
               
        strSQL = "SELECT an_access_zone FROM ar_altname WHERE an_customer = " & Val(sCust)
        
        Set rsTemp = t_dbMainDatabase.OpenRecordset(strSQL, dbOpenSnapshot, dbSQLPassThrough)
   
        sZone = szEMPTY
        If Not rsTemp Is Nothing Then
            If rsTemp.RecordCount > 0 Then
                If Not IsNull(rsTemp!an_access_zone) Then
                    sZone = Trim(rsTemp!an_access_zone)
                End If
            End If
        End If
        
        If sZone = szEMPTY Then
            sAccess = "E" 'zone is not defined for the customer yet! do as usual
        Else
             strSQL = "SELECT ara_privilege FROM ar_access WHERE ara_access_zone = " & tfnSQLString(sZone)
             strSQL = strSQL & " AND ara_userid = " & tfnSQLString(sUser)
                
             sAccess = szEMPTY
             Set rsTemp = t_dbMainDatabase.OpenRecordset(strSQL, dbOpenSnapshot, dbSQLPassThrough)
            
             If Not rsTemp Is Nothing Then
                 If rsTemp.RecordCount > 0 Then
                     If Not IsNull(rsTemp!ara_privilege) Then
                         sAccess = UCase(Trim(rsTemp!ara_privilege))
                     End If
                 End If
             End If
        End If
    Else ' if this future is off, user has full access
        sAccess = "E"
    End If
    
    tfnGet_AR_Access_Flag = sAccess
    Exit Function

ErrorTrap:
    MsgBox "There is an error in checking customer access privilege." & vbCrLf & "Error Code: " & Err.Number & vbCrLf & " Error Desc: " & Err.Description, vbCrLf
    Err.Clear
    tfnGet_AR_Access_Flag = szEMPTY
    
    On Error GoTo 0
End Function

'
'Function : tfnCheckMySize - maintains a readable minimum size for all resolutions
'Variables: pointer to the form that received resize event
'Return   : none
'

Public Sub tfnCheckMySize(frmForm As Form)
    
    #If DEVELOP Then
        MsgBox "Remove your FORM_RESIZE event!"
    #End If

End Sub
'
'Function : tfnLockElasticControls - turn on elastic controls turned off at design time
'Variables: pointer to the form with eleastic controls
'Return   : none
'
Public Sub tfnLockElasticControls(frmForm As Form)
    
    #If DEVELOP Then
        MsgBox "Remove all calls to tfnLockElasticControls!"
    #End If

End Sub

Public Function tfnLockRow(sProgramID As String, _
                           sTable As String, _
                           sSql As String, _
                           Optional vShowMsg As Variant, _
                           Optional sLockedUser As String = "") As Boolean

    Const SUB_NAME = "tfnLockRow"
    Const sErrID = "Lock Row"

    Dim nPos1 As Integer
    Dim nPos2 As Integer
    Dim strSQL As String
    Dim rsTemp As Recordset
    Dim sCriteria As String
    Dim sUserID As String
    Dim sTemp As String
    Dim t_lLockHandle As Long     'Handle for row lock routine
    Dim I As Integer

    #If FACTOR_MENU = 1 Then
        tfnLockRow = True
        Exit Function
    #End If
    
    tfnLockRow = False
    
    #If DEVELOP Then
        If Trim(sTable) = "" Then
            MsgBox "You have to provide the table name in which you want to lock a row", , sErrID
        End If
        If Trim(sProgramID) = "" Then
            MsgBox "You have to provide the program ID to lock a row", , sErrID
        End If
        If Trim(sSql) = "" Then
            MsgBox "You have to provide the criteria or the SQL to lock a row", , sErrID
        End If
        On Error GoTo errTableName
        strSQL = "SELECT * FROM " & sTable & " WHERE ROWID = 1"
        Set rsTemp = t_dbMainDatabase.OpenRecordset(strSQL, dbOpenSnapshot, dbSQLPassThrough)
        rsTemp.Close
        sUserID = tfnGetNamedString(t_dbMainDatabase.Connect, "UID")
    #Else
        If t_oleObject Is Nothing Then
            sUserID = tfnGetNamedString(t_dbMainDatabase.Connect, "UID")
        Else
            sUserID = t_oleObject.UserName
        End If
    #End If

    'Get the where clause
    strSQL = UCase(sSql)
    nPos2 = InStr(strSQL, " WHERE ")
    If nPos2 > 0 Then
        nPos1 = InStr(strSQL, " ORDER ")
        If nPos1 = 0 Then
            nPos1 = Len(sSql) + 1
        End If
        sCriteria = Mid(sSql, nPos2 + 7, nPos1 - nPos2 - 7)
    Else
        sCriteria = sSql
    End If
    
    #If DEVELOP Then
        If Len(sCriteria) > 80 Then
            MsgBox "The criteria is too long." & vbKeyReturn & "Probably, you need to remove the field names", vbOKOnly
            Exit Function
        End If
    #End If
    
    sTemp = LCase(Trim(sTable))
    
    For I = 0 To nHandleCount - 1
        If sTemp = arryLockHandles(I).m_sTable Then
            tfnLockRow = True
            Exit Function
        End If
    Next I

    On Error GoTo errOpenRecord
    strSQL = "EXECUTE PROCEDURE lock_row(" & tfnSQLString(sTemp) & ", " & tfnSQLString(sProgramID) & ", " & tfnSQLString(sUserID) & ", " & tfnSQLString(sCriteria) & ")"
    Set rsTemp = t_dbMainDatabase.OpenRecordset(strSQL, dbOpenSnapshot, dbSQLPassThrough)
    
    If rsTemp.RecordCount > 0 Then
        t_lLockHandle = rsTemp.Fields(0)
        
        If t_lLockHandle = 0 Then
            If Trim(rsTemp.Fields(1)) = "" Then
                #If DEVELOP Then
                    MsgBox "Make sure you logged on a database with locking procedures setup", vbOKOnly
                #End If
            Else
                Dim bShowMsg As Boolean
                
                If IsMissing(vShowMsg) Then
                    bShowMsg = True
                Else
                    bShowMsg = vShowMsg
                End If
                
                'david 01/12/2001
                'return the user id that locks the record(s)
                sLockedUser = Trim(rsTemp.Fields(1))
                
                If bShowMsg Then
                    MsgBox "The record you have selected is locked by " & sLockedUser & "." & vbCrLf & "Select another record for edit or try again later.", vbOKOnly
                End If
            End If
        End If
    End If
    
    rsTemp.Close
    Set rsTemp = Nothing
    
    If t_lLockHandle > 0 Then
        If I >= nHandleCount Then
            If nHandleCount = 0 Then
                nHandleCount = 1
                ReDim arryLockHandles(nHandleCount - 1)
            Else
                nHandleCount = nHandleCount + 1
                ReDim Preserve arryLockHandles(nHandleCount - 1)
            End If
        End If
        
        tfnLockRow = True
        arryLockHandles(I).m_sTable = sTemp
        arryLockHandles(I).m_lHandle = t_lLockHandle
    End If
    Exit Function
 
errOpenRecord:
    #If NO_ERROR_HANDLER Then
        MsgBox "Cannot lock table"
    #Else
        If Not objErrHandler Is Nothing Then
            tfnErrHandler SUB_NAME, strSQL
        End If
    #End If
    Err.Clear
    Exit Function

errTableName:
    #If DEVELOP Then
        MsgBox "Please make sure the table name for locking is correct", vbOKOnly, App.Title
    #End If
    Err.Clear
End Function

'
'Function : tfnResizeFonts - resize fonts
'Variables: pointer to the form, pointer to FontNames and FontSizes arrays
'Return   : none
'
Public Sub tfnResizeFonts(frmMyForm As Form, nFontSizes() As Integer)

    #If DEVELOP Then
        MsgBox "Remove all calls to tfnResizeFonts!"
    #End If

End Sub
'
'Function : tfnStoreFontInfo - saves design time font information
'Variables: pointer to the form, pointer to FontNames and FontSizes arrays
'Return   : none
'
Public Sub tfnStoreFontInfo(frmForm As Form, arrayFontSizes() As Integer)

    #If DEVELOP Then
        MsgBox "Remove all calls to tfnStoreFontInfo!"
    #End If

End Sub

Public Function tfnUnlockRow(Optional vTable As Variant) As Boolean
    Const SUB_NAME = "tfnUnlockRow"
    
    #If FACTOR_MENU = 1 Then
        tfnUnlockRow = True
        Exit Function
    #End If
    
    If nHandleCount <= 0 Then
        Exit Function
    End If

    Dim strSQL As String
    Dim rsTemp As Recordset
    
    tfnUnlockRow = False
    On Error GoTo errUnlock
    
    If IsMissing(vTable) Then
        While nHandleCount > 0
            strSQL = "EXECUTE PROCEDURE unlock_row(" & CStr(arryLockHandles(nHandleCount - 1).m_lHandle) & ")"
            Set rsTemp = t_dbMainDatabase.OpenRecordset(strSQL, dbOpenSnapshot, dbSQLPassThrough)
            If rsTemp.RecordCount > 0 Then
                If rsTemp.Fields(0) > 0 Then
                    arryLockHandles(nHandleCount - 1).m_sTable = ""
                    arryLockHandles(nHandleCount - 1).m_lHandle = -1
                    nHandleCount = nHandleCount - 1
                Else
                    rsTemp.Close
                    Exit Function
                End If
            Else
                rsTemp.Close
                Exit Function
            End If
        Wend
        
        ReDim arryLockHandles(0)
        rsTemp.Close
    Else
        Dim sTable As String
        Dim I As Long
        Dim j As Long
        
        sTable = LCase(Trim(vTable))
        
        For I = 0 To nHandleCount - 1
            If sTable = arryLockHandles(I).m_sTable Then
                strSQL = "EXECUTE PROCEDURE unlock_row(" & CStr(arryLockHandles(I).m_lHandle) & ")"
                Set rsTemp = t_dbMainDatabase.OpenRecordset(strSQL, dbOpenSnapshot, dbSQLPassThrough)
                If rsTemp.RecordCount > 0 Then
                    If rsTemp.Fields(0) > 0 Then
                        arryLockHandles(I).m_sTable = ""
                        arryLockHandles(I).m_lHandle = -1
                        nHandleCount = nHandleCount - 1
                    Else
                        rsTemp.Close
                        Exit Function
                    End If
                Else
                    rsTemp.Close
                    Exit Function
                End If
                
                Exit For
            End If
        Next I
        
        If I < UBound(arryLockHandles) Then
            For j = I + 1 To UBound(arryLockHandles)
                arryLockHandles(j - 1).m_sTable = arryLockHandles(j).m_sTable
                arryLockHandles(j - 1).m_lHandle = arryLockHandles(j).m_lHandle
            Next j
        End If
        
        If nHandleCount > 0 Then
            ReDim Preserve arryLockHandles(nHandleCount - 1)
        Else
            ReDim arryLockHandles(0)
        End If
    End If
    
    Set rsTemp = Nothing
    tfnUnlockRow = True
    Exit Function

errUnlock:
    #If NO_ERROR_HANDLER Then
        MsgBox "Cannot unlock table"
    #Else
        If Not objErrHandler Is Nothing Then
            tfnErrHandler SUB_NAME, strSQL
        End If
    #End If
End Function

'update program version
Public Sub tfnUpdateVersion()
#If FACTOR_MENU < 0 Then
    Dim sProgramName As String, sMajorVersion As String, sSql As String, rsTemp As Recordset
    Dim sMinorVersion As String, sRevision As String, sUserName As String
    #If DEVELOP Then
        Dim nSpot As Integer
    #End If
    On Error GoTo ErrorInMakeExeOption
    #If DEVELOP Then
        nSpot = 1
    #End If
    
    sProgramName = UCase(Trim(App.FileDescription))
    sMajorVersion = Trim(CStr(App.Major))
    sMinorVersion = Trim(CStr(App.Minor))
    sRevision = Trim(CStr(App.Revision))
    
    On Error GoTo 0
    
    If sProgramName = "" Then
        #If DEVELOP Then
            nSpot = 2
        #End If
        GoTo ErrorInMakeExeOption
    Else
        If InStr(sProgramName, ".") > 1 Then
            sProgramName = Mid(sProgramName, 1, InStr(sProgramName, ".") - 1)
        End If
    End If
    
    If sMajorVersion = "" Or sMinorVersion = "" Then
        #If DEVELOP Then
            nSpot = 3
        #End If
        GoTo ErrorInMakeExeOption
    End If
    
    If sProgramName <> UCase(Trim(App.EXEName)) Then
        #If FACTOR_MENU >= 0 Then
            MsgBox "Error in tfnUpdateVersion(): EXE file name and File Description in VB Make Exe Option are NOT match.", vbCritical
        #Else
            #If DEVELOP Then
                nSpot = 4
            #End If
            GoTo ErrorExecuteSQL
        #End If
        Exit Sub
    End If
    
    sUserName = tfnGetUserName()
    
    sSql = "EXECUTE PROCEDURE pro_versions ( '" & UCase(sProgramName) & "', "
    sSql = sSql & "'" & sMajorVersion & "." & sMinorVersion & "." & sRevision
    sSql = sSql & "', 'P', '" & sUserName & "')"
    
    On Error GoTo ErrorExecuteSQL
    #If DEVELOP Then
        nSpot = 5
    #End If
    Set rsTemp = t_dbMainDatabase.OpenRecordset(sSql, dbOpenSnapshot, dbSQLPassThrough)
    On Error GoTo 0
    
    If rsTemp.RecordCount = 0 Then
        #If DEVELOP Then
            nSpot = 6
        #End If
        GoTo ErrorExecuteSQL
    Else
        If rsTemp.Fields(0) = 0 Then
            #If Not NO_ERROR_HANDLER Then
                tfnErrHandler "tfnUpdateVersion", sSql, rsTemp.Fields(1)
            #End If
            MsgBox "Version Update Failed! " & vbCrLf & vbCrLf & "pro_version Error.", vbExclamation
        End If
    End If

    Exit Sub
    
ErrorInMakeExeOption:
#If FACTOR_MENU >= 0 Then
    MsgBox "Error in tfnUpdateVersion(): Make Exe Option parameter(s) has not been setup.", vbCritical
    Exit Sub
#End If
ErrorExecuteSQL:
    #If DEVELOP Then
        MsgBox "Version Update Failed! " & vbCrLf & vbCrLf & "pro_version Error. at spot: " & CStr(nSpot), vbExclamation
        If nSpot = 4 Then
            MsgBox "The program Name is : " & sProgramName & vbCrLf & vbCrLf & "but the application name is : " & UCase(Trim(App.EXEName))
        End If
    #Else
        #If Not NO_ERROR_HANDLER Then
            tfnErrHandler "tfnUpdateVersion", sSql
        #End If
    #End If
#End If
End Sub

'
'Function : tfnUpdateStatusBar - updates the status bar CAPS, NUM, and SCRL panes
'Variables: pointer to the form the status bar is on
'Return   : none
'
Public Sub tfnUpdateStatusBar(frmForm As Form, Optional bRefreshStatus As Boolean = True)
    Dim sOn As String
    
    On Error Resume Next
    
    sOn = " on " & Trim(tfnGetDataSourceName)
    
    If Trim(sOn) <> "on" Then
        If InStr(frmForm.Caption, sOn) <= 0 Then
            frmForm.Caption = frmForm.Caption & sOn
        End If
    End If
    
    If Not bRefreshStatus Then
        Exit Sub
    End If
    
    Dim intKeyStatus As Integer
    
    intKeyStatus = GetKeyState(VK_CAPITAL)

    If intKeyStatus = 1 Then
        frmForm.ffraStatusbar.PanelCaption(2) = "CAPS"
    Else
        frmForm.ffraStatusbar.PanelCaption(2) = szEMPTY
    End If

    DoEvents

    intKeyStatus = GetKeyState(VK_SCROLL)

    If intKeyStatus = 1 Then
        frmForm.ffraStatusbar.PanelCaption(0) = "SCRL"
    Else
        frmForm.ffraStatusbar.PanelCaption(0) = szEMPTY
    End If

    DoEvents

    intKeyStatus = GetKeyState(VK_NUMLOCK)

    If intKeyStatus = 1 Then
        frmForm.ffraStatusbar.PanelCaption(1) = "NUM"
    Else
        frmForm.ffraStatusbar.PanelCaption(1) = szEMPTY
    End If
End Sub

'===========================
'oleObject Support Functions
'===========================

'
'Function : tfnExecuteProgram - wrapper around oleObject.RunExe function - displays consistant error message
'Variables: pointer to the oleObject, program name to run from application context menu/toolbar
'Return   : true if application launched, false if not
'
Public Function tfnExecuteProgram(oleObject As Object, szProgram As String) As Boolean

    If oleObject.IsExeValid(szProgram) = True Then 'if user has the privledges to run application
        
        If oleObject.RunExe(szProgram) = False Then 'attempt to run the application failed
            MsgBox szRUNEXE_ERROR & szProgram, vbCritical, szRUNEXE_TITLE 'display message to the user
            tfnExecuteProgram = False 'return error code
        Else
            tfnExecuteProgram = True 'else return application launched ok
        End If
    
    Else
        MsgBox oleObject.ErrorMessage, vbCritical, szRUNEXE_TITLE 'display oleObject error message
        tfnExecuteProgram = False 'return error code
    End If

End Function
'
'Function : tfnOpenDatabase() - opens the database
'Variables: none
'Return   : database handle
'
'david 02/09/00
'changed function to handle background processing
'if no parameter is supplied, this function will show error message box
'other for backgroud process, these two parameters are actually REQUIRED.
'Pass a False to bShowMsgBox to suppress the error message box, in turn,
'return the error message to the calling function.
Public Function tfnOpenDatabase(Optional bShowMsgBox As Boolean = True, _
                                 Optional sErrMsg As String = "") As Boolean
    Dim I As Integer
    
    #If FACTOR_MENU = 1 Then
        tfnOpenDatabase = True
        Exit Function
    #ElseIf FACTOR_MENU = 0 Then                              'for developement allow the standard ODBC Connect Dialog Box
        If Trim(t_szConnect) = "" Then
            t_szConnect = szODBC                              'ODBC string activate ODBC Connect Dialog Box
        End If
    #ElseIf FACTOR_MENU < 0 Then                              'use the FACTOR Menu oleObject for Database connection string
        If Trim(t_szConnect) = "" Then
            On Error Resume Next
            Set t_oleObject = CreateObject(t_szOLEObjectName) 'get the handle to the oleObject internal to the FACTOR Main Menu
            t_oleObject.EXEName = App.EXEName
            t_szConnect = t_oleObject.MainConnectString       'get the FACTOR Main Menu connect string
        End If
    #End If

    On Error GoTo ERROR_CONNECTING 'set the runtime error handler for database connection

    Set t_engFactor = New DBEngine 'create a new dDBEngine
    
    
    Set t_wsWorkSpace = t_engFactor.Workspaces(0) 'set the default workspace handle
    t_engFactor.IniPath = tfnGetSystemDir 'put the path in engine ini variable
    
    Set t_dbMainDatabase = t_wsWorkSpace.OpenDatabase("", False, False, t_szConnect)
    
    
    tfnOpenDatabase = True
    Exit Function

ERROR_CONNECTING:
    If t_oleObject Is Nothing Then
        If bShowMsgBox Then
            MsgBox fnShowODBCError(), vbCritical
        Else
            sErrMsg = fnShowODBCError()
        End If
    Else
        If bShowMsgBox Then
            MsgBox Err.Description, vbOKOnly + vbCritical, szCONNECTION_ERROR
        Else
            sErrMsg = Err.Description
        End If
    End If

    tfnOpenDatabase = False

End Function

Private Function fnShowODBCError() As String
    Dim I As Integer
    Dim sMsgs As String
    Dim sNumbers As String
    Dim sODBCErrors As String
    
    If Err.Number = 3146 Then
        With t_engFactor.Errors
            If .Count > 0 Then
                For I = 0 To .Count - 2
                    sMsgs = sMsgs & "Number: " & .Item(I).Number & Space(5) & .Item(I).Description & vbCrLf
                Next
            End If
            If .Count <= 2 Then
                sNumbers = ""
            Else
                sNumbers = "s"
            End If
        End With
        sODBCErrors = "The following error" & sNumbers & " occurred while doing an ODBC query:" & vbCrLf & vbCrLf _
                       & vbCrLf & sMsgs
    Else
        sODBCErrors = Err.Description
    End If

    fnShowODBCError = sODBCErrors
    Err.Clear

End Function

Public Function tfnRound(vTemp As Variant, _
                         Optional vPrec As Variant) As Variant

    Dim fTempD As Double
    Dim sFmt As String
    Dim nPrec As Integer
    Dim fOffset As Double
    Dim sTemp As String
    
    If IsMissing(vPrec) Then
        nPrec = 0
    Else
        nPrec = vPrec
    End If
    If IsNull(vTemp) Then
        tfnRound = 0
    Else
        If Trim(vTemp) = "" Then
            tfnRound = 0
        Else
            If IsNumeric(vTemp) Then
                If nPrec >= 0 Then
                    sFmt = "#0." & String(nPrec, "#")
                    If VarType(vTemp) = vbDouble And vTemp <> 0 And Abs(vTemp) < 100000 And nPrec = 2 Then
                        'If format with 2 decimal point places, we suppose that it is dealing with money
                        fTempD = CDbl(vTemp)
                        fOffset = Sgn(vTemp) * 10 ^ (Log(Abs(vTemp)) / Log10 - 7.375)
                        tfnRound = Val(Format(vTemp + fOffset, sFmt))
                    Else
                        sTemp = CStr(vTemp)
                        tfnRound = Val(Format(sTemp, sFmt))
                    End If
                Else
                    sTemp = CStr(vTemp)
                    tfnRound = Val(Format(sTemp, "#"))
                End If
            Else
                tfnRound = 0
            End If
        End If
    End If
End Function

Public Function tfnOpenLocalDatabase(Optional bShowMsgBox As Boolean = True, _
                                 Optional sErrMsg As String = "") As DataBase

'#####################################################################
'# Modified 10-30-01 Robert Atwood to implement Multi-Company factmenu
'# (Must read factor.mdb from c:\factor\<datasourcename>\factor.mdb
'#####################################################################
    Dim sWinSysDir As String

    #If DEVELOP Then
        sWinSysDir = LOCAL_FACTOR_PATH
    #Else
        sWinSysDir = LOCAL_FACTOR_PATH & UCase(Trim(tfnGetDataSourceName)) + "\"
    #End If
    
    #If FACTOR_MENU <> 1 Then
        On Error GoTo ERROR_CONNECTING 'set the runtime error handler for database connection
    
        'david 11/15/2001
        If t_engFactor Is Nothing Then
            Set t_engFactor = New DBEngine 'create a new dDBEngine
            t_engFactor.IniPath = sWinSysDir 'put the path in engine ini variable
        End If
        
        If t_wsWorkSpace Is Nothing Then
            Set t_wsWorkSpace = t_engFactor.Workspaces(0) 'set the default workspace handle
        End If
        
        If Not fnCopyFactorMDB() Then
            sErrMsg = "Could not create new local database"
            
            If bShowMsgBox Then
                MsgBox sErrMsg + ".", vbExclamation
            End If
            
            Exit Function
        End If
  
        Set tfnOpenLocalDatabase = t_wsWorkSpace.OpenDatabase(sWinSysDir & "factor.mdb")
        On Error GoTo 0
        Exit Function
    
ERROR_CONNECTING:
        
        If bShowMsgBox Then
            MsgBox Err.Description, vbOKOnly + vbCritical, "Local Access " & szCONNECTION_ERROR
        Else
            sErrMsg = Err.Description
        End If
        
        Set tfnOpenLocalDatabase = Nothing
    
        On Error GoTo 0
    #End If
End Function

'
'Function : tfnAuthorizeExecute() - prevents released application from executing outside FACTOR Main Menu
'Variables: command line string - passing by FACTOR Main Menu in Shell function
'Return   : true if handshake ok, false if not
'
Public Function tfnAuthorizeExecute(szHandShake As String) As Boolean
 
#If FACTOR_MENU >= 0 Then             'during development bypass handshake requirement
        tfnAuthorizeExecute = True   'return ok to run application
#ElseIf FACTOR_MENU < 0 Then 'released application can only be run from FACTOR Main Menu
    If szHandShake = t_szHandShake Then 'and only if you know the secret hand shake
        tfnAuthorizeExecute = True      'handshake ok, return ok to run application to caller
    Else  'you don't know squat!
        If Trim(t_szConnect) = "" Then
            MsgBox szRUN_ERROR, vbOKOnly + vbCritical, App.Title 'display error message to the user
            tfnAuthorizeExecute = False 'return error flag
        Else
            tfnAuthorizeExecute = True
        End If
    End If
#End If

End Function

'========================================
'Global General Purpose Support Functions
'========================================

'
'Function : tfnCenterForm - centers a form in the screen
'Variables: pointer to the form, optional pointer to parent form
'Return   : none
'
Sub tfnCenterForm(frmCurrent As Form, Optional vParentForm As Variant)
  
    If IsMissing(vParentForm) Then
        frmCurrent.Left = (Screen.Width - frmCurrent.Width) \ 2
        frmCurrent.Top = (Screen.Height - frmCurrent.Height) \ 2
    Else
        
        If vParentForm.Width > frmCurrent.Width Then
            frmCurrent.Left = vParentForm.Left + (vParentForm.Width - frmCurrent.Width) \ 2
        Else
            frmCurrent.Left = (Screen.Width - frmCurrent.Width) \ 2
        End If

        If vParentForm.Height > frmCurrent.Height Then
            frmCurrent.Top = vParentForm.Top + (vParentForm.Height - frmCurrent.Height) \ 2
        Else
            frmCurrent.Top = (Screen.Height - frmCurrent.Height) \ 2
        End If
    End If
    
End Sub
'
'Function        : tfnConfirm - msgbox wrapper
'Passed Variables: message to display
'Returns         : true for yes, false for no
'
Public Function tfnConfirm(szMessage As String, Optional vDefaultButton As Variant) As Boolean
  
  Dim nStyle As Long
  'vDefaultButton -- added by WJ
  If IsMissing(vDefaultButton) Then
    nStyle = vbYesNo + vbQuestion ' put focus on Yes
  Else
    nStyle = vbYesNo + vbQuestion + Val(vDefaultButton) 'Put Focus to Yes or No
  End If
  If MsgBox(szMessage, nStyle, App.Title) = vbYes Then
    tfnConfirm = True
  Else
    tfnConfirm = False
  End If
  
End Function

'added by xijian on 1/13/00
Public Function tfnBuildMultiLines(sParam() As String, _
                           sSrc As String, _
                           sDelim As String, _
                           Optional vStart As Variant, _
                           Optional vEnd As Variant)
                          
    If Trim(sSrc) = "" Then
        Exit Function
    End If

    Const nArrayInc As Integer = 5
    Dim i1 As Integer
    Dim i2 As Integer
    Dim k As Integer
    Dim nEnd As Integer
    Dim sTemp As String
    
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
        i2 = InStr(i1, sSrc, sDelim)
        If i2 >= i1 And i2 <= nEnd Then
            If k > UBound(sParam) Then
                ReDim Preserve sParam(k + nArrayInc)
            End If
            sTemp = Mid$(sSrc, i1, i2 - i1)
            If sTemp <> "" Or sDelim <> " " Then
                sParam(k) = sTemp
                k = k + 1
            End If
            i1 = i2 + Len(sDelim)
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
    tfnBuildMultiLines = UBound(sParam) + 1
End Function

Public Function tfnGetMultiLines(rsTemp As Recordset, Optional fieldNum As Variant) As String
    Dim sTemp As String
    
    If rsTemp.RecordCount > 0 Then
        If IsMissing(fieldNum) Then
            fieldNum = 0
        End If
        'first line
        If Not IsNull(rsTemp.Fields(fieldNum)) Then
            'david 04/30/2002
            sTemp = RTrim$(fnRemoveChr0(rsTemp.Fields(fieldNum)))
            '''''''''''''''''
        Else
            sTemp = ""
        End If
        
        rsTemp.MoveNext
        
        'the rest
        While Not rsTemp.EOF
            If Not IsNull(rsTemp.Fields(fieldNum)) Then
                'david 04/30/2002
                sTemp = sTemp + vbCrLf + RTrim$(fnRemoveChr0(rsTemp.Fields(fieldNum)))
                '''''''''''''''''
            Else
                sTemp = sTemp + vbCrLf + ""
            End If
            rsTemp.MoveNext
        Wend
        
    End If
    
    tfnGetMultiLines = sTemp
End Function

'''end add


'
'Function        : tfnCancelExit - msgbox wrapper
'Passed Variables: Exit./Cancel message to display
'Returns         : true for yes, false for no
'
Public Function tfnCancelExit(szMessage As String) As Boolean
  
  If MsgBox(szMessage, vbYesNo + vbQuestion + vbDefaultButton2 + vbApplicationModal, App.Title) = vbYes Then
    tfnCancelExit = True
  Else
    tfnCancelExit = False
  End If
  
End Function
'
'Function        : tLockWin - traps extra mouse clicks - prevents stack overflows during long wait periods
'Passed Variables: optional pointer to current form, no paramter will unlock previous locked window
'Returns         : none
'
Public Sub tfnLockWin(Optional frmCurrent As Variant)

    Static frmSaved As Form 'pointer to any active form

    On Error Resume Next 'turn off the default runtime error handler

    If Not frmSaved Is Nothing Then          'if a previous form locked
        EnableWindow frmSaved.hwnd, -1       'disable the lock on window/form
        Set frmSaved = Nothing               'clear the pointer to the static form
        Screen.MousePointer = DEFAULT_CURSOR 'set the cursor back to the
    End If

    If Not IsMissing(frmCurrent) Then          'if a pointer to a form is valid
        Set frmSaved = frmCurrent              'save the pointer in the local static variable
        Screen.MousePointer = HOURGLASS_CURSOR 'set the mouse to the hourglass
        EnableWindow frmCurrent.hwnd, 0        'lock the window
    End If

End Sub
'
'Function        : tfnWaitSeconds
'Passed Variables: Number of seconds to wait
'Returns         : none
'
Public Sub tfnWaitSeconds(nSecondsToWait As Integer)
    
    Dim lStartTime As Long
    
    lStartTime = Timer
    
    Do While Timer < lStartTime + nSecondsToWait + 1
        
        DoEvents
    Loop

End Sub
'
'Function        : tfnLog - file log function
'Passed Variables: string to save in file, optional name of file to save data
'Returns         : true for yes, false for no
'
Public Sub tfnLog(szLogEntry As String, Optional szFilename As Variant)

    Dim nFileNumber As Integer
    
    On Error Resume Next
    
    nFileNumber = FreeFile
    
    If IsMissing(szFilename) Then
        Open szLOG_FILE_NAME For Append As #nFileNumber
    Else
        Open szFilename For Append As #nFileNumber
    End If
    
    Print #nFileNumber, szLogEntry

    Close nFileNumber
    
End Sub

'
'Function        : tfnIsFile - tests if file exists
'Passed Variables: filename
'Returns         : true if exists, false if not
'
Public Function tfnIsFile(ByVal szFilename As String) As Boolean
    
    Dim nLength As Integer
    
    On Error Resume Next

    nLength = Len(Dir(szFilename))
    
    If Err Or nLength = 0 Then
        tfnIsFile = False
    Else
        tfnIsFile = True
    End If

End Function
'
'Function : tfnStripNULL - strips off the NULL terminator on C strings, returned from windows api calls
'Variables: NULL terminated string
'Return   : original string with the null removed
'
Public Function tfnStripNULL(ByRef szString As String) As String
  
    Dim nPos As Integer 'position of the NULL terminator
   
    If Len(szString) = Null Or Len(szString) = 0 Then 'make sure string is valid
        szString = szEMPTY 'if not set the string to an empty string
    Else
    
        nPos = InStr(szString, Chr(0)) 'get the position of the NULL terminator
        
        If nPos > 0 Then 'if nPos is greater than 0 then a NULL was found
           szString = Left(szString, nPos - 1) 'strip off the NULL terminator
        End If 'if string did not have a NULL do not change it
    
    End If
    
    tfnStripNULL = szString 'return the string

End Function
'
'Function : tfnParseString - parses a delimited string, default is a space
'Variables: string to parse, optional delimiter, NOTE: original string is destroyed in the process, converion constant
'Return   : first deliminted substring in main string
'
Public Function tfnParseString(ByRef szMainString As String, Optional vDelimiter As Variant, Optional vConvertTo As Variant) As String

    Dim nPos As Integer        'position of the delimiter
    Dim szDelimiter As String  'delimiter to search for in the main string
    Dim szBuffer As String     'string buffer
    
    If IsMissing(vDelimiter) Then 'set the delimiter to as space if none was passed
        szDelimiter = szSPACE
    Else
        szDelimiter = vDelimiter
    End If
    
    szBuffer = Left(szMainString, 1)
    
    Do While szBuffer = szDelimiter
        szMainString = Mid(szMainString, 1)
        szBuffer = Left(szMainString, 1)
    Loop
    
    nPos = InStr(szMainString, szDelimiter) 'search for a delimiter

    If nPos > 0 Then 'if delimiter found
        szBuffer = Left(szMainString, nPos - 1) 'parse of the substring
        szMainString = Mid(szMainString, nPos + 1)   'remove the substring and delimiter from the main string
    Else
        szBuffer = szMainString 'return the last substring
        szMainString = szEMPTY 'empty the string
    End If
    
    If Not IsMissing(vConvertTo) Then 'if conversion constants passed
        If vConvertTo = vbUpperCase Or vConvertTo = vbLowerCase Or vConvertTo = vbProperCase Then
            szBuffer = StrConv(szBuffer, vConvertTo) 'convert to the case constant sent, if its valid
        End If
    End If
    
    tfnParseString = szBuffer

End Function
'
'Function : tfnGetSystemDir - gets the windows system directory
'Variables: optional variable to add a slash to the end of the path
'Return   : the path to the windows system directory
'
Public Function tfnGetSystemDir(Optional vAddSlash As Variant) As String
'# Modified 10-30-01 Robert Atwood to allow INI files to be stored in c:\factor instead of
'# system directory
    Dim nLength As Integer     'length returned from API call
    Dim szSystemDir As String  'temp string to hold system directory
    
    szSystemDir = Space(MAX_STRING_LENGTH) 'set the string to a fixed length for API call, pad with spaces
    
    'nLength = GetSystemDirectory(szSystemDir, MAX_STRING_LENGTH) 'call the API function to get the system directory
  
    'szSystemDir = Left(szSystemDir, nLength) 'trim off the excess spaces
    szSystemDir = LOCAL_FACTOR_PATH & UCase(Trim(tfnGetDataSourceName))
    If Not IsMissing(vAddSlash) Then
        If Right(szSystemDir, 1) <> szSLASH And vAddSlash = True Then 'add a slash if it needs one
            szSystemDir = szSystemDir + szSLASH
        End If
    End If
    
    tfnGetSystemDir = szSystemDir 'return system directory back to the caller

End Function
'
'Function : tfnGetWindowsDir - gets the windows directory
'Variables: optional variable to add a slash to the end of the path
'Return   : the path to the windows directory
'
Public Function tfnGetWindowsDir(Optional vAddSlash As Variant) As String
    '#Modified to return the factor directory always now
'    Dim nLength As Integer      'length returned from API call
'    Dim szWindowsDir As String  'temp string to hold windows directory
'
'    szWindowsDir = Space(MAX_STRING_LENGTH) 'set the string to a fixed length for API call, pad with spaces
'
'    nLength = GetWindowsDirectory(szWindowsDir, MAX_STRING_LENGTH) 'get the current windows directory
'
'    szWindowsDir = Left(szWindowsDir, nLength) 'trim off the excess spaces
'
'    If Not IsMissing(vAddSlash) Then
'        If Right(szWindowsDir, 1) <> szSLASH And vAddSlash = True Then 'add a slash if it needs one
'            szWindowsDir = szWindowsDir + szSLASH
'        End If
'    End If
    '######################################
    'Modified 10-31-01 Robert Atwood
    'for Multi-Company Factmentu
    '######################################
    tfnGetWindowsDir = LOCAL_FACTOR_PATH
  '  tfnGetWindowsDir = szWindowsDir 'return windows directory back to the caller

End Function

'Function : tfnGetAppDir - returns the application directory path
'Variables: optional variable to add a slash to the end of the path
'Return   : directory path
'
Public Function tfnGetAppDir(Optional vAddSlash As Variant) As String
    
    Dim szTemp As String 'temp to hold the path
        
    szTemp = App.Path 'use the App object to retrieve the path
        
    If Not IsMissing(vAddSlash) Then
        If Right(szTemp, 1) <> szSLASH And vAddSlash = True Then 'add a slash if it needs one
            szTemp = szTemp + szSLASH
        End If
    End If
    
    tfnGetAppDir = szTemp 'return the path

End Function
'
'Function : tfnReadINI - reads a value from a windows INI file
'Variables: [section], [key], and ini file name
'Return   : the [value] for the [section] and [key] sent
'
Public Function tfnReadINI(szSection As String, szKey As String, szINIFile As String) As String

    #If Win32 Then
        Dim nLength As Long 'length of the value returned for api call
    #Else
        Dim nLength As Integer 'length of the value returned for api call
    #End If
    
    Dim szINI As String    'string to hold the value retrieved

    szINI = Space(MAX_STRING_LENGTH) 'clear and make the string fixed length
    
    'get the [value] for the [section], [key], and ini file sent
    nLength = GetPrivateProfileString(szSection, szKey, szEMPTY, szINI, MAX_STRING_LENGTH, szINIFile)
    
    If nLength <> 0 Then 'if length positive [value] has been found
        szINI = Left(szINI, nLength) 'make it a basic string
    Else
        szINI = ""
    End If
    
    tfnReadINI = szINI 'return the value

End Function
'
'Function : tfnWriteINI - writes a value to a windows INI file
'Variables: [section], [key], [value], and ini file name
'Return   : status of api call
'
Public Function tfnWriteINI(szSection As String, szKey As String, szValue As String, szINIFile As String) As Boolean

    Dim bStatus As Boolean 'status returned from api call
    
    'write the [value] for the [section], [key], and ini file sent
    bStatus = WritePrivateProfileString(szSection, szKey, szValue, szINIFile)
    
    tfnWriteINI = bStatus

End Function
'
'Function : max - returns the maximum of the 2 values passed
'Variables: two variables types of any kind
'Return   : the max of the 2
'
Public Function max(a As Variant, b As Variant) As Variant
    max = -a * (a >= b) - b * (a < b)
End Function
'
'Function : min - returns the minimum of the 2 values passed
'Variables: two variable types of any kind
'Return   : the min of the 2
'
Public Function min(a As Variant, b As Variant) As Variant
    min = -a * (a <= b) - b * (a > b)
End Function
'
'Function : LOWORD - lower 2 bytes of a long
'Variables: long variable
'Return   : integer value of lower 2 bytes
'
Public Function LOWORD(lVal As Long) As Integer
    LOWORD = lVal And MAX_INT
End Function
'
'Function : HIWORD - gets the upper 2 bytes of a long
'Variables: long variable
'Return   : integer value of upper 2 bytes
'
Public Function HIWORD(lVal As Long) As Integer
    HIWORD = lVal& \ MAX_INT
End Function
'
'Function : tfnFixAmpersand - adds ampersand to a string with an ampersand - override default button behavior
'Variables: string to check for ampersand
'Return   : text with any single ampersands replaced with double ampersands
'
Public Function tfnFixAmpersand(ByVal szTextIn As String) As String
    
    Dim szTemp As String 'temp string to hold converted string
    Dim nPos As Integer  'holds the position of the ampersand

    nPos = InStr(szTextIn, szAMPERSAND) 'search for an ampersand
    
    If nPos <> 0 Then 'if no ampersand found just return the original string
        
        szTemp = szEMPTY 'clear the temp string
        
        Do While nPos <> 0 'search for all the ampersnads in the string
            szTemp = szTemp + Left(szTextIn, nPos) + szAMPERSAND 'add another ampersand next to the other
            szTextIn = Mid(szTextIn, nPos + 1) 'strip off substring saved in szTemp
            nPos = InStr(szTextIn, szAMPERSAND) 'search for next ampersand
        Loop
        
        szTemp = szTemp + szTextIn 'save the last part of the original string
        
        tfnFixAmpersand = szTemp 'return the modified string
        Exit Function
        
    End If
    
    tfnFixAmpersand = szTextIn 'no ampersand found return the original string

End Function
'
'Function : tfnIsNull - test for NULL value
'Variables: object to test
'Return   : true if NULL, false if not
'
Public Function tfnIsNull(Value As Variant) As Boolean
    
    Dim szTest As String
    
    On Error GoTo NULL_ERROR
    szTest = Value
        
    tfnIsNull = False
    Exit Function

NULL_ERROR:
    If Err.Number = 94 Then
        tfnIsNull = True
    Else
        tfnIsNull = False
    End If
End Function

'======================
'Resource File Fuctions
'======================

'
'Function : tfnSetToolBarPics -
'Variables: pointer to the form
'Return   : none
'
Public Sub tfnSetToolBarPics(frmWindow As Form)

    On Error Resume Next
    
    Call tfnSetButtonPic(frmWindow.cmdTBPrint, PRINT_DOWN)
    Call tfnSetButtonPic(frmWindow.cmdTBCopy, COPY_DOWN)
    Call tfnSetButtonPic(frmWindow.cmdTBCancel, CANCEL_DOWN)
    Call tfnSetButtonPic(frmWindow.cmdTBExit, EXIT_UP)
    Call tfnSetButtonPic(frmWindow.cmdTBHelp, HELP_UP)
        
End Sub
'
'Function : tfnSetFormLookups -
'Variables: pointer to the form
'Return   : none
'
Public Sub tfnSetFormLookups(frmWindow As Form)
    
    Dim nIndex As Integer
    
    On Error Resume Next
    
    For nIndex = 0 To frmWindow.Controls.Count
        
        If Left(CStr(frmWindow.Controls(nIndex).Tag), 6) = "LOOKUP" Then
            Call tfnSetButtonPic(frmWindow.Controls(nIndex), SEARCH_DOWN)
        End If
    
    Next nIndex
        
End Sub
'
'Function : tfnSetButtonPic - gets a icon from a resource file
'Variables: pointer to the control, resource ID
'Return   : none
'
Public Function tfnSetButtonPic(cmdButton As Control, ResID As Integer) As Boolean

    On Error GoTo ERROR_BadResource 'set runtime error handler
    
    cmdButton.Picture = frmContext.LoadPicture(ResID)
    tfnSetButtonPic = False 'return no error occured reading resource
    Exit Function

ERROR_BadResource:
    tfnSetButtonPic = True 'error reading resource file
End Function

Public Function tfnSQLString(ByVal vTemp As Variant, Optional vNoQuotes As Variant) As String
'
' Properly quotes and formats an SQL string.  If vNoQuotes is present, the result WILL NOT BE QUOTED
' for each ' character found, insert a double ''.  Leave "%* alone
    
    Dim nIdx As Integer
    Dim nPos As Integer
    Dim szParameter As String
    
    If IsNull(vTemp) Then
        szParameter = ""
    Else
        szParameter = vTemp
    End If

    nIdx = 1
    nPos = InStr(nIdx, szParameter, "'")
    
    While nPos <> 0
        szParameter = Left(szParameter, nPos) & "'" & Right(szParameter, Len(szParameter) - nPos)
        nIdx = nPos + 2
        nPos = InStr(nIdx, szParameter, "'")
    Wend
    
    ' quote the whole string - optional
    If IsMissing(vNoQuotes) Then
        tfnSQLString = "'" & szParameter & "'"
    Else
        tfnSQLString = szParameter
    End If

End Function

Public Function tfnSQLCheckPercent(ByRef szParameter As String) As String
'
' adds extra % to string if string uses LIKE statement in the SQL
'

    Dim nIdx As Integer
    Dim nPos As Integer
    
    ' for each %
    nIdx = 1
    nPos = InStr(nIdx, szParameter, "%")
    
    While nPos <> 0
        szParameter = Left(szParameter, nPos) & "%" & Right(szParameter, Len(szParameter) - nPos)
        nIdx = nPos + 2
        nPos = InStr(nIdx, szParameter, "%")
    Wend
    
    tfnSQLCheckPercent = szParameter

End Function

' tfnDisableFormSystemClose
'
' This function disables the close command on the system menu for a form
'
' INPUT:    frmForm - the form to be affected
'
' OUTPUT:   none

Public Sub tfnDisableFormSystemClose(ByRef frmForm As Form, Optional vCloseSize As Variant, Optional bChangeColor As Boolean = True)
    
    #If Win32 Then
        Dim nCode As Long
    #Else
        Dim nCode As Integer
    #End If
    
    Dim bCloseSize As Boolean
    
    If IsMissing(vCloseSize) Then
        bCloseSize = True
    Else
        bCloseSize = vCloseSize
    End If
    
    nCode = GetSystemMenu(frmForm.hwnd, False)
    
    'david 10/27/00
    'the following does not work in windows2000
    'Call ModifyMenu(nCode, SC_CLOSE, 1, 0, "&Close")
    subDisableSystemClose frmForm
    
    'the following work in windows98 ONLY! It does not work in windows2000
    If bCloseSize Then
        Call ModifyMenu(nCode, SC_SIZE, 1, 0, "&Size")
        Call ModifyMenu(nCode, SC_MAXIMIZE, 1, 0, "Ma&ximize")
    End If
    
    'for
    
    'david 10/26/00
    If bChangeColor Then
        tfnFixBackColor frmForm
    End If
    
End Sub

' tfnEnableTBButton
'
' This function enables a toolbar button
'
' INPUT:    ctrlTB      - toolbar button control
'           nResIdx     - resource file index for the bitmap
'
' OUTPUT:   none

Public Sub tfnEnableTBButton(ByRef ctrlTB As Control, ByVal nResIdx As Integer)

    frmContext.ButtonEnabled(nResIdx) = True
End Sub

' tfnDisableTBButton
'
' This function disables a toolbar button
'
' INPUT:    ctrlTB      - toolbar button control
'           nResIdx     - resource file index for the bitmap
'
' OUTPUT:   none

Public Sub tfnDisableTBButton(ByRef ctrlTB As Control, ByVal nResIdx As Integer)

    frmContext.ButtonEnabled(nResIdx) = False
    
End Sub

' tfnTBButtonEnabled
'
' This function can be used to determine is a toolbar button is enabled or not
'
' INPUT:    ctrlTB      - toolbar button control
'
' OUTPUT:   True - button is enabled / False - button is disabled

Public Function tfnTBButtonEnabled(ByVal nID As Integer)

    tfnTBButtonEnabled = frmContext.ButtonEnabled(nID)
    
End Function

'tfnFormatCaption
'
' This function will replicate all the ampersand that is in the input string so that
' the output string can be assigned to the caption of the label or videoelastic control.
'
' use ONLY when you want to assign a string to a LABEL
'
Public Function tfnFormatCaption(ByVal sText As String) As String
    Dim sTemp As String, nPosi As Integer
    
    On Error Resume Next
    sTemp = ""
    Do
        nPosi = InStr(sText, "&")
        If nPosi = 0 Then
            tfnFormatCaption = sTemp & sText
            Exit Function
        End If
        sTemp = Left(sText, nPosi) & "&"
        sText = Mid(sText, nPosi + 1)
    Loop
End Function

'tfnRun
'
' This function is used to run a stand alone program
'
Public Function tfnRun(szExeName As String, _
                       Optional vWindowStyle As Integer = SW_SHOWNORMAL, _
                       Optional bHandShake As Boolean = True, _
                       Optional sParms As String = "") As Boolean

    Dim szCmd As String
    Dim hTempInstance As Long
    
    If InStr(szExeName, "\") <= 0 Then
        #If FACTOR_MENU < 0 Then
            Const gszBINROOT As String = ".\bin\"
        #Else
            Const gszBINROOT As String = "g:\program\factmenu\bin\"
        #End If
    
        szExeName = LCase(Trim(szExeName))
        szCmd = gszBINROOT & szExeName & IIf(InStr(szExeName, ".") = 0, ".exe", "")
    Else
        szCmd = szExeName
    End If

    Const SHELL_OK As Long = 32
    
    On Error GoTo ErrorRun
    
    'check further for the EXE that is in BIN
    If Dir(szCmd) <> "" Then
        'append hand sake string
        If bHandShake Then
            szCmd = szCmd & " " & t_szHandShake
        End If
        
        If Trim(sParms) <> "" Then
            szCmd = szCmd & " " & Trim(sParms)
        End If
        
        hTempInstance = Shell(szCmd, vWindowStyle) 'run the program selected, save the instance handle
        If hTempInstance > SHELL_OK Or hTempInstance < 0 Then 'if hInstance greater than 32 application is running
            tfnRun = True 'application running
            Exit Function
        Else
            tfnRun = False 'application failed to launch
            Exit Function
        End If
    Else
        #If NO_ERROR_HANDLER Then
            MsgBox "Cannot execute program"
        #Else
            tfnErrHandler "tfnRun", 60007, " - " & gszBINROOT & szExeName
        #End If
        tfnRun = False 'application failed to launch
        Exit Function
    End If

    Exit Function

ErrorRun:
    #If NO_ERROR_HANDLER Then
        MsgBox "Cannot execute program" & vbCrLf & Err.Description
    #Else
        tfnErrHandler "tfnRun"
    #End If
End Function

'david 09/28/00
Public Function tfnGetHostName() As String
    'return the current HostName as was logged into factmenu
    #If DEVELOP Or (FACTOR_MENU >= 0) Then
        tfnGetHostName = "ssfactor"
        If t_dbMainDatabase Is Nothing Then Exit Function
            
        tfnGetHostName = tfnGetNamedString(t_dbMainDatabase.Connect, "HOST")
        
        If Trim(tfnGetHostName) = "" Then
            tfnGetHostName = tfnGetNamedString(t_dbMainDatabase.Connect, "SRVR")
        End If
    #Else
'        If t_oleObject Is Nothing Then
            If Not t_dbMainDatabase Is Nothing Then
                tfnGetHostName = tfnGetNamedString(t_dbMainDatabase.Connect, "HOST")
                
                If Trim(tfnGetHostName) = "" Then
                    tfnGetHostName = tfnGetNamedString(t_dbMainDatabase.Connect, "SRVR")
                End If
            End If
'        Else
            'may be not implemented yet
'            tfnGetHostName = t_oleObject.Host
'        End If
    #End If
    
End Function
Public Function tfnGetPassword() As String
    'return the current HostName as was logged into factmenu
    #If DEVELOP Or (FACTOR_MENU >= 0) Then
        tfnGetPassword = "ssfactor"
        If t_dbMainDatabase Is Nothing Then Exit Function
            
        tfnGetPassword = tfnGetNamedString(t_dbMainDatabase.Connect, "PWD")
    #Else
'        If t_oleObject Is Nothing Then
            If Not t_dbMainDatabase Is Nothing Then
                tfnGetPassword = tfnGetNamedString(t_dbMainDatabase.Connect, "PWD")
            End If
'        Else
            'may be not implemented yet
'            tfnGetPassword = t_oleObject.Password
'        End If
    #End If
    
End Function

Public Function tfnGetDataSourceName() As String
    'return the current DataSource Name as was logged into factmenu
    'Robert Atwood 10-29-01
    tfnGetDataSourceName = "ssfactor"
    
    #If DEVELOP Or (FACTOR_MENU >= 0) Then
        If t_dbMainDatabase Is Nothing Then Exit Function
            
        tfnGetDataSourceName = tfnGetNamedString(t_dbMainDatabase.Connect, "DSN")
    #Else
            If Not t_oleObject Is Nothing Then
                tfnGetDataSourceName = t_oleObject.FactorPath
            Else
                'david 11/15/2001
                If Not t_dbMainDatabase Is Nothing Then
                    tfnGetDataSourceName = tfnGetNamedString(t_dbMainDatabase.Connect, "DSN")
                Else
                    tfnGetDataSourceName = ""
                End If
                ''''''''''''''''''''
            End If
    #End If
    
End Function


'david 10/26/00
Public Sub tfnFixBackColor(ByRef frmMain As Form)
    Dim ctrl As Control
    
    On Error Resume Next
    
    frmMain.BackColor = &H8000000F
    
    For Each ctrl In frmMain.Controls
        If TypeOf ctrl Is FactorFrame Then
            If ctrl.BackColor <> &H800000 Then
                ctrl.BackColor = &H8000000F
            End If
        ElseIf TypeOf ctrl Is Label Then
            If ctrl.BorderStyle = 0 Then
                ctrl.BackColor = &H8000000F
            End If
        End If
    Next
End Sub

'david 10/27/00
Public Sub subDisableSystemClose(frmMain As Form)
    Dim hSysMenu As Long
    Dim nCnt As Long
    
    hSysMenu = GetSystemMenu(frmMain.hwnd, False)
    
    If hSysMenu Then
        nCnt = GetMenuItemCount(hSysMenu)
        If nCnt Then
            RemoveMenu hSysMenu, nCnt - 1, MF_BYPOSITION Or MF_REMOVE
            RemoveMenu hSysMenu, nCnt - 2, MF_BYPOSITION Or MF_REMOVE
            DrawMenuBar frmMain.hwnd
        End If
    End If
End Sub

Public Function fnCopyFactorMDB(Optional bShowError As Boolean = True, _
                                Optional sErrMsg As String = "") As Boolean
'##############################################################################
'# Modified to use c:\factor\<datasourcename>\factor.mdb 10-30-01 Robert Atwood
'##############################################################################

    Dim sFactorDir As String
    Dim sWinSysDir As String
    Dim bCopy As Boolean
    
    Dim lFactorMajor As Long
    Dim lFactorMinor As Long
    Dim lFactorRev As Long
    Dim lWinSysMajor As Long
    Dim lWinSysMinor As Long
    Dim lWinSysRev As Long
    
    sFactorDir = LOCAL_FACTOR_PATH
    sWinSysDir = LOCAL_FACTOR_PATH & UCase(Trim(tfnGetDataSourceName)) + "\"
    
    On Error Resume Next
    
    If Dir(sFactorDir + "FACTOR.MDB", vbNormal) = "" Then
        sErrMsg = "FACTOR.MDB does not exist in '" + sFactorDir + "'."
        If bShowError Then
            MsgBox sErrMsg, vbExclamation
        End If
        Exit Function
    End If
    
    If Dir(sWinSysDir + "FACTOR.MDB", vbNormal) = "" Then
        If Dir(sWinSysDir, vbDirectory) = "" Then
            MkDir sWinSysDir
        End If
        
        bCopy = True
    Else
        'check the database version to see we need to copy
        subGetLocalDBVersion lFactorMajor, lFactorMinor, lFactorRev, sFactorDir + "FACTOR.MDB"
        subGetLocalDBVersion lWinSysMajor, lWinSysMinor, lWinSysRev, sWinSysDir + "FACTOR.MDB"
        
        If lFactorMajor > lWinSysMajor Then
            bCopy = True
        Else
            If lFactorMinor > lWinSysMinor Then
                bCopy = True
            Else
                If lFactorRev > lWinSysRev Then
                    bCopy = True
                End If
            End If
        End If
    End If
    
    If bCopy Then
        Dim lRet As Long
        Dim sSRCFile As String
        Dim sDestFile As String
        
        sSRCFile = sFactorDir + "FACTOR.MDB"
        sDestFile = sWinSysDir + "FACTOR.MDB"
        
        lRet = CopyFile(sSRCFile, sDestFile, 0)
    
        If lRet = 0 Then
            sErrMsg = "Failed to copy FACTOR.MDB to '" + sWinSysDir + "'"
             MsgBox sErrMsg, vbExclamation
        End If
    End If
    
    fnCopyFactorMDB = True
    Exit Function
    
errFileInUsed:
    sErrMsg = "'" + sFactorDir + "FACTOR.MDB" + "' is in use by another program."
    
    If bShowError Then
        MsgBox sErrMsg, vbExclamation
    End If
End Function

Private Sub subGetLocalDBVersion(lMajor As Long, _
                                 lMinor As Long, _
                                 lRevision As Long, _
                                 sDBPath As String)

    Dim engLocal As New DBEngine
    Dim dbLocal As DataBase
    Dim wsLocal As Workspace
    Dim strSQL As String
    Dim rsTemp As Recordset
    
    On Error GoTo errOpenDB
    Set wsLocal = engLocal.Workspaces(0)
    Set dbLocal = wsLocal.OpenDatabase(sDBPath, , True)
    strSQL = "SELECT nMajor, nMinor, nRevision FROM SysVersion"
    Set rsTemp = dbLocal.OpenRecordset(strSQL)
    
    If rsTemp.RecordCount > 0 Then
        If Not IsNull(rsTemp!nMajor) Then
            lMajor = Trim(rsTemp!nMajor)
        End If
        If Not IsNull(rsTemp!nMinor) Then
            lMinor = Trim(rsTemp!nMinor)
        End If
        If Not IsNull(rsTemp!nRevision) Then
            lRevision = Trim(rsTemp!nRevision)
        End If
    End If
    
    dbLocal.Close
    Set dbLocal = Nothing
    Set wsLocal = Nothing
    Set engLocal = Nothing
    
    Exit Sub

errExitHere:
    lMajor = -1
    lMinor = -1
    lRevision = -1
    
    Exit Sub

errOpenDB:
    If Err.Number = 3051 Then
        On Error GoTo errSetAttr
        SetAttr sDBPath, vbNormal
        Resume
    Else
        Resume errExitHere
    End If

errSetAttr:
    Resume errExitHere
End Sub


''''''''''''''''''''Sam Zheng on 08/24/2001 ''''''''''''''''''''''''''''''
' for inv cross reference:
'step 1: call subCreateTemp_inv_header in form_load( before setup combos;
'step 2: in each setup combos' SQL, replace inv_header by tmpx_inv_header;
'step 3: in product code validate function, call fnInv_xref_check(sCode)
'        just after the empty( if stext = "" then ... end if) checking.

Public Sub subCreateTemp_inv_header()
    Dim strSQL As String
    
    On Error Resume Next
    strSQL = "drop table tmpx_inv_header"
    t_dbMainDatabase.ExecuteSQL strSQL
    
    strSQL = " select * from inv_header into temp tmpx_inv_header "
    t_dbMainDatabase.ExecuteSQL strSQL
    
    If Not tfnNeed_inv_xref Then
       Exit Sub
    End If
    
    strSQL = " insert into tmpx_inv_header " _
           & " select ivx_old_nbr, ivh_link,ivh_prodtcl,ivh_print,ivh_xref," _
           & " ivh_desc,ivh_class,ivh_spec_part,ivh_uom_stock,ivh_uom_sales," _
           & " ivh_uom_pricing,ivh_brand,ivh_uom_purch,ivh_assoc_prodlnk," _
           & " ivh_fet_amt,ivh_rpt_factor,ivh_active,ivh_stocking " _
           & " from inv_header, inv_xref " _
           & " where ivx_new_nbr = ivh_product"
    t_dbMainDatabase.ExecuteSQL strSQL
End Sub

Public Function fnInv_xref_check(ByVal sCode As String) As String
    Dim strSQL As String
    Dim rsTemp As Recordset
    
    fnInv_xref_check = Trim(sCode)
    If Not tfnNeed_inv_xref Then
        Exit Function
    End If
    
    strSQL = "select ivx_new_nbr from inv_xref where ivx_old_nbr= " _
           & tfnSQLString(sCode)
    Set rsTemp = t_dbMainDatabase.OpenRecordset(strSQL, dbOpenSnapshot, SQL_PASSTHROUGH)
    If Not rsTemp.EOF Then
        If Not IsNull(rsTemp!ivx_new_nbr) Then
            fnInv_xref_check = Trim$(rsTemp!ivx_new_nbr)
        End If
    End If
    rsTemp.Close
End Function

Public Function tfnNeed_inv_xref() As Boolean
    Dim strSQL As String
    Dim rsTemp As Recordset
    On Error GoTo ErrTrap
    
    If Not (SYS_PARM_6005 = "Y" Or SYS_PARM_6005 = "N") Then
        SYS_PARM_6005 = "N"
        strSQL = "SELECT parm_field FROM sys_parm WHERE parm_nbr = 6005"
        Set rsTemp = t_dbMainDatabase.OpenRecordset(strSQL, dbOpenSnapshot, SQL_PASSTHROUGH)
        If Not rsTemp.EOF Then
            If Not IsNull(rsTemp!parm_field) And UCase(Trim$("" & rsTemp!parm_field)) = "Y" Then
                SYS_PARM_6005 = "Y"
            End If
        End If
        rsTemp.Close
    End If
    
    If SYS_PARM_6005 = "Y" Then
        tfnNeed_inv_xref = True
    Else
        tfnNeed_inv_xref = False
    End If
    
    Exit Function
    
ErrTrap:
    tfnNeed_inv_xref = False

End Function

'nMinWidth x nMinHeight
'640 x 480
'800 x 600
'1024 x 768
Public Function fnScreenResolution(Optional nMinWidth As Integer = 0, _
                                   Optional nMinHeight As Integer = 0, _
                                   Optional bSetScreenResolution As Boolean = False, _
                                   Optional bAskToChange As Boolean = True) As Boolean

    Static nScreenWidth As Integer
    Static nScreenHeight As Integer
    
    Static bChangeResolutionTo800x600 As Boolean
    
    If Not bSetScreenResolution Then  'restore screen resolution
        If bChangeResolutionTo800x600 Then
            If nScreenWidth > 0 And nScreenHeight > 0 Then
                fnScreenResolution = fnSetScreenResolution(nScreenWidth, nScreenHeight)
            End If
        End If
    End If
    
    nScreenWidth = GetSystemMetrics(SM_CXSCREEN)
    nScreenHeight = GetSystemMetrics(SM_CYSCREEN)

    'error ???
    If nScreenWidth = 0 Or nScreenHeight = 0 Then
        fnScreenResolution = True
        Exit Function
    End If
    
    If nScreenWidth >= nMinWidth And nScreenHeight >= nMinHeight Then
        fnScreenResolution = True
        Exit Function
    End If
    
    If bAskToChange Then
        bChangeResolutionTo800x600 = MsgBox("This program is designed to run on the windows with screen resolution" _
            + " of 800x600 or higher." + vbCrLf + vbCrLf + "Your screen resolution is " + CStr(nScreenWidth) + "x" _
            + CStr(nScreenHeight) + "." + vbCrLf + vbCrLf + "Do you want the program change the screen resolution" _
            + " to 800x600?", vbQuestion + vbYesNo + vbDefaultButton2)
    End If
    
    If Not bChangeResolutionTo800x600 Then
        Exit Function
    End If
    
    'change screen resolution to 800x600
    fnScreenResolution = fnSetScreenResolution(800, 600)
End Function

Public Function fnSetScreenResolution(nScreenWidth As Integer, nScreenHeight As Integer) As Boolean
    'Code:
    Dim typDevM As typDevMODE
    Dim lngResult As Long
    Dim intAns    As Integer
    
    ' Retrieve info about the current graphics mode
    ' on the current display device.
    lngResult = EnumDisplaySettings(0, 0, typDevM)
    
    ' Set the new resolution. Don't change the color
    ' depth so a restart is not necessary.
    With typDevM
        .dmFields = DM_PELSWIDTH Or DM_PELSHEIGHT
        .dmPelsWidth = CLng(nScreenWidth)  'ScreenWidth (640,800,1024, etc)
        .dmPelsHeight = CLng(nScreenHeight) 'ScreenHeight (480,600,768, etc)
    End With
    
    ' Change the display settings to the specified graphics mode.
    lngResult = ChangeDisplaySettings(typDevM, CDS_TEST)
    
    Select Case lngResult
    Case DISP_CHANGE_RESTART
        MsgBox "You need to restart your computer to apply these changes.", vbInformation + vbSystemModal, "Screen Resolution"
        
        'intAns = MsgBox("You must restart your computer to apply these changes." & _
            vbCrLf & vbCrLf & "Do you want to restart now?", _
            vbYesNo + vbSystemModal, "Screen Resolution")
        'If intAns = vbYes Then Call ExitWindowsEx(EWX_REBOOT, 0)
    Case DISP_CHANGE_SUCCESSFUL
        Call ChangeDisplaySettings(typDevM, CDS_UPDATEREGISTRY)
        Call SendMessage(HWND_BROADCAST, WM_DISPLAYCHANGE, SPI_SETNONCLIENTMETRICS, ByVal 0&)
        'MsgBox "Screen resolution changed.", vbInformation + vbSystemModal, "Resolution Changed"
        fnSetScreenResolution = True
    Case Else
        MsgBox "Mode not supported", vbSystemModal, "Error"
    End Select
End Function

'david 04/30/2002
'#367575
Public Function fnRemoveChr0(vText) As String
    Dim sText As String
    Dim sTemp As String
    Dim sChar As String
    Dim I As Long
    
    sText = vText & ""
    
    sTemp = ""
    
    If sText <> "" Then
        If InStrB(sText, Chr(0)) > 0 Then
            For I = 1 To Len(sText)
                sChar = Mid(sText, I, 1)
                
                If sChar <> Chr(0) Then
                    sTemp = sTemp + sChar
                End If
            Next I
        
            sTemp = RTrim(sText)
        Else
            sTemp = sText
        End If
    End If
    
    fnRemoveChr0 = sTemp
End Function
'''''''''''''''


