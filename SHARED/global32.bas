'
'               Visual Basic Declarations of CRPE32.DLL
'               =====================================
'
'       File:         GLOBAL32.BAS
'
'       Author:       Seagate Software Information Management Group, Inc.
'       Date:         15 Apr 92
'
'       Purpose:      This file presents the API to the Crystal Reports
'                     Print Engine DLL (Professional).
'
'       Language:     Visual Basic for Windows
'
'       Copyright (c) 1992-1997 Seagate Software Information Management Group, Inc.
'
'       Revisions:
'
'          CCS  15 Apr 92  -  Original Development
'          KYL  12 Jul 92  -  Modified Existing Declarations
'                             Added Missing Declarations
'          KYL  27 Aug 92  -  Converted to CRPE32.DLL
'          CRD  08 Feb 93  -  Added new calls for 2.0 and Global declares for samples
'          CRD  25 Feb 93  -  Added new calls for 2.0 Pro
'          RBC  23 Apr 93  -  Added more new calls, rearranged to match CRPE.H
'          DVA  22 Dec 93  -  Added new calls for 3.0
'          TW   15 Mar 94  -  3.0 call reorganization
'          RS   28 Aug 95  -  32-bit update
'          JEA  12 May 96  -  Revised for 5.0
'                             Added the following Error Codes
'                               PE_ERR_BADSECTIONHEIGHT     = 560
'                               PE_ERR_BADVALUETYPE         = 561
'                               PE_ERR_INVALIDSUBREPORTNAME = 562
'                               PE_ERR_FIELDEXIST           = 563
'                               PE_ERR_NOPARENTWINDOW       = 564
'                               PE_ERR_INVALIDZOOMFACTOR    = 565
'                             Added the Following Constants
'                               PE_WORD_LEN = 2
'                               PE_PF_NAME_LEN = 256
'                               PE_PF_PROMPT_LEN = 256
'                               PE_PF_VALUE_LEN = 256
'                               PE_PF_NUMBER = 0
'                               PE_PF_CURRENCY = 1
'                               PE_PF_BOOLEAN = 2
'                               PE_PF_DATE = 3
'                               PE_PF_STRING = 4
'                               PE_SIZEOF_VARINFO_TYPE
'                               PE_SUBREPORT_NAME_LEN = 128
'                               PE_SIZEOF_SUBREPORT_INFO
'                               PE_PARAMETER_NAME_LEN = 128
'                               PE_SIZEOF_PARAMETER_INFO
'                               PE_SECT_PAGE_HEADER = 2
'                               PE_SECT_PAGE_FOOTER = 7
'                               PE_SECT_REPORT_HEADER = 1
'                               PE_SECT_REPORT_FOOTER = 8
'                               PE_SECT_GROUP_HEADER = 3
'                               PE_SECT_GROUP_FOOTER = 5
'                               PE_SECT_DETAIL = 4
'                             Added the following Structures
'                               PEParameterFieldInfo
'                               PESubreportInfo
'                               PEParameterInfo
'                             Added the following Declarations
'                               PEGetNParameterFields
'                               PEGetNthParameterField
'                               PESetNthParameterField
'                               PEGetNSubreportsInSection
'                               PEGetNthSubreportInSection
'                               PEGetSubreportInfo
'                               PEOpenSubreport
'                               PECloseSubreport
'                               PESetDialogParentWindow
'                               PEEnableProgressDialog
'                               PEGetNPages
'                               PEShowNthPage
'                               PEGetNSections
'                               PEGetSectionCode
'                               PEGetNthParamInfo
'                             Added the following Functions
'                               PE_SECTION_CODE
'                               PE_SECTION_TYPE
'                               PE_GROUP_N
'                               PE_SECTION_N
'          JEA 13 May 96  -  Changed PELogOnServerWithPrivateInfo 2nd parameter from
'                            PrivateInfo As Any to ByVal PrivateInfo As Long
'          JEA 02 Jun 96  -  Added Following Constants
'                              PE_UNCHANGED_COLOR = -2
'                            Added Following Structure Members
'                              PESectionOptions.underlaySection
'                              PESectionOptions.backgroundColor
'          JEA 27 Jun 96  -  Added the Following Constants
'                              PE_SIZEOF_EXPORT_OPTIONS
'                              PE_SIZEOF_GRAPH_DATA_INFO
'                              PE_SIZEOF_GRAPH_TEXT_INFO
'                              PE_SIZEOF_GRAPH_OPTIONS
'                              PE_SIZEOF_JOB_INFO
'                              PE_SIZEOF_PRINT_FILE_OPTIONS
'                              PE_SIZEOF_CHAR_SEP_FILE_OPTIONS
'                              PE_SIZEOF_PRINT_OPTIONS
'                              PE_SIZEOF_SECTION_OPTIONS
'          JEA 08 Jul 96  -  Added reserve members to these structures for alignment:
'                              Type PEPrintFileOptions
'                                  StructSize As Integer   ' initialize to # of bytes in PEPrintFileOptions
'
'                                  UseReportNumberFmt As Integer
'                                  reserved1 As Integer     ' reserved - do not set
'                                  UseReportDateFormat As Integer
'                                  reserved2 As Integer     ' reserved - do not set
'                              End Type
'
'                              Type PECharSepFileOptions
'                                  StructSize As Integer   ' initialize to PE_SIZEOF_CHAR_SEP_FILE_OPTIONS
'
'                                  UseReportNumberFmt As Integer
'                                  reserved1 As Integer     ' reserved - do not set
'                                  UseReportDateFormat As Integer
'                                  reserved2 As Integer     ' reserved - do not set
'
'                                  StringDelimiter As String * 1
'                                  FieldDelimiter As String * PE_FIELDDELIMLEN
'                              End Type
'         JEA 09 Jul 96  -  Changed Jobinfo Structure declaration
'                           from
'                           PEJobInfo
'                               PrintEnded As Integer
'                           PEJobInfo
'                               PrintEnded As Long
'
'         JEA 11 Jul 96  -  Removed declarations for the following calls
'                            PESetLineHeight
'                            PEGetNLinesInSection
'                            PEGetLineHeight
'                            For old code:
'                               PESetLineHeight now calls PESetMinimumSectionHeight
'                               PEGetLineHeight now calls PEGetMinimumSectionHeight
'                               PEGetNLineInSection always returns 1.
'
'                            Changed the DEVMODE argument of the follow calls to As Any
'                               PESelectPrinter
'                               PEGetSelectedPrinter
'          JEA 02 Aug 96  -  Added the following error code:
'                               Global Const PE_ERR_PAGESIZEOVERFLOW = 567
'                               Global Const PE_ERR_LOWSYSTEMRESOURCES = 568
'                         -  Added the following format formula name constants:
'                               Global Const SECTION_VISIBILITY = 58
'                               Global Const NEW_PAGE_BEFORE = 60
'                               Global Const NEW_PAGE_AFTER = 61
'                               Global Const KEEP_SECTION_TOGETHER = 62
'                               Global Const SUPPRESS_BLANK_SECTION = 63
'                               Global Const RESET_PAGE_N_AFTER = 64
'                               Global Const PRINT_AT_BOTTOM_OF_PAGE = 65
'                               Global Const UNDERLAY_SECTION = 66
'                               Global Const SECTION_BACK_COLOUR = 67
'                         -  Added the following function
'Declare Function PESetSectionFormatFormula Lib "crpe32.dll" (ByVal printJob%, ByVal sectionCode%, ByVal formulaName%, ByVal formulaString$) As Integer
'                         -  Changed structure member from
'                              PESectionOptions.SuppressBlankLines
'                              to PESectionOptions.SuppressBlankSection
'          JEA 07 Aug 96  -  Added Function PEVBGetVersion
'          JEA 09 Aug 96  -  Added Function To re-direct PEGetJobStatus
'                            Changed PEJobInfo.numRecordsRead,
'                                    PEJobInfo.numRecordsSelected,
'                                    PEJobInfo.numRecordsPrinted back to single Longs.
'          JEA 23 Aug 96  -  Added
'                            showArea As Integer
'                            freeFormPlacement As Integer
'                            members to PESectionOptions structure
'          JEA 28 Aug 96  -  Changed the Following Declarations:
'                            Changed
'                            Declare Function PEHasSavedData Lib "crpe32.dll" (ByVal printJob%, HasSavedData%) As Integer
'                            to
'                            Declare Function PEHasSavedData Lib "crpe32.dll" (ByVal printJob%, HasSavedData As Long) As Integer
'                            Changed
'                            Declare Function PEPrintControlsShowing Lib "crpe32.dll" (ByVal printJob%, ControlsShowing%) As Integer
'                            to
'                            Declare Function PEPrintControlsShowing Lib "crpe32.dll" (ByVal printJob%, ControlsShowing As Long) As Integer
'          JEA 07 OCT 96  -  Changed
'                            PEGraphOptions.ShowDataValue
'                            PEGraphOptions.ShowGridLine
'                            PEGraphOptions.VerticalBars
'                            PEGraphOptions.ShowLegend to Longs to match the CRPE32.DLL which
'                            expects 4 Byte BOOLS.
'                            Changed
'                            Global Const PE_SIZEOF_GRAPH_OPTIONS = 5 * PE_WORD_LEN + 2 * 8 + PE_GRAPH_TEXT_LEN
'                            to
'                            Global Const PE_SIZEOF_GRAPH_OPTIONS = PE_WORD_LEN + 2 * 8 + 4 * 4 + PE_GRAPH_TEXT_LEN
'                            to accomodate new longs.
'          CGH 10 OCT 96  -  Changed PEPrintReport window size parameters to Longs to accomadate CW_USEDEFAULT
'          JEA 15 Oct 96  -  Added the following declarations
'                              PESetAreaFormat
'                              PESetAreaFormatFormula
'                              PEGetAreaFormat
'                            Added function to create area codes:
'                              Function PE_AREA_CODE(sectionType%, groupN%) As Integer
'          JEA 04 Nov 96  -  Added constant
'                              PE_SIZEOF_PARAMETER_FIELD_INFO
'          JEA 09 JAN 97  -  Added new naming convention (PE_FFN) for format formula name constants.
'                         -  Added new format formula name constant PE_FFN_SHOW_AREA = 59
'                         -  Added the following constants
'                              Global Const PE_VI_STRING_LEN = 256
'                              Global Const PE_BYTE_LEN = 1
'                              Global Const PE_LONG_LEN = 4
'                              Global Const PE_DOUBLE_LEN = 8
'                              Global Const PE_VI_STRING_LEN = 256
'                              Global Const PE_SIZEOF_VALUE_INFO = 1 * PE_BYTE_LEN + _
'                                                                 15 * PE_WORD_LEN + _
'                                                                  2 * PE_LONG_LEN + _
'                                                                  2 * PE_DOUBLE_LEN + _
'                                                                  1 * PE_VI_STRING_LEN
'                         -  Added the following Type
'                              Type PEValueInfo
'                         -  Added the following declarations
'                              Declare Function PEConvertPFInfoToVInfo...
'                              Declare Function PEConvertVInfoToPFInfo...
'                         -  Added the following Error Codes
'                             Global Const PE_ERR_BADGROUPNUMBER = 570
'                             Global Const PE_ERR_INVALIDNEGATIVEVALUE = 572
'                             Global Const PE_ERR_INVALIDMEMORYPOINTER = 573
'                             Global Const PE_ERR_INVALIDPARAMETERNUMBER = 594
'                             Global Const PE_ERR_SQLSERVERNOTOPENED = 599

'          JEA 14 JAN 97  -  Changed
'                              PE_SECTION_TYPE = ((sectionCode) / 1000)
'                                to
'                              PE_SECTION_TYPE = ((sectionCode) \ 1000)
'                                to truncate instead of round.
'                            Changed
'                              PE_SECTION_N = (((sectionCode / 25) Mod 40))
'                                to
'                              PE_SECTION_N = (((sectionCode \ 25) Mod 40))
'                                to truncate instead of round.
'          JEA 31 JAN 97  - Added constant
'                              Global Const PE_PF_REPORT_NAME_LEN = 128
'                         - Added Structure Members
'                              PEParameterFieldInfo.ReportName
'                              PEParameterFieldInfo.needsCurrentValue
'          JEA 29 APR 97  - Added Constant
'                              PE_SIZEOF_WINDOW_OPTIONS
'                              PE_FFN_AREASECTION_VISIBILITY
'                           Added Type
'                              PEWindowOptions
'                           Added Functions
'                              PEGetSectionFormatFormula
'                              PEGetAreaFormatFormula
'                              PEGetWindowOptions
'                              PESetWindowOptions
'          JEA 20 MAY 97  - Added Constants
'                              PE_VI_NUMBER
'                              PE_VI_CURRENCY
'                              PE_VI_BOOLEAN
'                              PE_VI_DATE
'                              PE_VI_STRING
'                              PE_VI_DATETIME
'                              PE_VI_TIME
'                              PE_VI_INTEGER
'                              PE_VI_COLOR
'                              PE_VI_CHAR
'                              PE_VI_LONG
'                              PE_VI_NOVALUE
'                           Added Structure Members
'                              PEWindowOptions.hasPrintSetupButton
'                              PEWindowOptions.hasRefreshButton
'                              PEValueInfo.ignored
'                              PEValueInfo.viLong
'                           Updated Constants
'                              PE_SIZEOF_WINDOW_OPTIONS
'                              PE_SIZEOF_VALUE_INFO
'          JEA 14 JUL 97  - Added Constants
'                              PE_GO_TBN_ALL_GROUPS_UNSORTED
'                              PE_GO_TBN_ALL_GROUPS_SORTED
'                              PE_GO_TBN_TOP_N_GROUPS
'                              PE_GO_TBN_BOTTOM_N_GROUPS
'                              PE_SF_ORIGINAL
'                              PE_SF_SPECIFIED
'                              PE_SIZEOF_GROUP_OPTIONS
'                              PE_SI_APPLICATION_NAME_LEN
'                              PE_SI_TITLE_LEN
'                              PE_SI_SUBJECT_LEN
'                              PE_SI_AUTHOR_LEN
'                              PE_SI_KEYWORDS_LEN
'                              PE_SI_COMMENTS_LEN
'                              PE_SI_REPORT_TEMPLATE_LEN
'                              PE_SIZEOF_REPORT_SUMMARY_INFO
'                              PE_PT_LONGVARCHAR
'                              PE_PT_BINARY
'                              PE_PT_VARBINARY
'                              PE_PT_LONGVARBINARY
'                              PE_PT_BIGINT
'                              PE_PT_TINYINT
'                              PE_PT_BIT
'                              PE_PT_CHAR
'                              PE_PT_NUMERIC
'                              PE_PT_DECIMAL
'                              PE_PT_INTEGER
'                              PE_PT_SMALLINT
'                              PE_PT_FLOAT
'                              PE_PT_REAL
'                              PE_PT_DOUBLE
'                              PE_PT_DATE
'                              PE_PT_TIME
'                              PE_PT_TIMESTAMP
'                              PE_PT_VARCHAR
'                           Added Types
'                              PEGroupOptions
'                              PEReportSummaryInfo
'                           Added Functions
'                              PEGetGroupOptions
'                              PESetGroupOptions
'                              PEGetReportSummaryInfo
'                              PESetReportSummaryInfo


Global Const PE_UNCHANGED_COLOR = -2

Global Const PE_ERR_NOERROR = 0

Global Const PE_ERR_NOTENOUGHMEMORY = 500
Global Const PE_ERR_INVALIDJOBNO = 501
Global Const PE_ERR_INVALIDHANDLE = 502
Global Const PE_ERR_STRINGTOOLONG = 503
Global Const PE_ERR_NOSUCHREPORT = 504
Global Const PE_ERR_NODESTINATION = 505
Global Const PE_ERR_BADFILENUMBER = 506
Global Const PE_ERR_BADFILENAME = 507
Global Const PE_ERR_BADFIELDNUMBER = 508
Global Const PE_ERR_BADFIELDNAME = 509
Global Const PE_ERR_BADFORMULANAME = 510
Global Const PE_ERR_BADSORTDIRECTION = 511
Global Const PE_ERR_ENGINENOTOPEN = 512
Global Const PE_ERR_INVALIDPRINTER = 513
Global Const PE_ERR_PRINTFILEEXISTS = 514
Global Const PE_ERR_BADFORMULATEXT = 515
Global Const PE_ERR_BADGROUPSECTION = 516
Global Const PE_ERR_ENGINEBUSY = 517
Global Const PE_ERR_BADSECTION = 518
Global Const PE_ERR_NOPRINTWINDOW = 519
Global Const PE_ERR_JOBALREADYSTARTED = 520
Global Const PE_ERR_BADSUMMARYFIELD = 521
Global Const PE_ERR_NOTENOUGHSYSRES = 522
Global Const PE_ERR_BADGROUPCONDITION = 523
Global Const PE_ERR_JOBBUSY = 524
Global Const PE_ERR_BADREPORTFILE = 525
Global Const PE_ERR_NODEFAULTPRINTER = 526
Global Const PE_ERR_SQLSERVERERROR = 527
Global Const PE_ERR_BADLINENUMBER = 528
Global Const PE_ERR_DISKFULL = 529
Global Const PE_ERR_FILEERROR = 530
Global Const PE_ERR_INCORRECTPASSWORD = 531
Global Const PE_ERR_BADDATABASEDLL = 532
Global Const PE_ERR_BADDATABASEFILE = 533
Global Const PE_ERR_ERRORINDATABASEDLL = 534
Global Const PE_ERR_DATABASESESSION = 535
Global Const PE_ERR_DATABASELOGON = 536
Global Const PE_ERR_DATABASELOCATION = 537
Global Const PE_ERR_BADSTRUCTSIZE = 538
Global Const PE_ERR_BADDATE = 539
Global Const PE_ERR_BADEXPORTDLL = 540
Global Const PE_ERR_ERRORINEXPORTDLL = 541
Global Const PE_ERR_PREVATFIRSTPAGE = 542
Global Const PE_ERR_NEXTATLASTPAGE = 543
Global Const PE_ERR_CANNOTACCESSREPORT = 544
Global Const PE_ERR_USERCANCELLED = 545
Global Const PE_ERR_OLE2NOTLOADED = 546
Global Const PE_ERR_BADCROSSTABGROUP = 547
Global Const PE_ERR_NOCTSUMMARIZEDFIELD = 548
Global Const PE_ERR_DESTINATIONNOTEXPORT = 549
Global Const PE_ERR_INVALIDPAGENUMBER = 550
Global Const PE_ERR_NOTSTOREDPROCEDURE = 552
Global Const PE_ERR_INVALIDPARAMETER = 553
Global Const PE_ERR_GRAPHNOTFOUND = 554
Global Const PE_ERR_INVALIDGRAPHTYPE = 555
Global Const PE_ERR_INVALIDGRAPHDATA = 556
Global Const PE_ERR_CANNOTMOVEGRAPH = 557
Global Const PE_ERR_INVALIDGRAPHTEXT = 558
Global Const PE_ERR_INVALIDGRAPHOPT = 559

'New Error Codes For 5.0
Global Const PE_ERR_BADSECTIONHEIGHT = 560
Global Const PE_ERR_BADVALUETYPE = 561
Global Const PE_ERR_INVALIDSUBREPORTNAME = 562
Global Const PE_ERR_NOPARENTWINDOW = 564     'dialog parent window
Global Const PE_ERR_INVALIDZOOMFACTOR = 565  'zoom factor
Global Const PE_ERR_PAGESIZEOVERFLOW = 567
Global Const PE_ERR_LOWSYSTEMRESOURCES = 568
Global Const PE_ERR_BADGROUPNUMBER = 570
Global Const PE_ERR_INVALIDNEGATIVEVALUE = 572
Global Const PE_ERR_INVALIDMEMORYPOINTER = 573
Global Const PE_ERR_INVALIDPARAMETERNUMBER = 594
Global Const PE_ERR_SQLSERVERNOTOPENED = 599


Global Const PE_ERR_NOTIMPLEMENTED = 999

'Constants using to calculate structure size constants
Global Const PE_BYTE_LEN = 1
Global Const PE_WORD_LEN = 2
Global Const PE_LONG_LEN = 4
Global Const PE_DOUBLE_LEN = 8

' Open, print and close report (used when no changes needed to report)
' --------------------------------------------------------------------

Declare Function PEPrintReport Lib "crpe32.dll" (ByVal RptName$, ByVal Printer%, ByVal Window%, ByVal title$, ByVal Lft&, ByVal Top&, ByVal Wdth&, ByVal Height&, ByVal style As Long, ByVal PWindow As Long) As Integer


' Open and close print engine
' ---------------------------

Declare Function PEOpenEngine Lib "crpe32.dll" () As Integer
Declare Sub PECloseEngine Lib "crpe32.dll" ()
Declare Function PECanCloseEngine Lib "crpe32.dll" () As Integer


' Get version info
' ----------------

Global Const PE_GV_DLL = 100      ' values for version parameter of PEGetVersion
Global Const PE_GV_ENGINE = 200

Declare Function PEGetVersion Lib "crpe32.dll" (ByVal version%) As Integer


' Open and close print job (i.e. report)
' --------------------------------------

Declare Function PEOpenPrintJob Lib "crpe32.dll" (ByVal RptName$) As Integer
Declare Sub PEClosePrintJob Lib "crpe32.dll" (ByVal printJob%)


' Start and cancel print job (i.e. print the report, usually after changing report)
' ---------------------------------------------------------------------------------

Declare Function PEStartPrintJob Lib "crpe32.dll" (ByVal printJob%, ByVal WaitOrNot%) As Integer

Declare Sub PECancelPrintJob Lib "crpe32.dll" (ByVal printJob%)

' Print job status
' ----------------
 
Declare Function PEIsPrintJobFinished Lib "crpe32.dll" (ByVal printJob%) As Integer

' To work around the problem of 4 - Byte alignment the PEGetJobStatus
' call has been re-declared locally. When your application calls PEGetJobStatus
' it is calling a function in this file which in turn calls CRPE32.DLL.
Global Const PE_JOBNOTSTARTED = 1
Global Const PE_JOBINPROGRESS = 2
Global Const PE_JOBCOMPLETED = 3
Global Const PE_JOBFAILED = 4
Global Const PE_JOBCANCELLED = 5

Type PEJobInfo
    StructSize As Integer  ' initialize to PE_SIZEOF_JOB_INFO

    NumRecordsRead As Long
    NumRecordsSelected As Long
    NumRecordsPrinted As Long

    DisplayPageN As Integer
    LatestPageN As Integer
    StartPageN As Integer

    PrintEnded As Long
End Type

Type SplitPEJobInfo
    StructSize As Integer  ' initialize to PE_SIZEOF_JOB_INFO

    NumRecordsRead1 As Integer
    NumRecordsRead2 As Integer
    NumRecordsSelected1 As Integer
    NumRecordsSelected2 As Integer
    NumRecordsPrinted1 As Integer
    NumRecordsPrinted2 As Integer

    DisplayPageN As Integer
    LatestPageN As Integer
    StartPageN As Integer

    PrintEnded As Long
End Type

Global Const PE_SIZEOF_JOB_INFO = 10 * PE_WORD_LEN + 4

Declare Function RealPEGetJobStatus Lib "crpe32.dll" Alias "PEGetJobStatus" (ByVal printJob%, JobInfo As SplitPEJobInfo) As Integer

' Controlling dialogs
' -------------------

Declare Function PESetDialogParentWindow Lib "crpe32.dll" (ByVal printJob%, ByVal parentWindow As Long) As Integer

Declare Function PEEnableProgressDialog Lib "crpe32.dll" (ByVal printJob%, ByVal enable%) As Integer


' Print job error codes and messages
' ----------------------------------

Declare Function PEGetErrorCode Lib "crpe32.dll" (ByVal printJob%) As Integer

Declare Function PEGetErrorText Lib "crpe32.dll" (ByVal printJob%, textHandle As Long, textLength%) As Integer

Declare Function PEGetHandleString Lib "crpe32.dll" (ByVal textHandle As Long, ByVal Buffer$, ByVal BufferLength%) As Integer


' Setting the print date
' ----------------------

Declare Function PEGetPrintDate Lib "crpe32.dll" (ByVal printJob%, Date_Year%, Date_Month%, Date_Day%) As Integer

Declare Function PESetPrintDate Lib "crpe32.dll" (ByVal printJob%, ByVal Date_Year%, ByVal Date_Month%, ByVal Date_Day%) As Integer

' Encoding and Decoding Section Codes
' -----------------------------------

Global Const PE_ALLSECTIONS = 0

'Section types for use with PE_SECTION_CODE, PE_SECTION_TYPE, PE_GROUP_N and PE_SECTION_N functions
Global Const PE_SECT_PAGE_HEADER = 2
Global Const PE_SECT_PAGE_FOOTER = 7
Global Const PE_SECT_REPORT_HEADER = 1
Global Const PE_SECT_REPORT_FOOTER = 8
Global Const PE_SECT_GROUP_HEADER = 3
Global Const PE_SECT_GROUP_FOOTER = 5
Global Const PE_SECT_DETAIL = 4

'The old section constants with comment showing them in terms of the new:
'(Note that PE_GRANDTOTALSECTION and PE_SUMMARYSECTION both map
' to PE_SECT_REPORT_FOOTER.)

Global Const PE_HEADERSECTION = 2000  'PE_SECTION_CODE (PE_SECT_PAGE_HEADER,   0, 0)
Global Const PE_FOOTERSECTION = 7000  'PE_SECTION_CODE (PE_SECT_PAGE_FOOTER,   0, 0)
Global Const PE_TITLESECTION = 1000   'PE_SECTION_CODE (PE_SECT_REPORT_HEADER, 0, 0)
Global Const PE_SUMMARYSECTION = 8000 'PE_SECTION_CODE (PE_SECT_REPORT_FOOTER, 0, 0)
Global Const PE_GROUPHEADER = 3000    'PE_SECTION_CODE (PE_SECT_GROUP_HEADER,  0, 0)
Global Const PE_GROUPFOOTER = 5000    'PE_SECTION_CODE (PE_SECT_GROUP_FOOTER,  0, 0)
Global Const PE_DETAILSECTION = 4000  'PE_SECTION_CODE (PE_SECT_DETAIL,        0, 0)
Global Const PE_GRANDTOTALSECTION = PE_SUMMARYSECTION


' Controlling group conditions (i.e. group breaks)
' ------------------------------------------------
Global Const PE_SF_MAX_NAME_LENGTH = 50

Global Const PE_SF_DESCENDING = 0
Global Const PE_SF_ASCENDING = 1
Global Const PE_SF_ORIGINAL = 2 'only for group condition
Global Const PE_SF_SPECIFIED = 3 'only for group condition

' use PE_ANYCHANGE for all field types except Date
Global Const PE_GC_ANYCHANGE = 0

' use these constants for Date fields
Global Const PE_GC_DAILY = 0
Global Const PE_GC_WEEKLY = 1
Global Const PE_GC_BIWEEKLY = 2
Global Const PE_GC_SEMIMONTHLY = 3
Global Const PE_GC_MONTHLY = 4
Global Const PE_GC_QUARTERLY = 5
Global Const PE_GC_SEMIANNUALLY = 6
Global Const PE_GC_ANNUALLY = 7

' use these constants for Boolean fields
Global Const PE_GC_TOYES = 1
Global Const PE_GC_TONO = 2
Global Const PE_GC_EVERYYES = 3
Global Const PE_GC_EVERYNO = 4
Global Const PE_GC_NEXTISYES = 5
Global Const PE_GC_NEXTISNO = 6

Declare Function PESetGroupCondition Lib "crpe32.dll" (ByVal printJob%, ByVal sectionCode%, ByVal ConditionField$, ByVal condition%, ByVal sortDirection%) As Integer

Declare Function PEGetNGroups Lib "crpe32.dll" (ByVal printJob%) As Integer

' for PEGetGroupCondition, Condition% encodes both
' the condition and the type of the condition field
Global Const PE_GC_CONDITIONMASK = &HFF
Global Const PE_GC_TYPEMASK = &HF00

Global Const PE_GC_TYPEOTHER = &H0
Global Const PE_GC_TYPEDATE = &H200
Global Const PE_GC_TYPEBOOLEAN = &H400

Declare Function PEGetGroupCondition Lib "crpe32.dll" (ByVal printJob%, ByVal sectionCode%, ConditionFieldHandle As Long, ConditionFieldLength%, condition%, sortDirection%) As Integer

Global Const PE_FIELD_NAME_LEN = 512

Global Const PE_GO_TBN_ALL_GROUPS_UNSORTED = 0
Global Const PE_GO_TBN_ALL_GROUPS_SORTED = 1
Global Const PE_GO_TBN_TOP_N_GROUPS = 2
Global Const PE_GO_TBN_BOTTOM_N_GROUPS = 3

Type PEGroupOptions
    StructSize As Integer
    'when setting, pass a PE_GC_ constant, or PE_UNCHANGED for no change.
    'when getting, use PE_GC_TYPEMASK and PE_GC_CONDITIONMASK to
    'decode the condition.
    condition As Integer
    fieldName As String * PE_FIELD_NAME_LEN ' formula form, or empty for no change.
    sortDirection As Integer                ' a PE_SF_ const, or PE_UNCHANGED for no change.
    repeatGroupHeader As Integer            ' BOOL value, or PE_UNCHANGED for no change.
    keepGroupTogether As Integer            ' BOOL value, or PE_UNCHANGED for no change.
    topOrBottomNGroups As Integer           ' a PE_GO_TBN_ constant, or PE_UNCHANGED for no change.
    topOrBottomNSortFieldName As String * PE_FIELD_NAME_LEN ' formula form, or empty for no change.
    nTopOrBottomGroups As Integer           ' the number of groups to keep, 0 for all, or PE_UNCHANGED for no change.
    discardOtherGroups As Integer           ' BOOL value, or PE_UNCHANGED for no change.
End Type

Global Const PE_SIZEOF_GROUP_OPTIONS = 8 * PE_WORD_LEN + _
                                       2 * PE_FIELD_NAME_LEN

Declare Function PEGetGroupOptions Lib "crpe32.dll" (ByVal printJob%, ByVal groupN%, groupOptions As PEGroupOptions) As Integer
                                
Declare Function PESetGroupOptions Lib "crpe32.dll" (ByVal printJob%, ByVal groupN%, groupOptions As PEGroupOptions) As Integer

' Controlling formulas, selection formulas and group selection formulas
' ---------------------------------------------------------------------

Declare Function PEGetNFormulas Lib "crpe32.dll" (ByVal printJob%) As Integer

Declare Function PEGetNthFormula Lib "crpe32.dll" (ByVal printJob%, ByVal FormulaN%, NameHandle As Long, NameLength%, textHandle As Long, textLength%) As Integer

Declare Function PEGetFormula Lib "crpe32.dll" (ByVal printJob%, ByVal formulaName$, textHandle As Long, textLength%) As Integer

Declare Function PESetFormula Lib "crpe32.dll" (ByVal printJob%, ByVal formulaName$, ByVal formulaString$) As Integer

Declare Function PECheckFormula Lib "crpe32.dll" (ByVal printJob%, ByVal formulaName$) As Integer

Declare Function PEGetSelectionFormula Lib "crpe32.dll" (ByVal printJob%, textHandle As Long, textLength%) As Integer

Declare Function PESetSelectionFormula Lib "crpe32.dll" (ByVal printJob%, ByVal formulaString$) As Integer

Declare Function PECheckSelectionFormula Lib "crpe32.dll" (ByVal printJob%) As Integer

Declare Function PEGetGroupSelectionFormula Lib "crpe32.dll" (ByVal printJob%, textHandle As Long, textLength%) As Integer

Declare Function PESetGroupSelectionFormula Lib "crpe32.dll" (ByVal printJob%, ByVal formulaString$) As Integer

Declare Function PECheckGroupSelectionFormula Lib "crpe32.dll" (ByVal printJob%) As Integer

' Controlling Parameter Fields
' ----------------------------

Global Const PE_PF_REPORT_NAME_LEN = 128
Global Const PE_PF_NAME_LEN = 256
Global Const PE_PF_PROMPT_LEN = 256
Global Const PE_PF_VALUE_LEN = 256

Global Const PE_PF_NUMBER = 0
Global Const PE_PF_CURRENCY = 1
Global Const PE_PF_BOOLEAN = 2
Global Const PE_PF_DATE = 3
Global Const PE_PF_STRING = 4

Type PEParameterFieldInfo
    'Initialize to PE_SIZEOF_PARAMETER_FIELD_INFO.
    StructSize As Integer

    'PE_PF_ constant
    valueType As Integer

    'Indicate the default value is set in PEParameterFieldInfo.
    DefaultValueSet As Integer

    'Indicate the current value is set in PEParameterFieldInfo.
    CurrentValueSet As Integer

    'All strings are null-terminated.
    Name As String * PE_PF_NAME_LEN
    Prompt As String * PE_PF_PROMPT_LEN

    ' Could be Number, Date, DateTime, Time, Boolean, or String
    DefaultValue As String * PE_PF_VALUE_LEN
    CurrentValue As String * PE_PF_VALUE_LEN

    'name of report where the field belongs, only used in PEGetNthParameterField
    ReportName As String * PE_PF_REPORT_NAME_LEN

    'returns false (0) if parameter is linked, not in use, or has current value set
    needsCurrentValue As Integer
End Type
    
Global Const PE_SIZEOF_VARINFO_TYPE = 5 * PE_WORD_LEN + PE_PF_NAME_LEN + PE_PF_PROMPT_LEN + 2 * PE_PF_VALUE_LEN + PE_PF_REPORT_NAME_LEN
Global Const PE_SIZEOF_PARAMETER_FIELD_INFO = 5 * PE_WORD_LEN + PE_PF_NAME_LEN + PE_PF_PROMPT_LEN + 2 * PE_PF_VALUE_LEN + PE_PF_REPORT_NAME_LEN

Declare Function PEGetNParameterFields Lib "crpe32.dll" (ByVal printJob%) As Integer

Declare Function PEGetNthParameterField Lib "crpe32.dll" (ByVal printJob%, ByVal varN%, varInfo As PEParameterFieldInfo) As Integer

Declare Function PESetNthParameterField Lib "crpe32.dll" (ByVal printJob%, ByVal varN%, varInfo As PEParameterFieldInfo) As Integer

'*** Converting parameterInfo default value or current value into value info ****
Global Const PE_VI_STRING_LEN = 256

' define value type
Global Const PE_VI_NUMBER = 0
Global Const PE_VI_CURRENCY = 1
Global Const PE_VI_BOOLEAN = 2
Global Const PE_VI_DATE = 3
Global Const PE_VI_STRING = 4
Global Const PE_VI_DATETIME = 5
Global Const PE_VI_TIME = 6
Global Const PE_VI_INTEGER = 7
Global Const PE_VI_COLOR = 8
Global Const PE_VI_CHAR = 9
Global Const PE_VI_LONG = 10
Global Const PE_VI_NOVALUE = 100

Type PEValueInfo
    StructSize As Integer
    valueType As Integer  'a PE_VI_ constant
    viNumber As Double
    viCurrency As Double
    viBoolean As Long
    viString As String * PE_VI_STRING_LEN
    viDate(0 To 2) As Integer ' year, month, day
    viDateTime(0 To 5) As Integer ' year, month, day, hour, minute, second
    viTime(0 To 2) As Integer  ' hour, minute, second
    viColor As Long
    viInteger As Integer
    viC As Byte
    ignored As Byte 'for 4 byte alignment. ignored.
    viLong As Long
End Type

Global Const PE_SIZEOF_VALUE_INFO = 2 * PE_BYTE_LEN + _
                                   15 * PE_WORD_LEN + _
                                    3 * PE_LONG_LEN + _
                                    2 * PE_DOUBLE_LEN + _
                                    1 * PE_VI_STRING_LEN
                                     
Declare Function PEConvertPFInfoToVInfo Lib "crpe32.dll" (ByVal value As Any, ByVal valueType%, valueInfo As PEValueInfo) As Integer
Declare Function PEConvertVInfoToPFInfo Lib "crpe32.dll" (valueInfo As PEValueInfo, valueType%, ByVal value As Any) As Integer


' Controlling sort order and group sort order
' -------------------------------------------

Global Const PE_SF_MAXNAMELEN = 50  ' maximum length of a sort field name

Global Const PE_SF_DESC = 0         ' values for the Direction parameter
Global Const PE_SF_ASC = 1

Declare Function PEGetNSortFields Lib "crpe32.dll" (ByVal printJob%) As Integer

Declare Function PEGetNthSortField Lib "crpe32.dll" (ByVal printJob%, ByVal SortNumber%, NameHandle As Long, NameLength%, Direction%) As Integer

Declare Function PESetNthSortField Lib "crpe32.dll" (ByVal printJob%, ByVal SortNumber%, ByVal SortFieldName$, ByVal Direction%) As Integer

Declare Function PEDeleteNthSortField Lib "crpe32.dll" (ByVal printJob%, ByVal SortFieldN%) As Integer

Declare Function PEGetNGroupSortFields Lib "crpe32.dll" (ByVal printJob%) As Integer

Declare Function PEGetNthGroupSortField Lib "crpe32.dll" (ByVal printJob%, ByVal SortFieldN%, NameHandle As Long, NameLength%, Direction%) As Integer

Declare Function PESetNthGroupSortField Lib "crpe32.dll" (ByVal printJob%, ByVal SortFieldN%, ByVal SortGroupName$, ByVal Direction%) As Integer

Declare Function PEDeleteNthGroupSortField Lib "crpe32.dll" (ByVal printJob%, ByVal SortFieldN%) As Integer


' Controlling databases
' ---------------------
'
' The following functions allow retrieving and updating database info
' in an opened report, so that a report can be printed using different
' session, server, database, user and/or table location settings.  Any
' changes made to the report via these functions are not permanent, and
' only last as long as the report is open.
'
' The following database functions (except for PELogOnServer and
' PELogOffServer) must be called after PEOpenPrintJob and before
' PEStartPrintJob.

' The function PEGetNTables is called to fetch the number of tables in
' the report.  This includes all PC databases (e.g. Paradox, xBase)
' as well as SQL databases (e.g. SQL Server, Oracle, Netware).

Declare Function PEGetNTables Lib "crpe32.dll" (ByVal printJob%) As Integer

' The function PEGetNthTableType allows the application to determine the
' type of each table.  The application can test DBType (equal to
' PE_DT_STANDARD or PE_DT_SQL), or test the database DLL name used to
' create the report.  DLL names have the following naming convention:
'     - PDB*.DLL for standard (non-SQL) databases,
'     - PDS*.DLL for SQL databases.
'
' In the case of ODBC (PDSODBC.DLL) the DescriptiveName includes the
' ODBC data source name.

Global Const PE_DLL_NAME_LEN = 64
Global Const PE_FULL_NAME_LEN = 256
Global Const PE_SIZEOF_TABLE_TYPE = 324 ' # bytes in PETableType

Global Const PE_DT_STANDARD = 1  ' values for DBType
Global Const PE_DT_SQL = 2

Type PETableType
    StructSize As Integer   ' initialize to # bytes in PETableType

    DLLName As String * PE_DLL_NAME_LEN
    DescriptiveName  As String * PE_FULL_NAME_LEN

    DBType As Integer
End Type

Declare Function PEGetNthTableType Lib "crpe32.dll" (ByVal printJob%, ByVal TableN%, TableType As PETableType) As Integer

' The functions PEGetNthTableSessionInfo and PESetNthTableSessionInfo
' are only used when connecting to MS Access databases (which require a
' session to be opened first)

Global Const PE_SESS_USERID_LEN = 128
Global Const PE_SESS_PASSWORD_LEN = 128
Global Const PE_SIZEOF_SESSION_INFO = 262  ' # bytes in PESessionInfo

Type PESessionInfo
    'initialize to # bytes in PESessionInfo
    StructSize As Integer

    ' Password is undefined when getting information from report.
    UserID As String * PE_SESS_USERID_LEN
    Password As String * PE_SESS_PASSWORD_LEN

    ' SessionHandle is undefined when getting information from report.
    ' When setting information, if it is = 0 the UserID and Password
    ' settings are used, otherwise the SessionHandle is used.
    SessionHandle As Long
End Type

Declare Function PEGetNthTableSessionInfo Lib "crpe32.dll" (ByVal printJob%, ByVal TableN%, SessionInfo As PESessionInfo) As Integer

Declare Function PESetNthTableSessionInfo Lib "crpe32.dll" (ByVal printJob%, ByVal TableN%, SessionInfo As PESessionInfo, ByVal PropagateAcrossTables%) As Integer

' Logging on is performed when printing the report, but the correct
' log on information must first be set using PESetNthTableLogOnInfo.
' Only the password is required, but the server, database, and
' user names may optionally be overriden as well.
'
' If the parameter propagateAcrossTables is TRUE, the new log on info
' is also applied to any other tables in this report that had the
' same original server and database names as this table.  If FALSE
' only this table is updated.
'
' Logging off is performed automatically when the print job is closed.

Global Const PE_SERVERNAME_LEN = 128
Global Const PE_DATABASENAME_LEN = 128
Global Const PE_USERID_LEN = 128
Global Const PE_PASSWORD_LEN = 128
Global Const PE_SIZEOF_LOGON_INFO = 514  ' # bytes in PELogOnInfo

Type PELogOnInfo
    ' initialize to # bytes in PELogOnInfo
    StructSize As Integer

    ' For any of the following values an empty string ("") means to use
    ' the value already set in the report.  To override a value in the
    ' report use a non-empty string (e.g. "Server A").
    '
    ' For Netware SQL, pass the dictionary path name in ServerName and
    ' data path name in DatabaseName.

    ServerName As String * PE_SERVERNAME_LEN
    DatabaseName  As String * PE_DATABASENAME_LEN
    UserID As String * PE_USERID_LEN

    ' Password is undefined when getting information from report.

    Password  As String * PE_PASSWORD_LEN
End Type

Declare Function PEGetNthTableLogOnInfo Lib "crpe32.dll" (ByVal printJob%, ByVal TableN%, LogOnInfo As PELogOnInfo) As Integer
Declare Function PESetNthTableLogOnInfo Lib "crpe32.dll" (ByVal printJob%, ByVal TableN%, LogOnInfo As PELogOnInfo, ByVal Propagate%) As Integer

' A table's location is fetched and set using PEGetNthTableLocation and
' PESetNthTableLocation.  This name is database-dependent, and must be
' formatted correctly for the expected database.  For example:
'     - Paradox: "c:\crw\ORDERS.DB"
'     - SQL Server: "publications.dbo.authors"

Global Const PE_TABLE_LOCATION_LEN = 256
Global Const PE_SIZEOF_TABLE_LOCATION = 258  ' # bytes in PETableLocation

Type PETableLocation
    ' initialize to # bytes in PETableLocation
    StructSize As Integer
    Location  As String * PE_TABLE_LOCATION_LEN
End Type

Declare Function PEGetNthTableLocation Lib "crpe32.dll" (ByVal printJob%, ByVal TableN%, Location As PETableLocation) As Integer
Declare Function PESetNthTableLocation Lib "crpe32.dll" (ByVal printJob%, ByVal TableN%, Location As PETableLocation) As Integer

' If report based on a SQL Stored Procedure, use PEGetNParams to fetch the
' number of parameters, and PEGetNthParam and PESetNthParam to fetch and
' set individual parameters.  All parameter values are encoded as strings.

' If a parameter value is NULL, using PEGetNthParams will return a NULL
' for the textHandle and a zero value for the textLength.
'
' If you wish to SET a parameter to NULL then set the ParamValue to "CRWNULL"
' when using the PESetNthParam Api call.
' eg. PESetNthParam(myJobId, myParamNum, "CRWNULL")

Global Const PE_PARAMETER_NAME_LEN = 128
Global Const PE_PT_LONGVARCHAR = -1
Global Const PE_PT_BINARY = -2
Global Const PE_PT_VARBINARY = -3
Global Const PE_PT_LONGVARBINARY = -4
Global Const PE_PT_BIGINT = -5
Global Const PE_PT_TINYINT = -6
Global Const PE_PT_BIT = -7
Global Const PE_PT_CHAR = 1
Global Const PE_PT_NUMERIC = 2
Global Const PE_PT_DECIMAL = 3
Global Const PE_PT_INTEGER = 4
Global Const PE_PT_SMALLINT = 5
Global Const PE_PT_FLOAT = 6
Global Const PE_PT_REAL = 7
Global Const PE_PT_DOUBLE = 8
Global Const PE_PT_DATE = 9
Global Const PE_PT_TIME = 10
Global Const PE_PT_TIMESTAMP = 11
Global Const PE_PT_VARCHAR = 12

Type PEParameterInfo
     'Initialize to PE_SIZEOF_PARAMETER_INFO.
     StructSize As Integer

     Type As Integer
     
     'String is null-terminated.
     Name As String * PE_PARAMETER_NAME_LEN
End Type

Global Const PE_SIZEOF_PARAMETER_INFO = 2 * PE_WORD_LEN + PE_PARAMETER_NAME_LEN

Declare Function PEGetNParams Lib "crpe32.dll" (ByVal printJob%) As Integer
Declare Function PEGetNthParam Lib "crpe32.dll" (ByVal printJob%, ByVal paramN%, textHandle As Long, textLength%) As Integer
Declare Function PEGetNthParamInfo Lib "crpe32.dll" (ByVal printJob%, ByVal paramN%, paramInfo As PEParameterInfo) As Integer
Declare Function PESetNthParam Lib "crpe32.dll" (ByVal printJob%, ByVal paramN%, ByVal ParamValue$) As Integer

' The function PETestNthTableConnectivity tests whether a database
' table's settings are valid and ready to be reported on.  It returns
' true if the database session, log on, and location info is all
' correct.
'
' This is useful, for example, in prompting the user and testing a
' server password before printing begins.
'
' This function may require a significant amount of time to complete,
' since it will first open a user session (if required), then log onto
' the database server (if required), and then open the appropriate
' database table (to test that it exists).  It does not read any data,
' and closes the table immediately once successful.  Logging off is
' performed when the print job is closed.
'
' If it fails in any of these steps, the error code set indicates
' which database info needs to be updated using functions above:
'    - If it is unable to begin a session, PE_ERR_DATABASESESSION is set,
'      and the application should update with PESetNthTableSessionInfo.
'    - If it is unable to log onto a server, PE_ERR_DATABASELOGON is set,
'      and the application should update with PESetNthTableLogOnInfo.
'    - If it is unable open the table, PE_ERR_DATABASELOCATION is set,
'      and the application should update with PESetNthTableLocation.

Declare Function PETestNthTableConnectivity Lib "crpe32.dll" (ByVal printJob%, ByVal TableN%) As Integer

' PELogOnServer and PELogOffServer can be called at any time to log on
' and off of a database server.  These functions are not required if
' function PESetNthTableLogOnInfo above was already used to set the
' password for a table.
'
' These functions require a database DLL name, which can be retrieved
' using PEGetNthTableType above.
'
' This function can also be used for non-SQL tables, such as password-
' protected Paradox tables.  Call this function to set the password
' for the Paradox DLL before beginning printing.
'
' Note: When printing using PEStartPrintJob the ServerName passed in
' PELogOnServer must agree exactly with the server name stored in the
' report.  If this is not true use PESetNthTableLogOnInfo to perform
' logging on instead.

Declare Function PELogOnServer Lib "crpe32.dll" (ByVal DLLName$, LogOnInfo As PELogOnInfo) As Integer
Declare Function PELogOffServer Lib "crpe32.dll" (ByVal DLLName$, LogOnInfo As PELogOnInfo) As Integer
Declare Function PELogOnSQLServerWithPrivateInfo Lib "crpe32.dll" (ByVal DLLName$, ByVal PrivateInfo As Long) As Integer


' Overriding SQL query in report
' ------------------------------
'
' PEGetSQLQuery returns the same query as appears in the Show SQL Query
' dialog in CRW, in syntax specific to the database driver you are using.
'
' PESetSQLQuery is mostly useful for reports with SQL queries that
' were explicitly edited in the Show SQL Query dialog in CRW, i.e. those
' reports that needed database-specific selection criteria or joins.
' (Otherwise it is usually best to continue using function calls such as
' PESetSelectionFormula and let CRW build the SQL query automatically.)
'
' PESetSQLQuery has the same restrictions as editing in the Show SQL
' Query dialog; in particular that changes are accepted in the FROM and
' WHERE clauses but ignored in the SELECT list of fields.

Declare Function PEGetSQLQuery Lib "crpe32.dll" (ByVal printJob%, textHandle As Long, textLength%) As Integer

Declare Function PESetSQLQuery Lib "crpe32.dll" (ByVal printJob%, ByVal QueryString$) As Integer


' Saved data
' ----------
'
' Use PEHasSavedData to find out if a report currently has saved data
' associated with it.  This may or may not be TRUE when a print job is
' first opened from a report file.  Since data is saved during a print,
' this will always be TRUE immediately after a report is printed.
'
' Use PEDiscardSavedData to release the saved data associated with a
' report.  The next time the report is printed, it will get current data
' from the database.
'
' The default behavior is for a report to use its saved data, rather than
' refresh its data from the database when printing a report.

Declare Function PEHasSavedData Lib "crpe32.dll" (ByVal printJob%, HasSavedData As Long) As Integer

Declare Function PEDiscardSavedData Lib "crpe32.dll" (ByVal printJob%) As Integer


' Report title
' ------------

Declare Function PEGetReportTitle Lib "crpe32.dll" (ByVal printJob%, TitleHandle As Long, TitleLength%) As Integer
Declare Function PESetReportTitle Lib "crpe32.dll" (ByVal printJob%, ByVal title$) As Integer


' Controlling print to window
' ---------------------------

Declare Function PEOutputToWindow Lib "crpe32.dll" (ByVal printJob%, ByVal title$, ByVal Left As Long, ByVal Top As Long, ByVal Width As Long, ByVal Height As Long, ByVal style As Long, ByVal PWindow As Long) As Integer

Type PEWindowOptions
    StructSize As Integer            'initialize to PE_SIZEOF_WINDOW_OPTIONS

    hasGroupTree As Integer          '0 or 1, except use PE_UNCHANGED for no change
    canDrillDown As Integer          '0 or 1, except use PE_UNCHANGED for no change
    hasNavigationControls As Integer '0 or 1, except use PE_UNCHANGED for no change
    hasCancelButton As Integer       '0 or 1, except use PE_UNCHANGED for no change
    hasPrintButton As Integer        '0 or 1, except use PE_UNCHANGED for no change
    hasExportButton As Integer       '0 or 1, except use PE_UNCHANGED for no change
    hasZoomControl As Integer        '0 or 1, except use PE_UNCHANGED for no change
    hasCloseButton As Integer        '0 or 1, except use PE_UNCHANGED for no change
    hasProgressControls As Integer   '0 or 1, except use PE_UNCHANGED for no change
    hasSearchButton As Integer       '0 or 1, except use PE_UNCHANGED for no change
    hasPrintSetupButton As Integer   '0 or 1, except use PE_UNCHANGED for no change
    hasRefreshButton As Integer      '0 or 1, except use PE_UNCHANGED for no change
End Type

Global Const PE_SIZEOF_WINDOW_OPTIONS = 13 * PE_WORD_LEN

Declare Function PEGetWindowOptions Lib "crpe32.dll" (ByVal printJob%, Options As PEWindowOptions) As Integer

Declare Function PESetWindowOptions Lib "crpe32.dll" (ByVal printJob%, Options As PEWindowOptions) As Integer

Declare Function PEGetWindowHandle Lib "crpe32.dll" (ByVal printJob%) As Integer

Declare Sub PECloseWindow Lib "crpe32.dll" (ByVal printJob%)


' Controlling printed pages
' -------------------------

Declare Function PEShowNextPage Lib "crpe32.dll" (ByVal printJob%) As Integer
Declare Function PEShowFirstPage Lib "crpe32.dll" (ByVal printJob%) As Integer
Declare Function PEShowPreviousPage Lib "crpe32.dll" (ByVal printJob%) As Integer
Declare Function PEShowLastPage Lib "crpe32.dll" (ByVal printJob%) As Integer
Declare Function PEGetNPages Lib "crpe32.dll" (ByVal printJob%) As Integer
Declare Function PEShowNthPage Lib "crpe32.dll" (ByVal printJob%, ByVal pageN%) As Integer

Global Const PE_ZOOM_FULL_SIZE = 0
Global Const PE_ZOOM_SIZE_FIT_ONE_SIDE = 1
Global Const PE_ZOOM_SIZE_FIT_BOTH_SIDES = 2

Declare Function PEZoomPreviewWindow Lib "crpe32.dll" (ByVal printJob%, ByVal ZoomLevel%) As Integer
' ZoomLevel is a percent from 25 to 400 or a PE_ZOOM_ constant

' Controlling print window when print control buttons hidden
' ----------------------------------------------------------

Declare Function PEShowPrintControls Lib "crpe32.dll" (ByVal printJob%, ByVal ShowPrintControls%) As Integer

Declare Function PEPrintControlsShowing Lib "crpe32.dll" (ByVal printJob%, ControlsShowing As Long) As Integer

Declare Function PEPrintWindow Lib "crpe32.dll" (ByVal printJob%, ByVal WaitNoWait%) As Integer

Declare Function PEExportPrintWindow Lib "crpe32.dll" (ByVal printJob%, ByVal ToMail%, ByVal WaitUntilDone%) As Integer

Declare Function PENextPrintWindowMagnification Lib "crpe32.dll" (ByVal printJob%) As Integer


' Changing printer selection
' --------------------------

Declare Function PESelectPrinter Lib "crpe32.dll" (ByVal printJob%, ByVal PrinterDriver$, ByVal PrinterName$, ByVal PortName$, DevMode As Any) As Integer

Declare Function PEGetSelectedPrinter Lib "crpe32.dll" (ByVal printJob%, driverHandle As Long, driverLength%, printerHandle As Long, printerLength%, porthandle As Long, portlength%, DevMode As Any) As Integer


' Controlling print to printer
' ----------------------------

Declare Function PEOutputToPrinter Lib "crpe32.dll" (ByVal printJob%, ByVal nCopies%) As Integer

Declare Function PESetNDetailCopies Lib "crpe32.dll" (ByVal printJob%, ByVal nDetailCopies%) As Integer

Declare Function PEGetNDetailCopies Lib "crpe32.dll" (ByVal printJob%, nDetailCopies%) As Integer

' Extension to PESetPrintOptions function: If the 2nd parameter
' (pointer to PEPrintOptions) is set to 0 (null) the function prompts
' the user for these options.
'
' With this change, you can get the behaviour of the print-to-printer
' button in the print window by calling PESetPrintOptions with a
' null pointer and then calling PEPrintWindow.

Global Const PE_MAXPAGEN = 65535

Global Const PE_UNCOLLATED = 0
Global Const PE_COLLATED = 1
Global Const PE_DEFAULTCOLLATION = 2

Type PEPrintOptions
    StructSize As Integer   ' initialize to # bytes in PEPrintOptions

    ' page and copy numbers are 1-origin
    ' use 0 to preserve the existing settings
    StartPageN As Integer
    stopPageN As Integer

    nReportCopies As Integer
    collation As Integer
End Type

Global Const PE_SIZEOF_PRINT_OPTIONS = 5 * PE_WORD_LEN

Declare Function PESetPrintOptions Lib "crpe32.dll" (ByVal printJob%, Options As PEPrintOptions) As Integer

Declare Function PEGetPrintOptions Lib "crpe32.dll" (ByVal printJob%, Options As PEPrintOptions) As Integer


' Controlling print to file and export
' ------------------------------------

Global Const PE_FT_RECORD = 0
Global Const PE_FT_TABSEPARATED = 1
Global Const PE_FT_TEXT = 2
Global Const PE_FT_DIF = 3
Global Const PE_FT_CSV = 4
Global Const PE_FT_CHARSEPARATED = 5
Global Const PE_FT_TABFORMATTED = 6

' Use for all types except PE_FT_CHARSEPARATED
Type PEPrintFileOptions
    StructSize As Integer   ' initialize to # of bytes in PEPrintFileOptions

    UseReportNumberFmt As Integer
    reserved1 As Integer     ' reserved - do not set
    UseReportDateFormat As Integer
    reserved2 As Integer     ' reserved - do not set
End Type

Global Const PE_SIZEOF_PRINT_FILE_OPTIONS = 3 * PE_WORD_LEN

Global Const PE_FIELDDELIMLEN = 17

' Use for PE_FT_CHARSEPARATED
Type PECharSepFileOptions
    StructSize As Integer   ' initialize to PE_SIZEOF_CHAR_SEP_FILE_OPTIONS

    UseReportNumberFmt As Integer
    reserved1 As Integer     ' reserved - do not set
    UseReportDateFormat As Integer
    reserved2 As Integer     ' reserved - do not set

    StringDelimiter As String * 1
    FieldDelimiter As String * PE_FIELDDELIMLEN
End Type

Global Const PE_SIZEOF_CHAR_SEP_FILE_OPTIONS = 3 * PE_WORD_LEN + 1 * 1 + PE_FIELDDELIMLEN

Declare Function PEOutputToFile Lib "crpe32.dll" (ByVal printJob%, ByVal OutputFilePath$, ByVal FileType%, Options As Any) As Integer

Type PEExportOptions
    StructSize As Integer   'initialize to # bytes in PEExportOptions

    FormatDLLName As String * PE_DLL_NAME_LEN
    FormatType1 As Integer
    FormatType2 As Integer
    FormatOptions1 As Integer
    FormatOptions2 As Integer

    DestinationDLLName As String * PE_DLL_NAME_LEN
    DestinationType1 As Integer
    DestinationType2 As Integer
    DestinationOptions1 As Integer
    DestinationOptions2 As Integer

    ' following are set by PEGetExportOptions,
    ' and ignored by PEExportTo.
    NFormatOptionsBytes As Integer
    NDestinationOptionsBytes As Integer
End Type

Global Const PE_SIZEOF_EXPORT_OPTIONS = 11 * PE_WORD_LEN + 2 * PE_DLL_NAME_LEN

Declare Function PEGetExportOptions Lib "crpe32.dll" (ByVal printJob%, ExportOptions As PEExportOptions) As Integer

Declare Function PEExportTo Lib "crpe32.dll" (ByVal printJob%, ExportOptions As PEExportOptions) As Integer


' Setting page margins
' --------------------

Global Const PE_SM_DEFAULT = &H8000

Declare Function PESetMargins Lib "crpe32.dll" (ByVal printJob%, ByVal LeftMargin%, ByVal RightMargin%, ByVal TopMargin%, ByVal BottomMargin%) As Integer

Declare Function PEGetMargins Lib "crpe32.dll" (ByVal printJob%, LeftMargin%, RightMargin%, TopMargin%, BottomMargin%) As Integer


'Report Summary Info
'-------------------

Global Const PE_SI_APPLICATION_NAME_LEN = 128
Global Const PE_SI_TITLE_LEN = 128
Global Const PE_SI_SUBJECT_LEN = 128
Global Const PE_SI_AUTHOR_LEN = 128
Global Const PE_SI_KEYWORDS_LEN = 128
Global Const PE_SI_COMMENTS_LEN = 512
Global Const PE_SI_REPORT_TEMPLATE_LEN = 128

Type PEReportSummaryInfo
    StructSize As Integer
    applicationName As String * PE_SI_APPLICATION_NAME_LEN ' read only.
    title As String * PE_SI_TITLE_LEN
    subject As String * PE_SI_SUBJECT_LEN
    author As String * PE_SI_AUTHOR_LEN
    keywords As String * PE_SI_KEYWORDS_LEN
    comments As String * PE_SI_COMMENTS_LEN
    reportTemplate As String * PE_SI_REPORT_TEMPLATE_LEN
End Type
Const PE_SIZEOF_REPORT_SUMMARY_INFO = 1 * PE_WORD_LEN + _
                                      PE_SI_APPLICATION_NAME_LEN + _
                                      PE_SI_TITLE_LEN + _
                                      PE_SI_SUBJECT_LEN + _
                                      PE_SI_AUTHOR_LEN + _
                                      PE_SI_KEYWORDS_LEN + _
                                      PE_SI_COMMENTS_LEN + _
                                      PE_SI_REPORT_TEMPLATE_LEN
                                
Declare Function PEGetReportSummaryInfo Lib "crpe32.dll" (ByVal printJob%, summaryInfo As PEReportSummaryInfo) As Integer

Declare Function PESetReportSummaryInfo Lib "crpe32.dll" (ByVal printJob%, summaryInfo As PEReportSummaryInfo) As Integer

' Setting section height and format
' ---------------------------------

Declare Function PEGetNSections Lib "crpe32.dll" (ByVal printJob%) As Integer

Declare Function PEGetSectionCode Lib "crpe32.dll" (ByVal printJob%, ByVal sectionN%) As Integer

' MinimumHeight is in twips - 1440 twips to the inch
Declare Function PESetMinimumSectionHeight Lib "crpe32.dll" (ByVal printJob%, ByVal sectionCode%, ByVal MinimumHeight%) As Integer
Declare Function PEGetMinimumSectionHeight Lib "crpe32.dll" (ByVal printJob%, ByVal sectionCode%, MinimumHeight%) As Integer

Type PESectionOptions
    StructSize As Integer   ' initialize to PE_SIZEOF_SECTION_OPTIONS

    ' use 0 to turn off, 1 to turn on and PE_UNCHANGED to preserve each attribute
    Visible As Integer
    NewPageBefore As Integer
    NewPageAfter As Integer
    KeepTogether As Integer
    SuppressBlankSection As Integer
    ResetPageNAfter As Integer
    PrintAtBottomOfPage As Integer
    backgroundColor As Long   ' Use PE_UNCHANGED_COLOR to preserve the
                              ' existing color.
    underlaySection As Integer
    showArea As Integer
    freeFormPlacement As Integer
End Type

Global Const PE_SIZEOF_SECTION_OPTIONS = 11 * PE_WORD_LEN + 1 * 4

'Format formula name
'Old naming convention
Global Const SECTION_VISIBILITY = 58
Global Const NEW_PAGE_BEFORE = 60
Global Const NEW_PAGE_AFTER = 61
Global Const KEEP_SECTION_TOGETHER = 62
Global Const SUPPRESS_BLANK_SECTION = 63
Global Const RESET_PAGE_N_AFTER = 64
Global Const PRINT_AT_BOTTOM_OF_PAGE = 65
Global Const UNDERLAY_SECTION = 66
Global Const SECTION_BACK_COLOUR = 67

'New naming convention
Global Const PE_FFN_AREASECTION_VISIBILITY = 58
Global Const PE_FFN_SECTION_VISIBILITY = 58
Global Const PE_FFN_SHOW_AREA = 59
Global Const PE_FFN_NEW_PAGE_BEFORE = 60
Global Const PE_FFN_NEW_PAGE_AFTER = 61
Global Const PE_FFN_KEEP_SECTION_TOGETHER = 62
Global Const PE_FFN_SUPPRESS_BLANK_SECTION = 63
Global Const PE_FFN_RESET_PAGE_N_AFTER = 64
Global Const PE_FFN_PRINT_AT_BOTTOM_OF_PAGE = 65
Global Const PE_FFN_UNDERLAY_SECTION = 66
Global Const PE_FFN_SECTION_BACK_COLOUR = 67

Declare Function PEGetSectionFormatFormula Lib "crpe32.dll" (ByVal printJob%, ByVal sectionCode%, ByVal formulaName%, textHandle As Long, textLength%) As Integer

Declare Function PESetSectionFormatFormula Lib "crpe32.dll" (ByVal printJob%, ByVal sectionCode%, ByVal formulaName%, ByVal formulaString$) As Integer

Declare Function PEGetSectionFormat Lib "crpe32.dll" (ByVal printJob%, ByVal sectionCode%, Options As PESectionOptions) As Integer

Declare Function PESetSectionFormat Lib "crpe32.dll" (ByVal printJob%, ByVal sectionCode%, Options As PESectionOptions) As Integer

' Setting area format
' -------------------
                                                                                                               
Declare Function PEGetAreaFormatFormula Lib "crpe32.dll" (ByVal printJob%, ByVal areaCode%, ByVal formulaName%, textHandle As Long, textLength%) As Integer

Declare Function PESetAreaFormatFormula Lib "crpe32.dll" (ByVal printJob%, ByVal areaCode%, ByVal formulaName%, ByVal formulaString$) As Integer

Declare Function PEGetAreaFormat Lib "crpe32.dll" (ByVal printJob%, ByVal areaCode%, Options As PESectionOptions) As Integer

Declare Function PESetAreaFormat Lib "crpe32.dll" (ByVal printJob%, ByVal areaCode%, Options As PESectionOptions) As Integer


' Setting font info
' -----------------

' values for ScopeCode - may be ORed together
Global Const PE_FIELDS = 1
Global Const PE_TEXT = 2

Global Const PE_UNCHANGED = -1

' to preserve the existing setting, use the following
'   for FontFamily%    use  FF_DONTCARE
'   for FontPitch%     use  DEFAULT_PITCH
'   for CharSet%       use  DEFAULT_CHARSET
'   for PointSize%     use  0
'   for isItalic%      use  PE_UNCHANGED
'   for isUnderlined%  use  PE_UNCHANGED
'   for isStruckOut%   use  PE_UNCHANGED
'   for Weight%        use  0
Declare Function PESetFont Lib "crpe32.dll" (ByVal printJob%, ByVal sectionCode%, ByVal ScopeCode%, ByVal FaceName$, ByVal FontFamily%, ByVal FontPitch%, ByVal CharSet%, ByVal PointSize%, ByVal isItalic%, ByVal isUnderlined%, ByVal isStruckOut%, ByVal Weight%) As Integer


' Setting Graph/Chart info
' ------------------------
'
' Two parameters are passed to uniquely identify the graph:
'      - section code
'      - graph number in that section
'
' The section code includes whether it is a header or footer, and the
' graph number starts at 0, 1...  The graph number identifies the graph
' by its position in the section
'      - looking top down first,
'      - then left to right if they have the same top.

' Graph Types

Global Const PE_SIDE_BY_SIDE_BAR_GRAPH = 0
Global Const PE_STACKED_BAR_GRAPH = 2
Global Const PE_PERCENT_BAR_GRAPH = 3
Global Const PE_FAKED_3D_SIDE_BY_SIDE_BAR_GRAPH = 4
Global Const PE_FAKED_3D_STACKED_BAR_GRAPH = 5
Global Const PE_FAKED_3D_PERCENT_BAR_GRAPH = 6
Global Const PE_PIE_GRAPH = 40
Global Const PE_MULTIPLE_PIE_GRAPH = 42
Global Const PE_PROPORTIONAL_MULTI_PIE_GRAPH = 43
Global Const PE_LINE_GRAPH = 80
Global Const PE_AREA_GRAPH = 120
Global Const PE_THREED_BAR_GRAPH = 160
Global Const PE_USER_DEFINED_GRAPH = 500
Global Const PE_UNKNOWN_TYPE_GRAPH = 1000

' Graph Directions.
Global Const PE_GRAPH_ROWS_ONLY = 0
Global Const PE_GRAPH_COLS_ONLY = 1
Global Const PE_GRAPH_MIXED_ROW_COL = 2
Global Const PE_GRAPH_MIXED_COL_ROW = 3
Global Const PE_GRAPH_UNKNOWN_DIRECTION = 20

' Graph constant for rowGroupN, colGroupN, summarizedFieldN in PEGraphDataInfo
Global Const PE_GRAPH_DATA_NULL_SELECTION = -1

' Graph text max length
Global Const PE_GRAPH_TEXT_LEN = 128

Type PEGraphDataInfo
    StructSize        As Integer  ' initialize to # bytes in PEGraphDataInfo
    RowGroupN         As Integer  ' group number in report.
    ColGroupN         As Integer  ' group number in report.
    SummarizedFieldN  As Integer  ' summarized field number for the group
                                  ' where the graph stays.
    GraphDirection    As Integer  ' For normal group/total report, the direction,
                                  ' is always GRAPH_MIXED_ROW_COL.  For CrossTab
                                  ' report all four options will change the
                                  ' graph data.
End Type

Global Const PE_SIZEOF_GRAPH_DATA_INFO = 5 * PE_WORD_LEN

Type PEGraphTextInfo
    StructSize        As Integer  ' initialize to # bytes in PEGraphTextInfo
    GraphTitle        As String * PE_GRAPH_TEXT_LEN
    GraphSubTitle     As String * PE_GRAPH_TEXT_LEN
    GraphFootNote     As String * PE_GRAPH_TEXT_LEN
    GraphGroupsTitle  As String * PE_GRAPH_TEXT_LEN
    GraphSeriesTitle  As String * PE_GRAPH_TEXT_LEN
    GraphXAxisTitle   As String * PE_GRAPH_TEXT_LEN
    GraphYAxisTitle   As String * PE_GRAPH_TEXT_LEN
    GraphZAxisTitle   As String * PE_GRAPH_TEXT_LEN
End Type

Global Const PE_SIZEOF_GRAPH_TEXT_INFO = PE_WORD_LEN * 8 * PE_GRAPH_TEXT_LEN

Type PEGraphOptions
    StructSize     As Integer  ' initialize to # bytes in PEGraphOptions
    GraphMaxValue  As Double
    GraphMinValue  As Double
    ShowDataValue  As Long  ' Show data values on risers.
    ShowGridLine   As Long
    VerticalBars   As Long
    ShowLegend     As Long
    FontFaceName   As String * PE_GRAPH_TEXT_LEN
End Type

Global Const PE_SIZEOF_GRAPH_OPTIONS = PE_WORD_LEN + 2 * 8 + 4 * 4 + PE_GRAPH_TEXT_LEN

Declare Function PEGetGraphType Lib "crpe32.dll" (ByVal printJob%, ByVal sectionCode%, ByVal GraphN%, GraphType%) As Integer
Declare Function PEGetGraphData Lib "crpe32.dll" (ByVal printJob%, ByVal sectionCode%, ByVal GraphN%, GraphDataInfo As PEGraphDataInfo) As Integer
Declare Function PEGetGraphText Lib "crpe32.dll" (ByVal printJob%, ByVal sectionCode%, ByVal GraphN%, GraphTextInfo As PEGraphTextInfo) As Integer
Declare Function PEGetGraphOptions Lib "crpe32.dll" (ByVal printJob%, ByVal sectionCode%, ByVal GraphN%, GraphOptions As PEGraphOptions) As Integer

Declare Function PESetGraphType Lib "crpe32.dll" (ByVal printJob%, ByVal sectionCode%, ByVal GraphN%, GraphType%) As Integer
Declare Function PESetGraphData Lib "crpe32.dll" (ByVal printJob%, ByVal sectionCode%, ByVal GraphN%, GraphDataInfo As PEGraphDataInfo) As Integer
Declare Function PESetGraphText Lib "crpe32.dll" (ByVal printJob%, ByVal sectionCode%, ByVal GraphN%, GraphTextInfo As PEGraphTextInfo) As Integer
Declare Function PESetGraphOptions Lib "crpe32.dll" (ByVal printJob%, ByVal sectionCode%, ByVal GraphN%, GraphOptions As PEGraphOptions) As Integer

' Subreports
' ----------
Declare Function PEGetNSubreportsInSection Lib "crpe32.dll" (ByVal printJob%, ByVal sectionCode%) As Integer

Declare Function PEGetNthSubreportInSection Lib "crpe32.dll" (ByVal printJob%, ByVal sectionCode%, ByVal subreportN%) As Long

Global Const PE_SUBREPORT_NAME_LEN = 128

Type PESubreportInfo
    StructSize As Integer            ' Initialize to PE_SIZEOF_SUBREPORT_INFO.
    
    'Strings are null-terminated.
    Name As String * PE_SUBREPORT_NAME_LEN
End Type

Global Const PE_SIZEOF_SUBREPORT_INFO = PE_WORD_LEN + PE_SUBREPORT_NAME_LEN

Declare Function PEGetSubreportInfo Lib "crpe32.dll" (ByVal printJob%, ByVal subreportHandle As Long, subreportInfo As PESubreportInfo) As Integer

Declare Function PEOpenSubreport Lib "crpe32.dll" (ByVal parentJob%, ByVal subreportName$) As Integer
Declare Function PECloseSubreport Lib "crpe32.dll" (ByVal printJob%) As Integer
' End Of Declarations

Function PE_SECTION_CODE(sectionType%, groupN%, sectionN%) As Integer
' A function to create section codes:
' (This representation allows up to 25 groups and 40 sections of a given
' type, although Crystal Reports itself has no such limitations.)
    PE_SECTION_CODE = (((sectionType) * 1000) + ((groupN) Mod 25) + (((sectionN) Mod 40) * 25))
End Function

Function PE_AREA_CODE(sectionType%, groupN%) As Integer
'A function to create area codes:
    PE_AREA_CODE = PE_SECTION_CODE(sectionType, groupN, 0)
End Function

Function PE_GROUP_N(sectionCode%) As Integer
' Function to decode Group Number from section codes:
    PE_GROUP_N = ((sectionCode) Mod 25)
End Function

Function PE_SECTION_N(sectionCode) As Integer
' Function to decode Section Number from section codes:
   PE_SECTION_N = (((sectionCode \ 25) Mod 40))
End Function

Function PE_SECTION_TYPE(sectionCode%) As Integer
' Function to decode type from section codes:
    PE_SECTION_TYPE = ((sectionCode) \ 1000)
End Function

'Function to simplify PEGetVersion
Function PEVBGetVersion(ByVal component%) As Single
    Dim version As Integer
    Dim major As Integer
    Dim minor As Integer
    version = PEGetVersion(component)
    If version = 0 Then
        PEVBGetVersion = 0
    Else
        major = version / 256
        minor = version Mod 256
        PEVBGetVersion = major + (minor / 10)
    End If
End Function


Function PEGetJobStatus(ByVal job As Integer, info As PEJobInfo) As Integer
' To work around the problem of 4 - Byte alignment the PEGetJobStatus
' call has been re-declared here. When your application calls PEGetJobStatus
' it is calling this function which in turn calls CRPE32.DLL.
Dim splitinfo As SplitPEJobInfo
Dim temp1 As Long
Dim temp2 As Long

splitinfo.StructSize = PE_SIZEOF_JOB_INFO
PEGetJobStatus = RealPEGetJobStatus(job, splitinfo)
If PEGetJobStatus <> -1 Then
    temp1 = splitinfo.NumRecordsRead1
    If temp1 < 0 Then
        temp1 = 65536 + temp1
    End If
    temp2 = splitinfo.NumRecordsRead2
    If temp2 < 0 Then
        temp2 = 65536 + temp2
    End If
    temp2 = temp2 * 65536
    info.NumRecordsRead = temp1 + temp2
    
    temp1 = splitinfo.NumRecordsSelected1
    If temp1 < 0 Then
        temp1 = 65536 + temp1
    End If
    temp2 = splitinfo.NumRecordsSelected2
    If temp2 < 0 Then
        temp2 = 65536 + temp2
    End If
    temp2 = temp2 * 65536
    info.NumRecordsSelected = temp1 + temp2
    
    temp1 = splitinfo.NumRecordsPrinted1
    If temp1 < 0 Then
        temp1 = 65536 + temp1
    End If
    temp2 = splitinfo.NumRecordsPrinted2
    If temp2 < 0 Then
        temp2 = 65536 + temp2
    End If
    info.NumRecordsPrinted = temp1 + temp2
    info.LatestPageN = splitinfo.LatestPageN
    info.StartPageN = splitinfo.StartPageN
    info.DisplayPageN = splitinfo.DisplayPageN
    info.PrintEnded = splitinfo.PrintEnded
End If
End Function

