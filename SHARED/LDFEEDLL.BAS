Attribute VB_Name = "modLDFEEDLL"
Option Explicit

#If DEVELOP Then
    Declare Function LDFEE_Attach Lib "ldfee16d.dll" _
        (ByVal hInst As Integer, ByVal hWnd As Integer, _
        ByVal szConnect As String) As Integer
    
    Declare Function LDFEE_Detach Lib "ldfee16d.dll" _
        (ByVal hInst As Integer, ByVal hWnd As Integer) As Integer
        
    Declare Function LDFEE_AddRecord Lib "ldfee16d.dll" _
        (ByVal hInst As Integer, ByVal hWnd As Integer, _
        ByRef bpLDFBalRec As CLDFBalRec) As Integer
        
    Declare Function LDFEE_Calculate Lib "ldfee16d.dll" _
        (ByVal hInst As Integer, ByVal hWnd As Integer, _
        ByRef nRecCount As Integer) As Integer
    
    Declare Function LDFEE_GetLoadingFees Lib "ldfee16d.dll" _
        (ByVal hInst As Integer, ByVal hWnd As Integer, _
        ByRef bpLDFRecords As CLoadingFeeRec) As Integer
#Else
    Declare Function LDFEE_Attach Lib "ldfee16.dll" _
        (ByVal hInst As Integer, ByVal hWnd As Integer, _
        ByVal szConnect As String) As Integer
    
    Declare Function LDFEE_Detach Lib "ldfee16.dll" _
        (ByVal hInst As Integer, ByVal hWnd As Integer) As Integer
        
    Declare Function LDFEE_AddRecord Lib "ldfee16.dll" _
        (ByVal hInst As Integer, ByVal hWnd As Integer, _
        ByRef bpLDFBalRec As CLDFBalRec) As Integer
        
    Declare Function LDFEE_Calculate Lib "ldfee16.dll" _
        (ByVal hInst As Integer, ByVal hWnd As Integer, _
        ByRef nRecCount As Integer) As Integer
    
    Declare Function LDFEE_GetLoadingFees Lib "ldfee16.dll" _
        (ByVal hInst As Integer, ByVal hWnd As Integer, _
        ByRef bpLDFRecords As CLoadingFeeRec) As Integer
#End If

' Error codes
Public Const LDFEE_OK = &H0                 ' All's well
Public Const LDFEE_NOMEM = &H7FFF           ' Out of memory
Public Const LDFEE_BADPARMS = &H7FFE        ' Bad parameters passed
Public Const LDFEE_NODBASE = &H7FFD         ' No database available
Public Const LDFEE_NOAPPNODE = &H7FFC       ' Process not registered
Public Const LDFEE_OUTOFRANGE = &H7FFB      ' Index out of range
Public Const LDFEE_NOLDFEENODE = &H7FFA     ' No loading fee structure allocated
Public Const LDFEE_TASKBLOCKED = &H7FF9     ' DLL crtical section in use

Public Const LDFEE_OUTOFMEM = 1
Public Const LDFEE_DUPENTRY = 2             ' Duplicate balance record
Public Const LDFEE_NOINVHEADER = 4          ' No inv_header for product
Public Const LDFEE_NOINVCLASS = 5           ' No inv_classes for product
Public Const LDFEE_NOPRODTAXCLASS = 6       ' No tax class for product
Public Const LDFEE_NOLDFEERECS = 7          ' No loading fee records available
Public Const LDFEE_OUTOFSYNC = 8            ' FATAL FATAL ERROR!
Public Const LDFEE_NODBASEPTR = 9           ' No database supplied
Public Const LDFEE_BADSQL = 10              ' Error in query
Public Const LDFEE_NOTAXRECORDS = 11        ' No tx_table records exist
Public Const LDFEE_NOARECORDS = 12          ' No loading fee records exist
Public Const LDFEE_NOSYSPARM = 13           ' Can't get system parameter

Public Type CLDFBalRec
    lBillOfLading As Long
    lDelivery As Long
    lBOLProdLink As Long
    lDELProdLink As Long
    dNetAmount As Double
    sBOLUseGroup As String * 8
    sDELUseGroup As String * 8
End Type

Public Type CLoadingFeeRec
    lBillOfLading As Long
    lDelivery As Long
    lProductLink As Long
    sUseGroup As String * 8
    larryGlAcct(0 To 3) As Long
    carryCharge(0 To 3) As Byte
    dFee As Double
End Type
