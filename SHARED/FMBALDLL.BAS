Attribute VB_Name = "modFMBALDLL"
Option Explicit

#If DEVELOP Then
    Declare Function BALANCE_Attach Lib "fmbal16d.dll" _
        (ByVal dPrec As Double, ByRef hProc As Long) As Integer
        
    Declare Function BALANCE_Detach Lib "fmbal16d.dll" _
        (ByVal hProc As Long) As Integer
        
    Declare Function BALANCE_AddLine Lib "fmbal16d.dll" _
        (ByVal hProc As Long, ByVal lBOLNumber As Long, _
         ByVal lDELNumber As Long, ByVal lProdLink As Long, _
         ByVal dNetAmount As Double, _
         ByVal dGrossAmount As Double) As Integer
             
    Declare Function BALANCE_UpdateLine Lib "fmbal16d.dll" _
        (ByVal hProc As Long, ByVal lBOLNumber As Long, _
         ByVal lDELNumber As Long, ByVal lProdLink As Long, _
         ByVal dNetAmount As Double, _
         ByVal dGrossAmount As Double) As Integer
         
    Declare Function BALANCE_DeleteLine Lib "fmbal16d.dll" _
        (ByVal hProc As Long, ByVal lBOLNumber As Long, _
         ByVal lDELNumber As Long, ByVal lProdLink As Long _
        ) As Integer
    
    Declare Function BALANCE_GetUnbalancedCount Lib "fmbal16d.dll" _
        (ByVal hProc As Long, ByRef nCount As Integer) As Integer
        
    Declare Function BALANCE_GetUnbalancedRecords Lib "fmbal16d.dll" _
        (ByVal hProc As Long, ByRef pBalArray As CUnbalancedRec) _
        As Integer
#Else
    Declare Function BALANCE_Attach Lib "fmbal16.dll" _
        (ByVal dPrec As Double, ByRef hProc As Long) As Integer
        
    Declare Function BALANCE_Detach Lib "fmbal16.dll" _
        (ByVal hProc As Long) As Integer
        
    Declare Function BALANCE_AddLine Lib "fmbal16.dll" _
        (ByVal hProc As Long, ByVal lBOLNumber As Long, _
         ByVal lDELNumber As Long, ByVal lProdLink As Long, _
         ByVal dNetAmount As Double, _
         ByVal dGrossAmount As Double) As Integer
             
    Declare Function BALANCE_UpdateLine Lib "fmbal16.dll" _
        (ByVal hProc As Long, ByVal lBOLNumber As Long, _
         ByVal lDELNumber As Long, ByVal lProdLink As Long, _
         ByVal dNetAmount As Double, _
         ByVal dGrossAmount As Double) As Integer
         
    Declare Function BALANCE_DeleteLine Lib "fmbal16.dll" _
        (ByVal hProc As Long, ByVal lBOLNumber As Long, _
         ByVal lDELNumber As Long, ByVal lProdLink As Long _
        ) As Integer
    
    Declare Function BALANCE_GetUnbalancedCount Lib "fmbal16.dll" _
        (ByVal hProc As Long, ByRef nCount As Integer) As Integer
        
    Declare Function BALANCE_GetUnbalancedRecords Lib "fmbal16.dll" _
        (ByVal hProc As Long, ByRef pBalArray As CUnbalancedRec) _
        As Integer
#End If

Public Type CUnbalancedRec
    lBOLNumber As Long
    lProductLink As Long
    dNetAmount As Double
    dGrossAmount As Double
    
    ' NOTE: a positive amount indicates that there is
    ' more product on the BOL side than on the DEL side
End Type

