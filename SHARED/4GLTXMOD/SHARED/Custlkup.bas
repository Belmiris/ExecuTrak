Attribute VB_Name = "modCustLookUp"
'***********************************************************'
'
' Copyright (c) 1996 FACTOR, A Division of W.R.Hess Company
'
' Module name   : CUSTLKUP.BAS
' Date          : DEC 29, 1998
' Programmer(s) : Qinggang Ma
'
' An interface module to call Customer Lookup OLE.
'
'1. Setup Sub:     CreateCustomerLookUp
'2. Cleanup Sub:  DestroyCustomerLookUp
'3. Check Availability Function:     CustLookupAvailable (Boolean)
'4. Call Back Sub: BackFromCustomerLookup
'5. Sub to popup Customer Lookup screen:
'       ShowCustLookup
'6. Event call (F3 to popup the sceen):
'       CustKeyDown(KeyCode As Integer, Shift As Integer)
'7. Function to check the action of the operator:
'       CustDataFilled (boolean)
'       True:   A customer number has been selected
'       Fasle:  Canceled
'8. Functions to retrieve results:
'       CustAddress1    : Return the first line of addredss
'       CustAddress2    : Return the second line of addredss
'       CustAttention   : Return the attention line
'       CustCity        : Return the customer city code
'       CustCityName    : Return the customer city name
'       CustFirstName   : Return the first name
'       CustLastName    : Return the last name
'       CustNewNumber   : Return the new (Factor) number of the customer
'       CustOldNumber   : Return the old number of the customer (if any)
'       CustPhone       : Return the phone number of the customer
'       CustState       : Return the customer state code
'       CustStateNmae   : Return the customer state name
'       CustZipCode     : Return the zip code of the customer

Option Explicit
    #Const INCLUDE_SOURCE = False
    
    Private objCustomerLookup As Object
    
Public Function CustDataFilled() As Boolean

    If Not objCustomerLookup Is Nothing Then
        CustDataFilled = objCustomerLookup.DataFilled
    Else
        #If INCLUDE_SOURCE Then
            CustDataFilled = frmCustLookUp.DataFilled
        #End If
    End If

End Function

Public Function CustNewNumber() As Long
    If Not objCustomerLookup Is Nothing Then
        CustNewNumber = objCustomerLookup.FactorNumber
    Else
        #If INCLUDE_SOURCE Then
            CustNewNumber = val(frmCustLookUp.txtCustomerNumber)
        #End If
    End If
End Function

Public Function CustOldNumber() As String
    If Not objCustomerLookup Is Nothing Then
        CustOldNumber = objCustomerLookup.OldNumber
    Else
        #If INCLUDE_SOURCE Then
            CustOldNumber = frmCustLookUp.txtOldCust
        #End If
    End If
End Function

Public Function CustAddress1() As String
    If Not objCustomerLookup Is Nothing Then
        CustAddress1 = objCustomerLookup.Address1
    Else
        #If INCLUDE_SOURCE Then
            CustAddress1 = frmCustLookUp.txtAddress1
        #End If
    End If
End Function

Public Function CustAddress2() As String
    If Not objCustomerLookup Is Nothing Then
        CustAddress2 = objCustomerLookup.Address2
    Else
        #If INCLUDE_SOURCE Then
            CustAddress2 = frmCustLookUp.txtAddress2
        #End If
    End If
End Function

Public Function CustAttention() As String
    If Not objCustomerLookup Is Nothing Then
        CustAttention = objCustomerLookup.Attention
    Else
        #If INCLUDE_SOURCE Then
            CustAttention = frmCustLookUp.txtAttention
        #End If
    End If
End Function

Public Function CustCity() As String
    If Not objCustomerLookup Is Nothing Then
        CustCity = objCustomerLookup.City
    Else
        #If INCLUDE_SOURCE Then
            CustCity = frmCustLookUp.txtCity
        #End If
    End If
End Function

Public Function CustCityName() As String
    If Not objCustomerLookup Is Nothing Then
        CustCityName = objCustomerLookup.CityName
    Else
        #If INCLUDE_SOURCE Then
            CustCityName = frmCustLookUp.txtCityName
        #End If
    End If
End Function

Public Function CustFirstName() As String
    If Not objCustomerLookup Is Nothing Then
        CustFirstName = objCustomerLookup.FirstName
    Else
        #If INCLUDE_SOURCE Then
            CustFirstName = frmCustLookUp.txtFirstName
        #End If
    End If
End Function

Public Sub CustKeyDown(KeyCode As Integer, _
                       Shift As Integer, _
                       Optional txtCust As Variant, _
                       Optional sCriteria As String = "")
    If KeyCode = vbKeyF3 Then
        ShowCustLookup txtCust, sCriteria
    End If

End Sub

Public Function CustLastName() As String
    If Not objCustomerLookup Is Nothing Then
        CustLastName = objCustomerLookup.LastName
    Else
        #If INCLUDE_SOURCE Then
            CustLastName = frmCustLookUp.txtLastName
        #End If
    End If
End Function

Public Function CustPhone() As String
    If Not objCustomerLookup Is Nothing Then
        CustPhone = objCustomerLookup.Phone
    Else
        #If INCLUDE_SOURCE Then
            CustPhone = frmCustLookUp.txtPhone
        #End If
    End If
End Function

Public Function CustState() As String
    If Not objCustomerLookup Is Nothing Then
        CustState = objCustomerLookup.State
    Else
        #If INCLUDE_SOURCE Then
            CustState = frmCustLookUp.txtState
        #End If
    End If
End Function

Public Function CustStateName() As String
    If Not objCustomerLookup Is Nothing Then
        CustStateName = objCustomerLookup.StateName
    Else
        #If INCLUDE_SOURCE Then
            CustStateName = frmCustLookUp.txtStateName
        #End If
    End If
End Function

Public Function CustZipCode() As String
    If Not objCustomerLookup Is Nothing Then
        CustZipCode = objCustomerLookup.ZipCode
    Else
        #If INCLUDE_SOURCE Then
            CustZipCode = frmCustLookUp.txtZipCode
        #End If
    End If
End Function

Public Function CustLookupAvailable() As Boolean
    If objCustomerLookup Is Nothing Then
        CustLookupAvailable = False
    Else
        CustLookupAvailable = True
    End If
End Function

Public Sub ShowCustLookup(Optional txtCust As Variant, _
                            Optional sCriteria As String = "")
    If Not objCustomerLookup Is Nothing Then
        On Error GoTo tryOld1
        objCustomerLookup.ShowLookup txtCust, sCriteria
    End If
    Exit Sub

tryOld1:
    Resume tryOld2
tryOld2:
    On Error Resume Next
    objCustomerLookup.ShowLookup
End Sub

Public Sub CreateCustomerLookUp(frmCaller As Form)

    If objCustomerLookup Is Nothing Then
        #If INCLUDE_SOURCE Then
            Set objCustomerLookup = New CustomerLookup
        #Else
            On Error Resume Next
            Set objCustomerLookup = CreateObject("CUSTLKUP.CustomerLookup")
        #End If
        If Not objCustomerLookup Is Nothing Then
            With objCustomerLookup
                Set .MainForm = frmCaller
                Set .CDatabase = t_dbMainDatabase
                Set .CDBEngine = t_engFactor
            End With
        End If
    End If
End Sub

Public Sub DestroyCustomerLookUp()

    Set objCustomerLookup = Nothing
    
End Sub



