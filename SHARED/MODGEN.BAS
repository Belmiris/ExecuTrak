Attribute VB_Name = "modGEN"
'*************************************************************************
' Copyright (c) 1996 FACTOR, A Division of W.R.Hess Company
' Moduel Name: modGen
' Programmer: Weigong Jiang
' Nov. 26 ,1996
'
' A Small module for generating customer and vendor numbers
' in AR and AP File maintenance Based on the lastname entered
'
' public functions:
' (1) GetNextAr(szLastname as string)--- returns Customer number
' (2) GetNextAp(szLastname as string) --- returns vendor number
'

'*************************************************************************
Option Explicit

Public Const CHECK_AR = 100
Public Const CHECK_AP = 200

Private Type MyType
   ch As String
   nu As Long
End Type

Private Letter(1 To 27) As MyType
Private Sub Set_Up_Letter()
   Dim i As Integer
   Letter(1).ch = " "
   For i = 2 To 27 ' set letters 'a' to 'z'
     Letter(i).ch = Chr(Asc("a") + i - 2)
   Next i
   Letter(1).nu = 1
   Letter(2).nu = 4
   Letter(3).nu = 9
   Letter(4).nu = 12
   Letter(5).nu = 15
   Letter(6).nu = 18
   Letter(7).nu = 23
   Letter(8).nu = 26
   Letter(9).nu = 29
   Letter(10).nu = 32
   Letter(11).nu = 37
   Letter(12).nu = 40
   Letter(13).nu = 43
   Letter(14).nu = 46
   Letter(15).nu = 49
   Letter(16).nu = 52
   Letter(17).nu = 57
   Letter(18).nu = 60
   Letter(19).nu = 63
   Letter(20).nu = 66
   Letter(21).nu = 71
   Letter(22).nu = 74
   Letter(23).nu = 77
   Letter(24).nu = 80
   Letter(25).nu = 83
   Letter(26).nu = 86
   Letter(27).nu = 89
End Sub

' this function generate a long integer according to a char(2)
Private Function Let_Val(szChar1 As String, szChar2 As String) As Long
    Dim i As Long
    Set_Up_Letter
    If szChar1 = "" Then
       Let_Val = 1000000 ' 7_digit
    ElseIf Asc(szChar1) > Asc("z") Then
       Let_Val = 93000000 ' 8_digit
    ElseIf Asc(szChar1) < Asc("a") Then
       Let_Val = 93000000
    Else
       For i = 2 To 27
         If szChar1 = Letter(i).ch Then
            Let_Val = Letter(i).nu * 1000000
            Exit For
          End If
       Next i
    End If
    
    If szChar2 = "" Then
       Let_Val = Let_Val + 10000 ' 5_digit
    ElseIf Asc(szChar2) > Asc("z") Then
       Let_Val = Let_Val + 930000 ' 6_digit
    ElseIf Asc(szChar1) < Asc("a") Then
       Let_Val = Let_Val + 930000
    Else
       For i = 2 To 27
         If szChar2 = Letter(i).ch Then
            Let_Val = Let_Val + Letter(i).nu * 10000
            Exit For
          End If
       Next i
    End If
End Function
Private Function GetMaxAr(nFrom As Long, nTo As Long) As Long
    Dim szSQL As String
    Dim rsTemp As Recordset
    
    szSQL = "SELECT MAX(an_customer) max_ar FROM ar_altname WHERE an_customer BETWEEN " _
            & nFrom & " AND " & nTo
    On Error GoTo Errtrap
    Set rsTemp = t_dbMainDatabase.OpenRecordset(szSQL, dbOpenSnapshot, dbSQLPassThrough)
    If rsTemp.RecordCount > 0 Then
        If Not IsNull(rsTemp!max_ar) Then
            GetMaxAr = Val(rsTemp!max_ar)
        Else
            GetMaxAr = nFrom
        End If
    Else
       GetMaxAr = nFrom
    End If
    Exit Function
Errtrap:
    tfnErrHandler szSQL, "GetMaxAr"
End Function
Private Function GetMaxAp(nFrom As Long, nTo As Long) As Long
    Dim szSQL As String
    Dim rsTemp As Recordset
    szSQL = "SELECT MAX(pn_alt) ap_max FROM p_altname WHERE pn_alt BETWEEN " & CStr(nFrom) & " AND " & CStr(nTo)
    On Error GoTo Errtrap
    Set rsTemp = t_dbMainDatabase.OpenRecordset(szSQL, dbOpenSnapshot, dbSQLPassThrough)
    If rsTemp.RecordCount > 0 Then
        If Not IsNull(rsTemp!ap_max) Then
           GetMaxAp = Val(rsTemp!ap_max)
        Else
            GetMaxAp = nFrom
        End If
    Else
       GetMaxAp = nFrom
    End If
    Exit Function
Errtrap:
    tfnErrHandler szSQL, "GetMaxAp"
End Function

Public Function IsInDatabase(nNumber As Long, nArAp As Integer) As Boolean
    Dim szSQL As String
    Dim rsTemp As Recordset
    Dim bFound As Boolean
    
    If nArAp = CHECK_AP Then
        szSQL = "INSERT INTO p_altname(pn_alt) VALUES (" & CStr(nNumber) & ")"
    ElseIf nArAp = CHECK_AR Then
        szSQL = "INSERT INTO ar_altname(an_customer) VALUES (" & CStr(nNumber) & ")"
    End If
    On Error GoTo Errtrap
    t_dbMainDatabase.Execute szSQL, dbSQLPassThrough
    bFound = (t_dbMainDatabase.RecordsAffected = 0)
    IsInDatabase = bFound
    Exit Function
Errtrap:
    If t_engFactor.Errors(0).Number = -239 Then
        'Duplicate record found
        IsInDatabase = True
    Else
        tfnErrHandler szSQL, "IsInDatabase"
    End If
End Function

Public Function GetNextAr(szLastname As String, Optional vLock As Variant) As Long
    Dim a As String
    Dim b As String
    Dim X As Long
    Dim Y As Long
    Dim z As Long
    Dim i As Long
    Dim LOCK_RECORD As Boolean
    
    If IsMissing(vLock) Then
        LOCK_RECORD = False
    Else
        LOCK_RECORD = vLock
    End If
    
    szLastname = LCase(Trim(szLastname))
   
    If IsNull(szLastname) Then
       GetNextAr = GetMaxAr(1, 100000000)
       Exit Function
    Else 'len(szlastname) could be 0
       a = Mid(szLastname, 1, 1)
       b = Mid(szLastname, 2, 1)
    End If
    X = Let_Val(a, b)
    Y = 100000000 '9_digit
    
    For i = 1 To 26
       If Letter(i).ch = b Then
           Y = Let_Val(a, Letter(i + 1).ch)
           Exit For
       End If
    Next i
    z = GetMaxAr(X, Y)
    subCheckTheNumber z, CHECK_AR
    If LOCK_RECORD Then
        While IsInDatabase(z, CHECK_AR)
            subCheckTheNumber z, CHECK_AR
        Wend
    End If
    GetNextAr = z
    
End Function
Public Function GetNextAp(szLastname As String, Optional vLock As Variant) As Long
    Dim a As String
    Dim b As String
    Dim X As Long
    Dim Y As Long
    Dim z As Long
    Dim i As Long
    Dim LOCK_RECORD As Boolean
    If IsMissing(vLock) Then
        LOCK_RECORD = False
    Else
        LOCK_RECORD = vLock
    End If
    szLastname = LCase(Trim(szLastname))
   
    If IsNull(szLastname) Then
       GetNextAp = GetMaxAp(1, 100000000)
       Exit Function
    Else 'len(szlastname) could be 0
       a = Mid(szLastname, 1, 1)
       b = Mid(szLastname, 2, 1)
    End If
    X = Let_Val(a, b)
    Y = 100000000 '9_digit
    
    For i = 1 To 26
       If Letter(i).ch = b Then
           Y = Let_Val(a, Letter(i + 1).ch)
           Exit For
       End If
    Next i
    z = GetMaxAp(X, Y)
    subCheckTheNumber z, CHECK_AP
    If LOCK_RECORD Then
        While IsInDatabase(z, CHECK_AP)
            subCheckTheNumber z, CHECK_AP
        Wend
    End If
    
    GetNextAp = z
    
End Function

Private Sub subCheckTheNumber(lNumber As Long, nApAr As Integer)
    Const NUM_ELEVEN = 11
    
    Dim lRemdr As Long
    Dim lTemp As Long
    Dim bOK As Boolean
    
    bOK = False
    Do
        lNumber = lNumber + 1
        lTemp = Fix(lNumber / NUM_ELEVEN)
        lRemdr = lNumber - lTemp * NUM_ELEVEN
        Select Case nApAr
            Case CHECK_AP
                If lRemdr = 7 Or lRemdr = 9 Then
                    bOK = True
                End If
            Case CHECK_AR
                If lRemdr = 6 Or lRemdr = 8 Then
                    bOK = True
                End If
        End Select
    Loop Until bOK
End Sub

Public Function UnlockAr(ByVal nAR As Long) As Boolean
    Const SUB_NAME = "UnlockAr"
    
    Dim strSQL As String
    
    strSQL = "DELETE FROM ar_altname WHERE an_customer = " & CStr(nAR) _
           & " AND an_cust IS Null"
    On Error GoTo errUnlockAr
    t_dbMainDatabase.Execute strSQL, dbSQLPassThrough
    UnlockAr = True
    Exit Function
errUnlockAr:
    tfnErrHandler strSQL, SUB_NAME
    UnlockAr = False
End Function


Public Function UnlockAp(ByVal nAP As Long) As Boolean
    Const SUB_NAME = "UnlockAp"
    
    Dim strSQL As String
    
    strSQL = "DELETE FROM p_altname WHERE pn_alt = " & CStr(nAP) _
           & " AND pn_vendor IS Null"
    On Error GoTo errUnlockAp
    t_dbMainDatabase.Execute strSQL, dbSQLPassThrough
    UnlockAp = True
    Exit Function
errUnlockAp:
    tfnErrHandler strSQL, SUB_NAME
    UnlockAp = False
End Function

