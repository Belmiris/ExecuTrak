VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsComboDropDown"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***********************************************************'
'
' Copyright (c) 1996 FACTOR, A Division of W.R.Hess Company
'
' Module name   : DROPDOWN.CLS
' Date          : August 19, 1996
' Programmer(s) : Qinggang Ma
'
' This module defines a class for a combo dropdown.

Option Explicit
    Private Const MASK_MESSAGE = 1
    Private Const MASK_ON_OFF = 2
    Private Const MASK_CLEARED = 4
    Private Const MASK_ACTIVE = 8
    Private Const MASK_SEARCH_ON_RETURN = &H10
    Private Const MASK_EXTEND_RESET = &H20
    Private Const MASK_HOR_SCROLL = &H40
    Private Const MASK_REQUERY = &H80
    Private Const MASK_DATA_ENTERED = &H100

    Private sSql As String
    Private sErrorMessage As String
    Private colCombos As Collection
    Private nTableLeft As Integer
    Private nTableTop As Integer
    Private nTableWidth As Integer
    Private nVisibleRows As Integer
    Private nPosition As Integer
    Private nHeadLines As Integer
    Private nDatabase As Integer
    Private nFlags As Integer
'    Private bNoDataMessage As Boolean

Property Let ComboOn(bStatus As Boolean)
    If bStatus Then
        nFlags = nFlags Or MASK_ON_OFF
    Else
        nFlags = nFlags And Not MASK_ON_OFF
    End If
End Property

Property Let ExtendReset(bStatus As Boolean)
    If bStatus Then
        nFlags = nFlags Or MASK_EXTEND_RESET
    Else
        nFlags = nFlags And Not MASK_EXTEND_RESET
    End If
End Property

Property Let Cleared(bStatus As Boolean)
    If bStatus Then
        nFlags = nFlags Or MASK_CLEARED
    Else
        nFlags = nFlags And Not MASK_CLEARED
    End If
End Property

Property Let Active(bStatus As Boolean)
    If bStatus Then
        nFlags = nFlags Or MASK_ACTIVE
    Else
        nFlags = nFlags And Not MASK_ACTIVE
    End If
End Property

Property Let dbType(nType As Integer)
    nDatabase = nType
End Property

Property Get ComboOn() As Boolean
    If (nFlags And MASK_ON_OFF) = MASK_ON_OFF Then
        ComboOn = True
    Else
        ComboOn = False
    End If
End Property

Property Get DataEntered() As Boolean
    If (nFlags And MASK_DATA_ENTERED) = MASK_DATA_ENTERED Then
        DataEntered = True
    Else
        DataEntered = False
    End If
End Property

Property Let DataEntered(bStatus As Boolean)
    If bStatus Then
        nFlags = nFlags Or MASK_DATA_ENTERED
    Else
        nFlags = nFlags And Not MASK_DATA_ENTERED
    End If
End Property

Property Get ReQuery() As Boolean
    If (nFlags And MASK_REQUERY) = MASK_REQUERY Then
        ReQuery = True
    Else
        ReQuery = False
    End If
End Property

Property Let ReQuery(bStatus As Boolean)
    If bStatus Then
        nFlags = nFlags Or MASK_REQUERY
    Else
        nFlags = nFlags And Not MASK_REQUERY
    End If
End Property


Property Get ExtendReset() As Boolean
    If (nFlags And MASK_EXTEND_RESET) = MASK_EXTEND_RESET Then
        ExtendReset = True
    Else
        ExtendReset = False
    End If
End Property


Property Get Cleared() As Boolean
    If (nFlags And MASK_CLEARED) = MASK_CLEARED Then
        Cleared = True
    Else
        Cleared = False
    End If
End Property


Property Get Active() As Boolean
    If (nFlags And MASK_ACTIVE) = MASK_ACTIVE Then
        Active = True
    Else
        Active = False
    End If
End Property

Property Get dbType() As Integer
    dbType = nDatabase
End Property

Property Get SearchOnReturn() As Boolean
    If (nFlags And MASK_SEARCH_ON_RETURN) = MASK_SEARCH_ON_RETURN Then
        SearchOnReturn = True
    Else
        SearchOnReturn = False
    End If
End Property


Property Let SearchOnReturn(bTemp As Boolean)
    If bTemp Then
        nFlags = nFlags Or MASK_SEARCH_ON_RETURN
    Else
        nFlags = nFlags And Not MASK_SEARCH_ON_RETURN
    End If
End Property

Property Get ErrorMessage() As String
    ErrorMessage = sErrorMessage
End Property
Property Let ErrorMessage(sMsg As String)
    sErrorMessage = sMsg
End Property

Property Let ShowNoDataMessage(bStatus As Boolean)
    If bStatus Then
        nFlags = nFlags Or MASK_MESSAGE
    Else
        nFlags = nFlags And Not MASK_MESSAGE
    End If
End Property

Property Get ShowNoDataMessage() As Boolean
    If (nFlags And MASK_MESSAGE) = MASK_MESSAGE Then
        ShowNoDataMessage = True
    Else
        ShowNoDataMessage = False
    End If
End Property

Property Let Position(nPos As Integer)
    nPosition = nPos
End Property
Property Get Position() As Integer
    Position = nPosition
End Property


Property Let TableLeft(L As Integer)
    nTableLeft = L
End Property
Property Get TableLeft() As Integer
    TableLeft = nTableLeft
End Property


Property Set Combos(c As Collection)
    Set colCombos = c
End Property
Property Get Combos() As Collection
    Set Combos = colCombos
End Property

Property Let TableTop(T As Integer)
    nTableTop = T
End Property
Property Get TableTop() As Integer
    TableTop = nTableTop
End Property

Property Let TableWidth(w As Long)
    nTableWidth = w
End Property
Property Get TableWidth() As Long
    TableWidth = nTableWidth
End Property
Property Let SQL(s As String)
    sSql = s
End Property
Property Get SQL() As String
    SQL = sSql
End Property

Property Get HeadLines() As Integer
    HeadLines = nHeadLines
End Property

Property Let HeadLines(nLines As Integer)
    nHeadLines = nLines
End Property

Property Let VisibleRows(nRows As Integer)
    nVisibleRows = nRows
End Property
Property Get VisibleRows() As Integer
    VisibleRows = nVisibleRows
End Property


Private Sub Class_Initialize()
    sErrorMessage = ""
End Sub


Private Sub Class_Terminate()
    If Not colCombos Is Nothing Then
        While colCombos.Count > 0
            colCombos.Remove 1
        Wend
    End If
End Sub


Property Get HorScroll() As Boolean
    If (nFlags And MASK_HOR_SCROLL) = MASK_HOR_SCROLL Then
        HorScroll = True
    Else
        HorScroll = False
    End If
End Property

Property Let HorScroll(bTemp As Boolean)
    If bTemp Then
        nFlags = nFlags Or MASK_HOR_SCROLL
    Else
        nFlags = nFlags And Not MASK_HOR_SCROLL
    End If
End Property

