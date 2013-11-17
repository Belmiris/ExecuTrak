VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsColumns"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***********************************************************'
'
' Copyright (c) 1996 FACTOR, A Division of W.R.Hess Company
'
' Module name   : COLUME.CLS
' Date          : May 16, 1996
' Programmer(s) : Qinggang Ma
'
' This module defines a class used for by the floating dropdown class module.

Option Explicit
    
Private Const FLAG_VISIBLE = 1
Private Const FLAG_CURSOR_CONTROL = 2
Private Const FLAG_STATUS = 4
Private Const FLAG_UNIQUE = 8
Private Const FLAG_POPULATE_BOX = 16
Private Const FLAG_MESSAGE = 32
Private Const FLAG_CAPTION = 64

Private nColumn As Integer
Private sCaption As String
Private sDataField As String
'david 10/22/2003
Private sSourceDataField As String
'''''''''''''''''
Private nColumnType As Integer
Private nColumnWidth As Integer

Private sSql As String
Private nDropDownType As Integer
Private nECDirection As Integer
Public m_nFlags As Integer

'david 10/25/2001
Private sOrderByString As String
'''''''''''''''''

'david 01/28/2002
Private bSortAscending As Boolean
'''''''''''''''''

'david 11/16/2013  #6035-772641
'0:not set, use global, -1:Yes, show dropdown on key Enter, 1:No, goto next column on key Enter
Private nSearchOnReturn As Integer
'''''''''''''''''''''''''''''''
'

Property Get Caption() As String
    Caption = sCaption
End Property

Property Let Caption(sCap As String)
    sCaption = sCap
End Property


Property Get ColType() As Integer
    ColType = nColumnType
End Property

Property Let ColType(nType As Integer)
    nColumnType = nType
End Property

Property Get ColumnIndex() As Integer
    ColumnIndex = nColumn
End Property

Property Let ColumnIndex(nCol As Integer)
    nColumn = nCol
End Property

Property Get ColWidth() As Integer
    ColWidth = nColumnWidth
End Property

Property Let ColWidth(nWidth As Integer)
    nColumnWidth = nWidth
End Property

Property Get DataField() As String
    DataField = sDataField
End Property

Property Let DataField(sField As String)
    sDataField = sField
End Property

'david 10/22/2003
Property Get SourceDataField() As String
    SourceDataField = sSourceDataField
End Property

Property Let SourceDataField(sField As String)
    sSourceDataField = sField
End Property
'''''''''''''''''

Property Get DropDownType() As Integer
    DropDownType = nDropDownType
End Property

Property Let DropDownType(nType As Integer)
    nDropDownType = nType
End Property

Property Let ExtDirection(nDir As Integer)
    nECDirection = nDir
End Property

Property Get ExtDirection() As Integer
    ExtDirection = nECDirection
End Property

Property Get SQL() As String
    SQL = sSql
End Property

Property Let SQL(strSQL As String)
    sSql = strSQL
End Property

Property Let Visible(bStatus As Boolean)
    If bStatus Then
        m_nFlags = m_nFlags Or FLAG_VISIBLE
    Else
        m_nFlags = m_nFlags And Not FLAG_VISIBLE
    End If
End Property

Property Let CursorControl(bStatus As Boolean)
    If bStatus Then
        m_nFlags = m_nFlags Or FLAG_CURSOR_CONTROL
    Else
        m_nFlags = m_nFlags And Not FLAG_CURSOR_CONTROL
    End If
End Property

Property Let Unique(bStatus As Boolean)
    If bStatus Then
        m_nFlags = m_nFlags Or FLAG_UNIQUE
    Else
        m_nFlags = m_nFlags And Not FLAG_UNIQUE
    End If
End Property

Property Let PopulateBox(bStatus As Boolean)
    If bStatus Then
        m_nFlags = m_nFlags Or FLAG_POPULATE_BOX
    Else
        m_nFlags = m_nFlags And Not FLAG_POPULATE_BOX
    End If
End Property

Property Let NoDataMessage(bStatus As Boolean)
    If bStatus Then
        m_nFlags = m_nFlags Or FLAG_MESSAGE
    Else
        m_nFlags = m_nFlags And Not FLAG_MESSAGE
    End If
End Property

Property Let ShowCaption(bStatus As Boolean)
    If bStatus Then
        m_nFlags = m_nFlags Or FLAG_CAPTION
    Else
        m_nFlags = m_nFlags And Not FLAG_CAPTION
    End If
End Property

Property Let Status(bStatus As Boolean)
    If bStatus Then
        m_nFlags = m_nFlags Or FLAG_STATUS
    Else
        m_nFlags = m_nFlags And Not FLAG_STATUS
    End If
End Property

Property Get Visible() As Boolean
    If m_nFlags And FLAG_VISIBLE Then
        Visible = True
    Else
        Visible = False
    End If
End Property

Property Get CursorControl() As Boolean
    If m_nFlags And FLAG_CURSOR_CONTROL Then
        CursorControl = True
    Else
        CursorControl = False
    End If
End Property

Property Get Unique() As Boolean
    If m_nFlags And FLAG_UNIQUE Then
        Unique = True
    Else
        Unique = False
    End If
End Property

Property Get PopulateBox() As Boolean
    If m_nFlags And FLAG_POPULATE_BOX Then
        PopulateBox = True
    Else
        PopulateBox = False
    End If
End Property

Property Get NoDataMessage() As Boolean
    If m_nFlags And FLAG_MESSAGE Then
        NoDataMessage = True
    Else
        NoDataMessage = False
    End If
End Property

Property Get ShowCaption() As Boolean
    If (m_nFlags And FLAG_CAPTION) = FLAG_CAPTION Then
        ShowCaption = True
    Else
        ShowCaption = False
    End If
End Property

Property Get Status() As Boolean
    If m_nFlags And FLAG_STATUS Then
        Status = True
    Else
        Status = False
    End If
End Property

Property Get OrderByFields() As String
    OrderByFields = sOrderByString
End Property

Property Let OrderByFields(sFields As String)
    sOrderByString = sFields
End Property

Property Get PrevOrderIsAscending() As Boolean
    PrevOrderIsAscending = bSortAscending
End Property

Property Let PrevOrderIsAscending(bFields As Boolean)
    bSortAscending = bFields
End Property

'0:not set, use global, -1:Yes, show dropdown on key Enter, 1:No, goto next column on key Enter
Property Let SearchOnReturn(nValue As Integer)
    nSearchOnReturn = nValue
End Property

Property Get SearchOnReturn() As Integer
    '0:not set, use global, -1:Yes, show dropdown on key Enter, 1:No, goto next column on key Enter
    SearchOnReturn = nSearchOnReturn
End Property
