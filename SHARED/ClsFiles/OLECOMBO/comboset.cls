VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsComboSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***********************************************************'
'
' Copyright (c) 1996 FACTOR, A Division of W.R.Hess Company
'
' Module name   : COMBOSET.CLS
' Date          : August 19, 1996
' Programmer(s) : Qinggang Ma
'
' This module defines a class which contains a text box and a command button
' corresponding to a column in the dropdown true grid.

Option Explicit
    
Private Const MASK_EXTN_DIR = &H3
Private Const MASK_KEY_FIELD = &H4
Private Const MASK_APPEND_CRTR = &H8
Private Const MASK_GET_TEXT = &H10
Private Const MASK_TEXT_CHANGED = &H20
Private Const MASK_TEXT_VALID = &H40
Private Const MASK_AUTO_WIDTH = &H80
Private Const MASK_VISIBLE = &H100

Private txtBox As Control
Private cmdButton As Control
Private sColCaption As String
Private sDataField As String
Private sFieldValue As String
Private sAliasName As String
Private sRegPattern As String
Private nFieldType As Integer
Private nColumnWidth As Integer
Private nOrderFlag As Integer
Private nFlags As Integer

'david 01/28/2002
Private sOrderByString As String
Private bSortAscending As Boolean
''''''''''''''''''
'

Property Let Caption(sCap As String)
    sColCaption = sCap
End Property

Property Get Caption() As String
    Caption = sColCaption
End Property

Property Let AliasName(sName As String)
    sAliasName = sName
End Property

Property Get AliasName() As String
    AliasName = sAliasName
End Property

Property Set Button(b As Object)
    Set cmdButton = b
End Property

Property Get Button() As Object
    Set Button = cmdButton
End Property

Property Get ComboType() As Integer
    ComboType = nFieldType
End Property

Property Let ComboType(nType As Integer)
    nFieldType = nType
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

Property Let ExtDir(nTemp As Integer)
    nFlags = (nFlags And Not MASK_EXTN_DIR) Or nTemp
End Property

Property Get FieldValue() As String
    FieldValue = sFieldValue
End Property

Property Let FieldValue(sTemp As String)
    sFieldValue = sTemp
End Property

Property Let KeyField(bTemp As Boolean)
    If bTemp Then
        nFlags = nFlags Or MASK_KEY_FIELD
    Else
        nFlags = nFlags And Not MASK_KEY_FIELD
    End If
End Property

Property Let AppendCriteria(bTemp As Boolean)
    If bTemp Then
        nFlags = nFlags Or MASK_APPEND_CRTR
    Else
        nFlags = nFlags And Not MASK_APPEND_CRTR
    End If
End Property

Property Let Populate(bTemp As Boolean)
    If bTemp Then
        nFlags = nFlags Or MASK_GET_TEXT
    Else
        nFlags = nFlags And Not MASK_GET_TEXT
    End If
End Property

Property Let Valid(bTemp As Boolean)
    If bTemp Then
        nFlags = nFlags Or MASK_TEXT_VALID
    Else
        nFlags = nFlags And Not MASK_TEXT_VALID
    End If
End Property


Property Get KeyField() As Boolean
    If (nFlags And MASK_KEY_FIELD) = MASK_KEY_FIELD Then
        KeyField = True
    Else
        KeyField = False
    End If
End Property

Property Get AppendCriteria() As Boolean
    If (nFlags And MASK_APPEND_CRTR) = MASK_APPEND_CRTR Then
        AppendCriteria = True
    Else
        AppendCriteria = False
    End If
End Property

Property Get Populate() As Boolean
    If (nFlags And MASK_GET_TEXT) = MASK_GET_TEXT Then
        Populate = True
    Else
        Populate = False
    End If
End Property

Property Get Valid() As Boolean
    If (nFlags And MASK_TEXT_VALID) = MASK_TEXT_VALID Then
        Valid = True
    Else
        Valid = False
    End If
End Property

Property Get Changed() As Boolean
    If (nFlags And MASK_TEXT_CHANGED) = MASK_TEXT_CHANGED Then
        Changed = True
    Else
        Changed = False
    End If
End Property

Property Get AutoWidth() As Boolean
    If (nFlags And MASK_AUTO_WIDTH) = MASK_AUTO_WIDTH Then
        AutoWidth = True
    Else
        AutoWidth = False
    End If
End Property

Property Let Changed(bTemp As Boolean)
    If bTemp Then
        nFlags = nFlags Or MASK_TEXT_CHANGED
    Else
        nFlags = nFlags And Not MASK_TEXT_CHANGED
    End If
End Property

Property Let AutoWidth(bTemp As Boolean)
    If bTemp Then
        nFlags = nFlags Or MASK_AUTO_WIDTH
    Else
        nFlags = nFlags And Not MASK_AUTO_WIDTH
    End If
End Property

Property Get ExtDir() As Integer
    ExtDir = nFlags And MASK_EXTN_DIR
End Property

Property Let OrderFlag(nFlag As Integer)
    nOrderFlag = nFlag
End Property

Property Get OrderFlag() As Integer
    OrderFlag = nOrderFlag
End Property

Property Get RegExpPattern() As String
    RegExpPattern = sRegPattern
End Property

Property Let RegExpPattern(sPattern As String)
    sRegPattern = sPattern
End Property

Property Set Textbox(Box As Object)
    Set txtBox = Box
End Property

Property Get Textbox() As Object
    Set Textbox = txtBox
End Property

Property Get Visible() As Boolean
    If (nFlags And MASK_VISIBLE) = MASK_VISIBLE Then
        Visible = True
    Else
        Visible = False
    End If
End Property

Property Let Visible(bTemp As Boolean)
    If bTemp Then
        nFlags = nFlags Or MASK_VISIBLE
    Else
        nFlags = nFlags And Not MASK_VISIBLE
    End If
End Property

Private Sub Class_Initialize()
    Set txtBox = Nothing
    Set cmdButton = Nothing
    sColCaption = ""
    sDataField = ""
    sAliasName = ""
    sRegPattern = ""
End Sub

Private Sub Class_Terminate()
    Set txtBox = Nothing
    Set cmdButton = Nothing
End Sub

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

