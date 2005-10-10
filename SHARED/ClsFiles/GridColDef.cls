VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GridColDef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Member Variables for Properties
Private m_Key           As String
Private m_Header        As String
Private m_FieldName     As String
Private m_WidthPct      As Single
Private m_Visible       As Boolean
Private m_HiddenField   As Boolean
Private m_RegExp        As String
Private m_Prompt        As String
Private m_Editable      As Boolean
Private m_Locked        As Boolean
Private m_Alignment     As AlignmentConstants
Private m_DisplayFormat As String
Private m_Index         As Integer
Public Property Let Alignment(ByVal pAlignment As AlignmentConstants)
    m_Alignment = pAlignment
End Property
Public Property Get Alignment() As AlignmentConstants
    Alignment = m_Alignment
End Property
Public Property Let DisplayFormat(ByVal Format As String)
    m_DisplayFormat = Format
End Property
Public Property Get DisplayFormat() As String
    DisplayFormat = m_DisplayFormat
End Property
Public Property Let Editable(ByVal AllowEdit As Boolean)
    m_Editable = AllowEdit
End Property
Public Property Get Editable() As Boolean
    Editable = m_Editable
End Property
Public Property Let FieldName(ByVal pFieldName As String)
    m_FieldName = pFieldName
End Property
Public Property Get FieldName() As String
    FieldName = m_FieldName
End Property
Public Property Let Header(ByVal pHeader As String)
    m_Header = pHeader
End Property
Public Property Get Header() As String
    Header = m_Header
End Property
Public Property Let HiddenField(ByVal IsHiddenField As Boolean)
    m_HiddenField = IsHiddenField
End Property
Public Property Get HiddenField() As Boolean
    HiddenField = m_HiddenField
End Property
Friend Property Let Index(ByVal pIndex As Integer)
    m_Index = pIndex
End Property
Public Property Get Index() As Integer
    Index = m_Index
End Property
Public Property Let Key(ByVal pKey As String)
    If LenB(m_Key) = 0 Then
        m_Key = pKey
    End If
End Property
Public Property Get Key() As String
    Key = m_Key
End Property
Public Property Let Prompt(ByVal pPrompt As String)
    m_Prompt = pPrompt
End Property
Public Property Get Prompt() As String
    Prompt = m_Prompt
End Property
Public Property Let RegExp(ByVal Exp As String)
    m_RegExp = Exp
End Property
Public Property Get RegExp() As String
    RegExp = m_RegExp
End Property
Public Property Let Visible(ByVal pVisible As Boolean)
    m_Visible = pVisible
End Property
Public Property Get Visible() As Boolean
    Visible = m_Visible
End Property
Public Property Let WidthPct(ByVal Pct As Single)
    m_WidthPct = Pct
End Property
Public Property Get WidthPct() As Single
    WidthPct = m_WidthPct
End Property
Private Sub Class_Initialize()
    m_Visible = True
    m_Editable = True
End Sub
Public Property Let Locked(ByVal pLocked As Boolean)
    m_Locked = pLocked
End Property
Public Property Get Locked() As Boolean
    Locked = m_Locked
End Property
