VERSION 5.00
Object = "{C75015E0-2232-11D3-B440-0060971E99AF}#1.0#0"; "FACTFRM.OCX"
Object = "{01028C21-0000-0000-0000-000000000046}#4.0#0"; "TG32OV.OCX"
Begin VB.Form clsCRITERIA 
   Caption         =   "Criteria"
   ClientHeight    =   5985
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10170
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5985
   ScaleWidth      =   10170
   StartUpPosition =   3  'Windows Default
   Begin FACTFRMLib.FactorFrame efraFilterPopup 
      Height          =   3015
      Left            =   1380
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   1035
      Visible         =   0   'False
      Width           =   7560
      _Version        =   65536
      _ExtentX        =   13335
      _ExtentY        =   5318
      _StockProps     =   77
      BackColor       =   8388608
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelOuter      =   0
      Caption         =   "  Select Criteria"
      Style           =   7
      TitleBarHeight  =   0
      BeginProperty PanelFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin FACTFRMLib.FactorFrame efraForFocus 
         Height          =   2175
         Left            =   150
         TabIndex        =   7
         Top             =   195
         Width           =   7290
         _Version        =   65536
         _ExtentX        =   12859
         _ExtentY        =   3836
         _StockProps     =   77
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   5
         BeginProperty PanelFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin VB.TextBox txtFilter 
            Height          =   360
            HelpContextID   =   1000
            Left            =   165
            Locked          =   -1  'True
            TabIndex        =   13
            Top             =   990
            Width           =   6915
         End
         Begin VB.TextBox txtOp 
            Height          =   360
            HelpContextID   =   1001
            Left            =   180
            TabIndex        =   3
            Top             =   1680
            Width           =   570
         End
         Begin VB.TextBox txtValue 
            Height          =   360
            HelpContextID   =   1002
            Left            =   1395
            TabIndex        =   5
            Top             =   1680
            Width           =   5340
         End
         Begin FACTFRMLib.FactorFrame cmdOp 
            Height          =   360
            HelpContextID   =   1001
            Left            =   765
            TabIndex        =   4
            TabStop         =   0   'False
            Top             =   1680
            Width           =   360
            _Version        =   65536
            _ExtentX        =   635
            _ExtentY        =   635
            _StockProps     =   77
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CaptionPos      =   4
            Picture         =   "clsCRITERIA.frx":0000
            Style           =   3
            BorderWidth     =   4
            BeginProperty PanelFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin FACTFRMLib.FactorFrame cmdValue 
            Height          =   360
            HelpContextID   =   1002
            Left            =   6750
            TabIndex        =   6
            TabStop         =   0   'False
            Top             =   1680
            Width           =   360
            _Version        =   65536
            _ExtentX        =   635
            _ExtentY        =   635
            _StockProps     =   77
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CaptionPos      =   4
            Picture         =   "clsCRITERIA.frx":00F2
            Style           =   3
            BorderWidth     =   4
            BeginProperty PanelFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Label lblColumnName 
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Height          =   360
            Left            =   180
            TabIndex        =   22
            Top             =   315
            UseMnemonic     =   0   'False
            Width           =   6915
         End
         Begin VB.Label Label 
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Column Name"
            ForeColor       =   &H00000000&
            Height          =   270
            Index           =   0
            Left            =   180
            TabIndex        =   21
            Top             =   75
            UseMnemonic     =   0   'False
            Width           =   2010
         End
         Begin VB.Label Label 
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Operator"
            ForeColor       =   &H00000000&
            Height          =   270
            Index           =   2
            Left            =   180
            TabIndex        =   20
            Top             =   1410
            UseMnemonic     =   0   'False
            Width           =   1320
         End
         Begin VB.Label Label1 
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Value"
            ForeColor       =   &H00000000&
            Height          =   270
            Index           =   3
            Left            =   1440
            TabIndex        =   19
            Top             =   1410
            UseMnemonic     =   0   'False
            Width           =   2055
         End
         Begin VB.Label Label 
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Filter"
            ForeColor       =   &H00000000&
            Height          =   270
            Index           =   16
            Left            =   180
            TabIndex        =   18
            Top             =   750
            UseMnemonic     =   0   'False
            Width           =   825
         End
      End
      Begin FACTFRMLib.FactorFrame cmdFilterOK 
         Height          =   390
         HelpContextID   =   16
         Left            =   4641
         TabIndex        =   9
         Top             =   2490
         Width           =   1305
         _Version        =   65536
         _ExtentX        =   2302
         _ExtentY        =   688
         _StockProps     =   77
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   0   'False
         Caption         =   "O&K"
         CaptionPos      =   4
         PicturePos      =   3
         ShowFocusRect   =   -1  'True
         Style           =   3
         BorderWidth     =   4
         BeginProperty PanelFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin FACTFRMLib.FactorFrame cmdFilterCancel 
         Height          =   390
         HelpContextID   =   15
         Left            =   6135
         TabIndex        =   11
         Top             =   2490
         Width           =   1305
         _Version        =   65536
         _ExtentX        =   2302
         _ExtentY        =   688
         _StockProps     =   77
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "&Cancel"
         CaptionPos      =   4
         PicturePos      =   3
         ShowFocusRect   =   -1  'True
         Style           =   3
         BorderWidth     =   4
         BeginProperty PanelFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin FACTFRMLib.FactorFrame cmdAdd 
         Height          =   390
         HelpContextID   =   10
         Left            =   165
         TabIndex        =   8
         Top             =   2490
         Width           =   1305
         _Version        =   65536
         _ExtentX        =   2302
         _ExtentY        =   688
         _StockProps     =   77
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   0   'False
         Caption         =   "&Add"
         CaptionPos      =   4
         PicturePos      =   3
         ShowFocusRect   =   -1  'True
         Style           =   3
         BorderWidth     =   4
         BeginProperty PanelFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin FACTFRMLib.FactorFrame cmdClear 
         Height          =   390
         HelpContextID   =   20
         Left            =   1657
         TabIndex        =   12
         Top             =   2490
         Width           =   1305
         _Version        =   65536
         _ExtentX        =   2302
         _ExtentY        =   688
         _StockProps     =   77
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "C&lear"
         CaptionPos      =   4
         PicturePos      =   3
         ShowFocusRect   =   -1  'True
         Style           =   3
         BorderWidth     =   4
         BeginProperty PanelFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin FACTFRMLib.FactorFrame cmdRefresh 
         Height          =   390
         HelpContextID   =   14
         Left            =   3149
         TabIndex        =   10
         Top             =   2490
         Width           =   1305
         _Version        =   65536
         _ExtentX        =   2302
         _ExtentY        =   688
         _StockProps     =   77
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "&Refresh"
         CaptionPos      =   4
         PicturePos      =   3
         ShowFocusRect   =   -1  'True
         Style           =   3
         BorderWidth     =   4
         BeginProperty PanelFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin FACTFRMLib.FactorFrame efraTabBack 
      Height          =   5955
      Left            =   0
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   0
      Width           =   10155
      _Version        =   65536
      _ExtentX        =   17912
      _ExtentY        =   10504
      _StockProps     =   77
      BackColor       =   8388608
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelOuter      =   0
      BeginProperty PanelFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin FACTFRMLib.FactorFrame cmdCancelBtn 
         Height          =   390
         HelpContextID   =   15
         Left            =   8700
         TabIndex        =   2
         Top             =   5430
         Width           =   1305
         _Version        =   65536
         _ExtentX        =   2302
         _ExtentY        =   688
         _StockProps     =   77
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "&Cancel"
         CaptionPos      =   4
         PicturePos      =   3
         ShowFocusRect   =   -1  'True
         Style           =   3
         BorderWidth     =   4
         BeginProperty PanelFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin FACTFRMLib.FactorFrame cmdOkBtn 
         Height          =   390
         HelpContextID   =   16
         Left            =   7200
         TabIndex        =   1
         Top             =   5430
         Width           =   1305
         _Version        =   65536
         _ExtentX        =   2302
         _ExtentY        =   688
         _StockProps     =   77
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   0   'False
         Caption         =   "O&K"
         CaptionPos      =   4
         PicturePos      =   3
         ShowFocusRect   =   -1  'True
         Style           =   3
         BorderWidth     =   4
         BeginProperty PanelFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin FACTFRMLib.FactorFrame efra0TabBase 
         Height          =   5175
         Left            =   120
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   120
         Width           =   9945
         _Version        =   65536
         _ExtentX        =   17542
         _ExtentY        =   9128
         _StockProps     =   77
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   5
         BeginProperty PanelFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin DBTrueGrid.TDBGrid tblCriteria 
            Height          =   5040
            HelpContextID   =   1003
            Left            =   60
            OleObjectBlob   =   "clsCRITERIA.frx":02F4
            TabIndex        =   0
            Top             =   60
            Width           =   9810
         End
      End
   End
   Begin VB.Data datComboDropDown 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   384
      Left            =   432
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1764
      Visible         =   0   'False
      Width           =   1920
   End
   Begin DBTrueGrid.TDBGrid tblComboDropdown 
      Bindings        =   "clsCRITERIA.frx":192F
      Height          =   2484
      Left            =   0
      OleObjectBlob   =   "clsCRITERIA.frx":194E
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   0
      Width           =   3756
   End
End
Attribute VB_Name = "clsCRITERIA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_frmParent As Form
Private tgmCriteria As clsTGSpreadSheet
Private cValidate As cValidateInput

Const TYPE_DATE As String = "1"
Const TYPE_STRING As String = "2"
Const TYPE_DECIMAL As String = "3"
Const TYPE_LONG As String = "4"
Const TYPE_INT As String = "5"
Const TYPE_PHONE As String = "6"

Private Const nCol_ColumnName As Integer = 0
Private Const nCol_Filter As Integer = 1
Private Const nCol_Sort As Integer = 2
Private nCol_OldFilter As Integer
Private nCol_OldSort As Integer
Private nCol_DefaultSort As Integer
Private nCol_FieldName As Integer
Private nCol_FieldType As Integer
Private nCol_NdxTgCombo As String
Private nCol_RegExp As Integer
Private nCol_ComboSQL As String
Private nCol_sPreReqFields As Integer

Private MAX_COL As Integer

Private m_nLastAction As Integer
Private m_sOriginalSQL As String
Private m_sCriteria As String
Private m_sOrderBy As String
Private m_sReturnSQL As String

Private Const t_szCANCEL As String = "Cancel"
Private Const t_szOk As String = "OK"
Private m_sStatusMsg As String

Private m_nCurrRow As Integer

Private sRegExp As String
Private sRegExpOp As String
Private nClickCol As Integer
Private ndxPrevActiveCombo As Integer

'for Debug only
'Private tgcDropdown As clsComboControl
Private tgcDropdown As Object
'
Private Const nDB_LOCAL As Integer = 0
Private Const nDB_REMOTE As Integer = 1
'

Public Sub Setup(frmParent As Form, strSQL As String, Optional bForceSetup As Boolean = False)
    If tgmCriteria Is Nothing Then
        Set m_frmParent = frmParent
        subSetupGrid
        subSetupCombos  'Set up the combo dropdowns
        subSetupValidationSub
    End If
    
    If bForceSetup Then
        tgmCriteria.ClearData
    End If
    
    'setup already been done, exit
    If tgmCriteria.RowCount > 0 Then
        Exit Sub
    End If
    
    MainSQL = strSQL
    
    'build temp table for operator dropdown list
    'build temp table for sort option dropdown list
End Sub

'returns false if error
Public Function AddColumn(ByVal sColName As String, ByVal sFieldName As String, _
                        ByVal sFieldType As String, ByVal sDropdownSQL As String, _
                        Optional ByVal sRegExp As String = "", _
                        Optional ByVal sDefaultFilter As String = "", _
                        Optional ByVal sDefaultSort As String = "", _
                        Optional ByVal sPreRequisiteFields As String = "") As Boolean
    
    On Error GoTo errTrap
    
    sColName = fnCStr(sColName)
    sFieldName = fnCStr(sFieldName)
    sFieldType = fnCStr(sFieldType)
    sDropdownSQL = fnCStr(sDropdownSQL)
    
    If sColName = "" Or sFieldName = "" Or sFieldType = "" Or sDropdownSQL = "" Then
        MsgBox "Column Name, Field Name, sFieldType and Dropdown SQL are required.", vbCritical
        Exit Function
    End If
    
    Dim i As Integer
    Dim ndxPreReqField As Integer
    
    ndxPreReqField = -1
    
    For i = 0 To tgmCriteria.RowCount - 1
        If tgmCriteria.CellValue(nCol_FieldName, i) = sFieldName Then
            MsgBox "Field has already been defined for Col: " + tgmCriteria.CellValue(nCol_ColumnName, i) + ".", vbCritical
            Exit Function
        End If
    Next
    
    'build the row before add into the grid
    ReDim aryData(MAX_COL, 0)
    
    aryData(nCol_ColumnName, 0) = sColName
    aryData(nCol_Filter, 0) = sDefaultFilter
    aryData(nCol_Sort, 0) = sDefaultSort
    aryData(nCol_OldFilter, 0) = sDefaultFilter
    aryData(nCol_OldSort, 0) = sDefaultSort
    aryData(nCol_DefaultSort, 0) = sDefaultSort
    aryData(nCol_FieldName, 0) = sFieldName
    aryData(nCol_FieldType, 0) = sFieldType
    aryData(nCol_RegExp, 0) = sRegExp
    aryData(nCol_ComboSQL, 0) = sDropdownSQL
    aryData(nCol_sPreReqFields, 0) = sPreRequisiteFields
    
    'create a tg combo the column
    Dim nType As Integer
    Dim aryTemp() As String
    
    If Not tgcDropdown Is Nothing Then
        With tgcDropdown
            aryData(nCol_NdxTgCombo, 0) = .AddCombo(sDropdownSQL)
            
            Select Case Left(sFieldType, 1)
            Case TYPE_DATE
                nType = .SQL_DATE_TYPE
            Case TYPE_STRING
                aryTemp = Split(sFieldType, "-")
                
                If aryTemp(1) = "" Then
                    nType = .SQL_STRING_TYPE
                Else
                    nType = .SQL_STRING_TYPE(Val(aryTemp(1)))
                End If
            Case TYPE_DECIMAL
                aryTemp = Split(sFieldType, "-")
                
                If aryTemp(1) = "" Then
                    nType = .SQL_DECIMAL_TYPE
                Else
                    aryTemp = Split(sFieldType, ",")
                    nType = .SQL_DECIMAL_TYPE(Val(aryTemp(0)), Val(aryTemp(1)))
                End If
            Case TYPE_LONG
                nType = .SQL_LONG_TYPE
            Case TYPE_INT
                nType = .SQL_INT_TYPE
            Case TYPE_PHONE
                nType = .SQL_PHONE_TYPE
            Case Else  'regular expression
                MsgBox "Field Type is not valid.", vbCritical
                Exit Function
            End Select
            
            .AddComboBox txtValue, cmdValue, sFieldName, nType
        End With
    End If
    
    'add entry into the grid
    tgmCriteria.AllowAddNew = False
    tgmCriteria.FillWithArray aryData, False
    tgmCriteria.AllowAddNew = False
    
    AddColumn = True
    
    Exit Function
    
errTrap:
    MsgBox "Error in AddColumn()." + vbCrLf + vbCrLf + "Err Code: " & Err.Number & ", Err Desc: " & Err.Description
End Function

'the command button (e.g. cmdQuery) that activate this popup form needs to be disabled.
'otherwise it will take two clicks to cancel/hide this form.
Public Sub ShowPopUp(cmdButton As Control)
    m_sStatusMsg = m_frmParent.ffraStatusbar.Caption
    
    If Not cmdButton Is Nothing Then
        cmdButton.Enabled = False
        DoEvents
    End If
    
    Me.Show vbModal
    Screen.MousePointer = vbDefault

    If Not cmdButton Is Nothing Then
        cmdButton.Enabled = True
        DoEvents
        subSetFocus cmdButton
    End If
End Sub

Property Let MainSQL(strSQL As String)
    m_sOriginalSQL = strSQL
End Property

Property Let ReturnSQL(strSQL As String)
    m_sReturnSQL = strSQL
End Property

Public Sub ClearFilter(Optional sColName As String = "")
    Dim i As Integer
    
    For i = 0 To tgmCriteria.RowCount - 1
        If sColName = "" Or sColName = tgmCriteria.CellValue(nCol_ColumnName, i) Then
            tgmCriteria.CellValue(nCol_Filter, i) = ""
            tgmCriteria.CellValue(nCol_OldFilter, i) = ""
            tgmCriteria.CellValue(nCol_Sort, i) = tgmCriteria.CellValue(nCol_DefaultSort, i)
            tgmCriteria.CellValue(nCol_OldSort, i) = tgmCriteria.CellValue(nCol_DefaultSort, i)
            
            If sColName <> "" Then
                Exit For
            End If
        End If
    Next
    
    If tgmCriteria.RowCount > 0 Then
        tblCriteria.Bookmark = 0
    End If

    m_sCriteria = ""
    m_sOrderBy = ""
    m_sReturnSQL = ""
End Sub

Property Get LastAction() As Integer
    LastAction = m_nLastAction
End Property

Public Function Criteria() As String
    Criteria = m_sCriteria
End Function

Property Get OrderBy() As String
    OrderBy = m_sOrderBy
End Property

Property Get DataChanged() As Boolean
    Dim i As Integer
    DataChanged = False
    For i = 0 To tgmCriteria.RowCount - 1
        If fnCStr(tgmCriteria.CellValue(nCol_Filter, i)) <> fnCStr(tgmCriteria.CellValue(nCol_OldFilter, i)) Or _
            fnCStr(tgmCriteria.CellValue(nCol_Sort, i)) <> fnCStr(tgmCriteria.CellValue(nCol_OldSort, i)) Then
            DataChanged = True
            Exit Function
        End If
    Next i
End Property

Property Get MainSQL() As String
    MainSQL = fnCStr(m_sOriginalSQL)
End Property

'this function will return the original SQL + selection criteria
Property Get ReturnSQL() As String
    ReturnSQL = m_sReturnSQL
End Property

'more public functions
Property Get SQL_DATE_TYPE() As String
    SQL_DATE_TYPE = TYPE_DATE
End Property

Public Function SQL_STRING_TYPE(Optional vLen As Variant) As String
    If Not IsMissing(vLen) Then
        If Val(vLen) <= 0 Or Val(vLen) > 32767 Then MsgBox "String Length is not valid."
        
        SQL_STRING_TYPE = TYPE_STRING + "-" & vLen
    Else
        SQL_STRING_TYPE = TYPE_STRING + "-"
    End If
End Function

Public Function SQL_DECIMAL_TYPE(Optional ByVal vWhole As Variant, _
                                 Optional ByVal vDecimal As Variant) As String
    
    Dim sWhole As String
    Dim sDecimal As String
    
    
    If Not IsMissing(vWhole) Then
        sWhole = fnCStr(vWhole)
        If Val(sWhole) <= 0 Then
            sWhole = ""
        End If
    End If
    If Not IsMissing(vDecimal) Then
        sDecimal = fnCStr(vDecimal)
    End If
    
    SQL_DECIMAL_TYPE = TYPE_DECIMAL + "-" & sWhole & "," & sDecimal
End Function

Property Get SQL_LONG_TYPE() As String
    SQL_LONG_TYPE = TYPE_LONG
End Property

Property Get SQL_INT_TYPE() As String
    SQL_INT_TYPE = TYPE_INT
End Property

Property Get SQL_PHONE_TYPE() As String
    SQL_PHONE_TYPE = TYPE_PHONE
End Property

'Public Static Function
'aryFieldPrefixTable contains (sFieldPrefix As String, sTable As String) pairs
Public Function PrefixTableName(ByVal strSQL As String, ParamArray aryFieldPrefixTable() As Variant) As String
    Dim i As Integer
    Dim j As Integer
    Dim sFieldPrefix As String
    Dim sTable As String
    Dim sTemp As String
    
    If fnCStr(strSQL) = "" Then
        Exit Function
    End If
    
    i = -1
    
    On Error Resume Next
    i = UBound(aryFieldPrefixTable)
    On Error GoTo 0
    
    If i < 0 Then
        MsgBox "Field Prefix and Table Name is missing.", vbExclamation
        Exit Function
    End If
    
    If (i + 1) Mod 2 <> 0 Then
        MsgBox "Field Prefix and Table Name mismatch.", vbExclamation
        Exit Function
    End If
    
    j = -1
    sTemp = strSQL
    
    For i = 0 To ((i + 1) / 2) - 1
        j = j + 1
        sFieldPrefix = aryFieldPrefixTable(j)
        j = j + 1
        sTable = aryFieldPrefixTable(j)
    
        sTemp = Replace(sTemp, sFieldPrefix, sTable + "." + sFieldPrefix)
    Next i
    
    PrefixTableName = sTemp
End Function

Private Sub cmdRefresh_Click()
    txtFilter.Text = tgmCriteria.CellValue(nCol_Filter, tgmCriteria.GetCurrentRowNumber)
    cmdClear.Enabled = txtFilter.Text <> ""
    cmdFilterOK.Enabled = False
    cmdRefresh.Enabled = False
    txtOp = ""
    txtValue = ""
    subSetFocus txtOp
End Sub

'Main Form controls Event functions
Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
    cmdOkBtn.Enabled = True
    tgmCriteria.ResetFlags
    subSetFocus tblCriteria
End Sub

Private Sub Form_Load()
    tfnDisableFormSystemClose Me
    tfnCenterForm Me
    
    m_sReturnSQL = ""
End Sub

Private Sub cmdCancelBtn_GotFocus()
    m_frmParent.tfnSetStatusBarMessage t_szCANCEL
End Sub

Private Sub cmdOkBtn_GotFocus()
    m_frmParent.tfnSetStatusBarMessage t_szOk
End Sub

Private Sub cmdAdd_GotFocus()
    m_frmParent.tfnSetStatusBarMessage "Add"
End Sub

Private Sub cmdClear_GotFocus()
    m_frmParent.tfnSetStatusBarMessage "Clear"
End Sub

Private Sub cmdFilterCancel_GotFocus()
    m_frmParent.tfnSetStatusBarMessage t_szCANCEL
End Sub

Private Sub cmdFilterOK_GotFocus()
    m_frmParent.tfnSetStatusBarMessage t_szOk
End Sub

Private Sub cmdCancelBtn_Click()
    Screen.MousePointer = vbHourglass
    m_frmParent.tfnSetStatusBarMessage ""
    
    If DataChanged() Then
        If Not tfnCancelExit(t_szCANCEL_MESSAGE) Then
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
    End If

    'put old filter value back
    Dim i As Integer
    
    For i = 0 To tgmCriteria.RowCount - 1
        tgmCriteria.CellValue(nCol_Filter, i) = tgmCriteria.CellValue(nCol_OldFilter, i)
        tgmCriteria.CellValue(nCol_Sort, i) = tgmCriteria.CellValue(nCol_OldSort, i)
    Next
    
    m_nLastAction = vbCancel
    m_frmParent.tfnSetStatusBarMessage m_sStatusMsg
    Screen.MousePointer = vbDefault
    Me.Hide
End Sub

Private Sub cmdOkBtn_Click()
    Screen.MousePointer = vbHourglass
    m_frmParent.tfnSetStatusBarMessage ""
    
    'save the filter
    If DataChanged() Then
        Dim i As Integer
        
        For i = 0 To tgmCriteria.RowCount - 1
            tgmCriteria.CellValue(nCol_OldFilter, i) = tgmCriteria.CellValue(nCol_Filter, i)
            tgmCriteria.CellValue(nCol_OldSort, i) = tgmCriteria.CellValue(nCol_Sort, i)
        Next
    End If
    
    subBuildCriteria m_sCriteria, m_sOrderBy
    
    m_sReturnSQL = fnCStr(m_sOriginalSQL)
    
    If m_sCriteria <> "" Then
        If InStr(UCase(m_sReturnSQL), " WHERE ") > 0 Then
            m_sReturnSQL = m_sReturnSQL + " AND "
        Else
            m_sReturnSQL = m_sReturnSQL + " WHERE "
        End If
        
        m_sReturnSQL = m_sReturnSQL + m_sCriteria
    End If
    
    If m_sOrderBy <> "" Then
        m_sReturnSQL = m_sReturnSQL + " " + m_sOrderBy
    End If
    
    m_nLastAction = vbOK
    m_frmParent.tfnSetStatusBarMessage m_sStatusMsg
    Screen.MousePointer = vbDefault
    Me.Hide
End Sub

Private Sub cmdAdd_Click()
    'make filter string
    If fnRefreshFilterString() Then
        If tgmCriteria.CellValue(nCol_Filter, tgmCriteria.GetCurrentRowNumber) <> "" Then
            cmdRefresh.Enabled = True
        End If
    
        txtOp = ""
        txtValue = ""
        cValidate.ResetFlags
        
        cmdClear.Enabled = txtFilter.Text <> ""
        cmdFilterOK.Enabled = fnFilterChanged()
        cmdAdd.Enabled = False
        
        subSetFocus txtOp
    End If
End Sub

Private Sub cmdClear_Click()
    txtFilter.Text = ""
    If tgmCriteria.CellValue(nCol_Filter, tgmCriteria.GetCurrentRowNumber) <> "" Then
        cmdRefresh.Enabled = True
    End If
    
    cmdFilterOK.Enabled = fnFilterChanged()
    txtOp = ""
    txtValue = ""
    subSetFocus txtOp
    cmdClear.Enabled = False
End Sub

Private Sub cmdFilterCancel_Click()
    If fnFilterChanged() Then
        If Not tfnCancelExit(t_szCANCEL_MESSAGE) Then
            Exit Sub
        End If
    End If
    
    subHideFilterPopup
End Sub

Private Sub cmdFilterOK_Click()
    If fnFilterChanged() Then
        tgmCriteria.CellValue(nCol_Filter, tgmCriteria.GetCurrentRowNumber) = txtFilter.Text
    End If
    subSetGridButtonStatus
    subHideFilterPopup
End Sub

Private Function fnRefreshFilterString() As Boolean
    Dim sOP As String
    Dim sValue As String
    Dim aryTemp() As String
    Dim i As Integer
    Dim sNewOp As String
    
    sOP = fnCStr(txtOp)
    
    If sOP = "" Then
        Exit Function
    End If
    
    If fnValidFilter() <> "" Then
        txtFilter.Text = ""
    End If
    
    sValue = fnCStr(txtValue)
    
    If sValue <> "" Then
        Select Case Left(tgmCriteria.CellValue(nCol_FieldType, m_nCurrRow), 1)
        Case TYPE_STRING
            sValue = tfnSQLString(sValue)
        Case TYPE_DATE
            sValue = tfnDateString(sValue, True)
        Case TYPE_DECIMAL, TYPE_LONG, TYPE_INT
            sValue = sValue
        Case TYPE_PHONE
            sValue = tfnSQLString(sValue)
        End Select
    Else
        sValue = "NULL"
    End If
    
    sNewOp = ""

    If fnCStr(txtFilter.Text) = "" Then
        txtFilter.Text = sOP + sValue
    Else
        aryTemp = Split(txtFilter.Text, ";")
        
        For i = 0 To UBound(aryTemp)
            If Left(aryTemp(i), Len(sOP)) = sOP Then
                Exit For
            End If
        Next i
        
        If i > UBound(aryTemp) Then
            sNewOp = sOP + sValue
        Else
            'if not already exist in the filter, add it to the filter
            If InStr("," + Mid(aryTemp(i), Len(sOP) + 1) + ",", "," + sValue + ",") <= 0 Then
                aryTemp(i) = aryTemp(i) + "," + sValue
            End If
        End If
        
        sValue = ""
        
        For i = 0 To UBound(aryTemp)
            If sValue <> "" Then
                sValue = sValue + ";"
            End If
            
            sValue = sValue + fnCStr(aryTemp(i))
        Next i
        
        If sNewOp <> "" Then
            sValue = sValue + ";" + sNewOp
        End If
        
        txtFilter.Text = sValue
    End If
    
    fnRefreshFilterString = True
End Function

Private Function fnFilterChanged() As Boolean
    If txtFilter.Text <> fnCStr(tgmCriteria.CellValue(nCol_Filter, m_nCurrRow)) Then
        fnFilterChanged = True
    End If
End Function

Private Sub subSetupGrid()
    Dim myWidth As Variant
    Dim myField As Variant
    Dim myColCaption As Variant
    Dim myAlignment As Variant
    Dim nWidth As Integer
    Dim i As Integer
    
    With tblCriteria
        While .Columns.Count > 0
            .Columns.Remove 0
        Wend
        
        myColCaption = Array("Column Name", "Filter", "Sort Option")
        myWidth = Array(0.202, 0.64, 0.148)
        myField = Array("", "", "")
        myAlignment = Array(0, 0, 0)
    
        nWidth = .Width - 252
        
        For i = 0 To UBound(myWidth)
            .Columns.Add i
            
            With .Columns(i)
                .Caption = myColCaption(i)
                .Width = myWidth(i) * nWidth
                .Alignment = myAlignment(i)
                .HeadAlignment = 2
                .DataField = myField(i)
                .Visible = True
            End With
        Next i
        
        .ExtendRightColumn = True
    End With

    Set tgmCriteria = New clsTGSpreadSheet
    
    With tgmCriteria
        Set .Form = Me
        Set .StatusBar = m_frmParent.ffraStatusbar
        Set .Table = tblCriteria
        'Set .engFactor = t_engFactor
        .AutoRefreshCellValue = True
        .SetupTable False
        .AllowAddNew = False
        .AddEditColumn nCol_Sort, "Enter Sort Option: A for Ascending or D for Descending", "^([AD])$"
        nCol_OldFilter = .AddHiddenField("OldFilter")
        nCol_OldSort = .AddHiddenField("OldSort")
        nCol_DefaultSort = .AddHiddenField("DefaultSort")
        nCol_FieldName = .AddHiddenField("FieldName")
        nCol_FieldType = .AddHiddenField("FiledType")
        nCol_RegExp = .AddHiddenField("RegExp")
        nCol_NdxTgCombo = .AddHiddenField("NdxTgCombo")
        nCol_ComboSQL = .AddHiddenField("ComboSQL")
        nCol_sPreReqFields = .AddHiddenField("NdxPreReqField")
        
        MAX_COL = nCol_sPreReqFields
        .ClearData
    End With
End Sub

'sBeginWith - WHERE, AND or ""
Private Sub subBuildCriteria(ByRef sRetCriteria As String, _
                             ByRef sRetOrderBy As String, _
                             Optional nRow As Integer = -1)
    Const SUB_NAME As String = "subBuildCriteria"
    
    Dim sCriteria As String
    Dim sOrderBy As String
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim aryTemp() As String
    Dim aryTemp2() As String
    Dim sColumnName As String
    Dim sOperator As String
    Dim sValue As String
    Dim sTemp As String
    
    sRetCriteria = ""
    sRetOrderBy = ""
    sCriteria = ""
    sOrderBy = ""
    
    For i = 0 To tgmCriteria.RowCount - 1
        If nRow < 0 Or (nRow >= 0 And nRow = i) Then
            sColumnName = fnCStr(tgmCriteria.CellValue(nCol_FieldName, i))
            sTemp = ""
            
            If fnCStr(tgmCriteria.CellValue(nCol_Filter, i)) <> "" Then
                aryTemp = Split(fnCStr(tgmCriteria.CellValue(nCol_Filter, i)), ";")
                
                For j = 0 To UBound(aryTemp)
                    sOperator = Left(fnCStr(aryTemp(j)), 1)
                    sValue = Mid(fnCStr(aryTemp(j)), 2)
                    
                    'Here we are checking it is < or <>
                    If Left(sValue, 1) = ">" Then
                        sOperator = "<>"
                        sValue = Mid(sValue, 2)
                    End If
                    
                    aryTemp2 = Split(sValue, ",")
                    
                    For k = 0 To UBound(aryTemp2)
                        If fnCStr(aryTemp2(k)) <> "" Then
                            If sTemp <> "" Then
                                sTemp = sTemp & " OR "
                            End If
                            
                            If UCase(aryTemp2(k)) = "NULL" Then
                                sTemp = sTemp & sColumnName & " IS" & IIf(sOperator = "=", "", " NOT") & " NULL"
                            Else
                                sTemp = sTemp & sColumnName & " " & sOperator & " " & fnCStr(aryTemp2(k))
                            End If
                        End If
                    Next k
                Next j
            End If
            
            If sTemp <> "" Then
                If sCriteria <> "" Then
                    sCriteria = sCriteria & " AND "
                End If
                
                sCriteria = sCriteria + "(" + sTemp + ")"
            End If
            
            If nRow >= 0 Then
                Exit For
            End If
            
            If fnCStr(tgmCriteria.CellValue(nCol_Sort, i)) <> "" Then
                If sOrderBy = "" Then
                    sOrderBy = "ORDER BY " & sColumnName
                Else
                    sOrderBy = sOrderBy & ", " & sColumnName
                End If
                
                If fnCStr(tgmCriteria.CellValue(nCol_Sort, i)) = "D" Then
                    sOrderBy = sOrderBy & " DESC"
                End If
            End If
        End If
    Next i
    
    If nRow < 0 Then
        sRetOrderBy = sOrderBy
    End If
    
    sRetCriteria = sCriteria
End Sub

Private Sub Form_Unload(CANCEL As Integer)
    Set cValidate = Nothing
    Set tgmCriteria = Nothing
End Sub

Private Sub tblCriteria_AfterColEdit(ByVal ColIndex As Integer)
    tgmCriteria.AfterColEdit ColIndex
    subSetGridButtonStatus
End Sub

Private Sub tblCriteria_BeforeColEdit(ByVal ColIndex As Integer, ByVal KeyAscii As Integer, CANCEL As Integer)
    cmdOkBtn.Enabled = False
    tgmCriteria.BeforeColEdit ColIndex, KeyAscii, CANCEL
End Sub

Private Sub tblCriteria_Change()
    cmdOkBtn.Enabled = False
    tgmCriteria.Change
End Sub

Private Sub tblCriteria_FirstRowChange()
    On Error Resume Next
    If tgmCriteria.RowCount < 1 Then
        Exit Sub
    End If
    tgmCriteria.FirstRowChange
End Sub

Private Sub tblCriteria_GotFocus()
    tgmCriteria.GotFocus
    If tblCriteria.Col <> nCol_Sort Then
        m_frmParent.tfnSetStatusBarMessage "Click on 'Filter' column to enter/change selection criteria"
    End If
    subSetGridButtonStatus
End Sub

Private Sub tblCriteria_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If Not tblCriteria.EditActive Then
            If tgmCriteria.IsEndOfTable() Then
                If cmdOkBtn.Enabled Then
                    subSetFocus cmdOkBtn
                    Exit Sub
                End If
            End If
        End If
    End If
    
    tgmCriteria.KeyDown KeyCode, Shift
End Sub

Private Sub tblCriteria_KeyPress(KeyAscii As Integer)
    If tblCriteria.Col = nCol_Sort Then
        KeyAscii = fnUcase(KeyAscii)
    End If
    If Not tgmCriteria.Keypress(KeyAscii) Then
       KeyAscii = 0
    End If
End Sub

Private Sub tblCriteria_LostFocus()
    tgmCriteria.LostFocus
End Sub

Private Sub tblCriteria_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Debug.Print "x="; X; ", nClickCol = "; nClickCol
    nClickCol = tblCriteria.ColContaining(X)
End Sub

Private Sub tblCriteria_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    tgmCriteria.RowColChange LastRow, LastCol
    
    If tblCriteria.Col <> nCol_Sort Then
        m_frmParent.tfnSetStatusBarMessage "Click on 'Filter' column to enter/change selection criteria"
    End If

    If tgmCriteria.RowCount > 0 Then
        If nClickCol = nCol_Filter Then
            'popup the the filter form
            subShowFilterPopup
        End If
    End If
End Sub

Private Sub tblCriteria_SelChange(CANCEL As Integer)
    CANCEL = True
End Sub

Private Sub tblCriteria_UnboundDeleteRow(Bookmark As Variant)
    tgmCriteria.UnboundDeleteRow Bookmark
End Sub

Private Sub tblCriteria_UnboundAddData(ByVal RowBuf As RowBuffer, NewRowBookmark As Variant)
    tgmCriteria.UnboundAddData RowBuf, NewRowBookmark
End Sub

Private Sub tblCriteria_UnboundReadData(ByVal RowBuf As RowBuffer, StartLocation As Variant, ByVal ReadPriorRows As Boolean)
    If Not tgmCriteria Is Nothing Then
        tgmCriteria.ReadData RowBuf, StartLocation, ReadPriorRows
    End If
End Sub

Private Function fnUcase(ByVal kCode As Integer) As Integer
    fnUcase = Asc(UCase(Chr(kCode)))
End Function

Private Sub subShowFilterPopup()
    Dim i As Integer
    
    If ndxPrevActiveCombo >= 0 Then
        For i = 0 To tgmCriteria.RowCount - 1
            ndxPrevActiveCombo = tfnRound(tgmCriteria.CellValue(nCol_NdxTgCombo, i))
            tgcDropdown.ComboActive(ndxPrevActiveCombo) = False
        Next i
    End If
        
    tblCriteria.Enabled = False
    cmdFilterOK.Enabled = False
    cmdRefresh.Enabled = False
    cmdAdd.Enabled = False
    
    m_nCurrRow = tgmCriteria.GetCurrentRowNumber
    
    lblColumnName = tgmCriteria.CellValue(nCol_ColumnName, m_nCurrRow)
    txtFilter.Text = tgmCriteria.CellValue(nCol_Filter, m_nCurrRow)
    
    sRegExp = tgmCriteria.CellValue(nCol_RegExp, m_nCurrRow)
    
    ndxPrevActiveCombo = tfnRound(tgmCriteria.CellValue(nCol_NdxTgCombo, m_nCurrRow))
    tgcDropdown.ComboActive(ndxPrevActiveCombo) = True
    
    cmdClear.Enabled = txtFilter.Text <> ""
    'To set the Regular Expression for Operation Text Box
    SubSetOpRegExp
    
    nClickCol = 2
    txtOp = ""
    txtValue = ""
    cValidate.ResetFlags
    
    efraTabBack.Enabled = False
    efraFilterPopup.Visible = True
    efraFilterPopup.ZOrder 0
    DoEvents
    subSetFocus txtOp
    nClickCol = -1
End Sub

Private Sub subHideFilterPopup()
    efraTabBack.Enabled = True
    tblCriteria.Enabled = True
    efraFilterPopup.Visible = False
    subSetFocus tblCriteria
    nClickCol = -1
End Sub

Private Function fnCStr(v) As String
    fnCStr = Trim(v & "")
End Function

'TG Combo Dropdown functions
Private Sub subSetupCombos()
    #If ProtoType Then
        #If NO_OLECOMBO Then
            Exit Sub
        #End If
    #End If
    
    On Error GoTo errTrap
    
    'Setup the combos
    ' Set the general properties of the combos
    If tgcDropdown Is Nothing Then
        'for Debug only
        'Set tgcDropdown = New clsComboControl
        Set tgcDropdown = CreateObject(t_szOLECOMBO)
        
        subCreateTmpOperator
        
        If Not tgcDropdown Is Nothing Then
            With tgcDropdown
                Set .DBEngine = t_engFactor
                Set .Form = Me
                Set .DataBase = t_dbMainDatabase
                Set .DataLink = datComboDropDown
                Set .Table = tblComboDropDown
                
                'Set .LocalDatabase = dbLocal (Your local database object variable)
                'Add the combos below
                
                .AddCombo ("SELECT op FROM temp_operator")
                .AddComboBox txtOp, cmdOp, "op", .SQL_STRING_TYPE(2)
                .SetOrderingNone txtOp
                
                'more will be added in AddColumn() function - ON THE FLY
            End With
        End If
    End If

    Exit Sub

errTrap:
    tfnErrHandler "Criteria.subSetupCombos"
End Sub

Private Sub tblComboDropdown_Click()
    tgcDropdown.Click tblComboDropDown
End Sub

Private Sub tblComboDropDown_GotFocus()
    tgcDropdown.GotFocus tblComboDropDown
End Sub

Private Sub tblComboDropDown_KeyPress(KeyAscii As Integer)
    tgcDropdown.Keypress tblComboDropDown, KeyAscii
End Sub

Private Sub tblComboDropDown_LostFocus()
    tgcDropdown.LostFocus tblComboDropDown
End Sub

Private Sub tblComboDropDown_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    tgcDropdown.TableMouseUp Y
End Sub

Private Sub tblComboDropDown_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If Not tgcDropdown Is Nothing Then
        tgcDropdown.RowColChange
    End If
End Sub

Private Sub tblComboDropDown_SelChange(CANCEL As Integer)
    tgcDropdown.SelChange CANCEL
End Sub

Private Sub txtFilter_Change()
    If ActiveControl Is txtFilter Then
        If txtFilter.Tag = "" Then
            txtFilter.Tag = "Changed"
        End If
        
        cmdFilterOK.Enabled = False
    End If
End Sub

Private Sub txtFilter_DblClick()
    'unlock the textbox and allow edit
    txtFilter.Locked = False
    m_frmParent.tfnSetStatusBarMessage "Enter or change Filter"
End Sub

Private Sub txtFilter_GotFocus()
    If txtFilter.Locked Then
        m_frmParent.tfnSetStatusBarMessage "Filter"
    Else
        Dim sErrMsg As String
        sErrMsg = fnValidFilter()
        
        If sErrMsg = "" Then
            m_frmParent.tfnSetStatusBarMessage "Enter or change Filter"
        Else
            m_frmParent.tfnSetStatusBarError sErrMsg, True
        End If
    End If
End Sub

Private Sub txtFilter_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If Not txtFilter.Locked Then
            Dim sErrMsg As String
            sErrMsg = fnValidFilter()
            
            If sErrMsg = "" Then
                cmdFilterOK.Enabled = True
                subSetFocus cmdFilterOK
            Else
                If txtFilter.Tag = "Changed" Then
                    m_frmParent.tfnSetStatusBarError sErrMsg
                    subSetFocus txtFilter
                Else
                    SendKeys "{TAB}", True
                End If
            End If
        Else
            subSetFocus txtOp
        End If
        
        txtFilter.Tag = ""
    End If
End Sub

Private Sub txtFilter_LostFocus()
    If ActiveControl Is cmdFilterCancel Then
        txtFilter.Tag = ""
        txtFilter.Locked = True
        Exit Sub
    End If
    
    If Not txtFilter.Locked Then
        Dim sErrMsg As String
        sErrMsg = fnValidFilter()
        
        If sErrMsg = "" Then
            cmdFilterOK.Enabled = True
            txtFilter.Locked = True
        Else
            If txtFilter.Tag = "Changed" Then
                m_frmParent.tfnSetStatusBarError sErrMsg
                subSetFocus txtFilter
            End If
        End If
    End If
    
    txtFilter.Tag = ""
End Sub

Private Sub txtOp_Click()
    tgcDropdown.Click txtOp
End Sub

Private Sub txtOp_Change()
    tfnRegExpControlChange txtOp, sRegExpOp
    
    cValidate.Change txtOp
    
    If ActiveControl Is txtOp Then
        cmdAdd.Enabled = False
    End If
End Sub

Private Sub txtOp_GotFocus()
    tfnRegExpControlGotFocus txtOp, sRegExpOp
    
    cValidate.GotFocus txtOp
    tgcDropdown.GotFocus txtOp
    If tgcDropdown.SingleRecordSelected Then
        subSetFocus txtValue
    End If
End Sub

Private Sub txtOp_KeyPress(KeyAscii As Integer)
    #If ProtoType Then
        Exit Sub
    #End If
    
    If Not tfnRegExpControlKeyPress(txtOp, KeyAscii, sRegExpOp) Then
        KeyAscii = 0
        Exit Sub
    End If
    If KeyAscii = vbKeyReturn Then
        subLoadOpSQL
        
        If txtOp <> "" Then
            KeyAscii = 0
            subSetFocus txtValue
        End If
    End If
    If Not tgcDropdown.Keypress(txtOp, KeyAscii) Then
        If KeyAscii = vbKeyReturn Then
            If tgcDropdown.SingleRecordSelected Then
                subSetButtonStatus
                subSetFocus txtValue
            End If
        End If
        KeyAscii = 0
    Else
        cValidate.Keypress txtOp, KeyAscii
    End If
End Sub

Private Sub txtOp_LostFocus()
    #If ProtoType Then
        Exit Sub
    #End If
    If cValidate.LostFocus(txtOp, cmdOp) Then
        subSetButtonStatus
    End If
End Sub

Private Sub cmdOp_Click()
    #If ProtoType Then
        Exit Sub
    #End If
    
    If txtOp = "" Then
        subLoadOpSQL
        tgcDropdown.Click cmdOp
    Else
        subSetFocus txtValue
    End If
End Sub

Private Sub txtValue_Click()
    tgcDropdown.Click txtValue
End Sub

Private Sub txtValue_Change()
    If sRegExp <> "" Then
        tfnRegExpControlChange txtValue, sRegExp
    End If
    
    cValidate.Change txtValue
    
    If ActiveControl Is txtValue Then
        cmdAdd.Enabled = False
    End If
End Sub

Private Sub txtValue_GotFocus()
    If sRegExp <> "" Then
        tfnRegExpControlGotFocus txtValue, sRegExp
    End If
    
    cValidate.GotFocus txtValue
    tgcDropdown.GotFocus txtValue
    If tgcDropdown.SingleRecordSelected Then
        subSetButtonStatus
        
        If cmdAdd.Enabled Then
            subSetFocus cmdAdd
        End If
    End If
End Sub

Private Sub txtValue_KeyPress(KeyAscii As Integer)
    #If ProtoType Then
        Exit Sub
    #End If
    
    If sRegExp <> "" Then
        If Not tfnRegExpControlKeyPress(txtValue, KeyAscii, sRegExp) Then
            KeyAscii = 0
            Exit Sub
        End If
    End If
            
    If KeyAscii = vbKeyReturn Then
        Screen.MousePointer = vbHourglass
        subFixComboSQL
    End If
    
    If Not tgcDropdown.Keypress(txtValue, KeyAscii) Then
        If KeyAscii = vbKeyReturn Then
            If tgcDropdown.SingleRecordSelected Then
                subSetButtonStatus
                
                If cmdAdd.Enabled Then
                    subSetFocus cmdAdd
                End If
            End If
            
            Screen.MousePointer = vbDefault
        End If
        KeyAscii = 0
    Else
        cValidate.Keypress txtValue, KeyAscii
    End If
End Sub

Private Sub txtValue_LostFocus()
    #If ProtoType Then
        Exit Sub
    #End If
    If cValidate.LostFocus(txtValue, cmdValue) Then
        subSetButtonStatus
        
        If cmdAdd.Enabled Then
            If ActiveControl Is cmdFilterCancel Then
                subSetFocus cmdAdd
            End If
        End If
    End If
End Sub

Private Sub cmdValue_Click()
    #If ProtoType Then
        Exit Sub
    #End If
    
    subFixComboSQL
    tgcDropdown.Click cmdValue
End Sub

Private Sub subFixComboSQL()
    Dim strSQL As String
    Dim aryTemp() As String
    Dim i As Integer
    Dim sCriteria As String
    
    On Error GoTo errTrap
    
    i = tgmCriteria.GetCurrentRowNumber
    aryTemp = Split(tgmCriteria.CellValue(nCol_sPreReqFields, i), ",")
    
    If fnCStr(aryTemp(0)) = "" Then
        Exit Sub
    End If
    
    'strSQL = tgcDropdown.ComboSQL(txtValue)
    strSQL = tgmCriteria.CellValue(nCol_ComboSQL, i)
    
    If strSQL = "" Then
        Exit Sub
    End If
    
    For i = 0 To UBound(aryTemp)
        sCriteria = fnGetCriteriaByFieldName(aryTemp(i))
        
        If sCriteria <> "" Then
            If InStr(UCase(strSQL), " WHERE ") > 0 Then
                strSQL = strSQL + " AND "
            Else
                strSQL = strSQL + " WHERE "
            End If
        
            strSQL = strSQL + sCriteria
        End If
    Next
    
    tgcDropdown.ComboSQL(txtValue) = strSQL
    
    Exit Sub
    
errTrap:
    'do nothing
End Sub

Private Function fnGetCriteriaByFieldName(sField As String) As String
    Dim i As Integer
    Dim sCriteria As String
    
    For i = 0 To tgmCriteria.RowCount - 1
        If tgmCriteria.CellValue(nCol_FieldName, i) = sField Then
            subBuildCriteria sCriteria, "", i
            Exit For
        End If
    Next i
    
    fnGetCriteriaByFieldName = sCriteria
End Function

'Validation functions
Private Sub subSetupValidationSub()

    'Set up the validation class
    Set cValidate = New cValidateInput
    With cValidate
        Set .Form = Me
        Set .StatusBar = m_frmParent.ffraStatusbar
        .GreenMessageDefault = False
        .AddEditBox txtOp, "Enter Operator. E.g. =, <> and etc."
        .AddEditBox txtValue, "Enter Value"
        .MinTabIndex = 0
        .ESCControl = cmdFilterCancel
        .ESCControl = tblComboDropDown
        Set .LastBox = txtValue
        Set .ControlForFocus = efraForFocus
        .SetFirstControls cmdFilterOK, cmdFilterCancel
    End With
End Sub

Private Sub efraForFocus_GotFocus()
    If cValidate Is Nothing Then
        Exit Sub
    End If
    
    cValidate.GotFocus efraForFocus
End Sub

Private Sub subSelectIt(v As Textbox)
    On Error Resume Next
    v.SelStart = 0
    v.SelLength = Len(fnCStr(v.Text))
End Sub

Public Function fnInvalidData(txtBox As Textbox) As Boolean
    #If ProtoType Then
        Exit Function
    #End If
    
    fnInvalidData = False
    Select Case txtBox.TabIndex
        Case txtOp.TabIndex
            fnInvalidData = Not fnValidOperator(txtBox)
        Case txtValue.TabIndex
            fnInvalidData = Not fnValidValue(txtBox)
    End Select
End Function

Private Sub subSetButtonStatus()
    If cValidate.FirstInvalidInput < 0 And fnCStr(txtOp) <> "" Then
        cmdAdd.Enabled = True
    Else
        cmdAdd.Enabled = False
    End If
    
    'work around!!! to cover the red message
    If cValidate.ValidInput(txtValue) Then
        txtValue_GotFocus
    End If
End Sub

Private Function fnAddBeginEndQuotes(sIn As String) As String
    Dim ary() As String
    Dim i As Integer
    Dim sTemp As String
    Dim sRet As String
    Dim bAdd As Boolean
    
    ary = Split(sIn, ",")
    
    For i = 0 To UBound(ary)
        sTemp = fnCStr(ary(i))
        bAdd = False
        
        If Left(sTemp, 1) = "'" And Right(sTemp, 1) = "'" Then
            bAdd = True
        ElseIf sTemp <> "" Then
            sTemp = tfnSQLString(sTemp)
            bAdd = True
        End If
        
        If bAdd Then
            If sRet <> "" Then
                sRet = sRet + ", "
            End If
            
            sRet = sRet + sTemp
        End If
    Next i
    
    fnAddBeginEndQuotes = sRet
End Function

Private Function fnRemvBeginEndQuotes(ByVal sIn As String) As String
    If Left(sIn, 1) = "'" Then
        sIn = Mid(sIn, 2)
    End If
    If Right(sIn, 1) = "'" Then
        sIn = Left(sIn, Len(sIn) - 1)
    End If
    
    fnRemvBeginEndQuotes = sIn
End Function

Private Sub subCreateTmpOperator()
   Const SUB_NAME As String = "subCreateTmpOperator"
   
   Dim strSQL As String
    
    On Error GoTo LetUsGo:
    strSQL = "DROP TABLE temp_operator"
    t_dbMainDatabase.ExecuteSQL strSQL
LetUsGo:
   
   On Error GoTo errCreateTable
   
   strSQL = "CREATE TEMP TABLE temp_operator (op char(2), seq smallint)"
   t_dbMainDatabase.ExecuteSQL strSQL
    
   strSQL = "INSERT INTO temp_operator VALUES ('=', 1)"
   t_dbMainDatabase.ExecuteSQL strSQL
   
   strSQL = "INSERT INTO temp_operator VALUES ('<', 2)"
   t_dbMainDatabase.ExecuteSQL strSQL
   
   strSQL = "INSERT INTO temp_operator VALUES ('>', 3)"
   t_dbMainDatabase.ExecuteSQL strSQL
   
   strSQL = "INSERT INTO temp_operator VALUES ('<>', 4)"
   t_dbMainDatabase.ExecuteSQL strSQL
   
   Exit Sub
errCreateTable:
   tfnErrHandler SUB_NAME, strSQL
End Sub

Private Function GetRecordSet(rsTemp As Recordset, strSQL As String, _
                   Optional nDB As Variant, Optional szCalledFrom As Variant, _
                   Optional bShowErrow As Variant) As Long
    On Error GoTo SQLError
    If IsMissing(nDB) Then
       nDB = nDB_REMOTE
    End If
    Select Case nDB
        Case nDB_LOCAL
            Set rsTemp = dbLocal.OpenRecordset(strSQL, dbOpenSnapshot)
        Case nDB_REMOTE
            Set rsTemp = t_dbMainDatabase.OpenRecordset(strSQL, dbOpenSnapshot, dbSQLPassThrough)
    End Select
    If rsTemp.RecordCount > 0 Then
       rsTemp.MoveLast
       rsTemp.MoveFirst
    End If
    GetRecordSet = rsTemp.RecordCount
    Exit Function
SQLError:
    If IsMissing(szCalledFrom) Then
        szCalledFrom = ""
    End If
    If IsMissing(bShowErrow) Then
        bShowErrow = True
    End If
    GetRecordSet = -1
    tfnErrHandler "GetRecordSet," & szCalledFrom, strSQL, bShowErrow
    On Error GoTo 0
End Function

'return error if any
Private Function fnValidFilter() As String
    Dim sTemp As String
    Dim j As Integer
    Dim k As Integer
    Dim aryTemp() As String
    Dim aryTemp2() As String
    Dim sOP As String
    Dim sValue As String
    Dim sFilter As String
    Dim sNewFilter As String
    Dim sString As String
    
    sFilter = fnCStr(txtFilter.Text)
        
    If fnCStr(sFilter) = "" Then
        'no selection criteria, okay
        txtFilter.Text = sFilter
    Else
        aryTemp = Split(sFilter, ";")
        sNewFilter = ""
        
        For j = 0 To UBound(aryTemp)
            sOP = Left(fnCStr(aryTemp(j)), 1)
            sValue = Mid(fnCStr(aryTemp(j)), 2)
            
            If sOP = "" Then
                fnValidFilter = "Operator is missing"
                Exit Function
            End If
            
            'Here we are checking it is < or <>
            If Left(sValue, 1) = ">" Then
                If sOP <> "<" Then
                    fnValidFilter = "Invalid operator '" + sOP + Left(sValue, 1) + "' encountered"
                    Exit Function
                End If
                
                sOP = "<>"
                sValue = Mid(sValue, 2)
            End If
            
            Select Case sOP
            Case "<", ">", "=", "<>"
                'good
            Case Else
                fnValidFilter = "Invalid operator '" + sOP + "' encountered"
                Exit Function
            End Select
            
            If sOP = "<" Or sOP = ">" Then
                If Left(tgmCriteria.CellValue(nCol_FieldType, m_nCurrRow), 1) = TYPE_STRING Then
                    fnValidFilter = sOP + " is not allow for char type column"
                    Exit Function
                End If
            End If
                
            'reuse variable
            sFilter = ""
            aryTemp2 = Split(sValue, ",")
            
            For k = 0 To UBound(aryTemp2)
                sTemp = fnCStr(aryTemp2(k))
                
                If sTemp = "" Then
                    fnValidFilter = "Value is missing for operator '" + sOP + "'"
                    Exit Function
                Else
                    If UCase(sTemp) = "NULL" Then
                        If sOP = "<" Or sOP = ">" Then
                            fnValidFilter = "NULL is not valid for operator " + sOP
                            Exit Function
                        End If
                        
                        sTemp = "NULL"
                    End If
                    
                    Select Case Left(tgmCriteria.CellValue(nCol_FieldType, m_nCurrRow), 1)
                    Case TYPE_DATE
                        If sTemp <> "NULL" Then
                            sTemp = tfnFormatDate(fnRemvBeginEndQuotes(sTemp))
                            
                            If Not IsValidDate(sTemp) Then
                                fnValidFilter = "Value is not in date format"
                                Exit Function
                            End If
                            
                            sTemp = tfnDateString(sTemp, True)
                        End If
                    Case TYPE_STRING, TYPE_PHONE
                        If sTemp <> "NULL" Then
                            sTemp = fnAddBeginEndQuotes(sTemp)
                        End If
                    Case TYPE_DECIMAL, TYPE_INT, TYPE_LONG
                        If sTemp <> "NULL" Then
                            If Not IsNumeric(sTemp) Then
                                fnValidFilter = "Value entered is not in numeric format"
                                Exit Function
                            End If
                        End If
                    End Select
                
                    If sFilter <> "" Then
                        sFilter = sFilter + ","
                    End If
                    
                    sFilter = sFilter + sTemp
                End If
            Next k
            
            If sNewFilter <> "" Then
                sNewFilter = sNewFilter + ";"
            End If
        
            sNewFilter = sNewFilter + sOP + sFilter
        Next j
        
        txtFilter.Text = sNewFilter
        txtFilter.Tag = ""
    End If
End Function

Private Function fnValidOperator(txtBox As Textbox) As Boolean
    Const FUNC_NAME As String = "fnValidOperator"
    
    Dim strSQL As String
    Dim rsTemp As Recordset
    
    fnValidOperator = False
    
    If fnCStr(txtBox) <> "" Then
        strSQL = "SELECT op FROM temp_operator WHERE op=" & tfnSQLString(fnCStr(txtBox))
        
        Select Case Left(tgmCriteria.CellValue(nCol_FieldType, m_nCurrRow), 1)
        Case TYPE_STRING, TYPE_PHONE
            strSQL = strSQL & " AND op IN ('=','<>')"
        End Select
        
        If GetRecordSet(rsTemp, strSQL, , FUNC_NAME) < 1 Then
            cValidate.SetErrorMessage txtBox, "Invalid Operator"
            Exit Function
        End If
    
        cValidate.ResetFlags txtValue
        fnValidOperator = True
    End If
End Function

Private Function fnValidValue(txtBox As Textbox) As Boolean
    Const FUNC_NAME As String = "fnValidValue"
    
    Dim strSQL As String
    Dim rsTemp As Recordset
    Dim sExtraCriteria As String
    Dim sField As String
    
    fnValidValue = False
    
    If fnCStr(txtBox) = "" Then
        If fnCStr(txtOp) <> "=" And fnCStr(txtOp) <> "<>" Then
            Exit Function
        End If
    Else
        Select Case Left(tgmCriteria.CellValue(nCol_FieldType, m_nCurrRow), 1)
        Case TYPE_DATE
            If Not IsValidDate(txtBox.Text) Then
                cValidate.SetErrorMessage txtBox, "Invalid " & tgmCriteria.CellValue(nCol_ColumnName, m_nCurrRow)
                Exit Function
            Else
                txtBox = tfnFormatDate(txtBox)
            End If
        Case TYPE_STRING
        Case TYPE_DECIMAL, TYPE_INT, TYPE_LONG
            If Not IsNumeric(txtBox) Then
                cValidate.SetErrorMessage txtBox, "Value entered is not in numeric format"
                Exit Function
            End If
        End Select
    End If
    
    strSQL = tgcDropdown.ComboSQL(txtBox)
    
    sExtraCriteria = fnGetCriteria(tgmCriteria.CellValue(nCol_FieldName, m_nCurrRow), txtBox.Text)
    
    If sExtraCriteria <> "" Then
        sField = fnGetFieldByAlias(strSQL, tgmCriteria.CellValue(nCol_FieldName, m_nCurrRow))
        
        If sField <> "" Then
            sExtraCriteria = Replace(sExtraCriteria, tgmCriteria.CellValue(nCol_FieldName, m_nCurrRow), sField)
        End If
        
        If InStr(UCase(strSQL), " WHERE ") > 0 Then
            strSQL = strSQL + " AND "
        Else
            strSQL = strSQL + " WHERE "
        End If
        
        strSQL = strSQL & sExtraCriteria
    End If
    
    If GetRecordSet(rsTemp, strSQL) < 1 Then
        cValidate.SetErrorMessage txtBox, "Invalid " & tgmCriteria.CellValue(nCol_ColumnName, m_nCurrRow)
        Exit Function
    End If
'''
'''    If fnCStr(txtOp) = "" Then
'''        cValidate.SetErrorMessage txtOp, "Enter Operator"
'''        Exit Function
'''    End If
    
    fnValidValue = True
End Function

Private Function fnGetCriteria(sFieldName As String, ByVal sValue As String) As String
    Const FUNC_NAME As String = "fnGetCriteria"
    
    sValue = fnCStr(sValue)
    
    fnGetCriteria = "(" & sFieldName
    
    Select Case Left(tgmCriteria.CellValue(nCol_FieldType, m_nCurrRow), 1)
        Case TYPE_STRING, TYPE_PHONE
            If sValue = "" Then
                fnGetCriteria = fnGetCriteria & " IS NULL" & " OR " & sFieldName & " = ''"
            Else
                fnGetCriteria = fnGetCriteria & " = " & tfnSQLString(sValue)
            End If
        Case TYPE_DATE
            If sValue = "" Then
                fnGetCriteria = fnGetCriteria & " IS NULL"
            Else
                fnGetCriteria = fnGetCriteria & " = " & tfnDateString(sValue, True)
            End If
        Case TYPE_DECIMAL, TYPE_LONG, TYPE_INT
            If sValue = "" Then
                fnGetCriteria = fnGetCriteria & " IS NULL"
            Else
                fnGetCriteria = fnGetCriteria & " = " & txtValue
            End If
        End Select

        fnGetCriteria = fnGetCriteria & ")"
End Function

Private Function fnGetFieldByAlias(strSQL As String, sAlias As String) As String
    Dim nPosi As Integer
    Dim nLen As Integer
    Dim i As Integer
    Dim aryFields() As String
    Dim sTemp As String
    
    nPosi = InStr(UCase(strSQL), "SELECT ")
    
    If nPosi > 0 Then
        sTemp = Trim(Mid(strSQL, nPosi + 7))
        nPosi = InStr(UCase(sTemp), "DISTINCT ")
        
        If nPosi > 0 Then
            sTemp = Trim(Mid(sTemp, nPosi + 9))
        End If
        
        nPosi = InStr(UCase(sTemp), " FROM ")
    
        If nPosi > 0 Then
            sTemp = Trim(Left(sTemp, nPosi - 1))
            aryFields() = Split(sTemp, ",")
            
            For i = 0 To UBound(aryFields)
                sTemp = Trim(aryFields(i))
                nPosi = InStr(UCase(sTemp), " AS ")
                nLen = 4
                
                If nPosi <= 0 Then
                    nPosi = InStr(sTemp, " ")
                    nLen = 1
                End If
        
                If nPosi > 0 Then
                    If Trim(Mid(sTemp, nPosi + nLen)) = sAlias Then
                        fnGetFieldByAlias = Trim(Left(sTemp, nPosi - 1))
                        Exit Function
                    End If
                End If
            Next i
        End If
    End If
End Function

Public Function fnValidCellValue(Table As TDBGrid, ByVal nCol As Integer, ByVal lRow As Long, sText As String) As Boolean
    Const SUB_NAME As String = "fnValidCellValue"
    fnValidCellValue = True         'Assume the value will be valid
    sText = fnCStr(sText)
    Select Case nCol
        Case nCol_Sort
            If sText <> szEMPTY Then
                If sText <> "A" And sText <> "D" Then
                    tgmCriteria.ErrorMessage(nCol) = "Invalid Sort Option"
                    fnValidCellValue = False
                End If
            End If
    End Select
End Function

Private Sub subSetGridButtonStatus()
    If tgmCriteria.ValidData Then
        cmdOkBtn.Enabled = True
    Else
        cmdOkBtn.Enabled = False
    End If
End Sub

Private Sub subLoadOpSQL()

    Dim strSQL As String
    strSQL = "SELECT op, seq FROM temp_operator"
    Select Case Left(tgmCriteria.CellValue(nCol_FieldType, m_nCurrRow), 1)
        Case TYPE_STRING, TYPE_PHONE
            strSQL = strSQL & " WHERE op IN ('=','<>')"
    End Select
    
    strSQL = strSQL & " ORDER BY seq"
    tgcDropdown.ComboSQL(txtOp) = strSQL
End Sub

Private Sub SubSetOpRegExp()
    sRegExpOp = "^((=)|(<)|(>)|(<>))$"
    Select Case Left(tgmCriteria.CellValue(nCol_FieldType, m_nCurrRow), 1)
        Case TYPE_DATE, TYPE_STRING
            sRegExpOp = "^((=)|(<>))$"
    End Select
End Sub

Private Function IsValidDate(ByVal sDate As String) As Boolean
    
    If sDate = "" Then
        Exit Function
    End If
    sDate = tfnFormatDate(sDate)
    If SRegExpMatch(szDatePattern, sDate) <> 0 Then
        Exit Function
    End If
    If Not IsDate(tfnDateString(sDate)) Then
        Exit Function
    End If
    IsValidDate = True
End Function
