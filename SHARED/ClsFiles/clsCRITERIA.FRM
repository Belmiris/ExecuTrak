VERSION 5.00
Object = "{C75015E0-2232-11D3-B440-0060971E99AF}#1.0#0"; "factfrm.ocx"
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
   StartUpPosition =   2  'CenterScreen
   Begin FACTFRMLib.FactorFrame efraFilterPopup 
      Height          =   3030
      Left            =   1365
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   2295
      Visible         =   0   'False
      Width           =   8640
      _Version        =   65536
      _ExtentX        =   15240
      _ExtentY        =   5345
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
         TabIndex        =   10
         Top             =   195
         Width           =   8355
         _Version        =   65536
         _ExtentX        =   14737
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
         Begin VB.TextBox txtDesc 
            Height          =   360
            HelpContextID   =   1004
            Left            =   3585
            TabIndex        =   8
            Top             =   1680
            Width           =   4230
         End
         Begin VB.TextBox txtFilter 
            Height          =   360
            HelpContextID   =   1000
            Left            =   165
            Locked          =   -1  'True
            TabIndex        =   16
            Top             =   990
            Width           =   8025
         End
         Begin VB.TextBox txtOp 
            Height          =   360
            HelpContextID   =   1001
            Left            =   180
            TabIndex        =   4
            Top             =   1680
            Width           =   1065
         End
         Begin VB.TextBox txtValue 
            Height          =   360
            HelpContextID   =   1002
            Left            =   1710
            TabIndex        =   6
            Top             =   1680
            Width           =   1425
         End
         Begin FACTFRMLib.FactorFrame cmdOp 
            Height          =   360
            HelpContextID   =   1001
            Left            =   1260
            TabIndex        =   5
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
            Left            =   3150
            TabIndex        =   7
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
         Begin FACTFRMLib.FactorFrame cmdDesc 
            Height          =   360
            HelpContextID   =   1004
            Left            =   7830
            TabIndex        =   9
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
            Picture         =   "clsCRITERIA.frx":02F4
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
         Begin VB.Label lblDesc 
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Description"
            ForeColor       =   &H00000000&
            Height          =   270
            Left            =   3585
            TabIndex        =   26
            Top             =   1410
            UseMnemonic     =   0   'False
            Width           =   4605
         End
         Begin VB.Label lblColumnName 
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Height          =   360
            Left            =   180
            TabIndex        =   25
            Top             =   315
            UseMnemonic     =   0   'False
            Width           =   8025
         End
         Begin VB.Label Label 
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Column Name"
            ForeColor       =   &H00000000&
            Height          =   270
            Index           =   0
            Left            =   180
            TabIndex        =   24
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
            TabIndex        =   23
            Top             =   1410
            UseMnemonic     =   0   'False
            Width           =   1320
         End
         Begin VB.Label Label 
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Value"
            ForeColor       =   &H00000000&
            Height          =   270
            Index           =   3
            Left            =   1710
            TabIndex        =   22
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
            Index           =   1
            Left            =   180
            TabIndex        =   21
            Top             =   750
            UseMnemonic     =   0   'False
            Width           =   825
         End
      End
      Begin FACTFRMLib.FactorFrame cmdFilterOK 
         Height          =   390
         HelpContextID   =   16
         Left            =   5406
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
         Left            =   7155
         TabIndex        =   14
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
         Left            =   1912
         TabIndex        =   15
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
         Left            =   3659
         TabIndex        =   13
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
      TabIndex        =   17
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
      Begin FACTFRMLib.FactorFrame cmdClearAll 
         Height          =   390
         HelpContextID   =   1005
         Left            =   5700
         TabIndex        =   3
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
         Caption         =   "C&lear Filters"
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
         TabIndex        =   18
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
            OleObjectBlob   =   "clsCRITERIA.frx":04F6
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
      Bindings        =   "clsCRITERIA.frx":1B31
      Height          =   2484
      Left            =   0
      OleObjectBlob   =   "clsCRITERIA.frx":1B50
      TabIndex        =   19
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
Private nCol_NdxTgCombo As Integer
Private nCol_RegExp As Integer
Private nCol_ComboSQL As Integer
Private nCol_CriteriaField As Integer
Private nCol_Operators As Integer
Private nCol_Enabled As Integer
Private nCol_AllowEditSortCol As Integer
Private nCol_XtraColName As Integer
Private nCol_XtraRegExp As Integer
Private nCol_sPreReqFields As Integer

Private MAX_COL As Integer

Private m_nLastAction As Integer
Private m_sOriginalSQL As String
Private m_sCriteria As String
Private m_sOrderBy As String
Private m_sReturnSQL As String
Private m_bOrderByRequired As Boolean

Private Const t_szCANCEL As String = "Cancel"
Private Const t_szOk As String = "OK"
Private m_sStatusMsg As String

Private m_nCurrRow As Integer

Private sRegExp As String
Private sRegExpDesc As String
Private sRegExpOp As String
Private nClickCol As Integer
Private ndxPrevActiveCombo As Integer

'for Debug only
'Private tgcDropdown As clsComboControl
Private tgcDropdown As Object
'
Private Const nDB_LOCAL As Integer = 0
Private Const nDB_REMOTE As Integer = 1

Private sngDefaultValueLeft As Single
Private sngDefaultValueWidth As Single
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
                        Optional ByVal sPreRequisiteFields As String = "", _
                        Optional ByVal sCriteriaField As String = "", _
                        Optional ByVal bAllowEditSortCol As Boolean = True, _
                        Optional ByVal sOperators As String = "") As Boolean
    
    On Error GoTo errTrap
    
    sColName = fnCStr(sColName)
    sFieldName = fnCStr(sFieldName)
    sFieldType = fnCStr(sFieldType)
    sDropdownSQL = fnCStr(sDropdownSQL)
    
    If sColName = "" Or sFieldName = "" Or sFieldType = "" Or sDropdownSQL = "" Then
        MsgBox "Column Name, Field Name, Field Type and Dropdown SQL are required.", vbCritical
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
    aryData(nCol_CriteriaField, 0) = fnCStr(sCriteriaField)
    aryData(nCol_Operators, 0) = fnCStr(sOperators)
    aryData(nCol_Enabled, 0) = True
    aryData(nCol_AllowEditSortCol, 0) = bAllowEditSortCol
    aryData(nCol_sPreReqFields, 0) = sPreRequisiteFields
    
    'create a tg combo the column
    Dim nType As Integer
    Dim aryTemp() As String
    
    If Not tgcDropdown Is Nothing Then
        With tgcDropdown
            aryData(nCol_NdxTgCombo, 0) = .AddCombo(sDropdownSQL, , tgcDropdown.DropUp)
            
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
            
            'get the alias
            sFieldName = fnGetFieldName(sFieldName, "A")
            
            .AddComboBox txtValue, cmdValue, sFieldName, nType, , sFieldName
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

Public Sub AddExtraColumn(ByVal sPrimaryFieldName As String, ByVal sColName As String, _
                          ByVal sFieldName As String, ByVal sFieldType As String, _
                          Optional ByVal sRegExp As String = "", _
                          Optional ByVal sColCaption1 As String = "", _
                          Optional ByVal sColCaption2 As String = "")
    
    On Error GoTo errTrap
    
    Dim i As Integer
    Dim ndxComboID As Integer
    
    sColName = fnCStr(sColName)
    sFieldName = fnCStr(sFieldName)
    sFieldType = fnCStr(sFieldType)
    sRegExp = fnCStr(sRegExp)
    sColCaption1 = fnCStr(sColCaption1)
    sColCaption2 = fnCStr(sColCaption2)
    
    If sPrimaryFieldName = "" Then
        MsgBox "FieldName for Primary Column (e.g. Product Code/Customer Number), is required.", vbCritical
        Exit Sub
    End If
    
    If sColName = "" Or sFieldName = "" Or sFieldType = "" Then
        MsgBox "Column Name, Field Name, and Field Type are required.", vbCritical
        Exit Sub
    End If
    
    With tgmCriteria
        For i = 0 To .RowCount - 1
            If .CellValue(nCol_FieldName, i) = sPrimaryFieldName Then
                ndxComboID = tfnRound(.CellValue(nCol_NdxTgCombo, i))
                
                If ndxComboID <= 0 Then
                    MsgBox "ComboDropdown not setup for Primary Column!", vbCritical
                    Exit Sub
                End If
                
                Exit For
            End If
        Next
        
        If i > .RowCount - 1 Then
            MsgBox "Primary Column (e.g. Product Code/Customer Number) not found!", vbCritical
            Exit Sub
        End If
        
        'save 'desc box' setup
        .CellValue(nCol_XtraColName, i) = sColName
        .CellValue(nCol_XtraRegExp, i) = sRegExp
        
        If sColCaption1 = "" Then
            sColCaption1 = fnCStr(.CellValue(nCol_ColumnName, i))
        End If
        If sColCaption2 = "" Then
            sColCaption2 = sColName
        End If
    End With
    
    'create a tg combo the column
    Dim nType As Integer
    Dim aryTemp() As String
    
    If Not tgcDropdown Is Nothing Then
        With tgcDropdown
            .CurrentCombo = ndxComboID
            
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
                Exit Sub
            End Select
            
            .AddComboBox txtDesc, cmdDesc, sFieldName, nType, , sFieldName
            .SetColumnCaptions sColCaption1, sColCaption2
        End With
    End If
    
    Exit Sub
    
errTrap:
    MsgBox "Error in AddColumn()." + vbCrLf + vbCrLf + "Err Code: " & Err.Number & ", Err Desc: " & Err.Description
End Sub

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
    
    With tgmCriteria
        For i = 0 To .RowCount - 1
            If sColName = "" Or sColName = .CellValue(nCol_ColumnName, i) Then
                .CellValue(nCol_Filter, i) = ""
                .CellValue(nCol_OldFilter, i) = ""
                
                If .CellValue(nCol_Enabled, i) Then
                    .CellValue(nCol_Sort, i) = .CellValue(nCol_DefaultSort, i)
                Else
                    .CellValue(nCol_Filter, i) = "N/A"
                    .CellValue(nCol_Sort, i) = "N/A"
                End If
                
                .CellValue(nCol_OldSort, i) = .CellValue(nCol_Sort, i)
                
                If sColName <> "" Then
                    Exit For
                End If
            End If
        Next
        
        If .RowCount > 0 Then
            tblCriteria.Bookmark = 0
        End If
    End With
    
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
        If fnCStr(tgmCriteria.CellValue(nCol_Filter, i)) <> "N/A" Then
            If fnCStr(tgmCriteria.CellValue(nCol_Filter, i)) <> fnCStr(tgmCriteria.CellValue(nCol_OldFilter, i)) _
                Then
                DataChanged = True
                Exit Function
            End If
        End If
        
        If fnCStr(tgmCriteria.CellValue(nCol_Sort, i)) <> "N/A" Then
            If fnCStr(tgmCriteria.CellValue(nCol_Sort, i)) <> fnCStr(tgmCriteria.CellValue(nCol_OldSort, i)) Then
                DataChanged = True
                Exit Function
            End If
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

Property Let AllowEditSortCol(ByVal sColName As String, ByVal bStatus As Boolean)
    sColName = fnCStr(sColName)
    
    If sColName = "" Then
        MsgBox "Column Name is required.", vbCritical
        Exit Property
    End If
    
    Dim i As Integer
    
    With tgmCriteria
        For i = 0 To .RowCount - 1
            If .CellValue(nCol_ColumnName, i) = sColName Then
                .CellValue(nCol_AllowEditSortCol, i) = bStatus
                Exit Property
            End If
        Next
    End With

    MsgBox "Column Name not found in the Criteria Screen.", vbCritical
End Property

Property Get AllowEditSortCol(ByVal sColName As String) As Boolean
    sColName = fnCStr(sColName)
    
    If sColName = "" Then
        MsgBox "Column Name is required.", vbCritical
        Exit Property
    End If
    
    Dim i As Integer
    
    With tgmCriteria
        For i = 0 To .RowCount - 1
            If .CellValue(nCol_ColumnName, i) = sColName Then
                AllowEditSortCol = .CellValue(nCol_AllowEditSortCol, i)
                Exit Property
            End If
        Next
    End With

    MsgBox "Column Name not found in the Criteria Screen.", vbCritical
End Property

Property Let ColEnabled(ByVal sColName As String, ByVal bStatus As Boolean)
    sColName = fnCStr(sColName)
    
    If sColName = "" Then
        MsgBox "Column Name is required.", vbCritical
        Exit Property
    End If
    
    Dim i As Integer
    
    With tgmCriteria
        For i = 0 To .RowCount - 1
            If .CellValue(nCol_ColumnName, i) = sColName Then
                .CellValue(nCol_Enabled, i) = bStatus
                
                If bStatus Then
                    .CellValue(nCol_Sort, i) = .CellValue(nCol_DefaultSort, i)
                Else
                    .CellValue(nCol_Filter, i) = "N/A"
                    .CellValue(nCol_Sort, i) = "N/A"
                End If
                
                .CellValue(nCol_OldSort, i) = .CellValue(nCol_Sort, i)
                
                Exit Property
            End If
        Next
    End With

    MsgBox "Column Name not found in the Criteria Screen.", vbCritical
End Property

Property Get ColEnabled(ByVal sColName As String) As Boolean
    sColName = fnCStr(sColName)
    
    If sColName = "" Then
        MsgBox "Column Name is required.", vbCritical
        Exit Property
    End If
    
    Dim i As Integer
    
    With tgmCriteria
        For i = 0 To .RowCount - 1
            If .CellValue(nCol_ColumnName, i) = sColName Then
                ColEnabled = .CellValue(nCol_Enabled, i)
                Exit Property
            End If
        Next
    End With

    MsgBox "Column Name not found in the Criteria Screen.", vbCritical
End Property

Property Let Operators(ByVal sColName As String, ByVal sOperators As String)
    sColName = fnCStr(sColName)
    
    If sColName = "" Then
        MsgBox "Column Name is required.", vbCritical
        Exit Property
    End If
    
    Dim i As Integer
    
    With tgmCriteria
        For i = 0 To .RowCount - 1
            If .CellValue(nCol_ColumnName, i) = sColName Then
                .CellValue(nCol_Operators, i) = fnCStr(sOperators)
                Exit Property
            End If
        Next
    End With

    MsgBox "Column Name not found in the Criteria Screen.", vbCritical
End Property

Property Get Operators(ByVal sColName As String) As String
    sColName = fnCStr(sColName)
    
    If sColName = "" Then
        MsgBox "Column Name is required.", vbCritical
        Exit Property
    End If
    
    Dim i As Integer
    
    With tgmCriteria
        For i = 0 To .RowCount - 1
            If .CellValue(nCol_ColumnName, i) = sColName Then
                Operators = .CellValue(nCol_Operators, i)
                Exit Property
            End If
        Next
    End With

    MsgBox "Column Name not found in the Criteria Screen.", vbCritical
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

Property Let OrderByRequired(bYesNo As Boolean)
    m_bOrderByRequired = bYesNo
End Property

Property Get OrderByRequired() As Boolean
    OrderByRequired = m_bOrderByRequired
End Property

'returns A for Ascending or D for Descending, "" for no sort
Property Get SortOrder(vColIndexColName As Variant) As String
    Dim sColName As String
    Dim lRow As Long
    
    With tgmCriteria
        If VarType(vColIndexColName) = vbString Then
            sColName = fnCStr(vColIndexColName)
            
            If sColName = "" Then
                MsgBox "Column Name is required.", vbCritical
                Exit Property
            End If
            
            For lRow = 0 To .RowCount - 1
                If .CellValue(nCol_ColumnName, lRow) = sColName Then
                    SortOrder = fnCStr(.CellValue(nCol_Sort, lRow))
                    
                    If SortOrder = "" And m_bOrderByRequired Then
                        SortOrder = fnCStr(.CellValue(nCol_DefaultSort, lRow))
                    End If
                    
                    Exit Property
                End If
            Next
            
            If lRow > .RowCount - 1 Then
                MsgBox "Column Name '" + sColName + "' not found.", vbCritical
                Exit Property
            End If
        Else
            lRow = vColIndexColName
            
            If lRow < 0 Or lRow > .RowCount - 1 Then
                MsgBox "Column Index '" & lRow & "' is not valid.", vbCritical
                Exit Property
            End If
            
            SortOrder = fnCStr(.CellValue(nCol_Sort, lRow))
            
            If SortOrder = "" And m_bOrderByRequired Then
                SortOrder = fnCStr(.CellValue(nCol_DefaultSort, lRow))
            End If
        End If
    End With
End Property

Private Sub cmdClearAll_Click()
    If Not tfnCancelExit("Are you sure you want to clear the Filters and Sort Options?") Then
        Screen.MousePointer = vbDefault
        Exit Sub
    End If

    Dim i As Integer
    
    With tgmCriteria
        For i = 0 To .RowCount - 1
            .CellValue(nCol_Filter, i) = ""
            
            If .CellValue(nCol_Enabled, i) Then
                .CellValue(nCol_Sort, i) = .CellValue(nCol_DefaultSort, i)
            Else
                .CellValue(nCol_Filter, i) = "N/A"
                .CellValue(nCol_Sort, i) = "N/A"
            End If
        Next
    End With
End Sub

Private Sub cmdClearAll_GotFocus()
    m_frmParent.tfnSetStatusBarMessage "Clear All Filters"
End Sub

Private Sub cmdRefresh_Click()
    txtFilter.Text = tgmCriteria.CellValue(nCol_Filter, tgmCriteria.GetCurrentRowNumber)
    cmdClear.Enabled = txtFilter.Text <> ""
    cmdFilterOK.Enabled = False
    cmdRefresh.Enabled = False
    txtOp = ""
    txtValue = ""
    txtDesc = ""
    cValidate.ResetFlags

    subSetFocus txtOp
End Sub

'Main Form controls Event functions
Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
    subSetClearAllBtnStatus
    nClickCol = -1
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
    
    'the 'order by' in MainSQL will override the one bult from the criteria screen
    m_sReturnSQL = fnCStr(m_sOriginalSQL)
    
    'take out the 'order by' from MainSQL if any
    Dim nPosi As Integer
    
    nPosi = InStr(UCase(m_sReturnSQL), "ORDER BY")
    
    If nPosi > 0 Then
        m_sOrderBy = Trim(Mid(m_sReturnSQL, nPosi))
        m_sReturnSQL = Trim(Left(m_sReturnSQL, nPosi - 1))
    End If
        
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
    Dim lRow As Long
    
    If Not fnValidValue(txtValue) Then
        subSetFocus txtValue
        cmdAdd.Enabled = False
        Exit Sub
    End If
    
    'make filter string
    lRow = tgmCriteria.GetCurrentRowNumber
    
    If fnRefreshFilterString() Then
        If tgmCriteria.CellValue(nCol_Filter, lRow) <> "" And tgmCriteria.CellValue(nCol_Filter, lRow) <> "N/A" Then
            cmdRefresh.Enabled = True
        End If
    
        txtOp = ""
        txtValue = ""
        txtDesc = ""
        cValidate.ResetFlags
        
        cmdClear.Enabled = txtFilter.Text <> ""
        cmdFilterOK.Enabled = fnFilterChanged()
        cmdAdd.Enabled = False
        
        'subSetFocus txtOp
        subSetFocus cmdFilterOK
    End If
End Sub

Private Sub cmdClear_Click()
    If Not tfnCancelExit("Are you sure you want to clear the Filter value?") Then
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    
    Dim lRow As Long
    
    lRow = tgmCriteria.GetCurrentRowNumber
    
    txtFilter.Text = ""
    If tgmCriteria.CellValue(nCol_Filter, lRow) <> "" And tgmCriteria.CellValue(nCol_Filter, lRow) <> "N/A" Then
        cmdRefresh.Enabled = True
    End If
    
    cmdFilterOK.Enabled = fnFilterChanged()
    txtOp = ""
    txtValue = ""
    txtDesc = ""
    cValidate.ResetFlags
    
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
    Dim sTemp As String
    
    sOP = fnCStr(txtOp)
    
    If sOP = "" Then
        Exit Function
    End If
    
    If fnValidFilter() <> "" Then
        txtFilter.Text = ""
    End If
    
    sValue = fnCStr(txtValue)
    
    If sValue <> "" Then
        If sOP = "BETWEEN" Then
            sValue = fnFixBetweenValue(sValue, Left(tgmCriteria.CellValue(nCol_FieldType, m_nCurrRow), 1))
        Else
            Select Case Left(tgmCriteria.CellValue(nCol_FieldType, m_nCurrRow), 1)
            Case TYPE_STRING
                If sOP = "IN" Then
                    sValue = fnRemvMatchingBeginEndChar(sValue, "(", ")")
                Else
                    sValue = fnAddBeginEndQuotes(sValue)
                End If
            Case TYPE_DATE
                sValue = tfnDateString(tfnFormatDate(sValue), True)
            Case TYPE_DECIMAL, TYPE_LONG, TYPE_INT
                sValue = sValue
            Case TYPE_PHONE
                sValue = fnAddBeginEndQuotes(sValue)
            End Select
        End If
    Else
        sValue = "NULL"
    End If
    
    sNewOp = ""

    If fnCStr(txtFilter.Text) = "" Then
        If sOP = "IN" Then
            txtFilter.Text = sOP + "(" + sValue + ")"
        Else
            txtFilter.Text = sOP + sValue
        End If
    Else
        aryTemp = Split(txtFilter.Text, ";")
        
        For i = 0 To UBound(aryTemp)
            If Left(aryTemp(i), Len(sOP)) = sOP Then
                Exit For
            End If
        Next i
        
        If i > UBound(aryTemp) Then
            If sOP = "IN" Then
                sNewOp = sOP + "(" + sValue + ")"
            Else
                sNewOp = sOP + sValue
            End If
        Else
            If sOP = "LIKE" Or sOP = "MATCHES" Or sOP = "BETWEEN" Then
                aryTemp(i) = sOP + sValue
            ElseIf sOP = "IN" Then
                sTemp = Mid(aryTemp(i), 3)
                sTemp = fnRemvMatchingBeginEndChar(sTemp, "(", ")")
                sValue = sTemp + "," + sValue
                aryTemp(i) = sOP + "(" + sValue + ")"
            Else
                'if not already exist in the filter, add it to the filter
                If InStr("," + Mid(aryTemp(i), Len(sOP) + 1) + ",", "," + sValue + ",") <= 0 Then
                    aryTemp(i) = aryTemp(i) + "," + sValue
                End If
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

Private Function fnFixBetweenValue(ByVal sValue As String, sDataType As String) As String
    Dim aryTemp() As String
    Dim sTemp As String
    Dim i As Integer
    Dim sRet As String
    
    sValue = fnRemvMatchingBeginEndChar(sValue, "(", ")")
    sValue = Replace(sValue, " and ", ",")
    
    aryTemp = Split(sValue, ",")
    sRet = ""
    
    If UBound(aryTemp) = 1 Then
    For i = 0 To UBound(aryTemp)
        Select Case sDataType
        Case TYPE_STRING, TYPE_PHONE
            sRet = sRet + fnAddBeginEndQuotes(aryTemp(i))
        Case TYPE_DATE
            sRet = sRet + tfnDateString(tfnFormatDate(aryTemp(i)), True)
        Case TYPE_DECIMAL, TYPE_LONG, TYPE_INT
            sRet = sRet + aryTemp(i)
        End Select
        
        If i = 1 Then
            Exit For
        End If
        
        sRet = sRet + ","
    Next i
    End If

    fnFixBetweenValue = "(" + sRet + ")"
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
        .AddEditColumn nCol_Sort, "Enter Sort Opt:(A)scending/(D)escending. Click on" _
            + " 'Filter' col to enter/change sel. criteria", "^(([AD])|(\N\/\A))$"
        nCol_OldFilter = .AddHiddenField("OldFilter")
        nCol_OldSort = .AddHiddenField("OldSort")
        nCol_DefaultSort = .AddHiddenField("DefaultSort")
        nCol_FieldName = .AddHiddenField("FieldName")
        nCol_FieldType = .AddHiddenField("FiledType")
        nCol_RegExp = .AddHiddenField("RegExp")
        nCol_NdxTgCombo = .AddHiddenField("NdxTgCombo")
        nCol_ComboSQL = .AddHiddenField("ComboSQL")
        nCol_CriteriaField = .AddHiddenField("CriteriaField")
        nCol_Operators = .AddHiddenField("Operators")
        nCol_Enabled = .AddHiddenField("Enabled")
        nCol_AllowEditSortCol = .AddHiddenField("AllowEditSortCol")
        nCol_XtraColName = .AddHiddenField("XtraColName")
        nCol_XtraRegExp = .AddHiddenField("XtraRegExp")
        nCol_sPreReqFields = .AddHiddenField("sPreReqFields")
        
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
    
    With tgmCriteria
        For i = 0 To .RowCount - 1
            If nRow < 0 Or (nRow >= 0 And nRow = i) Then
                sColumnName = fnCStr(.CellValue(nCol_FieldName, i))
                'get the real field
                sColumnName = fnGetFieldName(sColumnName, "R")
                
                sTemp = ""
                
                If fnCStr(.CellValue(nCol_Filter, i)) <> "" And fnCStr(.CellValue(nCol_Filter, i)) <> "N/A" Then
                    aryTemp = Split(fnCStr(.CellValue(nCol_Filter, i)), ";")
                    
                    For j = 0 To UBound(aryTemp)
                        If Left(fnCStr(aryTemp(j)), 7) = "BETWEEN" Then
                            sOperator = "BETWEEN"
                            sValue = fnCStr(Mid(fnCStr(aryTemp(j)), 8))
                        ElseIf Left(fnCStr(aryTemp(j)), 7) = "MATCHES" Then
                            sOperator = "MATCHES"
                            sValue = Mid(fnCStr(aryTemp(j)), 8)
                        ElseIf Left(fnCStr(aryTemp(j)), 4) = "LIKE" Then
                            sOperator = "LIKE"
                            sValue = Mid(fnCStr(aryTemp(j)), 5)
                        ElseIf Left(fnCStr(aryTemp(j)), 2) = "IN" Then
                            sOperator = "IN"
                            sValue = Mid(fnCStr(aryTemp(j)), 3)
                        Else
                            sOperator = Left(fnCStr(aryTemp(j)), 1)
                            sValue = Mid(fnCStr(aryTemp(j)), 2)
                        End If
                        
                        'Here we are checking it is < or <>
                        If Left(sValue, 1) = ">" Or Left(sValue, 1) = "=" Then
                            sOperator = sOperator + Left(sValue, 1)
                            sValue = Mid(sValue, 2)
                        End If
                        
                        If sOperator = "BETWEEN" Then
                            sValue = fnRemvMatchingBeginEndChar(sValue, "(", ")")
                            sValue = Replace(sValue, ",", " AND ")
                            
                            If sTemp <> "" Then
                                sTemp = sTemp & " OR "
                            End If
                            
                            sTemp = sTemp & sColumnName & " " & sOperator & " " & sValue
                        ElseIf sOperator = "IN" Then
                            If sTemp <> "" Then
                                sTemp = sTemp & " OR "
                            End If
                            
                            sTemp = sTemp & sColumnName & " " & sOperator & " " & sValue
                        Else
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
                        End If
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
                
                If .CellValue(nCol_Enabled, i) Then
                    If fnCStr(.CellValue(nCol_Sort, i)) <> "" Or m_bOrderByRequired Then
                        If sOrderBy = "" Then
                            sOrderBy = "ORDER BY " & sColumnName
                        Else
                            sOrderBy = sOrderBy & ", " & sColumnName
                        End If
                        
                        If fnCStr(.CellValue(nCol_Sort, i)) = "D" Then
                            sOrderBy = sOrderBy & " DESC"
                        ElseIf fnCStr(.CellValue(nCol_DefaultSort, i)) = "D" Then
                            sOrderBy = sOrderBy & " DESC"
                        End If
                    End If
                End If
            End If
        Next i
    End With
    
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
    Dim lRow As Long
    
    lRow = tgmCriteria.GetCurrentRowNumber
    
    If ColIndex = nCol_Sort Then
        If Not tgmCriteria.CellValue(nCol_AllowEditSortCol, lRow) Then
            CANCEL = True
            Exit Sub
        End If
        
        If Not tgmCriteria.CellValue(nCol_Enabled, lRow) Then
            CANCEL = True
            Exit Sub
        End If
    End If

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
    Dim lRow As Long
    
    tgmCriteria.GotFocus
    
    lRow = tgmCriteria.GetCurrentRowNumber
    
    If Not tgmCriteria.CellValue(nCol_Enabled, lRow) Then
        m_frmParent.tfnSetStatusBarMessage "Selection Criteria is not applied"
        Exit Sub
    End If
    
    If tblCriteria.Col <> nCol_Sort Then
        m_frmParent.tfnSetStatusBarMessage "Click on 'Filter' column to enter/change selection criteria"
    ElseIf tblCriteria.Col = nCol_Sort Then
        If Not tgmCriteria.CellValue(nCol_AllowEditSortCol, lRow) Then
            m_frmParent.tfnSetStatusBarMessage "Click on 'Filter' column to enter/change selection criteria"
            Exit Sub
        Else
            'override the 'Data is valid' message
            If tgmCriteria.ValidCell(nCol_Sort, lRow) Then
                m_frmParent.tfnSetStatusBarMessage "Enter Sort Opt:(A)scending/(D)escending. Click on" _
                    + " 'Filter' col to enter/change sel. criteria"
            End If
        End If
    End If
    
    subSetGridButtonStatus
End Sub

Private Sub tblCriteria_KeyDown(KeyCode As Integer, Shift As Integer)
    If Not tblCriteria.EditActive Then
        If KeyCode = vbKeyReturn Then
            If tgmCriteria.IsEndOfTable() Then
                If cmdOkBtn.Enabled Then
                    subSetFocus cmdOkBtn
                    Exit Sub
                End If
            End If
        ElseIf KeyCode = vbKeySpace Then
            'popup the the filter form
            subShowFilterPopup
        Else
            nClickCol = -1
        End If
    End If
    
    tgmCriteria.KeyDown KeyCode, Shift
End Sub

Private Sub tblCriteria_KeyPress(KeyAscii As Integer)
    If tblCriteria.Col = nCol_Sort Then
        KeyAscii = fnUcase(KeyAscii)
        
        If KeyAscii <> vbKeyBack And KeyAscii <> vbKeyReturn Then
            If SRegExpMatch("^[AD]$", Chr(KeyAscii)) <> 0 Then
                Beep
                KeyAscii = 0
            End If
        End If
    End If
    If Not tgmCriteria.Keypress(KeyAscii) Then
       KeyAscii = 0
    End If
End Sub

Private Sub tblCriteria_LostFocus()
    tgmCriteria.LostFocus
End Sub

Private Sub tblCriteria_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    nClickCol = tblCriteria.ColContaining(x)
'Debug.Print "x="; x; ", nClickCol = "; nClickCol
End Sub

Private Sub tblCriteria_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    Dim lRow As Long
    
    tgmCriteria.RowColChange LastRow, LastCol
    
    lRow = tgmCriteria.GetCurrentRowNumber
    
    If tgmCriteria.CellValue(nCol_Enabled, lRow) Then
        If tblCriteria.Col <> nCol_Sort Then
            m_frmParent.tfnSetStatusBarMessage "Click on 'Filter' column to enter/change selection criteria"
        ElseIf tblCriteria.Col = nCol_Sort Then
            If Not tgmCriteria.CellValue(nCol_AllowEditSortCol, tgmCriteria.GetCurrentRowNumber) Then
                m_frmParent.tfnSetStatusBarMessage "Click on 'Filter' column to enter/change selection criteria"
            Else
                'override the 'Data is valid' message
                If tgmCriteria.ValidCell(nCol_Sort, lRow) Then
                    m_frmParent.tfnSetStatusBarMessage "Enter Sort Opt:(A)scending/(D)escending. Click on" _
                        + " 'Filter' col to enter/change sel. criteria"
                End If
            End If
        End If
    
        If tgmCriteria.RowCount > 0 Then
            If nClickCol = nCol_Filter Or nClickCol = nCol_ColumnName Then
                'popup the the filter form
                subShowFilterPopup
            End If
        End If
    Else
        m_frmParent.tfnSetStatusBarMessage "Selection Criteria is not applied"
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
    Dim sCaption As String
    
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
    
    sCaption = tgmCriteria.CellValue(nCol_XtraColName, m_nCurrRow)
    
    If sCaption = "" Then  'no extra column/textbox
        lblDesc.Visible = False
        txtDesc.Visible = False
        cmdDesc.Visible = False
        cmdValue.Left = cmdDesc.Left
        txtValue.Width = sngDefaultValueWidth + cmdValue.Width + txtDesc.Width + 70
    Else
        cmdValue.Left = sngDefaultValueLeft
        txtValue.Width = sngDefaultValueWidth
        sRegExpDesc = tgmCriteria.CellValue(nCol_XtraRegExp, m_nCurrRow)
        
        lblDesc.Caption = sCaption
        lblDesc.Visible = True
        txtDesc.Visible = True
        cmdDesc.Visible = True
    End If
    
    ndxPrevActiveCombo = tfnRound(tgmCriteria.CellValue(nCol_NdxTgCombo, m_nCurrRow))
    tgcDropdown.ComboActive(ndxPrevActiveCombo) = True
    
    cmdClear.Enabled = txtFilter.Text <> ""
    'To set the Regular Expression for Operation Text Box
    SubSetOpRegExp
    
    nClickCol = 2
    txtOp = ""
    txtValue = ""
    txtDesc = ""
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
    'subSetFocus tblCriteria
    subSetFocus cmdOkBtn
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
        'for Debugging only
        'Set tgcDropdown = New clsComboControl
        Set tgcDropdown = CreateObject(t_szOLECOMBO)
        
        subCreateTmpOperator
        
        If Not tgcDropdown Is Nothing Then
            With tgcDropdown
                Set .DBEngine = t_engFactor
                Set .Form = Me
                Set .DataBase = t_dbMainDatabase
                Set .DataLink = datComboDropDown
                Set .Table = tblComboDropdown
                .SortOnColumnClick = True
                
                'Set .LocalDatabase = dbLocal (Your local database object variable)
                'Add the combos below
                
                .AddCombo "SELECT op FROM temp_operator", , tgcDropdown.DropUp
                .AddComboBox txtOp, cmdOp, "op", .SQL_STRING_TYPE(10)
                .SetOrderingNone txtOp
                
                'more will be added in AddColumn() function - ON THE FLY
            End With
        End If
    
        'save the default left and width
        sngDefaultValueLeft = cmdValue.Left
        sngDefaultValueWidth = txtValue.Width
    End If

    Exit Sub

errTrap:
    tfnErrHandler "Criteria.subSetupCombos"
End Sub

Private Sub tblComboDropdown_Click()
    tgcDropdown.Click tblComboDropdown
End Sub

Private Sub tblComboDropDown_GotFocus()
    tgcDropdown.GotFocus tblComboDropdown
End Sub

Private Sub tblComboDropDown_HeadClick(ByVal ColIndex As Integer)
    tgcDropdown.HeadClick ColIndex
End Sub

Private Sub tblComboDropDown_KeyPress(KeyAscii As Integer)
    tgcDropdown.Keypress tblComboDropdown, KeyAscii
End Sub

Private Sub tblComboDropDown_LostFocus()
    tgcDropdown.LostFocus tblComboDropdown
End Sub

Private Sub tblComboDropDown_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    tgcDropdown.TableMouseUp y
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
                If txtFilter.Text <> "" Then
                    cmdFilterOK.Enabled = True
                    subSetFocus cmdFilterOK
                Else
                    SendKeys "{TAB}", True
                End If
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
            If txtFilter.Text <> "" Then
                cmdFilterOK.Enabled = True
                txtFilter.Locked = True
            End If
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
        txtOp.Tag = "subSetFocus txtValue"
    End If
End Sub

Private Sub txtOp_KeyPress(KeyAscii As Integer)
    #If ProtoType Then
        Exit Sub
    #End If
    
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    
    If Not tfnRegExpControlKeyPress(txtOp, KeyAscii, sRegExpOp) Then
        KeyAscii = 0
        Exit Sub
    End If
    If KeyAscii = vbKeyReturn Then
        subLoadOpSQL
        
        If txtOp <> "" Then
            KeyAscii = 0
            subSetFocus txtValue
            txtOp.Tag = "subSetFocus txtValue"
        End If
    End If
    If Not tgcDropdown.Keypress(txtOp, KeyAscii) Then
        If KeyAscii = vbKeyReturn Then
            If tgcDropdown.SingleRecordSelected Then
                subSetButtonStatus
                subSetFocus txtValue
                txtOp.Tag = "subSetFocus txtValue"
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
    If cValidate.LostFocus(txtOp, cmdOp, tblComboDropdown) Then
        subSetButtonStatus
        
        If fnCStr(tgmCriteria.CellValue(nCol_XtraColName, m_nCurrRow)) <> "" Then
            If InStr("IN|LIKE|MATCHES|BETWEEN", txtOp.Text) > 0 Then
                lblDesc.Visible = False
                txtDesc.Visible = False
                cmdDesc.Visible = False
                cmdValue.Left = cmdDesc.Left
                txtValue.Width = sngDefaultValueWidth + cmdValue.Width + txtDesc.Width + 70
            Else
                cmdValue.Left = sngDefaultValueLeft
                txtValue.Width = sngDefaultValueWidth
                lblDesc.Visible = True
                txtDesc.Visible = True
                cmdDesc.Visible = True
                
                If txtOp.Tag = "subSetFocus txtValue" Then
                    subSetFocus txtValue
                End If
                
                txtOp.Tag = ""
            End If
        End If
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
    If InStr("IN|LIKE|MATCHES|BETWEEN", txtOp.Text) > 0 Then
        tfnRegExpControlChange txtValue, "^(P{0,500})$"
    Else
        If sRegExp <> "" Then
            tfnRegExpControlChange txtValue, sRegExp
        End If
    End If
    
    cValidate.Change txtValue
    
    If ActiveControl Is txtValue Then
        cmdAdd.Enabled = False
    End If
End Sub

Private Sub txtValue_GotFocus()
    If InStr("IN|LIKE|MATCHES|BETWEEN", txtOp.Text) > 0 Then
        tfnRegExpControlGotFocus txtValue, "^(P{0,500})$"
    Else
        If sRegExp <> "" Then
            tfnRegExpControlGotFocus txtValue, sRegExp
        End If
    End If
    
    cValidate.GotFocus txtValue
    
    If InStr("IN|LIKE|MATCHES|BETWEEN", txtOp.Text) <= 0 Then
        tgcDropdown.GotFocus txtValue
        If tgcDropdown.SingleRecordSelected Then
            subSetButtonStatus
            
            If cmdAdd.Enabled Then
                subSetFocus cmdAdd
            End If
        End If
    End If
End Sub

Private Sub txtValue_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Or KeyCode = vbKeyBack Then
        txtDesc.Text = ""
    End If
End Sub

Private Sub txtValue_KeyPress(KeyAscii As Integer)
    #If ProtoType Then
        Exit Sub
    #End If
    
    If InStr("IN|LIKE|MATCHES|BETWEEN", txtOp.Text) > 0 Then
        If Not tfnRegExpControlKeyPress(txtValue, KeyAscii, "^(P{0,500})$") Then
            KeyAscii = 0
            Exit Sub
        End If
        
        If KeyAscii = vbKeyReturn Then
            KeyAscii = 0
            subSetButtonStatus
            
            If cmdAdd.Enabled Then
                subSetFocus cmdAdd
            Else
                SendKeys "{TAB}", True
            End If
        Else
            cValidate.Keypress txtValue, KeyAscii
            txtDesc.Text = ""
        End If
        
        Exit Sub
    End If
    
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
        txtDesc.Text = ""
    End If
End Sub

Private Sub txtValue_LostFocus()
    #If ProtoType Then
        Exit Sub
    #End If
    If cValidate.LostFocus(txtValue, cmdValue, txtDesc, cmdDesc, tblComboDropdown) Then
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
    
    If InStr("IN|LIKE|MATCHES|BETWEEN", txtOp.Text) > 0 Then
        Exit Sub
    End If
    
    subFixComboSQL
    tgcDropdown.Click cmdValue
End Sub

Private Sub txtDesc_Click()
    tgcDropdown.Click txtDesc
End Sub

Private Sub txtDesc_Change()
    If sRegExpDesc <> "" Then
        tfnRegExpControlChange txtDesc, sRegExpDesc
    End If
    
    cValidate.Change txtDesc
End Sub

Private Sub txtDesc_GotFocus()
    If sRegExpDesc <> "" Then
        tfnRegExpControlGotFocus txtDesc, sRegExpDesc
    End If
    
    cValidate.GotFocus txtDesc
    tgcDropdown.GotFocus txtDesc
    If tgcDropdown.SingleRecordSelected Then
        subSetButtonStatus
        
        If cmdAdd.Enabled Then
            subSetFocus cmdAdd
        End If
    End If
End Sub

Private Sub txtDesc_KeyPress(KeyAscii As Integer)
    #If ProtoType Then
        Exit Sub
    #End If
    
    If sRegExpDesc <> "" Then
        If Not tfnRegExpControlKeyPress(txtDesc, KeyAscii, sRegExpDesc) Then
            KeyAscii = 0
            Exit Sub
        End If
    End If
            
    If KeyAscii = vbKeyReturn Then
        Screen.MousePointer = vbHourglass
        subFixComboSQL
    End If
    
    If InStr("IN|LIKE|MATCHES|BETWEEN", txtOp.Text) > 0 Then
        cValidate.Keypress txtDesc, KeyAscii
    Else
        If Not tgcDropdown.Keypress(txtDesc, KeyAscii) Then
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
            cValidate.Keypress txtDesc, KeyAscii
        End If
    End If
End Sub

Private Sub txtDesc_LostFocus()
    #If ProtoType Then
        Exit Sub
    #End If
    If cValidate.LostFocus(txtValue, cmdValue, txtDesc, cmdDesc, tblComboDropdown) Then
        subSetButtonStatus
        
        If cmdAdd.Enabled Then
            If ActiveControl Is cmdFilterCancel Then
                subSetFocus cmdAdd
            End If
        End If
    End If
End Sub

Private Sub cmdDesc_Click()
    #If ProtoType Then
        Exit Sub
    #End If
    
    If InStr("IN|LIKE|MATCHES|BETWEEN", txtOp.Text) > 0 Then
        Exit Sub
    End If
    
    subFixComboSQL
    tgcDropdown.Click cmdDesc
End Sub

Private Sub subFixComboSQL()
    Dim strSQL As String
    Dim aryTemp() As String
    Dim i As Integer
    Dim sCriteria As String
    
    On Error GoTo errTrap
    
    i = tgmCriteria.GetCurrentRowNumber
    aryTemp = Split(tgmCriteria.CellValue(nCol_sPreReqFields, i), ",")
    
    If UBound(aryTemp) < 0 Then
        Exit Sub
    End If
    
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

Private Function fnGetCriteriaByFieldName(ByVal sField As String) As String
    Dim i As Integer
    Dim sCriteria As String
    
    'get real field name
    sField = fnGetFieldName(sField, "R")
    
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
        .ESCControl = tblComboDropdown
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
    Dim aryTemp() As String
    Dim i As Integer
    Dim sTemp As String
    Dim sRet As String
    Dim bAdd As Boolean
    
    
    aryTemp = Split(sIn, ",")
    sRet = ""
    
    For i = 0 To UBound(aryTemp)
        sTemp = Trim(aryTemp(i))
        bAdd = False
        
        If sTemp <> "" Then
            If (Left(sTemp, 1) = "'" And Right(sTemp, 1) = "'") Then
                'do nothing
            ElseIf (Left(sTemp, 1) = """" And Right(sTemp, 1) = """") Then
                If Len(sTemp) > 2 Then
                    sTemp = "'" + Mid(sTemp, 2, Len(sTemp) - 2) + "'"
                Else
                    sTemp = "''"
                End If
            Else
                sTemp = tfnSQLString(sTemp)
            End If
            
            If InStr(sRet, sTemp) <= 0 Then
                bAdd = True
            End If
            
            If bAdd Then
                If sRet <> "" Then
                    sRet = sRet + ","
                End If
                sRet = sRet + sTemp
            End If
        End If
    Next i
    
    fnAddBeginEndQuotes = sRet
End Function
'''
'''Private Function fnRemvBeginEndQuotes(ByVal sIn As String) As String
'''    If Left(sIn, 1) = "'" And Right(sIn, 1) = "'" Then
'''        If Len(sIn) > 2 Then
'''            sIn = Mid(sIn, 2, Len(sIn) - 2)
'''        Else
'''            sIn = ""
'''        End If
'''    End If
'''
'''    fnRemvBeginEndQuotes = sIn
'''End Function

Private Function fnRemvMatchingBeginEndChar(ByVal sIn As String, _
                                            ByVal sChar1 As String, _
                                            Optional ByVal sChar2 As String = "") As String
    
    If sChar2 = "" Then
        sChar2 = sChar1
    End If
    
    If Left(sIn, 1) = sChar1 And Right(sIn, 1) = sChar2 Then
        If Len(sIn) > 2 Then
            sIn = Mid(sIn, 2, Len(sIn) - 2)
        Else
            sIn = ""
        End If
    End If
    
    fnRemvMatchingBeginEndChar = sIn
End Function

Private Function fnConvtToRegExp(sIn As String) As String
    Dim ary() As String
    Dim i As Integer
    Dim sTemp As String
    Dim sRet As String
    
    ary = Split(sIn, ",")
    
    For i = 0 To UBound(ary)
        sTemp = fnCStr(ary(i))
        
        If sTemp <> "" Then
            If sRet <> "" Then
                sRet = sRet + "|"
            End If
            
            sRet = sRet + "(" + sTemp + ")"
        End If
    Next i
    
    If sRet <> "" Then
        fnConvtToRegExp = "^(" + sRet + ")$"
    Else
        fnConvtToRegExp = ""
    End If
End Function

Private Sub subCreateTmpOperator()
   Const SUB_NAME As String = "subCreateTmpOperator"
   
   Dim strSQL As String
    
    On Error GoTo LetUsGo:
    strSQL = "DROP TABLE temp_operator"
    t_dbMainDatabase.ExecuteSQL strSQL
LetUsGo:
   
   On Error GoTo errCreateTable
   
   strSQL = "CREATE TEMP TABLE temp_operator (op char(10), seq smallint)"
   t_dbMainDatabase.ExecuteSQL strSQL
    
   strSQL = "INSERT INTO temp_operator VALUES ('=', 1)"
   t_dbMainDatabase.ExecuteSQL strSQL
   
   strSQL = "INSERT INTO temp_operator VALUES ('>=', 2)"
   t_dbMainDatabase.ExecuteSQL strSQL
   
   strSQL = "INSERT INTO temp_operator VALUES ('<=', 3)"
   t_dbMainDatabase.ExecuteSQL strSQL
   
   strSQL = "INSERT INTO temp_operator VALUES ('<', 4)"
   t_dbMainDatabase.ExecuteSQL strSQL
   
   strSQL = "INSERT INTO temp_operator VALUES ('>', 5)"
   t_dbMainDatabase.ExecuteSQL strSQL
   
   strSQL = "INSERT INTO temp_operator VALUES ('<>', 6)"
   t_dbMainDatabase.ExecuteSQL strSQL
   
   strSQL = "INSERT INTO temp_operator VALUES ('IN', 7)"
   t_dbMainDatabase.ExecuteSQL strSQL
   
   strSQL = "INSERT INTO temp_operator VALUES ('LIKE', 8)"
   t_dbMainDatabase.ExecuteSQL strSQL
   
   strSQL = "INSERT INTO temp_operator VALUES ('MATCHES', 9)"
   t_dbMainDatabase.ExecuteSQL strSQL
   
   strSQL = "INSERT INTO temp_operator VALUES ('BETWEEN', 10)"
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
            If Left(fnCStr(aryTemp(j)), 7) = "BETWEEN" Then
                sOP = "BETWEEN"
                sValue = Mid(fnCStr(aryTemp(j)), 8)
            ElseIf Left(fnCStr(aryTemp(j)), 7) = "MATCHES" Then
                sOP = "MATCHES"
                sValue = Mid(fnCStr(aryTemp(j)), 8)
            ElseIf Left(fnCStr(aryTemp(j)), 4) = "LIKE" Then
                sOP = "LIKE"
                sValue = Mid(fnCStr(aryTemp(j)), 5)
            ElseIf Left(fnCStr(aryTemp(j)), 2) = "IN" Then
                sOP = "IN"
                sValue = Mid(fnCStr(aryTemp(j)), 3)
            Else
                sOP = Left(fnCStr(aryTemp(j)), 1)
                sValue = Mid(fnCStr(aryTemp(j)), 2)
            End If
            
            If sOP = "" Then
                fnValidFilter = "Operator is missing"
                Exit Function
            End If
            
            'Here we are checking it is < or <>
            If InStr("IN|LIKE|MATCHES|BETWEEN", sOP) <= 0 Then
                If Left(sValue, 1) = ">" Then
                    If sOP <> "<" Then
                        fnValidFilter = "Invalid operator '" + sOP + Left(sValue, 1) + "' encountered"
                        Exit Function
                    End If
                    
                    sOP = "<>"
                    sValue = Mid(sValue, 2)
                ElseIf Left(sValue, 1) = "=" Then
                    sOP = sOP + "="
                    sValue = Mid(sValue, 2)
                End If
            End If
            
            Select Case sOP
            Case "<", ">", "=", "<>", ">=", "<=", "IN", "LIKE", "MATCHES", "BETWEEN"
                'good
            Case Else
                fnValidFilter = "Invalid operator '" + sOP + "' encountered"
                Exit Function
            End Select
            
            If sOP = "<" Or sOP = ">" Or sOP = ">=" Or sOP = "<=" Then
                If Left(tgmCriteria.CellValue(nCol_FieldType, m_nCurrRow), 1) = TYPE_STRING Then
                    fnValidFilter = sOP + " is not allow for char type column"
                    Exit Function
                End If
            End If
            
            If sOP = "IN" Then
                sValue = fnRemvMatchingBeginEndChar(sValue, "(", ")")
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
                        If sOP = "<" Or sOP = ">" Or ">=" Or "<=" Then
                            fnValidFilter = "NULL is not valid for operator " + sOP
                            Exit Function
                        End If
                        
                        sTemp = "NULL"
                    End If
                    
                    Select Case Left(tgmCriteria.CellValue(nCol_FieldType, m_nCurrRow), 1)
                    Case TYPE_DATE
                        If sTemp <> "NULL" Then
                            sTemp = tfnFormatDate(fnRemvMatchingBeginEndChar(sTemp, "'"))
                            
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
        
            If sOP = "IN" Then
                sNewFilter = sNewFilter + sOP + "(" + sFilter + ")"
            Else
                sNewFilter = sNewFilter + sOP + sFilter
            End If
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
        subLoadOpSQL False
        
        strSQL = tgcDropdown.ComboSQL(txtBox)
        strSQL = strSQL + " AND op = " + tfnSQLString(UCase(txtBox.Text))
        
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
    Dim aryTemp() As String
    Dim sTemp As String
    Dim i As Integer
    
    fnValidValue = False
    
    If fnCStr(txtBox) = "" Then
        If fnCStr(txtOp) <> "=" And fnCStr(txtOp) <> "<>" Then
            Exit Function
        End If
    Else
        If txtOp.Text = "BETWEEN" Then
            If fnFixBetweenValue(txtBox.Text, Left(tgmCriteria.CellValue(nCol_FieldType, m_nCurrRow), 1)) <> "()" Then
                'format is valid
                fnValidValue = True
            End If
            Exit Function
        End If
        
        Select Case Left(tgmCriteria.CellValue(nCol_FieldType, m_nCurrRow), 1)
        Case TYPE_DATE
            If Not IsValidDate(txtBox.Text) Then
                cValidate.SetErrorMessage txtBox, "Invalid " & tgmCriteria.CellValue(nCol_ColumnName, m_nCurrRow)
                Exit Function
            Else
                txtBox = tfnFormatDate(txtBox)
            End If
        Case TYPE_STRING
            If InStr("IN|LIKE|MATCHES|BETWEEN", txtOp.Text) > 0 Then
                If txtOp.Text = "BETWEEN" Then
                    If InStr(txtBox.Text, ",") <= 0 And InStr(LCase(txtBox.Text), " and ") <= 0 Then
                        'format is not valid
                        Exit Function
                    End If
                End If
                
                fnValidValue = True
                Exit Function
            End If
        Case TYPE_DECIMAL, TYPE_INT, TYPE_LONG
            If InStr(txtBox.Text, ",") > 0 Then
                cValidate.SetErrorMessage txtBox, "Only one Value is allowed"
                Exit Function
            End If
            
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

Private Function fnGetCriteria(ByVal sFieldName As String, ByVal sValue As String) As String
    Const FUNC_NAME As String = "fnGetCriteria"
    
    sFieldName = fnGetFieldName(sFieldName, "R")
    sValue = fnCStr(sValue)
    
    fnGetCriteria = "(" & sFieldName
    
    Select Case Left(tgmCriteria.CellValue(nCol_FieldType, m_nCurrRow), 1)
        Case TYPE_STRING, TYPE_PHONE
            If sValue = "" Then
                fnGetCriteria = fnGetCriteria & " IS NULL" & " OR " & sFieldName & " = ''"
            Else
                fnGetCriteria = fnGetCriteria & " = " & fnAddBeginEndQuotes(sValue)
            End If
        Case TYPE_DATE
            If sValue = "" Then
                fnGetCriteria = fnGetCriteria & " IS NULL"
            Else
                fnGetCriteria = fnGetCriteria & " = " & tfnDateString(tfnFormatDate(sValue), True)
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
            If tgmCriteria.CellValue(nCol_Enabled, lRow) And sText <> szEMPTY Then
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

Private Sub subLoadOpSQL(Optional bAddOrderBy As Boolean = True)
    Dim strSQL As String
    Dim sOP As String
    
    strSQL = "SELECT op, seq FROM temp_operator WHERE 1=1"
        
    'operators overriden
    If tgmCriteria.CellValue(nCol_Operators, m_nCurrRow) <> "" Then
        sOP = fnAddBeginEndQuotes(tgmCriteria.CellValue(nCol_Operators, m_nCurrRow))
    Else
        Select Case Left(tgmCriteria.CellValue(nCol_FieldType, m_nCurrRow), 1)
        Case TYPE_STRING, TYPE_PHONE
            sOP = "'=','<>','IN','LIKE','MATCHES','BETWEEN'"
        Case Else
            sOP = "'=','<=','>=','<','>','<>','BETWEEN'"
        End Select
    End If
    
    If sOP <> "" Then
        strSQL = strSQL & " AND op IN " + "(" + sOP + ")"
    End If
    
    If bAddOrderBy Then
        strSQL = strSQL & " ORDER BY seq"
    End If
    
    tgcDropdown.ComboSQL(txtOp) = strSQL
End Sub

Private Sub SubSetOpRegExp()
    'operators overriden
    If tgmCriteria.CellValue(nCol_Operators, m_nCurrRow) <> "" Then
        sRegExpOp = fnConvtToRegExp(tgmCriteria.CellValue(nCol_Operators, m_nCurrRow))
    Else
        Select Case Left(tgmCriteria.CellValue(nCol_FieldType, m_nCurrRow), 1)
        Case TYPE_STRING, TYPE_PHONE
            sRegExpOp = "^((=)|(<>)|(IN)|(LIKE)|(MATCHES)|(BETWEEN))$"
        Case Else
            sRegExpOp = "^((=)|(<=)|(>=)|(<)|(>)|(<>)|(BETWEEN))$"
        End Select
    End If
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

Private Sub subSetClearAllBtnStatus()
    Dim i As Integer
    
    With tgmCriteria
        For i = 0 To .RowCount - 1
            If fnCStr(.CellValue(nCol_Filter, i)) <> "" And fnCStr(.CellValue(nCol_Filter, i)) <> "N/A" Then
                cmdClearAll.Enabled = True
                Exit Sub
            End If
        Next i
    End With
            
    cmdClearAll.Enabled = False
End Sub

Private Function fnGetFieldName(sFieldNameWithAlias As String, sRealOrAlias As String) As String
    Dim nPosi As Integer
    nPosi = InStr(sFieldNameWithAlias, " ")
    
    If nPosi > 0 Then
        If sRealOrAlias = "R" Then
            'real field name
            fnGetFieldName = Left(sFieldNameWithAlias, nPosi - 1)
        Else  'A
            'alias field name
            fnGetFieldName = Mid(sFieldNameWithAlias, nPosi + 1)
        End If
    Else
        fnGetFieldName = sFieldNameWithAlias
    End If
End Function
