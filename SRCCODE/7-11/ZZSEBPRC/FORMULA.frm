VERSION 5.00
Object = "{C75015E0-2232-11D3-B440-0060971E99AF}#1.0#0"; "FACTFRM.OCX"
Begin VB.Form frmFORMULA 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Bonus Formula Details"
   ClientHeight    =   3840
   ClientLeft      =   2760
   ClientTop       =   3756
   ClientWidth     =   8604
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3840
   ScaleWidth      =   8604
   ShowInTaskbar   =   0   'False
   Begin FACTFRMLib.FactorFrame FactorFrame1 
      Height          =   3288
      Left            =   60
      TabIndex        =   0
      Top             =   48
      Width           =   8460
      _Version        =   65536
      _ExtentX        =   14922
      _ExtentY        =   5800
      _StockProps     =   77
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelOuter      =   5
      BeginProperty PanelFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Formula :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   0
         Left            =   72
         TabIndex        =   24
         Top             =   84
         Width           =   1656
      End
      Begin VB.Label lblPercent 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1800
         TabIndex        =   23
         Top             =   1692
         Width           =   2208
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Percent(pct) :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   2
         Left            =   72
         TabIndex        =   22
         Top             =   1692
         Width           =   1656
      End
      Begin VB.Label lblACondition 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1800
         TabIndex        =   21
         Top             =   1284
         Width           =   6576
      End
      Begin VB.Label lblAFormula 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1800
         TabIndex        =   20
         Top             =   888
         Width           =   6576
      End
      Begin VB.Label lblCondition 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1800
         TabIndex        =   19
         Top             =   492
         Width           =   6576
      End
      Begin VB.Label lblFormula 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1800
         TabIndex        =   18
         Top             =   84
         Width           =   6576
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Adj Condition :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   7
         Left            =   72
         TabIndex        =   17
         Top             =   1284
         Width           =   1656
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Adj Formula :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   8
         Left            =   72
         TabIndex        =   16
         Top             =   888
         Width           =   1656
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Condition :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   9
         Left            =   72
         TabIndex        =   15
         Top             =   492
         Width           =   1656
      End
      Begin VB.Label lblAmount2 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1800
         TabIndex        =   14
         Top             =   2892
         Width           =   2208
      End
      Begin VB.Label lblAmount1 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1800
         TabIndex        =   13
         Top             =   2484
         Width           =   2208
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Amount2(amt2) :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   12
         Left            =   72
         TabIndex        =   12
         Top             =   2892
         Width           =   1656
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Amount1(amt1) :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   13
         Left            =   72
         TabIndex        =   11
         Top             =   2484
         Width           =   1656
      End
      Begin VB.Label lblVar1 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   6168
         TabIndex        =   10
         Top             =   2088
         Width           =   2208
      End
      Begin VB.Label lblVariable1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Variable1(v1) :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   4044
         TabIndex        =   9
         Top             =   2088
         Width           =   2088
      End
      Begin VB.Label lblDollar 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1800
         TabIndex        =   8
         Top             =   2088
         Width           =   2208
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Dollar(dol) :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   17
         Left            =   72
         TabIndex        =   7
         Top             =   2088
         Width           =   1656
      End
      Begin VB.Label lblMaxTotal 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   6168
         TabIndex        =   6
         Top             =   1692
         Width           =   2208
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Max. Total(mxt) :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   19
         Left            =   4044
         TabIndex        =   5
         Top             =   1692
         Width           =   2088
      End
      Begin VB.Label lblVar3 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   6168
         TabIndex        =   4
         Top             =   2892
         Width           =   2208
      End
      Begin VB.Label lblVar2 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   6168
         TabIndex        =   3
         Top             =   2484
         Width           =   2208
      End
      Begin VB.Label lblVariable3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Variable3(v3) :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   4044
         TabIndex        =   2
         Top             =   2892
         Width           =   2088
      End
      Begin VB.Label lblVariable2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Variable2(v2) :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   4044
         TabIndex        =   1
         Top             =   2484
         Width           =   2088
      End
   End
   Begin FACTFRMLib.FactorFrame cmdClose 
      Height          =   396
      HelpContextID   =   2901
      Left            =   7212
      TabIndex        =   25
      Top             =   3396
      Width           =   1308
      _Version        =   65536
      _ExtentX        =   2302
      _ExtentY        =   688
      _StockProps     =   77
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "&Close"
      CaptionPos      =   4
      PicturePos      =   3
      ShowFocusRect   =   -1  'True
      Style           =   3
      BorderWidth     =   4
      BeginProperty PanelFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmFORMULA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdClose_GotFocus()
    frmZZSEBPRC.tfnSetStatusBarMessage "Close Bonus Formula Details"
End Sub

Private Sub Form_Load()
    Move frmZZSEBPRC.Left + 162, frmZZSEBPRC.Top + 1160
    subSetFocus frmFORMULA.cmdClose
    DoEvents
End Sub

Public Function fnLoadBonusFormula(sCode As String, nLevel As Integer) As Boolean
    Const SUB_NAME As String = "fnLoadBonusFormula"
    Dim strSQL As String
    Dim rsTemp As Recordset
    
    fnLoadBonusFormula = False
    subClearFormulaDetails
    
    If Trim(sCode) = "" Or nLevel = 0 Then
        Exit Function
    End If
    
    strSQL = "SELECT * FROM bonus_formula WHERE bf_bonus_code = " & tfnSQLString(Trim(sCode))
    strSQL = strSQL & " AND bf_level = " & tfnRound(nLevel)
    
    If GetRecordSet(rsTemp, strSQL, , SUB_NAME) < 0 Then
        MsgBox "Failed to access database to get formula details", vbExclamation
        Exit Function
    End If
    
    If rsTemp.RecordCount = 0 Then
        MsgBox "No details available", vbExclamation
        Exit Function
    End If
    
    If rsTemp.RecordCount = 1 Then
        lblPercent = tfnRound(rsTemp!bf_percent, DEFAULT_DECIMALS)
        lblDollar = tfnRound(rsTemp!bf_dollar, 2)
        lblAmount1 = tfnRound(rsTemp!bf_amount1, DEFAULT_DECIMALS)
        lblAmount2 = tfnRound(rsTemp!bf_amount2, DEFAULT_DECIMALS)
        lblMaxTotal = tfnRound(rsTemp!bf_max_total, 2)
        lblVariable1 = fnGetField(rsTemp!bf_variable1) & "(v1) :"
        lblVariable2 = fnGetField(rsTemp!bf_variable2)
        If lblVariable2 = "" Then
            lblVariable2 = "Not Used :"
        Else
            lblVariable2 = lblVariable2 & "(v2) :"
        End If
        lblVariable3 = fnGetField(rsTemp!bf_variable3)
        If lblVariable3 = "" Then
            lblVariable3 = "Not Used :"
        Else
            lblVariable3 = lblVariable3 & "(v3) :"
        End If
        lblFormula = fnGetField(rsTemp!bf_formula)
        lblCondition = fnGetField(rsTemp!bf_condition)
        lblAFormula = fnGetField(rsTemp!bf_adj_formula)
        lblACondition = fnGetField(rsTemp!bf_adj_condition)
    End If
    
    fnLoadBonusFormula = True
    subSetFocus cmdClose

End Function

Private Sub subClearFormulaDetails()
    lblPercent = ""
    lblDollar = ""
    lblAmount1 = ""
    lblAmount2 = ""
    lblMaxTotal = ""
    lblVar1 = ""
    lblVar2 = ""
    lblVar3 = ""
    lblFormula = ""
    lblCondition = ""
    lblAFormula = ""
    lblACondition = ""
End Sub
