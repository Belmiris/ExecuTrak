VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ETFMTaxRecord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***********************************************************'
'
' Copyright (c) 1996 FACTOR, A Division of W.R.Hess Company
'
' Module name   : EDIRECRD.CLS
' Date          : Apr. 14, 1998
' Programmer    : Qinggang Ma
'
' Defines a record for EDI fuel movement tax information
'
Option Explicit

    Public TaxScheduleCode As String
    Public ProductCode As String
    Public CarrierCode As String
    Public CarrierName As String
    Public CarrierFedID As String
    Public OriginCity As String
    Public OriginState As String
    Public DestCity As String
    Public DestState As String
    Public BuySellCode As String
    Public BuySellName As String
    Public BuySellFedID As String
    Public DetailType As String
    Public LadingNumber As String
    Public LadingDate As String
    Public Quantity As String
    Public TTTrnNumber As Long
    
    'FOr internal use
    Public Movement As Long

