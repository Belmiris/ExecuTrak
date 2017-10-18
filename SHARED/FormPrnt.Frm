VERSION 5.00
Begin VB.Form FormCapturer 
   AutoRedraw      =   -1  'True
   Caption         =   "Form1"
   ClientHeight    =   2205
   ClientLeft      =   3390
   ClientTop       =   2685
   ClientWidth     =   3960
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2205
   ScaleWidth      =   3960
   Begin VB.PictureBox pctHold 
      BorderStyle     =   0  'None
      Height          =   1224
      Left            =   408
      ScaleHeight     =   1230
      ScaleWidth      =   3135
      TabIndex        =   0
      Top             =   492
      Visible         =   0   'False
      Width           =   3132
   End
End
Attribute VB_Name = "FormCapturer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Useage of this module:

    '1. Include this module to the project
    '2. When printing is needed, call the Public Sub PrintScreen,
    '   passing in the form to be printed. Example:
    '       FormCapturer.ScreenPrint frmArepaymt
    '   or  FormCapturer.ScreenPrint Me (If called with the form)
    
Option Explicit
    
    Private Type DispInterface
       Data1 As Long
       Data2 As Integer
       Data3 As Integer
       Data4(7) As Byte
    End Type
     
    Private Type BmpPicture
       Size As Long
       Type As Long
       hBmp As Long
       hPal As Long
       Reserved As Long
    End Type
    
    Private Declare Function BitBlt Lib "GDI32" ( _
        ByVal hDestDC As Long, _
        ByVal x As Long, _
        ByVal y As Long, _
        ByVal nWidth As Long, _
        ByVal nHeight As Long, _
        ByVal hSrcDC As Long, _
        ByVal XSrc As Long, _
        ByVal YSrc As Long, _
        ByVal dwRop As Long) As Long
    
    Private Declare Function CreateCompatibleBitmap Lib "GDI32" ( _
        ByVal hDC As Long, _
        ByVal nWidth As Long, _
        ByVal nHeight As Long) As Long

    Private Declare Function CreateCompatibleDC Lib "GDI32" ( _
        ByVal hDC As Long) As Long

    Private Declare Function DeleteDC Lib "GDI32" ( _
        ByVal hDC As Long) As Long
    
    Private Declare Function GetWindowDC Lib "user32" ( _
        ByVal hwnd As Long) As Long

    Private Declare Function OleCreatePictureIndirect Lib "olepro32.dll" ( _
       PicDesc As BmpPicture, _
       IDisp As DispInterface, _
       ByVal fPictureOwnsHandle As Long, _
       IPic As IPicture) As Long
    
    Private Declare Function ReleaseDC Lib "user32" ( _
        ByVal hwnd As Long, _
        ByVal hDC As Long) As Long
    
    Private Declare Function SelectObject Lib "GDI32" ( _
        ByVal hDC As Long, _
        ByVal hObject As Long) As Long


Private Function CaptureWindow(ByVal hWndForm As Long, _
                               ByVal lLeft As Long, _
                               ByVal lTop As Long, _
                               ByVal lWidth As Long, _
                               ByVal lHeight As Long) As Picture
    
    Dim hDCMemory As Long
    Dim hBmpHold As Long
    Dim hBmpPrev As Long
    Dim hDCForm As Long
    Dim hPalPrev As Long
    
    hDCForm = GetWindowDC(hWndForm)
    
    ' Create a device context in the memory
    hDCMemory = CreateCompatibleDC(hDCForm)
    ' Create a bitmap and place it in the memory DC
    hBmpHold = CreateCompatibleBitmap(hDCForm, lWidth, lHeight)
    hBmpPrev = SelectObject(hDCMemory, hBmpHold)
    
    ' Copy the on-screen image into the memory DC
    BitBlt hDCMemory, 0, 0, lWidth, lHeight, hDCForm, _
       lLeft, lTop, vbSrcCopy
    
    SelectObject hDCMemory, hBmpPrev
    
    'Clean the memory
    DeleteDC hDCMemory
    ReleaseDC hWndForm, hDCForm

    Set CaptureWindow = CreateBitmapPicture(hBmpHold, 0)
End Function
 
Private Function CreateBitmapPicture(ByVal hBmpSrc As Long, _
                                     ByVal hPalette As Long) As Picture

    Dim PicHold As BmpPicture
    Dim vbPicture As IPicture
    Dim IDisp As DispInterface
    
    ' Fill in with IDispatch Interface ID
    With IDisp
       .Data1 = &H20400
       .Data4(0) = &HC0
       .Data4(7) = &H46
    End With
    
    ' Fill Pic with necessary parts
    With PicHold
       .Size = Len(PicHold)         ' Length of structure
       .Type = vbPicTypeBitmap      ' Type of Picture (bitmap)
       .hBmp = hBmpSrc              ' Handle to bitmap
       .hPal = hPalette             ' Handle to palette (may be null)
    End With
    
    ' Create Picture object
    OleCreatePictureIndirect PicHold, IDisp, 1, vbPicture
    
    ' Return the new Picture object
    Set CreateBitmapPicture = vbPicture
End Function

Public Sub PrintScreen(frmSrc As Form)
    
    Dim nOX1 As Long
    Dim nOY1 As Long
    Dim nOX2 As Long
    Dim nOY2 As Long
    Dim nMargin As Integer
    Dim fFactor As Single
    Dim LEFT_MARGIN As Long
    Dim TOP_MARGIN As Long
        
    SelectPrinter.ShowDialog False, frmSrc
    If SelectPrinter.Canceled Then
        MsgBox "Print operation canceled"
        Exit Sub
    End If
    
    Screen.MousePointer = vbHourglass
    'Get the size of input form
    nOX1 = 0
    nOX2 = frmSrc.Width / Screen.TwipsPerPixelX
    nOY1 = 0
    nOY2 = frmSrc.Height / Screen.TwipsPerPixelY
    
    LEFT_MARGIN = Printer.ScaleWidth / 24
    nMargin = frmSrc.Width - frmSrc.ScaleWidth
    
    'Capture the form
    Set pctHold.Picture = CaptureWindow(frmSrc.hwnd, nOX1, nOY1, nOX2, nOY2)
    
    With Printer
        nOX1 = .ScaleWidth - LEFT_MARGIN * 2
        fFactor = nOX1 / frmSrc.Width
        nOY1 = frmSrc.Height * fFactor
        TOP_MARGIN = (.ScaleHeight - nOY1) / 2
        'Paint the picture on to the printer object and print
        .PaintPicture pctHold.Picture, LEFT_MARGIN, TOP_MARGIN, nOX1, nOY1, 0, 0, frmSrc.Width, frmSrc.Height
        .EndDoc
    End With
    Screen.MousePointer = vbDefault
End Sub



