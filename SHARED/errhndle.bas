Attribute VB_Name = "modErrorHandler"
Option Explicit

'Error Handler Class constant and object
Public Const NO_LOG = 0
Public Const LOG_TO_FILE = 1
Public Const LOG_TO_DB = 2
Private Const RASTERCAPS = 38
Private Const RC_PALETTE = &H100
Private Const SIZEPALETTE = 104


Public objErrHandler As clsErrorHandler

'screen caption constant, type and function declaration
'''''''''''''''''''''  vvvvvvvvvvvvvvvvvvvvvvvv  '''''''''''''''''''''''''
    Private Const CHUNK_SIZE = 32000        'Size of file transfer pieces
 
    Private Type GUID
       Data1 As Long
       Data2 As Integer
       Data3 As Integer
       Data4(7) As Byte
    End Type
    
    Private Type PALETTEENTRY
       peRed As Byte
       peGreen As Byte
       peBlue As Byte
       peFlags As Byte
    End Type
    
    Private Type LOGPALETTE
       palVersion As Integer
       palNumEntries As Integer
       palPalEntry(255) As PALETTEENTRY  ' Enough for 256 colors
    End Type

#If Win32 Then
     Private Type PicBmp
        Size As Long
        Type As Long
        hBmp As Long
        hPal As Long
        Reserved As Long
     End Type

     Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
     End Type

     Private Declare Function BitBlt Lib "GDI32" ( _
        ByVal hDCDest As Long, _
        ByVal XDest As Long, _
        ByVal YDest As Long, _
        ByVal nWidth As Long, _
        ByVal nHeight As Long, _
        ByVal hDCSrc As Long, _
        ByVal XSrc As Long, _
        ByVal YSrc As Long, _
        ByVal dwRop As Long) As Long
     
     Private Declare Function CreateCompatibleBitmap Lib "GDI32" ( _
        ByVal hDC As Long, _
        ByVal nWidth As Long, _
        ByVal nHeight As Long) As Long
     
     Private Declare Function CreateCompatibleDC Lib "GDI32" ( _
        ByVal hDC As Long) As Long

     Private Declare Function CreatePalette Lib "GDI32" ( _
        lpLogPalette As LOGPALETTE) As Long
     
     Private Declare Function DeleteDC Lib "GDI32" ( _
        ByVal hDC As Long) As Long
        
     Private Declare Function GetDC Lib "user32" ( _
        ByVal hwnd As Long) As Long
        
     Private Declare Function GetDesktopWindow Lib "user32" () As Long

     Private Declare Function GetDeviceCaps Lib "GDI32" ( _
        ByVal hDC As Long, ByVal iCapabilitiy As Long) As Long
        
     Private Declare Function GetSystemPaletteEntries Lib "GDI32" ( _
        ByVal hDC As Long, _
        ByVal wStartIndex As Long, _
        ByVal wNumEntries As Long, _
        lpPaletteEntries As PALETTEENTRY) _
        As Long
        
     Private Declare Function GetWindowDC Lib "user32" ( _
        ByVal hwnd As Long) As Long
     
     Private Declare Function SelectObject Lib "GDI32" ( _
        ByVal hDC As Long, ByVal hObject As Long) As Long
        
     Private Declare Function GetForegroundWindow Lib "user32" () _
        As Long
        
     Private Declare Function GetWindowRect Lib "user32" ( _
        ByVal hwnd As Long, _
        lpRect As RECT) As Long
     
     Private Declare Function OleCreatePictureIndirect Lib "olepro32.dll" ( _
        PicDesc As PicBmp, _
        RefIID As GUID, _
        ByVal fPictureOwnsHandle As Long, _
        IPic As IPicture) As Long
     
     Private Declare Function ReleaseDC Lib "user32" ( _
        ByVal hwnd As Long, ByVal hDC As Long) As Long
     
     Private Declare Function RealizePalette Lib "GDI32" ( _
        ByVal hDC As Long) As Long
        
     Private Declare Function SelectPalette Lib "GDI32" ( _
        ByVal hDC As Long, _
        ByVal hPalette As Long, _
        ByVal bForceBackground As Long) As Long
#Else
    Private Type PicBmp
       Size As Integer
       Type As Integer
       hBmp As Integer
       hPal As Integer
       Reserved As Integer
    End Type
    
    Private Type RECT
       Left As Integer
       Top As Integer
       Right As Integer
       Bottom As Integer
    End Type
    
    Private Declare Function BitBlt Lib "GDI" ( _
        ByVal hDCDest As Integer, _
        ByVal XDest As Integer, _
        ByVal YDest As Integer, _
        ByVal nWidth As Integer, _
        ByVal nHeight As Integer, _
        ByVal hDCSrc As Integer, _
        ByVal XSrc As Integer, _
        ByVal YSrc As Integer, _
        ByVal dwRop As Long) As Integer

    Private Declare Function CreateCompatibleBitmap Lib "GDI" ( _
        ByVal hDC As Integer, _
        ByVal nWidth As Integer, _
        ByVal nHeight As Integer) As Integer
        
    Private Declare Function CreateCompatibleDC Lib "GDI" ( _
        ByVal hDC As Integer) As Integer
        
    Private Declare Function CreatePalette Lib "GDI" ( _
        lpLogPalette As LOGPALETTE) As Integer
        
    Private Declare Function DeleteDC Lib "GDI" ( _
        ByVal hDC As Integer) As Integer

    Private Declare Function GetDC Lib "USER" ( _
        ByVal hwnd As Integer) As Integer

    Private Declare Function GetDesktopWindow Lib "USER" () As Integer
    
    Private Declare Function GetDeviceCaps Lib "GDI" ( _
        ByVal hDC As Integer, _
        ByVal iCapabilitiy As Integer) As Integer
        
    Private Declare Function GetSystemPaletteEntries Lib "GDI" ( _
        ByVal hDC As Integer, _
        ByVal wStartIndex As Integer, _
        ByVal wNumEntries As Integer, _
        lpPaletteEntries As PALETTEENTRY) As Integer
        
    Private Declare Function GetWindowDC Lib "USER" ( _
        ByVal hwnd As Integer) As Integer

    Private Declare Function SelectObject Lib "GDI" ( _
        ByVal hDC As Integer, _
        ByVal hObject As Integer) As Integer
        
    Private Declare Function GetForegroundWindow Lib "USER" _
        Alias "GetActiveWindow" () As Integer
        
    Private Declare Function GetWindowRect Lib "USER" ( _
        ByVal hwnd As Integer, _
        lpRect As RECT) As Integer

    Private Declare Function OleCreatePictureIndirect Lib "oc25.dll" ( _
       PictDesc As PicBmp, _
       RefIID As GUID, _
       ByVal fPictureOwnsHandle As Integer, _
       IPic As IPicture) As Integer
    
    Private Declare Function ReleaseDC Lib "USER" ( _
        ByVal hwnd As Integer, _
        ByVal hDC As Integer) As Integer

    Private Declare Function RealizePalette Lib "USER" ( _
        ByVal hDC As Integer) As Integer

    Private Declare Function SelectPalette Lib "USER" ( _
        ByVal hDC As Integer, _
        ByVal hPalette As Integer, ByVal _
        bForceBackground As Integer) As Integer
#End If
'''''''''''''''''''''  ^^^^^^^^^^^^^^^^^^^^^^^   '''''''''''''''''''''''''

Public Function tfnErrHandler(sProcName As String, _
                         Optional vSQL As Variant, _
                         Optional vExecProcMessage As Variant, _
                         Optional vShowMessage As Variant) As Boolean
    
    If Not objErrHandler Is Nothing Then
        If IsMissing(vSQL) Then
            tfnErrHandler = objErrHandler.ErrorHandler(sProcName)
        Else
            If IsMissing(vExecProcMessage) Then
                tfnErrHandler = objErrHandler.ErrorHandler(sProcName, vSQL)
            Else
                If IsMissing(vShowMessage) Then
                    tfnErrHandler = objErrHandler.ErrorHandler(sProcName, vSQL, vExecProcMessage)
                Else
                    tfnErrHandler = objErrHandler.ErrorHandler(sProcName, vSQL, vExecProcMessage, vShowMessage)
                End If
            End If
        End If
    End If
End Function

'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' CreateBitmapPicture
'    - Creates a bitmap type Picture object from a bitmap and palette
'
' hBmp
'    - Handle to a bitmap
'
' hPal
'    - Handle to a Palette
'    - Can be null if the bitmap doesn't use a palette
'
' Returns
'    - Returns a Picture object containing the bitmap
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
#If Win32 Then
Private Function CreateBitmapPicture(ByVal hBmp As Long, _
   ByVal hPal As Long) As Picture

    Dim r As Integer
#Else
Private Function CreateBitmapPicture(ByVal hBmp As Integer, _
   ByVal hPal As Integer) As Picture

    Dim r As Integer
#End If
    Dim Pic As PicBmp
    ' IPicture requires a reference to "Standard OLE Types"
    Dim IPic As IPicture
    Dim IID_IDispatch As GUID
    
   ' Fill in with IDispatch Interface ID
   With IID_IDispatch
      .Data1 = &H20400
      .Data4(0) = &HC0
      .Data4(7) = &H46
   End With
 
   ' Fill Pic with necessary parts
   With Pic
      .Size = Len(Pic)          ' Length of structure
      .Type = vbPicTypeBitmap   ' Type of Picture (bitmap)
      .hBmp = hBmp              ' Handle to bitmap
      .hPal = hPal              ' Handle to palette (may be null)
   End With
 
   ' Create Picture object
   r = OleCreatePictureIndirect(Pic, IID_IDispatch, 1, IPic)
 
   ' Return the new Picture object
   Set CreateBitmapPicture = IPic
End Function
 
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' CaptureWindow
'    - Captures any portion of a window
'
' hWndSrc
'    - Handle to the window to be captured
'
' Client
'    - If True CaptureWindow captures from the client area of the window
'    - If False CaptureWindow captures from the entire window
'
' LeftSrc, TopSrc, WidthSrc, HeightSrc
'    - Specify the portion of the window to capture
'    - Dimensions need to be specified in pixels
'
' Returns
'    - Returns a Picture object containing a bitmap of the specified
'      portion of the window that was captured
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
#If Win32 Then
Private Function CaptureWindow(ByVal hWndSrc As Long, _
   ByVal Client As Boolean, ByVal LeftSrc As Long, _
   ByVal TopSrc As Long, ByVal WidthSrc As Long, _
   ByVal HeightSrc As Long) As Picture

    Dim hDCMemory As Long
    Dim hBmp As Long
    Dim hBmpPrev As Long
    Dim r As Long
    Dim hDCSrc As Long
    Dim hPal As Long
    Dim hPalPrev As Long
    Dim RasterCapsScrn As Long
    Dim HasPaletteScrn As Long
    Dim PaletteSizeScrn As Long
    Dim LogPal As LOGPALETTE
#Else
Private Function CaptureWindow(ByVal hWndSrc As Integer, _
   ByVal Client As Boolean, ByVal LeftSrc As Integer, _
   ByVal TopSrc As Integer, ByVal WidthSrc As Long, _
   ByVal HeightSrc As Long) As Picture
   
    Dim hDCMemory As Integer
    Dim hBmp As Integer
    Dim hBmpPrev As Integer
    Dim r As Integer
    Dim hDCSrc As Integer
    Dim hPal As Integer
    Dim hPalPrev As Integer
    Dim RasterCapsScrn As Integer
    Dim HasPaletteScrn As Integer
    Dim PaletteSizeScrn As Integer
    Dim LogPal As LOGPALETTE
#End If
   ' Depending on the value of Client get the proper device context
   If Client Then
      hDCSrc = GetDC(hWndSrc) ' Get device context for client area
   Else
      hDCSrc = GetWindowDC(hWndSrc) ' Get device context for entire window
   End If
 
   ' Create a memory device context for the copy process
   hDCMemory = CreateCompatibleDC(hDCSrc)
   ' Create a bitmap and place it in the memory DC
   hBmp = CreateCompatibleBitmap(hDCSrc, WidthSrc, HeightSrc)
   hBmpPrev = SelectObject(hDCMemory, hBmp)
 
   ' Get screen properties
   RasterCapsScrn = GetDeviceCaps(hDCSrc, RASTERCAPS) ' Raster capabilities
   HasPaletteScrn = RasterCapsScrn And RC_PALETTE       ' Palette support
   PaletteSizeScrn = GetDeviceCaps(hDCSrc, SIZEPALETTE) ' Size of palette
 
   ' If the screen has a palette make a copy and realize it
   If HasPaletteScrn And (PaletteSizeScrn = 256) Then
      ' Create a copy of the system palette
      LogPal.palVersion = &H300
      LogPal.palNumEntries = 256
      r = GetSystemPaletteEntries(hDCSrc, 0, 256, LogPal.palPalEntry(0))
      hPal = CreatePalette(LogPal)
      ' Select the new palette into the memory DC and realize it
      hPalPrev = SelectPalette(hDCMemory, hPal, 0)
      r = RealizePalette(hDCMemory)
   End If
 
   ' Copy the on-screen image into the memory DC
   r = BitBlt(hDCMemory, 0, 0, WidthSrc, HeightSrc, hDCSrc, _
      LeftSrc, TopSrc, vbSrcCopy)
 
   ' Remove the new copy of the the on-screen image
   hBmp = SelectObject(hDCMemory, hBmpPrev)
 
   ' If the screen has a palette get back the palette that was selected
   ' in previously
   If HasPaletteScrn And (PaletteSizeScrn = 256) Then
      hPal = SelectPalette(hDCMemory, hPalPrev, 0)
   End If
 
   ' Release the device context resources back to the system
   r = DeleteDC(hDCMemory)
   r = ReleaseDC(hWndSrc, hDCSrc)
 
   ' Call CreateBitmapPicture to create a picture object from the bitmap
   ' and palette handles.  Then return the resulting picture object.
   Set CaptureWindow = CreateBitmapPicture(hBmp, hPal)
End Function
 
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' CaptureForm
'    - Captures an entire form including title bar and border
'
' frmSrc
'    - The Form object to capture
'
' Returns
'    - Returns a Picture object containing a bitmap of the entire form
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
Public Function CaptureForm(frmSrc As Form) As Picture
   ' Call CaptureWindow to capture the entire form given it's window
   ' handle and then return the resulting Picture object
   Set CaptureForm = CaptureWindow(frmSrc.hwnd, False, 0, 0, _
      frmSrc.ScaleX(frmSrc.Width, vbTwips, vbPixels), _
      frmSrc.ScaleY(frmSrc.Height, vbTwips, vbPixels))
End Function
'--------------------------------------------------------------------------

'Extracts data from field and places it into a file stream
'
Public Function FieldToFileStream(iFileNumber As Integer, fdObject As Field) As Boolean
   Dim sChunkHolder As String
   Dim lChunkCount As Long
   Dim lChunkRemainder As Long
   Dim i As Long
   
   lChunkCount = fdObject.FieldSize() \ CHUNK_SIZE
   lChunkRemainder = fdObject.FieldSize() Mod CHUNK_SIZE
   For i = 0 To lChunkCount - 1
      sChunkHolder = fdObject.GetChunk(i * CHUNK_SIZE, CHUNK_SIZE)
      Put iFileNumber, , sChunkHolder
   Next
   If lChunkRemainder > 0 Then
      sChunkHolder = fdObject.GetChunk(lChunkCount * CHUNK_SIZE, lChunkRemainder) 'Get remaining data
      Put iFileNumber, , sChunkHolder
   End If
   FieldToFileStream = True  'Can be modified to return errors
End Function

'Copies the remaining portion of an open file stream to a
'database field
'
Public Function FileStreamToField(iFileNumber As Integer, fdObject As Field) As Boolean
   Dim sChunkHolder As String
   Dim lChunkCount As Long
   Dim lChunkRemainder As Long
   Dim i As Long
   
   sChunkHolder = Space$(CHUNK_SIZE)
   lChunkCount = (LOF(iFileNumber) - Seek(iFileNumber) + 1) \ CHUNK_SIZE
   lChunkRemainder = (LOF(iFileNumber) - Seek(iFileNumber) + 1) Mod CHUNK_SIZE
   For i = 1 To lChunkCount
      Get iFileNumber, , sChunkHolder
      fdObject.AppendChunk (sChunkHolder)
   Next
   If lChunkRemainder > 0 Then
      sChunkHolder = Space$(lChunkRemainder)
      Get iFileNumber, , sChunkHolder
      fdObject.AppendChunk (sChunkHolder)
   End If
   FileStreamToField = True           'Can be modified to return errors
End Function


