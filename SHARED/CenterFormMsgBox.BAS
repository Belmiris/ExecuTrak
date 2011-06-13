Attribute VB_Name = "modCenterFormMsgBox"
Option Explicit

Private m_Form1 As Form

'misc API constants
Private Const WH_CBT = 5
Private Const GWL_HINSTANCE = (-6)
Private Const HCBT_ACTIVATE = 5

'UDT for passing data through the hook
Private Type MSGBOX_HOOK_PARAMS
   hwndOwner   As Long
   hHook       As Long
End Type

Private Type RECT
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type

'need this declared at module level as
'it is used in the call and the hook proc
Private mhp As MSGBOX_HOOK_PARAMS

Private Declare Function GetCurrentThreadId Lib "kernel32" () As Long

Private Declare Function GetWindowLong Lib "user32" _
   Alias "GetWindowLongA" _
  (ByVal hwnd As Long, _
   ByVal nIndex As Long) As Long

Private Declare Function MessageBox Lib "user32" _
   Alias "MessageBoxA" _
  (ByVal hwnd As Long, _
   ByVal lpText As String, _
   ByVal lpCaption As String, _
   ByVal wType As Long) As Long
   
Private Declare Function SetWindowsHookEx Lib "user32" _
   Alias "SetWindowsHookExA" _
  (ByVal idHook As Long, _
   ByVal lpfn As Long, _
   ByVal hmod As Long, _
   ByVal dwThreadId As Long) As Long
   
Private Declare Function UnhookWindowsHookEx Lib "user32" _
   (ByVal hHook As Long) As Long

Private Declare Function MoveWindow Lib "user32" _
  (ByVal hwnd As Long, _
   ByVal X As Long, _
   ByVal Y As Long, _
   ByVal nWidth As Long, _
   ByVal nHeight As Long, _
   ByVal bRepaint As Long) As Long
   
Private Declare Function GetWindowRect Lib "user32" _
  (ByVal hwnd As Long, _
   lpRect As RECT) As Long

Private Const MB_USERICON = &H80&

Private Type MsgBoxParams
    cbSize As Long
    hwndOwner As Long
    hInstance As Long
    lpszText As Long
    lpszCaption As Long
    dwStyle As Long
    lpszIcon As Long
    dwContextHelpId As Long
    lpfnMsgBoxCallback As Long
    dwLanguageId As Long
End Type

Private Declare Function MessageBoxIndirectW Lib "user32" (lpMsgBoxParams As MsgBoxParams) As Long
'

' note: I didn't bother to go ahead and start hacking with HelpFile and Context
Public Function MsgBox(ByVal Prompt As String, Optional ByVal Buttons As VbMsgBoxStyle = vbOKOnly, _
                       Optional ByVal Title As String, Optional ByVal ResourceIcon As String, _
                       Optional Form1 As Form = Nothing, Optional ByVal nMode As VBRUN.FormShowConstants = vbModal, _
                       Optional ByVal bCenterForm As Boolean = True) _
                       As VbMsgBoxResult
    Dim hwndOwner As Long
    Dim udtMsgBox As MsgBoxParams
    
    Set m_Form1 = Form1
    
    If m_Form1 Is Nothing Then Set m_Form1 = Screen.ActiveForm
    ' if no owner is specified, try to use the active form
    If hwndOwner = 0 Then If Not m_Form1 Is Nothing Then hwndOwner = m_Form1.hwnd
    
    With udtMsgBox
        .cbSize = Len(udtMsgBox)
        ' important to set owner to get behavior similar to the native MsgBox
        If nMode = vbModeless Then
            ' not a modal message box, thus no owner
            .hwndOwner = 0
        Else
            .hwndOwner = hwndOwner
        End If
        
        .hInstance = App.hInstance
        ' set the message
        .lpszText = StrPtr(Prompt)
        ' if no title is given, use the application title like the native MsgBox
        
        If LenB(Title) = 0 Then Title = App.Title
        .lpszCaption = StrPtr(Title)
        
        ' thought this would be a nice feature addition
        If LenB(ResourceIcon) = 0& Then
            .dwStyle = Buttons
        Else
            .dwStyle = (Buttons Or MB_USERICON) And Not (&H70&)
            .lpszIcon = StrPtr(ResourceIcon)
        End If
        
        If bCenterForm Then
            If hwndOwner = 0 And (Not Screen.ActiveForm Is Nothing) Then
                hwndOwner = Screen.ActiveForm.hwnd
            End If
            
            If hwndOwner <> 0 Then
                Dim hInstance As Long
                Dim hThreadId As Long
                
                'Set up the hook
                hInstance = GetWindowLong(hwndOwner, GWL_HINSTANCE)
                hThreadId = GetCurrentThreadId()
                
                'set up the MSGBOX_HOOK_PARAMS values
                'By specifying a Windows hook as one
                'of the params, we can intercept messages
                'sent by Windows and thereby manipulate
                'the dialog
                With mhp
                    .hwndOwner = hwndOwner
                    .hHook = SetWindowsHookEx(WH_CBT, _
                                              AddressOf MsgBoxHookProc, _
                                              hInstance, hThreadId)
                End With
            End If
        End If
    End With
    
    ' show the message box
    MsgBox = MessageBoxIndirectW(udtMsgBox)
End Function

Public Function MsgBoxHookProc(ByVal uMsg As Long, _
                               ByVal wParam As Long, _
                               ByVal lParam As Long) As Long
      
   Dim rc As RECT
   
  'temporary vars for demo
   Dim newLeft As Long
   Dim newTop As Long
   Dim dlgWidth As Long
   Dim dlgHeight As Long
   Dim scrWidth As Long
   Dim scrHeight As Long
   Dim frmLeft As Long
   Dim frmTop As Long
   Dim frmWidth As Long
   Dim frmHeight As Long
   Dim hwndMsgBox As Long
   
  'When the message box is about to be shown,
  'centre the dialog
   If uMsg = HCBT_ACTIVATE Then
   
     'in a HCBT_ACTIVATE message, wParam holds
     'the handle to the messagebox
      hwndMsgBox = wParam
              
     'Just as was done in other API hook demos,
     'position the dialog centered in the calling
     'parent form
     
      Call GetWindowRect(hwndMsgBox, rc)
      
      frmLeft = m_Form1.Left \ Screen.TwipsPerPixelX
      frmTop = m_Form1.Top \ Screen.TwipsPerPixelY
      frmWidth = m_Form1.Width \ Screen.TwipsPerPixelX
      frmHeight = m_Form1.Height \ Screen.TwipsPerPixelX

      dlgWidth = rc.Right - rc.Left
      dlgHeight = rc.Bottom - rc.Top
      
      scrWidth = Screen.Width \ Screen.TwipsPerPixelX
      scrHeight = Screen.Height \ Screen.TwipsPerPixelY
      
      newLeft = frmLeft + ((frmWidth - dlgWidth) \ 2)
      newTop = frmTop + ((frmHeight - dlgHeight) \ 2)
      
      Call MoveWindow(hwndMsgBox, newLeft, newTop, dlgWidth, dlgHeight, True)
      
      
     'done with the dialog so release the hook
      UnhookWindowsHookEx mhp.hHook
               
   End If
   
  'return False to let normal
  'processing continue
   MsgBoxHookProc = False

End Function


