Attribute VB_Name = "modMainEdit"
'***********************************************************'
'
' Copyright (c) 1997 FACTOR, A Division of W.R.Hess Company
'
' Module name   : ModMainEdit
'
' Programmer: Weigong Jiang
'
' Date: 12/09/97
' Note: We use Sendmessage to implement,instead of using
' textbox.seltext, cut ,paste and copy. The reason is that we
' may want undo
' Purpose: Since the Cut,Copy and Paste are on the Main edit menu
' of Template and the implementation of Cut,Copy and Paste is not
' correct(although easy to fix). I give this simple module to
' provide an easy fix.
'
' Usage:
' (a) In mnuMainEdit (template) Call
'           ToggleCopy mnuCopy
'           ToggleCut mnuCut
'           TogglePaste mnuPaste
'           Toggleundo mnuundo (if available)
' (b) In mnuCut_click Call Cut_click
'     In paste_click Call Cut_click
'     In copy_click Call Copy_click
'
Option Explicit

'
' Windows API call used to control textbox
'
#If Win16 Then
   Private Declare Function SendMessage Lib "User" (ByVal hWnd As Integer, ByVal wMsg As Integer, ByVal wParam As Integer, lParam As Any) As Long
#ElseIf Win32 Then
   Private Declare Function SendMessage Lib "USER32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
#End If
'
' Edit Control Messages
'
Private Const WM_CUT = &H300
Private Const WM_COPY = &H301
Private Const WM_PASTE = &H302
Private Const WM_CLEAR = &H303
Private Const WM_UNDO = &H304

#If Win16 Then
   Const EM_CANUNDO = &H416     'WM_USER + 22
   Const EM_GETMODIFY = &H408   'WM_USER + 8
#ElseIf Win32 Then
   Const EM_CANUNDO = &HC6
   Const EM_GETMODIFY = &HB8
#End If

Public Function IsTextChanged(txt As Textbox) As Boolean
    If SendMessage(txt.hWnd, EM_GETMODIFY, 0, 0&) Then
        IsTextChanged = True
    Else
        IsTextChanged = False
    End If
End Function

Public Sub Undo_Click()
    subPerformEdit WM_UNDO
End Sub
Public Sub Cut_Click()
    subPerformEdit WM_CUT
End Sub
Public Sub Copy_Click()
    subPerformEdit WM_COPY
End Sub
Public Sub Paste_Click()
    subPerformEdit WM_PASTE
End Sub
Public Sub Delete_Click()
    subPerformEdit WM_CLEAR
End Sub
Public Sub ToggleCut(mnu As Control)
   
   If TypeOf Screen.ActiveControl Is Textbox Then
       If Screen.ActiveControl.Locked Then
           mnu.Enabled = False
       Else
           mnu.Enabled = (Screen.ActiveControl.SelLength > 0)
       End If
   Else
      mnu.Enabled = False
   End If
End Sub
Public Sub ToggleUndo(mnu As Control)
   
   If TypeOf Screen.ActiveControl Is Textbox Then
       If Screen.ActiveControl.Locked Then
          mnu.Enabled = False
       Else
          mnu.Enabled = SendMessage(Screen.ActiveControl.hWnd, EM_CANUNDO, 0, 0&)
       End If
   Else
      mnu.Enabled = False
   End If
End Sub
Public Sub TogglePaste(mnu As Control)
   If TypeOf Screen.ActiveControl Is Textbox Then
       If Screen.ActiveControl.Locked Then
          mnu.Enabled = False
       Else
          mnu.Enabled = Clipboard.GetFormat(vbCFText)
       End If
   Else
      mnu.Enabled = False
   End If
End Sub
Public Sub ToggleCopy(mnu As Control)
   
   If TypeOf Screen.ActiveControl Is Textbox Then
     ' If Screen.ActiveControl.Locked Then
     '    mnu.Enabled = False
     ' Else
         mnu.Enabled = (Screen.ActiveControl.SelLength > 0)
     ' End If
   Else
      mnu.Enabled = False
   End If
End Sub
Public Sub ToggleDelete(mnu As Control)
   
   If TypeOf Screen.ActiveControl Is Textbox Then
      If Screen.ActiveControl.Locked Then
         mnu.Enabled = False
      Else
         mnu.Enabled = (Screen.ActiveControl.SelLength > 0)
      End If
   Else
      mnu.Enabled = False
   End If
End Sub

Private Sub subPerformEdit(ByVal nIndex As Long)
   If TypeOf Screen.ActiveControl Is Textbox Then
      If nIndex = WM_COPY Then
            Call SendMessage(Screen.ActiveControl.hWnd, nIndex, 0, 0&)
      Else
            If Screen.ActiveControl.Locked Then
               Beep
            Else
              Call SendMessage(Screen.ActiveControl.hWnd, nIndex, 0, 0&)
            End If
      End If
   Else
      Beep
   End If
End Sub



