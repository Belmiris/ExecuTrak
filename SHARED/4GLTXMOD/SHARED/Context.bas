Attribute VB_Name = "modContextMenus"
'***********************************************************'
'
' Copyright (c) 1996 FACTOR, A Division of W.R.Hess Company
'
' Module name   : CONTEXT.BAS
' Date          : Feb 21, 1995
' Programmer(s) : Jeffrey Wedekind
'
' This module implements context menu functions
'
' Functions:
'
Option Explicit
'
'Function : tfnShowPopup - shows a context menu
'Variables: pointer to the form, submenu index
'Return   : none
'
Public Sub tfnShowPopup(frmMenuForm As Form, nSubMenu As Integer)
        
    ' This is where to put menu
    Dim nWhereX, nWhereY As Single
    
    'Top Level Menu
    Dim hMenu As Integer
    
    'Sub Menu to popup
    Dim hSubMenu As Integer
    
    'Window rectangle to show popup in
    Dim rctMainWindow As RECT
    
    'Pos of Current Mouse Pointer
    Dim pntPosition As POINTAPI
    
    'A TEmp variable to hold the return value from TrackPopup
    Dim nTemp As Integer
    
    'Get the Mouse pointer position
    Call GetCursorPos(pntPosition)
    
    'Set up the variables to hold the screen size
    'Screen .ScaleMode = TWIPS
    rctMainWindow.Top = 0
    rctMainWindow.Left = 0
    rctMainWindow.Right = Screen.Width \ Screen.TwipsPerPixelX
    rctMainWindow.Bottom = Screen.Height \ Screen.TwipsPerPixelY
    
    ' Don't Put the menu right under the mousepointer
    nWhereX = pntPosition.x + 3
    nWhereY = pntPosition.y + 5
    
    'Get the top level menu
    hMenu = GetMenu(frmMenuForm.hwnd)
    
    'Get the submenu that they want
    hSubMenu = GetSubMenu(hMenu, nSubMenu)
    
    'Popup the menu
    nTemp = TrackPopupMenu(hSubMenu, 2, nWhereX, nWhereY, 0, frmMenuForm.hwnd, rctMainWindow)

End Sub

