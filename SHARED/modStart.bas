Attribute VB_Name = "modStart"
Option Explicit

Public Sub main()

    'If Command = t_szHandShake Then
    If tfnAuthorizeExecute(Command, False) Then
        Load frmMain
    Else
        frmSplash.Caption = "Select Data Sources"
        frmSplash.Show vbModal
    End If
End Sub

Public Sub subShowMainForm()
    frmMain.Show
End Sub

