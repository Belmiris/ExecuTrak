Attribute VB_Name = "Module1"
Option Explicit

Public Sub main()

    If Command = t_szHandShake Then
        Load frmMain "Name of the Main Form"
    Else
        frmSplash.Caption = "Select Data Sources"
    End If
End Sub

Public Sub subShowMainForm()
    frmMain.Show
End Sub

