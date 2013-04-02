Attribute VB_Name = "modFRCMDUtil"
'#######################################################################
'#Module   : modFRCMDUtil
'#Programmer : Robert Atwood
'#Date       : 07/23/04
'#Magic#     : 411480
'#Req's      : FRCMD component loaded on a form
'#Description: This module is a container for useful functions that take
'#             advantage of the RCMD component (Remote Command Protocol
'#             OCX in the component picker).  Just drop it on a form
'#             and you're
'#             All functions should call the setup subroutine prior to
'#             executing
'#######################################################################

Public Sub subSetupFRCMD(ctlFRCMD As FRCMD)
    If ctlFRCMD.ConnectStr = "" Then
        ctlFRCMD.ConnectStr = t_dbMainDatabase.Connect
    End If
    ctlFRCMD.Is4GECommand = False
End Sub

'#######################################################################
'# Function fnGetServerDateAndTime
'# Argument     : NONE
'# Magic        : 441678
'# Programmer   : Robert Atwood
'# Description  : Retrieves server Date and time using date command
'#######################################################################
Public Function fnGetServerDateAndTime(ctlFRCMD As FRCMD) As Date
    Dim sDate As String
    ctlFRCMD.Is4GECommand = False
    On Error GoTo ErrOut
    'Vijaya on 02/13/07 #543523 Changed from "Date +%D" to "date +'%D %H:%M:%S'"
    sDate = ctlFRCMD.Execute("date +'%D %H:%M:%S'")
    If IsDate(sDate) Then
        fnGetServerDateAndTime = sDate
    Else
        GoTo ErrOut
    End If
    fnGetServerDateAndTime = sDate
    Exit Function
ErrOut:
    MsgBox "Unable to retrieve time from database server.  Using local system time", vbOKOnly, "AWCRTINT"
    fnGetServerDateAndTime = Now
    On Error GoTo 0
End Function
