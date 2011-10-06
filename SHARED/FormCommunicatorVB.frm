VERSION 5.00
Begin VB.Form FormCommunicatorVB 
   ClientHeight    =   3735
   ClientLeft      =   -29940
   ClientTop       =   -29550
   ClientWidth     =   5325
   LinkTopic       =   "Form1"
   ScaleHeight     =   3735
   ScaleWidth      =   5325
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.TextBox TextBox 
      Height          =   495
      Index           =   0
      Left            =   180
      TabIndex        =   0
      Top             =   300
      Width           =   1215
   End
End
Attribute VB_Name = "FormCommunicatorVB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'//////////////////////////////////////////////////////////
' WINDOWS APIs

Private Declare Function IsWindow Lib "USER32" (ByVal hWnd As Long) As Long

'//////////////////////////////////////////////////////////
' CONSTANTS

Private Const windowText As String = "58BE1C53-3EA8-4A9D-8739-1361E078BF29"
Private Const connectRequest As String = "426096F4-DD07-4A81-B924-A1B26F011A23"
Private Const connectAccepted As String = "A97750A9-F259-48DF-B4E1-86F6602EF3BE"
Private Const connectDenied As String = "FAD8C045-7B90-4537-BDB6-0E256266A049"
Private Const messageResponse As String = "8BFEC2E8-AD6C-419C-983A-034AF71A1AA3"

'//////////////////////////////////////////////////////////
' VARIABLES

Private progId As String

Private appInformation() As String
Private connections() As CommunicationConnection
Private nextConnectionIndex As Integer
Private requestedIndex As Integer
Private idSession As String
Private responseReceived As Boolean

Public RequestAccepted As Boolean

'//////////////////////////////////////////////////////////
' EVENTS

Public Event connectRequest(ByRef Accept As Boolean, ApplicationId As String, parameters() As String)

Public Event MessageReceived(sender As String, args As String, response As String)

Public Event messageResponse(sender As String, args As String)


'//////////////////////////////////////////////////////////
' FORM IS LOADING

Private Sub Form_Load()
    ReDim connections(1)
    ReDim appInformation(0, 2)
    
    idSession = NewGUID()
    AddInformation "SESSION", SessionId
End Sub

'//////////////////////////////////////////////////////////
'
' PROPERTIES
'
'//////////////////////////////////////////////////////////

'//////////////////////////////////////////////////////////
' PROGRAM ID PROPERTY

Public Property Get ApplicationId() As String
    ApplicationId = progId
End Property

Public Property Let ApplicationId(ByVal value As String)
    If Len(progId) < 1 Then
        progId = value
        Me.Caption = windowText & progId & "|" & CStr(TextBox(0).hWnd)
    Else
        MsgBox "You cannot change a Communicator ApplicationId once it has been set", vbCritical
    End If
End Property

'//////////////////////////////////////////////////////////
' PROGRAM SESSION PROPERTY

Public Property Get SessionId() As String
    SessionId = idSession
End Property

'//////////////////////////////////////////////////////////
' INFORMATION PROPERTY

Public Property Get InfoItemCount() As Integer
    If Not ArrayIsInitialized(appInformation) Then
        ReDim appInformation(0, 0)
    End If
    InfoItemCount = UBound(appInformation, 1)
End Property


'//////////////////////////////////////////////////////////
'
' EVENTS
'
'//////////////////////////////////////////////////////////

'//////////////////////////////////////////////////////////
' MESSAGE TEXT BOX CHANGED EVENT

Private Sub TextBox_Change(Index As Integer)
    Dim txtBox As TextBox
    Dim idxAccept, idxDenied, idxHwnd, idxRequest, idxResponse As Integer
    Dim x As Integer
    Dim hWnd As Long
    Dim Accept As Boolean
    Dim parameters() As String
    Dim message As String
    Dim response As String
    
    If TextBox(Index).text = "" Then Exit Sub
    
    Set txtBox = TextBox(Index)
    
    message = TextBox(Index).text
    If Not CommArrayFromString(message, parameters) Then Exit Sub
    
    idxAccept = IndexOfKey(parameters, connectAccepted)
    idxDenied = IndexOfKey(parameters, connectDenied)
    idxHwnd = IndexOfKey(parameters, "HWND")
    idxRequest = IndexOfKey(parameters, connectRequest)
    idxResponse = IndexOfKey(parameters, messageResponse)
    
    '//////////////////////////////////////////////////////
    ' Connection Request ?
    
    If (idxRequest > -1) Then
        If idxHwnd > -1 Then
            hWnd = CLng(parameters(idxHwnd, 1))
            If IsWindow(hWnd) Then
                Accept = False
                RaiseEvent connectRequest(Accept, parameters(idxRequest, 1), parameters)
                If Accept Then
                    x = AddConnection(hWnd, parameters(idxRequest, 1), parameters)
                    response = connectAccepted & "=" & progId & "|HWND=" & CStr(connections(x).hwndTextBox)
                    SendCommunication hWnd, response
                End If
            End If
        End If
        GoTo FINISHED
    End If
    
    '//////////////////////////////////////////////////////
    ' Connection Request Accepted ?
    
    If (idxAccept > -1) Then
        If idxHwnd > -1 Then
            hWnd = CLng(parameters(idxHwnd, 1))
            If IsWindow(hWnd) Then
                x = AddConnection(hWnd, parameters(idxAccept, 1), parameters)
                RequestAccepted = True
            End If
        End If
        GoTo FINISHED
    End If
    
    '//////////////////////////////////////////////////////
    ' Response to Message
    
    If (idxResponse > -1) Then
        RaiseEvent messageResponse(parameters(idxResponse, 1), message)
        GoTo FINISHED
    End If
    
    '//////////////////////////////////////////////////////
    ' Message
    
    response = ""
    x = ConnectionIndexFromTextBox(TextBox(Index).hWnd)
    If x > -1 Then
        RaiseEvent MessageReceived(connections(x).progId, message, response)
        If Len(response) > 0 Then
            response = messageResponse + "=" + progId + "|" + response
            SendCommunication connections(x).hwndApplication, response
        End If
        GoTo FINISHED
    End If
    
FINISHED:
    
End Sub

'//////////////////////////////////////////////////////////
' METHODS
'//////////////////////////////////////////////////////////

'//////////////////////////////////////////////////////////
' Add new connection element and give it a textbox

Private Function AddConnection(hwndApplication As Long, progId As String, info() As String) As Integer
    If (nextConnectionIndex >= UBound(connections)) Then
        ReDim Preserve connections(nextConnectionIndex + 10)
    End If
    
    Load TextBox(TextBox.Count)
    
    connections(nextConnectionIndex).Accepted = True
    connections(nextConnectionIndex).progId = progId
    connections(nextConnectionIndex).hwndApplication = hwndApplication
    connections(nextConnectionIndex).hwndTextBox = TextBox(TextBox.Count - 1).hWnd
    connections(nextConnectionIndex).Information = info
    AddConnection = nextConnectionIndex
    
    nextConnectionIndex = nextConnectionIndex + 1
    
End Function

'//////////////////////////////////////////////////////////
' ADD A KEY VALUE PAIR TO INFORMATION

Public Sub AddInformation(key As String, value As String)
    Dim tmparray() As String
    Dim idx As Integer
    Dim i As Integer
    
    idx = IndexOfKey(appInformation, key)
    If idx > -1 Then
        appInformation(idx, 1) = value
    Else
        idx = UBound(appInformation, 1)
        If Len(appInformation(idx, 0)) < 1 Then
            appInformation(idx, 0) = key
            appInformation(idx, 1) = value
            Exit Sub
        End If
        
        ReDim tmparray(idx + 1, 2)
        For i = 0 To idx
            tmparray(i, 0) = appInformation(i, 0)
            tmparray(i, 1) = appInformation(i, 1)
        Next
        
        ReDim appInformation(idx + 1, 2)
        For i = 0 To idx
            appInformation(i, 0) = tmparray(i, 0)
            appInformation(i, 1) = tmparray(i, 1)
        Next
        
        appInformation(idx + 1, 0) = key
        appInformation(idx + 1, 1) = value
    End If
End Sub

'//////////////////////////////////////////////////////////
' GET INDEX OF CONNECTION TO THE PROGRAM WITH THIS NAME

Public Function ConnectionIndexByAppId(progId As String) As Integer
    Dim i As Integer
    ConnectionIndexByAppId = -1
    For i = 0 To UBound(connections)
        If connections(i).progId = progId Then
            If IsWindow(connections(i).hwndApplication) Then
                ConnectionIndexByAppId = i
                Exit For
            End If
        End If
    Next
End Function

'//////////////////////////////////////////////////////////
' SEE AN INFORMATION KEY

Public Function InfoItemKey(idx As Integer) As String
    If idx > UBound(appInformation, 1) Then
        MsgBox "invalid index sent to InfoItemKey", vbCritical
    Else
        InfoItemKey = appInformation(idx, 0)
    End If
End Function

'//////////////////////////////////////////////////////////
' SEE AN INFORMATION VALUE

Public Function InfoItemValue(idx As Integer) As String
    If idx > UBound(appInformation, 1) Then
        MsgBox "invalid index sent to InfoItemValue", vbCritical
    Else
        InfoItemValue = appInformation(idx, 1)
    End If
End Function

'//////////////////////////////////////////////////////////
' GET THE PROGRAM ID ASSIGNED TO THE TEXTBOX HANDLE

Private Function ProgramIdFromTextBox(hWnd As Long) As String
    Dim i As Integer
    
    ProgramIdFromTextBox = -1
    
    For i = 0 To UBound(connections)
        If connections(i).hwndTextBox = hWnd Then
            ProgramIdFromTextBox = connections(i).progId
            Exit For
        End If
    Next
End Function

'//////////////////////////////////////////////////////////
' GET THE INDEX OF THE CONNECTION ASSIGNED TO TEXTBOX HANDLE

Private Function ConnectionIndexFromTextBox(hWnd As Long) As Integer
    Dim i As Integer
    
    For i = 0 To UBound(connections)
        If connections(i).hwndTextBox = hWnd Then
            ConnectionIndexFromTextBox = i
            Exit For
        End If
    Next
End Function

'//////////////////////////////////////////////////////////
' REMOVE A KEY VALUE PAIR TO INFORMATION

Public Sub RemoveInformation(key As String)
    Dim i As Integer
    Dim newSize As Integer
    Dim tmparray() As String
    Dim remake As Boolean
    
    newSize = -1
    ReDim tmparray(UBound(appInformation, 1), 2)
    
    For i = 0 To UBound(appInformation, 1)
        If StrComp(appInformation(i, 0), key, vbTextCompare) <> 0 Then
            newSize = newSize + 1
            tmparray(newSize, 0) = appInformation(i, 0)
            tmparray(newSize, 1) = appInformation(i, 1)
        Else
            remake = True
        End If
    Next
        
    If remake Then
        ReDim appInformation(newSize, 2)
        For i = 0 To newSize
            appInformation(newSize, 0) = tmparray(i, 0)
            appInformation(newSize, 1) = tmparray(i, 1)
        Next
    End If
End Sub

'//////////////////////////////////////////////////////////
' LOCATE A PROGRAM WITH THIS ID AND REQUEST A CONNECTION

Public Function RequestConnectionId(progId As String) As Boolean
    Dim hWnd As Long
    Dim i As Integer
    
    RequestConnectionId = ConnectionIndexByAppId(progId) > -1
    
    If Not RequestConnectionId Then
        If RequestConnection(Me, progId) Then
            RequestConnectionId = ConnectionIndexByAppId(progId) > -1
        End If
    End If
End Function

'//////////////////////////////////////////////////////////
' SEND A MESSAGE TO AN APPLICATION

Public Function SendMessage(progId As String, message As String) As Boolean
    Dim idx As Integer
    
    idx = ConnectionIndexByAppId(progId)
    If (idx < 0) Then Exit Function
    
    If IsWindow(connections(idx).hwndApplication) = 0 Then Exit Function
    SendCommunication connections(idx).hwndApplication, message
    SendMessage = True
End Function
