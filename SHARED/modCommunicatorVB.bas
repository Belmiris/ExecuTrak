Attribute VB_Name = "modCommunicatorVB"
Option Explicit

'//////////////////////////////////////////////////////////
' WINDOWS APIs
Private Declare Function CoCreateGuid Lib "OLE32.DLL" (pGuid As GUID) As Long

Private Declare Function EnumWindows Lib "user32.dll" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long

Private Declare Function GetWindowText Lib "user32.dll" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long

Private Declare Function GetWindowTextLength Lib "USER32" Alias "GetWindowTextLengthA" (ByVal hWnd As Long) As Long

Private Declare Function FindWindow Lib "USER32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long

Private Declare Function IsWindow Lib "USER32" (ByVal hWnd As Long) As Long

Private Declare Function SendMessage Lib "USER32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As String) As Long

Private Const WM_SETTEXT As Long = &HC

'//////////////////////////////////////////////////////////
' STRUCTURES
Public Type CommunicationConnection
    Accepted As Boolean
    Information() As String
    hwndApplication As Long
    hwndTextBox As Long
    progId As String
End Type

Public Type CommunicatorConnectRequestEventArgs
    FromProgId As String
    Accept As Boolean
End Type

Public Type CommunicatorSendMessageEventArgs
    messages() As String
End Type

Public Type CommunicatorReceiveEventArgs
    FromProgId As String
    messages() As String
    Response() As String
End Type

Public Type CommunicatorReceivedResponseEventArgs
    FromProgId As String
    Responses() As String
End Type

Private Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(7) As Byte
End Type

'//////////////////////////////////////////////////////////
' CONSTANTS
Private Const windowText As String = "58BE1C53-3EA8-4A9D-8739-1361E078BF29"
Private Const connectRequest As String = "426096F4-DD07-4A81-B924-A1B26F011A23"
Private Const connectAccepted As String = "A97750A9-F259-48DF-B4E1-86F6602EF3BE"
Private Const connectDenied As String = "FAD8C045-7B90-4537-BDB6-0E256266A049"
Private Const messageResponse As String = "8BFEC2E8-AD6C-419C-983A-034AF71A1AA3"

'//////////////////////////////////////////////////////////
' VARIABLES

Private m_formCommunicator As FormCommunicatorVB
Private windowSearch As String
Private windowsFound As Collection
Private connectionAccepted As Boolean

'//////////////////////////////////////////////////////////
' DETERMINE WHETHER OR NOT AN ARRAY IS INITIALIZED
Public Function ArrayIsInitialized(someArray() As String) As Boolean
    If ((Not someArray) = -1) Then
        ArrayIsInitialized = False
    Else
        ArrayIsInitialized = True
    End If
End Function

'//////////////////////////////////////////////////////////
' SEARCH FOR WINDOW WITH TEXT STARTING WITH windowSearch
Public Function EnumWindowsProc(ByVal hWnd As Long, ByVal lParam As Long) As Long
    Dim wlength, x, l As Long
    Dim hwndApp As Long
    Dim buffer As String
    Dim test As String
    Dim arr() As String
    
    l = Len(windowSearch)
    
    wlength = GetWindowTextLength(hWnd) + 1
    If wlength >= l Then
        buffer = Space$(wlength)
        x = GetWindowText(hWnd, buffer, wlength)
        test = Mid$(buffer, 1, l)
        If StrComp(windowSearch, test) = 0 Then
            arr = Split(buffer, "|")
            hwndApp = CLng(arr(1))
            Call windowsFound.Add(hwndApp)
        End If
    End If
    
    EnumWindowsProc = True
    
End Function

'//////////////////////////////////////////////////////////
' TURN A PIPE DELIMITED NAME-VALUE PAIR INTO A 2D ARRAY
Public Function CommArrayFromString(message As String, messages() As String) As Boolean
    Dim hold As New Collection
    Dim arr() As String
    Dim messageCount As Integer
    Dim e, i As Integer
    
    arr = Split(message, "|")
    messageCount = UBound(arr)
    
    For i = 0 To messageCount
        e = InStr(1, arr(i), "=")
        If e > 0 Then
            hold.Add arr(i)
        End If
    Next
    
    If hold.Count < 1 Then Exit Function
    
    ReDim messages(hold.Count - 1, 2)
    
    For i = 1 To hold.Count
        e = InStr(1, hold(i), "=")
        messages(i - 1, 0) = Mid$(hold(i), 1, e - 1)
        messages(i - 1, 1) = Mid$(hold(i), e + 1)
    Next
    
    CommArrayFromString = True
    
End Function

'//////////////////////////////////////////////////////////
' Get the position of the 'key' in a 2D string array
Public Function IndexOfKey(array2d() As String, Key As String) As Integer
    Dim i As Integer
    
    IndexOfKey = -1
    For i = LBound(array2d, 1) To UBound(array2d, 1)
        If StrComp(Key, array2d(i, 0)) = 0 Then
            IndexOfKey = i
            Exit Function
        End If
    Next
    
End Function

'//////////////////////////////////////////////////////////
' CREATE A GUID: (c) 2000 Gus Molina
Public Function NewGUID() As String
    Dim udtGUID As GUID
    
    If (CoCreateGuid(udtGUID) = 0) Then
        NewGUID = _
            String(8 - Len(Hex$(udtGUID.Data1)), "0") & Hex$(udtGUID.Data1) & _
            String(4 - Len(Hex$(udtGUID.Data2)), "0") & Hex$(udtGUID.Data2) & _
            String(4 - Len(Hex$(udtGUID.Data3)), "0") & Hex$(udtGUID.Data3) & _
            IIf((udtGUID.Data4(0) < &H10), "0", "") & Hex$(udtGUID.Data4(0)) & _
            IIf((udtGUID.Data4(1) < &H10), "0", "") & Hex$(udtGUID.Data4(1)) & _
            IIf((udtGUID.Data4(2) < &H10), "0", "") & Hex$(udtGUID.Data4(2)) & _
            IIf((udtGUID.Data4(3) < &H10), "0", "") & Hex$(udtGUID.Data4(3)) & _
            IIf((udtGUID.Data4(4) < &H10), "0", "") & Hex$(udtGUID.Data4(4)) & _
            IIf((udtGUID.Data4(5) < &H10), "0", "") & Hex$(udtGUID.Data4(5)) & _
            IIf((udtGUID.Data4(6) < &H10), "0", "") & Hex$(udtGUID.Data4(6)) & _
            IIf((udtGUID.Data4(7) < &H10), "0", "") & Hex$(udtGUID.Data4(7))
    End If
End Function

'//////////////////////////////////////////////////////////
' INITIALIZE A COMMUNICATOR FORM
Public Function OpenCommunicator(frm As FormCommunicatorVB, progId As String) As FormCommunicatorVB
    Call CloseCommunicator(frm)
    Set frm = New FormCommunicatorVB
    Load frm
    frm.ApplicationId = progId
    Set OpenCommunicator = frm
End Function

'//////////////////////////////////////////////////////////
' SHUT DOWN A COMMUNICATOR FORM
Public Sub CloseCommunicator(frm As FormCommunicatorVB)
    If Not frm Is Nothing Then
        On Error Resume Next
        Unload frm
        Set frm = Nothing
        On Error GoTo 0
    End If
End Sub

'//////////////////////////////////////////////////////////
' LOOK FOR APPS WITH PROGID AND ASK THEM FOR A CONNECTION
Public Function RequestConnection(frm As FormCommunicatorVB, progId As String) As Boolean
    Dim lParam As Long
    Dim hWnd As Long
    Dim i As Integer
    Dim message As String
    
    Set m_formCommunicator = frm
    m_formCommunicator.RequestAccepted = False
    
    Set windowsFound = New Collection
    windowSearch = windowText + progId + "|"
    Call EnumWindows(AddressOf EnumWindowsProc, 0)
    
    Set m_formCommunicator = Nothing
    
    message = connectRequest & "=" & frm.ApplicationId & "|HWND=" & CStr(frm.TextBox(0).hWnd)
    
    For i = 0 To frm.InfoItemCount
        message = message & "|" + frm.InfoItemKey(i) + "=" + frm.InfoItemValue(i)
    Next
    
    For i = 1 To windowsFound.Count
        hWnd = windowsFound(i)
        SendMessage hWnd, WM_SETTEXT, 0, message
        
        If frm.RequestAccepted Then
            RequestConnection = True
            Exit For
        End If
    Next
    
End Function

'//////////////////////////////////////////////////////////
' Send a message to a window
Public Sub SendCommunication(hWnd As Long, message As String)

    SendMessage hWnd, WM_SETTEXT, 0, message
    
End Sub
