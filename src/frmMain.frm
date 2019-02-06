VERSION 5.00
Object = "{33101C00-75C3-11CF-A8A0-444553540000}#1.0#0"; "cswsk32.ocx"
Begin VB.Form frmMain 
   Caption         =   "IMSCROSSCOMM V3.0"
   ClientHeight    =   3585
   ClientLeft      =   270
   ClientTop       =   660
   ClientWidth     =   5265
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3585
   ScaleWidth      =   5265
   StartUpPosition =   1  'CenterOwner
   Begin SocketWrenchCtrl.Socket sckSend 
      Left            =   2160
      Top             =   3120
      _Version        =   65536
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      AutoResolve     =   -1  'True
      Backlog         =   5
      Binary          =   -1  'True
      Blocking        =   -1  'True
      Broadcast       =   0   'False
      BufferSize      =   0
      HostAddress     =   ""
      HostFile        =   ""
      HostName        =   ""
      InLine          =   0   'False
      Interval        =   0
      KeepAlive       =   0   'False
      Library         =   ""
      Linger          =   0
      LocalPort       =   0
      LocalService    =   ""
      Protocol        =   0
      RemotePort      =   0
      RemoteService   =   ""
      ReuseAddress    =   0   'False
      Route           =   -1  'True
      Timeout         =   0
      Type            =   1
      Urgent          =   0   'False
   End
   Begin SocketWrenchCtrl.Socket sockServer 
      Index           =   0
      Left            =   3360
      Top             =   3120
      _Version        =   65536
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      AutoResolve     =   -1  'True
      Backlog         =   5
      Binary          =   -1  'True
      Blocking        =   -1  'True
      Broadcast       =   0   'False
      BufferSize      =   0
      HostAddress     =   ""
      HostFile        =   ""
      HostName        =   ""
      InLine          =   0   'False
      Interval        =   0
      KeepAlive       =   0   'False
      Library         =   ""
      Linger          =   0
      LocalPort       =   0
      LocalService    =   ""
      Protocol        =   0
      RemotePort      =   0
      RemoteService   =   ""
      ReuseAddress    =   0   'False
      Route           =   -1  'True
      Timeout         =   0
      Type            =   1
      Urgent          =   0   'False
   End
   Begin SocketWrenchCtrl.Socket sckReceiveBroadcast 
      Left            =   1560
      Top             =   3120
      _Version        =   65536
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      AutoResolve     =   -1  'True
      Backlog         =   5
      Binary          =   -1  'True
      Blocking        =   -1  'True
      Broadcast       =   0   'False
      BufferSize      =   0
      HostAddress     =   ""
      HostFile        =   ""
      HostName        =   ""
      InLine          =   0   'False
      Interval        =   0
      KeepAlive       =   0   'False
      Library         =   ""
      Linger          =   0
      LocalPort       =   0
      LocalService    =   ""
      Protocol        =   0
      RemotePort      =   0
      RemoteService   =   ""
      ReuseAddress    =   0   'False
      Route           =   -1  'True
      Timeout         =   0
      Type            =   1
      Urgent          =   0   'False
   End
   Begin VB.CheckBox chkDebug 
      Caption         =   "Debug"
      Height          =   250
      Left            =   120
      TabIndex        =   3
      Top             =   3240
      Width           =   1200
   End
   Begin VB.ListBox lstAzione 
      Height          =   2205
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   4995
   End
   Begin VB.Timer tmrHost 
      Interval        =   1000
      Left            =   2760
      Top             =   3120
   End
   Begin VB.Label lblTimeout 
      Caption         =   "IDLE Timeout :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2880
      TabIndex        =   4
      ToolTipText     =   "If a client is idle, the connection will be shutdown after the timeout has expired"
      Top             =   480
      Width           =   2175
   End
   Begin VB.Label lblConnessioni 
      Caption         =   "Client Connected :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   2655
   End
   Begin VB.Label lblStato 
      Caption         =   "State :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2655
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()

    'ascolto messaggio di broadcast
    sckReceiveBroadcast.AddressFamily = AF_INET
    sckReceiveBroadcast.Binary = True
    sckReceiveBroadcast.Blocking = True
    sckReceiveBroadcast.Protocol = IPPROTO_UDP
    sckReceiveBroadcast.SocketType = SOCK_DGRAM

    sckReceiveBroadcast.Blocking = False
    sckReceiveBroadcast.LocalPort = 6999
    sckReceiveBroadcast.Broadcast = False
    'Create Socket to broadcast message
    sckReceiveBroadcast.Open

    sckSend.Protocol = IPPROTO_IP
    sckSend.SocketType = SOCK_DGRAM
    sckSend.Action = SOCKET_OPEN
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    Dim nClientIndex As Integer

    For nClientIndex = 1 To g_nActiveClients
        DropClient nClientIndex, "Drop per chiusura"
    Next nClientIndex

    ExitCrossComm
    End
End Sub

Private Sub Form_Resize()
    Dim width As Integer
    Dim minWidth As Integer
    Dim maxWidth As Integer
    Dim height As Integer
    Dim minHeight As Integer
    Dim maxHeight As Integer
    Dim wState As Integer
    
    If Me.WindowState = 0 Then
        minWidth = 5500
        
        maxWidth = 15000
        
        minHeight = 4100
        maxHeight = 8000
        
        width = Me.width
        If width < minWidth Then
            width = minWidth
        End If
        If width > maxWidth Then
            width = maxWidth
        End If
        
        height = Me.height
        If height < minHeight Then
            height = minHeight
        End If
        If height > maxHeight Then
            height = maxHeight
        End If
        
        Me.width = width
        Me.height = height
        lstAzione.width = width - 400
        lstAzione.height = Me.height - 1600
        chkDebug.Top = Me.height - chkDebug.height - 500
    End If
End Sub

Private Sub sckReceiveBroadcast_Read(DataLength As Integer, IsUrgent As Integer)
    Dim strMess As String
    Dim cbBuffer As Integer
    Dim intError As Integer
    
    On Error GoTo errReadBroadcast
    
    'Read a recieved broadcast and display it
    cbBuffer = sckReceiveBroadcast.Read(strMess, 512)
    
    If (cbBuffer > 0) Then
        If strMess = "WHEREAREYOU?" Then
            
            sckSend.AddressFamily = AF_INET
            sckSend.Binary = True
            sckSend.Blocking = False
            sckSend.LocalPort = IPPORT_ECHO
            sckSend.RemotePort = 7000
            sckSend.HostAddress = sckReceiveBroadcast.PeerAddress
            
            If sSerialNO = "" Then
                ShowVar "$KR_SERIALNO", sSerialNO
                ShowVar "$MODEL_NAME[]", sModelName
            End If
            
            strMess = "KUKA|" & sModelName & "|" & sSerialNO
            intMess = Len(strMess)
            
            intError = sckSend.Write(strMess, intMess)
            
        End If
    End If
    
    On Error GoTo 0
    Exit Sub
    
errReadBroadcast:
    On Error GoTo 0
    
    addMessage "Error in ReadBroadcast"
    
End Sub

Private Sub sockServer_Accept(Index As Integer, SocketId As Integer)
    Dim nClient As Integer

    On Error GoTo errAccept

    For nClient = 1 To g_nLastClient
        If sockServer(nClient).Connected = False Then
            sockServer(nClient).Accept = SocketId
            
            Exit Sub
        End If
    Next

    g_nLastClient = g_nLastClient + 1
    
    Load sockServer(g_nLastClient)
    sockServer(g_nLastClient).AutoResolve = False
    sockServer(g_nLastClient).Blocking = False
    sockServer(g_nLastClient).Accept = SocketId
    
    sockServer(g_nLastClient).KeepAlive = True
    
    On Error GoTo 0
    Exit Sub
    
errAccept:
    On Error GoTo 0
    
    addMessage "Error in Accept"
    
End Sub

Private Sub sockServer_Connect(Index As Integer)

    On Error GoTo errConnect
    
    '
    ' Check the number of active clients against the maximum number
    ' of clients specified by the user
    '
    If g_nActiveClients < g_nMaxClients Then
        g_nActiveClients = g_nActiveClients + 1
        'LogMsg ("Connected on Index " & Index & ", Active Clients: " & g_nActiveClients)
    '        Debug.Print "Connected on Index: " & Index
    '        Debug.Print "Active Clients Currently: " & g_nActiveClients
    '        Debug.Print
    Else
        'LogMsg ("Maximum number of clients exceeded, dropping #" & Index)
        sockServer(Index).Disconnect
        Exit Sub
    End If
    
    'memorizzo il tempo di arrivo del messaggio
    lLastReceiveDate(Index) = Now()
    
    UpdateForm
    On Error GoTo 0
    Exit Sub
    
errConnect:
    On Error GoTo 0
    
    addMessage "Error in Connect"
    
End Sub


Private Sub sockServer_Disconnect(Index As Integer)
    DropClient Index, "Disconnect Event"
    UpdateForm
End Sub

Private Sub sockServer_LastError(Index As Integer, ErrorCode As Integer, ErrorString As String, Response As Integer)
    
    If ErrorCode = 24048 Then
        MsgBox "PORTA GIA' IN USO", vbOKOnly, "KUKAVARPROXY"
        End
    End If
    
    DropClient Index, "LastError Event"
End Sub

Private Sub sockServer_Read(Index As Integer, DataLength As Integer, IsUrgent As Integer)
    Dim strBuffer As String
    Dim cchBuffer As Integer
    
    Static bReadPending As Boolean
    
    On Error GoTo errRead
    
    cchBuffer = sockServer(Index).Read(strBuffer, 2048)
   
    If (cchBuffer < 0) Then
        
        'If a would block error occurs exit the sub and let
        'another read event read when the client is ready
        If (sockServer(Index).LastError = WSAEWOULDBLOCK) Then
            Exit Sub
        End If
        
        'If an in progress error occurs exit the sub and let
        'another read event read when the client is ready
        If (sockServer(Index).LastError = WSAEINPROGRESS) Then
            Exit Sub
        End If

        'There has been an error in the read
        'print the error and quit
        DropClient Index, "Error occurred durring the read event: Error #", sockServer(Index).LastError
        UpdateForm
        Exit Sub
    End If
        
    Dim sValueToWrite As String
    Dim lLunghezzaBlocco As Long
    Dim lLunghezzaBloccoToWrite As Long
    Dim sAzione As String
    Dim lMsgID As Long
    Dim sMsgID As String
    Dim nFunction As Integer
        
    'i primi due byte indicano l'ID del messaggio ricevuto
    
    lMsgID = CLng(Asc(Mid(strBuffer, 1, 1))) * &H100 + Asc(Mid(strBuffer, 2, 1))
    sMsgID = Mid(strBuffer, 1, 2)
    
    lLunghezzaBlocco = Asc(Mid(strBuffer, 3, 1)) * &H100 + Asc(Mid(strBuffer, 4, 1))
    nMsg = 0
    
    While Len(strBuffer) >= lLunghezzaBlocco
        'DoEvents
        nMsg = nMsg + 1
        strMsg = Mid(strBuffer, 6, lLunghezzaBlocco)
        nFunction = CInt(Asc(Mid(strBuffer, 5, 1)))
        'Debug.Print "Msg separato numero " & nMsg & ": " & strMsg
        
        If Not bReadPending Then
            bReadPending = True
            If readMsg(nFunction, strMsg, sValueToWrite, sAzione) Then
            
                lLunghezzaBloccoToWrite = Len(sValueToWrite)
                sValueToWrite = sMsgID & longToWord(lLunghezzaBloccoToWrite) & sValueToWrite
    '                For a = 1 To Len(sValueToWrite)
    '                    Debug.Print "Carattere (" & a - 1 & ")= " & Asc(Mid(sValueToWrite, a, 1))
    '                Next a
                
                If frmMain.sockServer(Index).Write(sValueToWrite, Len(sValueToWrite)) < 0 Then
                    DropClient Index, "Error occurred durring the read event: Error #", sockServer(Index).LastError
                End If
                
                If frmMain.chkDebug.Value = 1 Then
                    addMessage "ID=" & Format(lMsgID, "00000") & " " & sAzione
                End If
            
            End If
            
            bReadPending = False
            
        Else
        
            addMessage "!!Reading pending!!"
        
        End If
        
        strBuffer = Right(strBuffer, Len(strBuffer) - (lLunghezzaBlocco + 4))
        If Len(strBuffer) > 3 Then
            lMsgID = CLng(Asc(Mid(strBuffer, 1, 1))) * &H100 + Asc(Mid(strBuffer, 2, 1))
            sMsgID = Mid(strBuffer, 1, 2)
            lLunghezzaBlocco = Asc(Mid(strBuffer, 3, 1)) * &H100 + Asc(Mid(strBuffer, 4, 1))
            
            'If frmMain.chkDebug.Value = 1 Then
            '    addMessage "Split msg: " & nMsg & " ID=" & lMsgID & " " & sAzione
            'End If
            
            Debug.Print "Split msg number " & nMsg & ": " & strMsg
        End If
        
    Wend
    
    'Debug.Print "Lunghezza attesa: " & lLunghezzaBlocco & " lunghezza ricevuta: " & Len(strBuffer)
        
    'memorizzo il tempo di arrivo del messaggio
    lLastReceiveDate(Index) = Now()
    
    On Error GoTo 0
    Exit Sub
    
errRead:
    On Error GoTo 0
    
    addMessage "Error in Read"
    addMessage "Msg ID " & lMsgID
    addMessage "Msg Len " & lLunghezzaBlocco
    addMessage "Buffer Len " & Len(strBuffer)
    addMessage strMsg
    
    bReadPending = False
    
End Sub

Public Sub DropClient(Index As Integer, Prompt As String, Optional nError As Integer)
    Dim strDisconnect As String
    
    On Error GoTo errDropClient
    
    If sockServer(Index).Connected Then
        strDisconnect = "Socket " & Index & " (handle = " & sockServer(Index).Handle & ") disconnected"
        sockServer(Index).Disconnect
        g_nActiveClients = g_nActiveClients - 1
    Else
        strDisconnect = "******************* Socket " & Index & " (handle = " & sockServer(Index).Handle & ") previously disconnected"
    End If
    
    If Prompt <> "" Then
        strDisconnect = strDisconnect & ": " & Prompt
    End If
    
    If Not IsMissing(nError) And nError <> 0 Then
        'LogError strDisconnect, nError, sockServer(Index).PeerAddress, PageStats(Index).lByte
    Else
        'LogMsg (strDisconnect)
    End If
    
    '    Debug.Print "Disconected on Index: " & Index
    '    Debug.Print "Active Clients Currently: " & g_nActiveClients
    '    Debug.Print "Location: " & Prompt
    '    Debug.Print
    
    If (g_nActiveClients < 0) Then
        'LogMsg ("***************** Active Clients < 0 **************")
    End If
    
    'Metto a zero la variabile del tempo di ricezione
    'If g_nActiveClients > 0 Then
        lLastReceiveDate(Index) = Now()
    'End If
    
    UpdateForm
    On Error GoTo 0
    Exit Sub
    
errDropClient:
    On Error GoTo 0
    
    addMessage "Error in ropClient"
    
End Sub

Private Sub tmrHost_Timer()

    Dim nClientIndex As Integer
    Dim lSeconds As Long

    On Error GoTo errTmrHost

    For nClientIndex = 1 To g_nActiveClients
        'when the remote host dies due to a hardware software problem
        'without closing the connection, the service remains in the state of
        'connected, the only way to restore the connection is
        'terminate and restart the program.
        'This condition is normal in the type of TCP communication. Usually
        'the connections are opened and closed after exchanging the data.
        'To prevent certain connections from hanging, I enter a check
        'on receiving requests from clients.
        'If the last data request is older than 60 seconds (60000 milliseconds)
        'I forcibly close the connection. The client will have to reopen
        'the same to request the data again.
        
        lSeconds = DateDiff("s", lLastReceiveDate(nClientIndex), Now())
        If lSeconds * 1000 > lTimeOutRequest Then
            'The reception timeout has expired. Disconnect the client.
            addMessage "Closing connection with client " & Format(nClientIndex, 0) & " (idle timeout)"
            DropClient nClientIndex, "Error occurred for timeout #", 0
        End If

    Next nClientIndex

    On Error GoTo 0
    Exit Sub

errTmrHost:
        On Error GoTo 0
        
        addMessage "Error in tmrHost"

End Sub


