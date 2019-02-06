Attribute VB_Name = "basMain"
'-----------------------------------------------------------------
'IMTS s.r.l.
'Via del Tratturello Tarantino, 6
'Tel 099/4725996 Fax 099/4729035
'email service@imts.eu
'web www.imts.eu
'-----------------------------------------------------------------
' Modified by : Lionel du Peloux (feb. 2019)
'-----------------------------------------------------------------

'dichiarazione oggetto crosscomm
Public CrossCommands As New cCrossComm

'Separatore di valori per comunicazione TCP/IP
Public Const sSeparatore = "|"

Public Const lTimeOutRequest = 30000

Public g_nMaxClients As Integer
Public g_nLastClient As Integer
Public g_nActiveClients As Integer

'Array che contiene il tempo di ricezione del messaggio
Public lLastReceiveDate() As Date

Public sSerialNO As String
Public sModelName As String

Public CrossIsConnected As Boolean
Public Sub ExitCrossComm()
    'Unload Cross-Communication
    If Not (CrossCommands Is Nothing) Then
        On Error Resume Next
        CrossCommands.ServerOff
        Set CrossCommands = Nothing
    End If
End Sub

Public Function Connect(ByVal nMode As Integer) As Boolean
'effettuo la connessione con il CrossComm del robot KUKA
'La connessione viene sempre effettuata in modalite sincrona
'Nel corso dell'esecuzione del programma, la connessione deve essere
'verificata ad intervalli di tempo.

Dim sValue As String

On Error GoTo Fehler
    'Connect to KRC
    
    Dim bRetVal As Boolean
    bRetVal = False
    
    If Not CrossCommands.CrossIsConnected Then
        If CrossCommands.Init(frmMain) Then
            Connect = CrossCommands.ConnectToCross("KUKAVARPROXY", nMode)
            ShowVar "$KR_SERIALNO", sSerialNO
            ShowVar "$MODEL_NAME[]", sModelName
            
        Else
            Connect = False
        End If
    Else
        Connect = True
    End If
    
On Error GoTo 0

Ende:
    InitCrossComm = bRetVal
    Exit Function

Fehler:
    '**VERIFICARE
    MsgBox Err.Description, vbCritical, Err.Number
    Resume Ende

End Function


Public Function longToWord(ByVal lValue As Long) As String
    varl = lValue And &HFF
    varh = (lValue And &HFF00) / (&H100)
    varh = varh And &HFF
    longToWord = Chr(varh) & Chr(varl)
End Function

Public Function readMsg(ByVal nFunction As Integer, ByVal strBuffer As String, ByRef sValueToWrite As String, ByRef sAzione As String) As Boolean

    Dim sVariableName As String
    Dim sVariableValue As String
    Dim bWriteAllOk As Boolean
    Dim sString(9) As String
    Dim sArrayValue As String
    Dim sArrayValueMessage As String
    
    If splitMsg(strBuffer, sString) Then
    
        sVariableName = sString(0)
        
    '    addMessage "Buffer " & strBuffer
    '    addMessage "Funzione " & nFunction
    '    addMessage "Nome variabile " & sVariableName
        
        Select Case nFunction
        Case 0
            'lettura variabile
            If sVariableName = "PING" Then
                ' PING-PONG keepalive to see if host is down
                ' if the proxy receives a request to read "PING" it just sends back a "PONG" answer (no crosscomm call)
                sVariableValue = "PONG"
                sValueToWrite = Chr(nFunction) & longToWord(Len(sVariableValue)) & sVariableValue & longToWord(Len(Chr(1))) & Chr(1)
            Else
                If ShowVar(sVariableName, sVariableValue) Then
                    sValueToWrite = Chr(nFunction) & longToWord(Len(sVariableValue)) & sVariableValue & longToWord(Len(Chr(1))) & Chr(1)
                Else
                    sValueToWrite = Chr(nFunction) & longToWord(Len(sVariableValue)) & longToWord(Len(Chr(0))) & Chr(0)
                End If
            End If

            sAzione = "Read: " & sVariableName & "=" & sVariableValue
            
        Case 1
            'scrittura variabile
            
            sVariableValue = sString(1)
            
            If SetVar(sVariableName, sVariableValue) Then
                sValueToWrite = Chr(nFunction) & longToWord(Len(sVariableValue)) & sVariableValue & longToWord(Len(Chr(1))) & Chr(1)
            Else
                sValueToWrite = Chr(nFunction) & longToWord(Len(sVariableValue)) & sVariableValue & longToWord(Len(Chr(0))) & Chr(0)
            End If
            
            sAzione = "Write: " & sVariableName & "=" & sVariableValue
            
        Case 2
            'lettura e formattazione di una variabile array destinata al PLC
            
            If ShowVar(sVariableName, sVariableValue) Then
            
                Dim vVettore As Variant
                Dim nDimVettore As Integer
                Dim sMsg As String
                
                vVettore = Split(sVariableValue, " ")
                nDimVettore = UBound(vVettore)
    
                If nDimVettore > 0 Then
                    For a = 0 To nDimVettore
                        sMsg = sMsg & longToWord(CLng(vVettore(a)))
                    Next a
                End If
                sValueToWrite = Chr(nFunction) & longToWord(Len(sMsg)) & sMsg & longToWord(Len(Chr(1))) & Chr(1)
            Else
                sValueToWrite = Chr(nFunction) & longToWord(Len(sMsg)) & sMsg & longToWord(Len(Chr(0))) & Chr(0)
            End If
            
            sAzione = "Read array: " & sVariableName & "=" & sVariableValue
            
        Case 3
            'scrittura di una variabile array destinata al PLC
            'il nome variabile deve essere inviato senza le parentesi quadre

            sVariableValue = sString(1)
            nDimVettore = Len(sVariableValue) / 2

            For a = 0 To nDimVettore - 1
                sArrayValue = CStr(wordToLong(Mid(sVariableValue, (2 * a) + 1, 2)))
                If Right(sVariableName, 2) = "[]" Then
                    If SetVar(Left(sVariableName, Len(sVariableName) - 2) & "[" & a + 1 & "]", sArrayValue) Then
                        bWriteAllOk = True
                    Else
                        bWriteAllOk = False
                    End If
                    'addMessage Left(sVariableName, Len(sVariableName) - 2) & "[" & a + 1 & "] = " & CStr(wordToLong(Mid(sVariableValue, (2 * a) + 1, 2)))
                Else
                    If SetVar(sVariableName & "[" & a + 1 & "]", sArrayValue) Then
                        bWriteAllOk = True
                    Else
                        bWriteAllOk = False
                    End If
                End If
                sArrayValueMessage = sArrayValueMessage & Asc(sArrayValue) & " "
            Next a

            If bWriteAllOk Then
                sValueToWrite = Chr(nFunction) & longToWord(Len(sMsg)) & sMsg & longToWord(Len(Chr(1))) & Chr(1)
            Else
                sValueToWrite = Chr(nFunction) & longToWord(Len(sMsg)) & sMsg & longToWord(Len(Chr(0))) & Chr(0)
            End If

            sAzione = "Write array: " & sVariableName & "=" & sVariableValue
            
        End Select
        
        readMsg = True
    
    End If
    
End Function

Public Function splitMsg(ByVal sMsg As String, ByRef sArrayString() As String) As Boolean

Dim nStartPos As Integer
Dim nLunghezzaMsg As Integer
Dim nIndex As String
Const nLenMsg = 2

On Error GoTo errsplitMsg

'sMsg = longToWord(4) & "CIAO" & longToWord(8) & "CIAOciao"

nStartPos = 1
nIndex = 0
nLunghezzaMsg = Len(sMsg)

Do While nStartPos < nLunghezzaMsg
    lLunghezzaBlocco = Asc(Mid(sMsg, nStartPos, 1)) * &H100 + Asc(Mid(sMsg, nStartPos + 1, 1))
    
    sArrayString(nIndex) = Mid(sMsg, nStartPos + nLenMsg, lLunghezzaBlocco)
    
    nStartPos = nStartPos + nLenMsg + lLunghezzaBlocco
    nIndex = nIndex + 1

Loop

splitMsg = True
On Error GoTo 0

Exit Function

errsplitMsg:

splitMsg = False
On Error GoTo 0

End Function


Public Function wordToLong(ByVal sValue As String) As Long
    Dim varl As Long
    Dim varh As Long
    
    If Len(sValue) = 2 Then
        varl = Asc(Mid(sValue, 2, 1))
        varh = Asc(Mid(sValue, 1, 1))
    
        varh = (varh) * (&H100)
        wordToLong = varh + varl
    Else
        wordToLong = 0
    End If
End Function
Public Sub Main()

Dim OSType As WindowsVersion

'Prepare our listening socket
frmMain.sockServer(0).AddressFamily = AF_INET
frmMain.sockServer(0).Protocol = IPPROTO_IP
frmMain.sockServer(0).SocketType = SOCK_STREAM
frmMain.sockServer(0).Blocking = False
frmMain.sockServer(0).AutoResolve = False
frmMain.sockServer(0).LocalPort = 7000

frmMain.sockServer(0).Listen

g_nMaxClients = 10
g_nLastClient = 0
g_nActiveClients = 0

ReDim lLastReceiveDate(g_nMaxClients)

'Visualizzazione finestra IMSCROSSCOMM

frmMain.Caption = "KukavarProxy " & App.Major & "." & App.Minor & "." & App.Revision & " | " & GetOSVersion(OSType)

Load frmMain
frmMain.Show

UpdateForm

End Sub


Public Sub UpdateForm()

    Dim sStato As String
    
    On Error GoTo errUpdateForm

    If frmMain.sockServer(0).Listening Then
        sStato = "Listening..."
    Else
        sStato = "Disconnected"
    End If
    
    frmMain.lblStato.Caption = "State: " & sStato
    frmMain.lblConnessioni.Caption = "Clients: " & g_nActiveClients & "/" & g_nMaxClients
    frmMain.lblTimeout.Caption = "IDLE Timeout : " & lTimeOutRequest / 1000 & " s"
    
    On Error GoTo 0
    Exit Sub
    
errUpdateForm:
    On Error GoTo 0
    
    addMessage "Error in UpdateForm"

End Sub
Public Function SetVar(ByVal sVariableName As String, ByVal sVariableValue As String) As Boolean
    
    On Error GoTo errSetVar
    
    If Connect(0) Then
        SetVar = CrossCommands.SetVar(sVariableName, sVariableValue)
    End If
    
    On Error GoTo 0
    Exit Function
    
errSetVar:
    On Error GoTo 0
    
    addMessage "Error in SetVar"
    
End Function

Public Function ShowVar(ByVal sVariableName As String, ByRef sVariableValue As String) As Boolean
'Effettuo la lettura di una variabile

On Error GoTo errShowVar

Dim stParam As String

If Connect(0) Then
    If CrossCommands.ShowVar(sVariableName, stParam) Then
        sVariableValue = ExtractVariableValue(stParam)
    Else
        sVariableValue = ""
    End If
    ShowVar = True
Else
    sVariableValue = ""
    ShowVar = False
End If

On Error GoTo 0
Exit Function

errShowVar:
    On Error GoTo 0
    
    addMessage "Error in ShowVar"
    
End Function


Public Sub Wait(msWait As Long)
    Dim msEnd As Long
    msEnd = GetTickCount() + msWait
    Do
        DoEvents
    Loop While GetTickCount() < msEnd
End Sub
Public Function ExtractVariableValue(ByVal VarString As String)
    
    On Error GoTo errExtractVariable
    
    Dim Pos
    If Len(VarString) > 0 Then
        Pos = InStr(VarString, "=")
        If Pos > 0 Then
            ExtractVariableValue = Trim(Mid$(VarString, Pos + 1))
        Else
            ExtractVariableValue = ""
        End If
    Else
        ExtractVariableValue = ""
    End If
    
    On Error GoTo 0
    Exit Function
    
errExtractVariable:
    On Error GoTo 0
    
    addMessage "Error in ExtractVariableValue"
    
End Function

Public Sub addMessage(ByVal sMsg As String)

    On Error Resume Next

    frmMain.lstAzione.AddItem Format(Now, "hh:mm:ss") & " " & sMsg
    frmMain.lstAzione.Selected(frmMain.lstAzione.ListCount - 1) = True
        
    If frmMain.lstAzione.ListCount > 100 Then
        frmMain.lstAzione.RemoveItem 0
    End If
    
    On Error GoTo 0
    
End Sub



