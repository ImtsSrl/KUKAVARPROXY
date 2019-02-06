VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cCrossComm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'=================================================================
' Class CrossCommand
'-----------------------------------------------------------------
' Guengoer, 05.11.2001
'=================================================================
' Implements KUKA Cross KRC Library (WBC_KrcLib)
' The class encapsulates the Cross KRC Library and offers
' functionality compatible to CrossCommEXE
'-----------------------------------------------------------------
' V1.0.8 2005-04-11
'     HandleLastError: Fix for interface IKAsyncError
' V1.0.7 2005-01-27
'     SetInfoOn: 1st SetInfoOn didn't provide sResult --> Fixed
' V1.0.5  DownLoadMemToRobot: Error with lowercase filename fixed
'=================================================================

Option Explicit

' Constants
Public Enum eOpenMode
    eOpenModeRead = &H1
    eOpenModeWrite = &H2
    eOpenModeReadWrite = &H3
    eOpenModeExclusive = &H4
    eOpenModeData = &H8
    eOpenModeSource = &H10
    eOpenModeRaw = &H20
    eOpenModeCancel0 = &H40
    eOpenModeCancel1 = &H80
    eOpenModeReSelect0 = &H100
    eOpenModeReSelect1 = &H200
End Enum

Public Enum eEditMode
    eEditModeFullEdit = &H0
    eEditModeDatKor = &H1
    eEditModeProKor = &H2
    eEditModeReadOnly = &H3
End Enum

Public Enum eFileAttrib
    eItemAttribReadOnly = vbReadOnly
    eItemAttribSystem = vbSystem
    eItemAttribHidden = vbHidden
    eItemAttribArchiv = vbArchive
    eItemAttribDirectory = vbDirectory
End Enum

Public Enum eItemType
    eItemTypeUnknown = &H0
    eItemTypeDir = &H1
    eItemTypeVirtualDir = &H2
    eItemTypeArchiv = &H4
    eItemTypeBinFile = &H8
    eItemTypeTextFile = &H10
    eItemTypeModule = &H20
    eItemTypeRaw = &H40
End Enum

Public Enum IO_TYPE
    IoRestart = -1
    IBusReset = 1
    CanReset = 2
    BoschReset = 3
    PerceptronReset = 4
End Enum

Public Enum Msg_Type
    MELD_ADE = -1
    MELD_H = 1
    MELD_Q = 2
    MELD_Z = 3
End Enum

Private Const nTimeOut_Standard = 3000
Private Const nTimeOut_Dir = 10000
Private Const nTimeOut_UpDownload = 70000

Private objServiceFactory As IKServiceFactory
Private itfSyncvar As ICKSyncVar
Private itfSyncFile As ICKSyncFile
Private itfSyncEdit As ICKSyncEdit
Private itfSyncSelect As ICKSyncSelect
Private itfSyncIo As ICKSyncIo
Private itfAsyncVar As ICKAsyncVar
Private itfAdviseMessage As ICKAdviseMessage
    
Dim m_Parent As Object
Dim sLocalConnectName As String 'Client-ID
Dim sLastError As String         'last cross error
Dim m_CrossIsConnected As Boolean
Dim m_KrcOk As Boolean
Dim CSetInfo As New Collection
Dim CSetInfoItems() As New Collection
Dim m_nID As Long
Dim nC_LocalMode As Integer

'for callback-functions SetInfoOn, SetInfoOff, RobMessage
Implements ICKCallbackVar
Implements IKCallbackError
Implements ICKConsumeMessage

Public Function Init(ByRef oParent As Object) As Boolean
On Error GoTo Err_Init
    
    Debug.Print "--- Init ---"
    If Not (oParent Is Nothing) Then
        Set m_Parent = oParent
        Set objServiceFactory = New KrcServiceFactory
        'CALLBACK
        ReDim CSetInfoItems(1)
    Else
        'object is missing
        Init = False
        Exit Function
    End If
    
Exit_Init:
    Init = True
    m_CrossIsConnected = True
    Exit Function
Err_Init:
    If Err.Number <> 0 Then Init = False
    Err.Clear
    m_CrossIsConnected = False
    Exit Function
End Function

Public Function ConnectToCross(ByVal sConnectName As String, Optional nC_Mode As Integer = 0) As Boolean
On Error GoTo Err_ConnectToCross
    Dim itfConsumeMessage As cCrossComm
    Dim nMessage As Long
    
    Debug.Print "--- ConnectToCross ---"
    sLocalConnectName = sConnectName
    nC_LocalMode = nC_Mode

    Set itfSyncvar = Nothing
    Set itfSyncvar = objServiceFactory.GetService("WBC_KrcLib.SyncVar", sLocalConnectName)
    
    Set itfSyncFile = Nothing
    Set itfSyncFile = objServiceFactory.GetService("WBC_KrcLib.SyncFile", sLocalConnectName)
    
    Set itfSyncEdit = Nothing
    Set itfSyncEdit = objServiceFactory.GetService("WBC_KrcLib.SyncEdit", sLocalConnectName)
    
    Set itfSyncSelect = Nothing
    Set itfSyncSelect = objServiceFactory.GetService("WBC_KrcLib.SyncSelect", sLocalConnectName)
    
    Set itfSyncIo = Nothing
    Set itfSyncIo = objServiceFactory.GetService("WBC_KrcLib.SyncIo", sLocalConnectName)
    
    Set itfAsyncVar = Nothing
    Set itfAsyncVar = objServiceFactory.GetService("WBC_KrcLib.AsyncVar", sLocalConnectName)
    
    'advise for all message types
    If nC_LocalMode = -1 Then
        Set itfAdviseMessage = Nothing
        Set itfAdviseMessage = objServiceFactory.GetService("WBC_KrcLib.AdviseMessage", sLocalConnectName)
        
        Set itfConsumeMessage = New cCrossComm
        itfConsumeMessage.InitConsumeMessage m_Parent 'set parent object for callback class
        
        nMessage = CLng(eMessageTypeInfo) Or CLng(eMessageTypeState) Or CLng(eMessageTypeEvent) Or _
                  CLng(eMessageTypeQuitt) Or CLng(eMessageTypeWait) Or CLng(eMessageTypeDialog)
        itfAdviseMessage.Advise itfConsumeMessage, nMessage
    End If
    
Exit_ConnectToCross:
    ConnectToCross = True
    m_CrossIsConnected = True
    m_KrcOk = True
    Exit Function
Err_ConnectToCross:
    If Err.Number <> 0 Then ConnectToCross = False
    Err.Clear
    m_CrossIsConnected = False
    m_KrcOk = False
    Exit Function
End Function

Function InitConsumeMessage(ByRef oParent As Object) As Boolean
'set parent object for ICKConsumeMessage callback class
On Error GoTo Err_InitConsumeMessage
    
    Debug.Print "--- Init Consume Message ---"
    If Not (oParent Is Nothing) Then
        Set m_Parent = oParent
    Else
        'object is missing
        InitConsumeMessage = False
        Exit Function
    End If
    
Exit_InitConsumeMessage:
    InitConsumeMessage = True
    Exit Function
Err_InitConsumeMessage:
    If Err.Number <> 0 Then InitConsumeMessage = False
    Err.Clear
    Exit Function
End Function

Public Function ServerOff() As Boolean
    Debug.Print "--- ServerOff ---"
    Set itfSyncvar = Nothing
    Set itfSyncFile = Nothing
    Set itfSyncEdit = Nothing
    Set itfSyncSelect = Nothing
    Set itfSyncIo = Nothing
    Set itfAsyncVar = Nothing
    Set itfAdviseMessage = Nothing
    Set objServiceFactory = Nothing
End Function

'--------------------------------------------------------------------------------
'Variable Handling
Public Function SetVar(ByVal strVarName, strNewValue As String, Optional Timeout As Long) As Boolean
On Error GoTo Err_SetVar

    Debug.Print "--- SetVar ---"
    itfSyncvar.SetVar strVarName, strNewValue

Exit_SetVar:
    SetVar = True
    Exit Function
Err_SetVar:
    HandleLastError itfSyncvar
    Err.Clear
    SetVar = False
    Exit Function
End Function

Public Function ShowVar(ByVal strVarName As String, ByRef strVarValue As String, Optional Timeout As Long) As Boolean
On Error GoTo Err_ShowVar

    Debug.Print "--- ShowVar ---"
    strVarValue = strVarName + "=" + itfSyncvar.ShowVar(strVarName)

Exit_ShowVar:
    ShowVar = True
    Exit Function
Err_ShowVar:
    HandleLastError itfSyncvar
    Err.Clear
    ShowVar = False
    Exit Function
End Function

Public Function SetInfoOn(sVariableName As String, sResult As String, Optional vTimeOut As Long = nTimeOut_Standard) As Boolean
'SetInfo count max. 75
On Error GoTo Err_SetInfoOnAsync
    Dim objCallbackVar As cCrossComm
    Dim bVarExists As Boolean
    Dim i As Integer
    
    Debug.Print "--- SetInfoOn ---"
    Set objCallbackVar = Me
    bVarExists = False
    
    'check if SetInfo for sVariableName is already active
    If CSetInfo.Count = 0 Then
        m_nID = m_nID + 1 'unique nID for SetInfo
        ReDim Preserve CSetInfoItems(m_nID + 1)
        
        'add variable to collection
        CSetInfo.Add CSetInfoItems(m_nID), sVariableName
        CSetInfo.Item(sVariableName).Add sVariableName, "sVariableName"
        CSetInfo.Item(sVariableName).Add CStr(m_nID), "m_nID"
        
        itfAsyncVar.SetInfo objCallbackVar, m_nID, sVariableName, 200

        On Error GoTo Err_SetInfoOnSync
        sResult = itfSyncvar.ShowVar(sVariableName)
        On Error GoTo Err_SetInfoOnAsync
    Else
        'check if sVariableName exists
        For i = 1 To CSetInfo.Count
            If sVariableName = CSetInfo.Item(i).Item("sVariableName") Then
                bVarExists = True
                Exit For
            End If
        Next
        If Not bVarExists Then
            m_nID = m_nID + 1 'unique nID for SetInfo
            ReDim Preserve CSetInfoItems(m_nID + 1)
            
            'add variable to collection
            CSetInfo.Add CSetInfoItems(m_nID), sVariableName
            CSetInfo.Item(sVariableName).Add sVariableName, "sVariableName"
            CSetInfo.Item(sVariableName).Add CStr(m_nID), "m_nID"
            
            itfAsyncVar.SetInfo objCallbackVar, m_nID, sVariableName, 200
            
            On Error GoTo Err_SetInfoOnSync
            sResult = itfSyncvar.ShowVar(sVariableName)
            On Error GoTo Err_SetInfoOnAsync
        End If
    End If

Exit_SetInfoOn:
    SetInfoOn = True
    Exit Function
Err_SetInfoOnAsync:
    HandleLastError itfAsyncVar
    Err.Clear
    SetInfoOn = False
    Exit Function
Err_SetInfoOnSync:
    HandleLastError itfSyncvar
    Err.Clear
    SetInfoOn = False
    Exit Function
End Function

Public Function SetInfoOff(sVariableName As String, Optional vTimeOut As Long = nTimeOut_Standard) As Boolean
On Error GoTo Err_SetInfoOff
    Dim nID As Long
    
    Debug.Print "--- SetInfoOff ---"
    'determine of nID from index of the item
    nID = CLng(CSetInfo.Item(sVariableName).Item("m_nID"))
    'remove sVariableName from collection
    CSetInfo.Remove sVariableName
    
    itfAsyncVar.Cancel (nID)

Exit_SetInfoOff:
    SetInfoOff = True
    Exit Function
Err_SetInfoOff:
    HandleLastError itfAsyncVar
    Err.Clear
    SetInfoOff = False
    Exit Function
End Function
'Variable Handling end
'--------------------------------------------------------------------------------


'--------------------------------------------------------------------------------
'File-Handling
Public Function GetRobotDirectory(ByRef sDir As String, Optional vMask As String, Optional vTimeOut = nTimeOut_Dir) As Boolean
On Error GoTo Err_GetRobotDirectory
'Read Directory from Robot
    Dim sLocalMask As String
    Dim nType As Long
    Dim nListFlag As Long
    Dim varInfo As Variant
    Dim varNames As Variant
    Dim i As Integer
    Dim strMsg As String
                
    Debug.Print "--- GetRobotDirectory ---"
    ' get robot mask
    If IsMissing(vMask) Then
        sLocalMask = "KRC:\"
    Else
        sLocalMask = "KRC:\" & vMask
    End If
    
    nType = CLng(eItemTypeUnknown)  'All directories and files
    
    nListFlag = CLng(eListFlagLong) Or CLng(eListFlagRecursive) 'Detailed and Recursive List
    
    itfSyncFile.NameList sLocalMask, nType, nListFlag, varNames, varInfo
    
    For i = LBound(varNames) To UBound(varNames)
        strMsg = strMsg + CStr(varNames(i)) + Chr(10)
    Next i
    
    sDir = strMsg   'return sDir
    
Exit_GetRobotDirectory:
    GetRobotDirectory = True
    Exit Function
Err_GetRobotDirectory:
    HandleLastError itfSyncFile
    Err.Clear
    GetRobotDirectory = False
    Exit Function
End Function

Public Function DownLoadDiskToRobot(sFileName As String, Optional vTimeOut As Long = nTimeOut_UpDownload) As Boolean
'gueltige werte fuer sFileName:Quellstring | Zielstring; Ohne Wildcards
'                                                 wird ein Zielstring angegebe muss dies ein gueltiger  KRC-Pfad sein:
'                                                     - KRC:\R1\[FOLGEN;SYSTEM;UPs;...\]DATEINAME.EXT
'                                                 Quellstring; ohne bzw.  mit  Wildcards (*,?)
'                                                 wird nur ein Quelle angegen, wird wenn vorhanden die entsprechende Datei ueberschrieben
'                                                 enthaelt die Quelle einen Namen  der vom Filemanager in ein entsprechendes Verzeichnis eingeordnet werden kann
'                                                 wird demensprechend die Datei kopiert andernfalls wird versucht die Datei nach KRC:\R1 zu kopieren
'Download from a File from HD to Robot
On Error GoTo Err_DownLoadDiskToRobot

'    Dim srcName As String
'    Dim destName As String
    Dim onlyFileName As String
    Dim tmpstr As String
    Dim IsWildCard As Boolean
    Dim IsDestPath As Boolean
    Dim i, c, c1 As Integer
    
    Dim m_sFileName As String
    Dim sSrcName As String
    Dim sDestPath As String
    Dim sDestName As String
    Dim sKrcName As String
    Dim sFullname As String
    Dim nFlags As Long
    Dim bDummy As Boolean
    
    Debug.Print "--- DownLoadDiskToRobot ---"
    IsDestPath = False
    IsWildCard = False
    DownLoadDiskToRobot = False
    If Not m_CrossIsConnected Then DownLoadDiskToRobot = False: Exit Function
    If (InStr(1, sFileName, "*") <> 0) Or (InStr(1, sFileName, "?") <> 0) Then IsWildCard = True
    i = InStr(1, sFileName, "|")
    If i > 0 Then IsDestPath = True
      If IsWildCard Then
        'Pfadangaben mit Wildcards
        If IsDestPath Then
            'es wird nur in Standardverzeichnisse geschrieben
            DownLoadDiskToRobot = False
            Exit Function
        Else
            On Error Resume Next
            tmpstr = Dir(sFileName, vbNormal)
            Do While tmpstr <> ""
                If (tmpstr <> ".") And (tmpstr <> "..") Then
                    If Err.Number <> 0 Then
                        DownLoadDiskToRobot = False
                        On Error GoTo 0
                        Exit Function
                    End If
                    On Error GoTo 0
                    onlyFileName = tmpstr
                    If Not GetFileName(onlyFileName) Then
                        DownLoadDiskToRobot = False
                        Exit Function
                    Else
                        sSrcName = tmpstr
                        
                        sFullname = itfSyncFile.GetFullName("/R1/" & onlyFileName)
                        
                        If Not sFullname = "" Then
                            tmpstr = onlyFileName
                            If onlyFileName <> "" Then
                                sDestName = sFullname 'onlyFileName
                             Else
                                sDestName = "KRC:\R1\" & tmpstr
                            End If
                        Else
                            sDestName = "KRC:\R1\" & tmpstr
                        End If
                    End If
                    c = InStr(1, sFileName, "\")
                    Do While c <> 0
                        c1 = c
                        c = InStr(c + 1, sFileName, "\")
                    Loop
                    sSrcName = Left(sFileName, c1) & tmpstr
                    'existing file will be overwritten
                    nFlags = CLng(eCopyFlagOverwriteExist)
                    itfSyncFile.Copy sSrcName, sDestName, nFlags
                End If
                tmpstr = Dir
            Loop
            DownLoadDiskToRobot = True
        End If  'IsDestPath
      Else
        'Pfadangaben ohne Wildcards
        If IsDestPath Then
          sSrcName = Trim$(Left$(sFileName, i - 1))
          sDestName = Trim$(Right$(sFileName, Len(sFileName) - i))
        Else
          onlyFileName = sFileName
          If Not GetFileName(onlyFileName) Then
              DownLoadDiskToRobot = False
              Exit Function
          Else
              sSrcName = sFileName
              sFullname = itfSyncFile.GetFullName("/R1/" & onlyFileName)
              tmpstr = onlyFileName
              If Not sFullname = "" Then
                  If onlyFileName <> "" Then
                    sDestName = sFullname 'onlyFileName
                  Else
                    sDestName = "KRC:\R1\" & tmpstr
                  End If
              Else
                  sDestName = "KRC:\R1\" & tmpstr
              End If
        End If
      End If  'IsWildCard
      
      'existing file will be overwritten
      nFlags = CLng(eCopyFlagOverwriteExist)
      itfSyncFile.Copy sSrcName, sDestName, nFlags
  End If
    
Exit_DownLoadDiskToRobot:
    DownLoadDiskToRobot = True
    Exit Function
Err_DownLoadDiskToRobot:
    HandleLastError itfSyncFile
    Err.Clear
    DownLoadDiskToRobot = False
    Exit Function
End Function

Public Function DownLoadMemToRobot(sFileName As String, sFile As String, Optional vTimeOut As Long = nTimeOut_UpDownload, Optional bKonvert As Boolean = False) As Boolean
'Download from a File from Memory to Robot
  On Error GoTo Err_DownLoadMemToRobot

    Dim m_sFileName As String
    Dim m_sFile As String
    Dim sSrcName As String
    Dim nFlags As Long
    Dim sDir As String
    Dim arrData() As Byte
    
    Debug.Print "--- DownLoadMemToRobot  ---"
    DownLoadMemToRobot = False
    If Not m_CrossIsConnected Then DownLoadMemToRobot = False: Exit Function
    m_sFileName = sFileName
    If Not GetFileName(m_sFileName) Then Exit Function
    
    'check if file exists
    GetRobotDirectory sDir
    If InStr(1, UCase(sDir), UCase(m_sFileName)) > 0 Then
        sSrcName = itfSyncFile.GetFullName("/R1/" + m_sFileName)
    Else
        sSrcName = "KRC:\R1\" + UCase(m_sFileName)
    End If
    
    'convert sFile to Byte-Array
    Str2ByteArray sFile, arrData
    
    nFlags = CLng(eCopyFlagOverwriteExist Or eCopyFlagOverwriteReadonly)
    itfSyncFile.CopyMem2File arrData, sSrcName, nFlags
    
Exit_DownLoadMemToRobot:
    DownLoadMemToRobot = True
    Exit Function
Err_DownLoadMemToRobot:
    HandleLastError itfSyncFile
    Err.Clear
    DownLoadMemToRobot = False
    Exit Function
End Function

Public Function UpLoadFromRobotToDisk(sFileName As String, sPath As String, Optional sOptions As String, Optional vTimeOut As Long = nTimeOut_UpDownload) As Boolean
On Error GoTo Err_UpLoadFromRobotToDisk
    
'Upload from a file from Robot to Disk
'sPath: format e.g. c:\temp\
    Dim sLocalOptions As String
    Dim srcFile As String
    Dim destFile As String
    Dim sFile As String
    
    Dim sSrcName As String
    Dim sDestName As String
    Dim nFlags As Long
    Dim m_sFileName As String
    Dim bDummy As Boolean
     
    Debug.Print "--- UpLoadFromRobotToDisk ---"
    If Right(sPath, 1) <> "\" Then sPath = sPath & "\"
    ' leave if no Cross is connected
    If Not m_CrossIsConnected Then UpLoadFromRobotToDisk = False: Exit Function
    srcFile = sFileName
    bDummy = GetFileName(srcFile)
    sSrcName = itfSyncFile.GetFullName("/R1/" & srcFile)
    If Not sSrcName = "" Then
        destFile = sFileName
        If GetFileName(destFile) Then
            sDestName = sPath & destFile
            'existing file will be overwritten
            nFlags = CLng(eCopyFlagOverwriteExist Or eCopyFlagOverwriteReadonly)
            itfSyncFile.Copy sSrcName, sDestName, nFlags
         End If
    Else
        srcFile = "krc:\r1\" & sFileName
        sDestName = sPath & sFileName
        'existing file will be overwritten
        nFlags = CLng(eCopyFlagOverwriteExist Or eCopyFlagOverwriteReadonly)
        itfSyncFile.Copy sSrcName, sDestName, nFlags
    End If
    
Exit_UpLoadFromRobotToDisk:
    UpLoadFromRobotToDisk = True
    Exit Function
Err_UpLoadFromRobotToDisk:
    HandleLastError itfSyncFile
    Err.Clear
    UpLoadFromRobotToDisk = False
    Exit Function
End Function

Public Function UpLoadFromRobotToMem(sFileName As String, sFile As String, Optional sOptions As String, Optional vTimeOut As Long = nTimeOut_UpDownload, Optional bKonvert As Boolean = False) As Boolean
'sFileName: "KRC:\R1\" is standard
'path for files in subdirectories must be in sFileName, e.g. sFileName="Program\test.src"

On Error GoTo Err_UpLoadFromRobotToMem
    Dim sSrcName As String
    Dim m_sFileName As String
    Dim bDummy As Boolean
    Dim nHandle As Long
    Dim strModule As String
    Dim strErrKey As String
    Dim nOpenMode As Long
    Dim eEditMode As EKEditMode
    Dim nFlags As Long
    
    Dim nLines As Long
    Dim eCodeType As EKCodeType
    
    Dim nStart, nCount, i As Long
    Dim vLines As Byte
    
    Debug.Print "--- UpLoadFromRobotToMem ---"
    m_sFileName = sFileName
    bDummy = GetFileName(m_sFileName)
    sSrcName = itfSyncFile.GetFullName("/R1/" + m_sFileName)
   
    'NEU: CopyFile2Mem
    nFlags = CLng(eCopyFlagOverwriteExist Or eCopyFlagOverwriteReadonly)
    sFile = ByteArrayToString(itfSyncFile.CopyFile2Mem(sSrcName, nFlags))

Exit_UpLoadFromRobotToMem:
    UpLoadFromRobotToMem = True
    Exit Function
Err_UpLoadFromRobotToMem:
    HandleLastError itfSyncFile
    Err.Clear
    UpLoadFromRobotToMem = False
    Exit Function
End Function

Public Function DeleteRobotProgram(sPrgName As String, Optional vTimeOut As Long = nTimeOut_Dir) As Boolean
'sFileName: "KRC:\R1\" is standard
'path for files in subdirectories must be in sPrgName, e.g. sPrgName="Program\test.src"

On Error GoTo Err_DeleteRobotProgram
    
    Dim sSrcName As String
    Dim bDestroyAlways As Boolean
    
    Debug.Print "--- DeleteRobotProgram ---"
    sSrcName = sPrgName
    
    If Not GetFileName(sSrcName) Then
        DeleteRobotProgram = False
        Exit Function
    End If
    sSrcName = "/R1/" & sSrcName
    sSrcName = itfSyncFile.GetFullName(sSrcName)
    
    bDestroyAlways = True
    
    itfSyncFile.Delete sSrcName, bDestroyAlways
    
Exit_DeleteRobotProgram:
    DeleteRobotProgram = True
    Exit Function
Err_DeleteRobotProgram:
    HandleLastError itfSyncFile
    Err.Clear
    DeleteRobotProgram = False
    Exit Function
End Function
' File-Handling end
'--------------------------------------------------------------------------------

'--------------------------------------------------------------------------------
' Program-Handling
Public Function RobotLevelStop(Optional vTimeOut As Long = nTimeOut_Standard) As Boolean
On Error GoTo Err_RobotLevelStop
    
    Debug.Print "--- RobotLevelStop ---"
    itfSyncSelect.Stop eInterpreterRobot
    
Exit_RobotLevelStop:
    RobotLevelStop = True
    Exit Function
Err_RobotLevelStop:
    HandleLastError itfSyncSelect
    Err.Clear
    RobotLevelStop = False
    Exit Function
End Function

Public Function ControlLevelStop(Optional vTimeOut As Long = nTimeOut_Standard) As Boolean
On Error GoTo Err_ControlLevelStop
    
    Debug.Print "--- ControlLevelStop ---"
    itfSyncSelect.Stop eInterpreterSubmit
    
Exit_ControlLevelStop:
    ControlLevelStop = True
    Exit Function
Err_ControlLevelStop:
    HandleLastError itfSyncSelect
    Err.Clear
    ControlLevelStop = False
    Exit Function
End Function

Public Function RunControlLevel(Optional vPrgName As String, Optional vTimeOut As Long = nTimeOut_Standard) As Boolean
On Error GoTo Err_RunControlLevel
    
    Dim bShowVar As Boolean
    Dim sLocalPrgName As String
    
    Debug.Print "--- RunControlLevel ---"
    If vPrgName = "" Then
        bShowVar = ShowVar("$PRO_I_O[]", sLocalPrgName)
    Else
        sLocalPrgName = vPrgName
    End If
    ' get text after "="
    If (InStr(sLocalPrgName, "=") > 0) Then sLocalPrgName = Mid(sLocalPrgName, InStr(sLocalPrgName, "=") + 1)
    ' cut off brackets
    If (InStr(sLocalPrgName, "(") > 0) Then sLocalPrgName = Left(sLocalPrgName, InStr(sLocalPrgName, "(") - 1)
    ' cut off "
    While (InStr(sLocalPrgName, Chr(34)) > 0)
        sLocalPrgName = Left(sLocalPrgName, InStr(sLocalPrgName, Chr(34)) - 1) & Mid(sLocalPrgName, InStr(sLocalPrgName, Chr(34)) + 1)
    Wend

    itfSyncSelect.Select sLocalPrgName, "", False

Exit_RunControlLevel:
    RunControlLevel = True
    Exit Function
Err_RunControlLevel:
    HandleLastError itfSyncSelect
    Err.Clear
    RunControlLevel = False
    Exit Function
End Function

Public Function SelectModul(strfile As String, Optional strParam As String, Optional bImplizitCancel As Boolean = False, Optional vTimeOut As Long = nTimeOut_Standard) As Boolean
On Error GoTo Err_SelectModul
     
    Debug.Print "--- SelectModul ---"
    If Not GetFileName(strfile) Then
        SelectModul = False
        Exit Function
    End If
    strfile = "/R1/" & strfile
    strfile = itfSyncFile.GetFullName(strfile)
    
    itfSyncSelect.Select strfile, strParam, bImplizitCancel
    
Exit_SelectModul:
    SelectModul = True
    Exit Function
Err_SelectModul:
    HandleLastError itfSyncSelect
    Err.Clear
    SelectModul = False
    Exit Function
End Function

Public Function CancelModul(Optional vTimeOut As Long = nTimeOut_Standard) As Boolean
On Error GoTo Err_CancelModul
    
    Debug.Print "--- CancelModul ---"
    'vTimeOut = -1: stop and cancel SUBMIT-Interpreter
    If vTimeOut = -1 Then
        ControlLevelStop
        itfSyncSelect.Cancel eInterpreterSubmit
    Else
        itfSyncSelect.Cancel eInterpreterRobot
    End If
    
Exit_CancelModul:
    CancelModul = True
    Exit Function
Err_CancelModul:
    HandleLastError itfSyncSelect
    Err.Clear
    CancelModul = False
    Exit Function
End Function
' Program-Handling end
'--------------------------------------------------------------------------------

'--------------------------------------------------------------------------------
' System-Handling
Public Function ConfirmAll(Optional vTimeOut As Long = nTimeOut_Standard) As Boolean
'On Error GoTo Err_ConfirmAll
    
    Dim itfAsyncError As IKAsyncError
    Dim nMsgHandle As Long
    Dim objConsumeMessage As cCrossComm
    Dim nID As Long
    
    Debug.Print "--- ConfirmAll ---"
    Set objConsumeMessage = Me
    nMsgHandle = 0
    nID = 0
    Set itfAsyncError = itfAdviseMessage
    itfAsyncError.Confirm objConsumeMessage, nID, nMsgHandle
    
Exit_ConfirmAll:
    ConfirmAll = True
    Set itfAsyncError = Nothing
    Set objConsumeMessage = Nothing
    Exit Function
Err_ConfirmAll:
    Err.Clear
    ConfirmAll = False
    Set itfAsyncError = Nothing
    Set objConsumeMessage = Nothing
    Exit Function
End Function

Public Function KrcOk(Optional vTimeOut As Long = nTimeOut_Standard) As Boolean
On Error GoTo Err_KrcOk

    KrcOk = m_KrcOk
    
Exit_KrcOk:
    KrcOk = True
    Exit Function
Err_KrcOk:
    Err.Clear
    KrcOk = False
    Exit Function
End Function

Public Function IO_Restart(nBus As IO_TYPE, Optional vTimeOut As Long = nTimeOut_Standard) As Boolean
On Error GoTo Err_IO_Restart
    Dim nParam As Long
    
    Debug.Print "--- IO_Restart ---"
    nParam = 0
    
    itfSyncIo.IoControl eIOControlRestart, nBus, nParam
    
Exit_IO_Restart:
    IO_Restart = True
    Exit Function
Err_IO_Restart:
    HandleLastError itfSyncIo
    Err.Clear
    IO_Restart = False
    Exit Function
End Function
' System-Handling end
'--------------------------------------------------------------------------------

'--------------------------------------------------------------------------------
' Properties
Public Property Get CrossIsConnected() As Boolean
' tell if cross is connected
    CrossIsConnected = m_CrossIsConnected
End Property

Public Property Get CrossError() As String
' last cross error
    CrossError = sLastError
End Property
' Properties end
'--------------------------------------------------------------------------------


'--------------------------------------------------------------------------------
' Error-Handling
Private Sub HandleLastError(ByRef ItfErr)
On Error GoTo Err_HandleLastError

    Dim m_ItfSyncErr As IKSyncError
    Dim m_ItfAsyncErr As IKAsyncError
    
    Dim strMsg As String
    Dim i As Integer
    Dim errData As TKMessage
    
    Debug.Print "--- HandleLastError ---"
    
    If TypeOf ItfErr Is IKSyncError Then
        Set m_ItfSyncErr = ItfErr
        errData = m_ItfSyncErr.GetLastError
    End If
    If TypeOf ItfErr Is IKAsyncError Then
        Set m_ItfAsyncErr = ItfErr
        'errData = m_ItfAsyncErr.GetLastError
    End If

    strMsg = "Error: bDBCause=" + CStr(errData.bDBCause)
    strMsg = strMsg + " , bstrCause=" + CStr(errData.bstrCause)
    strMsg = strMsg + " , bstrDBMessage=" + CStr(errData.bstrDBMessage)
    strMsg = strMsg + " , bstrDBModule=" + CStr(errData.bstrDBModule)
    strMsg = strMsg + " , eMessageType=" + CStr(errData.eMessageType)
    strMsg = strMsg + " , nBiasTimeStamp=" + CStr(errData.nBiasTimeStamp)
    strMsg = strMsg + " , nHighTimeStamp=" + CStr(errData.nHighTimeStamp)
    strMsg = strMsg + " , nLowTimeStamp=" + CStr(errData.nLowTimeStamp)
    strMsg = strMsg + " , nMessage=" + CStr(errData.nMessage)
    strMsg = strMsg + " , nMessageHandle=" + CStr(errData.nMessageHandle)

    For i = LBound(errData.varParams) To UBound(errData.varParams)
        strMsg = strMsg + " , varParams(" + CStr(i) + ") =" + CStr(errData.varParams(i))
    Next i

    For i = LBound(errData.varDBParams) To UBound(errData.varDBParams)
        strMsg = strMsg + " , varDBParams(" + CStr(i) + ") =" + CStr(errData.varDBParams(i))
    Next i

    sLastError = strMsg

Exit_HandleLastError:
    Exit Sub
Err_HandleLastError:
    strMsg = CStr(Err.Number)
    Err.Clear
    GoTo Exit_HandleLastError
End Sub
' Error-Handling end
'--------------------------------------------------------------------------------

Private Function GetFileName(ByRef sName As String, Optional mask As String) As Boolean
  'liest aus einer kompletten Pfadangabe mit Namen und Extention den Filenamen aus
  Dim n As Long
  Dim char, name  As String
  If Not ((sName Like "*.*") Or (sName Like "*")) Then GetFileName = False: sName = "": Exit Function
    For n = Len(sName) To 1 Step -1
      char = Mid(sName, n, 1)
      If mask <> "*." Then
        If char = "\" Or char = "/" Then Exit For
      End If
      name = char + name
    Next
  If mask = "" Then
    If Not (name Like "*.*") Then
        sName = "": GetFileName = False
    Else
        sName = name: GetFileName = True
    End If
  Else
    If mask = "*." Then
        If InStr(1, name, ".") > 0 Then
            GetFileName = False: sName = "": Exit Function
        Else
            sName = name: GetFileName = True
        End If
    Else
        If (name Like mask) Then
            sName = name: GetFileName = True
        Else
            GetFileName = False: sName = "": Exit Function
        End If
    End If
End If
End Function

Private Function ByteArrayToString(bytArray() As Byte) As String
    Dim sAns As String
    Dim iPos As String
    
    sAns = StrConv(bytArray, vbUnicode)
    iPos = InStr(sAns, Chr(0))
    If iPos > 0 Then sAns = Left(sAns, iPos - 1)
    ByteArrayToString = sAns
 End Function

Public Sub Str2ByteArray(StringIn As String, ByteArray() As Byte)
    ByteArray = StrConv(StringIn, vbFromUnicode)
End Sub

Private Sub Class_Terminate()
    ServerOff
End Sub


'--------------------------------------------------------------------------------
' clsCallbackVar begin
Private Sub ICKCallbackVar_OnSetInfo(ByVal nID As Long, ByVal bstrVal As String)
On Error Resume Next
'## ToDo
    Dim sVarName As String
    Dim i As Long
    
    'get variable name
    For i = 1 To CSetInfo.Count
        If nID = CSetInfo.Item(i).Item("m_nID") Then
            sVarName = CSetInfo.Item(i).Item("sVariableName")
            Exit For
        End If
    Next
    
    bstrVal = sVarName + "=" + bstrVal
    
    'call client routine
    m_Parent.SetInfoResult bstrVal
    m_Parent.Cross1_CmdMessage bstrVal, 0
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ICKCallbackVar_OnSetMultiVar(ByVal nID As Long)
    'Empty
End Sub

Private Sub ICKCallbackVar_OnSetVar(ByVal nID As Long)
    'Empty
End Sub

Private Sub ICKCallbackVar_OnShowMultiVar(ByVal nID As Long, ByVal varVals As Variant)
    'Empty
End Sub

Private Sub ICKCallbackVar_OnShowVar(ByVal nID As Long, ByVal bstrVal As String)
    'Empty
End Sub

Private Sub IKCallbackError_OnDialog(ByVal nID As Long, pDialogData As WBC_KrcLib.TKDialog)
    'Empty
End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub IKCallbackError_OnError(ByVal nID As Long, pErrorData As WBC_KrcLib.TKMessage)
    'Empty
End Sub
' clsCallbackVar end
'--------------------------------------------------------------------------------


'--------------------------------------------------------------------------------
' clsConsumeMessage begin
Private Sub ICKConsumeMessage_OnAddMessage(pMessage As WBC_KrcLib.TKMessage)
On Error Resume Next
    Dim strMsg As String
    Dim i As Integer

    strMsg = CStr(pMessage.nMessageHandle)
    strMsg = strMsg + "," + CStr(pMessage.bDBCause)
    strMsg = strMsg + "," + CStr(pMessage.bstrCause)
    strMsg = strMsg + "," + CStr(pMessage.bstrDBMessage)
    strMsg = strMsg + "," + CStr(pMessage.bstrDBModule)
    strMsg = strMsg + "," + CStr(pMessage.eMessageType)
    strMsg = strMsg + "," + CStr(pMessage.nBiasTimeStamp)
    strMsg = strMsg + "," + CStr(pMessage.nHighTimeStamp)
    strMsg = strMsg + "," + CStr(pMessage.nLowTimeStamp)
    strMsg = strMsg + "," + CStr(pMessage.nMessage)

    For i = LBound(pMessage.varParams) To UBound(pMessage.varParams)
        strMsg = strMsg + "," + CStr(i) + "=" + CStr(pMessage.varParams(i))
    Next i

    For i = LBound(pMessage.varDBParams) To UBound(pMessage.varDBParams)
        strMsg = strMsg + "," + CStr(i) + "=" + CStr(pMessage.varDBParams(i))
    Next i

    'call client routine, either RobMessage or Cross1_Message
    m_Parent.RobMessage pMessage.eMessageType, strMsg
    m_Parent.Cross1_Message pMessage.eMessageType, strMsg
End Sub

Private Sub ICKConsumeMessage_OnAddDialog(pDialogData As WBC_KrcLib.TKDialog)
    'Empty
End Sub

Private Sub ICKConsumeMessage_OnSubMessage(ByVal nMessageHandle As Long)
    'Empty
End Sub
' clsConsumeMessage end
'--------------------------------------------------------------------------------

