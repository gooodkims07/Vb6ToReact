Attribute VB_Name = "adoConst"
Option Explicit

Public adoConnect       As ADODB.Connection
Public adoConnectActive       As ADODB.Connection   'modPassw UPDATE_LOGIN_STATUS에서만 사용
Public adoSet           As ADODB.Recordset
Public adoCommand       As ADODB.Command
Public adoFunc          As ADODB.Recordset

Public lngExeCount      As Long
Public strOK            As String

Public GsConnUser           As ConnUserId
Public GsConnPass           As String
Public GsDataSrc        As String

Public Const PROVIDER_OracleOLE = "OraOLEDB.Oracle.1"
Public Const PROVIDER_Microsoft = "Microsoft OLE DB Provider for Oracle"


'Public Function adoDbConnect(ByVal sUser As String, ByVal sPassword As String, ByVal sDataSRC As String) As Integer
'    Dim sConString          As String
'    Dim strProvider         As String
'
'
'    sConString = ""
'    GsUser = sUser
'    GsPass = sPassword
'    GsDataSrc = sDataSRC
'
'    sConString = sConString & "Provider=" & PROVIDER_Microsoft & ";"
'    sConString = sConString & "User ID=" & sUser & ";"
'    sConString = sConString & "Data Source=" & sDataSRC & ";"
'    sConString = sConString & "Persist Security info=False"
'
'    On Error GoTo DBConnect_Error
'
'    Set adoConnect = New ADODB.Connection
'    adoConnect.CursorLocation = adUseClient
'    adoConnect.Open sConString, sUser, sPassword
'
'    Exit Function
'
'
'DBConnect_Error:
'    MsgBox adoConnect.Errors(0).Number & vbCrLf & _
'           adoConnect.Errors(0).Description & vbCrLf & vbCrLf & _
'           "ConnectString : " & sConString
'
'    Return
'
'End Function

Public Function adoDbConnectEnc(pConnUserId As ConnUserId, ByVal sEncPassword As String, ByVal sDataSRC As String) As Integer
    Dim sConString          As String
    Dim strProvider         As String
    
     Dim sUser               As String
    Dim sPassword           As String
    
    On Error GoTo DecryptErrorHandler
    sUser = GetConnUserId(pConnUserId)
    sPassword = DecryptCredential(sEncPassword)
    On Error GoTo 0
    
    sConString = ""
    GsConnUser = pConnUserId
    GsConnPass = sEncPassword
    GsDataSrc = sDataSRC
   
    sConString = sConString & "Provider=" & PROVIDER_Microsoft & ";"
    sConString = sConString & "User ID=" & sUser & ";"
    sConString = sConString & "Data Source=" & sDataSRC & ";"
    sConString = sConString & "Persist Security info=False"

    On Error GoTo DBConnect_Error
    
    Set adoConnect = New ADODB.Connection
    adoConnect.CursorLocation = adUseClient
    adoConnect.Open sConString, sUser, sPassword
    
    Exit Function
    
DecryptErrorHandler:
   MsgBox "접속 정보가 올바르지 않습니다." & vbCrLf & _
                                       "전산개발팀으로 연락 바랍니다.", vbExclamation, "경고"
    Exit Function
DBConnect_Error:
    MsgBox adoConnect.Errors(0).Number & vbCrLf & _
           adoConnect.Errors(0).Description & vbCrLf & vbCrLf & _
           "ConnectString : " & sConString
           
    Return

End Function

Public Function adoDbDisconnect() As Integer
    
    On Error GoTo Disconnect_Error
    
    adoConnect.Close
    If Not adoConnect Is Nothing Then
        Set adoConnect = Nothing
    End If
    
    Exit Function

Disconnect_Error:
    Resume Next
    Return
    
End Function

Public Function adoSetOpen(ByVal sSql As String, ByRef sAdoset As ADODB.Recordset, Optional ByVal sRecCnt As Integer = 0, _
                                                                                   Optional ByVal sGbOpen As Integer = 0) As Integer
    
    Dim strOpen     As String
       
    On Error GoTo SetOpen_Error
       
    If adoConnect Is Nothing Then
        Set adoConnect = Nothing
        Set adoConnect = New ADODB.Connection
        Call adoDbConnectEnc(GsConnUser, GsConnPass, GsDataSrc)
    ElseIf adoConnect.State = 0 Then
        Call adoDbConnectEnc(GsConnUser, GsConnPass, GsDataSrc)
    End If
    
    Set sAdoset = New ADODB.Recordset

    If sRecCnt > 0 Then 'Read할 Record 수
        sAdoset.MaxRecords = sRecCnt
    End If
       
    Select Case Val(sGbOpen)
        Case 0: strOpen = adOpenStatic
        Case 1: strOpen = adOpenDynamic
        Case 2: strOpen = adOpenForwardOnly
        Case 3: strOpen = adOpenKeyset
    End Select
    
    Call sAdoset.Open(sSql, adoConnect, strOpen, adLockReadOnly, adCmdText)

    If sAdoset.RecordCount = 0 Then
        adoSetOpen = False
    Else
        adoSetOpen = True
    End If
        
    Exit Function
    
    
SetOpen_Error:
    
    adoSetOpen = False
    MsgBox adoConnect.Errors(0).Number & vbCrLf & _
           adoConnect.Errors(0).Description & vbCrLf & _
           sSql
    
End Function

Public Function adoCmd(ByVal sSql As String, ByVal cmdType As Integer) As Integer
    
    Dim strOpen     As String
       
    On Error GoTo SetOpen_Error
       
    If adoConnect Is Nothing Then
        Set adoConnect = Nothing
        Set adoConnect = New ADODB.Connection
        Call adoDbConnectEnc(GsConnUser, GsConnUser, GsDataSrc)
    ElseIf adoConnect.State = 0 Then
        Call adoDbConnectEnc(GsConnUser, GsConnUser, GsDataSrc)
    End If
    
    Set adoCommand = New ADODB.Command
    
    With adoCommand
        .ActiveConnection = adoConnect
        .CommandText = sSql
        .CommandType = cmdType
    End With
    
    adoCmd = True
    Exit Function
    
SetOpen_Error:
    adoCmd = False
    MsgBox adoConnect.Errors(0).Number & vbCrLf & _
           adoConnect.Errors(0).Description & vbCrLf & _
           sSql
    
End Function

Public Function adoExecute(ByVal sSql As String, Optional nRetCount As Integer) As Integer
    
    
    On Error GoTo SetOpen_Error
    
    adoExecute = True
    Call adoConnect.Execute(sSql, nRetCount, adCmdText + ADODB.adExecuteNoRecords)
    adoExecute = True
    Exit Function
    
SetOpen_Error:
    adoExecute = False
    MsgBox adoConnect.Errors(0).Number & vbCrLf & _
           adoConnect.Errors(0).Description & vbCrLf & _
           sSql
    Exit Function
    Return
    
End Function


Public Function adoSetClose(ByRef sAdoset As ADODB.Recordset) As Integer
    
    On Error GoTo SetClose_Error
    
    sAdoset.Close
    If Not sAdoset Is Nothing Then Set sAdoset = Nothing
    
    Exit Function
    
    
SetClose_Error:
    MsgBox adoConnect.Errors(0).Number & vbCrLf & _
           adoConnect.Errors(0).Description
    adoSetClose = False
    
    Exit Function
    Return

End Function

' RECORDSET을 이용하여 UPDATE BATCH를 하는 경우의  RECORDSET생성 함수 (ADOOPENSET을 사용하지않는다)

Public Function CreateRecordset(argSql As String) As ADODB.Recordset
    
    Dim RsTemp As ADODB.Recordset
    
    Set RsTemp = New ADODB.Recordset
    Screen.MousePointer = vbHourglass
    
    On Error GoTo OpenError:
    
    Set CreateRecordset = Nothing
    
    adoConnect.CursorLocation = adUseClient
    
    Call RsTemp.Open(argSql, adoConnect, adOpenStatic, adLockBatchOptimistic, adCmdText)
    
    If Not RsTemp Is Nothing Then Set CreateRecordset = RsTemp
    Screen.MousePointer = vbDefault
    Exit Function
            
OpenError:
    
    MsgBox adoConnect.Errors(0).Number & vbCrLf & _
           adoConnect.Errors(0).Description & vbCrLf & _
           argSql
    Set CreateRecordset = Nothing
        Exit Function
        
End Function

Public Function UpdateBatch(ByVal rsSource As ADODB.Recordset) As Boolean

    '/UPDATE BATCH는 CREATERECORDSET함수를 이용하여 SET을 생성한 경우만 가능하다
    
    Dim RsTemp  As ADODB.Recordset
    Dim i       As Integer
    Dim StrTemp As String
    Dim blError As Boolean
    
    UpdateBatch = True
    
    Set RsTemp = rsSource
    If RsTemp.Supports(adUpdateBatch) = False Then Exit Function 'MsgBox gsMSGUPDATEBATCH, vbCritical, "ERROR"
    
    On Error GoTo FindError:
    RsTemp.Filter = adFilterAffectedRecords
    RsTemp.UpdateBatch
    
    Set RsTemp.ActiveConnection = Nothing
    Set RsTemp = Nothing
        
    Exit Function
    
FindError:

    If adoConnect.Errors.COUNT > 0 Then
        
        'INSERT나 UPDATE 처리과정중에 생긴 ERROR (AdoConnection 개체에 내장되어 있는 에러집합)
        MsgBox adoConnect.Errors(0).Number & vbCrLf & adoConnect.Errors(0).Description & vbCrLf
        UpdateBatch = False
        Resume Next
        Exit Function
    Else
        MsgBox Err.Description
    End If
End Function


Public Function adoDbConnectActiveEnc(pConnUserId As ConnUserId, ByVal sEncPassword As String, ByVal sDataSRC As String) As Integer
    Dim sConString          As String
    Dim strProvider         As String
    
    Dim sUser               As String
    Dim sPassword           As String
    
    On Error GoTo DecryptErrorHandler
    sUser = GetConnUserId(pConnUserId)
    sPassword = DecryptCredential(sEncPassword)
    On Error GoTo 0
    
    
    sConString = ""
    GsConnUser = pConnUserId
    GsConnPass = sEncPassword
    GsDataSrc = sDataSRC
   
    sConString = sConString & "Provider=" & PROVIDER_Microsoft & ";"
    sConString = sConString & "User ID=" & sUser & ";"
    sConString = sConString & "Data Source=" & sDataSRC & ";"
    sConString = sConString & "Persist Security info=False"

    On Error GoTo DBConnect_Error
    
    Set adoConnectActive = New ADODB.Connection
    adoConnectActive.CursorLocation = adUseClient
    adoConnectActive.Open sConString, sUser, sPassword
    
    Exit Function
    
DecryptErrorHandler:
   MsgBox "접속 정보가 올바르지 않습니다." & vbCrLf & _
                                       "전산개발팀으로 연락 바랍니다.", vbExclamation, "경고"
    Exit Function
DBConnect_Error:
    MsgBox adoConnectActive.Errors(0).Number & vbCrLf & _
           adoConnectActive.Errors(0).Description & vbCrLf & vbCrLf & _
           "ConnectString : " & sConString
           
    Exit Function

End Function

Public Function adoDbDisconnectActive() As Integer
    
    On Error GoTo Disconnect_Error
    
    adoConnectActive.Close
    If Not adoConnectActive Is Nothing Then
        Set adoConnectActive = Nothing
    End If
    
    Exit Function

Disconnect_Error:
    Resume Next
    Return
    
End Function



Public Function adoCmdActive(ByVal sSql As String, ByVal cmdType As Integer) As Integer
    
    Dim strOpen     As String
       
    On Error GoTo SetOpen_Error
       
    If adoConnectActive Is Nothing Then
        Set adoConnectActive = Nothing
        Set adoConnectActive = New ADODB.Connection
        Call adoDbConnectActiveEnc(GsConnUser, GsConnPass, GsDataSrc)
    ElseIf adoConnectActive.State = 0 Then
        Call adoDbConnectActiveEnc(GsConnUser, GsConnPass, GsDataSrc)
    End If
    
    Set adoCommand = New ADODB.Command
    
    With adoCommand
        .ActiveConnection = adoConnectActive
        .CommandText = sSql
        .CommandType = cmdType
    End With
    
    adoCmdActive = True
    Exit Function
    
SetOpen_Error:
    adoCmdActive = False
    MsgBox adoConnectActive.Errors(0).Number & vbCrLf & _
           adoConnectActive.Errors(0).Description & vbCrLf & _
           sSql
    
End Function



