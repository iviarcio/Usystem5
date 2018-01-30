VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 1  'vbDataSource
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDigiFort"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private WithEvents rs As ADODB.Recordset
Attribute rs.VB_VarHelpID = -1
'Private WinHttpReq As WinHttp.WinHttpRequest

Private Sub Class_GetDataMember(DataMember As String, Data As Object)
   Set Data = rs
End Sub

Private Sub Class_Initialize()
 ' Create an instance of the Recordset.
   Set rs = New ADODB.Recordset

   ' Set the properties of the Recordset.
   With rs
      .Fields.Append "UID", adVarChar, 10          ' Identificação do Sensor (clsModule.UID)
      .Fields.Append "EvTime", adVarChar, 10       ' Hora (hh:mm:ss) do Evento
      .Fields.Append "Sensor", adVarChar, 20       ' Descrição do Sensor (clsModule.mLocal)
      .Fields.Append "Evento", adVarChar, 20       ' Tipo do Evento no Sensor (string)
      .Fields.Append "Loja", adVarChar, 50         ' Identificação do Local
      .Fields.Append "Message", adVarChar, 50      ' Informações de Câmera/Monitor
      .Fields.Append "Status", adInteger           ' Status do Sensor (index)
      .Fields.Append "Color", adInteger
      .CursorType = adOpenStatic
      .LockType = adLockOptimistic
      .Open
   End With
   
End Sub

Public Property Get Recordset() As Recordset
   Set Recordset = rs
End Property

Public Function AddEvent(ByVal fUID As String, ByVal fSensor As String, ByVal idStatus As Integer, ByVal evStr As String, _
                    ByVal fColor As Integer, ByVal fLoja As String, ByVal camera_monitor As String) As Boolean
   
    Dim lHoraEv As String
    lHoraEv = Format(Now(), "hh:mm:ss")
    
    'Check duplicate events
    AddEvent = False
    Dim strCompare As String
    strCompare = "UID= '" & fUID & "'"
    rs.Find strCompare
    If Not rs.EOF Then
        If rs.Fields("Evento") = evStr Then
            If DateDiff("s", rs.Fields("EvTime"), lHoraEv) <= 30 Then
                Exit Function
            End If
        End If
    End If
    
    'If No duplicate then...
    With rs
      .AddNew
      .Fields.Item("Color") = cColor(fColor)
      .Fields.Item("UID") = fUID
      .Fields.Item("EvTime") = lHoraEv
      .Fields.Item("Sensor") = fSensor
      .Fields.Item("Evento") = evStr
      .Fields.Item("Loja") = fLoja
      .Fields.Item("Message") = camera_monitor
      .Fields.Item("Status") = idStatus
      .Update
    End With
    On Error Resume Next
    Dim lModule As clsModule
    Set lModule = lstModule.Item(fUID)
    ShowCamera lModule, fLoja
    On Error GoTo 0
    AddEvent = True
End Function

Public Sub MoveFirst()
   rs.MoveFirst
End Sub

Public Sub MoveNext()
   rs.MoveNext
   If rs.EOF = True Then
      rs.MoveFirst
   End If
End Sub

Public Sub MovePrevious()
   rs.MovePrevious
   If rs.BOF = True Then
      rs.MoveFirst
   End If
End Sub

Public Sub MoveLast()
   rs.MoveLast
End Sub

Public Sub Delete()
   rs.Delete adAffectCurrent
   rs.MoveNext
End Sub

'Public Sub ShowCamera(ByVal crModule As clsModule, ByVal fLoja As String)
'   Dim strURL As String
'   On Error GoTo camError
'   With crModule
'      If .ServerAddress <> "" And .Camera <> "" And .Monitor <> "" Then
'         If .telaCheia Then .SpotNumber = -1 Else .SpotNumber = NextSpotNumber()
'         If .senha = "" Then
'            strURL = "http://" & .ServerAddress & "/Interface/VirtualMatrix/ShowObject?" & _
'                  "MonitorID=" & .Monitor & "&SpotNumber=" & .SpotNumber & _
'                  "&ObjectType=0&ObjectName=" & .Camera & "&ResponseFormat=Text&AuthUser=" & .user
'         Else
'            strURL = "http://" & .ServerAddress & "/Interface/VirtualMatrix/ShowObject?" & _
'                  "MonitorID=" & .Monitor & "&SpotNumber=" & .SpotNumber & _
'                  "&ObjectType=0&ObjectName=" & .Camera & _
'                  "&ResponseFormat=Text&AuthUser=" & .user & "&AuthPass=" & .senha
'         End If
'      Else
'         Exit Sub
'      End If
'   End With
''   MsgBox "Comando a ser enviado ao CFTV:" & vbCrLf & _
''   strURL & vbCrLf & "Clique em Ok para continuar.", vbInformation + vbOKOnly
'   Set WinHttpReq = New WinHttpRequest
'   WinHttpReq.SetTimeouts "2000", "3000", "2000", "5000"  ' Resolve, Connect, Send and Receive
'   WinHttpReq.Open "POST", strURL, False
'   WinHttpReq.SetRequestHeader "ContentType", "text/plain; encoding='utf-8'"
'   WinHttpReq.SetRequestHeader "Content-Length", Len(strURL)
'   WinHttpReq.Send ""
'   Do While WinHttpReq.WaitForResponse(5)
'     Sleep 200
'     DoEvents
'   Loop
'   If WinHttpReq.status <> 200 Then
'       ForNet.StatusBar1.Panels.Item(2).Text = "Falha na visualização da Câmera " & crModule.Camera & " para " & fLoja & " (" & crModule.mLocal & ")"
'   End If
'   'MsgBox WinHttpReq.ResponseText, vbInformation + vbOKOnly
'   Set WinHttpReq = Nothing
'   Exit Sub
'
'camError:
'    If Err.Number <> 0 Then
'        ForNet.StatusBar1.Panels.Item(2).Text = "Falha na visualização da Câmera em " & fLoja & ". Erro: " & Err.Description
'        Err.Clear
'    End If
'
'End Sub

