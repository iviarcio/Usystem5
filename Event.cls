VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEvent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public SensorId As String        'Identificação do sensor -- obsoleto
Public evDate As String          'Data do evento (dd/mm/yyyy hh:mm:ss)
Public evStr As String           'Informações sobre o Evento
Public evDescr As String         'Descrição da Entidade onde ocorreu o evento
Public evTipo As Integer         'Tipo do Evento (Intrusão, pânico, ...)
Public evKind As kindEvent
Public comm As Integer           'Número da porta que recebeu o evento
Public evHex As String
Public evIn As String
Public sHeader As String
Public ssData As String
Public sUIDo As String           'Identificaçao do sensor (modulo) onde ocorreu o evento.
Public sUIDh As String           'Identificação do Receiver ou Repeater que encaminhou o evento.
Public sPTI As String
Public sClasse As String
Public sStat1 As String
Public sStat0 As String
Public ssLevel As String
Public ssMargin As String
Public sDeviceName As String
Public sMID As String
Public evCritico As Boolean      'Indicação de Evento crítico
Public evScope As Byte            '0: Real, 1: Acidental, 2: Teste
Public evAcao As String          'Ação tomada no tratamento de evento crítico
Public evObs As String           'Observações adicionais
Public evTreat As String         'Hora do tratamento de evento crítico
Public evUser As String          'Usuário que tratou o evento crítico

'Salva o Evento (checksum = 0k e duplicidade = false) na tabela LastEvents do Cadastro
Public Sub Persist()

   Dim lcm As New ADODB.Command
   Set lcm.ActiveConnection = cnDB
   lcm.CommandType = adCmdText

   'Salva evento no Serial Receiver
   If sHeader = H_Serial Then
      lcm.CommandText = "INSERT INTO LastEvents (Header, sData, STAT0, DeviceName, DateEvent) VALUES ('" & _
                        sHeader & bp & ssData & bp & sStat0 & bp & sDeviceName & bp & evDate & "')"
      lcm.Execute
    
   'Salva evento nos Repeaters
   ElseIf sHeader = H_Device And sMID = MID_Repeater Then
      lcm.CommandText = "INSERT INTO LastEvents (Header, UIDo, UIDh, PTI, STAT1, STAT0, sLevel, sMargin, DeviceName, DateEvent) VALUES ('" & _
                        sHeader & bp & sUIDo & bp & sUIDh & bp & sClasse & bp & sStat1 & bp & sStat0 & bp & ssLevel & _
                        bp & ssMargin & bp & sDeviceName & bp & evDate & "')"
      lcm.Execute

   'Salva evento nos Sensores
   ElseIf sHeader = H_Device And sMID = MID_Sensor Then
      lcm.CommandText = "INSERT INTO LastEvents (Header, UIDo, UIDh, PTI, STAT1, STAT0, sLevel, sMargin, DeviceName, DateEvent) VALUES ('" & _
                        sHeader & bp & sUIDo & bp & sUIDh & bp & sPTI & bp & sStat1 & bp & sStat0 & bp & ssLevel & _
                        bp & ssMargin & bp & sDeviceName & bp & evDate & "')"
      lcm.Execute
        
   End If
   DoEvents
   Control_LastEvents
   
End Sub

'Procedure que trata o eventos de Sensor
Public Sub TreatSensor()

'  Significado dos bits de sStat1:
'   Bit 7 -> 0: Normal, 1: EOL Tamper
'   Bits 4, 5 e 6 reservados
'   Bit 3 -> 0: normal, 1 = Alarme p/ 4 sensores
'   Bit 2 -> 0: normal, 1 = Alarme p/ 3 sensores
'   Bit 1 -> 0: normal, 1 = Alarme p/ 2 sensores
'   Bit 0 -> 0: normal, 1 = Alarme Principal
'
'  Significado dos bits de sStat0:
'   Bit 7 reservado
'   Bit 6 -> 0: normal, 1: Bateria Fraca
'   Bit 5 -> 0: normal, 1: Tamper
'   Bit 4 -> 0: status modificado, 1: Status s/ modificação
'   Bit 3 -> 0: normal, 1: Reset do Sensor
'   Bits 0, 1 e 2 reservados

'  First, verify the receiver
   Dim tDisplay As clsDisplay
   Dim tDevice As clsDevice
   Dim tReceiver As clsModule
   Set tReceiver = Nothing
   On Error Resume Next
   Set tReceiver = lstModule.Item(sUIDh)
   On Error GoTo 0
   If tReceiver Is Nothing Then
      'receiver não cadastrado
      Set tDevice = Nothing
      On Error Resume Next
      Set tDevice = lstDevice.Item(sUIDh)
      On Error GoTo 0
      If tDevice Is Nothing Then
         Set tDevice = New clsDevice
         tDevice.sUID = sUIDh
         tDevice.Serial = CStr(Int("&H" + Right(sUIDh, 6)))
         tDevice.tipo = "Receiver"
         tDevice.evDate = evDate
         tDevice.level = ""
         tDevice.margin = ""
         tDevice.comm = comm
         tDevice.name = "Serial Receiver"
         lstDevice.Add Item:=tDevice, Key:=sUIDh
         
         Set tDisplay = New clsDisplay
         tDisplay.dispMode = sxErSound
         tDisplay.dispFile = ""
         tDisplay.dispStr = "Warning: Device (Receiver/Repeater) não cadastrado (" & sUIDh & ")"
         tDisplay.dispImg = sxImgAlert
         Insert_Display tDisplay, False, False, fNoSound:=True
         If m_Debug Then Make_Service "Device (Receiver/Repeater) não cadastrado (" & sUIDh & ")", " "
                  
      Else
         ' Nothing to do. Message already displayed!
      End If
   Else
      tReceiver.mStatAtiv = True
      tReceiver.mLastAtiv = evDate
      tReceiver.NivelSinal = ssLevel
   End If
   
'  Now, verify the Sensor
   Dim serialSensor As String
   serialSensor = CStr(Int("&H" + Right(sUIDo, 6)))
   Dim tModule As clsModule
   Set tModule = Nothing
   On Error Resume Next
   Set tModule = lstModule.Item(sUIDo)
   On Error GoTo 0
   If tModule Is Nothing Then
      'modulo/device não identificado/cadastrado
      Set tDevice = Nothing
      On Error Resume Next
      Set tDevice = lstDevice.Item(sUIDo)
      On Error GoTo 0
      If tDevice Is Nothing Then
         Set tDevice = New clsDevice
         tDevice.sUID = sUIDo
         tDevice.Serial = serialSensor
         tDevice.tipo = "Sensor"
         tDevice.comm = comm
         tDevice.name = sDeviceName
         tDevice.level = ssLevel
         tDevice.margin = ssMargin
         tDevice.recep = CStr(Int("&H" + Right(sUIDh, 6)))
         tDevice.evDate = evDate
         lstDevice.Add Item:=tDevice, Key:=sUIDo
         
         Set tDisplay = New clsDisplay
         tDisplay.dispMode = sxErSound
         tDisplay.dispFile = ""
         tDisplay.dispStr = "Erro: Evento em Sensor não cadastrado (" & sUIDo & ")"
         tDisplay.dispImg = sxImgAlert
         Insert_Display tDisplay, False, False, fNoSound:=True
         If m_Debug Then Make_Service "Evento em Sensor não cadastrado (" & sUIDo & ")", " "
         
         Exit Sub
      Else
         ' Nothing to do. Message already displayed!
         Exit Sub
      End If
   End If
    
'  Flag p/ Persistencia do evento
   Dim flag As Boolean
   flag = False
      
   Dim status1 As Byte
   status1 = "&H" + sStat1
   Dim status0 As Byte
   status0 = "&H" + sStat0
   
   With tModule
      .mStatAtiv = True
      .mLastAtiv = evDate
      .NivelSinal = ssLevel
      
      If .mPTI = "" Then .AtualizaPTI sDeviceName
     
      'Trata Bateria Fraca
      Dim lowbat As Byte
      lowbat = status0 And &H50
      If lowbat = &H40 Then
         evStr = "Bateria Fraca no Sensor: " + serialSensor
         .TreatBateria True, Me
         flag = True
      ElseIf lowbat = 0 Then
         evStr = "Bateria normalizada."
         .TreatBateria False, Me
      End If
      
      'Trata Reset no Sensor
      Dim reset As Byte
      reset = status0 And &H18
      If reset = &H8 Then
         evStr = "Ocorreu Reset no Sensor: " + serialSensor
         .TreatReset True, Me
         flag = True
      ElseIf reset = 0 Then
         evStr = ""
         .TreatReset False, Me
      End If
         
      'Trata o Alarme de Vandalismo ou Tamper
      Dim tampa1 As Byte
      tampa1 = status1 And &H80
      Dim tampa2 As Byte
      tampa2 = status0 And &H30
      
      If tampa1 = &H80 Then
         evStr = "Alarme de Vandalismo no Sensor: " + serialSensor
         .TreatTampa True, Me
         flag = True
      ElseIf tampa2 = &H20 Then
         evStr = "Sensor de Intrusão retirado do local: " + serialSensor
         .TreatTampa True, Me
         flag = True
      ElseIf tampa2 = 0 Then
         evStr = "Sensor normalizado: " + serialSensor
         .TreatTampa False, Me
      End If
      
         
      'Trata os demais Alarmes
'      evStr = ""
      Dim ls1, ls2, ls3, ls4 As Byte
      ls1 = status1 And &H1
      ls2 = status1 And &H2
      ls3 = status1 And &H4
      ls4 = status1 And &H8
      If (ls1 = 0) And (ls2 = 0) And (ls3 = 0) And (ls4 = 0) Then
         .TreatAlarme stFechado, Me
      Else
         evStr = "( "
         If ls1 = 1 Then evStr = evStr + "Alarme 1 "
         If ls2 = 2 Then evStr = evStr + "Alarme 2 "
         If ls3 = 4 Then evStr = evStr + "Alarme 3 "
         If ls4 = 8 Then evStr = evStr + "Alarme 4 "
         evStr = evStr + ")"
         flag = True

         Dim lEntity As clsEntity
         Set lEntity = lstEntity.Item(CStr(.mEntity))
         
         On Error Resume Next
         If .critico Then
            Dim showEvt As Boolean
            Dim lColor As Integer
            lColor = .crColor
            showEvt = tFort.AddEvent(.UID, .mLocal, stAberto, evStr, _
                           lColor, lEntity.vDescr, .Camera & " / " & .Monitor)
                           
            If showEvt Then
                frmQueue.Show
            Else
                ForNet.Update_Display "Evento em " & .mLocal & " - " & _
                lEntity.vDescr & " descartado por duplicidade!", sxImgInform, False, True
                Exit Sub
            End If
         ElseIf .ShowEvent_in_Camera Then
            ShowCamera tModule, lEntity.vDescr
            If tModule.popup And .SpotNumber <> -1 Then
               .SpotNumber = FreeSpotNumber(.SpotNumber)
            End If
         End If
         On Error GoTo 0
         
         .TreatAlarme stAberto, Me
                 
      End If
    
    End With
   
End Sub

'Procedure que trata o eventos do Serial Receiver
Public Sub TreatReceiver()

'  Significado dos bits de sStat0:
'   Bit 7 -> 0: normal, 1: Ruido excessivo
'   Bit 5 -> 0: normal, 1: Alarme de Tamper
'   Bit 4 -> 0: status modificado, 1: Status s/ modificação
'   Bit 3 -> 0; normal, 1: ocorreu Reset
'   Bit 0 -> 0: normal, 1: Falha de Link

   'A msg de supervisão do Receiver não vem com seu UID
   'Busca-se se existe algum Receiver cadastrado no Sistema
   Dim device As clsModule
   Dim tModule As clsModule
   Set tModule = Nothing
   For Each device In lstModule
      If device.mTipo = d_Receiver Then
         Set tModule = device
         sUIDo = tModule.UID
         Exit For
      End If
   Next
   If tModule Is Nothing Then
      'Serial Receiver não identificado/cadastrado
      Dim tDevice As clsDevice
      Set tDevice = New clsDevice
      tDevice.sUID = ""
      tDevice.Serial = "Não Cadastrado"
      tDevice.tipo = "Receiver"
      tDevice.comm = comm
      tDevice.name = sDeviceName
      tDevice.evDate = evDate
      On Error Resume Next
      lstDevice.Add Item:=tDevice, Key:=sUIDo
      On Error GoTo 0
      
      Set tDisplay = New clsDisplay
      tDisplay.dispMode = sxErSound
      tDisplay.dispFile = ""
      tDisplay.dispStr = "Não existe Receptor cadastrado no Sistema"
      tDisplay.dispImg = sxImgAlert
      Insert_Display tDisplay, False, False, fNoSound:=True
      If m_Debug Then Make_Service "Não existe Receptor cadastrado no Sistema", " "
      
      Exit Sub
   End If
    
'  Flag p/ Persistencia do evento
   Dim flag As Boolean
   flag = False
      
   Dim status0 As Byte
   status0 = "&H" + sStat0
   
   With tModule
      .mStatAtiv = True
      .mLastAtiv = evDate
      .NivelSinal = ssLevel
      'Trata Ruído Excessivo
      Dim has_ruido As Byte
      has_ruido = status0 And &H90
      If has_ruido = &H80 And Not .SRuido Then
         evStr = "Ruído excessivo em Receptor. "
         .TreatRuido True, Me
         flag = True
      ElseIf has_ruido = 0 Then
         evStr = ""
         .TreatRuido False, Me
      End If
      
      'Trata o Alarme de Tamper
      Dim tampa As Byte
      tampa = status0 And &H30
      If tampa = &H20 Then
         evStr = "Tampa Aberta em Receptor. "
         .TreatTampa True, Me
         flag = True
      ElseIf tampa = 0 Then
         evStr = "Tampa Fechada. "
         .TreatTampa False, Me
      End If
      
      'Trata Reset no Serial
      Dim reset As Byte
      reset = status0 And &H18
      If reset = &H8 Then
         evStr = "Ocorreu Reset em Receptor. "
         .TreatReset True, Me
         flag = True
      ElseIf reset = 0 Then
         evStr = ""
         .TreatReset False, Me
      End If
      
      'Trata Falha de Link
      Dim link As Byte
      link = status0 And &H11
      If link = 1 Then
         evStr = "Falha de Link em Receptor. "
         .TreatLink True, Me
         flag = True
      ElseIf link = 0 Then
         .TreatLink False, Me
      End If
    
      'If flag Then Me.Persist
    End With
   
End Sub

'Procedure que trata o eventos do Repeater
Public Sub TreatRepeater()

'  Significado dos bits de sStat0:
'   Bit 7 -> 0: normal, 1: Ruido excessivo
'   Bit 6 -> 0: normal, 1: Bateria Fraca
'   Bit 5 -> 0: normal, 1: Alarme de Tamper
'   Bit 4 -> 0: status modificado, 1: Status s/ modificação
'   Bit 1 -> 0: normal, 1: Perda de rede AC

   Dim serialRp As String
   serialRp = CStr(Int("&H" + Right(sUIDo, 6)))
   
   Dim tModule As clsModule
   Set tModule = Nothing
   On Error Resume Next
   Set tModule = lstModule.Item(sUIDo)
   On Error GoTo 0
   If tModule Is Nothing Then
      'modulo/device não identificado/cadastrado
      Dim tDevice As clsDevice
      Set tDevice = Nothing
      On Error Resume Next
      Set tDevice = lstDevice.Item(sUIDo)
      On Error GoTo 0
      If tDevice Is Nothing Then
         Set tDevice = New clsDevice
         tDevice.sUID = sUIDo
         tDevice.Serial = serialRp
         tDevice.tipo = "Repeater"
         tDevice.level = ssLevel
         tDevice.margin = ssMargin
         tDevice.name = sDeviceName
         tDevice.comm = comm
         tDevice.recep = CStr(Int("&H" + Right(sUIDh, 6)))
         tDevice.evDate = evDate
         lstDevice.Add Item:=tDevice, Key:=sUIDo
         
         Set tDisplay = New clsDisplay
         tDisplay.dispMode = sxErSound
         tDisplay.dispFile = ""
         tDisplay.dispStr = "Erro: Repetidor não cadastrado (" & sUIDo & ")"
         tDisplay.dispImg = sxImgAlert
         Insert_Display tDisplay, False, False, fNoSound:=True
         If m_Debug Then Make_Service "Repetidor não cadastrado (" & sUIDo & ")", " "
         
         Exit Sub
      Else
         ' Nothing to do. Message already displayed!
         Exit Sub
      End If
   End If
    
'  Flag p/ Persistencia do evento
   Dim flag As Boolean
   flag = False
      
   Dim status0 As Byte
   status0 = "&H" + sStat0
   
   With tModule
      .mStatAtiv = True
      .mLastAtiv = evDate
      .NivelSinal = ssLevel
      
      'Trata Ruído Excessivo
      Dim ruido As Byte
      ruido = status0 And &H90
      If ruido = &H80 Then
         evStr = "Ruído excessivo no Repetidor: " + serialRp
         .TreatRuido True, Me
         flag = True
      ElseIf ruido = 0 Then
         evStr = ""
         .TreatRuido False, Me
      End If
      
      'Trata Bateria Fraca
      Dim lowbat As Byte
      lowbat = status0 And &H50
      If lowbat = &H40 Then
         evStr = "Bateria Fraca no Repetidor: " + serialRp
         .TreatBateria True, Me
         flag = True
      ElseIf lowbat = 0 Then
         evStr = "Bateria normalizada."
         .TreatBateria False, Me
      End If
      
      'Trata o Alarme de Tamper
      Dim tampa As Byte
      tampa = status0 And &H30
      If tampa = &H20 Then
         evStr = "Tampa Aberta no Repetidor: " + serialRp
         .TreatTampa True, Me
         flag = True
      ElseIf tampa = 0 Then
         evStr = "Tampa Fechada. "
         .TreatTampa False, Me
      End If
      
      'Trata Perda de Rede AC
      Dim rede As Byte
      rede = status0 And &H12
      If rede = 2 Then
         evStr = "Perda de Rede AC no Repetidor: " + serialRp
         .TreatRedeAC True, Me
         flag = True
      ElseIf rede = 0 Then
         .TreatRedeAC False, Me
      End If
      
      'If flag Then Me.Persist
    End With
   
End Sub
