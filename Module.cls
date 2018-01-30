VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsModule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum typeStatus
   stDesabilitada = 0
   stDesativada = 1
   stAtivada = 2
   stProgramada = 3
End Enum

Public Enum typeVerify
   vrSimples = 0
   vrDupla = 1
End Enum

Public Enum typeLogica
   lgNenhuma = 0
   lgAND = 1
End Enum

Public Serial_Number As String   'No. de série do Módulo (Device)
Public UID As String             'Identificação do Módulo - UID
Public mNumero As Integer        'Número da Zona (Módulo)
Public mPTI As String            'PTI
Public mLocal As String          'Local da Zona
Public mEntity As Long           'Id para a Entidade
Public mTipo As typeSensor       'Tipo do Módulo/device (Receiver, Repeater ou Sensor)
Public mCheck As typeVerify      'Tipo de Verificação
Public mJanela As Byte           'Janela de Tempo de Verificação
Public mLogica As typeLogica     'Tipo de Lógica
Public mNumLogica As Byte        'UID da Zona relacionada
Public mLocalLogica As String    'Local da Zona relacionada
Public mArquivo As String        'Arquivo de som
Public mChkAtiv As Boolean       'Flag de verificação de Atividade
Public mtempoAtiv As Long        'Tempo de verificação de Atividade (1 a 59 minutos)
Public mStatAtiv As Boolean      'Flag de Atividade
Public mLastAtiv As Date         'Data/Hora da última atividade
Public evDate As Date            'Data/Hora do último evento

Private mflagDupla As Boolean    'Flag de dupla verificação
Private mService As clsService   'Reference to current Service
Private mInicial As typeStatus   'Status inicial do Módulo
Private mStatus As typeStatus    'Status corrente do Módulo
Private tampa As Boolean         'False = tampa fechada, True = tampa aberta
Private reset As Boolean         'False = normal, True = indica a ocorrencia de reset
Private link As Boolean          'False = normal, True = indica falha de link
Private ruido As Boolean         'false = normal, True = indica ruido excessivo
Private rede As Boolean          'false = normal, True = indica falha de rede AC
Private lowbat As Boolean        'false = bateria Ok, True = bateria fraca
Private mSZona As typeZona       'Situação corrente da Zona
Private nivel_Sinal As Byte      'Níveis de sinal 0 a 8
Public critico As Boolean        'Sinaliza evento crítico (será apresentado via Popup)
Public crColor As Integer        'Cor de fundo do evento na tela de Popups (0: vbRed, ...)
Public crKind As Byte            '0 =  real, 1 = acidental, 2 = teste
Public crAcao As String          'Última ação tomada
Public crObs As String           'Última observação
Public crUser As String          'Usuário que restritou a ação
Public crTreat As String         'hh:mm:ss do tratamento do último evento

Public ServerAddress As String   'Endereço servidor digiFort
Public Camera As String          'Camera associada ao módulo
Public Monitor As String         'Monitor para visualização da Camera
Public SpotNumber As Integer     'Spot Number na Matriz Virtual. -1 para tela cheia
Public telaCheia As Boolean      'Indicativo de visualização em tela cheia
Public user As String            'Usuário de acesso ao servidor CFTV
Public senha As String           'Senha de acesso ao Servidor CFTV
Public popup As Boolean          'Sinaliza se mostra a camera ou não no Monitor selecionado

Private Sub Class_Initialize()
   mTipo = s_Incendio
   mInicial = stDesabilitada
   mStatus = stDesabilitada
   mSZona = stNone
   mArquivo = ""
   mflagDupla = False
   lowbat = False: rede = False
   tampa = False: reset = False
   link = False: ruido = False
   'nivel_Sinal = 8              'sem indicação de nivel, i.e., nivel ausente
   critico = False
   crColor = 4                   'vbWhite
   ServerAddress = ""
   Camera = ""
   Monitor = ""
   SpotNumber = -1
   telaCheia = False
   user = "admin"
   senha = ""
   popup = False
End Sub

Public Property Let NivelSinal(fNewValue As String)
    On Error GoTo Signal_Error
    nivel_Sinal = Int("&H" & fNewValue) - &H1E
    If nivel_Sinal > 8 Then nivel_Sinal = 8
    Exit Property
Signal_Error:
    nivel_Sinal = 0
End Property

Public Property Get NivelSinal() As String
    NivelSinal = nivel_Sinal
End Property

Public Property Let flagDupla(fNewValue As Boolean)
   mflagDupla = fNewValue
   Dim lEntity As clsEntity
   Set lEntity = lstEntity.Item(CStr(mEntity))
   If mCheck = vrDupla And Not mflagDupla Then
      If mSZona <> stFechado Then
         lEntity.UpdateColor clearDisp:=False
         'FIXME: EventRegistry lEntity
      End If
   End If
End Property

Public Property Get status() As typeStatus
   status = mStatus
End Property

'FIXME: Mudar para String (afeta Base de Dados e Grid)
Public Property Get STampa() As Byte
   If tampa Then
      STampa = 1
   Else
      STampa = 0
   End If
End Property

'FIXME: Mudar para String (afeta Base de Dados e Grid)
Public Property Get SLowBat() As Byte
   If lowbat Then
      SLowBat = 1
   Else
      SLowBat = 0
   End If
End Property

Public Property Let status(fNewValue As typeStatus)
   Dim lEntity As clsEntity
   Set lEntity = lstEntity.Item(CStr(mEntity))
   If mInicial = stAtivada Then
      If fNewValue = stDesativada Then
         If qResponse = sxQNone Then
            Load frmQuestion
            frmQuestion.lblSensor = lEntity.vDescr & " / " & Me.mLocal
            frmQuestion.Show vbModal
         End If
         Select Case qResponse
            Case sxQAll
               'nothing to do
            Case sxQYes
               qResponse = sxQNone
            Case sxQNo
               qResponse = sxQNone
               Exit Property
            Case sxQCancel
               Exit Property
         End Select
      End If
   End If
   lEntity.ClearStatus Me
   mStatus = fNewValue
   lEntity.UpdateStatus Me
   lEntity.UpdateColor clearDisp:=False
End Property

Public Property Get SZona() As typeZona
      SZona = mSZona
End Property

Public Property Get SRuido() As Boolean
    SRuido = ruido
End Property

Public Sub AtualizaPTI(ByVal vNewValue As String)
   mPTI = vNewValue
   Dim lcm As New ADODB.Command
   Set lcm.ActiveConnection = cnDB
   lcm.CommandType = adCmdText
   lcm.CommandText = "UPDATE Sensor SET PTI = '" & mPTI & "' WHERE UID = '" & UID & "'"
   lcm.Execute
End Sub

Private Sub AtualizaPlanta(fEntity As clsEntity, ByVal vNewValue As typeZona, ByVal fPersist As Boolean)
   fEntity.ClearStatus Me
   mSZona = vNewValue
   fEntity.UpdateStatus Me
   fEntity.SetInatividade
   fEntity.UpdateColor clearDisp:=False
   If mTipo = s_Intrusao Then
      If mSZona = stFechado And Not fEntity.hasIntrusOpen(1) Then
         fEntity.Mode = 1
      ElseIf (mSZona = stAberto Or mSZona = stCurto) And fEntity.hasIntrusOpen(1) Then
         fEntity.Mode = 2
      End If
   End If
   If fPersist Then
      Dim lcm As New ADODB.Command
      Set lcm.ActiveConnection = cnDB
      lcm.CommandType = adCmdText
      lcm.CommandText = "UPDATE Sensor SET SZona = " & mSZona & " WHERE " & _
                        "UID = '" & UID & "'"
      lcm.Execute
   End If
End Sub

Private Sub RondaRegistry(fEntity As clsEntity, fRonda As clsRonda, fEvent As clsEvent)
   Dim tPercurso As clsPercurso
   Set tPercurso = lstPercurso.Item(CStr(fRonda.idPercurso))
   fEvent.evDescr = fEntity.vDescr
   fEvent.evTipo = mTipo
   Dim kind_Ronda As Integer
   If mSZona = stAberto Then
      fEvent.evStr = "Ronda no horário:" & fRonda.descrRonda & ", Percurso " & tPercurso.descrPercurso
      kind_Ronda = 0
   Else
      Set tDisplay = New clsDisplay
      If mSZona = stCurto Then
         fEvent.evStr = "Ronda fora do intervalo: " & fRonda.descrRonda & ", Percurso " & tPercurso.descrPercurso
         tDisplay.dispStr = "Ronda fora do intervalo: " & fEntity.vDescr & ", Percurso " & tPercurso.descrPercurso
         kind_Ronda = 1
      ElseIf mSZona = stFalha Then
         fEvent.evStr = "Ronda não executada: " & fRonda.descrRonda & ", Percurso " & tPercurso.descrPercurso
         tDisplay.dispStr = "Ronda não executada: " & fEntity.vDescr & ", Percurso " & tPercurso.descrPercurso
         kind_Ronda = 2
      End If
      tDisplay.dispImg = sxImgAlert
      Insert_Display tDisplay, False, False
   End If
   fEntity.EventAdd fEvent
   EventAdd fEvent
   Dim lcmd As New ADODB.Command
   Set lcmd.ActiveConnection = cnDB
   lcmd.CommandType = adCmdText
   lcmd.CommandText = "INSERT INTO EvtRonda (Date_Ronda, Hour_Ronda, Descr_Ronda, " & _
                      "fk_Entity, fk_Ronda, fk_Percurso, kind_Ronda) VALUES ('" & _
                      fEvent.evDate & "', '" & fEvent.evDate & "', '" & fEvent.evDescr & "', " & _
                      fRonda.idEntity & ", " & fRonda.idRonda & ", " & fRonda.idPercurso & _
                      ", " & kind_Ronda & ")"
   lcmd.Execute
End Sub

Private Sub EventRegistry(fEntity As clsEntity, fEvent As clsEvent)
   Dim isPanico As Boolean
   Set tDisplay = New clsDisplay
   isPanico = False
   If mSZona <> stFechado Then
      If mTipo = s_Panico Then
         isPanico = True
         tDisplay.dispMode = sxPnSound
         tDisplay.dispFile = mArquivo
      ElseIf mTipo = s_Sistema Then
         tDisplay.dispMode = sxPnSound    'Treat Escada like Panico
         tDisplay.dispFile = mArquivo
      Else
         tDisplay.dispMode = sxAlSound
         tDisplay.dispFile = mArquivo
      End If
   End If
   If mSZona = stCurto Then 'FIXME
      tDisplay.dispStr = "Zona em Curto: " & fEntity.vDescr & " - " & _
                         strTipo(mTipo) & " - " & mLocal
      tDisplay.dispImg = sxImgAlert
   ElseIf mSZona = stAberto Then
      tDisplay.dispStr = "Alarme de " & strTipo(mTipo) & " na " & _
      fEntity.vDescr & " (sensor: " & mLocal & ")"
      tDisplay.dispImg = sxImgAlert
   End If
   Insert_Display tDisplay, isPanico, critico
   'Detailed report will be made in frmQueue if Event is critico.
   'Here the event is treated as normal event
'   fEvent.evCritico = False
   fEvent.evDescr = fEntity.vDescr
   fEvent.evTipo = mTipo
   fEntity.EventAdd fEvent
   EventAdd fEvent
   
End Sub

Public Sub TreatRuido(ByVal fvalue As Boolean, fEvent As clsEvent)
   If ruido <> fvalue Then
      ruido = fvalue
      Set tDisplay = New clsDisplay
      Dim lEntity As clsEntity
      Set lEntity = lstEntity.Item(CStr(mEntity))
      If Not ruido Then
         AtualizaPlanta lEntity, stFechado, True
         tDisplay.dispStr = ""
         tDisplay.dispMode = sxNoSound
      Else
         AtualizaPlanta lEntity, stFalha, True
         tDisplay.dispMode = sxErSound
         tDisplay.dispFile = ""
         tDisplay.dispStr = fEvent.evStr & lEntity.vDescr & " - " & mLocal
         tDisplay.dispImg = sxImgAlert
      End If
      Insert_Display tDisplay, False, False
      fEvent.evDescr = lEntity.vDescr
      fEvent.evTipo = mTipo
      Set tEvent = New clsEvent
      With tEvent
          .sUIDo = fEvent.sUIDo
          .evDescr = fEvent.evDescr
          .evDate = fEvent.evDate
          .evStr = fEvent.evStr
          .evTipo = fEvent.evTipo
      End With
      lEntity.EventAdd tEvent
      EventAdd tEvent
   End If
End Sub

Public Sub TreatReset(ByVal fvalue As Boolean, fEvent As clsEvent)
   If reset <> fvalue Then
      reset = fvalue
      Dim lEntity As clsEntity
      Set lEntity = lstEntity.Item(CStr(mEntity))
      Set tDisplay = New clsDisplay
      If Not reset Then
         AtualizaPlanta lEntity, stFechado, True
         tDisplay.dispStr = ""
         tDisplay.dispMode = sxNoSound
      Else
         AtualizaPlanta lEntity, stFalha, True
         tDisplay.dispMode = sxErSound
         tDisplay.dispFile = ""
         tDisplay.dispStr = fEvent.evStr & lEntity.vDescr & " - " & mLocal
         tDisplay.dispImg = sxImgAlert
      End If
      Insert_Display tDisplay, False, False
      fEvent.evDescr = lEntity.vDescr
      fEvent.evTipo = mTipo
      Set tEvent = New clsEvent
      With tEvent
          .sUIDo = fEvent.sUIDo
          .evDescr = fEvent.evDescr
          .evDate = fEvent.evDate
          .evStr = fEvent.evStr
          .evTipo = fEvent.evTipo
      End With
      lEntity.EventAdd tEvent
      EventAdd tEvent
   End If
End Sub

Public Sub TreatLink(ByVal fvalue As Boolean, fEvent As clsEvent)
   If link <> fvalue Then
      link = fvalue
      Dim lEntity As clsEntity
      Set lEntity = lstEntity.Item(CStr(mEntity))
      Set tDisplay = New clsDisplay
      If Not link Then
         AtualizaPlanta lEntity, stFechado, True
         tDisplay.dispStr = ""
         tDisplay.dispMode = sxNoSound
      Else
         AtualizaPlanta lEntity, stFalha, True
         tDisplay.dispMode = sxErSound
         tDisplay.dispFile = ""
         tDisplay.dispStr = fEvent.evStr & lEntity.vDescr & " - " & mLocal
         tDisplay.dispImg = sxImgAlert
      End If
      Insert_Display tDisplay, False, False
      fEvent.evDescr = lEntity.vDescr
      fEvent.evTipo = mTipo
      lEntity.EventAdd fEvent
      EventAdd fEvent
   End If
End Sub

Public Sub TreatRedeAC(ByVal fvalue As Boolean, fEvent As clsEvent)
   If rede <> fvalue Then
      rede = fvalue
      Dim lEntity As clsEntity
      Set lEntity = lstEntity.Item(CStr(mEntity))
      Set tDisplay = New clsDisplay
      If Not rede Then
         AtualizaPlanta lEntity, stFechado, True
         tDisplay.dispStr = ""
         tDisplay.dispMode = sxNoSound
      Else
         AtualizaPlanta lEntity, stFalha, True
         tDisplay.dispMode = sxErSound
         tDisplay.dispFile = ""
         tDisplay.dispStr = fEvent.evStr & lEntity.vDescr & " - " & mLocal
         tDisplay.dispImg = sxImgAlert
      End If
      Insert_Display tDisplay, False, False
      fEvent.evDescr = lEntity.vDescr
      fEvent.evTipo = mTipo
      Set tEvent = New clsEvent
      With tEvent
          .sUIDo = fEvent.sUIDo
          .evDescr = fEvent.evDescr
          .evDate = fEvent.evDate
          .evStr = fEvent.evStr
          .evTipo = fEvent.evTipo
      End With
      lEntity.EventAdd tEvent
      EventAdd tEvent
   End If
End Sub

Public Sub TreatBateria(ByVal fvalue As Boolean, fEvent As clsEvent)
   If lowbat <> fvalue Then
      lowbat = fvalue
      Dim lEntity As clsEntity
      Set lEntity = lstEntity.Item(CStr(mEntity))
      Set tDisplay = New clsDisplay
      If Not lowbat Then
         AtualizaPlanta lEntity, stFechado, True
         tDisplay.dispStr = ""
         tDisplay.dispMode = sxNoSound
     Else
         AtualizaPlanta lEntity, stFalha, True
         tDisplay.dispMode = sxErSound
         tDisplay.dispFile = ""
         tDisplay.dispStr = fEvent.evStr & " (" & lEntity.vDescr & " - " & mLocal & ")"
         tDisplay.dispImg = sxImgAlert
      End If
      Insert_Display tDisplay, False, False
      fEvent.evDescr = lEntity.vDescr
      fEvent.evTipo = mTipo
      'copy the event because fEvent will be modified
      Set tEvent = New clsEvent
      With tEvent
          .sUIDo = fEvent.sUIDo
          .evDescr = fEvent.evDescr
          .evDate = fEvent.evDate
          .evStr = fEvent.evStr
          .evTipo = fEvent.evTipo
      End With
      lEntity.EventAdd tEvent
      EventAdd tEvent
   End If
End Sub

Public Sub TreatTampa(ByVal fvalue As Boolean, fEvent As clsEvent)
   If tampa <> fvalue Then
      tampa = fvalue
      Dim lEntity As clsEntity
      Set lEntity = lstEntity.Item(CStr(mEntity))
      If Not tampa Then
         lEntity.flagVandalismo = False
         AtualizaPlanta lEntity, stFechado, True
      Else
         lEntity.flagVandalismo = True
         AtualizaPlanta lEntity, stTamper, True
      End If
      
      Dim isPanico As Boolean
      isPanico = False
      Set tDisplay = New clsDisplay
      
      If mSZona = stTamper Then
         Select Case mTipo
            Case s_Panico
               isPanico = True
               tDisplay.dispMode = sxTPanico
            Case s_Intrusao
               tDisplay.dispMode = sxTIntrus
            Case s_Incendio
               tDisplay.dispMode = sxTIncend
            Case s_Emergencia
               tDisplay.dispMode = sxTEmerge
            Case s_Sistema
               tDisplay.dispMode = sxTSistem
            Case Else
               tDisplay.dispMode = sxErSound
         End Select
         tDisplay.dispFile = ""
         tDisplay.dispStr = fEvent.evStr & " - " & lEntity.vDescr & " - " & mLocal
         tDisplay.dispImg = sxImgAlert
      Else
         tDisplay.dispStr = fEvent.evStr & " - " & lEntity.vDescr & " - " & mLocal
         tDisplay.dispMode = sxEvSound
      End If
      
      Insert_Display tDisplay, isPanico, False
      fEvent.evDescr = lEntity.vDescr
      fEvent.evTipo = mTipo
      Set tEvent = New clsEvent
      With tEvent
          .sUIDo = fEvent.sUIDo
          .evDescr = fEvent.evDescr
          .evDate = fEvent.evDate
          .evStr = fEvent.evStr
          .evTipo = fEvent.evTipo
      End With
      lEntity.EventAdd tEvent
      EventAdd tEvent
   End If
End Sub

Public Sub TreatAlarme(ByVal fvalue As Byte, fEvent As clsEvent)
   If (mSZona <> fvalue) Or critico Then
      Dim lEntity As clsEntity
      Set lEntity = lstEntity.Item(CStr(mEntity))
      If mSZona = stFalha Then
         AtualizaPlanta lEntity, fvalue, True
      Else
         If mTipo = s_Intrusao Then
            mk_Report lEntity, fvalue
         End If
         If mTipo = s_Ronda Then
            If fvalue = stFechado Then
               AtualizaPlanta lEntity, stFechado, False
            Else
               Dim lronda As clsRonda
               Set lronda = GetRonda(lEntity.vId)
               If lronda Is Nothing Then Exit Sub
               lronda.ticketPassagem = True
               If fvalue = stFalha Then
                  AtualizaPlanta lEntity, stFalha, False
                  RondaRegistry lEntity, lronda, fEvent
                  DoEvents
                  AtualizaPlanta lEntity, stFechado, False
               ElseIf fvalue = stAberto Then
                  If lronda.Occur_in_Interval Then
                     AtualizaPlanta lEntity, stAberto, False
                     RondaRegistry lEntity, lronda, fEvent
                  Else
                     'Aqui, stCurto é usado para indicar que ocorreu a passagem
                     'de ronda fora do intervalo
                     AtualizaPlanta lEntity, stCurto, False
                     RondaRegistry lEntity, lronda, fEvent
                  End If
               End If
            End If
         ElseIf (mStatus = stAtivada) Or (mTipo = s_Panico) Then
            If ((fvalue = stAberto) Or (fvalue = stCurto)) Then
               If mCheck = vrDupla Then
                  If Not mflagDupla Then
                     mflagDupla = True
                     Insert_Service fEvent, check_dupla, fvalue
                     Exit Sub
                  Else
                     'desliga o serviço e o flag de dupla verificação
                     Remove_Service check_dupla
                     mflagDupla = False
                  End If
               End If
               If mLogica = lgAND Then
                  Dim lLogica As clsModule
                  Set lLogica = lstModule.Item(CStr(mNumLogica))
                  If lLogica.SZona = stFechado Then
                     Exit Sub
                  End If
               End If
               AtualizaPlanta lEntity, fvalue, True
               EventRegistry lEntity, fEvent
            Else  ' fValue = stFechado
               AtualizaPlanta lEntity, stFechado, (mSZona <> fvalue)
               If mSZona <> fvalue Then
                  If fEvent.evStr = "" Then
                    fEvent.evStr = "Sensor normalizado."
                  End If
                  EventRegistry lEntity, fEvent
               End If
            End If
         Else ' mStatus <> stAtivada
            AtualizaPlanta lEntity, fvalue, True
         End If
      End If
   ElseIf mflagDupla Then
      mService.stype = fvalue
   End If

End Sub

Public Property Get SInicial() As typeStatus
   SInicial = mInicial
End Property

Public Property Let SInicial(ByVal vNewValue As typeStatus)
   mInicial = vNewValue
   mStatus = mInicial
End Property

Public Property Let InitZStatus(ByVal vNewValue As typeZona)
   If mSZona <> vNewValue Then
      Dim lEntity As clsEntity
      Set lEntity = lstEntity.Item(CStr(mEntity))
      lEntity.ClearStatus Me
      mSZona = vNewValue
      lEntity.UpdateStatus Me
      If mCheck = vrDupla And mflagDupla = True Then
         Exit Property
      ElseIf mLogica = lgAND Then
         Dim lLogica As clsModule
         Set lLogica = lstModule.Item(CStr(mNumLogica))
         If lLogica.SZona <> mSZona Then
            Exit Property
         Else
            lEntity.UpdateColor clearDisp:=False
         End If
      Else
         lEntity.UpdateColor clearDisp:=False
      End If
      If mTipo = s_Intrusao Then
         If mSZona = stFechado And Not lEntity.hasIntrusOpen(1) Then
            lEntity.Mode = 1
         ElseIf (mSZona = stAberto) And lEntity.hasIntrusOpen(1) Then
            lEntity.Mode = 2
         End If
      End If
   End If
End Property

Private Sub Insert_Service(fEvent As clsEvent, ByVal fSkind As typeService, Optional ByVal ftype As typeZona)
   If fSkind = check_dupla Then
      Dim lService As New clsService
      lService.sModule = UID
      lService.sEvent = fEvent
      'insere controle do tempo de janela e o tipo de evento
      lService.stime = mJanela
      lService.stype = ftype
      lService.skind = fSkind
      lstService.Add lService
      Set mService = lService
      ForNet.trmService.Enabled = True
   End If
End Sub

Public Sub Remove_Service(fSkind As typeService)
   If lstService.Count > 0 Then
      Dim IdxService As Integer
      IdxService = 0
      Dim lsrv As clsService
      For Each lsrv In lstService
         IdxService = IdxService + 1
         With lsrv
            If (.sModule = UID) And (.skind = fSkind) Then
              lstService.Remove IdxService
              Set mService = Nothing
              IdxService = IdxService - 1
            End If
         End With
      Next
   End If
End Sub

Public Sub SetFalha()
   mStatAtiv = False
   Dim lEvent As New clsEvent
   'lEvent.SensorId = Me.UID
   lEvent.sUIDo = Me.UID
   lEvent.evStr = "Evento de Inatividade: "
   lEvent.evDate = Format(Now, "dd/mm/yyyy hh:mm:ss")
   Dim lEntity As clsEntity
   Set lEntity = lstEntity.Item(CStr(mEntity))
   AtualizaPlanta lEntity, stFalha, True
   Set tDisplay = New clsDisplay
   tDisplay.dispMode = sxTSistem
   tDisplay.dispFile = ""
   tDisplay.dispStr = lEvent.evStr & lEntity.vDescr & " - " & mLocal
   tDisplay.dispImg = sxImgAlert
   Insert_Display tDisplay, False, False
   lEvent.evDescr = lEntity.vDescr
   lEvent.evTipo = mTipo
   lEntity.EventAdd lEvent
   EventAdd lEvent
End Sub


Private Sub mk_Report(ByVal fEntity As clsEntity, ByVal fNewValue As typeZona)
   'Verifica se transição de fechado p/ aberto
   If (fNewValue = stAberto Or fNewValue = stCurto) And (SZona = stFechado) Then
      'Verifica se Intervalo de Abertura
      If DateDiff("n", Time, m_dTOpen(curWeekday)) > 0 Then
         'Verifica se a Entidade não se encontra aberta por outra Zona de Intrusão
         If Not fEntity.hasIntrusOpen(1) Then
            If Not fEntity.hasAccessOpen Then
               fEntity.hasAccessOpen = True
               fEntity.OpenTime = Now()               'save the first open Occur
               fEntity.OpenLast = fEntity.OpenTime    'Initially, last = first
            Else
               fEntity.OpenLast = Now()      'save the last open Occur
            End If
            fEntity.Register_OpenEvents
         End If
      End If
   'Verifica se transição de aberto p/ fechado
   ElseIf (SZona = stAberto Or SZona = stCurto) And (fNewValue = stFechado) Then
      'Verifica se Intervalo de Fechamento
      If DateDiff("n", m_dTOpen(curWeekday), Time) > 0 And _
         DateDiff("n", Time, m_dTClose(curWeekday)) > 0 Then
         'Verifica se as demais Zonas de Intrusão se encontram fechadas
         If Not fEntity.hasIntrusOpen(2) Then
            If Not fEntity.hasAccessClose Then
               fEntity.hasAccessClose = True
               fEntity.CloseTime = Now()              'Save the first Close Occur
               fEntity.CloseLast = fEntity.CloseTime  'Initially, last = first
            Else
               fEntity.CloseLast = Now()              'Save the last Close Occur
            End If
            fEntity.Register_CloseEvents
         End If
      End If
   End If
End Sub

Public Function ShowEvent_in_Camera() As Boolean
    If (mTipo = s_Panico) And (ServerAddress <> "") And (Camera <> "") And (Monitor <> "") Then
       ShowEvent_in_Camera = True
    Else
       ShowEvent_in_Camera = False
    End If
End Function
