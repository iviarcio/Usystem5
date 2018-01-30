VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEntity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public vId As Long                  'Identificação da Entidade
Public vDescr As String             'Descrição da Entidade
Public vResp As String              'Nome do responsável
Public vTel1 As String              'Telefone Interno
Public vTel2 As String              'Telefone Externo
' new to 1.0.27
Public message As String            'Campo de observação, utilizado p/ indicar Câmera
' end new
Private vFloor As Integer           'Piso/Andar que a entidade está localizada
Private vStatus As Boolean          'Status da Entidade (Monitoração On/Off)
Private lngHandle As Long           'Handle para a região associada à entidade

' new to 1.0.34
Public colBack As BackColor        'Cor de fundo associada à região
' end new

Private lastColor As BackColor      'Última cor de fundo

'Número de Vertices da Região e respectivas cordenadas
Private intVertices As Integer
Private coordinates(1 To 24) As POINTAPI

'Lista de Dispositivos cadastrados na entidade
Public localModule As New Collection


'Lista de eventos na Entidade
Public localEvent As New Collection

'Indicação de Vandalismo em Zona da Entidade
Public flagVandalismo As Boolean

'Indicação de Inatividade
Public flagInativo As Boolean

'Indicação de Abertura e Fechamento
Public hasAccessOpen As Boolean
Public hasAccessClose As Boolean
Public OpenTime As Date
Public OpenLast As Date
Public CloseTime As Date
Public CloseLast As Date

'Indicação de Loja Aberta ou fechada
Public Mode As Byte                          '0: sem inf., 1: fechado, 2: aberto

Private intZonas As Integer                  'Número de Zonas
Private intDesabilitada As Integer           'Número de Zonas desabilitadas
Private intZAtiva As Integer                 'Número de Zonas Ativadas
Private intDesativAberta As Integer          'Número de Zonas Desativadas porem Abertas
Private intAtivAberta As Integer             'Número de Zonas Ativadas e Abertas
Private intZTipo(0 To nTSensor) As Integer   'Número de Zonas por tipo
Private intIntrusOpen As Integer             'Número de Zonas de Intrusão abertas

Private Sub Class_Initialize()
   vStatus = True
   Mode = 0
   flagInativo = False
   flagVandalismo = False
   vFloor = vbEmpty
   colBack = colGray
   lastColor = colGray
   intVertices = 0
   OpenTime = Now
   OpenLast = OpenTime
   CloseTime = Now
   CloseLast = CloseTime
   intZonas = 0
   intDesabilitada = 0
   intZAtiva = 0
   intDesativAberta = 0
   intAtivAberta = 0
   intZTipo(s_Incendio) = 0
   intZTipo(s_Intrusao) = 0
   intZTipo(s_Emergencia) = 0
   intZTipo(s_Panico) = 0
   intZTipo(s_Sistema) = 0
   intIntrusOpen = 0
End Sub

Public Property Get floor() As Integer
   floor = vFloor
End Property

Public Property Let floor(ByVal vNewValue As Integer)
   vFloor = vNewValue
End Property

Public Property Get status() As Boolean
   status = vStatus
End Property

Public Property Let status(ByVal vNewValue As Boolean)
   vStatus = vNewValue
   UpdateColor clearDisp:=False
End Property

Public Property Get Handle() As Long
   Handle = lngHandle
End Property

Public Property Let Handle(ByVal vNewValue As Long)
   lngHandle = vNewValue
End Property

Public Property Get BackGround() As BackColor
   BackGround = colBack
End Property

Public Property Get Vertices() As Integer
   Vertices = intVertices
End Property

Public Property Let Vertices(ByVal vNewValue As Integer)
   intVertices = vNewValue
End Property

Public Sub Dump(fEvento As typeEvento)
   Select Case fEvento
      Case stAdded
         Dim lci As New ADODB.Command
         Set lci.ActiveConnection = cnDB
         lci.CommandType = adCmdText
         lci.CommandText = "INSERT INTO Entity (fk_Floor, Descr_Entity, Resp_Entity, Tel1_Entity, " & _
                           "Tel2_Entity, message, nVertices, X01, Y01, X02, Y02, X03, Y03, X04, Y04, X05, Y05, " & _
                           "X06, Y06, X07, Y07, X08, Y08, X09, Y09, X10, Y10, X11, Y11, X12, Y12, " & _
                           "X13, Y13, X14, Y14, X15, Y15, X16, Y16, X17, Y17, X18, Y18, X19, Y19, " & _
                           "X20, Y20, X21, Y21, X22, Y22, X23, Y23, X24, Y24) Values (" & vFloor & ", '" & vDescr & "', '" & _
                           vResp & "', '" & vTel1 & "', '" & vTel2 & "', '" & message & "', 0, " & coordinates(1).x & ", " & coordinates(1).Y & ", " & _
                           coordinates(2).x & ", " & coordinates(2).Y & ", " & coordinates(3).x & ", " & coordinates(3).Y & ", " & _
                           coordinates(4).x & ", " & coordinates(4).Y & ", " & coordinates(5).x & ", " & coordinates(5).Y & ", " & _
                           coordinates(6).x & ", " & coordinates(6).Y & ", " & coordinates(7).x & ", " & coordinates(7).Y & ", " & _
                           coordinates(8).x & ", " & coordinates(8).Y & ", " & coordinates(9).x & ", " & coordinates(9).Y & ", " & _
                           coordinates(10).x & ", " & coordinates(10).Y & ", " & coordinates(11).x & ", " & coordinates(11).Y & ", " & _
                           coordinates(12).x & ", " & coordinates(12).Y & ", " & coordinates(13).x & ", " & coordinates(13).Y & ", " & _
                           coordinates(14).x & ", " & coordinates(14).Y & ", " & coordinates(15).x & ", " & coordinates(15).Y & ", " & _
                           coordinates(16).x & ", " & coordinates(16).Y & ", " & coordinates(17).x & ", " & coordinates(17).Y & ", " & _
                           coordinates(18).x & ", " & coordinates(18).Y & ", " & coordinates(19).x & ", " & coordinates(19).Y & ", " & _
                           coordinates(20).x & ", " & coordinates(20).Y & ", " & coordinates(21).x & ", " & coordinates(21).Y & ", " & _
                           coordinates(22).x & ", " & coordinates(22).Y & ", " & coordinates(23).x & ", " & coordinates(23).Y & ", " & _
                           coordinates(24).x & ", " & coordinates(24).Y & ")"
         lci.Execute
         
         Dim rsEntity As ADODB.Recordset
         Set rsEntity = New ADODB.Recordset
         rsEntity.CursorLocation = adUseClient
         rsEntity.CursorType = adOpenStatic
         rsEntity.LockType = adLockReadOnly
         rsEntity.Open "SELECT * FROM Entity WHERE (nVertices = 0)", cnDB
         If Not rsEntity.EOF Then
            vId = rsEntity("cp_Entity")
            rsEntity.Close
            lci.CommandText = "UPDATE Entity SET nVertices = " & intVertices - 1 & " WHERE (nVertices = 0)"
            lci.Execute
         Else
            MsgBox "Dados incorretos.", sxExclamation, sxProname
         End If
         Set lci = Nothing
      Case stModified
         Dim lcm As New ADODB.Command
         Set lcm.ActiveConnection = cnDB
         lcm.CommandType = adCmdText
         lcm.CommandText = "UPDATE Entity SET Entity.Descr_Entity = '" & vDescr & _
                            "', Entity.Resp_Entity = '" & vResp & "', Entity.Tel1_Entity = '" & _
                            vTel1 & "', Entity.Tel2_Entity = '" & vTel2 & _
                            "', Entity.message = '" & message & _
                            "' WHERE (Entity.cp_Entity =" & vId & ")"
         lcm.Execute
      Case stRemoved
         Dim lcr As New ADODB.Command
         Set lcr.ActiveConnection = cnDB
         lcr.CommandType = adCmdText
         lcr.CommandText = "DELETE FROM  Entity WHERE (cp_Entity =" & vId & ")"
         lcr.Execute
  End Select
End Sub

Public Function Top() As Long
   Top = coordinates(1).Y
End Function

Public Function Left() As Long
   Left = coordinates(1).x
End Function

Public Sub Set_Coordinates(ByVal fIndex As Integer, fX As Long, fY As Long)
   coordinates(fIndex).x = fX
   coordinates(fIndex).Y = fY
End Sub

Public Function Get_Coordinates(ByVal fIndex As Integer, fX As Long, fY As Long) As Boolean
   If fIndex <= intVertices Then
      Get_Coordinates = True
      fX = coordinates(fIndex).x
      fY = coordinates(fIndex).Y
   Else
      Get_Coordinates = False
   End If
End Function

Public Function hasZonas() As Boolean
   hasZonas = (intZonas > intDesabilitada)
End Function

Public Function hasZonasDesativadas() As Boolean
   hasZonasDesativadas = (intZonas - intDesabilitada) > intZAtiva
End Function

Public Function hasZonasAtivadas() As Boolean
   hasZonasAtivadas = (intZAtiva > 0)
End Function

Public Function hasModules(fTipo As typeSensor) As Boolean
   hasModules = (intZTipo(fTipo) > 0)
End Function

Public Function hasIntrusOpen(ByVal n As Integer) As Boolean
   hasIntrusOpen = (intIntrusOpen >= n)
End Function

Public Sub Add(fModule As clsModule, fIndex As String)
   localModule.Add Item:=fModule, Key:=fIndex
   UpdateStatus fModule
   UpdateColor clearDisp:=False
End Sub

'Public Sub UpdateBase(fBase As clsBase)
'   If localBase.Count > 0 Then
'      Dim lBase As clsBase
'      For Each lBase In localBase
'         If lBase.SysId = fBase.SysId Then
'            'Base already in the list
'            Exit Sub
'         End If
'      Next
'   End If
'   localBase.Add Item:=fBase, Key:=CStr(fBase.SysId)
'End Sub

Public Sub Remove(fModule As clsModule)
   If localModule.Count > 0 Then
      Dim Index As Integer
      Index = 1
      Dim lModule As clsModule
      For Each lModule In localModule
         If lModule.UID = fModule.UID Then
            ClearStatus fModule
            localModule.Remove Index
            UpdateColor clearDisp:=False
            Exit Sub
         End If
         Index = Index + 1
      Next
   End If
End Sub

Public Sub UpdateColor(ByVal clearDisp As Boolean)
   lastColor = colBack
   If vStatus Then
      If intDesabilitada = intZonas Then
         colBack = colGray
      ElseIf flagVandalismo Then
         colBack = colRed
         If vFloor <> m_iCurPiso Then
            ForNet.lblPiso(vFloor).ForeColor = lblColor.cRed
         End If
      ElseIf intAtivAberta > 0 Then
         colBack = colRed
         If vFloor <> m_iCurPiso Then
            ForNet.lblPiso(vFloor).ForeColor = lblColor.cRed
         End If
      ElseIf intDesativAberta > 0 Then
         If Not clearDisp And (lastColor = colRed) Then
            colBack = colRedYellow
         Else
            colBack = colYellow
         End If
      ElseIf intZAtiva > 0 Then
         If Not clearDisp And (lastColor = colRed) Then
            colBack = colRedGreen
            If vFloor <> m_iCurPiso Then
               ForNet.lblPiso(vFloor).ForeColor = lblColor.cRed
            End If
         ElseIf flagInativo Then
            colBack = colRed
         Else
            colBack = colGreen
         End If
      Else
         colBack = colBlue
      End If
   Else
      colBack = colGray
   End If
End Sub

Public Sub ClearStatus(fModule As clsModule)
   With fModule
      intZonas = intZonas - 1
      If .status = stDesabilitada Then
         intDesabilitada = intDesabilitada - 1
      Else
         intZTipo(.mTipo) = intZTipo(.mTipo) - 1
         If .status = stAtivada Then
            intZAtiva = intZAtiva - 1
            If (.SZona = stAberto) Or (.SZona = stCurto) Or (.SZona = stTamper) Or (.SZona = stFalha) Then
               intAtivAberta = intAtivAberta - 1
            End If
         ElseIf (.SZona = stAberto) Or (.SZona = stCurto) Or (.SZona = stTamper) Or (.SZona = stFalha) Then
            intDesativAberta = intDesativAberta - 1
         End If
         If .mTipo = s_Intrusao Then
            If (.SZona = stAberto) Or (.SZona = stCurto) Then
               intIntrusOpen = intIntrusOpen - 1
            End If
         End If
      End If
   End With
End Sub

Public Sub UpdateStatus(fModule As clsModule)
   With fModule
      intZonas = intZonas + 1
      If .status = stDesabilitada Then
         intDesabilitada = intDesabilitada + 1
      Else
         intZTipo(.mTipo) = intZTipo(.mTipo) + 1
         If .status = stAtivada Then
            intZAtiva = intZAtiva + 1
            If (.SZona = stAberto) Or (.SZona = stTamper) Or (.SZona = stFalha) Then
               intAtivAberta = intAtivAberta + 1
            End If
         ElseIf (.SZona = stAberto) Or (.SZona = stTamper) Or (.SZona = stFalha) Then
            intDesativAberta = intDesativAberta + 1
         End If
         If .mTipo = s_Intrusao Then
            If (.SZona = stAberto) Then
               intIntrusOpen = intIntrusOpen + 1
            End If
         End If
      End If
   End With
End Sub

Public Sub SetInatividade()
'Se todas as zonas da entidade estiverem em falha e com flag de atividade desligado então
'a entidade é considerada em Inatividade.
   Dim cM As clsModule
   flagInativo = True
   For Each cM In localModule
      If cM.SZona <> stFalha Then
         flagInativo = False
         Exit Sub
      End If
   Next
   For Each cM In localModule
      If cM.mStatAtiv Then
         flagInativo = False
         Exit Sub
      End If
   Next
End Sub


Public Sub Activate(fwichModule As typeSensor)
   Dim lModule As clsModule
   For Each lModule In localModule
      If (fwichModule = s_All) Or (lModule.mTipo = fwichModule) Then
         'If fwichModule = s_Atuador Then
            'Para atuadores, indicar sempre para ativar, independente do status
            'lModule.status = stAtivada
         'Else
         If (lModule.status = stDesativada) Then  'And (lModule.mTipo <> s_Atuador) Then
            lModule.status = stAtivada
         End If
      End If
   Next
End Sub

Public Sub Deactivate(fwichModule As typeSensor)
   Dim lModule As clsModule
   For Each lModule In localModule
      If (fwichModule = s_All) Or (lModule.mTipo = fwichModule) Then
         'If fwichModule = s_Atuador Then
            'Para atuadores, indicar sempre para desativar, independente do status
            'lModule.status = stDesativada
         'Else
         If (lModule.status = stAtivada) Then  'And (lModule.mTipo <> s_Atuador) Then
            lModule.status = stDesativada
         End If
      End If
   Next
End Sub

'FIXME
Public Sub ChangeSZona(fStatus As typeStatus)
'
' This routine is used only for test purposes
'
   Dim lModule As clsModule
   Dim lEvent As New clsEvent
   For Each lModule In localModule
      If lModule.status = fStatus Then
         If lModule.mTipo = s_Panico Or lModule.mTipo = s_Sistema Then
            lEvent.evDate = Now
            If lModule.SZona = stFechado Then
               lModule.TreatAlarme stAberto, lEvent
            Else
               lModule.TreatAlarme stFechado, lEvent
            End If
            DoEvents
         End If
      End If
   Next
End Sub

Public Sub EventAdd(fEvent As clsEvent)
   If localEvent.Count = 0 Then
      localEvent.Add fEvent
   Else
      localEvent.Add fEvent, , Before:=1
      If localEvent.Count > 20 Then
         localEvent.Remove 20
      End If
   End If
End Sub

Public Sub Register_OpenEvents()
   Dim lcmd As New ADODB.Command
   Set lcmd.ActiveConnection = cnDB
   lcmd.CommandType = adCmdText
   lcmd.CommandText = "UPDATE Entity SET AccessOpen = " & hasAccessOpen & _
                      ", OpenTime ='" & OpenTime & "', OpenLast = '" & OpenLast & _
                      "'  WHERE (cp_Entity = " & vId & ")"
   lcmd.Execute
End Sub

Public Sub Register_CloseEvents()
   Dim lcmd As New ADODB.Command
   Set lcmd.ActiveConnection = cnDB
   lcmd.CommandType = adCmdText
   lcmd.CommandText = "UPDATE Entity SET AccessClose = " & hasAccessClose & _
                      ", CloseTime ='" & CloseTime & "', CloseLast = '" & CloseLast & _
                      "'  WHERE (cp_Entity = " & vId & ")"
   lcmd.Execute
End Sub

