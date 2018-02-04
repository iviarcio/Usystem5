Attribute VB_Name = "ForAux"
'
'Módulo de rotinas auxiliares do USystem
'
Option Explicit

' new to 1.0.27
Private SpotUsed(1 To 4) As Boolean
Private Const maxSpot As Integer = 4
' end new
Public TraceDump As Boolean
Public FileNumber As Integer
Public ProgressCounter As Integer
Public Const strKeyCode = "@!#$%&^*{}+?"

' new to 1.0.42
Private sizeLastEvents As Integer
' end new

' new to 1.0.31
Private Const GWL_EXSTYLE       As Long = (-20)
Private Const LWA_COLORKEY      As Long = &H1
Private Const LWA_Defaut         As Long = &H2
Private Const WS_EX_LAYERED     As Long = &H80000
' end new

Public Function Transparency(ByVal hWnd As Long, Optional ByVal Col As Long = vbBlack, _
                             Optional ByVal PcTransp As Byte = 255, Optional ByVal TrMode As Boolean = True) As Boolean
' Return : True if there is no error.
' hWnd   : hWnd of the window to make transparent
' Col : Color to make transparent if TrMode=False
' PcTransp  : 0 Ã  255 >> 0 = transparent  -:- 255 = Opaque
   Dim DisplayStyle As Long
   On Error GoTo lblExit
   DisplayStyle = GetWindowLong(hWnd, GWL_EXSTYLE)
   If DisplayStyle <> (DisplayStyle Or WS_EX_LAYERED) Then
      DisplayStyle = (DisplayStyle Or WS_EX_LAYERED)
      Call SetWindowLong(hWnd, GWL_EXSTYLE, DisplayStyle)
   End If
   Transparency = (SetLayeredWindowAttributes(hWnd, Col, PcTransp, IIf(TrMode, LWA_COLORKEY Or LWA_Defaut, LWA_COLORKEY)) <> 0)
         
lblExit:
   If Not Err.Number = 0 Then Err.Clear
   
End Function

Public Sub ActiveTransparency(M As Form, D As Boolean, f As Boolean, _
                              T_Transparency As Integer, Optional Color As Long)
Dim B As Boolean
        If D And f Then
        'Makes color (here the background color of the shape) transparent
        'upon value of T_Transparency
            B = Transparency(M.hWnd, Color, T_Transparency, False)
        ElseIf D Then
            'Makes form, including all components, transparent
            'upon value of T_Transparency
            B = Transparency(M.hWnd, 0, T_Transparency, True)
        Else
            'Restores the form opaque.
            B = Transparency(M.hWnd, , 255, True)
        End If
End Sub

Public Sub GetEntity(ByVal fFloor As Integer, x As Long, Y As Long)
   'Retorna a região que se encontra nas coordenadas X e Y.
   'Se não houver nenhuma região, retorna nothing.
   'Afeta o objeto tentity
   Dim lngRet As Long
   For Each tEntity In lstEntity
      If tEntity.floor = fFloor Then
         lngRet = PtInRegion(tEntity.Handle, x, Y)
         If lngRet <> 0 Then
            Exit Sub
         End If
      End If
   Next
   Set tEntity = Nothing
End Sub

'FIXME
Public Sub EventAdd(fEvent As clsEvent)
   If fEvent.evCritico Then
      'Event already inserted on the last events collection
   Else
      If lstEvent.Count = 0 Then
         lstEvent.Add fEvent
      Else
         lstEvent.Add fEvent, , Before:=1
         If lstEvent.Count > 100 Then
            lstEvent.Remove 100
         End If
      End If
   End If
   'Now, register in the DataBase
   Dim cM As clsModule
   Set cM = lstModule.Item(fEvent.sUIDo)
   Dim cE As clsEntity
   Set cE = lstEntity.Item(CStr(cM.mEntity))
   Dim lcmd As New ADODB.Command
   Set lcmd.ActiveConnection = cnDB
   lcmd.CommandType = adCmdText
   
   If fEvent.evCritico Then
   lcmd.CommandText = "INSERT INTO Critico (fk_Entity, fk_Sensor, Date_Event, Hour_Event, Descr_Event, crKind, crUser, crAcao, crObs, crTreat) VALUES (" & _
                      cM.mEntity & ", '" & cM.UID & "', '" & fEvent.evDate & "','" & fEvent.evDate & _
                      "', '" & fEvent.evStr & "', " & fEvent.evScope & ", '" & fEvent.evUser & "', '" & _
                      fEvent.evAcao & "', '" & fEvent.evObs & "', '" & fEvent.evTreat & "')"
   
   Else
   lcmd.CommandText = "INSERT INTO Event (fk_Entity, fk_Sensor, Date_Event, Hour_Event, Descr_Event, kind_Event) VALUES (" & _
                      cM.mEntity & ", '" & cM.UID & "', '" & fEvent.evDate & "', '" & fEvent.evDate & _
                      "','" & fEvent.evStr & "', " & fEvent.evKind & ")"
   End If
   lcmd.Execute

   
End Sub

Public Function GenerateId() As Integer
   Static gId As Integer
   gId = gId + 1
   GenerateId = gId
End Function

Public Sub Make_Service(fDescr As String, fUser As String, Optional fEntity As Variant)
   Dim lEntity As Integer
   Dim lDescr_Entity As String
   If IsMissing(fEntity) Then
      lEntity = 0
      lDescr_Entity = "-"
   Else
      lEntity = CInt(fEntity)
      lDescr_Entity = Left$(tEntity(lEntity).Descr, 70)
   End If
   Dim lcmd As New ADODB.Command
   Set lcmd.ActiveConnection = cnDB
   lcmd.CommandType = adCmdText
   lcmd.CommandText = "INSERT INTO Service (Descr_Service, Date_Service, fk_Entity, " & _
                      "[User]) VALUES ('" & fDescr & "', '" & Now & "', " & lEntity & ", '" & fUser & "')"
   lcmd.Execute
End Sub

Public Sub Save_Last_Activities(ByVal shareable As Boolean)
   Dim cM As clsModule
   Dim lcmd As New ADODB.Command
   Set lcmd.ActiveConnection = cnDB
   lcmd.CommandType = adCmdText
   lcmd.ActiveConnection.BeginTrans
   For Each cM In lstModule
      With cM
         lcmd.CommandText = "UPDATE Sensor SET Last_Ativ = '" & .mLastAtiv & _
                   "' WHERE (UID = '" & .UID & "')"
         lcmd.Execute
         If Not shareable Then
            IncProgress
         Else
            DoEvents
         End If
      End With
   Next
   lcmd.ActiveConnection.CommitTrans
End Sub

Public Sub Aux_Initialize()
' new to 1.0.27
   InitSpotNumber
' end new
   TraceDump = False
   FileNumber = 0
   
'  Busca as configurações
   Dim lds As New ADODB.Recordset
   lds.Open "SELECT * FROM Config", cnDB, adOpenStatic, adLockReadOnly
   'Get the information about Backups
   m_bBackupAuto = lds("BackupAuto")
   m_sHorario = lds("Horario")
   'Get the information about label position of pisos
   m_bPisoLeft = lds("PisoLeft")
   lds.Close
   
' new to 1.0.42
   sizeLastEvents = 0
' end new
   
End Sub

Public Sub Dump_Entity_Status(ByVal fOpen As Boolean)
   Dim lcmd As New ADODB.Command
   Set lcmd.ActiveConnection = cnDB
   lcmd.ActiveConnection.BeginTrans
   lcmd.CommandType = adCmdText

   Dim lEntity As clsEntity
   Dim lPiso As clsPiso
   Dim lreport As Long
   lreport = DateDiff("d", CDate("01/01/2000"), Date)
   If fOpen Then
      lcmd.CommandText = "DELETE FROM AccessOpen WHERE Report=" & lreport
      lcmd.Execute
   Else
      lcmd.CommandText = "DELETE FROM AccessClose WHERE Report=" & lreport
      lcmd.Execute
   End If
   lcmd.ActiveConnection.CommitTrans
   DoEvents
   lcmd.ActiveConnection.BeginTrans
   For Each lEntity In lstEntity
      With lEntity
         If .hasModules(s_Intrusao) And Not .flagInativo Then
            Set lPiso = lstPiso.Item(CStr(.floor))
            If fOpen Then
               If .hasAccessOpen Then
                  lcmd.CommandText = "INSERT INTO AccessOpen (fk_Entity, Report, Descr_Floor, has_Access, " & _
                     "Date_Open, Date_Open_Last) VALUES (" & .vId & ", " & lreport & ", '" & _
                     lPiso.rCaption & "', " & .hasAccessOpen & ", '" & .OpenTime & "', '" & .OpenLast & "')"
                  lcmd.Execute
               End If
            Else
               If .hasAccessClose Then
                   lcmd.CommandText = "INSERT INTO AccessClose (fk_Entity, Report, Descr_Floor, has_Access, " & _
                     "Date_Close, Date_Close_Last) VALUES (" & .vId & ", " & lreport & ", '" & _
                     lPiso.rCaption & "', " & .hasAccessClose & ", '" & .CloseTime & "', '" & .CloseLast & "')"
                  lcmd.Execute
                End If
            End If
         End If
      End With
   Next
   lcmd.ActiveConnection.CommitTrans
   DoEvents
End Sub

Public Sub Clear_Entity_Status(ByVal fOpen As Boolean)
   Dim lcmd As New ADODB.Command
   Set lcmd.ActiveConnection = cnDB
   lcmd.CommandType = adCmdText
   For Each tEntity In lstEntity
      With tEntity
         If fOpen Then
            .hasAccessOpen = False
         Else
            .hasAccessClose = False
         End If
      End With
   Next
   If fOpen Then
      lcmd.CommandText = "UPDATE Entity SET AccessOpen = False"
   Else
      lcmd.CommandText = "UPDATE Entity SET AccessClose = False"
   End If
   lcmd.Execute
End Sub

Public Sub Dump_Lojas(ByVal fOpen As Boolean)
   Dim lcmd As New ADODB.Command
   Set lcmd.ActiveConnection = cnDB
   lcmd.ActiveConnection.BeginTrans
   lcmd.CommandType = adCmdText
   If fOpen Then
      lcmd.CommandText = "DELETE FROM LojaAberta"
   Else
      lcmd.CommandText = "DELETE FROM LojaFechada"
   End If
   lcmd.Execute
   lcmd.ActiveConnection.CommitTrans
   DoEvents
   lcmd.ActiveConnection.BeginTrans
   Dim lEntity As clsEntity
   Dim lPiso As clsPiso
   For Each lEntity In lstEntity
      With lEntity
         Set lPiso = lstPiso.Item(CStr(.floor))
         If fOpen Then
            If .hasIntrusOpen(1) And (.Mode = 2) Then
               lcmd.CommandText = "INSERT INTO LojaAberta (fk_Entity, fk_Floor, LastDate) VALUES (" & _
                                  .vId & ", " & .floor & ", '" & .OpenLast & "')"
               lcmd.Execute
            End If
         Else
            If Not .hasIntrusOpen(1) And (.Mode = 1) Then
               lcmd.CommandText = "INSERT INTO LojaFechada (fk_Entity, fk_Floor, LastDate) VALUES (" & _
                                  .vId & ", " & .floor & ", '" & .CloseLast & "')"
               lcmd.Execute
            End If
         End If
      End With
   Next
   lcmd.ActiveConnection.CommitTrans
   DoEvents
End Sub

Public Sub Save_Other_Configs(ByVal fLeftPos As Boolean)
   Dim lcmd As New ADODB.Command
   Set lcmd.ActiveConnection = cnDB
   lcmd.CommandType = adCmdText
   lcmd.CommandText = "UPDATE Config SET PisoLeft = " & fLeftPos & _
                      ", BackupAuto = " & m_bBackupAuto & ", Horario = '" & _
                      m_sHorario & "' WHERE (cp_Config = 1)"
   lcmd.Execute
End Sub

Public Sub DBEvent_CleanUp(ByVal fInterval As Long)
   Screen.MousePointer = vbHourglass
   Dim lcmd As New ADODB.Command
   Set lcmd.ActiveConnection = cnDB
   lcmd.ActiveConnection.BeginTrans
   lcmd.CommandType = adCmdText
   Dim ldif As Long
   ldif = DateDiff("d", CDate("01/01/2000"), Date) - fInterval
   If ldif > 0 Then
      lcmd.CommandText = "DELETE FROM AccessOpen WHERE (Report < " & ldif & ")"
      lcmd.Execute
      lcmd.CommandText = "DELETE FROM AccessClose WHERE (Report < " & ldif & ")"
      lcmd.Execute
   End If
   'Note: To pass date value in "#" format to Access, you need to represent
   'this value in english format, i.e., mm/dd/yyyy.
   Dim pastDate As String
   pastDate = Format$(DateAdd("d", -fInterval, Date), "mm/dd/yyyy")
   lcmd.CommandText = "DELETE FROM Event WHERE (Event.Date_Event < #" & pastDate & "#)"
   lcmd.Execute
   lcmd.CommandText = "DELETE FROM Service WHERE (Service.Date_Service < #" & pastDate & "#)"
   lcmd.Execute
   lcmd.ActiveConnection.CommitTrans
   Screen.MousePointer = vbDefault
End Sub

Public Sub LastEvents_CleanUp(ByVal fvalue As Integer)
   Screen.MousePointer = vbHourglass
   Dim lcmd As New ADODB.Command
   Set lcmd.ActiveConnection = cnDB
   lcmd.ActiveConnection.BeginTrans
   lcmd.CommandType = adCmdText
   lcmd.CommandText = "DELETE * FROM lastEvents WHERE pk_Event NOT IN (SELECT pk_Event FROM " & _
   "(SELECT TOP " & fvalue & " pk_Event FROM LastEvents ORDER BY pk_Event DESC) Foo );"
   DoEvents
   lcmd.Execute
   DoEvents
   lcmd.ActiveConnection.CommitTrans
   Screen.MousePointer = vbDefault
End Sub


Public Function XOREncryption(fCode As String, fData As String) As String
   Dim lonDataPtr As Long
   Dim intXORValue1 As Integer
   Dim intXORValue2 As Integer
   Dim strDataOut As String
   For lonDataPtr = 1 To Len(fData)
      intXORValue1 = Asc(Mid$(fData, lonDataPtr, 1))
      intXORValue2 = Asc(Mid$(fCode, ((lonDataPtr Mod Len(fCode)) + 1), 1))
      strDataOut = strDataOut & Chr(intXORValue1 Xor intXORValue2)
   Next lonDataPtr
   XOREncryption = strDataOut
End Function

Public Sub IncProgress()
   frmSplash.ProgressBar1.Value = ProgressCounter
   ProgressCounter = ProgressCounter + 1
   If ProgressCounter >= frmSplash.ProgressBar1.Max Then
      frmSplash.ProgressBar1.Max = ProgressCounter + 50
   End If
End Sub

Public Sub DataBase_Copy(fBase As String, fResult As String)
   On Error Resume Next
   If Dir(fResult) <> "" Then Kill fResult
   On Error GoTo JetError
   FileCopy fBase, fResult
   Exit Sub
JetError:
   If Err.Number = 70 Then
      'Permition Denied, try binary copy
      Dim lSize As Long
      Dim fn As Integer
      Dim bData() As Byte
      lSize = FileLen(fBase)
      If lSize > 0 Then
         ReDim bData(lSize - 1) As Byte
         fn = FreeFile
         Open fBase For Binary Access Read As fn
         Get fn, , bData
         Close fn
      End If
      fn = FreeFile
      Open fResult For Binary Access Write As fn
      If lSize > 0 Then Put fn, , bData
      Close fn
      Erase bData
      Exit Sub
   Else
      Err.Raise Err.Number
   End If
End Sub

Public Function DataBase_Compact(fBase As String, fResult As String) As Boolean
   On Error Resume Next
   If Dir(fResult) <> "" Then Kill fResult
   
    'Need to close the database?
    cnDB.Close
    
    On Error GoTo CompactError
    Dim je As jro.JetEngine
    Set je = New jro.JetEngine
    je.CompactDatabase _
      "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & fBase & ";Jet OLEDB:Database Password=DEPFwm89", _
      "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & fResult & ";Jet OLEDB:Database Password=DEPFwm89"
   DataBase_Compact = True
   cnDB.Open
   
   Exit Function
CompactError:
   MsgBox Err.Description & " " & sxContact, sxCritical, sxProname
   DataBase_Compact = False
End Function

Public Sub SetHourGlass(f As Form)
  ForNet.MousePointer = vbHourglass
  f.MousePointer = vbHourglass
End Sub

Public Sub ResetMouse(f As Form)
  ForNet.MousePointer = vbDefault
  f.MousePointer = vbDefault
End Sub

'Função que converte string recebido em ASC para HEX
Public Function Char_to_Hex(StringIn As String, STamanho As Integer) As String

    Dim i As Integer
    Dim temp As String
    
    Char_to_Hex = ""
    For i = 1 To STamanho
        temp = ""
        temp = Hex(Asc(Mid(StringIn, i, 1)))
        If Len(temp) = 1 Then temp = "0" & temp
        Char_to_Hex = Char_to_Hex & temp
    Next i

End Function

'Função que calcula o CheckSum do String
Public Function Verifica_CheckSum(ByVal strg As String) As Boolean

    Dim i As Integer
    Dim lCS As Integer
    
    lCS = 0

   'Calcula o CheckSun
   For i = 1 To Len(strg) - 1
      lCS = lCS + Asc(Mid$(strg, i, 1))
      If lCS > 255 Then lCS = lCS - 256
   Next
   
   'Compara o CheckSun
   If Asc(Right$(strg, 1)) = lCS Then
      Verifica_CheckSum = True
   Else
      Verifica_CheckSum = False
   End If
   
End Function

'Função que formata o string para apresentar na tela de comunicação
Public Function Formata_Mensagem(StringIn As String, ChkSum As Boolean) As String

    Dim Header As String
    Dim Size As String
    Dim sData As String
    Dim sMID As String
    Dim Result As String
    Dim UIDo As String
    Dim Dec_UIDo As String
    Dim UIDh As String
    Dim Dec_UIDh As String
    Dim PTI As String
    Dim Stat1 As String
    Dim Stat0 As String
    Dim level As String
    Dim margin As String
    Dim CkSum As String
    Dim DeviceName As String
    Dim Classe As String
    
    If ChkSum Then
        Result = "  OK"
    Else
        Result = "  ERRO"
    End If
   
    'Formata a mensagem de saída
    Header = Left(StringIn, 2)
    Size = Mid(StringIn, 3, 2)
    sMID = Mid(StringIn, 5, 2)

    'Trata o Serial Receiver
    If Header = H_Serial Then
        sData = Mid(StringIn, 5, 2)
        Stat0 = Mid(StringIn, 7, 2)
        CkSum = Mid(StringIn, 9, 2)
        Formata_Mensagem = "Serial Receiver ->  " & "HEADER = " & H_Serial & "  DATA = " & sData & "  STAT0 = " & Stat0 & "  CKSUM = " & CkSum & Result
    
    'Trata os Repeaters
    ElseIf Header = H_Device And sMID = MID_Repeater Then
        UIDo = Mid(StringIn, 5, 8)
        Dec_UIDo = CStr(Int("&H" + Right(UIDo, 6)))
        UIDh = Mid(StringIn, 13, 8)
        Dec_UIDh = CStr(Int("&H" + Right(UIDh, 6)))
        Classe = Mid(StringIn, 21, 2)
        Stat1 = Mid(StringIn, 23, 2)
        Stat0 = Mid(StringIn, 25, 2)
        level = Mid(StringIn, 27, 2)
        margin = Mid(StringIn, 29, 2)
        CkSum = Mid(StringIn, 31, 2)
        
        Formata_Mensagem = "High-Power Repeater -> " & "HEADER = " & H_Device & "  Size = " & Size & "  UIDo = " & UIDo & " / " & Dec_UIDo & _
                                       "  UIDh = " & UIDh & " / " & Dec_UIDh & _
                                       "  CLASSE = " & Classe & "  STAT1 = " & Stat1 & "  STAT0 = " & Stat0 & "  Level = " & level & _
                                       "  Margin = " & margin & "  CKSUM = " & CkSum & Result

    'Trata os Sensores
    ElseIf Header = H_Device And sMID = MID_Sensor Then
        UIDo = Mid(StringIn, 5, 8)
        Dec_UIDo = CStr(Int("&H" + Right(UIDo, 6)))
        UIDh = Mid(StringIn, 13, 8)
        Dec_UIDh = CStr(Int("&H" + Right(UIDh, 6)))
        PTI = Mid(StringIn, 21, 2)
        Stat1 = Mid(StringIn, 23, 2)
        Stat0 = Mid(StringIn, 25, 2)
        level = Mid(StringIn, 27, 2)
        margin = Mid(StringIn, 29, 2)
        CkSum = Mid(StringIn, 31, 2)
        
        'Busca o nome do Device na base
        DeviceName = Device_Name(PTI)
        
        Formata_Mensagem = "Sensor " & DeviceName & " -> " & "HEADER = " & H_Device & "  Size = " & Size & "  UIDo = " & UIDo & " / " & Dec_UIDo & _
                                       "  UIDh = " & UIDh & " / " & Dec_UIDh & _
                                       "  PTI = " & PTI & "  STAT1 = " & Stat1 & "  STAT0 = " & Stat0 & "  Level = " & level & "  Margin = " & margin & _
                                       "  CKSUM = " & CkSum & Result
    End If

End Function

'Rotina que busca o nome do Disposivito (Sensor) que enviou o dado
Public Function Device_Name(sPTI As String) As String
    Dim tPTI As clsPTI
    Set tPTI = Nothing
    On Error Resume Next
    Set tPTI = lstPTI.Item(sPTI)
    If tPTI Is Nothing Then
        Device_Name = ""
    Else
        Device_Name = tPTI.sProduct
    End If
    On Error GoTo 0
End Function

'Rotina que controla as diretivas de segurança e as condições das chaves
Public Sub Security_Check()
   
   'Verifica se existe registo de chave
   If gstChecksum = "" And gstCondorID = "" Then
        Key_check = 0
        Exit Sub
   End If
       
   'Existe uma chave registrada. Entra para a verificação
   If Chave Is Nothing Then Set Chave = New clsproteq
   Dim TestaChave As String
   Dim TestaData As String
   
   'Verifica se a chave esta registrada e conectada ou não
   TestaChave = Chave.Verifica(codUser)
   
    If TestaChave = "0" Then
      
      'A Chave esta registrada, porem não pode ser encontrada
      Key_check = 0
    
    Else
      
      'Chave está presente. No entanto é necessário verificar se o código
      'da chave corresponde ao código de registro.
      
      If TestaChave = gstCondorID Then
      
         'Valida se o número da chave confere com o código de registro Key_check = 1, caso contrário = 0
         Key_check = Match(gstChecksum, gstCondorID)
         
      Else
      
        'Codigo da chave de segurança não confere
         Key_check = 0
         
      End If
      
   End If
   
End Sub

Public Function Match(ByVal lCodigo As String, ByVal lKeyCode As String) As Byte
   
   Dim ndat As Long
   Dim nval As Long
   Dim i As Integer
   Dim lcode As String
   Dim lKey As String
   Dim ldig As String
      
   'Constante definida por ser a versão 01 (simples definição)
   ndat = 10
   
   For i = 1 To 7
     'Busca cada dígito do código da chave
      nval = Val(Mid(lKeyCode, i, 1))
      
      If nval = 0 Then
          lcode = ndat
          'Se o valor do dígito do código for 0 pega-se a constante ndat
          'o calculo atual mais a posição do dígito como multiplicador
          ndat = ndat * (Len(lcode) + i)
      Else
          'Se o valor do dígito do código for <> 0, multiplica com o resultado anterior
          'e acrescenta-se a posição do digito ao produto
          ndat = (ndat * nval) + i
      End If
      
   Next i
   
   'Separa o 5 digitos da direita
   lKey = Right(ndat, 5)
         
   'Atribui a este dígito mais significativo uma letra da lista abaixo
    lcode = ""
    For i = 1 To 5
        ldig = Val(Mid(lKey, i, 1))
        Select Case ldig
             Case 0
                 ldig = "A"
             Case 1
                 ldig = "B"
             Case 2
                 ldig = "C"
             Case 3
                 ldig = "D"
             Case 4
                 ldig = "E"
             Case 5
                 ldig = "F"
             Case 6
                 ldig = "G"
             Case 7
                 ldig = "H"
             Case 8
                 ldig = "I"
             Case 9
                 ldig = "J"
             Case ""
                 ldig = "X"
        End Select
        lcode = lcode & ldig
    Next i
       
   'Compara o código calculado com o registrado
   If lcode = lCodigo Then
        Match = 1
   Else
        Match = 0
   End If
    
End Function

' new to 1.0.27
Private Sub InitSpotNumber()
   Dim i As Integer
   For i = 1 To maxSpot
      SpotUsed(i) = False
   Next
End Sub

Public Function NextSpotNumber() As Integer
   Dim i As Integer
   For i = 1 To maxSpot
      If Not SpotUsed(i) Then
         SpotUsed(i) = True
         NextSpotNumber = i
         Exit Function
      End If
   Next
   NextSpotNumber = -1
End Function

Public Function FreeSpotNumber(ByVal Value As Integer) As Integer
   If Value <> -1 Then
      SpotUsed(Value) = False
   End If
   FreeSpotNumber = -1
End Function

Public Sub Control_LastEvents()
    ' keep only two thousand events
    Static ntimes As Integer ' init = 0
    sizeLastEvents = sizeLastEvents + 1
    If sizeLastEvents >= 2100 Then
        sizeLastEvents = 2000
        LastEvents_CleanUp 2000
        ntimes = ntimes + 1
    End If
    ' compact DB after two thousand deleted events
    If ntimes >= 20 Then
        ntimes = 0
        CompactDB
    End If
End Sub

Public Sub CompactDB()
    Dim lCadastro As String
    Dim cptfile As String
    cnDB.Close
    Set cnDB = Nothing
    DoEvents
    On Error Resume Next
    lCadastro = m_sPath & "\" & USystemDB
    cptfile = m_sPath & "\CptDB5.mdb"
    Kill cptfile
    DoEvents
    On Error GoTo CompactError
    Dim je As jro.JetEngine
    Set je = New jro.JetEngine
    je.CompactDatabase _
      "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & m_sPath & "\" & m_sDatabase & ";Jet OLEDB:Database Password=DEPFwm89", _
      "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & cptfile & ";Jet OLEDB:Database Password=DEPFwm89"
    DoEvents
    On Error Resume Next
    Kill m_sTmpFileDB
    Name lCadastro As m_sTmpFileDB
    FileCopy cptfile, lCadastro
    SetAttr lCadastro, vbNormal
    DoEvents
    Set cnDB = New ADODB.Connection
    cnDB.ConnectionString = m_sDatabase
    cnDB.ConnectionTimeout = 45
    cnDB.Open
    On Error GoTo 0
    Exit Sub
   
CompactError:
   MsgBox Err.Description & " " & sxContact, sxCritical, sxProname
   cnDB.Open
   
End Sub


Public Sub ShowCamera(ByVal crModule As clsModule, ByVal fLoja As String)
   Dim WinHttpReq As WinHttp.WinHttpRequest
   Dim strURL As String
   On Error GoTo camError
   With crModule
      If Not .popup Then Exit Sub
      If .ServerAddress <> "" And .Camera <> "" And .Monitor <> "" Then
         If .telaCheia Then .SpotNumber = -1 Else .SpotNumber = NextSpotNumber()
         If .senha = "" Then
            strURL = "http://" & .ServerAddress & "/Interface/VirtualMatrix/ShowObject?" & _
                  "MonitorID=" & .Monitor & "&SpotNumber=" & .SpotNumber & _
                  "&ObjectType=0&ObjectName=" & .Camera & "&ResponseFormat=Text&AuthUser=" & .user
         Else
            strURL = "http://" & .ServerAddress & "/Interface/VirtualMatrix/ShowObject?" & _
                  "MonitorID=" & .Monitor & "&SpotNumber=" & .SpotNumber & _
                  "&ObjectType=0&ObjectName=" & .Camera & _
                  "&ResponseFormat=Text&AuthUser=" & .user & "&AuthPass=" & .senha
         End If
      Else
         Exit Sub
      End If
   End With
   Set WinHttpReq = New WinHttpRequest
   WinHttpReq.SetTimeouts "2000", "5000", "2000", "20000"  ' Resolve, Connect, Send and Receive
   WinHttpReq.Open "POST", strURL, False
   WinHttpReq.SetRequestHeader "ContentType", "text/plain; encoding='utf-8'"
   WinHttpReq.SetRequestHeader "Content-Length", Len(strURL)
   WinHttpReq.Send ""
   WinHttpReq.WaitForResponse (60)
   DoEvents
   If WinHttpReq.status <> 200 Then
       ForNet.StatusBar1.Panels.Item(2).Text = "Falha na visualização da Câmera " & crModule.Camera & " para " & fLoja & " (" & crModule.mLocal & ")"
   End If
   Set WinHttpReq = Nothing
   Exit Sub
   
camError:
    If Err.Number <> 0 Then
        Err.Clear
        ForNet.StatusBar1.Panels.Item(2).Text = "Falha na visualização da Câmera em " & fLoja & ". Erro: " & Err.Description
        Resume Next
    End If

End Sub



