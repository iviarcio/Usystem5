VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum StatusType
   bOKS = 0          'Base Ok, sem periférico, sem dado
   bOKW = 1          'Base Ok, com WIRED presente, sem dado
   bOKF = 2          'Base Ok, com FA403 presente, sem dado (Wireless)
   bOKD = 3          'Base Ok, com dado
   bNOK = 4          'Base not Ok
End Enum

Public Enum BaseType
   wired = 0
   wireless = 1
End Enum

Public SysId As Long                'Identificação da Base
Public receptorID As Byte           'Identificação do receptor (p/ FA403: Wireless)
'Public flag_Online As Boolean       'Indicação de Base OnLine
'Public flag_Ativ As Boolean         'Indicação de Atividade
'Public flag_Tabela As Boolean       'Indicação de Tabela
'Public flag_NoComm As Boolean       'Indicação de ausencia de Comunicaçao
'Public flag_Monitor As Boolean      'Indicação de Base em Monitoração
Public flag_Status As StatusType    'Indicação do último Status da Base
Public nivel_Sinal As Integer       'Nivel de Sinal
Public categ_base As BaseType       'Tipo da Base (se wired ou wireless)
Public flag_dupla As Boolean        'Controla o pedido de tabela em dupla verificação
'Public IsTable As Boolean           'Indicação de recebimento de tabela
'Private iCount As Integer           'Contador de On-lines
'Private iNextC As Integer           'Próximo Contador de On-Lines
'Private iCRC As Integer             'Contador de Falhas de CRC
'Public iFalha As Integer            'Contador de Falhas de Inatividade
'Public VersaoRede As String         'Versão do nó da Rede MSB.LSB
'Public VersaoPeriferico As String   'Versão do Periférico MSB.LSB

'Lista de Módulos com o mesmo SysID
Public localModule As New Collection

'Public Sub Add(fModule As clsModule)
'   localModule.Add Item:=fModule, Key:=CStr(fModule.mNumero)
'End Sub

'Public Sub Remove(fModule As clsModule)
'   Dim Index As Integer
'   Index = 1
'   Dim lModule As clsModule
'   For Each lModule In localModule
'      If lModule.Idm = fModule.Idm Then
'         localModule.Remove Index
'      End If
'      Index = Index + 1
'   Next
'   If localModule.Count = 0 Then
'      Index = 1
'      Dim lBase As clsBase
'      For Each lBase In lstBase
'         If lBase.SysId = SysId Then
'            'Remove a base das Entidades
'            BaseCleanUp lBase
'            'Remove a base
'            lstBase.Remove Index
'            m_iBase = m_iBase - 1
'            Exit Sub
'         End If
'         Index = Index + 1
'      Next
'      If lstBase.Count = 0 Then
'         'stop the varredura
'         m_bCommStatus = False
'      End If
'   End If
'End Sub

'Public Function IncLimit() As Boolean
'   If iLimit <> 0 Then
'      iCount = iCount + 1
'      If iCount >= iLimit Then
'         iCount = iNextC
'         iNextC = 0
'      End If
'      IncLimit = (iCount = iNextC)
'   Else
'      IncLimit = False
'   End If
'End Function
   
'Private Sub BaseCleanUp(fBase As clsBase)
'   For Each tEntity In lstEntity
'      tEntity.Clean fBase
'   Next
'End Sub

Private Sub Class_Initialize()
'   flag_Online = False
'   flag_Tabela = False
'   flag_NoComm = True
'   flag_Monitor = False
   flag_dupla = False
'   IsTable = False
   receptorID = 0
   flag_Status = bOKS
'   iCount = iLimit
'   iNextC = -m_iBase + 1
'   iCRC = 0
'   iFalha = 0
End Sub

'Public Sub InCRC_Error()
'   iCRC = iCRC + 1
'   If iCRC > iLimit Then
'      iCRC = 0
''      flag_Online = False
'      flag_Status = bNOK
'   End If
'End Sub

'Public Sub Base_Chg(ByVal fNew_SysId As Byte)
'   'First, make changes in the Database
'   Dim lcmd As New ADODB.Command
'   Set lcmd.ActiveConnection = cnCD
'   lcmd.CommandType = adCmdText
'   lcmd.CommandText = "UPDATE Sensor SET Base_Sensor = " & fNew_SysId & _
'                      " WHERE (Base_Sensor = " & SysId & ")"
'   lcmd.Execute
'   lcmd.CommandText = "UPDATE Sensor SET Base_Logica = " & fNew_SysId & _
'                      " WHERE (Base_Logica = " & SysId & ")"
'   lcmd.Execute
'   lcmd.CommandText = "UPDATE Sensor SET Base_PGM = " & fNew_SysId & _
'                      " WHERE (Base_PGM = " & SysId & ")"
'   lcmd.Execute
'   lcmd.CommandText = "UPDATE Base SET SysID = " & fNew_SysId & _
'                      ", Rx = " & fNew_SysId & " WHERE (SysID = " & SysId & ")"
'   lcmd.Execute
'   Set lcmd.ActiveConnection = Nothing
'   'Then, make changes in the module Collection
'   Dim lModule As clsModule
'   For Each lModule In lstModule
'      With lModule
'         If .mBase = SysId Then
'            .mBase = fNew_SysId
'         End If
'         If .mBaseLogica = SysId Then
'            .mBaseLogica = fNew_SysId
'         End If
'         If .mBasePGM = SysId Then
'            .mBasePGM = fNew_SysId
'         End If
'      End With
'   Next
'   'Last, change the base identification
'   SysId = fNew_SysId
'End Sub

'Public Sub ToLimit()
'   'Force Base to ask table
'   iCount = iLimit
'End Sub

'Public Sub UpdateVersion(ByVal rMSB As Byte, ByVal rLSB As Byte, _
'                         ByVal pMSB As Byte, ByVal pLSB As Byte)
'   Dim lcm As New ADODB.Command
'   Set lcm.ActiveConnection = cnCD
'   lcm.CommandType = adCmdText
'   If CStr(rMSB) & "." & CStr(rLSB) <> VersaoRede Then
'      VersaoRede = CStr(rMSB) & "." & CStr(rLSB)
'      lcm.CommandText = "UPDATE Base SET VRede = '" & VersaoRede & "' WHERE " & _
'                        "SysID = " & SysId
'      lcm.Execute
'   End If
'   If CStr(pMSB) & "." & CStr(pLSB) <> VersaoPeriferico Then
'      VersaoPeriferico = CStr(pMSB) & "." & CStr(pLSB)
'      lcm.CommandText = "UPDATE Base SET VPeriferico = '" & VersaoPeriferico & "' WHERE " & _
'                        "SysID = " & SysId
'      lcm.Execute
'   End If
'End Sub

'Public Sub UpdateReceptor(ByVal fRX As Byte)
'   If receptorID <> fRX Then
'      receptorID = fRX
'      Dim lcm As New ADODB.Command
'      Set lcm.ActiveConnection = cnCD
'      lcm.CommandType = adCmdText
'      lcm.CommandText = "UPDATE Base SET Rx = " & receptorID & " WHERE " & _
'                        "SysID = " & SysId
'      lcm.Execute
'   End If
'End Sub

'Public Sub CheckPersistency()
'
'   Dim lrsBase As New ADODB.Recordset
'
'   lrsBase.CursorLocation = adUseClient
'   lrsBase.CursorType = adOpenStatic
'   lrsBase.LockType = adLockReadOnly
'
'   Dim lsql As String
'   lsql = "SELECT * FROM Base WHERE ([SysID]= " & SysId & ");"
'   lrsBase.Open lsql, cnCD
'   lrsBase.ActiveConnection = Nothing
'
'   If Not lrsBase.EOF Then
'      'Get latest status
'      receptorID = lrsBase("Rx")
''      VersaoRede = lrsBase("VRede")
''      VersaoPeriferico = lrsBase("VPeriferico")
'   Else
'      'Create the Base register
'      lsql = "INSERT INTO Base (SysID, Rx) " & _
'             "VALUES (" & SysId & ", " & receptorID & ");"
'      Dim lcm As New ADODB.Command
'      Set lcm.ActiveConnection = cnCD
'      lcm.CommandType = adCmdText
'      lcm.CommandText = lsql
'      lcm.Execute
'   End If
'
'End Sub
