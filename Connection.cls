VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsConnection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public ErroDB As Boolean
Public FileNotMdb As Boolean

'Executa uma stored procedure passada como argumento na base de dados
Public Sub ExecSp(ByVal fs_SpStr As String)
   On Error GoTo TratarErro
   
   Dim ltry As Boolean
   ltry = True
   Dim oCmd As New ADODB.Command
   With oCmd
      Set .ActiveConnection = cnDB
      .CommandText = fs_SpStr
      .CommandType = adCmdText
      .CommandTimeout = 0
TryExecSp:
      .Execute
   End With
   Exit Sub
   
TratarErro:
   If ltry Then
      ltry = False
      DoEvents
      Resume TryExecSp
   End If
   
   Dim ErrNro As Long
   ErrNro = Err.Number - vbObjectError
   On Error GoTo 0
   Select Case ErrNro
      Case 1000
         Err.Raise Err.Number, _
                   Err.Source, _
                   Err.Description, ""
      Case Else
         ErroNet = True
         Err.Raise vbObjectError + 1001, _
                   "Erro na conexão com a base de dados so USystem.", _
                   "Forneça a seguinte informação ao seu suporte técnico. " & _
                   "(" & 1001 & ") " & Err.Description & ": " & fs_SpStr & " .", ""
   End Select

End Sub

'Executa uma stored procedure na base de dados e retorna um RECORDSET
Public Function ExecSpGetRs(ByVal fs_SpStr As String) As ADODB.Recordset
   Dim ltry As Boolean
   ltry = True
   Dim oRs As New ADODB.Recordset
   
   On Error GoTo TratarErro
   
TryGetRs:
   oRs.Open fs_SpStr, cnDB, adOpenDynamic, adLockOptimistic, adCmdText
   Set ExecSpGetRs = oRs
   Exit Function
   
TratarErro:
   If ltry Then
      DoEvents
      ltry = False
      Resume TryGetRs
   End If
   Dim ErrNro
   ErrNro = Err.Number - vbObjectError
   On Error GoTo 0
   Select Case ErrNro
      Case 1000
         Err.Raise Err.Number, _
                   Err.Source, _
                   Err.Description, ""
      Case Else
         ErroNet = True
         Err.Raise vbObjectError + 1002, _
                   "Erro de leitura na base de dados do USystem.", _
                   "Forneça a seguinte informação ao seu suporte técnico. " & _
                   "(" & 1002 & ") " & Err.Description & ": " & fs_SpStr & " .", ""
   End Select
   
End Function

'Executa uma stored procedure na base de dados e retorna um recordset estático (read-only)
Public Function Snapshot(ByVal fs_SpStr As String) As ADODB.Recordset
   Dim ltry As Boolean
   ltry = True
   Dim oRs As New ADODB.Recordset
   
   On Error GoTo TratarErro
trySnapshot:
   oRs.Open fs_SpStr, cnDB, adOpenStatic, adLockReadOnly, adCmdText
   Set Snapshot = oRs
   Exit Function
TratarErro:
   If ltry Then
      DoEvents
      ltry = False
      Resume trySnapshot
   End If
   Dim ErrNro
   ErrNro = Err.Number - vbObjectError
   On Error GoTo 0
   Select Case ErrNro
      Case 1000
         Err.Raise Err.Number, _
                   Err.Source, _
                   Err.Description, ""
      Case Else
         ErroNet = True
         Err.Raise vbObjectError + 1003, _
                   "Erro de leitura na base de dados do USystem.", _
                   "Forneça a seguinte informação ao seu suporte técnico. " & _
                   "(" & 1003 & ") " & Err.Description & ": " & fs_SpStr & " .", ""
   End Select
End Function

'Executa uma stored procedure passada como argumento na base CADASTRO
Public Sub ExecQuery(ByVal fs_SpStr As String)
   Dim ltry As Boolean
   ltry = True
   On Error GoTo TratarErro
   Dim oCmd As New ADODB.Command
   
   With oCmd
      Set .ActiveConnection = cnDB
      .CommandText = fs_SpStr
      .CommandType = adCmdStoredProc
      .CommandTimeout = 0
tryQuery:
      .Execute
   End With
   Exit Sub
   
TratarErro:
   If ltry Then
      DoEvents
      ltry = False
      Resume tryQuery
   End If
   Dim ErrNro
   ErrNro = Err.Number - vbObjectError
   On Error GoTo 0
   Select Case ErrNro
      Case 1000
         Err.Raise Err.Number, _
                   Err.Source, _
                   Err.Description, ""
      Case Else
         ErroNet = True
         Err.Raise vbObjectError + 1007, _
                   "Erro de leitura na base de dados do Usystem.", _
                   "Forneça a seguinte informação ao seu suporte técnico. " & _
                   "(" & 1007 & ") " & Err.Description & ": " & fs_SpStr & " .", ""
   End Select
End Sub
