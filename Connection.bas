Attribute VB_Name = "modConnection"
Public strDCn As String
Public strECn As String
Public strWCn As String
Public oConn As ADODB.Connection
Public oCmd As ADODB.Command

Private Sub Connection_Constructor()
'Constroi os strings de conexão com os bancos de dados
   strDCn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & _
             LogOn.cboServidor.Text & _
            "\Clientes80.mdb;Persist Security Info=False"
   strECn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & _
             LogOn.cboServidor.Text & _
            "\Eventos80.mdb;Persist Security Info=False"
   strWCn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & _
             LogOn.cboServidor.Text & _
            "\CondorWeb80.mdb;Persist Security Info=False"
End Sub

Public Sub Connect()
   On Error GoTo TratarErro
   Connection_Constructor
   Set oConn = New ADODB.Connection
   With oConn
      .ConnectionString = strDCn
      .ConnectionTimeout = 45
      .CursorLocation = adUseClient
      .Open
   End With
   Exit Sub
TratarErro:
   On Error GoTo 0
   Err.Raise vbObjectError + 1000, _
             "Scre.Connection.Connect", _
             "Falha na conexão do banco de dados! Contate seu suporte técnico.", ""
End Sub

Public Sub Disconnect(ByVal fTp As Integer)
   oConn.Close
End Sub

Public Sub ExecSp(ByVal fs_SpStr As String)
' executa a stored procedure passada no argumento
   On Error GoTo TratarErro
   Set oCmd = New ADODB.Command
   With oCmd
      Set .ActiveConnection = oConn
      .CommandText = fs_SpStr
      .CommandType = adCmdText
      .CommandTimeout = 0
      .Execute
   End With
   Exit Sub
TratarErro:
   Dim ErrNro
   ErrNro = Err.Number - vbObjectError
   On Error GoTo 0
   Select Case ErrNro
      Case 1000
         Err.Raise Err.Number, _
                   Err.Source, _
                   Err.Description, ""
      Case Else
         Err.Raise vbObjectError + 1002, _
                   "Scre.Global.ExecSp", _
                   "Falha na execução da stored procedure " & fs_SpStr & " .", ""
   End Select
End Sub

Public Function ExecSpGetRs(ByVal fs_SpStr As String) As ADODB.Recordset
'executa a stored procedure passada no argumento e retorna um recordset
   Dim oRs As New ADODB.Recordset
   On Error GoTo TratarErro
   oRs.Open fs_SpStr, oConn, adOpenDynamic, adLockOptimistic, adCmdText
   Set ExecSpGetRs = oRs
   Exit Function
TratarErro:
   Dim ErrNro
   ErrNro = Err.Number - vbObjectError
   On Error GoTo 0
   Select Case ErrNro
      Case 1000
         Err.Raise Err.Number, _
                   Err.Source, _
                   Err.Description, ""
      Case Else
         Err.Raise vbObjectError + 1003, _
                   "Scre.Global.ExecSp", _
                   "Falha na execução da stored procedure " & fs_SpStr & " .", ""
   End Select
End Function

Public Function ExecCmd(ByVal fs_SpStr As String) As ADODB.Recordset
'executa o comando passado no argumento e retorna um recordset
   Dim oRs As New ADODB.Recordset
   On Error GoTo TratarErro
   oRs.Open fs_SpStr, oConn, adOpenDynamic, adLockOptimistic, adCmdText
   Set ExecCmd = oRs
   Exit Function
TratarErro:
   Dim ErrNro
   ErrNro = Err.Number - vbObjectError
   On Error GoTo 0
   Select Case ErrNro
      Case 1000
         Err.Raise Err.Number, _
                   Err.Source, _
                   Err.Description, ""
      Case Else
         Err.Raise vbObjectError + 1003, _
                   "Scre.Global.ExecSp", _
                   "Falha na execução do comando " & fs_SpStr & " .", ""
   End Select
End Function

Public Function Snapshot(ByVal fs_SpStr As String) As ADODB.Recordset
'executa o comando passado no argumento e retorna um recordset
   Dim oRs As New ADODB.Recordset
   On Error GoTo TratarErro
   oRs.Open fs_SpStr, oConn, adOpenStatic, adLockReadOnly, adCmdText
   Set Snapshot = oRs
   Exit Function
TratarErro:
   Dim ErrNro
   ErrNro = Err.Number - vbObjectError
   On Error GoTo 0
   Select Case ErrNro
      Case 1000
         Err.Raise Err.Number, _
                   Err.Source, _
                   Err.Description, ""
      Case Else
         Err.Raise vbObjectError + 1003, _
                   "Scre.Global.ExecSp", _
                   "Falha na execução do comando " & fs_SpStr & " .", ""
   End Select
End Function
