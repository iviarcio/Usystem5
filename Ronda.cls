VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRonda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public idRonda As Long
Public descrRonda As String
Public idPercurso As Long
Public idEntity As Long
Public intervalo As Integer
Public status As Boolean
Public descrEntity As String
Public ticketPassagem As Boolean

Public Function Occur_in_Interval() As Boolean
   Dim lPercurso As clsPercurso
   Set lPercurso = lstPercurso.Item(CStr(idPercurso))
   Occur_in_Interval = DateDiff("n", DateAdd("n", intervalo - lPercurso.desvio, lPercurso.Horario), Time) >= 0 And _
      DateDiff("n", Time, DateAdd("n", intervalo + lPercurso.desvio, lPercurso.Horario)) >= 0
End Function

Public Sub Insert()
   Dim lcm As New ADODB.Command
   Set lcm.ActiveConnection = cnDB
   lcm.CommandType = adCmdText
   lcm.CommandText = "INSERT INTO Ronda (DescrRonda, fkPercurso, fkEntity, Intervalo, Status) VALUES ('" & _
                     descrRonda & "', " & idPercurso & ", " & idEntity & ", " & _
                     intervalo & ", " & status & ")"
   lcm.Execute
   Dim rsRonda As ADODB.Recordset
   Set rsRonda = New ADODB.Recordset
   rsRonda.CursorLocation = adUseClient
   rsRonda.CursorType = adOpenStatic
   rsRonda.LockType = adLockReadOnly
   rsRonda.Open "SELECT * FROM Ronda WHERE (descrRonda = '" & descrRonda & _
                   "' AND fkPercurso = " & idPercurso & ");", cnDB
   If Not rsRonda.EOF Then
      idRonda = rsRonda("cpRonda")
   End If
   rsRonda.Close
End Sub

Public Sub Update()
   Dim lcm As New ADODB.Command
   Set lcm.ActiveConnection = cnDB
   lcm.CommandType = adCmdText
   lcm.CommandText = "UPDATE Ronda SET descrRonda = '" & descrRonda & _
                     "', fkPercurso = " & idPercurso & ", fkEntity = " & idEntity & _
                     ", Intervalo = " & intervalo & ", Status = " & status & _
                     " Where (cpRonda = " & idRonda & ");"
   lcm.Execute
End Sub

Public Sub Remove()
   Dim lcm As New ADODB.Command
   Set lcm.ActiveConnection = cnDB
   lcm.CommandType = adCmdText
   lcm.CommandText = "DELETE  FROM Ronda Where (cpRonda = " & _
                     idRonda & ")"
   lcm.Execute
End Sub


