VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPercurso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public idPercurso As Long
Public descrPercurso As String
Public Horario As Date
Public desvio As Integer
Public valSegSex As Boolean
Public valSab As Boolean
Public valDom As Boolean
Public status As Boolean

Public lstHorario As New Collection
Public lstRonda As New Collection

Public Sub RemoveRondas()
   Dim lR As clsRonda
   For Each lR In lstRonda
      lR.Remove
      lstRonda.Remove 1
   Next
End Sub

Public Sub LoadRondas()
   Dim tronda As clsRonda
   Dim rsRonda As ADODB.Recordset
   Set rsRonda = New ADODB.Recordset
   rsRonda.CursorLocation = adUseClient
   rsRonda.CursorType = adOpenStatic
   rsRonda.LockType = adLockReadOnly
   rsRonda.Open "SELECT * FROM Ponto_de_Ronda Where fkPercurso = " & idPercurso & ";", cnDB
   While Not rsRonda.EOF
      Set tronda = New clsRonda
      On Error Resume Next
      With tronda
         .idRonda = rsRonda("cpRonda")
         .descrRonda = rsRonda("DescrRonda")
         .idPercurso = rsRonda("fkPercurso")
         .idEntity = rsRonda("fkEntity")
         .intervalo = rsRonda("Intervalo")
         .status = rsRonda("Status")
         .descrEntity = rsRonda("Descr_Entity")
         .ticketPassagem = False
      End With
      On Error GoTo 0
      lstRonda.Add Item:=tronda, Key:=CStr(rsRonda("cpRonda"))
      rsRonda.MoveNext
   Wend
   rsRonda.Close
End Sub

Public Sub CleanTickets()
   Dim tronda As clsRonda
   For Each tronda In lstRonda
      tronda.ticketPassagem = False
   Next
   Dim tHorario As clsHorario
   Dim Index As Integer
   Index = 1
   For Each tHorario In lstHorario
      If Index <> 1 Then
         tHorario.ticketHorario = False
      End If
      Index = Index + 1
   Next
   Horario = lstHorario.Item(1).Horario
End Sub

'FIXME
Public Sub CheckRonda(fPassagem As Integer)
   Dim tronda As clsRonda
   Dim tEntity As clsEntity
   Dim lModule As clsModule
   For Each tronda In lstRonda
      If fPassagem - (tronda.intervalo + desvio) >= 0 Then
         If tronda.status And Not tronda.ticketPassagem Then
            'gera evento de ronda não executada
            Set tEntity = lstEntity.Item(CStr(tronda.idEntity))
            For Each lModule In tEntity.localModule
               If lModule.mTipo = s_Ronda Then
                  'aqui "stFalha" é usada para indicar que não ocorreu a passagem de ronda
                  lModule.evDate = Now
                  'lModule.SZona = stFalha
               End If
            Next
         End If
      End If
   Next
End Sub

Public Function MaxInterval() As Integer
   Dim lastInterval  As Integer
   Dim tronda As clsRonda
   lastInterval = 0
   For Each tronda In lstRonda
      If tronda.status Then
         If tronda.intervalo > lastInterval Then
            lastInterval = tronda.intervalo
         End If
      End If
   Next
   MaxInterval = lastInterval + desvio + 1
End Function

Public Function Active() As Boolean
   If status Then
      If IsHolliday() Then
         Active = valDom
      ElseIf curWeekday = vbSunday Then
         Active = valDom
      ElseIf curWeekday = vbSaturday Then
         Active = valSab
      Else
         Active = valSegSex
      End If
   Else
      Active = False
   End If
End Function

Public Sub Insert()
   Dim lcm As New ADODB.Command
   Set lcm.ActiveConnection = cnDB
   lcm.CommandType = adCmdText
   lcm.CommandText = "INSERT INTO Percurso (DescrPercurso, Horario) VALUES ('" & _
                     descrPercurso & "', '" & Horario & "')"
   lcm.Execute
   Dim rsPercurso As ADODB.Recordset
   Set rsPercurso = New ADODB.Recordset
   rsPercurso.CursorLocation = adUseClient
   rsPercurso.CursorType = adOpenStatic
   rsPercurso.LockType = adLockReadOnly
   rsPercurso.Open "SELECT * FROM Percurso WHERE (descrPercurso = '" & descrPercurso & _
                   "' AND Horario = #" & Horario & "#);", cnDB
   If Not rsPercurso.EOF Then
      idPercurso = rsPercurso("cpPercurso")
   End If
   rsPercurso.Close
End Sub

Public Sub Update()
   Dim lcm As New ADODB.Command
   Set lcm.ActiveConnection = cnDB
   lcm.CommandType = adCmdText
   lcm.CommandText = "UPDATE Percurso SET descrPercurso = '" & descrPercurso & _
                     "', horario = #" & Horario & "#, desvio = " & desvio & _
                     ", valSegSex = " & valSegSex & ", valSab = " & valSab & _
                     ", valDomFer = " & valDom & ", Status = " & status & _
                     " Where (cpPercurso = " & idPercurso & ");"
   lcm.Execute
End Sub

Public Sub Remove()
   Dim lcm As New ADODB.Command
   Set lcm.ActiveConnection = cnDB
   lcm.CommandType = adCmdText
   lcm.CommandText = "DELETE  FROM Percurso Where (cpPercurso = " & _
                     idPercurso & ")"
   lcm.Execute
End Sub

Public Sub Insert_Horario(fHour As String)
   Dim tHorario As clsHorario
   Set tHorario = New clsHorario
   tHorario.Horario = fHour
   tHorario.ticketHorario = False
   lstHorario.Add Item:=tHorario
End Sub

Public Sub Save_Horarios()
   Dim lcm As New ADODB.Command
   Set lcm.ActiveConnection = cnDB
   lcm.CommandType = adCmdText
   Dim tHorario As clsHorario
   For Each tHorario In lstHorario
      lcm.CommandText = "INSERT INTO RdHorario (fkPercurso, Horario) VALUES (" & _
                        idPercurso & ", '" & tHorario.Horario & "')"
      lcm.Execute
   Next
End Sub

Public Sub Remove_Horarios()
   On Error Resume Next
   Dim tHorario As clsHorario
   For Each tHorario In lstHorario
      lstHorario.Remove 1
   Next
   Dim lcm As New ADODB.Command
   Set lcm.ActiveConnection = cnDB
   lcm.CommandType = adCmdText
   lcm.CommandText = "DELETE  FROM RdHorario Where (fkPercurso = " & _
                     idPercurso & ")"
   lcm.Execute
   On Error GoTo 0
End Sub

Public Sub Load_Horarios()
   Dim tHorario As clsHorario
   Dim rsHorario As ADODB.Recordset
   Set rsHorario = New ADODB.Recordset
   rsHorario.CursorLocation = adUseClient
   rsHorario.CursorType = adOpenStatic
   rsHorario.LockType = adLockReadOnly
   rsHorario.Open "SELECT * FROM RdHorario WHERE fkPercurso =" & idPercurso & " ORDER BY Horario;", cnDB
   Dim Index As Integer
   Index = 1
   While Not rsHorario.EOF
      Set tHorario = New clsHorario
      On Error Resume Next
      With tHorario
         .idHorario = Index
         .Horario = rsHorario("Horario")
         .ticketHorario = False
      End With
      On Error GoTo 0
      lstHorario.Add Item:=tHorario, Key:=CStr(tHorario.idHorario)
      rsHorario.MoveNext
      Index = Index + 1
   Wend
   rsHorario.Close
   On Error Resume Next
   Horario = "00:00"
   Horario = lstHorario.Item(1).Horario
   On Error GoTo 0
End Sub

Public Sub NextHorario()
   Dim tHorario As clsHorario
   For Each tHorario In lstHorario
      If Not tHorario.ticketHorario Then
         Horario = tHorario.Horario
         tHorario.ticketHorario = True
         'clear tickets de ronda
         Dim tronda As clsRonda
         For Each tronda In lstRonda
            tronda.ticketPassagem = False
         Next
         Exit Sub
      End If
   Next
End Sub
