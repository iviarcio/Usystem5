Attribute VB_Name = "mdRonda"
Option Explicit

Public Sub Clear_Tickets_Percurso()
   Dim tPercurso As clsPercurso
   For Each tPercurso In lstPercurso
      tPercurso.CleanTickets
   Next
End Sub

Public Sub Treat_Percurso_Ronda()
   Dim tPercurso As clsPercurso
   For Each tPercurso In lstPercurso
      If tPercurso.Active Then
         If DateDiff("n", tPercurso.Horario, Time) >= 0 Then
            If DateDiff("n", Time, DateAdd("n", tPercurso.MaxInterval, tPercurso.Horario)) >= 0 Then
               tPercurso.CheckRonda DateDiff("n", tPercurso.Horario, Time)
            Else
               tPercurso.NextHorario
            End If
         End If
      End If
   Next
End Sub

Public Function GetRonda(lEntity As Integer) As clsRonda
   Dim lPercurso As clsPercurso
   Dim lronda As clsRonda
   For Each lPercurso In lstPercurso
      For Each lronda In lPercurso.lstRonda
         If lronda.idEntity = lEntity Then
            Set GetRonda = lronda
            Exit Function
         End If
      Next
   Next
End Function

Public Function IsHolliday() As Boolean
   IsHolliday = False   'for while
End Function
