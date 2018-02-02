VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsQueue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' this is the collection that holds the message values
Private msgValues As New Collection

' add a new msg value to the queue, skiping duplicate ones
Public Function Enqueue(Value As String, commPort As Integer) As Boolean

    'Se a mensagem recebida for um evento de Sensor ou Repeater ela devera conter 32 posicoes.
    'Separa somente:

        'HEADER+SIZE+UIDo+UIDh+PTI+STAT1+STAT0 com 26 posicoes

    'desconsiderando o Nivel de Sinal, a Margem de Ruido e o CheckSum cuja variacao nao configura como
    'nao duplicidade. Caso contrario, enfilera a mensagem, pois pode ser o HeartBeat do Serial Receiver
    'cujo comprimento = 10.

    Dim tEvent As New clsEvent
    If Len(Value) > 26 Then
        tEvent.evHex = Left(Value, 26)
    Else
        tEvent.evHex = Value
    End If

    'Check duplicate messages
    Enqueue = False
    Dim v As clsEvent
    For Each v In msgValues
        If v.evHex = tEvent.evHex Then Exit Function
    Next

    'If No duplicate then...
    With tEvent
        .evHex = Value ' full message
        'split message values
        .comm = commPort
        .sHeader = Left(Value, 2)
        .sMID = Mid(Value, 5, 2)
        If .sHeader = H_Serial Then     'Treat Serial Receiver
            .ssData = Mid(Value, 5, 2)
            .sStat0 = Mid(Value, 7, 2)
            .sDeviceName = "Receiver"
        ElseIf .sHeader = H_Device And .sMID = MID_Repeater Then    'Treat Repeaters
            .sUIDo = Mid(Value, 5, 8)
            .sUIDh = Mid(Value, 13, 8)
            .sClasse = Mid(Value, 21, 2)
            .sStat1 = Mid(Value, 23, 2)
            .sStat0 = Mid(Value, 25, 2)
            .ssLevel = Mid(Value, 27, 2)
            .ssMargin = Mid(Value, 29, 2)
            .sDeviceName = "Repeater"
        ElseIf .sHeader = H_Device And .sMID = MID_Sensor Then    'Treat Sensors
            .sUIDo = Mid(Value, 5, 8)
            .sUIDh = Mid(Value, 13, 8)
            .sPTI = Mid(Value, 21, 2)
            .sStat1 = Mid(Value, 23, 2)
            .sStat0 = Mid(Value, 25, 2)
            .ssLevel = Mid(Value, 27, 2)
            .ssMargin = Mid(Value, 29, 2)
            .sDeviceName = Device_Name(.sPTI)    'Seek device name on base
        End If
        .evDate = Format(Now, "dd/mm/yyyy hh:mm:ss")
    End With
    msgValues.Add tEvent
    Enqueue = True
End Function

' Read the next value from the queue and remove it
' Raises an error if the queue is empty
Public Function Dequeue() As clsEvent
'   DoEvents
   Set Dequeue = msgValues.Item(1)
   msgValues.Remove 1
End Function

' Return the next value from the queue, without removing it
' raises an error if queue is empty
Public Function Peek() As clsEvent
   Set Peek = msgValues.Item(1)
End Function

' Return the number of values in the queue
Public Function Count() As Integer
   Count = msgValues.Count
End Function

' clear the queue
Public Sub Clear()
   Set msgValues = New Collection
End Sub
