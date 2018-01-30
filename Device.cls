VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDevice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Esta classe retem os módulos não cadastrados
' na execução corrente

Public sUID As String
Public Serial As String
Public evDate As String
Public tipo As String
Public level As String
Public margin As String
Public name As String
Public recep As String
Public comm As Integer



