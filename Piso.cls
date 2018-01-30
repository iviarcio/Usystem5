VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPiso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public f As frmPlanta            'Referência à planta
Public n As Integer              'número do piso/andar
Public rStep As Single            'Posição do Label que referência o Piso
Public rCaption As String        'Descrição do Piso/Andar
Public c_bSetores As Boolean     'Indica se Entidades (setores) foram criadas

