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

Public f As frmPlanta            'Refer�ncia � planta
Public n As Integer              'n�mero do piso/andar
Public rStep As Single            'Posi��o do Label que refer�ncia o Piso
Public rCaption As String        'Descri��o do Piso/Andar
Public c_bSetores As Boolean     'Indica se Entidades (setores) foram criadas

