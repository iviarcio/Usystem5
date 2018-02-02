VERSION 5.00
Object = "{A4749554-0441-4E64-8A03-3323601631C7}#1.0#0"; "LaVolpeAlphaImg2.ocx"
Begin VB.Form frmGrupo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Seleção de Grupo"
   ClientHeight    =   2190
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   5070
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2190
   ScaleWidth      =   5070
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox lstTipoGrupo 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "Grupo.frx":0000
      Left            =   120
      List            =   "Grupo.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   840
      Width           =   3975
   End
   Begin VB.Label Label21 
      Caption         =   "Selecione o grupo no qual a ação terá efeito. Tecle em ""Ok"" para executar ou selecione "" X"" para cancelar a operação."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   795
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   4095
      WordWrap        =   -1  'True
   End
   Begin LaVolpeAlphaImg.AlphaImgCtl cmdCancel 
      Height          =   720
      Left            =   4320
      ToolTipText     =   "Cancelar"
      Top             =   840
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   1270
      Image           =   "Grupo.frx":0004
      Effects         =   "Grupo.frx":0C50
   End
   Begin LaVolpeAlphaImg.AlphaImgCtl cmdOk 
      Height          =   720
      Left            =   4320
      ToolTipText     =   "Selecionar"
      Top             =   0
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   1270
      Image           =   "Grupo.frx":0C68
      Effects         =   "Grupo.frx":1807
   End
End
Attribute VB_Name = "frmGrupo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    tGrupo = -1
    Unload Me
End Sub

Private Sub cmdOk_Click()
    tGrupo = lstTipoGrupo.ListIndex
    Unload Me
End Sub

Private Sub Form_Load()
    tGrupo = -1
    LoadGrupo
End Sub

Private Sub LoadGrupo()
    ' First, insert the name "Geral"
    lstTipoGrupo.AddItem ("Geral (não especificado)")
    Dim rsGrupo As ADODB.Recordset
    Set rsGrupo = New ADODB.Recordset
    rsGrupo.CursorLocation = adUseClient
    rsGrupo.CursorType = adOpenStatic
    rsGrupo.LockType = adLockReadOnly
    rsGrupo.Open "Select * From Grupo", cnDB
    While Not rsGrupo.EOF
       lstTipoGrupo.AddItem (rsGrupo("Descrição"))
       rsGrupo.MoveNext
    Wend
    rsGrupo.Close
End Sub

