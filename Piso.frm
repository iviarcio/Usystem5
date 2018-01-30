VERSION 5.00
Object = "{A4749554-0441-4E64-8A03-3323601631C7}#1.0#0"; "LaVolpeAlphaImg2.ocx"
Begin VB.Form frmPiso 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Dados do Piso/Andar"
   ClientHeight    =   1170
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   4935
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Piso.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1170
   ScaleWidth      =   4935
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtDescrPiso 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1320
      MaxLength       =   20
      TabIndex        =   3
      Top             =   600
      Width           =   2535
   End
   Begin VB.TextBox txtNumPiso 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3240
      TabIndex        =   2
      Top             =   120
      Width           =   615
   End
   Begin LaVolpeAlphaImg.AlphaImgCtl cmdExit 
      Height          =   720
      Left            =   4080
      Top             =   240
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   1270
      Image           =   "Piso.frx":030A
      Effects         =   "Piso.frx":100F
   End
   Begin VB.Label Label2 
      Caption         =   "Descrição:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Número do Piso/Andar:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2535
   End
End
Attribute VB_Name = "frmPiso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdExit_Click()
   m_bUserUnload = True
   Unload Me
End Sub

Private Sub cmdExit_MouseEnter()
   cmdExit.SetRedraw = False
   cmdExit.GrayScale = lvicSepia
   cmdExit.LightnessPct = -20
   cmdExit.SetRedraw = True
End Sub

Private Sub cmdExit_MouseExit()
   cmdExit.SetRedraw = False
   cmdExit.GrayScale = lvicNoGrayScale
   cmdExit.LightnessPct = 0
   cmdExit.SetRedraw = True
End Sub

Private Sub Form_Load()
   Left = (Screen.Width - Width) / 2   ' Center form horizontally.
   Top = (Screen.Height - Height) / 2  ' Center form vertically.
   cmdExit.Enabled = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If m_bUserUnload Then
      Cancel = True
      Hide
   End If
End Sub

Private Sub txtNumPiso_Change()
    cmdExit.Enabled = False
    If Not IsNumeric(txtNumPiso) Then
        Beep
        txtNumPiso = ""
    ElseIf txtNumPiso <= 0 Then
        Beep
        txtNumPiso = ""
    Else
        cmdExit.Enabled = True
    End If
End Sub

