VERSION 5.00
Object = "{A4749554-0441-4E64-8A03-3323601631C7}#1.0#0"; "LaVolpeAlphaImg2.ocx"
Begin VB.Form frmComm 
   Appearance      =   0  'Flat
   Caption         =   "Comunicação da rede USystemEco"
   ClientHeight    =   6585
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11640
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Moveable        =   0   'False
   ScaleHeight     =   6585
   ScaleWidth      =   11640
   Begin VB.TextBox txtLocal 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   6000
      Visible         =   0   'False
      Width           =   7215
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5700
      ItemData        =   "Comm.frx":0000
      Left            =   0
      List            =   "Comm.frx":0002
      TabIndex        =   0
      ToolTipText     =   "Janela que mostra a comunicação entre o micro e as receptoras e/ou módulos GPRS/IP"
      Top             =   0
      Width           =   11565
   End
   Begin LaVolpeAlphaImg.AlphaImgCtl cmdPrint 
      Height          =   720
      Left            =   9000
      ToolTipText     =   "Imprimi os eventos de comunicação"
      Top             =   5835
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   1270
      Image           =   "Comm.frx":0004
      Effects         =   "Comm.frx":11AF
   End
   Begin LaVolpeAlphaImg.AlphaImgCtl cmdExit 
      Height          =   720
      Left            =   10800
      ToolTipText     =   "Finaliza a tela de comunicação"
      Top             =   5760
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   1270
      Image           =   "Comm.frx":11C7
      Effects         =   "Comm.frx":1ECC
   End
   Begin LaVolpeAlphaImg.AlphaImgCtl cmdlimpar 
      Height          =   720
      Left            =   9960
      ToolTipText     =   "Limpa os eventos da tela"
      Top             =   5760
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   1270
      Image           =   "Comm.frx":1EE4
      Effects         =   "Comm.frx":31E5
   End
End
Attribute VB_Name = "frmComm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdExit_Click()
   Unload frmComm
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

Private Sub cmdlimpar_Click()
   List1.Clear
End Sub

Private Sub cmdlimpar_MouseEnter()
   cmdlimpar.SetRedraw = False
   cmdlimpar.GrayScale = lvicSepia
   cmdlimpar.LightnessPct = -20
   cmdlimpar.SetRedraw = True
End Sub

Private Sub cmdlimpar_MouseExit()
   cmdlimpar.SetRedraw = False
   cmdlimpar.GrayScale = lvicNoGrayScale
   cmdlimpar.LightnessPct = 0
   cmdlimpar.SetRedraw = True
End Sub

Private Sub Form_Load()
    Left = 50
    Top = 50
        
    m_bShowComm = True
            
End Sub

Private Sub Form_Unload(Cancel As Integer)
    m_bShowComm = False
End Sub

Private Sub List1_DblClick()
   List1.Clear
End Sub

Private Sub cmdPrint_Click()
   Load frmDir
   Set frmDir.fCaller = Me
   frmDir.Caption = "Diretório para salvar a Amostra de Eventos da Rede de Comunicação"
   frmDir.Show vbModal

   Dim fileName As String
   fileName = txtLocal.Text & "\Eventos_" & CStr(Day(Now)) & "_" & CStr(Month(Now)) & ".txt"
    
   On Error GoTo FileError
   Dim fso As New FileSystemObject
   Dim txtfile As TextStream
   Set txtfile = fso.CreateTextFile(fileName, True)
   txtfile.WriteLine ("Amostra de Eventos da Rede de Comunicação. ")
   txtfile.WriteBlankLines (1)
   txtfile.WriteLine ("Data: " + CStr(Day(Now)) + "/" + CStr(Month(Now)) + "/" + CStr(Year(Now)))
   txtfile.WriteBlankLines (3)
   ' Stop to show new events
   m_bShowComm = False
   ' Print all events presents on listbox
   Dim i As Integer
   For i = 0 To List1.ListCount - 1
      txtfile.WriteLine (List1.List(i))
   Next i
   ' Now, go back showing new events
   m_bShowComm = True
   txtfile.WriteBlankLines (1)
   txtfile.Close
   MsgBox "Dados salvo no arquivo: " & fileName, sxInformation, sxProname
   Exit Sub
FileError:
   MsgBox "Diretório inválido/protegido ou arquivo existente!, sxExclamation, sxProname"
End Sub

