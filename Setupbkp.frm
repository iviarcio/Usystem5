VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{A4749554-0441-4E64-8A03-3323601631C7}#1.0#0"; "LaVolpeAlphaImg2.ocx"
Begin VB.Form frmSetupbkp 
   Caption         =   "Setup"
   ClientHeight    =   3210
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5910
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Setupbkp.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3210
   ScaleWidth      =   5910
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      Caption         =   "Local do Backup"
      Height          =   975
      Left            =   0
      TabIndex        =   7
      Top             =   1200
      Width           =   5895
      Begin VB.TextBox txtLocal 
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Width           =   4575
      End
      Begin LaVolpeAlphaImg.AlphaImgCtl cmdLocal 
         Height          =   720
         Left            =   5040
         ToolTipText     =   "Procurar/Selecionar"
         Top             =   160
         Width           =   720
         _ExtentX        =   1270
         _ExtentY        =   1270
         Image           =   "Setupbkp.frx":0442
         Effects         =   "Setupbkp.frx":1601
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Backup"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   975
      Left            =   0
      TabIndex        =   2
      Top             =   120
      Width           =   5895
      Begin VB.CheckBox chkAuto 
         Alignment       =   1  'Right Justify
         Caption         =   "Realizar Backup Automático ?"
         ForeColor       =   &H8000000D&
         Height          =   315
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   2535
      End
      Begin MSMask.MaskEdBox mskHorario 
         Height          =   315
         Left            =   4200
         TabIndex        =   4
         Top             =   360
         Visible         =   0   'False
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "99:99"
         PromptChar      =   "_"
      End
      Begin VB.Label lblHorario 
         Caption         =   "Diariamente as "
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   3000
         TabIndex        =   6
         Top             =   390
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label lblHorario2 
         Caption         =   "horas"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   4920
         TabIndex        =   5
         Top             =   390
         Visible         =   0   'False
         Width           =   495
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Manter os Dados"
      Height          =   855
      Left            =   0
      TabIndex        =   0
      Top             =   2280
      Width           =   4695
      Begin VB.TextBox txtDados 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   3000
         TabIndex        =   1
         Text            =   "30"
         Top             =   345
         Width           =   495
      End
      Begin VB.Label Label3 
         Caption         =   "dias."
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   3600
         TabIndex        =   10
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Label2 
         Caption         =   "Manter os dados dos últimos"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   840
         TabIndex        =   9
         Top             =   360
         Width           =   2295
      End
   End
   Begin LaVolpeAlphaImg.AlphaImgCtl cmdExit 
      Height          =   720
      Left            =   5040
      ToolTipText     =   "Fechar Configuração de Backup"
      Top             =   2400
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   1270
      Image           =   "Setupbkp.frx":1619
      Effects         =   "Setupbkp.frx":231E
   End
End
Attribute VB_Name = "frmSetupbkp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private f_bChange As Boolean

Private Sub chkAuto_Click()
   lblHorario.Visible = (chkAuto.Value = vbChecked)
   lblHorario2.Visible = (chkAuto.Value = vbChecked)
   mskHorario.Visible = (chkAuto.Value = vbChecked)
   m_bBackupAuto = chkAuto.Value = vbChecked
End Sub

Private Sub cmdExit_Click()
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

Private Sub cmdLocal_Click()
   Load frmDir
   Set frmDir.fCaller = Me
   frmDir.Caption = "Diretório para Backup Automático"
   frmDir.Show vbModal
   If Not (m_sBPath = txtLocal) Then
      m_sBPath = txtLocal
      Call SaveSetting("USystemEco", "Options", "Backup", m_sBPath)
   End If
End Sub

Private Sub cmdLocal_MouseEnter()
   cmdLocal.SetRedraw = False
   cmdLocal.GrayScale = lvicSepia
   cmdLocal.LightnessPct = -20
   cmdLocal.SetRedraw = True
End Sub

Private Sub cmdLocal_MouseExit()
   cmdLocal.SetRedraw = False
   cmdLocal.GrayScale = lvicNoGrayScale
   cmdLocal.LightnessPct = 0
   cmdLocal.SetRedraw = True
End Sub


Private Sub Form_Load()
   Dim success As Long
   success = SetWindowPos(frmSetupbkp.hwnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
   curhwnd = frmSetupbkp.hwnd
   Left = (Screen.Width - Width) / 2   ' Center form horizontally.
   Top = (Screen.Height - Height) / 2  ' Center form vertically.
   If m_bBackupAuto Then
      chkAuto = vbChecked
      mskHorario = m_sHorario
   Else
      chkAuto = vbUnchecked
   End If
   txtDados = m_iEvKeep
   txtLocal = m_sBPath
   f_bChange = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If f_bChange Then
      m_iEvKeep = txtDados
      Dim lcmd As New ADODB.Command
      Set lcmd.ActiveConnection = cnCD
      lcmd.CommandType = adCmdText
      lcmd.CommandText = "UPDATE Horario SET KeepData=" & m_iEvKeep
      lcmd.Execute
      If Not (m_sBPath = txtLocal) Then
         m_sBPath = txtLocal
         Call SaveSetting("USystemEco", "Options", "Backup", m_sBPath)
      End If
   End If
End Sub

Private Sub mskHorario_Change()
   m_sHorario = mskHorario
End Sub

Private Sub txtDados_Change()
   If IsNumeric(txtDados) Then f_bChange = True
End Sub

Private Sub txtLocal_Change()
   f_bChange = True
End Sub
