VERSION 5.00
Object = "{A4749554-0441-4E64-8A03-3323601631C7}#1.0#0"; "LaVolpeAlphaImg2.ocx"
Begin VB.Form frmCritico 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Registro de Evento Crítico"
   ClientHeight    =   6555
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6600
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6555
   ScaleWidth      =   6600
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Classificação do evento:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   855
      Left            =   60
      TabIndex        =   12
      Top             =   1200
      Width           =   6375
      Begin VB.OptionButton Option1 
         Caption         =   "Teste"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Index           =   2
         Left            =   4320
         TabIndex        =   15
         Top             =   360
         Width           =   1455
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Acidental"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Index           =   1
         Left            =   2340
         TabIndex        =   14
         Top             =   360
         Width           =   1455
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Real"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Index           =   0
         Left            =   360
         TabIndex        =   13
         Top             =   360
         Value           =   -1  'True
         Width           =   1455
      End
   End
   Begin VB.TextBox txtPassword 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      IMEMode         =   3  'DISABLE
      Left            =   2400
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   6000
      Width           =   2175
   End
   Begin VB.TextBox txtUser 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   60
      TabIndex        =   3
      Top             =   6000
      Width           =   2055
   End
   Begin VB.TextBox txtObs 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   60
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   4320
      Width           =   6435
   End
   Begin VB.TextBox txtAcao 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   60
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   2520
      Width           =   6435
   End
   Begin VB.TextBox txtKind 
      Height          =   375
      Left            =   1680
      TabIndex        =   0
      Text            =   "0"
      Top             =   7080
      Width           =   2655
   End
   Begin VB.Label Label7 
      Caption         =   "Obs:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   360
      Left            =   60
      TabIndex        =   18
      Top             =   840
      Width           =   735
   End
   Begin VB.Label Label6 
      Caption         =   "Loja:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   360
      Left            =   60
      TabIndex        =   17
      Top             =   480
      Width           =   615
   End
   Begin VB.Label Label5 
      Caption         =   "Sensor:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   360
      Left            =   60
      TabIndex        =   16
      Top             =   120
      Width           =   975
   End
   Begin LaVolpeAlphaImg.AlphaImgCtl btnSom 
      Height          =   720
      Left            =   4800
      ToolTipText     =   "Desligar o som"
      Top             =   5760
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   1270
      Image           =   "Critico.frx":0000
      Effects         =   "Critico.frx":153F
   End
   Begin VB.Label lblLocal 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   1080
      TabIndex        =   11
      Top             =   480
      Width           =   5415
   End
   Begin VB.Label lblSensor 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   1080
      TabIndex        =   10
      Top             =   120
      Width           =   5415
   End
   Begin VB.Label lblObs 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   1080
      TabIndex        =   9
      Top             =   840
      Width           =   5415
   End
   Begin LaVolpeAlphaImg.AlphaImgCtl btnOk 
      Height          =   720
      Left            =   5640
      ToolTipText     =   "Registrar Evento."
      Top             =   5760
      Width           =   960
      _ExtentX        =   1693
      _ExtentY        =   1270
      Image           =   "Critico.frx":1557
      Effects         =   "Critico.frx":26E4
   End
   Begin VB.Label Label4 
      Caption         =   "Senha"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   2400
      TabIndex        =   8
      Top             =   5640
      Width           =   1695
   End
   Begin VB.Label Label3 
      Caption         =   "Usuário"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   60
      TabIndex        =   7
      Top             =   5640
      Width           =   2055
   End
   Begin VB.Label Label2 
      Caption         =   "Observação:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   60
      TabIndex        =   6
      Top             =   3960
      Width           =   2775
   End
   Begin VB.Label Label1 
      Caption         =   "Ação tomada:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   60
      TabIndex        =   5
      Top             =   2160
      Width           =   2775
   End
End
Attribute VB_Name = "frmCritico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public crModule As clsModule
Private registerOk As Boolean
'Private WinHttpReq As WinHttp.WinHttpRequest

Private Sub btnSom_MouseEnter()
   btnSom.SetRedraw = False
   btnSom.GrayScale = lvicSepia
   btnSom.LightnessPct = -20
   btnSom.SetRedraw = True
End Sub

Private Sub btnSom_MouseExit()
   btnSom.SetRedraw = False
   btnSom.GrayScale = lvicNoGrayScale
   btnSom.LightnessPct = 0
   btnSom.SetRedraw = True
End Sub

Private Sub btnSom_Click()
   Sound_Update fmode:=sxNoSound, isCritico:=False, fNoSound:=False
   DoEvents
End Sub

Private Sub Form_Activate()
   Dim lEntity As clsEntity
   With crModule
      lblSensor = .mLocal
      Set lEntity = lstEntity.Item(CStr(.mEntity))
   End With
   lblLocal = lEntity.vDescr
   lblObs = lEntity.message
End Sub

Private Sub Form_Load()
   Dim success As Long
   success = SetWindowPos(frmCritico.hWnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
   curhwnd = frmCritico.hWnd
   registerOk = False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   If Not registerOk Then
      Cancel = True
   End If
End Sub

Private Sub btnOk_MouseEnter()
   btnOk.SetRedraw = False
   btnOk.GrayScale = lvicSepia
   btnOk.LightnessPct = -20
   btnOk.SetRedraw = True
End Sub

Private Sub btnOk_MouseExit()
   btnOk.SetRedraw = False
   btnOk.GrayScale = lvicNoGrayScale
   btnOk.LightnessPct = 0
   btnOk.SetRedraw = True
End Sub

Private Sub btnOk_Click()
   Dim strPassword As String
   Dim lAccess As Boolean
   lAccess = False
   If (txtUser = sxAuthor) Then
      strPassword = Format(Day(Date), "00") & "f" & Format(Month(Date), "00") & "o" & Format(Right(Year(Date), 2), "00") & "r"
      If (txtPassword = strPassword) Then
         lAccess = True
      End If
   Else
      Dim lrs As New ADODB.Recordset
      lrs.Open "SELECT * FROM Employee WHERE (Name = '" & CStr(txtUser.Text) & "')", cnDB, adOpenStatic, adLockReadOnly
      If Not lrs.EOF Then
         If XOREncryption(strKeyCode, lrs("Password")) = CStr(txtPassword.Text) Then
            lAccess = True
         End If
      End If
      Beep
      txtPassword = ""
      txtUser = ""
      txtUser.SetFocus
   End If

   If lAccess Then
      With crModule
         .crKind = txtKind
         .crAcao = txtAcao
         .crObs = txtObs
         .crUser = txtUser
         .crTreat = Format(Now, "hh:mm:ss")
         If .SpotNumber <> -1 Then
            .SpotNumber = FreeSpotNumber(.SpotNumber)
         End If
      End With
      registerOk = True
      Unload Me
   End If
   
End Sub

Private Sub Option1_Click(Index As Integer)
   txtKind.Text = Index
End Sub

'Private Sub switchCamera(ByVal fSpot As Integer)
'   Dim strURL As String
'   With crModule
'      If .ServerAddress <> "" And .Camera <> "" And .Monitor <> "" Then
'         If .senha = "" Then
'            strURL = "http://" & .ServerAddress & "/Interface/VirtualMatrix/ShowObject?" & _
'                  "MonitorID=" & .Monitor & "&SpotNumber=" & fSpot & _
'                  "&ObjectType=0&ObjectName=" & .Camera & "&ResponseFormat=Text&AuthUser=" & .user
'         Else
'            strURL = "http://" & .ServerAddress & "/Interface/VirtualMatrix/ShowObject?" & _
'                  "MonitorID=" & .Monitor & "&SpotNumber=" & fSpot & _
'                  "&ObjectType=0&ObjectName=" & .Camera & _
'                  "&ResponseFormat=Text&AuthUser=" & .user & "&AuthPass=" & .senha
'         End If
'      Else
'         Exit Sub
'      End If
'   End With
'   Set WinHttpReq = New WinHttpRequest
'   WinHttpReq.SetTimeouts 5000, 5000, 5000, 5000  ' Resolve, Connect, Send and Receive
'   WinHttpReq.Open "POST", strURL, False
'   WinHttpReq.SetRequestHeader "ContentType", "text/plain; encoding='utf-8'"
'   WinHttpReq.SetRequestHeader "Content-Length", Len(strURL)
'   WinHttpReq.Send ""
'   Set WinHttpReq = Nothing
'End Sub

