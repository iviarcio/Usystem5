VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form frmRonda 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Relatórios de Ronda"
   ClientHeight    =   3435
   ClientLeft      =   2265
   ClientTop       =   2565
   ClientWidth     =   5850
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Ronda.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3435
   ScaleWidth      =   5850
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdCleanUp 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   45
      Picture         =   "Ronda.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Limpar base de dados de Eventos de Ronda."
      Top             =   2640
      Width           =   735
   End
   Begin VB.CommandButton cmdPrint 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4260
      Picture         =   "Ronda.frx":0884
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Imprimir os Eventos de Ronda de acordo com a Seleção"
      Top             =   2640
      Width           =   735
   End
   Begin VB.CommandButton cmdExit 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5070
      Picture         =   "Ronda.frx":0CC6
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Fechar Controle de Relatórios de Ronda."
      Top             =   2640
      Width           =   735
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   1455
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   5775
      _Version        =   65536
      _ExtentX        =   10186
      _ExtentY        =   2566
      _StockProps     =   15
      BackColor       =   14215660
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderWidth     =   0
      Begin VB.TextBox txtHourEnd 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "HH:mm:ss"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   4
         EndProperty
         Height          =   300
         Left            =   2175
         TabIndex        =   12
         Text            =   "23:59"
         Top             =   930
         Width           =   960
      End
      Begin VB.OptionButton optTipo 
         Caption         =   "Não Executados ou Fora do Previsto"
         ForeColor       =   &H00000000&
         Height          =   540
         Index           =   1
         Left            =   3705
         TabIndex        =   11
         Top             =   772
         Width           =   1965
      End
      Begin VB.OptionButton optTipo 
         Caption         =   "Todos"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   0
         Left            =   3705
         TabIndex        =   10
         Top             =   450
         Value           =   -1  'True
         Width           =   1185
      End
      Begin VB.TextBox txtData 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "HH:mm:ss"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   4
         EndProperty
         Height          =   300
         Left            =   120
         TabIndex        =   8
         Top             =   480
         Width           =   1365
      End
      Begin VB.TextBox txtHourInit 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "HH:mm:ss"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   4
         EndProperty
         Height          =   300
         Left            =   2175
         TabIndex        =   5
         Text            =   "00:00"
         Top             =   495
         Width           =   960
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         Caption         =   "Data"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   210
         TabIndex        =   9
         Top             =   75
         Width           =   1140
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         Caption         =   "Tipos de Eventos"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   3675
         TabIndex        =   7
         Top             =   75
         Width           =   1620
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         Caption         =   "Intervalo de Hora"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   1920
         TabIndex        =   6
         Top             =   75
         Width           =   1605
      End
   End
   Begin Threed.SSPanel SSPanel2 
      Height          =   1035
      Left            =   0
      TabIndex        =   13
      Top             =   1560
      Width           =   5775
      _Version        =   65536
      _ExtentX        =   10186
      _ExtentY        =   1826
      _StockProps     =   15
      BackColor       =   14215660
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderWidth     =   0
      Begin VB.ComboBox cblPonto 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "Ronda.frx":0FD0
         Left            =   3195
         List            =   "Ronda.frx":0FD2
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   420
         Width           =   2460
      End
      Begin VB.ComboBox cblPercurso 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "Ronda.frx":0FD4
         Left            =   195
         List            =   "Ronda.frx":0FD6
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   420
         Width           =   2460
      End
      Begin VB.Label lblPonto 
         Appearance      =   0  'Flat
         Caption         =   "Ponto"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   3255
         TabIndex        =   15
         Top             =   90
         Width           =   1155
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         Caption         =   "Percurso"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   225
         TabIndex        =   14
         Top             =   90
         Width           =   1155
      End
   End
   Begin VB.Image imgCleanUp 
      Height          =   480
      Left            =   855
      Picture         =   "Ronda.frx":0FD8
      Top             =   2790
      Width           =   480
   End
   Begin VB.Label lblCleanUp 
      Appearance      =   0  'Flat
      Caption         =   "obs.: Limpeza de todos os registros até a data especificada (inclusive)!"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   690
      Left            =   1440
      TabIndex        =   3
      Top             =   2715
      Width           =   2010
   End
End
Attribute VB_Name = "frmRonda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim frm As New frmViewReport9

Private Sub cblPercurso_Click()
   If cblPercurso.ListIndex <> -1 And cblPercurso.ListIndex <> 0 Then
      cblPonto.Clear
      cblPonto.AddItem "Todos"
      cblPonto.ItemData(cblPonto.NewIndex) = 0
      Dim tPercurso As clsPercurso
      Set tPercurso = lstPercurso.Item(CStr(cblPercurso.ItemData(cblPercurso.ListIndex)))
      Dim tPonto As clsRonda
      For Each tPonto In tPercurso.lstRonda
         cblPonto.AddItem tPonto.descrRonda
         cblPonto.ItemData(cblPonto.NewIndex) = tPonto.idRonda
      Next
      cblPonto.ListIndex = 0
   Else
      cblPonto.Clear
      cblPonto.AddItem "Todos"
      cblPonto.ItemData(cblPonto.NewIndex) = 0
      cblPonto.ListIndex = 0
   End If
End Sub

Private Sub cmdCleanUp_Click()
   Dim success As Long, lInterval As Long, pastDate As String, lWhere As String
   success = SetWindowPos(frmRonda.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, FLAGS)
   If MsgBox("Confirma a limpeza dos registros de Ronda conforme indicado?", sxQuestion, sxProname) = vbYes Then
      If Not IsDate(txtData) Then
         MsgBox "Data não é válida!", sxExclamation, sxProname
         success = SetWindowPos(frmReport.hWnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
         curhwnd = frmReport.hWnd
         Exit Sub
      End If
      Screen.MousePointer = vbHourglass
      'Note: To pass date value in "#" format to Access, you need to represent
      'this value in english format, i.e., mm/dd/yyyy.
      Dim stDate As String
      stDate = Format$(txtData, "mm/dd/yyyy")
      Dim lcmd As New ADODB.Command
      Set lcmd.ActiveConnection = cnDB
      lcmd.CommandType = adCmdText
      lcmd.CommandText = "DELETE FROM EvtRonda WHERE (Date_Ronda <= #" & stDate & "#)"
      lcmd.Execute
      Screen.MousePointer = vbDefault
      MsgBox "Limpeza de eventos de Ronda executada com sucesso", sxExclamation, sxProname
   End If
   success = SetWindowPos(frmRonda.hWnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
   curhwnd = frmReport.hWnd
End Sub

Private Sub cmdExit_Click()
   Unload Me
End Sub

Private Sub cmdPrint_Click()
   
   'Faz a consistência de Data
   If Not IsDate(txtData) Then
      MsgBox "A Data não é uma data válida!", sxExclamation, sxProname
      Exit Sub
   End If
   
   'Faz a consistência de Hora
   If Not IsDate(txtHourInit) And IsDate(txtHourEnd) Then
      MsgBox "O intervalo não é um intervalo válido!", sxExclamation, sxProname
      Exit Sub
   End If
   
   frm.DataEvt = txtData            'ParameterFields(0)
   frm.intervalo = txtHourInit      'ParameterFields(1)
   frm.Intervalo2 = txtHourEnd      'ParameterFields(2)
      
   'Verifica o Percurso se Todos ou um específico - ParameterFields(3)
   If cblPercurso.ListIndex <> -1 And cblPercurso.ListIndex <> 0 Then
      frm.Percurso = cblPercurso.ItemData(cblPercurso.ListIndex)
   Else
      frm.Percurso = 0
   End If
      
   'Verifica o Ponto se Todos ou um específico - 'ParameterFields(4)
   If cblPonto.ListIndex <> -1 And cblPonto.ListIndex <> 0 Then
      frm.Ponto = cblPonto.ItemData(cblPonto.ListIndex)
   Else
      frm.Ponto = 0
   End If

   If optTipo(0) Then
      'Opção de Todos os Eventos
      frm.SetTipo = g_iRptEvRonda
   Else
      'Não executados ou fora do intervalo
      frm.SetTipo = g_iRptExRonda
      frm.SetSelection = "({EvtRonda.kind_Ronda}=1 or {EvtRonda.kind_Ronda}=2)"
   End If
      
   SetHourGlass Me
   frm.WindowState = vbMaximized
   frm.Show
   ResetMouse Me
   
End Sub
   
Private Sub Form_Activate()
   cmdCleanUp.Visible = (m_tAccess = sxSystem)
   imgCleanUp.Visible = (m_tAccess = sxSystem)
   lblCleanUp.Visible = (m_tAccess = sxSystem)
   cmdCleanUp.Enabled = True
   cblPercurso.Visible = False
   cblPercurso.Clear
   cblPercurso.AddItem "Todos"
   cblPercurso.ItemData(cblPercurso.NewIndex) = 0
   Dim tPercurso As clsPercurso
   For Each tPercurso In lstPercurso
      cblPercurso.AddItem tPercurso.descrPercurso
      cblPercurso.ItemData(cblPercurso.NewIndex) = tPercurso.idPercurso
   Next
   cblPercurso.ListIndex = 0
   cblPercurso.Visible = True
End Sub

Private Sub Form_Load()
   Dim success As Long
   success = SetWindowPos(frmRonda.hWnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
   curhwnd = frmRonda.hWnd
   Left = (Screen.Width - Width) / 2   ' Center form horizontally.
   Top = (Screen.Height - Height) / 2  ' Center form vertically.
   txtData = Format$(Date, "dd/mm/yyyy")
End Sub

