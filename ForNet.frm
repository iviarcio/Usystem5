VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{A4749554-0441-4E64-8A03-3323601631C7}#1.0#0"; "LaVolpeAlphaImg2.ocx"
Begin VB.MDIForm ForNet 
   Appearance      =   0  'Flat
   BackColor       =   &H8000000C&
   ClientHeight    =   12345
   ClientLeft      =   165
   ClientTop       =   810
   ClientWidth     =   19365
   Icon            =   "ForNet.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picPattern 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   150
      Index           =   7
      Left            =   0
      Picture         =   "ForNet.frx":000C
      ScaleHeight     =   150
      ScaleWidth      =   19365
      TabIndex        =   11
      Top             =   150
      Visible         =   0   'False
      Width           =   19365
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BorderStyle     =   0  'None
      Height          =   820
      Left            =   0
      ScaleHeight     =   825
      ScaleWidth      =   19365
      TabIndex        =   10
      Top             =   1200
      Width           =   19365
      Begin LaVolpeAlphaImg.AlphaImgCtl btnAction 
         Height          =   720
         Index           =   16
         Left            =   18480
         ToolTipText     =   "Indicação de Sistema Online"
         Top             =   40
         Width           =   720
         _ExtentX        =   1270
         _ExtentY        =   1270
         Image           =   "ForNet.frx":010E
         Blend           =   192
         Effects         =   "ForNet.frx":0EF0
      End
      Begin LaVolpeAlphaImg.AlphaImgCtl btnAction 
         Height          =   720
         Index           =   14
         Left            =   11880
         ToolTipText     =   "Geração de Eventos Manual"
         Top             =   40
         Width           =   720
         _ExtentX        =   1270
         _ExtentY        =   1270
         Image           =   "ForNet.frx":0F08
         Effects         =   "ForNet.frx":1CF8
      End
      Begin LaVolpeAlphaImg.AlphaImgCtl btnAction 
         Height          =   720
         Index           =   13
         Left            =   11040
         ToolTipText     =   "Trigger de Comunicação"
         Top             =   40
         Width           =   720
         _ExtentX        =   1270
         _ExtentY        =   1270
         Image           =   "ForNet.frx":1D10
         Effects         =   "ForNet.frx":2ECF
      End
      Begin LaVolpeAlphaImg.AlphaImgCtl btnAction 
         Height          =   720
         Index           =   12
         Left            =   10200
         ToolTipText     =   "Informações de Módulos não cadastrados"
         Top             =   40
         Width           =   720
         _ExtentX        =   1270
         _ExtentY        =   1270
         Image           =   "ForNet.frx":2EE7
         Effects         =   "ForNet.frx":41F6
      End
      Begin LaVolpeAlphaImg.AlphaImgCtl btnAction 
         Height          =   720
         Index           =   11
         Left            =   9360
         ToolTipText     =   "Informações de Zonas Inativas"
         Top             =   40
         Width           =   720
         _ExtentX        =   1270
         _ExtentY        =   1270
         Image           =   "ForNet.frx":420E
         Effects         =   "ForNet.frx":4FE5
      End
      Begin LaVolpeAlphaImg.AlphaImgCtl btnAction 
         Height          =   720
         Index           =   10
         Left            =   8520
         ToolTipText     =   "Desligar o som"
         Top             =   40
         Width           =   720
         _ExtentX        =   1270
         _ExtentY        =   1270
         Image           =   "ForNet.frx":4FFD
         Effects         =   "ForNet.frx":653C
      End
      Begin LaVolpeAlphaImg.AlphaImgCtl btnAction 
         Height          =   720
         Index           =   9
         Left            =   7680
         ToolTipText     =   "Desativar indicação de disparo"
         Top             =   40
         Width           =   720
         _ExtentX        =   1270
         _ExtentY        =   1270
         Image           =   "ForNet.frx":6554
         Effects         =   "ForNet.frx":755C
      End
      Begin LaVolpeAlphaImg.AlphaImgCtl btnAction 
         Height          =   720
         Index           =   8
         Left            =   6840
         ToolTipText     =   "Desativar todas as Zonas"
         Top             =   40
         Width           =   720
         _ExtentX        =   1270
         _ExtentY        =   1270
         Image           =   "ForNet.frx":7574
         Effects         =   "ForNet.frx":8302
      End
      Begin LaVolpeAlphaImg.AlphaImgCtl btnAction 
         Height          =   720
         Index           =   7
         Left            =   6000
         ToolTipText     =   "Ativar todas as Zonas"
         Top             =   40
         Width           =   720
         _ExtentX        =   1270
         _ExtentY        =   1270
         Image           =   "ForNet.frx":831A
         Effects         =   "ForNet.frx":907E
      End
      Begin LaVolpeAlphaImg.AlphaImgCtl btnAction 
         Height          =   720
         Index           =   6
         Left            =   5160
         ToolTipText     =   "Relatório de Locais Fechados"
         Top             =   40
         Width           =   720
         _ExtentX        =   1270
         _ExtentY        =   1270
         Image           =   "ForNet.frx":9096
         Effects         =   "ForNet.frx":9C23
      End
      Begin LaVolpeAlphaImg.AlphaImgCtl btnAction 
         Height          =   720
         Index           =   5
         Left            =   4320
         ToolTipText     =   "Relatório de Locais Abertos"
         Top             =   40
         Width           =   720
         _ExtentX        =   1270
         _ExtentY        =   1270
         Image           =   "ForNet.frx":9C3B
         Effects         =   "ForNet.frx":A7E3
      End
      Begin LaVolpeAlphaImg.AlphaImgCtl btnAction 
         Height          =   720
         Index           =   4
         Left            =   3480
         ToolTipText     =   "Cadastro de Rondas"
         Top             =   40
         Width           =   720
         _ExtentX        =   1270
         _ExtentY        =   1270
         Image           =   "ForNet.frx":A7FB
         Effects         =   "ForNet.frx":BE67
      End
      Begin LaVolpeAlphaImg.AlphaImgCtl btnAction 
         Height          =   720
         Index           =   15
         Left            =   12720
         ToolTipText     =   "Sair do Sistema"
         Top             =   40
         Width           =   720
         _ExtentX        =   1270
         _ExtentY        =   1270
         Image           =   "ForNet.frx":BE7F
         Effects         =   "ForNet.frx":CB84
      End
      Begin LaVolpeAlphaImg.AlphaImgCtl btnAction 
         Height          =   720
         Index           =   3
         Left            =   2640
         ToolTipText     =   "Últimos Eventos"
         Top             =   40
         Width           =   720
         _ExtentX        =   1270
         _ExtentY        =   1270
         Image           =   "ForNet.frx":CB9C
         Effects         =   "ForNet.frx":DB22
      End
      Begin LaVolpeAlphaImg.AlphaImgCtl btnAction 
         Height          =   720
         Index           =   2
         Left            =   1800
         ToolTipText     =   "Status dos Locais"
         Top             =   40
         Width           =   720
         _ExtentX        =   1270
         _ExtentY        =   1270
         Image           =   "ForNet.frx":DB3A
         Effects         =   "ForNet.frx":ECD6
      End
      Begin LaVolpeAlphaImg.AlphaImgCtl btnAction 
         Height          =   720
         Index           =   1
         Left            =   960
         ToolTipText     =   "Cadastro de Locais"
         Top             =   40
         Width           =   720
         _ExtentX        =   1270
         _ExtentY        =   1270
         Image           =   "ForNet.frx":ECEE
         Effects         =   "ForNet.frx":FE40
      End
      Begin LaVolpeAlphaImg.AlphaImgCtl btnAction 
         Height          =   720
         Index           =   0
         Left            =   120
         ToolTipText     =   "Troca de Operador"
         Top             =   40
         Width           =   720
         _ExtentX        =   1270
         _ExtentY        =   1270
         Image           =   "ForNet.frx":FE58
         Effects         =   "ForNet.frx":10E7A
      End
   End
   Begin VB.Timer trmAquire 
      Enabled         =   0   'False
      Index           =   3
      Interval        =   100
      Left            =   9720
      Top             =   4680
   End
   Begin VB.Timer trmAquire 
      Enabled         =   0   'False
      Index           =   2
      Interval        =   100
      Left            =   9120
      Top             =   4680
   End
   Begin VB.Timer trmAquire 
      Enabled         =   0   'False
      Index           =   1
      Interval        =   100
      Left            =   8520
      Top             =   4680
   End
   Begin VB.Timer trmSecurity 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   4680
      Top             =   5160
   End
   Begin VB.Timer trmLastEvents 
      Interval        =   700
      Left            =   3240
      Top             =   5160
   End
   Begin VB.Timer trmPanico 
      Enabled         =   0   'False
      Interval        =   350
      Left            =   3720
      Top             =   5160
   End
   Begin Threed.SSPanel pctPiso 
      Align           =   3  'Align Left
      Height          =   9810
      Left            =   0
      TabIndex        =   8
      Top             =   2025
      Width           =   1605
      _Version        =   65536
      _ExtentX        =   2831
      _ExtentY        =   17304
      _StockProps     =   15
      BackColor       =   13160660
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.Label lblPiso 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pavimento Terreo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C000&
         Height          =   480
         Index           =   0
         Left            =   -30
         TabIndex        =   9
         Top             =   0
         Width           =   1515
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Timer trmInativos 
      Interval        =   30000
      Left            =   2760
      Top             =   5160
   End
   Begin VB.Timer trmAquire 
      Enabled         =   0   'False
      Index           =   0
      Interval        =   100
      Left            =   7920
      Top             =   4680
   End
   Begin VB.Timer trmService 
      Interval        =   1000
      Left            =   2280
      Top             =   5160
   End
   Begin VB.PictureBox picPattern 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   150
      Index           =   6
      Left            =   0
      Picture         =   "ForNet.frx":10E92
      ScaleHeight     =   150
      ScaleWidth      =   19365
      TabIndex        =   7
      Top             =   300
      Visible         =   0   'False
      Width           =   19365
   End
   Begin VB.PictureBox picPattern 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   150
      Index           =   5
      Left            =   0
      Picture         =   "ForNet.frx":10F94
      ScaleHeight     =   150
      ScaleWidth      =   19365
      TabIndex        =   6
      Top             =   450
      Visible         =   0   'False
      Width           =   19365
   End
   Begin VB.PictureBox picPattern 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   150
      Index           =   4
      Left            =   0
      Picture         =   "ForNet.frx":11096
      ScaleHeight     =   150
      ScaleWidth      =   19365
      TabIndex        =   5
      Top             =   600
      Visible         =   0   'False
      Width           =   19365
   End
   Begin VB.PictureBox picPattern 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   150
      Index           =   3
      Left            =   0
      Picture         =   "ForNet.frx":11198
      ScaleHeight     =   150
      ScaleWidth      =   19365
      TabIndex        =   4
      Top             =   750
      Visible         =   0   'False
      Width           =   19365
   End
   Begin VB.PictureBox picPattern 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   150
      Index           =   2
      Left            =   0
      Picture         =   "ForNet.frx":1129A
      ScaleHeight     =   150
      ScaleWidth      =   19365
      TabIndex        =   3
      Top             =   900
      Visible         =   0   'False
      Width           =   19365
   End
   Begin VB.PictureBox picPattern 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   150
      Index           =   1
      Left            =   0
      Picture         =   "ForNet.frx":1139C
      ScaleHeight     =   150
      ScaleWidth      =   19365
      TabIndex        =   2
      Top             =   1050
      Visible         =   0   'False
      Width           =   19365
   End
   Begin VB.PictureBox picPattern 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   150
      Index           =   0
      Left            =   0
      Picture         =   "ForNet.frx":1149E
      ScaleHeight     =   150
      ScaleWidth      =   19365
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   19365
   End
   Begin VB.Timer trmRefresh 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   4200
      Top             =   5160
   End
   Begin MSCommLib.MSComm CommA 
      Index           =   0
      Left            =   7920
      Tag             =   "0"
      Top             =   5160
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      CommPort        =   3
      DTREnable       =   0   'False
      BaudRate        =   19200
   End
   Begin MSComDlg.CommonDialog cdl 
      Left            =   3600
      Top             =   4365
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   510
      Left            =   0
      TabIndex        =   0
      Top             =   11835
      Width           =   19365
      _ExtentX        =   34158
      _ExtentY        =   900
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   5115
            MinWidth        =   5115
            Picture         =   "ForNet.frx":115A0
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Bevel           =   0
            Object.Width           =   22410
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1773
            MinWidth        =   1764
            Picture         =   "ForNet.frx":118BA
            Text            =   "0"
            TextSave        =   "0"
            Object.ToolTipText     =   "Lista de Eventos Pendentes"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Object.Width           =   1411
            MinWidth        =   1411
            TextSave        =   "16:43"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   2822
            MinWidth        =   2822
            Picture         =   "ForNet.frx":11D0C
            Object.ToolTipText     =   "Desliga o som "
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSCommLib.MSComm CommA 
      Index           =   1
      Left            =   8520
      Tag             =   "0"
      Top             =   5160
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      CommPort        =   3
      DTREnable       =   0   'False
      BaudRate        =   19200
   End
   Begin MSCommLib.MSComm CommA 
      Index           =   2
      Left            =   9120
      Tag             =   "0"
      Top             =   5160
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      CommPort        =   3
      DTREnable       =   0   'False
      BaudRate        =   19200
   End
   Begin MSCommLib.MSComm CommA 
      Index           =   3
      Left            =   9720
      Tag             =   "0"
      Top             =   5160
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      CommPort        =   3
      DTREnable       =   0   'False
      BaudRate        =   19200
   End
   Begin VB.Menu mnuOption 
      Caption         =   "&Opções"
      Begin VB.Menu mnuOperador 
         Caption         =   "Trocar de Operador"
      End
      Begin VB.Menu mnuPisos 
         Caption         =   "Seleção dos Pisos"
         Begin VB.Menu mnuLeft 
            Caption         =   "Ajustar à esquerda"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuBottom 
            Caption         =   "Ajustar ao rodapé"
         End
      End
      Begin VB.Menu mnuReport 
         Caption         =   "Programação de Abertura/Fechamento"
      End
      Begin VB.Menu mnuSeparator3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Sair"
         Shortcut        =   ^F
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Ferramentas"
      Begin VB.Menu mnuPisoAdd 
         Caption         =   "Adicionar Piso/Andar"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuPlantaAdd 
         Caption         =   "Adicionar/Alterar Planta"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuPisoRemove 
         Caption         =   "Remover Piso/Andar"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuSeparator1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuModeDesign 
         Caption         =   "&Design"
      End
      Begin VB.Menu mnuEntityCreate 
         Caption         =   "Criar Entidades"
         Enabled         =   0   'False
         Shortcut        =   ^E
      End
      Begin VB.Menu mnuSeparator2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBackup 
         Caption         =   "Backup && Restore"
      End
      Begin VB.Menu mnuSep19 
         Caption         =   "-"
      End
      Begin VB.Menu mnuConfig 
         Caption         =   "Comunicação"
      End
      Begin VB.Menu mnuSeparator4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCOperador 
         Caption         =   "Cadastro de Operadores"
      End
   End
   Begin VB.Menu mnuCadastro 
      Caption         =   "&Cadastros"
      Begin VB.Menu mnuCGrupos 
         Caption         =   "de Grupos"
      End
      Begin VB.Menu mnuSep15 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCServico 
         Caption         =   "Serviços (não disponível)"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuCRonda 
         Caption         =   "Rondas (não disponível)"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu mnuInform 
      Caption         =   "&Visualizar"
      Begin VB.Menu mnuLastEvents 
         Caption         =   "Últimos Eventos"
      End
      Begin VB.Menu mnuCritico 
         Caption         =   "Eventos Críticos"
      End
      Begin VB.Menu mnuSep8 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRLocOpen 
         Caption         =   "Locais/Lojas Abertas"
      End
      Begin VB.Menu mnuRLClose 
         Caption         =   "Locais/Lojas Fechadas"
      End
      Begin VB.Menu mnuSep11 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRStatus 
         Caption         =   "Status"
         Begin VB.Menu mnuSLocais 
            Caption         =   "Locais/Lojas"
         End
         Begin VB.Menu mnuIZonas 
            Caption         =   "Sensores"
            Begin VB.Menu mnuIZIncendio 
               Caption         =   "Incêndio"
            End
            Begin VB.Menu mnuIZIntrusao 
               Caption         =   "Intrusão"
            End
            Begin VB.Menu mnuIZEmergencia 
               Caption         =   "Emergência"
            End
            Begin VB.Menu mnuIZPanico 
               Caption         =   "Pânico"
            End
            Begin VB.Menu mnuIZSistema 
               Caption         =   "Sistema"
            End
         End
      End
      Begin VB.Menu mnuSep10 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRInativos 
         Caption         =   "Sensores/Receptores Inativos"
      End
      Begin VB.Menu mnuBaseStatus 
         Caption         =   "Módulos não Cadastrados"
      End
   End
   Begin VB.Menu mnuRelatorios 
      Caption         =   "&Relatórios"
      Begin VB.Menu mnuRPTGeral 
         Caption         =   "Gerais"
      End
      Begin VB.Menu mnuRCadLocais 
         Caption         =   "de Lojas Cadastradas"
      End
      Begin VB.Menu mnuRCadZonas 
         Caption         =   "de Sensores Cadastrados"
      End
   End
   Begin VB.Menu mnuMonitorActivate 
      Caption         =   "&Ativação"
      Begin VB.Menu mnuAZIncendio 
         Caption         =   "Incêndio"
      End
      Begin VB.Menu mnuAZIntrusao 
         Caption         =   "Intrusão"
      End
      Begin VB.Menu mnuAZEmergencia 
         Caption         =   "Emergência"
      End
      Begin VB.Menu mnuAZPanico 
         Caption         =   "Pânico"
      End
      Begin VB.Menu mnuAZSistema 
         Caption         =   "Sistema"
      End
   End
   Begin VB.Menu mnuMonitorDeactivate 
      Caption         =   "&Desativação"
      Begin VB.Menu mnuDZIncendio 
         Caption         =   "Incêndio"
      End
      Begin VB.Menu mnuDZIntrusao 
         Caption         =   "Intrusão"
      End
      Begin VB.Menu mnuDZEmergencia 
         Caption         =   "Emergência"
      End
      Begin VB.Menu mnuDZPanico 
         Caption         =   "Pânico"
      End
      Begin VB.Menu mnuDZSistema 
         Caption         =   "Sistema"
      End
   End
   Begin VB.Menu mnuAjuda 
      Caption         =   "Ajuda"
      Begin VB.Menu mnuHelp 
         Caption         =   "Tópicos de Ajuda"
         Shortcut        =   ^H
      End
      Begin VB.Menu mnuSeparator5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "Sobre a Rede USystemEco"
      End
   End
   Begin VB.Menu mnuEntity 
      Caption         =   "&Entidade/Loja"
      Visible         =   0   'False
      Begin VB.Menu mnuEntityProperties 
         Caption         =   "Propriedades"
      End
      Begin VB.Menu mnuLEvent 
         Caption         =   "Últimos Eventos"
      End
      Begin VB.Menu mnuSeparator6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuZConfig 
         Caption         =   "Configuração de Sensores"
      End
      Begin VB.Menu mnuEntityDelete 
         Caption         =   "Remover"
      End
      Begin VB.Menu mnuSeparator7 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEntityActivate 
         Caption         =   "Ativar Sensores"
         Begin VB.Menu mnuAEIncendio 
            Caption         =   "Incêndio"
         End
         Begin VB.Menu mnuAEIntrusao 
            Caption         =   "Intrusão"
         End
         Begin VB.Menu mnuAEEmergencia 
            Caption         =   "Emergência"
         End
         Begin VB.Menu mnuAEPanico 
            Caption         =   "Pânico"
         End
         Begin VB.Menu mnuAESistema 
            Caption         =   "Sistema"
         End
      End
      Begin VB.Menu mnuEntityDeactivate 
         Caption         =   "Desativar Sensores"
         Begin VB.Menu mnuDEIncêndio 
            Caption         =   "Incêndio"
         End
         Begin VB.Menu mnuDEIntrusao 
            Caption         =   "Intrusão"
         End
         Begin VB.Menu mnuDEEmergencia 
            Caption         =   "Emergência"
         End
         Begin VB.Menu mnuDEPanico 
            Caption         =   "Pânico"
         End
         Begin VB.Menu mnuDESistema 
            Caption         =   "Sistema"
         End
      End
      Begin VB.Menu mnuSeparator9 
         Caption         =   "-"
      End
      Begin VB.Menu mnuZStatus 
         Caption         =   "Status dos Sensores"
      End
      Begin VB.Menu mnuReset 
         Caption         =   "Reset de Sensores Wired"
      End
      Begin VB.Menu mnuSeparator8 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMonitor 
         Caption         =   "Monitoração"
         Begin VB.Menu mnuMDesativo 
            Caption         =   "Sensores Desativados"
         End
         Begin VB.Menu mnuMAtivo 
            Caption         =   "Sensores Ativados"
         End
      End
   End
End
Attribute VB_Name = "ForNet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Portas de Comunicação e Configurações das mesmas
Private fPort(0 To 3) As Integer
Private fSett(0 To 3) As String
Private fComm(0 To 3) As Boolean

'Controle de Acesso aos Pisos
Private cLblStep As Integer

'Controle de Display
Public noDisp As Boolean

'Os buffers de Comunicação devem ser preservados (para recuperar o eventual truncamento de mensagens).
Private CommA_Buffer(0 To 3) As String

Private segCounter As Byte

Private Sub SetAppearence(btn As AlphaImgCtl, flag As Boolean)
   If flag Then
      btn.GrayScale = lvicNoGrayScale
   Else
      btn.GrayScale = lvicGreenMask
   End If
   btn.Enabled = flag
End Sub

Private Sub btnAction_Click(Index As Integer)
   Select Case Index
      Case 0
         mnuOperador_Click
      Case 1
         mnuRCadLocais_Click
      Case 2
         mnuSLocais_Click
      Case 3
         mnuLastEvents_Click
      Case 4
         mnuCRonda_Click
      Case 5
         mnuRLocOpen_Click
      Case 6
         mnuRLClose_Click
      Case 7
         MonitorActivate s_All, True
      Case 8
         MonitorDeactivate s_All, True
      Case 9
         Disparo_Clear
      Case 10
         noDisp = False
         Sound_Update fmode:=sxNoSound, isCritico:=False, fNoSound:=False
         DoEvents
         Remove_Display
         Unload frmPanico
         On Error GoTo 0
      Case 11
         mnuRInativos_Click
      Case 12
         mnuBaseStatus_Click
      Case 13
         Load frmComm
         frmComm.Show
         frmComm.ZOrder 0
      Case 14
         SimulaEventos
         'Load frmPanico
      Case 15
         mnuExit_Click
   End Select
End Sub

Private Sub btnAction_MouseEnter(Index As Integer)
   If btnAction(Index).Enabled Then
      btnAction(Index).SetRedraw = False
      btnAction(Index).GrayScale = lvicSepia
      btnAction(Index).LightnessPct = -20
      btnAction(Index).SetRedraw = True
   End If
End Sub

Private Sub btnAction_MouseExit(Index As Integer)
   If btnAction(Index).Enabled Then
      btnAction(Index).SetRedraw = False
      btnAction(Index).GrayScale = lvicNoGrayScale
      btnAction(Index).LightnessPct = 0
      btnAction(Index).SetRedraw = True
   End If
End Sub

'Rotina chamada quando algum conteúdo é detectado no buffer de comunicação da porta serial mapeada na CommA
Private Sub CommA_OnComm(Index As Integer)

    Select Case CommA(Index).CommEvent
      
        Case comTxFull
            MsgBox "Buffer de transmissão cheio na Com" & CommA(Index).commPort & NL & CONTATO, vbOKOnly + vbInformation, USVersion
            
        Case comRxOver
            MsgBox "Estouro do Buffer de recepção da Com" & CommA(Index).commPort & NL & CONTATO, vbOKOnly + vbInformation, USVersion
            
        Case comRxParity
            MsgBox "Está ocorrendo erro de Paridade na Com" & CommA(Index).commPort & NL & CONTATO, vbOKOnly + vbInformation, USVersion
                        
        Case comEvReceive
            On Error Resume Next
            trmAquire(Index).Enabled = True
            Do
                DoEvents
                CommA_Buffer(Index) = CommA_Buffer(Index) & CommA(Index).Input
            Loop Until trmAquire(Index).Enabled = False
            'Chama a rotina de recebimento de mensagem para consumir CommA_Buffer(Index)
            Buffer_Received (Index)
            On Error GoTo 0
            
    End Select

End Sub

'Rotina que recebe o Buffer e trata todos os strings, mesmo que concatenados.
Private Sub Buffer_Received(Index As Integer)
        
    Dim Size As Integer
    Dim Sizef As Integer
    Dim CheckSum As Boolean
    Dim localBuffer As String
    Dim DadosHex As String
    Dim DadosComm As String
    Dim Duplicidade As Boolean
                  
    While CommA_Buffer(Index) <> ""
    
        'Lê o tamanho da primeira mensagem (segundo byte) do String, sem CheckSum
        If Len(CommA_Buffer(Index)) >= 2 Then
            Size = Asc(Mid(CommA_Buffer(Index), 2, 1))
            'Tamanho da mensagem com CheckSum
            Sizef = Size + 1
        Else
            Exit Sub
        End If
        
        'Verfica se o tamanho lido na mensagem é menor ou igual ao string.
        'Sai se tamanho lido for superior ao comprimento do Buffer (mensagem truncada?)
        If Len(CommA_Buffer(Index)) < Sizef Then Exit Sub
        
        'Separa a primeira mensagem
        localBuffer = Left(CommA_Buffer(Index), Sizef)
        'Remove a mensagem lida do Buffer
        CommA_Buffer(Index) = Right(CommA_Buffer(Index), Len(CommA_Buffer(Index)) - Size - 1)
        
        'Converte a mensagem de ASC para HEX
        DadosHex = Char_to_Hex(localBuffer, Sizef)
        'Verifica se o CheckSum está correto para entrar na rotina de tratamento
        CheckSum = Verifica_CheckSum(localBuffer)
        
        'Enfilera (FIFO) a mensagem se já não esta na Queue. Se estiver descarta como duplicada.
        If CheckSum Then
            Duplicidade = Not tQueue.Enqueue(DadosHex, CommA(Index).commPort)
            'Verifica se mostra a comunicação na tela frmComm
            If m_bShowComm And Not Duplicidade Then
                'Formata a mensagem de acordo com o tipo do evento (Serial Receiver, Repeater ou Sensor)
                DadosComm = "[Com" & CommA(Index).commPort & "] " & Formata_Mensagem(DadosHex, CheckSum)
                'Insere a mensagem na tela de comunicação
                If DadosHex <> "" Then
                    If Len(DadosComm) > 80 Then
                        frmComm.List1.AddItem Left(DadosComm, 79)
                        frmComm.List1.AddItem Right(DadosComm, Len(DadosComm) - 79)
                    Else
                        frmComm.List1.AddItem DadosComm
                    End If
                    frmComm.List1.AddItem ""
                End If
            End If
        End If
        
    Wend
    
End Sub


Private Sub lblPiso_Click(Index As Integer)
   lblPiso(m_iCurPiso).ForeColor = lblColor.cBlack
   lblPiso(Index).ForeColor = lblColor.cBlue
   m_iCurPiso = Index
   Set tpiso = lstPiso.Item(CStr(m_iCurPiso))
   tpiso.f.Show
   tpiso.f.WindowState = vbMaximized
End Sub

'Carrega o Form principal
Private Sub MDIForm_Load()

   m_iCurPiso = 0
   m_bDesignMode = False
   m_bUserUnload = True
   m_tAccess = sxOperator
   DoEvents
   
   'Fecha a tela de splash
   Unload frmSplash
   
   Screen.MousePointer = vbDefault
   LogOn_Display False              'Modifica m_tAccess, m_bShutDown
   If m_bShutDown Then
      Unload Me
      End
   End If
   
   Dim hMenu As Long
   hMenu = GetSystemMenu(ForNet.hWnd, False)
   DeleteMenu hMenu, SC_MAXIMIZE, MF_BYCOMMAND
   DeleteMenu hMenu, SC_MINIMIZE, MF_BYCOMMAND
   DeleteMenu hMenu, SC_SIZE, MF_BYCOMMAND
   DeleteMenu hMenu, SC_MOVE, MF_BYCOMMAND
   DeleteMenu hMenu, SC_RESTORE, MF_BYCOMMAND
   DeleteMenu hMenu, SC_NEXTWINDOW, MF_BYCOMMAND
   DeleteMenu hMenu, SC_CLOSE, MF_BYCOMMAND
   DeleteMenu hMenu, 0, MF_BYPOSITION

   Dim lStyle As Long
   lStyle = GetWindowLong(ForNet.hWnd, GWL_STYLE)
   lStyle = lStyle And (Not WS_MAXIMIZEBOX)
   lStyle = lStyle And (Not WS_MINIMIZEBOX)
   SetWindowLong ForNet.hWnd, GWL_STYLE, lStyle
   
   Dim Thwnd As Long
   Thwnd = FindWindow("Shell_traywnd", "")
   Call SetWindowPos(Thwnd, 0, 0, 0, 0, 0, SWP_HIDEWINDOW)
   ForNet.Height = Screen.Height
   ForNet.Width = Screen.Width
   ForNet.Left = 0
   ForNet.Top = 0
   
   StatusBar1.Panels.Item(1).Text = strAccess(m_tAccess) & m_sUser
   
   'Ajusta Caption da tela principal
   Set_Caption
   
   'Cria os padrões de preenchimento.
   Dim i As Integer
   For i = 0 To 7
      lngFill(i) = CreatePatternBrush(picPattern(i).Picture)
   Next i
   
   ' Ajusta a posição da botão online
   btnAction(16).Left = Picture1.Width - 760
   
   'Ajusta a posição do menu (Superior ou Lateral)
   mnuLeft.Checked = m_bPisoLeft
   mnuBottom.Checked = Not m_bPisoLeft
   
   'Ajusta o setup de posicionamento dos pisos em relação a Menu
   Pisos_Setup
   
   'Carrega as plantas de todos os pisos
   Load_Pisos
   
'  Configura as portas de comunicação COMM
   Comm_Setup
   
   'Carrega o formulário de eventos críticos
   Load frmQueue
   
   'Ativa o timer de redesenho das entidades
   trmRefresh.Enabled = True
   
   'Ativa o timer de varredura da tabela de Eventos (tQueue)
   trmLastEvents.Enabled = True
   
   'Ativa o timer de varredura da segurança
   trmSecurity.Enabled = True
      
End Sub

Public Sub Load_Pisos()
   'Cria referência ao form que irá conter a primeira planta
   Dim firstf As frmPlanta
   Set firstf = Nothing
   
   'Carrega todas as plantas
   Dim f As frmPlanta
   Dim rsFloor As New ADODB.Recordset
   
   rsFloor.Open "SELECT * FROM Floor ORDER BY cp_Floor ASC", cnDB, adOpenStatic, adLockReadOnly
   
   While Not rsFloor.EOF
      m_iCurPiso = rsFloor("cp_Floor")
      Set tpiso = New clsPiso
      lstPiso.Add Item:=tpiso, Key:=CStr(m_iCurPiso)
      tpiso.n = m_iCurPiso
      tpiso.rCaption = rsFloor("Descr_Floor")
      tpiso.rStep = m_iLastTop
      m_iLastTop = m_iLastTop + cLblStep
      Set f = New frmPlanta
      Set tpiso.f = f
      Load f
      f.curPiso = m_iCurPiso
      f.Caption = tpiso.rCaption
      Set f.Picture = LoadPicture(App.Path & "\Pisos\" & rsFloor("Picture_Floor"))
      If firstf Is Nothing Then
         Set firstf = f
      End If
      
      Load lblPiso(m_iCurPiso)
      If m_bPisoLeft Then
         lblPiso(m_iCurPiso).Top = tpiso.rStep
      Else
         lblPiso(m_iCurPiso).Left = tpiso.rStep
      End If
      lblPiso(m_iCurPiso).ForeColor = lblColor.cBlack
      lblPiso(m_iCurPiso).Caption = tpiso.rCaption
      lblPiso(m_iCurPiso).Visible = True
      rsFloor.MoveNext
   
   Wend
   
   Set tpiso = Nothing
   pctPiso.Visible = lstPiso.Count > 1
   
   If Not firstf Is Nothing Then
      m_iCurPiso = firstf.curPiso
      lblPiso(m_iCurPiso).ForeColor = lblColor.cBlue
      firstf.Show
   End If

   Status_Menu m_bDesignMode
   
End Sub

Public Sub Comm_Setup()
   'Busca a configuração atual
   Dim lds As New ADODB.Recordset
   lds.Open "Select * From Config", cnDB, adOpenStatic, adLockReadOnly
   Dim i As Integer
   For i = 0 To 3
      fComm(i) = lds("Enabled")
      fPort(i) = lds("Comm")
      fSett(i) = m_sBaud(lds("BaudRate")) & np & m_sParity(lds("Parity")) & np & m_sData(lds("DataBits")) & np & m_sStop(lds("StopBits"))
      lds.MoveNext
   Next i
   lds.Close
   m_bChange = False
'  Chama a rotina de inicialização da portas de comunicação
   Comm_Init
End Sub

Private Sub Comm_Init()
    
    On Error Resume Next
    Dim i As Integer
    For i = 0 To 3
      CommA(i).PortOpen = False
      m_bCommStatus = False
      If fComm(i) Then
         ' Select Comm Port
         CommA(i).commPort = fPort(i)
         ' baud, parity, data, and stop bit.
         CommA(i).Settings = fSett(i)
         ' Data Terminal Ready line is set to high (on) when the port is opened,
         CommA(i).DTREnable = True
         ' Tell the control to read the entire buffer when Input is used.
         CommA(i).InputLen = 0
         ' Tell the control we will use the iterrupt method (against polling)
         CommA(i).RThreshold = 1
         ' Open the port.
         CommA(i).PortOpen = True
         m_bCommStatus = True
      End If
   Next i
   On Error GoTo 0
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    Dim Thwnd As Long
    Thwnd = FindWindow("Shell_traywnd", "")
    Call SetWindowPos(Thwnd, 0, 0, 0, 0, 0, SWP_SHOWWINDOW)
End Sub

Private Sub mnuBackup_Click()
   frmService.Show
End Sub

Private Sub mnuBaseStatus_Click()
   frmDevices.Show
End Sub

Private Sub SimulaEventos()
   Dim ct As Integer
   Dim numEvents As Integer
   Dim lok As Boolean
   Dim DadosHex As String
   Dim DadosComm As String
   Dim Sensor(4) As String
   Dim i As Integer
   Dim j As Integer
   Dim loopControl As Boolean
   
   ct = 4 'default
   On Error Resume Next
   ct = InputBox("case number: ", "0..4=Receiver, 5..9=Repeater, 10..16=Sensor")
   numEvents = InputBox("Número de Eventos:", "Número de Eventos (default=1)", 1)
   
   Sensor(0) = InputBox("Número do Sensor:", "Alarme", "B28BCBA9")
   
   i = 0
     For j = 1 To numEvents
         
      Select Case ct
         'Serial Receiver
         Case 0:
             DadosHex = "11040080" 'Ruido excessivo + Status. -80 + 90 = Ruido
         Case 1:
            DadosHex = "11040020" 'Alarme de Tamper + Status. -20+30 = Tamper
         Case 2:
            DadosHex = "11040008" 'Ocorreu Reset + Status. -08 + 18 = Reset
         Case 3:
            DadosHex = "11040001" 'Falha no Link + status. -01 + 11 = link
            ct = 4
         Case 4:
            DadosHex = "11040000" 'Status ok.
            ct = 3
         
         'Repeater
         Case 5:
            DadosHex = "1330015F13FC00FCAD4A0000802626"  'Ruido excessivo. -80+90= Ruido
         Case 6:
            DadosHex = "1330015F13FC00FCAD4A0000402626"  'Bateria Fraca. -40+50= Bateria
         Case 7:
            DadosHex = "1330015F13FC00FCAD4A0000202626"  'Tamper. -20+30= Tamper
         Case 8:
            DadosHex = "1330015F13FC00FCAD4A0000022626"  'Perda Rede AC. -02+12= Rede
            ct = 9
         Case 9:
            DadosHex = "1330015F13FC00FCAD4A0000002626"  'Status ok.
            ct = 8
            
         'Sensores
         Case 10:
            DadosHex = "1330" + Sensor(i) + "00FCAD4A2880002626"  'Alarme de Vandalismo.
         Case 11:
            DadosHex = "1330" + Sensor(i) + "00FCAD4A2800202626"  'Tamper. -20+30 = Tamper
         Case 12:
            DadosHex = "1330" + Sensor(i) + "00FCAD4A2800402626"  'Bateria. -40+50= Bateria
         Case 13:
            DadosHex = "1330" + Sensor(i) + "00FCAD4A2800082626"  'Reset. -08+18= Reset
         Case 14:
            DadosHex = "1330" + Sensor(i) + "00FCAD4A2801003547"  'Alarme 1
            ct = 16
         Case 15:
            DadosHex = "1330" + Sensor(i) + "00FCAD4A2802002226"  'Alarme 2.
         Case 16:
            DadosHex = "1330" + Sensor(i) + "00FCAD4A2800004726"  'Status Ok.
            ct = 14
      End Select
      
      lok = tQueue.Enqueue(DadosHex, 1)
      DoEvents
      If m_bShowComm Then
         'Formata a mensagem de acordo com o tipo do evento (Serial Receiver, Repeater ou Sensor)
         DadosComm = Formata_Mensagem(DadosHex, True)
         'Insere a mensagem na tela de comunicação
         If DadosHex <> "" Then
            If Len(DadosComm) > 96 Then
               frmComm.List1.AddItem Left(DadosComm, 95)
               frmComm.List1.AddItem Right(DadosComm, Len(DadosComm) - 95)
            Else
               frmComm.List1.AddItem DadosComm
            End If
            frmComm.List1.AddItem ""
         End If
      End If
      DoEvents
      Sleep 4000 ' to sleep for 4 seconds
      DoEvents
     Next j
End Sub

Private Sub mnuBottom_Click()
   mnuLeft.Checked = False
   mnuBottom.Checked = True
   m_bPisoLeft = False
   Pisos_Setup
End Sub

Private Sub mnuCGrupos_Click()
    MsgBox "O cadastro de Grupos está em desenvolvimento.", vbInformation + vbOKOnly
End Sub

Private Sub mnuCritico_Click()
   Screen.MousePointer = vbHourglass
   DoEvents
   Dim frm As New frmViewReport9
   frm.SetTipo = g_iRptCritico
   frm.WindowState = vbMaximized
   frm.Show
   Screen.MousePointer = vbDefault
End Sub

Private Sub mnuCRonda_Click()
   MsgBox "Ronda não está disponível nesta versão do USystem", vbInformation + vbOKOnly
   'frmCRonda.Show
End Sub

Private Sub mnuCServico_Click()
   MsgBox "Serviços não está disponível nesta versão do USystem", vbInformation + vbOKOnly
End Sub

Private Sub mnuLeft_Click()
   mnuLeft.Checked = True
   mnuBottom.Checked = False
   m_bPisoLeft = True
   Pisos_Setup
End Sub

Private Sub mnuRCadLocais_Click()
   frmCadastro.Show
End Sub

Private Sub mnuRCadZonas_Click()
   frmCZonas.Show
End Sub

Private Sub mnuReport_Click()
   frmProgR.Show
End Sub

Private Sub mnuRInativos_Click()
   frmInativos.Show
End Sub

Private Sub mnuRLClose_Click()
   Screen.MousePointer = vbHourglass
   Dump_Lojas fOpen:=False
   DoEvents
   Dim frm As New frmViewReport9
   frm.SetTipo = g_iRptLFechados
   frm.WindowState = vbMaximized
   frm.Show
   Screen.MousePointer = vbDefault
End Sub

Private Sub mnuRLocOpen_Click()
   Screen.MousePointer = vbHourglass
   Dump_Lojas fOpen:=True
   DoEvents
   Dim frm As New frmViewReport9
   frm.SetTipo = g_iRptLAbertos
   frm.WindowState = vbMaximized
   frm.Show
   Screen.MousePointer = vbDefault
End Sub

Private Sub mnuRonda_Click()
   frmRonda.Show
End Sub

Private Sub mnuRPTGeral_Click()
   frmReport.Show
End Sub

Private Sub mnuOperador_Click()
   LogOn_Display True
End Sub

Private Sub LogOn_Display(ByVal fchg As Boolean)
   m_bPermition = False
   m_bShutDown = False
   m_Debug = False
   If fchg Then
      Load LogOn
      LogOn.Caption = "Troca de Operador"
      Update_Display "Mudança de Operador", sxImgInform, False
   Else
      Update_Display "Acesso ao USystem5", sxImgInform, False
   End If
   LogOn.Show vbModal
   If m_bPermition Then
      If fchg Then
         Make_Service "Mudança de Operador", strAccess(m_tAccess) & m_sUser
      Else
         Make_Service "Acesso ao USystem5", strAccess(m_tAccess) & m_sUser
      End If
      Update_Display "Benvindo à rede de Monitoramento USystem", sxImgInform, False
      StatusBar1.Panels.Item(1).Text = strAccess(m_tAccess) & m_sUser
      SetStatus m_tAccess
   ElseIf fchg Then
      m_bPermition = True
   ElseIf m_bShutDown Then
      If fchg Then Make_Service "Desativação da rede USystem", strAccess(m_tAccess) & m_sUser
      cnDB.Close: Set cnDB = Nothing
      m_bUserUnload = False
   Else
      m_bShutDown = True
   End If
End Sub

Private Sub mnuAEIncendio_Click()
   MonitorActivate s_Incendio, False
End Sub

Private Sub mnuAEIntrusao_Click()
   MonitorActivate s_Intrusao, False
End Sub

Private Sub mnuAEEmergencia_Click()
   MonitorActivate s_Emergencia, False
End Sub

Private Sub mnuAEPanico_Click()
   MonitorActivate s_Panico, False
End Sub

Private Sub mnuAESistema_Click()
   MonitorActivate s_Sistema, False
End Sub

Private Sub mnuAZEmergencia_Click()
   MonitorActivate s_Emergencia, True
End Sub

Private Sub mnuAZIncendio_Click()
   MonitorActivate s_Incendio, True
End Sub

Private Sub mnuAZIntrusao_Click()
   MonitorActivate s_Intrusao, True
End Sub

Private Sub mnuAZPanico_Click()
   MonitorActivate s_Panico, True
End Sub

Private Sub mnuAZSistema_Click()
   MonitorActivate s_Sistema, True
End Sub

Private Sub mnuCOperador_Click()
   frmAccess.Show vbModal
End Sub

Private Sub mnuDEIncendio_Click()
   MonitorDeactivate s_Incendio, False
End Sub

Private Sub mnuDEIntrusao_Click()
   MonitorDeactivate s_Intrusao, False
End Sub

Private Sub mnuDEEmergencia_Click()
   MonitorDeactivate s_Emergencia, False
End Sub

Private Sub mnuDEPanico_Click()
   MonitorDeactivate s_Panico, False
End Sub

Private Sub mnuDESistema_Click()
   MonitorDeactivate s_Sistema, False
End Sub

Private Sub mnuDZEmergencia_Click()
   MonitorDeactivate s_Emergencia, True
End Sub

Private Sub mnuDZIncendio_Click()
   MonitorDeactivate s_Incendio, True
End Sub

Private Sub mnuDZIntrusao_Click()
   MonitorDeactivate s_Intrusao, True
End Sub

Private Sub mnuDZPanico_Click()
   MonitorDeactivate s_Panico, True
End Sub

Private Sub mnuDZSistema_Click()
   MonitorDeactivate s_Sistema, True
End Sub

Private Sub mnuEntityCreate_Click()
   Me.ActiveForm.MousePointer = vbCrosshair
   m_DragState = StateDragging
End Sub

Private Sub MonitorUpdate(fStatus As Boolean)
   Dim lEntity As clsEntity
   For Each lEntity In lstEntity
      lEntity.status = fStatus
   Next
End Sub

Public Sub MonitorActivate(fmode As typeSensor, fAll As Boolean)
    tGrupo = -1
    frmGrupo.Show vbModal
    If tGrupo <> -1 Then
        If fAll Then
            Dim lEntity As clsEntity
            For Each lEntity In lstEntity
                lEntity.Activate fmode, tGrupo
            Next
        Else
            tEntity.Activate fmode, tGrupo
        End If
    End If
End Sub

Public Sub MonitorDeactivate(fmode As typeSensor, fAll As Boolean)
    qResponse = sxQNone
    tGrupo = -1
    frmGrupo.Show vbModal
    If tGrupo <> -1 Then
        If fAll Then
            Dim lEntity As clsEntity
            For Each lEntity In lstEntity
               lEntity.Deactivate fmode, tGrupo
            Next
        Else
            tEntity.Deactivate fmode, tGrupo
        End If
    End If
End Sub

Private Sub mnuIZEmergencia_Click()
   frmStatus.fZona = s_Emergencia
   frmStatus.Caption = "Situação corrente dos Sensores de Emergência"
   Set frmStatus.localModule = lstModule
   frmStatus.Show
End Sub

Private Sub mnuIZIncendio_Click()
   frmStatus.fZona = s_Incendio
   frmStatus.Caption = "Situação corrente dos Sensores de Incêndio"
   Set frmStatus.localModule = lstModule
   frmStatus.Show
End Sub

Private Sub mnuIZIntrusao_Click()
   frmStatus.fZona = s_Intrusao
   frmStatus.Caption = "Situação corrente dos Sensores de Intrusão"
   Set frmStatus.localModule = lstModule
   frmStatus.Show
End Sub

Private Sub mnuIZPanico_Click()
   frmStatus.fZona = s_Panico
   frmStatus.Caption = "Situação corrente dos Sensores de Pânico"
   Set frmStatus.localModule = lstModule
   frmStatus.Show
End Sub

Private Sub mnuIZSistema_Click()
   frmStatus.fZona = s_Sistema
   frmStatus.Caption = "Situação corrente dos Sensores de Sistema"
   Set frmStatus.localModule = lstModule
   frmStatus.Show
End Sub

Private Sub mnuLastEvents_Click()
   frmLastEvents.fEntity = False
   frmLastEvents.Caption = "Últimos Eventos - USystem5"
   Set frmLastEvents.lastEvents = lstEvent
   frmLastEvents.Show
End Sub

Private Sub mnuLEvent_Click()
   frmLastEvents.fEntity = True
   frmLastEvents.Caption = "Últimos Eventos - " & tEntity.vDescr
   Set frmLastEvents.lastEvents = tEntity.localEvent
   frmLastEvents.Show
End Sub

Private Sub mnuMAtivo_Click()
   tEntity.ChangeSZona stAtivada
End Sub

Private Sub mnuMDesativo_Click()
   tEntity.ChangeSZona stDesativada
End Sub

Private Sub mnuPisoRemove_Click()
   Dim lPiso As Integer
   lPiso = ForNet.ActiveForm.curPiso
   Dim lDescr As String
   lDescr = ForNet.ActiveForm.Caption
   If MsgBox("Confirma a remoção do " & lDescr & "?" & Chr(13) & Chr(10) & _
             "Isto implica em remover todas as Entidades da planta!", sxQuestion, sxProname) = vbYes Then
      Update_Display "Aguarde...", sxImgInform, False
      Dim lcr As New ADODB.Command
      Set lcr.ActiveConnection = cnDB
      lcr.CommandType = adCmdText
      lcr.CommandText = "DELETE FROM Entity WHERE (fk_Floor =" & lPiso & ")"
      lcr.Execute
      lcr.CommandText = "DELETE FROM Floor WHERE (cp_Floor = " & lPiso & ")"
      lcr.Execute
      Status_Menu m_bDesignMode
      m_bUserUnload = False
      Dim tpiso As clsPiso
      For Each tpiso In lstPiso
         Unload tpiso.f
         Unload lblPiso(tpiso.n)
         lstPiso.Remove 1
      Next
      m_bUserUnload = True
      DBase_ReOpen fIsRestore:=False, fPiso:=lPiso
      m_iCurPiso = 0
      m_iLastTop = 90
      Load_Pisos
   End If
End Sub

Private Sub mnuPlantaAdd_Click()
'  provoca a geração de erro se o usuário selecionar "Cancel"
   cdl.CancelError = True
   On Error GoTo adderror
'  prepara flags
   cdl.FLAGS = cdlOFNHideReadOnly Or cdlOFNFileMustExist
'  prepara titulo da caixa de diálogo
   cdl.DialogTitle = "Indicar figura a ser carregada"
'  prepara filtros
   cdl.Filter = "Todas as figuras (*.*)|*.*|Bitmaps " & _
                     "(*.bmp)|*.bmp"
'  especifica o filtro padrão
   cdl.FilterIndex = 2
'  diretório default
   cdl.InitDir = App.Path & "\Pisos"
'  mostra o diálogo Open
   cdl.ShowOpen
   If cdl.fileName <> "" Then
      Dim lPos As Integer
      lPos = InStr(cdl.fileName, cdl.FileTitle)
      Dim lpath As String
      lpath = Left$(cdl.fileName, lPos - 1)
      If UCase$(lpath) <> UCase$(App.Path) & "\PISOS\" Then
         FileCopy cdl.fileName, App.Path & "\Pisos\" & cdl.FileTitle
      End If
      Dim cmFloor As New ADODB.Command
      Set cmFloor.ActiveConnection = cnDB
      cmFloor.CommandType = adCmdText
      cmFloor.CommandText = "UPDATE Floor SET Picture_Floor = '" & cdl.FileTitle & _
                            "' WHERE cp_Floor = " & ForNet.ActiveForm.curPiso
      cmFloor.Execute
      Set ForNet.ActiveForm.Picture = LoadPicture(App.Path & "\Pisos\" & cdl.FileTitle)
   End If
   Exit Sub
adderror:
   MsgBox Err.Description & sxContact, sxExclamation, sxProname
End Sub

Private Sub mnuPisoAdd_Click()
   Dim f As New frmPlanta
   Load frmPiso
tryPiso:
   On Error GoTo NumError
   m_bUserUnload = False
   frmPiso.Show vbModal
   If m_bUserUnload Then
      Dim cmFloor As New ADODB.Command
      Set cmFloor.ActiveConnection = cnDB
      cmFloor.CommandType = adCmdText
      cmFloor.CommandText = "INSERT INTO Floor (cp_Floor, Descr_Floor) VALUES (" & _
                            frmPiso.txtNumPiso & ", '" & frmPiso.txtDescrPiso & "')"
      cmFloor.Execute
      m_bUserUnload = False
      On Error Resume Next
      'Here, m_iCurPiso may be invalid
      lblPiso(m_iCurPiso).ForeColor = lblColor.cBlack
      On Error GoTo 0
      Set tpiso = New clsPiso
      m_iCurPiso = frmPiso.txtNumPiso
      lstPiso.Add Item:=tpiso, Key:=CStr(m_iCurPiso)
      tpiso.rCaption = frmPiso.txtDescrPiso
      tpiso.n = m_iCurPiso
      tpiso.rStep = m_iLastTop
      m_iLastTop = m_iLastTop + 800      '390
      tpiso.c_bSetores = False
      Set tpiso.f = f
      Unload frmPiso
      Load f
      f.curPiso = m_iCurPiso
      f.Caption = tpiso.rCaption
      Load lblPiso(m_iCurPiso)
      lblPiso(m_iCurPiso).ForeColor = lblColor.cBlue
      lblPiso(m_iCurPiso).Top = tpiso.rStep
      lblPiso(m_iCurPiso).Caption = tpiso.rCaption
      lblPiso(m_iCurPiso).Visible = True
      pctPiso.Visible = lstPiso.Count > 1
      mnuPlantaAdd_Click
   End If
   Exit Sub
NumError:
   MsgBox "O número do Piso/Andar deve ser único!", sxExclamation, sxProname
   Resume tryPiso
End Sub

Private Sub mnuConfig_Click()
   frmSetup.Show vbModal
End Sub

Private Sub mnuEntityDelete_Click()
   ForNet.ActiveForm.Entity_Delete
End Sub

Private Sub mnuEntityProperties_Click()
   ForNet.ActiveForm.Entity_Edit
End Sub

Private Sub mnuExit_Click()
    Dim lmsg As String
    Dim Thwnd As Long

    lmsg = "A saída do programa pode resultar em perda de dados armazenados em memória. Deseja Continuar?"
    On Error Resume Next
    Dim success As Long
   
    success = SetWindowPos(curhwnd, HWND_NOTOPMOST, 0, 0, 0, 0, FLAGS)
    If MsgBox(lmsg, vbYesNo + vbDefaultButton2 + vbQuestion, sxProname) = vbYes Then
   
        Screen.MousePointer = vbHourglass
        frmSplash.Show
        frmSplash.lblMsg = "Salvando Configurações do Sistema. Aguarde..."
        ProgressCounter = 0
        IncProgress
        frmSplash.ProgressBar1.Visible = True
        DoEvents
        Save_Last_Activities shareable:=False
        Make_Service "Desativação da rede USystem", strAccess(m_tAccess) & m_sUser
        Save_Other_Configs fLeftPos:=mnuLeft.Checked
        IncProgress
        cnDB.Close: Set cnDB = Nothing
        IncProgress
        m_bUserUnload = False
        Unload frmSplash
        Screen.MousePointer = vbDefault
        
        Thwnd = FindWindow("Shell_traywnd", "")
        Call SetWindowPos(Thwnd, 0, 0, 0, 0, 0, SWP_SHOWWINDOW)
    
        End
    End If
   
End Sub

Private Sub Status_Menu(fStatus As Boolean)
   mnuModeDesign.Checked = fStatus
   mnuCOperador.Enabled = True 'fStatus
   mnuPisoAdd.Enabled = fStatus
   mnuPisoRemove.Enabled = fStatus And (lstPiso.Count > 0)
   mnuPlantaAdd.Enabled = fStatus
   mnuEntityDelete.Enabled = fStatus
   mnuMonitorActivate.Enabled = Not fStatus
   mnuMonitorDeactivate.Enabled = Not fStatus
   mnuEntityCreate.Enabled = fStatus
   MonitorUpdate Not fStatus
End Sub

Public Sub Update_Display(fStr As String, ByVal fImgIndex As Integer, ByVal fEvent As Boolean, Optional ByVal fforce As Boolean = False)
   If fforce Or Not m_UpdateLock Then
      If Not fEvent Then
         If noDisp And Not fforce Then Exit Sub
      Else
         noDisp = True
      End If
      If StatusBar1.Panels.Item(2).Text <> fStr Then
         StatusBar1.Panels.Item(2).Text = fStr
         If fImgIndex = sxImgNone Then
           Set StatusBar1.Panels.Item(2).Picture = Nothing
         End If
      End If
   End If
End Sub

Private Sub mnuModeDesign_Click()

    m_bDesignMode = Not m_bDesignMode
    Status_Menu m_bDesignMode
    
    If Not m_bDesignMode Then
      On Error Resume Next
      MousePointer = vbDefault
      Me.ActiveForm.MousePointer = vbDefault
    End If
    
End Sub

Private Sub mnuSLocais_Click()
   frmStLocal.Caption = "Situação corrente dos Locais"
   Set frmStLocal.localModule = lstModule
   frmStLocal.Show
End Sub

Private Sub mnuZConfig_Click()
   reloadForm = False
   Load frmRegister
   Set frmRegister.fEntity = tEntity   'retain the current Entity
   frmRegister.Caption = "Configuração de Zona - " & tEntity.vDescr
   frmRegister.Show
End Sub

Private Sub mnuZStatus_Click()
   frmStLocal.Caption = "Situação corrente da Zonas - " & tEntity.vDescr
   Set frmStLocal.localModule = tEntity.localModule
   frmStLocal.Show
End Sub

Private Sub StatusBar1_PanelClick(ByVal Panel As MSComctlLib.Panel)
   If Panel.Index = 2 Then
      Update_Display "", sxImgNone, False
   End If
End Sub

'Timer que controla os Inativos
Private Sub trmInativos_Timer()
   Static flagOpen As Boolean       'Init = false
   Static flagClose As Boolean      'Init = false
   
   If Weekday(Date) <> curWeekday Then
      'Clear_Tickets_Percurso
      Clear_Entity_Status fOpen:=True
      Clear_Entity_Status fOpen:=False
      curWeekday = Weekday(Date)
      flagOpen = False
   End If

   'Chama a rotina que controla as rondas
   'Treat_Percurso_Ronda
   
   'Chama a rotina que controla os objetos inativos
   Treat_Inativos
   
   Static dumpActivity As Integer
   dumpActivity = dumpActivity + 1
   If dumpActivity >= 10 Then
      dumpActivity = 0
      Save_Last_Activities shareable:=True
   End If
   
   If Abs(DateDiff("s", m_dTOpen(curWeekday), Time)) <= 15 And Not flagOpen Then
      flagOpen = True
      Dump_Entity_Status fOpen:=True
      flagClose = False
   End If
   
   If Abs(DateDiff("s", m_dTClose(curWeekday), Time)) <= 15 And Not flagClose Then
      flagClose = True
      Dump_Entity_Status fOpen:=False
   End If
   
   'Controla a execução de backup automático
   If m_bBackupAuto Then
      If DateDiff("n", m_sHorario, Time) = 0 Then
         frmService.Automatic_Backup
         DBEvent_CleanUp fInterval:=m_iEvKeep
      End If
   End If
      
End Sub

'Timer que verifica se ha algum evento na fila para ser tratado.
Private Sub trmLastEvents_Timer()
    Dim tEvent As clsEvent
    
    trmLastEvents.Enabled = False
    DoEvents
    While tQueue.Count > 0
        Set tEvent = tQueue.Dequeue
        'Verifica diretivas de segurança. Descarta evento caso Key_check = 0 e BYPASS = 0
        If Key_check = 1 Or BYPASS = 1 Then
            With tEvent
                If .sHeader = H_Serial Then
                    .TreatReceiver
                ElseIf .sHeader = H_Device And .sMID = MID_Repeater Then
                    .TreatRepeater
                ElseIf .sHeader = H_Device And .sMID = MID_Sensor Then
                    .TreatSensor
                End If
                '.Persist  (disabled, here. See clsEvent. Marcio, jul/2012)
            End With
        Else
            Set tDisplay = New clsDisplay
            tDisplay.dispMode = sxErSound
            tDisplay.dispFile = ""
            tDisplay.dispStr = tQueue.Count + 1 & " evento(s) descartado(s) por problemas na chave de segurança!"
            tDisplay.dispImg = sxImgAlert
            Insert_Display tDisplay, False, True
            If m_Debug Then Make_Service tQueue.Count + 1 & " evento(s) descartado(s) por problemas na chave de segurança!", " "
            tQueue.Clear
        End If
        DoEvents
    Wend
    trmLastEvents.Enabled = True

End Sub

Private Sub trmPanico_Timer()
'   If lstPanico.Count > 0 Then
'      Dim IdxPanico As Integer
'      IdxPanico = 0
'      Dim lpnc As clsPanico
'      For Each lpnc In lstPanico
'         IdxPanico = IdxPanico + 1
'         With lpnc
'            .pTime = .pTime - 1
'            If .pTime <= 0 Then
'               Dim lEvent As New clsEvent
'               .lModule.TreatAlarme stFechado, .lEvent
'               lstPanico.Remove IdxPanico
'               ' quando se remove um objeto, o indice deve ser decrementado de 1
'               IdxPanico = IdxPanico - 1
'            End If
'         End With
'      Next
'   Else
'      trmPanico.Enabled = False
'   End If
End Sub

'Timers que controlam a janela de entrada de informações nos buffers de comunicacao
Private Sub trmAquire_Timer(Index As Integer)
    trmAquire(Index).Enabled = False
End Sub

'Timer que redesenha as entidades na tela
Private Sub trmRefresh_Timer()
    Static flagToogle As Boolean     'Init = false
    Static flagInvert As Boolean     'Init = false
    
    On Error Resume Next
    Me.ActiveForm.Redesenha
    On Error GoTo 0
    flagToogle = Not flagToogle
    If flagToogle Then
        ' Animation Online button
        flagInvert = Not flagInvert
        btnAction(16).Effects.Invert = flagInvert
    End If
End Sub

'Timer que chama a verificação da segurança a cada 5 minutos
Private Sub trmSecurity_Timer()
    segCounter = segCounter + 1
    If segCounter = TempoCheckSecurity Then
        Security_Check
        Set_Caption
        segCounter = 0
    End If
End Sub

'Timer que controla o tempo dos serviços (dupla verificaçao, etc.)
Private Sub trmService_Timer()

   If lstService.Count > 0 Then
      Dim IdxService As Integer
      Dim lsrv As clsService
      
      IdxService = 0
      
      For Each lsrv In lstService
      IdxService = IdxService + 1
         With lsrv
            If .stime = 0 Then
                Service_Treat lsrv
            Else
                .stime = .stime - 1
            End If
         End With
      Next
      
   End If
      
End Sub

Private Sub SetStatus(ByVal fAccess As typeAccess)

    If fAccess = sxSystem Then
        Dim i As Integer
        For i = 11 To 15
            SetAppearence btnAction(i), True
        Next i
    Else
        SetAppearence btnAction(12), (m_tAccess = sxAdministrador) Or (m_tAccess = sxSupervisor)
        SetAppearence btnAction(13), (m_tAccess = sxAdministrador)
        SetAppearence btnAction(14), False  '(m_tAccess = sxAdministrador)
        SetAppearence btnAction(11), (fAccess <> sxOperator)
        SetAppearence btnAction(15), (fAccess <> sxOperator)
    End If
    
    mnuExit.Enabled = fAccess <> sxOperator
    mnuConfig.Enabled = (fAccess = sxAdministrador) Or (fAccess = sxSystem)
    mnuModeDesign.Enabled = (fAccess = sxAdministrador) Or (fAccess = sxSystem)
    mnuPisos.Enabled = fAccess <> sxOperator
    mnuBackup.Enabled = fAccess <> sxOperator
    mnuBaseStatus.Enabled = fAccess <> sxOperator
   
End Sub

Private Sub Disparo_Clear()
   Dim lEntity As clsEntity
   For Each lEntity In lstEntity
      lEntity.UpdateColor clearDisp:=True
   Next
End Sub

Private Sub Service_Treat(fsrv As clsService)
   If fsrv.skind = check_dupla Then
      Dim lModule As clsModule
      Set lModule = lstModule.Item(CStr(fsrv.sModule))
      lModule.TreatAlarme fsrv.stype, fsrv.sEvent
      lModule.Remove_Service fSkind:=check_dupla
      lModule.flagDupla = False
   End If
End Sub

'Rotina que carrega o setup de posicionamento dos PISOS
Public Sub Pisos_Setup()
   If m_bPisoLeft Then
      pctPiso.Align = vbAlignLeft
      pctPiso.Width = 1800
      lblPiso(0).Left = 0
      lblPiso(0).Top = -600
      cLblStep = 600    '390
      m_iLastTop = 90
   Else
      pctPiso.Align = vbAlignBottom
      pctPiso.Height = 800
      lblPiso(0).Left = -1600
      lblPiso(0).Top = 70
      cLblStep = 1600      '1030
      m_iLastTop = 600
   End If
   
   If pctPiso.Visible Then
      'Was loaded
      Dim tpiso As clsPiso
      For Each tpiso In lstPiso
         tpiso.rStep = m_iLastTop
         m_iLastTop = m_iLastTop + cLblStep
         If m_bPisoLeft Then
            lblPiso(tpiso.n).Left = 0
            lblPiso(tpiso.n).Top = tpiso.rStep
         Else
            lblPiso(tpiso.n).Top = 70
            lblPiso(tpiso.n).Left = tpiso.rStep
         End If
      Next
   End If
   
End Sub

Public Sub Treat_Inativos()
   If m_bCommStatus Then
      Dim cM As clsModule
      'Dim cE As clsEntity
      Dim hasInativos As Boolean
      hasInativos = False
      For Each cM In lstModule
         With cM
            If .mChkAtiv And .mStatAtiv Then
               If DateDiff("n", .mLastAtiv, Now) > .mtempoAtiv Then
                  .SetFalha
                  hasInativos = True
                 ' Set cE = lstEntity.Item(CStr(.mEntity))
                 ' cE.SetInatividade
               End If
            End If
         End With
      Next
      If hasInativos Then
         btnAction(11).GrayScale = lvicNoGrayScale
         Update_Display "Há módulos/dispositivos inativos na rede USystem", sxImgInform, False
      Else
         btnAction(11).GrayScale = lvicGreenMask
      End If
   End If
End Sub

'Preenche o Caption da tela principal
Public Sub Set_Caption()
   Caption = "Sistema de Segurança USystem (versão 5.0." & App.Revision & ")"
   
   If Trim(gstCompany) <> "" Then
        Caption = Caption & " -- " & gstCompany & " -- "
   End If

   If BYPASS = 1 Then
        Caption = Caption & " (BYPASSED)"
   ElseIf Key_check = 1 Then
        Caption = Caption & " (REGISTRO Ok)"
   Else
        Caption = Caption & " (SEM REGISTRO)"
   End If

End Sub
