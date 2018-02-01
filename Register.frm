VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{A4749554-0441-4E64-8A03-3323601631C7}#1.0#0"; "LaVolpeAlphaImg2.ocx"
Begin VB.Form frmRegister 
   Caption         =   "Configuração de Zona"
   ClientHeight    =   8115
   ClientLeft      =   165
   ClientTop       =   810
   ClientWidth     =   10065
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8115
   ScaleWidth      =   10065
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtGrupo 
      Height          =   300
      Left            =   1080
      MaxLength       =   3
      TabIndex        =   67
      TabStop         =   0   'False
      Top             =   9600
      Width           =   1065
   End
   Begin VB.TextBox txtColor 
      Height          =   300
      Left            =   7200
      MaxLength       =   3
      TabIndex        =   57
      TabStop         =   0   'False
      Top             =   9240
      Width           =   1065
   End
   Begin VB.TextBox txtTipoDevice 
      Height          =   300
      Left            =   7200
      MaxLength       =   3
      TabIndex        =   40
      TabStop         =   0   'False
      Top             =   8760
      Width           =   1065
   End
   Begin MSComDlg.CommonDialog cdl 
      Left            =   6030
      Top             =   9180
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox txtTipoZona 
      Height          =   300
      Left            =   6000
      MaxLength       =   3
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   8775
      Width           =   1065
   End
   Begin VB.TextBox txtSaidaPGM 
      Height          =   300
      Left            =   4815
      MaxLength       =   3
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   9210
      Width           =   1035
   End
   Begin VB.TextBox txtTipoPGM 
      Height          =   300
      Left            =   3525
      MaxLength       =   3
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   9195
      Width           =   1035
   End
   Begin VB.TextBox txtControlePGM 
      Height          =   300
      Left            =   2325
      MaxLength       =   3
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   9180
      Width           =   1035
   End
   Begin VB.TextBox txtTipoLogica 
      Height          =   300
      Left            =   1065
      MaxLength       =   3
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   9150
      Width           =   1035
   End
   Begin VB.TextBox txtJanela 
      Height          =   300
      Left            =   4800
      MaxLength       =   3
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   8775
      Width           =   1035
   End
   Begin VB.TextBox txtCheck 
      Height          =   300
      Left            =   3510
      MaxLength       =   3
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   8775
      Width           =   1035
   End
   Begin VB.TextBox txtDebounce 
      Height          =   300
      Left            =   2295
      MaxLength       =   3
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   8775
      Width           =   1065
   End
   Begin VB.TextBox txtInicialZona 
      Height          =   300
      Left            =   1050
      MaxLength       =   3
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   8760
      Width           =   1065
   End
   Begin VB.Frame frameZona 
      Caption         =   "Zona "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7215
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   9855
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
         ItemData        =   "Register.frx":0000
         Left            =   3120
         List            =   "Register.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   66
         Top             =   960
         Width           =   1935
      End
      Begin VB.TextBox txtNumeroZona 
         Alignment       =   2  'Center
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
         Left            =   8500
         MaxLength       =   15
         TabIndex        =   64
         Text            =   "No. Zona"
         Top             =   360
         Width           =   1215
      End
      Begin VB.Frame itico 
         Caption         =   "Tratamento "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Left            =   120
         TabIndex        =   43
         Top             =   4440
         Width           =   9615
         Begin VB.CheckBox chkPopup 
            Caption         =   "Popup ?"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   315
            Left            =   240
            TabIndex        =   62
            Top             =   360
            Width           =   1455
         End
         Begin VB.TextBox txtUser 
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
            IMEMode         =   3  'DISABLE
            Left            =   960
            MaxLength       =   255
            TabIndex        =   61
            Text            =   "admin"
            Top             =   1410
            Width           =   1575
         End
         Begin VB.TextBox txtPasswd 
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
            IMEMode         =   3  'DISABLE
            Left            =   4560
            MaxLength       =   255
            PasswordChar    =   "*"
            TabIndex        =   59
            Top             =   1410
            Width           =   1575
         End
         Begin VB.CheckBox chkCritico 
            Caption         =   "Sensor Crítico ?"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   315
            Left            =   2400
            TabIndex        =   53
            Top             =   360
            Width           =   1575
         End
         Begin VB.Frame fraColor 
            Caption         =   "Selecione a cor para evento em Sensor Crítico "
            Height          =   735
            Left            =   4800
            TabIndex        =   48
            Top             =   120
            Width           =   4695
            Begin VB.OptionButton optColor 
               Caption         =   "azul"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFF80&
               Height          =   315
               Index           =   1
               Left            =   1320
               TabIndex        =   52
               Top             =   240
               Width           =   855
            End
            Begin VB.OptionButton optColor 
               Caption         =   "vermellho"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   315
               Index           =   0
               Left            =   120
               TabIndex        =   51
               Top             =   240
               Value           =   -1  'True
               Width           =   1095
            End
            Begin VB.OptionButton optColor 
               Caption         =   "amarelo"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H0000FFFF&
               Height          =   315
               Index           =   2
               Left            =   2280
               TabIndex        =   50
               Top             =   240
               Width           =   1095
            End
            Begin VB.OptionButton optColor 
               Caption         =   "verde"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H0000FF00&
               Height          =   315
               Index           =   3
               Left            =   3600
               TabIndex        =   49
               Top             =   240
               Width           =   975
            End
         End
         Begin VB.TextBox txtServerAddress 
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
            Left            =   4560
            MaxLength       =   255
            TabIndex        =   47
            Text            =   "192.168.10.1:8601"
            Top             =   1000
            Width           =   1575
         End
         Begin VB.TextBox txtCamera 
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
            Left            =   7200
            MaxLength       =   255
            TabIndex        =   46
            Top             =   1000
            Width           =   2175
         End
         Begin VB.TextBox txtMonitor 
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
            Left            =   7200
            MaxLength       =   255
            TabIndex        =   45
            Top             =   1410
            Width           =   2175
         End
         Begin VB.CheckBox chkTelaCheia 
            Caption         =   "Visualização em tela cheia"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   315
            Left            =   240
            TabIndex        =   44
            Top             =   700
            Width           =   2415
         End
         Begin VB.Label Label2 
            Caption         =   "Usuário:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   240
            TabIndex        =   60
            Top             =   1440
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "Senha de Acesso:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   3000
            TabIndex        =   58
            Top             =   1440
            Width           =   1455
         End
         Begin VB.Image imgTratamento 
            Height          =   480
            Left            =   4080
            Picture         =   "Register.frx":0004
            Top             =   280
            Width           =   480
         End
         Begin VB.Label Label22 
            Caption         =   "Endereço IP do Servidor CFTV + "":"" +  Porta de Acesso:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   240
            TabIndex        =   56
            Top             =   1070
            Width           =   4215
         End
         Begin VB.Label Label23 
            Caption         =   "Câmera:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   6480
            TabIndex        =   55
            Top             =   1070
            Width           =   855
         End
         Begin VB.Label Label24 
            Caption         =   "Monitor:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   6480
            TabIndex        =   54
            Top             =   1440
            Width           =   735
         End
      End
      Begin VB.ComboBox lstTipoDevice 
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
         ItemData        =   "Register.frx":0446
         Left            =   1080
         List            =   "Register.frx":0465
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   960
         Width           =   1335
      End
      Begin VB.TextBox txtSerialNumber 
         Alignment       =   2  'Center
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
         Left            =   5520
         MaxLength       =   8
         TabIndex        =   3
         Text            =   "00000000"
         Top             =   960
         Width           =   1095
      End
      Begin VB.TextBox PTI 
         Alignment       =   2  'Center
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
         Left            =   8500
         Locked          =   -1  'True
         MaxLength       =   15
         TabIndex        =   5
         Text            =   "Serial Receiver"
         Top             =   960
         Width           =   1215
      End
      Begin VB.TextBox txtArquivo 
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
         Left            =   1560
         MaxLength       =   255
         TabIndex        =   14
         Top             =   6600
         Width           =   7095
      End
      Begin VB.TextBox txtLocalZona 
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
         Left            =   1080
         MaxLength       =   70
         TabIndex        =   1
         Top             =   360
         Width           =   5850
      End
      Begin VB.ComboBox lstDescrZona 
         Height          =   315
         ItemData        =   "Register.frx":04C1
         Left            =   1080
         List            =   "Register.frx":04C3
         Style           =   2  'Dropdown List
         TabIndex        =   36
         Top             =   360
         Width           =   6135
      End
      Begin VB.TextBox txtReceptor 
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
         Left            =   8400
         Locked          =   -1  'True
         MaxLength       =   15
         TabIndex        =   16
         Text            =   "txtReceptor"
         Top             =   3720
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox txtUID 
         Alignment       =   2  'Center
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
         Left            =   7080
         Locked          =   -1  'True
         MaxLength       =   8
         TabIndex        =   4
         Text            =   "B2HHHHHH"
         Top             =   960
         Width           =   975
      End
      Begin VB.Frame frameAtividade 
         Caption         =   "Controle de Atividade: "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   120
         TabIndex        =   33
         Top             =   3360
         Width           =   6255
         Begin VB.Frame fraAtividade 
            Height          =   735
            Left            =   1920
            TabIndex        =   34
            Top             =   120
            Width           =   4095
            Begin VB.TextBox txtTempo 
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
               Left            =   3120
               MaxLength       =   2
               TabIndex        =   13
               Top             =   240
               Width           =   735
            End
            Begin VB.Label lblTempo 
               Caption         =   "Tempo de verificação (1 .. 59 minutos):"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   255
               Left            =   120
               TabIndex        =   41
               Top             =   300
               Width           =   3135
            End
         End
         Begin VB.CheckBox chkAtividade 
            Caption         =   "Verificar Atividade"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   315
            Left            =   240
            TabIndex        =   12
            Top             =   360
            Width           =   1695
         End
      End
      Begin VB.Frame frameLogica 
         Caption         =   "Tipo de Lógica: "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   120
         TabIndex        =   23
         Top             =   2160
         Width           =   9615
         Begin VB.TextBox txtLocalLogica 
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
            Left            =   4680
            MaxLength       =   70
            TabIndex        =   11
            Top             =   360
            Width           =   4695
         End
         Begin VB.TextBox txtNumeroLogica 
            Alignment       =   2  'Center
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
            Left            =   3360
            MaxLength       =   3
            TabIndex        =   10
            Top             =   360
            Width           =   495
         End
         Begin VB.ComboBox lstTipoLogica 
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
            ItemData        =   "Register.frx":04C5
            Left            =   360
            List            =   "Register.frx":04CF
            Style           =   2  'Dropdown List
            TabIndex        =   9
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            Caption         =   "Local:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   4080
            TabIndex        =   22
            Top             =   390
            Width           =   495
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            Caption         =   "Zona associada:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   1920
            TabIndex        =   21
            Top             =   390
            Width           =   1335
         End
      End
      Begin VB.ComboBox lstJanela 
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
         ItemData        =   "Register.frx":04E1
         Left            =   8400
         List            =   "Register.frx":04F5
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   1680
         Width           =   1215
      End
      Begin VB.ComboBox lstCheck 
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
         ItemData        =   "Register.frx":0519
         Left            =   5040
         List            =   "Register.frx":0523
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   1680
         Width           =   1335
      End
      Begin VB.ComboBox lstInicial 
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
         ItemData        =   "Register.frx":0537
         Left            =   1320
         List            =   "Register.frx":0547
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   1680
         Width           =   1335
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Grupo: "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   2520
         TabIndex        =   65
         Top             =   1000
         Width           =   495
      End
      Begin VB.Label lblZona 
         Alignment       =   1  'Right Justify
         Caption         =   "Zona:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   7840
         TabIndex        =   63
         Top             =   405
         Width           =   615
      End
      Begin LaVolpeAlphaImg.AlphaImgCtl btnReg 
         Height          =   720
         Index           =   6
         Left            =   8880
         ToolTipText     =   "buscar arquivo de som"
         Top             =   6315
         Width           =   720
         _ExtentX        =   1270
         _ExtentY        =   1270
         Image           =   "Register.frx":057A
         Effects         =   "Register.frx":1739
      End
      Begin VB.Label lblReceptor 
         Alignment       =   1  'Right Justify
         Caption         =   "Receptor:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   7440
         TabIndex        =   42
         Top             =   3730
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label lblArquivoSom 
         Caption         =   "Arquivo de Som:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   240
         TabIndex        =   37
         Top             =   6630
         Width           =   1215
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00800000&
         X1              =   120
         X2              =   9600
         Y1              =   1440
         Y2              =   1440
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         Caption         =   "Dispositivo:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   120
         TabIndex        =   39
         Top             =   1000
         Width           =   855
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "SNº:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   5000
         TabIndex        =   38
         Top             =   1000
         Width           =   495
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         Caption         =   "Descrição:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   120
         TabIndex        =   35
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label20 
         Alignment       =   1  'Right Justify
         Caption         =   "PTI:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   8100
         TabIndex        =   15
         Top             =   990
         Width           =   375
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         Caption         =   "UID:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   6680
         TabIndex        =   17
         Top             =   1005
         Width           =   375
      End
      Begin VB.Label lblJanela 
         Alignment       =   1  'Right Justify
         Caption         =   "Janela de verificação:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   6240
         TabIndex        =   20
         Top             =   1710
         Width           =   2055
      End
      Begin VB.Label lblVerificação 
         Alignment       =   1  'Right Justify
         Caption         =   "Tipo de Verificação:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   3240
         TabIndex        =   19
         Top             =   1710
         Width           =   1695
      End
      Begin VB.Label lblInicial 
         Alignment       =   1  'Right Justify
         Caption         =   "Status Inicial:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   1710
         Width           =   1095
      End
   End
   Begin LaVolpeAlphaImg.AlphaImgCtl btnReg 
      Height          =   720
      Index           =   3
      Left            =   2640
      ToolTipText     =   "Inserir"
      Top             =   40
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   1270
      Image           =   "Register.frx":1751
      Effects         =   "Register.frx":2346
   End
   Begin LaVolpeAlphaImg.AlphaImgCtl btnReg 
      Height          =   720
      Index           =   4
      Left            =   3480
      ToolTipText     =   "Remover"
      Top             =   40
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   1270
      Image           =   "Register.frx":235E
      Effects         =   "Register.frx":3174
   End
   Begin LaVolpeAlphaImg.AlphaImgCtl btnReg 
      Height          =   720
      Index           =   5
      Left            =   4320
      ToolTipText     =   "Sair da Configuração de Zonas"
      Top             =   40
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   1270
      Image           =   "Register.frx":318C
      Effects         =   "Register.frx":3E91
   End
   Begin LaVolpeAlphaImg.AlphaImgCtl btnReg 
      Height          =   720
      Index           =   0
      Left            =   120
      ToolTipText     =   "Alterar"
      Top             =   40
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   1270
      Image           =   "Register.frx":3EA9
      Effects         =   "Register.frx":4C69
   End
   Begin LaVolpeAlphaImg.AlphaImgCtl btnReg 
      Height          =   720
      Index           =   1
      Left            =   960
      ToolTipText     =   "Atualizar"
      Top             =   40
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   1270
      Image           =   "Register.frx":4C81
      Effects         =   "Register.frx":5ADF
   End
   Begin LaVolpeAlphaImg.AlphaImgCtl btnReg 
      Height          =   720
      Index           =   2
      Left            =   1800
      ToolTipText     =   "Cancelar"
      Top             =   40
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   1270
      Image           =   "Register.frx":5AF7
      Effects         =   "Register.frx":69A9
   End
   Begin VB.Menu mnuConfig 
      Caption         =   "Configuração"
      Begin VB.Menu mnuAlterar 
         Caption         =   "Alterar"
      End
      Begin VB.Menu mnuSalvar 
         Caption         =   "Salvar"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuS1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuInserir 
         Caption         =   "Inserir"
      End
      Begin VB.Menu mnuCancelar 
         Caption         =   "Cancelar"
      End
      Begin VB.Menu mnuS2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRemover 
         Caption         =   "Remover"
      End
      Begin VB.Menu mnuS3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Fechar"
         Shortcut        =   ^F
      End
   End
End
Attribute VB_Name = "frmRegister"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public fEntity As clsEntity
Private fModule As clsModule
Private changeOff As Boolean
Private fEditMode As ADODB.EditModeEnum
Private rsSensor As New ADODB.Recordset

Private Const invalidDataSet As Integer = 9999
Private Const invalidUpdate  As Integer = 9898
Private Const invalidLoadReg As Integer = 9797
Private Const invalidSaveReg As Integer = 3021

Private Sub SetAppearence(btn As AlphaImgCtl, flag As Boolean)
   If flag Then
      btn.GrayScale = lvicNoGrayScale
   Else
      btn.GrayScale = lvicGreenMask
   End If
   btn.Enabled = flag
End Sub

Private Sub btnReg_Click(Index As Integer)
   Select Case Index
      Case 0
         mnuAlterar_Click
      Case 1
         mnuSalvar_Click
      Case 2
         mnuCancelar_Click
      Case 3
         mnuInserir_Click
      Case 4
         mnuRemover_Click
      Case 5
         mnuExit_Click
      Case 6
         cmdBusca
   End Select
End Sub

Private Sub btnReg_MouseEnter(Index As Integer)
   If btnReg(Index).Enabled Then
      btnReg(Index).SetRedraw = False
      btnReg(Index).GrayScale = lvicSepia
      btnReg(Index).LightnessPct = -20
      btnReg(Index).SetRedraw = True
   End If
End Sub

Private Sub btnReg_MouseExit(Index As Integer)
   If btnReg(Index).Enabled Then
      btnReg(Index).SetRedraw = False
      btnReg(Index).GrayScale = lvicNoGrayScale
      btnReg(Index).LightnessPct = 0
      btnReg(Index).SetRedraw = True
   End If
End Sub

Private Sub chkAtividade_Click()
   If chkAtividade.Value = vbChecked Then
      fraAtividade.Visible = True
   Else
      fraAtividade.Visible = False
   End If
End Sub

Private Sub chkCritico_Click()
   If chkCritico.Value = vbChecked Then
      imgTratamento.Visible = True
      fraColor.Visible = True
   Else
      imgTratamento.Visible = False
      fraColor.Visible = False
   End If
End Sub

Private Sub cmdBusca()
'  provoca a geração de erro se o usuário selecionar "Cancel"
   cdl.CancelError = True
   On Error GoTo findHandler
'  prepara flags
   cdl.FLAGS = cdlOFNHideReadOnly Or cdlOFNFileMustExist
'  prepara titulo da caixa de diálogo
   cdl.DialogTitle = "Indicar arquivo de mensagem"
'  prepara filtros
   cdl.Filter = "Todos (*.*)|*.*|Arquivos de Som " & _
                     "(*.wav)|*.wav"
'  especifica o filtro padrão
   cdl.FilterIndex = 2
'  diretório default
   cdl.InitDir = App.Path & "\Mensagens"
'  mostra o diálogo Open
   cdl.ShowOpen
   If InStr(cdl.InitDir, cdl.fileName) <> 0 Then
      txtArquivo = Mid$(cdl.fileName, Len(cdl.InitDir) + 2)
   Else
      txtArquivo = cdl.fileName
   End If
   Exit Sub
findHandler:
   If Err.Number <> 32755 Then 'This error occur when Cancel was selected
      Screen.MousePointer = vbDefault
      MsgBox "Error: " & Err.Description, sxImgInform, sxProname
      Exit Sub
   End If
End Sub

Private Sub Form_Activate()
   LoadGrupo
   LoadSensor
   UpdateForm fRefill:=True
End Sub

Private Sub Form_Load()
   changeOff = False
   Dim success As Long
   success = SetWindowPos(frmRegister.hWnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
   curhwnd = frmRegister.hWnd
   Left = (Screen.Width - Width) / 2   ' Center form horizontally.
   Top = (Screen.Height - Height) / 2  ' Center form vertically.
End Sub

Private Sub lstCheck_Click()
   If Not changeOff Then txtCheck = lstCheck.ListIndex
End Sub

Private Sub lstDescrZona_Click()
   Dim lSensor As String
   lSensor = lstDescrZona.ItemData(lstDescrZona.ListIndex)
   rsSensor.MoveFirst
   Dim lFound As Boolean
   lFound = False
   While Not lFound
      If rsSensor("Numero_Sensor") = lSensor Then
         lFound = True
      Else
         rsSensor.MoveNext
      End If
   Wend
   Load_Register
End Sub

Private Sub lstTipoGrupo_Click()
    If Not changeOff Then txtGrupo = lstTipoGrupo.ListIndex
End Sub

Private Sub lstInicial_Click()
   If Not changeOff Then txtInicialZona = lstInicial.ListIndex
End Sub

Private Sub lstJanela_Click()
   On Error Resume Next
   If Not changeOff Then txtJanela = lstJanela.ItemData(lstJanela.ListIndex)
   On Error GoTo 0
End Sub

Private Sub lstTipoLogica_Click()
   If Not changeOff Then txtTipoLogica = lstTipoLogica.ListIndex
End Sub

Private Sub lstTipoDevice_Click()
    If Not changeOff Then txtTipoDevice = lstTipoDevice.ListIndex
End Sub

Private Sub mnuAlterar_Click()
   mnuSalvar.Enabled = True
   mnuCancelar.Enabled = True
   mnuAlterar.Enabled = False
   mnuInserir.Enabled = False
   mnuExit.Enabled = False
   
   SetAppearence btnReg(0), False
   SetAppearence btnReg(1), True
   SetAppearence btnReg(2), True
   SetAppearence btnReg(3), False
   SetAppearence btnReg(4), False
   SetAppearence btnReg(5), False
   SetAppearence btnReg(6), True
    
   fEditMode = adEditInProgress
   Status_Controls True
   txtNumeroZona.Enabled = True
   txtSerialNumber.Enabled = True
   
End Sub

Private Sub mnuCancelar_Click()
   LoadSensor
   UpdateForm fRefill:=False
End Sub

Private Sub mnuExit_Click()
   Unload Me
End Sub

Private Sub mnuInserir_Click()
   mnuSalvar.Enabled = True
   mnuCancelar.Enabled = True
   mnuAlterar.Enabled = False
   mnuInserir.Enabled = False
   mnuRemover.Enabled = False
   mnuExit.Enabled = False
   
   SetAppearence btnReg(0), False
   SetAppearence btnReg(1), True
   SetAppearence btnReg(2), True
   SetAppearence btnReg(3), False
   SetAppearence btnReg(4), False
   SetAppearence btnReg(5), False
   SetAppearence btnReg(6), True
      
   Status_Controls True
   Clear_Register
   fEditMode = adEditAdd
   Default_Values
   chkAtividade.Value = vbUnchecked
   chkAtividade.Enabled = True
   fraAtividade.Visible = False
   
   chkCritico.Value = vbUnchecked
   chkPopup.Value = vbUnchecked
   chkTelaCheia.Value = vbUnchecked
   txtLocalZona.SetFocus
End Sub

Private Sub mnuRemover_Click()
   fEntity.Remove fModule
   lstModule.Remove fModule.UID
   
   Set fModule = Nothing
   Dim cmSensor As New ADODB.Command
   Set cmSensor.ActiveConnection = cnDB
   cmSensor.CommandType = adCmdText
   cmSensor.CommandText = "DELETE FROM Sensor WHERE (Serial_Number ='" & txtSerialNumber & "')"
   cmSensor.Execute
   DoEvents
   Make_Service "Remoção do Sensor SN= " & txtSerialNumber.Text, strAccess(m_tAccess) & m_sUser
   LoadSensor
   DoEvents
   If rsSensor.BOF Or rsSensor.EOF Then
      rsSensor.Requery
      DoEvents
   End If
   If rsSensor.BOF Or rsSensor.EOF Then
      Clear_Register
      lstDescrZona.Clear
      mnuRemover.Enabled = False
      mnuAlterar.Enabled = False
      SetAppearence btnReg(0), False
      SetAppearence btnReg(4), False
  Else
      lstDescrZona_Refill
      Load_Register
      mnuRemover.Enabled = True
      mnuAlterar.Enabled = True
      SetAppearence btnReg(0), True
      SetAppearence btnReg(4), True
   End If
   mnuCancelar.Enabled = False
   mnuSalvar.Enabled = False
   mnuInserir.Enabled = True
   mnuExit.Enabled = True
   
   Status_Controls False
   SetAppearence btnReg(1), False
   SetAppearence btnReg(3), True
   SetAppearence btnReg(2), False
   SetAppearence btnReg(5), True
   
End Sub

Private Sub mnuSalvar_Click()
   On Error GoTo TreatError
   If Verify_Consistency() Then
      If Update_Register() Then
         mnuRemover.Enabled = True
         mnuAlterar.Enabled = True
         mnuInserir.Enabled = True
         mnuCancelar.Enabled = False
         mnuSalvar.Enabled = False
         mnuExit.Enabled = True
         fEditMode = adEditNone
         Status_Controls False
         SetAppearence btnReg(0), True
         SetAppearence btnReg(1), False
         SetAppearence btnReg(3), True
         SetAppearence btnReg(2), False
         SetAppearence btnReg(4), True
         SetAppearence btnReg(5), True
         SetAppearence btnReg(6), False
      End If
   End If
   Exit Sub
TreatError:
   Select Case Err.Number
      Case invalidDataSet
         MsgBox "A rede USystem ira executar uma atualização na base de Dados!" & Chr$(13) & Chr$(10) & _
                "Este formulário será fechado e reaberto, em seguida. Caso o registro" & Chr$(13) & Chr$(10) & _
                "não seja carregado corretamente, repita a operação de reabrir o formulário", sxInformation, sxProname
         reloadForm = True
         Unload Me
         Exit Sub
      Case invalidSaveReg
         reloadForm = True
         Unload Me
         DoEvents
         Exit Sub
      Case Else
         MsgBox Err.Description, sxInformation, sxProname
   End Select
End Sub

Private Sub txtCheck_Change()
   On Error Resume Next
   lstCheck.ListIndex = txtCheck
   On Error GoTo 0
End Sub

Private Sub txtGrupo_Change()
   If IsNumeric(txtGrupo) Then
        lstTipoGrupo.ListIndex = txtGrupo
End Sub

Private Sub txtInicialZona_Change()
   On Error Resume Next
   lstInicial.ListIndex = txtInicialZona
   On Error GoTo 0
End Sub

Private Sub txtJanela_Change()
   On Error Resume Next
   Select Case txtJanela
      Case 10
         lstJanela.ListIndex = 0
      Case 20
         lstJanela.ListIndex = 1
      Case 40
         lstJanela.ListIndex = 2
      Case 80
         lstJanela.ListIndex = 3
   End Select
   On Error GoTo 0
End Sub

Private Sub txtNumeroLogica_Change()
   If Len(txtNumeroLogica) = 3 Then
      If (CInt(txtNumeroLogica) = CInt(txtNumeroZona)) Then
         Beep
         MsgBox "Não é permitido efetuar função AND com a própria Zona!", sxExclamation, sxProname
         txtNumeroLogica = ""
      Else
         Verify_Zona_AND
      End If
   End If
End Sub

'Procedure que consiste a entrada do Número da Zona, somente numeros.
Private Sub txtNumeroLogica_KeyPress(KeyAscii As Integer)
   If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then
      KeyAscii = 0
   End If
End Sub

'Procedure que padroniza o formato do número da zona logica em 3 digitos
Private Sub txtNumeroLogica_LostFocus()
    If Len(txtNumeroLogica) = 1 Then
        txtNumeroLogica = "00" & txtNumeroLogica
    ElseIf Len(txtNumeroLogica) = 2 Then
        txtNumeroLogica = "0" & txtNumeroLogica
    End If
End Sub

'Procedure que ao final do preenchimento do número da zona coloca sempre o número com 3 dígitos.
Private Sub txtNumeroZona_LostFocus()
    If Len(txtNumeroZona) = 1 Then
        txtNumeroZona = "00" & txtNumeroZona
    ElseIf Len(txtNumeroZona) = 2 Then
        txtNumeroZona = "0" & txtNumeroZona
    End If
End Sub

'Procedure que consiste a entrada do Número da Zona, somente numeros.
Private Sub txtNumeroZona_KeyPress(KeyAscii As Integer)
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub

'Procedure que ao final do preenchimento do n. de série,'calcula o UID e preenche o campo na tela de cadastro da zona.
Private Sub txtSerialNumber_Change()
   If Len(txtSerialNumber) = 8 Then
       txtUID = Hex(txtSerialNumber)
       txtUID = Right(txtUID, 6)
       If txtTipoDevice = 7 Then
           txtUID = "01" & txtUID
       ElseIf txtTipoDevice = 8 Then
           txtUID = "00" & txtUID
       Else
           txtUID = "B2" & txtUID
       End If
   Else
       txtUID = ""
   End If
End Sub

'Procedure que consiste a entrada do Número de Serie do dispositivo (Serial Receiver, Repeaters e Sensores)
Private Sub txtSerialNumber_KeyPress(KeyAscii As Integer)
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtTipoLogica_Change()
   On Error Resume Next
   lstTipoLogica.ListIndex = txtTipoLogica
   On Error GoTo 0
End Sub

'Metodo que ajusta a tela de configuração dependendo do dispositivo a ser cadastrado, rede ou sensor
Private Sub txtTipoDevice_Change()
   If IsNumeric(txtTipoDevice) Then
        lstTipoDevice.ListIndex = txtTipoDevice
        'Dispositivo é um Receptor ou um Repetidor
'        If txtTipoDevice = 7 Or txtTipoDevice = 8 Then
'             ConfigAdjust False
'             frmRegister.Caption = "Configuração de Rede"
'             frameAtividade.Top = 1560
'             lblArquivoSom.Top = 2790
'             txtArquivo.Top = 2760
'             btnReg(6).Top = 2480
'             frameZona.Height = 3255
'             frmRegister.Height = 4980
'        Else
'             'Dispositivo é um Sensor
'             ConfigAdjust True
'             frmRegister.Caption = "Configuração de Zona"
'             frameAtividade.Top = 3360
'             lblArquivoSom.Top = 4590
'             txtArquivo.Top = 4560
'             btnReg(6).Top = 4280
'             frameZona.Height = 5175
'             frmRegister.Height = 6890
'        End If
    End If
End Sub

'Metodo que ajusta a tela de configuração dependendo do dispositivo a ser cadastrado, rede ou sensor
Private Sub ConfigAdjust(fEnabled As Boolean)
    lblZona.Visible = fEnabled
    lblInicial.Visible = fEnabled
    lblVerificação.Visible = fEnabled
    lblJanela.Visible = fEnabled
    txtNumeroZona.Visible = fEnabled
    lstInicial.Visible = fEnabled
    lstCheck.Visible = fEnabled
    lstJanela.Visible = fEnabled
    frameLogica.Visible = fEnabled
End Sub

Private Sub Status_Controls(fEnabled As Boolean)
'
'  Coloca todos os controles do Form no estado "habilitado",
'  ou "desabilitado", isto é, permite a edição ou não.
'
   Dim c As Control
   For Each c In Controls
      If TypeOf c Is TextBox Or _
         TypeOf c Is ComboBox Or _
         TypeOf c Is CheckBox Or _
         TypeOf c Is OptionButton Then
         c.Enabled = fEnabled
      End If
   Next
   lstDescrZona.Enabled = True      'Always
End Sub

Private Sub Clear_Register()
   changeOff = True
   Dim c As Control
   For Each c In Controls
      If TypeOf c Is TextBox Then
         c.Text = ""
      ElseIf TypeOf c Is CheckBox Then
         c.Value = vbUnchecked
      ElseIf TypeOf c Is ComboBox Then
         If c.name <> "lstDescrZona" Then
            c.ListIndex = -1
         End If
      End If
   Next
   changeOff = False
End Sub

Private Function Update_Register() As Boolean
    Update_Register = True
    Dim lInsert As Boolean
    Dim txtStr As String
    lInsert = (fEditMode = adEditAdd)
    Dim lTipoDevice As Integer
    If IsNumeric(txtTipoDevice) Then
        lTipoDevice = txtTipoDevice
    Else
        lTipoDevice = 0
    End If
    Dim lGrupo As Integer
    If IsNumeric(txtGrupo) Then
        lGrupo = txtGrupo
    Else
        lGrupo = 0
    End If
    Dim lUID As String
    Dim lPTI As String
    If lTipoDevice = 8 Then
        'Receiver
'        txtUID = "00" & txtUID
        lPTI = "Receptor"
    ElseIf lTipoDevice = 7 Then
       'Repeater
'        txtUID = "01" & txtUID
        lPTI = "Repetidor"
    Else 'Sensor
'        txtUID = "B2" & txtUID
        lPTI = PTI
    End If
    lUID = txtUID
    Dim lReceptor As Byte
    If IsNumeric(txtReceptor) Then
        lReceptor = txtReceptor
    Else
        lReceptor = 0
    End If
    Dim lInicialZona As Long
    If IsNumeric(txtInicialZona) Then
        lInicialZona = txtInicialZona
    Else
        lInicialZona = 1
    End If
    Dim lCheck As Byte
    If IsNumeric(txtCheck) Then
        lCheck = txtCheck
    Else
        lCheck = 0
    End If
    Dim lJanela As Byte
    If IsNumeric(txtJanela) Then
        lJanela = txtJanela
    Else
        lJanela = 40
    End If
    Dim lTipoLogica As Byte
    If IsNumeric(txtTipoLogica) Then
        lTipoLogica = txtTipoLogica
    Else
        lTipoLogica = 0
    End If
    Dim lNumeroLogica As Byte
    If IsNumeric(txtNumeroLogica) Then
        lNumeroLogica = txtNumeroLogica
    Else
        lNumeroLogica = 1
    End If
    Dim lAtividade As Boolean, lTempo As Integer
    If chkAtividade.Value = vbUnchecked Then
        lAtividade = False
        fraAtividade.Visible = False
        lTempo = 0
    Else
        lAtividade = True
        fraAtividade.Visible = True
        lTempo = txtTempo
    End If
      
   Dim lCritico As Boolean, lPopup As Boolean, lTelaCheia As Boolean, lColor As Integer
   If chkPopup.Value = vbUnchecked Then
      lPopup = False
   Else
      lPopup = True
   End If
   If chkCritico.Value = vbUnchecked Then
      lCritico = False
      lColor = 4 'vbWhite
   Else
      lCritico = True
      If optColor(0) Then
         lColor = 0 'vbRed
      ElseIf optColor(1) Then
         lColor = 1 'vbBlue
      ElseIf optColor(2) Then
         lColor = 2 'vbYellow
      Else
         lColor = 3 'vbGreen
      End If
   End If
   
    lTelaCheia = (chkTelaCheia.Value = vbChecked)
    
    If lInsert Then
        txtStr = "INSERT INTO Sensor (fk_Entity, Numero_Sensor, Serial_Number, " & _
                "Local_Sensor, Local_Logica, Arquivo, UID, PTI, " & _
                "Receptor, Tipo_Sensor, fk_grupo, Inicial_Sensor, Check_Sensor, " & _
                "Janela_Sensor, Tipo_Logica, Numero_Logica, " & "chk_Atividade, chk_Tempo, " & _
                "popup, critico, color, servidor, camera, monitor, telacheia, user_cftv, senha)" & _
                " VALUES (" & fEntity.vId & np & txtNumeroZona & rp & txtSerialNumber & bp & _
                txtLocalZona & bp & txtLocalLogica & bp & Trim(txtArquivo) & bp & lUID & bp & lPTI & lp & _
                lReceptor & np & lTipoDevice & np & lGrupo & np & lInicialZona & np & lCheck & np & lJanela & np & lTipoLogica & np & _
                lNumeroLogica & np & lAtividade & np & lTempo & ", " & lPopup & ", " & lCritico & ", " & lColor & _
                ", '" & txtServerAddress & "', '" & txtCamera & "', '" & txtMonitor & "', " & lTelaCheia & _
                ", '" & txtUser & "', '" & txtPasswd & "')"
        oCnn.ExecSp txtStr
        Make_Service "Inserção do Sensor SN= " & txtSerialNumber.Text, strAccess(m_tAccess) & m_sUser
    Else
        txtStr = "UPDATE Sensor SET Numero_Sensor=" & txtNumeroZona & ", Serial_Number='" & txtSerialNumber & _
                "', Local_Sensor='" & txtLocalZona & "', Local_Logica='" & txtLocalLogica & _
                "', Arquivo='" & Trim(txtArquivo) & " ', UID='" & lUID & "', PTI='" & lPTI & _
                "', Receptor=" & lReceptor & ", Tipo_Sensor=" & lTipoDevice & ", Inicial_Sensor=" & _
                lInicialZona & ", Check_Sensor=" & lCheck & ", fk_grupo=" & lGrupo & _
                ", Janela_Sensor=" & lJanela & ", Tipo_Logica=" & lTipoLogica & ", Numero_Logica=" & _
                lNumeroLogica & ", chk_Atividade=" & lAtividade & ", chk_Tempo=" & lTempo & _
                ", popup=" & lPopup & ", critico=" & lCritico & ", color=" & lColor & ", servidor='" & _
                txtServerAddress & "', camera='" & txtCamera & "', monitor='" & txtMonitor & _
                "', telaCheia=" & lTelaCheia & ", user_cftv='" & txtUser & "', senha='" & txtPasswd & _
                "' WHERE (Serial_Number ='" & rsSensor("Serial_Number") & "')"
        oCnn.ExecSp txtStr
        Make_Service "Alteração no Sensor SN= " & txtSerialNumber.Text, strAccess(m_tAccess) & m_sUser
   End If
      
   If lInsert Then
      Dim fModule As New clsModule
      With fModule
         .Serial_Number = txtSerialNumber
         .UID = lUID
         .mNumero = txtNumeroZona
         .mLocal = txtLocalZona
         .mEntity = fEntity.vId
         .mTipo = lTipoDevice
         .SInicial = lInicialZona
         .mCheck = lCheck
         .mJanela = lJanela
         .mLogica = lTipoLogica
         .mNumLogica = lNumeroLogica
         .mLocalLogica = txtLocalLogica
         .mArquivo = rsSensor("Arquivo")
         .mLastAtiv = rsSensor("Last_Ativ")
         .mChkAtiv = lCheck And (.SInicial <> stDesabilitada)
         .mtempoAtiv = lTempo * 60
         .popup = lPopup
         .critico = lCritico
         .crColor = lColor
         .ServerAddress = txtServerAddress
         .Camera = txtCamera
         .Monitor = txtMonitor
         .telaCheia = lTelaCheia
         .user = txtUser
         .senha = txtPasswd
         .grupo = lGrupo
      End With
      fEntity.Add fModule, lUID
      lstModule.Add Item:=fModule, Key:=lUID
   
   Else
     ModuleUpdate lInsert
   End If
   
End Function

Private Sub Load_Register()
    On Error GoTo LoadError
    Clear_Register
    txtNumeroZona = rsSensor("Numero_Sensor")
    If Len(txtNumeroZona) = 1 Then
        txtNumeroZona = "00" & txtNumeroZona
    ElseIf Len(txtNumeroZona) = 2 Then
        txtNumeroZona = "0" & txtNumeroZona
    End If
    txtTipoDevice = rsSensor("Tipo_Sensor")
    txtSerialNumber = rsSensor("Serial_Number")
    txtReceptor = rsSensor("Receptor")
    If Not IsNull(rsSensor("PTI")) Or rsSensor("PTI") <> "" Then
        PTI = rsSensor("PTI")
    Else
        PTI = ""
    End If
    If Not IsNull(rsSensor("Local_Sensor")) Then
        txtLocalZona = rsSensor("Local_Sensor")
    Else
        txtLocalZona = ""
    End If
    txtInicialZona = rsSensor("Inicial_Sensor")
    txtCheck = rsSensor("Check_Sensor")
    If txtCheck <> 0 Then
        txtJanela = rsSensor("Janela_Sensor")
    End If
    txtTipoLogica = rsSensor("Tipo_Logica")
    If txtTipoLogica <> 0 Then
        txtNumeroLogica = rsSensor("Numero_Logica")
        If Len(txtNumeroLogica) = 1 Then
            txtNumeroLogica = "00" & txtNumeroLogica
        ElseIf Len(txtNumeroLogica) = 2 Then
            txtNumeroLogica = "0" & txtNumeroLogica
        End If
        If Not IsNull(rsSensor("Local_Logica")) Then
            txtLocalLogica = rsSensor("Local_Logica")
        End If
    Else
        txtNumeroLogica = ""
        txtLocalLogica = ""
    End If
    If Not IsNull(rsSensor("Arquivo")) Then
        txtArquivo = rsSensor("Arquivo")
    Else
        txtArquivo = ""
    End If
    If rsSensor("chk_Atividade") Then
        chkAtividade.Value = vbChecked
        fraAtividade.Visible = True
        txtTempo = rsSensor("chk_Tempo")
    Else
        chkAtividade.Value = vbUnchecked
        fraAtividade.Visible = False
    End If
   
    If rsSensor("popup") Then
        chkPopup.Value = vbChecked
    Else
        chkPopup.Value = vbUnchecked
    End If
    
    If rsSensor("critico") Then
        chkCritico.Value = vbChecked
        imgTratamento.Visible = True
        fraColor.Visible = True
        optColor(rsSensor("color")).Value = True
        txtColor = rsSensor("color")
    Else
        chkCritico.Value = vbUnchecked
        imgTratamento.Visible = False
        fraColor.Visible = False
    End If
   
    If rsSensor("telaCheia") Then
        chkTelaCheia.Value = vbChecked
    Else
        chkTelaCheia.Value = vbUnchecked
    End If
    
    If Not IsNull(rsSensor("servidor")) Then
        txtServerAddress = rsSensor("servidor")
    Else
        txtServerAddress = ""
    End If
    If Not IsNull(rsSensor("camera")) Then
        txtCamera = rsSensor("camera")
    Else
        txtCamera = ""
    End If
    If Not IsNull(rsSensor("monitor")) Then
        txtMonitor = rsSensor("monitor")
    Else
        txtMonitor = ""
    End If
    If Not IsNull(rsSensor("user_cftv")) Then
        txtUser = rsSensor("user_cftv")
    Else
        txtUser = ""
    End If
    If Not IsNull(rsSensor("senha")) Then
        txtPasswd = rsSensor("senha")
    Else
        txtPasswd = ""
    End If
    
    txtGrupo = rsSensor("fk_grupo")
    
   If txtUID <> rsSensor("UID") Then
      Err.Raise invalidLoadReg
   Else
      Set fModule = lstModule.Item(CStr(txtUID))
   End If
   Exit Sub
   
LoadError:
   Select Case Err.Number
      Case invalidLoadReg
         'there are some inconsistency with database and Usystem
         'Update the Database and fix the objects on Usystem
         On Error Resume Next
         Set fModule = lstModule.Item(rsSensor("UID"))
         lstModule.Remove fModule.UID
         fEntity.Remove fModule
         fModule.UID = txtUID
         lstModule.Add Item:=fModule, Key:=txtUID
         fEntity.Add fModule, txtUID
         Dim txtStr As String
         txtStr = "UPDATE Sensor SET UID='" & txtUID & "' WHERE (Serial_Number ='" & rsSensor("Serial_Number") & "')"
         oCnn.ExecSp txtStr
         Sync2Sensor
         On Error GoTo 0
      Case Else
         Err.Raise invalidUpdate
   End Select
End Sub

Private Sub ModuleUpdate(fInsertMode As Boolean)
    If fInsertMode Then
       Sync2Sensor
       Set fModule = New clsModule
'   Else
'      SyncSensor txtSerialNumber
'      Set fModule = lstModule.Item(txtUID)
'      fEntity.ClearStatus fModule
'   End If
        With fModule
           .Serial_Number = rsSensor("Serial_Number")
           .UID = rsSensor("UID")
           .mNumero = rsSensor("Numero_Sensor")
           .mLocal = rsSensor("Local_Sensor")
           .mEntity = rsSensor("fk_Entity")
           .mTipo = rsSensor("Tipo_Sensor")
           .SInicial = rsSensor("Inicial_Sensor")
           .mCheck = rsSensor("Check_Sensor")
           .mJanela = rsSensor("Janela_Sensor")
           .mLogica = rsSensor("Tipo_Logica")
           .mNumLogica = rsSensor("Numero_Logica")
           .mLocalLogica = rsSensor("Local_Logica")
           .mArquivo = rsSensor("Arquivo")
           .mLastAtiv = rsSensor("Last_Ativ")
           .mChkAtiv = rsSensor("chk_Atividade") And (.SInicial <> stDesabilitada)
           .mtempoAtiv = rsSensor("chk_Tempo") * 60
           .popup = rsSensor("popup")
           .critico = rsSensor("critico")
           .crColor = rsSensor("color")
           .ServerAddress = rsSensor("servidor")
           .Camera = rsSensor("camera")
           .Monitor = rsSensor("monitor")
           .telaCheia = rsSensor("telaCheia")
           .user = rsSensor("user_cftv")
           .senha = rsSensor("senha")
        End With
'   If fInsertMode Then
        fEntity.Add fModule, rsSensor("UID")
        lstModule.Add Item:=fModule, Key:=rsSensor("UID")
    Else
        Dim lColor As Integer
        If optColor(0) Then
           lColor = 0 'vbRed
        ElseIf optColor(1) Then
           lColor = 1 'vbBlue
        ElseIf optColor(2) Then
           lColor = 2 'vbYellow
        Else
           lColor = 3 'vbGreen
        End If
        With fModule
            .popup = chkPopup.Value = vbChecked
            .critico = chkCritico.Value = vbChecked
            .telaCheia = chkTelaCheia.Value = vbChecked
            .ServerAddress = txtServerAddress
            .Camera = txtCamera
            .Monitor = txtMonitor
            .user = txtUser
            .senha = txtPasswd
            .crColor = lColor
        End With
        fEntity.UpdateStatus fModule
    End If
    fEntity.UpdateColor clearDisp:=False
    ForNet.Treat_Inativos
End Sub

Private Sub Default_Values()
    txtTipoDevice = 0
    txtTipoDevice = 0
    txtInicialZona = 1
    txtNumeroZona = 1
    txtCheck = 0
    txtTipoLogica = 0
    chkAtividade.Value = vbUnchecked
    fraAtividade.Visible = False
    chkPopup.Value = vbUnchecked
    chkCritico.Value = vbUnchecked
    chkTelaCheia.Value = vbUnchecked
    imgTratamento.Visible = False
    fraColor.Visible = False
    txtColor = 4   'vbWhite
    txtGrupo = 0
End Sub

Private Sub Verify_Zona_AND()
   Dim rsZAND As New ADODB.Recordset
   rsZAND.Open "SELECT Numero_Sensor, fk_Entity FROM Sensor WHERE (Numero_Sensor = " & CInt(txtNumeroLogica) & _
                         " AND fk_Entity = " & fEntity.vId & ");", cnDB
   If rsZAND.EOF Then
      Beep
      MsgBox "Não é permitido efetuar função AND com uma Zona ainda " & _
             "não cadastrada!", sxExclamation, sxProname
      txtNumeroLogica = ""
   End If
   rsZAND.Close
End Sub

'Procedure que atualiza a lista de Zonas após uma inserção na entidade.
Private Sub lstDescrZona_Refill()
   Dim lSensor As String
   lSensor = rsSensor("Serial_Number")
   lstDescrZona.Clear
   rsSensor.MoveFirst
   While Not rsSensor.EOF
      lstDescrZona.AddItem rsSensor("Local_Sensor")
      lstDescrZona.ItemData(lstDescrZona.NewIndex) = rsSensor("Numero_Sensor")
      rsSensor.MoveNext
   Wend
   Sync1Sensor lSensor
End Sub

Private Function Verify_Consistency() As Boolean
   Verify_Consistency = False
   If oCnn Is Nothing Then Set oCnn = New clsConnection
   Dim lds As New ADODB.Recordset
   'Fixa o preenchimento do número do receptor como sendo 1 (como sendo somente um Serial Receiver)
   'Permite no futuro colocar mais de um Serial Receiver
   txtReceptor = "1"
   'Verifica o preenchimento do número da zona. Não pode ser nulo, ou 0.
   'Os controles de preechimento dos campos não permitem caracteres diferente de numerais.
   If Len(txtNumeroZona) = 0 Then
       MsgBox "O número de Zona é obrigatório.", sxExclamation, sxProname
       Exit Function
   ElseIf txtNumeroZona = "000" Then
       MsgBox "O número de Zona deve ser entre 001 e 999.", sxExclamation, sxProname
       Exit Function
   End If
   'Verifica o preenchimento do número do Serial Number tem 8 dígitos.
   If Len(txtSerialNumber) <> 8 Then
       MsgBox "O Número de Série de 8 digitos é obrigatório.", sxExclamation, sxProname
       Exit Function
   End If
   'Verifica a existência de duplicidade de Dispositivos e Zonas no cadastro.
   
   'Duplicidade de Zona em um mesma entidade
'   If fEditMode = adEditAdd Then
'       'Quando for inserção
'       Set lds = oCnn.ExecSpGetRs("SELECT Numero_Sensor, fk_Entity FROM Sensor WHERE (Numero_Sensor = " & CInt(txtNumeroZona) & _
'                 " AND fk_Entity = " & fEntity.vId & ")")
'   Else
'       'Quando for edição
'       Set lds = oCnn.ExecSpGetRs("SELECT Numero_Sensor, fk_Entity, Serial_Number FROM Sensor WHERE (Numero_Sensor = " & CInt(txtNumeroZona) & _
'                 " AND fk_Entity = " & fEntity.vId & " AND Serial_Number <> '" & rsSensor("Serial_Number") & "')")
'   End If
'   If Not lds.EOF Then
'       MsgBox "Cadastramento duplicado não permitido." & vbCrLf & _
'               "Zona com o mesmo número já existente nesta entidade. ", sxExclamation, sxProname
'       lds.Close
'       Exit Function
'   End If
'   lds.Close
        
   'Verifica a duplicidade de Dispositivos na rede
   If fEditMode = adEditAdd Then
       'Quando for inserção
       Set lds = oCnn.ExecSpGetRs("SELECT Sensor.Serial_Number, Entity.Descr_Entity FROM Sensor INNER JOIN Entity " & _
                 "ON Sensor.fk_Entity = Entity.cp_Entity WHERE (Sensor.Serial_Number = '" & txtSerialNumber & "')")
   Else
       'Quando for edição
       Set lds = oCnn.ExecSpGetRs("SELECT Sensor.Serial_Number, Entity.Descr_Entity FROM Sensor INNER JOIN Entity " & _
                 "ON Sensor.fk_Entity = Entity.cp_Entity WHERE (Sensor.Serial_Number = '" & txtSerialNumber & _
                 "' AND Serial_Number <> '" & txtSerialNumber & "')")
   End If
   If Not lds.EOF Then
       MsgBox "Cadastramento duplicado não permitido." & vbCrLf & _
               "Dispositivo com o mesmo SNº já existente na entidade " & lds("Descr_Entity"), sxExclamation, sxProname
       lds.Close
       Exit Function
   End If
   lds.Close
         
   'Faz consistencia da escolha de dulpla verificação de evento (usado em sensores infra passivo)
   If txtCheck = vrDupla Then
       If Not IsNumeric(txtJanela) Then
           txtJanela = 40
       End If
   Else
       txtJanela = ""
       lstJanela.ListIndex = -1
   End If
   
   'Consistência aprovada
   Verify_Consistency = True
End Function

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

Private Sub LoadSensor()
   On Error Resume Next
   rsSensor.Close
   On Error GoTo 0
   DoEvents
   rsSensor.Open "SELECT * FROM Sensor WHERE (fk_Entity = " & fEntity.vId & ") ORDER BY Serial_Number ASC", cnDB
   If rsSensor.EOF And rsSensor.BOF Then
      rsSensor.Requery
   ElseIf rsSensor.EOF Then
      DoEvents
      rsSensor.MoveFirst
   End If
End Sub

Private Sub SyncSensor(ByVal fSensor As String)
   Dim lFound As Boolean
   lFound = False
   LoadSensor
   While Not lFound
      If rsSensor("Serial_Number") = fSensor Then
         lFound = True
      Else
         rsSensor.MoveNext
      End If
   Wend
   Load_Register
End Sub

Private Sub Sync1Sensor(ByVal fSensor As String)
   Dim lFound As Boolean
   lFound = False
   rsSensor.MoveFirst
   While Not lFound
      If rsSensor("Serial_Number") = fSensor Then
         lFound = True
      Else
         rsSensor.MoveNext
      End If
   Wend
End Sub

Private Sub Sync2Sensor()
   Dim i As Integer
   LoadSensor
   DoEvents
   DoEvents
   If rsSensor.BOF And rsSensor.EOF Then
      rsSensor.Requery
   ElseIf rsSensor.EOF Then
      On Error Resume Next
      rsSensor.MoveFirst
   End If
   On Error GoTo 0
   Dim lFound As Boolean
   lFound = False
   While Not lFound
      If rsSensor.EOF Then
         Err.Raise invalidDataSet
      End If
      If rsSensor("Numero_Sensor") = CInt(txtNumeroZona) Then
         If rsSensor("Serial_Number") = txtSerialNumber And rsSensor("Receptor") = txtReceptor Then
            lFound = True
         Else
            rsSensor.MoveNext
         End If
      Else
         rsSensor.MoveNext
      End If
   Wend
End Sub

Private Sub UpdateForm(ByVal fRefill As Boolean)
   On Error GoTo TreatUpdate
   If rsSensor.EOF Then
      Clear_Register
      mnuRemover.Enabled = False
      mnuAlterar.Enabled = False
      Set fModule = Nothing
      SetAppearence btnReg(0), False
      SetAppearence btnReg(4), False
      lstDescrZona.Clear
   Else
      If fRefill Then lstDescrZona_Refill
      Load_Register
      mnuRemover.Enabled = True
      mnuAlterar.Enabled = True
      SetAppearence btnReg(0), True
      SetAppearence btnReg(4), True
   End If
   fEditMode = adEditNone
   mnuCancelar.Enabled = False
   mnuSalvar.Enabled = False
   mnuInserir.Enabled = True
   mnuExit.Enabled = True
   
   Status_Controls False
   SetAppearence btnReg(1), False
   SetAppearence btnReg(3), True
   SetAppearence btnReg(2), False
   SetAppearence btnReg(5), True
   Exit Sub
   
TreatUpdate:
   Select Case Err.Number
      Case invalidUpdate
         Set fModule = New clsModule
         With fModule
            .Serial_Number = rsSensor("Serial_Number")
            .UID = rsSensor("UID")
            .mNumero = rsSensor("Numero_Sensor")
            .mLocal = rsSensor("Local_Sensor")
            .mEntity = rsSensor("fk_Entity")
            .mTipo = rsSensor("Tipo_Sensor")
            .SInicial = rsSensor("Inicial_Sensor")
            .mCheck = rsSensor("Check_Sensor")
            .mJanela = rsSensor("Janela_Sensor")
            .mLogica = rsSensor("Tipo_Logica")
            .mNumLogica = rsSensor("Numero_Logica")
            .mLocalLogica = rsSensor("Local_Logica")
            .mArquivo = rsSensor("Arquivo")
            .mLastAtiv = rsSensor("Last_Ativ")
            .mChkAtiv = rsSensor("chk_Atividade") And (.SInicial <> stDesabilitada)
            .mtempoAtiv = rsSensor("chk_Tempo") * 60
            .critico = rsSensor("critico")
            .crColor = rsSensor("color")
            On Error Resume Next
               .ServerAddress = rsSensor("servidor")
               .Camera = rsSensor("camera")
               .Monitor = rsSensor("monitor")
               .telaCheia = rsSensor("telaCheia")
               .user = rsSensor("user_cftv")
               .senha = rsSensor("senha")
               .popup = rsSensor("popup")
               .grupo = rsSensor("grupo")
            On Error GoTo 0
         End With
         fEntity.Add fModule, rsSensor("UID")
         lstModule.Add Item:=fModule, Key:=rsSensor("UID")
         fEntity.UpdateColor clearDisp:=False
         ForNet.Treat_Inativos
         Resume Next
      Case Else
         MsgBox Err.Description, sxInformation, sxProname
      End Select
   
End Sub

