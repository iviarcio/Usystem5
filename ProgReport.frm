VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{A4749554-0441-4E64-8A03-3323601631C7}#1.0#0"; "LaVolpeAlphaImg2.ocx"
Begin VB.Form frmProgR 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Programação de Abertura/Fechamento"
   ClientHeight    =   4020
   ClientLeft      =   2265
   ClientTop       =   2565
   ClientWidth     =   5865
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
   Icon            =   "ProgReport.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4020
   ScaleWidth      =   5865
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtKeepData 
      Height          =   375
      Left            =   3160
      TabIndex        =   46
      Top             =   3360
      Width           =   495
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3195
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5835
      _ExtentX        =   10292
      _ExtentY        =   5636
      _Version        =   393216
      Style           =   1
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Segunda a Sexta"
      TabPicture(0)   =   "ProgReport.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame3"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame8"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "  Sábado   "
      TabPicture(1)   =   "ProgReport.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame4"
      Tab(1).Control(1)=   "Frame5"
      Tab(1).Control(2)=   "Frame9"
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "  Domingo  "
      TabPicture(2)   =   "ProgReport.frx":047A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame6"
      Tab(2).Control(1)=   "Frame7"
      Tab(2).Control(2)=   "Frame10"
      Tab(2).ControlCount=   3
      Begin VB.Frame Frame10 
         Caption         =   "Limites de Abertura e Fechamento"
         ForeColor       =   &H00800000&
         Height          =   855
         Left            =   -74880
         TabIndex        =   41
         Top             =   2235
         Width           =   4800
         Begin MSMask.MaskEdBox mskLDomOpen 
            Height          =   285
            Left            =   945
            TabIndex        =   42
            Top             =   360
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   503
            _Version        =   393216
            MaxLength       =   8
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "##:##:##"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskLDomClose 
            Height          =   285
            Left            =   3240
            TabIndex        =   43
            Top             =   345
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   503
            _Version        =   393216
            MaxLength       =   8
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "##:##:##"
            PromptChar      =   "_"
         End
         Begin VB.Label Label16 
            Caption         =   "Fechamento:"
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   2175
            TabIndex        =   45
            Top             =   345
            Width           =   975
         End
         Begin VB.Label Label11 
            Caption         =   "Abertura:"
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   120
            TabIndex        =   44
            Top             =   360
            Width           =   735
         End
      End
      Begin VB.Frame Frame9 
         Caption         =   "Limites de Abertura e Fechamento"
         ForeColor       =   &H00800000&
         Height          =   855
         Left            =   -74865
         TabIndex        =   36
         Top             =   2250
         Width           =   4800
         Begin MSMask.MaskEdBox mskLSabOpen 
            Height          =   285
            Left            =   945
            TabIndex        =   37
            Top             =   360
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   503
            _Version        =   393216
            MaxLength       =   8
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "##:##:##"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskLSabClose 
            Height          =   285
            Left            =   3240
            TabIndex        =   38
            Top             =   345
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   503
            _Version        =   393216
            MaxLength       =   8
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "##:##:##"
            PromptChar      =   "_"
         End
         Begin VB.Label Label6 
            Caption         =   "Fechamento:"
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   2175
            TabIndex        =   40
            Top             =   345
            Width           =   975
         End
         Begin VB.Label Label5 
            Caption         =   "Abertura:"
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   120
            TabIndex        =   39
            Top             =   360
            Width           =   735
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "Limites de Abertura e Fechamento"
         ForeColor       =   &H00800000&
         Height          =   855
         Left            =   135
         TabIndex        =   31
         Top             =   2250
         Width           =   4800
         Begin MSMask.MaskEdBox mskLSegOpen 
            Height          =   285
            Left            =   945
            TabIndex        =   32
            Top             =   360
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   503
            _Version        =   393216
            MaxLength       =   8
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "##:##:##"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskLSegClose 
            Height          =   285
            Left            =   3240
            TabIndex        =   33
            Top             =   345
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   503
            _Version        =   393216
            MaxLength       =   8
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "##:##:##"
            PromptChar      =   "_"
         End
         Begin VB.Label Label3 
            Caption         =   "Abertura:"
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   120
            TabIndex        =   35
            Top             =   360
            Width           =   735
         End
         Begin VB.Label Label2 
            Caption         =   "Fechamento:"
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   2175
            TabIndex        =   34
            Top             =   345
            Width           =   975
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "Intervalo de Fechamento"
         ForeColor       =   &H00800000&
         Height          =   855
         Left            =   -74880
         TabIndex        =   24
         Top             =   1320
         Width           =   4815
         Begin MSMask.MaskEdBox mskDomCloseInit 
            Height          =   285
            Left            =   960
            TabIndex        =   25
            Top             =   360
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   503
            _Version        =   393216
            MaxLength       =   8
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "##:##:##"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskDomCloseEnd 
            Height          =   285
            Left            =   3240
            TabIndex        =   26
            Top             =   360
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   503
            _Version        =   393216
            MaxLength       =   8
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "##:##:##"
            PromptChar      =   "_"
         End
         Begin VB.Label Label20 
            Caption         =   "Final: "
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   2760
            TabIndex        =   28
            Top             =   360
            Width           =   615
         End
         Begin VB.Label Label19 
            Caption         =   "Inicio:"
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   360
            TabIndex        =   27
            Top             =   360
            Width           =   615
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Intervalo de Abertura"
         ForeColor       =   &H00800000&
         Height          =   855
         Left            =   -74880
         TabIndex        =   19
         Top             =   360
         Width           =   4815
         Begin MSMask.MaskEdBox mskDomOpenInit 
            Height          =   285
            Left            =   960
            TabIndex        =   20
            Top             =   360
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   503
            _Version        =   393216
            MaxLength       =   8
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "##:##:##"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskDomOpenEnd 
            Height          =   285
            Left            =   3240
            TabIndex        =   21
            Top             =   360
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   503
            _Version        =   393216
            MaxLength       =   8
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "##:##:##"
            PromptChar      =   "_"
         End
         Begin VB.Label Label18 
            Caption         =   "Inicio:"
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   360
            TabIndex        =   23
            Top             =   360
            Width           =   615
         End
         Begin VB.Label Label17 
            Caption         =   "Final: "
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   2760
            TabIndex        =   22
            Top             =   360
            Width           =   615
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Intervalo de Fechamento"
         ForeColor       =   &H00800000&
         Height          =   855
         Left            =   -74880
         TabIndex        =   14
         Top             =   1320
         Width           =   4815
         Begin MSMask.MaskEdBox mskSabCloseInit 
            Height          =   285
            Left            =   960
            TabIndex        =   15
            Top             =   360
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   503
            _Version        =   393216
            MaxLength       =   8
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "##:##:##"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskSabCloseEnd 
            Height          =   285
            Left            =   3240
            TabIndex        =   16
            Top             =   360
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   503
            _Version        =   393216
            MaxLength       =   8
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "##:##:##"
            PromptChar      =   "_"
         End
         Begin VB.Label Label15 
            Caption         =   "Final: "
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   2760
            TabIndex        =   18
            Top             =   360
            Width           =   615
         End
         Begin VB.Label Label14 
            Caption         =   "Inicio:"
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   360
            TabIndex        =   17
            Top             =   360
            Width           =   615
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Intervalo de Abertura"
         ForeColor       =   &H00800000&
         Height          =   855
         Left            =   -74880
         TabIndex        =   11
         Top             =   360
         Width           =   4815
         Begin MSMask.MaskEdBox mskSabOpenInit 
            Height          =   285
            Left            =   960
            TabIndex        =   12
            Top             =   360
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   503
            _Version        =   393216
            MaxLength       =   8
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "##:##:##"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskSabOpenEnd 
            Height          =   285
            Left            =   3240
            TabIndex        =   29
            Top             =   360
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   503
            _Version        =   393216
            MaxLength       =   8
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "##:##:##"
            PromptChar      =   "_"
         End
         Begin VB.Label Label12 
            Caption         =   "Final: "
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   2760
            TabIndex        =   30
            Top             =   360
            Width           =   615
         End
         Begin VB.Label Label13 
            Caption         =   "Inicio:"
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   360
            TabIndex        =   13
            Top             =   360
            Width           =   615
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Intervalo de Fechamento"
         ForeColor       =   &H00800000&
         Height          =   855
         Left            =   120
         TabIndex        =   6
         Top             =   1320
         Width           =   4815
         Begin MSMask.MaskEdBox mskSegCloseInit 
            Height          =   285
            Left            =   960
            TabIndex        =   7
            Top             =   360
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   503
            _Version        =   393216
            MaxLength       =   8
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "##:##:##"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskSegCloseEnd 
            Height          =   285
            Left            =   3240
            TabIndex        =   8
            Top             =   360
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   503
            _Version        =   393216
            MaxLength       =   8
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "##:##:##"
            PromptChar      =   "_"
         End
         Begin VB.Label Label10 
            Caption         =   "Inicio:"
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   360
            TabIndex        =   10
            Top             =   360
            Width           =   615
         End
         Begin VB.Label Label9 
            Caption         =   "Final: "
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   2760
            TabIndex        =   9
            Top             =   360
            Width           =   615
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Intervalo de Abertura"
         ForeColor       =   &H00800000&
         Height          =   855
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   4815
         Begin MSMask.MaskEdBox mskSegOpenInit 
            Height          =   285
            Left            =   960
            TabIndex        =   2
            Top             =   360
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   503
            _Version        =   393216
            MaxLength       =   8
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "##:##:##"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskSegOpenEnd 
            Height          =   285
            Left            =   3240
            TabIndex        =   3
            Top             =   360
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   503
            _Version        =   393216
            MaxLength       =   8
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "##:##:##"
            PromptChar      =   "_"
         End
         Begin VB.Label Label8 
            Caption         =   "Final: "
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   2760
            TabIndex        =   5
            Top             =   360
            Width           =   615
         End
         Begin VB.Label Label7 
            Caption         =   "Inicio:"
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   360
            TabIndex        =   4
            Top             =   360
            Width           =   615
         End
      End
   End
   Begin VB.Label Label4 
      Caption         =   "dias"
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
      Height          =   375
      Left            =   3840
      TabIndex        =   48
      Top             =   3360
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Manter dados de eventos por "
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
      Height          =   375
      Left            =   120
      TabIndex        =   47
      Top             =   3360
      Width           =   3015
   End
   Begin LaVolpeAlphaImg.AlphaImgCtl cmdExit 
      Height          =   720
      Left            =   5040
      ToolTipText     =   "Fechar Programação de Abertura/Fechamento"
      Top             =   3240
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   1270
      Image           =   "ProgReport.frx":0496
      Effects         =   "ProgReport.frx":119B
   End
End
Attribute VB_Name = "frmProgR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private f_bChange As Boolean

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

Private Sub Form_Activate()
   Dim lrs As New ADODB.Recordset
   On Error Resume Next
   lrs.Open "SELECT * FROM Horario", cnDB, adOpenStatic, adLockReadOnly
   If Not lrs.EOF Then
      mskSegOpenInit = lrs("SegSexOpenInit")
      mskSegOpenEnd = lrs("SegSexOpenEnd")
      mskSegCloseInit = lrs("SegSexCloseInit")
      mskSegCloseEnd = lrs("SegSexCloseEnd")
      mskSabOpenInit = lrs("SabOpenInit")
      mskSabOpenEnd = lrs("SabOpenEnd")
      mskSabCloseInit = lrs("SabCloseInit")
      mskSabCloseEnd = lrs("SabCloseEnd")
      mskDomOpenInit = lrs("DomOpenInit")
      mskDomOpenEnd = lrs("DomOpenEnd")
      mskDomCloseInit = lrs("DomCloseInit")
      mskDomCloseEnd = lrs("DomCloseEnd")
      mskLSegOpen = lrs("TransLSegOpen")
      mskLSegClose = lrs("TransLSegClose")
      mskLSabOpen = lrs("TransLSabOpen")
      mskLSabClose = lrs("TransLSabClose")
      mskLDomOpen = lrs("TransLDomOpen")
      mskLDomClose = lrs("TransLDomClose")
      txtKeepData = lrs("KeepData")
   Else
      MsgBox sxDatabase + Chr$(13) + sxContact, sxCritical, sxProname
   End If
   lrs.Close
   f_bChange = False
End Sub

Private Sub Form_Load()
   Dim success As Long
   success = SetWindowPos(frmProgR.hWnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
   curhwnd = frmProgR.hWnd
   Left = (Screen.Width - Width) / 2   ' Center form horizontally.
   Top = (Screen.Height - Height) / 2  ' Center form vertically.
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If f_bChange Then
      Dim lcmd As New ADODB.Command
      Set lcmd.ActiveConnection = cnDB
      lcmd.CommandType = adCmdText
      lcmd.CommandText = "UPDATE Horario SET SegSexOpenInit=#" & mskSegOpenInit & _
            "#, SegSexOpenEnd=#" & mskSegOpenEnd & "#, SegSexCloseInit=#" & _
            mskSegCloseInit & "#, SegSexCloseEnd=#" & mskSegCloseEnd & _
            "#, SabOpenInit=#" & mskSabOpenInit & "#, SabOpenEnd=#" & _
            mskSabOpenEnd & "#, SabCloseInit=#" & mskSabCloseInit & _
            "#, SabCloseEnd=#" & mskSabCloseEnd & "#, DomOpenInit=#" & _
            mskDomOpenInit & "#, DomOpenEnd=#" & mskDomOpenEnd & "#, " & _
            "DomCloseInit=#" & mskDomCloseInit & "#, DomCloseEnd=#" & mskDomCloseEnd & _
            "#, TransLSegOpen=#" & mskLSegOpen & "#, TransLSegClose=#" & mskLSegClose & _
            "#, TransLSabOpen=#" & mskLSabOpen & "#, TransLSabClose=#" & mskLSabClose & _
            "#, TransLDomOpen=#" & mskLDomOpen & "#, TransLDomClose=#" & mskLDomClose & _
            "#, KeepData=" & m_iEvKeep
      lcmd.Execute
      Data_CleanUp
      Make_Service "Alteração nos horários de Abertura e Fechamento", strAccess(m_tAccess) & m_sUser
   End If
End Sub

Private Sub mskDomCloseEnd_Change()
   f_bChange = True
End Sub

Private Sub mskDomCloseInit_Change()
   f_bChange = True
End Sub

Private Sub mskDomOpenEnd_Change()
   f_bChange = True
End Sub

Private Sub mskDomOpenInit_Change()
   f_bChange = True
End Sub

Private Sub mskSabCloseEnd_Change()
   f_bChange = True
End Sub

Private Sub mskSabCloseInit_Change()
   f_bChange = True
End Sub

Private Sub mskSabOpenEnd_Change()
   f_bChange = True
End Sub

Private Sub mskSabOpenInit_Change()
   f_bChange = True
End Sub

Private Sub mskSegCloseEnd_Change()
   f_bChange = True
End Sub

Private Sub mskSegCloseInit_Change()
   f_bChange = True
End Sub

Private Sub mskSegOpenEnd_Change()
   f_bChange = True
End Sub

Private Sub mskSegOpenInit_Change()
   f_bChange = True
End Sub

Private Sub mskLSegClose_Change()
   f_bChange = True
End Sub

Private Sub mskLSegOpen_Change()
   f_bChange = True
End Sub

Private Sub mskLSabClose_Change()
   f_bChange = True
End Sub

Private Sub mskLSabOpen_Change()
   f_bChange = True
End Sub
Private Sub mskLDomClose_Change()
   f_bChange = True
End Sub

Private Sub mskLDomOpen_Change()
   f_bChange = True
End Sub

Private Sub txtKeepData_Change()
   m_iEvKeep = txtKeepData
   f_bChange = True
End Sub
