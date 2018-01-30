VERSION 5.00
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Object = "{A4749554-0441-4E64-8A03-3323601631C7}#1.0#0"; "LaVolpeAlphaImg2.ocx"
Begin VB.Form frmEntity 
   AutoRedraw      =   -1  'True
   Caption         =   "Cadastro"
   ClientHeight    =   5550
   ClientLeft      =   2055
   ClientTop       =   3000
   ClientWidth     =   9930
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Entity.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5550
   ScaleWidth      =   9930
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1395
      Left            =   120
      TabIndex        =   8
      Top             =   960
      Width           =   9030
      Begin VB.TextBox txtDescr 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   1560
         MaxLength       =   70
         TabIndex        =   1
         Top             =   240
         Width           =   4455
      End
      Begin VB.TextBox txtResp 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   1560
         MaxLength       =   50
         TabIndex        =   3
         Top             =   840
         Width           =   4455
      End
      Begin VB.TextBox txtFone 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   7440
         MaxLength       =   20
         TabIndex        =   5
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox txtFone2 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   7440
         MaxLength       =   20
         TabIndex        =   7
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "Local:"
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
         Left            =   720
         TabIndex        =   0
         Top             =   300
         Width           =   735
      End
      Begin VB.Label Label4 
         Caption         =   "Responsável:"
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
         TabIndex        =   2
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label5 
         Caption         =   "Tel. local:"
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
         Left            =   6600
         TabIndex        =   4
         Top             =   300
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Tel. externo:"
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
         Left            =   6600
         TabIndex        =   6
         Top             =   840
         Width           =   735
      End
   End
   Begin TrueOleDBGrid80.TDBGrid tdbg1 
      Height          =   2960
      Left            =   120
      TabIndex        =   9
      Top             =   2500
      Width           =   9735
      _ExtentX        =   17171
      _ExtentY        =   5212
      _LayoutType     =   4
      _RowHeight      =   18
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   " Entity"
      Columns(0).DataField=   ""
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   " Zona"
      Columns(1).DataField=   ""
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   " Descrição"
      Columns(2).DataField=   ""
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   " Tipo"
      Columns(3).DataField=   ""
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   " Modo"
      Columns(4).DataField=   ""
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   " Status"
      Columns(5).DataField=   ""
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   6
      Splits(0)._UserFlags=   0
      Splits(0).ExtendRightColumn=   -1  'True
      Splits(0).MarqueeStyle=   3
      Splits(0).AllowRowSizing=   0   'False
      Splits(0).RecordSelectors=   0   'False
      Splits(0).RecordSelectorWidth=   688
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).ScrollBars=   2
      Splits(0).AllowColSelect=   0   'False
      Splits(0).AlternatingRowStyle=   -1  'True
      Splits(0).DividerColor=   14215660
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=6"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=1191"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1085"
      Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
      Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=256"
      Splits(0)._ColumnProps(6)=   "Column(0).Visible=0"
      Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(8)=   "Column(1).Width=1482"
      Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=1376"
      Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
      Splits(0)._ColumnProps(12)=   "Column(1)._ColStyle=256"
      Splits(0)._ColumnProps(13)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(14)=   "Column(2).Width=6297"
      Splits(0)._ColumnProps(15)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(16)=   "Column(2)._WidthInPix=6191"
      Splits(0)._ColumnProps(17)=   "Column(2)._EditAlways=0"
      Splits(0)._ColumnProps(18)=   "Column(2)._ColStyle=256"
      Splits(0)._ColumnProps(19)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(20)=   "Column(3).Width=2223"
      Splits(0)._ColumnProps(21)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(22)=   "Column(3)._WidthInPix=2117"
      Splits(0)._ColumnProps(23)=   "Column(3)._EditAlways=0"
      Splits(0)._ColumnProps(24)=   "Column(3)._ColStyle=256"
      Splits(0)._ColumnProps(25)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(26)=   "Column(4).Width=2619"
      Splits(0)._ColumnProps(27)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(28)=   "Column(4)._WidthInPix=2514"
      Splits(0)._ColumnProps(29)=   "Column(4)._EditAlways=0"
      Splits(0)._ColumnProps(30)=   "Column(4)._ColStyle=65792"
      Splits(0)._ColumnProps(31)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(32)=   "Column(5).Width=2752"
      Splits(0)._ColumnProps(33)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(34)=   "Column(5)._WidthInPix=2646"
      Splits(0)._ColumnProps(35)=   "Column(5)._EditAlways=0"
      Splits(0)._ColumnProps(36)=   "Column(5)._ColStyle=65792"
      Splits(0)._ColumnProps(37)=   "Column(5).Order=6"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   0
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=11.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Tahoma"
      PrintInfos(0).PageFooterFont=   "Size=11.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Tahoma"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      AllowUpdate     =   0   'False
      DataMode        =   4
      DefColWidth     =   0
      HeadLines       =   1
      FootLines       =   1
      RowDividerStyle =   0
      MultipleLines   =   0
      EmptyRows       =   -1  'True
      CellTipsWidth   =   0
      DeadAreaBackColor=   13160660
      RowDividerColor =   14215660
      RowSubDividerColor=   14215660
      DirectionAfterEnter=   1
      DirectionAfterTab=   1
      MaxRows         =   250000
      ViewColumnCaptionWidth=   0
      ViewColumnWidth =   0
      _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
      _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=0,.valignment=0,.bgcolor=&H80000005&"
      _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
      _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
      _StyleDefs(3)   =   ":id=0,.borderColor=&H0&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=29,.bold=0,.fontsize=1125,.italic=0"
      _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(8)   =   ":id=1,.fontname=Tahoma"
      _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=33"
      _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=30,.bold=0,.fontsize=1125,.italic=0"
      _StyleDefs(11)  =   ":id=2,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(12)  =   ":id=2,.fontname=Tahoma"
      _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=31"
      _StyleDefs(14)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(15)  =   "SelectedStyle:id=6,.parent=1,.namedParent=32"
      _StyleDefs(16)  =   "EditorStyle:id=7,.parent=1"
      _StyleDefs(17)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=34,.bgcolor=&H80000005&"
      _StyleDefs(18)  =   ":id=8,.fgcolor=&H8000000D&"
      _StyleDefs(19)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=35,.bgcolor=&H95FFFF&"
      _StyleDefs(20)  =   "OddRowStyle:id=10,.parent=1,.namedParent=36"
      _StyleDefs(21)  =   "RecordSelectorStyle:id=25,.parent=2,.namedParent=27"
      _StyleDefs(22)  =   "FilterBarStyle:id=28,.parent=1,.namedParent=48"
      _StyleDefs(23)  =   "Splits(0).Style:id=71,.parent=1,.bold=0,.fontsize=1125,.italic=0,.underline=0"
      _StyleDefs(24)  =   ":id=71,.strikethrough=0,.charset=0"
      _StyleDefs(25)  =   ":id=71,.fontname=Tahoma"
      _StyleDefs(26)  =   "Splits(0).CaptionStyle:id=80,.parent=4"
      _StyleDefs(27)  =   "Splits(0).HeadingStyle:id=72,.parent=2"
      _StyleDefs(28)  =   "Splits(0).FooterStyle:id=73,.parent=3"
      _StyleDefs(29)  =   "Splits(0).InactiveStyle:id=74,.parent=5"
      _StyleDefs(30)  =   "Splits(0).SelectedStyle:id=76,.parent=6"
      _StyleDefs(31)  =   "Splits(0).EditorStyle:id=75,.parent=7"
      _StyleDefs(32)  =   "Splits(0).HighlightRowStyle:id=77,.parent=8"
      _StyleDefs(33)  =   "Splits(0).EvenRowStyle:id=78,.parent=9"
      _StyleDefs(34)  =   "Splits(0).OddRowStyle:id=79,.parent=10"
      _StyleDefs(35)  =   "Splits(0).RecordSelectorStyle:id=26,.parent=25"
      _StyleDefs(36)  =   "Splits(0).FilterBarStyle:id=47,.parent=28"
      _StyleDefs(37)  =   "Splits(0).Columns(0).Style:id=84,.parent=71,.alignment=0,.locked=0"
      _StyleDefs(38)  =   "Splits(0).Columns(0).HeadingStyle:id=81,.parent=72,.alignment=0"
      _StyleDefs(39)  =   "Splits(0).Columns(0).FooterStyle:id=82,.parent=73,.alignment=3"
      _StyleDefs(40)  =   "Splits(0).Columns(0).EditorStyle:id=83,.parent=75"
      _StyleDefs(41)  =   "Splits(0).Columns(1).Style:id=92,.parent=71,.alignment=0,.locked=0"
      _StyleDefs(42)  =   "Splits(0).Columns(1).HeadingStyle:id=89,.parent=72,.alignment=0"
      _StyleDefs(43)  =   "Splits(0).Columns(1).FooterStyle:id=90,.parent=73,.alignment=3"
      _StyleDefs(44)  =   "Splits(0).Columns(1).EditorStyle:id=91,.parent=75"
      _StyleDefs(45)  =   "Splits(0).Columns(2).Style:id=96,.parent=71,.alignment=0,.locked=0"
      _StyleDefs(46)  =   "Splits(0).Columns(2).HeadingStyle:id=93,.parent=72,.alignment=0"
      _StyleDefs(47)  =   "Splits(0).Columns(2).FooterStyle:id=94,.parent=73,.alignment=3"
      _StyleDefs(48)  =   "Splits(0).Columns(2).EditorStyle:id=95,.parent=75"
      _StyleDefs(49)  =   "Splits(0).Columns(3).Style:id=100,.parent=71,.alignment=0,.locked=0"
      _StyleDefs(50)  =   "Splits(0).Columns(3).HeadingStyle:id=97,.parent=72,.alignment=0"
      _StyleDefs(51)  =   "Splits(0).Columns(3).FooterStyle:id=98,.parent=73,.alignment=3"
      _StyleDefs(52)  =   "Splits(0).Columns(3).EditorStyle:id=99,.parent=75"
      _StyleDefs(53)  =   "Splits(0).Columns(4).Style:id=104,.parent=71"
      _StyleDefs(54)  =   "Splits(0).Columns(4).HeadingStyle:id=101,.parent=72"
      _StyleDefs(55)  =   "Splits(0).Columns(4).FooterStyle:id=102,.parent=73"
      _StyleDefs(56)  =   "Splits(0).Columns(4).EditorStyle:id=103,.parent=75"
      _StyleDefs(57)  =   "Splits(0).Columns(5).Style:id=24,.parent=71"
      _StyleDefs(58)  =   "Splits(0).Columns(5).HeadingStyle:id=21,.parent=72"
      _StyleDefs(59)  =   "Splits(0).Columns(5).FooterStyle:id=22,.parent=73"
      _StyleDefs(60)  =   "Splits(0).Columns(5).EditorStyle:id=23,.parent=75"
      _StyleDefs(61)  =   "Named:id=29:Normal"
      _StyleDefs(62)  =   ":id=29,.parent=0"
      _StyleDefs(63)  =   "Named:id=30:Heading"
      _StyleDefs(64)  =   ":id=30,.parent=29,.valignment=2,.bgcolor=&H808000&,.fgcolor=&H80000012&"
      _StyleDefs(65)  =   ":id=30,.wraptext=-1"
      _StyleDefs(66)  =   "Named:id=31:Footing"
      _StyleDefs(67)  =   ":id=31,.parent=29,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(68)  =   "Named:id=32:Selected"
      _StyleDefs(69)  =   ":id=32,.parent=29,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(70)  =   "Named:id=33:Caption"
      _StyleDefs(71)  =   ":id=33,.parent=30,.alignment=2"
      _StyleDefs(72)  =   "Named:id=34:HighlightRow"
      _StyleDefs(73)  =   ":id=34,.parent=29,.bgcolor=&H80000008&,.fgcolor=&H80000005&"
      _StyleDefs(74)  =   "Named:id=35:EvenRow"
      _StyleDefs(75)  =   ":id=35,.parent=29,.bgcolor=&HFFFF&"
      _StyleDefs(76)  =   "Named:id=36:OddRow"
      _StyleDefs(77)  =   ":id=36,.parent=29"
      _StyleDefs(78)  =   "Named:id=27:RecordSelector"
      _StyleDefs(79)  =   ":id=27,.parent=30"
      _StyleDefs(80)  =   "Named:id=48:FilterBar"
      _StyleDefs(81)  =   ":id=48,.parent=29"
   End
   Begin LaVolpeAlphaImg.AlphaImgCtl btnCad 
      Height          =   720
      Index           =   3
      Left            =   2640
      ToolTipText     =   "Configurar"
      Top             =   120
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   1270
      Image           =   "Entity.frx":030A
      Effects         =   "Entity.frx":139E
   End
   Begin LaVolpeAlphaImg.AlphaImgCtl btnCad 
      Height          =   720
      Index           =   5
      Left            =   4320
      ToolTipText     =   "Últimos Eventos"
      Top             =   120
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   1270
      Image           =   "Entity.frx":13B6
      Effects         =   "Entity.frx":233C
   End
   Begin LaVolpeAlphaImg.AlphaImgCtl btnCad 
      Height          =   720
      Index           =   4
      Left            =   3480
      ToolTipText     =   "Status"
      Top             =   120
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   1270
      Image           =   "Entity.frx":2354
      Effects         =   "Entity.frx":34F0
   End
   Begin LaVolpeAlphaImg.AlphaImgCtl btnCad 
      Height          =   720
      Index           =   2
      Left            =   1800
      ToolTipText     =   "Cancelar"
      Top             =   120
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   1270
      Image           =   "Entity.frx":3508
      Effects         =   "Entity.frx":43BA
   End
   Begin LaVolpeAlphaImg.AlphaImgCtl btnCad 
      Height          =   720
      Index           =   1
      Left            =   960
      ToolTipText     =   "Atualizar"
      Top             =   120
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   1270
      Image           =   "Entity.frx":43D2
      Effects         =   "Entity.frx":5230
   End
   Begin LaVolpeAlphaImg.AlphaImgCtl btnCad 
      Height          =   720
      Index           =   0
      Left            =   120
      ToolTipText     =   "Alterar"
      Top             =   120
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   1270
      Image           =   "Entity.frx":5248
      Effects         =   "Entity.frx":6008
   End
   Begin LaVolpeAlphaImg.AlphaImgCtl btnCad 
      Height          =   720
      Index           =   8
      Left            =   6840
      ToolTipText     =   "Sair do Cadastro de Entidades"
      Top             =   120
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   1270
      Image           =   "Entity.frx":6020
      Effects         =   "Entity.frx":6D25
   End
   Begin LaVolpeAlphaImg.AlphaImgCtl btnCad 
      Height          =   720
      Index           =   6
      Left            =   5160
      ToolTipText     =   "Ativar todas as Zonas"
      Top             =   120
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   1270
      Image           =   "Entity.frx":6D3D
      Effects         =   "Entity.frx":7AA1
   End
   Begin LaVolpeAlphaImg.AlphaImgCtl btnCad 
      Height          =   720
      Index           =   7
      Left            =   6000
      ToolTipText     =   "Desativar todas as Zonas"
      Top             =   120
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   1270
      Image           =   "Entity.frx":7AB9
      Effects         =   "Entity.frx":8847
   End
   Begin VB.Menu mnuEntidade 
      Caption         =   "&Entidade"
      Begin VB.Menu mnuAlterar 
         Caption         =   "Editar"
      End
      Begin VB.Menu mnuUpdate 
         Caption         =   "Salvar"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuCancel 
         Caption         =   "Cancelar"
      End
      Begin VB.Menu mnuRemove 
         Caption         =   "Remover"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu mnuS3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Fechar"
         Shortcut        =   ^F
      End
   End
   Begin VB.Menu mnuOption 
      Caption         =   "&Opções"
      Begin VB.Menu mnuConfig 
         Caption         =   "Configurar Zonas"
      End
      Begin VB.Menu mnuStatus 
         Caption         =   "Relatar Status das Zonas"
      End
      Begin VB.Menu mnuLastEvent 
         Caption         =   "Últimos Eventos"
      End
      Begin VB.Menu mnuS1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAtivar 
         Caption         =   "Ativar Zonas"
         Begin VB.Menu mnuActIncendio 
            Caption         =   "Incêndio"
         End
         Begin VB.Menu mnuActIntrusao 
            Caption         =   "Intrusão"
         End
         Begin VB.Menu mnuActEmergencia 
            Caption         =   "Emergência"
         End
         Begin VB.Menu mnuActPanico 
            Caption         =   "Pânico"
         End
         Begin VB.Menu mnuActSistema 
            Caption         =   "Sistema"
         End
      End
      Begin VB.Menu mnuDesativ 
         Caption         =   "Desativar Zonas"
         Begin VB.Menu mnuDeacIncendio 
            Caption         =   "Incêndio"
         End
         Begin VB.Menu mnuDeacIntrusao 
            Caption         =   "Intrusão"
         End
         Begin VB.Menu mnuDeacEmergencia 
            Caption         =   "Emergência"
         End
         Begin VB.Menu mnuDeacPanico 
            Caption         =   "Pânico"
         End
         Begin VB.Menu mnuDeacSistema 
            Caption         =   "Sistema"
         End
      End
   End
   Begin VB.Menu mnuMonitor 
      Caption         =   "&Monitor"
      Begin VB.Menu mnuMDesativ 
         Caption         =   "Zonas Desativadas"
      End
      Begin VB.Menu mnuMAtivo 
         Caption         =   "Zonas Ativadas"
      End
   End
End
Attribute VB_Name = "frmEntity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private fChange As Boolean
Private mList As XArrayDB
Public fEntity As clsEntity

Private Sub SetAppearence(btn As AlphaImgCtl, flag As Boolean)
   If flag Then
      btn.GrayScale = lvicNoGrayScale
   Else
      btn.GrayScale = lvicGreenMask
   End If
   btn.Enabled = flag
End Sub

Private Sub btnCad_Click(Index As Integer)
   Select Case Index
      Case 0
         mnuAlterar_Click
      Case 1
         mnuUpdate_Click
      Case 2
         mnuCancel_Click
      Case 3
         mnuConfig_Click
      Case 4
         mnuStatus_Click
      Case 5
         mnuLastEvent_Click
      Case 6
         fEntity.Activate s_All
         mnuDesativ.Enabled = True
         mnuAtivar.Enabled = fEntity.hasZonasDesativadas
         Grid_Load
      Case 7
         qResponse = sxQNone
         fEntity.Deactivate s_All
         mnuDesativ.Enabled = fEntity.hasZonasAtivadas
         mnuAtivar.Enabled = True
         Grid_Load
      Case 8
         mnuExit_Click
   End Select
End Sub

Private Sub btnCad_MouseEnter(Index As Integer)
   If btnCad(Index).Enabled Then
      btnCad(Index).SetRedraw = False
      btnCad(Index).GrayScale = lvicSepia
      btnCad(Index).LightnessPct = -20
      btnCad(Index).SetRedraw = True
   End If
End Sub

Private Sub btnCad_MouseExit(Index As Integer)
   If btnCad(Index).Enabled Then
      btnCad(Index).SetRedraw = False
      btnCad(Index).GrayScale = lvicNoGrayScale
      btnCad(Index).LightnessPct = 0
      btnCad(Index).SetRedraw = True
   End If
End Sub

Private Sub Form_Activate()
   Dim success As Long
   success = SetWindowPos(frmEntity.hWnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
   curhwnd = frmEntity.hWnd
   Entity_Load
   Grid_Load
   Me.Height = 6360  '4020
   
   If reloadForm Then
      Call mnuConfig_Click
   End If
   
End Sub

Private Sub Form_Load()
   Left = (Screen.Width - Width) / 2   ' Center form horizontally.
   Top = (Screen.Height - Height) / 2  ' Center form vertically.
'   tdbg1.EvenRowStyle.BackColor = &H80FFFF
'   tdbg1.OddRowStyle.BackColor = &HC0FFFF
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If fChange Then
      fEntity.Dump stModified
      fChange = False
   End If
   If m_bUserUnload Then
      Cancel = True
      Me.Hide
   End If
End Sub

Private Sub mnuActEmergencia_Click()
   fEntity.Activate s_Emergencia
   Make_Service "Ativação dos Sensores de Emergência", strAccess(m_tAccess) & m_sUser, fEntity
   mnuDesativ.Enabled = True
   mnuAtivar.Enabled = fEntity.hasZonasDesativadas
   Grid_Load
End Sub

Private Sub mnuActIncendio_Click()
   fEntity.Activate s_Incendio
   Make_Service "Ativação dos Sensores de Incêndio", strAccess(m_tAccess) & m_sUser, fEntity
   mnuDesativ.Enabled = True
   mnuAtivar.Enabled = fEntity.hasZonasDesativadas
   Grid_Load
End Sub

Private Sub mnuActIntrusao_Click()
   fEntity.Activate s_Intrusao
   Make_Service "Ativação dos Sensores de Intrusão", strAccess(m_tAccess) & m_sUser, fEntity
   mnuDesativ.Enabled = True
   mnuAtivar.Enabled = fEntity.hasZonasDesativadas
   Grid_Load
End Sub

Private Sub mnuActPanico_Click()
   fEntity.Activate s_Panico
   Make_Service "Ativação dos Sensores de Pânico", strAccess(m_tAccess) & m_sUser, fEntity
   mnuDesativ.Enabled = True
   mnuAtivar.Enabled = fEntity.hasZonasDesativadas
   Grid_Load
End Sub

Private Sub mnuActSistema_Click()
   fEntity.Activate s_Sistema
   Make_Service "Ativação dos Sensores de Sistema", strAccess(m_tAccess) & m_sUser, fEntity
   mnuDesativ.Enabled = True
   mnuAtivar.Enabled = fEntity.hasZonasDesativadas
   Grid_Load
End Sub

Private Sub mnuAlterar_Click()
   Control_Settings True
End Sub

Private Sub mnuCancel_Click()
   Entity_Load
End Sub

Private Sub mnuConfig_Click()
   reloadForm = False
   Load frmRegister
   Set frmRegister.fEntity = fEntity   'retain the current Entity
   frmRegister.Caption = "Configuração de Zona - " & fEntity.vDescr
   frmRegister.Show
End Sub

Private Sub mnuDeacEmergencia_Click()
   qResponse = sxQNone
   fEntity.Deactivate s_Emergencia
   Make_Service "Desativação dos Sensores de Emergencia", strAccess(m_tAccess) & m_sUser, fEntity
   mnuDesativ.Enabled = fEntity.hasZonasAtivadas
   mnuAtivar.Enabled = True
   Grid_Load
End Sub

Private Sub mnuDeacIncendio_Click()
   qResponse = sxQNone
   fEntity.Deactivate s_Incendio
   Make_Service "Desativação dos Sensores de Incêndio", strAccess(m_tAccess) & m_sUser, fEntity
   mnuDesativ.Enabled = fEntity.hasZonasAtivadas
   mnuAtivar.Enabled = True
   Grid_Load
End Sub

Private Sub mnuDeacIntrusao_Click()
   qResponse = sxQNone
   fEntity.Deactivate s_Intrusao
   Make_Service "Desativação dos Sensores de Intrusão", strAccess(m_tAccess) & m_sUser, fEntity
   mnuDesativ.Enabled = fEntity.hasZonasAtivadas
   mnuAtivar.Enabled = True
   Grid_Load
End Sub

Private Sub mnuDeacPanico_Click()
   qResponse = sxQNone
   fEntity.Deactivate s_Panico
   Make_Service "Desativação dos Sensores de Pânico", strAccess(m_tAccess) & m_sUser, fEntity
   mnuDesativ.Enabled = fEntity.hasZonasAtivadas
   mnuAtivar.Enabled = True
   Grid_Load
End Sub

Private Sub mnuDeacSistema_Click()
   qResponse = sxQNone
   fEntity.Deactivate s_Sistema
   Make_Service "Desativação dos Sensores de Sistema", strAccess(m_tAccess) & m_sUser, fEntity
   mnuDesativ.Enabled = fEntity.hasZonasAtivadas
   mnuAtivar.Enabled = True
   Grid_Load
End Sub

Private Sub mnuExit_Click()
   Unload Me
End Sub

Private Sub mnuLastEvent_Click()
   frmLastEvents.fEntity = True
   frmLastEvents.Caption = "Últimos Eventos - " & fEntity.vDescr
   Set frmLastEvents.lastEvents = fEntity.localEvent
   frmLastEvents.NEntity = fEntity.vId
   frmLastEvents.Show
End Sub

Private Sub mnuMAtivo_Click()
   fEntity.ChangeSZona stAtivada
End Sub

Private Sub mnuMDesativ_Click()
   fEntity.ChangeSZona stDesativada
End Sub

Private Sub mnuRemove_Click()
   Set tEntity = fEntity
   ForNet.ActiveForm.Entity_Delete
   Unload Me
End Sub

Private Sub mnuStatus_Click()
   Load frmZonas
   Set frmZonas.lEntity = fEntity
   frmZonas.Show
End Sub

Private Sub mnuUpdate_Click()
   Entity_Update
   Control_Settings False
End Sub

Private Sub Entity_Load()
   With fEntity
      txtDescr = .vDescr
      txtResp = .vResp
      txtFone = .vTel1
      txtFone2 = .vTel2
      mnuAtivar.Enabled = .hasZonasDesativadas
      mnuDesativ.Enabled = .hasZonasAtivadas
      mnuActIncendio.Enabled = .hasModules(s_Incendio)
      mnuActIntrusao.Enabled = .hasModules(s_Intrusao)
      mnuActEmergencia.Enabled = .hasModules(s_Emergencia)
      mnuActPanico.Enabled = .hasModules(s_Panico)
      mnuActSistema.Enabled = .hasModules(s_Sistema)
      mnuDeacIncendio.Enabled = mnuActIncendio.Enabled
      mnuDeacIntrusao.Enabled = mnuActIntrusao.Enabled
      mnuDeacEmergencia.Enabled = mnuActEmergencia.Enabled
      mnuDeacPanico.Enabled = mnuActPanico.Enabled
      mnuDeacSistema.Enabled = mnuActSistema.Enabled
   End With
   Control_Settings False
End Sub

Private Sub Entity_Update()
   With fEntity
      If txtDescr <> "" Then
         .vDescr = txtDescr
      Else
         .vDescr = "Loja " & .vId
         txtDescr = .vDescr
      End If
      .vResp = txtResp
      .vTel1 = txtFone
      .vTel2 = txtFone2
   End With
   fEntity.Dump stModified
End Sub

Private Sub Control_Settings(ByVal fvalue As Boolean)
   txtDescr.Enabled = fvalue And (m_tAccess <> sxOperator)
   txtResp.Enabled = fvalue
   txtFone.Enabled = fvalue
   txtFone2.Enabled = fvalue
   
   mnuMonitor.Visible = (m_sUser = sxAuthor)
   mnuConfig.Visible = (m_tAccess <> sxOperator)
   mnuRemove.Visible = (m_tAccess <> sxOperator)
   
   mnuRemove.Enabled = True
   mnuConfig.Enabled = Not fvalue
   mnuAlterar.Enabled = Not fvalue
   mnuStatus.Enabled = Not fvalue
   mnuAtivar.Enabled = Not fvalue
   mnuDesativ.Enabled = Not fvalue
   mnuLastEvent.Enabled = Not fvalue
   mnuExit.Enabled = Not fvalue
   
   mnuCancel.Enabled = fvalue
   mnuUpdate.Enabled = fvalue
   
   Dim i As Integer
   For i = 4 To 8
      SetAppearence btnCad(i), Not fvalue
   Next i
   SetAppearence btnCad(0), Not fvalue
   SetAppearence btnCad(1), fvalue
   SetAppearence btnCad(2), fvalue
   SetAppearence btnCad(3), Not fvalue And (m_tAccess <> sxOperator)
   
'   Toolbar1.Buttons("Config").Visible = Not fvalue And (m_tAccess <> sxOperator)
'   Toolbar1.Buttons("Alterar").Enabled = Not fvalue
'   Toolbar1.Buttons("Status").Enabled = Not fvalue
'   Toolbar1.Buttons("AZ_All").Enabled = Not fvalue
'   Toolbar1.Buttons("DZ_All").Enabled = Not fvalue
'   Toolbar1.Buttons("LastEvents").Enabled = Not fvalue
'   Toolbar1.Buttons("Exit").Enabled = Not fvalue
'   Toolbar1.Buttons("Cancelar").Enabled = fvalue
'   Toolbar1.Buttons("Atualizar").Enabled = fvalue
      
End Sub

Private Sub Grid_Load()
   On Error Resume Next
   Set mList = Nothing
   On Error GoTo 0
   Set mList = New XArrayDB
   If fEntity.localModule.Count >= 1 Then
      ' Allocate space for rows, 6 columns
      mList.ReDim 0, fEntity.localModule.Count - 1, 0, 5
      Dim mRow As Integer
      mRow = 0
      Dim cM As clsModule
      For Each cM In fEntity.localModule
         With cM
            mList(mRow, 0) = CStr(.mEntity)
            mList(mRow, 1) = CStr(.mNumero)
            mList(mRow, 2) = .mLocal
            mList(mRow, 3) = strTipo(.mTipo)
            mList(mRow, 4) = strModo(.status)
            mList(mRow, 5) = strStatus(.SZona)
            mRow = mRow + 1
         End With
      Next
      mList.ReDim 0, mRow - 1, 0, 5
      mList.QuickSort 0, mRow - 1, 4, XORDER_ASCEND, XTYPE_STRING, 3, XORDER_ASCEND, XTYPE_STRING
   End If
   tdbg1.Array = mList
   tdbg1.ReBind
End Sub
