VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmCRonda 
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cadastro de Rondas"
   ClientHeight    =   4170
   ClientLeft      =   2265
   ClientTop       =   2565
   ClientWidth     =   8895
   ClipControls    =   0   'False
   Icon            =   "CRonda.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4170
   ScaleWidth      =   8895
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox pctPrint 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   4320
      Picture         =   "CRonda.frx":030A
      ScaleHeight     =   330
      ScaleWidth      =   1140
      TabIndex        =   25
      Top             =   3810
      Width           =   1140
      Begin VB.Label lblPrint 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Imprimir"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   60
         TabIndex        =   26
         ToolTipText     =   "Exibe o Relatório de Configuração de Ronda."
         Top             =   30
         Width           =   1005
      End
   End
   Begin TrueOleDBGrid80.TDBDropDown TDBDropDownRonda 
      Bindings        =   "CRonda.frx":14CC
      Height          =   1740
      Left            =   4785
      Negotiate       =   -1  'True
      TabIndex        =   24
      Top             =   1980
      Width           =   2115
      _ExtentX        =   3731
      _ExtentY        =   3069
      _LayoutType     =   4
      _RowHeight      =   17
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "Id"
      Columns(0).DataField=   "cp_Entity"
      Columns(0).DataWidth=   11
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "Local"
      Columns(1).DataField=   "Descr_Entity"
      Columns(1).DataWidth=   255
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   2
      Splits(0)._UserFlags=   0
      Splits(0).ExtendRightColumn=   -1  'True
      Splits(0).AnchorRightColumn=   -1  'True
      Splits(0).MarqueeStyle=   3
      Splits(0).AllowRowSizing=   0   'False
      Splits(0).RecordSelectors=   0   'False
      Splits(0).RecordSelectorWidth=   979
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).ScrollBars=   2
      Splits(0).DividerColor=   15790320
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=2"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=847"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=767"
      Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
      Splits(0)._ColumnProps(5)=   "Column(0).Visible=0"
      Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(7)=   "Column(0)._AlignLeft=0"
      Splits(0)._ColumnProps(8)=   "Column(1).Width=2725"
      Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=2646"
      Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
      Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
      Splits.Count    =   1
      AllowRowSizing  =   0   'False
      Appearance      =   0
      BorderStyle     =   1
      ColumnHeaders   =   0   'False
      DataMode        =   0
      DefColWidth     =   0
      Enabled         =   -1  'True
      HeadLines       =   1
      RowDividerStyle =   2
      LayoutName      =   ""
      LayoutFileName  =   ""
      LayoutURL       =   ""
      EmptyRows       =   -1  'True
      ListField       =   "Descr_Entity"
      DataField       =   ""
      IntegralHeight  =   0   'False
      FetchRowStyle   =   0   'False
      AlternatingRowStyle=   0   'False
      DataMember      =   ""
      ColumnFooters   =   0   'False
      FootLines       =   1
      DeadAreaBackColor=   14215660
      ValueTranslate  =   0   'False
      RowDividerColor =   15790320
      RowSubDividerColor=   15790320
      _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
      _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
      _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
      _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
      _StyleDefs(3)   =   ":id=0,.borderColor=&H0&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=29,.bgcolor=&HFFFFCC&,.bold=0,.fontsize=900"
      _StyleDefs(7)   =   ":id=1,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(8)   =   ":id=1,.fontname=Tahoma"
      _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=33"
      _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=30,.bgcolor=&HFF8000&"
      _StyleDefs(11)  =   ":id=2,.fgcolor=&HFFFFFF&,.bold=-1,.fontsize=975,.italic=0,.underline=0"
      _StyleDefs(12)  =   ":id=2,.strikethrough=0,.charset=0"
      _StyleDefs(13)  =   ":id=2,.fontname=Tahoma"
      _StyleDefs(14)  =   "FooterStyle:id=3,.parent=1,.namedParent=31"
      _StyleDefs(15)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(16)  =   "SelectedStyle:id=6,.parent=1,.namedParent=32"
      _StyleDefs(17)  =   "EditorStyle:id=7,.parent=1"
      _StyleDefs(18)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=34"
      _StyleDefs(19)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=35"
      _StyleDefs(20)  =   "OddRowStyle:id=10,.parent=1,.namedParent=36"
      _StyleDefs(21)  =   "RecordSelectorStyle:id=37,.parent=2,.namedParent=39"
      _StyleDefs(22)  =   "FilterBarStyle:id=40,.parent=1,.namedParent=42"
      _StyleDefs(23)  =   "Splits(0).Style:id=11,.parent=1"
      _StyleDefs(24)  =   "Splits(0).CaptionStyle:id=20,.parent=4"
      _StyleDefs(25)  =   "Splits(0).HeadingStyle:id=12,.parent=2"
      _StyleDefs(26)  =   "Splits(0).FooterStyle:id=13,.parent=3"
      _StyleDefs(27)  =   "Splits(0).InactiveStyle:id=14,.parent=5"
      _StyleDefs(28)  =   "Splits(0).SelectedStyle:id=16,.parent=6"
      _StyleDefs(29)  =   "Splits(0).EditorStyle:id=15,.parent=7"
      _StyleDefs(30)  =   "Splits(0).HighlightRowStyle:id=17,.parent=8"
      _StyleDefs(31)  =   "Splits(0).EvenRowStyle:id=18,.parent=9"
      _StyleDefs(32)  =   "Splits(0).OddRowStyle:id=19,.parent=10"
      _StyleDefs(33)  =   "Splits(0).RecordSelectorStyle:id=38,.parent=37"
      _StyleDefs(34)  =   "Splits(0).FilterBarStyle:id=41,.parent=40"
      _StyleDefs(35)  =   "Splits(0).Columns(0).Style:id=24,.parent=11"
      _StyleDefs(36)  =   "Splits(0).Columns(0).HeadingStyle:id=21,.parent=12"
      _StyleDefs(37)  =   "Splits(0).Columns(0).FooterStyle:id=22,.parent=13"
      _StyleDefs(38)  =   "Splits(0).Columns(0).EditorStyle:id=23,.parent=15"
      _StyleDefs(39)  =   "Splits(0).Columns(1).Style:id=28,.parent=11"
      _StyleDefs(40)  =   "Splits(0).Columns(1).HeadingStyle:id=25,.parent=12"
      _StyleDefs(41)  =   "Splits(0).Columns(1).FooterStyle:id=26,.parent=13"
      _StyleDefs(42)  =   "Splits(0).Columns(1).EditorStyle:id=27,.parent=15"
      _StyleDefs(43)  =   "Named:id=29:Normal"
      _StyleDefs(44)  =   ":id=29,.parent=0"
      _StyleDefs(45)  =   "Named:id=30:Heading"
      _StyleDefs(46)  =   ":id=30,.parent=29,.valignment=2,.bgcolor=&HFF8000&,.fgcolor=&HFFFFFF&"
      _StyleDefs(47)  =   ":id=30,.wraptext=-1,.bold=0,.fontsize=975,.italic=0,.underline=0"
      _StyleDefs(48)  =   ":id=30,.strikethrough=0,.charset=0"
      _StyleDefs(49)  =   ":id=30,.fontname=MS Sans Serif"
      _StyleDefs(50)  =   "Named:id=31:Footing"
      _StyleDefs(51)  =   ":id=31,.parent=29,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(52)  =   "Named:id=32:Selected"
      _StyleDefs(53)  =   ":id=32,.parent=29,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(54)  =   "Named:id=33:Caption"
      _StyleDefs(55)  =   ":id=33,.parent=30,.alignment=2"
      _StyleDefs(56)  =   "Named:id=34:HighlightRow"
      _StyleDefs(57)  =   ":id=34,.parent=29,.bgcolor=&H80000008&,.fgcolor=&H80000005&"
      _StyleDefs(58)  =   "Named:id=35:EvenRow"
      _StyleDefs(59)  =   ":id=35,.parent=29,.bgcolor=&HFFFF00&"
      _StyleDefs(60)  =   "Named:id=36:OddRow"
      _StyleDefs(61)  =   ":id=36,.parent=29"
      _StyleDefs(62)  =   "Named:id=39:RecordSelector"
      _StyleDefs(63)  =   ":id=39,.parent=30"
      _StyleDefs(64)  =   "Named:id=42:FilterBar"
      _StyleDefs(65)  =   ":id=42,.parent=29"
   End
   Begin MSAdodcLib.Adodc AdodcRonda 
      Height          =   330
      Left            =   90
      Top             =   4635
      Width           =   2265
      _ExtentX        =   3995
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "AdodcRonda"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.PictureBox pctSave 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   6600
      Picture         =   "CRonda.frx":14E5
      ScaleHeight     =   330
      ScaleWidth      =   1140
      TabIndex        =   22
      Top             =   3810
      Width           =   1140
      Begin VB.Label lblSave 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Salvar"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   75
         TabIndex        =   23
         Top             =   30
         Width           =   1005
      End
   End
   Begin VB.PictureBox pctEdit 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   5460
      Picture         =   "CRonda.frx":26A7
      ScaleHeight     =   330
      ScaleWidth      =   1140
      TabIndex        =   20
      Top             =   3810
      Width           =   1140
      Begin VB.Label lblEdit 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Editar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   60
         TabIndex        =   21
         Top             =   30
         Width           =   1005
      End
   End
   Begin VB.PictureBox pctExit 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   7740
      Picture         =   "CRonda.frx":3869
      ScaleHeight     =   330
      ScaleWidth      =   1140
      TabIndex        =   18
      Top             =   3810
      Width           =   1140
      Begin VB.Label lblExit 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Sair"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   60
         TabIndex        =   19
         Top             =   30
         Width           =   1005
      End
   End
   Begin VB.PictureBox pctRemove 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   345
      Left            =   2160
      Picture         =   "CRonda.frx":4A2B
      ScaleHeight     =   345
      ScaleWidth      =   2160
      TabIndex        =   16
      Top             =   3810
      Width           =   2160
      Begin VB.Label lblRemove 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Remover Percurso Percurso"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   60
         TabIndex        =   17
         Top             =   45
         Width           =   2025
      End
   End
   Begin VB.PictureBox pctInsert 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   345
      Left            =   15
      Picture         =   "CRonda.frx":6EDD
      ScaleHeight     =   345
      ScaleWidth      =   2145
      TabIndex        =   14
      Top             =   3810
      Width           =   2145
      Begin VB.Label lblInsert 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Inserir Percurso"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   30
         TabIndex        =   15
         Top             =   45
         Width           =   1965
      End
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   1545
      Left            =   2505
      TabIndex        =   1
      Top             =   90
      Width           =   6360
      _Version        =   65536
      _ExtentX        =   11218
      _ExtentY        =   2725
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
      Begin VB.CommandButton cmdDelete 
         Enabled         =   0   'False
         Height          =   315
         Left            =   5940
         Picture         =   "CRonda.frx":938F
         Style           =   1  'Graphical
         TabIndex        =   29
         ToolTipText     =   "Remove Horário"
         Top             =   435
         Width           =   330
      End
      Begin VB.CommandButton cmdOk 
         Enabled         =   0   'False
         Height          =   330
         Left            =   5565
         Picture         =   "CRonda.frx":97D1
         Style           =   1  'Graphical
         TabIndex        =   28
         ToolTipText     =   "Insere Horário"
         Top             =   420
         Width           =   330
      End
      Begin VB.ComboBox lstHorario 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   4350
         Locked          =   -1  'True
         Sorted          =   -1  'True
         TabIndex        =   27
         Top             =   420
         Width           =   1200
      End
      Begin VB.TextBox txtDesvio 
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
         Enabled         =   0   'False
         Height          =   300
         Left            =   4390
         TabIndex        =   11
         Top             =   1155
         Width           =   600
      End
      Begin VB.OptionButton optStatus 
         Caption         =   "Ativado"
         Enabled         =   0   'False
         ForeColor       =   &H00808000&
         Height          =   255
         Index           =   0
         Left            =   195
         TabIndex        =   8
         Top             =   405
         Width           =   1185
      End
      Begin VB.CheckBox chkSegSex 
         Caption         =   "Segunda à Sexta"
         Enabled         =   0   'False
         ForeColor       =   &H00808000&
         Height          =   300
         Left            =   1950
         TabIndex        =   7
         Top             =   405
         Width           =   1995
      End
      Begin VB.OptionButton optStatus 
         Caption         =   "Desativado"
         Enabled         =   0   'False
         ForeColor       =   &H00808000&
         Height          =   255
         Index           =   1
         Left            =   195
         TabIndex        =   6
         Top             =   772
         Width           =   1320
      End
      Begin VB.CheckBox chkDom 
         Caption         =   "Domingos e Feriados"
         Enabled         =   0   'False
         ForeColor       =   &H00808000&
         Height          =   300
         Left            =   1950
         TabIndex        =   3
         Top             =   1140
         Width           =   1995
      End
      Begin VB.CheckBox chkSab 
         Caption         =   "Sábado"
         Enabled         =   0   'False
         ForeColor       =   &H00808000&
         Height          =   300
         Left            =   1950
         TabIndex        =   2
         Top             =   765
         Width           =   1995
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   " -"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   4180
         TabIndex        =   31
         Top             =   1180
         Width           =   240
      End
      Begin VB.Label Label6 
         Caption         =   " +"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   4175
         TabIndex        =   30
         Top             =   1125
         Width           =   240
      End
      Begin VB.Label Label5 
         Caption         =   "(Minutos)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   5100
         TabIndex        =   12
         Top             =   1170
         Width           =   795
      End
      Begin VB.Label Label4 
         Caption         =   "Desvio"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   380
         Left            =   4200
         TabIndex        =   10
         Top             =   840
         Width           =   1680
      End
      Begin VB.Label Label3 
         Caption         =   "Horários"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4200
         TabIndex        =   9
         Top             =   45
         Width           =   1680
      End
      Begin VB.Label Label2 
         Caption         =   "Status"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   75
         TabIndex        =   5
         Top             =   15
         Width           =   990
      End
      Begin VB.Label Label1 
         Caption         =   "Opções"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1800
         TabIndex        =   4
         Top             =   15
         Width           =   1680
      End
   End
   Begin TrueOleDBGrid80.TDBGrid tdbgPercurso 
      Height          =   3690
      Left            =   75
      TabIndex        =   0
      Top             =   60
      Width           =   2370
      _ExtentX        =   4180
      _ExtentY        =   6509
      _LayoutType     =   4
      _RowHeight      =   17
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "Percursos"
      Columns(0).DataField=   ""
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "idPercurso"
      Columns(1).DataField=   ""
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "horario"
      Columns(2).DataField=   ""
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "desvio"
      Columns(3).DataField=   ""
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "valSegSex"
      Columns(4).DataField=   ""
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "valSab"
      Columns(5).DataField=   ""
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "valDom"
      Columns(6).DataField=   ""
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(7)._VlistStyle=   0
      Columns(7)._MaxComboItems=   5
      Columns(7).Caption=   "Status"
      Columns(7).DataField=   ""
      Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   8
      Splits(0)._UserFlags=   0
      Splits(0).AllowRowSizing=   0   'False
      Splits(0).RecordSelectors=   0   'False
      Splits(0).RecordSelectorWidth=   979
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).ScrollBars=   2
      Splits(0).AllowColSelect=   0   'False
      Splits(0).DividerColor=   15790320
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=8"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=4392"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=4313"
      Splits(0)._ColumnProps(4)=   "Column(0).AllowSizing=0"
      Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=260"
      Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(7)=   "Column(1).Width=2725"
      Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=2646"
      Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(11)=   "Column(2).Width=2725"
      Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=2646"
      Splits(0)._ColumnProps(14)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(15)=   "Column(3).Width=2725"
      Splits(0)._ColumnProps(16)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(17)=   "Column(3)._WidthInPix=2646"
      Splits(0)._ColumnProps(18)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(19)=   "Column(4).Width=2725"
      Splits(0)._ColumnProps(20)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(21)=   "Column(4)._WidthInPix=2646"
      Splits(0)._ColumnProps(22)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(23)=   "Column(5).Width=2725"
      Splits(0)._ColumnProps(24)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(25)=   "Column(5)._WidthInPix=2646"
      Splits(0)._ColumnProps(26)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(27)=   "Column(6).Width=2725"
      Splits(0)._ColumnProps(28)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(29)=   "Column(6)._WidthInPix=2646"
      Splits(0)._ColumnProps(30)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(31)=   "Column(7).Width=2725"
      Splits(0)._ColumnProps(32)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(33)=   "Column(7)._WidthInPix=2646"
      Splits(0)._ColumnProps(34)=   "Column(7).Order=8"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   0
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      AllowUpdate     =   0   'False
      Appearance      =   0
      BorderStyle     =   0
      DataMode        =   4
      DefColWidth     =   0
      HeadLines       =   1
      FootLines       =   1
      RowDividerStyle =   6
      MultipleLines   =   0
      EmptyRows       =   -1  'True
      CellTipsWidth   =   0
      MultiSelect     =   0
      DeadAreaBackColor=   16777152
      RowDividerColor =   15790320
      RowSubDividerColor=   15790320
      DirectionAfterEnter=   1
      DirectionAfterTab=   1
      MaxRows         =   250000
      ViewColumnCaptionWidth=   0
      ViewColumnWidth =   0
      _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
      _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
      _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
      _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
      _StyleDefs(3)   =   ":id=0,.borderColor=&H0&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=29,.bgcolor=&HFFFFE0&"
      _StyleDefs(7)   =   "CaptionStyle:id=4,.parent=2,.namedParent=33,.bgcolor=&H80000002&,.bold=-1"
      _StyleDefs(8)   =   ":id=4,.fontsize=975,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(9)   =   ":id=4,.fontname=Tahoma"
      _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=30,.bgcolor=&H808000&"
      _StyleDefs(11)  =   ":id=2,.fgcolor=&HFFFFFF&,.bold=0,.fontsize=975,.italic=0,.underline=0"
      _StyleDefs(12)  =   ":id=2,.strikethrough=0,.charset=0"
      _StyleDefs(13)  =   ":id=2,.fontname=Tahoma"
      _StyleDefs(14)  =   "FooterStyle:id=3,.parent=1,.namedParent=31"
      _StyleDefs(15)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&HFFB400&,.fgcolor=&H80000009&,.bold=-1"
      _StyleDefs(16)  =   ":id=5,.fontsize=975,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(17)  =   ":id=5,.fontname=Tahoma"
      _StyleDefs(18)  =   "SelectedStyle:id=6,.parent=1,.namedParent=32,.bgcolor=&HC0C000&"
      _StyleDefs(19)  =   "EditorStyle:id=7,.parent=1"
      _StyleDefs(20)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=34,.bgcolor=&H80000003&"
      _StyleDefs(21)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=35"
      _StyleDefs(22)  =   "OddRowStyle:id=10,.parent=1,.namedParent=36"
      _StyleDefs(23)  =   "RecordSelectorStyle:id=61,.parent=2,.namedParent=63"
      _StyleDefs(24)  =   "FilterBarStyle:id=64,.parent=1,.namedParent=66"
      _StyleDefs(25)  =   "Splits(0).Style:id=11,.parent=1"
      _StyleDefs(26)  =   "Splits(0).CaptionStyle:id=20,.parent=4,.bgcolor=&HFF8000&,.fgcolor=&HFFFFFF&"
      _StyleDefs(27)  =   "Splits(0).HeadingStyle:id=12,.parent=2,.bgcolor=&HFFFFE0&"
      _StyleDefs(28)  =   "Splits(0).FooterStyle:id=13,.parent=3"
      _StyleDefs(29)  =   "Splits(0).InactiveStyle:id=14,.parent=5"
      _StyleDefs(30)  =   "Splits(0).SelectedStyle:id=16,.parent=6,.bgcolor=&H808000&"
      _StyleDefs(31)  =   "Splits(0).EditorStyle:id=15,.parent=7"
      _StyleDefs(32)  =   "Splits(0).HighlightRowStyle:id=17,.parent=8"
      _StyleDefs(33)  =   "Splits(0).EvenRowStyle:id=18,.parent=9,.bgcolor=&H808000&"
      _StyleDefs(34)  =   "Splits(0).OddRowStyle:id=19,.parent=10"
      _StyleDefs(35)  =   "Splits(0).RecordSelectorStyle:id=62,.parent=61"
      _StyleDefs(36)  =   "Splits(0).FilterBarStyle:id=65,.parent=64"
      _StyleDefs(37)  =   "Splits(0).Columns(0).Style:id=24,.parent=11,.bold=0,.fontsize=975,.italic=0"
      _StyleDefs(38)  =   ":id=24,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(39)  =   ":id=24,.fontname=Tahoma"
      _StyleDefs(40)  =   "Splits(0).Columns(0).HeadingStyle:id=21,.parent=12,.alignment=0"
      _StyleDefs(41)  =   ":id=21,.bgcolor=&HFF8000&,.bold=0,.fontsize=975,.italic=0,.underline=0"
      _StyleDefs(42)  =   ":id=21,.strikethrough=0,.charset=0"
      _StyleDefs(43)  =   ":id=21,.fontname=Tahoma"
      _StyleDefs(44)  =   "Splits(0).Columns(0).FooterStyle:id=22,.parent=13"
      _StyleDefs(45)  =   "Splits(0).Columns(0).EditorStyle:id=23,.parent=15"
      _StyleDefs(46)  =   "Splits(0).Columns(1).Style:id=28,.parent=11"
      _StyleDefs(47)  =   "Splits(0).Columns(1).HeadingStyle:id=25,.parent=12"
      _StyleDefs(48)  =   "Splits(0).Columns(1).FooterStyle:id=26,.parent=13"
      _StyleDefs(49)  =   "Splits(0).Columns(1).EditorStyle:id=27,.parent=15"
      _StyleDefs(50)  =   "Splits(0).Columns(2).Style:id=52,.parent=11"
      _StyleDefs(51)  =   "Splits(0).Columns(2).HeadingStyle:id=49,.parent=12"
      _StyleDefs(52)  =   "Splits(0).Columns(2).FooterStyle:id=50,.parent=13"
      _StyleDefs(53)  =   "Splits(0).Columns(2).EditorStyle:id=51,.parent=15"
      _StyleDefs(54)  =   "Splits(0).Columns(3).Style:id=48,.parent=11"
      _StyleDefs(55)  =   "Splits(0).Columns(3).HeadingStyle:id=45,.parent=12"
      _StyleDefs(56)  =   "Splits(0).Columns(3).FooterStyle:id=46,.parent=13"
      _StyleDefs(57)  =   "Splits(0).Columns(3).EditorStyle:id=47,.parent=15"
      _StyleDefs(58)  =   "Splits(0).Columns(4).Style:id=44,.parent=11"
      _StyleDefs(59)  =   "Splits(0).Columns(4).HeadingStyle:id=41,.parent=12"
      _StyleDefs(60)  =   "Splits(0).Columns(4).FooterStyle:id=42,.parent=13"
      _StyleDefs(61)  =   "Splits(0).Columns(4).EditorStyle:id=43,.parent=15"
      _StyleDefs(62)  =   "Splits(0).Columns(5).Style:id=40,.parent=11"
      _StyleDefs(63)  =   "Splits(0).Columns(5).HeadingStyle:id=37,.parent=12"
      _StyleDefs(64)  =   "Splits(0).Columns(5).FooterStyle:id=38,.parent=13"
      _StyleDefs(65)  =   "Splits(0).Columns(5).EditorStyle:id=39,.parent=15"
      _StyleDefs(66)  =   "Splits(0).Columns(6).Style:id=56,.parent=11"
      _StyleDefs(67)  =   "Splits(0).Columns(6).HeadingStyle:id=53,.parent=12"
      _StyleDefs(68)  =   "Splits(0).Columns(6).FooterStyle:id=54,.parent=13"
      _StyleDefs(69)  =   "Splits(0).Columns(6).EditorStyle:id=55,.parent=15"
      _StyleDefs(70)  =   "Splits(0).Columns(7).Style:id=60,.parent=11"
      _StyleDefs(71)  =   "Splits(0).Columns(7).HeadingStyle:id=57,.parent=12"
      _StyleDefs(72)  =   "Splits(0).Columns(7).FooterStyle:id=58,.parent=13"
      _StyleDefs(73)  =   "Splits(0).Columns(7).EditorStyle:id=59,.parent=15"
      _StyleDefs(74)  =   "Named:id=29:Normal"
      _StyleDefs(75)  =   ":id=29,.parent=0"
      _StyleDefs(76)  =   "Named:id=30:Heading"
      _StyleDefs(77)  =   ":id=30,.parent=29,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(78)  =   ":id=30,.wraptext=-1"
      _StyleDefs(79)  =   "Named:id=31:Footing"
      _StyleDefs(80)  =   ":id=31,.parent=29,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(81)  =   "Named:id=32:Selected"
      _StyleDefs(82)  =   ":id=32,.parent=29,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(83)  =   "Named:id=33:Caption"
      _StyleDefs(84)  =   ":id=33,.parent=30,.alignment=2"
      _StyleDefs(85)  =   "Named:id=34:HighlightRow"
      _StyleDefs(86)  =   ":id=34,.parent=29,.bgcolor=&H80000008&,.fgcolor=&H80000005&"
      _StyleDefs(87)  =   "Named:id=35:EvenRow"
      _StyleDefs(88)  =   ":id=35,.parent=29,.bgcolor=&HFFFF00&"
      _StyleDefs(89)  =   "Named:id=36:OddRow"
      _StyleDefs(90)  =   ":id=36,.parent=29"
      _StyleDefs(91)  =   "Named:id=63:RecordSelector"
      _StyleDefs(92)  =   ":id=63,.parent=30"
      _StyleDefs(93)  =   "Named:id=66:FilterBar"
      _StyleDefs(94)  =   ":id=66,.parent=29"
   End
   Begin TrueOleDBGrid80.TDBGrid tdbgRonda 
      Height          =   1980
      Left            =   2520
      TabIndex        =   13
      Top             =   1740
      Width           =   6345
      _ExtentX        =   11192
      _ExtentY        =   3493
      _LayoutType     =   4
      _RowHeight      =   17
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "Pontos de Ronda"
      Columns(0).DataField=   ""
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "Id"
      Columns(1).DataField=   ""
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "Local"
      Columns(2).DataField=   ""
      Columns(2).DropDown=   "TDBDropDownRonda"
      Columns(2).DropDown.vt=   8
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "Intervalo"
      Columns(3).DataField=   ""
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   84
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "Ativado"
      Columns(4).DataField=   ""
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "idRonda"
      Columns(5).DataField=   ""
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "idPercurso"
      Columns(6).DataField=   ""
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   7
      Splits(0)._UserFlags=   0
      Splits(0).AllowRowSizing=   0   'False
      Splits(0).RecordSelectorWidth=   979
      Splits(0)._SavedRecordSelectors=   -1  'True
      Splits(0).ScrollBars=   2
      Splits(0).AllowColSelect=   0   'False
      Splits(0).DividerColor=   15790320
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=7"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=3334"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=3254"
      Splits(0)._ColumnProps(4)=   "Column(0).AllowSizing=0"
      Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=260"
      Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(7)=   "Column(1).Width=847"
      Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=767"
      Splits(0)._ColumnProps(10)=   "Column(1).Visible=0"
      Splits(0)._ColumnProps(11)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(12)=   "Column(2).Width=3201"
      Splits(0)._ColumnProps(13)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(14)=   "Column(2)._WidthInPix=3122"
      Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(16)=   "Column(2).DropDownList=1"
      Splits(0)._ColumnProps(17)=   "Column(3).Width=1773"
      Splits(0)._ColumnProps(18)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(19)=   "Column(3)._WidthInPix=1693"
      Splits(0)._ColumnProps(20)=   "Column(3)._ColStyle=1"
      Splits(0)._ColumnProps(21)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(22)=   "Column(4).Width=1826"
      Splits(0)._ColumnProps(23)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(24)=   "Column(4)._WidthInPix=1746"
      Splits(0)._ColumnProps(25)=   "Column(4)._ColStyle=1"
      Splits(0)._ColumnProps(26)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(27)=   "Column(5).Width=2725"
      Splits(0)._ColumnProps(28)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(29)=   "Column(5)._WidthInPix=2646"
      Splits(0)._ColumnProps(30)=   "Column(5).Visible=0"
      Splits(0)._ColumnProps(31)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(32)=   "Column(6).Width=2725"
      Splits(0)._ColumnProps(33)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(34)=   "Column(6)._WidthInPix=2646"
      Splits(0)._ColumnProps(35)=   "Column(6).Visible=0"
      Splits(0)._ColumnProps(36)=   "Column(6).Order=7"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   0
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      AllowUpdate     =   0   'False
      Appearance      =   0
      BorderStyle     =   0
      DataMode        =   4
      DefColWidth     =   0
      HeadLines       =   1
      FootLines       =   1
      RowDividerStyle =   6
      MultipleLines   =   0
      EmptyRows       =   -1  'True
      CellTipsWidth   =   0
      MultiSelect     =   0
      DeadAreaBackColor=   16777152
      RowDividerColor =   15790320
      RowSubDividerColor=   15790320
      DirectionAfterEnter=   1
      DirectionAfterTab=   1
      MaxRows         =   250000
      ViewColumnCaptionWidth=   0
      ViewColumnWidth =   0
      _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
      _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
      _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
      _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
      _StyleDefs(3)   =   ":id=0,.borderColor=&H0&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=29,.bgcolor=&HFFFFE0&"
      _StyleDefs(7)   =   "CaptionStyle:id=4,.parent=2,.namedParent=33"
      _StyleDefs(8)   =   "HeadingStyle:id=2,.parent=1,.namedParent=30,.bgcolor=&HFFB56A&"
      _StyleDefs(9)   =   ":id=2,.fgcolor=&HFFFFFF&,.bold=-1,.fontsize=975,.italic=0,.underline=0"
      _StyleDefs(10)  =   ":id=2,.strikethrough=0,.charset=0"
      _StyleDefs(11)  =   ":id=2,.fontname=Tahoma"
      _StyleDefs(12)  =   "FooterStyle:id=3,.parent=1,.namedParent=31"
      _StyleDefs(13)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&HFFB400&,.fgcolor=&H80000009&,.bold=-1"
      _StyleDefs(14)  =   ":id=5,.fontsize=975,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(15)  =   ":id=5,.fontname=Tahoma"
      _StyleDefs(16)  =   "SelectedStyle:id=6,.parent=1,.namedParent=32,.bgcolor=&HC0C000&,.bold=0"
      _StyleDefs(17)  =   ":id=6,.fontsize=975,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(18)  =   ":id=6,.fontname=MS Sans Serif"
      _StyleDefs(19)  =   "EditorStyle:id=7,.parent=1"
      _StyleDefs(20)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=34"
      _StyleDefs(21)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=35"
      _StyleDefs(22)  =   "OddRowStyle:id=10,.parent=1,.namedParent=36"
      _StyleDefs(23)  =   "RecordSelectorStyle:id=57,.parent=2,.namedParent=59"
      _StyleDefs(24)  =   "FilterBarStyle:id=60,.parent=1,.namedParent=62"
      _StyleDefs(25)  =   "Splits(0).Style:id=11,.parent=1"
      _StyleDefs(26)  =   "Splits(0).CaptionStyle:id=20,.parent=4"
      _StyleDefs(27)  =   "Splits(0).HeadingStyle:id=12,.parent=2,.bgcolor=&HFF8000&,.fgcolor=&HFFFFFF&"
      _StyleDefs(28)  =   "Splits(0).FooterStyle:id=13,.parent=3"
      _StyleDefs(29)  =   "Splits(0).InactiveStyle:id=14,.parent=5"
      _StyleDefs(30)  =   "Splits(0).SelectedStyle:id=16,.parent=6"
      _StyleDefs(31)  =   "Splits(0).EditorStyle:id=15,.parent=7"
      _StyleDefs(32)  =   "Splits(0).HighlightRowStyle:id=17,.parent=8"
      _StyleDefs(33)  =   "Splits(0).EvenRowStyle:id=18,.parent=9"
      _StyleDefs(34)  =   "Splits(0).OddRowStyle:id=19,.parent=10"
      _StyleDefs(35)  =   "Splits(0).RecordSelectorStyle:id=58,.parent=57"
      _StyleDefs(36)  =   "Splits(0).FilterBarStyle:id=61,.parent=60"
      _StyleDefs(37)  =   "Splits(0).Columns(0).Style:id=24,.parent=11,.locked=0,.bold=0,.fontsize=975"
      _StyleDefs(38)  =   ":id=24,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(39)  =   ":id=24,.fontname=Tahoma"
      _StyleDefs(40)  =   "Splits(0).Columns(0).HeadingStyle:id=21,.parent=12,.alignment=0,.bold=0"
      _StyleDefs(41)  =   ":id=21,.fontsize=975,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(42)  =   ":id=21,.fontname=Tahoma"
      _StyleDefs(43)  =   "Splits(0).Columns(0).FooterStyle:id=22,.parent=13"
      _StyleDefs(44)  =   "Splits(0).Columns(0).EditorStyle:id=23,.parent=15"
      _StyleDefs(45)  =   "Splits(0).Columns(1).Style:id=48,.parent=11"
      _StyleDefs(46)  =   "Splits(0).Columns(1).HeadingStyle:id=45,.parent=12"
      _StyleDefs(47)  =   "Splits(0).Columns(1).FooterStyle:id=46,.parent=13"
      _StyleDefs(48)  =   "Splits(0).Columns(1).EditorStyle:id=47,.parent=15"
      _StyleDefs(49)  =   "Splits(0).Columns(2).Style:id=56,.parent=11,.locked=0"
      _StyleDefs(50)  =   "Splits(0).Columns(2).HeadingStyle:id=53,.parent=12"
      _StyleDefs(51)  =   "Splits(0).Columns(2).FooterStyle:id=54,.parent=13"
      _StyleDefs(52)  =   "Splits(0).Columns(2).EditorStyle:id=55,.parent=15"
      _StyleDefs(53)  =   "Splits(0).Columns(3).Style:id=28,.parent=11,.alignment=2"
      _StyleDefs(54)  =   "Splits(0).Columns(3).HeadingStyle:id=25,.parent=12"
      _StyleDefs(55)  =   "Splits(0).Columns(3).FooterStyle:id=26,.parent=13"
      _StyleDefs(56)  =   "Splits(0).Columns(3).EditorStyle:id=27,.parent=15"
      _StyleDefs(57)  =   "Splits(0).Columns(4).Style:id=40,.parent=11,.alignment=2"
      _StyleDefs(58)  =   "Splits(0).Columns(4).HeadingStyle:id=37,.parent=12"
      _StyleDefs(59)  =   "Splits(0).Columns(4).FooterStyle:id=38,.parent=13"
      _StyleDefs(60)  =   "Splits(0).Columns(4).EditorStyle:id=39,.parent=15"
      _StyleDefs(61)  =   "Splits(0).Columns(5).Style:id=44,.parent=11"
      _StyleDefs(62)  =   "Splits(0).Columns(5).HeadingStyle:id=41,.parent=12"
      _StyleDefs(63)  =   "Splits(0).Columns(5).FooterStyle:id=42,.parent=13"
      _StyleDefs(64)  =   "Splits(0).Columns(5).EditorStyle:id=43,.parent=15"
      _StyleDefs(65)  =   "Splits(0).Columns(6).Style:id=52,.parent=11"
      _StyleDefs(66)  =   "Splits(0).Columns(6).HeadingStyle:id=49,.parent=12"
      _StyleDefs(67)  =   "Splits(0).Columns(6).FooterStyle:id=50,.parent=13"
      _StyleDefs(68)  =   "Splits(0).Columns(6).EditorStyle:id=51,.parent=15"
      _StyleDefs(69)  =   "Named:id=29:Normal"
      _StyleDefs(70)  =   ":id=29,.parent=0"
      _StyleDefs(71)  =   "Named:id=30:Heading"
      _StyleDefs(72)  =   ":id=30,.parent=29,.valignment=2,.bgcolor=&HFFB56A&,.fgcolor=&HFFFFFF&"
      _StyleDefs(73)  =   ":id=30,.wraptext=-1,.bold=-1,.fontsize=975,.italic=0,.underline=0"
      _StyleDefs(74)  =   ":id=30,.strikethrough=0,.charset=0"
      _StyleDefs(75)  =   ":id=30,.fontname=MS Sans Serif"
      _StyleDefs(76)  =   "Named:id=31:Footing"
      _StyleDefs(77)  =   ":id=31,.parent=29,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(78)  =   "Named:id=32:Selected"
      _StyleDefs(79)  =   ":id=32,.parent=29,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(80)  =   "Named:id=33:Caption"
      _StyleDefs(81)  =   ":id=33,.parent=30,.alignment=2"
      _StyleDefs(82)  =   "Named:id=34:HighlightRow"
      _StyleDefs(83)  =   ":id=34,.parent=29,.bgcolor=&H80000008&,.fgcolor=&H80000005&"
      _StyleDefs(84)  =   "Named:id=35:EvenRow"
      _StyleDefs(85)  =   ":id=35,.parent=29,.bgcolor=&HFFFF00&"
      _StyleDefs(86)  =   "Named:id=36:OddRow"
      _StyleDefs(87)  =   ":id=36,.parent=29"
      _StyleDefs(88)  =   "Named:id=59:RecordSelector"
      _StyleDefs(89)  =   ":id=59,.parent=30"
      _StyleDefs(90)  =   "Named:id=62:FilterBar"
      _StyleDefs(91)  =   ":id=62,.parent=29"
   End
End
Attribute VB_Name = "frmCRonda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mPercurso As XArrayDB
Private mRonda As XArrayDB
Private fEditMode As Boolean

Private Sub chkDom_Click()
   mPercurso(tdbgPercurso.Row, 6) = (chkDom.Value = vbChecked)
End Sub

Private Sub chkSab_Click()
   mPercurso(tdbgPercurso.Row, 5) = (chkSab.Value = vbChecked)
End Sub

Private Sub chkSegSex_Click()
   mPercurso(tdbgPercurso.Row, 4) = (chkSegSex.Value = vbChecked)
End Sub

Private Sub cmdDelete_Click()
   If lstHorario.ListIndex <> -1 Then
      lstHorario.RemoveItem lstHorario.ListIndex
      On Error Resume Next
      lstHorario.ListIndex = 0
   End If
End Sub

Private Sub cmdOk_Click()
   If IsDate(lstHorario.Text) Then
      lstHorario.AddItem lstHorario.Text
      lstHorario.ListIndex = 0
   End If
End Sub

Private Sub lblEdit_Click()
   If lblEdit.Caption = "Editar" Then
      EnableControls True
      lblEdit.Caption = "Cancelar"
      lblSave.Enabled = True
      fEditMode = True
   Else  'Cancel Mode
      EnableControls False
      lblEdit.Caption = "Editar"
      lblSave.Enabled = False
      fEditMode = False
      tdbgPercurso.Close
      Form_Activate
   End If
End Sub

Private Sub EnableControls(ByVal fEnabled As Boolean)
   tdbgPercurso.AllowUpdate = fEnabled
   tdbgRonda.AllowUpdate = fEnabled
   tdbgRonda.AllowAddNew = fEnabled
   optStatus(0).Enabled = fEnabled
   optStatus(1).Enabled = fEnabled
   chkSegSex.Enabled = fEnabled
   chkSab.Enabled = fEnabled
   chkDom.Enabled = fEnabled
   txtDesvio.Enabled = fEnabled
   lstHorario.Locked = Not fEnabled
   cmdOk.Enabled = fEnabled
   cmdDelete.Enabled = fEnabled
End Sub

Private Sub lblExit_Click()
   Unload Me
End Sub

Private Sub Form_Load()
'   Dim success As Long
'   success = SetWindowPos(frmCRonda.hwnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
'   curhwnd = frmCRonda.hwnd
   Left = (Screen.Width - Width) / 2   ' Center form horizontally.
   Top = (Screen.Height - Height) / 2  ' Center form vertically.
End Sub

Private Sub Form_Activate()
   Dim cP As clsPercurso
   Dim mRow As Integer
   Dim mCol As Integer
   fEditMode = False
   Set mPercurso = Nothing
   lblRemove.Enabled = (lstPercurso.Count >= 1)
   If lstPercurso.Count >= 1 Then
      ' Allocate space for rows, 8 columns
      Set mPercurso = New XArrayDB
      mPercurso.ReDim 0, lstPercurso.Count - 1, 0, 7
      mRow = 0
      For Each cP In lstPercurso
         With cP
            mPercurso(mRow, 0) = CStr(.descrPercurso)
            mPercurso(mRow, 1) = .idPercurso
            mPercurso(mRow, 2) = .Horario
            mPercurso(mRow, 3) = .desvio
            mPercurso(mRow, 4) = .valSegSex
            mPercurso(mRow, 5) = .valSab
            mPercurso(mRow, 6) = .valDom
            mPercurso(mRow, 7) = .status
            mRow = mRow + 1
         End With
      Next
      mPercurso.ReDim 0, mRow - 1, 0, 7
      tdbgPercurso.Array = mPercurso
      tdbgPercurso.ReBind
   End If
   
   With AdodcRonda
      .ConnectionString = cnDB
      .CommandType = adCmdText
      .CursorType = adOpenStatic
      .LockType = adLockReadOnly
      .RecordSource = "SELECT * FROM Entity_Ronda"
      .Refresh
   End With

End Sub

Private Sub lblInsert_Click()
   Dim tPercurso As New clsPercurso
   With tPercurso
      .descrPercurso = "Novo Percurso"
      .Horario = Time
      .Insert
   End With
   If lstPercurso.Count = 0 Then
      Set mPercurso = New XArrayDB
   End If
   lstPercurso.Add Item:=tPercurso, Key:=CStr(tPercurso.idPercurso)
   mPercurso.ReDim 0, lstPercurso.Count - 1, 0, 7
   Dim mRow As Integer
   mRow = lstPercurso.Count - 1
   With tPercurso
      mPercurso(mRow, 0) = CStr(.descrPercurso)
      mPercurso(mRow, 1) = .idPercurso
      mPercurso(mRow, 2) = .Horario
      mPercurso(mRow, 3) = 0
      mPercurso(mRow, 4) = False
      mPercurso(mRow, 5) = False
      mPercurso(mRow, 6) = False
      mPercurso(mRow, 7) = False
   End With
   tdbgPercurso.Array = mPercurso
   tdbgPercurso.ReBind
   tdbgPercurso.MoveLast
   lblEdit_Click
End Sub

Private Sub lblPrint_Click()
   SetHourGlass Me
   Dim frm As New frmViewReport9
   frm.SetTipo = g_iRptCRonda
   frm.WindowState = vbMaximized
   frm.Show
   ResetMouse Me

'   Dim success As Long
'   success = SetWindowPos(frmCRonda.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, FLAGS)
'   On Error GoTo ErrorHandler
'   rpt1.ReportFileName = m_sPath & "\Config Ronda.rpt"
'   rpt1.WindowTitle = "Relatório de Configuração de Ronda"
'   rpt1.Destination = crptToWindow
'   rpt1.WindowState = crptMaximized
'   rpt1.WindowParentHandle = ForNet.hwnd
'   rpt1.WindowBorderStyle = crptFixedSingle
'   rpt1.WindowControlBox = False
'   rpt1.WindowShowCloseBtn = True
'   rpt1.action = 1
'   Exit Sub
'ErrorHandler:
'   MsgBox rpt1.LastErrorString, sxExclamation, sxProname
'   Resume Next
End Sub

Private Sub lblRemove_Click()
   Dim tPercurso As clsPercurso
   Set tPercurso = lstPercurso.Item(CStr(mPercurso(tdbgPercurso.Row, 1)))
   tPercurso.Remove_Horarios
   tPercurso.RemoveRondas
   tPercurso.Remove
   lstPercurso.Remove CStr(mPercurso(tdbgPercurso.Row, 1))
   lstHorario.Clear
   tdbgRonda.Close
   tdbgPercurso.Close
   Form_Activate
End Sub

Private Sub lblSave_Click()
   EnableControls False
   lblSave.Enabled = False
   lblEdit.Caption = "Editar"
   fEditMode = False
   Dim mRow As Integer
   mRow = tdbgPercurso.Row
   Dim tPercurso As clsPercurso
   Set tPercurso = lstPercurso.Item(CStr(mPercurso(mRow, 1)))
   With tPercurso
      tdbgPercurso.Update
      .descrPercurso = mPercurso(mRow, 0)
      .Horario = mPercurso(mRow, 2)
      .desvio = mPercurso(mRow, 3)
      .valSegSex = mPercurso(mRow, 4)
      .valSab = mPercurso(mRow, 5)
      .valDom = mPercurso(mRow, 6)
      .status = mPercurso(mRow, 7)
   End With
   tPercurso.Update
   Save_Rondas (mRow)
   Save_Horarios (mRow)
End Sub

Private Sub Save_Rondas(fRow As Integer)
   tdbgRonda.Update
   Dim tPercurso As clsPercurso
   Set tPercurso = lstPercurso.Item(CStr(mPercurso(fRow, 1)))
   Dim tronda As clsRonda
   Dim i As Integer
   For i = 0 To mRonda.UpperBound(1)
      Set tronda = Nothing
      On Error Resume Next
      If tPercurso.lstRonda.Count > 0 Then
         Set tronda = tPercurso.lstRonda.Item(CStr(mRonda(i, 5)))
      End If
      On Error GoTo 0
      If tronda Is Nothing Then
         'new, create one
         Set tronda = New clsRonda
         With tronda
            .descrRonda = mRonda(i, 0)
            .idEntity = mRonda(i, 1)
            .intervalo = mRonda(i, 3)
            .status = mRonda(i, 4)
            .idPercurso = tPercurso.idPercurso
         End With
         tronda.Insert
         mRonda(i, 5) = tronda.idRonda
         tPercurso.lstRonda.Add tronda, CStr(tronda.idRonda)
      Else
         'update it
         With tronda
            .descrRonda = mRonda(i, 0)
            .idEntity = mRonda(i, 1)
            .intervalo = mRonda(i, 3)
            .status = mRonda(i, 4)
         End With
         tronda.Update
      End If
   Next i

End Sub

Private Sub Save_Horarios(fRow As Integer)
   Dim tPercurso As clsPercurso
   Set tPercurso = lstPercurso.Item(CStr(mPercurso(fRow, 1)))
   tPercurso.Remove_Horarios
   Dim i As Integer
   For i = 0 To lstHorario.ListCount - 1
      tPercurso.Insert_Horario lstHorario.List(i)
   Next i
   tPercurso.Save_Horarios
End Sub

Private Sub lstHorario_KeyPress(KeyAscii As Integer)
   If KeyAscii = vbKeyReturn Then
      cmdOk_Click
   End If
End Sub

Private Sub optStatus_Click(Index As Integer)
   mPercurso(tdbgPercurso.Row, 7) = optStatus(0).Value
End Sub

Private Sub tdbgPercurso_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
   If tdbgPercurso.Bookmark <> -1 Then
      UpdateFields (tdbgPercurso.Bookmark)
   End If
End Sub

Private Sub UpdateFields(fRow As Integer)
   optStatus(0).Value = mPercurso(fRow, 7)
   If mPercurso(fRow, 4) Then
      chkSegSex.Value = vbChecked
   Else
      chkSegSex.Value = vbUnchecked
   End If
   If mPercurso(fRow, 5) Then
      chkSab.Value = vbChecked
   Else
      chkSab.Value = vbUnchecked
   End If
   If mPercurso(fRow, 6) Then
      chkDom.Value = vbChecked
   Else
      chkDom.Value = vbUnchecked
   End If
   txtDesvio.Text = mPercurso(fRow, 3)
   UpdateRonda (fRow)
   UpdateHorario (fRow)
End Sub

Private Sub UpdateRonda(fRow As Integer)
   Dim cP As clsPercurso
   Dim cR As clsRonda
   Dim mRow As Integer
   Dim mCol As Integer
   Set mRonda = Nothing
   Set cP = lstPercurso.Item(CStr(mPercurso(fRow, 1)))
'   If cP.lstRonda.Count > 0 Then
      ' Allocate space for rows, 7 columns
      Set mRonda = New XArrayDB
      mRonda.ReDim 0, cP.lstRonda.Count - 1, 0, 6
      mRow = 0
      For Each cR In cP.lstRonda
         With cR
            mRonda(mRow, 0) = CStr(.descrRonda)
            mRonda(mRow, 1) = .idEntity
            mRonda(mRow, 2) = .descrEntity
            mRonda(mRow, 3) = .intervalo
            mRonda(mRow, 4) = .status
            mRonda(mRow, 5) = .idRonda
            mRonda(mRow, 6) = .idPercurso
            mRow = mRow + 1
         End With
      Next
      mRonda.ReDim 0, mRow - 1, 0, 6
      tdbgRonda.Array = mRonda
      tdbgRonda.ReBind
'   Else
'      Set mRonda = New XArrayDB
'      mRonda.ReDim 0, -1, 0, 6
'      tdbgRonda.Array = mRonda
'   End If
   tdbgRonda.Refresh
End Sub

Private Sub UpdateHorario(fRow As Integer)
   Dim cP As clsPercurso
   Set cP = lstPercurso.Item(CStr(mPercurso(fRow, 1)))
   lstHorario.Clear
   Dim ch As clsHorario
   For Each ch In cP.lstHorario
      lstHorario.AddItem ch.Horario
   Next
   On Error Resume Next
   lstHorario.ListIndex = 0
End Sub

Private Sub tdbgRonda_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
   If ColIndex = 2 And fEditMode Then
      tdbgRonda.Columns(1).Value = TDBDropDownRonda.Columns(0).Value
   End If
End Sub

Private Sub txtDesvio_Change()
   If IsNumeric(txtDesvio.Text) Then
      mPercurso(tdbgPercurso.Row, 3) = txtDesvio.Text
   End If
End Sub

