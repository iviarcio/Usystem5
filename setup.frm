VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{A4749554-0441-4E64-8A03-3323601631C7}#1.0#0"; "LaVolpeAlphaImg2.ocx"
Begin VB.Form frmSetup 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Configuração das portas de Comunicação"
   ClientHeight    =   5640
   ClientLeft      =   1410
   ClientTop       =   1650
   ClientWidth     =   8010
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "setup.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5640
   ScaleWidth      =   8010
   Begin TabDlg.SSTab SSTab1 
      Height          =   4575
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   8070
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Comunicação A"
      TabPicture(0)   =   "setup.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label7(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "SSFrame3(0)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "chkPorta(0)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lstPort(0)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "Comunicação B"
      TabPicture(1)   =   "setup.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label7(1)"
      Tab(1).Control(1)=   "SSFrame3(1)"
      Tab(1).Control(2)=   "chkPorta(1)"
      Tab(1).Control(3)=   "lstPort(1)"
      Tab(1).ControlCount=   4
      TabCaption(2)   =   "Comunicação C"
      TabPicture(2)   =   "setup.frx":0342
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label7(2)"
      Tab(2).Control(1)=   "SSFrame3(2)"
      Tab(2).Control(2)=   "chkPorta(2)"
      Tab(2).Control(3)=   "lstPort(2)"
      Tab(2).ControlCount=   4
      TabCaption(3)   =   "Comunicação D"
      TabPicture(3)   =   "setup.frx":035E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Label7(3)"
      Tab(3).Control(1)=   "SSFrame3(3)"
      Tab(3).Control(2)=   "chkPorta(3)"
      Tab(3).Control(3)=   "lstPort(3)"
      Tab(3).ControlCount=   4
      Begin VB.ComboBox lstPort 
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
         Index           =   3
         ItemData        =   "setup.frx":037A
         Left            =   -74640
         List            =   "setup.frx":03C5
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   1560
         Width           =   2895
      End
      Begin VB.ComboBox lstPort 
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
         Index           =   2
         ItemData        =   "setup.frx":0460
         Left            =   -74640
         List            =   "setup.frx":04AB
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   1560
         Width           =   2895
      End
      Begin VB.ComboBox lstPort 
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
         Index           =   1
         ItemData        =   "setup.frx":0546
         Left            =   -74640
         List            =   "setup.frx":0591
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   1560
         Width           =   2895
      End
      Begin VB.ComboBox lstPort 
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
         Index           =   0
         ItemData        =   "setup.frx":062C
         Left            =   360
         List            =   "setup.frx":0677
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   1560
         Width           =   2895
      End
      Begin VB.CheckBox chkPorta 
         Caption         =   "Porta Habilitada"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   -74640
         TabIndex        =   4
         Top             =   600
         Width           =   2535
      End
      Begin VB.CheckBox chkPorta 
         Caption         =   "Porta Habilitada"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   -74640
         TabIndex        =   3
         Top             =   600
         Width           =   2535
      End
      Begin VB.CheckBox chkPorta 
         Caption         =   "Porta Habilitada"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   -74640
         TabIndex        =   2
         Top             =   600
         Width           =   2535
      End
      Begin VB.CheckBox chkPorta 
         Caption         =   "Porta Habilitada"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   360
         TabIndex        =   1
         Top             =   600
         Width           =   2535
      End
      Begin Threed.SSFrame SSFrame3 
         Height          =   2775
         Index           =   1
         Left            =   -71160
         TabIndex        =   13
         Top             =   1320
         Width           =   3255
         _Version        =   65536
         _ExtentX        =   5741
         _ExtentY        =   4895
         _StockProps     =   14
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin VB.ComboBox lstBaud 
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
            Height          =   390
            Index           =   1
            ItemData        =   "setup.frx":0712
            Left            =   1440
            List            =   "setup.frx":0728
            Style           =   2  'Dropdown List
            TabIndex        =   17
            Top             =   240
            Width           =   1575
         End
         Begin VB.ComboBox lstParity 
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
            Height          =   390
            Index           =   1
            ItemData        =   "setup.frx":0752
            Left            =   1440
            List            =   "setup.frx":0765
            Style           =   2  'Dropdown List
            TabIndex        =   16
            Top             =   840
            Width           =   1575
         End
         Begin VB.ComboBox lstData 
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
            Height          =   390
            Index           =   1
            ItemData        =   "setup.frx":0787
            Left            =   1440
            List            =   "setup.frx":079A
            Style           =   2  'Dropdown List
            TabIndex        =   15
            Top             =   1440
            Width           =   1575
         End
         Begin VB.ComboBox lstStop 
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
            Height          =   390
            Index           =   1
            ItemData        =   "setup.frx":07AD
            Left            =   1440
            List            =   "setup.frx":07BA
            Style           =   2  'Dropdown List
            TabIndex        =   14
            Top             =   2040
            Width           =   1575
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Caption         =   "&Velocidade:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   21
            Top             =   270
            Width           =   1215
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Caption         =   "&Paridade:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   20
            Top             =   870
            Width           =   1215
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Caption         =   "&Dados:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   255
            Index           =   1
            Left            =   480
            TabIndex        =   19
            Top             =   1470
            Width           =   855
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Caption         =   "Stop Bi&t:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   255
            Index           =   1
            Left            =   600
            TabIndex        =   18
            Top             =   2070
            Width           =   735
         End
      End
      Begin Threed.SSFrame SSFrame3 
         Height          =   2775
         Index           =   0
         Left            =   3840
         TabIndex        =   22
         Top             =   1320
         Width           =   3255
         _Version        =   65536
         _ExtentX        =   5741
         _ExtentY        =   4895
         _StockProps     =   14
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin VB.ComboBox lstStop 
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
            Height          =   390
            Index           =   0
            ItemData        =   "setup.frx":07C9
            Left            =   1440
            List            =   "setup.frx":07D6
            Style           =   2  'Dropdown List
            TabIndex        =   26
            Top             =   2040
            Width           =   1575
         End
         Begin VB.ComboBox lstData 
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
            Height          =   390
            Index           =   0
            ItemData        =   "setup.frx":07E5
            Left            =   1440
            List            =   "setup.frx":07F8
            Style           =   2  'Dropdown List
            TabIndex        =   25
            Top             =   1440
            Width           =   1575
         End
         Begin VB.ComboBox lstParity 
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
            Height          =   390
            Index           =   0
            ItemData        =   "setup.frx":080B
            Left            =   1440
            List            =   "setup.frx":081E
            Style           =   2  'Dropdown List
            TabIndex        =   24
            Top             =   840
            Width           =   1575
         End
         Begin VB.ComboBox lstBaud 
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
            Height          =   390
            Index           =   0
            ItemData        =   "setup.frx":0840
            Left            =   1440
            List            =   "setup.frx":0856
            Style           =   2  'Dropdown List
            TabIndex        =   23
            Top             =   240
            Width           =   1575
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Caption         =   "Stop Bi&t:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   255
            Index           =   0
            Left            =   600
            TabIndex        =   30
            Top             =   2070
            Width           =   735
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Caption         =   "&Dados:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   255
            Index           =   0
            Left            =   480
            TabIndex        =   29
            Top             =   1470
            Width           =   855
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Caption         =   "&Paridade:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   28
            Top             =   870
            Width           =   1215
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Caption         =   "&Velocidade:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   27
            Top             =   270
            Width           =   1215
         End
      End
      Begin Threed.SSFrame SSFrame3 
         Height          =   2775
         Index           =   2
         Left            =   -71160
         TabIndex        =   31
         Top             =   1320
         Width           =   3255
         _Version        =   65536
         _ExtentX        =   5741
         _ExtentY        =   4895
         _StockProps     =   14
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin VB.ComboBox lstStop 
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
            Height          =   390
            Index           =   2
            ItemData        =   "setup.frx":0880
            Left            =   1440
            List            =   "setup.frx":088D
            Style           =   2  'Dropdown List
            TabIndex        =   35
            Top             =   2040
            Width           =   1575
         End
         Begin VB.ComboBox lstData 
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
            Height          =   390
            Index           =   2
            ItemData        =   "setup.frx":089C
            Left            =   1440
            List            =   "setup.frx":08AF
            Style           =   2  'Dropdown List
            TabIndex        =   34
            Top             =   1440
            Width           =   1575
         End
         Begin VB.ComboBox lstParity 
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
            Height          =   390
            Index           =   2
            ItemData        =   "setup.frx":08C2
            Left            =   1440
            List            =   "setup.frx":08D5
            Style           =   2  'Dropdown List
            TabIndex        =   33
            Top             =   840
            Width           =   1575
         End
         Begin VB.ComboBox lstBaud 
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
            Height          =   390
            Index           =   2
            ItemData        =   "setup.frx":08F7
            Left            =   1440
            List            =   "setup.frx":090D
            Style           =   2  'Dropdown List
            TabIndex        =   32
            Top             =   240
            Width           =   1575
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Caption         =   "Stop Bi&t:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   255
            Index           =   2
            Left            =   600
            TabIndex        =   39
            Top             =   2070
            Width           =   735
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Caption         =   "&Dados:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   255
            Index           =   2
            Left            =   480
            TabIndex        =   38
            Top             =   1470
            Width           =   855
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Caption         =   "&Paridade:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   37
            Top             =   870
            Width           =   1215
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Caption         =   "&Velocidade:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   36
            Top             =   270
            Width           =   1215
         End
      End
      Begin Threed.SSFrame SSFrame3 
         Height          =   2775
         Index           =   3
         Left            =   -71160
         TabIndex        =   40
         Top             =   1320
         Width           =   3255
         _Version        =   65536
         _ExtentX        =   5741
         _ExtentY        =   4895
         _StockProps     =   14
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin VB.ComboBox lstStop 
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
            Height          =   390
            Index           =   3
            ItemData        =   "setup.frx":0937
            Left            =   1440
            List            =   "setup.frx":0944
            Style           =   2  'Dropdown List
            TabIndex        =   44
            Top             =   2040
            Width           =   1575
         End
         Begin VB.ComboBox lstData 
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
            Height          =   390
            Index           =   3
            ItemData        =   "setup.frx":0953
            Left            =   1440
            List            =   "setup.frx":0966
            Style           =   2  'Dropdown List
            TabIndex        =   43
            Top             =   1440
            Width           =   1575
         End
         Begin VB.ComboBox lstParity 
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
            Height          =   390
            Index           =   3
            ItemData        =   "setup.frx":0979
            Left            =   1440
            List            =   "setup.frx":098C
            Style           =   2  'Dropdown List
            TabIndex        =   42
            Top             =   840
            Width           =   1575
         End
         Begin VB.ComboBox lstBaud 
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
            Height          =   390
            Index           =   3
            ItemData        =   "setup.frx":09AE
            Left            =   1440
            List            =   "setup.frx":09C4
            Style           =   2  'Dropdown List
            TabIndex        =   41
            Top             =   240
            Width           =   1575
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Caption         =   "Stop Bi&t:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   255
            Index           =   3
            Left            =   600
            TabIndex        =   48
            Top             =   2070
            Width           =   735
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Caption         =   "&Dados:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   255
            Index           =   3
            Left            =   480
            TabIndex        =   47
            Top             =   1470
            Width           =   855
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Caption         =   "&Paridade:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   46
            Top             =   870
            Width           =   1215
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Caption         =   "&Velocidade:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   45
            Top             =   270
            Width           =   1215
         End
      End
      Begin VB.Label Label7 
         Caption         =   "Porta de Comunicação"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Index           =   3
         Left            =   -74640
         TabIndex        =   12
         Top             =   1200
         Width           =   2295
      End
      Begin VB.Label Label7 
         Caption         =   "Porta de Comunicação"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Index           =   2
         Left            =   -74640
         TabIndex        =   10
         Top             =   1200
         Width           =   2295
      End
      Begin VB.Label Label7 
         Caption         =   "Porta de Comunicação"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Index           =   1
         Left            =   -74640
         TabIndex        =   8
         Top             =   1200
         Width           =   2295
      End
      Begin VB.Label Label7 
         Caption         =   "Porta de Comunicação"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Index           =   0
         Left            =   360
         TabIndex        =   6
         Top             =   1200
         Width           =   2295
      End
   End
   Begin LaVolpeAlphaImg.AlphaImgCtl cmdExit 
      Height          =   720
      Left            =   7200
      ToolTipText     =   "Fechar Configuração das Portas Comm"
      Top             =   4800
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   1270
      Image           =   "setup.frx":09EE
      Effects         =   "setup.frx":16F3
   End
   Begin VB.Label Label1 
      Caption         =   "Obs.: As modificações nas portas de Comunicação terão efeito somente após a reinicialização do Sistema."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   615
      Left            =   120
      TabIndex        =   49
      Top             =   4800
      Width           =   6735
   End
End
Attribute VB_Name = "frmSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private fInit As Boolean

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

Private Sub Form_Load()
   Dim success As Long
   success = SetWindowPos(frmSetup.hWnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
   curhwnd = frmSetup.hWnd
   Left = (Screen.Width - Width) / 2   ' Center form horizontally.
   Top = (Screen.Height - Height) / 2  ' Center form vertically.
   
   fInit = True
   
'  Get the current configuration
   Dim i As Integer
   Dim lds As New ADODB.Recordset
   lds.Open "SELECT * FROM Config", cnDB, adOpenStatic, adLockReadOnly
   For i = 0 To 3
      If lds("Enabled") Then chkPorta(i).Value = vbChecked
      lstPort(i).ListIndex = lds("Comm") - 1
      lstBaud(i).ListIndex = lds("BaudRate")
      lstParity(i).ListIndex = lds("Parity")
      lstData(i).ListIndex = lds("DataBits")
      lstStop(i).ListIndex = lds("StopBits")
      lds.MoveNext
   Next
   
   fInit = False
   
   Me.MousePointer = vbDefault
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If m_bChange Then
      'Save the current configuration
      Dim i As Integer
      Dim lds As New ADODB.Recordset
      lds.Open "SELECT * FROM Config", cnDB, adOpenKeyset, adLockOptimistic
      For i = 0 To 3
         lds("Enabled") = chkPorta(i).Value
         lds("Comm") = lstPort(i).ItemData(lstPort(i).ListIndex)
         lds("BaudRate") = lstBaud(i).ListIndex
         lds("Parity") = lstParity(i).ListIndex
         lds("DataBits") = lstData(i).ListIndex
         lds("StopBits") = lstStop(i).ListIndex
         lds.Update
         lds.MoveNext
      Next
      lds.Close
      Make_Service "Alteração nas portas de Comunicação", strAccess(m_tAccess) & m_sUser
   End If
End Sub

Private Sub lstBaud_Click(Index As Integer)
   If Not fInit Then m_bChange = True
End Sub

Private Sub lstData_Click(Index As Integer)
   If Not fInit Then m_bChange = True
End Sub

Private Sub lstParity_Click(Index As Integer)
   If Not fInit Then m_bChange = True
End Sub

Private Sub lstPort_Click(Index As Integer)
   If Not fInit Then m_bChange = True
End Sub

Private Sub lstStop_Click(Index As Integer)
   If Not fInit Then m_bChange = True
End Sub

Private Sub chkPorta_Click(Index As Integer)
   If Not fInit Then m_bChange = True
End Sub

