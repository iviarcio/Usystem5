VERSION 5.00
Begin VB.Form frmConfig 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Configuração"
   ClientHeight    =   3360
   ClientLeft      =   1755
   ClientTop       =   1740
   ClientWidth     =   5370
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   Icon            =   "Config.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3360
   ScaleWidth      =   5370
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame5 
      Caption         =   "Hardware: "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   975
      Left            =   2835
      TabIndex        =   20
      Top             =   2310
      Width           =   2460
      Begin VB.CheckBox chkRTS 
         Caption         =   "RTSEnable"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   315
         TabIndex        =   22
         Top             =   615
         Value           =   1  'Checked
         Width           =   1950
      End
      Begin VB.CheckBox chkDTR 
         Caption         =   "DTREnable"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   315
         TabIndex        =   21
         Top             =   270
         Value           =   1  'Checked
         Width           =   1950
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Handshaking: "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   960
      Left            =   105
      TabIndex        =   18
      Top             =   2310
      Width           =   2625
      Begin VB.ComboBox lstHandshaking 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "Config.frx":030A
         Left            =   150
         List            =   "Config.frx":031A
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   270
         Width           =   2310
      End
   End
   Begin VB.Frame Frame3 
      Height          =   2295
      Left            =   2280
      TabIndex        =   9
      Top             =   0
      Width           =   3015
      Begin VB.ComboBox lstBaud 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1155
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   240
         Width           =   1605
      End
      Begin VB.ComboBox lstParity 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1155
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   765
         Width           =   1605
      End
      Begin VB.ComboBox lstData 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1155
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   1290
         Width           =   1605
      End
      Begin VB.ComboBox lstStop 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1155
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   1785
         Width           =   1605
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         Caption         =   "&Baud Rate:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   270
         Left            =   120
         TabIndex        =   17
         Top             =   270
         Width           =   1050
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         Caption         =   "Par&ity:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   270
         Left            =   540
         TabIndex        =   16
         Top             =   795
         Width           =   630
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         Caption         =   "&Data Bits:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   270
         Left            =   240
         TabIndex        =   15
         Top             =   1305
         Width           =   930
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         Caption         =   "Stop Bi&ts:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   270
         Left            =   270
         TabIndex        =   14
         Top             =   1800
         Width           =   915
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Tempos: "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   855
      Left            =   120
      TabIndex        =   5
      Top             =   1425
      Width           =   2055
      Begin VB.TextBox txtSeg 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   1
         Left            =   1515
         TabIndex        =   6
         TabStop         =   0   'False
         Text            =   "(ms)"
         Top             =   405
         Width           =   360
      End
      Begin VB.TextBox txtEspera 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   900
         TabIndex        =   7
         Top             =   375
         Width           =   1020
      End
      Begin VB.Label lblBounce 
         Appearance      =   0  'Flat
         Caption         =   "&Espera:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   90
         TabIndex        =   8
         Top             =   405
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Porta: "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   2055
      Begin VB.OptionButton optComm 
         Caption         =   "Com1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   765
      End
      Begin VB.OptionButton optComm 
         Caption         =   "Com2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   990
         TabIndex        =   3
         Top             =   255
         Width           =   765
      End
      Begin VB.OptionButton optComm 
         Caption         =   "Com3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   2
         Left            =   120
         TabIndex        =   2
         Top             =   630
         Width           =   765
      End
      Begin VB.OptionButton optComm 
         Caption         =   "Com4"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   3
         Left            =   990
         TabIndex        =   1
         Top             =   630
         Width           =   765
      End
   End
End
Attribute VB_Name = "frmConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Controla as alterações
Private fport As Integer
Private fbaud As String
Private fdata As String
Private fparity As String
Private fstop As String
Private fHandshaking As Integer
Private fDTR As Boolean
Private fRTS As Boolean
Private gChange As Boolean

Private Sub Form_Load()
   'Init Values
   lstBaud.AddItem "2400"
   lstBaud.AddItem "4800"
   lstBaud.AddItem "9600"
   lstBaud.AddItem "19200"
   lstBaud.AddItem "38400"
   lstBaud.AddItem "57600"
   lstBaud.AddItem "115200"
   lstParity.AddItem "E"  'Even
   lstParity.AddItem "O"  'Odd
   lstParity.AddItem "N"  'None
   lstParity.AddItem "M"  'Mark
   lstParity.AddItem "S"  'Space
   lstData.AddItem "4"
   lstData.AddItem "5"
   lstData.AddItem "6"
   lstData.AddItem "7"
   lstData.AddItem "8"
   lstStop.AddItem "1"
   lstStop.AddItem "1.5"
   lstStop.AddItem "2"
   LoadConfig
   gChange = False
End Sub

Private Sub LoadConfig()
   Dim v1, v2, v3 As Integer
   On Error Resume Next
   'Tempo de Repetição e Espera
   txtEspera = uControl.gstEspera
   'read the default port
   fport = CInt(uControl.gstCommPort)
   optComm(fport - 1).Value = True
   'read the default baud rate
   v1 = InStr(1, uControl.gstCommSett, ",")
   fbaud = Left$(uControl.gstCommSett, v1 - 1)
   SetControl lstBaud, fbaud
   'read the default parity
   fparity = Mid$(uControl.gstCommSett, v1 + 1, 1)
   SetControl lstParity, fparity
   'read the default data bits
   v2 = InStr(v1 + 1, uControl.gstCommSett, ",")
   fdata = Mid$(uControl.gstCommSett, v2 + 1, 1)
   SetControl lstData, fdata
   'read the stop bits
   v3 = InStr(v2 + 1, uControl.gstCommSett, ",")
   fstop = Mid$(uControl.gstCommSett, v3 + 1, 1)
   SetControl lstStop, fstop
   
   fHandshaking = uControl.iHandshaking
   lstHandshaking.ListIndex = fHandshaking
   
   fDTR = uControl.bDTREnable
   If fDTR Then
      chkDTR.Value = vbChecked
   Else
      chkDTR.Value = vbUnchecked
   End If
   
   fRTS = uControl.bRTSEnable
   If fRTS Then
      chkRTS.Value = vbChecked
   Else
      chkRTS.Value = vbUnchecked
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If gChange Then
      uControl.gstCommPort = CStr(fport)
      uControl.gstCommSett = fbaud & "," & fparity & "," & fdata & "," & fstop
      uControl.gstEspera = txtEspera
      uControl.iHandshaking = fHandshaking
      uControl.bDTREnable = fDTR
      uControl.bRTSEnable = fRTS
      Save_Config
   End If
End Sub

Private Sub lstHandshaking_Click()
   fHandshaking = lstHandshaking.ListIndex
   gChange = True
End Sub

Private Sub lstBaud_Click()
   fbaud = lstBaud.Text
   gChange = True
End Sub

Private Sub lstData_Click()
   fdata = lstData.Text
   gChange = True
End Sub

Private Sub lstParity_Click()
   fparity = lstParity.Text
   gChange = True
End Sub

Private Sub lstStop_Click()
   fstop = lstStop.Text
   gChange = True
End Sub

Private Sub SetControl(fc As Control, ftext As String)
   Dim i%
   For i% = 0 To fc.ListCount
      If fc.List(i%) = ftext Then fc.ListIndex = i%
   Next i%
End Sub

Private Sub optComm_Click(Index As Integer)
   gChange = True
   fport = Index + 1
End Sub

Private Sub txtEspera_Change()
   gChange = True
End Sub

Private Sub txtEspera_KeyPress(KeyAscii As Integer)
   If KeyAscii = vbKeyReturn Then
      KeyAscii = 0
      lstBaud.SetFocus
   ElseIf ((KeyAscii < vbKey0) Or (KeyAscii > vbKey9)) And (KeyAscii <> vbKeyBack) Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Sub txtRepeticao_Change()
    gChange = True
End Sub

Private Sub txtRepeticao_KeyPress(KeyAscii As Integer)
   If KeyAscii = vbKeyReturn Then
      KeyAscii = 0
      txtEspera.SetFocus
   ElseIf ((KeyAscii < vbKey0) Or (KeyAscii > vbKey9)) And (KeyAscii <> vbKeyBack) Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Sub chkDTR_Click()
   If chkDTR.Value = vbChecked Then
      fDTR = True
   Else
      fDTR = False
   End If
   gChange = True
End Sub

Private Sub chkRTS_Click()
   If chkRTS.Value = vbChecked Then
      fRTS = True
   Else
      fRTS = False
   End If
   gChange = True
End Sub

Private Sub Save_Config()
   SaveSetting "uControl", "Options", "CommPort", uControl.gstCommPort
   SaveSetting "uControl", "Options", "CommSett", uControl.gstCommSett
   SaveSetting "uControl", "Options", "Espera", uControl.gstEspera
   SaveSetting "uControl", "Options", "Handshaking", CStr(uControl.iHandshaking)
   SaveSetting "uControl", "Options", "DTREnable", CStr(uControl.bDTREnable)
   SaveSetting "uControl", "Options", "RTSEnable", CStr(uControl.bRTSEnable)
End Sub

