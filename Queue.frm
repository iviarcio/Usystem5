VERSION 5.00
Object = "{5C8CED40-8909-11D0-9483-00A0C91110ED}#1.0#0"; "MSDATREP.OCX"
Begin VB.Form frmQueue 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Eventos Críticos"
   ClientHeight    =   4650
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5925
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4650
   ScaleWidth      =   5925
   Visible         =   0   'False
   Begin MSDataRepeaterLib.DataRepeater DataRepeater1 
      Height          =   3840
      Left            =   0
      TabIndex        =   6
      Top             =   840
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   6773
      _StreamID       =   -1412567295
      _Version        =   393216
      CaptionStyle    =   0
      BeginProperty RepeatedControlName {21FC0FC0-1E5C-11D1-A327-00AA00688B10} 
         _StreamID       =   -1412567295
         _Version        =   65536
         Name            =   "prjCritico.ctlCritico"
      EndProperty
      RepeaterBindings=   8
      BeginProperty RepeaterBinding(0) {7D21A594-FC9B-11D0-A320-00AA00688B10} 
         _StreamID       =   -1412567295
         _Version        =   65536
         PropertyName    =   "UID"
         DataField       =   "UID"
      EndProperty
      BeginProperty RepeaterBinding(1) {7D21A594-FC9B-11D0-A320-00AA00688B10} 
         _StreamID       =   -1412567295
         _Version        =   65536
         PropertyName    =   "Color"
         DataField       =   "Color"
      EndProperty
      BeginProperty RepeaterBinding(2) {7D21A594-FC9B-11D0-A320-00AA00688B10} 
         _StreamID       =   -1412567295
         _Version        =   65536
         PropertyName    =   "Evento"
         DataField       =   "Evento"
      EndProperty
      BeginProperty RepeaterBinding(3) {7D21A594-FC9B-11D0-A320-00AA00688B10} 
         _StreamID       =   -1412567295
         _Version        =   65536
         PropertyName    =   "EvTime"
         DataField       =   "EvTime"
      EndProperty
      BeginProperty RepeaterBinding(4) {7D21A594-FC9B-11D0-A320-00AA00688B10} 
         _StreamID       =   -1412567295
         _Version        =   65536
         PropertyName    =   "Loja"
         DataField       =   "Loja"
      EndProperty
      BeginProperty RepeaterBinding(5) {7D21A594-FC9B-11D0-A320-00AA00688B10} 
         _StreamID       =   -1412567295
         _Version        =   65536
         PropertyName    =   "Message"
         DataField       =   "Message"
      EndProperty
      BeginProperty RepeaterBinding(6) {7D21A594-FC9B-11D0-A320-00AA00688B10} 
         _StreamID       =   -1412567295
         _Version        =   65536
         PropertyName    =   "Sensor"
         DataField       =   "Sensor"
      EndProperty
      BeginProperty RepeaterBinding(7) {7D21A594-FC9B-11D0-A320-00AA00688B10} 
         _StreamID       =   -1412567295
         _Version        =   65536
         PropertyName    =   "Status"
         DataField       =   "Status"
      EndProperty
   End
   Begin VB.Label Label5 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Local / Loja"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   420
      Left            =   280
      TabIndex        =   3
      Top             =   420
      Width           =   2685
   End
   Begin VB.Label Label7 
      BorderStyle     =   1  'Fixed Single
      Height          =   840
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   280
   End
   Begin VB.Label Label6 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Câmera / Monitor"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   420
      Left            =   2955
      TabIndex        =   4
      Top             =   420
      Width           =   2940
   End
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Evento"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   420
      Left            =   3555
      TabIndex        =   2
      Top             =   0
      Width           =   2340
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Sensor"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   420
      Left            =   1530
      TabIndex        =   1
      Top             =   0
      Width           =   2025
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Horário"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   420
      Left            =   285
      TabIndex        =   0
      Top             =   0
      Width           =   1230
   End
End
Attribute VB_Name = "frmQueue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
   Dim success As Long
   success = SetWindowPos(frmQueue.hWnd, HWND_TOPMOST, ForNet.Width - 6300, ForNet.Height - 5300, 0, 0, FLAGS)
   curhwnd = frmQueue.hWnd
End Sub

Private Sub Form_Load()
   Set tFort = New clsDigiFort
   Set DataRepeater1.DataSource = tFort
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   If Not IsNull(DataRepeater1.CurrentRecord) Then
      Cancel = True
   End If
End Sub

Private Sub DataRepeater1_Click()
   If Not IsNull(DataRepeater1.CurrentRecord) Then
      tFort.Recordset.Bookmark = DataRepeater1.CurrentRecord
      ChangeEvent tFort.Recordset
   End If
   
   If IsNull(DataRepeater1.CurrentRecord) Then
      Me.Hide
   End If
   
End Sub

Private Sub ChangeEvent(ByVal rs As Recordset)
   Dim lModule As clsModule
   Dim txId As String
   Dim txStatus As Integer
   Dim lEntity As clsEntity
   Dim lEvent As clsEvent
   
   txId = rs("UID")
   txStatus = rs("Status")
   Set lModule = lstModule.Item(txId)
   Set lEntity = lstEntity.Item(CStr(lModule.mEntity))
   lModule.evDate = rs("EvTime")
   
   Load frmCritico
   Set frmCritico.crModule = lModule
   frmCritico.Show vbModal
   
   Set lEvent = New clsEvent
   With lEvent
      .sUIDo = lModule.UID
      .evDate = Format(Now, "dd/mm/yyyy hh:mm:ss")
      .evDescr = lEntity.vDescr
      .evCritico = True
      .evTipo = lModule.mTipo
      .evScope = lModule.crScope
      .evAcao = lModule.crAcao
      .evObs = lModule.crObs
      .evTreat = lModule.crTreat
      .evUser = lModule.crUser
   End With
   EventAdd lEvent            'Persist the critical event on Database (modAux)
   rs.Delete adAffectCurrent
   
End Sub

