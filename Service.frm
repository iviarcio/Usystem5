VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{A4749554-0441-4E64-8A03-3323601631C7}#1.0#0"; "LaVolpeAlphaImg2.ocx"
Begin VB.Form frmService 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Serviços de Backup & Restore"
   ClientHeight    =   3225
   ClientLeft      =   1725
   ClientTop       =   1995
   ClientWidth     =   5895
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   Icon            =   "Service.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3225
   ScaleWidth      =   5895
   Begin VB.Frame Frame4 
      Caption         =   "Manter os Dados"
      Height          =   855
      Left            =   0
      TabIndex        =   11
      Top             =   3360
      Width           =   4695
      Begin VB.TextBox txtDados 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   3000
         TabIndex        =   12
         Text            =   "30"
         Top             =   345
         Width           =   495
      End
      Begin VB.Label Label2 
         Caption         =   "Manter os dados dos últimos"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   360
         TabIndex        =   14
         Top             =   360
         Width           =   2775
      End
      Begin VB.Label Label3 
         Caption         =   "dias."
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   3600
         TabIndex        =   13
         Top             =   360
         Width           =   375
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Local do Backup"
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
      Left            =   0
      TabIndex        =   9
      Top             =   1200
      Width           =   5895
      Begin VB.TextBox txtLocal 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Width           =   4575
      End
      Begin LaVolpeAlphaImg.AlphaImgCtl cmdLocal 
         Height          =   720
         Left            =   5040
         ToolTipText     =   "Procurar/Selecionar Local de Backup"
         Top             =   150
         Width           =   720
         _ExtentX        =   1270
         _ExtentY        =   1270
         Image           =   "Service.frx":030A
         Effects         =   "Service.frx":14C9
      End
   End
   Begin VB.CommandButton cmdRestore 
      Caption         =   "&Restore"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4080
      Picture         =   "Service.frx":14E1
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Recuperar a Base  de Dados selecionada"
      Top             =   2400
      Width           =   855
   End
   Begin VB.CommandButton cmdBackup 
      Caption         =   "&Backup"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3120
      Picture         =   "Service.frx":17EB
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Realizar operação de Backup da Base selecionada"
      Top             =   2400
      Width           =   855
   End
   Begin VB.Frame Frame1 
      Caption         =   "Backup"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   975
      Left            =   0
      TabIndex        =   2
      Top             =   120
      Width           =   5895
      Begin VB.CheckBox chkAuto 
         Alignment       =   1  'Right Justify
         Caption         =   "Realizar Backup Automático ?"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   315
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   2655
      End
      Begin MSMask.MaskEdBox mskHorario 
         Height          =   315
         Left            =   4440
         TabIndex        =   4
         Top             =   360
         Visible         =   0   'False
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "99:99"
         PromptChar      =   "_"
      End
      Begin VB.Label lblHorario2 
         Caption         =   "horas"
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
         Left            =   5160
         TabIndex        =   8
         Top             =   390
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label lblHorario 
         Caption         =   "Diariamente as "
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
         TabIndex        =   5
         Top             =   390
         Visible         =   0   'False
         Width           =   1095
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Base de Dados"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   855
      Left            =   0
      TabIndex        =   0
      Top             =   2280
      Width           =   3015
      Begin VB.CheckBox chkCadastro 
         Caption         =   "Cadastro && Eventos"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   480
         TabIndex        =   1
         Top             =   360
         Value           =   1  'Checked
         Width           =   2055
      End
   End
   Begin MSComDlg.CommonDialog cdlSync 
      Left            =   5160
      Top             =   3600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin LaVolpeAlphaImg.AlphaImgCtl cmdExit 
      Height          =   720
      Left            =   5040
      ToolTipText     =   "Fechar Configuração de Backup"
      Top             =   2300
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   1270
      Image           =   "Service.frx":1AF5
      Effects         =   "Service.frx":27FA
   End
End
Attribute VB_Name = "frmService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Private f_bChange As Boolean   'removed from version 1.0.32

Private fDataBase As String
Private fResult As String
Private fBackup As Boolean

Private Sub cmdBackup_Click()
'  provoca a geração de erro se o usuário selecionar "Cancel"
   cdlSync.CancelError = True
   On Error GoTo BackupHandler
'  prepara flags
   cdlSync.FLAGS = cdlOFNExtensionDifferent + cdlOFNOverwritePrompt + _
   cdlOFNPathMustExist + cdlOFNHideReadOnly
'  prepara título para a caixa de diálogo e filtros
   If chkCadastro.Value = vbChecked Then
      cdlSync.DialogTitle = "Indicar pasta e nome para Backup da BAse de Dados"
      cdlSync.fileName = "USystemDB5.mdb"
   End If
   cdlSync.Filter = "Todos (*.*)|*.*|Base de dados " & "(*.mdb)|*.mdb"
'  especifica o filtro padrão
   cdlSync.FilterIndex = 2
'  pasta default
   On Error Resume Next  'NÃO REMOVER!!!
   Dim tmp$, dst$
   If Not (m_sBPath = txtLocal) Then
      m_sBPath = txtLocal
      Call SaveSetting("USystem5", "Options", "Backup", m_sBPath)
   End If
   dst$ = m_sBPath
   tmp$ = Dir(dst$, vbDirectory)
   If tmp$ = "" Then
      MkDir Left$(dst$, Len(dst$))
   End If
   On Error GoTo BackupHandler
   cdlSync.InitDir = dst$
'  mostra o diálogo Save
   cdlSync.ShowSave
   fResult = cdlSync.fileName
   If fResult = "" Then
      Screen.MousePointer = vbDefault
      Exit Sub
   ElseIf (Format$(fResult, ">") = Format$(m_sDatabase, ">")) Then
      Screen.MousePointer = vbDefault
      MsgBox "Não é permitido realizar o backup sobre a base de dados corrente!", sxInformation, sxProname
      Exit Sub
   Else
      Screen.MousePointer = vbHourglass
      cnDB.Close
      DoEvents
      Backup_Restore fBckp:=True, fSilent:=False, fQuery:=True
      DoEvents
      fBackup = True
      Screen.MousePointer = vbDefault
   End If
   Unload Me
   Exit Sub
BackupHandler:
   Screen.MousePointer = vbDefault
   If Err.Number <> 32755 Then
      MsgBox "Error: " & Err.Description, sxInformation, sxProname
   End If
   Exit Sub
End Sub

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


Private Sub cmdLocal_Click()
   Load frmDir
   Set frmDir.fCaller = Me
   frmDir.Caption = "Diretório para Backup Automático"
   frmDir.Show vbModal
   If Not (m_sBPath = txtLocal) Then
      m_sBPath = txtLocal
      Call SaveSetting("USystem5", "Options", "Backup", m_sBPath)
   End If
End Sub

Private Sub cmdLocal_MouseEnter()
   cmdLocal.SetRedraw = False
   cmdLocal.GrayScale = lvicSepia
   cmdLocal.LightnessPct = -20
   cmdLocal.SetRedraw = True
End Sub

Private Sub cmdLocal_MouseExit()
   cmdLocal.SetRedraw = False
   cmdLocal.GrayScale = lvicNoGrayScale
   cmdLocal.LightnessPct = 0
   cmdLocal.SetRedraw = True
End Sub

Private Sub cmdRestore_Click()
   Dim lres%
   lres% = MsgBox("A base de dados corrente será substituída!" & vbCr & vbLf & _
                   "Deseja proseguir? ", sxQuestion, sxProname)
   If lres% = vbYes Then
'     provoca a geração de erro se o usuário selecionar "Cancel"
      cdlSync.CancelError = True
      On Error GoTo RestoreError
'     prepara flags
      cdlSync.FLAGS = cdlOFNHideReadOnly + cdlOFNFileMustExist
'     prepara titulo da caixa de diálogo
      cdlSync.DialogTitle = "Indicar base de dados a restaurar"
'     prepara filtros
      cdlSync.Filter = "Todos (*.*)|*.*|Base de dados " & _
                        "(*.mdb)|*.mdb"
'     especifica o filtro padrão
      cdlSync.FilterIndex = 2
'     diretório default
      If Not (m_sBPath = txtLocal) Then
         m_sBPath = txtLocal
         Call SaveSetting("USystem5", "Options", "Backup", m_sBPath)
      End If
      cdlSync.InitDir = m_sBPath
'     mostra o diálogo Open
      cdlSync.ShowOpen
      fResult = cdlSync.fileName
      If (Format$(fResult, ">") = Format$(m_sDatabase, ">")) Then
         Screen.MousePointer = vbDefault
         MsgBox "Não é permitido realizar o restore a partir da base de dados corrente!", sxInformation, sxProname
         Exit Sub
      End If
      Screen.MousePointer = vbHourglass
      ForNet.Update_Display "Aguarde...", sxImgInform, False
      cnDB.Close
      Backup_Restore fBckp:=False, fSilent:=False, fQuery:=True
      DoEvents
      ForNet.Update_Display "", sxImgNone, False
      Screen.MousePointer = vbDefault
   End If
   Unload Me
   Exit Sub
RestoreError:
   Screen.MousePointer = vbDefault
   ForNet.Update_Display "", sxImgNone, False
   If Err.Number <> 32755 Then
      MsgBox "Error: " & Err.Description, sxInformation, sxProname
   End If
   Exit Sub
End Sub

Private Sub Backup_Restore(ByVal fBckp As Boolean, _
                           ByVal fSilent As Boolean, _
                           ByVal fQuery As Boolean)
   Dim lcad As Boolean, levn As Boolean, lerr As Boolean
   Dim lattr As Integer, lastError As Long
   Dim lmsg As String
   Dim lTemp As String
   Err.Clear
   lastError = Err.Number
   lcad = (chkCadastro.Value = vbChecked)
   Screen.MousePointer = vbHourglass
   ForNet.Update_Display "Aguarde...", sxImgInform, False
   On Error Resume Next
   If fBckp Then
      lTemp = App.Path & "\tmpDB5.mdb"
      lattr = GetAttr(fResult)
      If Err.Number <> 53 Then
         On Error GoTo BckError
         If lattr = vbReadOnly And Not fSilent Then
            If MsgBox("A base de Dados na unidade de destino está" & vbCr & vbLf & _
                      "protegida contra escrita! Realizar Backup assim mesmo?", _
                      sxQuestion, sxProname) = vbYes Then
               SetAttr fResult, vbNormal
               If lcad Then
                  DataBase_Copy fDataBase, lTemp
                  If DataBase_Compact(lTemp, fResult) Then
                     'nothing
                  Else
                     DataBase_Copy lTemp, fResult
                  End If
               End If
               SetAttr fResult, vbReadOnly
            Else
               Err.Raise vbObjectError + 610, , "Cancelado pelo Operador"
               GoTo BckExit
            End If
         Else
            SetAttr fResult, vbNormal
            If lcad Then
               DataBase_Copy fDataBase, lTemp
               If DataBase_Compact(lTemp, fResult) Then
                  'nothing
               Else
                  DataBase_Copy lTemp, fResult
               End If
            End If
            SetAttr fResult, vbReadOnly
         End If
      Else 'Err.number = 53 {file not found}
         On Error GoTo BckError
         If lcad Then
            DataBase_Copy fDataBase, lTemp
            If DataBase_Compact(lTemp, fResult) Then
               'nothing
            Else
               DataBase_Copy lTemp, fResult
            End If
         End If
         SetAttr fResult, vbReadOnly
      End If
   Else    'Restore
      Dim lrs As New ADODB.Recordset
      On Error GoTo BckError
      If lcad Then
         On Error Resume Next
         Kill m_sTmpFileDB
         Name fDataBase As m_sTmpFileDB
         FileCopy fResult, fDataBase
         SetAttr fDataBase, vbNormal
'        Testa se a nova base de dados é do USystem5
         lmsg = "O arquivo selecionado não corresponde à base de dados do USystem5"
         On Error GoTo BckError
         cnDB.Open
         lrs.Open "SELECT Version FROM Admin", cnDB, adOpenStatic, adLockReadOnly
         If Not lrs.EOF Then
            If lrs("Version") <> "DB" & curVersion Then Err.Raise vbObjectError + 620, , lmsg
         Else
            Err.Raise vbObjectError + 620, , lmsg
         End If
         lrs.Close
         cnDB.Close
      End If
   End If
BckExit:
   If fQuery Then
      If fBckp Then
         cnDB.Open
      Else
         DBase_ReOpen fIsRestore:=True
      End If
   End If
   Screen.MousePointer = vbDefault
   ForNet.Update_Display "", sxImgNone, False
   If Not fSilent Then
      If lastError = 0 Then
         If fBckp Then
            MsgBox "Backup realizado com sucesso!", sxInformation, sxProname
         Else
            MsgBox "Restore realizado com sucesso!", sxInformation, sxProname
         End If
      End If
   ElseIf lastError = 0 And fQuery Then
      Make_Service "Operação de Backup Automática", " "
   End If
   Exit Sub
BckError:
   Screen.MousePointer = vbDefault
   ForNet.Update_Display "", sxImgNone, False
   lastError = Err.Number
   If Err.Number = 61 Then
      lmsg = "Não há espaço suficiente na unidade de destino."
   ElseIf (Err.Number = 70) Or (Err.Number = 75) Then
      lmsg = "Unidade protegida contra escrita."
   ElseIf Err.Number = 53 Then
      lmsg = "Arquivo não encontrado."
   ElseIf Err.Number = vbObjectError + 620 Then
      'wrong mdb
      lmsg = Err.Description
      If lcad Then cnDB.Close
   ElseIf Err.Number = 3343 Then
      'nothing, not a mdb file
   Else
      lmsg = Err.Description
   End If
   If Not fSilent Then
      If fBckp Then
         MsgBox "Erro na geração do Backup!" & vbCr & vbLf & lmsg, sxCritical, sxProname
      Else
         MsgBox "Erro na recuperação do Backup!" & vbCr & vbLf & lmsg, sxCritical, sxProname
         On Error Resume Next
         If lcad Then
            Kill fDataBase
            Name m_sTmpFileDB As fDataBase
         End If
      End If
   Else
      Make_Service "Erro no Backup Automático! " & lmsg, " "
   End If
   Resume BckExit
End Sub

Public Sub Automatic_Backup()
   On Error Resume Next  'NÃO REMOVER!!!
   Dim strData As String
   strData = CStr(Day(Now)) & "_" & CStr(Month(Now))
   Dim tmp$, dst$
   dst$ = m_sBPath
   tmp$ = Dir(dst$, vbDirectory)
   If tmp$ = "" Then
      MkDir Left$(dst$, Len(dst$))
   Else
      Kill m_sBPath & "\USystemDB5_bkp.mdb"
      Name m_sBPath & "\USystemDB5_" & strData & ".mdb" As m_sBPath & "\USystemDB5_bkp.mdb"
   End If
      
   chkCadastro.Value = vbChecked
   fResult = m_sBPath & "\USystemDB5_" & strData & ".mdb"
   cnDB.Close
   Backup_Restore fBckp:=True, fSilent:=True, fQuery:=True
   DoEvents
   
End Sub

Private Sub chkAuto_Click()
   lblHorario.Visible = (chkAuto.Value = vbChecked)
   lblHorario2.Visible = (chkAuto.Value = vbChecked)
   mskHorario.Visible = (chkAuto.Value = vbChecked)
   m_bBackupAuto = chkAuto.Value = vbChecked
End Sub

Private Sub Form_Load()
   Dim success As Long
   success = SetWindowPos(frmService.hWnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
   curhwnd = frmService.hWnd
   Left = (Screen.Width - Width) / 2   ' Center form horizontally.
   Top = (Screen.Height - Height) / 2  ' Center form vertically.
   If m_bBackupAuto Then
      chkAuto = vbChecked
      mskHorario = m_sHorario
   Else
      chkAuto = vbUnchecked
   End If
   
   txtLocal = m_sBPath
   fDataBase = m_sPath & "\USystemDB5.mdb"
   fBackup = False
End Sub

Private Sub Form_Unload(Cancel As Integer)

   If Not (m_sBPath = txtLocal) Then
      m_sBPath = txtLocal
      Call SaveSetting("USystem5", "Options", "Backup", m_sBPath)
   End If
End Sub

Private Sub mskHorario_Change()
   m_sHorario = mskHorario
End Sub
