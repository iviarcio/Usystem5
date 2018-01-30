VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{A4749554-0441-4E64-8A03-3323601631C7}#1.0#0"; "LaVolpeAlphaImg2.ocx"
Begin VB.Form frmReport 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Relatórios Gerais"
   ClientHeight    =   3495
   ClientLeft      =   2265
   ClientTop       =   2565
   ClientWidth     =   6105
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
   Icon            =   "Report.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3495
   ScaleWidth      =   6105
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraTipo 
      Caption         =   "Tipo de Evento"
      ForeColor       =   &H00800000&
      Height          =   1695
      Left            =   2565
      TabIndex        =   1
      Top             =   15
      Width           =   3420
      Begin VB.CheckBox chkTipo 
         Caption         =   "Abertura/Fechamento"
         Height          =   315
         Index           =   6
         Left            =   120
         TabIndex        =   10
         Top             =   1320
         Width           =   2535
      End
      Begin VB.CheckBox chkTipo 
         Caption         =   "Operação"
         Height          =   315
         Index           =   5
         Left            =   2145
         TabIndex        =   7
         Top             =   960
         Width           =   1095
      End
      Begin VB.CheckBox chkTipo 
         Caption         =   "Sistema"
         Height          =   315
         Index           =   4
         Left            =   2145
         TabIndex        =   6
         Top             =   615
         Width           =   1095
      End
      Begin VB.CheckBox chkTipo 
         Caption         =   "Incêndio"
         Height          =   315
         Index           =   0
         Left            =   2145
         TabIndex        =   5
         Top             =   255
         Width           =   1095
      End
      Begin VB.CheckBox chkTipo 
         Caption         =   "Emergência"
         Height          =   315
         Index           =   2
         Left            =   120
         TabIndex        =   4
         Top             =   960
         Width           =   1215
      End
      Begin VB.CheckBox chkTipo 
         Caption         =   "Pânico"
         Height          =   315
         Index           =   3
         Left            =   120
         TabIndex        =   3
         Top             =   600
         Width           =   1095
      End
      Begin VB.CheckBox chkTipo 
         Caption         =   "Intrusão"
         Height          =   315
         Index           =   1
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Value           =   1  'Checked
         Width           =   1095
      End
   End
   Begin VB.Frame fraData 
      Caption         =   "Data da Pesquisa"
      ForeColor       =   &H00800000&
      Height          =   1695
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2505
      Begin MSMask.MaskEdBox mskInicial 
         Height          =   315
         Left            =   840
         TabIndex        =   8
         Top             =   480
         Visible         =   0   'False
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskFinal 
         Height          =   315
         Left            =   840
         TabIndex        =   9
         Top             =   1080
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label lblInicial 
         Caption         =   "Inicial:"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   510
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label lblFinal 
         Caption         =   "Final: "
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   360
         TabIndex        =   12
         Top             =   1110
         Width           =   375
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Local"
      ForeColor       =   &H00800000&
      Height          =   855
      Left            =   15
      TabIndex        =   11
      Top             =   1755
      Width           =   5970
      Begin VB.ComboBox lstLocal 
         BackColor       =   &H0080FFFF&
         Height          =   315
         Left            =   405
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   330
         Width           =   5220
      End
   End
   Begin LaVolpeAlphaImg.AlphaImgCtl cmdCleanUp 
      Height          =   720
      Left            =   120
      ToolTipText     =   "Limpar base de dados de Eventos"
      Top             =   2640
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   1270
      Image           =   "Report.frx":0442
      Effects         =   "Report.frx":1743
   End
   Begin LaVolpeAlphaImg.AlphaImgCtl cmdPrint 
      Height          =   720
      Left            =   4440
      ToolTipText     =   "Visualizar Relatório de acordo com a seleção"
      Top             =   2685
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   1270
      Image           =   "Report.frx":175B
      Effects         =   "Report.frx":2906
   End
   Begin LaVolpeAlphaImg.AlphaImgCtl cmdExit 
      Height          =   720
      Left            =   5325
      ToolTipText     =   "Fechar geração de relatórios"
      Top             =   2685
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   1270
      Image           =   "Report.frx":291E
      Effects         =   "Report.frx":3623
   End
   Begin VB.Image imgCleanUp 
      Height          =   480
      Left            =   1020
      Picture         =   "Report.frx":363B
      Top             =   2805
      Width           =   480
   End
   Begin VB.Label lblCleanUp 
      Caption         =   "obs.: Para a limpeza só é  considerada a data final!"
      ForeColor       =   &H00800000&
      Height          =   480
      Left            =   1560
      TabIndex        =   14
      Top             =   2835
      Width           =   1905
   End
End
Attribute VB_Name = "frmReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lInit0 As String
Dim lEnd0 As String
Dim lok As Boolean
Dim lInit As String
Dim lEnd As String
Dim success As Long
Dim frm As New frmViewReport9

Private mList As XArrayDB
Private fNoUpdate As Boolean

Private Sub chkTipo_Click(Index As Integer)
   Dim i As Integer
   If Index = 5 Then
      If chkTipo(Index).Value = vbChecked Then
         fNoUpdate = True
         For i = 0 To 4
            chkTipo(i).Value = vbUnchecked
         Next i
         chkTipo(6).Value = vbUnchecked
         fNoUpdate = False
         lblInicial.Visible = True
         mskInicial.Visible = True
         lstLocal.Enabled = False
         lstLocal.ListIndex = 0
      Else
         lstLocal.Enabled = True
      End If
   ElseIf Index = 6 Then
      If chkTipo(Index).Value = vbChecked Then
         fNoUpdate = True
         For i = 0 To 5
            chkTipo(i).Value = vbUnchecked
         Next i
         fNoUpdate = False
         lblInicial.Visible = lstLocal.ListIndex <> 0
         mskInicial.Visible = lstLocal.ListIndex <> 0
      Else
         lstLocal.Enabled = True
'         lblInicial.Visible = True
'         mskInicial.Visible = True
      End If
   ElseIf Not fNoUpdate Then
      chkTipo(5).Value = vbUnchecked
      chkTipo(6).Value = vbUnchecked
      lblInicial.Visible = lstLocal.ListIndex <> 0
      mskInicial.Visible = lstLocal.ListIndex <> 0
   End If
End Sub

Private Sub cmdCleanUp_Click()
   Dim success As Long, lInterval As Long, pastDate As String, lWhere As String
   success = SetWindowPos(frmReport.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, FLAGS)
   If MsgBox("Confirma a limpeza dos registros da Base de dados conforme a seleção?", sxQuestion, sxProname) = vbYes Then
      If IsDate(mskFinal) Then
         lInterval = DateDiff("d", mskFinal, Date)
         If lInterval >= 0 Then
            'Note: To pass date value in "#" format to Access, you need to represent
            'this value in english format, i.e., mm/dd/yyyy.
            pastDate = Format$(DateAdd("d", -lInterval, Date), "mm/dd/yyyy")
         Else
            MsgBox "Data Final não é válida!", sxExclamation, sxProname
            success = SetWindowPos(frmReport.hWnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
            curhwnd = frmReport.hWnd
            Exit Sub
         End If
      Else
         MsgBox "Data Final não é válida!", sxExclamation, sxProname
         success = SetWindowPos(frmReport.hWnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
         curhwnd = frmReport.hWnd
         Exit Sub
      End If
      
      Screen.MousePointer = vbHourglass
      Dim lcmd As New ADODB.Command
      Set lcmd.ActiveConnection = cnDB
      lcmd.CommandType = adCmdText
      If chkTipo(6).Value = vbChecked Then
         If lstLocal.ItemData(lstLocal.ListIndex) = 0 Then
            lWhere = ""
         Else
            lWhere = " AND fk_Entity = " & lstLocal.ItemData(lstLocal.ListIndex)
         End If
         Dim ldif As Long
         ldif = DateDiff("d", CDate("01/01/2000"), Date) - lInterval
         If ldif > 0 Then
            lcmd.ActiveConnection.BeginTrans
            lcmd.CommandText = "DELETE FROM AccessOpen WHERE (Report < " & ldif & lWhere & ")"
            lcmd.Execute
            lcmd.CommandText = "DELETE FROM AccessClose WHERE (Report < " & ldif & lWhere & ")"
            lcmd.Execute
            lcmd.ActiveConnection.CommitTrans
         End If
      ElseIf chkTipo(5).Value = vbChecked Then
         lcmd.ActiveConnection.BeginTrans
         lcmd.CommandText = "DELETE FROM Service WHERE (Service.Date_Service < #" & pastDate & "#)"
         lcmd.Execute
         lcmd.ActiveConnection.CommitTrans
      Else
         Dim i As Integer, j As Integer
         Dim ltipo As Integer
         ltipo = 0
         j = 1
         For i = 0 To 4
            If chkTipo(i).Value = vbChecked Then
               ltipo = ltipo + j
            End If
            j = j * 2
         Next i
         ChooseEvent_CleanUp fInterval:=lInterval, fTipo:=ltipo, _
                             fClient:=lstLocal.ItemData(lstLocal.ListIndex)
      End If
      Screen.MousePointer = vbDefault
      MsgBox "Limpeza de eventos na base de dados executada com sucesso", sxExclamation, sxProname
   End If
   success = SetWindowPos(frmReport.hWnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
   curhwnd = frmReport.hWnd
End Sub

Private Sub cmdCleanUp_MouseEnter()
   cmdCleanUp.SetRedraw = False
   cmdCleanUp.GrayScale = lvicSepia
   cmdCleanUp.LightnessPct = -20
   cmdCleanUp.SetRedraw = True
End Sub

Private Sub cmdCleanUp_MouseExit()
   cmdCleanUp.SetRedraw = False
   cmdCleanUp.GrayScale = lvicNoGrayScale
   cmdCleanUp.LightnessPct = 0
   cmdCleanUp.SetRedraw = True
End Sub

Private Sub ChooseEvent_CleanUp(ByVal fInterval As Long, ByVal fTipo As Integer, _
                                ByVal fClient As Integer)
   Dim lcmd As New ADODB.Command
   Set lcmd.ActiveConnection = cnDB
   lcmd.ActiveConnection.BeginTrans
   lcmd.CommandType = adCmdText
   'Note: To pass date value in "#" format to Access, you need to represent
   'this value in english format, i.e., mm/dd/yyyy.
   Dim pastDate As String
   pastDate = Format$(DateAdd("d", -fInterval, Date), "mm/dd/yyyy")
   If fTipo >= 31 Then
      If fClient = 0 Then
         lcmd.CommandText = "DELETE FROM Event WHERE (Event.Date_Event < #" & pastDate & "#)"
      Else
         lcmd.CommandText = "DELETE FROM Event WHERE (Event.Date_Event < #" & _
                             pastDate & "# AND Event.fk_Entity = " & fClient & ")"
      End If
      lcmd.Execute
   ElseIf fTipo > 0 Then
      Dim i As Integer, j As Integer
      Dim lWhere As String
      Dim lcomp As String
      lcomp = " AND ("
      lWhere = "DELETE FROM Event WHERE (Event.Date_Event < #" & pastDate & "# "
      j = 16
      For i = 4 To 0 Step -1
         If fTipo \ j > 0 Then
            lWhere = lWhere & lcomp & "Event.Tipo_Sensor = '" & strTipo(i) & "'"
            lcomp = " OR "
         End If
         fTipo = fTipo Mod j
         j = j / 2
      Next i
      If fClient = 0 Then
         lcmd.CommandText = lWhere & "))"
      Else
         lcmd.CommandText = lWhere & ") AND Event.fk_Entity = " & fClient & ")"
      End If
      lcmd.Execute
   End If
   lcmd.ActiveConnection.CommitTrans
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

Private Sub cmdPrint_Click()
      
   'Relatórios de Abertura e Fechamento (Ùnico e Todos)
   If chkTipo(6).Value = vbChecked Then
      
      lok = False
      lok = Access_Prepare(lstLocal.ListIndex = 0)
       
      If lstLocal.ListIndex <> 0 Then
         frm.SetTipo = g_iRptAFUnico
         If lok = True Then
            Dim lSelection As String
            lSelection = frm.SetSelection
            frm.SetSelection = lSelection & " AND ({fk_Entity}= " & lstLocal.ItemData(lstLocal.ListIndex) & ")"
            SetHourGlass Me
            frm.WindowState = vbMaximized
            frm.Show
            ResetMouse Me
         Else
            MsgBox "Erro na geração do relatório!", vbOKOnly + vbInformation, USVersion
         End If
      Else
         If lok = True Then
            SetHourGlass Me
            frm.SetTipo = g_iRptAFTodos
            SetHourGlass Me
            frm.WindowState = vbMaximized
            frm.Show
            ResetMouse Me
         End If
      
      End If
   
   'Relatório de Operação
   ElseIf chkTipo(5).Value = vbChecked Then
        
      frm.SetTipo = g_iRptOperacao
      lok = Service_Prepare(lInit, lEnd)
      If lok = True Then
         SetHourGlass Me
         frm.WindowState = vbMaximized
         frm.Show
         ResetMouse Me
      Else
         MsgBox "Erro na geração do relatório!", vbOKOnly + vbInformation, USVersion
      End If
     
   'Relatório de Eventos (Único e Todos)
   Else
      
      lok = Event_Prepare
      If lok = True Then
        If lstLocal.ListIndex <> 0 Then
           frm.SetTipo = g_iRptEventosUnico
           SetHourGlass Me
           frm.WindowState = vbMaximized
           frm.Show
           ResetMouse Me
        Else
           frm.SetTipo = g_iRptEventos
           SetHourGlass Me
           frm.WindowState = vbMaximized
           frm.Show
           ResetMouse Me
        End If
       End If
   End If
   
End Sub
   
Private Sub cmdPrint_MouseEnter()
   cmdPrint.SetRedraw = False
   cmdPrint.GrayScale = lvicSepia
   cmdPrint.LightnessPct = -20
   cmdPrint.SetRedraw = True
End Sub

Private Sub cmdPrint_MouseExit()
   cmdPrint.SetRedraw = False
   cmdPrint.GrayScale = lvicNoGrayScale
   cmdPrint.LightnessPct = 0
   cmdPrint.SetRedraw = True
End Sub

  
Private Function Event_Prepare() As Boolean
   Dim lok As Boolean, hasDate As Boolean, lformula As String
   
   lok = True
   hasDate = False
   
   If mskInicial.ClipText = "" Or lstLocal.ListIndex = 0 Then
      
      If IsDate(mskFinal) Then
         Dim lSignal As String
         If lstLocal.ListIndex <> 0 Then
            lSignal = " <= "
         Else
            lSignal = " = "
         End If
         
         frm.SetSelection = "{Date_Event}" & lSignal & ExtractDate(mskFinal)
         hasDate = True
      Else
         MsgBox "Data Final não é válida!", sxExclamation, sxProname
         lok = False
      End If
      
   ElseIf IsDate(mskFinal) And IsDate(mskInicial) Then
   
      If DateDiff("n", mskInicial, mskFinal) >= 0 Then
         frm.SetSelection = "{Date_Event}>=" & ExtractDate(mskInicial) & " AND " & _
                            "{Date_Event}<=" & ExtractDate(mskFinal)
                            
         hasDate = True
      Else
         MsgBox "Data Inicial maior que a Data Final!", sxExclamation, sxProname
         lok = False
      End If
      
   ElseIf (mskInicial <> "") Or (mskFinal <> "") Then
      MsgBox "Data Inicial ou Final não são válidas!", sxExclamation, sxProname
      lok = False
   End If
   
   Dim hasTipo As Boolean
   hasTipo = False
   Dim i As Integer
   For i = 0 To 4
      If chkTipo(i).Value = vbChecked Then
         If hasTipo Then
            lformula = frm.SetSelection
            frm.SetSelection = lformula & " OR {Tipo_Sensor}='" & strTipo(i) & "'"
         ElseIf hasDate Then
            lformula = frm.SetSelection
            frm.SetSelection = lformula & " AND ({Tipo_Sensor}='" & strTipo(i) & "'"
         Else
            frm.SetSelection = "({Tipo_Sensor}='" & strTipo(i) & "'"
         End If
         hasTipo = True
      End If
   Next i
   
   If hasTipo Then
      lformula = frm.SetSelection
      frm.SetSelection = lformula & ")"
   End If
   
   Event_Prepare = lok
   
End Function

Private Function Service_Prepare(ByRef lInit As String, ByRef lEnd As String) As Boolean
    
    Service_Prepare = False
    
    'Valida os conteudos da data inicial
    If IsDate(mskInicial) And _
        Left(mskInicial, 2) > 0 And Left(mskInicial, 2) < 32 And _
        Mid(mskInicial, 4, 2) > 0 And Mid(mskInicial, 4, 2) < 13 Then
        lInit0 = mskInicial & " " & "00:00:00"
    Else
        Me.Hide
        MsgBox "A data inicial é inválida!", vbOKOnly + vbInformation, USVersion
        Me.Show
        ResetMouse Me
        Exit Function
    End If
      
    'Valida os conteudos da data final
    If IsDate(mskFinal) And _
        Left(mskFinal, 2) > 0 And Left(mskFinal, 2) < 32 And _
        Mid(mskFinal, 4, 2) > 0 And Mid(mskFinal, 4, 2) < 13 Then
        lEnd0 = mskFinal & " " & "23:59:59"
    Else
        Me.Hide
        MsgBox "A data final é inválida!", vbOKOnly + vbInformation, USVersion
        Me.Show
        ResetMouse Me
        Exit Function
    End If

    'Para pesquisa via Cristal Report deve-se inverter o dd/mm para mm/dd
    lInit = Format$(lInit0, "mm/dd/yyyy hh:mm:ss")
    lEnd = Format$(lEnd0, "mm/dd/yyyy hh:mm:ss")
    
    If DateDiff("s", lInit0, lEnd0) >= 0 Then
        'Sai da função sem erro
        frm.SetSelection = "{Date_Service} >= #" & lInit & "# AND {Date_Service} <= #" & lEnd & "#"
        Service_Prepare = True
    Else
        Me.Hide
        MsgBox "A data final deve ser maior ou igual à data inicial!", vbOKOnly + vbInformation, USVersion
        Me.Show
        ResetMouse Me
        Exit Function
    End If

End Function

Private Function Access_Prepare(ByVal fAll As Boolean) As Boolean

   Access_Prepare = False
   
   If IsDate(mskFinal) Then
   
      If DateDiff("d", Now, mskFinal) = 0 Then
         If DateDiff("n", Time, m_dTOpen(curWeekday)) > 0 Then
            'nothing, report already made
         Else
            'Make temporary dump
            Me.MousePointer = vbHourglass
            Dump_Entity_Status fOpen:=True
            Me.MousePointer = vbDefault
         End If
         If DateDiff("n", Time, m_dTClose(curWeekday)) > 0 Then
            'nothing, report already made
         ElseIf DateDiff("n", Time, m_dTClose(curWeekday)) > 0 Then
            'Make temporary dump
            Me.MousePointer = vbHourglass
            Dump_Entity_Status fOpen:=False
            Me.MousePointer = vbDefault
         End If
      End If
      
      Dim lreport1 As Long, lreport2 As Long
      
      If fAll Then
         lreport1 = DateDiff("d", CDate("01/01/2000"), mskFinal)
         frm.SetSelection = "{Report}=" & lreport1
      ElseIf Not IsDate(mskInicial) Then
         lreport1 = DateDiff("d", CDate("01/01/2000"), mskFinal)
         frm.SetSelection = "{Report}<=" & lreport1
      Else
         lreport1 = DateDiff("d", CDate("01/01/2000"), mskInicial)
         lreport2 = DateDiff("d", CDate("01/01/2000"), mskFinal)
         frm.SetSelection = "{Report}>=" & lreport1 & _
                                 " AND {Report}<= " & lreport2
      End If
      
      Dim lOL As String, lOH As String, lCL As String, lCH As String
      GetCurTimes lOL, lOH, lCL, lCH
      frm.OpenLow = lOL
      frm.OpenHigh = lOH
      frm.CloseLow = lCL
      frm.CloseHigh = lCH
      Access_Prepare = True
      
'      rpt1.ParameterFields(0) = "OpenLow;" & lOL & ";TRUE"
'      rpt1.ParameterFields(1) = "OpenHigh;" & lOH & ";TRUE"
'      rpt1.ParameterFields(2) = "CloseLow;" & lCL & ";TRUE"
'      rpt1.ParameterFields(3) = "CloseHigh;" & lCH & ";TRUE"

   Else
   
      MsgBox "Data fornecida não é válida!", sxExclamation, sxProname
      
   End If
   
End Function

Private Sub Form_Activate()
   cmdCleanUp.Visible = (m_tAccess = sxSystem)
   imgCleanUp.Visible = (m_tAccess = sxSystem)
   lblCleanUp.Visible = (m_tAccess = sxSystem)
   cmdCleanUp.Enabled = True
   fNoUpdate = False
   Set mList = New XArrayDB
   mList.ReDim 0, lstEntity.Count - 1, 0, 1
   Dim mRow As Integer
   mRow = 0
   Dim cE As clsEntity
   For Each cE In lstEntity
      mList(mRow, 0) = cE.vId
      mList(mRow, 1) = cE.vDescr
      mRow = mRow + 1
   Next
   mList.QuickSort 0, mRow - 1, 1, XORDER_ASCEND, XTYPE_STRING
   lstLocal.Visible = False
   lstLocal.Clear
   lstLocal.AddItem "TODOS"
   lstLocal.ItemData(lstLocal.NewIndex) = 0
   Dim Index As Integer
   For Index = 0 To mRow - 1
      lstLocal.AddItem mList(Index, 1)
      lstLocal.ItemData(lstLocal.NewIndex) = mList(Index, 0)
   Next
   lstLocal.ListIndex = 0
   lstLocal.Visible = True
End Sub

Private Sub Form_Load()
   Dim success As Long
   success = SetWindowPos(frmReport.hWnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
   curhwnd = frmReport.hWnd
   Left = (Screen.Width - Width) / 2   ' Center form horizontally.
   Top = (Screen.Height - Height) / 2  ' Center form vertically.
   mskFinal = Format$(Date, "dd/mm/yyyy")
   mskInicial = mskFinal
End Sub

Private Function ExtractDate(fText As String) As String
   Dim lPos As Integer
   lPos = InStr(1, fText, " ")
   If lPos <> 0 Then
      Dim lDate As String
      lDate = CDate(Left$(fText, lPos))
   Else
      lDate = CDate(fText)
   End If
   ExtractDate = "Date(" & Year(lDate) & ", " & Month(lDate) & ", " & Day(lDate) & ")"
End Function

Private Function ExtractHour(fText As String, ByVal fInit As Boolean) As String
   Dim lPos As Integer
   lPos = InStr(1, fText, " ")
   If lPos <> 0 Then
      Dim lDate As String
      lDate = CDate(Left$(fText, lPos))
   Else
      lDate = CDate(fText)
   End If
   If fInit Then
      ExtractHour = "DateTime(" & Year(lDate) & ", " & Month(lDate) & ", " & Day(lDate) & _
                  ", 00, 00, 00)"
   Else
      ExtractHour = "DateTime(" & Year(lDate) & ", " & Month(lDate) & ", " & Day(lDate) & _
                  ", " & Hour(Now()) & ", " & Minute(Now()) & ", 00)"
   End If
End Function

Private Sub GetCurTimes(fOL As String, fOH As String, fCL As String, fCH As String)
   Dim lrs As New ADODB.Recordset
   lrs.Open "SELECT * FROM Horario", cnDB, adOpenStatic, adLockReadOnly
   If Weekday(mskFinal) = vbSunday Then
      fOL = lrs("DomOpenInit")
      fOH = lrs("DomOpenEnd")
      fCL = lrs("DomCloseInit")
      fCH = lrs("DomCloseEnd")
   ElseIf Weekday(mskFinal) = vbSaturday Then
      fOL = lrs("SabOpenInit")
      fOH = lrs("SabOpenEnd")
      fCL = lrs("SabCloseInit")
      fCH = lrs("SabCloseEnd")
   Else
      fOL = lrs("SegSexOpenInit")
      fOH = lrs("SegSexOpenEnd")
      fCL = lrs("SegSexCloseInit")
      fCH = lrs("SegSexCloseEnd")
   End If
   lrs.Close
End Sub

Private Sub lstLocal_Click()
   If chkTipo(5).Value = vbUnchecked Then
      lblInicial.Visible = lstLocal.ListIndex <> 0
      mskInicial.Visible = lstLocal.ListIndex <> 0
   Else
      lblInicial.Visible = True
      mskInicial.Visible = True
   End If
End Sub

