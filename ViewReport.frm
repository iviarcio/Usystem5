VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form frmViewReport 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   9120
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13455
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "ViewReport.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   9120
   ScaleWidth      =   13455
   Begin VB.TextBox txtImpressora 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   250
      Left            =   7260
      TabIndex        =   7
      Text            =   "Impressora Selecionada"
      Top             =   45
      Width           =   4290
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   250
      Left            =   6120
      TabIndex        =   6
      Text            =   "Imprimindo em:"
      Top             =   45
      Width           =   1080
   End
   Begin VB.ComboBox cboPrinterDuplex 
      Height          =   315
      Left            =   9450
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   720
      Visible         =   0   'False
      Width           =   2115
   End
   Begin VB.ComboBox cboPaperOrientation 
      Height          =   315
      Left            =   9750
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   315
      Width           =   1170
   End
   Begin VB.ComboBox cboPaperSize 
      Height          =   315
      ItemData        =   "ViewReport.frx":030A
      Left            =   7650
      List            =   "ViewReport.frx":030C
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   315
      Width           =   2010
   End
   Begin VB.ComboBox cboPaperSource 
      Height          =   315
      Left            =   11010
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   315
      Width           =   2115
   End
   Begin VB.PictureBox pctPrinterSettings 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   4965
      Picture         =   "ViewReport.frx":030E
      ScaleHeight     =   330
      ScaleWidth      =   360
      TabIndex        =   1
      ToolTipText     =   "Printer Setup"
      Top             =   255
      Visible         =   0   'False
      Width           =   360
   End
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   9615
      Left            =   -60
      TabIndex        =   0
      Top             =   0
      Width           =   13440
      DisplayGroupTree=   0   'False
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   0   'False
      EnableNavigationControls=   -1  'True
      EnableStopButton=   0   'False
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   0   'False
      EnableSearchControl=   0   'False
      EnableRefreshButton=   0   'False
      EnableDrillDown =   0   'False
      EnableAnimationControl=   0   'False
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   0   'False
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   -1  'True
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
   End
End
Attribute VB_Name = "frmViewReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private intTipoRpt As Integer
Private intTipoAction As Integer
Private intWho As Long
Private strSelection As String
Private strDate As String

Private Sub Form_Load()
    
   Dim success As Long

   Left = 50
   Top = 50
   
   Screen.MousePointer = vbHourglass
   DoEvents
   
   If oCnn Is Nothing Then
         Set oCnn = New clsConnection
   End If
   
   'Esconde todos os Pisos existentes
   Dim rFloor As New ADODB.Recordset
   Set rFloor = oCnn.ExecSpGetRsCa("SELECT * FROM Floor")
   While Not rFloor.EOF
'        lblPiso(m_iCurPiso).Visible = False
        rFloor.MoveNext
   Wend
   rFloor.Close
      
   Select Case intTipoRpt
   
        Case g_iRptCLocais:
            success = SetWindowPos(frmCadastro.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, FLAGS)
            Hide

            Set adors = oCnn.ExecSpGetRsCa("SELECT * FROM Entity")
            Set ADOrs1 = oCnn.ExecSpGetRsCa("Select * From Floor")
            Set Report = New RptCLocais
            Report.ReportTitle = "Relatório de Cadastro de Locais"
            Me.Caption = "Relatório de Cadastro de Locais"
            Set CrDatabase = Report.Database
            Set CrDatabaseTables = CrDatabase.Tables
            Set CrDatabaseTable = CrDatabaseTables.Item(1)
            CrDatabaseTable.SetPrivateData 3, adors
            Set CrDatabaseTable = CrDatabaseTables.Item(2)
            CrDatabaseTable.SetPrivateData 3, ADOrs1
            
        Case g_iRptCAtuadores:
            success = SetWindowPos(frmCAtuador.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, FLAGS)
            Hide

            Set adors = oCnn.ExecSpGetRsCa("SELECT * FROM Entity")
            Set ADOrs1 = oCnn.ExecSpGetRsCa("Select * From Floor")
            Set ADOrs2 = oCnn.ExecSpGetRsCa("Select * From Sensor")
            Set Report = New RptCAtuadores
            Report.ReportTitle = "Relatório de Cadastro de Atuadores"
            Me.Caption = "Relatório de Cadastro de Atuadores"
            Set CrDatabase = Report.Database
            Set CrDatabaseTables = CrDatabase.Tables
            Set CrDatabaseTable = CrDatabaseTables.Item(1)
            CrDatabaseTable.SetPrivateData 3, adors
            Set CrDatabaseTable = CrDatabaseTables.Item(2)
            CrDatabaseTable.SetPrivateData 3, ADOrs1
            Set CrDatabaseTable = CrDatabaseTables.Item(3)
            CrDatabaseTable.SetPrivateData 3, ADOrs2
                  
        Case g_iRptEventos:
            Set adors = oCnn.ExecCmdGetRsEv("SELECT * FROM Event")
            Set Report = New RptEventos
            Report.ReportTitle = "Relatório de Abertura e Fechamento"
            Me.Caption = "Relatório de Abertura e Fechamento"
            Report.Database.SetDataSource adors

        Case g_iRptCZonas
            success = SetWindowPos(frmCZonas.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, FLAGS)
            Hide

            Set adors = oCnn.ExecSpGetRsCa("SELECT * FROM Entity")
            Set ADOrs1 = oCnn.ExecSpGetRsCa("Select * From Floor")
            Set ADOrs2 = oCnn.ExecSpGetRsCa("Select * From Sensor")
            Set Report = New RptCZonas
            Report.ReportTitle = "Relatório de Cadastro de Zonas"
            Me.Caption = "Relatório de Cadastro de Zonas"
            Set CrDatabase = Report.Database
            Set CrDatabaseTables = CrDatabase.Tables
            Set CrDatabaseTable = CrDatabaseTables.Item(1)
            CrDatabaseTable.SetPrivateData 3, adors
            Set CrDatabaseTable = CrDatabaseTables.Item(2)
            CrDatabaseTable.SetPrivateData 3, ADOrs1
            Set CrDatabaseTable = CrDatabaseTables.Item(3)
            CrDatabaseTable.SetPrivateData 3, ADOrs2
            
       Case g_iRptCServicos:
            success = SetWindowPos(frmCAction.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, FLAGS)
            Hide
            
            Set adors = oCnn.ExecSpGetRsCa("SELECT * FROM Actions")
            Set Report = New RptServicos
            Report.ReportTitle = "Relatório de serviços programados"
            Me.Caption = "Relatório de serviços programados"
            Report.Database.SetDataSource adors
            
        Case g_iRptCRonda
            success = SetWindowPos(frmCRonda.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, FLAGS)
            Hide

            Set adors = oCnn.ExecSpGetRsCa("SELECT * FROM Percurso")
            Set ADOrs1 = oCnn.ExecSpGetRsCa("Select * From Ronda")
            Set ADOrs2 = oCnn.ExecSpGetRsCa("Select * From RdHorario")
            Set ADOrs3 = oCnn.ExecSpGetRsCa("Select * From Entity")
            Set Report = New RptCRonda
            Report.ReportTitle = "Relatório de Cadastro de Rondas"
            Me.Caption = "Relatório de Cadastro de Rondas"
            Set CrDatabase = Report.Database
            Set CrDatabaseTables = CrDatabase.Tables
            Set CrDatabaseTable = CrDatabaseTables.Item(1)
            CrDatabaseTable.SetPrivateData 3, adors
            Set CrDatabaseTable = CrDatabaseTables.Item(2)
            CrDatabaseTable.SetPrivateData 3, ADOrs1
            Set CrDatabaseTable = CrDatabaseTables.Item(3)
            CrDatabaseTable.SetPrivateData 3, ADOrs2
            Set CrDatabaseTable = CrDatabaseTables.Item(4)
            CrDatabaseTable.SetPrivateData 3, ADOrs3
            

'      Case g_iRptAFUnico:
'         Set adors = oCnn.ExecSpGetRs("SELECT AccessOpen.*, AccessClose.* " & _
'                     "FROM AccessOpen INNER JOIN AccessClose ON (AccessOpen.fk_Entity = AccessClose.fk_Entity) " & _
'                     "AND (AccessOpen.Report = AccessClose.Report);")
'         Set Report = New rptAFUnico
'         Report.ReportTitle = "Relatório de Abertura e Fechamento"
'         Me.Caption = "Relatório de Abertura e Fechamento"
'         Report.Database.SetDataSource adors
      
'      Case g_iRptAberto:
'         Set adors = oCnn.ExecSpGetRs("Select * From FichaCadastro")
'         Set Report = New rptAberto
'         Report.ReportTitle = "Clientes Abertos"
'         Me.Caption = "Relatório de Clientes Abertos"
'         Report.Database.SetDataSource adors
'
'      Case g_iRptFechado:
'         Set adors = oCnn.ExecSpGetRs("Select * From FichaCadastro")
'         Set Report = New rptFechado
'         Report.ReportTitle = "Clientes Fechados"
'         Me.Caption = "Relatório de Clientes Fechados"
'         Report.Database.SetDataSource adors
'
'      Case g_iRptResume:
'         Dim ADOrs7(0 To 1) As ADODB.Recordset
'         Dim ADOrs6 As ADODB.Recordset
'         Set adors = oCnn.ExecSpGetRs("Select * From Cadastro")
'         Set ADOrs1 = oCnn.ExecSpGetRs("Select * From Municipio")
'         Set ADOrs6 = oCnn.ExecSpGetRs("Select * From Manuais")
'         Set ADOrs7(0) = oCnn.ExecSpGetRs("Select * From Contatos")
'         Set ADOrs7(1) = oCnn.ExecSpGetRs("Select * From Acesso")
'         Set Report = New rptResume
'         Report.ReportTitle = "Ficha Resumida do Cliente"
'         Me.Caption = "Ficha Resumida do Cliente"
'         Set CrDatabase = Report.Database
'         Set CrDatabaseTables = CrDatabase.Tables
'         Set CrDatabaseTable = CrDatabaseTables.Item(1)
'         CrDatabaseTable.SetPrivateData 3, adors
'         Set CrDatabaseTable = CrDatabaseTables.Item(2)
'         CrDatabaseTable.SetPrivateData 3, ADOrs1
'         Set CrDatabaseTable = CrDatabaseTables.Item(3)
'         CrDatabaseTable.SetPrivateData 3, ADOrs6
'         'looping through each section and each reportobject in each section.
'         'when it finds a subreport object, it sets it to a report object
'         'the data source is then assigned using .SetPrivateData method
'         Set CrSections = Report.Sections
'         Dim XX As Integer, YY As Integer, ZZ As Integer
'         ZZ = 0
'         For XX = 1 To CrSections.Count
'            Set CrSection = CrSections.Item(XX)
'            Set CrReportObj = CrSection.ReportObjects
'            For YY = 1 To CrReportObj.Count
'               If CrReportObj.Item(YY).Kind = crSubreportObject Then
'                  Set CrSubreportObj = CrReportObj.Item(YY)
'                  Set crSubreport = CrSubreportObj.OpenSubreport
'                  'We need to set the database, database tables and database table of the subreport
'                  'to get to the method that we need: SetPrivateData
'                  Set CrDatabase = crSubreport.Database
'                  Set CrDatabaseTables = CrDatabase.Tables
'                  Set CrDatabaseTable = CrDatabaseTables.Item(1)
'                  CrDatabaseTable.SetPrivateData 3, ADOrs7(ZZ)
'                  ZZ = ZZ + 1
'               End If
'            Next
'         Next
'
'      Case g_iRptClient:
'         Set adors = oCnn.ExecSpGetRs("Select * From FichaCadastro")
'         Set Report = New rptClient
'         Report.ReportTitle = "Lista de Clientes"
'         Me.Caption = "Lista de Clientes"
'         Report.Database.SetDataSource adors
'
'      Case g_iRptAtividade:
'         Set adors = oCnn.ExecSpGetRs("Select * From ListaInativos")
'         Set Report = New rptAtividade
'         Report.ReportTitle = "Relatório de Clientes Inativos"
'         Me.Caption = "Relatório de Clientes Inativos"
'         Report.Database.SetDataSource adors
'
'      Case g_iRptLastEvent:
'         If intWho <> 0 Then
'            Set adors = oCnn.ExecCmdGetRs("SELECT TOP 40 Event.Codigo, Event.Cliente, " & _
'                        "Event.Tipo, Event.EventDescr, Event.DateStr " & _
'                        "FROM Event WHERE (((Event.Codigo)=" & intWho & ")) ORDER BY " & _
'                        "Event.DateStr DESC;")
'         Else
'            Set adors = oCnn.ExecCmdGetRs("SELECT TOP 100 Event.Codigo, Event.Cliente, " & _
'                        "Event.Tipo, Event.EventDescr, Event.DateStr " & _
'                        "FROM Event ORDER BY Event.DateStr DESC;")
'         End If
'         Set Report = New rptLastEvents
'         Report.ReportTitle = "Últimos Eventos"
'         Me.Caption = "Últimos Eventos"
'         Report.Database.SetDataSource adors
'
'      Case g_iRptLabel:
'         Set adors = oCnn.ExecSpGetRs("Select * From Etiquetas")
'         Set Report = New rptLabel
'         Report.ReportTitle = "Etiquetas de mala direta"
'         Me.Caption = "Etiquetas de mala direta"
'         Report.Database.SetDataSource adors
'
'      Case g_iRptOS:
'         Set adors = oCnn.ExecSpGetRs("Select * From ListaOS")
'         Set ADOrs1 = oCnn.ExecSpGetRs("Select * From Municipio")
'         Set Report = New rptOS
'         Report.ReportTitle = "Ordem de Serviço do Cliente"
'         Me.Caption = "Ordem de Serviço do Cliente"
'         Set CrDatabase = Report.Database
'         Set CrDatabaseTables = CrDatabase.Tables
'         Set CrDatabaseTable = CrDatabaseTables.Item(1)
'         CrDatabaseTable.SetPrivateData 3, adors
'         Set CrDatabaseTable = CrDatabaseTables.Item(2)
'         CrDatabaseTable.SetPrivateData 3, ADOrs1
'
'      Case g_iRptPendentes:
'         Set adors = oCnn.ExecSpGetRs("Select * From Cadastro")
'         Set Report = New rptPendentes
'         Report.ReportTitle = "O.S. Pendentes"
'         Me.Caption = "O.S. Pendentes"
'         Report.Database.SetDataSource adors
'
'      Case g_iRptBasico:
'         Set adors = oCnn.ExecSpGetRs("Select * From EventRadionics")
'         Set Report = New rptBasico
'         Report.ReportTitle = "Cadastro de Eventos Básicos"
'         Me.Caption = "Cadastro de Eventos Básicos"
'         Report.Database.SetDataSource adors
'
'      Case g_iRptContact:
'         Set adors = oCnn.ExecSpGetRs("Select * From EventContactId")
'         Set Report = New rptContact
'         Report.ReportTitle = "Cadastro de Eventos ContactId"
'         Me.Caption = "Cadastro de Eventos ContactId"
'         Report.Database.SetDataSource adors
'
'      Case g_iRptEvent:
'         If intWho = 0 Then
'            Set adors = oCnn.ExecCmdGetRs("Select * From Event Where " & strDate)
'            Set ADOrs1 = oCnn.ExecSpGetRs("Select * From Cadastro")
'         Else
'            Set adors = oCnn.ExecCmdGetRs("Select * From Event Where (((Event.Codigo) = " & intWho & ") AND " & strDate & ")")
''            Set ADOrs1 = oCnn.ExecSpGetRs("Select * From Cadastro Where (Codigo = " & intWho & ");")
'         End If
'         Set Report = New rptEvent
'         Report.ReportTitle = "Relatório de Ocorrências"
'         Me.Caption = "Relatório de Ocorrências"
'         Set CrDatabase = Report.Database
'         Set CrDatabaseTables = CrDatabase.Tables
'         Set CrDatabaseTable = CrDatabaseTables.Item(1)
'         CrDatabaseTable.SetPrivateData 3, adors
'         Set CrDatabaseTable = CrDatabaseTables.item(2)
'         CrDatabaseTable.SetPrivateData 3, ADOrs1
'
'      Case g_iRptService:
'        Set adors = oCnn.ExecCmdGetRs("Select * From Service")
''         Set adors = oCnn.ExecCmdGetRs("Select * From Service Where " & strSelection)
'         Set ADOrs1 = oCnn.ExecSpGetRs("Select * From Workstation")
'         Set Report = New rptService
'         Report.ReportTitle = "Relatório de Serviços"
'         Me.Caption = "Relatório de Serviços"
'         Set CrDatabase = Report.Database
'         Set CrDatabaseTables = CrDatabase.Tables
'         Set CrDatabaseTable = CrDatabaseTables.Item(1)
'         CrDatabaseTable.SetPrivateData 3, adors
'         Set CrDatabaseTable = CrDatabaseTables.Item(2)
'         CrDatabaseTable.SetPrivateData 3, ADOrs1
'
'      Case g_iRptEstatistic:
'         Set adors = oCnn.ExecCmdGetRs("Select * From Event")
'         Set ADOrs1 = oCnn.ExecSpGetRs("Select * From Cadastro")
'         Set Report = New rptEstatistica
'         Report.ReportTitle = "Estatística de Ocorrências"
'         Me.Caption = "Estatística de Ocorrências"
'         Set CrDatabase = Report.Database
'         Set CrDatabaseTables = CrDatabase.Tables
'         Set CrDatabaseTable = CrDatabaseTables.Item(1)
'         CrDatabaseTable.SetPrivateData 3, adors
'         Set CrDatabaseTable = CrDatabaseTables.Item(2)
'         CrDatabaseTable.SetPrivateData 3, ADOrs1
'
'      Case g_iRptEstGeral:
'         Set adors = oCnn.ExecCmdGetRs("Select * From Event")
'         Set Report = New rptEstGeral
'         Report.ReportTitle = "Estatística de Ocorrências"
'         Me.Caption = "Estatística de Ocorrências"
'         Report.Database.SetDataSource adors
'
'      Case g_iRptEvolucao:
'         Set adors = oCnn.ExecSpGetRs("Select * From Cadastro")
'         Set Report = New rptEvolucao
'         Report.ReportTitle = "Relatório de Evolução de Clientes"
'         Me.Caption = "Relatório de Evolução de Clientes"
'         Report.Database.SetDataSource adors
'
'      Case g_iRptEmail:
'         Set adors = oCnn.ExecSpGetRs("Select * From Cadastro")
'         Set Report = New rptEmail
'         Report.ReportTitle = "Relatório de E-Mails"
'         Me.Caption = "Relatório de E-Mails"
'         Report.Database.SetDataSource adors
'
'      Case g_iRptContrato:
'         Set adors = oCnn.ExecSpGetRs("Select * From Cadastro")
'         Set Report = New rptContratos
'         Report.ReportTitle = "Controle de Contratos"
'         Me.Caption = "Controle de Contratos"
'         Report.Database.SetDataSource adors
'
'      Case g_iRptFeriados:
'         Set adors = oCnn.ExecSpGetRs("Select * From Feriado")
'         Set ADOrs1 = oCnn.ExecSpGetRs("Select * From Municipio")
'         Dim ADOrs3(0 To 1) As ADODB.Recordset
'         Set ADOrs3(0) = oCnn.ExecSpGetRs("Select * From Nacional")
'         Set ADOrs3(1) = oCnn.ExecSpGetRs("Select * From Estado")
'         Set Report = New rptFeriados
'         Set CrDatabase = Report.Database
'         Set CrDatabaseTables = CrDatabase.Tables
'         Set CrDatabaseTable = CrDatabaseTables.Item(1)
'         CrDatabaseTable.SetPrivateData 3, adors
'         Set CrDatabaseTable = CrDatabaseTables.Item(2)
'         CrDatabaseTable.SetPrivateData 3, ADOrs1
'         Set CrDatabaseTable = CrDatabaseTables.Item(3)
'         CrDatabaseTable.SetPrivateData 3, ADOrs3(0)
'         Set CrDatabaseTable = CrDatabaseTables.Item(4)
'         CrDatabaseTable.SetPrivateData 3, ADOrs3(1)
'         Report.ReportTitle = "Relatório de Feriados por Municipio"
'         Me.Caption = "Relatório de Feriados por Municipio"
'
'         Case g_iRptListaClientes:
'         Set adors = oCnn.ExecSpGetRs("Select * From Cadastro")
'         Set Report = New rptListaClientes
'         Report.ReportTitle = "Lista de Clientes"
'         Me.Caption = "Lista de Clientes"
'         Report.Database.SetDataSource adors
   End Select
   
   If strSelection <> "" Then
      Report.RecordSelectionFormula = strSelection
   End If
   
   CRViewer1.ReportSource = Report
   Report.PaperSize = GetSetting(USVersion, "Options", "PaperSize", 9)
   Report.RightMargin = Report.LeftMargin

   If intTipoAction = 0 Then
      CRViewer1.ViewReport
   ElseIf intTipoAction = 1 Then
      CRViewer1.PrintReport
   End If
      
   GetPrinterOptions
   Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Resize()
   CRViewer1.Top = 0
   CRViewer1.Left = 0
   CRViewer1.Height = ScaleHeight
   CRViewer1.Width = ScaleWidth
End Sub

Public Property Let SetTipo(tipo As Integer)
   intTipoRpt = tipo
End Property

Public Property Let SetAction(action As Integer)
   intTipoAction = action
End Property

Public Property Let SetSelection(formula As String)
   strSelection = formula
End Property

Public Property Let SetWho(fClient As Long)
   intWho = fClient
End Property

Public Property Let SetDate(fDate As String)
   strDate = fDate
End Property

' Changes the size of the paper for the report.
' Note that your printer may not accept all the available paper sizes.
Private Sub cboPaperSize_Click()
   Report.PaperSize = cboPaperSize.ItemData(cboPaperSize.ListIndex)    ' Set the papersize option
   SaveSetting USVersion, "Options", "PaperSize", Report.PaperSize
   If Me.Visible Then
      CRViewer1.Refresh
   End If
End Sub

' Changes the paper orientation for the displayed report.
Private Sub cboPaperOrientation_Click()
   Report.PaperOrientation = cboPaperOrientation.ItemData(cboPaperOrientation.ListIndex)
   SaveSetting USVersion, "Options", "PaperOrientation", Report.PaperOrientation
   If Me.Visible Then CRViewer1.Refresh
End Sub

' Changes the paper bin source for the displayed report.
' To enumerate the printer bins available for your printer, see EnumPrinterBins
' in Utilities.bas
' Note that your printer may override this setting to accommodate the papersize setting.
Private Sub cboPaperSource_Click()
   Report.PaperSource = cboPaperSource.ItemData(cboPaperSource.ListIndex)
   SaveSetting USVersion, "Options", "PaperSource", Report.PaperSource
End Sub

' Changes the printer duplex setting for the displayed report.
Private Sub cboPrinterDuplex_Click()
   Report.PrinterDuplex = cboPrinterDuplex.ItemData(cboPrinterDuplex.ListIndex)
   SaveSetting USVersion, "Options", "PrinterDuplex", Report.PrinterDuplex
End Sub

' Enumerate all the available printer options for the report
' Note that GetPaperSource, GetPaperSize, GetPrinterDuplex, GetPaperOrientation will not retrieve
' accurate settings unless the printer settings have been saved in the report or the properties have
' been set some place in code.  In this case, rptCustomer has printer settings saved in it.
Private Sub GetPrinterOptions()
   Dim i As Integer                            ' Counter
   Dim l  As Integer
   Dim PaperSource As Integer
   Dim PrinterDuplex As CRPrinterDuplexType
   Dim PaperOrientation As CRPaperOrientation
   Dim PaperSize As CRPaperSize
   
'   lblPrinterName.Caption = "Page Setup for: " & Report.PrinterName       ' Display the printer name

    'Display the selected priter
    If Printers.Count = 0 Then
        txtImpressora.Text = "Não existe Impressora instalada"
    Else
        txtImpressora.Text = Report.PrinterName
    End If

   'Display the list of available printer bins in the cboPaperSource combo box.
   EnumPrinterBins Report.PrinterName, cboPaperSource
   PaperSource = GetSetting(USVersion, "Options", "PaperSource", 0)         'Verifica se exista setup já armazenado no Reg
   If PaperSource = 0 Then
     SaveSetting USVersion, "Options", "PaperSource", 7                     'Salva setup no Reg
     PaperSource = 7                                                        'Get the report's paper source
   End If
   ' Cycle through the combo box and select the correct currently selected type of papersource in the report
   With cboPaperSource
      For i = 0 To .ListCount - 1
         If .ItemData(i) = PaperSource Then .ListIndex = i
      Next i
   End With
   
   ' Display the list of available printer duplexing types in the cboPrinterDuplex combo box.
   ' Addcbo is a helper function to make the code cleaner
   ' Addcbo format:   <combo name>, <item caption>, <.itemdata(.listindex) to assign>
   Addcbo cboPrinterDuplex, "Simplex", crPRDPSimplex
   Addcbo cboPrinterDuplex, "Horizontal", crPRDPHorizontal
   Addcbo cboPrinterDuplex, "Vertical", crPRDPVertical
   
   PrinterDuplex = GetSetting(USVersion, "Options", "PrinterDuplex", 0)     'Verifica se exista setup já armazenado no Reg
   If PrinterDuplex = 0 Then
     SaveSetting USVersion, "Options", "PrinterDuplex", 1                   'Salva setup no Reg
     PrinterDuplex = 1
   End If
   ' Cycle through the combo box and select the correct currently selected type of printer duplexing in the report
   With cboPrinterDuplex
      For i = 0 To .ListCount - 1
         If .ItemData(i) = PrinterDuplex Then .ListIndex = i
      Next i
   End With
   
   ' Display the list of available paper orientations in the cboPaperOrientation combo box.
   Addcbo cboPaperOrientation, "Retrato", crPortrait
   Addcbo cboPaperOrientation, "Paisagem", crLandscape
   
   PaperOrientation = GetSetting(USVersion, "Options", "PaperOrientation", 0)             'Verifica se exista setup já armazenado no Reg
   If PaperOrientation = 0 Then
     SaveSetting USVersion, "Options", "PaperOrientation", 1                              'Salva setup default
     PaperOrientation = 1
   End If
   
   ' Cycle through the combo box and select the correct currently selected type of paper orientation in the report
   With cboPaperOrientation
      For i = 0 To .ListCount - 1
         If .ItemData(i) = PaperOrientation Then .ListIndex = i
      Next i
   End With
       
'   Add the large number of supported paper sizes to the cboPaperSize combobox
'   Addcbo cboPaperSize, "Default", crDefaultPaperSize
   Addcbo cboPaperSize, "Letter 8 1/2 x 11 pol", crPaperLetter
'   Addcbo cboPaperSize, "Small Letter", crPaperLetterSmall
   Addcbo cboPaperSize, "Legal 8 1/2 x 14 pol", crPaperLegal
'   Addcbo cboPaperSize, "10x14", crPaper10x14
'   Addcbo cboPaperSize, "11x17", crPaper11x17
'   Addcbo cboPaperSize, "A3", crPaperA3
   Addcbo cboPaperSize, "A4 297 x 210 mm", crPaperA4
'   Addcbo cboPaperSize, "A4 Small", crPaperA4Small
'   Addcbo cboPaperSize, "A5", crPaperA5
'   Addcbo cboPaperSize, "B4", crPaperB4
'   Addcbo cboPaperSize, "B5", crPaperB5
'   Addcbo cboPaperSize, "C Sheet", crPaperCsheet
'   Addcbo cboPaperSize, "D Sheet", crPaperDsheet
'   Addcbo cboPaperSize, "Envelope C4", crPaperEnvelopeC4
'   Addcbo cboPaperSize, "Executive", crPaperExecutive
'   Addcbo cboPaperSize, "Fanfold Legal German", crPaperFanfoldLegalGerman
'   Addcbo cboPaperSize, "Fanfold Standard German", crPaperFanfoldStdGerman
'   Addcbo cboPaperSize, "Folio", crPaperFolio
'   Addcbo cboPaperSize, "Ledger", crPaperLedger
'   Addcbo cboPaperSize, "Note", crPaperNote
'   Addcbo cboPaperSize, "Quarto", crPaperQuarto

PaperSize = GetSetting(USVersion, "Options", "PaperSize", 0)          'Verifica se exista setup já armazenado no Reg
   If PaperSize = 0 Then
     SaveSetting USVersion, "Options", "PaperSize", 9                  'Salva setup no Reg
     PaperSize = 9
   End If
   
   ' Cycle through the combo box and select the correct currently selected type of paper size in the report
   With cboPaperSize
      For i = 0 To .ListCount - 1
         If .ItemData(i) = PaperSize Then .ListIndex = i
      Next i
   End With
End Sub

'A small helper function for the ShowPrinterOption functions that
'helps reduce the amount of code to write
'Addcbo format:   <combo name to add item to>, <item caption>, <.itemdata(.listindex) to assign>
Private Sub Addcbo(cbo As ComboBox, Name As String, Index As Integer)
   cbo.AddItem Name                        ' Add the name of the item to the combo box
   cbo.ItemData(cbo.NewIndex) = Index      ' Set the .itemdata(.listindex) for later retrieval
End Sub

' Call the Printer Setup dialog.  This dialog does not reflect
' changes that we may have made via the PaperSource, PrinterDuplex
' and PaperSize methods, since this method changes the **Printer Settings**,
' not the **Report Printer Settings**.  The two sets of methods are
' independent and are intended for use in different situations.
Private Sub pctPrinterSettings_Click()
   On Error Resume Next
   Report.PrinterSetup Me.hwnd
End Sub
