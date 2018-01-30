VERSION 5.00
Object = "{8767A745-088E-4CA6-8594-073D6D2DE57A}#9.2#0"; "crviewer9.dll"
Begin VB.Form frmViewReport9 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8055
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11775
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmViewReport9.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8055
   ScaleWidth      =   11775
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
      Left            =   4920
      Picture         =   "frmViewReport9.frx":0642
      ScaleHeight     =   330
      ScaleWidth      =   360
      TabIndex        =   7
      ToolTipText     =   "Printer Setup"
      Top             =   240
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   250
      Left            =   6360
      TabIndex        =   6
      Text            =   "Imprimindo em:"
      Top             =   30
      Width           =   1080
   End
   Begin VB.TextBox txtImpressora 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   250
      Left            =   7560
      TabIndex        =   5
      Text            =   "Impressora Selecionada"
      Top             =   30
      Width           =   3975
   End
   Begin VB.ComboBox cboPrinterDuplex 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   9360
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   720
      Visible         =   0   'False
      Width           =   2115
   End
   Begin VB.ComboBox cboPaperSource 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   7200
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   720
      Visible         =   0   'False
      Width           =   2115
   End
   Begin VB.ComboBox cboPaperOrientation 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   10080
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   280
      Width           =   1410
   End
   Begin VB.ComboBox cboPaperSize 
      Appearance      =   0  'Flat
      Height          =   315
      ItemData        =   "frmViewReport9.frx":0DB4
      Left            =   7920
      List            =   "frmViewReport9.frx":0DB6
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   280
      Width           =   2010
   End
   Begin CRVIEWER9LibCtl.CRViewer9 CRViewer1 
      Height          =   8055
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11775
      lastProp        =   500
      _cx             =   20770
      _cy             =   14208
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
      EnableDrillDown =   -1  'True
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
      LaunchHTTPHyperlinksInNewBrowser=   -1  'True
   End
End
Attribute VB_Name = "frmViewReport9"
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
Private strDataPeriodoIni As String
Private strDataPeriodoFim As String
Public OpenLow As String
Public OpenHigh As String
Public CloseLow As String
Public CloseHigh As String
Public DataEvt As String
Public intervalo As String
Public Intervalo2 As String
Public Percurso As Integer
Public Ponto As Integer

Private Sub Form_Load()

   Dim success As Long

   Left = 50
   Top = 50
   
   Screen.MousePointer = vbHourglass
   DoEvents
   
   If oCnn Is Nothing Then
         Set oCnn = New clsConnection
   End If
   
   Select Case intTipoRpt
   
      Case g_iRptCLocais:
         success = SetWindowPos(frmCadastro.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, FLAGS)
         Hide

         Set adors = oCnn.ExecSpGetRs("SELECT * FROM Entity")
         Set ADOrs1 = oCnn.ExecSpGetRs("SELECT * From Floor")
         Set Report = New RptCLocais
         Report.ReportTitle = "Relatório de Locais/Lojas Cadastradas"
         Me.Caption = "Relatório de Locais/Lojas Cadastradas"
         Set CrDatabase = Report.Database
         Set CrDatabaseTables = CrDatabase.Tables
         Set CrDatabaseTable = CrDatabaseTables.Item(1)
         CrDatabaseTable.SetPrivateData 3, adors
         Set CrDatabaseTable = CrDatabaseTables.Item(2)
         CrDatabaseTable.SetPrivateData 3, ADOrs1
                            
      Case g_iRptLAbertos:
         success = SetWindowPos(frmReport.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, FLAGS)
         Hide

         Set adors = oCnn.ExecSpGetRs("SELECT * FROM C_LojaAberta")
         Set Report = New RptLocaisAbertos
         Report.ReportTitle = "Relatório de Lojas Abertas"
         Me.Caption = "Relatório de Lojas Abertas"
         Report.Database.SetDataSource adors
          
      Case g_iRptLFechados:
         success = SetWindowPos(frmReport.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, FLAGS)
         Hide

         Set adors = oCnn.ExecSpGetRs("SELECT * FROM C_LojaFechada")
         Set Report = New RptLocaisFechados
         Report.ReportTitle = "Relatório de Lojas Fechadas"
         Me.Caption = "Relatório de Lojas Fechadas"
         Report.ParameterFields(1).AddCurrentValue gstCompany
         Report.Database.SetDataSource adors
      
      Case g_iRptEventos:
         success = SetWindowPos(frmReport.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, FLAGS)
         Hide

         Set adors = oCnn.ExecSpGetRs("SELECT * FROM c_Event")
         'Set adors = oCnn.ExecSpGetRs("SELECT Event.*, Sensor.Tipo_Sensor FROM Event " & _
         '                             "INNER JOIN Sensor ON Event.fk_Sensor = Sensor.Serial_Number")
         Set Report = New RptEventos
         Report.ReportTitle = "Relatório de Eventos"
         Me.Caption = "Relatório de Eventos"
         Report.Database.SetDataSource adors
          
      Case g_iRptUEventos:
         success = SetWindowPos(frmReport.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, FLAGS)
         Hide

         Set adors = oCnn.ExecSpGetRs("SELECT * FROM C_Event")
         Set Report = New RptUEventos
         Report.ReportTitle = "Relatório: Últimos Eventos"
         Me.Caption = "Relatório: Últimos Eventos"
         Report.ParameterFields(1).AddCurrentValue gstCompany
         Report.Database.SetDataSource adors
          
      Case g_iRptSCZonas:
         success = SetWindowPos(frmReport.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, FLAGS)
         Hide

         Set adors = oCnn.ExecSpGetRs("SELECT * FROM C_StatusCorrente")
         Set Report = New RptSituacaoZonas
         Report.ReportTitle = "Relatório: Situação Corrente dos Sensores/Receptores"
         Me.Caption = "Relatório: Situação corrente das Zonas"
         Report.Database.SetDataSource adors
                        
      Case g_iRptSCLocais:
         success = SetWindowPos(frmReport.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, FLAGS)
         Hide

         Set adors = oCnn.ExecSpGetRs("SELECT * FROM C_StatusCorrente")
         Set Report = New RptSituacaoZonas
         Report.ReportTitle = "Relatório: Situação corrente das Lojas"
         Me.Caption = "Relatório: Situação corrente das Lojas"
         Report.Database.SetDataSource adors
              
      Case g_iRptEventosUnico:
         success = SetWindowPos(frmReport.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, FLAGS)
         Hide

         Set adors = oCnn.ExecSpGetRs("SELECT * FROM Event WHERE fk_Entity = " & _
                     frmReport.lstLocal.ItemData(frmReport.lstLocal.ListIndex))
         Set Report = New RptEventosUnico
         Report.ReportTitle = "Relatório de Eventos"
         Me.Caption = "Relatório de Eventos"
         Report.Database.SetDataSource adors

      Case g_iRptCZonas:
         success = SetWindowPos(frmCZonas.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, FLAGS)
         Hide

         Set adors = oCnn.ExecSpGetRs("SELECT * FROM Entity")
         Set ADOrs1 = oCnn.ExecSpGetRs("SELECT * From Floor")
         Set ADOrs2 = oCnn.ExecSpGetRs("SELECT * From Sensor")
         Set Report = New RptCZonas
         Report.ReportTitle = "Relatório de Sensores/Receptores Cadastrados"
         Me.Caption = "Relatório de Sensores/Receptores Cadastrados"
         Set CrDatabase = Report.Database
         Set CrDatabaseTables = CrDatabase.Tables
         Set CrDatabaseTable = CrDatabaseTables.Item(1)
         CrDatabaseTable.SetPrivateData 3, adors
         Set CrDatabaseTable = CrDatabaseTables.Item(2)
         CrDatabaseTable.SetPrivateData 3, ADOrs1
         Set CrDatabaseTable = CrDatabaseTables.Item(3)
         CrDatabaseTable.SetPrivateData 3, ADOrs2
                      
      Case g_iRptCRonda:
         success = SetWindowPos(frmCRonda.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, FLAGS)
         Hide

         Set adors = oCnn.ExecSpGetRs("SELECT * FROM Percurso")
         Set ADOrs1 = oCnn.ExecSpGetRs("SELECT * From Ronda")
         Set ADOrs2 = oCnn.ExecSpGetRs("SELECT * From RdHorario")
         Set ADOrs3 = oCnn.ExecSpGetRs("SELECT * From Entity")
         Set Report = New RptCRonda
         Report.ReportTitle = "Relatório de Rondas Cadastradas"
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
          
      Case g_iRptOperacao:
         success = SetWindowPos(frmReport.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, FLAGS)
         Hide
         
         Set adors = oCnn.ExecSpGetRs("SELECT * FROM Service")
         Set Report = New RptOperacao
         Report.ReportTitle = "Log de Operações"
         Me.Caption = "Log de Operações"
         Report.Database.SetDataSource adors
          
      Case g_iRptAFUnico:
         success = SetWindowPos(frmReport.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, FLAGS)
         Hide

         Set adors = oCnn.ExecSpGetRs("SELECT * FROM C_AccessClose")
         Set ADOrs1 = oCnn.ExecSpGetRs("SELECT * FROM C_AccessOpen")
         Set Report = New RptAFUnico
         Me.Caption = "Relatório de Abertura e Fechamento"
         Report.ReportTitle = "Relatório de Abertura e Fechamento"
         Report.ParameterFields(1).AddCurrentValue OpenLow
         Report.ParameterFields(2).AddCurrentValue OpenHigh
         Report.ParameterFields(3).AddCurrentValue CloseLow
         Report.ParameterFields(4).AddCurrentValue CloseHigh
         Set CrDatabase = Report.Database
         Set CrDatabaseTables = CrDatabase.Tables
         Set CrDatabaseTable = CrDatabaseTables.Item(1)
         CrDatabaseTable.SetPrivateData 3, adors
         Set CrDatabaseTable = CrDatabaseTables.Item(2)
         CrDatabaseTable.SetPrivateData 3, ADOrs1
          
      Case g_iRptAFTodos:
         success = SetWindowPos(frmReport.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, FLAGS)
         Hide

         Set adors = oCnn.ExecSpGetRs("SELECT * FROM C_AccessClose")
         Set ADOrs1 = oCnn.ExecSpGetRs("SELECT * FROM C_AccessOpen")
         Set Report = New RptAFTodos
         Me.Caption = "Relatório de Abertura e Fechamento"
         Report.ReportTitle = "Relatório de Abertura e Fechamento"
         Report.ParameterFields(1).AddCurrentValue OpenLow
         Report.ParameterFields(2).AddCurrentValue OpenHigh
         Report.ParameterFields(3).AddCurrentValue CloseLow
         Report.ParameterFields(4).AddCurrentValue CloseHigh
         Set CrDatabase = Report.Database
         Set CrDatabaseTables = CrDatabase.Tables
         Set CrDatabaseTable = CrDatabaseTables.Item(1)
         CrDatabaseTable.SetPrivateData 3, adors
         Set CrDatabaseTable = CrDatabaseTables.Item(2)
         CrDatabaseTable.SetPrivateData 3, ADOrs1
          
      Case g_iRptEvRonda:
         success = SetWindowPos(frmRonda.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, FLAGS)
         Hide

         Set adors = oCnn.ExecSpGetRs("SELECT * FROM Percurso")
         Set ADOrs1 = oCnn.ExecSpGetRs("SELECT * FROM Ronda")
         Set ADOrs2 = oCnn.ExecSpGetRs("SELECT * FROM EvtRonda")
         Set Report = New RptEvRonda
         Me.Caption = "Relatório de Eventos de Ronda"
         Report.ReportTitle = "Relatório de Eventos de Ronda"
         Report.ParameterFields(1).AddCurrentValue DataEvt
         Report.ParameterFields(2).AddCurrentValue intervalo
         Report.ParameterFields(3).AddCurrentValue Intervalo2
         Report.ParameterFields(4).AddCurrentValue Percurso
         Report.ParameterFields(5).AddCurrentValue Ponto
         Set CrDatabase = Report.Database
         Set CrDatabaseTables = CrDatabase.Tables
         Set CrDatabaseTable = CrDatabaseTables.Item(1)
         CrDatabaseTable.SetPrivateData 3, adors
         Set CrDatabaseTable = CrDatabaseTables.Item(2)
         CrDatabaseTable.SetPrivateData 3, ADOrs1
         Set CrDatabaseTable = CrDatabaseTables.Item(3)
         CrDatabaseTable.SetPrivateData 3, ADOrs2
          
       Case g_iRptExRonda:
         success = SetWindowPos(frmRonda.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, FLAGS)
         Hide

         Set adors = oCnn.ExecSpGetRs("SELECT * FROM Percurso")
         Set ADOrs1 = oCnn.ExecSpGetRs("SELECT * FROM Ronda")
         Set ADOrs2 = oCnn.ExecSpGetRs("SELECT * FROM EvtRonda")
         Set Report = New RptEvRonda
         Me.Caption = "Relatório de Exceções de Ronda"
         Report.ReportTitle = "Relatório de Exceções de Ronda"
         Report.ParameterFields(1).AddCurrentValue DataEvt
         Report.ParameterFields(2).AddCurrentValue intervalo
         Report.ParameterFields(3).AddCurrentValue Intervalo2
         Report.ParameterFields(4).AddCurrentValue Percurso
         Report.ParameterFields(5).AddCurrentValue Ponto
         Set CrDatabase = Report.Database
         Set CrDatabaseTables = CrDatabase.Tables
         Set CrDatabaseTable = CrDatabaseTables.Item(1)
         CrDatabaseTable.SetPrivateData 3, adors
         Set CrDatabaseTable = CrDatabaseTables.Item(2)
         CrDatabaseTable.SetPrivateData 3, ADOrs1
         Set CrDatabaseTable = CrDatabaseTables.Item(3)
         CrDatabaseTable.SetPrivateData 3, ADOrs2
      
       Case g_iRptZInativas:
         success = SetWindowPos(frmInativos.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, FLAGS)
         Hide

         Set adors = oCnn.ExecSpGetRs("SELECT * FROM Entity")
         Set ADOrs1 = oCnn.ExecSpGetRs("SELECT * From Floor")
         Set ADOrs2 = oCnn.ExecSpGetRs("SELECT * From Sensor")
         Set Report = New RptZonasInativas
         Report.ReportTitle = "Relatório de Zonas Inativas"
         Me.Caption = "Relatório de Zonas Inativas"
         Set CrDatabase = Report.Database
         Set CrDatabaseTables = CrDatabase.Tables
         Set CrDatabaseTable = CrDatabaseTables.Item(1)
         CrDatabaseTable.SetPrivateData 3, adors
         Set CrDatabaseTable = CrDatabaseTables.Item(2)
         CrDatabaseTable.SetPrivateData 3, ADOrs1
         Set CrDatabaseTable = CrDatabaseTables.Item(3)
         CrDatabaseTable.SetPrivateData 3, ADOrs2
         
       Case g_iRptCritico:
         success = SetWindowPos(frmReport.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, FLAGS)
         Hide
         Set adors = oCnn.ExecSpGetRs("SELECT * FROM C_Critico ORDER BY cp_Critico DESC;")
         Set Report = New RptCritico
         Report.ReportTitle = "Relatório de Eventos Criticos"
         Me.Caption = "Relatório de Eventos Criticos"
         Report.Database.SetDataSource adors
                  
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

Public Property Get SetSelection() As String
   SetSelection = strSelection
End Property

Public Property Let SetWho(fClient As Long)
   intWho = fClient
End Property

Public Property Let SetDate(fDate As String)
   strDate = fDate
End Property

Public Property Let DataPeriodoIni(fDate As String)
   strDataPeriodoIni = fDate
End Property

Public Property Let DataPeriodoFim(fDate As String)
   strDataPeriodoFim = fDate
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
   
'   Display the selected priter
If Printers.Count = 0 Then
     txtImpressora.Text = "Não existe Impressora instalada"
Else
     txtImpressora.Text = Report.PrinterName
End If

   ' Display the list of available printer bins in the cboPaperSource combo box.
   EnumPrinterBins Report.PrinterName, cboPaperSource
   PaperSource = GetSetting(USVersion, "Options", "PaperSource", 0)         'Verifica se exista setup já armazenado no Reg
   If PaperSource = 0 Then
     PaperSource = Report.PaperSource    ' Get the report's paper source
     SaveSetting USVersion, "Options", "PaperSource", 7                     'Salva setup no Reg
   End If
   With cboPaperSource
      For i = 0 To .ListCount - 1                   ' Cycle through the combo box and select the correct currently selected type of papersource in the report
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
   If PrinterDuplex = 0 Then                                        ' Get the report's printer duplex setting
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
   Addcbo cboPaperSize, "Envelope DL", crPaperEnvelopeDL
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
Private Sub Addcbo(cbo As ComboBox, name As String, Index As Integer)
   cbo.AddItem name                        ' Add the name of the item to the combo box
   cbo.ItemData(cbo.NewIndex) = Index      ' Set the .itemdata(.listindex) for later retrieval
End Sub

' Call the Printer Setup dialog.  This dialog does not reflect
' changes that we may have made via the PaperSource, PrinterDuplex
' and PaperSize methods, since this method changes the **Printer Settings**,
' not the **Report Printer Settings**.  The two sets of methods are
' independent and are intended for use in different situations.
Private Sub pctPrinterSettings_Click()
   On Error Resume Next
   Report.PrinterSetup Me.hWnd
End Sub


