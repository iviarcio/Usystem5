Attribute VB_Name = "Utilities"
Option Explicit

' API functions and constants used in EnumPrinterBins
Private Declare Function OpenPrinter Lib "winspool.drv" Alias _
   "OpenPrinterA" (ByVal pPrinterName As String, phPrinter As Long, _
   ByVal pDefault As Long) As Long
Private Declare Function ClosePrinter Lib "winspool.drv" ( _
   ByVal hPrinter As Long) As Long
Private Declare Function DeviceCapabilities Lib "winspool.drv" _
   Alias "DeviceCapabilitiesA" (ByVal lpDeviceName As String, _
   ByVal lpPort As String, ByVal iIndex As Long, lpOutput As Any, _
   ByVal dev As Long) As Long

Private Const DC_BINS = 6
Private Const DC_BINNAMES = 12

'Constantes que armazenan o número referente ao tipo do relatório
'utilizadas pelo CRViewer do Crystal
Public Const g_iRptAFUnico As Integer = 1       'Relatório de Abertura e Fechamento Único       OK
Public Const g_iRptAFTodos As Integer = 2       'Relatório de Abertura e Fechamento Todos       OK
Public Const g_iRptCLocais As Integer = 3       'Relatório de Cadastro de Locais                OK
Public Const g_iRptCZonas As Integer = 4        'Relatório de Cadastro de Zonas                 OK
Public Const g_iRptEventos As Integer = 5       'Relatório de Eventos                           OK
Public Const g_iRptEventosUnico As Integer = 6  'Relatório Eventos Únicos                       OK
Public Const g_iRptLAbertos As Integer = 7      'Relatório de Locais Abertos                    Ok
Public Const g_iRptLFechados As Integer = 8     'Relatório de Locais Fechados                   Ok
Public Const g_iRptOperacao As Integer = 9      'Relatório de Operações                         OK
Public Const g_iRptSCLocais As Integer = 10     'Relatório de Situação Corrente Locais          Ok
Public Const g_iRptSCZonas As Integer = 11      'Relatório de Situação Corrente Zonas           Ok
Public Const g_iRptUEventos As Integer = 12     'Relatório de Últimos Eventos                   Ok
Public Const g_iRptZInativas As Integer = 13    'Relatório de Zonas Inativas                    Ok
Public Const g_iRptCRonda As Integer = 14       'Relatório de Config Ronda                      OK
Public Const g_iRptEvRonda As Integer = 15      'Relatório de Eventos Ronda                     Ok
Public Const g_iRptExRonda As Integer = 16      'Relatório de Exceções Eventos                  Ok
Public Const g_iRptCritico As Integer = 17      'Relatório de Eventos Críticos                  Ok

'Variaveis usadas para a visualização dos relatórios
Public adors As New ADODB.Recordset
Public ADOrs1 As ADODB.Recordset
Public ADOrs2 As ADODB.Recordset
Public ADOrs3 As ADODB.Recordset

Public Report As Object
Public CrDatabase As CRAXDRT.Database
Public CrDatabaseTables As CRAXDRT.DatabaseTables
Public CrDatabaseTable As CRAXDRT.DatabaseTable
Public CrSections As CRAXDRT.Sections
Public CrSection As CRAXDRT.Section
Public CrReportObj As CRAXDRT.ReportObjects
Public CrSubreportObj As CRAXDRT.SubreportObject
Public CrSubreport As CRAXDRT.Report

Public m_ShowObs As Boolean
Public Account_Bd As Boolean


' Add a list of the available paper sources for <PrinterName> to
' the combobox <cbo>
Public Sub EnumPrinterBins(PrinterName As String, cbo As ComboBox)
    Dim prn As Printer
    Dim hPrinter As Long                ' Handle of the selected printer
    Dim dwbins As Long                  ' The number of paperbins in the printer
    Dim i As Long                       ' counter
    Dim nameslist As String             ' The string returned with all the bin names
    Dim NameBin As String               ' The parsed bin name
    Dim numBin() As Integer             ' Used as part of the DeviceCapabilities API call
     
    cbo.Clear
    For Each prn In Printers
        ' Look through all the currently installed printers
        If prn.DeviceName = PrinterName Then
            ' We've found our printer -- open a handle to it
            If OpenPrinter(prn.DeviceName, hPrinter, 0) <> 0 Then
                ' Get the available bin numbers
                dwbins = DeviceCapabilities(prn.DeviceName, prn.Port, _
                                        DC_BINS, ByVal vbNullString, 0)
                ReDim numBin(1 To dwbins)
                nameslist = String(24 * dwbins, 0)
                dwbins = DeviceCapabilities(prn.DeviceName, prn.Port, _
                                        DC_BINS, numBin(1), 0)
                dwbins = DeviceCapabilities(prn.DeviceName, prn.Port, _
                                        DC_BINNAMES, ByVal nameslist, 0)
                For i = 1 To dwbins
                    ' For each bin number, add its corresponding name
                    ' to our combo box
                    NameBin = Mid(nameslist, 24 * (i - 1) + 1, 24)
                    NameBin = Left(NameBin, InStr(1, NameBin, Chr(0)) - 1)
                    cbo.AddItem NameBin
                    cbo.ItemData(cbo.NewIndex) = numBin(i)
                Next i
                ' Close the printer
                Call ClosePrinter(hPrinter)
            Else
                ' OpenPrinter failed, so we can't retrieve information about it
                cbo.AddItem prn.DeviceName & "  <Unavailable>"
            End If
        End If
    Next prn
End Sub

'A short procedure to give forms a nice teal shading.
Public Sub Dither(frm As Form)
   Dim intLoop As Integer
   ' Set the pen parameters
   frm.DrawStyle = vbInsideSolid
   frm.DrawMode = vbCopyPen
   frm.ScaleMode = vbPixels
   frm.DrawWidth = 8
   frm.ScaleWidth = 256
   For intLoop = 0 To 255
      frm.Line (intLoop, 0)-(intLoop - 1, Screen.Height), RGB(0, intLoop, intLoop), B
   Next intLoop
End Sub

