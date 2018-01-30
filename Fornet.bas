Attribute VB_Name = "ForBas"
Option Explicit

'Programa desenvolvido por:
'   José D. Favoretto Jr.
'   Marcio Machado pereira
'Copyright © 1995-2016 FOR Segurança Eletrônica

'Versão corrente do Usystem
Public Const curVersion = "5.0"
Public Const USVersion = "USystem " & curVersion

'Parametro BYPASS = 0, segurança OK, BYPASS = 1, segurança desativada
Public Const BYPASS = 1

'Variáveis de Registro e controle da Segurança
Public gstCondorID As String
Public gstChecksum As String
Public gstCompany As String
Public Const codUser = "20W44YR00"
Public Chave As clsproteq
Public Key_check As Integer     'Key_ckeck = 1 registro OK, Key_check = 0 sem registro

'Constantes e Tipos utilizados para criar as regiões poligonais
Public Const MAX_POINTS = 24
Public Type POINTAPI
    x As Long
    Y As Long
End Type
Public Type MATRIZ
    varPoints(1 To MAX_POINTS) As POINTAPI
End Type

'Tempo (minutos) da varredura da verificação da segurança
Public Const TempoCheckSecurity = 5

'Database Name
Public Const USystemDB = "USystemDB5.mdb"

'New Line, created in Main subroutine
Public NL As String * 2

'Connection object, created in Main subroutine
Public oCnn As clsConnection

'Constantes utilizadas pelas API's do Windows declaradas abaixo
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2
Public Const WS_MINIMIZEBOX = &H20000
Public Const WS_MAXIMIZEBOX = &H10000
Public Const WS_SYSMENU = &H80000
Public Const GWL_STYLE = (-16)
Public Const SWP_HIDEWINDOW = &H80
Public Const SWP_SHOWWINDOW = &H40
Public Const MF_BYCOMMAND = &H0&
Public Const MF_BYPOSITION = &H400&
Public Const SC_ARRANGE = &HF110
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOMOVE = &H2
Public Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE
Public Const SWP_FRAMECHANGED = &H20
Public Const SWP_NOZORDER = &H4
Public Const SC_RESTORE = &HF120
Public Const SC_SIZE = &HF000
Public Const SC_MOVE = &HF010
Public Const SC_MAXIMIZE = &HF030
Public Const SC_MINIMIZE = &HF020
Public Const SC_CLOSE = &HF060
Public Const SC_NEXTWINDOW = &HF040
Public Const SC_PREVWINDOW = &HF050

'API's do Windows utilizadas no USystem
Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal Y As Long) As Long
Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Declare Function CreatePolygonRgn Lib "gdi32" (lpPoint As MATRIZ, ByVal nCount As Long, ByVal nPolyFillMode As Long) As Long
Declare Function CreatePatternBrush Lib "gdi32" (ByVal hBitmap As Long) As Long
Declare Function PtInRegion Lib "gdi32" (ByVal hRgn As Long, ByVal x As Long, ByVal Y As Long) As Long
Declare Function InvertRgn Lib "gdi32" (ByVal hdc As Long, ByVal hRgn As Long) As Long
Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hdc As Long) As Long
Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Declare Function FillRgn Lib "gdi32" (ByVal hdc As Long, ByVal hRgn As Long, ByVal hBrush As Long) As Long
Declare Function FrameRgn Lib "gdi32" (ByVal hdc As Long, ByVal hRgn As Long, ByVal hBrush As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Declare Function PaintRgn Lib "gdi32" (ByVal hdc As Long, ByVal hRgn As Long) As Long
Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Declare Function GetSystemMenu Lib "user32.dll" (ByVal hWnd As Long, ByVal bRevert As Long) As Long
Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hWnd As Long, ByVal crKey As Long, ByVal bDefaut As Byte, ByVal dwFlags As Long) As Long
Declare Function DeleteMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long

'Sleep function
Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)

'Constantes utilizadas pelas regiões que são uma representação
'gráfica das Entidades do USystem
Public Enum BackColor
   colGray = 0
   colBlue = 1
   colGreen = 2
   colYellow = 3
   colRed = 4
   colRedGreen = 5
   colRedYellow = 6
   colGreenRed = 7
End Enum

'Cores utilizadas nos labels que representam um piso ou andar
Public Enum lblColor
   cRed = &HFF&
   cYellow = &HFFFF&
   cBlack = &H404040
   cBlue = &HC0C000
End Enum

' new to 1.0.27
'Cores usadas no eventos críticos
Public cColor(0 To 4) As ColorConstants
' end new

'Tipo que controla o modo de desenho da região
Public Enum ControlState
    StateNothing = 0
    StateDragging = 1
    StateSizing = 2
End Enum
   
'Tipos de Sensores do USystem
Public Const nTSensor = 8
Public Enum typeSensor
   s_Incendio = 0
   s_Intrusao = 1
   s_Emergencia = 2
   s_Panico = 3
   s_Sistema = 4
   s_Ronda = 5
   s_All = 6
   d_Repeater = 7
   d_Receiver = 8
End Enum

'Estado dos Sensores
Public Enum typeZona
   stFechado = 0
   stAberto = 1
   stCurto = 2
   stFalha = 3
   stTamper = 4
   stNone = 5
End Enum

Public Enum typeDirection
   ImgToEntity = 0
   EntityToImg = 1
   ImgToModule = 2
   ModuleToImg = 3
   PctToEntity = 4
   EntityToPct = 5
End Enum

Public Enum typeEvento
   stDefault = 0
   stAdded = 1
   stModified = 2
   stRemoved = 3
End Enum

Public Enum typeAccess
   sxOperator = 0
   sxAdministrator = 1
   sxSystem = 2
End Enum

'Tipo e variável utilizada para a desativação de Zonas Inicialmente ativadas
Public Enum typeQuestion
   sxQNone = 0
   sxQYes = 1
   sxQAll = 2
   sxQNo = 3
   sxQCancel = 4
End Enum
Public qResponse As typeQuestion

'Constantes usadas em Regiões
Public Const ALTERNATE = 1
Public Const WINDING = 2

'Strings que denotam o tipo de acesso ao USystem
Public strAccess(0 To 2) As String

'Strings que denotam os tipos de sensores e status dos mesmos
Public strTipo(0 To 8) As String
Public strStatus(0 To 5) As String
Public strModo(0 To 3) As String
Public strInfo(0 To 1) As String

'Constantes utilizadas nas caixas de mensagens
Public Const sxQuestion = vbYesNo + vbQuestion + vbDefaultButton2
Public Const sxExclamation = vbOKOnly + vbExclamation
Public Const sxInformation = vbOKOnly + vbInformation
Public Const sxCritical = vbOKOnly + vbCritical
Public Const sxInitialize = "Erro no arquivo de inicialização."
Public Const sxDatabase = "Erro na base de dados da rede USystem."
Public Const sxContact = "Contate o seu suporte técnico mais próximo!"
Public Const sxProname = "USystem"
Public Const sxAuthor = "FOR"
Public Const sxRegistro = "REGISTRO"
Public Const CONTATO = "Contate o seu suporte técnico!"

'Indices para as figuras mostradas na barra de status do objeto Imagelist1 (ForNet)
Public Const sxImgInform = 14
Public Const sxImgQuestion = 15
Public Const sxImgAlert = 16
Public Const sxImgOccur = 17
Public Const sxImgNone = 0
Public Const sxImgEntity = 7

Public Const bp As String = "','"     'usado na concatenação de Strings ambos os lados
Public Const rp As String = ",'"      'string somente do lado direito
Public Const lp As String = "',"      'string somente do lado esquerdo
Public Const np As String = ","       'somente numeros em ambos os lados

'Localização (caminho) do USystem5 e do Backup automático
Public m_sPath As String
Public m_sBPath As String

'Strings de Conexão e Conexões disponíveis
Public m_sDatabase As String
Public cnDB As ADODB.Connection
Public ErroNet As Boolean
Public ACCESSNET As Boolean
Public FATALACCESS As Boolean

'Objetos de Persistencia
Public rsBack As ADODB.Recordset
Public rsEvent As ADODB.Recordset
Public rsOccur As ADODB.Recordset

Public m_sTmpFileDB As String
Public m_sCommSett(0 To 3) As String
Public m_sCommPort(0 To 3) As String
Public m_sCommEnabled(0 To 3) As Boolean
Public m_sEvSound As String
Public m_sExSound As String
Public m_sErSound As String
Public m_sBgSound As String
Public m_sMsSound As String
Public m_sPnSound As String

'Controle de acesso às funções do USystem
Public m_tAccess As typeAccess
Public m_bChange As Boolean
Public m_bPermition As Boolean
Public m_bShutDown As Boolean
Public m_bUserUnload As Boolean
Public m_bDesignMode As Boolean
Public m_DragState As ControlState

Public m_bBackupAuto As Boolean
Public m_sHorario As String
Public m_dTOpen(vbSunday To vbSaturday) As Date
Public m_dTClose(vbSunday To vbSaturday) As Date
Public m_iEvKeep As Integer

Public m_Debug As Boolean

'Usuário corrente
Public m_sUser As String

'Lista de Entidades e entidade corrente
Public lstEntity As New Collection
Public tEntity As clsEntity

'Lista de Sensores (devices) e Sensor (módulo) corrente
Public lstModule As New Collection
Public tModule As clsModule

'Lista de Devices (módulos) não cadastrados
Public lstDevice As New Collection

'Lista de Eventos e Evento corrente
Public lstEvent As New Collection
Public tEvent As clsEvent

'Lista de Serviços de dupla verificação
Public lstService As New Collection

'Lista de Eventos de Pânico para temporização de fechamento automático
Public lstPanico As New Collection

'Fila de mensagens correntes First-in/First-out, sem duplicação, ainda não tratadas
Public tQueue As New clsQueue

'Tabela PTI (Product Type Identification)
Public lstPTI As New Collection

'Lista de Pisos e Piso corrente
Public lstPiso As New Collection
Public tpiso As clsPiso
Public m_iLastTop As Integer
Public m_iCurPiso As Long
Public m_bPisoLeft As Boolean

'Lista de Percursos de Ronda
Public lstPercurso As New Collection

'Display de Eventos, Display corrente e Lock do Display
Public lstDisplay As New Collection
Public tDisplay As clsDisplay
Public nDisplay As Integer
Public m_UpdateLock As Boolean

'Handle dos padrões de preenchimento para as regiões
Public lngFill(0 To 7) As Long

'Array de ctes usadas na Configuração da Serial
Public m_sBaud(0 To 5)  As String
Public m_sParity(0 To 4)  As String * 1
Public m_sData(0 To 4)  As String * 1
Public m_sStop(0 To 2)  As String
Public Const BufferSize = 1024

'Heart Beat
Public gHeart As String * 2
Public gHeartBeat As Integer

'T/R message Buffer
Public tBuffer$, rBuffer$

'Communication Control
Public m_bCommStatus As Boolean

'USystem Date/time Controls
Public ForTime As Date
Public ForData As Date

'Current handle of topmost form
Public curhwnd As Long

'Control the reload of a form
Public reloadForm As Boolean

'Current week day
Public curWeekday As Integer

'Controle do timer para leitura da esposta dos módulos
Public bFlag As Boolean

'Para uso da tela que mostra a comunicação da porta RS232
Public m_bShowComm As Boolean

Public Const H_Ack = "10"
Public Const H_Serial = "11"
Public Const H_Device = "13"
Public Const MID_Serial = "00"
Public Const MID_Repeater = "01"
Public Const MID_Sensor = "B2"

' new to 1.0.27
'Enfilera eventos críticos para serem tratados pelo operador
Public tFort As clsDigiFort
Public colBind As BindingCollection
' end new

Sub Main()
   'Testa se o programa já está sendo executado
   If App.PrevInstance Then
      MsgBox "A rede USystem já está sendo executada!" & Chr$(13) & Chr$(10) & _
             "Esta segunda ativação será terminada.", sxInformation, sxProname
      End
   End If
   
   NL = Chr$(13) & Chr$(10)
   
   'Inicia a tela de abertura
   m_UpdateLock = False
   nDisplay = 0
   frmSplash.Show
   frmSplash.lblMsg = "Iniciando o Sistema. Aguarde..."
   ProgressCounter = 0
   IncProgress
   frmSplash.ProgressBar1.Visible = True
   DoEvents
   frmSplash.MousePointer = vbHourglass
   Screen.MousePointer = vbHourglass
   DoEvents
   
   Dim rnapp As Integer
   rnapp = App.Revision
   Dim rnreg As String
   rnreg = GetSetting("USystem5", "Options", "Revision")
   If rnreg <> "" Then
      If CInt(rnreg) < rnapp Then
         DeleteSetting "USystem5", "Options"
         Register_Me
      End If
   Else
      Register_Me
   End If
   
   'Busca o path das Bases de Dados e dos Relatórios
   m_sPath = App.Path & "\Database"
   'Set the temporary name used in Backup & Restore
   m_sTmpFileDB = m_sPath & "\USystemDB5.tmp"
   'Get the path for automatic Backup
   m_sBPath = GetSetting("Usystem5", "Options", "Backup", App.Path & "\Backup")
   
   'Get the name os Databases
   m_sDatabase = GetSetting("USystem5", "Options", "DataBase")
   m_tAccess = sxOperator 'Default
   m_bChange = False
    
   'Obtem e verifica as informações de Licença
   gstCondorID = GetSetting("USystem5", "Options", "License")
   gstChecksum = GetSetting("USystem5", "Options", "Checksum")
   
   'Obtem e verifica as informações da Empresa
   gstCompany = GetSetting("USystem5", "Options", "Company")
   
   'Inicia a verificação da base de dados do USystem e remove Read Only
   If m_sDatabase <> "" Then
      Dim lattr As Integer
      lattr = GetAttr(m_sPath & "\" & m_sDatabase)
      If Err.Number <> 53 Then
         If lattr And vbReadOnly Then
            'A base de Dados está protegida contra escrita!
            SetAttr m_sPath & "\" & m_sDatabase, vbNormal
         End If
      Else
         Unload frmSplash
         Screen.MousePointer = vbDefault
         MsgBox sxDatabase + Chr$(13) + sxContact, sxCritical, sxProname
         End
      End If
         
      'm_sDatabase = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & m_sPath & "\" & m_sDatabase & ";Persist Security Info=False"
      m_sDatabase = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & m_sPath & "\" & m_sDatabase & ";Jet OLEDB:Database Password=DEPFwm89"
    
      'Estabelece a conexão com a base de dados
      Set cnDB = New ADODB.Connection
      cnDB.ConnectionString = m_sDatabase
      cnDB.ConnectionTimeout = 45
      cnDB.Open
      
      IncProgress
      
      'Verifica a consistência da base de dados (tipo e versão)
      Dim lrs As New ADODB.Recordset
      lrs.Open "SELECT * FROM Admin", cnDB, adOpenStatic, adLockReadOnly
      If Not lrs.EOF Then
         If lrs("Version") <> "DB" & curVersion Then
            Unload frmSplash
            Screen.MousePointer = vbDefault
            MsgBox sxDatabase + Chr$(13) + sxContact, sxCritical, sxProname
            End
         End If
      Else
         Unload frmSplash
         Screen.MousePointer = vbDefault
         MsgBox sxDatabase + Chr$(13) + sxContact, sxCritical, sxProname
         End
      End If
      lrs.Close
      
      IncProgress
            
      curWeekday = Weekday(Date)
      
      Sound_Init              'Inicializa constantes de som
      Ctes_Init               'Inicializa constantes do sistema (status, operadores, etc.)
      Aux_Initialize          'Busca configurações de backup
      Load_Entities           'Carrega as entidades
      Load_Percursos          'Carrega os percursos de ronda
      Load_PTI                'Carrega a Tabela PTI
      Event_Populate          'Carrega os últimos eventos ocorridos
      Data_CleanUp            'Horários de abertura
      'LastEvents_CleanUp 100  'Limpa a tabela últimos eventos exceto os ultimos 100 registros
      
      frmSplash.ProgressBar1.Visible = False
      DoEvents
      
      ForNet.Show             'Carrega a tela principal
      
      Sound_Update fmode:=sxBgSound, isCritico:=False, fNoSound:=False
      
   Else
      Unload frmSplash
      Screen.MousePointer = vbDefault
      MsgBox sxInitialize + Chr$(13) + sxContact, sxCritical, sxProname
      End
   End If
End Sub

Private Sub Register_Me()
   SaveSetting "USystem5", "Options", "Revision", CStr(App.Revision)
   SaveSetting "USystem5", "Options", "DataBase", "USystemDB5.mdb"
End Sub

Private Sub Ctes_Init()
   strAccess(0) = "Operador "
   strAccess(1) = "Administrador "
   strAccess(2) = "Sistema "

   strTipo(0) = "Incêndio"
   strTipo(1) = "Intrusão"
   strTipo(2) = "Emergência"
   strTipo(3) = "Pânico"
   strTipo(4) = "Sistema"
   strTipo(5) = "Ronda"
   strTipo(6) = " "
   strTipo(7) = "Repeater"
   strTipo(8) = "Receiver"
   
   strStatus(0) = "Ok "
   strStatus(1) = "Aberto "
   strStatus(2) = "Curto "
   strStatus(3) = " "
   strStatus(4) = "Vandalismo"
   strStatus(5) = " "
   
   strModo(0) = "Desabilitada"
   strModo(1) = "Desativada"
   strModo(2) = "Ativada"
   strModo(3) = "Programada"
   
   strInfo(0) = "Permanente"
   strInfo(1) = "Temporizado"
   
   m_sBaud(0) = "1200"
   m_sBaud(1) = "2400"
   m_sBaud(2) = "4800"
   m_sBaud(3) = "9600"
   m_sBaud(4) = "19200"
   m_sBaud(5) = "38400"

   m_sParity(0) = "E"  'Even
   m_sParity(1) = "O"  'Odd
   m_sParity(2) = "N"  'None
   m_sParity(3) = "M"  'Mark
   m_sParity(4) = "S"  'Space
   m_sData(0) = "4"
   m_sData(1) = "5"
   m_sData(2) = "6"
   m_sData(3) = "7"
   m_sData(4) = "8"
   m_sStop(0) = "1"
   m_sStop(1) = "1.5"
   m_sStop(2) = "2"
   
   cColor(0) = vbRed
   cColor(1) = &HFFFF80  'vbBlue
   cColor(2) = vbYellow
   cColor(3) = vbGreen
   cColor(4) = vbWhite

   curhwnd = 0
   qResponse = sxQNone
   
End Sub

'Rotina que carrega as entidades
Private Sub Load_Entities()

   Dim rsEntity As ADODB.Recordset
   Set rsEntity = New ADODB.Recordset
   rsEntity.CursorLocation = adUseClient
   rsEntity.CursorType = adOpenStatic
   rsEntity.LockType = adLockReadOnly
   rsEntity.Open "Select * From Entity Order By cp_Entity", cnDB
   
'  Design the entity collection
   While Not rsEntity.EOF
      Set tEntity = New clsEntity
      On Error Resume Next
      With tEntity
         .vId = rsEntity("cp_Entity")
         .vDescr = rsEntity("Descr_Entity")
         .vResp = rsEntity("Resp_Entity")
         .vTel1 = rsEntity("Tel1_Entity")
         .vTel2 = rsEntity("Tel2_Entity")
         .floor = rsEntity("fk_Floor")
         .message = rsEntity("message")
         .Vertices = rsEntity("nVertices")
         .Set_Coordinates 1, rsEntity("X01"), rsEntity("Y01")
         .Set_Coordinates 2, rsEntity("X02"), rsEntity("Y02")
         .Set_Coordinates 3, rsEntity("X03"), rsEntity("Y03")
         .Set_Coordinates 4, rsEntity("X04"), rsEntity("Y04")
         .Set_Coordinates 5, rsEntity("X05"), rsEntity("Y05")
         .Set_Coordinates 6, rsEntity("X06"), rsEntity("Y06")
         .Set_Coordinates 7, rsEntity("X07"), rsEntity("Y07")
         .Set_Coordinates 8, rsEntity("X08"), rsEntity("Y08")
         .Set_Coordinates 9, rsEntity("X09"), rsEntity("Y09")
         .Set_Coordinates 10, rsEntity("X10"), rsEntity("Y10")
         .Set_Coordinates 11, rsEntity("X11"), rsEntity("Y11")
         .Set_Coordinates 12, rsEntity("X12"), rsEntity("Y12")
         .Set_Coordinates 13, rsEntity("X13"), rsEntity("Y13")
         .Set_Coordinates 14, rsEntity("X14"), rsEntity("Y14")
         .Set_Coordinates 15, rsEntity("X15"), rsEntity("Y15")
         .Set_Coordinates 16, rsEntity("X16"), rsEntity("Y16")
         .Set_Coordinates 17, rsEntity("X17"), rsEntity("Y17")
         .Set_Coordinates 18, rsEntity("X18"), rsEntity("Y18")
         .Set_Coordinates 19, rsEntity("X19"), rsEntity("Y19")
         .Set_Coordinates 20, rsEntity("X20"), rsEntity("Y20")
         .Set_Coordinates 21, rsEntity("X21"), rsEntity("Y21")
         .Set_Coordinates 22, rsEntity("X22"), rsEntity("Y22")
         .Set_Coordinates 23, rsEntity("X23"), rsEntity("Y23")
         .Set_Coordinates 24, rsEntity("X24"), rsEntity("Y24")
         .hasAccessOpen = rsEntity("AccessOpen")
         .hasAccessClose = rsEntity("AccessClose")
         .OpenTime = rsEntity("OpenTime")
         .OpenLast = rsEntity("OpenLast")
         .CloseTime = rsEntity("CloseTime")
         .CloseLast = rsEntity("CloseLast")
         .flagInativo = False
      End With
      On Error GoTo 0
      lstEntity.Add Item:=tEntity, Key:=CStr(rsEntity("cp_Entity"))
      rsEntity.MoveNext
      IncProgress
   Wend
   rsEntity.Close
   On Error GoTo 0
   Entity_Populate
End Sub

Private Sub Entity_Populate()

   Dim tEntity As clsEntity
   Dim lrsModule As New ADODB.Recordset
   Dim lsql As String
   
   'When start USystem, all modules are considered actives
   Dim initAtividade As Date
   initAtividade = Now
   
   lrsModule.CursorLocation = adUseClient
   lrsModule.CursorType = adOpenStatic
   lrsModule.LockType = adLockReadOnly
   
   For Each tEntity In lstEntity
   
      lsql = "SELECT * FROM Sensor WHERE ([fk_Entity]= " & tEntity.vId & ");"
      lrsModule.Open lsql, cnDB
      lrsModule.ActiveConnection = Nothing
      
      While Not lrsModule.EOF
      
         Set tModule = New clsModule
         
         With tModule
            .Serial_Number = lrsModule("Serial_Number")
            .UID = lrsModule("UID")
            .mNumero = lrsModule("Numero_Sensor")
            .mEntity = lrsModule("fk_Entity")
            .mTipo = lrsModule("Tipo_Sensor")
            If IsNull(lrsModule("Local_Sensor")) Then
               .mLocal = ""
            Else
               .mLocal = lrsModule("Local_Sensor")
            End If
            .SInicial = lrsModule("Inicial_Sensor")
            .mCheck = lrsModule("Check_Sensor")
            .mJanela = lrsModule("Janela_Sensor")
            .mLogica = lrsModule("Tipo_Logica")
            .mNumLogica = lrsModule("Numero_Logica")
            If IsNull(lrsModule("Local_Logica")) Then
               .mLocalLogica = ""
            Else
               .mLocalLogica = lrsModule("Local_Logica")
            End If
            If IsNull(lrsModule("Arquivo")) Then
               .mArquivo = ""
            Else
               .mArquivo = lrsModule("Arquivo")
            End If
            If IsNull(lrsModule("PTI")) Then
               .mPTI = ""
            Else
               .mPTI = lrsModule("PTI")
            End If
            .mChkAtiv = lrsModule("chk_Atividade") And (.SInicial <> stDesabilitada)
            .mtempoAtiv = lrsModule("chk_Tempo")
            .mStatAtiv = True
            .mLastAtiv = initAtividade
            .critico = lrsModule("critico")
            .crColor = lrsModule("color")
            On Error Resume Next
               .ServerAddress = lrsModule("servidor")
               .Camera = lrsModule("camera")
               .Monitor = lrsModule("monitor")
               .telaCheia = lrsModule("telaCheia")
               .user = lrsModule("user_cftv")
               .senha = lrsModule("senha")
            On Error GoTo 0
            .popup = lrsModule("popup")
         End With
         
         tEntity.Add tModule, lrsModule("UID")
         lstModule.Add Item:=tModule, Key:=lrsModule("UID")
         tModule.InitZStatus = lrsModule("SZona")
         lrsModule.MoveNext
         
      Wend
      
      lrsModule.Close
      IncProgress
      
   Next
   
End Sub

'Rotina que carrega os percursos programados para Ronda
Private Sub Load_Percursos()
   Dim tPercurso As clsPercurso
   Dim rsPercurso As ADODB.Recordset
   Set rsPercurso = New ADODB.Recordset
   rsPercurso.CursorLocation = adUseClient
   rsPercurso.CursorType = adOpenStatic
   rsPercurso.LockType = adLockReadOnly
   rsPercurso.Open "Select * From Percurso", cnDB
   While Not rsPercurso.EOF
      Set tPercurso = New clsPercurso
      On Error Resume Next
      With tPercurso
         .idPercurso = rsPercurso("cpPercurso")
         .descrPercurso = rsPercurso("DescrPercurso")
         .Horario = rsPercurso("Horario")
         .desvio = rsPercurso("Desvio")
         .valSegSex = rsPercurso("ValSegSex")
         .valSab = rsPercurso("ValSab")
         .valDom = rsPercurso("ValDomFer")
         .status = rsPercurso("Status")
      End With
      On Error GoTo 0
      'Rotina que carrega os Rorários das Rondas
      tPercurso.Load_Horarios
      
      'Rotina que carrega os Atributos de cada Ponto de Ronda
      tPercurso.LoadRondas
      
      lstPercurso.Add Item:=tPercurso, Key:=CStr(rsPercurso("cpPercurso"))
      rsPercurso.MoveNext
      IncProgress
   Wend
   rsPercurso.Close
End Sub

'Rotina que carrega a tabela PTI (Product Type Identification)
Private Sub Load_PTI()
   Dim tPTI As clsPTI
   Dim rsPTI As ADODB.Recordset
   Set rsPTI = New ADODB.Recordset
   rsPTI.CursorLocation = adUseClient
   rsPTI.CursorType = adOpenStatic
   rsPTI.LockType = adLockReadOnly
   rsPTI.Open "Select * From Tabela_PTI", cnDB
   While Not rsPTI.EOF
      Set tPTI = New clsPTI
      On Error Resume Next
      With tPTI
         .sPTI = rsPTI("PTI")
         .sProduct = rsPTI("Produto")
         .sDescription = rsPTI("Descricao")
      End With
      On Error GoTo 0
      lstPTI.Add Item:=tPTI, Key:=rsPTI("PTI")
      rsPTI.MoveNext
      IncProgress
   Wend
   rsPTI.Close
End Sub

'Rotina que carrega os Últimos Eventos ocorridos
Private Sub Event_Populate()
   Dim idxTipo As Integer
   Dim rsTipo As String
   Dim rsEvent As ADODB.Recordset
   Set rsEvent = New ADODB.Recordset
   rsEvent.CursorLocation = adUseClient
   rsEvent.CursorType = adOpenStatic
   rsEvent.LockType = adLockReadOnly
   rsEvent.Open "SELECT TOP 100 * FROM Event INNER JOIN (Entity INNER JOIN Sensor ON Entity.cp_Entity = Sensor.fk_Entity)" & _
   " ON (Event.fk_Sensor = Sensor.Serial_Number) AND (Event.fk_Entity = Entity.cp_Entity) ORDER BY Date_Event DESC, Hour_Event DESC", cnDB

'  Design the lstEvent collection
   On Error Resume Next
   Dim nCount As Integer
   nCount = 0
   While Not rsEvent.EOF
      nCount = nCount + 1
      If nCount <= 100 Then
         Set tEvent = New clsEvent
         tEvent.sUIDo = rsEvent("fk_Sensor")
         tEvent.evDate = rsEvent("Date_Event")
         rsTipo = rsEvent("Tipo_Sensor")
         For idxTipo = 0 To 8
            If rsTipo = strTipo(idxTipo) Then
               tEvent.evTipo = idxTipo
               Exit For
            End If
         Next idxTipo
         tEvent.evStr = rsEvent("Descr_Event")
         tEvent.evDescr = rsEvent("Descr_Entity")
         lstEvent.Add tEvent
         Set tEntity = lstEntity.Item(CStr(rsEvent("fk_Entity")))
         tEntity.EventAdd tEvent
         rsEvent.MoveNext
      Else
         rsEvent.MoveLast
         rsEvent.MoveNext
      End If
      IncProgress
   Wend
   rsEvent.Close
   On Error GoTo 0
End Sub

Public Sub Data_CleanUp()
   Dim lrs As New ADODB.Recordset
   lrs.Open "SELECT * FROM Horario", cnDB, adOpenStatic, adLockReadOnly
   If Not lrs.EOF Then
      m_dTOpen(vbSunday) = lrs("TransLDomOpen")
      m_dTClose(vbSunday) = lrs("TransLDomClose")
      m_dTOpen(vbSaturday) = lrs("TransLSabOpen")
      m_dTClose(vbSaturday) = lrs("TransLSabClose")
      Dim i As Integer
      For i = vbMonday To vbFriday
         m_dTOpen(i) = lrs("TransLSegOpen")
         m_dTClose(i) = lrs("TransLSegClose")
      Next i
      m_iEvKeep = lrs("KeepData")
   Else
      MsgBox sxDatabase + Chr$(13) + sxContact, sxCritical, sxProname
   End If
   lrs.Close
End Sub

Public Sub DBase_ReOpen(ByVal fIsRestore As Boolean, Optional fPiso As Integer)
      
   Screen.MousePointer = vbHourglass
   frmSplash.Show
   If fIsRestore Then
      frmSplash.lblMsg = "Reconfigurando o Sistema"
   Else
      frmSplash.lblMsg = "Removendo o Piso " & fPiso & " e todos seus componentes."
   End If
   DoEvents
   
   If fIsRestore Then
      cnDB.Open
   End If
   
   Dim lM As clsModule
   For Each lM In lstModule
      lstModule.Remove 1
   Next
   Dim lE As clsEntity
   For Each lE In lstEntity
      lstEntity.Remove 1
   Next

   Dim lp As clsPercurso
   For Each lp In lstPercurso
      lp.RemoveRondas
      lstPercurso.Remove 1
   Next
      
   If fIsRestore Then
      Dim lPiso As clsPiso
      For Each lPiso In lstPiso
         Unload lPiso.f
         Unload ForNet.lblPiso(lPiso.n)
         lstPiso.Remove 1
      Next
   End If
   
   Load_Entities
   Load_Percursos
   Load_PTI
   
   If fIsRestore Then
      m_bUserUnload = True
      m_iCurPiso = 0
      m_iLastTop = 90
      ForNet.Pisos_Setup
      ForNet.Load_Pisos
      On Error Resume Next
      ForNet.ActiveForm.Entity_Refresh
      On Error GoTo 0
   End If
   
   Unload frmSplash
   Screen.MousePointer = vbDefault
   
End Sub
