VERSION 5.00
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Object = "{A4749554-0441-4E64-8A03-3323601631C7}#1.0#0"; "LaVolpeAlphaImg2.ocx"
Begin VB.Form frmCadastro 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Relatório de Locais/Lojas Cadastradas"
   ClientHeight    =   5580
   ClientLeft      =   2265
   ClientTop       =   2565
   ClientWidth     =   10170
   ClipControls    =   0   'False
   Icon            =   "Cadastro.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5580
   ScaleWidth      =   10170
   ShowInTaskbar   =   0   'False
   Begin TrueOleDBGrid80.TDBGrid tdbg1 
      Height          =   4575
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   8070
      _LayoutType     =   4
      _RowHeight      =   18
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   " Entity"
      Columns(0).DataField=   ""
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   " Piso"
      Columns(1).DataField=   ""
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   " Local"
      Columns(2).DataField=   ""
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   " Responsável"
      Columns(3).DataField=   ""
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   " Tel. local"
      Columns(4).DataField=   ""
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   " Tel. externo"
      Columns(5).DataField=   ""
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   6
      Splits(0)._UserFlags=   0
      Splits(0).ExtendRightColumn=   -1  'True
      Splits(0).MarqueeStyle=   3
      Splits(0).AllowRowSizing=   0   'False
      Splits(0).RecordSelectorWidth=   688
      Splits(0)._SavedRecordSelectors=   -1  'True
      Splits(0).ScrollBars=   2
      Splits(0).AllowColSelect=   0   'False
      Splits(0).AlternatingRowStyle=   -1  'True
      Splits(0).DividerColor=   14215660
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=6"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=1191"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1085"
      Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
      Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=256"
      Splits(0)._ColumnProps(6)=   "Column(0).Visible=0"
      Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(8)=   "Column(1).Width=1323"
      Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=1217"
      Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
      Splits(0)._ColumnProps(12)=   "Column(1)._ColStyle=256"
      Splits(0)._ColumnProps(13)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(14)=   "Column(2).Width=5556"
      Splits(0)._ColumnProps(15)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(16)=   "Column(2)._WidthInPix=5450"
      Splits(0)._ColumnProps(17)=   "Column(2)._EditAlways=0"
      Splits(0)._ColumnProps(18)=   "Column(2)._ColStyle=256"
      Splits(0)._ColumnProps(19)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(20)=   "Column(3).Width=4075"
      Splits(0)._ColumnProps(21)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(22)=   "Column(3)._WidthInPix=3969"
      Splits(0)._ColumnProps(23)=   "Column(3)._EditAlways=0"
      Splits(0)._ColumnProps(24)=   "Column(3)._ColStyle=256"
      Splits(0)._ColumnProps(25)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(26)=   "Column(4).Width=2805"
      Splits(0)._ColumnProps(27)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(28)=   "Column(4)._WidthInPix=2699"
      Splits(0)._ColumnProps(29)=   "Column(4)._EditAlways=0"
      Splits(0)._ColumnProps(30)=   "Column(4)._ColStyle=256"
      Splits(0)._ColumnProps(31)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(32)=   "Column(5).Width=2302"
      Splits(0)._ColumnProps(33)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(34)=   "Column(5)._WidthInPix=2196"
      Splits(0)._ColumnProps(35)=   "Column(5)._EditAlways=0"
      Splits(0)._ColumnProps(36)=   "Column(5)._ColStyle=65792"
      Splits(0)._ColumnProps(37)=   "Column(5).Order=6"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   0
      PrintInfos(0).PageHeaderFont=   "Size=11.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Tahoma"
      PrintInfos(0).PageFooterFont=   "Size=11.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Tahoma"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      AllowUpdate     =   0   'False
      DataMode        =   4
      DefColWidth     =   0
      HeadLines       =   1
      FootLines       =   1
      RowDividerStyle =   0
      MultipleLines   =   0
      CellTipsWidth   =   0
      DeadAreaBackColor=   13160660
      RowDividerColor =   14215660
      RowSubDividerColor=   14215660
      DirectionAfterEnter=   1
      DirectionAfterTab=   1
      MaxRows         =   250000
      ViewColumnCaptionWidth=   0
      ViewColumnWidth =   0
      _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
      _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=0,.valignment=0,.bgcolor=&H80000005&"
      _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
      _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
      _StyleDefs(3)   =   ":id=0,.borderColor=&H0&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=29,.bold=0,.fontsize=1125,.italic=0"
      _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(8)   =   ":id=1,.fontname=Tahoma"
      _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=33"
      _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=30"
      _StyleDefs(11)  =   "FooterStyle:id=3,.parent=1,.namedParent=31"
      _StyleDefs(12)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(13)  =   "SelectedStyle:id=6,.parent=1,.namedParent=32"
      _StyleDefs(14)  =   "EditorStyle:id=7,.parent=1"
      _StyleDefs(15)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=34,.bgcolor=&H80000005&"
      _StyleDefs(16)  =   ":id=8,.fgcolor=&H8000000D&"
      _StyleDefs(17)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=35"
      _StyleDefs(18)  =   "OddRowStyle:id=10,.parent=1,.namedParent=36"
      _StyleDefs(19)  =   "RecordSelectorStyle:id=71,.parent=2,.namedParent=73"
      _StyleDefs(20)  =   "FilterBarStyle:id=74,.parent=1,.namedParent=76"
      _StyleDefs(21)  =   "Splits(0).Style:id=37,.parent=1"
      _StyleDefs(22)  =   "Splits(0).CaptionStyle:id=46,.parent=4"
      _StyleDefs(23)  =   "Splits(0).HeadingStyle:id=38,.parent=2"
      _StyleDefs(24)  =   "Splits(0).FooterStyle:id=39,.parent=3"
      _StyleDefs(25)  =   "Splits(0).InactiveStyle:id=40,.parent=5"
      _StyleDefs(26)  =   "Splits(0).SelectedStyle:id=42,.parent=6"
      _StyleDefs(27)  =   "Splits(0).EditorStyle:id=41,.parent=7"
      _StyleDefs(28)  =   "Splits(0).HighlightRowStyle:id=43,.parent=8"
      _StyleDefs(29)  =   "Splits(0).EvenRowStyle:id=44,.parent=9"
      _StyleDefs(30)  =   "Splits(0).OddRowStyle:id=45,.parent=10"
      _StyleDefs(31)  =   "Splits(0).RecordSelectorStyle:id=72,.parent=71"
      _StyleDefs(32)  =   "Splits(0).FilterBarStyle:id=75,.parent=74"
      _StyleDefs(33)  =   "Splits(0).Columns(0).Style:id=24,.parent=37,.alignment=0,.locked=0"
      _StyleDefs(34)  =   "Splits(0).Columns(0).HeadingStyle:id=21,.parent=38,.alignment=0"
      _StyleDefs(35)  =   "Splits(0).Columns(0).FooterStyle:id=22,.parent=39,.alignment=3"
      _StyleDefs(36)  =   "Splits(0).Columns(0).EditorStyle:id=23,.parent=41"
      _StyleDefs(37)  =   "Splits(0).Columns(1).Style:id=28,.parent=37,.alignment=0,.locked=0"
      _StyleDefs(38)  =   "Splits(0).Columns(1).HeadingStyle:id=25,.parent=38,.alignment=0"
      _StyleDefs(39)  =   "Splits(0).Columns(1).FooterStyle:id=26,.parent=39,.alignment=3"
      _StyleDefs(40)  =   "Splits(0).Columns(1).EditorStyle:id=27,.parent=41"
      _StyleDefs(41)  =   "Splits(0).Columns(2).Style:id=50,.parent=37,.alignment=0,.locked=0"
      _StyleDefs(42)  =   "Splits(0).Columns(2).HeadingStyle:id=47,.parent=38,.alignment=0"
      _StyleDefs(43)  =   "Splits(0).Columns(2).FooterStyle:id=48,.parent=39,.alignment=3"
      _StyleDefs(44)  =   "Splits(0).Columns(2).EditorStyle:id=49,.parent=41"
      _StyleDefs(45)  =   "Splits(0).Columns(3).Style:id=54,.parent=37,.alignment=0,.locked=0"
      _StyleDefs(46)  =   "Splits(0).Columns(3).HeadingStyle:id=51,.parent=38,.alignment=0"
      _StyleDefs(47)  =   "Splits(0).Columns(3).FooterStyle:id=52,.parent=39,.alignment=3"
      _StyleDefs(48)  =   "Splits(0).Columns(3).EditorStyle:id=53,.parent=41"
      _StyleDefs(49)  =   "Splits(0).Columns(4).Style:id=58,.parent=37,.alignment=0,.locked=0"
      _StyleDefs(50)  =   "Splits(0).Columns(4).HeadingStyle:id=55,.parent=38,.alignment=0"
      _StyleDefs(51)  =   "Splits(0).Columns(4).FooterStyle:id=56,.parent=39,.alignment=3"
      _StyleDefs(52)  =   "Splits(0).Columns(4).EditorStyle:id=57,.parent=41"
      _StyleDefs(53)  =   "Splits(0).Columns(5).Style:id=70,.parent=37"
      _StyleDefs(54)  =   "Splits(0).Columns(5).HeadingStyle:id=67,.parent=38"
      _StyleDefs(55)  =   "Splits(0).Columns(5).FooterStyle:id=68,.parent=39"
      _StyleDefs(56)  =   "Splits(0).Columns(5).EditorStyle:id=69,.parent=41"
      _StyleDefs(57)  =   "Named:id=29:Normal"
      _StyleDefs(58)  =   ":id=29,.parent=0"
      _StyleDefs(59)  =   "Named:id=30:Heading"
      _StyleDefs(60)  =   ":id=30,.parent=29,.valignment=2,.bgcolor=&H808000&,.fgcolor=&H80000012&"
      _StyleDefs(61)  =   ":id=30,.wraptext=-1"
      _StyleDefs(62)  =   "Named:id=31:Footing"
      _StyleDefs(63)  =   ":id=31,.parent=29,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(64)  =   "Named:id=32:Selected"
      _StyleDefs(65)  =   ":id=32,.parent=29,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(66)  =   "Named:id=33:Caption"
      _StyleDefs(67)  =   ":id=33,.parent=30,.alignment=2"
      _StyleDefs(68)  =   "Named:id=34:HighlightRow"
      _StyleDefs(69)  =   ":id=34,.parent=29,.bgcolor=&H80000008&,.fgcolor=&H80000005&"
      _StyleDefs(70)  =   "Named:id=35:EvenRow"
      _StyleDefs(71)  =   ":id=35,.parent=29,.bgcolor=&HFFFF&"
      _StyleDefs(72)  =   "Named:id=36:OddRow"
      _StyleDefs(73)  =   ":id=36,.parent=29"
      _StyleDefs(74)  =   "Named:id=73:RecordSelector"
      _StyleDefs(75)  =   ":id=73,.parent=30"
      _StyleDefs(76)  =   "Named:id=76:FilterBar"
      _StyleDefs(77)  =   ":id=76,.parent=29"
   End
   Begin LaVolpeAlphaImg.AlphaImgCtl cmdExit 
      Height          =   720
      Left            =   9405
      ToolTipText     =   "Fechar Relatório de  Locais/Lojas"
      Top             =   4680
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   1270
      Image           =   "Cadastro.frx":0442
      Effects         =   "Cadastro.frx":1147
   End
   Begin LaVolpeAlphaImg.AlphaImgCtl cmdPrint 
      Height          =   720
      Left            =   8520
      ToolTipText     =   "Relatório de Locais/Lojas"
      Top             =   4680
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   1270
      Image           =   "Cadastro.frx":115F
      Effects         =   "Cadastro.frx":230A
   End
End
Attribute VB_Name = "frmCadastro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mList As XArrayDB

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
   SetHourGlass Me
   Dim frm As New frmViewReport9
   frm.SetTipo = g_iRptCLocais
   frm.WindowState = vbMaximized
   frm.Show
   ResetMouse Me
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

Private Sub Form_Activate()
   Dim cE As clsEntity
   Dim mRow As Integer
   Dim mCol As Integer
   Set mList = Nothing
   cmdPrint.Enabled = (lstEntity.Count >= 1)
   If lstEntity.Count >= 1 Then
      ' Allocate space for rows, 6 columns
      Set mList = New XArrayDB
      mList.ReDim 0, lstEntity.Count - 1, 0, 5
      mRow = 0
      For Each cE In lstEntity
         With cE
            mList(mRow, 0) = CStr(.vId)
            mList(mRow, 1) = CStr(.floor)
            mList(mRow, 2) = .vDescr
            mList(mRow, 3) = .vResp
            mList(mRow, 4) = .vTel1
            mList(mRow, 5) = .vTel2
            mRow = mRow + 1
         End With
      Next
      mList.ReDim 0, mRow - 1, 0, 5
      mList.QuickSort 0, mRow - 1, 1, XORDER_ASCEND, XTYPE_STRING, 2, XORDER_ASCEND, XTYPE_STRING
      tdbg1.Array = mList
      tdbg1.ReBind
   End If
End Sub

Private Sub Form_Load()
   Dim success As Long
   success = SetWindowPos(frmCadastro.hWnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
   curhwnd = frmCadastro.hWnd
   Left = (Screen.Width - Width) / 2   ' Center form horizontally.
   Top = (Screen.Height - Height) / 2  ' Center form vertically.
   tdbg1.EvenRowStyle.BackColor = &H80FFFF
   tdbg1.OddRowStyle.BackColor = &HC0FFFF
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set mList = Nothing
End Sub

Private Sub tdbg1_SelChange(Cancel As Integer)
   Dim tEntity As clsEntity
   On Error Resume Next
   Set tEntity = lstEntity.Item(tdbg1.Columns(0))
   If Not (tEntity Is Nothing) Then
      Load frmEntity
      With frmEntity
         Set .fEntity = tEntity   'retain the Entity
         .mnuRemove.Enabled = m_bDesignMode
         .mnuMonitor.Visible = (m_sUser = sxAuthor)
      End With
      frmEntity.Show
   End If
End Sub

