VERSION 5.00
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Object = "{A4749554-0441-4E64-8A03-3323601631C7}#1.0#0"; "LaVolpeAlphaImg2.ocx"
Begin VB.Form frmStLocal 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Situação Corrente dos Locais"
   ClientHeight    =   5490
   ClientLeft      =   2265
   ClientTop       =   2565
   ClientWidth     =   8790
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
   Icon            =   "StLocal.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5490
   ScaleWidth      =   8790
   ShowInTaskbar   =   0   'False
   Begin TrueOleDBGrid80.TDBGrid tdbg1 
      Height          =   4575
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   8070
      _LayoutType     =   4
      _RowHeight      =   18
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "Entity"
      Columns(0).DataField=   ""
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   " Local"
      Columns(1).DataField=   ""
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   " Tipo"
      Columns(2).DataField=   ""
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   16
      Columns(3)._MaxComboItems=   5
      Columns(3).ValueItems(0)._DefaultItem=   0
      Columns(3).ValueItems(0).Value=   "0"
      Columns(3).ValueItems(0).Value.vt=   8
      Columns(3).ValueItems(0).DisplayValue=   "Fechado"
      Columns(3).ValueItems(0).DisplayValue.vt=   8
      Columns(3).ValueItems(0)._PropDict=   "_DefaultItem,517,2"
      Columns(3).ValueItems(1)._DefaultItem=   0
      Columns(3).ValueItems(1).Value=   "1"
      Columns(3).ValueItems(1).Value.vt=   8
      Columns(3).ValueItems(1).DisplayValue=   "Aberto"
      Columns(3).ValueItems(1).DisplayValue.vt=   8
      Columns(3).ValueItems(1)._PropDict=   "_DefaultItem,517,2"
      Columns(3).ValueItems(2)._DefaultItem=   0
      Columns(3).ValueItems(2).Value=   "2"
      Columns(3).ValueItems(2).Value.vt=   8
      Columns(3).ValueItems(2).DisplayValue=   "Curto"
      Columns(3).ValueItems(2).DisplayValue.vt=   8
      Columns(3).ValueItems(2)._PropDict=   "_DefaultItem,517,2"
      Columns(3).ValueItems(3)._DefaultItem=   0
      Columns(3).ValueItems(3).Value=   "3"
      Columns(3).ValueItems(3).Value.vt=   8
      Columns(3).ValueItems(3).DisplayValue=   "Falha"
      Columns(3).ValueItems(3).DisplayValue.vt=   8
      Columns(3).ValueItems(3)._PropDict=   "_DefaultItem,517,2"
      Columns(3).ValueItems(4)._DefaultItem=   0
      Columns(3).ValueItems(4).Value=   "4"
      Columns(3).ValueItems(4).Value.vt=   8
      Columns(3).ValueItems(4).DisplayValue=   "Tamper"
      Columns(3).ValueItems(4).DisplayValue.vt=   8
      Columns(3).ValueItems(4)._PropDict=   "_DefaultItem,517,2"
      Columns(3).ValueItems(5)._DefaultItem=   0
      Columns(3).ValueItems(5).Value=   "5"
      Columns(3).ValueItems(5).Value.vt=   8
      Columns(3).ValueItems(5).DisplayValue=   "_ "
      Columns(3).ValueItems(5).DisplayValue.vt=   8
      Columns(3).ValueItems(5)._PropDict=   "_DefaultItem,517,2"
      Columns(3).ValueItems.Count=   6
      Columns(3).Caption=   " Status"
      Columns(3).DataField=   ""
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   4
      Splits(0)._UserFlags=   0
      Splits(0).ExtendRightColumn=   -1  'True
      Splits(0).MarqueeStyle=   3
      Splits(0).AllowRowSizing=   0   'False
      Splits(0).RecordSelectorWidth=   979
      Splits(0)._SavedRecordSelectors=   -1  'True
      Splits(0).ScrollBars=   2
      Splits(0).AllowColSelect=   0   'False
      Splits(0).AlternatingRowStyle=   -1  'True
      Splits(0).DividerColor=   15790320
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=4"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=1191"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1085"
      Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
      Splits(0)._ColumnProps(5)=   "Column(0).AllowSizing=0"
      Splits(0)._ColumnProps(6)=   "Column(0)._ColStyle=256"
      Splits(0)._ColumnProps(7)=   "Column(0).Visible=0"
      Splits(0)._ColumnProps(8)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(9)=   "Column(1).Width=8096"
      Splits(0)._ColumnProps(10)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(11)=   "Column(1)._WidthInPix=7990"
      Splits(0)._ColumnProps(12)=   "Column(1)._EditAlways=0"
      Splits(0)._ColumnProps(13)=   "Column(1).AllowSizing=0"
      Splits(0)._ColumnProps(14)=   "Column(1)._ColStyle=256"
      Splits(0)._ColumnProps(15)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(16)=   "Column(2).Width=2725"
      Splits(0)._ColumnProps(17)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(18)=   "Column(2)._WidthInPix=2619"
      Splits(0)._ColumnProps(19)=   "Column(2)._EditAlways=0"
      Splits(0)._ColumnProps(20)=   "Column(2).AllowSizing=0"
      Splits(0)._ColumnProps(21)=   "Column(2)._ColStyle=256"
      Splits(0)._ColumnProps(22)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(23)=   "Column(3).Width=1931"
      Splits(0)._ColumnProps(24)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(25)=   "Column(3)._WidthInPix=1826"
      Splits(0)._ColumnProps(26)=   "Column(3)._EditAlways=0"
      Splits(0)._ColumnProps(27)=   "Column(3).AllowSizing=0"
      Splits(0)._ColumnProps(28)=   "Column(3)._ColStyle=256"
      Splits(0)._ColumnProps(29)=   "Column(3).Order=4"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   0
      PrintInfos(0).Name=   "piInternal 0"
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
      RowDividerColor =   15790320
      RowSubDividerColor=   15790320
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
      _StyleDefs(19)  =   "RecordSelectorStyle:id=63,.parent=2,.namedParent=65"
      _StyleDefs(20)  =   "FilterBarStyle:id=66,.parent=1,.namedParent=68"
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
      _StyleDefs(31)  =   "Splits(0).RecordSelectorStyle:id=64,.parent=63"
      _StyleDefs(32)  =   "Splits(0).FilterBarStyle:id=67,.parent=66"
      _StyleDefs(33)  =   "Splits(0).Columns(0).Style:id=24,.parent=37,.alignment=0,.locked=0"
      _StyleDefs(34)  =   "Splits(0).Columns(0).HeadingStyle:id=21,.parent=38,.alignment=0"
      _StyleDefs(35)  =   "Splits(0).Columns(0).FooterStyle:id=22,.parent=39,.alignment=3"
      _StyleDefs(36)  =   "Splits(0).Columns(0).EditorStyle:id=23,.parent=41"
      _StyleDefs(37)  =   "Splits(0).Columns(1).Style:id=50,.parent=37,.alignment=0,.locked=0"
      _StyleDefs(38)  =   "Splits(0).Columns(1).HeadingStyle:id=47,.parent=38,.alignment=0"
      _StyleDefs(39)  =   "Splits(0).Columns(1).FooterStyle:id=48,.parent=39,.alignment=3"
      _StyleDefs(40)  =   "Splits(0).Columns(1).EditorStyle:id=49,.parent=41"
      _StyleDefs(41)  =   "Splits(0).Columns(2).Style:id=58,.parent=37,.alignment=0,.locked=0"
      _StyleDefs(42)  =   "Splits(0).Columns(2).HeadingStyle:id=55,.parent=38,.alignment=0"
      _StyleDefs(43)  =   "Splits(0).Columns(2).FooterStyle:id=56,.parent=39,.alignment=3"
      _StyleDefs(44)  =   "Splits(0).Columns(2).EditorStyle:id=57,.parent=41"
      _StyleDefs(45)  =   "Splits(0).Columns(3).Style:id=62,.parent=37,.alignment=0,.locked=0"
      _StyleDefs(46)  =   "Splits(0).Columns(3).HeadingStyle:id=59,.parent=38,.alignment=0"
      _StyleDefs(47)  =   "Splits(0).Columns(3).FooterStyle:id=60,.parent=39,.alignment=3"
      _StyleDefs(48)  =   "Splits(0).Columns(3).EditorStyle:id=61,.parent=41"
      _StyleDefs(49)  =   "Named:id=29:Normal"
      _StyleDefs(50)  =   ":id=29,.parent=0"
      _StyleDefs(51)  =   "Named:id=30:Heading"
      _StyleDefs(52)  =   ":id=30,.parent=29,.valignment=2,.bgcolor=&H808000&,.fgcolor=&H80000012&"
      _StyleDefs(53)  =   ":id=30,.wraptext=-1"
      _StyleDefs(54)  =   "Named:id=31:Footing"
      _StyleDefs(55)  =   ":id=31,.parent=29,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(56)  =   "Named:id=32:Selected"
      _StyleDefs(57)  =   ":id=32,.parent=29,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(58)  =   "Named:id=33:Caption"
      _StyleDefs(59)  =   ":id=33,.parent=30,.alignment=2"
      _StyleDefs(60)  =   "Named:id=34:HighlightRow"
      _StyleDefs(61)  =   ":id=34,.parent=29,.bgcolor=&HFF0000&,.fgcolor=&H80000005&"
      _StyleDefs(62)  =   "Named:id=35:EvenRow"
      _StyleDefs(63)  =   ":id=35,.parent=29,.bgcolor=&HFFFF&"
      _StyleDefs(64)  =   "Named:id=36:OddRow"
      _StyleDefs(65)  =   ":id=36,.parent=29"
      _StyleDefs(66)  =   "Named:id=65:RecordSelector"
      _StyleDefs(67)  =   ":id=65,.parent=30"
      _StyleDefs(68)  =   "Named:id=68:FilterBar"
      _StyleDefs(69)  =   ":id=68,.parent=29"
   End
   Begin LaVolpeAlphaImg.AlphaImgCtl cmdExit 
      Height          =   720
      Left            =   7920
      ToolTipText     =   "Fechar Visualização dos Locais"
      Top             =   4680
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   1270
      Image           =   "StLocal.frx":0442
      Effects         =   "StLocal.frx":1147
   End
   Begin LaVolpeAlphaImg.AlphaImgCtl cmdPrint 
      Height          =   720
      Left            =   7080
      ToolTipText     =   "Visualizar Situação Corrente dos Locais"
      Top             =   4680
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   1270
      Image           =   "StLocal.frx":115F
      Effects         =   "StLocal.frx":230A
   End
End
Attribute VB_Name = "frmStLocal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public localModule As New Collection
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
   Screen.MousePointer = vbHourglass
   'First prepare the table
   Dim lcmd As New ADODB.Command
   Set lcmd.ActiveConnection = cnDB
   lcmd.CommandType = adCmdText
   lcmd.CommandText = "DELETE FROM StatusCorrente"
   lcmd.Execute
   Dim cM As clsModule
   For Each cM In localModule
      With cM
         lcmd.CommandText = "INSERT INTO StatusCorrente (fk_Sensor, Status, Sinal, Bateria, Tampa) VALUES ('" & _
                            .Serial_Number & "', " & .SZona & ", " & .NivelSinal & ", " & .SLowBat & ", " & .STampa & ")"
         lcmd.Execute
      End With
   Next
   'Now Print the report
   Screen.MousePointer = vbHourglass
   Dim frm As New frmViewReport9
   frm.SetTipo = g_iRptSCLocais
   frm.WindowState = vbMaximized
   frm.Show
   Screen.MousePointer = vbDefault
   
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
   Dim fE As clsEntity
   Dim cM As clsModule
   Dim mRow As Integer
   Dim mCol As Integer
   Set mList = Nothing
   cmdPrint.Enabled = (localModule.Count >= 1)
   If localModule.Count >= 1 Then
      ' Allocate space for rows, 4 columns
      Set mList = New XArrayDB
      mList.ReDim 0, localModule.Count - 1, 0, 3
      mRow = 0
      For Each cM In localModule
         With cM
            Set fE = lstEntity.Item(CStr(.mEntity))
            mList(mRow, 0) = CStr(.mEntity)
            mList(mRow, 1) = fE.vDescr
            mList(mRow, 2) = strTipo(.mTipo)
            mList(mRow, 3) = .SZona
            mRow = mRow + 1
         End With
      Next
      mList.QuickSort 0, mRow - 1, 1, XORDER_ASCEND, XTYPE_STRING
      tdbg1.Array = mList
      tdbg1.ReBind
   End If
End Sub

Private Sub Form_Load()
   Dim success As Long
   success = SetWindowPos(frmStLocal.hWnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
   curhwnd = frmStLocal.hWnd
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
      Load frmZonas
      Set frmZonas.lEntity = tEntity
      frmZonas.Show
   End If
   On Error GoTo 0
End Sub
