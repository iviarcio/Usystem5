VERSION 5.00
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "todg6.ocx"
Begin VB.Form frmBstatus 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Status da Rede"
   ClientHeight    =   5115
   ClientLeft      =   2265
   ClientTop       =   2565
   ClientWidth     =   6870
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
   Icon            =   "Bstatus.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5115
   ScaleWidth      =   6870
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   2000
      Left            =   120
      Top             =   4680
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "&Sair"
      Height          =   735
      Left            =   6120
      Picture         =   "Bstatus.frx":0742
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Fechar Situação Corrente da Rede"
      Top             =   4320
      Width           =   735
   End
   Begin TrueOleDBGrid60.TDBGrid tdbg1 
      Height          =   4215
      Left            =   0
      OleObjectBlob   =   "Bstatus.frx":0A4C
      TabIndex        =   1
      Top             =   0
      Width           =   6855
   End
End
Attribute VB_Name = "frmBstatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mList As XArrayDB

Private Sub cmdExit_Click()
   Unload Me
End Sub

Private Sub Form_Activate()
   Dim mRow As Integer
   Dim mCol As Integer
   
'   If lstBase.Count >= 1 Then
   
      ' Allocate space for rows, 6 columns
      Set mList = New XArrayDB
      mList.ReDim 0, lstBase.Count - 1, 0, 5
      mRow = 0
      
'      For Each cB In lstBase
      
'         With cB
         
'            mList(mRow, 0) = .SysId
            
'            If .flag_Ativ Then
'               mList(mRow, 1) = "Comunicação"
'            Else
'               mList(mRow, 1) = "Falha"
'            End If
            
'            mList(mRow, 2) = .VersaoRede
            
'            Select Case .flag_Status
'
'               Case bOKS
'                  mList(mRow, 3) = "Nenhum"
'                  mList(mRow, 4) = " _ "
'               Case bOKW
'                  mList(mRow, 3) = "Wired"
'                  mList(mRow, 4) = " _ "
'               Case bOKF
'                  mList(mRow, 3) = "Wireless"
'                  mList(mRow, 4) = .receptorID
'            End Select
'
''            mList(mRow, 5) = .VersaoPeriferico
'            mRow = mRow + 1
            
'         End With
         
'      Next
      
      mList.ReDim 0, mRow - 1, 0, 5
      mList.QuickSort 0, mRow - 1, 0, XORDER_ASCEND, XTYPE_INTEGER
      tdbg1.Array = mList
      tdbg1.ReBind
      
'   End If
End Sub

Private Sub Form_Load()
   Dim success As Long
   success = SetWindowPos(frmBstatus.hwnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
   curhwnd = frmBstatus.hwnd
   Left = (Screen.Width - Width) / 2   ' Center form horizontally.
   Top = (Screen.Height - Height) / 2  ' Center form vertically.
   tdbg1.EvenRowStyle.BackColor = &H80FFFF
   tdbg1.OddRowStyle.BackColor = &HC0FFFF
End Sub

Private Sub Timer1_Timer()
   Set mList = Nothing
   Form_Activate
End Sub
