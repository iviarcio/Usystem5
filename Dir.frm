VERSION 5.00
Object = "{A4749554-0441-4E64-8A03-3323601631C7}#1.0#0"; "LaVolpeAlphaImg2.ocx"
Begin VB.Form frmDir 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Diretório para Backup Automático"
   ClientHeight    =   3945
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7440
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Dir.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3945
   ScaleWidth      =   7440
   StartUpPosition =   3  'Windows Default
   Begin VB.DriveListBox Drive1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   120
      TabIndex        =   1
      Top             =   3360
      Width           =   6255
   End
   Begin VB.DirListBox Dir1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3090
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6255
   End
   Begin LaVolpeAlphaImg.AlphaImgCtl cmdOk 
      Height          =   720
      Left            =   6600
      ToolTipText     =   "Selecionar"
      Top             =   3120
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   1270
      Image           =   "Dir.frx":0442
      Effects         =   "Dir.frx":0FE1
   End
   Begin LaVolpeAlphaImg.AlphaImgCtl cmdCancel 
      Height          =   720
      Left            =   6600
      ToolTipText     =   "Cancelar"
      Top             =   2160
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   1270
      Image           =   "Dir.frx":0FF9
      Effects         =   "Dir.frx":1C45
   End
End
Attribute VB_Name = "frmDir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public fCaller As Form

Private Sub cmdCancel_Click()
   Unload Me
End Sub

Private Sub cmdCancel_MouseEnter()
   cmdCancel.SetRedraw = False
   cmdCancel.GrayScale = lvicSepia
   cmdCancel.LightnessPct = -20
   cmdCancel.SetRedraw = True
End Sub

Private Sub cmdCancel_MouseExit()
   cmdCancel.SetRedraw = False
   cmdCancel.GrayScale = lvicNoGrayScale
   cmdCancel.LightnessPct = 0
   cmdCancel.SetRedraw = True
End Sub

Private Sub cmdOk_Click()
   fCaller.txtLocal = Dir1.List(Dir1.ListIndex)
   Unload Me
End Sub

Private Sub cmdOk_MouseEnter()
   cmdOk.SetRedraw = False
   cmdOk.GrayScale = lvicSepia
   cmdOk.LightnessPct = -20
   cmdOk.SetRedraw = True
End Sub

Private Sub cmdOk_MouseExit()
   cmdOk.SetRedraw = False
   cmdOk.GrayScale = lvicNoGrayScale
   cmdOk.LightnessPct = 0
   cmdOk.SetRedraw = True
End Sub

Private Sub Drive1_Change()
   On Error GoTo DriveError
   Dir1.Path = Drive1.Drive
   Exit Sub
DriveError:
   MsgBox "O dispositivo não está disponível!", sxInformation, sxProname
   Drive1.Drive = "C:"
End Sub

Private Sub Form_Load()
    Left = (Screen.Width - Width) / 2   ' Center form horizontally.
    Top = (Screen.Height - Height) / 2  ' Center form vertically.
End Sub
