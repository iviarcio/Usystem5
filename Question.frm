VERSION 5.00
Begin VB.Form frmQuestion 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "USystemEco"
   ClientHeight    =   1800
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6285
   Icon            =   "Question.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1800
   ScaleWidth      =   6285
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdAll 
      Caption         =   "Sim p/ todos"
      Height          =   405
      Left            =   1855
      TabIndex        =   3
      Top             =   1215
      Width           =   1170
   End
   Begin VB.CommandButton CmdNo 
      Caption         =   "Não"
      Height          =   405
      Left            =   3230
      TabIndex        =   2
      Top             =   1215
      Width           =   1170
   End
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Não p/ todos"
      Height          =   405
      Left            =   4605
      TabIndex        =   1
      Top             =   1215
      Width           =   1170
   End
   Begin VB.CommandButton cmdYes 
      Caption         =   "Sim"
      Default         =   -1  'True
      Height          =   405
      Left            =   480
      TabIndex        =   0
      Top             =   1215
      Width           =   1170
   End
   Begin VB.Label lblSensor 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   615
      Left            =   990
      TabIndex        =   6
      Top             =   180
      Width           =   5100
   End
   Begin VB.Label Label2 
      Caption         =   "deve ser de alta segurança. Você confirma a desativação?"
      Height          =   300
      Left            =   255
      TabIndex        =   5
      Top             =   840
      Width           =   5835
   End
   Begin VB.Label Label1 
      Caption         =   "A Zona "
      Height          =   255
      Left            =   285
      TabIndex        =   4
      Top             =   180
      Width           =   645
   End
End
Attribute VB_Name = "frmQuestion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CmdAll_Click()
   qResponse = sxQAll
   Unload Me
End Sub

Private Sub cmdCancel_Click()
   qResponse = sxQCancel
   Unload Me
End Sub

Private Sub CmdNo_Click()
   qResponse = sxQNo
   Unload Me
End Sub

Private Sub cmdYes_Click()
   qResponse = sxQYes
   Unload Me
End Sub

Private Sub Form_Load()
   Dim success As Long
   success = SetWindowPos(frmQuestion.hwnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
   curhwnd = frmQuestion.hwnd
   Left = (Screen.Width - Width) / 2   ' Center form horizontally.
   Top = (Screen.Height - Height) / 2  ' Center form vertically.
End Sub
