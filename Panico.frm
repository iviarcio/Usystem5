VERSION 5.00
Begin VB.Form frmPanico 
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   ClientHeight    =   2010
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12540
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   DrawStyle       =   5  'Transparent
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2010
   ScaleWidth      =   12540
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer trmpanico 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   840
      Top             =   3000
   End
   Begin VB.Image Image1 
      Height          =   2130
      Left            =   0
      Picture         =   "Panico.frx":0000
      Stretch         =   -1  'True
      Top             =   -120
      Width           =   12570
   End
End
Attribute VB_Name = "frmPanico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private fvisible As Boolean

Private Sub Form_Load()
   Dim i As Integer
   ActiveTransparency Me, True, False, 255, Me.BackColor
   Me.Show
   Me.ZOrder 0
   fvisible = True
   trmpanico.Enabled = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
   trmpanico.Enabled = False
   fvisible = False
End Sub

Private Sub trmPanico_Timer()
   fvisible = Not fvisible
   frmPanico.Visible = fvisible
End Sub

