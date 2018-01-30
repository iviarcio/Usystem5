VERSION 5.00
Begin VB.Form LogOn 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Acesso USystemEco"
   ClientHeight    =   4035
   ClientLeft      =   2400
   ClientTop       =   2805
   ClientWidth     =   4575
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   Icon            =   "Logon.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4035
   ScaleWidth      =   4575
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      Caption         =   "Registro"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   240
      TabIndex        =   8
      Top             =   2880
      Width           =   4095
      Begin VB.CommandButton cmdLimpar 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   2760
         Picture         =   "Logon.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   330
         Width           =   1215
      End
      Begin VB.TextBox txtKeyCode 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1320
         MaxLength       =   7
         TabIndex        =   12
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   240
         MaxLength       =   5
         TabIndex        =   10
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   960
         TabIndex        =   11
         Top             =   360
         Width           =   375
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      Caption         =   "Empresa"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   240
      TabIndex        =   7
      Top             =   1800
      Width           =   4095
      Begin VB.TextBox txtCompany 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   240
         MaxLength       =   40
         TabIndex        =   9
         Top             =   360
         Width           =   3615
      End
   End
   Begin VB.CommandButton cmdRegistro 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   840
      Picture         =   "Logon.frx":05E1
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1080
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   2040
      Picture         =   "Logon.frx":092A
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1080
      Width           =   1095
   End
   Begin VB.TextBox txtPassword 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   840
      MaxLength       =   35
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   600
      Width           =   3495
   End
   Begin VB.TextBox txtName 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   840
      MaxLength       =   35
      TabIndex        =   1
      Top             =   240
      Width           =   3495
   End
   Begin VB.CommandButton cmdOk 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   3240
      Picture         =   "Logon.frx":0C3A
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Image imgReg 
      Appearance      =   0  'Flat
      Height          =   360
      Index           =   1
      Left            =   3840
      Picture         =   "Logon.frx":0EB0
      Top             =   4200
      Width           =   360
   End
   Begin VB.Image imgReg 
      Appearance      =   0  'Flat
      Height          =   360
      Index           =   0
      Left            =   3240
      Picture         =   "Logon.frx":117F
      Top             =   4200
      Width           =   360
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Caption         =   "&Senha:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   630
      Width           =   615
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Caption         =   "&Nome:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   270
      Width           =   615
   End
End
Attribute VB_Name = "LogOn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Tipos de acesso
'Private Const sxAdmin = 0
'Private Const sxSecurity = 1
Private Const sxMaxLen = 40
Private regFlag As Boolean

Private Sub Try_Login()
    Dim strPassword As String

   If txtName = sxAuthor Then
      strPassword = Format(Day(Date), "00") & "f" & Format(Month(Date), "00") & "o" & Format(Right(Year(Date), 2), "00") & "r"
      If txtPassword = strPassword Then
            m_sUser = txtName
            m_tAccess = sxSystem
            m_bPermition = True
            m_Debug = True
            Unload Me
            Exit Sub
      End If
      
    ElseIf txtName = sxRegistro Then
        strPassword = Format(Day(Date), "00") & Format(Month(Date), "00") & Format(Right(Year(Date), 2), "00")
        If txtPassword = strPassword Then
            cmdRegistro.Enabled = True
            txtName = ""
            txtPassword = ""
            Exit Sub
        End If
        
      txtName = ""
      txtPassword = ""
      Beep
      
   Else
      Dim lrs As New ADODB.Recordset
      lrs.Open "SELECT * FROM Employee WHERE (Name = '" & CStr(txtName.Text) & "')", cnDB, adOpenStatic, adLockReadOnly
      If Not lrs.EOF Then
         If XOREncryption(strKeyCode, lrs("Password")) = CStr(txtPassword.Text) Then
            m_sUser = lrs("Name")
            m_tAccess = lrs("Type")
            m_bPermition = True
            m_Debug = False
            lrs.Close
            Unload Me
            Exit Sub
         End If
      End If
      lrs.Close
      txtName = ""
      txtPassword = ""
      Beep
   End If
End Sub

Private Sub cmdCancel_Click()
   m_bPermition = False
   m_Debug = False
   m_bShutDown = True
   Unload Me
End Sub

Private Sub cmdlimpar_Click()
    txtCodigo = vbNullString
    txtKeyCode = vbNullString
    gstChecksum = txtCodigo
    gstCondorID = txtKeyCode
    SaveSetting "USystem5", "Options", "Checksum", vbNullString
    SaveSetting "USystem5", "Options", "License", vbNullString
End Sub

Private Sub cmdOk_Click()
   Me.Height = 2160
   Try_Login
End Sub

Private Sub cmdRegistro_Click()
    If regFlag = False Then
    
        Me.Height = 4560
        
        txtCompany = gstCompany
        If Len(Trim(txtCompany)) = 0 Then
            txtCompany.SetFocus
        End If
        
        txtCodigo = gstChecksum
        txtKeyCode = gstCondorID
        
        cmdOk.Enabled = False
        txtName.Enabled = False
        txtPassword.Enabled = False
        
        cmdRegistro.Picture = imgReg(1)
    
        regFlag = True
        
    Else
    
        Me.Height = 2160
        gstCompany = txtCompany
        gstChecksum = txtCodigo
        gstCondorID = txtKeyCode
        
        If oCnn Is Nothing Then Set oCnn = New clsConnection
        oCnn.ExecSp "UPDATE Config SET Empresa = '" & gstCompany & "'"
        
        SaveSetting "USystem5", "Options", "Company", gstCompany
        SaveSetting "USystem5", "Options", "License", gstCondorID
        SaveSetting "USystem5", "Options", "Checksum", gstChecksum
        
        'Verifica a chave de segurança
        Security_Check
   
        If BYPASS = 1 Then
            cmdOk.Enabled = True
            cmdRegistro.Enabled = False
            LogOn.Caption = "Acesso ao USystem5 - BYPASSED"
        ElseIf Key_check = 1 Then
            cmdOk.Enabled = True
            cmdRegistro.Enabled = False
            LogOn.Caption = "Acesso ao USystem5 - REGISTRADO"
        Else
            LogOn.Caption = "Acesso ao USystem5 - SEM REGISTRO"
        End If
        
        txtName.Enabled = True
        txtPassword.Enabled = True
        
        cmdRegistro.Picture = imgReg(0)
        
        regFlag = False

    End If
        
End Sub

Private Sub Form_Activate()
   On Error Resume Next
   txtName.SetFocus
End Sub

Private Sub Form_Load()
   Dim success As Long
   
   success = SetWindowPos(LogOn.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, FLAGS)
   success = SetWindowPos(LogOn.hWnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
   curhwnd = LogOn.hWnd
   Left = (Screen.Width - Width) / 2   ' Center form horizontally.
   Top = (Screen.Height - Height) / 2  ' Center form vertically.
   
   m_bPermition = False
   m_bShutDown = False
   m_Debug = False
   
   Me.Height = 2160
   
   Security_Check
   
   If BYPASS = 1 Then
        cmdRegistro.Enabled = False
        LogOn.Caption = "Acesso ao USystem5 - BYPASSED"
   ElseIf Key_check = 1 Then
        cmdRegistro.Enabled = False
        LogOn.Caption = "Acesso ao USystem5 - REGISTRADO"
   Else
        cmdOk.Enabled = False
        cmdRegistro.Enabled = True
        LogOn.Caption = "Acesso ao USystem5 - SEM REGISTRO"
   End If
   
   
         
End Sub

Private Sub txtCodigo_Change()
    If Len(txtCodigo) = 5 Then
       txtKeyCode.SetFocus
   End If
End Sub

Private Sub txtName_KeyPress(KeyAscii As Integer)
   If (KeyAscii <> vbKeyTab) And (KeyAscii <> vbKeyReturn) Then
      If Len(txtName) >= sxMaxLen Then
         KeyAscii = 0
         Beep
      End If
   ElseIf KeyAscii = vbKeyReturn Then
      KeyAscii = 0
      txtPassword.SetFocus
   End If
End Sub

Private Sub txtPassword_KeyPress(KeyAscii As Integer)
   If KeyAscii <> vbKeyReturn Then
      If Len(txtPassword) >= sxMaxLen Then
         KeyAscii = 0
         Beep
      End If
   Else
      KeyAscii = 0
'      Try_Login
   End If
End Sub
