VERSION 5.00
Object = "{A4749554-0441-4E64-8A03-3323601631C7}#1.0#0"; "LaVolpeAlphaImg2.ocx"
Begin VB.Form frmAccess 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro de Acesso"
   ClientHeight    =   4005
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6405
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Access.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4005
   ScaleWidth      =   6405
   Begin VB.Frame Frame1 
      Caption         =   "Nível de &Acesso"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   1455
      Left            =   120
      TabIndex        =   2
      Top             =   2420
      Width           =   5055
      Begin VB.ListBox lstType 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   870
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   3495
      End
      Begin LaVolpeAlphaImg.AlphaImgCtl btnAccess 
         Height          =   720
         Index           =   0
         Left            =   3960
         ToolTipText     =   "Alterar o nível de acesso"
         Top             =   360
         Width           =   720
         _ExtentX        =   1270
         _ExtentY        =   1270
         Image           =   "Access.frx":030A
         Effects         =   "Access.frx":1433
      End
   End
   Begin VB.Frame fraUser 
      Caption         =   "Usuário"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   2415
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   6135
      Begin VB.ComboBox lstEmployee 
         BackColor       =   &H00FFFFFF&
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
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   600
         Width           =   3255
      End
      Begin LaVolpeAlphaImg.AlphaImgCtl btnAccess 
         Height          =   720
         Index           =   6
         Left            =   4440
         ToolTipText     =   "Alterar a senha do usuário selecionado"
         Top             =   1440
         Width           =   960
         _ExtentX        =   1693
         _ExtentY        =   1270
         Image           =   "Access.frx":144B
         Effects         =   "Access.frx":25D8
      End
      Begin LaVolpeAlphaImg.AlphaImgCtl btnAccess 
         Height          =   720
         Index           =   2
         Left            =   5040
         ToolTipText     =   "Excluir o usuário selecionado"
         Top             =   360
         Width           =   720
         _ExtentX        =   1270
         _ExtentY        =   1270
         Image           =   "Access.frx":25F0
         Effects         =   "Access.frx":3708
      End
      Begin LaVolpeAlphaImg.AlphaImgCtl btnAccess 
         Height          =   720
         Index           =   4
         Left            =   3960
         ToolTipText     =   "Inserir um novo usuário"
         Top             =   360
         Width           =   720
         _ExtentX        =   1270
         _ExtentY        =   1270
         Image           =   "Access.frx":3720
         Effects         =   "Access.frx":46B1
      End
   End
   Begin VB.Frame fraNew 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   120
      TabIndex        =   4
      Top             =   0
      Width           =   6135
      Begin VB.TextBox txt1 
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
         Left            =   1680
         MaxLength       =   50
         TabIndex        =   6
         Top             =   360
         Width           =   3135
      End
      Begin VB.TextBox txt2 
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
         IMEMode         =   3  'DISABLE
         Left            =   1680
         MaxLength       =   30
         PasswordChar    =   "*"
         TabIndex        =   8
         Top             =   960
         Width           =   3135
      End
      Begin VB.TextBox txt3 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         IMEMode         =   3  'DISABLE
         Left            =   1680
         MaxLength       =   30
         PasswordChar    =   "*"
         TabIndex        =   10
         Top             =   1545
         Width           =   3135
      End
      Begin LaVolpeAlphaImg.AlphaImgCtl btnAccess 
         Height          =   720
         Index           =   1
         Left            =   5160
         ToolTipText     =   "Cancelar"
         Top             =   360
         Width           =   720
         _ExtentX        =   1270
         _ExtentY        =   1270
         Image           =   "Access.frx":46C9
         Effects         =   "Access.frx":57AD
      End
      Begin LaVolpeAlphaImg.AlphaImgCtl btnAccess 
         Height          =   720
         Index           =   5
         Left            =   5160
         ToolTipText     =   "Confirma a inserção do novo Usuário"
         Top             =   1440
         Width           =   720
         _ExtentX        =   1270
         _ExtentY        =   1270
         Image           =   "Access.frx":57C5
         Effects         =   "Access.frx":68DD
      End
      Begin VB.Label lbl1 
         Alignment       =   1  'Right Justify
         Caption         =   "&Nome:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label lbl2 
         Alignment       =   1  'Right Justify
         Caption         =   "&Senha:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label lbl3 
         Alignment       =   1  'Right Justify
         Caption         =   "&Confirmação:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   1560
         Width           =   1455
      End
   End
   Begin LaVolpeAlphaImg.AlphaImgCtl btnAccess 
      Height          =   960
      Index           =   3
      Left            =   5280
      ToolTipText     =   "Fechar Cadastro de Acesso"
      Top             =   2760
      Width           =   960
      _ExtentX        =   1693
      _ExtentY        =   1693
      Image           =   "Access.frx":68F5
      Effects         =   "Access.frx":7C14
   End
End
Attribute VB_Name = "frmAccess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const NONE = -1
Private Const EMPNEW = 0
Private Const ALTPAS = 1

Private fmode As Integer
Private fupdate As Integer

Private rsEmployee As New ADODB.Recordset

Private Sub SetAppearence(btn As AlphaImgCtl, flag As Boolean)
   If flag Then
      btn.GrayScale = lvicNoGrayScale
   Else
      btn.GrayScale = lvicGreenMask
   End If
   btn.Enabled = flag
End Sub

Private Sub btnAccess_Click(Index As Integer)
   Select Case Index
      Case 0
         cmdAlter
      Case 1
         cmdCancel
      Case 2
         cmdDelete
      Case 3
         Unload Me
      Case 4
         cmdNew
      Case 5
         cmdOk
      Case 6
         cmdPassword
   End Select
End Sub

Private Sub btnAccess_MouseEnter(Index As Integer)
   If btnAccess(Index).Enabled Then
      btnAccess(Index).SetRedraw = False
      btnAccess(Index).GrayScale = lvicSepia
      btnAccess(Index).LightnessPct = -20
      btnAccess(Index).SetRedraw = True
   End If
End Sub

Private Sub btnAccess_MouseExit(Index As Integer)
   If btnAccess(Index).Enabled Then
      btnAccess(Index).SetRedraw = False
      btnAccess(Index).GrayScale = lvicNoGrayScale
      btnAccess(Index).LightnessPct = 0
      btnAccess(Index).SetRedraw = True
   End If
End Sub

Private Sub Insert_Employee()
   Dim lcm As New ADODB.Command
   Set lcm.ActiveConnection = cnDB
   lcm.CommandType = adCmdText
   lcm.CommandText = "INSERT INTO Employee (Name, Type, [Password]) VALUES ('" & _
                     txt1 & "', " & lstType.ListIndex & ", '" & XOREncryption(strKeyCode, txt2) & "')"
   lcm.Execute
   Make_Service "Inclusão do Operador: " & txt1.Text, strAccess(m_tAccess) & m_sUser
   Query_Employee
End Sub

Private Sub Alter_Passwd()
    Dim lok As Boolean
    Dim fStr As String
    lok = False
'    If IsNull(rsEmployee("Password")) And (txt1 = "") Then
        lok = (txt2 = txt3) And txt2 <> ""
'    ElseIf rsEmployee("Password") = txt1 Then
'        lok = (txt2 = txt3) And txt2 <> ""
'    End If
    If lok Then
      Dim lcm As New ADODB.Command
      Set lcm.ActiveConnection = cnDB
      lcm.CommandType = adCmdText
      lcm.CommandText = "UPDATE Employee SET [Password] = '" & XOREncryption(strKeyCode, txt2) & "' WHERE " & _
                        "Employee.cp_Employee = " & rsEmployee("cp_Employee")
      lcm.Execute
      Make_Service "Alteração de Senha do Usuário: " & rsEmployee("Name"), strAccess(m_tAccess) & m_sUser
      Query_Employee
    Else
        Beep
        MsgBox "Dados incorretos.", sxExclamation, sxProname
    End If
End Sub

Private Sub cmdAlter()
   If (m_tAccess = sxSupervisor) And (lstType.ListIndex = 2) Then
      Beep
      Exit Sub
   End If
   Dim lcm As New ADODB.Command
   Set lcm.ActiveConnection = cnDB
   lcm.CommandType = adCmdText
   lcm.CommandText = "UPDATE Employee SET Type = " & lstType.ListIndex & " WHERE " & _
                     "Employee.cp_Employee = " & rsEmployee("cp_Employee")
   lcm.Execute
   SetAppearence btnAccess(0), False
   Make_Service "Alteração de Nível de Acesso: " & rsEmployee("Name"), strAccess(m_tAccess) & m_sUser
   Query_Employee
End Sub

Private Sub cmdPassword()
   If (m_tAccess = sxSupervisor) And (lstType.ListIndex = 2) Then
      Beep
      Exit Sub
   End If
   SetAppearence btnAccess(0), False
   fmode = ALTPAS
   lbl1.Visible = False
   txt1.Visible = False
   lbl2 = "Nova &Senha:"
   txt2 = ""
   txt3 = ""
   SetAppearence btnAccess(5), True
   btnAccess(5).ToolTipText = "Comfirma a alteração de senha."
   fraNew.Caption = "Alteração de Senha: "
   fraNew.ZOrder 0
   txt2.SetFocus
End Sub

Private Sub New_Employee()
   Dim crt$
   If (txt1 = "") Or (txt3 = "") Or (txt2 <> txt3) Then
      Beep
      MsgBox "Dados incorretos.", sxExclamation, sxProname
   Else
      Dim lrs As New ADODB.Recordset
      lrs.Open "SELECT * FROM Employee WHERE (Name = '" & CStr(txt1.Text) & "')", cnDB, adOpenStatic, adLockReadOnly
      If Not lrs.EOF Then
         lrs.Close
         Beep
         MsgBox "O Operador/Administrador já está cadastrado!", sxExclamation, sxProname
         Exit Sub
      Else
         lrs.Close
         Insert_Employee
      End If
   End If
End Sub

Private Sub Update_LstEmployee()
    lstEmployee.Clear
    If m_tAccess <> sxOperator Then
      rsEmployee.MoveFirst
      While Not rsEmployee.EOF
          lstEmployee.AddItem rsEmployee("Name")
          rsEmployee.MoveNext
      Wend
      rsEmployee.MoveFirst
    Else
      rsEmployee.Filter = "Name = '" & m_sUser & "'"
      lstEmployee.AddItem rsEmployee("Name")
    End If
    lstEmployee.ListIndex = 0
End Sub

Private Sub Update_LstType()
    lstType.Selected(rsEmployee("Type")) = True
    SetAppearence btnAccess(0), False
End Sub

Private Sub Form_Load()
   lstType.AddItem "Operador"
   lstType.AddItem "Administrador"
   lstType.AddItem "Sistema"
   lstType.Enabled = m_tAccess <> sxOperator
   fmode = NONE
   
   Query_Employee
   
   SetAppearence btnAccess(0), False
   SetAppearence btnAccess(4), m_tAccess <> sxOperator
   SetAppearence btnAccess(2), m_tAccess <> sxOperator
   Left = (Screen.Width - Width) / 2   ' Center form horizontally.
   Top = (Screen.Height - Height) / 2  ' Center form vertically.
End Sub

Private Sub lstEmployee_Click()
   rsEmployee.MoveFirst
   Dim lFound As Boolean
   lFound = False
   While Not lFound
      If rsEmployee("Name") = lstEmployee.Text Then
         lFound = True
      Else
         rsEmployee.MoveNext
      End If
   Wend
   Update_LstType
End Sub

Private Sub lstType_Click()
   If fmode <> EMPNEW Then
      If m_tAccess = sxAdministrador Then
         If rsEmployee("Name") <> "FOR" Then
            SetAppearence btnAccess(0), True
         Else
            Beep
            Update_LstType
         End If
      Else
         Beep
         Update_LstType
      End If
   ElseIf m_tAccess <> sxAdministrador And lstType.ListIndex = 2 Then
      Beep
      lstType.ListIndex = 1
   End If
End Sub

Private Sub cmdCancel()
    SetAppearence btnAccess(0), False
    fraUser.ZOrder 0
    SetAppearence btnAccess(5), False
    fmode = NONE
End Sub

Private Sub cmdDelete()
   SetAppearence btnAccess(0), False
   If rsEmployee("Name") = "FOR" Then
      MsgBox "Sistema 'FOR' não pode ser excluído.", sxInformation, sxProname
      Exit Sub
   ElseIf (m_tAccess = sxSupervisor) And (lstType.ListIndex = 2) Then
      MsgBox "Administrador não pode excluir Usuário de Sistema!", sxInformation, sxProname
      Exit Sub
   End If
   Dim res%
   res% = MsgBox("Deseja excluir o Usuário selecionado?", sxQuestion, sxProname)
   If res% = vbYes Then
      Dim lcm As New ADODB.Command
      Set lcm.ActiveConnection = cnDB
      Make_Service "Exclusão de Operador: " & rsEmployee("Name"), strAccess(m_tAccess) & m_sUser
      lcm.CommandType = adCmdText
      lcm.CommandText = "DELETE FROM Employee WHERE " & _
                        "Employee.cp_Employee = " & rsEmployee("cp_Employee")
      lcm.Execute
      Query_Employee
   End If
End Sub

Private Sub cmdNew()
    SetAppearence btnAccess(0), False
    fmode = EMPNEW
    lbl1.Visible = True
    lbl1 = "&Nome:"
    txt1.Visible = True
    txt1.PasswordChar = ""
    txt1 = ""
    lbl2 = "&Senha:"
    txt2 = ""
    txt3 = ""
    SetAppearence btnAccess(5), True
    btnAccess(5).ToolTipText = "Confirma a inserção do novo Usuário"
    fraNew.Caption = "Novo &Usuário: "
    fraNew.ZOrder 0
    lstType.ListIndex = 0
    txt1.SetFocus
End Sub

Private Sub cmdOk()
    SetAppearence btnAccess(0), False
    If fmode = EMPNEW Then
        New_Employee
    Else
        Alter_Passwd
    End If
    fraUser.ZOrder 0
    fmode = NONE
End Sub

Private Sub txt1_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        txt2.SetFocus
    ElseIf KeyAscii <> vbKeyTab Then
        If Len(txt1) >= 50 Then
            KeyAscii = 0
            Beep
        End If
    End If
End Sub

Private Sub txt2_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        txt3.SetFocus
    ElseIf KeyAscii <> vbKeyTab Then
        If Len(txt2) >= 50 Then
            KeyAscii = 0
            Beep
        End If
    End If
End Sub

Private Sub txt3_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        Call cmdOk
    ElseIf KeyAscii <> vbKeyTab Then
        If Len(txt3) >= 50 Then
            KeyAscii = 0
            Beep
        End If
    End If
End Sub

Private Sub Query_Employee()
   On Error Resume Next
   rsEmployee.Close
   On Error GoTo 0
   rsEmployee.Filter = ""
   rsEmployee.Open "SELECT * FROM Employee ORDER BY Name", cnDB, adOpenStatic, adLockReadOnly
   On Error GoTo 0
   Update_LstEmployee
   Update_LstType
End Sub
