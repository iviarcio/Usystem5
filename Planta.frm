VERSION 5.00
Begin VB.Form frmPlanta 
   AutoRedraw      =   -1  'True
   ClientHeight    =   9360
   ClientLeft      =   60
   ClientTop       =   120
   ClientWidth     =   12225
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   624
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   815
   Visible         =   0   'False
   WindowState     =   2  'Maximized
End
Attribute VB_Name = "frmPlanta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Identificação do Piso e referencia à classe
Public curPiso As Integer
Private rPiso As clsPiso

'Handle para o Form
Private lngHDC As Long

'Coordenadas auxiliares para a contrução da Entidade
Private m_t As MATRIZ
Private m_iVertices As Integer
Private bValidadeReg As Boolean

Private Sub Form_Activate()
   Entity_Refresh
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
   If KeyAscii = vbKeyEscape And m_bDesignMode Then
      If m_iVertices >= 1 Then
         m_iVertices = m_iVertices - 1
         Tracejado
         If m_iVertices = 0 Then m_DragState = StateDragging
      End If
   End If
End Sub

Private Sub Form_Load()
   'Armazena o handle do Form
   lngHDC = GetDC(Me.hWnd)
   'Demais incializações
   m_DragState = StateNothing
'   m_bSetores = False
   bValidadeReg = False
   m_iVertices = 0
   Set rPiso = lstPiso.Item(CStr(m_iCurPiso))
   rPiso.c_bSetores = False
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
   If m_bDesignMode And Button = vbLeftButton And m_DragState <> StateNothing Then
      m_iVertices = m_iVertices + 1
      If m_iVertices > MAX_POINTS Then
         MsgBox "Excedido o número de vértices da região! ", sxCritical, sxProname
         m_iVertices = m_iVertices - 2
         Tracejado
         Exit Sub
      End If
      m_t.varPoints(m_iVertices).x = x
      m_t.varPoints(m_iVertices).Y = Y
      If m_iVertices = 1 Then
         Me.CurrentX = x
         Me.CurrentY = Y
         Me.MousePointer = vbCrosshair
         m_DragState = StateSizing
         bValidadeReg = False
      Else
         Me.Line -(x, Y)
         If Abs(x - m_t.varPoints(1).x) <= 2 And Abs(Y - m_t.varPoints(1).Y) <= 2 Then
            If Not bValidadeReg Then
               MsgBox "Uma região deve conter 3 ou mais vértices diferentes!", sxCritical, sxProname
               m_iVertices = m_iVertices - 1
               Tracejado
            Else
               Me.Cls
               Entity_Create
               m_iVertices = 0
               Me.MousePointer = vbDefault
               m_DragState = StateNothing
            End If
         ElseIf m_iVertices >= 3 Then
            Dim d1 As Single, d2 As Single, delta As Single
            On Error GoTo DivisionError
            d1 = (m_t.varPoints(m_iVertices).x - m_t.varPoints(m_iVertices - 1).x) / _
                 (m_t.varPoints(m_iVertices).Y - m_t.varPoints(m_iVertices - 1).Y)
            d2 = (m_t.varPoints(m_iVertices).x - m_t.varPoints(m_iVertices - 2).x) / _
                 (m_t.varPoints(m_iVertices).Y - m_t.varPoints(m_iVertices - 2).Y)
            On Error GoTo 0
            delta = Abs(d1 - d2)
DeltaLabel:
            If delta >= 0.1 Then bValidadeReg = True
         End If
      End If
   ElseIf Button = vbRightButton Then
      'Retorna a entidade localizada em x,y.
      GetEntity curPiso, CLng(x), CLng(Y)
      If tEntity Is Nothing Then
      Else
         Load frmEntity
         With frmEntity
            Set .fEntity = tEntity   'retain the Entity
            .mnuRemove.Enabled = m_bDesignMode
            .mnuMonitor.Visible = (m_sUser = sxAuthor)
         End With
         frmEntity.Show
      End If
   End If
   Exit Sub
DivisionError:
   delta = 0
   Resume DeltaLabel
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
   If m_DragState = StateNothing Then
      'Retorna a entidade localizada em x,y.
      GetEntity curPiso, CLng(x), CLng(Y)
      If tEntity Is Nothing Then
         Me.MousePointer = vbDefault
         Exit Sub
      End If
      'Se existe uma região em x,y ...
      Me.MousePointer = vbIconPointer
      ForNet.Update_Display tEntity.vDescr, sxImgEntity, False
      Dim lngRet As Long
      lngRet = SelectObject(lngHDC, lngFill(tEntity.BackGround))
      lngRet = FillRgn(lngHDC, tEntity.Handle, lngFill(tEntity.BackGround))
   ElseIf m_DragState = StateSizing Then
      Tracejado
      Me.Line -(x, Y)
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If m_bUserUnload Then
      Cancel = True
   End If
End Sub

Public Sub Entity_Delete()
   If MsgBox("Confirma a remoção da Entidade? Isto implica em remover todos" & _
             " os módulos da entidade!", sxQuestion, sxProname) = vbYes Then
      Dim lModule As clsModule
      Dim lEntity As clsEntity
      Set lEntity = tEntity
      For Each lModule In lEntity.localModule
         lEntity.Remove lModule
         lstModule.Remove lModule.UID
      Next
      lEntity.Dump stRemoved
      lstEntity.Remove (CStr(lEntity.vId))
      Dim lngRet As Long
      lngRet = DeleteObject(lEntity.Handle)
      Set lEntity = Nothing
      ForNet.Update_Display "", sxImgNone, False
      Me.Refresh
   End If
End Sub

Private Sub Entity_Create()
   Dim i As Integer
   Set tEntity = New clsEntity
   With tEntity
      .Vertices = m_iVertices
      For i = 1 To m_iVertices
         .Set_Coordinates i, m_t.varPoints(i).x, m_t.varPoints(i).Y
      Next i
      .Handle = CreatePolygonRgn(m_t, .Vertices, ALTERNATE)
      .floor = curPiso
      .Dump stAdded
   End With
'  Seleciona o padrão de preenchimento.
   Dim lngRet As Long
   lngRet = SelectObject(lngHDC, lngFill(0))
   'Preenche a região com o padrão selecionado
   lngRet = FillRgn(lngHDC, tEntity.Handle, lngFill(0))
   With tEntity
      .vDescr = "Loja " & .vId
      .Dump stModified
   End With
   lstEntity.Add Item:=tEntity, Key:=CStr(tEntity.vId)
   Load frmEntity
   With frmEntity
      Set .fEntity = tEntity   'retain the Entity
      .mnuRemove.Enabled = m_bDesignMode
      .mnuMonitor.Visible = (m_sUser = sxAuthor)
      .Show
   End With
End Sub

Public Sub Entity_Edit()
   Load frmEntity
   With frmEntity
      Set .fEntity = tEntity
      .mnuRemove.Enabled = m_bDesignMode
      .mnuMonitor.Visible = (m_sUser = sxAuthor)
      .Show
   End With
End Sub

Public Sub Redesenha()
   'Utilizada para atualizar a exibição das regiões
   'quando o form é redimensionado ou foi ocultado
   'por outra janela. As regiões são apenas novamente
   'preenchidas, não sendo necessário recriá-las.
   'Mesmo estando a propriedade AutoRedraw definida como
   'true, o VB somente atualiza a exibição de controles
   'e do conteúdo da propriedade Picture. Ou seja, o
   'conteúdo das regiões não pertence ao VB, e sim ao
   'sistema, sendo assim, não serão redesenhados.
   Dim lngRet As Long
   Dim lEntity As clsEntity
   For Each lEntity In lstEntity
      If lEntity.floor = curPiso Then
         lngRet = SelectObject(lngHDC, lngFill(lEntity.BackGround))
         lngRet = FillRgn(lngHDC, lEntity.Handle, lngFill(lEntity.BackGround))
         If lEntity.BackGround = colRedGreen Then
            lEntity.colBack = colGreenRed
         ElseIf lEntity.BackGround = colGreenRed Then
            lEntity.colBack = colRedGreen
         End If
      End If
   Next
End Sub

Private Sub Tracejado()
   'Utilizada para atualizar a exibição da região
   'quando a mesma é redimensionada
   Dim i As Integer
   Cls
   If m_iVertices >= 1 Then
      Me.CurrentX = m_t.varPoints(1).x
      Me.CurrentY = m_t.varPoints(1).Y
      For i = 2 To m_iVertices
         Me.Line -(m_t.varPoints(i).x, m_t.varPoints(i).Y)
         Me.CurrentX = m_t.varPoints(i).x
         Me.CurrentY = m_t.varPoints(i).Y
      Next i
   ElseIf m_iVertices = 0 Then
      Me.MousePointer = vbDefault
   End If
End Sub

Public Sub Entity_Refresh()
   Dim lngRet As Long
   Dim mt As MATRIZ
   Dim i As Integer
   If Not rPiso.c_bSetores Then   'need to create all Entities!
      Dim lEntity As clsEntity
      For Each lEntity In lstEntity
         With lEntity
            If .floor = curPiso Then
               For i = 1 To .Vertices
                  .Get_Coordinates i, mt.varPoints(i).x, mt.varPoints(i).Y
               Next i
               .Handle = CreatePolygonRgn(mt, .Vertices, ALTERNATE)
               'Seleciona o padrão de preenchimento.
               lngRet = SelectObject(lngHDC, lngFill(.BackGround))
               'Preenche a região com o padrão selecionado
               lngRet = FillRgn(lngHDC, .Handle, lngFill(.BackGround))
               rPiso.c_bSetores = True
            End If
         End With
      Next
   End If
   m_DragState = StateNothing
End Sub
