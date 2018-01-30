Attribute VB_Name = "SysSound"
Option Explicit

'Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, _
'                                                                     ByVal uFlags As Long) As Long

' This call was modified to the one below! see:
' http://msdn.microsoft.com/en-us/library/windows/desktop/dd743680%28v=vs.85%29.aspx

'Sound API declaration
Public Declare Function sndPlaySound Lib "winmm.dll" Alias "PlaySoundA" _
    (ByVal lpszName As String, _
     ByVal hModule As Long, _
     ByVal dwFlags As Long) As Long

'Sound API Constants
Private Const sndApplication As Long = &H80
Private Const sndAlias As Long = &H10000
Private Const sndAliasID As Long = &H110000
Private Const sndASync As Long = &H1
Private Const sndFilename As Long = &H20000
Private Const sndLoop As Long = &H8
Private Const sndMemory As Long = &H4
Private Const sndNoDefault As Long = &H2
Private Const sndNoStop As Long = &H10
Private Const sndNoWait As Long = &H2000
Private Const sndPurge As Long = &H40              ' to stop a non-waveform sound (not supported)
Private Const sndResource As Long = &H40004
Private Const sndSync As Long = &H0

'Requires Windows Vista or Later
Private Const sndSEntry As Long = &H80000          ' SoundSentry is an accessibility feature that causes the computer
                                                   ' to display visual cue when a sound is played.
Private Const SND_ALIAS_START = 0
Private Const SND_ALIAS_SYSTEMASTERISK = 10835     ' "SystemAsterisk" event.
Private Const SND_ALIAS_SYSTEMDEFAULT = 17491      ' "SystemDefault" event.
Private Const SND_ALIAS_SYSTEMEXCLAMATION = 8531   ' "SystemExclamation" event.
Private Const SND_ALIAS_SYSTEMEXIT = 17747         ' "SystemExit" event.
Private Const SND_ALIAS_SYSTEMHAND = 18515         ' "SystemHand" event.
Private Const SND_ALIAS_SYSTEMQUESTION = 16211     ' "SystemQuestion" event.
Private Const SND_ALIAS_SYSTEMSTART = 21331        ' "SystemStart" event.
Private Const SND_ALIAS_SYSTEMWELCOME = 22355      ' "SystemWelcome" event.

Private Const sndActive As Long = 1&

'Sound Modes used by USystem
Public Enum SoundMode
   sxEvSound
   sxExSound
   sxErSound
   sxBgSound
   sxNoSound
   sxCtSound
   sxAlSound
   sxMsSound
   sxPnSound
   sxTIntrus
   sxTPanico
   sxTIncend
   sxTEmerge
   sxTSistem
End Enum

'Sound Variables used by USystem
Public m_sEvSound As String
Public m_sExSound As String
Public m_sErSound As String
Public m_sBgSound As String
Public m_sAlSound As String
Public m_sPnSound As String

Public m_sTIntrus As String
Public m_sTPanico As String
Public m_sTIncend As String
Public m_sTEmerge As String
Public m_sTSistem As String

'Sound Controls used by USystem
Private fDummy As Long
Private fSound As Long
Private fLoopC As Long

'Maximum number of tentatives to emit async loop sound
Private Const maxTentative = 5

Public Sub Sound_Init()
   fSound = 0&
   
   m_sBgSound = App.Path & "\Mensagens\Melorise.wav"
   m_sEvSound = App.Path & "\Mensagens\Alert.wav"
   m_sAlSound = App.Path & "\Mensagens\Siren1.wav"
   m_sErSound = App.Path & "\Mensagens\Phnbusy.wav"
   m_sExSound = App.Path & "\Mensagens\spkgliss.wav"
   m_sPnSound = App.Path & "\Mensagens\Alert.wav"
   
   m_sTIntrus = App.Path & "\Mensagens\Tampa_intrusao.wav"
   m_sTPanico = App.Path & "\Mensagens\Tampa_panico.wav"
   m_sTIncend = App.Path & "\Mensagens\Tampa_incendio.wav"
   m_sTEmerge = App.Path & "\Mensagens\Tampa_emergencia.wav"
   m_sTSistem = App.Path & "\Mensagens\Tampa_sistema.wav"
   
   If IsWindowsVistaOrGreater Then
      fLoopC = sndFilename Or sndLoop Or sndASync  ' or sndSEntry     # commented because didin't work
   Else
      fLoopC = sndFilename Or sndLoop Or sndASync
   End If
   
End Sub

Public Sub Sound_Update(ByVal fmode As SoundMode, ByVal isCritico As Boolean, ByVal fNoSound As Boolean, Optional fFile As String)
   
   If fNoSound Then Exit Sub
   
   Dim ltentative As Integer
   Dim lFile As String
   lFile = ""
   If Not IsMissing(fFile) Then
      On Error GoTo FileError
      If Left$(fFile, 1) <> " " Then
         If Dir(fFile) <> "" Then
            lFile = fFile
         End If
      End If
   End If
DisplaySound:
   Select Case fmode
      Case sxBgSound:
         fDummy = sndPlaySound(vbNullString, 0, 0)
         DoEvents
         fDummy = sndPlaySound(m_sBgSound, 0, sndFilename Or sndASync)
         fSound = 0&
      Case sxExSound:
         If fSound <> sndActive Then
            fDummy = sndPlaySound(vbNullString, 0, 0)
            DoEvents
            fDummy = sndPlaySound(m_sExSound, 0, sndFilename Or sndASync)
         End If
      Case sxEvSound:
         If fSound <> sndActive Then
            fDummy = sndPlaySound(vbNullString, 0, 0)
            DoEvents
            If lFile = "" Then
               fDummy = sndPlaySound(m_sEvSound, 0, sndFilename Or sndASync)
            Else
               fDummy = sndPlaySound(lFile, 0, sndFilename Or sndASync)
            End If
         End If
      Case sxErSound:
         fDummy = sndPlaySound(vbNullString, 0, 0)
         DoEvents
         fSound = sndPlaySound(m_sErSound, 0, fLoopC)
         If fSound <> sndActive Then
            ltentative = 0
            While (fSound <> sndActive) And (ltentative < maxTentative)
               DoEvents
               ltentative = ltentative + 1
               fSound = sndPlaySound(m_sErSound, 0, fLoopC)
            Wend
         End If
      Case sxNoSound:
         If fSound = sndActive Then
            fDummy = sndPlaySound(vbNullString, 0, 0)
            DoEvents
            fDummy = sndPlaySound(vbNullString, 0, 0)
            DoEvents
            fSound = 0&
         End If
       Case sxAlSound:
         If fSound <> sndActive Then
            fDummy = sndPlaySound(vbNullString, 0, 0)
            DoEvents
            If lFile = "" Then
               fSound = sndPlaySound(m_sAlSound, 0, fLoopC)
               If fSound <> sndActive Then
                  ltentative = 0
                  While (fSound <> sndActive) And (ltentative < maxTentative)
                     DoEvents
                     ltentative = ltentative + 1
                     fSound = sndPlaySound(m_sAlSound, 0, fLoopC)
                  Wend
               End If
            Else
               fSound = sndPlaySound(lFile, 0, fLoopC)
               If fSound <> sndActive Then
                  ltentative = 0
                  While (fSound <> sndActive) And (ltentative < maxTentative)
                     DoEvents
                     ltentative = ltentative + 1
                     fSound = sndPlaySound(lFile, 0, fLoopC)
                  Wend
               End If
            End If
         End If
       Case sxPnSound:
         fDummy = sndPlaySound(vbNullString, 0, 0)
         DoEvents
         If Not isCritico Then Load frmPanico
         DoEvents
         If lFile = "" Then
            fSound = sndPlaySound(m_sPnSound, 0, fLoopC)
            If fSound <> sndActive Then
               ltentative = 0
               While (fSound <> sndActive) And (ltentative < maxTentative)
                  DoEvents
                  ltentative = ltentative + 1
                  fSound = sndPlaySound(m_sPnSound, 0, fLoopC)
               Wend
            End If
         Else
            fSound = sndPlaySound(lFile, 0, fLoopC)
            If fSound <> sndActive Then
               ltentative = 0
               While (fSound <> sndActive) And (ltentative < maxTentative)
                  DoEvents
                  ltentative = ltentative + 1
                  fSound = sndPlaySound(lFile, 0, fLoopC)
               Wend
            End If
         End If
       Case sxTPanico:
         fDummy = sndPlaySound(vbNullString, 0, 0)
         DoEvents
         If Not isCritico Then Load frmPanico
         DoEvents
         fSound = sndPlaySound(m_sTPanico, 0, fLoopC)
         If fSound <> sndActive Then
            ltentative = 0
            While (fSound <> sndActive) And (ltentative < maxTentative)
               DoEvents
               ltentative = ltentative + 1
               fSound = sndPlaySound(m_sTPanico, 0, fLoopC)
            Wend
         End If
       Case sxTIntrus:
         If fSound <> sndActive Then
            fDummy = sndPlaySound(vbNullString, 0, 0)
            DoEvents
            fSound = sndPlaySound(m_sTIntrus, 0, fLoopC)
            If fSound <> sndActive Then
               ltentative = 0
               While (fSound <> sndActive) And (ltentative < maxTentative)
                  DoEvents
                  ltentative = ltentative + 1
                  fSound = sndPlaySound(m_sTIntrus, 0, fLoopC)
               Wend
            End If
         End If
       Case sxTIncend:
         If fSound <> sndActive Then
            fDummy = sndPlaySound(vbNullString, 0, 0)
            DoEvents
            fSound = sndPlaySound(m_sTIncend, 0, fLoopC)
            If fSound <> sndActive Then
               ltentative = 0
               While (fSound <> sndActive) And (ltentative < maxTentative)
                  DoEvents
                  ltentative = ltentative + 1
                  fSound = sndPlaySound(m_sTIncend, 0, fLoopC)
               Wend
            End If
         End If
       Case sxTEmerge:
         If fSound <> sndActive Then
            fDummy = sndPlaySound(vbNullString, 0, 0)
            DoEvents
            fSound = sndPlaySound(m_sTEmerge, 0, fLoopC)
            If fSound <> sndActive Then
               ltentative = 0
               While (fSound <> sndActive) And (ltentative < maxTentative)
                  DoEvents
                  ltentative = ltentative + 1
                  fSound = sndPlaySound(m_sTEmerge, 0, fLoopC)
               Wend
            End If
         End If
       Case sxTSistem:
         If fSound <> sndActive Then
            fDummy = sndPlaySound(vbNullString, 0, 0)
            DoEvents
            fSound = sndPlaySound(m_sTSistem, 0, fLoopC)
            If fSound <> sndActive Then
               ltentative = 0
               While (fSound <> sndActive) And (ltentative < maxTentative)
                  DoEvents
                  ltentative = ltentative + 1
                  fSound = sndPlaySound(m_sTSistem, 0, fLoopC)
               Wend
            End If
         End If
  End Select
  DoEvents
  Exit Sub
FileError:
   'File Not Found
   Err.Clear
   lFile = ""
   Resume DisplaySound
End Sub

Private Sub Show_Display(ByVal fIsCritico As Boolean, ByVal fNoSound As Boolean)
   On Error Resume Next
   Dim lDisplay As clsDisplay
   Set lDisplay = lstDisplay.Item(1)
   m_UpdateLock = False 'To show the message
   ForNet.Update_Display lDisplay.dispStr, lDisplay.dispImg, True
   If lDisplay.dispFile <> "" Then
      Sound_Update lDisplay.dispMode, fIsCritico, fNoSound, lDisplay.dispFile
   ElseIf Dir(App.Path & "\Mensagens\" & lDisplay.dispFile) <> "" Then
      Sound_Update lDisplay.dispMode, fIsCritico, fNoSound, App.Path & "\Mensagens\" & lDisplay.dispFile
   ElseIf lDisplay.dispMode <> sxNoSound Then
      Sound_Update lDisplay.dispMode, fIsCritico, fNoSound
   End If
   m_UpdateLock = True
End Sub

Public Sub Insert_Display(fDisp As clsDisplay, fIsPanico As Boolean, fIsCritico As Boolean, Optional fNoSound As Boolean = False)
    On Error GoTo sndError
    If fIsPanico Or fDisp.dispMode = sxPnSound Then
        If lstDisplay.Count >= 20 Then
            lstDisplay.Remove 1
        End If
        lstDisplay.Add Item:=fDisp
        If Not m_UpdateLock Then Show_Display fIsCritico, fNoSound
        nDisplay = lstDisplay.Count
        ForNet.StatusBar1.Panels.Item(3).Text = nDisplay
   ElseIf Not m_UpdateLock Then
      ForNet.Update_Display fDisp.dispStr, fDisp.dispImg, True
      If fDisp.dispFile = "" Then
         If fDisp.dispMode <> sxNoSound Then
            Sound_Update fDisp.dispMode, fIsCritico, fNoSound
         End If
      ElseIf Dir(fDisp.dispFile) <> "" Then
         Sound_Update fDisp.dispMode, fIsCritico, fNoSound, fDisp.dispFile
      Else
         Sound_Update fDisp.dispMode, fIsCritico, fNoSound, App.Path & "\Mensagens\" & fDisp.dispFile
      End If
   End If
   Exit Sub
sndError:
   Sound_Update fDisp.dispMode, fIsCritico, fNoSound
End Sub

Public Sub Remove_Display()
   On Error Resume Next
   lstDisplay.Remove 1
   nDisplay = lstDisplay.Count
   If nDisplay > 0 Then
      Show_Display False, False
   Else
      m_UpdateLock = False
   End If
   ForNet.StatusBar1.Panels.Item(3).Text = nDisplay
End Sub

