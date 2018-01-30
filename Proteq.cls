VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsproteq"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Declare Function C500 Lib "C50032.DLL" (ByVal Entra As String) As Integer
Dim ChaveIni As Integer
Dim strData As Variant
Dim Psenha As String
Dim offset As Long
Dim ch As Integer
Dim i As Integer

'Rotina que faz a leitura do código da chave e sua inicialização
Public Function Verifica(CodAcesso) As String
Dim A As String

A = String$(15, Chr$(0))

A = Chr$(3) & CodAcesso

    i = C500(A)
    
    If i <> 0 Then
        Verifica = "0"
    Else
        Verifica = Mid$(A$, 3, 7)
    End If
    
End Function

'Rotina que faz a leitura da data de expiração gravado na memória da chave
Public Function Leitura(ByVal DateString As Byte) As String

'    Dim i As Integer
    Dim A As String * 10
    Dim EndInicial As Integer
    Dim EndFinal As Integer
    Dim E As Integer
    Dim memoria As String
    Dim Temp1 As String
    Dim Temp2 As String

    A = String$(15, Chr$(0))
    memoria = String$(20, Chr$(0))

    If DateString = 1 Then
        EndInicial = 0
        EndFinal = 3
    ElseIf DateString = 2 Then
        EndInicial = 4
        EndFinal = 7
    End If
    
    'Le da memória fisica a primeira ou segunda sequência de datas (CONDOR ou CONDOR MAPS)
    For E = EndInicial To EndFinal
        A = Chr$(1) & "@@" & Chr$(E)
        i = C500(A)
        memoria = Hex(Asc(Mid$(A, 2, 1))) & Hex(Asc(Mid$(A, 3, 1)))
        If Len(Temp2) = 2 Or Len(Temp2) = 5 Then
            Temp2 = Temp2 & "/" & memoria
        Else
            Temp2 = Temp2 & memoria
        End If
    Next E
        If Not IsDate(Temp2) Then
            Leitura = "00/00/0000"
        Else
            Leitura = Temp2
        End If
    
End Function

'Rotina que faz a escrita de expiração na chave
Public Function Escrita(ByVal dExpira As String, ByVal DateString As Byte) As Boolean
'    Dim i As Integer
    Dim A As String
    Dim D As Integer
    Dim DadoInicial As Integer
    Dim DadoFinal As Integer
    Dim EndInicial As Integer
    Dim memoria As String
    
    A = String$(15, Chr$(0))
    memoria = String$(16, Chr$(0))
   
    memoria = Left(dExpira, 2) & Mid(dExpira, 4, 2) & Right(dExpira, 4)
           
    DadoInicial = 0
    DadoFinal = 3
    
    If DateString = 1 Then
        EndInicial = 0
    ElseIf DateString = 2 Then
        EndInicial = 4
    End If
   
    For D = DadoInicial To DadoFinal
        A = Chr$(2) & Chr$(Mid$(memoria, 2 * D + 1, 1)) & Chr$(Mid$(memoria, 2 * D + 2, 1)) & Chr$(EndInicial)
        i = C500(A)
        If i <> 0 Then
            Escrita = False
            Exit Function
        End If
        EndInicial = EndInicial + 1
        Next D
        Escrita = True
    
End Function

Public Function CriptSenha(dData As Variant, strKey As String) As Variant
    Dim StrCampo As String
    Dim Soma As Byte
    
    Const MIN_ASC = 48
    Const MAX_ASC = 57
    Const NUM_ASC = MAX_ASC - MIN_ASC + 1

    Soma = 0
    StrCampo = ""
    
    strData = Left(dData, 2) & Mid(dData, 4, 2) & Right(dData, 4)
    strData = strData & strKey
    
    For i = 1 To 8
        Soma = Soma + Mid(strData, i, 1)
    Next i
    If Soma < 13 Then Soma = Soma * 3
    
    ChaveIni = INICondor
    
    offset = NumericPassword(ChaveIni)
    Rnd -1
    Randomize offset
    str_len = Len(strData)
    For i = 1 To str_len
        ch = Asc(Mid$(strData, i, 1))
        If ch >= MIN_ASC And ch <= MAX_ASC Then
            ch = ch - MIN_ASC
            offset = Int((NUM_ASC + 1) * Rnd)
            ch = ((ch + offset) Mod NUM_ASC)
            ch = ch + MIN_ASC
            StrCampo = StrCampo & Chr$(ch)
        End If
    Next i
    
    StrCampo = StrCampo * Soma
    StrCampo = Left(Soma, 1) & StrCampo & Right(Soma, 1)
    
    CriptSenha = StrCampo
End Function

Public Function DeCriptSenhaKey(Fsenha As String, ByVal DateString As Byte) As String

    Const MIN_ASC = 48
    Const MAX_ASC = 57
    Const NUM_ASC = MAX_ASC - MIN_ASC + 1
    
    'Verifica qual a ChaveIni será usada, Condor = 1 e CondorMaps=2
    If DateString = 1 Then
        ChaveIni = INICondor
    ElseIf DateString = 2 Then
        ChaveIni = INICondorMaps
    End If
    
    Psenha = Replace$(Fsenha, "-", "")

    strData = Mid(Psenha, 3, 12)
    
    to_text = ""
    offset = NumericPassword(ChaveIni)
    Rnd -1
    Randomize offset
    str_len = Len(strData)
    For i = 1 To str_len
        ch = Asc(Mid$(strData, i, 1))
        If ch >= MIN_ASC And ch <= MAX_ASC Then
            ch = ch - MIN_ASC
            offset = Int((NUM_ASC + 1) * Rnd)
            ch = ((ch - offset) Mod NUM_ASC)
            If ch < 0 Then ch = ch + NUM_ASC
            ch = ch + MIN_ASC
            to_text = to_text & Chr$(ch)
        End If
    Next i
    
    DeCriptSenhaKey = Right(to_text, 4)
    
End Function

Public Function DeCriptSenhaData(ByVal Fsenha As String, ByVal DateString As Byte) As Date

    Dim dData As String
    Dim fData As String
    
    Const MIN_ASC = 48
    Const MAX_ASC = 57
    Const NUM_ASC = MAX_ASC - MIN_ASC + 1
    
    On Error GoTo ErrorTreat
    
    'Verifica qual a ChaveIni será usada, Condor = 1 e CondorMaps=2
    If DateString = 1 Then
        ChaveIni = INICondor
    ElseIf DateString = 2 Then
        ChaveIni = INICondorMaps
    End If
    
    Psenha = Replace$(Fsenha, "-", "")
    
    strData = Mid(Psenha, 3, 12)
    
    to_text = ""
    offset = NumericPassword(ChaveIni)
    Rnd -1
    Randomize offset
    str_len = Len(strData)
    For i = 1 To str_len
        ch = Asc(Mid$(strData, i, 1))
        If ch >= MIN_ASC And ch <= MAX_ASC Then
            ch = ch - MIN_ASC
            offset = Int((NUM_ASC + 1) * Rnd)
            ch = ((ch - offset) Mod NUM_ASC)
            If ch < 0 Then ch = ch + NUM_ASC
            ch = ch + MIN_ASC
            to_text = to_text & Chr$(ch)
        End If
    Next i
    
    fData = Left(to_text, 8)
    
    For i = 1 To Len(fData)
        dData = dData & Mid(to_text, i, 1)
        If i = 2 Or i = 4 Then
            dData = dData & ","
        End If
    Next i

    DeCriptSenhaData = CVDate(dData)
    Exit Function
    
ErrorTreat:
    frmValidade.lblinforma2 = vbCr & "Código de registro inválido!"
    
End Function

Private Function NumericPassword(ByVal Password As String) As Long
    Dim Value As Long
    Dim shift1 As Long
    Dim shift2 As Long

    str_len = Len(Password)
    For i = 1 To str_len
        ' Adiciona a próxima letra
        ch = Asc(Mid$(Password, i, 1))
        Value = Value Xor (ch * 2 ^ shift1)
        Value = Value Xor (ch * 2 ^ shift2)

        ' Change the shift offsets.
        shift1 = (shift1 + 7) Mod 19
        shift2 = (shift2 + 13) Mod 23
    Next i
    NumericPassword = Value
End Function
        
      
