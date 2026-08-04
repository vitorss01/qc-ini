Attribute VB_Name = "mRegistros"
Option Explicit
' ===== RESULTADOS NAO CONFORMES (Sprint NC) =====
'
' O ANALISTA marca um resultado como nao conforme. O resultado SAI DOS CALCULOS
' e some da view, mas a linha PERMANECE no DB_Resultados com o valor original
' intacto -- ISO 15189 8.4.2 exige que o valor original siga recuperavel apos a
' emenda. A aba Registros passa a exibi-lo como ocorrencia documentada.
'
' ENCAIXE NA ESTRUTURA EXISTENTE, sem mexer em formula nenhuma. A aba Registros
' ja tinha as colunas certas; so faltava usa-las com este proposito:
'
'   B  Data          data da corrida
'   C  Analito
'   D  Nivel
'   E  RUN           era "Corrida (Seq)" -- passa a receber o RUN
'   F  Resultado NC  era "Rep 1" -- o grafico ja plota esta coluna como marcador
'   G,H              livres para repeticoes posteriores, digitadas depois
'   I  Foi calibrado?
'   J  Parecer Tecnico  (era "Observacao")
'   K  Responsavel      login do sistema
'   M  Tipo de NC       (era "Status")
'
' Nao foi preciso trocar Rep1/2/3 por uma coluna unica: 3.240 formulas do Calc
' dependem desses nomes, e o marcador do grafico ja aponta para F. Reaproveitar
' a estrutura entrega a funcionalidade sem tocar em nada disso.
'
' TRES CAMADAS EM UMA OPERACAO SO:
'   1. mDados.ExcluirLogico  -> muda o Status no banco (sai dos calculos)
'   2. mAuditoria.Auditar    -> Event Store, encadeado por hash
'   3. mLogDB.RegistrarLogDB -> tabela do banco, por origem
' As tres compartilham o ID_Auditoria. A escrita na Registros e apresentacao.

Public Const REG As String = "Registros"
Public Const REG_R0 As Long = 4
Public Const REG_RN As Long = 203
Public Const RG_NUM As Long = 1
Public Const RG_DATA As Long = 2
Public Const RG_ANALITO As Long = 3
Public Const RG_NIVEL As Long = 4
Public Const RG_RUN As Long = 5
Public Const RG_RESULT As Long = 6
Public Const RG_CALIB As Long = 9
Public Const RG_PARECER As Long = 10
Public Const RG_RESP As Long = 11
Public Const RG_TIPO As Long = 13

' Primeira linha livre da view, a partir de B4.
Public Function PrimeiraLinhaLivreReg() As Long
    Dim ws As Worksheet, i As Long
    Set ws = ThisWorkbook.Sheets(REG)
    For i = REG_R0 To REG_RN
        If Trim$(CStr(ws.Cells(i, RG_DATA).Value)) = "" And _
           Trim$(CStr(ws.Cells(i, RG_ANALITO).Value)) = "" Then
            PrimeiraLinhaLivreReg = i
            Exit Function
        End If
    Next i
    PrimeiraLinhaLivreReg = 0          ' cheia
End Function

' Registra UM resultado como nao conforme.
' Devolve o ID_Auditoria, ou "" se nao houve o que registrar.
Public Function MarcarNaoConforme(ByVal run As Long, ByVal nivel As Long, _
                                  ByVal analito As String, _
                                  ByVal tipoNC As String, _
                                  ByVal parecer As String) As String
    Dim ws As Worksheet, dados As Variant, i As Long, lin As Long
    Dim dtCorrida As Variant, lote As String, valor As Variant, stAntes As String
    Dim idEv As String, linReg As Long

    If Not ParecerValido(parecer) Then
        MarcarNaoConforme = ""
        Exit Function
    End If

    Set ws = ThisWorkbook.Sheets(BANCO)
    dados = CarregarDB()
    If IsEmpty(dados) Then MarcarNaoConforme = "": Exit Function

    ' localiza o registro no banco
    lin = 0
    For i = 1 To UBound(dados, 1)
        If CLng(dados(i, COL_RUN)) = run And CLng(dados(i, COL_NIVEL)) = nivel Then
            If UCase$(Trim$(CStr(dados(i, COL_ANALITO)))) = UCase$(Trim$(analito)) Then
                lin = BANCO_R0 + i - 1
                dtCorrida = dados(i, COL_DATA)
                lote = CStr(dados(i, COL_LOTE))
                valor = dados(i, COL_RESULT)
                stAntes = Trim$(CStr(dados(i, COL_STATUS)))
                Exit For
            End If
        End If
    Next i
    If lin = 0 Then MarcarNaoConforme = "": Exit Function

    ' 1. sai dos calculos, com o valor original preservado
    ws.Cells(lin, COL_STATUS).Value = tipoNC

    ' 2. Event Store
    idEv = Auditar(CAT_DADO, AC_EXCLUSAO, "mRegistros", _
                   run, dtCorrida, "", lote, nivel, analito, _
                   valor, valor, stAntes, tipoNC, tipoNC, parecer)

    ' 3. tabela do banco, por origem
    RegistrarLogDB ORIG_RESULTADOS, idEv, AC_EXCLUSAO, _
                   run, dtCorrida, nivel, analito, lote, valor, _
                   stAntes, tipoNC, parecer

    ' 4. apresentacao na aba Registros
    linReg = PrimeiraLinhaLivreReg()
    If linReg > 0 Then
        Dim wr As Worksheet, prot As Boolean
        Set wr = ThisWorkbook.Sheets(REG)
        prot = wr.ProtectContents
        If prot Then wr.Unprotect Password:="qcini2025"
        wr.Cells(linReg, RG_DATA).Value = dtCorrida
        wr.Cells(linReg, RG_DATA).NumberFormat = "dd/mm/yyyy"
        wr.Cells(linReg, RG_ANALITO).Value = analito
        wr.Cells(linReg, RG_NIVEL).Value = nivel
        wr.Cells(linReg, RG_RUN).Value = run
        wr.Cells(linReg, RG_RESULT).Value = valor
        wr.Cells(linReg, RG_CALIB).Value = "Nao"
        wr.Cells(linReg, RG_PARECER).Value = parecer
        wr.Cells(linReg, RG_RESP).Value = UsuarioSistema()
        wr.Cells(linReg, RG_TIPO).Value = tipoNC
        If prot Then wr.Protect Password:="qcini2025", UserInterfaceOnly:=True, _
                                DrawingObjects:=False, Contents:=True, Scenarios:=True
    End If

    MarcarNaoConforme = idEv
End Function

' Remove uma ocorrencia da VITRINE Registros. Nunca do banco.
'
' O analista pode ter classificado errado. Limpar a tela nao pode significar
' apagar evidencia: a linha do DB_Resultados fica onde esta, e a remocao vira
' evento proprio no Event Store e no LOG_Registros.
Public Function ExcluirRegistroNC(ByVal linha As Long, ByVal parecer As String) As String
    Dim wr As Worksheet, idEv As String, prot As Boolean
    Dim run As Long, nivel As Long, analito As String, valor As Variant
    Dim dtCorrida As Variant, tipoNC As String

    If Not ParecerValido(parecer) Then ExcluirRegistroNC = "": Exit Function
    If linha < REG_R0 Or linha > REG_RN Then ExcluirRegistroNC = "": Exit Function

    Set wr = ThisWorkbook.Sheets(REG)
    If Trim$(CStr(wr.Cells(linha, RG_ANALITO).Value)) = "" Then ExcluirRegistroNC = "": Exit Function

    dtCorrida = wr.Cells(linha, RG_DATA).Value
    analito = CStr(wr.Cells(linha, RG_ANALITO).Value)
    nivel = Val(wr.Cells(linha, RG_NIVEL).Value)
    run = Val(wr.Cells(linha, RG_RUN).Value)
    valor = wr.Cells(linha, RG_RESULT).Value
    tipoNC = CStr(wr.Cells(linha, RG_TIPO).Value)

    idEv = Auditar(CAT_DADO, "REGISTRO_NC_REMOVIDO", "mRegistros", _
                   run, dtCorrida, "", "", nivel, analito, _
                   valor, valor, tipoNC, tipoNC, _
                   "Remocao de ocorrencia da aba Registros", parecer)

    RegistrarLogDB ORIG_REGISTROS, idEv, "REGISTRO_NC_REMOVIDO", _
                   run, dtCorrida, nivel, analito, "", valor, _
                   tipoNC, tipoNC, parecer

    prot = wr.ProtectContents
    If prot Then wr.Unprotect Password:="qcini2025"
    wr.Range(wr.Cells(linha, RG_DATA), wr.Cells(linha, RG_TIPO)).ClearContents
    If prot Then wr.Protect Password:="qcini2025", UserInterfaceOnly:=True, _
                            DrawingObjects:=False, Contents:=True, Scenarios:=True

    ExcluirRegistroNC = idEv
End Function

' Ocorrencias visiveis, para popular o formulario de exclusao.
' Devolve "linha|data|analito|nivel|RUN|resultado|tipo" por item.
Public Function ListaRegistrosNC() As Collection
    Dim wr As Worksheet, i As Long, c As Collection, s As String
    Set c = New Collection
    Set wr = ThisWorkbook.Sheets(REG)
    For i = REG_R0 To REG_RN
        If Trim$(CStr(wr.Cells(i, RG_ANALITO).Value)) <> "" Then
            s = CStr(i) & "|"
            If IsDate(wr.Cells(i, RG_DATA).Value) Then
                s = s & Format$(wr.Cells(i, RG_DATA).Value, "dd/mm/yyyy")
            End If
            s = s & "|" & CStr(wr.Cells(i, RG_ANALITO).Value) & _
                "|" & CStr(wr.Cells(i, RG_NIVEL).Value) & _
                "|" & CStr(wr.Cells(i, RG_RUN).Value) & _
                "|" & CStr(wr.Cells(i, RG_RESULT).Value) & _
                "|" & CStr(wr.Cells(i, RG_TIPO).Value)
            c.Add s
        End If
    Next i
    Set ListaRegistrosNC = c
End Function

' Estados NAO ELEGIVEIS do Cfg_Status: sao as opcoes validas de tipo de NC.
' Lidos da tabela, nunca escritos em codigo (ADR-006).
Public Function TiposNaoConformidade() As Collection
    Dim ws As Worksheet, i As Long, c As Collection, nome As String
    Set c = New Collection
    Set ws = ThisWorkbook.Sheets(CFG)
    For i = CFG_R0 To CFG_RN
        nome = Trim$(CStr(ws.Cells(i, CFG_C_STATUS).Value))
        If nome <> "" Then
            If UCase$(Trim$(CStr(ws.Cells(i, CFG_C_ELEG).Value))) <> "SIM" Then c.Add nome
        End If
    Next i
    Set TiposNaoConformidade = c
End Function

' Resultados de uma corrida: alimenta o formulario sem digitacao.
' Devolve "analito|nivel|resultado|status" por item.
Public Function ResultadosDaCorrida(ByVal run As Long, ByVal nivel As Long) As Collection
    Dim dados As Variant, i As Long, c As Collection
    Set c = New Collection
    dados = CarregarDB()
    If IsEmpty(dados) Then Set ResultadosDaCorrida = c: Exit Function
    For i = 1 To UBound(dados, 1)
        If CLng(dados(i, COL_RUN)) = run And CLng(dados(i, COL_NIVEL)) = nivel Then
            If Trim$(CStr(dados(i, COL_ANALITO))) <> "" Then
                c.Add CStr(dados(i, COL_ANALITO)) & "|" & CStr(nivel) & "|" & _
                      CStr(dados(i, COL_RESULT)) & "|" & CStr(dados(i, COL_STATUS))
            End If
        End If
    Next i
    Set ResultadosDaCorrida = c
End Function

' Cabecalho da corrida, para o formulario preencher sozinho.
' Devolve "data|lote" ou "" se o RUN nao existir.
Public Function CabecalhoDaCorrida(ByVal run As Long) As String
    Dim dados As Variant, i As Long
    dados = CarregarDB()
    If IsEmpty(dados) Then CabecalhoDaCorrida = "": Exit Function
    For i = 1 To UBound(dados, 1)
        If CLng(dados(i, COL_RUN)) = run Then
            CabecalhoDaCorrida = Format$(dados(i, COL_DATA), "dd/mm/yyyy") & "|" & _
                                 CStr(dados(i, COL_LOTE))
            Exit Function
        End If
    Next i
    CabecalhoDaCorrida = ""
End Function

' ===================== PONTOS DE ENTRADA =====================
' Chamados pelos botoes das abas. Ficam aqui, e nao no modulo do formulario,
' para que o botao dependa de um nome estavel mesmo se o formulario for
' reconstruido pelo pipeline.

Public Sub AbrirFormNaoConforme()
    frmResultadoNaoConforme.Show
End Sub

Public Sub AbrirFormExcluirRegistroNC()
    frmExcluirRegistroNC.Show
End Sub
