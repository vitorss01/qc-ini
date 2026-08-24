Attribute VB_Name = "mBI"
Option Explicit
' ===== CAMADA DE DADOS PARA BI (ADR-026) =====
'
' O QUE ESTA ABA E, E O QUE ELA NAO E
'
' BI_Data NAO e uma copia da planilha. E uma TABELA FATO: uma linha por
' resultado de controle, granularidade (Lote, RUN, Nivel, Analito), com chaves
' estaveis que nao dependem de posicao de celula. O Power BI le daqui e de
' mais lugar nenhum.
'
' O Excel continua sendo a fonte operacional. Esta camada nao calcula nada que
' o QC_INI ja nao calcule -- ela DESNORMALIZA o que o motor produz para que o
' BI nao precise reimplementar a regra. Onde ha conta, a conta e a mesma, e o
' teste de reconciliacao (mBI.ReconciliarComCalc) prova isso a cada build.
'
' DE ONDE VEM CADA COISA
'
'   resultado, data, status, RUN, nivel, lote .. DB_Resultados (A:G)
'   media e DP alvo POR LOTE ................... LotesStore, bloco do lote
'   Fonte / ETp / CVTp ......................... Analitos S / T / U (ADR-028)
'   Bias (assinado) ............................ mCEQ.BiasEQ, do EP (ADR-030)
'   area e unidade ............................. Analitos (B, C)
'   Westgard ................................... produzido por mEstatistica.AvaliarWestgard
'
' ALVO POR LOTE, E NAO O ALVO DA TELA
'
' A aba Analitos mostra as metas do LOTE ATIVO -- e so dele. O historico dos
' demais lotes vive no LotesStore, em blocos de 40 linhas (mLotes.SalvarViewNoBloco).
' Ler o alvo da tela para um resultado de 2026 com o lote de 2027 carregado daria
' um Z errado, plausivel e silencioso. Aqui o alvo vem sempre do bloco do lote
' A QUE O RESULTADO PERTENCE.
'
' WESTGARD: UMA UNICA FONTE DE VERDADE
'
' A camada BI prepara as series por lote/analito e chama diretamente
' mEstatistica.AvaliarWestgard. Ela nao possui uma segunda implementacao das
' regras. Assim, 1_3s, 2_2s, R_4s, 4_1s, 10x e o alerta 1_2s usam exatamente o
' mesmo motor que alimenta Painel, Estatistica e Eventos_Westgard.
'
' O QUE NAO EXISTE E POR ISSO NAO ESTA AQUI
'
' Equipamento e Setor. A producao nao tem as abas Corridas/Cfg_Status, e nenhuma
' outra guarda esses campos. Coluna sem origem e coluna que alguem vai filtrar e
' obter conclusao errada.

Public Const BI_ABA As String = "BI_Data"
Public Const BI_TABELA As String = "tblBI_Fato"
Public Const BI_CAB As Long = 1
Public Const BI_R0 As Long = 2
Public Const BI_NCOL As Long = 84

Private Const LS_CAP As Long = 40          ' linhas por bloco no LotesStore
Private Const LS_C0 As Long = 3            ' coluna C = Analitos!E (Media N1)

Private Type BI_SYSTEMTIME
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
End Type

#If VBA7 Then
Private Declare PtrSafe Sub GetSystemTime Lib "kernel32" (ByRef lpSystemTime As BI_SYSTEMTIME)
#Else
Private Declare Sub GetSystemTime Lib "kernel32" (ByRef lpSystemTime As BI_SYSTEMTIME)
#End If

Private Function Cab() As Variant
    Cab = Array( _
        "ID_Result", "ID_Corrida", "Data", "Ano", "Mes", "Trimestre", "Competencia", _
        "ID_Analito", "Analito", "Area", "Unidade", _
        "ID_Lote", "Lote", "Nivel", "RUN", _
        "Resultado", "Status", "Ativo", _
        "Media_Alvo", "DP_Alvo", "Z", _
        "Lim_m3s", "Lim_m2s", "Lim_m1s", "Lim_p1s", "Lim_p2s", "Lim_p3s", _
        "CVtp_pct", "BIAStp_pct", "ETp_pct", _
        "W_1_3s", "W_2_2s", "W_R_4s", "W_4_1s", "W_10x", "A_1_2s", "Veredito", _
        "Produto", "WorkbookID", "VersaoContrato", "AtualizadoEmUTC", "FonteArquivo", _
        "ID_Result_Global", "ID_Corrida_Global", "ID_Lote_Global", "ID_Analito_Global", _
        "N_Observado", "Media_Observada", "DP_Observado", "CV_Observado_pct", _
        "Bias_Observado_pct", "ET_Observado_pct", "Sigma", _
        "Fonte_Especificacao", "ID_Especificacao", "Vigencia_Inicio", "Vigencia_Fim", _
        "Situacao_Especificacao", "Usuario_Atualizacao", "Tipo_Evento", _
        "Bias_Observado_abs_pct", "Classificacao_Sigma", _
        "Margem_ETp_pp", "Margem_ETp_pct", "Status_Margem_ETp", _
        "Provedor_EQA", "Ano_EQA", "Rodada_EQA", _
        "DPM_Teorico", "Yield_Teorico", _
        "Regra_Westgard_Recomendada", "N_Controle_Recomendado", _
        "RunSize_Max_Recomendado", "Frequencia_QC_Descricao", _
        "Cobertura_Motor_Westgard", "Referencia_Plano_QC", _
        "Sigma_Plano", "Nivel_Governante", "Classificacao_Sigma_Plano", _
        "Usar_1_3s", "Usar_2_2s", "Usar_R_4s", "Usar_4_1s", "Usar_8x")
End Function

Private Function AgoraUTC() As Date
    Dim st As BI_SYSTEMTIME
    GetSystemTime st
    AgoraUTC = DateSerial(st.wYear, st.wMonth, st.wDay) + _
               TimeSerial(st.wHour, st.wMinute, st.wSecond)
End Function

Private Function ValorNomeBI(ByVal nome As String) As String
    Dim nm As Name, s As String
    On Error Resume Next
    Set nm = ThisWorkbook.Names(nome)
    On Error GoTo 0
    If nm Is Nothing Then Exit Function
    On Error Resume Next
    ValorNomeBI = Trim$(CStr(nm.RefersToRange.Value))
    On Error GoTo 0
    If Len(ValorNomeBI) > 0 Then Exit Function
    s = nm.RefersTo
    If Left$(s, 1) = "=" Then s = Mid$(s, 2)
    If Left$(s, 1) = Chr$(34) And Right$(s, 1) = Chr$(34) Then _
        s = Mid$(s, 2, Len(s) - 2)
    ValorNomeBI = Trim$(s)
End Function

' Indice do lote no registro de lotes (mesma conta de mLotes.BlocoDoLote, que e
' Private). Devolve 0 quando o lote nao esta registrado.
Private Function BlocoDoLoteBI(ByVal loteCore As String) As Long
    Dim rng As Range, c As Range, k As Long
    If Len(Trim$(loteCore)) = 0 Then Exit Function
    On Error Resume Next
    Set rng = ThisWorkbook.Names("regLoteCol").RefersToRange
    On Error GoTo 0
    If rng Is Nothing Then Exit Function
    For Each c In rng
        k = k + 1
        If Trim$(CStr(c.Value)) = Trim$(loteCore) Then BlocoDoLoteBI = k: Exit Function
    Next c
End Function

' Media e DP do analito NAQUELE lote, no nivel pedido.
' Devolve False quando nao ha alvo utilizavel -- e o chamador deixa Z vazio, em
' vez de dividir por zero e publicar um numero.
Private Function AlvoDoLote(ByVal iBloco As Long, ByVal iAnalito As Long, _
                            ByVal nivel As Long, ByRef media As Double, _
                            ByRef dp As Double) As Boolean
    Dim ws As Worksheet, r As Long, cM As Long
    Dim vM As Variant, vD As Variant
    If iBloco < 1 Or iAnalito < 1 Or nivel < 1 Or nivel > 3 Then Exit Function
    Set ws = ThisWorkbook.Sheets("LotesStore")
    r = 2 + (iBloco - 1) * LS_CAP + (iAnalito - 1)
    cM = LS_C0 + (nivel - 1) * 2          ' C/D = N1, E/F = N2, G/H = N3
    vM = ws.Cells(r, cM).Value
    vD = ws.Cells(r, cM + 1).Value
    If Not IsNumeric(vM) Or Not IsNumeric(vD) Then Exit Function
    If Len(Trim$(CStr(vM))) = 0 Or Len(Trim$(CStr(vD))) = 0 Then Exit Function
    media = CDbl(vM)
    dp = CDbl(vD)
    AlvoDoLote = (dp > 0)
End Function

' Reconstroi BI_Data inteira. O(n) sobre o banco, uma escrita em bloco.
Public Sub AtualizarBIData()
    Dim ws As Worksheet, wsA As Worksheet
    Dim dados As Variant, i As Long, n As Long, ult As Long
    Dim saida() As Variant, k As Long
    Dim idxAnalito As Object, idxBloco As Object, especCV As Object
    Dim especBIAS As Object, especET As Object, especFonte As Object
    Dim especAno As Object, especID As Object, especSituacao As Object
    Dim especCVA As Object            ' CVTp lido da Analitos!U (ADR-028)
    Dim area As Object, unid As Object
    Dim zPorChave As Object, runsPorGrupo As Object, flagsPorChave As Object
    Dim prot As Boolean, protEstava As Boolean
    Dim produto As String, workbookID As String, versaoContrato As String
    Dim atualizadoUTC As Date, fonteArquivo As String, usuarioAtual As String

    Set wsA = ThisWorkbook.Sheets("Analitos")
    Set ws = GarantirAba()
    produto = ValorNomeBI("biProduto")
    workbookID = ValorNomeBI("biWorkbookID")
    versaoContrato = ValorNomeBI("biContratoVersao")
    If produto = "" Or workbookID = "" Or versaoContrato = "" Then
        Err.Raise vbObjectError + 580, "mBI.AtualizarBIData", _
                  "Identidade BI ausente. Execute preparar_contrato_bi antes de atualizar."
    End If
    atualizadoUTC = AgoraUTC()
    fonteArquivo = ThisWorkbook.Name
    usuarioAtual = UsuarioSistema()

    ' --- catalogo de analitos: indice na Analitos, area, unidade -----------
    Set idxAnalito = CreateObject("Scripting.Dictionary")
    Set area = CreateObject("Scripting.Dictionary")
    Set unid = CreateObject("Scripting.Dictionary")
    Set especET = CreateObject("Scripting.Dictionary")
    Set especFonte = CreateObject("Scripting.Dictionary")
    Set especAno = CreateObject("Scripting.Dictionary")
    Set especID = CreateObject("Scripting.Dictionary")
    Set especSituacao = CreateObject("Scripting.Dictionary")
    Set especCVA = CreateObject("Scripting.Dictionary")
    especCVA.CompareMode = 1
    idxAnalito.CompareMode = 1: area.CompareMode = 1: unid.CompareMode = 1
    especET.CompareMode = 1: especFonte.CompareMode = 1: especAno.CompareMode = 1
    especID.CompareMode = 1: especSituacao.CompareMode = 1
    For i = 4 To 43
        Dim nm As String
        nm = Trim$(CStr(wsA.Cells(i, 1).Value))
        If Len(nm) > 0 Then
            If Not idxAnalito.Exists(nm) Then
                idxAnalito.Add nm, i - 3                  ' 1..40, alinhado ao LotesStore
                area.Add nm, CStr(wsA.Cells(i, 2).Value)
                unid.Add nm, CStr(wsA.Cells(i, 3).Value)
                ' ADR-028: a especificacao vigente vem da Analitos, e so dela.
                '
                ' Estas colunas MUDARAM de lugar quando o bloco foi reorganizado,
                ' e o codigo antigo continuou lendo 18 e 17 -- que hoje sao
                ' "ETp VB" e "Desemp.". O BI publicava o NIVEL DE DESEMPENHO
                ' (OTI/DES/MIN) no campo Fonte_Especificacao e um ETp que nao era
                ' o em uso. Numero plausivel e errado, direto no painel do gestor.
                '
                '   S (19) = ETp fonte           CLIA | VB | FAB
                '   T (20) = ETp em uso final %  <- a meta oficial
                '   U (21) = CVTp%
                If IsNumeric(wsA.Cells(i, 20).Value) Then especET.Add nm, wsA.Cells(i, 20).Value
                If IsNumeric(wsA.Cells(i, 21).Value) Then especCVA.Add nm, wsA.Cells(i, 21).Value
                especFonte.Add nm, CStr(wsA.Cells(i, 19).Value)
                especAno.Add nm, ValorNomeBI("espAno")
                especID.Add nm, ""
                especSituacao.Add nm, "ANALITOS_VIGENTE"
            End If
        End If
    Next i

    ' --- Eng_Especificacoes saiu de cena (ADR-028) -----------------------
    '
    ' Este bloco relia CVtp / BIAStp / ETp do Eng_Especificacoes e SOBRESCREVIA
    ' o que tinha acabado de ler da Analitos. Era a segunda fonte de verdade que
    ' o ADR-027 eliminou no Painel e na Estatistica, sobrevivendo justamente na
    ' camada que o gestor ve.
    '
    ' BIAStp continua vazio de proposito: nem a Analitos nem o historico por ano
    ' informam bias permitido. Preencher com zero seria inventar meta.
    Set especCV = especCVA
    Set especBIAS = CreateObject("Scripting.Dictionary")
    especBIAS.CompareMode = 1

    ' --- cache de blocos por lote ----------------------------------------
    Set idxBloco = CreateObject("Scripting.Dictionary")
    idxBloco.CompareMode = 1

    ' --- banco ------------------------------------------------------------
    ult = UltimaLinhaBanco()
    If ult < BANCO_R0 Then
        LimparCorpo ws
        Exit Sub
    End If
    dados = ThisWorkbook.Sheets(BANCO).Range( _
        ThisWorkbook.Sheets(BANCO).Cells(BANCO_R0, COL_RUN), _
        ThisWorkbook.Sheets(BANCO).Cells(ult, COL_STATUS)).Value
    n = UBound(dados, 1)
    ReDim saida(1 To n, 1 To BI_NCOL)

    ' Os filtros de EQA sao os mesmos para o lote inteiro: ler uma vez fora
    ' do laco, e nao 90.000 vezes dentro dele.
    Dim provedorEQA As String, anoEQA As String, rodadaEQA As String
    provedorEQA = ValorNomeBI("eqProvedor")
    anoEQA = ValorNomeBI("eqAnoEP")
    rodadaEQA = ValorNomeBI("eqRodada")

    ' PASSO 1 -- Z de cada resultado, guardado por (lote|run|analito|nivel).
    ' Westgard 2_2s e R_4s comparam NIVEIS DA MESMA CORRIDA, entao o Z do outro
    ' nivel precisa existir antes de julgar qualquer linha. Por isso duas
    ' passadas, e nao uma.
    Set zPorChave = CreateObject("Scripting.Dictionary")
    zPorChave.CompareMode = 1
    Set runsPorGrupo = CreateObject("Scripting.Dictionary")
    runsPorGrupo.CompareMode = 1

    For i = 1 To n
        Dim an As String, lote As String, nucleo As String
        Dim nv As Long, run As Long, st As String
        an = Trim$(CStr(dados(i, COL_ANALITO)))
        If Len(an) = 0 Then GoTo proxima1
        st = Trim$(CStr(dados(i, COL_STATUS)))
        If st <> ST_ATIVO Then GoTo proxima1
        If Not IsNumeric(dados(i, COL_RESULT)) Then GoTo proxima1
        lote = Trim$(CStr(dados(i, COL_LOTE)))
        If Len(lote) < 6 Then GoTo proxima1
        nucleo = NucleoLote(lote)
        nv = CLng(Val(CStr(dados(i, COL_NIVEL))))
        run = CLng(Val(CStr(dados(i, COL_RUN))))

        Dim iB As Long
        If idxBloco.Exists(nucleo) Then
            iB = idxBloco(nucleo)
        Else
            iB = BlocoDoLoteBI(nucleo)
            idxBloco.Add nucleo, iB
        End If

        Dim md As Double, sd As Double
        If Not idxAnalito.Exists(an) Then GoTo proxima1
        If AlvoDoLote(iB, idxAnalito(an), nv, md, sd) Then
            zPorChave(nucleo & "|" & run & "|" & UCase$(an) & "|" & nv) = _
                (CDbl(dados(i, COL_RESULT)) - md) / sd
            Dim grupo As String, runsGrupo As Object
            grupo = nucleo & "|" & UCase$(an)
            If Not runsPorGrupo.Exists(grupo) Then
                Set runsGrupo = CreateObject("Scripting.Dictionary")
                runsPorGrupo.Add grupo, runsGrupo
            Else
                Set runsGrupo = runsPorGrupo(grupo)
            End If
            If Not runsGrupo.Exists(CStr(run)) Then runsGrupo.Add CStr(run), run
        End If
proxima1:
    Next i

    ' O proprio motor avalia as series. mBI apenas indexa a saida por registro.
    Set flagsPorChave = FlagsDoMotor(zPorChave, runsPorGrupo)

    ' PASSO 2 -- monta a linha da tabela fato
    For i = 1 To n
        Dim an2 As String, lt As String, nuc As String, st2 As String
        Dim nv2 As Long, run2 As Long
        an2 = Trim$(CStr(dados(i, COL_ANALITO)))
        If Len(an2) = 0 Then GoTo proxima2
        lt = Trim$(CStr(dados(i, COL_LOTE)))
        nuc = ""
        If Len(lt) >= 6 Then nuc = NucleoLote(lt)
        st2 = Trim$(CStr(dados(i, COL_STATUS)))
        nv2 = CLng(Val(CStr(dados(i, COL_NIVEL))))
        run2 = CLng(Val(CStr(dados(i, COL_RUN))))

        k = k + 1
        saida(k, 1) = nuc & "|" & run2 & "|" & nv2 & "|" & UCase$(an2)   ' ID_Result
        saida(k, 2) = nuc & "|" & run2                                   ' ID_Corrida

        Dim dt As Date, temData As Boolean
        temData = IsDate(dados(i, COL_DATA))
        If temData Then
            dt = CDate(dados(i, COL_DATA))
            saida(k, 3) = Int(CDbl(dt))
            saida(k, 4) = Year(dt)
            saida(k, 5) = Month(dt)
            saida(k, 6) = Int((Month(dt) - 1) / 3) + 1
            saida(k, 7) = Format$(dt, "yyyy-mm")
        End If

        saida(k, 8) = UCase$(an2)
        saida(k, 9) = an2
        If area.Exists(an2) Then saida(k, 10) = area(an2)
        If unid.Exists(an2) Then saida(k, 11) = unid(an2)
        saida(k, 12) = nuc
        saida(k, 13) = lt
        saida(k, 14) = nv2
        saida(k, 15) = run2
        If IsNumeric(dados(i, COL_RESULT)) Then saida(k, 16) = CDbl(dados(i, COL_RESULT))
        saida(k, 17) = st2
        saida(k, 18) = IIf(st2 = ST_ATIVO, 1, 0)

        Dim md2 As Double, sd2 As Double, temAlvo As Boolean
        Dim iB2 As Long
        iB2 = 0
        If idxBloco.Exists(nuc) Then iB2 = idxBloco(nuc)
        temAlvo = False
        If idxAnalito.Exists(an2) Then
            temAlvo = AlvoDoLote(iB2, idxAnalito(an2), nv2, md2, sd2)
        End If
        If temAlvo Then
            saida(k, 19) = md2
            saida(k, 20) = sd2
            saida(k, 22) = md2 - 3 * sd2
            saida(k, 23) = md2 - 2 * sd2
            saida(k, 24) = md2 - sd2
            saida(k, 25) = md2 + sd2
            saida(k, 26) = md2 + 2 * sd2
            saida(k, 27) = md2 + 3 * sd2
        End If

        Dim chave As String, z As Variant, temZ As Boolean
        chave = nuc & "|" & run2 & "|" & UCase$(an2) & "|" & nv2
        temZ = zPorChave.Exists(chave)
        If temZ Then
            z = zPorChave(chave)
            saida(k, 21) = z
        End If

        If especCV.Exists(an2) Then saida(k, 28) = especCV(an2)
        If especBIAS.Exists(an2) Then saida(k, 29) = especBIAS(an2)
        If especET.Exists(an2) Then saida(k, 30) = especET(an2)

        ' --- Westgard: transporte da saida do motor -----------------------
        Dim wf As Variant
        If flagsPorChave.Exists(chave) Then
            wf = flagsPorChave(chave)
            saida(k, 31) = wf(0)
            saida(k, 32) = wf(1)
            saida(k, 33) = wf(2)
            saida(k, 34) = wf(3)
            saida(k, 35) = wf(4)
            saida(k, 36) = wf(5)
        Else
            saida(k, 31) = "": saida(k, 32) = "": saida(k, 33) = ""
            saida(k, 34) = "": saida(k, 35) = "": saida(k, 36) = ""
        End If
        If Len(Trim$(CStr(saida(k, 16)))) = 0 Then
            saida(k, 37) = ""
        ElseIf flagsPorChave.Exists(chave) Then
            saida(k, 37) = wf(6)
        Else
            saida(k, 37) = "NAO AVALIADO"
        End If

        ' --- identidade global e contrato versionado ----------------------
        saida(k, 38) = produto
        saida(k, 39) = workbookID
        saida(k, 40) = versaoContrato
        saida(k, 41) = atualizadoUTC
        saida(k, 42) = fonteArquivo
        saida(k, 43) = workbookID & "|RESULT|" & CStr(saida(k, 1))
        saida(k, 44) = workbookID & "|RUN|" & CStr(saida(k, 2))
        saida(k, 45) = workbookID & "|LOTE|" & nuc
        saida(k, 46) = produto & "|ANALITO|" & UCase$(an2)

        ' --- desempenho calculado pelo mesmo motor do Excel ---------------
        Dim eb As Variant, nObs As Long, mediaObs As Double, dpObs As Double
        Dim cvObs As Double, biasObs As Double, etObs As Double, sigmaObs As Double
        eb = mEstatistica.EstatBasica(an2, nv2, nuc)
        nObs = CLng(eb(0))
        saida(k, 47) = nObs
        If nObs > 0 Then
            mediaObs = CDbl(eb(1))
            saida(k, 48) = mediaObs
        End If
        If nObs > 1 Then
            dpObs = CDbl(eb(2))
            saida(k, 49) = dpObs
            If mediaObs <> 0 Then
                cvObs = mEstatistica.CalcularCV(dpObs, mediaObs)
                saida(k, 50) = cvObs
            End If
        End If
        ' --- BIAS: vem do ENSAIO DE PROFICIENCIA, nao do alvo do lote (ADR-030)
        '
        ' Antes era CalcularBias(mediaObs, alvoDoLote): a deriva do controle
        ' INTERNO contra a media atribuida ao proprio lote. Isso nao e erro
        ' sistematico -- erro sistematico se mede contra um valor externo e
        ' independente, o consenso do grupo no EP. O painel do gestor publicava
        ' um Sigma construido sobre o bias errado.
        '
        ' A coluna publica o bias ASSINADO, que informa a direcao do desvio.
        ' CalcularErroTotal e CalcularSigma ja aplicam Abs() internamente, entao
        ' a magnitude entra certa nas metricas sem destruir o sinal na exibicao.
        '
        ' "SEM EP" (texto) quando nao ha rodada utilizavel: sem bias nao ha ET
        ' nem Sigma. Zero seria exatidao inventada.
        Dim biasEP As Variant, anoBI As Variant
        anoBI = saida(k, 4)          ' o ano do proprio resultado
        biasEP = mCEQ.BiasEQ(an2, anoBI, "SIGNED")
        If IsNumeric(biasEP) And nObs > 0 Then
            biasObs = CDbl(biasEP)
            saida(k, 51) = biasObs
            If nObs > 1 And mediaObs <> 0 Then
                etObs = mEstatistica.CalcularErroTotal(cvObs, biasObs)
                saida(k, 52) = etObs
                If especET.Exists(an2) Then
                    If IsNumeric(especET(an2)) And CDbl(especET(an2)) <> 0 And cvObs <> 0 Then
                        sigmaObs = mEstatistica.CalcularSigma(CDbl(especET(an2)), biasObs, cvObs)
                        saida(k, 53) = sigmaObs
                    End If
                End If
            End If
        End If

        If especFonte.Exists(an2) Then saida(k, 54) = especFonte(an2)
        If especID.Exists(an2) Then saida(k, 55) = especID(an2)
        If especAno.Exists(an2) Then
            If IsNumeric(especAno(an2)) And CLng(Val(CStr(especAno(an2)))) > 0 Then
                saida(k, 56) = DateSerial(CLng(especAno(an2)), 1, 1)
                saida(k, 57) = DateSerial(CLng(especAno(an2)), 12, 31)
            End If
        End If
        If especSituacao.Exists(an2) Then saida(k, 58) = especSituacao(an2)
        saida(k, 59) = usuarioAtual
        saida(k, 60) = "RESULTADO_CQI"

        ' --- ADR-033: magnitude do bias, classificacao e orcamento de erro ---
        '
        ' A coluna 51 publica o bias ASSINADO, que informa a direcao do desvio.
        ' Uma medida de BI que faca AVERAGE sobre ela cancela desvios opostos e
        ' devolve um vies falso perto de zero -- foi por isso que a
        ' consolidacao de rodadas em mCEQ passou a somar magnitudes. A coluna
        ' 61 carrega o |bias| justamente para as medidas agregarem o que deve
        ' ser agregado.
        '
        ' As classificacoes NAO sao recalculadas aqui: sao as mesmas funcoes que
        ' a Estatistica e o Painel chamam (mQualidade). Se as faixas mudarem,
        ' mudam nos tres ao mesmo tempo.
        If IsNumeric(biasEP) And nObs > 0 Then saida(k, 61) = Abs(CDbl(biasEP))
        saida(k, 62) = mQualidade.ClassificarSigma(saida(k, 53))
        If especET.Exists(an2) Then
            saida(k, 63) = mQualidade.MargemETp(especET(an2), saida(k, 52))
            saida(k, 64) = mQualidade.MargemETpPct(especET(an2), saida(k, 52))
        End If
        saida(k, 65) = mQualidade.ClassificarMargem(saida(k, 64))

        ' --- ADR-035: do Sigma ate a decisao operacional -------------------
        '
        ' O BI recebe a CADEIA INTEIRA, e nao so o Sigma. Um numero Sigma
        ' isolado nao decide nada; o que decide e quantas regras rodar,
        ' quantos controles medir e quantos pacientes podem passar entre
        ' eventos de CQ. Recalcular isso em DAX criaria uma segunda escada
        ' que divergiria da planilha no primeiro ajuste de faixa.
        '
        ' Os tres primeiros campos dizem de QUAL recorte de EQA veio o bias:
        ' sem isso, um relatorio filtrado por CAP/C-B e outro por Controllab
        ' pareceriam a mesma coisa.
        saida(k, 66) = provedorEQA
        saida(k, 67) = anoEQA
        saida(k, 68) = rodadaEQA
        ' As colunas 69..76 e 77..84 NAO sao preenchidas aqui: dependem do
        ' Sigma do PIOR nivel do analito, que so se conhece depois de varrer
        ' todas as linhas. Ver PreencherPlanoDoPiorNivel, abaixo do laco.
proxima2:
    Next i

    ' --- ADR-040: o plano de CQ vem do PIOR nivel, nao da propria linha ---
    PreencherPlanoDoPiorNivel saida, k

    ' --- escrita, com a protecao tratada (mesma licao do ADR-025) ---------
    prot = ws.ProtectContents
    On Error GoTo restaura
    If prot Then ws.Unprotect Password:="qcini2025"

    LimparCorpo ws
    If k > 0 Then
        Dim bloco() As Variant, r2 As Long, c2 As Long
        ReDim bloco(1 To k, 1 To BI_NCOL)
        For r2 = 1 To k
            For c2 = 1 To BI_NCOL
                bloco(r2, c2) = saida(r2, c2)
            Next c2
        Next r2
        ws.Range(ws.Cells(BI_R0, 1), ws.Cells(BI_R0 + k - 1, BI_NCOL)).Value = bloco
        ws.Range(ws.Cells(BI_R0, 3), ws.Cells(BI_R0 + k - 1, 3)).NumberFormatLocal = "aaaa-mm-dd;@"
        ws.Range(ws.Cells(BI_R0, 13), ws.Cells(BI_R0 + k - 1, 13)).NumberFormat = "@"
        ws.Range(ws.Cells(BI_R0, 41), ws.Cells(BI_R0 + k - 1, 41)).NumberFormatLocal = "aaaa-mm-dd hh:mm:ss;@"
        ws.Range(ws.Cells(BI_R0, 56), ws.Cells(BI_R0 + k - 1, 57)).NumberFormatLocal = "aaaa-mm-dd;@"
    End If
    AjustarTabela ws, k

restaura:
    Dim nErr As Long, sErr As String
    nErr = Err.Number: sErr = Err.Description
    On Error Resume Next
    If prot Then
        ws.Protect Password:="qcini2025", UserInterfaceOnly:=True, _
                   DrawingObjects:=False, Contents:=True, Scenarios:=True
    End If
    On Error GoTo 0
    If nErr <> 0 Then Err.Raise nErr, "mBI.AtualizarBIData", sErr
End Sub

' Avalia cada serie lote/analito usando a API publica do motor. A funcao
' devolve um dicionario por chave de resultado com:
'   Array(1_3s, 2_2s, R_4s, 4_1s, 10x, alerta_1_2s, veredito)
Private Function FlagsDoMotor(ByVal zPorChave As Object, _
                              ByVal runsPorGrupo As Object) As Object
    Dim flags As Object, runsGrupo As Object
    Dim g As Variant, p As Variant, chave As String
    Dim runs() As Long, nRun As Long, i As Long, j As Long, nv As Long
    Dim tmp As Long, nlvLocal As Long, rejeitado As Boolean
    Dim z() As Double, temDado() As Boolean
    Dim r13 As Variant, r22 As Variant, rR4 As Variant
    Dim r41 As Variant, r10 As Variant, a12 As Variant

    Set flags = CreateObject("Scripting.Dictionary")
    flags.CompareMode = 1
    nlvLocal = mEstatistica.NLV

    For Each g In runsPorGrupo.Keys
        Set runsGrupo = runsPorGrupo(g)
        nRun = runsGrupo.Count
        If nRun < 1 Then GoTo proximoGrupo

        ReDim runs(1 To nRun)
        i = 0
        Dim rk As Variant
        For Each rk In runsGrupo.Keys
            i = i + 1
            runs(i) = CLng(runsGrupo(rk))
        Next rk
        For i = 1 To nRun - 1
            For j = i + 1 To nRun
                If runs(j) < runs(i) Then
                    tmp = runs(i): runs(i) = runs(j): runs(j) = tmp
                End If
            Next j
        Next i

        ReDim z(0 To nlvLocal - 1, 1 To nRun)
        ReDim temDado(0 To nlvLocal - 1, 1 To nRun)
        ReDim r13(0 To nlvLocal - 1, 1 To nRun)
        ReDim r22(0 To nlvLocal - 1, 1 To nRun)
        ReDim rR4(0 To nlvLocal - 1, 1 To nRun)
        ReDim r41(0 To nlvLocal - 1, 1 To nRun)
        ReDim r10(0 To nlvLocal - 1, 1 To nRun)
        ReDim a12(0 To nlvLocal - 1, 1 To nRun)

        p = Split(CStr(g), "|")
        For i = 1 To nRun
            For nv = 1 To nlvLocal
                chave = CStr(p(0)) & "|" & runs(i) & "|" & CStr(p(1)) & "|" & nv
                If zPorChave.Exists(chave) Then
                    z(nv - 1, i) = CDbl(zPorChave(chave))
                    temDado(nv - 1, i) = True
                End If
            Next nv
        Next i

        mEstatistica.AvaliarWestgard z, temDado, nRun, r13, r22, rR4, r41, r10, a12

        For i = 1 To nRun
            For nv = 1 To nlvLocal
                If temDado(nv - 1, i) Then
                    rejeitado = (r13(nv - 1, i) = 1 Or r22(nv - 1, i) = 1 Or _
                                  rR4(nv - 1, i) = 1 Or r41(nv - 1, i) = 1 Or _
                                  r10(nv - 1, i) = 1)
                    chave = CStr(p(0)) & "|" & runs(i) & "|" & CStr(p(1)) & "|" & nv
                    flags(chave) = Array(IIf(r13(nv - 1, i) = 1, 1, 0), _
                                         IIf(r22(nv - 1, i) = 1, 1, 0), _
                                         IIf(rR4(nv - 1, i) = 1, 1, 0), _
                                         IIf(r41(nv - 1, i) = 1, 1, 0), _
                                         IIf(r10(nv - 1, i) = 1, 1, 0), _
                                         IIf(a12(nv - 1, i) = 1, 1, 0), _
                                         IIf(rejeitado, "REJEITADO", "OK"))
                End If
            Next nv
        Next i
proximoGrupo:
    Next g

    Set FlagsDoMotor = flags
End Function

Private Function GarantirAba() As Worksheet
    Dim ws As Worksheet, i As Long, cb As Variant
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(BI_ABA)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = BI_ABA
    End If
    cb = Cab()
    For i = 0 To UBound(cb)
        ws.Cells(BI_CAB, i + 1).Value = cb(i)
    Next i
    ws.Rows(BI_CAB).Font.Bold = True
    Set GarantirAba = ws
End Function

Private Sub LimparCorpo(ByVal ws As Worksheet)
    Dim ult As Long
    ult = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If ult >= BI_R0 Then
        ws.Range(ws.Cells(BI_R0, 1), ws.Cells(ult, BI_NCOL)).ClearContents
    End If
End Sub

' Tabela ESTRUTURADA, nao faixa de celulas.
'
' O Power Query referencia a tabela pelo NOME. Uma faixa A1:AH9999 quebraria a
' cada linha a mais ou a menos; o ListObject acompanha sozinho e o M nao muda.
Private Sub AjustarTabela(ByVal ws As Worksheet, ByVal nLinhas As Long)
    Dim lo As ListObject, alvo As Range
    Dim ultima As Long
    ultima = BI_R0 + IIf(nLinhas > 0, nLinhas - 1, 0)
    Set alvo = ws.Range(ws.Cells(BI_CAB, 1), ws.Cells(ultima, BI_NCOL))
    On Error Resume Next
    Set lo = ws.ListObjects(BI_TABELA)
    On Error GoTo 0
    If lo Is Nothing Then
        Set lo = ws.ListObjects.Add(xlSrcRange, alvo, , xlYes)
        lo.Name = BI_TABELA
        lo.TableStyle = "TableStyleMedium2"
    Else
        lo.Resize alvo
    End If
End Sub

' Reconciliacao integral entre a fonte operacional e a tabela fato. Compara
' cardinalidade, identidade, resultado e status de TODAS as linhas, nao apenas
' o analito selecionado no Painel. Devolve "comparados|divergencias|primeira".
Public Function ReconciliarBancoBI() As String
    Dim wsB As Worksheet, dados As Variant, idx As Object, globais As Object
    Dim i As Long, ultB As Long, comp As Long, div As Long, prim As String
    Dim chave As String, nuc As String, an As String, lin As Long
    Set wsB = ThisWorkbook.Sheets(BI_ABA)
    Set idx = CreateObject("Scripting.Dictionary")
    Set globais = CreateObject("Scripting.Dictionary")
    idx.CompareMode = 1: globais.CompareMode = 1

    ultB = wsB.Cells(wsB.Rows.Count, 1).End(xlUp).Row
    For i = BI_R0 To ultB
        chave = Trim$(CStr(wsB.Cells(i, 1).Value))
        If chave <> "" Then
            If idx.Exists(chave) Then
                div = div + 1
                If prim = "" Then prim = "ID_Result duplicado no BI: " & chave
            Else
                idx.Add chave, i
            End If
        End If
        chave = Trim$(CStr(wsB.Cells(i, 43).Value))
        If chave <> "" Then
            If globais.Exists(chave) Then
                div = div + 1
                If prim = "" Then prim = "ID_Result_Global duplicado: " & chave
            Else
                globais.Add chave, True
            End If
        End If
    Next i

    dados = CarregarDB()
    If Not IsEmpty(dados) Then
        For i = 1 To UBound(dados, 1)
            an = Trim$(CStr(dados(i, COL_ANALITO)))
            If an <> "" Then
                nuc = NucleoLote(CStr(dados(i, COL_LOTE)))
                chave = nuc & "|" & CLng(Val(CStr(dados(i, COL_RUN)))) & "|" & _
                        CLng(Val(CStr(dados(i, COL_NIVEL)))) & "|" & UCase$(an)
                If Not idx.Exists(chave) Then
                    div = div + 1
                    If prim = "" Then prim = "registro ausente no BI: " & chave
                Else
                    comp = comp + 1
                    lin = CLng(idx(chave))
                    If CStr(wsB.Cells(lin, 17).Value) <> CStr(dados(i, COL_STATUS)) Then
                        div = div + 1
                        If prim = "" Then prim = "status diverge: " & chave
                    End If
                    If Not MesmoValorBI(wsB.Cells(lin, 16).Value, dados(i, COL_RESULT)) Then
                        div = div + 1
                        If prim = "" Then prim = "resultado diverge: " & chave
                    End If
                End If
            End If
        Next i
    End If
    If idx.Count <> comp Then
        div = div + Abs(idx.Count - comp)
        If prim = "" Then prim = "cardinalidade BI=" & idx.Count & " banco=" & comp
    End If
    ReconciliarBancoBI = CStr(comp) & "|" & CStr(div) & "|" & prim
End Function

Private Function MesmoValorBI(ByVal a As Variant, ByVal b As Variant) As Boolean
    If IsNumeric(a) And IsNumeric(b) And Trim$(CStr(a)) <> "" And Trim$(CStr(b)) <> "" Then
        MesmoValorBI = (Abs(CDbl(a) - CDbl(b)) < 0.0000001)
    Else
        MesmoValorBI = (Trim$(CStr(a)) = Trim$(CStr(b)))
    End If
End Function

' ---------------------------------------------------------------------------
' RECONCILIACAO: a camada BI tem de concordar com o motor, nao aproximar-se dele
'
' Compara, para o analito e o lote que a aba Calc esta exibindo, o Z e o veredito
' de Westgard linha a linha. Devolve "comparados|divergencias|primeira".
'
' Qualquer divergencia e defeito: as duas contas saem do mesmo dado e da mesma
' regra. Se divergirem, uma das duas esta errada -- e num sistema de CQ nao da
' para saber qual sem investigar.
Public Function ReconciliarComCalc() As String
    Dim wsC As Worksheet, wsB As Worksheet
    Dim selAn As String, lote As String
    Dim i As Long, ultB As Long, comp As Long, div As Long, prim As String
    Dim idx As Object, chave As String

    Set wsC = ThisWorkbook.Sheets("Calc")
    Set wsB = ThisWorkbook.Sheets(BI_ABA)
    selAn = Trim$(CStr(ThisWorkbook.Names("selAnalito").RefersToRange.Value))
    lote = Trim$(CStr(ThisWorkbook.Names("loteAtivo").RefersToRange.Value))
    If Len(selAn) = 0 Then ReconciliarComCalc = "0|0|sem analito selecionado": Exit Function

    ' indexa a BI_Data por RUN|NIVEL para o analito e lote em tela
    Set idx = CreateObject("Scripting.Dictionary")
    idx.CompareMode = 1
    ultB = wsB.Cells(wsB.Rows.Count, 1).End(xlUp).Row
    For i = BI_R0 To ultB
        If UCase$(Trim$(CStr(wsB.Cells(i, 9).Value))) = UCase$(selAn) Then
            If Trim$(CStr(wsB.Cells(i, 12).Value)) = lote Then
                chave = CStr(wsB.Cells(i, 15).Value) & "|" & CStr(wsB.Cells(i, 14).Value)
                If Not idx.Exists(chave) Then idx.Add chave, i
            End If
        End If
    Next i

' Calc: cada nivel ocupa 22 colunas; Z, cinco flags e veredito sao comparados.
    Dim nv As Long, colZ As Long, colV As Long, colF As Long, f As Long
    For i = 3 To 182
        If Len(Trim$(CStr(wsC.Cells(i, 2).Value))) = 0 Then GoTo prox
        For nv = 1 To mEstatistica.NLV
            colZ = 7 + (nv - 1) * 22
            colF = 11 + (nv - 1) * 22
            colV = 16 + (nv - 1) * 22
            If Len(Trim$(CStr(wsC.Cells(i, colZ).Value))) = 0 Then GoTo proxNivel
            chave = CStr(wsC.Cells(i, 2).Value) & "|" & CStr(nv)
            If Not idx.Exists(chave) Then
                div = div + 1
                If Len(prim) = 0 Then prim = "chave ausente no BI: RUN " & wsC.Cells(i, 2).Value & " N" & nv
                GoTo proxNivel
            End If
            comp = comp + 1
            Dim zC As Double, zB As Double, vC As String, vB As String
            zC = CDbl(wsC.Cells(i, colZ).Value)
            zB = CDbl(wsB.Cells(idx(chave), 21).Value)
            vC = Trim$(CStr(wsC.Cells(i, colV).Value))
            vB = Trim$(CStr(wsB.Cells(idx(chave), 37).Value))
            If Abs(zC - zB) > 0.000001 Or vC <> vB Then
                div = div + 1
                If Len(prim) = 0 Then
                    prim = "RUN " & wsC.Cells(i, 2).Value & " N" & nv & _
                           " Z calc=" & Format$(zC, "0.000000") & " bi=" & Format$(zB, "0.000000") & _
                           " veredito calc=" & vC & " bi=" & vB
                End If
            End If
            For f = 0 To 4
                If CLng(Val(CStr(wsC.Cells(i, colF + f).Value))) <> _
                   CLng(Val(CStr(wsB.Cells(idx(chave), 31 + f).Value))) Then
                    div = div + 1
                    If Len(prim) = 0 Then
                        prim = "flag " & (f + 1) & " diverge: RUN " & _
                               wsC.Cells(i, 2).Value & " N" & nv
                    End If
                End If
            Next f
proxNivel:
        Next nv
prox:
    Next i
    ReconciliarComCalc = CStr(comp) & "|" & CStr(div) & "|" & prim
End Function


' --- ADR-040: o plano de CQ vem do PIOR nivel do analito -----------------
'
' O DEFEITO QUE ISTO CORRIGE
'
' Ate aqui o mBI preenchia regras, N e run size a partir do Sigma da PROPRIA
' LINHA, ou seja, do Sigma daquele nivel. O Lactato tem Sigma 6,99 no nivel 1
' e 1,83 no nivel 2: as linhas do nivel 1 publicavam "1_3s, N=2, run size
' 1000" -- o CQ mais leve que existe, num analito cujo nivel 2 nao sustenta
' nem 3 Sigma. Mil pacientes entre eventos de controle.
'
' O ADR-038 ja tinha resolvido isso NA PLANILHA (Sigma_Plano = MIN entre os
' niveis validos). O contrato do BI ficou para tras e publicava a versao
' errada -- Excel e Power BI diriam coisas diferentes sobre o mesmo analito.
'
' POR QUE AQUI E NAO EM DAX
'
' A regra e uma decisao de negocio e o motor e a unica camada de calculo
' (ADR-019). Reimplementar o MIN entre niveis no Power BI criaria uma segunda
' copia que divergiria no primeiro ajuste de faixa -- exatamente o que o
' ADR-027 passou uma sessao inteira eliminando.
'
' O AGRUPAMENTO
'
' Chave = (ID_Analito, ID_Lote). Colapsa os niveis, que e o que "o pior nivel
' governa" quer dizer, e preserva o lote: dois lotes do mesmo analito sao
' materiais diferentes e nao devem herdar o plano um do outro.
Private Sub PreencherPlanoDoPiorNivel(ByRef saida As Variant, ByVal n As Long)
    If n <= 0 Then Exit Sub

    Dim pior As Object, nivelPior As Object
    Set pior = CreateObject("Scripting.Dictionary")
    pior.CompareMode = 1
    Set nivelPior = CreateObject("Scripting.Dictionary")
    nivelPior.CompareMode = 1

    Dim k As Long, ch As String, s As Double, sp As Double

    For k = 1 To n
        If SigmaValidoBI(saida(k, 53)) Then
            ch = CStr(saida(k, 8)) & "|" & CStr(saida(k, 12))
            s = CDbl(saida(k, 53))
            If Not pior.Exists(ch) Then
                pior(ch) = s
                nivelPior(ch) = saida(k, 14)
            ElseIf s < CDbl(pior(ch)) Then
                pior(ch) = s
                nivelPior(ch) = saida(k, 14)
            End If
        End If
    Next k

    For k = 1 To n
        ch = CStr(saida(k, 8)) & "|" & CStr(saida(k, 12))
        If pior.Exists(ch) Then
            sp = CDbl(pior(ch))
            saida(k, 77) = sp
            saida(k, 78) = "Nivel " & CStr(nivelPior(ch))
            saida(k, 79) = mQualidade.ClassificarSigma(sp)
            saida(k, 69) = mPlanoQC.DPMdoSigma(sp)
            saida(k, 70) = mPlanoQC.RendimentoDoSigma(sp)
            saida(k, 71) = mPlanoQC.PlanoQC(sp, "REGRAS")
            saida(k, 72) = mPlanoQC.PlanoQC(sp, "N")
            saida(k, 73) = mPlanoQC.PlanoQC(sp, "RUNSIZE")
            saida(k, 74) = mPlanoQC.PlanoQC(sp, "FREQUENCIA")
            saida(k, 75) = mPlanoQC.CoberturaWestgard(sp)
            saida(k, 76) = mPlanoQC.PlanoQC(sp, "REFERENCIA")
            saida(k, 80) = mPlanoQC.RegraNoPlano(sp, "1_3s")
            saida(k, 81) = mPlanoQC.RegraNoPlano(sp, "2_2s")
            saida(k, 82) = mPlanoQC.RegraNoPlano(sp, "R_4s")
            saida(k, 83) = mPlanoQC.RegraNoPlano(sp, "4_1s")
            saida(k, 84) = mPlanoQC.RegraNoPlano(sp, "8x")
        Else
            ' Sem Sigma valido em nenhum nivel: o plano fica VAZIO e o rotulo
            ' diz SEM DADOS. Nao ha faixa a aplicar, e escolher a mais rigorosa
            ' "por seguranca" inventaria uma conclusao que o dado nao sustenta.
            saida(k, 78) = "SEM DADOS"
            saida(k, 79) = "SEM DADOS"
            saida(k, 80) = False
            saida(k, 81) = False
            saida(k, 82) = False
            saida(k, 83) = False
            saida(k, 84) = False
        End If
    Next k
End Sub


' Este Sigma pode governar um plano de CQ?
'
' Zero e a armadilha central: um Sigma exatamente zero nao existe num metodo
' que produziu resultado -- ele nasce de ETp, bias e CV. Zero na coluna e
' ausencia de dado disfarcada, e aceita-lo faria o analito cair na faixa
' "reavaliar metodo" por falta de informacao, que e conclusao diferente de
' "o metodo e ruim". Texto, erro de celula e estouro caem pelo mesmo motivo.
Private Function SigmaValidoBI(ByVal v As Variant) As Boolean
    SigmaValidoBI = False
    If IsEmpty(v) Or IsNull(v) Then Exit Function
    If IsError(v) Then Exit Function
    If Not IsNumeric(v) Then Exit Function

    Dim d As Double
    On Error GoTo fora
    d = CDbl(v)
    On Error GoTo 0

    If d <= 0 Then Exit Function
    If d > 1E+15 Then Exit Function      ' infinito / estouro
    SigmaValidoBI = True
    Exit Function
fora:
End Function
