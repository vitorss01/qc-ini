# -*- coding: utf-8 -*-
"""patch_vba.py -- Etapa 1 (multi-lote): gera os modulos VBA alterados.

Entrada : fontes extraidas do QC_Bioquimica.xlsm de producao (oletools), em
          <entrada>/  (um arquivo por modulo, nome = nome do modulo).
Saida   : etapa1_multilote/src/  (mEstatistica.bas, mEstatPeriodo.bas, mBI.bas,
          Planilha7.cls, Planilha3.cls, EstaPastaDeTrabalho.cls). O mLotes.bas
          novo e escrito a mao e ja esta em src/.

Cada troca e feita por ANCORA EXATA e conferida: se a ancora nao aparece o
numero esperado de vezes, o script para. Patch que "passa" sem ter casado e
exatamente o defeito que o ADR-047 descreveu (a correcao que nunca embarcou).
"""
import os
import re
import sys

AQUI = os.path.dirname(os.path.abspath(__file__))
SRC = os.path.join(AQUI, 'src')


def ler(d, nome):
    with open(os.path.join(d, nome), encoding='utf-8', newline='') as f:
        return f.read().replace('\r\n', '\n')


def gravar(nome, txt):
    with open(os.path.join(SRC, nome), 'w', encoding='utf-8', newline='') as f:
        f.write(txt.replace('\n', '\r\n'))


def troca(txt, velho, novo, vezes=1, rotulo=''):
    n = txt.count(velho)
    if n != vezes:
        raise SystemExit(f'ANCORA [{rotulo}] casou {n}x, esperado {vezes}x:\n{velho[:200]}')
    return txt.replace(velho, novo)


def troca_bloco(txt, inicio, fim, novo, rotulo=''):
    """Substitui do inicio (inclusive) ate fim (exclusive). Ambos unicos."""
    a = txt.count(inicio)
    if a != 1:
        raise SystemExit(f'ANCORA-INICIO [{rotulo}] casou {a}x')
    i = txt.index(inicio)
    j = txt.index(fim, i)
    return txt[:i] + novo + txt[j:]


# ---------------------------------------------------------------------------
def patch_mestatistica(t):
    # --- AlvoAnalito: le o alvo DO LOTE pela chave; ETp pelo cabecalho -------
    t = troca_bloco(t,
        "' Media/DP configurados (valores-alvo do lote) para o nivel.\n",
        "' ============================ ESTATISTICA BASICA",
        """' Media/DP-alvo DO LOTE pedido, para o nivel (ADR-049).
'
' ANTES: lia Analitos!E:J -- a tela do lote carregado --, qualquer que fosse o
' lote dos resultados avaliados. Com o Painel num ano de outro lote, o z de um
' lote era calculado com a media/DP de outro. Numero plausivel, errado, calado.
' AGORA: a fonte e mLotes.ParametrosLote(lote, analito, nivel), pela chave.
' Devolve False quando o lote nao tem parametro valido: quem chama NAO calcula z.
'
' ETp: vem da coluna "ETp em uso final %", achada pelo CABECALHO. O indice fixo
' 18 apontava para "ETp VB" desde que as colunas de especificacao mudaram.
Public Function AlvoAnalito(ByVal analito As String, ByVal nivel As Long, _
                            ByRef alvoMedia As Double, ByRef alvoDP As Double, _
                            ByRef etp As Double, ByVal lote As String) As Boolean
    Dim r As Long, cE As Long
    alvoMedia = 0: alvoDP = 0: etp = 0
    r = LinhaAnalito(analito)
    If r > 0 Then
        cE = ColunaEtpEmUso()
        If cE > 0 And cE <= UBound(mAnal, 2) Then
            If IsNumeric(mAnal(r, cE)) And Not IsEmpty(mAnal(r, cE)) Then etp = CDbl(mAnal(r, cE))
        End If
    End If
    AlvoAnalito = mLotes.ParametrosLote(lote, analito, nivel, alvoMedia, alvoDP)
End Function

' Coluna de "ETp em uso final %" na aba Analitos, pelo rotulo da linha 3.
Private Function ColunaEtpEmUso() As Long
    Static c As Long
    Dim ws As Worksheet, j As Long, s As String
    If c > 0 Then ColunaEtpEmUso = c: Exit Function
    Set ws = ThisWorkbook.Sheets("Analitos")
    For j = 1 To 20
        s = UCase$(Replace(Trim$(CStr(ws.Cells(3, j).Value)), "  ", " "))
        If Left$(s, 10) = "ETP EM USO" Then c = j: Exit For
    Next j
    If c = 0 Then c = 20                     ' T, layout de 08/2026
    ColunaEtpEmUso = c
End Function

' Ordem CRONOLOGICA das corridas: data, e o RUN so desempata (ADR-049).
'
' O RUN e unico por (data, lote), mas NAO e cronologico: dado historico
' importado depois ganha RUN maior. No arquivo real, marco/2025 do lote 8974
' tem RUN 151..168 e junho/2026 tem RUN ~114 -- ordenar por RUN punha 2025
' DEPOIS de 2026, e as regras de sequencia (4_1s, 8x) liam uma serie fora do
' tempo. Um so comparador para o motor, os eventos e o BI.
Public Function CorridaAntes(ByVal d1 As Double, ByVal r1 As Long, _
                             ByVal d2 As Double, ByVal r2 As Long) As Boolean
    If d1 < d2 Then
        CorridaAntes = True
    ElseIf d1 = d2 Then
        CorridaAntes = (r1 < r2)
    End If
End Function

Public Sub OrdenarCorridas(ByRef runs() As Long, ByRef dts() As Double, ByVal n As Long)
    Dim a As Long, b As Long, tr As Long, td As Double
    For a = 2 To n
        tr = runs(a): td = dts(a): b = a - 1
        Do While b >= 1
            If Not CorridaAntes(td, tr, dts(b), runs(b)) Then Exit Do
            runs(b + 1) = runs(b): dts(b + 1) = dts(b): b = b - 1
        Loop
        runs(b + 1) = tr: dts(b + 1) = td
    Next a
End Sub

""", 'AlvoAnalito')

    # --- InvalidarCache tambem invalida a store de lotes ---------------------
    t = troca(t, "Public Sub InvalidarCache()\n    Set mCache = Nothing\n",
              "Public Sub InvalidarCache()\n    mLotes.InvalidarLotes\n    Set mCache = Nothing\n",
              rotulo='InvalidarCache')

    # --- AtualizarCalc: lote do Painel, janela cronologica, parametro ausente --
    t = troca(t, """    GarantirDB
    Dim anoPainel As Long
    anoPainel = CLng(Val(ThisWorkbook.Sheets("Painel").Range("I3").Value))
    lote = LoteDoAnoOuAtivo(analito, anoPainel)
""", """    GarantirDB
    ' ADR-049: o lote e o do PAINEL (loteSel), escolhido pelo usuario. A versao
    ' anterior deduzia o lote pelo ano (LoteDoAnoOuAtivo) enquanto o Calc
    ' filtrava pelo lote ativo: grafico e veredicto podiam falar de lotes
    ' diferentes na mesma tela.
    lote = mLotes.LotePainel()
""", rotulo='AtualizarCalc lote')

    t = troca(t, """    ws.Range("C1").Value = analito
    ws.Range("E1").Value = lote
    ws.Range("G1").Value = Now
    ws.Range("I1").Value = 0
    If analito = "" Or IsEmpty(mDB) Then GoTo restaura
""", """    ws.Range("C1").Value = analito
    ws.Range("E1").NumberFormat = "@"       ' lote "010" nao pode virar 10
    ws.Range("E1").Value = lote
    ws.Range("G1").Value = Now
    ws.Range("I1").Value = 0
    ws.Range("K1").Value = 0
    ws.Range("M1").Value = 0
    ws.Range("O1").Value = ""
    ws.Range("Q1").Value = ""
    If analito = "" Or IsEmpty(mDB) Then GoTo restaura
    ws.Range("O1").Value = LotesNoPeriodo(analito)
""", rotulo='AtualizarCalc cabecalho')

    t = troca_bloco(t,
        "    ' ---- descobrir corridas (RUN) elegiveis do analito no lote ----\n",
        "    Set ordem = CreateObject(\"Scripting.Dictionary\")\n",
        """    ' ---- descobrir TODAS as corridas (RUN) elegiveis do analito no lote ----
    ' ADR-049: sem teto na descoberta. O teto antigo (NK) cortava pela ordem do
    ' banco, e num lote com mais de 180 corridas as MAIS NOVAS sumiam do grafico
    ' e do Westgard, sem aviso.
    Set seen = CreateObject("Scripting.Dictionary")
    Dim nTodas As Long, runsT() As Long, dtsT() As Double
    ReDim runsT(1 To UBound(mDB, 1)): ReDim dtsT(1 To UBound(mDB, 1))
    For i = 1 To UBound(mDB, 1)
        If StrComp(Trim$(CStr(mDB(i, COL_ANALITO))), analito, 1) = 0 Then
            If NucleoLote(CStr(mDB(i, COL_LOTE))) = lote Then
                If EhElegivel(mDB(i, COL_STATUS)) Then
                    If Not seen.Exists(CStr(mDB(i, COL_RUN))) Then
                        nTodas = nTodas + 1
                        runsT(nTodas) = CLng(Val(mDB(i, COL_RUN)))
                        If IsDate(mDB(i, COL_DATA)) Then dtsT(nTodas) = CDbl(CDate(mDB(i, COL_DATA)))
                        seen.Add CStr(mDB(i, COL_RUN)), nTodas
                    End If
                End If
            End If
        End If
    Next i
    ws.Range("K1").Value = nTodas
    If nTodas = 0 Then GoTo restaura
    OrdenarCorridas runsT, dtsT, nTodas

    ' ---- janela: as NK corridas mais recentes ate o fim do periodo ----
    ' O contexto ANTERIOR ao periodo fica (as regras de sequencia precisam dele);
    ' o que passa do fim do periodo sai. Se o proprio periodo tiver mais de NK
    ' corridas, cortam-se as mais antigas dele e o numero vai para M1, que o
    ' Painel mostra.
    Dim fimJan As Long, iniJan As Long, nCort As Long, vAte As Variant
    fimJan = nTodas
    On Error Resume Next
    vAte = ThisWorkbook.Names("filtroAte").RefersToRange.Value
    On Error GoTo restaura
    If IsDate(vAte) Then
        Do While fimJan >= 1
            If Int(dtsT(fimJan)) <= Int(CDbl(CDate(vAte))) Then Exit Do
            fimJan = fimJan - 1
        Loop
    End If
    If fimJan < 1 Then GoTo restaura
    iniJan = fimJan - NK + 1
    If iniJan < 1 Then iniJan = 1
    For i = 1 To iniJan - 1
        If PassaFiltro(dtsT(i)) Then nCort = nCort + 1
    Next i
    ws.Range("M1").Value = nCort
    nRun = fimJan - iniJan + 1
    ReDim runs(1 To NK): ReDim dts(1 To NK)
    For i = 1 To nRun
        runs(i) = runsT(iniJan + i - 1)
        dts(i) = dtsT(iniJan + i - 1)
    Next i

""", 'AtualizarCalc descoberta')

    t = troca(t, """    ' ---- alvos e z-scores ----
    ReDim alvoM(0 To NLV - 1): ReDim alvoS(0 To NLV - 1)
    For t = 0 To NLV - 1
        AlvoAnalito analito, t + 1, alvoM(t), alvoS(t), etp
        For i = 1 To nRun
            If temDado(t, i) Then z(t, i) = CalcularZ(valor(t, i), alvoM(t), alvoS(t))
        Next i
    Next t

    ' ---- Westgard ----
    ReDim r13(0 To NLV - 1, 1 To nRun): ReDim r2sMulti(0 To NLV - 1, 1 To nRun)
    ReDim rR4(0 To NLV - 1, 1 To nRun): ReDim r1sMulti(0 To NLV - 1, 1 To nRun)
    ReDim rSeq(0 To NLV - 1, 1 To nRun): ReDim a12(0 To NLV - 1, 1 To nRun)
    AvaliarWestgard z, temDado, nRun, r13, r2sMulti, rR4, r1sMulti, rSeq, a12
""", """    ' ---- alvos e z-scores ----
    ' Nivel sem media/DP do lote NAO entra no Westgard: z = 0 seria lido como
    ' "no alvo" e o grafico sairia verde sobre um lote nunca configurado.
    Dim parOK() As Boolean, temAval() As Boolean, stPar As String
    ReDim alvoM(0 To NLV - 1): ReDim alvoS(0 To NLV - 1): ReDim parOK(0 To NLV - 1)
    ReDim temAval(0 To NLV - 1, 1 To nRun)
    For t = 0 To NLV - 1
        parOK(t) = AlvoAnalito(analito, t + 1, alvoM(t), alvoS(t), etp, lote)
        stPar = stPar & IIf(t > 0, ";", "") & "N" & (t + 1) & "=" & IIf(parOK(t), "OK", "SEM")
        For i = 1 To nRun
            If temDado(t, i) And parOK(t) Then
                z(t, i) = CalcularZ(valor(t, i), alvoM(t), alvoS(t))
                temAval(t, i) = True
            End If
        Next i
    Next t
    ws.Range("Q1").Value = stPar

    ' ---- Westgard ----
    ReDim r13(0 To NLV - 1, 1 To nRun): ReDim r2sMulti(0 To NLV - 1, 1 To nRun)
    ReDim rR4(0 To NLV - 1, 1 To nRun): ReDim r1sMulti(0 To NLV - 1, 1 To nRun)
    ReDim rSeq(0 To NLV - 1, 1 To nRun): ReDim a12(0 To NLV - 1, 1 To nRun)
    AvaliarWestgard z, temAval, nRun, r13, r2sMulti, rR4, r1sMulti, rSeq, a12
""", rotulo='AtualizarCalc alvos')

    t = troca(t, """        For i = 1 To nRun
            If temDado(t, i) Then
                outLvl(i, 1) = IIf(r13(t, i) = 1, 1, 0)""", """        For i = 1 To nRun
            If temDado(t, i) And Not temAval(t, i) Then
                ' resultado existe, parametro do lote nao: sem veredicto
                outLvl(i, 1) = 0: outLvl(i, 2) = 0: outLvl(i, 3) = 0
                outLvl(i, 4) = 0: outLvl(i, 5) = 0: outLvl(i, 6) = 0
                outLvl(i, 7) = "SEM PARAMETROS"
            ElseIf temDado(t, i) Then
                outLvl(i, 1) = IIf(r13(t, i) = 1, 1, 0)""", rotulo='AtualizarCalc publicar')

    # (a ordenacao antiga por RUN ficava entre a descoberta e 'Set ordem' e
    #  saiu junto com o bloco de descoberta acima)

    # --- AtualizarPainelEng: alvo do lote em analise --------------------------
    t = troca(t, "        AlvoAnalito analito, t + 1, alvoM, alvoS, etp\n        bias = CalcularBias(media, alvoM)\n        et = CalcularErroTotal(cv, bias)\n",
              "        AlvoAnalito analito, t + 1, alvoM, alvoS, etp, mLotes.LotePainel()\n        bias = CalcularBias(media, alvoM)\n        et = CalcularErroTotal(cv, bias)\n",
              rotulo='AtualizarPainelEng')

    # --- AtualizarEstatisticaAba: alvo do lote do filtro ----------------------
    t = troca(t, "                AlvoAnalito analito, t, alvoM, alvoS, etp\n",
              "                ' lote vazio = todos os lotes: nao ha UM alvo, entao nao ha bias\n"
              "                AlvoAnalito analito, t, alvoM, alvoS, etp, loteF\n",
              rotulo='AtualizarEstatisticaAba')

    # --- AnalisarViolacoes: cronologico + alvo do lote ------------------------
    t = troca(t, """        Do While j >= 1
            If rs(j) <= tr Then Exit Do
            rs(j + 1) = rs(j): ys(j + 1) = ys(j): ds(j + 1) = ds(j): j = j - 1
        Loop""", """        Do While j >= 1
            If Not CorridaAntes(td2, tr, ds(j), rs(j)) Then Exit Do
            rs(j + 1) = rs(j): ys(j + 1) = ys(j): ds(j + 1) = ds(j): j = j - 1
        Loop""", rotulo='AnalisarViolacoes ordem')
    t = troca(t, "    AlvoAnalito analito, nivel, alvoM, alvoS, etp\n",
              "    If Not AlvoAnalito(analito, nivel, alvoM, alvoS, etp, loteCore) Then Exit Function\n",
              rotulo='AnalisarViolacoes alvo')
    t = troca(t, "    a = AnalisarViolacoes(analito, nivel, LoteAtivoCore())\n",
              "    a = AnalisarViolacoes(analito, nivel, mLotes.LotePainel())\n", rotulo='DetalheViolacao')

    # --- RegistrarEventosWestgard: lote do Painel, cronologico, sem parametro -
    t = troca(t, "    If IsEmpty(mDB) Then GoTo restaura\n    lote = LoteAtivoCore()\n",
              "    If IsEmpty(mDB) Then GoTo restaura\n    lote = mLotes.LotePainel()\n"
              "    ws.Range(\"K2\").Value = \"Lote:\"\n    ws.Range(\"L2\").NumberFormat = \"@\"\n    ws.Range(\"L2\").Value = lote\n",
              rotulo='Eventos lote')
    t = troca(t, """        Dim runs() As Long, nRun As Long, vistos As Object, item As Variant
        Set vistos = CreateObject("Scripting.Dictionary")
        ReDim runs(1 To col.Count)
        nRun = 0
        For j = 1 To col.Count
            item = col(j)
            If Not vistos.Exists(CStr(item(1))) Then
                vistos.Add CStr(item(1)), 1
                nRun = nRun + 1
                runs(nRun) = CLng(item(1))
            End If
        Next j
        If nRun > 0 Then
            Dim tr As Long
            For j = 2 To nRun
                tr = runs(j): k = j - 1
                Do While k >= 1
                    If runs(k) <= tr Then Exit Do
                    runs(k + 1) = runs(k): k = k - 1
                Loop
                runs(k + 1) = tr
            Next j""", """        Dim runs() As Long, nRun As Long, vistos As Object, item As Variant
        Dim dtRun() As Double
        Set vistos = CreateObject("Scripting.Dictionary")
        ReDim runs(1 To col.Count): ReDim dtRun(1 To col.Count)
        nRun = 0
        For j = 1 To col.Count
            item = col(j)
            If Not vistos.Exists(CStr(item(1))) Then
                vistos.Add CStr(item(1)), 1
                nRun = nRun + 1
                runs(nRun) = CLng(item(1))
                dtRun(nRun) = CDbl(item(3))
            End If
        Next j
        If nRun > 0 Then
            OrdenarCorridas runs, dtRun, nRun   ' cronologico (ADR-049)""", rotulo='Eventos ordem')
    t = troca(t, """            Dim alvoM() As Double, alvoS() As Double
            ReDim alvoM(0 To NLV - 1): ReDim alvoS(0 To NLV - 1)
            Dim mm As Double, ss As Double, ee As Double
            For t = 0 To NLV - 1
                AlvoAnalito analitoN, t + 1, mm, ss, ee
                alvoM(t) = mm: alvoS(t) = ss
            Next t
""", """            Dim alvoM() As Double, alvoS() As Double, alvoOK() As Boolean
            ReDim alvoM(0 To NLV - 1): ReDim alvoS(0 To NLV - 1): ReDim alvoOK(0 To NLV - 1)
            Dim mm As Double, ss As Double, ee As Double
            For t = 0 To NLV - 1
                alvoOK(t) = AlvoAnalito(analitoN, t + 1, mm, ss, ee, lote)
                alvoM(t) = mm: alvoS(t) = ss
            Next t
""", rotulo='Eventos alvos')
    t = troca(t, """                If Not td(t, i) Then           ' repetido no mesmo (nivel,RUN): fica o primeiro
                    td(t, i) = True
                    vl(t, i) = CDbl(item(2))
                    dtv(t, i) = CDbl(item(3))
                    If alvoS(t) > 0 Then zz(t, i) = CalcularZ(vl(t, i), alvoM(t), alvoS(t))
                End If""", """                ' repetido no mesmo (nivel,RUN): fica o primeiro. Nivel sem
                ' parametro do lote nao entra: z = 0 nao e "no alvo" (ADR-049).
                If Not td(t, i) And alvoOK(t) Then
                    td(t, i) = True
                    vl(t, i) = CDbl(item(2))
                    dtv(t, i) = CDbl(item(3))
                    zz(t, i) = CalcularZ(vl(t, i), alvoM(t), alvoS(t))
                End If""", rotulo='Eventos z')

    # --- Recalculo do analito: UM recalculo no fim, nao um por escrita ---------
    t = troca(t, """Public Sub RecalcularAnalitoAtual()
    Application.ScreenUpdating = False
    AtualizarCalc
    RegistrarEventosWestgard
    AtualizarPainelEng
    AtualizarEixos
    Application.ScreenUpdating = True
End Sub""", """' ADR-049: o motor grava em blocos (Eng_Saida, Eventos_Westgard). Com o calculo
' em automatico, CADA bloco disparava o recalculo das formulas dependentes --
' medido em ~11 s por recalculo no arquivo real, pago uma vez por escrita.
' Trocar de lote ou de analito virava minutos. Agora: calculo manual durante o
' motor, UM recalculo no fim, e o modo anterior restaurado tambem no erro. O
' resultado e o mesmo: nenhuma rotina do motor le formula que ele mesmo altere.
Public Sub RecalcularAnalitoAtual()
    Dim calcAntes As Long, nE As Long, sE As String
    calcAntes = Application.Calculation
    Application.ScreenUpdating = False
    On Error GoTo fim
    Application.Calculation = xlCalculationManual
    AtualizarCalc
    RegistrarEventosWestgard
    AtualizarPainelEng
fim:
    nE = Err.Number: sE = Err.Description
    On Error Resume Next
    Application.Calculation = calcAntes
    If calcAntes = xlCalculationAutomatic Then Application.Calculate
    AtualizarEixos
    Application.ScreenUpdating = True
    On Error GoTo 0
    If nE <> 0 Then Err.Raise nE, "mEstatistica.RecalcularAnalitoAtual", sE
End Sub""", rotulo='RecalcularAnalitoAtual')
    t = troca(t, """Public Sub AtualizarEstatistica()
    Application.ScreenUpdating = False
    InvalidarCache
    AtualizarCalc
    RegistrarEventosWestgard
    AtualizarPainelEng
    AtualizarEstatisticaAba
    AtualizarEixos
    Application.ScreenUpdating = True
End Sub""", """Public Sub AtualizarEstatistica()
    Dim calcAntes As Long, nE As Long, sE As String
    calcAntes = Application.Calculation
    Application.ScreenUpdating = False
    On Error GoTo fim
    Application.Calculation = xlCalculationManual     ' ver RecalcularAnalitoAtual
    InvalidarCache
    AtualizarCalc
    RegistrarEventosWestgard
    AtualizarPainelEng
    AtualizarEstatisticaAba
fim:
    nE = Err.Number: sE = Err.Description
    On Error Resume Next
    Application.Calculation = calcAntes
    If calcAntes = xlCalculationAutomatic Then Application.Calculate
    AtualizarEixos
    Application.ScreenUpdating = True
    On Error GoTo 0
    If nE <> 0 Then Err.Raise nE, "mEstatistica.AtualizarEstatistica", sE
End Sub""", rotulo='AtualizarEstatistica')

    # --- PassaFiltro por DIA (a hora gravada nao decide o periodo) ------------
    t = troca(t, """    If IsDate(de) Then If d < CDate(de) Then PassaFiltro = False: Exit Function
    If IsDate(ate) Then If d > CDate(ate) Then PassaFiltro = False: Exit Function""",
    """    ' ADR-049: por DIA. Resultado gravado com hora (a importacao de 2025 veio
    ' com 03:00) ficava fora do ultimo dia do periodo: 29/05 03:00 > 29/05 00:00.
    ' mEstatPeriodo ja truncava; o Painel nao.
    If IsDate(de) Then If Int(dserial) < Int(CDbl(CDate(de))) Then PassaFiltro = False: Exit Function
    If IsDate(ate) Then If Int(dserial) > Int(CDbl(CDate(ate))) Then PassaFiltro = False: Exit Function""",
    rotulo='PassaFiltro dia')

    # --- LoteDoAnoOuAtivo sai: o lote e escolhido, nao deduzido ---------------
    t = troca_bloco(t, "Public Function LoteDoAnoOuAtivo(",
                    "' Grava o provedor de EQA",
                    "' LoteDoAnoOuAtivo removida (ADR-049): o lote em analise e escolhido no\n"
                    "' Painel (loteSel), nao deduzido do ano.\n\n\n\n", 'LoteDoAnoOuAtivo')

    # --- LotesNoPeriodo: dica quando o lote escolhido nao tem dado no periodo --
    t = t.rstrip('\n') + """

' Lotes com resultado elegivel do analito dentro do filtro do Painel, com o
' numero de corridas: "8973 (39) · 8974 (18)". Vai para Eng_Saida!O1 e o
' Painel mostra quando o lote escolhido nao tem corrida no periodo -- em vez
' de trocar o lote sozinho, como o LoteDoAnoOuAtivo fazia.
Public Function LotesNoPeriodo(ByVal analito As String) As String
    Dim i As Long, nl As String, k As String, d As Object, runs As Object, s As String, x As Variant
    GarantirDB
    If IsEmpty(mDB) Then Exit Function
    Set d = CreateObject("Scripting.Dictionary")
    Set runs = CreateObject("Scripting.Dictionary")
    For i = 1 To UBound(mDB, 1)
        If StrComp(Trim$(CStr(mDB(i, COL_ANALITO))), analito, 1) = 0 Then
            If EhElegivel(mDB(i, COL_STATUS)) And IsDate(mDB(i, COL_DATA)) Then
                If PassaFiltro(CDbl(CDate(mDB(i, COL_DATA)))) Then
                    nl = NucleoLote(CStr(mDB(i, COL_LOTE)))
                    k = nl & "|" & CStr(mDB(i, COL_RUN))
                    If Not runs.Exists(k) Then
                        runs.Add k, 1
                        d(nl) = d(nl) + 1
                    End If
                End If
            End If
        End If
    Next i
    For Each x In d.Keys
        s = s & IIf(Len(s) > 0, " " & ChrW(183) & " ", "") & x & " (" & d(x) & ")"
    Next x
    LotesNoPeriodo = s
End Function

' Spinner do Painel e De/Ate: o motor tem de acompanhar o analito e o periodo
' em tela. Antes o spinner so reajustava o eixo e o Eng_Saida ficava no
' analito anterior -- o Calc agora se recusa a plotar um motor de outro
' analito/lote, entao quem muda a tela precisa refazer o motor.
Public Sub PainelMudou()
    On Error Resume Next
    GarantirMotorDoPainel
    AtualizarEixos
End Sub

' Painel ativado: se o motor ficou para tras (outro analito ou outro lote),
' refaz antes de mostrar.
Public Sub GarantirMotorDoPainel()
    Dim eng As Worksheet
    On Error Resume Next
    Set eng = ThisWorkbook.Sheets("Eng_Saida")
    If StrComp(Trim$(CStr(eng.Range("C1").Value)), _
               Trim$(CStr(ThisWorkbook.Names("selAnalito").RefersToRange.Value)), 1) <> 0 _
       Or Trim$(CStr(eng.Range("E1").Value)) <> mLotes.LotePainel() Then
        RecalcularAnalitoAtual
    End If
End Sub
"""
    return t


def patch_mestatperiodo(t):
    t = troca(t, "            EstatPeriodo = BiasContraAlvo(analito, nivel, media)\n",
              "            EstatPeriodo = BiasContraAlvo(analito, nivel, media, Trim$(CStr(lote)))\n",
              rotulo='EP bias')
    t = troca(t, "            bias = BiasContraAlvo(analito, nivel, media)\n",
              "            bias = BiasContraAlvo(analito, nivel, media, Trim$(CStr(lote)))\n",
              rotulo='EP et')
    t = troca_bloco(t, "' Bias % contra o alvo do lote, lido de LotesStore.\n",
                    "' ---------------------------------------------------------------------------\n' LIMITE VINDO DO MOTOR",
                    """' Bias % contra o alvo DO LOTE do filtro (ADR-049).
'
' Lote vazio = estatistica de TODOS os lotes juntos: nao existe UM alvo para
' isso, entao nao ha bias. Antes, o alvo era sempre o do lote ativo, e o bias
' de um recorte de outro lote saia contra a media errada.
'
' Devolve Empty quando nao ha alvo -- e nao zero. Zero seria lido como
' "exatidao perfeita" num analito que nunca teve alvo cadastrado.
Public Function BiasContraAlvo(ByVal analito As String, ByVal nivel As Variant, _
                               ByVal mediaObs As Double, _
                               Optional ByVal lote As String = "") As Variant
    Dim alvo As Variant
    If Len(Trim$(lote)) = 0 Then Exit Function
    alvo = AlvoDoLote(analito, nivel, lote)
    If Not IsNumeric(alvo) Then Exit Function
    If IsEmpty(alvo) Then Exit Function
    If CDbl(alvo) = 0 Then Exit Function
    BiasContraAlvo = (mediaObs - CDbl(alvo)) / CDbl(alvo) * 100#
End Function

' Media-alvo do lote para (analito, nivel). A leitura e UMA so no sistema:
' mLotes.ParametrosLote, pela chave LOTE|ANALITO. As duas implementacoes que
' existiam (varredura por idx aqui, aritmetica de bloco no mBI) concordavam
' por coincidencia -- a divida registrada no ADR-045 fecha aqui.
' lote omitido = lote em analise no Painel.
Public Function AlvoDoLote(ByVal analito As String, ByVal nivel As Variant, _
                           Optional ByVal lote As String = "") As Variant
    Dim m As Double, s As Double
    If Len(Trim$(lote)) = 0 Then lote = mLotes.LotePainel()
    If mLotes.ParametrosLote(lote, analito, CLng(Val(CStr(nivel))), m, s) Then AlvoDoLote = m
End Function

""", 'EP AlvoDoLote')
    return t


def patch_mbi(t):
    t = troca(t, """        Dim iB As Long
        If idxBloco.Exists(nucleo) Then
            iB = idxBloco(nucleo)
        Else
            iB = BlocoDoLoteBI(nucleo)
            idxBloco.Add nucleo, iB
        End If

        Dim md As Double, sd As Double
        If Not idxAnalito.Exists(an) Then GoTo proxima1
        If AlvoDoLote(iB, idxAnalito(an), nv, md, sd) Then""", """        Dim md As Double, sd As Double
        If Not idxAnalito.Exists(an) Then GoTo proxima1
        ' alvo do lote A QUE O RESULTADO PERTENCE, pela chave (ADR-049)
        If mLotes.ParametrosLote(nucleo, an, nv, md, sd) Then""", rotulo='BI passo1')
    t = troca(t, "            If Not runsGrupo.Exists(CStr(run)) Then runsGrupo.Add CStr(run), run\n",
              "            ' guarda a DATA da corrida: a serie do motor e cronologica (ADR-049)\n"
              "            If Not runsGrupo.Exists(CStr(run)) Then\n"
              "                If IsDate(dados(i, COL_DATA)) Then\n"
              "                    runsGrupo.Add CStr(run), CDbl(CDate(dados(i, COL_DATA)))\n"
              "                Else\n"
              "                    runsGrupo.Add CStr(run), 0#\n"
              "                End If\n"
              "            End If\n", rotulo='BI runsGrupo')
    t = troca(t, """        Dim iB2 As Long
        iB2 = 0
        If idxBloco.Exists(nuc) Then iB2 = idxBloco(nuc)
        temAlvo = False
        If idxAnalito.Exists(an2) Then
            temAlvo = AlvoDoLote(iB2, idxAnalito(an2), nv2, md2, sd2)
        End If""", """        temAlvo = False
        If idxAnalito.Exists(an2) Then
            temAlvo = mLotes.ParametrosLote(nuc, an2, nv2, md2, sd2)
        End If""", rotulo='BI passo2')
    t = troca(t, """        ReDim runs(1 To nRun)
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
""", """        ReDim runs(1 To nRun)
        Dim dtsG() As Double
        ReDim dtsG(1 To nRun)
        i = 0
        Dim rk As Variant
        For Each rk In runsGrupo.Keys
            i = i + 1
            runs(i) = CLng(rk)
            dtsG(i) = CDbl(runsGrupo(rk))
        Next rk
        ' mesma ordem cronologica do motor do Painel (ADR-049)
        mEstatistica.OrdenarCorridas runs, dtsG, nRun
""", rotulo='BI FlagsDoMotor')
    t = troca(t, '    lote = Trim$(CStr(ThisWorkbook.Names("loteAtivo").RefersToRange.Value))\n    If Len(selAn) = 0 Then ReconciliarComCalc',
              '    lote = mLotes.LotePainel()        ' + "' o lote que o Calc esta mostrando\n" + '    If Len(selAn) = 0 Then ReconciliarComCalc',
              rotulo='BI reconciliar')
    # a funcao de bloco nao tem mais chamador
    t = troca_bloco(t, "' Media e DP do analito NAQUELE lote, no nivel pedido.\n",
                    "Public Sub AtualizarBIData()",
                    "' AlvoDoLote por aritmetica de bloco removida (ADR-049): o alvo vem de\n"
                    "' mLotes.ParametrosLote, pela chave LOTE|ANALITO, como no resto do sistema.\n\n",
                    'BI AlvoDoLote')
    t = troca(t, "'   media e DP alvo POR LOTE ................... LotesStore, bloco do lote\n",
              "'   media e DP alvo POR LOTE ................... LotesStore, chave LOTE|ANALITO (mLotes)\n",
              rotulo='BI cabecalho')
    # memo do bias de EP por (analito, ano): era uma varredura das abas de EQA POR
    # LINHA do banco -- o BI nao terminava em 20 min com 100 mil linhas (T10)
    t = troca(t, """        biasEP = mCEQ.BiasEQ(an2, anoBI, "SIGNED")
""", """        ' ADR-049: memo por (analito, ano). BiasEQ varre as abas de EQA inteiras;
        ' chamado por LINHA do banco, o BI passava de 20 min com 100 mil linhas.
        ' So ha ~60 combinacoes (analito, ano), e o resultado e o mesmo.
        Dim memoBiasBI As Object, kBias As String
        If memoBiasBI Is Nothing Then
            Set memoBiasBI = CreateObject("Scripting.Dictionary")
            memoBiasBI.CompareMode = 1
        End If
        kBias = UCase$(an2) & "|" & CStr(anoBI)
        If memoBiasBI.Exists(kBias) Then
            biasEP = memoBiasBI(kBias)
        Else
            biasEP = mCEQ.BiasEQ(an2, anoBI, "SIGNED")
            memoBiasBI.Add kBias, biasEP
        End If
""", rotulo='BI memo bias')
    # plano de CQ por GRUPO (analito, lote), nao por linha
    t = troca_bloco(t, """    For k = 1 To n
        ch = CStr(saida(k, 8)) & "|" & CStr(saida(k, 12))
        If pior.Exists(ch) Then
            sp = CDbl(pior(ch))
""", """        Else
            ' Sem Sigma valido em nenhum nivel""", """    ' ADR-049: o plano depende SO do Sigma do grupo (analito, lote). Calcular
    ' por LINHA repetia ~13 leituras do Cfg_PlanoQC para cada resultado do
    ' banco -- o BI nao terminava em 20 min. Agora: uma vez por grupo, copiada.
    Dim memoPlano As Object, pl As Variant, jj As Long
    Set memoPlano = CreateObject("Scripting.Dictionary")
    memoPlano.CompareMode = 1
    For k = 1 To n
        ch = CStr(saida(k, 8)) & "|" & CStr(saida(k, 12))
        If pior.Exists(ch) Then
            sp = CDbl(pior(ch))
            If Not memoPlano.Exists(ch) Then
                ReDim pl(69 To 84)
                pl(77) = sp
                pl(78) = "Nivel " & CStr(nivelPior(ch))
                pl(79) = mQualidade.ClassificarSigma(sp)
                pl(69) = mPlanoQC.DPMdoSigma(sp)
                pl(70) = mPlanoQC.RendimentoDoSigma(sp)
                pl(71) = mPlanoQC.PlanoQC(sp, "REGRAS")
                pl(72) = mPlanoQC.PlanoQC(sp, "N")
                pl(73) = mPlanoQC.PlanoQC(sp, "RUNSIZE")
                pl(74) = mPlanoQC.PlanoQC(sp, "FREQUENCIA")
                pl(75) = mPlanoQC.CoberturaWestgard(sp)
                pl(76) = mPlanoQC.PlanoQC(sp, "REFERENCIA")
                For jj = 1 To 5
                    pl(79 + jj) = mPlanoQC.RegraNoPlano(sp, mEstatistica.NomeRegraWestgard(jj))
                Next jj
                memoPlano.Add ch, pl
            End If
            pl = memoPlano(ch)
            For jj = 69 To 84
                saida(k, jj) = pl(jj)
            Next jj
""", rotulo='BI plano por grupo')
    # GarantirAba escrevia o cabecalho da BI_Data PROTEGIDA: erro 1004 e a macro
    # parada no depurador (visto em 19/09/2026, arquivo de producao)
    t = troca(t, """    cb = Cab()
    For i = 0 To UBound(cb)
        ws.Cells(BI_CAB, i + 1).Value = cb(i)
    Next i
    ws.rows(BI_CAB).Font.Bold = True
    Set GarantirAba = ws""", """    cb = Cab()
    ' ADR-046/049: a BI_Data fica protegida no arquivo de producao. Escrever o
    ' cabecalho sem destrancar dava erro 1004 e deixava a macro PARADA no
    ' depurador -- "Atualizar BI" nunca terminava. Mesma guarda do resto.
    Dim protCab As Boolean
    protCab = LiberarEscrita(ws)
    On Error GoTo restauraCab
    For i = 0 To UBound(cb)
        ws.Cells(BI_CAB, i + 1).Value = cb(i)
    Next i
    ws.rows(BI_CAB).Font.Bold = True
restauraCab:
    Dim nEc As Long, sEc As String
    nEc = Err.Number: sEc = Err.Description
    RestaurarProtecao ws, protCab
    On Error GoTo 0
    If nEc <> 0 Then Err.Raise nEc, "mBI.GarantirAba", sEc
    Set GarantirAba = ws""", rotulo='BI GarantirAba protecao')
    return t


def patch_painel(t):
    # Painel (Planilha7): E3 = lote em analise; De/Ate refazem o motor.
    t = troca(t, """    If Not Intersect(Target, Me.Range("B3,G3,G4,M3,M4,N3,N4")) Is Nothing Then
        Application.EnableEvents = False
        AtualizarEixos
        Application.EnableEvents = True
    End If""", """    If Not Intersect(Target, Me.Range("E3")) Is Nothing Then
        ' ADR-049: lote em analise. Carrega media/DP/limites DESSE lote.
        mLotes.TrocarLoteAnalise
    End If
    If Not Intersect(Target, Me.Range("B3")) Is Nothing Then
        ' analito: refaz o motor so se ele ficou para tras (o spinner ja refaz)
        Application.EnableEvents = False
        mEstatistica.GarantirMotorDoPainel
        AtualizarEixos
        Application.EnableEvents = True
    End If
    If Not Intersect(Target, Me.Range("G3,G4")) Is Nothing Then
        ' periodo: a janela de corridas do motor depende do fim do periodo
        Application.EnableEvents = False
        mEstatistica.RecalcularAnalitoAtual
        Application.EnableEvents = True
    End If
    If Not Intersect(Target, Me.Range("M3,M4,N3,N4")) Is Nothing Then
        Application.EnableEvents = False
        AtualizarEixos
        Application.EnableEvents = True
    End If""", rotulo='Painel change')
    t = troca(t, "Private Sub Worksheet_Activate()\n    AtualizarEixos\n    HookCharts\nEnd Sub",
              "Private Sub Worksheet_Activate()\n    mEstatistica.GarantirMotorDoPainel\n    AtualizarEixos\n    HookCharts\nEnd Sub",
              rotulo='Painel activate')
    return t


def patch_workbook(t):
    t = troca(t, "    mDados.AtualizarListasAno\n    On Error GoTo 0\n",
              "    mDados.AtualizarListasAno\n    mLotes.SincronizarLotesAoAbrir\n    On Error GoTo 0\n",
              rotulo='Workbook_Open')
    return t


ANALITOS_CLS = """
' ADR-049: Media/DP digitados aqui pertencem ao lote em tela (loteParam) e sao
' gravados NA HORA na LotesStore, com rastro no Audit_Log.
Private Sub Worksheet_Change(ByVal Target As Range)
    If Intersect(Target, Me.Range("E4:J43")) Is Nothing Then Exit Sub
    mLotes.ParametroEditado Target
End Sub
"""


def main(entrada):
    gravar('mEstatistica.bas', patch_mestatistica(ler(entrada, 'mEstatistica.bas')))
    gravar('mEstatPeriodo.bas', patch_mestatperiodo(ler(entrada, 'mEstatPeriodo.bas')))
    gravar('mBI.bas', patch_mbi(ler(entrada, 'mBI.bas')))
    gravar('Planilha7.cls', patch_painel(ler(entrada, 'Planilha7.cls')))
    gravar('EstaPastaDeTrabalho.cls', patch_workbook(ler(entrada, 'EstaPastaDeTrabalho.cls')))
    p3 = ler(entrada, 'Planilha3.cls')
    if 'Worksheet_Change' in p3:
        raise SystemExit('Planilha3 ja tem Worksheet_Change -- revisar antes de sobrescrever')
    gravar('Planilha3.cls', p3.rstrip('\n') + '\n' + ANALITOS_CLS)
    # nada sobrou chamando o que saiu
    for nome in ('mEstatistica.bas', 'mEstatPeriodo.bas', 'mBI.bas'):
        s = ler(SRC, nome)
        for proibido in ('LoteDoAnoOuAtivo(', 'BlocoDoLoteBI(nucleo)', 'AlvoDoLote(iB'):
            if proibido in s:
                raise SystemExit(f'{nome} ainda contem {proibido}')
    print('ok: modulos gerados em', SRC)


if __name__ == '__main__':
    main(sys.argv[1])
