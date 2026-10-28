# -*- coding: utf-8 -*-
"""qa_westgard_sequencia.py -- ADR-083: as regras de Westgard do MESMO nivel andam pela SEQUENCIA DE ATIVOS do nivel;
as regras ENTRE niveis continuam pareando pelo RUN. Roda numa COPIA ja instalada (instalar_adr083.py); nada e salvo.

Uso: python qa_westgard_sequencia.py <Bioquimica|Hematologia> <copia_instalada.xlsm> <saida.json>

Parte 1 -- MOTOR, pela PONTE: funcoes QA083_* acrescentadas ao mEstatistica DESTA copia (como a ponte de
_arquitetura/scripts_fase3/testar_westgard_por_modulo.py, mas dentro do modulo, para chamar tambem os detectores
privados), z/td montados no teste. O motor ANTIGO (controle) e o mesmo modulo com AvaliarWestgard, Det_2_2s_AcrossRun,
Det_MesmoNivel_Lado e Det_MesmoNivel_Sequencia trocados pelo texto do ADR-081 (MOTOR_ANTIGO, abaixo).
  W00 pre-condicao: o mEstatistica e o mWestgardKnowledge da copia = fontes (ADR-083 instalado); a ponte passa no lint
  W01 (a) N1 +2,3 / corrida so de N2 / N1 +2,3: 2_2s ACROSS do N1 na posicao 4 (runIni 2), marca so em (N1,4), sem
      "nao avaliavel"; motor antigo: 0 eventos e NaoAvaliavel=2 (Bio; na Hema o detector direto, W11)
  W02 ativo intermediario do mesmo nivel interrompe (+2,3 +0,5 +2,4): 0 de 2_2s ACROSS nos dois motores
  W03 lados opostos com corrida so de N2 no meio: 0 de 2_2s ACROSS e 0 de R_4s (R_4s nunca cruza RUN)
  W04 (c) regra ENTRE niveis continua pelo RUN: niveis em corridas diferentes nao formam 2_2s/2of3_2s WITHIN_RUN
  W05 (c) NxR oficial nao compacta: corrida incompleta deixa 4_1s N2_R2 (Bio) / 6x N3_R2 (Hema) nao avaliavel
  W06 complementar com lacunas entra no trace com " (n obs)" e NAO marca (8x SAME_LEVEL_R8 / 3_1s e 6x SAME_LEVEL);
      motor antigo: nada
  W07 Hema 2of3_2s SAME_LEVEL_R3 = 2 ou mais de 3 (+2,3 +0,5 +2,4): 1 COMPLEMENTAR; antigo (3 de 3): 0. Nas duas areas
      (detector direto): a observacao que FECHA a janela passa do limite -- +0,1 +2,3 +2,4 +0,5 = 1 evento (run 3) e
      -2,3 -2,4 +2,3 = 0
  W08 contadores zerados a cada janela: 1 observacao do outro lado no meio impede TODA janela que a contem
  W09 R_4s, 2_2s/2of3_2s WITHIN e NxR identicos nos dois motores numa serie mista (marcas e trace entre niveis)
  W10 SeqDoNivel devolve as posicoes avaliaveis do nivel
  W11 (f) Bio = Hema: os detectores do mesmo nivel chamados direto dao a MESMA saida canonica nas duas areas, e o
      texto dos procedimentos compartilhados e identico em src/bio e src/hema
Parte 2 -- LOTE REAL (o do Painel e ate 3 outros com mais resultados), motor antigo x novo:
  W12 (b) superconjunto: toda linha antiga de Eventos_Westgard (14 colunas) continua; as novas so de detector do mesmo nivel
  W13 (c) detectores ENTRE niveis e 1_3s: linhas identicas
  W14 (d) Hema: nenhum evento OFICIAL muda e o bloco de flags do Eng_Saida do analito em tela e identico;
      Bio: so o 2_2s ACROSS muda (R22s/Veredicto)
  W15 EventosWestgardOficiais/AuditarMudancasWestgard (unidade): evento oficial levado para um RUN fora do eixo = 1
      WESTGARD_EVENTO_CORRIDA_NOVA ([CORRIDA_NOVA]) + 1 WESTGARD_EVENTO_REMOVIDO ([RETROATIVO]) no Audit_Log com RUN,
      analito, lote, nivel, regra e SEM AU_CHAVE/AU_SEQ; foto sem a referencia das corridas = WESTGARD_EVENTO_NOVO
      [NAO_CLASSIFICADO]; J2 vazio (aba em reconstrucao) = foto vazia e nada auditado; escopo diferente = nada auditado
Parte 3 -- PONTA A PONTA, pelo caminho real (manual digitado, InativarPorId/ReativarPorId + AtualizarSoEntradas):
  W20 pre-condicao: manuais em dois dias livres viram as corridas planejadas (D2: N1 -2,3 | so N2 | N1 -2,4;
      D1: N1 +2,3 | N1 +0,5 | N1 +2,4, com os outros niveis em media-0,3DP), todos ATIVOS
  W21 (a) D2: 2_2s ACROSS OFICIAL do N1 no RUN da 3a corrida com RUN_Inicial da 1a (Bio); Hema: 2of3_2s
      SAME_LEVEL_R3 COMPLEMENTAR e nenhum OFICIAL novo; motor antigo na mesma base: nada disso
  W22 (e) inativar o N1 intermediario de D1: Bio ganha EXATAMENTE o 2_2s ACROSS (RUN da 3a, inicio da 1a) e o Audit_Log
      tem WESTGARD_EVENTO_NOVO [RETROATIVO] sem AU_CHAVE (precisa da chamada em mIntegracao.AtualizarDadosCore); Hema:
      nenhum OFICIAL muda
  W23 (e) reativar: eventos OFICIAIS = os de antes da inativacao e o Audit_Log tem WESTGARD_EVENTO_REMOVIDO (Bio)
  W24 (g) com o N1 inativado: Eng_Saida (R22s/Veredicto), Calc (flag e veredicto), engPainel e MetricasWestgard
      coerentes com Eventos_Westgard nas corridas do teste (e a BI_Data, se o arquivo tiver identidade de BI). Nada passa
      calado: as corridas A/B/C TEM de estar no Calc, o N1 do engPainel tem de ser numero e, na Bioquimica com
      identidade de BI, a linha da corrida C tem de existir na BI_Data
Parte 4 -- FOTO POR LOTE x EQUIPAMENTO (opcao completa do ADR-083, aba Westgard_Fotos), num lote L2 que NAO e o do
Painel (manuais num dia livre do analito; o Painel e a aba Eventos_Westgard ficam no lote em que estavam):
  W30 pre-condicao: a aba Westgard_Fotos existe (instalar_adr083.py), protegida, foto completa (J2) com linhas CONTROLE;
      mIntegracao chama mEstatistica.AuditarFotosWestgard
  W31 manuais de L2 viram as corridas planejadas, ATIVAS (Bio: N1 +2,3 | +0,5 | +2,4; Hema: N1 +3,5), e o ATUALIZAR
      audita a regra nova de L2 como rotina: Hema = WESTGARD_EVENTO_CORRIDA_NOVA 1_3s [CORRIDA_NOVA] com o lote L2
  W32 (a)(d) inativar (InativarPorId + AtualizarSoEntradas) o resultado de L2: Bio = WESTGARD_EVENTO_NOVO 2_2s ACROSS
      [RETROATIVO] na 3a corrida; Hema = WESTGARD_EVENTO_REMOVIDO 1_3s [RETROATIVO] -- com lote L2, equipamento,
      nivel, analito e sem AU_CHAVE; TEMPOS do Sync_Log com westgard_fotos RECALC >= 1; Eventos_Westgard, loteSel e
      selEquipamento continuam no lote/equipamento de antes
  W33 (b)(d) reativar: Bio = WESTGARD_EVENTO_REMOVIDO; Hema = WESTGARD_EVENTO_NOVO [RETROATIVO]; Painel intacto
  W34 (c) ATUALIZAR sem mudanca (AtualizarSoEntradas de novo): nenhuma linha WESTGARD_* no Audit_Log, RECALC=0 e
      MUDANCAS=0 no TEMPOS do Sync_Log, com o tempo gasto anotado (motor PQ: o DB_SEAC pode trazer resultado novo --
      so linhas de rotina sao toleradas)
  W35 (e) a foto bate com a VARREDURA do instalador (RegistrarEventosWestgard lote a lote, com o Seletor do Painel)
      no lote do Painel, em L2 e em ate 2 outros lotes x equipamento; as linhas CORRIDAS de L2 tem as corridas do teste
  W36 meta de desempenho: sem mudanca, a foto custa < 1 s no ATUALIZAR (o tempo medido vai na evidencia)
  (f) Bio e Hema: a suite roda nas duas areas; o W11 confere o texto identico dos procedimentos da foto
  W99 o arquivo de entrada nao foi alterado
Sai com 1 se algum falhar.
"""
import collections
import datetime as dt
import json
import os
import re
import sys
import tempfile
import time
import traceback

AQUI = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, AQUI)
from qa_final import QA, reg, RES, sha256_arquivo  # noqa: E402
sys.path.insert(0, os.path.join(AQUI, '..'))
sys.path.insert(0, os.path.join(AQUI, '..', 'estudos'))
import instalar_integracao as ii  # noqa: E402
import codigo_atual  # noqa: E402
import lint_vba  # noqa: E402
from instalar_adr083 import Seletor, ler_eventos, ler_foto, conferir_foto, LOTE_DA_ABA  # noqa: E402

SRC = os.path.join(AQUI, '..', 'src')
ABA_EV = 'Eventos_Westgard'
C_DATA, C_RUN, C_ANALITO, C_NIVEIS, C_REGRA, C_DETECTOR = 1, 2, 3, 4, 5, 6
C_ESCOPO, C_CLASSE, C_N, C_R, C_RUNINI, C_EVID = 7, 8, 9, 10, 11, 12
TB_MAN = 'tblResultados_Manuais'
MOTIVO_MAN = 'Falha do interfaceamento'
# Audit_Log (mAuditoria.AU_*)
AU_ACAO, AU_RUN, AU_LOTE, AU_NIVEL, AU_ANALITO, AU_STANT, AU_STNOVO, AU_MOTIVO = 7, 10, 13, 14, 15, 21, 22, 23
AU_CHAVE, AU_SEQ, AU_PARECER, AU_EQUIP = 9, 16, 24, 12
ABA_FOTOS = 'Westgard_Fotos'
NEF = 7                                           # campos por nivel no Eng_Saida
MESMO_NIVEL = {'bio': {('2_2s', 'ACROSS_RUN_SAME_LEVEL'), ('4_1s', 'SAME_LEVEL_R4'), ('8x', 'SAME_LEVEL_R8')},
               'hema': {('2of3_2s', 'SAME_LEVEL_R3'), ('3_1s', 'SAME_LEVEL_R3'), ('6x', 'SAME_LEVEL_R6')}}
COMPARTILHADOS = ['SeqDoNivel', 'DetalheJanelaNivel', 'AvaliarWestgard', 'Det_2_2s_AcrossRun', 'Det_MesmoNivel_Lado',
                  'Det_MesmoNivel_Sequencia', 'Det_R_4s', 'Det_2_2s_WithinRun', 'Det_2of3_2s_WithinRun', 'Det_3_1s_N3R1',
                  'Det_4_1s_N2R2', 'Det_JanelaNR', 'CorridaCompleta', 'EventosWestgardOficiais',
                  'AuditarMudancasWestgard', 'RegistrarMudancaWestgard', 'ParteEscopoWg', 'ChaveEventoWestgard',
                  'TxtCelWg', 'RunTxtWg', 'DataIsoWg', 'RegistrarEventosWestgard', 'ChaveMetaWg', 'ClasseEventoNovoWg',
                  'GravarMudancaWg',
                  # ADR-083 (opcao completa): a foto por lote x equipamento
                  'EventosEscopoWg', 'AcrescentarSerieWg', 'CompararEventosWg', 'ResumoNaoDetalhadosWg', 'DescEventoWg',
                  'AuditarFotosWestgard', 'GravarFotosWestgard', 'SincronizarFotosWg', 'AbaFotosWg', 'GarantirDicWg',
                  'EscopoDaLinhaWg', 'PesoAnalitoWg', 'ImpressoesDigitaisWg', 'LerControleWg', 'LerAntesWg',
                  'EventosAtuaisWg', 'CalcularEscopoWg', 'GravarFotoWg', 'AcrescentarCorridasWg']

# ---------------------------------------------------------------- VBA da ponte (so numa COPIA, nunca salva)
PONTE_BASE = r'''
' ===== PONTE DO QA (testes/qa_westgard_sequencia.py, ADR-083) -- so numa COPIA de teste, nunca salva =====
' itens = "t|i|z;t|i|z;..." com z de ponto decimal (Val nao depende do formato regional)
Private Sub QA083_Montar(ByVal itens As String, ByVal nRun As Long, ByRef zq() As Double, ByRef tq() As Boolean)
    Dim pq As Variant, cq As Variant, kq As Long, tt As Long, iq As Long
    ReDim zq(0 To NLV - 1, 1 To nRun)
    ReDim tq(0 To NLV - 1, 1 To nRun)
    pq = Split(itens, ";")
    For kq = LBound(pq) To UBound(pq)
        If Len(Trim$(CStr(pq(kq)))) > 0 Then
            cq = Split(CStr(pq(kq)), "|")
            tt = CLng(Val(cq(0))): iq = CLng(Val(cq(1)))
            If tt >= 0 And tt <= NLV - 1 Then
                If iq >= 1 And iq <= nRun Then
                    zq(tt, iq) = Val(cq(2))
                    tq(tt, iq) = True
                End If
            End If
        End If
    Next kq
End Sub

' "NOME=t,i t,i ..." com as posicoes marcadas (= 1) de uma saida do motor
Private Function QA083_Marcas(ByVal rotulo As String, ByRef mq As Variant) As String
    Dim tt As Long, iq As Long, sq As String
    For tt = LBound(mq, 1) To UBound(mq, 1)
        For iq = LBound(mq, 2) To UBound(mq, 2)
            If Not IsEmpty(mq(tt, iq)) Then
                If mq(tt, iq) = 1 Then sq = sq & IIf(Len(sq) > 0, " ", "") & CStr(tt) & "," & CStr(iq)
            End If
        Next iq
    Next tt
    QA083_Marcas = rotulo & "=" & sq
End Function

Public Function QA083_Avaliar(ByVal itens As String, ByVal nRun As Long) As String
    Dim zq() As Double, tq() As Boolean
    Dim q13 As Variant, q2 As Variant, qR4 As Variant, q41 As Variant, qSq As Variant, q12 As Variant
    On Error GoTo falhaQ
    QA083_Montar itens, nRun, zq, tq
    ReDim q13(0 To NLV - 1, 1 To nRun): ReDim q2(0 To NLV - 1, 1 To nRun)
    ReDim qR4(0 To NLV - 1, 1 To nRun): ReDim q41(0 To NLV - 1, 1 To nRun)
    ReDim qSq(0 To NLV - 1, 1 To nRun): ReDim q12(0 To NLV - 1, 1 To nRun)
    AvaliarWestgard zq, tq, nRun, q13, q2, qR4, q41, qSq, q12
    QA083_Avaliar = "OK|" & QA083_Marcas("R13", q13) & ";" & QA083_Marcas("R2", q2) & ";" & _
                    QA083_Marcas("R4S", qR4) & ";" & QA083_Marcas("R41", q41) & ";" & _
                    QA083_Marcas("SEQ", qSq) & ";" & QA083_Marcas("A12", q12)
    Exit Function
falhaQ:
    QA083_Avaliar = "ERRO|" & Err.Number & " " & Err.Description
End Function

Public Function QA083_Trace() As String
    On Error Resume Next
    QA083_Trace = TraceWestgard()
End Function

Public Function QA083_NaoAval() As String
    On Error Resume Next
    QA083_NaoAval = NaoAvaliaveisWestgard()
End Function
'''

PONTE_NOVA = r'''
' Um detector do MESMO nivel chamado direto (texto identico nas duas areas -- prova Bio = Hema). Os nomes de
' regra/detector "QA083" nao existem em DETECTORES: o evento sai COMPLEMENTAR e a marca so vem do 2_2s ACROSS.
Public Function QA083_Detector(ByVal qual As String, ByVal itens As String, ByVal nRun As Long) As String
    Dim zq() As Double, tq() As Boolean, sq As Variant
    On Error GoTo falhaQ
    QA083_Montar itens, nRun, zq, tq
    ReDim sq(0 To NLV - 1, 1 To nRun)
    Set gTrace = New Collection
    Set gNaoAval = CreateObject("Scripting.Dictionary")
    gNaoAval.CompareMode = 1
    Select Case qual
        Case "2_2S_ACROSS"
            Det_2_2s_AcrossRun zq, tq, nRun, sq
        Case "LADO4"
            Det_MesmoNivel_Lado zq, tq, nRun, "QA083", "QA083_LADO4", 4, sq
        Case "SEQ2DE3"
            Det_MesmoNivel_Sequencia zq, tq, nRun, "QA083", "QA083_SEQ2DE3", 3, 2, 2#, sq
        Case Else
            QA083_Detector = "ERRO|detector desconhecido"
            Exit Function
    End Select
    QA083_Detector = "OK|" & QA083_Marcas("M", sq)
    Exit Function
falhaQ:
    QA083_Detector = "ERRO|" & Err.Number & " " & Err.Description
End Function

Public Function QA083_Seq(ByVal itens As String, ByVal nRun As Long, ByVal nivel0 As Long) As String
    Dim zq() As Double, tq() As Boolean, pq() As Long, mq As Long, kq As Long, sq As String
    On Error GoTo falhaQ
    QA083_Montar itens, nRun, zq, tq
    mq = SeqDoNivel(tq, nivel0, nRun, pq)
    For kq = 1 To mq
        sq = sq & IIf(kq > 1, ",", "") & CStr(pq(kq))
    Next kq
    QA083_Seq = "OK|" & sq
    Exit Function
falhaQ:
    QA083_Seq = "ERRO|" & Err.Number & " " & Err.Description
End Function

' Foto dos eventos oficiais e auditoria contra ela, como mIntegracao.AtualizarDadosCore fara.
Public Function QA083_Auditoria(ByVal acao As String) As String
    Static foto As Object
    Dim kq As Variant, sq As String
    On Error GoTo falhaQ
    Select Case acao
        Case "FOTO"
            Set foto = EventosWestgardOficiais()
            For Each kq In foto.Keys
                If IsObject(foto(kq)) Then                  ' "#RUNS": as corridas avaliadas (ADR-083)
                    sq = sq & CStr(kq) & Chr$(9) & CStr(foto(kq).Count) & " corrida(s)" & vbLf
                Else
                    sq = sq & CStr(kq) & Chr$(9) & CStr(foto(kq)) & vbLf
                End If
            Next kq
            QA083_Auditoria = "OK|" & sq
        Case "SEMREF"
            If foto.Exists("#RUNS") Then foto.Remove "#RUNS"
            QA083_Auditoria = "OK|"
        Case "AUDITAR"
            AuditarMudancasWestgard foto
            QA083_Auditoria = "OK|"
        Case Else
            QA083_Auditoria = "ERRO|acao desconhecida"
    End Select
    Exit Function
falhaQ:
    QA083_Auditoria = "ERRO|" & Err.Number & " " & Err.Description
End Function
'''

# ---------------------------------------------------------------- motor do ADR-081 (controle), texto original
MOTOR_ANTIGO = {
    'AvaliarWestgard': r'''Public Sub AvaliarWestgard(ByRef z() As Double, ByRef temDado() As Boolean, ByVal nRun As Long, _
                           ByRef r13 As Variant, ByRef r2sMulti As Variant, ByRef rR4 As Variant, _
                           ByRef r1sMulti As Variant, ByRef rSeq As Variant, ByRef a12 As Variant)
    Set gTrace = New Collection
    Set gNaoAval = CreateObject("Scripting.Dictionary")
    gNaoAval.CompareMode = 1

    Det_1_3s z, temDado, nRun, r13, a12
    Det_R_4s z, temDado, nRun, rR4

    If NLV >= 3 Then
        Det_2of3_2s_WithinRun z, temDado, nRun, r2sMulti
        Det_3_1s_N3R1 z, temDado, nRun, r1sMulti
        Det_6x_N3R2 z, temDado, nRun, rSeq
        ' Complementares: calculados e registrados, NAO consolidados.
        Det_MesmoNivel_Sequencia z, temDado, nRun, "2of3_2s", "SAME_LEVEL_R3", 3, 2#, r2sMulti
        Det_MesmoNivel_Sequencia z, temDado, nRun, "3_1s", "SAME_LEVEL_R3", 3, 1#, r1sMulti
        Det_MesmoNivel_Lado z, temDado, nRun, "6x", "SAME_LEVEL_R6", 6, rSeq
    Else
        Det_2_2s_WithinRun z, temDado, nRun, r2sMulti
        Det_2_2s_AcrossRun z, temDado, nRun, r2sMulti
        Det_4_1s_N2R2 z, temDado, nRun, r1sMulti
        Det_8x_N2R4 z, temDado, nRun, rSeq
        Det_MesmoNivel_Sequencia z, temDado, nRun, "4_1s", "SAME_LEVEL_R4", 4, 1#, r1sMulti
        Det_MesmoNivel_Lado z, temDado, nRun, "8x", "SAME_LEVEL_R8", 8, rSeq
    End If
End Sub
''',
    'Det_2_2s_AcrossRun': r'''Private Sub Det_2_2s_AcrossRun(ByRef z() As Double, ByRef td() As Boolean, ByVal nRun As Long, _
                               ByRef r2sMulti As Variant)
    Dim t As Long, i As Long
    For t = 0 To NLV - 1
        For i = 2 To nRun
            If Not (td(t, i) And td(t, i - 1)) Then
                NaoAvaliavel "2_2s", "ACROSS_RUN_SAME_LEVEL"
            ElseIf (z(t, i) > 2 And z(t, i - 1) > 2) Or _
                   (z(t, i) < -2 And z(t, i - 1) < -2) Then
                r2sMulti(t, i) = 1
                Evid "2_2s", "ACROSS_RUN_SAME_LEVEL", "WITHIN_MATERIAL_ACROSS_RUN", i, _
                     CStr(t + 1), i - 1, _
                     "N" & CStr(t + 1) & " " & Format$(z(t, i - 1), "0.00") & _
                     "->" & Format$(z(t, i), "0.00")
            End If
        Next i
    Next t
End Sub
''',
    'Det_MesmoNivel_Lado': r'''Private Sub Det_MesmoNivel_Lado(ByRef z() As Double, ByRef td() As Boolean, ByVal nRun As Long, _
                                ByVal regra As String, ByVal detector As String, _
                                ByVal n As Long, ByRef saida As Variant)
    Dim t As Long, i As Long, k As Long, acima As Long, abaixo As Long
    Dim ativo As Boolean
    ativo = DetectorAtivo(regra, detector)
    For t = 0 To NLV - 1
        For i = n To nRun
            Dim ok As Boolean
            ok = True
            For k = i - n + 1 To i
                If Not td(t, k) Then ok = False
            Next k
            If Not ok Then
                NaoAvaliavel regra, detector
                GoTo prox
            End If
            acima = 0: abaixo = 0
            For k = i - n + 1 To i
                If z(t, k) > 0 Then acima = acima + 1
                If z(t, k) < 0 Then abaixo = abaixo + 1
            Next k
            If acima = n Or abaixo = n Then
                If ativo Then saida(t, i) = 1
                Evid regra, detector, "WITHIN_MATERIAL_ACROSS_RUN", i, _
                     CStr(t + 1), i - n + 1, _
                     "N" & CStr(t + 1) & " runs " & CStr(i - n + 1) & ".." & CStr(i)
            End If
prox:
        Next i
    Next t
End Sub
''',
    'Det_MesmoNivel_Sequencia': r'''Private Sub Det_MesmoNivel_Sequencia(ByRef z() As Double, ByRef td() As Boolean, ByVal nRun As Long, _
                                     ByVal regra As String, ByVal detector As String, _
                                     ByVal n As Long, ByVal limite As Double, ByRef saida As Variant)
    Dim t As Long, i As Long, k As Long, acima As Long, abaixo As Long
    Dim ativo As Boolean
    ativo = DetectorAtivo(regra, detector)
    For t = 0 To NLV - 1
        For i = n To nRun
            Dim ok As Boolean
            ok = True
            For k = i - n + 1 To i
                If Not td(t, k) Then ok = False
            Next k
            If Not ok Then
                NaoAvaliavel regra, detector
                GoTo prox
            End If
            acima = 0: abaixo = 0
            For k = i - n + 1 To i
                If z(t, k) > limite Then acima = acima + 1
                If z(t, k) < -limite Then abaixo = abaixo + 1
            Next k
            If acima = n Or abaixo = n Then
                If ativo Then saida(t, i) = 1
                Evid regra, detector, "WITHIN_MATERIAL_ACROSS_RUN", i, _
                     CStr(t + 1), i - n + 1, _
                     "N" & CStr(t + 1) & " runs " & CStr(i - n + 1) & ".." & CStr(i)
            End If
prox:
        Next i
    Next t
End Sub
''',
}


# ---------------------------------------------------------------- utilidades
def txt(v):
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return '' if v is None else str(v).strip()


def tem_identidade_bi(wb):
    """O arquivo tem a BI ligada (nomes bi* + modulo mBI)? -- veio do instalar_adr083, que nao toca mais a BI_Data."""
    for n in ('biProduto', 'biWorkbookID', 'biContratoVersao'):
        try:
            v = wb.Names(n).RefersToRange.Value
        except Exception:                                      # noqa: BLE001 -- nome ausente ou constante
            try:
                v = str(wb.Names(n).RefersTo or '').lstrip('=').strip('"')
            except Exception:                                  # noqa: BLE001
                return False
        if not txt(v):
            return False
    return 'mBI' in {c.Name for c in wb.VBProject.VBComponents}


def iint(v):
    try:
        return int(float(v))
    except (TypeError, ValueError):
        return None


def num(v):
    return isinstance(v, float) or (isinstance(v, int) and not isinstance(v, bool) and v > -2146820000)


def norm_cel(v):
    if hasattr(v, 'year'):
        return v.replace(tzinfo=None).strftime('%Y-%m-%d %H:%M:%S')
    if isinstance(v, float):
        return round(v, 9)
    return '' if v is None else str(v)


def itens(pontos):
    return ';'.join(f'{t}|{i}|{float(z)!r}' for (t, i), z in sorted(pontos.items()))


def marcas(saida):
    """'OK|R13=;R2=0,4;...' -> {'R2': {(0, 4)}, ...}."""
    if not str(saida).startswith('OK|'):
        raise RuntimeError(f'ponte: {saida}')
    out = {}
    for parte in str(saida)[3:].split(';'):
        if '=' in parte:
            k, v = parte.split('=', 1)
            out[k] = {tuple(int(x) for x in p.split(',')) for p in v.split(' ') if p}
    return out


def procs_de(texto):
    t = texto.replace('\r\n', '\n')
    out = {}
    for m in re.finditer(r'^((?:Public |Private )?(?:Sub|Function) (\w+)\(.*?^End (?:Sub|Function)\n)', t, re.M | re.S):
        out[m.group(2)] = m.group(1)
    return out


def trocar_procs(texto, novos):
    t = texto.replace('\r\n', '\n')
    if not t.endswith('\n'):
        t += '\n'
    for nome, proc in novos.items():
        rx = re.compile(r'^(?:Public |Private )?(?:Sub|Function) ' + re.escape(nome) + r'\(.*?^End (?:Sub|Function)\n',
                        re.M | re.S)
        achados = list(rx.finditer(t))
        if len(achados) != 1:
            raise RuntimeError(f'procedimento {nome}: {len(achados)} ocorrencia(s) no modulo')
        p = proc.replace('\r\n', '\n')
        t = t[:achados[0].start()] + (p if p.endswith('\n') else p + '\n') + t[achados[0].end():]
    return t.replace('\n', '\r\n')


def lint(texto):
    d = tempfile.mkdtemp(prefix='qa083_')
    p = os.path.join(d, 'mEstatistica.bas')
    with open(p, 'w', encoding='utf-8', newline='') as f:
        f.write('Attribute VB_Name = "mEstatistica"\r\n' + texto)
    return lint_vba.checar(p)


class Motor:
    """Troca o texto do mEstatistica da COPIA: 'novo' (o instalado + ponte completa), 'antigo' (procedimentos do ADR-081
    + ponte base) ou 'original' (exatamente o instalado). Compila sempre (erro de compilacao = falha do teste)."""

    def __init__(self, q):
        self.q = q
        self.vbp = q.wb.VBProject
        cm = self.vbp.VBComponents('mEstatistica').CodeModule
        self.original = cm.Lines(1, cm.CountOfLines).replace('\r\n', '\n').replace('\n', '\r\n')
        self.atual = 'original'

    def texto(self, variante):
        if variante == 'novo':
            return self.original.rstrip('\r\n') + '\r\n' + PONTE_BASE.replace('\n', '\r\n') + PONTE_NOVA.replace('\n', '\r\n')
        if variante == 'antigo':
            return trocar_procs(self.original, MOTOR_ANTIGO).rstrip('\r\n') + '\r\n' + PONTE_BASE.replace('\n', '\r\n')
        return self.original

    def usar(self, variante):
        if variante == self.atual:
            return
        t = self.texto(variante)
        if variante != 'original':
            prob = lint(t)
            if prob:
                raise RuntimeError(f'ponte/motor {variante} reprovado no lint: {prob[:5]}')
        ii.substituir_codigo(self.vbp, 'mEstatistica', t, criar=False)
        try:
            ok = ii.compilar_vba(self.q.ex, self.q.wb)
        except SystemExit as e:                       # o cao de guarda da compilacao encerra o Excel com SystemExit
            raise RuntimeError(f'mEstatistica ({variante}): {e}') from None
        if ok is False:
            raise RuntimeError(f'mEstatistica ({variante}) nao compilou')
        self.atual = variante


# ---------------------------------------------------------------- leituras de planilha
def eventos(q):
    ws = q.ws(ABA_EV)
    n = int(float(ws.Range('J2').Value or 0))
    if n <= 0:
        return []
    v = ws.Range(ws.Cells(4, 1), ws.Cells(3 + n, 14)).Value
    return [tuple(r) for r in v if txt(r[C_REGRA - 1])]


def oficiais(linhas):
    return {(txt(r[C_ANALITO - 1]).upper(), txt(r[C_NIVEIS - 1]), txt(r[C_REGRA - 1]), txt(r[C_DETECTOR - 1]),
             iint(r[C_RUN - 1]), iint(r[C_RUNINI - 1])) for r in linhas if txt(r[C_CLASSE - 1]) == 'OFICIAL'}


def contagem(linhas):
    return collections.Counter(tuple(norm_cel(x) for x in r) for r in linhas)


def audit_ult(q):
    al = q.ws('Audit_Log')
    return al.Cells(al.Rows.Count, 1).End(-4162).Row


def audit_desde(q, r0):
    al = q.ws('Audit_Log')
    ult = audit_ult(q)
    if ult <= r0:
        return []
    v = al.Range(al.Cells(r0 + 1, 1), al.Cells(ult, 24)).Value
    return [{'acao': txt(r[AU_ACAO - 1]), 'run': iint(r[AU_RUN - 1]), 'lote': txt(r[AU_LOTE - 1]),
             'nivel': iint(r[AU_NIVEL - 1]), 'analito': txt(r[AU_ANALITO - 1]), 'st_ant': txt(r[AU_STANT - 1]),
             'st_novo': txt(r[AU_STNOVO - 1]), 'motivo': txt(r[AU_MOTIVO - 1]), 'chave': txt(r[AU_CHAVE - 1]),
             'seq': txt(r[AU_SEQ - 1]), 'parecer': txt(r[AU_PARECER - 1]), 'equip': txt(r[AU_EQUIP - 1]).upper()}
            for r in v]


def mesmo_lote(a, b):
    """Lote canonico x lote da base: "010" e "10" sao o mesmo numero (instalar_adr083.mesmo_lote)."""
    if a == b:
        return True
    try:
        return float(a) == float(b)
    except ValueError:
        return False


def flags_por_run(q):
    e = q.eng_saida()
    return {r: list(e['flags'][i]) for i, r in enumerate(e['runs'])}


def serial(d):
    return float((dt.datetime.combine(d, dt.time()) - dt.datetime(1899, 12, 30)).days)


# ---------------------------------------------------------------- a suite
def executar(produto, caminho):
    hash0 = sha256_arquivo(caminho)
    q = QA(produto, caminho)
    area = 'bio' if q.bio else 'hema'
    nlv = q.nlv
    try:
        def run(mac, *a, teto=600):
            return q.run(mac, *a, teto=teto)

        # ======================================================== W00
        cm = q.wb.VBProject.VBComponents('mEstatistica').CodeModule
        no_arquivo = cm.Lines(1, cm.CountOfLines)
        fonte = ii.codigo(os.path.join(SRC, area, 'mEstatistica.bas'))
        igual_est = codigo_atual.normalizar(no_arquivo).lower() == codigo_atual.normalizar(fonte).lower()
        cmk = q.wb.VBProject.VBComponents('mWestgardKnowledge').CodeModule
        igual_wk = codigo_atual.normalizar(cmk.Lines(1, cmk.CountOfLines)).lower() == \
            codigo_atual.normalizar(ii.codigo(os.path.join(SRC, 'comum', 'mWestgardKnowledge.bas'))).lower()
        cmi = q.wb.VBProject.VBComponents('mIntegracao').CodeModule
        integra = 'AuditarFotosWestgard' in cmi.Lines(1, cmi.CountOfLines)      # ADR-083 (opcao completa)
        motor = Motor(q)
        prob_novo, prob_antigo = lint(motor.texto('novo')), lint(motor.texto('antigo'))
        reg('W00 Pré-condição: mEstatistica e mWestgardKnowledge da cópia = fontes (ADR-083 instalado); a ponte e o motor '
            'de controle (ADR-081) passam no lint', igual_est and igual_wk and not prob_novo and not prob_antigo,
            {'mEstatistica_igual_fonte': igual_est, 'mWestgardKnowledge_igual_fonte': igual_wk,
             'mIntegracao_chama_AuditarFotosWestgard': integra, 'lint_novo': prob_novo[:3],
             'lint_antigo': prob_antigo[:3]})
        if not (igual_est and igual_wk):
            raise RuntimeError('a cópia não tem o VBA do ADR-083 (rodar instalar_adr083.py nela antes)')

        def av(pontos, nrun):
            r = run('mEstatistica.QA083_Avaliar', itens(pontos), nrun)
            tr = [l for l in str(run('mEstatistica.QA083_Trace') or '').split('\n') if l.strip()]
            na = str(run('mEstatistica.QA083_NaoAval') or '')
            return marcas(r), tr, na

        def det(qual, pontos, nrun):
            r = run('mEstatistica.QA083_Detector', qual, itens(pontos), nrun)
            tr = [l for l in str(run('mEstatistica.QA083_Trace') or '').split('\n') if l.strip()]
            return marcas(r), tr

        def com(tr, prefixo):
            return [l for l in tr if l.startswith(prefixo)]

        # cenarios (nivel 0-based, posicao 1-based)
        C01 = ({(0, 1): 0.1, (0, 2): 2.3, (0, 4): 2.3, (1, 1): 0.2, (1, 2): 0.1, (1, 3): 0.4, (1, 4): -0.1}, 4)
        C02 = ({(0, 1): 2.3, (0, 2): 0.5, (0, 3): 2.4, (1, 1): 0.1, (1, 2): 0.1, (1, 3): 0.1}, 3)
        C03 = ({(0, 1): 2.3, (1, 1): 0.1, (1, 2): 0.0, (0, 3): -2.3, (1, 3): 0.1}, 3)
        if q.bio:
            C04 = ({(0, 1): 2.3, (1, 2): 2.4}, 2)
            C05 = ({(0, 1): 1.3, (0, 3): 1.4, (1, 1): 1.2, (1, 2): 1.4, (1, 3): 1.5}, 3)
            C06 = ({**{(0, i): 0.5 for i in (1, 2, 4, 5, 7, 8, 10, 11)}, **{(1, i): -0.5 for i in (3, 6, 9)}}, 11)
            C08 = ({(0, i): (-0.5 if i == 8 else 0.5) for i in range(1, 16)}, 15)
            C09 = ({(0, 1): 2.2, (1, 1): -2.1, (0, 2): 2.3, (1, 2): 2.4, (0, 3): 1.3, (1, 3): 1.2, (0, 4): 1.4,
                    (1, 4): 1.5, **{(t, i): 0.5 for i in range(5, 9) for t in (0, 1)}, (0, 10): 0.3}, 10)
        else:
            C04 = ({(0, 1): 2.3, (1, 2): 2.4, (2, 3): 2.5}, 3)
            C05 = ({(0, 1): 0.4, (1, 1): 0.5, (2, 1): 0.6, (1, 2): 0.7, (2, 2): 0.2, (0, 3): 0.3, (1, 3): 0.5,
                    (2, 3): 0.4}, 3)
            C06 = ({**{(0, i): 1.2 for i in (1, 3, 5)}, **{(t, i): 0.2 for i in range(1, 6) for t in (1, 2)}}, 5)
            C06b = ({**{(0, i): 0.3 for i in (1, 3, 5, 7, 9, 11)}, **{(t, i): -0.2 for i in range(1, 12) for t in (1, 2)}}, 11)
            C07 = ({(0, 1): 2.3, (0, 2): 0.5, (0, 3): 2.4, **{(t, i): 0.1 for i in (1, 2, 3) for t in (1, 2)}}, 3)
            C08 = ({(0, i): (-0.3 if i == 6 else 0.3) for i in range(1, 12)}, 11)
            C09 = ({(0, 1): 2.3, (1, 1): 2.4, (2, 1): 0.5, (0, 2): 2.3, (1, 2): 0.2, (2, 2): -2.2, (0, 3): 1.2,
                    (1, 3): 1.4, (2, 3): 1.1, **{(t, i): 0.4 for i in (4, 5) for t in (0, 1, 2)}, (0, 7): 0.3}, 7)

        # ======================================================== motor ANTIGO (controle)
        motor.usar('antigo')
        ant = {}
        ant['C01'] = av(*C01)
        ant['C02'] = av(*C02)
        ant['C06'] = av(*C06)
        if not q.bio:
            ant['C06b'] = av(*C06b)
            ant['C07'] = av(*C07)
        ant['C09'] = av(*C09)

        # ---- lote real com o motor antigo
        p = q.ws('Painel')
        lote_sel = q.nome('loteSel')
        lote_sel_orig = (lote_sel.HasFormula, lote_sel.Formula, lote_sel.Value)
        lote_painel = txt(run('mLotes.LotePainel'))
        fin = q.ler('tblCQ_Final')
        eq_painel = txt(run('mEstatistica.EquipFiltro')).upper()
        cont_lote = collections.Counter(txt(r['LOTE']) for r in fin if r['PARTICIPA_ESTATISTICA'] == 'SIM'
                                        and (not eq_painel or txt(r['EQUIPAMENTO']).upper() == eq_painel))
        lotes_reais = [lote_painel] + [L for L, _ in cont_lote.most_common() if L and L != lote_painel][:3]

        def por_lote(L):
            q.ex.xl.EnableEvents = False
            try:
                lote_sel.Worksheet.Unprotect(ii.SENHA)
                lote_sel.Value = "'" + L
            finally:
                q.ex.xl.EnableEvents = True

        def devolver_lote():
            q.ex.xl.EnableEvents = False
            try:
                lote_sel.Worksheet.Unprotect(ii.SENHA)
                tem_f, f, v = lote_sel_orig
                if tem_f:
                    lote_sel.Formula = f
                elif v in (None, ''):
                    lote_sel.ClearContents()
                elif isinstance(v, str):
                    lote_sel.Value = "'" + v
                else:
                    lote_sel.Value = v
            finally:
                q.ex.xl.EnableEvents = True

        def capturar_lotes():
            out = {}
            run('mEstatistica.InvalidarCache')
            for L in lotes_reais:
                por_lote(L)
                run('mEstatistica.RegistrarEventosWestgard', teto=900)
                out[L] = eventos(q)
            devolver_lote()
            run('mEstatistica.InvalidarCache')
            run('mEstatistica.RecalcularAnalitoAtual', teto=900)
            return out, flags_por_run(q), txt(q.ws('Eng_Saida').Range('C1').Value)

        p.Unprotect(ii.SENHA)
        reais_ant, flags_ant, an_tela_ant = capturar_lotes()

        # ======================================================== motor NOVO
        motor.usar('novo')
        nov = {}
        nov['C01'] = av(*C01)
        nov['C02'] = av(*C02)
        nov['C03'] = av(*C03)
        nov['C04'] = av(*C04)
        nov['C05'] = av(*C05)
        nov['C06'] = av(*C06)
        if not q.bio:
            nov['C06b'] = av(*C06b)
            nov['C07'] = av(*C07)
        nov['C08'] = av(*C08)
        nov['C09'] = av(*C09)

        # ---- W01
        m1, tr1, na1 = nov['C01']
        ma1, tra1, naa1 = ant['C01']
        d1, dtr1 = det('2_2S_ACROSS', *C01)
        if q.bio:
            lin = com(tr1, '2_2s|ACROSS_RUN_SAME_LEVEL|WITHIN_MATERIAL_ACROSS_RUN|OFICIAL|run=4|runIni=2|niveis=1|N=1 R=2')
            outras2 = {k for k in m1.get('R2', set())} - {(0, 4)}
            ok1 = (len(lin) == 1 and len(com(tr1, '2_2s|ACROSS')) == 1 and m1.get('R2') == {(0, 4)}
                   and '2_2s|ACROSS_RUN_SAME_LEVEL' not in na1
                   and not com(tra1, '2_2s|ACROSS') and not ma1.get('R2') and '2_2s|ACROSS_RUN_SAME_LEVEL=2' in naa1)
            ev1 = {'trace_novo': com(tr1, '2_2s'), 'R2_novo': sorted(m1.get('R2', [])), 'outras_marcas_2s': sorted(outras2),
                   'nao_avaliaveis_novo': na1, 'trace_antigo': com(tra1, '2_2s'), 'R2_antigo': sorted(ma1.get('R2', [])),
                   'nao_avaliaveis_antigo': naa1}
        else:
            ok1 = d1.get('M') == {(0, 4)} and len(com(dtr1, '2_2s|ACROSS_RUN_SAME_LEVEL|WITHIN_MATERIAL_ACROSS_RUN|')) == 1 \
                and any('|run=4|runIni=2|niveis=1|' in l for l in dtr1)
            ev1 = {'detector_direto': sorted(d1.get('M', [])), 'trace': dtr1,
                   'obs': 'a Hematologia nao tem 2_2s ACROSS na matriz; o detector compartilhado e chamado direto'}
        reg('W01 (a) N1 > +2DP em duas observações consecutivas do N1 com uma corrida só de N2 entre elas: 2_2s ACROSS '
            'detectado AGORA na posição real (marca só em N1/pos 4, RUN_Inicial = pos 2, sem "não avaliável"); com o '
            'motor do ADR-081 não era (NaoAvaliavel = 2)', ok1, ev1)

        # ---- W02
        m2, tr2, _ = nov['C02']
        ma2, tra2, _ = ant['C02']
        d2, _ = det('2_2S_ACROSS', *C02)
        reg('W02 Ativo intermediário do mesmo nível interrompe a regra (+2,3 +0,5 +2,4): nenhum 2_2s ACROSS, nos dois '
            'motores e no detector direto', not com(tr2, '2_2s|ACROSS') and not com(tra2, '2_2s|ACROSS') and not d2.get('M'),
            {'trace_novo': com(tr2, '2_2s'), 'trace_antigo': com(tra2, '2_2s'), 'detector_direto': sorted(d2.get('M', []))})

        # ---- W03
        m3, tr3, _ = nov['C03']
        d3, _ = det('2_2S_ACROSS', *C03)
        reg('W03 Lados opostos com corrida só de N2 no meio: nenhum 2_2s ACROSS e nenhum R_4s (R_4s nunca cruza RUN)',
            not com(tr3, '2_2s|ACROSS') and not com(tr3, 'R_4s') and not m3.get('R4S') and not d3.get('M'),
            {'trace': tr3, 'R4S': sorted(m3.get('R4S', []))})

        # ---- W04
        m4, tr4, na4 = nov['C04']
        if q.bio:
            ok4 = not com(tr4, '2_2s|WITHIN_RUN') and '2_2s|WITHIN_RUN=2' in na4 and not com(tr4, '2_2s|ACROSS')
        else:
            ok4 = not com(tr4, '2of3_2s|WITHIN_RUN') and '2of3_2s|WITHIN_RUN=3' in na4
        reg('W04 (c) Regra ENTRE níveis continua pareando pelo RUN: níveis em corridas diferentes não formam o 2_2s/2of3_2s '
            'WITHIN_RUN (não avaliável por corrida)', ok4, {'trace': tr4, 'nao_avaliaveis': na4})

        # ---- W05
        m5, tr5, na5 = nov['C05']
        if q.bio:
            ok5 = not com(tr5, '4_1s|N2_R2') and '4_1s|N2_R2=2' in na5 and not m5.get('R41')
        else:
            ok5 = not com(tr5, '6x|N3_R2') and '6x|N3_R2=2' in na5 and not m5.get('SEQ')
        reg('W05 (c) NxR oficial NÃO compacta: a corrida incompleta deixa o 4_1s N2_R2 (Bio) / 6x N3_R2 (Hema) não '
            'avaliável, como antes', ok5, {'trace': tr5, 'nao_avaliaveis': na5})

        # ---- W06
        m6, tr6, _ = nov['C06']
        ma6, tra6, _ = ant['C06']
        if q.bio:
            l6 = com(tr6, '8x|SAME_LEVEL_R8|WITHIN_MATERIAL_ACROSS_RUN|COMPLEMENTAR|run=11|runIni=1|niveis=1|')
            ok6 = (len(l6) == 1 and l6[0].endswith('N1 runs 1..11 (8 obs)') and not m6.get('SEQ')
                   and not com(tr6, '8x|N2_R4') and not com(tra6, '8x|SAME_LEVEL_R8'))
            ev6 = {'trace_novo': com(tr6, '8x'), 'SEQ': sorted(m6.get('SEQ', [])), 'trace_antigo': com(tra6, '8x')}
        else:
            l6 = com(tr6, '3_1s|SAME_LEVEL_R3|WITHIN_MATERIAL_ACROSS_RUN|COMPLEMENTAR|run=5|runIni=1|niveis=1|')
            m6b, tr6b, _ = nov['C06b']
            ma6b, tra6b, _ = ant['C06b']
            l6b = com(tr6b, '6x|SAME_LEVEL_R6|WITHIN_MATERIAL_ACROSS_RUN|COMPLEMENTAR|run=11|runIni=1|niveis=1|')
            todas = [m6.get(k) for k in ('R13', 'R2', 'R4S', 'R41', 'SEQ')] + [m6b.get(k) for k in ('R13', 'R2', 'R4S', 'R41', 'SEQ')]
            ok6 = (len(l6) == 1 and l6[0].endswith('N1 runs 1..5 (3 obs)') and not any(todas)
                   and not com(tr6, '3_1s|N3_R1') and not com(tra6, '3_1s|SAME_LEVEL_R3')
                   and len(l6b) == 1 and l6b[0].endswith('N1 runs 1..11 (6 obs)')
                   and not any('niveis=1|' in l for l in com(tra6b, '6x|SAME_LEVEL_R6')))
            ev6 = {'trace_3_1s': com(tr6, '3_1s'), 'trace_6x_N1': [l for l in com(tr6b, '6x') if 'niveis=1|' in l],
                   'marcas': todas, 'trace_antigo': com(tra6, '3_1s') + [l for l in com(tra6b, '6x') if 'niveis=1|' in l]}
        reg('W06 Complementar do mesmo nível com lacunas entra no trace como COMPLEMENTAR com " (n obs)" e NÃO marca '
            '(veredicto inalterado); o motor antigo não o via', ok6, ev6)

        # ---- W07
        # ADR-083: com minimo < n a observacao que FECHA a janela tem de passar do limite, do lado que disparou -- sem
        # isso +2,3 +2,4 +0,5 dava um 2o evento fechando dentro do controle e -2,3 -2,4 +2,3 um "abaixo" no lado oposto
        d7b, dt7b = det('SEQ2DE3', {(0, 1): 0.1, (0, 2): 2.3, (0, 3): 2.4, (0, 4): 0.5}, 4)
        l7b = com(dt7b, 'QA083|QA083_SEQ2DE3|')
        d7c, dt7c = det('SEQ2DE3', {(0, 1): -2.3, (0, 2): -2.4, (0, 3): 2.3}, 3)
        l7c = com(dt7c, 'QA083|QA083_SEQ2DE3|')
        fecha7 = len(l7b) == 1 and '|run=3|' in l7b[0] and not l7c and not d7b.get('M') and not d7c.get('M')
        ev_fecha7 = {'mais0,1_2,3_2,4_0,5': l7b, 'menos2,3_2,4_mais2,3': l7c}
        if not q.bio:
            m7, tr7, _ = nov['C07']
            ma7, tra7, _ = ant['C07']
            l7 = com(tr7, '2of3_2s|SAME_LEVEL_R3|WITHIN_MATERIAL_ACROSS_RUN|COMPLEMENTAR|run=3|runIni=1|niveis=1|')
            ok7 = (len(l7) == 1 and l7[0].endswith('N1 runs 1..3') and not m7.get('R2')
                   and not com(tra7, '2of3_2s|SAME_LEVEL_R3'))
            reg('W07 Hema: 2of3_2s SAME_LEVEL_R3 = 2 ou mais de 3 (+2,3 +0,5 +2,4): 1 evento COMPLEMENTAR, sem marca; o '
                'motor antigo (3 de 3) não o via; a observação de fechamento passa do limite (+0,1 +2,3 +2,4 +0,5 = 1 '
                'evento; −2,3 −2,4 +2,3 = 0)', ok7 and fecha7,
                {'trace_novo': com(tr7, '2of3'), 'trace_antigo': com(tra7, '2of3'), 'fechamento': ev_fecha7})
        else:
            d7, dt7 = det('SEQ2DE3', {(0, 1): 2.3, (0, 2): 0.5, (0, 3): 2.4}, 3)
            ok7 = len(com(dt7, 'QA083|QA083_SEQ2DE3|WITHIN_MATERIAL_ACROSS_RUN|COMPLEMENTAR|run=3|runIni=1|niveis=1|')) == 1
            reg('W07 Bio: o detector compartilhado "2 ou mais de 3" (usado pelo 2of3_2s SAME_LEVEL_R3 da Hema) dispara com '
                '+2,3 +0,5 +2,4; a observação de fechamento passa do limite (+0,1 +2,3 +2,4 +0,5 = 1 evento; −2,3 −2,4 '
                '+2,3 = 0)', ok7 and fecha7, {'trace': dt7, 'fechamento': ev_fecha7})

        # ---- W08
        m8, tr8, _ = nov['C08']
        pre8 = '8x|SAME_LEVEL_R8' if q.bio else '6x|SAME_LEVEL_R6'
        reg('W08 Contadores zerados a cada janela: uma observação do outro lado no meio impede toda janela que a contém '
            f'({pre8} do N1 não dispara)', not [l for l in com(tr8, pre8) if 'niveis=1|' in l],
            {'trace': com(tr8, pre8)})

        # ---- W09
        m9, tr9, na9 = nov['C09']
        ma9, tra9, naa9 = ant['C09']
        entre = ('R_4s|', '2_2s|WITHIN_RUN', '2of3_2s|WITHIN_RUN', '4_1s|N2_R2', '8x|N2_R4', '3_1s|N3_R1', '6x|N3_R2',
                 '1_3s|')
        t9n = sorted(l for l in tr9 if l.startswith(entre))
        t9a = sorted(l for l in tra9 if l.startswith(entre))
        chaves9 = ('R13', 'R4S', 'A12') + (('R41', 'SEQ') if q.bio else ('R41', 'SEQ', 'R2'))
        mk9 = {k: (sorted(m9.get(k, [])), sorted(ma9.get(k, []))) for k in chaves9}
        reg('W09 (c) R_4s, 2_2s/2of3_2s WITHIN_RUN, NxR e 1_3s idênticos nos dois motores numa série mista (trace e marcas)',
            t9n == t9a and bool(t9n) and all(a == b for a, b in mk9.values()),
            {'trace_entre_niveis_novo': t9n, 'igual_antigo': t9n == t9a, 'marcas_novo_x_antigo': mk9})

        # ---- W10
        s10a = run('mEstatistica.QA083_Seq', itens({(0, 1): 0.1, (0, 2): 2.3, (0, 4): 2.3, (1, 3): 0.4}), 4, 0)
        s10b = run('mEstatistica.QA083_Seq', itens({(0, 1): 0.1, (0, 2): 2.3, (0, 4): 2.3, (1, 3): 0.4}), 4, 1)
        s10c = run('mEstatistica.QA083_Seq', '', 3, 0)
        reg('W10 SeqDoNivel devolve as posições avaliáveis do nível, em ordem (vazio sem dado)',
            s10a == 'OK|1,2,4' and s10b == 'OK|3' and s10c == 'OK|', {'N1': s10a, 'N2': s10b, 'vazio': s10c})

        # ---- W11 (f)
        def canon(saida_tr):
            m_, tr_ = saida_tr
            linhas = []
            for l in tr_:
                c = l.split('|')
                linhas.append('|'.join([c[0], c[1], c[2]] + c[4:7] + [c[-1].replace(',', '.')]))
            return {'M': sorted(m_.get('M', [])), 'trace': linhas}
        f_obtido = {
            'across': canon(det('2_2S_ACROSS', {(0, 1): 0.1, (0, 2): 2.3, (0, 4): 2.3, (1, 3): 0.4}, 4)),
            'lado4': canon(det('LADO4', {(0, 1): 0.5, (0, 2): 0.6, (0, 4): 0.7, (0, 6): 0.8, (1, 3): -1.0}, 6)),
            'seq2de3': canon(det('SEQ2DE3', {(0, 1): 2.3, (0, 3): 0.5, (0, 5): 2.4, (1, 2): 0.1}, 5)),
        }
        f_esperado = {
            'across': {'M': [(0, 4)], 'trace': ['2_2s|ACROSS_RUN_SAME_LEVEL|WITHIN_MATERIAL_ACROSS_RUN|run=4|runIni=2|'
                                                'niveis=1|N1 2.30->2.30']},
            'lado4': {'M': [], 'trace': ['QA083|QA083_LADO4|WITHIN_MATERIAL_ACROSS_RUN|run=6|runIni=1|niveis=1|'
                                         'N1 runs 1..6 (4 obs)']},
            'seq2de3': {'M': [], 'trace': ['QA083|QA083_SEQ2DE3|WITHIN_MATERIAL_ACROSS_RUN|run=5|runIni=1|niveis=1|'
                                           'N1 runs 1..5 (3 obs)']},
        }
        pb = procs_de(open(os.path.join(SRC, 'bio', 'mEstatistica.bas'), encoding='utf-8').read())
        ph = procs_de(open(os.path.join(SRC, 'hema', 'mEstatistica.bas'), encoding='utf-8').read())
        dif_txt = [n for n in COMPARTILHADOS if pb.get(n) is None or pb.get(n) != ph.get(n)]
        f_cmp = {k: (f_obtido[k] == f_esperado[k]) for k in f_esperado}
        reg('W11 (f) Bio = Hema: os detectores do mesmo nível chamados direto dão a MESMA saída canônica nas duas áreas, e '
            'o texto dos procedimentos compartilhados (motor, SeqDoNivel, trilha) é idêntico em src/bio e src/hema',
            all(f_cmp.values()) and not dif_txt,
            {'saida_canonica': {k: {'obtido': f_obtido[k], 'ok': f_cmp[k]} for k in f_obtido},
             'procedimentos_diferentes': dif_txt})

        # ======================================================== Parte 2: lote real, antigo x novo
        reais_nov, flags_nov, an_tela_nov = capturar_lotes()
        viol, novas_fora, entre_dif, of_dif = {}, {}, {}, {}
        for L in lotes_reais:
            A, D = reais_ant.get(L, []), reais_nov.get(L, [])
            falta = contagem(A) - contagem(D)
            if falta:
                viol[L] = [list(k) for k in list(falta)[:3]]
            sobra = contagem(D) - contagem(A)
            fora = {(k[C_REGRA - 1], k[C_DETECTOR - 1]) for k in sobra} - MESMO_NIVEL[area]
            if fora:
                novas_fora[L] = sorted(fora)
            ent_a = contagem([r for r in A if (txt(r[C_REGRA - 1]), txt(r[C_DETECTOR - 1])) not in MESMO_NIVEL[area]])
            ent_d = contagem([r for r in D if (txt(r[C_REGRA - 1]), txt(r[C_DETECTOR - 1])) not in MESMO_NIVEL[area]])
            if ent_a != ent_d:
                entre_dif[L] = {'so_antigo': len(ent_a - ent_d), 'so_novo': len(ent_d - ent_a)}
            oa, od = oficiais(A), oficiais(D)
            of_dif[L] = {'oficiais_antes': len(oa), 'oficiais_depois': len(od), 'removidos': sorted(oa - od)[:3],
                         'novos': len(od - oa), 'regras_novas': sorted({k[2] + '/' + k[3] for k in od - oa})}
        n_linhas = {L: (len(reais_ant.get(L, [])), len(reais_nov.get(L, []))) for L in lotes_reais}
        reg('W12 (b) Superconjunto num lote real: toda linha antiga de Eventos_Westgard (14 colunas) continua no motor novo; '
            'as linhas novas são só de detector do MESMO nível',
            not viol and not novas_fora and any(a for a, _ in n_linhas.values()),
            {'lotes': lotes_reais, 'linhas_antigo_novo': n_linhas, 'antigas_que_sumiram': viol,
             'novas_fora_do_mesmo_nivel': novas_fora})
        reg('W13 (c) Detectores ENTRE níveis (R_4s, WITHIN_RUN, N2_R2/N2_R4/N3_R1/N3_R2) e 1_3s: linhas idênticas nos dois '
            'motores', not entre_dif, {'diferencas': entre_dif})
        if q.bio:
            so_22 = all(set(v['regras_novas']) <= {'2_2s/ACROSS_RUN_SAME_LEVEL'} and not v['removidos']
                        for v in of_dif.values())
            # flags do analito em tela: so o campo 2 (R22s) e o 7 (Veredicto) podem mudar
            fl_ok, fl_ev = True, []
            for r_, fa in flags_ant.items():
                fn = flags_nov.get(r_)
                if fn is None:
                    continue
                for t in range(nlv):
                    for c_ in range(NEF):
                        if fa[t * NEF + c_] != fn[t * NEF + c_] and c_ not in (1, 6):
                            fl_ok = False
                            fl_ev.append((r_, t + 1, c_ + 1, fa[t * NEF + c_], fn[t * NEF + c_]))
            reg('W14 (d) Bio, lote real: entre os eventos OFICIAIS só entram 2_2s ACROSS (nenhum sai); no Eng_Saida do '
                'analito em tela só mudam R22s e Veredicto', so_22 and fl_ok and an_tela_ant == an_tela_nov,
                {'oficiais_por_lote': of_dif, 'flags_fora_de_R22s_Veredicto': fl_ev[:6], 'analito_em_tela': an_tela_nov})
        else:
            iguais_of = all(not v['novos'] and not v['removidos'] for v in of_dif.values())
            iguais_fl = flags_ant == flags_nov
            reg('W14 (d) Hema, lote real: NENHUM evento OFICIAL muda (todos os SAME_LEVEL são complementares) e o bloco de '
                'flags do Eng_Saida do analito em tela é idêntico', iguais_of and iguais_fl and bool(flags_nov),
                {'oficiais_por_lote': of_dif, 'flags_iguais': iguais_fl, 'analito_em_tela': an_tela_nov,
                 'corridas': len(flags_nov)})

        # ---- W15 trilha (unidade): foto, muda um evento oficial de RUN, audita
        ws_ev = q.ws(ABA_EV)

        def refazer15():
            run('mEstatistica.InvalidarCache')
            run('mEstatistica.RegistrarEventosWestgard', teto=900)

        def mexer15():
            """Leva o 1o evento OFICIAL da aba (ou um sintetico, se o lote nao tiver) para o RUN 99999999, que nao e
            corrida do eixo. Devolve (linha, evento, RUN original ou None)."""
            linhas_ = eventos(q)
            n_ = len(linhas_)
            ws_ev.Unprotect(ii.SENHA)
            r_ = next((i for i, r in enumerate(linhas_) if txt(r[C_CLASSE - 1]) == 'OFICIAL'), None)
            if r_ is not None:
                ws_ev.Cells(4 + r_, C_RUN).Value = 99999999
                return r_, linhas_[r_], iint(linhas_[r_][C_RUN - 1])
            alvo_ = ('', 26010101.0, 'QA083', '1', '1_3s', 'INDIVIDUAL', 'WITHIN_RUN', 'OFICIAL', 1.0, 1.0,
                     26010101.0, 'QA', 'Erro', 3.5)              # lote sem evento oficial: acrescenta um (so entra)
            ws_ev.Range(ws_ev.Cells(4 + n_, 1), ws_ev.Cells(4 + n_, 14)).Value = (tuple(alvo_),)
            ws_ev.Cells(4 + n_, C_RUN).Value = 99999999
            ws_ev.Range('J2').Value = n_ + 1
            return None, alvo_, None

        ws_ev.Unprotect(ii.SENHA)
        refazer15()
        foto = str(run('mEstatistica.QA083_Auditoria', 'FOTO'))
        a0 = audit_ult(q)
        r_of, alvo15, run_orig = mexer15()
        r_aud = str(run('mEstatistica.QA083_Auditoria', 'AUDITAR'))
        aud15 = audit_desde(q, a0)
        niv15 = txt(alvo15[C_NIVEIS - 1])
        nv_esp = iint(niv15) if ',' not in niv15 else None
        an15 = txt(alvo15[C_ANALITO - 1])
        novo15 = [x for x in aud15 if x['acao'] == 'WESTGARD_EVENTO_CORRIDA_NOVA']
        rem15 = [x for x in aud15 if x['acao'] == 'WESTGARD_EVENTO_REMOVIDO']
        lote15 = txt(ws_ev.Range('L2').Value)

        def bate(x, run_, classe):
            return (x['run'] == run_ and x['analito'] == an15 and x['lote'] == lote15
                    and (x['nivel'] == nv_esp or (nv_esp is None and x['nivel'] in (None, 0)))
                    and x['motivo'].startswith(txt(alvo15[C_REGRA - 1]) + ' ' + txt(alvo15[C_DETECTOR - 1]))
                    and x['parecer'].startswith(f'ADR-083 [{classe}]') and not x['chave'] and not x['seq'])
        ok15a = ('#RUNS' in foto and r_aud == 'OK|' and len(novo15) == 1
                 and bate(novo15[0], 99999999, 'CORRIDA_NOVA') and novo15[0]['st_novo'] == 'OFICIAL'
                 and (run_orig is None and not rem15 or len(rem15) == 1 and bate(rem15[0], run_orig, 'RETROATIVO')
                      and rem15[0]['st_ant'] == 'OFICIAL') and len(aud15) == (1 if run_orig is None else 2))
        # escopo diferente: nada auditado
        run('mEstatistica.QA083_Auditoria', 'FOTO')
        ws_ev.Unprotect(ii.SENHA)
        ws_ev.Range('L2').Value = 'QA083_OUTRO_LOTE'
        if r_of is not None:
            ws_ev.Cells(4 + r_of, C_RUN).Value = run_orig
        a1 = audit_ult(q)
        run('mEstatistica.QA083_Auditoria', 'AUDITAR')
        aud15b = audit_desde(q, a1)
        # foto sem a referencia das corridas avaliadas (o motor ainda nao rodou na sessao): NOVO / NAO_CLASSIFICADO
        refazer15()
        run('mEstatistica.QA083_Auditoria', 'FOTO')
        run('mEstatistica.QA083_Auditoria', 'SEMREF')
        a2 = audit_ult(q)
        mexer15()
        run('mEstatistica.QA083_Auditoria', 'AUDITAR')
        aud15c = audit_desde(q, a2)
        nov15c = [x for x in aud15c if x['acao'] == 'WESTGARD_EVENTO_NOVO']
        ok15c = (len(nov15c) == 1 and bate(nov15c[0], 99999999, 'NAO_CLASSIFICADO')
                 and not [x for x in aud15c if x['acao'] == 'WESTGARD_EVENTO_CORRIDA_NOVA'])
        # J2 vazio = aba em reconstrucao (RegistrarEventosWestgard parou no meio): foto vazia, nada auditado
        refazer15()
        run('mEstatistica.QA083_Auditoria', 'FOTO')
        mexer15()
        ws_ev.Unprotect(ii.SENHA)
        ws_ev.Range('J2').ClearContents()
        a3 = audit_ult(q)
        run('mEstatistica.QA083_Auditoria', 'AUDITAR')
        aud15d = audit_desde(q, a3)
        foto_vazia = str(run('mEstatistica.QA083_Auditoria', 'FOTO')) == 'OK|'
        refazer15()
        reg('W15 Trilha (unidade): EventosWestgardOficiais + AuditarMudancasWestgard -- evento OFICIAL levado para um RUN '
            'fora do eixo = 1 WESTGARD_EVENTO_CORRIDA_NOVA + 1 WESTGARD_EVENTO_REMOVIDO [RETROATIVO] com RUN, analito, '
            'lote, nível e regra, sem AU_CHAVE/AU_SEQ; sem a referência = NOVO [NAO_CLASSIFICADO]; J2 vazio = nada; '
            'foto de outro lote/equipamento = nada auditado', ok15a and not aud15b and ok15c and not aud15d and foto_vazia,
            {'foto': foto[:300], 'evento_alterado': [norm_cel(x) for x in alvo15[:8]], 'audit': aud15,
             'audit_escopo_diferente': aud15b, 'audit_sem_referencia': aud15c, 'audit_j2_vazio': aud15d,
             'foto_com_j2_vazio_vazia': foto_vazia})

        # ======================================================== Parte 3: ponta a ponta
        motor.usar('original')
        run('mEstatistica.InvalidarCache')
        cadastro = [txt(v[0]) for v in q.espec().Range('A4:A43').Value]
        lote = txt(run('mLotes.LotePainel'))
        equip = txt(run('mEstatistica.EquipFiltro')).upper()
        an_sel = txt(q.nome('selAnalito').Value)
        gap = float(q.cfg('GAP_CORRIDA_MIN') or 60)
        tol = float(q.cfg('TOLERANCIA_CONFLITO_MIN') or 0)

        def params(an, nv):
            m_ = run('mLotes.MediaDoLote', lote, an, nv)
            s_ = run('mLotes.DPDoLote', lote, an, nv)
            return (float(m_), float(s_)) if num(m_) and num(s_) and float(s_) > 0 else None

        da = collections.defaultdict(set)
        cont_an = collections.Counter()
        equip_dado = {}
        for r in fin:
            a_ = txt(r['ANALITO']).upper()
            if hasattr(r['DATA_HORA'], 'year'):
                da[a_].add(r['DATA_HORA'].date())
            if txt(r['LOTE']) == lote and r['PARTICIPA_ESTATISTICA'] == 'SIM' and iint(r['NIVEL']) == 1 \
                    and (not equip or txt(r['EQUIPAMENTO']).upper() == equip):
                cont_an[a_] += 1
                equip_dado[a_] = r['EQUIPAMENTO']
        cand = [a for a in cadastro if a and cont_an[a.upper()] >= 5 and params(a, 1)]
        cand.sort(key=lambda a: (a.upper() != an_sel.upper(), -cont_an[a.upper()]))
        if not cand:
            raise RuntimeError(f'pré-condição ausente: nenhum analito do lote {lote} com média/DP do N1 e >= 5 resultados')
        an = cand[0]
        par = {nv: params(an, nv) for nv in range(1, nlv + 1)}
        hoje = dt.date.today()
        dias_lote = sorted(r['DATA_HORA'].date() for r in fin if txt(r['ANALITO']).upper() == an.upper()
                           and txt(r['LOTE']) == lote and hasattr(r['DATA_HORA'], 'year'))
        ocup = da[an.upper()]
        livres = []
        d_ = min(dias_lote[-1] + dt.timedelta(days=30), hoje - dt.timedelta(days=1))
        while d_ >= dias_lote[0] - dt.timedelta(days=400) and len(livres) < 2:
            if d_ not in ocup:
                livres.append(d_)
            d_ -= dt.timedelta(days=1)
        if len(livres) < 2:
            raise RuntimeError('pré-condição ausente: dois dias livres para o analito')
        D1, D2 = livres[0], livres[1]
        passo = max(gap, tol) + 30
        horas = [dt.time(6, 0)]
        for _ in range(2):
            t_ = dt.datetime.combine(D1, horas[-1]) + dt.timedelta(minutes=passo)
            horas.append(t_.time())
        if horas[-1] <= horas[0] or horas[-1] > dt.time(23, 0):
            raise RuntimeError(f'pré-condição ausente: GAP_CORRIDA_MIN = {gap} não cabe três corridas num dia')
        eq_man = equip_dado.get(an.upper()) or (equip or str(q.cfg('EQUIPAMENTO_PADRAO') or ''))

        def valor(nv, zz):
            pm = par.get(nv)
            if pm:
                return round(pm[0] + zz * pm[1], 6)
            vals = sorted(r['RESULTADO'] for r in fin if txt(r['ANALITO']).upper() == an.upper()
                          and iint(r['NIVEL']) == nv and num(r['RESULTADO']))
            return vals[len(vals) // 2] if vals else 1.0

        plano = []                                       # (rotulo, dia, hora, nivel, z)
        outros = list(range(2, nlv + 1))
        for (rot, dd_, h, z1) in (('A2', D2, horas[0], -2.3), ('M2', D2, horas[1], None), ('C2', D2, horas[2], -2.4),
                                  ('A', D1, horas[0], 2.3), ('B', D1, horas[1], 0.5), ('C', D1, horas[2], 2.4)):
            if z1 is not None:
                plano.append((rot, dd_, h, 1, z1))
            for nv in outros:
                plano.append((rot, dd_, h, nv, -0.3))
        # Painel: analito do teste e periodo ate D1 (o Calc/Eng_Saida mostram as corridas do teste)
        q.ex.xl.EnableEvents = False
        try:
            p.Unprotect(ii.SENHA)
            p.Range('B3').Value = cadastro.index(an) + 1
            p.Range('G3').Value = serial(min(D1, D2) - dt.timedelta(days=120))
            p.Range('G4').Value = serial(D1)
        finally:
            q.ex.xl.EnableEvents = True
        ids = {}
        q.ex.xl.EnableEvents = True
        for rot, dd_, h, nv, zz in plano:
            lin = q.escrever_linha(TB_MAN, {
                'DATA': dt.datetime.combine(dd_, dt.time()), 'HORA': (h.hour * 60 + h.minute) / 1440.0,
                'EQUIPAMENTO': eq_man, 'LOTE': lote, 'NIVEL': nv, 'ANALITO': an, 'RESULTADO': valor(nv, zz),
                'MOTIVO': MOTIVO_MAN})
            ids[(rot, nv)] = txt(q.lo(TB_MAN).ListColumns('ID_REGISTRO').DataBodyRange.Cells(lin, 1).Value)
        q.atualizar()
        f0 = {txt(r['ID_REGISTRO']): r for r in q.ler('tblCQ_Final')}
        run_de = {}
        ruins20 = []
        for (rot, nv), i_ in ids.items():
            r = f0.get(i_)
            if r is None or r['PARTICIPA_ESTATISTICA'] != 'SIM':
                ruins20.append((rot, nv, i_, None if r is None else r['PARTICIPA_ESTATISTICA']))
                continue
            run_de.setdefault(rot, set()).add(iint(r['RUN']))
        estrutura = {k: sorted(v) for k, v in run_de.items()}
        rr = {k: (list(v)[0] if len(v) == 1 else None) for k, v in run_de.items()}
        ok20 = (not ruins20 and all(rr.get(k) for k in ('A2', 'M2', 'C2', 'A', 'B', 'C'))
                and rr['A2'] < rr['M2'] < rr['C2'] and rr['A'] < rr['B'] < rr['C'])
        reg('W20 Pré-condição ponta a ponta: manuais em dois dias livres viram as corridas planejadas (D2: N1 −2,3 | só os '
            'outros níveis | N1 −2,4; D1: N1 +2,3 | +0,5 | +2,4), todos ATIVOS', ok20,
            {'analito': an, 'lote': lote, 'equip': eq_man, 'D1': D1, 'D2': D2, 'horas': [str(h) for h in horas],
             'RUN_por_corrida': estrutura, 'nao_ativos': ruins20, 'parametros': par})
        if not ok20:
            raise RuntimeError('cenário ponta a ponta não montou as corridas planejadas')

        AN = an.upper()
        ev0 = eventos(q)
        # so o analito do teste: resultado novo de outro analito (ATUALIZAR no modo PQ le o DB_SEAC) nao interfere
        of0 = {k for k in oficiais(ev0) if k[0] == AN}
        # ---- W21 (a) -- o controle com o motor antigo roda no fim (W25), na mesma base
        k21 = (AN, '1', '2_2s', 'ACROSS_RUN_SAME_LEVEL', rr['C2'], rr['A2'])
        comp21 = [r for r in ev0 if txt(r[C_ANALITO - 1]).upper() == AN and txt(r[C_REGRA - 1]) == '2of3_2s'
                  and txt(r[C_DETECTOR - 1]) == 'SAME_LEVEL_R3' and iint(r[C_RUN - 1]) == rr['C2']
                  and txt(r[C_NIVEIS - 1]) == '1']
        of_d2 = sorted(k for k in of0 if k[0] == AN and k[4] in (rr['A2'], rr['M2'], rr['C2']))
        if q.bio:
            ok21 = k21 in of0
            ev21 = {'esperado': k21, 'oficiais_nas_corridas_D2': of_d2}
        else:
            ok21 = bool(comp21) and not [k for k in of_d2 if k[1] == '1' and k[4] == rr['C2']]
            ev21 = {'complementar_2of3': [[norm_cel(x) for x in r] for r in comp21], 'oficiais_nas_corridas_D2': of_d2}

        # ---- W22 (e) inativar o N1 intermediario de D1
        idB = ids[('B', 1)]
        a22 = audit_ult(q)
        r22 = str(run('mIntegracao.InativarPorId', idB, 'QA ADR-083: sequência de ativos do nível', True, ''))
        s22 = str(run('mIntegracao.AtualizarSoEntradas', an, teto=1200))
        ev1_ = eventos(q)
        of1 = {k for k in oficiais(ev1_) if k[0] == AN}
        aud22 = [x for x in audit_desde(q, a22) if x['acao'].startswith('WESTGARD_') and x['analito'].upper() == AN]
        k22 = (AN, '1', '2_2s', 'ACROSS_RUN_SAME_LEVEL', rr['C'], rr['A'])
        if q.bio:
            ok22 = (r22.startswith('OK|') and s22.startswith('OK|') and of1 - of0 == {k22} and not (of0 - of1)
                    and any(x['acao'] == 'WESTGARD_EVENTO_NOVO' and x['run'] == rr['C'] and x['nivel'] == 1
                            and x['analito'].upper() == AN and x['motivo'] == '2_2s ACROSS_RUN_SAME_LEVEL'
                            and x['parecer'].startswith('ADR-083 [RETROATIVO]') and not x['chave'] for x in aud22))
        else:
            ok22 = r22.startswith('OK|') and s22.startswith('OK|') and of1 == of0 and not aud22
        reg('W22 (e) Inativar (InativarPorId + AtualizarSoEntradas) o N1 intermediário entre dois N1 > +2DP: Bio ganha '
            'EXATAMENTE o 2_2s ACROSS (RUN da 3ª corrida, início na 1ª) e o Audit_Log tem WESTGARD_EVENTO_NOVO; Hema: '
            'nenhum OFICIAL muda e nada é auditado', ok22,
            {'ID_inativado': idB, 'retornos': [r22[:80], s22[:80]], 'oficiais_novos': sorted(of1 - of0),
             'oficiais_removidos': sorted(of0 - of1), 'esperado_bio': k22, 'audit_westgard': aud22,
             'mIntegracao_chama_AuditarFotosWestgard': integra})

        # ---- W24 (g) coerencia com o N1 inativado
        e1 = q.eng_saida()
        calc = q.ws('Calc')
        bcol = [iint(v[0]) for v in calc.Range('B3:B182').Value]
        marc1 = collections.defaultdict(set)
        for k in of1:
            if k[0] == AN:
                for nv in k[1].split(','):
                    marc1[(int(nv), k[4])].add(k[2])
        inc24 = []
        fora24 = []
        n_calc24 = 0
        for rot in ('A', 'B', 'C', 'A2', 'M2', 'C2'):
            r_ = rr[rot]
            if r_ not in e1['runs']:
                if rot in ('A', 'B', 'C'):
                    inc24.append((rot, r_, 'fora do Eng_Saida'))
                else:                                   # D2 pode ficar antes da janela de 180 corridas do Calc
                    fora24.append((rot, r_))
                continue
            sl = e1['runs'].index(r_)
            for t in range(nlv):
                ver = txt(e1['flags'][sl][t * NEF + 6])
                tem_ponto = num(e1['val'][sl][t]) if t < len(e1['val'][sl]) else False
                esp = ('REJEITADO' if marc1.get((t + 1, r_)) else 'OK') if tem_ponto else ''
                if ver not in (esp, 'SEM PARAMETROS' if tem_ponto and not par.get(t + 1) else esp):
                    inc24.append((rot, r_, t + 1, 'Eng_Saida', ver, esp))
                r22s = iint(e1['flags'][sl][t * NEF + 1])
                esp22 = 1 if '2_2s' in marc1.get((t + 1, r_), set()) or '2of3_2s' in marc1.get((t + 1, r_), set()) else 0
                if tem_ponto and r22s != esp22 and ver != 'SEM PARAMETROS':
                    inc24.append((rot, r_, t + 1, 'R22s', r22s, esp22))
                if r_ in bcol:
                    i_ = bcol.index(r_) + 3
                    vc = txt(calc.Cells(i_, 16 + 22 * t).Value)
                    n_calc24 += 1
                    if (vc or ver) and vc != ver and not (vc.upper() == 'ALERTA' and ver == 'OK'):
                        inc24.append((rot, r_, t + 1, 'Calc', vc, ver))
                elif rot in ('A', 'B', 'C') and t == 0:
                    # A/B/C sao as corridas mais novas do periodo do Painel: TEM de estar no Calc (senao nada foi visto)
                    inc24.append((rot, r_, 'fora do Calc (B3:B182)'))
        pn = q.nome('engPainel').Value
        rej_n1 = any(txt(e1['flags'][i][6]) == 'REJEITADO' for i in range(e1['n']) if e1['fil'][i] in (1, 1.0))
        pn_n1 = bool(num(pn[0][1]) and pn[0][1])
        pn_ok = pn_n1 and txt(pn[0][9]) == ('REJEITADO' if rej_n1 else 'OK')
        met = str(run('mEstatistica.MetricasWestgard', an, 1))
        n_of_n1 = sum(1 for k in of1 if k[0] == AN and '1' in k[1].split(','))
        met_ok = met.split('|')[0] == str(n_of_n1)
        bi = 'sem identidade de BI'
        try:
            if tem_identidade_bi(q.wb):
                run('mBI.AtualizarBIData', teto=1200)
                wsb = q.ws('BI_Data')
                ult = wsb.Cells(wsb.Rows.Count, 1).End(-4162).Row
                vb = wsb.Range(wsb.Cells(1, 1), wsb.Cells(ult, 37)).Value
                lin_bi = [r for r in vb if txt(r[8]).upper() == AN and mesmo_lote(txt(r[11]), lote)
                          and iint(r[14]) == rr['C'] and iint(r[13]) == 1]
                bi = {'linhas': len(lin_bi), 'R22s': [iint(r[31]) for r in lin_bi], 'veredicto': [txt(r[36]) for r in lin_bi]}
                # A BI_Data (mBI.FlagsDoMotor) agrupa nucleo|ANALITO SEM equipamento (limitacao anterior ao ADR-083, avisada
                # pelo instalar_adr083): com o analito em DOIS equipamentos no lote, a sequencia do N1 da BI intercala as
                # duas maquinas e o 2_2s pode divergir legitimamente do Painel (por equipamento) -- so evidencia, nao falha
                eqs_bi = {txt(r['EQUIPAMENTO']).upper() for r in q.ler('tblCQ_Final')
                          if txt(r['ANALITO']).upper() == AN and mesmo_lote(txt(r['LOTE']), lote)
                          and r['PARTICIPA_ESTATISTICA'] == 'SIM'}
                if len(eqs_bi) > 1:
                    bi = dict(bi, nota=f'analito em {len(eqs_bi)} equipamentos no lote: BI sem equipamento (so evidencia)')
                elif q.bio and not lin_bi:
                    inc24.append(('BI', rr['C'], 1, 'linha da corrida C ausente na BI_Data'))
                elif q.bio and not (iint(lin_bi[0][31]) == 1 and txt(lin_bi[0][36]) == 'REJEITADO'):
                    inc24.append(('BI', rr['C'], 1, bi))
        except Exception as e_:                          # noqa: BLE001 -- a BI e opcional
            bi = f'BI nao conferida: {e_!r:.120}'
        reg('W24 (g) Com o N1 inativado: Eng_Saida (R22s/Veredicto), Calc (veredicto), engPainel (N1) e MetricasWestgard '
            'coerentes com os eventos OFICIAIS de Eventos_Westgard nas corridas do teste (BI_Data se houver identidade)',
            not inc24 and pn_ok and met_ok,
            {'incoerencias': inc24[:10], 'D2_fora_da_janela': fora24, 'celulas_Calc_conferidas': n_calc24,
             'engPainel_N1_numerico': pn_n1, 'engPainel_N1': [norm_cel(x) for x in pn[0][:18]],
             'MetricasWestgard_N1': met, 'oficiais_N1_na_aba': n_of_n1, 'BI': bi})

        # ---- W23 (e) reativar
        a23 = audit_ult(q)
        r23 = str(run('mIntegracao.ReativarPorId', idB, 'QA ADR-083: fim do teste da sequência'))
        s23 = str(run('mIntegracao.AtualizarSoEntradas', an, teto=1200))
        of2 = {k for k in oficiais(eventos(q)) if k[0] == AN}
        aud23 = [x for x in audit_desde(q, a23) if x['acao'].startswith('WESTGARD_') and x['analito'].upper() == AN]
        if q.bio:
            ok23 = (r23.startswith('OK|') and s23.startswith('OK|') and of2 == of0
                    and any(x['acao'] == 'WESTGARD_EVENTO_REMOVIDO' and x['run'] == rr['C'] and x['nivel'] == 1
                            and x['analito'].upper() == AN and x['parecer'].startswith('ADR-083 [RETROATIVO]')
                            and not x['chave'] for x in aud23))
        else:
            ok23 = r23.startswith('OK|') and s23.startswith('OK|') and of2 == of0 and not aud23
        reg('W23 (e) Reativar: os eventos OFICIAIS voltam EXATAMENTE aos de antes da inativação (a regra é função do estado '
            'atual) e o Audit_Log tem WESTGARD_EVENTO_REMOVIDO (Bio)', ok23,
            {'retornos': [r23[:80], s23[:80]], 'iguais_a_antes': of2 == of0, 'so_agora': sorted(of2 - of0)[:3],
             'so_antes': sorted(of0 - of2)[:3], 'audit_westgard': aud23})

        # ---- W21 (fecho): o motor ANTIGO na mesma base nao ve o caso (a)
        motor.usar('antigo')
        run('mEstatistica.InvalidarCache')
        run('mEstatistica.RegistrarEventosWestgard', teto=900)
        ev_ant = eventos(q)
        of_ant = oficiais(ev_ant)
        comp_ant = [r for r in ev_ant if txt(r[C_ANALITO - 1]).upper() == AN and txt(r[C_REGRA - 1]) == '2of3_2s'
                    and txt(r[C_DETECTOR - 1]) == 'SAME_LEVEL_R3' and iint(r[C_RUN - 1]) == rr['C2']
                    and txt(r[C_NIVEIS - 1]) == '1']
        motor.usar('original')
        run('mEstatistica.InvalidarCache')
        run('mEstatistica.AtualizarEstatistica', teto=900)
        antes_nao = (k21 not in of_ant) if q.bio else not comp_ant
        ev21.update({'motor_antigo_ve': not antes_nao})
        reg('W21 (a) Ponta a ponta, D2 (N1 −2,3 | corrida só dos outros níveis | N1 −2,4): Bio = 2_2s ACROSS OFICIAL do N1 '
            'no RUN da 3ª corrida com RUN_Inicial da 1ª; Hema = 2of3_2s SAME_LEVEL_R3 COMPLEMENTAR e nenhum OFICIAL novo '
            'no N1; o motor do ADR-081 na mesma base não via', ok21 and antes_nao, ev21)

        # ======================================================== Parte 4: foto por lote x equipamento (opcao completa)
        # O motor 'original' esta em vigor e a ultima atualizacao (W23) deixou a foto em dia com a base.
        ws_f = q.ws(ABA_FOTOS) if ABA_FOTOS in [w.Name for w in q.wb.Worksheets] else None
        foto0, nctrl0 = (ler_foto(q.wb) if ws_f is not None else (None, 0))
        ok30 = (ws_f is not None and bool(ws_f.ProtectContents) and foto0 is not None and nctrl0 > 0 and integra)
        reg('W30 Pré-condição da opção completa: a aba Westgard_Fotos existe (instalar_adr083.py), está protegida e tem a '
            'foto completa (J2) com linhas CONTROLE; o mIntegracao chama AuditarFotosWestgard', ok30,
            {'aba': ws_f is not None, 'protegida': bool(ws_f.ProtectContents) if ws_f is not None else None,
             'linhas_controle': nctrl0, 'eventos_na_foto': sum(len(s) for s in (foto0 or {}).values()),
             'mIntegracao_chama_AuditarFotosWestgard': integra,
             'obs': 'o QA revela todas as abas ao abrir (qa_final.QA.abrir): a ocultacao e conferida no log do instalador'})
        if not ok30:
            raise RuntimeError('a cópia não tem a foto por lote x equipamento (rodar instalar_adr083.py nela antes)')

        motor_vba = str(q.cfg('MOTOR_DADOS') or '').strip().upper() == 'VBA'
        eq_hema = '' if q.bio else txt(run('mEstatistica.EquipFiltro')).upper()
        sel_lote0 = txt(q.nome('loteSel').Value)
        sel_eq0 = txt(q.nome('selEquipamento').Value) if q.bio else ''

        def painel_agora():
            """(lote e equipamento da aba Eventos_Westgard, loteSel, selEquipamento, lote do Eng_Saida)."""
            return (txt(ws_ev.Range('L2').Value), txt(ws_ev.Range('N2').Value).upper(), txt(q.nome('loteSel').Value),
                    txt(q.nome('selEquipamento').Value) if q.bio else '', txt(q.ws('Eng_Saida').Range('E1').Value))

        def tempos_wg():
            """(segundos, RECALC, N, MUDANCAS, texto) do westgard_fotos no TEMPOS da ultima linha do Sync_Log."""
            linhas_log = q.ler('tblSyncLog')
            tx = txt(linhas_log[-1].get('TEMPOS')) if linhas_log else ''
            m_ = re.search(r'westgard_fotos ([\d.,]+)s \[RECALC=(\d+)/(\d+) MUDANCAS=(\d+)', tx)
            if not m_:
                return None, None, None, None, tx[-220:]
            return (float(m_.group(1).replace(',', '.')), int(m_.group(2)), int(m_.group(3)), int(m_.group(4)),
                    tx[-220:])

        # ---- W31: lote L2 (nao e o do Painel) e os manuais num dia livre do analito
        fin2 = q.ler('tblCQ_Final')
        cont2 = collections.Counter()
        eq_raw2 = {}
        for r in fin2:
            if r['PARTICIPA_ESTATISTICA'] != 'SIM' or iint(r['NIVEL']) != 1:
                continue
            L_, a_, e_ = txt(r['LOTE']), txt(r['ANALITO']), txt(r['EQUIPAMENTO'])
            if not L_ or not a_ or mesmo_lote(L_, lote):
                continue
            if q.bio and not e_:
                continue
            if not q.bio and eq_hema and e_.upper() != eq_hema:
                continue
            cont2[(L_, e_.upper(), a_.upper())] += 1
            eq_raw2[(L_, e_.upper(), a_.upper())] = r['EQUIPAMENTO']
        nome_cad = {a.upper(): a for a in cadastro if a}

        def params2(L_, an_, nv):
            m_ = run('mLotes.MediaDoLote', L_, an_, nv)
            s_ = run('mLotes.DPDoLote', L_, an_, nv)
            return (float(m_), float(s_)) if num(m_) and num(s_) and float(s_) > 0 else None

        alvo2 = None
        for (L_, e_, a_), n_ in cont2.most_common():
            if n_ < 5 or a_ not in nome_cad:
                continue
            if params2(L_, nome_cad[a_], 1):
                alvo2 = (L_, e_, nome_cad[a_])
                break
        if alvo2 is None:
            raise RuntimeError(f'pré-condição ausente: nenhum outro lote (além de {lote}) com média/DP do N1 e >= 5 '
                               'resultados de um analito do cadastro')
        L2, EQ2, an2 = alvo2
        AN2 = an2.upper()
        EQ_ESC = EQ2 if q.bio else eq_hema              # o equipamento do escopo na foto (Hema: o do EquipFiltro)
        par2 = {nv: params2(L2, an2, nv) for nv in range(1, nlv + 1)}
        ocup2 = {r['DATA_HORA'].date() for r in fin2 if txt(r['ANALITO']).upper() == AN2 and hasattr(r['DATA_HORA'], 'year')}
        dias2 = sorted(r['DATA_HORA'].date() for r in fin2 if txt(r['ANALITO']).upper() == AN2
                       and mesmo_lote(txt(r['LOTE']), L2) and hasattr(r['DATA_HORA'], 'year'))
        d_ = min(dias2[-1] + dt.timedelta(days=30), hoje - dt.timedelta(days=1))
        D3 = None
        while d_ >= dias2[0] - dt.timedelta(days=400):
            if d_ not in ocup2:
                D3 = d_
                break
            d_ -= dt.timedelta(days=1)
        if D3 is None:
            raise RuntimeError('pré-condição ausente: um dia livre para o analito do lote L2')

        def valor2(nv, zz):
            pm = par2.get(nv)
            if pm:
                return round(pm[0] + zz * pm[1], 6)
            vals = sorted(r['RESULTADO'] for r in fin2 if txt(r['ANALITO']).upper() == AN2
                          and iint(r['NIVEL']) == nv and num(r['RESULTADO']))
            return vals[len(vals) // 2] if vals else 1.0

        # Bio: o 2_2s ACROSS nasce ao inativar o N1 do meio; Hema: o 1_3s some ao inativar o N1 (nenhuma regra OFICIAL
        # da Hematologia nasce de uma inativacao -- todas as do mesmo nivel sao complementares)
        plano2 = ((('A', horas[0], 2.3), ('B', horas[1], 0.5), ('C', horas[2], 2.4)) if q.bio
                  else (('E', horas[0], 3.5),))
        painel_antes = painel_agora()
        a31 = audit_ult(q)
        ids2 = {}
        q.ex.xl.EnableEvents = True
        for rot, h, z1 in plano2:
            for nv in range(1, nlv + 1):
                lin = q.escrever_linha(TB_MAN, {
                    'DATA': dt.datetime.combine(D3, dt.time()), 'HORA': (h.hour * 60 + h.minute) / 1440.0,
                    'EQUIPAMENTO': eq_raw2[alvo2[:2] + (AN2,)], 'LOTE': L2, 'NIVEL': nv, 'ANALITO': an2,
                    'RESULTADO': valor2(nv, z1 if nv == 1 else -0.3), 'MOTIVO': MOTIVO_MAN})
                ids2[(rot, nv)] = txt(q.lo(TB_MAN).ListColumns('ID_REGISTRO').DataBodyRange.Cells(lin, 1).Value)
        q.atualizar()
        f2 = {txt(r['ID_REGISTRO']): r for r in q.ler('tblCQ_Final')}
        run2, ruins31 = {}, []
        for (rot, nv), i_ in ids2.items():
            r = f2.get(i_)
            if r is None or r['PARTICIPA_ESTATISTICA'] != 'SIM':
                ruins31.append((rot, nv, i_, None if r is None else r['PARTICIPA_ESTATISTICA']))
                continue
            run2.setdefault(rot, set()).add(iint(r['RUN']))
        rr2 = {k: (list(v)[0] if len(v) == 1 else None) for k, v in run2.items()}
        aud31 = [x for x in audit_desde(q, a31) if x['acao'].startswith('WESTGARD_') and x['analito'].upper() == AN2
                 and mesmo_lote(x['lote'], L2)]
        estrutura31 = not ruins31 and all(rr2.get(k) for k, _, _ in plano2)
        if q.bio:
            estrutura31 = estrutura31 and rr2['A'] < rr2['B'] < rr2['C']
            ok31 = estrutura31
        else:
            ok31 = estrutura31 and any(
                x['acao'] == 'WESTGARD_EVENTO_CORRIDA_NOVA' and x['run'] == rr2['E'] and x['nivel'] == 1
                and x['motivo'] == '1_3s INDIVIDUAL' and x['parecer'].startswith('ADR-083 [CORRIDA_NOVA]')
                and x['equip'] == EQ_ESC and not x['chave'] for x in aud31)
        reg('W31 Manuais do lote L2 (que NÃO é o do Painel) viram as corridas planejadas, ATIVAS; Hema: o ATUALIZAR audita '
            'o 1_3s novo de L2 como rotina (WESTGARD_EVENTO_CORRIDA_NOVA [CORRIDA_NOVA], lote L2)', ok31,
            {'lote_painel': lote, 'L2': L2, 'equip': EQ2, 'analito': an2, 'dia': D3, 'RUN_por_corrida': rr2,
             'nao_ativos': ruins31, 'parametros': par2, 'audit_westgard_L2': aud31})
        if not estrutura31:
            raise RuntimeError('cenário da Parte 4 não montou as corridas planejadas')

        # ---- W32 (a)(d): inativar o resultado de L2
        id32 = ids2[('B', 1)] if q.bio else ids2[('E', 1)]
        a32 = audit_ult(q)
        r32 = str(run('mIntegracao.InativarPorId', id32, 'QA ADR-083: foto por lote x equipamento', True, ''))
        s32 = str(run('mIntegracao.AtualizarSoEntradas', an2, teto=1200))
        aud32 = [x for x in audit_desde(q, a32) if x['acao'].startswith('WESTGARD_')]
        t32 = tempos_wg()
        painel32 = painel_agora()

        def linha_l2(lst, acao, run_, motivo, classe):
            return [x for x in lst if x['acao'] == acao and x['run'] == run_ and x['nivel'] == 1
                    and x['analito'].upper() == AN2 and mesmo_lote(x['lote'], L2) and x['equip'] == EQ_ESC
                    and x['motivo'] == motivo and x['parecer'].startswith(f'ADR-083 [{classe}]') and not x['chave']
                    and not x['seq']]
        if q.bio:
            l32 = linha_l2(aud32, 'WESTGARD_EVENTO_NOVO', rr2['C'], '2_2s ACROSS_RUN_SAME_LEVEL', 'RETROATIVO')
        else:
            l32 = linha_l2(aud32, 'WESTGARD_EVENTO_REMOVIDO', rr2['E'], '1_3s INDIVIDUAL', 'RETROATIVO')
        ok32 = (r32.startswith('OK|') and s32.startswith('OK|') and len(l32) == 1 and t32[1] is not None
                and t32[1] >= 1 and painel32 == painel_antes)
        reg('W32 (a)(d) Inativar um resultado do lote L2, fora do Painel: o Audit_Log tem a linha certa com o lote L2 (Bio: '
            'WESTGARD_EVENTO_NOVO 2_2s ACROSS na 3ª corrida; Hema: WESTGARD_EVENTO_REMOVIDO 1_3s), [RETROATIVO], com '
            'equipamento, nível e analito e sem AU_CHAVE; o Sync_Log anota o lote recalculado; Eventos_Westgard e o '
            'Painel continuam no lote/equipamento de antes', ok32,
            {'ID_inativado': id32, 'retornos': [r32[:80], s32[:80]], 'linha_esperada': l32,
             'audit_westgard': aud32[:8], 'westgard_fotos_TEMPOS': t32, 'painel_antes': painel_antes,
             'painel_depois': painel32})

        # ---- W33 (b)(d): reativar
        a33 = audit_ult(q)
        r33 = str(run('mIntegracao.ReativarPorId', id32, 'QA ADR-083: fim do teste da foto'))
        s33 = str(run('mIntegracao.AtualizarSoEntradas', an2, teto=1200))
        aud33 = [x for x in audit_desde(q, a33) if x['acao'].startswith('WESTGARD_')]
        painel33 = painel_agora()
        if q.bio:
            l33 = linha_l2(aud33, 'WESTGARD_EVENTO_REMOVIDO', rr2['C'], '2_2s ACROSS_RUN_SAME_LEVEL', 'RETROATIVO')
        else:
            l33 = linha_l2(aud33, 'WESTGARD_EVENTO_NOVO', rr2['E'], '1_3s INDIVIDUAL', 'RETROATIVO')
        ok33 = r33.startswith('OK|') and s33.startswith('OK|') and len(l33) == 1 and painel33 == painel_antes
        reg('W33 (b)(d) Reativar o resultado de L2: o Audit_Log tem a linha inversa com o lote L2 (Bio: '
            'WESTGARD_EVENTO_REMOVIDO; Hema: WESTGARD_EVENTO_NOVO [RETROATIVO]); o Painel não muda', ok33,
            {'retornos': [r33[:80], s33[:80]], 'linha_esperada': l33, 'audit_westgard': aud33[:8],
             'painel_depois': painel33})

        # ---- W34 (c): ATUALIZAR sem mudanca
        a34 = audit_ult(q)
        s34 = str(run('mIntegracao.AtualizarSoEntradas', an2, teto=1200))
        aud34 = [x for x in audit_desde(q, a34) if x['acao'].startswith('WESTGARD_')]
        t34 = tempos_wg()
        if motor_vba:
            ok34 = s34.startswith('OK|') and not aud34 and t34[1] == 0 and t34[3] == 0 and t34[0] is not None
        else:
            # motor PQ: AtualizarSoEntradas e o ATUALIZAR normal e le o DB_SEAC -- resultado novo e rotina, nada mais
            ok34 = (s34.startswith('OK|') and t34[0] is not None
                    and all(x['acao'] == 'WESTGARD_EVENTO_CORRIDA_NOVA' for x in aud34))
        reg('W34 (c) ATUALIZAR sem mudança: nenhuma linha WESTGARD_* no Audit_Log e nenhum lote recalculado (RECALC=0, '
            'MUDANCAS=0 no TEMPOS do Sync_Log), com o tempo da foto anotado', ok34,
            {'motor_VBA': motor_vba, 'retorno': s34[:80], 'audit_westgard': aud34[:8], 'westgard_fotos_TEMPOS': t34,
             'painel': painel_agora() == painel_antes})

        # ---- W36: meta de desempenho (sem mudanca)
        reg('W36 Meta: sem mudança, a foto por lote x equipamento custa menos de 1 s no ATUALIZAR (Sync_Log, TEMPOS)',
            t34[0] is not None and t34[0] < 1.0 and (t34[1] == 0 or not motor_vba),
            {'segundos': t34[0], 'lotes_x_equip_com_resultado': t34[2], 'recalculados': t34[1]})

        # ---- W35 (e): a foto bate com a varredura do instalador (Seletor do Painel + RegistrarEventosWestgard)
        foto35, nctrl35 = ler_foto(q.wb)
        escopos = sorted((foto35 or {}).items(), key=lambda kv: -len(kv[1]))
        alvos35 = ([(eq_painel, lote)] if eq_painel else []) if q.bio else [(None, lote)]
        alvos35.append((EQ2 if q.bio else None, L2))
        for (L_, e_), _ in escopos:
            alvo_ = (e_ if q.bio else None, L_)
            if len(alvos35) >= 4:
                break
            if alvo_ not in alvos35 and (not q.bio or e_):
                alvos35.append(alvo_)
        dep35 = {}
        p.Unprotect(ii.SENHA)
        q.ex.xl.EnableEvents = False
        try:
            sel = Seletor(q.wb, q.bio)
            run('mEstatistica.InvalidarCache')
            try:
                for eq_, L_ in alvos35:
                    sel.por(eq_, L_)
                    run('mEstatistica.RegistrarEventosWestgard', teto=900)
                    lt_, _eqa, linhas_ = ler_eventos(q.wb)
                    LOTE_DA_ABA[(eq_, L_)] = lt_
                    dep35[(eq_, L_)] = linhas_
            finally:
                sel.devolver()
        finally:
            q.ex.xl.EnableEvents = True
        run('mEstatistica.InvalidarCache')
        run('mEstatistica.AtualizarEstatistica', teto=900)
        dif35, n35 = conferir_foto(foto35 or {}, dep35, q.bio, eq_hema)
        ws_f = q.ws(ABA_FOTOS)
        nf = int(float(txt(ws_f.Range('J2').Value) or 0))
        vf = ws_f.Range(ws_f.Cells(4, 1), ws_f.Cells(3 + max(nf, 1), 6)).Value if nf else []
        corr_l2 = set()
        for r in vf:
            if txt(r[0]) == 'CORRIDAS' and mesmo_lote(txt(r[1]), L2) and txt(r[2]).upper() == EQ_ESC \
                    and txt(r[3]) == AN2:
                corr_l2 |= {iint(x) for x in txt(r[4]).split(',') if x.strip()}
        corr_ok = all(rr2[k] in corr_l2 for k, _, _ in plano2)
        reg('W35 (e) A foto bate com a varredura do instalador (RegistrarEventosWestgard lote a lote pelo Seletor do Painel) '
            'no lote do Painel, em L2 e em até 2 outros lotes x equipamento; as linhas CORRIDAS de L2 têm as corridas do '
            'teste; o Painel volta ao lote de antes', not dif35 and n35 >= 2 and corr_ok and painel_agora() == painel_antes,
            {'alvos': alvos35, 'conferidos': n35, 'divergencias': dif35[:5], 'linhas_controle': nctrl35,
             'corridas_L2_na_foto': corr_ok, 'painel': painel_agora(), 'loteSel_inicial': sel_lote0,
             'selEquipamento_inicial': sel_eq0})
    finally:
        q.fechar(salvar=False)
        hash1 = sha256_arquivo(caminho)
        reg('W99 O arquivo de entrada não foi alterado (nada salvo)', hash0 == hash1,
            {'sha256_antes': hash0[:16], 'sha256_depois': hash1[:16]})


def main(produto, caminho, saida):
    try:
        executar(produto, caminho)
    except Exception as ex:                              # noqa: BLE001
        reg('EXECUCAO', False, {'erro': repr(ex)[:300], 'traceback': traceback.format_exc()[-1500:]})
    with open(saida, 'w', encoding='utf-8') as f:
        json.dump(RES, f, ensure_ascii=False, indent=1, default=str)
    n_ok = sum(1 for r in RES if r['resultado'] == 'PASS')
    print(f'\n=== {produto} Westgard na sequência de ativos (ADR-083): {n_ok} PASS / {len(RES) - n_ok} FAIL ===',
          flush=True)
    sys.exit(0 if RES and n_ok == len(RES) else 1)


if __name__ == '__main__':
    main(sys.argv[1], os.path.abspath(sys.argv[2]), sys.argv[3])
