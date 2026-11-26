# -*- coding: utf-8 -*-
"""QA do ADR-103 (entrega 49) -- alvos a VARIOS analitos (mesmo lote e periodo). Roda numa COPIA instalada
(instalar_adr102 + instalar_adr103); nada e salvo.

Uso: python qa_adr103.py <Bioquimica|Hematologia> <copia_instalada.xlsm> [saida.json]

Pelo caminho real: B98/E98/G98 escritos na secao; o botao e chamado pelo seu nucleo (mAlvosRef.AplicarAlvosVarios) com
as respostas dadas por mAlvosRef.DefinirRespostasTeste (Interactive = False: nenhuma caixa abre). Oraculo
INDEPENDENTE por analito: N/Media/DP do periodo pelo calcular_alvos do aplicar_alvos_ciq.py sobre a tblCQ_Final lida
aqui (PARTICIPA = SIM, lote e analito exatos, todos os equipamentos, DATA entre as datas, inclusive).

  V00 pre-condicoes: lote com >= 3 analitos de chave unica na LotesStore e N >= recomendado em todos os niveis
  V01 janela: "Selecionar todos" marca o catalogo inteiro; desmarcar um; marcar um; "Desmarcar todos" desabilita Aplicar
  V02 cancelar o modo e cancelar a confirmacao: CANCELADO e a LotesStore inteira igual
  V03 3 analitos, periodo inteiro, Media e DP: OK|3; cada analito = SEU oraculo (Media/DP por nivel); os nao
      selecionados e as outras linhas da LotesStore iguais; os 3 analitos com valores diferentes entre si; trilha
      PARAMETRO_LOTE_CALCULADO de cada um e UMA PARAMETRO_LOTE_CALCULADO_VARIOS; confirmacao cita periodo e quantidade
  V04 analito sem registro no lote e analito inexistente na lista: ficam de fora (bloqueados), os outros sao gravados
  V05 janela curta (N baixo): a ciencia coletiva NAO deixa de fora os temporarios (nada muda neles) e grava os
      demais = oraculo; SIM grava os temporarios com origem MANUAL ... TEMPORARIO e CONDICAO TEMPORARIO
  V06 o botao de UM analito continua igual (mesmo resultado do AplicarAlvos sem argumentos = oraculo do analito do
      Painel)
  S01 tabela de teste do usuario (6,00 / 5,99 / 5,00 / 4,99 / 4,00 Marginal / 3,99 / 3,00 / 2,99)
  S02 classificacao e status literais nas fronteiras +-1e-6 (oraculo proprio); vazio e texto = ""
  S03 Estatistica M = regra institucional em todas as linhas com Sigma; Painel mostra a classe
  P01 C4:D4 sem a informacao; P02 H4:I4 mescladas e centralizadas; P03 status por nivel na coluna tecnica oculta AB
      (stNivel), Q4 lendo dela, Classe Sigma em J6:J (J:K mesclados) com 5 cores, larguras de J/K inalteradas; P04 sem #REF! no topo
  B01 tabela de digitacao (N114) com a estrutura da grade; vazia = recusa; B02 cancelar nao transfere; B03 so
      valores (formula vira valor), linhas/colunas certas, RefBulaStore, digitacao intacta, trilha; B04 sobrescrever
      avisa; B05 invalidos recusam tudo; B06 LotesStore intacta
  Y00 arquivo de entrada intacto
"""
import collections
import datetime as dt
import json
import os
import sys
import traceback

AQUI = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, AQUI)
sys.path.insert(0, os.path.join(AQUI, '..'))
sys.path.insert(0, os.path.join(AQUI, '..', '..', 'agentes', 'ferramentas'))
from qa_final import QA, reg, RES, sha256_arquivo  # noqa: E402
from instalar_integracao import SENHA  # noqa: E402
import aplicar_alvos_ciq as oraculo  # noqa: E402
import painel_layout  # noqa: E402

R_ENT = 98


def txt(v):
    return oraculo.txt(v)


def U(v):
    return txt(v).upper()


def num(v):
    return oraculo.num(v)


def perto(a, b, tol=1e-9):
    return num(a) and num(b) and abs(float(a) - float(b)) <= tol * max(1.0, abs(float(b)))


def dia(v):
    return v.date() if isinstance(v, dt.datetime) else None


def executar(produto, caminho, saida):
    h0 = sha256_arquivo(caminho)
    q = QA(produto, caminho)
    try:
        xl, nlv = q.ex.xl, q.nlv
        p = q.ws('Painel')
        ls = q.ws('LotesStore')

        def run(m, *a, teto=900):
            return q.run('mAlvosRef.' + m if '.' not in m else m, *a, teto=teto)

        def resp(s):
            run('DefinirRespostasTeste', s)

        def aud0():
            return int(q.run('mAuditoria.UltimaLinhaAudit', teto=60))

        def trilha(r0, acao):
            al = q.ws('Audit_Log')
            r1 = aud0()
            if r1 <= r0:
                return []
            return [l for l in al.Range(al.Cells(r0 + 1, 1), al.Cells(r1, 25)).Value
                    if txt(l[6]) == acao and txt(l[7]) == 'mAlvosRef']

        def secao(lote, ini, fim):
            xl.EnableEvents = False
            try:
                p.Range(f'B{R_ENT}').Value = lote
                p.Range(f'E{R_ENT}').Value = ini if ini else None
                p.Range(f'G{R_ENT}').Value = fim if fim else None
            finally:
                xl.EnableEvents = True

        def ult_ls():
            return ls.Cells(ls.Rows.Count, 16).End(-4162).Row

        def foto_ls():
            u = ult_ls()
            return [tuple(r) for r in ls.Range(f'A2:AO{u}').Value]

        def linha_de(chave):
            u = ult_ls()
            return [i for i, v in enumerate(ls.Range(f'P2:P{u}').Value, 2) if U(v[0]) == chave][0]

        # ------------------------------------------------------------ V00 pre-condicoes
        if not run('SecaoInstalada', teto=60):
            raise RuntimeError('secao do ADR-102 nao instalada no Painel')
        lim = txt(run('LimitesNTexto', teto=60)).split('|')
        rec = int(lim[0])
        rows = q.ler('tblCQ_Final')
        sim = [r for r in rows if U(r.get('PARTICIPA_ESTATISTICA')) == 'SIM' and num(r.get('RESULTADO'))]
        catalogo = [txt(v[0]) for v in q.espec().Range('A4:A43').Value if txt(v[0])]
        cat_u = {U(x): x for x in catalogo}
        u = ult_ls()
        chaves = collections.Counter(U(v[0]) for v in ls.Range(f'P2:P{u}').Value if txt(v[0]))
        cont = collections.Counter((U(r['LOTE']), U(r['ANALITO']), int(float(r['NIVEL']))) for r in sim
                                   if txt(r.get('NIVEL')))
        por_lote = collections.defaultdict(list)
        for lt in {k[0] for k in cont}:
            for an in cat_u:
                if chaves.get(f'{lt}|{an}') == 1 and all(cont.get((lt, an, nv), 0) >= rec for nv in range(1, nlv + 1)):
                    por_lote[lt].append(an)
        cand = sorted(((len(v), lt) for lt, v in por_lote.items()), reverse=True)
        if not cand or cand[0][0] < 3:
            reg('V00 Pre-condicoes: lote com >= 3 analitos de chave unica e N >= recomendado', False,
                {'melhor': cand[:3]})
            return
        LT = cand[0][1]
        lote = next(txt(r['LOTE']) for r in sim if U(r['LOTE']) == LT)
        ANS = sorted(por_lote[LT], key=lambda a: -sum(cont.get((LT, a, nv), 0) for nv in range(1, nlv + 1)))[:3]
        nomes = [cat_u[a] for a in ANS]
        rs = {a: [r for r in sim if U(r['LOTE']) == LT and U(r['ANALITO']) == a] for a in cat_u}

        def oraculo_an(a, ini, fim):
            sel = [r for r in rs[a] if (ini is None or (dia(r.get('DATA_HORA')) and dia(r['DATA_HORA']) >= ini.date()))
                   and (fim is None or (dia(r.get('DATA_HORA')) and dia(r['DATA_HORA']) <= fim.date()))]
            c = oraculo.calcular_alvos(sel, nlv, n=10 ** 9, minimo=2)
            out = {}
            for nv in range(1, nlv + 1):
                x = c.get((LT, a, nv))
                out[nv] = (x['n_lote'], x['media_periodo'], x['dp_periodo']) if x else (0, None, None)
            return out

        def classe(o):
            """Como a VBA: nivel gravavel = N >= 2 e DP > 0; TEMPORARIO = gravavel com N < recomendado."""
            grav = [nv for nv, (n, m, s) in o.items() if n >= 2 and num(s) and s > 0]
            if not grav:
                return 'BLOQUEADO'
            return 'TEMP' if any(o[nv][0] < rec for nv in grav) else 'OK'

        datas = sorted(d for a in ANS for d in (dia(r.get('DATA_HORA')) for r in rs[a]) if d)
        d_ini = dt.datetime.combine(datas[0], dt.time())
        d_fim = dt.datetime.combine(datas[-1], dt.time())
        reg('V00 Pre-condicoes: lote com >= 3 analitos de chave unica e N >= recomendado em todos os niveis', True,
            {'lote': lote, 'analitos': nomes, 'periodo': [str(d_ini.date()), str(d_fim.date())], 'recomendado': rec})

        # ------------------------------------------------------------ V01 janela
        def jan(acoes):
            return str(run('TesteJanelaVarios', acoes, teto=60) or '').split(chr(9))   # sem strip: a selecao vazia comeca com TAB
        j1, j2, j3, j4 = jan('TODOS'), jan(f'TODOS|-{catalogo[0]}'), jan(f'NENHUM|+{nomes[1]}'), jan('TODOS|NENHUM')
        reg('V01 Janela: Selecionar todos = catalogo inteiro; desmarcar um; marcar um; Desmarcar todos desabilita Aplicar',
            j1[0].split('|') == catalogo and j1[2] == '1' and j2[0].split('|') == catalogo[1:]
            and j3[0] == nomes[1] and j3[2] == '1' and j4[0] == '' and j4[2] == '0'
            and j1[1].startswith(f'{len(catalogo)} de {len(catalogo)}'),
            {'todos': j1[1], 'menos_um': j2[1], 'um': j3[1:], 'nenhum': j4[1:]})

        # ------------------------------------------------------------ V02 cancelar
        secao(lote, d_ini, d_fim)
        antes = foto_ls()
        r0 = aud0()
        resp('')
        c1 = txt(run('AplicarAlvosVarios', '|'.join(nomes)))
        resp('1|NAO')
        c2 = txt(run('AplicarAlvosVarios', '|'.join(nomes)))
        resp('')
        reg('V02 Cancelar o modo e a confirmacao: CANCELADO, LotesStore inteira igual, sem trilha',
            c1 == 'CANCELADO|transferencia' and c2 == 'CANCELADO|confirmacao' and foto_ls() == antes
            and not trilha(r0, 'PARAMETRO_LOTE_CALCULADO'), {'modo': c1, 'confirmacao': c2})

        # ------------------------------------------------------------ V03 tres analitos, periodo inteiro
        oras = {a: oraculo_an(a, d_ini, d_fim) for a in ANS}
        linhas = {a: linha_de(f'{LT}|{a}') for a in ANS}
        antes = foto_ls()
        r0 = aud0()
        resp('1|SIM|SIM')
        r = txt(run('AplicarAlvosVarios', '|'.join(nomes)))
        pergunta = txt(run('UltimaPergunta', teto=60))
        resp('')
        depois = foto_ls()
        dif = []
        for a in ANS:
            ln = depois[linhas[a] - 2]
            for nv in range(1, nlv + 1):
                n, m, s = oras[a][nv]
                if n >= 2 and num(s) and s > 0:
                    if not (perto(ln[2 + (nv - 1) * 2], m) and perto(ln[3 + (nv - 1) * 2], s)):
                        dif.append({'analito': a, 'nivel': nv, 'LS': ln[2 + (nv - 1) * 2:4 + (nv - 1) * 2],
                                    'oraculo': (m, s)})
        outras = [i + 2 for i, (x, y) in enumerate(zip(antes, depois)) if x != y and i + 2 not in linhas.values()]
        medias = [depois[linhas[a] - 2][2] for a in ANS]
        t1 = trilha(r0, 'PARAMETRO_LOTE_CALCULADO')
        t2 = trilha(r0, 'PARAMETRO_LOTE_CALCULADO_VARIOS')
        an_trilha = {U(l[14]) for l in t1}
        reg('V03 Tres analitos, Media e DP: OK|3; cada um = SEU oraculo; outras linhas da LotesStore iguais; valores '
            'distintos entre analitos; trilha de cada analito + 1 resumo; confirmacao com periodo e quantidade',
            r.startswith('OK|3|0|0|0|') and not dif and not outras and len(set(medias)) == 3
            and an_trilha == set(ANS) and len(t2) == 1
            and 'Analitos selecionados: 3' in pergunta and d_ini.strftime('%d/%m/%Y') in pergunta,
            {'retorno': r[:300], 'diferencas': dif[:6], 'outras_linhas_alteradas': outras[:10], 'medias_N1': medias,
             'trilha_analitos': sorted(an_trilha), 'resumo': len(t2), 'pergunta': pergunta[:400]})

        # ------------------------------------------------------------ V04 sem registro / inexistente
        sem_reg = next((cat_u[a] for a in cat_u if chaves.get(f'{LT}|{a}', 0) == 0), None)
        lista = [nomes[0]] + ([sem_reg] if sem_reg else []) + ['QA_ANALITO_INEXISTENTE']
        antes = foto_ls()
        resp('1|SIM|SIM')
        r = txt(run('AplicarAlvosVarios', '|'.join(lista)))
        resp('')
        depois = foto_ls()
        nb = len(lista) - 1
        outras = [i + 2 for i, (x, y) in enumerate(zip(antes, depois)) if x != y and i + 2 != linhas[ANS[0]]]
        reg('V04 Analito sem registro no lote e inexistente: ficam de fora (bloqueados), o outro e gravado',
            r.startswith(f'OK|1|{nb}|0|0|') and 'SEM_REGISTRO' in r and not outras,
            {'lista': lista, 'retorno': r[:300], 'outras_linhas_alteradas': outras[:10]})

        # ------------------------------------------------------------ V05 janela curta: temporarios
        fim_curto = None
        for d in sorted(set(datas)):
            fimd = dt.datetime.combine(d, dt.time())
            cls = [classe(oraculo_an(a, d_ini, fimd)) for a in ANS]
            if 'TEMP' in cls:
                fim_curto = fimd
                break
        if fim_curto is None:
            reg('V05 Janela curta com analito TEMPORARIO', False, {'motivo': 'nenhuma janela com N baixo'})
        else:
            oc = {a: oraculo_an(a, d_ini, fim_curto) for a in ANS}
            cl = {a: classe(oc[a]) for a in ANS}
            secao(lote, d_ini, fim_curto)
            antes = foto_ls()
            resp('1|SIM|NAO')
            r_nao = txt(run('AplicarAlvosVarios', '|'.join(nomes)))
            meio = foto_ls()
            n_ok = sum(1 for a in ANS if cl[a] == 'OK')
            n_t = sum(1 for a in ANS if cl[a] == 'TEMP')
            n_b = sum(1 for a in ANS if cl[a] == 'BLOQUEADO')
            temp_intactos = all(meio[linhas[a] - 2] == antes[linhas[a] - 2] for a in ANS if cl[a] != 'OK')
            resp('1|SIM|SIM')
            r_sim = txt(run('AplicarAlvosVarios', '|'.join(nomes)))
            resp('')
            fim_ = foto_ls()
            dif = []
            for a in ANS:
                if cl[a] == 'BLOQUEADO':
                    continue
                ln = fim_[linhas[a] - 2]
                for nv in range(1, nlv + 1):
                    n, m, s = oc[a][nv]
                    if n >= 2 and num(s) and s > 0 and not (perto(ln[2 + (nv - 1) * 2], m) and perto(ln[3 + (nv - 1) * 2], s)):
                        dif.append({'analito': a, 'nivel': nv})
                if cl[a] == 'TEMP':
                    w = U(ln[22])
                    conds = [U(ln[26 + (nv - 1) * 5 + 3]) for nv in range(1, nlv + 1)]
                    if 'TEMPORARIO' not in w or not any('TEMPORARIO' in c for c in conds):
                        dif.append({'analito': a, 'origem': w[:120], 'condicoes': conds})
            reg('V05 Janela curta: ciencia NAO deixa os temporarios de fora (intactos) e grava os demais; SIM grava os '
                'temporarios (= oraculo; origem e CONDICAO TEMPORARIO)',
                r_nao.startswith(f'OK|{n_ok}|{n_b}|{n_t}|0|') and temp_intactos
                and r_sim.startswith(f'OK|{n_ok + n_t}|{n_b}|0|0|') and not dif and n_t > 0,
                {'ate': str(fim_curto.date()), 'classes': cl, 'nao': r_nao[:200], 'sim': r_sim[:200], 'dif': dif[:6]})

        # ------------------------------------------------------------ V06 botao de um analito igual
        a1 = ANS[2]
        pos = [U(x) for x in [txt(v[0]) for v in q.espec().Range('A4:A43').Value]].index(a1) + 1
        p.Unprotect(SENHA)
        xl.EnableEvents = False
        p.Range('B3').Value = pos
        xl.EnableEvents = True
        q.run('mEstatistica.PainelMudou', teto=900)
        secao(lote, d_ini, d_fim)
        ora = oraculo_an(a1, d_ini, d_fim)
        resp('1|SIM|SIM')
        r = txt(run('AplicarAlvos'))
        resp('')
        ln = foto_ls()[linhas[a1] - 2]
        ok6 = r.startswith('OK|') and all(perto(ln[2 + (nv - 1) * 2], ora[nv][1]) and perto(ln[3 + (nv - 1) * 2], ora[nv][2])
                                          for nv in range(1, nlv + 1) if ora[nv][0] >= 2 and num(ora[nv][2]) and ora[nv][2] > 0)
        reg('V06 Botao de UM analito (AplicarAlvos sem argumentos) continua igual: OK e = oraculo do analito do Painel',
            ok6 and U(q.nome('selAnalito').Value) == a1, {'analito': a1, 'retorno': r[:200]})

        # ------------------------------------------------------------ S01..S03 regra institucional do Sigma
        casos = [(6.00, 'Excelente — Classe mundial'), (5.99, 'Excelente'), (5.00, 'Bom'), (4.99, 'Bom'),
                 (4.00, 'Marginal'), (3.99, 'Marginal'), (3.00, 'Marginal'), (2.99, 'Insatisfatório')]
        got = [(s_, txt(q.run('mQualidade.ClassificarSigma', s_, teto=30))) for s_, _ in casos]
        reg('S01 Tabela de teste do usuario (sigma = 4,00 -> Marginal, decisao de 26/11/2026)',
            [g for _, g in got] == [c for _, c in casos], {'obtido': got})
        ESC = [('Excelente — Classe mundial', 'Excelente'), ('Excelente', 'Alto desempenho'),
               ('Bom', 'Desempenho adequado'), ('Marginal', 'Atenção'),
               ('Insatisfatório', 'método mais instável ou tendencioso')]

        def faixa(s_):
            return ESC[0] if s_ >= 6 else ESC[1] if 5 < s_ < 6 else ESC[2] if 4 < s_ <= 5 else \
                ESC[3] if 3 <= s_ <= 4 else ESC[4]
        pts = [x + d for x in (3.0, 4.0, 5.0, 6.0) for d in (-1e-6, 0.0, 1e-6)] + [0.0, -2.0, 2.5, 7.0, 100.0]
        dif = [(s_, txt(q.run('mQualidade.ClassificarSigma', s_, teto=30)), txt(q.run('mQualidade.StatusSigma', s_,
                                                                                         teto=30)), faixa(s_))
               for s_ in pts]
        dif = [d for d in dif if (d[1], d[2]) != d[3]]
        vaz = [txt(q.run('mQualidade.ClassificarSigma', v_, teto=30)) for v_ in ('', 'abc')]
        reg('S02 Classificacao e status literais nas fronteiras (3, 4, 5, 6 e +-1e-6) e fora delas; vazio/texto = ""',
            not dif and vaz == ['', ''], {'diferencas': dif, 'vazio_texto': vaz})
        e = q.ws('Estatística')
        xl.Calculate()
        lin = [list(x) for x in e.Range('A14:M133').Value]
        dif = [(14 + i, x[11], x[12]) for i, x in enumerate(lin) if num(x[11]) and txt(x[12]) != faixa(float(x[11]))[0]]
        n_s = sum(1 for x in lin if num(x[11]))
        cj = painel_layout.col_classe(p)
        pj = [(txt(p.Range(f'I{6 + i}').Value), txt(p.Range(f'{cj}{6 + i}').Value)) for i in range(1, nlv + 1)]
        pdif = [x for x in pj if x[1] and x[1] not in [c for c, _ in ESC]]
        reg('S03 Estatistica M = regra institucional em TODAS as linhas com Sigma; Painel mostra a classe da coluna M',
            n_s > 0 and not dif and not pdif, {'linhas_com_sigma': n_s, 'diferencas': dif[:10], 'painel': pj})

        # ------------------------------------------------------------ P01..P04 topo do Painel
        c4 = (txt(p.Range('C4').Formula), p.Range('C4').MergeCells, txt(p.Range('D4').Formula))
        reg('P01 C4:D4 sem a informação (vazias, sem mesclagem)', c4 == ('', False, ''), {'C4_D4': c4})
        h4 = p.Range('H4')
        reg('P02 H4:I4 mescladas e centralizadas (horizontal e vertical); a formula do aviso continua em H4',
            h4.MergeCells and h4.MergeArea.Address == '$H$4:$I$4' and h4.MergeArea.HorizontalAlignment == -4108
            and h4.MergeArea.VerticalAlignment == -4108 and txt(h4.Formula).startswith('=IF('),
            {'merge': h4.MergeArea.Address, 'H': h4.MergeArea.HorizontalAlignment, 'V': h4.MergeArea.VerticalAlignment})
        cs = painel_layout.col_status(p)
        q4 = txt(p.Range('Q4').Formula)
        st = [txt(painel_layout.status_nivel(p, i)) for i in range(1, nlv + 1)]
        reg('P03 Status por nivel fora da interface (coluna tecnica oculta, nome stNivel); Q4 continua lendo o status; '
            'J6:J = Classe Sigma (formula da Estatistica M) com as 5 cores, em J:K mesclados (larguras '
            'inalteradas: o corte do K4 e conferido pelo qa_adr090 V02)',
            painel_layout.layout_novo(p) and cs == 'AB' and p.Columns('AB').Hidden
            and f'$AB$7:$AB${6 + nlv}' in q4 and all(s_ for s_ in st)
            and all('Estatística!$M$' in txt(p.Range(f'J{6 + i}').Formula) for i in range(1, nlv + 1))
            and p.Range('J7').FormatConditions.Count == 5
            and all(txt(p.Range(f'K{r_}').Formula) == '' for r_ in range(6, 7 + nlv))
            and all(p.Range(f'J{r_}').MergeArea.Address == f'$J${r_}:$K${r_}' for r_ in range(6, 7 + nlv))
            and txt(q.wb.Names('stNivel').RefersTo).replace(' ', '') == f'=Painel!$AB$7:$AB${6 + nlv}',
            {'status': st, 'Q4': q4[:90], 'J6': txt(p.Range('J6').Value)})
        ruins = [c_.Address for c_ in p.Range('A1:AB12').Cells if '#REF' in txt(c_.Formula).upper()]
        reg('P04 Nenhum #REF! no topo do Painel; Q4 com valor', not ruins and txt(p.Range('Q4').Value) != '',
            {'ref': ruins[:6], 'Q4': txt(p.Range('Q4').Value)})

        # ------------------------------------------------------------ B01..B06 digitacao da bula e transferencia
        D, RR1, NR = 13, 115, 40
        ncol = 1 + nlv * 3 + 1
        secao(lote, d_ini, d_fim)
        run('AtualizarVistaReferencia')
        rb = q.ws('RefBulaStore')

        def foto_rb():
            u_ = rb.Cells(rb.Rows.Count, 3).End(-4162).Row
            return [tuple(x) for x in rb.Range(f'A1:O{max(u_, 1)}').Value]

        def limpar_dig():
            xl.EnableEvents = False
            p.Range(p.Cells(RR1, 2 + D), p.Cells(RR1 + NR - 1, ncol + D)).ClearContents()
            xl.EnableEvents = True
        cabg = [p.Cells(114, c_).Value for c_ in range(1, ncol + 1)]
        cabd = [p.Cells(114, c_ + D).Value for c_ in range(1, ncol + 1)]
        ang = [txt(p.Cells(r_, 1).Value) for r_ in range(RR1, RR1 + NR)]
        and_ = [txt(p.Cells(r_, 1 + D).Value) for r_ in range(RR1, RR1 + NR)]
        limpar_dig()
        ls0, rb0 = foto_ls(), foto_rb()
        resp('SIM')
        r_vazia = txt(run('TransferirReferenciaCore'))
        reg('B01 Tabela de digitacao a partir de N114 com a mesma estrutura (cabecalho e analitos da grade); vazia: '
            'recusa, nada transferido',
            cabd == cabg and and_ == ang and r_vazia.startswith('ERRO|') and 'vazia' in r_vazia and foto_rb() == rb0,
            {'cab_grade': cabg, 'cab_dig': cabd, 'retorno': r_vazia[:120]})
        # dois analitos com chave do lote: N1 Media/LI/LS e a Fonte; uma celula com FORMULA (deve ir o VALOR)
        alvos_b = [r_ for r_ in range(RR1, RR1 + NR) if U(p.Cells(r_, 1).Value) in ANS][:2]
        grade0 = {r_: [p.Cells(r_, c_).Value for c_ in range(2, ncol + 1)] for r_ in alvos_b}
        xl.EnableEvents = False
        vals = {}
        for k_, r_ in enumerate(alvos_b):
            base = 100.0 + 10 * k_
            vals[r_] = (base, base - 5, base + 5)
            p.Cells(r_, 2 + D).Value = base
            p.Cells(r_, 3 + D).Formula = f'={base - 5}+0'          # formula na digitacao: vai so o valor
            p.Cells(r_, 4 + D).Value = base + 5
            p.Cells(r_, ncol + D).Value = 'Bula QA'
        xl.EnableEvents = True
        dig0 = [tuple(x) for x in p.Range(p.Cells(RR1, 1 + D), p.Cells(RR1 + NR - 1, ncol + D)).Formula]
        # a grade ja tinha valor diferente nessas celulas? entao a confirmacao tem de avisar a substituicao
        havia = any(v0 not in (None, '') for r_ in alvos_b for v0 in grade0[r_])
        resp('NAO')
        r_nao = txt(run('TransferirReferenciaCore'))
        grade_nao = {r_: [p.Cells(r_, c_).Value for c_ in range(2, ncol + 1)] for r_ in alvos_b}
        reg('B02 Cancelar a confirmacao: nada transferido (grade e RefBulaStore iguais)',
            r_nao == 'CANCELADO|confirmacao' and grade_nao == grade0 and foto_rb() == rb0, {'retorno': r_nao})
        r0 = aud0()
        resp('SIM')
        r_sim = txt(run('TransferirReferenciaCore'))
        pg = txt(run('UltimaPergunta', teto=60))
        resp('')
        ok_g, sem_formula = True, True
        for r_ in alvos_b:
            for j_, v_ in enumerate(vals[r_]):
                c_ = p.Cells(r_, 2 + j_)
                if not perto(c_.Value, v_):
                    ok_g = False
                if c_.HasFormula:
                    sem_formula = False
            if txt(p.Cells(r_, ncol).Value) != 'Bula QA':
                ok_g = False
        st_rb = {U(x[2]): x for x in foto_rb()[1:] if txt(x[2])}
        ok_rb = all(perto(st_rb.get(f'{LT}|{U(p.Cells(r_, 1).Value)}', [None] * 15)[3 + j_], v_)
                    for r_ in alvos_b for j_, v_ in enumerate(vals[r_]))
        dig1 = [tuple(x) for x in p.Range(p.Cells(RR1, 1 + D), p.Cells(RR1 + NR - 1, ncol + D)).Formula]
        aud = trilha(r0, 'REFERENCIA_BULA_ALTERADA')
        reg('B03 Transferir: SO VALORES nas linhas/colunas certas (formula da digitacao chega como valor, sem formula '
            'na grade); RefBulaStore do lote gravada; digitacao intacta; trilha REFERENCIA_BULA_ALTERADA; LotesStore igual',
            r_sim == f'OK|{3 * len(alvos_b) + len(alvos_b)}|{len(alvos_b)}' and ok_g and sem_formula and ok_rb
            and dig1 == dig0 and len(aud) >= 3 * len(alvos_b) and foto_ls() == ls0
            and (('SUBSTITU' in pg) == havia),
            {'retorno': r_sim, 'grade_ok': ok_g, 'sem_formula': sem_formula, 'store_ok': ok_rb, 'trilha': len(aud),
             'digitacao_intacta': dig1 == dig0})
        # sobrescrever: confirmacao avisa
        xl.EnableEvents = False
        p.Cells(alvos_b[0], 2 + D).Value = vals[alvos_b[0]][0] + 1
        xl.EnableEvents = True
        resp('SIM')
        r_sob = txt(run('TransferirReferenciaCore'))
        pg = txt(run('UltimaPergunta', teto=60))
        resp('')
        reg('B04 Valor ja transferido: a confirmacao avisa que sera SUBSTITUIDO; confirmado, o novo valor vai para a grade',
            r_sob.startswith('OK|') and 'SUBSTITU' in pg and perto(p.Cells(alvos_b[0], 2).Value, vals[alvos_b[0]][0] + 1),
            {'retorno': r_sob, 'pergunta': pg[-160:]})
        # invalidos: texto em coluna numerica; LI > LS
        grade_b = [p.Cells(alvos_b[1], c_).Value for c_ in range(2, ncol + 1)]
        rb_b = foto_rb()
        xl.EnableEvents = False
        p.Cells(alvos_b[1], 2 + D).Value = 'abc'
        xl.EnableEvents = True
        resp('SIM')
        r_txt = txt(run('TransferirReferenciaCore'))
        xl.EnableEvents = False
        p.Cells(alvos_b[1], 2 + D).Value = vals[alvos_b[1]][0]
        p.Cells(alvos_b[1], 3 + D).Value = 999.0                     # LI > LS
        xl.EnableEvents = True
        resp('SIM')
        r_lils = txt(run('TransferirReferenciaCore'))
        resp('')
        reg('B05 Valores invalidos (texto em coluna numerica; LI > LS): recusa a transferencia INTEIRA, nada muda',
            r_txt.startswith('ERRO|') and 'número' in r_txt and r_lils.startswith('ERRO|') and 'LI maior' in r_lils
            and [p.Cells(alvos_b[1], c_).Value for c_ in range(2, ncol + 1)] == grade_b and foto_rb() == rb_b,
            {'texto': r_txt[:160], 'li_ls': r_lils[:160]})
        reg('B06 A transferencia NUNCA grava na LotesStore (A2:AO inteira igual)', foto_ls() == ls0, {})
        limpar_dig()
    except Exception as e:                                # noqa: BLE001
        reg('Z99 Execucao sem excecao', False, {'erro': repr(e), 'trace': traceback.format_exc()[-1500:]})
    finally:
        try:
            q.run('mAlvosRef.DefinirRespostasTeste', '', teto=30)
        except Exception:                                 # noqa: BLE001
            pass
        q.fechar(salvar=False)
    h1 = sha256_arquivo(caminho)
    reg('Y00 Arquivo de entrada intacto (SHA-256)', h0 == h1, {'antes': h0[:16], 'depois': h1[:16]})
    if saida:
        with open(saida, 'w', encoding='utf-8') as f:
            json.dump(RES, f, ensure_ascii=False, indent=1, default=str)
    falhas = [r for r in RES if r['resultado'] == 'FAIL']
    print(f'\n=== {produto} ADR-103 (alvos a varios analitos): {len(RES) - len(falhas)} PASS / {len(falhas)} FAIL ===',
          flush=True)
    return falhas


if __name__ == '__main__':
    f = executar(sys.argv[1], os.path.abspath(sys.argv[2]), sys.argv[3] if len(sys.argv) > 3 else None)
    sys.exit(1 if f else 0)
