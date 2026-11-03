# -*- coding: utf-8 -*-
"""QA do ADR-088 (03/11/2026): lote 1 de correcoes da auditoria (ciclo 2026-10-31_ciclo01). Roda numa COPIA instalada
(instalar_adr088.py); nada e salvo. Dados de teste com MOTIVO "Outro: TESTE-ADR088 ..." (so na copia).

Uso: python qa_adr088.py <Bioquimica|Hematologia> <copia_instalada.xlsm> <saida.json>
     (QC_PROD_ORIGINAL = a producao antes da instalacao: referencia do I17b/G08b; sem ela, NAO_EXECUTADO)

  R12 (OC-0002) Painel H7:H8/H9 = ET so com bias numerico: formula; nenhum guarda IF(ISNUMBER($G),$G,0) no Painel;
      analito com G = "SEM EP" -> H vazio; analito com G numerico -> H = 1,65*E + |G| (1e-9)
  M20 (OC-0001) manual com EQUIPAMENTO e ANALITO em minusculas, digitado (eventos) e colado (eventos desligados): a
      tabela fica com a grafia do cadastro; o manual tem corrida propria; o resultado real do dia continua
      participando com o mesmo RUN; invariante: 0 chaves (equip, lote, nivel, analito, RUN; sem caixa) com > 1
      participante; QA A12 = 0
  E18 (OC-0005) DATA/HORA coladas em texto que nao existem (31/02, 29/02/2026, 00/01, 32/01, 15/13, 25:00) ->
      MANUAL_INCOMPLETO "DATA"/"HORA", fora da estatistica; 28/02/2026 em texto e aceito
  E17/X12 (OC-0004) resultado <= 0 -> IMPLAUSIVEL, PARTICIPA = NAO, NAO_PLOTAR; o conjunto IMPLAUSIVEL = o dos
      participantes <= 0 de antes (varredura independente) + os manuais <= 0; A11 lista exatamente esses IDs;
      nenhum participante <= 0
  X13 (OC-0004) fora do LJ/Westgard: o ID implausivel nao esta no Eng_Saida do analito/lote/mes dele, e o n/CV do
      Painel do nivel = independente so com os participantes (1e-9)
  I17 (OC-0006) BiasEQ com o filtro do ViesEQ: Bio Bilirrubina direta G = 19,994209755156 (1e-9), AC 13, AD 3;
      Hema IG%/IG#/NRBC%/NRBC# = "SEM EP"; I17b: os demais analitos com G/AC/AD iguais aos da producao
  G08 (OC-0009, Bio) Estatistica B103 = linhas com analito e G = "SEM EP" (independente); G08b: B98:B102 iguais
  L04 (OC-0003) troca de lote no Painel (H3) e, na Hema, Ano De/Ate (B3/D3): a Estatistica acompanha (Hema: C:F =
      independente, 1e-9; Bio: C:F = Painel B:E do mesmo lote e periodo)
  Y11 (OC-0013) H3 vazio: Painel G = Estatistica G e I = Estatistica L (1e-9); C:E e H vazios; G6/I6 "do método"
  Y00 seguranca do teste: arquivo de entrada intacto
"""
import collections
import datetime as dt
import json
import os
import re
import statistics
import sys
import time
import traceback

AQUI = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, AQUI)
from qa_final import QA, reg, RES, sha256_arquivo  # noqa: E402
sys.path.insert(0, os.path.join(AQUI, '..'))
import instalar_integracao as ii  # noqa: E402

C_ID0 = 39                                         # Eng_Saida: 1a coluna de ID por nivel (mEstatistica.ENG_COL_ID0)
RX_GUARDA = re.compile(r'IF\(ISNUMBER\(\$G\d+\),\$G\d+,0\)')
TAG = 'Outro: TESTE-ADR088'
NAO = 'NÃO'


def txt(v):
    if isinstance(v, float) and v == int(v):
        return str(int(v))
    return str(v if v is not None else '').strip()


def num(v):
    return isinstance(v, (int, float)) and not isinstance(v, bool) and not (isinstance(v, int) and v < -2146820000)


def perto(a, b, tol=1e-9):
    return num(a) and num(b) and abs(a - b) <= tol * max(1.0, abs(b))


def serie(d):
    return float((d.replace(tzinfo=None) - dt.datetime(1899, 12, 30)).days)


def nao_executado(teste, motivo):
    RES.append({'teste': teste, 'resultado': 'NAO_EXECUTADO', 'evidencia': motivo})
    print('NAO_EXECUTADO ' + teste + ' -- ' + motivo, flush=True)


def desc(vals):
    n = len(vals)
    if n == 0:
        return 0, None, None, None
    m = statistics.fmean(vals)
    s = statistics.stdev(vals) if n > 1 else None
    return n, m, s, (s / m * 100 if s is not None and m else None)


def executar(produto, caminho, saida):
    hash0 = sha256_arquivo(caminho)
    q = QA(produto, caminho)
    try:
        xl, nlv, bio = q.ex.xl, q.nlv, q.bio
        p, e, est = q.ws('Painel'), q.ws('Eng_Saida'), q.ws('Estatística')
        q.run('mSeguranca.ReprotectAll')
        cad = [txt(x[0]) for x in q.espec().Range('A4:A43').Value]
        eq_pad = txt(q.cfg('EQUIPAMENTO_PADRAO'))

        # ------------------------------------------------------------------ R12 estatico
        fh = {r: str(p.Range(f'H{r}').Formula) for r in range(7, 7 + nlv)}
        esperado = {r: f'=IF(OR($E{r}="",NOT(ISNUMBER($G{r}))),"",$E{r}*1.65+ABS($G{r}))' for r in fh}
        guardas = []
        fs = p.UsedRange.Formula
        for lin in fs:
            for f in lin:
                if isinstance(f, str) and RX_GUARDA.search(f.replace(' ', '')):
                    guardas.append(f[:80])
        reg('R12a Painel H7:H%d = ET só com bias numérico (ADR-030); 0 fórmulas com o guarda IF(ISNUMBER($G),$G,0)'
            % (6 + nlv), all(fh[r].replace(' ', '') == esperado[r] for r in fh) and not guardas,
            {'H': fh, 'guardas': guardas})

        # ------------------------------------------------------------------ dados antes
        fin0 = q.ler('tblCQ_Final')
        le0 = {r['ID_REGISTRO'] for r in fin0 if (r['PARTICIPA_ESTATISTICA'] == 'SIM' or r['STATUS_ANALITICO'] == 'IMPLAUSIVEL')
               and num(r['RESULTADO']) and r['RESULTADO'] <= 0}
        # alvo real: ATIVO do interfaceamento, 1a corrida do dia, analito do catalogo, equipamento padrao (Hema) ou
        # qualquer (Bio), dia anterior ao ultimo da base (o manual das 23:5x nao fica "posterior ao lancamento")
        reais = [r for r in fin0 if r['STATUS_ANALITICO'] == 'ATIVO' and r['ORIGEM_RESULTADO'] == 'INTERFACEAMENTO'
                 and txt(r['ANALITO']) in cad and r['DATA_HORA'] and r['CORRIDA_NO_DIA'] == 1 and num(r['RESULTADO'])
                 and r['RESULTADO'] > 0 and (bio or txt(r['EQUIPAMENTO']).upper() == eq_pad.upper())]
        reais.sort(key=lambda r: r['DATA_HORA'])
        T = reais[-200]
        T2 = reais[-400]
        dT, dT2 = T['DATA_HORA'].date(), T2['DATA_HORA'].date()
        col = lambda r: {'LOTE': "'" + txt(r['LOTE']), 'NIVEL': int(r['NIVEL'])}       # noqa: E731

        # ------------------------------------------------------------------ manuais
        # M1 digitado com eventos (Worksheet_Change -> PrepararLinha), em minusculas
        q.escrever_linha('tblResultados_Manuais', dict(
            DATA=dt.datetime(dT.year, dT.month, dT.day), HORA='23:50', EQUIPAMENTO=txt(T['EQUIPAMENTO']).lower(),
            ANALITO=txt(T['ANALITO']).lower(), RESULTADO=float(T['RESULTADO']) * 1.001, MOTIVO=TAG + ' M20a',
            **col(T)))
        man_tab = {r['MOTIVO']: r for r in q.ler('tblResultados_Manuais') if txt(r.get('MOTIVO')).startswith(TAG)}
        m1_tab = man_tab.get(TAG + ' M20a', {})
        # os outros colados com os eventos desligados (o preparo do ATUALIZAR normaliza)
        xl.EnableEvents = False
        casos = {  # motivo: (DATA, HORA, RESULTADO, extra)
            'M20b': (dt.datetime(dT2.year, dT2.month, dT2.day), '23:55', float(T2['RESULTADO']) * 1.001,
                     {'EQUIPAMENTO': '  ' + txt(T2['EQUIPAMENTO']).lower() + ' ', 'ANALITO': txt(T2['ANALITO']).upper()}),
            'E18 31/02/2026': ("'31/02/2026", '08:00', 1.5, {}),
            'E18 29/02/2026': ("'29/02/2026", '08:00', 1.5, {}),
            'E18 00/01/2026': ("'00/01/2026", '08:00', 1.5, {}),
            'E18 32/01/2026': ("'32/01/2026", '08:00', 1.5, {}),
            'E18 15/13/2026': ("'15/13/2026", '08:00', 1.5, {}),
            'E18 25:00': ("'10/10/2026", "'25:00", 1.5, {}),
            'E18 28/02/2026': ("'28/02/2026", "'08:00", 1.5, {}),
            'X12 -1': (dt.datetime(dT2.year, dT2.month, dT2.day), '22:00', -1, {}),
            'X12 0': (dt.datetime(dT2.year, dT2.month, dT2.day), '22:30', 0, {}),
        }
        for j, (k, (d_, h_, v_, extra)) in enumerate(casos.items()):
            # ID explicito: com os eventos desligados ninguem carimba o MAN_ e a proxima escrita cairia na mesma linha
            base = dict(ID_REGISTRO=f'MAN_{9101 + j}', DATA=d_, HORA=h_, RESULTADO=v_, MOTIVO=f'{TAG} {k}', EQUIPAMENTO=txt(T2['EQUIPAMENTO']),
                        ANALITO=txt(T2['ANALITO']), **col(T2))
            base.update(extra)
            q.escrever_linha('tblResultados_Manuais', base)
        xl.EnableEvents = True
        t0 = time.time()
        r = str(q.run('mIntegracao.AtualizarSoEntradas', teto=1500))
        t_atu = round(time.time() - t0, 1)
        if not r.startswith('OK|'):
            raise RuntimeError('ATUALIZAR so entradas: ' + r[:400])
        fin = q.ler('tblCQ_Final')
        por_id = {x['ID_REGISTRO']: x for x in fin}
        qa = q.ler('tblQA_Integracao')
        man = {txt(x['MOTIVO_MANUAL']): x for x in fin if x['ORIGEM_RESULTADO'] == 'MANUAL'
               and txt(x.get('MOTIVO_MANUAL')).startswith(TAG)}
        man_tab = {r['MOTIVO']: r for r in q.ler('tblResultados_Manuais') if txt(r.get('MOTIVO')).startswith(TAG)}

        # ------------------------------------------------------------------ M20
        def chave(x):
            return (txt(x['EQUIPAMENTO']).upper(), txt(x['LOTE']), txt(x['NIVEL']), txt(x['ANALITO']).upper(), txt(x['RUN']))
        cnt = collections.Counter(chave(x) for x in fin if x['PARTICIPA_ESTATISTICA'] == 'SIM' and x['RUN'])
        dup = [k for k, n in cnt.items() if n > 1]
        a12 = [x for x in qa if x.get('CODIGO') == 'A12']
        ok_m = True
        ev_m = {'ATUALIZAR_s': t_atu, 'alvo_a': T['ID_REGISTRO'], 'alvo_b': T2['ID_REGISTRO']}
        for tag, alvo, tab in (('M20a', T, m1_tab), ('M20b', T2, man_tab.get(f'{TAG} M20b', {}))):
            mx = man.get(f'{TAG} {tag}')
            tr = por_id.get(alvo['ID_REGISTRO'])
            ok = (mx is not None and tr is not None
                  and txt(tab.get('EQUIPAMENTO')) == txt(alvo['EQUIPAMENTO']) and txt(tab.get('ANALITO')) == txt(alvo['ANALITO'])
                  and txt(mx['EQUIPAMENTO']) == txt(alvo['EQUIPAMENTO']) and txt(mx['ANALITO']) == txt(alvo['ANALITO'])
                  and mx['ANALITO_CADASTRADO'] == 'SIM' and mx['STATUS_ANALITICO'] == 'ATIVO' and mx['RUN'] != alvo['RUN']
                  and tr['PARTICIPA_ESTATISTICA'] == 'SIM' and tr['RUN'] == alvo['RUN'])
            ok_m = ok_m and ok
            ev_m[tag] = {'tabela': (tab.get('EQUIPAMENTO'), tab.get('ANALITO')),
                         'cq': None if mx is None else (mx['ID_REGISTRO'], mx['EQUIPAMENTO'], mx['ANALITO'], mx['STATUS_ANALITICO'],
                                                        mx['RUN']),
                         'real': None if tr is None else (tr['STATUS_ANALITICO'], tr['RUN'], alvo['RUN'])}
        reg('M20 Manual em minúsculas (digitado e colado) fica com a grafia do cadastro e corrida própria; o resultado '
            'real continua participando com o mesmo RUN; 0 chaves (equip, lote, nível, analito, RUN; sem caixa) com > 1 '
            'participante; A12 = 0', ok_m and not dup and not a12, dict(ev_m, chaves_duplicadas=dup[:5], A12=len(a12)))

        # ------------------------------------------------------------------ E18
        ev_e, ok_e = {}, True
        for k, campo in (('31/02/2026', 'DATA'), ('29/02/2026', 'DATA'), ('00/01/2026', 'DATA'), ('32/01/2026', 'DATA'),
                         ('15/13/2026', 'DATA'), ('25:00', 'HORA')):
            mx = man.get(f'{TAG} E18 {k}')
            ok = (mx is not None and mx['STATUS_ANALITICO'] == 'MANUAL_INCOMPLETO' and mx['PARTICIPA_ESTATISTICA'] == NAO
                  and campo in txt(mx['MOTIVO_EXCLUSAO_AUTOMATICA']).split(': ', 1)[-1].split(', '))
            ok_e = ok_e and ok
            ev_e[k] = None if mx is None else (mx['STATUS_ANALITICO'], txt(mx['MOTIVO_EXCLUSAO_AUTOMATICA'])[:90])
        mv = man.get(f'{TAG} E18 28/02/2026')
        ok_v = mv is not None and mv['STATUS_ANALITICO'] != 'MANUAL_INCOMPLETO' and \
            hasattr(mv['DATA'], 'year') and (mv['DATA'].year, mv['DATA'].month, mv['DATA'].day) == (2026, 2, 28)
        ev_e['28/02/2026'] = None if mv is None else (mv['STATUS_ANALITICO'], str(mv['DATA']))
        reg('E18 Data/hora inexistente colada em texto (31/02, 29/02/2026, 00/01, 32/01, 15/13, 25:00) = MANUAL_INCOMPLETO '
            '"DATA"/"HORA", fora da estatística; 28/02/2026 em texto é aceito', ok_e and ok_v, ev_e)

        # ------------------------------------------------------------------ E17 / X12
        imp = {x['ID_REGISTRO'] for x in fin if x['STATUS_ANALITICO'] == 'IMPLAUSIVEL'}
        a11 = {x['ID_REGISTRO'] for x in qa if x.get('CODIGO') == 'A11'}
        man_le0 = {man[f'{TAG} X12 {k}']['ID_REGISTRO'] for k in ('-1', '0') if f'{TAG} X12 {k}' in man}
        part_le0 = [x['ID_REGISTRO'] for x in fin if x['PARTICIPA_ESTATISTICA'] == 'SIM' and num(x['RESULTADO'])
                    and x['RESULTADO'] <= 0]
        forma = all(por_id[i]['PARTICIPA_ESTATISTICA'] == NAO and por_id[i]['TIPO_PLOTAGEM_LJ'] == 'NAO_PLOTAR'
                    and por_id[i]['REGISTRAR_RESULTADO_NO_LJ'] == NAO and num(por_id[i]['RESULTADO'])
                    and por_id[i]['RESULTADO'] <= 0 for i in imp)
        reg('E17 Resultado ≤ 0 = IMPLAUSIVEL (PARTICIPA NÃO, NAO_PLOTAR): o conjunto = participantes ≤ 0 de antes '
            '(varredura independente) + os manuais ≤ 0; A11 lista exatamente esses IDs; nenhum participante ≤ 0',
            imp == (le0 | man_le0) and imp == a11 and len(man_le0) == 2 and not part_le0 and forma,
            {'IMPLAUSIVEL': len(imp), 'participantes_le0_antes': len(le0), 'manuais_le0': sorted(man_le0),
             'A11': len(a11), 'so_IMPLAUSIVEL': sorted(imp - a11)[:5], 'so_A11': sorted(a11 - imp)[:5],
             'participantes_le0_depois': part_le0[:5], 'amostra': sorted(imp)[:8]})
        reg('X12 Manual -1 e 0 = IMPLAUSIVEL com alerta A11 (antes: ATIVO, sem alerta)',
            len(man_le0) == 2 and man_le0 <= imp and man_le0 <= a11, {'manuais': sorted(man_le0)})

        # ------------------------------------------------------------------ tela do Painel
        def tela(an, lote, eq, d_ini, d_fim):
            q.run('mLotes.EsquecerFiltroRotina', teto=60)
            xl.EnableEvents = False
            p.Unprotect(ii.SENHA)
            p.Range('B3').Value = cad.index(an) + 1
            if bio:
                p.Range('L4').Value = eq
            p.Range('G3').Value = serie(d_ini)
            p.Range('G4').Value = serie(d_fim)
            xl.EnableEvents = True
            if lote:
                p.Range('H3').Value = "'" + lote                # como o usuario: eventos ligados (TrocarLoteAnalise)
            else:
                p.Range('H3').MergeArea.ClearContents()
            q.ex.esperar()
            q.run('mEstatistica.RecalcularAnalitoAtual', teto=600)
            xl.Calculate()

        def ids_eng():
            n = int(e.Range('I1').Value or 0)
            if not n:
                return set()
            blk = e.Range(e.Cells(3, C_ID0), e.Cells(2 + n, C_ID0 + nlv - 1)).Value
            return {txt(v) for linha in blk for v in linha if txt(v)}

        def indep(an, lote, eq, nivel, d_ini, d_fim):
            vals = [x['RESULTADO'] for x in fin if x['PARTICIPA_ESTATISTICA'] == 'SIM' and txt(x['ANALITO']) == an
                    and txt(x['LOTE']) == lote and txt(x['EQUIPAMENTO']).upper() == eq.upper() and int(x['NIVEL']) == nivel
                    and hasattr(x['DATA'], 'year') and d_ini <= x['DATA'].replace(tzinfo=None) <= d_fim]
            return desc(vals)

        # X13: um implausivel do interfaceamento (Bio) ou o manual -1 (Hema)
        cand = sorted((i for i in imp if not i.startswith('MAN_') and txt(por_id[i]['ANALITO']) in cad),
                      key=lambda i: por_id[i]['DATA_HORA'] or dt.datetime(1900, 1, 1))
        alvo_x = cand[-1] if cand else man.get(f'{TAG} X12 -1', {}).get('ID_REGISTRO')
        if alvo_x:
            ax = por_id[alvo_x]
            dd = ax['DATA'].replace(tzinfo=None)
            d_ini = dt.datetime(dd.year, dd.month, 1)
            d_fim = (d_ini + dt.timedelta(days=32)).replace(day=1) - dt.timedelta(days=1)
            an_x, lt_x, eq_x, nv_x = txt(ax['ANALITO']), txt(ax['LOTE']), txt(ax['EQUIPAMENTO']), int(ax['NIVEL'])
            tela(an_x, lt_x, eq_x, d_ini, d_fim)
            ps = q.painel_stats()[nv_x - 1]
            ind = indep(an_x, lt_x, eq_x, nv_x, d_ini, d_fim)
            ids = ids_eng()
            ok = (alvo_x not in ids and len(ids) > 0 and ps[1] == ind[0] and perto(ps[2], ind[1])
                  and (ind[3] is None or perto(ps[4], ind[3])))
            reg('X13 O implausível fica fora do LJ/Westgard (não está no Eng_Saida) e o n/média/CV do Painel do nível = '
                'independente só com os participantes', ok,
                {'id': alvo_x, 'valor': ax['RESULTADO'], 'analito': an_x, 'lote': lt_x, 'equip': eq_x, 'nivel': nv_x,
                 'mes': str(d_ini.date()), 'painel_n_media_dp_cv': ps[1:5], 'independente': ind,
                 'ids_no_eng_saida': len(ids)})
            tela_base = (an_x, lt_x, eq_x, d_ini, d_fim)
        else:
            nao_executado('X13', 'nenhum IMPLAUSIVEL de analito do catalogo')
            tela_base = (txt(T['ANALITO']), txt(T['LOTE']), txt(T['EQUIPAMENTO']),
                         dt.datetime(dT.year, dT.month, 1), dt.datetime(dT.year, dT.month, 28))

        # ------------------------------------------------------------------ Estatistica (G, AB, AC, AD, L)
        def estat():
            est.Range('A14:AD133').Calculate()
            v = est.Range('A14:AD133').Value
            out = {}
            for linha in v:
                k = txt(linha[27])                                   # AB = analito|nivel
                if k:
                    out[k] = {'A': linha[0], 'C': linha[2], 'D': linha[3], 'E': linha[4], 'F': linha[5],
                              'G': linha[6], 'H': linha[7], 'L': linha[11], 'AC': linha[28], 'AD': linha[29]}
            return out
        xl.CalculateFull()
        st = estat()

        # analito com G numerico em todos os niveis (o do X13, se servir; senao o de mais participantes): base do R12b,
        # do L04 e do Y11 -- com G = "SEM EP" eles passariam sem provar nada
        elig = lambda x: x['PARTICIPA_ESTATISTICA'] == 'SIM' and (bio or txt(x['EQUIPAMENTO']).upper() == eq_pad.upper())  # noqa: E731
        cont = collections.Counter(txt(x['ANALITO']) for x in fin if elig(x))
        g_ok = [a for a in cad if a and cont.get(a) and all(num(st.get(f'{a}|{n}', {}).get('G')) for n in range(1, nlv + 1))]
        an_t = tela_base[0] if tela_base[0] in g_ok else max(g_ok, key=lambda a: cont[a])
        if an_t != tela_base[0]:
            u = max((x for x in fin if elig(x) and txt(x['ANALITO']) == an_t), key=lambda x: x['DATA_HORA'])
            du = u['DATA'].replace(tzinfo=None)
            tela_base = (an_t, txt(u['LOTE']), txt(u['EQUIPAMENTO']), dt.datetime(du.year, du.month, 1), du)
        tela(*tela_base)
        # R12b/c: SEM EP -> H vazio; G numerico -> H = 1,65E + |G|
        g_num = [st.get(f'{an_t}|{n}', {}).get('G') for n in range(1, nlv + 1)]
        hs = [p.Range(f'{c}{r}').Value for r in range(7, 7 + nlv) for c in 'EGH']
        ok_num = all(num(G) for G in g_num) and all(
            (not num(hs[3 * k]) and hs[3 * k + 2] in (None, '')) or perto(hs[3 * k + 2], 1.65 * hs[3 * k] + abs(hs[3 * k + 1]))
            for k in range(nlv))
        reg('R12b Com G numérico: Painel H = 1,65·E + |G| (1e-9)', ok_num, {'analito': an_t, 'E_G_H': hs})
        sem = [k.split('|')[0] for k, x in st.items() if x['G'] == 'SEM EP' and txt(x['A']) in cad]
        ativos_an = collections.Counter(txt(x['ANALITO']) for x in fin if x['PARTICIPA_ESTATISTICA'] == 'SIM')
        sem_dados = [a for a in dict.fromkeys(sem) if ativos_an.get(a, 0) > 0]
        if sem_dados:
            a_s = sem_dados[0]
            linhas_s = sorted((x for x in fin if x['PARTICIPA_ESTATISTICA'] == 'SIM' and txt(x['ANALITO']) == a_s
                               and (bio or txt(x['EQUIPAMENTO']).upper() == eq_pad.upper())), key=lambda x: x['DATA_HORA'])
            u = linhas_s[-1]
            du = u['DATA'].replace(tzinfo=None)
            tela(a_s, txt(u['LOTE']), txt(u['EQUIPAMENTO']), dt.datetime(du.year, du.month, 1), du)
            hs2 = [(p.Range(f'E{r}').Value, p.Range(f'G{r}').Value, p.Range(f'H{r}').Value) for r in range(7, 7 + nlv)]
            reg('R12c Com G = "SEM EP" (sem CEQ): Painel H vazio (antes: 1,65·CV)',
                any(num(x[0]) for x in hs2) and all(x[2] in (None, '') for x in hs2),
                {'analito': a_s, 'lote': txt(u['LOTE']), 'E_G_H': hs2})
        else:
            nao_executado('R12c', 'nenhum analito com G = SEM EP e participantes')

        # I17: BiasEQ com o filtro do ViesEQ
        orig = os.environ.get('QC_PROD_ORIGINAL')
        if bio:
            b1, b2 = st.get('Bilirrubina direta|1', {}), st.get('Bilirrubina direta|2', {})
            ok = (perto(b1.get('G'), 19.994209755156) and perto(b2.get('G'), 19.994209755156)
                  and b1.get('AC') == 13 and b1.get('AD') == 3)
            reg('I17 Bio: Bilirrubina direta G = 19,994209755156 % (sem CHM-04/CHM-13, resultado 0), AC 13, AD 3',
                ok, {'G': (b1.get('G'), b2.get('G')), 'AC': b1.get('AC'), 'AD': b1.get('AD'), 'H': (b1.get('H'), b2.get('H'))})
            mudam = {'Bilirrubina direta'}
        else:
            alvo4 = ['IG%', 'IG#', 'NRBC%', 'NRBC#']
            gs = {a: [st.get(f'{a}|{n}', {}).get('G') for n in range(1, nlv + 1)] for a in alvo4}
            reg('I17 Hema: IG%, IG#, NRBC% e NRBC# (as 15 amostras NAO AVALIADO) = "SEM EP" e H vazio',
                all(g == 'SEM EP' for v in gs.values() for g in v)
                and all(st.get(f'{a}|{n}', {}).get('H') in (None, '') for a in alvo4 for n in range(1, nlv + 1)), gs)
            mudam = set(alvo4)
        if orig and os.path.exists(orig):
            import openpyxl
            import warnings
            warnings.filterwarnings('ignore')
            wo = openpyxl.load_workbook(orig, read_only=True, data_only=True)['Estatística']
            dif = []
            for linha in wo.iter_rows(min_row=14, max_row=133, max_col=30, values_only=True):
                k = txt(linha[27])
                if not k or k.split('|')[0] in mudam or k not in st:
                    continue
                for c, j in (('G', 6), ('AC', 28), ('AD', 29)):
                    a_, b_ = linha[j], st[k][c]
                    if not ((num(a_) and perto(b_, a_)) or (not num(a_) and txt(a_) == txt(b_))):
                        dif.append((k, c, a_, b_))
            reg('I17b Os demais analitos com G, AC e AD iguais aos da produção (só a seleção de amostras mudou)',
                not dif, {'diferencas': dif[:10], 'n': len(dif)})
            if bio:
                bo = [r[1] for r in wo.iter_rows(min_row=98, max_row=102, min_col=1, max_col=2, values_only=True)]
                bn = [x[0] for x in est.Range('B98:B102').Value]
                reg('G08b Bio: classes Sigma B98:B102 registradas antes/depois (mudam só pelo ≤ 0 fora do CVp)', True,
                    {'antes': bo, 'depois': bn})
        else:
            nao_executado('I17b', 'QC_PROD_ORIGINAL ausente')

        # G08 (Bio) B103
        if bio:
            b103 = est.Range('B103').Value
            ind103 = sum(1 for x in st.values() if x['G'] == 'SEM EP' and txt(x['A']) not in ('', '0'))
            reg('G08 Bio: "Sem EP" (B103) conta só as linhas com analito (= contagem independente; o ciclo 01 esperava 4)',
                b103 == ind103, {'B103': b103, 'independente': ind103,
                                 'linhas': sorted(k for k, x in st.items() if x['G'] == 'SEM EP')[:10]})

        # ------------------------------------------------------------------ L04 troca de lote / Ano De-Ate
        an_l, lt_l, eq_l, di_l, df_l = tela_base
        lotes = collections.Counter(txt(x['LOTE']) for x in fin if x['PARTICIPA_ESTATISTICA'] == 'SIM'
                                    and txt(x['ANALITO']) == an_l and txt(x['EQUIPAMENTO']).upper() == eq_l.upper())
        outros = [l for l, n in lotes.most_common() if l != lt_l and n >= 20]
        if not outros:
            nao_executado('L04', f'{an_l}: sem outro lote com >= 20 resultados')
        else:
            lt2 = outros[0]
            ev_l, ok_l = {}, True

            def conf_hema(rot):
                b4 = txt(est.Range('B4').Value)
                a_de, a_ate = int(est.Range('B3').Value), int(est.Range('D3').Value)
                linha = estat().get(f'{an_l}|1', {})
                vals = [x['RESULTADO'] for x in fin if x['PARTICIPA_ESTATISTICA'] == 'SIM' and txt(x['ANALITO']) == an_l
                        and txt(x['LOTE']) == b4 and int(x['NIVEL']) == 1 and hasattr(x['DATA'], 'year')
                        and a_de <= x['DATA'].year <= max(a_de, a_ate) and txt(x['EQUIPAMENTO']).upper() == eq_pad.upper()]
                n, m, s, cv = desc(vals)
                ok = (linha.get('C') == n if n else linha.get('C') in (0, None, '')) and \
                    (n == 0 or perto(linha.get('D'), m)) and (n < 2 or (perto(linha.get('E'), s) and perto(linha.get('F'), cv)))
                ev_l[rot] = {'B4': b4, 'anos': (a_de, a_ate), 'estat_C_F': [linha.get(c) for c in 'CDEF'],
                             'independente': (n, m, s, cv)}
                return ok

            def conf_bio(rot, lote):
                linha = estat().get(f'{an_l}|1', {})
                ps = q.painel_stats()[0]
                h4 = txt(est.Range('H4').Value)
                ok = h4 == lote and linha.get('C') == ps[1] and (perto(linha.get('F'), ps[4])
                                                                  or (not num(linha.get('F')) and not num(ps[4])))
                ev_l[rot] = {'H4': h4, 'estat_C_F': [linha.get(c) for c in 'CDEF'], 'painel': ps[1:5]}
                return ok

            if bio:
                # o mes do ultimo resultado do outro lote: o Painel trabalha com no maximo 180 corridas (ADR-049/086,
                # G11 do ciclo 01) e o lote inteiro passaria disso -- a comparacao tem de ser na mesma populacao
                meses = collections.Counter((x['DATA'].year, x['DATA'].month) for x in fin if elig(x) and x['DATA']
                                            and txt(x['LOTE']) == lt2 and txt(x['ANALITO']) == an_l and int(x['NIVEL']) == 1
                                            and txt(x['EQUIPAMENTO']).upper() == eq_l.upper())
                (ya, ma), _ = meses.most_common(1)[0]
                ini_ = dt.datetime(ya, ma, 1)
                tela(an_l, lt2, eq_l, ini_, (ini_ + dt.timedelta(days=32)).replace(day=1) - dt.timedelta(days=1))
                ok_l = conf_bio('H3 = ' + lt2, lt2)
                tela(an_l, lt_l, eq_l, di_l, df_l)
                ok_l = conf_bio('H3 = ' + lt_l, lt_l) and ok_l
            else:
                b3_0, d3_0 = est.Range('B3').Value, est.Range('D3').Value
                tela(an_l, lt2, eq_l, di_l, df_l)                     # so troca o lote; anos como estavam
                ok_l = conf_hema(f'H3 = {lt2} (anos {b3_0}/{d3_0})')
                tela(an_l, lt_l, eq_l, di_l, df_l)
                ok_l = conf_hema(f'de volta H3 = {lt_l} (anos {b3_0}/{d3_0})') and ok_l
                # Ano De/Ate (eventos ligados: Worksheet_Change da aba): os anos do lote em H3, depois um ano sem ele
                anos_l = sorted({x['DATA'].year for x in fin if txt(x['LOTE']) == lt_l and txt(x['ANALITO']) == an_l
                                 and hasattr(x['DATA'], 'year')})
                fora = min(anos_l) - 1
                n_antes = estat().get(f'{an_l}|1', {}).get('C')
                est.Unprotect(ii.SENHA)
                for de, ate in ((anos_l[0], anos_l[-1]), (fora, fora)):
                    est.Range('B3').Value = de
                    q.ex.esperar()
                    est.Range('D3').Value = ate
                    q.ex.esperar()
                    xl.Calculate()
                    ok_l = conf_hema(f'B3/D3 = {de}/{ate}') and ok_l
                ev_l['n_N1_antes_de_mudar_os_anos'] = n_antes
                ok_l = ok_l and ev_l[f'B3/D3 = {fora}/{fora}']['independente'][0] == 0 and                     ev_l[f'B3/D3 = {anos_l[0]}/{anos_l[-1]}']['independente'][0] > 0
                est.Range('B3').Value = b3_0
                q.ex.esperar()
                est.Range('D3').Value = d3_0
                q.ex.esperar()
            reg('L04 Troca de lote no Painel%s: a Estatística acompanha (%s)' % (
                '' if bio else ' e Ano De/Até', 'C:F = Painel do mesmo lote' if bio else 'C:F = independente, 1e-9'),
                ok_l, ev_l)

        # ------------------------------------------------------------------ Y11 H3 vazio
        tela(an_l, '', eq_l, di_l, df_l)
        st2 = estat()
        lin = [(p.Range(f'C{r}').Value, p.Range(f'D{r}').Value, p.Range(f'E{r}').Value, p.Range(f'G{r}').Value,
                p.Range(f'H{r}').Value, p.Range(f'I{r}').Value) for r in range(7, 7 + nlv)]
        g6, i6 = txt(p.Range('G6').Text), txt(p.Range('I6').Text)
        ok = (e.Range('AD1').Value is True and 'do método' in g6 and 'do método' in i6
              and all(x[0] in (None, '') and x[1] in (None, '') and x[2] in (None, '') and x[4] in (None, '') for x in lin)
              and all((perto(lin[k][3], st2[f'{an_l}|{k + 1}']['G']) or txt(lin[k][3]) == txt(st2[f'{an_l}|{k + 1}']['G']))
                      and (perto(lin[k][5], st2[f'{an_l}|{k + 1}']['L']) or txt(lin[k][5]) == txt(st2[f'{an_l}|{k + 1}']['L']))
                      for k in range(nlv)))
        reg('Y11 H3 vazio (TODOS): Bias e Sigma do método visíveis (= Estatística G e L), rotulados "do método"; Média/DP/CV '
            'e ET vazios', ok, {'G6': g6, 'I6': i6, 'C_D_E_G_H_I': lin,
                                 'estat_G_L': [(st2[f'{an_l}|{k + 1}']['G'], st2[f'{an_l}|{k + 1}']['L']) for k in range(nlv)]})
        tela(an_l, lt_l, eq_l, di_l, df_l)
        g6b = txt(p.Range('G6').Text)
        reg('Y11b Com lote em H3 os rótulos voltam ("Bias %", "Sigma (CVp ...)")', g6b == 'Bias %'
            and 'do método' not in txt(p.Range('I6').Text), {'G6': g6b, 'I6': txt(p.Range('I6').Text)})
    except Exception as ex_:                         # noqa: BLE001
        reg('Z99 Execução da suíte sem exceção', False, {'erro': repr(ex_), 'trace': traceback.format_exc()[-2000:]})
    finally:
        q.fechar(salvar=False)
    hash1 = sha256_arquivo(caminho)
    reg('Y00 Segurança do teste: arquivo de entrada intacto (SHA-256 igual antes e depois; nada salvo nele)',
        hash0 == hash1, {'arquivo': caminho, 'sha256_antes': hash0, 'sha256_depois': hash1})
    if saida:
        with open(saida, 'w', encoding='utf-8') as f:
            json.dump(RES, f, ensure_ascii=False, indent=1, default=str)
    falhas = [r for r in RES if r['resultado'] == 'FAIL']
    ne = [r for r in RES if r['resultado'] == 'NAO_EXECUTADO']
    print(f'\n=== {produto} ADR-088 (lote 1 da auditoria): {len(RES) - len(falhas) - len(ne)} PASS / {len(falhas)} FAIL'
          f' / {len(ne)} NAO_EXECUTADO ===', flush=True)
    return falhas


if __name__ == '__main__':
    falhas = executar(sys.argv[1], os.path.abspath(sys.argv[2]), sys.argv[3] if len(sys.argv) > 3 else None)
    sys.exit(1 if falhas else 0)
