# -*- coding: utf-8 -*-
"""QA do ADR-074 -- aba Analitos com UM BLOCO POR LOTE. Roda numa COPIA instalada; nada e salvo nela.

Uso: python qa_blocos_lote.py <Bioquimica|Hematologia> <copia_instalada.xlsm> [saida.json]

Os testes de aceitacao A..J do pedido (14/10/2026), pelo caminho do usuario (eventos ligados: digitar no bloco,
escolher o lote no Painel, trocar o ano de referencia):
  BA Persistencia: cada lote do cadastro tem UM bloco, na ordem do cadastro; E:J de todos os blocos = LotesStore
  BB Independencia: Media/DP digitados no bloco do lote X gravam SO em X (store, trilha, retrato, origem); o 1o
     lote nao muda; valor invalido nao grava e a formula volta
  BC Graficos: lote X no Painel -> Calc (AX1/AY1, AZ1/BA1) e as 7 linhas de cada nivel = media de X +- k DP;
     de volta ao 1o lote -> os parametros dele
  BD Parametros ausentes: lote sem Media/DP -> bloco vazio, Calc sem media, Painel "SEM MEDIA/DP"; nada herdado
  BE Inclusao: lote novo no cadastro -> UM bloco novo, com validade, anos e botoes; sincronizar de novo nao duplica
  BF Referencias anuais: ano CLIA 2026 no bloco de X -> especificacoes do bloco = biblioteca 2026; ETp do lote
     muda; trilha; retrato na store; motor do lote em analise segue o ano do lote; ano fora da biblioteca recusado
  BG Abrir/recolher: so as linhas do bloco; formulas intactas; abrir todos / recolher todos
  BH Integridade: nenhum #REF! novo em formulas e nomes; nenhum erro nos blocos; nomes do motor na Especificações
  BI Reabertura: salvar copia, fechar, abrir: blocos, Media/DP, anos e botoes continuam; digitar continua gravando
  BJ Regressao: ETp da Estatistica = ETp em uso do motor; LotesStore integra; arquivo de entrada intacto
Bioquimica e Hematologia (esta depois do ADR-075).
"""
import datetime as dt
import json
import os
import sys
import tempfile
import traceback

AQUI = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, AQUI)
from qa_final import QA, reg, RES, sha256_arquivo  # noqa: E402

R0, PASSO, NAN = 2, 54, 40
O_AN, O_ANO, O_NOMES, O_CLIA = 4, 45, 46, 47


def num(v):
    return isinstance(v, (int, float)) and not isinstance(v, bool) and not (isinstance(v, int) and v < -2146820000)


def erro(v):
    return isinstance(v, int) and v < -2146820000


def perto(a, b, tol=1e-9):
    return num(a) and num(b) and abs(a - b) <= tol * max(1.0, abs(b))


def txt(v):
    if isinstance(v, float) and v == int(v):
        return str(int(v))
    return str(v if v is not None else '').strip()


def blocos(b):
    out, h = [], R0
    while txt(b.Cells(h, 1).Value).upper() == 'LOTE':
        out.append((txt(b.Cells(h, 2).Value), h))
        h += PASSO
    return out


def cadastro(q):
    vistos, out = set(), []
    for v in q.nome('regLoteCol').Value:
        s = txt(v[0])
        if s and s.upper() not in vistos:
            vistos.add(s.upper())
            out.append(s)
    return out


def store(q):
    ls = q.ws('LotesStore')
    ult = ls.Cells(ls.Rows.Count, 16).End(-4162).Row
    out = {}
    for i, linha in enumerate(ls.Range(f'A2:W{max(2, ult)}').Value, 2):
        if linha[15]:
            out[str(linha[15]).upper()] = (i, linha)
    return out


def trilha(q, n=12):
    aud = q.ws('Audit_Log')
    ua = aud.Cells(aud.Rows.Count, 1).End(-4162).Row
    return ' '.join(str(aud.Cells(i, c).Value or '') for i in range(max(2, ua - n), ua + 1) for c in range(1, 25))


def escolher_lotes(q, rows, eq, cad, lote0):
    cont = {}
    for r in rows:
        if str(r.get('PARTICIPA_ESTATISTICA', '')).upper() != 'SIM':
            continue
        if eq and str(r.get('EQUIPAMENTO', '')).upper() != eq:
            continue
        cont[txt(r['LOTE'])] = cont.get(txt(r['LOTE']), 0) + 1
    cand = sorted(((cont.get(l, 0), l) for l in cad if l != lote0), reverse=True)
    if not cand or cand[0][0] == 0:
        raise RuntimeError('pre-condicao: nenhum outro lote do cadastro com resultados')
    return cand[0][1], cand[0][0]


def selecionar_lote(q, lote, pos_an):
    p = q.ws('Painel')
    p.Range('H3').Value = "'" + lote                         # como o usuario: eventos ligados
    q.ex.esperar()
    p.Range('B3').Value = pos_an
    q.run('mEstatistica.PainelMudou', teto=300)
    q.ex.xl.Calculate()
    q.ex.esperar()


def linhas_calc(q):
    c = q.ws('Calc')
    v = c.Range('Q3:W182').Value
    v2 = c.Range('AM3:AS182').Value
    return v, v2


def executar(produto, caminho, saida):
    hash0 = sha256_arquivo(caminho)
    q = QA(produto, caminho)
    reaberto = None
    copia = None
    try:
        q.run('mSeguranca.ReprotectAll')                  # sessao de usuario: macros escrevem (UserInterfaceOnly)
        nomes_abas = [w.Name for w in q.wb.Worksheets]
        # media/DP do grafico no Calc: Bioquimica AX1/AY1 (N1), AZ1/BA1 (N2); Hematologia BT1/BU1, BV1/BW1
        cel_md = ('AX1', 'AY1', 'AZ1', 'BA1') if produto.startswith('Bio') else ('BT1', 'BU1', 'BV1', 'BW1')
        b, e, p = q.ws('Analitos'), q.espec(), q.ws('Painel')
        eq = str(q.run('mEstatistica.EquipFiltro', teto=60) or '').strip().upper()
        rows = q.ler('tblCQ_Final')
        cad = cadastro(q)
        lote0 = txt(q.nome('loteParam').Value)
        catalogo = [txt(v[0]) for v in e.Range('A4:A43').Value]

        # ------------------------------------------------------------ BA persistencia
        bl = blocos(b)
        st = store(q)
        dif = []
        for lote, h in bl:
            vals = b.Range(b.Cells(h + O_AN, 1), b.Cells(h + O_AN + NAN - 1, 10)).Value
            for linha in vals:
                an = txt(linha[0])
                if not an:
                    continue
                esp = st.get(f'{lote}|{an}'.upper(), (0, (None,) * 23))[1][2:8]
                for j in range(6):
                    v, s = linha[4 + j], esp[j]
                    if (num(s) and not perto(v, s)) or (not num(s) and v not in ('', None)):
                        dif.append((lote, an, j, v, s))
        conf = str(q.run('mLotes.ConferirLotesStore', teto=60))
        com_formula = [lote for lote, h in bl
                       if any(str(f).startswith('=') for linha in b.Range(b.Cells(h + O_AN, 5), b.Cells(h + O_AN + NAN - 1, 10)).Formula
                              for f in linha)]
        reg('BA Persistência: cada lote do cadastro tem UM bloco (na ordem do cadastro); Média/DP (E:J) de TODOS os blocos '
            'são VALORES (sem fórmula) iguais aos da LotesStore (chave LOTE|ANALITO, sem duplicata)',
            [l for l, _ in bl] == cad and len(set(l for l, _ in bl)) == len(bl) and not dif and conf.startswith('ok')
            and not com_formula,
            {'blocos': len(bl), 'cadastro': len(cad), 'primeiros': [l for l, _ in bl][:5], 'divergencias': dif[:6],
             'ConferirLotesStore': conf, 'blocos_com_formula_em_EJ': com_formula[:6]})

        # ------------------------------------------------------------ BN preenchimento inicial (ADR-076)
        marca = None
        try:
            marca = str(q.wb.Names('blPreencHist').RefersTo)
        except Exception:
            pass
        ref = marca.split('ref ')[-1].strip('"= ') if marca else ''
        grupos = {}
        for r_ in rows:
            if str(r_.get('PARTICIPA_ESTATISTICA', '')).upper() != 'SIM' or not num(r_.get('RESULTADO')):
                continue
            k_ = (txt(r_['LOTE']).upper(), txt(r_['ANALITO']).upper(), int(r_['NIVEL']))
            grupos.setdefault(k_, []).append(float(r_['RESULTADO']))
        import statistics as stt
        conf_n, dif_n, lotes_n = 0, [], set()
        for chave, (i_, linha) in st.items():
            if not str(linha[22] or '').startswith('preenchimento inicial'):
                continue
            lote_, an_ = txt(linha[0]), txt(linha[14])
            for nv in range(1, q.nlv + 1):
                m_, s_ = linha[2 + (nv - 1) * 2], linha[3 + (nv - 1) * 2]
                if not num(m_):
                    continue
                xs = grupos.get((lote_.upper(), an_.upper(), nv), [])
                ok_m = bool(xs) and perto(m_, stt.mean(xs), 1e-9)
                dref = q.run('mLotes.DPDoLote', ref, an_, nv) if ref and lote_ != ref else ''
                if f'N{nv}: ' in str(linha[22]) and 'DP do lote' in str(linha[22]).split(f'N{nv}: ')[1].split(' | ')[0]:
                    ok_s = perto(s_, dref, 1e-9)
                elif num(s_):
                    ok_s = len(xs) >= 20 and perto(s_, stt.stdev(xs), 1e-7)
                else:
                    ok_s = len(xs) < 20 and not num(dref)
                conf_n += 1
                lotes_n.add(lote_)
                if not (ok_m and ok_s):
                    dif_n.append((lote_, an_, nv, m_, stt.mean(xs) if xs else None, s_, dref, len(xs)))
        reg('BN Preenchimento inicial (pedido de 15/10/2026): Média = média de todos os resultados Ativos do próprio lote '
            '(recalculada aqui de forma independente); DP = o do lote em uso (ou, sem ele, o DP do próprio lote com '
            'n >= 20); origem registrada; nada sobrescrito',
            bool(marca) and conf_n > 0 and not dif_n,
            {'marca': marca, 'lote_referencia_DP': ref, 'pares_conferidos': conf_n, 'lotes': len(lotes_n),
             'divergencias': dif_n[:6]})

        # ------------------------------------------------------------ BM largura X = W
        wx, ww = float(b.Columns('X').ColumnWidth), float(b.Columns('W').ColumnWidth)
        reg('BM Formatação: coluna X com a mesma largura da coluna W', abs(wx - ww) < 0.01, {'X': wx, 'W': ww})

        # ------------------------------------------------------------ BB independencia
        loteX, nX = escolher_lotes(q, rows, eq, cad, lote0)
        hX, h0 = dict(bl)[loteX], dict(bl)[lote0]
        an = next(a for a in catalogo if a and num(q.run('mLotes.MediaDoLote', lote0, a, 1))
                  and num(q.run('mLotes.MediaDoLote', lote0, a, 2)))
        i_an = catalogo.index(an)
        rX, r0 = hX + O_AN + i_an, h0 + O_AN + i_an
        m0 = [q.run('mLotes.MediaDoLote', lote0, an, n) for n in (1, 2)]
        s0 = [q.run('mLotes.DPDoLote', lote0, an, n) for n in (1, 2)]
        novos = [round(m0[0] * 1.1, 4), round(s0[0] * 1.3, 4), round(m0[1] * 0.9, 4), round(s0[1] * 1.2, 4)]
        for j, v in enumerate(novos):
            b.Cells(rX, 5 + j).Value = v                      # digitando no bloco do lote X
            q.ex.esperar()
        mX = [q.run('mLotes.MediaDoLote', loteX, an, n) for n in (1, 2)]
        sX = [q.run('mLotes.DPDoLote', loteX, an, n) for n in (1, 2)]
        m0b = [q.run('mLotes.MediaDoLote', lote0, an, n) for n in (1, 2)]
        q.ex.xl.Calculate()
        st = store(q)
        linX = st.get(f'{loteX}|{an}'.upper(), (0, [None] * 23))[1]
        sem_formula = not any(bool(b.Cells(rX, 5 + j).HasFormula) for j in range(4))
        tela_ok = all(perto(b.Cells(rX, 5 + j).Value, novos[j]) for j in range(4))
        tela0_ok = perto(b.Cells(r0, 5).Value, m0[0]) and perto(b.Cells(r0, 6).Value, s0[0])
        # invalido: DP negativo -> nao grava, a celula volta ao valor gravado
        b.Cells(rX, 6).Value = -1
        q.ex.esperar()
        inval_ok = perto(q.run('mLotes.DPDoLote', loteX, an, 1), novos[1]) and perto(b.Cells(rX, 6).Value, novos[1])
        reg('BB Independência: Média/DP digitados no bloco do lote X gravam só em X (LotesStore, trilha '
            'PARAMETRO_LOTE_ALTERADO, retrato "CLIA aaaa · VB aaaa", origem "bloco do lote") e o VALOR digitado fica na '
            'célula (sem fórmula); o 1º lote não muda; DP inválido não grava e a célula volta ao valor gravado',
            [perto(mX[0], novos[0]), perto(sX[0], novos[1]), perto(mX[1], novos[2]), perto(sX[1], novos[3])] == [True] * 4
            and m0b == m0 and sem_formula and tela_ok and tela0_ok and inval_ok
            and 'PARAMETRO_LOTE_ALTERADO' in trilha(q) and str(linX[22] or '').startswith('digitado na aba Analitos (bloco')
            and str(linX[17] or '').startswith('CLIA ') and ' VB ' in str(linX[17] or ''),
            {'loteX': loteX, 'resultados_X': nX, 'analito': an, 'digitado': novos, 'lido_X': mX + sX, 'lote0': lote0,
             'lote0_antes': m0, 'lote0_depois': m0b, 'celulas_sem_formula': sem_formula, 'origem': linX[22],
             'retrato': linX[16:20], 'dp_invalido_recusado': inval_ok})

        # ------------------------------------------------------------ BC graficos
        selecionar_lote(q, loteX, i_an + 1)
        c = q.ws('Calc')
        ax = [c.Range(x).Value for x in cel_md]
        v1, v2 = linhas_calc(q)
        lin1 = [l for l in v1 if num(l[3])]
        lin2 = [l for l in v2 if num(l[3])]
        exp1 = [mX[0] + k * sX[0] for k in range(-3, 4)]
        exp2 = [mX[1] + k * sX[1] for k in range(-3, 4)]
        ok1 = bool(lin1) and all(all(perto(a, e_) for a, e_ in zip(l, exp1)) for l in lin1)
        ok2 = bool(lin2) and all(all(perto(a, e_) for a, e_ in zip(l, exp2)) for l in lin2)
        st7 = [str(p.Range(x).Value or '') for x in ('J7', 'J8')]
        selecionar_lote(q, lote0, i_an + 1)
        ax0 = [c.Range(x).Value for x in cel_md]
        reg('BC Gráficos: com o lote X no Painel, média e DP do gráfico (Calc, N1 e N2) e as 7 linhas de cada '
            'nível (-3DP..+3DP) são as do lote X; de volta ao 1º lote, as dele',
            perto(ax[0], mX[0]) and perto(ax[1], sX[0]) and perto(ax[2], mX[1]) and perto(ax[3], sX[1])
            and ok1 and ok2 and not any('SEM M' in s.upper() for s in st7)
            and perto(ax0[0], m0[0]) and perto(ax0[1], s0[0]) and perto(ax0[2], m0[1]) and perto(ax0[3], s0[1]),
            {'lote': loteX, 'analito': an, 'Calc_media_dp_N1_N2': dict(zip(cel_md, ax)), 'pontos_N1': len(lin1), 'pontos_N2': len(lin2),
             'linhas_N1_ok': ok1, 'linhas_N2_ok': ok2, 'status_painel': st7, 'volta_lote0': ax0})

        # ------------------------------------------------------------ BC2 alternancia entre lotes
        # o analito com Media/DP (N1) no maior numero de lotes anteriores (pela LotesStore), alem do 1o lote
        st2 = store(q)
        def tem_n1(l, a):
            lin = st2.get(f'{l}|{a}'.upper(), (0, [None] * 23))[1]
            return num(lin[2]) and num(lin[3]) and lin[3] > 0
        melhor = max(((sum(1 for l in cad if l not in (lote0, loteX) and tem_n1(l, a)), a)
                      for a in catalogo if a and tem_n1(lote0, a)), default=(0, None))
        an_alt = melhor[1]
        outros = [l for l in cad if l not in (lote0, loteX) and an_alt and tem_n1(l, an_alt)][:2]
        if len(outros) < 2 and an_alt and tem_n1(loteX, an_alt):
            outros = (outros + [loteX])[:2]
        seq = ([lote0] + outros) * 2
        dif2 = []
        for l in seq:
            selecionar_lote(q, l, catalogo.index(an_alt) + 1)
            got = [c.Range(x).Value for x in cel_md]
            exp = [q.run('mLotes.MediaDoLote', l, an_alt, 1), q.run('mLotes.DPDoLote', l, an_alt, 1),
                   q.run('mLotes.MediaDoLote', l, an_alt, 2), q.run('mLotes.DPDoLote', l, an_alt, 2)]
            stl = [str(p.Range(x).Value or '') for x in ('J7', 'J8')]
            ok_n1 = perto(got[0], exp[0]) and perto(got[1], exp[1]) and 'SEM M' not in stl[0].upper()
            ok_n2 = (perto(got[2], exp[2]) and perto(got[3], exp[3])) if num(exp[3]) else (got[2] in ('', None))
            if not (ok_n1 and ok_n2):
                dif2.append((l, got, exp, stl))
        selecionar_lote(q, lote0, i_an + 1)
        reg('BC2 Alternância: lote em uso e dois lotes anteriores, duas voltas -- a média e o DP do gráfico são sempre os '
            'do lote escolhido (N1; N2 quando o lote o tem); nenhuma vez "SEM MÉDIA/DP" no N1; nada preso ao lote anterior',
            len(outros) == 2 and not dif2, {'sequencia': seq, 'analito': an_alt, 'divergencias': dif2[:4]})

        # ------------------------------------------------------------ BE inclusao (lote novo, sem resultados)
        n0 = len(blocos(b))
        novo_lote = 'QA074'
        r = str(q.run('mApp.CadastrarLote', novo_lote, '31/12/2027', teto=300))
        bl2 = blocos(b)
        hN = dict(bl2).get(novo_lote)
        val = b.Cells(hN, 6).Value if hN else None
        nomes_shp = [s.Name for s in b.Shapes]
        n_sinc = q.run('mLotesBlocos.SincronizarBlocos', teto=120)
        reg('BE Inclusão: lote cadastrado ganha UM bloco novo, no fim, com a validade do cadastro, os anos de referência '
            'e os botões abrir/recolher e CIQ; sincronizar de novo não duplica',
            r.startswith('OK') and len(bl2) == n0 + 1 and bl2[-1][0] == novo_lote and hasattr(val, 'year')
            and val.year == 2027 and f'blkTgl_{hN}' in nomes_shp and f'blkCIQ_{hN}' in nomes_shp
            and num(b.Cells(hN + O_ANO, 4).Value) and num(b.Cells(hN + O_ANO, 10).Value) and int(n_sinc or 0) == 0
            and len(blocos(b)) == n0 + 1,
            {'retorno': r, 'blocos_antes': n0, 'blocos_depois': len(bl2), 'linha': hN, 'validade': str(val),
             'anos': (b.Cells(hN + O_ANO, 4).Value, b.Cells(hN + O_ANO, 10).Value) if hN else None,
             'sincronizar_de_novo': n_sinc})

        # ------------------------------------------------------------ BD ausentes (o lote novo, sem Media/DP)
        loteY, hY = novo_lote, hN
        selecionar_lote(q, loteY, i_an + 1)
        axY = [c.Range(x).Value for x in cel_md]
        v1, _ = linhas_calc(q)
        sem_linha = all(not num(l[3]) for l in v1)
        stY = [str(p.Range(x).Value or '') for x in ('J7', 'J8')]
        rY = hY + O_AN + i_an
        blocoY = b.Range(b.Cells(rY, 5), b.Cells(rY, 10)).Value[0] + b.Range(b.Cells(rY, 25), b.Cells(rY, 38)).Value[0]
        reg('BD Parâmetros ausentes: lote sem Média/DP -> Calc sem média/DP, nenhuma linha de limite, Painel "SEM MÉDIA/DP" '
            'e bloco vazio (nada herdado de outro lote)',
            all(v in ('', None) for v in axY) and sem_linha and all('SEM M' in s.upper() for s in stY)
            and all(v in ('', None) for v in blocoY) and q.run('mLotes.MediaDoLote', loteY, an, 1) in ('', None),
            {'loteY': loteY, 'Calc': axY, 'status_painel': stY, 'bloco': [v for v in blocoY if v not in ('', None)][:4]})
        selecionar_lote(q, lote0, i_an + 1)

        # ------------------------------------------------------------ BF referencias anuais
        yX = hX + O_ANO
        lib = {}                                             # ano -> deslocamento da biblioteca (Especificações!B46:AO89)
        for k, rr in enumerate((46, 55, 64, 73, 82)):
            lib[int(e.Range(f'B{rr}').Value)] = rr
        anos = sorted(lib)
        ano_novo = next(a for a in anos if a != int(b.Cells(yX, 4).Value))
        b.Cells(yX, 4).Value = ano_novo
        q.ex.esperar()
        q.ex.xl.Calculate()
        rr = lib[ano_novo]
        nomes_lib = [txt(v) for v in e.Range(f'B{rr + 1}:AO{rr + 1}').Value[0]]
        clia_lib = e.Range(f'B{rr + 2}:AO{rr + 2}').Value[0]
        nomes_bl = [txt(v) for v in b.Range(b.Cells(hX + O_NOMES, 2), b.Cells(hX + O_NOMES, 41)).Value[0]]
        clia_bl = b.Range(b.Cells(hX + O_CLIA, 2), b.Cells(hX + O_CLIA, 41)).Value[0]
        dif6 = []
        for nm_, v in zip(nomes_bl, clia_bl):
            if not nm_:
                continue
            esp_ = clia_lib[nomes_lib.index(nm_)] if nm_ in nomes_lib else None
            if (num(esp_) and not perto(v, esp_)) or (not num(esp_) and v not in ('', None)):
                dif6.append((nm_, v, esp_))
        k_ok = []
        for i, nm_ in enumerate(catalogo):
            if nm_:
                esp_ = clia_lib[nomes_lib.index(nm_)] if nm_ in nomes_lib else None
                kv = b.Cells(hX + O_AN + i, 11).Value
                k_ok.append((num(esp_) and perto(kv, esp_)) or (not num(esp_) and kv in ('', None)))
        st = store(q)
        retr = str(st.get(f'{loteX}|{an}'.upper(), (0, [None] * 23))[1][17] or '')
        # motor: o lote em analise passa a ser X -> anoCLIAemUso = ano do bloco de X
        selecionar_lote(q, loteX, i_an + 1)
        ano_motor = q.nome('anoCLIAemUso').Value
        k_motor = [(e.Range(f'K{4 + i}').Value, b.Cells(hX + O_AN + i, 11).Value) for i, nm_ in enumerate(catalogo) if nm_]
        motor_ok = all((perto(a, bb) if num(bb) else a in (0, '', None)) for a, bb in k_motor)
        # ano fora da biblioteca
        vb0 = b.Cells(yX, 10).Value
        b.Cells(yX, 10).Value = 2099
        q.ex.esperar()
        recusado = b.Cells(yX, 10).Value == vb0
        selecionar_lote(q, lote0, i_an + 1)
        reg('BF Referências anuais: o ano CLIA escolhido no bloco do lote X traz as especificações daquele ano da '
            'biblioteca (linhas do bloco e coluna ETp CLIA), fica na trilha e no retrato da LotesStore; com X em análise '
            'o motor usa o ano de X; ano fora da biblioteca é recusado',
            not dif6 and all(k_ok) and 'ANO_REFERENCIA_LOTE_ALTERADO' in trilha(q) and retr.startswith(f'CLIA {ano_novo}')
            and num(ano_motor) and int(ano_motor) == ano_novo and motor_ok and recusado,
            {'lote': loteX, 'ano_novo': ano_novo, 'divergencias_linha_CLIA': dif6[:6], 'ETp_CLIA_ok': sum(k_ok),
             'retrato': retr, 'anoCLIAemUso_com_X': ano_motor, 'motor_igual_bloco': motor_ok, 'ano_2099_recusado': recusado})

        # ------------------------------------------------------------ BG abrir/recolher
        def formulas(h):
            return sum(1 for linha in b.Range(b.Cells(h, 1), b.Cells(h + 53, 56)).Formula
                       for f in linha if str(f).startswith('='))
        f_antes = formulas(hX)
        q.run('mLotesBlocos.AbrirBloco', hX, True)
        aberto = [bool(b.Rows(hX + k).Hidden) for k in (1, 20, 53)]
        vis0 = bool(b.Rows(h0 + 10).Hidden)
        q.run('mLotesBlocos.AbrirBloco', hX, False)
        fechado = [bool(b.Rows(hX + k).Hidden) for k in (1, 20, 53)]
        cab_vis = not bool(b.Rows(hX).Hidden) and not bool(b.Rows(hX + PASSO).Hidden)
        vis0_depois = bool(b.Rows(h0 + 10).Hidden)
        txt_tgl = str(b.Shapes(f'blkTgl_{hX}').TextFrame2.TextRange.Text)
        f_depois = formulas(hX)
        q.run('mLotesBlocos.ExpandirTodos')
        todos_abertos = not any(bool(b.Rows(h + 10).Hidden) for _, h in blocos(b))
        q.run('mLotesBlocos.RecolherTodos')
        todos_fechados = all(bool(b.Rows(h + 10).Hidden) and not bool(b.Rows(h).Hidden) for _, h in blocos(b))
        reg('BG Abrir/recolher: o botão do bloco mexe só nas linhas DELE (o cabeçalho e os outros blocos ficam), nenhuma '
            'fórmula se perde; abrir todos / recolher todos',
            aberto == [False] * 3 and fechado == [True] * 3 and cab_vis and vis0 == vis0_depois and f_antes == f_depois
            and f_antes > 1000 and 'abrir' in txt_tgl and todos_abertos and todos_fechados,
            {'bloco': loteX, 'aberto': aberto, 'recolhido': fechado, 'formulas_antes_depois': (f_antes, f_depois),
             'texto_botao': txt_tgl.encode('ascii', 'replace').decode(), 'todos_abertos': todos_abertos, 'todos_recolhidos': todos_fechados})

        # ------------------------------------------------------------ BH integridade
        refs = []
        for aba in ('Analitos', 'Especificações', 'Estatística', 'Calc', 'Painel', 'Início', 'Configuração'):
            if aba not in nomes_abas:
                continue
            fs = q.ws(aba).UsedRange.Formula
            fs = fs if isinstance(fs, tuple) else ((fs,),)
            refs += [(aba, str(f)[:80]) for linha in fs for f in linha if '#REF!' in str(f)]
        nomes_ref = []
        for n in q.wb.Names:
            try:
                if '#REF' in str(n.RefersTo):
                    nomes_ref.append(n.Name)
            except Exception:
                pass
        base = os.environ.get('QC_PROD_ORIGINAL')
        nomes_ref_base = None
        if base and os.path.exists(base):
            import zipfile
            import re
            x = zipfile.ZipFile(base).read('xl/workbook.xml').decode('utf-8', 'replace')
            nomes_ref_base = [m.group(1) for m in re.finditer(r'<definedName name="([^"]+)"[^>]*>[^<]*#REF', x)]
            nomes_ref = [n for n in nomes_ref if n not in nomes_ref_base]
        erros_bl = []
        for lote, h in blocos(b):
            for linha in b.Range(b.Cells(h, 1), b.Cells(h + 53, 56)).Value:
                if any(erro(v) for v in linha):
                    erros_bl.append(lote)
                    break
        motor_nomes = {nm_: str(q.wb.Names(nm_).RefersTo) for nm_ in ('etpOficial', 'cvtpOficial', 'espFonte', 'lstAnalitos')}
        reg('BH Integridade: nenhum #REF! em fórmulas (Analitos, Especificações, Estatística, Calc, Painel, Início, '
            'Configuração) nem nome novo quebrado; nenhum erro nos blocos; nomes do motor na Especificações',
            not refs and not nomes_ref and not erros_bl and all('Especifica' in v for v in motor_nomes.values()),
            {'formulas_REF': refs[:6], 'nomes_REF_novos': nomes_ref[:6], 'nomes_REF_ja_na_producao': nomes_ref_base,
             'blocos_com_erro': erros_bl[:6], 'nomes_motor': motor_nomes})

        # ------------------------------------------------------------ BJ regressao (antes de fechar)
        e_ws = q.ws('Estatística')
        q.ex.xl.Calculate()
        estat = {}
        for linha in e_ws.Range(f'A14:K{13 + NAN * q.nlv}').Value:       # so a tabela principal (analito x nivel)
            if linha[0] and linha[1] == 1 and txt(linha[0]) not in estat:
                estat[txt(linha[0])] = linha[10]
        dif_j = []
        for i, nm_ in enumerate(catalogo):
            if nm_ and nm_ in estat:
                t_m = e.Range(f'T{4 + i}').Value
                if (num(t_m) and not perto(estat[nm_], t_m)) or (not num(t_m) and estat[nm_] != '-'):
                    dif_j.append((nm_, estat[nm_], t_m))
        conf = str(q.run('mLotes.ConferirLotesStore', teto=60))
        reg('BJ Regressão: o ETp da Estatística é o ETp em uso do motor (lote em análise) e a LotesStore continua íntegra',
            not dif_j and conf.startswith('ok'), {'divergencias': dif_j[:6], 'ConferirLotesStore': conf})

        # ------------------------------------------------------------ BI reabertura
        copia = os.path.join(tempfile.gettempdir(), f'qa074_reab_{os.getpid()}.xlsm')
        q.wb.SaveCopyAs(copia)
        q.fechar(salvar=False)
        q = None
        reaberto = QA(produto, copia)
        reaberto.run('mSeguranca.ReprotectAll')
        b2 = reaberto.ws('Analitos')
        bl3 = blocos(b2)
        h3 = dict(bl3).get(loteX)
        v_ok = all(perto(b2.Cells(h3 + O_AN + i_an, 5 + j).Value, novos[j]) for j in range(4)) if h3 else False
        ano_ok = h3 and int(b2.Cells(h3 + O_ANO, 4).Value) == ano_novo
        acoes = {s.Name: str(s.OnAction) for s in b2.Shapes if s.Name in (f'blkTgl_{h3}', f'blkCIQ_{h3}')}
        novo2 = round(novos[0] * 1.01, 4)
        b2.Cells(h3 + O_AN + i_an, 5).Value = novo2
        reaberto.ex.esperar()
        grava = perto(reaberto.run('mLotes.MediaDoLote', loteX, an, 1), novo2)
        reg('BI Reabertura: depois de salvar, fechar e abrir, os blocos, o Média/DP digitado, o ano de referência e os '
            'botões continuam; digitar no bloco continua gravando na LotesStore',
            len(bl3) == len(blocos(b2)) and len(bl3) >= len(cad) and v_ok and ano_ok and len(acoes) == 2
            and all('mLotesBlocos.' in a for a in acoes.values()) and grava,
            {'blocos': len(bl3), 'valores_ok': v_ok, 'ano': b2.Cells(h3 + O_ANO, 4).Value if h3 else None,
             'botoes': acoes, 'digitar_depois_de_reabrir': grava})
    except Exception as ex_geral:
        reg('B00 Execução interrompida: os testes seguintes a este ponto não rodaram', False,
            {'erro': f'{type(ex_geral).__name__}: {ex_geral}', 'traceback': traceback.format_exc()[-1500:]})
    finally:
        for x in (q, reaberto):
            if x is not None:
                try:
                    x.fechar(salvar=False)
                except Exception:
                    pass
        if copia and os.path.exists(copia):
            try:
                os.remove(copia)
            except Exception:
                pass
    hash1 = sha256_arquivo(caminho)
    reg('BK Segurança do teste: arquivo de entrada intacto (SHA-256 igual antes e depois; nada salvo nele)',
        hash0 == hash1, {'arquivo': caminho, 'sha256_antes': hash0, 'sha256_depois': hash1})
    if saida:
        with open(saida, 'w', encoding='utf-8') as f:
            json.dump(RES, f, ensure_ascii=False, indent=1, default=str)
    falhas = [r for r in RES if r['resultado'] == 'FAIL']
    print(f'\n=== {produto} blocos por lote (ADR-074): {len(RES) - len(falhas)} PASS / {len(falhas)} FAIL ===', flush=True)
    return falhas


if __name__ == '__main__':
    falhas = executar(sys.argv[1], os.path.abspath(sys.argv[2]), sys.argv[3] if len(sys.argv) > 3 else None)
    sys.exit(1 if falhas else 0)
