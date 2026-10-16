# -*- coding: utf-8 -*-
"""QA do ADR-078 -- as linhas do alvo (Media e -3DP..+3DP) aparecem sem depender das corridas, e trocar de lote
nunca grava na LotesStore. Roda numa COPIA instalada; nada e salvo.

Uso: python qa_linhas_alvo.py <Bioquimica|Hematologia> <copia_instalada.xlsm> [saida.json]

O caso do usuario (15/10/2026): trocou 8974 -> 8973 no Painel com a Media/DP digitada e o grafico nao mostrou nada.
  A01 lote SEM resultado na base, Media/DP digitada no bloco (eventos ligados): store e Calc com os valores; em cada
      grafico as 7 linhas com 2 pontos (x = 0 e 21, y = Media + k*DP); nenhum ponto de resultado; eixo X fixo 0..21;
      eixo Y = Media +- 3,3 DP; status e Q4 "AGUARDANDO 1a CORRIDA"; A39 "sem nenhum resultado na base"; PNG
  A02 o mesmo lote, analito SEM Media/DP: nenhuma linha (Y das linhas = NA), status "SEM MEDIA/DP"
  A03 lote em uso com um periodo sem corrida (2000): status "SEM CORRIDA NO PERIODO" e as 7 linhas continuam
  A04 trocar de lote nao grava: 6 trocas (em uso <-> sem resultado <-> outro), inclusive com a tela do motor
      (Especificacoes E:J) DESATUALIZADA de proposito -- a LotesStore (C:H de todas as linhas) fica identica
  A05 marca TESTE (ADR-076): nivel com valor de TESTE -> status "(alvo TESTE)" e o aviso no A39; a marca e por campo
      (Media/DP): trocar os dois no bloco tira a marca SO daquele nivel
  A06 filtro da rotina: ir a um lote sem resultado no periodo do lote em uso ajusta o filtro com aviso (H4, A39);
      voltar ao lote em uso devolve o filtro de antes (De/Ate, Ano, Periodo, analito, equipamento)
  A07 Hematologia: a troca de lote NUNCA escreve o EQUIPAMENTO_PADRAO da integracao (selEquipamento)
  A08 eventos presos (macro interrompida): a navegacao (mApp.Ir) religa os eventos
  A09 resgate (ADR-078): constante digitada no bloco com os eventos desligados vai para a LotesStore quando a store
      esta vazia (layout antigo e de valores); com outro valor na store: CONFLITO, nao grava
  A10 seguranca do teste: arquivo de entrada intacto
"""
import json
import os
import sys
import tempfile
import traceback

AQUI = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, AQUI)
from qa_final import QA, reg, RES, sha256_arquivo  # noqa: E402

LIM = ['-3s', '-2s', '-1s', 'Média', '+1s', '+2s', '+3s']
KS = [-3, -2, -1, 0, 1, 2, 3]
PONTOS = ('Resultado', 'OK', 'Violação')
R0, PASSO, NAN, O_AN = 2, 54, 40, 4


def num(v):
    return isinstance(v, (int, float)) and not isinstance(v, bool) and not (isinstance(v, int) and v < -2146820000)


def perto(a, b, tol=1e-9):
    return num(a) and num(b) and abs(a - b) <= tol * max(1.0, abs(b))


def txt(v):
    if isinstance(v, float) and v == int(v):
        return str(int(v))
    return str(v if v is not None else '').strip()


def series(ch):
    out = {}
    for i in range(1, ch.SeriesCollection().Count + 1):
        s = ch.SeriesCollection(i)
        try:
            xs, ys = list(s.XValues), list(s.Values)
        except Exception:
            xs, ys = [], []
        pares = [(x, y) for x, y in zip(xs, ys) if num(x) and num(y)]
        out.setdefault(str(s.Name), []).append({'n': len(xs), 'pares': pares})
    return out


def store_valores(q):
    ls = q.ws('LotesStore')
    ult = ls.Cells(ls.Rows.Count, 16).End(-4162).Row
    return [tuple(r) for r in ls.Range(f'A2:H{max(2, ult)}').Value] + [tuple(r) for r in ls.Range(f'P2:P{max(2, ult)}').Value]


def blocos(b):
    out, h = {}, R0
    while txt(b.Cells(h, 1).Value).upper() == 'LOTE':
        out[txt(b.Cells(h, 2).Value)] = h
        h += PASSO
    return out


def executar(produto, caminho, saida):
    hash0 = sha256_arquivo(caminho)
    q = QA(produto, caminho)
    nlv = q.nlv
    png = None
    try:
        q.run('mSeguranca.ReprotectAll')
        p, c, b, e = q.ws('Painel'), q.ws('Calc'), q.ws('Analitos'), q.espec()
        cab = [str(v or '').strip() for v in c.Range('A2:DZ2').Value[0]]
        cel = {}
        for i in range(1, nlv + 1):
            lin1 = [str(c.Cells(1, k).Formula or '') for k in range(1, 131)]
            cm = next(k + 1 for k, f in enumerate(lin1) if f'>0,INDEX(lsMedN{i},' in f.replace(' ', ''))
            cs = next(k + 1 for k, f in enumerate(lin1) if f'>0,INDEX(lsDPN{i},' in f.replace(' ', ''))
            cel[i] = (cm, cs)
        nomes_cat = [txt(v[0]) for v in e.Range('A4:A43').Value]
        rows = q.ler('tblCQ_Final')
        n_lote = {}
        for r in rows:
            n_lote[txt(r['LOTE'])] = n_lote.get(txt(r['LOTE']), 0) + 1
        cad = []
        for v in q.nome('regLoteCol').Value:
            s = txt(v[0])
            if s and s not in cad:
                cad.append(s)
        lote0 = txt(q.nome('loteAtivo').Value)                  # o lote EM USO (Configuracao C20)
        if lote0 not in cad:
            lote0 = txt(q.nome('loteParam').Value) or cad[0]
        eq0 = txt(q.nome('selEquipamento').Value)

        def trocar(lote):
            p.Range('H3').Value = "'" + lote                    # como o usuario: eventos ligados
            q.ex.esperar()
            q.ex.xl.Calculate()

        def ir_analito(an):
            p.Range('B3').Value = nomes_cat.index(an) + 1
            q.run('mEstatistica.PainelMudou', teto=300)
            q.ex.xl.Calculate()

        def md(i):
            return c.Cells(1, cel[i][0]).Value, c.Cells(1, cel[i][1]).Value

        def status():
            return [txt(p.Range(f'J{6 + i}').Value) for i in range(1, nlv + 1)]

        # ------------------------------------------------------------ A01 o caso do usuario
        sem = [l for l in cad if n_lote.get(l, 0) == 0]
        criado = False
        if not sem:
            r = str(q.run('mApp.CadastrarLote', 'QA078', '31/12/2027', teto=300))
            if not r.startswith('OK'):
                raise RuntimeError(f'pre-condicao: cadastrar lote sem resultado: {r}')
            sem, criado = ['QA078'], True
        L = '8973' if '8973' in sem else sem[0]
        an = 'Albumina' if 'Albumina' in nomes_cat else next(x for x in nomes_cat if x)
        an2 = next(x for x in nomes_cat if x and x != an)
        # alvo plausivel: o do lote em uso, se houver; senao 10 +- 0,5 (o teste so confere a geometria)
        alvos = []
        for i in range(1, nlv + 1):
            m0, s0 = q.run('mLotes.MediaDoLote', lote0, an, i), q.run('mLotes.DPDoLote', lote0, an, i)
            alvos.append((round(float(m0), 2), round(float(s0), 3)) if num(m0) and num(s0) and s0 > 0
                         else (10.0 + i, 0.5))
        if L == '8973' and an == 'Albumina' and nlv == 2:
            alvos = [(3.90, 0.17), (2.53, 0.13)]                   # os valores do relato do usuario
        h = blocos(b).get(L)
        if not h:
            q.run('mLotesBlocos.SincronizarBlocos', teto=120)
            h = blocos(b).get(L)
        r_an = h + O_AN + nomes_cat.index(an)
        r_an2 = h + O_AN + nomes_cat.index(an2)
        # analito 2 sem alvo no lote L (A02)
        for j in range(6):
            if txt(b.Cells(r_an2, 5 + j).Value):
                b.Cells(r_an2, 5 + j).Value = None
                q.ex.esperar()
        for i, (m, s) in enumerate(alvos, 1):
            b.Cells(r_an, 5 + 2 * (i - 1)).Value = m                # digitado no bloco, celula a celula
            q.ex.esperar()
            b.Cells(r_an, 6 + 2 * (i - 1)).Value = s
            q.ex.esperar()
        ir_analito(an)
        trocar(L)
        q.run('mUI.AtualizarEixos', teto=120)
        st, q4, a39 = status(), txt(p.Range('Q4').Value), txt(p.Range('A39').Value)
        xmin, xmax = q.nome('xrunmin').Value, q.nome('xrunmax').Value
        ruins, ev1 = [], {'lote': L, 'analito': an, 'alvos': alvos, 'status': st, 'Q4': q4, 'A39': a39[:240],
                          'xrunmin': xmin, 'xrunmax': xmax}
        for i, (m, s) in enumerate(alvos, 1):
            if not (perto(q.run('mLotes.MediaDoLote', L, an, i), m) and perto(q.run('mLotes.DPDoLote', L, an, i), s)):
                ruins.append(f'N{i}: LotesStore sem o valor digitado')
            if not (perto(md(i)[0], m) and perto(md(i)[1], s)):
                ruins.append(f'N{i}: Calc {md(i)} != {(m, s)}')
            ch = p.ChartObjects(i).Chart
            se = series(ch)
            for nm, k in zip(LIM, KS):
                pr = se.get(nm, [{'pares': []}])[0]['pares']
                if len(pr) != 2:
                    ruins.append(f'N{i} {nm}: {len(pr)} par(es)')
                    continue
                if not (perto(pr[0][0], 0) and perto(pr[1][0], 21)):
                    ruins.append(f'N{i} {nm}: x {pr[0][0]}..{pr[1][0]}')
                if not all(perto(y, m + k * s) for _, y in pr):
                    ruins.append(f'N{i} {nm}: y {[y for _, y in pr]} != {m + k * s}')
            for nm in PONTOS:
                if any(x['pares'] for x in se.get(nm, [])):
                    ruins.append(f'N{i} {nm}: tem ponto')
            ax, ay = ch.Axes(1), ch.Axes(2)
            if ax.MinimumScaleIsAuto or not (perto(ax.MinimumScale, 0) and perto(ax.MaximumScale, 21)):
                ruins.append(f'N{i}: eixo X {ax.MinimumScale}..{ax.MaximumScale} auto={ax.MinimumScaleIsAuto}')
            if not (perto(ay.MinimumScale, m - 3.3 * s, 1e-6) and perto(ay.MaximumScale, m + 3.3 * s, 1e-6)):
                ruins.append(f'N{i}: eixo Y {ay.MinimumScale}..{ay.MaximumScale}')
            if not st[i - 1].startswith('AGUARDANDO 1ª CORRIDA'):
                ruins.append(f'N{i}: status {st[i - 1]!r}')
        if not q4.startswith('AGUARDANDO'):
            ruins.append(f'Q4 {q4!r}')
        if 'sem nenhum resultado na base' not in a39:
            ruins.append('A39 sem o aviso')
        try:
            png = os.path.join(os.path.dirname(saida) if saida else tempfile.gettempdir(),
                               f'adr078_{produto}_{L}_N1.png')
            p.ChartObjects(1).Chart.Export(png)
        except Exception as ex_png:
            png = f'erro: {ex_png}'
        ev1['png'] = png
        reg('A01 Caso do usuário: lote sem nenhum resultado, Média/DP digitada no bloco -- a LotesStore e o Calc recebem os '
            'valores e cada gráfico mostra a média e as 6 linhas de limite (2 pontos cada, de 0 a 21; y = Média + k·DP), '
            'sem ponto de resultado; eixo Y = Média ± 3,3 DP; status "AGUARDANDO 1ª CORRIDA"', not ruins,
            dict(ev1, problemas=ruins[:8]))

        # ------------------------------------------------------------ A02 sem alvo: nada inventado
        ir_analito(an2)
        st2 = status()
        se_vazio = []
        for i in range(1, nlv + 1):
            se = series(p.ChartObjects(i).Chart)
            se_vazio.append(sum(len(se.get(nm, [{'pares': []}])[0]['pares']) for nm in LIM))
        reg('A02 Mesmo lote, analito SEM Média/DP: nenhuma linha no gráfico e status "SEM MÉDIA/DP" (nada de outro lote)',
            all(n == 0 for n in se_vazio) and all(s.startswith('SEM MÉDIA/DP') for s in st2)
            and all(md(i) == ('', '') or not num(md(i)[0]) for i in range(1, nlv + 1)),
            {'lote': L, 'analito': an2, 'pares_nas_linhas': se_vazio, 'status': st2})

        # ------------------------------------------------------------ A03 lote em uso, periodo sem corrida
        ir_analito(an)
        trocar(lote0)
        tem_alvo = [num(md(i)[0]) for i in range(1, nlv + 1)]
        g3, g4 = p.Range('G3').Value, p.Range('G4').Value
        p.Range('G3').Value = '01/01/2000'
        q.ex.esperar()
        p.Range('G4').Value = '31/12/2000'
        q.ex.esperar()
        q.ex.xl.Calculate()
        st3 = status()
        pares3 = []
        for i in range(1, nlv + 1):
            se = series(p.ChartObjects(i).Chart)
            pares3.append([len(se.get(nm, [{'pares': []}])[0]['pares']) for nm in LIM])
        ok3 = all((s.startswith('SEM CORRIDA NO PERÍODO') and all(n == 2 for n in pr)) if ta else s.startswith('SEM MÉDIA/DP')
                  for s, pr, ta in zip(st3, pares3, tem_alvo)) and any(tem_alvo)
        reg('A03 Lote em uso com um período sem corrida: status "SEM CORRIDA NO PERÍODO" e as 7 linhas do alvo continuam '
            'no gráfico (2 pontos cada)', ok3, {'lote': lote0, 'analito': an, 'status': st3, 'pares': pares3})
        p.Range('G3').Value = g3
        q.ex.esperar()
        p.Range('G4').Value = g4
        q.ex.esperar()

        # ------------------------------------------------------------ A04 trocar nao grava
        outro = next((l for l in cad if l not in (lote0, L) and n_lote.get(l, 0) > 0), None)
        antes = store_valores(q)
        seq = [L, lote0, outro, L, outro, lote0] if outro else [L, lote0, L, lote0]
        for k, l in enumerate(seq):
            if k == 2:
                # tela do motor DESATUALIZADA de proposito (eventos desligados): antes, a troca gravava ESTA tela no lote
                q.ex.xl.EnableEvents = False
                lin = 4 + nomes_cat.index(an)
                e.Unprotect('qcini2025')
                e.Range(f'E{lin}').Value = 123456.0
                q.ex.xl.EnableEvents = True
            trocar(l)
        depois = store_valores(q)
        dif = [(a, d) for a, d in zip(antes, depois) if a != d]
        reg('A04 Trocar de lote não grava na LotesStore: 6 trocas (em uso, sem resultado, outro), uma delas com a tela do '
            'motor desatualizada de propósito -- todas as linhas da LotesStore idênticas',
            not dif and len(antes) == len(depois), {'sequencia': seq, 'linhas': len(antes), 'diferencas': dif[:4]})

        # ------------------------------------------------------------ A05 marca TESTE (bit 1 = Media, bit 2 = DP)
        ls = q.ws('LotesStore')
        ult = ls.Cells(ls.Rows.Count, 16).End(-4162).Row

        def bit(v):
            return int(v) if num(v) else 0
        marcados = [(txt(r[0]), txt(r[14]), [bit(r[23 + k]) for k in range(3)]) for r in ls.Range(f'A2:Z{ult}').Value
                    if any(bit(r[23 + k]) > 0 for k in range(nlv))]
        if marcados:
            lt, at, fl = next(((a_, b_, f) for a_, b_, f in marcados if a_ in cad and b_ in nomes_cat and a_ != lote0),
                              marcados[0])
            nv = next(k for k in range(nlv) if fl[k] > 0) + 1
            ir_analito(at)
            trocar(lt)
            if txt(q.nome('selAnalito').Value) != at:          # o ajuste do ADR-077 pode ter mudado o analito
                ir_analito(at)
            st5, a395 = status(), txt(p.Range('A39').Value)
            ok_a = st5[nv - 1].endswith('(alvo TESTE)')
            ok_b = 'ALVO DE TESTE' in a395
            h5 = blocos(b).get(lt)
            r5 = h5 + O_AN + nomes_cat.index(at)
            m_old, s_old = b.Cells(r5, 5 + 2 * (nv - 1)).Value, b.Cells(r5, 6 + 2 * (nv - 1)).Value
            b.Cells(r5, 5 + 2 * (nv - 1)).Value = round(float(m_old) * 1.01, 4) if num(m_old) else 1.0
            q.ex.esperar()
            so_media = bool(q.run('mLotes.AlvoDeTeste', lt, at, nv))
            ok_c = so_media == bool(fl[nv - 1] & 2)            # so a Media trocada: continua TESTE se o DP era de teste
            b.Cells(r5, 6 + 2 * (nv - 1)).Value = round(float(s_old) * 1.01, 5) if num(s_old) and s_old > 0 else 0.5
            q.ex.esperar()
            q.ex.xl.Calculate()
            ok_d = not q.run('mLotes.AlvoDeTeste', lt, at, nv)
            outros = [k for k in range(1, nlv + 1) if k != nv]
            ok_e = all(bool(q.run('mLotes.AlvoDeTeste', lt, at, k)) == (fl[k - 1] > 0) for k in outros)
            st5b = status()
            trocar(lote0)                                       # volta ao lote em uso: o filtro da rotina volta
            reg('A05 Marca TESTE (ADR-076): o nível com valor de TESTE aparece "(alvo TESTE)" no Painel, com o aviso no A39; '
                'trocar só a Média mantém a marca se o DP era de teste; trocar a Média e o DP tira a marca só daquele nível',
                ok_a and ok_b and ok_c and ok_d and ok_e and not st5b[nv - 1].endswith('(alvo TESTE)'),
                {'lote': lt, 'analito': at, 'nivel': nv, 'bits': fl, 'status_antes': st5, 'status_depois': st5b,
                 'A39_tem_aviso': ok_b, 'teste_depois_so_da_media': so_media, 'outros_niveis_mantidos': ok_e})
        else:
            reg('A05 Marca TESTE (ADR-076): nenhum nível marcado nesta cópia (sem valores de TESTE) -- não aplicável', True,
                {'marcados': 0})

        # ------------------------------------------------------------ A06 filtro da rotina
        ir_analito(an)
        trocar(lote0)
        # o filtro da rotina, escolhido pelo usuario: o ano do ultimo resultado do lote em uso neste analito
        sim = [r for r in rows if str(r.get('PARTICIPA_ESTATISTICA', '')).upper() == 'SIM']
        n_sim = {}
        for r in sim:
            n_sim[txt(r['LOTE'])] = n_sim.get(txt(r['LOTE']), 0) + 1
        datas0 = [r['DATA_HORA'] for r in sim if txt(r['LOTE']) == lote0 and txt(r.get('ANALITO')).upper() == an.upper()
                  and hasattr(r.get('DATA_HORA'), 'year')]
        if datas0:
            p.Range('K3').Value = max(datas0).year
            q.ex.esperar()
            p.Range('K4').Value = 'Ano inteiro'
            q.ex.esperar()
            q.ex.xl.Calculate()
        q.run('mLotes.EsquecerFiltroRotina', teto=60)
        F0 = str(q.run('mLotes.FiltroAtualTexto', teto=60))
        eq_rot = txt(q.nome('selEquipamento').Value)
        de, ate = p.Range('G3').Value, p.Range('G4').Value
        cand = None
        if de or ate:
            for l in cad:
                if l in (lote0,) or n_sim.get(l, 0) == 0:
                    continue
                dentro = False
                for r in sim:
                    if txt(r['LOTE']) != l:
                        continue
                    if txt(r.get('ANALITO')).upper() != an.upper():
                        continue
                    d = r.get('DATA_HORA')
                    if hasattr(d, 'date'):
                        dd = d.replace(tzinfo=None)
                        if (not de or dd >= de.replace(tzinfo=None)) and (not ate or dd <= ate.replace(tzinfo=None, hour=23, minute=59)):
                            dentro = True
                            break
                if not dentro:
                    cand = l
                    break
        if cand:
            trocar(cand)
            aviso, h4 = str(q.run('mLotes.AvisoPainel', teto=60)), txt(p.Range('H4').Value)
            F1 = str(q.run('mLotes.FiltroAtualTexto', teto=60))
            trocar(lote0)
            F2 = str(q.run('mLotes.FiltroAtualTexto', teto=60))
            aviso2 = str(q.run('mLotes.AvisoPainel', teto=60))
            reg('A06 Filtro da rotina: o lote sem resultado no período do lote em uso ajusta o filtro, com aviso no Painel '
                '(H4 e A39); voltar ao lote em uso devolve o filtro de antes (analito, De/Até, Ano, Período, equipamento)',
                bool(aviso) and 'aviso' in h4 and F1 != F0 and F2 == F0 and 'restaurado' in aviso2,
                {'lote_em_uso': lote0, 'visitado': cand, 'filtro_antes': F0, 'filtro_no_visitado': F1,
                 'filtro_de_volta': F2, 'aviso': aviso[:160], 'H4': h4, 'aviso_de_volta': aviso2})
        else:
            reg('A06 Filtro da rotina: nenhum lote do cadastro fora do período em tela -- não aplicável', True,
                {'filtro': F0})

        # ------------------------------------------------------------ A07 Hematologia: EQUIPAMENTO_PADRAO intocado
        eq1 = txt(q.nome('selEquipamento').Value)
        hoje_sheet = q.nome('selEquipamento').Worksheet.Name
        reg('A07 A troca de lote nunca escreve o equipamento da integração (Hematologia: Cfg_Integracao); na Bioquímica o '
            'seletor do Painel (L4) pode ser ajustado e é devolvido ao voltar ao lote em uso',
            (eq1 == eq0) if hoje_sheet != 'Painel' else (eq1 == eq_rot),
            {'aba_do_selEquipamento': hoje_sheet, 'antes': eq0, 'rotina_A06': eq_rot, 'depois': eq1})

        # ------------------------------------------------------------ A08 eventos presos
        # o usuario (Interactive = True; a automacao do QA roda com False e a guarda nao age nela)
        q.ex.xl.EnableEvents = False
        q.ex.xl.Interactive = True
        try:
            q.run('mApp.Ir', 'Painel', teto=60)
            ok8 = bool(q.ex.xl.EnableEvents)
        finally:
            q.ex.xl.Interactive = False
            q.ex.xl.EnableEvents = True
        reg('A08 Eventos presos fora de operação (macro interrompida): a navegação religa os eventos -- a digitação de '
            'Média/DP volta a chegar à LotesStore', ok8, {'eventos_depois_de_Ir': ok8})

        # ------------------------------------------------------------ A09 resgate de digitacao
        def cab(r):
            p_ = r.split('\n')[0].split('|')
            return [int(x) for x in p_[1:4]] if len(p_) >= 4 else None
        h9 = blocos(b).get(lote0)
        r9 = h9 + O_AN + nomes_cat.index(an2)
        evitar = {an2, nomes_cat[nomes_cat.index(an2) + 1] if nomes_cat.index(an2) + 1 < len(nomes_cat) else ''}
        anC = next((x for x in nomes_cat if x and x not in evitar and num(q.run('mLotes.ValorDoLote', lote0, x, 1))), None)
        rA = h9 + O_AN + nomes_cat.index(anC) if anC else None
        b.Cells(r9, 5).Value = None                                 # store vazia neste campo (eventos ligados)
        q.ex.esperar()
        vazio9 = q.run('mLotes.ValorDoLote', lote0, an2, 1)
        conf_store = q.run('mLotes.ValorDoLote', lote0, anC, 1) if anC else None
        hL = blocos(b).get(L)
        rL = hL + O_AN + nomes_cat.index(an2)
        vazioL = q.run('mLotes.ValorDoLote', L, an2, 1)
        q.ex.xl.EnableEvents = False
        b.Unprotect('qcini2025')
        b.Cells(r9 + 1, 5).Formula = '=IF(1=1,"","")'               # (a) o layout antigo: E:J com formula
        b.Cells(r9, 5).Value = 7.7777                               #     digitado com os eventos desligados
        if anC:
            b.Cells(rA, 5).Value = conf_store + 1.0                 # (b) outro numero onde a store ja tem valor
        b.Cells(rL, 5).Value = 5.5                                  # (c) bloco so de valores (ADR-076)
        q.ex.xl.EnableEvents = True
        r = str(q.run('mLotesBlocos.ResgatarDigitados', teto=300))
        d9 = q.run('mLotes.ValorDoLote', lote0, an2, 1)
        dA = q.run('mLotes.ValorDoLote', lote0, anC, 1) if anC else None
        dL = q.run('mLotes.ValorDoLote', L, an2, 1)
        r2 = str(q.run('mLotesBlocos.ResgatarDigitados', teto=300))
        c1, c2 = cab(r), cab(r2)
        reg('A09 Resgate: Média digitada no bloco com os eventos desligados -- store vazia: vai para a LotesStore (layout '
            'antigo com fórmula e layout de valores); store com outro valor: CONFLITO, não grava e aparece no relatório; '
            'a 2ª execução não recupera mais nada',
            r.startswith('OK|') and c1 is not None and c1[0] >= 2 and vazio9 in ('', None)
            and vazioL in ('', None) and perto(d9, 7.7777) and perto(dL, 5.5)
            and ((c1[1] >= 1 and perto(dA, conf_store) and 'CONFLITO' in r) if anC else True)
            and c2 is not None and c2[0] == 0,
            {'lote_em_uso': lote0, 'lote_sem_resultado': L, 'analito': an2, 'store_depois': {'formula': d9, 'valores': dL},
             'conflito': {'analito': anC, 'store_antes_depois': (conf_store, dA)}, 'retorno': r[:400],
             'segunda': r2[:120]})
    except Exception:
        traceback.print_exc()
        reg('AXX execução do QA sem exceção', False, {'erro': traceback.format_exc()[-800:]})
    finally:
        q.fechar(salvar=False)
    hash1 = sha256_arquivo(caminho)
    reg('A10 Segurança do teste: arquivo de entrada intacto (SHA-256 igual antes e depois; nada salvo nele)',
        hash0 == hash1, {'arquivo': caminho, 'sha256_antes': hash0, 'sha256_depois': hash1})
    if saida:
        with open(saida, 'w', encoding='utf-8') as f:
            json.dump({'resultados': RES}, f, ensure_ascii=False, indent=1, default=str)
    falhas = [r for r in RES if r['resultado'] == 'FAIL']
    print(f'\n=== {produto} linhas do alvo (ADR-078): {len(RES) - len(falhas)} PASS / {len(falhas)} FAIL ===', flush=True)
    return falhas


if __name__ == '__main__':
    falhas = executar(sys.argv[1], os.path.abspath(sys.argv[2]), sys.argv[3] if len(sys.argv) > 3 else None)
    sys.exit(1 if falhas else 0)
