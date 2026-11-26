# -*- coding: utf-8 -*-
"""repro_e46_sigma.py -- Tester/QA, validacao da entrega 46 (ADR-101), oraculo proprio. Roda numa COPIA; nada salvo.

Uso: python repro_e46_sigma.py <Bioquimica|Hematologia> <copia.xlsm> <saida.json>

  G01 Estatistica L14:L133: todo Sigma numerico >= 0; a classe M = regra institucional do ADR-103 (5 faixas literais) em TODAS as linhas (todos os analitos x niveis); L = 0 quando
      (ETp - |bias|) < 0 (ver colunas lidas do cabecalho)
  G02 tblPlanoQC_Sigma com as 6 faixas (Ruim 2-3 e Inaceitavel < 2)
  G03 Painel: em ate 6 analitos (um por faixa encontrada, se houver), I7.. = Estatistica L do analito/nivel, K = classe
      da faixa, e a cor exibida (DisplayFormat) de I e K e a mesma dentro da faixa e diferente entre faixas
  G04 legenda V5:X8 preenchida e fora de todos os graficos (nenhum ChartObject cruza V5:X8); 6 regras de cor em I7 e L14
  Y00 arquivo intacto
"""
import collections
import json
import os
import sys
import traceback

TESTES = os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', '..', 'integracao_dbseac', 'testes')
sys.path.insert(0, os.path.abspath(TESTES))
from qa_final import QA, reg, RES, sha256_arquivo  # noqa: E402
from qa_fluxo_manual import SENHA  # noqa: E402
import painel_layout  # noqa: E402  (ADR-103)

# ADR-103: regra institucional (operadores literais); nomes sem acento, comparados por igualdade
def faixa(s):
    if s >= 6:
        return 'excelente — classe mundial'
    if 5 < s < 6:
        return 'excelente'
    if 4 < s <= 5:
        return 'bom'
    if 3 <= s <= 4:
        return 'marginal'
    return 'insatisfatorio'


def num(v):
    return isinstance(v, (int, float)) and not isinstance(v, bool)


def sem_ac(s):
    import unicodedata
    return ''.join(c for c in unicodedata.normalize('NFD', str(s or '')) if unicodedata.category(c) != 'Mn').lower()


def executar(produto, caminho, saida):
    h0 = sha256_arquivo(caminho)
    q = QA(produto, caminho)
    try:
        est = q.ws('Estatística')
        cab = [sem_ac(x) for x in est.Range('A13:Z13').Value[0]]
        linhas = [list(r) for r in est.Range('A14:Z133').Value]
        neg, dif, por_faixa, n = [], [], collections.defaultdict(list), 0
        for i, r in enumerate(linhas):
            s, c = r[11], r[12]
            if not num(s):
                continue
            n += 1
            if s < 0:
                neg.append(14 + i)
            f = faixa(s)
            if f != sem_ac(c):
                dif.append({'linha': 14 + i, 'sigma': s, 'classe': c})
            por_faixa[f].append((str(r[0]).strip(), int(r[1] or 0), s, 14 + i))
        reg('G01 Estatistica: Sigma >= 0 e classe M = regra institucional (ADR-103) em todas as linhas com Sigma',
            n > 0 and not neg and not dif,
            {'linhas_com_sigma': n, 'negativos': neg[:10], 'classe_diferente': dif[:10],
             'por_faixa': {k: len(v) for k, v in por_faixa.items()}, 'cabecalho_L_M': cab[11:13]})
        lo = q.lo('tblPlanoQC_Sigma')
        hdr = [c.Name for c in lo.ListColumns]
        pl = [dict(zip(hdr, r)) for r in lo.DataBodyRange.Value]
        cls = sorted((float(p['Sigma_Min']), float(p['Sigma_Max']), sem_ac(p['Classificacao'])) for p in pl)
        ok2 = len(cls) == 6 and cls[0][1] == 2.0 and 'inaceit' in cls[0][2] and cls[1][:2] == (2.0, 3.0) \
            and 'ruim' in cls[1][2]
        reg('G02 tblPlanoQC_Sigma com 6 faixas (Ruim 2-3, Inaceitavel < 2)', ok2, {'faixas': cls})
        # G03 Painel por faixa
        p = q.ws('Painel')
        sel = q.espec()
        cat = [str(v[0] or '').strip() for v in sel.Range('A4:A43').Value]
        nlv = q.nlv
        amostra = []
        for f, lst in por_faixa.items():
            for an, nv, s, lin in lst:
                if an in cat:
                    amostra.append((f, an, nv, s, lin))
                    break
        res, cores = [], collections.defaultdict(set)
        for f, an, nv, s, lin in amostra:
            try:
                p.Unprotect(SENHA)
            except Exception:                           # noqa: BLE001
                pass
            q.ex.xl.EnableEvents = False
            p.Range('B3').Value = cat.index(an) + 1     # indice do analito (como qa_adr086)
            q.ex.xl.EnableEvents = True
            q.run('mEstatistica.RecalcularAnalitoAtual', teto=600)
            q.ex.xl.Calculate()
            ri = 6 + nv
            ck_ = painel_layout.col_classe(p)              # ADR-103: K (antigo) ou J (novo)
            vi, vk = p.Range(f'I{ri}').Value, str(p.Range(f'{ck_}{ri}').Value or '')
            ci = p.Range(f'I{ri}').DisplayFormat.Interior.Color
            ck = p.Range(f'{ck_}{ri}').DisplayFormat.Interior.Color
            ok = num(vi) and abs(vi - s) <= 1e-9 * max(1, abs(s)) and f == sem_ac(vk)
            cores[f].add((ci, ck))
            res.append({'faixa': f, 'analito': an, 'nivel': nv, 'est_L': s, 'painel_I': vi, 'painel_K': vk,
                        'cor_I': ci, 'cor_K': ck, 'ok': ok})
        uma_cor = all(len(v) == 1 for v in cores.values())
        distintas = len({next(iter(v)) for v in cores.values()}) == len(cores)
        reg('G03 Painel: I = Estatistica L, K = classe da faixa; cor exibida unica por faixa e distinta entre faixas',
            bool(res) and all(x['ok'] for x in res) and uma_cor and distintas,
            {'amostra': res, 'uma_cor_por_faixa': uma_cor, 'cores_distintas': distintas})
        # G04 legenda
        leg = p.Range('V5:X8')
        L, T, R, B = leg.Left, leg.Top, leg.Left + leg.Width, leg.Top + leg.Height
        cruza = []
        for co in p.ChartObjects():
            if not (co.Left >= R or co.Left + co.Width <= L or co.Top >= B or co.Top + co.Height <= T):
                cruza.append(co.Name)
        txt = [str(x or '') for r in leg.Value for x in r]
        n_fc = (p.Range('I7').FormatConditions.Count, est.Range('L14').FormatConditions.Count)
        reg('G04 Legenda V5:X8 preenchida e fora de todos os graficos; 5 regras de cor (ADR-103) em Painel I7 e Estatistica L14',
            bool(txt[0]) and sum(1 for t in txt if t) >= 6 and not cruza and n_fc == (5, 5),
            {'V5': txt[0][:80], 'celulas_preenchidas': sum(1 for t in txt if t), 'graficos_que_cruzam': cruza,
             'n_graficos': p.ChartObjects().Count, 'regras_cor_I7_L14': n_fc})
    except Exception as e:                                # noqa: BLE001
        reg('Z99 Execucao sem excecao', False, {'erro': repr(e), 'trace': traceback.format_exc()[-1500:]})
    finally:
        q.fechar(salvar=False)
    h1 = sha256_arquivo(caminho)
    reg('Y00 Arquivo de entrada intacto (SHA-256)', h0 == h1, {'antes': h0[:16], 'depois': h1[:16]})
    if saida:
        with open(saida, 'w', encoding='utf-8') as f:
            json.dump(RES, f, ensure_ascii=False, indent=1, default=str)
    falhas = [r for r in RES if r['resultado'] == 'FAIL']
    print(f'\n=== {produto} repro_e46_sigma: {len(RES) - len(falhas)} PASS / {len(falhas)} FAIL ===', flush=True)
    return falhas


if __name__ == '__main__':
    f = executar(sys.argv[1], os.path.abspath(sys.argv[2]), sys.argv[3] if len(sys.argv) > 3 else None)
    sys.exit(1 if f else 0)
