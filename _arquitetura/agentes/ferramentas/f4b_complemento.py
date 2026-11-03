# -*- coding: utf-8 -*-
"""F4 do Usuario Simulado -- complemento do f4_periodos_lotes.py (mesma copia, outra sessao do Excel).

Uso (um Excel por vez):
    python -u f4b_complemento.py <copia.xlsm> <cq.pkl> <pasta_saida>

O que faz (pelo caminho do usuario: celulas de filtro do Painel com os eventos ligados):
  A  periodos longos que o roteiro principal nao alcancou (os rotulos reais do K4 sao "Ano inteiro" e
     "Tudo (sem filtro)"): Glicose / DIMENSION 1 / lote 8974 (mais de 180 corridas). Compara o Painel com TODOS os
     dados brutos do periodo e com os dados so da janela de 180 corridas, e guarda o aviso (A39/H4).
  B  H3 vazio (ADR-086) num periodo com dois lotes de Glicose no DIMENSION 1 (set-out/2026: 8977 e 5659); e o
     lote 5659 sozinho em H3 (o lote aparece na lista H3 e tem linhas MATRIZ=SORO mapeadas para Glicose).
  V  imagens do Painel e da Estatistica exportadas pelo PROPRIO Excel da copia (ExportAsFixedFormat -> PNG, como o
     testes/auditoria_abas.py) em tres estados, mais a auditoria automatica de formas/textos (auditar_aba).
Nada e salvo na copia. Calculo independente: o mesmo do f4_periodos_lotes (statistics sobre a tblCQ_Final bruta).
Autor: AGENTE_USUARIO_SIMULADO, ciclo 2026-10-31_ciclo01 (2o passe).
"""
import collections
import datetime as dt
import json
import os
import pickle
import sys
import traceback

AQUI = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, AQUI)
import f4_periodos_lotes as f4  # noqa: E402
from f4_periodos_lotes import QA, Bruto, Tela, cmp_painel, cmp_eng_escopo, janela, stats, j, reg, OUT, log  # noqa: E402
import auditoria_abas as aa  # noqa: E402  (o f4 ja pos testes/ no sys.path)
from qa_etl_recebimento import SENHA  # noqa: E402


def executar(copia, pkl, pasta):
    os.makedirs(pasta, exist_ok=True)
    d = pickle.load(open(pkl, 'rb'))
    b = Bruto(d['rows'])
    q = QA('Bioquimica', copia)
    OUT['excel_pid'] = q.ex.pid
    AN, EQ = 'Glicose', 'DIMENSION 1'
    try:
        t = Tela(q, b)
        q.run('mSeguranca.ReprotectAll', teto=120)
        t.muda('C3', AN)
        t.muda('L4', EQ)
        # ------------------------------------------------ A: Ano inteiro / Tudo, lote 8974
        for (lab, d0, d1) in [('Ano inteiro', dt.date(2026, 1, 1), dt.date(2026, 12, 31)),
                              ('Tudo (sem filtro)', None, None),
                              ('S1', dt.date(2026, 1, 1), dt.date(2026, 6, 30))]:
            try:
                t.muda('H3', "'8974")
                t.muda('K3', 2026)
                dtc = t.muda('K4', lab)
                est = t.estado()
                dd0, dd1 = (d0 or dt.date(1900, 1, 1)), (d1 or dt.date(2100, 1, 1))
                eixo = b.corridas(AN, EQ, '8974')
                jan, cort = janela(eixo, d1)
                win = []
                for t_ in range(2):
                    vals = [c['v'][t_ + 1][0] for c in jan if (t_ + 1) in c['v'] and dd0 <= c['dh'].date() <= dd1]
                    win.append(stats(vals))
                todas = [stats([float(e['val']) for e in b.sel(AN, EQ, '8974', t_ + 1, dd0, dd1)]) for t_ in range(2)]
                cmp_painel(t, est, AN, EQ, '8974', dd0, dd1, f'A-{lab[:4]}',
                           f'Período "{lab}" 2026 — Glicose DIM 1 lote 8974 (mais de 180 corridas): Painel = todos os dados brutos do período',
                           {'tempo_s': dtc, 'K4': est['K4'], 'A39': est['A39'], 'corridas_elegiveis_antes_da_janela': cort,
                            'estatistica_so_da_janela_180': win, 'estatistica_todas': todas,
                            'painel_igual_janela': [abs((est['engPainel'][k][2] or 0) - (win[k]['media'] or 0)) < 1e-9 and est['engPainel'][k][1] == win[k]['n'] for k in range(2)]})
            except Exception as ex:
                OUT['erros'].append(f'A {lab}: {ex!r} {traceback.format_exc()[-800:]}')
        # ------------------------------------------------ B: H3 vazio com 8977 + 5659 (set-out/2026)
        try:
            t.muda('K3', 2026)
            t.muda('G3', f4.serie(dt.date(2026, 9, 1)))
            t.muda('G4', f4.serie(dt.date(2026, 10, 31)))
            d0, d1 = dt.date(2026, 9, 1), dt.date(2026, 10, 31)
            dtc = t.muda('H3', None)
            est = t.estado()
            eng, fora, faltam = cmp_eng_escopo(t, AN, EQ, None, d0, d1, todos=True)
            cq = {e['id']: e for e in b.elig}
            mat = {f4.txt(r['ID_REGISTRO']): f4.txt(r.get('MATRIZ')) for r in d['rows']}
            pts = collections.Counter()
            exemplos = collections.defaultdict(list)
            for x in eng:
                if x['fil'] != 1:
                    continue
                for t_ in range(2):
                    i = x['id'][t_]
                    if i in cq:
                        k = (cq[i]['lote'], t_ + 1, mat.get(i))
                        pts[k] += 1
                        if len(exemplos[k]) < 3:
                            exemplos[k].append({'id': i, 'valor': x['val'][t_], 'alvo': x['alvo'][t_], 'flags': x['flags'][t_]})
            ind = [len(b.sel(AN, EQ, None, t_ + 1, d0, d1)) for t_ in range(2)]
            reg('B-TODOS-SETOUT', 'H3 vazio, Glicose DIM 1, 01/09–31/10/2026: o motor junta os lotes do período (8977 e 5659) '
                'com o alvo de cada um; n = todos os lotes',
                'APROVADO' if (not fora and not faltam and [est['engPainel'][k][1] for k in range(2)] == ind) else 'REPROVADO',
                {'H4': est['H4'], 'A39': est['A39'], 'n_qc': [est['engPainel'][k][1] for k in range(2)], 'n_independente': ind,
                 'pontos_por_lote_nivel_matriz': {f'{k[0]}|N{k[1]}|{k[2]}': v for k, v in pts.items()},
                 'exemplos': {f'{k[0]}|N{k[1]}|{k[2]}': v for k, v in exemplos.items()},
                 'linha_painel': est['B7:T8'], 'fora': fora[:5], 'faltam': faltam[:5], 'tempo_s': dtc})
            OUT['todos_setout_estado'] = j(est)
            try:
                t.p.Unprotect(SENHA)              # passo tecnico: PageSetup da exportacao (nada e gravado)
                OUT['img_todos'] = aa.exportar_png(t.p, pasta, 45, 30)
                os.replace(os.path.join(pasta, 'Painel.png'), os.path.join(pasta, 'Painel_todos_setout.png'))
            except Exception as ex:
                OUT['img_todos'] = repr(ex)
            # lote 5659 sozinho
            dtc = t.muda('H3', "'5659")
            est = t.estado()
            m1, s1 = t.alvo('5659', AN, 1)
            m2, s2 = t.alvo('5659', AN, 2)
            g0, g1 = f4.dia(est['G3']), f4.dia(est['G4'])
            cmp_painel(t, est, AN, EQ, '5659', g0 or d0, g1 or d1, 'B-5659',
                       'Lote 5659 em H3 (lote de controle de URINA na Configuração? linhas SORO mapeadas para Glicose): '
                       'Painel = dados brutos do lote',
                       {'tempo_s': dtc, 'alvo_cadastro': [[m1, s1], [m2, s2]], 'H4': est['H4'], 'A39': est['A39'],
                        'C3': est['C3'], 'L4': est['L4'], 'K4': est['K4'], 'status': [est['B7:T8'][0][9], est['B7:T8'][1][9]]})
            OUT['estado_5659'] = j(est)
            # lotes do cadastro (Configuracao) -- onde o 5659 esta e com que descricao
            try:
                ws = q.ws('Configuração')
                ur = ws.UsedRange.Value
                OUT['config_5659'] = [list(map(f4.txt, l))[:12] for l in ur if any(f4.txt(c) == '5659' for c in l)]
            except Exception as ex:
                OUT['config_5659'] = repr(ex)
        except Exception as ex:
            OUT['erros'].append(f'B: {ex!r} {traceback.format_exc()[-1200:]}')
        # ------------------------------------------------ F: formulas do Painel e celulas de filtro (o que o usuario pode mexer)
        try:
            OUT['painel_formulas_7'] = j({t.p.Cells(7, c_).Address.replace('$', ''): str(t.p.Cells(7, c_).Formula)[:400]
                                          for c_ in range(1, 21)})
            cel_info = {}
            for ws_, cels in ((t.p, ('B3', 'C3', 'G3', 'G4', 'H3', 'I3', 'K3', 'K4', 'L4', 'M4')),
                              (t.est, ('B4', 'D4', 'F4', 'H4', 'B5', 'D5', 'N4', 'P4'))):
                for c_ in cels:
                    r_ = ws_.Range(c_)
                    info = {'formula': str(r_.Formula)[:200], 'valor': str(r_.Value)[:60], 'bloqueada': bool(r_.Locked)}
                    try:
                        info['validacao'] = [r_.Validation.Type, str(r_.Validation.Formula1)[:200]]
                    except Exception:
                        info['validacao'] = None
                    cel_info[f'{ws_.Name}!{c_}'] = info
            OUT['celulas_filtro'] = j(cel_info)
        except Exception as ex:
            OUT['erros'].append(f'F: {ex!r} {traceback.format_exc()[-800:]}')
        # ------------------------------------------------ E2: Estatistica pelo caminho do usuario (so B4 Visao e F4)
        try:
            t.muda('H3', "'8977")
            t.muda('K3', 2026)
            t.muda('K4', 'Set')
            ws = t.est
            prot_est = bool(ws.ProtectContents)
            res = []
            for (vis, parte) in (('MENSAL', 9), ('TRIMESTRAL', 3), ('PERSONALIZADO', None)):
                erro = None
                try:
                    t.muda('B4', vis, aba=ws)
                    if parte is not None:
                        t.muda('F4', parte, aba=ws)
                except Exception as ex:
                    erro = repr(ex)[:200]
                q.ex.xl.Calculate()
                linhas, filtro = t.estat_linhas()
                res.append({'visao': vis, 'parte': parte, 'erro_escrita': erro, 'filtro': filtro,
                            'glicose': [linhas.get((AN, 1)), linhas.get((AN, 2))]})
            OUT['estatistica_caminho_usuario'] = j({'aba_protegida': prot_est, 'casos': res})
        except Exception as ex:
            OUT['erros'].append(f'E2: {ex!r} {traceback.format_exc()[-800:]}')
        # ------------------------------------------------ V: imagens e auditoria visual
        try:
            t.muda('H3', "'8977")
            t.muda('K3', 2026)
            t.muda('K4', 'Out')
            t.p.Unprotect(SENHA)
            OUT['img_out_8977'] = aa.exportar_png(t.p, pasta, 45, 30)
            OUT['achados_painel'] = aa.auditar_aba(t.p)
            os.replace(os.path.join(pasta, 'Painel.png'), os.path.join(pasta, 'Painel_out_8977.png'))
            t.muda('K4', 'T3')
            OUT['img_t3_8977'] = aa.exportar_png(t.p, pasta, 45, 30)
            os.replace(os.path.join(pasta, 'Painel.png'), os.path.join(pasta, 'Painel_t3_8977.png'))
            ws = t.est
            ws.Unprotect(SENHA)
            OUT['achados_estatistica'] = aa.auditar_aba(ws)
            OUT['img_estatistica'] = aa.exportar_png(ws, pasta, 60, 30)
        except Exception as ex:
            OUT['erros'].append(f'V: {ex!r} {traceback.format_exc()[-1200:]}')
    finally:
        q.fechar(False)
        log('Excel encerrado (PID', OUT.get('excel_pid'), ')')
        with open(os.path.join(pasta, 'f4b.json'), 'w', encoding='utf-8') as f:
            json.dump(j(OUT), f, ensure_ascii=False, indent=1)


if __name__ == '__main__':
    executar(sys.argv[1], sys.argv[2], sys.argv[3])
