# -*- coding: utf-8 -*-
"""Mede a troca de analito no Painel como o USUARIO a sente, e o custo de cada parte.

Uso: python medir_spinner.py <Bioquimica|Hematologia> <copia.xlsm> [n_trocas] [--visivel] [--json saida.json]

Clique completo (o que o usuario espera): o cronometro comeca ANTES de mudar o valor do spinner
-- a celula ligada (B3) recalcula toda a cadeia do selAnalito em calculo automatico ANTES da macro
-- e termina quando o PainelMudou volta e o Excel para de calcular. --visivel liga a janela e a
renderizacao dos graficos (o custo real na tela do usuario).

Cenarios: frio (1a troca depois de abrir: simula o Workbook_Open e o login), quente (n trocas),
lista C3 (escolha pelo nome, com eventos) e componentes (cada rotina do motor sozinha, em calculo
manual). Nada e salvo.
"""
import json
import os
import statistics
import sys
import time

AQUI = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, AQUI)
sys.path.insert(0, os.path.join(AQUI, '..', '..', 'etapa1_multilote'))
import xlh  # noqa: E402


def esperar_calculo(xl, teto=120):
    t0 = time.time()
    while time.time() - t0 < teto:
        try:
            if xl.CalculationState == 0:
                return
        except Exception:
            pass
        time.sleep(0.005)


def resumo(ts):
    ts = sorted(ts)
    return {'mediana_ms': round(statistics.median(ts) * 1000), 'p90_ms': round(ts[int(0.9 * (len(ts) - 1))] * 1000),
            'max_ms': round(max(ts) * 1000), 'n': len(ts)}


def medir(produto, caminho, n=12, visivel=False):
    ex = xlh.Excel(visivel=visivel)
    out = {'produto': produto, 'visivel': visivel}
    try:
        wb = ex.abrir(caminho)
        try:
            wb.Unprotect('qcini2025')
        except Exception:
            pass
        for ws in wb.Worksheets:
            if ws.Visible != -1:
                ws.Visible = -1
        ex.xl.EnableEvents = True
        ex.xl.Calculation = -4105
        run = lambda m, *a: ex.run("'" + wb.Name + "'!" + m, *a, teto=300)
        # o que o Workbook_Open + login fazem (sem o dialogo)
        run('mApp.Reset')
        run('mDados.AtualizarListasAno')
        run('mLotes.SincronizarLotesAoAbrir')
        run('mSeguranca.ReprotectAll')
        for nome in ('mEstatistica.AquecerMotor',):
            try:
                t0 = time.perf_counter(); run(nome); out['aquecer_s'] = round(time.perf_counter() - t0, 3)
            except Exception:
                out['aquecer_s'] = None
        p = wb.Worksheets('Painel')
        p.Activate()
        sp = p.Spinners('Spinner 1')
        na = int(ex.xl.WorksheetFunction.CountA(wb.Worksheets('Analitos').Range('A4:A43')))
        out['analitos'] = na
        out['spinner_min_max'] = [int(sp.Min), int(sp.Max)]
        v0 = int(sp.Value)

        def clique(v):
            t0 = time.perf_counter()
            sp.Value = v                      # celula ligada: recalculo automatico da cadeia
            esperar_calculo(ex.xl)
            t1 = time.perf_counter()
            run('mEstatistica.PainelMudou')   # OnAction do spinner
            esperar_calculo(ex.xl)
            t2 = time.perf_counter()
            return t1 - t0, t2 - t1, t2 - t0

        prox = lambda k: 1 + ((v0 - 1 + k) % na)
        a, b, tot = clique(prox(1))
        out['frio'] = {'link_ms': round(a * 1000), 'macro_ms': round(b * 1000), 'total_ms': round(tot * 1000),
                       'analito': str(p.Range('C3').Value)}
        print(f'  frio : {tot * 1000:6.0f} ms (link {a * 1000:.0f} + macro {b * 1000:.0f})  {p.Range("C3").Value}', flush=True)
        tl, tm, tt = [], [], []
        for k in range(2, n + 2):
            a, b, tot = clique(prox(k))
            tl.append(a); tm.append(b); tt.append(tot)
            print(f'  quente {k - 1:2d}: {tot * 1000:6.0f} ms (link {a * 1000:.0f} + macro {b * 1000:.0f})  {p.Range("C3").Value}', flush=True)
        out['quente_total'] = resumo(tt)
        out['quente_link'] = resumo(tl)
        out['quente_macro'] = resumo(tm)
        nomes = [str(wb.Worksheets('Analitos').Cells(4 + ((v0 + k) % na), 1).Value) for k in range(4)]
        tlst = []
        for nome in nomes:
            t0 = time.perf_counter()
            p.Range('C3').Value = nome
            esperar_calculo(ex.xl)
            tlst.append(time.perf_counter() - t0)
        out['lista'] = resumo(tlst)
        # componentes, em calculo manual (o motor como roda dentro da operacao)
        comp = {}
        ex.xl.Calculation = -4135
        for k in range(3):
            sp.Value = prox(n + 3 + k)
            for nome in ('mEstatistica.AtualizarCalc', 'mEstatistica.RegistrarEventosWestgard',
                         'mEstatistica.AtualizarPainelEng'):
                t0 = time.perf_counter(); run(nome); comp.setdefault(nome.split('.')[1], []).append(time.perf_counter() - t0)
            t0 = time.perf_counter(); ex.xl.Calculate(); esperar_calculo(ex.xl)
            comp.setdefault('Application.Calculate', []).append(time.perf_counter() - t0)
            t0 = time.perf_counter(); run('mUI.AtualizarEixos'); comp.setdefault('AtualizarEixos', []).append(time.perf_counter() - t0)
        ex.xl.Calculation = -4105
        out['componentes_ms'] = {k: round(statistics.median(v) * 1000) for k, v in comp.items()}
        print(f'== {produto}: frio {out["frio"]["total_ms"]} ms | quente {out["quente_total"]} | lista {out["lista"]}'
              f' | componentes {out["componentes_ms"]}', flush=True)
        wb.Close(False)
        return out
    finally:
        ex.fechar()


if __name__ == '__main__':
    a = sys.argv[1:]
    n = int(a[2]) if len(a) > 2 and a[2].isdigit() else 12
    r = medir(a[0], os.path.abspath(a[1]), n, '--visivel' in a)
    if '--json' in a:
        with open(a[a.index('--json') + 1], 'w', encoding='utf-8') as f:
            json.dump(r, f, ensure_ascii=False, indent=1)
    print(json.dumps(r, ensure_ascii=False))
