# -*- coding: utf-8 -*-
"""valida_r7_estat_hema.py -- Tester/QA, ciclo 2026-11-16_reteste07, VALIDAR (OC-0048, complemento).

O valida_r7_adr097.py exercitou o A5 com o lote em analise que a copia abre (522611, so 2025). Aqui, numa COPIA nova
da Hematologia, o lote em analise e trocado pelo caminho do usuario (Painel H3, eventos ligados -> TrocarLoteAnalise) e
conferidos os casos que faltaram:
  (a) lote de 2026 (625311 ou o primeiro lote so de 2026) com Ano De/Ate = 2025: A5 "nao tem resultado em 2025 (tem em
      2026)"; com 2026/2026 o A5 fica vazio; 2025..2026 vazio; B3 2026 > D3 2025 -> aviso "usa so 2026";
  (b) lote com resultado em dois anos (se existir): Ano De/Ate cobrindo so um -> A5 "tambem tem resultado em ...";
  em todos: a grade (n, media) = oraculo proprio da tblCQ_Final por ano (os numeros nao mudam, so o aviso).
Uso: python -u valida_r7_estat_hema.py
Regras: so copia; nada e salvo; fecha so o PID criado pelo qa_final.QA; nenhuma captura de tela.
"""
import datetime as dt
import json
import os
import shutil
import sys
import time

sys.argv = [sys.argv[0], 'Hematologia'] + sys.argv[1:]
AQUI = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, AQUI)
import valida_r7_adr097 as V  # noqa: E402
import qa_final  # noqa: E402


def anos_por_lote(rows):
    out = {}
    for r in rows:
        if V.up(r.get('EQUIPAMENTO')) != V.EQ.upper() or V.up(r.get('PARTICIPA_ESTATISTICA')) != 'SIM':
            continue
        if not V.num(r.get('RESULTADO')) or not V.num(r.get('DATA')):
            continue
        a = (V.EPOCA + dt.timedelta(days=int(r['DATA']))).year
        d = out.setdefault(V.t(r.get('LOTE')), {})
        d[a] = d.get(a, 0) + 1
    return out


def main():
    os.makedirs(V.PASTA, exist_ok=True)
    src = os.path.join(V.C.REPO, 'QC_Hematologia.xlsm')
    dst = os.path.join(V.PASTA, f'Hematologia_estat_{time.strftime("%H%M%S")}.xlsm')
    L = V.Log(os.path.join(V.PASTA, 'valida_r7_estat_Hematologia.json'))
    h0 = V.sha_full(src)
    shutil.copy2(src, dst)
    L['origem'] = {'sha256_antes': h0[:16], 'sha_copia': V.sha_full(dst)[:16]}
    qa = qa_final.QA('Hematologia', dst)
    L['excel_pid'] = qa.ex.pid
    try:
        xl = qa.ex.xl
        rows = V.final(qa)
        apl = anos_por_lote(rows)
        L['anos_por_lote'] = {k: v for k, v in apl.items() if k}
        so26 = '625311' if apl.get('625311', {}).keys() == {2026} else next(
            (k for k, v in apl.items() if set(v) == {2026}), None)
        dois = next((k for k, v in sorted(apl.items(), key=lambda kv: -sum(kv[1].values())) if len(v) >= 2), None)
        L['lotes'] = {'so_2026': so26, 'dois_anos': dois}
        e = qa.ws('Estatística')
        p = qa.ws('Painel')

        def trocar(lote):
            xl.EnableEvents = True
            p.Range('H3').FormulaLocal = "'" + lote
            qa.ex.esperar()
            time.sleep(1.0)
            xl.Calculate()
            qa.ex.esperar()
            return V.t(e.Range('B4').Value)

        def caso(lote, b3, d3):
            xl.EnableEvents = True
            for c, v in (('B3', b3), ('D3', d3)):
                e.Range(c).Value = v
                qa.ex.esperar()
                time.sleep(0.6)
            xl.Calculate()
            qa.ex.esperar()
            g = V.grade(e)
            a_ef = set(range(b3, max(b3, d3) + 1)) if d3 >= b3 else {b3}
            pior, ndif = 0.0, []
            for (an, nv), x in g.items():
                o = V.oraculo(rows, an, nv, lote=lote, anos=a_ef)
                if x['n'] != o['n']:
                    ndif.append(f'{an}|{nv}:{x["n"]}x{o["n"]}')
                elif x['n']:
                    rm = V.rel(x['media'], o.get('media'))
                    if rm is not None:
                        pior = max(pior, rm)
            return {'B3': V.t(e.Range('B3').Text), 'D3': V.t(e.Range('D3').Text), 'B4': V.t(e.Range('B4').Value),
                    'A5': V.t(e.Range('A5').Value), 'A5_cor': e.Range('A5').Font.Color, 'linhas_grade': len(g),
                    'n_total_grade': sum(x['n'] for x in g.values() if V.num(x['n'])),
                    'n_diferentes': ndif[:10], 'max_rel_media': pior}

        res = {}
        if so26:
            res['B4_depois_troca_so26'] = trocar(so26)
            for b3, d3 in ((2025, 2025), (2026, 2026), (2025, 2026), (2026, 2025)):
                res[f'so26_{b3}_{d3}'] = caso(so26, b3, d3)
        if dois:
            anos = sorted(apl[dois])
            res['B4_depois_troca_dois'] = trocar(dois)
            res[f'dois_{anos[0]}_{anos[0]}'] = caso(dois, anos[0], anos[0])
            res[f'dois_{anos[-1]}_{anos[-1]}'] = caso(dois, anos[-1], anos[-1])
            res[f'dois_{anos[0]}_{anos[-1]}'] = caso(dois, anos[0], anos[-1])
        L['casos'] = res
    finally:
        qa.fechar()
        h1 = V.sha_full(src)
        L['fim'] = {'sha256_depois': h1[:16], 'producao_intacta': h1 == h0}
        try:
            os.remove(dst)
        except Exception:                                            # noqa: BLE001
            pass


if __name__ == '__main__':
    main()
