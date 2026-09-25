# -*- coding: utf-8 -*-
"""prova_v2.py -- a camada de desempenho (ADR-050) muda o TEMPO, nao o RESULTADO.

Uso: python prova_v2.py <arquivo_etapa1_grande.xlsm>   (copia; nada e salvo)

1. fotografa, com o codigo da Etapa 1, o que o usuario ve em varias
   combinacoes (lote x analito x periodo): Calc inteiro, Painel, faixa O3,
   Eng_Saida, Eventos_Westgard, Estatistica e a lista da Liberacao;
2. instala a v2 (VBA + Calc lendo o motor + Liberacao por valor);
3. fotografa as MESMAS combinacoes e compara celula a celula;
4. mede: clique do spinner, troca de lote em analise, troca de lote em uso,
   importacao de uma corrida.
"""
import datetime as dt
import os
import shutil
import sys
import time

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), '..'))
import xlh  # noqa: E402
import v2   # noqa: E402

PROV = 'IFERROR(INDEX(Analitos!$AR$4:$AR$43,MATCH($A14,Analitos!$A$4:$A$43,0)),"CAP")'
TMP = os.path.join(os.environ['TMP'], 'qcw', 'perf')
CRLF, LF = chr(13) + chr(10), chr(10)
os.makedirs(TMP, exist_ok=True)


def num(v):
    return isinstance(v, (int, float)) and not isinstance(v, bool)


def igual(a, b):
    if num(a) and num(b):
        return abs(a - b) <= 1e-9 * max(1.0, abs(b))
    if hasattr(a, 'year') and hasattr(b, 'year'):
        return abs((a - b).total_seconds()) < 1
    return (a if a not in (None, '') else '') == (b if b not in (None, '') else '')


def foto(ex, wb):
    sh = {s.Name: s for s in wb.Worksheets}
    ex.xl.Calculate()
    return {
        'calc': sh['Calc'].Range('A1:BI182').Value,
        'painel': sh['Painel'].Range('A3:Y11').Value,
        'eng1': sh['Eng_Saida'].Range('A1:Q1').Value,
        'eng': sh['Eng_Saida'].Range('A3:AB182').Value,
        'engstat': sh['Eng_Saida'].Range('A185:U186').Value,
        'eventos': sh['Eventos_Westgard'].Range('A1:N9000').Value,
        'estat': sh['Estatística'].Range('A14:AB93').Value,
        'liber': sh['Liberação'].Range('A4:B203').Value,
    }


def comparar(rot, a, b):
    difs = []
    for k in a:
        va, vb = a[k], b[k]
        for i, (ra, rb) in enumerate(zip(va, vb)):
            for j, (x, y) in enumerate(zip(ra, rb)):
                if not igual(x, y):
                    difs.append((k, i, j, x, y))
    print(f'  {rot}: {"IGUAL" if not difs else str(len(difs)) + " diferencas"}')
    for d in difs[:8]:
        print('     ', d)
    return difs


def selecionar(ex, lote=None, analito=None, de=None, ate=None, pai=None):
    if de is not None or ate is not None:
        pai.Range('G3').Value = de if de else ''
        pai.Range('G4').Value = ate if ate else ''
    if lote:
        ex.run('mTesteEtapa1.T_Selecionar', lote)
    if analito:
        ex.run('mTesteEtapa1.T_Analito', analito)
    ex.run('mEstatistica.RecalcularAnalitoAtual')
    ex.esperar()


COMBOS = []


def montar_combos(wb):
    db = wb.Sheets('DB_Resultados')
    ult = db.Cells(db.Rows.Count, 1).End(-4162).Row
    datas = [r[0] for r in db.Range(f'B4:B{ult}').Value if hasattr(r[0], 'year')]
    fim = max(datas).replace(tzinfo=None)
    fim = dt.datetime(fim.year, fim.month, fim.day)
    ini = fim - dt.timedelta(days=30)
    lotes = sorted({str(r[0])[3:-2] for r in db.Range(f'D4:D{ult}').Value if r[0]})
    ans = [wb.Sheets('Analitos').Cells(r, 1).Value for r in range(4, 44)]
    ans = [a for a in ans if a]
    l1, l2 = lotes[0], lotes[-1]
    lm = lotes[len(lotes) // 2]
    COMBOS.extend([
        (lm, ans[0], None, None),
        (lm, ans[2], None, None),
        (l2, ans[0], None, None),
        (l2, ans[5], ini, fim),
        (l1, ans[0], ini, fim),
        (l1, ans[7], None, None),
    ])
    globals()['LOTES'] = lotes


def main(src):
    dst = os.path.join(TMP, 'prova_v2.xlsm')
    shutil.copyfile(src, dst)
    ex = xlh.Excel()
    wb = ex.abrir(dst)
    try:
        sh = {s.Name: s for s in wb.Worksheets}
        pai, est = sh['Painel'], sh['Estatística']
        if 'mTesteEtapa1' not in [c.Name for c in wb.VBProject.VBComponents]:
            with open(os.path.join(os.path.dirname(os.path.abspath(__file__)), 'mTesteEtapa1.bas'), encoding='utf-8') as fh:
                txt = CRLF.join(l for l in fh.read().replace(CRLF, LF).split(LF) if not l.startswith('Attribute '))
            comp = wb.VBProject.VBComponents.Add(1)
            comp.Name = 'mTesteEtapa1'
            comp.CodeModule.AddFromString(txt)
        for s in wb.Worksheets:
            if s.ProtectContents:
                s.Unprotect('qcini2025')
        for col in ('R', 'S', 'T', 'AC', 'AD'):
            f = est.Range(f'{col}14').Formula
            if 'eqProvedor' in f:
                est.Range(f'{col}14:{col}93').Formula = f.replace('eqProvedor', PROV)
        ex.xl.CalculateFull(); ex.esperar()
        ex.run('mEstatistica.AtualizarEstatistica', teto=600)
        montar_combos(wb)
        antes = []
        for lote, an, de, ate in COMBOS:
            selecionar(ex, lote, an, de or '', ate or '', pai)
            antes.append(foto(ex, wb))
        print('fotos ANTES:', len(antes))

        # ---- instala a v2 ----
        vbp = wb.VBProject
        for m in ('mApp', 'mUI', 'mBanco', 'mLotes', 'mEstatistica', 'mOperacao', 'mImportar', 'mCEQ', 'mDados'):
            v2.substituir_codigo(vbp, m, v2.codigo(m + '.bas'))
        v2.aplicar_planilhas(ex, wb, 2)
        ex.xl.CalculateFull(); ex.esperar()
        ex.run('mEstatistica.AtualizarEstatistica', teto=600)
        print('v2 instalada')

        difs = 0
        for (lote, an, de, ate), a in zip(COMBOS, antes):
            selecionar(ex, lote, an, de or '', ate or '', pai)
            difs += len(comparar(f'{lote} {an} {de and de.date()}', a, foto(ex, wb)))
        print('TOTAL de diferencas:', difs)

        # ---- tempos ----
        selecionar(ex, LOTES[-1], None, '', '', pai)
        tt = []
        for idx in (3, 5, 1, 7, 2, 9, 4, 6):
            t0 = time.time()
            pai.Range('B3').Value = idx
            ex.run('PainelMudou')
            ex.esperar()
            tt.append(time.time() - t0)
        print('spinner (8 cliques): ' + ' '.join(f'{x:.2f}' for x in tt) + f'  | max {max(tt):.2f}s')
        for lote in (LOTES[0], LOTES[-1], LOTES[len(LOTES) // 2]):
            t0 = time.time(); ex.run('mTesteEtapa1.T_Selecionar', lote); ex.esperar()
            print(f'troca de lote em analise -> {lote}: {time.time() - t0:.2f}s')
        for lote in (LOTES[0], LOTES[-1]):
            t0 = time.time(); ex.run('mTesteEtapa1.T_LoteEmUso', lote); ex.esperar()
            print(f'troca de lote EM USO -> {lote}: {time.time() - t0:.2f}s')
        # ---- importar UMA corrida (todos os analitos x 2 niveis) pela aba Importar ----
        imp = sh['Importar']
        col_imp, c = {}, 5
        while imp.Cells(4, c).Value not in (None, ''):
            nm = imp.Cells(3, c).Value
            if nm:
                col_imp[str(nm).strip()] = c
            c += 1
        ncol = max(col_imp.values())
        db = sh['DB_Resultados']
        ult = db.Cells(db.Rows.Count, 1).End(-4162).Row
        fim = max(r[0] for r in db.Range(f'B4:B{ult}').Value if hasattr(r[0], 'year')).replace(tzinfo=None)
        for rodada in range(2):
            dia = fim + dt.timedelta(days=3 + rodada)
            area = []
            for nv in (1, 2):
                row = [None] * (ncol - 1)
                row[0] = dia.strftime('%d/%m/%Y'); row[1] = nv; row[2] = LOTES[-1]
                for an, cc in col_imp.items():
                    row[cc - 2] = '12,5' if nv == 1 else '37,5'
                area.append(row)
            rg = imp.Range(imp.Cells(5, 2), imp.Cells(4 + len(area), ncol))
            rg.NumberFormat = '@'
            rg.Value = area
            t0 = time.time()
            r = ex.run('mImportar.ExecutarImportacao', True, teto=900); ex.esperar()
            print(f'importar 1 corrida ({len(col_imp)} analitos x 2 niveis): {time.time() - t0:.2f}s -> {r}')
        print('calculo ao fim:', ex.xl.Calculation, '(automatico = -4105)')
        # ---- sem formula varrendo o banco? ----
        for nmx in ('Calc', 'Liberação', 'Painel', 'Estatística'):
            ws = sh[nmx]
            try:
                fr = ws.UsedRange.SpecialCells(-4123)
                txt = [c.Formula for c in fr if any(k in c.Formula for k in ('rRUN', 'rData', 'rValor', 'rAnalito', 'rLote', 'rStatus', 'DB_Resultados'))]
            except Exception:
                txt = []
            print(f'   formulas sobre o banco em {nmx}: {len(txt)}')
    finally:
        ex.fechar()


if __name__ == '__main__':
    main(os.path.abspath(sys.argv[1]))
