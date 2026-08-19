# -*- coding: utf-8 -*-
"""testar_ssot_propagacao.py - prova a cadeia R -> S/T -> Estatistica -> Sigma -> Painel

Troca a fonte em Analitos!R e confere que TODA a cadeia acompanha, sem tocar em
mais nada. E o requisito do item 17 do pedido: mudar R deve bastar.

Confere tambem o item 16 -- nenhuma aba pode exibir um ETp/CVTp diferente do que
Analitos!S/T dizem.

Uso: python testar_ssot_propagacao.py <arquivo.xlsm>
"""
import io
import os
import sys
import time
import shutil
import subprocess

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', write_through=True)
import win32com.client as w

falhas = []


def ck(nome, cond, det=''):
    print(('  OK   ' if cond else '  FALHA') + '  ' + nome + (('  -> ' + det) if det else ''))
    if not cond:
        falhas.append(nome)


def novo_excel():
    for t in range(1, 9):
        try:
            xl = w.DispatchEx('Excel.Application')
            xl.Visible = False
            xl.DisplayAlerts = False
            xl.EnableEvents = False
            xl.AutomationSecurity = 1
            return xl
        except Exception:
            if t in (2, 4):
                subprocess.call(['powershell', '-NoProfile', '-Command',
                                 'Start-Process excel.exe -WindowStyle Hidden'],
                                stderr=subprocess.DEVNULL)
                time.sleep(8)
            time.sleep(2.5 * t)
    raise RuntimeError('Excel COM nao subiu')


def tenta(fn, vezes=8):
    ult = None
    for i in range(vezes):
        try:
            return fn()
        except Exception as e:
            ult = e
            s = str(e).lower()
            if 'rejeitada' not in s and 'rejected' not in s:
                raise
            time.sleep(1.0 + 0.8 * i)
    raise ult


def main(caminho):
    copia = os.path.join(os.environ.get('TEMP', '.'), 'prop_' + os.path.basename(caminho))
    shutil.copy(caminho, copia)
    xl = novo_excel()
    wb = xl.Workbooks.Open(copia)
    try:
        try:
            wb.Unprotect('qcini2025')
        except Exception:
            pass
        for nm in ('Analitos', 'Painel', 'Estatística'):
            try:
                wb.Worksheets(nm).Unprotect('qcini2025')
            except Exception:
                pass
        an = wb.Worksheets('Analitos')
        es = wb.Worksheets('Estatística')
        pa = wb.Worksheets('Painel')
        xl.Calculation = -4105

        # analito de teste: precisa ter dado nas TRES fontes para o teste valer
        alvo = None
        for i in range(4, 44):
            nome = an.Cells(i, 1).Value
            if not nome:
                continue
            k = an.Cells(i, 11).Value      # K = TEa CLIA
            q = an.Cells(i, 17).Value      # Q = ETp VB
            m = an.Cells(i, 13).Value      # M = ET FAB
            if all(isinstance(x, (int, float)) and x for x in (k, q, m)):
                alvo = (i, nome)
                break
        if alvo is None:
            # sem um analito com as tres, usa um com CLIA e VB
            for i in range(4, 44):
                nome = an.Cells(i, 1).Value
                if nome and isinstance(an.Cells(i, 11).Value, (int, float)) \
                        and isinstance(an.Cells(i, 17).Value, (int, float)):
                    alvo = (i, nome)
                    break
        lin, nome = alvo
        print('analito de teste: %s (linha %d)' % (nome, lin))
        print('  K(TEa CLIA)=%r  Q(ETp VB)=%r  M(ET FAB)=%r'
              % (an.Cells(lin, 11).Value, an.Cells(lin, 17).Value, an.Cells(lin, 13).Value))
        orig = an.Cells(lin, 18).Value

        # localiza a linha da Estatistica desse analito, nivel 1
        linEst = None
        for r in range(14, 94):
            if str(es.Cells(r, 1).Value or '').strip() == str(nome).strip() \
                    and str(es.Cells(r, 2).Value or '').strip() in ('1', '1.0'):
                linEst = r
                break
        if linEst is None:
            for r in range(14, 94):
                if str(es.Cells(r, 1).Value or '').strip() == str(nome).strip():
                    linEst = r
                    break
        print('  linha na Estatistica: %s' % linEst)

        pa.Range('C3').Value = nome
        tenta(lambda: xl.CalculateFull())

        print()
        for fonte in ('CLIA', 'VB', 'FAB'):
            an.Cells(lin, 18).Value = fonte
            tenta(lambda: xl.CalculateFull())
            S = an.Cells(lin, 19).Value
            T = an.Cells(lin, 20).Value
            O = es.Cells(linEst, 15).Value if linEst else None
            G = es.Cells(linEst, 7).Value if linEst else None
            sig = es.Cells(linEst, 18).Value if linEst else None
            pF = pa.Cells(7, 6).Value
            pI = pa.Cells(7, 9).Value
            print('R=%-5s | S(ETp)=%-28r T(CVTp)=%-20r' % (fonte, S, T))
            print('        | Estat O=%-28r G=%-20r Sigma=%r' % (O, G, sig))
            print('        | Painel F7=%-26r Sigma I7=%r' % (pF, pI))

            ck('R=%s: Estatistica!O == Analitos!S' % fonte,
               (O == S) or (O in (None, '') and S in (None, '')), '%r vs %r' % (O, S))
            ck('R=%s: Estatistica!G == Analitos!T' % fonte,
               (G == T) or (G in (None, '') and T in (None, '')), '%r vs %r' % (G, T))
            ck('R=%s: Painel!F7 == Analitos!S' % fonte,
               (pF == S) or (pF in (None, '') and S in (None, '')), '%r vs %r' % (pF, S))
            if isinstance(S, str):
                ck('R=%s: ETp em texto -> Sigma VAZIO (nao zero, nao erro)' % fonte,
                   sig in (None, ''), repr(sig))
            print()

        an.Cells(lin, 18).Value = orig
        tenta(lambda: xl.CalculateFull())
        print('fonte restaurada para %r' % orig)

        # ---- item 16: nenhuma divergencia em NENHUM analito ---------------
        print()
        print('=== CONSISTENCIA GLOBAL (todos os analitos, nivel 1) ===')
        div = 0
        for r in range(14, 94):
            nm = str(es.Cells(r, 1).Value or '').strip()
            if not nm:
                continue
            ia = None
            for i in range(4, 44):
                if str(an.Cells(i, 1).Value or '').strip() == nm:
                    ia = i
                    break
            if ia is None:
                continue
            if es.Cells(r, 15).Value != an.Cells(ia, 19).Value:
                div += 1
                if div <= 3:
                    print('   DIVERGE %s: Estat O=%r vs Analitos S=%r'
                          % (nm, es.Cells(r, 15).Value, an.Cells(ia, 19).Value))
            if es.Cells(r, 7).Value != an.Cells(ia, 20).Value:
                div += 1
                if div <= 3:
                    print('   DIVERGE %s: Estat G=%r vs Analitos T=%r'
                          % (nm, es.Cells(r, 7).Value, an.Cells(ia, 20).Value))
        ck('nenhum ETp/CVTp diverge de Analitos S/T', div == 0, '%d divergencia(s)' % div)

        # ---- item 18: nada quebrado ---------------------------------------
        print()
        ref = 0
        for ws in (an, es, pa):
            for cel in ws.UsedRange:
                v = cel.Text
                if isinstance(v, str) and '#REF!' in v:
                    ref += 1
        ck('nenhum #REF! em Analitos/Estatistica/Painel', ref == 0, str(ref))
    finally:
        try:
            wb.Close(False)
        except Exception:
            pass
        try:
            xl.Quit()
        except Exception:
            pass

    print()
    print('=' * 62)
    if falhas:
        print('FALHAS (%d):' % len(falhas))
        for f in falhas:
            print('   - %s' % f)
        sys.exit(1)
    print('PROPAGACAO R -> S/T -> ESTATISTICA -> SIGMA -> PAINEL: INTEGRA')


if __name__ == '__main__':
    main(sys.argv[1])
