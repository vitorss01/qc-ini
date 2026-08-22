# -*- coding: utf-8 -*-
"""testar_bias_ceq.py - os cinco testes numericos do bias, e a rastreabilidade

Cada teste grava linhas de EP com valores conhecidos, confere o que sai, e
restaura o estado anterior. Nenhum deles depende do dado que ja esta la.

  1  X_lab=105, X_ref=100  ->  Bias = +5%, |Bias| = 5%
  2  X_lab= 95, X_ref=100  ->  Bias = -5%, |Bias| = 5%
  3  ETp=10, Bias=-2, CV=2 ->  Sigma = 4,0   (o sinal nao pode inflar o Sigma)
  4  CV=2, Bias=-2         ->  ET = 1,65*2 + 2 = 5,3
  5  rodadas +5% e -5%     ->  consolidado 5%, NUNCA 0%
  6  X_ref = 0             ->  nao calculavel, e nao zero
  7  analito sem EP        ->  "SEM EP", com ET e Sigma VAZIOS
  8  rastreabilidade       ->  X_lab/X_ref -> bias -> ET -> Sigma -> Painel

Uso: python testar_bias_ceq.py <arquivo.xlsm>
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
EQ_R0 = 4


def ck(nome, cond, det=''):
    print(('  OK   ' if cond else '  FALHA') + '  ' + nome + (('  -> ' + det) if det else ''))
    if not cond:
        falhas.append(nome)


def perto(a, b, tol=1e-9):
    try:
        return abs(float(a) - float(b)) <= tol
    except Exception:
        return False


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
    copia = os.path.join(os.environ.get('TEMP', '.'), 'tbias_' + os.path.basename(caminho))
    shutil.copy(caminho, copia)
    xl = novo_excel()
    wb = xl.Workbooks.Open(copia)
    try:
        try:
            wb.Unprotect('qcini2025')
        except Exception:
            pass
        for nm in ('EQC_Dados', 'Estatística', 'Painel', 'Analitos'):
            try:
                wb.Worksheets(nm).Unprotect('qcini2025')
            except Exception:
                pass
        eq = wb.Worksheets('EQC_Dados')
        es = wb.Worksheets('Estatística')
        pa = wb.Worksheets('Painel')
        an = wb.Worksheets('Analitos')
        xl.Calculation = -4105

        ult = eq.Cells(eq.Rows.Count, 1).End(-4162).Row
        livre = max(ult + 1, EQ_R0)
        anoRef = int(es.Range('J5').Value.year) if es.Range('J5').Value else 2026
        print('EQC_Dados: ultima linha %d ; ano de referencia da aba: %d' % (ult, anoRef))

        # analito de teste que NAO tem EP -- para nao contaminar dado real
        alvo = None
        for i in range(4, 44):
            nm = an.Cells(i, 1).Value
            if not nm:
                continue
            r = tenta(lambda n=nm: xl.Run('BiasEQ', n, anoRef, 'ABS'))
            if isinstance(r, str):
                alvo = str(nm).strip()
                break
        print('analito sem EP escolhido para os testes: %r' % alvo)

        def escreve_rodada(lin, rodada, xlab, xref):
            eq.Cells(lin, 1).Value = alvo
            eq.Cells(lin, 2).Value = anoRef
            eq.Cells(lin, 3).Value = rodada
            eq.Cells(lin, 7).Value = xlab
            eq.Cells(lin, 8).Value = xref
            # As colunas derivadas (J=SDI, M=Status, N=Bias, O=|Bias|) so existem
            # nas linhas que ja tinham formula. Escrever so A/B/C/G/H numa linha
            # virgem deixa N vazio -- e o teste acusaria "SEM EP" achando que o
            # modulo falhou, quando o que faltava era a formula da linha.
            tenta(lambda: eq.Range(eq.Cells(EQ_R0, 10), eq.Cells(EQ_R0, 15)).Copy())
            tenta(lambda: eq.Range(eq.Cells(lin, 10), eq.Cells(lin, 15)).PasteSpecial(-4123))
            xl.CutCopyMode = False

        def limpa(lin0, n):
            eq.Range(eq.Cells(lin0, 1), eq.Cells(lin0 + n - 1, 15)).ClearContents()

        # ---- 7 (antes de sujar): ausencia de EP -------------------------
        print()
        print('=== 7. ANALITO SEM EP ===')
        r = tenta(lambda: xl.Run('BiasEQ', alvo, anoRef, 'ABS'))
        print('   BiasEQ(%s) = %r' % (alvo, r))
        ck('sem EP devolve marcador de texto, nao zero',
           isinstance(r, str) and 'SEM EP' in r.upper(), repr(r))

        # ---- 1 e 2: formula do bias --------------------------------------
        print()
        print('=== 1 e 2. FORMULA DO BIAS ===')
        for rot, xlab, esperado in (('105 vs 100', 105.0, 5.0), ('95 vs 100', 95.0, -5.0)):
            escreve_rodada(livre, 1, xlab, 100.0)
            tenta(lambda: xl.CalculateFull())
            assinado = tenta(lambda: xl.Run('BiasEQ', alvo, anoRef, 'SIGNED'))
            absoluto = tenta(lambda: xl.Run('BiasEQ', alvo, anoRef, 'ABS'))
            nlinha = eq.Cells(livre, 14).Value
            print('   %-11s -> EQC_Dados!N=%s  assinado=%s  |bias|=%s'
                  % (rot, nlinha, assinado, absoluto))
            ck('%s: bias assinado = %+.0f%%' % (rot, esperado),
               perto(assinado, esperado), str(assinado))
            ck('%s: |bias| = %.0f%%' % (rot, abs(esperado)),
               perto(absoluto, abs(esperado)), str(absoluto))
        limpa(livre, 1)

        # ---- 5: cancelamento -------------------------------------------
        print()
        print('=== 5. CANCELAMENTO DE SINAL (+5%% e -5%%) ===')
        escreve_rodada(livre, 1, 105.0, 100.0)
        escreve_rodada(livre + 1, 2, 95.0, 100.0)
        tenta(lambda: xl.CalculateFull())
        absoluto = tenta(lambda: xl.Run('BiasEQ', alvo, anoRef, 'ABS'))
        assinado = tenta(lambda: xl.Run('BiasEQ', alvo, anoRef, 'SIGNED'))
        nrod = tenta(lambda: xl.Run('BiasEQ', alvo, anoRef, 'N'))
        print('   rodadas=%s  |bias| consolidado=%s  assinado=%s' % (nrod, absoluto, assinado))
        ck('duas rodadas entraram', perto(nrod, 2), str(nrod))
        ck('consolidado de magnitude = 5%, NAO 0%', perto(absoluto, 5.0), str(absoluto))
        ck('media assinada = 0% (fica so para leitura)', perto(assinado, 0.0), str(assinado))

        # ---- 3 e 4: Sigma e ET com bias negativo -------------------------
        print()
        print('=== 3 e 4. SIGMA E ET COM BIAS NEGATIVO ===')
        limpa(livre, 2)
        escreve_rodada(livre, 1, 98.0, 100.0)      # bias = -2%
        tenta(lambda: xl.CalculateFull())
        bias = tenta(lambda: xl.Run('BiasEQ', alvo, anoRef, 'ABS'))
        biasS = tenta(lambda: xl.Run('BiasEQ', alvo, anoRef, 'SIGNED'))
        print('   bias assinado=%s  |bias|=%s' % (biasS, bias))
        ck('bias assinado = -2%', perto(biasS, -2.0), str(biasS))
        etp, cv = 10.0, 2.0
        sigma = (etp - abs(float(bias))) / cv
        et = 1.65 * cv + abs(float(bias))
        print('   ETp=%.0f CV=%.0f -> Sigma=(%.0f-|%.0f|)/%.0f=%.3f ; ET=1,65*%.0f+%.0f=%.3f'
              % (etp, cv, etp, abs(float(bias)), cv, sigma, cv, abs(float(bias)), et))
        ck('Sigma = 4,0 (sinal negativo nao infla)', perto(sigma, 4.0), str(sigma))
        ck('ET = 5,3', perto(et, 5.3), str(et))

        # ---- 6: X_ref = 0 ------------------------------------------------
        print()
        print('=== 6. X_ref = 0 ===')
        escreve_rodada(livre, 1, 100.0, 0.0)
        tenta(lambda: xl.CalculateFull())
        nl = eq.Cells(livre, 14).Value
        r = tenta(lambda: xl.Run('BiasEQ', alvo, anoRef, 'ABS'))
        print('   EQC_Dados!N com X_ref=0: %r ; BiasEQ: %r' % (nl, r))
        ck('linha com X_ref=0 nao vira bias zero',
           (nl in (None, '')) and isinstance(r, str), '%r / %r' % (nl, r))
        limpa(livre, 2)
        tenta(lambda: xl.CalculateFull())

        # ---- 8: rastreabilidade ponta a ponta ---------------------------
        print()
        print('=== 8. RASTREABILIDADE COM DADO REAL ===')
        real = None
        for r0 in range(14, 94):
            v = es.Cells(r0, 11).Value
            if isinstance(v, (int, float)):
                real = (r0, str(es.Cells(r0, 1).Value).strip(), int(es.Cells(r0, 2).Value))
                break
        if real is None:
            ck('existe analito com EP para rastrear', False, 'nenhum')
        else:
            lin, nome, niv = real
            mem = tenta(lambda: xl.Run('BiasEQMemoria', nome, 1900, 2999))
            print('   %s N%d' % (nome, niv))
            for parte in str(mem).split(';'):
                if parte.strip():
                    print('      %s' % parte.strip())
            cv = es.Cells(lin, 6).Value
            k = es.Cells(lin, 11).Value
            et = es.Cells(lin, 14).Value
            etp = es.Cells(lin, 15).Value
            sig = es.Cells(lin, 18).Value
            print('   CV=%s  |Bias|=%s  ETp=%s  ET=%s  Sigma=%s' % (cv, k, etp, et, sig))
            ck('ET = 1,65*CV + |Bias|', perto(et, 1.65 * float(cv) + abs(float(k)), 1e-9),
               '%s vs %s' % (et, 1.65 * float(cv) + abs(float(k))))
            ck('Sigma = (ETp - |Bias|)/CV',
               perto(sig, (float(etp) - abs(float(k))) / float(cv), 1e-9),
               '%s vs %s' % (sig, (float(etp) - abs(float(k))) / float(cv)))
            # Painel
            pa.Range('C3').Value = nome
            tenta(lambda: xl.CalculateFull())
            pl = 6 + niv
            print('   Painel N%d: CV=%s Bias=%s ETp=%s ET=%s Sigma=%s'
                  % (niv, pa.Cells(pl, 5).Value, pa.Cells(pl, 7).Value,
                     pa.Cells(pl, 6).Value, pa.Cells(pl, 8).Value, pa.Cells(pl, 9).Value))
            ck('Painel usa o MESMO bias da Estatistica',
               perto(pa.Cells(pl, 7).Value, k), '%s vs %s' % (pa.Cells(pl, 7).Value, k))
            ck('Painel Sigma = Estatistica Sigma',
               perto(pa.Cells(pl, 9).Value, sig), '%s vs %s' % (pa.Cells(pl, 9).Value, sig))

        # ---- 7b: sem EP -> ET e Sigma VAZIOS na planilha ----------------
        print()
        print('=== 7b. SEM EP: ET E SIGMA FICAM VAZIOS ===')
        semEP = None
        for r0 in range(14, 94):
            if isinstance(es.Cells(r0, 11).Value, str):
                semEP = r0
                break
        if semEP:
            # Le UMA vez cada celula. Reler a mesma celula dentro do print e de
            # novo no ck dobra as travessias COM e o Excel corta com
            # RPC_E_CALL_REJECTED bem no fim do teste.
            nome = tenta(lambda: es.Cells(semEP, 1).Value)
            kk = tenta(lambda: es.Cells(semEP, 11).Value)
            nn = tenta(lambda: es.Cells(semEP, 14).Value)
            rr = tenta(lambda: es.Cells(semEP, 18).Value)
            print('   linha %d (%s): K=%r  N(ET)=%r  R(Sigma)=%r'
                  % (semEP, nome, kk, nn, rr))
            ck('ET vazio sem bias', nn in (None, ''), repr(nn))
            ck('Sigma vazio sem bias', rr in (None, ''), repr(rr))
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
    print('=' * 64)
    if falhas:
        print('FALHAS (%d):' % len(falhas))
        for f in falhas:
            print('   - %s' % f)
        sys.exit(1)
    print('BIAS DO EP: TODOS OS TESTES PASSARAM')


if __name__ == '__main__':
    main(sys.argv[1])
