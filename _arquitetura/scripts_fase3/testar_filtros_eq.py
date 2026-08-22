# -*- coding: utf-8 -*-
"""testar_filtros_eq.py - os tres filtros do controle externo respondem?

Nao basta a formula existir: o teste TROCA provedor, ano e rodada e confere que
o numero muda de acordo -- e que continua batendo com a media calculada a mao a
partir das proprias linhas da EQC_Dados.

Cria rodadas de CAP com valores conhecidos para que exista o que comparar entre
os dois provedores; ao fim, remove o que criou.

Uso: python testar_filtros_eq.py <arquivo.xlsm>
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
    copia = os.path.join(os.environ.get('TEMP', '.'), 'tfeq_' + os.path.basename(caminho))
    shutil.copy(caminho, copia)
    xl = novo_excel()
    wb = xl.Workbooks.Open(copia)
    try:
        try:
            wb.Unprotect('qcini2025')
        except Exception:
            pass
        for nm in ('EQC_Dados', 'Estatística'):
            try:
                wb.Worksheets(nm).Unprotect('qcini2025')
            except Exception:
                pass
        eq = wb.Worksheets('EQC_Dados')
        es = wb.Worksheets('Estatística')
        xl.Calculation = -4105

        ult = eq.Cells(eq.Rows.Count, 1).End(-4162).Row
        alvo = str(eq.Cells(EQ_R0, 1).Value).strip()
        ano = int(eq.Cells(EQ_R0, 2).Value)
        print('analito de teste: %s (ja tem Controllab em %d)' % (alvo, ano))

        # ---- rodadas de CAP com bias conhecido --------------------------
        # A = +4%, B = -8%, C = +6%  ->  |bias| medio = 6, assinado = +2/3
        novas = [('A', 104.0), ('B', 92.0), ('C', 106.0)]
        base = ult + 1
        for k, (rod, xlab) in enumerate(novas):
            lin = base + k
            eq.Cells(lin, 1).Value = alvo
            eq.Cells(lin, 2).Value = ano
            eq.Cells(lin, 3).Value = rod
            eq.Cells(lin, 5).Value = 'CAP'
            eq.Cells(lin, 6).Value = '01'
            eq.Cells(lin, 7).Value = xlab
            eq.Cells(lin, 8).Value = 100.0
            eq.Cells(lin, 9).Value = 5.0        # SD grupo -> SDI = (xlab-100)/5
            eq.Cells(lin, 11).Value = 90.0      # limites
            eq.Cells(lin, 12).Value = 110.0
            # Copiar J..P inteiro sobrescreveria K e L, que sao DADO DIGITADO
            # (os limites do grupo), nao formula. Foi o que fez o teste acusar
            # "3 fora dos limites" onde so uma amostra estava fora: as tres
            # herdaram os limites da linha 4, de outra escala.
            for faixa in ((10, 10), (13, 16)):      # J ; M..P
                tenta(lambda f=faixa: eq.Range(eq.Cells(EQ_R0, f[0]),
                                               eq.Cells(EQ_R0, f[1])).Copy())
                tenta(lambda l=lin, f=faixa: eq.Range(eq.Cells(l, f[0]),
                                                      eq.Cells(l, f[1])).PasteSpecial(-4123))
                xl.CutCopyMode = False
        tenta(lambda: xl.CalculateFull())
        print('3 rodadas de CAP criadas: A=+4%%, B=-8%%, C=+6%% (SD grupo 5)')

        def bias(prov, rod, modo='ABS'):
            return tenta(lambda: xl.Run('BiasEQ', alvo, ano, modo, prov, rod))

        print()
        print('=== 1. FILTRO DE PROVEDOR ===')
        cap = bias('CAP', 'TODAS')
        ctl = bias('Controllab', 'TODAS')
        todos = bias('', 'TODAS')
        print('   CAP=%s   Controllab=%s   todos=%s' % (cap, ctl, todos))
        ck('CAP consolida as 3 rodadas: |bias| = 6%', perto(cap, 6.0), str(cap))
        ck('Controllab difere do CAP', not perto(cap, ctl), '%s vs %s' % (cap, ctl))
        ck('sem filtro fica entre os dois',
           min(float(cap), float(ctl)) <= float(todos) <= max(float(cap), float(ctl)),
           str(todos))

        print()
        print('=== 2. FILTRO DE RODADA ===')
        for rod, esperado in (('A', 4.0), ('B', 8.0), ('C', 6.0)):
            v = bias('CAP', rod)
            s = bias('CAP', rod, 'SIGNED')
            print('   CAP rodada %s: |bias|=%s  assinado=%s' % (rod, v, s))
            ck('rodada %s: |bias| = %.0f%%' % (rod, esperado), perto(v, esperado), str(v))
        ck('rodada B tem sinal negativo', perto(bias('CAP', 'B', 'SIGNED'), -8.0),
           str(bias('CAP', 'B', 'SIGNED')))
        ck('rodada inexistente (D) devolve SEM EP',
           isinstance(bias('CAP', 'D'), str), repr(bias('CAP', 'D')))

        print()
        print('=== 3. MEDIA DAS MAGNITUDES, NAO DOS ASSINADOS ===')
        a = bias('CAP', 'TODAS', 'ABS')
        s = bias('CAP', 'TODAS', 'SIGNED')
        print('   |bias| medio=%s   assinado medio=%s' % (a, s))
        ck('|bias| = (4+8+6)/3 = 6', perto(a, 6.0), str(a))
        ck('assinado = (4-8+6)/3 = 0,667', perto(s, (4 - 8 + 6) / 3.0), str(s))
        ck('as duas NAO coincidem (o sinal seria cancelado)', not perto(a, s),
           '%s vs %s' % (a, s))

        print()
        print('=== 4. STATUS SDI (limite |2|) ===')
        # SDI: A=(104-100)/5=0,8  B=(92-100)/5=-1,6  C=(106-100)/5=1,2 -> max 1,6 OK
        st = tenta(lambda: xl.Run('StatusSDIeq', alvo, ano, 'CAP', 'TODAS'))
        mx = tenta(lambda: xl.Run('SDIeq', alvo, ano, 'MAX', 'CAP', 'TODAS'))
        print('   max |SDI| = %s -> %s' % (mx, st))
        ck('max |SDI| = 1,6', perto(mx, 1.6), str(mx))
        ck('status aprova com |SDI| <= 2', 'OK' in str(st), str(st))
        # agora uma amostra fora
        eq.Cells(base, 7).Value = 115.0        # SDI = 3,0 e fora do limite 110
        tenta(lambda: xl.CalculateFull())
        st2 = tenta(lambda: xl.Run('StatusSDIeq', alvo, ano, 'CAP', 'TODAS'))
        lim2 = tenta(lambda: xl.Run('StatusLimitesEQ', alvo, ano, 'CAP', 'TODAS'))
        print('   com amostra de 115: SDI -> %s ; limites -> %s' % (st2, lim2))
        ck('|SDI| > 2 reprova', 'FORA' in str(st2), str(st2))
        ck('fora dos limites reprova', 'NAO OK' in str(lim2), str(lim2))

        print()
        print('=== 5. OS FILTROS DA ABA MOVEM A ESTATISTICA ===')
        linEst = None
        for r in range(14, 94):
            if str(es.Cells(r, 1).Value or '').strip() == alvo:
                linEst = r
                break
        es.Range('N4').Value = ano
        for prov in ('CAP', 'Controllab'):
            es.Range('L4').Value = prov
            es.Range('P4').Value = 'TODAS'
            tenta(lambda: xl.CalculateFull())
            k = es.Cells(linEst, 11).Value
            esperado = bias(prov, 'TODAS')
            print('   provedor=%-11s Estatistica!K=%s  (BiasEQ=%s)  T=%s'
                  % (prov, k, esperado, str(es.Cells(linEst, 20).Value)[:34]))
            ck('K acompanha o provedor %s' % prov, perto(k, esperado),
               '%s vs %s' % (k, esperado))
        es.Range('L4').Value = 'CAP'
        es.Range('P4').Value = 'B'
        tenta(lambda: xl.CalculateFull())
        k = es.Cells(linEst, 11).Value
        print('   rodada=B -> Estatistica!K=%s' % k)
        ck('K acompanha a rodada B', perto(k, 8.0), str(k))
        print('   resumo: %r' % es.Cells(5, 11).Value)
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
    print('FILTROS DO CONTROLE EXTERNO: TODOS RESPONDEM')


if __name__ == '__main__':
    main(sys.argv[1])
