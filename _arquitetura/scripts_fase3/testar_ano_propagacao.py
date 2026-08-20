# -*- coding: utf-8 -*-
"""testar_ano_propagacao.py - prova a cadeia S2 -> A4:W43 -> Estatistica/Painel

Os blocos de 2025 e 2026 sao HOJE identicos (conferido: 0 celulas diferentes).
Trocar o ano sem mais nada nao provaria nada -- os numeros ficariam iguais por
coincidencia, e o teste passaria sem testar. Por isso o teste GRAVA UM MARCADOR
no bloco de 2026, confere que ele aparece na ponta, e depois restaura.

Cobre os testes 1 a 4 do pedido: troca de ano, segundo ano, integridade e
reabertura.

Uso: python testar_ano_propagacao.py <arquivo.xlsm>
"""
import io
import os
import sys
import time
import shutil
import subprocess

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', write_through=True)
import win32com.client as w

MARCADOR = 99.0
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


def destravar(wb):
    try:
        wb.Unprotect('qcini2025')
    except Exception:
        pass
    for s in wb.Worksheets:
        try:
            s.Unprotect('qcini2025')
        except Exception:
            pass


def abrir(caminho):
    xl = novo_excel()
    wb = xl.Workbooks.Open(caminho)
    destravar(wb)
    xl.Calculation = -4105
    return xl, wb


def main(caminho):
    copia = os.path.join(os.environ.get('TEMP', '.'), 'anoprop_' + os.path.basename(caminho))
    shutil.copy(caminho, copia)
    xl, wb = abrir(copia)
    try:
        an = wb.Worksheets('Analitos')
        es = wb.Worksheets('Estatística')
        pa = wb.Worksheets('Painel')

        # analito de teste: com fonte CLIA e ETp CLIA numerico
        alvo = None
        for i in range(4, 44):
            nome = an.Cells(i, 1).Value
            if nome and str(an.Cells(i, 19).Value or '').strip() == 'CLIA' \
                    and isinstance(an.Cells(i, 11).Value, (int, float)) \
                    and an.Cells(i, 11).Value:
                alvo = (i, str(nome).strip())
                break
        lin, nome = alvo
        print('analito de teste: %s (linha %d, fonte CLIA)' % (nome, lin))

        # coluna dele no bloco historico
        colHist = None
        for c in range(2, 42):
            if str(an.Cells(47, c).Value or '').strip() == nome:
                colHist = c
                break
        print('coluna no historico: %d' % colHist)

        linEst = None
        for r in range(14, 94):
            if str(es.Cells(r, 1).Value or '').strip() == nome:
                linEst = r
                break
        pa.Range('C3').Value = nome
        tenta(lambda: xl.CalculateFull())

        def leitura():
            return dict(
                K=an.Cells(lin, 11).Value, T=an.Cells(lin, 20).Value,
                U=an.Cells(lin, 21).Value,
                O=es.Cells(linEst, 15).Value if linEst else None,
                G=es.Cells(linEst, 7).Value if linEst else None,
                sig=es.Cells(linEst, 18).Value if linEst else None,
                pF=pa.Cells(7, 6).Value, pSig=pa.Cells(7, 9).Value,
                status=an.Range('T2').Value)

        print()
        print('=== 1. BASE: ano 2025 ===')
        an.Range('S2').Value = 2025
        tenta(lambda: xl.CalculateFull())
        b = leitura()
        print('   K=%r T(ETp)=%r U(CVTp)=%r | Estat O=%r G=%r Sigma=%r | Painel F=%r'
              % (b['K'], b['T'], b['U'], b['O'], b['G'], b['sig'], b['pF']))
        ck('status confirma 2025', 'ano 2025 localizado' in str(b['status']), str(b['status']))
        ck('Estatistica!O == Analitos!T', b['O'] == b['T'], '%r vs %r' % (b['O'], b['T']))
        ck('Painel!F7 == Analitos!T', b['pF'] == b['T'], '%r vs %r' % (b['pF'], b['T']))

        print()
        print('=== 2. MARCADOR no bloco de 2026 (CLIA ETp%% = %g) ===' % MARCADOR)
        origHist = an.Cells(57, colHist).Value      # linha "CLIA ETp%" do bloco 2026
        an.Cells(57, colHist).Value = MARCADOR
        an.Range('S2').Value = 2026
        tenta(lambda: xl.CalculateFull())
        d = leitura()
        print('   K=%r T(ETp)=%r U(CVTp)=%r | Estat O=%r G=%r Sigma=%r | Painel F=%r'
              % (d['K'], d['T'], d['U'], d['O'], d['G'], d['sig'], d['pF']))
        ck('status confirma 2026', 'ano 2026 localizado' in str(d['status']), str(d['status']))
        ck('Analitos!K seguiu o ano', d['K'] == MARCADOR, repr(d['K']))
        ck('Analitos!T (ETp em uso) seguiu', d['T'] == MARCADOR, repr(d['T']))
        ck('Estatistica!O seguiu', d['O'] == MARCADOR, repr(d['O']))
        ck('Painel!F7 seguiu', d['pF'] == MARCADOR, repr(d['pF']))
        ck('Sigma mudou com o ETp', d['sig'] != b['sig'], '%r -> %r' % (b['sig'], d['sig']))
        ck('CVTp tambem acompanhou (K/3)', abs((d['U'] or 0) - MARCADOR / 3.0) < 1e-9, repr(d['U']))

        print()
        print('=== 3. ANO NAO CADASTRADO ===')
        an.Range('S2').Value = 2099
        tenta(lambda: xl.CalculateFull())
        s = an.Range('T2').Value
        k = an.Cells(lin, 11).Value
        print('   status=%r  K=%r' % (s, k))
        ck('ano inexistente e DENUNCIADO', 'NAO CADASTRADO' in str(s).upper(), str(s))
        ck('sem #REF!/#VALOR! em K', not (isinstance(k, str) and k.startswith('#')), repr(k))

        print()
        print('=== 4. RESTAURA ===')
        an.Cells(57, colHist).Value = origHist
        an.Range('S2').Value = 2025
        tenta(lambda: xl.CalculateFull())
        v = leitura()
        ck('volta ao valor de 2025', v['T'] == b['T'], '%r vs %r' % (v['T'], b['T']))

        print()
        print('=== 5. INTEGRIDADE (Analitos, Estatistica, Painel) ===')
        # SpecialCells(formulas, 16) devolve SO as celulas com erro.
        #
        # Iterar UsedRange celula a celula percorre dezenas de milhares de
        # objetos COM e o Excel devolve RPC_E_CALL_REJECTED no meio -- o teste
        # morria sem terminar de conferir. Deixar o proprio Excel filtrar e
        # instantaneo e nao depende de sorte.
        ruins = {}
        for ws in (an, es, pa):
            try:
                erradas = tenta(lambda w=ws: w.UsedRange.SpecialCells(-4123, 16))
            except Exception:
                continue                      # nenhuma celula com erro
            for cel in erradas:
                t = str(tenta(lambda c=cel: c.Text))
                ruins[t] = ruins.get(t, 0) + 1
        print('   erros encontrados: %s' % (ruins if ruins else 'nenhum'))
        ck('zero #REF!', not any('#REF!' in k for k in ruins), str(ruins))
        ck('zero #NOME?', not any(k.startswith('#NOME') or k.startswith('#NAME')
                                  for k in ruins), str(ruins))

        wb.Save()
        tenta(lambda: wb.Close(SaveChanges=True))
        xl.Quit()
        time.sleep(3)

        print()
        print('=== 6. FECHAR E REABRIR ===')
        xl, wb = abrir(copia)
        an = wb.Worksheets('Analitos')
        es = wb.Worksheets('Estatística')
        print('   S2=%r  status=%r  K=%r  T=%r'
              % (an.Range('S2').Value, an.Range('T2').Value,
                 an.Cells(lin, 11).Value, an.Cells(lin, 20).Value))
        ck('ano preservado apos reabrir', an.Range('S2').Value == 2025,
           repr(an.Range('S2').Value))
        ck('ETp preservado apos reabrir', an.Cells(lin, 20).Value == b['T'],
           '%r vs %r' % (an.Cells(lin, 20).Value, b['T']))
        ck('Estatistica coerente apos reabrir',
           es.Cells(linEst, 15).Value == an.Cells(lin, 20).Value,
           '%r vs %r' % (es.Cells(linEst, 15).Value, an.Cells(lin, 20).Value))
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
    print('PROPAGACAO S2 -> A4:W43 -> ESTATISTICA/PAINEL: INTEGRA')


if __name__ == '__main__':
    main(sys.argv[1])
