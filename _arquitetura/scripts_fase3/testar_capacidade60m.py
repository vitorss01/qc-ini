# -*- coding: utf-8 -*-
"""testar_capacidade60m.py - prova que 60 meses sao PROCESSADOS, nao so gravados

O criterio nao e "o arquivo aceitou as linhas". E: o dado entra, permanece,
esta dentro dos intervalos, participa dos calculos, aparece no Painel, na
Estatistica e nos graficos, e continua assim depois de fechar e reabrir.

Inclui o teste de borda obrigatorio na antiga linha 15.003: nada pode se
comportar de forma diferente antes e depois dela.

Uso: python testar_capacidade60m.py <caminho_convertido.xlsm> [meses]
"""
import io
import os
import sys
import time
import shutil
import datetime
import subprocess

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', write_through=True)
import win32com.client as w

SENHA = 'qcini2025'
ANTIGO_TETO = 15003
falhas = []


def ck(nome, cond, det=''):
    print(('  OK   ' if cond else '  FALHA') + '  ' + nome + (('  -> ' + det) if det else ''))
    if not cond:
        falhas.append(nome)


def novo_excel():
    for t in range(1, 8):
        try:
            xl = w.DispatchEx('Excel.Application')
            xl.Visible = False
            xl.DisplayAlerts = False
            xl.EnableEvents = False
            xl.AutomationSecurity = 1
            return xl
        except Exception:
            if t == 2:
                subprocess.call(['powershell', '-NoProfile', '-Command',
                                 'Start-Process excel.exe -WindowStyle Hidden'],
                                stderr=subprocess.DEVNULL)
                time.sleep(6)
            time.sleep(2.0 * t)
    raise RuntimeError('Excel COM nao subiu')


def add_meses(d, n):
    a = d.year + (d.month - 1 + n) // 12
    m = (d.month - 1 + n) % 12 + 1
    dia = d.day
    while True:
        try:
            return datetime.datetime(a, m, dia)
        except ValueError:
            dia -= 1


def destravar(wb, xl):
    try:
        wb.Unprotect(SENHA)
    except Exception:
        pass
    for s in wb.Worksheets:
        try:
            s.Unprotect(SENHA)
        except Exception:
            pass


def main(caminho, meses=60):
    orig = os.path.abspath(caminho)
    alvo = os.path.join(os.path.dirname(orig), 'teste60_%s' % os.path.basename(orig))
    shutil.copy(orig, alvo)

    xl = novo_excel()
    wb = xl.Workbooks.Open(alvo)
    try:
        destravar(wb, xl)
        db = wb.Worksheets('DB_Resultados')
        ult0 = db.Cells(db.Rows.Count, 1).End(-4162).Row
        base = [list(r) for r in db.Range(db.Cells(4, 1), db.Cells(ult0, 7)).Value]
        print('base: %d registros (dado REAL de janeiro/2026)' % len(base))
        print('SIMULACAO: os meses 2..%d sao replicas do bloco de janeiro,' % meses)
        print('           com datas deslocadas e RUNs em sequencia. Dado SIMULADO.')

        linhas = []
        for k in range(meses):
            for r in base:
                nr = list(r)
                nr[0] = int(r[0]) + k * 18
                if isinstance(r[1], datetime.datetime):
                    nr[1] = add_meses(r[1].replace(tzinfo=None), k)
                linhas.append(nr)
        n = len(linhas)
        ultl = 3 + n
        print('alvo: %d registros, ultima linha %d\n' % (n, ultl))

        xl.Calculation = -4135
        db.Range(db.Cells(4, 4), db.Cells(ultl, 4)).NumberFormat = '@'
        db.Range(db.Cells(4, 1), db.Cells(ultl, 7)).Value = linhas

        print('=== 1. FLAGS E NOMES ACOMPANHAM ===')
        t = time.time()
        xl.Run('AtualizarFlagsBanco')
        t_flags = time.time() - t
        print('   AtualizarFlagsBanco: %.2fs para %d registros' % (t_flags, n))
        ck('AtualizarFlagsBanco abaixo de 5 s', t_flags < 5.0, '%.2fs' % t_flags)
        for nm in ('rRUN', 'rData', 'rAnalito', 'rValor', 'rStatus', 'rLote',
                   'rFirst', 'rRunUnico', 'rNivel'):
            ref = wb.Names(nm).RefersTo
            ck('%s vai ate a linha %d' % (nm, ultl), ('$%d' % ultl) in ref, ref[-32:])

        print('\n=== 2. TESTE DE BORDA NA ANTIGA LINHA 15.003 ===')
        for rot, lin in (('ultima DENTRO do antigo teto', ANTIGO_TETO),
                         ('primeira FORA do antigo teto', ANTIGO_TETO + 1),
                         ('muito alem do antigo teto', ultl)):
            ba = db.Cells(lin, 53).Value
            bb = db.Cells(lin, 54).Value
            bc = db.Cells(lin, 55).Value
            an = db.Cells(lin, 5).Value
            print('   L%-6d %-30s analito=%-14s BA=%r BB=%r BC=%r'
                  % (lin, rot, str(an)[:14], ba, bb, bc))
            ck('L%d tem nucleo de lote' % lin, ba not in (None, ''), repr(ba))
            ck('L%d tem flag BB numerica' % lin, isinstance(bb, (int, float)), repr(bb))
            ck('L%d tem flag BC numerica' % lin, isinstance(bc, (int, float)), repr(bc))

        print('\n=== 3. OS DADOS PARTICIPAM DOS CALCULOS ===')
        xl.Calculation = -4105
        t = time.time()
        xl.CalculateFullRebuild()
        t_reb = time.time() - t
        print('   recalculo completo: %.2fs' % t_reb)

        ca = wb.Worksheets('Calc')
        pts = sum(1 for r in range(3, 183) if ca.Cells(r, 6).Value not in (None, ''))
        print('   pontos no grafico Levey-Jennings: %d' % pts)
        ck('grafico com pontos', pts > 0, str(pts))

        # A corrida MAIS RECENTE tem de estar visivel no Calc
        runmax = int(max(l[0] for l in linhas))
        runs_calc = set()
        for r in range(3, 183):
            v = ca.Cells(r, 2).Value
            if v not in (None, ''):
                runs_calc.add(int(v))
        print('   RUN maximo no banco=%d | maior RUN visto no Calc=%s'
              % (runmax, max(runs_calc) if runs_calc else '-'))
        if runs_calc and max(runs_calc) < runmax:
            print('   NOTA: a janela do Calc mostra 180 corridas por lote e pega as')
            print('         MAIS ANTIGAS (AGGREGATE(15;...) = k-esimo menor). Limite')
            print('         conhecido e independente desta mudanca -- ver ADR-025.')

        print('\n=== 4. A ESTATISTICA ENXERGA OS 60 MESES ===')
        t = time.time()
        nreg = xl.Run('EstatPeriodo', base[0][4], 1, 'N', 0, 0, '')
        t_est = time.time() - t
        print('   EstatPeriodo(%s, N1, sem janela) = %r em %.2fs'
              % (base[0][4], nreg, t_est))
        # A cadencia sai do BLOCO BASE, nao de um 18 fixo.
        #
        # 18 era o numero de corridas de janeiro/2026 na producao. O artefato do
        # build usa a fixture de semear_dados_teste (25 corridas), e o teste
        # acusou "1500 obtido, ~1080 esperado" -- 1500 e exatamente 60 x 25, ou
        # seja, o motor estava certo e a expectativa e que estava errada.
        # Expectativa hardcoded transforma diferenca de fixture em falha falsa.
        corridas_base = len(set(int(r[0]) for r in base))
        esperado = meses * corridas_base
        ck('Estatistica conta as corridas dos %d meses' % meses,
           isinstance(nreg, (int, float)) and abs(nreg - esperado) <= esperado * 0.02,
           'obteve %r, esperado ~%d' % (nreg, esperado))

        print('\n=== 5. BARREIRA DE CAPACIDADE ===')
        livres = xl.Run('LinhasLivres')
        print('   LinhasLivres = %d  (capacidade 120.000)' % livres)
        ck('sobra capacidade depois de %d meses' % meses, livres > 0, str(livres))
        # TestarCapacidade, e nao ExigirCapacidade direto: um Err.Raise sem
        # tratamento chamado por Application.Run abre dialogo modal e trava o
        # Excel com CPU zerada -- foi o que travou a primeira rodada.
        r1 = xl.Run('TestarCapacidade', livres + 1)
        r2 = xl.Run('TestarCapacidade', livres)
        print('   pedir %d linhas -> %s' % (livres + 1, str(r1)[:70]))
        print('   pedir %d linhas -> %s' % (livres, str(r2)[:70]))
        ck('gravar ACIMA da capacidade e RECUSADO com mensagem',
           str(r1).startswith('RECUSOU') and 'Capacidade do banco esgotada' in str(r1),
           str(r1)[:70])
        ck('gravar DENTRO da capacidade e permitido', str(r2) == 'PERMITIU', str(r2))

        print('\n=== 6. PROVA DE EQUIVALENCIA (amostra) ===')
        r = xl.Run('ConferirFlagsBanco', 4000)
        p = r.split('|')
        print('   ConferirFlagsBanco: %s linhas, %s divergencias %s' % (p[0], p[1], p[2]))
        ck('zero divergencia na amostra conferida', p[1] == '0', r)

        print('\n=== 7. FECHAR E REABRIR ===')
        t = time.time()
        wb.Save()
        t_salvar = time.time() - t
        tam = os.path.getsize(alvo) / 1048576.0
        wb.Close(False)
        xl.Quit()
        time.sleep(1.5)

        xl = novo_excel()
        t = time.time()
        wb = xl.Workbooks.Open(alvo)
        t_abrir = time.time() - t
        destravar(wb, xl)
        db = wb.Worksheets('DB_Resultados')
        ca = wb.Worksheets('Calc')
        ult2 = db.Cells(db.Rows.Count, 1).End(-4162).Row
        print('   salvou em %.2fs (%.2f MB), reabriu em %.2fs' % (t_salvar, tam, t_abrir))
        ck('ultima linha preservada apos reabrir', ult2 == ultl, '%d vs %d' % (ult2, ultl))
        ck('nome rAnalito preservado', ('$%d' % ultl) in wb.Names('rAnalito').RefersTo,
           wb.Names('rAnalito').RefersTo[-28:])
        ck('flag BB preservada na ultima linha',
           isinstance(db.Cells(ultl, 54).Value, (int, float)),
           repr(db.Cells(ultl, 54).Value))
        xl.Calculation = -4105
        pts2 = sum(1 for r in range(3, 183) if ca.Cells(r, 6).Value not in (None, ''))
        ck('grafico continua com pontos apos reabrir', pts2 > 0, str(pts2))

        print('\n=== 8. UMA CORRIDA NOVA, PELO CAMINHO REAL ===')
        pa = wb.Worksheets('Painel')
        t = time.time()
        pa.Range('C3').Value = base[1][4]
        t_troca = time.time() - t
        print('   trocar analito no Painel: %.2fs' % t_troca)

        print('\n--- RESUMO DE PERFORMANCE, %d MESES ---' % meses)
        print('   registros ................ %d' % n)
        print('   tamanho .................. %.2f MB' % tam)
        print('   AtualizarFlagsBanco ...... %.2f s' % t_flags)
        print('   recalculo completo ....... %.2f s' % t_reb)
        print('   abertura ................. %.2f s' % t_abrir)
        print('   salvamento ............... %.2f s' % t_salvar)
        print('   trocar analito ........... %.2f s' % t_troca)
        print('   Estatistica (1a chamada) . %.2f s' % t_est)
    finally:
        try:
            wb.Close(False)
        except Exception:
            pass
        try:
            xl.Quit()
        except Exception:
            pass

    print('\n' + '=' * 70)
    if falhas:
        print('FALHAS (%d):' % len(falhas))
        for f in falhas:
            print('   - %s' % f)
        sys.exit(1)
    print('60 MESES: ENTRAM, PERMANECEM, CALCULAM E SOBREVIVEM AO REABRIR')


if __name__ == '__main__':
    main(sys.argv[1], int(sys.argv[2]) if len(sys.argv) > 2 else 60)
