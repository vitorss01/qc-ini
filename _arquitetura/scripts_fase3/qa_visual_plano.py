# -*- coding: utf-8 -*-
"""qa_visual_plano.py - ADR-035: o que a tela MOSTRA

Le o TEXTO RENDERIZADO de cada celula, e nao o valor que a formula devolve. E a
unica forma de pegar "####", "0" no lugar de vazio, texto cortado e rotulo que
some -- defeitos que so existem na renderizacao.

LIMITE DECLARADO: este script nao ve a tela. Ele confere largura, texto exibido,
sobreposicao de objetos e destino de hyperlink. Julgamento estetico -- se o
bloco "parece" bom -- exige abrir o arquivo, e isso nao e feito aqui.

Uso: python qa_visual_plano.py <arquivo.xlsm>
"""
import io
import os
import sys
import time
import subprocess

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', write_through=True)
import win32com.client as w

problemas = []


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
            if t in (1, 3, 5):
                subprocess.call(['powershell', '-NoProfile', '-Command',
                                 'Start-Process excel.exe -WindowStyle Minimized'],
                                stderr=subprocess.DEVNULL)
                time.sleep(10)
            time.sleep(2.5 * t)
    raise RuntimeError('Excel COM nao subiu')


TRANS = ('rejeitada', 'rejected', 'membro n', 'member not found', 'busy')


def tenta(fn, vezes=10):
    ult = None
    for i in range(vezes):
        try:
            return fn()
        except Exception as e:
            ult = e
            if not any(t in str(e).lower() for t in TRANS):
                raise
            time.sleep(1.0 + 0.8 * i)
    raise ult


def col(n):
    s = ''
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def main(caminho):
    xl = novo_excel()
    wb = xl.Workbooks.Open(os.path.abspath(caminho))
    try:
        es = wb.Worksheets('Estatística')
        pa = wb.Worksheets('Painel')
        xl.Calculation = -4105
        tenta(lambda: pa.Range('B3').__setattr__('Value', 4))
        tenta(lambda: es.Range('L4').__setattr__('Value', 'CAP'))
        tenta(lambda: es.Range('N4').__setattr__('Value', 2025))
        tenta(lambda: es.Range('P4').__setattr__('Value', 'TODAS'))
        tenta(lambda: xl.CalculateFull())

        print('=== SELETORES DA ESTATISTICA ===')
        for ref, rot in (('L4', 'Provedor'), ('N4', 'Ano'), ('P4', 'Rodada')):
            v = tenta(lambda r=ref: es.Range(r).Text)
            try:
                lista = tenta(lambda r=ref: es.Range(r).Validation.Formula1)
            except Exception:
                lista = '(sem validação)'
            print('   %-10s = %-12s  lista: %s' % (rot, v, str(lista)[:70]))
            if str(lista).startswith('(sem'):
                problemas.append('seletor %s sem lista de validação' % rot)
        print('   eco do filtro: %s' % tenta(lambda: es.Range('K5').Text))
        print('   status Westgard: M7=%r  M8=%r'
              % (tenta(lambda: pa.Range('M7').Text),
                 tenta(lambda: pa.Range('M8').Text)))
        print('   status Westgard: M7=%r  M8=%r'
              % (tenta(lambda: pa.Range('M7').Text),
                 tenta(lambda: pa.Range('M8').Text)))

        print()
        print('=== PAINEL: analito %s ==='
              % tenta(lambda: pa.Range('C3').Text))
        def bloco(titulo):
            for r in range(1, 200):
                for c in range(1, 34):
                    v = tenta(lambda x=r, y=c: pa.Cells(x, y).Value)
                    if v and titulo.lower() in str(v).lower():
                        return r, c
            problemas.append('bloco %r nao encontrado no Painel' % titulo)
            return None

        blocos = []
        for tit, alt, c1 in (('DESEMPENHO SIX SIGMA', 3, 17),
                             ('PLANO DE CQ RECOMENDADO', 3, 16),
                             ('ERRO TOTAL vs ETp', 3, 15),
                             ('MARGEM CRÍTICA', 4, 12),
                             ('SIGMA × DPM', 11, 12)):
            achado = bloco(tit)
            if achado:
                r0, c0 = achado
                blocos.append((tit, r0, r0 + alt + 2, c0, c0 + 8))
        for nome, r0, r1, c0, c1 in blocos:
            print('   --- %s ---' % nome)
            for r in range(r0, r1 + 1):
                celulas = []
                for c in range(c0, c1 + 1):
                    t = str(tenta(lambda x=r, y=c: pa.Cells(x, y).Text))
                    if t:
                        celulas.append(t[:22])
                if celulas:
                    print('      %s' % ' | '.join(celulas))

        print()
        print('=== BASE CIENTIFICA / HYPERLINKS ===')
        n = tenta(lambda: pa.Hyperlinks.Count)
        print('   hyperlinks no Painel: %d' % n)
        for i in range(1, n + 1):
            h = tenta(lambda k=i: pa.Hyperlinks(k))
            print('      %-6s %-58s -> %s'
                  % (h.Range.Address, str(h.TextToDisplay)[:58],
                     str(h.Address)[:60]))
            if not str(h.Address).startswith('http'):
                problemas.append('hyperlink sem destino http em %s'
                                 % h.Range.Address)
        if n < 4:
            problemas.append('esperados 4 hyperlinks, encontrados %d' % n)

        print()
        print('=== NOTAS OBRIGATORIAS PRESENTES? ===')
        alvos = [('Painel', pa, 'short-term Sigma', 'nota do 1,5 SD'),
                 ('Painel', pa, 'run size é recomendação', 'nota de run size'),
                 ('Painel', pa, 'número TOTAL de medições', 'nota do N'),
                 ('Painel', pa, 'R_4s', 'nota run size vs R_4s'),
                 ('Estatística', es, 'short-term Sigma', 'nota do 1,5 SD')]
        for nomeAba, ws, trecho, rot in alvos:
            achou = False
            for r in range(1, 150):
                for c in range(1, 20):
                    t = str(tenta(lambda x=r, y=c: ws.Cells(x, y).Value) or '')
                    if trecho in t:
                        achou = True
                        break
                if achou:
                    break
            print('   %-26s em %-12s: %s' % (rot, nomeAba,
                                             'sim' if achou else 'NAO'))
            if not achou:
                problemas.append('%s ausente em %s' % (rot, nomeAba))

        print()
        print('=== CELULAS EXIBINDO #### (coluna estreita) ===')
        cortadas = []
        for ws, nome, r1, c1 in ((pa, 'Painel', 95, 18),
                                 (es, 'Estatística', 140, 28)):
            for r in range(1, r1 + 1):
                for c in range(1, c1 + 1):
                    t = str(tenta(lambda a=ws, x=r, y=c: a.Cells(x, y).Text))
                    if t and set(t) == {'#'}:
                        cortadas.append('%s!%s%d' % (nome, col(c), r))
        print('   %s' % (cortadas if cortadas else 'nenhuma'))
        problemas.extend(cortadas)

        print()
        print('=== OBJETOS SOBREPOSTOS AOS BLOCOS NOVOS ===')
        for sh in list(pa.Shapes):
            l, t2, w2, h2 = sh.Left, sh.Top, sh.Width, sh.Height
            rB, cB = bloco('DESEMPENHO SIX SIGMA')
            topoBloco = tenta(lambda: pa.Rows(rB).Top)
            colJ = tenta(lambda: pa.Columns(cB).Left)
            invade = (t2 + h2 > topoBloco) and (l + w2 > colJ)
            print('   %-22s L=%.0f T=%.0f %.0fx%.0f  invade o bloco? %s'
                  % (str(sh.Name)[:22], l, t2, w2, h2, 'SIM' if invade else 'nao'))
            if invade:
                problemas.append('objeto %s sobre o primeiro bloco' % sh.Name)

        print()
        print('=== TROCA DE ANALITO ATUALIZA O CARD? ===')
        antes = []
        for idx in (1, 4, 9):
            tenta(lambda i=idx: pa.Range('B3').__setattr__('Value', i))
            tenta(lambda: xl.CalculateFull())
            nome = str(tenta(lambda: pa.Range('C3').Text))
            sg = str(tenta(lambda: pa.Range('V10').Text))
            dp = str(tenta(lambda: pa.Range('X10').Text))
            rg = str(tenta(lambda: pa.Range('T22').Text))
            print('   analito %-2d %-24s sigma=%-8s dpm=%-10s regras=%s'
                  % (idx, nome[:24], sg, dp, rg[:28]))
            antes.append((nome, sg, dp))
        if len(set(x[0] for x in antes)) != 3:
            problemas.append('trocar o analito nao mudou o nome exibido')

        print()
        print('=== TROCA DE PROVEDOR ATUALIZA? ===')
        for prov in ('CAP', 'Controllab', 'CAP'):
            tenta(lambda p=prov: es.Range('L4').__setattr__('Value', p))
            tenta(lambda: xl.CalculateFull())
            print('   provedor=%-11s Estatistica G14=%-10s Painel L13=%s'
                  % (prov, str(tenta(lambda: es.Cells(14, 7).Text))[:10],
                     str(tenta(lambda: pa.Range('T10').Text))[:10]))
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
    print('=' * 68)
    if problemas:
        print('PROBLEMAS DE RENDERIZACAO (%d):' % len(problemas))
        for p in problemas:
            print('   - %s' % p)
        sys.exit(1)
    print('QA DE RENDERIZACAO: SEM PROBLEMAS')
    print('LIMITE: julgamento estetico exige abrir o arquivo; nao foi feito aqui.')


if __name__ == '__main__':
    main(sys.argv[1])
