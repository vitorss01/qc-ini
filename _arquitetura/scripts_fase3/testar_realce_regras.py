# -*- coding: utf-8 -*-
"""testar_realce_regras.py - ADR-037: o realce acende a regra certa?

Trabalha numa COPIA em TEMP.

NAO VEJO COR NA TELA. O que da para provar, e o que este teste prova:

  1. a DECISAO -- a matriz resolvida em Cfg_PlanoQC!J10:P10 diz, para cada
     Sigma, quais regras estao no plano;
  2. a LIGACAO -- cada celula de regra do Painel tem uma condicao que le
     essa matriz, com o formato pedido (verde escuro, fonte branca, negrito);
  3. a PRIORIDADE -- a condicao de violacao tem prioridade MAIOR que a de
     recomendacao, entao uma regra violada nao fica verde;
  4. a COERENCIA -- o realce e o texto de Regra_Westgard_Recomendada saem da
     mesma fonte e nunca divergem.

O que este teste NAO prova: que a cor renderizada na tela e agradavel. Isso
exige abrir o arquivo.

Uso: python testar_realce_regras.py <arquivo.xlsm>
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
CFG = 'Cfg_PlanoQC'
MAT_C0, MAT_N = 10, 5          # J..N
REGRAS_PAINEL = [('G6', '1-3S'), ('H6', '2-2S'), ('I6', 'R4S'),
                 ('J6', '4-1S'), ('K6', '8X')]

VERDE_ESCURO = 0x14 + 0x6C * 256 + 0x43 * 65536
BRANCO = 255 + 255 * 256 + 255 * 65536


def ck(nome, cond, det=''):
    print(('  OK   ' if cond else '  FALHA') + '  ' + nome +
          (('  -> ' + det) if det else ''))
    if not cond:
        falhas.append(nome)


def txt(v):
    return '' if v is None else str(v).strip()


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


def main(caminho):
    copia = os.path.join(os.environ.get('TEMP', '.'),
                         'trr_' + os.path.basename(caminho))
    shutil.copy(caminho, copia)
    xl = novo_excel()
    wb = xl.Workbooks.Open(copia)
    try:
        pa = wb.Worksheets('Painel')
        es = wb.Worksheets('Estatística')
        an = wb.Worksheets('Analitos')
        cfg = wb.Worksheets(CFG)
        cfg.Visible = -1
        xl.Calculation = -4105

        rotulos = [txt(tenta(lambda c=c: cfg.Cells(3, c).Value))
                   for c in range(MAT_C0, MAT_C0 + MAT_N)]

        def ativas():
            return [tenta(lambda c=c: cfg.Cells(10, c).Value)
                    for c in range(MAT_C0, MAT_C0 + MAT_N)]

        def acesas():
            return [r for r, v in zip(rotulos, ativas()) if v == 1]

        def poeSigma(s):
            for ref in ('V10', 'V11'):
                tenta(lambda r=ref, v=s: pa.Range(r).__setattr__('Value', v))
            tenta(lambda: xl.CalculateFull())

        # ================================================================
        print('=== A MATRIZ E DERIVADA DO TEXTO, NAO ESCRITA AO LADO ===')
        print('   rotulos da matriz: %s' % rotulos)
        for r in range(4, 9):
            f = txt(tenta(lambda x=r: cfg.Cells(x, MAT_C0).Formula))
            ck('linha %d: a bandeira LE a coluna Regras' % r,
               'SEARCH' in f and '$D' in f, f[:64])
        ck('a matriz tem exatamente as 5 regras do projeto',
           rotulos == ['1-3S', '2-2S', 'R4S', '4-1S', '8X'],
           str(rotulos))

        print()
        print('=== CADA CELULA DE REGRA TEM AS DUAS CONDICOES ===')
        for ref, rot in REGRAS_PAINEL:
            fcs = tenta(lambda r=ref: pa.Range(r).FormatConditions)
            n = tenta(lambda f=fcs: f.Count)
            if n < 2:
                ck('%s (%s): 2 condicoes' % (ref, rot), False, '%d' % n)
                continue
            f1 = tenta(lambda f=fcs: f.Item(1))
            f2 = tenta(lambda f=fcs: f.Item(2))
            form1 = txt(tenta(lambda x=f1: x.Formula1))
            form2 = txt(tenta(lambda x=f2: x.Formula1))
            p1 = tenta(lambda x=f1: x.Priority)
            p2 = tenta(lambda x=f2: x.Priority)
            viol = f1 if 'SUM(' in form1 else f2
            reco = f2 if 'SUM(' in form1 else f1
            pViol = p1 if 'SUM(' in form1 else p2
            pReco = p2 if 'SUM(' in form1 else p1
            ck('%s (%s): condicao de VIOLACAO le as contagens do nivel'
               % (ref, rot), 'SUM($%s$7:$%s$8)' % (ref[0], ref[0])
               in form1 + form2, (form1 + ' / ' + form2)[:70])
            ck('%s (%s): condicao de RECOMENDACAO le a matriz' % (ref, rot),
               'regrasAtivas' in form1 + form2 and 'regrasRotulos'
               in form1 + form2)
            ck('%s (%s): VIOLACAO tem prioridade sobre RECOMENDACAO'
               % (ref, rot), pViol < pReco, 'violacao=%s recomendacao=%s'
               % (pViol, pReco))
            ck('%s (%s): recomendada = verde escuro, fonte branca, negrito'
               % (ref, rot),
               int(tenta(lambda x=reco: x.Interior.Color)) == VERDE_ESCURO
               and int(tenta(lambda x=reco: x.Font.Color)) == BRANCO
               and tenta(lambda x=reco: x.Font.Bold) is True)
            ck('%s (%s): violada NAO usa a mesma cor da recomendada'
               % (ref, rot),
               int(tenta(lambda x=viol: x.Interior.Color)) != VERDE_ESCURO)

        print()
        print('=== TESTES 1 a 5: qual regra acende em cada Sigma ===')
        casos = [
            ('TESTE 1  Sigma 6,00', 6.00, ['1-3S']),
            ('TESTE 2  Sigma 5,50', 5.50, ['1-3S', '2-2S', 'R4S']),
            ('TESTE 3  Sigma 4,50', 4.50, ['1-3S', '2-2S', 'R4S', '4-1S']),
            ('TESTE 4  Sigma 3,50', 3.50, ['1-3S', '2-2S', 'R4S', '4-1S', '8X']),
            ('TESTE 5  Sigma 2,99', 2.99, ['1-3S', '2-2S', 'R4S', '4-1S', '8X']),
        ]
        for rot, sg, esperado in casos:
            poeSigma(sg)
            got = acesas()
            texto = txt(tenta(lambda s=sg: xl.Run('PlanoQC', s, 'REGRAS')))
            print('   %s -> acesas %s' % (rot, got))
            print('        texto do plano: %r' % (texto or '(faixa sem regras)'))
            ck('%s acende %s' % (rot, esperado), got == esperado, str(got))

        print()
        print('=== O REALCE E O TEXTO NUNCA DIVERGEM ===')
        for sg in (6.0, 5.5, 4.5, 3.5, 3.0):
            poeSigma(sg)
            got = set(acesas())
            texto = txt(tenta(lambda s=sg: xl.Run('PlanoQC', s, 'REGRAS')))
            # normaliza o rotulo exibido para o token do texto
            mapa = {'1-3S': '1_3s', '2-2S': '2_2s', 'R4S': 'R_4s',
                    '4-1S': '4_1s', '8X': '8x'}
            doTexto = set(k for k, v in mapa.items() if v in texto)
            ck('Sigma %.2f: acesas == texto do plano' % sg, got == doTexto,
               '%s vs %s' % (sorted(got), sorted(doTexto)))

        print()
        print('=== TESTE 5b: abaixo de 3 Sigma, estrategia intensiva e alerta ===')
        poeSigma(2.99)
        motor = txt(tenta(lambda: cfg.Range('D1').Value))
        status = txt(tenta(lambda: cfg.Cells(8, 9).Value))
        alerta = txt(tenta(lambda: pa.Range('F11').Value))
        print('   motor executa : %s' % motor)
        print('   Status_Plano_QC: %r' % status)
        print('   alerta no Painel: %r' % alerta[:90])
        ck('acendem exatamente as regras que o motor suporta',
           set(acesas()) == {'1-3S', '2-2S', 'R4S', '4-1S', '8X'},
           str(acesas()))
        ck('Status_Plano_QC da faixa < 3 = REAVALIAR MÉTODO',
           status == 'REAVALIAR MÉTODO', status)
        ck('o Painel mostra o alerta de desempenho inadequado',
           'INADEQUADO' in alerta.upper(), alerta[:60])
        ck('e NAO atribui run size', txt(tenta(
            lambda: xl.Run('PlanoQC', 2.99, 'RUNSIZE'))) == '')

        print()
        print('=== TESTE 6: violacao vence recomendacao ===')
        poeSigma(6.20)
        ck('em Sigma 6,20 a 1-3S esta no plano', '1-3S' in acesas(),
           str(acesas()))
        # forca uma violacao de 1_3s no nivel 1
        guardaG7 = tenta(lambda: pa.Range('G7').Formula)
        tenta(lambda: pa.Range('G7').__setattr__('Value', 2))
        tenta(lambda: xl.CalculateFull())
        fcs = tenta(lambda: pa.Range('G6').FormatConditions)
        f1 = tenta(lambda: fcs.Item(1))
        f2 = tenta(lambda: fcs.Item(2))
        viol = f1 if 'SUM(' in txt(tenta(lambda: f1.Formula1)) else f2
        reco = f2 if viol is f1 else f1
        # Range.Evaluate nao resolve nome definido nem referencia relativa do
        # mesmo jeito que a formatacao condicional resolve -- devolveu #NOME?.
        # A verdade de cada condicao se calcula direto da fonte que ela le.
        somaG = (tenta(lambda: pa.Range('G7').Value) or 0) +                 (tenta(lambda: pa.Range('G8').Value) or 0)
        vViol = somaG > 0
        vReco = '1-3S' in acesas()
        print('   com G7=2: SUM(G7:G8)=%s -> violacao=%s | 1-3S no plano=%s'
              % (somaG, vViol, vReco))
        ck('a condicao de violacao passou a valer', vViol is True, str(vViol))
        ck('a de recomendacao tambem seria verdadeira -- e por isso a '
           'prioridade importa', vReco is True, str(vReco))
        ck('e a de violacao tem prioridade menor (vence)',
           tenta(lambda: viol.Priority) < tenta(lambda: reco.Priority),
           '%s < %s' % (tenta(lambda: viol.Priority),
                        tenta(lambda: reco.Priority)))
        tenta(lambda: pa.Range('G7').__setattr__('Formula', guardaG7))

        print()
        print('=== TESTE DINAMICO: trocar de analito muda tudo sozinho ===')
        # devolve as formulas de Sigma
        for ref, cl in (('V10', 'L'), ('V11', 'L')):
            pass
        tenta(lambda: pa.Range('V10').__setattr__(
            'Formula',
            '=IFERROR(INDEX(Estatística!$L$14:$L$93,'
            'MATCH(selAnalito&"|"&1,Estatística!$AB$14:$AB$93,0)),"")'))
        tenta(lambda: pa.Range('V11').__setattr__(
            'Formula',
            '=IFERROR(INDEX(Estatística!$L$14:$L$93,'
            'MATCH(selAnalito&"|"&2,Estatística!$AB$14:$AB$93,0)),"")'))
        tenta(lambda: es.Range('L4').__setattr__('Value', 'CAP'))
        tenta(lambda: es.Range('N4').__setattr__('Value', 2025))
        tenta(lambda: es.Range('P4').__setattr__('Value', 'TODAS'))
        tenta(lambda: xl.CalculateFull())

        vistos = []
        print('   %-3s %-24s %-8s %-20s %-30s %-6s %-7s %-11s'
              % ('#', 'analito', 'sigma', 'classificação', 'regras acesas',
                 'N', 'run', 'DPM'))
        for idx in range(1, 32):
            tenta(lambda i=idx: pa.Range('B3').__setattr__('Value', i))
            tenta(lambda: xl.CalculateFull())
            nome = txt(tenta(lambda: pa.Range('C3').Value))
            sg = tenta(lambda: cfg.Range('B1').Value)
            if not isinstance(sg, (int, float)):
                continue
            cls = txt(tenta(lambda: pa.Range('W10').Value))
            ac = acesas()
            n = txt(tenta(lambda: pa.Range('U22').Text))
            run = txt(tenta(lambda: pa.Range('V22').Text))
            dpm = txt(tenta(lambda: pa.Range('X10').Text))
            faixa = ('>=6' if sg >= 6 else '5-6' if sg >= 5 else
                     '4-5' if sg >= 4 else '3-4' if sg >= 3 else '<3')
            if faixa not in [v[0] for v in vistos]:
                vistos.append((faixa, nome, sg, cls, ac, n, run, dpm))
                print('   %-3d %-24s %-8.3f %-20s %-30s %-6s %-7s %-11s'
                      % (idx, nome[:24], sg, cls[:20], '/'.join(ac)[:30],
                         n, run, dpm))
            if len(vistos) >= 5:
                break
        ck('foram vistas pelo menos 3 faixas diferentes de Sigma',
           len(vistos) >= 3, '%d faixas: %s' % (len(vistos),
                                                [v[0] for v in vistos]))
        # As faixas <3 e 3-4 acendem o MESMO conjunto de propósito: o plano
        # Marginal e 1_3s/2_2s/R_4s/4_1s/8x, e a estrategia intensiva abaixo
        # de 3 e o conjunto que o motor suporta -- os mesmos cinco. Exigir
        # conjuntos distintos reprovaria o comportamento correto.
        esperadoPorFaixa = {
            '>=6': ['1-3S'],
            '5-6': ['1-3S', '2-2S', 'R4S'],
            '4-5': ['1-3S', '2-2S', 'R4S', '4-1S'],
            '3-4': ['1-3S', '2-2S', 'R4S', '4-1S', '8X'],
            '<3': ['1-3S', '2-2S', 'R4S', '4-1S', '8X'],
        }
        for faixa, nome, sg, cls, ac, n, run, dpm in vistos:
            ck('%s (Sigma %.3f, %s) acende %s'
               % (faixa, sg, nome[:18], esperadoPorFaixa[faixa]),
               ac == esperadoPorFaixa[faixa], str(ac))
        ck('o conjunto aceso mudou entre pelo menos 3 faixas',
           len(set(tuple(v[4]) for v in vistos)) >= 3,
           str(sorted(set(tuple(v[4]) for v in vistos))))

        print()
        print('=== o Painel diz QUAL nivel governa o plano? ===')
        # O card mostra a classificacao do NIVEL 1; o realce usa o MENOR Sigma
        # dos dois niveis. Lactato exibia "Classe mundial" com as cinco regras
        # acesas -- cada metade certa, o conjunto contraditorio. O texto de
        # apoio precisa dizer qual Sigma manda.
        apoio = txt(tenta(lambda: pa.Range('F10').Value))
        print('   F10: %s' % apoio[:150])
        ck('o texto de apoio nomeia o nivel que governa',
           'PIOR nível' in apoio and ('nível 1' in apoio or 'nível 2' in apoio),
           apoio[:90])
        ck('e mostra o Sigma que esta mandando',
           'Sigma' in apoio and any(c.isdigit() for c in apoio), apoio[-60:])
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
    print('=' * 70)
    if falhas:
        print('FALHAS (%d):' % len(falhas))
        for f in falhas:
            print('   - %s' % f)
        sys.exit(1)
    print('REALCE DINAMICO DAS REGRAS: TODAS AS PROVAS PASSARAM')
    print('LIMITE: a cor renderizada na tela nao foi vista; foram provadas a '
          'decisao, a ligacao, a prioridade e a coerencia com o texto.')


if __name__ == '__main__':
    main(sys.argv[1])
