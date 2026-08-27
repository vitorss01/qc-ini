# -*- coding: utf-8 -*-
"""testar_alvo_do_lote.py - o teto do LotesStore vem do dado, nao do numero 200

O DEFEITO (ADR-045)

mEstatPeriodo.AlvoDoLote varria "For i = 2 To 200". O LotesStore guarda 40
linhas por lote (mBI.LS_CAP), entao 200 cobre quatro blocos e meio: do 5o lote
em diante o alvo caia fora da varredura e a funcao devolvia VAZIO -- sem Z, sem
bias e sem erro.

A correcao existia em src_producao desde o ADR-045 e nunca chegou ao arquivo,
porque instalar_estat_periodo.ps1 nao e chamado pelo build_all.

COMO SE PROVA ALCANCE AQUI -- e por que as tentativas obvias dao falso negativo

AlvoDoLote NAO procura o analito pelo nome no LotesStore. Resolve o nome para um
INDICE na aba Analitos (linhas 4..43) e depois procura a linha cujo lote e o
ativo E cujo idx bate. Sai no PRIMEIRO casamento.

  - analito inventado    -> nao existe em Analitos, a funcao sai antes de varrer
  - linha nova la embaixo -> a linha antiga casa primeiro e responde por ela
  - linha isolada em 240  -> o vao antes dela dispara o Exit For do branco

Entao: neutraliza-se o casamento antigo mudando SO a coluna idx daquela linha
(a coluna 1 continua preenchida, entao nao abre buraco), preenche-se o bloco de
forma CONTIGUA ate 240 -- que e a forma do dado real, blocos de 40 linhas -- e
poe-se o par lote+idx la no fim. Se a resposta for o valor escrito em 240, a
varredura chegou. Depois desfaz tudo e confere que voltou ao de antes.

Nada e salvo: o teste fecha o arquivo sem gravar.
"""
import io
import os
import sys
import time
import subprocess

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', write_through=True)
import win32com.client as w

TRANS = ('rejeitada', 'rejected', 'membro n', 'member not found', 'busy')


def tenta(fn, vezes=10):
    u = None
    for i in range(vezes):
        try:
            return fn()
        except Exception as e:
            u = e
            if not any(t in str(e).lower() for t in TRANS):
                raise
            time.sleep(1.0 + 0.8 * i)
    raise u


def novo():
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


def main(caminho):
    caminho = os.path.abspath(caminho)
    xl = novo()
    wb = xl.Workbooks.Open(caminho)
    ok = fail = 0

    def reg(nome, esp, obt, cond):
        nonlocal ok, fail
        if cond:
            ok += 1
        else:
            fail += 1
        print('  %-5s %-44s esp %-20s obt %s'
              % ('OK' if cond else 'FALHA', nome[:44], str(esp)[:20], str(obt)[:32]))

    try:
        wsL = wb.Worksheets('LotesStore')
        wsA = wb.Worksheets('Analitos')

        ult = int(tenta(lambda: wsL.Cells(wsL.Rows.Count, 1).End(-4162).Row))
        ANALITO = str(tenta(lambda: wsA.Cells(4, 1).Value) or '').strip()
        IDX_ALVO = 1
        LIN = max(ult + 2, 240)
        ALVO = 4.321

        print('ultima linha do LotesStore hoje : %d' % ult)
        print('analito de indice 1 em Analitos : %r' % ANALITO)

        loteAtivo = str(tenta(lambda: xl.Run('LoteAtivoCore')) or '').strip()
        print('lote ativo                      : %r' % loteAtivo)
        if not loteAtivo:
            raise SystemExit('sem lote ativo: nao da para provar o alcance')

        antiga = 0
        for r_ in range(2, ult + 1):
            l_ = str(tenta(lambda x=r_: wsL.Cells(x, 1).Value) or '').strip()
            i_ = tenta(lambda x=r_: wsL.Cells(x, 2).Value)
            try:
                bate = int(float(i_)) == IDX_ALVO
            except (TypeError, ValueError):
                bate = False
            if l_ == loteAtivo and bate:
                antiga = r_
                break
        if not antiga:
            raise SystemExit('nao achei a linha do lote ativo com idx %d' % IDX_ALVO)

        antes = tenta(lambda: xl.Run('AlvoDoLote', ANALITO, 1))
        print('linha %d responde hoje          : %r' % (antiga, antes))
        print()

        idxAntigo = tenta(lambda: wsL.Cells(antiga, 2).Value)
        tenta(lambda: wsL.Cells(antiga, 2).__setattr__('Value', 999))

        ini = ult + 1
        col1 = [[loteAtivo] for _ in range(LIN - ini + 1)]
        tenta(lambda: wsL.Range(wsL.Cells(ini, 1),
                                wsL.Cells(LIN, 1)).__setattr__('Value', col1))
        tenta(lambda: wsL.Cells(LIN, 2).__setattr__('Value', IDX_ALVO))
        tenta(lambda: wsL.Cells(LIN, 3).__setattr__('Value', ALVO))
        conf = int(tenta(lambda: wsL.Cells(wsL.Rows.Count, 1).End(-4162).Row))
        print('bloco contiguo de %d ate %d; ultima linha agora: %d'
              % (ini, LIN, conf))
        print()

        r = tenta(lambda: xl.Run('AlvoDoLote', ANALITO, 1))
        chegou = isinstance(r, (int, float)) and abs(float(r) - ALVO) < 1e-9
        reg('varredura alcanca a linha %d' % LIN, ALVO, repr(r), chegou)

        # provas estruturais. Sem caixa: o VBE reescreve a caixa dos
        # identificadores ao importar -- Application.Run volta como
        # Application.run, do mesmo jeito que Str$ vira stR$ e ws.Rows vira
        # ws.rows. Comparacao sensivel a caixa reprova codigo correto.
        mod = wb.VBProject.VBComponents('mEstatPeriodo').CodeModule
        tl = mod.Lines(1, mod.CountOfLines).lower()
        reg('sem teto fixo "To 200" na varredura', 'ausente',
            'presente' if 'to 200' in tl else 'ausente', 'to 200' not in tl)
        reg('teto derivado de End(xlUp)', 'presente',
            'presente' if 'end(xlup).row' in tl else 'ausente',
            'end(xlup).row' in tl)
        reg('LimEspec usa vinculo tardio', 'Application.Run',
            'presente' if 'application.run("especcvtp"' in tl else 'ausente',
            'application.run("especcvtp"' in tl)

        # desfaz tudo
        tenta(lambda: wsL.Range(wsL.Cells(ini, 1),
                                wsL.Cells(LIN, 3)).ClearContents())
        tenta(lambda: wsL.Cells(antiga, 2).__setattr__('Value', idxAntigo))
        depois = tenta(lambda: xl.Run('AlvoDoLote', ANALITO, 1))
        reg('desfeito: volta a responder o de antes', repr(antes), repr(depois),
            depois == antes)

        # e a Estatistica continua sem celula em erro
        est = None
        for ws in wb.Worksheets:
            if ws.Name.lower().startswith('estat'):
                est = ws
                break
        tenta(lambda: xl.CalculateFullRebuild())
        ruins = []
        for r_ in range(1, 100):
            for c in range(1, 32):
                t = str(tenta(lambda x=r_, y=c: est.Cells(x, y).Text))
                if t.startswith('#') and t.strip('#') != '':
                    ruins.append('%s%d=%s' % (chr(64 + c), r_, t))
        reg('Estatistica sem celula em erro', '0',
            '%d %s' % (len(ruins), ruins[:3]), not ruins)
    finally:
        try:
            wb.Close(False)   # nunca salva: e teste
        except Exception:
            pass
        try:
            xl.Quit()
        except Exception:
            pass

    print()
    print('%d OK, %d FALHA' % (ok, fail))
    sys.exit(1 if fail else 0)


main(sys.argv[1])
