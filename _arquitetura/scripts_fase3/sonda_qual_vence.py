# -*- coding: utf-8 -*-
"""sonda_qual_vence.py - qual AtualizarEstatistica o mOperacao realmente chama

NAO SE DEDUZ, MEDE-SE.

Ha duas Public Sub AtualizarEstatistica na Bioquimica:

  mUI          4 linhas: On Error Resume Next + Sheets("Estatistica").Calculate
  mEstatistica 10 linhas: InvalidarCache, AtualizarCalc, RegistrarEventosWestgard,
               AtualizarPainelEng, AtualizarEstatisticaAba, AtualizarEixos

mOperacao.AtualizarOperacao -- "cadeia unica chamada apos qualquer gravacao/
exclusao" -- chama sem qualificar, de um TERCEIRO modulo. Se casar no stub, o
motor de Westgard nao roda depois de lancar resultado.

DISCRIMINADOR: so a versao do mEstatistica chama AtualizarEixos, que escreve
MinimumScale/MaximumScale nos graficos. Desregula-se o eixo de um grafico e
roda-se AtualizarOperacao:

  escala volta ao valor de antes -> venceu mEstatistica (o motor rodou)
  escala continua desregulada    -> venceu o stub do mUI

O tempo de execucao entra como corroboracao: o stub e instantaneo, o motor nao.

O arquivo e fechado SEM salvar.
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
    try:
        # onde estao as duas copias
        nomes = []
        for c in wb.VBProject.VBComponents:
            try:
                cm = c.CodeModule
                txt = cm.Lines(1, cm.CountOfLines)
            except Exception:
                continue
            if 'sub atualizarestatistica(' in txt.lower():
                nomes.append(c.Name)
        print('modulos que definem AtualizarEstatistica: %s' % nomes)
        if len(nomes) < 2:
            print('menos de duas copias: nao ha ambiguidade a medir neste produto')

        pa = wb.Worksheets('Painel')
        ch = None
        for o in pa.ChartObjects():
            ch = o
            break
        if ch is None:
            raise SystemExit('Painel sem grafico: sem discriminador')
        eixo = ch.Chart.Axes(2)   # xlValue
        antes = float(tenta(lambda: eixo.MinimumScale))
        print('grafico %r  MinimumScale antes: %r' % (ch.Name, antes))

        DESREG = antes - 12345.0
        tenta(lambda: eixo.__setattr__('MinimumScale', DESREG))
        print('desregulado para: %r' % float(tenta(lambda: eixo.MinimumScale)))
        print()

        print('chamando AtualizarOperacao ...')
        t0 = time.time()
        try:
            tenta(lambda: xl.Run('AtualizarOperacao'), vezes=2)
            erro = None
        except Exception as e:
            erro = str(e)[:180]
        dt = time.time() - t0
        print('   levou %.2fs   erro: %s' % (dt, erro or 'nenhum'))

        depois = float(tenta(lambda: ch.Chart.Axes(2).MinimumScale))
        print('MinimumScale depois: %r' % depois)
        print()

        voltou = abs(depois - DESREG) > 1e-9
        if voltou:
            print('VEREDITO: venceu mEstatistica.AtualizarEstatistica')
            print('          o eixo foi reescrito, entao AtualizarEixos rodou.')
        else:
            print('VEREDITO: venceu o STUB do mUI')
            print('          o eixo continuou desregulado: AtualizarEixos NAO rodou,')
            print('          logo o motor tambem nao. Depois de gravar resultado, a')
            print('          cadeia apenas recalcula a aba Estatistica.')
        print('          (tempo %.2fs -- o stub e instantaneo, o motor nao)' % dt)

        tenta(lambda: ch.Chart.Axes(2).__setattr__('MinimumScale', antes))
    finally:
        try:
            wb.Close(False)
        except Exception:
            pass
        try:
            xl.Quit()
        except Exception:
            pass


main(sys.argv[1])
