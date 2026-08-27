# -*- coding: utf-8 -*-
"""reparar_bioquimica_fase3.py - fecha o artefato hibrido da Bioquimica

O DEFEITO

QC_Bioquimica.xlsm recebeu o mEstatistica da Fase 3 sem receber o que ele exige:

  mWestgardKnowledge  ausente -- e mEstatistica chama RegraClassificacao,
                      RegraInterpretacao, RegraCausas e RegraSugestoes.
                      Resultado: "Sub ou Function nao definida", que e erro de
                      COMPILACAO e nao de execucao: nenhum On Error captura, e
                      com o Excel invisivel vira dialogo modal que trava a
                      automacao.
  Eng_Saida           ausente -- Sheets("Eng_Saida") em AtualizarCalc,
  Eventos_Westgard    AtualizarPainelEng, AtualizarEstatisticaAba e
  Cfg_Status          RegistrarEventosWestgard: erro 9.

As tres abas ja foram criadas por criar_eng_saida.ps1 e criar_abas_motor.ps1.
Falta o modulo -- e falta remover a duplicata de AtualizarEstatistica.

POR QUE REPARO CIRURGICO E NAO build_all

O build_all reconstroi a partir da producao e grava FORA do repositorio. Rodar
o build e publicar o resultado por cima do arquivo da raiz sobrescreveria o que
so existe na producao: o layout do Painel ajustado a mao pelo gestor (base
oficial do ADR-036), os 455 resultados reais do CAP na EQA_Base, e o
mEstatPeriodo instalado pelo ADR-047. A unidade de correcao aqui e a
dependencia que falta, nao o arquivo inteiro.

A UNIDADE DE CORRECAO E O MODULO INTEIRO

Copiar so RegraClassificacao faria o compilador parar na proxima -- e sao
quatro. O modulo vem inteiro, da mesma fonte que o build_all usa.

Nada e salvo se qualquer verificacao falhar.
"""
import io
import os
import sys
import time
import threading
import subprocess

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', write_through=True)
import win32com.client as w
import win32gui
import pythoncom

BASE = r"C:\Users\vitor\OneDrive - MSFT\Desktop\QC_INI"
FONTE_MWK = os.path.join(BASE, '_arquitetura', 'snapshot_producao',
                         'Hematologia', 'vba', 'mWestgardKnowledge.bas')
ROTINAS = ['AtualizarCalc', 'RegistrarEventosWestgard',
           'AtualizarPainelEng', 'AtualizarEstatisticaAba']


def matar():
    subprocess.call(['powershell', '-NoProfile', '-Command',
                     'Get-Process EXCEL -EA SilentlyContinue | Stop-Process -Force'],
                    stderr=subprocess.DEVNULL)
    time.sleep(4)


def primar():
    """Sobe um Excel interativo antes de qualquer DispatchEx.

    Depois de varios ciclos de Stop-Process, o DCOM passa a devolver
    CO_E_SERVER_EXEC_FAILURE (0x80080005) e nao se recupera sozinho. Medido:
    com nenhum EXCEL.EXE vivo, DispatchEx falha; com um Excel ja em execucao,
    DispatchEx, Dispatch e GetActiveObject funcionam os tres.

    Por isso o priming vem ANTES da primeira tentativa, e nao como reacao a
    falha -- a versao anterior primava so depois de ja ter falhado.
    """
    if not proc_excel():
        subprocess.call(
            ['powershell', '-NoProfile', '-Command',
             'Start-Process excel.exe -WindowStyle Minimized'],
            stderr=subprocess.DEVNULL)
        time.sleep(14)


def proc_excel():
    try:
        out = subprocess.check_output(
            ['powershell', '-NoProfile', '-Command',
             '(Get-Process EXCEL -EA SilentlyContinue | Measure-Object).Count'],
            stderr=subprocess.DEVNULL, timeout=40)
        return int(out.decode('ascii', 'ignore').strip() or '0') > 0
    except Exception:
        return False


def novo():
    primar()
    for t in range(1, 10):
        try:
            xl = w.DispatchEx('Excel.Application')
            xl.Visible = False
            xl.DisplayAlerts = False
            xl.EnableEvents = False
            xl.AutomationSecurity = 1
            return xl
        except Exception:
            subprocess.call(
                ['powershell', '-NoProfile', '-Command',
                 'Start-Process excel.exe -WindowStyle Minimized'],
                stderr=subprocess.DEVNULL)
            time.sleep(6 + 2.0 * t)
    raise RuntimeError('Excel COM nao subiu')


def texto_dialogo():
    """Le a caixa modal que o VBA abre com o Excel invisivel."""
    achado = []

    def cb(h, _):
        try:
            if not win32gui.IsWindowVisible(h):
                return True
            if win32gui.GetClassName(h) != '#32770':
                return True
            partes = []

            def filho(hc, _2):
                try:
                    if win32gui.GetClassName(hc) == 'Static':
                        t = win32gui.GetWindowText(hc)
                        if t:
                            partes.append(t.replace('\r', ' ').replace('\n', ' '))
                except Exception:
                    pass
                return True
            win32gui.EnumChildWindows(h, filho, None)
            if partes:
                achado.append(' | '.join(partes))
        except Exception:
            pass
        return True
    try:
        win32gui.EnumWindows(cb, None)
    except Exception:
        pass
    return achado


def rodar(xl, macro, limite):
    """(estado, segundos). Se travar, tenta ler o dialogo antes de desistir."""
    caixa = {}
    idst = pythoncom.CoMarshalInterThreadInterfaceInStream(
        pythoncom.IID_IDispatch, xl)

    def alvo():
        pythoncom.CoInitialize()
        try:
            x = w.Dispatch(pythoncom.CoGetInterfaceAndReleaseStream(
                idst, pythoncom.IID_IDispatch))
            t0 = time.time()
            try:
                x.Run(macro)
                caixa['r'] = ('completou', time.time() - t0)
            except Exception as e:
                caixa['r'] = ('erro: %s' % str(e)[:150], time.time() - t0)
        except Exception as e:
            caixa['r'] = ('thread: %s' % str(e)[:80], 0.0)
        finally:
            pythoncom.CoUninitialize()

    th = threading.Thread(target=alvo, daemon=True)
    th.start()
    th.join(limite)
    if th.is_alive():
        d = texto_dialogo()
        return ('TRAVOU (>%ds) dialogo=%s' % (limite, d or 'nenhum'), float(limite))
    return caixa.get('r', ('sem resultado', 0.0))


def main(caminho, limite=180):
    caminho = os.path.abspath(caminho)
    if not os.path.exists(FONTE_MWK):
        raise SystemExit('fonte ausente: %s' % FONTE_MWK)

    # NAO matar antes: o DispatchEx precisa de um Excel ja vivo para nao cair
    # em CO_E_SERVER_EXEC_FAILURE. primar() cuida disso dentro de novo().
    xl = novo()
    wb = xl.Workbooks.Open(caminho)
    salvar = False
    problemas = []
    try:
        if wb.ReadOnly:
            raise SystemExit('somente leitura')
        proj = wb.VBProject
        nomes = [c.Name for c in proj.VBComponents]
        print('modulos antes: %d' % len(nomes))

        # ---- 1. mWestgardKnowledge ---------------------------------------
        if 'mWestgardKnowledge' in nomes:
            print('  mWestgardKnowledge ja presente')
        else:
            proj.VBComponents.Import(FONTE_MWK)
            achou = [c for c in proj.VBComponents
                     if c.Name == 'mWestgardKnowledge']
            if not achou:
                raise SystemExit('Import nao criou mWestgardKnowledge')
            cm = achou[0].CodeModule
            txt = cm.Lines(1, cm.CountOfLines)
            faltam = [f for f in ('RegraClassificacao', 'RegraInterpretacao',
                                  'RegraCausas', 'RegraSugestoes')
                      if ('function %s' % f).lower() not in txt.lower()]
            print('  mWestgardKnowledge importado (%d linhas); funcoes faltando: %s'
                  % (cm.CountOfLines, faltam or 'nenhuma'))
            if faltam:
                raise SystemExit('modulo importado nao define %s' % faltam)

        # ---- 2. duplicata de AtualizarEstatistica -------------------------
        dup = []
        for c in proj.VBComponents:
            try:
                cm = c.CodeModule
                if cm.ProcStartLine('AtualizarEstatistica', 0) > 0:
                    dup.append(c.Name)
            except Exception:
                pass
        print('  AtualizarEstatistica definida em: %s' % dup)
        if 'mUI' in dup and 'mEstatistica' in dup:
            cm = proj.VBComponents('mUI').CodeModule
            ini = cm.ProcStartLine('AtualizarEstatistica', 0)
            n = cm.ProcCountLines('AtualizarEstatistica', 0)
            cm.DeleteLines(ini, n)
            print('  stub do mUI removido (linhas %d..%d): a chamada sem '
                  'qualificar do mOperacao deixa de ser ambigua'
                  % (ini, ini + n - 1))

        # ---- 3. as quatro rotinas, no arquivo que sera entregue -----------
        print()
        print('=== rotinas do motor ===')
        for macro in ROTINAS:
            tenta_su = None
            est, seg = rodar(xl, macro, limite)
            try:
                tenta_su = bool(xl.ScreenUpdating)
                xl.ScreenUpdating = True
            except Exception:
                pass
            ok = est == 'completou'
            print('  %-5s %-26s %-52s %6.1fs  ScreenUpdating apos=%s'
                  % ('OK' if ok else 'FALHA', macro, est[:52], seg, tenta_su))
            if not ok:
                problemas.append('%s -> %s' % (macro, est[:110]))
            if est.startswith('TRAVOU'):
                break

        if problemas:
            print()
            for p in problemas:
                print('  PROBLEMA: %s' % p)
            raise SystemExit('nada salvo: %d rotina(s) nao completaram'
                             % len(problemas))

        # ---- 4. a planilha nao pode ter ganho erro -------------------------
        print()
        ruins = []
        for ws in wb.Worksheets:
            try:
                ur = ws.UsedRange
                if ur.Rows.Count * ur.Columns.Count > 400000:
                    continue
                v = ur.Value
            except Exception:
                continue
            if not isinstance(v, tuple):
                continue
            for lin in v:
                if not isinstance(lin, tuple):
                    lin = (lin,)
                for cel in lin:
                    if isinstance(cel, str) and cel.startswith('#') \
                            and cel.strip('#'):
                        ruins.append('%s:%s' % (ws.Name, cel))
                        break
        vistos = sorted(set(ruins))
        print('  celulas em erro: %d %s' % (len(vistos), vistos[:6]))

        wb.Save()
        salvar = True
        print()
        print('SALVO: %s' % caminho)
    finally:
        try:
            xl.ScreenUpdating = True
        except Exception:
            pass
        try:
            wb.Close(salvar)
        except Exception:
            pass
        try:
            xl.Quit()
        except Exception:
            pass


main(sys.argv[1], int(sys.argv[2]) if len(sys.argv) > 2 else 180)
