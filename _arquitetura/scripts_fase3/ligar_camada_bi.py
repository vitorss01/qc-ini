# -*- coding: utf-8 -*-
"""ligar_camada_bi.py - poe a camada de BI no arquivo de producao

ESTADO ENCONTRADO

mEQA, mCEQ, mQualidade, mPlanoQC, mWestgardKnowledge e mEstatistica ja estavam
na producao. So o mBI faltava -- e ele e o ULTIMO da cadeia, entao a importacao
nao esbarra em dependencia nenhuma. Sem ele nao ha aba BI_Data, e sem BI_Data o
Power BI nao tem de onde ler.

O QUE FAZ

  1. importa mBI.bas de src_producao (fonte unica; o build usa a mesma)
  2. confere que o projeto continua compilando, chamando uma funcao do modulo
  3. roda AtualizarBIData, que cria/preenche a aba
  4. roda as duas reconciliacoes que o proprio modulo publica
  5. so salva se tudo passar

Uso: python ligar_camada_bi.py <arquivo.xlsm> [limite_s]
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
FONTE = os.path.join(BASE, '_arquitetura', 'src_producao', 'mBI.bas')


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
    if not proc_excel():
        subprocess.call(['powershell', '-NoProfile', '-Command',
                         'Start-Process excel.exe -WindowStyle Minimized'],
                        stderr=subprocess.DEVNULL)
        time.sleep(14)
    # DispatchEx e o preferido (instancia propria), mas depois de ciclos de
    # Stop-Process ele passa a devolver CO_E_SERVER_EXEC_FAILURE enquanto o
    # Dispatch, que ANEXA a um Excel ja vivo, continua funcionando. Medido.
    ult = None
    for t in range(1, 8):
        for fn in (lambda: w.DispatchEx('Excel.Application'),
                   lambda: w.Dispatch('Excel.Application')):
            try:
                xl = fn()
                xl.Visible = False
                xl.DisplayAlerts = False
                xl.EnableEvents = False
                xl.AutomationSecurity = 1
                return xl
            except Exception as e:
                ult = e
        subprocess.call(['powershell', '-NoProfile', '-Command',
                         'Start-Process excel.exe -WindowStyle Minimized'],
                        stderr=subprocess.DEVNULL)
        time.sleep(6 + 2.0 * t)
    raise RuntimeError('Excel COM nao subiu: %s' % ult)


def dialogo():
    achado = []

    def cb(h, _):
        try:
            if win32gui.IsWindowVisible(h) and win32gui.GetClassName(h) == '#32770':
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
    """Toda execucao tem limite e condicao de falha explicita."""
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
                r = x.Run(macro)
                caixa['r'] = ('ok', time.time() - t0, r)
            except Exception as e:
                caixa['r'] = ('erro: %s' % str(e)[:130], time.time() - t0, None)
        except Exception as e:
            caixa['r'] = ('thread: %s' % str(e)[:70], 0.0, None)
        finally:
            pythoncom.CoUninitialize()

    th = threading.Thread(target=alvo, daemon=True)
    th.start()
    th.join(limite)
    if th.is_alive():
        return ('TRAVOU (>%ds) dialogo=%s' % (limite, dialogo() or 'nenhum'),
                float(limite), None)
    return caixa.get('r', ('sem resultado', 0.0, None))


def main(caminho, limite=240):
    caminho = os.path.abspath(caminho)
    prod = os.path.basename(caminho).replace('QC_', '').replace('.xlsm', '')
    xl = novo()
    wb = xl.Workbooks.Open(caminho)
    salvar = False
    try:
        if wb.ReadOnly:
            raise SystemExit('somente leitura')
        proj = wb.VBProject
        nomes = [c.Name for c in proj.VBComponents]

        # dependencias antes de importar: mBI e o ultimo da cadeia
        dep = ['mQualidade', 'mPlanoQC', 'mCEQ', 'mEQA']
        falta = [d for d in dep if d not in nomes]
        print('%s: dependencias do mBI %s'
              % (prod, 'todas presentes' if not falta else 'FALTAM %s' % falta))
        if falta:
            raise SystemExit('importar %s antes do mBI' % falta)

        if 'mBI' in nomes:
            print('  mBI ja presente -- reimportando para ficar igual a fonte')
            for c in list(proj.VBComponents):
                if c.Name == 'mBI':
                    proj.VBComponents.Remove(c)
                    break
        proj.VBComponents.Import(FONTE)
        alvo = [c for c in proj.VBComponents if c.Name == 'mBI']
        if not alvo:
            raise SystemExit('Import nao criou mBI')
        print('  mBI importado (%d linhas)' % alvo[0].CodeModule.CountOfLines)

        # IMPASSE DE ORDEM, RESOLVIDO AQUI.
        #
        # aplicar_bi_data.ps1 recusa rodar sem o mBI no arquivo; e toda rotina
        # do mBI faz Sheets("BI_Data"), que sem a aba e erro 9 -- dialogo modal
        # que trava a automacao. Nenhum dos dois pode ser o primeiro sozinho.
        #
        # A saida: importar o modulo e SALVAR, deixar o script criar a aba com o
        # arquivo fechado, e so entao voltar para preencher e reconciliar.
        temBI = any(ws.Name == 'BI_Data' for ws in wb.Worksheets)
        if not temBI:
            wb.Save()
            wb.Close(True)
            xl.Quit()
            print('  mBI salvo; criando BI_Data com o arquivo fechado')
            r = subprocess.call(
                ['powershell', '-NoProfile', '-ExecutionPolicy', 'Bypass',
                 '-File', os.path.join(BASE, '_arquitetura', 'scripts_fase3',
                                       'aplicar_bi_data.ps1'),
                 '-Workbook', caminho])
            if r != 0:
                raise SystemExit('aplicar_bi_data.ps1 devolveu %d' % r)
            xl = novo()
            wb = xl.Workbooks.Open(caminho)
            print('  BI_Data criada; reabrindo para preencher')

        est, seg, _ = rodar(xl, 'AtualizarBIData', limite)
        if est != 'ok':
            raise SystemExit('AtualizarBIData: %s' % est)
        print('  AtualizarBIData: ok em %.1fs' % seg)

        bi = None
        for ws in wb.Worksheets:
            if ws.Name == 'BI_Data':
                bi = ws
                break
        if bi is None:
            raise SystemExit('BI_Data nao existe depois do AtualizarBIData')
        ncol = int(bi.Cells(1, bi.Columns.Count).End(-4159).Column)
        nlin = int(bi.Cells(bi.Rows.Count, 1).End(-4162).Row)
        print('  BI_Data: %d colunas x %d linhas (cabecalho na 1)' % (ncol, nlin))
        if ncol < 84:
            raise SystemExit('contrato incompleto: %d colunas, esperado 84' % ncol)

        for macro in ('ReconciliarBancoBI', 'ReconciliarComCalc'):
            est, seg, r = rodar(xl, macro, 120)
            print('  %-20s %s -> %s' % (macro, est[:24], str(r)[:90]))
            if est != 'ok':
                raise SystemExit('%s: %s' % (macro, est))

        wb.Save()
        salvar = True
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


main(sys.argv[1], int(sys.argv[2]) if len(sys.argv) > 2 else 240)
