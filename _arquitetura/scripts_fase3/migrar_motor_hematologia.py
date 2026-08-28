# -*- coding: utf-8 -*-
"""migrar_motor_hematologia.py - a Hematologia passa ao grao de evento

O QUE MUDA

A producao rodava o mEstatistica antigo (1617 linhas): sem Eng_Saida, sem
MetricasWestgard, e com um RegistrarEventosWestgard que grava
(corrida x analito x nivel). Nesse formato R_4s, 2of3_2s, 3_1s e 6x sao
ESTRUTURALMENTE INVISIVEIS -- todas exigem ver os niveis juntos.

O motor da linhagem de build (1965 linhas) agrupa a serie por ANALITO com todos
os niveis juntos e grava o grao de evento em 14 campos, com detector, escopo,
N, R, corrida inicial e evidencia.

O QUE PRECISA ESTAR NO LUGAR ANTES

  Eng_Saida     criada por criar_eng_saida.ps1 -NLV 3
  Cfg_Status    ja existia
  Eventos_Westgard  ja existia
  LiberarEscrita / RestaurarProtecao  NAO existiam na producao -- o motor novo
                usa o par do ADR-046, e sem ele nao compila

A guarda de producao usa a senha em LITERAL, nao a constante SENHA_PROT do
trecho de referencia: essa constante nao existe no mSeguranca de producao, e
importar o trecho como esta derrubaria o projeto inteiro.

Nada e salvo se qualquer etapa falhar.
"""
import io
import os
import re
import sys
import time
import threading
import subprocess

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', write_through=True)
import win32com.client as w
import win32gui
import win32con
import pythoncom

BASE = r"C:\Users\vitor\OneDrive - MSFT\Desktop\QC_INI"
# PRODUTO VEM POR ARGUMENTO: os dois estavam no grao legado.
#
# A Bioquimica gravava as regras CONCATENADAS numa celula ("13s+22s+41s+10x") --
# o defeito que o ADR-045 descreve -- e ainda com 10x, aposentado no ADR-041.
PRODUTO = sys.argv[1] if len(sys.argv) > 1 else 'Hematologia'
ALVO = os.path.join(BASE, 'QC_%s.xlsm' % PRODUTO)
MOTOR = os.path.join(BASE, '_arquitetura', 'src_hardening1', PRODUTO,
                     'mEstatistica.bas')

GUARDA = '''

' ============== ESCRITA EM ABA PROTEGIDA (ADR-046) ==============
'
' ReprotectAll aplica UserInterfaceOnly:=True, que e a configuracao certa, mas
' o Excel NAO PERSISTE essa flag ao salvar. Reaberto o arquivo, a aba volta
' protegida tambem para o VBA ate Workbook_Open rodar LockApp -- e Workbook_Open
' nao roda em automacao, porque todo script abre com EnableEvents = False.
'
' Por isso a garantia nao pode ser "alguem rodou LockApp antes": cada rotina que
' escreve trata a propria janela.
'
' SENHA EM LITERAL, e nao a constante SENHA_PROT do trecho de referencia: essa
' constante nao existe neste modulo, e uma referencia a nome inexistente e erro
' de COMPILACAO -- que derruba o projeto inteiro, nao so a rotina.
Public Function LiberarEscrita(ByVal ws As Worksheet) As Boolean
    LiberarEscrita = ws.ProtectContents
    If LiberarEscrita Then ws.Unprotect Password:="qcini2025"
End Function


' Nao propaga erro: e chamada no caminho de limpeza, onde mascarar a excecao
' original seria pior do que falhar em reproteger.
Public Sub RestaurarProtecao(ByVal ws As Worksheet, ByVal estava As Boolean)
    If Not estava Then Exit Sub
    On Error Resume Next
    ws.Protect Password:="qcini2025", UserInterfaceOnly:=True, _
               DrawingObjects:=False, Contents:=True, Scenarios:=True
End Sub
'''


def proc_excel():
    try:
        return int(subprocess.check_output(
            ['powershell', '-NoProfile', '-Command',
             '(Get-Process EXCEL -EA SilentlyContinue | Measure-Object).Count'],
            stderr=subprocess.DEVNULL, timeout=40).decode('ascii', 'ignore').strip() or '0') > 0
    except Exception:
        return False


def novo():
    if not proc_excel():
        subprocess.call(['powershell', '-NoProfile', '-Command',
                         'Start-Process excel.exe -WindowStyle Minimized'],
                        stderr=subprocess.DEVNULL)
        time.sleep(14)
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


def caixa():
    achado = []

    def cb(h, _):
        try:
            if not win32gui.IsWindowVisible(h) or win32gui.GetClassName(h) != '#32770':
                return True
            if 'Visual Basic' not in win32gui.GetWindowText(h):
                return True
            ps = []

            def f(hc, _2):
                try:
                    if win32gui.GetClassName(hc) == 'Static':
                        t = win32gui.GetWindowText(hc)
                        if t:
                            ps.append(t.replace('\r', ' ').replace('\n', ' '))
                except Exception:
                    pass
                return True
            win32gui.EnumChildWindows(h, f, None)
            if ps:
                achado.append((h, ' | '.join(ps)))
        except Exception:
            pass
        return True
    try:
        win32gui.EnumWindows(cb, None)
    except Exception:
        pass
    return achado[0] if achado else (None, None)


def dispensar(h):
    alvo = []

    def f(hc, _):
        try:
            if win32gui.GetClassName(hc) == 'Button':
                t = win32gui.GetWindowText(hc).replace('&', '').strip().lower()
                if t in ('ok', 'fim', 'end'):
                    alvo.append(hc)
        except Exception:
            pass
        return True
    win32gui.EnumChildWindows(h, f, None)
    win32gui.PostMessage(alvo[0] if alvo else h,
                         win32con.BM_CLICK if alvo else win32con.WM_CLOSE, 0, 0)


def rodar(xl, macro, limite=300):
    """Roda com limite. Se abrir caixa modal, LE e dispensa em vez de travar."""
    idst = pythoncom.CoMarshalInterThreadInterfaceInStream(
        pythoncom.IID_IDispatch, xl)
    fim = {}

    def alvo():
        pythoncom.CoInitialize()
        try:
            x = w.Dispatch(pythoncom.CoGetInterfaceAndReleaseStream(
                idst, pythoncom.IID_IDispatch))
            t0 = time.time()
            try:
                x.Run(macro)
                fim['r'] = ('ok', time.time() - t0)
            except Exception as e:
                fim['r'] = ('erro: %s' % str(e)[:110], time.time() - t0)
        except Exception as e:
            fim['r'] = ('thread: %s' % str(e)[:70], 0.0)
        finally:
            pythoncom.CoUninitialize()

    th = threading.Thread(target=alvo, daemon=True)
    th.start()
    for _ in range(limite * 2):
        if not th.is_alive():
            break
        time.sleep(0.5)
        h, txt = caixa()
        if h:
            # O VBE SELECIONA o identificador ofensor. Ler ANTES de dispensar:
            # depois do clique a selecao se perde e sobra so "Variavel nao
            # definida", que nao diz qual.
            onde = ''
            try:
                pane = xl.VBE.ActiveCodePane
                comp = pane.CodeModule.Parent.Name
                sel = pane.GetSelection()
                l1, c1, c2 = int(sel[0]), int(sel[1]), int(sel[3])
                linha = pane.CodeModule.Lines(l1, 1)
                onde = ' [%s l.%d: %r em %r]' % (
                    comp, l1, linha[c1 - 1:c2 - 1], linha.strip()[:70])
            except Exception as e:
                onde = ' [selecao ilegivel: %s]' % str(e)[:50]
            dispensar(h)
            th.join(8)
            return ('DIALOGO: %s%s' % (txt, onde), 0.0)
    th.join(3)
    return fim.get('r', ('TRAVOU', float(limite)))


def main():
    xl = novo()
    wb = xl.Workbooks.Open(ALVO)
    salvar = False
    try:
        proj = wb.VBProject

        # ---- 1. a guarda do ADR-046 -------------------------------------
        cm = proj.VBComponents('mSeguranca').CodeModule
        txt = cm.Lines(1, cm.CountOfLines)
        if 'Function LiberarEscrita' in txt:
            print('mSeguranca: guarda ja presente')
        else:
            cm.AddFromString(GUARDA)
            print('mSeguranca: LiberarEscrita/RestaurarProtecao acrescentadas')

        # ---- 1b. mAuditoria ---------------------------------------------
        #
        # O motor novo chama Auditar CAT_SIS quando o buffer de eventos estoura.
        # CAT_SIS e a rotina Auditar vivem no mAuditoria, que NAO existia na
        # Hematologia -- e referencia a nome inexistente e erro de COMPILACAO,
        # que derruba o projeto inteiro e nao so o caminho de erro.
        #
        # O VBE apontou o nome exato (mEstatistica l.1786), o que evitou
        # descobrir as dependencias uma por uma, a cada abertura do Excel.
        AUD = os.path.join(BASE, '_arquitetura', 'src_hardening1', 'mAuditoria.bas')
        tem_aud = any(c.Name == 'mAuditoria' for c in proj.VBComponents)
        if tem_aud:
            print('mAuditoria: ja presente')
        else:
            proj.VBComponents.Import(AUD)
            achou = [c for c in proj.VBComponents if c.Name == 'mAuditoria']
            if not achou:
                raise SystemExit('Import nao criou mAuditoria')
            print('mAuditoria importado (%d linhas)'
                  % achou[0].CodeModule.CountOfLines)

        # ---- 2. o motor -------------------------------------------------
        antes = None
        for c in list(proj.VBComponents):
            if c.Name == 'mEstatistica':
                antes = c.CodeModule.CountOfLines
                proj.VBComponents.Remove(c)
                break
        print('mEstatistica anterior: %s linhas -- removido' % antes)
        proj.VBComponents.Import(MOTOR)
        novo_mod = [c for c in proj.VBComponents if c.Name == 'mEstatistica']
        if not novo_mod:
            raise SystemExit('Import nao criou mEstatistica')
        n = novo_mod[0].CodeModule.CountOfLines
        print('mEstatistica novo: %d linhas' % n)

        # 'aS' e a palavra reservada 'As': se veio de novo, sai de novo
        t2 = novo_mod[0].CodeModule.Lines(1, n)
        k = len(re.findall(r'(?<![A-Za-z0-9_])aS(?![A-Za-z0-9_])', t2))
        if k:
            novo_mod[0].CodeModule.DeleteLines(1, n)
            novo_mod[0].CodeModule.AddFromString(
                re.sub(r'(?<![A-Za-z0-9_])aS(?![A-Za-z0-9_])', 'alvoS', t2))
            print('  aS -> alvoS (%d) no modulo importado' % k)

        # ---- 3. o motor tem de rodar ------------------------------------
        print()
        for macro in ('AtualizarCalc', 'RegistrarEventosWestgard',
                      'AtualizarPainelEng', 'AtualizarEstatisticaAba'):
            est, seg = rodar(xl, macro)
            print('  %-5s %-26s %-46s %6.1fs'
                  % ('OK' if est == 'ok' else 'FALHA', macro, str(est)[:46], seg))
            if est != 'ok':
                raise SystemExit('%s: %s' % (macro, est))

        # ---- 4. o grao mudou? --------------------------------------------
        ev = wb.Worksheets('Eventos_Westgard')
        cab = [str(ev.Cells(3, c).Value or '') for c in range(1, 16)]
        cab = [c for c in cab if c]
        ult = int(ev.Cells(ev.Rows.Count, 1).End(-4162).Row)
        print()
        print('  Eventos_Westgard: %d colunas, %d linhas de evento'
              % (len(cab), max(0, ult - 3)))
        print('    %s' % ' | '.join(cab))
        # O CABECALHO NAO E A PROVA -- O DADO E.
        #
        # RegistrarEventosWestgard so reescreve da linha 4 para baixo, entao o
        # rotulo da linha 3 continua o antigo mesmo com o motor novo gravando 14
        # colunas. Medir o cabecalho reprovaria uma migracao que deu certo.
        # A largura do DADO e que diz qual motor escreveu.
        larg = 0
        for c in range(1, 15):
            if ev.Cells(4, c).Value not in (None, ''):
                larg = c
        print('    largura do dado na linha 4: %d coluna(s)' % larg)
        if larg < 12:
            raise SystemExit('ainda no grao legado (dado com %d colunas)' % larg)

        regras = set()
        for r in range(4, min(ult, 4000) + 1):
            v = ev.Cells(r, 5).Value
            if v:
                regras.add(str(v).strip())
        print('    regras registradas: %s' % sorted(regras))

        wb.Save()
        salvar = True
        print()
        print('SALVO: %s' % ALVO)
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


main()
