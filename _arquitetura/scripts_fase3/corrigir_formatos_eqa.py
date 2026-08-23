# -*- coding: utf-8 -*-
"""corrigir_formatos_eqa.py - ADR-034: casas decimais, sem passar pelo locale

O PROBLEMA

Range.NumberFormat deveria receber o codigo INVARIANTE ("0.00"), mas nesta
instalacao ele se comporta como NumberFormatLocal: devolve "Geral" em vez de
"General". Num Excel pt-BR, o ponto e separador de MILHAR e a virgula e
decimal -- entao "0.00" foi lido como "milhar, sem decimais", e o Excel
normalizou para "#.000". Resultado: as colunas numericas nasceram com o numero
de casas errado.

O QUE ESTE SCRIPT FAZ

Mede primeiro (imprime o que a celula EXIBE, que e o unico fato que importa
para o gestor), aplica os codigos em NumberFormatLocal com virgula decimal, e
mede de novo. Se o depois nao for melhor que o antes, nao salva.

Uso: python corrigir_formatos_eqa.py <arquivo.xlsm>
"""
import io
import os
import sys
import time
import subprocess

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', write_through=True)
import win32com.client as w

# coluna -> (rotulo, codigo pt-BR, casas decimais esperadas)
FORMATOS = [
    (3,  'Ano',              '0',                 0),
    (6,  'Seu Resultado',    '0,00',              2),
    (7,  'Média/Valor-alvo', '0,000',             3),
    (8,  'SD',               '0,000',             3),
    (9,  'SDI',              '+0,0;-0,0;0,0',     1),
    (10, 'Limite Inferior',  '0,00',              2),
    (11, 'Limite Superior',  '0,00',              2),
    (13, 'Nº Laboratórios',  '0',                 0),
    (15, 'Bias (%)',         '0,00',              2),
    (16, '|Bias| (%)',       '0,00',              2),
    (17, 'Página Fonte',     '0',                 0),
]

ABAS_DADOS = ['EQA.CAP_Dados', 'EQA.Controllab_Dados']
ABAS_DERIV = ['EQA.CAP_Nao_Aceitaveis', 'EQA.Controllab_Desvios']


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


def tenta(fn, vezes=10):
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


def casas(txt):
    """Quantas casas decimais o texto exibido tem."""
    t = str(txt)
    if ',' in t:
        return len(t.split(',')[-1])
    return 0


def amostrar(ws, linhas):
    """O que as celulas EXIBEM hoje -- e nao o codigo de formato guardado."""
    saida = {}
    for c, rot, _cod, _esp in FORMATOS:
        vistos = []
        for r in linhas:
            v = tenta(lambda rr=r, cc=c: ws.Cells(rr, cc).Value)
            if isinstance(v, (int, float)):
                vistos.append(str(tenta(lambda rr=r, cc=c: ws.Cells(rr, cc).Text)))
            if len(vistos) >= 3:
                break
        saida[c] = vistos
    return saida


def main(caminho):
    caminho = os.path.abspath(caminho)
    xl = novo_excel()
    wb = xl.Workbooks.Open(caminho)
    salvou = False
    if wb.ReadOnly:
        wb.Close(False)
        xl.Quit()
        raise SystemExit('Somente leitura: %s' % caminho)
    try:
        xl.Calculation = -4135
        cap = wb.Worksheets('EQA.CAP_Dados')
        linhas = list(range(2, 40))

        print('=== ANTES: o que a EQA.CAP_Dados exibe ===')
        antes = amostrar(cap, linhas)
        errado = 0
        for c, rot, cod, esp in FORMATOS:
            ex = antes[c]
            ok = all(casas(t) == esp for t in ex) if ex else True
            if not ok:
                errado += 1
            print('   %-18s esperado %d casa(s) | exibindo %s %s'
                  % (rot, esp, ex, '' if ok else '<- ERRADO'))
        print('colunas com casas decimais erradas: %d' % errado)

        # ---- aplicar, agora por NumberFormatLocal -----------------------
        print()
        for nome in ABAS_DADOS:
            ws = wb.Worksheets(nome)
            for c, rot, cod, esp in FORMATOS:
                tenta(lambda w2=ws, cc=c, k=cod:
                      w2.Range(w2.Cells(2, cc), w2.Cells(3000, cc))
                      .__setattr__('NumberFormatLocal', k))
            print('%s: %d colunas reformatadas' % (nome, len(FORMATOS)))
        # as derivadas espelham as mesmas colunas, deslocadas para a linha 5
        for nome in ABAS_DERIV:
            ws = wb.Worksheets(nome)
            for c, rot, cod, esp in FORMATOS:
                tenta(lambda w2=ws, cc=c, k=cod:
                      w2.Range(w2.Cells(5, cc), w2.Cells(210, cc))
                      .__setattr__('NumberFormatLocal', k))
            print('%s: %d colunas reformatadas' % (nome, len(FORMATOS)))
        # e a base consolidada, que o Power BI le
        base = wb.Worksheets('EQA_Base')
        visAntes = base.Visible
        base.Visible = -1
        for c, cod in ((7, '0,00'), (8, '0,000'), (9, '0,000'),
                       (10, '+0,0;-0,0;0,0'), (11, '0,00'), (12, '0,00'),
                       (16, '0,00'), (17, '0,00'), (18, '0'), (2, '0')):
            tenta(lambda cc=c, k=cod:
                  base.Range(base.Cells(2, cc), base.Cells(5001, cc))
                  .__setattr__('NumberFormatLocal', k))
        base.Visible = visAntes
        print('EQA_Base: 10 colunas reformatadas')

        xl.Calculation = -4105
        tenta(lambda: xl.CalculateFullRebuild())

        print()
        print('=== DEPOIS ===')
        depois = amostrar(cap, linhas)
        errado2 = 0
        for c, rot, cod, esp in FORMATOS:
            ex = depois[c]
            ok = all(casas(t) == esp for t in ex) if ex else True
            if not ok:
                errado2 += 1
            print('   %-18s esperado %d casa(s) | exibindo %s %s'
                  % (rot, esp, ex, '' if ok else '<- AINDA ERRADO'))
        print('colunas com casas decimais erradas: %d (era %d)' % (errado2, errado))

        # nenhuma coluna pode ter ficado estreita demais e virar ####
        cortadas = []
        for c, rot, cod, esp in FORMATOS:
            for r in (2, 3, 200, 400):
                t = str(tenta(lambda cc=c, rr=r: cap.Cells(rr, cc).Text))
                if t and set(t) == {'#'}:
                    cortadas.append('%s (linha %d)' % (rot, r))
        print('celulas exibindo ####: %s' % (cortadas if cortadas else 'nenhuma'))

        if errado2 > errado or cortadas:
            raise SystemExit('piorou ou apareceu #### -- nada salvo')

        wb.Save()
        salvou = True
        print('SALVO: %s' % caminho)
    finally:
        try:
            wb.Close(salvou)
        except Exception:
            pass
        try:
            xl.Quit()
        except Exception:
            pass


if __name__ == '__main__':
    main(sys.argv[1])
