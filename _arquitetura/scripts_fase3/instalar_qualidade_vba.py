# -*- coding: utf-8 -*-
"""instalar_qualidade_vba.py - ADR-033: uma escada so, para os tres consumidores

O QUE FAZ

  1. importa mQualidade no ARQUIVO DE TRABALHO (so ele -- ver MODULOS abaixo)
  2. troca os IF encadeados da Estatistica!M/P e do Painel!O por chamadas a
     mQualidade.ClassificarSigma e mQualidade.ClassificarMargem
  3. prova que compilou: chama as funcoes por xl.Run antes de salvar
  4. cronometra o recalculo, para o custo da UDF nao passar despercebido

POR QUE TROCAR IF ENCADEADO POR UDF

Enquanto a escada estava escrita como IF(L>=6;...;IF(L>=5;...)) em tres lugares
-- Estatistica, Painel e BI --, mexer numa faixa exigia lembrar dos tres. O
primeiro esquecido divergiria em silencio, e a divergencia so apareceria quando
o gestor comparasse a planilha com o relatorio. Agora existe uma implementacao,
em mQualidade, versionada em src_producao e importada pelo build.

O CUSTO: 164 chamadas de UDF por recalculo (80 linhas x 2 colunas na
Estatistica, mais 4 no Painel). Cronometrado abaixo.

Uso: python instalar_qualidade_vba.py <arquivo.xlsm>
"""
import io
import os
import sys
import time
import subprocess

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', write_through=True)
import win32com.client as w

RAIZ = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
# SO mQualidade. mBI NAO entra aqui -- e a licao desta etapa.
#
# O arquivo de trabalho nao tem mBI nem mEstatistica, e nao tem aba BI_Data:
# a camada de BI e montada pelo build (aplicar_vba.ps1 + aplicar_bi_data.ps1)
# dentro do ARTEFATO, a partir de src_producao e src_hardening1. Importar mBI
# aqui faz o projeto parar de compilar, porque mBI chama mEstatistica, que nao
# esta neste arquivo -- e o resultado e uma caixa "Erro de compilacao: 'Sub' ou
# 'Function' nao definida" que, numa instancia com Visible = False, trava o
# script em silencio, com o Excel ocioso e respondendo.
#
# O ADR-033 no BI e mudanca de FONTE (mBI.bas 60 -> 65 colunas, o porteiro em
# aplicar_bi_data.ps1 e o esquema em gerar_pbip.py). Quem prova isso e o build,
# nao este script.
MODULOS = [('mQualidade', os.path.join(RAIZ, 'src_producao', 'mQualidade.bas'))]

EST_R0, EST_RN = 14, 93
LST_R0, LST_N = 106, 16


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


def eh_erro(txt):
    t = str(txt)
    return t.startswith('#') and t.strip('#') != ''


def main(caminho):
    caminho = os.path.abspath(caminho)
    for nome, arq in MODULOS:
        if not os.path.exists(arq):
            raise SystemExit('fonte ausente: %s' % arq)
    xl = novo_excel()
    wb = xl.Workbooks.Open(caminho)
    salvou = False
    if wb.ReadOnly:
        wb.Close(False)
        xl.Quit()
        raise SystemExit('Somente leitura (aberto no Excel?): %s' % caminho)
    try:
        xl.Calculation = -4135
        es = wb.Worksheets('Estatística')
        pa = wb.Worksheets('Painel')

        # ---- 1. modulos --------------------------------------------------
        vbp = wb.VBProject
        for nome, arq in MODULOS:
            for c in list(vbp.VBComponents):
                if c.Name == nome:
                    vbp.VBComponents.Remove(c)
            vbp.VBComponents.Import(arq)
            print('importado: %s' % nome)

        # ---- 2. a escada passa a ser chamada, nao copiada ----------------
        n = 0
        for r in range(EST_R0, EST_RN + 1):
            tenta(lambda rr=r: es.Cells(rr, 13).__setattr__(
                'Formula',
                '=IF(NOT(ISNUMBER($L{0})),"",mQualidade.ClassificarSigma($L{0}))'
                .format(rr)))
            tenta(lambda rr=r: es.Cells(rr, 16).__setattr__(
                'Formula',
                '=IF(NOT(ISNUMBER($O{0})),"",mQualidade.ClassificarMargem($O{0}))'
                .format(rr)))
            n += 2
        for lin in (12, 13):
            tenta(lambda l=lin: pa.Cells(l, 15).__setattr__(
                'Formula',
                '=IF(NOT(ISNUMBER($N{0})),"",mQualidade.ClassificarSigma($N{0}))'
                .format(l)))
            n += 1
        for lin in (18, 19):
            tenta(lambda l=lin: pa.Cells(l, 15).__setattr__(
                'Formula',
                '=IF(NOT(ISNUMBER($N{0})),"",mQualidade.ClassificarMargem($N{0}))'
                .format(l)))
            n += 1
        print('%d celulas passaram a chamar mQualidade' % n)

        # ---- 3. compila? prova por chamada direta ------------------------
        provas = [('ClassificarSigma', (6.0,), 'Classe mundial'),
                  ('ClassificarSigma', (2.99,), 'Inadequado'),
                  ('ClassificarSigma', ('',), ''),
                  ('ClassificarMargem', (0.0,), 'Margem critica'),
                  ('ClassificarMargem', (-1.0,), 'ETp excedido'),
                  ('ClassificarMargem', (50.0,), 'Dentro do orcamento')]
        for fn, args, esperado in provas:
            got = tenta(lambda f=fn, a=args: xl.Run(f, *a))
            marca = 'OK  ' if str(got) == esperado else 'FALHA'
            print('   %s %s%s = %r (esperado %r)' % (marca, fn, args, got, esperado))
            if str(got) != esperado:
                raise SystemExit('mQualidade nao respondeu como esperado -- nada salvo')

        # ---- 4. recalculo cronometrado ----------------------------------
        xl.Calculation = -4105
        t0 = time.time()
        tenta(lambda: xl.CalculateFullRebuild())
        dt = time.time() - t0
        print('recalculo completo: %.1fs' % dt)

        # ---- 5. o BI nao mora aqui: so confere que continua assim --------
        temBI = False
        for ws in wb.Worksheets:
            if ws.Name == 'BI_Data':
                temBI = True
        print('aba BI_Data neste arquivo: %s (esperado: nao -- ela nasce no build)'
              % ('SIM' if temBI else 'nao'))
        if temBI:
            raise SystemExit('BI_Data apareceu no arquivo de trabalho -- '
                             'o pressuposto desta etapa mudou, revise antes de salvar')

        # ---- 6. conferencia ---------------------------------------------
        erros = []
        for r in range(EST_R0, EST_RN + 1):
            for c in range(7, 21):
                t = str(tenta(lambda rr=r, cc=c: es.Cells(rr, cc).Text))
                if eh_erro(t):
                    erros.append('Estatistica!%d,%d=%s' % (r, c, t))
        for r in range(LST_R0 + 2, LST_R0 + 2 + LST_N):
            for c in range(1, 9):
                t = str(tenta(lambda rr=r, cc=c: es.Cells(rr, cc).Text))
                if eh_erro(t):
                    erros.append('Estatistica!%d,%d=%s' % (r, c, t))
        for r in range(5, 26):
            for c in range(1, 20):
                t = str(tenta(lambda rr=r, cc=c: pa.Cells(rr, cc).Text))
                if eh_erro(t):
                    erros.append('Painel!%d,%d=%s' % (r, c, t))
        print('celulas em erro: %d' % len(erros))
        for e in erros[:8]:
            print('   %s' % e)
        if erros:
            raise SystemExit('erro de formula -- nada salvo')
        if dt > 30:
            print('AVISO: recalculo em %.1fs -- acima do esperado' % dt)

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
