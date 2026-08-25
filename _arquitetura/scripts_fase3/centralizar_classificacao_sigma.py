# -*- coding: utf-8 -*-
"""centralizar_classificacao_sigma.py - ADR-043: uma escada de Sigma so

O DEFEITO

A coluna Classificacao da Estatistica da Hematologia carregava a escada
FIXA NA CELULA:

    =IF($L7>=6,"Excelente",IF($L7>=4,"Bom",IF($L7>=3,"Marginal","Inaceitavel")))

Tres erros nela:

  1. >=6 devolve "Excelente". A faixa e "Classe mundial".
  2. A faixa 5 a <6 NAO EXISTE: Sigma 5,5 nao passa no >=6, cai no >=4 e
     aparece como "Bom". Um metodo excelente exibido como apenas bom.
  3. Abaixo de 3 diz "Inaceitavel"; o resto do projeto diz "Desempenho
     inadequado". Duas etiquetas para o mesmo estado.

A Bioquimica ja usava mQualidade.ClassificarSigma -- foi modernizada pelo
ADR-033 e a Hematologia ficou para tras. Este script iguala.

A CLASSIFICACAO OFICIAL, uma para todo o QC_INI

    Sigma < 3 .......... Desempenho inadequado
    3 <= Sigma < 4 ..... Marginal
    4 <= Sigma < 5 ..... Bom
    5 <= Sigma < 6 ..... Excelente
    Sigma >= 6 ......... Classe mundial

Uso: python centralizar_classificacao_sigma.py <arquivo.xlsm>
"""
import io
import os
import re
import sys
import time

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8',
                              write_through=True)
import win32com.client as w

SENHA = 'qcini2025'
ABA = 'Estat' + chr(0x00ED) + 'stica'

# A escada fixa, em qualquer variacao de rotulo. Reconhecer pelo PADRAO e nao
# pelo texto exato: o mesmo defeito ja apareceu com "Inaceitavel" e poderia
# aparecer com outra etiqueta sem deixar de ser a mesma escada duplicada.
PADRAO_ESCADA = re.compile(
    r'IF\s*\(\s*\$?[A-Z]+\d+\s*>=\s*6\s*,\s*"[^"]*"', re.IGNORECASE)


def tenta(fn, vezes=6):
    ult = None
    for i in range(vezes):
        try:
            return fn()
        except Exception as e:
            ult = e
            if 'rejeitada' not in str(e).lower() and 'rejected' not in str(e).lower():
                raise
            time.sleep(1.0 + 0.8 * i)
    raise ult


def main():
    caminho = os.path.abspath(sys.argv[1])
    xl = w.DispatchEx('Excel.Application')
    xl.Visible = False
    xl.DisplayAlerts = False
    xl.EnableEvents = False
    xl.AutomationSecurity = 1
    wb = xl.Workbooks.Open(caminho)
    salvou = False
    try:
        if wb.ReadOnly:
            raise SystemExit('somente leitura: %s' % caminho)

        estrutura = bool(wb.ProtectStructure)
        if estrutura:
            wb.Unprotect(SENHA)

        ws = wb.Worksheets(ABA)
        reprot = bool(ws.ProtectContents)
        if reprot:
            try:
                ws.Unprotect(SENHA)
            except Exception:
                ws.Unprotect()

        # acha a coluna de Sigma pelo cabecalho, e nao por posicao fixa: os
        # dois produtos tem layouts diferentes na Estatistica
        rCab = cSigma = 0
        for r in range(1, 21):
            for c in range(1, 31):
                # Os dois produtos rotulam a coluna de forma diferente:
                # "Sigma" na Hematologia, "SIX SIGMA" na Bioquimica (ADR-033).
                # Procurar so por um dos dois faria o script "nao achar" o
                # cabecalho e passar em branco pelo produto errado.
                if str(ws.Cells(r, c).Text).strip().upper() in ('SIGMA', 'SIX SIGMA'):
                    rCab, cSigma = r, c
                    break
            if rCab:
                break
        if not rCab:
            raise SystemExit('nao achei o cabecalho "Sigma" na %s' % ABA)

        cClasse = cSigma + 1
        rotulo = str(ws.Cells(rCab, cClasse).Text).strip()
        print('cabecalho na linha %d; Sigma col %d; classificacao col %d (%r)'
              % (rCab, cSigma, cClasse, rotulo))

        trocadas = ja = 0
        vazias = 0
        for r in range(rCab + 1, rCab + 200):
            f = str(ws.Cells(r, cClasse).Formula)
            if not f:
                vazias += 1
                if vazias > 20:
                    break
                continue
            vazias = 0
            if 'ClassificarSigma' in f:
                ja += 1
                continue
            if PADRAO_ESCADA.search(f):
                # Address vem como PROPRIEDADE pelo COM tardio, nao metodo:
                # chama-lo estoura "'str' object is not callable".
                letra = ''
                n = cSigma
                while n > 0:
                    n, resto = divmod(n - 1, 26)
                    letra = chr(65 + resto) + letra
                col = '%s%d' % (letra, r)
                nova = ('=IF(NOT(ISNUMBER(${0})),"",mQualidade.ClassificarSigma(${0}))'
                        .format(col))
                tenta(lambda rr=r, nn=nova:
                      ws.Cells(rr, cClasse).__setattr__('Formula', nn))
                trocadas += 1

        print('formulas trocadas: %d ; ja centralizadas: %d' % (trocadas, ja))

        tenta(lambda: xl.CalculateFullRebuild())

        # prova nas fronteiras, pela funcao que as celulas agora chamam
        print()
        print('   %-7s %s' % ('Sigma', 'classificacao'))
        esperado = {2.99: 'Desempenho inadequado', 3.0: 'Marginal',
                    3.99: 'Marginal', 4.0: 'Bom', 4.99: 'Bom',
                    5.0: 'Excelente', 5.99: 'Excelente',
                    6.0: 'Classe mundial', 6.01: 'Classe mundial'}
        falhas = 0
        for s in sorted(esperado):
            got = str(xl.Run('ClassificarSigma', s))
            ok = (got == esperado[s])
            if not ok:
                falhas += 1
            print('   %-7s %-24s %s' % (s, got, 'PASS' if ok else
                                        'FAIL (esperado %s)' % esperado[s]))
        if falhas:
            raise SystemExit('%d fronteira(s) erradas -- nada salvo' % falhas)

        erros = 0
        for r in range(rCab + 1, rCab + 200):
            if str(ws.Cells(r, cClasse).Text).startswith('#'):
                erros += 1
        if erros:
            raise SystemExit('%d celula(s) em erro -- nada salvo' % erros)

        if reprot:
            try:
                ws.Protect(SENHA)
            except Exception:
                pass
        if estrutura:
            try:
                wb.Protect(SENHA, True, False)
            except Exception:
                pass

        wb.Save()
        salvou = True
        print('\nSALVO: %s' % caminho)
    finally:
        try:
            wb.Close(salvou)
        except Exception:
            pass
        try:
            xl.Quit()
        except Exception:
            pass
    return 0


if __name__ == '__main__':
    sys.exit(main())
