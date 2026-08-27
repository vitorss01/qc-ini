# -*- coding: utf-8 -*-
"""ligar_carimbo_estatistica.py - ADR-049: a celula declara que depende do banco

POR QUE

O carimbo de conteudo (instalar_carimbo_banco.ps1) faz o CACHE do
mEstatPeriodo invalidar quando o banco muda. Mas a celula continuava sem
declarar essa dependencia, entao o Excel nao a marcava suja: so
CalculateFullRebuild atualizava a Estatistica.

E o fluxo real nao faz rebuild -- mDados chama Application.Calculate depois de
gravar. Medido, apos excluir um resultado:

    F9 (Application.Calculate) ......... n = 25   (errado)
    CalculateFullRebuild ............... n = 24   (certo)

Passando DB_Carimbo como 8o argumento, a celula entra no grafo de calculo do
Excel e o F9 volta a refrescar.

POR QUE NAO O INTERVALO DE DADOS COMO ARGUMENTO

Foi a proposta inicial (DB_Resultados!A:F em 320 formulas). Faria centenas de
UDFs dependerem de milhoes de celulas. DB_Carimbo e UMA celula que o Excel ja
recalcula quando o intervalo muda -- mesma dependencia, custo desprezivel.

IDEMPOTENTE: quem ja tem o argumento nao e tocado.

Uso: python ligar_carimbo_estatistica.py <arquivo.xlsm>
"""
import io
import os
import sys

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8',
                              write_through=True)
import win32com.client as w

ABA = 'Estat' + chr(0x00ED) + 'stica'
SENHA = 'qcini2025'


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

        # o nome tem de existir antes: sem ele a formula fica #NOME?
        try:
            wb.Names('DB_Carimbo').RefersToRange.Value
        except Exception:
            raise SystemExit('DB_Carimbo ausente -- rode '
                             'instalar_carimbo_banco.ps1 antes')

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

        ur = ws.UsedRange
        ultLin = ur.Row + ur.Rows.Count - 1
        ultCol = ur.Column + ur.Columns.Count - 1
        trocadas = ja = 0
        for r in range(1, min(ultLin, 300) + 1):
            for c in range(1, min(ultCol, 40) + 1):
                f = ws.Cells(r, c).Formula
                if not isinstance(f, str) or 'EstatPeriodo(' not in f:
                    continue
                if 'DB_Carimbo' in f:
                    ja += 1
                    continue
                # Recorta o fecho da chamada de EstatPeriodo e acrescenta o
                # argumento. Procurar o ")" a partir do nome da funcao evita
                # confundir com o ")" do IF externo.
                i = f.find('EstatPeriodo(')
                j = f.find(')', i)
                if j < 0:
                    continue
                novo = f[:j] + ',DB_Carimbo' + f[j:]
                ws.Cells(r, c).Formula = novo
                trocadas += 1

        print('formulas ligadas ao carimbo: %d  (ja ligadas: %d)'
              % (trocadas, ja))

        xl.CalculateFullRebuild()
        erros = []
        for r in range(1, min(ultLin, 300) + 1):
            for c in range(1, min(ultCol, 40) + 1):
                t = str(ws.Cells(r, c).Text or '')
                if t.startswith('#') and len(t) > 1:
                    erros.append('%s%d = %s' % (chr(64 + c) if c <= 26 else '?',
                                                r, t))
        if erros:
            raise SystemExit('celulas em erro apos a troca -- nada salvo:\n  %s'
                             % '\n  '.join(erros[:8]))

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
    return 0


if __name__ == '__main__':
    sys.exit(main())
