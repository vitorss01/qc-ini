# -*- coding: utf-8 -*-
"""materializar_cfg_westgard.py - ADR-044: a aba e projecao, nao fonte

POR QUE A ABA NAO PODE SER EDITAVEL

A tentacao obvia seria dar ao laboratorio uma aba onde ele configura os
detectores. Nesta fase isso seria um risco: alguem trocaria "6x N3_R2" por
"6x SAME_LEVEL_R6" e mudaria a regra de negocio sem que nada acusasse -- o
motor continuaria rodando N3_R2 e a planilha diria outra coisa. Seria
recriar, na interface, exatamente o problema que os ADR-041 a 044
eliminaram no codigo.

Entao: a fonte canonica e DetectoresWestgard(), no motor. Esta aba e uma
MATERIALIZACAO para auditoria e para o ETL do Power BI, gerada a partir
dela, protegida e oculta.

O TESTE DE SINCRONISMO

--conferir compara a aba com a fonte e falha se divergirem. E o que impede a
projecao de virar uma segunda verdade: se alguem editar a aba a mao, o teste
acusa na proxima execucao.

Uso:
    python materializar_cfg_westgard.py <arquivo.xlsm>
    python materializar_cfg_westgard.py <arquivo.xlsm> --conferir
"""
import io
import os
import sys
import time

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8',
                              write_through=True)
import win32com.client as w

ABA = 'Cfg_Westgard_Escopo'
SENHA = 'qcini2025'

CAB = ['Area', 'Niveis_QC', 'Regra', 'Detector', 'Escopo', 'N', 'R',
       'Limiar', 'Ativo', 'Oficial', 'Complementar', 'Descricao',
       'Versao_Regra']

VERSAO = 'ADR-044'

# Colunas da fonte: Area|Niveis|Regra|Detector|Ativo|N|R|Escopo|Limiar
I_AREA, I_NIV, I_REGRA, I_DET, I_ATIVO, I_N, I_R, I_ESC, I_LIM = range(9)

DESCRICAO = {
    'INDIVIDUAL': 'Resultado isolado alem do limite, sem depender de outro nivel nem da corrida anterior',
    'WITHIN_RUN': 'Avaliado apenas dentro da corrida atual',
    'ACROSS_RUN_SAME_LEVEL': 'Corridas consecutivas do mesmo nivel de controle',
    'N2_R2': '2 niveis x 2 corridas = 4 observacoes',
    'N2_R4': '2 niveis x 4 corridas = 8 observacoes',
    'N3_R1': '3 niveis da mesma corrida',
    'N3_R2': '3 niveis x 2 corridas = 6 observacoes',
    'SAME_LEVEL_R3': 'Complementar: 3 corridas seguidas do mesmo nivel',
    'SAME_LEVEL_R4': 'Complementar: 4 corridas seguidas do mesmo nivel',
    'SAME_LEVEL_R6': 'Complementar: 6 corridas seguidas do mesmo nivel',
    'SAME_LEVEL_R8': 'Complementar: 8 corridas seguidas do mesmo nivel',
}


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


def linhas_da_fonte(tabela):
    fora = []
    for bruto in tabela.split(';'):
        c = bruto.split('|')
        if len(c) < 9:
            continue
        ativo = (c[I_ATIVO] == '1')
        fora.append([
            c[I_AREA], int(c[I_NIV]), c[I_REGRA], c[I_DET], c[I_ESC],
            int(c[I_N]), int(c[I_R]), int(c[I_LIM]),
            ativo, ativo, not ativo,
            DESCRICAO.get(c[I_DET], ''), VERSAO,
        ])
    return fora


def ler_aba(ws, n):
    fora = []
    for r in range(4, 4 + n):
        linha = []
        for c in range(1, len(CAB) + 1):
            v = ws.Cells(r, c).Value
            linha.append(v)
        if linha[0] is None:
            break
        fora.append(linha)
    return fora


def comparavel(linha):
    """Normaliza para comparar aba x fonte sem tropecar em tipo: o Excel
    devolve numero como float e booleano como bool."""
    fora = []
    for v in linha:
        if isinstance(v, bool):
            fora.append('V' if v else 'F')
        elif isinstance(v, float) and v == int(v):
            fora.append(str(int(v)))
        else:
            fora.append('' if v is None else str(v).strip())
    return fora


def main():
    caminho = os.path.abspath(sys.argv[1])
    conferir = '--conferir' in sys.argv

    xl = w.DispatchEx('Excel.Application')
    xl.Visible = False
    xl.DisplayAlerts = False
    xl.EnableEvents = False
    xl.AutomationSecurity = 1
    wb = xl.Workbooks.Open(caminho, 0, conferir)
    salvou = False
    try:
        tabela = str(xl.Run('DetectoresWestgard'))
        area = str(xl.Run('AreaDoProduto'))
        fonte = [L for L in linhas_da_fonte(tabela)
                 if L[0].upper() == area.upper()]
        print('%s: %d detectores na fonte canonica' % (area, len(fonte)))

        existe = False
        for i in range(1, wb.Worksheets.Count + 1):
            if wb.Worksheets(i).Name == ABA:
                existe = True
                break

        if conferir:
            if not existe:
                raise SystemExit('FALHA: aba %s nao existe' % ABA)
            ws = wb.Worksheets(ABA)
            naAba = ler_aba(ws, len(fonte) + 5)
            if len(naAba) != len(fonte):
                raise SystemExit('FALHA: aba tem %d linhas, fonte tem %d'
                                 % (len(naAba), len(fonte)))
            divergencias = 0
            for i, (a, f) in enumerate(zip(naAba, fonte)):
                if comparavel(a) != comparavel(f):
                    divergencias += 1
                    print('  linha %d diverge:' % (i + 4))
                    print('    aba   : %s' % comparavel(a))
                    print('    fonte : %s' % comparavel(f))
            if divergencias:
                raise SystemExit('FALHA: %d linha(s) divergem da fonte canonica'
                                 % divergencias)
            print('SINCRONIZADA: aba identica a DetectoresWestgard()')
            return 0

        estrutura = bool(wb.ProtectStructure)
        if estrutura:
            wb.Unprotect(SENHA)

        if not existe:
            ws = wb.Worksheets.Add()
            ws.Name = ABA
        else:
            ws = wb.Worksheets(ABA)
            if ws.ProtectContents:
                try:
                    ws.Unprotect(SENHA)
                except Exception:
                    ws.Unprotect()
        ws.Visible = -1                      # visivel para poder escrever
        tenta(lambda: ws.Cells.ClearContents())

        tenta(lambda: ws.Cells(1, 1).__setattr__(
            'Value', 'CONFIGURACAO WESTGARD POR ESCOPO - PROJECAO AUTOMATICA'))
        tenta(lambda: ws.Cells(2, 1).__setattr__(
            'Value', 'Gerada de mEstatistica.DetectoresWestgard(). NAO editar: '
                     'a fonte e o motor, e materializar_cfg_westgard.py --conferir '
                     'falha se esta aba divergir dele.'))
        for j, h in enumerate(CAB, start=1):
            tenta(lambda jj=j, hh=h: ws.Cells(3, jj).__setattr__('Value', hh))
        ws.Range(ws.Cells(3, 1), ws.Cells(3, len(CAB))).Font.Bold = True

        for i, L in enumerate(fonte):
            for j, v in enumerate(L, start=1):
                tenta(lambda rr=4 + i, jj=j, vv=v:
                      ws.Cells(rr, jj).__setattr__('Value', vv))

        ws.Columns.AutoFit()
        print('%d linhas materializadas' % len(fonte))

        try:
            ws.Protect(SENHA)
        except Exception:
            pass
        ws.Visible = 2                       # xlSheetVeryHidden
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
