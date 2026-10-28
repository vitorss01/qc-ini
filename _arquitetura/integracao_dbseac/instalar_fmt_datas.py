"""Datas com o ano visivel (28/10/2026). Pedido do usuario: na aba DB as colunas DATA (H), DATA_HORA e RECEBIDO_EM (R)
mostravam 'dd/mm/yyyy' -- o 'yyyy' foi gravado na lingua do Excel e o Excel em portugues o leu como TEXTO.

Em TODAS as tabelas do arquivo, a coluna cujo formato tem o 'y' literal (\\y) recebe o formato certo pelo
camada_dados.fmt_data (codigo do ano do Excel instalado). A linha 5 da aba DB passa a montar o ano com YEAR().
Idempotente: a 2a execucao registra 0 mudancas. Nao toca dados, BI nem VBA.

Uso: python instalar_fmt_datas.py <Bioquimica|Hematologia> <arquivo.xlsm> [--sem-salvar]
"""
import os
import sys

AQUI = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, AQUI)
sys.path.insert(0, os.path.join(AQUI, '..', 'etapa1_multilote'))
import xlh  # noqa: E402
import instalar_integracao as ii  # noqa: E402
from instalar_seguranca_usuarios import foto_protecao, devolver_protecao  # noqa: E402
from camada_dados import fmt_data, formulas_resumo, LINHA_CAB  # noqa: E402

# formato pretendido por nome de coluna (o mesmo do camada_dados)
FMT = {'DATA': 'dd/mm/yyyy', 'DATA_HORA': 'dd/mm/yyyy hh:mm:ss', 'RECEBIDO_EM': 'dd/mm/yyyy hh:mm',
       'DATA_INATIVACAO': 'dd/mm/yyyy hh:mm', 'REGISTRADO_EM': 'dd/mm/yyyy hh:mm'}


def log(m):
    print(m, flush=True)


def main(produto, caminho, salvar=True):
    caminho = os.path.abspath(caminho)
    mud = []
    ex = xlh.Excel()
    log(f'EXCEL_PID {ex.pid}')
    try:
        wb = ex.abrir(caminho)
        if wb.ReadOnly:
            raise SystemExit(f'arquivo aberto SOMENTE LEITURA (outra instancia o segura): {caminho}')
        estrutura = bool(wb.ProtectStructure)
        foto = foto_protecao(wb)
        ii.desproteger_tudo(wb)
        ano = str(ex.xl.International[19 - 1] or 'y')            # xlYearCode
        for ws in wb.Worksheets:
            for lo in ws.ListObjects:
                for col in lo.ListColumns:
                    rng = col.DataBodyRange
                    if rng is None:
                        continue
                    # por esta automacao o formato volta na lingua local: 'yyyy' num Excel cujo ano e 'a' = texto
                    nf = str(rng.Cells(1, 1).NumberFormatLocal)
                    if ano != 'y' and 'y' in nf.lower():
                        fmt = FMT.get(col.Name, 'dd/mm/yyyy hh:mm' if 'h' in nf else 'dd/mm/yyyy')
                        fmt_data(rng, fmt)
                        if '#' in str(rng.Cells(1, 1).Text):              # data com hora mais larga que a coluna
                            rng.EntireColumn.ColumnWidth = max(rng.EntireColumn.ColumnWidth, 19)
                        mud.append(f'{ws.Name}!{lo.Name}[{col.Name}]: {nf} -> {rng.Cells(1, 1).Text}')
        f_db = formulas_resumo()['DB']
        c5 = wb.Worksheets('DB').Cells(LINHA_CAB - 1, 1)
        if str(c5.Formula) != f_db:
            c5.Formula = f_db
            mud.append(f'DB!A5: {c5.Text[:120]}')
        for m in mud:
            log('   ' + m)
        log(f'MUDANCAS {len(mud)}')
        devolver_protecao(wb, foto)
        if estrutura:
            wb.Protect(ii.SENHA, True, False)
        if salvar and mud:
            wb.Save()
            log('salvo')
    finally:
        ex.fechar()


if __name__ == '__main__':
    a = sys.argv[1:]
    main(a[0], a[1], salvar='--sem-salvar' not in a)
