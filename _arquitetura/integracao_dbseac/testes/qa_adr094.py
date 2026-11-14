# -*- coding: utf-8 -*-
"""QA do ADR-094 (13/11/2026): data-hora com segundos (dd/mm/aaaa hh:mm:ss) em tudo o que o usuario ve; campo so de
data continua dd/mm/aaaa. Roda numa COPIA instalada (instalar_adr094.py); nada e salvo.

Uso: python qa_adr094.py <Bioquimica|Hematologia> <copia_instalada.xlsm> <saida.json>

  F01 cada coluna de data/data-hora das tabelas (lista do instalar_adr094.COLUNAS): o TEXTO exibido = o VALOR
      formatado aqui de forma independente (dia antes do mes; segundos nas de data-hora; so data nas de data; sem ####),
      na 1a, na ultima e em ate 60 linhas espalhadas -- pega o 'm/d/yyyy' das linhas da carga de 11/11 (Bioquimica)
  F02 linha NOVA da tabela (como a carga do motor VBA faz: lo.Resize) herda o formato: tblCQ_Final e tblDB_Recebimento
  F03 Inicio "Ultimo resultado recebido" e DB A5 mostram data-hora com segundos
  F04 Westgard_Fotos DATA_FOTO e LotesStore REGISTRADO_EM com segundos
  F05 depois de um ATUALIZAR (so entradas): nenhum DETALHE do QA_INTEGRACAO com data-hora sem segundos
  (a dica do ponto e a ficha da janela com segundos: qa_inativar N02/N06, qa_inativar_ux U01, qa_adr085 Z01)
  Y00 seguranca do teste: arquivo de entrada intacto
"""
import datetime as dt
import json
import os
import re
import sys
import traceback

AQUI = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, AQUI)
from qa_final import QA, reg, RES, sha256_arquivo  # noqa: E402
sys.path.insert(0, os.path.join(AQUI, '..'))
import instalar_integracao as ii  # noqa: E402
import instalar_adr094 as i94  # noqa: E402

RX_DH = re.compile(r'\d{2}/\d{2}/\d{4} \d{2}:\d{2}:\d{2}')
RX_SEM_SEG = re.compile(r'\b\d{2}/\d{2}/\d{4} \d{2}:\d{2}(?!:\d)')


def esperado(v, fmt):
    if fmt == i94.DH:
        return (v.replace(tzinfo=None) + dt.timedelta(microseconds=500000)).strftime('%d/%m/%Y %H:%M:%S')
    return v.strftime('%d/%m/%Y')


def executar(produto, caminho, saida):
    hash0 = sha256_arquivo(caminho)
    q = QA(produto, caminho)
    try:
        xl = q.ex.xl
        wb = q.wb

        # ------------------------------------------------------------ F01
        ruins, conferidas, por_col = [], 0, {}
        for tab, c, fmt in i94.COLUNAS:
            _, lo = i94.achar_tabela(wb, tab)
            rng = lo.ListColumns(c).DataBodyRange
            if rng is None:
                por_col[f'{tab}[{c}]'] = 'vazia'
                continue
            n = rng.Rows.Count
            linhas = sorted({1, n} | {1 + k * (n - 1) // 59 for k in range(60)} if n > 1 else {1})
            k = 0
            for i in linhas:
                cel = rng.Cells(i, 1)
                v = cel.Value
                if v is None or not hasattr(v, 'strftime'):
                    continue
                t, e = str(cel.Text), esperado(v, fmt)
                k += 1
                if t != e:
                    ruins.append(f'{tab}[{c}] linha {i}: exibe {t!r}, valor {e}')
            conferidas += k
            por_col[f'{tab}[{c}]'] = k
        reg('F01 Data/data-hora das tabelas: texto exibido = valor (dd/mm/aaaa hh:mm:ss nas de data-hora; dd/mm/aaaa nas de '
            'data; dia antes do mês; sem ####), 1ª, última e até 60 linhas espalhadas de cada coluna',
            conferidas > 0 and not ruins, {'conferidas': conferidas, 'por_coluna': por_col, 'divergencias': ruins[:12]})

        # ------------------------------------------------------------ F02 linha nova herda o formato
        xl.EnableEvents = False
        f02 = {}
        okf2 = True
        teste_v = dt.datetime(2026, 11, 13, 8, 23, 52)
        for tab, cols in (('tblCQ_Final', ('DATA', 'DATA_HORA', 'RECEBIDO_EM', 'DATA_INATIVACAO')),
                          ('tblDB_Recebimento', ('DATA', 'DATA_HORA', 'RECEBIDO_EM'))):
            ws, lo = i94.achar_tabela(wb, tab)
            ws.Unprotect(ii.SENHA)
            n = lo.ListRows.Count
            lo.Resize(lo.HeaderRowRange.Resize(n + 2, lo.ListColumns.Count))
            for c in cols:
                fmt = [f for t_, c_, f in i94.COLUNAS if t_ == tab and c_ == c][0]
                cel = lo.ListColumns(c).DataBodyRange.Cells(n + 1, 1)
                cel.Value = teste_v
                t = str(cel.Text)
                e = esperado(teste_v, fmt)
                f02[f'{tab}[{c}]'] = t
                okf2 = okf2 and t == e
            # desfaz (o F05 roda um ATUALIZAR: linha de teste na tabela entraria na carga)
            lo.DataBodyRange.Rows(n + 1).Resize(2).ClearContents()
            lo.Resize(lo.HeaderRowRange.Resize(n + 1, lo.ListColumns.Count))
            if lo.ListRows.Count != n:
                raise RuntimeError(f'F02: {tab} nao voltou a {n} linhas')
        reg('F02 Linha nova da tabela (como a carga faz: lo.Resize) herda o formato da coluna: 13/11/2026 08:23:52 nas de '
            'data-hora, 13/11/2026 nas de data', okf2, f02)

        # ------------------------------------------------------------ F03 formulas
        ini = [w for w in wb.Worksheets if w.Name == 'Início'][0]
        xl.Calculate()
        c_ini = [str(ini.Cells(r, 3).Text) for r in range(20, 32) if i94.INICIO_NOVO in str(ini.Cells(r, 3).Formula)]
        a5 = str(wb.Worksheets('DB').Range('A5').Text)
        n_dh_a5 = len(RX_DH.findall(a5))
        reg('F03 Início "Último resultado recebido" e DB A5 (último resultado / última chegada) com segundos',
            len(c_ini) == 1 and (c_ini[0] == 'nenhum' or bool(RX_DH.fullmatch(c_ini[0]))) and n_dh_a5 == 2
            and not RX_SEM_SEG.search(a5), {'Início': c_ini, 'DB_A5': a5[:200]})

        # ------------------------------------------------------------ F04 abas tecnicas
        ws = wb.Worksheets('Westgard_Fotos')
        ult = ws.Cells(ws.Rows.Count, 6).End(-4162).Row
        foto = str(ws.Cells(ult, 6).Text) if ult > 3 else ''
        ls = wb.Worksheets('LotesStore')
        cu = next(c for c in range(1, 40) if i94.coluna_cab(ls, c, 'REGISTRADO_EM'))
        ul = ls.Cells(ls.Rows.Count, cu).End(-4162).Row
        reg_em = str(ls.Cells(ul, cu).Text)
        ok4 = (foto == '' or bool(RX_DH.fullmatch(foto))) and (bool(RX_DH.fullmatch(reg_em)) or reg_em == 'REGISTRADO_EM')
        reg('F04 Westgard_Fotos DATA_FOTO e LotesStore REGISTRADO_EM com segundos', ok4,
            {'Westgard_Fotos_ultima': foto, 'LotesStore_ultimo': reg_em})

        # ------------------------------------------------------------ F05 textos do QA
        xl.EnableEvents = True
        r_at = str(q.run('mIntegracao.AtualizarSoEntradas', teto=1500))
        qa = q.ler('tblQA_Integracao')
        sem = [(x.get('CODIGO'), str(x.get('DETALHE'))[:160]) for x in qa if RX_SEM_SEG.search(str(x.get('DETALHE') or ''))]
        com = sum(1 for x in qa if RX_DH.search(str(x.get('DETALHE') or '')))
        reg('F05 Depois de um ATUALIZAR: nenhum DETALHE do QA_INTEGRACAO com data-hora sem segundos', r_at.startswith('OK|')
            and not sem, {'atualizar': r_at[:80], 'detalhes_com_data_hora': com, 'sem_segundos': sem[:8]})
    except Exception as e:                                     # noqa: BLE001
        reg('F99 Execução interrompida: os testes seguintes a este ponto não rodaram', False,
            {'erro': f'{type(e).__name__}: {e}', 'traceback': traceback.format_exc()[-1500:]})
    finally:
        q.fechar(salvar=False)
    hash1 = sha256_arquivo(caminho)
    reg('Y00 Segurança do teste: arquivo de entrada intacto (SHA-256 igual antes e depois; nada salvo nele)', hash0 == hash1,
        {'arquivo': caminho, 'sha256_antes': hash0, 'sha256_depois': hash1})
    if saida:
        with open(saida, 'w', encoding='utf-8') as f:
            json.dump(RES, f, ensure_ascii=False, indent=1, default=str)
    falhas = [r for r in RES if r['resultado'] == 'FAIL']
    print(f'\n=== {produto} ADR-094 (data-hora com segundos): {len(RES) - len(falhas)} PASS / {len(falhas)} FAIL ===',
          flush=True)
    return falhas


if __name__ == '__main__':
    falhas = executar(sys.argv[1], os.path.abspath(sys.argv[2]), sys.argv[3] if len(sys.argv) > 3 else None)
    sys.exit(1 if falhas else 0)
