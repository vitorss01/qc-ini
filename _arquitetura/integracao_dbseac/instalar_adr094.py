# -*- coding: utf-8 -*-
"""instalar_adr094.py -- ADR-094 (13/11/2026, pedido do usuario): toda data-hora mostrada com segundos
(dd/mm/aaaa hh:mm:ss, ex.: 13/11/2026 08:23:52). Campo que e so data continua dd/mm/aaaa. SO APRESENTACAO: nenhum valor,
chave, RUN, regra ou criterio muda.

Uso:  python instalar_adr094.py <Bioquimica|Hematologia> <arquivo.xlsm> [--sem-salvar]

Idempotente (a 2a execucao registra 0 mudancas). Passos:
  1. consulta QA_INTEGRACAO regravada de pq/ e conferida (DataHoraTxt ao segundo, igual ao mEtlQA.DH);
  2. todo o VBA atual das fontes (codigo_atual.py), compilacao exigida (textos com segundos: dica do ponto, ficha da
     janela, mensagem do ATUALIZAR, assinatura, historico, calibracao, ultimo registro do lote, Westgard_Fotos);
  3. planilhas -- formato por TABELA + COLUNA (nunca so pelo nome: a DATA dos comentarios e data-hora), na coluna
     inteira (as linhas novas da carga herdam o formato da coluna da tabela). Corrige junto o defeito da Bioquimica:
     DATA/RECEBIDO_EM/DATA_INATIVACAO com 'm/d/yyyy' nas linhas da carga de 11/11 (mes na frente);
     Westgard_Fotos F, LotesStore REGISTRADO_EM (coluna inteira: o mLotes grava Now sem formatar); formulas DB A5 e
     Inicio "Ultimo resultado";
     coluna alargada so se a data-hora nao couber (####);
  4. fumaca (texto exibido = valor, dia antes do mes); 5. devolve a protecao e salva com os eventos desligados.
"""
import os
import re
import sys

AQUI = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, AQUI)
sys.path.insert(0, os.path.join(AQUI, '..', 'etapa1_multilote'))
import xlh  # noqa: E402
import pqlib  # noqa: E402
import instalar_integracao as ii  # noqa: E402
import codigo_atual  # noqa: E402
import camada_dados as cd  # noqa: E402
from instalar_seguranca_usuarios import foto_protecao, devolver_protecao  # noqa: E402

MUDANCAS = []
CONSULTAS = ['QA_INTEGRACAO']
DH = 'dd/mm/yyyy hh:mm:ss'
D = 'dd/mm/yyyy'
# (tabela, coluna, formato) -- a fonte unica do que o usuario ve; o teste qa_adr094 le esta lista
COLUNAS = [
    ('tblCQ_Final', 'DATA', D), ('tblCQ_Final', 'DATA_HORA', DH), ('tblCQ_Final', 'RECEBIDO_EM', DH),
    ('tblCQ_Final', 'DATA_INATIVACAO', DH),
    ('tblDB_Recebimento', 'DATA', D), ('tblDB_Recebimento', 'DATA_HORA', DH), ('tblDB_Recebimento', 'RECEBIDO_EM', DH),
    ('tblDB_Organizado', 'DATA', D), ('tblDB_Organizado', 'DATA_HORA', DH),
    ('tblQA_Integracao', 'DATA_HORA', DH),
    ('tblResultados_Manuais', 'REGISTRADO_EM', DH),
    ('tblInativacao_NaoConformes', 'DATA_INATIVACAO', DH),
    ('tblComentariosTecnicos', 'DATA', DH),                  # carimbo com Now: data-hora, apesar do nome
    ('tblSyncLog', 'DATA_MIN', DH), ('tblSyncLog', 'DATA_MAX', DH),
    # INICIO/FIM ja tinham segundos desde o ADR-079, mas na largura 12 (19 caracteres nao cabem: '####'); a aba
    # Sync_Log e visivel a todos depois do login (mSeguranca.UnlockApp) -- achado do qa_adr094 F06, 14/11
    ('tblSyncLog', 'INICIO', DH), ('tblSyncLog', 'FIM', DH),
]
INICIO_ANTIGO = 'TEXT(MAX(tblCQ_Final[DATA_HORA]),"dd/mm/aaaa hh:mm")'
INICIO_NOVO = 'TEXT(MAX(tblCQ_Final[DATA_HORA]),"dd/mm/aaaa hh:mm:ss")'


def log(m):
    print(m, flush=True)


def normalizar(t):
    return '\n'.join(l.rstrip() for l in str(t).replace('\r\n', '\n').replace('\r', '\n').strip().split('\n'))


def consultas(wb):
    wb.Queries.FastCombine = True
    for q in CONSULTAS:
        atual = pqlib.query_existe(wb, q)
        if atual is None:
            raise SystemExit(f'consulta {q} nao existe no arquivo -- nada foi salvo')
        if normalizar(atual.Formula) != normalizar(pqlib.m_de(q)):
            pqlib.gravar_query(wb, q, pqlib.m_de(q))
            MUDANCAS.append(f'consulta {q}')
    conferir_consultas(wb)


def conferir_consultas(wb):
    for q in CONSULTAS:
        if normalizar(pqlib.query_existe(wb, q).Formula) != normalizar(pqlib.m_de(q)):
            raise SystemExit(f'consulta {q}: formula no arquivo difere da fonte depois de gravar -- nada foi salvo')


def achar_tabela(wb, nome):
    for ws in wb.Worksheets:
        for lo in ws.ListObjects:
            if lo.Name == nome:
                return ws, lo
    raise SystemExit(f'tabela {nome} nao encontrada -- nada foi salvo')


def local(app, fmt):
    return fmt.replace('y', str(app.International[19 - 1] or 'y'))          # xlYearCode (ver camada_dados.fmt_data)


SONDA = 46384.99999                                     # 28/12/2026 23:59:59 (serial OLE, sem fuso: Value2)


LARG_DH = [0.0]                                         # largura em que a data-hora coube nesta execucao


def caber(ws, rng, oque, pode_sondar=True):
    """Alarga a coluna (nunca estreita) ate a data-hora caber: '####' e o que o auditoria_abas acusa.
    14/11/2026 (entrega 40, qa_adr094 F02/F06): coluna VAZIA hoje (DATA_INATIVACAO, REGISTRADO_EM dos manuais) nunca
    mostrava '#' na amostra e ficava estreita -- o 1o valor gravado depois aparecia '####'. Coluna vazia: sonda
    temporaria (SONDA) na 1a celula, devolvida vazia em seguida."""
    col = rng.EntireColumn
    c1 = rng.Cells(1, 1)
    sonda = pode_sondar and c1.Value is None and not c1.HasFormula
    if sonda:
        c1.Value2 = SONDA
    try:
        for _ in range(16):
            amostra = [c1, rng.Cells(rng.Rows.Count, 1)]
            if not any(str(c.Text).startswith('#') for c in amostra if c.Value is not None):
                if c1.Value is not None and str(c1.NumberFormat).endswith('ss'):
                    LARG_DH[0] = max(LARG_DH[0], float(col.ColumnWidth))
                return
            col.ColumnWidth = float(col.ColumnWidth) + 1
            MUDANCAS.append(f'{oque}: coluna alargada para {col.ColumnWidth}' + (' (sonda: coluna vazia)' if sonda else ''))
    finally:
        if sonda:
            c1.ClearContents()


def formatar(app, ws, rng, fmt, oque, pode_sondar=True):
    alvo = local(app, fmt)
    if str(rng.NumberFormatLocal or '') != alvo:          # Null (formatos misturados) tambem cai aqui
        rng.NumberFormatLocal = alvo
        MUDANCAS.append(f'{oque} -> {fmt}')
    caber(ws, rng, oque, pode_sondar)


def colunas(app, wb):
    for tab, c, fmt in COLUNAS:
        ws, lo = achar_tabela(wb, tab)
        try:
            col = lo.ListColumns(c)
        except Exception:                                   # noqa: BLE001
            raise SystemExit(f'{tab}[{c}] nao existe -- nada foi salvo')
        rng = col.DataBodyRange
        vazia = rng is None
        if vazia:                                           # tabela vazia: a linha de insercao guarda o formato
            if lo.InsertRowRange is None:
                log(f'   {tab}[{c}]: tabela vazia sem linha de insercao (formato fica para a 1a linha)')
                continue
            rng = lo.InsertRowRange.Cells(1, col.Index)
        if tab == 'tblSyncLog':                             # = instalar_adr079: lin .. lin + 4000
            rng = ws.Range(rng.Cells(1, 1), rng.Cells(4001, 1))
        # tabela vazia: escrever na linha de insercao criaria uma linha -- sem sonda; a coluna recebe a largura em
        # que a data-hora ja coube nesta execucao (tblCQ_Final[DATA_HORA] vem antes na lista e tem dados)
        formatar(app, ws, rng, fmt, f'{ws.Name}!{tab}[{c}]', pode_sondar=not vazia)
        if vazia and fmt == DH and LARG_DH[0] > float(rng.EntireColumn.ColumnWidth):
            rng.EntireColumn.ColumnWidth = LARG_DH[0]
            MUDANCAS.append(f'{ws.Name}!{tab}[{c}]: tabela vazia, coluna alargada para {LARG_DH[0]}')


def coluna_cab(ws, col, cab, linhas=range(1, 13)):
    for r in linhas:
        if str(ws.Cells(r, col).Value or '').strip().upper() == cab.upper():
            return r
    return None


def avulsas(app, wb):
    # Westgard_Fotos F (o VBA reaplica o mesmo formato a cada foto: mEstatistica)
    ws = wb.Worksheets('Westgard_Fotos')
    r = coluna_cab(ws, 6, 'DATA_FOTO')
    if r is None:
        raise SystemExit('Westgard_Fotos: cabecalho DATA_FOTO nao achado na coluna F -- nada foi salvo')
    ult = max(r + 1, ws.Cells(ws.Rows.Count, 6).End(-4162).Row)
    formatar(app, ws, ws.Range(ws.Cells(r + 1, 6), ws.Cells(ult, 6)), DH, 'Westgard_Fotos!F')
    # LotesStore REGISTRADO_EM (carimbo do registro de Media/DP)
    ws = wb.Worksheets('LotesStore')
    cu = None
    for c in range(1, 40):
        if coluna_cab(ws, c, 'REGISTRADO_EM'):
            cu = c
            break
    if cu is None:
        raise SystemExit('LotesStore: coluna REGISTRADO_EM nao achada -- nada foi salvo')
    r = coluna_cab(ws, cu, 'REGISTRADO_EM')
    formatar(app, ws, ws.Range(ws.Cells(r + 1, cu), ws.Cells(ws.Rows.Count, cu)), DH, 'LotesStore!REGISTRADO_EM')
    # Eventos_Westgard NAO: o instante do evento vem do texto da descricao de Westgard (ISO ao minuto, lido de volta
    # pelas fotos); mostrar ':00' seria precisao falsa (revisao independente, 13/11)


def formulas(wb):
    f_db = cd.formulas_resumo()['DB']
    c5 = wb.Worksheets('DB').Cells(cd.LINHA_CAB - 1, 1)
    if str(c5.Formula) != f_db:
        c5.Formula = f_db
        MUDANCAS.append('DB!A5 (ultimo resultado / ultima chegada com segundos)')
    ini = [w for w in wb.Worksheets if w.Name == 'Início'][0]
    achou = False
    for r in range(20, 32):
        f = str(ini.Cells(r, 3).Formula)
        if INICIO_NOVO in f:
            achou = True
        elif INICIO_ANTIGO in f:
            ini.Cells(r, 3).Formula = f.replace(INICIO_ANTIGO, INICIO_NOVO)
            MUDANCAS.append(f'Início!C{r} (último resultado recebido com segundos)')
            achou = True
    if not achou:
        raise SystemExit('Início: formula do "Último resultado recebido" nao achada -- nada foi salvo')


RX_DH = re.compile(r'^\d{2}/\d{2}/\d{4} \d{2}:\d{2}:\d{2}$')
RX_D = re.compile(r'^\d{2}/\d{2}/\d{4}$')


def fumaca(ex, wb):
    conferir_consultas(wb)
    ex.xl.Calculate()
    ruins = []
    for tab, c, fmt in COLUNAS:
        _, lo = achar_tabela(wb, tab)
        rng = lo.ListColumns(c).DataBodyRange
        if rng is None:
            continue
        n = rng.Rows.Count
        for i in sorted({1, max(1, n // 2), n}):
            cel = rng.Cells(i, 1)
            v = cel.Value
            if v is None or not hasattr(v, 'strftime'):
                continue
            esp = v.strftime('%d/%m/%Y %H:%M:%S' if fmt == DH else '%d/%m/%Y')
            t = str(cel.Text)
            # o .Value do Excel arredonda ao segundo; o texto exibido tambem: comparar o texto com o valor
            if not (RX_DH if fmt == DH else RX_D).match(t) or (fmt == D and t != esp) or (fmt == DH and t[:10] != esp[:10]):
                ruins.append(f'{tab}[{c}] linha {i}: {t!r} (valor {esp})')
    for r in range(20, 32):
        ini = [w for w in wb.Worksheets if w.Name == 'Início'][0]
        if INICIO_NOVO in str(ini.Cells(r, 3).Formula):
            t = str(ini.Cells(r, 3).Text)
            if t != 'nenhum' and not RX_DH.match(t):
                ruins.append(f'Início!C{r}: {t!r}')
    if ruins:
        raise SystemExit('fumaca: ' + ' | '.join(ruins[:8]) + ' -- nada foi salvo')
    log('   ok: texto exibido = valor (dia/mes/ano hh:mm:ss) nas colunas de data-hora; so data nas de data')


def main(produto, caminho, salvar=True):
    try:
        sys.stdout.reconfigure(errors='replace')
    except Exception:                                   # noqa: BLE001
        pass
    caminho = os.path.abspath(caminho)
    ex = xlh.Excel()
    log(f'EXCEL_PID {ex.pid}')
    try:
        wb = ex.abrir(caminho)
        if wb.ReadOnly:
            raise SystemExit(f'arquivo aberto SOMENTE LEITURA (outra instancia o segura): {caminho}')
        estrutura = bool(wb.ProtectStructure)
        foto = foto_protecao(wb)
        ii.desproteger_tudo(wb)
        ex.xl.EnableEvents = False
        log('1. consulta QA_INTEGRACAO (pq/)')
        consultas(wb)
        log('2. VBA (todo o codigo atual das fontes: codigo_atual.py)')
        codigo_atual.instalar(ex, wb, produto, log)
        ex.xl.Calculation = -4135
        log('3. planilhas')
        colunas(ex.xl, wb)
        avulsas(ex.xl, wb)
        formulas(wb)
        log('4. fumaca')
        ex.xl.Calculation = -4105
        fumaca(ex, wb)
        ex.xl.EnableEvents = False
        devolver_protecao(wb, foto)
        if estrutura:
            wb.Protect(ii.SENHA, True, False)
        log(f'MUDANCAS {len(MUDANCAS)}' + (': ' + '; '.join(MUDANCAS) if MUDANCAS else ''))
        if salvar:
            antes = os.path.getmtime(wb.FullName)
            wb.Save()
            if os.path.getmtime(wb.FullName) == antes:
                raise SystemExit('o Save nao chegou ao arquivo (data de gravacao igual) -- conferir')
            log('salvo')
    finally:
        ex.fechar()


if __name__ == '__main__':
    a = sys.argv[1:]
    if len(a) < 2:
        raise SystemExit(__doc__)
    main(a[0], a[1], salvar='--sem-salvar' not in a)
