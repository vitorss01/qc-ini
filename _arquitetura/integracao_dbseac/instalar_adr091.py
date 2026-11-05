# -*- coding: utf-8 -*-
"""instalar_adr091.py -- ADR-091 (05/11/2026): regras de material por LOTE (tblLotesMaterial).

Uso:  python instalar_adr091.py <Bioquimica|Hematologia> <arquivo.xlsm> [--atualizar] [--sem-salvar]

Idempotente (a 2a execucao registra 0 mudancas). Passos:
  1. consultas DB_CQ_FINAL e QA_INTEGRACAO regravadas de pq/ e conferidas (MATRIZ do lote antes do de/para,
     LOTE_FORA_DO_CQ / LOTE_OUTRO_MATERIAL, A13, A14 -- motor reserva PQ);
  2. todo o VBA atual das fontes (codigo_atual.py): mEtlCQ (LerLotesMaterial, as mesmas regras) e mEtlQA (A13/A14);
     EXIGE compilacao;
  3. tblLotesMaterial (LOTE, REGRA, MATRIZ, ANALITOS, NOTA) na aba Cfg_Integracao, com a carga inicial de
     config_produtos.py (Bio: 8353 SO_ANALITOS PCR; 6353/6354/5659 MATRIZ URINA; 8177/8794 FORA_DO_CQ. Hema: vazia).
     Ja existe: NADA e reescrito (a tabela e do usuario); so confere as colunas;
  4. fumaca: consultas conferidas de novo; mEtlCQ.LotesMaterialResumo responde e nao tem aviso (A14) na carga
     inicial; com --atualizar, o ATUALIZAR so das entradas conclui e: LOTE_* = achados A13 (mesmos IDs); nenhum
     participante nos lotes FORA_DO_CQ nem fora da lista do SO_ANALITOS; os lotes da regra MATRIZ so com essa matriz;
     A14 = 0;
  5. devolve a protecao de cada aba e a da estrutura como estavam; salva com os eventos desligados e confere a data de
     gravacao. Qualquer falha para ANTES de salvar.
Sem --atualizar, a tblCQ_Final gravada continua a de antes ate o proximo ATUALIZAR (que aplica as regras).
"""
import os
import sys
import time

AQUI = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, AQUI)
sys.path.insert(0, os.path.join(AQUI, '..', 'etapa1_multilote'))
import xlh  # noqa: E402
import pqlib  # noqa: E402
import instalar_integracao as ii  # noqa: E402
import codigo_atual  # noqa: E402
import config_produtos as cp  # noqa: E402
import camada_dados as cd  # noqa: E402
from instalar_seguranca_usuarios import foto_protecao, devolver_protecao  # noqa: E402

MUDANCAS = []
CONSULTAS = ['DB_CQ_FINAL', 'QA_INTEGRACAO']
TB = 'tblLotesMaterial'
ABA = 'Cfg_Integracao'
REGRAS = ('MATRIZ', 'SO_ANALITOS', 'FORA_DO_CQ')
STATUS_LOTE = ('LOTE_FORA_DO_CQ', 'LOTE_OUTRO_MATERIAL')
NOTA_ABA = ('REGRAS DE MATERIAL POR LOTE (ADR-091): MATRIZ = o lote usa esta matriz (antes do de/para) | '
            'SO_ANALITOS = so participa nos analitos listados (separe com ;) | FORA_DO_CQ = o lote inteiro fica fora. '
            'Editar e ATUALIZAR; linha com problema aparece no QA (A14).')


def log(m):
    print(m, flush=True)


def normalizar(t):
    return '\n'.join(l.rstrip() for l in str(t).replace('\r\n', '\n').replace('\r', '\n').strip().split('\n'))


def up(v):
    if isinstance(v, float) and v == int(v):
        v = int(v)
    t = str(v if v is not None else '').strip()
    return t.upper() or None


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


def destino_livre(ws, ncol):
    """1a coluna, a partir de L (depois do de/para em E:J), com ncol colunas vazias da linha 1 ate o fim do usado."""
    ult = ws.UsedRange.Row + ws.UsedRange.Rows.Count + 50
    col = 12
    while col < 200:
        rng = ws.Range(ws.Cells(1, col), ws.Cells(ult, col + ncol))     # + 1 coluna de folga a direita
        if int(ws.Application.WorksheetFunction.CountA(rng)) == 0:
            return col
        col += 1
    raise SystemExit('Cfg_Integracao sem espaco livre para a tblLotesMaterial -- nada foi salvo')


def tabela_lotes(wb, produto):
    lo = pqlib.tabela(wb, TB)
    if lo is not None:
        cab = [str(c.Name) for c in lo.ListColumns]
        falta = [c for c in cp.COLUNAS_LOTES_MATERIAL if c not in cab]
        if falta:
            raise SystemExit(f'{TB} existe sem as colunas {falta} -- nada foi salvo')
        log(f'   {TB} existia ({lo.ListRows.Count} linha(s), preservada)')
        return lo
    ws = wb.Worksheets(ABA)
    linhas = cp.linhas_lotes_material(produto)
    ncol = len(cp.COLUNAS_LOTES_MATERIAL)
    col = destino_livre(ws, ncol)
    r0 = cd.LINHA_CAB
    # LOTE e texto (nunca numero: '08353' ficaria 8353); o resto tambem, para o que o usuario digitar
    ws.Range(ws.Cells(r0 + 1, col), ws.Cells(r0 + max(1, len(linhas)), col + ncol - 1)).NumberFormat = '@'
    lo = pqlib.criar_tabela_dados(ws, TB, cp.COLUNAS_LOTES_MATERIAL, linhas, ws.Cells(r0, col).Address)
    cd.formatar_cabecalho_tabela(lo)
    for j, w in enumerate((10, 14, 12, 34, 60)):
        ws.Columns(col + j).ColumnWidth = w
    lo.DataBodyRange.Locked = False             # editavel com a aba protegida (o administrador mantem as regras)
    nota = ws.Cells(r0 - 1, col)
    if nota.Value in (None, ''):
        nota.Value = NOTA_ABA
        nota.WrapText = False
        nota.Font.Size = 8
        nota.Font.Italic = True
    MUDANCAS.append(f'{TB} criada em {ABA}!{ws.Cells(r0, col).Address} ({len(linhas)} regra(s))')
    return lo


def conferir_carga(wb, produto):
    """A tabela recem-criada tem exatamente a carga inicial; a existente so e listada (pode ter sido editada)."""
    lo = pqlib.tabela(wb, TB)
    reg = [r for r in pqlib.ler_tabela(lo) if any(r.get(c) not in (None, '') for c in ('LOTE', 'REGRA', 'MATRIZ', 'ANALITOS'))]
    log('   regras: ' + ('; '.join(f'{r["LOTE"]} {r["REGRA"]} {r.get("MATRIZ") or ""}{r.get("ANALITOS") or ""}'.strip()
                                    for r in reg) or '(nenhuma)'))
    return reg


def fumaca(ex, wb, produto, atualizar, recem_criada):
    conferir_consultas(wb)
    nome = wb.Name
    res = str(ex.run(f"'{nome}'!mEtlCQ.LotesMaterialResumo", teto=120))
    log(f'   LotesMaterialResumo = {res[:300]}')
    contagem, _, avisos = res.partition('|')
    regras = conferir_carga(wb, produto)
    if recem_criada:
        esperado = cp.linhas_lotes_material(produto)
        n = {k: len({l[0] for l in esperado if l[1] == k}) for k in REGRAS}
        if contagem != f'MATRIZ={n["MATRIZ"]};FORA_DO_CQ={n["FORA_DO_CQ"]};SO_ANALITOS={n["SO_ANALITOS"]}' or avisos:
            raise SystemExit(f'fumaca: carga inicial lida como {res[:300]!r} (esperado {n}, sem avisos) -- nada foi salvo')
        if [[up(r[c]) for c in ('LOTE', 'REGRA', 'MATRIZ', 'ANALITOS')] for r in regras] != \
                [[up(x) for x in l[:4]] for l in esperado]:
            raise SystemExit('fumaca: a tblLotesMaterial gravada difere da carga inicial -- nada foi salvo')
    elif avisos:
        log(f'   ATENCAO: a tblLotesMaterial existente tem linha(s) ignorada(s)/suspeita(s) (QA A14): {avisos[:300]}')
    if not atualizar:
        return
    ex.xl.EnableEvents = True
    try:
        t0 = time.time()
        r = str(ex.run(f"'{nome}'!mIntegracao.AtualizarSoEntradas", teto=1500))
    finally:
        ex.xl.EnableEvents = False
    if not r.startswith('OK|'):
        raise SystemExit('ATUALIZAR (so entradas) falhou: ' + r[:600])
    fin = pqlib.ler_tabela(pqlib.tabela(wb, 'tblCQ_Final'))
    qa = pqlib.ler_tabela(pqlib.tabela(wb, 'tblQA_Integracao'))
    lote_st = {x['ID_REGISTRO'] for x in fin if x.get('STATUS_ANALITICO') in STATUS_LOTE}
    a13 = {x['ID_REGISTRO'] for x in qa if x.get('CODIGO') == 'A13'}
    a14 = [x for x in qa if x.get('CODIGO') == 'A14']
    mz = {up(x['LOTE']): up(x['MATRIZ']) for x in regras if up(x['REGRA']) == 'MATRIZ' and up(x['MATRIZ'])}
    fora = {up(x['LOTE']) for x in regras if up(x['REGRA']) == 'FORA_DO_CQ'}
    so = {}
    for x in regras:
        if up(x['REGRA']) == 'SO_ANALITOS':
            so.setdefault(up(x['LOTE']), set()).update(a for a in (up(p) for p in str(x['ANALITOS'] or '').split(';')) if a)
    ruins = []
    for x in fin:
        lt = up(x.get('LOTE'))
        if lt in mz and up(x.get('MATRIZ')) != mz[lt]:
            ruins.append(('MATRIZ', x['ID_REGISTRO'], x.get('MATRIZ')))
        if x.get('PARTICIPA_ESTATISTICA') == 'SIM' and (lt in fora or (lt in so and up(x.get('ANALITO')) not in so[lt])):
            ruins.append(('PARTICIPA', x['ID_REGISTRO'], x.get('ANALITO')))
    cont = {s: sum(1 for x in fin if x.get('STATUS_ANALITICO') == s) for s in STATUS_LOTE}
    log(f'   ATUALIZAR so entradas OK ({time.time() - t0:.1f}s): {cont}, A13 {len(a13)}, A14 {len(a14)}, '
        f'violacoes {len(ruins)} {ruins[:3]}')
    if lote_st != a13 or ruins or (recem_criada and a14):
        raise SystemExit('fumaca: LOTE_* != A13, regra nao aplicada ou A14 na carga inicial -- nada foi salvo')


def main(produto, caminho, atualizar=False, salvar=True):
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
        log('1. consultas DB_CQ_FINAL e QA_INTEGRACAO (pq/)')
        consultas(wb)
        log('2. VBA (todo o codigo atual das fontes: codigo_atual.py)')
        codigo_atual.instalar(ex, wb, produto, log)
        ex.xl.Calculation = -4135
        log(f'3. {TB} ({ABA})')
        recem = pqlib.tabela(wb, TB) is None
        tabela_lotes(wb, produto)
        log('4. fumaca')
        fumaca(ex, wb, produto, atualizar, recem)
        ex.xl.EnableEvents = False
        ex.xl.Calculation = -4105
        devolver_protecao(wb, foto)
        if estrutura:
            wb.Protect(ii.SENHA, True, False)
        log(f'MUDANCAS {len(MUDANCAS)}' + (': ' + '; '.join(MUDANCAS) if MUDANCAS else ''))
        if salvar:
            # eventos desligados: com eles o Workbook_BeforeSave cancela o Save e engole o erro
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
    main(a[0], a[1], atualizar='--atualizar' in a, salvar='--sem-salvar' not in a)
