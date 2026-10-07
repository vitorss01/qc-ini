# -*- coding: utf-8 -*-
"""instalar_adr070.py -- ADR-070: aba Inativar simples (ID + ANALITO + MOTIVO) e ID na dica do Levey-Jennings.

Uso:  python instalar_adr070.py <Bioquimica|Hematologia> <arquivo.xlsm> [--sem-salvar]

Idempotente (rodar de novo nao muda nada). So:
  1. instala todo o VBA atual das fontes (codigo_atual.py -- inclui a clsCht, agora versionada); EXIGE compilacao;
  2. regrava as consultas DB_CQ_FINAL e QA_INTEGRACAO a partir de pq/ e confere;
  3. aba Inativar: migra a tblInativacao_NaoConformes para o layout novo (camada_dados.montar_entrada):
     ID_REGISTRO | ANALITO | REGISTRAR - LJ | MOTIVO | DATA_INATIVACAO | USUARIO, sem as 8 colunas de formula.
     Toda linha existente fica, na mesma posicao, com ID, caixa, data e usuario; o ANALITO das linhas antigas
     vem da tblCQ_Final pelo ID; o MOTIVO fica vazio (a justificativa antiga continua valendo pelos
     COMENTARIOS_TECNICOS). Conferido linha a linha;
  4. Digitar Resultados e COMENTARIOS_TECNICOS: as colunas de formula (formato '@' guardava a formula como
     TEXTO) voltam a calcular; conferido que nenhuma celula de formula mostra "=...";
  5. textos do topo das abas Inativar, COMENTARIOS_TECNICOS e Principal - Resultados;
  6. Eng_Saida: cabecalhos das colunas novas (AM..BJ = ID e DATA_HORA de cada ponto do LJ);
  7. MODO_FONTE = HISTORICO: ATUALIZAR DADOS e confere que o conjunto de INATIVADOS (e a plotagem de cada um) nao
     mudou e que nenhuma linha migrada caiu em E10/E11. Em SEAC a consulta nova vale no proximo ATUALIZAR DADOS
     (como no ADR-058: o instalador nao le a rede); roda so o motor do Painel;
  8. fumaca da dica: mUI.DicaPonto de um ponto do grafico N1 devolve "ID ... RUN ..." com o ID e o RUN do Eng_Saida;
  9. devolve a protecao de cada aba e a da estrutura como estavam; salva.
Qualquer falha para ANTES de salvar.
"""
import os
import sys
import time
import unicodedata

AQUI = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, AQUI)
sys.path.insert(0, os.path.join(AQUI, '..', 'etapa1_multilote'))
import xlh  # noqa: E402
import pqlib  # noqa: E402
import camada_dados as cd  # noqa: E402
import instalar_integracao as ii  # noqa: E402
from instalar_seguranca_usuarios import foto_protecao, devolver_protecao  # noqa: E402

TB = 'tblInativacao_NaoConformes'
CONSULTAS = ['DB_CQ_FINAL', 'QA_INTEGRACAO']
CAB_ENG = (['N1 ID', 'N2 ID', 'N3 ID']
           + [f'N{n} X{s} ID' for n in (1, 2, 3) for s in (1, 2, 3)]
           + ['N1 DATA_HORA', 'N2 DATA_HORA', 'N3 DATA_HORA']
           + [f'N{n} X{s} DATA_HORA' for n in (1, 2, 3) for s in (1, 2, 3)])
COL_ENG0 = 39                     # AM (mEstatistica.ENG_COL_ID0)


def log(*a):
    print(*a, flush=True)


def normalizar(t):
    return '\n'.join(l.rstrip() for l in str(t).replace('\r\n', '\n').replace('\r', '\n').strip().split('\n'))


def coluna(lo, nome):
    n = lo.ListRows.Count
    if not n:
        return []
    v = lo.ListColumns(nome).DataBodyRange.Value2
    return [r[0] for r in v] if n > 1 else [v]


def modo_fonte(wb):
    """MODO_FONTE do CFG sem acento (o mIntegracao aceita 'Histórico')."""
    for r in pqlib.ler_tabela(pqlib.tabela(wb, 'tblConfigIntegracao')):
        if str(r['CHAVE']).strip().upper() == 'MODO_FONTE':
            t = unicodedata.normalize('NFKD', str(r['VALOR'] or '')).encode('ascii', 'ignore').decode().strip().upper()
            return t or 'SEAC'
    return 'SEAC'


def inativados(wb):
    lo = pqlib.tabela(wb, 'tblCQ_Final')
    ids, st, pl = coluna(lo, 'ID_REGISTRO'), coluna(lo, 'STATUS_ANALITICO'), coluna(lo, 'TIPO_PLOTAGEM_LJ')
    return {i: p for i, s, p in zip(ids, st, pl) if s == 'INATIVADO'}


def conferir_migracao(lo, mig):
    """Toda linha antiga continua na mesma posicao com ID, caixa, data, usuario e o ANALITO preenchido."""
    if not mig.get('migrada'):
        return 'tabela ja no layout novo (nada migrado)'
    novo = {h: coluna(lo, h) for h in ('ID_REGISTRO', 'ANALITO', 'REGISTRAR - LJ', 'MOTIVO', 'DATA_INATIVACAO', 'USUARIO')}
    ruins = []
    for k, antiga in enumerate(mig['antigas']):
        for h, v in antiga.items():
            atual = novo[h][k]
            if (v in (None, '') and atual in (None, '')) or v == atual:
                continue
            if isinstance(v, float) and isinstance(atual, float) and abs(v - atual) < 1e-9:
                continue
            ruins.append((k + 1, h, v, atual))
        if novo['MOTIVO'][k] not in (None, ''):
            ruins.append((k + 1, 'MOTIVO', None, novo['MOTIVO'][k]))
    if ruins:
        raise SystemExit(f'migracao da {TB} divergiu (linha, coluna, antes, depois): {ruins[:10]}')
    return (f"{mig['linhas']} linha(s) preservada(s) na mesma posicao, {mig['ids']} com ID; "
            f"ANALITO preenchido pela tblCQ_Final em {mig['analito_preenchido']}")


def conferir_formulas(wb):
    """Nenhuma celula de formula das abas de entrada mostra o texto da formula (o defeito do formato '@')."""
    ruins, n = [], 0
    for nome, (_aba, cols) in cd.ENTRADAS.items():
        lo = pqlib.tabela(wb, nome)
        for h, _w, _f, tipo in cols:
            if tipo in ('in', 'auto'):
                continue
            rng = lo.ListColumns(h).DataBodyRange
            for i in (1, 2, rng.Rows.Count):
                cel = rng.Cells(i, 1)
                n += 1
                v = cel.Value
                if not cel.HasFormula or (isinstance(v, str) and v.startswith('=')):
                    ruins.append((nome, h, i, str(v)[:60]))
    if ruins:
        raise SystemExit(f'colunas de formula sem calcular: {ruins[:10]}')
    lo = pqlib.tabela(wb, TB)
    if any(c.DataBodyRange.Cells(1, 1).HasFormula for c in lo.ListColumns):
        raise SystemExit(f'{TB} ainda tem coluna de formula')
    return n


def main(produto, caminho, salvar=True):
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
        modo = modo_fonte(wb)
        antes = inativados(wb)
        log(f'MODO_FONTE = {modo}; inativados na tblCQ_Final antes: {len(antes)}')

        log('1. VBA (todo o codigo atual das fontes: codigo_atual.py)')
        import codigo_atual
        codigo_atual.instalar(ex, wb, produto, log)

        log('2. consultas')
        wb.Queries.FastCombine = True
        for q in CONSULTAS:
            pqlib.gravar_query(wb, q, pqlib.m_de(q))
        for q in CONSULTAS:
            if normalizar(pqlib.query_existe(wb, q).Formula) != normalizar(pqlib.m_de(q)):
                raise SystemExit(f'consulta {q}: formula no arquivo difere da fonte depois de gravar')
        log(f'   {", ".join(CONSULTAS)}: identicas a pq/')

        log('3. aba Inativar: layout novo (ID | ANALITO | REGISTRAR - LJ | MOTIVO | DATA_INATIVACAO | USUARIO)')
        mig = cd.migrar_inativacao(wb, pqlib.tabela(wb, TB), cd.ENTRADAS[TB][1])   # no layout novo: nada
        lo = cd.montar_entrada(wb, produto, TB)            # formatos, travas, validacoes, caixa de selecao
        if mig.get('migrada'):
            cd.igualar_largura_cabecalho(lo.Parent, len(cd.ENTRADAS[TB][1]), mig['largura_cab'])
        log('   ' + conferir_migracao(lo, mig))

        log('4. Digitar Resultados e COMENTARIOS_TECNICOS: colunas de formula calculando')
        for nome in ('tblResultados_Manuais', 'tblComentariosTecnicos'):
            cd.montar_entrada(wb, produto, nome)
        log(f'   {conferir_formulas(wb)} celulas de formula conferidas (nenhuma mostra "=..."); Inativar sem formula')

        log('5. textos do topo')
        for nome in ('Inativar', 'COMENTARIOS_TECNICOS', 'Principal - Resultados'):
            cd.textos_topo(wb, nome)
        log('   Inativar: ' + str(wb.Worksheets('Inativar').Range('A3').Value))

        log('6. Eng_Saida: colunas AM..BJ (ID e DATA_HORA de cada ponto)')
        eng = wb.Worksheets('Eng_Saida')
        eng.Range(eng.Cells(2, COL_ENG0), eng.Cells(2, COL_ENG0 + len(CAB_ENG) - 1)).Value = tuple([tuple(CAB_ENG)])
        eng.Range(eng.Cells(1, COL_ENG0), eng.Cells(1, COL_ENG0)).Value = \
            'DE QUEM E O PONTO: ID_REGISTRO e DATA_HORA de cada ponto e de cada X (ADR-070) -- dica do grafico'

        ex.xl.EnableEvents = True
        ex.xl.Calculation = -4105
        if modo == 'HISTORICO':
            log('7. ATUALIZAR DADOS em modo historico')
            t0 = time.time()
            r = str(ex.run("'" + wb.Name + "'!mIntegracao.AtualizarDadosAutomatico", teto=2400))
            if not r.startswith('OK|') or 'concluída' not in r:
                raise SystemExit('ATUALIZAR DADOS falhou: ' + r[:800])
            depois = inativados(wb)
            if depois != antes:
                so_a = sorted(set(antes) - set(depois))[:10]
                so_d = sorted(set(depois) - set(antes))[:10]
                mud = [(i, antes[i], depois[i]) for i in set(antes) & set(depois) if antes[i] != depois[i]][:10]
                raise SystemExit(f'o conjunto de INATIVADOS mudou com a migracao: saiu {so_a}, entrou {so_d}, plotagem {mud}')
            qa = pqlib.ler_tabela(pqlib.tabela(wb, 'tblQA_Integracao'))
            e1011 = [x for x in qa if x['CODIGO'] in ('E10', 'E11')]
            if e1011:
                raise SystemExit(f'linhas da Inativar caindo em E10/E11 depois da migracao: {e1011[:5]}')
            log(f'   OK ({time.time() - t0:.0f}s): {len(depois)} inativado(s), mesmos IDs e mesma plotagem; sem E10/E11')
        else:
            log('7. MODO_FONTE = SEAC: a consulta nova vale no proximo ATUALIZAR DADOS; motor do Painel')
            ex.run("'" + wb.Name + "'!mEstatistica.InvalidarCache", teto=300)
            ex.run("'" + wb.Name + "'!mEstatistica.AtualizarCalc", teto=900)

        log('8. fumaca da dica do grafico (mUI.DicaPonto)')
        log('   ' + fumaca_dica(ex, wb))
        ex.xl.EnableEvents = False

        devolver_protecao(wb, foto)
        if estrutura:
            wb.Protect(ii.SENHA, True, False)
        if salvar:
            wb.Save()
            log('salvo')
    finally:
        ex.fechar()


def fumaca_dica(ex, wb):
    eng = wb.Worksheets('Eng_Saida')
    n = int(eng.Range('I1').Value or 0)
    if n == 0:
        return 'Painel sem corridas no motor: nada a conferir (a dica e conferida pelo qa_inativar)'
    ch = wb.Worksheets('Painel').ChartObjects(1).Chart
    nomes = [str(s.Name) for s in ch.SeriesCollection()]
    serie = nomes.index('Resultado') + 1
    for slot in range(n, 0, -1):
        v = eng.Cells(2 + slot, 25).Value
        if isinstance(v, float):
            idv = str(eng.Cells(2 + slot, COL_ENG0).Value or '')
            run = eng.Cells(2 + slot, 2).Value
            d = str(ex.run("'" + wb.Name + "'!mUI.DicaPonto", 1, serie, slot, teto=120))
            curto = idv.split('-')[-1]
            run_txt = str(int(run)) if isinstance(run, float) else str(run)
            if not idv or not d.startswith(f'ID {curto} ') or f'RUN {run_txt}' not in d:
                raise SystemExit(f'DicaPonto(1, {serie}, {slot}) = {d!r}; esperado ID {curto} e RUN {run_txt}')
            return f'DicaPonto(N1, ponto {slot}) = {d}'
    return 'nenhum ponto N1 com valor na janela do motor'


if __name__ == '__main__':
    a = sys.argv[1:]
    main(a[0], a[1], salvar='--sem-salvar' not in a)
