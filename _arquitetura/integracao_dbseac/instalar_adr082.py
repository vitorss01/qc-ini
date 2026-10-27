# -*- coding: utf-8 -*-
"""instalar_adr082.py -- ADR-082: o resultado MANUAL substitui um resultado do interfaceamento por ESCOLHA do usuario.

Uso:  python instalar_adr082.py <Bioquimica|Hematologia> <arquivo.xlsm> [--atualizar] [--sem-salvar]

Decisao do usuario (22/10/2026; revoga a D09b de 06/10 "manual nunca substitui automatico"): coluna opcional
SUBSTITUI_ID na aba Digitar Resultados. Preenchida com o ID de um resultado do interfaceamento com a mesma chave (e com
MOTIVO), o manual vale e o automatico sai da estatistica e do grafico (SUBSTITUIDO_MANUAL); quem decide e o motor de
dados (pq/DB_CQ_FINAL.m e src/comum/mEtlCQ.bas, iguais).

Idempotente (a 2a execucao registra 0 mudancas). So:
  1. tblResultados_Manuais: coluna SUBSTITUI_ID no FIM da tabela (formato texto, destravada, cabecalho no estilo das
     outras colunas de digitacao, mensagem de ajuda na celula) e os textos do topo da aba (camada_dados.ABAS);
  2. linha 5 da Principal - Resultados: contador "substituídos por manual" (camada_dados.formulas_resumo);
  3. consultas DB_CQ_FINAL e QA_INTEGRACAO regravadas de pq/ e conferidas (motor MOTOR_DADOS = PQ e a reserva);
  4. todo o VBA atual das fontes (codigo_atual.py: mEtlCQ/mEtlQA com a regra, mIntegracao com o preenchimento ao
     digitar o SUBSTITUI_ID e a trilha, mUI com a dica e o aviso da janela de inativar); EXIGE compilacao;
  5. fumaca: coluna e consultas conferidas de novo; mIntegracao.SubstituidoPeloManual/ManualQueSubstitui e
     mUI.EfeitoInativar/EfeitoReativar respondem
     sem erro; com --atualizar, o ATUALIZAR DADOS tem de concluir (conta SUBSTITUIDO_MANUAL e os achados E12/I04);
  6. devolve a protecao de cada aba e a da estrutura como estavam.
Sem retroatividade: manual antigo nao tem SUBSTITUI_ID, entao nada muda no historico (a tblCQ_Final nao precisa ser
refeita). Qualquer falha para ANTES de salvar.
"""
import os
import sys
import time

AQUI = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, AQUI)
sys.path.insert(0, os.path.join(AQUI, '..', 'etapa1_multilote'))
import xlh  # noqa: E402
import pqlib  # noqa: E402
import camada_dados as cd  # noqa: E402
import instalar_integracao as ii  # noqa: E402
from instalar_seguranca_usuarios import foto_protecao, devolver_protecao  # noqa: E402

TB_MAN = 'tblResultados_Manuais'
ABA_MAN = 'Digitar Resultados'
COL = 'SUBSTITUI_ID'
CONSULTAS = ['DB_CQ_FINAL', 'QA_INTEGRACAO']
MUDANCAS = []


def log(*a):
    print(*a, flush=True)


def normalizar(t):
    return '\n'.join(l.rstrip() for l in str(t).replace('\r\n', '\n').replace('\r', '\n').strip().split('\n'))


def norm_f(f):
    return str(f or '').replace('_xlfn.', '').replace(' ', '').upper()


def definicao_coluna():
    """(largura, formato, tipo) da SUBSTITUI_ID em camada_dados.ENTRADAS -- uma fonte so para instalacao nova e esta."""
    for h, w, fmt, tipo in cd.ENTRADAS[TB_MAN][1]:
        if h == COL:
            return w, fmt, tipo
    raise SystemExit(f'camada_dados.ENTRADAS sem a coluna {COL}')


def estilo_cab(c):
    return (c.Interior.Color, bool(c.Font.Bold), float(c.Font.Size), c.Font.Color, bool(c.WrapText), c.VerticalAlignment)


def coluna(wb):
    """Passo 1: a coluna SUBSTITUI_ID no fim da tblResultados_Manuais, formatada como as de digitacao."""
    lo = pqlib.tabela(wb, TB_MAN)
    if lo is None:
        raise SystemExit(f'{TB_MAN} nao encontrada: instale antes a camada de dados (instalar_integracao.py)')
    ws = lo.Parent
    largura, fmt, tipo = definicao_coluna()
    nomes = [c.Name for c in lo.ListColumns]
    if COL not in nomes:
        lo.ListColumns.Add().Name = COL                        # ListColumns.Add() sem posicao = no fim
        MUDANCAS.append(f'coluna {COL} (fim da tabela)')
        log(f'   {COL}: coluna criada no fim da tabela')
    elif nomes[-1] != COL:
        # nunca move coluna com digitacao do usuario: so registra
        log(f'   {COL}: ja existia (posicao {nomes.index(COL) + 1} de {len(nomes)}) -- mantida onde esta')
    col = lo.ListColumns(COL)
    rng = col.DataBodyRange
    if rng.NumberFormat != fmt:
        rng.NumberFormat = fmt
        MUDANCAS.append(f'{COL}: formato texto')
    if rng.Locked is not False:
        rng.Locked = False                                      # digitado: destravado com a aba protegida
        MUDANCAS.append(f'{COL}: destravada')
    cx = ws.Columns(rng.Column)
    if abs(float(cx.ColumnWidth) - largura) > 0.01:
        cx.ColumnWidth = largura
        MUDANCAS.append(f'{COL}: largura')
    # cabecalho no estilo das outras colunas de digitacao (a referencia e a DATA, tipo 'in')
    hr = lo.HeaderRowRange
    cab_ref = hr.Cells(1, [c.Name for c in lo.ListColumns].index('DATA') + 1)
    cab_novo = hr.Cells(1, [c.Name for c in lo.ListColumns].index(COL) + 1)
    if estilo_cab(cab_novo) != estilo_cab(cab_ref) or cab_novo.Interior.Color != cd.tema.ACAO:
        tipos = {h: (t if t in ('in', 'auto') else 'f') for h, _, _, t in cd.ENTRADAS[TB_MAN][1]}
        cd.formatar_cabecalho_tabela(lo, [tipos.get(c.Name, 'f') for c in lo.ListColumns])
        MUDANCAS.append(f'{COL}: cabecalho')
    # mensagem de ajuda (Validation de so mensagem, como a da aba Inativar)
    titulo, msg = cd.DICA_SUBST
    try:
        v = rng.Cells(1, 1).Validation
        atual = (str(v.InputTitle), str(v.InputMessage), bool(v.ShowInput))
        ultima = rng.Cells(rng.Rows.Count, 1).Validation
        atual_fim = (str(ultima.InputTitle), str(ultima.InputMessage))
    except Exception:                                           # noqa: BLE001 -- sem validacao na celula
        atual, atual_fim = None, None
    if atual != (titulo, msg, True) or atual_fim != (titulo, msg):
        rng.Validation.Delete()
        rng.Validation.Add(0, 1, 1)                             # xlValidateInputOnly
        rng.Validation.InputTitle = titulo
        rng.Validation.InputMessage = msg
        rng.Validation.ShowInput = True
        MUDANCAS.append(f'{COL}: mensagem de ajuda')
    return lo


def textos(wb):
    """Passo 1 (fim) e 2: textos do topo da Digitar Resultados e o contador da Principal - Resultados."""
    ws = [w for w in wb.Worksheets if w.Name == ABA_MAN][0]
    _, tit, sub, instr = [a for a in cd.ABAS if a[0] == ABA_MAN][0]
    esperado = [tit, sub] + list(instr[:2])
    atual = [str(ws.Range(f'A{i}').Value or '') for i in range(1, 5)]
    if atual != esperado:
        cd.textos_topo(wb, ABA_MAN)
        MUDANCAS.append(f'{ABA_MAN}: textos do topo')
    nome = 'Principal - Resultados'
    f = cd.formulas_resumo()[nome]
    wp = [w for w in wb.Worksheets if w.Name == nome][0]
    if norm_f(wp.Range('A5').Formula2) != norm_f(f):
        wp.Range('A5').Formula2 = f
        MUDANCAS.append(f'{nome}: contador da linha 5')


def consultas(wb):
    """Passo 3: DB_CQ_FINAL e QA_INTEGRACAO identicas a pq/."""
    wb.Queries.FastCombine = True
    for q in CONSULTAS:
        atual = pqlib.query_existe(wb, q)
        if atual is None:
            raise SystemExit(f'consulta {q} nao existe no arquivo: instale antes a camada de dados (instalar_integracao.py)')
        if normalizar(atual.Formula) != normalizar(pqlib.m_de(q)):
            pqlib.gravar_query(wb, q, pqlib.m_de(q))
            MUDANCAS.append(f'consulta {q}')
    conferir_consultas(wb)


def conferir_consultas(wb):
    for q in CONSULTAS:
        if normalizar(pqlib.query_existe(wb, q).Formula) != normalizar(pqlib.m_de(q)):
            raise SystemExit(f'consulta {q}: formula no arquivo difere da fonte depois de gravar')


def fumaca(ex, wb, atualizar):
    """Passo 5: confere o que foi feito (resultado, nao intencao)."""
    lo = pqlib.tabela(wb, TB_MAN)
    nomes = [c.Name for c in lo.ListColumns]
    rng = lo.ListColumns(COL).DataBodyRange
    if COL not in nomes or rng.NumberFormat != '@' or rng.Locked is not False:
        raise SystemExit(f'fumaca: coluna {COL} ausente, fora do formato texto ou travada')
    conferir_consultas(wb)
    ex.xl.EnableEvents = True
    try:
        r1 = ex.run(f"'{wb.Name}'!mIntegracao.SubstituidoPeloManual", 'MAN_0000', teto=120)
        r2 = ex.run(f"'{wb.Name}'!mUI.EfeitoInativar", 'MAN_0000', teto=120)
        # revisao do ADR-082: a janela de reativar tambem avisa (X de um automatico substituido continua fora)
        r3 = ex.run(f"'{wb.Name}'!mIntegracao.ManualQueSubstitui", '999999999', teto=120)
        r4 = ex.run(f"'{wb.Name}'!mUI.EfeitoReativar", '999999999', teto=120)
        if any(str(r or '') != '' for r in (r1, r2, r3, r4)):
            raise SystemExit(f'fumaca: ID inexistente nao deveria substituir nada ({r1!r}, {r2!r}, {r3!r}, {r4!r})')
        log('   mIntegracao.SubstituidoPeloManual/ManualQueSubstitui e mUI.EfeitoInativar/EfeitoReativar respondem')
        if atualizar:
            ex.xl.Calculation = -4105
            t0 = time.time()
            r = str(ex.run(f"'{wb.Name}'!mIntegracao.AtualizarDadosAutomatico", teto=1500))
            if not r.startswith('OK|') or 'concluída' not in r:
                raise SystemExit('ATUALIZAR DADOS falhou: ' + r[:800])
            fin = pqlib.ler_tabela(pqlib.tabela(wb, 'tblCQ_Final'))
            qa = pqlib.ler_tabela(pqlib.tabela(wb, 'tblQA_Integracao'))
            n_sub = sum(1 for x in fin if x.get('STATUS_ANALITICO') == 'SUBSTITUIDO_MANUAL')
            log(f'   ATUALIZAR DADOS OK ({time.time() - t0:.1f}s): {n_sub} SUBSTITUIDO_MANUAL; '
                f'E12 = {sum(1 for x in qa if x.get("CODIGO") == "E12")}, '
                f'I04 = {sum(1 for x in qa if x.get("CODIGO") == "I04")}')
    finally:
        ex.xl.EnableEvents = False


def main(produto, caminho, atualizar=False, salvar=True):
    try:
        sys.stdout.reconfigure(errors='replace')
    except Exception:                                           # noqa: BLE001
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

        log(f'1. {TB_MAN}: coluna {COL}')
        coluna(wb)
        log(f'2. textos do topo da aba {ABA_MAN} e contador "substituídos por manual" da Principal - Resultados')
        textos(wb)
        log(f'3. consultas {", ".join(CONSULTAS)} (pq/)')
        consultas(wb)
        log('   identicas a pq/')
        log('4. VBA (todo o codigo atual das fontes: codigo_atual.py)')
        import codigo_atual
        codigo_atual.instalar(ex, wb, produto, log)
        log('5. fumaca')
        fumaca(ex, wb, atualizar)
        log(f'MUDANCAS {len(MUDANCAS)}' + (': ' + '; '.join(MUDANCAS) if MUDANCAS else ''))
        devolver_protecao(wb, foto)
        if estrutura:
            wb.Protect(ii.SENHA, True, False)
        if salvar:
            wb.Save()
            log('salvo')
    finally:
        ex.fechar()


if __name__ == '__main__':
    a = sys.argv[1:]
    main(a[0], a[1], atualizar='--atualizar' in a, salvar='--sem-salvar' not in a)
