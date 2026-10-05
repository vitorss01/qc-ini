# -*- coding: utf-8 -*-
"""instalar_seguranca_usuarios.py -- ADR-059: papeis e sessao num QC_INI ja integrado (ADR-057/058).

Uso:  python instalar_seguranca_usuarios.py <Bioquimica|Hematologia> <arquivo.xlsm> [--sem-salvar]

Incremental. So:
  1. troca o modulo mSeguranca pelo de src/<produto>/ e compila;
  2. aba Usuarios: TRAVA as celulas da sessao (currentUser/currentPapel) e a tabela de
     usuarios (A4:E53: login, nome, funcao, hash, rubrica) -- so o VBA escreve nelas;
     DESTRAVA so a area de cadastro (cadLogin/cadNome/cadSenha/cadPapel); OCULTA a
     coluna do hash (ninguem precisa ler hash; com a aba protegida, so o ADM reexibe);
  3. limpa a sessao e o formulario de cadastro: o arquivo nao e distribuido com a
     identidade de quem o usou por ultimo;
  4. devolve a protecao de CADA aba exatamente como estava (conteudo, objetos, cenarios,
     filtro, ordenacao, formatacao) e a da estrutura -- a Audit_Log precisa de filtro e
     ordenacao para o auditor, e um Protect uniforme tiraria isso sem ninguem notar;
  5. CONFERE: texto do modulo identico a fonte; travas e coluna oculta lidas de volta.
Qualquer falha para ANTES de salvar.
"""
import os
import sys

AQUI = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, AQUI)
sys.path.insert(0, os.path.join(AQUI, '..', 'etapa1_multilote'))
import xlh  # noqa: E402
import instalar_integracao as ii  # noqa: E402
from instalar_modo_historico import normalizar  # noqa: E402

TRAVAR = ['N1:N2', 'A4:E53']
CADASTRO = ['cadLogin', 'cadNome', 'cadSenha', 'cadPapel']
COL_HASH = 4


def log(*a):
    print(*a, flush=True)


def foto_protecao(wb):
    """Estado de protecao de cada aba, para devolver exatamente como estava."""
    out = {}
    for ws in wb.Worksheets:
        pr = ws.Protection
        out[ws.Name] = dict(
            conteudo=bool(ws.ProtectContents), objetos=bool(ws.ProtectDrawingObjects), cenarios=bool(ws.ProtectScenarios),
            AllowFormattingCells=pr.AllowFormattingCells, AllowFormattingColumns=pr.AllowFormattingColumns,
            AllowFormattingRows=pr.AllowFormattingRows, AllowInsertingColumns=pr.AllowInsertingColumns,
            AllowInsertingRows=pr.AllowInsertingRows, AllowInsertingHyperlinks=pr.AllowInsertingHyperlinks,
            AllowDeletingColumns=pr.AllowDeletingColumns, AllowDeletingRows=pr.AllowDeletingRows,
            AllowSorting=pr.AllowSorting, AllowFiltering=pr.AllowFiltering, AllowUsingPivotTables=pr.AllowUsingPivotTables)
    return out


def devolver_protecao(wb, foto):
    for ws in wb.Worksheets:
        f = foto.get(ws.Name)
        if not f or not (f['conteudo'] or f['objetos'] or f['cenarios']):
            continue
        # Protect(Password, DrawingObjects, Contents, Scenarios, UserInterfaceOnly, AllowFormattingCells,
        #         AllowFormattingColumns, AllowFormattingRows, AllowInsertingColumns, AllowInsertingRows,
        #         AllowInsertingHyperlinks, AllowDeletingColumns, AllowDeletingRows, AllowSorting,
        #         AllowFiltering, AllowUsingPivotTables)
        ws.Protect(ii.SENHA, f['objetos'], f['conteudo'], f['cenarios'], True,
                   f['AllowFormattingCells'], f['AllowFormattingColumns'], f['AllowFormattingRows'],
                   f['AllowInsertingColumns'], f['AllowInsertingRows'], f['AllowInsertingHyperlinks'],
                   f['AllowDeletingColumns'], f['AllowDeletingRows'], f['AllowSorting'],
                   f['AllowFiltering'], f['AllowUsingPivotTables'])
    depois = foto_protecao(wb)
    dif = [n for n in foto if foto[n] != depois.get(n)]
    if dif:
        raise SystemExit(f'protecao nao voltou ao estado original em: {dif}')


def main(produto, caminho, salvar=True):
    caminho = os.path.abspath(caminho)
    ex = xlh.Excel()
    log(f'EXCEL_PID {ex.pid}')
    try:
        wb = ex.abrir(caminho)
        if wb.ReadOnly:          # outro Excel segura o arquivo: o Save 'funcionaria' sem gravar nada
            raise SystemExit(f'arquivo aberto SOMENTE LEITURA (outra instancia o segura): {caminho}')
        estrutura = bool(wb.ProtectStructure)
        foto = foto_protecao(wb)
        log(f'   protecao lida de {len(foto)} abas; estrutura protegida: {estrutura}')
        ii.desproteger_tudo(wb)

        log('1. VBA (todo o codigo atual das fontes: codigo_atual.py)')
        import codigo_atual
        codigo_atual.instalar(ex, wb, produto, log)

        log('2. aba Usuarios: travas, area de cadastro, coluna do hash')
        ws = wb.Worksheets('Usuarios')
        for a in TRAVAR:
            ws.Range(a).Locked = True
        for n in CADASTRO:
            wb.Names(n).RefersToRange.Locked = False
        ws.Columns(COL_HASH).Hidden = True

        log('3. sessao e formulario de cadastro limpos')
        for n in ('currentUser', 'currentPapel', 'cadMsg') + tuple(CADASTRO):
            wb.Names(n).RefersToRange.Value = ''

        # conferencia lida de volta, antes de reproteger
        falhas = [a for a in TRAVAR if ws.Range(a).Locked is not True]
        falhas += [n for n in CADASTRO if wb.Names(n).RefersToRange.Locked is not False]
        if not ws.Columns(COL_HASH).Hidden:
            falhas.append('coluna do hash visivel')
        if any(wb.Names(n).RefersToRange.Value not in (None, '') for n in ('currentUser', 'currentPapel', 'cadSenha')):
            falhas.append('sessao/senha nao limpas')
        if falhas:
            raise SystemExit(f'conferencia da aba Usuarios falhou: {falhas}')
        log('   N1:N2 e A4:E53 travadas; cadastro destravado; hash oculto; sessao vazia')

        log('4. protecao devolvida aba a aba')
        devolver_protecao(wb, foto)
        if estrutura:
            wb.Protect(ii.SENHA, True, False)
        if bool(wb.ProtectStructure) != estrutura:
            raise SystemExit('protecao da estrutura nao voltou ao estado original')
        log(f'   {sum(1 for f in foto.values() if f["conteudo"])} abas reprotegidas com as mesmas permissoes')
        if salvar:
            wb.Save()
            log('salvo')
    finally:
        ex.fechar()


if __name__ == '__main__':
    a = sys.argv[1:]
    main(a[0], a[1], salvar='--sem-salvar' not in a)
