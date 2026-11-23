# -*- coding: utf-8 -*-
"""instalar_adr099.py -- ADR-099 (22/11/2026): coluna MATERIAL no cadastro de lotes (Configuração!F25:F125).

Uso:  python instalar_adr099.py <Bioquimica|Hematologia> <arquivo.xlsm> [--sem-salvar]

A entrada amigavel da regra MATRIZ da tblLotesMaterial (ADR-091), que continua a UNICA fonte dos dois motores (VBA e
PQ ficam iguais por construcao). Idempotente (a 2a execucao registra 0 mudancas). Passos:
  1. todo o VBA atual das fontes (codigo_atual.py): mMaterialLote (novo), Configuracao.cls (F26:F125 -> MaterialEditado),
     mIntegracao (espelho no ATUALIZAR), mLotes (espelho ao abrir), mApp (lote novo); EXIGE compilacao;
  2. Configuração: F25 "Material" (estilo do D25), F26:F125 texto, destravadas, com a lista das matrizes do de/para
     (Bio: SORO;URINA; Hema: SANGUE TOTAL). Aborta se F25:F125 tiver conteudo que nao seja desta coluna;
  3. carga inicial = SO o espelho das regras MATRIZ existentes (nenhuma regra nova: nenhum resultado muda);
  4. fumaca: LotesMaterialResumo igual ao de antes; F = regra de 100 % dos lotes do cadastro; a tblLotesMaterial
     inalterada;
  5. devolve a protecao como estava e salva com os eventos desligados. Qualquer falha para ANTES de salvar.
"""
import os
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
CFG = 'Configuração'
R0, R1, C_LOTE, C_MAT = 26, 125, 3, 6
TITULO = 'Material'
MSG = ('O material vale para TODOS os resultados do lote (interfaceamento e manuais), a partir do próximo ATUALIZAR. '
       'Vazio = padrão do equipamento.')


def log(m):
    print(m, flush=True)


def tabela_lm(wb):
    lo = pqlib.tabela(wb, 'tblLotesMaterial')
    if lo is None:
        raise SystemExit('tblLotesMaterial nao existe (instalar_adr091 antes) -- nada foi salvo')
    return [tuple(str(c or '').strip() for c in l) for l in (lo.DataBodyRange.Value if lo.ListRows.Count else ())]


def coluna(wb, ex):
    ws = wb.Worksheets(CFG)
    f25 = str(ws.Range('F25').Value or '')
    if f25 != TITULO:
        alheio = [ws.Cells(r, C_MAT).Address for r in range(25, R1 + 1) if ws.Cells(r, C_MAT).Formula not in (None, '')]
        if alheio:
            raise SystemExit(f'{CFG}!F25:F125 ocupada por outro conteudo ({alheio[:5]}) -- nada foi salvo')
        ws.Range('F25').Value = TITULO
        MUDANCAS.append('Configuração F25 "Material"')
    h = ws.Range('F25')
    h.Font.Bold = ws.Range('D25').Font.Bold
    h.Interior.Color = ws.Range('D25').Interior.Color
    h.Font.Color = ws.Range('D25').Font.Color
    if ws.Columns(C_MAT).ColumnWidth < 14:
        ws.Columns(C_MAT).ColumnWidth = 16
    corpo = ws.Range(ws.Cells(R0, C_MAT), ws.Cells(R1, C_MAT))
    if str(corpo.Cells(1, 1).NumberFormat) != '@':
        corpo.NumberFormat = '@'
    corpo.Locked = False
    lista = str(ex.run(f"'{wb.Name}'!mMaterialLote.MatrizesPermitidas", teto=60) or '')
    if not lista:
        raise SystemExit('MatrizesPermitidas vazia (de/para sem MATRIZ e sem MATRIZ_PADRAO) -- nada foi salvo')
    f1 = cd.sep_lista(ex.xl).join(lista.split(';'))
    atual = None
    try:
        atual = str(corpo.Cells(1, 1).Validation.Formula1)
    except Exception:                                   # noqa: BLE001 -- sem validacao
        pass
    if atual != f1:
        corpo.Validation.Delete()
        corpo.Validation.Add(3, 1, 1, f1)                # xlValidateList, xlValidAlertStop, xlBetween
        corpo.Validation.IgnoreBlank = True
        corpo.Validation.InCellDropdown = True
        corpo.Validation.InputTitle = 'Material do lote'
        corpo.Validation.InputMessage = MSG
        corpo.Validation.ErrorTitle = 'Material não reconhecido'
        corpo.Validation.ErrorMessage = 'Escolha da lista: ' + ', '.join(lista.split(';')) + ' (ou deixe vazio).'
        MUDANCAS.append(f'Configuração F26:F125 lista ({lista})')
    return lista


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
        tab0 = tabela_lm(wb)
        log('1. VBA (todo o codigo atual das fontes: codigo_atual.py)')
        codigo_atual.instalar(ex, wb, produto, log)
        nome = wb.Name
        res0 = str(ex.run(f"'{nome}'!mEtlCQ.LotesMaterialResumo", teto=120))
        log('2. Configuração: coluna Material')
        lista = coluna(wb, ex)
        log('3. carga inicial (so o espelho das regras MATRIZ)')
        r = str(ex.run(f"'{nome}'!mMaterialLote.CargaInicial", teto=300))
        if not r.startswith('OK|'):
            raise SystemExit(f'CargaInicial: {r[:300]} -- nada foi salvo')
        _, n, sem = (r.split('|') + ['', ''])[:3]
        if int(n or 0):
            MUDANCAS.append(f'Material espelhado em {n} lote(s)')
        log(f'   espelhados {n}; lotes do cadastro sem regra MATRIZ (vazio = padrao): {sem[:400]}')
        log('4. fumaca')
        res1 = str(ex.run(f"'{nome}'!mEtlCQ.LotesMaterialResumo", teto=120))
        if res1 != res0 or tabela_lm(wb) != tab0:
            raise SystemExit(f'fumaca: a tblLotesMaterial mudou ({res0!r} -> {res1!r}) -- nada foi salvo')
        ws = wb.Worksheets(CFG)
        ruins = []
        for rr in range(R0, R1 + 1):
            lote = str(ws.Cells(rr, C_LOTE).Value or '').strip()
            if isinstance(ws.Cells(rr, C_LOTE).Value, float) and ws.Cells(rr, C_LOTE).Value == int(ws.Cells(rr, C_LOTE).Value):
                lote = str(int(ws.Cells(rr, C_LOTE).Value))
            esp = str(ex.run(f"'{nome}'!mMaterialLote.MatrizRegraDoLote", lote) or '') if lote else ''
            if str(ws.Cells(rr, C_MAT).Value or '').strip().upper() != esp:
                ruins.append((rr, lote, ws.Cells(rr, C_MAT).Value, esp))
        if ruins:
            raise SystemExit(f'fumaca: coluna Material diferente da regra em {ruins[:5]} -- nada foi salvo')
        if ws.Range('F26').Locked is not False or ws.Range('F25').Locked is not True or ws.Range('C26').Locked is not True:
            raise SystemExit('fumaca: travas da Configuração erradas (F26 livre, F25/C26 travadas) -- nada foi salvo')
        log(f'   ok: lista {lista}; LotesMaterialResumo {res1[:120]}')
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
