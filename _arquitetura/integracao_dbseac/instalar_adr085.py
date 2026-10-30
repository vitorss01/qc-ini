# -*- coding: utf-8 -*-
"""instalar_adr085.py -- ADR-085 (29/10/2026): o Painel como tela unica de trabalho.

Uso:  python instalar_adr085.py <Bioquimica|Hematologia> <arquivo.xlsm> [--sem-salvar]

Idempotente (a 2a execucao registra 0 mudancas). Passos:
  1. janela frmInativar com as acoes explicitas (nao conforme / inativar / comentario; no X: reativar / tirar o X /
     comentario) -- instalar_frm_inativar.criar_form;
  2. todo o VBA atual das fontes (codigo_atual.py: clsCht que cancela o "Formatar" do Excel em todo duplo clique,
     mUI.CliqueNoGrafico/DicaCompleta, mIntegracao.GravarComentarioTecnico/AplicarComentarioNaBase/TrocarPlotagemPorId,
     mEtlCQ.ComentarioConsolidado, mEstatistica.RegrasDoPonto, mCalibracao.AssinarRegistro, Registros.cls) -- EXIGE
     compilacao;
  3. Painel: sai o "Historico" da coluna U (U6 + uma linha por nivel). ANTES confere que nenhuma formula da pasta
     (outra aba, nome definido ou o proprio Painel) le essas celulas -- se ler, PARA sem salvar. A coluna 21 do
     engPainel (o texto do historico) continua calculada pelo motor: nada que a auditoria usa e apagado;
  4. Registros: I3 "Assinatura (duplo clique)" -- a assinatura pelo login (mCalibracao.AssinarRegistro);
  5. devolve a protecao de cada aba e a da estrutura como estavam. Qualquer falha para ANTES de salvar.
"""
import os
import re
import sys

AQUI = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, AQUI)
sys.path.insert(0, os.path.join(AQUI, '..', 'etapa1_multilote'))
import xlh  # noqa: E402
import instalar_integracao as ii  # noqa: E402
import codigo_atual  # noqa: E402
import instalar_frm_inativar  # noqa: E402
from instalar_seguranca_usuarios import foto_protecao, devolver_protecao  # noqa: E402

MUDANCAS = []
CAB_HIST = 'Histórico'
CAB_ASSIN = 'Assinatura (duplo clique)'


def log(m):
    print(m, flush=True)


def _formulas(ws):
    """(endereco, formula) de cada celula com formula na area usada."""
    ur = ws.UsedRange
    f = ur.Formula
    if not isinstance(f, tuple):
        f = ((f,),)
    r0, c0 = ur.Row, ur.Column
    for i, linha in enumerate(f):
        for j, v in enumerate(linha):
            if isinstance(v, str) and v.startswith('='):
                yield (r0 + i, c0 + j), v


def dependencias_coluna_u(wb, nlv):
    """Formulas que leem Painel!U6:U<6+nlv> -- fora as proprias celulas do historico."""
    ult = 6 + nlv
    alvo = {(r, 21) for r in range(6, ult + 1)}
    cel = r'\$?U\$?([6-9])(?![0-9])'
    rx_outra = re.compile(r"(?:'?Painel'?)!" + cel, re.I)
    rx_mesma = re.compile(r"(?<![A-Za-z0-9_$!.])" + cel)
    achados = []
    for ws in wb.Worksheets:
        for (r, c), f in _formulas(ws):
            if ws.Name == 'Painel' and (r, c) in alvo:
                continue
            refs = [int(m.group(1)) for m in rx_outra.finditer(f)]
            if ws.Name == 'Painel':
                refs += [int(m.group(1)) for m in rx_mesma.finditer(f)]
            if any(6 <= x <= ult for x in refs):
                achados.append(f'{ws.Name}!R{r}C{c}: {f[:120]}')
    for n in wb.Names:
        refs = [int(m.group(1)) for m in rx_outra.finditer(str(n.RefersTo))]
        if any(6 <= x <= ult for x in refs):
            achados.append(f'nome {n.Name}: {n.RefersTo}')
    return achados


def painel_sem_historico(wb, nlv):
    p = wb.Worksheets('Painel')
    ult = 6 + nlv
    faixa = p.Range(f'U6:U{ult}')
    vals = [str(r[0] or '') for r in faixa.Formula]
    if not any(vals):
        return                                           # ja removido
    if str(p.Range('U6').Value or '').strip() != CAB_HIST:
        raise SystemExit(f'Painel!U6 inesperado ({p.Range("U6").Value!r}): layout diferente -- nada foi feito')
    for k, f in enumerate(vals[1:], start=1):
        if f and 'engPainel' not in f:
            raise SystemExit(f'Painel!U{6 + k} inesperado ({f[:80]!r}): layout diferente -- nada foi feito')
    deps = dependencias_coluna_u(wb, nlv)
    if deps:
        raise SystemExit('Painel!U (Historico) e lida por outras formulas -- nada foi feito: ' + ' | '.join(deps[:5]))
    faixa.ClearContents()
    # a celula fica com o fundo e as bordas da vizinha vazia de baixo (sem cabecalho orfao no meio do bloco)
    p.Range(f'U{ult + 1}').Copy()
    faixa.PasteSpecial(-4122)                            # xlPasteFormats
    wb.Application.CutCopyMode = False
    MUDANCAS.append(f'Painel U6:U{ult} (Historico) removido')


def registros(wb):
    ws = wb.Worksheets('Registros')
    if 'tblRegistros' in [lo.Name for lo in ws.ListObjects]:
        return                                       # ADR-087: a aba virou tabela -- o adr087 e o dono do layout
    v = str(ws.Range('I3').Value or '').strip()
    if v == CAB_ASSIN:
        return
    if v != 'Assinatura':
        raise SystemExit(f'Registros!I3 inesperado ({v!r}): layout diferente -- nada foi feito')
    ws.Range('I3').Value = CAB_ASSIN
    MUDANCAS.append('Registros I3')


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
        nlv = 2 if produto.startswith('Bio') else 3
        log('1. janela frmInativar: acoes explicitas (nao conforme / inativar / comentario)')
        n = instalar_frm_inativar.criar_form(wb, log)
        if n:
            MUDANCAS.append('frmInativar')
        log('2. VBA (todo o codigo atual das fontes: codigo_atual.py)')
        codigo_atual.instalar(ex, wb, produto, log)
        painel_sem_historico(wb, nlv)
        log(f'3. Painel: coluna U (Historico) fora da tela; dependencias conferidas')
        registros(wb)
        log('4. Registros: I3 "Assinatura (duplo clique)"')
        devolver_protecao(wb, foto)
        if estrutura:
            wb.Protect(ii.SENHA, True, False)
        log(f'MUDANCAS {len(MUDANCAS)}' + (': ' + '; '.join(MUDANCAS) if MUDANCAS else ''))
        if salvar:
            wb.Save()
            log('salvo')
    finally:
        ex.fechar()


if __name__ == '__main__':
    a = sys.argv[1:]
    if len(a) < 2:
        raise SystemExit(__doc__)
    main(a[0], a[1], salvar='--sem-salvar' not in a)
