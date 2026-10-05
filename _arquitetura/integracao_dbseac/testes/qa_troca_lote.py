# -*- coding: utf-8 -*-
"""QA de TROCA DE LOTE (QUALITY_GATE 4.3) -- roda numa COPIA ja instalada.

Uso: python qa_troca_lote.py <Bioquimica|Hematologia> <copia.xlsm> [saida.json]

Pelo caminho real: escreve o lote em Painel!H3 com eventos ligados (Worksheet_Change ->
mLotes.TrocarLoteAnalise), como o usuario faz na lista. Pergunta do achado ALTO #20 da
FASE3A: o lote novo herda em silencio a media/DP do anterior? E a ida e volta preserva os
parametros do lote original? Trocar de lote nao e alterar parametro: nada de
PARAMETRO_LOTE_ALTERADO na trilha. Nada e salvo.

Sessao: o arquivo entregue e salvo blindado; aqui se faz o que o login faz (ReprotectAll,
protecao com UserInterfaceOnly nesta sessao).
"""
import json
import os
import sys

AQUI = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, AQUI)
from qa_final import QA, reg, RES  # noqa: E402

LOTE_TESTE = 'QA9999'
SEM = 'SEM MÉDIA/DP'


def executar(produto, caminho, saida):
    q = QA(produto, caminho)
    try:
        q.run('mSeguranca.ReprotectAll')
        p = q.ws('Painel')
        lote1 = str(p.Range('H3').Value)
        analito = p.Range('C3').Value
        cad = q.wb.Names('regLoteCol').RefersToRange
        lotes = [str(c.Value) for c in cad.Cells if c.Value not in (None, '')]
        sem_par = [L for L in lotes if L != lote1 and all(
            q.run('mLotes.MediaDoLote', L, analito, n) == '' for n in range(1, q.nlv + 1))]
        if sem_par:
            lote2 = sem_par[0]
        else:
            # so um lote cadastrado: registra um lote de teste SEM parametros (so nesta copia)
            ws = cad.Worksheet
            ws.Unprotect('qcini2025')
            livre = next(c for c in cad.Cells if c.Value in (None, ''))
            livre.Value = LOTE_TESTE
            q.run('mSeguranca.ReprotectAll')
            lote2 = LOTE_TESTE
        par1 = [(q.run('mLotes.MediaDoLote', lote1, analito, n), q.run('mLotes.DPDoLote', lote1, analito, n))
                for n in range(1, q.nlv + 1)]
        st = lambda: [p.Range(f'J{6 + n}').Text for n in range(1, q.nlv + 1)]
        media_painel = lambda: [p.Range(f'C{6 + n}').Text for n in range(1, q.nlv + 1)]
        st1 = st()
        marca = q.ws('Audit_Log').Cells(q.ws('Audit_Log').Rows.Count, 1).End(-4162).Row

        p.Range('H3').Value = lote2                    # o usuario escolhe o outro lote na lista
        st2 = st()
        guarda2 = [q.run('mLotes.MediaDoLote', lote2, analito, n) for n in range(1, q.nlv + 1)]
        reg(f'L01 Troca para lote sem média/DP ({lote2}): o Painel diz "{SEM}" em todo nível, nada herdado de {lote1}',
            str(p.Range('H3').Value) == lote2 and all(s == SEM for s in st2) and all(g == '' for g in guarda2),
            {'analito': analito, 'lote_antes': lote1, 'status_antes': st1, 'lote_depois': lote2, 'status_depois': st2})

        p.Range('H3').Value = lote1                    # e volta
        st3 = st()
        par3 = [(q.run('mLotes.MediaDoLote', lote1, analito, n), q.run('mLotes.DPDoLote', lote1, analito, n))
                for n in range(1, q.nlv + 1)]
        reg(f'L02 Volta para {lote1}: parâmetros intactos (ida e volta não grava a tela vazia por cima) e Painel avaliando de novo',
            par3 == par1 and st3 == st1 and any(s != SEM for s in st3),
            {'parametros_antes': par1, 'parametros_depois': par3, 'status': st3})

        al = q.ws('Audit_Log')
        ult = al.Cells(al.Rows.Count, 1).End(-4162).Row
        acoes = [al.Cells(r, 7).Value for r in range(marca + 1, ult + 1)]
        reg('L03 Trocar de lote não é alterar parâmetro: nenhum PARAMETRO_LOTE_ALTERADO na trilha',
            'PARAMETRO_LOTE_ALTERADO' not in acoes, {'eventos_na_troca': acoes})
    finally:
        q.fechar(salvar=False)
    if saida:
        with open(saida, 'w', encoding='utf-8') as f:
            json.dump(RES, f, ensure_ascii=False, indent=1, default=str)
    falhas = [r for r in RES if r['resultado'] == 'FAIL']
    print(f'\n=== {produto} troca de lote: {len(RES) - len(falhas)} PASS / {len(falhas)} FAIL ===', flush=True)
    return falhas


if __name__ == '__main__':
    falhas = executar(sys.argv[1], os.path.abspath(sys.argv[2]), sys.argv[3] if len(sys.argv) > 3 else None)
    sys.exit(1 if falhas else 0)
