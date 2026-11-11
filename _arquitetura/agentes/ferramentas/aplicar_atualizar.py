# -*- coding: utf-8 -*-
"""Roda o ATUALIZAR DADOS na PRODUCAO pelo mesmo caminho seguro da entrega e mede o efeito.

Uso: python aplicar_atualizar.py <Bioquimica|Hematologia> <saida.json>

1. copia a producao (raiz do repositorio) para %TEMP%\\qcw\\atualizar\\ e guarda o sha256;
2. na copia: login tecnico (qa_final.QA), le a tblCQ_Final (antes), mIntegracao.AtualizarDadosAutomatico
   (le o DB_SEAC da origem configurada no arquivo), le a tblCQ_Final (depois);
3. grava no JSON: contagem por STATUS_ANALITICO, participantes por lote/analito e as diferencas;
4. salva, blinda (blindar_entrega.py) e verifica a blindagem;
5. confere que a producao nao mudou desde a copia, faz backup e troca, conferindo o sha256 depois.
Encerra so o Excel que abriu. Nunca grava direto na producao.
"""
import collections
import datetime
import hashlib
import json
import os
import shutil
import sys

AQUI = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.abspath(os.path.join(AQUI, '..', '..', '..'))
INT = os.path.join(REPO, '_arquitetura', 'integracao_dbseac')
sys.path.insert(0, INT)
sys.path.insert(0, os.path.join(INT, 'testes'))
sys.path.insert(0, os.path.join(REPO, '_arquitetura', 'etapa1_multilote'))
import blindar_entrega  # noqa: E402
from qa_final import QA  # noqa: E402


def sha(p):
    return hashlib.sha256(open(p, 'rb').read()).hexdigest()


def foto(q):
    rows = q.ler('tblCQ_Final')
    st = collections.Counter(str(r.get('STATUS_ANALITICO') or '') for r in rows)
    part = collections.Counter()
    for r in rows:
        if str(r.get('PARTICIPA_ESTATISTICA') or '').upper() == 'SIM':
            part[f"{str(r.get('LOTE') or '').strip()}|{str(r.get('ANALITO') or '').strip()}"] += 1
    return {'linhas': len(rows), 'status': dict(st), 'participantes': dict(part)}


def main(produto, saida):
    prod = os.path.join(REPO, f'QC_{produto}.xlsm')
    pasta = os.path.join(os.environ['TEMP'], 'qcw', 'atualizar')
    os.makedirs(pasta, exist_ok=True)
    copia = os.path.join(pasta, f'QC_{produto}.xlsm')
    shutil.copy2(prod, copia)
    h0 = sha(prod)
    print('producao', h0[:16], flush=True)

    q = QA(produto, copia)
    try:
        antes = foto(q)
        resumo, seg = q.atualizar(teto=1800)
        print(f'ATUALIZAR: {resumo} ({seg:.0f} s)', flush=True)
        depois = foto(q)
        q.ex.xl.EnableEvents = False
        q.wb.Save()
    finally:
        q.ex.fechar()

    chaves = set(antes['participantes']) | set(depois['participantes'])
    mudou = {k: [antes['participantes'].get(k, 0), depois['participantes'].get(k, 0)] for k in sorted(chaves)
             if antes['participantes'].get(k, 0) != depois['participantes'].get(k, 0)}
    rel = {'produto': produto, 'quando': datetime.datetime.now().isoformat(timespec='minutes'),
           'sha_antes': h0, 'resumo_atualizar': resumo, 'segundos': round(seg),
           'linhas': [antes['linhas'], depois['linhas']],
           'status_antes': antes['status'], 'status_depois': depois['status'],
           'participantes_que_mudaram': mudou}

    blindar_entrega.blindar(copia)
    if not blindar_entrega.verificar(copia):
        raise SystemExit('ERRO: a copia nao ficou blindada; nada foi trocado')
    if sha(prod) != h0:
        raise SystemExit('ERRO: a producao mudou durante a aplicacao; nada foi trocado')
    stamp = datetime.datetime.now().strftime('%Y-%m-%d_%H%M')
    bk = os.path.join(REPO, f'_backup_pre_atualizar_{stamp}')
    os.makedirs(bk)
    shutil.copy2(prod, os.path.join(bk, os.path.basename(prod)))
    shutil.copy2(copia, prod)
    h1 = sha(prod)
    if h1 != sha(copia):
        raise SystemExit('ERRO: sha da producao difere da copia aplicada')
    rel['sha_depois'], rel['backup'] = h1, os.path.basename(bk)
    with open(saida, 'w', encoding='utf-8') as f:
        json.dump(rel, f, ensure_ascii=False, indent=1)
    print(f'APLICADO {produto}: sha {h0[:16]} -> {h1[:16]} (backup {bk})', flush=True)
    print('status antes ', antes['status'], flush=True)
    print('status depois', depois['status'], flush=True)
    print('participantes que mudaram:', len(mudou), flush=True)


if __name__ == '__main__':
    if len(sys.argv) != 3:
        print(__doc__)
        sys.exit(2)
    main(*sys.argv[1:])
