# -*- coding: utf-8 -*-
"""Aplica na PRODUCAO uma inativacao pedida pelo usuario, pelo mesmo caminho seguro da entrega.

Uso: python aplicar_inativacao.py <Bioquimica|Hematologia> <ID> <ANALITO> <MOTIVO>

1. copia a producao (raiz do repositorio) para %TEMP%\\qcw\\inativar\\ e guarda o sha256;
2. na copia: login tecnico (qa_final.QA), mIntegracao.RegistrarInativacao (aba Inativar, carimbo e trilha),
   mIntegracao.AtualizarSoEntradas (sem ler o DB_SEAC), confere o STATUS do ID na tblCQ_Final;
3. salva, blinda (blindar_entrega.py) e verifica a blindagem;
4. confere que a producao nao mudou desde a copia, faz backup e troca, conferindo o sha256 depois.
Encerra so o Excel que abriu. Nunca grava direto na producao.
"""
import datetime
import hashlib
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


def main(produto, id_, analito, motivo):
    prod = os.path.join(REPO, f'QC_{produto}.xlsm')
    pasta = os.path.join(os.environ['TEMP'], 'qcw', 'inativar')
    os.makedirs(pasta, exist_ok=True)
    copia = os.path.join(pasta, f'QC_{produto}.xlsm')
    shutil.copy2(prod, copia)
    h0 = sha(prod)
    print('producao', h0[:16], flush=True)

    q = QA(produto, copia)
    try:
        xl = q.ex
        nome = q.wb.Name
        r = xl.run(f"'{nome}'!mIntegracao.RegistrarInativacao", id_, analito, motivo)
        print('RegistrarInativacao:', r, flush=True)
        if not str(r).startswith('OK'):
            raise SystemExit(f'ERRO ao registrar: {r}')
        r = xl.run(f"'{nome}'!mIntegracao.AtualizarSoEntradas", teto=900)
        print('AtualizarSoEntradas:', r, flush=True)
        st = None
        for ws in q.wb.Worksheets:
            try:
                t = ws.ListObjects('tblCQ_Final')
            except Exception:
                continue
            ids = [c[0] for c in t.ListColumns('ID_REGISTRO').DataBodyRange.Value]
            sts = [c[0] for c in t.ListColumns('STATUS_ANALITICO').DataBodyRange.Value]
            part = [c[0] for c in t.ListColumns('PARTICIPA_ESTATISTICA').DataBodyRange.Value]
            for i, v in enumerate(ids):
                if str(v) == id_:
                    st = (sts[i], part[i])
            break
        print('tblCQ_Final', id_, '->', st, flush=True)
        if not st or st[0] != 'INATIVADO':
            raise SystemExit(f'ERRO: {id_} nao ficou INATIVADO na tblCQ_Final: {st}')
        xl.xl.EnableEvents = False
        q.wb.Save()
    finally:
        q.ex.fechar()

    blindar_entrega.blindar(copia)
    if not blindar_entrega.verificar(copia):
        raise SystemExit('ERRO: a copia nao ficou blindada; nada foi trocado')
    if sha(prod) != h0:
        raise SystemExit('ERRO: a producao mudou durante a aplicacao; nada foi trocado')
    stamp = datetime.datetime.now().strftime('%Y-%m-%d_%H%M')
    bk = os.path.join(REPO, f'_backup_pre_inativacao_{stamp}')
    os.makedirs(bk)
    shutil.copy2(prod, os.path.join(bk, os.path.basename(prod)))
    shutil.copy2(copia, prod)
    h1, hc = sha(prod), sha(copia)
    if h1 != hc:
        raise SystemExit('ERRO: sha da producao difere da copia aplicada')
    print(f'APLICADO {produto} {id_}: sha {h0[:16]} -> {h1[:16]} (backup {bk})', flush=True)


if __name__ == '__main__':
    if len(sys.argv) != 5:
        print(__doc__)
        sys.exit(2)
    main(*sys.argv[1:])
