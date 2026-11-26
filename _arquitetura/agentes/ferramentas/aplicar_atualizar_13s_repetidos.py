# -*- coding: utf-8 -*-
"""ACAO PONTUAL (26/11/2026, pedido do usuario; NAO e regra nova do sistema): ATUALIZAR na producao e inativar os
resultados que violaram 1-3s E foram repetidos, pelo mesmo caminho seguro do aplicar_atualizar/aplicar_inativacao.

Uso: python aplicar_atualizar_13s_repetidos.py <Bioquimica|Hematologia> <pasta_saida_fora_do_repo> [--simular] [--desde=AAAA-MM-DD]
(escopo escolhido pelo usuario em 26/11/2026: --desde=2026-11-01, so novembro/2026)

Criterio (decisao do usuario em 26/11/2026):
  - resultado ATIVO e participante; alvo = media/DP do lote (mLotes.MediaDoLote/DPDoLote, o mesmo do LJ);
  - viola 1-3s: |resultado - media| > 3 DP;
  - REPETIDO: existe outro resultado ATIVO do mesmo analito, nivel, lote e equipamento, no MESMO dia e com hora
    POSTERIOR, dentro de +-3 DP. So entao o que ficou fora e inativado (X vermelho), com motivo que cita o repetido.
  - violacao cuja repeticao tambem ficou fora, ou sem repeticao: NAO e inativada (fica listada para o RT).
Passos: copia a producao, ATUALIZAR (le o DB_SEAC), detecta, mIntegracao.RegistrarInativacao para cada um,
AtualizarSoEntradas, confere STATUS=INATIVADO, salva, blinda, confere que a producao nao mudou, backup e troca (sha).
A lista com IDs (dados do laboratorio) vai so para a pasta de saida, nunca para o repositorio.
"""
import collections
import datetime
import hashlib
import json
import os
import re
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

TETO_INATIVACOES = 100      # acima disso algo esta errado (alvo ausente/errado): nada e trocado


def sha(p):
    return hashlib.sha256(open(p, 'rb').read()).hexdigest()


def num(v):
    try:
        return float(v)
    except Exception:
        return None


def nivel(v):
    m = re.search(r'\d', str(v or ''))
    return int(m.group(0)) if m else 0


def detectar(q, rows, desde=None):
    alvo = {}

    def par(lote, analito, nv):
        k = (lote, analito, nv)
        if k not in alvo:
            m = num(q.run('mLotes.MediaDoLote', lote, analito, nv))
            s = num(q.run('mLotes.DPDoLote', lote, analito, nv))
            alvo[k] = (m, s) if (m is not None and s and s > 0) else None
        return alvo[k]

    grupos = collections.defaultdict(list)
    for r in rows:
        if str(r.get('STATUS_ANALITICO') or '').upper() != 'ATIVO':
            continue
        if str(r.get('PARTICIPA_ESTATISTICA') or '').upper() != 'SIM':
            continue
        res, dh, d = num(r.get('RESULTADO')), r.get('DATA_HORA'), r.get('DATA')
        lote, an, nv = str(r.get('LOTE') or '').strip(), str(r.get('ANALITO') or '').strip(), nivel(r.get('NIVEL'))
        if res is None or not lote or not an or not nv or not isinstance(dh, datetime.datetime):
            continue
        p = par(lote, an, nv)
        if not p:
            continue
        z = (res - p[0]) / p[1]
        dia = d.date() if isinstance(d, datetime.datetime) else dh.date()
        grupos[(an, nv, lote, str(r.get('EQUIPAMENTO') or '').strip(), dia)].append(
            {'id': str(r.get('ID_REGISTRO')), 'dh': dh, 'res': res, 'z': z, 'an': an, 'nv': nv, 'lote': lote})

    inativar, manter = [], []
    for k, lst in grupos.items():
        lst.sort(key=lambda x: x['dh'])
        for i, x in enumerate(lst):
            if abs(x['z']) <= 3 or (desde and x['dh'].date() < desde):
                continue
            rep = [y for y in lst[i + 1:] if y['dh'] > x['dh']]
            ok = [y for y in rep if abs(y['z']) <= 3]
            if ok:
                inativar.append(dict(x, repetido=ok[0]['id'], z_rep=ok[0]['z']))
            else:
                manter.append(dict(x, motivo='repeticao tambem fora de 3DP' if rep else 'sem repeticao no dia'))
    sem_alvo = sorted({f'{a}|N{n}|{l}' for (l, a, n), v in alvo.items() if v is None})
    return inativar, manter, sem_alvo


def main(produto, saida, simular, desde=None):
    os.makedirs(saida, exist_ok=True)
    prod = os.path.join(REPO, f'QC_{produto}.xlsm')
    pasta = os.path.join(os.environ['TEMP'], 'qcw', 'atualizar13s')
    os.makedirs(pasta, exist_ok=True)
    copia = os.path.join(pasta, f'QC_{produto}.xlsm')
    shutil.copy2(prod, copia)
    h0 = sha(prod)
    print('producao', h0[:16], flush=True)
    rel = {'produto': produto, 'quando': datetime.datetime.now().isoformat(timespec='minutes'), 'sha_antes': h0}

    q = QA(produto, copia)
    try:
        resumo, seg = q.atualizar(teto=1800)
        print(f'ATUALIZAR: {resumo} ({seg:.0f} s)', flush=True)
        rel['resumo_atualizar'], rel['segundos_atualizar'] = resumo, round(seg)
        rows = q.ler('tblCQ_Final')
        inativar, manter, sem_alvo = detectar(q, rows, desde)
        rel['desde'] = str(desde)
        print(f'1-3s repetidos a inativar: {len(inativar)}; 1-3s mantidos: {len(manter)}; '
              f'combinacoes sem alvo: {len(sem_alvo)}', flush=True)
        for x in ([] if simular else inativar):
            print(f"  INATIVAR {x['an']} N{x['nv']} lote {x['lote']} {x['dh']:%d/%m/%Y %H:%M} z={x['z']:+.2f} "
                  f"-> repetido z={x['z_rep']:+.2f}", flush=True)
        for x in ([] if simular else manter):
            print(f"  MANTIDO  {x['an']} N{x['nv']} lote {x['lote']} {x['dh']:%d/%m/%Y %H:%M} z={x['z']:+.2f} "
                  f"({x['motivo']})", flush=True)
        if len(inativar) > TETO_INATIVACOES and not simular:
            raise SystemExit(f'ERRO: {len(inativar)} inativacoes passam do teto {TETO_INATIVACOES}; nada trocado')
        feitos = []
        if not simular:
            hoje = datetime.date.today().strftime('%d/%m/%Y')
            for x in inativar:
                mot = (f"Violação 1-3s (z={x['z']:+.2f}); resultado repetido (ID {x['repetido']}, "
                       f"z={x['z_rep']:+.2f}) dentro dos limites. Inativação pontual autorizada pelo usuário em {hoje}.")
                r = str(q.run('mIntegracao.RegistrarInativacao', x['id'], x['an'], mot))
                if not r.startswith('OK'):
                    raise SystemExit(f"ERRO ao registrar {x['id']}: {r}; nada trocado")
                feitos.append(x['id'])
            if feitos:
                r = q.run('mIntegracao.AtualizarSoEntradas', teto=1800)
                print('AtualizarSoEntradas:', r, flush=True)
                st = {str(r.get('ID_REGISTRO')): str(r.get('STATUS_ANALITICO') or '') for r in q.ler('tblCQ_Final')}
                ruins = [i for i in feitos if st.get(i) != 'INATIVADO']
                if ruins:
                    raise SystemExit(f'ERRO: {len(ruins)} ID(s) nao ficaram INATIVADO; nada trocado')
                print(f'conferido: {len(feitos)} INATIVADO na tblCQ_Final', flush=True)
            q.ex.xl.EnableEvents = False
            q.wb.Save()
        cont = collections.Counter(str(r.get('STATUS_ANALITICO') or '') for r in q.ler('tblCQ_Final'))
        rel['status_depois'] = dict(cont)
    finally:
        q.ex.fechar()

    conv = lambda L: [{k: (v.isoformat() if isinstance(v, datetime.datetime) else v) for k, v in x.items()} for x in L]
    rel.update({'inativados': conv(inativar) if not simular else [], 'candidatos': conv(inativar),
                'mantidos': conv(manter), 'sem_alvo': sem_alvo, 'simular': simular})
    arq = os.path.join(saida, f'atualizar13s_{produto}.json')
    if simular:
        json.dump(rel, open(arq, 'w', encoding='utf-8'), ensure_ascii=False, indent=1, default=str)
        print('SIMULACAO: nada trocado ->', arq, flush=True)
        return
    blindar_entrega.blindar(copia)
    if not blindar_entrega.verificar(copia):
        raise SystemExit('ERRO: a copia nao ficou blindada; nada foi trocado')
    if sha(prod) != h0:
        raise SystemExit('ERRO: a producao mudou durante a aplicacao; nada foi trocado')
    stamp = datetime.datetime.now().strftime('%Y-%m-%d_%H%M')
    bk = os.path.join(REPO, f'_backup_pre_atualizar13s_{stamp}')
    os.makedirs(bk)
    shutil.copy2(prod, os.path.join(bk, os.path.basename(prod)))
    shutil.copy2(copia, prod)
    h1 = sha(prod)
    if h1 != sha(copia):
        raise SystemExit('ERRO: sha da producao difere da copia aplicada')
    rel['sha_depois'], rel['backup'] = h1, os.path.basename(bk)
    json.dump(rel, open(arq, 'w', encoding='utf-8'), ensure_ascii=False, indent=1, default=str)
    print(f'APLICADO {produto}: sha {h0[:16]} -> {h1[:16]} (backup {bk}); {len(feitos)} inativados', flush=True)
    print('tamanho', os.path.getsize(prod), 'sha', h1, flush=True)


if __name__ == '__main__':
    a = [x for x in sys.argv[1:] if not x.startswith('--')]
    if len(a) != 2:
        print(__doc__)
        sys.exit(2)
    ds = [x.split('=', 1)[1] for x in sys.argv if x.startswith('--desde=')]
    main(a[0], a[1], '--simular' in sys.argv, datetime.date.fromisoformat(ds[0]) if ds else None)
