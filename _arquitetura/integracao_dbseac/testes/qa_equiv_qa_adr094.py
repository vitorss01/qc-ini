# -*- coding: utf-8 -*-
"""QA do ADR-094 nos DOIS motores: a tblQA_Integracao do motor VBA (mEtlQA) = a do Power Query (QA_INTEGRACAO.m),
achado por achado, com o DETALHE (que agora leva a data-hora ao segundo). Roda o estudos/equivalencia_qa.py --fixtures
numa copia da copia instalada (o arquivo de entrada nao e aberto); nada e salvo.

Uso: python qa_equiv_qa_adr094.py <Bioquimica|Hematologia> <copia_instalada.xlsm> <saida.json>

  Q01 conjunto de achados igual nos dois motores (as 8 colunas), mesma ordem das partes, e os casos das fixtures do
      ADR-082/088/091 conferidos (ok) -- a regra de sempre, agora com os textos novos
  Q02 achados com data-hora no DETALHE (E01, A03, E06, A07): existem (0 = FAIL: nao exercitado), mesma contagem nos
      dois motores e todos com segundos no PQ e no VBA
  Y00 seguranca do teste: arquivo de entrada intacto
"""
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile

AQUI = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, AQUI)
from qa_final import reg, RES, sha256_arquivo  # noqa: E402

COD_DH = ('E01', 'A03', 'E06', 'A07')
RX_DH = re.compile(r'\d{2}/\d{2}/\d{4} \d{2}:\d{2}:\d{2}')
RX_SEM = re.compile(r'\b\d{2}/\d{2}/\d{4} \d{2}:\d{2}(?!:\d)')


def executar(produto, caminho, saida):
    hash0 = sha256_arquivo(caminho)
    tmp = tempfile.mkdtemp(prefix='eqqa094_')
    copia = os.path.join(tmp, os.path.basename(caminho))
    shutil.copy2(caminho, copia)
    js = os.path.join(tmp, 'equiv_qa.json')
    est = os.path.join(AQUI, '..', 'estudos')
    r = subprocess.run([sys.executable, '-u', '-W', 'ignore', 'equivalencia_qa.py', produto, copia, js, '--fixtures'],
                       cwd=est, capture_output=True, text=True, encoding='utf-8', errors='replace')
    print(r.stdout[-3000:], flush=True)
    if r.returncode != 0 or not os.path.exists(js):
        reg('Q01 Equivalência do QA (VBA = PQ) com as fixtures', False,
            {'erro': 'equivalencia_qa.py nao concluiu', 'rc': r.returncode, 'stderr': r.stderr[-1500:]})
    else:
        d = json.load(open(js, encoding='utf-8'))
        ok1 = (not d.get('so_no_pq') and not d.get('so_no_vba') and d.get('ordem_das_partes_igual') is True
               and d.get('linhas', {}).get('pq') == d.get('linhas', {}).get('vba')
               and all((d.get(k) or {}).get('ok') is True for k in ('adr082', 'adr088', 'adr091')))
        reg('Q01 Equivalência do QA: o mesmo conjunto de achados nos dois motores (com o DETALHE), a mesma ordem das partes, '
            'e os casos do ADR-082/088/091 conferidos', ok1,
            {'linhas': d.get('linhas'), 'so_no_pq': d.get('so_no_pq'), 'so_no_vba': d.get('so_no_vba'),
             'ordem_igual': d.get('ordem_das_partes_igual'),
             'ok': {k: (d.get(k) or {}).get('ok') for k in ('adr082', 'adr088', 'adr091')}})
        n_pq = {c: d.get('por_codigo_pq', {}).get(c, 0) for c in COD_DH}
        n_vba = {c: d.get('por_codigo_vba', {}).get(c, 0) for c in COD_DH}
        total = sum(n_pq.values())
        dets = d.get('detalhes_dh') or {}
        ruins = {m_: [x[:160] for x in dets.get(m_, []) if RX_SEM.search(x) or not RX_DH.search(x)] for m_ in ('pq', 'vba')}
        # 0 achados com data-hora = o texto novo nao foi exercitado: FAIL (nao passar calado)
        reg('Q02 Achados com data-hora no DETALHE (E01/A03/E06/A07): existem nesta base, mesma contagem nos dois motores e '
            'TODOS com dd/mm/aaaa hh:mm:ss (nenhum sem segundos), no PQ e no VBA',
            total > 0 and n_pq == n_vba and not ruins['pq'] and not ruins['vba'],
            {'por_codigo_pq': n_pq, 'por_codigo_vba': n_vba, 'conferidos': {m_: len(dets.get(m_, [])) for m_ in ('pq', 'vba')},
             'fora_do_formato': {m_: v[:5] for m_, v in ruins.items()}, 'exemplo': (dets.get('pq') or [''])[0][:200]})
    shutil.rmtree(tmp, ignore_errors=True)
    hash1 = sha256_arquivo(caminho)
    reg('Y00 Segurança do teste: arquivo de entrada intacto (SHA-256 igual antes e depois; nada salvo nele)', hash0 == hash1,
        {'arquivo': caminho, 'sha256_antes': hash0, 'sha256_depois': hash1})
    if saida:
        with open(saida, 'w', encoding='utf-8') as f:
            json.dump(RES, f, ensure_ascii=False, indent=1, default=str)
    falhas = [x for x in RES if x['resultado'] == 'FAIL']
    print(f'\n=== {produto} ADR-094 equivalência do QA (VBA = PQ): {len(RES) - len(falhas)} PASS / {len(falhas)} FAIL ===',
          flush=True)
    return falhas


if __name__ == '__main__':
    falhas = executar(sys.argv[1], os.path.abspath(sys.argv[2]), sys.argv[3] if len(sys.argv) > 3 else None)
    sys.exit(1 if falhas else 0)
