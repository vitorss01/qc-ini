# -*- coding: utf-8 -*-
"""Linha de base do ATUALIZAR DADOS (reengenharia da camada de dados, 20/10/2026). Roda numa COPIA; nada e salvo.

Uso: python bench_atualizar.py <Bioquimica|Hematologia> <arquivo.xlsm> <saida.json> [--n 3] [--local <DB_SEAC.xlsm>]

Abre a copia como o QA (qa_final.QA), roda mIntegracao.AtualizarDadosAutomatico N vezes seguidas (o usuario
clicando de novo) e registra, por execucao:
  - tempo total de parede (Python, em volta da chamada) e os tempos por camada que o proprio VBA mede
    (tblDB_Recebimento, tblDB_Organizado, tblCQ_Final, tblQA_Integracao, motor);
  - o resto (preparar entradas, estado auditavel, lotes, auditoria) = total - soma das etapas;
  - resultados recebidos, novos, linhas da tblCQ_Final, data minima/maxima e linhas por mes (sem dado clinico);
  - os ciclos da tarefa agendada do DB_SEAC que se sobrepuseram (log da automacao, so data/hora/duracao).
--local: aponta CAMINHO_DB_SEAC para uma copia local do DB_SEAC (isola a leitura pela rede).
Tambem mede a abertura do arquivo e um SaveAs da copia (o usuario salva depois do ATUALIZAR).
"""
import datetime as dt
import json
import os
import re
import sys
import time

AQUI = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(AQUI, '..', 'testes'))
from qa_final import QA  # noqa: E402

# pasta de logs da automacao do DB_SEAC (caminho de rede do laboratorio: fora do repositorio publico)
LOG_DIR = os.environ.get('QC_LOG_DB_SEAC', '')
ETAPAS = ['tblDB_Recebimento', 'tblDB_Organizado', 'tblCQ_Final', 'tblQA_Integracao', 'motor']


def ciclos_agendados(ini, fim):
    """Ciclos da automacao do DB_SEAC que cruzaram [ini, fim] (fim do ciclo e duracao vem do log)."""
    out = []
    for d in {ini.date(), fim.date()}:
        arq = os.path.join(LOG_DIR, f'DB_SEAC_{d:%Y-%m-%d}.log')
        if not os.path.exists(arq):
            continue
        for linha in open(arq, encoding='utf-8', errors='replace'):
            m = re.match(r'(\d\d/\d\d/\d{4} \d\d:\d\d:\d\d) \| (\w+) \| (\d+)s', linha)
            if not m:
                continue
            t_fim = dt.datetime.strptime(m.group(1), '%d/%m/%Y %H:%M:%S')
            t_ini = t_fim - dt.timedelta(seconds=int(m.group(3)))
            if t_fim >= ini and t_ini <= fim:
                out.append({'inicio': t_ini.strftime('%H:%M:%S'), 'fim': t_fim.strftime('%H:%M:%S'),
                            'resultado': m.group(2), 's': int(m.group(3))})
    return out


def perfil_dados(q):
    """Contagens agregadas da tblCQ_Final (nada clinico): linhas, ids distintos, data min/max, linhas por mes."""
    lo = q.wb.Worksheets('Principal - Resultados').ListObjects('tblCQ_Final')
    cab = [str(c.Name) for c in lo.ListColumns]
    n = lo.ListRows.Count
    if n == 0:
        return {'linhas': 0}
    datas = lo.ListColumns('DATA_HORA').DataBodyRange.Value
    ids = lo.ListColumns('ID_REGISTRO').DataBodyRange.Value
    meses = {}
    ds = []
    for (v,) in datas:
        if hasattr(v, 'year'):
            ds.append(v)
            k = f'{v.year}-{v.month:02d}'
            meses[k] = meses.get(k, 0) + 1
    return {'linhas': n, 'colunas': len(cab), 'ids_distintos': len({str(i[0]) for i in ids}),
            'data_min': str(min(ds))[:19] if ds else None, 'data_max': str(max(ds))[:19] if ds else None,
            'por_mes': dict(sorted(meses.items())[-6:])}


def main(produto, arquivo, saida, n=3, local=None):
    res = {'produto': produto, 'arquivo': os.path.basename(arquivo), 'modo': 'local' if local else 'rede',
           'inicio': dt.datetime.now().isoformat(timespec='seconds'), 'execucoes': []}
    t0 = time.perf_counter()
    q = QA(produto, arquivo)
    res['abrir_s'] = round(time.perf_counter() - t0, 1)
    try:
        q.run('mSeguranca.ReprotectAll')
        if local:
            q.cfg('CAMINHO_DB_SEAC', local)
        res['modo_fonte'] = q.cfg('MODO_FONTE')
        res['perfil_antes'] = perfil_dados(q)
        for i in range(n):
            ini = dt.datetime.now()
            t = time.perf_counter()
            r = str(q.run('mIntegracao.AtualizarDadosAutomatico', teto=3600))
            total = time.perf_counter() - t
            fim = dt.datetime.now()
            etapas = {}
            m = re.search(r'Tempo: (.*)$', r, re.S)
            if m:
                for parte in m.group(1).split('·'):
                    mm = re.match(r'\s*(\S+) ([\d,\.]+)s', parte)
                    if mm:
                        etapas[mm.group(1)] = float(mm.group(2).replace(',', '.'))
            rec = re.search(r'Resultados recebidos: ([\d\.]+) \(novos nesta atualização: (-?\d+)\)', r)
            ex = {'n': i + 1, 'inicio': ini.strftime('%H:%M:%S'), 'fim': fim.strftime('%H:%M:%S'),
                  'ok': r.startswith('OK|'), 'total_s': round(total, 1), 'etapas_s': etapas,
                  'outros_s': round(total - sum(etapas.values()), 1),
                  'recebidos': int(rec.group(1).replace('.', '')) if rec else None,
                  'novos': int(rec.group(2)) if rec else None,
                  'ciclos_db_seac': ciclos_agendados(ini, fim)}
            if not ex['ok']:
                ex['erro'] = r[:400]
            res['execucoes'].append(ex)
            print(json.dumps(ex, ensure_ascii=False), flush=True)
        res['perfil_depois'] = perfil_dados(q)
        # o usuario salva depois: mede um SaveAs da copia (outro nome, ao lado)
        destino = os.path.splitext(arquivo)[0] + '_salvo.xlsm'
        q.ex.xl.DisplayAlerts = False
        t = time.perf_counter()
        q.wb.SaveCopyAs(destino)
        res['salvar_copia_s'] = round(time.perf_counter() - t, 1)
        res['tamanho_salvo_bytes'] = os.path.getsize(destino)
        os.remove(destino)
    finally:
        q.fechar(salvar=False)
    res['fim'] = dt.datetime.now().isoformat(timespec='seconds')
    with open(saida, 'w', encoding='utf-8') as f:
        json.dump(res, f, ensure_ascii=False, indent=1, default=str)
    print('RESUMO', json.dumps({k: res[k] for k in ('produto', 'modo', 'abrir_s', 'salvar_copia_s')}, ensure_ascii=False))


if __name__ == '__main__':
    a = sys.argv[1:]
    n = int(a[a.index('--n') + 1]) if '--n' in a else 3
    local = a[a.index('--local') + 1] if '--local' in a else None
    main(a[0], os.path.abspath(a[1]), a[2], n, local)
