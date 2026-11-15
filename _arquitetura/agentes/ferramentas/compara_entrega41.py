# -*- coding: utf-8 -*-
"""compara_entrega41.py -- Tester/QA, reteste04, VALIDAR: "nenhum valor de estatistica muda sem motivo".

Compara, com os MESMOS dados, o arquivo de antes da entrega 41 (_entrega_2026-11-14_0715/<produto>/producao_original)
e o arquivo instalado pela entrega 41 (_entrega_2026-11-14_0715/<produto>/QC_<produto>.xlsm). Cada um e copiado para
%TEMP%/qcw/auditoria/2026-11-14_reteste04/tester/ e aberto sozinho (um Excel por vez, so o meu PID e fechado); o Painel
recebe os mesmos filtros; le-se Estatistica A14:AU133 e Painel B7:J9 (Value2). Nada e salvo.

Uso: python -u compara_entrega41.py <Bioquimica|Hematologia>
"""
import datetime as dt
import json
import os
import shutil
import sys
import time

AQUI = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, AQUI)
import repro_comum as C  # noqa: E402

PROD = sys.argv[1]
BIO = PROD == 'Bioquimica'
PASTA = os.path.join(os.environ['TEMP'], 'qcw', 'auditoria', '2026-11-14_reteste04', 'tester')
ENT = os.path.join(C.REPO, '_entrega_2026-11-14_0715', PROD)
ARQ = f'QC_{PROD}.xlsm'
if BIO:
    ESTADOS = [('salvo', []),
               ('Glicose 8977 DIM1 2026', [('L4', 'DIMENSION 1'), ('C3', 'Glicose'), ('H3', "'8977"), ('K3', 2026),
                                           ('K4', 'Ano inteiro')]),
               ('Cloro 8974 DIM2 Tudo', [('L4', 'DIMENSION 2'), ('C3', 'Cloro'), ('H3', "'8974"), ('K4', 'Tudo (sem filtro)')])]
else:
    ESTADOS = [('salvo', []),
               ('WBC 2026', [('C3', 'WBC'), ('K3', 2026), ('K4', 'Ano inteiro')]),
               ('PLT Tudo', [('C3', 'PLT'), ('K4', 'Tudo (sem filtro)')])]


def ler(caminho, rot):
    dst = os.path.join(PASTA, f'{PROD}_cmp_{rot}.xlsm')
    shutil.copy2(caminho, dst)
    h = C.sha(dst)
    qa = C.qa_final.QA(PROD, dst)
    pid = qa.ex.pid
    out = {'sha12': h, 'pid': pid, 'estados': {}}
    try:
        p = qa.ws('Painel')
        p.Activate()
        qa.ex.esperar()
        for nome, seq in ESTADOS:
            for c, v in seq:
                C.por(qa, p, c, v, local=False)
            time.sleep(1)
            qa.ex.xl.Calculate()
            out['estados'][nome] = {'estat': qa.ws('Estatística').Range('A14:AU133').Value2,
                                    'painel': p.Range(f'B7:J{6 + (2 if BIO else 3)}').Value2,
                                    'filtro': [str(p.Range(c).Value) for c in ('C3', 'H3', 'K3', 'K4')]}
    finally:
        qa.fechar(salvar=False)
        os.remove(dst)
    return out


def main():
    t0 = time.time()
    a = ler(os.path.join(ENT, 'producao_original', ARQ), 'antes41')
    b = ler(os.path.join(ENT, ARQ), 'depois41')
    res = {'antes': {'sha12': a['sha12'], 'pid': a['pid']}, 'depois': {'sha12': b['sha12'], 'pid': b['pid']}, 'estados': {}}
    for nome in a['estados']:
        ea, eb = a['estados'][nome], b['estados'][nome]
        dif = []
        for parte in ('estat', 'painel'):
            for i, (la, lb) in enumerate(zip(ea[parte], eb[parte])):
                for j, (x, y) in enumerate(zip(la, lb)):
                    if x != y:
                        dif.append([parte, i, j, x, y])
        nnum = sum(1 for l in ea['estat'] for x in l if isinstance(x, float))
        res['estados'][nome] = {'filtro_antes': ea['filtro'], 'filtro_depois': eb['filtro'], 'celulas_numericas': nnum,
                                'diferencas': len(dif), 'exemplos': dif[:15],
                                # fora da coluna B das linhas sem analito (OC-0049 da Hema), com a linha da aba e o analito
                                'outras': [[p_, i + 14, j + 1, ea['estat'][i][0], x, y] if p_ == 'estat' else [p_, i, j, x, y]
                                           for p_, i, j, x, y in dif if not (p_ == 'estat' and j == 1 and i >= 84)]}
    res['segundos'] = round(time.time() - t0)
    s = json.dumps(res, ensure_ascii=False, indent=1, default=str)
    open(os.path.join(PASTA, f'compara_entrega41_{PROD}.json'), 'w', encoding='utf-8').write(s)
    print(s[:4000])


if __name__ == '__main__':
    main()
