# -*- coding: utf-8 -*-
"""Gera a tabela PASS/FAIL (markdown) a partir dos JSON de qa_final.py.
Uso: python tabela_qa.py qa_Hematologia.json [qa_Bioquimica.json ...] > tabela.md"""
import json
import sys


def resumo(ev, n=260):
    s = json.dumps(ev, ensure_ascii=False, default=str)
    s = s.replace('|', '¦').replace('\n', ' ')
    return s if len(s) <= n else s[:n] + '…'


def main(arquivos):
    for arq in arquivos:
        res = json.load(open(arq, encoding='utf-8'))
        nome = arq.replace('\\', '/').split('/')[-1].replace('qa_', '').replace('.json', '')
        ok = sum(1 for r in res if r['resultado'] == 'PASS')
        print(f'\n### {nome}: {ok} PASS / {len(res) - ok} FAIL\n')
        print('| Teste | Resultado | Evidência |')
        print('|---|---|---|')
        for r in res:
            print(f"| {r['teste']} | **{r['resultado']}** | {resumo(r['evidencia'])} |")


if __name__ == '__main__':
    main(sys.argv[1:])
