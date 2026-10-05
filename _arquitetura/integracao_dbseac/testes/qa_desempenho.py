# -*- coding: utf-8 -*-
"""QA de DESEMPENHO da troca de analito (ADR-062) -- trava permanente contra regressao.

Uso: python qa_desempenho.py <Bioquimica|Hematologia> <copia.xlsm> [saida.json]

Mede pelo caminho do usuario (medir_spinner.py, Excel oculto) e reprova acima dos tetos:
  1a troca depois do login  <= 600 ms   (era 1,7-2,3 s antes do aquecimento no login)
  trocas seguintes, mediana <= 300 ms   (era ~320 ms oculto / ~425 ms visivel)
  trocas seguintes, maximo  <= 600 ms
  escolha pela lista (C3)   <= 450 ms   (era 480-725 ms)
  spinner limitado aos analitos cadastrados (Max = n)
Reprova tambem se uma formula nova ou uma varredura do banco voltar ao clique (o erro do
eqProvedor no ADR-050): e exatamente o que estes tetos pegam.
"""
import json
import os
import sys

AQUI = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, AQUI)
import medir_spinner  # noqa: E402

TETOS = {'frio_ms': 600, 'quente_mediana_ms': 300, 'quente_max_ms': 600, 'lista_mediana_ms': 450}
RES = []


def reg(teste, ok, evid):
    RES.append({'teste': teste, 'resultado': 'PASS' if ok else 'FAIL', 'evidencia': evid})
    print(('PASS ' if ok else 'FAIL ') + teste + ' -- ' + json.dumps(evid, ensure_ascii=False, default=str)[:500], flush=True)


def executar(produto, caminho, saida):
    r = medir_spinner.medir(produto, caminho, 12, False)
    reg('P01 1ª troca de analito depois do login (motor aquecido no login) dentro do teto',
        r['frio']['total_ms'] <= TETOS['frio_ms'], {'frio_ms': r['frio']['total_ms'], 'aquecer_no_login_s': r.get('aquecer_s'),
                                                     'teto_ms': TETOS['frio_ms']})
    reg('P02 Trocas seguintes pelo spinner: mediana e máximo dentro do teto',
        r['quente_total']['mediana_ms'] <= TETOS['quente_mediana_ms'] and r['quente_total']['max_ms'] <= TETOS['quente_max_ms'],
        {'quente': r['quente_total'], 'componentes_ms': r['componentes_ms'],
         'tetos': [TETOS['quente_mediana_ms'], TETOS['quente_max_ms']]})
    reg('P03 Escolha do analito pela lista (C3) dentro do teto',
        r['lista']['mediana_ms'] <= TETOS['lista_mediana_ms'], {'lista': r['lista'], 'teto_ms': TETOS['lista_mediana_ms']})
    reg('P04 Spinner limitado aos analitos cadastrados (não cai em posição vazia)',
        r['spinner_min_max'] == [1, r['analitos']], {'spinner': r['spinner_min_max'], 'analitos': r['analitos']})
    if saida:
        with open(saida, 'w', encoding='utf-8') as f:
            json.dump({'resultados': RES, 'medicao': r}, f, ensure_ascii=False, indent=1, default=str)
    falhas = [x for x in RES if x['resultado'] == 'FAIL']
    print(f'\n=== {produto} desempenho: {len(RES) - len(falhas)} PASS / {len(falhas)} FAIL ===', flush=True)
    return falhas


if __name__ == '__main__':
    falhas = executar(sys.argv[1], os.path.abspath(sys.argv[2]), sys.argv[3] if len(sys.argv) > 3 else None)
    sys.exit(1 if falhas else 0)
