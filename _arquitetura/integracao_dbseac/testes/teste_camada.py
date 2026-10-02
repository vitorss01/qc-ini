# -*- coding: utf-8 -*-
"""Fase 1 -- prova da camada Power Query numa COPIA (nunca no arquivo de producao).

Uso: python teste_camada.py <Hematologia|Bioquimica> <copia.xlsm> [--limite N] [--refresh K] [--salvar]

Monta a camada (camada_dados.montar), mede cada consulta e confere as invariantes:
  ID_REGISTRO unico e nao vazio ............ chave estavel
  <=1 PARTICIPA=SIM por (equip,lote,nivel,analito,RUN)   (motor nao pode divergir)
  RUN preenchido em toda linha de corrida real
  todo recebido aparece na final (nada some)
  K refreshes seguidos: mesma contagem, mesmos IDs, mesmos RUN (idempotencia)
Imprime uma linha JSON por medida para o relatorio.
"""
import collections
import json
import os
import sys
import time

AQUI = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(AQUI, '..'))
sys.path.insert(0, os.path.join(AQUI, '..', '..', 'etapa1_multilote'))
import xlh  # noqa: E402
import pqlib  # noqa: E402
import camada_dados as cd  # noqa: E402


def arg(nome, padrao=None, tipo=str):
    if nome in sys.argv:
        return tipo(sys.argv[sys.argv.index(nome) + 1])
    return padrao


def conferir(rec, fin, rotulo):
    ids = [r['ID_REGISTRO'] for r in fin]
    vazios = sum(1 for i in ids if not i)
    dup = sum(1 for v in collections.Counter(ids).values() if v > 1)
    part = collections.Counter((r['EQUIPAMENTO'], r['LOTE'], r['NIVEL'], r['ANALITO'], r['RUN'])
                               for r in fin if r['PARTICIPA_ESTATISTICA'] == 'SIM')
    multi = sum(1 for v in part.values() if v > 1)
    sem_run = sum(1 for r in fin if r['STATUS_ANALITICO'] in ('ATIVO', 'INATIVADO', 'SEM_VALOR') and not r['RUN'])
    ids_rec = {r['ID_REGISTRO'] for r in rec}
    sumidos = len(ids_rec - set(ids))
    st = dict(collections.Counter(r['STATUS_ANALITICO'] for r in fin))
    pl = dict(collections.Counter(r['TIPO_PLOTAGEM_LJ'] for r in fin))
    ok = vazios == 0 and dup == 0 and multi == 0 and sem_run == 0 and sumidos == 0
    print(json.dumps({'medida': rotulo, 'linhas_final': len(fin), 'ids_vazios': vazios, 'ids_duplicados': dup,
                      'chave_motor_com_mais_de_1': multi, 'corrida_sem_run': sem_run, 'recebidos_sumidos': sumidos,
                      'status': st, 'plotagem': pl, 'OK': ok}, ensure_ascii=False, default=str), flush=True)
    return ok


def main():
    produto, copia = sys.argv[1], os.path.abspath(sys.argv[2])
    limite = arg('--limite', None, int)
    k = arg('--refresh', 0, int)
    xl = xlh.Excel()
    print('EXCEL_PID', xl.pid, flush=True)
    try:
        wb = xl.abrir(copia)
        t0 = time.time()
        tempos = cd.montar(wb, produto, carregar=True, limite=limite)
        print(json.dumps({'medida': 'carga', 'produto': produto, 'limite': limite, 'tempos': tempos,
                          'total_s': round(time.time() - t0, 1)}, ensure_ascii=False), flush=True)
        rec = pqlib.ler_tabela(pqlib.tabela(wb, 'tblDB_Recebimento'))
        if limite:
            rec = rec[:limite]
        fin = pqlib.ler_tabela(pqlib.tabela(wb, 'tblCQ_Final'))
        conferir(rec, fin, 'carga')
        qa = pqlib.ler_tabela(pqlib.tabela(wb, 'tblQA_Integracao'))
        print(json.dumps({'medida': 'qa', 'achados': dict(collections.Counter(f"{r['SEVERIDADE']} {r['CODIGO']}" for r in qa))},
                         ensure_ascii=False), flush=True)
        base = {r['ID_REGISTRO']: (r['RUN'], r['STATUS_ANALITICO']) for r in fin}
        nrec0 = len(pqlib.ler_tabela(pqlib.tabela(wb, 'tblDB_Recebimento')))
        for i in range(1, k + 1):
            ts = {}
            for q in cd.ORDEM_CARGA:
                lo = pqlib.tabela(wb, cd.SAIDAS[q][1])
                ts[q] = round(pqlib.atualizar(lo), 1)
            nrec = len(pqlib.ler_tabela(pqlib.tabela(wb, 'tblDB_Recebimento')))
            fin_i = pqlib.ler_tabela(pqlib.tabela(wb, 'tblCQ_Final'))
            atual = {r['ID_REGISTRO']: (r['RUN'], r['STATUS_ANALITICO']) for r in fin_i}
            iguais = atual == base
            print(json.dumps({'medida': f'refresh_{i}', 'tempos': ts, 'recebimento_linhas': nrec,
                              'recebimento_antes': nrec0, 'final_linhas': len(fin_i),
                              'ids_runs_status_identicos': iguais}, ensure_ascii=False), flush=True)
        if '--salvar' in sys.argv:
            wb.Save()
    finally:
        xl.fechar()


if __name__ == '__main__':
    main()
