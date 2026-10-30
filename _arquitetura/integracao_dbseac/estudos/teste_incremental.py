# -*- coding: utf-8 -*-
"""ADR-079: sincronizacao incremental, idempotencia, falhas e reserva (PQ) numa COPIA ja instalada. Nada salvo.

Uso: python teste_incremental.py <Bioquimica|Hematologia> <copia.xlsm> <saida.json>

  S1 incremental  tira do recebimento os K mais recentes e 40 do meio (resultado que "chega atrasado"); o ATUALIZAR
                  tem de recebe-los de volta: mesmas linhas, mesmos valores, so o RECEBIDO_EM novo; a Principal volta
                  igual (fora o RECEBIDO_EM dessas linhas);
  S2 idempotente  ATUALIZAR de novo: 0 novos, OK_SEM_MUDANCA, nada gravado;
  S3 origem ruim  CAMINHO_DB_SEAC inexistente: ERRO, NADA gravado (tabelas identicas), Sync_Log = ERRO;
  S4 trava        MODO HISTORICO com 30% do recebimento apagado: a Principal encolheria -> ERRO, nada gravado;
  S5 recuperacao  volta SEAC: os 30% sao recebidos de novo e a Principal volta inteira;
  S6 reserva      MOTOR_DADOS = PQ: o caminho antigo continua funcionando e registra no Sync_Log;
Cada passo cronometrado. Sai com 1 se algum falhar.
"""
import json
import os
import random
import shutil
import sys
import time

AQUI = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(AQUI, '..', 'testes'))
sys.path.insert(0, os.path.join(AQUI, '..'))
from qa_final import QA  # noqa: E402
import instalar_integracao as ii  # noqa: E402

K_RECENTES = 1200
K_MEIO = 40


def tabela(q, nome):
    lo = q.lo(nome)
    cab = [c.Name for c in lo.ListColumns]
    lin = [list(r) for r in lo.DataBodyRange.Value2] if lo.ListRows.Count else []
    return cab, lin


def por_id(t, ignorar=()):
    cab, lin = t
    i = cab.index('ID_REGISTRO')
    keep = [j for j, c in enumerate(cab) if c not in ignorar]
    return {r[i]: tuple(None if r[j] in ('', None) else r[j] for j in keep) for r in lin}


def ultimo_log(q):
    cab, lin = tabela(q, 'tblSyncLog')
    return dict(zip(cab, lin[-1])) if lin else {}


def apagar_linhas(q, nome, linhas):
    """linhas = indices 1-based da tabela; apaga de baixo para cima (blocos contiguos de uma vez)."""
    lo = q.lo(nome)
    lo.Parent.Unprotect(ii.SENHA)
    linhas = sorted(set(linhas), reverse=True)
    n_antes = lo.ListRows.Count
    blocos, ini, fim = [], None, None
    for r in linhas:
        if fim is None:
            ini = fim = r
        elif r == ini - 1:
            ini = r
        else:
            blocos.append((ini, fim))
            ini = fim = r
    if fim is not None:
        blocos.append((ini, fim))
    for a, b in blocos:
        # Range.Resize(k) pelo COM vira Item(k) (uma celula): faixa explicita
        lo.Parent.Range(lo.DataBodyRange.Rows(a), lo.DataBodyRange.Rows(b)).Delete()
    if lo.ListRows.Count != n_antes - len(linhas):
        raise SystemExit(f'apagar_linhas: esperado {n_antes - len(linhas)} linhas, ficaram {lo.ListRows.Count}')


def main(produto, arq, saida):
    res, falhas = {'produto': produto, 'passos': {}}, []

    def passo(nome, ok, **info):
        lg = ultimo_log(q)
        info['tempos'] = {k: lg.get(k) for k in ('T_ORIGEM', 'T_RECEBIMENTO', 'T_CQ', 'T_VALIDAR', 'T_GRAVAR', 'T_QA',
                                                 'T_MOTOR', 'TOTAL_S')}
        res['passos'][nome] = dict(ok=bool(ok), **info)
        print(f'{"OK   " if ok else "FALHA"} {nome}: {info}', flush=True)
        if not ok:
            falhas.append(nome)

    q = QA(produto, arq)
    try:
        q.cfg('MOTOR_DADOS', 'VBA')
        # o DB_SEAC de producao recebe resultados enquanto o teste roda (entrega 33: 2 novos no meio do S1); os passos
        # comparam com a partida, entao a origem precisa ficar parada: copia congelada, so nesta copia de teste
        foto = os.path.join(os.environ['TEMP'], 'qcw', f'db_seac_foto_{produto}')
        os.makedirs(foto, exist_ok=True)
        origem = q.cfg('CAMINHO_DB_SEAC')
        caminho = os.path.join(foto, os.path.basename(str(origem)))
        shutil.copyfile(str(origem), caminho)
        q.cfg('CAMINHO_DB_SEAC', caminho)
        r, s = q.atualizar()                                    # estado de partida em dia
        rec0, cq0 = tabela(q, 'tblDB_Recebimento'), tabela(q, 'tblCQ_Final')
        n0 = len(rec0[1])
        res['partida'] = {'recebimento': n0, 'cq': len(cq0[1]), 's': round(s, 1)}

        # ---- S1 incremental
        random.seed(79)
        meio = random.sample(range(int(n0 * 0.80), n0 - K_RECENTES), K_MEIO)
        apagar_linhas(q, 'tblDB_Recebimento', list(range(n0 - K_RECENTES + 1, n0 + 1)) + meio)
        r, s = q.atualizar()
        lg = ultimo_log(q)
        rec1, cq1 = tabela(q, 'tblDB_Recebimento'), tabela(q, 'tblCQ_Final')
        a, b = por_id(rec0, ('RECEBIDO_EM',)), por_id(rec1, ('RECEBIDO_EM',))
        dif = [k for k in a if a.get(k) != b.get(k)]
        ca, cb = por_id(cq0, ('RECEBIDO_EM',)), por_id(cq1, ('RECEBIDO_EM',))
        difcq = [k for k in ca if ca.get(k) != cb.get(k)] + [k for k in cb if k not in ca]
        ordem_ok = rec1[1] == sorted(rec1[1], key=lambda x: (x[rec1[0].index('DATA_HORA')] or -1,
                                                              x[rec1[0].index('ID_ORIGEM')] or -1))
        passo('S1_incremental', lg.get('NOVOS') == K_RECENTES + K_MEIO and len(rec1[1]) == n0 and not dif and not difcq
              and ordem_ok, s=round(s, 1), novos=lg.get('NOVOS'), linhas=len(rec1[1]), recebimento_dif=len(dif),
              cq_dif=len(difcq), ordem_ok=ordem_ok, gravadas=lg.get('GRAVADAS'), amostra=dif[:3] + difcq[:3])

        # ---- S2 idempotente
        r, s = q.atualizar()
        lg = ultimo_log(q)
        passo('S2_idempotente', lg.get('NOVOS') == 0 and lg.get('RESULTADO') == 'OK_SEM_MUDANCA' and lg.get('GRAVADAS') == 0,
              s=round(s, 1), resultado=lg.get('RESULTADO'), novos=lg.get('NOVOS'), gravadas=lg.get('GRAVADAS'))

        # ---- S3 origem inacessivel
        antes = (tabela(q, 'tblDB_Recebimento'), tabela(q, 'tblCQ_Final'), tabela(q, 'tblQA_Integracao'))
        q.cfg('CAMINHO_DB_SEAC', os.path.join(os.environ['TEMP'], 'nao_existe', 'DB_SEAC.xlsm'))
        t0 = time.time()
        try:
            q.atualizar()
            erro = None
        except RuntimeError as e:
            erro = str(e)
        s = time.time() - t0
        depois = (tabela(q, 'tblDB_Recebimento'), tabela(q, 'tblCQ_Final'), tabela(q, 'tblQA_Integracao'))
        lg = ultimo_log(q)
        passo('S3_origem_inacessivel', erro is not None and antes == depois and lg.get('RESULTADO') == 'ERRO',
              s=round(s, 1), erro=(erro or '')[:160], tabelas_intactas=antes == depois, log=lg.get('RESULTADO'),
              log_erro=str(lg.get('ERRO'))[:120])
        q.cfg('CAMINHO_DB_SEAC', caminho)

        # ---- S4 trava contra encolher (historico + 30% apagado)
        q.cfg('MODO_FONTE', 'HISTORICO')
        n = q.lo('tblDB_Recebimento').ListRows.Count
        corte = int(n * 0.30)
        apagar_linhas(q, 'tblDB_Recebimento', list(range(n - corte + 1, n + 1)))
        antes = (tabela(q, 'tblDB_Recebimento'), tabela(q, 'tblCQ_Final'))
        try:
            q.atualizar()
            erro = None
        except RuntimeError as e:
            erro = str(e)
        depois = (tabela(q, 'tblDB_Recebimento'), tabela(q, 'tblCQ_Final'))
        passo('S4_trava_encolher', erro is not None and 'cairia' in erro and antes == depois,
              erro=(erro or '')[:200], tabelas_intactas=antes == depois)

        # ---- S5 recuperacao
        q.cfg('MODO_FONTE', 'SEAC')
        r, s = q.atualizar()
        lg = ultimo_log(q)
        rec5, cq5 = tabela(q, 'tblDB_Recebimento'), tabela(q, 'tblCQ_Final')
        a5 = por_id(rec5, ('RECEBIDO_EM',))
        c5 = por_id(cq5, ('RECEBIDO_EM',))
        passo('S5_recuperacao', len(rec5[1]) == n0 and a5 == a and c5 == ca, s=round(s, 1), novos=lg.get('NOVOS'),
              recebimento=len(rec5[1]), iguais_ao_inicio=(a5 == a, c5 == ca))

        # ---- S6 reserva PQ
        q.cfg('MOTOR_DADOS', 'PQ')
        r, s = q.atualizar()
        lg = ultimo_log(q)
        cq6 = por_id(tabela(q, 'tblCQ_Final'), ('RECEBIDO_EM',))
        passo('S6_reserva_pq', lg.get('MOTOR_DADOS') == 'PQ' and lg.get('RESULTADO') == 'OK' and cq6 == ca,
              s=round(s, 1), log=(lg.get('MOTOR_DADOS'), lg.get('RESULTADO')), cq_igual=cq6 == ca)
        q.cfg('MOTOR_DADOS', 'VBA')
    finally:
        q.fechar(salvar=False)
    res['falhas'] = falhas
    with open(saida, 'w', encoding='utf-8') as f:
        json.dump(res, f, ensure_ascii=False, indent=1, default=str)
    print('RESULTADO:', 'PASS' if not falhas else 'FAIL ' + ', '.join(falhas), flush=True)
    sys.exit(1 if falhas else 0)


if __name__ == '__main__':
    main(sys.argv[1], os.path.abspath(sys.argv[2]), sys.argv[3])
