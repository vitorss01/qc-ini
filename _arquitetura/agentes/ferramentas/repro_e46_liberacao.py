# -*- coding: utf-8 -*-
"""repro_e46_liberacao.py -- Tester/QA, validacao da entrega 46 (24/11/2026). Cenario PROPRIO da Liberacao amarrada
ao RUN, NAO coberto por qa_liberacao_run (L01-L03) nem por qa_fluxo_manual (F05e/F07b), mais a integridade da copia.
Roda numa COPIA da producao; nada e salvo. Dados de teste: usuario local QA_TESTER46 (senha sintetica, so na copia),
manuais com MOTIVO 'TESTE-AUDITORIA' (valor = mediana dos N1 ativos do escopo).

Uso: python repro_e46_liberacao.py <Bioquimica|Hematologia> <copia.xlsm> <saida.json>   (REPRO_LOGIN=1: estado de sessao com login)

  T00 integridade da copia como foi entregue (antes de qualquer ATUALIZAR): STATUS_ANALITICO so os documentados;
      todo par (ID, analito) da aba Inativar esta INATIVADO na tblCQ_Final; nenhum nao-ATIVO participa da estatistica
  T01 assinatura REAL (mSeguranca.AssinarCom, usuario local ANALISTA): INICIADO + FINALIZADO num RUN do meio da lista
      e INICIADO no 1o RUN; Audit_Log ASSINATURA_* com 'RUN <n>' = coluna A da linha assinada
  T02 DOIS manuais em dois dias passados sem resultado do lote ANTES do RUN do meio -> 2 RUNs novos; a assinatura do meio
      desce 2 linhas e fica ao lado do MESMO RUN; a do 1o RUN nao se mexe; nenhuma linha assinada sem assinatura nova;
      nenhuma LIBERACAO_ASSINATURA_SEM_CORRIDA; 2 RESULTADO_MANUAL_INCLUIDO
  T03 troca do lote em uso ida e volta (mApp.UsarLote -> outro lote -> lote original): a LiberStore guarda o RUN
      (coluna E) e as assinaturas voltam ao lado dos MESMOS RUNs
  T04 FlushLoteAtual (o caminho do BeforeSave): no bloco da LiberStore, cada assinatura tem na coluna E o RUN da linha
  T05 retirada dos dois manuais (Delete) + ATUALIZAR: lista de RUNs = a inicial e as assinaturas nas linhas iniciais
  T06 Audit_Log: as linhas ASSINATURA_* continuam apontando os RUNs que hoje estao ao lado das assinaturas
  Y00 arquivo de entrada intacto
"""
import collections
import datetime as dt
import hashlib
import json
import os
import statistics
import sys
import traceback

TESTES = os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', '..', 'integracao_dbseac', 'testes')
sys.path.insert(0, os.path.abspath(TESTES))
from qa_final import QA, reg, RES, sha256_arquivo  # noqa: E402
from qa_fluxo_manual import SENHA, apagar_linha, nucleo_lote, TB_INAT, TB_MAN  # noqa: E402  (so utilidades COM)

LOGIN = 'QA_TESTER46'
SENHA_T = 'QaTester46#Sint'          # sintetica, so existe nesta copia
MOTIVO = 'TESTE-AUDITORIA'
ABA = 'Liberação'


def txt(v):
    if isinstance(v, float) and v == int(v):
        return str(int(v))
    return '' if v is None else str(v).strip()


def iint(v):
    try:
        return int(float(v))
    except (TypeError, ValueError):
        return None


def dia_de_run(run):
    return dt.date(2000 + run // 1000000, (run // 10000) % 100, (run // 100) % 100)


def executar(produto, caminho, saida):
    h0 = sha256_arquivo(caminho)
    q = QA(produto, caminho)
    try:
        if os.environ.get('REPRO_LOGIN') == '1':
            # o estado de uma sessao com login (LockApp/UnlockApp -> ReprotectAll: UserInterfaceOnly nas abas). Sem ele
            # (automacao pura), a Hema perde as assinaturas na troca de lote: ver diag_e46_liberstore.py
            q.run('mSeguranca.ReprotectAll', teto=300)
        lib = q.ws(ABA)
        aud = q.ws('Audit_Log')

        def ler_lib():
            return [[txt(x) for x in r] for r in lib.Range('A4:F203').Value]

        def runs_de(l):
            return [iint(r[0]) for r in l if iint(r[0])]

        def assin_por_run(l):
            return {iint(r[0]): tuple(r[2:6]) for r in l if iint(r[0]) and any(r[2:6])}

        def assin_por_linha(l):
            return {i: (iint(r[0]), tuple(r[2:6])) for i, r in enumerate(l) if any(r[2:6])}

        def ult_aud():
            return int(q.run('mAuditoria.UltimaLinhaAudit'))

        def aud_desde(r0):
            r1 = ult_aud()
            if r1 <= r0:
                return []
            return [(txt(l[6]), txt(l[23])) for l in aud.Range(aud.Cells(r0 + 1, 1), aud.Cells(r1, 24)).Value]

        def atualizar(rot):
            r = str(q.run('mIntegracao.AtualizarSoEntradas', '', teto=1800))
            if not r.startswith('OK|'):
                raise RuntimeError(f'ATUALIZAR ({rot}): {r[:200]}')

        # ---------------------------------------------------------------- T00 integridade como entregue
        fin = q.ler('tblCQ_Final')
        st = collections.Counter(txt(r.get('STATUS_ANALITICO')) for r in fin)
        idx = {(txt(r['ID_REGISTRO']).upper(), txt(r['ANALITO']).upper()): r for r in fin}
        inat = [r for r in q.ler(TB_INAT) if txt(r.get('ID_REGISTRO'))]
        nao_inat, nao_achado = [], 0
        for r in inat:
            k = (txt(r['ID_REGISTRO']).upper(), txt(r.get('ANALITO')).upper())
            f = idx.get(k)
            if f is None:
                nao_achado += 1
            elif txt(f.get('STATUS_ANALITICO')) != 'INATIVADO':
                nao_inat.append({'id': k[0], 'analito': k[1], 'status': txt(f.get('STATUS_ANALITICO'))})
        inat_part = sum(1 for r in fin if txt(r.get('STATUS_ANALITICO')) != 'ATIVO'
                        and txt(r.get('PARTICIPA_ESTATISTICA')).upper() == 'SIM')
        # status do ETL (src/comum/mEtlCQ.bas ~l.674-692 e ADR-091/092): so ATIVO participa
        validos = {'ATIVO', 'INATIVADO', 'IMPLAUSIVEL', 'LOTE_FORA_DO_CQ', 'LOTE_OUTRO_MATERIAL', 'MANUAL_INCOMPLETO',
                   'DUPLICIDADE_ORIGEM', 'CONFLITO_MANUAL', 'SUBSTITUIDO_MANUAL', 'SEM_DATA_HORA', 'SEM_VALOR'}
        reg('T00 Integridade da copia entregue: STATUS so os documentados; Inativar -> INATIVADO; nenhum nao-ATIVO na '
            'estatistica', set(st) <= validos and not nao_inat and inat_part == 0,
            {'status': dict(st), 'linhas_inativar': len(inat), 'inativar_sem_resultado': nao_achado,
             'inativar_nao_INATIVADO': nao_inat[:10], 'nao_ATIVO_participando': inat_part})

        # ---------------------------------------------------------------- preparo
        modo0 = q.cfg('MODO_FONTE', 'HISTORICO')
        q.cfg('MOTOR_DADOS', 'VBA')
        atualizar('base')
        lote = txt(q.run('mDados.LoteAtivoCore'))
        eqf = txt(q.run('mEstatistica.EquipFiltro')).upper()
        l0 = ler_lib()
        runs0 = runs_de(l0)
        if len(runs0) < 6:
            reg('T01 pre-condicao ausente: menos de 6 corridas na Liberacao', False, {'lote': lote, 'n': len(runs0)})
            return
        sig0 = assin_por_linha(l0)
        fin = q.ler('tblCQ_Final')
        do_lote = [r for r in fin if nucleo_lote(txt(r['LOTE'])) == lote
                   and (not eqf or txt(r.get('EQUIPAMENTO')).upper() == eqf)]
        dias_com = {r['DATA'].date() for r in do_lote if hasattr(r.get('DATA'), 'year')}
        primeiro = dia_de_run(runs0[0])
        # RUN do meio: o mais a frente na lista que tenha ao menos 2 dias vazios entre o 1o RUN e ele
        alvo, dias_vazios = None, []
        for j in range(len(runs0) // 2, len(runs0) - 1):
            d_alvo = dia_de_run(runs0[j])
            vaz = [primeiro + dt.timedelta(days=k) for k in range(1, (d_alvo - primeiro).days)
                   if primeiro + dt.timedelta(days=k) not in dias_com]
            vaz = [d for d in vaz if runs0[0] < int(d.strftime('%y%m%d')) * 100 + 1 < runs0[j]]
            if len(vaz) >= 2:
                alvo, dias_vazios = j, vaz
                break
        if alvo is None:
            reg('T02 pre-condicao ausente: sem 2 dias vazios antes de um RUN do meio', False, {'lote': lote})
            return
        r_meio, r_pri = runs0[alvo], runs0[0]
        # dois dias vazios bem separados entre o 1o RUN e o do meio
        d1, d2 = dias_vazios[len(dias_vazios) // 3], dias_vazios[(2 * len(dias_vazios)) // 3]
        if d1 == d2:
            d1, d2 = dias_vazios[0], dias_vazios[-1]

        # usuario local de teste (ANALISTA) com a rubrica em texto
        us = q.ws('Usuarios')
        us.Unprotect(SENHA)
        iu = next(i for i in range(4, 54) if not us.Cells(i, 1).Value)
        us.Cells(iu, 1).Value = LOGIN
        us.Cells(iu, 2).Value = 'QA Tester 46'
        us.Cells(iu, 3).Value = 'ANALISTA'
        us.Cells(iu, 4).Value = hashlib.sha256(SENHA_T.encode()).hexdigest()
        us.Cells(iu, 5).Value = 'RUBRICA-QA46'
        us.Protect(SENHA, False, True, True, True)

        # ---------------------------------------------------------------- T01 assinatura real
        a0 = ult_aud()
        res = {}
        for (j, col, tipo) in ((alvo, 'C', 'INICIADO'), (alvo, 'E', 'FINALIZADO'), (0, 'C', 'INICIADO')):
            res[f'{col}{4 + j}'] = str(q.run('mSeguranca.AssinarCom', lib.Range(f'{col}{4 + j}'), LOGIN, SENHA_T, tipo,
                                             teto=300))
        l1 = ler_lib()
        s1 = assin_por_run(l1)
        ev1 = aud_desde(a0)
        ass1 = [d for a, d in ev1 if a.startswith('ASSINATURA_')]
        ok1 = (all(v == 'OK' for v in res.values()) and r_meio in s1 and r_pri in s1
               and LOGIN in s1[r_meio][1] and LOGIN in s1[r_meio][3] and LOGIN in s1[r_pri][1]
               and sum(f'RUN {r_meio}' in d for d in ass1) == 2 and sum(f'RUN {r_pri}' in d for d in ass1) == 1)
        reg('T01 Assinatura real (AssinarCom) INICIADO+FINALIZADO no RUN do meio e INICIADO no 1o RUN; Audit_Log com o RUN',
            ok1, {'lote': lote, 'equip': eqf, 'corridas': len(runs0), 'RUN_meio': r_meio, 'linha_meio': 4 + alvo,
                  'RUN_1o': r_pri, 'retornos': res, 'trilhas_ASSINATURA': len(ass1),
                  'assinaturas_previas_na_producao': len(sig0)})
        sig_ini = assin_por_run(l1)
        linha_ini = {iint(r[0]): i for i, r in enumerate(l1) if iint(r[0]) in sig_ini}

        # ---------------------------------------------------------------- T02 dois manuais antes do RUN do meio
        cont = collections.Counter((txt(r['ANALITO']), txt(r['EQUIPAMENTO']), txt(r.get('MATRIZ')), txt(r['LOTE']))
                                   for r in do_lote if txt(r.get('STATUS_ANALITICO')) == 'ATIVO' and iint(r['NIVEL']) == 1)
        (an, eq, mt, lt), _ = cont.most_common(1)[0]
        vals = [float(r['RESULTADO']) for r in do_lote if txt(r['ANALITO']) == an and iint(r['NIVEL']) == 1
                and txt(r.get('STATUS_ANALITICO')) == 'ATIVO' and isinstance(r.get('RESULTADO'), (int, float))]
        v = round(statistics.median(vals), 4)
        q.ex.xl.EnableEvents = True
        linhas_man = []
        for d in (d1, d2):
            val = {'DATA': dt.datetime.combine(d, dt.time()), 'HORA': 0.5, 'EQUIPAMENTO': eq}
            if mt:
                val['MATRIZ'] = mt
            val.update({'LOTE': lt, 'NIVEL': 1, 'ANALITO': an, 'RESULTADO': v, 'MOTIVO': MOTIVO})
            linhas_man.append(q.escrever_linha(TB_MAN, val))
        a1 = ult_aud()
        atualizar('T02')
        l2 = ler_lib()
        runs2 = runs_de(l2)
        s2 = assin_por_run(l2)
        novos = sorted(set(runs2) - set(runs0))
        lin2 = {iint(r[0]): i for i, r in enumerate(l2) if iint(r[0]) in s2}
        ev2 = aud_desde(a1)
        acs2 = collections.Counter(a for a, _ in ev2)
        so_essas = set(assin_por_linha(l2)) == {lin2[k] for k in s2}
        ok2 = (len(novos) == 2 and s2 == sig_ini and lin2.get(r_meio) == linha_ini[r_meio] + 2
               and lin2.get(r_pri) == linha_ini[r_pri] and so_essas
               and acs2.get('LIBERACAO_ASSINATURA_SEM_CORRIDA', 0) == 0 and acs2.get('RESULTADO_MANUAL_INCLUIDO', 0) == 2
               and txt(l2[linha_ini[r_meio]][0]) != str(r_meio))
        reg('T02 Dois RUNs novos antes do RUN assinado: a assinatura desce 2 linhas com o MESMO RUN; a do 1o RUN fica; '
            'sem trilha de orfa; 2 RESULTADO_MANUAL_INCLUIDO', ok2,
            {'analito': an, 'dias_manuais': [str(d1), str(d2)], 'valor': v, 'RUNs_novos': novos,
             'linha_meio_antes': 4 + linha_ini[r_meio], 'linha_meio_depois': None if r_meio not in lin2 else 4 + lin2[r_meio],
             'RUN_na_linha_antiga_agora': l2[linha_ini[r_meio]][0], 'assinaturas_iguais_por_RUN': s2 == sig_ini,
             'linhas_assinadas_so_as_esperadas': so_essas, 'acoes_audit': dict(acs2)})

        # ---------------------------------------------------------------- T03 troca de lote ida e volta
        cad = [txt(c.Value) for c in q.wb.Names('regLoteCol').RefersToRange]
        outro = next((c for c in cad if c and c != lote), None)
        if outro is None:
            reg('T03 pre-condicao ausente: so um lote no cadastro', False, {'lote': lote})
        else:
            q.run('mApp.UsarLote', outro, teto=1800)
            lote_mid = txt(q.run('mDados.LoteAtivoCore'))
            l_outro = ler_lib()
            q.run('mApp.UsarLote', lote, teto=1800)
            l3 = ler_lib()
            s3 = assin_por_run(l3)
            lin3 = {iint(r[0]): i for i, r in enumerate(l3) if iint(r[0]) in s3}
            nosso_no_outro = sum(1 for r in l_outro if LOGIN in r[3] or LOGIN in r[5])
            ok3 = (lote_mid == outro and txt(q.run('mDados.LoteAtivoCore')) == lote and s3 == sig_ini
                   and lin3 == lin2 and runs_de(l3) == runs2 and nosso_no_outro == 0)
            reg('T03 Troca do lote em uso ida e volta: assinaturas voltam ao lado dos MESMOS RUNs (LiberStore com o RUN); '
                'nada do teste aparece no outro lote', ok3,
                {'outro_lote': outro, 'em_uso_no_meio': lote_mid, 'assinaturas_iguais': s3 == sig_ini,
                 'linhas_iguais_ao_T02': lin3 == lin2, 'lista_igual_T02': runs_de(l3) == runs2,
                 'assinaturas_do_teste_no_outro_lote': nosso_no_outro})

        # ---------------------------------------------------------------- T04 FlushLoteAtual -> LiberStore coluna E
        q.run('mLotes.FlushLoteAtual', teto=600)
        cad = [txt(c.Value) for c in q.wb.Names('regLoteCol').RefersToRange]
        ib = cad.index(lote) + 1
        nlb = 200
        rb0 = 2 + (ib - 1) * nlb
        stl = q.ws('LiberStore')
        blk = [[txt(x) for x in r] for r in stl.Range(stl.Cells(rb0, 1), stl.Cells(rb0 + nlb - 1, 5)).Value]
        pares = {iint(r[4]): tuple(r[0:4]) for r in blk if any(r[0:4])}
        reg('T04 FlushLoteAtual: no bloco do lote na LiberStore, cada assinatura tem o seu RUN na coluna E', pares == sig_ini,
            {'bloco': ib, 'linha0': rb0, 'RUNs_com_assinatura_no_bloco': sorted(k for k in pares if k),
             'esperado': sorted(sig_ini)})

        # ---------------------------------------------------------------- T05 retirada dos manuais
        for rr in sorted(linhas_man, reverse=True):
            apagar_linha(q, TB_MAN, rr)
        a5 = ult_aud()
        atualizar('T05')
        l5 = ler_lib()
        s5 = assin_por_run(l5)
        lin5 = {iint(r[0]): i for i, r in enumerate(l5) if iint(r[0]) in s5}
        acs5 = collections.Counter(a for a, _ in aud_desde(a5))
        ok5 = runs_de(l5) == runs0 and s5 == sig_ini and lin5 == linha_ini and acs5.get('RESULTADO_MANUAL_RETIRADO', 0) == 2
        reg('T05 Manuais retirados: lista = a inicial e as assinaturas nas linhas iniciais; 2 RESULTADO_MANUAL_RETIRADO',
            ok5, {'lista_igual': runs_de(l5) == runs0, 'assinaturas_iguais': s5 == sig_ini,
                  'linhas': {k: 4 + v for k, v in lin5.items()}, 'acoes_audit': dict(acs5)})

        # ---------------------------------------------------------------- T06 trilha x tela
        ev6 = aud_desde(a0)
        ass6 = [d for a, d in ev6 if a.startswith('ASSINATURA_')]
        runs_trilha = collections.Counter(int(d.split('RUN ')[1].split()[0]) for d in ass6 if 'RUN ' in d)
        tela = collections.Counter()
        for r in l5:
            if iint(r[0]):
                tela[iint(r[0])] += (1 if r[3] else 0) + (1 if r[5] else 0)
        tela = +tela
        reg('T06 Audit_Log: as trilhas ASSINATURA_* apontam os RUNs que estao ao lado das assinaturas na tela',
            runs_trilha == tela, {'trilha': dict(runs_trilha), 'tela': dict(tela)})
        q.cfg('MODO_FONTE', modo0)
    except Exception as e:                                # noqa: BLE001
        reg('Z99 Execucao sem excecao', False, {'erro': repr(e), 'trace': traceback.format_exc()[-1500:]})
    finally:
        q.fechar(salvar=False)
    h1 = sha256_arquivo(caminho)
    reg('Y00 Arquivo de entrada intacto (SHA-256)', h0 == h1, {'antes': h0[:16], 'depois': h1[:16]})
    if saida:
        with open(saida, 'w', encoding='utf-8') as f:
            json.dump(RES, f, ensure_ascii=False, indent=1, default=str)
    falhas = [r for r in RES if r['resultado'] == 'FAIL']
    print(f'\n=== {produto} repro_e46_liberacao: {len(RES) - len(falhas)} PASS / {len(falhas)} FAIL ===', flush=True)
    return falhas


if __name__ == '__main__':
    f = executar(sys.argv[1], os.path.abspath(sys.argv[2]), sys.argv[3] if len(sys.argv) > 3 else None)
    sys.exit(1 if f else 0)
