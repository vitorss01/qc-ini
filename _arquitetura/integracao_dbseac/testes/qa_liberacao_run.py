# -*- coding: utf-8 -*-
"""QA da Liberacao amarrada ao RUN (correcao P1 da entrega 46, 23-24/11/2026) -- os dois caminhos de borda que o
qa_fluxo_manual (F05e/F07b) nao exercita. Roda numa COPIA instalada; nada e salvo.

Uso: python qa_liberacao_run.py <Bioquimica|Hematologia> <copia_instalada.xlsm> <saida.json>

  L01 corrida assinada que PERDE todo resultado ativo do lote em uso (todos inativados): a assinatura nao some -- vai
      para o fim da lista, ao lado do mesmo RUN, com "(sem resultado ativo)", aviso em H2 e UMA linha
      LIBERACAO_ASSINATURA_SEM_CORRIDA; a assinatura da ULTIMA corrida listada continua ao lado do seu RUN (com o lote
      acima de 200 corridas, a linha da reserva sai da corrida mais nova SEM assinatura)
  L02 um 2o ATUALIZAR nao repete a trilha nem muda a lista
  L03 desfazer as inativacoes: a corrida volta ao seu lugar com a assinatura; a lista volta a ser a do inicio
  Y00 arquivo de entrada intacto
"""
import json
import os
import sys
import traceback

AQUI = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, AQUI)
from qa_final import QA, reg, RES, sha256_arquivo  # noqa: E402
from qa_fluxo_manual import lib_ler, lib_gravar, apagar_linha, nucleo_lote, TB_INAT  # noqa: E402

SEM = '(sem resultado ativo)'
MARCA = 'QA_LIBERACAO_RUN: marcador do RUN '


def iint(v):
    try:
        return int(float(v))
    except (TypeError, ValueError):
        return None


def txt(v):
    if isinstance(v, float) and v == int(v):
        return str(int(v))
    return '' if v is None else str(v).strip()


def executar(produto, caminho, saida):
    h0 = sha256_arquivo(caminho)
    q = QA(produto, caminho)
    try:
        def aud0():
            return int(q.run('mAuditoria.UltimaLinhaAudit'))

        def acoes_desde(r0):
            al = q.ws('Audit_Log')
            r1 = aud0()
            return [txt(l[6]) for l in al.Range(al.Cells(r0 + 1, 1), al.Cells(r1, 7)).Value] if r1 > r0 else []

        def atualizar():
            r = str(q.run('mIntegracao.AtualizarSoEntradas', '', teto=1800))
            if not r.startswith('OK|'):
                raise RuntimeError('ATUALIZAR: ' + r[:200])

        def onde(lib, run):
            return next((i for i, l in enumerate(lib) if iint(l[0]) == run), None)

        atualizar()                                      # a copia pode ter entradas ainda nao aplicadas
        lote = txt(q.ws('Configuração').Range('C20').Value)
        eqf = txt(q.run('mEstatistica.EquipFiltro')).upper()
        lib0 = lib_ler(q)
        runs0 = [iint(l[0]) for l in lib0 if iint(l[0])]
        if len(runs0) < 3:
            reg('L01 pre-condicao ausente: a Liberacao do lote em uso tem menos de 3 corridas', False,
                {'lote': lote, 'corridas': len(runs0)})
            return
        fin = q.ler('tblCQ_Final')
        ativos = {}
        for r in fin:
            if txt(r.get('STATUS_ANALITICO')) == 'ATIVO' and nucleo_lote(txt(r['LOTE'])) == lote \
                    and (not eqf or txt(r.get('EQUIPAMENTO')).upper() == eqf) and iint(r.get('RUN')):
                ativos.setdefault(iint(r['RUN']), []).append(r)
        # a corrida que vai perder os ativos: a de MENOS resultados, longe das pontas da lista
        cand = [x for x in runs0[1:-1] if x in ativos]
        rx = min(cand, key=lambda x: (len(ativos[x]), x))
        rult = runs0[-1]
        ix, iu = onde(lib0, rx), onde(lib0, rult)
        lib_gravar(q, ix, 'QA', MARCA + str(rx))
        lib_gravar(q, iu, 'QA', MARCA + str(rult))
        # inativa todos os ativos da corrida rx (linhas na aba Inativar)
        linhas_inat = []
        for r in ativos[rx]:
            linhas_inat.append(q.escrever_linha(TB_INAT, {'ID_REGISTRO': txt(r['ID_REGISTRO']), 'ANALITO': txt(r['ANALITO']),
                                                          'MOTIVO': 'QA liberacao run'}))
        a0 = aud0()
        atualizar()
        lib1 = lib_ler(q)
        h2 = txt(q.ws('Liberação').Range('H2').Value)
        jx, ju = onde(lib1, rx), onde(lib1, rult)
        ac1 = acoes_desde(a0)
        n_orf = ac1.count('LIBERACAO_ASSINATURA_SEM_CORRIDA')
        ultimo_listado = max((i for i, l in enumerate(lib1) if iint(l[0])), default=None)
        ok1 = (jx is not None and txt(lib1[jx][3]) == MARCA + str(rx) and txt(lib1[jx][1]) == SEM
               and jx == ultimo_listado and ju is not None and txt(lib1[ju][3]) == MARCA + str(rult)
               and str(rx) in h2 and n_orf == 1)
        reg('L01 corrida assinada sem resultado ativo: a assinatura vai para o fim com "(sem resultado ativo)", aviso e '
            '1 trilha; a assinatura da ultima corrida listada continua ao lado do seu RUN',
            ok1, {'lote': lote, 'equip': eqf, 'corridas_antes': len(runs0), 'RUN_sem_ativo': rx,
                  'n_inativados': len(linhas_inat), 'linha_do_RUN_depois': None if jx is None else 4 + jx,
                  'conteudo': None if jx is None else lib1[jx][:4], 'RUN_ultimo': rult,
                  'linha_ultimo_depois': None if ju is None else 4 + ju,
                  'conteudo_ultimo': None if ju is None else lib1[ju][:4], 'H2': h2[:200], 'trilhas': n_orf})
        # L02
        a1 = aud0()
        atualizar()
        lib2 = lib_ler(q)
        reg('L02 um 2o ATUALIZAR nao repete a trilha nem muda a lista',
            lib2 == lib1 and acoes_desde(a1).count('LIBERACAO_ASSINATURA_SEM_CORRIDA') == 0,
            {'lista_igual': lib2 == lib1, 'trilhas_novas': acoes_desde(a1).count('LIBERACAO_ASSINATURA_SEM_CORRIDA')})
        # L03
        for rr in sorted(linhas_inat, reverse=True):
            apagar_linha(q, TB_INAT, rr)
        atualizar()
        lib3 = lib_ler(q)
        kx, ku = onde(lib3, rx), onde(lib3, rult)
        ok3 = (kx == ix and txt(lib3[kx][3]) == MARCA + str(rx) and ku == iu and txt(lib3[ku][3]) == MARCA + str(rult)
               and [iint(l[0]) for l in lib3 if iint(l[0])] == runs0)
        reg('L03 desfeitas as inativacoes, a corrida volta ao seu lugar com a assinatura e a lista volta a do inicio',
            ok3, {'linha_RUN': None if kx is None else 4 + kx, 'esperada': 4 + ix,
                  'lista_igual_ao_inicio': [iint(l[0]) for l in lib3 if iint(l[0])] == runs0})
    except Exception as e:                           # noqa: BLE001
        reg('Z99 Execucao da suite sem excecao', False, {'erro': repr(e), 'trace': traceback.format_exc()[-1500:]})
    finally:
        q.fechar(salvar=False)
    h1 = sha256_arquivo(caminho)
    reg('Y00 Seguranca do teste: arquivo de entrada intacto (SHA-256 igual antes e depois)', h0 == h1,
        {'antes': h0, 'depois': h1})
    if saida:
        with open(saida, 'w', encoding='utf-8') as f:
            json.dump(RES, f, ensure_ascii=False, indent=1, default=str)
    falhas = [r for r in RES if r['resultado'] == 'FAIL']
    print(f'\n=== {produto} Liberacao pelo RUN: {len(RES) - len(falhas)} PASS / {len(falhas)} FAIL ===', flush=True)
    return falhas


if __name__ == '__main__':
    falhas = executar(sys.argv[1], os.path.abspath(sys.argv[2]), sys.argv[3] if len(sys.argv) > 3 else None)
    sys.exit(1 if falhas else 0)
