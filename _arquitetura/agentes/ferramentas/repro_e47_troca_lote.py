# -*- coding: utf-8 -*-
"""repro_e47_troca_lote.py -- Tester/QA 24/11/2026, validacao da entrega 47 (A-01). Cenario PROPRIO, diferente do
repro_e46_liberacao: SEM login (automacao pura, abas protegidas sem UserInterfaceOnly), DUAS trocas seguidas para
lotes diferentes (L0 -> L1 -> L2 -> L0), depois Save da COPIA (caminho Workbook_BeforeSave: FlushLoteAtual,
TrancarParaGravar, RestaurarAposGravar) e nova troca na mesma sessao, depois fechar, REABRIR a copia salva e trocar
de novo. Marcadores sinteticos 'QA47' / 'TESTE-AUDITORIA ...' nas colunas C/D da Liberacao (so na copia).

Uso: python repro_e47_troca_lote.py <produto> <copia.xlsm> <saida.json>
  P01 L0 -> L1 -> L2 -> L0: marcadores do L0 voltam ao lado dos MESMOS RUNs; L2 nao mostra marcador de outro lote
  P02 de volta ao L1 (troca seguida): o marcador do L1 volta ao lado do mesmo RUN
  P03 protecao ativa (ProtectContents=True) em LiberStore, Liberacao e RegistrosStore depois de CADA troca
  P04 Save da copia (BeforeSave) e troca L0 -> L1 -> L0 na mesma sessao: marcadores intactos, protecao ativa
  P05 copia salva reaberta: L0 com os marcadores ao abrir; L0 -> L1 -> L0 intactos; protecao ativa
  P06 LiberStore (bloco do L0) tem os marcadores com o RUN na coluna E
"""
import json
import os
import sys
import traceback

TESTES = os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', '..', 'integracao_dbseac', 'testes')
sys.path.insert(0, os.path.abspath(TESTES))
from qa_final import QA, reg, RES  # noqa: E402
from qa_fluxo_manual import lib_gravar  # noqa: E402

ABA = 'Liberação'
ABAS_P = ('LiberStore', ABA, 'RegistrosStore')


def t(v):
    if isinstance(v, float) and v == int(v):
        return str(int(v))
    return '' if v is None else str(v).strip()


def main(produto, caminho, saida):
    prot_log = []
    q = None
    try:
        q = QA(produto, caminho)
        lib = q.ws(ABA)

        def lote():
            return t(q.run('mDados.LoteAtivoCore'))

        def ler():
            return [[t(x) for x in r] for r in lib.Range('A4:D203').Value]

        def marc():
            return {r[0]: (r[2], r[3]) for r in ler() if r[0] and r[2] == 'QA47'}

        def prot(rot):
            st = {}
            for n in ABAS_P:
                try:
                    w = q.ws(n)
                    st[n] = (bool(w.ProtectContents), bool(w.ProtectionMode))
                except Exception as e:      # noqa: BLE001
                    st[n] = repr(e)[:80]
            prot_log.append((rot, st))
            return all(isinstance(v, tuple) and v[0] for v in st.values())

        def usar(lt, rot):
            q.run('mApp.UsarLote', lt, teto=1800)
            ok = lote() == lt
            return ok and prot(rot)

        def marcar(rot, linhas):
            rr = ler()
            com_run = [i for i, r in enumerate(rr) if r[0]]
            feitos = {}
            for k in linhas:
                if k < len(com_run):
                    i = com_run[k]
                    lib_gravar(q, i, 'QA47', f'TESTE-AUDITORIA {rot} RUN {rr[i][0]}')
                    feitos[rr[i][0]] = ('QA47', f'TESTE-AUDITORIA {rot} RUN {rr[i][0]}')
            return feitos

        def salvar_copia():
            # REPRO_SAVE=manual: a MESMA sequencia do Workbook_BeforeSave (EstaPastaDeTrabalho.cls l.23-47) chamada
            # passo a passo, porque o wb.Save() pela COM com eventos ligados nao chegou ao disco (ver diag_save)
            h0 = os.path.getmtime(caminho)
            if os.environ.get('REPRO_SAVE') == 'manual':
                q.ex.xl.EnableEvents = False
                q.run('mLotes.FlushLoteAtual', teto=600)
                q.run('mSeguranca.TrancarParaGravar', teto=600)
                q.wb.Save()
                q.run('mSeguranca.RestaurarAposGravar', teto=600)
                q.ex.xl.EnableEvents = True
            else:
                q.ex.xl.EnableEvents = True
                q.wb.Save()
            prot_log.append(('gravou no disco', os.path.getmtime(caminho) != h0))

        cad = [t(c.Value) for c in q.wb.Names('regLoteCol').RefersToRange]
        L0 = lote()
        outros = [c for c in cad if c and c != L0]
        prot('abertura (sem login)')
        m0 = marcar('L0', [0, 2])
        # L1 e L2: os dois primeiros lotes com corridas na Liberacao
        L1 = L2 = None
        m1 = {}
        ok_trocas = True
        for c in outros:
            ok_trocas &= usar(c, f'-> {c}')
            if any(r[0] for r in ler()):
                if L1 is None:
                    L1 = c
                    m1 = marcar('L1', [1])
                else:
                    L2 = c
                    break
        if L1 is None or L2 is None:
            reg('P00 pre-condicao: 2 outros lotes com corridas', False, {'L0': L0, 'L1': L1, 'L2': L2})
            return
        # garante ordem L0 -> L1 -> L2 -> L0 a partir daqui (ja estamos no L2)
        m_l2 = marc()
        ok_trocas &= usar(L0, f'-> {L0} (volta)')
        v0 = marc()
        reg('P01 L0 -> L1 -> L2 -> L0 sem login: marcadores do L0 voltam ao lado dos MESMOS RUNs; L2 sem marcador alheio',
            v0 == m0 and len(m0) == 2 and not m_l2,
            {'produto': produto, 'L0': L0, 'L1': L1, 'L2': L2, 'esperado_L0': m0, 'obtido_L0': v0, 'marcadores_no_L2': m_l2,
             'lotes_percorridos': outros[:outros.index(L2) + 1]})
        ok_trocas &= usar(L1, f'-> {L1}')
        v1 = marc()
        ok_trocas &= usar(L0, f'-> {L0}')
        v0b = marc()
        reg('P02 troca seguida de volta ao L1: o marcador do L1 volta ao lado do mesmo RUN; L0 intacto de novo',
            v1 == m1 and len(m1) == 1 and v0b == m0, {'esperado_L1': m1, 'obtido_L1': v1, 'L0': v0b == m0})
        reg('P03 protecao (ProtectContents) ativa em LiberStore/Liberacao/RegistrosStore depois de cada troca (sessao 1)',
            ok_trocas, {'trocas': len(prot_log) - 1, 'estados': prot_log[-3:]})

        # ---------------- Save (BeforeSave) e nova troca na mesma sessao
        salvar_copia()
        p_save = prot('depois do Save')
        ok4 = usar(L1, f'pos-save -> {L1}')
        v1s = marc()
        ok4 &= usar(L0, f'pos-save -> {L0}')
        v0s = marc()
        reg('P04 Save da copia (BeforeSave) e L0 -> L1 -> L0 na mesma sessao: marcadores intactos e protecao ativa',
            p_save and ok4 and v1s == m1 and v0s == m0,
            {'protecao_depois_save': prot_log[-3][1], 'L1': v1s == m1, 'L0': v0s == m0, 'estados': prot_log[-2:]})
        salvar_copia()                 # grava com L0 em uso e os marcadores
        q.fechar(salvar=False)
        q = None

        # ---------------- reabre a copia salva
        q = QA(produto, caminho)
        lib = q.ws(ABA)
        pa = prot('reaberta (sem login)')
        la = lote()
        va = marc()
        ok5 = usar(L1, f'reaberta -> {L1}')
        v1r = marc()
        ok5 &= usar(L0, f'reaberta -> {L0}')
        v0r = marc()
        reg('P05 copia salva reaberta: L0 com os marcadores; L0 -> L1 -> L0 intactos; protecao ativa',
            pa and ok5 and la == L0 and va == m0 and v1r == m1 and v0r == m0,
            {'lote_ao_abrir': la, 'L0_ao_abrir': va == m0, 'L1': v1r == m1, 'L0': v0r == m0, 'estados': prot_log[-3:]})
        q.run('mLotes.FlushLoteAtual', teto=600)
        cad = [t(c.Value) for c in q.wb.Names('regLoteCol').RefersToRange]
        rb0 = 2 + cad.index(L0) * 200
        stl = q.ws('LiberStore')
        blk = [[t(x) for x in r] for r in stl.Range(stl.Cells(rb0, 1), stl.Cells(rb0 + 199, 5)).Value]
        st = {r[4]: (r[0], r[1]) for r in blk if r[0] == 'QA47'}
        reg('P06 LiberStore, bloco do L0: marcadores com o RUN na coluna E', st == m0,
            {'bloco_linha0': rb0, 'obtido': st})
    except Exception as e:             # noqa: BLE001
        reg('Z99 execucao sem excecao', False, {'erro': repr(e)[:300], 'trace': traceback.format_exc()[-1200:]})
    finally:
        if q is not None:
            q.fechar(salvar=False)
    with open(saida, 'w', encoding='utf-8') as f:
        json.dump({'res': RES, 'protecao': prot_log}, f, ensure_ascii=False, indent=1, default=str)
    falhas = [r for r in RES if r['resultado'] == 'FAIL']
    print('=== %s repro_e47_troca_lote: %d PASS / %d FAIL ===' % (produto, len(RES) - len(falhas), len(falhas)), flush=True)
    return falhas


if __name__ == '__main__':
    sys.exit(1 if main(sys.argv[1], os.path.abspath(sys.argv[2]), os.path.abspath(sys.argv[3])) else 0)
