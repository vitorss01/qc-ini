# -*- coding: utf-8 -*-
"""valida_adr089.py -- Tester/QA, ciclo 2026-11-04_reteste02, modo VALIDAR (ADR-089, commit 2facec9, entrega 36).

Escrito pelo Tester/QA (quem corrigiu foi o agente programador; nao reaproveita qa_adr089.py). Reaproveita
repro_comum (ciclo 01) e qa_final.QA. Uma sessao de Excel por produto, sobre COPIA da producao atual em
%TEMP%/qcw/auditoria/2026-11-04_reteste02/. Nada e salvo; fechar() encerra so o PID que o QA criou; sem captura de tela.

Roteiro (os passos dos grupos G5-G8 do tester_reproduzir.md do ciclo 01, refeitos sobre a correcao):
  L51  OC-0051 SEM UserInterfaceOnly (arquivo recem-aberto): lote -> K3/K4 -> H3 vazio = TODOS; ampliacao lote A -> B
       -> vazio; e o mesmo COM UserInterfaceOnly (sessao de login).
  C1-C5 OC-0027/0032 (sem UIO: as funcoes liberam a protecao sozinhas): registrar 110, duplicidade, Corrigir 105,
       Excluir; G a cada passo; trilha CEQ_ALTERADO/CEQ_EXCLUIDO/EQA_BASE_CONSOLIDADA.
  C6   OC-0028 (com UIO): registrar 110, editar F na tabela (Select + escrita, eventos ligados) -> Z3/Z4/AA1, G parado;
       ATUALIZAR DADOS COMPLETO (AtualizarDadosAutomatico, le o DB_SEAC so para leitura) -> consolida, G = G da correcao,
       Z2 = Z1.
  C7   OC-0029: aviso AZUL/AMBAR, ConfirmacoesCEQ, OK "fora do bias", Estatistica parada; CAP fora do mapa, Controllab,
       analito novo; ampliacao: edicao direta no Controllab + botao Consolidar.
  C8   OC-0030: todos os pares de prefixo do mapa (calculados aqui a partir do W:Y), unidade do par x unidade propria.
  C10  OC-0032 PROVEDOR_ALTERADO (Bio Painel M4; Hema Estatistica R4).
  FIM  integridade da cadeia do Audit_Log; producao intacta.
Dados de teste: rodada 'TESTE-AUDITORIA R2'/'TESTE-AUDITORIA CTL2', amostras TA-*, arquivo TESTE-AUDITORIA.pdf.
Uso: python -u valida_adr089.py <Bioquimica|Hematologia>
"""
import datetime as dt
import os
import sys
import time
import traceback
from collections import Counter

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import repro_comum as C  # noqa: E402

PROD = sys.argv[1]
BIO = PROD == 'Bioquimica'
NLV = 2 if BIO else 3
C.PASTA = os.path.join(os.environ['TEMP'], 'qcw', 'auditoria', '2026-11-04_reteste02')
ROD, ROD_CTL, ANO = 'TESTE-AUDITORIA R2', 'TESTE-AUDITORIA CTL2', '2025'
if BIO:
    AN, NOME, UN, CANON = 'Proteína total', 'Protein, total, serum', 'g/dL', 'Proteína total'
    FORA = [('CAP', 'Protein, total', 'g/dL'), ('Controllab', 'Proteína total', 'g/dL'), ('CAP', 'Hemoglobin A1c', '%')]
else:
    AN, NOME, UN, CANON = 'RBC', 'Red Blood Cell Count', 'x 10E12/L', 'RBC'
    FORA = [('CAP', 'Red Blood Cell', 'x 10E12/L'), ('Controllab', 'RBC', 'x 10E12/L'), ('CAP', 'nRBC', '/100 WBC')]
KEYS = [f'{AN}|{n}' for n in range(1, NLV + 1)]


def t(v):
    return '' if v is None else str(v)


def G(qa):
    e = C.estat_tabela(qa)
    return {k: e[k]['G_bias'] for k in KEYS}


def estat_snap(qa):
    e = C.estat_tabela(qa)
    return {k: (v['G_bias'], v['H_et'], v['L_sigma']) for k, v in e.items()}


def mudaram(a, b):
    return sorted(k for k in a if a[k] != b.get(k))


class Audit:
    def __init__(self, qa):
        self.qa = qa
        self.n0 = self.lo().ListRows.Count

    def lo(self):
        return self.qa.lo('tblAuditoria')

    def novos(self):
        lo = self.lo()
        n = lo.ListRows.Count
        if n <= self.n0:
            return []
        cab = [str(c.Name) for c in lo.ListColumns]
        v = lo.DataBodyRange.Rows(f'{self.n0 + 1}:{n}').Value
        self.n0 = n
        out = []
        for l in v:
            d = {cab[i]: C.limpo(x) for i, x in enumerate(l) if x not in (None, '')}
            txt = ' | '.join(str(x) for x in d.values())
            acao = next((str(x) for x in l if str(x) in ('CEQ_REGISTRADO', 'CEQ_ALTERADO', 'CEQ_EXCLUIDO',
                                                        'EQA_BASE_CONSOLIDADA', 'PROVEDOR_ALTERADO')), None)
            out.append({'acao': acao, 'txt': txt[:700]})
        return out


def painel_estado(qa, p, rot):
    eng = qa.ws('Eng_Saida')
    qa.ex.esperar()
    return {'passo': rot, 'H3': t(p.Range('H3').Value), 'E1': t(eng.Range('E1').Value), 'AD1': eng.Range('AD1').Value,
            'J7': [t(x[0]) for x in C.bloco(p, f'J7:J{6 + NLV}')], 'H4': t(p.Range('H4').Value)[:70],
            'G6': t(p.Range('G6').Value), 'A39': t(p.Range('A39').Value)[:90], 'cfgG2': t(qa.ws('Configuração').Range('G2').Value)}


def seq_l51(qa, p, seq, rot):
    out = []
    for c, v in seq:
        p.Range(c).Value = v
        qa.ex.esperar()
        time.sleep(0.5)
        if c == 'H3' or c == 'K4':
            out.append(painel_estado(qa, p, f'{rot}:{c}={v}'))
    time.sleep(2)
    p.Range('H3').MergeArea.ClearContents()
    qa.ex.esperar()
    time.sleep(2)
    out.append(painel_estado(qa, p, f'{rot}:H3 vazio'))
    ok = (not any('DESATUALIZADO' in ' '.join(d['J7']) for d in out) and out[-1]['AD1'] is True
          and 'TODOS' in out[-1]['H4'] and 'método' in out[-1]['G6']
          and all(d['E1'].lstrip("'") == d['H3'].lstrip("'") for d in out[:-1] if d['H3']))
    return {'passos': out, 'ok': ok}


def main():
    src, copia, h0 = C.copiar(PROD)
    L = C.Log(os.path.join(C.PASTA, f'valida_adr089_{PROD}.json'))
    L['ambiente'] = {'producao': os.path.basename(src), 'sha_prod': h0, 'copia': copia}
    senha = C.senha_planilha()
    qa = C.qa_final.QA(PROD, copia)
    L['excel_pid'] = qa.ex.pid
    try:
        p = qa.ws('Painel')
        est = qa.ws('Estatística')
        cap, ctl = qa.ws('EQA.CAP_Dados'), qa.ws('EQA.Controllab_Dados')
        base = qa.ws('EQA_Base')
        aud = Audit(qa)
        p.Activate()
        qa.ex.esperar()
        for w in (p, est):
            try:
                w.Unprotect(senha)
            except Exception:
                pass
        L['estado_inicial'] = {'Configuracao_ProtectContents': qa.ws('Configuração').ProtectContents,
                               'Configuracao_UIO': qa.ws('Configuração').ProtectionMode,
                               'CAP_Z2': t(cap.Range('Z2').Value)[:90], 'CAP_Z3': t(cap.Range('Z3').Value)[:90],
                               'EQA_Z1': t(base.Range('Z1').Value)[:90], 'pendente': qa.run('mEQA.ConsolidacaoPendente')}

        # ============================================================ L51 (OC-0051) sem UIO, depois com UIO
        if BIO:
            s1 = [('L4', 'DIMENSION 2'), ('C3', 'Hemoglobina glicada'), ('H3', "'8590"), ('K3', 2026), ('K4', 'Set')]
            s2 = [('C3', 'Glicose'), ('L4', 'DIMENSION 1'), ('H3', "'8974"), ('H3', "'8977"), ('K3', 2026), ('K4', 'Out')]
        else:
            s1 = [('C3', 'WBC'), ('H3', "'602911"), ('K3', 2026), ('K4', 'Mar')]
            s2 = [('C3', 'WBC'), ('H3', "'536411"), ('H3', "'602911"), ('K3', 2026), ('K4', 'Mar')]
        L['L51_semUIO_lote_K3K4_vazio'] = seq_l51(qa, p, s1, 'semUIO')
        L['L51_semUIO_loteA_loteB_vazio'] = seq_l51(qa, p, s2, 'semUIO_AB')
        L['L51_audit'] = aud.novos()

        # ============================================================ estado do CEQ
        if BIO:
            for c, v in (('C3', AN), ('L4', 'DIMENSION 2'), ('H3', "'8977"), ('K3', 2026), ('K4', 'Out')):
                C.por(qa, p, c, v, local=False)
        else:
            C.por(qa, p, 'C3', AN, local=False)
        mapa = [m for m in C.bloco(base, 'W2:Y400') if m[1]]
        L['mapa_n'] = len(mapa)

        # ============================================================ C1-C5 OC-0027 (sem UIO)
        oc27 = {}
        e0 = estat_snap(qa)
        g0 = G(qa)
        oc27['B0'] = g0
        c110 = C.campos_ceq('CAP', ROD, ANO, NOME, 'TA-11', '110', unid=UN)
        c105 = C.campos_ceq('CAP', ROD, ANO, NOME, 'TA-11', '105', unid=UN)
        oc27['aviso_nome_mapeado'] = t(qa.run('mCEQForm.AvisoAnalitoCEQ', 'CAP', NOME))
        oc27['confirmacoes_nome_mapeado'] = t(qa.run('mCEQForm.ConfirmacoesCEQ', c110))
        r1 = t(qa.run('mCEQForm.SalvarCEQ', c110))
        lin = int(r1.split('|')[1])
        g1 = G(qa)
        oc27['C1_registrar_110'] = {'ret': r1, 'G': g1, 'audit': aud.novos()}
        r2 = t(qa.run('mCEQForm.SalvarCEQ', c105))
        oc27['C2_duplicidade'] = {'ret': r2, 'cita_corrigir': 'Corrigir selecionado' in r2, 'audit': aud.novos()}
        reg = qa.run('mCEQForm.RegistrosDaRodada', 'CAP', ROD, ANO)
        oc27['lista_rodada'] = [list(x) for x in reg] if isinstance(reg, tuple) else reg
        oc27['descricao'] = t(qa.run('mCEQForm.DescricaoCEQ', 'CAP', lin))
        r3 = t(qa.run('mCEQForm.CorrigirCEQ', lin, c105))
        g3 = G(qa)
        lc = qa.run('mCEQForm.LinhaCEQ', 'CAP', lin)
        oc27['C3_corrigir_105'] = {'ret': r3, 'G': g3, 'linha_F': cap.Range(f'F{lin}').Value,
                                   'eqa_base': [(x.get('Resultado') or x) for x in C.eqa_teste(qa)[0]][:3],
                                   'LinhaCEQ': list(lc) if isinstance(lc, tuple) else lc, 'audit': aud.novos()}
        r3b = t(qa.run('mCEQForm.CorrigirCEQ', lin, C.campos_ceq('CAP', ROD, ANO, NOME, 'TA-11', 'abc', unid=UN)))
        oc27['C3b_corrigir_invalido'] = {'ret': r3b, 'F_inalterado': cap.Range(f'F{lin}').Value, 'audit': aud.novos()}
        r4 = t(qa.run('mCEQForm.ExcluirCEQ', 'CAP', lin))
        g4 = G(qa)
        e4 = estat_snap(qa)
        oc27['C4_excluir'] = {'ret': r4, 'G': g4, 'G_igual_B0_exato': g4 == g0, 'estat_linhas_dif_B0': mudaram(e0, e4),
                              'linha_vazia': [cap.Cells(lin, c).Value for c in (1, 4, 6, 18)],
                              'eqa_teste': len(C.eqa_teste(qa)[0]), 'audit': aud.novos()}
        r5 = t(qa.run('mCEQForm.ExcluirCEQ', 'CAP', lin))
        oc27['C5_excluir_de_novo'] = {'ret': r5, 'audit': aud.novos()}
        L['OC0027'] = oc27

        # ============================================================ com UIO (sessao de login)
        qa.run('mSeguranca.ReprotectAll')
        p.Activate()
        for w in (p, est):
            try:
                w.Unprotect(senha)
            except Exception:
                pass
        L['L51_comUIO_lote_K3K4_vazio'] = seq_l51(qa, p, s1, 'comUIO')
        if BIO:
            for c, v in (('C3', AN), ('L4', 'DIMENSION 2'), ('H3', "'8977"), ('K3', 2026), ('K4', 'Out')):
                C.por(qa, p, c, v, local=False)
        else:
            C.por(qa, p, 'C3', AN, local=False)
        aud.novos()

        # ============================================================ C6 OC-0028
        oc28 = {'G_antes': G(qa)}
        r6 = t(qa.run('mCEQForm.SalvarCEQ', c110))
        lin6 = int(r6.split('|')[1])
        oc28['registrar_110'] = {'ret': r6, 'reaproveitou_linha_excluida': lin6 == lin, 'G': G(qa), 'audit': aud.novos()}
        gx = G(qa)
        cap.Activate()
        qa.ex.esperar()
        cap.Range(f'F{lin6}').Select()
        qa.ex.esperar()
        cap.Range(f'F{lin6}').Value = 105
        qa.ex.esperar()
        time.sleep(0.5)
        oc28['edicao_direta'] = {'Z2': t(cap.Range('Z2').Value)[:90], 'Z3': t(cap.Range('Z3').Value),
                                 'Z3_ambar_negrito': (cap.Range('Z3').DisplayFormat.Interior.Color,
                                                      cap.Range('Z3').DisplayFormat.Font.Bold,
                                                      cap.Range('Z3').DisplayFormat.Font.Color),
                                 'Z4': cap.Range('Z4').Value, 'AA1': base.Range('AA1').Value,
                                 'pendente': qa.run('mEQA.ConsolidacaoPendente'), 'G': G(qa),
                                 'G_parado': G(qa) == gx, 'audit': aud.novos()}
        p.Activate()
        t0 = time.time()
        ra = t(qa.run('mIntegracao.AtualizarDadosAutomatico', teto=1800))
        oc28['atualizar_completo'] = {'ret': ra[:300], 's': round(time.time() - t0, 1)}
        g6 = G(qa)
        oc28['depois_atualizar'] = {'G': g6, 'G_igual_G_correcao_janela': g6 == g3,
                                    'Z2': t(cap.Range('Z2').Value)[:120], 'Z1': t(base.Range('Z1').Value)[:120],
                                    'Z2_igual_Z1': t(cap.Range('Z2').Value) == t(base.Range('Z1').Value),
                                    'Z3': t(cap.Range('Z3').Value)[:90], 'pendente': qa.run('mEQA.ConsolidacaoPendente'),
                                    'audit': aud.novos()}
        # sem pendencia: o ATUALIZAR nao consolida de novo (so a funcao, sem o motor)
        oc28['ConsolidarSePendente_sem_pendencia'] = {'ret': qa.run('mEQA.ConsolidarSePendente'), 'audit': aud.novos()}
        L['OC0028'] = oc28
        for w in (p, est):
            try:
                w.Unprotect(senha)
            except Exception:
                pass

        # ============================================================ C7 OC-0029
        oc29 = {}
        for i, (prov, nome, un) in enumerate(FORA):
            rod = ROD if prov == 'CAP' else ROD_CTL
            cp = C.campos_ceq(prov, rod, ANO, nome, f'TA-2{i}', '120', unid=un)
            a = estat_snap(qa)
            av = t(qa.run('mCEQForm.AvisoAnalitoCEQ', prov, nome))
            cf = t(qa.run('mCEQForm.ConfirmacoesCEQ', cp))
            r = t(qa.run('mCEQForm.SalvarCEQ', cp))
            b = estat_snap(qa)
            oc29[f'{prov}:{nome}'] = {'aviso': av, 'confirmacao': cf, 'ret': r, 'estat_mudou': mudaram(a, b),
                                      'MapearAnalito': t(qa.run('mEQA.MapearAnalito', prov, nome)), 'audit': aud.novos()}
        oc29['aviso_mapeado_CAP'] = t(qa.run('mCEQForm.AvisoAnalitoCEQ', 'CAP', NOME))
        oc29['aviso_vazio'] = t(qa.run('mCEQForm.AvisoAnalitoCEQ', 'CAP', ''))
        # ampliacao OC-0028 no Controllab: edicao direta + botao Consolidar
        k = next(x for x in oc29 if x.startswith('Controllab:'))
        lctl = int(oc29[k]['ret'].split('|')[1])
        ctl.Activate()
        qa.ex.esperar()
        ctl.Range(f'F{lctl}').Select()
        qa.ex.esperar()
        ctl.Range(f'F{lctl}').Value = 131
        qa.ex.esperar()
        amp = {'Z3_ctl': t(ctl.Range('Z3').Value), 'pendente': qa.run('mEQA.ConsolidacaoPendente'), 'audit': aud.novos()}
        qa.run('mEQA.AtualizarEQABase')
        amp.update({'Z3_ctl_depois': t(ctl.Range('Z3').Value)[:90], 'pendente_depois': qa.run('mEQA.ConsolidacaoPendente'),
                    'Z2_ctl_igual_Z1': t(ctl.Range('Z2').Value) == t(base.Range('Z1').Value), 'audit_botao': aud.novos()})
        oc29['amp_controllab_edicao_direta'] = amp
        L['OC0029'] = oc29

        # ============================================================ C8 OC-0030: pares do mapa
        cap_ = [m for m in mapa if str(m[0]).strip().upper() == 'CAP'] or mapa
        dm = {str(m[1]).strip(): str(m[2]).strip() for m in cap_}
        lo = cap.ListObjects(1)
        dados = lo.DataBodyRange.Value
        uni = {}
        for l in dados:
            nm, u = t(l[3]).strip(), t(l[13]).strip()
            if nm in dm and u and 'SIMULA' not in t(l[17]).upper():
                uni.setdefault(dm[nm], Counter())[u] += 1
        moda = {c: u.most_common(1)[0][0] for c, u in uni.items()}
        pares = []
        for a_ in dm:
            for b_ in dm:
                if a_ != b_ and b_.upper().startswith(a_.upper()) and dm[a_] != dm[b_]:
                    pares.append((a_, b_))
        oc30 = {'pares': pares, 'resultado': []}
        for a_, b_ in pares:
            ua, ub = moda.get(dm[a_]), moda.get(dm[b_])
            d = {'curto': a_, 'canon': dm[a_], 'longo': b_, 'canon_par': dm[b_], 'unid_propria': ua, 'unid_par': ub,
                 'aviso': t(qa.run('mCEQForm.AvisoAnalitoCEQ', 'CAP', a_))}
            d['aviso_cita_par'] = b_ in d['aviso'] and 'existe também' in d['aviso']
            if ub:
                d['conf_unid_par'] = t(qa.run('mCEQForm.ConfirmacoesCEQ', C.campos_ceq('CAP', ROD, ANO, a_, 'TA-30', '110', unid=ub)))
            if ua:
                d['conf_unid_propria'] = t(qa.run('mCEQForm.ConfirmacoesCEQ', C.campos_ceq('CAP', ROD, ANO, a_, 'TA-30', '110', unid=ua)))
            d['aviso_longo'] = t(qa.run('mCEQForm.AvisoAnalitoCEQ', 'CAP', b_))[:160]
            oc30['resultado'].append(d)
        if not BIO:
            cl = C.campos_ceq('CAP', ROD, ANO, 'Lymphs', 'TA-31', '110', unid='x 10E9/L')
            a = estat_snap(qa)
            oc30['Lymphs_x10E9L'] = {'aviso': t(qa.run('mCEQForm.AvisoAnalitoCEQ', 'CAP', 'Lymphs')),
                                     'confirmacao': t(qa.run('mCEQForm.ConfirmacoesCEQ', cl)),
                                     'ret': t(qa.run('mCEQForm.SalvarCEQ', cl))}
            b = estat_snap(qa)
            oc30['Lymphs_x10E9L']['estat_mudou'] = mudaram(a, b)
            oc30['Lymphs_x10E9L']['audit'] = aud.novos()
        L['OC0030'] = oc30

        # ============================================================ C10 OC-0032 PROVEDOR_ALTERADO
        prov = {}
        if BIO:
            p.Activate()
            antes = t(p.Range('M4').Value)
            ga = G(qa)
            C.por(qa, p, 'M4', 'Controllab', local=False)
            gb = G(qa)
            C.por(qa, p, 'M4', antes or 'CAP', local=False)
            prov = {'M4_antes': antes, 'G_CAP': ga, 'G_Controllab': gb, 'G_volta': G(qa) == ga, 'audit': aud.novos()}
        else:
            est.Activate()
            qa.ex.esperar()
            antes = t(est.Range('R4').Value)
            est.Range('R4').Select()
            est.Range('R4').Value = 'Controllab'
            qa.ex.esperar()
            est.Range('R4').Select()
            est.Range('R4').Value = antes or 'CAP'
            qa.ex.esperar()
            prov = {'R4_antes': antes, 'audit': aud.novos()}
        L['OC0032_provedor'] = prov
        L['FIM'] = {'VerificarIntegridadeLog': t(qa.run('mAuditoria.VerificarIntegridadeLog'))[:300]}
    except Exception:
        L['ERRO'] = traceback.format_exc()
    finally:
        qa.fechar()
        L['sha_prod_depois'] = C.sha(src)
        L['fim'] = dt.datetime.now().isoformat(timespec='seconds')


if __name__ == '__main__':
    main()
