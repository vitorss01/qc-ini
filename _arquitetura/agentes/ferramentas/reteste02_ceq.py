# -*- coding: utf-8 -*-
"""reteste02_ceq.py -- Usuario Simulado, reteste 02 (ADR-089): OC-0027/0028/0029/0030/0032 (CEQ) e OC-0051 (H3).
Independente dos scripts do Tester (nao usa valida_adr089.py nem repro_*).

Uso (um Excel por vez; cada etapa abre e fecha o SEU Excel pelo PID; nunca em paralelo com outro agente):
    set PYTHONIOENCODING=utf-8
    python -u reteste02_ceq.py <Bioquimica|Hematologia> [ceq,h3,calc]
Pasta: %TEMP%\\qcw\\auditoria\\2026-11-04_reteste02\\usuario\\<produto>\\ (copias novas da producao, shutil.copy2;
o sha256 curto da producao e conferido antes e depois). DB_SEAC: o congelado do reteste 01 (copia de arquivo).

  ceq   copia nova; login de um ANALISTA de teste (o caminho do usuario); ATUALIZAR inicial (o 1o depois da entrega);
        fluxo do ciclo 01 (f2_ceq.py / f2_ceq_hema.py), pelo nucleo que os botoes da janela chamam (mCEQForm):
        C0  desenho do frmCEQ (botoes Corrigir/Excluir, lista com 5 colunas) -- so leitura do Designer;
        C1  registro novo TESTE-AUDITORIA R1 (110/100) -> lista da rodada (5a coluna = linha) -> LinhaCEQ;
        C2  duplicidade (mensagem nova) -> CorrigirCEQ (105) como ANALISTA, mesma linha, CEQ_ALTERADO;
        C3  edicao DIRETA na tabela (120; aba liberada como o ADM faria) -> Z3/Z4/AA1 -> bias parado ->
            ATUALIZAR (ConsolidarSePendente) -> Z2/Z3 -> bias novo; edicao sem selecao ("?");
        C4  fora do mapa (CAP e Controllab): AvisoAnalitoCEQ, ConfirmacoesCEQ, OK "Gravado fora do bias";
        C5  pares (Hema "Lymphs" x 10E9/L; Bio "Lactate"): aviso azul com o par e confirmacao de unidade;
        C6  PROVEDOR_ALTERADO (Bio Painel M4; Hema Estatistica R4, selecionando antes);
        C7  ExcluirCEQ de tudo o que foi criado -> Estatistica volta bit a bit ao inicio; trilha CEQ_EXCLUIDO.
        Retratos SaveCopyAs (eventos desligados) S_corr, S_atu, S_fim para a calculadora.
  h3    OC-0051: copia nova SEM login (sem UserInterfaceOnly, como no reteste 01), ATUALIZAR, a sequencia do
        reteste01_painel_h3.py (lote em H3 -> K3/K4 -> apagar H3) e o fluxo vizinho (lote A -> lote B -> vazio).
  calc  calc_independente.py sobre os retratos (sem Excel): as linhas do analito (G, H, L, T, AC, AD).
Marcacao: rodada/amostra/arquivo "TESTE-AUDITORIA"; valores sinteticos 110/105/120/130/150 sobre alvo 100.
Senha: importada da suite de testes, nunca escrita aqui. Nenhum caminho de rede impresso.
Autor: AGENTE_USUARIO_SIMULADO, 2026-11-04_reteste02.
"""
import datetime as dt
import hashlib
import json
import os
import shutil
import subprocess
import sys
import time
import traceback

AQUI = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.abspath(os.path.join(AQUI, '..', '..', '..'))
TESTES = os.path.join(REPO, '_arquitetura', 'integracao_dbseac', 'testes')
sys.path.insert(0, TESTES)
sys.path.insert(0, os.path.join(TESTES, '..'))
from qa_final import QA  # noqa: E402
from qa_etl_recebimento import SENHA  # noqa: E402
from qa_inativar_ux import criar_usuario, fazer_login  # noqa: E402

PROD = sys.argv[1]
ETAPAS = (sys.argv[2] if len(sys.argv) > 2 else 'ceq,h3,calc').split(',')
BIO = PROD.startswith('Bio')
W = os.path.join(os.environ['TEMP'], 'qcw', 'auditoria', '2026-11-04_reteste02', 'usuario', PROD)
CONG = os.path.join(os.environ['TEMP'], 'qcw', 'auditoria', '2026-11-04_reteste01', 'usuario', PROD,
                    'DB_SEAC_congelado.xlsm')
ORIGEM = os.path.join(REPO, f'QC_{PROD}.xlsm')
MARCA = 'TESTE-AUDITORIA'
if BIO:
    AN, NOME, UNID, CTRL, FORA, PAR, PAR_UN = 'Glicose', 'Glucose, serum', 'mg/dL', 'Ureia', 'Glucose', 'Lactate', 'U/L'
else:
    AN, NOME, UNID, CTRL, FORA, PAR, PAR_UN = 'LYMPH#', 'Lymphs Absolute', 'x 10E9/L', 'LYMPH%', 'Lymphs Abs', \
        'Lymphs', 'x 10E9/L'
EST_RG = 'A14:AU97'
COLS = ('G', 'H', 'L', 'T', 'AC', 'AD')


def log(*a):
    print(time.strftime('%H:%M:%S'), *a, flush=True)


def sha(c):
    return hashlib.sha256(open(c, 'rb').read()).hexdigest()[:12]


def limpo(v):
    if isinstance(v, dt.datetime):
        return v.replace(tzinfo=None).isoformat()
    if isinstance(v, (list, tuple)):
        return [limpo(x) for x in v]
    if isinstance(v, dict):
        return {str(k): limpo(x) for k, x in v.items()}
    return v


def ci(letra):
    n = 0
    for ch in letra:
        n = n * 26 + ord(ch) - 64
    return n - 1


def copia(nome):
    os.makedirs(W, exist_ok=True)
    cp = os.path.join(W, nome)
    shutil.copy2(ORIGEM, cp)
    return cp


class S:
    def __init__(self, q):
        self.q, self.xl, self.wb = q, q.ex.xl, q.wb
        self.est = q.ws('Estatística')

    def run(self, m, *a):
        return self.xl.Run(f"'{self.wb.Name}'!mCEQForm.{m}", *a)

    def linhas(self, *ans):
        v = self.est.Range(EST_RG).Value
        out = {}
        for i, r in enumerate(v):
            a = str(r[0] or '').strip()
            if a in ans:
                out[f'{a}|{r[1]}'] = {c: r[ci(c)] for c in COLS}
        return out

    def estat_tudo(self):
        return [list(r) for r in self.est.Range(EST_RG).Value]

    def lo(self, prov):
        nome = 'tblEQA_CAP_Dados' if prov == 'CAP' else 'tblEQA_Controllab_Dados'
        for ws in self.wb.Worksheets:
            for t in ws.ListObjects:
                if t.Name == nome:
                    return t

    def zz(self, prov='CAP'):
        ws = self.lo(prov).Parent
        b = self.q.ws('EQA_Base')
        return {'Z1': ws.Range('Z1').Value, 'Z2': ws.Range('Z2').Value, 'Z3': ws.Range('Z3').Value,
                'Z4': ws.Range('Z4').Value, 'AA1': b.Range('AA1').Value, 'Base_Z1': b.Range('Z1').Value,
                'Z3_negrito': bool(ws.Range('Z3').Font.Bold), 'Z3_cor': ws.Range('Z3').DisplayFormat.Font.Color}

    def base_teste(self):
        b = self.q.ws('EQA_Base')
        n = int(self.xl.WorksheetFunction.CountA(b.Range('A2:A5001')))
        rows = [list(r) for r in b.Range(f'A2:U{1 + max(n, 1)}').Value] if n else []
        return n, [r for r in rows if MARCA in str(r[2] or '') or MARCA in str(r[5] or '')]

    def aud_ult(self):
        al = self.q.ws('Audit_Log')
        return al.Cells(al.Rows.Count, 1).End(-4162).Row

    def aud_desde(self, r0):
        al = self.q.ws('Audit_Log')
        ult = self.aud_ult()
        if ult <= r0:
            return []
        cab = al.Range(al.Cells(3, 1), al.Cells(3, 20)).Value[0]   # linha 3 = cabecalho (1-2 = titulo)
        return [{str(cab[i]): r[i] for i in range(20) if r[i] not in (None, '') and cab[i]}
                for r in al.Range(al.Cells(r0 + 1, 1), al.Cells(ult, 20)).Value]


def campos(prov='CAP', an=None, res='110', un=None, rod=None, am='TA-01', ano='2025', alvo='100'):
    return [prov, rod or f'{MARCA} R1', ano, an or NOME, am, res, alvo, '2', '5', '90', '110', 'Acceptable', '50',
            un or UNID, '1', f'{MARCA}.pdf']


def save_copy(q, nome):
    ev = q.ex.xl.EnableEvents
    q.ex.xl.EnableEvents = False
    try:
        q.wb.SaveCopyAs(os.path.join(W, nome))
    finally:
        q.ex.xl.EnableEvents = ev


def etapa_ceq():
    out = {'producao_sha12_antes': sha(ORIGEM), 'erros': []}
    cp = copia(f'{PROD}_ceq.xlsm')
    q = QA(PROD, cp)
    out['excel_pid'] = q.ex.pid
    s = S(q)
    criados = []   # (prov, linha)

    def gravar():
        json.dump(limpo(out), open(os.path.join(W, 'ceq.json'), 'w', encoding='utf-8'), ensure_ascii=False,
                  indent=1, default=str)

    def passo(nome, fn):
        log('==>', nome)
        try:
            out[nome] = fn()
        except Exception as e:  # noqa: BLE001
            out[nome] = {'ERRO': repr(e), 'trace': traceback.format_exc()[-1500:]}
            out['erros'].append(nome)
            log('ERRO', nome, repr(e))
        gravar()

    try:
        q.cfg('CAMINHO_DB_SEAC', CONG)
        criar_usuario(q, 'TA_R02_ANALISTA', 'ANALISTA')
        out['login'] = fazer_login(q, 'TA_R02_ANALISTA')
        out['papel'] = q.nome('currentPapel').Value
        q.ex.xl.EnableEvents = True
        t0 = time.time()
        r, _ = q.atualizar(teto=1800)
        out['atualizar_inicial'] = {'s': round(time.time() - t0, 1), 'r': str(r)[:300]}
        ano = s.est.Range('N4').Value
        ano = str(int(ano)) if isinstance(ano, float) else '2025'
        out['ano_ep'] = ano
        out['R4_ou_AR'] = s.est.Range('R4').Value
        B0 = s.estat_tudo()
        out['B0'] = s.linhas(AN, CTRL)
        out['B0_zz'] = s.zz()
        gravar()

        if os.environ.get('R02_SOBRESCREVER_P'):   # achado da 1a passada: |Bias| (P, formula) sobrescrito
            out['nota'] = 'R02_SOBRESCREVER_P ligado'

        def c0():
            comp = q.wb.VBProject.VBComponents('frmCEQ')
            d = {}
            for c in comp.Designer.Controls:
                cap = ''
                try:
                    cap = c.Caption
                except Exception:  # noqa: BLE001
                    pass
                if cap and any(k in cap for k in ('Corrigir', 'Excluir', 'Salvar', 'Cancelar', 'Fechar')) or \
                        c.Name.lower().startswith(('lst', 'lbl', 'cmd', 'btn')):
                    x = {'cap': cap, 'top': c.Top, 'left': c.Left, 'w': c.Width, 'enabled': c.Enabled}
                    try:
                        x['colunas'] = [c.ColumnCount, c.ColumnWidths]
                    except Exception:  # noqa: BLE001
                        pass
                    d[c.Name] = x
            return d
        passo('C0_desenho_frmCEQ', c0)

        def c1():
            a0 = s.aud_ult()
            r = str(s.run('SalvarCEQ', campos(ano=ano)))
            lin = int(r.split('|')[1]) if r.startswith('OK|') else 0
            if lin:
                criados.append(('CAP', lin))
            reg = s.run('RegistrosDaRodada', 'CAP', f'{MARCA} R1', ano)
            return {'ret': r, 'linha': lin, 'lista_rodada': limpo(reg), 'LinhaCEQ': limpo(s.run('LinhaCEQ', 'CAP', lin)),
                    'descricao': s.run('DescricaoCEQ', 'CAP', lin), 'estat': s.linhas(AN, CTRL), 'zz': s.zz(),
                    'trilha': s.aud_desde(a0)}
        passo('C1_registro', c1)
        lin = out['C1_registro'].get('linha') if isinstance(out.get('C1_registro'), dict) else 0

        def c2():
            a0 = s.aud_ult()
            dup = str(s.run('SalvarCEQ', campos(ano=ano)))
            corr = str(s.run('CorrigirCEQ', lin, campos(ano=ano, res='105')))
            # corrigir para a chave de OUTRO registro real tem de ser recusado (pega a 1a linha real da tabela)
            lo = s.lo('CAP')
            r0 = lo.DataBodyRange.Row
            real = [str(x or '') for x in lo.Parent.Range(lo.Parent.Cells(r0, lo.Range.Column),
                                                           lo.Parent.Cells(r0, lo.Range.Column + 17)).Value[0]]
            real_c = list(s.run('LinhaCEQ', 'CAP', r0) or [])
            corr_dup = str(s.run('CorrigirCEQ', lin, real_c)) if real_c else 'sem linha real'
            r = {'duplicidade': dup, 'corrigir_105': corr, 'corrigir_para_chave_real': corr_dup,
                 'linha_real_usada': r0, 'LinhaCEQ_depois': limpo(s.run('LinhaCEQ', 'CAP', lin)),
                 'base_teste': s.base_teste(), 'estat': s.linhas(AN, CTRL), 'zz': s.zz(), 'trilha': s.aud_desde(a0)}
            save_copy(q, 'S_corr.xlsm')
            return r
        passo('C2_duplicidade_e_correcao', c2)

        def c3():
            lo = s.lo('CAP')
            ws = lo.Parent
            cres = lo.Range.Column + 6 - 1
            ws.Unprotect(SENHA)          # o ANALISTA nao tem a aba liberada; e o ADM quem edita direto (ciclo 01)
            q.ex.xl.EnableEvents = True
            ws.Activate()
            a0 = s.aud_ult()
            ws.Cells(lin, cres).Select()
            ws.Cells(lin, cres).Value = 120
            q.ex.esperar()
            q.ex.xl.Calculate()
            r = {'apos_editar': {'zz': s.zz(), 'estat': s.linhas(AN), 'base_teste': s.base_teste()[1],
                                 'trilha': s.aud_desde(a0)}}
            # edicao sem selecionar antes (o "antes" vira "?"): a amostra, celula vizinha
            a1 = s.aud_ult()
            q.ex.xl.Goto(ws.Range('Z10'))
            ws.Cells(lin, lo.Range.Column + 5 - 1).Value = 'TA-01'   # mesmo valor (nao deve registrar)
            ws.Cells(lin, lo.Range.Column + 17 - 1).Value = '2'      # Pagina (Q) 1 -> 2 sem selecionar
            q.ex.esperar()
            r['sem_selecao'] = s.aud_desde(a1)
            a2 = s.aud_ult()
            q.ws('EQA_Base').Activate()
            t0 = time.time()
            ra, _ = q.atualizar(teto=1800)
            r['atualizar'] = {'s': round(time.time() - t0, 1), 'r': str(ra)[:300]}
            q.ex.xl.Calculate()
            r['apos_atualizar'] = {'zz': s.zz(), 'estat': s.linhas(AN, CTRL), 'base_teste': s.base_teste()[1],
                                   'trilha': [x for x in s.aud_desde(a2) if 'CEQ' in str(x) or 'EQA' in str(x)]}
            save_copy(q, 'S_atu.xlsm')
            # 2o ATUALIZAR sem pendencia: nao consolida de novo
            a3 = s.aud_ult()
            t0 = time.time()
            q.atualizar(teto=1800)
            r['atualizar_sem_pendencia'] = {'s': round(time.time() - t0, 1),
                                            'consolidou': any('EQA_BASE_CONSOLIDADA' in str(x) for x in s.aud_desde(a3)),
                                            'zz': s.zz()}
            return r
        passo('C3_edicao_direta_atualizar', c3)

        def c4():
            r = {}
            for prov, nome in (('CAP', FORA), ('Controllab', NOME), ('CAP', NOME)):
                r[f'aviso|{prov}|{nome}'] = s.run('AvisoAnalitoCEQ', prov, nome)
                r[f'confirma|{prov}|{nome}'] = s.run('ConfirmacoesCEQ', campos(prov, an=nome, ano=ano))
            antes = s.linhas(AN)
            a0 = s.aud_ult()
            for prov, nome, am in (('CAP', FORA, 'TA-02'), ('Controllab', NOME, 'TA-03')):
                g = str(s.run('SalvarCEQ', campos(prov, an=nome, res='150', am=am, ano=ano)))
                r[f'gravar|{prov}|{nome}'] = g
                if g.startswith('OK|'):
                    criados.append((prov, int(g.split('|')[1])))
            q.ex.xl.Calculate()
            r['estat_igual'] = s.linhas(AN) == antes
            r['estat'] = s.linhas(AN)
            r['trilha'] = s.aud_desde(a0)
            return r
        passo('C4_fora_do_mapa', c4)

        def c5():
            r = {'aviso_par': s.run('AvisoAnalitoCEQ', 'CAP', PAR),
                 'confirma_par_unid': s.run('ConfirmacoesCEQ', campos(an=PAR, un=PAR_UN, ano=ano)),
                 'aviso_nome_certo': s.run('AvisoAnalitoCEQ', 'CAP', NOME),
                 'confirma_nome_certo': s.run('ConfirmacoesCEQ', campos(ano=ano))}
            if not BIO:   # o fluxo original do ciclo 01: o usuario responde Sim e grava "Lymphs" 150/100
                ctrl0 = s.linhas(CTRL)
                g = str(s.run('SalvarCEQ', campos(an=PAR, un=PAR_UN, res='150', am='TA-04', ano=ano)))
                r['gravar_Lymphs_apos_Sim'] = g
                if g.startswith('OK|'):
                    criados.append(('CAP', int(g.split('|')[1])))
                q.ex.xl.Calculate()
                r['LYMPH%_antes'] = ctrl0
                r['LYMPH%_depois'] = s.linhas(CTRL)
                save_copy(q, 'S_par.xlsm')
            return r
        passo('C5_pares', c5)

        def c6():
            a0 = s.aud_ult()
            r = {}
            if BIO:
                p = q.ws('Painel')
                p.Activate()
                p.Range('C3').Value = AN
                q.ex.esperar()
                m4 = p.Range('M4').Value
                r['M4_antes'] = m4
                for v in ('Controllab', m4):
                    try:
                        p.Range('M4').Value = v
                    except Exception as e:  # noqa: BLE001
                        r['M4_sem_unprotect'] = repr(e)
                        p.Unprotect(SENHA)
                        p.Range('M4').Value = v
                    q.ex.esperar()
                r['M4_depois'] = p.Range('M4').Value
            else:
                e = s.est
                e.Activate()
                r4 = e.Range('R4').Value
                r['R4_antes'] = r4
                for v in ('Controllab', r4):
                    e.Range('R4').Select()
                    try:
                        e.Range('R4').Value = v
                    except Exception as ex:  # noqa: BLE001
                        r['R4_sem_unprotect'] = repr(ex)
                        e.Unprotect(SENHA)
                        e.Range('R4').Value = v
                    q.ex.esperar()
                r['R4_depois'] = e.Range('R4').Value
            r['trilha'] = s.aud_desde(a0)
            return r
        passo('C6_provedor', c6)

        def c7():
            a0 = s.aud_ult()
            r = {'excluir': {}}
            for prov, ln in reversed(criados):
                r['excluir'][f'{prov}|{ln}'] = [s.run('DescricaoCEQ', prov, ln), str(s.run('ExcluirCEQ', prov, ln))]
            r['excluir_de_novo'] = str(s.run('ExcluirCEQ', 'CAP', lin)) if lin else None
            q.ex.xl.Calculate()
            B9 = s.estat_tudo()
            difs = [(14 + i, j, B0[i][j], B9[i][j]) for i in range(len(B0)) for j in range(len(B0[i]))
                    if B0[i][j] != B9[i][j]]
            r['estat_bit_a_bit_igual_B0'] = not difs
            r['difs'] = difs[:20]
            r['estat'] = s.linhas(AN, CTRL)
            r['base_teste'] = s.base_teste()
            r['zz'] = s.zz()
            r['trilha'] = s.aud_desde(a0)
            # a linha vazia e reaproveitada pelo proximo registro?
            g = str(s.run('SalvarCEQ', campos(ano=ano, rod=f'{MARCA} R9')))
            r['reuso'] = g
            if g.startswith('OK|'):
                r['reuso_excluir'] = str(s.run('ExcluirCEQ', 'CAP', int(g.split('|')[1])))
            save_copy(q, 'S_fim.xlsm')
            return r
        passo('C7_exclusao', c7)
    finally:
        q.fechar(salvar=False)
        out['producao_sha12_depois'] = sha(ORIGEM)
        gravar()


def etapa_h3():
    out = {'producao_sha12_antes': sha(ORIGEM)}
    cp = copia(f'{PROD}_h3.xlsm')
    q = QA(PROD, cp)
    out['excel_pid'] = q.ex.pid
    try:
        q.cfg('CAMINHO_DB_SEAC', CONG)
        q.atualizar(teto=1800)
        p = q.ws('Painel')
        p.Activate()
        q.ex.esperar()
        try:
            p.Unprotect(SENHA)
        except Exception:  # noqa: BLE001
            pass
        q.ex.xl.EnableEvents = True
        nl = 2 if BIO else 3

        def tela(rot):
            q.ex.esperar()
            d = {c: p.Range(c).Value for c in ('C3', 'H3', 'K3', 'K4', 'G6', 'H6', 'I6', 'H4', 'L4')}
            d['linhas'] = [list(x)[:10] for x in p.Range('A7:J%d' % (6 + nl)).Value]
            d['A39'] = p.Range('A39').Value
            d['Config_G2'] = q.ws('Configuração').Range('G2').Value
            out[rot] = d

        seq = ([('L4', 'DIMENSION 2'), ('C3', 'Hemoglobina glicada'), ('H3', "'8590"), ('K3', 2026), ('K4', 'Set')]
               if BIO else [('C3', 'WBC'), ('H3', "'602911"), ('K3', 2026), ('K4', 'Mar')])
        for c, v in seq:
            p.Range(c).Value = v
            q.ex.esperar()
        tela('lote')
        time.sleep(3)
        p.Range('H3').MergeArea.ClearContents()
        q.ex.esperar()
        time.sleep(3)
        tela('h3_vazio')
        # vizinho: lote A -> lote B -> vazio (o fluxo do f4) e voltar a um lote
        viz = (("'8974", "'8977") if BIO else ("'536411", "'602911"))
        if BIO:
            p.Range('C3').Value = 'Glicose'
            q.ex.esperar()
        for i, v in enumerate(viz):
            p.Range('H3').Value = v
            q.ex.esperar()
            tela(f'viz_lote{i + 1}')
        p.Range('H3').MergeArea.ClearContents()
        q.ex.esperar()
        time.sleep(2)
        tela('viz_vazio')
        p.Range('H3').Value = viz[0]
        q.ex.esperar()
        tela('viz_volta_lote')
    finally:
        q.fechar(salvar=False)
        out['producao_sha12_depois'] = sha(ORIGEM)
        json.dump(limpo(out), open(os.path.join(W, 'h3.json'), 'w', encoding='utf-8'), ensure_ascii=False, indent=1,
                  default=str)
    for k, v in out.items():
        print(k, json.dumps(limpo(v), ensure_ascii=False, default=str)[:700])


def etapa_calc():
    res = {}
    for nome in ('S_corr', 'S_atu', 'S_par', 'S_fim'):
        x = os.path.join(W, nome + '.xlsm')
        if not os.path.exists(x):
            continue
        js = os.path.join(W, nome + '_calc.json')
        t0 = time.time()
        p = subprocess.run([sys.executable, '-u', os.path.join(AQUI, 'calc_independente.py'), x, '--json', js],
                           capture_output=True, text=True, encoding='utf-8', errors='replace')
        d = json.load(open(js, encoding='utf-8'))
        sel = {}
        for L in d['linhas']:
            if L['analito'] in (AN, CTRL):
                sel[f"{L['analito']}|{L['nivel']}"] = {c: L['comp'].get(c) for c in COLS if c in L['comp']}
        res[nome] = {'s': round(time.time() - t0, 1), 'resumo': [ln for ln in p.stdout.splitlines()
                                                                   if ln.startswith(('RESUMO', 'PAINEL', 'DIF'))][:30],
                     'linhas': sel}
        log(nome, res[nome]['resumo'][:3])
    json.dump(res, open(os.path.join(W, 'calc.json'), 'w', encoding='utf-8'), ensure_ascii=False, indent=1,
              default=str)


if __name__ == '__main__':
    for e in ETAPAS:
        {'ceq': etapa_ceq, 'h3': etapa_h3, 'calc': etapa_calc}[e]()
