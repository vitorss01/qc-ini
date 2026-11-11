# -*- coding: utf-8 -*-
"""valida_adr090_091.py -- Tester/QA, ciclo 2026-11-11_reteste03, modo VALIDAR.

ADR-090 (commits bbcd5a4/105ecbc, entrega 37): OC-0012, 0019, 0021, 0023, 0024, 0031, 0033, 0034, 0035, 0036, 0040,
0043, 0044, 0049, 0050, 0052, 0053, 0054, 0055. ADR-091 (commit bfcc5db, entrega 38): tblLotesMaterial.

Escrito pelo Tester/QA (quem corrigiu foi o agente programador). NAO reaproveita qa_adr090.py nem qa_adr091.py: os
cenarios sao os ORIGINAIS dos pareceres (parecer_ux G01..G32, parecer_cq G11, reteste02) com oraculos proprios
(contagens independentes na tblCQ_Final, largura medida do texto, leitura da caixa de mensagem da propria instancia).
Infraestrutura reaproveitada: qa_final.QA (abre a copia e revela as abas, imprime EXCEL_PID), repro_comum (meus
utilitarios do ciclo 01), auditoria_abas.largura_pt (medida da fonte), qa_inativar_ux.criar_usuario/fazer_login
(usuario de teste so na copia).

Regras: so COPIA da producao em %TEMP%/qcw/auditoria/2026-11-11_reteste03/; nada e salvo; fechar() encerra so o PID que
o QA criou; nenhuma captura de tela (a caixa de mensagem e lida pelo texto das janelas do PID desta sessao).
Dados de teste: manuais com MOTIVO/USUARIO "TESTE-AUDITORIA", valores sinteticos 123,456 / 111,111 / -1; CEQ com
rodada "TESTE-AUDITORIA C-A 2025", amostras TA-*, arquivo TESTE-AUDITORIA.pdf.

Uso: python -u valida_adr090_091.py <Bioquimica|Hematologia>
"""
import collections
import datetime as dt
import json
import os
import re
import sys
import threading
import time
import traceback
import unicodedata

AQUI = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, AQUI)
import repro_comum as C  # noqa: E402
sys.path.insert(0, C.TESTES)
import auditoria_abas  # noqa: E402
from qa_inativar_ux import criar_usuario, fazer_login  # noqa: E402

PROD = sys.argv[1]
BIO = PROD == 'Bioquimica'
NLV = 2 if BIO else 3
C.PASTA = os.path.join(os.environ['TEMP'], 'qcw', 'auditoria', '2026-11-11_reteste03')
GUIA = os.path.join(C.REPO, '_arquitetura', 'integracao_dbseac', 'GUIA_DE_USO.md')
USR = 'TESTE_AUDITORIA_R3'
TAG = 'TESTE-AUDITORIA'
MOTIVO = 'Outro: TESTE-AUDITORIA'
AMBAR = 255 + 235 * 256 + 156 * 65536
AMBAR_FONTE = 156 + 87 * 256
ROD, ANO = 'TESTE-AUDITORIA C-A 2025', '2025'
if BIO:
    AN_CEQ, NOME_CEQ, UN_CEQ = 'Proteína total', 'Protein, total, serum', 'g/dL'
else:
    AN_CEQ, NOME_CEQ, UN_CEQ = 'RBC', 'Red Blood Cell Count', 'x 10E12/L'
LOTE_ST = ('LOTE_FORA_DO_CQ', 'LOTE_OUTRO_MATERIAL')
PRECEDE_LOTE = ('INATIVADO', 'SEM_VALOR', 'SEM_DATA_HORA', 'SUBSTITUIDO_MANUAL', 'CONFLITO_MANUAL', 'DUPLICIDADE_ORIGEM',
                'MANUAL_INCOMPLETO', 'DUPLICIDADE')


def t(v):
    if isinstance(v, float) and v == int(v):
        return str(int(v))
    return '' if v is None else str(v).strip()


def up(v):
    s = unicodedata.normalize('NFC', t(v)).upper()
    return s


def num(v):
    return isinstance(v, (int, float)) and not isinstance(v, bool)


def milhar(n):
    return f'{n:,}'.replace(',', '.')


def fmt2(x):
    s = f'{x:,.2f}'
    return s.replace(',', '#').replace('.', ',').replace('#', '.')


class Audit:
    """Linhas novas da tblAuditoria desde a ultima chamada (texto inteiro + acao reconhecida)."""
    RX = re.compile(r'(RESULTADO_MANUAL_[A-Z_]+|CEQ_[A-Z_]+|EQA_BASE_CONSOLIDADA|PROVEDOR_[A-Z_]+)')

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
        v = lo.DataBodyRange.Rows(f'{self.n0 + 1}:{n}').Value
        self.n0 = n
        out = []
        for l in v:
            txt = ' | '.join(t(C.limpo(x)) for x in l if x not in (None, ''))
            m = self.RX.search(txt)
            out.append({'acao': m.group(1) if m else None, 'txt': txt[:1500]})
        return out


class Vigia(threading.Thread):
    """Le e fecha caixas de dialogo (#32770) e UserForms (ThunderDFrame) do PID desta sessao. Nao captura imagem."""

    def __init__(self, pid, teto=45):
        super().__init__(daemon=True)
        self.pid, self.teto = pid, teto
        self.vistos = []
        self.parar = threading.Event()

    def run(self):
        import win32con
        import win32gui
        import win32process
        t0 = time.time()
        while not self.parar.is_set() and time.time() - t0 < self.teto:
            achou = []

            def cb(h, _):
                try:
                    _, p = win32process.GetWindowThreadProcessId(h)
                    if p == self.pid and win32gui.GetClassName(h) in ('#32770', 'ThunderDFrame'):
                        achou.append(h)
                except Exception:
                    pass
                return True
            try:
                win32gui.EnumWindows(cb, None)
            except Exception:
                pass
            for h in achou:
                filhos = []

                def cc(ch, _):
                    try:
                        tx = win32gui.GetWindowText(ch)
                        if tx:
                            filhos.append(tx)
                    except Exception:
                        pass
                    return True
                try:
                    win32gui.EnumChildWindows(h, cc, None)
                except Exception:
                    pass
                self.vistos.append({'classe': win32gui.GetClassName(h), 'titulo': win32gui.GetWindowText(h),
                                    'textos': filhos})
                try:
                    win32gui.PostMessage(h, win32con.WM_CLOSE, 0, 0)
                except Exception:
                    pass
            if achou:
                time.sleep(1.0)
            time.sleep(0.25)


def manual_cols(qa):
    lo = qa.lo('tblResultados_Manuais')
    return lo, [str(c.Name) for c in lo.ListColumns]


def ler_manual(qa, linha):
    lo, cab = manual_cols(qa)
    out = {}
    for c in cab:
        cel = lo.ListColumns(c).DataBodyRange.Cells(linha, 1)
        out[c] = (t(C.limpo(cel.Value)), t(cel.Text))
    return out


def escrever_celula(qa, linha, col, v):
    lo, _ = manual_cols(qa)
    cel = lo.ListColumns(col).DataBodyRange.Cells(linha, 1)
    if isinstance(v, str):
        cel.FormulaLocal = v
    else:
        cel.Value = v
    qa.ex.esperar()


def main():
    src, copia, h0 = C.copiar(PROD)
    copia2 = os.path.join(C.PASTA, f'{PROD}_valida.xlsm')
    if os.path.exists(copia2):
        os.remove(copia2)
    os.replace(copia, copia2)
    copia = copia2
    suf = ('_so091' if '--so091' in sys.argv else '') + ('_' + sys.argv[sys.argv.index('--secoes') + 1].replace(',', '_')
                                                        if '--secoes' in sys.argv else '')
    L = C.Log(os.path.join(C.PASTA, f'valida_{PROD}{suf}.json'))
    L['ambiente'] = {'producao': os.path.basename(src), 'sha_prod_12': h0, 'copia': copia}
    senha = C.senha_planilha()
    qa = C.qa_final.QA(PROD, copia)
    L['excel_pid'] = qa.ex.pid
    xl = qa.ex.xl
    try:
        if '--so091' in sys.argv:
            L['linhas_tblCQ_Final'] = len(fin0 := qa.ler('tblCQ_Final'))
            secao(L, 'ADR091', s_adr091, qa, fin0)
        else:
            rodar(qa, L, senha, xl)
    except Exception as e:  # noqa: BLE001
        L['ERRO_FATAL'] = {'erro': repr(e), 'trace': traceback.format_exc()[-3000:]}
    finally:
        try:
            qa.fechar(salvar=False)
        except Exception:
            pass
    h1 = C.sha(src)
    L['producao_intacta'] = {'sha_antes': h0, 'sha_depois': h1, 'igual': h0 == h1}
    print('FIM', PROD, flush=True)


def secao(L, nome, fn, *a):
    t0 = time.time()
    try:
        r = fn(*a)
    except Exception as e:  # noqa: BLE001
        r = {'ERRO': repr(e), 'trace': traceback.format_exc()[-2500:]}
    if isinstance(r, dict):
        r['_s'] = round(time.time() - t0, 1)
    L[nome] = r
    return r


# =================================================================================================== secoes
def s_estaticas(qa, senha):
    p, e = qa.ws('Painel'), qa.ws('Estatística')
    out = {}
    # ---- OC-0043: J7:J9 "REJEITADO" (e os outros status) cabem com o Reduzir para caber (<= 30 %)
    j = {}
    for c in [f'J{7 + k}' for k in range(NLV)]:
        cel = p.Range(c)
        ma = cel.MergeArea
        larg = float(ma.Width) - 3.0
        f = cel.DisplayFormat.Font
        medidas = {}
        for s in ('REJEITADO', 'ALERTA', 'OK', 'SEM CORRIDA NO PERÍODO', 'AGUARDANDO 1ª CORRIDA',
                  'SEM CORRIDA DESTE ANALITO'):
            w = auditoria_abas.largura_pt(s, f.Name, bool(f.Bold), float(f.Size))
            medidas[s] = {'texto_pt': round(w, 1), 'fator': round(min(1.0, larg / w), 3)}
        j[c] = {'shrink': cel.ShrinkToFit, 'wrap': cel.WrapText, 'largura_pt': round(larg, 1), 'fonte':
                [f.Name, bool(f.Bold), float(f.Size)], 'valor_atual': t(cel.Value), 'medidas': medidas}
    out['O43'] = {'celulas': j, 'ok_rejeitado': all(v['shrink'] is True and v['medidas']['REJEITADO']['fator'] >= 0.70
                                                    for v in j.values()),
                  'longos_alem_30pct': sorted({s for v in j.values() for s, m in v['medidas'].items()
                                               if m['fator'] < 0.70})}
    # ---- OC-0044: A:B centralizadas na vertical
    linhas = [r for r in range(14, 134) if e.Cells(r, 1).HasFormula]
    fora = [r for r in linhas if e.Range(e.Cells(r, 1), e.Cells(r, 2)).VerticalAlignment != -4108]
    alvo44 = {}
    for r in linhas:
        nm = t(e.Cells(r, 1).Value)
        if nm in ('Lactato', 'Ácido úrico', 'Cálcio'):
            alvo44[f'{nm} N{t(e.Cells(r, 2).Value)} (L{r})'] = [e.Cells(r, 1).VerticalAlignment,
                                                                 e.Cells(r, 2).VerticalAlignment, float(e.Rows(r).RowHeight)]
    out['O44'] = {'linhas_com_formula': len(linhas), 'fora_do_centro': fora[:10], 'exemplos_do_parecer': alvo44,
                  'ok': bool(linhas) and not fora}
    # ---- OC-0036: AT = "x,xx / y,yy" com os valores das celulas de origem
    f14 = t(e.Range('AT14').Formula)
    refs = re.findall(r'FIXED\(\s*\$?([A-Z]{1,3})\$?14\b', f14)
    at_rows, ruins, conferidos = [], [], 0
    for r in linhas:
        tx = t(e.Cells(r, 46).Text)
        if not tx:
            continue
        at_rows.append((r, tx))
        if re.search(r'\b0\d\d\b', tx) and ',' not in tx:
            ruins.append((r, tx, 'zero-preenchido'))
        if len(refs) == 2:
            vals = [e.Range(f'{c}{r}').Value for c in refs]
            esp = ' / '.join(fmt2(v) if num(v) else t(v) for v in vals)
            if all(num(v) for v in vals):
                conferidos += 1
                if esp != tx:
                    ruins.append((r, tx, esp))
    out['O36'] = {'formula_AT14': f14[:220], 'refs': refs, 'n_textos': len(at_rows), 'conferidos_com_origem': conferidos,
                  'exemplos': at_rows[:5], 'ruins': ruins[:8],
                  'ok': bool(at_rows) and not ruins and any(re.match(r'-?[\d.]*\d,\d\d / ', x) for _, x in at_rows)}
    # ---- OC-0049 (Hema): linhas sem analito vazias
    if not BIO:
        vazias, sujas = [], []
        for r in linhas:
            if t(e.Cells(r, 1).Text) == '':
                vazias.append(r)
                vis = {e.Cells(1, c).Address.replace('$', '').rstrip('1') + str(r): t(e.Cells(r, c).Text)
                       for c in range(1, 48)}
                lixo = {k: v for k, v in vis.items() if v != ''}
                if lixo:
                    sujas.append((r, dict(list(lixo.items())[:6])))
        out['O49'] = {'linhas_vazias': len(vazias), 'primeira': vazias[:1], 'ultima': vazias[-1:],
                      'com_texto_visivel_A_AU': sujas[:6], 'n_sujas': len(sujas),
                      'ok_A_G_R_S_T': all(all(t(e.Cells(r, c).Text) == '' for c in (1, 7, 18, 19, 20)) for r in vazias)
                      and bool(vazias)}
    # ---- OC-0050: Principal A5
    pr = qa.ws('Principal - Resultados')
    lo = qa.lo('tblCQ_Final')
    n = lo.ListRows.Count
    part = lo.ListColumns('PARTICIPA_ESTATISTICA').DataBodyRange.Value
    sim = sum(1 for x in part if x[0] == 'SIM')
    a5 = t(pr.Range('A5').Text)
    out['O50'] = {'A5': a5[:200], 'linhas': n, 'participam': sim,
                  'ok': a5.startswith(f'Total: {milhar(n)}') and f'participam: {milhar(sim)}' in a5
                  and not re.search(r'\d,0\b', a5)}
    # ---- OC-0040 (Bio, G22): Estatistica "Ano 0" / D4 = Painel!I3 / visoes MENSAL-TRIMESTRAL
    cab = {}
    for rr in range(2, 7):
        for cc in range(1, 12):
            cel = e.Cells(rr, cc)
            if t(cel.Formula) or t(cel.Text):
                cab[cel.Address.replace('$', '')] = {'f': t(cel.Formula)[:160], 'txt': t(cel.Text)[:120],
                                                     'locked': bool(cel.Locked)}
    ano0 = {k: v for k, v in cab.items() if re.search(r'\bAno 0\b', v['txt'])}
    out['O40'] = {'cabecalho_estatistica': cab, 'com_Ano_0': ano0, 'Painel_I3': t(p.Range('I3').Value),
                  'Painel_K3': t(p.Range('K3').Value), 'D4': cab.get('D4'),
                  'D4_ref_I3': 'I3' in (cab.get('D4') or {}).get('f', ''),
                  'protegida': bool(e.ProtectContents)}
    return out


def s_guia(qa):
    g = open(GUIA, encoding='utf-8').read()
    s3 = g[g.find('## 3. Digitar'):g.find('## 4. ')]
    s5 = g[g.find('## 5. Controle externo'):g.find('## 6. O que')]
    s6 = g[g.find('## 6. O que'):g.find('## 7. Itens')]
    termos = ['Registrar resultado do CEQ', 'Salvar e próximo', 'Corrigir selecionado', 'Excluir selecionado',
              'Consolidar', 'ATUALIZAR DADOS', 'Z3', 'valor-alvo', 'FÓRMULA SOBRESCRITA', 'login']
    no_guia = {k: (k in s5 or k in s6) for k in termos}
    # os botoes e rotulos que o GUIA cita existem no arquivo?
    capt = []
    for ws in qa.wb.Worksheets:
        try:
            for sh in ws.Shapes:
                try:
                    tx = t(sh.TextFrame2.TextRange.Text)
                except Exception:
                    tx = ''
                if tx:
                    capt.append((ws.Name, tx[:80]))
        except Exception:
            pass
    frm = []
    try:
        for c in qa.wb.VBProject.VBComponents('frmCEQ').Designer.Controls:
            try:
                frm.append(t(c.Caption))
            except Exception:
                pass
    except Exception as ex:
        frm.append('ERRO ' + repr(ex))
    todos = ' || '.join(x for _, x in capt) + ' || ' + ' || '.join(frm)
    existe = {k: (k.upper() in todos.upper()) for k in ['Registrar resultado do CEQ', 'Salvar e próximo',
                                                        'Corrigir selecionado', 'Excluir selecionado', 'Consolidar',
                                                        'ATUALIZAR DADOS']}
    onde = {k: [w for w, x in capt if k.upper() in x.upper()][:4] for k in existe}
    return {'s3_equipamento': ('equipamento' in s3.lower() and 'âmbar' in s3), 'termos_s5_s6': no_guia,
            'manda_digitar_nas_abas': ('digite cada rodada nas abas' in g) or ('digite nas abas EQA' in g),
            'R4_hema_no_guia': 'R4 vale para **todos os analitos**' in s5 or 'R4 vale para todos' in s5,
            'botoes_existem': existe, 'onde': onde, 'frm_captions': [x for x in frm if x][:40]}


def s_sem_sessao(qa):
    lo_cap = qa.lo('tblEQA_CAP_Dados')
    n0 = lo_cap.ListRows.Count
    usados0 = sum(1 for r in lo_cap.ListColumns(4).DataBodyRange.Value if r[0] not in (None, ''))
    aud = Audit(qa)
    sess = qa.run('mSeguranca.SessaoAtiva')
    r = t(qa.run('mCEQForm.SalvarCEQ', C.campos_ceq('CAP', ROD, ANO, NOME_CEQ, 'TA-33', '110', unid=UN_CEQ)))
    usados1 = sum(1 for x in lo_cap.ListColumns(4).DataBodyRange.Value if x[0] not in (None, ''))
    return {'SessaoAtiva': sess, 'retorno': r, 'linhas_usadas': [usados0, usados1], 'linhas_tabela': [n0, lo_cap.ListRows.Count],
            'audit': aud.novos(), 'ok': r.startswith('ERRO|0|Sessão encerrada') and usados0 == usados1}


def abrir_ceq_sem_sessao(qa):
    """AbrirCEQ sem login: a caixa da propria instancia e lida e fechada pelo Vigia."""
    vg = Vigia(qa.ex.pid, teto=40)
    vg.start()
    t0 = time.time()
    erro = None
    try:
        qa.ex.run("'" + qa.wb.Name + "'!mCEQForm.AbrirCEQ", teto=60)
    except Exception as ex:  # noqa: BLE001
        erro = repr(ex)
    dur = time.time() - t0
    vg.parar.set()
    vg.join(3)
    caixas = vg.vistos
    msg = ' '.join(' '.join(c['textos']) for c in caixas)
    return {'caixas': caixas, 'erro_run': erro, 'segundos': round(dur, 1),
            'ok': any(c['classe'] == '#32770' for c in caixas) and 'Faça o login para registrar resultados do CEQ' in msg
            and not any(c['classe'] == 'ThunderDFrame' for c in caixas)}


def contar_periodo(fin, analito, lote, equip, de, ate):
    cont = collections.Counter()
    for r in fin:
        if up(r['ANALITO']) != up(analito) or r['PARTICIPA_ESTATISTICA'] != 'SIM' or not num(r['RESULTADO']):
            continue
        if lote is not None and t(r['LOTE']) != t(lote):
            continue
        if equip and up(r['EQUIPAMENTO']) != up(equip):
            continue
        d = r['DATA_HORA']
        if not isinstance(d, dt.datetime):
            continue
        if isinstance(de, dt.datetime) and d.date() < de.date():
            continue
        if isinstance(ate, dt.datetime) and d.date() > ate.date():
            continue
        cont[int(r['NIVEL'] or 0)] += 1
    return ' · '.join(f'N{k} {cont.get(k, 0)}' for k in range(1, NLV + 1)), cont


def painel_O12(qa, seq, limpar_h3, fin):
    p, eng = qa.ws('Painel'), qa.ws('Eng_Saida')
    for c, v in seq:
        C.por(qa, p, c, v, local=False)
    if limpar_h3:
        p.Range('H3').MergeArea.ClearContents()
        qa.ex.esperar()
        time.sleep(1)
    cort = int(eng.Range('M1').Value or 0)
    af1 = t(eng.Range('AF1').Value)
    a39 = t(p.Range('A39').Value)
    ps = qa.painel_stats()
    de, ate = qa.nome('filtroDe').Value, qa.nome('filtroAte').Value
    de = C.limpo(de) if isinstance(de, dt.datetime) else None
    ate = C.limpo(ate) if isinstance(ate, dt.datetime) else None
    todos = eng.Range('AD1').Value is True
    an = t(qa.nome('selAnalito').Value)
    lote = None if todos else t(eng.Range('E1').Value).lstrip("'")
    eq = t(eng.Range('S1').Value) if BIO else None
    indep, cont = contar_periodo(fin, an, lote, eq, de, ate)
    n_jan = [ps[k][1] for k in range(NLV)]
    return {'filtro': {'analito': an, 'lote': lote, 'todos': todos, 'equip': eq, 'de': str(de), 'ate': str(ate),
                       'K4': t(p.Range('K4').Value)},
            'cortadas_M1': cort, 'AF1': af1, 'independente': indep, 'n_janela_painel': n_jan, 'A39': a39[-300:],
            'ok': cort > 0 and 'indicadores do Painel' in a39 and af1 == indep and f'Período inteiro: {af1}' in a39
            and all(n_jan[k] is not None and n_jan[k] <= cont.get(k + 1, 0) for k in range(NLV))}


def s_O12(qa, fin):
    out = {}
    if BIO:
        out['original_Glicose_8974_DIM1_Tudo'] = painel_O12(
            qa, [('L4', 'DIMENSION 1'), ('C3', 'Glicose'), ('H3', "'8974"), ('K4', 'Tudo (sem filtro)')], False, fin)
        # ampliacao: o maior (analito, lote, equipamento) com > 180 por nivel, fora o do parecer
        g = collections.Counter((t(r['ANALITO']), t(r['LOTE']), t(r['EQUIPAMENTO']))
                                for r in fin if r['PARTICIPA_ESTATISTICA'] == 'SIM' and int(r['NIVEL'] or 0) == 1)
        cand = [(k, n) for k, n in g.most_common(40) if n > 200 and k != ('Glicose', '8974', 'DIMENSION 1')
                and k[0] not in ('Glicose',)]
        if cand:
            (an, lt, eq), n = cand[0]
            out['ampliacao'] = painel_O12(qa, [('L4', eq), ('C3', an), ('H3', "'" + lt), ('K4', 'Tudo (sem filtro)')],
                                          False, fin)
            out['ampliacao']['escolhido'] = [an, lt, eq, n]
    else:
        out['original_WBC_TODOS_2026'] = painel_O12(qa, [('C3', 'WBC'), ('K3', 2026), ('K4', 'Ano inteiro')], True, fin)
        out['ampliacao_RBC_TODOS_2026'] = painel_O12(qa, [('C3', 'RBC'), ('K3', 2026), ('K4', 'Ano inteiro')], True, fin)
    out['ok'] = all(v.get('ok') for v in out.values() if isinstance(v, dict))
    return out


def s_O41(qa, senha):
    """Bonus: o que o ADR-090 implementou como "OC-0040" e o G23 (registrado como OC-0041)."""
    p = qa.ws('Painel')
    C.por(qa, p, 'K3', 2026, local=False)
    C.por(qa, p, 'K4', 'S1', local=False)
    p.Protect(senha, True, True)
    passos = []
    try:
        p.Range('G3').Value = dt.datetime(2026, 9, 1)
        p.Range('G4').Value = dt.datetime(2026, 10, 31)
        qa.ex.esperar()
        passos.append(['De/Até 01/09–31/10/2026', t(p.Range('K3').Value), t(p.Range('K4').Value)])
        p.Range('K4').Value = 'Out'
        qa.ex.esperar()
        passos.append(['K4 Out', t(p.Range('K3').Value), t(p.Range('K4').Value), str(p.Range('G3').Value)[:10],
                       str(p.Range('G4').Value)[:10]])
    finally:
        p.Unprotect(senha)
    return {'passos': passos, 'ok': passos[0][1:] == ['2026', 'Personalizado'] and passos[1][3:] == ['2026-10-01', '2026-10-31']}


def s_adr091(qa, fin):
    out = {}
    lo = qa.lo('tblLotesMaterial')
    cab = [str(c.Name) for c in lo.ListColumns]
    regras = []
    if lo.ListRows.Count:
        for l in lo.DataBodyRange.Value:
            d = dict(zip(cab, [t(x) for x in l]))
            if any(d.values()):
                regras.append(d)
    out['tabela'] = {'aba': lo.Parent.Name, 'colunas': cab, 'n': len(regras), 'regras': regras}
    out['resumo_vba'] = t(qa.run('mEtlCQ.LotesMaterialResumo'))
    fora = {up(r['LOTE']) for r in regras if up(r['REGRA']) == 'FORA_DO_CQ'}
    so = collections.defaultdict(set)
    for r in regras:
        if up(r['REGRA']) == 'SO_ANALITOS':
            for a in r['ANALITOS'].split(';'):
                if a.strip():
                    so[up(r['LOTE'])].add(up(a))
    mz = {}
    for r in regras:
        if up(r['REGRA']) == 'MATRIZ' and up(r['LOTE']) not in mz:
            mz[up(r['LOTE'])] = up(r['MATRIZ'])
    esp = qa.espec()
    cat = {up(x[0]) for x in esp.Range('A4:A43').Value if x[0]}
    st = collections.Counter(r['STATUS_ANALITICO'] for r in fin)
    out['status_copia'] = dict(st)
    # ---- regra por regra
    lotes = collections.defaultdict(list)
    for r in fin:
        lotes[up(r['LOTE'])].append(r)
    viol = []
    det = {}
    for lt in sorted(fora):
        rs = lotes.get(lt, [])
        c = collections.Counter(r['STATUS_ANALITICO'] for r in rs)
        part = sum(1 for r in rs if r['PARTICIPA_ESTATISTICA'] == 'SIM')
        for r in rs:
            if r['STATUS_ANALITICO'] != 'LOTE_FORA_DO_CQ' and r['STATUS_ANALITICO'] not in PRECEDE_LOTE:
                viol.append(('FORA', lt, r['ID_REGISTRO'], r['STATUS_ANALITICO']))
        det[f'FORA {lt}'] = {'linhas': len(rs), 'status': dict(c), 'participam': part}
    for lt, ans in sorted(so.items()):
        rs = lotes.get(lt, [])
        dentro = [r for r in rs if up(r['ANALITO']) in ans]
        outros = [r for r in rs if up(r['ANALITO']) not in ans]
        for r in outros:
            if r['STATUS_ANALITICO'] != 'LOTE_OUTRO_MATERIAL' and r['STATUS_ANALITICO'] not in PRECEDE_LOTE:
                viol.append(('SO', lt, r['ID_REGISTRO'], r['ANALITO'], r['STATUS_ANALITICO']))
            if r['PARTICIPA_ESTATISTICA'] == 'SIM':
                viol.append(('SO-participa', lt, r['ID_REGISTRO']))
        for r in dentro:
            if r['STATUS_ANALITICO'] in LOTE_ST:
                viol.append(('SO-dentro', lt, r['ID_REGISTRO'], r['STATUS_ANALITICO']))
        det[f'SO {lt}'] = {'linhas': len(rs), 'dentro': dict(collections.Counter(r['STATUS_ANALITICO'] for r in dentro)),
                           'dentro_participam': sum(1 for r in dentro if r['PARTICIPA_ESTATISTICA'] == 'SIM'),
                           'outros': dict(collections.Counter(r['STATUS_ANALITICO'] for r in outros)),
                           'outros_analitos': len({r['ANALITO'] for r in outros}),
                           'outros_participam': sum(1 for r in outros if r['PARTICIPA_ESTATISTICA'] == 'SIM')}
    # de/para da URINA (ANALITO_ORIGEM -> ANALITO_QCINI)
    dp = qa.ler('tblDeParaAnalitos')
    depara = {(up(d['MATRIZ']), up(d['ANALITO_ORIGEM'])): t(d['ANALITO_QCINI']) for d in dp if d.get('ANALITO_ORIGEM')}
    for lt, m in sorted(mz.items()):
        rs = lotes.get(lt, [])
        mzs = collections.Counter(up(r['MATRIZ']) for r in rs)
        sericos = [r for r in rs if up(r['ANALITO']) in cat]
        sericos_part = [r for r in sericos if r['PARTICIPA_ESTATISTICA'] == 'SIM']
        # de/para da matriz da regra; ANALITO_QCINI vazio = "sem cadastro": o ANALITO fica o da origem (CRE2, UCFP...)
        dp_err = [(r['ID_REGISTRO'], r['ANALITO_ORIGEM'], r['ANALITO'], depara[(m, up(r['ANALITO_ORIGEM']))]) for r in rs
                  if r['ORIGEM_RESULTADO'] == 'INTERFACEAMENTO' and (m, up(r['ANALITO_ORIGEM'])) in depara
                  and (depara[(m, up(r['ANALITO_ORIGEM']))] or t(r['ANALITO_ORIGEM'])) != t(r['ANALITO'])]
        if any(k != m for k in mzs):
            viol.append(('MATRIZ', lt, dict(mzs)))
        if sericos_part:
            viol.append(('MATRIZ-serico-participa', lt, len(sericos_part)))
        if dp_err:
            viol.append(('MATRIZ-depara', lt, dp_err[:3]))
        det[f'MATRIZ {lt}'] = {'linhas': len(rs), 'matriz': dict(mzs), 'em_analito_serico': len(sericos),
                               'serico_participa': len(sericos_part),
                               'analitos': dict(collections.Counter(t(r['ANALITO']) for r in rs).most_common(8)),
                               'participam': sum(1 for r in rs if r['PARTICIPA_ESTATISTICA'] == 'SIM')}
    # LOTE_* so nos lotes com regra; colunas derivadas
    regra_lotes = fora | set(so)
    lote_rows = [r for r in fin if r['STATUS_ANALITICO'] in LOTE_ST]
    fora_regra = [r['ID_REGISTRO'] for r in lote_rows if up(r['LOTE']) not in regra_lotes]
    deriv = collections.Counter((r['STATUS_ANALITICO'], r['PARTICIPA_ESTATISTICA'], t(r['REGISTRAR_RESULTADO_NO_LJ']),
                                 t(r['TIPO_PLOTAGEM_LJ']), t(r['MOTIVO_EXCLUSAO_AUTOMATICA'])[:70]) for r in lote_rows)
    sem_run = sum(1 for r in lote_rows if not num(r['RUN']))
    leq0 = [(r['ID_REGISTRO'], r['RESULTADO'], r['STATUS_ANALITICO']) for r in fin
            if up(r['LOTE']) in regra_lotes and num(r['RESULTADO']) and r['RESULTADO'] <= 0]
    out['por_regra'] = det
    out['violacoes'] = viol[:30]
    out['n_violacoes'] = len(viol)
    out['LOTE_total'] = len(lote_rows)
    out['LOTE_fora_dos_lotes_com_regra'] = fora_regra[:10]
    out['colunas_derivadas_LOTE'] = [list(k) + [n] for k, n in deriv.items()]
    out['LOTE_sem_RUN'] = sem_run
    out['resultado_leq0_nos_lotes_com_regra'] = leq0[:20]
    # ---- QA A13 / A14
    qai = qa.ler('tblQA_Integracao')
    a13 = [x for x in qai if x.get('CODIGO') == 'A13']
    a14 = [x for x in qai if x.get('CODIGO') == 'A14']
    ids_lote = {t(r['ID_REGISTRO']) for r in lote_rows}
    ids_a13 = set()
    for x in a13:
        for v in x.values():
            if t(v) in ids_lote:
                ids_a13.add(t(v))
    out['A13'] = {'n': len(a13), 'severidades': dict(collections.Counter(x.get('SEVERIDADE') for x in a13)),
                  'ids_LOTE_cobertos': len(ids_a13), 'ids_LOTE': len(ids_lote),
                  'exemplo': {k: t(v)[:90] for k, v in (a13[0].items() if a13 else [])}}
    out['A14'] = {'n': len(a14), 'textos': [' | '.join(t(v) for v in x.values() if v)[:200] for x in a14[:5]]}
    # ---- cenario original da OC-0007: lotes de urina/PCR/fora ativos em analitos sericos
    orig = ['5659', '6353', '6354', '8353', '8177', '8794']
    ser = collections.Counter()
    for r in fin:
        if up(r['LOTE']) in orig and up(r['ANALITO']) in cat and r['PARTICIPA_ESTATISTICA'] == 'SIM':
            if up(r['LOTE']) == '8353' and up(r['ANALITO']) in so.get('8353', set()):
                continue
            ser[(t(r['LOTE']), t(r['ANALITO']))] += 1
    out['OC0007_sericos_participando'] = {f'{a}|{b}': n for (a, b), n in ser.items()}
    # ---- ampliacao: lotes candidatos fora da tabela
    por_lote_an = collections.defaultdict(collections.Counter)
    for r in fin:
        if r['PARTICIPA_ESTATISTICA'] == 'SIM' and up(r['ANALITO']) in cat:
            por_lote_an[t(r['LOTE'])][t(r['ANALITO'])] += 1
    painel_unico = {lt: {'analitos': len(c), 'max_por_analito': max(c.values()), 'total': sum(c.values())}
                    for lt, c in por_lote_an.items() if len(c) >= 8 and max(c.values()) <= 4 and up(lt) not in regra_lotes
                    and up(lt) not in mz}
    familia = {lt: dict(c.most_common(5)) for lt, c in por_lote_an.items()
               if lt[:3] in ('635', '565', '835') and up(lt) not in regra_lotes and up(lt) not in mz}
    pcr_fora = collections.Counter(t(r['LOTE']) for r in fin if up(r['ANALITO']) == up('Proteína C-reativa')
                                   and r['PARTICIPA_ESTATISTICA'] == 'SIM')
    out['ampliacao'] = {'paineis_rodados_uma_vez_fora_da_tabela': painel_unico, 'familias_urina_pcr_fora_da_tabela': familia,
                        'lotes_de_PCR_participando': dict(pcr_fora)}
    if BIO:
        out['ok'] = (not viol and not fora_regra and len(a13) == len(lote_rows) and len(ids_a13) == len(ids_lote)
                     and not a14 and not ser and len(regras) > 0)
    else:
        out['ok'] = (len(regras) == 0 and not lote_rows and not a13 and not a14)
    return out


def s_manuais(qa, fin, aud):
    """OC-0019/0021/0023/0024 + manuais do ADR-091 (Bio)."""
    out = {}
    p = qa.ws('Painel')
    dr = qa.ws('Digitar Resultados')
    lo, cab = manual_cols(qa)
    out['colunas'] = cab
    out['instrucao_A3'] = t(dr.Range('A3').Value)[:400]
    if BIO:
        C.por(qa, p, 'L4', 'DIMENSION 2', local=False)
        C.por(qa, p, 'C3', 'Glicose', local=False)
    sel_eq = t(qa.nome('selEquipamento').Value)
    out['selEquipamento'] = sel_eq
    an = 'Glicose' if BIO else 'WBC'
    reais = sorted([r for r in fin if r['STATUS_ANALITICO'] == 'ATIVO' and r['ORIGEM_RESULTADO'] == 'INTERFACEAMENTO'
                    and t(r['ANALITO']) == an and int(r['NIVEL'] or 0) == 1 and isinstance(r['DATA_HORA'], dt.datetime)
                    and (not BIO or t(r['EQUIPAMENTO']) == 'DIMENSION 1')], key=lambda r: r['DATA_HORA'])
    R = reais[-40]
    dR = R['DATA_HORA']
    out['referencia_real'] = {k: t(R[k]) for k in ('ID_REGISTRO', 'EQUIPAMENTO', 'LOTE', 'NIVEL', 'ANALITO', 'RESULTADO',
                                                   'DATA_HORA', 'RUN')}
    dia = dR.strftime('%d/%m/%Y')
    dr.Activate()
    qa.ex.xl.EnableEvents = True
    linhas = {}

    def novo(rot, campos, sem_equip=False):
        lo2, r = C.linha_livre_manual(qa)
        for col, v in campos:
            escrever_celula(qa, r, col, v)
        linhas[rot] = r
        return r

    # M1: sem EQUIPAMENTO (G01) e sem NIVEL (G05: a chave "0|0|GLICOSE" do ciclo 01). O MOTIVO nao e obrigatorio no
    # manual comum (mEtlCQ: so com SUBSTITUI_ID) -- 1a rodada desta validacao, em que o manual sem MOTIVO ficou ATIVO.
    r1 = novo('M1_incompleto', [('DATA', dia), ('HORA', '03:17'), ('LOTE', "'" + t(R['LOTE'])),
                                ('ANALITO', an), ('RESULTADO', '123,456'), ('MOTIVO', MOTIVO)])
    ceq = lo.ListColumns('EQUIPAMENTO').DataBodyRange.Cells(r1, 1)
    eq1 = {'valor': t(ceq.Value), 'fundo': ceq.Interior.Color, 'fonte': ceq.Font.Color,
           'nota': t(ceq.Comment.Text()) if ceq.Comment is not None else ''}
    out['O19_auto'] = eq1
    if BIO:
        escrever_celula(qa, r1, 'EQUIPAMENTO', 'DIMENSION 1')
        eq2 = {'valor': t(ceq.Value), 'fundo': ceq.Interior.Color, 'nota': ceq.Comment is not None}
        out['O19_digitado'] = eq2
        out['O19_ok'] = (eq1['valor'] == sel_eq == 'DIMENSION 2' and eq1['fundo'] == AMBAR and eq1['fonte'] == AMBAR_FONTE
                         and eq1['nota'] == f'Preenchido pelo Painel com {sel_eq}: confira antes de ATUALIZAR DADOS.'
                         and eq2['valor'] == 'DIMENSION 1' and eq2['fundo'] != AMBAR and not eq2['nota']
                         and 'EQUIPAMENTO' in out['instrucao_A3'])
    else:
        out['O19_ok'] = (eq1['valor'] == sel_eq and eq1['fundo'] != AMBAR and eq1['nota'] == ''
                         and 'EQUIPAMENTO' in out['instrucao_A3'])
    # M2: conflito (mesma chave do real, +5 min, com MOTIVO)
    hm = (dR + dt.timedelta(minutes=5)).strftime('%H:%M')
    novo('M2_conflito', [('DATA', dia), ('HORA', hm), ('EQUIPAMENTO', t(R['EQUIPAMENTO'])), ('LOTE', "'" + t(R['LOTE'])),
                         ('NIVEL', '1'), ('ANALITO', an), ('RESULTADO', '111,111'), ('MOTIVO', MOTIVO)])
    # M7: implausivel num lote comum (ampliacao da trilha)
    novo('M7_implausivel', [('DATA', dia), ('HORA', '03:29'), ('EQUIPAMENTO', t(R['EQUIPAMENTO'])),
                            ('LOTE', "'" + t(R['LOTE'])), ('NIVEL', '1'), ('ANALITO', an), ('RESULTADO', '-1'),
                            ('MOTIVO', MOTIVO)])
    # M8: sem HORA (outro campo obrigatorio)
    novo('M8_sem_hora', [('DATA', dia), ('EQUIPAMENTO', t(R['EQUIPAMENTO'])), ('LOTE', "'" + t(R['LOTE'])), ('NIVEL', '1'),
                         ('ANALITO', an), ('RESULTADO', '123,456'), ('MOTIVO', MOTIVO)])
    if BIO:
        base = [('DATA', dia), ('EQUIPAMENTO', 'DIMENSION 1'), ('NIVEL', '1'), ('MOTIVO', MOTIVO)]
        novo('M3_8353_Glicose', base + [('HORA', '03:21'), ('LOTE', "'8353"), ('ANALITO', 'Glicose'), ('RESULTADO', '123,456')])
        novo('M4_8353_PCR', base + [('HORA', '03:23'), ('LOTE', "'8353"), ('ANALITO', 'Proteína C-reativa'),
                                    ('RESULTADO', '1,234')])
        novo('M5_8794_menos1', base + [('HORA', '03:25'), ('LOTE', "'8794"), ('ANALITO', 'Glicose'), ('RESULTADO', '-1')])
        novo('M6_6354_Creatinina', base + [('HORA', '03:27'), ('LOTE', "'6354"), ('ANALITO', 'Creatinina'),
                                           ('RESULTADO', '12,345')])
    # a validacao de dados da aba aceitaria o que foi escrito pelo COM? (Validation.Value da celula)
    val = {}
    for k, r in linhas.items():
        for col in ('LOTE', 'ANALITO', 'NIVEL', 'EQUIPAMENTO'):
            cel = lo.ListColumns(col).DataBodyRange.Cells(r, 1)
            try:
                vv = cel.Validation
                val[f'{k}.{col}'] = [vv.Type, t(vv.Formula1)[:60], bool(vv.Value)]
            except Exception:
                val[f'{k}.{col}'] = 'sem validacao'
    out['validacao_dados'] = val
    aud.novos()
    t0 = time.time()
    r_at1 = t(qa.run('mIntegracao.AtualizarSoEntradas', teto=1500))
    out['atualizar1'] = {'ret': r_at1[:120], 's': round(time.time() - t0, 1)}
    trilha1 = aud.novos()
    leit1 = {k: ler_manual(qa, r) for k, r in linhas.items()}
    ids = {k: v['ID_REGISTRO'][0] for k, v in leit1.items()}
    out['ids'] = ids
    fin1 = qa.final_por_id()
    out['manuais_1'] = {k: {'STATUS': v['STATUS'][0], 'RUN_txt': v['RUN'][1], 'DETALHE_txt': v['DETALHE'][1][:160],
                            'MATRIZ': v.get('MATRIZ', ('', ''))[0],
                            'final': {c: t(fin1.get(ids[k], {}).get(c)) for c in
                                      ('STATUS_ANALITICO', 'PARTICIPA_ESTATISTICA', 'ANALITO', 'MATRIZ',
                                       'TIPO_PLOTAGEM_LJ', 'MOTIVO_EXCLUSAO_AUTOMATICA', 'RUN')}}
                        for k, v in leit1.items()}
    out['trilha_1'] = [{'acao': x['acao'], 'id': next((k for k, i in ids.items() if i and i in x['txt']), None),
                        'txt': x['txt'][:1200]} for x in trilha1 if x['acao']]
    # O24: ResumoQA
    rq = t(qa.run('mIntegracao.ResumoQA'))
    qai = qa.ler('tblQA_Integracao')
    err = collections.Counter(x['CODIGO'] for x in qai if x['SEVERIDADE'] == 'ERRO')
    ngov = sum(err.get(c, 0) for c in ('E01', 'E02', 'E03', 'E10', 'E11'))
    out['O24'] = {'resumo': rq.split('\r\n'), 'erros_por_codigo': dict(err), 'gov': ngov, 'len': len(rq),
                  'E10': [' | '.join(t(v) for v in x.values() if v)[:260] for x in qai
                          if x['SEVERIDADE'] == 'ERRO' and x['CODIGO'] in ('E01', 'E02', 'E03', 'E10', 'E11')][:6]}
    lin = rq.split('\r\n')
    out['O24']['ok'] = (err.get('E04', 0) > 0 and any(x.startswith(f'• {err["E04"]} resultado(s) manual(is) incompleto(s)')
                                                       for x in lin)
                        and 'Há ERRO DE GOVERNANÇA' not in rq
                        and (('erro(s) de governança' in rq) == (ngov > 0))
                        and (ngov == 0 or any(x.startswith(f'• {ngov} erro(s) de governança') for x in lin))
                        and len(rq) <= 1024)
    # M1: completar o NIVEL e atualizar de novo
    escrever_celula(qa, linhas['M1_incompleto'], 'NIVEL', '1')
    t0 = time.time()
    r_at2 = t(qa.run('mIntegracao.AtualizarSoEntradas', teto=1500))
    out['atualizar2'] = {'ret': r_at2[:120], 's': round(time.time() - t0, 1)}
    trilha2 = aud.novos()
    m1b = ler_manual(qa, linhas['M1_incompleto'])
    out['M1_depois'] = {'STATUS': m1b['STATUS'][0], 'RUN_txt': m1b['RUN'][1], 'DETALHE_txt': m1b['DETALHE'][1]}
    out['trilha_2'] = [{'acao': x['acao'], 'txt': x['txt'][:300]} for x in trilha2 if x['acao'] and ids['M1_incompleto'] in x['txt']]
    # O21 / O23 vereditos
    m = out['manuais_1']
    t1 = collections.defaultdict(list)
    for x in out['trilha_1']:
        t1[x['id']].append(x['acao'])
    out['O23_ok'] = (m['M1_incompleto']['STATUS'] == 'MANUAL_INCOMPLETO' and t1['M1_incompleto'] == ['RESULTADO_MANUAL_RECUSADO']
                     and any('NIVEL' in x['txt'] for x in out['trilha_1'] if x['id'] == 'M1_incompleto')
                     and m['M8_sem_hora']['STATUS'] == 'MANUAL_INCOMPLETO' and t1['M8_sem_hora'] == ['RESULTADO_MANUAL_RECUSADO']
                     and m['M2_conflito']['STATUS'] == 'CONFLITO_MANUAL'
                     and t1['M2_conflito'] == ['RESULTADO_MANUAL_EM_CONFLITO']
                     and out['M1_depois']['STATUS'] == 'ATIVO'
                     and [x['acao'] for x in out['trilha_2']] == ['RESULTADO_MANUAL_INCLUIDO'])
    out['O23_trilha_por_manual'] = dict(t1)
    out['O21_ok'] = (m['M1_incompleto']['RUN_txt'] == '' and m['M2_conflito']['RUN_txt'] == '' and m['M8_sem_hora']['RUN_txt'] == ''
                     and m['M1_incompleto']['DETALHE_txt'] not in ('', '0') and 'obrigatório' in m['M1_incompleto']['DETALHE_txt']
                     and out['M1_depois']['DETALHE_txt'] == '' and out['M1_depois']['RUN_txt'] not in ('', '0'))
    # O21: varredura das linhas que ja existiam (RUN/DETALHE com "0")
    n = lo.ListRows.Count
    vr = lo.ListColumns('RUN').DataBodyRange
    vd = lo.ListColumns('DETALHE').DataBodyRange
    vi = lo.ListColumns('ID_REGISTRO').DataBodyRange.Value
    zeros = []
    sem_acento = []
    for i in range(1, min(n, 400) + 1):
        if vi[i - 1][0] in (None, ''):
            continue
        a, b = t(vr.Cells(i, 1).Text), t(vd.Cells(i, 1).Text)
        if a == '0' or b == '0':
            zeros.append((t(vi[i - 1][0]), a, b))
        if re.search(r'obrigatorio|invalido|lancamento|ambiguo|numerico', b):
            sem_acento.append((t(vi[i - 1][0]), b[:80]))
    out['O21_varredura'] = {'linhas': n, 'com_zero': zeros[:10], 'sem_acento': sem_acento[:10]}
    out['O21_ok'] = out['O21_ok'] and not zeros and not sem_acento
    # ADR-091 manuais (Bio)
    if BIO:
        f = {k: v['final'] for k, v in m.items()}
        out['adr091_manuais_ok'] = (f['M3_8353_Glicose']['STATUS_ANALITICO'] == 'LOTE_OUTRO_MATERIAL'
                                    and f['M3_8353_Glicose']['PARTICIPA_ESTATISTICA'] != 'SIM'
                                    and f['M4_8353_PCR']['STATUS_ANALITICO'] == 'ATIVO'
                                    and f['M4_8353_PCR']['PARTICIPA_ESTATISTICA'] == 'SIM'
                                    and f['M5_8794_menos1']['STATUS_ANALITICO'] == 'LOTE_FORA_DO_CQ'
                                    and f['M6_6354_Creatinina']['MATRIZ'] == 'URINA')
        # o manual de urina com o nome do analito de SORO entra no Painel/Estatistica da Creatinina?
        p.Activate()
        for c, v in (('L4', 'DIMENSION 1'), ('C3', 'Creatinina'), ('H3', "'6354"), ('K4', 'Tudo (sem filtro)')):
            C.por(qa, p, c, v, local=False)
        time.sleep(1)
        eng = qa.ws('Eng_Saida')
        est = C.estat_tabela(qa)
        try:
            lista = t(p.Range('H3').Validation.Formula1)[:200]
        except Exception:
            lista = 'sem validacao'
        out['M6_no_painel'] = {'H3': t(p.Range('H3').Value), 'E1': t(eng.Range('E1').Value), 'engPainel': qa.painel_stats(),
                               'J7': t(p.Range('J7').Value), 'A39': t(p.Range('A39').Value)[:260], 'lista_H3': lista,
                               'estat_Creatinina': {k: [v['n'], v['media']] for k, v in est.items()
                                                    if k.startswith('Creatinina|')}}
    return out


def s_ceq(qa, senha, aud):
    out = {}
    cap = qa.ws('EQA.CAP_Dados')
    base = qa.ws('EQA_Base')
    lo_cap = qa.lo('tblEQA_CAP_Dados')
    # ---- OC-0034: listas sem a simulacao (oraculo: Arquivo fonte da propria linha)
    o34 = {}
    for prov, tb in (('CAP', 'tblEQA_CAP_Dados'), ('Controllab', 'tblEQA_Controllab_Dados')):
        try:
            lo = qa.lo(tb)
        except Exception:
            continue
        if lo.ListRows.Count == 0:
            continue
        v = lo.DataBodyRange.Value
        res = {}
        for col, rot in ((2, 'rodada'), (4, 'analito'), (12, 'avaliacao')):
            reais = {up(r[col - 1]) for r in v if t(r[col - 1]) and 'SIMULA' not in up(r[17])}
            simu = {up(r[col - 1]) for r in v if t(r[col - 1]) and 'SIMULA' in up(r[17])} - reais
            lista = {up(x) for x in (qa.run('mCEQForm.DistintosCEQ', prov, col) or ())}
            res[rot] = {'n_lista': len(lista), 'n_esperado': len(reais), 'igual': lista == reais,
                        'so_simulacao': sorted(simu)[:10], 'simulacao_na_lista': sorted(simu & lista)[:10]}
        o34[prov] = res
    out['O34'] = o34
    out['O34_ok'] = bool(o34) and all(x['igual'] and not x['simulacao_na_lista'] for p in o34.values() for x in p.values())
    # ---- OC-0052: pares do mapa por prefixo SEM espaco (oraculo: mapa lido da EQA_Base W:Y)
    mapa = [(t(m[0]), t(m[1]), t(m[2])) for m in C.bloco(base, 'W2:Y600') if t(m[1])]
    o52 = {'mapa_linhas': len(mapa), 'cabecalho': [t(x) for x in C.bloco(base, 'W1:Y1')[0]]}
    for prov in ('CAP', 'Controllab'):
        nomes = {}
        for pv, nm, cn in mapa:
            if up(pv) == up(prov) and nm not in nomes:
                nomes[nm] = cn
        pares = []
        for c, cc in nomes.items():
            palavra = [l for l, cl in nomes.items() if up(l).startswith(up(c) + ' ') and up(cl) != up(cc)]
            literal = [l for l, cl in nomes.items() if up(l).startswith(up(c)) and len(l) > len(c) and up(cl) != up(cc)]
            if literal and not palavra:
                av = t(qa.run('mCEQForm.AvisoAnalitoCEQ', prov, c))
                pares.append({'digitado': c, 'pares_literais': literal[:3], 'aviso': av[-140:],
                              'cita_par': any(f'“{l}”' in av for l in literal)})
        o52[prov] = {'nomes': len(nomes), 'pares_sem_espaco': pares}
    o52['MCH_hema'] = t(qa.run('mCEQForm.AvisoAnalitoCEQ', 'CAP', 'MCH')) if not BIO else None
    o52['par_de_palavra'] = t(qa.run('mCEQForm.AvisoAnalitoCEQ', 'CAP', 'Lymphs' if not BIO else 'Lactate'))
    out['O52'] = o52
    tod = [p for k in ('CAP', 'Controllab') for p in o52.get(k, {}).get('pares_sem_espaco', [])]
    out['O52_ok'] = all(p['cita_par'] for p in tod) and (BIO or 'MCHC' in (o52['MCH_hema'] or ''))
    out['O52_n_pares'] = len(tod)
    # ---- OC-0031 / OC-0055 (sessao ativa)
    g0 = {k: v['G_bias'] for k, v in C.estat_tabela(qa).items() if k.startswith(AN_CEQ + '|')}
    c_sem = C.campos_ceq('CAP', ROD, ANO, NOME_CEQ, 'TA-31', '110', alvo='', unid=UN_CEQ)
    c_com = C.campos_ceq('CAP', ROD, ANO, NOME_CEQ, 'TA-55', '110', unid=UN_CEQ)
    conf_sem = t(qa.run('mCEQForm.ConfirmacoesCEQ', c_sem, 0))
    conf_com = t(qa.run('mCEQForm.ConfirmacoesCEQ', c_com, 0))
    r_sem = t(qa.run('mCEQForm.SalvarCEQ', c_sem))
    g1 = {k: v['G_bias'] for k, v in C.estat_tabela(qa).items() if k.startswith(AN_CEQ + '|')}
    a_sem = aud.novos()
    r_com = t(qa.run('mCEQForm.SalvarCEQ', c_com))
    g2 = {k: v['G_bias'] for k, v in C.estat_tabela(qa).items() if k.startswith(AN_CEQ + '|')}
    a_com = aud.novos()
    l_sem = int(r_sem.split('|')[1]) if r_sem.startswith('OK|') else 0
    l_com = int(r_com.split('|')[1]) if r_com.startswith('OK|') else 0
    desc = t(qa.run('mCEQForm.DescricaoCEQ', 'CAP', l_com)) if l_com else ''
    regs = qa.run('mCEQForm.RegistrosDaRodada', 'CAP', ROD, ANO)
    regs = [[t(x) for x in l] for l in regs] if isinstance(regs, tuple) else regs
    des = qa.wb.VBProject.VBComponents('frmCEQ').Designer
    lbl = {n: t(des.Controls(n).Caption) for n in ('lblAlvo', 'lblSub')}
    out['O31'] = {'G0': g0, 'G_depois_sem_alvo': g1, 'G_depois_com_alvo': g2, 'confirmacao_sem_alvo': conf_sem,
                  'confirmacao_com_alvo': conf_com, 'salvar_sem_alvo': r_sem, 'salvar_com_alvo': r_com, 'rotulos': lbl,
                  'audit': [x['acao'] for x in a_sem + a_com]}
    out['O31_ok'] = ('NÃO entra no bias' in conf_sem and 'valor-alvo' not in conf_com and r_sem.startswith('OK|')
                     and 'sem valor-alvo' in r_sem and g1 == g0 and g2 != g0 and lbl['lblAlvo'] == 'Valor-alvo ¹'
                     and 'valor-alvo' in lbl['lblSub'])
    # O55
    l_exc = t(qa.run('mCEQForm.ExcluirCEQ', 'CAP', l_com)) if l_com else ''
    a_exc = aud.novos()
    textos55 = [desc] + [' '.join(x) for x in (regs or [])] + [x['txt'] for x in a_com + a_exc]
    out['O55'] = {'descricao': desc, 'registros_da_rodada': regs, 'excluir': l_exc,
                  'trilha': [x['txt'][:260] for x in a_com + a_exc],
                  'repete_ano': [x[:160] for x in textos55 if '2025 2025' in x]}
    out['O55_ok'] = bool(desc) and 'C-A 2025' in desc and not out['O55']['repete_ano']
    out['linha_sem_alvo'] = l_sem
    # ---- OC-0053: sobrescrever |Bias| (P) numa linha real mapeada (cenario do reteste02)
    cap.Activate()
    cap.Unprotect(senha)
    v = lo_cap.DataBodyRange.Value
    r0 = lo_cap.DataBodyRange.Row
    mapeados = {up(nm) for pv, nm, cn in mapa if up(pv) == 'CAP'}
    rr = next(r0 + i for i, r in enumerate(v) if t(r[3]) and up(r[3]) in mapeados and num(r[15] if len(r) > 15 else None)
              and 'SIMULA' not in up(r[17]) and TAG not in t(r[1]))
    an53 = t(cap.Range(f'D{rr}').Value)
    canon53 = next((cn for pv, nm, cn in mapa if up(pv) == 'CAP' and up(nm) == up(an53)), '')
    vinfo = {}
    for c in ('O', 'P', 'S', 'T', 'U', 'V', 'W', 'X'):
        try:
            vv = cap.Range(f'{c}{rr}').Validation
            vinfo[c] = [vv.Type, t(vv.Formula1)]
        except Exception as ex:
            vinfo[c] = ['sem', repr(ex)[:60]]
    z3a = t(cap.Range('Z3').Value)
    fP, vP = t(cap.Range(f'P{rr}').Formula), cap.Range(f'P{rr}').Value
    gA = {k: v_['G_bias'] for k, v_ in C.estat_tabela(qa).items() if k.startswith(canon53 + '|')}
    aud.novos()
    cap.Range(f'P{rr}').Value = 2.55
    qa.ex.esperar()
    qa.ex.xl.Calculate()
    z3b = t(cap.Range('Z3').Value)
    valid_apos = None
    try:
        valid_apos = bool(cap.Range(f'P{rr}').Validation.Value)
    except Exception:
        pass
    gB = {k: v_['G_bias'] for k, v_ in C.estat_tabela(qa).items() if k.startswith(canon53 + '|')}
    tr53 = aud.novos()
    cap.Range(f'P{rr}').Formula = fP
    qa.ex.xl.Calculate()
    z3c = t(cap.Range('Z3').Value)
    gC = {k: v_['G_bias'] for k, v_ in C.estat_tabela(qa).items() if k.startswith(canon53 + '|')}
    out['O53'] = {'linha': rr, 'analito_cap': an53, 'canonico': canon53, 'validacao': vinfo, 'P_antes': vP,
                  'formula_P': fP[:120], 'Z3': [z3a[:120], z3b[:160], z3c[:120]], 'validacao_aceita_2,55': valid_apos,
                  'G_antes': gA, 'G_sobrescrito': gB, 'G_restaurado': gC, 'trilha': [x['txt'][:300] for x in tr53]}
    out['O53_ok'] = (all(x[0] == 7 and x[1].replace(' ', '') == '=1=0' for x in vinfo.values())
                     and 'FÓRMULA SOBRESCRITA' in z3b and any('CEQ_FORMULA_SOBRESCRITA' in x['txt'] and f'P{rr}' in x['txt']
                                                             for x in tr53)
                     and z3c == z3a and gC == gA)
    # ---- OC-0054: edicao direta sem selecionar antes
    qa.ws('Painel').Activate()
    cap.Activate()
    cap.Range('AB1').Select()
    z4a = cap.Range('Z4').Value
    f0 = cap.Range(f'F{rr}').Value
    cap.Range(f'F{rr}').Value = f0
    qa.ex.esperar()
    t54a = aud.novos()
    z4b = cap.Range('Z4').Value
    z3m = t(cap.Range('Z3').Value)
    cap.Range('AB1').Select()
    cap.Range(f'F{rr}').Value = float(f0) + 1
    qa.ex.esperar()
    t54b = aud.novos()
    z3d = t(cap.Range('Z3').Value)
    antes_txt = [t(f0), t(f0).replace('.', ',')]
    out['O54'] = {'F_antes': f0, 'mesmo_valor_trilha': [x['txt'][:200] for x in t54a], 'Z4': [z4a, z4b],
                  'Z3_mesmo_valor': z3m[:120], 'diferente_trilha': [x['txt'][:400] for x in t54b], 'Z3_diferente': z3d[:120]}
    out['O54_ok'] = (not t54a and z4b == z4a and 'NÃO CONSOLIDADA' not in z3m.upper()
                     and len(t54b) == 1 and t54b[0]['acao'] == 'CEQ_ALTERADO'
                     and any(f'{a} →' in t54b[0]['txt'] for a in antes_txt) and '? →' not in t54b[0]['txt'])
    qa.ex.xl.EnableEvents = False
    cap.Range(f'F{rr}').Value = f0
    qa.ex.xl.EnableEvents = True
    return out


def s_fim_sem_sessao(qa, l_sem, aud):
    out = {}
    lo_cap = qa.lo('tblEQA_CAP_Dados')
    qa.run('mSeguranca.Logout')
    qa.ex.esperar()
    out['SessaoAtiva_apos_logout'] = qa.run('mSeguranca.SessaoAtiva')
    out['salvar'] = t(qa.run('mCEQForm.SalvarCEQ', C.campos_ceq('CAP', ROD, ANO, NOME_CEQ, 'TA-34', '110', unid=UN_CEQ)))
    if l_sem:
        antes = t(qa.ws('EQA.CAP_Dados').Range(f'F{l_sem}').Value)
        out['corrigir_sem_sessao'] = t(qa.run('mCEQForm.CorrigirCEQ', l_sem,
                                              C.campos_ceq('CAP', ROD, ANO, NOME_CEQ, 'TA-31', '115', alvo='', unid=UN_CEQ)))
        out['F_antes_depois'] = [antes, t(qa.ws('EQA.CAP_Dados').Range(f'F{l_sem}').Value)]
        out['excluir_sem_sessao'] = t(qa.run('mCEQForm.ExcluirCEQ', 'CAP', l_sem))
        out['A_depois'] = t(qa.ws('EQA.CAP_Dados').Range(f'A{l_sem}').Value)
    out['audit'] = [x['txt'][:260] for x in aud.novos()]
    out['abrir'] = abrir_ceq_sem_sessao(qa)
    return out


SECOES = sys.argv[sys.argv.index('--secoes') + 1].split(',') if '--secoes' in sys.argv else None


def quer(n):
    return SECOES is None or n in SECOES


def rodar(qa, L, senha, xl):
    p, e = qa.ws('Painel'), qa.ws('Estatística')
    p.Activate()
    qa.ex.esperar()
    for w in (p, e):
        try:
            w.Unprotect(senha)
        except Exception:
            pass
    fin = qa.ler('tblCQ_Final')
    L['linhas_tblCQ_Final'] = len(fin)
    if quer('estaticas'):
        secao(L, 'estaticas', s_estaticas, qa, senha)
    if quer('O35_guia'):
        secao(L, 'O35_guia', s_guia, qa)
    if quer('O33_sem_sessao_inicio'):
        secao(L, 'O33_sem_sessao_inicio', s_sem_sessao, qa)
    if quer('ADR091'):
        secao(L, 'ADR091', s_adr091, qa, fin)
    if quer('O12'):
        secao(L, 'O12', s_O12, qa, fin)
    if BIO and quer('O41'):
        secao(L, 'O41_G23_bonus', s_O41, qa, senha)
    # login (usuario de teste so na copia)
    criar_usuario(qa, USR, 'ANALISTA')
    L['login'] = fazer_login(qa, USR)
    for w in (p, e):
        try:
            w.Unprotect(senha)
        except Exception:
            pass
    aud = Audit(qa)
    if quer('manuais'):
        secao(L, 'manuais', s_manuais, qa, fin, aud)
    ceq = secao(L, 'ceq', s_ceq, qa, senha, aud) if quer('ceq') else {}
    if quer('fim'):
        secao(L, 'fim_sem_sessao', s_fim_sem_sessao, qa, (ceq or {}).get('linha_sem_alvo', 0), aud)


if __name__ == '__main__':
    t0 = time.time()
    main()
    print(f'tempo {time.time() - t0:.0f}s', flush=True)
