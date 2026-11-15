# -*- coding: utf-8 -*-
"""valida_reteste04.py -- Tester/QA, ciclo 2026-11-14_reteste04, modo VALIDAR (entrega 41, commit 1a72892).

Itens: Bio OC-0012, 0019, 0023, 0024, 0031, 0033, 0034, 0035, 0036, 0040, 0041, 0043, 0044, 0050, 0052, 0054, 0055,
0056, 0057, 0058, 0059, 0066, ADR-091, ADR-093, ADR-094; Hema OC-0023, 0024, 0031, 0033, 0034, 0035, 0036, 0041, 0049,
0050, 0052, 0054, 0055, 0057, 0058, 0066, ADR-091, ADR-094.

Escrito pelo Tester/QA. NAO reaproveita os scripts do Usuario Simulado (reteste04_usuario_*.py) nem as suites qa_adr09x.
Reaproveita so a infraestrutura: qa_final.QA (abre a copia, revela as abas, imprime EXCEL_PID), repro_comum e as classes
Audit/Vigia do meu valida_adr090_091.py (leitura da trilha e das caixas do MEU PID), auditoria_abas.largura_pt (medida
da fonte), qa_inativar_ux.criar_usuario/fazer_login (usuario de teste so na copia).
Cenarios VARIADOS em relacao ao reteste 03 e ao Usuario Simulado: analito do CEQ escolhido pelo mapa (nao Proteina
total/Glicose/RBC/HGB), Painel em DIMENSION 1 no ambar, conflito +3 min, oraculo do ADR-093 pelo de/para x recebimento
(um aceito alem da Creatinina e um recusado escolhidos pela varredura), regras de teste em lotes da Hema escolhidos pela
varredura, oraculo da janela de 180 corridas por RUN com media/DP/CV, oraculo do bias em duas etapas (k rodadas), e
varredura propria de TODAS as tabelas para o ADR-094.

Regras: so COPIA nova em %TEMP%/qcw/auditoria/2026-11-14_reteste04/tester/; nada e salvo; fechar() encerra so o PID
que o QA criou; nenhuma captura de tela. Dados de teste: usuario TESTE_AUDITORIA_R4V, MOTIVO "Outro: TESTE-AUDITORIA",
valores 222,222 / 211,111 / -1 / 50 / 12,345 ...; CEQ rodada "TESTE-AUDITORIA R4V 2025", amostras TV-*.

Uso: python -u valida_reteste04.py <Bioquimica|Hematologia> [--secoes a,b,c]
"""
import collections
import datetime as dt
import json
import os
import re
import statistics as stt
import sys
import time
import traceback

AQUI = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, AQUI)
import repro_comum as C  # noqa: E402
sys.path.insert(0, C.TESTES)
import auditoria_abas  # noqa: E402
from qa_inativar_ux import criar_usuario, fazer_login  # noqa: E402
from valida_adr090_091 import Audit, Vigia  # noqa: E402  (minhas, ciclo reteste03)

PROD = sys.argv[1]
BIO = PROD == 'Bioquimica'
NLV = 2 if BIO else 3
C.PASTA = os.path.join(os.environ['TEMP'], 'qcw', 'auditoria', '2026-11-14_reteste04', 'tester')
GUIA = os.path.join(C.REPO, '_arquitetura', 'integracao_dbseac', 'GUIA_DE_USO.md')
ADR090 = os.path.join(C.REPO, '_arquitetura', 'integracao_dbseac', 'ADR-090.md')
USR = 'TESTE_AUDITORIA_R4V'
MOTIVO = 'Outro: TESTE-AUDITORIA'
AMBAR = 255 + 235 * 256 + 156 * 65536
AMBAR_FONTE = 156 + 87 * 256
ROD, ANO = 'TESTE-AUDITORIA R4V 2025', '2025'
LOTE_ST = ('LOTE_FORA_DO_CQ', 'LOTE_OUTRO_MATERIAL')
PRECEDE_LOTE = ('INATIVADO', 'SEM_VALOR', 'SEM_DATA_HORA', 'SUBSTITUIDO_MANUAL', 'CONFLITO_MANUAL', 'DUPLICIDADE_ORIGEM',
                'MANUAL_INCOMPLETO', 'DUPLICIDADE')
SECOES = sys.argv[sys.argv.index('--secoes') + 1].split(',') if '--secoes' in sys.argv else None
GLOBAL = {}
SEGREDOS = [x for x in (os.environ.get('USERNAME'), os.environ.get('COMPUTERNAME')) if x]
try:
    import win32api
    SEGREDOS.append(win32api.GetUserNameEx(3))   # nome de exibicao do Windows (vai para o Audit_Log da copia)
except Exception:  # noqa: BLE001
    pass
SEGREDOS = [x for x in SEGREDOS if x and len(x) > 3]


def quer(n):
    return SECOES is None or n in SECOES


def t(v):
    if isinstance(v, float) and v == int(v):
        return str(int(v))
    return '' if v is None else str(v).strip()


def up(v):
    return t(v).upper()


def num(v):
    return isinstance(v, (int, float)) and not isinstance(v, bool)


def milhar(n):
    return f'{n:,}'.replace(',', '.')


def fmt2(x):
    return f'{x:,.2f}'.replace(',', '#').replace('.', ',').replace('#', '.')


def ser(d):
    """Data como serial do Excel (o pywin32 passa datetime ingenuo pelo fuso: OC-0066)."""
    return float((d - dt.datetime(1899, 12, 30)).days)


def rel(a, b):
    try:
        return abs(a - b) / max(abs(a), abs(b), 1e-300)
    except Exception:
        return None


class Log(C.Log):
    def gravar(self):
        s = json.dumps(self.R, ensure_ascii=False, indent=1, default=str)
        for x in SEGREDOS:
            s = re.sub(re.escape(x), '<omitido>', s, flags=re.I)
        open(self.saida, 'w', encoding='utf-8').write(s)


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


def estat_snapshot(qa):
    return qa.ws('Estatística').Range('A14:AU133').Value2


def diff_snap(a, b):
    d = []
    for i, (la, lb) in enumerate(zip(a, b)):
        for j, (x, y) in enumerate(zip(la, lb)):
            if x != y:
                d.append((14 + i, j + 1, x, y))
    return d


def estat_linhas(qa):
    est = qa.ws('Estatística')
    v = est.Range('A14:AU133').Value2
    out = {}
    for i, l in enumerate(v, 14):
        if l[0] and num(l[1]):
            k = f'{l[0]}|{int(l[1])}'
            out.setdefault(k, {'linha': i, 'n': l[2], 'media': l[3], 'dp': l[4], 'cv': l[5], 'G': l[6], 'AC': l[28]})
    return out


def mdcv(v):
    v = [float(x) for x in v]
    if len(v) < 2:
        return {'n': len(v)}
    m = stt.fmean(v)
    s = stt.stdev(v)
    return {'n': len(v), 'media': m, 'dp': s, 'cv': s / abs(m) * 100}


def dentro(d, de, ate):
    if not isinstance(d, dt.datetime):
        return False
    if isinstance(de, dt.datetime) and d.date() < de.date():
        return False
    if isinstance(ate, dt.datetime) and d.date() > ate.date():
        return False
    return True


def eleg(fin, an, lote, eq, de, ate):
    """Resultados elegiveis (PARTICIPA SIM, numerico) do analito/lote/equipamento/periodo -- oraculo proprio."""
    out = []
    for r in fin:
        if up(r['ANALITO']) != up(an) or r['PARTICIPA_ESTATISTICA'] != 'SIM' or not num(r['RESULTADO']):
            continue
        if lote is not None and t(r['LOTE']) != t(lote):
            continue
        if eq and up(r['EQUIPAMENTO']) != up(eq):
            continue
        if not dentro(r['DATA_HORA'], de, ate):
            continue
        out.append(r)
    return out


# ============================================================================================ estaticas
def s_estaticas(qa):
    p, e = qa.ws('Painel'), qa.ws('Estatística')
    out = {}
    linhas = [r for r in range(14, 134) if e.Cells(r, 1).HasFormula]
    # OC-0036: AT = "x,xx / y,yy" com as celulas de origem formatadas por mim
    f14 = t(e.Range('AT14').Formula)
    refs = re.findall(r'\$?([A-Z]{1,3})\$?14\b', f14)
    refs = [x for x in dict.fromkeys(refs) if x != 'A']
    txts, ruins, conf = [], [], 0
    for r in linhas:
        tx = t(e.Cells(r, 46).Text)
        if not tx:
            continue
        txts.append((r, tx))
        if re.search(r'(^|\s)0\d\d(\s|$)', tx):
            ruins.append((r, tx, 'zero-preenchido'))
        partes = tx.split(' / ')
        if len(partes) == 2 and len(refs) >= 2:
            vals = [e.Range(f'{c}{r}').Value2 for c in refs[:2]]
            if all(num(v) for v in vals):
                conf += 1
                esp = ' / '.join(fmt2(v) for v in vals)
                if esp != tx:
                    ruins.append((r, tx, esp))
    padrao = re.compile(r'^-?\d{1,3}(\.\d{3})*,\d{2} / -?\d{1,3}(\.\d{3})*,\d{2}$')
    fora_padrao = [(r, x) for r, x in txts if not padrao.match(x)]
    out['O36'] = {'formula_AT14': f14[:200], 'refs': refs[:4], 'n_textos': len(txts), 'conferidos': conf,
                  'exemplos': txts[:4], 'ruins': ruins[:6], 'fora_do_padrao': fora_padrao[:8],
                  'todos_textos': sorted({x for _, x in txts})[:90],
                  'ok': bool(txts) and not ruins and not fora_padrao}
    # OC-0050: Principal A5 x contagem propria
    pr = qa.ws('Principal - Resultados')
    out['O50'] = {'A5': t(pr.Range('A5').Text)[:220]}
    # OC-0044 (Bio)
    if BIO:
        fora = [r for r in linhas if e.Range(e.Cells(r, 1), e.Cells(r, 2)).VerticalAlignment != -4108]
        altas = [(r, t(e.Cells(r, 1).Value), float(e.Rows(r).RowHeight)) for r in linhas if e.Rows(r).RowHeight > 20][:6]
        out['O44'] = {'linhas_analito': len(linhas), 'fora_do_centro': fora[:10], 'linhas_altas': altas,
                      'ok': bool(linhas) and not fora}
    # OC-0049 (Hema): linhas sem analito sem texto de A a AU
    if not BIO:
        vaz, sujas = [], []
        for r in range(14, 134):
            if t(e.Cells(r, 1).Text) == '':
                vaz.append(r)
                tx = e.Range(e.Cells(r, 1), e.Cells(r, 47))
                vals = [t(tx.Cells(1, c).Text) for c in range(1, 48)]
                if any(vals):
                    sujas.append((r, [(c + 1, v) for c, v in enumerate(vals) if v][:5]))
        com = [r for r in range(14, 134) if t(e.Cells(r, 1).Text)]
        niveis_ok = all(t(e.Cells(r, 2).Text) in ('1', '2', '3') for r in com)
        out['O49'] = {'linhas_vazias': [vaz[:1], vaz[-1:], len(vaz)], 'sujas': sujas[:5], 'n_sujas': len(sujas),
                      'linhas_com_analito': len(com), 'niveis_1_2_3': niveis_ok, 'B98_formula': t(e.Range('B98').Formula),
                      'ok': bool(vaz) and not sujas and niveis_ok}
    # OC-0043 (Bio): todos os textos que a formula do J7 pode mostrar, medidos na fonte real
    if BIO:
        j = {}
        f7 = t(p.Range('J7').Formula)
        literais = sorted({x.replace('""', '"') for x in re.findall(r'"((?:[^"]|"")*)"', f7)} - {'', ','})
        estados = [x for x in literais if re.match(r'^[A-ZÀ-Ú0-9ª ./]{2,}$', x)] + ['REJEITADO', 'ALERTA', 'OK']
        for c in [f'J{7 + k}' for k in range(NLV)]:
            cel = p.Range(c)
            larg = float(cel.MergeArea.Width) - 3.0
            f = cel.DisplayFormat.Font
            # conservador: mede em NEGRITO (a formatacao condicional pode por negrito conforme o estado)
            med = {s: round(min(1.0, larg / auditoria_abas.largura_pt(s, f.Name, True, float(f.Size))), 3)
                   for s in estados}
            j[c] = {'shrink': cel.ShrinkToFit, 'wrap': cel.WrapText, 'largura_util_pt': round(larg, 1),
                    'fonte': [f.Name, bool(f.Bold), float(f.Size)], 'fator': med}
        q4f = t(p.Range('Q4').Formula)
        pior = min(v for c in j.values() for v in c['fator'].values())
        out['O43'] = {'estados_da_formula': estados, 'celulas': j, 'pior_fator': pior,
                      'Q4_cobre': {s: (f'"{s}"' in q4f) for s in literais},
                      'ok': all(c['shrink'] and not c['wrap'] for c in j.values()) and pior >= 0.70}
    # OC-0040 (Bio): cabecalho da Estatistica
    if BIO:
        cab = {}
        for rr in range(2, 7):
            for cc in range(1, 12):
                cel = e.Cells(rr, cc)
                if t(cel.Formula) or t(cel.Text):
                    cab[cel.Address.replace('$', '')] = [t(cel.Formula)[:120], t(cel.Text)[:100], bool(cel.Locked)]
        out['O40_cabecalho'] = cab
    return out


def s_docs():
    g = open(GUIA, encoding='utf-8').read()
    i3, i4 = g.find('## 3. '), g.find('## 4. ')
    i5, i6 = g.find('## 5. '), g.find('## 6. ')
    s3, s56 = g[i3:i4], g[i5:g.find('## 7. ')]
    termos = ['Registrar resultado do CEQ', 'login', 'valor-alvo', 'Corrigir', 'Excluir', 'Consolidar', 'Z3',
              'ATUALIZAR', 'FÓRMULA SOBRESCRITA']
    manda = [m.group(0) for m in re.finditer(r'[^.\n]*digit[ea][^.\n]*(nas abas|na aba) EQA[^.\n]*', g)]
    nao_digite = 'não digite' in s56.lower()
    a = open(ADR090, encoding='utf-8').read().split('\n')
    oc40 = [(i + 1, l.strip()[:200]) for i, l in enumerate(a) if 'OC-0040' in l]
    oc41 = [(i + 1, l.strip()[:200]) for i, l in enumerate(a) if 'OC-0041' in l]
    g23 = [(i + 1, l.strip()[:200]) for i, l in enumerate(a) if 'G23' in l]
    oc = open(os.path.join(C.REPO, '_arquitetura', 'agentes', 'OCORRENCIAS.md'), encoding='utf-8').read()
    reg40 = re.search(r'^\| OC-0040 \|[^|]*\|[^|]*\|([^|]*)\|', oc, re.M)
    reg41 = re.search(r'^\| OC-0041 \|[^|]*\|[^|]*\|([^|]*)\|', oc, re.M)
    # G23 nunca pode aparecer junto de "OC-0040" sem a ressalva de que o nome antigo foi trocado
    g23_mal = [x for x in g23 if 'OC-0040' in x[1] and 'OC-0041' not in x[1]]
    return {'O35': {'s3_equipamento': 'EQUIPAMENTO' in s3.upper() and 'âmbar' in s3,
                    'termos_s5_6': {k: k in s56 for k in termos}, 'manda_digitar_nas_abas': manda[:3],
                    'nao_digite_nelas': nao_digite,
                    'ok': 'EQUIPAMENTO' in s3.upper() and all(k in s56 for k in termos) and not manda and nao_digite},
            'O59': {'linhas_OC-0040': oc40, 'linhas_OC-0041': oc41, 'linhas_G23': g23, 'G23_rotulado_so_OC-0040': g23_mal,
                    'registro_OC-0040': t(reg40.group(1)) if reg40 else None,
                    'registro_OC-0041': t(reg41.group(1)) if reg41 else None,
                    'ok': bool(oc41) and not g23_mal and reg41 is not None and 'K3' in reg41.group(1)}}


def s_botoes(qa):
    capt = []
    for ws in qa.wb.Worksheets:
        try:
            for sh in ws.Shapes:
                try:
                    tx = t(sh.TextFrame2.TextRange.Text)
                except Exception:
                    tx = ''
                if tx:
                    capt.append((ws.Name, tx[:60], t(getattr(sh, 'OnAction', ''))[:60]))
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
    tudo = ' || '.join(x for _, x, _ in capt).upper() + ' || ' + ' || '.join(frm).upper()
    alvo = ['Registrar resultado do CEQ', 'Consolidar', 'ATUALIZAR', 'Salvar', 'Corrigir', 'Excluir']
    return {'existem': {k: k.upper() in tudo for k in alvo},
            'onde': {k: sorted({w for w, x, _ in capt if k.upper() in x.upper()})[:6] for k in alvo},
            'frm': [x for x in frm if x][:30]}


# ============================================================================================ sessao / CEQ
def campos(rod, an, am, res, alvo='100', unid=''):
    return C.campos_ceq('CAP', rod, ANO, an, am, res, alvo=alvo, unid=unid)


def s_sem_sessao(qa, nome_ceq, unid):
    lo = qa.lo('tblEQA_CAP_Dados')
    u0 = sum(1 for r in lo.ListColumns(4).DataBodyRange.Value if r[0] not in (None, ''))
    aud = Audit(qa)
    sess = qa.run('mSeguranca.SessaoAtiva')
    r = t(qa.run('mCEQForm.SalvarCEQ', campos(ROD, nome_ceq, 'TV-00', '110', unid=unid)))
    u1 = sum(1 for x in lo.ListColumns(4).DataBodyRange.Value if x[0] not in (None, ''))
    ab = abrir_ceq(qa)
    tr = aud.novos()
    return {'SessaoAtiva': sess, 'salvar': r, 'linhas_usadas': [u0, u1], 'trilha': [x['txt'][:200] for x in tr],
            'abrir': ab, 'ok': r.startswith('ERRO|0|') and 'login' in r.lower() and u0 == u1 and not tr and ab['ok']}


def abrir_ceq(qa):
    vg = Vigia(qa.ex.pid, teto=40)
    vg.start()
    erro = None
    try:
        qa.ex.run("'" + qa.wb.Name + "'!mCEQForm.AbrirCEQ", teto=60)
    except Exception as ex:  # noqa: BLE001
        erro = repr(ex)
    time.sleep(1.5)
    vg.parar.set()
    vg.join(3)
    msg = ' '.join(' '.join(c['textos']) for c in vg.vistos)
    return {'caixas': vg.vistos, 'erro_run': erro,
            'ok': any(c['classe'] == '#32770' for c in vg.vistos) and 'login' in msg.lower()
            and not any(c['classe'] == 'ThunderDFrame' for c in vg.vistos)}


def escolher_ceq(qa):
    """Analito do CEQ: um nome do CAP mapeado cujo canonico tem bias numerico na Estatistica, fora dos ja usados."""
    base = qa.ws('EQA_Base')
    mapa = [(t(m[0]), t(m[1]), t(m[2])) for m in C.bloco(base, 'W2:Y600') if t(m[1])]
    est = estat_linhas(qa)
    usados = {'PROTEÍNA TOTAL', 'GLICOSE', 'RBC', 'HGB', 'ALT (SGPT)', 'BASO%', 'WBC'}
    lo = qa.lo('tblEQA_CAP_Dados')
    cab = [t(c.Name) for c in lo.ListColumns]
    v = lo.DataBodyRange.Value
    iu = next((i for i, c in enumerate(cab) if 'UNID' in c.upper()), None)
    cands = []
    for pv, nm, cn in mapa:
        if up(pv) != 'CAP' or up(cn) in usados:
            continue
        k = f'{cn}|1'
        if k in est and num(est[k]['G']):
            reais = [r for r in v if up(r[3]) == up(nm)]
            if reais:
                cands.append((cn, nm, t(reais[0][iu]) if iu is not None else ''))
    cands.sort()
    return cands[len(cands) // 2] if cands else None, len(cands), mapa


def rodadas_base(qa, canon):
    """Rodadas distintas (ano 2025, CAP) do canonico na EQA_Base com |bias| numerico -- conferencia do k."""
    b = qa.ws('EQA_Base')
    v = C.bloco(b, 'A1:U6001')
    cab = [up(c) for c in v[0]]

    def col(*nomes):
        for n in nomes:
            for i, c in enumerate(cab):
                if n in c:
                    return i
        return None
    ia, ip, ir, iy, ib = (col('ANALITO_CANONICO', 'CANONICO'), col('PROVEDOR'), col('RODADA'), col('ANO'),
                          col('|BIAS|', 'BIAS_ABS', 'ABS'))
    rod = set()
    for l in v[1:]:
        if ia is None or up(l[ia]) != up(canon):
            continue
        if ip is not None and up(l[ip]) != 'CAP':
            continue
        if iy is not None and t(l[iy]) != ANO:
            continue
        if ib is not None and not num(l[ib]):
            continue
        rod.add(t(l[ir]) if ir is not None else '?')
    return {'cabecalho': [t(c) for c in v[0]][:21], 'k_base': len(rod), 'rodadas': sorted(rod)[:12]}


def s_ceq(qa, senha, aud, esc):
    out = {}
    canon, nome, unid = esc
    out['escolhido'] = {'canonico': canon, 'nome_cap': nome, 'unidade': unid}
    # OC-0034
    o34 = {}
    for prov, tb in (('CAP', 'tblEQA_CAP_Dados'), ('Controllab', 'tblEQA_Controllab_Dados')):
        try:
            lo = qa.lo(tb)
        except Exception:
            continue
        if lo.ListRows.Count == 0:
            o34[prov] = 'vazia'
            continue
        v = lo.DataBodyRange.Value
        res = {}
        for col, rot in ((2, 'rodada'), (4, 'analito'), (12, 'avaliacao')):
            reais = {up(r[col - 1]) for r in v if t(r[col - 1]) and 'SIMULA' not in up(r[17])}
            simu = {up(r[col - 1]) for r in v if t(r[col - 1]) and 'SIMULA' in up(r[17])} - reais
            lista = {up(x) for x in (qa.run('mCEQForm.DistintosCEQ', prov, col) or ())}
            res[rot] = {'n_lista': len(lista), 'n_reais': len(reais), 'igual': lista == reais,
                        'simulacao_na_lista': sorted(simu & lista)[:6], 'n_so_simulacao': len(simu)}
        o34[prov] = res
    out['O34'] = o34
    out['O34_ok'] = all(x['igual'] and not x['simulacao_na_lista'] for p in o34.values() if isinstance(p, dict)
                        for x in p.values())
    # OC-0052: pares por prefixo literal no mapa real; na Bio (sem par real) um par sintetico na copia
    base = qa.ws('EQA_Base')
    mapa = [(t(m[0]), t(m[1]), t(m[2])) for m in C.bloco(base, 'W2:Y600') if t(m[1])]
    nomes = {}
    for pv, nm, cn in mapa:
        if up(pv) == 'CAP':
            nomes.setdefault(nm, cn)
    pares = []
    for c, cc in nomes.items():
        lit = [l for l, cl in nomes.items() if up(l).startswith(up(c)) and len(l) > len(c) and up(cl) != up(cc)]
        if lit:
            av = t(qa.run('mCEQForm.AvisoAnalitoCEQ', 'CAP', c))
            pares.append({'nome': c, 'par': lit[:2], 'aviso': av[-160:], 'cita': any(f'“{l}”' in av for l in lit)})
    o52 = {'mapa_CAP': len(nomes), 'pares_reais': pares}
    if not any(p for p in pares if not any(up(x).startswith(up(p['nome']) + ' ') for x in p['par'])):
        # sintetico: "TESTE-AUDITORIA R4V" (-> canonico A) e "TESTE-AUDITORIA R4VX" (-> canonico B), apagado depois
        livre = next(i for i in range(2, 600) if not t(base.Range(f'X{i}').Value))
        dois = [cn for cn in dict.fromkeys(cn for _, _, cn in mapa)][:2]
        ev = qa.ex.xl.EnableEvents
        qa.ex.xl.EnableEvents = False
        try:
            base.Unprotect(senha)
        except Exception:
            pass
        base.Range(f'W{livre}:Y{livre}').Value = (('CAP', 'TESTE-AUDITORIA R4V', dois[0]),)
        base.Range(f'W{livre + 1}:Y{livre + 1}').Value = (('CAP', 'TESTE-AUDITORIA R4VX', dois[1]),)
        qa.ex.xl.EnableEvents = ev
        a1 = t(qa.run('mCEQForm.AvisoAnalitoCEQ', 'CAP', 'TESTE-AUDITORIA R4V'))
        a2 = t(qa.run('mCEQForm.AvisoAnalitoCEQ', 'CAP', 'TESTE-AUDITORIA R4VX'))
        qa.ex.xl.EnableEvents = False
        base.Range(f'W{livre}:Y{livre + 1}').ClearContents()
        qa.ex.xl.EnableEvents = ev
        o52['sintetico'] = {'linhas': [livre, livre + 1], 'canonicos': dois, 'aviso_menor': a1, 'aviso_maior': a2,
                            'ok': '“TESTE-AUDITORIA R4VX”' in a1 and 'existe também' not in a2}
    o52['par_palavra'] = t(qa.run('mCEQForm.AvisoAnalitoCEQ', 'CAP', 'Lactate' if BIO else 'Lymphs'))[-160:]
    out['O52'] = o52
    out['O52_ok'] = all(p['cita'] for p in pares) and (o52.get('sintetico', {}).get('ok', True)) and \
        'existe também' in o52['par_palavra'] and (BIO or any(p['nome'] == 'MCH' for p in pares))
    # OC-0031: sem alvo -> G igual bit a bit; com alvo -> G = (k*G0 + b)/(k+1), k inteiro = rodadas da base
    k0 = f'{canon}|1'
    s0 = estat_snapshot(qa)
    rb = rodadas_base(qa, canon)
    g0 = {k: v['G'] for k, v in estat_linhas(qa).items() if k.startswith(canon + '|')}
    c_sem = campos(ROD, nome, 'TV-01', '110', alvo='', unid=unid)
    c_com = campos(ROD, nome, 'TV-02', '110', alvo='100', unid=unid)
    conf_sem = t(qa.run('mCEQForm.ConfirmacoesCEQ', c_sem, 0))
    conf_com = t(qa.run('mCEQForm.ConfirmacoesCEQ', c_com, 0))
    r_sem = t(qa.run('mCEQForm.SalvarCEQ', c_sem))
    s1 = estat_snapshot(qa)
    a_sem = aud.novos()
    r_com = t(qa.run('mCEQForm.SalvarCEQ', c_com))
    g2 = {k: v['G'] for k, v in estat_linhas(qa).items() if k.startswith(canon + '|')}
    a_com = aud.novos()
    G0, G2 = g0[k0], g2[k0]
    kf = (10.0 - G2) / (G2 - G0) if G2 != G0 else None
    k_base = rb['k_base']
    esp2 = (k_base * G0 + 10.0) / (k_base + 1) if k_base else None
    l_sem = int(r_sem.split('|')[1]) if r_sem.startswith('OK|') else 0
    l_com = int(r_com.split('|')[1]) if r_com.startswith('OK|') else 0
    des = qa.wb.VBProject.VBComponents('frmCEQ').Designer
    lbl = {n: t(des.Controls(n).Caption) for n in ('lblAlvo', 'lblSub')}
    out['O31'] = {'G0': G0, 'diff_sem_alvo': diff_snap(s0, s1)[:5], 'confirmacao_sem_alvo': conf_sem,
                  'confirmacao_com_alvo': conf_com, 'salvar_sem': r_sem, 'salvar_com': r_com, 'G_com_alvo': G2,
                  'k_pela_formula': kf, 'rodadas_base': rb, 'G_esperado_k_base': esp2, 'rel': rel(esp2, G2) if esp2 else None,
                  'mesmo_G_em_todos_os_niveis': len({round(x, 12) for x in g2.values() if num(x)}) == 1, 'rotulos': lbl,
                  'trilha': [x['acao'] for x in a_sem + a_com]}
    k_int = round(kf) if kf is not None and abs(kf - round(kf)) < 1e-6 else None
    out['O31']['k_inteiro'] = k_int
    out['O31_ok'] = ('NÃO entra no bias' in conf_sem and conf_com == '' and r_sem.startswith('OK|')
                     and 'sem valor-alvo' in r_sem and not [x for x in diff_snap(s0, s1) if x[1] == 7]  # G (bias)
                     and k_int is not None
                     and k_int == k_base and 'Valor-alvo' in lbl['lblAlvo'])
    if k_int and k_int != k_base:
        k_base = k_int
    # OC-0055: descricao, lista da rodada, trilha, duplicidade
    desc = t(qa.run('mCEQForm.DescricaoCEQ', 'CAP', l_com)) if l_com else ''
    regs = qa.run('mCEQForm.RegistrosDaRodada', 'CAP', ROD, ANO)
    regs = [[t(x) for x in l] for l in regs] if isinstance(regs, tuple) else regs
    dup = t(qa.run('mCEQForm.SalvarCEQ', c_com))
    a_dup = aud.novos()
    tx55 = [desc, dup] + [' '.join(x) for x in (regs or [])] + [x['txt'] for x in a_com]
    out['O55'] = {'descricao': desc, 'registros': regs, 'duplicidade': dup, 'trilha_dup': [x['acao'] for x in a_dup],
                  'repete_ano': [x[:160] for x in tx55 if '2025 2025' in x]}
    out['O55_ok'] = bool(desc) and ROD in desc and not out['O55']['repete_ano'] and dup.startswith('ERRO|') \
        and ROD in dup
    # OC-0057: sem sessao, Corrigir e Excluir recusados; com sessao, corrigir muda G pelo oraculo
    lo_cap = qa.lo('tblEQA_CAP_Dados')
    cap = qa.ws('EQA.CAP_Dados')
    n0 = lo_cap.ListRows.Count
    qa.run('mSeguranca.Logout')
    qa.ex.esperar()
    sa = qa.run('mSeguranca.SessaoAtiva')
    fA = t(cap.Range(f'F{l_com}').Value)
    rc = t(qa.run('mCEQForm.CorrigirCEQ', l_com, campos(ROD, nome, 'TV-02', '105', unid=unid)))
    fB = t(cap.Range(f'F{l_com}').Value)
    rx = t(qa.run('mCEQForm.ExcluirCEQ', 'CAP', l_com))
    aB = t(cap.Range(f'A{l_com}').Value)
    tr57 = aud.novos()
    gS = {k: v['G'] for k, v in estat_linhas(qa).items() if k.startswith(canon + '|')}
    out['O57'] = {'SessaoAtiva': sa, 'corrigir': rc, 'F': [fA, fB], 'excluir': rx, 'A_depois': aB,
                  'linhas': [n0, lo_cap.ListRows.Count], 'trilha': [x['txt'][:200] for x in tr57], 'G_igual': gS == g2}
    ok57a = (rc.startswith('ERRO|0|') and rx.startswith('ERRO|0|') and fA == fB and aB == 'CAP'
             and not [x for x in tr57 if (x['acao'] or '').startswith(('CEQ_', 'EQA_'))]
             and gS == g2 and sa is False)
    fazer_login(qa, USR)
    rc2 = t(qa.run('mCEQForm.CorrigirCEQ', l_com, campos(ROD, nome, 'TV-02', '105', unid=unid)))
    g3 = {k: v['G'] for k, v in estat_linhas(qa).items() if k.startswith(canon + '|')}
    esp3 = (k_base * G0 + 5.0) / (k_base + 1) if k_base else None
    tr57b = aud.novos()
    out['O57']['com_sessao'] = {'corrigir': rc2, 'G': g3[k0], 'G_esperado': esp3, 'rel': rel(esp3, g3[k0]) if esp3 else None,
                                'trilha': [x['txt'][:220] for x in tr57b]}
    # limpeza: excluir os dois registros de teste
    rx2 = t(qa.run('mCEQForm.ExcluirCEQ', 'CAP', l_com))
    rx3 = t(qa.run('mCEQForm.ExcluirCEQ', 'CAP', l_sem)) if l_sem else ''
    tr_exc = aud.novos()
    s4 = estat_snapshot(qa)
    d4 = diff_snap(s0, s4)
    out['O57']['limpeza'] = {'excluir': [rx2, rx3], 'trilha': [x['acao'] for x in tr_exc],
                             'estatistica_volta_bit_a_bit': not d4, 'diff': d4[:5]}
    out['O57_ok'] = (ok57a and rc2.startswith('OK|') and esp3 is not None and rel(esp3, g3[k0]) <= 1e-9
                     and rx2.startswith('OK|') and not d4)
    out['linha_sem_alvo'] = l_sem
    return out


def s_O54(qa, senha, aud):
    cap = qa.ws('EQA.CAP_Dados')
    lo = qa.lo('tblEQA_CAP_Dados')
    v = lo.DataBodyRange.Value
    r0 = lo.DataBodyRange.Row
    # uma linha real diferente das usadas antes (linha 2): a 6a linha real com Seu Resultado numerico
    reais = [r0 + i for i, r in enumerate(v) if t(r[3]) and num(r[5]) and 'SIMULA' not in up(r[17])
             and 'TESTE' not in up(r[1])]
    rr = reais[5]
    cap.Activate()
    try:
        cap.Unprotect(senha)
    except Exception:
        pass
    cap.Range('AB1').Select()
    f0 = cap.Range(f'F{rr}').Value
    z3a = t(cap.Range('Z3').Value)
    aud.novos()
    cap.Range(f'F{rr}').Value = f0
    qa.ex.esperar()
    t1 = aud.novos()
    z3b = t(cap.Range('Z3').Value)
    cap.Range('AB1').Select()
    cap.Range(f'F{rr}').Value = float(f0) + 2
    qa.ex.esperar()
    t2 = aud.novos()
    z3c = t(cap.Range('Z3').Value)
    qa.ex.xl.EnableEvents = False
    cap.Range(f'F{rr}').Value = f0
    qa.ex.xl.EnableEvents = True
    antes = {t(f0), t(f0).replace('.', ',')}
    return {'linha': rr, 'analito': t(cap.Range(f'D{rr}').Value), 'F': f0, 'mesmo_valor_trilha': [x['txt'][:200] for x in t1],
            'Z3': [z3a[:90], z3b[:90], z3c[:120]], 'diferente_trilha': [x['txt'][:300] for x in t2],
            'ok': (not t1 and 'NÃO CONSOLIDADA' not in z3b.upper() and len(t2) == 1 and t2[0]['acao'] == 'CEQ_ALTERADO'
                   and any(f'{a} →' in t2[0]['txt'] for a in antes) and '? →' not in t2[0]['txt']
                   and 'NÃO CONSOLIDADA' in z3c.upper())}


# ============================================================================================ Painel
def painel_janela(qa, seq, fin, limpar_h3=False):
    p, eng = qa.ws('Painel'), qa.ws('Eng_Saida')
    for c, v in seq:
        C.por(qa, p, c, v, local=False)
    if limpar_h3:
        p.Range('H3').MergeArea.ClearContents()
        qa.ex.esperar()
        time.sleep(1)
    m1 = int(eng.Range('M1').Value or 0)
    af1 = t(eng.Range('AF1').Value)
    a39 = t(p.Range('A39').Value)
    ps = qa.painel_stats()
    de, ate = qa.nome('filtroDe').Value, qa.nome('filtroAte').Value
    de = C.limpo(de) if isinstance(de, dt.datetime) else None
    ate = C.limpo(ate) if isinstance(ate, dt.datetime) else None
    an = t(qa.nome('selAnalito').Value)
    lote = t(eng.Range('E1').Value).lstrip("'")
    eq = t(eng.Range('S1').Value) if BIO else None
    el = eleg(fin, an, lote, eq, de, ate)
    cont = collections.Counter(int(r['NIVEL'] or 0) for r in el)
    indep = ' · '.join(f'N{k} {cont.get(k, 0)}' for k in range(1, NLV + 1))
    runs = sorted({int(r['RUN']) for r in el if num(r['RUN'])}, reverse=True)
    # corridas so com X vermelho (inativado plotado) no periodo tambem ocupam lugar? (registro, nao criterio)
    runs_x = sorted({int(r['RUN']) for r in fin if up(r['ANALITO']) == up(an) and t(r['LOTE']) == lote
                     and (not eq or up(r['EQUIPAMENTO']) == up(eq)) and num(r['RUN']) and dentro(r['DATA_HORA'], de, ate)
                     and (r['PARTICIPA_ESTATISTICA'] == 'SIM' or 'X' in up(r['TIPO_PLOTAGEM_LJ']))}, reverse=True)
    jan = set(runs[:180])
    por_nivel = []
    for k in range(1, NLV + 1):
        vals = [r['RESULTADO'] for r in el if int(r['NIVEL'] or 0) == k and int(r['RUN']) in jan]
        o = mdcv(vals)
        sist = ps[k - 1]
        por_nivel.append({'nivel': k, 'meu': o, 'painel': list(sist[1:5]),
                          'rel_media': rel(o.get('media'), sist[2]) if num(sist[2]) else None,
                          'rel_dp': rel(o.get('dp'), sist[3]) if num(sist[3]) else None})
    ok_ind = all(x['meu']['n'] == x['painel'][0] and (x['rel_media'] or 0) <= 1e-9 and (x['rel_dp'] or 0) <= 1e-9
                 for x in por_nivel)
    return {'filtro': {'analito': an, 'lote': lote, 'equip': eq, 'de': str(de), 'ate': str(ate), 'K4': t(p.Range('K4').Value)},
            'M1': m1, 'AF1': af1, 'AF1_meu': indep, 'corridas_elegiveis': len(runs), 'corridas_com_X': len(runs_x),
            'cortadas_meu': max(0, len(runs) - 180), 'janela': por_nivel, 'A39': a39[-320:],
            'ok': (m1 == max(0, len(runs) - 180) and af1 == indep and 'indicadores do Painel' in a39
                   and f'Período inteiro: {af1}' in a39 and ok_ind) if m1 else None}


def s_O12(qa, fin):
    out = {'original': painel_janela(qa, [('L4', 'DIMENSION 1'), ('C3', 'Glicose'), ('H3', "'8974"),
                                          ('K4', 'Tudo (sem filtro)')], fin)}
    g = collections.Counter((t(r['ANALITO']), t(r['LOTE']), t(r['EQUIPAMENTO'])) for r in fin
                            if r['PARTICIPA_ESTATISTICA'] == 'SIM' and int(r['NIVEL'] or 0) == 1)
    # ampliacao no DIMENSION 2 (o reteste 03 e o Usuario usaram so o DIMENSION 1)
    cand = [(k, n) for k, n in g.most_common(80) if n > 230 and k[2] == 'DIMENSION 2' and k[0] not in ('Glicose', 'Cloro')]
    if cand:
        (an, lt, eq), n = cand[0]
        out['ampliacao_DIM2'] = painel_janela(qa, [('L4', eq), ('C3', an), ('H3', "'" + lt), ('K4', 'Tudo (sem filtro)')], fin)
        out['ampliacao_DIM2']['escolhido'] = [an, lt, eq, n]
    # vizinho: um mes (sem corte)
    out['vizinho_mes'] = painel_janela(qa, [('L4', 'DIMENSION 1'), ('C3', 'Glicose'), ('H3', "'8974"), ('K3', 2025),
                                            ('K4', 'Nov')], fin)
    viz = out['vizinho_mes']
    viz['ok'] = viz['M1'] == 0 and viz['AF1_meu'].startswith('N1') and 'fora do gráfico' not in viz['A39'] and \
        all(x['meu']['n'] == (x['painel'][0] or 0) for x in viz['janela'])
    out['ok'] = all(v.get('ok') for v in out.values() if isinstance(v, dict))
    return out


def s_O41(qa, senha):
    p = qa.ws('Painel')
    prot0 = bool(p.ProtectContents)
    if not prot0:
        p.Protect(senha, True, True)
    passos = []
    try:
        for c, v in (('K3', 2026), ('K4', 'T1')):
            p.Range(c).Value = v
            qa.ex.esperar()
        passos.append(['K3 2026 + T1', t(p.Range('K3').Value), t(p.Range('K4').Value), str(p.Range('G3').Value)[:10],
                       str(p.Range('G4').Value)[:10]])
        p.Range('G3').Value = ser(dt.datetime(2026, 8, 15))
        qa.ex.esperar()
        p.Range('G4').Value = ser(dt.datetime(2026, 10, 20))
        qa.ex.esperar()
        passos.append(['De/Até 15/08–20/10/2026', t(p.Range('K3').Value), t(p.Range('K4').Value)])
        p.Range('K4').Value = 'Set'
        qa.ex.esperar()
        passos.append(['K4 Set', t(p.Range('K3').Value), t(p.Range('K4').Value), str(p.Range('G3').Value)[:10],
                       str(p.Range('G4').Value)[:10]])
        p.Range('G3').Value = ser(dt.datetime(2025, 11, 10))
        qa.ex.esperar()
        p.Range('G4').Value = ser(dt.datetime(2026, 2, 5))
        qa.ex.esperar()
        e = qa.ws('Estatística')
        ano0 = [c.Address for c in e.Range('A1:AU13') if 'Ano 0' in t(c.Text)]
        passos.append(['De/Até 10/11/2025–05/02/2026', t(p.Range('K3').Value), t(p.Range('K4').Value), ano0])
    finally:
        if not prot0:
            p.Unprotect(senha)
    ok = (passos[0][3:] == ['2026-01-01', '2026-03-31'] and passos[1][2] == 'Personalizado'
          and passos[2][3:] == ['2026-09-01', '2026-09-30'] and passos[3][2] == 'Personalizado' and passos[3][1] == ''
          and not passos[3][3])
    return {'Painel_protegido_ao_abrir': prot0, 'passos': passos, 'ok': ok}


def s_O40(qa, fin):
    """Bio: a Estatistica segue o Painel; mes e trimestre alcancaveis; nunca 'Ano 0'; n/media/DP/CV = oraculo."""
    p, e = qa.ws('Painel'), qa.ws('Estatística')
    est_txt = lambda: {c: t(e.Range(c).Text)[:80] for c in ('A4', 'B4', 'D4', 'B5', 'D5', 'H5', 'J5')}  # noqa: E731
    out = {'D4_formula': t(e.Range('D4').Formula), 'B4_travada': bool(e.Range('B4').Locked),
           'protegida': bool(e.ProtectContents)}
    estados = []
    for rot, seq in (('Ago/2026', [('L4', 'DIMENSION 1'), ('C3', 'Glicose'), ('H3', "'8977"), ('K3', 2026), ('K4', 'Ago')]),
                     ('T2/2026', [('K4', 'T2')]),
                     ('Ano 2026', [('K4', 'Ano inteiro')]),
                     ('De/Até 20/12/2025–10/01/2026', [('G3', ser(dt.datetime(2025, 12, 20))), ('G4', ser(dt.datetime(2026, 1, 10)))]),
                     ('Tudo', [('K4', 'Tudo (sem filtro)')])):
        for c, v in seq:
            C.por(qa, p, c, v, local=False)
        time.sleep(0.5)
        de, ate = qa.nome('filtroDe').Value, qa.nome('filtroAte').Value
        de = C.limpo(de) if isinstance(de, dt.datetime) else None
        ate = C.limpo(ate) if isinstance(ate, dt.datetime) else None
        el = eleg(fin, 'Glicose', '8977', 'DIMENSION 1', de, ate)
        el_l = estat_linhas(qa)
        niv = []
        for k in (1, 2):
            o = mdcv([r['RESULTADO'] for r in el if int(r['NIVEL'] or 0) == k])
            s = el_l.get(f'Glicose|{k}', {})
            niv.append({'nivel': k, 'meu': o, 'estat': [s.get('n'), s.get('media'), s.get('dp'), s.get('cv')],
                        'ok': o['n'] == s.get('n') and (o['n'] < 2 or (rel(o['media'], s['media']) <= 1e-9
                                                                       and rel(o['dp'], s['dp']) <= 1e-9
                                                                       and rel(o['cv'], s['cv']) <= 1e-9))})
        ano0 = [c.Address.replace('$', '') for c in e.Range('A1:AU13') if re.search(r'\bAno 0\b', t(c.Text))]
        estados.append({'estado': rot, 'K3': t(p.Range('K3').Value), 'K4': t(p.Range('K4').Value), 'de': str(de)[:10],
                        'ate': str(ate)[:10], 'cab': est_txt(), 'Ano_0': ano0, 'niveis': niv})
    out['estados'] = estados
    try:
        e.Range('B4').Value = 'MENSAL'
        out['B4_escrita'] = 'ACEITA (defeito?)'
    except Exception as ex:  # noqa: BLE001
        out['B4_escrita'] = 'recusada: ' + repr(ex)[:80]
    out['ok'] = (all(not s['Ano_0'] and all(n['ok'] for n in s['niveis']) for s in estados)
                 and estados[3]['K3'] == '' and estados[3]['cab']['D4'] == '' and 'recusada' in out['B4_escrita']
                 and '2026' in estados[0]['cab']['D4'])
    return out


def s_O43_estados(qa):
    """Estados reais provocados pelo filtro (Bio) e a forma curta/longa mostrada."""
    p = qa.ws('Painel')
    out = []
    for rot, seq in (('REJEITADO? 8974 Tudo', [('L4', 'DIMENSION 1'), ('C3', 'Glicose'), ('H3', "'8974"), ('K4', 'Tudo (sem filtro)')]),
                     ('periodo vazio Fev/2024', [('K3', 2024), ('K4', 'Fev')]),
                     ('lote urina 6353', [('H3', "'6353"), ('K4', 'Tudo (sem filtro)')])):
        for c, v in seq:
            C.por(qa, p, c, v, local=False)
        j = [p.Range(f'J{7 + k}') for k in range(NLV)]
        out.append({'estado': rot, 'J': [t(c.Text) for c in j], 'Q4': t(p.Range('Q4').Text),
                    'cor': [c.DisplayFormat.Interior.Color for c in j]})
    return out


# ============================================================================================ ADR-091
def s_adr091(qa, fin, regras_extra=None):
    lo = qa.lo('tblLotesMaterial')
    cab = [str(c.Name) for c in lo.ListColumns]
    regras = []
    if lo.ListRows.Count:
        for l in lo.DataBodyRange.Value:
            d = dict(zip(cab, [t(x) for x in l]))
            if any(d.values()):
                regras.append(d)
    fora = {up(r['LOTE']) for r in regras if up(r['REGRA']) == 'FORA_DO_CQ'}
    so = collections.defaultdict(set)
    for r in regras:
        if up(r['REGRA']) == 'SO_ANALITOS':
            for a in r['ANALITOS'].split(';'):
                if a.strip():
                    so[up(r['LOTE'])].add(up(a))
    mz = {up(r['LOTE']): up(r['MATRIZ']) for r in regras if up(r['REGRA']) == 'MATRIZ'}
    cat = {up(x[0]) for x in qa.espec().Range('A4:A43').Value if x[0]}
    lotes = collections.defaultdict(list)
    for r in fin:
        lotes[up(r['LOTE'])].append(r)
    viol, det = [], {}
    for lt in sorted(fora):
        rs = lotes.get(lt, [])
        for r in rs:
            if r['STATUS_ANALITICO'] != 'LOTE_FORA_DO_CQ' and r['STATUS_ANALITICO'] not in PRECEDE_LOTE:
                viol.append(('FORA', lt, r['ID_REGISTRO'], r['STATUS_ANALITICO']))
            if r['PARTICIPA_ESTATISTICA'] == 'SIM':
                viol.append(('FORA-participa', lt, r['ID_REGISTRO']))
        det[f'FORA {lt}'] = {'linhas': len(rs), 'status': dict(collections.Counter(r['STATUS_ANALITICO'] for r in rs))}
    for lt, ans in sorted(so.items()):
        rs = lotes.get(lt, [])
        dn = [r for r in rs if up(r['ANALITO']) in ans]
        ou = [r for r in rs if up(r['ANALITO']) not in ans]
        for r in ou:
            if r['STATUS_ANALITICO'] != 'LOTE_OUTRO_MATERIAL' and r['STATUS_ANALITICO'] not in PRECEDE_LOTE:
                viol.append(('SO', lt, r['ID_REGISTRO'], r['STATUS_ANALITICO']))
            if r['PARTICIPA_ESTATISTICA'] == 'SIM':
                viol.append(('SO-participa', lt, r['ID_REGISTRO']))
        for r in dn:
            if r['STATUS_ANALITICO'] in LOTE_ST:
                viol.append(('SO-dentro', lt, r['ID_REGISTRO']))
        det[f'SO {lt}'] = {'dentro': dict(collections.Counter(r['STATUS_ANALITICO'] for r in dn)),
                           'dentro_participam': sum(1 for r in dn if r['PARTICIPA_ESTATISTICA'] == 'SIM'),
                           'outros': dict(collections.Counter(r['STATUS_ANALITICO'] for r in ou))}
    for lt, m in sorted(mz.items()):
        rs = lotes.get(lt, [])
        mzs = collections.Counter(up(r['MATRIZ']) for r in rs)
        ser = [r for r in rs if up(r['ANALITO']) in cat and r['PARTICIPA_ESTATISTICA'] == 'SIM']
        if any(k != m for k in mzs):
            viol.append(('MATRIZ', lt, dict(mzs)))
        if ser:
            viol.append(('MATRIZ-serico', lt, len(ser), [r['ID_REGISTRO'] for r in ser[:3]]))
        det[f'MATRIZ {lt}'] = {'linhas': len(rs), 'matriz': dict(mzs), 'serico_participa': len(ser)}
    lote_rows = [r for r in fin if r['STATUS_ANALITICO'] in LOTE_ST]
    regra_lotes = fora | set(so) | set(mz)
    fora_regra = [r['ID_REGISTRO'] for r in lote_rows if up(r['LOTE']) not in regra_lotes]
    qai = qa.ler('tblQA_Integracao')
    a13 = [x for x in qai if x.get('CODIGO') == 'A13']
    return {'regras': len(regras), 'por_tipo': dict(collections.Counter(up(r['REGRA']) for r in regras)), 'detalhe': det,
            'violacoes': viol[:20], 'n_violacoes': len(viol), 'LOTE_total': len(lote_rows), 'LOTE_fora_de_regra': fora_regra[:5],
            'A13': len(a13), 'linhas': len(fin),
            'ok': not viol and not fora_regra and len(a13) == len(lote_rows)}


# ============================================================================================ ADR-093 oraculo
def oraculo_093(qa):
    """Para cada analito do catalogo: o que deve acontecer com um manual dele no lote 6354 (MATRIZ URINA)."""
    dp = qa.ler('tblDeParaAnalitos')
    rec = qa.ler('tblDB_Recebimento')
    cat = [t(x[0]) for x in qa.espec().Range('A4:A43').Value if x[0]]
    orig = collections.defaultdict(set)
    tem_dp = set()
    for d in dp:
        if t(d.get('ANALITO_QCINI')):
            orig[up(d['ANALITO_QCINI'])].add(up(d['ANALITO_ORIGEM']))
            # linha da matriz SEM ANALITO_QCINI ("sem cadastro no QC_INI") nao e de/para: e assim que a urina chega
            tem_dp.add((up(d.get('MATRIZ')), up(d.get('ANALITO_ORIGEM'))))
    colm = 'MATRIZ' if rec and 'MATRIZ' in rec[0] else None
    colo = next((c for c in ('ANALITO_ORIGEM', 'ANALITO', 'EXAME') if rec and c in rec[0]), None)
    recebidos = {(up(r.get(colm)), up(r.get(colo))) for r in rec} if colm and colo else set()
    res = {}
    for a in cat:
        cods = orig.get(up(a), set())
        if len(cods) != 1:
            res[a] = ('RECUSA', f'{len(cods)} codigos')
            continue
        c = next(iter(cods))
        if ('URINA', c) in tem_dp:
            res[a] = ('RECUSA', f'{c} tem de/para na URINA')
        elif ('URINA', c) not in recebidos:
            res[a] = ('RECUSA', f'{c} nunca recebido na URINA')
        else:
            res[a] = ('ACEITA', c)
    return res, {'colunas_rec': [colm, colo], 'pares_urina': sorted(x[1] for x in recebidos if x[0] == 'URINA')[:30]}


# ============================================================================================ manuais / ATUALIZAR
def manual_cols(qa):
    lo = qa.lo('tblResultados_Manuais')
    return lo, [str(c.Name) for c in lo.ListColumns]


def ler_manual(qa, linha):
    lo, cab = manual_cols(qa)
    return {c: (t(C.limpo(lo.ListColumns(c).DataBodyRange.Cells(linha, 1).Value)),
                t(lo.ListColumns(c).DataBodyRange.Cells(linha, 1).Text)) for c in cab}


def escrever(qa, linha, col, v):
    lo, _ = manual_cols(qa)
    cel = lo.ListColumns(col).DataBodyRange.Cells(linha, 1)
    if isinstance(v, str):
        cel.FormulaLocal = v
    else:
        cel.Value = v
    qa.ex.esperar()


def regra_teste(qa, senha, linhas):
    """Hema: regras de teste na tblLotesMaterial (NOTA TESTE-AUDITORIA), so na copia."""
    lo = qa.lo('tblLotesMaterial')
    ws = lo.Parent
    try:
        ws.Unprotect(senha)
    except Exception:
        pass
    ev = qa.ex.xl.EnableEvents
    qa.ex.xl.EnableEvents = False
    cab = [str(c.Name) for c in lo.ListColumns]
    try:
        for d in linhas:
            vals = lo.DataBodyRange.Value if lo.ListRows.Count else ()
            i = next((k + 1 for k, l in enumerate(vals) if not any(t(x) for x in l)), None)
            if i is None:
                lo.ListRows.Add()
                i = lo.ListRows.Count
            for c, v in d.items():
                lo.ListColumns(c).DataBodyRange.Cells(i, 1).Value = v
    finally:
        qa.ex.xl.EnableEvents = ev
    return cab


def s_manuais(qa, fin, aud, senha):
    out = {}
    p, dr = qa.ws('Painel'), qa.ws('Digitar Resultados')
    lo, cab = manual_cols(qa)
    out['A3'] = t(dr.Range('A3').Value)[:300]
    an = 'Glicose' if BIO else 'WBC'
    if BIO:
        C.por(qa, p, 'L4', 'DIMENSION 1', local=False)
        C.por(qa, p, 'C3', 'Glicose', local=False)
    sel_eq = t(qa.nome('selEquipamento').Value)
    reais = sorted([r for r in fin if r['STATUS_ANALITICO'] == 'ATIVO' and r['ORIGEM_RESULTADO'] == 'INTERFACEAMENTO'
                    and t(r['ANALITO']) == an and int(r['NIVEL'] or 0) == 1 and isinstance(r['DATA_HORA'], dt.datetime)
                    and (not BIO or t(r['EQUIPAMENTO']) == 'DIMENSION 1')], key=lambda r: r['DATA_HORA'])
    R = reais[-25]
    dR = R['DATA_HORA']
    dia = dR.strftime('%d/%m/%Y')
    LT = "'" + t(R['LOTE'])
    EQ = t(R['EQUIPAMENTO'])
    out['referencia'] = {k: t(R[k]) for k in ('ID_REGISTRO', 'EQUIPAMENTO', 'LOTE', 'ANALITO', 'RESULTADO', 'DATA_HORA', 'RUN')}
    # regras de teste (Hema) e oraculo do ADR-093 (Bio)
    if not BIO:
        cont = collections.Counter(t(r['LOTE']) for r in fin)
        cand = sorted([l for l, n in cont.items() if n >= 20 and l != t(R['LOTE'])], key=lambda l: cont[l])
        lfora, lso = cand[0], cand[1]
        regra_teste(qa, senha, [{'LOTE': lfora, 'REGRA': 'FORA_DO_CQ', 'NOTA': 'TESTE-AUDITORIA R4V'},
                                {'LOTE': lso, 'REGRA': 'SO_ANALITOS', 'ANALITOS': 'WBC', 'NOTA': 'TESTE-AUDITORIA R4V'}])
        out['regras_teste'] = {'FORA_DO_CQ': [lfora, cont[lfora]], 'SO_ANALITOS_WBC': [lso, cont[lso]]}
    else:
        o93, info = oraculo_093(qa)
        aceitos = [a for a, (k, _) in o93.items() if k == 'ACEITA']
        recus = [a for a, (k, m) in o93.items() if k == 'RECUSA' and 'nunca' in m]
        out['oraculo_093'] = {'aceitos': {a: o93[a][1] for a in aceitos}, 'recusados_amostra': {a: o93[a][1] for a in recus[:8]},
                              'info': info}
        outro_aceito = next((a for a in aceitos if a not in ('Creatinina', 'Glicose', 'Cloro')), None)
        recusado = next((a for a in recus if a not in ('Lactato', 'Bilirrubina total')), recus[0] if recus else None)
        out['escolhidos_093'] = {'Creatinina': o93.get('Creatinina'), 'outro_aceito': [outro_aceito, o93.get(outro_aceito)],
                                 'recusado': [recusado, o93.get(recusado)]}
    dr.Activate()
    qa.ex.xl.EnableEvents = True
    L = {}

    def novo(rot, cps):
        _, r = C.linha_livre_manual(qa)
        for col, v in cps:
            escrever(qa, r, col, v)
        L[rot] = r
        return r

    base_eq = [] if not BIO else [('EQUIPAMENTO', 'DIMENSION 1')]
    # M1: sem EQUIPAMENTO e sem NIVEL
    r1 = novo('M1_sem_nivel', [('DATA', dia), ('HORA', '02:41'), ('LOTE', LT), ('ANALITO', an), ('RESULTADO', '222,222'),
                               ('MOTIVO', MOTIVO)])
    ce = lo.ListColumns('EQUIPAMENTO').DataBodyRange.Cells(r1, 1)
    out['O19_auto'] = {'valor': t(ce.Value), 'fundo': ce.Interior.Color, 'fonte': ce.Font.Color,
                       'nota': t(ce.Comment.Text()) if ce.Comment is not None else '', 'selEquipamento': sel_eq}
    if BIO:
        escrever(qa, r1, 'EQUIPAMENTO', 'DIMENSION 2')
        out['O19_digitado'] = {'valor': t(ce.Value), 'fundo': ce.Interior.Color, 'nota': ce.Comment is not None}
        a = out['O19_auto']
        out['O19_ok'] = (a['valor'] == sel_eq == 'DIMENSION 1' and a['fundo'] == AMBAR and a['fonte'] == AMBAR_FONTE
                         and a['nota'] == f'Preenchido pelo Painel com {sel_eq}: confira antes de ATUALIZAR DADOS.'
                         and out['O19_digitado']['valor'] == 'DIMENSION 2' and out['O19_digitado']['fundo'] != AMBAR
                         and not out['O19_digitado']['nota'] and 'EQUIPAMENTO' in out['A3'])
        escrever(qa, r1, 'EQUIPAMENTO', 'DIMENSION 1')
    novo('M2_sem_hora', [('DATA', dia)] + base_eq + [('LOTE', LT), ('NIVEL', '1'), ('ANALITO', an),
                                                     ('RESULTADO', '222,222'), ('MOTIVO', MOTIVO)])
    hm = (dR + dt.timedelta(minutes=3)).strftime('%H:%M')
    novo('M3_conflito', [('DATA', dia), ('HORA', hm), ('EQUIPAMENTO', EQ), ('LOTE', LT), ('NIVEL', '1'), ('ANALITO', an),
                         ('RESULTADO', '211,111'), ('MOTIVO', MOTIVO)])
    novo('M4_implausivel', [('DATA', dia), ('HORA', '02:47')] + base_eq + [('LOTE', LT), ('NIVEL', '1'), ('ANALITO', an),
                                                                           ('RESULTADO', '-1'), ('MOTIVO', MOTIVO)])
    novo('M5_texto', [('DATA', dia), ('HORA', '02:49')] + base_eq + [('LOTE', LT), ('NIVEL', '1'), ('ANALITO', an),
                                                                     ('RESULTADO', 'xyz'), ('MOTIVO', MOTIVO)])
    if BIO:
        b = [('DATA', dia), ('EQUIPAMENTO', 'DIMENSION 1'), ('NIVEL', '1'), ('MOTIVO', MOTIVO)]
        novo('M6_8794_fora', b + [('HORA', '02:51'), ('LOTE', "'8794"), ('ANALITO', 'Glicose'), ('RESULTADO', '50')])
        novo('M7_8353_glicose', b + [('HORA', '02:53'), ('LOTE', "'8353"), ('ANALITO', 'Glicose'), ('RESULTADO', '51')])
        novo('M8_6354_creatinina', b + [('HORA', '02:55'), ('LOTE', "'6354"), ('ANALITO', 'Creatinina'), ('RESULTADO', '12,345')])
        oa, rc = out['escolhidos_093']['outro_aceito'][0], out['escolhidos_093']['recusado'][0]
        if oa:
            novo('M9_6354_outro_aceito', b + [('HORA', '02:57'), ('LOTE', "'6354"), ('ANALITO', oa), ('RESULTADO', '3,21')])
        if rc:
            novo('M10_6354_recusado', b + [('HORA', '02:59'), ('LOTE', "'6354"), ('ANALITO', rc), ('RESULTADO', '4,32')])
    else:
        lfora, lso = out['regras_teste']['FORA_DO_CQ'][0], out['regras_teste']['SO_ANALITOS_WBC'][0]
        b = [('DATA', dia), ('NIVEL', '1'), ('MOTIVO', MOTIVO)]
        novo('M6_fora', b + [('HORA', '02:51'), ('LOTE', "'" + lfora), ('ANALITO', 'WBC'), ('RESULTADO', '7,11')])
        novo('M7_so_RBC', b + [('HORA', '02:53'), ('LOTE', "'" + lso), ('ANALITO', 'RBC'), ('RESULTADO', '4,22')])
        novo('M8_so_WBC', b + [('HORA', '02:55'), ('LOTE', "'" + lso), ('ANALITO', 'WBC'), ('RESULTADO', '7,33')])
    # inativar um ponto real e comentar outro (OC-0066 / ADR-094)
    pts = [r for r in reais[-60:-30] if r['ID_REGISTRO'] != R['ID_REGISTRO']]
    ri = t(qa.run('mIntegracao.InativarPorId', t(pts[0]['ID_REGISTRO']), 'TESTE-AUDITORIA R4V inativação', False, ''))
    rcm = t(qa.run('mIntegracao.GravarComentarioTecnico', t(pts[1]['ID_REGISTRO']), 'TESTE-AUDITORIA R4V comentário', ''))
    out['inativar'] = {'id': t(pts[0]['ID_REGISTRO']), 'ret': ri[:120]}
    out['comentar'] = {'id': t(pts[1]['ID_REGISTRO']), 'ret': rcm[:120]}
    sync0 = qa.lo('tblSyncLog').ListRows.Count
    aud.novos()
    t0 = time.time()
    r_at = t(qa.run('mIntegracao.AtualizarDadosAutomatico', teto=1800))
    out['atualizar_completo'] = {'ret': r_at[:300], 's': round(time.time() - t0, 1)}
    tr1 = aud.novos()
    lei = {k: ler_manual(qa, r) for k, r in L.items()}
    ids = {k: v['ID_REGISTRO'][0] for k, v in lei.items()}
    fin1 = qa.ler('tblCQ_Final')
    fid = {r['ID_REGISTRO']: r for r in fin1}
    out['ids'] = ids
    out['m'] = {k: {'STATUS': v['STATUS'][0], 'RUN_txt': v['RUN'][1], 'DETALHE': v['DETALHE'][1][:220],
                    'REGISTRADO_EM_txt': v.get('REGISTRADO_EM', ('', ''))[1],
                    'final': {c: t(fid.get(ids[k], {}).get(c)) for c in ('STATUS_ANALITICO', 'PARTICIPA_ESTATISTICA',
                                                                          'ANALITO', 'ANALITO_ORIGEM', 'ANALITO_QCINI',
                                                                          'BLOCO', 'MATRIZ', 'MOTIVO_EXCLUSAO_AUTOMATICA')}}
                for k, v in lei.items()}
    trilha = collections.defaultdict(list)
    for x in tr1:
        if x['acao'] and x['acao'].startswith('RESULTADO_MANUAL'):
            k = next((k for k, i in ids.items() if i and re.search(re.escape(i) + r'\b', x['txt'])), None)
            trilha[k].append([x['acao'], x['txt'][:400]])
    out['trilha_1'] = dict(trilha)
    qai = qa.ler('tblQA_Integracao')
    qa_por = collections.defaultdict(list)
    for x in qai:
        txt = ' '.join(t(v) for v in x.values())
        for k, i in ids.items():
            if i and re.search(re.escape(i) + r'\b', txt):
                qa_por[k].append(t(x.get('CODIGO')))
    out['qa_por_manual'] = dict(qa_por)
    # ResumoQA (OC-0024)
    rq = t(qa.run('mIntegracao.ResumoQA'))
    err = collections.Counter(x['CODIGO'] for x in qai if x['SEVERIDADE'] == 'ERRO')
    ngov = sum(err.get(c, 0) for c in ('E01', 'E02', 'E03', 'E10', 'E11'))
    lin = rq.split('\r\n')
    out['O24'] = {'resumo': lin, 'erros': dict(err), 'gov': ngov,
                  'ok': (err.get('E04', 0) > 0 and any(x.startswith(f'• {err["E04"]} resultado(s) manual(is) incompleto(s)')
                                                       for x in lin)
                         and 'ERRO DE GOVERNANÇA' not in rq.upper().replace('ERRO(S) DE GOVERNANÇA', '')
                         and (ngov == 0) == ('governança' not in rq))}
    GLOBAL['fin1'] = fin1
    # O23 / O58 vereditos por manual
    m = out['m']
    esp_trilha = {'M1_sem_nivel': 'RESULTADO_MANUAL_RECUSADO', 'M2_sem_hora': 'RESULTADO_MANUAL_RECUSADO',
                  'M5_texto': 'RESULTADO_MANUAL_RECUSADO', 'M3_conflito': 'RESULTADO_MANUAL_EM_CONFLITO',
                  'M4_implausivel': 'RESULTADO_MANUAL_FORA_DA_ESTATISTICA'}
    esp_status = {'M1_sem_nivel': 'MANUAL_INCOMPLETO', 'M2_sem_hora': 'MANUAL_INCOMPLETO', 'M5_texto': 'MANUAL_INCOMPLETO',
                  'M3_conflito': 'CONFLITO_MANUAL', 'M4_implausivel': 'IMPLAUSIVEL'}
    if BIO:
        esp_trilha.update({'M6_8794_fora': 'RESULTADO_MANUAL_FORA_DA_ESTATISTICA',
                           'M7_8353_glicose': 'RESULTADO_MANUAL_FORA_DA_ESTATISTICA',
                           'M8_6354_creatinina': 'RESULTADO_MANUAL_INCLUIDO'})
        esp_status.update({'M6_8794_fora': 'LOTE_FORA_DO_CQ', 'M7_8353_glicose': 'LOTE_OUTRO_MATERIAL',
                           'M8_6354_creatinina': 'ATIVO'})
        if 'M9_6354_outro_aceito' in L:
            esp_trilha['M9_6354_outro_aceito'] = 'RESULTADO_MANUAL_INCLUIDO'
            esp_status['M9_6354_outro_aceito'] = 'ATIVO'
        if 'M10_6354_recusado' in L:
            esp_trilha['M10_6354_recusado'] = 'RESULTADO_MANUAL_FORA_DA_ESTATISTICA'
            esp_status['M10_6354_recusado'] = 'LOTE_OUTRO_MATERIAL'
    else:
        esp_trilha.update({'M6_fora': 'RESULTADO_MANUAL_FORA_DA_ESTATISTICA', 'M7_so_RBC': 'RESULTADO_MANUAL_FORA_DA_ESTATISTICA',
                           'M8_so_WBC': 'RESULTADO_MANUAL_INCLUIDO'})
        esp_status.update({'M6_fora': 'LOTE_FORA_DO_CQ', 'M7_so_RBC': 'LOTE_OUTRO_MATERIAL', 'M8_so_WBC': 'ATIVO'})
    conf = {}
    for k in L:
        obt_t = [a for a, _ in trilha.get(k, [])]
        conf[k] = {'status': [m[k]['final']['STATUS_ANALITICO'] or m[k]['STATUS'], esp_status.get(k)],
                   'trilha': [obt_t, esp_trilha.get(k)],
                   'ok': (m[k]['final']['STATUS_ANALITICO'] or m[k]['STATUS']) == esp_status.get(k)
                   and obt_t == [esp_trilha.get(k)]}
    out['conferencia_1'] = conf
    # ADR-093 (Bio)
    if BIO:
        def ok93(k, cod):
            f = m[k]['final']
            return (f['STATUS_ANALITICO'] == 'ATIVO' and f['PARTICIPA_ESTATISTICA'] == 'SIM' and up(f['ANALITO']) == up(cod)
                    and up(f['ANALITO_ORIGEM']) == up(cod) and f['ANALITO_QCINI'] == '' and f['BLOCO'] == ''
                    and up(f['MATRIZ']) == 'URINA' and f"lancado como '{cod}'" in f['MOTIVO_EXCLUSAO_AUTOMATICA']
                    and 'I05' in qa_por.get(k, []) and 'A13' not in qa_por.get(k, [])
                    and any('lancado como' in tx for _, tx in trilha.get(k, [])))
        cre = out['escolhidos_093']['Creatinina']
        r93 = {'M8_creatinina': ok93('M8_6354_creatinina', cre[1]) if cre and cre[0] == 'ACEITA' else 'oraculo nao aceita'}
        if 'M9_6354_outro_aceito' in L:
            r93['M9_outro'] = ok93('M9_6354_outro_aceito', out['escolhidos_093']['outro_aceito'][1][1])
        if 'M10_6354_recusado' in L:
            f = m['M10_6354_recusado']['final']
            r93['M10_recusado'] = (f['STATUS_ANALITICO'] == 'LOTE_OUTRO_MATERIAL' and f['PARTICIPA_ESTATISTICA'] != 'SIM'
                                   and 'A13' in qa_por.get('M10_6354_recusado', []))
        out['ADR093'] = r93
        out['ADR093_ok'] = all(v is True for v in r93.values())
    # sync log / inativacao
    out['sync'] = {'linhas': [sync0, qa.lo('tblSyncLog').ListRows.Count]}
    # completar o NIVEL do M1 e atualizar so entradas: INCLUIDO uma vez, nada repetido
    escrever(qa, L['M1_sem_nivel'], 'NIVEL', '1')
    t0 = time.time()
    r2 = t(qa.run('mIntegracao.AtualizarSoEntradas', teto=1500))
    out['atualizar_so_entradas'] = {'ret': r2[:160], 's': round(time.time() - t0, 1)}
    tr2 = [x for x in aud.novos() if x['acao'] and x['acao'].startswith('RESULTADO_MANUAL')]
    m1 = ler_manual(qa, L['M1_sem_nivel'])
    out['M1_depois'] = {'STATUS': m1['STATUS'][0], 'RUN_txt': m1['RUN'][1], 'DETALHE': m1['DETALHE'][1]}
    out['trilha_2'] = [[x['acao'], next((k for k, i in ids.items() if i and re.search(re.escape(i) + r'\b', x['txt'])), None)]
                       for x in tr2]
    out['O23_ok'] = (all(conf[k]['ok'] for k in ('M1_sem_nivel', 'M2_sem_hora', 'M3_conflito', 'M5_texto'))
                     and out['M1_depois']['STATUS'] == 'ATIVO'
                     and out['trilha_2'] == [['RESULTADO_MANUAL_INCLUIDO', 'M1_sem_nivel']])
    out['O58_ok'] = all(conf[k]['ok'] for k in conf if esp_trilha.get(k) == 'RESULTADO_MANUAL_FORA_DA_ESTATISTICA') and \
        not any(a == 'RESULTADO_MANUAL_INCLUIDO' for k in conf if esp_trilha.get(k) == 'RESULTADO_MANUAL_FORA_DA_ESTATISTICA'
                for a in conf[k]['trilha'][0])
    out['linhas_manual'] = L
    return out


# ============================================================================================ ADR-094 / OC-0066
RX_DH = re.compile(r'^\d{2}/\d{2}/\d{4} \d{2}:\d{2}:\d{2}$')
RX_D = re.compile(r'^\d{2}/\d{2}/\d{4}$')
FORA_ESCOPO = {('tblAuditoria', 'Data'), ('tblAuditoria', 'DataCorrida')}
ABAS_FORA = ('BI_Data', 'Eventos_Westgard', 'Registros')


def de_serial(x):
    d = dt.datetime(1899, 12, 30) + dt.timedelta(days=float(x))
    return (d + dt.timedelta(microseconds=500000)).replace(microsecond=0)


def s_adr094(qa, extra_ids):
    """Varredura propria: toda tabela de toda aba; colunas com formato de data; 1a, ultima, 25 espalhadas e as linhas
    tocadas pelo teste. Texto com '#' = FAIL; valor com hora -> texto dd/mm/aaaa hh:mm:ss igual ao meu; so data -> dd/mm/aaaa."""
    achados, conferidas, colunas = [], 0, []
    for ws in qa.wb.Worksheets:
        for lo in ws.ListObjects:
            if lo.DataBodyRange is None:
                continue
            n = lo.ListRows.Count
            cab = [str(c.Name) for c in lo.ListColumns]
            idc = next((i for i, c in enumerate(cab) if c in ('ID_REGISTRO', 'ID')), None)
            alvo_lin = set()
            if idc is not None and extra_ids:
                ids = lo.ListColumns(idc + 1).DataBodyRange.Value2
                ids = ids if isinstance(ids, tuple) else ((ids,),)
                alvo_lin = {i + 1 for i, x in enumerate(ids) if t(x[0]) in extra_ids}
            for ci, c in enumerate(cab):
                rng = lo.ListColumns(ci + 1).DataBodyRange
                nf = str(rng.Cells(1, 1).NumberFormat)
                v2 = rng.Value2
                v2 = v2 if isinstance(v2, tuple) else ((v2,),)
                preench = [i + 1 for i, x in enumerate(v2) if num(x[0])]
                if not preench:
                    continue
                if not re.search(r'(d|y|a){1,4}', nf.lower()) or not re.search(r'[/-]', nf):
                    continue
                passo = max(1, len(preench) // 25)
                amostra = sorted(set(preench[::passo] + preench[-3:] + [i for i in alvo_lin if i in preench]))
                if ws.Name == 'Sync_Log' or lo.Name == 'tblSyncLog':
                    amostra = sorted(set(amostra + preench[-5:]))
                colunas.append(f'{ws.Name}!{lo.Name}[{c}] ({len(preench)}/{n}) {nf}')
                for i in amostra:
                    cel = rng.Cells(i, 1)
                    tx = t(cel.Text)
                    x = v2[i - 1][0]
                    d = de_serial(x)
                    fr = float(x) % 1
                    hora = 1e-7 < fr < 1 - 1e-7
                    esc = (lo.Name, c) in FORA_ESCOPO or ws.Name in ABAS_FORA or c.upper() == 'HORA'
                    conferidas += 1
                    if '#' in tx:
                        achados.append({'onde': f'{ws.Name}!{lo.Name}[{c}] linha {i}', 'texto': tx, 'tipo': '####',
                                        'fora_do_escopo': esc, 'largura': float(cel.ColumnWidth)})
                    elif c.upper() == 'HORA':
                        continue
                    elif hora and not (RX_DH.match(tx) and tx == d.strftime('%d/%m/%Y %H:%M:%S')):
                        achados.append({'onde': f'{ws.Name}!{lo.Name}[{c}] linha {i}', 'texto': tx,
                                        'esperado': d.strftime('%d/%m/%Y %H:%M:%S'), 'tipo': 'data-hora', 'fora_do_escopo': esc})
                    elif c.upper() == 'DATA' and lo.Name != 'tblComentariosTecnicos' and not esc and not RX_D.match(tx):
                        achados.append({'onde': f'{ws.Name}!{lo.Name}[{c}] linha {i}', 'texto': tx,
                                        'esperado': d.strftime('%d/%m/%Y'), 'tipo': 'so-data', 'fora_do_escopo': esc})
                    elif not hora and not (RX_D.match(tx) or RX_DH.match(tx)) and not esc:
                        achados.append({'onde': f'{ws.Name}!{lo.Name}[{c}] linha {i}', 'texto': tx,
                                        'esperado': d.strftime('%d/%m/%Y'), 'tipo': 'data', 'fora_do_escopo': esc})
    db = qa.ws('DB')
    a5 = t(db.Range('A5').Text)
    sl = qa.lo('tblSyncLog')
    cabs = [str(c.Name) for c in sl.ListColumns]
    ult = {c: t(sl.ListColumns(c).DataBodyRange.Cells(sl.ListRows.Count, 1).Text) for c in cabs
           if c in ('INICIO', 'FIM', 'DATA_MIN', 'DATA_MAX')}
    fq = qa.lo('tblCQ_Final')
    ci = [str(c.Name) for c in fq.ListColumns].index('DATA_INATIVACAO') + 1
    dina = []
    ids = fq.ListColumns('ID_REGISTRO').DataBodyRange.Value2
    for i, x in enumerate(ids, 1):
        if t(x[0]) in extra_ids:
            cel = fq.ListColumns(ci).DataBodyRange.Cells(i, 1)
            if num(cel.Value2):
                dina.append([t(x[0]), t(cel.Text), float(cel.ColumnWidth)])
    no_esc = [a for a in achados if not a['fora_do_escopo']]
    return {'colunas_data': len(colunas), 'celulas_conferidas': conferidas, 'achados_no_escopo': no_esc[:20],
            'achados_fora_do_escopo': [a for a in achados if a['fora_do_escopo']][:10], 'DB_A5': a5,
            'DB_A5_ok': len(re.findall(r'\d{2}/\d{2}/\d{4} \d{2}:\d{2}:\d{2}', a5)) >= 2,
            'sync_ultima_linha': ult, 'DATA_INATIVACAO_teste': dina, 'colunas': colunas,
            'ok': not no_esc and len(re.findall(r'\d{2}/\d{2}/\d{4} \d{2}:\d{2}:\d{2}', a5)) >= 2
            and all(RX_DH.match(v) for v in ult.values()) and bool(dina) and all(RX_DH.match(x[1]) for x in dina)}


def s_O50(qa, fin):
    pr = qa.ws('Principal - Resultados')
    a5 = t(pr.Range('A5').Text)
    n = len(fin)
    sim = sum(1 for r in fin if r['PARTICIPA_ESTATISTICA'] == 'SIM')
    ina = sum(1 for r in fin if r['STATUS_ANALITICO'] == 'INATIVADO')
    man = sum(1 for r in fin if up(r['ORIGEM_RESULTADO']) == 'MANUAL')
    sub = sum(1 for r in fin if r['STATUS_ANALITICO'] == 'SUBSTITUIDO_MANUAL')
    tp = collections.Counter(t(r['TIPO_PLOTAGEM_LJ']) for r in fin)
    esp = f'Total: {milhar(n)} · participam: {milhar(sim)}'
    a5n = re.sub(r'\s+', ' ', a5)
    return {'A5': a5, 'meu': {'total': n, 'participam': sim, 'inativados': ina, 'manuais': man, 'substituidos': sub,
                              'tipo_plotagem': dict(tp)},
            'ok': a5n.startswith(esp) and f'inativados: {milhar(ina)}' in a5n and f'manuais: {milhar(man)}' in a5n
            and f'X vermelho: {milhar(tp.get("X_VERMELHO", 0))}' in a5n
            and f'substituídos por manual: {milhar(sub)}' in a5n and not re.search(r'\d,0\b', a5n)}


# ============================================================================================ main
def main():
    src = os.path.join(C.REPO, f'QC_{PROD}.xlsm')
    os.makedirs(C.PASTA, exist_ok=True)
    h0 = C.sha(src)
    copia = os.path.join(C.PASTA, f'{PROD}_valida_{time.strftime("%H%M%S")}.xlsm')
    import shutil
    shutil.copy2(src, copia)
    hc = C.sha(copia)
    suf = ('_' + '_'.join(SECOES)) if SECOES else ''
    L = Log(os.path.join(C.PASTA, f'valida_{PROD}{suf}.json'))
    L['ambiente'] = {'producao': os.path.basename(src), 'sha_prod_12': h0, 'sha_copia_12': hc, 'copia': os.path.basename(copia)}
    senha = C.senha_planilha()
    qa = C.qa_final.QA(PROD, copia)
    L['excel_pid'] = qa.ex.pid
    try:
        rodar(qa, L, senha)
    except Exception as e:  # noqa: BLE001
        L['ERRO_FATAL'] = {'erro': repr(e), 'trace': traceback.format_exc()[-3000:]}
    finally:
        try:
            qa.fechar(salvar=False)
        except Exception:
            pass
    L['producao_intacta'] = {'sha_antes': h0, 'sha_depois': C.sha(src), 'igual': h0 == C.sha(src)}
    try:
        os.remove(copia)
        L['copia_apagada'] = True
    except Exception as e:  # noqa: BLE001
        L['copia_apagada'] = repr(e)
    print('FIM', PROD, flush=True)


def rodar(qa, L, senha):
    p = qa.ws('Painel')
    p.Activate()
    qa.ex.esperar()
    fin = qa.ler('tblCQ_Final')
    L['linhas_tblCQ_Final'] = len(fin)
    L['painel_protegido'] = bool(p.ProtectContents)
    if quer('estaticas'):
        secao(L, 'estaticas', s_estaticas, qa)
        secao(L, 'O50_inicio', s_O50, qa, fin)
        secao(L, 'docs', s_docs)
        secao(L, 'botoes', s_botoes, qa)
    esc, ncand, _ = escolher_ceq(qa)
    L['ceq_escolha'] = {'escolhido': esc, 'candidatos': ncand}
    if quer('sessao'):
        secao(L, 'O33_sem_sessao', s_sem_sessao, qa, esc[1], esc[2])
    if quer('adr091'):
        secao(L, 'ADR091_inicio', s_adr091, qa, fin)
    if BIO and quer('O12'):
        secao(L, 'O12', s_O12, qa, fin)
    if BIO and quer('O40'):
        secao(L, 'O40', s_O40, qa, fin)
    if BIO and quer('O43'):
        secao(L, 'O43_estados', s_O43_estados, qa)
    if quer('O41'):
        secao(L, 'O41', s_O41, qa, senha)
    criar_usuario(qa, USR, 'ANALISTA')
    L['login'] = fazer_login(qa, USR)
    aud = Audit(qa)
    if quer('ceq'):
        ceq = secao(L, 'ceq', s_ceq, qa, senha, aud, esc)
        secao(L, 'O54', s_O54, qa, senha, aud)
    if quer('manuais'):
        man = secao(L, 'manuais', s_manuais, qa, fin, aud, senha)
        fin1 = GLOBAL.pop('fin1', None)
        if fin1:
            secao(L, 'ADR091_depois', s_adr091, qa, fin1)
            secao(L, 'O50_depois', s_O50, qa, qa.ler('tblCQ_Final'))
            ids = set((man.get('ids') or {}).values()) | {man['inativar']['id'], man['comentar']['id']}
            secao(L, 'ADR094', s_adr094, qa, ids)
            if not BIO:
                secao(L, 'O49_depois', lambda: s_estaticas(qa).get('O49'))


if __name__ == '__main__':
    t0 = time.time()
    main()
    print(f'tempo {time.time() - t0:.0f}s', flush=True)
