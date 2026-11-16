# -*- coding: utf-8 -*-
"""reteste06_usuario_f1_hema.py -- Usuario Simulado, ciclo 2026-11-15_reteste06, frente F1, RETESTE da OC-0073 (Hematologia).

Irmao do reteste06_usuario_f1.py (Bioquimica): o MESMO roteiro, com o que muda na Hematologia -- 3 niveis, equipamento
unico XN1000 (sem L4 no Painel), coluna MATRIZ (preenchida pelo PrepararLinha), Painel A7:T9, Estatistica por ANO
(B3/D3, que o roteiro poe em 2026 como o usuario) e lote B4, Eng_Saida com 3 niveis (valor Y..AA, ID AM..AO, alvo BW..CB, marcas de Westgard C..W).

OC-0073 (ADR-096, entrega 43): linha manual copiada e colada com o ID_REGISTRO de outra. Antes ela aparecia ATIVA na
tela e o motor a deixava CONFLITO_MANUAL ("vale a 1a linha"), fora da estatistica/LJ/Westgard sem aviso.

Fluxo (aba Digitar Resultados, eventos LIGADOS, Application.Interactive = True durante a digitacao para que o aviso
do EntradaMudou apareca; as caixas sao lidas pelo TEXTO das janelas do PID desta sessao e fechadas no OK -- nunca
captura de tela):
  S0  origem digitada celula a celula (WBC, XN1000, lote corrente, N1, 14/11/2026 08:10) pelo usuario 1; troca de login
      para o usuario 2 (o resto e feito por ele)
  P1  colar sem editar (A..MOTIVO, o bloco destravado) na proxima linha livre
  P1b colar a linha inteira da tabela pela COM (o caminho do Tester)
  P1c a linha inteira com a protecao como o usuario a ve (Protect sem UserInterfaceOnly): o Excel recusa?
  P2  colar e editar: NIVEL 2 e RESULTADO = media + 4 DP do lote (deve dar 1-3s no Westgard)
  P2c colar e trocar so o RESULTADO (mesma chave e instante)
  P2d colar e trocar HORA e RESULTADO
  P3a digitar o ID da origem numa linha vazia (antes dos dados)
  P3b sobrescrever o ID de uma linha preenchida (N3) com a variante "man_<n>" (minusculo, sem zeros)
  P4  bloco de 3 linhas colado de uma vez
  P5  mover (recortar A..MOTIVO e colar 3 linhas abaixo da livre)
  P6  ID fora do padrao MAN_ ("TESTEAUD9") repetido em duas linhas
  P7  repeticao com os eventos DESLIGADOS (o "dado antigo"; HORA 15:00, N2)
  ATUALIZAR DADOS #1 (nucleo do botao: mIntegracao.AtualizarDadosAutomatico, le o DB_SEAC so para leitura)
  S13 correcao do P7 como a mensagem manda (apagar o ID) e apagar a linha do P1 (A..MOTIVO)
  ATUALIZAR DADOS #2
Modo 'carimbo' (fluxo vizinho R6-01 da Bio): digita A, apaga A..MOTIVO, espera 150 s, botao Digitar Resultados, digita B
onde o cursor parou (data/hora = carimbo de A + 1 min), controle C numa linha nunca usada, ATUALIZAR.
Leituras: linha da tela (STATUS/RUN/DETALHE em Text), tblCQ_Final, tblQA_Integracao, tblAuditoria nova, Painel (C3/H3/
K3/K4: engPainel, A7:T9), Eng_Saida (pontos do LJ por ID, alvo, marcas), aba Estatistica (WBC N1..N3, janela B3/D3/B4).
Oraculo proprio (Python puro, statistics): n/media/DP/CV dos PARTICIPA = SIM da tblCQ_Final no filtro do Painel
(WBC, XN1000, lote, nov/2026) e na janela da Estatistica (lote B4, anos B3..D3); 1-3s pelo alvo do lote
(mLotes.MediaDoLote/DPDoLote).

Uso: python -u reteste06_usuario_f1_hema.py <nova|base|carimbo|carimbo_base>
  nova = copia da producao atual (raiz); base = copia do backup anterior a entrega 43 (_backup_pre_integracao_2026-11-15_2059).
Regras: so COPIAS em %TEMP%/qcw/auditoria/2026-11-15_reteste06/; nada e salvo; fechar() encerra so o PID do QA;
usuarios TESTE_AUDITORIA_R6H/R6HB/R6HC (ANALISTA) criados so na copia; valores sinteticos com 3 decimais.
"""
import datetime as dt
import os
import shutil
import statistics as stt
import sys
import threading
import time
import traceback

AQUI = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, AQUI)
_argv = sys.argv[:]
sys.argv = [sys.argv[0], 'Hematologia']
import valida_adr090_091 as V  # noqa: E402
sys.argv = _argv
import repro_comum as C  # noqa: E402
from qa_inativar_ux import criar_usuario, fazer_login  # noqa: E402

assert not V.BIO
PROD = 'Hematologia'
NLV = 3
CICLO = '2026-11-15_reteste06'
PASTA = os.path.join(os.environ['TEMP'], 'qcw', 'auditoria', CICLO)
C.PASTA = PASTA
BACKUP = os.path.join(C.REPO, '_backup_pre_integracao_2026-11-15_2059', f'QC_{PROD}.xlsm')
USR1, USR2, USR3 = 'TESTE_AUDITORIA_R6H', 'TESTE_AUDITORIA_R6HB', 'TESTE_AUDITORIA_R6HC'
EQ, AN, DIA = 'XN1000', 'WBC', '14/11/2026'
MOTIVO = 'Outro'
t, up, num = V.t, V.up, V.num


# ----------------------------------------------------------------------------------------------- caixas (sem imagem)
class Caixas(threading.Thread):
    """Le o texto das caixas (#32770) do PID desta sessao e clica OK (ou fecha). Nao captura imagem."""

    def __init__(self, pid, teto=3600):
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
                    if p == self.pid and win32gui.GetClassName(h) == '#32770' and win32gui.IsWindowVisible(h):
                        achou.append(h)
                except Exception:  # noqa: BLE001
                    pass
                return True
            try:
                win32gui.EnumWindows(cb, None)
            except Exception:  # noqa: BLE001
                pass
            for h in achou:
                textos, botoes = [], {}

                def cc(ch, _):
                    try:
                        tx = win32gui.GetWindowText(ch)
                        if win32gui.GetClassName(ch) == 'Button':
                            botoes[tx.replace('&', '').strip().upper()] = ch
                        elif tx:
                            textos.append(tx)
                    except Exception:  # noqa: BLE001
                        pass
                    return True
                try:
                    win32gui.EnumChildWindows(h, cc, None)
                except Exception:  # noqa: BLE001
                    pass
                self.vistos.append({'titulo': win32gui.GetWindowText(h), 'texto': '\n'.join(textos),
                                    'botoes': sorted(botoes), 'quando': time.strftime('%H:%M:%S')})
                try:
                    if 'OK' in botoes:
                        win32gui.PostMessage(botoes['OK'], win32con.BM_CLICK, 0, 0)
                    else:
                        win32gui.PostMessage(h, win32con.WM_CLOSE, 0, 0)
                except Exception:  # noqa: BLE001
                    pass
            time.sleep(0.6 if achou else 0.2)


# ----------------------------------------------------------------------------------------------- tabela manual
def lo_m(qa):
    return qa.lo('tblResultados_Manuais')


def cel(qa, r, col):
    return lo_m(qa).ListColumns(col).DataBodyRange.Cells(r, 1)


def faixa(qa, r1, r2, c1='ID_REGISTRO', c2='MOTIVO'):
    ws = qa.ws('Digitar Resultados')
    return ws.Range(cel(qa, r1, c1), cel(qa, r2, c2))


def linha_tabela(qa, r):
    return lo_m(qa).ListRows(r).Range


def escrever(qa, r, col, v):
    c = cel(qa, r, col)
    if isinstance(v, str) and col != 'ID_REGISTRO':
        c.FormulaLocal = v
    else:
        c.Value = v
    qa.ex.esperar()
    time.sleep(0.25)


def ler(qa, r):
    d = V.ler_manual(qa, r)
    return {k: (v[0] if v[0] == v[1] or k not in ('RUN', 'STATUS', 'DETALHE', 'REGISTRADO_EM', 'DATA', 'HORA')
                else {'v': v[0], 'txt': v[1]}) for k, v in d.items()}


def idd(qa, r):
    return t(cel(qa, r, 'ID_REGISTRO').Value)


def livres(qa, k=1, depois=0):
    """k linhas livres CONSECUTIVAS (ID e DATA vazios), a partir da linha 'depois'+1."""
    lo = lo_m(qa)
    ids = lo.ListColumns('ID_REGISTRO').DataBodyRange.Value
    dts = lo.ListColumns('DATA').DataBodyRange.Value
    n = len(ids)
    for i in range(depois, n - k + 1):
        if all(ids[j][0] in (None, '') and dts[j][0] in (None, '') for j in range(i, i + k)):
            return i + 1
    raise RuntimeError('sem linhas livres')


def marca(qa):
    try:
        return t(qa.wb.Names('qcUltimoIdManual').RefersTo)
    except Exception as e:  # noqa: BLE001
        return repr(e)[:80]


# ----------------------------------------------------------------------------------------------- painel e oraculo
def oraculo(fin, lote, ano=2026, mes=11, anos=None):
    """n/media/DP/CV dos PARTICIPA = SIM (WBC, XN1000, lote) no mes (Painel) ou nos anos [a0, a1] (Estatistica)."""
    out = []
    for nv in range(1, NLV + 1):
        def no_periodo(d):
            if not isinstance(d, dt.datetime):
                return False
            if anos:
                return anos[0] <= d.year <= anos[1]
            return d.year == ano and d.month == mes
        vals = [float(r['RESULTADO']) for r in fin
                if up(r.get('PARTICIPA_ESTATISTICA')) == 'SIM' and t(r.get('ANALITO')) == AN and up(r.get('EQUIPAMENTO')) == EQ
                and t(r.get('LOTE')) == lote and num(r.get('NIVEL')) and int(r['NIVEL']) == nv
                and no_periodo(r.get('DATA_HORA')) and num(r.get('RESULTADO'))]
        n = len(vals)
        m = stt.fmean(vals) if n else None
        s = stt.stdev(vals) if n > 1 else None
        out.append({'nivel': nv, 'n': n, 'soma': sum(vals), 'media': m, 'dp': s,
                    'cv_pct': (100 * s / m) if (s is not None and m) else None})
    return out


def rel(a, b):
    if not num(a) or not num(b):
        return None
    return abs(a - b) / max(abs(b), 1e-300)


def eng_linhas(qa):
    e = qa.ws('Eng_Saida')
    n = int(e.Range('I1').Value or 0)
    if not n:
        return []
    blk = e.Range(e.Cells(3, 1), e.Cells(2 + n, 80)).Value
    out = []
    for k, r in enumerate(blk):
        out.append({'slot': k + 1, 'run': t(r[1]), 'val': [r[24 + j] for j in range(NLV)],
                    'id': [t(r[38 + j]) for j in range(NLV)], 'cor': [r[62 + j] for j in range(NLV)],
                    'alvo': [(r[74 + 2 * j], r[75 + 2 * j]) for j in range(NLV)],
                    'flags': [list(r[2 + j * 7:2 + j * 7 + 7]) for j in range(NLV)]})
    return out


def painel(qa, lote, ids, fin):
    p = qa.ws('Painel')
    for c, v in (('C3', AN), ('H3', "'" + lote), ('K3', 2026), ('K4', 'Nov')):
        C.por(qa, p, c, v, local=False)
    time.sleep(1.0)
    qa.ex.esperar()
    eng = eng_linhas(qa)
    meus = [x for x in eng if any(i and any(i == j or i.startswith(j + '~') for j in ids) for i in x['id'])]
    e = qa.ws('Estatística')
    cab0 = {a: t(e.Range(a).Text) for a in ('B3', 'D3', 'B4')}
    # o usuario poe o Ano De/Ate da aba Estatistica (B3/D3, celulas destravadas) no ano dos dados (2026)
    for c in ('B3', 'D3'):
        C.por(qa, e, c, 2026, local=False)
    time.sleep(1.0)
    qa.ex.esperar()
    cab = {a: t(e.Range(a).Text) for a in ('B3', 'D3', 'B4', 'K5', 'N4', 'P4', 'R4')}
    cab['antes_de_digitar_o_ano'] = cab0
    est = {}
    for nv in range(1, NLV + 1):
        try:
            est[f'N{nv}'] = list(qa.estat_motor(AN, nv) or [])
        except Exception as ex:  # noqa: BLE001
            est[f'N{nv}'] = repr(ex)[:120]
    # oraculo da janela da Estatistica (lote B4, anos B3..D3)
    o_est = None
    try:
        a0, a1 = int(float(e.Range('B3').Value)), int(float(e.Range('D3').Value))
        o_est = {'lote': t(e.Range('B4').Value), 'anos': [a0, a1],
                 'oraculo': oraculo(fin, t(e.Range('B4').Value), anos=(a0, a1))}
    except Exception as ex:  # noqa: BLE001
        o_est = {'erro': repr(ex)[:160]}
    return {'filtro': {c: t(p.Range(c).Text) for c in ('C3', 'H3', 'K3', 'K4', 'G3', 'G4')},
            'engPainel': [list(l) for l in qa.nome('engPainel').Value][:NLV],
            'A7T9': [[x for x in l] for l in p.Range('A7:T9').Value],
            'eng_n_pontos': len(eng), 'eng_S1': t(qa.ws('Eng_Saida').Range('S1').Value),
            'eng_E1': t(qa.ws('Eng_Saida').Range('E1').Value),
            'eng_meus_pontos': meus, 'eng_ultimos': eng[-3:], 'estatistica': est, 'estatistica_cab': cab,
            'estatistica_oraculo': o_est}


def comparar(o, pa):
    out = []
    for i in range(NLV):
        ep = pa['engPainel'][i]
        es = pa['estatistica'].get(f'N{i + 1}')
        oe = (pa.get('estatistica_oraculo') or {}).get('oraculo')
        d = {'nivel': i + 1, 'oraculo_mes': (o[i]['n'], o[i]['media'], o[i]['dp'], o[i]['cv_pct']),
             'engPainel': ep[:5],
             'rel_media_painel': rel(ep[2], o[i]['media']), 'rel_dp_painel': rel(ep[3], o[i]['dp']),
             'n_painel_ok': ep[1] == o[i]['n']}
        if isinstance(es, list) and len(es) >= 3 and oe:
            d['estatistica'] = es
            d['oraculo_estatistica'] = (oe[i]['n'], oe[i]['media'], oe[i]['dp'])
            d['n_estat_ok'] = es[0] == oe[i]['n']
            d['rel_media_estat'] = rel(es[1], oe[i]['media'])
            d['rel_dp_estat'] = rel(es[2], oe[i]['dp'])
        out.append(d)
    return out


# ----------------------------------------------------------------------------------------------- leituras pos-ATUALIZAR
def depois(qa, linhas, msg):
    fin = qa.ler('tblCQ_Final')
    tela = {k: ler(qa, r) for k, r in linhas.items()}
    ids = sorted({t(v['ID_REGISTRO']) for v in tela.values() if t(v['ID_REGISTRO'])})
    meus = [r for r in fin if any(t(r['ID_REGISTRO']) == i or t(r['ID_REGISTRO']).startswith(i + '~') for i in ids)
            or t(r['ID_REGISTRO']).startswith('MAN_INVALIDO')]
    cols = ('ID_REGISTRO', 'STATUS_ANALITICO', 'PARTICIPA_ESTATISTICA', 'NIVEL', 'RESULTADO', 'DATA_HORA', 'RUN',
            'MOTIVO_EXCLUSAO_AUTOMATICA', 'EQUIPAMENTO', 'MATRIZ', 'LOTE')
    fq = [{c: (r[c].isoformat(sep=' ') if isinstance(r.get(c), dt.datetime) else r.get(c)) for c in cols} for r in meus]
    qai = []
    try:
        for x in qa.ler('tblQA_Integracao'):
            s = ' | '.join(t(v) for v in x.values() if v not in (None, ''))
            if any(i in s for i in ids) or 'TESTEAUD9' in s or 'MAN_INVALIDO' in s or 'repetido' in s:
                qai.append(s[:600])
    except Exception as ex:  # noqa: BLE001
        qai = [repr(ex)[:200]]
    return fin, {'msg': msg, 'tela': tela, 'tblCQ_Final': fq, 'n_tblCQ_Final': len(fin), 'qa_integracao': qai,
                 'tem_til_L': [r['ID_REGISTRO'] for r in fq if '~L' in t(r['ID_REGISTRO'])]}


# ----------------------------------------------------------------------------------------------- fluxo
def escolher_lote(qa, fin0):
    cand = sorted([r for r in fin0 if up(r.get('PARTICIPA_ESTATISTICA')) == 'SIM' and t(r.get('ANALITO')) == AN
                   and up(r.get('EQUIPAMENTO')) == EQ and isinstance(r.get('DATA_HORA'), dt.datetime)
                   and up(r.get('ORIGEM_RESULTADO')).startswith('INTERF')], key=lambda r: r['DATA_HORA'])
    for r in reversed(cand):
        lt = t(r['LOTE'])
        o = oraculo(fin0, lt)
        if all(o[i]['n'] >= 3 for i in range(NLV)):
            return lt, cand[-1]['DATA_HORA']
    return None, cand[-1]['DATA_HORA'] if cand else None


def fluxo(qa, L, rotulo):
    senha = C.senha_planilha()
    fin0 = qa.ler('tblCQ_Final')
    L['tblCQ_Final_inicio'] = len(fin0)
    lote, ultimo = escolher_lote(qa, fin0)
    L['lote'] = lote
    L['ultimo_real_WBC'] = str(ultimo)
    alvo = {}
    for nv in range(1, NLV + 1):
        alvo[nv] = (qa.run('mLotes.MediaDoLote', lote, AN, nv, teto=60), qa.run('mLotes.DPDoLote', lote, AN, nv, teto=60))
    L['alvo_lote'] = {f'N{k}': v for k, v in alvo.items()}
    o_ini = oraculo(fin0, lote)

    def val(nv, k):
        m, s = alvo[nv]
        if not num(m) or not num(s) or not s:
            m, s = o_ini[nv - 1]['media'], o_ini[nv - 1]['dp']
        return round(m + k * s, 3)
    V1 = {'S0': val(1, 0.1), 'P2c': val(1, -0.4), 'P2d': val(1, 0.3), 'P3a': val(1, -0.2), 'P6a': val(1, 0.05),
          'P6b': val(1, -0.05)}
    V2 = {'P2': val(2, 4.0), 'P5': val(2, 0.2), 'P7': val(2, 0.1)}
    V3 = {'P3b': val(3, -0.3)}
    L['valores'] = {'N1': V1, 'N2': V2, 'N3': V3}
    fmt = lambda x: f'{x:.3f}'.replace('.', ',')  # noqa: E731

    # ---- antes
    pa = painel(qa, lote, [], fin0)
    L['antes'] = {'oraculo': o_ini, 'painel': pa, 'comparacao': comparar(o_ini, pa)}
    lo = lo_m(qa)
    ws = qa.ws('Digitar Resultados')
    cab = [str(c.Name) for c in lo.ListColumns]
    L['ambiente_tabela'] = {'colunas': cab, 'linhas_com_id': sum(1 for x in lo.ListColumns('ID_REGISTRO').DataBodyRange.Value
                                                                 if x[0] not in (None, '')),
                            'ProtectContents': bool(ws.ProtectContents), 'ProtectionMode_UIOnly': bool(ws.ProtectionMode),
                            'Locked_por_coluna': {c: lo.ListColumns(c).DataBodyRange.Cells(1, 1).Locked for c in cab},
                            'faixa_ID_MOTIVO': t(faixa(qa, 1, 1).Address),
                            'marca_qcUltimoIdManual': marca(qa)}

    # ---- digitacao
    qa.run('mIntegracao.IrDigitarResultados', teto=120)
    L['botao_inicio'] = {'aba': t(qa.ex.xl.ActiveSheet.Name), 'celula': t(qa.ex.xl.ActiveCell.Address)}
    qa.ex.xl.EnableEvents = True
    qa.ex.xl.Interactive = True
    vig = Caixas(qa.ex.pid)
    vig.start()
    linhas, passos = {}, {}
    k0 = [0]

    def passo(nome, extra=None):
        d = {'caixas': vig.vistos[k0[0]:], 'marca': marca(qa)}
        k0[0] = len(vig.vistos)
        if extra:
            d.update(extra)
        passos[nome] = d
        L['passos'] = passos

    def novo(rot, campos, r=None):
        r = r or livres(qa)
        for col, v in campos:
            escrever(qa, r, col, v)
        linhas[rot] = r
        return r

    def base_campos(hora, nv, v):
        return [('DATA', DIA), ('HORA', hora), ('EQUIPAMENTO', EQ), ('LOTE', "'" + lote), ('NIVEL', str(nv)),
                ('ANALITO', AN), ('RESULTADO', fmt(v)), ('MOTIVO', MOTIVO)]

    try:
        r0 = novo('S0_origem', base_campos('08:10', 1, V1['S0']))
        orig0 = ler(qa, r0)
        id0 = t(orig0['ID_REGISTRO'])
        passo('S0_origem', {'linha': r0, 'tela': orig0})
        qa.ex.xl.Interactive = False
        L['login_usuario2'] = fazer_login(qa, USR2)
        qa.run('mIntegracao.IrDigitarResultados', teto=120)
        qa.ex.xl.EnableEvents = True
        qa.ex.xl.Interactive = True
        time.sleep(1.5)

        # P1 colar sem editar
        r1 = livres(qa)
        faixa(qa, r0, r0).Copy(faixa(qa, r1, r1))
        qa.ex.esperar()
        qa.ex.xl.CutCopyMode = False
        linhas['P1_colar_AK'] = r1
        passo('P1_colar_AK', {'linha': r1, 'tela': ler(qa, r1), 'origem_igual': ler(qa, r0) == orig0})

        # P1b linha inteira pela COM
        r2 = livres(qa)
        linha_tabela(qa, r0).Copy(linha_tabela(qa, r2))
        qa.ex.esperar()
        qa.ex.xl.CutCopyMode = False
        linhas['P1b_colar_linha_inteira_COM'] = r2
        passo('P1b_colar_linha_inteira_COM', {'linha': r2, 'tela': ler(qa, r2), 'origem_igual': ler(qa, r0) == orig0})

        # P1c linha inteira com a protecao como o usuario a ve
        r3 = livres(qa)
        ev3 = {'linha': r3}
        if ws.ProtectContents:
            try:
                ws.Protect(senha, False, True, True, False)
                ev3['ProtectionMode_UIOnly_durante'] = bool(ws.ProtectionMode)
                try:
                    linha_tabela(qa, r0).Copy(linha_tabela(qa, r3))
                    ev3['linha_inteira'] = 'ACEITA'
                except Exception as ex:  # noqa: BLE001
                    ev3['linha_inteira'] = 'RECUSADA: ' + repr(ex)[:300]
                qa.ex.xl.CutCopyMode = False
                ev3['tela_depois_linha_inteira'] = {'ID': idd(qa, r3), 'DATA': t(cel(qa, r3, 'DATA').Value)}
                r3b = livres(qa, 1, r3)
                try:
                    faixa(qa, r0, r0).Copy(faixa(qa, r3b, r3b))
                    ev3['bloco_AK'] = 'ACEITO'
                except Exception as ex:  # noqa: BLE001
                    ev3['bloco_AK'] = 'RECUSADO: ' + repr(ex)[:300]
                qa.ex.xl.CutCopyMode = False
                ev3['linha_AK'] = r3b
                ev3['tela_AK'] = ler(qa, r3b)
                linhas['P1c_AK_protegida'] = r3b
                if ev3['tela_depois_linha_inteira']['ID']:
                    linhas['P1c_linha_inteira_protegida'] = r3
            finally:
                ws.Protect(senha, False, True, True, True)
                ev3['ProtectionMode_UIOnly_restaurado'] = bool(ws.ProtectionMode)
        else:
            ev3['nota'] = 'aba sem protecao de conteudo'
        passo('P1c_protecao_do_usuario', ev3)

        # P2 colar e editar
        r4 = livres(qa)
        faixa(qa, r0, r0).Copy(faixa(qa, r4, r4))
        qa.ex.esperar()
        qa.ex.xl.CutCopyMode = False
        id_colado = idd(qa, r4)
        escrever(qa, r4, 'NIVEL', '2')
        escrever(qa, r4, 'RESULTADO', fmt(V2['P2']))
        linhas['P2_colar_editar'] = r4
        passo('P2_colar_editar', {'linha': r4, 'id_logo_apos_colar': id_colado, 'tela': ler(qa, r4)})

        # P2c so o RESULTADO
        r5 = livres(qa)
        faixa(qa, r0, r0).Copy(faixa(qa, r5, r5))
        qa.ex.esperar()
        qa.ex.xl.CutCopyMode = False
        escrever(qa, r5, 'RESULTADO', fmt(V1['P2c']))
        linhas['P2c_so_resultado'] = r5
        passo('P2c_so_resultado', {'linha': r5, 'tela': ler(qa, r5)})

        # P2d HORA e RESULTADO
        r6 = livres(qa)
        faixa(qa, r0, r0).Copy(faixa(qa, r6, r6))
        qa.ex.esperar()
        qa.ex.xl.CutCopyMode = False
        escrever(qa, r6, 'HORA', '16:10')
        escrever(qa, r6, 'RESULTADO', fmt(V1['P2d']))
        linhas['P2d_hora_resultado'] = r6
        passo('P2d_hora_resultado', {'linha': r6, 'tela': ler(qa, r6)})

        # P3a ID da origem digitado numa linha vazia
        r7 = livres(qa)
        escrever(qa, r7, 'ID_REGISTRO', id0)
        id_dig = idd(qa, r7)
        for col, v in base_campos('09:20', 1, V1['P3a']):
            escrever(qa, r7, col, v)
        linhas['P3a_id_digitado_vazia'] = r7
        passo('P3a_id_digitado_vazia', {'linha': r7, 'digitado': id0, 'id_logo_apos_digitar': id_dig, 'tela': ler(qa, r7)})

        # P3b variante minuscula sem zeros sobre uma linha preenchida (N3)
        r8 = novo('P3b_id_sobrescrito', base_campos('10:30', 3, V3['P3b']))
        id_auto = idd(qa, r8)
        variante = 'man_' + str(int(id0[4:]))
        escrever(qa, r8, 'ID_REGISTRO', variante)
        passo('P3b_id_sobrescrito', {'linha': r8, 'id_auto_antes': id_auto, 'digitado': variante, 'tela': ler(qa, r8)})

        # P4 bloco de 3 linhas
        r9 = livres(qa, 3)
        faixa(qa, r0, r0 + 2).Copy(faixa(qa, r9, r9 + 2))
        qa.ex.esperar()
        qa.ex.xl.CutCopyMode = False
        for k in range(3):
            linhas[f'P4_bloco_{k + 1}'] = r9 + k
        passo('P4_bloco_3', {'linhas': [r9, r9 + 1, r9 + 2], 'origem_ids': [idd(qa, r0 + k) for k in range(3)],
                             'ids': [idd(qa, r9 + k) for k in range(3)]})

        # P5 mover
        r10 = novo('P5_mover_origem', base_campos('11:40', 2, V2['P5']))
        id5 = idd(qa, r10)
        u10 = {c: t(cel(qa, r10, c).Value) for c in ('USUARIO', 'REGISTRADO_EM')}
        d10 = livres(qa, 1, livres(qa) + 2)
        faixa(qa, r10, r10).Cut(faixa(qa, d10, d10))
        qa.ex.esperar()
        del linhas['P5_mover_origem']
        linhas['P5_movida'] = d10
        cab_ak = cab[:cab.index('MOTIVO') + 1]
        passo('P5_mover', {'de': r10, 'para': d10, 'id_antes': id5, 'id_depois': idd(qa, d10),
                           'origem_AK_vazia': all(cel(qa, r10, c).Value in (None, '') for c in cab_ak),
                           'origem_resto': {c: t(cel(qa, r10, c).Value) for c in ('USUARIO', 'REGISTRADO_EM')},
                           'carimbo_antes': u10, 'tela_destino': ler(qa, d10)})

        # P6 ID invalido repetido
        r11 = novo('P6a_id_invalido', base_campos('12:50', 1, V1['P6a']))
        escrever(qa, r11, 'ID_REGISTRO', 'TESTEAUD9')
        r12 = novo('P6b_id_invalido', base_campos('13:30', 1, V1['P6b']))
        escrever(qa, r12, 'ID_REGISTRO', 'TESTEAUD9')
        passo('P6_id_invalido_repetido', {'linhas': [r11, r12], 'ids': [idd(qa, r11), idd(qa, r12)]})

        # P7 eventos desligados
        qa.ex.xl.EnableEvents = False
        r13 = livres(qa)
        faixa(qa, r0, r0).Copy(faixa(qa, r13, r13))
        qa.ex.xl.CutCopyMode = False
        for col, v in (('HORA', '15:00'), ('NIVEL', '2'), ('RESULTADO', fmt(V2['P7']))):
            escrever(qa, r13, col, v)
        qa.ex.xl.EnableEvents = True
        linhas['P7_eventos_desligados'] = r13
        passo('P7_eventos_desligados', {'linha': r13, 'tela': ler(qa, r13)})
    finally:
        qa.ex.xl.Interactive = False
        qa.ex.xl.EnableEvents = True
    L['linhas'] = linhas
    L['antes_atualizar'] = {k: ler(qa, r) for k, r in linhas.items()}
    L['origem_intacta_antes_atualizar'] = ler(qa, r0) == orig0

    # ---- ATUALIZAR #1
    aud = V.Audit(qa)
    try:
        msg1, s1 = qa.atualizar()
    except Exception as ex:  # noqa: BLE001
        msg1, s1 = 'ERRO ' + repr(ex), -1
    fin1, d1 = depois(qa, linhas, msg1)
    d1['s'] = round(s1, 1)
    d1['auditoria'] = aud.novos()
    d1['caixas_no_atualizar'] = vig.vistos[k0[0]:]
    k0[0] = len(vig.vistos)
    L['atualizar1'] = d1
    o1 = oraculo(fin1, lote)
    p1 = painel(qa, lote, [t(v['ID_REGISTRO']) for v in d1['tela'].values()], fin1)
    L['depois1'] = {'oraculo': o1, 'painel': p1, 'comparacao': comparar(o1, p1)}

    # ---- S13 correcao do P7 (apagar o ID) e apagar a linha do P1
    qa.run('mIntegracao.IrDigitarResultados', teto=120)
    qa.ex.xl.EnableEvents = True
    qa.ex.xl.Interactive = True
    try:
        id7_antes = idd(qa, linhas['P7_eventos_desligados'])
        cel(qa, linhas['P7_eventos_desligados'], 'ID_REGISTRO').ClearContents()
        qa.ex.esperar()
        time.sleep(0.5)
        id7_depois = idd(qa, linhas['P7_eventos_desligados'])
        id1 = idd(qa, linhas['P1_colar_AK'])
        faixa(qa, linhas['P1_colar_AK'], linhas['P1_colar_AK']).ClearContents()
        qa.ex.esperar()
        time.sleep(0.5)
        passo('S13_correcao', {'P7_id_antes': id7_antes, 'P7_id_depois': id7_depois, 'P1_id_apagado': id1,
                               'P1_linha_depois': ler(qa, linhas['P1_colar_AK'])})
    finally:
        qa.ex.xl.Interactive = False
    L['P1_id_apagado'] = id1

    # ---- ATUALIZAR #2
    aud2 = V.Audit(qa)
    try:
        msg2, s2 = qa.atualizar()
    except Exception as ex:  # noqa: BLE001
        msg2, s2 = 'ERRO ' + repr(ex), -1
    fin2, d2 = depois(qa, linhas, msg2)
    d2['s'] = round(s2, 1)
    d2['auditoria'] = aud2.novos()
    d2['P1_apagado_ainda_na_tblCQ_Final'] = [r['ID_REGISTRO'] for r in fin2 if t(r['ID_REGISTRO']).split('~')[0] == id1]
    vig.parar.set()
    d2['caixas_no_atualizar'] = vig.vistos[k0[0]:]
    L['atualizar2'] = d2
    o2 = oraculo(fin2, lote)
    p2 = painel(qa, lote, [t(v['ID_REGISTRO']) for v in d2['tela'].values()], fin2)
    L['depois2'] = {'oraculo': o2, 'painel': p2, 'comparacao': comparar(o2, p2)}
    L['origem_intacta_fim'] = ler(qa, r0)


def fluxo_carimbo(qa, L, rotulo):
    """R6-01 da Bio na Hema: apagar A..MOTIVO deixa USUARIO/REGISTRADO_EM e o proximo manual nessa linha e recusado?"""
    fin0 = qa.ler('tblCQ_Final')
    lote, _ = escolher_lote(qa, fin0)
    L['lote'] = lote
    qa.run('mIntegracao.IrDigitarResultados', teto=120)
    qa.ex.xl.EnableEvents = True

    def campos(data, hora, nv, v):
        return [('DATA', data), ('HORA', hora), ('EQUIPAMENTO', EQ), ('LOTE', "'" + lote), ('NIVEL', str(nv)),
                ('ANALITO', AN), ('RESULTADO', v), ('MOTIVO', MOTIVO)]
    m1 = qa.run('mLotes.MediaDoLote', lote, AN, 1, teto=60)
    m2 = qa.run('mLotes.MediaDoLote', lote, AN, 2, teto=60)
    fa = lambda x: f'{x:.3f}'.replace('.', ',')  # noqa: E731
    vA, vB, vC = fa(round(m1 + 0.001, 3)), fa(round(m1 + 0.002, 3)), fa(round(m2 + 0.003, 3))
    L['valores'] = {'A': vA, 'B': vB, 'C': vC}
    rA = livres(qa)
    for c, v in campos('14/11/2026', '07:40', 1, vA):
        escrever(qa, rA, c, v)
    a = ler(qa, rA)
    L['A_digitado'] = {'linha': rA, 'tela': a}
    faixa(qa, rA, rA).ClearContents()
    qa.ex.esperar()
    L['A_apagado'] = ler(qa, rA)
    t0 = a['REGISTRADO_EM']['v'] if isinstance(a['REGISTRADO_EM'], dict) else a['REGISTRADO_EM']
    t0 = dt.datetime.fromisoformat(str(t0)[:19])
    time.sleep(150)
    qa.run('mIntegracao.IrDigitarResultados', teto=120)
    L['botao_parou_em'] = {'aba': t(qa.ex.xl.ActiveSheet.Name), 'celula': t(qa.ex.xl.ActiveCell.Address),
                           'linha_da_tabela': qa.ex.xl.ActiveCell.Row - lo_m(qa).DataBodyRange.Row + 1}
    qa.ex.xl.EnableEvents = True
    hB = (t0 + dt.timedelta(minutes=1)).replace(second=0)
    if hB <= t0:
        hB += dt.timedelta(minutes=1)
    agora = dt.datetime.now()
    rB = L['botao_parou_em']['linha_da_tabela']
    for c, v in campos(hB.strftime('%d/%m/%Y'), hB.strftime('%H:%M'), 1, vB):
        escrever(qa, rB, c, v)
    L['B_digitado'] = {'linha': rB, 'data_hora_digitada': hB.isoformat(sep=' '),
                       'momento_da_digitacao': agora.isoformat(sep=' ', timespec='seconds'),
                       'T0_carimbo_de_A': t0.isoformat(sep=' '), 'tela': ler(qa, rB)}
    rC = livres(qa, 1, rB)
    for c, v in campos(hB.strftime('%d/%m/%Y'), hB.strftime('%H:%M'), 2, vC):
        escrever(qa, rC, c, v)
    L['C_controle'] = {'linha': rC, 'tela': ler(qa, rC)}
    msg, s = qa.atualizar()
    fin = qa.ler('tblCQ_Final')
    ids = [t(ler(qa, r)['ID_REGISTRO']) for r in (rB, rC)]
    L['atualizar'] = {'msg': msg, 's': round(s, 1),
                      'tela': {k: ler(qa, r) for k, r in (('B', rB), ('C', rC))},
                      'tblCQ_Final': [{c: (x[c].isoformat(sep=' ') if isinstance(x.get(c), dt.datetime) else x.get(c))
                                       for c in ('ID_REGISTRO', 'STATUS_ANALITICO', 'PARTICIPA_ESTATISTICA', 'NIVEL',
                                                 'RESULTADO', 'DATA_HORA', 'MOTIVO_EXCLUSAO_AUTOMATICA')}
                                      for x in fin if t(x['ID_REGISTRO']).split('~')[0] in ids],
                      'qa': [' | '.join(t(v) for v in x.values() if v not in (None, ''))[:400]
                             for x in qa.ler('tblQA_Integracao') if any(i in ' '.join(t(v) for v in x.values()) for i in ids)]}


def rodar(rotulo):
    os.makedirs(PASTA, exist_ok=True)
    base = rotulo.endswith('base')
    origem = BACKUP if base else os.path.join(C.REPO, f'QC_{PROD}.xlsm')
    destino = os.path.join(PASTA, f'{PROD}_F1_{rotulo}.xlsm')
    if os.path.exists(destino):
        os.remove(destino)
    h0 = C.sha(origem)
    shutil.copy2(origem, destino)
    L = C.Log(os.path.join(PASTA, f'{PROD}_F1_{rotulo}.json'))
    L['ambiente'] = {'origem': os.path.relpath(origem, C.REPO), 'sha_origem_12': h0, 'copia': os.path.basename(destino),
                     'sha_copia_12': C.sha(destino)}
    qa = C.qa_final.QA(PROD, destino)
    L['excel_pid'] = qa.ex.pid
    try:
        qa.ws('Painel').Activate()
        qa.ex.esperar()
        senha = C.senha_planilha()
        if rotulo.startswith('carimbo'):
            criar_usuario(qa, USR3, 'ANALISTA')
            L['login'] = fazer_login(qa, USR3)
            fluxo_carimbo(qa, L, rotulo)
        else:
            criar_usuario(qa, USR1, 'ANALISTA')
            criar_usuario(qa, USR2, 'ANALISTA')
            L['login_usuario1'] = fazer_login(qa, USR1)
            for w in ('Painel', 'Estatística'):
                try:
                    qa.ws(w).Unprotect(senha)
                except Exception:  # noqa: BLE001
                    pass
            fluxo(qa, L, rotulo)
    except Exception as ex:  # noqa: BLE001
        L['ERRO_FATAL'] = {'erro': repr(ex), 'trace': traceback.format_exc()[-3000:]}
    finally:
        try:
            qa.ex.xl.Interactive = False
        except Exception:  # noqa: BLE001
            pass
        try:
            qa.fechar(salvar=False)
        except Exception:  # noqa: BLE001
            pass
    L['origem_intacta'] = {'sha_antes': h0, 'sha_depois': C.sha(origem), 'igual': h0 == C.sha(origem)}
    try:
        os.remove(destino)
        L['copia_apagada'] = True
    except Exception as ex:  # noqa: BLE001
        L['copia_apagada'] = repr(ex)[:120]
    print('FIM', rotulo, flush=True)


if __name__ == '__main__':
    t0 = time.time()
    rodar(sys.argv[1])
    print(f'tempo {time.time() - t0:.0f}s', flush=True)
