# -*- coding: utf-8 -*-
"""reteste06_usuario_f1.py -- Usuario Simulado, ciclo 2026-11-15_reteste06, frente F1, RETESTE da OC-0073 (Bioquimica).

OC-0073 (ADR-096, entrega 43): linha manual copiada e colada com o ID_REGISTRO de outra. Antes ela aparecia ATIVA na
tela e o motor a deixava CONFLITO_MANUAL ("vale a 1a linha"), fora da estatistica/LJ/Westgard sem aviso.

Fluxo (aba Digitar Resultados, eventos LIGADOS, Application.Interactive = True durante a digitacao para que o aviso
do EntradaMudou apareca; as caixas sao lidas pelo TEXTO das janelas do PID desta sessao e fechadas no OK -- nunca
captura de tela):
  S0  origem digitada celula a celula (Glicose, DIMENSION 1, lote corrente, N1, 14/11/2026 08:10) pelo usuario 1;
      troca de login para o usuario 2 (o resto e feito por ele)
  S1  P1  colar sem editar: copia A:K (ID..MOTIVO, o bloco destravado) da origem para a proxima linha livre
  S2  P1b colar a linha inteira da tabela (A:Q) pela COM (o caminho do Tester no reteste 05)
  S3  P1c a linha inteira com a protecao como o usuario a ve (Protect sem UserInterfaceOnly): o Excel recusa?
  S4  P2  colar e editar: A:K da origem; NIVEL 2 e RESULTADO = media+4DP do lote (deve dar 1-3s no Westgard)
  S5  P2c colar e trocar so o RESULTADO (mesma chave e instante)
  S6  P2d colar e trocar HORA e RESULTADO
  S7  P3a digitar o ID da origem numa linha vazia (antes dos dados)
  S8  P3b sobrescrever o ID de uma linha preenchida com a variante "man_<n>" (minusculo, sem zeros) do ID da origem
  S9  P4  bloco de 3 linhas (A:K) colado de uma vez
  S10 P5  mover (recortar A:K e colar 3 linhas abaixo da livre)
  S11 P6  ID fora do padrao MAN_ ("TESTEAUD9") repetido em duas linhas
  S12 P7  repeticao com os eventos DESLIGADOS (A:K da origem; HORA 15:00, N2) -- o "dado antigo"
  ATUALIZAR DADOS #1 (nucleo do botao: mIntegracao.AtualizarDadosAutomatico, le o DB_SEAC so para leitura)
  S13 correcao do P7 como a mensagem manda (apagar o ID) e apagar a linha do P1 (A:K)
  ATUALIZAR DADOS #2
Leituras: linha da tela (STATUS/RUN/DETALHE em Text), tblCQ_Final, tblQA_Integracao (A01/E04), tblAuditoria nova,
Painel (L4/C3/H3/K3/K4 = DIMENSION 1/Glicose/lote/2026/Nov: engPainel, A7:T8), Eng_Saida (pontos do LJ por ID, alvo,
marcas Westgard), aba Estatistica (Glicose N1/N2).
Oraculo proprio (Python puro, statistics): n/media/DP/CV dos PARTICIPA = SIM da tblCQ_Final no filtro do Painel, e a
conta "antes + valores novos"; 1-3s pelo alvo do lote (mLotes.MediaDoLote/DPDoLote).

Uso: python -u reteste06_usuario_f1.py <nova|base>
  nova = copia da producao atual (raiz); base = copia do backup anterior a entrega 43 (_backup_pre_integracao_2026-11-15_1802).
Regras: so COPIAS em %TEMP%/qcw/auditoria/2026-11-15_reteste06/; nada e salvo; fechar() encerra so o PID do QA;
usuarios TESTE_AUDITORIA_R6 e TESTE_AUDITORIA_R6B (ANALISTA) criados so na copia; valores sinteticos com 3 decimais.
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
sys.argv = [sys.argv[0], 'Bioquimica']
import valida_adr090_091 as V  # noqa: E402
sys.argv = _argv
import repro_comum as C  # noqa: E402
from qa_inativar_ux import criar_usuario, fazer_login  # noqa: E402

PROD = 'Bioquimica'
CICLO = '2026-11-15_reteste06'
PASTA = os.path.join(os.environ['TEMP'], 'qcw', 'auditoria', CICLO)
C.PASTA = PASTA
BACKUP = os.path.join(C.REPO, '_backup_pre_integracao_2026-11-15_1802', f'QC_{PROD}.xlsm')
USR1, USR2 = 'TESTE_AUDITORIA_R6', 'TESTE_AUDITORIA_R6B'
EQ, AN, DIA = 'DIMENSION 1', 'Glicose', '14/11/2026'
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
def oraculo(fin, lote):
    out = []
    for nv in (1, 2):
        vals = [float(r['RESULTADO']) for r in fin
                if up(r.get('PARTICIPA_ESTATISTICA')) == 'SIM' and t(r.get('ANALITO')) == AN and up(r.get('EQUIPAMENTO')) == EQ
                and t(r.get('LOTE')) == lote and num(r.get('NIVEL')) and int(r['NIVEL']) == nv
                and isinstance(r.get('DATA_HORA'), dt.datetime) and r['DATA_HORA'].year == 2026 and r['DATA_HORA'].month == 11
                and num(r.get('RESULTADO'))]
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
        out.append({'slot': k + 1, 'run': t(r[1]), 'val': [r[24], r[25]], 'id': [t(r[38]), t(r[39])],
                    'x': [r[29], r[30]], 'alvo': [(r[74], r[75]), (r[76], r[77])],
                    'flags': [list(r[2:9]), list(r[9:16])]})
    return out


def painel(qa, lote, ids):
    p = qa.ws('Painel')
    for c, v in (('L4', EQ), ('C3', AN), ('H3', "'" + lote), ('K3', 2026), ('K4', 'Nov')):
        C.por(qa, p, c, v, local=False)
    time.sleep(1.0)
    qa.ex.esperar()
    eng = eng_linhas(qa)
    meus = [x for x in eng if any(i and any(i == j or i.startswith(j + '~') for j in ids) for i in x['id'])]
    est = {}
    for nv in (1, 2):
        try:
            est[f'N{nv}'] = list(qa.estat_motor(AN, nv) or [])
        except Exception as ex:  # noqa: BLE001
            est[f'N{nv}'] = repr(ex)[:120]
    e = qa.ws('Estatística')
    cab = {a: t(e.Range(a).Text) for a in ('B4', 'D4', 'B5', 'D5', 'H5', 'J5')}
    return {'filtro': {c: t(p.Range(c).Text) for c in ('L4', 'C3', 'H3', 'K3', 'K4')},
            'engPainel': [list(l) for l in qa.nome('engPainel').Value][:2],
            'A7T8': [[x for x in l] for l in p.Range('A7:T8').Value],
            'eng_n_pontos': len(eng), 'eng_S1': t(qa.ws('Eng_Saida').Range('S1').Value),
            'eng_E1': t(qa.ws('Eng_Saida').Range('E1').Value),
            'eng_meus_pontos': meus, 'eng_ultimos': eng[-3:], 'estatistica': est, 'estatistica_cab': cab}


# ----------------------------------------------------------------------------------------------- leituras pos-ATUALIZAR
def depois(qa, linhas, msg):
    fin = qa.ler('tblCQ_Final')
    tela = {k: ler(qa, r) for k, r in linhas.items()}
    ids = sorted({t(v['ID_REGISTRO']) for v in tela.values() if t(v['ID_REGISTRO'])})
    meus = [r for r in fin if any(t(r['ID_REGISTRO']) == i or t(r['ID_REGISTRO']).startswith(i + '~') for i in ids)
            or t(r['ID_REGISTRO']).startswith('MAN_INVALIDO')]
    cols = ('ID_REGISTRO', 'STATUS_ANALITICO', 'PARTICIPA_ESTATISTICA', 'NIVEL', 'RESULTADO', 'DATA_HORA', 'RUN',
            'MOTIVO_EXCLUSAO_AUTOMATICA', 'EQUIPAMENTO', 'LOTE')
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
def fluxo(qa, L, rotulo):
    senha = C.senha_planilha()
    fin0 = qa.ler('tblCQ_Final')
    L['tblCQ_Final_inicio'] = len(fin0)
    # lote corrente da serie (o do resultado mais recente ATIVO do interfaceamento) com os dois niveis em Nov/2026
    cand = sorted([r for r in fin0 if up(r.get('PARTICIPA_ESTATISTICA')) == 'SIM' and t(r.get('ANALITO')) == AN
                   and up(r.get('EQUIPAMENTO')) == EQ and isinstance(r.get('DATA_HORA'), dt.datetime)
                   and up(r.get('ORIGEM_RESULTADO')).startswith('INTERF')], key=lambda r: r['DATA_HORA'])
    lote = None
    for r in reversed(cand):
        lt = t(r['LOTE'])
        o = oraculo(fin0, lt)
        if o[0]['n'] >= 3 and o[1]['n'] >= 3:
            lote = lt
            break
    L['lote'] = lote
    alvo = {}
    for nv in (1, 2):
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
    V2 = {'P2': val(2, 4.0), 'P3b': val(2, -0.3), 'P5': val(2, 0.2), 'P7': val(2, 0.1)}
    L['valores'] = {'N1': V1, 'N2': V2}
    fmt = lambda x: f'{x:.3f}'.replace('.', ',')  # noqa: E731

    # ---- antes
    o0 = oraculo(fin0, lote)
    pa = painel(qa, lote, [])
    L['antes'] = {'oraculo': o0, 'painel': pa,
                  'oraculo_x_painel': [{'n': (o0[i]['n'], pa['engPainel'][i][1]),
                                        'rel_media': rel(pa['engPainel'][i][2], o0[i]['media']),
                                        'rel_dp': rel(pa['engPainel'][i][3], o0[i]['dp'])} for i in range(2)]}
    lo = lo_m(qa)
    ws = qa.ws('Digitar Resultados')
    cab = [str(c.Name) for c in lo.ListColumns]
    L['ambiente_tabela'] = {'colunas': cab, 'linhas_com_id': sum(1 for x in lo.ListColumns('ID_REGISTRO').DataBodyRange.Value
                                                                 if x[0] not in (None, '')),
                            'ProtectContents': bool(ws.ProtectContents), 'ProtectionMode_UIOnly': bool(ws.ProtectionMode),
                            'Locked_por_coluna': {c: lo.ListColumns(c).DataBodyRange.Cells(1, 1).Locked for c in cab},
                            'marca_qcUltimoIdManual': marca(qa)}

    # ---- digitacao
    qa.run('mIntegracao.IrDigitarResultados', teto=120)
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
        # troca de usuario (o resto e o usuario 2)
        qa.ex.xl.Interactive = False
        L['login_usuario2'] = fazer_login(qa, USR2)
        qa.run('mIntegracao.IrDigitarResultados', teto=120)
        qa.ex.xl.EnableEvents = True
        qa.ex.xl.Interactive = True
        time.sleep(1.5)

        # S1 P1 colar sem editar (A:K)
        r1 = livres(qa)
        faixa(qa, r0, r0).Copy(faixa(qa, r1, r1))
        qa.ex.esperar()
        qa.ex.xl.CutCopyMode = False
        linhas['P1_colar_AK'] = r1
        passo('P1_colar_AK', {'linha': r1, 'tela': ler(qa, r1), 'origem_igual': ler(qa, r0) == orig0})

        # S2 P1b linha inteira A:Q pela COM
        r2 = livres(qa)
        linha_tabela(qa, r0).Copy(linha_tabela(qa, r2))
        qa.ex.esperar()
        qa.ex.xl.CutCopyMode = False
        linhas['P1b_colar_linha_inteira_COM'] = r2
        passo('P1b_colar_linha_inteira_COM', {'linha': r2, 'tela': ler(qa, r2), 'origem_igual': ler(qa, r0) == orig0})

        # S3 P1c linha inteira com a protecao como o usuario a ve (sem UserInterfaceOnly)
        r3 = livres(qa)
        ev3 = {'linha': r3}
        if ws.ProtectContents:
            try:
                ws.Protect(senha, False, True, True, False)   # Password, DrawingObjects, Contents, Scenarios, UIOnly
                ev3['ProtectionMode_UIOnly_durante'] = bool(ws.ProtectionMode)
                try:
                    linha_tabela(qa, r0).Copy(linha_tabela(qa, r3))
                    ev3['linha_inteira'] = 'ACEITA'
                except Exception as ex:  # noqa: BLE001
                    ev3['linha_inteira'] = 'RECUSADA: ' + repr(ex)[:300]
                qa.ex.xl.CutCopyMode = False
                ev3['tela_depois_linha_inteira'] = {'ID': idd(qa, r3), 'DATA': t(cel(qa, r3, 'DATA').Value)}
                # e o bloco A:K (so celulas destravadas) na mesma situacao
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

        # S4 P2 colar e editar (NIVEL 2, RESULTADO media+4DP)
        r4 = livres(qa)
        faixa(qa, r0, r0).Copy(faixa(qa, r4, r4))
        qa.ex.esperar()
        qa.ex.xl.CutCopyMode = False
        id_colado = idd(qa, r4)
        escrever(qa, r4, 'NIVEL', '2')
        escrever(qa, r4, 'RESULTADO', fmt(V2['P2']))
        linhas['P2_colar_editar'] = r4
        passo('P2_colar_editar', {'linha': r4, 'id_logo_apos_colar': id_colado, 'tela': ler(qa, r4)})

        # S5 P2c colar e trocar so o RESULTADO
        r5 = livres(qa)
        faixa(qa, r0, r0).Copy(faixa(qa, r5, r5))
        qa.ex.esperar()
        qa.ex.xl.CutCopyMode = False
        escrever(qa, r5, 'RESULTADO', fmt(V1['P2c']))
        linhas['P2c_so_resultado'] = r5
        passo('P2c_so_resultado', {'linha': r5, 'tela': ler(qa, r5)})

        # S6 P2d colar e trocar HORA e RESULTADO
        r6 = livres(qa)
        faixa(qa, r0, r0).Copy(faixa(qa, r6, r6))
        qa.ex.esperar()
        qa.ex.xl.CutCopyMode = False
        escrever(qa, r6, 'HORA', '16:10')
        escrever(qa, r6, 'RESULTADO', fmt(V1['P2d']))
        linhas['P2d_hora_resultado'] = r6
        passo('P2d_hora_resultado', {'linha': r6, 'tela': ler(qa, r6)})

        # S7 P3a ID da origem digitado numa linha vazia (antes dos dados)
        r7 = livres(qa)
        escrever(qa, r7, 'ID_REGISTRO', id0)
        id_dig = idd(qa, r7)
        for col, v in base_campos('09:20', 1, V1['P3a']):
            escrever(qa, r7, col, v)
        linhas['P3a_id_digitado_vazia'] = r7
        passo('P3a_id_digitado_vazia', {'linha': r7, 'digitado': id0, 'id_logo_apos_digitar': id_dig, 'tela': ler(qa, r7)})

        # S8 P3b variante minuscula sem zeros sobre uma linha preenchida
        r8 = novo('P3b_id_sobrescrito', base_campos('10:30', 2, V2['P3b']))
        id_auto = idd(qa, r8)
        variante = 'man_' + str(int(id0[4:]))
        escrever(qa, r8, 'ID_REGISTRO', variante)
        passo('P3b_id_sobrescrito', {'linha': r8, 'id_auto_antes': id_auto, 'digitado': variante, 'tela': ler(qa, r8)})

        # S9 P4 bloco de 3 linhas (A:K de r0..r2)
        r9 = livres(qa, 3)
        faixa(qa, r0, r0 + 2).Copy(faixa(qa, r9, r9 + 2))
        qa.ex.esperar()
        qa.ex.xl.CutCopyMode = False
        for k in range(3):
            linhas[f'P4_bloco_{k + 1}'] = r9 + k
        passo('P4_bloco_3', {'linhas': [r9, r9 + 1, r9 + 2], 'origem_ids': [idd(qa, r0 + k) for k in range(3)],
                             'ids': [idd(qa, r9 + k) for k in range(3)]})

        # S10 P5 mover (recortar A:K e colar 3 linhas abaixo da proxima livre)
        r10 = novo('P5_mover_origem', base_campos('11:40', 2, V2['P5']))
        id5 = idd(qa, r10)
        u10 = {c: t(cel(qa, r10, c).Value) for c in ('USUARIO', 'REGISTRADO_EM')}
        d10 = livres(qa, 1, livres(qa) + 2)
        faixa(qa, r10, r10).Cut(faixa(qa, d10, d10))
        qa.ex.esperar()
        del linhas['P5_mover_origem']
        linhas['P5_movida'] = d10
        passo('P5_mover', {'de': r10, 'para': d10, 'id_antes': id5, 'id_depois': idd(qa, d10),
                           'origem_AK_vazia': all(cel(qa, r10, c).Value in (None, '') for c in cab[:cab.index('MOTIVO') + 1]),
                           'origem_resto': {c: t(cel(qa, r10, c).Value) for c in ('USUARIO', 'REGISTRADO_EM')},
                           'carimbo_antes': u10, 'tela_destino': ler(qa, d10)})

        # S11 P6 ID invalido repetido
        r11 = novo('P6a_id_invalido', base_campos('12:50', 1, V1['P6a']))
        escrever(qa, r11, 'ID_REGISTRO', 'TESTEAUD9')
        r12 = novo('P6b_id_invalido', base_campos('13:30', 1, V1['P6b']))
        escrever(qa, r12, 'ID_REGISTRO', 'TESTEAUD9')
        passo('P6_id_invalido_repetido', {'linhas': [r11, r12], 'ids': [idd(qa, r11), idd(qa, r12)]})

        # S12 P7 repeticao com os eventos desligados
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
    p1 = painel(qa, lote, [t(v['ID_REGISTRO']) for v in d1['tela'].values()])
    L['depois1'] = {'oraculo': o1, 'painel': p1}

    # ---- S13 correcao do P7 (apagar o ID) e apagar a linha do P1 (A:K)
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
    p2 = painel(qa, lote, [t(v['ID_REGISTRO']) for v in d2['tela'].values()])
    L['depois2'] = {'oraculo': o2, 'painel': p2}
    L['origem_intacta_fim'] = ler(qa, r0)


def rodar(rotulo):
    os.makedirs(PASTA, exist_ok=True)
    origem = BACKUP if rotulo == 'base' else os.path.join(C.REPO, f'QC_{PROD}.xlsm')
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
        criar_usuario(qa, USR1, 'ANALISTA')
        criar_usuario(qa, USR2, 'ANALISTA')
        L['login_usuario1'] = fazer_login(qa, USR1)
        senha = C.senha_planilha()
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
