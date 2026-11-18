# -*- coding: utf-8 -*-
"""reteste07_usuario_f1_hema.py -- Usuario Simulado, ciclo 2026-11-16_reteste07, frente F1, RETESTE da OC-0078 (Hematologia).

Irmao do reteste07_usuario_f1.py (Bioquimica): o MESMO roteiro, com os auxiliares da Hema (reteste06_usuario_f1_hema:
3 niveis, XN1000 sem L4, Painel A7:T9, Estatistica por ano B3/D3, Eng_Saida com 3 niveis). Analito WBC; niveis
A N1, D N2, M N3, G N1, F N1, K N2; SUBSTITUI_ID = numero de um HEM- real N1 (vale) e N2 (vizinho C11).

OC-0078 (ADR-097, entrega 44): a linha manual apagada (Delete de ID..MOTIVO ou de DATA..MOTIVO) ou movida (recorte)
guardava USUARIO e REGISTRADO_EM (colunas travadas). O proximo manual digitado nela herdava o carimbo: era recusado
("DATA/HORA posterior ao lancamento", MANUAL_INCOMPLETO) ou ficava no nome de quem apagou.

Fluxo (aba Digitar Resultados, eventos LIGADOS, Application.Interactive = True na digitacao para que os avisos do
EntradaMudou aparecam; caixas lidas pelo TEXTO das janelas do PID desta sessao e fechadas no OK -- nunca captura de
tela). Dois usuarios ANALISTA criados so na copia: U1 = TESTE_AUDITORIA_R7, U2 = TESTE_AUDITORIA_R7B.
  FASE 1 (U1): digita A, D, M, G, F (Glicose, DIMENSION 1, lote corrente, hoje, 03:0x) nas linhas r0..r0+4,
               deixa r0+5 e r0+6 livres e digita o controle K em r0+7. ATUALIZAR #1 (todos devem ficar ATIVO).
  FASE 2a (U1): apaga A (Delete ID..MOTIVO) e D (Delete so DATA..MOTIVO, o ID fica na selecao de fora).
  FASE 2b (U2, >= 2,5 min depois do carimbo de A): botao Digitar Resultados -> digita B onde o cursor parou, com
               DATA/HORA = carimbo de A + 1 min (o caso que era recusado); digita E na linha de D (carimbo de D + 1 min).
  FASE 2c (U2): recorta A:K de M para uma linha livre abaixo; digita N na origem (carimbo de M + 2 min).
  FASE 2d (U2): Ctrl+D (FillDown de A:K) de F para r0+5 e r0+6 e troca HORA e RESULTADO das duas de baixo.
  FASE 2e (U2): corrige o RESULTADO de G (manual de U1, ja integrado) na propria linha.
  FASE 2f (U2): ADR-082 -- numa linha livre digita primeiro o SUBSTITUI_ID (numero de um resultado real N1 do lote,
               ATIVO), depois RESULTADO e MOTIVO. Vizinho C11: outra linha com SUBSTITUI_ID de um real N2 e Delete de
               ID..MOTIVO -> o SUBSTITUI_ID deve sair junto.
  FASE 2g: vizinhos -- Delete de 2 linhas de uma vez (U2 digita X1 e X2 e apaga A:K das duas); carimbo orfao legado
               (eventos desligados) numa linha vazia -> a passada do ATUALIZAR deve limpar.
  ATUALIZAR #2 -> tela, tblCQ_Final (STATUS, MOTIVO, USUARIO_MANUAL), QA, Audit_Log, Painel/Estatistica/Eng_Saida.
Oraculo proprio (Python puro, statistics): n/media/DP/CV dos PARTICIPA = SIM da tblCQ_Final (Glicose, DIMENSION 1, lote,
nivel, nov/2026) e a conta a mao "antes - retirados - substituido + novos".

Uso: python -u reteste07_usuario_f1.py <nova|base>
  nova = copia da producao atual (raiz); base = copia do backup anterior a entrega 44
  (_backup_pre_integracao_2026-11-16_0959, entrega 43), para provar que o roteiro pega o defeito.
Regras: so COPIAS em %TEMP%/qcw/auditoria/2026-11-16_reteste07/; nada e salvo; fechar() encerra so o PID do QA;
valores sinteticos com 3 decimais (alvo do lote + k DP); MOTIVO "Outro".
"""
import datetime as dt
import os
import shutil
import sys
import time
import traceback

AQUI = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, AQUI)
import reteste06_usuario_f1_hema as R  # noqa: E402

C, V, t, up, num = R.C, R.V, R.t, R.up, R.num
PROD = 'Hematologia'
CICLO = '2026-11-16_reteste07'
PASTA = os.path.join(os.environ['TEMP'], 'qcw', 'auditoria', CICLO)
R.PASTA = C.PASTA = PASTA
BACKUP = os.path.join(C.REPO, '_backup_pre_integracao_2026-11-16_0959', f'QC_{PROD}.xlsm')
U1, U2 = 'TESTE_AUDITORIA_R7H', 'TESTE_AUDITORIA_R7HB'
EQ, AN = R.EQ, R.AN
MOTIVO = 'Outro'
HOJE = dt.date.today()


def fmt(x):
    return f'{x:.3f}'.replace('.', ',')


def dtx(v):
    if isinstance(v, dict):
        v = v.get('v')
    s = t(v)
    try:
        return dt.datetime.fromisoformat(s[:19])
    except Exception:  # noqa: BLE001
        return None


def carimbo(qa, r):
    c1 = R.cel(qa, r, 'USUARIO').Value
    c2 = R.cel(qa, r, 'REGISTRADO_EM').Value
    if hasattr(c2, 'tzinfo') and c2 is not None:
        c2 = c2.replace(tzinfo=None)
    return {'USUARIO': t(c1), 'REGISTRADO_EM': c2.isoformat(sep=' ') if hasattr(c2, 'isoformat') else t(c2)}


def estado(qa, r):
    d = {'ID': R.idd(qa, r)}
    d.update(carimbo(qa, r))
    try:
        d['SUBSTITUI_ID'] = t(R.cel(qa, r, 'SUBSTITUI_ID').Value)
    except Exception:  # noqa: BLE001
        d['SUBSTITUI_ID'] = '(sem coluna)'
    d['DATA'] = t(R.cel(qa, r, 'DATA').Text)
    d['HORA'] = t(R.cel(qa, r, 'HORA').Text)
    d['RESULTADO'] = t(R.cel(qa, r, 'RESULTADO').Text)
    return d


def fluxo(qa, L):
    fin0 = qa.ler('tblCQ_Final')
    L['tblCQ_Final_inicio'] = len(fin0)
    L['colunas_tblCQ_Final'] = list(fin0[0].keys()) if fin0 else []
    cand = sorted([r for r in fin0 if up(r.get('PARTICIPA_ESTATISTICA')) == 'SIM' and t(r.get('ANALITO')) == AN
                   and up(r.get('EQUIPAMENTO')) == EQ and isinstance(r.get('DATA_HORA'), dt.datetime)
                   and up(r.get('ORIGEM_RESULTADO')).startswith('INTERF')], key=lambda r: r['DATA_HORA'])
    lote, _ult = R.escolher_lote(qa, fin0)
    L['lote'] = lote
    L['ultimo_real'] = cand[-1]['DATA_HORA'].isoformat(sep=' ') if cand else None
    alvo = {nv: (qa.run('mLotes.MediaDoLote', lote, AN, nv, teto=60), qa.run('mLotes.DPDoLote', lote, AN, nv, teto=60))
            for nv in (1, 2, 3)}
    L['alvo_lote'] = {f'N{k}': v for k, v in alvo.items()}

    o_ini = R.oraculo(fin0, lote)

    def val(nv, k):
        m, s = alvo[nv]
        if not num(m) or not num(s) or not s:
            m, s = o_ini[nv - 1]['media'], o_ini[nv - 1]['dp']
        return round(m + k * s, 3)

    # resultados reais do lote em nov/2026 para o SUBSTITUI_ID (N1 vale; N2 e o vizinho C11)
    reais = [r for r in cand if t(r['LOTE']) == lote and up(r.get('STATUS_ANALITICO')) == 'ATIVO'
             and r['DATA_HORA'].year == 2026 and r['DATA_HORA'].month == 11]
    if not any(int(r['NIVEL']) == 1 for r in reais) or not any(int(r['NIVEL']) == 2 for r in reais):
        reais = [r for r in cand if t(r['LOTE']) == lote and up(r.get('STATUS_ANALITICO')) == 'ATIVO']
    sub1 = next(r for r in reversed(reais) if int(r['NIVEL']) == 1)
    sub2 = next(r for r in reversed(reais) if int(r['NIVEL']) == 2)
    L['substituidos'] = {k: {c: (x[c].isoformat(sep=' ') if isinstance(x.get(c), dt.datetime) else x.get(c))
                             for c in ('ID_REGISTRO', 'NIVEL', 'RESULTADO', 'DATA_HORA', 'STATUS_ANALITICO')}
                         for k, x in (('N1', sub1), ('N2', sub2))}
    dia = HOJE.strftime('%d/%m/%Y')

    def campos(hora, nv, v, data=dia):
        return [('DATA', data), ('HORA', hora), ('EQUIPAMENTO', EQ), ('LOTE', "'" + lote), ('NIVEL', str(nv)),
                ('ANALITO', AN), ('RESULTADO', fmt(v)), ('MOTIVO', MOTIVO)]

    o0 = R.oraculo(fin0, lote)
    p0 = R.painel(qa, lote, [], fin0)
    L['antes'] = {'oraculo': o0, 'painel': p0, 'comparacao': R.comparar(o0, p0)}

    qa.run('mIntegracao.IrDigitarResultados', teto=120)
    qa.ex.xl.EnableEvents = True
    qa.ex.xl.Interactive = True
    vig = R.Caixas(qa.ex.pid)
    vig.start()
    k0 = [0]
    passos = {}

    def passo(nome, d):
        d['caixas'] = vig.vistos[k0[0]:]
        k0[0] = len(vig.vistos)
        passos[nome] = d
        L['passos'] = passos

    def digitar(r, cs):
        for col, v in cs:
            R.escrever(qa, r, col, v)

    linhas = {}
    vals = {}
    # ------------------------------------------------------------------ FASE 1 (U1)
    try:
        r0 = R.livres(qa, 8)
        plano = [('A', r0, '03:01', 1, 0.1), ('D', r0 + 1, '03:02', 2, -0.1), ('M', r0 + 2, '03:03', 3, 0.2),
                 ('G', r0 + 3, '03:04', 1, -0.2), ('F', r0 + 4, '03:05', 1, 0.15), ('K', r0 + 7, '03:08', 2, 0.05)]
        f1 = {}
        for nome, r, h, nv, k in plano:
            v = val(nv, k)
            vals[nome] = (nv, v)
            digitar(r, campos(h, nv, v))
            linhas[nome] = r
            f1[nome] = estado(qa, r)
        passo('F1_digitacao_U1', {'r0': r0, 'linhas': f1})
    finally:
        qa.ex.xl.Interactive = False
    L['linhas'] = dict(linhas)
    aud = V.Audit(qa)
    msg1, s1 = qa.atualizar()
    fin1, d1 = R.depois(qa, {k: r for k, r in linhas.items()}, msg1)
    d1['tblCQ_Final_meus'] = meus(fin1, {R.idd(qa, r) for r in linhas.values()})
    d1['s'] = round(s1, 1)
    d1['auditoria'] = aud.novos()
    L['atualizar1'] = d1
    ids1 = {k: R.idd(qa, r) for k, r in linhas.items()}
    L['ids_fase1'] = ids1
    T0 = {k: dtx(carimbo(qa, r)['REGISTRADO_EM']) for k, r in linhas.items()}
    L['T0'] = {k: v.isoformat(sep=' ') if v else None for k, v in T0.items()}

    # ------------------------------------------------------------------ FASE 2a (U1 apaga)
    qa.run('mIntegracao.IrDigitarResultados', teto=120)
    qa.ex.xl.EnableEvents = True
    qa.ex.xl.Interactive = True
    try:
        rA, rD = linhas['A'], linhas['D']
        antesA, antesD = estado(qa, rA), estado(qa, rD)
        R.faixa(qa, rA, rA).ClearContents()                                  # Delete de ID..MOTIVO
        qa.ex.esperar()
        time.sleep(0.4)
        R.faixa(qa, rD, rD, 'DATA', 'MOTIVO').ClearContents()                # Delete so de DATA..MOTIVO (ID fica fora)
        qa.ex.esperar()
        time.sleep(0.4)
        passo('F2a_apagar_U1', {'A_antes': antesA, 'A_depois_delete_ID_MOTIVO': estado(qa, rA),
                                'D_antes': antesD, 'D_depois_delete_DATA_MOTIVO': estado(qa, rD),
                                'marca': R.marca(qa)})
    finally:
        qa.ex.xl.Interactive = False

    # ------------------------------------------------------------------ FASE 2b (U2 digita na linha apagada)
    L['login_U2'] = R.fazer_login(qa, U2)
    falta = (T0['A'] + dt.timedelta(minutes=2, seconds=30) - dt.datetime.now()).total_seconds()
    if falta > 0:
        time.sleep(falta)
    qa.run('mIntegracao.IrDigitarResultados', teto=120)
    botao = {'aba': t(qa.ex.xl.ActiveSheet.Name), 'celula': t(qa.ex.xl.ActiveCell.Address),
             'linha_da_tabela': qa.ex.xl.ActiveCell.Row - R.lo_m(qa).DataBodyRange.Row + 1}
    qa.ex.xl.EnableEvents = True
    qa.ex.xl.Interactive = True
    try:
        rB = botao['linha_da_tabela']
        hB = (T0['A'] + dt.timedelta(minutes=1)).replace(second=0)
        if hB <= T0['A']:
            hB += dt.timedelta(minutes=1)
        vals['B'] = (1, val(1, -0.15))
        agoraB = dt.datetime.now()
        digitar(rB, campos(hB.strftime('%H:%M'), 1, vals['B'][1], hB.strftime('%d/%m/%Y')))
        linhas['B'] = rB
        eB = estado(qa, rB)
        rDl = linhas['D']
        hE = (T0['D'] + dt.timedelta(minutes=1)).replace(second=0)
        if hE <= T0['D']:
            hE += dt.timedelta(minutes=1)
        vals['E'] = (2, val(2, 0.12))
        agoraE = dt.datetime.now()
        digitar(rDl, campos(hE.strftime('%H:%M'), 2, vals['E'][1], hE.strftime('%d/%m/%Y')))
        linhas['E'] = rDl
        passo('F2b_digitar_na_linha_apagada_U2', {'botao_parou_em': botao, 'B': {'linha': rB, 'data_hora': hB.isoformat(sep=' '),
                                                  'momento': agoraB.isoformat(sep=' ', timespec='seconds'), 'tela': eB},
                                                  'E': {'linha': rDl, 'data_hora': hE.isoformat(sep=' '),
                                                        'momento': agoraE.isoformat(sep=' ', timespec='seconds'),
                                                        'tela': estado(qa, rDl)}, 'marca': R.marca(qa)})

        # -------------------------------------------------------------- FASE 2c (recorte)
        rM = linhas['M']
        antesM = estado(qa, rM)
        dM = R.livres(qa, 1, r0 + 8 + 2)
        R.faixa(qa, rM, rM).Cut(R.faixa(qa, dM, dM))
        qa.ex.esperar()
        time.sleep(0.5)
        origem_logo = estado(qa, rM)
        destino_logo = estado(qa, dM)
        hN = (T0['M'] + dt.timedelta(minutes=2)).replace(second=0)
        vals['N'] = (3, val(3, -0.22))
        agoraN = dt.datetime.now()
        digitar(rM, campos(hN.strftime('%H:%M'), 3, vals['N'][1], hN.strftime('%d/%m/%Y')))
        linhas['N'] = rM
        linhas['M'] = dM
        passo('F2c_recorte', {'de': rM, 'para': dM, 'M_antes': antesM, 'origem_logo_apos_recorte': origem_logo,
                              'destino_logo_apos_recorte': destino_logo, 'N_na_origem': {
                                  'data_hora': hN.isoformat(sep=' '), 'momento': agoraN.isoformat(sep=' ', timespec='seconds'),
                                  'tela': estado(qa, rM)}})

        # -------------------------------------------------------------- FASE 2d (Ctrl+D)
        rF = linhas['F']
        antesF = estado(qa, rF)
        R.faixa(qa, rF, rF + 2).FillDown()
        qa.ex.esperar()
        time.sleep(0.6)
        logo = [estado(qa, rF + k) for k in range(3)]
        vals['F1'] = (1, val(1, 0.25))
        vals['F2'] = (1, val(1, -0.25))
        R.escrever(qa, rF + 1, 'HORA', '03:21')
        R.escrever(qa, rF + 1, 'RESULTADO', fmt(vals['F1'][1]))
        R.escrever(qa, rF + 2, 'HORA', '03:22')
        R.escrever(qa, rF + 2, 'RESULTADO', fmt(vals['F2'][1]))
        linhas['F1'], linhas['F2'] = rF + 1, rF + 2
        passo('F2d_ctrl_d', {'origem_antes': antesF, 'logo_apos_filldown': logo,
                             'depois_edicao': [estado(qa, rF + k) for k in range(3)]})

        # -------------------------------------------------------------- FASE 2e (editar G)
        rG = linhas['G']
        antesG = estado(qa, rG)
        vals['G_novo'] = (1, val(1, -0.3))
        R.escrever(qa, rG, 'RESULTADO', fmt(vals['G_novo'][1]))
        passo('F2e_editar_G', {'antes': antesG, 'depois': estado(qa, rG), 'tela': R.ler(qa, rG)})

        # -------------------------------------------------------------- FASE 2f (ADR-082)
        rS = R.livres(qa, 1, dM)
        num1 = t(sub1['ID_REGISTRO']).split('-')[-1]
        R.escrever(qa, rS, 'SUBSTITUI_ID', num1)
        apos_subst = estado(qa, rS)
        apos_subst_tela = R.ler(qa, rS)
        vals['S'] = (1, round(float(sub1['RESULTADO']) + 0.123, 3))
        R.escrever(qa, rS, 'RESULTADO', fmt(vals['S'][1]))
        R.escrever(qa, rS, 'MOTIVO', MOTIVO)
        linhas['S'] = rS
        rS2 = R.livres(qa, 1, rS)
        num2 = t(sub2['ID_REGISTRO']).split('-')[-1]
        R.escrever(qa, rS2, 'SUBSTITUI_ID', num2)
        s2_pre = estado(qa, rS2)
        R.escrever(qa, rS2, 'RESULTADO', fmt(round(float(sub2['RESULTADO']) + 0.321, 3)))
        s2_antes = estado(qa, rS2)
        R.faixa(qa, rS2, rS2).ClearContents()                                # Delete de ID..MOTIVO (SUBSTITUI_ID no fim)
        qa.ex.esperar()
        time.sleep(0.4)
        passo('F2f_substitui', {'S': {'linha': rS, 'digitado': num1, 'logo_apos_SUBSTITUI_ID': apos_subst,
                                      'tela_logo': apos_subst_tela, 'depois': estado(qa, rS)},
                                'S2_C11': {'linha': rS2, 'digitado': num2, 'apos_subst': s2_pre, 'antes_delete': s2_antes,
                                           'depois_delete_ID_MOTIVO': estado(qa, rS2)}})

        # -------------------------------------------------------------- FASE 2g (vizinhos)
        rX = R.livres(qa, 2, rS2)
        digitar(rX, campos('03:40', 1, val(1, 0.0)))
        digitar(rX + 1, campos('03:41', 2, val(2, 0.0)))
        xantes = [estado(qa, rX), estado(qa, rX + 1)]
        R.faixa(qa, rX, rX + 1).ClearContents()
        qa.ex.esperar()
        time.sleep(0.4)
        xdepois = [estado(qa, rX), estado(qa, rX + 1)]
        # orfao legado (eventos desligados, como um arquivo antigo)
        rO = R.livres(qa, 1, rX + 2)
        qa.ex.xl.EnableEvents = False
        R.cel(qa, rO, 'USUARIO').Value = 'TESTE_AUDITORIA_ORFAO'
        R.cel(qa, rO, 'REGISTRADO_EM').Value = dt.datetime(2026, 11, 1, 6, 0, 0)
        qa.ex.xl.EnableEvents = True
        passo('F2g_vizinhos', {'delete_2_linhas': {'linhas': [rX, rX + 1], 'antes': xantes, 'depois': xdepois},
                               'orfao': {'linha': rO, 'antes_atualizar': estado(qa, rO)}})
        linhas['_orfao'] = rO
    finally:
        qa.ex.xl.Interactive = False
        qa.ex.xl.EnableEvents = True
    L['linhas'] = dict(linhas)
    L['valores'] = vals
    L['tela_antes_atualizar2'] = {k: estado(qa, r) for k, r in linhas.items()}

    # ------------------------------------------------------------------ ATUALIZAR #2
    aud2 = V.Audit(qa)
    msg2, s2 = qa.atualizar()
    vig.parar.set()
    lin_tela = {k: r for k, r in linhas.items() if not k.startswith('_')}
    fin2, d2 = R.depois(qa, lin_tela, msg2)
    d2['s'] = round(s2, 1)
    d2['auditoria'] = aud2.novos()
    d2['orfao_depois'] = estado(qa, linhas['_orfao'])
    d2['carimbo_tela'] = {k: estado(qa, r) for k, r in lin_tela.items()}
    ids_todos = set(ids1.values()) | {t(v['ID']) for v in d2['carimbo_tela'].values()}
    ids_todos |= {t(sub1['ID_REGISTRO']), t(sub2['ID_REGISTRO'])}
    cols = [c for c in (fin2[0].keys() if fin2 else []) if c in (
        'ID_REGISTRO', 'STATUS_ANALITICO', 'PARTICIPA_ESTATISTICA', 'NIVEL', 'RESULTADO', 'DATA_HORA', 'RUN',
        'MOTIVO_EXCLUSAO_AUTOMATICA', 'USUARIO_MANUAL', 'REGISTRADO_EM', 'SUBSTITUI_ID', 'SUBSTITUIDO_POR')
        or 'USUARIO' in c or 'REGISTR' in c or 'SUBST' in c]
    d2['tblCQ_Final_meus'] = [{c: (r[c].isoformat(sep=' ') if isinstance(r.get(c), dt.datetime) else r.get(c)) for c in cols}
                              for r in fin2 if t(r['ID_REGISTRO']).split('~')[0] in ids_todos]
    d2['ids_fase1_ausentes'] = [i for i in ids1.values() if not any(t(r['ID_REGISTRO']).split('~')[0] == i for r in fin2)]
    qai = []
    for x in qa.ler('tblQA_Integracao'):
        s = ' | '.join(t(v) for v in x.values() if v not in (None, ''))
        if any(i and i in s for i in ids_todos) or 'posterior' in s:
            qai.append(s[:600])
    d2['qa_integracao'] = qai
    L['atualizar2'] = d2
    p2 = R.painel(qa, lote, sorted(i for i in ids_todos if i), fin2)
    o2 = R.oraculo(fin2, lote)
    L['depois2'] = {'oraculo': o2, 'painel': p2, 'comparacao': R.comparar(o2, p2)}


def meus(fin, ids):
    cols = ('ID_REGISTRO', 'STATUS_ANALITICO', 'PARTICIPA_ESTATISTICA', 'NIVEL', 'RESULTADO', 'DATA_HORA', 'RUN',
            'MOTIVO_EXCLUSAO_AUTOMATICA', 'USUARIO_MANUAL')
    return [{c: (r[c].isoformat(sep=' ') if isinstance(r.get(c), dt.datetime) else r.get(c)) for c in cols if c in r}
            for r in fin if t(r['ID_REGISTRO']).split('~')[0] in ids]


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
        R.criar_usuario(qa, U1, 'ANALISTA')
        R.criar_usuario(qa, U2, 'ANALISTA')
        L['login_U1'] = R.fazer_login(qa, U1)
        senha = C.senha_planilha()
        for w in ('Painel', 'Estatística'):
            try:
                qa.ws(w).Unprotect(senha)
            except Exception:  # noqa: BLE001
                pass
        fluxo(qa, L)
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
