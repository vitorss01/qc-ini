# -*- coding: utf-8 -*-
"""repro_p01_carimbo_residual.py -- Tester/QA, ciclo 2026-11-15_reteste06, modo REPRODUZIR da P-01 (Bio R6-01 e
Hema R6-01): apagar ou mover uma linha manual deixa USUARIO e REGISTRADO_EM e o proximo manual digitado nela herda
o carimbo (recusado por "DATA/HORA posterior ao lancamento" e atribuido ao usuario antigo).

Escrito pelo Tester/QA. NAO reaproveita reteste06_carimbo_residual.py nem reteste06_usuario_f1*.py (Usuario Simulado).
Reaproveita so infraestrutura: qa_final.QA (abre a copia, revela as abas, imprime EXCEL_PID), repro_comum (meu),
Vigia (meu, le e fecha caixas do MEU PID, sem imagem) e qa_inativar_ux.criar_usuario/fazer_login (usuario so na copia).

Diferencas de roteiro em relacao ao Usuario Simulado (ampliacao):
  * dois usuarios: T1 digita e apaga; T2 (outro login) digita na linha apagada -> mostra a atribuicao errada;
  * quatro linhas de origem: L0 (apagar ID..MOTIVO, depois digitar B com DATA/HORA > carimbo), L1 (recortar
    ID..MOTIVO e colar adiante; digitar B2 na origem, OUTRO analito e OUTRO nivel), L2 (apagar; digitar D com DATA/HORA
    ANTERIOR ao carimbo: entra ATIVO, mas atribuido a T1), L3 (apagar e deixar vazia: o ATUALIZAR/PrepararEntradas
    limpa o resto?);
  * controle C numa linha nunca usada, mesmo analito/nivel/DATA/HORA do B;
  * oraculo proprio: regra do motor (mEtlCQ: DATA/HORA > REGISTRADO_EM -> MANUAL_INCOMPLETO) aplicada em Python aos
    valores lidos da tela, comparada com o que o motor gravou.

Uso: python -u repro_p01_carimbo_residual.py <Bioquimica|Hematologia>
Regras: so COPIA nova em %TEMP%/qcw/auditoria/2026-11-15_reteste06/tester/; nada e salvo; fechar() encerra so o PID
que o QA criou; nenhuma captura de tela. Dados de teste: usuarios TESTE_AUDITORIA_R6T1 e TESTE_AUDITORIA_R6T2
(ANALISTA, so na copia), OBS/COMENTARIO "TESTE-AUDITORIA P01", valores sinteticos de 3 decimais (media recente do
lote + 0,001..0,004), DATAs 14/11/2026 e 15/11/2026.
"""
import datetime as dt
import json
import os
import re
import shutil
import sys
import time
import traceback

AQUI = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, AQUI)
import repro_comum as C  # noqa: E402
from valida_adr090_091 import Vigia  # noqa: E402
sys.path.insert(0, C.TESTES)
from qa_inativar_ux import criar_usuario, fazer_login  # noqa: E402

PROD = sys.argv[1]
BIO = PROD == 'Bioquimica'
CICLO = '2026-11-15_reteste06'
C.PASTA = os.path.join(os.environ['TEMP'], 'qcw', 'auditoria', CICLO, 'tester')
T1, T2 = 'TESTE_AUDITORIA_R6T1', 'TESTE_AUDITORIA_R6T2'
EPOCA = dt.datetime(1899, 12, 30)
SEGREDOS = [x for x in (os.environ.get('USERNAME'), os.environ.get('COMPUTERNAME')) if x and len(x) > 3]


def t(v):
    if isinstance(v, float) and v == int(v):
        return str(int(v))
    return '' if v is None else str(v).strip()


def num(v):
    return isinstance(v, (int, float)) and not isinstance(v, bool)


def dser(x):
    return (EPOCA + dt.timedelta(days=float(x))) if num(x) else None


def fmt(x):
    d = dser(x)
    return d.strftime('%d/%m/%Y %H:%M:%S') if d else t(x)


class Log(C.Log):
    def gravar(self):
        s = json.dumps(self.R, ensure_ascii=False, indent=1, default=str)
        for x in SEGREDOS:
            s = re.sub(re.escape(x), '<omitido>', s, flags=re.I)
        open(self.saida, 'w', encoding='utf-8').write(s)


# ------------------------------------------------------------------------------------------------ tabela manual
def cols(qa):
    lo = qa.lo('tblResultados_Manuais')
    return lo, [str(c.Name) for c in lo.ListColumns]


def cel(qa, r, col):
    lo, _ = cols(qa)
    return lo.ListColumns(col).DataBodyRange.Cells(r, 1)


def linha(qa, r):
    lo, cab = cols(qa)
    rg = lo.ListRows(r).Range
    v2 = rg.Value2[0]
    out = {'_planilha': rg.Row}
    for i, c in enumerate(cab):
        x = v2[i]
        if c in ('REGISTRADO_EM',) or (c in ('DATA', 'HORA') and num(x)):
            out[c] = x
            out[c + '_txt'] = t(rg.Cells(1, i + 1).Text)
        else:
            out[c] = x if num(x) else t(x)
    return out


def escrever(qa, r, col, v):
    c = cel(qa, r, col)
    c.FormulaLocal = v
    qa.ex.esperar()
    time.sleep(0.25)


def digitar(qa, r, campos, obs):
    for col, v in campos:
        escrever(qa, r, col, v)
    if obs:
        escrever(qa, r, obs, 'TESTE-AUDITORIA P01')
    return linha(qa, r)


def faixa_id_motivo(qa, r):
    lo, cab = cols(qa)
    a = cel(qa, r, 'ID_REGISTRO')
    b = cel(qa, r, 'MOTIVO')
    return a.Worksheet.Range(a, b)


# ------------------------------------------------------------------------------------------------ dados de referencia
def final(qa):
    lo = qa.lo('tblCQ_Final')
    cab = [str(c.Name) for c in lo.ListColumns]
    v = lo.DataBodyRange.Value2
    return cab, [dict(zip(cab, l)) for l in v]


def ref(rows, analito, nivel, equip):
    c = [r for r in rows if t(r.get('ANALITO')).upper() == analito.upper() and t(r.get('NIVEL')) == str(nivel)
         and t(r.get('PARTICIPA_ESTATISTICA')).upper() == 'SIM' and num(r.get('RESULTADO'))
         and num(r.get('DATA_HORA')) and t(r.get('ORIGEM_RESULTADO')).upper().startswith('INTERF')
         and (not equip or t(r.get('EQUIPAMENTO')).upper() == equip.upper())]
    c.sort(key=lambda r: r['DATA_HORA'])
    if not c:
        return None
    lote = t(c[-1]['LOTE'])
    vals = [r['RESULTADO'] for r in c if t(r['LOTE']) == lote][-20:]
    return {'analito': analito, 'nivel': nivel, 'lote': lote, 'ultimo': fmt(c[-1]['DATA_HORA']),
            'media_recente': sum(vals) / len(vals), 'n_ref': len(vals)}


def outro_analito(rows, principal, equip):
    cont = {}
    for r in rows:
        an = t(r.get('ANALITO'))
        if an.upper() == principal.upper() or t(r.get('PARTICIPA_ESTATISTICA')).upper() != 'SIM':
            continue
        if equip and t(r.get('EQUIPAMENTO')).upper() != equip.upper():
            continue
        if not num(r.get('DATA_HORA')) or not t(r.get('ORIGEM_RESULTADO')).upper().startswith('INTERF'):
            continue
        cont[an] = max(cont.get(an, 0), r['DATA_HORA'])
    return sorted(cont, key=lambda a: -cont[a])


def dec(x):
    return f'{x:.3f}'.replace('.', ',')


def campos_de(rf, data, hora, valor, equip):
    c = [('DATA', data), ('HORA', hora)]
    if equip:
        c.append(('EQUIPAMENTO', equip))
    c += [('LOTE', "'" + rf['lote']), ('NIVEL', str(rf['nivel'])), ('ANALITO', rf['analito']),
          ('RESULTADO', dec(valor)), ('MOTIVO', 'Outro')]
    return c


def oraculo(lin):
    """Regra do motor (mEtlCQ, linha 'If Ms(dh) > Ms(regc)'): DATA+HORA > REGISTRADO_EM -> incompleto."""
    d, h, reg = lin.get('DATA'), lin.get('HORA'), lin.get('REGISTRADO_EM')
    if not (num(d) and num(h) and num(reg)):
        return 'indeterminado'
    dh = int(d) + (h - int(h))
    return 'RECUSA (DATA/HORA posterior ao lancamento)' if round(dh * 86400000) > round(reg * 86400000) else 'ACEITA'


# ------------------------------------------------------------------------------------------------ roteiro
def roteiro(qa, L):
    out = {}
    lo, cab = cols(qa)
    out['colunas'] = cab
    obs = next((c for c in cab if any(k in c.upper() for k in ('OBS', 'COMENT'))), None)
    ws = lo.Parent
    out['aba'] = {'nome': ws.Name, 'ProtectContents': bool(ws.ProtectContents)}
    eq = 'DIMENSION 1' if BIO else ''
    _, rows = final(qa)
    pri = 'Glicose' if BIO else 'WBC'
    seg = outro_analito(rows, pri, eq)
    nseg = 2 if BIO else 3
    rA2 = next(x for x in (ref(rows, a, nseg, eq) for a in seg) if x)
    R = {'A': ref(rows, pri, 1, eq), 'A2': rA2, 'A3': ref(rows, pri, 2, eq)}
    out['referencias'] = R
    # travamento das colunas (o que o usuario consegue apagar)
    r_livre = C.linha_livre_manual(qa)[1]
    out['travadas'] = {c: bool(cel(qa, r_livre, c).Locked) for c in cab}

    # --- T1 digita A, A2, A3, A4 em 4 linhas seguidas, a partir da linha onde o botao leva
    fazer_login(qa, T1)
    out['login_T1'] = t(qa.nome('currentUser').Value)
    qa.run('mIntegracao.IrDigitarResultados', teto=120)
    ac = qa.ex.xl.ActiveCell
    r0 = ac.Row - lo.DataBodyRange.Row + 1
    out['botao_1'] = {'aba': qa.ex.xl.ActiveSheet.Name, 'celula': ac.Address.replace('$', ''), 'linha_tabela': r0}
    qa.ex.xl.EnableEvents = True
    vig = Vigia(qa.ex.pid, teto=900)
    vig.start()
    orig = {}
    orig['A'] = digitar(qa, r0, campos_de(R['A'], '14/11/2026', '07:40:00', R['A']['media_recente'] + 0.001, eq), obs)
    orig['A2'] = digitar(qa, r0 + 1, campos_de(R['A2'], '14/11/2026', '07:41:00', R['A2']['media_recente'] + 0.002,
                                               eq), obs)
    orig['A3'] = digitar(qa, r0 + 2, campos_de(R['A3'], '14/11/2026', '07:42:00', R['A3']['media_recente'] + 0.003,
                                               eq), obs)
    orig['A4'] = digitar(qa, r0 + 3, campos_de(R['A'], '14/11/2026', '07:43:00', R['A']['media_recente'] + 0.004,
                                               eq), obs)
    out['origem_digitada'] = orig
    # linhas de controle e destino do recorte: nunca usadas
    rC, rDest = r0 + 6, r0 + 9
    out['controle_antes'] = linha(qa, rC)
    out['destino_antes'] = linha(qa, rDest)

    # --- apagar ID..MOTIVO (L0, L2, L3) e recortar ID..MOTIVO de L1 para rDest, com a aba protegida e eventos ligados
    acoes = {}
    for rot, r in (('L0_apagar', r0), ('L2_apagar', r0 + 2), ('L3_apagar_e_deixar', r0 + 3)):
        try:
            rg = faixa_id_motivo(qa, r)
            acoes[rot] = {'faixa': rg.Address.replace('$', '')}
            rg.ClearContents()
            qa.ex.esperar()
            time.sleep(0.3)
            acoes[rot]['ok'] = True
        except Exception as e:  # noqa: BLE001
            acoes[rot] = {'ERRO': repr(e)}
        acoes[rot]['linha_depois'] = linha(qa, r)
    try:
        rg = faixa_id_motivo(qa, r0 + 1)
        dest = cel(qa, rDest, 'ID_REGISTRO')
        acoes['L1_recortar'] = {'faixa': rg.Address.replace('$', ''), 'destino': dest.Address.replace('$', '')}
        rg.Cut(dest)
        qa.ex.esperar()
        time.sleep(0.5)
        qa.ex.xl.CutCopyMode = False
        acoes['L1_recortar']['ok'] = True
    except Exception as e:  # noqa: BLE001
        acoes['L1_recortar'] = {'ERRO': repr(e)}
    acoes['L1_recortar']['origem_depois'] = linha(qa, r0 + 1)
    acoes['L1_recortar']['destino_depois'] = linha(qa, rDest)
    out['acoes'] = acoes

    # --- espera: o B tem de ter DATA/HORA depois do carimbo de A e antes de agora
    regA = orig['A']['REGISTRADO_EM']
    regA2 = orig['A2']['REGISTRADO_EM']
    alvo = dser(regA) + dt.timedelta(seconds=75)
    while dt.datetime.now() < alvo:
        time.sleep(2)

    # --- T2 (outro usuario) entra e clica no botao Digitar Resultados
    fazer_login(qa, T2)
    out['login_T2'] = t(qa.nome('currentUser').Value)
    qa.run('mIntegracao.IrDigitarResultados', teto=120)
    ac = qa.ex.xl.ActiveCell
    out['botao_2'] = {'aba': qa.ex.xl.ActiveSheet.Name, 'celula': ac.Address.replace('$', ''),
                      'linha_tabela': ac.Row - lo.DataBodyRange.Row + 1, 'e_a_linha_apagada_L0': ac.Row - lo.DataBodyRange.Row + 1 == r0}
    qa.ex.xl.EnableEvents = True
    hB = (dser(regA) + dt.timedelta(seconds=40))
    hB2 = (dser(regA2) + dt.timedelta(seconds=40))
    novos = {}
    out['inicio_digitacao_T2'] = dt.datetime.now().strftime('%d/%m/%Y %H:%M:%S')
    novos['B_em_L0'] = digitar(qa, r0, campos_de(R['A'], hB.strftime('%d/%m/%Y'), hB.strftime('%H:%M:%S'),
                                                  R['A']['media_recente'] - 0.001, eq), obs)
    novos['B2_em_L1_origem_do_recorte'] = digitar(qa, r0 + 1, campos_de(R['A2'], hB2.strftime('%d/%m/%Y'),
                                                                        hB2.strftime('%H:%M:%S'),
                                                                        R['A2']['media_recente'] - 0.002, eq), obs)
    novos['D_em_L2_hora_anterior'] = digitar(qa, r0 + 2, campos_de(R['A3'], '15/11/2026', '06:00:00',
                                                                   R['A3']['media_recente'] - 0.003, eq), obs)
    novos['C_controle_linha_nova'] = digitar(qa, rC, campos_de(R['A'], hB.strftime('%d/%m/%Y'), hB.strftime('%H:%M:%S'),
                                                               R['A']['media_recente'] - 0.004, eq), obs)
    novos['L3_vazia'] = linha(qa, r0 + 3)
    out['fim_digitacao_T2'] = dt.datetime.now().strftime('%d/%m/%Y %H:%M:%S')
    vig.parar.set()
    out['caixas_na_digitacao'] = vig.vistos
    out['antes_atualizar'] = novos
    out['oraculo'] = {k: oraculo(v) for k, v in novos.items() if v.get('DATA')}
    out['oraculo']['destino_recorte'] = oraculo(linha(qa, rDest))
    out['heranca'] = {k: {'USUARIO': v.get('USUARIO'), 'REGISTRADO_EM': v.get('REGISTRADO_EM_txt'),
                          'usuario_logado': out['login_T2'], 'herdou_usuario': v.get('USUARIO') == T1}
                      for k, v in novos.items()}

    # --- ATUALIZAR DADOS (nucleo do botao)
    vig2 = Vigia(qa.ex.pid, teto=1800)
    vig2.start()
    t0 = time.time()
    try:
        msg, s = qa.atualizar()
        out['atualizar'] = {'msg': msg[:2500], 's': round(s, 1)}
    except Exception as e:  # noqa: BLE001
        out['atualizar'] = {'ERRO': repr(e), 's': round(time.time() - t0, 1)}
    vig2.parar.set()
    out['caixas_atualizar'] = vig2.vistos
    rr = {'B_em_L0': r0, 'B2_em_L1_origem_do_recorte': r0 + 1, 'D_em_L2_hora_anterior': r0 + 2, 'L3_vazia': r0 + 3,
          'C_controle_linha_nova': rC, 'A2_movido_destino': rDest}
    out['depois_atualizar'] = {k: linha(qa, r) for k, r in rr.items()}
    ids = [t(v.get('ID_REGISTRO')) for v in out['depois_atualizar'].values() if t(v.get('ID_REGISTRO'))]
    out['ids'] = ids
    cabF, rowsF = final(qa)
    keep = [c for c in ('ID_REGISTRO', 'ANALITO', 'NIVEL', 'LOTE', 'EQUIPAMENTO', 'RUN', 'DATA_HORA', 'RESULTADO',
                        'PARTICIPA_ESTATISTICA', 'TIPO_PLOTAGEM_LJ', 'STATUS_ANALITICO', 'ORIGEM_RESULTADO',
                        'MOTIVO_EXCLUSAO_AUTOMATICA', 'USUARIO_MANUAL', 'REGISTRADO_EM') if c in cabF]
    out['tblCQ_Final'] = [{c: (fmt(r[c]) if c in ('DATA_HORA', 'REGISTRADO_EM') else r[c]) for c in keep}
                          for r in rowsF if any(t(r.get('ID_REGISTRO')).startswith(i) for i in ids)]
    try:
        out['QA_Integracao'] = [{k: t(v)[:300] for k, v in x.items() if v not in (None, '')}
                                for x in qa.ler('tblQA_Integracao')
                                if any(i in ' '.join(t(v) for v in x.values()) for i in ids)][:20]
    except Exception as e:  # noqa: BLE001
        out['QA_Integracao'] = repr(e)
    aud = {}
    for w in qa.wb.Worksheets:
        for l in w.ListObjects:
            if 'AUDIT' in l.Name.upper():
                try:
                    cabA = [str(c.Name) for c in l.ListColumns]
                    vA = l.DataBodyRange.Value2 if l.DataBodyRange is not None else []
                    hit = []
                    for x in vA[-400:]:
                        s = ' '.join(t(y) for y in x)
                        if any(i in s for i in ids) or 'TESTE_AUDITORIA_R6T' in s:
                            hit.append({cabA[j]: (fmt(y) if num(y) and 40000 < y < 60000 else t(y)[:250])
                                        for j, y in enumerate(x) if y not in (None, '')})
                    aud[l.Name] = hit[-25:]
                except Exception as e:  # noqa: BLE001
                    aud[l.Name] = repr(e)
    out['auditoria'] = aud
    return out


def main():
    os.makedirs(C.PASTA, exist_ok=True)
    src = os.path.join(C.REPO, f'QC_{PROD}.xlsm')
    h0 = C.sha(src)
    copia = os.path.join(C.PASTA, f'{PROD}_P01_{time.strftime("%H%M%S")}.xlsm')
    shutil.copy2(src, copia)
    L = Log(os.path.join(C.PASTA, f'repro_P01_{PROD}.json'))
    L['ambiente'] = {'producao': os.path.basename(src), 'sha_prod_12': h0, 'copia': os.path.basename(copia),
                     'sha_copia_12': C.sha(copia)}
    qa = C.qa_final.QA(PROD, copia)
    L['excel_pid'] = qa.ex.pid
    try:
        criar_usuario(qa, T1, 'ANALISTA')
        criar_usuario(qa, T2, 'ANALISTA')
        qa.ex.xl.EnableEvents = True
        try:
            L['roteiro'] = roteiro(qa, L)
        except Exception as e:  # noqa: BLE001
            L['ERRO'] = {'erro': repr(e), 'trace': traceback.format_exc()[-3000:]}
    finally:
        try:
            qa.fechar(salvar=False)
        except Exception:  # noqa: BLE001
            pass
    L['producao_intacta'] = {'sha_antes': h0, 'sha_depois': C.sha(src), 'igual': h0 == C.sha(src)}
    try:
        os.remove(copia)
        L['copia_apagada'] = True
    except Exception as e:  # noqa: BLE001
        L['copia_apagada'] = repr(e)
    print('FIM', PROD, flush=True)


if __name__ == '__main__':
    t0 = time.time()
    main()
    print(f'tempo {time.time() - t0:.0f}s', flush=True)
