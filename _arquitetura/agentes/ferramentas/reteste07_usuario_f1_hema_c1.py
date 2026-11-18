# -*- coding: utf-8 -*-
"""reteste07_usuario_f1_hema_c1.py -- Usuario Simulado, ciclo 2026-11-16_reteste07, F1 Hematologia, COMPLEMENTO da OC-0078.

Por que existe: no roteiro principal (reteste07_usuario_f1_hema.py, copia de 17/11) o manual B do criterio 1 (Delete de
ID..MOTIVO e manual novo na mesma linha) ganhou ID e carimbo novos, mas ficou CONFLITO_MANUAL "coincide com resultado
do interfaceamento (30 min)" porque a DATA/HORA escolhida (carimbo antigo + 1 min = 08:40) caiu a 29,6 min de um
resultado real N1 do dia (HEM-379943, 08:10). E a regra correta do motor, nao o defeito. Este complemento refaz SO o
criterio 1 escolhendo o NIVEL cuja DATA/HORA alvo fica a mais de 30 min de qualquer resultado real do dia, para provar
o "fica ATIVO e entra na Estatistica/LJ". Fluxo vizinho do ADR-097 §4: "apagar so o ID de uma linha com dado gera ID
novo e mantem o carimbo" (C04 do qa_adr097), que o roteiro principal nao fez.

Fluxo (aba Digitar Resultados, eventos ligados, Interactive = True na digitacao; caixas lidas pelo texto e fechadas):
  U1 digita A (nivel escolhido, hoje 03:11) e C (N2, hoje 03:12). ATUALIZAR #1.
  U1 apaga A (Delete de ID..MOTIVO). U1 apaga so o ID de C (C04).
  U2 (>= 2,5 min depois do carimbo de A): botao Digitar Resultados -> digita B onde o cursor parou, DATA/HORA =
     carimbo de A + 1 min. ATUALIZAR #2.
Leituras: tela, tblCQ_Final (lista completa do filtro WBC/XN1000/lote/nov antes e depois, para a conta a mao),
Audit_Log novo, QA, Painel/Estatistica/Eng_Saida e o oraculo (statistics).
Uso: python -u reteste07_usuario_f1_hema_c1.py   (copia NOVA da producao; nada salvo; fecha so o PID do QA)
"""
import datetime as dt
import os
import shutil
import sys
import time
import traceback

AQUI = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, AQUI)
import reteste07_usuario_f1_hema as M  # noqa: E402

R, C, V, t, up, num = M.R, M.C, M.V, M.t, M.up, M.num
PROD, PASTA, EQ, AN = M.PROD, M.PASTA, M.EQ, M.AN
U1, U2 = 'TESTE_AUDITORIA_R7HC', 'TESTE_AUDITORIA_R7HD'


def iso(v):
    return v.isoformat(sep=' ') if isinstance(v, dt.datetime) else v


def filtro(fin, lote):
    out = []
    for r in fin:
        if t(r.get('ANALITO')) == AN and up(r.get('EQUIPAMENTO')) == EQ and t(r.get('LOTE')) == lote \
                and isinstance(r.get('DATA_HORA'), dt.datetime) and r['DATA_HORA'].year == 2026 and r['DATA_HORA'].month == 11:
            out.append({c: iso(r.get(c)) for c in ('ID_REGISTRO', 'NIVEL', 'RESULTADO', 'DATA_HORA', 'STATUS_ANALITICO',
                                                   'PARTICIPA_ESTATISTICA', 'RUN', 'USUARIO_MANUAL',
                                                   'MOTIVO_EXCLUSAO_AUTOMATICA')})
    return out


def fluxo(qa, L):
    fin0 = qa.ler('tblCQ_Final')
    lote, _ = R.escolher_lote(qa, fin0)
    L['lote'] = lote
    alvo = {nv: (qa.run('mLotes.MediaDoLote', lote, AN, nv, teto=60), qa.run('mLotes.DPDoLote', lote, AN, nv, teto=60))
            for nv in (1, 2, 3)}

    def val(nv, k):
        m, s = alvo[nv]
        return round(m + k * s, 3)

    dia = M.HOJE.strftime('%d/%m/%Y')

    def campos(hora, nv, v, data=dia):
        return [('DATA', data), ('HORA', hora), ('EQUIPAMENTO', EQ), ('LOTE', "'" + lote), ('NIVEL', str(nv)),
                ('ANALITO', AN), ('RESULTADO', M.fmt(v)), ('MOTIVO', M.MOTIVO)]

    def digitar(r, cs):
        for col, v in cs:
            R.escrever(qa, r, col, v)

    # nivel: o que tem o resultado real de hoje mais distante de "agora + 3 min" (a DATA/HORA de B sera ~ isso)
    alvo_b = dt.datetime.now() + dt.timedelta(minutes=3)
    dist = {}
    for nv in (1, 2, 3):
        hs = [r['DATA_HORA'] for r in fin0 if t(r.get('ANALITO')) == AN and up(r.get('EQUIPAMENTO')) == EQ
              and num(r.get('NIVEL')) and int(r['NIVEL']) == nv and isinstance(r.get('DATA_HORA'), dt.datetime)
              and r['DATA_HORA'].date() == M.HOJE and not t(r.get('ID_REGISTRO')).startswith('MAN')]
        dist[nv] = min([abs((alvo_b - h).total_seconds()) / 60 for h in hs], default=9999)
    nvA = max(dist, key=dist.get)
    L['distancia_min_por_nivel'] = dist
    L['nivel_A_B'] = nvA
    L['filtro_antes'] = filtro(fin0, lote)
    o0 = R.oraculo(fin0, lote)

    qa.run('mIntegracao.IrDigitarResultados', teto=120)
    qa.ex.xl.EnableEvents = True
    qa.ex.xl.Interactive = True
    vig = R.Caixas(qa.ex.pid)
    vig.start()
    try:
        r0 = R.livres(qa, 2)
        vA, vC = val(nvA, 0.1), val(2, -0.12)
        digitar(r0, campos('03:11', nvA, vA))
        digitar(r0 + 1, campos('03:12', 2, vC))
        L['F1'] = {'A': M.estado(qa, r0), 'C': M.estado(qa, r0 + 1)}
    finally:
        qa.ex.xl.Interactive = False
    msg1, s1 = qa.atualizar()
    fin1 = qa.ler('tblCQ_Final')
    L['atualizar1'] = {'msg': msg1, 's': round(s1, 1), 'tela': {'A': R.ler(qa, r0), 'C': R.ler(qa, r0 + 1)}}
    T0 = M.dtx(M.carimbo(qa, r0)['REGISTRADO_EM'])
    idA, idC = R.idd(qa, r0), R.idd(qa, r0 + 1)
    carC = M.carimbo(qa, r0 + 1)

    qa.run('mIntegracao.IrDigitarResultados', teto=120)
    qa.ex.xl.EnableEvents = True
    qa.ex.xl.Interactive = True
    try:
        R.faixa(qa, r0, r0).ClearContents()
        qa.ex.esperar()
        time.sleep(0.4)
        R.cel(qa, r0 + 1, 'ID_REGISTRO').ClearContents()                     # C04: apaga so o ID de uma linha com dado
        qa.ex.esperar()
        time.sleep(0.6)
        L['F2a'] = {'A_depois_delete': M.estado(qa, r0), 'C_antes': {'ID': idC, **carC},
                    'C_depois_apagar_so_ID': M.estado(qa, r0 + 1), 'caixas': list(vig.vistos)}
    finally:
        qa.ex.xl.Interactive = False

    L['login_U2'] = R.fazer_login(qa, U2)
    falta = (T0 + dt.timedelta(minutes=2, seconds=30) - dt.datetime.now()).total_seconds()
    if falta > 0:
        time.sleep(falta)
    # confere de novo a distancia com os reais do ATUALIZAR #1
    hB = (T0 + dt.timedelta(minutes=1)).replace(second=0)
    if hB <= T0:
        hB += dt.timedelta(minutes=1)
    hs = [r['DATA_HORA'] for r in fin1 if t(r.get('ANALITO')) == AN and up(r.get('EQUIPAMENTO')) == EQ
          and num(r.get('NIVEL')) and int(r['NIVEL']) == nvA and isinstance(r.get('DATA_HORA'), dt.datetime)
          and r['DATA_HORA'].date() == M.HOJE and not t(r.get('ID_REGISTRO')).startswith('MAN')]
    L['B_distancia_min_real'] = min([abs((hB - h).total_seconds()) / 60 for h in hs], default=9999)
    qa.run('mIntegracao.IrDigitarResultados', teto=120)
    botao = {'aba': t(qa.ex.xl.ActiveSheet.Name), 'celula': t(qa.ex.xl.ActiveCell.Address),
             'linha_da_tabela': qa.ex.xl.ActiveCell.Row - R.lo_m(qa).DataBodyRange.Row + 1}
    qa.ex.xl.EnableEvents = True
    qa.ex.xl.Interactive = True
    try:
        rB = botao['linha_da_tabela']
        vB = val(nvA, -0.15)
        agora = dt.datetime.now()
        digitar(rB, campos(hB.strftime('%H:%M'), nvA, vB, hB.strftime('%d/%m/%Y')))
        L['F2b'] = {'botao': botao, 'B': {'linha': rB, 'nivel': nvA, 'valor': vB, 'data_hora': iso(hB),
                                          'momento': agora.isoformat(sep=' ', timespec='seconds'), 'tela': M.estado(qa, rB)},
                    'T0_A': iso(T0)}
    finally:
        qa.ex.xl.Interactive = False
        vig.parar.set()
    aud = V.Audit(qa)
    msg2, s2 = qa.atualizar()
    fin2 = qa.ler('tblCQ_Final')
    L['atualizar2'] = {'msg': msg2, 's': round(s2, 1), 'tela': {'B': R.ler(qa, rB), 'C': R.ler(qa, r0 + 1)},
                       'auditoria': aud.novos(), 'idA': idA, 'idC_antigo': idC}
    L['filtro_depois'] = filtro(fin2, lote)
    idB, idCn = R.idd(qa, rB), R.idd(qa, r0 + 1)
    L['ids'] = {'A': idA, 'B': idB, 'C_antigo': idC, 'C_novo': idCn}
    p2 = R.painel(qa, lote, [idB, idCn], fin2)
    o2 = R.oraculo(fin2, lote)
    L['antes_oraculo'] = o0
    L['depois'] = {'oraculo': o2, 'painel': p2, 'comparacao': R.comparar(o2, p2)}


def rodar():
    os.makedirs(PASTA, exist_ok=True)
    origem = os.path.join(C.REPO, f'QC_{PROD}.xlsm')
    destino = os.path.join(PASTA, f'{PROD}_F1_c1.xlsm')
    if os.path.exists(destino):
        os.remove(destino)
    h0 = C.sha(origem)
    shutil.copy2(origem, destino)
    L = C.Log(os.path.join(PASTA, f'{PROD}_F1_c1.json'))
    L['ambiente'] = {'origem': os.path.basename(origem), 'sha_origem_12': h0, 'copia': os.path.basename(destino),
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
    print('FIM c1', flush=True)


if __name__ == '__main__':
    t0 = time.time()
    rodar()
    print(f'tempo {time.time() - t0:.0f}s', flush=True)
