# -*- coding: utf-8 -*-
"""reteste06_carimbo_residual.py -- Usuario Simulado, ciclo 2026-11-15_reteste06, F1 (Bioquimica): fluxo vizinho da OC-0073.

Pergunta: quando o usuario APAGA uma linha manual (seleciona ID..MOTIVO, A:K, e tecla Delete -- o que a mensagem do
ATUALIZAR e o A01 do ADR-096 mandam fazer com a linha repetida) ou a MOVE (recortar A:K), as colunas travadas USUARIO e
REGISTRADO_EM (L:M) ficam na linha. O proximo manual digitado nessa linha herda o carimbo antigo (Carimbar so preenche
o vazio)? E o motor, que usa REGISTRADO_EM como o instante do lancamento ("DATA/HORA posterior ao lancamento"), recusa
um manual valido por isso?

Fluxo (eventos ligados, como o usuario; usuario TESTE_AUDITORIA_R6C, ANALISTA, so na copia):
  1. digita o manual A (Glicose, DIMENSION 1, lote corrente, N1, 14/11/2026 07:40) -> carimbo T0
  2. apaga A:K da linha A (Delete)  -> le L:M
  3. espera ~2,5 min; mIntegracao.IrDigitarResultados (o botao) -> em que linha o cursor para?
  4. digita o manual B na linha onde o botao parou: DATA = hoje, HORA = T0 + 1 min (anterior ao momento da digitacao)
  5. controle: o manual C igual ao B (N2), numa linha que nunca foi usada
  6. ATUALIZAR DADOS (nucleo do botao) -> STATUS/DETALHE de B e C; tblCQ_Final; QA E04
Uso: python -u reteste06_carimbo_residual.py [nova|base]   (nova = producao atual; base = backup anterior a entrega 43;
copia em %TEMP%/qcw/auditoria/2026-11-15_reteste06/)
Nada e salvo; fechar() encerra so o PID do QA; nenhuma captura de tela.
"""
import datetime as dt
import os
import shutil
import sys
import time
import traceback

AQUI = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, AQUI)
import reteste06_usuario_f1 as R  # noqa: E402

C, V, t = R.C, R.V, R.t
USR = 'TESTE_AUDITORIA_R6C'


def rodar(rotulo='nova'):
    os.makedirs(R.PASTA, exist_ok=True)
    origem = R.BACKUP if rotulo == 'base' else os.path.join(C.REPO, f'QC_{R.PROD}.xlsm')
    suf = '' if rotulo == 'nova' else '_base'
    destino = os.path.join(R.PASTA, f'{R.PROD}_F1_carimbo{suf}.xlsm')
    if os.path.exists(destino):
        os.remove(destino)
    h0 = C.sha(origem)
    shutil.copy2(origem, destino)
    L = C.Log(os.path.join(R.PASTA, f'{R.PROD}_F1_carimbo{suf}.json'))
    L['ambiente'] = {'origem': os.path.basename(origem), 'sha_origem_12': h0, 'copia': os.path.basename(destino)}
    qa = C.qa_final.QA(R.PROD, destino)
    L['excel_pid'] = qa.ex.pid
    try:
        R.criar_usuario(qa, USR, 'ANALISTA')
        L['login'] = R.fazer_login(qa, USR)
        lote = '8977'
        qa.run('mIntegracao.IrDigitarResultados', teto=120)
        qa.ex.xl.EnableEvents = True

        def campos(data, hora, nv, v):
            return [('DATA', data), ('HORA', hora), ('EQUIPAMENTO', R.EQ), ('LOTE', "'" + lote), ('NIVEL', str(nv)),
                    ('ANALITO', R.AN), ('RESULTADO', v), ('MOTIVO', R.MOTIVO)]
        rA = R.livres(qa)
        for c, v in campos('14/11/2026', '07:40', 1, '79,321'):
            R.escrever(qa, rA, c, v)
        a = R.ler(qa, rA)
        L['A_digitado'] = {'linha': rA, 'tela': a}
        R.faixa(qa, rA, rA).ClearContents()          # Delete em A:K (o que o usuario consegue apagar)
        qa.ex.esperar()
        L['A_apagado'] = R.ler(qa, rA)
        t0 = a['REGISTRADO_EM']['v'] if isinstance(a['REGISTRADO_EM'], dict) else a['REGISTRADO_EM']
        t0 = dt.datetime.fromisoformat(str(t0)[:19])
        time.sleep(150)
        qa.run('mIntegracao.IrDigitarResultados', teto=120)
        L['botao_parou_em'] = {'aba': t(qa.ex.xl.ActiveSheet.Name), 'celula': t(qa.ex.xl.ActiveCell.Address),
                               'linha_da_tabela': qa.ex.xl.ActiveCell.Row - R.lo_m(qa).DataBodyRange.Row + 1}
        qa.ex.xl.EnableEvents = True
        hB = (t0 + dt.timedelta(minutes=1)).replace(second=0)
        if hB <= t0:
            hB += dt.timedelta(minutes=1)
        agora = dt.datetime.now()
        rB = L['botao_parou_em']['linha_da_tabela']
        for c, v in campos(hB.strftime('%d/%m/%Y'), hB.strftime('%H:%M'), 1, '79,654'):
            R.escrever(qa, rB, c, v)
        L['B_digitado'] = {'linha': rB, 'data_hora_digitada': hB.isoformat(sep=' '), 'momento_da_digitacao': agora.isoformat(sep=' ', timespec='seconds'),
                           'T0_carimbo_de_A': t0.isoformat(sep=' '), 'tela': R.ler(qa, rB)}
        rC = R.livres(qa, 1, rB)
        for c, v in campos(hB.strftime('%d/%m/%Y'), hB.strftime('%H:%M'), 2, '274,321'):
            R.escrever(qa, rC, c, v)
        L['C_controle'] = {'linha': rC, 'tela': R.ler(qa, rC)}
        msg, s = qa.atualizar()
        fin = qa.ler('tblCQ_Final')
        ids = [t(R.ler(qa, r)['ID_REGISTRO']) for r in (rB, rC)]
        L['atualizar'] = {'msg': msg, 's': round(s, 1),
                          'tela': {k: R.ler(qa, r) for k, r in (('B', rB), ('C', rC))},
                          'tblCQ_Final': [{c: (x[c].isoformat(sep=' ') if isinstance(x.get(c), dt.datetime) else x.get(c))
                                           for c in ('ID_REGISTRO', 'STATUS_ANALITICO', 'PARTICIPA_ESTATISTICA', 'NIVEL',
                                                     'RESULTADO', 'DATA_HORA', 'MOTIVO_EXCLUSAO_AUTOMATICA')}
                                          for x in fin if t(x['ID_REGISTRO']).split('~')[0] in ids],
                          'qa': [' | '.join(t(v) for v in x.values() if v not in (None, ''))[:400]
                                 for x in qa.ler('tblQA_Integracao') if any(i in ' '.join(t(v) for v in x.values()) for i in ids)]}
    except Exception as ex:  # noqa: BLE001
        L['ERRO_FATAL'] = {'erro': repr(ex), 'trace': traceback.format_exc()[-3000:]}
    finally:
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
    print('FIM', flush=True)


if __name__ == '__main__':
    rodar(sys.argv[1] if len(sys.argv) > 1 else 'nova')
