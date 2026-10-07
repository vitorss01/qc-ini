# -*- coding: utf-8 -*-
"""QA de CASOS EXTREMOS (QUALITY_GATE 4.4) na arquitetura do ADR-057 -- roda numa COPIA ja instalada.

Uso: python qa_casos_extremos.py <Bioquimica|Hematologia> <copia_instalada.xlsm> [saida.json]

Na arquitetura antiga o 4.4 falava de banco vazio, RUN duplicado e status invalido no
DB_Resultados. Hoje a unica porta de dado digitado e a aba Digitar Resultados e a
unica porta de exclusao e a aba Inativar; e por elas que o dado ruim chega. Cada caso
entra PELO CAMINHO REAL (celula a celula, evento ligado, ATUALIZAR DADOS) e o teste
exige a classificacao certa na Principal - Resultados, o achado certo no QA e que
nada ruim participe da estatistica. Nada e salvo no arquivo.

Funciona fora da rede do laboratorio: o arquivo deve estar com MODO_FONTE = HISTORICO
(ADR-058) ou com o DB_SEAC acessivel.
"""
import datetime as dt
import json
import os
import sys
import traceback

AQUI = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, AQUI)
from qa_final import QA, reg, RES, NAO, sha256_arquivo, sem_acento  # noqa: E402

SIM = 'SIM'


def executar(produto, caminho, saida):
    hash0 = sha256_arquivo(caminho)
    q = QA(produto, caminho)
    try:
        fin0 = q.final_por_id()
        # um resultado ATIVO real do interfaceamento, de analito cadastrado, serve de molde
        molde = next(r for r in fin0.values() if r['ORIGEM_RESULTADO'] == 'INTERFACEAMENTO'
                     and r['STATUS_ANALITICO'] == 'ATIVO' and r['ANALITO_CADASTRADO'] == SIM and r['LOTE'])
        base = {'LOTE': molde['LOTE'], 'NIVEL': 1, 'ANALITO': molde['ANALITO'], 'EQUIPAMENTO': molde['EQUIPAMENTO']}
        # Madrugada (00:07..06:07, ja passada) do dia mais recente em que a chave do molde (equipamento,
        # lote, analito) NAO tem resultado do interfaceamento entre 23:00 da vespera e 07:00: nenhum caso
        # vira CONFLITO_MANUAL por acaso, nem com a janela de conflito atravessando a meia-noite (D09c).
        # Antes o teste so aceitava HOJE e abortava (SystemExit, sem JSON) no laboratorio, onde hoje ha dado.
        chave_m = (molde['EQUIPAMENTO'], str(molde['LOTE']), molde['ANALITO'])
        inter_m = [r['DATA_HORA'] for r in fin0.values() if r['ORIGEM_RESULTADO'] == 'INTERFACEAMENTO'
                   and r['DATA_HORA'] and (r['EQUIPAMENTO'], str(r['LOTE']), r['ANALITO']) == chave_m]
        agora = dt.datetime.now()
        d0 = dt.datetime(agora.year, agora.month, agora.day) - dt.timedelta(days=1 if agora.hour < 7 else 0)
        dia = next((c for c in (d0 - dt.timedelta(days=k) for k in range(120))
                    if not any(c - dt.timedelta(hours=1) <= h <= c + dt.timedelta(hours=7) for h in inter_m)), None)
        if dia is None:
            raise RuntimeError('pré-condição ausente: nenhum dia nos últimos 120 sem interfaceamento da chave do molde '
                               f'{chave_m} na madrugada')
        print('DIA_DOS_MANUAIS', dia.date(), 'chave', chave_m, flush=True)

        def manual(hora, **kw):
            v = dict(base, DATA=dia, HORA=dt.time(hora, 7).strftime('%H:%M'), RESULTADO=molde['RESULTADO'])
            v.update(kw)
            return q.escrever_linha('tblResultados_Manuais', v)

        futuro = dt.datetime.now() + dt.timedelta(days=40)
        manual(1, DATA=dt.datetime(futuro.year, futuro.month, futuro.day))            # a) data futura
        manual(2, RESULTADO='abc')                                                     # b) resultado nao numerico
        manual(3, NIVEL=q.nlv + 1)                                                     # c) nivel inexistente
        manual(4, ID_REGISTRO='XYZ_12')                                                # d) ID fora do padrao
        manual(5, ANALITO='ANALITO_QUE_NAO_EXISTE')                                    # e) analito sem cadastro
        res_virg = str(molde['RESULTADO']).replace('.', ',')
        manual(6, RESULTADO=res_virg)                                                  # f) decimal com virgula (pt-BR)
        # i) X10/D07: '1.5' gravado como TEXTO (apostrofo = texto na celula, sem mexer em NumberFormat).
        # 00:07 do dia escolhido: sempre no passado e longe dos outros casos (01:07..06:07)
        lin10 = manual(0, RESULTADO="'1.5")
        cel10 = q.lo('tblResultados_Manuais').ListColumns('RESULTADO').DataBodyRange.Cells(lin10, 1).Value
        q.escrever_linha('tblInativacao_NaoConformes', {'ID_REGISTRO': 'MAN_9999'})    # g) inativar ID inexistente
        rep_r = next(r for r in fin0.values() if r['ORIGEM_RESULTADO'] == 'INTERFACEAMENTO'
                     and r['STATUS_ANALITICO'] == 'ATIVO' and r['ID_REGISTRO'] != molde['ID_REGISTRO'])
        rep = rep_r['ID_REGISTRO']
        # h) mesmo ID inativado 2x (com o analito do resultado: ADR-070)
        q.escrever_linha('tblInativacao_NaoConformes', {'ID_REGISTRO': rep, 'ANALITO': rep_r['ANALITO']})
        q.escrever_linha('tblInativacao_NaoConformes', {'ID_REGISTRO': rep, 'ANALITO': rep_r['ANALITO']})
        q.escrever_linha('tblComentariosTecnicos', {'ID_REGISTRO': rep, 'COMENTARIO_TECNICO': 'QA casos extremos'})

        q.atualizar()
        fin = q.ler('tblCQ_Final')
        man = {r['_h']: r for r in (dict(x, _h=(x['DATA_HORA'].hour if x['DATA_HORA'] else None)) for x in fin
                                    if x['ORIGEM_RESULTADO'] == 'MANUAL' and x['ID_REGISTRO'] not in
                                    {i for i, v in fin0.items() if v['ORIGEM_RESULTADO'] == 'MANUAL'})}
        inval = [x for x in fin if str(x['ID_REGISTRO']).startswith('MAN_INVALIDO')]
        qa = q.ler('tblQA_Integracao')
        cod = lambda c: [x for x in qa if x['CODIGO'] == c]

        def fora(r):
            return r is not None and r['PARTICIPA_ESTATISTICA'] == NAO and r['TIPO_PLOTAGEM_LJ'] == 'NAO_PLOTAR'

        novos = [x for x in fin if x['ORIGEM_RESULTADO'] == 'MANUAL' and x['ID_REGISTRO'] not in fin0]
        fut = next((x for x in novos if x['DATA'] and x['DATA'].year == futuro.year and x['DATA'].month == futuro.month
                    and x['DATA'].day == futuro.day), None)
        reg('X01 Data futura: MANUAL_INCOMPLETO, fora da estatística e do LJ, E04 no QA',
            fut is not None and fut['STATUS_ANALITICO'] == 'MANUAL_INCOMPLETO' and fora(fut)
            and 'posterior ao lancamento' in str(fut['MOTIVO_EXCLUSAO_AUTOMATICA'])
            and any(x['ID_REGISTRO'] == fut['ID_REGISTRO'] for x in cod('E04')),
            {k: (fut or {}).get(k) for k in ('ID_REGISTRO', 'DATA_HORA', 'STATUS_ANALITICO', 'MOTIVO_EXCLUSAO_AUTOMATICA')})
        r = man.get(2)
        reg('X02 Resultado não numérico ("abc"): MANUAL_INCOMPLETO, fora',
            r is not None and r['STATUS_ANALITICO'] == 'MANUAL_INCOMPLETO' and fora(r) and 'RESULTADO' in str(r['MOTIVO_EXCLUSAO_AUTOMATICA']),
            {k: (r or {}).get(k) for k in ('ID_REGISTRO', 'STATUS_ANALITICO', 'MOTIVO_EXCLUSAO_AUTOMATICA')})
        r = man.get(3)
        reg(f'X03 Nível inexistente ({q.nlv + 1}): MANUAL_INCOMPLETO, fora',
            r is not None and r['STATUS_ANALITICO'] == 'MANUAL_INCOMPLETO' and fora(r) and 'NIVEL' in str(r['MOTIVO_EXCLUSAO_AUTOMATICA']),
            {k: (r or {}).get(k) for k in ('ID_REGISTRO', 'NIVEL', 'STATUS_ANALITICO', 'MOTIVO_EXCLUSAO_AUTOMATICA')})
        reg('X04 ID manual fora do padrão (XYZ_12): vira MAN_INVALIDO_L<n>, MANUAL_INCOMPLETO, fora',
            len(inval) == 1 and inval[0]['STATUS_ANALITICO'] == 'MANUAL_INCOMPLETO' and fora(inval[0]),
            [{k: x.get(k) for k in ('ID_REGISTRO', 'STATUS_ANALITICO', 'MOTIVO_EXCLUSAO_AUTOMATICA')} for x in inval])
        r = man.get(5)
        # D17: "fora do Painel" agora e assertado. O Painel escolhe o analito pelo INDICE na lista de
        # Analitos!A4:A43 (Painel!B3) e o motor so tem alvo para quem esta la: fora dessa lista o analito
        # nao pode ser exibido. Confere tambem que o analito em tela (Eng_Saida!C1) nao e ele.
        cadastro = [str(v[0] or '').strip().upper() for v in q.ws('Analitos').Range('A4:A43').Value]
        an_tela = str(q.ws('Eng_Saida').Range('C1').Value or '').strip().upper()
        fora_painel = bool([c for c in cadastro if c]) and 'ANALITO_QUE_NAO_EXISTE' not in cadastro \
            and an_tela != 'ANALITO_QUE_NAO_EXISTE'
        reg('X05 Analito sem cadastro: guardado, ANALITO_CADASTRADO = NÃO, I01 no QA, fora do Painel',
            r is not None and r['ANALITO_CADASTRADO'] == NAO and any(x['ANALITO'] == 'ANALITO_QUE_NAO_EXISTE' for x in cod('I01'))
            and fora_painel,
            dict({k: (r or {}).get(k) for k in ('ID_REGISTRO', 'ANALITO', 'ANALITO_CADASTRADO', 'STATUS_ANALITICO',
                                                 'PARTICIPA_ESTATISTICA')},
                 fora_do_painel=fora_painel, analito_em_tela=an_tela, analitos_no_cadastro=len([c for c in cadastro if c])))
        r = man.get(6)
        reg(f'X06 Decimal com vírgula ("{res_virg}"): lido como {molde["RESULTADO"]}, nunca como milhar (armadilha do item 2.2)',
            r is not None and r['STATUS_ANALITICO'] == 'ATIVO' and abs(float(r['RESULTADO']) - float(molde['RESULTADO'])) < 1e-9,
            {k: (r or {}).get(k) for k in ('ID_REGISTRO', 'RESULTADO', 'STATUS_ANALITICO')})
        reg('X07 Inativar ID inexistente (MAN_9999): E03 no QA, nenhum resultado afetado',
            any(x['ID_REGISTRO'] == 'MAN_9999' for x in cod('E03')) and 'MAN_9999' not in {x['ID_REGISTRO'] for x in fin},
            [x['DETALHE'] for x in cod('E03')][:3])
        fr = next((x for x in fin if x['ID_REGISTRO'] == rep), None)
        reg('X08 Mesmo ID inativado duas vezes: E02 no QA, uma única inativação, resultado não duplica',
            any(x['ID_REGISTRO'] == rep for x in cod('E02')) and sum(1 for x in fin if x['ID_REGISTRO'] == rep) == 1
            and fr['STATUS_ANALITICO'] == 'INATIVADO',
            {'id': rep, 'E02': [x['DETALHE'] for x in cod('E02')][:2], 'status': (fr or {}).get('STATUS_ANALITICO')})
        ruins = [x for x in novos if x['STATUS_ANALITICO'] == 'MANUAL_INCOMPLETO' and x['PARTICIPA_ESTATISTICA'] != NAO]
        # 7 manuais novos (a..f + X10); nada da base some; o que mais entrar so pode ser interfaceamento
        # (em MODO SEAC a mesma atualizacao pode receber resultados novos do DB_SEAC)
        ids_fin = [x['ID_REGISTRO'] for x in fin]
        sumidos = sorted(str(i) for i in set(fin0) - set(ids_fin))
        outros = [x['ID_REGISTRO'] for x in fin if x['ID_REGISTRO'] not in fin0 and x['ORIGEM_RESULTADO'] != 'MANUAL'
                  and x['ORIGEM_RESULTADO'] != 'INTERFACEAMENTO']
        reg('X09 Nenhum resultado inválido participa da estatística',
            not ruins and len(novos) == 7 and not sumidos and not outros and len(ids_fin) == len(set(ids_fin)),
            {'linhas': [len(fin0), len(fin)], 'manuais_novos': len(novos), 'sumidos': sumidos[:10],
             'novos_sem_origem_valida': outros[:10], 'participando_indevido': ruins[:3]})
        # X10 (D07): '1.5' em TEXTO e ambiguo em pt-BR (ponto = milhar): nunca pode virar 15 e participar
        r = man.get(0)
        pre10 = isinstance(cel10, str) and cel10.strip() == '1.5'
        e04_10 = [x for x in cod('E04') if r is not None and x['ID_REGISTRO'] == r['ID_REGISTRO']]
        amb10 = any('resultado ambiguo' in sem_acento(x['DETALHE'] or '').lower() for x in e04_10)
        res10 = (r or {}).get('RESULTADO')
        virou15 = isinstance(res10, (int, float)) and abs(float(res10) - 15.0) < 1e-9
        ev10 = {'celula_RESULTADO': cel10, 'tipo_celula': type(cel10).__name__,
                **{k: (r or {}).get(k) for k in ('ID_REGISTRO', 'RESULTADO', 'STATUS_ANALITICO', 'PARTICIPA_ESTATISTICA',
                                                  'TIPO_PLOTAGEM_LJ', 'MOTIVO_EXCLUSAO_AUTOMATICA')},
                'E04': [x['DETALHE'] for x in e04_10]}
        if not pre10:
            ev10['falha'] = 'pré-condição ausente: a célula RESULTADO não ficou com o TEXTO 1.5'
        if r is None:
            ev10['falha'] = 'pré-condição ausente: o manual das 00:07 não chegou à Principal - Resultados'
        reg("X10 RESULTADO manual '1.5' gravado como TEXTO: MANUAL_INCOMPLETO, fora da estatística e do LJ, "
            "E04 com 'RESULTADO ambíguo', nunca vira 15 (hoje falha: prova D07)",
            pre10 and r is not None and r['STATUS_ANALITICO'] == 'MANUAL_INCOMPLETO' and fora(r) and amb10 and not virou15,
            ev10)
    except Exception as ex_geral:
        # nenhum caso some calado: a interrupcao vira FAIL, e o hash/JSON saem mesmo assim
        reg('X00 Execução interrompida: os casos seguintes a este ponto não rodaram', False,
            {'erro': f'{type(ex_geral).__name__}: {ex_geral}', 'traceback': traceback.format_exc()[-1500:]})
    finally:
        q.fechar(salvar=False)
    hash1 = sha256_arquivo(caminho)
    reg('X11 Segurança do teste: arquivo de entrada intacto (SHA-256 igual antes e depois; nada salvo nele)',
        hash0 == hash1, {'arquivo': caminho, 'sha256_antes': hash0, 'sha256_depois': hash1})
    if saida:
        with open(saida, 'w', encoding='utf-8') as f:
            json.dump(RES, f, ensure_ascii=False, indent=1, default=str)
    falhas = [r for r in RES if r['resultado'] == 'FAIL']
    print(f'\n=== {produto} casos extremos: {len(RES) - len(falhas)} PASS / {len(falhas)} FAIL ===', flush=True)
    return falhas


if __name__ == '__main__':
    falhas = executar(sys.argv[1], os.path.abspath(sys.argv[2]), sys.argv[3] if len(sys.argv) > 3 else None)
    sys.exit(1 if falhas else 0)
