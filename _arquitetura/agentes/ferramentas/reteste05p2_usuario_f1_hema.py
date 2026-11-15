# -*- coding: utf-8 -*-
"""reteste05p2_usuario_f1_hema.py -- Usuario Simulado, ciclo 2026-11-15_reteste05_parte2, frente F1, RETESTE (Hematologia).

Irmao do reteste05p2_usuario_f1.py (Bioquimica), de onde reaproveita o oraculo de acentos, a leitura de datas, a
avaliacao de listas e o filtro da aba Registros. Correcao declarada: ADR-095, entrega 42, Hema 7bab8fc0.
  OC-0021  DETALHE sem acento. Hema: IMPLAUSIVEL, substituicao (ADR-082: valido, ja substituido, inexistente, ID manual,
           outra chave, sem MOTIVO, retransmissao se houver DUPLICIDADE_ORIGEM), regras de lote do ADR-091 (a Hema nao tem
           regra; o administrador da COPIA cadastra duas regras TESTE-AUDITORIA na tblLotesMaterial, como o GUIA manda:
           FORA_DO_CQ e SO_ANALITOS), linha inteira colada (ID repetido?). A troca de analito de urina (ADR-093) e so Bio.
           Manuais digitados celula a celula (eventos ligados), ATUALIZAR pelo nucleo do botao
           (mIntegracao.AtualizarDadosAutomatico, le o DB_SEAC so para leitura), DETALHE/RUN/STATUS (Text) e tblCQ_Final.
  OC-0068  Listas: Digitar Resultados EQUIPAMENTO/MATRIZ/MOTIVO, Registros H4/F7, Painel L4 (se houver). Oraculo dos itens
           validos: equipamentos e matrizes distintos da tblCQ_Final; os 4 motivos do GUIA/ADR-082. Manual com EQUIPAMENTO,
           MATRIZ e MOTIVO escolhidos pelo item do menu -> entra na serie certa do Painel (n +1, valor no Eng_Saida); a
           substituicao valida com o motivo do menu grava o motivo certo na trilha (S1 e QA I04).
  OC-0072  Ano literal: Text/NumberFormat apos o login (Eventos_Westgard A4:A5003, Registros E4/F4 digitados,
           Configuracao D/AD 26-125, Eng_Saida AC3:AC182, Westgard_Fotos F, Sync_Log, Estatistica, LotesStore U1) antes e
           depois do ATUALIZAR; XML (xml_formatos_listas) da copia sem Excel e do SaveCopyAs depois do uso.
A MESMA sequencia roda em duas copias, uma de cada vez (um Excel por vez):
  base41 = copia do backup anterior a entrega 42 (_backup_pre_integracao_2026-11-15_0458, entrega 41) -- o "antes";
  nova   = copia da producao atual (raiz).
  xml    = sem Excel: copia das duas e compara o XML (formatos com ano literal, x12ac:list, valores <v> dos alvos).
Regras: so COPIAS em %TEMP%/qcw/auditoria/2026-11-15_reteste05_parte2/; a copia em uso nunca e salva (o SaveCopyAs grava
um terceiro arquivo, so para ler o XML); fechar() encerra so o PID do QA; nenhuma captura de tela (caixas lidas pelo
texto das janelas do PID desta sessao e fechadas). Dados de teste: manuais com MOTIVO "Outro: TESTE-AUDITORIA" (ou o
item do menu, nos casos da OC-0068), valores 123,456 / 111,111 / 12,345 / -1; regras de lote com OBS TESTE-AUDITORIA;
usuario TESTE_AUDITORIA_R5P2H (ANALISTA) criado so na copia.
Uso: python -u reteste05p2_usuario_f1_hema.py <base41|nova|xml>  -> Hematologia_F1_<rotulo>.json na pasta do ciclo
"""
import collections
import datetime as dt
import json
import os
import re
import shutil
import sys
import time
import traceback
import zipfile

AQUI = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, AQUI)
_argv = sys.argv[:]
sys.argv = [sys.argv[0], 'Hematologia']
import valida_adr090_091 as V  # noqa: E402
import reteste05p2_usuario_f1 as B  # noqa: E402  (reaproveita oraculo/leituras; V ja carregado como Hematologia)
sys.argv = _argv
import repro_comum as C  # noqa: E402
import xml_formatos_listas as X  # noqa: E402
from qa_inativar_ux import criar_usuario, fazer_login  # noqa: E402

assert not V.BIO
PROD = 'Hematologia'
NLV = 3
AN = 'WBC'
CICLO = '2026-11-15_reteste05_parte2'
PASTA = os.path.join(os.environ['TEMP'], 'qcw', 'auditoria', CICLO)
C.PASTA = PASTA
BACKUP = os.path.join(C.REPO, '_backup_pre_integracao_2026-11-15_0458', f'QC_{PROD}.xlsm')
USR = 'TESTE_AUDITORIA_R5P2H'
MOTIVO = 'Outro: TESTE-AUDITORIA'
MOTIVOS = ['Falha do interfaceamento', 'Resultado não transmitido', 'Equipamento em contingência', 'Outro']
MESES = B.MESES
t, up, num = V.t, V.up, V.num
sem_acento_achados, acentuar, tira_acento = B.sem_acento_achados, B.acentuar, B.tira_acento


def secao(L, nome, fn, *a):
    print('>>', nome, time.strftime('%H:%M:%S'), flush=True)
    return V.secao(L, nome, fn, *a)


# ------------------------------------------------------------------------------------------- OC-0072
def alvos_0072(qa):
    out = {}
    nomes = [str(w.Name) for w in qa.wb.Worksheets]
    out['abas'] = [n for n in nomes if n in ('Eventos_Westgard', 'Sync_Log', 'Westgard_Fotos', 'Configuração', 'Eng_Saida',
                                             'Registros', 'Estatística', 'LotesStore')]

    def tenta(k, aba, ender):
        try:
            out[k] = B.ler_datas(qa, aba, ender)
        except Exception as ex:  # noqa: BLE001
            out[k] = {'erro': repr(ex)[:160]}
    tenta('Eventos_Westgard_A', 'Eventos_Westgard', 'A4:A5003')
    try:
        ew = qa.ws('Eventos_Westgard')
        out['Eventos_Westgard_A_fmt_A4_A5003'] = t(ew.Range('A4:A5003').NumberFormat)
        out['Eventos_Westgard_cab'] = [t(ew.Cells(3, j).Value) for j in range(1, 8)]
    except Exception as ex:  # noqa: BLE001
        out['Eventos_Westgard_A_fmt_A4_A5003'] = repr(ex)[:120]
    tenta('Configuracao_D', 'Configuração', 'D26:D125')
    tenta('Configuracao_AD', 'Configuração', 'AD26:AD125')
    try:
        out['Configuracao_fmt_D_AD'] = [t(qa.ws('Configuração').Range('D26:D125').NumberFormat),
                                        t(qa.ws('Configuração').Range('AD26:AD125').NumberFormat),
                                        bool(qa.ws('Configuração').Range('AD26').EntireColumn.Hidden)]
    except Exception as ex:  # noqa: BLE001
        out['Configuracao_fmt_D_AD'] = repr(ex)[:120]
    tenta('Eng_Saida_AC', 'Eng_Saida', 'AC3:AC182')
    try:
        ws = qa.ws('Westgard_Fotos')
        n = ws.Cells(ws.Rows.Count, 6).End(-4162).Row
        tenta('Westgard_Fotos_F_inicio', 'Westgard_Fotos', 'F4:F40')
        tenta('Westgard_Fotos_F_fim', 'Westgard_Fotos', f'F{max(4, n - 40)}:F{n}')
        out['Westgard_Fotos_F_fmt_coluna'] = t(ws.Range('F:F').NumberFormat)
    except Exception as ex:  # noqa: BLE001
        out['Westgard_Fotos'] = repr(ex)[:120]
    if 'Sync_Log' in nomes:
        sl = qa.ws('Sync_Log')
        m = sl.Cells(sl.Rows.Count, 1).End(-4162).Row
        for col in ('A', 'B', 'L', 'M'):
            tenta(f'Sync_Log_{col}', 'Sync_Log', f'{col}4:{col}{max(4, m)}')
        out['Sync_Log_ultima_linha'] = m
    if 'Estatística' in nomes:
        e = qa.ws('Estatística')
        out['Estatistica_E8F9C10'] = {a: {'txt': t(e.Range(a).Text), 'v2': e.Range(a).Value2,
                                          'fmt': t(e.Range(a).NumberFormat)} for a in ('E8', 'F8', 'E9', 'F9', 'C10')}
    if 'LotesStore' in nomes:
        tenta('LotesStore_U1', 'LotesStore', 'U1')
    out['n_ruins_total'] = sum(v.get('n_ruins', 0) for v in out.values() if isinstance(v, dict))
    return out


# ------------------------------------------------------------------------------------------- OC-0068
def s_listas(qa, fin):
    sep = t(qa.ex.xl.International[4])
    out = {'sep_lista_excel': sep}
    lo, _ = V.manual_cols(qa)
    _, r = C.linha_livre_manual(qa)
    # oraculo dos itens validos: o que existe nos dados (tblCQ_Final), nao a propria validacao
    eqs = sorted({t(x['EQUIPAMENTO']) for x in fin if t(x['EQUIPAMENTO'])})
    mzs = sorted({t(x['MATRIZ']) for x in fin if t(x.get('MATRIZ'))})
    out['oraculo'] = {'equipamentos_tblCQ_Final': eqs, 'matrizes_tblCQ_Final': mzs, 'motivos_GUIA': MOTIVOS}

    def av(cel, validos, invalidos):
        try:
            return B.avaliar_lista(qa, cel, validos, invalidos, sep)
        except Exception as ex:  # noqa: BLE001
            return {'erro': repr(ex)[:240], 'ok': False, 'msg': ''}
    inv_eq = ['XN9999', 'XN'] + ([sep.join(eqs), ','.join(eqs)] if len(eqs) > 1 else [eqs[0] + ',' + eqs[0]])
    out['DR_EQUIPAMENTO'] = av(lo.ListColumns('EQUIPAMENTO').DataBodyRange.Cells(r, 1), eqs, inv_eq)
    out['DR_MATRIZ'] = av(lo.ListColumns('MATRIZ').DataBodyRange.Cells(r, 1), mzs, ['SORO', 'PLASMA', 'SANGUE'])
    out['DR_MOTIVO'] = av(lo.ListColumns('MOTIVO').DataBodyRange.Cells(r, 1), MOTIVOS,
                          ['Outro: TESTE-AUDITORIA', ','.join(MOTIVOS), 'Falha', 'outro'])
    viz = {}
    for col in ('LOTE', 'ANALITO', 'NIVEL', 'DATA', 'HORA', 'RESULTADO', 'SUBSTITUI_ID'):
        c = lo.ListColumns(col).DataBodyRange.Cells(r, 1)
        try:
            viz[col] = [c.Validation.Type, t(c.Validation.Formula1)[:60], t(c.Validation.ErrorMessage)[:120]]
        except Exception as ex:  # noqa: BLE001
            viz[col] = 'sem validacao: ' + repr(ex)[:60]
    out['DR_vizinhos'] = viz
    rg = qa.ws('Registros')
    out['Registros_H4'] = av(rg.Range('H4'), ['Sim', 'Não'], ['Sim,Não', 'Talvez', 'S'])
    out['Registros_F7_foi_calibrado'] = av(rg.Range('F7'), ['Sim', 'Não'], ['Sim,Não'])
    p = qa.ws('Painel')
    try:
        vv = p.Range('L4').Validation
        out['Painel_L4'] = {'tipo': vv.Type, 'formula1': t(vv.Formula1)[:200], 'L4': t(p.Range('L4').Value),
                            'K4_rotulo': t(p.Range('L3').Value)}
    except Exception as ex:  # noqa: BLE001
        out['Painel_L4'] = {'sem_validacao': repr(ex)[:120], 'L4': t(p.Range('L4').Value), 'L3': t(p.Range('L3').Value)}
    for a in ('K4', 'M4', 'C3', 'H3', 'K3'):
        try:
            f1 = t(p.Range(a).Validation.Formula1)
            out[f'Painel_{a}_vizinho'] = f1[:200] if f1.startswith('=') else f1.split(sep)[:8]
        except Exception as ex:  # noqa: BLE001
            out[f'Painel_{a}_vizinho'] = 'sem validacao'
    chaves = ('DR_EQUIPAMENTO', 'DR_MATRIZ', 'DR_MOTIVO', 'Registros_H4')
    out['ok'] = all(out[k]['ok'] for k in chaves)
    out['msg_seta'] = {k: ('na seta da célula' in out[k].get('msg', '')) for k in chaves}
    return out


# ------------------------------------------------------------------------------------------- regras de lote (copia)
def cadastrar_regras(qa, fin, senha, lote_evitar):
    """Administrador da COPIA cadastra 2 regras TESTE-AUDITORIA (FORA_DO_CQ, SO_ANALITOS) na tblLotesMaterial."""
    lo = qa.lo('tblLotesMaterial')
    cab = [str(c.Name) for c in lo.ListColumns]
    out = {'colunas': cab, 'n_antes': lo.ListRows.Count,
           'linhas_antes': [[t(x) for x in l] for l in (lo.DataBodyRange.Value if lo.ListRows.Count else ())][:5]}
    cnt = collections.Counter(t(r['LOTE']) for r in fin if up(r['ANALITO']) == AN and r['PARTICIPA_ESTATISTICA'] == 'SIM')
    cand = sorted([lt for lt, n in cnt.items() if lt and lt not in lote_evitar and n >= 6], key=lambda lt: cnt[lt])
    fora, so = cand[0], cand[1]
    out['escolha'] = {'FORA_DO_CQ': fora, 'n_WBC_fora': cnt[fora], 'SO_ANALITOS': so, 'n_WBC_so': cnt[so],
                      'analitos_SO': 'RBC'}
    ws = lo.Parent
    out['aba'] = str(ws.Name)
    try:
        ws.Unprotect(senha)
    except Exception:  # noqa: BLE001
        pass
    for regra in (('FORA_DO_CQ', fora, ''), ('SO_ANALITOS', so, 'RBC')):
        # linha livre (a tabela pode ter linha vazia no fim)
        vals = lo.DataBodyRange.Value if lo.ListRows.Count else ()
        livre = next((i + 1 for i, l in enumerate(vals) if not any(x not in (None, '') for x in l)), None)
        if livre is None:
            lo.ListRows.Add()
            livre = lo.ListRows.Count
        row = lo.ListRows(livre).Range
        for j, c in enumerate(cab, 1):
            cu = up(c)
            v = None
            if cu == 'LOTE':
                v = "'" + regra[1]
            elif cu == 'REGRA':
                v = regra[0]
            elif cu == 'ANALITOS':
                v = regra[2] or None
            elif any(k in cu for k in ('OBS', 'COMENT', 'MOTIVO', 'NOTA', 'DESCR')):
                v = 'TESTE-AUDITORIA'
            if v is not None:
                row.Cells(1, j).FormulaLocal = v
                qa.ex.esperar()
    out['linhas_depois'] = [[t(x) for x in l] for l in lo.DataBodyRange.Value]
    try:
        out['resumo_vba'] = t(qa.run('mEtlCQ.LotesMaterialResumo'))[:300]
    except Exception as ex:  # noqa: BLE001
        out['resumo_vba'] = repr(ex)[:120]
    return out, fora, so


# ------------------------------------------------------------------------------------------- manuais
def painel_serie(qa, lote, ano, mes):
    p = qa.ws('Painel')
    for c, v in (('C3', AN), ('H3', "'" + lote), ('K3', ano), ('K4', MESES[mes - 1])):
        C.por(qa, p, c, v, local=False)
    time.sleep(1.2)
    qa.ex.esperar()
    es = qa.eng_saida()
    vals1 = [l[0] for l in es['val']] if es['n'] else []
    return {'H3': t(p.Range('H3').Value), 'K3': t(p.Range('K3').Value), 'K4': t(p.Range('K4').Value),
            'engPainel': qa.painel_stats(), 'eng_E1': t(es.get('lote')), 'n_pontos': es['n'],
            'pontos_N1': vals1,
            'tem_123456_N1': any(isinstance(x, float) and abs(x - 123.456) < 1e-9 for x in vals1)}


def s_manuais(qa, fin, sep, senha):
    out = {}
    p = qa.ws('Painel')
    lo, cab = V.manual_cols(qa)
    out['colunas'] = cab

    def reais():
        return sorted([r for r in fin if r['STATUS_ANALITICO'] == 'ATIVO' and r['ORIGEM_RESULTADO'] == 'INTERFACEAMENTO'
                       and t(r['ANALITO']) == AN and int(r['NIVEL'] or 0) == 1
                       and isinstance(r['DATA_HORA'], dt.datetime)], key=lambda r: r['DATA_HORA'])
    r1 = reais()
    R, S1, S2, S3 = r1[-40], r1[-41], r1[-42], r1[-43]
    R2 = r1[-30]
    dup = [r for r in fin if r['STATUS_ANALITICO'] == 'DUPLICIDADE_ORIGEM' and r['ORIGEM_RESULTADO'] == 'INTERFACEAMENTO']
    ref = lambda r: {k: t(r[k]) for k in ('ID_REGISTRO', 'EQUIPAMENTO', 'MATRIZ', 'LOTE', 'NIVEL', 'ANALITO',  # noqa: E731
                                          'RESULTADO', 'DATA_HORA', 'RUN', 'STATUS_ANALITICO')}
    out['referencias'] = {'R': ref(R), 'S1': ref(S1), 'S2': ref(S2), 'S3': ref(S3), 'R2': ref(R2),
                          'DUP': ref(dup[0]) if dup else None, 'n_DUPLICIDADE_ORIGEM': len(dup)}
    numid = lambda r: re.sub(r'\D', '', t(r['ID_REGISTRO']))  # noqa: E731
    dia, lote = R['DATA_HORA'].date(), t(R['LOTE'])
    dia2, lote2 = R2['DATA_HORA'].date(), t(R2['LOTE'])
    sdia, sdia2 = dia.strftime('%d/%m/%Y'), dia2.strftime('%d/%m/%Y')
    eqR = t(R['EQUIPAMENTO'])
    us, us2 = [], []
    # regras de lote TESTE-AUDITORIA (a Hema nao tem nenhuma)
    regras, lote_fora, lote_so = cadastrar_regras(qa, fin, senha, {lote, lote2, t(S1['LOTE']), t(S2['LOTE']),
                                                                    t(S3['LOTE'])})
    out['regras_lote'] = regras
    # Painel ANTES (OC-0068: a serie do manual)
    out['serie_antes'] = painel_serie(qa, lote2, dia2.year, dia2.month)
    C.por(qa, p, 'C3', AN, local=False)
    try:
        qa.run('mIntegracao.IrDigitarResultados', teto=120)
    except Exception as ex:  # noqa: BLE001
        out['ir_erro'] = repr(ex)[:200]
    out['aba_ativa'] = t(qa.wb.ActiveSheet.Name)
    out['celula_ativa'] = t(qa.ex.xl.ActiveCell.Address)
    qa.ex.xl.EnableEvents = True
    linhas = {}
    vig = B.Caixas(qa.ex.pid, teto=1800)
    vig.start()
    _, rr = C.linha_livre_manual(qa)
    f_eq = t(lo.ListColumns('EQUIPAMENTO').DataBodyRange.Cells(rr, 1).Validation.Formula1)
    f_mz = t(lo.ListColumns('MATRIZ').DataBodyRange.Cells(rr, 1).Validation.Formula1)
    f_mo = t(lo.ListColumns('MOTIVO').DataBodyRange.Cells(rr, 1).Validation.Formula1)
    item_eq = next(x for x in f_eq.split(sep) if eqR in x)
    item_mz = next(x for x in f_mz.split(sep) if t(R2['MATRIZ']) in x) if t(R2['MATRIZ']) else None
    item_mo1 = next(x for x in f_mo.split(sep) if 'Falha do interfaceamento' in x)
    item_mo2 = next(x for x in f_mo.split(sep) if 'Resultado não transmitido' in x)
    out['itens_menu_usados'] = {'EQUIPAMENTO': item_eq, 'MATRIZ': item_mz, 'MOTIVO_M01': item_mo1, 'MOTIVO_M07': item_mo2}
    eq1 = ('EQUIPAMENTO', eqR)

    def novo(rot, campos):
        _, r = C.linha_livre_manual(qa)
        for col, v in campos:
            V.escrever_celula(qa, r, col, v)
            time.sleep(0.2)
        linhas[rot] = r
        return r

    L1 = "'" + lote
    try:
        m01 = [('DATA', sdia2), ('HORA', B.hora_livre(fin, dia2, AN, us2)), ('EQUIPAMENTO', item_eq)]
        if item_mz:
            m01.append(('MATRIZ', item_mz))
        m01 += [('LOTE', "'" + lote2), ('NIVEL', '1'), ('ANALITO', AN), ('RESULTADO', '123,456'), ('MOTIVO', item_mo1)]
        novo('M01_ok_itens_do_menu', m01)
        novo('M02_implausivel', [('DATA', sdia), ('HORA', B.hora_livre(fin, dia, AN, us)), eq1, ('LOTE', L1),
                                 ('NIVEL', '1'), ('ANALITO', AN), ('RESULTADO', '-1'), ('MOTIVO', MOTIVO)])
        novo('M03_lote_FORA', [('DATA', sdia), ('HORA', B.hora_livre(fin, dia, AN, us)), eq1, ('LOTE', "'" + lote_fora),
                               ('NIVEL', '1'), ('ANALITO', AN), ('RESULTADO', '123,456'), ('MOTIVO', MOTIVO)])
        novo('M04_lote_SO', [('DATA', sdia), ('HORA', B.hora_livre(fin, dia, AN, us)), eq1, ('LOTE', "'" + lote_so),
                             ('NIVEL', '1'), ('ANALITO', AN), ('RESULTADO', '123,456'), ('MOTIVO', MOTIVO)])
        novo('M07_subst_ok_motivo_menu', [('SUBSTITUI_ID', numid(S1)), ('RESULTADO', '123,456'), ('MOTIVO', item_mo2)])
        r8 = novo('M08_subst_dup', [('SUBSTITUI_ID', numid(S1))])
        hs1 = S1['DATA_HORA'] + dt.timedelta(minutes=2)
        V.escrever_celula(qa, r8, 'HORA', hs1.strftime('%H:%M'))
        V.escrever_celula(qa, r8, 'RESULTADO', '111,111')
        V.escrever_celula(qa, r8, 'MOTIVO', MOTIVO)
        novo('M09_subst_inexistente', [('SUBSTITUI_ID', '999999999'), ('DATA', sdia),
                                       ('HORA', B.hora_livre(fin, dia, AN, us)), eq1, ('LOTE', L1), ('NIVEL', '2'),
                                       ('ANALITO', AN), ('RESULTADO', '123,456'), ('MOTIVO', MOTIVO)])
        id_m02 = V.ler_manual(qa, linhas['M02_implausivel'])['ID_REGISTRO'][0]
        novo('M10_subst_id_manual', [('SUBSTITUI_ID', id_m02), ('DATA', sdia), ('HORA', B.hora_livre(fin, dia, AN, us)),
                                     eq1, ('LOTE', L1), ('NIVEL', '2'), ('ANALITO', AN), ('RESULTADO', '123,456'),
                                     ('MOTIVO', MOTIVO)])
        r11 = novo('M11_subst_outra_chave', [('SUBSTITUI_ID', numid(S2))])
        V.escrever_celula(qa, r11, 'NIVEL', '2')
        V.escrever_celula(qa, r11, 'RESULTADO', '123,456')
        V.escrever_celula(qa, r11, 'MOTIVO', MOTIVO)
        novo('M12_subst_sem_motivo', [('SUBSTITUI_ID', numid(S3)), ('RESULTADO', '123,456')])
        if dup:
            novo('M13_subst_retransmissao', [('SUBSTITUI_ID', numid(dup[0])), ('RESULTADO', '12,345'), ('MOTIVO', MOTIVO)])
        novo('M14_data_futura', [('DATA', '31/12/2026'), ('HORA', '03:00'), eq1, ('LOTE', L1), ('NIVEL', '1'),
                                 ('ANALITO', AN), ('RESULTADO', '123,456'), ('MOTIVO', MOTIVO)])
        hc = (R['DATA_HORA'] + dt.timedelta(minutes=5)).strftime('%H:%M')
        novo('M15_conflito', [('DATA', sdia), ('HORA', hc), eq1, ('LOTE', L1), ('NIVEL', '1'), ('ANALITO', AN),
                              ('RESULTADO', '111,111'), ('MOTIVO', MOTIVO)])
        hrep = B.hora_livre(fin, dia, AN, us)
        for rot in ('M16_repetido_1', 'M17_repetido_2'):
            novo(rot, [('DATA', sdia), ('HORA', hrep), eq1, ('LOTE', L1), ('NIVEL', '3'), ('ANALITO', AN),
                       ('RESULTADO', '123,456'), ('MOTIVO', MOTIVO)])
        novo('M18_sem_nivel', [('DATA', sdia), ('HORA', B.hora_livre(fin, dia, AN, us)), eq1, ('LOTE', L1),
                               ('ANALITO', AN), ('RESULTADO', '123,456'), ('MOTIVO', MOTIVO)])
        # M19: o usuario copia a linha inteira do M09 e cola na linha livre (Ctrl+C / Ctrl+V), com o ID junto
        try:
            _, r19 = C.linha_livre_manual(qa)
            src = lo.ListRows(linhas['M09_subst_inexistente']).Range
            dst = lo.ListRows(r19).Range
            src.Copy(dst)
            qa.ex.esperar()
            time.sleep(0.5)
            linhas['M19_linha_colada_de_M09'] = r19
            out['M19_colagem'] = 'ok'
        except Exception as ex:  # noqa: BLE001
            out['M19_colagem'] = 'falhou: ' + repr(ex)[:200]
        try:
            qa.ex.xl.CutCopyMode = False
        except Exception:  # noqa: BLE001
            pass
        out['avisos_ao_digitar'] = list(vig.vistos)
        n_av = len(vig.vistos)
        out['pre'] = {k: {c: V.ler_manual(qa, r)[c][0] for c in ('ID_REGISTRO', 'DATA', 'HORA', 'EQUIPAMENTO', 'MATRIZ',
                                                                 'LOTE', 'NIVEL', 'ANALITO', 'RESULTADO', 'MOTIVO',
                                                                 'SUBSTITUI_ID')} for k, r in linhas.items()}
        aud = V.Audit(qa)
        t0 = time.time()
        ret = t(qa.run('mIntegracao.AtualizarDadosAutomatico', teto=1500))
        out['atualizar'] = {'ret': ret[:900], 's': round(time.time() - t0, 1)}
        out['avisos_no_atualizar'] = vig.vistos[n_av:]
        tr = aud.novos()
        out['trilha'] = [{'acao': x['acao'], 'txt': re.sub(r'\b[A-Za-z0-9_.-]+\\[A-Za-z0-9_.-]+', '<usr>', x['txt'])[:400]}
                         for x in tr if x['acao']]
    finally:
        vig.parar.set()
    leit = {k: V.ler_manual(qa, r) for k, r in linhas.items()}
    fin1 = qa.final_por_id()
    res = {}
    for k, v in leit.items():
        idm = v['ID_REGISTRO'][0]
        f = fin1.get(idm, {})
        det_txt = v['DETALHE'][1]
        mot = t(f.get('MOTIVO_EXCLUSAO_AUTOMATICA'))
        res[k] = {'ID': idm, 'STATUS': v['STATUS'][0], 'RUN_txt': v['RUN'][1], 'RUN_v': v['RUN'][0],
                  'DETALHE_txt': det_txt, 'final': {c: t(f.get(c)) for c in ('STATUS_ANALITICO', 'PARTICIPA_ESTATISTICA',
                                                                             'ANALITO', 'MATRIZ', 'EQUIPAMENTO', 'LOTE',
                                                                             'NIVEL', 'RUN')},
                  'MOTIVO_EXCL': mot, 'sem_acento': sem_acento_achados(det_txt),
                  'DETALHE_igual_MOTIVO_acentuado_por_mim': det_txt == acentuar(mot),
                  'MOTIVO_tem_acento': mot != tira_acento(mot),
                  'EQUIPAMENTO_MATRIZ_MOTIVO_tela': [v['EQUIPAMENTO'][0], v['MATRIZ'][0], v['MOTIVO'][0]]}
    out['manuais'] = res
    # M19 (linha colada com o ID junto): as linhas da tblCQ_Final com esse ID (o motor marca a 2a com ~L<n>)
    if 'M19_linha_colada_de_M09' in res:
        i19 = res['M19_linha_colada_de_M09']['ID']
        out['M19_final'] = {k: {c: t(v.get(c)) for c in ('STATUS_ANALITICO', 'PARTICIPA_ESTATISTICA', 'RUN',
                                                         'MOTIVO_EXCLUSAO_AUTOMATICA')}
                            for k, v in fin1.items() if t(k).startswith(i19)}
    s1 = fin1.get(t(S1['ID_REGISTRO']), {})
    out['S1_depois'] = {c: t(s1.get(c)) for c in ('STATUS_ANALITICO', 'PARTICIPA_ESTATISTICA', 'MOTIVO_EXCLUSAO_AUTOMATICA',
                                                  'RUN')}
    # efeito das regras TESTE nos reais dos lotes
    out['reais_nos_lotes_com_regra'] = {
        lt: dict(collections.Counter((t(r['ANALITO']) == AN, t(r['STATUS_ANALITICO']), t(r['PARTICIPA_ESTATISTICA']))
                                     for r in fin1.values() if t(r['LOTE']) == lt and t(r['ORIGEM_RESULTADO']) ==
                                     'INTERFACEAMENTO').most_common(6))
        for lt in (lote_fora, lote_so)}
    out['reais_nos_lotes_com_regra'] = {lt: [[list(k), n] for k, n in d.items()] for lt, d in
                                        out['reais_nos_lotes_com_regra'].items()}
    ex_motivo = collections.Counter(t(r['MOTIVO_EXCLUSAO_AUTOMATICA']) for r in fin1.values()
                                    if t(r['LOTE']) in (lote_fora, lote_so) and t(r['MOTIVO_EXCLUSAO_AUTOMATICA']))
    out['motivos_reais_nos_lotes_com_regra'] = ex_motivo.most_common(5)
    # QA_INTEGRACAO: I04/E12/E04 que citam os manuais
    try:
        ids = {v['ID'] for v in res.values() if v['ID']}
        qai = qa.ler('tblQA_Integracao')
        cit = []
        for x in qai:
            txt = ' | '.join(t(v) for v in x.values() if v not in (None, ''))
            if any(i in txt for i in ids) or t(S1['ID_REGISTRO']) in txt:
                cit.append(txt[:300])
        out['qa_integracao_manuais'] = cit[:40]
    except Exception as ex:  # noqa: BLE001
        out['qa_integracao_manuais'] = repr(ex)[:160]
    # varredura da aba (RUN/DETALHE "0", palavras sem acento)
    n = lo.ListRows.Count
    vr, vd = lo.ListColumns('RUN').DataBodyRange, lo.ListColumns('DETALHE').DataBodyRange
    vi = lo.ListColumns('ID_REGISTRO').DataBodyRange.Value
    zeros, sa = [], []
    for i in range(1, min(n, 400) + 1):
        if vi[i - 1][0] in (None, ''):
            continue
        a, b = t(vr.Cells(i, 1).Text), t(vd.Cells(i, 1).Text)
        if a == '0' or b == '0':
            zeros.append((t(vi[i - 1][0]), a, b))
        if sem_acento_achados(b):
            sa.append((t(vi[i - 1][0]), sem_acento_achados(b), b[:160]))
    out['varredura'] = {'linhas_com_id': sum(1 for x in vi if x[0] not in (None, '')), 'com_zero': zeros,
                        'com_palavra_sem_acento': sa}
    # Painel DEPOIS (OC-0068)
    out['serie_depois'] = painel_serie(qa, lote2, dia2.year, dia2.month)
    a, d = out['serie_antes'], out['serie_depois']
    out['M01_serie'] = {'dia': sdia2, 'lote': lote2,
                        'n_N1_antes_depois': [a['engPainel'][0][1], d['engPainel'][0][1]],
                        'media_N1_antes_depois': [a['engPainel'][0][2], d['engPainel'][0][2]],
                        'N2_N3_iguais': a['engPainel'][1:] == d['engPainel'][1:],
                        'valor_no_grafico_antes': a['tem_123456_N1'], 'valor_no_grafico_depois': d['tem_123456_N1']}
    return out


# ------------------------------------------------------------------------------------------- execucao
def rodar(rotulo):
    origem = BACKUP if rotulo == 'base41' else os.path.join(C.REPO, f'QC_{PROD}.xlsm')
    destino = os.path.join(PASTA, f'{PROD}_F1_{rotulo}.xlsm')
    os.makedirs(PASTA, exist_ok=True)
    if os.path.exists(destino):
        os.remove(destino)
    h0 = C.sha(origem)
    shutil.copy2(origem, destino)
    L = C.Log(os.path.join(PASTA, f'{PROD}_F1_{rotulo}.json'))
    L['ambiente'] = {'origem': os.path.relpath(origem, C.REPO), 'sha_origem_12': h0, 'copia': os.path.basename(destino),
                     'sha_copia_12': C.sha(destino)}
    senha = C.senha_planilha()
    qa = C.qa_final.QA(PROD, destino)
    L['excel_pid'] = qa.ex.pid
    try:
        p, e = qa.ws('Painel'), qa.ws('Estatística')
        p.Activate()
        qa.ex.esperar()
        fin = qa.ler('tblCQ_Final')
        L['linhas_tblCQ_Final'] = len(fin)
        L['ultimo_recebido'] = str(max((r['DATA_HORA'] for r in fin if isinstance(r['DATA_HORA'], dt.datetime)),
                                       default=None))
        criar_usuario(qa, USR, 'ANALISTA')
        L['login'] = fazer_login(qa, USR)
        for w in (p, e):
            try:
                w.Unprotect(senha)
            except Exception:  # noqa: BLE001
                pass
        sep = t(qa.ex.xl.International[4])
        secao(L, 'O72_apos_login', alvos_0072, qa)
        secao(L, 'O68_listas', s_listas, qa, fin)
        secao(L, 'O68_O72_registros', B.s_registros, qa, senha)
        secao(L, 'O21_O68_manuais', s_manuais, qa, fin, sep, senha)
        secao(L, 'O72_depois_atualizar', alvos_0072, qa)
        if rotulo == 'nova':
            sc = os.path.join(PASTA, f'{PROD}_F1_nova_SaveCopyAs.xlsm')
            if os.path.exists(sc):
                os.remove(sc)
            qa.wb.SaveCopyAs(sc)
            r = X.resumo(sc)
            L['O72_O68_xml_depois_do_uso'] = {'arquivo': os.path.basename(sc),
                                              'numFmt_com_letra_literal': r['numFmt_com_letra_literal'],
                                              'uso_literais': r['uso_literais'], 'x12ac_list': r['x12ac_list'],
                                              'formatos_de_data': r['formatos_de_data'],
                                              'validacoes_DR_Registros_Painel': [x for x in r['validacoes_lista'] if
                                                                                 x[0] in ('Digitar Resultados', 'Registros',
                                                                                          'Painel')]}
    except Exception as ex:  # noqa: BLE001
        L['ERRO_FATAL'] = {'erro': repr(ex), 'trace': traceback.format_exc()[-3000:]}
    finally:
        try:
            qa.fechar(salvar=False)
        except Exception:  # noqa: BLE001
            pass
    L['origem_intacta'] = {'sha_antes': h0, 'sha_depois': C.sha(origem), 'igual': h0 == C.sha(origem)}
    print('FIM', rotulo, flush=True)


def _col(ref):
    return re.match(r'[A-Z]+', ref).group(0)


def _lin(ref):
    return int(re.search(r'\d+', ref).group(0))


ALVOS_HEMA = {
    'Eventos_Westgard': lambda r: _col(r) == 'A' and 4 <= _lin(r) <= 5003,
    'Configuração': lambda r: _col(r) in ('D', 'AD') and 26 <= _lin(r) <= 125,
    'Eng_Saida': lambda r: _col(r) == 'AC' and 3 <= _lin(r) <= 182,
    'Registros': lambda r: r in ('E4', 'F4'),
    'Estatística': lambda r: r in ('E8', 'F8', 'E9', 'F9', 'C10'),
    'Sync_Log': lambda r: _col(r) in ('A', 'B', 'L', 'M') and _lin(r) >= 2,
    'Westgard_Fotos': lambda r: _col(r) == 'F',
    'LotesStore': lambda r: r == 'U1',
}


def xml():
    """Sem Excel: copia base41 e nova, varre o XML e compara os valores <v> dos alvos da OC-0072."""
    os.makedirs(PASTA, exist_ok=True)
    out = {}
    arq = {}
    for rot, origem in (('base41', BACKUP), ('nova', os.path.join(C.REPO, f'QC_{PROD}.xlsm'))):
        h0 = C.sha(origem)
        dst = os.path.join(PASTA, f'{PROD}_F1_xml_{rot}.xlsm')
        shutil.copy2(origem, dst)
        arq[rot] = dst
        r = X.resumo(dst)
        out[rot] = {'sha_origem_12': h0, 'sha_depois_12': C.sha(origem),
                    'numFmt_com_letra_literal': r['numFmt_com_letra_literal'], 'uso_literais': r['uso_literais'],
                    'x12ac_list': r['x12ac_list'], 'formatos_de_data': r['formatos_de_data'],
                    'validacoes_lista_DR_Registros_Painel': [x for x in r['validacoes_lista']
                                                             if x[0].split(' (')[0] in ('Digitar Resultados', 'Registros',
                                                                                        'Painel')]}
    a = {k: X.alvos(arq[k], ALVOS_HEMA) for k in arq}
    comp = {}
    for aba in ALVOS_HEMA:
        b, n = a['base41'].get(aba, {}), a['nova'].get(aba, {})
        comuns = set(b) & set(n)
        dif = [(r, b[r][0], n[r][0]) for r in sorted(comuns) if b[r][0] != n[r][0]]
        comp[aba] = {'n_base': len(b), 'n_nova': len(n), 'n_comuns': len(comuns), 'valores_diferentes': len(dif),
                     'dif_amostra': dif[:8],
                     'so_base_com_valor': [r for r in set(b) - set(n) if b[r][0] not in (None, '')][:8],
                     'so_nova_com_valor': [r for r in set(n) - set(b) if n[r][0] not in (None, '')][:8],
                     'formatos_base': dict(collections.Counter(x[1] for x in b.values())),
                     'formatos_nova': dict(collections.Counter(x[1] for x in n.values())),
                     'nao_vazios_nova': sum(1 for x in n.values() if x[0] not in (None, ''))}
    comp['_cols_Westgard_Fotos'] = {k: a[k].get('_cols_Westgard_Fotos') for k in a}
    out['alvos_comparacao'] = comp
    json.dump(out, open(os.path.join(PASTA, f'{PROD}_F1_xml.json'), 'w', encoding='utf-8'), ensure_ascii=False, indent=1,
              default=str)
    for k in arq:
        os.remove(arq[k])
    print(json.dumps({k: (v if k != 'alvos_comparacao' else {a2: {c: v[a2][c] for c in ('n_comuns', 'valores_diferentes',
                                                                                         'formatos_base', 'formatos_nova')}
                                                             for a2 in ALVOS_HEMA})
                      for k, v in out.items()}, ensure_ascii=False, default=str)[:6000])


if __name__ == '__main__':
    t0 = time.time()
    if sys.argv[1] == 'xml':
        xml()
    else:
        rodar(sys.argv[1])
    print(f'tempo {time.time() - t0:.0f}s', flush=True)
