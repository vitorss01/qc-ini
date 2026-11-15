# -*- coding: utf-8 -*-
"""reteste04_usuario_f2_hema.py -- Usuario Simulado, ciclo 2026-11-14_reteste04, frente F2, modo RETESTE (Hematologia).

Irmao do reteste04_usuario_f2.py (Bioquimica). Refaz numa COPIA NOVA da producao atual (entrega 41) o fluxo original de
OC-0049, OC-0057, OC-0058, OC-0066/ADR-094 e ADR-091, com o fluxo vizinho de cada um, e o pedido do ciclo: incluir e
corrigir um resultado de CEQ TESTE-AUDITORIA e conferir que bias, ET e Sigma do analito certo (e so dele) mudam.

Secoes (uma copia, um Excel, nada salvo na copia de trabalho; retratos SaveCopyAs com eventos desligados):
  O49        (arquivo recem-aberto) Estatistica A98:AU133 (linhas sem analito): nenhum texto; formulas A/B/G/R/S/T;
             linhas 14..97 com analito e nivel 1/2/3. Vizinho: um analito TESTE-AUDITORIA na Especificacoes!A32
             (como o administrador cadastraria; eventos desligados, Calculate) -> linhas 98..100 mostram o nome e o
             nivel; apagar -> vazias de novo; A14:AU97 igual bit a bit antes/depois.
  ceq        (login ANALISTA de teste) HGB / CAP "Hemoglobin", alvo 13: S0 -> SalvarCEQ 14,3 (+10 %) -> S1 ->
             CorrigirCEQ 13,65 (+5 %) -> S2 -> vizinhos: Controllab do mesmo analito sem mapa; linha de mapa
             TESTE-AUDITORIA do Controllab para HGB (EQA_Base W:Y, administrador) + amostra Controllab 26/13 (+100 %);
             CAP 2024 (+100 %) com Ano EP 2025 -> S3 (Estatistica nao pode mudar: R4 = CAP) -> R4 = Controllab
             (filtro do usuario) -> RX (so o HGB tem bias, os outros SEM EP) -> R4 = CAP (volta ao S3) ->
             duplicidade -> ExcluirCEQ de tudo + mapa apagado -> S4 (= S0 bit a bit).
             Diferenca celula a celula da Estatistica A14:AU133 e Painel B7:J9 (C3 = HGB).
  O57        Salvar com sessao -> Logout -> Salvar/Corrigir/Excluir devolvem ERRO|0 e nada muda; AbrirCEQ sem login
             (caixa lida pelo texto da janela do PID); vizinho: login de novo -> Corrigir e Excluir funcionam.
  adr091     tblLotesMaterial da Hema (vazia na producao) + baseline (nenhum LOTE_*). Regras TESTE-AUDITORIA na copia,
             como o administrador (aba Cfg_Integracao): 533811 FORA_DO_CQ, 536411 SO_ANALITOS "WBC".
  manuais    Digitar Resultados (navegacao pelo botao, eventos ligados): -1 (IMPLAUSIVEL), 533811 WBC (FORA_DO_CQ),
             533811 -1 (precedencia LOTE_ antes de IMPLAUSIVEL), 536411 RBC (LOTE_OUTRO_MATERIAL), 536411 WBC
             (ATIVO), 625311 WBC (ATIVO, controle) + inativar um ponto real (InativarPorId) + comentario tecnico
             (GravarComentarioTecnico) -> ATUALIZAR DADOS (AtualizarDadosAutomatico; DB_SEAC so lido) -> trilha por
             manual (OC-0058), QA A13/A14, oraculo proprio das regras sobre a tblCQ_Final -> retrato R1 ->
             2o ATUALIZAR (so entradas) sem trilha repetida -> Painel nos lotes com regra.
  adr094     varredura de todas as tabelas (funcao s_adr094 do reteste04_usuario_f2.py): nenhum "####", data-hora com
             segundos; DATA_INATIVACAO da Hema (coluna VAZIA na producao: o cenario exato da OC-0066) antes e depois.
  calc       (sem Excel) calc_independente.auditar_hema sobre os retratos.

Regras: so COPIA em %TEMP%/qcw/auditoria/2026-11-14_reteste04/; fechar() encerra so o PID do QA; nenhuma captura de
tela. Dados de teste: rodada "TESTE-AUDITORIA F2H R1", amostras TA-F2H-*, arquivo TESTE-AUDITORIA.pdf; manuais com
MOTIVO "Outro: TESTE-AUDITORIA"; regras da tblLotesMaterial com NOTA TESTE-AUDITORIA; analito "TESTE-AUDITORIA X".

Uso: python -u reteste04_usuario_f2_hema.py            (Excel + calc)
     python -u reteste04_usuario_f2_hema.py --parte2   (2a copia nova: regras + manuais + ATUALIZAR, ADR-094, Painel)
     python -u reteste04_usuario_f2_hema.py --so-calc  (so a calculadora sobre os retratos ja gravados)
Saida: %TEMP%/qcw/auditoria/2026-11-14_reteste04/Hematologia_F2.json (+ Hematologia_F2_calc.json)
"""
import collections
import datetime as dt
import json
import os
import re
import sys
import time
import traceback

AQUI = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, AQUI)
_argv = sys.argv[:]
sys.argv = [sys.argv[0], 'Hematologia']
import valida_adr090_091 as V  # noqa: E402  (infraestrutura: Audit, Vigia, ler/escrever manual, abrir_ceq_sem_sessao)
sys.argv = _argv
import repro_comum as C  # noqa: E402
import reteste04_usuario_f2 as B  # noqa: E402  (helpers: estat_all, diff_estat, linhas_an, por, s_adr094)
from qa_inativar_ux import criar_usuario, fazer_login  # noqa: E402

assert not V.BIO
PROD = 'Hematologia'
NLV = 3
CICLO = '2026-11-14_reteste04'
C.PASTA = os.path.join(os.environ['TEMP'], 'qcw', 'auditoria', CICLO)
USR = 'TESTE_AUDITORIA_F2H'
TAG = 'TESTE-AUDITORIA'
MOTIVO = 'Outro: TESTE-AUDITORIA'
AN, NOME, UN = 'HGB', 'Hemoglobin', 'g/dL'
ROD = 'TESTE-AUDITORIA F2H R1'
NOME_CTL = 'TESTE-AUDITORIA HGB CTL'
L_FORA, L_SO, L_SO_AN = '533811', '536411', 'WBC'
t, up, num = V.t, V.up, V.num
RETR = {}
por = B.por


def secao(L, nome, fn, *a):
    print('>>', nome, time.strftime('%H:%M:%S'), flush=True)
    return V.secao(L, nome, fn, *a)


def retrato(qa, nome):
    caminho = os.path.join(C.PASTA, f'{PROD}_F2_{nome}.xlsm')
    ev = qa.ex.xl.EnableEvents
    qa.ex.xl.EnableEvents = False
    try:
        qa.wb.SaveCopyAs(caminho)
    finally:
        qa.ex.xl.EnableEvents = ev
    RETR[nome] = caminho
    return os.path.basename(caminho)


def painel_an(qa):
    p = qa.ws('Painel')
    v = p.Range('A7:J9').Value
    cab = ['nivel', 'n', 'media', 'dp', 'cv', 'etp', 'bias', 'et', 'sigma', 'status']
    return {'C3': t(p.Range('C3').Value), 'H3': t(p.Range('H3').Value),
            'niveis': [dict(zip(cab, [str(x) if isinstance(x, dt.datetime) else x for x in l])) for l in v]}


def campos(prov='CAP', an=NOME, res='14,3', alvo='13', am='TA-F2H-01', ano='2025', rod=ROD):
    return C.campos_ceq(prov, rod, ano, an, am, res, alvo=alvo, sd='0,2', sdi='1', linf='12,2', lsup='13,8',
                        aval='Acceptable', nlabs='50', unid=UN, pag='1', arq=f'{TAG}.pdf')


def base_teste(qa):
    b = qa.ws('EQA_Base')
    v = b.Range('A1:U5001').Value
    cab = [t(x) for x in v[0]]
    return [{k: (x if not isinstance(x, dt.datetime) else str(x)) for k, x in zip(cab, r) if x not in (None, '')}
            for r in v[1:] if any(TAG in t(x) for x in r)]


def resumo_estat(qa):
    e = qa.ws('Estatística')
    return {a: t(e.Range(a).Text)[:200] for a in ('B3', 'D3', 'B4', 'N4', 'P4', 'R4', 'K5')}


# ================================================================================================ OC-0049
def s_o49(qa, senha):
    e = qa.ws('Estatística')
    out = {}
    tx = e.Range('A98:AU133')
    vals = tx.Value
    textos = []
    for i in range(36):
        for j in range(47):
            c = e.Cells(98 + i, 1 + j)
            s = t(c.Text)
            if s:
                textos.append(f'{B.col_letra(j)}{98 + i}="{s}"')
    out['A98_AU133_textos'] = textos[:40]
    out['A98_AU133_n_textos'] = len(textos)
    out['A98_AU133_valores_nao_vazios'] = sum(1 for l in vals for x in l if x not in (None, ''))
    out['formulas_98'] = {c: t(e.Range(f'{c}98').Formula)[:120] for c in ('A', 'B', 'C', 'G', 'K', 'R', 'S', 'T',
                                                                          'AF', 'AR')}
    est = B.estat_all(qa)
    com = [r for r in est[:84]]
    out['linhas_14_97'] = {'com_analito': sum(1 for r in com if t(r[0])),
                           'niveis_ok': all(r[1] == (i % 3) + 1 for i, r in enumerate(com)),
                           'analitos': sorted({t(r[0]) for r in com})}
    out['retrato'] = retrato(qa, 'E0')
    # vizinho: cadastrar um analito na 1a linha livre do catalogo (Especificacoes!A32) e apagar
    esp = qa.espec()
    a32 = esp.Range('A32')
    out['Especificacoes_A31_A32'] = [t(esp.Range('A31').Value), t(a32.Value), t(a32.Formula)[:60]]
    if not t(a32.Value) and not t(a32.Formula):
        try:
            esp.Unprotect(senha)
        except Exception:
            pass
        ev = qa.ex.xl.EnableEvents
        qa.ex.xl.EnableEvents = False
        try:
            a32.Value = 'TESTE-AUDITORIA X'
            qa.ex.xl.Calculate()
            out['vizinho_com_analito'] = [[t(e.Cells(r, c).Text) for c in (1, 2, 3, 7, 8, 11, 12, 18, 19, 20)]
                                          for r in (98, 99, 100, 101)]
            a32.ClearContents()
            qa.ex.xl.Calculate()
        finally:
            qa.ex.xl.EnableEvents = ev
        textos2 = [f'{B.col_letra(j)}{98 + i}' for i in range(36) for j in range(47)
                   if t(e.Cells(98 + i, 1 + j).Text)]
        out['vizinho_apagado_textos'] = textos2[:20]
        out['A14_AU97_igual_antes_depois'] = B.estat_all(qa)[:84] == est[:84]
    out['ok'] = (not textos and out['linhas_14_97']['com_analito'] == 84 and out['linhas_14_97']['niveis_ok'])
    v = out.get('vizinho_com_analito')
    out['ok_vizinho'] = bool(v) and v[0][0] == 'TESTE-AUDITORIA X' and [x[1] for x in v[:3]] == ['1', '2', '3'] \
        and not v[3][0] and not out.get('vizinho_apagado_textos') and out.get('A14_AU97_igual_antes_depois')
    return out


# ================================================================================================ CEQ
def s_ceq(qa, aud, senha):
    out = {}
    e, p = qa.ws('Estatística'), qa.ws('Painel')
    b = qa.ws('EQA_Base')
    mapa = [(t(m[0]), t(m[1]), t(m[2])) for m in C.bloco(b, 'W2:Y600') if t(m[1])]
    out['mapa_CAP_HGB'] = [nm for pv, nm, cn in mapa if up(pv) == 'CAP' and cn == AN]
    out['mapa_Controllab'] = [(nm, cn) for pv, nm, cn in mapa if up(pv) == 'CONTROLLAB'][:10]
    out['filtro'] = resumo_estat(qa)
    ano = t(e.Range('N4').Value)
    ano = ano if re.fullmatch(r'\d{4}', ano) else '2025'
    out['ano'] = ano
    por(qa, p, 'C3', AN)
    qa.ex.xl.Calculate()
    E0 = B.estat_all(qa)
    out['S0'] = {'estat': B.linhas_an(E0, AN), 'painel': painel_an(qa), 'K5': t(e.Range('K5').Text)[:200],
                 'retrato': retrato(qa, 'S0')}
    aud.novos()
    # ---- incluir (a janela: Aviso -> Confirmacoes -> Salvar)
    out['aviso_analito'] = t(qa.run('mCEQForm.AvisoAnalitoCEQ', 'CAP', NOME))
    out['confirmacoes'] = t(qa.run('mCEQForm.ConfirmacoesCEQ', campos(ano=ano), 0))
    r1 = t(qa.run('mCEQForm.SalvarCEQ', campos(ano=ano)))
    lin = int(r1.split('|')[1]) if r1.startswith('OK|') else 0
    E1a = B.estat_all(qa)
    qa.ex.xl.Calculate()
    E1 = B.estat_all(qa)
    cap = qa.ws('EQA.CAP_Dados')
    out['S1'] = {'salvar': r1, 'linha': lin,
                 'gravado_F_G': [cap.Range(f'F{lin}').Value, cap.Range(f'G{lin}').Value] if lin else None,
                 'descricao': t(qa.run('mCEQForm.DescricaoCEQ', 'CAP', lin)) if lin else '',
                 'lista_rodada': [[t(x) for x in l] for l in (qa.run('mCEQForm.RegistrosDaRodada', 'CAP', ROD, ano)
                                                               or [])],
                 'base_teste': base_teste(qa), 'Z3': t(cap.Range('Z3').Value)[:160],
                 'estat_sem_Calculate_igual_com': E1a == E1, 'K5': t(e.Range('K5').Text)[:200],
                 'estat': B.linhas_an(E1, AN), 'painel': painel_an(qa), 'diff_S0': B.diff_estat(E0, E1),
                 'trilha': [x['txt'][:400] for x in aud.novos()], 'retrato': retrato(qa, 'S1')}
    # ---- corrigir (mesma linha; 13,65 = +5 %)
    r2 = t(qa.run('mCEQForm.CorrigirCEQ', lin, campos(ano=ano, res='13,65'))) if lin else 'sem linha'
    qa.ex.xl.Calculate()
    E2 = B.estat_all(qa)
    out['S2'] = {'corrigir': r2, 'gravado_F': cap.Range(f'F{lin}').Value if lin else None,
                 'base_teste': base_teste(qa), 'estat': B.linhas_an(E2, AN), 'painel': painel_an(qa),
                 'diff_S1': B.diff_estat(E1, E2), 'trilha': [x['txt'][:500] for x in aud.novos()],
                 'retrato': retrato(qa, 'S2')}
    # ---- vizinhos: Controllab sem mapa, Controllab com mapa de teste, CAP de outro ano
    extra = []
    out['aviso_ctl_sem_mapa'] = t(qa.run('mCEQForm.AvisoAnalitoCEQ', 'Controllab', NOME))
    rc0 = t(qa.run('mCEQForm.SalvarCEQ', campos(prov='Controllab', res='26', am='TA-F2H-CTL0', ano=ano)))
    if rc0.startswith('OK|'):
        extra.append(('Controllab', int(rc0.split('|')[1])))
    b.Unprotect(senha)
    v = b.Range('W2:Y600').Value
    livre = next(i for i, l in enumerate(v, 2) if not any(l))
    ev = qa.ex.xl.EnableEvents
    qa.ex.xl.EnableEvents = False
    try:
        b.Range(f'W{livre}:Y{livre}').Value = (('Controllab', NOME_CTL, AN),)
    finally:
        qa.ex.xl.EnableEvents = ev
    out['mapa_teste_linha'] = livre
    out['aviso_ctl_mapeado'] = t(qa.run('mCEQForm.AvisoAnalitoCEQ', 'Controllab', NOME_CTL))
    rc = t(qa.run('mCEQForm.SalvarCEQ', campos(prov='Controllab', an=NOME_CTL, res='26', am='TA-F2H-CTL', ano=ano)))
    if rc.startswith('OK|'):
        extra.append(('Controllab', int(rc.split('|')[1])))
    ano_out = str(int(ano) - 1)
    ra = t(qa.run('mCEQForm.SalvarCEQ', campos(res='26', am='TA-F2H-ANO', ano=ano_out)))
    if ra.startswith('OK|'):
        extra.append(('CAP', int(ra.split('|')[1])))
    qa.ex.xl.Calculate()
    E3 = B.estat_all(qa)
    out['S3_vizinhos'] = {'controllab_sem_mapa': rc0, 'controllab_mapeado': rc, 'cap_outro_ano': [ano_out, ra],
                          'base_teste': base_teste(qa), 'diff_S2': B.diff_estat(E2, E3),
                          'estat': B.linhas_an(E3, AN), 'trilha': [x['acao'] for x in aud.novos()],
                          'retrato': retrato(qa, 'S3')}
    # ---- o usuario troca o provedor global (R4) para Controllab e volta
    r4 = t(e.Range('R4').Value)
    out['R4_antes'] = r4
    out['R4_controllab'] = por(qa, e, 'R4', 'Controllab')
    qa.ex.xl.Calculate()
    EX = B.estat_all(qa)
    dx = B.diff_estat(E3, EX)
    out['RX_controllab'] = {'R4': t(e.Range('R4').Value), 'K5': t(e.Range('K5').Text)[:200],
                            'estat': B.linhas_an(EX, AN), 'painel': painel_an(qa),
                            'G_por_analito': {f'{t(r[0])}|{t(r[1])}': r[6] for r in EX if t(r[0])},
                            'n_celulas_mudadas': len(dx), 'analitos_mudados': sorted({x['analito'] for x in dx}),
                            'retrato': retrato(qa, 'RX')}
    out['R4_volta'] = por(qa, e, 'R4', r4 or 'CAP')
    qa.ex.xl.Calculate()
    E3b = B.estat_all(qa)
    out['R4_volta_igual_S3'] = E3b == E3
    # ---- duplicidade
    out['duplicidade'] = t(qa.run('mCEQForm.SalvarCEQ', campos(ano=ano, res='13,65')))[:400]
    # ---- excluir tudo
    exc = {}
    for prov, ln in sorted(extra, key=lambda x: -x[1]):
        exc[f'{prov}|{ln}'] = t(qa.run('mCEQForm.ExcluirCEQ', prov, ln))[:160]
    exc[f'CAP|{lin}'] = t(qa.run('mCEQForm.ExcluirCEQ', 'CAP', lin))[:160] if lin else ''
    qa.ex.xl.EnableEvents = False
    try:
        b.Range(f'W{livre}:Y{livre}').ClearContents()
    finally:
        qa.ex.xl.EnableEvents = ev
    qa.ex.xl.Calculate()
    E4 = B.estat_all(qa)
    out['S4_excluido'] = {'excluir': exc, 'base_teste': base_teste(qa), 'diff_S0': B.diff_estat(E0, E4),
                          'igual_S0_bit_a_bit': E0 == E4, 'painel': painel_an(qa), 'K5': t(e.Range('K5').Text)[:200],
                          'trilha': [x['txt'][:300] for x in aud.novos()], 'retrato': retrato(qa, 'S4')}
    return out


# ================================================================================================ OC-0057
def s_sem_sessao(qa, aud):
    out = {}
    r = t(qa.run('mCEQForm.SalvarCEQ', campos(am='TA-F2H-SES', res='14,3')))
    lin = int(r.split('|')[1]) if r.startswith('OK|') else 0
    out['registro_com_sessao'] = r[:120]
    cap = qa.ws('EQA.CAP_Dados')
    qa.run('mSeguranca.Logout')
    qa.ex.esperar()
    out['SessaoAtiva'] = qa.run('mSeguranca.SessaoAtiva')
    aud.novos()
    n0 = qa.lo('tblEQA_CAP_Dados').ListRows.Count
    f0 = cap.Range(f'F{lin}').Value if lin else None
    out['salvar'] = t(qa.run('mCEQForm.SalvarCEQ', campos(am='TA-F2H-SES2')))[:160]
    out['corrigir'] = t(qa.run('mCEQForm.CorrigirCEQ', lin, campos(am='TA-F2H-SES', res='13'))) [:160] if lin else ''
    out['F_antes_depois'] = [f0, cap.Range(f'F{lin}').Value if lin else None]
    out['excluir'] = t(qa.run('mCEQForm.ExcluirCEQ', 'CAP', lin))[:160] if lin else ''
    out['A_depois'] = t(cap.Range(f'A{lin}').Value) if lin else ''
    out['linhas_tabela'] = [n0, qa.lo('tblEQA_CAP_Dados').ListRows.Count]
    out['trilha'] = [x['acao'] for x in aud.novos()]
    out['abrir_sem_login'] = V.abrir_ceq_sem_sessao(qa)
    out['relogin'] = fazer_login(qa, USR)
    aud.novos()
    out['corrigir_com_sessao'] = t(qa.run('mCEQForm.CorrigirCEQ', lin, campos(am='TA-F2H-SES', res='13')))[:160] \
        if lin else ''
    out['F_apos_corrigir'] = cap.Range(f'F{lin}').Value if lin else None
    out['excluir_com_sessao'] = t(qa.run('mCEQForm.ExcluirCEQ', 'CAP', lin))[:160] if lin else ''
    out['trilha_com_sessao'] = [x['acao'] for x in aud.novos()]
    out['base_teste_fim'] = base_teste(qa)
    out['ok'] = (out['salvar'].startswith('ERRO|0|') and out['corrigir'].startswith('ERRO|0|')
                 and out['excluir'].startswith('ERRO|0|') and out['F_antes_depois'][0] == out['F_antes_depois'][1]
                 and out['linhas_tabela'][0] == out['linhas_tabela'][1]
                 and not [a for a in out['trilha'] if a and a.startswith('CEQ_')] and out['abrir_sem_login']['ok'])
    out['ok_vizinho'] = (out['corrigir_com_sessao'].startswith('OK|') and out['excluir_com_sessao'].startswith('OK|'))
    return out


# ================================================================================================ ADR-091 + manuais + ATUALIZAR
def cols_final(qa):
    lo = qa.lo('tblCQ_Final')
    c = lo.ListColumns('DATA_INATIVACAO')
    rng = c.DataBodyRange
    v = rng.Value
    preench = [i for i, x in enumerate(v) if x[0] not in (None, '')]
    return {'largura': rng.Cells(1, 1).ColumnWidth, 'fmt': t(rng.Cells(1, 1).NumberFormat),
            'fmt_ultima': t(rng.Cells(len(v), 1).NumberFormat), 'preenchidas': len(preench),
            'textos': [t(rng.Cells(i + 1, 1).Text) for i in preench[:5]]}


def oraculo_regras(fin):
    precede = set(V.PRECEDE_LOTE)
    out, viol = {}, []
    rs = [r for r in fin if t(r['LOTE']) == L_FORA]
    bad = [r['ID_REGISTRO'] for r in rs if r['STATUS_ANALITICO'] != 'LOTE_FORA_DO_CQ'
           and r['STATUS_ANALITICO'] not in precede]
    part = [r['ID_REGISTRO'] for r in rs if r['PARTICIPA_ESTATISTICA'] == 'SIM']
    out[f'FORA_DO_CQ {L_FORA}'] = {'linhas': len(rs), 'status': dict(collections.Counter(r['STATUS_ANALITICO']
                                                                                         for r in rs))}
    if bad or part:
        viol.append({'lote': L_FORA, 'status_errado': bad[:5], 'participa': part[:5]})
    rs = [r for r in fin if t(r['LOTE']) == L_SO]
    fora = [r for r in rs if up(r['ANALITO']) != L_SO_AN]
    dentro = [r for r in rs if up(r['ANALITO']) == L_SO_AN]
    bad = [r['ID_REGISTRO'] for r in fora if r['STATUS_ANALITICO'] != 'LOTE_OUTRO_MATERIAL'
           and r['STATUS_ANALITICO'] not in precede]
    bad += [r['ID_REGISTRO'] for r in fora if r['PARTICIPA_ESTATISTICA'] == 'SIM']
    bad2 = [r['ID_REGISTRO'] for r in dentro if r['STATUS_ANALITICO'].startswith('LOTE_')]
    out[f'SO_ANALITOS {L_SO}'] = {'linhas': len(rs),
                                  'dentro': dict(collections.Counter(r['STATUS_ANALITICO'] for r in dentro)),
                                  'fora': dict(collections.Counter(r['STATUS_ANALITICO'] for r in fora))}
    if bad or bad2:
        viol.append({'lote': L_SO, 'fora_da_lista_errado': bad[:5], 'da_lista_excluido': bad2[:5]})
    outros = [r['ID_REGISTRO'] for r in fin if r['STATUS_ANALITICO'].startswith('LOTE_')
              and t(r['LOTE']) not in (L_FORA, L_SO)]
    if outros:
        viol.append({'LOTE_em_lote_sem_regra': outros[:5]})
    out['total_LOTE'] = dict(collections.Counter(r['STATUS_ANALITICO'] for r in fin
                                                 if r['STATUS_ANALITICO'].startswith('LOTE_')))
    out['violacoes'] = viol
    return out


def s_regras_manuais(qa, fin, aud, senha, L):
    out = {}
    p = qa.ws('Painel')
    lm = qa.lo('tblLotesMaterial')
    out['tblLotesMaterial_producao'] = {'linhas': lm.ListRows.Count,
                                        'cabecalho': [t(c.Name) for c in lm.ListColumns],
                                        'valores': [[t(x) for x in l] for l in (lm.DataBodyRange.Value or [])]}
    out['baseline_LOTE'] = dict(collections.Counter(r['STATUS_ANALITICO'] for r in fin
                                                    if r['STATUS_ANALITICO'].startswith('LOTE_')))
    out['DATA_INATIVACAO_antes'] = cols_final(qa)
    out['participa_antes'] = {lt: collections.Counter(r['PARTICIPA_ESTATISTICA'] for r in fin if t(r['LOTE']) == lt)
                              for lt in (L_FORA, L_SO)}
    # ---- regras como o administrador (aba Cfg_Integracao, tabela editavel)
    ws = qa.ws('Cfg_Integracao')
    try:
        ws.Unprotect(senha)
    except Exception:
        pass
    cab = [t(c.Name) for c in lm.ListColumns]
    regras = [{'LOTE': L_FORA, 'REGRA': 'FORA_DO_CQ', 'NOTA': f'{TAG}: lote fora do CQ (teste)'},
              {'LOTE': L_SO, 'REGRA': 'SO_ANALITOS', 'ANALITOS': L_SO_AN, 'NOTA': f'{TAG}: so WBC (teste)'}]
    for i, rg in enumerate(regras):
        vazio = lm.ListRows.Count >= i + 1 and not any(t(x) for x in lm.ListRows(i + 1).Range.Value[0])
        linha = lm.ListRows(i + 1) if vazio else lm.ListRows.Add()
        for k, v in rg.items():
            cel = linha.Range.Cells(1, cab.index(k) + 1)
            cel.NumberFormat = '@' if k == 'LOTE' else cel.NumberFormat
            cel.Value = v
            qa.ex.esperar()
    out['tblLotesMaterial_teste'] = [[t(x) for x in l] for l in lm.DataBodyRange.Value]
    # ---- manuais (Digitar Resultados, navegacao pelo botao)
    por(qa, p, 'C3', 'WBC')
    try:
        qa.run('mIntegracao.IrDigitarResultados', teto=120)
    except Exception as ex:  # noqa: BLE001
        out['ir_erro'] = repr(ex)[:120]
    out['aba_ativa'] = t(qa.wb.ActiveSheet.Name)

    def real(lote, an='WBC', k=-30):
        rr = sorted([r for r in fin if r['STATUS_ANALITICO'] == 'ATIVO' and r['ORIGEM_RESULTADO'] == 'INTERFACEAMENTO'
                     and t(r['ANALITO']) == an and int(r['NIVEL'] or 0) == 1 and t(r['LOTE']) == lote
                     and isinstance(r['DATA_HORA'], dt.datetime)], key=lambda r: r['DATA_HORA'])
        return rr[k if len(rr) >= abs(k) else len(rr) // 2], rr
    R_at, reais = real('625311')
    R_fo, _ = real(L_FORA)
    R_so, _ = real(L_SO)
    eq = t(R_at['EQUIPAMENTO'])
    dias = {'625311': R_at['DATA_HORA'].strftime('%d/%m/%Y'), L_FORA: R_fo['DATA_HORA'].strftime('%d/%m/%Y'),
            L_SO: R_so['DATA_HORA'].strftime('%d/%m/%Y')}
    out['dias'] = dias
    qa.ex.xl.EnableEvents = True
    linhas = {}

    def novo(rot, lote, an, res, hora, nivel='1'):
        _, r = C.linha_livre_manual(qa)
        for col, v in (('DATA', dias[lote]), ('HORA', hora), ('EQUIPAMENTO', eq), ('LOTE', "'" + lote),
                       ('NIVEL', nivel), ('ANALITO', an), ('RESULTADO', res), ('MOTIVO', MOTIVO)):
            V.escrever_celula(qa, r, col, v)
        linhas[rot] = r
    novo('A_implausivel_625311', '625311', 'WBC', '-1', '04:11')
    novo('B_533811_WBC_fora_do_cq', L_FORA, 'WBC', '7,77', '04:13')
    novo('C_533811_menos1', L_FORA, 'WBC', '-1', '04:15', nivel='2')
    novo('D_536411_RBC_outro_material', L_SO, 'RBC', '4,44', '04:17')
    novo('E_536411_WBC_ativo', L_SO, 'WBC', '7,77', '04:19')
    novo('F_625311_WBC_controle', '625311', 'WBC', '7,77', '04:21', nivel='2')
    # lista de validacao do LOTE aceita os lotes usados?
    lo_m, _ = V.manual_cols(qa)
    vl = {}
    for rot, r in linhas.items():
        try:
            vl[rot] = bool(lo_m.ListColumns('LOTE').DataBodyRange.Cells(r, 1).Validation.Value)
        except Exception as ex:  # noqa: BLE001
            vl[rot] = 'sem validacao ' + ex.__class__.__name__
    out['validacao_LOTE'] = vl
    # ---- inativar um ponto real e comentar outro (nucleo da janela do Painel)
    out['n_reais_WBC_N1_625311'] = len(reais)
    alvo_inat, alvo_com = reais[-12], reais[-13]   # 1a rodada: -40/-41 estourava (o lote 625311 tem < 41)
    out['inativar'] = {'id': t(alvo_inat['ID_REGISTRO']),
                       'ret': t(qa.run('mIntegracao.InativarPorId', t(alvo_inat['ID_REGISTRO']),
                                       'TESTE-AUDITORIA: inativação de teste (F2 Hema)', True, '', teto=120))[:200]}
    out['comentar'] = {'id': t(alvo_com['ID_REGISTRO']),
                       'ret': t(qa.run('mIntegracao.GravarComentarioTecnico', t(alvo_com['ID_REGISTRO']),
                                       'TESTE-AUDITORIA: comentário de teste (F2 Hema)', 'NÃO', teto=120))[:200]}
    aud.novos()
    sync0 = qa.lo('tblSyncLog').ListRows.Count
    vg = V.Vigia(qa.ex.pid, teto=600)
    vg.start()
    t0 = time.time()
    try:
        r_at = t(qa.run('mIntegracao.AtualizarDadosAutomatico', teto=1500))
    finally:
        vg.parar.set()
    out['atualizar1'] = {'ret': r_at[:700], 's': round(time.time() - t0, 1), 'caixas': vg.vistos[:5],
                         'sync_linhas': [sync0, qa.lo('tblSyncLog').ListRows.Count]}
    tr1 = aud.novos()
    leit = {k: V.ler_manual(qa, r) for k, r in linhas.items()}
    ids = {k: v['ID_REGISTRO'][0] for k, v in leit.items()}
    out['ids'] = ids
    fin2 = qa.ler('tblCQ_Final')
    finid = {t(r['ID_REGISTRO']): r for r in fin2}
    qai = qa.ler('tblQA_Integracao')
    res = {}
    for k, v in leit.items():
        i = ids[k]
        f = finid.get(i, {})
        res[k] = {'STATUS_aba': v['STATUS'][0], 'DETALHE_aba': v['DETALHE'][1][:260], 'RUN_aba': v['RUN'][1],
                  'final': {c: t(f.get(c)) for c in ('STATUS_ANALITICO', 'PARTICIPA_ESTATISTICA', 'ANALITO', 'LOTE',
                                                     'MOTIVO_EXCLUSAO_AUTOMATICA', 'REGISTRAR_RESULTADO_NO_LJ',
                                                     'TIPO_PLOTAGEM_LJ')},
                  'QA': [f"{t(x.get('CODIGO'))}|{t(x.get('SEVERIDADE'))}" for x in qai
                         if any(t(z) == i for z in x.values())],
                  'trilha': [x['acao'] for x in tr1 if x['acao'] and re.search(re.escape(i) + r'\b', x['txt'])]}
    out['resultado'] = res
    out['inativado_no_final'] = {c: t(finid.get(out['inativar']['id'], {}).get(c))
                                 for c in ('STATUS_ANALITICO', 'PARTICIPA_ESTATISTICA', 'DATA_INATIVACAO',
                                           'MOTIVO_INATIVACAO')}
    out['DATA_INATIVACAO_depois'] = cols_final(qa)
    qa_cod = collections.Counter(t(x.get('CODIGO')) for x in qai)
    out['QA_codigos'] = {k: qa_cod.get(k, 0) for k in ('A13', 'A14', 'E10', 'I05')}
    out['oraculo'] = oraculo_regras(fin2)
    out['A13_igual_LOTE'] = qa_cod.get('A13', 0) == sum(out['oraculo']['total_LOTE'].values())
    pr = qa.ws('Principal - Resultados')
    out['Principal_A5'] = t(pr.Range('A5').Text)[:220]
    out['trilha_outros'] = [x['acao'] for x in tr1 if x['acao'] and not any(
        re.search(re.escape(i) + r'\b', x['txt']) for i in ids.values() if i)]
    qa.ex.xl.Calculate()
    out['retrato'] = retrato(qa, 'R1')
    # 2o ATUALIZAR (so entradas)
    t0 = time.time()
    r2 = t(qa.run('mIntegracao.AtualizarSoEntradas', teto=1500))
    tr2 = aud.novos()
    out['atualizar2'] = {'ret': r2[:200], 's': round(time.time() - t0, 1),
                         'trilha_manual_repetida': [x['acao'] for x in tr2 if x['acao'] and 'MANUAL' in x['acao']]}
    L['_ids_manuais'] = ids
    return out, fin2


def s_painel_lotes(qa, fin):
    p = qa.ws('Painel')
    out = {}

    def foto(an, lote):
        time.sleep(1)
        d = {a: t(p.Range(a).Text)[:220] for a in ('C3', 'H3', 'K4', 'H4', 'J7', 'J8', 'J9', 'Q4', 'A39')}
        d['engPainel'] = qa.painel_stats()
        d['indep_participa_por_nivel'] = dict(collections.Counter(
            int(r['NIVEL'] or 0) for r in fin if t(r['LOTE']) == lote and t(r['ANALITO']) == an
            and r['PARTICIPA_ESTATISTICA'] == 'SIM' and num(r['RESULTADO'])))
        d['indep_total_por_nivel'] = dict(collections.Counter(
            int(r['NIVEL'] or 0) for r in fin if t(r['LOTE']) == lote and t(r['ANALITO']) == an))
        return d
    por(qa, p, 'K4', 'Tudo (sem filtro)')
    for an, lote in (('WBC', L_SO), ('RBC', L_SO), ('WBC', L_FORA), ('WBC', '625311')):
        por(qa, p, 'C3', an)
        por(qa, p, 'H3', "'" + lote)
        out[f'{an}_{lote}'] = foto(an, lote)
    return out


# ================================================================================================ calc
def calc(nomes=('E0', 'S0', 'S1', 'S2', 'S3', 'RX', 'S4', 'R1')):
    import calc_independente as CI
    res = {}
    for nome in nomes:
        cam = os.path.join(C.PASTA, f'{PROD}_F2_{nome}.xlsm')
        if not os.path.exists(cam):
            continue
        d, c, meta, linhas = CI.auditar_hema(cam, 1e-9)
        pain = meta.pop('painel_comparacoes', [])
        dif = [{'linha': L_['linha'], 'analito': L_['analito'], 'nivel': L_['nivel'], 'col': k, 'qc': x['qc'],
                'ind': x['ind'], 'dr': x['dif_rel']} for L_ in linhas for k, x in L_['comp'].items()
               if x['status'] != 'OK']
        tot = collections.Counter('OK' if x['status'] == 'OK' else 'DIF' for L_ in linhas for x in L_['comp'].values())
        ptot = collections.Counter('OK' if x['status'] == 'OK' else 'DIF' for x in pain)
        alvo = {f"{L_['analito']}|{L_['nivel']}": {k: x['qc'] for k, x in L_['comp'].items()}
                for L_ in linhas if L_['analito'] == AN}
        ind = {f"{L_['analito']}|{L_['nivel']}": {k: x['ind'] for k, x in L_['comp'].items()}
               for L_ in linhas if L_['analito'] == AN}
        b = c.bias_ceq(AN, provedor=d.f_provedor)
        cvp = {}
        si, sf = c.janela_sigma()
        for nv in (1, 2, 3):
            try:
                cvp[nv] = c.cv_pooled_ext(AN, nv, si, sf)
            except Exception as ex:  # noqa: BLE001
                cvp[nv] = repr(ex)[:80]
        res[nome] = {'meta': {k: v for k, v in meta.items() if k in ('lote', 'equipamento', 'ano_de', 'ano_ate',
                                                                       'provedor', 'ano_ep', 'rodada', 'janela_sigma',
                                                                       'janela_sigma_cache', 'janela_12m', 'n_cq',
                                                                       'n_eqa')},
                     'totais': dict(tot), 'painel_totais': dict(ptot), 'n_div': len(dif), 'divergencias': dif[:40],
                     'painel_div': [x for x in pain if x['status'] != 'OK'][:20],
                     'linhas_HGB_qc': alvo, 'linhas_HGB_ind': ind,
                     'bias_ind': ({k: b[k] for k in ('abs', 'signed', 'n', 'nrod', 'rodadas')} if b else None),
                     'cv_pooled_HGB': cvp}
        json.dump(res[nome], open(os.path.join(C.PASTA, f'calcH_{nome}.json'), 'w', encoding='utf-8'),
                  ensure_ascii=False, indent=1, default=str)
        print('calc', nome, dict(tot), 'painel', dict(ptot), flush=True)
    return res


def main():
    src, copia, h0 = C.copiar(PROD)
    destino = os.path.join(C.PASTA, f'{PROD}_F2.xlsm')
    if os.path.exists(destino):
        os.remove(destino)
    os.replace(copia, destino)
    L = C.Log(os.path.join(C.PASTA, f'{PROD}_F2.json'))
    L['ambiente'] = {'producao': os.path.basename(src), 'sha_prod_12': h0, 'copia': os.path.basename(destino)}
    senha = C.senha_planilha()
    qa = C.qa_final.QA(PROD, destino)
    L['excel_pid'] = qa.ex.pid
    try:
        p = qa.ws('Painel')
        p.Activate()
        qa.ex.esperar()
        fin = qa.ler('tblCQ_Final')
        L['linhas_tblCQ_Final'] = len(fin)
        L['ultimo_recebido'] = str(max((r['DATA_HORA'] for r in fin if isinstance(r['DATA_HORA'], dt.datetime)),
                                       default=None))
        secao(L, 'O49', s_o49, qa, senha)
        criar_usuario(qa, USR, 'ANALISTA')
        L['login'] = fazer_login(qa, USR)
        aud = V.Audit(qa)
        secao(L, 'ceq', s_ceq, qa, aud, senha)
        secao(L, 'O57_sem_sessao', s_sem_sessao, qa, aud)
        print('>> regras_manuais', time.strftime('%H:%M:%S'), flush=True)
        try:
            rm, fin2 = s_regras_manuais(qa, fin, aud, senha, L)
            L['regras_manuais'] = rm
        except Exception as ex:  # noqa: BLE001
            L['regras_manuais'] = {'ERRO': repr(ex), 'trace': traceback.format_exc()[-2500:]}
            fin2 = qa.ler('tblCQ_Final')
            rm = {}
        extra = dict(L.R.get('_ids_manuais', {}) or {})
        if isinstance(rm, dict) and rm.get('inativar'):
            extra['inativado'] = rm['inativar']['id']
            extra['comentado'] = rm['comentar']['id']
        secao(L, 'adr094', B.s_adr094, qa, extra)
        secao(L, 'painel_lotes', s_painel_lotes, qa, fin2)
    except Exception as ex:  # noqa: BLE001
        L['ERRO_FATAL'] = {'erro': repr(ex), 'trace': traceback.format_exc()[-3000:]}
    finally:
        try:
            qa.fechar(salvar=False)
        except Exception:
            pass
    del senha
    h1 = C.sha(src)
    L['producao_intacta'] = {'sha_antes': h0, 'sha_depois': h1, 'igual': h0 == h1}
    print('FIM excel', flush=True)


def parte2():
    """2a copia nova: so regras + manuais + inativar/comentar + ATUALIZAR, ADR-094 e Painel dos lotes (na 1a rodada a
    secao regras_manuais parou num indice antes do ATUALIZAR)."""
    src, copia, h0 = C.copiar(PROD)
    destino = os.path.join(C.PASTA, f'{PROD}_F2_p2.xlsm')
    if os.path.exists(destino):
        os.remove(destino)
    os.replace(copia, destino)
    L = C.Log(os.path.join(C.PASTA, f'{PROD}_F2_p2.json'))
    L['ambiente'] = {'producao': os.path.basename(src), 'sha_prod_12': h0, 'copia': os.path.basename(destino)}
    senha = C.senha_planilha()
    qa = C.qa_final.QA(PROD, destino)
    L['excel_pid'] = qa.ex.pid
    try:
        p = qa.ws('Painel')
        p.Activate()
        qa.ex.esperar()
        fin = qa.ler('tblCQ_Final')
        L['linhas_tblCQ_Final'] = len(fin)
        criar_usuario(qa, USR, 'ANALISTA')
        L['login'] = fazer_login(qa, USR)
        aud = V.Audit(qa)
        print('>> regras_manuais', time.strftime('%H:%M:%S'), flush=True)
        rm, fin2 = s_regras_manuais(qa, fin, aud, senha, L)
        L['regras_manuais'] = rm
        extra = dict(L.R.get('_ids_manuais', {}) or {})
        extra['inativado'] = rm['inativar']['id']
        extra['comentado'] = rm['comentar']['id']
        secao(L, 'adr094', B.s_adr094, qa, extra)
        secao(L, 'painel_lotes', s_painel_lotes, qa, fin2)
    except Exception as ex:  # noqa: BLE001
        L['ERRO_FATAL'] = {'erro': repr(ex), 'trace': traceback.format_exc()[-3000:]}
    finally:
        try:
            qa.fechar(salvar=False)
        except Exception:
            pass
    del senha
    h1 = C.sha(src)
    L['producao_intacta'] = {'sha_antes': h0, 'sha_depois': h1, 'igual': h0 == h1}
    print('FIM excel', flush=True)


if __name__ == '__main__':
    t0 = time.time()
    if '--parte2' in sys.argv:
        parte2()
        r = calc(('R1',))
        json.dump(r, open(os.path.join(C.PASTA, f'{PROD}_F2_calcR.json'), 'w', encoding='utf-8'),
                  ensure_ascii=False, indent=1, default=str)
        print(f'tempo {time.time() - t0:.0f}s', flush=True)
        sys.exit(0)
    if '--so-calc' not in sys.argv:
        main()
    r = calc()
    json.dump(r, open(os.path.join(C.PASTA, f'{PROD}_F2_calc.json'), 'w', encoding='utf-8'),
              ensure_ascii=False, indent=1, default=str)
    print(f'tempo {time.time() - t0:.0f}s', flush=True)
