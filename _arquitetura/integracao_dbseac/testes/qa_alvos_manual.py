# -*- coding: utf-8 -*-
"""QA do ADR-102 -- alvos do laboratorio por lote e periodo (manual) e vigilancia contra a referencia da bula/Unity no
rodape do Painel. Roda numa COPIA instalada (instalar_adr102); nada e salvo.

Uso: python qa_alvos_manual.py <Bioquimica|Hematologia> <copia_instalada.xlsm> [saida.json]

Pelo caminho real: as entradas da secao (B98, E98, G98) e a grade da referencia sao escritas com os eventos LIGADOS
(Worksheet_Change do Painel -> mAlvosRef); os botoes sao chamados pelos seus nucleos (mAlvosRef.AplicarAlvos,
mAlvosRef.RegistrarParecerCore) com as respostas das perguntas dadas por mAlvosRef.DefinirRespostasTeste
(Interactive = False: nenhuma caixa abre). Adendo 25/11: a 1a resposta do "Aplicar" e o modo (1 = Media e DP,
2 = so a Media, 3 = so o DP). Oraculo INDEPENDENTE: N/Media/DP do periodo pelo calcular_alvos do
aplicar_alvos_ciq.py sobre a tblCQ_Final lida aqui (PARTICIPA = SIM, lote e analito exatos, todos os equipamentos,
DATA entre as datas inclusive); fora da faixa contado aqui sobre a mesma tabela.

  A00 pre-condicoes: secao instalada, lote x analito com a chave UMA vez na LotesStore, N1 >= 90 no lote, Painel no
      analito; janelas de datas para N 2..19 e 20..89
  A01 B98/E98/G98 digitados -> a apuracao do Painel (N, Media, DP) = oraculo
  A02 N < 2 em todos os niveis: bloqueia (ERRO), nada gravado
  A03 lote vazio / inexistente / chave duplicada: abortam com mensagem clara, nada gravado
  A04 cancelar a confirmacao nao grava nada (nem trilha)
  A04b cancelar a pergunta do que transferir (sem resposta / resposta invalida) nao grava nada
  A05 N < 20 sem a ciencia de TEMPORARIO: nao grava
  A06 N < 20 com a ciencia: grava; origem MANUAL ... TEMPORARIO; AA:AO do nivel; trilha PARAMETRO_LOTE_CALCULADO
  A07 N 20..89: grava, ACEITAVEL, aviso do ideal >= 90
  A08 N >= 90: grava, IDEAL; Media/DP na LotesStore = oraculo; a VBA (mLotes.MediaDoLote/DPDoLote) le o novo alvo
  A09 cabecalhos AA1:AO1 e auditoria do nivel (N, datas, condicao, atualizado em)
  M01 modo 2 (so a Media): so a Media muda (= oraculo), o DP fica; origem [MEDIA]; trilha so da Media
  M02 modo 3 (so o DP) com N baixo: so o DP muda (= oraculo), a Media fica; origem [DP] ... TEMPORARIO
  M03 origem temporaria: coluna Origem atual (I100) amarela com o aviso "temporario: substituir ..."
  M04 excecao -> modo 2 -> modo 2: a marca do DP (CONDICAO_Nk por campo) e a heranca em W continuam; amarelo fica
  M05 W reescrita por outra gravacao (bloco da aba Analitos): o amarelo continua pela marca do CONDICAO_Nk
  A10 aplicar_alvos_ciq.selecionar_pares pula a linha com origem MANUAL (e a conta)
  E01 excecao: gravar sem liberar e recusado; o bloco abre TRANCADO (Locked, cinza)
  E02 liberar: cancelar nao destrava; confirmar destrava, marca Z105 e grava ALVO_MANUAL_LIBERADO
  E03 Lote anterior digitado puxa o DP da LotesStore desse lote (mesmo analito e nivel); sem justificativa: recusado
  E04 gravar: cancelar nao grava; confirmar grava so os campos preenchidos, origem EXCECAO TEMPORARIA ... media=BULA;
      DP=LOTE ANTERIOR <lote>, CONDICAO_N1 = EXCECAO_TEMPORARIA, Audit_Log por valor; outros niveis intactos
  E05 depois de gravar: campos limpos e TRANCADOS de novo; gravar de novo e recusado
  E06 origem EXCECAO TEMPORARIA: Origem atual amarela com o aviso
  E07 liberar e trocar o lote da secao tranca de novo
  E08 "Aplicar Alvos ao Lote" (Media e DP) substitui a excecao: origem MANUAL sem marca temporaria, CONDICAO IDEAL,
      amarelo sai
  E09 excecao so do DP sobre valor PROPRIO: pede a 2a confirmacao (cancelar nao grava); confirmada, a Media e o
      N/periodo ficam, CONDICAO_N1 = IDEAL+EXCECAO_TEMPORARIA_DP, Audit_Log com BULA, justificativa e PROPRIO
  E10 o proprio lote como "Lote anterior" e recusado (nada gravado)
  R01 referencia: resultados abaixo do LI e acima do LS viram PENDENTE (= oraculo); dentro da faixa nao
  R02 verificar de novo: nenhuma observacao nova (alerta so para nova)
  R03 corrigir o limite: as pendentes antigas viram RETIRADA; as que continuam fora entram com a chave nova
  R04 parecer curto recusado; parecer registrado da ciencia a todas as pendentes do lote (RefBulaParecer, CIENTE)
  R05 CIENTE nunca e apagada (limite corrigido de novo)
  R07 a referencia (grade, verificacao, parecer) NUNCA grava na LotesStore (A1:AO inteira igual antes e depois)
  R06 LS gravado direto na RefBulaStore -> ATUALIZAR (so entradas) gera as PENDENTES (= oraculo), linha no resumo, e
      o resumo ainda termina no retorno da foto do Westgard
  Y00 arquivo de entrada intacto
"""
import collections
import datetime as dt
import json
import os
import re
import sys
import traceback

AQUI = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, AQUI)
sys.path.insert(0, os.path.join(AQUI, '..'))
sys.path.insert(0, os.path.join(AQUI, '..', '..', 'agentes', 'ferramentas'))
from qa_final import QA, reg, RES, sha256_arquivo  # noqa: E402
from instalar_integracao import SENHA  # noqa: E402
import aplicar_alvos_ciq as oraculo  # noqa: E402

R_ENT, R_NV1, R_REF1, N_REF, R_PAR_TXT = 98, 100, 115, 40, 176        # adendo 25/11: referencia desceu 8 linhas
R_MAN_TIT, R_MAN1, R_MAN_JUST, C_MARCA = 105, 107, 110, 'Z'               # excecao temporaria (valores manuais)
AMARELO_TEMP = 156 * 65536 + 235 * 256 + 255                              # RGB(255, 235, 156) do mAlvosRef
FUTURO = (dt.datetime(2099, 1, 1), dt.datetime(2099, 1, 2))


class PreCondicao(Exception):
    pass


def txt(v):
    return oraculo.txt(v)


def U(v):
    return txt(v).upper()


def num(v):
    return oraculo.num(v)


def perto(a, b, tol=1e-9):
    return num(a) and num(b) and abs(float(a) - float(b)) <= tol * max(1.0, abs(float(b)))


def vnum(v):
    # valor do Audit_Log (numero ou texto com virgula/ponto) -> float; None se nao for numero
    if num(v):
        return float(v)
    try:
        return float(txt(v).replace(',', '.'))
    except ValueError:
        return None


def dia(v):
    return v.date() if isinstance(v, dt.datetime) else None


# Marca temporaria POR CAMPO no CONDICAO_Nk (especificacao do adendo, conferida contra o que a VBA grava):
# EXCECAO_TEMPORARIA / TEMPORARIO = os dois campos; senao tokens "+": condicao propria, <marca>_MEDIA, <marca>_DP.
EXC = 'EXCECAO_TEMPORARIA'


def marcas(cond):
    c = U(cond)
    if not c:
        return '', ''
    if c == EXC:
        return 'EXCECAO', 'EXCECAO'
    if c == 'TEMPORARIO':
        return 'TEMPORARIO', 'TEMPORARIO'
    tm = td = ''
    for t in (x.strip() for x in c.split('+')):
        if t == EXC + '_MEDIA':
            tm = 'EXCECAO'
        elif t == EXC + '_DP':
            td = 'EXCECAO'
        elif t == 'TEMPORARIO_MEDIA':
            tm = 'TEMPORARIO'
        elif t == 'TEMPORARIO_DP':
            td = 'TEMPORARIO'
    return tm, td


def propria(cond):
    return next((t.strip() for t in U(cond).split('+') if t.strip() in ('ACEITAVEL', 'IDEAL')), '')


def cond_das_marcas(tm, td, pr):
    def mc(t):
        return EXC if t == 'EXCECAO' else t
    if tm == td:
        return mc(tm) if tm else pr
    partes = [pr] if pr and (not tm or not td) else []
    if tm:
        partes.append(mc(tm) + '_MEDIA')
    if td:
        partes.append(mc(td) + '_DP')
    return '+'.join(partes)


def cond_apos_aplicar(ant, modo, c):
    tm, td = marcas(ant)
    h = 'TEMPORARIO' if c == 'TEMPORARIO' else ''
    if modo in (1, 2):
        tm = h
    if modo in (1, 3):
        td = h
    return cond_das_marcas(tm, td, propria(ant) if h else c)


def executar(produto, caminho, saida):
    h0 = sha256_arquivo(caminho)
    q = QA(produto, caminho)
    try:
        xl, nlv = q.ex.xl, q.nlv
        p = q.ws('Painel')
        ls = q.ws('LotesStore')

        def run(m, *a, teto=600):
            return q.run('mAlvosRef.' + m if '.' not in m else m, *a, teto=teto)

        def resp(s):
            run('DefinirRespostasTeste', s)

        def aud0():
            return int(q.run('mAuditoria.UltimaLinhaAudit', teto=60))

        def trilha(r0, acao):
            al = q.ws('Audit_Log')
            r1 = aud0()
            if r1 <= r0:
                return []
            return [l for l in al.Range(al.Cells(r0 + 1, 1), al.Cells(r1, 25)).Value
                    if txt(l[6]) == acao and txt(l[7]) == 'mAlvosRef']

        def sem_eventos(f):
            xl.EnableEvents = False
            try:
                f()
            finally:
                xl.EnableEvents = True

        def escrever_secao(lote, ini, fim, eventos=False):
            def f():
                p.Range(f'B{R_ENT}').Value = lote
                p.Range(f'E{R_ENT}').Value = ini if ini else None
                p.Range(f'G{R_ENT}').Value = fim if fim else None
            if eventos:
                f()
            else:
                sem_eventos(f)

        # ------------------------------------------------------------ A00 pre-condicoes
        if not run('SecaoInstalada', teto=60):
            raise PreCondicao('secao do ADR-102 nao instalada no Painel (instalar_adr102)', {})
        lim = txt(run('LimitesNTexto', teto=60)).split('|')
        rec, ideal = int(lim[0]), int(lim[1])
        rows = q.ler('tblCQ_Final')
        sim = [r for r in rows if U(r.get('PARTICIPA_ESTATISTICA')) == 'SIM' and num(r.get('RESULTADO'))]
        catalogo = [txt(v[0]) for v in q.espec().Range('A4:A43').Value if txt(v[0])]
        ult = ls.Cells(ls.Rows.Count, 16).End(-4162).Row
        chaves = collections.Counter(U(v[0]) for v in ls.Range(f'P2:P{ult}').Value if txt(v[0]))
        cont = collections.Counter((U(r['LOTE']), U(r['ANALITO']), int(float(r['NIVEL']))) for r in sim
                                   if txt(r.get('NIVEL')))
        cand = []
        for (lt, an, nv), n in cont.items():
            if nv != 1 or n < ideal + 5 or chaves.get(f'{lt}|{an}') != 1:
                continue
            if an not in [U(x) for x in catalogo]:
                continue
            cand.append((n, lt, an))
        ideal_ajustado = None
        if not cand:
            # Dados sem lote com N1 >= ideal + 5 (ex.: Hematologia, lotes curtos): o limite e configuravel
            # (Cfg_Integracao ALVO_N_IDEAL), entao a COPIA de teste usa ideal = N1 do maior lote - 5, o que tambem
            # prova que a regra le a configuracao. Exige N1 >= rec + 15 para sobrar janela ACEITAVEL.
            for (lt, an, nv), n in cont.items():
                if (nv == 1 and n >= rec + 15 and chaves.get(f'{lt}|{an}') == 1
                        and an in [U(x) for x in catalogo]):
                    cand.append((n, lt, an))
            if not cand:
                raise PreCondicao(f'nenhum lote x analito com N1 >= {rec + 15} ativos e chave unica na LotesStore', {})
            novo = max(cand)[0] - 5
            wsc = q.ws('Cfg_Integracao')
            wsc.Unprotect(SENHA)
            lo = wsc.ListObjects('tblConfigIntegracao')
            corpo = lo.DataBodyRange
            achou = False
            for i in range(1, corpo.Rows.Count + 1):
                if U(corpo.Cells(i, 1).Value) == 'ALVO_N_IDEAL':
                    sem_eventos(lambda: corpo.Cells(i, 2).__setattr__('Value', novo))
                    achou = True
                    break
            lim = txt(run('LimitesNTexto', teto=60)).split('|')
            if not achou or int(lim[1]) != novo:
                raise PreCondicao('ALVO_N_IDEAL da copia nao pode ser ajustado', {'lido': lim, 'pedido': novo})
            rec, ideal = int(lim[0]), int(lim[1])
            ideal_ajustado = novo
        cand.sort(reverse=True)
        _, LT, AN = cand[0]
        lote = next(txt(r['LOTE']) for r in sim if U(r['LOTE']) == LT)
        analito = next(x for x in catalogo if U(x) == AN)
        if U(q.nome('selAnalito').Value) != AN:
            pos = [U(x) for x in [txt(v[0]) for v in q.espec().Range('A4:A43').Value]].index(AN) + 1
            p.Unprotect(SENHA)
            sem_eventos(lambda: p.Range('B3').__setattr__('Value', pos))
            q.run('mEstatistica.PainelMudou', teto=900)
        if U(q.nome('selAnalito').Value) != AN:
            raise PreCondicao('o Painel nao passou a mostrar o analito escolhido', {'analito': analito})
        rs_lote = [r for r in sim if U(r['LOTE']) == LT and U(r['ANALITO']) == AN]

        def oraculo_periodo(ini, fim):
            sel = [r for r in rs_lote if (ini is None or (dia(r.get('DATA_HORA')) and dia(r['DATA_HORA']) >= ini.date()))
                   and (fim is None or (dia(r.get('DATA_HORA')) and dia(r['DATA_HORA']) <= fim.date()))]
            c = oraculo.calcular_alvos(sel, nlv, n=10 ** 9, minimo=2)
            out = {}
            for nv in range(1, nlv + 1):
                x = c.get((LT, AN, nv))
                out[nv] = (x['n_lote'], x['media_periodo'], x['dp_periodo']) if x else (0, None, None)
            return out

        datas1 = sorted(dia(r['DATA_HORA']) for r in rs_lote if int(float(r['NIVEL'])) == 1 and dia(r.get('DATA_HORA')))
        d_ini = dt.datetime.combine(datas1[0], dt.time())

        def janela(lo, hi):
            for d in sorted(set(datas1)):
                n = sum(1 for x in datas1 if x <= d)
                if lo <= n <= hi:
                    return dt.datetime.combine(d, dt.time())
            return None

        fim_temp, fim_aceit = janela(2, rec - 1), janela(rec, ideal - 1)
        reg('A00 Pre-condicoes: secao instalada, lote x analito com chave unica e N1 >= ideal, janelas de N',
            bool(fim_temp and fim_aceit),
            {'lote': lote, 'analito': analito, 'N1_total': len(datas1), 'limites': [rec, ideal],
             'janela_temporario': str(fim_temp), 'janela_aceitavel': str(fim_aceit),
             'ideal_ajustado_na_copia': ideal_ajustado})
        if not (fim_temp and fim_aceit):
            return
        rLs = [i for i, v in enumerate(ls.Range(f'P2:P{ult}').Value, 2) if U(v[0]) == f'{LT}|{AN}'][0]

        def linha_ls():
            return tuple(ls.Range(f'A{rLs}:AO{rLs}').Value[0])

        # ------------------------------------------------------------ A01 hook da secao
        escrever_secao(lote, d_ini, fim_aceit, eventos=True)
        ora = oraculo_periodo(d_ini, fim_aceit)
        vista = {nv: tuple(p.Range(f'B{R_NV1 + nv - 1}:D{R_NV1 + nv - 1}').Value[0]) for nv in range(1, nlv + 1)}
        ok1 = all(int(vista[nv][0] or 0) == ora[nv][0] and (ora[nv][0] < 1 or perto(vista[nv][1], ora[nv][1]))
                  and (ora[nv][0] < 2 or perto(vista[nv][2], ora[nv][2])) for nv in vista)
        reg('A01 Lote/De/Ate digitados (eventos ligados): apuracao do Painel (N, Media, DP) = oraculo; Z98 = analito',
            ok1 and U(p.Range(f'Z{R_ENT}').Value) == AN, {'vista': vista, 'oraculo': ora,
                                                           'Z98': txt(p.Range(f'Z{R_ENT}').Value)})

        # ------------------------------------------------------------ A02 N < 2
        antes = linha_ls()
        escrever_secao(lote, *FUTURO)
        resp('1|SIM|SIM')
        r = txt(run('AplicarAlvos'))
        resp('')
        reg('A02 N < 2 em todos os niveis: bloqueia com mensagem, nada gravado',
            r.startswith('ERRO|') and 'DP' in r and linha_ls() == antes, {'retorno': r[:200]})

        # ------------------------------------------------------------ A03 lote vazio / inexistente / duplicado
        res3 = {}
        escrever_secao('', d_ini, fim_aceit)
        res3['vazio'] = txt(run('AplicarAlvos'))
        escrever_secao('TODOS', d_ini, fim_aceit)
        res3['todos'] = txt(run('AplicarAlvos'))
        escrever_secao('QA_LOTE_INEXISTENTE', d_ini, fim_aceit)
        res3['inexistente'] = txt(run('AplicarAlvos'))
        ls.Unprotect(SENHA)
        dup = ls.Cells(ls.Rows.Count, 16).End(-4162).Row + 1
        ls.Cells(dup, 16).Value = f'{LT}|{AN}'
        q.run('mLotes.InvalidarLotes', teto=60)
        escrever_secao(lote, d_ini, fim_aceit)
        resp('1|SIM|SIM')
        res3['duplicado'] = txt(run('AplicarAlvos'))
        resp('')
        ls.Cells(dup, 16).ClearContents()
        ls.Protect(SENHA, False, True, True, True)
        q.run('mLotes.InvalidarLotes', teto=60)
        reg('A03 Lote vazio, TODOS, inexistente e chave duplicada abortam (ERRO com o motivo); nada gravado',
            res3['vazio'].startswith('ERRO|') and 'B98' in res3['vazio']
            and res3['todos'].startswith('ERRO|') and 'B98' in res3['todos']
            and res3['inexistente'].startswith('ERRO|') and 'registro' in res3['inexistente']
            and res3['duplicado'].startswith('ERRO|') and 'duplicada' in res3['duplicado'] and linha_ls() == antes,
            {k: v[:160] for k, v in res3.items()})

        # ------------------------------------------------------------ A04 cancelar
        a0 = aud0()
        escrever_secao(lote, d_ini, fim_aceit)
        resp('1|NAO')
        r = txt(run('AplicarAlvos'))
        resp('')
        reg('A04 Cancelar a confirmacao nao grava (LotesStore e trilha iguais)',
            r == 'CANCELADO|confirmacao' and linha_ls() == antes and not trilha(a0, 'PARAMETRO_LOTE_CALCULADO'),
            {'retorno': r, 'pergunta': txt(run('UltimaPergunta'))[:300]})

        # ------------------------------------------------------------ A04b cancelar a pergunta do que transferir
        a0 = aud0()
        resp('')                                      # sem resposta e Interactive = False: Cancelar
        r_sem = txt(run('AplicarAlvos'))
        perg_modo = txt(run('UltimaPergunta'))
        resp('9|SIM|SIM')                             # resposta fora de 1/2/3: Cancelar
        r_inv = txt(run('AplicarAlvos'))
        resp('')
        reg('A04b Cancelar a pergunta do que transferir (Cancelar ou resposta fora de 1/2/3) nao grava nada',
            r_sem == 'CANCELADO|transferencia' and r_inv == 'CANCELADO|transferencia' and linha_ls() == antes
            and not trilha(a0, 'PARAMETRO_LOTE_CALCULADO') and all(x in perg_modo for x in ('1 =', '2 =', '3 =')),
            {'sem_resposta': r_sem, 'invalida': r_inv, 'pergunta': perg_modo[:300]})

        # ------------------------------------------------------------ A05 / A06 temporario
        escrever_secao(lote, d_ini, fim_temp)
        resp('1|SIM|NAO')
        r = txt(run('AplicarAlvos'))
        resp('')
        reg('A05 N < recomendado sem a ciencia de TEMPORARIO: nao grava',
            r == 'CANCELADO|temporario' and linha_ls() == antes and 'TEMPOR' in U(run('UltimaPergunta')),
            {'retorno': r, 'pergunta': txt(run('UltimaPergunta'))[:300]})

        def cond_py(n, s):
            # oraculo independente da regra do N (ADR-102), mais o DP = 0 mantido
            if n < 2:
                return 'BLOQUEADO'
            if not (num(s) and float(s) > 0):
                return 'DP_ZERO'
            return 'TEMPORARIO' if n < rec else 'ACEITAVEL' if n < ideal else 'IDEAL'

        def cols_nivel(lin, nv):
            base = 26 + 5 * (nv - 1)
            return (lin[2 + 2 * (nv - 1)], lin[3 + 2 * (nv - 1)], lin[23 + nv - 1]) + tuple(lin[base:base + 5])

        def conferir(ini, fim, cond_n1, r, a0, antes_ls):
            ora = oraculo_periodo(ini, fim)
            lin = linha_ls()
            div = []
            for nv in range(1, nlv + 1):
                n, m, s = ora[nv]
                mv, sv = lin[2 + 2 * (nv - 1)], lin[3 + 2 * (nv - 1)]
                base = 26 + 5 * (nv - 1)
                c = cond_py(n, s)
                if f'N{nv}={c}' not in r:
                    div.append((nv, 'condicao no retorno', c))
                if c in ('BLOQUEADO', 'DP_ZERO'):
                    if cols_nivel(lin, nv) != cols_nivel(antes_ls, nv):
                        div.append((nv, 'nivel mantido foi alterado', cols_nivel(lin, nv)))
                else:
                    if not (perto(mv, m) and perto(sv, s)):
                        div.append((nv, 'valor', mv, sv, m, s))
                    if int(lin[base] or 0) != n or txt(lin[base + 3]) != c:
                        div.append((nv, 'auditoria', lin[base:base + 5]))
                    if dia(lin[base + 1]) != ini.date() or dia(lin[base + 2]) != fim.date():
                        div.append((nv, 'datas', lin[base + 1], lin[base + 2]))
                    if txt(lin[23 + nv - 1]):
                        div.append((nv, 'bit TESTE', lin[23 + nv - 1]))
                    if num(q.run('mLotes.MediaDoLote', lote, analito, nv, teto=60)) and not perto(
                            q.run('mLotes.MediaDoLote', lote, analito, nv, teto=60), m):
                        div.append((nv, 'VBA le outro alvo'))
            w = txt(lin[22])
            n_alt = int(r.split('|')[2]) if r.startswith('OK|') else -1
            tr = trilha(a0, 'PARAMETRO_LOTE_CALCULADO')
            return {'ok': r.startswith('OK|') and f'N1={cond_n1}' in r and w.upper().startswith('MANUAL')
                    and '[MEDIA+DP]' in w and r.endswith('|MEDIA+DP') and not div and len(tr) == n_alt,
                    'retorno': r, 'W': w[:200], 'divergencias': div[:6], 'trilha': [len(tr), n_alt], 'oraculo': ora}

        a0 = aud0()
        antes_ls = linha_ls()
        resp('1|SIM|SIM')
        r = txt(run('AplicarAlvos'))
        resp('')
        ev = conferir(d_ini, fim_temp, 'TEMPORARIO', r, a0, antes_ls)
        reg('A06 N < recomendado com a ciencia: grava = oraculo; origem MANUAL ... TEMPORARIO; AA:AO; trilha por valor',
            ev['ok'] and 'TEMPORARIO' in ev['W'], ev)

        def ciencia_ok(perg, ora_):
            # a ciencia de TEMPORARIO so aparece se algum nivel cair em TEMPORARIO pelo N (nunca pelo N1 aqui)
            temp_ = [nv for nv in ora_ if cond_py(*ora_[nv][::2]) == 'TEMPORARIO']
            return ('DECLARO' in U(perg)) == bool(temp_) and 'N1 (N =' not in U(perg)

        # ------------------------------------------------------------ A07 aceitavel
        escrever_secao(lote, d_ini, fim_aceit)
        a0 = aud0()
        antes_ls = linha_ls()
        resp('1|SIM|SIM')     # modo + confirmacao (+ ciencia, so se o N2/N3 cair em TEMPORARIO; a que sobra e limpa)
        r = txt(run('AplicarAlvos'))
        perg = txt(run('UltimaPergunta'))
        resp('')
        ev = conferir(d_ini, fim_aceit, 'ACEITAVEL', r, a0, antes_ls)
        aviso = txt(run('UltimoAviso'))
        reg(f'A07 N {rec}..{ideal - 1}: grava (ACEITAVEL) sem pedir ciencia de TEMPORARIO para o N1; informa o ideal '
            f'>= {ideal}', ev['ok'] and str(ideal) in aviso and ciencia_ok(perg, ev['oraculo']),
            dict(ev, aviso=aviso[:200], pergunta=perg[:200]))

        # ------------------------------------------------------------ A08 ideal
        fim_all = dt.datetime.combine(max(datas1), dt.time())
        escrever_secao(lote, d_ini, fim_all)
        a0 = aud0()
        antes_ls = linha_ls()
        resp('1|SIM|SIM')
        r = txt(run('AplicarAlvos'))
        perg = txt(run('UltimaPergunta'))
        resp('')
        ev = conferir(d_ini, fim_all, 'IDEAL', r, a0, antes_ls)
        reg(f'A08 N >= {ideal}: grava (IDEAL) sem ciencia de TEMPORARIO para o N1; LotesStore e leitura da VBA = '
            'oraculo', ev['ok'] and ciencia_ok(perg, ev['oraculo']), dict(ev, pergunta=perg[:200]))

        # ------------------------------------------------------------ M01 / M02 o que transferir (adendo)
        def so_um_campo(modo, ini, fim, resposta, cond_n1):
            escrever_secao(lote, ini, fim)
            a0_ = aud0()
            antes_ = linha_ls()
            resp(resposta)
            r_ = txt(run('AplicarAlvos'))
            resp('')
            ora_ = oraculo_periodo(ini, fim)
            lin = linha_ls()
            div = []
            for nv in range(1, nlv + 1):
                n, m, s = ora_[nv]
                mv, sv = lin[2 + 2 * (nv - 1)], lin[3 + 2 * (nv - 1)]
                m0, s0 = antes_[2 + 2 * (nv - 1)], antes_[3 + 2 * (nv - 1)]
                if cond_py(n, s) in ('BLOQUEADO', 'DP_ZERO'):
                    if cols_nivel(lin, nv) != cols_nivel(antes_, nv):
                        div.append((nv, 'nivel mantido foi alterado'))
                    continue
                if modo == 2 and not (perto(mv, m) and sv == s0):
                    div.append((nv, 'modo 2: Media = oraculo e DP igual', mv, m, sv, s0))
                if modo == 3 and not (perto(sv, s) and mv == m0):
                    div.append((nv, 'modo 3: DP = oraculo e Media igual', sv, s, mv, m0))
                esp_c = cond_apos_aplicar(antes_[26 + 5 * (nv - 1) + 3], modo, cond_py(n, s))
                if U(lin[26 + 5 * (nv - 1) + 3]) != esp_c:
                    div.append((nv, 'CONDICAO_N (marca por campo)', lin[26 + 5 * (nv - 1) + 3], esp_c))
            w = txt(lin[22])
            tr = trilha(a0_, 'PARAMETRO_LOTE_CALCULADO')
            n_alt = int(r_.split('|')[2]) if r_.startswith('OK|') else -1
            rot = {2: 'MEDIA', 3: 'DP'}[modo]
            campo, outro_campo = ('Media', 'DP') if modo == 2 else ('DP', 'Media')
            cand = [x for x in tr if any(re.fullmatch(campo + r' N\d', txt(c)) for c in x)]
            outro = [x for x in tr if any(re.fullmatch(outro_campo + r' N\d', txt(c)) for c in x)]
            return {'ok': r_.startswith('OK|') and r_.endswith('|' + rot) and f'N1={cond_n1}' in r_
                    and w.upper().startswith('MANUAL') and f'[{rot}]' in w and not div and len(tr) == n_alt
                    and len(cand) == len(tr) and not outro,
                    'retorno': r_, 'W': w[:220], 'divergencias': div[:6], 'trilha': [len(tr), n_alt], 'oraculo': ora_}

        ev = so_um_campo(2, d_ini, fim_aceit, '2|SIM|SIM', 'ACEITAVEL')
        reg('M01 Transferir so a Media (2): so a Media muda (= oraculo), o DP fica; CONDICAO_Nk com a marca por campo; origem '
            '[MEDIA]; trilha so da Media', ev['ok'], ev)
        ev = so_um_campo(3, d_ini, fim_temp, '3|SIM|SIM', 'TEMPORARIO')
        reg('M02 Transferir so o DP (3) com N baixo: so o DP muda (= oraculo), a Media fica; origem [DP] com '
            'TEMPORARIO; trilha so do DP', ev['ok'] and 'TEMPORARIO' in ev['W'], ev)
        cor = int(p.Range(f'I{R_NV1}').Interior.Color or 0)
        reg('M03 Origem temporaria: Origem atual da apuracao (I100) amarela com o aviso "temporario: substituir pelos '
            'valores do laboratorio"', cor == AMARELO_TEMP and 'SUBSTITUIR PELOS VALORES DO LABORAT' in U(
                p.Range(f'I{R_NV1}').Value), {'cor': cor, 'I100': txt(p.Range(f'I{R_NV1}').Value)[:200]})

        # ------------------------------------------------------------ A09 cabecalhos AA:AO
        cab = [txt(c) for c in ls.Range('AA1:AO1').Value[0]]
        esperado = [f'{n}{k}' for k in (1, 2, 3) for n in ('N_USADO_N', 'DATA_INI_N', 'DATA_FIM_N', 'CONDICAO_N',
                                                           'ATUALIZADO_EM_N')]
        quando = linha_ls()[30]
        reg('A09 LotesStore AA1:AO1 = auditoria por nivel; ATUALIZADO_EM do N1 = agora',
            cab == esperado and isinstance(quando, dt.datetime)
            and abs((quando.replace(tzinfo=None) - dt.datetime.now()).total_seconds()) < 3600,
            {'cabecalho': cab, 'atualizado_em': str(quando)})

        # ------------------------------------------------------------ A10 ferramenta pula MANUAL
        st = {}
        for i, linha in enumerate(ls.Range(f'A2:Z{ult}').Value, 2):
            if txt(linha[15]):
                st.setdefault(oraculo.chave(linha[15]), (i, tuple(linha)))
        _, pares, manuais = oraculo.selecionar_pares([LT], [AN], st, {(LT, AN, 1): {}}, nlv)
        reg('A10 aplicar_alvos_ciq pula a linha com origem MANUAL (contada) e nao a recalcula',
            manuais == [f'{LT}|{AN}'] and not pares, {'manuais': manuais, 'pares': pares})

        # ------------------------------------------------------------ E01..E08 excecao temporaria (adendo)
        def estado():
            return txt(run('EstadoManualTexto'))

        def trancado():
            return (bool(p.Range(f'B{R_MAN1}').Locked) and bool(p.Range(f'B{R_MAN_JUST}').Locked)
                    and bool(p.Range(f'E{R_MAN1}').Locked) and not txt(p.Range(f'{C_MARCA}{R_MAN_TIT}').Value)
                    and estado() == 'TRANCADA')

        escrever_secao(lote, d_ini, fim_all)
        run('BloquearManual')
        antes = linha_ls()
        r_e1 = txt(run('GravarManualCore'))
        reg('E01 Excecao: o bloco esta TRANCADO (Locked, cinza) e gravar sem liberar e recusado; nada gravado',
            trancado() and r_e1.startswith('ERRO|') and 'TRANCADA' in U(r_e1) and linha_ls() == antes
            and int(p.Range(f'B{R_MAN1}').Interior.Color or 0) == 217 * 65536 + 217 * 256 + 217,
            {'retorno': r_e1[:200], 'estado': estado()})

        a0 = aud0()
        resp('NAO')
        r_canc = txt(run('LiberarManualCore'))
        ainda = trancado()
        resp('SIM')
        r_lib = txt(run('LiberarManualCore'))
        resp('')
        perg_lib = txt(run('UltimaPergunta'))
        lib = trilha(a0, 'ALVO_MANUAL_LIBERADO')
        reg('E02 Liberar: cancelar nao destrava; confirmar (texto da excecao) destrava, marca Z105 e grava '
            'ALVO_MANUAL_LIBERADO no Audit_Log',
            r_canc == 'CANCELADO|confirmacao' and ainda and r_lib == f'OK|{LT}|{AN}'
            and not p.Range(f'B{R_MAN1}').Locked and not p.Range(f'B{R_MAN_JUST}').Locked
            and U(p.Range(f'{C_MARCA}{R_MAN_TIT}').Value).startswith(f'LIBERADO|{LT}|{AN}|') and len(lib) == 1
            and 'PRÓPRIOS' in U(perg_lib) and 'EXCEÇÃO' in U(perg_lib),
            {'cancelado': r_canc, 'liberado': r_lib, 'estado': estado(), 'trilha': len(lib)})

        # lote anterior: outro lote do mesmo analito com DP N1 > 0 na LotesStore (na falta, o proprio lote)
        ult = ls.Cells(ls.Rows.Count, 16).End(-4162).Row
        la, dp_ant = None, None
        for linha in ls.Range(f'A2:Z{ult}').Value:
            k = U(linha[15])
            if k.endswith(f'|{AN}') and k != f'{LT}|{AN}' and num(linha[3]) and float(linha[3]) > 0 and txt(linha[0]):
                la, dp_ant = txt(linha[0]), float(linha[3])
                break
        if la is None:
            la, dp_ant = lote, float(linha_ls()[3])
        media_bula = round(float(linha_ls()[2]) * 1.013 + 0.01, 4)
        antes = linha_ls()
        p.Range(f'B{R_MAN1}').Value = media_bula
        p.Range(f'E{R_MAN1}').Value = la                                     # eventos ligados: ManualEditado
        dp_tela, orig_tela = p.Range(f'C{R_MAN1}').Value, txt(p.Range(f'D{R_MAN1}').Value)
        r_sem = txt(run('GravarManualCore'))
        reg('E03 Lote anterior digitado: DP = o da LotesStore desse lote (mesmo analito e N1), origem "Lote anterior"; '
            'gravar sem justificativa e recusado (nada gravado, continua liberado)',
            perto(dp_tela, dp_ant) and orig_tela == 'Lote anterior' and r_sem.startswith('ERRO|')
            and 'JUSTIFICATIVA' in U(r_sem) and linha_ls() == antes and estado().startswith('LIBERADA'),
            {'lote_anterior': la, 'dp_store': dp_ant, 'dp_tela': dp_tela, 'origem': orig_tela,
             'retorno': r_sem[:200]})

        just_e4 = 'QA: inicio do lote, sem resultados proprios; media da bula e DP do lote anterior.'
        p.Range(f'B{R_MAN_JUST}').Value = just_e4
        # sem outro lote do analito na LotesStore: o proprio lote e recusado como "Lote anterior" (E10) -> DP da bula
        fonte_e4 = f'LOTE ANTERIOR {U(la)}'
        if U(la) == LT:
            p.Range(f'E{R_MAN1}').Value = None
            p.Range(f'D{R_MAN1}').Value = 'Bula'
            fonte_e4 = 'BULA'
        resp('NAO')
        r_canc = txt(run('GravarManualCore'))
        canc_ok = r_canc == 'CANCELADO|confirmacao' and linha_ls() == antes
        a0 = aud0()
        resp('SIM|SIM')        # confirmacao + (se a Media/DP atual for PROPRIA) a 2a confirmacao; a que sobra e limpa
        r_e4 = txt(run('GravarManualCore'))
        resp('')
        lin = linha_ls()
        w = txt(lin[22])
        tr = trilha(a0, 'PARAMETRO_LOTE_EXCECAO')
        outros = all(cols_nivel(lin, nv) == cols_nivel(antes, nv) for nv in range(2, nlv + 1))
        # conteudo da trilha: justificativa, fonte e antes -> depois de cada valor
        tr_txt = [' | '.join(txt(c) for c in l) for l in tr]
        tr_med = [l for l in tr if any(txt(c) == 'Media N1' for c in l)]
        tr_dp = [l for l in tr if any(txt(c) == 'DP N1' for c in l)]
        trilha_ok = (len(tr_med) == 1 and len(tr_dp) == 1
                     and all(just_e4 in t and '(EXCECAO TEMPORARIA, ' in t for t in tr_txt)
                     and any(perto(vnum(c), media_bula, 1e-6) for c in tr_med[0])
                     and any(perto(vnum(c), antes[2], 1e-6) for c in tr_med[0])
                     and any(perto(vnum(c), dp_ant, 1e-6) for c in tr_dp[0])
                     and 'MEDIA N1 = BULA' in U(' '.join(txt(c) for c in tr_med[0]))
                     and f'DP N1 = {fonte_e4}' in U(' '.join(txt(c) for c in tr_dp[0])))
        reg('E04 Gravar valores manuais: cancelar nao grava; confirmar grava Media (bula) e DP (lote anterior), origem '
            'EXCECAO TEMPORARIA ... N1 media=BULA; DP=LOTE ANTERIOR <lote>, CONDICAO_N1 = EXCECAO_TEMPORARIA, '
            'Audit_Log por valor (justificativa, fonte, antes -> depois); os outros niveis intactos',
            canc_ok and r_e4.startswith('OK|2|') and perto(lin[2], media_bula) and perto(lin[3], dp_ant)
            and U(w).startswith('EXCECAO TEMPORARIA ') and f'N1 MEDIA=BULA; DP={fonte_e4}' in U(w)
            and txt(lin[29]) == 'EXCECAO_TEMPORARIA' and lin[26] in (None, '') and len(tr) == 2 and trilha_ok
            and outros,
            {'cancelado': r_canc, 'retorno': r_e4, 'W': w[:220], 'AA_AE': lin[26:31], 'trilha': tr_txt[:2],
             'media': [lin[2], media_bula], 'dp': [lin[3], dp_ant]})

        r_de_novo = txt(run('GravarManualCore'))
        reg('E05 Depois de gravar: campos limpos e TRANCADOS de novo (cinza, sem marca); gravar de novo e recusado',
            trancado() and not txt(p.Range(f'B{R_MAN1}').Value) and not txt(p.Range(f'E{R_MAN1}').Value)
            and not txt(p.Range(f'B{R_MAN_JUST}').Value) and r_de_novo.startswith('ERRO|'),
            {'estado': estado(), 'retorno': r_de_novo[:160]})
        cor = int(p.Range(f'I{R_NV1}').Interior.Color or 0)
        reg('E06 Origem EXCECAO TEMPORARIA: Origem atual (I100) amarela com o aviso "temporario: substituir pelos '
            'valores do laboratorio"', cor == AMARELO_TEMP and 'SUBSTITUIR PELOS VALORES' in U(p.Range(f'I{R_NV1}').Value)
            and 'EXCECAO TEMPORARIA' in U(p.Range(f'I{R_NV1}').Value), {'cor': cor})

        resp('SIM')
        r_lib2 = txt(run('LiberarManualCore'))
        resp('')
        lib2 = estado()
        escrever_secao('QA_LOTE_INEXISTENTE', d_ini, fim_all, eventos=True)     # Worksheet_Change -> SecaoMudou
        depois = estado()
        escrever_secao(lote, d_ini, fim_all, eventos=True)
        reg('E07 Liberar e trocar o lote da secao (B98) tranca de novo a digitacao manual',
            r_lib2.startswith('OK|') and lib2.startswith('LIBERADA') and depois == 'TRANCADA' and trancado(),
            {'liberado': r_lib2, 'antes': lib2, 'depois': depois})

        # ------------------------------------------------------------ M04 excecao -> modo 2 -> modo 2
        def amarelo():
            return (int(p.Range(f'I{R_NV1}').Interior.Color or 0) == AMARELO_TEMP
                    and 'SUBSTITUIR PELOS VALORES' in U(p.Range(f'I{R_NV1}').Value))

        m4 = []
        for _ in range(2):
            escrever_secao(lote, d_ini, fim_aceit)
            resp('2|SIM|SIM')
            r_m4 = txt(run('AplicarAlvos'))
            resp('')
            lin = linha_ls()
            m4.append({'retorno': r_m4, 'cond_N1': txt(lin[29]), 'marcas': marcas(lin[29]), 'W': txt(lin[22])[:260],
                       'dp': lin[3], 'amarelo': amarelo()})
        reg('M04 Excecao (Media e DP) -> so a Media (2) -> so a Media (2) de novo: o DP da excecao continua marcado '
            '(CONDICAO_N1 = ACEITAVEL+EXCECAO_TEMPORARIA_DP), W herda "N1 DP (EXCECAO TEMPORARIA)", o DP nao muda e o '
            'amarelo fica nas duas vezes',
            all(x['retorno'].startswith('OK|') and x['marcas'] == ('', 'EXCECAO') and x['amarelo']
                and 'N1 DP (EXCECAO TEMPORARIA)' in U(x['W']) and perto(x['dp'], dp_ant) for x in m4)
            and U(m4[0]['cond_N1']) == 'ACEITAVEL+EXCECAO_TEMPORARIA_DP', {'rodadas': m4})

        # ------------------------------------------------------------ M05 W reescrita por outra gravacao
        w_m5 = 'digitado na aba Analitos (bloco do lote) QA'
        ls.Unprotect(SENHA)
        sem_eventos(lambda: ls.Cells(rLs, 23).__setattr__('Value', w_m5))
        ls.Protect(SENHA, False, True, True, True)
        run('AtualizarVistaAlvos')
        reg('M05 Origem W reescrita sem marca (bloco da aba Analitos): o amarelo continua pela marca do CONDICAO_N1',
            amarelo() and w_m5.upper() in U(p.Range(f'I{R_NV1}').Value),
            {'I100': txt(p.Range(f'I{R_NV1}').Value)[:200]})

        escrever_secao(lote, d_ini, fim_all)       # o M04 deixou a janela ACEITAVEL na secao; o E08 e no periodo todo
        a0 = aud0()
        antes_ls = linha_ls()
        resp('1|SIM|SIM')
        r = txt(run('AplicarAlvos'))
        resp('')
        ev = conferir(d_ini, fim_all, 'IDEAL', r, a0, antes_ls)
        cor = int(p.Range(f'I{R_NV1}').Interior.Color or 0)
        # o N2/N3 pode continuar TEMPORARIO pelo proprio N no periodo inteiro: o amarelo acompanha a origem
        temp_w = ('TEMPORARIO' in U(ev['W']) or 'EXCECAO TEMPORARIA' in U(ev['W'])
                  or any(any(marcas(linha_ls()[29 + 5 * k])) for k in range(nlv)))
        reg('E08 "Aplicar Alvos ao Lote" (Media e DP) substitui a excecao: origem MANUAL sem EXCECAO, CONDICAO_N1 = '
            'IDEAL, nenhum nivel com EXCECAO_TEMPORARIA; o amarelo so fica se a nova origem for TEMPORARIO pelo N',
            ev['ok'] and 'EXCECAO' not in U(ev['W']) and txt(linha_ls()[29]) == 'IDEAL'
            and not any('EXCECAO' in marcas(linha_ls()[29 + 5 * k]) for k in range(nlv))
            and (cor == AMARELO_TEMP) == temp_w
            and ('SUBSTITUIR' in U(p.Range(f'I{R_NV1}').Value)) == temp_w, dict(ev, cor=cor, temporaria=temp_w))

        # ------------------------------------------------------------ E09 excecao so do DP sobre valor PROPRIO
        resp('SIM')
        r_lib9 = txt(run('LiberarManualCore'))
        resp('')
        antes9 = linha_ls()
        dp9 = round(float(antes9[3]) * 1.07 + 0.001, 4)
        just9 = 'QA: DP da bula por excecao, media propria mantida.'
        p.Range(f'C{R_MAN1}').Value = dp9
        p.Range(f'D{R_MAN1}').Value = 'Bula'
        p.Range(f'B{R_MAN_JUST}').Value = just9
        resp('SIM|NAO')
        r_c9 = txt(run('GravarManualCore'))
        perg9 = txt(run('UltimaPergunta'))
        resp('')
        canc9 = r_c9 == 'CANCELADO|proprio' and linha_ls() == antes9 and 'PRÓPRIO' in U(perg9)
        a0 = aud0()
        resp('SIM|SIM')
        r_e9 = txt(run('GravarManualCore'))
        resp('')
        lin = linha_ls()
        tr9 = [' | '.join(txt(c) for c in l) for l in trilha(a0, 'PARAMETRO_LOTE_EXCECAO')]
        reg('E09 Excecao so do DP sobre valor PROPRIO: 2a confirmacao (cancelar nao grava); confirmada, so o DP muda, '
            'a Media e o N/periodo do N1 ficam, CONDICAO_N1 = IDEAL+EXCECAO_TEMPORARIA_DP, Audit_Log com BULA, '
            'justificativa e PROPRIO; amarelo',
            r_lib9.startswith('OK|') and canc9 and r_e9.startswith('OK|1|') and perto(lin[3], dp9)
            and lin[2] == antes9[2] and lin[26:29] == antes9[26:29] and U(lin[29]) == 'IDEAL+EXCECAO_TEMPORARIA_DP'
            and len(tr9) == 1 and just9 in tr9[0] and 'DP N1 = BULA' in U(tr9[0]) and 'PROPRIO' in U(tr9[0])
            and amarelo(),
            {'liberar': r_lib9, 'cancelado': r_c9, 'retorno': r_e9, 'AA_AE': lin[26:31], 'trilha': tr9[:1]})

        # ------------------------------------------------------------ E10 o proprio lote como "Lote anterior"
        resp('SIM')
        r_lib10 = txt(run('LiberarManualCore'))
        resp('')
        antes10 = linha_ls()
        p.Range(f'E{R_MAN1}').Value = lote                                   # eventos ligados: ManualEditado
        p.Range(f'C{R_MAN1}').Value = 1.234
        p.Range(f'D{R_MAN1}').Value = 'Lote anterior'
        p.Range(f'B{R_MAN_JUST}').Value = 'QA: tentativa com o proprio lote como lote anterior.'
        resp('SIM|SIM')
        r_e10 = txt(run('GravarManualCore'))
        resp('')
        run('BloquearManual')
        reg('E10 O proprio lote como "Lote anterior" e recusado; nada gravado',
            r_lib10.startswith('OK|') and r_e10.startswith('ERRO|') and 'PRÓPRIO LOTE' in U(r_e10)
            and linha_ls() == antes10, {'retorno': r_e10[:200]})

        # ------------------------------------------------------------ R01 referencia
        valores = sorted(float(r['RESULTADO']) for r in rs_lote if int(float(r['NIVEL'])) == 1)
        li, lsup = valores[len(valores) // 10], valores[-(len(valores) // 10) - 1]
        if li == valores[0]:
            li = valores[0] + 1e-9
        linha_grade = next((R_REF1 + i for i, v in enumerate(p.Range(f'A{R_REF1}:A{R_REF1 + N_REF - 1}').Value)
                            if U(v[0]) == AN), None)
        if linha_grade is None:
            raise PreCondicao(f'analito nao aparece na grade da referencia (A{R_REF1}:A{R_REF1 + N_REF - 1})',
                              {'analito': analito})
        media_ref = (li + lsup) / 2

        def obs():
            o = q.ws('RefBulaObs')
            u = o.Cells(o.Rows.Count, 1).End(-4162).Row
            return [tuple(l) for l in o.Range(f'A2:Q{u}').Value] if u >= 2 else []

        def do_lote(lst, st_=None):
            return [l for l in lst if U(l[2]) == LT and (st_ is None or U(l[10]) == st_)]

        ult_ls = ls.Cells(ls.Rows.Count, 16).End(-4162).Row + 5

        def foto_ls():
            return [tuple(l) for l in ls.Range(f'A1:AO{ult_ls}').Value]

        foto0 = foto_ls()
        p.Range(f'B{linha_grade}:D{linha_grade}').Value = [[media_ref, li, lsup]]     # eventos ligados
        ab = sum(1 for x in valores if x < li)
        ac = sum(1 for x in valores if x > lsup)
        o1 = do_lote(obs(), 'PENDENTE')
        sit = collections.Counter(txt(l[9]) for l in o1 if U(l[3]) == AN and int(l[4]) == 1)
        rb = q.ws('RefBulaStore')
        linha_rb = [tuple(l) for l in rb.Range('A2:O50').Value if U(l[2]) == f'{LT}|{AN}']
        reg('R01 Referencia digitada na grade: RefBulaStore gravada; abaixo do LI e acima do LS = PENDENTE (= oraculo); '
            'dentro da faixa nao gera observacao; alerta das novas',
            sit.get('ABAIXO DO LI', 0) == ab and sit.get('ACIMA DO LS', 0) == ac and len(linha_rb) == 1
            and perto(linha_rb[0][4], li) and 'NOVO' in U(run('UltimoAviso')),
            {'LI': li, 'LS': lsup, 'oraculo': [ab, ac], 'obs': dict(sit), 'store': linha_rb[:1],
             'aviso': txt(run('UltimoAviso'))[:200]})

        # ------------------------------------------------------------ R02 so nova alerta
        r2 = txt(run('VerificarReferencia'))
        reg('R02 Verificar de novo sem mudanca: 0 observacoes novas, as pendentes continuam',
            r2.startswith('OK|0|') and len(do_lote(obs(), 'PENDENTE')) == len(o1), {'retorno': r2})

        # ------------------------------------------------------------ R03 RETIRADA
        li2 = valores[0] - 1.0
        p.Range(f'C{linha_grade}').Value = li2                                        # eventos ligados
        o3 = do_lote(obs())
        ret = [l for l in o3 if U(l[10]) == 'RETIRADA']
        pend3 = [l for l in o3 if U(l[10]) == 'PENDENTE']
        reg('R03 LI corrigido: as pendentes antigas viram RETIRADA; as que continuam fora (acima do LS) entram de novo '
            'com a chave nova', len(ret) == len(o1) and len(pend3) == ac
            and all(txt(l[9]) == 'ACIMA DO LS' for l in pend3),
            {'retiradas': len(ret), 'pendentes_antes': len(o1), 'pendentes_agora': len(pend3), 'acima_oraculo': ac})

        # ------------------------------------------------------------ R04 parecer
        p.Range(f'A{R_PAR_TXT}').Value = 'curto'
        resp('SIM')
        r_curto = txt(run('RegistrarParecerCore'))
        resp('')
        texto = 'QA: limites da bula conferidos; resultados acima do LS em avaliacao, corrida aprovada pelo supervisor.'
        p.Range(f'A{R_PAR_TXT}').Value = texto
        a0 = aud0()
        resp('SIM')
        r4 = txt(run('RegistrarParecerCore'))
        resp('')
        o4 = do_lote(obs())
        pr = q.ws('RefBulaParecer')
        up = pr.Cells(pr.Rows.Count, 1).End(-4162).Row
        ultimo = tuple(pr.Range(f'A{up}:F{up}').Value[0]) if up >= 2 else ()
        ciente = [l for l in o4 if U(l[10]) == 'CIENTE']
        reg('R04 Parecer curto recusado; parecer registrado da ciencia a todas as pendentes do lote (CIENTE, id do '
            'parecer), RefBulaParecer com usuario e data/hora, trilha, campo do texto limpo',
            r_curto.startswith('ERRO|') and r4.startswith('OK|') and int(r4.split('|')[2]) == len(pend3)
            and len(ciente) == len(pend3) and not [l for l in o4 if U(l[10]) == 'PENDENTE']
            and all(txt(l[14]) == r4.split('|')[1] for l in ciente) and ultimo and txt(ultimo[5]) == texto
            and U(ultimo[3]) == LT and txt(ultimo[2]) and trilha(a0, 'REFERENCIA_PARECER_REGISTRADO')
            and not txt(p.Range(f'A{R_PAR_TXT}').Value),
            {'curto': r_curto[:120], 'retorno': r4, 'cientes': len(ciente), 'parecer': ultimo})

        # ------------------------------------------------------------ R05 CIENTE fica
        p.Range(f'D{linha_grade}').Value = valores[-1] + 1.0                          # LS acima de tudo
        o5 = do_lote(obs())
        reg('R05 CIENTE nunca e apagada nem retirada (limite corrigido de novo)',
            len([l for l in o5 if U(l[10]) == 'CIENTE']) == len(ciente)
            and not [l for l in o5 if U(l[10]) == 'PENDENTE'], {'status': dict(collections.Counter(U(l[10]) for l in o5))})
        foto1 = foto_ls()
        dif = [i for i, (a, b) in enumerate(zip(foto0, foto1), 1) if a != b]
        reg('R07 Bula/Unity e complemento: grade, verificacoes, RETIRADA e parecer nao gravaram NADA na LotesStore '
            '(A1:AO inteira igual)', not dif and len(foto0) == len(foto1), {'linhas_diferentes': dif[:10]})

        # ------------------------------------------------------------ R06 ATUALIZAR
        # LS do N1 gravado direto na RefBulaStore (eventos desligados): so o gancho do ATUALIZAR pode gerar as novas
        ls6 = valores[len(valores) // 2]
        acima6 = sum(1 for x in valores if x > ls6)
        rb = q.ws('RefBulaStore')
        r_rb = next(i for i, l in enumerate(rb.Range('C2:C200').Value, 2) if U(l[0]) == f'{LT}|{AN}')
        rb.Unprotect(SENHA)
        sem_eventos(lambda: rb.Cells(r_rb, 6).__setattr__('Value', ls6))            # F = N1_LS
        rb.Protect(SENHA, False, True, True, True)
        pend_antes = len(do_lote(obs(), 'PENDENTE'))
        r6 = txt(q.run('mIntegracao.AtualizarSoEntradas', '', teto=1800))
        wg, _k = oraculo.retorno_foto_wg(r6)
        pend6 = [l for l in do_lote(obs(), 'PENDENTE') if U(l[3]) == AN and int(l[4]) == 1]
        reg('R06 ATUALIZAR (so entradas) roda a verificacao no fim: as novas acima do LS viram PENDENTE (= oraculo), '
            'linha "Referência bula/Unity" no resumo, e o resumo ainda termina no retorno da foto do Westgard',
            r6.startswith('OK|') and bool(wg) and 'Referência bula/Unity' in r6 and pend_antes == 0
            and len(pend6) == acima6 and all(txt(l[9]) == 'ACIMA DO LS' for l in pend6),
            {'LS': ls6, 'oraculo_acima': acima6, 'pendentes': len(pend6), 'pendentes_antes': pend_antes,
             'resumo': r6[-300:]})
    except PreCondicao as e:
        reg('A00 pre-condicao ausente: ' + str(e.args[0]), False, e.args[1] if len(e.args) > 1 else {})
    except Exception as e:                           # noqa: BLE001
        reg('Z99 Execucao da suite sem excecao', False, {'erro': repr(e), 'trace': traceback.format_exc()[-1500:]})
    finally:
        q.fechar(salvar=False)
    h1 = sha256_arquivo(caminho)
    reg('Y00 Seguranca do teste: arquivo de entrada intacto (SHA-256 igual antes e depois)', h0 == h1,
        {'antes': h0, 'depois': h1})
    if saida:
        with open(saida, 'w', encoding='utf-8') as f:
            json.dump(RES, f, ensure_ascii=False, indent=1, default=str)
    falhas = [r for r in RES if r['resultado'] == 'FAIL']
    print(f'\n=== {produto} alvos manuais e referencia (ADR-102): {len(RES) - len(falhas)} PASS / {len(falhas)} FAIL ===',
          flush=True)
    return falhas


if __name__ == '__main__':
    falhas = executar(sys.argv[1], os.path.abspath(sys.argv[2]), sys.argv[3] if len(sys.argv) > 3 else None)
    sys.exit(1 if falhas else 0)
