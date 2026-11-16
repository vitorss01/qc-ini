# -*- coding: utf-8 -*-
"""valida_r6_adr096.py -- Tester/QA, ciclo 2026-11-15_reteste06, modo VALIDAR do ADR-096 (entrega 43).

Escrito pelo Tester/QA. NAO reaproveita reteste06_usuario_f1*.py / reteste06_carimbo_residual.py (Usuario Simulado) nem
o qa_adr095 (suite). Da infraestrutura usa so qa_final.QA (abre a copia, revela as abas, imprime EXCEL_PID),
repro_comum (meu: senha da aba, sha) e qa_inativar_ux.criar_usuario/fazer_login (usuario so na copia).

Ocorrencias validadas (criterios do Orquestrador):
  OC-0073  linha colada / colada e editada / ID MAN_ digitado repetido -> ID novo com aviso; ATUALIZAR; n e media
           (Painel, Estatistica) com oraculo proprio; LJ (Eng_Saida) e Westgard (1-3s); original intacto; mover mantem
           o ID; ID nao-MAN_ repetido -> MAN_INVALIDO/E04; repetido com eventos desligados -> CONFLITO_MANUAL, A01 novo,
           mensagem do ATUALIZAR conta; correcao apagando o ID. Ampliacao propria: colar ACIMA da origem, colar so
           valores (PasteSpecial), FillDown (a de cima fica), redigitar o proprio ID, editar sem tocar no ID, outro
           analito/lote que os do Usuario Simulado, e a reserva PQ (MOTOR_DADOS = PQ) comparada com o VBA.
  OC-0021  DETALHE na tela com " é resultado manual", " é retransmissão", "1ª linha" e a conjuncao intacta; MOTIVO do
           motor sem acento e igual no VBA e no PQ. Formula avaliada numa celula de rascunho com os textos dos motores.
  OC-0074  validacoes de lista (filtros do Registros, colunas da tblRegistros, Painel L4 Bio, Digitar Resultados):
           titulo, mensagem, ShowError, AlertStyle, itens aceitos e textos recusados (Validation.Value).
  OC-0075  (Bio) Estatistica H5/J5 com so o De / so o Ate: Text, Value2, Formula, formato e a grade A14:AU133.
  OC-0076  (Bio) aviso do corte de 180 corridas (Painel A39) com H3 vazio e com lote; n citado x oraculo.

Uso: python -u valida_r6_adr096.py <Bioquimica|Hematologia> <nova|base> [--sem-pq]
  nova = copia da producao atual (raiz); base = copia do backup anterior a entrega 43 (so as partes de leitura).
Regras: so COPIA nova em %TEMP%/qcw/auditoria/2026-11-15_reteste06/tester_validar/; nada e salvo; fechar() encerra so o
PID que o QA criou; nenhuma captura de tela. Dados de teste so na copia: usuario TESTE_AUDITORIA_R6V (ANALISTA), manuais
de 14/11/2026 entre 05:11 e 06:17 com valores sinteticos de 3 decimais, ID invalido TESTE-AUD-R6V.
"""
import datetime as dt
import hashlib
import json
import os
import re
import shutil
import statistics as stt
import sys
import threading
import time
import traceback

AQUI = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, AQUI)
import repro_comum as C  # noqa: E402
sys.path.insert(0, C.TESTES)
from qa_inativar_ux import criar_usuario, fazer_login  # noqa: E402
import qa_final  # noqa: E402

PROD = sys.argv[1]
MODO = sys.argv[2]
SEM_PQ = '--sem-pq' in sys.argv
SO_LISTAS = '--so-listas' in sys.argv
BIO = PROD == 'Bioquimica'
NLV = 2 if BIO else 3
CICLO = '2026-11-15_reteste06'
PASTA = os.path.join(os.environ['TEMP'], 'qcw', 'auditoria', CICLO, 'tester_validar')
BASES = {'Bioquimica': '_backup_pre_integracao_2026-11-15_1802', 'Hematologia': '_backup_pre_integracao_2026-11-15_2059'}
USR = 'TESTE_AUDITORIA_R6V'
EQ = 'DIMENSION 1' if BIO else 'XN1000'
EVITAR = {'GLICOSE', 'ALT (TGP)'} if BIO else {'WBC', 'BASO#'}     # analitos ja usados pelo Usuario Simulado / reproducao
DIA = '14/11/2026'
EPOCA = dt.datetime(1899, 12, 30)
SEGREDOS = [x for x in (os.environ.get('USERNAME'), os.environ.get('COMPUTERNAME')) if x and len(x) > 3]


def t(v):
    if isinstance(v, float) and v == int(v):
        return str(int(v))
    return '' if v is None else str(v).strip()


def up(v):
    return t(v).upper()


def num(v):
    return isinstance(v, (int, float)) and not isinstance(v, bool)


def dser(x):
    return (EPOCA + dt.timedelta(days=float(x))) if num(x) else None


def rel(a, b):
    if not num(a) or not num(b):
        return None
    return abs(a - b) / max(abs(a), abs(b), 1e-300)


def dec(x):
    return f'{x:.3f}'.replace('.', ',')


def sha_full(p):
    h = hashlib.sha256()
    with open(p, 'rb') as f:
        for b in iter(lambda: f.read(1 << 20), b''):
            h.update(b)
    return h.hexdigest()


class Log:
    def __init__(self, saida):
        self.saida, self.R = saida, {'inicio': dt.datetime.now().isoformat(timespec='seconds')}

    def __setitem__(self, k, v):
        self.R[k] = v
        print('##', k, json.dumps(v, ensure_ascii=False, default=str)[:1500], flush=True)
        self.gravar()

    def __getitem__(self, k):
        return self.R[k]

    def gravar(self):
        s = json.dumps(self.R, ensure_ascii=False, indent=1, default=str)
        for x in SEGREDOS:
            s = re.sub(re.escape(x), '<omitido>', s, flags=re.I)
        open(self.saida, 'w', encoding='utf-8').write(s)


class Vigia(threading.Thread):
    """Le o TEXTO e fecha caixas (#32770) e UserForms do PID desta sessao. Nao captura imagem."""

    def __init__(self, pid):
        super().__init__(daemon=True)
        self.pid, self.vistos, self.parar = pid, [], threading.Event()

    def run(self):
        import win32con
        import win32gui
        import win32process
        while not self.parar.is_set():
            achou = []

            def cb(h, _):
                try:
                    if win32process.GetWindowThreadProcessId(h)[1] == self.pid and \
                            win32gui.GetClassName(h) == '#32770' and win32gui.IsWindowVisible(h):
                        achou.append(h)
                except Exception:                                    # noqa: BLE001
                    pass
                return True
            try:
                win32gui.EnumWindows(cb, None)
            except Exception:                                        # noqa: BLE001
                pass
            for h in achou:
                tx = []

                def cc(ch, _):
                    try:
                        s = win32gui.GetWindowText(ch)
                        if s:
                            tx.append(s)
                    except Exception:                                # noqa: BLE001
                        pass
                    return True
                try:
                    win32gui.EnumChildWindows(h, cc, None)
                    self.vistos.append({'t': time.strftime('%H:%M:%S'), 'titulo': win32gui.GetWindowText(h),
                                        'textos': [x for x in tx if x not in ('OK',)]})
                    win32gui.PostMessage(h, win32con.WM_CLOSE, 0, 0)
                except Exception:                                    # noqa: BLE001
                    pass
            time.sleep(0.6 if achou else 0.2)


# =============================================================================================== leitura de tabelas
def tabela(qa, nome):
    lo = qa.lo(nome)
    cab = [str(c.Name) for c in lo.ListColumns]
    if lo.DataBodyRange is None:
        return cab, []
    v = lo.DataBodyRange.Value2
    return cab, [dict(zip(cab, l)) for l in v]


def final(qa):
    return tabela(qa, 'tblCQ_Final')[1]


def base_id(i):
    return t(i).split('~')[0]


def meus_final(rows, ids):
    ids = set(ids)
    return [{k: r.get(k) for k in ('ID_REGISTRO', 'STATUS_ANALITICO', 'PARTICIPA_ESTATISTICA', 'MOTIVO_EXCLUSAO_AUTOMATICA',
                                    'RUN', 'NIVEL', 'RESULTADO', 'ANALITO', 'LOTE', 'USUARIO_MANUAL', 'TIPO_PLOTAGEM_LJ')}
            for r in rows if base_id(r.get('ID_REGISTRO')) in ids]


def qa_meus(qa, ids):
    cab, rows = tabela(qa, 'tblQA_Integracao')
    out = []
    for r in rows:
        s = ' | '.join(t(v) for v in r.values())
        if any(i and i in s for i in ids):
            out.append({k: t(v) for k, v in r.items() if t(v)})
    return out


def oraculo(rows, analito, nivel, lote=None, ini=None, fim=None, equip=EQ):
    """Populacao independente: PARTICIPA = SIM, RESULTADO numerico, analito/equipamento/lote/nivel, DATA no dia [ini, fim]."""
    v = []
    for r in rows:
        if up(r.get('PARTICIPA_ESTATISTICA')) != 'SIM' or not num(r.get('RESULTADO')):
            continue
        if up(r.get('ANALITO')) != analito.upper() or t(r.get('NIVEL')) != str(nivel):
            continue
        if equip and up(r.get('EQUIPAMENTO')) != equip.upper():
            continue
        if lote is not None and t(r.get('LOTE')) != str(lote):
            continue
        d = r.get('DATA')
        if ini is not None and (not num(d) or int(d) < ini):
            continue
        if fim is not None and (not num(d) or int(d) > fim):
            continue
        v.append(r['RESULTADO'])
    o = {'n': len(v)}
    if v:
        o['media'] = stt.fmean(v)
    if len(v) > 1:
        o['dp'] = stt.stdev(v)
    return o


def serial(d):
    return (d - EPOCA).days


# =============================================================================================== OC-0074 listas
REFERS = {}


def itens_lista(xl, cel):
    f1 = t(cel.Validation.Formula1)
    if f1.startswith('='):
        try:
            try:
                REFERS[f1] = t(cel.Worksheet.Parent.Names(f1[1:]).RefersTo)
            except Exception:                                        # noqa: BLE001
                pass
            v = cel.Worksheet.Evaluate(f1[1:])
            v = v.Value if hasattr(v, 'Value') else v
        except Exception as e:                                       # noqa: BLE001
            return f1, f'erro {e}'
        if not isinstance(v, tuple):
            v = ((v,),)
        return f1, [x[0] for x in v if t(x[0])]                    # valor cru da lista (texto ou numero)
    return f1, [int(x) if x.isdigit() else x for x in f1.split(xl.International[5 - 1])]


def aceita(xl, cel, valor):
    ev = xl.EnableEvents
    xl.EnableEvents = False
    try:
        cel.Value = valor
        ok = bool(cel.Validation.Value)
        if not ok and isinstance(valor, str) and valor.isdigit():   # digitado vira numero
            cel.Value = int(valor)
            ok = bool(cel.Validation.Value) and 'numero'
        if not ok and isinstance(valor, str) and valor.isdigit():   # escolhido na seta: texto
            cel.Value = "'" + valor
            ok = bool(cel.Validation.Value) and 'texto'
        return ok
    except Exception as e:                                           # noqa: BLE001
        return f'erro {type(e).__name__}'
    finally:
        xl.EnableEvents = ev


def checar_validacao(qa, rotulo, cel, invalidos):
    xl = qa.ex.xl
    out = {'celula': f'{cel.Worksheet.Name}!{cel.Address}'}
    try:
        v = cel.Validation
        out.update({'Type': v.Type, 'AlertStyle': v.AlertStyle, 'ShowError': bool(v.ShowError),
                    'ErrorTitle': t(v.ErrorTitle), 'ErrorMessage': t(v.ErrorMessage)})
    except Exception as e:                                           # noqa: BLE001
        out['erro'] = f'sem validacao: {e}'
        return out
    f1, itens = itens_lista(xl, cel)
    out['Formula1'], out['itens'] = f1, itens
    antes = cel.Formula
    try:
        if isinstance(itens, list):
            out['aceitos'] = {t(i): aceita(xl, cel, i) for i in itens}
            out['itens'] = [t(i) for i in itens]
        out['invalidos'] = {i: aceita(xl, cel, i) for i in invalidos}
    finally:
        ev = xl.EnableEvents
        xl.EnableEvents = False
        cel.Formula = antes
        xl.EnableEvents = ev
    msg = out['ErrorMessage']
    out['ok_msg'] = (out['ShowError'] and out['AlertStyle'] == 1 and msg.startswith('Escolha ')
                     and 'na seta da célula.' in msg)
    out['refersto'] = REFERS.get(f1)
    out['ok_valores'] = (isinstance(itens, list) and len(itens) > 0 and all(x for x in out.get('aceitos', {}).values())
                         and all(x is False for x in out['invalidos'].values()))
    return out


def parte_listas(qa, L):
    wb = qa.wb
    senha = C.senha_planilha()
    ws = wb.Worksheets('Registros')
    try:
        ws.Unprotect(senha)            # so para escrever o valor de prova; a validacao (o que se confere) nao muda
    except Exception:                                                # noqa: BLE001
        pass
    lo = ws.ListObjects('tblRegistros')
    inval = ['TESTE-AUDITORIA-R6V', 'Sim,Não', 'Sim;Não', '9', 'DIMENSION 3', 'XN2000', 'ATIVO,INATIVO']
    res = {}
    for col, rot in ((2, 'filtro lote B4'), (3, 'filtro analito C4'), (7, 'filtro equipamento G4'),
                     (8, 'filtro Calibrado? H4'), (10, 'filtro status J4')):
        r = checar_validacao(qa, rot, ws.Cells(4, col), inval)
        r['ok_limpar'] = 'Limpar filtros' in r.get('ErrorMessage', '') and 'apague o filtro' in r.get('ErrorMessage', '')
        res[rot] = r
    cab = [str(c.Name) for c in lo.ListColumns]
    n = lo.ListRows.Count
    for idx in (4, 6, 10):
        rg = lo.ListColumns(idx).DataBodyRange
        cel = rg.Cells(rg.Rows.Count, 1)
        res[f'tblRegistros col {idx} {cab[idx - 1]}'] = checar_validacao(qa, cab[idx - 1], cel, inval)
    res['_tblRegistros'] = {'colunas': cab, 'linhas': n}
    res['_DETALHE_no_Registros'] = [c for c in cab if 'DETALHE' in c.upper()]
    if BIO:
        res['Painel L4'] = checar_validacao(qa, 'L4', wb.Worksheets('Painel').Range('L4'), inval)
    lm = qa.lo('tblResultados_Manuais')
    for col in ('EQUIPAMENTO', 'MATRIZ', 'MOTIVO', 'NIVEL'):
        try:
            rg = lm.ListColumns(col).DataBodyRange
            res[f'Digitar Resultados {col}'] = checar_validacao(qa, col, rg.Cells(rg.Rows.Count, 1), inval)
        except Exception as e:                                       # noqa: BLE001
            res[f'Digitar Resultados {col}'] = {'erro': str(e)}
    L['OC0074_listas'] = res


# =============================================================================================== OC-0075 / OC-0076 Painel
def por(qa, ws, addr, valor):
    c = ws.Range(addr)
    if valor is None:
        c.MergeArea.ClearContents()
    elif isinstance(valor, str):
        c.FormulaLocal = valor
    else:
        c.Value = valor
    qa.ex.esperar()
    time.sleep(0.4)


def ler_cel(ws, a):
    c = ws.Range(a)
    return {'text': t(c.Text), 'v2': c.Value2, 'f': t(c.Formula), 'fmt': t(c.NumberFormat)}


def grade_hash(ws):
    v = ws.Range('A14:AU133').Value2
    return hashlib.sha1(json.dumps(v, default=str).encode()).hexdigest()[:16]


def grade_linha(ws, analito, nivel):
    for l in ws.Range('A14:F133').Value2:
        if up(l[0]) == analito.upper() and t(l[1]) == str(nivel):
            return {'n': l[2], 'media': l[3], 'dp': l[4]}
    return None


def candidatos(rows, nlv, mes_ini, evitar):
    """(analito, lote) com mais resultados ativos do interfaceamento em nov/2026 no EQ, com os nlv niveis."""
    cont = {}
    for r in rows:
        if up(r.get('PARTICIPA_ESTATISTICA')) != 'SIM' or up(r.get('EQUIPAMENTO')) != EQ.upper():
            continue
        if not up(r.get('ORIGEM_RESULTADO')).startswith('INTERF') or not num(r.get('DATA')) or int(r['DATA']) < mes_ini:
            continue
        k = (t(r['ANALITO']), t(r['LOTE']))
        cont.setdefault(k, {}).setdefault(t(r['NIVEL']), 0)
        cont[k][t(r['NIVEL'])] += 1
    ok = [(k, sum(v.values())) for k, v in cont.items() if len(v) >= nlv and k[0].upper() not in evitar]
    return [k for k, _ in sorted(ok, key=lambda x: -x[1])]


def parte_estat_janela(qa, L, analitos):
    """OC-0075 (Bio): H5/J5 com so o De / so o Ate."""
    p, e = qa.ws('Painel'), qa.ws('Estatística')
    out = {}
    rows = final(qa)
    for an in analitos:
        por(qa, p, 'L4', EQ)
        por(qa, p, 'C3', an)
        for nome, g3, g4 in (('so_de_set', '01/09/2026', None), ('so_ate_set', None, '30/09/2026'),
                             ('ambos_set', '01/09/2026', '30/09/2026'), ('so_de_mai', '01/05/2026', None),
                             ('tudo', None, None)):
            por(qa, p, 'G4', None)
            por(qa, p, 'G3', None)
            if g3:
                por(qa, p, 'G3', g3)
            if g4:
                por(qa, p, 'G4', g4)
            qa.ex.xl.Calculate()
            qa.ex.esperar()
            cel = {a: ler_cel(e, a) for a in ('B5', 'D5', 'H5', 'I5', 'J5', 'H4')}
            cel['linha5'] = ' '.join(x for x in (cel['H5']['text'], cel['I5']['text'], cel['J5']['text']) if x)
            g = {nv: grade_linha(e, an, nv) for nv in (1, 2)}
            lote = t(e.Range('H4').Value)
            h5, j5 = e.Range('H5').Value2, e.Range('J5').Value2
            ini = int(h5) if num(h5) and h5 > 0 else None
            fim = int(j5) if num(j5) and j5 > 0 else None
            orc = {nv: oraculo(rows, an, nv, lote=lote, ini=ini, fim=fim) for nv in (1, 2)}
            cmp_ = {nv: {'n': (g[nv] or {}).get('n'), 'n_oraculo': orc[nv]['n'],
                         'rel_media': rel((g[nv] or {}).get('media'), orc[nv].get('media'))} for nv in (1, 2)}
            out[f'{an}|{nome}'] = {'painel': {'G3': t(p.Range('G3').Text), 'G4': t(p.Range('G4').Text),
                                              'K3': t(p.Range('K3').Text), 'K4': t(p.Range('K4').Text),
                                              'H3': t(p.Range('H3').Text)},
                                   'cel': cel, 'grade_hash': grade_hash(e), 'grade_vs_oraculo': cmp_,
                                   'tem_1900': '1900' in cel['linha5'] or '1900' in cel['B5']['text'] + cel['D5']['text']}
    L['OC0075_estatistica'] = out


def contar_corridas(rows, analito, equip, lote=None):
    s = set()
    for r in rows:
        if up(r.get('PARTICIPA_ESTATISTICA')) != 'SIM' or not num(r.get('RESULTADO')):
            continue
        if up(r.get('ANALITO')) != analito.upper() or up(r.get('EQUIPAMENTO')) != equip.upper():
            continue
        if lote is not None and t(r.get('LOTE')) != str(lote):
            continue
        s.add((t(r.get('RUN')), t(r.get('LOTE'))))
    return len(s)


def lotes_por_corridas(rows, analito, equip):
    c = {}
    for r in rows:
        if up(r.get('PARTICIPA_ESTATISTICA')) == 'SIM' and up(r.get('ANALITO')) == analito.upper() \
                and up(r.get('EQUIPAMENTO')) == equip.upper():
            c.setdefault(t(r.get('LOTE')), set()).add(t(r.get('RUN')))
    return sorted(((k, len(v)) for k, v in c.items()), key=lambda x: -x[1])


def parte_corte(qa, L, cenarios):
    """OC-0076: Painel A39 com H3 vazio e com lote, periodo inteiro ("Tudo")."""
    p, eng = qa.ws('Painel'), qa.ws('Eng_Saida')
    rows = final(qa)
    out = {}
    for an, equip, lote in cenarios:
        if BIO:
            por(qa, p, 'L4', equip)
        por(qa, p, 'C3', an)
        por(qa, p, 'G4', None)
        por(qa, p, 'G3', None)
        por(qa, p, 'H3', ("'" + lote) if lote else None)
        qa.ex.xl.Calculate()
        qa.ex.esperar()
        a39 = t(p.Range('A39').Text)
        af1, m1 = t(eng.Range('AF1').Value), eng.Range('M1').Value
        orc = {nv: oraculo(rows, an, nv, lote=lote or None, equip=equip)['n'] for nv in range(1, NLV + 1)}
        esperado_af1 = ' · '.join(f'N{nv} {orc[nv]}' for nv in range(1, NLV + 1))
        out[f'{an}|{equip}|{lote or "(todos)"}'] = {
            'H3': t(p.Range('H3').Text), 'K4': t(p.Range('K4').Text), 'A39': a39, 'A39_formula_fim': t(p.Range('A39').Formula)[-300:],
            'Eng_AF1': af1, 'Eng_M1': m1, 'oraculo_n_periodo': orc, 'AF1_igual_oraculo': af1 == esperado_af1,
            'corridas_oraculo': contar_corridas(rows, an, equip, lote or None),
            'cita_estatistica': 'Estatística' in a39, 'todos_os_lotes': '(todos os lotes)' in a39}
    L['OC0076_corte'] = out


# =============================================================================================== OC-0021 formula
LITERAIS = {   # texto do motor (sem acento) -> o que a tela deve mostrar (escrito a mao, nao vem da lista do produto)
    'SUBSTITUI_ID invalido: MAN_0042 e resultado manual (SUBSTITUI_ID aceita so resultado do interfaceamento)':
        'SUBSTITUI_ID inválido: MAN_0042 é resultado manual (SUBSTITUI_ID aceita só resultado do interfaceamento)',
    '123456 e retransmissao de 123455 (informe o ID original)':
        '123456 é retransmissão de 123455 (informe o ID original)',
    'ID_REGISTRO manual repetido (vale a 1a linha)': 'ID_REGISTRO manual repetido (vale a 1ª linha)',
    'resultado <= 0: implausivel para material de controle (fora da estatistica e do grafico)':
        'resultado <= 0: implausível para material de controle (fora da estatística e do gráfico)',
    'resultado manual repetido (mesma chave e instante)': 'resultado manual repetido (mesma chave e instante)',
    'o ID 777 nao existe nos resultados do interfaceamento': 'o ID 777 não existe nos resultados do interfaceamento',
    'campos obrigatorios ausentes/invalidos: DATA/HORA posterior ao lancamento':
        'campos obrigatórios ausentes/inválidos: DATA/HORA posterior ao lançamento',
}


def parte_detalhe_formula(qa, L):
    lo = qa.lo('tblResultados_Manuais')
    f = t(lo.ListColumns('DETALHE').DataBodyRange.Cells(1, 1).Formula)
    i0 = f.find('SUBSTITUTE(')
    j = f.rfind(',"atualize os dados"')
    expr = f[i0:j] if i0 >= 0 and j > i0 else None
    out = {'formula_len': len(f), 'formula_fim': f[-260:], 'n_SUBSTITUTE': f.count('SUBSTITUTE(')}
    if not expr:
        out['erro'] = 'nao achei a cadeia SUBSTITUTE'
        L['OC0021_formula'] = out
        return
    m = re.search(r'XLOOKUP\(.*?\)&""', expr)
    out['xlookup'] = m.group(0) if m else None
    ws = qa.wb.Worksheets.Add()
    try:
        res = {}
        motivos = sorted({re.sub(r'\d', '#', t(r.get('MOTIVO_EXCLUSAO_AUTOMATICA'))) for r in final(qa)
                          if t(r.get('MOTIVO_EXCLUSAO_AUTOMATICA'))})
        provas = list(LITERAIS) + [x for x in motivos if x not in LITERAIS]
        for k, s in enumerate(provas, 1):
            lit = '"' + s.replace('"', '""') + '"'
            ws.Range(f'A{k}').Formula2 = '=' + expr.replace(m.group(0), lit)
        qa.ex.xl.Calculate()
        for k, s in enumerate(provas, 1):
            res[s] = t(ws.Range(f'A{k}').Value)
        out['literais'] = {s: {'tela': res[s], 'esperado': LITERAIS[s], 'ok': res[s] == LITERAIS[s]} for s in LITERAIS}
        sobra = {}
        for s in provas[len(LITERAIS):]:
            r = res[s]
            if re.search(r'\b(nao|ja|so|estatistica|grafico|implausivel|lancado|lancamento|invalido|obrigatorio|'
                         r'retransmissao|substituido|catalogo)\b', r) or re.search(r'(?<![\w])\d+a(?![\w])', r) or \
                    re.search(r'#+ e (resultado|retransmiss)', r):
                sobra[s] = r
        out['motivos_distintos_da_copia'] = {s: res[s] for s in provas[len(LITERAIS):]}
        out['sobra_sem_acento'] = sobra
    finally:
        ev = qa.ex.xl.DisplayAlerts
        qa.ex.xl.DisplayAlerts = False
        ws.Delete()
        qa.ex.xl.DisplayAlerts = ev
    L['OC0021_formula'] = out


# =============================================================================================== OC-0073 manual
class Manual:
    def __init__(self, qa, vig):
        self.qa, self.vig = qa, vig
        self.lo = qa.lo('tblResultados_Manuais')
        self.ws = self.lo.Parent
        self.cab = [str(c.Name) for c in self.lo.ListColumns]
        self.k = 0

    def cel(self, r, col):
        return self.lo.ListColumns(col).DataBodyRange.Cells(r, 1)

    def faixa(self, r1, r2=None):
        return self.ws.Range(self.cel(r1, 'ID_REGISTRO'), self.cel(r2 or r1, 'MOTIVO'))

    def linha(self, r):
        rg = self.lo.ListRows(r).Range
        v = rg.Value2[0]
        d = {'_planilha': rg.Row}
        for i, c in enumerate(self.cab):
            if c in ('STATUS', 'DETALHE', 'RUN', 'DATA', 'HORA', 'REGISTRADO_EM'):
                d[c] = t(rg.Cells(1, i + 1).Text)
            else:
                d[c] = v[i] if num(v[i]) else t(v[i])
        return d

    def id(self, r):
        return t(self.cel(r, 'ID_REGISTRO').Value)

    def esc(self, r, col, v):
        c = self.cel(r, col)
        c.FormulaLocal = v
        self.qa.ex.esperar()
        time.sleep(0.25)

    def digitar(self, r, campos):
        for col, v in campos:
            self.esc(r, col, v)

    def avisos(self):
        time.sleep(1.2)
        novos = self.vig.vistos[self.k:]
        self.k = len(self.vig.vistos)
        return novos

    def livre(self):
        ids = self.lo.ListColumns('ID_REGISTRO').DataBodyRange.Value2
        dts = self.lo.ListColumns('DATA').DataBodyRange.Value2
        for i in range(len(ids)):
            if not t(ids[i][0]) and not t(dts[i][0]):
                return i + 1
        raise RuntimeError('tabela manual cheia')


def campos(lote, nv, an, hora, val, equip=True):
    c = [('DATA', DIA), ('HORA', hora)]
    if equip and BIO:
        c.append(('EQUIPAMENTO', EQ))
    return c + [('LOTE', "'" + lote), ('NIVEL', str(nv)), ('ANALITO', an), ('RESULTADO', dec(val)), ('MOTIVO', 'Outro')]


def painel_filtro(qa, an, lote):
    p = qa.ws('Painel')
    if BIO:
        por(qa, p, 'L4', EQ)
    por(qa, p, 'C3', an)
    por(qa, p, 'H3', "'" + lote)
    por(qa, p, 'K3', 2026)
    por(qa, p, 'K4', 'Nov')
    qa.ex.xl.Calculate()
    qa.ex.esperar()
    eng = [list(l) for l in qa.nome('engPainel').Value][:NLV]
    return {'filtro': {c: t(p.Range(c).Text) for c in ('L4', 'C3', 'H3', 'K3', 'K4', 'G3', 'G4')},
            'engPainel': [{'nivel': l[0], 'n': l[1], 'media': l[2], 'dp': l[3]} for l in eng],
            'A7': [[t(x) for x in l] for l in p.Range(f'A7:T{6 + NLV}').Value]}


def eng_slots(qa, ids):
    e = qa.ws('Eng_Saida')
    n = int(e.Range('I1').Value or 0)
    if not n:
        return {'n': 0}
    cab = [t(x) for x in e.Range(e.Cells(2, 1), e.Cells(2, 80)).Value[0]]
    blk = e.Range(e.Cells(3, 1), e.Cells(2 + n, 80)).Value
    out = []
    for k, r in enumerate(blk):
        idc = [t(x) for x in r[38:38 + NLV]]
        if any(base_id(i) in ids for i in idc if i):
            flags = {}
            for nv in range(NLV):
                for c in range(2 + 7 * nv, 9 + 7 * nv):
                    if r[c] not in (None, '', 0):
                        flags[f'N{nv + 1}:{cab[c] or c}'] = r[c]
            out.append({'slot': k + 1, 'run': t(r[1]), 'ids': idc, 'val': list(r[24:24 + NLV]), 'flags': flags})
    return {'n': n, 'C1': t(e.Range('C1').Value), 'E1': t(e.Range('E1').Value), 'meus': out}


def estat_hema(qa, an, lote):
    e = qa.ws('Estatística')
    for c in ('B3', 'D3'):
        por(qa, e, c, 2026)
    qa.ex.xl.Calculate()
    qa.ex.esperar()
    return {'B3': t(e.Range('B3').Text), 'D3': t(e.Range('D3').Text), 'B4': t(e.Range('B4').Text),
            'grade': {nv: grade_linha(e, an, nv) for nv in range(1, NLV + 1)}}


def parte_manual(qa, L):
    xl = qa.ex.xl
    rows0 = final(qa)
    nov1 = serial(dt.datetime(2026, 11, 1))
    cand = candidatos(rows0, NLV, nov1, EVITAR)
    (AN, LOTE) = None, None
    for an, lote in cand:
        alvo = [(qa.run('mLotes.MediaDoLote', lote, an, nv, teto=60), qa.run('mLotes.DPDoLote', lote, an, nv, teto=60))
                for nv in range(1, NLV + 1)]
        if all(num(a) and num(b) and b > 0 for a, b in alvo):
            AN, LOTE, ALVO = an, lote, alvo
            break
    (AN2, LOTE2) = next(((a, l) for a, l in cand if a != AN), (AN, LOTE))
    rec = {}
    for nv in range(1, NLV + 1):
        v = [r['RESULTADO'] for r in sorted((r for r in rows0 if up(r.get('ANALITO')) == AN.upper()
                                             and t(r.get('LOTE')) == LOTE and t(r.get('NIVEL')) == str(nv)
                                             and up(r.get('PARTICIPA_ESTATISTICA')) == 'SIM' and num(r.get('RESULTADO'))
                                             and num(r.get('DATA_HORA'))), key=lambda r: r['DATA_HORA'])][-20:]
        rec[nv] = stt.fmean(v)
    m2, s2 = ALVO[1]
    z4 = round(m2 + 4.0002 * s2, 3)
    if (z4 - m2) / s2 <= 3.9:
        z4 = round(m2 + 4.01 * s2, 3)
    V = {'v1': rec[1] + 0.002, 'v2': rec[1] - 0.003, 'v4': rec[2] + 0.001, 'v5': rec[1] + 0.004, 'v6': rec[1] - 0.001,
         'v7': rec[2] - 0.002, 'v8': rec[2] + 0.003, 'z4': z4}
    V = {k: round(x, 3) for k, x in V.items()}
    L['ref'] = {'analito': AN, 'lote': LOTE, 'alvo_lote': ALVO, 'media_recente': rec, 'valores': V,
                'z_do_z4': (V['z4'] - m2) / s2, 'analito_detalhe': AN2, 'lote_detalhe': LOTE2, 'candidatos': cand[:6]}

    pre = painel_filtro(qa, AN, LOTE)
    L['painel_antes'] = pre

    # ---------------------------------------------------------------- digitacao (eventos ligados, Interactive)
    criar_usuario(qa, USR, 'ANALISTA')
    L['login'] = fazer_login(qa, USR)
    qa.run('mIntegracao.IrDigitarResultados', teto=120)
    M = Manual(qa, VIG)
    f = M.livre()
    livres = all(not t(M.cel(f + i, 'ID_REGISTRO').Value) and not t(M.cel(f + i, 'DATA').Value) for i in range(22))
    L['tabela_manual'] = {'primeira_livre': f, 'planilha': M.cel(f, 'ID_REGISTRO').Row, '22_livres': livres,
                          'protegida': bool(M.ws.ProtectContents), 'cursor': t(xl.ActiveCell.Address)}
    xl.Interactive = True
    P = {}
    passos = {}
    try:
        rO = f + 2
        M.digitar(rO, campos(LOTE, 1, AN, '05:11', V['v1']))
        P['O'] = rO
        orig = M.linha(rO)
        passos['O_origem'] = {'linha': orig, 'avisos': M.avisos()}
        idO = orig['ID_REGISTRO']

        def igual_origem():
            a = M.linha(rO)
            return {k: a[k] for k in a if a[k] != orig[k] and k not in ('STATUS', 'DETALHE', 'RUN')}

        # S1 colar sem editar logo abaixo
        M.faixa(rO).Copy(M.faixa(f + 3))
        xl.CutCopyMode = False
        passos['S1_colar_sem_editar'] = {'id': M.id(f + 3), 'avisos': M.avisos(), 'origem_mudou': igual_origem()}
        P['S1'] = f + 3
        # S2 colar ACIMA da origem e editar HORA e RESULTADO
        M.faixa(rO).Copy(M.faixa(f))
        xl.CutCopyMode = False
        id_col = M.id(f)
        av = M.avisos()
        M.digitar(f, [('HORA', '05:17'), ('RESULTADO', dec(V['v2']))])
        passos['S2_colar_acima'] = {'id_ao_colar': id_col, 'id_depois_editar': M.id(f), 'avisos': av + M.avisos(),
                                    'origem_mudou': igual_origem(), 'id_origem_agora': M.id(rO)}
        P['S2'] = f
        # S3 colar so valores (PasteSpecial) e editar NIVEL 2 + RESULTADO media+4DP
        M.faixa(rO).Copy()
        M.faixa(f + 4).PasteSpecial(-4163)
        xl.CutCopyMode = False
        id_col = M.id(f + 4)
        av = M.avisos()
        M.digitar(f + 4, [('NIVEL', '2'), ('RESULTADO', dec(V['z4']))])
        passos['S3_colar_valores_editar'] = {'id_ao_colar': id_col, 'id_depois_editar': M.id(f + 4),
                                             'avisos': av + M.avisos(), 'origem_mudou': igual_origem()}
        P['S3'] = f + 4
        # S4 digitar o ID da origem (minusculo, sem zeros, com espacos) numa linha vazia, depois os dados
        n_o = int(idO.split('_')[1])
        M.esc(f + 5, 'ID_REGISTRO', f' man_{n_o} ')
        id_dig = M.id(f + 5)
        av = M.avisos()
        M.digitar(f + 5, campos(LOTE, 2, AN, '05:23', V['v4']))
        passos['S4_id_digitado'] = {'digitado': f' man_{n_o} ', 'id_apos_digitar': id_dig, 'id_final': M.id(f + 5),
                                    'avisos': av + M.avisos()}
        P['S4'] = f + 5
        # S5 linha nova digitada e FillDown de 3 linhas (a de cima fica com o ID)
        M.digitar(f + 6, campos(LOTE, 1, AN, '05:29', V['v5']))
        idP = M.id(f + 6)
        M.avisos()
        M.faixa(f + 6, f + 8).FillDown()
        ids_fd = [M.id(f + 6), M.id(f + 7), M.id(f + 8)]
        av = M.avisos()
        M.digitar(f + 7, [('HORA', '05:31'), ('RESULTADO', dec(V['v6']))])
        M.digitar(f + 8, [('NIVEL', '2'), ('HORA', '05:37'), ('RESULTADO', dec(V['v7']))])
        passos['S5_filldown'] = {'id_P': idP, 'ids_apos_filldown': ids_fd, 'avisos': av + M.avisos(),
                                 'ids_finais': [M.id(f + 6), M.id(f + 7), M.id(f + 8)]}
        P['P'], P['FD1'] = f + 6, f + 7
        # S7 sem falso positivo: editar o RESULTADO sem tocar no ID; redigitar o proprio ID
        M.esc(f + 6, 'RESULTADO', dec(V['v5']))
        a1 = M.id(f + 6)
        M.esc(f + 6, 'ID_REGISTRO', idP)
        passos['S7_sem_falso_positivo'] = {'id_P': idP, 'apos_editar_resultado': a1, 'apos_redigitar_id': M.id(f + 6),
                                           'avisos': M.avisos()}
        # S6 mover (recortar A:K da f+8 e colar em f+11)
        id_mov = M.id(f + 8)
        M.faixa(f + 8).Cut(M.faixa(f + 11))
        passos['S6_mover'] = {'id_antes': id_mov, 'id_destino': M.id(f + 11), 'origem_vazia': not M.id(f + 8)
                              and not t(M.cel(f + 8, 'DATA').Value), 'avisos': M.avisos(),
                              'carimbo_na_origem': {c: t(M.cel(f + 8, c).Text) for c in ('USUARIO', 'REGISTRADO_EM')}}
        P['MOV'] = f + 11
        # S8 ID fora do padrao repetido em duas linhas
        for r, h, v in ((f + 12, '05:47', V['v1'] + 0.005), (f + 13, '05:53', V['v1'] + 0.006)):
            M.esc(r, 'ID_REGISTRO', 'TESTE-AUD-R6V')
            M.digitar(r, campos(LOTE, 1, AN, h, v))
        passos['S8_id_invalido'] = {'ids': [M.id(f + 12), M.id(f + 13)], 'avisos': M.avisos()}
        P['INV1'], P['INV2'] = f + 12, f + 13
        # S9 repetido com os eventos DESLIGADOS (dado antigo)
        xl.EnableEvents = False
        M.faixa(rO).Copy(M.faixa(f + 14))
        xl.CutCopyMode = False
        M.cel(f + 14, 'HORA').FormulaLocal = '05:59'
        M.cel(f + 14, 'NIVEL').FormulaLocal = '2'
        M.cel(f + 14, 'RESULTADO').FormulaLocal = dec(V['v8'])
        xl.EnableEvents = True
        passos['S9_eventos_desligados'] = {'id': M.id(f + 14), 'linha': M.linha(f + 14), 'avisos': M.avisos()}
        P['REP'] = f + 14
        # D1 SUBSTITUI_ID = um MAN_ (outro analito)  /  D3 RESULTADO 0 (implausivel)  /  D2 retransmissao
        M.digitar(f + 15, campos(LOTE2, 1, AN2, '06:05', 1.234))
        M.esc(f + 15, 'SUBSTITUI_ID', idO)
        passos['D1_subst_man'] = {'linha': M.linha(f + 15), 'avisos': M.avisos()}
        M.digitar(f + 16, campos(LOTE2, 1, AN2, '06:11', 0))
        passos['D3_zero'] = {'linha': M.linha(f + 16), 'avisos': M.avisos()}
        dup = next((r for r in rows0 if up(r.get('STATUS_ANALITICO')) == 'DUPLICIDADE_ORIGEM' and t(r.get('ID_REGISTRO'))
                    and up(r.get('ORIGEM_RESULTADO')).startswith('INTERF')), None)
        if dup:
            M.esc(f + 17, 'SUBSTITUI_ID', t(dup['ID_REGISTRO']))
            lin = M.linha(f + 17)
            for col in ('RESULTADO',):
                if not t(lin.get(col)):
                    M.esc(f + 17, col, dec(dup['RESULTADO']) if num(dup.get('RESULTADO')) else '1')
            M.esc(f + 17, 'MOTIVO', 'Outro')
            passos['D2_subst_retransmissao'] = {'dup': {k: dup.get(k) for k in ('ID_REGISTRO', 'ANALITO', 'NIVEL', 'LOTE',
                                                                               'MOTIVO_EXCLUSAO_AUTOMATICA')},
                                                'linha': M.linha(f + 17), 'avisos': M.avisos()}
            P['D2'] = f + 17
        else:
            passos['D2_subst_retransmissao'] = 'sem DUPLICIDADE_ORIGEM na tblCQ_Final da copia'
        P['D1'], P['D3'] = f + 15, f + 16
        passos['origem_final'] = {'mudou': igual_origem()}
    finally:
        xl.Interactive = False
        xl.EnableEvents = True
    L['passos'] = passos
    tela0 = {k: M.linha(r) for k, r in P.items()}
    L['tela_antes_atualizar'] = tela0
    ids = {v['ID_REGISTRO'] for v in tela0.values() if v['ID_REGISTRO']}
    try:
        L['marca_qcUltimoIdManual'] = t(qa.wb.Names('qcUltimoIdManual').RefersTo)
    except Exception as e:  # noqa: BLE001
        L['marca_qcUltimoIdManual'] = f'sem nome: {e}'

    # ---------------------------------------------------------------- ATUALIZAR #1 (VBA)
    def depois(rotulo):
        rows = final(qa)
        R = {'tela': {k: M.linha(r) for k, r in P.items()}, 'final': meus_final(rows, ids), 'qa': qa_meus(qa, ids),
             'ResumoQA': t(qa.run('mIntegracao.ResumoQA', teto=120)), 'motor': t(qa.cfg('MOTOR_DADOS'))}
        pa = painel_filtro(qa, AN, LOTE)
        R['painel'] = pa
        nov30 = serial(dt.datetime(2026, 11, 30))
        R['oraculo_painel'] = {nv: oraculo(rows, AN, nv, lote=LOTE, ini=nov1, fim=nov30) for nv in range(1, NLV + 1)}
        R['painel_vs_oraculo'] = {nv: {'n': (pa['engPainel'][nv - 1]['n'], R['oraculo_painel'][nv]['n']),
                                       'rel_media': rel(pa['engPainel'][nv - 1]['media'], R['oraculo_painel'][nv].get('media')),
                                       'rel_dp': rel(pa['engPainel'][nv - 1]['dp'], R['oraculo_painel'][nv].get('dp'))}
                                  for nv in range(1, NLV + 1)}
        R['eng'] = eng_slots(qa, ids)
        if BIO:
            e = qa.ws('Estatística')
            lote_e = t(e.Range('H4').Value)
            h5, j5 = e.Range('H5').Value2, e.Range('J5').Value2
            g = {nv: grade_linha(e, AN, nv) for nv in (1, 2)}
            o = {nv: oraculo(rows, AN, nv, lote=lote_e, ini=int(h5) if num(h5) and h5 else None,
                             fim=int(j5) if num(j5) and j5 else None) for nv in (1, 2)}
            R['estatistica'] = {'H4': lote_e, 'H5': t(e.Range('H5').Text), 'J5': t(e.Range('J5').Text), 'grade': g,
                                'oraculo': o, 'rel_media': {nv: rel((g[nv] or {}).get('media'), o[nv].get('media'))
                                                            for nv in (1, 2)}}
        else:
            es = estat_hema(qa, AN, LOTE)
            lote_e = es['B4'] or LOTE
            o = {nv: oraculo(rows, AN, nv, lote=lote_e, ini=serial(dt.datetime(2026, 1, 1)),
                             fim=serial(dt.datetime(2026, 12, 31))) for nv in range(1, NLV + 1)}
            es['oraculo'] = o
            es['rel_media'] = {nv: rel((es['grade'][nv] or {}).get('media'), o[nv].get('media')) for nv in range(1, NLV + 1)}
            R['estatistica'] = es
        return R, rows

    msg1, s1 = qa.atualizar()
    R1, rows1 = depois('1')
    R1['mensagem_atualizar'], R1['segundos'] = msg1, round(s1, 1)
    L['atualizar1'] = R1

    # ---------------------------------------------------------------- reserva PQ: mesmo estado, motor PQ
    if not SEM_PQ:
        try:
            antes_motor = qa.cfg('MOTOR_DADOS', 'PQ')
            t0 = time.time()
            r = str(qa.run('mIntegracao.AtualizarDadosAutomatico', teto=1800))
            rows_pq = final(qa)
            fq = {t(x['ID_REGISTRO']): x for x in meus_final(rows_pq, ids)}
            fv = {t(x['ID_REGISTRO']): x for x in R1['final']}
            dif = {}
            for k in sorted(set(fq) | set(fv)):
                a, b = fv.get(k, {}), fq.get(k, {})
                d = {c: (a.get(c), b.get(c)) for c in ('STATUS_ANALITICO', 'PARTICIPA_ESTATISTICA',
                                                        'MOTIVO_EXCLUSAO_AUTOMATICA', 'RESULTADO', 'NIVEL')
                     if t(a.get(c)) != t(b.get(c))}
                if d:
                    dif[k] = d
            qpq = qa_meus(qa, ids)
            def a01(lst):
                return sorted(v for x in lst if 'A01' in x.values() for v in x.values()
                              if 'inative' in v or 'apague' in v or 'FORA da estatistica' in v)
            a01v, a01p = a01(R1['qa']), a01(qpq)
            L['reserva_pq'] = {'motor_antes': antes_motor, 'retorno': r[:600], 'segundos': round(time.time() - t0, 1),
                               'final_pq': list(fq.values()), 'diferencas_vba_pq': dif, 'qa_pq': qpq,
                               'A01_vba': a01v, 'A01_pq': a01p, 'A01_igual': a01v == a01p,
                               'tblCQ_Final_linhas': (len(rows1), len(rows_pq))}
        finally:
            qa.cfg('MOTOR_DADOS', 'VBA')

    # ---------------------------------------------------------------- correcao do repetido (apagar o ID) e ATUALIZAR #2
    xl.Interactive = True
    try:
        M.cel(P['REP'], 'ID_REGISTRO').ClearContents()
        qa.ex.esperar()
        L['correcao'] = {'id_novo': M.id(P['REP']), 'avisos': M.avisos(), 'linha': M.linha(P['REP'])}
    finally:
        xl.Interactive = False
    ids.add(M.id(P['REP']))
    msg2, s2_ = qa.atualizar()
    R2, _ = depois('2')
    R2['mensagem_atualizar'], R2['segundos'] = msg2, round(s2_, 1)
    L['atualizar2'] = R2
    L['avisos_todos'] = VIG.vistos


VIG = None


def main():
    global VIG
    os.makedirs(PASTA, exist_ok=True)
    src = os.path.join(C.REPO, f'QC_{PROD}.xlsm') if MODO == 'nova' else os.path.join(C.REPO, BASES[PROD], f'QC_{PROD}.xlsm')
    tag = time.strftime('%H%M%S')
    dst = os.path.join(PASTA, f'{PROD}_{MODO}_{tag}.xlsm')
    L = Log(os.path.join(PASTA, f'valida_r6_{PROD}_{MODO}{"_listas" if SO_LISTAS else ""}.json'))
    h0 = sha_full(src)
    shutil.copy2(src, dst)
    L['origem'] = {'arquivo': os.path.relpath(src, C.REPO), 'sha256_antes': h0[:16], 'copia': os.path.basename(dst),
                   'sha_copia': sha_full(dst)[:16]}
    qa = qa_final.QA(PROD, dst)
    L['excel_pid'] = qa.ex.pid
    VIG = Vigia(qa.ex.pid)
    VIG.start()
    try:
        rows = final(qa)
        L['tblCQ_Final'] = {'linhas': len(rows),
                            'hash_value2': hashlib.sha1(json.dumps([sorted((k, t(v)) for k, v in r.items()) for r in rows],
                                                                   default=str).encode()).hexdigest()[:16]}
        for nome, fn in (('listas', lambda: parte_listas(qa, L)),
                         ('detalhe_formula', lambda: parte_detalhe_formula(qa, L))):
            if SO_LISTAS and nome != 'listas':
                continue
            try:
                fn()
            except Exception as e:                                   # noqa: BLE001
                L[f'ERRO_{nome}'] = f'{type(e).__name__}: {e} | {traceback.format_exc()[-900:]}'
        try:
            if SO_LISTAS:
                raise StopIteration('so listas')
            if BIO:
                outro = next(a for a, l in candidatos(rows, 2, serial(dt.datetime(2026, 1, 1)), {'GLICOSE'}) if l == '8977')
                parte_estat_janela(qa, L, ['Glicose', outro])
                lt = lotes_por_corridas(rows, 'Glicose', 'DIMENSION 1')
                d2 = {}
                for r in rows:
                    if up(r.get('EQUIPAMENTO')) == 'DIMENSION 2' and up(r.get('PARTICIPA_ESTATISTICA')) == 'SIM':
                        d2.setdefault(t(r['ANALITO']), set()).add((t(r['RUN']), t(r['LOTE'])))
                an_d2 = sorted(d2, key=lambda a: -len(d2[a]))
                an_d2 = next((a for a in an_d2 if a.upper() != 'HEMOGLOBINA GLICADA'), an_d2[0])
                L['corte_escolhas'] = {'lotes_glicose_d1': lt[:4], 'analito_d2': an_d2, 'corridas_d2': len(d2[an_d2])}
                parte_corte(qa, L, [('Glicose', 'DIMENSION 1', None), ('Glicose', 'DIMENSION 1', lt[0][0]),
                                    (an_d2, 'DIMENSION 2', None), (outro, 'DIMENSION 1', None)])
            else:
                parte_corte(qa, L, [('WBC', EQ, None), ('PLT', EQ, None)])
        except Exception as e:                                       # noqa: BLE001
            L['ERRO_painel'] = f'{type(e).__name__}: {e} | {traceback.format_exc()[-900:]}'
        if MODO == 'nova' and not SO_LISTAS:
            try:
                parte_manual(qa, L)
            except Exception as e:                                   # noqa: BLE001
                L['ERRO_manual'] = f'{type(e).__name__}: {e} | {traceback.format_exc()[-1500:]}'
    finally:
        VIG.parar.set()
        qa.fechar()
        h1 = sha_full(src)
        L['fim'] = {'sha256_depois': h1[:16], 'producao_intacta': h1 == h0,
                    'hora': dt.datetime.now().isoformat(timespec='seconds')}
        try:
            os.remove(dst)
            L['copia_apagada'] = True
        except Exception as e:                                       # noqa: BLE001
            L['copia_apagada'] = f'nao: {e}'


if __name__ == '__main__':
    main()
