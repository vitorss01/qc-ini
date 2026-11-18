# -*- coding: utf-8 -*-
"""valida_r7_adr097.py -- Tester/QA, ciclo 2026-11-16_reteste07, modo VALIDAR do ADR-097 (entrega 44).

Escrito pelo Tester/QA. NAO reaproveita reteste07_usuario_f1*.py (Usuario Simulado) nem a suite qa_adr097. Da
infraestrutura usa so qa_final.QA (abre a copia, revela as abas, imprime EXCEL_PID), repro_comum (meu: senha, REPO) e
qa_inativar_ux.criar_usuario/fazer_login (usuario so na copia).

Ocorrencias (criterios do Orquestrador):
  OC-0078  carimbo da linha manual: (1) Delete de ID..MOTIVO (uma linha e duas areas), (2) Delete de DATA..MOTIVO
           deixando o ID, (3) recortar (origem e destino), (4) Ctrl+D (FillDown) a partir de uma linha carimbada,
           (5) edicao na propria linha, (6) SUBSTITUI_ID (Delete com o SUBSTITUI_ID fora da selecao; linha so com o
           SUBSTITUI_ID + MOTIVO/EQUIPAMENTO; substituicao valida); apagar so o ID; orfao legado limpo no ATUALIZAR.
           Reaproveitar a linha apagada por OUTRO usuario, com DATA/HORA de hoje POSTERIOR ao carimbo antigo (o caso
           que era recusado). ATUALIZAR (VBA) + reserva PQ; n/media do Painel e da Estatistica x oraculo proprio; LJ;
           nenhum manual antigo perde o carimbo; inativar um manual (fluxo vizinho).
  OC-0079  GUIA §3 (lido fora daqui) + travas das colunas (ID..MOTIVO livres; USUARIO..DETALHE travadas) + colar
           ID..MOTIVO sem erro; texto do aviso da colagem.
  OC-0080  DETALHE do ID fora do padrao (linhas reais + formula avaliada numa aba de rascunho para muitos IDs).
  OC-0081  colar e trocar so o RESULTADO -> CONFLITO_MANUAL; A01 novo (VBA e PQ); ResumoQA conta; corrigir a HORA.
  OC-0082  filtros do Registros (G4/J4 e os outros) com lista vazia -> "(nenhum registrado)"; texto livre recusado;
           escolher o item nao filtra; com dado, a lista e a de antes (oraculo: valores distintos da coluna).
  OC-0083  filtro de lote B4 em texto: lote digitado so com numeros aceito e filtra (contagem x oraculo); fora da
           lista recusado.
  OC-0048  (Hema) D3 >= B3 com mensagem; A5 vermelho nos casos do ADR-097 §6; grade x oraculo por ano.

Uso: python -u valida_r7_adr097.py <Bioquimica|Hematologia> [--sem-pq]
Regras: so COPIA nova em %TEMP%/qcw/auditoria/2026-11-16_reteste07/tester_validar/; nada e salvo; fechar() encerra so
o PID que o QA criou; nenhuma captura de tela. Dados de teste so na copia: usuarios TESTE_AUDITORIA_R7VA e _R7VB
(ANALISTA), manuais de 14/11/2026 (madrugada, horarios a mais de 40 min de qualquer real do analito) e de hoje (depois
do carimbo antigo), valores sinteticos de 3 decimais; IDs invalidos TESTE1, MAN_1,5, MAN_12%, MAN_1E3, MAN_-12.
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
SEM_PQ = '--sem-pq' in sys.argv
BIO = PROD == 'Bioquimica'
NLV = 2 if BIO else 3
CICLO = '2026-11-16_reteste07'
PASTA = os.path.join(os.environ['TEMP'], 'qcw', 'auditoria', CICLO, 'tester_validar')
UA, UB = 'TESTE_AUDITORIA_R7VA', 'TESTE_AUDITORIA_R7VB'
EQ = 'DIMENSION 1' if BIO else 'XN1000'
LOTE_PREF = '8977' if BIO else '625311'
EVITAR = {'GLICOSE', 'ALT (TGP)', 'MAGNÉSIO', 'MAGNESIO', 'GLUC'} if BIO else {'WBC', 'BASO#', 'BASO%', 'EO#'}
DIA_A = dt.datetime(2026, 11, 14)
EPOCA = dt.datetime(1899, 12, 30)
SEGREDOS = [x for x in (os.environ.get('USERNAME'), os.environ.get('COMPUTERNAME')) if x and len(x) > 3]
AVISO_097 = ('Mude o que for diferente antes de clicar em ATUALIZAR DADOS. Atenção: com a mesma data, hora, nível e '
             'analito de outra linha, ela fica como repetida (CONFLITO_MANUAL), mesmo com outro resultado. Se for uma '
             'nova dosagem, use a hora real dela.')
DET_INV = 'ID fora do padrão: apague-o (um novo é gerado)'


def t(v):
    if isinstance(v, float) and v == int(v):
        return str(int(v))
    return '' if v is None else str(v).strip()


def up(v):
    return t(v).upper()


def num(v):
    return isinstance(v, (int, float)) and not isinstance(v, bool)


def rel(a, b):
    if not num(a) or not num(b):
        return None
    return abs(a - b) / max(abs(a), abs(b), 1e-300)


def dec(x):
    return f'{x:.3f}'.replace('.', ',')


def serial(d):
    return (d - EPOCA).days


def sha_full(p):
    h = hashlib.sha256()
    with open(p, 'rb') as f:
        for b in iter(lambda: f.read(1 << 20), b''):
            h.update(b)
    return h.hexdigest()


def esp(s):
    return re.sub(r'\s+', ' ', t(s))


class Log:
    def __init__(self, saida):
        self.saida, self.R = saida, {'inicio': dt.datetime.now().isoformat(timespec='seconds')}

    def __setitem__(self, k, v):
        self.R[k] = v
        print('##', k, json.dumps(v, ensure_ascii=False, default=str)[:1800], flush=True)
        self.gravar()

    def __getitem__(self, k):
        return self.R[k]

    def get(self, k, d=None):
        return self.R.get(k, d)

    def gravar(self):
        s = json.dumps(self.R, ensure_ascii=False, indent=1, default=str)
        for x in SEGREDOS:
            s = re.sub(re.escape(x), '<omitido>', s, flags=re.I)
        open(self.saida, 'w', encoding='utf-8').write(s)


class Vigia(threading.Thread):
    """Le o TEXTO e fecha caixas (#32770) do PID desta sessao. Nao captura imagem."""

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
            time.sleep(0.5 if achou else 0.2)


# =============================================================================================== tabelas
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


CAMPOS_F = ('ID_REGISTRO', 'STATUS_ANALITICO', 'PARTICIPA_ESTATISTICA', 'MOTIVO_EXCLUSAO_AUTOMATICA', 'RUN', 'NIVEL',
            'RESULTADO', 'ANALITO', 'LOTE', 'USUARIO_MANUAL', 'ID_RELACIONADO', 'DATA', 'HORA', 'TIPO_PLOTAGEM_LJ')


def meus_final(rows, ids):
    ids = {x for x in ids if x}
    return {t(r.get('ID_REGISTRO')): {k: r.get(k) for k in CAMPOS_F if k in r}
            for r in rows if base_id(r.get('ID_REGISTRO')) in ids or t(r.get('ID_REGISTRO')) in ids}


def qa_meus(qa, ids):
    cab, rows = tabela(qa, 'tblQA_Integracao')
    out = []
    for r in rows:
        s = ' | '.join(t(v) for v in r.values())
        if any(i and re.search(re.escape(i) + r'(?!\d)', s) for i in ids):
            out.append({k: t(v) for k, v in r.items() if t(v)})
    return out


def oraculo(rows, analito, nivel, lote=None, ini=None, fim=None, equip=EQ, anos=None):
    v, ids = [], []
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
        if anos is not None:
            if not num(d) or (EPOCA + dt.timedelta(days=int(d))).year not in anos:
                continue
        v.append(r['RESULTADO'])
        ids.append(t(r.get('ID_REGISTRO')))
    o = {'n': len(v), 'ids': ids}
    if v:
        o['media'] = stt.fmean(v)
    if len(v) > 1:
        o['dp'] = stt.stdev(v)
    return o


def minutos_reais(rows, analito, dia):
    """Minutos do dia (serial) dos resultados do interfaceamento do analito no EQ (qualquer nivel)."""
    out = []
    for r in rows:
        if up(r.get('ANALITO')) != analito.upper() or up(r.get('EQUIPAMENTO')) != EQ.upper():
            continue
        if not up(r.get('ORIGEM_RESULTADO')).startswith('INTERF'):
            continue
        dh = r.get('DATA_HORA')
        if num(dh) and int(dh) == dia:
            out.append(round((dh - int(dh)) * 1440))
    return sorted(out)


def slots(reais, ini, fim, passo, n, folga=40):
    out, m = [], ini
    while m <= fim and len(out) < n:
        if all(abs(m - x) > folga for x in reais):
            out.append(m)
        m += passo
    return out


def hhmm(m):
    return f'{m // 60:02d}:{m % 60:02d}'


# =============================================================================================== OC-0080 formula
IDS_FORMULA = {
    'MAN_0012': 'valido', ' man_7 ': 'valido', 'man_00012': 'valido', 'MAN_12 ': 'valido', 'MAN_0': 'valido',
    'MAN_1,5': 'invalido', 'MAN_12%': 'invalido', 'MAN_1E3': 'invalido', 'MAN_-12': 'invalido', 'TESTE1': 'invalido',
    'MAN_': 'invalido', 'MAN_12A': 'invalido', 'MAN_1.5': 'invalido', 'MAN_+12': 'invalido', 'MAN_ 12': 'invalido',
    'HEM-12345': 'invalido', 'BIO-12345': 'invalido', '12345': 'invalido', 'MAN12': 'invalido', 'MAN_(12)': 'invalido',
    'MAN_$12': 'invalido', 'MAN_1/2': 'invalido', 'MAN_12/11': 'invalido', 'MAN_1e3': 'invalido', 'MAN_0x1A': 'invalido',
}


def parte_detalhe_formula(qa, L):
    lo = qa.lo('tblResultados_Manuais')
    f = t(lo.ListColumns('DETALHE').DataBodyRange.Cells(1, 1).Formula)
    out = {'formula_ini': f[:160], 'formula_fim': f[-420:], 'tem_texto_novo': DET_INV in f}
    try:
        sys.path.insert(0, os.path.join(C.REPO, '_arquitetura', 'integracao_dbseac'))
        import camada_dados
        ref = camada_dados.FORMULA_DETALHE
        norm = lambda x: re.sub(r'tblResultados_Manuais\[@', '[@', re.sub(r'\[@\[([A-Z_]+)\]\]', r'[@\g<1>]', x))  # noqa: E731
        out['igual_camada_dados'] = norm(f) == norm(ref)
    except Exception as e:                                           # noqa: BLE001
        out['igual_camada_dados'] = f'erro {e}'
    ws = qa.wb.Worksheets.Add()
    try:
        res = {}
        ks = list(IDS_FORMULA)
        for k, s in enumerate(ks, 1):
            ws.Range(f'A{k}').NumberFormat = '@'
            ws.Range(f'A{k}').Value = s
            fk = re.sub(r'(tblResultados_Manuais)?\[@\[?ID_REGISTRO\]?\]', f'$A{k}', f)
            ws.Range(f'B{k}').Formula2 = fk
        qa.ex.xl.Calculate()
        for k, s in enumerate(ks, 1):
            v = t(ws.Range(f'B{k}').Value)
            esper = 'atualize os dados' if IDS_FORMULA[s] == 'valido' else DET_INV
            res[repr(s)] = {'tela': v, 'esperado': esper, 'ok': v == esper}
        out['ids'] = res
        out['ok'] = all(x['ok'] for x in res.values())
    finally:
        ev = qa.ex.xl.DisplayAlerts
        qa.ex.xl.DisplayAlerts = False
        ws.Delete()
        qa.ex.xl.DisplayAlerts = ev
    L['OC0080_formula'] = out


# =============================================================================================== OC-0082/0083 Registros
def itens_lista(cel):
    f1 = t(cel.Validation.Formula1)
    if f1.startswith('='):
        try:
            v = cel.Worksheet.Evaluate(f1[1:])
            v = v.Value if hasattr(v, 'Value') else v
        except Exception as e:                                       # noqa: BLE001
            return f1, f'erro {e}'
        if not isinstance(v, tuple):
            v = ((v,),)
        return f1, [t(x[0]) for x in v if t(x[0])]
    return f1, [x for x in f1.split(';')]


def aceita(xl, cel, valor, como='valor'):
    ev = xl.EnableEvents
    xl.EnableEvents = False
    try:
        if como == 'local':
            cel.FormulaLocal = valor
        else:
            cel.Value = valor
        return bool(cel.Validation.Value)
    except Exception as e:                                           # noqa: BLE001
        return f'erro {type(e).__name__}'
    finally:
        xl.EnableEvents = ev


def info_valid(cel):
    v = cel.Validation
    try:
        return {'Type': v.Type, 'AlertStyle': v.AlertStyle, 'ShowError': bool(v.ShowError),
                'ErrorTitle': t(v.ErrorTitle), 'ErrorMessage': t(v.ErrorMessage), 'Formula1': t(v.Formula1),
                'NumberFormat': t(cel.NumberFormat), 'Locked': bool(cel.Locked)}
    except Exception as e:                                           # noqa: BLE001
        return {'erro': str(e)}


def parte_registros(qa, L):
    xl = qa.ex.xl
    ws = qa.ws('Registros')
    lo = ws.ListObjects('tblRegistros')
    cab = [str(c.Name) for c in lo.ListColumns]
    out = {'colunas': cab, 'linhas': lo.ListRows.Count, 'protegida': bool(ws.ProtectContents)}
    try:
        ws.Unprotect(C.senha_planilha())        # so para escrever as provas e, depois, devolver o dado da copia
    except Exception:                                                # noqa: BLE001
        pass
    qa.run('mCalibracao.LimparFiltrosRegistros', teto=120)
    CEL = {'B4 lote': 2, 'C4 analito': 3, 'G4 equipamento': 7, 'J4 status': 10}
    COLT = {2: 13, 3: 3, 7: 12, 10: 10}        # coluna da tabela filtrada (C_LOTE=13, C_ANALITO=3, C_EQUIP=12, C_STATUS=10)
    vals = lo.DataBodyRange.Value2 if lo.DataBodyRange is not None else []

    def distintos(ci, extra=()):
        s = {t(r[ci - 1]) for r in vals if t(r[ci - 1]) and len(t(r[ci - 1])) <= 80}
        s |= {x for x in extra if x}
        return sorted(s, key=lambda x: x.lower())

    cad = []
    try:
        cad = [t(x[0]) for x in qa.wb.Names('regLoteCol').RefersToRange.Value2]
    except Exception:                                                # noqa: BLE001
        pass
    antes = {}
    for rot, c in CEL.items():
        cel = ws.Cells(4, c)
        f1, it = itens_lista(cel)
        antes[rot] = {'valid': info_valid(cel), 'itens': it,
                      'oraculo': distintos(COLT[c], cad if c == 2 else ())}
        antes[rot]['lista_igual_oraculo'] = it == antes[rot]['oraculo'] if it != ['(nenhum registrado)'] else \
            antes[rot]['oraculo'] == []
    out['antes'] = antes
    # ---------- OC-0082: forca a lista vazia de equipamento e status (copia), reconstroi, prova, devolve o dado
    nlin = lo.ListRows.Count
    guard = {}
    vazio = {}
    if nlin:
        ev = xl.EnableEvents
        xl.EnableEvents = False
        try:
            for c in (12, 10):
                rg = lo.ListColumns(c).DataBodyRange
                guard[c] = (rg.Formula, rg.Value2)
                rg.ClearContents()
        finally:
            xl.EnableEvents = ev
    qa.run('mCalibracao.AtualizarListasFiltro', teto=120)
    for rot in ('G4 equipamento', 'J4 status'):
        c = CEL[rot]
        cel = ws.Cells(4, c)
        f1, it = itens_lista(cel)
        r = {'itens': it, 'valid': info_valid(cel)}
        r['texto_livre'] = {x: aceita(xl, cel, x) for x in ('TESTE-AUDITORIA-R7V', 'DIMENSION 3', 'XN2000', 'ATIVO')}
        r['item_aceito'] = aceita(xl, cel, '(nenhum registrado)')
        ev = xl.EnableEvents
        xl.EnableEvents = True
        try:
            cel.Value = '(nenhum registrado)'                # escolhido na seta (com os eventos: RegistrosMudou)
            qa.ex.esperar()
            time.sleep(0.5)
            r['retorno_filtro'] = t(qa.run('mCalibracao.AplicarFiltrosRegistros', teto=120))
            r['K4'] = t(ws.Range('K4').Text)
        finally:
            cel.ClearContents()
            xl.EnableEvents = ev
        msg = r['valid'].get('ErrorMessage', '')
        r['ok'] = (it == ['(nenhum registrado)'] and all(v is False for v in r['texto_livre'].values())
                   and r['item_aceito'] is True and r['K4'].startswith('Sem filtro')
                   and msg.startswith('Escolha ') and 'na seta da célula.' in msg)
        vazio[rot] = r
    out['lista_vazia'] = vazio
    # devolve
    if guard:
        ev = xl.EnableEvents
        xl.EnableEvents = False
        try:
            for c, (fo, v2) in guard.items():
                lo.ListColumns(c).DataBodyRange.Formula = fo
        finally:
            xl.EnableEvents = ev
    qa.run('mCalibracao.AtualizarListasFiltro', teto=120)
    depois = {}
    for rot, c in CEL.items():
        f1, it = itens_lista(ws.Cells(4, c))
        depois[rot] = {'itens': it, 'igual_antes': it == antes[rot]['itens']}
    out['com_dado_de_novo'] = depois
    # ---------- OC-0083: lote digitado so com numeros
    cel = ws.Cells(4, 2)
    lotes = antes['B4 lote']['itens']
    lote_num = next((x for x in lotes if x.isdigit() and any(t(r[12]) == x for r in vals)), None) or \
        next((x for x in lotes if x.isdigit()), None)
    r = {'NumberFormat': t(cel.NumberFormat), 'lote_escolhido': lote_num}
    if lote_num:
        ev = xl.EnableEvents
        xl.EnableEvents = True
        try:
            cel.FormulaLocal = lote_num                      # digitado
            qa.ex.esperar()
            time.sleep(0.5)
            r['tipo_valor'] = type(cel.Value).__name__
            r['validacao_ok'] = bool(cel.Validation.Value)
            r['K4'] = t(ws.Range('K4').Text)
            r['retorno'] = t(qa.run('mCalibracao.AplicarFiltrosRegistros', teto=120))
            orac = sum(1 for x in vals if t(x[12]) == lote_num)
            r['oraculo_linhas_do_lote'] = orac
            vis = t(r['retorno']).split('|')[0]
            r['visiveis'] = vis
            r['ok_filtra'] = vis.isdigit() and int(vis) == orac
        finally:
            cel.ClearContents()
            xl.EnableEvents = ev
        r['fora_da_lista'] = {x: aceita(xl, cel, x, 'local') for x in ('99999', '8977-EE', '1')}
        ev = xl.EnableEvents
        xl.EnableEvents = False
        cel.ClearContents()
        xl.EnableEvents = ev
        r['ok'] = (r['NumberFormat'] == '@' and r['validacao_ok'] and r['tipo_valor'] == 'str' and r['ok_filtra']
                   and all(v is False for v in r['fora_da_lista'].values()))
    out['lote_digitado'] = r
    qa.run('mCalibracao.LimparFiltrosRegistros', teto=120)
    L['OC0082_0083_registros'] = out


# =============================================================================================== OC-0048 (Hema)
def grade(ws):
    out = {}
    for l in ws.Range('A14:F133').Value2:
        if t(l[0]) and num(l[2]):
            out[(up(l[0]), t(l[1]))] = {'n': l[2], 'media': l[3], 'dp': l[4]}
    return out


def parte_estat_hema(qa, L, rows):
    xl = qa.ex.xl
    e = qa.ws('Estatística')
    d3 = e.Range('D3')
    out = {'D3_valid': info_valid(d3), 'B3_valid': info_valid(e.Range('B3')),
           'A5_fonte': {'cor': e.Range('A5').Font.Color, 'negrito': bool(e.Range('A5').Font.Bold)},
           'inicial': {c: t(e.Range(c).Text) for c in ('B3', 'D3', 'B4', 'A5')}}
    lote = t(e.Range('B4').Value)
    anos = {}
    for r in rows:
        if t(r.get('LOTE')) == lote and up(r.get('EQUIPAMENTO')) == EQ.upper() and num(r.get('RESULTADO')) \
                and num(r.get('DATA')) and up(r.get('PARTICIPA_ESTATISTICA')) == 'SIM':
            a = (EPOCA + dt.timedelta(days=int(r['DATA']))).year
            anos[a] = anos.get(a, 0) + 1
    anos_todos = {}
    for r in rows:
        if t(r.get('LOTE')) == lote and num(r.get('RESULTADO')) and num(r.get('DATA')):
            a = (EPOCA + dt.timedelta(days=int(r['DATA']))).year
            anos_todos[a] = anos_todos.get(a, 0) + 1
    out['lote_B4'] = lote
    out['anos_do_lote_participa_EQ'] = anos
    out['anos_do_lote_todas_linhas'] = anos_todos

    def por_ano(b3, d3v):
        ev = xl.EnableEvents
        xl.EnableEvents = True
        try:
            for c, v in (('B3', b3), ('D3', d3v)):
                e.Range(c).Value = v
                qa.ex.esperar()
                time.sleep(0.6)
            xl.Calculate()
            qa.ex.esperar()
        finally:
            xl.EnableEvents = ev
        g = grade(e)
        a_ef = set(range(b3, max(b3, d3v) + 1))
        cmp_ = {}
        piores = 0.0
        for (an, nv), x in g.items():
            o = oraculo(rows, an, nv, lote=lote, anos=a_ef)
            rm = rel(x['media'], o.get('media')) if x['n'] else None
            cmp_[f'{an}|{nv}'] = (x['n'], o['n'], rm)
            if x['n'] != o['n']:
                piores = max(piores, 1.0)
            elif rm is not None:
                piores = max(piores, rm)
        return {'B3': t(e.Range('B3').Text), 'D3': t(e.Range('D3').Text), 'A5': t(e.Range('A5').Value),
                'A5_cor': e.Range('A5').Font.Color, 'grade_hash': hashlib.sha1(
                    json.dumps(e.Range('A14:AU133').Value2, default=str).encode()).hexdigest()[:16],
                'linhas_grade': len(g), 'max_rel_ou_n_diferente': piores,
                'amostra': dict(list(cmp_.items())[:6]), 'n_diferentes': [k for k, v in cmp_.items() if v[0] != v[1]]}

    casos = {}
    ano_lote = sorted(anos)
    casos['2025_2025'] = por_ano(2025, 2025)
    casos['2026_2026'] = por_ano(2026, 2026)
    casos['2025_2026'] = por_ano(2025, 2026)
    casos['2026_2025_B3maior'] = por_ano(2026, 2025)
    casos['2024_2024'] = por_ano(2024, 2024)
    out['casos'] = casos
    out['grade_B3maior_igual_2026_2026'] = casos['2026_2025_B3maior']['grade_hash'] == casos['2026_2026']['grade_hash']
    # D3 < B3 recusado pela validacao (COM nao dispara a caixa: Validation.Value)
    e.Range('B3').Value = 2026
    out['D3_recusa'] = {str(v): aceita(xl, d3, v) for v in (2025, 2024)}
    out['D3_aceita'] = {str(v): aceita(xl, d3, v) for v in (2026, 2027)}
    ev = xl.EnableEvents
    xl.EnableEvents = True
    try:
        e.Range('B3').Value = 2026
        e.Range('D3').Value = 2026
        qa.ex.esperar()
    finally:
        xl.EnableEvents = ev
    out['ano_lote'] = ano_lote
    L['OC0048_estatistica'] = out


# =============================================================================================== OC-0078..0081 manual
class Manual:
    def __init__(self, qa, vig):
        self.qa, self.vig = qa, vig
        self.lo = qa.lo('tblResultados_Manuais')
        self.ws = self.lo.Parent
        self.cab = [str(c.Name) for c in self.lo.ListColumns]
        self.k = 0

    def cel(self, r, col):
        return self.lo.ListColumns(col).DataBodyRange.Cells(r, 1)

    def faixa(self, r1, r2=None, c1='ID_REGISTRO', c2='MOTIVO'):
        return self.ws.Range(self.cel(r1, c1), self.cel(r2 or r1, c2))

    def linha(self, r):
        rg = self.lo.ListRows(r).Range
        v = rg.Value2[0]
        d = {'_planilha': rg.Row}
        for i, c in enumerate(self.cab):
            if c in ('STATUS', 'DETALHE', 'RUN', 'DATA', 'HORA'):
                d[c] = t(rg.Cells(1, i + 1).Text)
            elif c == 'REGISTRADO_EM':
                d[c] = v[i]
                d['REGISTRADO_EM_txt'] = t(rg.Cells(1, i + 1).Text)
            else:
                d[c] = v[i] if num(v[i]) else t(v[i])
        return d

    def carimbo(self, r):
        return {'USUARIO': t(self.cel(r, 'USUARIO').Value), 'REGISTRADO_EM': self.cel(r, 'REGISTRADO_EM').Value2}

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

    def todas(self):
        """Carimbo de todas as linhas com ID ou carimbo (para 'nenhum manual antigo perde o carimbo')."""
        n = self.lo.ListRows.Count
        idv = self.lo.ListColumns('ID_REGISTRO').DataBodyRange.Value2
        us = self.lo.ListColumns('USUARIO').DataBodyRange.Value2
        rg = self.lo.ListColumns('REGISTRADO_EM').DataBodyRange.Value2
        out = {}
        for i in range(n):
            if t(idv[i][0]) or t(us[i][0]) or t(rg[i][0]):
                out[i + 1] = (t(idv[i][0]), t(us[i][0]), rg[i][0])
        return out


def campos(dia, hora, lote, nv, an, val, equip=True):
    c = [('DATA', dia.strftime('%d/%m/%Y')), ('HORA', hora)]
    if equip and BIO:
        c.append(('EQUIPAMENTO', EQ))
    return c + [('LOTE', "'" + lote), ('NIVEL', str(nv)), ('ANALITO', an), ('RESULTADO', dec(val)), ('MOTIVO', 'Outro')]


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
    return {'filtro': {c: t(p.Range(c).Text) for c in ('L4', 'C3', 'H3', 'K3', 'K4')},
            'engPainel': [{'nivel': l[0], 'n': l[1], 'media': l[2], 'dp': l[3]} for l in eng]}


def eng_slots(qa, ids):
    e = qa.ws('Eng_Saida')
    n = int(e.Range('I1').Value or 0)
    if not n:
        return {'n': 0}
    blk = e.Range(e.Cells(3, 1), e.Cells(2 + n, 80)).Value
    out = []
    for k, r in enumerate(blk):
        idc = [t(x) for x in r[38:38 + NLV]]
        if any(base_id(i) in ids for i in idc if i):
            out.append({'slot': k + 1, 'run': t(r[1]), 'ids': idc, 'val': list(r[24:24 + NLV])})
    return {'n': n, 'C1': t(e.Range('C1').Value), 'E1': t(e.Range('E1').Value), 'meus': out}


def audit(qa, r0, ids):
    ws = qa.ws('Audit_Log')
    ur = ws.UsedRange
    ult = ur.Row + ur.Rows.Count - 1
    nc = ur.Column + ur.Columns.Count - 1
    ini = max(1, ult - 1500)               # o End(xlUp) da coluna A para no ListObject: varre o fim da area usada
    v = ws.Range(ws.Cells(ini, 1), ws.Cells(ult, nc)).Value
    out = []
    for l in v:
        s = ' | '.join(t(x) for x in l)
        if any(i and re.search(re.escape(i) + r'(?!\d)', s) for i in ids):
            out.append(s[:400])
    return out, ult


def candidatos(rows, nlv, mes_ini, evitar, lote_pref):
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
    return [k for k, _ in sorted(ok, key=lambda x: (x[0][1] != lote_pref, -x[1]))]


def estat(qa, an, lote, rows):
    e = qa.ws('Estatística')
    if BIO:
        lote_e = t(e.Range('H4').Value)
        h5, j5 = e.Range('H5').Value2, e.Range('J5').Value2
        g = grade(e)
        o = {nv: oraculo(rows, an, nv, lote=lote_e, ini=int(h5) if num(h5) and h5 else None,
                         fim=int(j5) if num(j5) and j5 else None) for nv in (1, 2)}
        return {'H4': lote_e, 'H5': t(e.Range('H5').Text), 'J5': t(e.Range('J5').Text),
                'grade': {nv: g.get((an.upper(), str(nv))) for nv in (1, 2)},
                'oraculo_n': {nv: o[nv]['n'] for nv in (1, 2)},
                'rel_media': {nv: rel((g.get((an.upper(), str(nv))) or {}).get('media'), o[nv].get('media')) for nv in (1, 2)},
                'n_igual': {nv: (g.get((an.upper(), str(nv))) or {}).get('n') == o[nv]['n'] for nv in (1, 2)}}
    xl = qa.ex.xl
    ev = xl.EnableEvents
    xl.EnableEvents = True
    try:
        for c in ('B3', 'D3'):
            e.Range(c).Value = 2026
            qa.ex.esperar()
        xl.Calculate()
    finally:
        xl.EnableEvents = ev
    lote_e = t(e.Range('B4').Value) or lote
    g = grade(e)
    o = {nv: oraculo(rows, an, nv, lote=lote_e, anos={2026}) for nv in range(1, NLV + 1)}
    return {'B4': lote_e, 'B3': t(e.Range('B3').Text), 'D3': t(e.Range('D3').Text), 'A5': t(e.Range('A5').Value),
            'grade': {nv: g.get((an.upper(), str(nv))) for nv in range(1, NLV + 1)},
            'oraculo_n': {nv: o[nv]['n'] for nv in range(1, NLV + 1)},
            'rel_media': {nv: rel((g.get((an.upper(), str(nv))) or {}).get('media'), o[nv].get('media')) for nv in range(1, NLV + 1)},
            'n_igual': {nv: (g.get((an.upper(), str(nv))) or {}).get('n') == o[nv]['n'] for nv in range(1, NLV + 1)}}


def parte_manual(qa, L, rows0):
    xl = qa.ex.xl
    nov1, nov30 = serial(dt.datetime(2026, 11, 1)), serial(dt.datetime(2026, 11, 30))
    hoje = dt.datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
    cand = candidatos(rows0, NLV, nov1, EVITAR, LOTE_PREF)
    AN = LOTE = None
    agora_m = dt.datetime.now().hour * 60 + dt.datetime.now().minute
    escolha = []
    for an, lote in cand[:25]:
        rh = minutos_reais(rows0, an, serial(hoje))
        perto = [x for x in rh if x > agora_m - 50]
        escolha.append((an, lote, len(rh), perto))
        if lote == LOTE_PREF and not perto:
            AN, LOTE = an, lote
            break
    if AN is None:
        AN, LOTE = cand[0]
    REAIS_A = minutos_reais(rows0, AN, serial(DIA_A))
    SA = slots(REAIS_A, 60, 23 * 60, 5, 40)
    if len(SA) < 30:
        raise RuntimeError(f'poucos horarios livres em 14/11 para {AN}: {len(SA)}')
    rec = {}
    for nv in range(1, NLV + 1):
        v = [r['RESULTADO'] for r in sorted((r for r in rows0 if up(r.get('ANALITO')) == AN.upper()
                                             and t(r.get('LOTE')) == LOTE and t(r.get('NIVEL')) == str(nv)
                                             and up(r.get('PARTICIPA_ESTATISTICA')) == 'SIM' and num(r.get('RESULTADO'))
                                             and num(r.get('DATA_HORA'))), key=lambda r: r['DATA_HORA'])][-20:]
        rec[nv] = stt.fmean(v)
    L['ref'] = {'analito': AN, 'lote': LOTE, 'media_recente': rec, 'reais_14_11_min': REAIS_A,
                'escolha': escolha[:10], 'slots_14_11': [hhmm(x) for x in SA[:30]]}
    k = [0]

    def prox():
        h = hhmm(SA[k[0]])
        k[0] += 1
        return h

    def val(nv, d):
        return round(rec[nv] + d, 3)

    ids_todos = set()
    criar_usuario(qa, UA, 'ANALISTA')
    criar_usuario(qa, UB, 'ANALISTA')
    L['login_A'] = fazer_login(qa, UA)
    painel0 = painel_filtro(qa, AN, LOTE)
    L['painel_0'] = painel0
    qa.run('mIntegracao.IrDigitarResultados', teto=120)
    M = Manual(qa, VIG)
    f = M.livre()
    L['tabela_manual'] = {'colunas': M.cab, 'primeira_livre': f, 'planilha': M.cel(f, 'ID_REGISTRO').Row,
                          'protegida': bool(M.ws.ProtectContents), 'ProtectionMode': bool(M.ws.ProtectionMode),
                          'travas': {c: bool(M.cel(f, c).Locked) for c in M.cab}, 'linhas': M.lo.ListRows.Count,
                          'livres_30': all(not M.id(f + i) and not t(M.cel(f + i, 'DATA').Value) for i in range(30))}
    # mapa de linhas (relativas a f)
    R = {'r1': f, 'r2': f + 1, 'r3': f + 2, 'r8': f + 3, 'r9': f + 4, 'r4': f + 6, 'r5': f + 10, 'r6': f + 11,
         'r7': f + 12, 'rP': f + 13, 'cola': f + 14, 'inv1': f + 15, 'inv2': f + 16, 'inv3': f + 17, 'inv4': f + 18,
         'inv5': f + 19, 'val1': f + 20, 'val2': f + 21, 's1': f + 22, 's2': f + 23, 's3': f + 24, 'dest': f + 26,
         'orf': f + 28}
    L['linhas'] = R
    passos = {}
    xl.Interactive = True
    try:
        # ------------------------------------------------ fase A: usuario A digita os manuais antigos (14/11)
        A = {'r1': (1, 0.002), 'r2': (2, -0.002), 'r3': (1, 0.003), 'r8': (2, 0.001), 'r9': (1, -0.001),
             'r4': (1, 0.004), 'r5': (2, 0.003), 'r6': (1, -0.003), 'r7': (2, -0.001), 'rP': (1, 0.001)}
        if NLV == 3:
            A['r9'] = (3, 0.002)
            A['r7'] = (3, -0.002)
        dA = {}
        for nome, (nv, d) in A.items():
            r = R[nome]
            h = prox()
            M.digitar(r, campos(DIA_A, h, LOTE, nv, AN, val(nv, d)))
            dA[nome] = {'hora': h, 'nivel': nv, 'valor': val(nv, d), 'linha': M.linha(r)}
        passos['A_digitados'] = {k2: {'id': v['linha']['ID_REGISTRO'], 'USUARIO': v['linha']['USUARIO'],
                                      'REG': v['linha']['REGISTRADO_EM_txt'], 'STATUS': v['linha']['STATUS']}
                                 for k2, v in dA.items()}
        passos['A_avisos'] = M.avisos()
        # OC-0080: IDs fora do padrao e validos, digitados no ID e depois os dados
        invs = {'inv1': 'TESTE1', 'inv2': 'MAN_1,5', 'inv3': 'MAN_12%', 'inv4': 'MAN_1E3', 'inv5': 'MAN_-12',
                'val1': 'MAN_0012', 'val2': ' man_7 '}
        det0 = {}
        for nome, idd in invs.items():
            r = R[nome]
            M.esc(r, 'ID_REGISTRO', idd)
            nv = 1 if nome in ('inv1', 'inv3', 'inv5', 'val1') else 2
            h = prox()
            M.digitar(r, campos(DIA_A, h, LOTE, nv, AN, val(nv, 0.0005 * (1 + len(det0)))))
            l_ = M.linha(r)
            det0[nome] = {'digitado': idd, 'id': l_['ID_REGISTRO'], 'DETALHE': l_['DETALHE'], 'STATUS': l_['STATUS'],
                          'hora': h, 'nivel': nv}
        passos['OC0080_antes_atualizar'] = det0
        passos['OC0080_avisos'] = M.avisos()
    finally:
        xl.Interactive = False
    ids_A = {v['linha']['ID_REGISTRO'] for v in dA.values()}
    ids_todos |= ids_A | {v['id'] for v in det0.values()}
    carimbo_A = {k2: M.carimbo(R[k2]) for k2 in A}
    L['passos_A'] = passos
    al_r0 = qa.ws('Audit_Log').Cells(qa.ws('Audit_Log').Rows.Count, 1).End(-4162).Row

    msg1, s1 = qa.atualizar()
    rows1 = final(qa)
    L['atualizar1'] = {'segundos': round(s1, 1), 'mensagem': msg1[:1500],
                       'final': meus_final(rows1, ids_todos),
                       'tela': {k2: {c: M.linha(R[k2])[c] for c in ('ID_REGISTRO', 'STATUS', 'DETALHE', 'USUARIO')}
                                for k2 in list(A) + list(invs)},
                       'carimbos_A_iguais': {k2: M.carimbo(R[k2]) == carimbo_A[k2] for k2 in A}}
    snap1 = M.todas()

    # ------------------------------------------------ fase B: usuario B
    L['login_B'] = fazer_login(qa, UB)
    qa.run('mIntegracao.IrDigitarResultados', teto=120)
    xl.Interactive = True
    PB = {}
    try:
        M.avisos()
        # (5) edicao na propria linha (r5): o carimbo e o ID ficam
        c0 = M.carimbo(R['r5'])
        novo_v5 = round(dA['r5']['valor'] + 0.007, 3)
        M.esc(R['r5'], 'RESULTADO', dec(novo_v5))
        PB['5_editar'] = {'id_antes': dA['r5']['linha']['ID_REGISTRO'], 'id_depois': M.id(R['r5']),
                          'carimbo_antes': c0, 'carimbo_depois': M.carimbo(R['r5']), 'avisos': M.avisos(),
                          'novo_valor': novo_v5}
        # (4) Ctrl+D a partir de r4 (carimbo de A) para 2 linhas vazias
        c0 = M.carimbo(R['r4'])
        t_fd = time.time()
        M.faixa(R['r4'], R['r4'] + 2).FillDown()
        qa.ex.esperar()
        av = M.avisos()
        fd = {'r4': {'id': M.id(R['r4']), 'carimbo': M.carimbo(R['r4'])},
              'r4+1': {'id': M.id(R['r4'] + 1), 'carimbo': M.carimbo(R['r4'] + 1)},
              'r4+2': {'id': M.id(R['r4'] + 2), 'carimbo': M.carimbo(R['r4'] + 2)}}
        hfd = [prox(), prox()]
        M.esc(R['r4'] + 1, 'HORA', hfd[0])
        M.esc(R['r4'] + 2, 'HORA', hfd[1])
        M.esc(R['r4'] + 2, 'RESULTADO', dec(val(1, -0.004)))
        PB['4_filldown'] = {'carimbo_r4_antes': c0, 'apos': fd, 'avisos': av + M.avisos(), 'horas_novas': hfd,
                            'ids_finais': [M.id(R['r4']), M.id(R['r4'] + 1), M.id(R['r4'] + 2)],
                            'carimbos_finais': [M.carimbo(R['r4'] + i) for i in range(3)]}
        # OC-0081/0079: copiar ID..MOTIVO de rP (A) para 'cola' e trocar SO o RESULTADO
        M.faixa(R['rP']).Copy(M.faixa(R['cola']))
        xl.CutCopyMode = False
        qa.ex.esperar()
        id_col = M.id(R['cola'])
        av = M.avisos()
        M.esc(R['cola'], 'RESULTADO', dec(val(1, 0.006)))
        PB['81_colar_trocar_resultado'] = {'id_origem': M.id(R['rP']), 'id_colado': id_col, 'avisos': av + M.avisos(),
                                           'carimbo_colado': M.carimbo(R['cola']),
                                           'carimbo_origem': M.carimbo(R['rP']),
                                           'carimbo_origem_igual': M.carimbo(R['rP']) == carimbo_A['rP']}
        # (6) SUBSTITUI_ID
        reais = sorted((r for r in rows1 if up(r.get('ANALITO')) == AN.upper() and t(r.get('LOTE')) == LOTE
                        and up(r.get('EQUIPAMENTO')) == EQ.upper() and up(r.get('ORIGEM_RESULTADO')).startswith('INTERF')
                        and up(r.get('STATUS_ANALITICO')) == 'ATIVO' and num(r.get('DATA')) and nov1 <= r['DATA'] <= nov30
                        and num(r.get('RESULTADO'))), key=lambda r: r['DATA_HORA'])
        s_n1 = next(r for r in reais if t(r['NIVEL']) == '1')
        s_n2 = next(r for r in reais if t(r['NIVEL']) == '2')
        numero = lambda i: re.sub(r'^[A-Za-z]+-', '', t(i))  # noqa: E731
        # s1: substituicao valida digitada so com o numero
        M.esc(R['s1'], 'SUBSTITUI_ID', numero(s_n1['ID_REGISTRO']))
        l_pre = M.linha(R['s1'])
        v_s1 = round(s_n1['RESULTADO'] + 0.111, 3)
        M.esc(R['s1'], 'RESULTADO', dec(v_s1))
        M.esc(R['s1'], 'MOTIVO', 'Outro')
        PB['6_subst_valida'] = {'real': {k2: s_n1.get(k2) for k2 in ('ID_REGISTRO', 'NIVEL', 'RESULTADO', 'RUN')},
                                'pre_preenchida': {c: l_pre[c] for c in ('ID_REGISTRO', 'DATA', 'HORA', 'LOTE', 'NIVEL',
                                                                         'ANALITO', 'SUBSTITUI_ID')},
                                'linha': M.linha(R['s1']), 'avisos': M.avisos(), 'valor': v_s1}
        # s2: SUBSTITUI_ID + dado, depois Delete de ID..MOTIVO (o SUBSTITUI_ID fica fora da selecao)
        M.esc(R['s2'], 'SUBSTITUI_ID', numero(s_n2['ID_REGISTRO']))
        M.esc(R['s2'], 'RESULTADO', dec(round(s_n2['RESULTADO'] + 0.222, 3)))
        M.esc(R['s2'], 'MOTIVO', 'Outro')
        antes_s2 = M.linha(R['s2'])
        M.faixa(R['s2']).ClearContents()
        qa.ex.esperar()
        PB['6_delete_com_subst_fora'] = {'antes': {c: antes_s2[c] for c in ('ID_REGISTRO', 'SUBSTITUI_ID', 'USUARIO')},
                                         'depois': M.linha(R['s2']), 'avisos': M.avisos(),
                                         'real': t(s_n2['ID_REGISTRO'])}
        # s3: so o SUBSTITUI_ID (inexistente) e depois MOTIVO/EQUIPAMENTO: nada some
        M.esc(R['s3'], 'SUBSTITUI_ID', '999999987')
        a = M.linha(R['s3'])
        M.esc(R['s3'], 'MOTIVO', 'Outro')
        b = M.linha(R['s3'])
        if BIO:
            M.esc(R['s3'], 'EQUIPAMENTO', EQ)
        c_ = M.linha(R['s3'])
        # o usuario desiste: apaga ID..MOTIVO e o SUBSTITUI_ID (USUARIO..DETALHE sao travadas: duas areas)
        xl.Union(M.faixa(R['s3']), M.cel(R['s3'], 'SUBSTITUI_ID')).ClearContents()
        qa.ex.esperar()
        PB['6_so_subst'] = {'apos_subst': {k2: a[k2] for k2 in ('ID_REGISTRO', 'SUBSTITUI_ID', 'USUARIO', 'DATA')},
                            'apos_motivo': {k2: b[k2] for k2 in ('ID_REGISTRO', 'SUBSTITUI_ID', 'USUARIO', 'MOTIVO')},
                            'apos_equip': {k2: c_[k2] for k2 in ('ID_REGISTRO', 'SUBSTITUI_ID', 'USUARIO', 'EQUIPAMENTO')},
                            'apos_apagar': M.linha(R['s3']), 'avisos': M.avisos()}
        # apagar SO o ID de uma linha com dado (r6): ID novo, carimbo fica
        c0 = M.carimbo(R['r6'])
        id0 = M.id(R['r6'])
        M.cel(R['r6'], 'ID_REGISTRO').ClearContents()
        qa.ex.esperar()
        PB['so_o_id'] = {'id_antes': id0, 'id_depois': M.id(R['r6']), 'carimbo_antes': c0,
                         'carimbo_depois': M.carimbo(R['r6']), 'avisos': M.avisos()}
        # inativar o manual r7 (fluxo vizinho, aba Inativar)
        id7 = M.id(R['r7'])
        xl.Interactive = False
        rI = qa.escrever_linha('tblInativacao_NaoConformes', {'ID_REGISTRO': id7, 'MOTIVO': 'TESTE-AUDITORIA R7V'})
        qa.ex.esperar()
        loI = qa.lo('tblInativacao_NaoConformes')
        PB['inativar'] = {'id': id7, 'linha': rI, 'valores': [t(x) for x in loI.ListRows(rI).Range.Value[0]]}
        xl.Interactive = True
        qa.run('mIntegracao.IrDigitarResultados', teto=120)
        M.avisos()
        # orfao legado: carimbo numa linha vazia gravado com os eventos desligados
        xl.EnableEvents = False
        M.cel(R['orf'], 'USUARIO').Value = UA
        M.cel(R['orf'], 'REGISTRADO_EM').Value = dt.datetime.now() - dt.timedelta(minutes=30)
        xl.EnableEvents = True
        PB['orfao_gravado'] = M.carimbo(R['orf'])
        # (1) Delete de ID..MOTIVO de r1
        c1 = M.carimbo(R['r1'])
        M.faixa(R['r1']).ClearContents()
        qa.ex.esperar()
        PB['1_delete_id_motivo'] = {'id_antes': dA['r1']['linha']['ID_REGISTRO'], 'carimbo_antes': c1,
                                    'depois': M.linha(R['r1']), 'avisos': M.avisos()}
        # (1b) Delete de duas areas de uma vez (r8 e r9)
        uni = xl.Union(M.faixa(R['r8']), M.faixa(R['r9']))
        uni.ClearContents()
        qa.ex.esperar()
        PB['1b_delete_duas_areas'] = {'r8': M.linha(R['r8']), 'r9': M.linha(R['r9']), 'avisos': M.avisos()}
        # (2) Delete de DATA..MOTIVO de r2 (o ID fica fora da selecao)
        c2 = M.carimbo(R['r2'])
        M.faixa(R['r2'], c1='DATA').ClearContents()
        qa.ex.esperar()
        PB['2_delete_data_motivo'] = {'id_antes': dA['r2']['linha']['ID_REGISTRO'], 'carimbo_antes': c2,
                                      'depois': M.linha(R['r2']), 'avisos': M.avisos()}
        # (3) recortar r3 (manual de A, ja integrado) para 'dest'
        c3 = M.carimbo(R['r3'])
        M.faixa(R['r3']).Cut(M.faixa(R['dest']))
        qa.ex.esperar()
        PB['3_recortar'] = {'id_antes': dA['r3']['linha']['ID_REGISTRO'], 'carimbo_antes': c3,
                            'origem': M.linha(R['r3']), 'destino': M.linha(R['dest']), 'avisos': M.avisos()}
        # ---- reaproveitar as linhas apagadas, por B, HOJE, com hora POSTERIOR ao carimbo antigo de A
        carimbo_min = max(x['REGISTRADO_EM'] for x in (c1, c2, c3) if num(x['REGISTRADO_EM']))
        alvo_min = int(round((carimbo_min - int(carimbo_min)) * 1440)) + 1
        if int(carimbo_min) != serial(hoje):
            alvo_min = 1
        reais_h = minutos_reais(final(qa), AN, serial(hoje))
        while True:
            ag = dt.datetime.now()
            agm = ag.hour * 60 + ag.minute
            sh = slots(reais_h, alvo_min, agm - 1, 1, 3, folga=35)
            if len(sh) >= 3:
                break
            if agm - alvo_min > 240:
                raise RuntimeError(f'sem horario livre hoje depois do carimbo: reais {reais_h}')
            time.sleep(20)
        PB['horas_hoje'] = {'carimbo_antigo_max': carimbo_min, 'horas': [hhmm(x) for x in sh], 'reais_hoje': reais_h}
        reuso = {}
        for i, nv in enumerate((1, 2, 1)):
            qa.run('mIntegracao.IrDigitarResultados', teto=120)
            ac = xl.ActiveCell
            r = ac.Row - M.lo.DataBodyRange.Row + 1
            M.digitar(r, campos(hoje, hhmm(sh[i]), LOTE, nv, AN, val(nv, 0.0015 * (i + 1))))
            l_ = M.linha(r)
            reuso[f'B{i + 1}'] = {'ActiveCell': t(ac.Address), 'linha_rel': r,
                                  'qual': next((k2 for k2, v in R.items() if v == r), '?'), 'linha': l_,
                                  'nivel': nv, 'hora': hhmm(sh[i])}
        PB['reaproveitar'] = reuso
        PB['reaproveitar_avisos'] = M.avisos()
    finally:
        xl.Interactive = False
        xl.EnableEvents = True
    L['passos_B'] = PB
    ids_B = {M.id(R['r4'] + 1), M.id(R['r4'] + 2), M.id(R['cola']), M.id(R['s1']), M.id(R['r6']), M.id(R['dest'])} | \
        {v['linha']['ID_REGISTRO'] for v in reuso.values()}
    ids_todos |= ids_B | {t(s_n1['ID_REGISTRO']), t(s_n2['ID_REGISTRO'])}
    snapB = M.todas()
    # nenhum manual antigo perde o carimbo (as linhas que B nao tocou)
    tocadas = {R[x] for x in ('r1', 'r2', 'r3', 'r8', 'r9', 's2', 's3', 'orf', 'dest', 'cola', 's1')} | \
        {R['r4'] + 1, R['r4'] + 2} | {v['linha_rel'] for v in reuso.values()}
    L['carimbo_antigo_antes_atualizar2'] = {
        'iguais': {r: snap1[r] == snapB.get(r) for r in snap1 if r not in tocadas},
        'mudou': {r: (snap1[r], snapB.get(r)) for r in snap1 if r not in tocadas and snap1[r] != snapB.get(r)}}

    def depois(rot, snap_ref):
        rows = final(qa)
        tela = {k2: M.linha(r) for k2, r in R.items()}
        tela.update({'r4+1': M.linha(R['r4'] + 1), 'r4+2': M.linha(R['r4'] + 2)})
        for v in reuso.values():
            tela[f'reuso_{v["linha_rel"]}'] = M.linha(v['linha_rel'])
        D = {'final': meus_final(rows, ids_todos), 'qa': qa_meus(qa, ids_todos),
             'ResumoQA': t(qa.run('mIntegracao.ResumoQA', teto=120)), 'motor': t(qa.cfg('MOTOR_DADOS')),
             'tela': {k2: {c: v.get(c) for c in ('ID_REGISTRO', 'STATUS', 'DETALHE', 'USUARIO', 'REGISTRADO_EM_txt',
                                                  'SUBSTITUI_ID', 'RESULTADO', 'HORA', 'DATA')} for k2, v in tela.items()}}
        pa = painel_filtro(qa, AN, LOTE)
        D['painel'] = pa
        orc = {nv: oraculo(rows, AN, nv, lote=LOTE, ini=nov1, fim=nov30) for nv in range(1, NLV + 1)}
        D['painel_vs_oraculo'] = {nv: {'n': (pa['engPainel'][nv - 1]['n'], orc[nv]['n']),
                                       'rel_media': rel(pa['engPainel'][nv - 1]['media'], orc[nv].get('media')),
                                       'rel_dp': rel(pa['engPainel'][nv - 1]['dp'], orc[nv].get('dp'))}
                                  for nv in range(1, NLV + 1)}
        D['oraculo_ids'] = {nv: orc[nv]['ids'] for nv in range(1, NLV + 1)}
        D['eng'] = eng_slots(qa, ids_todos)
        D['estatistica'] = estat(qa, AN, LOTE, rows)
        D['audit'], _ = audit(qa, al_r0, ids_todos)
        snap = M.todas()
        D['carimbos_vs_ref'] = {r: snap_ref[r] == snap.get(r) for r in snap_ref}
        D['orfao'] = M.carimbo(R['orf'])
        return D, rows

    msg2, s2 = qa.atualizar()
    D2, rows2 = depois('2', snapB)
    D2['mensagem'], D2['segundos'] = msg2[:2500], round(s2, 1)
    # diferenca de populacao do Painel (ids) entre o atualizar1 e o 2
    o1 = {nv: set(oraculo(rows1, AN, nv, lote=LOTE, ini=nov1, fim=nov30)['ids']) for nv in range(1, NLV + 1)}
    D2['delta_ids'] = {nv: {'entraram': sorted(set(D2['oraculo_ids'][nv]) - o1[nv]),
                            'sairam': sorted(o1[nv] - set(D2['oraculo_ids'][nv]))} for nv in range(1, NLV + 1)}
    L['atualizar2'] = D2

    # ------------------------------------------------ reserva PQ (mesmo estado)
    if not SEM_PQ:
        try:
            antes_motor = qa.cfg('MOTOR_DADOS', 'PQ')
            t0 = time.time()
            r = str(qa.run('mIntegracao.AtualizarDadosAutomatico', teto=2400))
            rows_pq = final(qa)
            fq, fv = meus_final(rows_pq, ids_todos), D2['final']
            dif = {}
            for k2 in sorted(set(fq) | set(fv)):
                a, b = fv.get(k2, {}), fq.get(k2, {})
                d = {c: (a.get(c), b.get(c)) for c in ('STATUS_ANALITICO', 'PARTICIPA_ESTATISTICA', 'ID_RELACIONADO',
                                                        'MOTIVO_EXCLUSAO_AUTOMATICA', 'RESULTADO', 'NIVEL', 'USUARIO_MANUAL')
                     if t(a.get(c)) != t(b.get(c))}
                if d:
                    dif[k2] = d
            qpq = qa_meus(qa, ids_todos)

            def a01(lst):
                return sorted(v for x in lst if 'A01' in x.values() for v in x.values() if 'Digitar Resultados' in v
                              or 'inative' in v)
            L['reserva_pq'] = {'motor_antes': antes_motor, 'retorno': r[:800], 'segundos': round(time.time() - t0, 1),
                               'diferencas_vba_pq': dif, 'A01_vba': a01(D2['qa']), 'A01_pq': a01(qpq),
                               'A01_igual': a01(D2['qa']) == a01(qpq), 'ResumoQA_pq': t(qa.run('mIntegracao.ResumoQA')),
                               'linhas': (len(rows2), len(rows_pq))}
        except Exception as e:                                       # noqa: BLE001
            L['reserva_pq'] = {'erro': f'{type(e).__name__}: {e}'}
        finally:
            qa.cfg('MOTOR_DADOS', 'VBA')

    # ------------------------------------------------ fase C: corrigir a HORA da colada e apagar os IDs invalidos
    snapC0 = M.todas()
    xl.Interactive = True
    PC = {}
    try:
        hc = prox()
        M.esc(R['cola'], 'HORA', hc)
        PC['cola_hora'] = {'hora': hc, 'id': M.id(R['cola']), 'carimbo': M.carimbo(R['cola']), 'avisos': M.avisos()}
        ap = {}
        for nome in ('inv1', 'inv2', 'inv3', 'inv4', 'inv5'):
            antes = M.linha(R[nome])
            M.cel(R[nome], 'ID_REGISTRO').ClearContents()
            qa.ex.esperar()
            ap[nome] = {'DETALHE_antes': antes['DETALHE'], 'STATUS_antes': antes['STATUS'], 'id_antes': antes['ID_REGISTRO'],
                        'id_novo': M.id(R[nome]), 'carimbo': M.carimbo(R[nome])}
        PC['apagar_id_invalido'] = ap
        PC['avisos'] = M.avisos()
    finally:
        xl.Interactive = False
    ids_todos |= {v['id_novo'] for v in ap.values()}
    L['passos_C'] = PC
    msg3, s3 = qa.atualizar()
    D3, rows3 = depois('3', snapC0)
    D3['mensagem'], D3['segundos'] = msg3[:2500], round(s3, 1)
    L['atualizar3'] = D3
    L['avisos_todos'] = VIG.vistos


VIG = None


def main():
    global VIG
    os.makedirs(PASTA, exist_ok=True)
    src = os.path.join(C.REPO, f'QC_{PROD}.xlsm')
    tag = time.strftime('%H%M%S')
    dst = os.path.join(PASTA, f'{PROD}_nova_{tag}.xlsm')
    L = Log(os.path.join(PASTA, f'valida_r7_{PROD}.json'))
    h0 = sha_full(src)
    shutil.copy2(src, dst)
    L['origem'] = {'arquivo': os.path.relpath(src, C.REPO), 'sha256_antes': h0[:16], 'copia': os.path.basename(dst),
                   'sha_copia': sha_full(dst)[:16]}
    qa = qa_final.QA(PROD, dst)
    L['excel_pid'] = qa.ex.pid
    VIG = Vigia(qa.ex.pid)
    VIG.start()
    try:
        for nome, fn in (('detalhe_formula', lambda: parte_detalhe_formula(qa, L)),
                         ('registros', lambda: parte_registros(qa, L))):
            try:
                fn()
            except Exception as e:                                   # noqa: BLE001
                L[f'ERRO_{nome}'] = f'{type(e).__name__}: {e} | {traceback.format_exc()[-1200:]}'
        try:
            msg0, s0 = qa.atualizar()
            L['atualizar0'] = {'segundos': round(s0, 1), 'mensagem': msg0[:800]}
        except Exception as e:                                       # noqa: BLE001
            L['ERRO_atualizar0'] = f'{type(e).__name__}: {e}'
        rows0 = final(qa)
        L['tblCQ_Final_0'] = len(rows0)
        if not BIO:
            try:
                parte_estat_hema(qa, L, rows0)
            except Exception as e:                                   # noqa: BLE001
                L['ERRO_estat_hema'] = f'{type(e).__name__}: {e} | {traceback.format_exc()[-1200:]}'
        try:
            parte_manual(qa, L, rows0)
        except Exception as e:                                       # noqa: BLE001
            L['ERRO_manual'] = f'{type(e).__name__}: {e} | {traceback.format_exc()[-2000:]}'
            L['avisos_todos'] = VIG.vistos
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
