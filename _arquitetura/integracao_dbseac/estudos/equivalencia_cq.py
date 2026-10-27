# -*- coding: utf-8 -*-
"""Equivalencia do mEtlCQ (VBA) com o Power Query: a tblCQ_Final calculada em VBA = a gravada pelo PQ, linha a linha,
coluna a coluna. Roda numa COPIA cuja tblCQ_Final foi gerada pelo PQ com as mesmas entradas; nada e salvo.

Uso: python equivalencia_cq.py <Bioquimica|Hematologia> <copia.xlsm> <saida.json> [--n 3] [--fixtures]

--fixtures: casos de cada regra (manual, inativacao, comentario e, ADR-082, SUBSTITUI_ID: substituicao valida, chave
diferente, ID inexistente, ID de manual, dois manuais para o mesmo ID -- o 1o INATIVADO, que continua substituindo --,
interfaceamento INATIVADO (sem corrida propria: o RUN do manual), SEM_VALOR, SEM_DATA_HORA e retransmissao indicados,
manual sem MOTIVO). A copia precisa da coluna SUBSTITUI_ID (instalar_adr082.py).
Com elas, 'adr082_ok' confere o que o PQ decidiu contra o ESPERADO de cada caso (nao so PQ = VBA: se os dois motores
errassem igual, a equivalencia passaria) -- 'adr082_falhas' lista o que nao bateu.
"""
import datetime as dt
import json
import os
import sys
import tempfile
import time

AQUI = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(AQUI, '..', 'testes'))
from qa_final import QA  # noqa: E402
sys.path.insert(0, os.path.join(AQUI, '..'))
import instalar_integracao as ii  # noqa: E402

SRC = os.path.join(AQUI, '..', 'src', 'comum', 'mEtlCQ.bas')
TESTE = r'''
Option Explicit
Public Function RodarCQ() As String
    Dim t As Double, a As Variant, ws As Worksheet, cols As Variant
    t = Timer
    a = mEtlCQ.CalcularCQFinal()
    RodarCQ = Format$(Timer - t, "0.000") & "|" & mEtlCQ.TemposUltimo()
    Set ws = ThisWorkbook.Worksheets.Add
    ws.Name = "ETL_TESTE"
    cols = mEtlCQ.NomesColunas()
    ws.Range("A1").Resize(1, UBound(cols) + 1).Value = cols
    Dim j As Long, n As Long
    n = UBound(a, 1)
    For j = 1 To 40
        Select Case cols(j - 1)
            Case "DATA": ws.Cells(2, j).Resize(n, 1).NumberFormat = "dd/mm/yyyy"
            Case "HORA": ws.Cells(2, j).Resize(n, 1).NumberFormat = "hh:mm:ss"
            Case "DATA_HORA", "DATA_INATIVACAO", "RECEBIDO_EM": ws.Cells(2, j).Resize(n, 1).NumberFormat = "dd/mm/yyyy hh:mm:ss.000"
            Case "CORRIDA_NO_DIA", "RUN", "NIVEL", "RESULTADO", "ID_ORIGEM", "ITEM_ID"
            Case Else: ws.Cells(2, j).Resize(n, 1).NumberFormat = "@"
        End Select
    Next j
    ws.Range("A2").Resize(n, UBound(a, 2)).Value2 = a
End Function
'''


def importar(wb, caminho_bas):
    texto = open(caminho_bas, encoding='utf-8').read()
    tmp = os.path.join(tempfile.gettempdir(), os.path.basename(caminho_bas))
    with open(tmp, 'w', encoding='cp1252', newline='\r\n') as f:
        f.write(texto.replace('\r\n', '\n'))
    nome = os.path.splitext(os.path.basename(caminho_bas))[0]
    for c in list(wb.VBProject.VBComponents):
        if c.Name == nome:
            wb.VBProject.VBComponents.Remove(c)
    wb.VBProject.VBComponents.Import(tmp)


def escrever(q, tabela, linhas):
    """Grava linhas (dicts por nome de coluna) nas primeiras linhas da tabela de entrada, com eventos desligados."""
    lo = None
    for ws in q.wb.Worksheets:
        try:
            lo = ws.ListObjects(tabela)
            break
        except Exception:
            pass
    ws = lo.Parent
    ws.Unprotect(ii.SENHA)
    cab = [str(c.Name) for c in lo.ListColumns]
    for i, d in enumerate(linhas, 1):
        for k, v in d.items():
            if k in cab:
                lo.DataBodyRange.Cells(i, cab.index(k) + 1).Value = v


def fixtures(q):
    """Casos de cada regra da spec (manual, inativacao, comentario) sobre linhas reais da tabela; depois o PQ refaz a
    tblCQ_Final (a referencia) com essas entradas."""
    import datetime as _dt
    lo = q.wb.Worksheets('Principal - Resultados').ListObjects('tblCQ_Final')
    cab = [str(c.Name) for c in lo.ListColumns]
    d = lo.DataBodyRange.Value
    lo_m = q.wb.Worksheets('Digitar Resultados').ListObjects('tblResultados_Manuais')
    if 'SUBSTITUI_ID' not in [str(c.Name) for c in lo_m.ListColumns]:
        raise SystemExit('a copia nao tem a coluna SUBSTITUI_ID (ADR-082): rode o instalar_adr082.py nela antes')
    ativos = [r for r in d if r[cab.index('STATUS_ANALITICO')] == 'ATIVO']
    a, b, c = ativos[100], ativos[200], ativos[300]
    col = lambda r, n: r[cab.index(n)]
    dh = col(a, 'DATA_HORA').replace(tzinfo=None)
    num = lambda r: str(col(r, 'ID_REGISTRO')).split('-')[-1]
    q.ex.xl.EnableEvents = False
    subs = fixture_substitui(cab, d, a)
    m25 = next(x for x in subs if x['ID_REGISTRO'] == 'MAN_0025')
    escrever(q, 'tblInativacao_NaoConformes', [
        {'ID_REGISTRO': col(a, 'ID_REGISTRO'), 'ANALITO': col(a, 'ANALITO'), 'REGISTRAR - LJ': True, 'MOTIVO': 'QA motivo 1',
         'DATA_INATIVACAO': _dt.datetime(2026, 10, 20, 9, 0, 0), 'USUARIO': 'QA'},
        {'ID_REGISTRO': ' ' + num(b) + ' ', 'ANALITO': str(col(b, 'ANALITO')).lower(), 'REGISTRAR - LJ': False, 'MOTIVO': None,
         'USUARIO': 'QA'},
        {'ID_REGISTRO': str(col(c, 'ID_REGISTRO')).lower(), 'ANALITO': 'OUTRO ANALITO', 'REGISTRAR - LJ': 'SIM', 'MOTIVO': 'x'},
        {'ID_REGISTRO': col(a, 'ID_REGISTRO'), 'ANALITO': col(a, 'ANALITO'), 'MOTIVO': 'repetido (ignorado)'},
        # ADR-082 (revisao): o 1o manual de dois para o mesmo ID, inativado, continua substituindo (E12 do 2o diz isso)
        {'ID_REGISTRO': 'MAN_0025', 'ANALITO': m25['ANALITO'], 'REGISTRAR - LJ': True, 'MOTIVO': 'QA manual substituto inativado',
         'USUARIO': 'QA'},
    ])
    escrever(q, 'tblComentariosTecnicos', [
        {'ID_REGISTRO': col(b, 'ID_REGISTRO'), 'COMENTARIO_TECNICO': 'coment 1'},
        {'ID_REGISTRO': num(b), 'COMENTARIO_TECNICO': 'coment 2'},
        {'ID_REGISTRO': col(c, 'ID_REGISTRO'), 'COMENTARIO_TECNICO': '   '},
    ])
    base = {'EQUIPAMENTO': col(a, 'EQUIPAMENTO'), 'LOTE': col(a, 'LOTE'), 'NIVEL': col(a, 'NIVEL'), 'ANALITO': col(a, 'ANALITO'),
            'MOTIVO': 'QA manual', 'USUARIO': 'QA', 'REGISTRADO_EM': _dt.datetime(2026, 10, 20, 12, 0, 0)}
    dia = _dt.datetime(dh.year, dh.month, dh.day)
    hm = lambda m: (dh + _dt.timedelta(minutes=m)).strftime('%H:%M')
    escrever(q, 'tblResultados_Manuais', [
        dict(base, ID_REGISTRO='MAN_0001', DATA=dia, HORA=hm(5), RESULTADO=1.25),            # conflito (5 min)
        dict(base, ID_REGISTRO='man_2', DATA=dia, HORA=hm(45), RESULTADO='2,5'),             # valido (45 min)
        dict(base, ID_REGISTRO='MAN_0002', DATA=dia, HORA=hm(50), RESULTADO=3.0),            # ID repetido
        dict(base, ID_REGISTRO='X1', DATA=dia, HORA=hm(60), RESULTADO=1.0),                  # ID invalido
        dict(base, ID_REGISTRO='MAN_0003', DATA=dia, HORA=None, RESULTADO=1.0),              # sem hora
        dict(base, ID_REGISTRO='MAN_0004', DATA=dia, HORA=hm(70), RESULTADO='1.5'),          # ambiguo
        dict(base, ID_REGISTRO='MAN_0005', DATA=dia, HORA=hm(45), RESULTADO=2.6),            # mesma chave e instante
        dict(base, ID_REGISTRO='MAN_0006', DATA=_dt.datetime(2027, 1, 1), HORA='08:00', RESULTADO=1.0),  # futuro
        dict(base, ID_REGISTRO='MAN_0007', DATA=dia, HORA=hm(80), NIVEL=9, RESULTADO=1.0),   # nivel invalido
        dict(base, ID_REGISTRO='MAN_0008', DATA=dia, HORA=hm(90), RESULTADO='abc'),          # resultado invalido
    ] + fixture_fim_do_dia(cab, d) + subs)
    q.ex.xl.EnableEvents = True
    q.run('mIntegracao.Refrescar', 'tblCQ_Final', teto=900)


def fixture_fim_do_dia(cab, d):
    """qa_etl E11a (20/10/2026): manual N1 as 23:50 num dia cujo ultimo bloco so tem N2 -- corrida propria no fim.
    Usa o ultimo dia com resultado N1 do analito/lote da ultima linha ATIVA do interfaceamento."""
    import datetime as _dt
    col = lambda r, n: r[cab.index(n)]
    intf = [r for r in d if col(r, 'ORIGEM_RESULTADO') == 'INTERFACEAMENTO' and col(r, 'STATUS_ANALITICO') == 'ATIVO'
            and col(r, 'NIVEL') == 1 and hasattr(col(r, 'DATA_HORA'), 'year')]
    if not intf:
        return []
    # o caso real do E11a (Bioquimica, 14/10/2026: Bilirrubina total 8977 com o bloco das 09:14 so de N2) e o generico
    alvos = [r for r in intf if col(r, 'ANALITO') == 'Bilirrubina total' and str(col(r, 'LOTE')).strip() == '8977'][-1:]
    out = []
    for k, u in enumerate(alvos + [intf[-1]]):
        dh = col(u, 'DATA_HORA').replace(tzinfo=None)
        out.append({'ID_REGISTRO': f'MAN_{9 + k:04d}', 'EQUIPAMENTO': col(u, 'EQUIPAMENTO'), 'LOTE': col(u, 'LOTE'),
                    'NIVEL': 1, 'ANALITO': col(u, 'ANALITO'), 'DATA': _dt.datetime(dh.year, dh.month, dh.day),
                    'HORA': '23:50', 'RESULTADO': 999.5, 'MOTIVO': 'QA fim do dia', 'USUARIO': 'QA'})
    return out


# ADR-082: o que cada manual da fixture_substitui deve dar -- {MAN_00nn: (caso, ID do alvo)}; conferido em conferir_adr082
ESPERADO_ADR082 = {}


def fixture_substitui(cab, d, a):
    """ADR-082: SUBSTITUI_ID. a = o resultado ATIVO que a fixture inativa (vira o caso "interfaceamento INATIVADO
    substituido": continua INATIVADO, sem corrida propria -- o RUN do manual). Os alvos sao do interfaceamento, longe de
    a; a DATA/HORA do manual e a do alvo. Grava em ESPERADO_ADR082 o caso e o alvo de cada manual."""
    import datetime as _dt
    col = lambda r, n: r[cab.index(n)]                                          # noqa: E731
    dhr = lambda r: col(r, 'DATA_HORA').replace(tzinfo=None)                    # noqa: E731
    intf = [r for r in d if col(r, 'ORIGEM_RESULTADO') == 'INTERFACEAMENTO' and hasattr(col(r, 'DATA_HORA'), 'year')]
    atv = [r for r in intf if col(r, 'STATUS_ANALITICO') == 'ATIVO' and col(r, 'ID_REGISTRO') != col(a, 'ID_REGISTRO')]
    reg_em = _dt.datetime(2030, 1, 1, 12, 0, 0)          # depois de todo resultado: nenhum "posterior ao lancamento"

    def man(k, alvo, sub, **extra):
        base = {'ID_REGISTRO': f'MAN_{k:04d}', 'SUBSTITUI_ID': sub, 'EQUIPAMENTO': col(alvo, 'EQUIPAMENTO'),
                'MATRIZ': col(alvo, 'MATRIZ'), 'LOTE': col(alvo, 'LOTE'), 'NIVEL': col(alvo, 'NIVEL'),
                'ANALITO': col(alvo, 'ANALITO'), 'RESULTADO': 1.5, 'MOTIVO': 'QA substitui', 'USUARIO': 'QA',
                'REGISTRADO_EM': reg_em}
        if hasattr(col(alvo, 'DATA_HORA'), 'year'):
            dh = dhr(alvo)
            base.update(DATA=_dt.datetime(dh.year, dh.month, dh.day), HORA=dh.strftime('%H:%M:%S'))
        base.update(extra)
        return base

    num = lambda r: str(col(r, 'ID_REGISTRO')).split('-')[-1]                   # noqa: E731
    e, f, g, h = atv[400], atv[500], atv[600], atv[700]
    dh_a = dhr(a)
    longe = lambda m: dict(DATA=_dt.datetime(dh_a.year, dh_a.month, dh_a.day),   # noqa: E731
                           HORA=(dh_a + _dt.timedelta(minutes=m)).strftime('%H:%M'))
    out = [
        man(21, e, num(e)),                                                      # valido (numero puro)
        man(22, f, col(f, 'ID_REGISTRO'), NIVEL=2 if col(f, 'NIVEL') == 1 else 1),   # chave diferente (outro nivel)
        man(23, a, '999999999', **longe(200)),                                   # ID inexistente
        man(24, a, 'man_21', **longe(210)),                                      # ID de um manual
        man(25, g, col(g, 'ID_REGISTRO')),                                       # dois manuais: vale o 1o
        man(26, g, num(g), HORA=(dhr(g) + _dt.timedelta(minutes=2)).strftime('%H:%M:%S')),   # ... o 2o e CONFLITO
        man(27, a, col(a, 'ID_REGISTRO'), HORA=(dh_a + _dt.timedelta(minutes=1)).strftime('%H:%M:%S')),  # INATIVADO
        man(29, h, num(h), MOTIVO=None),                                         # sem MOTIVO: MANUAL_INCOMPLETO
    ]
    ESPERADO_ADR082.clear()
    ESPERADO_ADR082.update({
        'MAN_0021': ('valido', col(e, 'ID_REGISTRO')), 'MAN_0022': ('outra_chave', col(f, 'ID_REGISTRO')),
        'MAN_0023': ('inexistente', None), 'MAN_0024': ('de_manual', None),
        'MAN_0025': ('valido_man_inativado', col(g, 'ID_REGISTRO')), 'MAN_0026': ('repetido', col(g, 'ID_REGISTRO')),
        'MAN_0027': ('valido_inativado', col(a, 'ID_REGISTRO')), 'MAN_0029': ('sem_motivo', col(h, 'ID_REGISTRO'))})
    # alvos que dependem do dado real (pulados se a base nao tiver)
    sv = next((r for r in intf if col(r, 'STATUS_ANALITICO') == 'SEM_VALOR'), None)
    if sv is not None:
        out.append(man(28, sv, col(sv, 'ID_REGISTRO')))                          # SEM_VALOR substituido
        ESPERADO_ADR082['MAN_0028'] = ('valido', col(sv, 'ID_REGISTRO'))
    dup = next((r for r in intf if col(r, 'STATUS_ANALITICO') == 'DUPLICIDADE_ORIGEM'), None)
    if dup is not None:
        out.append(man(30, dup, col(dup, 'ID_REGISTRO')))                        # retransmissao: invalido
        ESPERADO_ADR082['MAN_0030'] = ('retransmissao', col(dup, 'ID_REGISTRO'))
    sdh = next((r for r in d if col(r, 'ORIGEM_RESULTADO') == 'INTERFACEAMENTO'
                and col(r, 'STATUS_ANALITICO') == 'SEM_DATA_HORA' and hasattr(col(r, 'DATA'), 'year')), None)
    if sdh is not None:
        dd = col(sdh, 'DATA').replace(tzinfo=None)
        out.append(man(31, sdh, col(sdh, 'ID_REGISTRO'), DATA=_dt.datetime(dd.year, dd.month, dd.day), HORA='12:00'))
        ESPERADO_ADR082['MAN_0031'] = ('valido', col(sdh, 'ID_REGISTRO'))
    return out


def conferir_adr082(cab, dados):
    """O que o PQ decidiu nos manuais da fixture_substitui x o ESPERADO de cada caso (ADR-082, com a revisao: o alvo
    inativado nao tem corrida propria -- fica com o RUN/corrida do manual). Devolve (ok, falhas)."""
    ix = {c: cab.index(c) for c in ('ID_REGISTRO', 'STATUS_ANALITICO', 'ID_RELACIONADO', 'MOTIVO_EXCLUSAO_AUTOMATICA',
                                    'RUN', 'CORRIDA_NO_DIA', 'INATIVACAO_REGISTRADA')}
    por_id = {str(r[ix['ID_REGISTRO']]): r for r in dados}
    v = lambda r, c: norm(c, r[ix[c]]) if r is not None else None              # noqa: E731
    falhas = []
    if not ESPERADO_ADR082:
        return False, ['fixture_substitui nao rodou (ESPERADO_ADR082 vazio)']
    for man, (caso, alvo) in sorted(ESPERADO_ADR082.items()):
        m, t = por_id.get(man), por_id.get(alvo) if alvo else None
        mot = str(v(m, 'MOTIVO_EXCLUSAO_AUTOMATICA') or '')
        if m is None:
            falhas.append(f'{man}: ausente da tblCQ_Final')
            continue
        if caso in ('valido', 'valido_inativado', 'valido_man_inativado'):
            # o manual inativado (fixtures: MAN_0025) continua substituindo
            st_m = 'INATIVADO' if caso == 'valido_man_inativado' else 'ATIVO'
            ok = (t is not None and v(m, 'STATUS_ANALITICO') == st_m and v(m, 'ID_RELACIONADO') == alvo
                  and v(t, 'ID_RELACIONADO') == man and mot == f'substitui o resultado do interfaceamento {alvo}')
            if ok and v(t, 'INATIVACAO_REGISTRADA') == 'SIM':
                st = v(t, 'STATUS_ANALITICO')
                ok = st in ('INATIVADO', 'SEM_VALOR', 'SEM_DATA_HORA')
                if st in ('INATIVADO', 'SEM_VALOR'):
                    ok = ok and v(t, 'RUN') is not None and \
                        (v(t, 'RUN'), v(t, 'CORRIDA_NO_DIA')) == (v(m, 'RUN'), v(m, 'CORRIDA_NO_DIA'))
            elif ok:
                ok = (v(t, 'STATUS_ANALITICO') == 'SUBSTITUIDO_MANUAL' and v(t, 'RUN') is None
                      and v(t, 'CORRIDA_NO_DIA') is None and caso != 'valido_inativado')
        elif caso == 'outra_chave':
            ok = 'SUBSTITUI_ID invalido: ' in mot and 'tem outra chave' in mot and t is not None \
                and v(t, 'STATUS_ANALITICO') != 'SUBSTITUIDO_MANUAL' and v(t, 'ID_RELACIONADO') != man
        elif caso == 'inexistente':
            ok = 'SUBSTITUI_ID invalido: o ID ' in mot and 'nao existe nos resultados do interfaceamento' in mot
        elif caso == 'de_manual':
            ok = 'SUBSTITUI_ID invalido: ' in mot and 'e resultado manual' in mot
        elif caso == 'repetido':
            ok = (v(m, 'STATUS_ANALITICO') == 'CONFLITO_MANUAL' and v(m, 'ID_RELACIONADO') == 'MAN_0025'
                  and f'outro manual (MAN_0025) ja substitui {alvo}' in mot)
        elif caso == 'sem_motivo':
            ok = (v(m, 'STATUS_ANALITICO') == 'MANUAL_INCOMPLETO' and 'MOTIVO (obrigatorio' in mot and t is not None
                  and v(t, 'STATUS_ANALITICO') != 'SUBSTITUIDO_MANUAL' and v(t, 'ID_RELACIONADO') is None)
        elif caso == 'retransmissao':
            ok = 'e retransmissao de' in mot and t is not None and v(t, 'STATUS_ANALITICO') == 'DUPLICIDADE_ORIGEM'
        else:
            ok = False
        if not ok:
            falhas.append(f'{man} ({caso}): manual {[str(v(m, c)) for c in ("STATUS_ANALITICO", "ID_RELACIONADO", "RUN")]} '
                          f'{mot[:120]!r} | alvo {alvo}: '
                          f'{[str(v(t, c)) for c in ("STATUS_ANALITICO", "ID_RELACIONADO", "RUN", "INATIVACAO_REGISTRADA")]}')
    # ninguem aponta para um manual invalido (nada substituido em silencio)
    invalidos = {k for k, (c, _) in ESPERADO_ADR082.items() if not c.startswith('valido')}
    for r in dados:
        if str(r[cab.index('ORIGEM_RESULTADO')]) == 'INTERFACEAMENTO' and str(v(r, 'ID_RELACIONADO') or '') in invalidos:
            falhas.append(f'{r[ix["ID_REGISTRO"]]}: aponta o manual invalido {v(r, "ID_RELACIONADO")}')
    return not falhas, falhas


def norm(c, v):
    if v is None or v == '':
        return None
    if hasattr(v, 'year'):
        v = v.replace(tzinfo=None)
        if c == 'DATA':
            return v.date()
        if c == 'HORA':
            return (v.hour, v.minute, v.second)
        return v.replace(microsecond=round(v.microsecond / 1000) * 1000 % 1000000)
    if isinstance(v, float) and c in ('CORRIDA_NO_DIA', 'RUN', 'NIVEL', 'ID_ORIGEM', 'ITEM_ID') and v == int(v):
        return int(v)
    return v


def main(produto, arq, saida, n=3, fix=False):
    import lint_vba
    erros = lint_vba.checar(SRC)
    if erros:                                    # erro de compilacao abriria a janela do VBA na tela do usuario
        raise SystemExit('lint VBA falhou, Excel NAO aberto: ' + ' | '.join(erros))
    res = {'produto': produto, 'tempos_s': []}
    q = QA(produto, arq)
    try:
        importar(q.wb, SRC)
        comp = q.wb.VBProject.VBComponents.Add(1)
        comp.Name = 'mTesteEtl'
        comp.CodeModule.AddFromString(TESTE)
        if ii.compilar_vba(q.ex, q.wb, teto=20) is False:
            raise SystemExit('VBA nao compilou')
        if fix:
            fixtures(q)
        pq = q.wb.Worksheets('Principal - Resultados').ListObjects('tblCQ_Final')
        cab_pq = [str(c.Name) for c in pq.ListColumns]
        dados_pq = pq.DataBodyRange.Value
        for i in range(n):
            if i > 0:
                q.ex.xl.DisplayAlerts = False
                q.wb.Worksheets('ETL_TESTE').Delete()
            r = str(q.run('mTesteEtl.RodarCQ', teto=900))
            res['tempos_s'].append(float(r.split('|')[0].replace(',', '.')))
            res.setdefault('fases', []).append(r.split('|')[1] if '|' in r else '')
            print('fases', res['fases'][-1], flush=True)
            print('tempo VBA', res['tempos_s'][-1], flush=True)
        ws = q.wb.Worksheets('ETL_TESTE')
        cab_vba = [str(v) for v in ws.Range(ws.Cells(1, 1), ws.Cells(1, 40)).Value[0]]
        if len(cab_vba) != 40 or len([c for c in cab_vba if c in cab_pq]) < 39:
            raise SystemExit(f'cabecalho VBA ilegivel: {cab_vba[:5]}')
        nl = len(dados_pq)
        nv = int(q.ex.xl.WorksheetFunction.CountA(ws.Range('A:A'))) - 1
        dados_vba = ws.Range(ws.Cells(2, 1), ws.Cells(nv + 1, 40)).Value
        res['linhas'] = {'pq': nl, 'vba': len(dados_vba)}
        div, ex = {}, {}
        comuns = [c for c in cab_vba if c in cab_pq]
        res['colunas_so_vba'] = [c for c in cab_vba if c not in cab_pq]
        res['colunas_pq_nao_comparadas'] = [c for c in cab_pq if c not in cab_vba]
        res['cab_vba'] = cab_vba
        for r in range(min(nl, len(dados_vba))):
            for c in comuns:
                a = norm(c, dados_pq[r][cab_pq.index(c)])
                b = norm(c, dados_vba[r][cab_vba.index(c)])
                if a != b:
                    div[c] = div.get(c, 0) + 1
                    ex.setdefault(c, (r + 2, str(dados_pq[r][0]), str(a), str(b)))
        res['divergencias'] = div
        ist = cab_pq.index('STATUS_ANALITICO')
        cont = {}
        for r in dados_pq:
            cont[r[ist]] = cont.get(r[ist], 0) + 1
        res['status_pq'] = cont
        cj = [cab_pq.index(c) for c in ('INATIVACAO_REGISTRADA', 'COMENTARIO_TECNICO', 'GOVERNANCA', 'ID_RELACIONADO')]
        res['marcas_pq'] = {'inativacao_SIM': sum(1 for r in dados_pq if r[cj[0]] == 'SIM'),
                            'com_comentario': sum(1 for r in dados_pq if r[cj[1]]),
                            'governanca_erro': sum(1 for r in dados_pq if str(r[cj[2]]).startswith('ERRO')),
                            'com_id_relacionado': sum(1 for r in dados_pq if r[cj[3]])}
        print('status PQ', cont, res['marcas_pq'], flush=True)
        if fix:
            # ADR-082: o que o PQ decidiu nos manuais do SUBSTITUI_ID (MAN_0021..MAN_0031) e nos resultados indicados
            ci = {c: cab_pq.index(c) for c in ('ID_REGISTRO', 'STATUS_ANALITICO', 'ID_RELACIONADO',
                                                'MOTIVO_EXCLUSAO_AUTOMATICA', 'RUN')}
            man = {r[ci['ID_REGISTRO']]: r for r in dados_pq
                   if str(r[ci['ID_REGISTRO']]).startswith('MAN_00') and 21 <= int(str(r[ci['ID_REGISTRO']])[4:8]) <= 31}
            rel = {r[ci['ID_RELACIONADO']] for r in man.values() if r[ci['ID_RELACIONADO']]}
            alvo = {r[ci['ID_REGISTRO']]: r for r in dados_pq if r[ci['ID_REGISTRO']] in rel}
            todos = dict(alvo, **man)
            res['adr082'] = {i: {c: str(todos[i][j]) for c, j in ci.items() if c != 'ID_REGISTRO'}
                             for i in sorted(todos, key=str)}
            res['adr082_substituidos_pq'] = cont.get('SUBSTITUIDO_MANUAL', 0)
            res['adr082_ok'], res['adr082_falhas'] = conferir_adr082(cab_pq, dados_pq)
            print('ADR-082 SUBSTITUIDO_MANUAL no PQ:', res['adr082_substituidos_pq'], '| esperado:',
                  'OK' if res['adr082_ok'] else 'FALHOU', flush=True)
            for x in res['adr082_falhas']:
                print('  ADR-082', x, flush=True)
        res['exemplos'] = ex
        print(json.dumps({k: res[k] for k in ('linhas', 'colunas_so_vba', 'colunas_pq_nao_comparadas', 'divergencias')}, ensure_ascii=False), flush=True)
        for c, e in ex.items():
            print('  ex', c, e, flush=True)
    finally:
        q.fechar(salvar=False)
    with open(saida, 'w', encoding='utf-8') as f:
        json.dump(res, f, ensure_ascii=False, indent=1, default=str)


if __name__ == '__main__':
    a = sys.argv[1:]
    main(a[0], os.path.abspath(a[1]), a[2], int(a[a.index('--n') + 1]) if '--n' in a else 3, '--fixtures' in a)
