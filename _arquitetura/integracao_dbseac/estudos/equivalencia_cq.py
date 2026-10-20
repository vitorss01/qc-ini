# -*- coding: utf-8 -*-
"""Equivalencia do mEtlCQ (VBA) com o Power Query: a tblCQ_Final calculada em VBA = a gravada pelo PQ, linha a linha,
coluna a coluna. Roda numa COPIA cuja tblCQ_Final foi gerada pelo PQ com as mesmas entradas; nada e salvo.

Uso: python equivalencia_cq.py <Bioquimica|Hematologia> <copia.xlsm> <saida.json> [--n 3]
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
    ativos = [r for r in d if r[cab.index('STATUS_ANALITICO')] == 'ATIVO']
    a, b, c = ativos[100], ativos[200], ativos[300]
    col = lambda r, n: r[cab.index(n)]
    dh = col(a, 'DATA_HORA').replace(tzinfo=None)
    num = lambda r: str(col(r, 'ID_REGISTRO')).split('-')[-1]
    q.ex.xl.EnableEvents = False
    escrever(q, 'tblInativacao_NaoConformes', [
        {'ID_REGISTRO': col(a, 'ID_REGISTRO'), 'ANALITO': col(a, 'ANALITO'), 'REGISTRAR - LJ': True, 'MOTIVO': 'QA motivo 1',
         'DATA_INATIVACAO': _dt.datetime(2026, 10, 20, 9, 0, 0), 'USUARIO': 'QA'},
        {'ID_REGISTRO': ' ' + num(b) + ' ', 'ANALITO': str(col(b, 'ANALITO')).lower(), 'REGISTRAR - LJ': False, 'MOTIVO': None,
         'USUARIO': 'QA'},
        {'ID_REGISTRO': str(col(c, 'ID_REGISTRO')).lower(), 'ANALITO': 'OUTRO ANALITO', 'REGISTRAR - LJ': 'SIM', 'MOTIVO': 'x'},
        {'ID_REGISTRO': col(a, 'ID_REGISTRO'), 'ANALITO': col(a, 'ANALITO'), 'MOTIVO': 'repetido (ignorado)'},
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
    ])
    q.ex.xl.EnableEvents = True
    q.run('mIntegracao.Refrescar', 'tblCQ_Final', teto=900)


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
