# -*- coding: utf-8 -*-
"""ADR-087: auditoria visual das telas (abas) de uma COPIA instalada -- nada e salvo.

Uso: python auditoria_abas.py <Bioquimica|Hematologia> <copia.xlsm> <pasta_saida> [abas separadas por ;]

Para cada tela:
  1. IMAGEM: exporta a area visivel de trabalho (A1 ate a ultima coluna/linha do conteudo da tela, limitada) para PDF
     e converte em PNG (pasta_saida/<aba>.png) -- a inspecao visual nao usa a tela do usuario;
  2. BOTOES/FORMAS: nenhum par de formas visiveis se sobrepoe (fora os graficos e o que fica dentro deles), todo botao
     cabe na area da tela e o texto dele cabe nele (medido com a fonte real);
  3. CELULAS: nenhum numero/data aparece como "####" (coluna estreita demais); texto que nao cabe na coluna e cuja
     vizinha da direita esta preenchida (ou que esta mesclado) = CORTADO, salvo com quebra de linha e altura bastante.
Sai com 1 se achar problema; o JSON em pasta_saida/auditoria_<produto>.json tem cada achado com a celula/forma.
"""
import json
import os
import sys
import time

AQUI = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, AQUI)
from qa_final import QA  # noqa: E402
sys.path.insert(0, os.path.join(AQUI, '..'))
import instalar_integracao as ii  # noqa: E402
from PIL import ImageFont  # noqa: E402

TELAS = ['Início', 'Painel', 'Registros', 'Inativar', 'COMENTARIOS_TECNICOS', 'Digitar Resultados', 'Liberação',
         'EQA.CAP_Resumo', 'EQA.Controllab_Resumo', 'EQA.CAP_Dados', 'EQA.Controllab_Dados', 'Estatística',
         'Analitos', 'Configuração', 'Principal - Resultados', 'Eventos_Westgard', 'Usuarios']
LIM_LIN, LIM_COL = 60, 30                       # area conferida/exportada de cada tela (o que cabe numa tela grande)
FONTES = {('segoe ui', False): 'segoeui.ttf', ('segoe ui', True): 'segoeuib.ttf', ('calibri', False): 'calibri.ttf',
          ('calibri', True): 'calibrib.ttf', ('arial', False): 'arial.ttf', ('arial', True): 'arialbd.ttf'}
_cache = {}


def largura_pt(texto, nome, bold, size):
    k = (str(nome).lower(), bool(bold))
    arq = FONTES.get(k, FONTES[('calibri', bool(bold))])
    kk = (arq, size)
    if kk not in _cache:
        _cache[kk] = ImageFont.truetype(os.path.join(r'C:\Windows\Fonts', arq), max(1, round(size * 96 / 72)))
    return _cache[kk].getlength(texto) * 72 / 96


def cruzam(a, b, folga=0.75):
    return a[0] < b[0] + b[2] - folga and b[0] < a[0] + a[2] - folga and a[1] < b[1] + b[3] - folga and b[1] < a[1] + a[3] - folga


def _linhas(texto, largura, nome, bold, size):
    n = 0
    for par in texto.split('\n'):
        atual, k = '', 1
        for p in par.split(' '):
            cand = (atual + ' ' + p).strip()
            if largura_pt(cand, nome, bold, size) <= largura or not atual:
                atual = cand
            else:
                k += 1
                atual = p
        n += k
    return n


def ret_texto(cel):
    """Retangulo (pt) onde o texto da celula e desenhado: alinhamento horizontal/vertical, quebra de linha e o
    transbordo para a direita (texto a esquerda sem quebra passa da celula). None = celula vazia."""
    area = cel.MergeArea
    c1 = area.Cells(1, 1)
    t = str(c1.Text or '')
    if not t.strip():
        return None
    f = c1.Font
    nome, bold, size = f.Name, f.Bold == -1 or f.Bold is True, float(f.Size or 11)
    L, T, W, H = float(area.Left), float(area.Top), float(area.Width), float(area.Height)
    if c1.WrapText:
        n = _linhas(t, W - 4, nome, bold, size)
        w = min(W - 4, max(largura_pt(x, nome, bold, size) for x in t.split('\n')))
    else:
        n = t.count('\n') + 1
        w = max(largura_pt(x, nome, bold, size) for x in t.split('\n'))
    if c1.ShrinkToFit and w > W - 4:
        w = W - 4
    h = min(n * size * 1.25, H)
    ha, va = c1.HorizontalAlignment, c1.VerticalAlignment
    numero = isinstance(c1.Value, (int, float)) and not isinstance(c1.Value, bool)
    if ha in (-4108, 7):
        x = L + (W - w) / 2
    elif ha == -4152 or (ha == 1 and numero):
        x = L + W - 2 - w
    else:
        x = L + 2
    if va == -4160:
        y = T + 1
    elif va == -4108:
        y = T + (H - h) / 2
    else:
        y = T + H - 1 - h
    return (x, y, max(w, 1), max(h, 1))


def valores_fora_do_eixo(ch):
    """Eixo de valores com minimo/maximo FIXOS que cortam dados do grafico: (min, max, eixo_min, eixo_max) ou None."""
    try:
        ax = ch.Axes(2)
        if ax.MaximumScaleIsAuto and ax.MinimumScaleIsAuto:
            return None
        vals = []
        for i in range(1, ch.SeriesCollection().Count + 1):
            s = ch.SeriesCollection(i)
            try:
                if s.AxisGroup != 1:
                    continue
                vals += [float(v) for v in (s.Values or ()) if isinstance(v, (int, float)) and not isinstance(v, bool)]
            except Exception:                     # noqa: BLE001
                continue
        if not vals:
            return None
        lo, hi = float(ax.MinimumScale), float(ax.MaximumScale)
        if max(vals) > hi + 1e-9 or min(vals) < lo - 1e-9:
            return (min(vals), max(vals), lo, hi)
    except Exception:                             # noqa: BLE001
        return None
    return None


def textos_sob(ws, rets, max_lin=300, max_col=80):
    """Retangulos de texto (ret_texto) de celulas preenchidas que cruzam algum dos retangulos dados."""
    if not rets:
        return []
    y0, y1 = min(r[1] for r in rets), max(r[1] + r[3] for r in rets)
    x1 = max(r[0] + r[2] for r in rets)
    linhas = [r for r in range(1, max_lin + 1)
              if not ws.Rows(r).Hidden and float(ws.Rows(r).Top) < y1 and float(ws.Rows(r).Top) + float(ws.Rows(r).RowHeight) > y0]
    out = []
    for r in linhas:
        for c in range(1, max_col + 1):
            cel = ws.Cells(r, c)
            if float(cel.Left) > x1:
                break
            if ws.Columns(c).Hidden or cel.Formula in (None, ''):
                continue
            if cel.MergeArea.Cells(1, 1).Address != cel.Address:
                continue
            rt = ret_texto(cel)
            if rt and any(cruzam(x, rt) for x in rets):
                out.append(rt)
    return out


def auditar_aba(ws):
    achados = []
    ur = ws.UsedRange
    ult_l = min(ur.Row + ur.Rows.Count - 1, LIM_LIN)
    ult_c = min(ur.Column + ur.Columns.Count - 1, LIM_COL)
    # ---- formas
    formas = []
    for sh in ws.Shapes:
        try:
            if not sh.Visible:
                continue
            tipo = int(sh.Type)
        except Exception:                         # noqa: BLE001
            continue
        r = (float(sh.Left), float(sh.Top), float(sh.Width), float(sh.Height))
        formas.append((str(sh.Name), tipo, r, sh))
    graficos = [f for f in formas if f[1] == 3]   # msoChart
    for i, (n1, t1, r1, s1) in enumerate(formas):
        if t1 == 3:
            continue
        if any(cruzam(r1, g[2]) and r1[0] >= g[2][0] - 1 and r1[1] >= g[2][1] - 1 for g in graficos):
            continue                               # dentro de um grafico (caixa de dica): parte dele
        for n2, t2, r2, s2 in formas[i + 1:]:
            if t2 == 3 or n2.startswith('qcAssin_') or n1.startswith('qcAssin_'):
                continue
            if cruzam(r1, r2):
                achados.append({'tipo': 'formas_sobrepostas', 'a': n1, 'b': n2, 'ret_a': r1, 'ret_b': r2})
        try:
            tf = s1.TextFrame2
            if tf.HasText:
                tr = tf.TextRange
                txt = str(tr.Text)
                w = max(largura_pt(x, tr.Font.Name, tr.Font.Bold == -1, float(tr.Font.Size)) for x in txt.split('\r'))
                livre = r1[2] - float(tf.MarginLeft) - float(tf.MarginRight)
                if not tf.WordWrap and w > livre + 1:
                    achados.append({'tipo': 'texto_da_forma_cortado', 'forma': n1, 'texto': txt[:60],
                                    'texto_pt': round(w), 'livre_pt': round(livre)})
        except Exception:                         # noqa: BLE001
            pass
    # ---- grafico por cima de celula preenchida; botao (inclusive a barra de navegacao) por cima do TEXTO de uma celula
    for n1, t1, r1, s1 in formas:
        if n1.startswith('qcAssin_'):
            continue
        if t1 != 3 and any(cruzam(r1, g[2]) and r1[0] >= g[2][0] - 1 and r1[1] >= g[2][1] - 1 for g in graficos):
            continue
        if t1 == 3:
            fora = valores_fora_do_eixo(s1.Chart)
            if fora:
                achados.append({'tipo': 'grafico_corta_valor', 'forma': n1, 'dados_min_max': fora[:2],
                                'eixo_min_max': fora[2:]})
            try:
                bloco = ws.Range(s1.TopLeftCell, s1.BottomRightCell)
                vb = bloco.Value
                vb = vb if isinstance(vb, tuple) else ((vb,),)
            except Exception:                     # noqa: BLE001
                continue
            for i, linha in enumerate(vb):
                for j, x in enumerate(linha):
                    if x in (None, ''):
                        continue
                    cel = ws.Cells(bloco.Row + i, bloco.Column + j)
                    rc = (float(cel.Left) + 2, float(cel.Top) + 2, max(float(cel.Width) - 4, 1), max(float(cel.Height) - 4, 1))
                    if cruzam(r1, rc) and not ws.Rows(cel.Row).Hidden and not ws.Columns(cel.Column).Hidden:
                        achados.append({'tipo': 'grafico_sobre_celulas', 'forma': n1,
                                        'celula': cel.Address.replace('$', ''), 'valor': str(x)[:40]})
            continue
        for rt in textos_sob(ws, [r1]):
            achados.append({'tipo': 'forma_sobre_texto', 'forma': n1, 'texto_em': [round(v) for v in rt]})
    # ---- celulas
    vals = ws.Range(ws.Cells(1, 1), ws.Cells(ult_l, ult_c)).Value
    for li in range(1, ult_l + 1):
        if ws.Rows(li).Hidden:
            continue
        for co in range(1, ult_c + 1):
            v = vals[li - 1][co - 1] if vals else None
            if v in (None, ''):
                continue
            if ws.Columns(co).Hidden:
                continue
            c = ws.Cells(li, co)
            t = str(c.Text)
            if t and set(t) == {'#'}:
                achados.append({'tipo': 'numero_####', 'celula': c.Address.replace('$', ''), 'valor': str(v)[:30]})
                continue
            if not isinstance(v, str):
                continue
            area = c.MergeArea
            if area.Cells(1, 1).Address != c.Address:
                continue
            larg = float(area.Width) - 4
            f = c.Font
            w = max(largura_pt(x, f.Name, bool(f.Bold), float(f.Size or 11)) for x in t.split('\n')) if t else 0
            if w <= larg:
                continue
            if area.Cells(1, 1).ShrinkToFit and w <= larg * 1.3:
                continue                           # o Excel reduz a fonte ate caber (ate 30%: continua legivel)
            if c.WrapText:
                linhas = 0
                for par in t.split('\n'):
                    linhas += max(1, int(-(-largura_pt(par, f.Name, bool(f.Bold), float(f.Size or 11)) // max(larg, 1))))
                if linhas * float(f.Size or 11) * 1.25 > float(area.Height) + 2:
                    achados.append({'tipo': 'texto_quebrado_sem_altura', 'celula': c.Address.replace('$', ''),
                                    'texto': t[:70], 'linhas': linhas, 'altura_pt': float(area.Height)})
                continue
            prox_col = area.Column + area.Columns.Count
            vizinha = ws.Cells(li, prox_col)
            ocupada = prox_col <= ult_c + 1 and str(vizinha.Formula or '') != ''
            alinh = c.HorizontalAlignment
            if ocupada or area.Columns.Count > 1 or alinh in (-4108, -4152, 7):   # centro/direita: corta dos dois lados
                achados.append({'tipo': 'texto_cortado', 'celula': c.Address.replace('$', ''), 'texto': t[:70],
                                'texto_pt': round(w), 'largura_pt': round(larg)})
    return achados


def exportar_png(ws, pasta, ult_l, ult_c):
    import pypdfium2 as pdfium
    pdf = os.path.join(pasta, f'{ws.Name}.pdf')
    png = os.path.join(pasta, f'{ws.Name}.png')
    try:
        ws.PageSetup.PrintArea = ws.Range(ws.Cells(1, 1), ws.Cells(ult_l, ult_c)).Address
        ws.PageSetup.Orientation = 2
        ws.PageSetup.Zoom = False
        ws.PageSetup.FitToPagesWide = 1
        ws.PageSetup.FitToPagesTall = 1
        ws.ExportAsFixedFormat(0, pdf, 0, True, False)     # xlTypePDF, qualidade padrao, sem abrir
        doc = pdfium.PdfDocument(pdf)
        img = doc[0].render(scale=2.2).to_pil()
        img.save(png)
        doc.close()
        os.remove(pdf)
        return png
    except Exception as e:                         # noqa: BLE001
        return f'falhou: {e!r}'


def main(produto, caminho, pasta, abas=None):
    os.makedirs(pasta, exist_ok=True)
    q = QA(produto, caminho)
    res = {}
    try:
        q.ex.xl.ScreenUpdating = True
        for nome in (abas or TELAS):
            try:
                ws = q.wb.Worksheets(nome)
            except Exception:                      # noqa: BLE001
                continue
            vis = ws.Visible
            ws.Visible = -1
            ws.Unprotect(ii.SENHA)
            ur = ws.UsedRange
            ult_l = ur.Row + ur.Rows.Count - 1
            ult_c = ur.Column + ur.Columns.Count - 1
            for sh in ws.Shapes:                   # a imagem inclui as formas (botoes alem do conteudo)
                try:
                    if sh.Visible:
                        ult_l = max(ult_l, sh.BottomRightCell.Row)
                        ult_c = max(ult_c, sh.BottomRightCell.Column)
                except Exception:                  # noqa: BLE001
                    pass
            ult_l, ult_c = min(ult_l, LIM_LIN), min(ult_c, LIM_COL)
            t0 = time.time()
            achados = auditar_aba(ws)
            png = exportar_png(ws, pasta, ult_l, ult_c)
            res[nome] = {'achados': achados, 'imagem': png, 's': round(time.time() - t0, 1)}
            print(f'{nome:24s} {len(achados):3d} achado(s)  {png}', flush=True)
            ws.Visible = vis
    finally:
        q.fechar(salvar=False)
    with open(os.path.join(pasta, f'auditoria_{produto}.json'), 'w', encoding='utf-8') as f:
        json.dump(res, f, ensure_ascii=False, indent=1, default=str)
    return 1 if any(v['achados'] for v in res.values()) else 0


if __name__ == '__main__':
    a = sys.argv[1:]
    sys.exit(main(a[0], os.path.abspath(a[1]), os.path.abspath(a[2]), a[3].split(';') if len(a) > 3 else None))
