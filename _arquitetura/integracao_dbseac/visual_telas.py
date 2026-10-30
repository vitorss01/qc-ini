# -*- coding: utf-8 -*-
"""ADR-087 (parte visual): nada cortado nem sobreposto nas telas -- chamado pelo instalar_adr087 (arquivo aberto, abas
desprotegidas). So formato: nenhum valor, formula ou regra muda.

Guiado pela PROPRIA auditoria (testes/auditoria_abas.auditar_aba): so mexe no que ela aponta, entao rodar de novo nao
muda nada (idempotente) e o que sobrar aparece no relatorio da auditoria.
  1. grades de dados (tabelas do Excel e as grades da Estatistica e dos Eventos_Westgard): largura de cada coluna pelo
     conteudo real (AutoFit do Excel nas linhas de dados, com teto), cabecalho com quebra de linha e altura que cabe,
     texto acima do teto quebra na propria linha (grades pequenas);
  2. barra de navegacao por cima de conteudo da linha 2 -> a linha 1 cresce ate a barra;
  3. formas sobrepostas: controle giratorio encurta ate o botao; forma que invade a barra vai para depois dela;
  4. grafico por cima de celulas preenchidas (resumos do CEQ) -> vai para a direita do conteudo;
  5. rotulo cortado: ate 15% a mais -> reduzir para caber (ShrinkToFit); mesclado -> quebra de linha com altura;
     senao a coluna alarga (ate 150 pt) e o resto quebra;
  6. aba Registros: faixa de ajuda legivel, cabecalho da tabela no padrao das tabelas do Painel, botao Limpar filtros.
Formas sao recolocadas onde estavam depois de mexer em larguras (as livres nao andam junto com as colunas).
"""
import os
import re
import sys

AQUI = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(AQUI, 'testes'))
import auditoria_abas as aud  # noqa: E402


def cor(r, g, b):
    return r + g * 256 + b * 65536


TEAL = cor(30, 110, 98)          # cabecalho das tabelas do Painel
CINZA = cor(89, 89, 89)
TETO_COL = 260                   # pt: coluna de grade nao passa disso (o resto quebra)
TETO_ROTULO = 150                # pt: rotulo fora de grade alarga a coluna ate isso
NAO_ALARGAR = {'Painel', 'Início'}   # telas desenhadas a mao: formas livres alinhadas as colunas
MOVER_GRAFICOS = {'EQA.CAP_Resumo', 'EQA.Controllab_Resumo'}
ENCOLHER = {'Eventos_Westgard'}  # coluna A larga demais (titulo antigo): pode estreitar ate o conteudo


def _formas(ws):
    out = []
    for sh in ws.Shapes:
        try:
            out.append((sh, int(sh.Placement), float(sh.Width), float(sh.Height)))
        except Exception:                          # noqa: BLE001
            pass
    return out


def _recolocar(fs):
    """Depois de mexer em larguras: a forma que acompanha as celulas (Placement 1/2) anda com elas (certo), mas a de
    'mover e redimensionar' (2) esticaria junto -- volta so o TAMANHO. Grafico estica de proposito: fica."""
    for sh, plc, w, h in fs:
        try:
            if plc == 2 and int(sh.Type) != 3 and (abs(sh.Width - w) > 0.5 or abs(sh.Height - h) > 0.5):
                sh.Width, sh.Height = w, h
        except Exception:                          # noqa: BLE001
            pass


def crescer_linha(ws, r, altura):
    """Muda a altura da linha r e leva junto as formas LIVRES (Placement 3) que estao abaixo dela (as outras o Excel
    ja leva): sem isto um botao livre ficava por cima do texto que desceu."""
    antes = float(ws.Rows(r).RowHeight)
    fundo = float(ws.Rows(r).Top) + antes
    altura = min(409, altura)
    if abs(altura - antes) < 0.5:
        return False
    livres = [s for s in ws.Shapes if int(s.Placement) == 3 and float(s.Top) >= fundo - 0.5]
    ws.Rows(r).RowHeight = altura
    d = float(ws.Rows(r).RowHeight) - antes
    for s in livres:
        s.Top = float(s.Top) + d
    return True


def _txt(c):
    try:
        return str(c.Text or '')
    except Exception:                              # noqa: BLE001
        return ''


def _larg_texto(c, t=None):
    t = _txt(c) if t is None else t
    f = c.Font
    return max((aud.largura_pt(x, f.Name, bool(f.Bold), float(f.Size or 11)) for x in t.split('\n')), default=0)


def _linhas(texto, largura, f):
    n = 0
    for par in texto.split('\n'):
        atual, k = '', 1
        for p in par.split(' '):
            cand = (atual + ' ' + p).strip()
            if aud.largura_pt(cand, f.Name, bool(f.Bold), float(f.Size or 11)) <= largura or not atual:
                atual = cand
            else:
                k += 1
                atual = p
        n += k
    return n


def _por_largura_pt(ws, c, pt):
    """ColumnWidth (em caracteres) que da 'pt' de largura -- a razao pt/caractere medida na propria coluna."""
    col = ws.Columns(c)
    # a largura em pt nao e proporcional a ColumnWidth (ha uma margem fixa em pixels): sem corrigir, a coluna fica
    # alguns pt aquem e a rodada seguinte alarga de novo. Repete a razao ate chegar (converge em 2-3 passos)
    for _ in range(4):
        w0, p0 = float(col.ColumnWidth), float(col.Width)
        if p0 <= 0 or abs(p0 - pt) < 1.0:
            return
        col.ColumnWidth = min(255, max(0.5, w0 * pt / p0))


def grade(ws, hr, c0, c1, r1, mud, encolher=False):
    """Uma grade de dados: cabecalho na linha hr, colunas c0..c1, dados ate r1."""
    if r1 <= hr:
        r1 = hr + 1
    ate = min(r1, hr + 3000)
    fs = _formas(ws)
    quebrar = []
    for c in range(c0, c1 + 1):
        col = ws.Columns(c)
        if col.Hidden:
            continue
        w0, p0 = float(col.ColumnWidth), float(col.Width)
        corpo = ws.Range(ws.Cells(hr + 1, c), ws.Cells(ate, c))
        if encolher and hr > 2:                     # a coluna pode estreitar: o que fica acima do cabecalho tambem conta
            corpo = ws.Application.Union(ws.Range(ws.Cells(2, c), ws.Cells(hr - 1, c)), corpo)
        tem_dado = ws.Application.WorksheetFunction.CountA(corpo) > 0
        corpo.Columns.AutoFit()                     # (coluna sem dado: o AutoFit nao mexe -- so o cabecalho conta)
        p_auto = float(col.Width) if tem_dado else 0.0
        cab = ws.Cells(hr, c)
        palavras = [p for p in re.split(r'[\s\n]+', _txt(cab)) if p]
        p_cab = max((_larg_texto(cab, p) for p in palavras), default=0) + 10
        quer = max(p_auto + 2 if tem_dado else 0.0, p_cab)
        if not encolher:
            quer = max(quer, p0)
        elif quer < p0 and p0 - quer <= 3:
            quer = p0                               # estreitar so vale a pena com folga (nao fica oscilando)
        if quer > TETO_COL:
            if p_auto > TETO_COL:
                quebrar.append(c)
            quer = max(TETO_COL, p0 if not encolher else 0)
        col.ColumnWidth = w0                       # volta e aplica so a decisao (AutoFit mexeu)
        if quer - p0 > 2.5 or p0 - quer > 3:          # a largura anda em pixels: margem contra a 'subida' a cada rodada
            _por_largura_pt(ws, c, quer)
            mud.append(f'{ws.Name} coluna {aud_col(c)} {p0:.0f}->{quer:.0f} pt')
    # cabecalho: quebra e altura para o maior numero de linhas
    cabs = ws.Range(ws.Cells(hr, c0), ws.Cells(hr, c1))
    precisa = 0.0
    for c in range(c0, c1 + 1):
        cel = ws.Cells(hr, c)
        t = _txt(cel)
        if not t or ws.Columns(c).Hidden:
            continue
        n = _linhas(t, float(cel.MergeArea.Width) - 6, cel.Font)
        precisa = max(precisa, n * float(cel.Font.Size or 11) * 1.3 + 4)
    if not cabs.WrapText:
        cabs.WrapText = True
        mud.append(f'{ws.Name} cabecalho (linha {hr}) com quebra')
    if precisa > float(ws.Rows(hr).RowHeight) + 0.5:
        crescer_linha(ws, hr, precisa)
        cabs.VerticalAlignment = -4108
        mud.append(f'{ws.Name} cabecalho (linha {hr}) altura {precisa:.0f} pt')
    if quebrar and (r1 - hr) <= 1500:
        for c in quebrar:
            corpo = ws.Range(ws.Cells(hr + 1, c), ws.Cells(r1, c))
            if not corpo.WrapText:
                corpo.WrapText = True
                mud.append(f'{ws.Name} coluna {aud_col(c)} quebra o texto longo')
        ws.Range(ws.Rows(hr + 1), ws.Rows(r1)).AutoFit()
    _recolocar(fs)


def aud_col(n):
    s = ''
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def _ultima_col(ws, linha, ate=80):
    v = ws.Range(ws.Cells(linha, 1), ws.Cells(linha, ate)).Value[0]
    ult = 0
    for i, x in enumerate(v, 1):
        if x not in (None, ''):
            ult = i
    return ult


def grades(ws, mud):
    los = list(ws.ListObjects)
    if los:
        for lo in los:
            hr = lo.HeaderRowRange.Row
            c0 = lo.Range.Column
            c1 = c0 + lo.Range.Columns.Count - 1
            r1 = lo.Range.Row + lo.Range.Rows.Count - 1
            grade(ws, hr, c0, c1, r1, mud, encolher=ws.Name in ENCOLHER)
        return
    marcas = {'Estatística': ('Analito',), 'Eventos_Westgard': ('RUN', 'Analito')}.get(ws.Name)
    if not marcas:
        return
    for hr in range(1, 25):
        v = [str(x or '').strip() for x in ws.Range(ws.Cells(hr, 1), ws.Cells(hr, 60)).Value[0]]
        if all(m in v for m in marcas):
            c0 = next(i for i, x in enumerate(v, 1) if x)
            c1 = _ultima_col(ws, hr, 60)
            r1 = int(ws.Cells(ws.Rows.Count, c0).End(-4162).Row)
            grade(ws, hr, c0, c1, r1, mud, encolher=ws.Name in ENCOLHER)
            return


def barra_nav(ws, mud):
    """A barra do alto (nav_* e os botoes ao lado dela) nao pode cobrir texto:
    - na linha 1 (cabecalho de tabela embaixo da barra, como nos Dados do CEQ): o texto vai para o pe da celula e a
      linha cresce o bastante para ele ficar abaixo da barra;
    - nas linhas de baixo: a linha 1 cresce ate o fim da barra."""
    nav = [s for s in ws.Shapes if str(s.Name).startswith('nav_')]
    if not nav:
        return
    fim1 = float(ws.Rows(1).Top) + float(ws.Rows(1).RowHeight)
    faixa = [s for s in ws.Shapes if int(s.Type) != 3 and s.Visible and float(s.Top) < fim1 - 0.5]
    rets = [(float(s.Left), float(s.Top), float(s.Width), float(s.Height)) for s in faixa]
    fundo = max(r[1] + r[3] for r in rets)
    sob, h_txt = [], 0.0
    for c in range(1, _ultima_col(ws, 1, 80) + 1):
        cel = ws.Cells(1, c)
        if cel.MergeArea.Cells(1, 1).Address != cel.Address:
            continue
        rt = aud.ret_texto(cel)
        if rt and any(aud.cruzam(r, rt) for r in rets):
            sob.append(cel)
        if rt:
            h_txt = max(h_txt, rt[3])
    if sob:
        # a linha toda no pe (uniforme): o cabecalho inteiro fica abaixo da barra
        ws.Range(ws.Cells(1, 1), ws.Cells(1, max(1, _ultima_col(ws, 1, 80)))).VerticalAlignment = -4107
        crescer_linha(ws, 1, fundo + 3 + h_txt + 3 - float(ws.Rows(1).Top))
        mud.append(f'{ws.Name} linha 1: texto abaixo da barra de navegacao ({len(sob)} celula(s))')
        return
    esq = min(r[0] for r in rets)
    dir_ = max(r[0] + r[2] for r in rets)
    r = 2
    invade = False
    while float(ws.Rows(r).Top) < fundo - 0.5 and r < 6:
        for c in range(1, 60):
            cel = ws.Cells(r, c)
            if float(cel.Left) > dir_:
                break
            if float(cel.Left) + float(cel.Width) > esq and _txt(cel):
                invade = True
        r += 1
    if invade:
        crescer_linha(ws, 1, fundo + 4 - float(ws.Rows(1).Top))
        mud.append(f'{ws.Name} linha 1 cresce ate a barra de navegacao ({fundo + 4:.0f} pt)')


def formas(ws, mud):
    fs = [s for s in ws.Shapes if s.Visible and int(s.Type) != 3]
    nav = [s for s in fs if str(s.Name).startswith('nav_')]
    fim_nav = max((float(s.Left) + float(s.Width) for s in nav), default=0)
    for a in fs:
        for b in fs:
            if a is b or str(a.Name) >= str(b.Name):
                continue
            ra = (float(a.Left), float(a.Top), float(a.Width), float(a.Height))
            rb = (float(b.Left), float(b.Top), float(b.Width), float(b.Height))
            if not aud.cruzam(ra, rb):
                continue
            na, nb = str(a.Name), str(b.Name)
            if na.startswith('qcAssin_') or nb.startswith('qcAssin_'):
                continue
            giro = [s for s in (a, b) if int(s.Type) == 8 and str(s.Name).lower().startswith('spinner')]
            if giro:
                g = giro[0]
                o = b if g is a else a
                if float(o.Top) > float(g.Top):
                    g.Height = float(o.Top) - float(g.Top) - 1
                    mud.append(f'{ws.Name} {g.Name} termina antes de {o.Name}')
                    continue
            if na.startswith('nav_') != nb.startswith('nav_'):
                o = b if na.startswith('nav_') else a
                o.Left = fim_nav + 8
                mud.append(f'{ws.Name} {o.Name} sai de baixo da barra de navegacao')


def graficos(ws, mud):
    if ws.Name not in MOVER_GRAFICOS:
        return
    for sh in ws.Shapes:
        if int(sh.Type) != 3:
            continue
        fora = aud.valores_fora_do_eixo(sh.Chart)
        if fora:
            ax = sh.Chart.Axes(2)
            if fora[0] >= 0:
                ax.MinimumScale = 0                # contagem: o eixo comeca no zero
            else:
                ax.MinimumScaleIsAuto = True
            ax.MaximumScaleIsAuto = True
            mud.append(f'{ws.Name} grafico {sh.Name}: eixo cortava barras (max {fora[1]:g} acima do teto {fora[3]:g})')
        r0, r1 = sh.TopLeftCell.Row, sh.BottomRightCell.Row
        c0, c1 = sh.TopLeftCell.Column, sh.BottomRightCell.Column
        vals = ws.Range(ws.Cells(r0, c0), ws.Cells(r1, c1)).Value
        if not any(x not in (None, '') for linha in vals for x in linha):
            continue
        ult = max(_ultima_col(ws, r, 60) for r in range(r0, r1 + 1))
        novo = float(ws.Cells(1, ult + 1).Left) + 12
        if abs(float(sh.Left) - novo) > 0.5:
            sh.Left = novo
            mud.append(f'{ws.Name} grafico {sh.Name} para depois da coluna {aud_col(ult)} (nao cobre mais celulas)')


def rotulos(ws, mud):
    for a in aud.auditar_aba(ws):
        if a['tipo'] == 'numero_####':
            c = ws.Range(a['celula'])
            col = ws.Columns(c.Column)
            w0, p0 = float(col.ColumnWidth), float(col.Width)
            c.Columns.AutoFit()
            if float(col.Width) < p0:
                col.ColumnWidth = w0
            else:
                mud.append(f'{ws.Name} coluna {aud_col(c.Column)} alarga para o numero de {a["celula"]}')
            continue
        if a['tipo'] not in ('texto_cortado', 'texto_quebrado_sem_altura'):
            continue
        c = ws.Range(a['celula'])
        area = c.MergeArea
        larg = float(area.Width) - 4
        w = _larg_texto(c)
        tam = float(c.Font.Size or 11)
        n = _linhas(_txt(c), larg - 2, c.Font)
        cabe_quebrando = n * tam * 1.25 <= float(area.Height) + 1
        if a['tipo'] == 'texto_cortado' and w <= larg * 1.15:
            area.ShrinkToFit = True
            mud.append(f'{ws.Name}!{a["celula"]} reduz para caber')
            continue
        if ws.Name in NAO_ALARGAR:
            # tela desenhada a mao: a altura das linhas fica (botoes livres estao alinhados a elas)
            if cabe_quebrando:
                area.WrapText = True
                mud.append(f'{ws.Name}!{a["celula"]} quebra em {n} linha(s) na altura que ja tem')
            elif w <= larg * 1.3:
                area.WrapText = False
                area.ShrinkToFit = True
                mud.append(f'{ws.Name}!{a["celula"]} reduz para caber')
            else:
                area.WrapText = True
                crescer_linha(ws, c.Row, float(ws.Rows(c.Row).RowHeight) + n * tam * 1.25 + 2 - float(area.Height))
                mud.append(f'{ws.Name}!{a["celula"]} quebra em {n} linha(s) (linha mais alta)')
            continue
        if area.Columns.Count > 1 or a['tipo'] == 'texto_quebrado_sem_altura':
            area.WrapText = True
            if not cabe_quebrando:
                crescer_linha(ws, c.Row, float(ws.Rows(c.Row).RowHeight) + n * tam * 1.3 + 4 - float(area.Height))
            mud.append(f'{ws.Name}!{a["celula"]} quebra em {n} linha(s)')
            continue
        fs = _formas(ws)
        quer = min(w + 8, TETO_ROTULO)
        if quer > float(area.Width):
            _por_largura_pt(ws, c.Column, quer)
        if w + 8 > TETO_ROTULO:
            c.WrapText = True
            n = _linhas(_txt(c), float(c.Width) - 6, c.Font)
            if n * tam * 1.25 > float(c.Height) + 1:
                crescer_linha(ws, c.Row, n * tam * 1.3 + 4)
        _recolocar(fs)
        mud.append(f'{ws.Name} coluna {aud_col(c.Column)} alarga para {a["celula"]}')


GRUPOS = {'Início': [lambda n: n.startswith('tile_') and n != 'tile_guias', lambda n: n == 'tile_guias']}


def formas_sobre_texto(ws, mud):
    """Botao/forma livre por cima de texto: desce (em grupo, quando sao varios lado a lado, como os blocos do Inicio)
    ate o texto ficar livre. Sem saida em 4 tentativas = volta como estava e fica no relatorio da auditoria."""
    achados = [a for a in aud.auditar_aba(ws) if a['tipo'] == 'forma_sobre_texto']
    if not achados:
        return
    nomes = {a['forma'] for a in achados if not a['forma'].startswith('nav_')}
    feitos = set()
    for n in sorted(nomes):
        if n in feitos:
            continue
        regra = next((g for g in GRUPOS.get(ws.Name, []) if g(n)), None)
        grupo = [s for s in ws.Shapes if (regra(str(s.Name)) if regra else str(s.Name) == n) and s.Visible]
        feitos |= {str(s.Name) for s in grupo}
        if any(int(s.Placement) != 3 for s in grupo):
            continue                               # forma presa as celulas: anda com elas, nao se move a mao
        tops = [float(s.Top) for s in grupo]
        for _ in range(4):
            rets = [(float(s.Left), float(s.Top), float(s.Width), float(s.Height)) for s in grupo]
            cob = aud.textos_sob(ws, rets)
            if not cob:
                break
            d = max(r[1] + r[3] for r in cob) + 6 - min(float(s.Top) for s in grupo)
            for s in grupo:
                s.Top = float(s.Top) + max(d, 1)
        rets = [(float(s.Left), float(s.Top), float(s.Width), float(s.Height)) for s in grupo]
        if aud.textos_sob(ws, rets):
            for s, t0 in zip(grupo, tops):
                s.Top = t0
            continue
        mud.append(f'{ws.Name} {", ".join(sorted(str(s.Name) for s in grupo))} desce(m) para nao cobrir texto')


# UTF-8 lido como latin-1: so as sequencias conhecidas (um 'Ã' legitimo, como em CALIBRAÇÃO, fica intacto)
MOJIBAKE = {'auditÃvel': 'auditável', 'Ã­': 'í', 'Ã¡': 'á', 'Ã©': 'é', 'Ã§': 'ç', 'Ã£': 'ã', 'Ã³': 'ó', 'Ãº': 'ú',
            'Ãª': 'ê', 'Ã´': 'ô', 'Ãµ': 'õ', 'Ã¢': 'â', 'Ã ': 'à', 'Ã‰': 'É', 'Ã‡': 'Ç', 'Ã“': 'Ó'}


def titulos(ws, mud):
    """Texto fixo com acentuacao corrompida (ex.: 'EstatÃ­stico') nas linhas do titulo."""
    for r in range(1, 4):
        for c in range(1, 12):
            cel = ws.Cells(r, c)
            f = str(cel.Formula or '')
            if 'Ã' not in f or f.startswith('='):
                continue
            novo = f
            for a, b in MOJIBAKE.items():
                novo = novo.replace(a, b)
            if novo != f:
                cel.Value = novo
                mud.append(f'{ws.Name}!{cel.Address.replace("$", "")} acentuacao corrigida')


def _int(v, padrao=-1):
    try:
        return int(v)
    except (TypeError, ValueError):
        return padrao


def painel(ws, mud):
    """C4 (contador de analitos) passava por baixo do controle giratorio: o texto termina antes dele e quebra."""
    giro = [s for s in ws.Shapes if int(s.Type) == 8 and str(s.Name).lower().startswith('spinner')]
    c4 = ws.Range('C4').MergeArea
    if not giro:
        return
    g = giro[0]
    if float(c4.Left) + float(c4.Width) <= float(g.Left) + 0.5:
        return
    ult = c4.Column
    while ult + 1 <= c4.Column + c4.Columns.Count - 1 and float(ws.Cells(4, ult + 1).Left) + float(
            ws.Cells(4, ult + 1).Width) <= float(g.Left) + 0.5:
        ult += 1
    f = ws.Range('C4').Formula
    fmt = (ws.Range('C4').Font.Name, ws.Range('C4').Font.Size, ws.Range('C4').Font.Color, ws.Range('C4').HorizontalAlignment)
    c4.UnMerge()
    novo = ws.Range(ws.Cells(4, 3), ws.Cells(4, ult))
    if ult > 3:
        novo.Merge()
    ws.Range('C4').Formula = f
    novo.Font.Name, novo.Font.Size, novo.Font.Color = fmt[0], fmt[1], fmt[2]
    novo.HorizontalAlignment = fmt[3]
    novo.WrapText = True
    novo.VerticalAlignment = -4108
    mud.append(f'Painel C4 termina antes do controle giratorio (C4:{aud_col(ult)}4, quebra)')


def registros(ws, mud):
    lo = ws.ListObjects('tblRegistros') if any(str(x.Name) == 'tblRegistros' for x in ws.ListObjects) else None
    a2 = ws.Range('A2').MergeArea
    cinza_a2 = cor(102, 102, 102)                  # o mesmo cinza que o instalar_adr087 grava (senao alternam a cada rodada)
    if _int(ws.Range('A2').Interior.Pattern) != -4142 or _int(ws.Range('A2').Font.Color) != cinza_a2:
        a2.Interior.Pattern = -4142
        a2.Font.Color = cinza_a2
        mud.append('Registros A2 legivel (sem a faixa escura)')
    faixa = ws.Range('A3:A5')
    if _int(faixa.Interior.Pattern) != -4142:
        faixa.Interior.Pattern = -4142
        ws.Range('A3').Font.Color = cor(102, 102, 102)
        mud.append('Registros A3:A5 sem a faixa escura (rotulo FILTROS visivel)')
    if lo is not None:
        cab = lo.HeaderRowRange
        if _int(cab.Interior.Color) != TEAL or str(cab.Font.Name) != 'Segoe UI':
            cab.Interior.Color = TEAL
            cab.Font.Name, cab.Font.Size, cab.Font.Bold, cab.Font.Color = 'Segoe UI', 9, True, 0xFFFFFF
            cab.VerticalAlignment = -4108
            mud.append('Registros cabecalho da tabela no padrao das tabelas do Painel')
        corpo = lo.DataBodyRange
        if corpo is not None and str(corpo.Font.Name) != 'Segoe UI':
            corpo.Font.Name, corpo.Font.Size = 'Segoe UI', 9
            mud.append('Registros corpo da tabela em Segoe UI 9')
    try:
        b = ws.Shapes('btnLimparFiltrosReg')
        if float(b.Height) < 16:
            b.Height = 17
            b.Top = float(ws.Rows(3).Top) + max(0.0, (float(ws.Rows(3).RowHeight) - 17) / 2)
            mud.append('Registros botao Limpar filtros com altura de botao')
    except Exception:                              # noqa: BLE001
        pass


def aplicar(wb, log=print, telas=None):
    """Aplica nas telas; devolve a lista de mudancas (vazia = ja estava tudo certo)."""
    mud = []
    for nome in (telas or aud.TELAS):
        try:
            ws = wb.Worksheets(nome)
        except Exception:                          # noqa: BLE001
            continue
        if nome == 'Registros':
            registros(ws, mud)
        if nome == 'Painel':
            painel(ws, mud)
        titulos(ws, mud)
        grades(ws, mud)
        formas(ws, mud)                            # antes da barra: o botao que sai de baixo dela muda o que ela cobre
        barra_nav(ws, mud)
        rotulos(ws, mud)
        graficos(ws, mud)
        formas_sobre_texto(ws, mud)
    log(f'   visual: {len(mud)} ajuste(s)')
    return mud
