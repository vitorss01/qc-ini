# -*- coding: utf-8 -*-
"""instalar_adr074.py -- ADR-074: aba Analitos com UM BLOCO POR LOTE (lote, validade, Media/DP, limites,
anos de referencia CLIA/VB, especificacoes, ET e Sigma do lote). Fase 1: Bioquimica.

Uso:  python instalar_adr074.py <Bioquimica|Hematologia> <arquivo.xlsm> [--sem-salvar]

Pedido do usuario (14/10/2026): cada lote cadastrado na Configuracao com o seu bloco na aba Analitos (lote e
validade herdados do cadastro, linha 2 / B2:D2 / F2:H2; ANO DE REFERENCIA do CLIA em D47:E47 e da Variacao
Biologica em J47:K47; proximo bloco na linha 56), botao para abrir/recolher cada bloco, e os limites e o
Sigma de cada lote calculados com a referencia daquele lote.

Idempotente. So:
  1. instala todo o VBA atual das fontes (codigo_atual.py; inclui o mLotesBlocos); EXIGE compilacao;
  2. Hematologia: so o VBA (fase 2: a Hematologia ainda nao tem a biblioteca anual de especificacoes);
  3. Bioquimica, MIGRACAO (uma vez): a aba Analitos vira "Especificações" (o motor: catalogo, escolha da fonte,
     biblioteca anual 2025-2029, lote em analise) -- o Excel ajusta sozinho as formulas, nomes e graficos que
     a liam -- e nasce a aba "Analitos" dos blocos, no mesmo lugar;
  4. Especificações: lista de anos da biblioteca (lstAnosEspec), anos em uso do lote em analise (anoCLIAemUso,
     anoVBemUso = os anos do BLOCO do lote, ou o ano padrao S2), K/M/N leem o ano CLIA e O/P o ano VB do lote
     (antes: um ano so, S2, para todos); S2 passa a ser o ANO PADRAO do lote novo;
  5. Analitos: mLotesBlocos.MontarBlocos refaz os blocos (formato e formulas atuais) preservando os anos ja
     escolhidos; linha 1 com titulo, barra de navegacao e os botoes Abrir todos / Recolher todos / Lote em
     analise / Especificacoes; linha 1 e coluna A congeladas; aberto so o bloco do lote em analise;
  6. FUMACA; devolve a protecao de cada aba e a da estrutura (a aba nova recebe a da antiga Analitos).
Nao cadastra nem calcula Media/DP de lote nenhum.
"""
import os
import sys

AQUI = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, AQUI)
sys.path.insert(0, os.path.join(AQUI, '..', 'etapa1_multilote'))
import xlh  # noqa: E402
import instalar_integracao as ii  # noqa: E402
from instalar_seguranca_usuarios import foto_protecao, devolver_protecao  # noqa: E402

ESP = 'Especificações'
BLOCOS = 'Analitos'
ANOS_LIB = (46, 55, 64, 73, 82)        # celulas B do ano de cada bloco da biblioteca (Especificações!B46:AO89)
COL_LISTA = 'AT'                       # Especificações!AT46:AT50 = lista dos anos (validacao dos blocos)
R0, PASSO, NAN = 2, 54, 40


def log(*a):
    print(*a, flush=True)


def nome(wb, nm, ref):
    alvo = ref.replace("'", '').replace('$', '').upper()
    for n in list(wb.Names):
        if n.Name == nm:
            if str(n.RefersTo).replace("'", '').replace('$', '').upper() == alvo:
                return False
            n.Delete()
            break
    wb.Names.Add(nm, ref)
    return True


def migrar(wb):
    """A aba Analitos (motor) passa a chamar Especificações; nasce a Analitos dos blocos antes dela."""
    if ii.sheet(wb, ESP) is not None:
        if ii.sheet(wb, BLOCOS) is None:
            raise SystemExit('Especificações existe mas a aba Analitos dos blocos nao: estado inesperado')
        return False
    a = wb.Worksheets(BLOCOS)
    if str(a.Range('A3').Value or '').strip() != 'Analito':
        raise SystemExit(f'Analitos!A3 = {a.Range("A3").Value!r}: nao e a aba do motor, nada alterado')
    cor_aba = a.Tab.Color
    a.Name = ESP
    try:
        novo = wb.Worksheets.Add(a)                 # Before = a antiga: a nova fica no mesmo lugar
    except Exception:
        novo = wb.Worksheets.Add(None, wb.Worksheets(wb.Worksheets.Count))
    novo.Name = BLOCOS
    # aba criada por automacao so ganha CodeName depois que o projeto VBA e consultado (codigo_atual precisa dele)
    _ = wb.VBProject.Name, wb.VBProject.VBComponents.Count
    if not novo.CodeName:
        raise SystemExit('a aba Analitos nova ficou sem CodeName: o modulo da aba nao seria instalado')
    try:
        if cor_aba not in (None, False, -4142):
            novo.Tab.Color = cor_aba
    except Exception:
        pass
    novo.Cells.Font.Name = 'Segoe UI'
    novo.Cells.Font.Size = 9
    return True


def motor(wb):
    """Especificações: lista de anos, anos em uso do lote em analise, K..P pelos anos do lote."""
    e = wb.Worksheets(ESP)
    feitos = []
    nossos = ('anos (lista)', 'CLIA em uso (lote em análise)', 'VB em uso (lote em análise)')
    usado = [str(e.Range(f'{COL_LISTA}{r}').Formula or '') for r in range(45, 56)]
    if any(f and not f.startswith(('=$B$', '=LET(')) and f not in nossos for f in usado):
        raise SystemExit(f'Especificações!{COL_LISTA}45:{COL_LISTA}55 ocupada: {usado}')
    e.Range(f'{COL_LISTA}45').Value = 'anos (lista)'
    e.Range(f'{COL_LISTA}45').Font.Size = 8
    e.Range(f'{COL_LISTA}45').Font.Italic = True
    for i, r in enumerate(ANOS_LIB):
        e.Range(f'{COL_LISTA}{46 + i}').Formula = f'=$B${r}'
    nome(wb, 'lstAnosEspec', f"='{ESP}'!${COL_LISTA}$46:${COL_LISTA}$50")
    # anos em uso do LOTE EM ANALISE = os do bloco dele (sem bloco ou sem ano: o ano padrao S2). Em celula, e nao
    # em nome: nome definido com LET e recusado pela automacao do Excel.
    for nm, c, r, rot in (('anoCLIAemUso', 'D', 52, 'CLIA em uso (lote em análise)'),
                          ('anoVBemUso', 'J', 54, 'VB em uso (lote em análise)')):
        e.Range(f'{COL_LISTA}{r}').Value = rot
        e.Range(f'{COL_LISTA}{r}').Font.Size = 8
        e.Range(f'{COL_LISTA}{r}').Font.Italic = True
        e.Range(f'{COL_LISTA}{r + 1}').Formula2 = (
            f'=LET(lin,MATCH(""&loteParam,{BLOCOS}!$B:$B,0),ano,IFERROR(INDEX({BLOCOS}!${c}:${c},lin+45),""),'
            f'IF(OR(ISERROR(lin),ano=""),espAno,ano))')
        e.Range(f'{COL_LISTA}{r + 1}').Font.Bold = True
        nome(wb, nm, f"='{ESP}'!${COL_LISTA}${r + 1}")
    n = 0
    for c, nm in (('K', 'anoCLIAemUso'), ('M', 'anoCLIAemUso'), ('N', 'anoCLIAemUso'),
                  ('O', 'anoVBemUso'), ('P', 'anoVBemUso')):
        for r in range(4, 44):
            f = str(e.Range(f'{c}{r}').Formula)
            if '$S$2' in f:
                e.Range(f'{c}{r}').Formula = f.replace('$S$2', nm)
                n += 1
            elif f and nm not in f:
                raise SystemExit(f'Especificações!{c}{r}: formula inesperada (sem $S$2 nem {nm}): {f[:120]}')
    feitos.append(f'{n} formulas K..P passaram a ler o ano do lote' if n else 'K..P ja leem o ano do lote')
    if str(e.Range('Q2').Value or '').upper().startswith('ESPECIFICA'):
        e.Range('Q2').Value = 'ANO PADRÃO (lote novo)'
    e.Range('A1').Value = 'ESPECIFICAÇÕES · CATÁLOGO E REFERÊNCIAS  (até 40)'
    e.Range('A2').Formula = ('="LOTE EM ANÁLISE "&loteParam&"  ·  Média/DP e especificações DESTE lote (CLIA "&anoCLIAemUso&'
                             '" · VB "&anoVBemUso&"). Para ver ou editar QUALQUER lote, use a aba Analitos (um bloco por lote).'
                             '  ·  Lote em uso p/ lançamentos: "&loteAtivo')
    return feitos


def linha1(wb, ex):
    """Titulo, barra de navegacao e os botoes da aba dos blocos."""
    import ux
    import tema
    ws = wb.Worksheets(BLOCOS)
    ws.Range('A1:H1').UnMerge()
    ws.Range('A1:H1').Merge()
    ws.Range('A1').Value = 'ANALITOS · MÉDIA/DP, LIMITES E ESPECIFICAÇÕES POR LOTE'
    f = ws.Range('A1').Font
    f.Name, f.Size, f.Bold, f.Color = 'Segoe UI', 13, True, tema.cor('#3CD6B0')
    ws.Range('A1').VerticalAlignment = -4108
    ws.Range('A1:BD1').Interior.Color = tema.cor('#13282C')
    ws.Rows(1).RowHeight = 30
    extras = [('nav_blkabrir', 'mLotesBlocos.ExpandirTodos', '▼ Abrir todos', 84),
              ('nav_blkrecolher', 'mLotesBlocos.RecolherTodos', '▶ Recolher todos', 96),
              ('nav_blklote', 'mLotesBlocos.IrLoteEmAnalise', 'Lote em análise ▸', 100),
              ('nav_blkespec', 'mLotesBlocos.IrEspecificacoes', 'Especificações ▸', 100)]
    ux.barra_navegacao(ws, extras)


def congelar(wb, ex):
    ws = wb.Worksheets(BLOCOS)
    ws.Activate()
    janela = wb.Application.ActiveWindow
    janela.FreezePanes = False
    janela.ScrollRow = 1
    janela.ScrollColumn = 1
    ws.Range('B2').Select()
    janela.FreezePanes = True
    janela.Zoom = 90
    ws.Range('B2').Select()


def blocos(wb):
    ws = wb.Worksheets(BLOCOS)
    out, h = [], R0
    while str(ws.Cells(h, 1).Value or '').strip().upper() == 'LOTE':
        out.append((str(ws.Cells(h, 2).Value or '').strip(), h))
        h += PASSO
    return out


def cadastro(wb):
    vistos, out = set(), []
    for v in wb.Names('regLoteCol').RefersToRange.Value:
        s = v[0]
        if s in (None, ''):
            continue
        s = str(int(s)) if isinstance(s, float) and s == int(s) else str(s).strip()
        if s and s.upper() not in vistos:
            vistos.add(s.upper())
            out.append(s)
    return out


def num(v):
    return isinstance(v, (int, float)) and not isinstance(v, bool) and not (isinstance(v, int) and v < -2146820000)


def fumaca(wb, ex):
    ruins = []
    ws, e, ls = wb.Worksheets(BLOCOS), wb.Worksheets(ESP), wb.Worksheets('LotesStore')
    bl = blocos(wb)
    cad = cadastro(wb)
    if [b[0] for b in bl][:len(cad)] != cad:
        ruins.append(('blocos != cadastro', [b[0] for b in bl][:6], cad[:6]))
    for nm in ('etpOficial', 'cvtpOficial', 'espFonte', 'espVigente', 'lstAnalitos', 'aInput', 'espAno'):
        try:
            if ESP not in str(wb.Names(nm).RefersTo):
                ruins.append(('nome fora da Especificações', nm, wb.Names(nm).RefersTo))
        except Exception:
            ruins.append(('nome ausente', nm))
    anoC = wb.Names('anoCLIAemUso').RefersToRange.Value
    anoV = wb.Names('anoVBemUso').RefersToRange.Value
    if not num(anoC) or not num(anoV):
        ruins.append(('anos em uso', anoC, anoV))
    if 'anoCLIAemUso' not in str(e.Range('K4').Formula) or 'anoVBemUso' not in str(e.Range('O4').Formula):
        ruins.append(('K4/O4 sem os anos do lote', e.Range('K4').Formula[:80]))
    # valores x store, limites, especificacao = motor no lote em analise
    ult = ls.Cells(ls.Rows.Count, 16).End(-4162).Row
    store = {}
    for linha in ls.Range(f'A2:P{max(2, ult)}').Value:
        if linha[15]:
            store[str(linha[15]).upper()] = linha[2:8]
    lote_p = str(wb.Names('loteParam').RefersToRange.Value or '').strip()
    conferidos = 0
    for lote, h in bl:
        vals = ws.Range(ws.Cells(h + 4, 1), ws.Cells(h + 43, 56)).Value
        for i, linha in enumerate(vals):
            an = linha[0]
            if not an:
                continue
            esperado = store.get(f'{lote}|{an}'.upper(), (None,) * 6)
            for j in range(6):
                v, s = linha[4 + j], esperado[j]
                if (num(s) and not (num(v) and abs(v - s) < 1e-12)) or (not num(s) and v not in ('', None)):
                    ruins.append(('E:J != LotesStore', lote, an, j, v, s))
                    break
            m, dp = linha[4], linha[5]
            if num(m) and num(dp) and dp > 0:
                lim = linha[24:31]
                exp = [m + k * dp for k in range(-3, 4)]
                if not all(num(a) and abs(a - b) < 1e-9 for a, b in zip(lim, exp)):
                    ruins.append(('limites N1', lote, an, lim, exp))
                conferidos += 1
            if any(isinstance(x, int) and x < -2146820000 for x in linha):
                ruins.append(('erro no bloco', lote, an))
        if lote == lote_p:
            for i in range(NAN):
                an = e.Range(f'A{4 + i}').Value
                if not an:
                    continue
                for cb, ce in (('K', 'K'), ('O', 'O'), ('P', 'P'), ('T', 'T'), ('U', 'U')):
                    vb, ve = ws.Range(f'{cb}{h + 4 + i}').Value, e.Range(f'{ce}{4 + i}').Value
                    vb = vb if num(vb) else str(vb or '')
                    ve = ve if num(ve) and ve != 0 else ('' if ce in 'KOP' and ve in (0, None, '') else str(ve or ''))
                    if num(vb) and num(ve):
                        if abs(vb - ve) > 1e-9:
                            ruins.append(('bloco != motor', lote, an, cb, vb, ve))
                    elif str(vb) != str(ve):
                        ruins.append(('bloco != motor', lote, an, cb, vb, ve))
    if conferidos == 0:
        ruins.append(('nenhum lote com Media/DP para conferir os limites',))
    for lote, h in bl[:3]:
        if ws.Cells(h + 45, 4).Value in (None, '') or ws.Cells(h + 45, 10).Value in (None, ''):
            ruins.append(('ano de referencia vazio', lote))
        nomes = [s.Name for s in ws.Shapes]
        if f'blkTgl_{h}' not in nomes or f'blkCIQ_{h}' not in nomes:
            ruins.append(('botoes do bloco', lote))
    if ruins:
        raise SystemExit(f'fumaca ADR-074 falhou: {ruins[:6]}')
    return len(bl), conferidos, anoC, anoV


def main(produto, caminho, salvar=True):
    caminho = os.path.abspath(caminho)
    ex = xlh.Excel()
    log(f'EXCEL_PID {ex.pid}')
    try:
        wb = ex.abrir(caminho)
        if wb.ReadOnly:
            raise SystemExit(f'arquivo aberto SOMENTE LEITURA (outra instancia o segura): {caminho}')
        estrutura = bool(wb.ProtectStructure)
        foto = foto_protecao(wb)
        ii.desproteger_tudo(wb)
        ex.xl.Calculation = -4135
        import codigo_atual

        if not produto.startswith('Bio'):
            log('1. VBA (todo o codigo atual das fontes: codigo_atual.py)')
            codigo_atual.instalar(ex, wb, produto, log)
            log('2. Hematologia: blocos por lote na fase 2 (sem biblioteca anual de especificacoes ainda); so o VBA')
        else:
            migrou = migrar(wb)
            if migrou:
                foto[ESP] = foto.get(BLOCOS)               # a protecao da antiga Analitos vai com ela
            foto[BLOCOS] = foto.get(ESP) or foto.get(BLOCOS)
            log(f'1. {"MIGRACAO: Analitos -> Especificações + Analitos nova (blocos)" if migrou else "ja migrada"}')
            log('2. VBA (todo o codigo atual das fontes: codigo_atual.py)')
            codigo_atual.instalar(ex, wb, produto, log)
            for f in motor(wb):
                log(f'3. Especificações: {f}; lstAnosEspec, anoCLIAemUso, anoVBemUso')
            ex.xl.Calculation = -4105
            ex.xl.Calculate()
            ex.esperar()
            r = ex.run(f"'{wb.Name}'!mLotesBlocos.MontarBlocos", 2, teto=900)
            if not str(r).startswith('OK|'):
                raise SystemExit(f'MontarBlocos: {r}')
            log(f'4. Analitos: {str(r).split("|")[1]} bloco(s) montado(s)')
            linha1(wb, ex)
            ex.xl.Calculate()
            ex.esperar()
            # aberto so o bloco do lote em analise
            ex.run(f"'{wb.Name}'!mLotesBlocos.RecolherTodos", teto=120)
            lote_p = str(wb.Names('loteParam').RefersToRange.Value or '').strip()
            h = dict(blocos(wb)).get(lote_p)
            if h:
                ex.run(f"'{wb.Name}'!mLotesBlocos.AbrirBloco", h, True, teto=120)
            congelar(wb, ex)
            nb, nconf, anoC, anoV = fumaca(wb, ex)
            log(f'5. fumaca OK: {nb} blocos = cadastro; {nconf} analito(s) com Media/DP conferidos (E:J = LotesStore, '
                f'limites = media +-k DP); lote em analise {lote_p}: bloco = motor (CLIA {anoC}, VB {anoV})')
            n2 = ex.run(f"'{wb.Name}'!mLotesBlocos.SincronizarBlocos", teto=120)
            if int(n2 or 0) != 0:
                raise SystemExit(f'SincronizarBlocos criou {n2} bloco(s) logo depois da montagem')
            wb.Worksheets(BLOCOS).Range('B2').Select()

        devolver_protecao(wb, foto)
        if estrutura:
            wb.Protect(ii.SENHA, True, False)
        if salvar:
            wb.Save()
            log('salvo')
    finally:
        ex.fechar()


if __name__ == '__main__':
    a = sys.argv[1:]
    main(a[0], a[1], salvar='--sem-salvar' not in a)
