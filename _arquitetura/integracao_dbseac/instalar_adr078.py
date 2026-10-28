# -*- coding: utf-8 -*-
"""instalar_adr078.py -- ADR-078: as linhas do alvo (Media e -3DP..+3DP) aparecem sem depender das corridas.

Uso:  python instalar_adr078.py <Bioquimica|Hematologia> <arquivo.xlsm> [--sem-salvar]

O defeito (relato do usuario, 15/10/2026): trocou o lote 8974 pelo 8973 no Painel, com a Media/DP digitada, e o
grafico nao mostrou nada. A Media/DP chegava a LotesStore e ao Calc (AX1..BA1 iguais ao digitado). O que falhava era
o desenho: as 7 linhas de cada nivel usavam como X a coluna das corridas (Calc!E3:E182, NA() sem corrida) e as
colunas de limite exigiam corrida no periodo (D = 1). Lote sem corrida (8973: 0 resultados) ou sem corrida no filtro
= grafico vazio; com 1 corrida, 1 ponto e nenhuma linha. Um lote por arquivo NAO resolveria.

Idempotente (a 2a execucao registra 0 mudancas). So:
  1. instala todo o VBA atual das fontes (codigo_atual.py); EXIGE compilacao;
  2. Calc, linhas 183:185 (area nova, nada a referencia): por nivel e por linha (-3s..+3s), DOIS pontos --
     x = bordas do eixo (xrunmin/xrunmax), y = Media + k*DP do lote. As linhas 3..182 nao mudam;
  3. Calc: sem corrida, a moldura do eixo X fica fixa de 0 a 21 (20 corridas, CLSI C24); com corridas, igual;
  4. Calc: alvoTeste{i} (o nivel do lote em tela esta com valor de TESTE do ADR-076); B3:B182 compara o
     equipamento sem diferenca de maiusculas/espacos;
  5. LotesStore X:Z (TESTE_N1..N3; nomes lsTesteN1..3) + migracao das linhas cuja ORIGEM diz TESTE;
  6. Eng_Saida: nomes engNLoteBase (W1, o VBA escreve), engAviso (Y1), engFiltroRotina (AA1);
  7. Painel: status por nivel (J7..), Q4, A39 e H4 dizem o caso (AGUARDANDO 1a CORRIDA, SEM CORRIDA NO PERIODO,
     SEM MEDIA/DP, alvo TESTE, filtro ajustado) e as formatacoes passam a "comeca com";
  8. graficos do Painel: as 7 series de limite de cada nivel passam a ler a area 183:185 -- SO a formula da serie
     (tipo, cor, traco, espessura e legenda ficam como estao);
  9. FUMACA no lote em tela; devolve a protecao de cada aba e a da estrutura como estavam.
"""
import os
import re
import sys

AQUI = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, AQUI)
sys.path.insert(0, os.path.join(AQUI, '..', 'etapa1_multilote'))
import xlh  # noqa: E402
import instalar_integracao as ii  # noqa: E402
from instalar_seguranca_usuarios import foto_protecao, devolver_protecao  # noqa: E402

LIM = ['-3s', '-2s', '-1s', 'Média', '+1s', '+2s', '+3s']
SUF = ['m3', 'm2', 'm1', 'med', 'p1', 'p2', 'p3']
KS = [-3, -2, -1, 0, 1, 2, 3]
R_ROT, R_X0, R_X1 = 183, 184, 185
LS_TESTE = 24                              # LotesStore X
ROT_183 = ('LINHAS DO ALVO (ADR-078): 2 pontos por linha; x = bordas do eixo (xrunmin/xrunmax); '
           'y = Média + k·DP do lote; não dependem das corridas')
MUDANCAS = []
# motor de outro equipamento (Calc!B e o status): os dois lados pelo TRIM do Excel; <> ja ignora maiusculas
EQ_DIF = 'TRIM(""&engEquip)<>TRIM(""&selEquipamento)'


def log(*a):
    print(*a, flush=True)


def col(n):
    s = ''
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def ncol(s):
    n = 0
    for ch in s:
        n = n * 26 + ord(ch) - 64
    return n


def norm(f):
    return str(f or '').replace('_xlfn.', '').replace(' ', '').upper()


def setf(rng, f, oque):
    """Escreve a formula so se difere (idempotente); conta a mudanca."""
    if norm(rng.Formula) != norm(f):
        rng.Formula = f
        MUDANCAS.append(oque)


def setv(rng, v, oque):
    if str(rng.Value if rng.Value is not None else '') != str(v):
        rng.Value = v
        MUDANCAS.append(oque)


def nome(wb, nm, ref):
    alvo = ref.replace("'", '').replace('$', '').upper()
    for n in list(wb.Names):
        if n.Name == nm:
            if str(n.RefersTo).replace("'", '').replace('$', '').upper() == alvo:
                return
            n.Delete()
            break
    wb.Names.Add(nm, ref)
    MUDANCAS.append(f'nome {nm}')


def endereco_nome(wb, nm):
    r = wb.Names(nm).RefersToRange
    return r.Worksheet.Name, str(r.Address)


# ---------------------------------------------------------------- validacoes (antes de escrever)
def mapa(wb, nlv):
    c = wb.Worksheets('Calc')
    cab = [str(v or '').strip() for v in c.Range('A2:DZ2').Value[0]]
    m = {'niveis': []}
    lin1 = [str(c.Cells(1, k).Formula or '') for k in range(1, 131)]
    for i in range(1, nlv + 1):
        cols = []
        for s in SUF:
            h = f'N{i} {s}'
            if h not in cab:
                raise SystemExit(f'Calc: cabecalho "{h}" nao encontrado na linha 2')
            cols.append(cab.index(h) + 1)
        if cols != list(range(cols[0], cols[0] + 7)):
            raise SystemExit(f'Calc: colunas de limite do N{i} fora de ordem: {cols}')
        cm = next((k + 1 for k, f in enumerate(lin1) if f'>0,INDEX(lsMedN{i},' in f.replace(' ', '')), 0)
        cs = next((k + 1 for k, f in enumerate(lin1) if f'>0,INDEX(lsDPN{i},' in f.replace(' ', '')), 0)
        if not cm or not cs:
            raise SystemExit(f'Calc: celulas de Media/DP do N{i} (linha 1, INDEX(lsMedN{i}/lsDPN{i})) nao encontradas')
        m['niveis'].append({'cols': cols, 'M': f'${col(cm)}$1', 'S': f'${col(cs)}$1'})
    kc = next((k + 1 for k, f in enumerate(lin1) if 'LSCHAVE' in f.upper().replace(' ', '')), 0)
    if not kc:
        raise SystemExit('Calc: celula da chave do lote (MATCH(...,lsChave)) nao encontrada na linha 1')
    m['chave'] = f'${col(kc)}$1'
    m['dec'] = f'${col(kc + 1)}$1'
    m['teste'] = [kc + 1 + i for i in range(1, nlv + 1)]
    for i, k in enumerate(m['teste'], 1):
        f = str(c.Cells(1, k).Formula or '')
        if f and 'LSTESTEN' not in f.upper():
            raise SystemExit(f'Calc!{col(k)}1 ocupada ({f[:60]}): nao e o alvoTeste{i}')
    for nm in ('xrunmin', 'xrunmax'):
        sh, ad = endereco_nome(wb, nm)
        if sh != 'Calc':
            raise SystemExit(f'nome {nm} fora do Calc ({sh}!{ad})')
        m[nm] = ad
    for nm in ('engLote', 'engPainel', 'engNTotal', 'engAnalito', 'engEquip', 'loteAnalise', 'selAnalito'):
        try:
            wb.Names(nm)
        except Exception:
            raise SystemExit(f'nome {nm} nao existe')
    # linhas 183:185 vazias ou ja do ADR-078
    a183 = str(c.Range(f'A{R_ROT}').Value or '')
    usado = [str(v or '') for row in c.Range(f'A{R_ROT}:DZ{R_X1}').Formula for v in row]
    if a183 and not a183.startswith('LINHAS DO ALVO (ADR-078)'):
        raise SystemExit(f'Calc!A{R_ROT} ocupada: {a183[:60]}')
    if not a183 and any(usado):
        raise SystemExit('Calc: linhas 183:185 tem conteudo que nao e do ADR-078')
    e = wb.Worksheets('Eng_Saida')
    for a, ok in (('X1', ('', 'aviso:')), ('Z1', ('', 'filtro da rotina:')), ('V1', ('', 'lote na base:'))):
        if str(e.Range(a).Value or '') not in ok:
            raise SystemExit(f'Eng_Saida!{a} ocupada: {e.Range(a).Value!r}')
    ls = wb.Worksheets('LotesStore')
    for k, h in enumerate(('TESTE_N1', 'TESTE_N2', 'TESTE_N3')):
        v = str(ls.Cells(1, LS_TESTE + k).Value or '')
        if v not in ('', h):
            raise SystemExit(f'LotesStore!{col(LS_TESTE + k)}1 ocupada: {v!r}')
    p = wb.Worksheets('Painel')
    if p.ChartObjects().Count != nlv:
        raise SystemExit(f'Painel: {p.ChartObjects().Count} graficos; esperados {nlv}')
    for i in range(1, nlv + 1):
        ch = p.ChartObjects(i).Chart
        t = ch.ChartTitle.Text if ch.HasTitle else ''
        if f'Nível {i}' not in t:
            raise SystemExit(f'Painel: o grafico {i} nao e o do Nivel {i} (titulo {t!r})')
    return m


# ---------------------------------------------------------------- Calc
def calc(wb, m, nlv):
    c = wb.Worksheets('Calc')
    setv(c.Range(f'A{R_ROT}'), ROT_183, 'Calc A183')
    setv(c.Range(f'A{R_X0}'), 'borda esq.', 'Calc A184')
    setv(c.Range(f'A{R_X1}'), 'borda dir.', 'Calc A185')
    xmin, xmax = m['xrunmin'], m['xrunmax']
    setf(c.Range(f'E{R_X0}'), f'=IF(ISNUMBER({xmin}),{xmin},NA())', 'Calc E184')
    setf(c.Range(f'E{R_X1}'), f'=IF(ISNUMBER({xmax}),{xmax},NA())', 'Calc E185')
    for i, nv in enumerate(m['niveis'], 1):
        M, S = nv['M'], nv['S']
        for cc, k in zip(nv['cols'], KS):
            y = f'{M}' if k == 0 else f'{M}+({k})*{S}'
            f = f'=IF(OR({M}="",{S}=""),NA(),{y})'
            for r in (R_X0, R_X1):
                setf(c.Cells(r, cc), f, f'Calc {col(cc)}{r}')
        nome(wb, f'limN{i}', f"=Calc!${col(nv['cols'][0])}${R_X0}:${col(nv['cols'][-1])}${R_X1}")
    nome(wb, 'limX', f'=Calc!$E${R_X0}:$E${R_X1}')
    # moldura do eixo X: sem corrida, 0..21 (20 corridas); com corrida, como antes
    setf(c.Range(xmin), '=IF(COUNT($E$3:$E$182)=0,0,AGGREGATE(5,6,$E$3:$E$182)-0.6)', 'xrunmin')
    setf(c.Range(xmax), '=IF(COUNT($E$3:$E$182)=0,21,AGGREGATE(4,6,$E$3:$E$182)+2.2)', 'xrunmax')
    # nivel com valor de TESTE (ADR-076) no lote em tela
    for i, k in enumerate(m['teste'], 1):
        M = m['niveis'][i - 1]['M']
        setf(c.Cells(1, k), f'=IF(OR({m["chave"]}=0,{M}=""),FALSE,N(INDEX(lsTesteN{i},{m["chave"]}))>0)',
             f'Calc {col(k)}1 alvoTeste{i}')
        nome(wb, f'alvoTeste{i}', f'=Calc!${col(k)}$1')
    # B3:B182: o equipamento em tela comparado sem diferenca de maiusculas/espacos (Hematologia: EQUIPAMENTO_PADRAO)
    rg = c.Range('B3:B182')
    fs = [r[0] for r in rg.Formula]
    antigos = ['(""&engEquip)<>(""&selEquipamento)', '(""&engEquip)<>UPPER(TRIM(""&selEquipamento))']
    novo = []
    for f in fs:
        g = f
        for a in antigos:
            g = g.replace(a, EQ_DIF)
        if EQ_DIF not in g:
            raise SystemExit('Calc!B3:B182: comparacao do equipamento fora do padrao esperado')
        novo.append(g)
    if novo != fs:
        rg.Formula = tuple((f,) for f in novo)
        MUDANCAS.append('Calc B3:B182 equipamento normalizado')


# ---------------------------------------------------------------- LotesStore e Eng_Saida
def lotesstore(wb):
    ls = wb.Worksheets('LotesStore')
    for k, h in enumerate(('TESTE_N1', 'TESTE_N2', 'TESTE_N3')):
        setv(ls.Cells(1, LS_TESTE + k), h, f'LotesStore {h}')
        nome(wb, f'lsTesteN{k + 1}', f'=LotesStore!${col(LS_TESTE + k)}$2:${col(LS_TESTE + k)}$4001')
    ult = ls.Cells(ls.Rows.Count, 16).End(-4162).Row
    n = 0
    if ult >= 2:
        rg = ls.Range(ls.Cells(2, 1), ls.Cells(ult, LS_TESTE + 2))
        v = [list(r) for r in rg.Value]
        mudou = False
        for r in v:
            org = str(r[22] or '')
            if not org.upper().startswith('TESTE'):
                continue
            segs = {s.strip()[:3]: s for s in org.split(' -- ', 1)[-1].split(' | ')}
            for nv in range(1, 4):
                seg, ct = segs.get(f'N{nv}:', ''), LS_TESTE - 1 + nv - 1
                bits = (1 if 'media de ' in seg else 0) | (2 if 'DP = ' in seg else 0)
                atual = r[ct] if isinstance(r[ct], (int, float)) and not isinstance(r[ct], bool) else 0
                # vazio, ou o 1 da 1a versao desta migracao: passa a dizer qual campo e de teste
                if bits and (r[ct] in (None, '') or atual == 1) and atual != bits:
                    r[ct] = bits
                    mudou = True
                    n += 1
        if mudou:
            ls.Range(ls.Cells(2, LS_TESTE), ls.Cells(ult, LS_TESTE + 2)).Value = \
                tuple(tuple(r[LS_TESTE - 1:LS_TESTE + 2]) for r in v)
            MUDANCAS.append(f'LotesStore: {n} nivel(is) marcados TESTE')
    return n


def eng_saida(wb):
    e = wb.Worksheets('Eng_Saida')
    setv(e.Range('V1'), 'lote na base:', 'Eng_Saida V1')
    setv(e.Range('X1'), 'aviso:', 'Eng_Saida X1')
    setv(e.Range('Z1'), 'filtro da rotina:', 'Eng_Saida Z1')
    nome(wb, 'engNLoteBase', '=Eng_Saida!$W$1')
    nome(wb, 'engAviso', '=Eng_Saida!$Y$1')
    nome(wb, 'engFiltroRotina', '=Eng_Saida!$AA$1')


# ---------------------------------------------------------------- Painel
def status_nivel(m, i):
    # ADR-084 (28/10/2026, pedido do usuario): o status mostra so o veredito -- sem o sufixo " (alvo TESTE)". A marca
    # TESTE continua na LotesStore (lsTesteN*) e no Calc (alvoTeste); so nao aparece mais no Painel
    M = 'Calc!' + m['niveis'][i - 1]['M']
    return ('=IF(OR((""&engLote)<>(""&loteAnalise),(""&engAnalito)<>(""&selAnalito),' + EQ_DIF + '),"DESATUALIZADO",'
            'IF(N(engNLoteBase)<0,"BASE SEM DADOS",'
            f'IF({M}="","SEM MÉDIA/DP",'
            'IF(N(engNLoteBase)=0,"AGUARDANDO 1ª CORRIDA",'
            'IF(N(engNTotal)=0,"SEM CORRIDA DESTE ANALITO",'
            'IF(COUNTIF(Calc!$D$3:$D$182,1)=0,"SEM CORRIDA NO PERÍODO",'
            f'IF(INDEX(engPainel,{i},10)="","SEM PONTO NO PERÍODO",INDEX(engPainel,{i},10))))))))')


def q4(nlv):
    J = f'$J$7:$J${6 + nlv}'
    return (f'=IF(COUNTIF({J},"DESATUALIZADO*")>0,"DESATUALIZADO",'
            f'IF(COUNTIF({J},"BASE SEM DADOS*")>0,"BASE SEM DADOS",'
            f'IF(COUNTIF({J},"REJEITADO*")>0,"REJEITADO",'
            f'IF(COUNTIF({J},"SEM MÉDIA/DP*")>0,"SEM MÉDIA/DP",'
            f'IF(COUNTIF({J},"AGUARDANDO*")>0,"AGUARDANDO 1ª CORRIDA",'
            f'IF(COUNTIF({J},"SEM *")>0,INDEX({J},MATCH("SEM *",{J},0)),'
            f'IF(COUNTIF({J},"OK*")>0,"OK","—")))))))')


def a39(m, nlv):
    Ms = ['Calc!' + n['M'] for n in m['niveis']]
    Ss = ['Calc!' + n['S'] for n in m['niveis']]
    dec = 'Calc!' + m['dec']
    nomes = ['N1', 'N2', 'N3'][:nlv]
    vazio = 'OR(' + ','.join(f'{x}=""' for x in Ms) + ')'
    todos = 'AND(' + ','.join(f'{x}=""' for x in Ms) + ')'
    quais = 'TRIM(' + '&'.join(f'IF({x}="","{n} ","")' for x, n in zip(Ms, nomes)) + ')'
    alvo = '"Lote "&loteAnalise&" · alvo ' + '&" · '.join(
        f'{n} "&FIXED({x},{dec})&" ± "&FIXED({s},{dec}+1)' for n, x, s in zip(nomes, Ms, Ss)) + '&". "'
    return (f'=IF({vazio},"⛔ LOTE "&loteAnalise&" · "&selAnalito&": SEM MÉDIA/DP em "&IF({todos},"{" e ".join(nomes) if nlv == 2 else "N1, N2 e N3"}",{quais})'
            f'&" — digite a Média e o DP no bloco do lote (aba Analitos). Westgard não avaliado. ",{alvo})'
            '&IF(selAnalito="","",IF(N(engNLoteBase)<0,"A base de resultados está vazia ou não foi carregada (ATUALIZAR DADOS). ",'
            'IF(N(engNLoteBase)=0,"Lote "&loteAnalise&" sem nenhum resultado na base: as linhas mostram o alvo cadastrado (confira antes da 1ª corrida). ",'
            'IF(N(engNTotal)=0,"O lote não tem resultado de "&selAnalito&" no "&engEquip&". ",'
            'IF(COUNTIF(Calc!$D$3:$D$182,1)=0,"Nenhuma corrida deste lote no período"&IF(engLotesPer<>""," (com dados: "&engLotesPer&")","")&". ","")))))'
            '&IF(engAviso<>"","FILTRO: "&engAviso&" ","")'
            '&IF(N(engCortadas)>0,"⚠ "&engCortadas&" corrida(s) antigas do período fora do gráfico (máx. 180). ","")'
            '&IF(OR($G$3<>"",$G$4<>""),"Filtro "&IF($G$3<>"",TEXT($G$3,"dd/mm/aa"),"início")&" a "&IF($G$4<>"",TEXT($G$4,"dd/mm/aa"),"hoje"),"Sem filtro de data")')


AZUL = (7949855, 16247773)          # fonte RGB(31,78,121), fundo RGB(221,235,247)
AMBAR = (22428, 10284031)           # fonte RGB(156,87,0),  fundo RGB(255,235,156)


def regras_texto(rng, extras):
    """Regras 'igual a' (OK / REJEITADO / SEM MEDIA/DP) viram 'comeca com', com as MESMAS cores; e os estados novos."""
    fcs = rng.FormatConditions
    antigas, novas = [], {}
    for i in range(1, fcs.Count + 1):
        fc = fcs(i)
        if fc.Type == 1:
            f1 = str(fc.Formula1 or '')
            mt = re.fullmatch(r'="(.*)"', f1)
            if mt:
                semfundo = fc.Interior.ColorIndex in (-4142, None) or fc.Interior.Pattern in (-4142,)
                antigas.append((mt.group(1), fc.Font.Color, None if semfundo else fc.Interior.Color, fc.Font.Bold))
        elif fc.Type == 9:
            novas[str(fc.Text)] = 1
    for txt, *_ in antigas:
        novas.pop(txt, None)
    if antigas:
        for i in range(fcs.Count, 0, -1):
            if fcs(i).Type == 1 and re.fullmatch(r'="(.*)"', str(fcs(i).Formula1 or '')):
                fcs(i).Delete()
    regras = [(t, fo, fu, b) for t, fo, fu, b in antigas] + \
             [(t, fo, fu, True) for t, (fo, fu) in extras if t not in novas]
    for t, fo, fu, b in regras:
        if t in novas:
            continue
        fc = rng.FormatConditions.Add(Type=9, String=t, TextOperator=2)
        fc.Font.Color = fo
        if b is not None:
            fc.Font.Bold = bool(b)
        if fu is not None:
            fc.Interior.Color = fu
        fc.StopIfTrue = False
        MUDANCAS.append(f'formatacao {rng.Address} "{t}"')


def painel(wb, m, nlv):
    p = wb.Worksheets('Painel')
    for i in range(1, nlv + 1):
        setf(p.Range(f'J{6 + i}'), status_nivel(m, i), f'Painel J{6 + i}')
    setf(p.Range('Q4'), q4(nlv), 'Painel Q4')
    setf(p.Range('A39'), a39(m, nlv), 'Painel A39')
    setf(p.Range('H4'), '=IF(engAviso<>"","▲ lote em análise · ver o aviso abaixo dos gráficos","▲ lote em análise")',
         'Painel H4')
    extras = [('AGUARDANDO', AZUL), ('SEM CORRIDA', AZUL), ('SEM PONTO', AZUL), ('DESATUALIZADO', AMBAR),
              ('BASE SEM DADOS', AMBAR)]
    regras_texto(p.Range(f'J7:J{6 + nlv}'), extras)
    regras_texto(p.Range('Q4'), extras)


# ---------------------------------------------------------------- graficos
# ADR-080: o X pode ja ser a coluna da corrida do lote do nivel (Calc CR..CT) -- preservada; antes era sempre E
RX = re.compile(r'^=SERIES\((.*),Calc!\$([A-Z]+)\$(\d+):\$[A-Z]+\$(\d+),Calc!\$([A-Z]+)\$(\d+):\$([A-Z]+)\$(\d+),(\d+)\)$')


def graficos(wb, m, nlv):
    p = wb.Worksheets('Painel')
    leg = {}
    for i in range(1, nlv + 1):
        ch = p.ChartObjects(i).Chart
        cols = m['niveis'][i - 1]['cols']
        leg[i] = ch.Legend.LegendEntries().Count if ch.HasLegend else None
        achadas = set()
        for s_i in range(1, ch.SeriesCollection().Count + 1):
            s = ch.SeriesCollection(s_i)
            nm = str(s.Name)
            if nm not in LIM:
                continue
            f = str(s.Formula)
            mt = RX.match(f)
            if not mt:
                raise SystemExit(f'grafico {i}, serie {nm}: formula fora do padrao: {f}')
            esperado = col(cols[LIM.index(nm)])
            if mt.group(5) != esperado:
                raise SystemExit(f'grafico {i}, serie {nm}: Y em {mt.group(5)}, esperado {esperado}')
            xc = mt.group(2) if mt.group(2) in ('E', 'CR', 'CS', 'CT') else 'E'
            novo = f'=SERIES({mt.group(1)},Calc!${xc}${R_X0}:${xc}${R_X1},Calc!${esperado}${R_X0}:${esperado}${R_X1},{mt.group(9)})'
            if f != novo:
                s.Formula = novo
                MUDANCAS.append(f'grafico {i} serie {nm}')
            achadas.add(nm)
        if achadas != set(LIM):
            raise SystemExit(f'grafico {i}: series de limite ausentes: {set(LIM) - achadas}')
    return leg


# ---------------------------------------------------------------- fumaca
def num(v):
    return isinstance(v, (int, float)) and not isinstance(v, bool) and not (isinstance(v, int) and v < -2146820000)


def bordas_nivel(wb, i, xmin, xmax):
    """ADR-080: bordas do eixo do nivel i (xrunmin<i>/xrunmax<i>); antes dele, o eixo comum."""
    try:
        a, b = wb.Names(f'xrunmin{i}').RefersToRange.Value, wb.Names(f'xrunmax{i}').RefersToRange.Value
        if num(a) and num(b):
            return a, b
    except Exception:
        pass
    return xmin, xmax


def fumaca(wb, ex, m, nlv, leg0):
    c, p = wb.Worksheets('Calc'), wb.Worksheets('Painel')
    xmin = c.Range(m['xrunmin']).Value
    xmax = c.Range(m['xrunmax']).Value
    if not (num(xmin) and num(xmax) and xmax > xmin):
        raise SystemExit(f'fumaca: xrunmin/xrunmax nao numericos ({xmin!r}, {xmax!r})')
    ruins, conf = [], 0
    for i in range(1, nlv + 1):
        ch = p.ChartObjects(i).Chart
        M = c.Range(m['niveis'][i - 1]['M']).Value
        S = c.Range(m['niveis'][i - 1]['S']).Value
        for s_i in range(1, ch.SeriesCollection().Count + 1):
            s = ch.SeriesCollection(s_i)
            nm = str(s.Name)
            if nm not in LIM:
                continue
            xs, ys = list(s.XValues), list(s.Values)
            k = KS[LIM.index(nm)]
            if len(xs) != 2 or len(ys) != 2:
                ruins.append((i, nm, 'pontos', len(xs), len(ys)))
                continue
            bx = bordas_nivel(wb, i, xmin, xmax)          # ADR-080: cada nivel tem as suas bordas
            if not (abs(xs[0] - bx[0]) < 1e-9 and abs(xs[1] - bx[1]) < 1e-9):
                ruins.append((i, nm, 'x', xs, bx))
            if num(M) and num(S):
                if not all(num(y) and abs(y - (M + k * S)) <= 1e-9 * max(1, abs(M)) for y in ys):
                    ruins.append((i, nm, 'y', ys, M + k * S))
                conf += 1
            elif any(num(y) for y in ys):
                ruins.append((i, nm, 'y sem alvo', ys))
        if ch.HasLegend and leg0.get(i) is not None and ch.Legend.LegendEntries().Count != leg0[i]:
            ruins.append((i, 'legenda', leg0[i], ch.Legend.LegendEntries().Count))
        t = str(p.Range(f'J{6 + i}').Value or '')
        if not t:
            ruins.append((i, 'status vazio'))
    for nm in ('engNLoteBase', 'engAviso', 'engFiltroRotina', 'limX', 'alvoTeste1', 'lsTesteN1'):
        wb.Names(nm).RefersToRange
    if ruins:
        raise SystemExit(f'fumaca ADR-078 falhou: {ruins[:6]}')
    return conf, xmin, xmax


def main(produto, caminho, salvar=True):
    try:
        sys.stdout.reconfigure(errors='replace')
    except Exception:
        pass
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
        nlv = 2 if produto.startswith('Bio') else 3
        import codigo_atual
        log('1. VBA (todo o codigo atual das fontes: codigo_atual.py)')
        codigo_atual.instalar(ex, wb, produto, log)
        m = mapa(wb, nlv)
        log(f'2. validado: {nlv} niveis; limites em ' +
            ', '.join(f'N{i + 1} {col(n["cols"][0])}:{col(n["cols"][-1])} (M {n["M"]}, DP {n["S"]})'
                      for i, n in enumerate(m['niveis'])) + f'; chave {m["chave"]}')
        nt = lotesstore(wb)                 # antes do Calc: os nomes lsTesteN* existem quando a formula e escrita
        eng_saida(wb)
        log(f'3. LotesStore X:Z (TESTE_N1..3), {nt} nivel(is) migrado(s); Eng_Saida: engNLoteBase, engAviso, engFiltroRotina')
        calc(wb, m, nlv)
        log('4. Calc: linhas do alvo (183:185), moldura do eixo X sem corrida (0..21), alvoTeste, equipamento normalizado')
        painel(wb, m, nlv)
        log('5. Painel: status J7.., Q4, A39, H4 e formatacoes')
        leg0 = graficos(wb, m, nlv)
        log(f'6. graficos: as 7 series de limite de cada nivel leem Calc 184:185 (legenda antes: {leg0})')
        ex.xl.Calculation = -4105
        ex.xl.Calculate()
        ex.esperar()
        for mac in ('mEstatistica.InvalidarCache', 'mEstatistica.RecalcularAnalitoAtual', 'mUI.AtualizarEixos'):
            ex.run(f"'{wb.Name}'!{mac}", teto=600)
        ex.xl.Calculate()
        ex.esperar()
        conf, xmin, xmax = fumaca(wb, ex, m, nlv, leg0)
        log(f'7. fumaca OK: lote {wb.Names("loteSel").RefersToRange.Value}: {conf} serie(s) de limite com 2 pontos '
            f'x = ({xmin}, {xmax}) e y = Media + k DP; status {[wb.Worksheets("Painel").Range(f"J{6 + i}").Value for i in range(1, nlv + 1)]}')
        log(f'MUDANCAS {len(MUDANCAS)}' + (': ' + '; '.join(MUDANCAS[:12]) + (' ...' if len(MUDANCAS) > 12 else '') if MUDANCAS else ''))
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
