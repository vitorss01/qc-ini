# -*- coding: utf-8 -*-
"""v2.py -- camada de DESEMPENHO (ADR-050), aplicada sobre a Etapa 1.

O que muda, e por que (medido com 5 anos de banco, ~100 mil resultados):

  Calc C/F/AB(/AX)  liam data e valor do BANCO por MAXIFS/SUMIFS -- 540
                    formulas varrendo 100 mil linhas a cada troca de analito.
                    Agora leem o que o motor publicou em Eng_Saida (engData,
                    engValN1..N3), por posicao. Plotam exatamente o que o motor
                    avaliou (ADR-049, decisao 4).
  Liberacao A:B     600 formulas AGGREGATE/MINIFS sobre o banco, refeitas a cada
                    gravacao. Viram valor, escrito por mLotes.AtualizarListaLiberacao
                    (chamada por mBanco.AtualizarFlagsBanco e por TrocarLote).
  VBA               mApp (casca: tela, calculo, eventos, status), indice do
                    banco, eventos de Westgard em cache por lote, filtro lido
                    uma vez, TrocarLote sem CalculateFull.
"""
import os

AQUI = os.path.dirname(os.path.abspath(__file__))
SRC = os.path.join(AQUI, 'src')

# coluna de valor do nivel t (1-based) no Calc, e o nome do valor no Eng_Saida
CALC_VAL = {1: 'F', 2: 'AB', 3: 'AX'}
ENG_VAL = {1: 'Y', 2: 'Z', 3: 'AA'}


def codigo(nome_arq, subst=None):
    with open(os.path.join(SRC, nome_arq), encoding='utf-8', newline='') as f:
        txt = f.read().replace('\r\n', '\n')
    for a, b in (subst or []):
        if txt.count(a) != 1:
            raise SystemExit(f'{nome_arq}: ancora de produto nao encontrada (1x): {a[:70]!r}')
        txt = txt.replace(a, b)
    linhas = txt.split('\n')
    return '\r\n'.join(l for l in linhas if not l.startswith('Attribute ')).strip('\r\n') + '\r\n'


def conferir_duplicados(arquivos):
    """Nome de macro repetido = "Ambiguous name detected": o projeto VBA inteiro
    para de compilar e QUALQUER macro passa a responder "nao e possivel executar".
    Sem Excel visivel isso vira uma caixa de dialogo invisivel e a automacao fica
    esperando para sempre -- foi o que aconteceu em 20/09/2026, por um corte errado
    num modulo. Dois segundos de conferencia aqui evitam quinze minutos de travamento.
    """
    import re
    vistos = {}
    for arq, subst in arquivos:
        txt = codigo(arq, subst)
        nomes = re.findall(r'^(?:Public\s+|Private\s+)?(?:Static\s+)?(?:Sub|Function|Property\s+\w+)\s+(\w+)',
                           txt, re.M)
        priv = set(re.findall(r'^Private\s+(?:Sub|Function)\s+(\w+)', txt, re.M))
        for n in nomes:
            if nomes.count(n) > 1:
                raise SystemExit(f'{arq}: {n} declarado {nomes.count(n)}x no mesmo modulo')
            if n in priv:
                continue                      # privado nao colide entre modulos
            if n.lower() in vistos and vistos[n.lower()] != arq:
                raise SystemExit(f'{n}: publico em {vistos[n.lower()]} e em {arq} '
                                 f'-- "Ambiguous name detected" derruba o projeto inteiro')
            vistos[n.lower()] = arq
    return len(vistos)


def garantir_modulo(vbp, nome):
    for c in vbp.VBComponents:
        if c.Name == nome:
            return c
    c = vbp.VBComponents.Add(1)          # vbext_ct_StdModule
    c.Name = nome
    return c


def substituir_codigo(vbp, comp, texto):
    cm = garantir_modulo(vbp, comp).CodeModule
    n = cm.CountOfLines
    if n:
        cm.DeleteLines(1, n)
    cm.AddFromString(texto)


def nome(wb, n, ref):
    try:
        wb.Names(n).Delete()
    except Exception:
        pass
    wb.Names.Add(n, ref)
    return wb.Names(n).RefersTo


def aplicar_planilhas(ex, wb, nlv):
    """Nomes engData/engValN*, Calc lendo o motor, Liberacao A:B como valor."""
    sh = {s.Name: s for s in wb.Worksheets}
    calc, eng, lib = sh['Calc'], sh['Eng_Saida'], sh['Liberação']
    nome(wb, 'engData', '=Eng_Saida!$AC$3:$AC$182')
    for t in range(1, nlv + 1):
        c = ENG_VAL[t]
        nome(wb, f'engValN{t}', f'=Eng_Saida!${c}$3:${c}$182')
    if eng.Range('AC2').Value in (None, ''):
        eng.Range('AC2').Value = 'Data'
    eng.Range('AC3:AC182').NumberFormat = 'dd/mm/yyyy hh:mm'

    # Calc: data e valores por nivel = o que o motor publicou, por posicao
    f = calc.Range('C3').Formula
    if 'engData' not in f:
        if 'MAXIFS(rData' not in f:
            raise SystemExit(f'Calc!C3 formato inesperado: {f[:120]}')
        calc.Range('C3:C182').Formula = '=IF($B3="","",IF(INDEX(engData,$A3)="","",INDEX(engData,$A3)))'
    for t in range(1, nlv + 1):
        col = CALC_VAL[t]
        f = calc.Range(f'{col}3').Formula
        if f'engValN{t}' in f:
            continue
        if 'SUMIFS(rValor' not in f or f'rNivel,{t},' not in f:
            raise SystemExit(f'Calc!{col}3 formato inesperado: {f[:120]}')
        calc.Range(f'{col}3:{col}182').Formula = (
            f'=IF($B3="","",IF(INDEX(engValN{t},$A3)="","",INDEX(engValN{t},$A3)))')

    # Liberacao A:B: de formula sobre o banco para valor escrito pelo VBA
    f = lib.Range('A4').Formula
    if isinstance(f, str) and f.startswith('=') and 'rRUN' not in f:
        raise SystemExit(f'Liberacao!A4 formato inesperado: {f[:120]}')
    ex.run('mLotes.AtualizarListaLiberacao', teto=300)
    if isinstance(lib.Range('A4').Formula, str) and lib.Range('A4').Formula.startswith('='):
        raise SystemExit('Liberacao!A4 ainda e formula depois de AtualizarListaLiberacao')
