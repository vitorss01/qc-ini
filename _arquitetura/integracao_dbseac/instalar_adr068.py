# -*- coding: utf-8 -*-
"""instalar_adr068.py -- ADR-068: analito SEM ETp cadastrado mostra "-" (nunca 0 nem Sigma negativo).

Uso:  python instalar_adr068.py <Bioquimica|Hematologia> <arquivo.xlsm> [--sem-salvar]

O defeito (achado no estudo D-01, 06/10/2026): a celula de ETp vazia na aba Analitos (ex.: Lipase) chega a
Estatistica!K como 0 -- INDEX de celula vazia devolve 0 -- e o Sigma vira (0 - |bias|)/CV, NEGATIVO, com
classe, DPM e plano de CQ calculados em cima dele.

Idempotente. So:
  1. instala todo o VBA atual das fontes (codigo_atual.py: motor com ETp/Sigma "-" sem ETp); EXIGE compilacao;
  2. Estatistica: K (ETp) = "-" quando o ETp cadastrado esta vazio, nao numerico ou <= 0; e as colunas que
     dependem do ETp -- L Sigma, M status, N/O margem, P status da margem, V DPM, W rendimento, X regras,
     Y N, Z run size, AA cobertura -- devolvem "-" quando K = "-". ET% (H) nao depende do ETp e fica;
  3. Painel: Sigma por nivel (I7:I9) "-" quando o ETp do nivel e "-"; caixa "Sigma do plano" (O4) "-";
     as cores da caixa passam a exigir NUMERO (texto "-" contava como >= 4 e ficava verde);
  4. Cfg_PlanoQC!B1 (sigmaDoPlano): "-" quando nenhum nivel tem Sigma porque falta o ETp;
  5. FUMACA: toda linha com ETp ausente tem K..AA em "-" e nenhuma tem Sigma negativo por ETp 0; nenhum erro;
  6. devolve a protecao de cada aba e a da estrutura como estavam.
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
from instalar_adr064 import LIN_TAB, ultima_linha_tabela, selecionar_a1  # noqa: E402

DEPENDENTES = ['L', 'M', 'N', 'O', 'P', 'V', 'W', 'X', 'Y', 'Z', 'AA']
RX_K = re.compile(r'^=IF\(\$A(\d+)="","",IFERROR\((.+),""\)\)$', re.S)


def log(*a):
    print(*a, flush=True)


def coluna_k(e, ult):
    f = str(e.Range(f'K{LIN_TAB}').Formula)
    if '"-"' in f:
        return 'ja instalada'
    m = RX_K.match(f)
    if not m:
        raise SystemExit(f'Estatistica!K{LIN_TAB}: formula inesperada, nada alterado: {f[:160]}')
    fonte = m.group(2)                                   # INDEX(<etp>,MATCH($A14,...))
    novo = (f'=IF(OR($A{LIN_TAB}="",$A{LIN_TAB}=0),"",IFERROR(IF(N({fonte})>0,{fonte},"-"),"-"))')
    e.Range(f'K{LIN_TAB}:K{ult}').Formula = novo           # relativo: o Excel propaga linha a linha
    return 'instalada'


def dependentes(e, ult):
    feitas = []
    for c in DEPENDENTES:
        f = str(e.Range(f'{c}{LIN_TAB}').Formula)
        if f.startswith(f'=IF($K{LIN_TAB}="-","-",'):
            continue
        if not f.startswith('='):
            raise SystemExit(f'Estatistica!{c}{LIN_TAB} sem formula: {f[:80]}')
        e.Range(f'{c}{LIN_TAB}:{c}{ult}').Formula = f'=IF($K{LIN_TAB}="-","-",{f[1:]})'
        feitas.append(c)
    return feitas


def painel(wb, nlv):
    p = wb.Worksheets('Painel')
    feitos = []
    for k in range(nlv):
        r = 7 + k
        f = str(p.Range(f'I{r}').Formula)
        if not f.startswith(f'=IF($F{r}="-","-",'):
            # forma antiga: (F - |G|)/E do motor; forma do ADR-071: le o Sigma de Estatistica!L pela chave AB
            antiga = 'ABS(' in f and f'$F{r}' in f
            adr071 = '!$L$14:$L$' in f and '!$AB$14:$AB$' in f and 'selAnalito' in f
            if not (antiga or adr071):
                raise SystemExit(f'Painel!I{r}: formula inesperada: {f[:120]}')
            p.Range(f'I{r}').Formula = f'=IF($F{r}="-","-",{f[1:]})'
            feitos.append(f'I{r}')
    f = str(p.Range('O4').Formula)
    if 'sigmaDoPlano="-"' not in f:
        p.Range('O4').Formula = '=IF(sigmaDoPlano="-","-",IF(NOT(ISNUMBER(sigmaDoPlano)),"",sigmaDoPlano))'
        feitos.append('O4')
    feitos += cores_o4(p)
    return feitos


def local(ws, formula):
    """FormatConditions le a formula no idioma da instalacao: traduz por uma celula auxiliar."""
    aux = ws.Range('ZZ1000')
    antes = aux.Formula
    aux.Formula = formula
    loc = aux.FormulaLocal
    aux.Formula = antes
    return loc


def cores_o4(p):
    """Regras 'valor da celula >= 4' e '< 3' em O4 viram expressoes que exigem numero (mesmo formato)."""
    fcs = p.Range('O4').FormatConditions
    trocas = []
    for i in range(fcs.Count, 0, -1):
        fc = fcs(i)
        if fc.Type != 1:                                  # xlCellValue
            continue
        op, f1 = fc.Operator, str(fc.Formula1).lstrip('=')
        sinal = {7: '>=', 6: '<', 5: '>', 8: '<='}.get(op)  # xlGreaterEqual, xlLess, xlGreater, xlLessEqual
        if not sinal:
            continue
        # so copia o que a regra tinha de fato (ColorIndex xlNone/None = sem preenchimento ou cor propria)
        interior = fc.Interior.Color if fc.Interior.ColorIndex not in (None, -4142) else None
        fonte = fc.Font.Color if fc.Font.ColorIndex not in (None, -4142, -4105) else None
        negrito = fc.Font.Bold
        expr = f'=AND(ISNUMBER($O$4),$O$4{sinal}{f1.replace(",", ".")})'
        fc.Delete()
        nova = fcs.Add(2, None, local(p, expr))           # xlExpression
        if interior is not None:
            nova.Interior.Color = interior
        if fonte is not None:
            nova.Font.Color = fonte
        if negrito is not None:
            nova.Font.Bold = negrito
        trocas.append(f'O4 {sinal}{f1}')
    return trocas


def plano(wb, nlv):
    c = wb.Worksheets('Cfg_PlanoQC')
    f = str(c.Range('B1').Formula)
    if '"-"' in f:
        return False
    faixa = f'Painel!$I$7:$I${6 + nlv}'
    if faixa not in f:
        raise SystemExit(f'Cfg_PlanoQC!B1: formula inesperada: {f[:120]}')
    c.Range('B1').Formula = (f'=IF(COUNT({faixa})=0,IF(COUNTIF({faixa},"-")>0,"-","SEM DADOS"),MIN({faixa}))')
    return True


def fumaca(wb, e, ult, produto):
    a = wb.Worksheets('Analitos')
    sem, com, ruins = [], 0, []
    for r in range(LIN_TAB, ult + 1):
        an = e.Range(f'A{r}').Value
        if an in (None, '', 0):
            continue
        k = e.Range(f'K{r}').Value
        if k == '-':
            sem.append(str(an))
            fora = [c for c in DEPENDENTES if e.Range(f'{c}{r}').Value != '-']
            if fora:
                ruins.append((r, an, 'dependentes sem "-"', fora))
        elif isinstance(k, float):
            com += 1
            if k <= 0:
                ruins.append((r, an, 'ETp <= 0 numerico', k))
        else:
            ruins.append((r, an, 'K inesperado', k))
        for c in ['K'] + DEPENDENTES:
            if str(e.Range(f'{c}{r}').Text).startswith('#'):
                ruins.append((r, an, 'erro', c))
    if ruins:
        raise SystemExit(f'fumaca ADR-068 falhou: {ruins[:6]}')
    return sorted(set(sem)), com


def main(produto, caminho, salvar=True):
    caminho = os.path.abspath(caminho)
    nlv = 2 if produto.startswith('Bio') else 3
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

        log('1. VBA (todo o codigo atual das fontes: codigo_atual.py)')
        import codigo_atual
        codigo_atual.instalar(ex, wb, produto, log)

        e = wb.Worksheets('Estatística')
        ult = ultima_linha_tabela(e)
        log(f'2. Estatistica (linhas {LIN_TAB}..{ult}): K {coluna_k(e, ult)}; dependentes {dependentes(e, ult) or "ja instaladas"}')
        log(f'3. Painel: {painel(wb, nlv) or "ja instalado"}')
        log(f'4. Cfg_PlanoQC!B1: {"instalada" if plano(wb, nlv) else "ja instalada"}')

        log('5. fumaca')
        ex.xl.Calculation = -4105
        ex.xl.CalculateFull()
        ex.esperar()
        sem, com = fumaca(wb, e, ult, produto)
        log(f'   {com} linhas com ETp numerico; sem ETp ("-" em K..AA): {", ".join(sem) or "nenhum"}')
        selecionar_a1(e)

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
