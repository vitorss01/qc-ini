# -*- coding: utf-8 -*-
"""instalar_adr069.py -- ADR-069: escolha da especificacao do ETp na aba Analitos da BIOQUIMICA (coluna S).

Uso:  python instalar_adr069.py <Bioquimica|Hematologia> <arquivo.xlsm> [--sem-salvar]

A Bioquimica escolhe, por analito, de onde vem o ETp em uso (coluna S "ESPQ FONTE": CLIA, VB ou FAB):
  T "ETp em uso final %" (nome etpOficial)  = K (ETp CLIA) | R (ETp VB) | M (ETp FAB)
  U "CVTp %"            (nome cvtpOficial) = V (Lim CV CLIA = K/3) | W (Lim CV VB = f x CVi) | N (CVTp FAB)
Defeitos corrigidos (06/10/2026):
  1. FAB TROCADO: T devolvia N (CVTp FAB) e U devolvia M (ETp FAB);
  2. o "-" de especificacao ausente nunca disparava: K, M, N, O, P vem da tabela de especificacoes por
     INDEX, que devolve 0 (nao "") para celula vazia -- T virava 0 e o Sigma saia negativo;
  3. R (ETp VB) com so um dos dois (CVi ou CVg) dava numero (a parcela do outro zerada): agora exige os dois.
Regras: especificacao <= 0 ou vazia = ausente = "-"; linha sem analito = "".
A Hematologia nao tem escolha (fonte por prioridade: ETp cfg > CLIA > VB, colunas P/Q/R) e nao e alterada aqui.

Idempotente. So:
  1. instala todo o VBA atual das fontes (codigo_atual.py); EXIGE compilacao;
  2. Bioquimica: reescreve R, T e U das linhas 4..43 (formula relativa, igual em todas as linhas);
  3. FUMACA: para cada analito e para CADA fonte (CLIA, VB, FAB) troca S, recalcula a aba e confere T e U
     contra K/R/M e V/W/N (ou "-"), e devolve S como estava; nenhum erro;
  4. devolve a protecao de cada aba e a da estrutura como estavam.
"""
import os
import sys

AQUI = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, AQUI)
sys.path.insert(0, os.path.join(AQUI, '..', 'etapa1_multilote'))
import xlh  # noqa: E402
import instalar_integracao as ii  # noqa: E402
from instalar_seguranca_usuarios import foto_protecao, devolver_protecao  # noqa: E402

L0, L1 = 4, 43
FONTES = ('CLIA', 'VB', 'FAB')
MARCA = 'ADR-069'


def log(*a):
    print(*a, flush=True)


def formulas(r):
    t = (f'=IF($A{r}="","",IF($S{r}="CLIA",IF(N($K{r})>0,$K{r},"-"),IF($S{r}="VB",IF(N($R{r})>0,$R{r},"-"),'
         f'IF($S{r}="FAB",IF(N($M{r})>0,$M{r},"-"),"-"))))')
    u = (f'=IF($A{r}="","",IF($S{r}="CLIA",IF(N($V{r})>0,$V{r},"-"),IF($S{r}="VB",IF(N($W{r})>0,$W{r},"-"),'
         f'IF($S{r}="FAB",IF(N($N{r})>0,$N{r},"-"),"-"))))')
    return t, u


def coluna_r(a):
    """R (ETp VB): exige CVi E CVg > 0. Embrulha a formula existente (a de cada desempenho OTI/DES/MIN)."""
    f = str(a.Range(f'R{L0}').Formula)
    if f.startswith(f'=IF(AND(N($O{L0})>0,N($P{L0})>0),'):
        return 'ja instalada'
    if 'SQRT' not in f.upper() or f'O{L0}' not in f:
        raise SystemExit(f'Analitos!R{L0}: formula inesperada, nada alterado: {f[:160]}')
    a.Range(f'R{L0}:R{L1}').Formula = f'=IF(AND(N($O{L0})>0,N($P{L0})>0),{f[1:]},"")'
    return 'instalada'


def esperado(a, r, fonte):
    k, m, n_, rr, v, w = (a.Range(f'{c}{r}').Value for c in ('K', 'M', 'N', 'R', 'V', 'W'))
    num = lambda x: isinstance(x, float) and x > 0
    src_t = {'CLIA': k, 'VB': rr, 'FAB': m}[fonte]
    src_u = {'CLIA': v, 'VB': w, 'FAB': n_}[fonte]
    return (src_t if num(src_t) else '-'), (src_u if num(src_u) else '-')


def fumaca(a):
    ruins, cont = [], {f: {'numero': 0, 'traco': 0} for f in FONTES}
    for r in range(L0, L1 + 1):
        if a.Range(f'A{r}').Value in (None, ''):
            if a.Range(f'T{r}').Value not in (None, '') or a.Range(f'U{r}').Value not in (None, ''):
                ruins.append((r, 'linha vazia com T/U', a.Range(f'T{r}').Value))
            continue
        orig = a.Range(f'S{r}').Value
        try:
            for fonte in FONTES:
                a.Range(f'S{r}').Value = fonte
                a.Calculate()
                et, eu = esperado(a, r, fonte)
                t, u = a.Range(f'T{r}').Value, a.Range(f'U{r}').Value
                okt = (t == et) if et == '-' else (isinstance(t, float) and abs(t - et) < 1e-9)
                oku = (u == eu) if eu == '-' else (isinstance(u, float) and abs(u - eu) < 1e-9)
                if not (okt and oku):
                    ruins.append((r, a.Range(f'A{r}').Value, fonte, ('T', t, et), ('U', u, eu)))
                cont[fonte]['traco' if et == '-' else 'numero'] += 1
        finally:
            a.Range(f'S{r}').Value = orig
    a.Calculate()
    return ruins, cont


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

        log('1. VBA (todo o codigo atual das fontes: codigo_atual.py)')
        import codigo_atual
        codigo_atual.instalar(ex, wb, produto, log)

        if produto.startswith('Bio'):
            a = wb.Worksheets('Analitos')
            for nome, col in (('etpOficial', 'T'), ('cvtpOficial', 'U'), ('espFonte', 'S')):
                ref = str(wb.Names(nome).RefersTo).replace('$', '')
                if ref != f'=Analitos!{col}{L0}:{col}{L1}':
                    raise SystemExit(f'nome {nome} aponta para {ref}, esperado Analitos!{col}{L0}:{col}{L1}')
            log(f'2. Analitos R (ETp VB exige CVi e CVg): {coluna_r(a)}')
            t, u = formulas(L0)
            mudou = str(a.Range(f'T{L0}').Formula) != t or str(a.Range(f'U{L0}').Formula) != u
            if mudou:
                a.Range(f'T{L0}:T{L1}').Formula = t
                a.Range(f'U{L0}:U{L1}').Formula = u
            log(f'   T (ETp em uso) e U (CVTp): {"instaladas" if mudou else "ja instaladas"}')
            log('3. fumaca: cada analito x CLIA / VB / FAB')
            ex.xl.Calculation = -4135
            ruins, cont = fumaca(a)
            ex.xl.Calculation = -4105
            if ruins:
                raise SystemExit(f'fumaca ADR-069 falhou: {ruins[:6]}')
            log(f'   OK: {cont} (numero = especificacao cadastrada; traco = ausente)')
            sel = {}
            for r in range(L0, L1 + 1):
                if a.Range(f'A{r}').Value:
                    sel.setdefault(str(a.Range(f'S{r}').Value), []).append(
                        f'{a.Range(f"A{r}").Value}={a.Range(f"T{r}").Text}')
            for k, v in sel.items():
                log(f'   em uso {k}: {", ".join(v)}')
        else:
            log('2. Hematologia: fonte do ETp por prioridade (P/Q/R), sem escolha -- nada a alterar')

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
