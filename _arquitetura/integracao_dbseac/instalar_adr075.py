# -*- coding: utf-8 -*-
"""instalar_adr075.py -- ADR-075: especificacoes da HEMATOLOGIA no modelo da Bioquimica (biblioteca anual por ano,
escolha da fonte CLIA/VB/FAB por analito). Pre-requisito dos blocos por lote (ADR-074) na Hematologia.

Uso:  python instalar_adr075.py <Bioquimica|Hematologia> <arquivo.xlsm> [--sem-salvar]

Pedido do usuario (14/10/2026): "a hematologia nao tem nem as especificacoes de qualidade que a bioquimica tem ...
adeque a hematologia para ficar similar a bioquimica".

Antes (Hematologia): K TEa CLIA, L CV Fab, M CVi, N CVg digitados direto na tabela; O desempenho; P ETp cfg; Q fonte
por prioridade (cfg -> CLIA -> VB); R ETp final; S/T limites de CV; U:W tabela fi/fb.
Depois (= Bioquimica): K ETp CLIA, L CVTp CLIA, M ETp FAB, N CVTp FAB, O CVi, P CVg (todos da BIBLIOTECA ANUAL,
linhas 46-89, pelo ano), Q desempenho, R ETp VB, S ESPQ FONTE (escolha CLIA/VB/FAB), T ETp em uso, U CVTp em uso,
V/W limites de CV, X:Z tabela fi/fb.

Como: tres colunas INSERIDAS (L, M e U) -- o Excel leva junto as ~2.500 formulas da Estatistica que liam a aba
(fonte Q->S, ETp R->T, limites S/T->V/W, CVi M->O); depois as formulas e a biblioteca. Idempotente (marca:
S3 = "ESPQ FONTE"). Bioquimica: so o VBA.
  1. instala todo o VBA atual das fontes; EXIGE compilacao;
  2. confere o layout antigo (cabecalhos) e guarda, por analito: TEa CLIA, CV fab, CVi, CVg, fonte e ETp final;
  3. insere as colunas; cabecalhos da Bioquimica; S = a fonte que valia (CLIA/VB; sem nenhuma, CLIA -> ETp "-");
  4. biblioteca anual 2025-2029 (ANO, nomes, CLIA ETp, CLIA 1/3, VB-CVI, VB-CVG, FAB ETp, FAB CVTp); 2025 recebe os
     valores que estavam digitados (CV fab -> FAB CVTp); 2026-2029 ficam para o laboratorio preencher;
  5. K..P leem a biblioteca pelo ano (S2, ano padrao; o ADR-074 troca pelo ano do lote em analise), R/T/U as
     formulas da Bioquimica; nomes etpOficial, cvtpOficial, espFonte, espVigente, espHistorico, espAno;
     Estatistica J (CVTp) passa a ler cvtpOficial, como na Bioquimica;
  6. FUMACA: o ETp em uso de cada analito e o MESMO ETp final de antes (e o da Estatistica), e a fonte tambem.
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
ANOS = (2025, 2026, 2027, 2028, 2029)
LIB0 = 46                                   # biblioteca: blocos de 9 linhas a partir da 46 (B46:AO89)
ANTIGO = {'K': 'TEa CLIA %', 'L': 'CV Fab %', 'M': 'CVi %', 'N': 'CVg %', 'O': 'Desemp.', 'P': 'ETp cfg %',
          'Q': 'ETp fonte', 'R': 'ETp final %', 'S': 'Lim CV CLIA %', 'T': 'Lim CV VB %', 'U': 'Desemp.', 'V': 'fi',
          'W': 'fb'}
NOVO = {'K': 'ETp CLIA %', 'L': 'CVTp CLIA %', 'M': 'ETp FAB %', 'N': 'CVTp FAB %', 'O': 'CVi %', 'P': 'CVg %',
        'Q': 'Desemp.', 'R': 'ETp VB', 'S': 'ESPQ FONTE', 'T': 'ETp em uso  final %', 'U': 'CVTp%',
        'V': 'Lim CV CLIA %', 'W': 'Lim CV VB %'}
CH = 'CHOOSE({1;2;3;4;5},$B$46,$B$55,$B$64,$B$73,$B$82)'


def log(*a):
    print(*a, flush=True)


def num(v):
    return isinstance(v, (int, float)) and not isinstance(v, bool) and not (isinstance(v, int) and v < -2146820000)


def da_biblioteca(r, desl, ano='$S$2'):
    """Valor do analito da linha r na linha `desl` do bloco do ano (o mesmo modelo da Bioquimica)."""
    blk = f'(MATCH({ano},{CH},0)-1)*9'
    return (f'=IFERROR(INDEX($B$46:$AO$89,{blk}+{desl},MATCH($A{r},INDEX($B$46:$AO$89,{blk}+2,0),0)),"")')


def formulas(r):
    return {
        'K': da_biblioteca(r, 3),
        'L': f'=IF(K{r}="","",K{r}/3)',
        'M': da_biblioteca(r, 7),
        'N': da_biblioteca(r, 8),
        'O': da_biblioteca(r, 5),
        'P': da_biblioteca(r, 6),
        'R': (f'=IF(AND(N($O{r})>0,N($P{r})>0),IFERROR(IF(Q{r}="ÓTI",1.65*(0.25*O{r})+0.125*SQRT(O{r}^2+P{r}^2),'
              f'IF(Q{r}="DES",1.65*(0.5*O{r})+0.25*SQRT(O{r}^2+P{r}^2),IF(Q{r}="MIN",1.65*(0.75*O{r})+0.375*'
              f'SQRT(O{r}^2+P{r}^2),""))),""),"")'),
        'T': (f'=IF($A{r}="","",IF($S{r}="CLIA",IF(N($K{r})>0,$K{r},"-"),IF($S{r}="VB",IF(N($R{r})>0,$R{r},"-"),'
              f'IF($S{r}="FAB",IF(N($M{r})>0,$M{r},"-"),"-"))))'),
        'U': (f'=IF($A{r}="","",IF($S{r}="CLIA",IF(N($V{r})>0,$V{r},"-"),IF($S{r}="VB",IF(N($W{r})>0,$W{r},"-"),'
              f'IF($S{r}="FAB",IF(N($N{r})>0,$N{r},"-"),"-"))))'),
    }


def nome(wb, nm, ref):
    alvo = ref.replace("'", '').replace('$', '').upper()
    for n in list(wb.Names):
        if n.Name == nm:
            if str(n.RefersTo).replace("'", '').replace('$', '').upper() == alvo:
                return
            n.Delete()
            break
    wb.Names.Add(nm, ref)


def ja_migrada(e):
    return str(e.Range('S3').Value or '').strip().upper() == 'ESPQ FONTE'


def conferir_antigo(e):
    ruins = [(c, e.Range(f'{c}3').Value) for c, rot in ANTIGO.items()
             if ' '.join(str(e.Range(f'{c}3').Value or '').split()) != rot]
    if ruins:
        raise SystemExit(f'Hematologia: layout inesperado na linha 3 da aba do motor, nada alterado: {ruins[:5]}')
    if any(e.Range(f'P{r}').Value not in (None, '') for r in range(L0, L1 + 1)):
        raise SystemExit('Hematologia: ha ETp cfg (P) digitado; o modelo da Bioquimica nao tem essa opcao. Nada alterado.')
    for r in range(LIB0, LIB0 + 45):
        if any(v not in (None, '') for v in e.Range(f'A{r}:AO{r}').Value[0]):
            if str(e.Range(f'A{LIB0}').Value or '').strip().upper() != 'ANO':
                raise SystemExit(f'Hematologia: linha {r} da aba do motor ocupada (a biblioteca vai em 46-89). Nada alterado.')


def antes(e):
    out = {}
    for r in range(L0, L1 + 1):
        an = str(e.Range(f'A{r}').Value or '').strip()
        if an:
            k, l_, m, n_, q, rr = (e.Range(f'{c}{r}').Value for c in ('K', 'L', 'M', 'N', 'Q', 'R'))
            out[r] = dict(an=an, clia=k, cvfab=l_, cvi=m, cvg=n_, fonte=str(q or ''), etp=rr)
    return out


def inserir(e):
    a2 = e.Range('A2').MergeArea
    for c in ('L', 'M', 'U'):
        e.Columns(c).Insert()
    # a nota da linha 2 nao pode cobrir Q2:S2 (ano padrao, como na Bioquimica)
    a2 = e.Range('A2').MergeArea
    if a2.Column + a2.Columns.Count - 1 >= 17:
        a2.UnMerge()
        e.Range('A2:P2').Merge()
    for c, rot in NOVO.items():
        e.Range(f'{c}3').Value = rot


def biblioteca(e, guardado):
    ordem = [str(e.Range(f'A{r}').Value or '').strip() for r in range(L0, L1 + 1)]
    for k, ano in enumerate(ANOS):
        b = LIB0 + 9 * k
        e.Range(f'A{b}').Value = 'ANO'
        e.Range(f'B{b}').Value = ano
        e.Range(f'A{b + 1}').Value = 'ESPECIFICAÇÕES DE QUALIDADE'
        e.Range(f'B{b + 1}').Formula2 = '=TRANSPOSE($A$4:$A$43)'
        for i, rot in enumerate(('CLIA ETp%', 'CLIA 1/3 ETp%', 'VB-CVI', 'VB-CVG', 'FAB ETp%', 'FAB CVTp%')):
            e.Range(f'A{b + 2 + i}').Value = rot
        e.Range(f'A{b}:A{b + 7}').Font.Bold = True
        e.Range(f'B{b}').Font.Bold = True
        e.Range(f'B{b}').Interior.Color = 0xCCFFFF
        e.Range(f'B{b + 1}:AO{b + 1}').WrapText = True
        e.Range(f'B{b + 1}:AO{b + 1}').Font.Size = 8
        e.Range(f'B{b + 2}:AO{b + 7}').Locked = False
    # 2025 = o que estava digitado na tabela
    b = LIB0
    seed = 0
    for r, d in guardado.items():
        j = ordem.index(d['an']) + 2                      # coluna B = 1o analito
        for desl, v in ((2, d['clia']), (4, d['cvi']), (5, d['cvg']), (7, d['cvfab'])):
            if num(v):
                e.Cells(b + desl, j).Value = v
                seed += 1
    return seed


def tabela(wb, e, guardado):
    for r in range(L0, L1 + 1):
        f = formulas(r)
        for c in ('K', 'L', 'M', 'N', 'O', 'P', 'R', 'T', 'U'):
            e.Range(f'{c}{r}').Formula2 = f[c]
        d = guardado.get(r)
        if d:
            e.Range(f'S{r}').Value = d['fonte'] if d['fonte'] in ('CLIA', 'VB', 'FAB') else 'CLIA'
    s = e.Range(f'S{L0}:S{L1}')
    s.Locked = False
    s.Validation.Delete()
    sep = wb.Application.International[5 - 1]
    s.Validation.Add(3, 1, 1, sep.join(('CLIA', 'VB', 'FAB')))
    s.Validation.InputTitle = 'Especificação do ETp'
    s.Validation.InputMessage = 'CLIA (coluna K), VB (R: CVi e CVg) ou FAB (M).'
    s.HorizontalAlignment = -4108
    for c in ('K', 'L', 'M', 'N', 'O', 'P', 'R', 'T', 'U'):
        e.Range(f'{c}{L0}:{c}{L1}').Locked = True
    if str(e.Range('Q2').Value or '').strip() == '':
        e.Range('Q2').Value = 'ESPECIFICAÇÕES DO ANO'
        e.Range('S2').Value = ANOS[0]
        e.Range('Q2').Font.Bold = True
        e.Range('S2').Font.Bold = True
        e.Range('S2').Interior.Color = 0xCCFFFF
        e.Range('S2').Locked = False
    esp = e.Name
    nome(wb, 'etpOficial', f"='{esp}'!$T${L0}:$T${L1}")
    nome(wb, 'cvtpOficial', f"='{esp}'!$U${L0}:$U${L1}")
    nome(wb, 'espFonte', f"='{esp}'!$S${L0}:$S${L1}")
    nome(wb, 'espVigente', f"='{esp}'!$A${L0}:$W${L1}")
    nome(wb, 'espHistorico', f"='{esp}'!$A$46:$AO$89")
    nome(wb, 'espAno', f"='{esp}'!$S$2")


def estatistica_cvtp(wb, e):
    """Estatistica J (CVTp) le cvtpOficial, como na Bioquimica (antes: limite VB ou CLIA, sem a fonte FAB)."""
    from instalar_adr064 import LIN_TAB, ultima_linha_tabela
    est = wb.Worksheets('Estatística')
    ult = ultima_linha_tabela(est)
    f = str(est.Range(f'J{LIN_TAB}').Formula)
    if 'cvtpOficial' in f:
        return 'ja le cvtpOficial'
    est.Range(f'J{LIN_TAB}:J{ult}').Formula = (f"=IF($A{LIN_TAB}=\"\",\"\",IFERROR(INDEX(cvtpOficial,MATCH($A{LIN_TAB},"
                                              f"'{e.Name}'!$A$4:$A$43,0)),\"\"))")
    return f'J{LIN_TAB}:J{ult} -> cvtpOficial'


def fumaca(wb, e, guardado):
    from instalar_adr064 import LIN_TAB
    ruins = []
    est = wb.Worksheets('Estatística')
    k_est = {}
    for linha in est.Range(f'A{LIN_TAB}:K{LIN_TAB + 119}').Value:
        if linha[0] and linha[1] == 1 and str(linha[0]) not in k_est:
            k_est[str(linha[0]).strip()] = linha[10]
    for r, d in guardado.items():
        t, s = e.Range(f'T{r}').Value, e.Range(f'S{r}').Value
        velho = d['etp']
        if num(velho):
            if not (num(t) and abs(t - velho) < 1e-9):
                ruins.append(('ETp mudou', d['an'], velho, t))
        elif t != '-':
            ruins.append(('sem ETp antes, com ETp agora', d['an'], velho, t))
        if d['fonte'] in ('CLIA', 'VB') and s != d['fonte']:
            ruins.append(('fonte mudou', d['an'], d['fonte'], s))
        ke = k_est.get(d['an'])
        if ke is not None and not ((num(ke) and num(t) and abs(ke - t) < 1e-9) or (ke == '-' and t == '-')):
            ruins.append(('Estatistica K != T', d['an'], ke, t))
        for c in ('K', 'L', 'M', 'N', 'O', 'P', 'R', 'T', 'U', 'V', 'W'):
            if str(e.Range(f'{c}{r}').Text).startswith('#'):
                ruins.append(('erro', d['an'], c))
    if ruins:
        raise SystemExit(f'fumaca ADR-075 falhou: {ruins[:6]}')
    return len(guardado)


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
            log('2. Bioquimica: ja tem a biblioteca anual e a escolha da fonte; so o VBA')
        else:
            e = wb.Worksheets(ii.aba_espec(wb))
            if ja_migrada(e):
                log(f'2. Hematologia: {e.Name} ja no modelo da Bioquimica')
                ex.xl.Calculation = -4105
                ex.xl.Calculate()
            else:
                ex.xl.Calculation = -4135
                conferir_antigo(e)
                guardado = antes(e)
                log(f'2. Hematologia: layout antigo conferido; {len(guardado)} analitos guardados (TEa CLIA, CV fab, CVi, CVg, fonte, ETp)')
                inserir(e)
                log('3. colunas L, M e U inseridas (formulas da Estatistica acompanham); cabecalhos da Bioquimica')
                n = biblioteca(e, guardado)
                log(f'4. biblioteca anual {ANOS[0]}-{ANOS[-1]} em {e.Name}!A46:AO89; {n} valor(es) de {ANOS[0]} vindos da tabela')
                tabela(wb, e, guardado)
                log(f'5. K..P pela biblioteca; R/T/U da Bioquimica; S = fonte que valia; nomes etpOficial/cvtpOficial/espFonte/espAno; '
                    f'{estatistica_cvtp(wb, e)}')
                ex.xl.Calculation = -4105
                ex.xl.Calculate()
                ex.esperar()
                nf = fumaca(wb, e, guardado)
                log(f'6. fumaca OK: ETp em uso e fonte de {nf} analitos = os de antes; Estatistica K = T')
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
