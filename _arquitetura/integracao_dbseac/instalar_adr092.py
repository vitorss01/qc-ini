# -*- coding: utf-8 -*-
"""instalar_adr092.py -- ADR-092 (12/11/2026): ocorrencias do reteste 03 (Tester, 11/11/2026).
Nenhum item muda criterio estatistico: o N-01 tira da estatistica um manual que entrava no analito errado (regra
conservadora: recusa, nao mapeia; decisao de regra pendente de confirmacao do CQ), o resto e trilha, sessao e exibicao.

Uso:  python instalar_adr092.py <Bioquimica|Hematologia> <arquivo.xlsm> [--sem-salvar]

Idempotente (a 2a execucao registra 0 mudancas). Passos:
  1. consultas DB_CQ_FINAL e QA_INTEGRACAO regravadas de pq/ e conferidas (N-01 no motor reserva PQ: manual num lote de
     regra MATRIZ com analito do catalogo sem de/para nessa matriz -> LOTE_OUTRO_MATERIAL; A13 com o texto proprio);
  2. todo o VBA atual das fontes (codigo_atual.py), compilacao exigida: mEtlCQ/mEtlQA (N-01), mCEQForm (N-02: Corrigir e
     Excluir exigem sessao), mIntegracao (N-03: trilha RESULTADO_MANUAL_FORA_DA_ESTATISTICA para IMPLAUSIVEL/LOTE_*);
  3. planilhas:
     - Painel J7:J9: regra de cor "comeca com SEM MÉD" com as cores da "SEM MÉDIA/DP" (a forma curta do OC-0043; as
       formulas curtas vem do instalar_adr078, que roda antes no entregar);
     - Bioquimica, Estatistica (OC-0040, UX G22): D4 = ano do Painel (K3), vazio em vez de "0"; A4 "Visão (segue o
       Painel)"; H5 = "histórico inteiro do lote" e I5/J5 vazios quando o Painel nao tem De/Ate;
     - Hematologia, Estatistica (OC-0049, UX G31): coluna B (nivel) vazia nas linhas sem analito;
  4. fumaca; 5. devolve a protecao como estava e salva com os eventos desligados. Qualquer falha para ANTES de salvar.
"""
import os
import re
import sys

AQUI = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, AQUI)
sys.path.insert(0, os.path.join(AQUI, '..', 'etapa1_multilote'))
import xlh  # noqa: E402
import pqlib  # noqa: E402
import instalar_integracao as ii  # noqa: E402
import codigo_atual  # noqa: E402
import camada_dados as cd  # noqa: E402
from instalar_seguranca_usuarios import foto_protecao, devolver_protecao  # noqa: E402

MUDANCAS = []
CONSULTAS = ['DB_CQ_FINAL', 'QA_INTEGRACAO']
SEM_MEDIA = (0x007A9C, 0xCEF4FF)            # fonte e fundo da regra "SEM MÉDIA/DP" do painel.py
# ADR-095 (reteste 04, P-09): sem De/Ate no Painel ("Tudo"), o D4 fica vazio -- dizia "Ano 2026" ao lado do "historico
# inteiro do lote". So exibicao: com o B4 = PERSONALIZADO o motor (mEstatistica, linha ~2132) le B5/D5, nunca o D4
D4 = '=IF(OR(Painel!K3="",AND(Painel!$G$3="",Painel!$G$4="")),"",Painel!K3)'
D4_ANTIGO = ('=Painel!I3', '=IF(Painel!K3="","",Painel!K3)')     # o de antes do ADR-092 e o do ADR-092
# ADR-095: B5/D5 (=Painel!G3/G4) com o Painel em "Tudo" mostravam "00/01/1900"; so o formato (o zero fica em branco)
FMT_B5D5 = 'dd/mm/yyyy;;'
# ADR-096 (R5-01): H5/J5 com valor 0 (so o Ate / so o De) dizem o lado aberto da janela; sem 'y' no texto (fmt_data)
FMT_H5 = 'dd/mm/yyyy;;"primeiro dado"'
FMT_J5 = 'dd/mm/yyyy;;"último dado"'
A4 = 'Visão (segue o Painel)'
HIST = 'histórico inteiro do lote'
H5_ANTIGO = '=JanelaInicio($B$4,$D$4,$F$4,$B$5,$D$5)'
J5_ANTIGO = '=JanelaFim($B$4,$D$4,$F$4,$B$5,$D$5)'
SEM_DEATE = 'AND(Painel!$G$3="",Painel!$G$4="")'
H5 = f'=IF({SEM_DEATE},"{HIST}",JanelaInicio($B$4,$D$4,$F$4,$B$5,$D$5))'
J5 = f'=IF({SEM_DEATE},"",JanelaFim($B$4,$D$4,$F$4,$B$5,$D$5))'
I5 = f'=IF({SEM_DEATE},"","a")'


def log(m):
    print(m, flush=True)


def norm(f):
    return str(f or '').replace('_xlfn.', '').replace(' ', '').upper()


def normalizar(t):
    return '\n'.join(l.rstrip() for l in str(t).replace('\r\n', '\n').replace('\r', '\n').strip().split('\n'))


def consultas(wb):
    wb.Queries.FastCombine = True
    for q in CONSULTAS:
        atual = pqlib.query_existe(wb, q)
        if atual is None:
            raise SystemExit(f'consulta {q} nao existe no arquivo -- nada foi salvo')
        if normalizar(atual.Formula) != normalizar(pqlib.m_de(q)):
            pqlib.gravar_query(wb, q, pqlib.m_de(q))
            MUDANCAS.append(f'consulta {q}')
    conferir_consultas(wb)


def conferir_consultas(wb):
    for q in CONSULTAS:
        if normalizar(pqlib.query_existe(wb, q).Formula) != normalizar(pqlib.m_de(q)):
            raise SystemExit(f'consulta {q}: formula no arquivo difere da fonte depois de gravar -- nada foi salvo')


def setf(rng, f, antigo, oque):
    atual = str(rng.Formula)
    if norm(atual) == norm(f):
        return
    antigos = antigo if isinstance(antigo, tuple) else (antigo,)
    if norm(atual) not in [norm(x) for x in antigos]:   # 'a' (constante do I5) tambem passa por aqui
        raise SystemExit(f'{oque}: formula inesperada {atual[:160]!r} -- nada foi salvo')
    rng.Formula = f
    MUDANCAS.append(oque)


def painel(wb):
    """Regra 'comeca com SEM MÉD' em J7:J9 (pega 'SEM MÉD/DP'), com as cores da regra 'SEM MÉDIA/DP'."""
    p = wb.Worksheets('Painel')
    nlv = 3 if p.Range('J9').HasFormula else 2
    rng = p.Range(f'J7:J{6 + nlv}')
    if 'SEM MÉD/DP' not in str(p.Range('J7').Formula):
        raise SystemExit('Painel J7 sem a forma curta "SEM MÉD/DP" (instalar_adr078 nao rodou antes) -- nada foi salvo')
    fcs = rng.FormatConditions
    for i in range(1, fcs.Count + 1):
        if fcs(i).Type == 9 and str(fcs(i).Text) == 'SEM MÉD':
            return
    fc = fcs.Add(9, None, None, None, 'SEM MÉD', 2)       # xlTextString, xlBeginsWith (posicional: ver instalar_adr078)
    fc.Font.Color, fc.Font.Bold, fc.Interior.Color = SEM_MEDIA[0], True, SEM_MEDIA[1]
    fc.StopIfTrue = False
    MUDANCAS.append(f'Painel {rng.Address} cor "SEM MÉD"')


def estatistica_bio(wb):
    e = wb.Worksheets('Estatística')
    setf(e.Range('D4'), D4, D4_ANTIGO, 'Estatística D4 (ano do Painel K3; vazio sem De/Até)')
    for c in ('B5', 'D5'):                                  # ADR-095: so o formato (formula e valor intactos)
        # NumberFormatLocal com o codigo de ano do Excel (cd.fmt_data): 'yyyy' por NumberFormat vira o texto 'yyyy' no
        # pt-BR (achado da revisao independente do ADR-095, 14/11)
        local = FMT_B5D5.replace('y', str(e.Application.International[19 - 1] or 'y'))
        if str(e.Range(c).NumberFormatLocal) != local:
            if not re.fullmatch(r'=Painel!\$?G\$?[34]', str(e.Range(c).Formula)):
                raise SystemExit(f'Estatística {c} inesperada: {e.Range(c).Formula!r} -- nada foi salvo')
            cd.fmt_data(e.Range(c), FMT_B5D5)
            MUDANCAS.append(f'Estatística {c} formato (zero em branco)')
    if str(e.Range('A4').Value or '') != A4:
        if str(e.Range('A4').Value or '').strip() != 'Visão':
            raise SystemExit(f'Estatística A4 inesperada: {e.Range("A4").Value!r} -- nada foi salvo')
        e.Range('A4').Value = A4
        e.Range('A4').ShrinkToFit = True
        MUDANCAS.append('Estatística A4 (Visão segue o Painel)')
    setf(e.Range('H5'), H5, H5_ANTIGO, 'Estatística H5 (histórico inteiro sem De/Até)')
    setf(e.Range('J5'), J5, J5_ANTIGO, 'Estatística J5')
    setf(e.Range('I5'), I5, 'a', 'Estatística I5 (sem o "a" quando é o histórico inteiro)')
    # ADR-096 (reteste 05, R5-01): com so o De (ou so o Ate) no Painel, JanelaFim (JanelaInicio) devolve 0 = sem limite
    # e a celula mostrava "00/01/1900". So o formato: o zero diz o que a janela e (o Tester conferiu em 15/11 que o
    # motor vai do De ao ultimo dado); formula e Value2 intactos (o setf acima garante a formula do ADR-092)
    for c, fmt in (('H5', FMT_H5), ('J5', FMT_J5)):
        local = fmt.replace('y', str(e.Application.International[19 - 1] or 'y'))
        if str(e.Range(c).NumberFormatLocal) != local:
            cd.fmt_data(e.Range(c), fmt)
            MUDANCAS.append(f'Estatística {c} formato (zero = {fmt.split(";")[-1]})')


MSG_D3 = 'Ano Até deve ser igual ou maior que Ano De.'
VERMELHO = 192                                          # RGB(192, 0, 0) no Font.Color do Excel


def estatistica_hema(wb):
    e = wb.Worksheets('Estatística')
    n = 0
    for r in range(14, 134):
        b = e.Cells(r, 2)
        f = str(b.Formula)
        alvo = None
        if re.fullmatch(r'[123]', f):
            alvo = f'=IF($A{r}="","",{f})'
        elif not re.fullmatch(rf'=IF\(\$A{r}="","",[123]\)', f):
            raise SystemExit(f'Estatística!B{r} inesperada: {f!r} -- nada foi salvo')
        if alvo:
            b.Formula = alvo
            n += 1
    if n:
        MUDANCAS.append(f'Estatística B (Hematologia): nível vazio sem analito ({n} linha(s))')
    # ADR-097 (OC-0048, G30 itens 1 e 2): Ano Até menor que o Ano De era aceito e o VBA o trocava calado pelo Ano De.
    # Validação por referência (sem nome de função: o Formula1 da validação é lido na língua do Excel)
    v = e.Range('D3').Validation
    try:
        tipo, f1, msg = int(v.Type), str(v.Formula1), str(v.ErrorMessage)
    except Exception:                                   # noqa: BLE001 (sem validação)
        tipo, f1, msg = -1, '', ''
    if (tipo, f1.replace('$', ''), msg) != (1, '=B3', MSG_D3):
        v.Delete()
        v.Add(1, 1, 7, '=$B$3')                         # xlValidateWholeNumber, xlValidAlertStop, xlGreaterEqual
        v.IgnoreBlank = True
        v.ErrorTitle = 'Ano Até'
        v.ErrorMessage = MSG_D3
        v.ShowError = True
        MUDANCAS.append('Estatística D3: Ano Até >= Ano De')
    # A5: aviso do lote fora do Ano De/Até (o VBA escreve o texto em AtualizarEstatisticaAba)
    a5 = e.Range('A5')
    if int(a5.Font.Color) != VERMELHO or not a5.Font.Bold:
        a5.Font.Color = VERMELHO
        a5.Font.Bold = True
        MUDANCAS.append('Estatística A5: aviso do ano em vermelho')


def fumaca(ex, wb, produto):
    conferir_consultas(wb)
    ex.xl.Calculate()
    p = wb.Worksheets('Painel')
    e = wb.Worksheets('Estatística')
    if produto.startswith('Bio'):
        d4, k3 = e.Range('D4').Text, p.Range('K3').Text
        sem = p.Range('G3').Value in (None, '') and p.Range('G4').Value in (None, '')
        b5d5 = str(e.Range('B5').Text) + str(e.Range('D5').Text)
        if (str(d4) != ('' if sem else str(k3)) or str(d4) == '0' or '1900' in b5d5
                or re.search(r'[A-Za-z]', b5d5)
                or any(t and not re.fullmatch(r'\d{2}/\d{2}/\d{4}', t)
                       for t in (str(e.Range('B5').Text), str(e.Range('D5').Text)))):
            raise SystemExit(f'fumaca: Estatística D4 = {d4!r}, Painel K3 = {k3!r}, B5/D5 = {b5d5!r} -- nada foi salvo')
        log(f'   Estatística D4 = {d4!r} (Painel K3), H5 = {e.Range("H5").Text!r}')
    else:
        sujas = [r for r in range(14, 134) if not str(e.Cells(r, 1).Text) and str(e.Cells(r, 2).Text)]
        cheias = [r for r in range(14, 134) if str(e.Cells(r, 1).Text) and str(e.Cells(r, 2).Text) not in ('1', '2', '3')]
        if sujas or cheias:
            raise SystemExit(f'fumaca: Estatística B sujas {sujas[:5]} / nivel ruim {cheias[:5]} -- nada foi salvo')
    log('   ok')


def main(produto, caminho, salvar=True):
    try:
        sys.stdout.reconfigure(errors='replace')
    except Exception:                                   # noqa: BLE001
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
        ex.xl.EnableEvents = False
        log('1. consultas DB_CQ_FINAL e QA_INTEGRACAO (pq/)')
        consultas(wb)
        log('2. VBA (todo o codigo atual das fontes: codigo_atual.py)')
        codigo_atual.instalar(ex, wb, produto, log)
        ex.xl.Calculation = -4135
        log('3. planilhas')
        painel(wb)
        if produto.startswith('Bio'):
            estatistica_bio(wb)
        else:
            estatistica_hema(wb)
        log('4. fumaca')
        ex.xl.Calculation = -4105
        fumaca(ex, wb, produto)
        ex.xl.EnableEvents = False
        devolver_protecao(wb, foto)
        if estrutura:
            wb.Protect(ii.SENHA, True, False)
        log(f'MUDANCAS {len(MUDANCAS)}' + (': ' + '; '.join(MUDANCAS) if MUDANCAS else ''))
        if salvar:
            antes = os.path.getmtime(wb.FullName)
            wb.Save()
            if os.path.getmtime(wb.FullName) == antes:
                raise SystemExit('o Save nao chegou ao arquivo (data de gravacao igual) -- conferir')
            log('salvo')
    finally:
        ex.fechar()


if __name__ == '__main__':
    a = sys.argv[1:]
    if len(a) < 2:
        raise SystemExit(__doc__)
    main(a[0], a[1], salvar='--sem-salvar' not in a)
