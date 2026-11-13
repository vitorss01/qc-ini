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
from instalar_seguranca_usuarios import foto_protecao, devolver_protecao  # noqa: E402

MUDANCAS = []
CONSULTAS = ['DB_CQ_FINAL', 'QA_INTEGRACAO']
SEM_MEDIA = (0x007A9C, 0xCEF4FF)            # fonte e fundo da regra "SEM MÉDIA/DP" do painel.py
D4 = '=IF(Painel!K3="","",Painel!K3)'
D4_ANTIGO = '=Painel!I3'
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
    if norm(atual) != norm(antigo):                     # 'a' (constante do I5) tambem passa por aqui
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
    setf(e.Range('D4'), D4, D4_ANTIGO, 'Estatística D4 (ano do Painel K3)')
    if str(e.Range('A4').Value or '') != A4:
        if str(e.Range('A4').Value or '').strip() != 'Visão':
            raise SystemExit(f'Estatística A4 inesperada: {e.Range("A4").Value!r} -- nada foi salvo')
        e.Range('A4').Value = A4
        e.Range('A4').ShrinkToFit = True
        MUDANCAS.append('Estatística A4 (Visão segue o Painel)')
    setf(e.Range('H5'), H5, H5_ANTIGO, 'Estatística H5 (histórico inteiro sem De/Até)')
    setf(e.Range('J5'), J5, J5_ANTIGO, 'Estatística J5')
    setf(e.Range('I5'), I5, 'a', 'Estatística I5 (sem o "a" quando é o histórico inteiro)')


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


def fumaca(ex, wb, produto):
    conferir_consultas(wb)
    ex.xl.Calculate()
    p = wb.Worksheets('Painel')
    e = wb.Worksheets('Estatística')
    if produto.startswith('Bio'):
        d4, k3 = e.Range('D4').Text, p.Range('K3').Text
        if str(d4) != str(k3) or str(d4) == '0':
            raise SystemExit(f'fumaca: Estatística D4 = {d4!r}, Painel K3 = {k3!r} -- nada foi salvo')
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
