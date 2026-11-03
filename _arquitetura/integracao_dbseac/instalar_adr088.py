# -*- coding: utf-8 -*-
"""instalar_adr088.py -- ADR-088 (03/11/2026): lote 1 de correcoes da auditoria (ciclo 2026-10-31_ciclo01).

Uso:  python instalar_adr088.py <Bioquimica|Hematologia> <arquivo.xlsm> [--atualizar] [--sem-salvar]

Idempotente (a 2a execucao registra 0 mudancas). Passos:
  1. consultas DB_CQ_FINAL e QA_INTEGRACAO regravadas de pq/ e conferidas (IMPLAUSIVEL, A11, A12 -- motor reserva PQ);
  2. todo o VBA atual das fontes (codigo_atual.py): mIntegracao (EQUIPAMENTO/MATRIZ/ANALITO do manual na grafia do
     cadastro, OC-0001), mEtlCQ (data/hora inexistente = MANUAL_INCOMPLETO, OC-0005; CIQ <= 0 = IMPLAUSIVEL, OC-0004),
     mEtlQA (A11/A12), mCEQ (BiasEQ com o filtro de amostras do ViesEQ, OC-0006), mLotes (troca de lote refaz a
     Estatistica, OC-0003) e o modulo da aba Estatistica (Hema: Ano De/Ate refazem C:F); EXIGE compilacao;
  3. Painel H7:H8 (Bio) / H7:H9 (Hema): ET so com bias numerico (OC-0002, restaura o ADR-030);
  4. Painel G6/I6: no modo TODOS (H3 vazio) o rotulo diz "do metodo" (OC-0013, emenda do ADR-086 dec. 4);
  5. Bio, Estatistica B103: "Sem EP" conta so as linhas com analito (OC-0009);
  6. fumaca: nenhuma formula do Painel com o guarda IF(ISNUMBER($G..),$G..,0); o nome engTodos existe e e o do ADR-086;
     consultas conferidas de novo; a Estatistica e o BiasEQ respondem sem erro; com --atualizar, o ATUALIZAR so das
     entradas (sem ler o DB_SEAC) conclui e IMPLAUSIVEL = achados A11;
  7. devolve a protecao de cada aba e a da estrutura como estavam; salva com os eventos desligados e confere a data de
     gravacao. Qualquer falha para ANTES de salvar.
Sem --atualizar, a tblCQ_Final gravada continua a de antes ate o proximo ATUALIZAR (que aplica as regras novas).
"""
import os
import re
import sys
import time

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
RX_GUARDA = re.compile(r'IF\(ISNUMBER\(\$G\d+\),\$G\d+,0\)')         # o guarda que o ADR-030 tinha tirado
G6_TODOS = '=IF(engTodos,"Bias % do método (CEQ)","Bias %")'
I6_TODOS = '"Sigma do método"'
B103 = '=COUNTIFS($G$14:$G$93,"SEM EP",$A$14:$A$93,"<>0",$A$14:$A$93,"<>")'


def log(m):
    print(m, flush=True)


def norm(f):
    return str(f or '').replace('_xlfn.', '').replace(' ', '').upper()


def normalizar(t):
    return '\n'.join(l.rstrip() for l in str(t).replace('\r\n', '\n').replace('\r', '\n').strip().split('\n'))


def h_et(r):
    return f'=IF(OR($E{r}="",NOT(ISNUMBER($G{r}))),"",$E{r}*1.65+ABS($G{r}))'


def setf(rng, f, oque):
    if norm(rng.Formula) != norm(f):
        rng.Formula = f
        MUDANCAS.append(oque)


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


def painel(wb, nlv):
    p = wb.Worksheets('Painel')
    # 3. ET (H) so com bias numerico
    for r in range(7, 7 + nlv):
        old = str(p.Range(f'H{r}').Formula)
        if not old.startswith('=IF(') or '1.65' not in old:
            raise SystemExit(f'Painel!H{r} inesperada: {old[:120]} -- nada foi salvo')
        setf(p.Range(f'H{r}'), h_et(r), f'Painel H{r} (ET so com bias numerico)')
    # 4. rotulos do modo TODOS (ADR-086 emendado): Bias e Sigma sao do metodo
    g6 = str(p.Range('G6').Formula)
    if 'engTodos' not in g6:
        if g6.strip() != 'Bias %':
            raise SystemExit(f'Painel!G6 inesperada: {g6!r} -- nada foi salvo')
        p.Range('G6').Formula = G6_TODOS
        MUDANCAS.append('Painel G6 (rotulo do modo TODOS)')
    elif norm(g6) != norm(G6_TODOS):
        p.Range('G6').Formula = G6_TODOS
        MUDANCAS.append('Painel G6 (rotulo do modo TODOS)')
    i6 = str(p.Range('I6').Formula)
    if 'engTodos' not in i6:
        if not i6.startswith('=IF(N(Sigma_Meses)'):
            raise SystemExit(f'Painel!I6 inesperada: {i6!r} -- nada foi salvo')
        p.Range('I6').Formula = f'=IF(engTodos,{I6_TODOS},' + i6[1:] + ')'
        MUDANCAS.append('Painel I6 (rotulo do modo TODOS)')
    elif norm(I6_TODOS) not in norm(i6):
        raise SystemExit(f'Painel!I6 com engTodos mas sem o rotulo esperado: {i6!r} -- nada foi salvo')


def estatistica_bio(wb):
    e = wb.Worksheets('Estatística')
    a103, old = str(e.Range('A103').Value or ''), str(e.Range('B103').Formula)
    if not a103.startswith('Sem EP'):
        raise SystemExit(f'Estatistica!A103 inesperada: {a103!r} -- nada foi salvo')
    if norm(old) != norm(B103):
        if norm(old) != norm('=COUNTIF($G$14:$G$93,"SEM EP")'):
            raise SystemExit(f'Estatistica!B103 inesperada: {old!r} -- nada foi salvo')
        e.Range('B103').Formula = B103
        MUDANCAS.append('Estatistica B103 (Sem EP so com analito)')


def fumaca(ex, wb, produto, nlv, atualizar):
    p = wb.Worksheets('Painel')
    ruins = []
    ur = p.UsedRange
    fs = ur.Formula
    for lin in (fs if isinstance(fs, tuple) else ((fs,),)):
        for f in lin:
            if isinstance(f, str) and RX_GUARDA.search(f.replace(' ', '')):
                ruins.append(f[:100])
    if ruins:
        raise SystemExit(f'fumaca: {len(ruins)} formula(s) do Painel com o guarda do bias 0: {ruins[:2]} -- nada foi salvo')
    if norm(wb.Names('engTodos').RefersTo) != norm('=N(Eng_Saida!$AD$1)<>0'):
        raise SystemExit(f'fumaca: nome engTodos = {wb.Names("engTodos").RefersTo!r} (esperado o do ADR-086)')
    for r in range(7, 7 + nlv):
        if norm(p.Range(f'H{r}').Formula) != norm(h_et(r)):
            raise SystemExit(f'fumaca: Painel!H{r} nao ficou como pedido')
    conferir_consultas(wb)
    nome = wb.Name
    an = str(wb.Worksheets('Especificações').Range('A4').Value or '').strip()
    ex.xl.Calculation = -4105
    b = ex.run(f"'{nome}'!mCEQ.BiasEQ", an, 'TODOS', 'ABS', 'CAP', 'TODAS', teto=300)
    log(f'   BiasEQ({an!r}, TODOS, ABS, CAP) = {b!r}')
    ex.run(f"'{nome}'!mEstatistica.AtualizarEstatisticaAba", teto=600)
    log('   mEstatistica.AtualizarEstatisticaAba sem erro')
    if atualizar:
        ex.xl.EnableEvents = True
        try:
            t0 = time.time()
            r = str(ex.run(f"'{nome}'!mIntegracao.AtualizarSoEntradas", teto=1500))
        finally:
            ex.xl.EnableEvents = False
        if not r.startswith('OK|'):
            raise SystemExit('ATUALIZAR (so entradas) falhou: ' + r[:600])
        fin = pqlib.ler_tabela(pqlib.tabela(wb, 'tblCQ_Final'))
        qa = pqlib.ler_tabela(pqlib.tabela(wb, 'tblQA_Integracao'))
        imp = {x['ID_REGISTRO'] for x in fin if x.get('STATUS_ANALITICO') == 'IMPLAUSIVEL'}
        a11 = {x['ID_REGISTRO'] for x in qa if x.get('CODIGO') == 'A11'}
        le0 = {x['ID_REGISTRO'] for x in fin if x.get('PARTICIPA_ESTATISTICA') == 'SIM'
               and isinstance(x.get('RESULTADO'), (int, float)) and x['RESULTADO'] <= 0}
        log(f'   ATUALIZAR so entradas OK ({time.time() - t0:.1f}s): IMPLAUSIVEL {len(imp)}, A11 {len(a11)}, '
            f'participantes <= 0: {len(le0)}, A12 {sum(1 for x in qa if x.get("CODIGO") == "A12")}')
        if imp != a11 or le0:
            raise SystemExit('fumaca: IMPLAUSIVEL != A11 ou ainda ha participante <= 0 -- nada foi salvo')


def main(produto, caminho, atualizar=False, salvar=True):
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
        nlv = 2 if produto.startswith('Bio') else 3
        log('1. consultas DB_CQ_FINAL e QA_INTEGRACAO (pq/)')
        consultas(wb)
        log('2. VBA (todo o codigo atual das fontes: codigo_atual.py)')
        codigo_atual.instalar(ex, wb, produto, log)
        ex.xl.Calculation = -4135
        log('3/4. Painel: ET so com bias numerico; rotulos "do método" no modo TODOS')
        painel(wb, nlv)
        if produto.startswith('Bio'):
            log('5. Estatistica B103 (Sem EP so com analito)')
            estatistica_bio(wb)
        log('6. fumaca')
        fumaca(ex, wb, produto, nlv, atualizar)
        ex.xl.EnableEvents = False
        ex.xl.Calculation = -4105
        devolver_protecao(wb, foto)
        if estrutura:
            wb.Protect(ii.SENHA, True, False)
        log(f'MUDANCAS {len(MUDANCAS)}' + (': ' + '; '.join(MUDANCAS) if MUDANCAS else ''))
        if salvar:
            # eventos desligados: com eles o Workbook_BeforeSave cancela o Save e engole o erro
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
    main(a[0], a[1], atualizar='--atualizar' in a, salvar='--sem-salvar' not in a)
