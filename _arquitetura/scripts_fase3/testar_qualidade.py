# -*- coding: utf-8 -*-
"""testar_qualidade.py - ADR-033: as 15 provas de bias, sigma e orcamento

Nenhuma prova aqui olha o dado que por acaso esta na pasta. Todas INJETAM
valores conhecidos (CV, ETp, rodadas de EP com bias definido) e conferem o que
sai contra a conta feita a mao. O trabalho e feito numa COPIA em TEMP: o
arquivo de producao nao e tocado.

As duas provas centrais sao a 2 e a 3. Elas separam o bias com sinal do bias em
magnitude. Com a formula que estava na pasta -- ET = 1,65*CV + Bias e
Sigma = (ETp - Bias)/CV -- um bias de -8% DIMINUIA o erro total e AUMENTAVA o
Sigma. Westgard 2018, pagina 3: o bias sempre encolhe o erro permitido, nunca
o aumenta. Prova: bias -8 e bias +8 tem de produzir ET igual e Sigma igual.

Uso: python testar_qualidade.py <arquivo.xlsm>
"""
import io
import os
import sys
import time
import shutil
import subprocess

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', write_through=True)
import win32com.client as w

falhas = []

# ADR-034: a fonte do EP deixou de ser a EQC_Dados. Este fixture injetava
# rodadas naquela aba; depois do repontamento do mCEQ para a EQA_Base, as
# linhas continuavam sendo gravadas e simplesmente nao eram lidas -- e as
# provas passaram a medir o bias REAL do CAP em vez do bias controlado.
# O produto estava certo; o instrumento e que apontava para o lugar antigo.
EQ_ABA = 'EQA.CAP_Dados'
EQ_R0 = 2
RODADA_TESTE = 'TESTE-2099'     # nao colide com C-A/C-B/C-C nem com o legado
ANO_TESTE = 2099
LIN = 14                    # primeira linha da tabela: analito 1, nivel 1

# Estatistica
C_CV, C_BIAS, C_ET, C_FONTE = 6, 7, 8, 9
C_CVTP, C_ETP, C_SIGMA, C_CLASSE = 10, 11, 12, 13
C_MPP, C_MPCT, C_MSTATUS, C_STCV = 14, 15, 16, 17


def ck(nome, cond, det=''):
    print(('  OK   ' if cond else '  FALHA') + '  ' + nome + (('  -> ' + det) if det else ''))
    if not cond:
        falhas.append(nome)


def perto(a, b, tol=1e-6):
    try:
        return abs(float(a) - float(b)) <= tol
    except Exception:
        return False


def novo_excel():
    for t in range(1, 9):
        try:
            xl = w.DispatchEx('Excel.Application')
            xl.Visible = False
            xl.DisplayAlerts = False
            xl.EnableEvents = False
            xl.AutomationSecurity = 1
            return xl
        except Exception:
            if t in (1, 3, 5):
                subprocess.call(['powershell', '-NoProfile', '-Command',
                                 'Start-Process excel.exe -WindowStyle Minimized'],
                                stderr=subprocess.DEVNULL)
                time.sleep(10)
            time.sleep(2.5 * t)
    raise RuntimeError('Excel COM nao subiu')


# Erros que significam 'o Excel esta ocupado, tente de novo' -- e nao
# 'sua chamada esta errada'. 'Membro nao encontrado' entrou na lista
# depois de aparecer de forma intermitente em Range() logo apos um
# recalculo pesado: o modelo de objetos fica parcialmente indisponivel
# enquanto o Excel se assenta.
TRANSITORIOS = ('rejeitada', 'rejected', 'membro n', 'member not found',
                'call was rejected', 'ocupado', 'busy')


def tenta(fn, vezes=10):
    ult = None
    for i in range(vezes):
        try:
            return fn()
        except Exception as e:
            ult = e
            s = str(e).lower()
            if not any(t in s for t in TRANSITORIOS):
                raise
            time.sleep(1.0 + 0.8 * i)
    raise ult


def main(caminho):
    copia = os.path.join(os.environ.get('TEMP', '.'),
                         'tq_' + os.path.basename(caminho))
    shutil.copy(caminho, copia)
    xl = novo_excel()
    wb = xl.Workbooks.Open(copia)
    try:
        eq = wb.Worksheets(EQ_ABA)
        ctl = wb.Worksheets('EQA.Controllab_Dados')
        base = wb.Worksheets('EQA_Base')
        base.Visible = -1
        es = wb.Worksheets('Estatística')
        pa = wb.Worksheets('Painel')
        an = wb.Worksheets('Analitos')
        xl.Calculation = -4105

        alvo = str(tenta(lambda: es.Cells(LIN, 1).Value)).strip()
        ano = ANO_TESTE
        tenta(lambda: es.Range('N4').__setattr__('Value', ano))
        ultEQ = tenta(lambda: eq.Cells(eq.Rows.Count, 4).End(-4162).Row)
        ultCTL = tenta(lambda: ctl.Cells(ctl.Rows.Count, 4).End(-4162).Row)

        # O mapa provedor->canonico e o que faz a linha chegar na
        # Estatistica. O analito de prova entra com mapeamento identidade.
        ultMapa = tenta(lambda: base.Cells(base.Rows.Count, 23).End(-4162).Row)
        for k, prov in enumerate(('CAP', 'Controllab')):
            for cc, vv in ((23, prov), (24, alvo), (25, alvo)):
                tenta(lambda c=cc, v=vv, kk=k:
                      base.Cells(ultMapa + 1 + kk, c).__setattr__('Value', v))
        print('analito de prova: %s   nivel %s   ano EP %d   (%s, linha %d)'
              % (alvo, tenta(lambda: es.Cells(LIN, 2).Value), ano,
                 EQ_ABA, ultEQ))

        # -- utilitarios de injecao -----------------------------------------
        # Rebuscar a aba pelo NOME a cada uso. O ponteiro guardado la em
        # cima envelhece quando a estrutura da pasta muda (aba reexibida,
        # tabela redimensionada) e o COM passa a devolver
        # 'Membro nao encontrado' em .Range -- mesmo continuando a atender
        # .Cells. Rebuscar custa nada.
        def folha(nome):
            return wb.Worksheets(nome)

        def bloco(titulo):
            """(linha, coluna) do titulo do bloco -- procurado em toda a area,
            porque o gestor moveu os blocos da coluna J para a R."""
            for r in range(1, 200):
                for c in range(1, 34):
                    v = tenta(lambda x=r, y=c: pa.Cells(x, y).Value)
                    if v and titulo.lower() in str(v).lower():
                        return r, c
            raise SystemExit('bloco %r nao encontrado' % titulo)

        def colDe(titulo, nome, marca='Nível'):
            r0, c0 = bloco(titulo)
            for r in range(r0, r0 + 8):
                if str(tenta(lambda x=r: pa.Cells(x, c0).Value) or '').strip() == marca:
                    for c in range(c0, c0 + 12):
                        t = str(tenta(lambda x=r, y=c: pa.Cells(x, y).Value) or '').strip()
                        if t == nome:
                            return r + 1, c
            raise SystemExit('coluna %r nao achada em %r' % (nome, titulo))

        def limpa_eq():
            for nome, u in ((EQ_ABA, ultEQ),
                            ('EQA.Controllab_Dados', ultCTL)):
                ref = 'A%d:R%d' % (u + 1, u + 40)
                tenta(lambda n=nome, r=ref: folha(n).Range(r).ClearContents())

        def poe_eq(rodada, xlab, xref=100.0, sdgrupo=5.0, prov='CAP',
                   desloc=0, aba='', base_lin=None):
            # estrutura canonica do ADR-034:
            # A provedor | B rodada | C ano | D analito | E amostra |
            # F resultado | G alvo | H SD | I SDI | J lim.inf | K lim.sup
            # O provedor vem da ABA, nao da coluna A: a consolidacao
            # carimba o nome da planilha em que a linha vive. Escrever
            # 'Controllab' numa linha da EQA.CAP_Dados nao a torna do
            # Controllab -- e por isso que a linha vai para a aba certa.
            ws = folha(aba) if aba else folha(EQ_ABA)
            lin = (base_lin if base_lin is not None else ultEQ) + 1 + desloc
            for c, v in ((1, prov), (2, str(rodada)), (3, ano), (4, alvo),
                         (5, '%02d' % (desloc + 1)), (6, xlab), (7, xref),
                         (8, sdgrupo), (9, (xlab - xref) / sdgrupo),
                         (10, 50.0), (11, 150.0), (12, 'Acceptable')):
                tenta(lambda l=lin, cc=c, vv=v, a=ws:
                      a.Cells(l, cc).__setattr__('Value', vv))
            for c, f in ((15, '=IFERROR((F{0}-G{0})/G{0}*100,"")'),
                         (16, '=IF(O{0}="","",ABS(O{0}))')):
                tenta(lambda cc=c, ff=f, l=lin, a=ws:
                      a.Cells(l, cc).__setattr__('Formula', ff.format(l)))

        def cenario(cv, etp, xlab, prov='CAP'):
            """CV e ETp entram como literal; o bias vem de uma rodada de EP."""
            limpa_eq()
            poe_eq(RODADA_TESTE, xlab, prov=prov)
            tenta(lambda: es.Range('L4').__setattr__('Value', prov))
            tenta(lambda: es.Range('P4').__setattr__('Value', RODADA_TESTE))
            tenta(lambda: es.Cells(LIN, C_CV).__setattr__('Value', cv))
            tenta(lambda: an.Cells(4, 20).__setattr__('Value', etp))
            # a EQA_Base e materializada por VBA: sem consolidar, o mCEQ le
            # a base anterior e a prova mede o estado errado
            tenta(lambda: xl.Run('AtualizarEQABase'))
            tenta(lambda: xl.CalculateFull())

        def ler(col):
            return tenta(lambda: es.Cells(LIN, col).Value)

        # ================================================================
        print('\n=== PROVA 1. o bias do EP chega na coluna G ===')
        cenario(cv=2.0, etp=12.0, xlab=104.0)          # bias +4%
        ck('G = |bias| = 4', perto(ler(C_BIAS), 4.0), str(ler(C_BIAS)))
        ck('T guarda o sinal (+4)', perto(tenta(lambda: es.Cells(LIN, 20).Value), 4.0),
           str(tenta(lambda: es.Cells(LIN, 20).Value)))

        print('\n=== PROVA 2. ABS() no ERRO TOTAL (o bias nunca encolhe o ET) ===')
        cenario(cv=2.0, etp=12.0, xlab=108.0)          # bias +8%
        et_pos = ler(C_ET)
        cenario(cv=2.0, etp=12.0, xlab=92.0)           # bias -8%
        et_neg = ler(C_ET)
        esperado = 1.65 * 2.0 + 8.0                    # 11,30
        print('   ET com bias +8%% = %s   ET com bias -8%% = %s   (a mao: %.2f)'
              % (et_pos, et_neg, esperado))
        ck('ET(+8) = 1,65*CV + |bias| = 11,30', perto(et_pos, esperado), str(et_pos))
        ck('ET(-8) = ET(+8): o sinal nao encolhe o ET', perto(et_neg, et_pos),
           '%s vs %s' % (et_neg, et_pos))
        antigo = 1.65 * 2.0 - 8.0                      # o que a formula anterior daria
        ck('a formula anterior daria %.2f -- corrigida' % antigo,
           not perto(et_neg, antigo), 'agora %s' % et_neg)

        print('\n=== PROVA 3. ABS() no SIGMA (bias negativo nao infla o Sigma) ===')
        sig_neg = ler(C_SIGMA)
        cenario(cv=2.0, etp=12.0, xlab=108.0)
        sig_pos = ler(C_SIGMA)
        esperado = (12.0 - 8.0) / 2.0                  # 2,00
        print('   Sigma com bias +8%% = %s   com bias -8%% = %s   (a mao: %.2f)'
              % (sig_pos, sig_neg, esperado))
        ck('Sigma = (ETp - |bias|)/CV = 2,00', perto(sig_pos, esperado), str(sig_pos))
        ck('Sigma(-8) = Sigma(+8)', perto(sig_neg, sig_pos),
           '%s vs %s' % (sig_neg, sig_pos))
        ck('a formula anterior daria %.2f -- corrigida' % ((12.0 + 8.0) / 2.0),
           not perto(sig_neg, 10.0), 'agora %s' % sig_neg)

        print('\n=== PROVA 4. Sigma bate com a conta em varios pontos ===')
        for cv, etp, xlab in ((1.5, 10.0, 103.0), (4.0, 20.0, 95.0), (0.8, 6.0, 101.0)):
            cenario(cv=cv, etp=etp, xlab=xlab)
            b = abs(xlab - 100.0)
            ck('CV=%.1f ETp=%.1f |bias|=%.1f -> Sigma=%.4f'
               % (cv, etp, b, (etp - b) / cv),
               perto(ler(C_SIGMA), (etp - b) / cv), str(ler(C_SIGMA)))

        print('\n=== PROVA 5. as 5 faixas de classificacao, nas bordas ===')
        for etp, sigma, classe in ((12.0, 6.0, 'Classe mundial'),
                                   (11.98, 5.99, 'Excelente'),
                                   (10.0, 5.0, 'Excelente'),
                                   (9.98, 4.99, 'Bom'),
                                   (8.0, 4.0, 'Bom'),
                                   (7.98, 3.99, 'Marginal'),
                                   (6.0, 3.0, 'Marginal'),
                                   (5.98, 2.99, 'Desempenho inadequado')):
            cenario(cv=2.0, etp=etp, xlab=100.0)       # bias 0
            ck('Sigma %.2f -> %s' % (sigma, classe),
               perto(ler(C_SIGMA), sigma, 1e-9) and str(ler(C_CLASSE)) == classe,
               'sigma=%s classe=%r' % (ler(C_SIGMA), ler(C_CLASSE)))

        print('\n=== PROVA 6. Sigma < 3 NAO reprova corrida ===')
        cenario(cv=2.0, etp=5.98, xlab=100.0)          # Sigma 2,99
        wg_antes = str(tenta(lambda: pa.Cells(7, 19).Value))
        tot_antes = tenta(lambda: pa.Cells(7, 18).Value)
        ck('classe = Inadequado', str(ler(C_CLASSE)) == 'Desempenho inadequado', str(ler(C_CLASSE)))
        ck('o veredito da corrida vem de Westgard, nao do Sigma',
           'REPROVA' not in wg_antes or (tot_antes or 0) > 0,
           'status=%r total de violacoes=%s' % (wg_antes, tot_antes))
        # nenhuma formula da pasta pode ler o Sigma para reprovar
        suspeitas = 0
        for r in range(LIN, 94):
            for c in (C_MSTATUS, C_STCV):
                f = str(tenta(lambda rr=r, cc=c: es.Cells(rr, cc).Formula))
                if '$L' in f or '$M' in f:
                    suspeitas += 1
        ck('nenhum status de corrida depende das colunas L/M', suspeitas == 0,
           '%d formulas suspeitas' % suspeitas)

        print('\n=== PROVA 7. a especificacao vem da Analitos S/T/U ===')
        cenario(cv=2.0, etp=17.0, xlab=100.0)
        tenta(lambda: an.Cells(4, 19).__setattr__('Value', 'FAB_TESTE'))
        tenta(lambda: an.Cells(4, 21).__setattr__('Value', 9.75))
        tenta(lambda: xl.CalculateFull())
        ck('I espelha Analitos!S', str(ler(C_FONTE)) == 'FAB_TESTE', str(ler(C_FONTE)))
        ck('K espelha Analitos!T', perto(ler(C_ETP), 17.0), str(ler(C_ETP)))
        ck('J espelha Analitos!U', perto(ler(C_CVTP), 9.75), str(ler(C_CVTP)))

        print('\n=== PROVA 8. Status CV compara CV observado com CVTp ===')
        ck('CV 2,00 <= CVTp 9,75 -> OK', str(ler(C_STCV)) == 'OK', str(ler(C_STCV)))
        tenta(lambda: an.Cells(4, 21).__setattr__('Value', 1.5))
        tenta(lambda: xl.CalculateFull())
        ck('CV 2,00 > CVTp 1,50 -> acusa', 'acima' in str(ler(C_STCV)), str(ler(C_STCV)))

        print('\n=== PROVA 9. margem = ETp - ET, em pontos e em porcentagem ===')
        cenario(cv=2.0, etp=12.0, xlab=104.0)          # ET = 3,3+4 = 7,30
        et, etp = 1.65 * 2.0 + 4.0, 12.0
        ck('margem (p.p.) = %.2f' % (etp - et), perto(ler(C_MPP), etp - et), str(ler(C_MPP)))
        ck('margem %% = %.4f' % ((etp - et) / etp * 100),
           perto(ler(C_MPCT), (etp - et) / etp * 100), str(ler(C_MPCT)))
        ck('ET conferido = %.2f' % et, perto(ler(C_ET), et), str(ler(C_ET)))

        print('\n=== PROVA 10. as 3 faixas de orcamento, nas bordas ===')
        # com CV=2 e bias=0, ET = 3,30 fixo; so o ETp muda
        for etp, sit in ((12.0, 'Dentro do orcamento'),
                         (3.3 / 0.9, 'Margem critica'),        # margem = 10,00%
                         (3.3 / 0.95, 'Margem critica'),       # margem =  5,00%
                         (3.3, 'Margem critica'),              # margem =  0,00%
                         (3.0, 'ETp excedido')):
            cenario(cv=2.0, etp=etp, xlab=100.0)
            ck('ETp=%.4f -> margem %s%% -> %s'
               % (etp, ('%.2f' % ler(C_MPCT)) if ler(C_MPCT) != '' else '-', sit),
               str(ler(C_MSTATUS)) == sit, repr(ler(C_MSTATUS)))

        print('\n=== PROVA 11. sem EP nao vira zero: vira SEM EP, e ET/Sigma ficam vazios ===')
        limpa_eq()
        tenta(lambda: xl.Run('AtualizarEQABase'))
        tenta(lambda: es.Cells(LIN, C_CV).__setattr__('Value', 2.0))
        tenta(lambda: es.Range('P4').__setattr__('Value', 'INEXISTENTE'))
        tenta(lambda: xl.CalculateFull())
        ck('G = "SEM EP" (texto, nao 0)', str(ler(C_BIAS)) == 'SEM EP', repr(ler(C_BIAS)))
        ck('ET fica vazio, nao 0', ler(C_ET) in ('', None), repr(ler(C_ET)))
        ck('Sigma fica vazio, nao 0', ler(C_SIGMA) in ('', None), repr(ler(C_SIGMA)))
        ck('classificacao fica vazia', ler(C_CLASSE) in ('', None), repr(ler(C_CLASSE)))
        ck('margem fica vazia', ler(C_MPP) in ('', None), repr(ler(C_MPP)))
        tenta(lambda: es.Range('P4').__setattr__('Value', 'TODAS'))

        print('\n=== PROVA 12. os filtros de EP ainda mandam nas colunas novas ===')
        limpa_eq()
        poe_eq(RODADA_TESTE + '-A', 104.0, prov='CAP', desloc=0)
        poe_eq(RODADA_TESTE + '-B', 112.0, prov='CAP', desloc=1)
        poe_eq(RODADA_TESTE + '-A', 102.0, prov='Controllab', desloc=0,
               aba='EQA.Controllab_Dados', base_lin=ultCTL)
        tenta(lambda: xl.CalculateFull())
        tenta(lambda: xl.Run('AtualizarEQABase'))
        tenta(lambda: es.Range('L4').__setattr__('Value', 'CAP'))
        for rod, esp in (('TODAS', 8.0), (RODADA_TESTE + '-A', 4.0),
                         (RODADA_TESTE + '-B', 12.0)):
            tenta(lambda r=rod: es.Range('P4').__setattr__('Value', r))
            tenta(lambda: xl.CalculateFull())
            tenta(lambda: xl.Run('AtualizarEQABase'))
            tenta(lambda: xl.CalculateFull())
            ck('CAP rodada %s -> |bias| %.0f' % (rod, esp), perto(ler(C_BIAS), esp),
               str(ler(C_BIAS)))
        tenta(lambda: es.Range('L4').__setattr__('Value', 'Controllab'))
        tenta(lambda: es.Range('P4').__setattr__('Value', 'TODAS'))
        tenta(lambda: xl.CalculateFull())
        ck('trocar de provedor muda o numero', perto(ler(C_BIAS), 2.0), str(ler(C_BIAS)))

        print('\n=== PROVA 13. a lista de criticos pega quem esta sinalizado ===')
        tenta(lambda: es.Range('L4').__setattr__('Value', 'CAP'))
        cenario(cv=2.0, etp=3.0, xlab=100.0)           # este analito estoura o ETp
        tenta(lambda: xl.CalculateFull())
        nExc = tenta(lambda: es.Cells(102, 6).Value)   # RES_R0+2+2 = 100? conferir rotulo
        rot = [str(tenta(lambda r=r: es.Cells(r, 4).Value)) for r in (98, 99, 100)]
        cnt = [tenta(lambda r=r: es.Cells(r, 6).Value) for r in (98, 99, 100)]
        print('   resumo do orcamento: %s' % list(zip(rot, cnt)))
        prim = str(tenta(lambda: es.Cells(108, 1).Value))
        sit = str(tenta(lambda: es.Cells(108, 6).Value))
        print('   primeira linha da lista: %r / %r' % (prim, sit))
        ck('o analito sinalizado aparece na lista', prim == alvo, repr(prim))
        ck('a lista traz a situacao dele', sit in ('ETp excedido', 'Margem critica'),
           repr(sit))
        ck('o contador de ETp excedido e >= 1',
           isinstance(cnt[2], (int, float)) and cnt[2] >= 1, str(cnt[2]))

        print('\n=== PROVA 14. o Painel reproduz a mesma conta ===')
        tenta(lambda: pa.Range('B3').__setattr__('Value', 1))   # analito 1 = alvo
        cenario(cv=2.0, etp=12.0, xlab=108.0)
        tenta(lambda: xl.CalculateFull())
        # ADR-035 moveu os blocos do Painel: o Six Sigma passou a comecar em
        # J12 (cabecalho) com N1 em 13, e o ET vs ETp foi para J26/27. Sem
        # atualizar aqui, o teste lia a LINHA DE CABECALHO e comparava numero
        # com o texto "Sigma" -- foi o que aconteceu.
        rS, cB = colDe('DESEMPENHO SIX SIGMA', 'Bias EQC (abs) %')
        _r, cEtp = colDe('DESEMPENHO SIX SIGMA', 'ETp %')
        _r, cCV = colDe('DESEMPENHO SIX SIGMA', 'CV % obs')
        _r, cSg = colDe('DESEMPENHO SIX SIGMA', 'Sigma')
        _r, cCl = colDe('DESEMPENHO SIX SIGMA', 'Classificação')
        rE, cET = colDe('ERRO TOTAL vs ETp', 'ET %')
        pBias = tenta(lambda: pa.Cells(rS, cB).Value)
        pEtp = tenta(lambda: pa.Cells(rS, cEtp).Value)
        pCV = tenta(lambda: pa.Cells(rS, cCV).Value)
        pSig = tenta(lambda: pa.Cells(rS, cSg).Value)
        pCls = str(tenta(lambda: pa.Cells(rS, cCl).Value))
        pET = tenta(lambda: pa.Cells(rE, cET).Value)
        print('   Painel N1: CV=%s |bias|=%s ETp=%s Sigma=%s %r ET=%s'
              % (pCV, pBias, pEtp, pSig, pCls, pET))
        ck('Painel usa o mesmo |bias| do EP', perto(pBias, 8.0), str(pBias))
        ck('Painel usa o mesmo ETp da Analitos', perto(pEtp, 12.0), str(pEtp))
        if isinstance(pCV, (int, float)) and pCV:
            ck('Painel Sigma = (ETp-|bias|)/CV do proprio periodo',
               perto(pSig, (12.0 - 8.0) / pCV), '%s vs %s' % (pSig, (12.0 - 8.0) / pCV))
            ck('Painel ET = 1,65*CV + |bias|', perto(pET, 1.65 * pCV + 8.0),
               '%s vs %s' % (pET, 1.65 * pCV + 8.0))
        else:
            ck('Painel sem CV no periodo -> Sigma vazio, nao zero',
               pSig in ('', None), repr(pSig))
        ck('a classificacao do Painel usa a mesma escada',
           pCls in ('', 'Classe mundial', 'Excelente', 'Bom', 'Marginal', 'Desempenho inadequado'),
           repr(pCls))

        print('\n=== PROVA 15. nenhuma celula em erro nas areas mexidas ===')
        erros = []
        for r in range(LIN, 94):
            for c in range(7, 21):
                t = str(tenta(lambda rr=r, cc=c: es.Cells(rr, cc).Text))
                if t.startswith('#') and t.strip('#') != '':
                    erros.append('Estatistica!%d,%d=%s' % (r, c, t))
        for r in list(range(96, 125)):
            for c in range(1, 9):
                t = str(tenta(lambda rr=r, cc=c: es.Cells(rr, cc).Text))
                if t.startswith('#') and t.strip('#') != '':
                    erros.append('Estatistica!%d,%d=%s' % (r, c, t))
        for r in list(range(5, 26)):
            for c in range(1, 20):
                t = str(tenta(lambda rr=r, cc=c: pa.Cells(rr, cc).Text))
                if t.startswith('#') and t.strip('#') != '':
                    erros.append('Painel!%d,%d=%s' % (r, c, t))
        ck('zero celulas em erro', not erros, '; '.join(erros[:6]))
    finally:
        try:
            wb.Close(False)
        except Exception:
            pass
        try:
            xl.Quit()
        except Exception:
            pass

    print('\n' + '=' * 66)
    if falhas:
        print('FALHAS (%d):' % len(falhas))
        for f in falhas:
            print('   - %s' % f)
        sys.exit(1)
    print('QUALIDADE: TODAS AS PROVAS PASSARAM')


if __name__ == '__main__':
    main(sys.argv[1])
