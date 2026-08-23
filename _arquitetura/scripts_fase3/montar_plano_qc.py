# -*- coding: utf-8 -*-
"""montar_plano_qc.py - ADR-035: do Sigma ate a decisao operacional

O QUE MONTA

  Cfg_PlanoQC (oculta)   tblPlanoQC_Sigma -- as faixas de Sigma com regra,
                         N, run size, frequencia e referencia. UMA tabela;
                         Estatistica, Painel e BI leem dela.
  Estatistica V..AB      DPM teorico, Rendimento teorico, regras, N, run
                         size, cobertura do motor, e a chave analito|nivel
  Estatistica 127+       tabela de referencia Sigma x DPM x Rendimento,
                         CALCULADA pela mesma funcao que a coluna usa
  Painel J10..           os cinco blocos da cadeia de decisao

POR QUE O PAINEL NAO RECALCULA NADA

A secao 36 da missao exige uma fonte da verdade so, e a 45 exige que o card
do Painel e a celula da Estatistica coincidam. As duas coisas so fecham se o
Painel LER a Estatistica em vez de repetir a conta -- e e o que ele faz, por
INDEX/MATCH na chave analito|nivel.

A consequencia precisa ficar dita na tela: o Sigma do Painel responde ao
periodo e ao filtro de EQA definidos na Estatistica, nao ao filtro de datas do
proprio Painel. O filtro do Painel continua mandando no grafico e nos
descritivos, que sao outra pergunta.

Uso: python montar_plano_qc.py <arquivo.xlsm>
"""
import io
import os
import sys
import time
import subprocess

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', write_through=True)
import win32com.client as w

RAIZ = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def cor(r, g, b):
    return r + g * 256 + b * 65536

AZUL = cor(0x26, 0x3B, 0x4D)
BRANCO = cor(255, 255, 255)
CINZA = cor(0xF2, 0xF4, 0xF7)
CINZA_TXT = cor(0x60, 0x6A, 0x78)
VERDE = cor(0x14, 0x6C, 0x43)

EST_R0, EST_RN = 14, 93
REF_R0 = 127                     # tabela educativa na Estatistica

# ------------------------------------------------------------------ tblPlanoQC
# Westgard Sigma Rules with Run Sizes.
#   Sigma_Min | Sigma_Max | Classificacao | Regras | N | RunSize | Frequencia | Ref
#
# Abaixo de 3 Sigma, N e run size ficam VAZIOS de proposito: preencher ali
# sugeriria existir plano de CQ estatistico capaz de sustentar o metodo.
PLANO = [
    (6.0, 999.0, 'Classe mundial', '1_3s', 2, 1000,
     'Até 1000 pacientes entre eventos de CQ',
     'Westgard & Westgard, 2019'),
    (5.0, 6.0, 'Excelente', '1_3s / 2_2s / R_4s', 2, 450,
     'Até 450 pacientes entre eventos de CQ',
     'Westgard & Westgard, 2019'),
    (4.0, 5.0, 'Bom', '1_3s / 2_2s / R_4s / 4_1s', 4, 200,
     'Até 200 pacientes entre eventos de CQ',
     'Westgard & Westgard, 2019'),
    (3.0, 4.0, 'Marginal', '1_3s / 2_2s / R_4s / 4_1s / 6x', 6, 45,
     'Até 45 pacientes entre eventos de CQ',
     'Westgard & Westgard, 2019; Peng et al., 2021'),
    (-999.0, 3.0, 'Desempenho inadequado', '', '', '',
     'CQ estatístico isolado pode ser insuficiente — investigar e melhorar '
     'o desempenho analítico ou reavaliar o método',
     'Westgard et al., 2018; CLSI C24-Ed4'),
]

CAB_PLANO = ['Sigma_Min', 'Sigma_Max', 'Classificacao', 'Regras', 'N_Controle',
             'RunSize_Max', 'Frequencia', 'Referencia']

# ---------------------------------------------------------- colunas novas
NOVAS = [
    (22, 'DPM teórico',
     '=IF(NOT(ISNUMBER($L{0})),"",mPlanoQC.DPMdoSigma($L{0}))', '#.##0'),
    (23, 'Rendimento teórico %',
     '=IF(NOT(ISNUMBER($L{0})),"",mPlanoQC.RendimentoDoSigma($L{0}))', '0,0000'),
    (24, 'Regras Westgard recomendadas',
     '=IF(NOT(ISNUMBER($L{0})),"",mPlanoQC.PlanoQC($L{0},"REGRAS"))', None),
    (25, 'N (medições de controle)',
     '=IF(NOT(ISNUMBER($L{0})),"",mPlanoQC.PlanoQC($L{0},"N"))', '0'),
    (26, 'Run Size máx (pacientes)',
     '=IF(NOT(ISNUMBER($L{0})),"",mPlanoQC.PlanoQC($L{0},"RUNSIZE"))', '#.##0'),
    (27, 'Cobertura do motor Westgard',
     '=IF(NOT(ISNUMBER($L{0})),"",mPlanoQC.CoberturaWestgard($L{0}))', None),
    (28, 'chave analito|nível',
     '=IF($A{0}="","",$A{0}&"|"&$B{0})', None),
]

SIGMAS_REF = [6.0, 5.5, 5.0, 4.5, 4.0, 3.5, 3.0, 2.5, 2.0]

REFS = [
    ('Westgard et al., 2018 — Sigma metrics e tabela Sigma × DPM',
     'https://doi.org/10.11613/BM.2018.020502'),
    ('Westgard & Westgard, 2019 — Sigma Rules with Run Sizes',
     'https://doi.org/10.1093/ajcp/aqy158'),
    ('Peng et al., 2021 — aplicação prática de regras, N e run size',
     'https://doi.org/10.1002/jcla.23665'),
    ('CLSI C24-Ed4 — SQC baseado em risco (base conceitual)',
     'https://clsi.org/standards/products/method-evaluation/documents/c24/'),
]


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


TRANSITORIOS = ('rejeitada', 'rejected', 'membro n', 'member not found', 'busy')


def tenta(fn, vezes=10):
    ult = None
    for i in range(vezes):
        try:
            return fn()
        except Exception as e:
            ult = e
            if not any(t in str(e).lower() for t in TRANSITORIOS):
                raise
            time.sleep(1.0 + 0.8 * i)
    raise ult


def col(n):
    s = ''
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def aba(wb, nome):
    try:
        return wb.Worksheets(nome)
    except Exception:
        return None


def eh_erro(t):
    t = str(t)
    return t.startswith('#') and t.strip('#') != ''


def cab(ws, linha, textos, col0=1, larguras=None):
    for i, t in enumerate(textos):
        c = ws.Cells(linha, col0 + i)
        c.Value = t
        c.Font.Bold = True
        c.Font.Color = BRANCO
        c.Interior.Color = AZUL
        c.HorizontalAlignment = -4108
        c.WrapText = True
        if larguras:
            ws.Columns(col0 + i).ColumnWidth = larguras[i]
    ws.Rows(linha).RowHeight = 32


def titulo(ws, lin, c0, texto, tam=12):
    cel = ws.Cells(lin, c0)
    cel.Value = texto
    cel.Font.Bold = True
    cel.Font.Size = tam


def nota(ws, lin, c0, texto):
    cel = ws.Cells(lin, c0)
    cel.Value = texto
    cel.Font.Italic = True
    cel.Font.Size = 9
    cel.Font.Color = CINZA_TXT


def main(caminho):
    caminho = os.path.abspath(caminho)
    xl = novo_excel()
    wb = xl.Workbooks.Open(caminho)
    salvou = False
    if wb.ReadOnly:
        wb.Close(False)
        xl.Quit()
        raise SystemExit('Somente leitura: %s' % caminho)
    try:
        xl.Calculation = -4135
        xl.ScreenUpdating = False

        # ---- 0. modulos --------------------------------------------------
        vbp = wb.VBProject
        for nome in ('mCEQ', 'mPlanoQC'):
            for c in list(vbp.VBComponents):
                if c.Name == nome:
                    vbp.VBComponents.Remove(c)
            vbp.VBComponents.Import(os.path.join(RAIZ, 'src_producao', nome + '.bas'))
            print('importado: %s' % nome)

        # ---- 1. Cfg_PlanoQC ----------------------------------------------
        cfg = aba(wb, 'Cfg_PlanoQC')
        if cfg is None:
            cfg = wb.Worksheets.Add()
            cfg.Name = 'Cfg_PlanoQC'
        cfg.Visible = -1
        tenta(lambda: cfg.Cells.Clear())
        titulo(cfg, 1, 1, 'tblPlanoQC_Sigma — plano de CQ recomendado por Sigma', 13)
        nota(cfg, 2, 1,
             'Fonte única. Estatística, Painel e Power BI leem daqui — mudar '
             'uma faixa é editar uma linha, e não reescrever fórmula em três abas.')
        cab(cfg, 3, CAB_PLANO, 1, [11, 11, 24, 30, 12, 13, 46, 34])
        for i, linha in enumerate(PLANO):
            for j, v in enumerate(linha):
                tenta(lambda r=4 + i, c=j + 1, vv=v:
                      cfg.Cells(r, c).__setattr__('Value', vv))
        nota(cfg, 4 + len(PLANO) + 1, 1, mPQ_nota_runsize())
        nota(cfg, 4 + len(PLANO) + 2, 1, mPQ_nota_dpm())
        for lo in list(cfg.ListObjects):
            lo.Unlist()
        lo = cfg.ListObjects.Add(
            1, cfg.Range(cfg.Cells(3, 1), cfg.Cells(3 + len(PLANO), 8)), None, 1)
        lo.Name = 'tblPlanoQC_Sigma'
        lo.TableStyle = 'TableStyleLight1'
        cfg.Visible = 0                    # oculta, mas nao veryHidden: e config
        print('Cfg_PlanoQC: tblPlanoQC_Sigma com %d faixas' % len(PLANO))

        # ---- 2. Estatistica: colunas novas -------------------------------
        es = aba(wb, 'Estatística')
        for ws2 in wb.Worksheets:
            if ws2.ProtectContents:
                for senha in ('qcini2025', None):
                    try:
                        ws2.Unprotect(senha) if senha else ws2.Unprotect()
                        break
                    except Exception:
                        pass
        larg = {22: 14, 23: 16, 24: 30, 25: 12, 26: 14, 27: 26, 28: 18}
        for c, tit, f, fmt in NOVAS:
            cel = es.Cells(13, c)
            cel.Value = tit
            cel.Font.Bold = True
            cel.Font.Color = BRANCO
            cel.Interior.Color = AZUL
            cel.WrapText = True
            es.Columns(c).ColumnWidth = larg[c]
            tenta(lambda cc=c, ff=f:
                  es.Range(es.Cells(EST_R0, cc), es.Cells(EST_RN, cc))
                  .__setattr__('Formula', ff.format(EST_R0)))
            if fmt:
                tenta(lambda cc=c, k=fmt:
                      es.Range(es.Cells(EST_R0, cc), es.Cells(EST_RN, cc))
                      .__setattr__('NumberFormatLocal', k))
        es.Columns(28).Hidden = True
        print('Estatistica V..AB: %d colunas novas (AB oculta, e a chave)'
              % len(NOVAS))

        # ---- 3. Estatistica: tabela de referencia Sigma x DPM ------------
        titulo(es, REF_R0, 1,
               'REFERÊNCIA — SIGMA × DPM TEÓRICO × RENDIMENTO TEÓRICO', 12)
        nota(es, REF_R0 + 1, 1,
             'Valores CALCULADOS pela mesma função que alimenta a coluna V. '
             'Conferem com a tabela publicada em Westgard et al., 2018. '
             'Servem de orientação: o DPM do analito usa o Sigma real, sem '
             'arredondar para a linha mais próxima.')
        cab(es, REF_R0 + 2, ['Sigma', 'DPM teórico', 'Rendimento teórico %'],
            1, [12, 16, 22])
        for i, sg in enumerate(SIGMAS_REF):
            r = REF_R0 + 3 + i
            tenta(lambda rr=r, v=sg: es.Cells(rr, 1).__setattr__('Value', v))
            tenta(lambda rr=r: es.Cells(rr, 2).__setattr__(
                'Formula', '=mPlanoQC.DPMdoSigma($A%d)' % rr))
            tenta(lambda rr=r: es.Cells(rr, 3).__setattr__(
                'Formula', '=mPlanoQC.RendimentoDoSigma($A%d)' % rr))
        tenta(lambda: es.Range('A%d:A%d' % (REF_R0 + 3, REF_R0 + 11))
              .__setattr__('NumberFormatLocal', '0,0'))
        tenta(lambda: es.Range('B%d:B%d' % (REF_R0 + 3, REF_R0 + 11))
              .__setattr__('NumberFormatLocal', '#.##0'))
        tenta(lambda: es.Range('C%d:C%d' % (REF_R0 + 3, REF_R0 + 11))
              .__setattr__('NumberFormatLocal', '0,00000'))
        nota(es, REF_R0 + 13, 1, mPQ_nota_dpm())
        print('Estatistica L%d: tabela de referencia com %d pontos'
              % (REF_R0, len(SIGMAS_REF)))

        # ---- 4. Painel ----------------------------------------------------
        pa = aba(wb, 'Painel')
        tenta(lambda: pa.Range('J10:Z70').Clear())
        for c in range(10, 18):
            pa.Columns(c).ColumnWidth = 17
        pa.Columns(10).ColumnWidth = 9

        # o Painel LE a Estatistica: uma fonte da verdade so
        def le(colEst, nivel):
            return ('=IFERROR(INDEX(Estatística!${0}$14:${0}$93,'
                    'MATCH(selAnalito&"|"&{1},Estatística!$AB$14:$AB$93,0)),"")'
                    .format(colEst, nivel))

        titulo(pa, 10, 10, 'DESEMPENHO SIX SIGMA — analito selecionado', 13)
        nota(pa, 11, 10,
             'Valores vindos da aba Estatística. O período e o filtro de EQA '
             '(provedor / ano / rodada) são os definidos lá; o filtro de datas '
             'deste Painel manda no gráfico e nos descritivos, não aqui.')
        cab(pa, 12, ['Nível', 'CV % obs', 'Bias EQC (abs) %', 'ETp %', 'Sigma',
                     'Classificação', 'DPM teórico', 'Rendimento teórico %'], 10,
            [9, 13, 16, 12, 12, 18, 15, 18])
        for i, niv in enumerate((1, 2)):
            r = 13 + i
            tenta(lambda rr=r, n=niv: pa.Cells(rr, 10).__setattr__('Value', 'N%d' % n))
            for j, (cl, fmt) in enumerate((('F', '0,00'), ('G', '0,00'),
                                           ('K', '0,00'), ('L', '0,00'),
                                           ('M', None), ('V', '#.##0'),
                                           ('W', '0,0000'))):
                tenta(lambda rr=r, cc=11 + j, ff=le(cl, niv):
                      pa.Cells(rr, cc).__setattr__('Formula', ff))
                if fmt:
                    tenta(lambda rr=r, cc=11 + j, k=fmt:
                          pa.Cells(rr, cc).__setattr__('NumberFormatLocal', k))
        nota(pa, 15, 10, mPQ_nota_dpm())

        titulo(pa, 17, 10, 'PLANO DE CQ RECOMENDADO PELO SIGMA', 13)
        cab(pa, 18, ['Nível', 'Sigma', 'Regras Westgard', 'N (medições)',
                     'Run Size máx', 'Frequência de CQ',
                     'Cobertura do motor'], 10,
            [9, 12, 28, 14, 14, 30, 24])
        for i, niv in enumerate((1, 2)):
            r = 19 + i
            tenta(lambda rr=r, n=niv: pa.Cells(rr, 10).__setattr__('Value', 'N%d' % n))
            tenta(lambda rr=r, ff=le('L', niv):
                  pa.Cells(rr, 11).__setattr__('Formula', ff))
            tenta(lambda rr=r: pa.Cells(rr, 11).__setattr__('NumberFormatLocal', '0,00'))
            tenta(lambda rr=r, ff=le('X', niv):
                  pa.Cells(rr, 12).__setattr__('Formula', ff))
            tenta(lambda rr=r, ff=le('Y', niv):
                  pa.Cells(rr, 13).__setattr__('Formula', ff))
            tenta(lambda rr=r, ff=le('Z', niv):
                  pa.Cells(rr, 14).__setattr__('Formula', ff))
            tenta(lambda rr=r: pa.Cells(rr, 15).__setattr__(
                'Formula',
                '=IF(NOT(ISNUMBER($K%d)),"",mPlanoQC.PlanoQC($K%d,"FREQUENCIA"))'
                % (rr, rr)))
            tenta(lambda rr=r, ff=le('AA', niv):
                  pa.Cells(rr, 16).__setattr__('Formula', ff))
        nota(pa, 21, 10, mPQ_nota_runsize())
        nota(pa, 22, 10,
             'N é o número TOTAL de medições de controle no evento — não o '
             'número de níveis. Como distribuir entre níveis, materiais e '
             'replicatas depende da configuração do laboratório.')
        nota(pa, 23, 10,
             'Run Size é quantos pacientes podem passar entre eventos de CQ. '
             'Não confundir com a regra R_4s.')

        titulo(pa, 25, 10, 'ERRO TOTAL vs ETp — orçamento de erro', 13)
        cab(pa, 26, ['Nível', 'ET %', 'ETp %', 'Margem (p.p.)', 'Margem %',
                     'Situação'], 10, [9, 13, 13, 15, 13, 22])
        for i, niv in enumerate((1, 2)):
            r = 27 + i
            tenta(lambda rr=r, n=niv: pa.Cells(rr, 10).__setattr__('Value', 'N%d' % n))
            for j, (cl, fmt) in enumerate((('H', '0,00'), ('K', '0,00'),
                                           ('N', '0,00'), ('O', '0,00'),
                                           ('P', None))):
                tenta(lambda rr=r, cc=11 + j, ff=le(cl, niv):
                      pa.Cells(rr, cc).__setattr__('Formula', ff))
                if fmt:
                    tenta(lambda rr=r, cc=11 + j, k=fmt:
                          pa.Cells(rr, cc).__setattr__('NumberFormatLocal', k))

        titulo(pa, 30, 10, 'MARGEM CRÍTICA — todos os analitos', 12)
        for i, (rot, cnt) in enumerate((('ETp excedido', 'ETp excedido'),
                                        ('Margem crítica (≤10%)', 'Margem critica'),
                                        ('Dentro do orçamento', 'Dentro do orcamento'))):
            tenta(lambda rr=31 + i, v=rot: pa.Cells(rr, 10).__setattr__('Value', v))
            tenta(lambda rr=31 + i, k=cnt: pa.Cells(rr, 12).__setattr__(
                'Formula', '=COUNTIF(Estatística!$P$14:$P$93,"%s")' % k))

        titulo(pa, 35, 10, 'SIGMA × DPM × RENDIMENTO — referência', 12)
        cab(pa, 36, ['Sigma', 'DPM teórico', 'Rendimento %'], 10, [9, 15, 16])
        for i, sg in enumerate(SIGMAS_REF):
            r = 37 + i
            tenta(lambda rr=r, v=sg: pa.Cells(rr, 10).__setattr__('Value', v))
            tenta(lambda rr=r: pa.Cells(rr, 11).__setattr__(
                'Formula', '=mPlanoQC.DPMdoSigma($J%d)' % rr))
            tenta(lambda rr=r: pa.Cells(rr, 12).__setattr__(
                'Formula', '=mPlanoQC.RendimentoDoSigma($J%d)' % rr))
        tenta(lambda: pa.Range('J37:J45').__setattr__('NumberFormatLocal', '0,0'))
        tenta(lambda: pa.Range('K37:K45').__setattr__('NumberFormatLocal', '#.##0'))
        tenta(lambda: pa.Range('L37:L45').__setattr__('NumberFormatLocal', '0,000'))

        titulo(pa, 48, 10, 'BASE CIENTÍFICA DO PLANO DE CQ', 12)
        nota(pa, 49, 10,
             'Metodologia: Westgard Sigma Rules with Run Sizes. '
             'Base conceitual de SQC baseado em risco: CLSI C24-Ed4.')
        for i, (texto, url) in enumerate(REFS):
            r = 50 + i
            cel = pa.Cells(r, 10)
            tenta(lambda c=cel: c.__setattr__('Value', ''))
            # Hyperlinks.Add e POSICIONAL: argumento nomeado nao vincula neste
            # dispatch (a mesma armadilha que mandou uma aba para outra pasta
            # no ADR-034).
            tenta(lambda c=cel, u=url, t=texto:
                  pa.Hyperlinks.Add(c, u, '', 'Abrir: ' + u, t + ' ↗'))
        print('Painel: Six Sigma (J12), Plano de CQ (J18), ET vs ETp (J26), '
              'margem (J30), referência (J36), base científica (J48)')

        # ---- 5. recalculo e conferencia -----------------------------------
        xl.Calculation = -4105
        t0 = time.time()
        tenta(lambda: xl.CalculateFullRebuild())
        print('recalculo completo: %.1fs' % (time.time() - t0))
        xl.ScreenUpdating = True

        print('   DPMdoSigma(6)   = %s' % tenta(lambda: xl.Run('DPMdoSigma', 6.0)))
        print('   DPMdoSigma(3)   = %s' % tenta(lambda: xl.Run('DPMdoSigma', 3.0)))
        print('   PlanoQC(5.5,N)  = %s' % tenta(lambda: xl.Run('PlanoQC', 5.5, 'N')))
        print('   Cobertura(3.5)  = %s' % tenta(lambda: xl.Run('CoberturaWestgard', 3.5)))

        erros = []
        for nome, r1, c1 in (('Estatística', 145, 28), ('Painel', 55, 18),
                             ('Cfg_PlanoQC', 12, 8)):
            ws = aba(wb, nome)
            vis = ws.Visible
            ws.Visible = -1
            for r in range(1, r1 + 1):
                for c in range(1, c1 + 1):
                    t = str(tenta(lambda a=ws, x=r, y=c: a.Cells(x, y).Text))
                    if eh_erro(t):
                        erros.append('%s!%s%d=%s' % (nome, col(c), r, t))
            ws.Visible = vis
        print('celulas em erro: %d' % len(erros))
        for e in erros[:10]:
            print('   %s' % e)
        if erros:
            raise SystemExit('erro de formula -- nada salvo')

        wb.Save()
        salvou = True
        print('SALVO: %s' % caminho)
    finally:
        try:
            wb.Close(salvou)
        except Exception:
            pass
        try:
            xl.Quit()
        except Exception:
            pass


def mPQ_nota_dpm():
    return ('DPM teórico estimado pelo Sigma: usa a convenção de short-term '
            'Sigma com deslocamento de 1,5 SD. É um benchmark teórico de '
            'desempenho, e não uma contagem observada de erros em resultados '
            'de pacientes.')


def mPQ_nota_runsize():
    return ('O run size é recomendação de planejamento de SQC baseada em '
            'desempenho Sigma e risco. Não substitui requisitos regulatórios, '
            'de acreditação, instruções do fabricante ou procedimentos '
            'internos mais restritivos.')


if __name__ == '__main__':
    main(sys.argv[1])
