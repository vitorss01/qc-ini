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

CRLF = chr(13) + chr(10)

# Tres faixas: abaixo de 1, quatro casas (senao "zero defeitos"); abaixo de
# 1000, uma casa (para 3,4 nao virar 3); acima, separador de milhar.
FMT_DPM = '[<1]0,0000;[<1000]0,0;#.##0'
FMT_YIELD = '0,0000'
FMT_YIELD_REF = '0,00000'

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
# 8x E A REGRA SEQUENCIAL DA MATRIZ DE DOIS NIVEIS (ADR-038)
#
# Com tres niveis a regra e 6x -- ver PLANO_HEMATOLOGIA. Nao sao sinonimos:
# o numero de medicoes consecutivas faz parte da definicao da regra.
#
# Uma regra de sequencia responde a pergunta: quantos resultados consecutivos
# do mesmo lado da media denunciam desvio sistematico. O laboratorio fechou a
# resposta em oito. E o que a tabela recomenda e o que o motor do Calc avalia.
#
# Abaixo de 3 Sigma, N e run size ficam VAZIOS de proposito: preencher ali
# sugeriria existir plano de CQ estatistico capaz de sustentar o metodo.
# ADR-041: HA UMA MATRIZ POR NUMERO DE NIVEIS, e elas nao sao intercambiaveis.
#
# As faixas existem para casar com a probabilidade de deteccao daquele DESENHO
# de controle. Com tres niveis por corrida o poder estatistico por evento e
# maior, entao a mesma protecao se obtem com menos regras -- e por isso a
# Hematologia sai de "todas as cinco" ja a partir de 4 Sigma, enquanto a
# Bioquimica so sai a partir de 4 com quatro regras e precisa de 5 para tres.
#
# N e o numero de controles por evento; R e quantas vezes o evento se repete.
PLANO_BIOQUIMICA = [
    (6.0, 999.0, 'Classe mundial', '1_3s', 2, 1000,
     'Até 1000 pacientes entre eventos de CQ',
     'Westgard & Westgard, 2019'),
    (5.0, 6.0, 'Excelente', '1_3s / 2_2s / R_4s', 2, 450,
     'Até 450 pacientes entre eventos de CQ',
     'Westgard & Westgard, 2019'),
    (4.0, 5.0, 'Bom', '1_3s / 2_2s / R_4s / 4_1s', 4, 200,
     'Até 200 pacientes entre eventos de CQ',
     'Westgard & Westgard, 2019'),
    (3.0, 4.0, 'Marginal', '1_3s / 2_2s / R_4s / 4_1s / 8x', 6, 45,
     'Até 45 pacientes entre eventos de CQ',
     'Westgard & Westgard, 2019; Peng et al., 2021'),
    # ADR-038: a faixa DECLARA as cinco regras. Antes ficava vazia enquanto o
    # realce acendia as cinco -- cor e texto discordando na mesma tela. N e
    # run size continuam vazios de proposito: nao ha plano estatistico que
    # sustente o metodo abaixo de 3 Sigma, e numero ali sugeriria que ha.
    (-999.0, 3.0, 'Desempenho inadequado', '1_3s / 2_2s / R_4s / 4_1s / 8x',
     '', '',
     'CQ estatístico isolado pode ser insuficiente — investigar e melhorar '
     'o desempenho analítico ou reavaliar o método',
     'Westgard et al., 2018; CLSI C24-Ed4'),
]


# Hematologia -- tres niveis de controle. Regras 2of3_2s, 3_1s e 6x; 8x NAO
# aparece aqui, e nao e sinonimo de 6x: o numero de medicoes consecutivas faz
# parte da definicao da regra.
#
# N = 3 em todas as faixas porque o desenho ja mede os tres niveis por evento.
# O que muda entre faixas e QUANTAS REGRAS rodam e, abaixo de 4 Sigma, quantas
# VEZES o evento se repete (R).
PLANO_HEMATOLOGIA = [
    (6.0, 999.0, 'Classe mundial', '1_3s', 3, 1,
     'Uma medicao em cada um dos 3 niveis por evento de CQ',
     'Westgard & Westgard, 2019'),
    (5.0, 6.0, 'Excelente', '1_3s / 2of3_2s / R_4s', 3, 1,
     'Uma medicao em cada um dos 3 niveis por evento de CQ',
     'Westgard & Westgard, 2019'),
    (4.0, 5.0, 'Bom', '1_3s / 2of3_2s / R_4s / 3_1s', 3, 1,
     'Uma medicao em cada um dos 3 niveis por evento de CQ',
     'Westgard & Westgard, 2019'),
    # DUAS faixas abaixo de 4, com AS MESMAS cinco regras.
    #
    # A matriz de tres niveis nao muda de regra entre 3 e 4 Sigma -- ja pede
    # as cinco. Mas a CLASSIFICACAO DE DESEMPENHO e outra pergunta, e ela tem
    # cinco faixas em todo o QC_INI: 3 a <4 e "Marginal", abaixo de 3 e
    # "Desempenho inadequado".
    #
    # Colapsar as duas numa faixa so obrigaria mQualidade.ClassificarSigma --
    # que le a classificacao DESTA tabela -- a chamar um metodo de Sigma 3,5
    # de "Desempenho inadequado", divergindo da Bioquimica para o mesmo
    # numero. Separar aqui mantem UMA escada de classificacao no projeto
    # inteiro, sem tabela paralela.
    (3.0, 4.0, 'Marginal', '1_3s / 2of3_2s / R_4s / 3_1s / 6x', 6, 1,
     'N=6 / R=1, ou N=3 / R=2 conforme a estrategia operacional do laboratorio',
     'Westgard & Westgard, 2019; Peng et al., 2021'),
    # Abaixo de 3 Sigma as regras continuam as cinco, e N e run size ficam
    # VAZIOS: nao ha plano estatistico que sustente o metodo, e numero ali
    # sugeriria que ha. O caminho e investigar o metodo.
    (-999.0, 3.0, 'Desempenho inadequado', '1_3s / 2of3_2s / R_4s / 3_1s / 6x',
     '', '',
     'CQ estatistico isolado pode ser insuficiente - investigar e melhorar '
     'o desempenho analitico ou reavaliar o metodo',
     'Westgard et al., 2018; CLSI C24-Ed4'),
]

PLANOS = {
    'Bioquimica': PLANO_BIOQUIMICA,
    'Hematologia': PLANO_HEMATOLOGIA,
}

# Resolvido em main() pelo arquivo alvo. O default existe apenas para o modulo
# poder ser importado; rodar sem passar pelo main escreveria a matriz errada.
PLANO = PLANO_BIOQUIMICA


def plano_do_arquivo(caminho):
    """O produto sai do NOME do arquivo alvo. Escrever a matriz errada nao
    quebra nada visivelmente -- so publica um plano de CQ que nao corresponde
    ao desenho de controle daquele setor."""
    nome = os.path.basename(caminho).lower()
    for produto, plano in PLANOS.items():
        if produto.lower() in nome:
            return produto, plano
    raise SystemExit('nao reconheci o produto em %s' % os.path.basename(caminho))


CAB_PLANO = ['Sigma_Min', 'Sigma_Max', 'Classificacao', 'Regras', 'N_Controle',
             'RunSize_Max', 'Frequencia', 'Referencia']

# ---------------------------------------------------------- colunas novas
NOVAS = [
    # abaixo de 1 defeito por milhao, quatro casas: senao Sigma 6,99 exibe
    # "0", que se le como "zero defeitos" em vez de "menos de um"
    (22, 'DPM teórico',
     '=IF(NOT(ISNUMBER($L{0})),"",mPlanoQC.DPMdoSigma($L{0}))',
     FMT_DPM),
    (23, 'Rendimento teórico %',
     '=IF(NOT(ISNUMBER($L{0})),"",mPlanoQC.RendimentoDoSigma($L{0}))',
     FMT_YIELD),
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



def importar_bas(vbp, caminho):
    """Importa um .bas convertendo o arquivo para cp1252.

    O VBA le .bas como ANSI da pagina de codigo do sistema. Um arquivo gravado
    em UTF-8 chega com cada acento virando dois caracteres: "nao aplicavel"
    escrito com til e acento apareceu na tela do Painel como dois simbolos por
    letra. Converter na hora da importacao resolve para todos os modulos de uma
    vez, e mantem os fontes legiveis em UTF-8 no repositorio.
    """
    import tempfile
    texto = io.open(caminho, encoding='utf-8', errors='replace').read()
    tmp = os.path.join(tempfile.gettempdir(),
                       'ansi_' + os.path.basename(caminho))
    with io.open(tmp, 'w', encoding='cp1252', errors='replace',
                 newline=CRLF) as fh:
        fh.write(texto)
    vbp.VBComponents.Import(tmp)


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
    global PLANO
    produto, PLANO = plano_do_arquivo(caminho)
    print('produto: %s (%d faixas de Sigma)' % (produto, len(PLANO)))
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
        for nome in ('mCEQ', 'mPlanoQC', 'mQualidade'):
            for c in list(vbp.VBComponents):
                if c.Name == nome:
                    vbp.VBComponents.Remove(c)
            importar_bas(vbp, os.path.join(RAIZ, 'src_producao', nome + '.bas'))
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

        # As colunas G..T nasceram com NumberFormat = "0.00" no ADR-033, e
        # nesta instalacao o Excel le esse codigo com convencao pt-BR, onde o
        # ponto e separador de MILHAR. O |Bias| de 2,70 aparecia como "003" e
        # o Sigma de 4,25 como "004". Vai por NumberFormatLocal, com virgula.
        FMT_EST = {7: '0,00', 8: '0,00', 10: '0,00', 11: '0,00', 12: '0,00',
                   14: '0,00', 15: '0,00', 20: '0,00',
                   3: '0', 4: '0,0000', 5: '0,0000', 6: '0,00'}
        for c, k in FMT_EST.items():
            tenta(lambda cc=c, kk=k:
                  es.Range(es.Cells(EST_R0, cc), es.Cells(EST_RN, cc))
                  .__setattr__('NumberFormatLocal', kk))
        print('Estatistica: %d colunas antigas reformatadas por '
              'NumberFormatLocal' % len(FMT_EST))
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
              .__setattr__('NumberFormatLocal', FMT_DPM))
        tenta(lambda: es.Range('C%d:C%d' % (REF_R0 + 3, REF_R0 + 11))
              .__setattr__('NumberFormatLocal', FMT_YIELD_REF))
        nota(es, REF_R0 + 13, 1, mPQ_nota_dpm())
        print('Estatistica L%d: tabela de referencia com %d pontos'
              % (REF_R0, len(SIGMAS_REF)))

        # ---- 3b. o rotulo do resumo acompanha a tabela --------------------
        #
        # tblPlanoQC_Sigma chama a faixa de "Desempenho inadequado"; a escada
        # antiga do mQualidade chamava "Inadequado". Duas etiquetas para a
        # mesma faixa fazem o COUNTIF do resumo contar zero. Agora a tabela
        # manda, e o rotulo do resumo e ajustado para bater com ela.
        trocados = 0
        for r in range(96, 106):
            v = tenta(lambda rr=r: es.Cells(rr, 1).Value)
            if str(v).strip() == 'Inadequado':
                tenta(lambda rr=r: es.Cells(rr, 1).__setattr__(
                    'Value', 'Desempenho inadequado'))
                trocados += 1
        print('resumo da Estatistica: %d rotulo(s) alinhado(s) com a tabela'
              % trocados)

        # ---- 3c. o eco do filtro apontava para a aba LEGADA ---------------
        #
        # A formula em K5 contava linhas da EQC_Dados, que continua na pasta
        # com os 90 registros de simulacao. Com CAP/2025/TODAS ela dizia
        # "90 linha(s) no banco de EP" enquanto a EQA_Base tem 545 -- numero
        # errado, e vindo da fonte que o mCEQ nem le mais.
        tenta(lambda: es.Range('K5').__setattr__(
            'Formula',
            '=ResumoFiltroEQ(eqProvedor,eqAnoEP,eqRodada)&" | "&'
            'TEXT(SUMPRODUCT((EQA_Base!$A$2:$A$5001=eqProvedor)*'
            '(EQA_Base!$B$2:$B$5001=eqAnoEP)*'
            '(EQA_Base!$T$2:$T$5001="SIM")*'
            'IF(eqRodada="TODAS",1,--(EQA_Base!$C$2:$C$5001=eqRodada))),"0")&'
            '" linha(s) analíticas na EQA_Base"'))
        print('Estatistica K5: eco repontado para a EQA_Base')

        # ---- 3d. a lista de rodadas so oferece o que alimenta analise -----
        #
        # As rodadas A, B e C vinham dos 90 registros de simulacao, marcados
        # com Uso_Analitico = NAO. Oferece-las no seletor convida o gestor a
        # escolher um filtro que nao devolve numero nenhum.
        baseA = aba(wb, 'EQA_Base')
        visB = baseA.Visible
        baseA.Visible = -1
        ultB = tenta(lambda: baseA.Cells(baseA.Rows.Count, 1).End(-4162).Row)
        anos, rods, provs = set(), set(), set()
        if ultB >= 2:
            dd = tenta(lambda: baseA.Range(baseA.Cells(2, 1),
                                           baseA.Cells(ultB, 20)).Value)
            for row in dd:
                if str(row[19]).strip().upper() != 'SIM':
                    continue
                if row[0]:
                    provs.add(str(row[0]).strip())
                if row[1]:
                    anos.add(int(row[1]))
                if row[2]:
                    rods.add(str(row[2]).strip())
        baseA.Visible = visB

        def valida(ref, lista):
            r = es.Range(ref)
            try:
                r.Validation.Delete()
            except Exception:
                pass
            r.Validation.Add(3, 1, 1, lista)
            r.Validation.InCellDropdown = True
            r.Validation.IgnoreBlank = True

        valida('L4', ','.join(sorted(provs)) or 'CAP')
        valida('N4', ','.join(str(a) for a in sorted(anos)) or '2025')
        valida('P4', 'TODAS,' + ','.join(sorted(rods)))
        print('seletores: provedor %s | anos %s | rodadas %s'
              % (sorted(provs), sorted(anos), sorted(rods)))
        tenta(lambda: es.Range('L4').__setattr__('Value', 'CAP'))
        tenta(lambda: es.Range('N4').__setattr__(
            'Value', max(anos) if anos else 2025))
        tenta(lambda: es.Range('P4').__setattr__('Value', 'TODAS'))

        # ---- 4. Painel ----------------------------------------------------
        pa = aba(wb, 'Painel')
        tenta(lambda: pa.Range('J10:Z70').Clear())
        tenta(lambda: pa.Range('J10:Z120').Clear())
        for c in range(10, 18):
            pa.Columns(c).ColumnWidth = 17
        pa.Columns(10).ColumnWidth = 9

        # ONDE OS BLOCOS PODEM COMECAR
        #
        # Os graficos de Levey-Jennings tem 1582 px de largura a partir de x=0:
        # atravessam a coluna J inteira. Qualquer bloco colocado em J10 fica
        # DEBAIXO deles -- invisivel, ainda que a formula esteja certa. A
        # primeira linha livre e medida, e nao arbitrada.
        fundo = 0.0
        for sh in list(pa.Shapes):
            if sh.Width > 400:            # os graficos, nao os controles
                fundo = max(fundo, sh.Top + sh.Height)
        BASE = 10
        for r in range(10, 200):
            if tenta(lambda x=r: pa.Rows(x).Top) > fundo + 12:
                BASE = r
                break
        print('graficos terminam em y=%.0f; blocos comecam na linha %d'
              % (fundo, BASE))

        # o Painel LE a Estatistica: uma fonte da verdade so
        def le(colEst, nivel):
            return ('=IFERROR(INDEX(Estatística!${0}$14:${0}$93,'
                    'MATCH(selAnalito&"|"&{1},Estatística!$AB$14:$AB$93,0)),"")'
                    .format(colEst, nivel))

        titulo(pa, BASE, 10, 'DESEMPENHO SIX SIGMA — analito selecionado', 13)
        nota(pa, BASE + 1, 10,
             'Valores vindos da aba Estatística. O período e o filtro de EQA '
             '(provedor / ano / rodada) são os definidos lá; o filtro de datas '
             'deste Painel manda no gráfico e nos descritivos, não aqui.')
        cab(pa, BASE + 2, ['Nível', 'CV % obs', 'Bias EQC (abs) %', 'ETp %', 'Sigma',
                     'Classificação', 'DPM teórico', 'Rendimento teórico %'], 10,
            [9, 13, 16, 12, 12, 18, 15, 18])
        for i, niv in enumerate((1, 2)):
            r = BASE + 3 + i
            tenta(lambda rr=r, n=niv: pa.Cells(rr, 10).__setattr__('Value', 'N%d' % n))
            for j, (cl, fmt) in enumerate((('F', '0,00'), ('G', '0,00'),
                                           ('K', '0,00'), ('L', '0,00'),
                                           ('M', None), ('V', FMT_DPM),
                                           ('W', FMT_YIELD))):
                tenta(lambda rr=r, cc=11 + j, ff=le(cl, niv):
                      pa.Cells(rr, cc).__setattr__('Formula', ff))
                if fmt:
                    tenta(lambda rr=r, cc=11 + j, k=fmt:
                          pa.Cells(rr, cc).__setattr__('NumberFormatLocal', k))
        nota(pa, BASE + 5, 10, mPQ_nota_dpm())

        titulo(pa, BASE + 7, 10, 'PLANO DE CQ RECOMENDADO PELO SIGMA', 13)
        cab(pa, BASE + 8, ['Nível', 'Sigma', 'Regras Westgard', 'N (medições)',
                     'Run Size máx', 'Frequência de CQ',
                     'Cobertura do motor'], 10,
            [9, 12, 28, 14, 14, 30, 24])
        for i, niv in enumerate((1, 2)):
            r = BASE + 9 + i
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
        nota(pa, BASE + 11, 10, mPQ_nota_runsize())
        nota(pa, BASE + 12, 10,
             'N é o número TOTAL de medições de controle no evento — não o '
             'número de níveis. Como distribuir entre níveis, materiais e '
             'replicatas depende da configuração do laboratório.')
        nota(pa, BASE + 13, 10,
             'Run Size é quantos pacientes podem passar entre eventos de CQ. '
             'Não confundir com a regra R_4s.')

        titulo(pa, BASE + 15, 10, 'ERRO TOTAL vs ETp — orçamento de erro', 13)
        cab(pa, BASE + 16, ['Nível', 'ET %', 'ETp %', 'Margem (p.p.)', 'Margem %',
                     'Situação'], 10, [9, 13, 13, 15, 13, 22])
        for i, niv in enumerate((1, 2)):
            r = BASE + 17 + i
            tenta(lambda rr=r, n=niv: pa.Cells(rr, 10).__setattr__('Value', 'N%d' % n))
            for j, (cl, fmt) in enumerate((('H', '0,00'), ('K', '0,00'),
                                           ('N', '0,00'), ('O', '0,00'),
                                           ('P', None))):
                tenta(lambda rr=r, cc=11 + j, ff=le(cl, niv):
                      pa.Cells(rr, cc).__setattr__('Formula', ff))
                if fmt:
                    tenta(lambda rr=r, cc=11 + j, k=fmt:
                          pa.Cells(rr, cc).__setattr__('NumberFormatLocal', k))

        titulo(pa, BASE + 20, 10, 'MARGEM CRÍTICA — todos os analitos', 12)
        for i, (rot, cnt) in enumerate((('ETp excedido', 'ETp excedido'),
                                        ('Margem crítica (≤10%)', 'Margem critica'),
                                        ('Dentro do orçamento', 'Dentro do orcamento'))):
            tenta(lambda rr=BASE + 21 + i, v=rot: pa.Cells(rr, 10).__setattr__('Value', v))
            tenta(lambda rr=BASE + 21 + i, k=cnt: pa.Cells(rr, 12).__setattr__(
                'Formula', '=COUNTIF(Estatística!$P$14:$P$93,"%s")' % k))

        titulo(pa, BASE + 25, 10, 'SIGMA × DPM × RENDIMENTO — referência', 12)
        cab(pa, BASE + 26, ['Sigma', 'DPM teórico', 'Rendimento %'], 10, [9, 15, 16])
        for i, sg in enumerate(SIGMAS_REF):
            r = BASE + 27 + i
            tenta(lambda rr=r, v=sg: pa.Cells(rr, 10).__setattr__('Value', v))
            tenta(lambda rr=r: pa.Cells(rr, 11).__setattr__(
                'Formula', '=mPlanoQC.DPMdoSigma($J%d)' % rr))
            tenta(lambda rr=r: pa.Cells(rr, 12).__setattr__(
                'Formula', '=mPlanoQC.RendimentoDoSigma($J%d)' % rr))
        tenta(lambda: pa.Range('J%d:J%d' % (BASE + 27, BASE + 35)).__setattr__('NumberFormatLocal', '0,0'))
        tenta(lambda: pa.Range('K%d:K%d' % (BASE + 27, BASE + 35)).__setattr__('NumberFormatLocal', FMT_DPM))
        tenta(lambda: pa.Range('L%d:L%d' % (BASE + 27, BASE + 35)).__setattr__('NumberFormatLocal', FMT_YIELD_REF))

        titulo(pa, BASE + 38, 10, 'BASE CIENTÍFICA DO PLANO DE CQ', 12)
        nota(pa, BASE + 39, 10,
             'Metodologia: Westgard Sigma Rules with Run Sizes. '
             'Base conceitual de SQC baseado em risco: CLSI C24-Ed4.')
        for i, (texto, url) in enumerate(REFS):
            r = BASE + 40 + i
            cel = pa.Cells(r, 10)
            tenta(lambda c=cel: c.__setattr__('Value', ''))
            # Hyperlinks.Add e POSICIONAL: argumento nomeado nao vincula neste
            # dispatch (a mesma armadilha que mandou uma aba para outra pasta
            # no ADR-034).
            tenta(lambda c=cel, u=url, t=texto:
                  pa.Hyperlinks.Add(c, u, '', 'Abrir: ' + u, t + ' ↗'))
        print('Painel a partir da linha %d: Six Sigma, plano de CQ, ET vs ETp, margem, referencia e base cientifica' % BASE)  # noqa
        print('(antigo) Six Sigma (J12), Plano de CQ (J18), ET vs ETp (J26), '
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
        for nome, r1, c1 in (('Estatística', 145, 28), ('Painel', BASE + 45, 18),
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
