# -*- coding: utf-8 -*-
"""acabar_modulo_eqa.py - ADR-034: resumos, derivadas, graficos e acabamento

O QUE FALTAVA DEPOIS DA MIGRACAO

  1. colunas de apoio na forma barata, e o status padronizado materializado
  2. EQA.CAP_Resumo e EQA.Controllab_Resumo, com matriz analito x rodada
     DINAMICA (nada fixado em 2025)
  3. EQA.CAP_Nao_Aceitaveis e EQA.Controllab_Desvios, derivadas por formula
  4. um grafico util em cada resumo
  5. botao de consolidar e aviso de base desatualizada nas abas de digitacao
  6. validacoes da Estatistica realinhadas com as rodadas que existem
  7. EQA_Base escondida (veryHidden)

DUAS ARMADILHAS QUE MOLDARAM ESTE SCRIPT

UDF DENTRO DE SUMPRODUCT RECEBE ARRAY. PadronizarStatus(faixa) nao e chamada
celula a celula: recebe a faixa inteira de uma vez, CStr(array) da erro de tipo
e o resumo todo vira #VALOR!. Por isso o status virou COLUNA (W), calculada uma
vez por linha, e os resumos usam COUNTIFS -- que ainda por cima e muito mais
rapido do que SUMPRODUCT.

ORDINAL POR SUMPRODUCT E O(n^3). A forma anterior era
SUMPRODUCT((COUNTIF(faixa;faixa)=1)*1): para a linha n sao n x n comparacoes.
Com 545 linhas ja pesa; com alguns anos de rodadas trava. A forma nova -- MAX
da propria coluna ate a linha anterior, mais um -- da o mesmo numero com duas
operacoes baratas por linha.

Uso: python acabar_modulo_eqa.py <arquivo.xlsm>
"""
import io
import os
import sys
import time
import subprocess

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', write_through=True)
import win32com.client as w

CRLF = chr(13) + chr(10)

RAIZ = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def cor(r, g, b):
    return r + g * 256 + b * 65536

AZUL_CAB = cor(0x26, 0x3B, 0x4D)
BRANCO = cor(255, 255, 255)
CINZA = cor(0xF2, 0xF4, 0xF7)
VERM_FONTE = cor(0xB4, 0x23, 0x18)
VERDE_FONTE = cor(0x14, 0x6C, 0x43)

NLIN = 5000          # ate onde as formulas dos resumos olham
NRODADAS = 8         # vagas de rodada na matriz
NANALITOS = 60       # vagas de analito na matriz
NDERIV = 200         # vagas nas abas derivadas
NCOLTAB = 24         # A..R canonicas + S..X de apoio
C_ESTADO = 26        # Z, com Y de respiro

CAB_DERIV = ['Provedor', 'Survey / Rodada', 'Ano', 'Analito', 'Amostra',
             'Seu Resultado', 'Média / Valor-alvo', 'SD', 'SDI',
             'Limite Inferior', 'Limite Superior', 'Avaliação',
             'Nº Laboratórios', 'Unidade', 'Bias (%)', '|Bias| (%)',
             'Página Fonte', 'Arquivo Fonte']
LARG_DERIV = [13, 15, 8, 27, 11, 13, 17, 11, 9, 14, 14, 15, 15, 17, 12, 12, 13, 48]
FMT_DERIV = [None, None, '0', None, None, '0.00', '0.000', '0.000',
             '+0.0;-0.0;0.0', '0.00', '0.00', None, '0', None, '0.00', '0.00',
             '0', None]

APOIO = [
    (19, 'ordem_desvio',
     '=IF($W2<>"NAO ACEITO","",COUNTIFS($W$2:$W2,"NAO ACEITO"))'),
    (20, 'ordem_rodada',
     '=IF($B2="","",IF(COUNTIF($B$2:$B2,$B2)=1,MAX($T$1:$T1)+1,""))'),
    (21, 'ordem_analito',
     '=IF($D2="","",IF(COUNTIF($D$2:$D2,$D2)=1,MAX($U$1:$U1)+1,""))'),
    (22, 'par_rodada_analito',
     '=IF($D2="","",IF(COUNTIFS($B$2:$B2,$B2,$D$2:$D2,$D2)=1,1,0))'),
    (23, 'status_padronizado',
     '=IF($D2="","",mEQA.PadronizarStatus($L2))'),
    (24, 'analito_com_desvio',
     '=IF(OR($D2="",$W2<>"NAO ACEITO"),"",'
     'IF(COUNTIFS($D$2:$D2,$D2,$W$2:$W2,"NAO ACEITO")=1,1,0))'),
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


def tenta(fn, vezes=8):
    ult = None
    for i in range(vezes):
        try:
            return fn()
        except Exception as e:
            ult = e
            s = str(e).lower()
            if 'rejeitada' not in s and 'rejected' not in s:
                raise
            time.sleep(1.0 + 0.8 * i)
    raise ult


def aba(wb, nome):
    try:
        return wb.Worksheets(nome)
    except Exception:
        return None


def eh_erro(t):
    t = str(t)
    return t.startswith('#') and t.strip('#') != ''


def col(n):
    s = ''
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def titulo(ws, texto, sub=''):
    tenta(lambda: ws.Cells(1, 1).__setattr__('Value', texto))
    ws.Cells(1, 1).Font.Size = 15
    ws.Cells(1, 1).Font.Bold = True
    if sub:
        ws.Cells(2, 1).Value = sub
        ws.Cells(2, 1).Font.Italic = True


def faixa_cab(ws, linha, textos, col0=1):
    for i, t in enumerate(textos):
        cel = ws.Cells(linha, col0 + i)
        cel.Value = t
        cel.Font.Bold = True
        cel.Font.Color = BRANCO
        cel.Interior.Color = AZUL_CAB
        cel.HorizontalAlignment = -4108
        cel.WrapText = True
    ws.Rows(linha).RowHeight = 30


def montar_resumo(wb, nome, origem, provedor, rot_rodada, rot_ok, rot_ruim,
                  cabecalho_texto):
    """Os dois resumos sao a MESMA analise. Muda o vocabulario, nao a
    profundidade -- a secao 5 da missao proibe um provedor sofisticado e o
    outro simplificado."""
    ws = aba(wb, nome)
    tenta(lambda: ws.Cells.Clear())
    o = "'%s'!" % origem
    B = o + '$B$2:$B$%d' % NLIN      # rodada
    C = o + '$C$2:$C$%d' % NLIN      # ano
    D = o + '$D$2:$D$%d' % NLIN      # analito
    T = o + '$T$2:$T$%d' % NLIN      # ordem_rodada
    U = o + '$U$2:$U$%d' % NLIN      # ordem_analito
    V = o + '$V$2:$V$%d' % NLIN      # par rodada+analito
    W = o + '$W$2:$W$%d' % NLIN      # status padronizado
    X = o + '$X$2:$X$%d' % NLIN      # analito com desvio (1a ocorrencia)

    titulo(ws, cabecalho_texto,
           'Calculado a partir de %s. Nada aqui é digitado, e nenhum ano '
           'está fixado: as colunas seguem as rodadas que existirem.' % origem)

    # ---- bloco 1: desempenho por rodada -------------------------------
    faixa_cab(ws, 4, [rot_rodada, 'Ano', 'Analitos', 'Resultados',
                      rot_ok, rot_ruim, '% ' + rot_ok])
    for i in range(NRODADAS):
        r = 5 + i
        ws.Cells(r, 1).Formula = (
            '=IFERROR(INDEX(%s,MATCH(%d,%s,0)),"")' % (B, i + 1, T))
        ws.Cells(r, 2).Formula = (
            '=IF($A%d="","",IFERROR(INDEX(%s,MATCH($A%d,%s,0)),""))'
            % (r, C, r, B))
        ws.Cells(r, 3).Formula = (
            '=IF($A%d="","",SUMIFS(%s,%s,$A%d))' % (r, V, B, r))
        ws.Cells(r, 4).Formula = '=IF($A%d="","",COUNTIF(%s,$A%d))' % (r, B, r)
        ws.Cells(r, 5).Formula = (
            '=IF($A%d="","",COUNTIFS(%s,$A%d,%s,"ACEITO"))' % (r, B, r, W))
        ws.Cells(r, 6).Formula = (
            '=IF($A%d="","",COUNTIFS(%s,$A%d,%s,"NAO ACEITO"))' % (r, B, r, W))
        ws.Cells(r, 7).Formula = (
            '=IF(OR($A%d="",$D%d=0),"",$E%d/$D%d*100)' % (r, r, r, r))
    rT = 5 + NRODADAS
    ws.Cells(rT, 1).Value = 'TOTAL'
    ws.Cells(rT, 3).Formula = '=SUM(%s)' % V
    ws.Cells(rT, 4).Formula = '=COUNTA(%s)' % B
    ws.Cells(rT, 5).Formula = '=COUNTIF(%s,"ACEITO")' % W
    ws.Cells(rT, 6).Formula = '=COUNTIF(%s,"NAO ACEITO")' % W
    ws.Cells(rT, 7).Formula = '=IF($D%d=0,"",$E%d/$D%d*100)' % (rT, rT, rT)
    ws.Range(ws.Cells(rT, 1), ws.Cells(rT, 7)).Font.Bold = True
    ws.Range(ws.Cells(rT, 1), ws.Cells(rT, 7)).Interior.Color = CINZA

    # ---- bloco 2: indicadores ------------------------------------------
    r0 = rT + 2
    faixa_cab(ws, r0, ['Indicador', 'Valor'], col0=9)
    ind = [
        ('Rodadas', '=COUNT(%s)' % T),
        ('Analitos distintos', '=COUNT(%s)' % U),
        ('Resultados', '=COUNTA(%s)' % B),
        (rot_ruim, '=COUNTIF(%s,"NAO ACEITO")' % W),
        ('Analitos com pelo menos 1 %s' % rot_ruim.lower(), '=SUM(%s)' % X),
        ('Não avaliados pelo provedor', '=COUNTIF(%s,"NAO AVALIADO")' % W),
        ('Sem analito canônico (fora da Estatística)',
         '=SUMPRODUCT((EQA_Base!$A$2:$A$5001="%s")*(EQA_Base!$E$2:$E$5001="")*'
         '(EQA_Base!$D$2:$D$5001<>""))' % provedor),
        ('Fora da análise (Uso_Analitico = NAO)',
         '=SUMPRODUCT((EQA_Base!$A$2:$A$5001="%s")*'
         '(EQA_Base!$T$2:$T$5001="NAO"))' % provedor),
    ]
    for i, (rot, f) in enumerate(ind):
        ws.Cells(r0 + 1 + i, 9).Value = rot
        ws.Cells(r0 + 1 + i, 10).Formula = f
    ws.Columns(9).ColumnWidth = 42
    ws.Columns(10).ColumnWidth = 12

    # ---- bloco 3: matriz analito x rodada, dinamica --------------------
    rM = r0 + 11
    ws.Cells(rM, 1).Value = 'MATRIZ ANALITO × ' + rot_rodada.upper()
    ws.Cells(rM, 1).Font.Bold = True
    ws.Cells(rM, 1).Font.Size = 12
    faixa_cab(ws, rM + 1, ['Analito'] + [''] * NRODADAS + ['Total de amostras'])
    for i in range(NRODADAS):
        ws.Cells(rM + 1, 2 + i).Formula = '=IF($A%d="","",$A%d)' % (5 + i, 5 + i)
    # Tambem por coluna: 60 x 10 celulas uma a uma ja bastam para a instancia
    # automatizada comecar a recusar chamadas.
    r0M = rM + 2
    r1M = rM + 1 + NANALITOS
    tenta(lambda: ws.Range(ws.Cells(r0M, 1), ws.Cells(r1M, 1)).__setattr__(
        'Formula',
        '=IFERROR(INDEX(%s,MATCH(ROW()-%d,%s,0)),"")' % (D, rM + 1, U)))
    for i in range(NRODADAS):
        c = 2 + i
        cab = '%s$%d' % (col(c), rM + 1)
        f = ('=IF(OR($A%d="",%s=""),"",IF(COUNTIFS(%s,%s,%s,$A%d)=0,"—",'
             'COUNTIFS(%s,%s,%s,$A%d)))'
             % (r0M, cab, B, cab, D, r0M, B, cab, D, r0M))
        tenta(lambda cc=c, ff=f: ws.Range(ws.Cells(r0M, cc),
                                          ws.Cells(r1M, cc))
              .__setattr__('Formula', ff))
    tenta(lambda: ws.Range(ws.Cells(r0M, 2 + NRODADAS),
                           ws.Cells(r1M, 2 + NRODADAS)).__setattr__(
        'Formula', '=IF($A%d="","",COUNTIF(%s,$A%d))' % (r0M, D, r0M)))

    ws.Columns(1).ColumnWidth = 30
    for i in range(NRODADAS):
        ws.Columns(2 + i).ColumnWidth = 13
    ws.Columns(2 + NRODADAS).ColumnWidth = 18
    ws.Range(ws.Cells(5, 7), ws.Cells(rT, 7)).NumberFormat = '0.0'
    ws.Activate()
    ws.Range('A5').Select()
    wb.Windows(1).FreezePanes = True
    ws.Range('A1').Select()
    return ws, rT


def montar_derivada(wb, nome, origem, tit, sub):
    """Deriva por formula, e nao por copia: a secao 20 da missao exige que a
    aba acompanhe a origem sozinha, sem banco duplicado."""
    ws = aba(wb, nome)
    tenta(lambda: ws.Cells.Clear())
    o = "'%s'!" % origem
    S = o + '$S$2:$S$%d' % NLIN
    titulo(ws, tit, sub)
    faixa_cab(ws, 4, CAB_DERIV)
    for i, larg in enumerate(LARG_DERIV):
        ws.Columns(i + 1).ColumnWidth = larg
        if FMT_DERIV[i]:
            ws.Range(ws.Cells(5, i + 1),
                     ws.Cells(4 + NDERIV, i + 1)).NumberFormat = FMT_DERIV[i]
    # Uma atribuicao POR COLUNA, e nao 3.600 por celula.
    #
    # Cell a cell o Excel devolveu RPC_E_CALL_REJECTED no meio do caminho: a
    # instancia automatizada fica ocupada e recusa a chamada seguinte. Com
    # ROW()-4 no lugar do indice literal, a mesma formula serve para as 200
    # linhas, e cada coluna vira UMA chamada.
    for c in range(1, 19):
        f = ('=IFERROR(INDEX(%s$2:%s$%d,MATCH(ROW()-4,%s,0)),"")'
             % (o + col(c), col(c), NLIN, S))
        tenta(lambda cc=c, ff=f: ws.Range(ws.Cells(5, cc),
                                          ws.Cells(4 + NDERIV, cc))
              .__setattr__('Formula', ff))
    ws.Cells(3, 1).Formula = (
        '=COUNT(%s)&" registro(s) — a lista acompanha %s sozinha"' % (S, origem))
    ws.Cells(3, 1).Font.Italic = True
    ws.Activate()
    ws.Range('A5').Select()
    wb.Windows(1).FreezePanes = True
    ws.Range('A1').Select()
    return ws


def grafico(ws, rT, tit):
    for ch in list(ws.ChartObjects()):
        ch.Delete()
    co = ws.ChartObjects().Add(620, 70, 430, 250)
    ct = co.Chart
    ct.ChartType = 51                                  # xlColumnClustered
    # Union das duas colunas -- rotulo e valor de uma vez. Criar serie vazia e
    # apagar sobras deixaria o grafico dependendo da ordem em que o Excel
    # numera as series.
    ct.SetSourceData(ws.Application.Union(
        ws.Range(ws.Cells(4, 1), ws.Cells(rT - 1, 1)),
        ws.Range(ws.Cells(4, 6), ws.Cells(rT - 1, 6))))
    ct.HasTitle = True
    ct.ChartTitle.Text = tit
    ct.ChartTitle.Font.Size = 12
    ct.HasLegend = False
    try:
        ct.SeriesCollection(1).Format.Fill.ForeColor.RGB = VERM_FONTE
    except Exception:
        pass
    return co


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
        vbp = wb.VBProject
        for nome in ('mEQA', 'mCEQ'):
            for c in list(vbp.VBComponents):
                if c.Name == nome:
                    vbp.VBComponents.Remove(c)
            importar_bas(vbp, os.path.join(RAIZ, 'src_producao', nome + '.bas'))
        print('modulos reimportados; CarimboEQA = %r'
              % tenta(lambda: xl.Run('CarimboEQA')))

        base = aba(wb, 'EQA_Base')
        if base.Visible != -1:
            base.Visible = -1        # visivel enquanto trabalhamos nela

        # ---- 1. colunas de apoio -----------------------------------------
        for nome, tab in (('EQA.CAP_Dados', 'tblEQA_CAP_Dados'),
                          ('EQA.Controllab_Dados', 'tblEQA_Controllab_Dados')):
            ws = aba(wb, nome)
            lo = ws.ListObjects(tab)
            ult = max(2, lo.Range.Rows.Count)
            if lo.ListColumns.Count < NCOLTAB:
                tenta(lambda l=lo, w2=ws, u=ult: l.Resize(
                    w2.Range(w2.Cells(1, 1), w2.Cells(u, NCOLTAB))))
            for c, cab, f in APOIO:
                cel = ws.Cells(1, c)
                cel.Value = cab
                cel.Font.Bold = True
                cel.Font.Color = BRANCO
                cel.Interior.Color = AZUL_CAB
                tenta(lambda cc=c, ff=f, w2=ws, u=ult:
                      w2.Range(w2.Cells(2, cc), w2.Cells(u, cc))
                      .__setattr__('Formula', ff))
                ws.Columns(c).Hidden = True
            print('%s: %d colunas de apoio, status materializado em W'
                  % (nome, len(APOIO)))

        xl.Calculation = -4105
        tenta(lambda: xl.CalculateFullRebuild())
        xl.Calculation = -4135

        # ---- 2. resumos --------------------------------------------------
        wsR, rT = montar_resumo(
            wb, 'EQA.CAP_Resumo', 'EQA.CAP_Dados', 'CAP',
            'Survey', 'Acceptable', 'Unacceptable',
            'CAP — Resumo da avaliação externa')
        print('EQA.CAP_Resumo montado (TOTAL na linha %d)' % rT)
        wsR2, rT2 = montar_resumo(
            wb, 'EQA.Controllab_Resumo', 'EQA.Controllab_Dados', 'Controllab',
            'Rodada', 'Conforme', 'Com desvio',
            'Controllab — Resumo da avaliação externa')
        print('EQA.Controllab_Resumo montado (TOTAL na linha %d)' % rT2)

        # ---- 3. derivadas -------------------------------------------------
        montar_derivada(
            wb, 'EQA.CAP_Nao_Aceitaveis', 'EQA.CAP_Dados',
            'CAP — Resultados Unacceptable',
            'Derivada de EQA.CAP_Dados. Não digite aqui: corrija na origem.')
        montar_derivada(
            wb, 'EQA.Controllab_Desvios', 'EQA.Controllab_Dados',
            'Controllab — Resultados com desvio',
            'Derivada de EQA.Controllab_Dados. O critério é o do próprio '
            'provedor, e não o vocabulário do CAP.')
        print('duas abas derivadas montadas (%d vagas cada)' % NDERIV)

        # ---- 4. graficos --------------------------------------------------
        xl.Calculation = -4105
        tenta(lambda: xl.CalculateFullRebuild())
        grafico(wsR, rT, 'Unacceptable por survey')
        grafico(wsR2, rT2, 'Resultados com desvio por rodada')
        print('um grafico em cada resumo')

        # ---- 5. botao e aviso nas abas de digitacao ----------------------
        for nome in ('EQA.CAP_Dados', 'EQA.Controllab_Dados'):
            ws = aba(wb, nome)
            for sh in list(ws.Shapes):
                sh.Delete()
            bt = ws.Shapes.AddShape(5, 700, 6, 200, 26)
            bt.TextFrame2.TextRange.Text = 'Consolidar na EQA_Base'
            bt.TextFrame2.TextRange.Font.Size = 10
            bt.TextFrame2.TextRange.Font.Bold = True
            bt.Fill.ForeColor.RGB = AZUL_CAB
            bt.Line.Visible = False
            bt.OnAction = 'AtualizarEQABase'
            bt.Placement = 3                            # xlFreeFloating
            prov = 'CAP' if 'CAP' in nome else 'Controllab'
            ws.Cells(1, C_ESTADO).Value = 'ESTADO DA CONSOLIDAÇÃO'
            ws.Cells(1, C_ESTADO).Font.Bold = True
            ws.Cells(2, C_ESTADO).Formula = '=mEQA.CarimboEQA()'
            # O aviso compara o que ESTA digitado com o que a base ja recebeu
            # deste provedor. Divergiu, a base esta atras -- e quem le a
            # Estatistica precisa saber disso antes de acreditar no numero.
            ws.Cells(3, C_ESTADO).Formula = (
                '=IF(COUNTA($D$2:$D$%d)=COUNTIF(EQA_Base!$A$2:$A$5001,"%s"),'
                '"base em dia",'
                '"BASE DESATUALIZADA — clique em Consolidar na EQA_Base")'
                % (NLIN, prov))
            ws.Cells(3, C_ESTADO).Font.Bold = True
            # FormatConditions.Add sobre o objeto devolvido por Cells() voltou
            # E_INVALIDARG; sobre um Range() montado explicitamente, funciona --
            # e a mesma chamada que o semaforo das abas de digitacao usa.
            #
            # A cor aqui e reforco, nao o recado: quem le "BASE DESATUALIZADA"
            # ja sabe o que fazer sem depender de vermelho. Por isso, se ainda
            # assim falhar, o script AVISA e segue, em vez de morrer por causa
            # de uma formatacao.
            alvo = ws.Range(ws.Cells(3, C_ESTADO), ws.Cells(3, C_ESTADO))
            try:
                alvo.FormatConditions.Delete()
                alvo.FormatConditions.Add(
                    2, None, '=ESQUERDA($Z$3;4)="BASE"').Font.Color = VERM_FONTE
                alvo.FormatConditions.Add(
                    2, None, '=ESQUERDA($Z$3;4)="base"').Font.Color = VERDE_FONTE
            except Exception as e:
                print('   AVISO: sem cor no aviso de %s (%s); o texto continua'
                      % (nome, str(e)[:60]))
            ws.Columns(25).ColumnWidth = 3
            ws.Columns(C_ESTADO).ColumnWidth = 62
            print('%s: botao de consolidar + carimbo e aviso em Z1:Z3' % nome)

        # ---- 6. validacoes da Estatistica ---------------------------------
        ultB = tenta(lambda: base.Cells(base.Rows.Count, 1).End(-4162).Row)
        anos, rodadas, provs = set(), set(), set()
        if ultB >= 2:
            d = tenta(lambda: base.Range(base.Cells(2, 1), base.Cells(ultB, 3)).Value)
            for row in d:
                if row[0]:
                    provs.add(str(row[0]).strip())
                if row[1]:
                    anos.add(int(row[1]))
                if row[2]:
                    rodadas.add(str(row[2]).strip())
        es = aba(wb, 'Estatística')

        # PROTECAO AINDA EXISTE, apesar do ADR-031 ter declarado a pasta
        # liberada. A EQC_Dados apareceu protegida na etapa anterior e a
        # Estatistica recusou Validation.Add com 0x800A03EC. Entao o estado e
        # medido e registrado, e nao suposto -- e a pasta sai desta etapa
        # realmente liberada, que e o que a fase de desenvolvimento pede.
        protegidas = []
        for ws2 in wb.Worksheets:
            if ws2.ProtectContents:
                protegidas.append(ws2.Name)
                for senha in ('qcini2025', None):
                    try:
                        ws2.Unprotect(senha) if senha else ws2.Unprotect()
                        break
                    except Exception:
                        pass
        print('abas que estavam protegidas: %s'
              % (protegidas if protegidas else 'nenhuma'))
        resta = [ws2.Name for ws2 in wb.Worksheets if ws2.ProtectContents]
        if resta:
            raise SystemExit('nao consegui desproteger: %s -- nada salvo' % resta)

        def valida(ref, lista):
            r = es.Range(ref)
            try:
                r.Validation.Delete()
            except Exception:
                pass
            r.Validation.Add(3, 1, 1, lista)
            r.Validation.InCellDropdown = True
            r.Validation.IgnoreBlank = True

        # VIRGULA no protocolo; o ponto-e-virgula so aparece na tela em pt-BR.
        # Escrever ";" aqui gravaria um item unico chamado "CAP;Controllab".
        valida('L4', ','.join(sorted(provs)) or 'CAP')
        valida('N4', ','.join(str(a) for a in sorted(anos)) or '2025')
        valida('P4', 'TODAS,' + ','.join(sorted(rodadas)))
        print('validacoes: provedor %s | anos %s | rodadas %s'
              % (sorted(provs), sorted(anos), sorted(rodadas)))
        if 'CAP' in provs:
            tenta(lambda: es.Range('L4').__setattr__('Value', 'CAP'))
        if anos:
            tenta(lambda: es.Range('N4').__setattr__('Value', max(anos)))
        tenta(lambda: es.Range('P4').__setattr__('Value', 'TODAS'))

        # ---- 7. recalculo e conferencia -----------------------------------
        xl.Calculation = -4105
        t0 = time.time()
        tenta(lambda: xl.CalculateFullRebuild())
        print('recalculo completo: %.1fs' % (time.time() - t0))
        xl.ScreenUpdating = True

        erros = []
        for nome, r1, c1 in (('EQA.CAP_Resumo', 90, 12),
                             ('EQA.Controllab_Resumo', 90, 12),
                             ('EQA.CAP_Nao_Aceitaveis', 40, 18),
                             ('EQA.Controllab_Desvios', 40, 18),
                             ('EQA.CAP_Dados', 30, 26),
                             ('EQA.Controllab_Dados', 5, 26),
                             ('Estatística', 40, 21)):
            ws = aba(wb, nome)
            for r in range(1, r1 + 1):
                for c in range(1, c1 + 1):
                    t = str(tenta(lambda w2=ws, rr=r, cc=c: w2.Cells(rr, cc).Text))
                    if eh_erro(t):
                        erros.append('%s!%s%d=%s' % (nome, col(c), r, t))
        print('celulas em erro: %d' % len(erros))
        for e in erros[:10]:
            print('   %s' % e)
        if erros:
            raise SystemExit('erro de formula -- nada salvo')

        comBias = sum(1 for r in range(14, 94)
                      if isinstance(tenta(lambda rr=r: es.Cells(rr, 7).Value),
                                    (int, float)))
        print('Estatistica: %d linhas com |Bias| numerico vindo da EQA_Base'
              % comBias)
        print('CAP_Resumo TOTAL: analitos=%s resultados=%s ok=%s ruim=%s'
              % (wsR.Cells(rT, 3).Value, wsR.Cells(rT, 4).Value,
                 wsR.Cells(rT, 5).Value, wsR.Cells(rT, 6).Value))
        nDeriv = tenta(lambda: aba(wb, 'EQA.CAP_Nao_Aceitaveis').Cells(3, 1).Value)
        print('CAP_Nao_Aceitaveis: %s' % nDeriv)

        # ---- 8. EQA_Base sai da experiencia do usuario -------------------
        base.Visible = 2                      # xlSheetVeryHidden
        print('EQA_Base: veryHidden')

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


if __name__ == '__main__':
    main(sys.argv[1])
