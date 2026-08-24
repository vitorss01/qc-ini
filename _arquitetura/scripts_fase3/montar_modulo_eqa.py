# -*- coding: utf-8 -*-
"""montar_modulo_eqa.py - ADR-034: as sete abas do modulo de controle externo

  EQA.CAP_Dados            digitacao do CAP, tabela tblEQA_CAP_Dados
  EQA.CAP_Resumo           indicadores + matriz analito x survey (dinamica)
  EQA.CAP_Nao_Aceitaveis   derivada, so Unacceptable
  EQA.Controllab_Dados     digitacao do Controllab, mesma estrutura
  EQA.Controllab_Resumo    mesmos indicadores, terminologia do provedor
  EQA.Controllab_Desvios   derivada, so o que exige atencao
  EQA_Base                 consolidada, veryHidden, alimenta mCEQ e o BI

TRES COLUNAS DE APOIO DENTRO DA TABELA (S, T, U -- ocultas)

As abas derivadas e as matrizes dos resumos precisam responder "qual e o k-esimo
X". Sem coluna de apoio isso exige formula matricial (Ctrl+Shift+Enter) ou
UNIQUE/SORT, que dependem da versao do Excel. As tres colunas numeram a primeira
ocorrencia -- de nao aceitavel, de rodada e de analito -- e o resto vira
INDEX/MATCH comum, que funciona em qualquer versao.

Elas ficam DENTRO da tabela de proposito: coluna calculada de tabela se propaga
sozinha quando o usuario digita uma linha nova. Fora da tabela, ele teria de
arrastar formula -- exatamente o que a secao 31 da missao proibe.

LARGURAS, FORMATOS E CORES saem da planilha de referencia
CAP_C3_2025_COMPLETO_ORGANIZADO.xlsx, lidos do arquivo e nao estimados:
cabecalho #263B4D com fonte branca, SDI com formato de sinal +0,0/-0,0,
Acceptable em verde discreto e Unacceptable em vermelho.

Uso: python montar_modulo_eqa.py <arquivo.xlsm>
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

# --- paleta lida da referencia (RGB -> inteiro BGR do COM) -------------------
def cor(r, g, b):
    return r + g * 256 + b * 65536

AZUL_CAB = cor(0x26, 0x3B, 0x4D)      # #263B4D  cabecalho
BRANCO = cor(255, 255, 255)
VERDE_FUNDO = cor(0xE7, 0xF6, 0xEC)
VERDE_FONTE = cor(0x14, 0x6C, 0x43)
VERM_FUNDO = cor(0xFD, 0xE2, 0xE2)
VERM_FONTE = cor(0xB4, 0x23, 0x18)
GUIA_CAP = cor(0x26, 0x3B, 0x4D)      # familia CAP
GUIA_CTL = cor(0x1F, 0x6F, 0x5C)      # familia Controllab
GUIA_BASE = cor(0x70, 0x70, 0x70)

# --- estrutura canonica das duas abas de digitacao --------------------------
# (coluna, titulo, largura, formato, e formula?)
COLS = [
    (1,  'Provedor',            13,  'General', None),
    (2,  'Survey / Rodada',     15,  'General', None),
    (3,  'Ano',                 8,   '0',       None),
    (4,  'Analito',             27,  'General', None),
    (5,  'Amostra',             11,  'General', None),
    (6,  'Seu Resultado',       13,  '0.00',    None),
    (7,  'Média / Valor-alvo',  17,  '0.000',   None),
    (8,  'SD',                  11,  '0.000',   None),
    (9,  'SDI',                 9,   '+0.0;-0.0;0.0', None),
    (10, 'Limite Inferior',     14,  '0.00',    None),
    (11, 'Limite Superior',     14,  '0.00',    None),
    (12, 'Avaliação',           15,  'General', None),
    (13, 'Nº Laboratórios',     15,  '0',       None),
    (14, 'Unidade',             17,  'General', None),
    (15, 'Bias (%)',            12,  '0.00',    '=IFERROR((F{0}-G{0})/G{0}*100,"")'),
    (16, '|Bias| (%)',          12,  '0.00',    '=IF(O{0}="","",ABS(O{0}))'),
    (17, 'Página Fonte',        13,  '0',       None),
    (18, 'Arquivo Fonte',       48,  'General', None),
]
C_APOIO = [
    (19, 'ordem_desvio',  '=IF($L{0}="","",IF(mEQA.PadronizarStatus($L{0})="NAO ACEITO",'
                          'COUNTIFS($L$2:$L{0},$L{0}),""))'),
    (20, 'ordem_rodada',  '=IF($B{0}="","",IF(COUNTIF($B$2:$B{0},$B{0})=1,'
                          'SUMPRODUCT((COUNTIF($B$2:$B{0},$B$2:$B{0})=1)*1),""))'),
    (21, 'ordem_analito', '=IF($D{0}="","",IF(COUNTIF($D$2:$D{0},$D{0})=1,'
                          'SUMPRODUCT((COUNTIF($D$2:$D{0},$D$2:$D{0})=1)*1),""))'),
]
NCOL = 21

CAB_BASE = ['Provedor', 'Ano', 'Rodada', 'Analito', 'Analito_Canonico', 'Amostra',
            'Resultado', 'Valor_Alvo', 'SD', 'SDI', 'Limite_Inferior',
            'Limite_Superior', 'Avaliacao_Original', 'Status_Padronizado',
            'Unidade', 'Bias', 'Bias_Abs', 'Pagina_Fonte', 'Arquivo_Fonte',
            'Uso_Analitico', 'Chave']

ABAS = [
    ('EQA.CAP_Dados', GUIA_CAP),
    ('EQA.CAP_Resumo', GUIA_CAP),
    ('EQA.CAP_Nao_Aceitaveis', GUIA_CAP),
    ('EQA.Controllab_Dados', GUIA_CTL),
    ('EQA.Controllab_Resumo', GUIA_CTL),
    ('EQA.Controllab_Desvios', GUIA_CTL),
    ('EQA_Base', GUIA_BASE),
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
    # Indexacao direta, e nao varredura da colecao: depois de Add/Move a
    # colecao iterada pode devolver ponteiros mortos.
    try:
        return wb.Worksheets(nome)
    except Exception:
        return None


def garantir(wb, nome, depois):
    ws = aba(wb, nome)
    if ws is None:
        ws = wb.Worksheets.Add(After=depois)
        ws.Name = nome
    return ws


def cabecalho(ws, cols, linha=1):
    for c, titulo, larg, fmt, _f in cols:
        cel = ws.Cells(linha, c)
        cel.Value = titulo
        cel.Font.Bold = True
        cel.Font.Color = BRANCO
        cel.Interior.Color = AZUL_CAB
        cel.HorizontalAlignment = -4108        # xlCenter
        ws.Columns(c).ColumnWidth = larg
        if fmt and fmt != 'General':
            ws.Range(ws.Cells(linha + 1, c), ws.Cells(2000, c)).NumberFormat = fmt


def montar_dados(wb, nome, provedor, tabela, guia):
    ws = aba(wb, nome)
    tenta(lambda: ws.Cells.Clear())
    cabecalho(ws, COLS)
    for c, titulo, formula in C_APOIO:
        cel = ws.Cells(1, c)
        cel.Value = titulo
        cel.Font.Bold = True
        cel.Font.Color = BRANCO
        cel.Interior.Color = AZUL_CAB

    # uma linha semente: ListObjects.Add exige corpo, e a tabela precisa das
    # colunas calculadas gravadas em pelo menos uma linha para propaga-las.
    ws.Cells(2, 1).Value = provedor
    for c, _t, _l, _f, formula in COLS:
        if formula:
            ws.Cells(2, c).Formula = formula.format(2)
    for c, _t, formula in C_APOIO:
        ws.Cells(2, c).Formula = formula.format(2)

    for lo in list(ws.ListObjects):
        lo.Unlist()
    lo = ws.ListObjects.Add(1, ws.Range(ws.Cells(1, 1), ws.Cells(2, NCOL)), None, 1)
    lo.Name = tabela
    lo.TableStyle = 'TableStyleMedium2'

    # as tres colunas de apoio sao aparelho interno, nao informacao do gestor
    for c, _t, _f in C_APOIO:
        ws.Columns(c).Hidden = True

    # cabecalho sempre a vista, e as colunas-chave (provedor..analito) tambem:
    # sem isso, rolar para a coluna de bias faz perder de vista qual analito
    # esta sendo lido.
    ws.Activate()
    ws.Range('E2').Select()
    wb.Windows(1).FreezePanes = True

    ws.Tab.Color = guia
    ws.Rows(1).RowHeight = 28
    ws.Range(ws.Cells(1, 1), ws.Cells(1, NCOL)).VerticalAlignment = -4108
    ws.Range('A1').Select()
    return ws, lo


def semaforo(ws, tabela):
    """Acceptable em verde discreto, Unacceptable em vermelho -- como a
    referencia. A regra le o status PADRONIZADO, nao o texto cru, para o
    Controllab acender o mesmo semaforo com a palavra dele."""
    faixa = ws.Range(ws.Cells(2, 12), ws.Cells(2000, 12))
    tenta(lambda: faixa.FormatConditions.Delete())
    f1 = faixa.FormatConditions.Add(2, None, '=mEQA.PadronizarStatus($L2)="NAO ACEITO"')
    f1.Interior.Color = VERM_FUNDO
    f1.Font.Color = VERM_FONTE
    f1.Font.Bold = True
    f2 = faixa.FormatConditions.Add(2, None, '=mEQA.PadronizarStatus($L2)="ACEITO"')
    f2.Interior.Color = VERDE_FUNDO
    f2.Font.Color = VERDE_FONTE
    # barra de dados no |Bias|, como na referencia: o olho encontra o desvio
    # grande sem precisar ordenar a coluna
    faixa2 = ws.Range(ws.Cells(2, 16), ws.Cells(2000, 16))
    tenta(lambda: faixa2.FormatConditions.Delete())
    faixa2.FormatConditions.AddDatabar()


def main(caminho):
    caminho = os.path.abspath(caminho)
    xl = novo_excel()
    wb = xl.Workbooks.Open(caminho)
    salvou = False
    # A protecao de ESTRUTURA impede Worksheets.Add e o erro que o Excel
    # devolve ("O metodo Add da classe Sheets falhou") nao diz por que. No
    # build da Bioquimica este script roda antes da etapa que tranca; rodando
    # sobre um artefato ja pronto, precisa destrancar e devolver o estado.
    estruturaProtegida = False
    try:
        estruturaProtegida = bool(wb.ProtectStructure)
        if estruturaProtegida:
            wb.Unprotect('qcini2025')
            print('estrutura destravada para a montagem (sera restaurada)')
    except Exception as e:
        raise SystemExit('nao consegui destravar a estrutura: %s' % e)
    if wb.ReadOnly:
        wb.Close(False)
        xl.Quit()
        raise SystemExit('Somente leitura (aberto no Excel?): %s' % caminho)
    try:
        xl.Calculation = -4135
        xl.ScreenUpdating = False

        # ---- 0. mEQA precisa existir ANTES: as formulas o chamam ---------
        vbp = wb.VBProject
        for nome in ('mEQA', 'mCEQ'):
            for c in list(vbp.VBComponents):
                if c.Name == nome:
                    vbp.VBComponents.Remove(c)
            importar_bas(vbp, os.path.join(RAIZ, 'src_producao', nome + '.bas'))
            print('importado: %s' % nome)
        print('   PadronizarStatus("Unacceptable") = %r'
              % tenta(lambda: xl.Run('PadronizarStatus', 'Unacceptable')))

        # ---- 1. criar as sete abas, na ordem, agrupadas ------------------
        nomeAncora = 'EQC_Dados' if aba(wb, 'EQC_Dados') else 'Registros'
        anterior = aba(wb, nomeAncora)
        # Duas etapas separadas: primeiro EXISTIR, depois ORDENAR.
        #
        # Add(After=...) nao honrou a ancora aqui -- a primeira aba foi parar
        # antes de Registros, e nao depois de EQC_Dados. Em vez de confiar no
        # argumento, cada aba e movida para depois da que ocupa um INDICE
        # conhecido, e o indice e relido do resultado. Assim a ordem final e
        # consequencia de uma medida, nao de uma suposicao.
        for nome, _g in ABAS:
            if aba(wb, nome) is None:
                ws = wb.Worksheets.Add()      # posicao aqui nao importa
                ws.Name = nome

        # ARGUMENTO NOMEADO NAO VINCULA NESTE DISPATCH.
        #
        # Add(After=x) ignorou a ancora, e Move(After=x) mandou a aba para uma
        # PASTA NOVA -- que e o que Move faz quando nao recebe Before nem
        # After. A aba sumiu do arquivo sem erro nenhum. So o primeiro
        # parametro POSICIONAL (Before) chega intacto.
        #
        # Entao a ordenacao usa Before: mover cada aba, na ordem, para
        # imediatamente antes de uma referencia fixa que deve ficar DEPOIS do
        # bloco. Cada uma cai antes da referencia e empurra as anteriores para
        # a esquerda, produzindo exatamente a sequencia pedida.
        REFERENCIA = 'Estatística'
        if aba(wb, REFERENCIA) is None:
            raise SystemExit('aba de referencia %r nao existe' % REFERENCIA)
        # ENCADEAMENTO, e nao ancora fixa. Mover cada aba para Before=REFERENCIA
        # depende de o Excel empurrar as anteriores para a esquerda, e esse
        # comportamento nao se sustentou aqui: as sete cairam fora do lugar e
        # em ordem inversa. Encadear -- a primeira antes da referencia, cada
        # seguinte DEPOIS da anterior -- fixa a sequencia sem depender disso.
        anterior = None
        for nome, guia in ABAS:
            ws = aba(wb, nome)
            if anterior is None:
                tenta(lambda s=ws: s.Move(aba(wb, REFERENCIA)))
            else:
                tenta(lambda s=ws, a=anterior: s.Move(None, aba(wb, a)))
            # Move invalida a referencia COM da aba movida: qualquer acesso
            # depois dela devolve 0x800A01A8 (objeto necessario).
            ws = aba(wb, nome)
            if ws is None:
                raise SystemExit('aba %r desapareceu; abas atuais: %s'
                                 % (nome, [x.Name for x in wb.Worksheets]))
            ws.Tab.Color = guia
            anterior = nome
        print('7 abas criadas e ordenadas imediatamente antes de %s' % REFERENCIA)

        # ---- 2. as duas abas de digitacao --------------------------------
        for nome, prov, tab, guia in (
                ('EQA.CAP_Dados', 'CAP', 'tblEQA_CAP_Dados', GUIA_CAP),
                ('EQA.Controllab_Dados', 'Controllab', 'tblEQA_Controllab_Dados', GUIA_CTL)):
            ws, lo = montar_dados(wb, nome, prov, tab, guia)
            semaforo(ws, tab)
            print('%s: tabela %s, %d colunas (3 de apoio ocultas)'
                  % (nome, tab, NCOL))

        # ---- 3. EQA_Base --------------------------------------------------
        base = aba(wb, 'EQA_Base')
        tenta(lambda: base.Cells.Clear())
        for i, t in enumerate(CAB_BASE):
            cel = base.Cells(1, i + 1)
            cel.Value = t
            cel.Font.Bold = True
            cel.Font.Color = BRANCO
            cel.Interior.Color = AZUL_CAB
            base.Columns(i + 1).ColumnWidth = 16
        base.Columns(4).ColumnWidth = 27
        base.Columns(5).ColumnWidth = 27
        base.Columns(19).ColumnWidth = 42
        base.Columns(21).ColumnWidth = 46
        for c, fmt in ((7, '0.00'), (8, '0.000'), (9, '0.000'),
                       (10, '+0.0;-0.0;0.0'), (11, '0.00'), (12, '0.00'),
                       (16, '0.00'), (17, '0.00'), (18, '0')):
            base.Range(base.Cells(2, c), base.Cells(5001, c)).NumberFormat = fmt
        # cabecalho do mapa de analitos, em W:Y
        for i, t in enumerate(['Provedor', 'Nome no provedor', 'Analito canônico']):
            cel = base.Cells(1, 23 + i)
            cel.Value = t
            cel.Font.Bold = True
            cel.Font.Color = BRANCO
            cel.Interior.Color = AZUL_CAB
            base.Columns(23 + i).ColumnWidth = 28
        base.Cells(1, 26).Value = 'ainda nao consolidada'
        base.Columns(26).ColumnWidth = 90
        print('EQA_Base: %d colunas + mapa de analitos em W:Y + carimbo em Z1'
              % len(CAB_BASE))

        # ---- 4. EQC_Dados vira legado, e some da experiencia -------------
        # NAO e apagada: 90 registros continuam la, intactos e auditaveis.
        # So deixa de ser consumida -- o mCEQ agora le a EQA_Base.
        eqc = aba(wb, 'EQC_Dados')
        if eqc is not None:
            # Destravar ANTES de escrever. A ordem antiga escrevia primeiro e
            # so entao media a protecao -- funcionava na Bioquimica porque ali
            # este script roda antes da etapa que protege as abas, e falhava
            # sobre um artefato ja pronto.
            prot = eqc.ProtectContents
            if prot:
                try:
                    eqc.Unprotect('qcini2025')
                except Exception:
                    eqc.Unprotect()
            tenta(lambda: eqc.Cells(1, 1).__setattr__(
                'Value', 'LEGADO (ADR-034) — esta aba não alimenta mais nada. '
                         'Preservada apenas como histórico; o módulo vigente é '
                         'EQA.CAP_Dados / EQA.Controllab_Dados / EQA_Base.'))
            # Font.Bold recusado aqui foi sintoma de aba protegida em etapas
            # anteriores deste projeto; entao o estado e medido, e nao suposto.
            try:
                eqc.Cells(1, 1).Font.Bold = True
                negrito = 'sim'
            except Exception as e:
                negrito = 'recusado (%s)' % str(e)[:40]
            eqc.Visible = 0                    # xlSheetHidden
            print('EQC_Dados: legado, oculta, dados intactos '
                  '(estava protegida: %s ; negrito: %s)' % (prot, negrito))

        xl.Calculation = -4105
        tenta(lambda: xl.CalculateFullRebuild())
        xl.ScreenUpdating = True

        # ---- 5. conferencia ----------------------------------------------
        faltando = [n for n, _g in ABAS if aba(wb, n) is None]
        if faltando:
            raise SystemExit('abas ausentes: %s -- nada salvo' % faltando)
        ordem = [ws.Name for ws in wb.Worksheets]
        print('ordem do arquivo: %s' % ' > '.join(ordem))
        i0 = ordem.index('EQA.CAP_Dados')
        esperada = [n for n, _g in ABAS]
        if ordem[i0:i0 + 7] != esperada:
            raise SystemExit('ordem errada: %s -- nada salvo' % ordem[i0:i0 + 7])
        print('ordem conferida: %s' % ' > '.join(ordem[i0:i0 + 7]))

        wb.Save()
        if estruturaProtegida:
            try:
                wb.Protect('qcini2025', True, False)
                print('estrutura retravada')
            except Exception as e:
                print('AVISO: nao retravei a estrutura: %s' % e)
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
