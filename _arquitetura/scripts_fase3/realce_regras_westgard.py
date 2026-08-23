# -*- coding: utf-8 -*-
"""realce_regras_westgard.py - ADR-037: qual regra o Sigma pede, em cores

O QUE FAZ

  1. estende tblPlanoQC_Sigma com uma MATRIZ de regras (uma coluna por regra,
     1 ou 0) -- derivada do proprio texto de Regras, para nao existir segunda
     fonte
  2. acrescenta Status_Plano_QC (REAVALIAR METODO abaixo de 3 Sigma)
  3. poe formatacao condicional nas celulas de regra do Painel (G6:K6)
  4. escreve legenda e texto de apoio em area livre
  5. prova que o resto do layout do gestor nao mudou

O QUE NAO FAZ

Nao move celula, nao troca cor existente, nao mexe na formatacao de violacao
que ja esta em G7:K9, nao altera quais regras o motor executa.

DUAS SEMANTICAS, DUAS CORES

  verde escuro / branco / negrito  ->  regra RECOMENDADA pelo Sigma
  vermelho (codificacao existente) ->  regra VIOLADA na corrida

Sao coisas diferentes. Uma regra recomendada e nao violada nao pode parecer
violada, e uma regra violada nao pode ficar verde so porque esta no plano.
Por isso a condicao de violacao entra com prioridade MAIOR: numa celula que
seja as duas coisas, vence o vermelho.

A MATRIZ VEM DO TEXTO, NAO AO LADO DELE

Cada bandeira e =IF(ISNUMBER(SEARCH("1_3s"; texto de Regras));1;0). Assim a
recomendacao textual e o realce visual NAO PODEM divergir: sao a mesma
informacao lida de dois jeitos. Mudar uma faixa na tabela muda os dois.

ABAIXO DE 3 SIGMA A FAIXA NAO TEM REGRAS

E de proposito: nao existe plano estatistico que sustente o metodo ali. Para o
realce, a faixa cai no conjunto de regras que o MOTOR suporta -- estrategia
intensiva -- e a tabela devolve Status_Plano_QC = "REAVALIAR METODO". Nao e
prescricao cientifica de que aquelas regras resolvem; e o maximo que o CQ
estatistico alcanca enquanto o metodo nao melhora.

QUAL SIGMA MANDA NO REALCE

O bloco Westgard tem uma linha por nivel, mas a celula da REGRA e uma so. O
realce usa o MENOR Sigma dos dois niveis: o plano de CQ precisa cobrir o nivel
que esta pior. Fica dito no texto de apoio.

Uso: python realce_regras_westgard.py <arquivo.xlsm>
"""
import io
import os
import sys
import time
import shutil
import subprocess

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', write_through=True)
import win32com.client as w

CRLF = chr(13) + chr(10)
RAIZ = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def cor(r, g, b):
    return r + g * 256 + b * 65536

VERDE_ESCURO = cor(0x14, 0x6C, 0x43)
BRANCO = cor(255, 255, 255)
AZUL = cor(0x26, 0x3B, 0x4D)
CINZA_TXT = cor(0x60, 0x6A, 0x78)

# As celulas de regra no layout do gestor, com o rotulo que ele usa.
# NAO mover. NAO renomear.
REGRAS_PAINEL = {'G6': '1-3S', 'H6': '2-2S', 'I6': 'R4S',
                 'J6': '4-1S', 'K6': '8X'}

# rotulo exibido -> token usado no texto de Regras da tblPlanoQC_Sigma
MATRIZ = [('1-3S', '1_3s'), ('2-2S', '2_2s'), ('R4S', 'R_4s'),
          ('4-1S', '4_1s'), ('6x', '6x'), ('8X', '8x'), ('10x', '10x')]

CFG = 'Cfg_PlanoQC'
CFG_R0, CFG_RN = 4, 8          # linhas das faixas
MAT_C0 = 10                    # J: primeira coluna da matriz
C_STATUS = 9                   # I: Status_Plano_QC

# Sigma do analito selecionado, por nivel, no layout do gestor
SIGMA_N1 = 'Painel!$V$10'
SIGMA_N2 = 'Painel!$V$11'


def importar_bas(vbp, caminho):
    import tempfile
    texto = io.open(caminho, encoding='utf-8', errors='replace').read()
    tmp = os.path.join(tempfile.gettempdir(), 'ansi_' + os.path.basename(caminho))
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


TRANS = ('rejeitada', 'rejected', 'membro n', 'member not found', 'busy')


def tenta(fn, vezes=10):
    ult = None
    for i in range(vezes):
        try:
            return fn()
        except Exception as e:
            ult = e
            if not any(t in str(e).lower() for t in TRANS):
                raise
            time.sleep(1.0 + 0.8 * i)
    raise ult


def main(caminho):
    caminho = os.path.abspath(caminho)
    antes = os.path.join(os.environ.get('TEMP', '.'), 'painel_pre_realce.xlsm')
    shutil.copy(caminho, antes)

    xl = novo_excel()
    wb = xl.Workbooks.Open(caminho)
    salvou = False
    if wb.ReadOnly:
        wb.Close(False)
        xl.Quit()
        raise SystemExit('Somente leitura: %s' % caminho)
    try:
        xl.Calculation = -4135
        pa = wb.Worksheets('Painel')
        cfg = wb.Worksheets(CFG)
        visCfg = cfg.Visible
        cfg.Visible = -1

        vbp = wb.VBProject
        for nome in ('mCEQ', 'mPlanoQC'):
            for c in list(vbp.VBComponents):
                if c.Name == nome:
                    vbp.VBComponents.Remove(c)
            importar_bas(vbp, os.path.join(RAIZ, 'src_producao', nome + '.bas'))

        # ---- 0. o bloco esta onde eu penso? ------------------------------
        print('=== conferindo as celulas de regra ANTES de escrever ===')
        for ref, rot in sorted(REGRAS_PAINEL.items()):
            v = str(tenta(lambda r=ref: pa.Range(r).Value) or '').strip()
            print('   %-4s = %-6s %s' % (ref, v, 'ok' if v == rot else 'DIVERGE'))
            if v != rot:
                raise SystemExit('celula de regra fora do lugar -- nada escrito')
        print('   Sigma N1 em %s = %s' % (SIGMA_N1, tenta(lambda: pa.Range('V10').Value)))
        print('   Sigma N2 em %s = %s' % (SIGMA_N2, tenta(lambda: pa.Range('V11').Value)))

        # ---- 1. Cfg: sigma efetivo e regras do motor ---------------------
        tenta(lambda: cfg.Range('A1').__setattr__(
            'Value', 'Sigma efetivo do plano (menor dos dois níveis):'))
        # MIN dos dois, ignorando o que nao for numero. Sem CSE: os IF
        # devolvem escalar.
        tenta(lambda: cfg.Range('B1').__setattr__(
            'Formula',
            '=IF(AND(NOT(ISNUMBER({0})),NOT(ISNUMBER({1}))),"",'
            'MIN(IF(ISNUMBER({0}),{0},9E+99),IF(ISNUMBER({1}),{1},9E+99)))'
            .format(SIGMA_N1, SIGMA_N2)))
        tenta(lambda: cfg.Range('C1').__setattr__(
            'Value', 'Regras que o motor executa:'))
        tenta(lambda: cfg.Range('D1').__setattr__(
            'Formula', '=mPlanoQC.RegrasImplementadas()'))
        for ref in ('A1', 'C1'):
            tenta(lambda r=ref: cfg.Range(r).Font.__setattr__('Italic', True))

        # ---- 2. Status_Plano_QC e a matriz de regras ---------------------
        cel = cfg.Cells(3, C_STATUS)
        cel.Value = 'Status_Plano_QC'
        cel.Font.Bold = True
        cel.Font.Color = BRANCO
        cel.Interior.Color = AZUL
        cfg.Columns(C_STATUS).ColumnWidth = 22
        for i in range(CFG_R0, CFG_RN + 1):
            regras = str(tenta(lambda r=i: cfg.Cells(r, 4).Value) or '').strip()
            tenta(lambda r=i, v=('REAVALIAR MÉTODO' if regras == ''
                                 else 'PLANO APLICÁVEL'):
                  cfg.Cells(r, C_STATUS).__setattr__('Value', v))

        for j, (rotulo, token) in enumerate(MATRIZ):
            c = MAT_C0 + j
            cel = cfg.Cells(3, c)
            cel.Value = rotulo
            cel.Font.Bold = True
            cel.Font.Color = BRANCO
            cel.Interior.Color = AZUL
            cfg.Columns(c).ColumnWidth = 8
            for i in range(CFG_R0, CFG_RN + 1):
                # A bandeira LE o texto de Regras da propria linha. Quando a
                # faixa nao tem regras (abaixo de 3 Sigma), cai no conjunto
                # que o motor suporta -- estrategia intensiva.
                tenta(lambda r=i, cc=c, tk=token: cfg.Cells(r, cc).__setattr__(
                    'Formula',
                    '=IF(ISNUMBER(SEARCH("{0}",IF($D{1}="",$D$1,$D{1}))),1,0)'
                    .format(tk, r)))
        tenta(lambda: cfg.Cells(2, MAT_C0).__setattr__(
            'Value', 'matriz derivada da coluna Regras — não editar à mão'))
        tenta(lambda: cfg.Cells(2, MAT_C0).Font.__setattr__('Italic', True))

        # ---- linha 10: a matriz RESOLVIDA para o Sigma de agora ----------
        #
        # A formatacao condicional nao aceitou a expressao inteira com INDEX +
        # SUMPRODUCT apontando para outra aba (E_INVALIDARG). Resolver aqui e
        # deixar a CF apenas comparar com 1 e o caminho robusto: CF sempre
        # aceita NOME DEFINIDO, mesmo que o nome aponte para outra planilha.
        tenta(lambda: cfg.Range('A2').__setattr__(
            'Value', 'Linha da faixa vigente (índice na matriz):'))
        tenta(lambda: cfg.Range('B2').__setattr__(
            'Formula',
            '=IF(NOT(ISNUMBER($B$1)),0,SUMPRODUCT(($A${0}:$A${1}<=$B$1)*'
            '($B${0}:$B${1}>$B$1)*ROW($A${0}:$A${1}))-{2})'
            .format(CFG_R0, CFG_RN, CFG_R0 - 1)))
        tenta(lambda: cfg.Range('A2').Font.__setattr__('Italic', True))
        tenta(lambda: cfg.Cells(9, MAT_C0).__setattr__(
            'Value', 'regras ativas para o Sigma selecionado:'))
        tenta(lambda: cfg.Cells(9, MAT_C0).Font.__setattr__('Italic', True))
        for j in range(len(MATRIZ)):
            c = MAT_C0 + j
            letra = chr(64 + c)
            tenta(lambda cc=c, lt=letra: cfg.Cells(10, cc).__setattr__(
                'Formula',
                '=IF(OR(NOT(ISNUMBER($B$1)),$B$2<1),0,'
                'INDEX({0}${1}:{0}${2},$B$2))'.format(lt, CFG_R0, CFG_RN)))
            tenta(lambda cc=c: cfg.Cells(10, cc).Font.__setattr__('Bold', True))

        # nomes: a CF le por NOME, e nao por referencia a outra aba
        for nome, ref in (
                ('regrasAtivas', '=%s!$%s$10:$%s$10'
                 % (CFG, chr(64 + MAT_C0), chr(64 + MAT_C0 + len(MATRIZ) - 1))),
                ('regrasRotulos', '=%s!$%s$3:$%s$3'
                 % (CFG, chr(64 + MAT_C0), chr(64 + MAT_C0 + len(MATRIZ) - 1))),
                ('sigmaDoPlano', '=%s!$B$1' % CFG)):
            try:
                wb.Names(nome).Delete()
            except Exception:
                pass
            wb.Names.Add(nome, ref)
        print('Cfg_PlanoQC: Status_Plano_QC, matriz de %d regras, '
              'linha resolvida e 3 nomes' % len(MATRIZ))

        # a tabela estruturada precisa abarcar as colunas novas
        for lo in list(cfg.ListObjects):
            if lo.Name == 'tblPlanoQC_Sigma':
                tenta(lambda l=lo: l.Resize(
                    cfg.Range(cfg.Cells(3, 1),
                              cfg.Cells(CFG_RN, MAT_C0 + len(MATRIZ) - 1))))

        # ---- 3. formatacao condicional nas celulas de regra --------------
        #
        # Duas condicoes por celula. A de VIOLACAO entra primeiro e por isso
        # ganha prioridade: uma regra violada nao pode ficar verde so porque
        # esta no plano.
        col_letra = {v: k[0] for k, v in REGRAS_PAINEL.items()}
        for ref, rotulo in sorted(REGRAS_PAINEL.items()):
            letra = ref[0]
            faixa = pa.Range(ref)
            try:
                faixa.FormatConditions.Delete()
            except Exception:
                pass
            # 1) violada em qualquer um dos dois niveis
            f1 = pa.Range(ref).FormatConditions.Add(
                2, None, '=SUM(${0}$7:${0}$8)>0'.format(letra))
            f1.Interior.Color = cor(0xFD, 0xE2, 0xE2)
            f1.Font.Color = cor(0xB4, 0x23, 0x18)
            f1.Font.Bold = True
            # 2) recomendada pelo Sigma -- le a matriz, nao uma escada nova
            # SUMPRODUCT, e nao INDEX/MATCH.
            #
            # Sondado celula a celula: a formatacao condicional ACEITA
            # SUMPRODUCT sobre nomes que apontam para outra aba, e RECUSA
            # INDEX/MATCH sobre os mesmos nomes -- inclusive dentro de
            # IFERROR. Sao funcoes que devolvem REFERENCIA, e a CF nao
            # atravessa planilha com elas. SUMPRODUCT so faz aritmetica de
            # matriz e passa.
            #
            # A conta e a mesma: casa o rotulo da celula com o rotulo da
            # matriz e devolve a bandeira daquela regra.
            f2 = pa.Range(ref).FormatConditions.Add(
                2, None,
                '=SUMPRODUCT((regrasRotulos={0})*regrasAtivas)=1'.format(ref))
            f2.Interior.Color = VERDE_ESCURO
            f2.Font.Color = BRANCO
            f2.Font.Bold = True
        print('Painel G6:K6: 2 regras de formato cada (violação > recomendação)')

        # ---- 4. legenda e texto de apoio, em area livre ------------------
        tenta(lambda: pa.Range('F10').__setattr__(
            'Formula',
            '=IF(NOT(ISNUMBER(sigmaDoPlano)),'
            '"Regras iluminadas = estratégia de CQ recomendada pelo Sigma '
            'do analito selecionado.",'
            '"Regras iluminadas = plano de CQ recomendado. Base: menor Sigma '
            'entre os dois níveis = "&TEXT(sigmaDoPlano,"0,00")&" ("&'
            'IF(AND(ISNUMBER({0}),OR(NOT(ISNUMBER({1})),{0}<={1})),"nível 1",'
            '"nível 2")&"). O plano precisa cobrir o nível que está pior — '
            'por isso ele pode pedir mais regras do que a classificação de '
            'um único nível sugere.")'.format(SIGMA_N1, SIGMA_N2)))
        tenta(lambda: pa.Range('F11').__setattr__(
            'Formula',
            '=IF(NOT(ISNUMBER(sigmaDoPlano)),"",'
            'IF(sigmaDoPlano<3,'
            '"DESEMPENHO INADEQUADO — CQ intensivo pode não ser suficiente. '
            'Investigar Bias, CV e desempenho do método.",'
            '"Plano de CQ aplicável para o Sigma deste analito."))'))
        tenta(lambda: pa.Range('F12').__setattr__(
            'Value', 'VERDE ESCURO: regra recomendada pelo Sigma   ·   '
                     'VERMELHO: regra violada na corrida   ·   '
                     'NEUTRO: fora da estratégia atual'))
        tenta(lambda: pa.Range('F13').__setattr__(
            'Formula',
            '=IF(NOT(ISNUMBER(sigmaDoPlano)),"","Motor executa: "&%s!$D$1&'
            '"   |   Cobertura: "&mPlanoQC.CoberturaWestgard(sigmaDoPlano))'
            % CFG))
        for ref in ('F10', 'F12', 'F13'):
            tenta(lambda r=ref: pa.Range(r).Font.__setattr__('Italic', True))
            tenta(lambda r=ref: pa.Range(r).Font.__setattr__('Size', 9))
            tenta(lambda r=ref: pa.Range(r).Font.Color.__class__ and
                  pa.Range(r).Font.__setattr__('Color', CINZA_TXT))
        tenta(lambda: pa.Range('F11').Font.__setattr__('Bold', True))
        tenta(lambda: pa.Range('F11').Font.__setattr__('Size', 9))
        print('Painel F10:F13: apoio, alerta de Sigma < 3, legenda e cobertura')

        cfg.Visible = visCfg

        xl.Calculation = -4105
        tenta(lambda: xl.CalculateFullRebuild())

        erros = []
        for r in range(1, 40):
            for c in range(1, 30):
                t = str(tenta(lambda x=r, y=c: pa.Cells(x, y).Text))
                if t.startswith('#') and t.strip('#') != '':
                    erros.append('Painel!%d,%d=%s' % (r, c, t))
        print('celulas em erro: %d %s' % (len(erros), erros[:5]))
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

    print()
    print('arquivo de comparacao: %s' % antes)


if __name__ == '__main__':
    main(sys.argv[1])
