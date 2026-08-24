# -*- coding: utf-8 -*-
"""padronizar_8x_pior_nivel.py - ADR-038: 8x definitivo, pior nivel governa

TRES DECISOES FECHADAS PELO GESTOR

  1. a familia de regra sequencial e 8x, e so 8x. 6x e 10x saem do projeto.
  2. o plano de CQ e governado pelo PIOR nivel: Sigma_Plano = MIN(N1, N2),
     entre os niveis validos. Sem nenhum valido, SEM DADOS.
  3. area visual sem funcao pode ser removida ou reconectada.

O DEFEITO QUE A DECISAO 2 CORRIGE

Lactato tem Sigma 6,99 no nivel 1 e 1,83 no nivel 2. O bloco de plano exibia a
linha do nivel 1 -- "1_3s, N=2, run size 1000" --, o CQ mais leve que existe,
enquanto o nivel 2 esta em 1,8 Sigma. Um plano lido assim autoriza mil
pacientes entre eventos de controle num metodo que nao sustenta nem 3 Sigma.

O bloco passa a mostrar UM plano, o do pior nivel, com a base declarada na
linha de baixo.

O QUE NAO MUDA

Posicao, cor, tamanho, borda, alinhamento, largura, altura, M7/M8, filtros. O
layout e do gestor.

Uso: python padronizar_8x_pior_nivel.py <arquivo.xlsm>
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

CFG = 'Cfg_PlanoQC'
CFG_R0, CFG_RN = 4, 8
MAT_C0 = 10                    # J
C_STATUS = 9                   # I

# A familia definitiva. Cinco regras, e so elas.
MATRIZ = [('1-3S', '1_3s'), ('2-2S', '2_2s'), ('R4S', 'R_4s'),
          ('4-1S', '4_1s'), ('8X', '8x')]
MAT_FIM = MAT_C0 + len(MATRIZ) - 1                     # N
LIXO_C0, LIXO_C1 = MAT_FIM + 1, MAT_C0 + 6             # O..P: 6x e 10x saem

REGRAS_PAINEL = {'G6': '1-3S', 'H6': '2-2S', 'I6': 'R4S',
                 'J6': '4-1S', 'K6': '8X'}

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


def letra(n):
    s = ''
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def main(caminho):
    caminho = os.path.abspath(caminho)
    antes = os.path.join(os.environ.get('TEMP', '.'), 'pre_adr038.xlsm')
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

        # A protecao volta sozinha entre etapas -- ja aconteceu com a
        # EQC_Dados e com a Estatistica. O estado e MEDIDO e registrado, e nao
        # suposto: sem isso, Range.Clear devolve 0x800A03EC e o script morre no
        # meio da alteracao.
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
        resta = [x.Name for x in wb.Worksheets if x.ProtectContents]
        if resta:
            raise SystemExit('nao consegui desproteger: %s -- nada escrito'
                             % resta)

        vbp = wb.VBProject
        for nome in ('mCEQ', 'mPlanoQC'):
            for c in list(vbp.VBComponents):
                if c.Name == nome:
                    vbp.VBComponents.Remove(c)
            importar_bas(vbp, os.path.join(RAIZ, 'src_producao', nome + '.bas'))

        # ---- 0. os pressupostos ainda valem? -----------------------------
        print('=== conferindo antes de escrever ===')
        for ref, rot in sorted(REGRAS_PAINEL.items()):
            v = str(tenta(lambda r=ref: pa.Range(r).Value) or '').strip()
            if v != rot:
                raise SystemExit('%s = %r, esperado %r -- nada escrito'
                                 % (ref, v, rot))
        print('   G6:K6 conferidas: %s' % [REGRAS_PAINEL[k]
                                           for k in sorted(REGRAS_PAINEL)])
        for ref, esperado in (('R20', 'PLANO DE CQ'), ('R21', 'Nível'),
                              ('S21', 'Sigma'), ('R15', 'ERRO TOTAL')):
            v = str(tenta(lambda r=ref: pa.Range(r).Value) or '')
            if esperado.lower() not in v.lower():
                raise SystemExit('%s = %r nao contem %r -- nada escrito'
                                 % (ref, v[:40], esperado))
        print('   blocos R15 e R20 conferidos')

        # ---- 1. 6x e 10x saem da matriz ----------------------------------
        antesRot = [str(tenta(lambda c=c: cfg.Cells(3, c).Value) or '')
                    for c in range(MAT_C0, MAT_C0 + 7)]
        print()
        print('=== 6x E 10x SAEM ===')
        print('   matriz antes : %s' % antesRot)
        # encolhe a tabela PRIMEIRO: limpar celula que ainda pertence ao
        # ListObject e briga desnecessaria com o Excel
        for lo in list(cfg.ListObjects):
            if lo.Name == 'tblPlanoQC_Sigma':
                tenta(lambda l=lo: l.Resize(
                    cfg.Range(cfg.Cells(3, 1), cfg.Cells(CFG_RN, MAT_FIM))))
        tenta(lambda: cfg.Range('%s1:%s%d' % (letra(LIXO_C0), letra(LIXO_C1),
                                              CFG_RN + 4)).Clear())
        for j, (rotulo, token) in enumerate(MATRIZ):
            c = MAT_C0 + j
            cel = cfg.Cells(3, c)
            cel.Value = rotulo
            cel.Font.Bold = True
            cel.Font.Color = BRANCO
            cel.Interior.Color = AZUL
            for i in range(CFG_R0, CFG_RN + 1):
                tenta(lambda r=i, cc=c, tk=token: cfg.Cells(r, cc).__setattr__(
                    'Formula',
                    '=IF(ISNUMBER(SEARCH("{0}",IF($D{1}="",$D$1,$D{1}))),1,0)'
                    .format(tk, r)))
            tenta(lambda cc=c, lt=letra(c): cfg.Cells(10, cc).__setattr__(
                'Formula',
                '=IF(OR(NOT(ISNUMBER($B$1)),$B$2<1),0,'
                'INDEX({0}${1}:{0}${2},$B$2))'.format(lt, CFG_R0, CFG_RN)))
        depoisRot = [str(tenta(lambda c=c: cfg.Cells(3, c).Value) or '')
                     for c in range(MAT_C0, MAT_C0 + 7)]
        print('   matriz depois: %s' % depoisRot)

        for lo in list(cfg.ListObjects):
            if lo.Name == 'tblPlanoQC_Sigma':
                tenta(lambda l=lo: l.Resize(
                    cfg.Range(cfg.Cells(3, 1), cfg.Cells(CFG_RN, MAT_FIM))))

        for nome, ref in (('regrasAtivas', '=%s!$%s$10:$%s$10'
                           % (CFG, letra(MAT_C0), letra(MAT_FIM))),
                          ('regrasRotulos', '=%s!$%s$3:$%s$3'
                           % (CFG, letra(MAT_C0), letra(MAT_FIM)))):
            try:
                wb.Names(nome).Delete()
            except Exception:
                pass
            wb.Names.Add(nome, ref)
        print('   nomes regrasAtivas/regrasRotulos agora cobrem %s..%s'
              % (letra(MAT_C0), letra(MAT_FIM)))

        # ---- 2. Sigma_Plano diz SEM DADOS quando nao ha nivel valido -----
        tenta(lambda: cfg.Range('B1').__setattr__(
            'Formula',
            '=IF(AND(NOT(ISNUMBER({0})),NOT(ISNUMBER({1}))),"SEM DADOS",'
            'MIN(IF(ISNUMBER({0}),{0},9E+99),IF(ISNUMBER({1}),{1},9E+99)))'
            .format(SIGMA_N1, SIGMA_N2)))
        # A3 e B3 sao os CABECALHOS Sigma_Min e Sigma_Max da tabela --
        # escrever ali destroi a tabela e faz o Excel devolver 0 no lugar do
        # texto. Foi o que aconteceu na primeira tentativa: o Painel exibia
        # "Governa o PIOR nível: 0". A celula do nivel que governa vai para
        # C2/D2, acima da tabela.
        tenta(lambda: cfg.Range('A3').__setattr__('Value', 'Sigma_Min'))
        tenta(lambda: cfg.Range('B3').__setattr__('Value', 'Sigma_Max'))
        tenta(lambda: cfg.Range('C2').__setattr__(
            'Value', 'Nível que governa o plano:'))
        tenta(lambda: cfg.Range('C2').Font.__setattr__('Italic', True))
        tenta(lambda: cfg.Range('D2').__setattr__(
            'Formula',
            '=IF(NOT(ISNUMBER($B$1)),"",IF(AND(ISNUMBER({0}),'
            'OR(NOT(ISNUMBER({1})),{0}<={1})),"nível 1","nível 2"))'
            .format(SIGMA_N1, SIGMA_N2)))
        print('   Sigma_Plano = MIN dos níveis válidos; sem nenhum, SEM DADOS')

        # ---- 2b. abaixo de 3 Sigma o TEXTO acompanha o realce ------------
        #
        # O realce acendia as cinco regras (estrategia intensiva) enquanto a
        # coluna Regras ficava vazia -- texto e cor discordando na mesma tela.
        # Agora a faixa declara a estrategia. N e run size continuam VAZIOS:
        # nao existe run size tabelado abaixo de 3 Sigma, e inventar um seria
        # autorizar intervalo de CQ que o metodo nao sustenta.
        intensivo = ' / '.join(t for _r, t in MATRIZ)
        for i in range(CFG_R0, CFG_RN + 1):
            smax = tenta(lambda r=i: cfg.Cells(r, 2).Value)
            if isinstance(smax, (int, float)) and smax <= 3:
                tenta(lambda r=i, v=intensivo:
                      cfg.Cells(r, 4).__setattr__('Value', v))
                print('   faixa < 3 Sigma: Regras = %s (N e run size seguem '
                      'vazios)' % intensivo)

        # ---- 3. o bloco do plano passa a mostrar UM plano, o do pior -----
        #
        # Antes eram duas linhas por nivel, e o Painel exibia a do nivel 1.
        # Com Lactato (N1 6,99 / N2 1,83) isso recomendava run size 1000 para
        # um metodo que nao sustenta 3 Sigma.
        tenta(lambda: pa.Range('R22').__setattr__('Value', 'Plano'))
        tenta(lambda: pa.Range('S22').__setattr__(
            'Formula', '=IF(NOT(ISNUMBER(%s!$B$1)),"",%s!$B$1)' % (CFG, CFG)))
        for c, campo in ((20, 'REGRAS'), (21, 'N'), (22, 'RUNSIZE'),
                         (23, 'FREQUENCIA')):
            tenta(lambda cc=c, k=campo: pa.Cells(22, cc).__setattr__(
                'Formula',
                '=IF(NOT(ISNUMBER({0}!$B$1)),"",mPlanoQC.PlanoQC({0}!$B$1,"{1}"))'
                .format(CFG, k)))
        tenta(lambda: pa.Cells(22, 24).__setattr__(
            'Formula',
            '=IF(NOT(ISNUMBER({0}!$B$1)),"",'
            'mPlanoQC.CoberturaWestgard({0}!$B$1))'.format(CFG)))

        # a linha de baixo deixa de repetir e passa a declarar a base
        tenta(lambda: pa.Range('R23').__setattr__('Value', 'Base'))
        tenta(lambda: pa.Range('S23').__setattr__(
            'Formula',
            '=IF(NOT(ISNUMBER({0}!$B$1)),"Sem Sigma válido em nenhum nível — '
            'sem plano.","Pior nível governa: "&{0}!$D$2&" (Sigma "&'
            'TEXT({0}!$B$1,"0,00")&").   N1 = "&'
            'IF(ISNUMBER({1}),TEXT({1},"0,00"),"—")&"   ·   N2 = "&'
            'IF(ISNUMBER({2}),TEXT({2},"0,00"),"—")&'
            'IF({0}!$B$1<3,"   ·   "&INDEX({0}!$I${3}:$I${4},{0}!$B$2),""))'
            .format(CFG, SIGMA_N1, SIGMA_N2, CFG_R0, CFG_RN)))
        for c in range(20, 25):
            tenta(lambda cc=c: pa.Cells(23, cc).ClearContents())
        tenta(lambda: pa.Range('S23').Font.__setattr__('Italic', True))
        tenta(lambda: pa.Range('S23').Font.__setattr__('Size', 9))
        print('   R22 = plano do pior nível ; R23 = base declarada')

        # ---- 4. realce: negrito E italico --------------------------------
        for ref in sorted(REGRAS_PAINEL):
            l = ref[0]
            try:
                pa.Range(ref).FormatConditions.Delete()
            except Exception:
                pass
            f1 = pa.Range(ref).FormatConditions.Add(
                2, None, '=SOMA(${0}$7:${0}$8)>0'.format(l))
            f1.Interior.Color = cor(0xFD, 0xE2, 0xE2)
            f1.Font.Color = cor(0xB4, 0x23, 0x18)
            f1.Font.Bold = True
            f2 = pa.Range(ref).FormatConditions.Add(
                2, None,
                '=SOMARPRODUTO((regrasRotulos={0})*regrasAtivas)=1'.format(ref))
            f2.Interior.Color = VERDE_ESCURO
            f2.Font.Color = BRANCO
            f2.Font.Bold = True
            f2.Font.Italic = True
        print('   G6:K6: recomendada = verde escuro, branca, NEGRITO + ITÁLICO')

        # ---- 5. texto de apoio acompanha a decisao ------------------------
        tenta(lambda: pa.Range('F10').__setattr__(
            'Formula',
            '=IF(NOT(ISNUMBER({0}!$B$1)),'
            '"Regras iluminadas = plano de CQ recomendado pelo Sigma.",'
            '"Regras iluminadas = plano de CQ recomendado. Governa o PIOR '
            'nível: "&{0}!$D$2&" (Sigma "&TEXT({0}!$B$1,"0,00")&").")'
            .format(CFG)))
        cfg.Visible = visCfg

        xl.Calculation = -4105
        tenta(lambda: xl.CalculateFullRebuild())

        erros = []
        for r in range(1, 40):
            for c in range(1, 30):
                t = str(tenta(lambda x=r, y=c: pa.Cells(x, y).Text))
                if t.startswith('#') and t.strip('#') != '':
                    erros.append('Painel!%s%d=%s' % (letra(c), r, t))
        print()
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
    print('arquivo de comparacao: %s' % antes)


if __name__ == '__main__':
    main(sys.argv[1])
