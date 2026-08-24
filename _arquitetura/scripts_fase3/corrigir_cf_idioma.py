# -*- coding: utf-8 -*-
"""corrigir_cf_idioma.py - a formatacao condicional fala pt-BR

O DEFEITO

FormatConditions.Add(..., Formula1) nesta instalacao se comporta como
FormulaLocal: espera os nomes de funcao no idioma do Excel. Escrita em ingles,
a condicao e ACEITA sem erro e nunca avalia verdadeira -- ou e recusada com
E_INVALIDARG, dependendo da funcao.

Consequencia: o realce inteiro de G6:K6 estava morto. Nem a recomendacao
(SUMPRODUCT) nem a violacao (SUM) pintavam. As provas anteriores liam a
FORMULA da condicao e a PRIORIDADE dela e concluiam por raciocinio qual
venceria; nenhuma media o que o Excel realmente pintava.

COMO FOI ISOLADO

Uma condicao por vez na mesma celula, medindo Range.DisplayFormat depois de
cada uma:

    =TRUE                                    nao pintou
    =$G$6="1-3S"                             PINTOU
    =SUM($G$7:$G$8)=0                        nao pintou
    =sigmaDoPlano>=5                         PINTOU
    =SUMPRODUCT((regrasRotulos=G6)*...)=1    nao pintou
    =INDEX(regrasAtivas,1)=1                 RECUSADA

O que pintava nao tinha nome de funcao. Refeito em pt-BR, tudo pinta --
inclusive o INDICE que antes era recusado.

E O MESMO PADRAO do NumberFormat, que ja se comportara como NumberFormatLocal
e fez a Estatistica exibir 2,70 como "003".

Uso: python corrigir_cf_idioma.py <arquivo.xlsm>
"""
import io
import os
import sys
import time
import subprocess

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', write_through=True)
import win32com.client as w


def cor(r, g, b):
    return r + g * 256 + b * 65536

VERDE = cor(0x14, 0x6C, 0x43)
BRANCO = cor(255, 255, 255)
VERM_FUNDO = cor(0xFD, 0xE2, 0xE2)
VERM_FONTE = cor(0xB4, 0x23, 0x18)
VERDE_FONTE = cor(0x14, 0x6C, 0x43)

REGRAS = [('G6', '1-3S'), ('H6', '2-2S'), ('I6', 'R4S'),
          ('J6', '4-1S'), ('K6', '8X')]


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
        es = wb.Worksheets('Estatística')

        protegidas = []
        for ws in wb.Worksheets:
            if ws.ProtectContents:
                protegidas.append(ws.Name)
                for s in ('qcini2025', None):
                    try:
                        ws.Unprotect(s) if s else ws.Unprotect()
                        break
                    except Exception:
                        pass
        print('abas que estavam protegidas: %d' % len(protegidas))

        # ---- 1. G6:K6, agora em pt-BR ------------------------------------
        for ref, rot in REGRAS:
            v = str(tenta(lambda r=ref: pa.Range(r).Value) or '').strip()
            if v != rot:
                raise SystemExit('%s = %r, esperado %r -- nada escrito'
                                 % (ref, v, rot))
        for ref, rot in REGRAS:
            l = ref[0]
            try:
                pa.Range(ref).FormatConditions.Delete()
            except Exception:
                pass
            # SOMA, e nao SUM: nesta instalacao a CF le o idioma local
            f1 = pa.Range(ref).FormatConditions.Add(
                2, None, '=SOMA(${0}$7:${0}$8)>0'.format(l))
            f1.Interior.Color = VERM_FUNDO
            f1.Font.Color = VERM_FONTE
            f1.Font.Bold = True
            # SOMARPRODUTO, e nao SUMPRODUCT
            f2 = pa.Range(ref).FormatConditions.Add(
                2, None,
                '=SOMARPRODUTO((regrasRotulos={0})*regrasAtivas)=1'.format(ref))
            f2.Interior.Color = VERDE
            f2.Font.Color = BRANCO
            f2.Font.Bold = True
            f2.Font.Italic = True
        print('Painel G6:K6: condicoes reescritas em pt-BR')

        # ---- 2. o aviso da Estatistica tinha o mesmo defeito --------------
        #
        # Foi RECUSADO no ADR-034 (LEFT em ingles) e ficou registrado como
        # "sem cor, o texto continua". A causa era esta.
        alvo = es.Range('Z3')
        try:
            alvo.FormatConditions.Delete()
        except Exception:
            pass
        try:
            c1 = alvo.FormatConditions.Add(2, None, '=ESQUERDA($Z$3;4)="BASE"')
            c1.Font.Color = VERM_FONTE
            c1.Font.Bold = True
            c2 = alvo.FormatConditions.Add(2, None, '=ESQUERDA($Z$3;4)="base"')
            c2.Font.Color = VERDE_FONTE
            print('Estatistica Z3: aviso de base desatualizada ganhou cor')
        except Exception as e:
            print('Estatistica Z3: ainda recusado (%s)' % str(e)[:50])

        xl.Calculation = -4105
        tenta(lambda: xl.CalculateFullRebuild())

        # ---- 3. prova pelo formato EFETIVO -------------------------------
        print()
        print('=== prova: DisplayFormat depois da correcao ===')
        guardaV = (tenta(lambda: pa.Range('V10').Formula),
                   tenta(lambda: pa.Range('V11').Formula))
        guardaG = {}
        for lin in (7, 8):
            for c in range(7, 12):
                guardaG[(lin, c)] = tenta(
                    lambda x=lin, y=c: pa.Cells(x, y).Formula)

        for sg, esperadas in ((6.0, ['1-3S']),
                              (5.5, ['1-3S', '2-2S', 'R4S']),
                              (4.5, ['1-3S', '2-2S', 'R4S', '4-1S']),
                              (2.99, ['1-3S', '2-2S', 'R4S', '4-1S', '8X'])):
            for r in ('V10', 'V11'):
                tenta(lambda x=r, v=sg: pa.Range(x).__setattr__('Value', v))
            for lin in (7, 8):
                for c in range(7, 12):
                    tenta(lambda x=lin, y=c: pa.Cells(x, y).__setattr__('Value', 0))
            tenta(lambda: xl.CalculateFull())
            verdes = [rot for ref, rot in REGRAS
                      if int(tenta(lambda r=ref:
                                   pa.Range(r).DisplayFormat.Interior.Color)) == VERDE]
            ok = verdes == esperadas
            print('   %s Sigma %.2f -> verdes %s' %
                  ('OK  ' if ok else 'FALHA', sg, verdes))
            if not ok:
                raise SystemExit('o realce ainda nao pinta -- nada salvo')

        # violacao vence
        tenta(lambda: pa.Range('V10').__setattr__('Value', 5.5))
        tenta(lambda: pa.Range('V11').__setattr__('Value', 5.5))
        tenta(lambda: pa.Range('G7').__setattr__('Value', 3))
        tenta(lambda: xl.CalculateFull())
        f = int(tenta(lambda: pa.Range('G6').DisplayFormat.Interior.Color))
        it = bool(tenta(lambda: pa.Range('G6').DisplayFormat.Font.Italic))
        print('   %s violacao vence: fundo #%06X (violacao=#%06X)'
              % ('OK  ' if f == VERM_FUNDO else 'FALHA', f, VERM_FUNDO))
        if f != VERM_FUNDO:
            raise SystemExit('a violacao nao venceu -- nada salvo')

        # devolve o que foi mexido
        tenta(lambda: pa.Range('V10').__setattr__('Formula', guardaV[0]))
        tenta(lambda: pa.Range('V11').__setattr__('Formula', guardaV[1]))
        for (lin, c), fx in guardaG.items():
            tenta(lambda x=lin, y=c, ff=fx:
                  pa.Cells(x, y).__setattr__('Formula', ff))
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


if __name__ == '__main__':
    main(sys.argv[1])
