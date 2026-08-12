# -*- coding: utf-8 -*-
"""instalar_capacidade60m.py - ADR-025: capacidade de 60 meses e fim do custo n^2

O QUE FAZ, em ordem

  1. captura os valores atuais de BB e BC (produzidos pelas formulas) para
     servir de GABARITO da prova de equivalencia depois
  2. importa o modulo mBanco
  3. liga AtualizarFlagsBanco ao fim de UpsertResultados e de ExcluirLogico,
     e poe a barreira ExigirCapacidade antes da gravacao
  4. apaga as formulas de BA:BC e chama AtualizarFlagsBanco para gravar os
     VALORES equivalentes em todo o historico existente
  5. redimensiona os intervalos nomeados r* para a ultima linha real
  6. compara valor novo x gabarito antigo, linha a linha
  7. so salva se a equivalencia for de 100%

Nada e salvo se qualquer conferencia falhar. Trabalha no arquivo passado por
argumento -- use sempre uma copia primeiro.

Uso: python instalar_capacidade60m.py <caminho.xlsm>
"""
import io
import os
import sys
import time
import subprocess

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', write_through=True)
import win32com.client as w

SENHA = 'qcini2025'
AQUI = os.path.dirname(os.path.abspath(__file__))
MODULO = os.path.join(os.path.dirname(AQUI), 'src_producao', 'mBanco.bas')

# --- trechos injetados -----------------------------------------------------
GUARDA_UPSERT = """        ' Barreira de capacidade (ADR-025). Antes de gravar, nao depois: aceitar o
        ' registro e descobrir o estouro em seguida deixaria o banco pela metade.
        ExigirCapacidade nAdd
"""

CHAMADA_FLAGS = """    ' As flags BA/BB/BC sao mantidas por mBanco desde o ADR-025. Recalcular aqui,
    ' e nao so nas linhas novas: uma linha inserida ou reativada muda quem e a
    ' "primeira ativa" das linhas seguintes.
    AtualizarFlagsBanco
"""


def novo_excel():
    for t in range(1, 8):
        try:
            xl = w.DispatchEx('Excel.Application')
            xl.Visible = False
            xl.DisplayAlerts = False
            xl.EnableEvents = False
            xl.AutomationSecurity = 1
            return xl
        except Exception:
            if t == 2:
                subprocess.call(['powershell', '-NoProfile', '-Command',
                                 'Start-Process excel.exe -WindowStyle Hidden'],
                                stderr=subprocess.DEVNULL)
                time.sleep(6)
            time.sleep(2.0 * t)
    raise RuntimeError('Excel COM nao subiu')


def modulo(wb, nome):
    for c in wb.VBProject.VBComponents:
        if c.Name == nome:
            return c
    return None


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
        if wb.ProtectStructure:
            wb.Unprotect(SENHA)
        db = wb.Worksheets('DB_Resultados')
        for t in (lambda: db.Unprotect(SENHA), lambda: db.Unprotect()):
            try:
                t()
                break
            except Exception:
                pass

        ult = db.Cells(db.Rows.Count, 1).End(-4162).Row
        print('banco: ultima linha com dado = %d  (%d registros)' % (ult, max(0, ult - 3)))
        if ult < 4:
            raise SystemExit('banco vazio -- nada a converter')

        # ---- 1. gabarito: o que as FORMULAS produzem hoje -------------------
        xl.Calculation = -4105
        xl.CalculateFullRebuild()
        gab = db.Range(db.Cells(4, 53), db.Cells(ult, 55)).Value
        gab = [[('' if v is None else v) for v in linha] for linha in gab]
        print('gabarito capturado das formulas: %d linhas x 3 colunas' % len(gab))

        # ---- 2. modulo mBanco ----------------------------------------------
        vbp = wb.VBProject
        antigo = modulo(wb, 'mBanco')
        if antigo is not None:
            vbp.VBComponents.Remove(antigo)
        vbp.VBComponents.Import(MODULO)
        print('modulo mBanco importado')

        # ---- 3. ligar aos pontos de escrita --------------------------------
        md = modulo(wb, 'mDados')
        if md is None:
            raise SystemExit('mDados nao encontrado')
        cm = md.CodeModule
        txt = cm.Lines(1, cm.CountOfLines)

        if 'AtualizarFlagsBanco' in txt:
            print('mDados ja estava ligado ao mBanco -- nada a inserir')
        else:
            # barreira: imediatamente antes do bloco que acrescenta os novos
            alvo = "    ' --- acrescenta os novos em um unico bloco ---"
            linhas = txt.split('\r\n')
            iAlvo = None
            for i, L in enumerate(linhas):
                if L.strip().startswith("' --- acrescenta os novos"):
                    iAlvo = i + 1
                    break
            if iAlvo is None:
                raise SystemExit('ancora do bloco de insercao nao encontrada em UpsertResultados')
            cm.InsertLines(iAlvo, GUARDA_UPSERT.rstrip('\n'))
            print('ExigirCapacidade inserida em UpsertResultados (linha %d)' % iAlvo)

            # chamada das flags: antes do rotulo Limpeza do Upsert
            txt = cm.Lines(1, cm.CountOfLines)
            linhas = txt.split('\r\n')
            iLimp = None
            for i, L in enumerate(linhas):
                if L.strip() == 'Limpeza:':
                    iLimp = i + 1
                    break
            if iLimp is None:
                raise SystemExit('rotulo Limpeza nao encontrado em UpsertResultados')
            cm.InsertLines(iLimp, CHAMADA_FLAGS.rstrip('\n'))
            print('AtualizarFlagsBanco inserida em UpsertResultados (linha %d)' % iLimp)

            # e no fim de ExcluirLogico, antes do return
            txt = cm.Lines(1, cm.CountOfLines)
            linhas = txt.split('\r\n')
            iExc = None
            for i, L in enumerate(linhas):
                if L.strip() == 'ExcluirLogico = n':
                    iExc = i + 1
                    break
            if iExc is None:
                raise SystemExit('fim de ExcluirLogico nao encontrado')
            cm.InsertLines(iExc, CHAMADA_FLAGS.rstrip('\n'))
            print('AtualizarFlagsBanco inserida em ExcluirLogico (linha %d)' % iExc)

        # ---- 4. formulas fora, valores dentro ------------------------------
        alvoF = db.Range(db.Cells(4, 53), db.Cells(db.Rows.Count, 55))
        alvoF.ClearContents()
        print('formulas BA:BC removidas de todo o provisionamento')

        xl.Run('AtualizarFlagsBanco')
        print('AtualizarFlagsBanco executada: BA/BB/BC gravados como VALOR')

        # ---- 5. conferir os nomes ------------------------------------------
        for n in ('rRUN', 'rData', 'rNivel', 'rAnalito', 'rValor', 'rStatus',
                  'rLote', 'rFirst', 'rRunUnico'):
            ref = wb.Names(n).RefersTo
            if ('$%d' % ult) not in ref:
                raise SystemExit('nome %s nao acompanhou a ultima linha: %s' % (n, ref))
        print('9 intervalos nomeados redimensionados para a linha %d' % ult)

        # ---- 6. prova de equivalencia --------------------------------------
        novo = db.Range(db.Cells(4, 53), db.Cells(ult, 55)).Value
        novo = [[('' if v is None else v) for v in linha] for linha in novo]
        div = []
        for i in range(len(gab)):
            for c in range(3):
                a, b = gab[i][c], novo[i][c]
                # "" da formula e celula vazia do VBA sao o mesmo estado logico
                if (a == '' or a is None) and (b == '' or b is None):
                    continue
                if str(a).strip() != str(b).strip():
                    div.append((4 + i, 'BA BB BC'.split()[c], a, b))
        print('\n=== PROVA DE EQUIVALENCIA ===')
        print('registros comparados : %d  (x3 colunas = %d celulas)'
              % (len(gab), len(gab) * 3))
        print('divergencias         : %d' % len(div))
        print('concordancia         : %.4f%%'
              % (100.0 * (len(gab) * 3 - len(div)) / max(1, len(gab) * 3)))
        for d in div[:10]:
            print('   L%d %s: formula=%r  vba=%r' % d)
        if div:
            raise SystemExit('EQUIVALENCIA < 100%% -- nada foi salvo')

        # conferencia independente, pelo caminho ingenuo
        r = xl.Run('ConferirFlagsBanco', 3000)
        print('ConferirFlagsBanco (caminho independente): %s' % r)
        if r.split('|')[1] != '0':
            raise SystemExit('ConferirFlagsBanco acusou divergencia -- nada foi salvo')

        xl.CalculateFullRebuild()
        if wb.ProtectStructure is False:
            wb.Protect(SENHA, True, False)
        wb.Save()
        salvou = True
        print('\nSALVO: %s' % caminho)
        print('tamanho: %.2f MB' % (os.path.getsize(caminho) / 1048576.0))
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
