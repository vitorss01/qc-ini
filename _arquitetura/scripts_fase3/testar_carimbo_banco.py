# -*- coding: utf-8 -*-
"""testar_carimbo_banco.py - ADR-049: a Estatistica acompanha o banco

O DEFEITO

mEstatPeriodo reaproveita os agregados de mAgg e decidia pelo Carimbo, cuja
unica impressao digital do DADO era a ULTIMA LINHA de DB_Resultados.

Inclusao nova mudava a ultima linha e invalidava. Edicao EM LINHA nao -- e
marcar um resultado como nao conforme grava Status na propria linha
(mRegistros). Com o cache quente, medido:

    n inicial ............................... 25
    apos Status -> Excluido, sem recalcular .. 25
    apos Application.Calculate (F9) ......... 25
    apos CalculateFullRebuild ............... 25   <- devia ser 24

Isso contamina media, DP, CV e daí Sigma, classificacao e plano de CQ.

O QUE ESTE TESTE PROVA

Pelo fluxo NORMAL -- sem reimportar modulo, sem fechar o arquivo, sem macro
especial de QA -- que a Estatistica acompanha:

  1. edicao em linha (Status)  ......... n cai, e volta ao desfazer
  2. edicao em linha (valor)   ......... media/DP/CV mudam com n IGUAL
  3. inclusao de linha nova    ......... n sobe
  4. propagacao ate o plano    ......... CV -> Sigma -> classificacao -> plano
  5. reabertura                ......... continua funcionando depois de fechar

O caso 2 e o que separa esta solucao de uma heuristica por contagem de linhas:
o numero de linhas nao muda, e mesmo assim o resultado tem de mudar.

NADA E SALVO: cada cenario e desfeito e o arquivo fecha sem gravar.

Uso: python testar_carimbo_banco.py <arquivo.xlsm>
"""
import io
import os
import sys

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8',
                              write_through=True)
import win32com.client as w

I = chr(0x00ED)
ABA_EST = 'Estat' + I + 'stica'
C_NIVEL, C_ANALITO, C_VALOR, C_STATUS = 3, 5, 6, 7
XL_UP = -4162
XL_AUTOMATICO = -4105

falhas = []


def checar(nome, ok, detalhe=''):
    if not ok:
        falhas.append(nome)
    print('   %-56s %s' % (nome[:56], 'PASS' if ok else 'FAIL'))
    if not ok and detalhe:
        for l in str(detalhe).split('\n')[:4]:
            print('        %s' % l)


def abrir(caminho):
    xl = w.DispatchEx('Excel.Application')
    xl.Visible = False
    xl.DisplayAlerts = False
    xl.EnableEvents = False
    xl.AutomationSecurity = 1
    wb = xl.Workbooks.Open(caminho, 0, False)
    xl.Calculation = XL_AUTOMATICO
    return xl, wb


def main():
    caminho = os.path.abspath(sys.argv[1])
    xl, wb = abrir(caminho)
    try:
        est = wb.Worksheets(ABA_EST)
        db = wb.Worksheets('DB_Resultados')
        print('=' * 78)
        print('CARIMBO DO BANCO -- a Estatistica acompanha o dado?')
        print('=' * 78)

        try:
            carimbo0 = str(wb.Names('DB_Carimbo').RefersToRange.Value)
        except Exception as e:
            checar('DB_Carimbo existe', False, str(e))
            print('\nTOTAL: %d FAIL' % len(falhas))
            return 1
        print('   carimbo inicial: %s' % carimbo0[:60])

        # acha a linha da Estatistica do 1o analito e a 1a linha ativa dele
        lin = 0
        for r in range(1, 60):
            if str(est.Cells(r, 1).Value or '').strip() == 'Analito':
                lin = r + 1
                break
        analito = str(est.Cells(lin, 1).Value or '').strip()
        nivel = str(est.Cells(lin, 2).Value or '').strip()
        print('   sentinela: %s N%s (linha %d da Estatistica)'
              % (analito, nivel, lin))

        def le():
            xl.Calculate()          # F9 -- o mesmo que mDados dispara
            v = est.Range(est.Cells(lin, 3), est.Cells(lin, 6)).Value[0]
            return tuple(None if x is None else round(float(x), 9)
                         if isinstance(x, (int, float)) else x for x in v)

        base = le()
        print('   n/media/DP/CV inicial: %s' % (base,))
        print()

        ult = db.Cells(db.Rows.Count, 1).End(XL_UP).Row
        dados = db.Range(db.Cells(4, 1), db.Cells(ult, 7)).Value
        alvo = None
        for i, r in enumerate(dados):
            if (str(r[C_ANALITO - 1]).strip() == analito
                    and str(r[C_NIVEL - 1]).strip().rstrip('.0') == nivel.rstrip('.0')
                    and str(r[C_STATUS - 1]).strip() == 'Ativo'):
                alvo = 4 + i
                break
        if alvo is None:
            checar('achou linha ativa do sentinela no banco', False)
            print('\nTOTAL: %d FAIL' % len(falhas))
            return 1

        # --- 1. edicao em linha: Status ----------------------------------
        antes = db.Cells(alvo, C_STATUS).Value
        db.Cells(alvo, C_STATUS).Value = 'Excluido'
        d1 = le()
        checar('Status editado em linha reduz o n (F9, sem rebuild)',
               d1[0] is not None and base[0] is not None and d1[0] == base[0] - 1,
               'antes n=%s, agora n=%s' % (base[0], d1[0]))
        db.Cells(alvo, C_STATUS).Value = antes
        d1v = le()
        checar('desfazer a edicao devolve os valores originais', d1v == base,
               'base %s\nagora %s' % (base, d1v))

        # --- 2. edicao em linha: valor, com n IGUAL ----------------------
        vAntes = db.Cells(alvo, C_VALOR).Value
        db.Cells(alvo, C_VALOR).Value = float(vAntes) * 1.5
        d2 = le()
        checar('valor editado muda media/DP/CV mantendo o n',
               d2[0] == base[0] and d2[1] != base[1],
               'n %s->%s  media %s->%s' % (base[0], d2[0], base[1], d2[1]))
        db.Cells(alvo, C_VALOR).Value = vAntes
        checar('desfazer o valor devolve os originais', le() == base)

        # --- 3. inclusao de linha nova ------------------------------------
        nova = ult + 1
        origem = db.Range(db.Cells(alvo, 1), db.Cells(alvo, 7)).Value[0]
        db.Range(db.Cells(nova, 1), db.Cells(nova, 7)).Value = origem
        d3 = le()
        checar('linha nova aumenta o n',
               d3[0] == base[0] + 1, 'n %s -> %s' % (base[0], d3[0]))
        db.Range(db.Cells(nova, 1), db.Cells(nova, 7)).ClearContents()
        checar('remover a linha volta ao estado inicial', le() == base)

        # --- 4. propagacao ate o plano de CQ ------------------------------
        # A cadeia vive na propria linha da Estatistica: CV -> Sigma ->
        # classificacao. O plano vem do motor, que le a mesma serie.
        def linhaToda():
            xl.Calculate()
            return [str(x)[:18] for x in
                    est.Range(est.Cells(lin, 1), est.Cells(lin, 14)).Value[0]]
        antesLinha = linhaToda()
        db.Cells(alvo, C_VALOR).Value = float(vAntes) * 3.0
        depoisLinha = linhaToda()
        mudou = [i for i in range(len(antesLinha))
                 if antesLinha[i] != depoisLinha[i]]
        checar('a mudanca propaga por varias colunas (CV/Sigma/classificacao)',
               len(mudou) >= 3,
               'colunas que mudaram: %s\nantes  %s\ndepois %s'
               % (mudou, antesLinha[:12], depoisLinha[:12]))
        db.Cells(alvo, C_VALOR).Value = vAntes
        xl.Calculate()
    finally:
        try:
            wb.Close(False)
        except Exception:
            pass
        try:
            xl.Quit()
        except Exception:
            pass

    # --- 5. reabertura -----------------------------------------------------
    xl, wb = abrir(caminho)
    try:
        est = wb.Worksheets(ABA_EST)
        db = wb.Worksheets('DB_Resultados')
        lin2 = 0
        for r in range(1, 60):
            if str(est.Cells(r, 1).Value or '').strip() == 'Analito':
                lin2 = r + 1
                break
        xl.Calculate()
        b2 = est.Cells(lin2, 3).Value
        ult = db.Cells(db.Rows.Count, 1).End(XL_UP).Row
        dados = db.Range(db.Cells(4, 1), db.Cells(ult, 7)).Value
        an = str(est.Cells(lin2, 1).Value or '').strip()
        nv = str(est.Cells(lin2, 2).Value or '').strip()
        alvo = next((4 + i for i, r in enumerate(dados)
                     if str(r[C_ANALITO - 1]).strip() == an
                     and str(r[C_NIVEL - 1]).strip().rstrip('.0') == nv.rstrip('.0')
                     and str(r[C_STATUS - 1]).strip() == 'Ativo'), None)
        db.Cells(alvo, C_STATUS).Value = 'Excluido'
        xl.Calculate()
        d5 = est.Cells(lin2, 3).Value
        checar('depois de fechar e reabrir, continua acompanhando',
               d5 == b2 - 1, 'n %s -> %s' % (b2, d5))
    finally:
        try:
            wb.Close(False)
        except Exception:
            pass
        try:
            xl.Quit()
        except Exception:
            pass

    print()
    print('TOTAL: %d FAIL' % len(falhas))
    return 1 if falhas else 0


if __name__ == '__main__':
    sys.exit(main())
