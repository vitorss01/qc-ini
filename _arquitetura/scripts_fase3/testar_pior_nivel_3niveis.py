# -*- coding: utf-8 -*-
"""testar_pior_nivel_3niveis.py - o pior de TRES niveis governa o plano

POR QUE ESTE TESTE EXISTE

PreencherPlanoDoPiorNivel foi escrito generico -- percorre os niveis que
existirem. "Generico" nao e prova: com dois niveis ele nunca exercitou o
caso de tres, e o proprio ADR-038 nasceu de um defeito em que o nivel errado
governava. O teste executa os cenarios exigidos, incluindo os de ausencia.

O QUE MEDE

  A. tres niveis validos ........... o MENOR governa, e o rotulo aponta ele
  B. um nivel ausente .............. governa o menor entre os presentes
  C. dois ausentes ................. governa o unico presente
  D. todos ausentes ................ SEM DADOS, plano vazio
  E. Sigma invalido ................ zero, texto e erro nao sao Sigma

A avaliacao roda sobre a MESMA rotina do motor (ponte VBA), e nao sobre uma
reimplementacao em Python -- reimplementar provaria o teste, nao o motor.

Uso: python testar_pior_nivel_3niveis.py <arquivo.xlsm>
"""
import io
import os
import sys

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8',
                              write_through=True)
import win32com.client as w

PONTE = '''
Public Function TestarPiorNivel(ByVal itens As String) As String
    ' itens: "nivel|sigma;nivel|sigma;..."  sigma vazio = ausente
    Dim saida() As Variant, p As Variant, c As Variant, k As Long, n As Long
    p = Split(itens, ";")
    n = UBound(p) - LBound(p) + 1
    ReDim saida(1 To n, 1 To 84)
    For k = LBound(p) To UBound(p)
        c = Split(CStr(p(k)), "|")
        saida(k + 1, 8) = "TESTE"          ' ID_Analito
        saida(k + 1, 12) = "L1"            ' ID_Lote
        saida(k + 1, 14) = CLng(c(0))      ' Nivel
        If Len(Trim$(CStr(c(1)))) > 0 Then
            If IsNumeric(Replace(CStr(c(1)), ".", ",")) Then
                saida(k + 1, 53) = CDbl(Replace(CStr(c(1)), ".", ","))
            Else
                saida(k + 1, 53) = CStr(c(1))     ' texto: tem de ser recusado
            End If
        End If
    Next k

    mBI.PreencherPlanoDoPiorNivel saida, n

    TestarPiorNivel = "plano=" & CStr(saida(1, 77)) & _
                      ";governa=" & CStr(saida(1, 78)) & _
                      ";classe=" & CStr(saida(1, 79)) & _
                      ";regras=" & CStr(saida(1, 71)) & _
                      ";N=" & CStr(saida(1, 72)) & _
                      ";run=" & CStr(saida(1, 73))
End Function
'''

CENARIOS = [
    ('A. N1=7,0 N2=6,5 N3=2,8', '1|7.0;2|6.5;3|2.8', '2.8', 'Nivel 3'),
    ('A. N1=3,5 N2=5,0 N3=6,5', '1|3.5;2|5.0;3|6.5', '3.5', 'Nivel 1'),
    ('B. N1 ausente, N2=4,2 N3=5,1', '1|;2|4.2;3|5.1', '4.2', 'Nivel 2'),
    ('C. so N3=6,1', '1|;2|;3|6.1', '6.1', 'Nivel 3'),
    ('D. todos ausentes', '1|;2|;3|', '', 'SEM DADOS'),
    ('E. zero nao e Sigma', '1|0;2|0;3|0', '', 'SEM DADOS'),
    ('E. texto nao e Sigma', '1|abc;2|;3|', '', 'SEM DADOS'),
    ('E. zero convive com valido', '1|0;2|4.5;3|', '4.5', 'Nivel 2'),
]


def campo(saida, chave):
    for parte in str(saida).split(';'):
        if parte.startswith(chave + '='):
            return parte.split('=', 1)[1]
    return ''


def main():
    caminho = os.path.abspath(sys.argv[1])
    xl = w.DispatchEx('Excel.Application')
    xl.Visible = False
    xl.DisplayAlerts = False
    xl.EnableEvents = False
    xl.AutomationSecurity = 1
    wb = xl.Workbooks.Open(caminho, 0, True)
    falhas = 0
    try:
        area = str(xl.Run('AreaDoProduto'))
        print('=' * 78)
        print('%s -- pior nivel governa o plano' % area)
        print('=' * 78)

        vbp = wb.VBProject
        for c in list(vbp.VBComponents):
            if c.Name == 'mTestePiorNivel':
                vbp.VBComponents.Remove(c)
                break
        m = vbp.VBComponents.Add(1)
        m.Name = 'mTestePiorNivel'
        m.CodeModule.AddFromString(PONTE)

        print()
        print('   %-34s %-8s %-12s %s' % ('cenario', 'plano', 'governa', ''))
        print('   ' + '-' * 72)
        for nome, itens, espPlano, espGov in CENARIOS:
            saida = xl.Run('TestarPiorNivel', itens)
            plano = campo(saida, 'plano')
            gov = campo(saida, 'governa')
            # normaliza decimal do VBA (virgula) para comparar
            planoN = plano.replace(',', '.')
            ok = (planoN.rstrip('0').rstrip('.') ==
                  espPlano.rstrip('0').rstrip('.')) and gov == espGov
            if not ok:
                falhas += 1
            print('   %-34s %-8s %-12s %s'
                  % (nome[:34], plano or '(vazio)', gov, 'PASS' if ok else 'FAIL'))
            if not ok:
                print('        esperado plano=%s governa=%s'
                      % (espPlano or '(vazio)', espGov))
            if nome.startswith('D') or nome.startswith('E. zero nao'):
                regras = campo(saida, 'regras')
                if regras:
                    falhas += 1
                    print('        FALHA: sem Sigma o plano deveria ficar vazio, veio %r'
                          % regras)
    finally:
        wb.Close(False)
        xl.Quit()

    print()
    print('TOTAL: %d PASS, %d FAIL' % (len(CENARIOS) - falhas, falhas))
    return 1 if falhas else 0


if __name__ == '__main__':
    sys.exit(main())
