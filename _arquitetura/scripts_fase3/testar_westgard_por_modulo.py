# -*- coding: utf-8 -*-
"""testar_westgard_por_modulo.py - ADR-041: as duas matrizes, provadas

POR QUE ESTE TESTE EXISTE

Ate aqui a unica prova de que as regras estavam certas era a funcao devolver
os nomes certos. Nome nao e semantica: foi assim que "o plano pede 8x e o
motor conta 10" atravessou o projeto respondendo cobertura TOTAL.

Este teste injeta series de z-score construidas para disparar UMA regra
especifica e confere se o motor dispara -- e, tao importante quanto, se NAO
dispara quando nao deve. A fronteira e o que prova o limiar: 7 consecutivos
nao podem acender 8x; 8 tem de acender.

O QUE MEDE

  A. Matriz por modulo ..... MatrizWestgard() devolve a familia do produto
  B. Anti-contaminacao ..... nenhuma regra da outra familia na matriz, e
                             nenhuma da propria familia faltando
  C. Limiares .............. o motor publica os numeros que o nome promete
  D. Semantica ............. series construidas acendem a regra certa
  E. Fronteira ............. n-1 consecutivos NAO acendem; n acende
  F. Cobertura ............. TOTAL so quando nome E limiar concordam

Uso:
    python testar_westgard_por_modulo.py <arquivo.xlsm>
"""
import io
import os
import sys

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8',
                              write_through=True)
import win32com.client as w

BIO = ['1_3s', '2_2s', 'R_4s', '4_1s', '8x']
HEMA = ['1_3s', '2of3_2s', 'R_4s', '3_1s', '6x']

resultados = []


def anota(bloco, teste, esperado, obtido):
    ok = str(esperado).strip() == str(obtido).strip()
    resultados.append((bloco, teste, esperado, obtido, ok))
    return ok


def novo_excel(caminho):
    xl = w.DispatchEx('Excel.Application')
    xl.Visible = False
    xl.DisplayAlerts = False
    xl.EnableEvents = False
    xl.AutomationSecurity = 1
    return xl, xl.Workbooks.Open(caminho, 0, True)


def serie(nlv, nrun, pontos):
    """Monta z(nivel, corrida) e temDado. pontos: {(nivel, corrida): z}.
    Corridas sem ponto ficam com temDado=False -- o motor nao pode tratar
    ausencia como zero, e o teste depende disso."""
    z = [[0.0] * (nrun + 1) for _ in range(nlv)]
    td = [[False] * (nrun + 1) for _ in range(nlv)]
    for (t, i), v in pontos.items():
        z[t][i] = float(v)
        td[t][i] = True
    return z, td


def avaliar(xl, nlv, nrun, pontos):
    """Chama AvaliarWestgard via uma macro-ponte e devolve quais regras
    dispararam. A ponte existe porque VBA nao aceita array 2D vindo do COM
    com os tipos que a assinatura pede (ByRef Double() e Boolean())."""
    z, td = serie(nlv, nrun, pontos)
    # monta o texto da serie: nivel|corrida|z por ponto
    itens = ';'.join('%d|%d|%r' % (t, i, z[t][i])
                     for t in range(nlv) for i in range(1, nrun + 1) if td[t][i])
    return xl.Run('TestarWestgardSerie', itens, nrun)


PONTE = '''
Public Function TestarWestgardSerie(ByVal itens As String, ByVal nRun As Long) As String
    Dim z() As Double, td() As Boolean
    Dim r13 As Variant, r22 As Variant, rR4 As Variant, r41 As Variant
    Dim r10 As Variant, a12 As Variant
    Dim p As Variant, campo As Variant, t As Long, i As Long, k As Long

    ReDim z(0 To NLV - 1, 1 To nRun)
    ReDim td(0 To NLV - 1, 1 To nRun)
    ReDim r13(0 To NLV - 1, 1 To nRun): ReDim r22(0 To NLV - 1, 1 To nRun)
    ReDim rR4(0 To NLV - 1, 1 To nRun): ReDim r41(0 To NLV - 1, 1 To nRun)
    ReDim r10(0 To NLV - 1, 1 To nRun): ReDim a12(0 To NLV - 1, 1 To nRun)

    p = Split(itens, ";")
    For k = LBound(p) To UBound(p)
        If Len(Trim$(CStr(p(k)))) > 0 Then
            campo = Split(CStr(p(k)), "|")
            t = CLng(campo(0)): i = CLng(campo(1))
            z(t, i) = CDbl(Replace(CStr(campo(2)), ".", ","))
            td(t, i) = True
        End If
    Next k

    AvaliarWestgard z, td, nRun, r13, r22, rR4, r41, r10, a12

    Dim saida As String
    saida = "1_3s=" & Soma(r13) & ";R2=" & Soma(r22) & ";R_4s=" & Soma(rR4) & _
            ";R4=" & Soma(r41) & ";TEND=" & Soma(r10) & ";A12=" & Soma(a12)
    TestarWestgardSerie = saida
End Function

Private Function Soma(ByRef m As Variant) As Long
    Dim t As Long, i As Long, s As Long
    For t = LBound(m, 1) To UBound(m, 1)
        For i = LBound(m, 2) To UBound(m, 2)
            If Not IsEmpty(m(t, i)) Then If m(t, i) = 1 Then s = s + 1
        Next i
    Next t
    Soma = s
End Function
'''


def instalar_ponte(wb):
    vbp = wb.VBProject
    for c in list(vbp.VBComponents):
        if c.Name == 'mTesteWestgard':
            vbp.VBComponents.Remove(c)
            break
    m = vbp.VBComponents.Add(1)      # vbext_ct_StdModule
    m.Name = 'mTesteWestgard'
    m.CodeModule.AddFromString(PONTE)


def conta(saida, chave):
    for parte in str(saida).split(';'):
        if parte.startswith(chave + '='):
            return int(parte.split('=')[1])
    return -1


def main():
    caminho = os.path.abspath(sys.argv[1])
    xl, wb = novo_excel(caminho)
    try:
        nlv = int(xl.Run('NiveisDoProduto')) if False else None
    except Exception:
        nlv = None
    try:
        matriz = str(xl.Run('MatrizWestgard'))
        regras = [r.strip() for r in matriz.split(';')]
        nlv = 3 if '2of3_2s' in regras else 2
        familia = HEMA if nlv == 3 else BIO
        modulo = 'HEMATOLOGIA' if nlv == 3 else 'BIOQUIMICA'

        print('=' * 78)
        print('%s -- %d niveis' % (modulo, nlv))
        print('=' * 78)
        print('   matriz: %s' % ' / '.join(regras))

        # ---- A. matriz do modulo
        anota('A', 'matriz do modulo', ';'.join(familia), matriz)

        # ---- B. anti-contaminacao
        alheias = BIO if nlv == 3 else HEMA
        alheias = [r for r in alheias if r not in familia] + ['10x']
        for r in familia:
            anota('B', 'propria presente: %s' % r, True, r in regras)
        for r in alheias:
            anota('B', 'alheia ausente: %s' % r, True, r not in regras)

        erro = str(xl.Run('ValidarMatrizWestgard') or '')
        anota('B', 'ValidarMatrizWestgard sem erro', '', erro)
        if erro:
            print('   VALIDACAO: %s' % erro)

        # ---- C. limiares publicados
        seq = int(xl.Run('LimiarSequencialWestgard'))
        um = int(xl.Run('LimiarUmSigmaWestgard'))
        print('   limiares: sequencial=%d  um-sigma=%d' % (seq, um))
        anota('C', 'limiar sequencial', 6 if nlv == 3 else 8, seq)
        anota('C', 'limiar de 1s', 3 if nlv == 3 else 4, um)

        # ---- D/E. semantica e fronteira
        instalar_ponte(wb)
        print()
        print('   %-38s %-8s %-8s %s' % ('cenario', 'esperado', 'obtido', ''))
        print('   ' + '-' * 70)

        def cenario(bloco, nome, pontos, nrun, chave, esperado):
            saida = avaliar(xl, nlv, nrun, pontos)
            got = conta(saida, chave)
            ok = (got > 0) if esperado else (got == 0)
            resultados.append((bloco, nome, 'dispara' if esperado else 'nao dispara',
                               '%d marca(s)' % got, ok))
            print('   %-38s %-8s %-8s %s'
                  % (nome, 'dispara' if esperado else 'nao', '%d' % got,
                     'PASS' if ok else 'FAIL'))

        # 1_3s: um ponto alem de 3s
        cenario('D', '1_3s: z=+3,4 num nivel', {(0, 1): 3.4}, 1, '1_3s', True)
        cenario('E', '1_3s: z=+2,9 nao dispara', {(0, 1): 2.9}, 1, '1_3s', False)

        if nlv == 3:
            # 2of3_2s: dois dos tres niveis alem de 2s, mesmo lado
            cenario('D', '2of3_2s: +2,3 +2,4 +0,5',
                    {(0, 1): 2.3, (1, 1): 2.4, (2, 1): 0.5}, 1, 'R2', True)
            cenario('E', '2of3_2s: so um nivel alem de 2s',
                    {(0, 1): 2.3, (1, 1): 0.4, (2, 1): 0.5}, 1, 'R2', False)
            # 3_1s: os tres niveis alem de 1s, mesmo lado
            cenario('D', '3_1s: +1,2 +1,4 +1,1',
                    {(0, 1): 1.2, (1, 1): 1.4, (2, 1): 1.1}, 1, 'R4', True)
            cenario('E', '3_1s: so dois niveis alem de 1s',
                    {(0, 1): 1.2, (1, 1): 1.4, (2, 1): 0.3}, 1, 'R4', False)
            # 6x: seis corridas consecutivas do mesmo lado, no mesmo nivel
            cenario('D', '6x: 6 corridas do mesmo lado',
                    {(0, i): 0.5 for i in range(1, 7)}, 6, 'TEND', True)
            cenario('E', '6x: 5 corridas NAO dispara',
                    {(0, i): 0.5 for i in range(1, 6)}, 5, 'TEND', False)
        else:
            # 2_2s: dois consecutivos no mesmo nivel
            cenario('D', '2_2s: +2,3 e +2,4 seguidos',
                    {(0, 1): 2.3, (0, 2): 2.4}, 2, 'R2', True)
            # 4_1s: quatro consecutivos alem de 1s
            cenario('D', '4_1s: 4 corridas alem de 1s',
                    {(0, i): 1.2 for i in range(1, 5)}, 4, 'R4', True)
            cenario('E', '4_1s: 3 corridas NAO dispara',
                    {(0, i): 1.2 for i in range(1, 4)}, 3, 'R4', False)
            # 8x: a fronteira que a secao 11 exige
            cenario('D', '8x: 8 corridas do mesmo lado',
                    {(0, i): 0.5 for i in range(1, 9)}, 8, 'TEND', True)
            cenario('E', '8x: 7 corridas NAO dispara',
                    {(0, i): 0.5 for i in range(1, 8)}, 7, 'TEND', False)

        # R_4s vale nos dois: amplitude entre niveis > 4 DP
        cenario('D', 'R_4s: +2,2 e -2,1 na mesma corrida',
                {(0, 1): 2.2, (1, 1): -2.1}, 1, 'R_4s', True)

        # ---- F. cobertura
        cob = str(xl.Run('CoberturaWestgard', 3.5))
        print()
        print('   CoberturaWestgard(3,5) = %s' % cob)
        anota('F', 'cobertura TOTAL', 'TOTAL', cob)

    finally:
        wb.Close(False)
        xl.Quit()

    print()
    print('=' * 78)
    blocos = {}
    for b, t, e, o, ok in resultados:
        d = blocos.setdefault(b, [0, 0])
        d[0 if ok else 1] += 1
    for b in sorted(blocos):
        print('   bloco %s: %d PASS, %d FAIL' % (b, blocos[b][0], blocos[b][1]))
    falhas = [r for r in resultados if not r[4]]
    for b, t, e, o, ok in falhas:
        print('   FAIL [%s] %s | esperado=%s | obtido=%s' % (b, t, e, o))
    print('TOTAL: %d PASS, %d FAIL'
          % (len(resultados) - len(falhas), len(falhas)))
    return 1 if falhas else 0


if __name__ == '__main__':
    sys.exit(main())
