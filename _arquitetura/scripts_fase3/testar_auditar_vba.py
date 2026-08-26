# -*- coding: utf-8 -*-
"""testar_auditar_vba.py - o auditor estatico reprova o que deve

POR QUE ESTE TESTE EXISTE

auditar_vba.py rodou sobre o projeto e devolveu quase nada. "Quase nada" tem
duas explicacoes possiveis -- o projeto esta limpo, ou o analisador esta cego --
e as duas se parecem exatamente igual na tela. Um analisador que nunca acusa e
pior do que nenhum: da a sensacao de cobertura sem a cobertura.

COMO

Copia os modulos para uma pasta temporaria, injeta UM defeito conhecido de cada
categoria, e confirma que a categoria correspondente aparece. Cada caso e
independente: a pasta e reconstruida a cada rodada.

Uso: python testar_auditar_vba.py <pasta_com_modulos_exportados>
"""
import io
import os
import shutil
import subprocess
import sys
import tempfile

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8',
                              write_through=True)

AUDITOR = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                       'auditar_vba.py')

# nome ; categoria esperada ; funcao que sabota
CASOS = []


def caso(nome, categoria):
    def deco(fn):
        CASOS.append((nome, categoria, fn))
        return fn
    return deco


def escrever(pasta, arquivo, linhas):
    io.open(os.path.join(pasta, arquivo), 'w', encoding='cp1252',
            newline='\r\n').write('\n'.join(linhas))


def ler(pasta, arquivo):
    return io.open(os.path.join(pasta, arquivo), encoding='cp1252',
                   errors='replace').read().replace('\r\n', '\n').split('\n')


@caso('chamada de Sub que nao existe', 'A. chamada sem definicao')
def _a(pasta):
    L = ler(pasta, 'mQualidade.bas')
    L.append('Public Sub IscaA()')
    L.append('    Call RotinaQueNuncaExistiu')
    L.append('End Sub')
    escrever(pasta, 'mQualidade.bas', L)


@caso('Application.Run de macro inexistente', 'A. chamada sem definicao')
def _a2(pasta):
    L = ler(pasta, 'mQualidade.bas')
    L.append('Public Sub IscaA2()')
    L.append('    Application.Run "MacroFantasma"')
    L.append('End Sub')
    escrever(pasta, 'mQualidade.bas', L)


@caso('mModulo.Membro inexistente', 'A. chamada sem definicao')
def _a3(pasta):
    L = ler(pasta, 'mQualidade.bas')
    L.append('Public Sub IscaA3()')
    L.append('    mBI.FuncaoQueNaoExisteNoBI 1, 2')
    L.append('End Sub')
    escrever(pasta, 'mQualidade.bas', L)


@caso('mesmo nome Public em dois modulos', 'B. nome PUBLICO duplicado')
def _b(pasta):
    L = ler(pasta, 'mQualidade.bas')
    L.append('Public Function ClassificarSigmaDuplo() As String')
    L.append('End Function')
    escrever(pasta, 'mQualidade.bas', L)
    L = ler(pasta, 'mPlanoQC.bas')
    L.append('Public Function ClassificarSigmaDuplo() As String')
    L.append('End Function')
    escrever(pasta, 'mPlanoQC.bas', L)


@caso('procedure repetida no mesmo modulo', 'C. procedure duplicada')
def _c(pasta):
    L = ler(pasta, 'mQualidade.bas')
    L += ['Public Sub IscaC()', 'End Sub',
          'Public Sub IscaC()', 'End Sub']
    escrever(pasta, 'mQualidade.bas', L)


@caso('Dim de modulo depois da 1a procedure', 'D. declaracao de modulo apos')
def _d(pasta):
    L = ler(pasta, 'mQualidade.bas')
    L.append('Private gIscaD As Object')
    escrever(pasta, 'mQualidade.bas', L)


@caso('On Error GoTo para label inexistente', 'F. GoTo/On Error para label')
def _f(pasta):
    L = ler(pasta, 'mQualidade.bas')
    L += ['Public Sub IscaF()',
          '    On Error GoTo labelQueNaoExiste',
          '    Debug.Print 1',
          'End Sub']
    escrever(pasta, 'mQualidade.bas', L)


@caso('desprotege e nunca reprotege', 'G. desprotege e NAO reprotege')
def _g(pasta):
    L = ler(pasta, 'mQualidade.bas')
    L += ['Public Sub IscaG()',
          '    Dim ws As Worksheet',
          '    Set ws = ThisWorkbook.Sheets("Calc")',
          '    ws.Unprotect Password:="x"',
          '    ws.Cells(1, 1).Value = 1',
          'End Sub']
    escrever(pasta, 'mQualidade.bas', L)


@caso('Exit Sub dentro da janela destrancada',
      'G. saida antecipada dentro da janela')
def _g2(pasta):
    L = ler(pasta, 'mQualidade.bas')
    L += ['Public Sub IscaG2()',
          '    Dim ws As Worksheet',
          '    Set ws = ThisWorkbook.Sheets("Calc")',
          '    ws.Unprotect Password:="x"',
          '    If 1 = 1 Then Exit Sub',
          '    ws.Protect Password:="x"',
          'End Sub']
    escrever(pasta, 'mQualidade.bas', L)


@caso('variavel usada sem Dim (Option Explicit)',
      'J. identificador nao declarado')
def _j(pasta):
    L = ler(pasta, 'mQualidade.bas')
    L += ['Public Sub IscaJ()',
          '    Dim x As Long',
          '    x = contadorQueNinguemDeclarou + 1',
          'End Sub']
    escrever(pasta, 'mQualidade.bas', L)


@caso('atribuicao a variavel nao declarada', 'J. identificador nao declarado')
def _j2(pasta):
    L = ler(pasta, 'mQualidade.bas')
    L += ['Public Sub IscaJ2()',
          '    totalSemDeclaracao = 7',
          'End Sub']
    escrever(pasta, 'mQualidade.bas', L)


def rodar(pasta):
    r = subprocess.run([sys.executable, '-u', AUDITOR, pasta],
                       capture_output=True)
    return (r.stdout or b'').decode('utf-8', 'replace')


def main():
    origem = os.path.abspath(sys.argv[1])
    print('=' * 78)
    print('O AUDITOR ESTATICO REPROVA O QUE DEVE?')
    print('=' * 78)
    print()

    base = rodar(origem)
    falhas = 0
    for nome, categoria, sabotar in CASOS:
        tmp = tempfile.mkdtemp(prefix='auditvba_')
        try:
            for f in os.listdir(origem):
                shutil.copy2(os.path.join(origem, f), tmp)
            sabotar(tmp)
            saida = rodar(tmp)
            achou = categoria in saida
            # nao basta detectar: nao pode ser algo que ja aparecia antes
            novo = achou and (categoria not in base
                              or saida.count(categoria) >= base.count(categoria))
            ok = achou and novo
            if not ok:
                falhas += 1
            print('   %-46s %s' % (nome[:46], 'PASS' if ok else 'FAIL'))
            if not ok:
                print('        esperava a categoria %r na saida' % categoria)
        finally:
            shutil.rmtree(tmp, ignore_errors=True)

    print()
    print('TOTAL: %d PASS, %d FAIL' % (len(CASOS) - falhas, falhas))
    return 1 if falhas else 0


if __name__ == '__main__':
    sys.exit(main())
