# -*- coding: utf-8 -*-
"""Verificacao ESTATICA de um modulo VBA antes de abri-lo no Excel (um erro de compilacao abre a janela do VBA na
tela do usuario). Pega os dois erros que ja escaparam:
  - nome declarado duas vezes no mesmo procedimento (o VBA nao diferencia maiusculas: 'c' e 'C' sao o mesmo);
  - variavel que recebe valor sem estar declarada (Option Explicit).
Uso: python lint_vba.py <modulo.bas> [...]   -> sai com 1 se achar problema.
"""
import re
import sys

PALAVRAS = {'if', 'then', 'else', 'elseif', 'end', 'for', 'each', 'in', 'next', 'do', 'loop', 'while', 'until', 'wend',
            'select', 'case', 'with', 'set', 'let', 'dim', 'redim', 'preserve', 'exit', 'function', 'sub', 'property',
            'on', 'error', 'resume', 'goto', 'call', 'erase', 'const', 'private', 'public', 'static', 'option'}


def procedimentos(t):
    t = re.sub(r' _\n\s*', ' ', t)                      # junta as linhas continuadas
    padrao = re.compile(r'^(?:Private |Public )?(?:Function|Sub) (\w+)\(([^\n]*)\)(?: As \w+(?:\(\))?)?[ \t]*\n(.*?)^End (?:Function|Sub)',
                        re.M | re.S)
    for m in padrao.finditer(t):
        yield m.group(1), m.group(2), m.group(3)


def nomes_modulo(t):
    n = set()
    for m in re.finditer(r'^(?:Private|Public|Dim) (?:Const )?(\w+)', t, re.M):
        n.add(m.group(1).lower())
    for m in re.finditer(r"^(?:Private|Public|Dim) (?!Const |Function |Sub |Declare |Type |Enum )([^\n']*)", t, re.M):
        for parte in m.group(1).split(',')[1:]:                         # Private a As Long, b As Double
            v = re.match(r'\s*(?:WithEvents\s+)?(\w+)', parte)
            if v:
                n.add(v.group(1).lower())
    for m in re.finditer(r'(?:,\s*)(\w+) As Long = ', t):                  # Const A As Long = 1, B As Long = 2
        n.add(m.group(1).lower())
    for m in re.finditer(r'^(?:Private |Public )?(?:Function|Sub) (\w+)', t, re.M):
        n.add(m.group(1).lower())
    return n


def checar(caminho):
    t = open(caminho, encoding='utf-8').read().replace('\r\n', '\n')
    mod = nomes_modulo(t)
    problemas = []
    for nome, params, corpo in procedimentos(t):
        decl = []
        for prm in params.split(','):
            mm = re.search(r'(\w+)(?:\(\))?\s+As\b', prm)
            if mm:
                decl.append(mm.group(1))
        for d in re.finditer(r'^\s*(?:Dim|Static) (.*)$', corpo, re.M):
            for parte in re.split(r',(?![^(]*\))', d.group(1)):
                mm = re.match(r'\s*(\w+)', parte)
                if mm:
                    decl.append(mm.group(1))
        baixos = [d.lower() for d in decl]
        for d in sorted({x for x in baixos if baixos.count(x) > 1}):
            problemas.append(f'{nome}: "{d}" declarado mais de uma vez (o VBA nao diferencia maiusculas)')
        for d in sorted({x for x in baixos if x in mod}):
            problemas.append(f'{nome}: local "{d}" tem o nome de um procedimento/constante do modulo (o VBA nao diferencia maiusculas)')
        for linha in corpo.split('\n'):
            limpa = re.sub(r"'.*$", '', re.sub(r'"[^"]*"', '""', linha))
            if re.search(r'\bThen\b\s*\S', limpa) and re.search(r'\bElseIf\b', limpa):
                problemas.append(f'{nome}: ElseIf em If de uma linha (o VBA so aceita Else): {linha.strip()[:60]}')
        # dentro da Function X, 'X(...) = valor' e lido como chamada a propria funcao (tipos incompativeis)
        for linha in corpo.split('\n'):
            limpa = re.sub(r"'.*$", '', re.sub(r'"[^"]*"', '""', linha))
            if re.search(r'(?:^|:|\bThen\b|\bElse\b)\s*' + re.escape(nome) + r'\s*\([^=]*\)\s*=(?!=)', limpa, re.I):
                problemas.append(f'{nome}: "{nome}(...) = " dentro da propria funcao (use uma variavel local): {linha.strip()[:60]}')
        conhecidos = set(baixos) | mod | {nome.lower()}
        for linha in corpo.split('\n'):
            sem_coment = re.sub(r"'.*$", '', re.sub(r'"[^"]*"', '""', linha))
            for trecho in re.split(r'(?<!\w)(?:Then|Else|:)(?!\w)', sem_coment):
                mm = re.match(r'\s*(?:Set\s+)?(\w+)\s*(?:\([^=]*\))?\s*=(?!=)', trecho)
                if mm:
                    alvo = mm.group(1).lower()
                    if alvo not in conhecidos and alvo not in PALAVRAS:
                        problemas.append(f'{nome}: "{mm.group(1)}" recebe valor sem estar declarada')
    return sorted(set(problemas))


if __name__ == '__main__':
    erros = []
    for arq in sys.argv[1:]:
        for p in checar(arq):
            erros.append(f'{arq}: {p}')
    print('\n'.join(erros) if erros else 'lint VBA: OK')
    sys.exit(1 if erros else 0)
