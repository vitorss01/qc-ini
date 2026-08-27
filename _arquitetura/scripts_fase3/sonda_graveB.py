# -*- coding: utf-8 -*-
"""sonda_graveB.py - as duas AtualizarEstatistica da Bioquimica, lado a lado"""
import io
import os
import re
import sys

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', write_through=True)

D = (r"C:\Users\vitor\AppData\Local\Temp\claude"
     r"\C--Users-vitor-OneDrive---MSFT-Desktop-QC-INI"
     r"\919a6d1a-4dbb-471e-9759-5f896c9a1c9b\scratchpad\vba_Bioquimica")


def ler(p):
    for enc in ('cp1252', 'utf-8'):
        try:
            return io.open(p, encoding=enc).read()
        except (UnicodeDecodeError, IOError):
            continue
    return ''


def corpo(cod, nome):
    m = re.search(r'^[ \t]*(?:Public |Private )?Sub[ \t]+%s[ \t]*\(' % nome,
                  cod, re.M | re.I)
    if not m:
        return None
    resto = cod[m.start():]
    f = re.search(r'^[ \t]*End Sub', resto, re.M)
    return resto[:f.end() if f else 1500]


for mod in ('mUI.bas', 'mEstatistica.bas'):
    cod = ler(os.path.join(D, mod))
    c = corpo(cod, 'AtualizarEstatistica')
    print('=' * 78)
    print('%s :: AtualizarEstatistica  (%d linhas)'
          % (mod, c.count('\n') + 1 if c else 0))
    print('=' * 78)
    if c:
        for l in c.split('\n')[:34]:
            print('   ' + l.rstrip()[:96])
    print()

print('=' * 78)
print('CONTEXTO DAS CHAMADAS SEM QUALIFICAR')
print('=' * 78)
for f, alvo in (('mOperacao.bas', 76), ('mUI.bas', 103)):
    cod = ler(os.path.join(D, f)).split('\n')
    print('--- %s, linhas %d a %d ---' % (f, alvo - 6, alvo + 3))
    for i in range(max(0, alvo - 7), min(len(cod), alvo + 3)):
        marca = '>>' if i + 1 == alvo else '  '
        print('  %s %4d  %s' % (marca, i + 1, cod[i].rstrip()[:92]))
    print()
