# -*- coding: utf-8 -*-
"""sonda_mimportar.py - as duas linhagens do mImportar

A versao COM a guarda do ADR-046 tem 323 linhas; a de producao, 385. A corrigida
e a mais CURTA, entao nao e a mesma com um patch: sao linhagens diferentes.
Antes de propagar qualquer coisa, e preciso saber o que a de producao tem a
mais -- se for funcionalidade, sobrescrever seria perda.
"""
import io
import os
import re
import sys
import difflib

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', write_through=True)

TMP = (r"C:\Users\vitor\AppData\Local\Temp\claude"
       r"\C--Users-vitor-OneDrive---MSFT-Desktop-QC-INI"
       r"\919a6d1a-4dbb-471e-9759-5f896c9a1c9b\scratchpad")
BASE = r"C:\Users\vitor\OneDrive - MSFT\Desktop\QC_INI\_arquitetura"

A = os.path.join(BASE, 'src_producao', 'mImportar.bas')          # producao
B = os.path.join(BASE, 'src_hardening1', 'mImportar.bas')        # com a guarda


def ler(p):
    for enc in ('cp1252', 'utf-8'):
        try:
            return io.open(p, encoding=enc).read()
        except (UnicodeDecodeError, IOError):
            continue
    return ''


def procs(cod):
    return [m.group(3) for m in re.finditer(
        r'^[ \t]*(Public |Private )?(Sub|Function)[ \t]+(\w+)', cod, re.M)]


a, b = ler(A), ler(B)
pa, pb = procs(a), procs(b)

print('=== procedures ===')
print('  src_producao   (%d linhas): %d procedures' % (a.count('\n'), len(pa)))
print('  src_hardening1 (%d linhas): %d procedures' % (b.count('\n'), len(pb)))
soA = [p for p in pa if p not in pb]
soB = [p for p in pb if p not in pa]
print('  so na producao      : %s' % (soA or 'nenhuma'))
print('  so no hardening1    : %s' % (soB or 'nenhuma'))

print()
print('=== o que a producao tem a mais, linha a linha ===')
la = [l.rstrip() for l in a.replace('\r\n', '\n').split('\n')]
lb = [l.rstrip() for l in b.replace('\r\n', '\n').split('\n')]
d = list(difflib.unified_diff(lb, la, 'hardening1(com guarda)', 'producao', n=1,
                              lineterm=''))
print('  %d linhas de diff' % len(d))
soment = [l for l in d if l.startswith('+') and not l.startswith('+++')]
somenos = [l for l in d if l.startswith('-') and not l.startswith('---')]
print('  so na producao: %d linhas | so no hardening1: %d linhas'
      % (len(soment), len(somenos)))
print()
print('  --- amostra do que a PRODUCAO tem e o hardening1 nao ---')
for l in soment[:40]:
    print('    ' + l[:100])
print()
print('  --- amostra do que o HARDENING1 tem e a producao nao ---')
for l in somenos[:40]:
    print('    ' + l[:100])
