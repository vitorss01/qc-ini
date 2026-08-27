# -*- coding: utf-8 -*-
"""sonda_graves.py - os dois GRAVES da Bioquimica, medidos

B. AtualizarEstatistica publico em mEstatistica E mUI
G. mImportar.MostrarErros / LimparAreaImport destrancam sem On Error

Para cada um: o que ha na producao, o que ha na fonte, e quem chama.
"""
import io
import os
import re
import sys

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', write_through=True)

TMP = (r"C:\Users\vitor\AppData\Local\Temp\claude"
       r"\C--Users-vitor-OneDrive---MSFT-Desktop-QC-INI"
       r"\919a6d1a-4dbb-471e-9759-5f896c9a1c9b\scratchpad")
BASE = r"C:\Users\vitor\OneDrive - MSFT\Desktop\QC_INI\_arquitetura"


def ler(p):
    for enc in ('cp1252', 'utf-8'):
        try:
            return io.open(p, encoding=enc).read()
        except (UnicodeDecodeError, IOError):
            continue
    return ''


def corpo(cod, nome):
    m = re.search(r'^\s*(?:Public |Private )?(?:Sub|Function)\s+%s\b' % nome,
                  cod, re.M | re.I)
    if not m:
        return None
    fim = re.search(r'^\s*End (?:Sub|Function)', cod[m.start():], re.M)
    return cod[m.start():m.start() + (fim.end() if fim else 600)]


print('=' * 78)
print('GRAVE B -- AtualizarEstatistica em dois modulos')
print('=' * 78)
for prod in ('Hematologia', 'Bioquimica', 'Imunologia'):
    d = os.path.join(TMP, 'vba_%s' % prod)
    if not os.path.isdir(d):
        continue
    onde = []
    for f in sorted(os.listdir(d)):
        cod = ler(os.path.join(d, f))
        for m in re.finditer(r'^\s*(Public |Private )?(Sub|Function)\s+'
                             r'(AtualizarEstatistica)\s*\(', cod, re.M | re.I):
            onde.append('%s (%s)' % (f, (m.group(1) or 'Public ').strip() or 'Public'))
    print('  %-12s %s' % (prod, onde or 'nao existe'))

print()
print('  quem chama AtualizarEstatistica sem qualificar (Bioquimica):')
d = os.path.join(TMP, 'vba_Bioquimica')
for f in sorted(os.listdir(d)):
    cod = ler(os.path.join(d, f))
    for i, l in enumerate(cod.split('\n'), 1):
        if re.search(r'(?<![.\w])AtualizarEstatistica\b', l) and \
           not re.search(r'(Sub|Function)\s+AtualizarEstatistica', l, re.I):
            qual = 'QUALIFICADA' if re.search(r'\w+\.AtualizarEstatistica', l) else 'SEM QUALIFICAR'
            print('     %-28s l.%-4d %-14s %s' % (f, i, qual, l.strip()[:60]))

print()
print('=' * 78)
print('GRAVE G -- mImportar destranca sem On Error')
print('=' * 78)
for rot, cam in (('producao (xlsm)', os.path.join(TMP, 'vba_Bioquimica', 'mImportar.bas')),
                 ('src_hardening1/Bioquimica',
                  os.path.join(BASE, 'src_hardening1', 'Bioquimica', 'mImportar.bas')),
                 ('src_hardening1', os.path.join(BASE, 'src_hardening1', 'mImportar.bas')),
                 ('src_producao', os.path.join(BASE, 'src_producao', 'mImportar.bas'))):
    if not os.path.exists(cam):
        print('  %-28s (arquivo ausente)' % rot)
        continue
    cod = ler(cam)
    print('  %-28s %d linhas' % (rot, cod.count('\n')))
    for nome in ('MostrarErros', 'LimparAreaImport'):
        c = corpo(cod, nome)
        if c is None:
            print('     %-18s ausente' % nome)
            continue
        temGuarda = bool(re.search(r'LiberarEscrita|RestaurarProtecao', c))
        temOnErr = bool(re.search(r'On Error', c))
        temUnp = bool(re.search(r'Unprotect', c, re.I))
        print('     %-18s Unprotect=%-5s On Error=%-5s guarda(mSeguranca)=%s'
              % (nome, temUnp, temOnErr, temGuarda))
