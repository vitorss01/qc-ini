# -*- coding: utf-8 -*-
"""checar_abas_citadas.py - todo Sheets("X") do VBA aponta para uma aba que existe?

Erro 9 -- "subscrito fora do intervalo" -- e o que o VBA devolve quando se pede
Sheets("Nome") de uma aba que nao existe. Com o Excel invisivel isso vira um
dialogo modal que trava a automacao sem dizer qual nome falhou.

Esta varredura antecipa o defeito sem abrir o Excel: cruza cada literal de nome
de aba citado no codigo com a lista real de abas do arquivo.

Um nome citado e ausente NAO e necessariamente um defeito -- pode estar num
caminho protegido por On Error, ou numa rotina que so roda em outro produto.
Por isso o relatorio separa "citado em rotina sem On Error" do resto.
"""
import io
import os
import re
import sys

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', write_through=True)
import openpyxl
from oletools.olevba import VBA_Parser

PRODUTOS = ['QC_Hematologia.xlsm', 'QC_Bioquimica.xlsm', 'QC_Imunologia.xlsm']
BASE = r"C:\Users\vitor\OneDrive - MSFT\Desktop\QC_INI"

# Sheets("X") / Worksheets("X") / Sheets("X").  -- e as constantes que guardam nome
RX_SHEET = re.compile(r'(?:Sheets|Worksheets)\s*\(\s*"([^"]+)"\s*\)')
RX_CONST = re.compile(r'^\s*(?:Public|Private)?\s*Const\s+(\w+)\s+As\s+String\s*=\s*"([^"]+)"',
                      re.M | re.I)
RX_SHEET_CONST = re.compile(r'(?:Sheets|Worksheets)\s*\(\s*(\w+)\s*\)')


def proc_de(cod, pos):
    """nome da procedure que contem a posicao, e se ela tem On Error"""
    ini = 0
    nome = '(nivel de modulo)'
    for m in re.finditer(r'^[ \t]*(?:Public |Private )?(?:Sub|Function)[ \t]+(\w+)',
                         cod, re.M):
        if m.start() > pos:
            break
        ini, nome = m.start(), m.group(1)
    fim = cod.find('\nEnd Sub', pos)
    f2 = cod.find('\nEnd Function', pos)
    if f2 != -1 and (fim == -1 or f2 < fim):
        fim = f2
    corpo = cod[ini:fim if fim > 0 else len(cod)]
    return nome, bool(re.search(r'On Error', corpo))


for arq in PRODUTOS:
    cam = os.path.join(BASE, arq)
    if not os.path.exists(cam):
        continue
    wb = openpyxl.load_workbook(cam, read_only=True)
    abas = set(wb.sheetnames)
    wb.close()

    p = VBA_Parser(cam)
    mods = {}
    for _, _, nome, cod in p.extract_macros():
        if nome:
            mods[nome.replace('.bas', '').replace('.cls', '').replace('.frm', '')] = cod or ''
    p.close()

    print()
    print('=' * 78)
    print('%s  (%d abas)' % (arq, len(abas)))
    print('=' * 78)

    faltando = []
    for mod, cod in sorted(mods.items()):
        consts = dict(RX_CONST.findall(cod))
        citados = []
        for m in RX_SHEET.finditer(cod):
            citados.append((m.group(1), m.start()))
        for m in RX_SHEET_CONST.finditer(cod):
            if m.group(1) in consts:
                citados.append((consts[m.group(1)], m.start()))
        for nome, pos in citados:
            if nome in abas:
                continue
            proc, temOnErr = proc_de(cod, pos)
            faltando.append((mod, proc, nome, temOnErr))

    if not faltando:
        print('  todos os nomes de aba citados existem')
        continue

    semguarda = [f for f in faltando if not f[3]]
    comguarda = [f for f in faltando if f[3]]

    print('  SEM On Error (erro 9 vira dialogo modal): %d' % len(semguarda))
    vistos = set()
    for mod, proc, nome, _ in semguarda:
        k = (mod, proc, nome)
        if k in vistos:
            continue
        vistos.add(k)
        print('     %-16s %-28s -> aba %r AUSENTE' % (mod, proc, nome))
    if comguarda:
        print('  com On Error (degrada, nao trava): %d' % len(comguarda))
        vistos = set()
        for mod, proc, nome, _ in comguarda:
            k = (mod, proc, nome)
            if k in vistos:
                continue
            vistos.add(k)
            print('     %-16s %-28s -> aba %r ausente' % (mod, proc, nome))
