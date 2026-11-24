# -*- coding: utf-8 -*-
"""conferir_assinaturas_liberacao.py -- entrega 46, passo 2 do QC_RELEASE_STATUS (P1 da Liberacao): a assinatura que
esta hoje na aba Liberacao esta ao lado do MESMO RUN com que foi assinada?

So LEITURA, numa COPIA (openpyxl, sem Excel, nada e salvo). Para cada evento ASSINATURA_* do Audit_Log cujo detalhe
cita a aba Liberacao ("signatario: ... | dd/mm/aaaa hh:nn:ss | Liberacao!C10 | RUN 26111602 | ..."), procura na aba o
texto com o mesmo instante e compara o RUN da coluna A dessa linha com o RUN do evento.

Uso: python conferir_assinaturas_liberacao.py <copia.xlsm> [<saida.json>]
Nao imprime nomes de signatarios (so contagens, RUNs e celulas).
"""
import json
import re
import shutil
import sys
import tempfile
import os

import openpyxl

RX_EV = re.compile(r'\|\s*(\d{2}/\d{2}/\d{4} \d{2}:\d{2}:\d{2})\s*\|\s*([^!|]*Libera[^!|]*)!([A-Z]+)(\d+)\s*\|\s*RUN\s*(\S+)')
RX_QUANDO = re.compile(r'(\d{2}/\d{2}/\d{4} \d{2}:\d{2}:\d{2})')


def aba(wb, prefixo):
    for n in wb.sheetnames:
        if n.lower().startswith(prefixo.lower()):
            return wb[n]
    return None


def main(a):
    origem = a[0]
    saida = a[1] if len(a) > 1 else None
    tmp = os.path.join(tempfile.mkdtemp(), os.path.basename(origem))
    shutil.copy2(origem, tmp)                              # le uma copia da copia: nada toca o arquivo dado
    wb = openpyxl.load_workbook(tmp, read_only=True, data_only=True, keep_vba=False)
    log = aba(wb, 'Audit_Log')
    lib = None
    for n in wb.sheetnames:
        if n.lower().startswith('libera') and 'store' not in n.lower():
            lib = wb[n]
    if log is None or lib is None:
        print('abas nao encontradas:', wb.sheetnames)
        return 2
    eventos = {}
    n_ev = 0
    for row in log.iter_rows(values_only=True):
        cels = [str(c) for c in row if c is not None]
        tipo = next((c for c in cels if c.startswith('ASSINATURA_')), None)
        if not tipo or tipo in ('ASSINATURA_RECUSADA', 'ASSINATURA_ANULADA'):
            continue
        for c in cels:
            m = RX_EV.search(c)
            if m:
                n_ev += 1
                eventos.setdefault(m.group(1), []).append({'tipo': tipo, 'celula': m.group(3) + m.group(4),
                                                           'run': m.group(5).strip()})
    # a aba: A = RUN; C:F = rubrica/texto das assinaturas
    achados, ok, sem_evento = [], 0, 0
    linhas = list(lib.iter_rows(min_row=4, max_row=203, min_col=1, max_col=6, values_only=True))
    for i, row in enumerate(linhas):
        run_linha = '' if row[0] is None else str(row[0]).strip().split('.')[0]
        for j in range(2, 6):
            v = row[j]
            if v is None:
                continue
            m = RX_QUANDO.search(str(v))
            if not m:
                continue
            evs = eventos.get(m.group(1), [])
            if not evs:
                sem_evento += 1
                continue
            runs_ev = {e['run'].split('.')[0] for e in evs}
            col = 'ABCDEF'[j]
            if run_linha in runs_ev:
                ok += 1
            else:
                achados.append({'celula': f'{col}{i + 4}', 'run_na_linha': run_linha, 'run_assinado': sorted(runs_ev),
                                'quando': m.group(1)})
    res = {'arquivo': os.path.basename(origem), 'eventos_assinatura_liberacao': n_ev, 'assinaturas_na_aba_conferidas': ok,
           'assinaturas_sem_evento_no_log': sem_evento, 'fora_do_run': achados}
    print(json.dumps(res, ensure_ascii=False, indent=1))
    if saida:
        json.dump(res, open(saida, 'w', encoding='utf-8'), ensure_ascii=False, indent=1)
    wb.close()
    return 1 if achados else 0


if __name__ == '__main__':
    sys.exit(main(sys.argv[1:]))
