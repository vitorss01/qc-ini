# -*- coding: utf-8 -*-
"""valida_r6_complemento.py -- Tester/QA, ciclo 2026-11-15_reteste06 (VALIDAR ADR-096), complemento da OC-0074.

O valida_r6_adr096.py mostrou, na copia nova e na base (igual), que:
  * na Bio o filtro de lote (Registros!B4) recusa todo lote da propria lista pela Validation.Value;
  * os filtros de equipamento (G4) e status (J4) tem a lista vazia (lstRegEquip/lstRegStatus = 1 celula vazia) e
    aceitam qualquer texto, entao a mensagem nova nunca aparece.
Este script so LE a celula e a lista e testa Validation.Value com o valor escrito de varias formas, para achar a causa.
Uso: python -u valida_r6_complemento.py <Bioquimica|Hematologia> <saida.json>
Regras: copia nova em %TEMP%/qcw/auditoria/2026-11-15_reteste06/tester_validar/; nada salvo; fecha so o PID do QA.
"""
import hashlib
import json
import os
import shutil
import sys
import time

AQUI = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, AQUI)
import repro_comum as C  # noqa: E402
sys.path.insert(0, C.TESTES)
import qa_final  # noqa: E402

PROD, SAIDA = sys.argv[1], sys.argv[2]
PASTA = os.path.join(os.environ['TEMP'], 'qcw', 'auditoria', '2026-11-15_reteste06', 'tester_validar')


def sha(p):
    h = hashlib.sha256()
    with open(p, 'rb') as f:
        for b in iter(lambda: f.read(1 << 20), b''):
            h.update(b)
    return h.hexdigest()


def main():
    src = os.path.join(C.REPO, f'QC_{PROD}.xlsm')
    dst = os.path.join(PASTA, f'{PROD}_compl_{time.strftime("%H%M%S")}.xlsm')
    h0 = sha(src)
    shutil.copy2(src, dst)
    R = {'origem': {'sha256_antes': h0[:16], 'copia': os.path.basename(dst)}}
    qa = qa_final.QA(PROD, dst)
    R['excel_pid'] = qa.ex.pid
    xl = qa.ex.xl
    try:
        ws = qa.wb.Worksheets('Registros')
        try:
            ws.Unprotect(C.senha_planilha())
        except Exception:                                            # noqa: BLE001
            pass
        out = {}
        for addr, col in (('B4', 27), ('G4', 29), ('J4', 30)):
            c = ws.Range(addr)
            d = {'NumberFormat': str(c.NumberFormat), 'MergeCells': bool(c.MergeCells),
                 'MergeArea': str(c.MergeArea.Address), 'Formula1': str(c.Validation.Formula1),
                 'IgnoreBlank': bool(c.Validation.IgnoreBlank), 'InCellDropdown': bool(c.Validation.InCellDropdown),
                 'ErrorMessage': str(c.Validation.ErrorMessage)}
            lst = []
            for r in range(2, 8):
                k = ws.Cells(r, col)
                lst.append({'cel': k.Address, 'Value2': k.Value2, 'tipo': type(k.Value2).__name__,
                            'Formula': str(k.Formula), 'Prefix': str(k.PrefixCharacter), 'NumberFormat': str(k.NumberFormat)})
            d['lista_AA'] = lst
            nome = str(c.Validation.Formula1).lstrip('=')
            try:
                d['RefersTo'] = str(qa.wb.Names(nome).RefersTo)
            except Exception as e:                                   # noqa: BLE001
                d['RefersTo'] = f'erro {e}'
            first = next((x['Value2'] for x in lst if x['Value2'] not in (None, '')), None)
            prov = {}
            antes = c.Formula
            ev = xl.EnableEvents
            xl.EnableEvents = False
            try:
                cands = [('texto_lista', first), ('numero', int(first) if isinstance(first, str) and first.isdigit() else None),
                         ('apostrofo', "'" + str(first) if first is not None else None),
                         ('lixo', 'TESTE-AUDITORIA-R6V')]
                for rot, v in cands:
                    if v is None:
                        continue
                    c.Value = v
                    prov[rot] = {'escrito': v, 'Value2': c.Value2, 'tipo': type(c.Value2).__name__,
                                 'Validation.Value': bool(c.Validation.Value)}
            finally:
                c.Formula = antes
                xl.EnableEvents = ev
            d['provas'] = prov
            out[addr] = d
        R['registros'] = out
        lo = ws.ListObjects('tblRegistros')
        R['tblRegistros'] = {'linhas': lo.ListRows.Count,
                             'Equipamento': [x[0] for x in lo.ListColumns('Equipamento').DataBodyRange.Value2],
                             'Status': [x[0] for x in lo.ListColumns('Status').DataBodyRange.Value2]}
        # linha nova da tabela (como a assinatura de uma calibracao): a validacao e a mensagem acompanham?
        ev = xl.EnableEvents
        xl.EnableEvents = False
        try:
            nova = lo.ListRows.Add()
            lin = {}
            for idx in (4, 6, 10):
                c = nova.Range.Cells(1, idx)
                try:
                    lin[str(lo.ListColumns(idx).Name)] = {'cel': c.Address, 'Formula1': str(c.Validation.Formula1),
                                                          'ErrorTitle': str(c.Validation.ErrorTitle),
                                                          'ErrorMessage': str(c.Validation.ErrorMessage)}
                    c.Value = 'TESTE-AUDITORIA-R6V'
                    lin[str(lo.ListColumns(idx).Name)]['lixo_aceito'] = bool(c.Validation.Value)
                    c.ClearContents()
                except Exception as e:                               # noqa: BLE001
                    lin[str(lo.ListColumns(idx).Name)] = f'sem validacao: {e}'
            R['linha_nova_tblRegistros'] = lin
        finally:
            xl.EnableEvents = ev
    finally:
        qa.fechar()
        h1 = sha(src)
        R['fim'] = {'sha256_depois': h1[:16], 'producao_intacta': h1 == h0}
        try:
            os.remove(dst)
            R['copia_apagada'] = True
        except Exception as e:                                       # noqa: BLE001
            R['copia_apagada'] = str(e)
        json.dump(R, open(SAIDA, 'w', encoding='utf-8'), ensure_ascii=False, indent=1, default=str)
        print(json.dumps(R, ensure_ascii=False, default=str)[:4000])


if __name__ == '__main__':
    main()
