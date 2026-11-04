# -*- coding: utf-8 -*-
# Reteste01 (usuario): Painel com a aba ATIVA (como o usuario), OC-0002 (HbA1c SEM EP) e OC-0013 (H3 vazio).
import json, os, shutil, sys, time
F = r'C:\Users\vitor.santos\OneDrive - MSFT\INI_DB_AJUSTES\OneDrive - MSFT\Desktop\QC_INI_PRONTO\qc-ini\_arquitetura\agentes\ferramentas'
sys.path.insert(0, F)
import f1_entrada_ciq as f1
PROD = sys.argv[1]
f1.P.update(f1.PRODUTOS[PROD])
W = os.path.join(os.environ['TEMP'], 'qcw', 'auditoria', '2026-11-04_reteste01', 'usuario', PROD)
cp = os.path.join(W, f'{PROD}_painel2.xlsm')
shutil.copy2(os.path.join(W, f'{PROD}_prep.xlsm'), cp)
q = f1.QA(PROD, cp)
print('EXCEL_PID', q.ex.pid, flush=True)
out = {}
try:
    p = q.ws('Painel')
    p.Activate(); q.ex.esperar()
    try: p.Unprotect(f1.SENHA)
    except Exception: pass
    q.ex.xl.EnableEvents = True
    nl = 2 if PROD == 'Bioquimica' else 3
    def tela(rot):
        q.ex.esperar()
        out[rot] = {c: p.Range(c).Value for c in ('C3', 'H3', 'K3', 'K4', 'G6', 'H6', 'I6', 'H4')}
        out[rot]['linhas'] = [list(x)[:10] for x in p.Range('A7:J%d' % (6 + nl)).Value]
        out[rot]['A39'] = p.Range('A39').Value
    seq = ([('L4', 'DIMENSION 2'), ('C3', 'Hemoglobina glicada'), ('H3', "'8590"), ('K3', 2026), ('K4', 'Set')] + ([('G3', 46266), ('G4', 46295)] if os.environ.get('G34') else [])
           if PROD == 'Bioquimica' else [('C3', 'WBC'), ('H3', "'602911"), ('K3', 2026), ('K4', 'Mar')])
    for c, v in seq:
        p.Range(c).Value = v; q.ex.esperar(); out.setdefault('eventos_apos', []).append((c, bool(q.ex.xl.EnableEvents)))
    tela('lote')
    time.sleep(3)
    p.Range('H3').MergeArea.ClearContents(); q.ex.esperar(); time.sleep(3)
    tela('h3_vazio')
finally:
    q.fechar(salvar=False)
json.dump(f1.j(out), open(os.path.join(W, 'painel2.json'), 'w', encoding='utf-8'), ensure_ascii=False, indent=1)
for k, v in out.items():
    print(k, json.dumps(f1.j(v), ensure_ascii=False)[:900])



