# -*- coding: utf-8 -*-
"""Complemento da F4 Hema (ciclo 2026-10-31_ciclo01): a aba Estatistica acompanha o lote/ano quando o usuario troca H3,
entra na aba pelo botao de navegacao (mApp.IrEstatistica) e muda Ano De/Ate? Lista tambem os modulos VBA de planilha
(linhas de codigo), para saber se a Estatistica tem evento proprio. Uso: python -u f4_hema_conf.py <copia.xlsm> <raiz_repo> <saida.json>.
Nada e salvo na copia. Autor: AGENTE_USUARIO_SIMULADO."""
import sys, os, json, datetime as dt
REPO = sys.argv[2]
sys.path.insert(0, os.path.join(REPO, '_arquitetura', 'integracao_dbseac', 'testes'))
sys.path.insert(0, os.path.join(REPO, '_arquitetura', 'agentes', 'ferramentas'))
from qa_final import QA
q = QA('Hematologia', sys.argv[1])
out = {'pid': q.ex.pid}
try:
    p = q.ws('Painel'); est = q.ws('Estatística')
    q.run('mSeguranca.ReprotectAll', teto=120)
    def wbc():
        v = est.Range('A14:F133').Value
        return {'B4': est.Range('B4').Value, 'B3': est.Range('B3').Value, 'D3': est.Range('D3').Value,
                'WBC': [list(r[2:6]) for r in v if r[0] == 'WBC'],
                'linhas_nome_estranho': [[i + 14, r[0], r[1], r[2]] for i, r in enumerate(v) if r[0] not in (None, '') and not isinstance(r[0], str) or r[0] in ('0',)]}
    out['antes'] = wbc()
    q.ex.xl.EnableEvents = True
    p.Range('H3').Value = "'619611"; q.ex.esperar()
    out['apos_H3'] = wbc()
    q.run('mApp.IrEstatistica', teto=120); q.ex.esperar()
    out['apos_IrEstatistica'] = wbc()
    est.Range('B3').Value = 2026; est.Range('D3').Value = 2026; q.ex.esperar(); q.ex.xl.Calculate()
    out['apos_B3D3'] = wbc()
    q.run('mApp.IrPainel', teto=120); q.ex.esperar(); q.run('mApp.IrEstatistica', teto=120); q.ex.esperar()
    out['apos_ir_e_voltar'] = wbc()
    a = est.Range('A14:B133').Value
    out['nomes_A'] = [[i + 14, a[i][0], a[i][1]] for i in range(len(a)) if a[i][0] not in (None, '') and (not isinstance(a[i][0], str) or a[i][0].strip() in ('0', ''))]
    try:
        comps = []
        for c in q.wb.VBProject.VBComponents:
            n = c.CodeModule.CountOfLines
            comps.append([c.Name, c.Type, n, c.Properties('Name').Value if c.Type == 100 else ''])
        out['vbcomp'] = comps
    except Exception as e:
        out['vbcomp'] = 'sem acesso: ' + repr(e)[:120]
finally:
    q.fechar(False)
json.dump(out, open(sys.argv[3], 'w', encoding='utf-8'), ensure_ascii=False, indent=1, default=str)
print('ok')
