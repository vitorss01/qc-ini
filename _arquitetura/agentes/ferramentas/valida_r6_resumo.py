# -*- coding: utf-8 -*-
"""valida_r6_resumo.py -- Tester/QA, ciclo 2026-11-15_reteste06: resume os JSON do valida_r6_adr096.py (nova x base).
Uso: python valida_r6_resumo.py <Bioquimica|Hematologia>   (le %TEMP%/qcw/auditoria/2026-11-15_reteste06/tester_validar/)
"""
import json
import os
import sys

P = os.path.join(os.environ['TEMP'], 'qcw', 'auditoria', '2026-11-15_reteste06', 'tester_validar')
PROD = sys.argv[1]


def ld(m):
    f = os.path.join(P, f'valida_r6_{PROD}_{m}.json')
    return json.load(open(f, encoding='utf-8')) if os.path.exists(f) else {}


N, B = ld('nova'), ld('base')
print('### origem', N.get('origem'), N.get('fim'), '| base', B.get('origem'), B.get('fim'))
print('### tblCQ_Final nova', N.get('tblCQ_Final'), 'base', B.get('tblCQ_Final'))
for k in [k for k in N if k.startswith('ERRO')]:
    print('!!!', k, N[k][:600])

print('\n### OC-0074 listas (nova | base)')
for k, v in N.get('OC0074_listas', {}).items():
    if k.startswith('_'):
        print(k, v)
        continue
    b = B.get('OC0074_listas', {}).get(k, {})
    if v.get('Type') != 3:
        print(f'{k}: Type {v.get("Type")} (nao e lista) msg={v.get("ErrorMessage")!r}')
        continue
    print(f'{k}: titulo={v.get("ErrorTitle")!r} msg={v.get("ErrorMessage")!r} show={v.get("ShowError")} style={v.get("AlertStyle")} '
          f'ok_msg={v.get("ok_msg")} ok_valores={v.get("ok_valores")} limpar={v.get("ok_limpar")}')
    print(f'    itens nova={len(v.get("itens") or [])} base={len(b.get("itens") or [])} iguais={v.get("itens") == b.get("itens")} '
          f'aceitos iguais={v.get("aceitos") == b.get("aceitos")} invalidos iguais={v.get("invalidos") == b.get("invalidos")} '
          f'refersto={v.get("refersto")!r}')
    nao = [a for a, x in (v.get('aceitos') or {}).items() if not x]
    num_ = [a for a, x in (v.get('aceitos') or {}).items() if x == 'numero']
    inv = [a for a, x in v.get('invalidos', {}).items() if x]
    if nao or inv or num_:
        print(f'    item recusado={nao} aceito_so_como_numero={num_[:5]}.. invalido_aceito={inv}')
    print(f'    base: titulo={b.get("ErrorTitle")!r} msg={b.get("ErrorMessage")!r}')

print('\n### OC-0021 formula')
f, fb = N.get('OC0021_formula', {}), B.get('OC0021_formula', {})
print('n_SUBSTITUTE nova', f.get('n_SUBSTITUTE'), 'base', fb.get('n_SUBSTITUTE'))
for s, v in f.get('literais', {}).items():
    print(v['ok'], '|', v['tela'], '| base:', fb.get('literais', {}).get(s, {}).get('ok'))
print('sobra sem acento (motivos da copia):', f.get('sobra_sem_acento'))
print('motivos distintos:', len(f.get('motivos_distintos_da_copia', {})))
for s, r in f.get('motivos_distintos_da_copia', {}).items():
    print('   ', s, '=>', r)

print('\n### OC-0075 Estatistica')
for k, v in N.get('OC0075_estatistica', {}).items():
    b = B.get('OC0075_estatistica', {}).get(k, {})
    c, cb = v['cel'], b.get('cel', {})
    same = {a: (c[a]['v2'] == cb.get(a, {}).get('v2'), c[a]['f'] == cb.get(a, {}).get('f')) for a in ('H5', 'I5', 'J5', 'B5', 'D5')}
    print(f'{k}: "{c["linha5"]}" | base "{cb.get("linha5")}" | H5 fmt {c["H5"]["fmt"]!r} J5 fmt {c["J5"]["fmt"]!r} | '
          f'v2/formula iguais {same} | grade igual {v["grade_hash"] == b.get("grade_hash")} | oraculo {v["grade_vs_oraculo"]} | 1900 {v["tem_1900"]}')

print('\n### OC-0076 corte')
for k, v in N.get('OC0076_corte', {}).items():
    b = B.get('OC0076_corte', {}).get(k, {})
    print(f'{k}: A39="{v["A39"]}"\n    base="{b.get("A39")}"\n    AF1={v["Eng_AF1"]} oraculo={v["oraculo_n_periodo"]} ok={v["AF1_igual_oraculo"]} '
          f'M1={v["Eng_M1"]} (base {b.get("Eng_M1")}) corridas={v["corridas_oraculo"]} cita_estat={v["cita_estatistica"]} todos={v["todos_os_lotes"]}')

if 'passos' in N:
    print('\n### OC-0073 ref', json.dumps(N['ref'], ensure_ascii=False)[:600])
    print('painel_antes', N['painel_antes']['engPainel'])
    for k, v in N['passos'].items():
        vv = dict(v) if isinstance(v, dict) else v
        if isinstance(vv, dict):
            vv.pop('linha', None)
        print(' ', k, json.dumps(vv, ensure_ascii=False)[:700])
    for rot in ('atualizar1', 'atualizar2'):
        R = N.get(rot)
        if not R:
            continue
        print(f'\n### {rot} ({R.get("segundos")} s) motor={R.get("motor")}')
        for k, l in R['tela'].items():
            print(f'   tela {k:4s} pl{l["_planilha"]} {l["ID_REGISTRO"]:14s} N{l["NIVEL"]} {l["HORA"]} {l["RESULTADO"]} STATUS={l["STATUS"]} '
                  f'RUN={l["RUN"]} DETALHE="{l["DETALHE"]}" USUARIO={l["USUARIO"]}')
        for r in R['final']:
            print('   final', r['ID_REGISTRO'], r['STATUS_ANALITICO'], r['PARTICIPA_ESTATISTICA'], r['RUN'], r['NIVEL'], r['RESULTADO'],
                  '|', r['MOTIVO_EXCLUSAO_AUTOMATICA'])
        for q in R['qa']:
            print('   qa', json.dumps(q, ensure_ascii=False)[:500])
        print('   ResumoQA:', R['ResumoQA'].replace('\r\n', ' / '))
        print('   mensagem:', R['mensagem_atualizar'].replace('\r\n', ' / ')[:1500])
        print('   painel', R['painel']['filtro'], R['painel']['engPainel'])
        print('   painel x oraculo', R['painel_vs_oraculo'])
        print('   painel A7', [l[:10] + l[12:20] for l in R['painel']['A7']])
        print('   estatistica', json.dumps(R['estatistica'], ensure_ascii=False)[:900])
        print('   eng', json.dumps(R['eng'], ensure_ascii=False)[:1800])
    if 'reserva_pq' in N:
        q = N['reserva_pq']
        print('\n### reserva PQ', q['segundos'], 's; linhas', q['tblCQ_Final_linhas'])
        print('   diferencas VBA x PQ:', q['diferencas_vba_pq'])
        print('   A01 igual:', q['A01_igual'], '\n   VBA:', q['A01_vba'], '\n   PQ :', q['A01_pq'])
        print('   retorno:', q['retorno'].replace('\r\n', ' / '))
    print('\n### correcao', json.dumps(N.get('correcao'), ensure_ascii=False)[:400])
    print('### avisos (todos)', len(N.get('avisos_todos', [])))
    for a in N.get('avisos_todos', []):
        print('   ', a)
