# -*- coding: utf-8 -*-
"""reteste05_usuario_f1.py -- Usuario Simulado, ciclo 2026-11-15_reteste05, frente F1, modo RETESTE (Bioquimica).

OC-0071 (R4-02 do reteste 04): Estatistica da Bio com o Painel em "Tudo (sem filtro)": B5/D5 mostravam "00/01/1900"
e o D4 "Ano 2026". Correcao declarada: ADR-095 par.3 (instalar_adr092: formato dd/mm/aaaa;; no B5/D5 e D4 com a
condicao SEM_DEATE; entrega 42). Criterio de aceite: em "Tudo" B5/D5/D4 vazios; com periodo, datas dd/mm/aaaa e ano
corretos; numeros da Estatistica identicos aos de antes da entrega (so exibicao).

Metodo: a MESMA sequencia de filtros do Painel (celulas L4, C3, H3, K3, K4, G3/G4, eventos ligados, como o usuario),
seguida de abrir a aba Estatistica (Activate), em DUAS copias novas, uma de cada vez (um Excel por vez):
  - BASE  = copia do backup da producao ANTERIOR a entrega 42 (_backup_pre_integracao_2026-11-15_0011, entrega 41);
  - NOVA  = copia da producao atual (raiz).
Em cada cenario le, da Estatistica: Formula/Text/Value2/NumberFormat do cabecalho A4:J5, o Text das linhas 1-12
(procura "1900", "yyyy", "Ano 0") e o Value2 de A1:AU200 (a tabela inteira); do Painel: G3/G4/K3/K4 e engPainel.
A comparacao BASE x NOVA e celula a celula, na precisao total do Value2 (tolerancia zero: o mesmo motor sobre os
mesmos dados). Confere tambem que a tblCQ_Final das duas copias e a mesma (n de linhas + hash do Value2).

Cenarios: S1 original (DIMENSION 1, Glicose, lote 8974, "Tudo"); S2 original 2a parte (K3 2026, K4 Set); vizinhos:
S3 trimestre T3/2026; S4 "Ano inteiro" 2026; S5 H3 vazio + "Tudo" (todos os lotes, ADR-086); S6 H3 vazio + Out/2026;
S7 De/Ate digitado atravessando o ano (01/12/2025-31/01/2026); S8 so o De digitado (01/09/2026, Ate vazio);
S9 volta a "Tudo" depois do personalizado (o cabecalho tem de limpar de novo).

Regras: so COPIAS em %TEMP%/qcw/auditoria/2026-11-15_reteste05/; nada e salvo; fechar() encerra so o PID do QA;
nenhuma captura de tela; nenhum dado de teste e gravado (so filtros).
Uso: python -u reteste05_usuario_f1.py  -> %TEMP%/qcw/auditoria/2026-11-15_reteste05/Bioquimica_F1.json
"""
import datetime as dt
import hashlib
import json
import os
import re
import shutil
import sys
import time
import traceback

AQUI = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, AQUI)
import repro_comum as C  # noqa: E402

PROD = 'Bioquimica'
CICLO = '2026-11-15_reteste05'
PASTA = os.path.join(os.environ['TEMP'], 'qcw', 'auditoria', CICLO)
C.PASTA = PASTA
BACKUP = os.path.join(C.REPO, '_backup_pre_integracao_2026-11-15_0011', f'QC_{PROD}.xlsm')
HDR = [f'{c}{r}' for r in (4, 5) for c in 'ABCDEFGHIJ']


def t(v):
    return '' if v is None else str(v).strip()


def cenarios():
    D = dt.datetime
    return [
        ('S1_original_Tudo', [('L4', 'DIMENSION 1'), ('C3', 'Glicose'), ('H3', "'8974"), ('K4', 'Tudo (sem filtro)')]),
        ('S2_original_Set2026', [('K3', 2026), ('K4', 'Set')]),
        ('S3_T3_2026', [('K4', 'T3')]),
        ('S4_Ano_inteiro_2026', [('K4', 'Ano inteiro')]),
        ('S5_H3_vazio_Tudo', [('H3', None), ('K4', 'Tudo (sem filtro)')]),
        ('S6_H3_vazio_Out2026', [('K3', 2026), ('K4', 'Out')]),
        ('S7_DeAte_atravessa_ano', [('H3', "'8974"), ('G3', D(2025, 12, 1)), ('G4', D(2026, 1, 31))]),
        ('S8_so_De', [('K3', 2026), ('K4', 'Set'), ('G4', None), ('G3', D(2026, 9, 1))]),
        ('S9_volta_Tudo', [('K4', 'Tudo (sem filtro)')]),
    ]


def aplicar(qa, p, passos):
    for c, v in passos:
        if v is None:
            p.Range(c).MergeArea.ClearContents()
            qa.ex.esperar()
            time.sleep(0.5)
        else:
            C.por(qa, p, c, v, local=False)
    time.sleep(1.5)
    qa.ex.esperar()


def ler(qa, p, e):
    e.Activate()                               # o usuario abre a aba Estatistica
    qa.ex.esperar()
    time.sleep(1)
    hdr = {}
    for a in HDR:
        c = e.Range(a)
        hdr[a] = {'f': t(c.Formula)[:140], 'txt': t(c.Text)[:120], 'v2': c.Value2, 'fmt': t(c.NumberFormat)}
    txt12 = ' '.join(t(e.Cells(r, k).Text) for r in range(1, 13) for k in range(1, 48))
    v = e.Range('A1:AU200').Value2
    tab = [list(l) for l in v]
    pn = {c: str(C.limpo(p.Range(c).Value)) for c in ('G3', 'G4', 'K3', 'K4', 'H3', 'C3', 'L4')}
    pn['G3_txt'], pn['G4_txt'] = t(p.Range('G3').Text), t(p.Range('G4').Text)
    try:
        eng = [list(x) for x in qa.nome('engPainel').Value2]
    except Exception as ex:  # noqa: BLE001
        eng = repr(ex)
    p.Activate()
    qa.ex.esperar()
    return {'hdr': hdr, 'tem_1900': '1900' in txt12, 'tem_yyyy': bool(re.search(r'y{2,4}', txt12)),
            'tem_Ano_0': bool(re.search(r'\bAno 0\b', txt12)), 'painel': pn, 'engPainel': eng, 'tab': tab}


def rodar(rotulo, origem):
    destino = os.path.join(PASTA, f'{PROD}_F1_{rotulo}.xlsm')
    if os.path.exists(destino):
        os.remove(destino)
    h0 = C.sha(origem)
    shutil.copy2(origem, destino)
    R = {'origem': os.path.relpath(origem, C.REPO), 'sha_origem_12': h0, 'copia': os.path.basename(destino),
         'cen': {}}
    senha = C.senha_planilha()
    qa = C.qa_final.QA(PROD, destino)
    R['excel_pid'] = qa.ex.pid
    try:
        p, e = qa.ws('Painel'), qa.ws('Estatística')
        for w in (p, e):
            try:
                w.Unprotect(senha)
            except Exception:
                pass
        p.Activate()
        qa.ex.esperar()
        lo = qa.lo('tblCQ_Final')
        dv = lo.DataBodyRange.Value2
        R['tblCQ_Final'] = {'linhas': len(dv), 'hash': hashlib.sha256(repr(dv).encode()).hexdigest()[:16]}
        for nome, passos in cenarios():
            print('>>', rotulo, nome, time.strftime('%H:%M:%S'), flush=True)
            try:
                aplicar(qa, p, passos)
                R['cen'][nome] = ler(qa, p, e)
            except Exception as ex:  # noqa: BLE001
                R['cen'][nome] = {'erro': repr(ex), 'trace': traceback.format_exc()[-1500:]}
    except Exception as ex:  # noqa: BLE001
        R['ERRO_FATAL'] = {'erro': repr(ex), 'trace': traceback.format_exc()[-3000:]}
    finally:
        try:
            qa.fechar(salvar=False)
        except Exception:
            pass
    R['sha_origem_depois_12'] = C.sha(origem)
    return R


def diff_tab(a, b):
    out = []
    for i, (la, lb) in enumerate(zip(a, b), 1):
        for j, (x, y) in enumerate(zip(la, lb)):
            if x != y and not (isinstance(x, float) and isinstance(y, float) and x == y):
                col = ''
                k = j + 1
                while k:
                    k, r = divmod(k - 1, 26)
                    col = chr(65 + r) + col
                out.append([f'{col}{i}', x, y])
    return out


def main():
    os.makedirs(PASTA, exist_ok=True)
    prod = os.path.join(C.REPO, f'QC_{PROD}.xlsm')
    L = {'inicio': dt.datetime.now().isoformat(timespec='seconds')}
    L['BASE'] = rodar('base41', BACKUP)          # um Excel por vez: BASE fecha antes de a NOVA abrir
    L['NOVA'] = rodar('nova', prod)
    comp = {}
    for nome, _ in cenarios():
        a, b = L['BASE']['cen'].get(nome, {}), L['NOVA']['cen'].get(nome, {})
        if 'tab' not in a or 'tab' not in b:
            comp[nome] = 'faltou leitura'
            continue
        d = diff_tab(a['tab'], b['tab'])
        comp[nome] = {'n_celulas_diferentes': len(d), 'diferencas': d[:40],
                      'engPainel_igual': a['engPainel'] == b['engPainel'],
                      'painel_igual': a['painel'] == b['painel']}
    L['comparacao'] = comp
    for k in ('BASE', 'NOVA'):
        for c in L[k]['cen'].values():
            c.pop('tab', None)
    L['fim'] = dt.datetime.now().isoformat(timespec='seconds')
    json.dump(L, open(os.path.join(PASTA, f'{PROD}_F1.json'), 'w', encoding='utf-8'), ensure_ascii=False, indent=1,
              default=str)
    print('FIM', flush=True)


if __name__ == '__main__':
    t0 = time.time()
    main()
    print(f'tempo {time.time() - t0:.0f}s', flush=True)
