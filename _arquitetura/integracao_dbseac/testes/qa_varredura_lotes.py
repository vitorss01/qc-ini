# -*- coding: utf-8 -*-
"""QA do ADR-077 -- troca TODOS os lotes do cadastro no Painel, como o usuario, e exige que cada lote com resultados
abra o grafico com pontos, a linha da media e as linhas de limite, com a media e o DP DAQUELE lote. Roda numa COPIA
instalada; nada e salvo.

Uso: python qa_varredura_lotes.py <Bioquimica|Hematologia> <copia_instalada.xlsm> [saida.json]

Pedido do usuario (15/10/2026): "troque de lote para ver todos eles se estao mostrando os limites e medias nos
graficos". A varredura anterior (antes do ADR-077) achou 15 de 21 lotes com o grafico VAZIO: filtro de ano 2026 para
lote de 2025, DIMENSION 1 para lote que so rodou no DIMENSION 2, analito que o lote nao tem.

  V01 todo lote com resultados: grafico com pontos e linhas de limite (no nivel que tem ponto); nenhum "SEM MEDIA/DP"
      nesse nivel; media e DP do grafico = os do lote (LotesStore)
  V02 o lote em uso (o 1o) com dado no filtro atual NAO teve o filtro mexido (periodo e equipamento como estavam)
  V03 lote sem nenhum resultado na base: o Painel diz "SEM MEDIA/DP" e nao inventa linha
  V04 seguranca do teste: arquivo de entrada intacto
"""
import json
import os
import sys

AQUI = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, AQUI)
from qa_final import QA, reg, RES, sha256_arquivo  # noqa: E402


def num(v):
    return isinstance(v, (int, float)) and not isinstance(v, bool) and not (isinstance(v, int) and v < -2146820000)


def perto(a, b, tol=1e-9):
    return num(a) and num(b) and abs(a - b) <= tol * max(1.0, abs(b))


def txt(v):
    if isinstance(v, float) and v == int(v):
        return str(int(v))
    return str(v if v is not None else '').strip()


def executar(produto, caminho, saida):
    hash0 = sha256_arquivo(caminho)
    q = QA(produto, caminho)
    linhas_tab = []
    try:
        q.run('mSeguranca.ReprotectAll')
        bio = produto.startswith('Bio')
        cel_md = ('AX1', 'AY1', 'AZ1', 'BA1') if bio else ('BT1', 'BU1', 'BV1', 'BW1')
        p, c = q.ws('Painel'), q.ws('Calc')
        cab = [str(v or '').strip() for v in c.Range('A2:CZ2').Value[0]]
        col = {k: cab.index(k) + 1 for k in ('N1 line', 'N1 m3', 'N2 line', 'N2 m3') if k in cab}
        rows = q.ler('tblCQ_Final')
        com_dado = {}
        for r in rows:
            if str(r.get('PARTICIPA_ESTATISTICA', '')).upper() == 'SIM':
                com_dado[txt(r['LOTE'])] = com_dado.get(txt(r['LOTE']), 0) + 1
        cad = []
        for v in q.nome('regLoteCol').Value:
            s = txt(v[0])
            if s and s not in cad:
                cad.append(s)
        lote0 = txt(q.nome('loteParam').Value)
        filtro0 = (txt(p.Range('G3').Value), txt(p.Range('G4').Value), txt(p.Range('L4').Value) if bio else '')
        ruins, sem_dado_ruins, lote0_ok = [], [], None
        for lote in cad + [lote0]:                          # o 1o de novo no fim: volta ao lote em uso
            p.Range('H3').Value = "'" + lote                # como o usuario: eventos ligados
            q.ex.esperar()
            q.ex.xl.Calculate()
            an = txt(q.nome('selAnalito').Value)
            md = [c.Range(x).Value for x in cel_md]
            cont = {}
            for k, cc in col.items():
                vals = c.Range(c.Cells(3, cc), c.Cells(182, cc)).Value
                cont[k] = sum(1 for v in vals if num(v[0]))
            st = [txt(p.Range(f'J{7 + i}').Value) for i in range(2)]
            exp = [q.run('mLotes.MediaDoLote', lote, an, 1), q.run('mLotes.DPDoLote', lote, an, 1),
                   q.run('mLotes.MediaDoLote', lote, an, 2), q.run('mLotes.DPDoLote', lote, an, 2)]
            filtro = (txt(p.Range('G3').Value), txt(p.Range('G4').Value), txt(p.Range('L4').Value) if bio else '')
            linha = {'lote': lote, 'resultados': com_dado.get(lote, 0), 'analito': an, 'filtro_de_ate_equip': filtro,
                     'pontos_N1': cont.get('N1 line'), 'limites_N1': cont.get('N1 m3'),
                     'pontos_N2': cont.get('N2 line'), 'limites_N2': cont.get('N2 m3'),
                     'media_dp_grafico': md, 'media_dp_lote': exp, 'status': st,
                     'aviso': str(q.ex.xl.StatusBar or '')[:200]}
            linhas_tab.append(linha)
            print(f"{lote:>8} n={linha['resultados']:5} {an[:20]:20} N1 {linha['pontos_N1']}/{linha['limites_N1']} "
                  f"N2 {linha['pontos_N2']}/{linha['limites_N2']} filtro={filtro} st={st}", flush=True)
            if com_dado.get(lote, 0) == 0:
                if not all('SEM M' in s.upper() for s in st) or any(num(v) for v in md):
                    sem_dado_ruins.append(linha)
                continue
            pts = (cont.get('N1 line', 0), cont.get('N2 line', 0))
            lim = (cont.get('N1 m3', 0), cont.get('N2 m3', 0))
            erro = []
            if sum(pts) == 0:
                erro.append('grafico sem pontos')
            for nv in (0, 1):
                if pts[nv] > 0:
                    if lim[nv] == 0:
                        erro.append(f'N{nv + 1} com pontos e sem linhas de limite')
                    if 'SEM M' in st[nv].upper():
                        erro.append(f'N{nv + 1} SEM MEDIA/DP')
                    if not (perto(md[2 * nv], exp[2 * nv]) and perto(md[2 * nv + 1], exp[2 * nv + 1])):
                        erro.append(f'N{nv + 1} media/DP do grafico != do lote')
            if erro:
                ruins.append((lote, an, erro))
            if lote == lote0 and lote0_ok is None:
                lote0_ok = (filtro == filtro0)
        reg('V01 Todos os lotes com resultados, trocados no Painel como o usuário faz: o gráfico abre com pontos, a linha '
            'da média e as linhas de limite (-3DP..+3DP) no nível que tem ponto, sem "SEM MÉDIA/DP", e a média/DP do '
            'gráfico são as do lote escolhido',
            not ruins and sum(1 for l in linhas_tab if l['resultados'] > 0) > 0,
            {'lotes_com_resultado': sum(1 for l in cad if com_dado.get(l, 0) > 0), 'lotes_com_problema': ruins[:8]})
        reg('V02 O lote em uso, que tem dado no filtro atual, não teve o filtro mexido (período e equipamento como estavam)',
            lote0_ok is True, {'lote': lote0, 'filtro_antes': filtro0})
        reg('V03 Lote sem nenhum resultado na base: o Painel diz "SEM MÉDIA/DP" e não desenha linha', not sem_dado_ruins,
            {'lotes_sem_resultado': [l for l in cad if com_dado.get(l, 0) == 0], 'problemas': sem_dado_ruins[:3]})
    finally:
        q.fechar(salvar=False)
    hash1 = sha256_arquivo(caminho)
    reg('V04 Segurança do teste: arquivo de entrada intacto (SHA-256 igual antes e depois; nada salvo nele)',
        hash0 == hash1, {'arquivo': caminho, 'sha256_antes': hash0, 'sha256_depois': hash1})
    if saida:
        with open(saida, 'w', encoding='utf-8') as f:
            json.dump({'resultados': RES, 'lotes': linhas_tab}, f, ensure_ascii=False, indent=1, default=str)
    falhas = [r for r in RES if r['resultado'] == 'FAIL']
    print(f'\n=== {produto} varredura de lotes (ADR-077): {len(RES) - len(falhas)} PASS / {len(falhas)} FAIL ===', flush=True)
    return falhas


if __name__ == '__main__':
    falhas = executar(sys.argv[1], os.path.abspath(sys.argv[2]), sys.argv[3] if len(sys.argv) > 3 else None)
    sys.exit(1 if falhas else 0)
