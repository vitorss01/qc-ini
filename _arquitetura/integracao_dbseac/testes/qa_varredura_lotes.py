# -*- coding: utf-8 -*-
"""QA do ADR-077 -- troca TODOS os lotes do cadastro no Painel, como o usuario, e exige que cada lote com resultados
abra o grafico com pontos, a linha da media e as linhas de limite, com a media e o DP DAQUELE lote. Roda numa COPIA
instalada; nada e salvo.

Uso: python qa_varredura_lotes.py <Bioquimica|Hematologia> <copia_instalada.xlsm> [saida.json]

Pedido do usuario (15/10/2026): "troque de lote para ver todos eles se estao mostrando os limites e medias nos
graficos". A varredura anterior (antes do ADR-077) achou 15 de 21 lotes com o grafico VAZIO: filtro de ano 2026 para
lote de 2025, DIMENSION 1 para lote que so rodou no DIMENSION 2, analito que o lote nao tem.

  V01 todo lote com resultados: grafico com pontos e, em todo nivel com Media/DP, as 7 linhas do alvo com 2 pontos
      cada NA SERIE DO GRAFICO (ADR-078: x = bordas do eixo); nenhum "SEM MEDIA/DP" no nivel que tem ponto; media e
      DP do grafico = os do lote (LotesStore); N1..N3 na Hematologia
  V02 o lote em uso (o 1o) com dado no filtro atual NAO teve o filtro mexido (periodo e equipamento como estavam)
  V03 lote sem nenhum resultado na base: nivel sem Media/DP -> "SEM MEDIA/DP" e nenhuma linha; nivel com Media/DP ->
      "AGUARDANDO 1a CORRIDA" e as 7 linhas do alvo (ADR-078)
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
        nlv = q.nlv
        cel_md = ('AX1', 'AY1', 'AZ1', 'BA1') if bio else ('BT1', 'BU1', 'BV1', 'BW1', 'BX1', 'BY1')
        LIM = ['-3s', '-2s', '-1s', 'Média', '+1s', '+2s', '+3s']
        p, c = q.ws('Painel'), q.ws('Calc')
        cab = [str(v or '').strip() for v in c.Range('A2:CZ2').Value[0]]
        col = {f'N{i} line': cab.index(f'N{i} line') + 1 for i in range(1, nlv + 1)}

        def pares_linhas(i):
            ch = p.ChartObjects(i).Chart
            out = []
            for s_i in range(1, ch.SeriesCollection().Count + 1):
                s = ch.SeriesCollection(s_i)
                if str(s.Name) in LIM:
                    out.append(sum(1 for x, y in zip(s.XValues, s.Values) if num(x) and num(y)))
            return out
        rows = q.ler('tblCQ_Final')
        com_dado, na_base = {}, {}
        for r in rows:                                     # V03: a mesma contagem do Painel (engNLoteBase: toda linha)
            na_base[txt(r['LOTE'])] = na_base.get(txt(r['LOTE']), 0) + 1
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
        # 21/10/2026: o lote em uso e conferido PRIMEIRO (V02) -- antes so na 1a vez que aparecia no cadastro; se ele nao
        # era o 1o do cadastro (producao com loteParam 8977, 17o), os lotes anteriores ja tinham ajustado o filtro (ADR-077)
        for k_lote, lote in enumerate([lote0] + cad + [lote0]):   # o em uso no inicio e de novo no fim
            p.Range('H3').Value = "'" + lote                # como o usuario: eventos ligados
            q.ex.esperar()
            q.ex.xl.Calculate()
            an = txt(q.nome('selAnalito').Value)
            md = [c.Range(x).Value for x in cel_md]
            cont = {}
            for k, cc in col.items():
                vals = c.Range(c.Cells(3, cc), c.Cells(182, cc)).Value
                cont[k] = sum(1 for v in vals if num(v[0]))
            st = [txt(p.Range(f'J{7 + i}').Value) for i in range(nlv)]
            exp = []
            for i in range(1, nlv + 1):
                exp += [q.run('mLotes.MediaDoLote', lote, an, i), q.run('mLotes.DPDoLote', lote, an, i)]
            lin = [pares_linhas(i) for i in range(1, nlv + 1)]
            filtro = (txt(p.Range('G3').Value), txt(p.Range('G4').Value), txt(p.Range('L4').Value) if bio else '')
            linha = {'lote': lote, 'resultados': com_dado.get(lote, 0), 'analito': an, 'filtro_de_ate_equip': filtro,
                     'pontos': [cont.get(f'N{i} line') for i in range(1, nlv + 1)], 'pares_linhas': lin,
                     'media_dp_grafico': md, 'media_dp_lote': exp, 'status': st,
                     'aviso': str(q.run('mLotes.AvisoPainel', teto=60) or '')[:200]}
            linhas_tab.append(linha)
            print(f"{lote:>8} n={linha['resultados']:5} {an[:20]:20} pontos {linha['pontos']} linhas "
                  f"{[sum(x) for x in lin]} filtro={filtro} st={st}", flush=True)
            tem_alvo = [num(md[2 * nv]) and num(md[2 * nv + 1]) for nv in range(nlv)]
            linhas_ok = [(all(n == 2 for n in lin[nv]) and len(lin[nv]) == 7) if tem_alvo[nv] else sum(lin[nv]) == 0
                         for nv in range(nlv)]
            if na_base.get(lote, 0) == 0:
                for nv in range(nlv):
                    esperado = 'AGUARDANDO 1ª CORRIDA' if tem_alvo[nv] else 'SEM MÉDIA/DP'
                    if not st[nv].startswith(esperado) or not linhas_ok[nv]:
                        sem_dado_ruins.append(linha)
                        break
                continue
            if com_dado.get(lote, 0) == 0:                 # so resultados que nao participam: nada a exigir no grafico
                continue
            pts = [cont.get(f'N{i} line', 0) for i in range(1, nlv + 1)]
            erro = []
            if sum(pts) == 0:
                erro.append('grafico sem pontos')
            for nv in range(nlv):
                if not linhas_ok[nv]:
                    erro.append(f'N{nv + 1} linhas do alvo {lin[nv]} (com Media/DP: {tem_alvo[nv]})')
                cad_md = num(exp[2 * nv]) and num(exp[2 * nv + 1])
                if pts[nv] > 0 and not cad_md:
                    # 29/10/2026: o lote nao tem Media/DP cadastrada neste nivel (8383 PCR N1 com 1 resultado) -- o
                    # certo e dizer "SEM MEDIA/DP" e nao desenhar linha; nada de media no grafico
                    if 'SEM M' not in st[nv].upper() or num(md[2 * nv]) or num(md[2 * nv + 1]):
                        erro.append(f'N{nv + 1} sem Media/DP cadastrada e o Painel nao diz SEM MEDIA/DP')
                elif pts[nv] > 0:
                    if 'SEM M' in st[nv].upper():
                        erro.append(f'N{nv + 1} SEM MEDIA/DP')
                    if not (perto(md[2 * nv], exp[2 * nv]) and perto(md[2 * nv + 1], exp[2 * nv + 1])):
                        erro.append(f'N{nv + 1} media/DP do grafico != do lote')
            if erro:
                ruins.append((lote, an, erro))
            if k_lote == 0:
                lote0_ok = (filtro == filtro0)
        reg('V01 Todos os lotes com resultados, trocados no Painel como o usuário faz: o gráfico abre com pontos e, em todo '
            'nível com Média/DP, a média e as linhas -3DP..+3DP com 2 pontos cada na série do gráfico; sem "SEM MÉDIA/DP" '
            'no nível que tem ponto; a média/DP do gráfico são as do lote escolhido',
            not ruins and sum(1 for l in linhas_tab if l['resultados'] > 0) > 0,
            {'lotes_com_resultado': sum(1 for l in cad if com_dado.get(l, 0) > 0), 'lotes_com_problema': ruins[:8]})
        reg('V02 O lote em uso, que tem dado no filtro atual, não teve o filtro mexido (período e equipamento como estavam)',
            lote0_ok is True, {'lote': lote0, 'filtro_antes': filtro0,
                               'filtro_depois': linhas_tab[0]['filtro_de_ate_equip'] if linhas_tab else None})
        reg('V03 Lote sem nenhum resultado na base: nível sem Média/DP diz "SEM MÉDIA/DP" e não desenha linha; nível com '
            'Média/DP diz "AGUARDANDO 1ª CORRIDA" e mostra as 7 linhas do alvo', not sem_dado_ruins,
            {'lotes_sem_resultado': [l for l in cad if na_base.get(l, 0) == 0], 'problemas': sem_dado_ruins[:3]})
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
