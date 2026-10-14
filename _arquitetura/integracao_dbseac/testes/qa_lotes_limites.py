# -*- coding: utf-8 -*-
"""QA do ADR-073 -- aba Analitos como central de limites por lote, com rastreabilidade. Roda numa COPIA instalada.

Uso: python qa_lotes_limites.py <Bioquimica|Hematologia> <copia_instalada.xlsm> [saida.json]

Pelo caminho real (eventos ligados): escolhe um lote SEM Media/DP no seletor da aba Analitos (J1), preenche os
vazios pelo CIQ (mLotes.PreencherVaziosCIQ, o nucleo do botao), digita um valor como o usuario e confere a
LotesStore, o Painel e os limites do grafico. Nada e salvo no arquivo.

  L01 seletor J1 = lote do Painel (H3) = lote dos parametros em tela; lote sem Media/DP: editor vazio, resumo diz
      "SEM media/DP" (por isso o grafico ficava sem limites)
  L02 preencher vazios pelo CIQ: Media/DP = media e DP amostral dos 20 PRIMEIROS resultados do lote (Ativos, ordem
      de RUN, equipamento do filtro), recalculados aqui de forma independente, analito a analito, nivel a nivel;
      par com menos de 20 resultados fica vazio
  L03 retrato da especificacao em cada linha gravada: fonte, ano, ETp, CVTp = os da aba Analitos agora; registrado
      hoje; origem "CIQ: 20 primeiros"
  L04 limites do grafico existem para o lote novo: eixo do nivel 1 (axmin1/axmax1) numerico e envolvendo media+-3DP;
      Painel sem "SEM MEDIA/DP" no nivel com parametro
  L05 digitar um valor (como o usuario): grava na LotesStore, origem "digitado", trilha PARAMETRO_LOTE_ALTERADO
  L06 preencher de novo nao sobrescreve nada (0 preenchidos; valores iguais)
  L07 voltar ao 1o lote: os parametros dele estao intactos e o seletor acompanha
  L08 seguranca do teste: arquivo de entrada intacto
"""
import datetime as dt
import json
import os
import statistics as st
import sys
import traceback

AQUI = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, AQUI)
from qa_final import QA, reg, RES, sha256_arquivo  # noqa: E402

N_ALVO = 20


def num(v):
    return isinstance(v, (int, float)) and not isinstance(v, bool) and not (isinstance(v, int) and v < -2146820000)


def perto(a, b, tol=1e-9):
    return num(a) and num(b) and abs(a - b) <= tol * max(1.0, abs(b))


def escolher(q, rows, eq):
    """Lote do cadastro (que nao e o em tela) com mais resultados Ativos no equipamento do filtro e sem Media/DP."""
    cad = [str(v[0]).strip() for v in q.nome('regLoteCol').Value if v[0] not in (None, '')]
    atual = str(q.nome('loteParam').Value or '').strip()
    cont = {}
    for r in rows:
        if str(r.get('PARTICIPA_ESTATISTICA', '')).upper() != 'SIM':
            continue
        if eq and str(r.get('EQUIPAMENTO', '')).upper() != eq:
            continue
        cont[str(r['LOTE']).strip()] = cont.get(str(r['LOTE']).strip(), 0) + 1
    cand = [(cont.get(l, 0), l) for l in cad if l != atual and l.replace('.0', '') != atual]
    cand.sort(reverse=True)
    for n, l in cand:
        if n >= N_ALVO and 'SEM media/DP' in str(q.run('mLotes.RegistroLote', l.replace('.0', ''), teto=60)):
            return l.replace('.0', ''), n
    return None, 0


def esperado(rows, lote, eq):
    grupos = {}
    for i, r in enumerate(rows):
        if str(r['LOTE']).strip().replace('.0', '') != lote:
            continue
        if str(r.get('PARTICIPA_ESTATISTICA', '')).upper() != 'SIM':
            continue
        if eq and str(r.get('EQUIPAMENTO', '')).upper() != eq:
            continue
        if not num(r.get('RESULTADO')):
            continue
        k = (str(r['ANALITO']).strip().upper(), int(r['NIVEL']))
        grupos.setdefault(k, []).append((float(r['RUN'] or 0), i, float(r['RESULTADO'])))
    out = {}
    for k, xs in grupos.items():
        if len(xs) < N_ALVO:
            out[k] = None
            continue
        xs.sort(key=lambda t: (t[0], t[1]))
        v = [t[2] for t in xs[:N_ALVO]]
        out[k] = (st.mean(v), st.stdev(v))
    return out


def executar(produto, caminho, saida):
    hash0 = sha256_arquivo(caminho)
    q = QA(produto, caminho)
    wb = q.wb
    try:
        q.run('mSeguranca.ReprotectAll')               # sessao de usuario: macros escrevem (UserInterfaceOnly)
        a = q.ws('Analitos')
        p = q.ws('Painel')
        eq = str(q.run('mEstatistica.EquipFiltro', teto=60) or '').strip().upper()
        rows = q.ler('tblCQ_Final')
        lote0 = str(q.nome('loteParam').Value or '').strip()
        nomes = [(r, str(a.Range(f'A{r}').Value).strip()) for r in range(4, 44) if a.Range(f'A{r}').Value]
        antes0 = {(an, nv): q.run('mLotes.MediaDoLote', lote0, an, nv) for r, an in nomes for nv in (1, 2)}
        lote, nlote = escolher(q, rows, eq)
        if not lote:
            raise RuntimeError('nenhum lote do cadastro sem Media/DP com resultados suficientes para o teste')

        # ------------------------------------------------------------ L01 seletor
        a.Range('J1').Value = "'" + lote                       # como o usuario: eventos ligados
        q.ex.esperar()
        sel, h3, par = str(a.Range('J1').Value), str(p.Range('H3').Value), str(q.nome('loteParam').Value)
        editor = [a.Range(f'E{r}:J{r}').Value[0] for r, _ in nomes]
        vazio = all(v in (None, '') for linha in editor for v in linha)
        resumo = str(a.Range('K1').Text)
        reg('L01 Seletor da aba Analitos (J1) = lote do Painel (H3) = lote dos parâmetros em tela; lote sem Média/DP '
            'mostra o editor vazio e o resumo diz "SEM media/DP" (era por isso que o gráfico ficava sem limites)',
            sel == lote and h3.replace('.0', '') == lote and par == lote and vazio and 'SEM media/DP' in resumo,
            {'lote': lote, 'resultados': nlote, 'J1': sel, 'H3': h3, 'loteParam': par, 'editor_vazio': vazio,
             'resumo': resumo[:160]})

        # ------------------------------------------------------------ L02 preencher pelo CIQ
        r1 = str(q.run('mLotes.PreencherVaziosCIQ', teto=600))
        esp = esperado(rows, lote, eq)
        dif, cheios, vazios_ok = [], 0, 0
        for r, an in nomes:
            for nv in (1, 2, 3)[:q.nlv]:
                m, s = q.run('mLotes.MediaDoLote', lote, an, nv), q.run('mLotes.DPDoLote', lote, an, nv)
                e = esp.get((an.upper(), nv))
                if e is None:
                    if m not in ('', None):
                        dif.append((an, nv, 'devia ficar vazio', m))
                    else:
                        vazios_ok += 1
                else:
                    if perto(m, e[0]) and perto(s, e[1]):
                        cheios += 1
                    else:
                        dif.append((an, nv, (m, s), e))
        reg('L02 Preencher vazios pelo CIQ: Média/DP = média e DP amostral dos 20 PRIMEIROS resultados do lote (Ativos, '
            'ordem de RUN, equipamento do filtro), recalculados aqui de forma independente; par com menos de 20 fica vazio',
            r1.startswith('OK|') and cheios > 0 and not dif,
            {'retorno': r1, 'pares_preenchidos_conferidos': cheios, 'vazios_corretos': vazios_ok,
             'equipamento': eq, 'divergencias': dif[:6]})

        # ------------------------------------------------------------ L03 retrato da especificacao
        ls = q.ws('LotesStore')
        ult = ls.Cells(ls.Rows.Count, 16).End(-4162).Row
        linhas = {str(ls.Cells(i, 16).Value or '').upper(): i for i in range(2, ult + 1)}
        cF = next(c for c in range(1, 60) if str(a.Cells(3, c).Value or '').upper().replace('  ', ' ').startswith(('ESPQ FONTE', 'ETP FONTE')))
        cE = next(c for c in range(1, 60) if str(a.Cells(3, c).Value or '').upper().replace('  ', ' ').startswith(('ETP EM USO', 'ETP FINAL')))
        hoje = dt.date.today()
        dif3, vistos = [], 0
        for r, an in nomes:
            i = linhas.get(f'{lote}|{an}'.upper())
            if not i or not any(num(x) for x in ls.Range(f'C{i}:H{i}').Value[0]):
                continue
            vistos += 1
            fonte, etp = str(a.Cells(r, cF).Value or '-'), a.Cells(r, cE).Value
            q_, s_, u_, w_ = ls.Cells(i, 17).Value, ls.Cells(i, 19).Value, ls.Cells(i, 21).Value, str(ls.Cells(i, 23).Value or '')
            ok = (str(q_) == fonte and (perto(s_, etp) if num(etp) and etp > 0 else s_ == '-')
                  and hasattr(u_, 'year') and u_.date() == hoje and w_.startswith('CIQ: 20 primeiros'))
            if not ok:
                dif3.append((an, q_, fonte, s_, etp, str(u_), w_[:40]))
        reg('L03 Cada linha gravada leva o retrato da especificação vigente (fonte, ETp) = o da aba Analitos agora, '
            'registrado hoje, origem "CIQ: 20 primeiros"', vistos > 0 and not dif3,
            {'linhas_conferidas': vistos, 'divergencias': dif3[:6]})

        # ------------------------------------------------------------ L04 limites no grafico
        an_ok = next((an for r, an in nomes if num(q.run('mLotes.MediaDoLote', lote, an, 1))), None)
        ev4 = {}
        ok4 = False
        if an_ok:
            pos = [an for _, an in nomes].index(an_ok) + 1
            p.Range('B3').Value = pos
            q.run('mEstatistica.PainelMudou', teto=300)
            q.ex.xl.Calculate()
            m1, s1 = q.run('mLotes.MediaDoLote', lote, an_ok, 1), q.run('mLotes.DPDoLote', lote, an_ok, 1)
            mn, mx = q.nome('axmin1').Value, q.nome('axmax1').Value
            st1 = str(p.Range('J7').Value or '')
            ok4 = num(mn) and num(mx) and mn <= m1 - 3 * s1 + 1e-9 and mx >= m1 + 3 * s1 - 1e-9 and 'SEM M' not in st1.upper()
            ev4 = {'analito': an_ok, 'media': m1, 'dp': s1, 'axmin1': mn, 'axmax1': mx, 'status_N1': st1,
                   'C3': str(p.Range('C3').Value)}
        reg('L04 Lote novo com Média/DP: o gráfico ganha limites (eixo do nível 1 numérico envolvendo média ± 3 DP) e o '
            'Painel deixa de dizer "SEM MÉDIA/DP"', ok4, ev4)

        # ------------------------------------------------------------ L05 digitar (caminho do usuario)
        r_an, an5 = next((r, an) for r, an in nomes if num(q.run('mLotes.MediaDoLote', lote, an, 1)))
        novo = round(float(q.run('mLotes.MediaDoLote', lote, an5, 1)) + 0.5, 4)
        a.Range(f'E{r_an}').Value = novo
        q.ex.esperar()
        i5 = linhas.get(f'{lote}|{an5}'.upper())
        m5 = q.run('mLotes.MediaDoLote', lote, an5, 1)
        org5 = str(ls.Cells(i5, 23).Value or '')
        aud = q.ws('Audit_Log')
        ua = aud.Cells(aud.Rows.Count, 1).End(-4162).Row
        trilha = ' '.join(str(aud.Cells(i, c).Value or '') for i in range(max(2, ua - 5), ua + 1) for c in range(1, 25))
        reg('L05 Digitar um valor como o usuário: grava na LotesStore, origem "digitado", trilha PARAMETRO_LOTE_ALTERADO',
            perto(m5, novo) and org5.startswith('digitado') and 'PARAMETRO_LOTE_ALTERADO' in trilha,
            {'analito': an5, 'digitado': novo, 'lido': m5, 'origem': org5})

        # ------------------------------------------------------------ L06 nao sobrescreve
        antes6 = {(an, nv): q.run('mLotes.MediaDoLote', lote, an, nv) for _, an in nomes for nv in (1, 2)}
        r6 = str(q.run('mLotes.PreencherVaziosCIQ', teto=600))
        depois6 = {(an, nv): q.run('mLotes.MediaDoLote', lote, an, nv) for _, an in nomes for nv in (1, 2)}
        reg('L06 Preencher de novo não sobrescreve nada (0 preenchidos; o valor digitado continua)',
            r6.startswith('OK|0|') and antes6 == depois6, {'retorno': r6})

        # ------------------------------------------------------------ L07 voltar ao 1o lote
        a.Range('J1').Value = "'" + lote0
        q.ex.esperar()
        depois0 = {(an, nv): q.run('mLotes.MediaDoLote', lote0, an, nv) for r, an in nomes for nv in (1, 2)}
        reg('L07 Voltar ao 1º lote: parâmetros dele intactos; seletor, Painel e editor acompanham',
            depois0 == antes0 and str(q.nome('loteParam').Value) == lote0 and str(p.Range('H3').Value).replace('.0', '') == lote0,
            {'lote0': lote0, 'iguais': depois0 == antes0})
    except Exception as ex_geral:
        reg('L00 Execução interrompida: os testes seguintes a este ponto não rodaram', False,
            {'erro': f'{type(ex_geral).__name__}: {ex_geral}', 'traceback': traceback.format_exc()[-1500:]})
    finally:
        q.fechar(salvar=False)
    hash1 = sha256_arquivo(caminho)
    reg('L08 Segurança do teste: arquivo de entrada intacto (SHA-256 igual antes e depois; nada salvo nele)',
        hash0 == hash1, {'arquivo': caminho, 'sha256_antes': hash0, 'sha256_depois': hash1})
    if saida:
        with open(saida, 'w', encoding='utf-8') as f:
            json.dump(RES, f, ensure_ascii=False, indent=1, default=str)
    falhas = [r for r in RES if r['resultado'] == 'FAIL']
    print(f'\n=== {produto} lotes/limites (ADR-073): {len(RES) - len(falhas)} PASS / {len(falhas)} FAIL ===', flush=True)
    return falhas


if __name__ == '__main__':
    falhas = executar(sys.argv[1], os.path.abspath(sys.argv[2]), sys.argv[3] if len(sys.argv) > 3 else None)
    sys.exit(1 if falhas else 0)
