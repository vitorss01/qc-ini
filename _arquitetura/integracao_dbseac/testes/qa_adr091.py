# -*- coding: utf-8 -*-
"""QA do ADR-091 (05/11/2026): regras de material por LOTE (tblLotesMaterial). Roda numa COPIA instalada
(instalar_adr091.py); nada e salvo. Dados de teste: regras e manuais com "TESTE-ADR091" (so na copia).

Uso: python qa_adr091.py <Bioquimica|Hematologia> <copia_instalada.xlsm> <saida.json>

Cada passo edita a tblLotesMaterial e roda o ATUALIZAR so das entradas (mIntegracao.AtualizarSoEntradas):
  T01 a tabela existe nos dois produtos com as colunas do ADR e a carga inicial do config_produtos (Bio 6 regras,
      Hema 0); LotesMaterialResumo sem aviso
  T02 SEM regras: 0 LOTE_*, 0 A13, 0 A14 (a referencia dos passos seguintes)
  T03 regras REAIS (Bio): MATRIZ -> os lotes 6353/6354/5659 so com URINA e nenhum resultado do interfaceamento num
      analito cadastrado (sairam do soro); SO_ANALITOS 8353 -> so a PCR pode participar; FORA_DO_CQ 8177/8794 -> 0
      participantes. Contagens por lote contra a tabela (antes/depois). Fora desses lotes, STATUS/PARTICIPA/RUN/MATRIZ/
      ANALITO iguais aos de sem regras. Hema: tudo igual ao de sem regras
  T04 regras de TESTE nos dois produtos (lotes reais escolhidos nos dados): SO_ANALITOS (analito em minusculas, com
      espacos), FORA_DO_CQ (lote com um INATIVADO: continua INATIVADO), MATRIZ (matriz sintetica), mais manuais em cada
      lote (inclusive resultado -1 num lote FORA_DO_CQ: o lote decide antes do IMPLAUSIVEL). Conjunto LOTE_* = o
      esperado (varredura independente); PARTICIPA = NAO e NAO_PLOTAR; A13 = os mesmos IDs; RUN dos LOTE_* igual ao
      de antes (contam na corrida); fora dos lotes de teste nada muda
  T05 linhas invalidas -> A14 exatamente nas linhas e com os textos esperados (LOTE vazio, REGRA desconhecida, MATRIZ
      sem matriz, MATRIZ repetida, SO_ANALITOS sem analito, analito fora do de/para)
  T06 apagar as regras de teste e ATUALIZAR: o interfaceamento volta ao estado das regras reais (reprocessa)
  Y00 seguranca do teste: arquivo de entrada intacto
"""
import collections
import datetime as dt
import json
import os
import sys
import time
import traceback

AQUI = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, AQUI)
from qa_final import QA, reg, RES, sha256_arquivo  # noqa: E402
sys.path.insert(0, os.path.join(AQUI, '..'))
import config_produtos as cp  # noqa: E402
import instalar_integracao as ii  # noqa: E402

TB = 'tblLotesMaterial'
TAG = 'TESTE-ADR091'
MZ_TESTE = 'TESTE-ADR091'
NAO = 'NÃO'
LOTE_ST = ('LOTE_FORA_DO_CQ', 'LOTE_OUTRO_MATERIAL')
# status que vem ANTES das regras do lote (o inativado continua INATIVADO)
PRECEDE = ('MANUAL_INCOMPLETO', 'DUPLICIDADE_ORIGEM', 'CONFLITO_MANUAL', 'SUBSTITUIDO_MANUAL', 'SEM_DATA_HORA',
           'SEM_VALOR', 'INATIVADO')
TXT = {'LOTE_FORA_DO_CQ': 'lote fora do CQ (regra FORA_DO_CQ da tblLotesMaterial; fora da estatistica e do grafico)',
       'LOTE_OUTRO_MATERIAL': 'lote de outro material neste analito (regra SO_ANALITOS da tblLotesMaterial; fora da '
                              'estatistica e do grafico)'}
# ADR-092 (N-01): manual num lote de regra MATRIZ com analito do catalogo sem de/para nessa matriz
TXT_MZSEM = ('manual com analito do catalogo sem cadastro na matriz do lote (regra MATRIZ da tblLotesMaterial; '
             'fora da estatistica e do grafico)')
DEPARA = {'cat': set(), 'linhas': []}       # preenchido na execucao (tblDeParaAnalitos da copia)


def n01(x, mz):
    """ADR-092 (N-01), varredura independente: manual + lote com regra MATRIZ + analito do catalogo + nenhuma linha do
    de/para com esse ANALITO_QCINI e MATRIZ vazia ou igual a do lote."""
    lt, an = up(x['LOTE']), txt(x['ANALITO'])
    if x.get('ORIGEM_RESULTADO') != 'MANUAL' or lt not in mz or an not in DEPARA['cat']:
        return False
    return not any(q_ == an and (not m_ or m_.upper() == mz[lt]) for m_, q_ in DEPARA['linhas'])


def txt(v):
    if isinstance(v, float) and v == int(v):
        return str(int(v))
    return str(v if v is not None else '').strip()


def up(v):
    return txt(v).upper() or None


def regras_de(linhas):
    """(mz, fora, so) das linhas validas -- as regras do ADR-091 (a 1a MATRIZ do lote vale)."""
    mz, fora, so = {}, set(), {}
    for l in linhas:
        lt, rg = up(l[0]), up(l[1])
        if not lt:
            continue
        if rg == 'MATRIZ' and up(l[2]) and lt not in mz:
            mz[lt] = up(l[2])
        elif rg == 'FORA_DO_CQ':
            fora.add(lt)
        elif rg == 'SO_ANALITOS':
            ans = {a for a in (up(p) for p in txt(l[3]).split(';')) if a}
            if ans:
                so.setdefault(lt, set()).update(ans)
    return mz, fora, so


def esperado_lote(x, fora, so, mz=None):
    if x['STATUS_ANALITICO'] in PRECEDE:
        return None
    lt = up(x['LOTE'])
    if lt in fora:
        return 'LOTE_FORA_DO_CQ'
    if lt in so and up(x['ANALITO']) not in so[lt]:
        return 'LOTE_OUTRO_MATERIAL'
    if mz and n01(x, mz):
        return 'LOTE_OUTRO_MATERIAL'                 # ADR-092 (N-01)
    return None


def executar(produto, caminho, saida):
    hash0 = sha256_arquivo(caminho)
    q = QA(produto, caminho)
    try:
        xl, bio = q.ex.xl, q.bio

        def gravar_regras(linhas):
            """A tblLotesMaterial passa a ter exatamente estas linhas (LOTE, REGRA, MATRIZ, ANALITOS, NOTA)."""
            xl.EnableEvents = False
            lo = q.lo(TB)
            ws = lo.Parent
            ws.Unprotect(ii.SENHA)
            if lo.ListRows.Count:
                lo.DataBodyRange.ClearContents()
            r0, c0 = lo.HeaderRowRange.Row, lo.HeaderRowRange.Column
            n = max(1, len(linhas))
            lo.Resize(ws.Range(ws.Cells(r0, c0), ws.Cells(r0 + n, c0 + lo.ListColumns.Count - 1)))
            cab = [str(c.Name) for c in lo.ListColumns]
            # celulas pela aba (o DataBodyRange de uma tabela recem-redimensionada e vazia vem None pelo COM)
            ws.Range(ws.Cells(r0 + 1, c0), ws.Cells(r0 + n, c0 + len(cab) - 1)).NumberFormat = '@'
            for i, l in enumerate(linhas, 1):
                for nome, v in zip(cp.COLUNAS_LOTES_MATERIAL, l):
                    if v is not None:
                        ws.Cells(r0 + i, c0 + cab.index(nome)).Value = v
            xl.EnableEvents = True

        def atualizar(rotulo):
            xl.EnableEvents = True
            t0 = time.time()
            r = str(q.run('mIntegracao.AtualizarSoEntradas', teto=1500))
            if not r.startswith('OK|'):
                raise RuntimeError(f'ATUALIZAR so entradas ({rotulo}): ' + r[:400])
            print(f'   ATUALIZAR {rotulo}: {time.time() - t0:.1f}s', flush=True)
            return q.ler('tblCQ_Final'), q.ler('tblQA_Integracao')

        def ver(fin, qa, linhas):
            """Varredura independente: conjunto LOTE_* esperado x obtido, PARTICIPA/TIPO/MOTIVO, A13 = LOTE_*."""
            mz, fora, so = regras_de(linhas)
            obt = {x['ID_REGISTRO']: x['STATUS_ANALITICO'] for x in fin if x['STATUS_ANALITICO'] in LOTE_ST}
            esp = {}
            for x in fin:
                e = esperado_lote(x, fora, so, mz) if x['STATUS_ANALITICO'] in LOTE_ST + PRECEDE + ('ATIVO', 'IMPLAUSIVEL') else None
                if e:
                    esp[x['ID_REGISTRO']] = e
            ruins = [x['ID_REGISTRO'] for x in fin if x['STATUS_ANALITICO'] in LOTE_ST and not (
                x['PARTICIPA_ESTATISTICA'] == NAO and x['TIPO_PLOTAGEM_LJ'] == 'NAO_PLOTAR'
                and x['REGISTRAR_RESULTADO_NO_LJ'] == NAO and txt(x['MOTIVO_EXCLUSAO_AUTOMATICA']) == (
                    TXT_MZSEM if x['STATUS_ANALITICO'] == 'LOTE_OUTRO_MATERIAL' and n01(x, mz) and not (
                        up(x['LOTE']) in so and up(x['ANALITO']) not in so[up(x['LOTE'])]) else TXT[x['STATUS_ANALITICO']]))]
            mz_ruim = [x['ID_REGISTRO'] for x in fin if up(x['LOTE']) in mz and up(x['MATRIZ']) != mz[up(x['LOTE'])]]
            a13 = {x['ID_REGISTRO'] for x in qa if x.get('CODIGO') == 'A13'}
            return {'obtido': obt, 'esperado': esp, 'ok': obt == esp and not ruins and not mz_ruim and a13 == set(obt),
                    'difs': sorted(set(obt.items()) ^ set(esp.items()))[:6], 'flags_ruins': ruins[:5],
                    'matriz_ruim': mz_ruim[:5], 'A13': len(a13), 'A13_igual': a13 == set(obt)}

        def por_lote(fin, lotes):
            c = collections.defaultdict(collections.Counter)
            for x in fin:
                lt = up(x['LOTE'])
                if lt in lotes:
                    c[lt]['linhas'] += 1
                    c[lt][x['STATUS_ANALITICO']] += 1
                    c[lt]['participa'] += x['PARTICIPA_ESTATISTICA'] == 'SIM'
            return {k: dict(v) for k, v in c.items()}

        def iguais_fora(fa, fb, lotes, cols=('STATUS_ANALITICO', 'PARTICIPA_ESTATISTICA', 'RUN', 'MATRIZ', 'ANALITO')):
            a = {x['ID_REGISTRO']: x for x in fa if up(x['LOTE']) not in lotes and not txt(x.get('MOTIVO_MANUAL')).startswith(TAG)}
            b = {x['ID_REGISTRO']: x for x in fb if up(x['LOTE']) not in lotes and not txt(x.get('MOTIVO_MANUAL')).startswith(TAG)}
            dif = [(i, c, a[i][c], b.get(i, {}).get(c)) for i in a for c in cols if i not in b or a[i][c] != b[i][c]]
            return set(a) == set(b) and not dif, dif[:5]

        dp_ = q.ler('tblDeParaAnalitos')
        DEPARA['linhas'] = [(txt(r.get('MATRIZ')), txt(r.get('ANALITO_QCINI'))) for r in dp_ if txt(r.get('ANALITO_ORIGEM'))]
        DEPARA['cat'] = {q_ for _, q_ in DEPARA['linhas'] if q_}

        # ------------------------------------------------------------------ T01
        lo = q.lo(TB)
        cab = [str(c.Name) for c in lo.ListColumns] if lo is not None else []
        reais = [[r.get(c) for c in cp.COLUNAS_LOTES_MATERIAL] for r in (q.ler(TB) if lo is not None else [])
                 if any(r.get(c) not in (None, '') for c in ('LOTE', 'REGRA', 'MATRIZ', 'ANALITOS'))]
        carga = cp.linhas_lotes_material(produto)
        resumo = str(q.run('mEtlCQ.LotesMaterialResumo', teto=120))
        reg('T01 tblLotesMaterial com as colunas do ADR e a carga inicial do config_produtos (%d regras); sem aviso (A14)'
            % len(cp.linhas_lotes_material(produto)),
            cab == cp.COLUNAS_LOTES_MATERIAL and [[up(v) for v in l[:4]] for l in reais] == [[up(v) for v in l[:4]] for l in carga]
            and resumo.endswith('|'), {'colunas': cab, 'regras': reais, 'resumo': resumo[:300]})
        fin_prod = q.ler('tblCQ_Final')                         # gravada antes da instalacao (sem as regras)

        # ------------------------------------------------------------------ T02 sem regras
        gravar_regras([])
        fin_sem, qa_sem = atualizar('sem regras')
        n_lote = sum(1 for x in fin_sem if x['STATUS_ANALITICO'] in LOTE_ST)
        a1314 = [x['CODIGO'] for x in qa_sem if x.get('CODIGO') in ('A13', 'A14')]
        reg('T02 Sem regras: 0 LOTE_*, 0 A13, 0 A14', n_lote == 0 and not a1314, {'LOTE_*': n_lote, 'A13/A14': len(a1314)})

        # ------------------------------------------------------------------ T03 regras reais
        gravar_regras(reais)
        fin_real, qa_real = atualizar('regras reais')
        mz, fora, so = regras_de(reais)
        lotes_reais = set(mz) | fora | set(so)
        v = ver(fin_real, qa_real, reais)
        ok_f, dif_f = iguais_fora(fin_sem, fin_real, lotes_reais)
        soro = [x['ID_REGISTRO'] for x in fin_real if up(x['LOTE']) in mz and x['ORIGEM_RESULTADO'] == 'INTERFACEAMENTO'
                and x['ANALITO_CADASTRADO'] == 'SIM']
        part_ruim = [x['ID_REGISTRO'] for x in fin_real if x['PARTICIPA_ESTATISTICA'] == 'SIM' and (
            up(x['LOTE']) in fora or (up(x['LOTE']) in so and up(x['ANALITO']) not in so[up(x['LOTE'])]))]
        a14 = [x for x in qa_real if x.get('CODIGO') == 'A14']
        ev = {'por_lote_antes_instalar': por_lote(fin_prod, lotes_reais), 'por_lote_sem_regras': por_lote(fin_sem, lotes_reais),
              'por_lote_com_regras': por_lote(fin_real, lotes_reais), 'varredura': {k: v[k] for k in v if k not in ('obtido', 'esperado')},
              'LOTE_*': len(v['obtido']), 'interface_URINA_em_analito_cadastrado': soro[:5], 'participa_indevido': part_ruim[:5],
              'A14': len(a14), 'fora_dos_lotes_igual_sem_regras': ok_f, 'dif_fora': dif_f}
        if bio:
            ok = v['ok'] and not soro and not part_ruim and not a14 and ok_f and len(v['obtido']) > 0
            ev['lotes_da_tabela_sem_dado'] = sorted(lt for lt in lotes_reais if lt not in ev['por_lote_com_regras'])
        else:
            ok = not v['obtido'] and not a14 and iguais_fora(fin_sem, fin_real, set())[0]
        reg('T03 Regras reais: %s' % ('MATRIZ URINA sai do soro; 8353 so PCR; 8177/8794 fora; resto igual ao de sem regras'
                                      if bio else 'nenhuma regra, tabela igual a sem regras'), ok, ev)

        # ------------------------------------------------------------------ T04 regras de teste
        intf = [x for x in fin_real if x['ORIGEM_RESULTADO'] == 'INTERFACEAMENTO' and up(x['LOTE'])
                and up(x['LOTE']) not in lotes_reais and hasattr(x['DATA_HORA'], 'year')]
        cont = collections.Counter(up(x['LOTE']) for x in intf if x['STATUS_ANALITICO'] == 'ATIVO')
        ans = collections.defaultdict(set)
        for x in intf:
            if x['STATUS_ANALITICO'] == 'ATIVO':
                ans[up(x['LOTE'])].add(txt(x['ANALITO']))
        inat = [x for x in intf if x['STATUS_ANALITICO'] == 'INATIVADO' and cont[up(x['LOTE'])] > 0]
        inat_nova = None
        if not inat:
            # base sem inativado do interfaceamento (Hema): a suite inativa um ATIVO do lote mais cheio (so na copia)
            lt0 = cont.most_common(1)[0][0]
            x0 = next(x for x in intf if up(x['LOTE']) == lt0 and x['STATUS_ANALITICO'] == 'ATIVO')
            q.escrever_linha('tblInativacao_NaoConformes', {'ID_REGISTRO': x0['ID_REGISTRO'], 'ANALITO': txt(x0['ANALITO']),
                                                            'MOTIVO': f'Outro: {TAG} inativado num lote FORA_DO_CQ'})
            inat, inat_nova = [x0], x0['ID_REGISTRO']
        L_fora = up(inat[-1]['LOTE'])
        livres = [lt for lt, _ in cont.most_common() if lt != L_fora]
        L_so = next(lt for lt in livres if len(ans[lt]) >= 2)
        L_mz = next(lt for lt in livres if lt != L_so)
        A = sorted(ans[L_so])[0]
        B = sorted(ans[L_so])[1]
        teste = [
            [L_so, 'so_analitos', None, '  ' + A.lower() + ' ', f'{TAG} SO_ANALITOS'],
            [L_fora, 'FORA_DO_CQ', None, None, f'{TAG} FORA_DO_CQ'],
            [L_mz.lower(), 'Matriz', MZ_TESTE.lower(), None, f'{TAG} MATRIZ'],
        ]
        n0 = len(reais)
        invalidas = [  # (linha, regra, texto do A14)
            (['', 'FORA_DO_CQ', None, None, TAG], None, 'FORA_DO_CQ', 'LOTE vazio; a linha foi ignorada'),
            ([L_mz, 'XYZ', None, None, TAG], L_mz, 'XYZ',
             "REGRA 'XYZ' desconhecida (use MATRIZ, SO_ANALITOS ou FORA_DO_CQ); a linha foi ignorada"),
            ([L_so, 'MATRIZ', None, None, TAG], L_so, 'MATRIZ', 'regra MATRIZ sem a MATRIZ; a linha foi ignorada'),
            ([L_mz, 'MATRIZ', 'OUTRA', None, TAG], L_mz, 'MATRIZ',
             f'o lote ja tem uma regra MATRIZ antes desta ({MZ_TESTE}); esta linha foi ignorada'),
            ([L_fora, 'SO_ANALITOS', None, ' ; ', TAG], L_fora, 'SO_ANALITOS', 'regra SO_ANALITOS sem ANALITOS; a linha foi ignorada'),
            ([L_so, 'SO_ANALITOS', None, f'{A}; Analito Inexistente ADR091', TAG], L_so, 'SO_ANALITOS',
             'analito(s) fora do de/para do QC_INI: Analito Inexistente ADR091 (nenhum resultado tem esse nome; confira a grafia)'),
        ]
        todas = reais + teste + [x[0] for x in invalidas]
        esp_a14 = sorted(f'linha {n0 + len(teste) + k} (lote {lt or "?"}, regra {rg}): {m}'
                         for k, (_, lt, rg, m) in enumerate(invalidas, 1))
        gravar_regras(todas)
        # manuais nos lotes de teste (eventos desligados, ID explicito, como o qa_adr088)
        base = {r: next(x for x in intf if up(x['LOTE']) == lt and txt(x['ANALITO']) == an and x['STATUS_ANALITICO'] == 'ATIVO')
                for r, lt, an in (('fora', L_fora, sorted(ans[L_fora])[0]), ('so_b', L_so, B), ('so_a', L_so, A),
                                  ('mz', L_mz, sorted(ans[L_mz])[0]))
                if any(up(x['LOTE']) == lt and txt(x['ANALITO']) == an and x['STATUS_ANALITICO'] == 'ATIVO' for x in intf)}
        casos = {'M1 fora': ('fora', 2.5, {}), 'M2 so outro analito': ('so_b', 2.5, {}), 'M3 so analito da lista': ('so_a', 2.5, {}),
                 'M4 matriz': ('mz', 2.5, {'MATRIZ': 'SORO' if bio else 'SANGUE TOTAL'}), 'M5 fora e -1': ('fora', -1, {})}
        xl.EnableEvents = False
        for j, (k, (b, val, extra)) in enumerate(casos.items()):
            if b not in base:
                continue
            x = base[b]
            d = x['DATA_HORA']
            q.escrever_linha('tblResultados_Manuais', dict(
                ID_REGISTRO=f'MAN_{9201 + j}', DATA=dt.datetime(d.year, d.month, d.day), HORA=f'23:{50 + j}',
                EQUIPAMENTO=txt(x['EQUIPAMENTO']), LOTE="'" + txt(x['LOTE']), NIVEL=int(x['NIVEL']), ANALITO=txt(x['ANALITO']),
                RESULTADO=val, MOTIVO=f'{TAG} {k}', **extra))
        xl.EnableEvents = True
        fin_t, qa_t = atualizar('regras de teste')
        v4 = ver(fin_t, qa_t, todas)
        lotes_t = {L_so, L_fora, L_mz}
        ok_f4, dif_f4 = iguais_fora(fin_real, fin_t, lotes_t)
        x_inat = next((x for x in fin_t if x['ID_REGISTRO'] == inat[-1]['ID_REGISTRO']), {})
        ant = {x['ID_REGISTRO']: x for x in fin_real}
        run_mudou = [x['ID_REGISTRO'] for x in fin_t if x['STATUS_ANALITICO'] in LOTE_ST and x['ID_REGISTRO'] in ant
                     and (x['RUN'], x['CORRIDA_NO_DIA']) != (ant[x['ID_REGISTRO']]['RUN'], ant[x['ID_REGISTRO']]['CORRIDA_NO_DIA'])]
        man = {txt(x['MOTIVO_MANUAL'])[len(TAG) + 1:]: x for x in fin_t if txt(x.get('MOTIVO_MANUAL')).startswith(TAG)}
        esp_man = {'M1 fora': lambda x: x['STATUS_ANALITICO'] == 'LOTE_FORA_DO_CQ',
                   'M2 so outro analito': lambda x: x['STATUS_ANALITICO'] == 'LOTE_OUTRO_MATERIAL',
                   'M3 so analito da lista': lambda x: x['STATUS_ANALITICO'] not in LOTE_ST,
                   'M4 matriz': lambda x: up(x['MATRIZ']) == MZ_TESTE,
                   'M5 fora e -1': lambda x: x['STATUS_ANALITICO'] == 'LOTE_FORA_DO_CQ'}
        ok_man = all(k in man and f(man[k]) for k, f in esp_man.items() if casos[k][0] in base)
        ok = (v4['ok'] and ok_f4 and x_inat.get('STATUS_ANALITICO') == 'INATIVADO' and not run_mudou and ok_man
              and any(s == 'LOTE_OUTRO_MATERIAL' for s in v4['obtido'].values())
              and any(s == 'LOTE_FORA_DO_CQ' for s in v4['obtido'].values()))
        reg('T04 Regras de teste (SO_ANALITOS em minusculas, FORA_DO_CQ com um INATIVADO, MATRIZ sintetica) e manuais: '
            'LOTE_* = varredura independente; NAO participa/plota; A13 = os mesmos IDs; inativado continua INATIVADO; '
            'RUN igual; fora dos lotes de teste nada muda',
            ok, {'lotes': {'SO_ANALITOS': (L_so, A), 'FORA_DO_CQ': L_fora, 'MATRIZ': L_mz},
                 'varredura': {k: v4[k] for k in v4 if k not in ('obtido', 'esperado')},
                 'por_status': dict(collections.Counter(v4['obtido'].values())), 'inativado': (inat[-1]['ID_REGISTRO'],
                 x_inat.get('STATUS_ANALITICO')), 'run_mudou': run_mudou[:5], 'fora_igual': ok_f4, 'dif_fora': dif_f4,
                 'manuais': {k: (x['ID_REGISTRO'], x['STATUS_ANALITICO'], x['MATRIZ']) for k, x in man.items()},
                 'manuais_ok': ok_man})

        # ------------------------------------------------------------------ T05 A14
        a14 = sorted(txt(x['DETALHE']) for x in qa_t if x.get('CODIGO') == 'A14')
        reg('T05 Linhas invalidas da tblLotesMaterial -> A14 nas linhas e com os textos esperados', a14 == esp_a14,
            {'obtido': a14, 'esperado': esp_a14})

        # ------------------------------------------------------------------ T06 apagar as regras de teste
        gravar_regras(reais)
        fin_r, qa_r = atualizar('regras de teste apagadas')
        intf_r = {x['ID_REGISTRO']: x for x in fin_r if x['ORIGEM_RESULTADO'] == 'INTERFACEAMENTO'}
        # RUN so fora dos lotes de teste: os manuais de teste (23:5x) continuam la e podem abrir uma corrida no fim do dia
        dif = [(i, c, x[c], intf_r.get(i, {}).get(c)) for i, x in
               ((x['ID_REGISTRO'], x) for x in fin_real if x['ORIGEM_RESULTADO'] == 'INTERFACEAMENTO')
               for c in ('STATUS_ANALITICO', 'PARTICIPA_ESTATISTICA', 'MATRIZ', 'ANALITO', 'RUN')
               if (c != 'RUN' or up(x['LOTE']) not in lotes_t) and i != inat_nova and x[c] != intf_r.get(i, {}).get(c)]
        man_r = [x['ID_REGISTRO'] for x in fin_r if txt(x.get('MOTIVO_MANUAL')).startswith(TAG) and x['STATUS_ANALITICO'] in LOTE_ST]
        reg('T06 Regras de teste apagadas + ATUALIZAR: o interfaceamento volta ao estado das regras reais; nenhum manual de '
            'teste com LOTE_*', not dif and not man_r and not [x for x in qa_r if x.get('CODIGO') == 'A14'],
            {'diferencas': dif[:6], 'manuais_LOTE': man_r})
    except Exception as ex_:                         # noqa: BLE001
        reg('Z99 Execução da suíte sem exceção', False, {'erro': repr(ex_), 'trace': traceback.format_exc()[-2000:]})
    finally:
        q.fechar(salvar=False)
    hash1 = sha256_arquivo(caminho)
    reg('Y00 Segurança do teste: arquivo de entrada intacto (SHA-256 igual antes e depois; nada salvo nele)',
        hash0 == hash1, {'arquivo': caminho, 'sha256_antes': hash0, 'sha256_depois': hash1})
    if saida:
        with open(saida, 'w', encoding='utf-8') as f:
            json.dump(RES, f, ensure_ascii=False, indent=1, default=str)
    falhas = [r for r in RES if r['resultado'] == 'FAIL']
    print(f'\n=== {produto} ADR-091 (material por lote): {len(RES) - len(falhas)} PASS / {len(falhas)} FAIL ===', flush=True)
    return falhas


if __name__ == '__main__':
    falhas = executar(sys.argv[1], os.path.abspath(sys.argv[2]), sys.argv[3] if len(sys.argv) > 3 else None)
    sys.exit(1 if falhas else 0)
