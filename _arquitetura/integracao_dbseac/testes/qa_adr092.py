# -*- coding: utf-8 -*-
"""QA do ADR-092 (11/11/2026): ocorrencias do reteste 03 (Tester) -- N-01, N-02, N-03, OC-0040, OC-0043, OC-0049.
Roda numa COPIA instalada (instalar_adr092.py); nada e salvo. Dados de teste com "TESTE-ADR092" (so na copia).

Uso: python qa_adr092.py <Bioquimica|Hematologia> <copia_instalada.xlsm> <saida.json>

  K01 N-01: manual num lote com regra MATRIZ (lote de teste, matriz URINA) e analito do catalogo -> o status e o da
      varredura independente (ADR-093: codigo de origem unico, sem de/para em URINA e ja recebido em URINA -> ATIVO nesse
      codigo, com a nota e o I05; senao LOTE_OUTRO_MATERIAL com o texto proprio e A13; Hema: o de/para nao tem matriz,
      entao ATIVO); o mesmo lote com um nome FORA do catalogo nao vira LOTE_*; o mesmo analito num lote comum fica ATIVO
  K02 N-02: sem login, CorrigirCEQ e ExcluirCEQ devolvem "ERRO|0|Sessão encerrada…" e nada muda (tabela e Audit_Log)
  K03 N-03: manual IMPLAUSIVEL (-1) e o LOTE_* do K01 -> trilha RESULTADO_MANUAL_FORA_DA_ESTATISTICA, nunca INCLUIDO
  K04 OC-0043: cada status de J7:J9 (as formas curtas da formula + OK/ALERTA/REJEITADO) cabe com reducao <= 30 % na
      largura util de J7 (fonte efetiva); "Reduzir para caber" ligado; o Q4 diz a forma longa de cada estado
  K05 OC-0040 (Bio): Estatistica D4 = Painel!K3 (nunca "Ano 0"), rotulo da Visao, periodo efetivo vazio = "historico
      inteiro do lote"
  K06 OC-0049 (Hema): linhas sem analito com A..T vazias; as linhas com analito mostram o nivel
  Y00 seguranca do teste: arquivo de entrada intacto
"""
import collections
import datetime as dt
import json
import os
import re
import sys
import traceback

AQUI = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, AQUI)
from qa_final import QA, reg, RES, sha256_arquivo  # noqa: E402
sys.path.insert(0, os.path.join(AQUI, '..'))
import config_produtos as cp  # noqa: E402
import instalar_integracao as ii  # noqa: E402
import auditoria_abas  # noqa: E402

TAG = 'TESTE-ADR092'
NAO = 'NÃO'
LOTE_T = 'QA092U'
TXT_MZCORR = ("analito '%s' digitado num lote de URINA: lancado como '%s' (o analito que o interfaceamento usa nessa "
              "matriz; regra MATRIZ da tblLotesMaterial)")                               # ADR-093, = mEtlCQ / DB_CQ_FINAL
TXT_MZSEM = ('manual com analito do catalogo sem cadastro na matriz do lote (regra MATRIZ da tblLotesMaterial; '
             'fora da estatistica e do grafico)')
SESSAO = 'ERRO|0|Sessão encerrada'
LONGOS = {'DESATUAL.': 'DESATUALIZADO', 'BASE VAZIA': 'BASE SEM DADOS', 'SEM MÉD/DP': 'SEM MÉDIA/DP',
          'AGUARDA 1ª': 'AGUARDANDO 1ª CORRIDA', 'SEM CORRIDA': 'SEM CORRIDA DESTE ANALITO',
          'PER. VAZIO': 'SEM CORRIDA NO PERÍODO', 'SEM PONTO': 'SEM PONTO NO PERÍODO'}


def txt(v):
    if isinstance(v, float) and v == int(v):
        return str(int(v))
    return '' if v is None else str(v).strip()


def executar(produto, caminho, saida):
    hash0 = sha256_arquivo(caminho)
    bio = produto.startswith('Bio')
    q = QA(produto, caminho)
    try:
        xl = q.ex.xl
        al = q.ws('Audit_Log')

        def xrun(macro, *a):
            return xl.Run(f"'{q.wb.Name}'!{macro}", *a)

        def aud0():
            return al.Cells(al.Rows.Count, 1).End(-4162).Row

        def aud(r0):
            ult = aud0()
            if ult <= r0:
                return []
            v = al.Range(al.Cells(r0 + 1, 1), al.Cells(ult, 30)).Value
            return [(txt(l[6]), ' | '.join(txt(c) for c in l)) for l in v]

        # ------------------------------------------------------------ K02 N-02 (antes de qualquer login)
        lo_cap = q.lo('tblEQA_CAP_Dados')
        vals0 = lo_cap.DataBodyRange.Value
        lin = None
        for i, r in enumerate(vals0, 1):
            if any(txt(c) for c in r[:6]):
                lin = lo_cap.DataBodyRange.Rows(i).Row
                break
        campos = ['CAP', 'QA-092', '2026', 'x', 'y', '1', '', '', '', '', '', 'A', '', '', '', 'qa_092.pdf']
        a0 = aud0()
        rc = str(xrun('mCEQForm.CorrigirCEQ', lin or 5, campos))
        rx = str(xrun('mCEQForm.ExcluirCEQ', 'CAP', lin or 5))
        vals1 = lo_cap.DataBodyRange.Value
        trilha = [a for a, _ in aud(a0) if a.startswith('CEQ_')]
        reg('K02 N-02: sem login, CorrigirCEQ e ExcluirCEQ devolvem "ERRO|0|Sessão encerrada…" e nada muda (tabela do CAP e '
            'Audit_Log)', rc.startswith(SESSAO) and rx.startswith(SESSAO) and vals0 == vals1 and not trilha and lin is not None,
            {'linha': lin, 'corrigir': rc[:80], 'excluir': rx[:80], 'tabela_igual': vals0 == vals1, 'trilha_CEQ': trilha})

        # ------------------------------------------------------------ K01 / K03 N-01 e N-03
        fin0 = q.ler('tblCQ_Final')
        dp = q.ler('tblDeParaAnalitos')
        linhas_dp = [(txt(r.get('MATRIZ')), txt(r.get('ANALITO_QCINI'))) for r in dp if txt(r.get('ANALITO_ORIGEM'))]
        trios_dp = [(txt(r.get('MATRIZ')), txt(r.get('ANALITO_QCINI')), txt(r.get('ANALITO_ORIGEM')))
                    for r in dp if txt(r.get('ANALITO_ORIGEM'))]
        cat = {a for _, a in linhas_dp if a}
        reais = sorted([r for r in fin0 if r['STATUS_ANALITICO'] == 'ATIVO' and r['ORIGEM_RESULTADO'] == 'INTERFACEAMENTO'
                        and txt(r['ANALITO']) in cat and isinstance(r['DATA_HORA'], dt.datetime)
                        and isinstance(r['RESULTADO'], float)
                        and r['DATA_HORA'].date() < dt.date.today() - dt.timedelta(days=1)],      # so dias fechados
                       key=lambda r: r['DATA_HORA'])
        T = reais[-200]
        an = txt(T['ANALITO'])
        mapeado_urina = any(a == an and (not m or m.upper() == 'URINA') for m, a in linhas_dp)
        esp_m1 = 'ATIVO' if mapeado_urina else 'LOTE_OUTRO_MATERIAL'
        # ADR-093 (decisao do usuario em 13/11/2026): codigo de origem UNICO do analito e sem de/para em URINA -> o manual
        # e aceito nesse codigo (o nome com que o resultado de urina chega do equipamento). Oraculo independente.
        cod_m1 = None
        if not mapeado_urina:
            cods = {o for _, a, o in trios_dp if a == an}
            # = ParaQcini (1a linha que casa; QCINI vazio = nada) e o codigo ja recebido do equipamento em URINA
            recebidos_urina = {txt(r.get('ANALITO_ORIGEM')) for r in fin0 if r.get('ORIGEM_RESULTADO') == 'INTERFACEAMENTO'
                               and txt(r.get('MATRIZ')).upper() == 'URINA'}
            if len(cods) == 1:
                c_ = next(iter(cods))
                pq_ = next((q_ or None for m_, q_, o in trios_dp if o == c_ and (not m_ or m_.upper() == 'URINA')), None)
                if pq_ is None and c_ in recebidos_urina:
                    cod_m1, esp_m1 = c_, 'ATIVO'
        # regra MATRIZ de teste (acrescentada as reais)
        lo_lm = q.lo('tblLotesMaterial')
        ws_lm = lo_lm.Parent
        reais_lm = [[r.get(c) for c in cp.COLUNAS_LOTES_MATERIAL] for r in q.ler('tblLotesMaterial')
                    if any(r.get(c) not in (None, '') for c in ('LOTE', 'REGRA', 'MATRIZ', 'ANALITOS'))]
        todas = reais_lm + [[LOTE_T, 'MATRIZ', 'URINA', None, TAG]]
        xl.EnableEvents = False
        ws_lm.Unprotect(ii.SENHA)
        if lo_lm.ListRows.Count:
            lo_lm.DataBodyRange.ClearContents()
        r0, c0 = lo_lm.HeaderRowRange.Row, lo_lm.HeaderRowRange.Column
        lo_lm.Resize(ws_lm.Range(ws_lm.Cells(r0, c0), ws_lm.Cells(r0 + len(todas), c0 + lo_lm.ListColumns.Count - 1)))
        cab = [str(c.Name) for c in lo_lm.ListColumns]
        ws_lm.Range(ws_lm.Cells(r0 + 1, c0), ws_lm.Cells(r0 + len(todas), c0 + len(cab) - 1)).NumberFormat = '@'
        for i, l in enumerate(todas, 1):
            for nome, v in zip(cp.COLUNAS_LOTES_MATERIAL, l):
                if v is not None:
                    ws_lm.Cells(r0 + i, c0 + cab.index(nome)).Value = v
        d = T['DATA_HORA']
        # REGISTRADO_EM explicito (com os eventos desligados nada carimba): sem ele, "DATA/HORA posterior ao lancamento"
        base = dict(DATA=dt.datetime(d.year, d.month, d.day), EQUIPAMENTO=txt(T['EQUIPAMENTO']), NIVEL=int(T['NIVEL']),
                    REGISTRADO_EM=dt.datetime.combine(d.date() + dt.timedelta(days=1), dt.time(0, 10)))
        casos = {'M1': dict(ID_REGISTRO='MAN_9921', HORA='23:51', LOTE="'" + LOTE_T, ANALITO=an, RESULTADO=12.345),
                 'M2': dict(ID_REGISTRO='MAN_9922', HORA='23:52', LOTE="'" + LOTE_T, ANALITO='URINA QA092', RESULTADO=12.345),
                 'M3': dict(ID_REGISTRO='MAN_9923', HORA='23:53', LOTE="'" + txt(T['LOTE']), ANALITO=an,
                            RESULTADO=float(T['RESULTADO'])),
                 'M4': dict(ID_REGISTRO='MAN_9924', HORA='23:54', LOTE="'" + txt(T['LOTE']), ANALITO=an, RESULTADO=-1)}
        q.lo('tblResultados_Manuais').Parent.Unprotect(ii.SENHA)       # REGISTRADO_EM e coluna travada
        for k, c in casos.items():
            q.escrever_linha('tblResultados_Manuais', dict(base, MOTIVO=f'{TAG} {k}', **c))
        xl.EnableEvents = True
        a1 = aud0()
        r_at = str(q.run('mIntegracao.AtualizarSoEntradas', teto=1500))
        fin1 = {txt(r['ID_REGISTRO']): r for r in q.ler('tblCQ_Final')}
        qa1 = q.ler('tblQA_Integracao')
        m = {k: fin1.get(c['ID_REGISTRO']) for k, c in casos.items()}
        st = {k: (v or {}).get('STATUS_ANALITICO') for k, v in m.items()}
        a13 = {txt(x.get('ID_REGISTRO')) for x in qa1 if x.get('CODIGO') == 'A13'}
        ok1 = (r_at.startswith('OK|') and st['M1'] == esp_m1 and st['M2'] not in ('LOTE_OUTRO_MATERIAL', 'LOTE_FORA_DO_CQ')
               and st['M3'] in ('ATIVO', 'CONFLITO_MANUAL') and st['M4'] == 'IMPLAUSIVEL')
        i05 = {txt(x.get('ID_REGISTRO')): txt(x.get('DETALHE')) for x in qa1 if x.get('CODIGO') == 'I05'}
        if esp_m1 == 'LOTE_OUTRO_MATERIAL':
            ok1 = ok1 and m['M1']['PARTICIPA_ESTATISTICA'] == NAO and m['M1']['TIPO_PLOTAGEM_LJ'] == 'NAO_PLOTAR' \
                and txt(m['M1']['MOTIVO_EXCLUSAO_AUTOMATICA']) == TXT_MZSEM and 'MAN_9921' in a13
        elif cod_m1:                                                     # ADR-093
            mot1 = txt(m['M1']['MOTIVO_EXCLUSAO_AUTOMATICA'])
            ok1 = ok1 and txt(m['M1']['ANALITO']) == cod_m1 and m['M1']['PARTICIPA_ESTATISTICA'] != NAO \
                and txt(m['M1']['MATRIZ']).upper() == 'URINA' and mot1 == TXT_MZCORR % (an, cod_m1) \
                and 'MAN_9921' not in a13 and 'MAN_9921' in i05 and mot1 in i05.get('MAN_9921', '')
        reg('K01 N-01/ADR-093: manual no lote de regra MATRIZ URINA com analito do catálogo -> %s (varredura independente '
            'do de/para; código de origem único = aceito nesse código, com I05); nome fora do catálogo no mesmo lote não '
            'vira LOTE_*; o analito num lote comum participa'
            % (esp_m1 + (f' como {cod_m1}' if cod_m1 else '')), ok1,
            {'analito': an, 'codigo_ADR093': cod_m1, 'analito_M1': txt((m['M1'] or {}).get('ANALITO')),
             'I05_M1': i05.get('MAN_9921'), 'mapeado_em_URINA': mapeado_urina, 'status': st, 'A13_M1': 'MAN_9921' in a13,
                           'motivo_M1': txt((m['M1'] or {}).get('MOTIVO_EXCLUSAO_AUTOMATICA')), 'atualizar': r_at[:80]})
        tr = aud(a1)
        acoes = {k: sorted({a for a, det in tr if c['ID_REGISTRO'] in det and a.startswith('RESULTADO_MANUAL')})
                 for k, c in casos.items()}
        fora = 'RESULTADO_MANUAL_FORA_DA_ESTATISTICA'
        ok3 = fora in acoes['M4'] and 'RESULTADO_MANUAL_INCLUIDO' not in acoes['M4']
        if esp_m1 == 'LOTE_OUTRO_MATERIAL':
            ok3 = ok3 and fora in acoes['M1'] and 'RESULTADO_MANUAL_INCLUIDO' not in acoes['M1']
        reg('K03 N-03: manual IMPLAUSIVEL (e o LOTE_* do K01) -> trilha RESULTADO_MANUAL_FORA_DA_ESTATISTICA, nunca INCLUIDO',
            ok3, {'acoes': acoes})

        # ------------------------------------------------------------ K04 OC-0043
        p = q.ws('Painel')
        nlv = 2 if bio else 3
        j7 = p.Range('J7')
        fonte = j7.DisplayFormat.Font
        larg = j7.MergeArea.Width - 3
        literais = [x for x in re.findall(r'"([^"]*)"', str(j7.Formula)) if x and x.upper() == x and len(x) > 1]
        status = sorted(set(literais) | {'OK', 'ALERTA', 'REJEITADO'})
        med = {s: round(auditoria_abas.largura_pt(s, fonte.Name, fonte.Bold, fonte.Size), 1) for s in status}
        corta = {s: round(larg / w, 2) for s, w in med.items() if w > 0 and larg / w < 0.70}
        q4 = str(p.Range('Q4').Formula)
        q4s = q4.replace(' ', '')
        # cada forma curta mapeada para a sua longa no Q4: COUNTIF(J,"<curta>")>0,"<longa>"
        sem_longo = [c for c, lg in LONGOS.items() if c in literais and f'"{c}")>0,"{lg}"'.replace(' ', '') not in q4s]
        shrink = all(p.Range(f'J{6 + i}').ShrinkToFit is True for i in range(1, nlv + 1))
        fcs = j7.FormatConditions
        cor_sem = [int(fcs(i).Interior.Color) for i in range(1, fcs.Count + 1)
                   if fcs(i).Type == 9 and 'SEM MÉD/DP'.startswith(str(fcs(i).Text))]
        reg('K04 OC-0043: cada status de J7:J%d cabe com redução <= 30 %% (forma curta); "Reduzir para caber" ligado; o '
            'Q4 diz a forma longa; "SEM MÉD/DP" com a cor do sem média' % (6 + nlv),
            not corta and not sem_longo and shrink and len(literais) >= 7 and 0xCEF4FF in cor_sem,
            {'largura_util_pt': round(larg, 1), 'fonte': [fonte.Name, fonte.Bold, fonte.Size], 'medidas_pt': med,
             'cortados': corta, 'sem_forma_longa_no_Q4': sem_longo, 'shrink': shrink, 'cores_SEM_MED': cor_sem})

        e = q.ws('Estatística')
        # ------------------------------------------------------------ K05 OC-0040 (Bio)
        if bio:
            d4f, d4, k3 = str(e.Range('D4').Formula), e.Range('D4').Value, p.Range('K3').Value
            a4 = txt(e.Range('A4').Value)
            xl.EnableEvents = False
            p.Unprotect(ii.SENHA)
            g3, g4, k3f = p.Range('G3').Formula, p.Range('G4').Formula, p.Range('K3').Formula
            try:
                p.Range('G3').Value = None
                p.Range('G4').Value = None
                xl.Calculate()
                per = ' '.join(txt(e.Range(c).Text) for c in ('H5', 'I5', 'J5'))
                p.Range('K3').Value = None                   # De/Ate em anos diferentes: o Painel deixa o K3 vazio
                xl.Calculate()
                d4_vazio = txt(e.Range('D4').Text)
                ano0 = bool(re.search(r'\bAno 0\b', ' '.join(txt(e.Range(c).Text) for c in ('C4', 'D4', 'E4'))))
            finally:
                p.Range('G3').Formula = g3
                p.Range('G4').Formula = g4
                p.Range('K3').Formula = k3f
                xl.Calculate()
                xl.EnableEvents = True
            reg('K05 OC-0040: Estatística D4 = ano do Painel (K3; vazio, nunca "0"), rótulo "Visão (segue o Painel)", '
                'período efetivo vazio = "histórico inteiro do lote"',
                bool(re.search(r'Painel!\$?K\$?3', d4f)) and txt(e.Range('D4').Text) == txt(p.Range('K3').Text)
                and txt(d4) not in ('0', '') and d4_vazio != '0' and not ano0
                and a4 == 'Visão (segue o Painel)' and per.strip() == 'histórico inteiro do lote',
                {'D4': [d4f, d4], 'Painel_K3': k3, 'D4_com_K3_vazio': d4_vazio, 'A4': a4, 'periodo_vazio': per})
        # ------------------------------------------------------------ K06 OC-0049 (Hema)
        else:
            v = e.Range('A14:T133').Value
            vazias = [14 + i for i, l in enumerate(v) if not txt(l[0])]
            sujas = [r for r in vazias if any(txt(e.Cells(r, c).Text) for c in range(1, 21))]
            cheias = [(14 + i, txt(l[1])) for i, l in enumerate(v) if txt(l[0])]
            nivel_ruim = [x for x in cheias if x[1] not in ('1', '2', '3')]
            reg('K06 OC-0049: linhas sem analito com A..T vazias; as linhas com analito mostram o nível',
                vazias and not sujas and cheias and not nivel_ruim,
                {'linhas_sem_analito': len(vazias), 'com_texto': sujas[:6], 'linhas_com_analito': len(cheias),
                 'nivel_ruim': nivel_ruim[:6]})
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
    print(f'\n=== {produto} ADR-092: {len(RES) - len(falhas)} PASS / {len(falhas)} FAIL ===', flush=True)
    return falhas


if __name__ == '__main__':
    falhas = executar(sys.argv[1], os.path.abspath(sys.argv[2]), sys.argv[3] if len(sys.argv) > 3 else None)
    sys.exit(1 if falhas else 0)
