# -*- coding: utf-8 -*-
"""QA FINAL da nova arquitetura (ADR-057) -- roda numa COPIA ja instalada.

Uso: python qa_final.py <Bioquimica|Hematologia> <copia_instalada.xlsm> [saida.json]

Exercita o sistema PELO CAMINHO REAL: escreve nas tabelas de entrada com os
eventos ligados (o Worksheet_Change carimba), roda mIntegracao.AtualizarDadosCore
(o mesmo nucleo do botao ATUALIZAR DADOS) e confere tblCQ_Final, Eng_Saida
(o que o Levey-Jennings plota), Calc, Estatistica, QA e Audit_Log.

Cada teste devolve PASS/FAIL com a evidencia. Nada e salvo no arquivo, exceto
no teste de fechar/reabrir (que trabalha numa copia propria).
"""
import collections
import json
import math
import os
import shutil
import sys
import threading
import time

AQUI = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(AQUI, '..'))
sys.path.insert(0, os.path.join(AQUI, '..', '..', 'etapa1_multilote'))
import xlh  # noqa: E402
import pqlib  # noqa: E402

RES = []
NAO = 'NÃO'


def reg(teste, ok, evid):
    RES.append({'teste': teste, 'resultado': 'PASS' if ok else 'FAIL', 'evidencia': evid})
    print(('PASS ' if ok else 'FAIL ') + teste + ' -- ' + json.dumps(evid, ensure_ascii=False, default=str)[:600], flush=True)


class QA:
    def __init__(self, produto, caminho):
        self.produto, self.caminho = produto, caminho
        self.bio = produto.startswith('Bio')
        self.nlv = 2 if self.bio else 3
        self.abrir()

    # ------------------------------------------------------------ infraestrutura
    def abrir(self):
        self.ex = xlh.Excel()
        print('EXCEL_PID', self.ex.pid, flush=True)
        self.wb = self.ex.abrir(self.caminho)
        self.ex.xl.EnableEvents = True
        self.ex.xl.Calculation = -4105

    def fechar(self, salvar=False):
        self.ex.fechar(salvar)

    def run(self, macro, *a, teto=1500):
        return self.ex.run("'" + self.wb.Name + "'!" + macro, *a, teto=teto)

    def atualizar(self):
        """O nucleo do botao ATUALIZAR DADOS pela entrada de automacao (sem dialogo).
        'ERRO|...' vira excecao aqui, para o teste que espera falha."""
        t0 = time.time()
        r = str(self.run('mIntegracao.AtualizarDadosAutomatico'))
        if r.startswith('ERRO|'):
            raise RuntimeError(r[5:])
        return r[3:], time.time() - t0

    def lo(self, nome):
        return pqlib.tabela(self.wb, nome)

    def ler(self, nome):
        # datas do COM vem com fuso; o teste compara com datas ingenuas
        out = []
        for r in pqlib.ler_tabela(self.lo(nome)):
            out.append({k: (v.replace(tzinfo=None) if hasattr(v, 'tzinfo') and getattr(v, 'tzinfo', None) else v)
                        for k, v in r.items()})
        return out

    def ws(self, nome):
        return self.wb.Worksheets(nome)

    def nome(self, n):
        return self.wb.Names(n).RefersToRange

    def final_por_id(self):
        return {r['ID_REGISTRO']: r for r in self.ler('tblCQ_Final')}

    def escrever_linha(self, tabela, valores):
        """Escreve na 1a linha vazia da tabela de entrada, celula a celula (dispara o evento)."""
        lo = self.lo(tabela)
        col0 = lo.ListColumns(1).DataBodyRange
        n = lo.ListRows.Count
        vals = col0.Value
        r = None
        for i in range(n):
            v = vals[i][0] if n > 1 else vals
            if v in (None, ''):
                r = i + 1
                break
        assert r, 'tabela de entrada cheia'
        for col, v in valores.items():
            lo.ListColumns(col).DataBodyRange.Cells(r, 1).Value = v
        return r

    def limpar_linha(self, tabela, coluna, valor):
        lo = self.lo(tabela)
        rng = lo.ListColumns(coluna).DataBodyRange
        vals = rng.Value
        for i, v in enumerate(vals, start=1):
            if v[0] == valor:
                rng.Cells(i, 1).ClearContents()       # o evento limpa o resto da linha (reativacao)
                for c in lo.ListColumns:
                    try:
                        cel = c.DataBodyRange.Cells(i, 1)
                        if not cel.HasFormula:
                            cel.ClearContents()
                    except Exception:
                        pass
                return i
        return None

    def eng_saida(self):
        eng = self.ws('Eng_Saida')
        n = int(eng.Range('I1').Value or 0)
        if n == 0:
            return {'n': 0, 'runs': [], 'val': [], 'x': [], 'fil': [], 'flags': []}
        def bloco(r):
            v = r.Value                      # 1 celula -> escalar; 1 linha -> tupla de tuplas
            if not isinstance(v, tuple):
                return [[v]]
            return [list(l) for l in v]
        runs = [int(r[0]) for r in bloco(eng.Range(f'B3:B{2 + n}'))]
        val = bloco(eng.Range(eng.Cells(3, 25), eng.Cells(2 + n, 24 + self.nlv)))
        x = bloco(eng.Range(eng.Cells(3, 30), eng.Cells(2 + n, 38)))
        fil = [r[0] for r in bloco(eng.Range(f'X3:X{2 + n}'))]
        flags = bloco(eng.Range(eng.Cells(3, 3), eng.Cells(2 + n, 2 + 7 * self.nlv)))
        return {'n': n, 'runs': runs, 'val': val, 'x': x, 'fil': fil, 'flags': flags,
                'analito': eng.Range('C1').Value, 'lote': eng.Range('E1').Value, 'equip': eng.Range('S1').Value,
                'stamp': eng.Range('G1').Value}

    def painel_stats(self):
        """engPainel: uma linha por nivel -- n, media, DP, CV (o que o Painel mostra)."""
        rng = self.nome('engPainel').Value
        return [tuple(round(float(c), 9) if isinstance(c, float) else c for c in linha[:5]) for linha in rng]

    def estat_motor(self, analito, nivel):
        """n/media/DP da aba Estatistica para (analito, nivel)."""
        est = self.ws('Estatística')
        a = est.Range('A14:F133').Value
        for linha in a:
            if str(linha[0] or '').strip().upper() == analito.upper() and int(linha[1] or 0) == nivel:
                return linha[2], linha[3], linha[4]
        return None

    def bloco_repeticoes(self):
        est = self.ws('Estatística')
        v = est.Range('A145:B190').Value
        return {str(a).strip(): b for a, b in v if a not in (None, '')}


def stats(vals):
    n = len(vals)
    if n == 0:
        return 0, None, None
    m = sum(vals) / n
    sd = math.sqrt(sum((v - m) ** 2 for v in vals) / (n - 1)) if n > 1 else None
    return n, m, sd


# =============================================================================
def executar(produto, caminho, saida):
    q = QA(produto, caminho)
    try:
        wb = q.wb
        abas = [w.Name for w in wb.Worksheets]
        vbp = [c.Name for c in wb.VBProject.VBComponents]
        # ---------------------------------------------------------------- T01 estrutura / legado
        novas = ['DB', 'DB_ORGANIZADO', 'Digitar Resultados', 'Inativar', 'COMENTARIOS_TECNICOS',
                 'Principal - Resultados', 'QA_INTEGRACAO', 'Cfg_Integracao']
        legado = ['Importar', 'Resultados', 'EQC_Dados', 'DB_Resultados', 'Cfg_Status']
        mods_fora = ['mBanco', 'mImportar', 'mOperacao', 'frmCorrida', 'frmMassa', 'frmExcluir', 'frmConfigEstatistica']
        nomes = []
        for n in wb.Names:
            try:
                nomes.append(n.Name)
            except Exception:
                pass
        reg_cab = [q.ws('Registros').Cells(3, c).Value for c in range(1, 11)]
        ok = (all(a in abas for a in novas) and not any(a in abas for a in legado)
              and not any(m in vbp for m in mods_fora)
              and not any(n in nomes for n in ('regRep1', 'regRep2', 'regRep3', 'rData', 'rValor', 'capBanco'))
              and 'Rep 1' not in reg_cab)
        # um botao por aba de dados, com o MESMO nome da aba, em cada uma das quatro
        quatro = ['DB', 'Inativar', 'Digitar Resultados', 'Principal - Resultados']
        botoes = {}
        for a in quatro:
            if a in abas:
                textos = []
                for sh in q.ws(a).Shapes:
                    try:
                        textos.append(str(sh.TextFrame.Characters().Text).strip())
                    except Exception:
                        pass
                botoes[a] = [t for t in quatro if t in textos]
        ok_bot = all(set(botoes.get(a, [])) == set(quatro) for a in quatro)
        reg('T01b Botões: cada aba de dados tem os botões DB · Inativar · Digitar Resultados · Principal - Resultados '
            '(mesmo nome das abas)', ok_bot, botoes)
        reg('T01 Estrutura: abas novas presentes; LEGADO/IMPORTAR/Resultados/DB_Resultados/Cfg_Status removidas; '
            'Rep 1-3 removidas; módulos legados removidos', ok,
            {'abas_novas': [a for a in novas if a in abas], 'legado_restante': [a for a in legado if a in abas],
             'modulos_legados_restantes': [m for m in mods_fora if m in vbp], 'cabecalho_Registros': reg_cab})

        # ---------------------------------------------------------------- base: atualizacao 1
        resumo, t = q.atualizar()
        rec = q.ler('tblDB_Recebimento')
        fin = q.ler('tblCQ_Final')
        org = q.lo('tblDB_Organizado')
        reg('T02 DB_RECEBIMENTO recebe o DB_SEAC com as colunas do DB_SEAC + ID_REGISTRO + RECEBIDO_EM',
            len(rec) > 0 and len({r['ID_REGISTRO'] for r in rec}) == len(rec),
            {'linhas': len(rec), 'colunas': list(rec[0].keys()) if rec else [], 'ids_unicos': len({r['ID_REGISTRO'] for r in rec})})
        cols_org = [c.Name for c in org.ListColumns]
        reg('T03 DB_ORGANIZADO horizontal (uma coluna por analito)', org.ListRows.Count > 0 and len(cols_org) > 12,
            {'linhas': org.ListRows.Count, 'colunas': cols_org[:12] + ['...'], 'n_colunas': len(cols_org)})
        ids = [r['ID_REGISTRO'] for r in fin]
        obrig = ['ID_REGISTRO', 'DATA', 'RUN', 'NIVEL', 'LOTE', 'ANALITO', 'RESULTADO', 'ORIGEM_RESULTADO', 'STATUS_ANALITICO',
                 'PARTICIPA_ESTATISTICA', 'REGISTRAR_RESULTADO_NO_LJ', 'TIPO_PLOTAGEM_LJ', 'COMENTARIO_TECNICO']
        part = collections.Counter((r['EQUIPAMENTO'], r['LOTE'], r['NIVEL'], r['ANALITO'], r['RUN'])
                                   for r in fin if r['PARTICIPA_ESTATISTICA'] == 'SIM')
        ok = (len(ids) == len(set(ids)) and all(c in fin[0] for c in obrig) and len(fin) >= len(rec)
              and not any(v > 1 for v in part.values()))
        reg('T04 DB_CQ_FINAL: uma linha por resultado, ID único, campos obrigatórios, <=1 participante por corrida/nível',
            ok, {'linhas': len(fin), 'colunas': len(fin[0]), 'ids_unicos': len(set(ids)),
                 'status': dict(collections.Counter(r['STATUS_ANALITICO'] for r in fin)),
                 'plotagem': dict(collections.Counter(r['TIPO_PLOTAGEM_LJ'] for r in fin)),
                 'tempo_atualizacao_s': round(t, 1), 'resumo': resumo})
        base_ids = {r['ID_REGISTRO']: (r['RUN'], r['STATUS_ANALITICO']) for r in fin}

        # ---------------------------------------------------------------- escolha do resultado real (visivel no LJ)
        # periodo do Painel = o periodo em que o lote em analise TEM dado (o lote cadastrado
        # pode ser antigo; os recentes ainda nao estao cadastrados -- ver QA I03)
        v_lote = q.wb.Application.Evaluate('loteAnalise')      # o nome devolve REFERENCIA, nao valor
        v_lote = v_lote.Value if hasattr(v_lote, 'Value') else v_lote
        lote_an = str(int(v_lote)) if isinstance(v_lote, float) else str(v_lote).strip()
        eq_pai = str(q.ws('Painel').Range('L4').Value) if q.bio else None
        # analito em tela = o que mais tem resultado no lote/equipamento em analise (o Painel
        # pode estar num analito cujo controle usa outro lote, ex. PCR na Bioquimica)
        import datetime as dtm
        cont = collections.Counter(r['ANALITO'] for r in fin if str(r['LOTE']) == lote_an
                                   and r['PARTICIPA_ESTATISTICA'] == 'SIM' and (eq_pai is None or r['EQUIPAMENTO'] == eq_pai))
        cadastro = [str(v[0] or '').strip() for v in q.ws('Analitos').Range('A4:A43').Value]
        an_top = next((a for a, _ in cont.most_common() if a in cadastro), None)
        dts_lote = [r['DATA'] for r in fin if str(r['LOTE']) == lote_an and r['PARTICIPA_ESTATISTICA'] == 'SIM'
                    and (eq_pai is None or r['EQUIPAMENTO'] == eq_pai) and r['ANALITO'] == an_top]
        if dts_lote:
            q.ex.xl.EnableEvents = False
            pai = q.ws('Painel')
            pai.Unprotect('qcini2025')
            pai.Range('B3').Value = cadastro.index(an_top) + 1
            # datas como NUMERO DE SERIE: o pywin32 converte datetime ingenuo por fuso (+3 h)
            serie = lambda d: float((d.replace(tzinfo=None) - dtm.datetime(1899, 12, 30)).days)
            pai.Range('G3').Value = serie(min(dts_lote))
            pai.Range('G4').Value = serie(max(dts_lote))
            q.ex.xl.EnableEvents = True
            q.run('mEstatistica.AtualizarEstatistica')
            q.run('RecalcularEstatPeriodo') if q.bio else None
        e0 = q.eng_saida()
        analito, lote = e0['analito'], e0['lote']
        alvo = None
        for i in range(e0['n'] - 1, -1, -1):
            if e0['fil'][i] == 1 and isinstance(e0['val'][i][0], float):
                alvo = (i, e0['runs'][i], e0['val'][i][0])
                break
        cands = [r for r in fin if r['ANALITO'] == analito and str(r['LOTE']) == str(lote) and int(r['NIVEL']) == 1
                 and alvo and int(r['RUN'] or 0) == alvo[1] and r['PARTICIPA_ESTATISTICA'] == 'SIM'
                 and (not q.bio or r['EQUIPAMENTO'] == e0['equip'])]
        reg('T05 Resultado real escolhido no LJ em tela (analito/lote/RUN)', bool(cands),
            {'analito': analito, 'lote': lote, 'equip': e0['equip'], 'slot': alvo, 'candidato': cands[0] if cands else None})
        if not cands:
            raise RuntimeError('sem resultado visivel no Painel para os testes de LJ')
        X = cands[0]
        xid, slot, xval = X['ID_REGISTRO'], alvo[0], alvo[2]
        st0 = q.painel_stats()
        rec_x0 = [r for r in rec if r['ID_REGISTRO'] == xid][0]

        # ---------------------------------------------------------------- ID estavel em varios refreshes
        for k in range(2):
            q.atualizar()
        fin2 = q.final_por_id()
        ok = all(fin2.get(i, {}).get('RUN') == v[0] for i, v in base_ids.items()) and len(fin2) == len(base_ids)
        reg('T06 ID estável: mesmos IDs e mesmo RUN após 3 atualizações', ok,
            {'ids': len(fin2), 'ids_base': len(base_ids), 'exemplo': (xid, fin2[xid]['RUN'])})

        # ---------------------------------------------------------------- inativacao (soft-delete) sem comentario
        q.escrever_linha('tblInativacao_NaoConformes', {'ID_REGISTRO': xid.split('-')[-1]})   # so o numero: o evento normaliza
        lin = [r for r in q.ler('tblInativacao_NaoConformes') if r['ID_REGISTRO']]
        q.atualizar()
        f = q.final_por_id()[xid]
        rec_x1 = [r for r in q.ler('tblDB_Recebimento') if r['ID_REGISTRO'] == xid][0]
        qa = q.ler('tblQA_Integracao')
        e01 = [r for r in qa if r['CODIGO'] == 'E01' and r['ID_REGISTRO'] == xid]
        reg('T07 Inativação por ID: soft-delete (origem intacta, registro continua, PARTICIPA=NÃO), checkbox padrão SIM, '
            'carimbo automático', f['STATUS_ANALITICO'] == 'INATIVADO' and f['PARTICIPA_ESTATISTICA'] == NAO
            and f['REGISTRAR_RESULTADO_NO_LJ'] == 'SIM' and f['TIPO_PLOTAGEM_LJ'] == 'X_VERMELHO'
            and rec_x0 == rec_x1 and lin and lin[0]['REGISTRAR - LJ'] is True and lin[0]['DATA_INATIVACAO'] and lin[0]['USUARIO'],
            {'linha_inativacao': lin[0] if lin else None, 'final': {k: f[k] for k in ('ID_REGISTRO', 'STATUS_ANALITICO',
             'PARTICIPA_ESTATISTICA', 'REGISTRAR_RESULTADO_NO_LJ', 'TIPO_PLOTAGEM_LJ', 'RESULTADO')}, 'origem_identica': rec_x0 == rec_x1})
        reg('T08 Governança: inativado SEM comentário = QA ERRO E01 (FAIL de governança detectado)', bool(e01),
            {'achado': e01[0] if e01 else None})

        # ---------------------------------------------------------------- LJ: X vermelho
        e1 = q.eng_saida()
        st1 = q.painel_stats()
        sl = e1['runs'].index(alvo[1]) if alvo[1] in e1['runs'] else None
        calc = q.ws('Calc')
        xcol = calc.Cells(3 + sl, 24).Value if sl is not None else None          # Calc!X (N1 x1)
        valcol = calc.Cells(3 + sl, 6).Value if sl is not None else None         # Calc!F (N1 valor)
        ok = (sl is not None and e1['x'][sl][0] == xval and e1['val'][sl][0] in (None, '')
              and xcol == xval and valcol in (None, '') and e1['flags'][sl][6] in (None, ''))
        reg('T09 LJ: inativado com REGISTRAR-LJ=SIM aparece como X na posição da corrida, fora da série normal e do Westgard',
            ok, {'slot': sl, 'RUN': alvo[1], 'engXN1': e1['x'][sl][0] if sl is not None else None,
                 'engValN1': e1['val'][sl][0] if sl is not None else None, 'Calc_X': xcol, 'Calc_valor': valcol,
                 'veredicto_Westgard_N1': e1['flags'][sl][6] if sl is not None else None})
        # reconciliacao: TODO resultado elegivel e TODO X da tabela final, na janela, esta no LJ
        def reconciliar(e, fin_atual):
            pos = {r: i for i, r in enumerate(e['runs'])}
            faltam, sobram, x_faltam = [], 0, []
            esperados = set()
            for r in fin_atual.values():
                if r['ANALITO'] != analito or str(r['LOTE']) != str(lote):
                    continue
                if q.bio and r['EQUIPAMENTO'] != e0['equip']:
                    continue
                run = int(r['RUN'] or 0)
                if run not in pos:
                    continue
                i, nv = pos[run], int(r['NIVEL']) - 1
                if r['PARTICIPA_ESTATISTICA'] == 'SIM':
                    esperados.add((i, nv))
                    if e['val'][i][nv] != r['RESULTADO']:
                        faltam.append((r['ID_REGISTRO'], run, r['RESULTADO'], e['val'][i][nv]))
                elif r['TIPO_PLOTAGEM_LJ'] == 'X_VERMELHO':
                    if r['RESULTADO'] not in e['x'][i][nv * 3:nv * 3 + 3]:
                        x_faltam.append((r['ID_REGISTRO'], run, r['RESULTADO']))
            for i, linha in enumerate(e['val']):
                for nv, v in enumerate(linha):
                    if isinstance(v, float) and (i, nv) not in esperados:
                        sobram += 1
            return len(esperados), faltam, sobram, x_faltam
        n_esp, faltam, sobram, x_faltam = reconciliar(e1, q.final_por_id())
        reg('T09b Principal - Resultados -> Levey-Jennings: todo resultado que participa está no ponto da sua corrida/nível, '
            'nenhum ponto sem origem, todo inativado com REGISTRAR-LJ aparece como X',
            n_esp > 0 and not faltam and sobram == 0 and not x_faltam,
            {'pontos_conferidos': n_esp, 'divergentes': faltam[:5], 'pontos_sem_origem': sobram, 'X_ausentes': x_faltam[:5],
             'janela_corridas': e1['n']})
        try:
            png = os.path.join(os.path.dirname(caminho), f'LJ_{produto}_X_vermelho.png')
            q.ws('Painel').ChartObjects(1).Chart.Export(png)
            reg('T09c Imagem do gráfico com o X exportada para conferência visual', os.path.exists(png), {'arquivo': png})
        except Exception as ex_png:
            reg('T09c Imagem do gráfico com o X exportada para conferência visual', False, {'erro': str(ex_png)})
        # estatistica do Painel recalculada independentemente da saida do motor
        vals = [v[0] for v, fl in zip(e1['val'], e1['fil']) if fl == 1 and isinstance(v[0], float)]
        n_ind, m_ind, sd_ind = stats(vals)
        p1 = st1[0]
        ok = (p1[1] == n_ind and abs(p1[2] - m_ind) < 1e-9 and (sd_ind is None or abs(p1[3] - sd_ind) < 1e-9)
              and st0[0][1] == n_ind + 1)
        reg('T10 X sem impacto estatístico: n/média/DP do Painel = população SIM (recalculada à parte); n caiu 1',
            ok, {'antes': st0[0], 'depois': p1, 'independente': (n_ind, m_ind, sd_ind)})

        # serie do grafico: X vermelho, sem linha
        ch = q.ws('Painel').ChartObjects(1).Chart
        sx = [s for s in ch.SeriesCollection() if str(s.Name).startswith('Não conforme')]
        ok = bool(sx) and all(s.MarkerStyle == -4168 and s.Format.Line.Visible == 0 for s in sx)
        cor = sx[0].MarkerForegroundColor if sx else None
        reg('T11 Gráfico: série "Não conforme (X)" com marcador X vermelho e SEM linha ligando à série', ok and cor == 255,
            {'series': [s.Name for s in sx], 'marcador': [s.MarkerStyle for s in sx], 'cor_BGR': cor,
             'tamanho': [s.MarkerSize for s in sx]})

        # ---------------------------------------------------------------- comentario -> QA PASS
        q.escrever_linha('tblComentariosTecnicos', {'ID_REGISTRO': xid,
                                                    'COMENTARIO_TECNICO': 'QA automatizado: resultado fora do esperado, repetido.'})
        q.atualizar()
        f = q.final_por_id()[xid]
        e01b = [r for r in q.ler('tblQA_Integracao') if r['CODIGO'] == 'E01' and r['ID_REGISTRO'] == xid]
        reg('T12 Comentário técnico: com justificativa o QA passa (E01 some), TEM_JUSTIFICATIVA=SIM',
            not e01b and f['TEM_JUSTIFICATIVA'] == 'SIM' and f['GOVERNANCA'] == 'OK',
            {'COMENTARIO_TECNICO': f['COMENTARIO_TECNICO'], 'GOVERNANCA': f['GOVERNANCA']})

        # ---------------------------------------------------------------- REGISTRAR-LJ = NAO
        lo = q.lo('tblInativacao_NaoConformes')
        idx = [i for i, v in enumerate(lo.ListColumns('ID_REGISTRO').DataBodyRange.Value, 1) if v[0] == xid][0]
        lo.ListColumns('REGISTRAR - LJ').DataBodyRange.Cells(idx, 1).Value = False
        q.atualizar()
        f = q.final_por_id()[xid]
        e2 = q.eng_saida()
        st2 = q.painel_stats()
        sl2 = e2['runs'].index(alvo[1]) if alvo[1] in e2['runs'] else None
        ok = (f['TIPO_PLOTAGEM_LJ'] == 'NAO_PLOTAR' and f['PARTICIPA_ESTATISTICA'] == NAO and xid in q.final_por_id()
              and (sl2 is None or e2['x'][sl2][0] in (None, '')) and st2[0] == st1[0])
        reg('T13 REGISTRAR-LJ desmarcado: não plota, não participa, continua auditável; estatística igual ao estado X',
            ok, {'TIPO_PLOTAGEM_LJ': f['TIPO_PLOTAGEM_LJ'], 'slot_no_LJ': sl2, 'painel_X': st1[0], 'painel_nao_plotar': st2[0]})

        # ---------------------------------------------------------------- reabilitacao
        q.limpar_linha('tblInativacao_NaoConformes', 'ID_REGISTRO', xid)
        q.atualizar()
        fin3 = q.final_por_id()
        f = fin3[xid]
        st3 = q.painel_stats()
        e3 = q.eng_saida()
        sl3 = e3['runs'].index(alvo[1])
        ok = (f['STATUS_ANALITICO'] == 'ATIVO' and f['PARTICIPA_ESTATISTICA'] == 'SIM' and f['TIPO_PLOTAGEM_LJ'] == 'NORMAL'
              and len(fin3) == len(base_ids) and f['RUN'] == alvo[1] and e3['val'][sl3][0] == xval and st3[0] == st0[0])
        reg('T14 Reabilitação: apagar o ID volta ATIVO/NORMAL, mesmo ID, sem duplicar, estatística volta ao valor inicial',
            ok, {'final': {k: f[k] for k in ('ID_REGISTRO', 'STATUS_ANALITICO', 'PARTICIPA_ESTATISTICA', 'TIPO_PLOTAGEM_LJ', 'RUN')},
                 'linhas': len(fin3), 'painel_inicial': st0[0], 'painel_reabilitado': st3[0]})

        # ---------------------------------------------------------------- manual MAN_0001
        dts = sorted({r['DATA_HORA'] for r in fin3.values() if r['ANALITO'] == analito and str(r['LOTE']) == str(lote)
                      and (not q.bio or r['EQUIPAMENTO'] == e0['equip'])})
        ult = dts[-1]
        dia = ult.date() if hasattr(ult, 'date') else ult
        import datetime as dtm
        mesmo_dia = [d for d in dts if d.date() == dia]
        hora_livre = dtm.time(23, 59)
        assert all(abs((dtm.datetime.combine(dia, hora_livre) - d).total_seconds()) > 3600 for d in mesmo_dia)
        man_valor = round(xval * 1.01, 2)
        q.escrever_linha('tblResultados_Manuais', {'DATA': dtm.datetime.combine(dia, dtm.time()), 'HORA': 23.0 / 24 + 59.0 / 1440,
                                                   'LOTE': str(lote), 'NIVEL': 1, 'ANALITO': analito, 'RESULTADO': man_valor,
                                                   'MOTIVO': 'Falha do interfaceamento'})
        man = [r for r in q.ler('tblResultados_Manuais') if r['ID_REGISTRO']]
        q.atualizar()
        fin4 = q.final_por_id()
        fm = fin4.get('MAN_0001')
        e4 = q.eng_saida()
        vis = fm is not None and fm['RUN'] in e4['runs'] and e4['val'][e4['runs'].index(fm['RUN'])][0] == man_valor
        reg('T15 Resultado manual MAN_0001: ORIGEM=MANUAL, PARTICIPA=SIM, RUN pela mesma regra, aparece no LJ',
            fm is not None and fm['ORIGEM_RESULTADO'] == 'MANUAL' and fm['PARTICIPA_ESTATISTICA'] == 'SIM' and vis
            and man and man[0]['ID_REGISTRO'] == 'MAN_0001' and man[0]['USUARIO'],
            {'linha_manual': man[0] if man else None, 'final': fm and {k: fm[k] for k in ('ID_REGISTRO', 'ORIGEM_RESULTADO',
             'STATUS_ANALITICO', 'PARTICIPA_ESTATISTICA', 'RUN', 'RESULTADO')}, 'no_LJ': vis})

        # ---------------------------------------------------------------- manual que depois chega pelo interfaceamento
        cand = [r for r in fin4.values() if r['ORIGEM_RESULTADO'] == 'INTERFACEAMENTO' and r['STATUS_ANALITICO'] == 'ATIVO'
                and r['ANALITO'] == analito and str(r['LOTE']) == str(lote) and int(r['NIVEL']) == 1][-5]
        d = cand['DATA_HORA'] + dtm.timedelta(minutes=5)
        q.escrever_linha('tblResultados_Manuais', {'DATA': dtm.datetime.combine(d.date(), dtm.time()),
                                                   'HORA': (d.hour * 60 + d.minute) / 1440.0, 'LOTE': str(lote), 'NIVEL': 1,
                                                   'ANALITO': analito, 'RESULTADO': cand['RESULTADO'],
                                                   'EQUIPAMENTO': cand['EQUIPAMENTO'], 'MOTIVO': 'Resultado não transmitido'})
        q.atualizar()
        fin5 = q.final_por_id()
        f2 = fin5.get('MAN_0002')
        fi = fin5[cand['ID_REGISTRO']]
        a01 = [r for r in q.ler('tblQA_Integracao') if r['CODIGO'] == 'A01' and r['ID_REGISTRO'] == 'MAN_0002']
        part5 = collections.Counter((r['EQUIPAMENTO'], r['LOTE'], r['NIVEL'], r['ANALITO'], r['RUN'])
                                    for r in fin5.values() if r['PARTICIPA_ESTATISTICA'] == 'SIM')
        ok = (f2 and f2['STATUS_ANALITICO'] == 'CONFLITO_MANUAL' and f2['PARTICIPA_ESTATISTICA'] == NAO
              and f2['ID_RELACIONADO'] == cand['ID_REGISTRO'] and fi['STATUS_ANALITICO'] == 'ATIVO' and a01
              and not any(v > 1 for v in part5.values()))
        reg('T16 Duplicidade manual x interfaceamento: chave (equip, matriz, lote, nível, analito, dia, ±30 min) -> '
            'CONFLITO_MANUAL; interfaceamento prevalece; QA A01; nada duplica em silêncio', ok,
            {'manual': f2 and {k: f2[k] for k in ('ID_REGISTRO', 'STATUS_ANALITICO', 'ID_RELACIONADO', 'MOTIVO_EXCLUSAO_AUTOMATICA')},
             'interface': {k: fi[k] for k in ('ID_REGISTRO', 'STATUS_ANALITICO')}, 'alerta': a01[0] if a01 else None})

        # ---------------------------------------------------------------- bloco de repeticoes 3/2/1
        est = q.ws('Estatística')
        lote_est = str(q.nome('Estat_Lote').Value)
        def valor_nome(n):
            try:
                v = q.wb.Names(n).RefersToRange.Value
            except Exception:
                v = q.wb.Application.Evaluate(n)
            if hasattr(v, 'year'):
                v = (v.replace(tzinfo=None) - dtm.datetime(1899, 12, 30)).days
            return float(v or 0)
        ini = valor_nome('Estat_Ini_Efetiva')
        fim_ = valor_nome('Estat_Fim_Efetiva')
        equip = str(q.ws('Painel').Range('L4').Value) if q.bio else None
        nomes_an = ['WBC', 'RBC', 'HGB'] if not q.bio else ['Glicose', 'Ureia', 'Creatinina']
        alvo_q = dict(zip(nomes_an, (3, 2, 1)))

        def no_filtro(r):
            d = r['DATA']
            dserial = (d - dtm.datetime(1899, 12, 30)).days if hasattr(d, 'year') else 0
            return (str(r['LOTE']) == lote_est and (ini == 0 or dserial >= ini) and (fim_ == 0 or dserial < int(fim_) + 1)
                    and (equip is None or r['EQUIPAMENTO'] == equip))
        antes_est = {a: q.estat_motor(a, 1) for a in nomes_an}
        escolhidos = {}
        for a, k in alvo_q.items():
            pool = [r for r in fin5.values() if r['ANALITO'] == a and r['PARTICIPA_ESTATISTICA'] == 'SIM' and int(r['NIVEL']) == 1
                    and r['ORIGEM_RESULTADO'] == 'INTERFACEAMENTO' and no_filtro(r)]
            escolhidos[a] = [r['ID_REGISTRO'] for r in pool[:k]]
            for i in escolhidos[a]:
                q.escrever_linha('tblInativacao_NaoConformes', {'ID_REGISTRO': i})
                q.escrever_linha('tblComentariosTecnicos', {'ID_REGISTRO': i, 'COMENTARIO_TECNICO': 'QA: repetição registrada.'})
        q.atualizar()
        bloco = q.bloco_repeticoes()
        ok = all(bloco.get(a) == k for a, k in alvo_q.items())
        reg(f'T17 Estatística: bloco RESULTADOS NÃO CONFORMES / REPETIÇÕES conta {alvo_q}', ok,
            {'bloco': {a: bloco.get(a) for a in nomes_an}, 'outros_nao_zero': {a: v for a, v in bloco.items() if v and a not in alvo_q},
             'ids': escolhidos, 'filtro': {'lote': lote_est, 'ini': ini, 'fim': fim_, 'equip': equip}})
        # estatistica antes x depois, recalculada independentemente
        fin6 = q.final_por_id()
        res = {}
        okc = True
        for a in nomes_an:
            depois = q.estat_motor(a, 1)
            if not q.bio:
                ano_de = int(est.Range('B3').Value or 0)
                ano_ate = int(est.Range('D3').Value or 0) or ano_de
                pop = [r['RESULTADO'] for r in fin6.values() if r['ANALITO'] == a and int(r['NIVEL']) == 1
                       and str(r['LOTE']) == lote_est and r['PARTICIPA_ESTATISTICA'] == 'SIM'
                       and (ano_de == 0 or ano_de <= r['DATA'].year <= ano_ate)]
            else:
                pop = [r['RESULTADO'] for r in fin6.values() if r['ANALITO'] == a and int(r['NIVEL']) == 1
                       and r['PARTICIPA_ESTATISTICA'] == 'SIM' and no_filtro(r)]
            n_i, m_i, sd_i = stats(pop)
            res[a] = {'antes': antes_est[a], 'depois': depois, 'independente': (n_i, m_i, sd_i)}
            if depois is None or depois[0] in (None, '') or int(depois[0]) != n_i or abs(float(depois[1]) - m_i) > 1e-6:
                okc = False
            if antes_est[a] and antes_est[a][0] not in (None, '') and int(antes_est[a][0]) - int(depois[0]) != alvo_q[a]:
                okc = False
        reg('T18 Estatística usa só PARTICIPA=SIM da DB_CQ_FINAL: n caiu exatamente o número de inativados; '
            'n/média batem com o cálculo independente', okc, res)

        # ---------------------------------------------------------------- idempotencia
        n0 = len(fin6)
        estado0 = {i: (r['RUN'], r['STATUS_ANALITICO'], r['TIPO_PLOTAGEM_LJ'], r['COMENTARIO_TECNICO']) for i, r in fin6.items()}
        tempos = []
        for k in range(4):
            _, tk = q.atualizar()
            tempos.append(round(tk, 1))
        fin7 = q.final_por_id()
        estado1 = {i: (r['RUN'], r['STATUS_ANALITICO'], r['TIPO_PLOTAGEM_LJ'], r['COMENTARIO_TECNICO']) for i, r in fin7.items()}
        reg('T19 Refresh idempotente (4x): mesma contagem, mesmos IDs/RUN, inativações, comentários e manuais mantidos',
            estado0 == estado1 and len(fin7) == n0, {'linhas': [n0, len(fin7)], 'tempos_s': tempos,
                                                    'manuais': [i for i in fin7 if i.startswith('MAN_')]})

        # ---------------------------------------------------------------- falha nao vira "concluida"
        cfg = q.lo('tblConfigIntegracao')
        q.ws('Cfg_Integracao').Unprotect('qcini2025')      # o administrador edita a configuracao
        vals = cfg.DataBodyRange.Value
        lin_c = [i for i, r in enumerate(vals, 1) if r[0] == 'CAMINHO_DB_SEAC'][0]
        cam = cfg.DataBodyRange.Cells(lin_c, 2).Value
        stamp0 = q.ws('Eng_Saida').Range('G1').Value
        cfg.DataBodyRange.Cells(lin_c, 2).Value = cam + '.NAO_EXISTE'
        erro = None
        try:
            q.atualizar()
        except Exception as ex:
            erro = str(ex)
        stamp1 = q.ws('Eng_Saida').Range('G1').Value
        n_falha = q.lo('tblCQ_Final').ListRows.Count
        cfg.DataBodyRange.Cells(lin_c, 2).Value = cam
        reg('T20 Refresh síncrono: falha de camada PARA o processo, não diz "concluída", não atualiza gráfico/estatística',
            erro is not None and 'DB_Recebimento' in erro and stamp0 == stamp1 and n_falha == len(fin7),
            {'erro': (erro or '')[:300], 'carimbo_motor_inalterado': stamp0 == stamp1, 'linhas_final': n_falha})

        # ---------------------------------------------------------------- auditoria
        al = q.ws('Audit_Log')
        ult_l = al.Cells(al.Rows.Count, 1).End(-4162).Row
        acoes = [al.Cells(r, 7).Value for r in range(max(4, ult_l - 60), ult_l + 1)]
        c = collections.Counter(acoes)
        reg('T21 Audit_Log: inativação, reativação, plotagem e manual registrados (ISO 15189 8.4)',
            c.get('RESULTADO_INATIVADO', 0) >= 1 and c.get('RESULTADO_REATIVADO', 0) >= 1 and c.get('PLOTAGEM_LJ_ALTERADA', 0) >= 1
            and any('MANUAL_INCLUIDO' in str(a) for a in acoes), dict(c))

        # ---------------------------------------------------------------- fonte unica (VBA e formulas)
        import re as _re
        proib = ('tblDB_Recebimento', 'tblDB_Organizado', 'tblResultados_Manuais', 'tblInativacao_NaoConformes',
                 'DB_SEAC', 'DB_Resultados', 'Importar', '"Resultados"')
        achados = []
        for comp in wb.VBProject.VBComponents:
            cm = comp.CodeModule
            n = cm.CountOfLines
            if not n:
                continue
            txt = cm.Lines(1, n)
            for i, l in enumerate(txt.split('\r\n'), 1):
                s = l.strip()
                if s.startswith("'"):
                    continue
                for p in proib:
                    if p in l and comp.Name != 'mIntegracao':
                        achados.append(f'{comp.Name}:{i}: {s[:90]}')
        for nm in ('Calc', 'Painel', 'Estatística', 'Eng_Saida', 'Liberação', 'Eventos_Westgard'):
            ur = q.ws(nm).UsedRange
            fs = ur.Formula
            for linha in fs:
                for fcel in linha:
                    if isinstance(fcel, str) and fcel.startswith('=') and any(p.strip('"') in fcel for p in proib):
                        achados.append(f'{nm}: {fcel[:90]}')
        reg('T22 Fonte única: nenhum módulo analítico nem fórmula de LJ/Estatística lê DB_SEAC, DB_RECEBIMENTO, manual, '
            'inativação, Resultados, Importar ou DB_Resultados', not achados, achados[:20])

        # ---------------------------------------------------------------- compilacao e formularios
        modulo = wb.VBProject.VBComponents.Add(1)
        modulo.CodeModule.AddFromString(
            'Public Function QA_Forms() As String\r\n'
            '    Dim f As Object, s As String, nome As Variant\r\n'
            '    For Each nome In Array("frmAssinar", "frmDev")\r\n'
            '        Set f = VBA.UserForms.Add(nome): s = s & nome & ";": Unload f\r\n'
            '    Next nome\r\n'
            '    QA_Forms = s\r\n'
            'End Function\r\n')
        nome_mod = modulo.Name
        forms = q.run(nome_mod + '.QA_Forms', teto=120)
        comp_ok = True
        try:
            ctl = wb.VBProject.VBE.CommandBars.FindControl(1, 578)     # Depurar > Compilar VBAProject
            feito = threading.Event()

            def cao():
                if not feito.wait(120):
                    q.ex.matar()
            threading.Thread(target=cao, daemon=True).start()
            if ctl is not None and ctl.Enabled:
                ctl.Execute()
            feito.set()
        except Exception as ex:
            comp_ok = False
            forms = f'{forms} | compilacao: {ex}'
        wb.VBProject.VBComponents.Remove(wb.VBProject.VBComponents(nome_mod))
        reg('T23 VBA compila (Depurar > Compilar) e os formulários que ficaram abrem', comp_ok and 'frmDev' in str(forms),
            {'formularios': forms})
    finally:
        q.fechar(salvar=False)

    # ---------------------------------------------------------------- fechar e reabrir (copia propria)
    copia = caminho.replace('.xlsm', '_reabrir.xlsm')
    shutil.copy2(caminho, copia)
    q2 = QA(produto, copia)
    try:
        r1, t1 = q2.atualizar()
        q2.fechar(salvar=True)
        q2 = QA(produto, copia)
        r2, t2 = q2.atualizar()
        fin = q2.ler('tblCQ_Final')
        e = q2.eng_saida()
        ok = len(fin) > 0 and e['n'] > 0 and 'concluída' in r2
        reg('T24 Fechar, reabrir e atualizar: camadas, motor, LJ e Estatística funcionando', ok,
            {'linhas_final': len(fin), 'corridas_no_LJ': e['n'], 'tempo_1a_s': round(t1, 1), 'tempo_apos_reabrir_s': round(t2, 1),
             'resumo': r2})
    finally:
        q2.fechar(salvar=False)
    if saida:
        with open(saida, 'w', encoding='utf-8') as f:
            json.dump(RES, f, ensure_ascii=False, indent=1, default=str)
    falhas = [r for r in RES if r['resultado'] == 'FAIL']
    print(f'\n=== {produto}: {len(RES) - len(falhas)} PASS / {len(falhas)} FAIL ===', flush=True)
    return falhas


if __name__ == '__main__':
    falhas = executar(sys.argv[1], os.path.abspath(sys.argv[2]), sys.argv[3] if len(sys.argv) > 3 else None)
    sys.exit(1 if falhas else 0)
