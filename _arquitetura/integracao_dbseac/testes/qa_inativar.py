# -*- coding: utf-8 -*-
"""QA do ADR-070 -- aba Inativar por ID + ANALITO + MOTIVO e ID na dica do Levey-Jennings. Roda numa COPIA instalada.

Uso: python qa_inativar.py <Bioquimica|Hematologia> <copia_instalada.xlsm> [saida.json]

Pelo caminho real: escreve na aba Inativar celula a celula com os eventos ligados (o Worksheet_Change normaliza o
ID e carimba), roda mIntegracao.AtualizarDadosAutomatico (o nucleo do botao ATUALIZAR DADOS) e confere a
Principal - Resultados, o QA_INTEGRACAO, o Eng_Saida (o que o grafico plota) e a dica do grafico (mUI.DicaPonto,
a mesma funcao que a clsCht chama no evento do mouse, com o mesmo (nivel, serie, ponto)). Nada e salvo no arquivo.

  N01 estrutura: a tabela so tem ID | ANALITO | REGISTRAR - LJ | MOTIVO | DATA_INATIVACAO | USUARIO, sem formula;
      ANALITO pela lista do cadastro; data/usuario travados; nas abas Digitar Resultados e COMENTARIOS_TECNICOS
      nenhuma celula de formula mostra o texto "=..." (defeito do formato '@')
  N02 dica de um ponto normal: ID curto, RUN, data/hora e valor = a linha da tblCQ_Final; linha de limite so RUN;
      ponto alem das corridas = ""
  N03 (1) ID (so o numero) + analito certo + motivo: INATIVADO, PARTICIPA = NAO, X vermelho no LJ na corrida certa
  N04 (3) o MOTIVO basta como justificativa (sem comentario tecnico): sem E01; sem motivo e sem comentario: E01
  N05 (2) analito errado: NAO inativa (ATIVO, ponto normal) e o QA acusa E10; sem analito: NAO inativa e E11
  N06 (5) dica do X: "ID ... INATIVADO ... RUN ... data ... valor" do resultado inativado
  N07 duplo clique (nucleo testavel): mIntegracao.RegistrarInativacao grava a linha (ID normalizado, analito,
      caixa marcada, motivo, carimbo), recusa o mesmo ID de novo, e a inativacao vale na atualizacao
  N08 corrigir o analito aplica; apagar o ID reativa e limpa a linha inteira (analito e motivo nao ficam)
  N09 seguranca do teste: arquivo de entrada intacto
  Revisao 07/10/2026:
  N10 analito SEM cadastro (Bio FERR..., Hema RET-HE/IPF...) continua inativavel: validacao de ANALITO em estilo
      Aviso (aceita o nome fora da lista) e o ID + nome de origem inativa na atualizacao, sem E10/E11
  N11 RegistrarInativacao nao reaproveita linha com carimbo residual (data/usuario sem ID): grava numa linha
      totalmente vazia, com caixa marcada e carimbo novo; a linha residual fica como estava
  N12 dica do LJ: a republicacao do motor (AtualizarCalc) apaga o texto da caixa "tip" e muda a geracao das dicas;
      InputBox ANSI: mUI.TextoAnsi troca simbolos fora do cp1252 por texto visivel (">=", "sigma", "delta", "[?]")
"""
import datetime as dt
import json
import os
import sys
import traceback

AQUI = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, AQUI)
from qa_final import QA, reg, RES, NAO, sha256_arquivo  # noqa: E402

SIM = 'SIM'
TB = 'tblInativacao_NaoConformes'
CAB = ['ID_REGISTRO', 'ANALITO', 'REGISTRAR - LJ', 'MOTIVO', 'DATA_INATIVACAO', 'USUARIO']
SEP = ' · '


def curto(i):
    s = str(i)
    p = s.rfind('-')
    return s[p + 1:] if p >= 0 and s[p + 1:].isdigit() else s


def partes(dica):
    return [p.strip() for p in str(dica or '').split(SEP.strip())]


def num_pt(t):
    try:
        return float(str(t).replace('.', '').replace(',', '.')) if ',' in str(t) else float(t)
    except (TypeError, ValueError):
        return None


def confere_dica(dica, r, inativado):
    """A dica bate com a linha r da tblCQ_Final? Devolve (ok, esperado)."""
    dh = r['DATA_HORA']
    esp = ['ID ' + curto(r['ID_REGISTRO'])] + (['INATIVADO'] if inativado else []) + \
          [f"RUN {int(r['RUN'])}", dh.strftime('%d/%m/%Y %H:%M'), str(r['RESULTADO'])]
    p = partes(dica)
    if len(p) != len(esp):
        return False, esp
    ok = p[:-1] == esp[:-1]
    v = num_pt(p[-1])
    ok = ok and v is not None and abs(v - float(r['RESULTADO'])) <= 1e-9 * max(1.0, abs(float(r['RESULTADO'])))
    return ok, esp


def executar(produto, caminho, saida):
    hash0 = sha256_arquivo(caminho)
    q = QA(produto, caminho)
    wb = q.wb
    try:
        # ------------------------------------------------------------ N01 estrutura
        lo = q.lo(TB)
        ws = lo.Parent
        cab = [c.Name for c in lo.ListColumns]
        formulas = [c.Name for c in lo.ListColumns if c.DataBodyRange.Cells(1, 1).HasFormula]
        trav = {c.Name: bool(c.DataBodyRange.Cells(1, 1).Locked) for c in lo.ListColumns}
        try:
            val_an = str(lo.ListColumns('ANALITO').DataBodyRange.Cells(1, 1).Validation.Formula1)
            estilo_an = int(lo.ListColumns('ANALITO').DataBodyRange.Cells(1, 1).Validation.AlertStyle)
        except Exception as e:                      # noqa: BLE001
            val_an = f'sem validacao: {e}'
            estilo_an = None
        try:
            caixa = int(lo.ListColumns('REGISTRAR - LJ').DataBodyRange.Cells(1, 1).CellControl.Type)
        except Exception:                           # noqa: BLE001 -- Excel sem caixa nativa: lista SIM/NAO
            caixa = None
        texto = [str(ws.Range(a).Value or '') for a in ('A3', 'A4')]
        mostra_formula = []
        for nome in ('tblResultados_Manuais', 'tblComentariosTecnicos'):
            l2 = q.lo(nome)
            for c in l2.ListColumns:
                cel = c.DataBodyRange.Cells(1, 1)
                v = cel.Value
                if isinstance(v, str) and v.startswith('='):
                    mostra_formula.append((nome, c.Name, v[:50]))
        ok = (cab == CAB and not formulas and 'lstAnalitos' in val_an
              and trav == {'ID_REGISTRO': False, 'ANALITO': False, 'REGISTRAR - LJ': False, 'MOTIVO': False,
                           'DATA_INATIVACAO': True, 'USUARIO': True}
              and texto[0].startswith('Digite o ID') and 'reativar, apague a linha' in texto[1] and not mostra_formula)
        reg('N01 Aba Inativar enxuta: ID | ANALITO | REGISTRAR - LJ | MOTIVO | DATA_INATIVACAO | USUARIO, sem fórmula, '
            'analito pela lista, carimbo travado; nenhuma célula de fórmula das abas de entrada mostra "=..."', ok,
            {'cabecalho': cab, 'colunas_com_formula': formulas, 'validacao_ANALITO': val_an, 'travadas': trav,
             'caixa_nativa_tipo': caixa, 'instrucoes': texto, 'formula_como_texto': mostra_formula})

        # ------------------------------------------------------------ Painel: analito/lote com dado (como o qa_final T05)
        fin0 = q.final_por_id()
        v_lote = wb.Application.Evaluate('loteAnalise')
        v_lote = v_lote.Value if hasattr(v_lote, 'Value') else v_lote
        lote_an = str(int(v_lote)) if isinstance(v_lote, float) else str(v_lote).strip()
        eq_pai = str(q.ws('Painel').Range('L4').Value) if q.bio else None
        import collections
        cont = collections.Counter(r['ANALITO'] for r in fin0.values() if str(r['LOTE']) == lote_an
                                   and r['PARTICIPA_ESTATISTICA'] == SIM and (eq_pai is None or r['EQUIPAMENTO'] == eq_pai))
        cadastro = [str(v[0] or '').strip() for v in q.ws('Analitos').Range('A4:A43').Value]
        an_top = next((a for a, _ in cont.most_common() if a in cadastro), None)
        dts = [r['DATA'] for r in fin0.values() if str(r['LOTE']) == lote_an and r['PARTICIPA_ESTATISTICA'] == SIM
               and (eq_pai is None or r['EQUIPAMENTO'] == eq_pai) and r['ANALITO'] == an_top]
        if dts:
            q.ex.xl.EnableEvents = False
            pai = q.ws('Painel')
            pai.Unprotect('qcini2025')
            pai.Range('B3').Value = cadastro.index(an_top) + 1
            serie = lambda d: float((d.replace(tzinfo=None) - dt.datetime(1899, 12, 30)).days)   # noqa: E731
            pai.Range('G3').Value = serie(min(dts))
            pai.Range('G4').Value = serie(max(dts))
            q.ex.xl.EnableEvents = True
            q.run('mEstatistica.AtualizarEstatistica')
        e0 = q.eng_saida()
        analito, lote = e0['analito'], e0['lote']

        def resultado_do_slot(e, i, fin):
            c = [r for r in fin.values() if r['ANALITO'] == analito and str(r['LOTE']) == str(lote)
                 and int(r['NIVEL'] or 0) == 1 and int(r['RUN'] or 0) == e['runs'][i] and r['PARTICIPA_ESTATISTICA'] == SIM
                 and (not q.bio or r['EQUIPAMENTO'] == e['equip'])]
            return c[0] if c else None
        slots = []
        for i in range(e0['n'] - 1, -1, -1):
            if e0['fil'][i] == 1 and isinstance(e0['val'][i][0], float):
                r = resultado_do_slot(e0, i, fin0)
                if r is not None and abs(float(r['RESULTADO']) - e0['val'][i][0]) < 1e-12:
                    slots.append((i, r))
            if len(slots) == 6:
                break
        if len(slots) < 6:
            raise RuntimeError(f'pré-condição ausente: menos de 6 pontos N1 visíveis no LJ ({analito}, lote {lote}): '
                               f'{len(slots)}')
        (sA, A), (sB, B), (sC, C), (sD, D), (sE, E), (sN, N) = slots
        outro = next(a for a in cadastro if a and a.upper() != str(analito).upper())
        ch = q.ws('Painel').ChartObjects(1).Chart
        nomes = [str(s.Name) for s in ch.SeriesCollection()]
        s_res = nomes.index('Resultado') + 1
        s_x = [k + 1 for k, n in enumerate(nomes) if 'conforme' in n.lower()]
        dica = lambda nv, s, p: str(q.run('mUI.DicaPonto', nv, s, p, teto=120))   # noqa: E731

        # ------------------------------------------------------------ N02 dica de ponto normal
        dN = dica(1, s_res, sN + 1)
        okN, espN = confere_dica(dN, N, False)
        d_lim = dica(1, 1, sN + 1)
        d_fora = dica(1, s_res, e0['n'] + 1)
        reg('N02 Dica do ponto normal (mUI.DicaPonto): ID curto, RUN real, data/hora e valor = Principal - Resultados; '
            'linha de limite = só a corrida; além das corridas = vazio',
            okN and d_lim.startswith(f"RUN {e0['runs'][sN]}") and d_fora == '',
            {'serie_Resultado': s_res, 'ponto': sN + 1, 'dica': dN, 'esperado': SEP.join(espN), 'dica_limite': d_lim,
             'dica_alem': d_fora, 'ID': N['ID_REGISTRO']})

        # ------------------------------------------------------------ escrever as linhas (evento ligado)
        motA = 'QA ADR-070: repetição (resultado fora do esperado)'
        rA = q.escrever_linha(TB, {'ID_REGISTRO': curto(A['ID_REGISTRO']), 'ANALITO': analito, 'MOTIVO': motA})
        rB = q.escrever_linha(TB, {'ID_REGISTRO': B['ID_REGISTRO'], 'ANALITO': outro, 'MOTIVO': 'QA ADR-070: analito errado'})
        rC = q.escrever_linha(TB, {'ID_REGISTRO': curto(C['ID_REGISTRO']), 'MOTIVO': 'QA ADR-070: sem analito'})
        rD = q.escrever_linha(TB, {'ID_REGISTRO': D['ID_REGISTRO'], 'ANALITO': f'  {str(analito).lower()} '})
        # N10: um resultado de analito SEM cadastro (ANALITO = nome de origem na tblCQ_Final), ATIVO
        sem_cad = next((r for r in fin0.values() if r['ANALITO'] and str(r['ANALITO']).strip() not in cadastro
                        and r['STATUS_ANALITICO'] == 'ATIVO' and r['PARTICIPA_ESTATISTICA'] == SIM
                        and str(r['ID_REGISTRO'] or '').strip()), None)
        rF = None
        if sem_cad is not None:
            rF = q.escrever_linha(TB, {'ID_REGISTRO': curto(sem_cad['ID_REGISTRO']), 'ANALITO': sem_cad['ANALITO'],
                                       'MOTIVO': 'QA revisao: analito sem cadastro'})
        # N11: linha com carimbo RESIDUAL (data e usuario sem ID/analito/motivo), como a migracao ou uma exclusao em
        # varias areas deixam -- a primeira linha "livre" pela regra antiga
        col_id = lo.ListColumns('ID_REGISTRO').DataBodyRange.Value
        r_res = next(i for i, v in enumerate(col_id, start=1) if v[0] in (None, ''))
        data_res = dt.datetime(2001, 1, 1, 8, 0)
        q.ex.xl.EnableEvents = False
        try:
            ws.Unprotect('qcini2025')
        except Exception:                                    # noqa: BLE001
            pass
        lo.ListColumns('DATA_INATIVACAO').DataBodyRange.Cells(r_res, 1).Value = data_res
        lo.ListColumns('USUARIO').DataBodyRange.Cells(r_res, 1).Value = 'QA_RESIDUO'
        q.ex.xl.EnableEvents = True
        residuo0 = {c: lo.ListColumns(c).DataBodyRange.Cells(r_res, 1).Value for c in CAB}
        linhas = {k: {c: lo.ListColumns(c).DataBodyRange.Cells(r, 1).Value for c in CAB}
                  for k, r in (('A', rA), ('B', rB), ('C', rC), ('D', rD))}
        motE = 'QA ADR-070: inativado pelo duplo clique no gráfico'
        regE = str(q.run('mIntegracao.RegistrarInativacao', curto(E['ID_REGISTRO']), analito, motE, teto=120))
        regE2 = str(q.run('mIntegracao.RegistrarInativacao', E['ID_REGISTRO'], analito, motE, teto=120))
        rE = int(regE.split('|')[1]) if regE.startswith('OK|') else None
        linE = {c: lo.ListColumns(c).DataBodyRange.Cells(rE, 1).Value for c in CAB} if rE else {}
        residuo1 = {c: lo.ListColumns(c).DataBodyRange.Cells(r_res, 1).Value for c in CAB}
        _, t1 = q.atualizar()
        fin1 = q.final_por_id()
        qa1 = q.ler('tblQA_Integracao')
        e1 = q.eng_saida()
        cod = lambda c, i: [x for x in qa1 if x['CODIGO'] == c and x['ID_REGISTRO'] == i]   # noqa: E731

        # ------------------------------------------------------------ N03 inativacao que vale
        fA = fin1[A['ID_REGISTRO']]
        slA = e1['runs'].index(int(A['RUN'])) if int(A['RUN']) in e1['runs'] else None
        x_ok = slA is not None and float(A['RESULTADO']) in e1['x'][slA][0:3] and e1['val'][slA][0] in (None, '')
        evento_ok = (linhas['A']['ID_REGISTRO'] == A['ID_REGISTRO'] and linhas['A']['REGISTRAR - LJ'] is True
                     and linhas['A']['DATA_INATIVACAO'] and linhas['A']['USUARIO'])
        reg('N03 (1) Inativar pelo número do ID + analito certo + motivo: INATIVADO, PARTICIPA=NÃO, X vermelho na corrida '
            'certa do LJ (fora da série normal); evento normalizou o ID, marcou a caixa e carimbou',
            fA['STATUS_ANALITICO'] == 'INATIVADO' and fA['PARTICIPA_ESTATISTICA'] == NAO
            and fA['TIPO_PLOTAGEM_LJ'] == 'X_VERMELHO' and fA['INATIVACAO_REGISTRADA'] == SIM and x_ok and evento_ok
            and not cod('E10', A['ID_REGISTRO']) and not cod('E11', A['ID_REGISTRO']),
            {'linha_Inativar': linhas['A'], 'final': {k: fA[k] for k in ('ID_REGISTRO', 'ANALITO', 'STATUS_ANALITICO',
             'PARTICIPA_ESTATISTICA', 'TIPO_PLOTAGEM_LJ', 'INATIVACAO_REGISTRADA', 'RUN', 'RESULTADO')},
             'slot_LJ': slA, 'X_N1': e1['x'][slA][0:3] if slA is not None else None,
             'valor_N1_no_slot': e1['val'][slA][0] if slA is not None else None, 'tempo_atualizacao_s': round(t1, 1)})

        # ------------------------------------------------------------ N04 motivo = justificativa
        fD = fin1[D['ID_REGISTRO']]
        reg('N04 (3) O MOTIVO da linha basta como justificativa (sem comentário técnico): TEM_JUSTIFICATIVA=SIM, '
            'GOVERNANCA OK, sem E01; o inativado sem motivo e sem comentário (analito em minúsculas e com espaços, '
            'aceito) continua com E01',
            fA['TEM_JUSTIFICATIVA'] == SIM and fA['GOVERNANCA'] == 'OK' and fA['MOTIVO_INATIVACAO'] == motA
            and not fA['COMENTARIO_TECNICO'] and not cod('E01', A['ID_REGISTRO'])
            and fD['STATUS_ANALITICO'] == 'INATIVADO' and fD['TEM_JUSTIFICATIVA'] == NAO and bool(cod('E01', D['ID_REGISTRO'])),
            {'A': {k: fA[k] for k in ('MOTIVO_INATIVACAO', 'COMENTARIO_TECNICO', 'TEM_JUSTIFICATIVA', 'GOVERNANCA')},
             'E01_A': cod('E01', A['ID_REGISTRO']), 'D': {k: fD[k] for k in ('STATUS_ANALITICO', 'TEM_JUSTIFICATIVA', 'GOVERNANCA')},
             'linha_D': linhas['D'], 'E01_D': [x['DETALHE'] for x in cod('E01', D['ID_REGISTRO'])]})

        # ------------------------------------------------------------ N05 analito errado / ausente
        fB, fC = fin1[B['ID_REGISTRO']], fin1[C['ID_REGISTRO']]
        slB = e1['runs'].index(int(B['RUN'])) if int(B['RUN']) in e1['runs'] else None
        ativo = lambda f: (f['STATUS_ANALITICO'] == 'ATIVO' and f['PARTICIPA_ESTATISTICA'] == SIM   # noqa: E731
                           and f['TIPO_PLOTAGEM_LJ'] == 'NORMAL' and f['INATIVACAO_REGISTRADA'] == NAO
                           and f['DATA_INATIVACAO'] in (None, '') and not f['MOTIVO_INATIVACAO'])
        e10, e11 = cod('E10', B['ID_REGISTRO']), cod('E11', C['ID_REGISTRO'])
        reg('N05 (2) ID com analito ERRADO não inativa (continua ATIVO, ponto normal no LJ) e o QA acusa E10 (ERRO); '
            'sem analito também não inativa e acusa E11',
            ativo(fB) and ativo(fC) and slB is not None and e1['val'][slB][0] == float(B['RESULTADO'])
            and len(e10) == 1 and e10[0]['SEVERIDADE'] == 'ERRO' and outro in e10[0]['DETALHE']
            and str(analito) in e10[0]['DETALHE'] and len(e11) == 1 and e11[0]['SEVERIDADE'] == 'ERRO'
            and not cod('E10', C['ID_REGISTRO']) and not cod('E11', B['ID_REGISTRO']),
            {'B': {k: fB[k] for k in ('ID_REGISTRO', 'ANALITO', 'STATUS_ANALITICO', 'TIPO_PLOTAGEM_LJ', 'INATIVACAO_REGISTRADA')},
             'analito_digitado_B': outro, 'E10': [x['DETALHE'] for x in e10],
             'C': {k: fC[k] for k in ('ID_REGISTRO', 'STATUS_ANALITICO', 'TIPO_PLOTAGEM_LJ')},
             'E11': [x['DETALHE'] for x in e11], 'ponto_B_no_LJ': e1['val'][slB][0] if slB is not None else None})

        # ------------------------------------------------------------ N06 dica do X
        dX, okX, espX = None, False, None
        if slA is not None:
            k = [i for i in range(3) if e1['x'][slA][i] == float(A['RESULTADO'])]
            if k and len(s_x) >= k[0] + 1:
                dX = dica(1, s_x[k[0]], slA + 1)
                okX, espX = confere_dica(dX, fA, True)
        dN1 = dica(1, s_res, slB + 1) if slB is not None else None
        okN1, _ = confere_dica(dN1, fB, False) if dN1 else (False, None)
        reg('N06 (5) Dica do X vermelho: "ID … · INATIVADO · RUN … · data hora · valor" do resultado inativado; '
            'o ponto normal ao lado continua com a sua dica',
            okX and okN1, {'serie_X': s_x, 'ponto': (slA + 1) if slA is not None else None, 'dica_X': dX,
                            'esperado': SEP.join(espX) if espX else None, 'dica_normal_B': dN1})

        # ------------------------------------------------------------ N07 duplo clique (nucleo)
        fE = fin1[E['ID_REGISTRO']]
        reg('N07 Inativar pelo gráfico (núcleo do duplo clique, mIntegracao.RegistrarInativacao): grava ID normalizado, '
            'analito, caixa marcada, motivo e carimbo; recusa o mesmo ID de novo; vale na atualização',
            regE.startswith('OK|') and regE2.startswith('ERRO|') and linE.get('ID_REGISTRO') == E['ID_REGISTRO']
            and linE.get('ANALITO') == analito and linE.get('REGISTRAR - LJ') is True and linE.get('MOTIVO') == motE
            and linE.get('DATA_INATIVACAO') and linE.get('USUARIO')
            and fE['STATUS_ANALITICO'] == 'INATIVADO' and fE['TIPO_PLOTAGEM_LJ'] == 'X_VERMELHO'
            and fE['MOTIVO_INATIVACAO'] == motE,
            {'retorno': regE, 'retorno_repetido': regE2, 'linha': linE,
             'final': {k: fE[k] for k in ('STATUS_ANALITICO', 'TIPO_PLOTAGEM_LJ', 'MOTIVO_INATIVACAO')}})

        # ------------------------------------------------------------ N10 analito sem cadastro
        if sem_cad is not None:
            fF = fin1.get(sem_cad['ID_REGISTRO'], {})
            okF = (fF.get('STATUS_ANALITICO') == 'INATIVADO' and fF.get('PARTICIPA_ESTATISTICA') == NAO
                   and not cod('E10', sem_cad['ID_REGISTRO']) and not cod('E11', sem_cad['ID_REGISTRO']))
            evF = {'ID': sem_cad['ID_REGISTRO'], 'ANALITO': sem_cad['ANALITO'], 'linha': rF,
                   'final': {k: fF.get(k) for k in ('STATUS_ANALITICO', 'PARTICIPA_ESTATISTICA', 'TIPO_PLOTAGEM_LJ')}}
        else:
            okF, evF = True, {'aviso': 'nenhum resultado ATIVO de analito sem cadastro nesta base: so a validacao conferida'}
        reg('N10 Analito SEM cadastro continua inativável: a validação de ANALITO é Aviso (aceita o nome de origem, ex. FERR, '
            'IPF) e ID + esse nome inativa na atualização, sem E10/E11 (a conferência ID × analito fica no Power Query)',
            estilo_an == 2 and okF, dict(evF, estilo_validacao=estilo_an))

        # ------------------------------------------------------------ N11 linha com carimbo residual
        reg('N11 RegistrarInativacao não reaproveita a linha com carimbo residual (data/usuário sem ID): grava numa linha '
            'TOTALMENTE vazia, com a caixa marcada e carimbo novo; a linha residual continua como estava',
            rE is not None and rE != r_res and residuo1 == residuo0 and linE.get('REGISTRAR - LJ') is True
            and linE.get('USUARIO') != 'QA_RESIDUO' and fE['TIPO_PLOTAGEM_LJ'] == 'X_VERMELHO',
            {'linha_residual': r_res, 'linha_gravada': rE, 'residuo_antes': residuo0, 'residuo_depois': residuo1,
             'linha_E': linE})

        # ------------------------------------------------------------ N08 corrigir o analito; reativar
        lo.ListColumns('ANALITO').DataBodyRange.Cells(rB, 1).Value = analito
        lo.ListColumns('ID_REGISTRO').DataBodyRange.Cells(rA, 1).ClearContents()      # evento: reativacao
        linA2 = {c: lo.ListColumns(c).DataBodyRange.Cells(rA, 1).Value for c in CAB}
        _, t2 = q.atualizar()
        fin2 = q.final_por_id()
        qa2 = q.ler('tblQA_Integracao')
        fB2, fA2 = fin2[B['ID_REGISTRO']], fin2[A['ID_REGISTRO']]
        reg('N08 Corrigido o analito, a inativação passa a valer (E10 some); apagar a célula do ID reativa (ATIVO, mesmo RUN) '
            'e limpa a linha inteira -- analito, caixa, motivo, data e usuário',
            fB2['STATUS_ANALITICO'] == 'INATIVADO' and not [x for x in qa2 if x['CODIGO'] == 'E10' and x['ID_REGISTRO'] == B['ID_REGISTRO']]
            and fA2['STATUS_ANALITICO'] == 'ATIVO' and fA2['TIPO_PLOTAGEM_LJ'] == 'NORMAL' and fA2['RUN'] == A['RUN']
            and all(v in (None, '') for v in linA2.values()),
            {'B_depois': {k: fB2[k] for k in ('STATUS_ANALITICO', 'TIPO_PLOTAGEM_LJ', 'MOTIVO_INATIVACAO')},
             'A_depois': {k: fA2[k] for k in ('STATUS_ANALITICO', 'TIPO_PLOTAGEM_LJ', 'RUN')}, 'linha_A_apos_apagar_ID': linA2,
             'tempo_atualizacao_s': round(t2, 1)})

        # ------------------------------------------------------------ N12 dica limpa na republicacao; InputBox ANSI
        ch1 = q.ws('Painel').ChartObjects(1).Chart
        try:
            tip = ch1.Shapes('tip')
        except Exception:                                    # noqa: BLE001
            tip = ch1.Shapes.AddTextbox(1, 6, 2, 480, 16)
            tip.Name = 'tip'
        tip.TextFrame2.TextRange.Text = 'ID 999999 · RUN 1 · QA dica antiga'
        g0 = q.run('mUI.GeracaoDica', teto=60)
        q.run('mEstatistica.AtualizarCalc')
        txt_tip = str(ch1.Shapes('tip').TextFrame2.TextRange.Text)
        g1 = q.run('mUI.GeracaoDica', teto=60)
        ansi = str(q.run('mUI.TextoAnsi', 'z \u2265 3\u03c3, \u0394 > ET; a\u00e7\u00e3o \u2013 ok \u2603', teto=60))
        ansi_esp = 'z >= 3sigma, delta > ET; a\u00e7\u00e3o \u2013 ok [?]'
        reg('N12 Dica do LJ: a republicação do motor (AtualizarCalc, como na troca de período/lote/analito) apaga o texto '
            'da caixa "tip" e muda a geração das dicas; InputBox ANSI: TextoAnsi troca ≥ σ Δ por texto visível, mantém '
            'acentos e o travessão do cp1252 e marca o resto como [?]',
            txt_tip == '' and isinstance(g1, (int, float)) and isinstance(g0, (int, float)) and g1 > g0 and ansi == ansi_esp,
            {'texto_tip_depois': txt_tip, 'geracao': [g0, g1], 'TextoAnsi': ansi, 'esperado': ansi_esp})
    except Exception as ex_geral:
        reg('N00 Execução interrompida: os testes seguintes a este ponto não rodaram', False,
            {'erro': f'{type(ex_geral).__name__}: {ex_geral}', 'traceback': traceback.format_exc()[-1500:]})
    finally:
        q.fechar(salvar=False)
    hash1 = sha256_arquivo(caminho)
    reg('N09 Segurança do teste: arquivo de entrada intacto (SHA-256 igual antes e depois; nada salvo nele)',
        hash0 == hash1, {'arquivo': caminho, 'sha256_antes': hash0, 'sha256_depois': hash1})
    if saida:
        with open(saida, 'w', encoding='utf-8') as f:
            json.dump(RES, f, ensure_ascii=False, indent=1, default=str)
    falhas = [r for r in RES if r['resultado'] == 'FAIL']
    print(f'\n=== {produto} inativar (ADR-070): {len(RES) - len(falhas)} PASS / {len(falhas)} FAIL ===', flush=True)
    return falhas


if __name__ == '__main__':
    falhas = executar(sys.argv[1], os.path.abspath(sys.argv[2]), sys.argv[3] if len(sys.argv) > 3 else None)
    sys.exit(1 if falhas else 0)
