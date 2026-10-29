# -*- coding: utf-8 -*-
"""QA do ADR-085 (29/10/2026): o Painel como tela unica -- duplo clique no ponto com acoes explicitas, comentario
tecnico automatico (fonte unica, sem duplicar), regras de Westgard no ponto, coluna U fora do Painel, Registros pelo ID
com assinatura. Roda numa COPIA instalada; nada e salvo no arquivo.

Uso: python qa_adr085.py <Bioquimica|Hematologia> <copia_instalada.xlsm> <saida.json>

  Z01 (teste 1) identificacao do ponto: a ficha da janela (mUI.FichaPonto) = o registro de origem pelo ID (ID, analito,
      equipamento, nivel, lote, corrida, RUN, data, hora, valor) e a serie/grafico de origem
  Z02 (teste 7) Westgard no ponto: mUI.RegrasNoPonto = as marcas que o motor publicou para aquela corrida e nivel
      (Eng_Saida), em todo ponto com violacao do grafico; a dica (DicaCompleta) ganha a linha "Westgard: ..." so nesses
  Z03 duplo clique: o clsCht instalado cancela o Excel (Cancel = True antes de tudo) no duplo clique e no clique
      direito; fora de um ponto de resultado (eixo, area, legenda, titulo, linha de limite) = so a dica, nenhuma janela
  Z04 (teste 2) NAO CONFORME: X vermelho no lugar do ponto, valor original intacto, comentario na aba
      COMENTARIOS_TECNICOS (ID, texto, analito, nivel, data, usuario, INATIVADO=SIM) e na Principal, trilha, Painel
  Z05 (teste 4) editar o comentario pelo X (acao COMENTAR): a MESMA linha (sem duplicar), texto novo na Principal na
      hora (sem ATUALIZAR, < 3 s), texto anterior no Audit_Log; um ATUALIZAR depois nao muda a linha da Principal
  Z06 (teste 3) X -> sem X (TIRAR_X): sai do grafico, continua INATIVADO, registro preservado
  Z07 (teste 3) INATIVAR direto (acao B) num ponto ativo: some do grafico, registro preservado com valor/ID/motivo
  Z08 so comentario num ponto ativo: estado e grafico iguais (mesmos pontos, n e veredito), comentario gravado
  Z09 (teste 11) recusas sem gravar nada: acao incompativel com o estado, texto vazio
  Z10 (teste 11) reversao: o estado falha depois do comentario gravado -> o comentario volta ao que era
  Z11 (teste 5) depois de cada acao: aba ativa = Painel, filtros (analito, periodo, lote, equipamento) iguais
  Z12 coluna U do Painel sem o historico, nenhum #REF! no Painel, engPainel continua com a coluna 21
  Z13 (teste 8) Registros: o ID preenche a linha (data/hora, analito, RUN, equipamento, valor); assinatura com login e
      senha (texto da identidade + ID em H, trilha); senha errada recusa; ID trocado anula a assinatura; sem ID recusa
  Z14 (teste 10) tempos medidos (comentario so, nao conforme) -- registrados, sem meta inventada
  Z00 seguranca do teste: arquivo de entrada intacto
"""
import json
import os
import re
import sys
import time
import traceback

AQUI = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, AQUI)
from qa_final import QA, reg, RES, sha256_arquivo  # noqa: E402
sys.path.insert(0, os.path.join(AQUI, '..'))
import instalar_integracao as ii  # noqa: E402
from qa_inativar_ux import (criar_usuario, fazer_login, preparar_painel, series, Leitor, x_do_id,  # noqa: E402
                            SENHA_QA, USUARIO_QA)

SEP = ' · '
TB_C = 'tblComentariosTecnicos'
FLAGS = {2: ['1-3s', '2-2s', 'R-4s', '4-1s', '8x'], 3: ['1-3s', '2de3-2s', 'R-4s', '3-1s', '6x']}


def coment(q):
    lo = q.lo(TB_C)
    cab = [c.Name for c in lo.ListColumns]
    if lo.ListRows.Count == 0:
        return []
    v = lo.DataBodyRange.Value
    return [dict(zip(cab, r), _linha=k) for k, r in enumerate(v, start=1)]


def coment_do(q, idf):
    return [r for r in coment(q) if str(r.get('ID_REGISTRO') or '').strip().upper() == idf.upper()]


def filtros(q):
    p = q.ws('Painel')
    return {c: p.Range(c).Value for c in ('B3', 'G3', 'G4', 'H3', 'L4')}


def executar(produto, caminho, saida):
    hash0 = sha256_arquivo(caminho)
    q = QA(produto, caminho)
    L = Leitor(q)
    tempos = {}

    def run(macro, *a, teto=300):
        return str(q.run(macro, *a, teto=teto))

    def acao(resp, nivel, serie, ponto, rotulo):
        time.sleep(2.2)                              # mUI.BLOQUEIO_SEG: clique logo depois e ignorado
        run('mUI.DefinirRespostaTeste', resp, teto=60)
        t0 = time.time()
        r = run('mUI.InativarPontoUI', nivel, serie, ponto, teto=900)
        tempos[rotulo] = round(time.time() - t0, 2)
        return r

    try:
        xl = q.ex.xl
        p = q.ws('Painel')
        nlv = q.nlv
        criar_usuario(q, USUARIO_QA, 'ANALISTA')
        if fazer_login(q, USUARIO_QA).upper() != USUARIO_QA:
            raise RuntimeError('login do usuário de teste falhou')
        preparar_painel(q)
        p.Activate()
        s_res, s_x = series(q, 1)
        base = L.eng_por_run()
        pts = [(rn, d['linha'], d['id'][0], d['cor'][0]) for rn, d in base.items() if d['cor'][0] is not None and d['id'][0]]
        if len(pts) < 10 or s_res is None:
            raise RuntimeError(f'pré-condição: menos de 10 pontos N1 no LJ ({len(pts)}) ou sem série Resultado')
        m = len(pts) // 2
        PA, PB, PC, PD = pts[m], pts[m - 3], pts[m + 3], pts[m - 5]
        fil0 = filtros(q)

        # ------------------------------------------------------------ Z01 identificacao
        ficha = run('mUI.FichaPonto', 1, s_res, PA[1] + 1, teto=60)
        r_a = L.cq({PA[2]})[PA[2]]
        dh = r_a['DATA_HORA']
        val = r_a['RESULTADO']
        idc = PA[2].split('-')[-1]
        esperado = [f'ID {idc}', str(r_a['ANALITO']).strip(), str(r_a['EQUIPAMENTO']).strip(), 'Nível 1',
                    f'Lote {str(r_a["LOTE"]).strip()}', f'Corrida {int(PA[3])}',
                    f'RUN {str(r_a["RUN"]).split(".")[0]}', f'Data {dh:%d/%m/%Y}', f'Hora {dh:%H:%M}',
                    'Gráfico N1', 'série "Resultado"']
        faltam = [e for e in esperado if e not in ficha]
        m_val = re.search(r'Valor ([-0-9,\.E+]+)', ficha)
        v_ficha = float(m_val.group(1).replace(',', '.')) if m_val else None
        reg('Z01 Duplo clique identifica o ponto pelo ID: a ficha da janela traz ID, analito, equipamento, nível, lote, '
            'corrida, RUN, data, hora, valor e o gráfico/série, iguais ao registro de origem (Principal - Resultados)',
            not faltam and v_ficha is not None and abs(v_ficha - float(val)) < 1e-9 and 'Estado atual: ATIVO' in ficha,
            {'ficha': ficha, 'faltando': faltam, 'valor_cq': val, 'valor_ficha': v_ficha})

        # ------------------------------------------------------------ Z02 Westgard no ponto
        eng = q.ws('Eng_Saida')
        n_eng = int(eng.Range('I1').Value or 0)
        flags = eng.Range(eng.Cells(3, 3), eng.Cells(2 + n_eng, 3 + 7 * nlv - 1)).Value if n_eng else ()
        z2, n_viol, n_ok = [], 0, 0
        for t in range(nlv):
            sr, _ = series(q, t + 1)
            if sr is None:
                continue
            for i in range(n_eng):
                fl = flags[i][t * 7:t * 7 + 7]
                ids = eng.Cells(3 + i, 39 + t).Value
                if not ids:
                    continue
                nomes = [FLAGS[2 if nlv == 2 else 3][k] for k in range(5) if fl[k] in (1, 1.0)]
                esp = ' + '.join(nomes) if nomes else ('1-2s (alerta, não rejeita)' if fl[5] in (1, 1.0) else '')
                got = run('mUI.RegrasNoPonto', t + 1, sr, i + 1, teto=60)
                dica = run('mUI.DicaCompleta', t + 1, sr, i + 1, teto=60)
                linha_w = ('\nWestgard: ' + esp) if esp else ''
                ok = got == esp and (dica.endswith(linha_w) if esp else 'Westgard' not in dica)
                if esp:
                    n_viol += 1
                elif n_ok < 3:
                    n_ok += 1
                else:
                    continue
                if not ok or len(z2) < 6:
                    z2.append({'nivel': t + 1, 'slot': i + 1, 'id': ids, 'flags': fl[:6], 'esperado': esp,
                               'obtido': got, 'dica': dica, 'ok': ok})
        reg('Z02 Ponto vermelho mostra as regras de Westgard violadas (todas, pela corrida e nível do motor): '
            'RegrasNoPonto = as marcas publicadas; a dica ganha "Westgard: ..." só nos pontos com violação',
            all(x['ok'] for x in z2) and len(z2) > 0,
            {'pontos_com_violacao': n_viol, 'amostra': z2[:10]})

        # ------------------------------------------------------------ Z03 duplo clique nunca abre o Excel
        cm = q.wb.VBProject.VBComponents('clsCht').CodeModule
        cod = cm.Lines(1, cm.CountOfLines)

        def primeira_instrucao(proc):
            mm = re.search(r'Private Sub ' + proc + r'\(.*?\)\s*\n(.*?)End Sub', cod, re.S)
            if not mm:
                return None
            for ln in mm.group(1).split('\n'):
                s = ln.strip()
                if s and not s.startswith("'") and not s.lower().startswith('dim '):
                    return s
        ins_dc, ins_rc = primeira_instrucao('c_BeforeDoubleClick'), primeira_instrucao('c_BeforeRightClick')
        fora = {}
        nomes1 = [str(s.Name) for s in p.ChartObjects(1).Chart.SeriesCollection()]
        s_lim = nomes1.index('+3s') + 1 if '+3s' in nomes1 else 1
        for nome, args in (('eixo', (21, 1, 0)), ('area_do_grafico', (2, 0, 0)), ('legenda', (24, 0, 0)),
                           ('area_de_plotagem', (19, 0, 0)), ('titulo', (4, 0, 0)),
                           ('linha_de_limite', (3, s_lim, 1)), ('serie_sem_ponto', (3, s_res, -1))):
            fora[nome] = run('mUI.CliqueNoGrafico', 1, *args, teto=60)
        no_ponto = run('mUI.CliqueNoGrafico', 1, 3, s_res, PA[1] + 1, teto=60)
        reg('Z03 Duplo clique/clique direito no gráfico nunca abre "Formatar" nem o menu do Excel: Cancel = True é a 1ª '
            'instrução dos dois eventos; fora de um ponto de resultado só aparece a dica; no ponto, a janela de registrar',
            ins_dc == 'Cancel = True' and ins_rc == 'Cancel = True'
            and all(v.startswith('DICA|') for v in fora.values())
            and no_ponto.startswith('JANELA|'),
            {'1a_instrucao_duplo_clique': ins_dc, '1a_instrucao_clique_direito': ins_rc, 'fora_do_ponto': fora,
             'no_ponto_sem_interface': no_ponto})

        # ------------------------------------------------------------ Z04 NAO CONFORME
        aud0 = L.audit_ult()
        c0 = coment_do(q, PA[2])
        txt4 = 'QA ADR-085: resultado fora do esperado;\nfrasco trocado e repetido (Δ ≥ 2s)'
        r4 = acao('ACAO|NC|' + txt4, 1, s_res, PA[1] + 1, 'nao_conforme_s')
        fil4, aba4 = filtros(q), str(xl.ActiveSheet.Name)
        depois4 = L.eng_por_run()
        rn_x, _, pos_x = x_do_id(depois4, 0, PA[2])
        f4 = L.cq({PA[2]})[PA[2]]
        c4 = coment_do(q, PA[2])
        acoes4 = [a['acao'] for a in L.audit_desde(aud0)]
        reg('Z04 Marcar como NÃO CONFORME: X vermelho no lugar do ponto (mesmo RUN), valor original intacto, comentário '
            'gravado sozinho na aba COMENTARIOS_TECNICOS (ID, texto, analito, nível, data, usuário, INATIVADO=SIM) e na '
            'Principal; trilha COMENTARIO_TECNICO_REGISTRADO + INATIVACAO_PEDIDA',
            r4.startswith('OK|') and 'não conforme' in r4 and rn_x == PA[0] and f4['STATUS_ANALITICO'] == 'INATIVADO'
            and f4['TIPO_PLOTAGEM_LJ'] == 'X_VERMELHO' and f4['RESULTADO'] == r_a['RESULTADO']
            and f4['COMENTARIO_TECNICO'] == txt4 and f4['MOTIVO_INATIVACAO'] == txt4 and f4['TEM_JUSTIFICATIVA'] == 'SIM'
            and len(c4) == len(c0) + (0 if c0 else 1) and len(c4) >= 1 and c4[0]['COMENTARIO_TECNICO'] == txt4
            and str(c4[0]['INATIVADO']) == 'SIM' and str(c4[0]['ANALITO']).strip() == str(r_a['ANALITO']).strip()
            and str(c4[0]['USUARIO'] or '').upper() == USUARIO_QA and c4[0]['DATA'] is not None
            and 'COMENTARIO_TECNICO_REGISTRADO' in acoes4 and 'INATIVACAO_PEDIDA' in acoes4,
            {'retorno': r4[:200], 'x_run': rn_x, 'x_pos': pos_x, 'cq': {k: f4[k] for k in (
                'STATUS_ANALITICO', 'TIPO_PLOTAGEM_LJ', 'RESULTADO', 'COMENTARIO_TECNICO', 'TEM_JUSTIFICATIVA')},
             'linhas_comentario': [{k: str(v) for k, v in r.items()} for r in c4], 'trilha': acoes4})
        z11 = [('nao_conforme', aba4, fil4)]

        # ------------------------------------------------------------ Z05 editar o comentario pelo X
        sx_pt = None
        for sx in s_x:
            for pt in range(1, n_eng + 2):
                d = run('mIntegracao.DadosDoPonto', 1, sx, pt, teto=60)
                if d.startswith('OK|') and d.split('|')[1] == PA[2]:
                    sx_pt = (sx, pt)
                    break
            if sx_pt:
                break
        if not sx_pt:
            raise RuntimeError(f'X do {PA[2]} não encontrado no gráfico')
        aud5 = L.audit_ult()
        txt5 = 'QA ADR-085: revisado -- reagente vencido confirmado'
        r5 = acao('ACAO|COMENTAR|' + txt5, 1, sx_pt[0], sx_pt[1], 'comentario_editado_s')
        f5 = L.cq({PA[2]})[PA[2]]
        c5 = coment_do(q, PA[2])
        aud5l = [a for a in L.audit_desde(aud5)]
        alt = [a for a in aud5l if a['acao'] == 'COMENTARIO_TECNICO_ALTERADO']
        fil5, aba5 = filtros(q), str(xl.ActiveSheet.Name)
        q.run('mIntegracao.AtualizarSoEntradas', teto=900)
        f5b = L.cq({PA[2]})[PA[2]]
        mesma = {k: (f5[k], f5b[k]) for k in f5 if f5[k] != f5b[k] and k not in ('RECEBIDO_EM',)}
        reg('Z05 Editar o comentário pelo X: a MESMA linha (sem duplicar), texto novo na Principal na hora (sem '
            'ATUALIZAR), o anterior no Audit_Log (COMENTARIO_TECNICO_ALTERADO); um ATUALIZAR depois não muda a linha',
            r5.startswith('OK|') and len(c5) == len(c4) and c5[0]['_linha'] == c4[0]['_linha']
            and c5[0]['COMENTARIO_TECNICO'] == txt5 and f5['COMENTARIO_TECNICO'] == txt5 and len(alt) == 1
            and txt4.split(';')[0] in str(alt[0]['parecer'] or '') + str(alt[0]['motivo'] or '') + str(alt)
            and not mesma and tempos['comentario_editado_s'] < 3.0,
            {'retorno': r5[:200], 'linhas': len(c5), 'linha': c5[0]['_linha'] if c5 else None, 'cq': f5['COMENTARIO_TECNICO'],
             'trilha': aud5l, 'diferencas_apos_atualizar': {k: [str(a), str(b)] for k, (a, b) in mesma.items()},
             'tempo_s': tempos['comentario_editado_s']})
        z11.append(('comentario', aba5, fil5))

        # ------------------------------------------------------------ Z06 tirar o X
        time.sleep(0.1)
        r6 = acao('ACAO|TIRAR_X|QA ADR-085: tirar do gráfico', 1, sx_pt[0], sx_pt[1], 'tirar_x_s')
        depois6 = L.eng_por_run()
        rn6, _, _ = x_do_id(depois6, 0, PA[2])
        normal6 = any(d['id'][0] == PA[2] for d in depois6.values())
        f6 = L.cq({PA[2]})[PA[2]]
        reg('Z06 Tirar o X (resultado não conforme -> fora do gráfico): sai do gráfico, continua INATIVADO, registro '
            'original preservado (valor, ID, motivo)',
            r6.startswith('OK|') and rn6 is None and not normal6 and f6['STATUS_ANALITICO'] == 'INATIVADO'
            and f6['TIPO_PLOTAGEM_LJ'] == 'NAO_PLOTAR' and f6['RESULTADO'] == r_a['RESULTADO']
            and f6['MOTIVO_INATIVACAO'] == txt4,
            {'retorno': r6[:200], 'cq': {k: str(f6[k]) for k in ('STATUS_ANALITICO', 'TIPO_PLOTAGEM_LJ', 'RESULTADO',
                                                                  'MOTIVO_INATIVACAO')}})
        z11.append(('tirar_x', str(xl.ActiveSheet.Name), filtros(q)))
        # limpeza do PA (reativar pela aba Inativar, como o usuario faria para um ponto fora do grafico)
        run('mIntegracao.ReativarPorId', PA[2], 'QA ADR-085: limpeza', teto=120)
        q.run('mIntegracao.AtualizarSoEntradas', teto=900)

        # ------------------------------------------------------------ Z07 INATIVAR direto (B)
        base7 = L.eng_por_run()
        lin7 = next(d['linha'] for d in base7.values() if d['id'][0] == PB[2])
        r_b = L.cq({PB[2]})[PB[2]]
        r7 = acao('ACAO|INATIVAR|QA ADR-085: inativar sem X', 1, s_res, lin7 + 1, 'inativar_s')
        depois7 = L.eng_por_run()
        f7 = L.cq({PB[2]})[PB[2]]
        sumiu = x_do_id(depois7, 0, PB[2])[0] is None and not any(d['id'][0] == PB[2] for d in depois7.values())
        reg('Z07 INATIVAR (opção B): some do gráfico; o registro fica com ID, valor, data, usuário e motivo',
            r7.startswith('OK|') and sumiu and f7['STATUS_ANALITICO'] == 'INATIVADO' and f7['TIPO_PLOTAGEM_LJ'] == 'NAO_PLOTAR'
            and f7['RESULTADO'] == r_b['RESULTADO'] and f7['DATA_HORA'] == r_b['DATA_HORA']
            and str(f7['USUARIO_INATIVACAO'] or '').upper() == USUARIO_QA and f7['MOTIVO_INATIVACAO'],
            {'retorno': r7[:200], 'cq': {k: str(f7[k]) for k in ('STATUS_ANALITICO', 'TIPO_PLOTAGEM_LJ', 'RESULTADO',
                                                                  'USUARIO_INATIVACAO', 'MOTIVO_INATIVACAO')}})
        z11.append(('inativar', str(xl.ActiveSheet.Name), filtros(q)))
        run('mIntegracao.ReativarPorId', PB[2], 'QA ADR-085: limpeza', teto=120)
        q.run('mIntegracao.AtualizarSoEntradas', teto=900)

        # ------------------------------------------------------------ Z08 so comentario num ponto ativo
        base8 = L.eng_por_run()
        lin8 = next(d['linha'] for d in base8.values() if d['id'][0] == PC[2])
        st8 = [str(p.Range(f'J{7 + t}').Value) for t in range(nlv)]
        inat8 = L.inativar()
        txt8 = 'QA ADR-085: só observação técnica'
        r8 = acao('ACAO|COMENTAR|' + txt8, 1, s_res, lin8 + 1, 'comentario_s')
        f8 = L.cq({PC[2]})[PC[2]]
        reg('Z08 Só comentário num ponto ativo: estado ATIVO, gráfico, n e veredito iguais, nada na aba Inativar; '
            'comentário na aba COMENTARIOS_TECNICOS e na Principal, sem ATUALIZAR',
            r8.startswith('OK|') and f8['STATUS_ANALITICO'] == 'ATIVO' and L.eng_por_run() == base8
            and [str(p.Range(f'J{7 + t}').Value) for t in range(nlv)] == st8 and L.inativar() == inat8
            and f8['COMENTARIO_TECNICO'] == txt8 and len(coment_do(q, PC[2])) == 1,
            {'retorno': r8[:200], 'tempo_s': tempos['comentario_s'], 'comentario_cq': f8['COMENTARIO_TECNICO']})
        z11.append(('so_comentario', str(xl.ActiveSheet.Name), filtros(q)))

        # ------------------------------------------------------------ Z09 recusas
        cz, iz = coment(q), L.inativar()
        r9a = acao('ACAO|REATIVAR|QA ADR-085: incompatível', 1, s_res, lin8 + 1, 'recusa_a')
        r9b = acao('ACAO|NC|   \n  ', 1, s_res, lin8 + 1, 'recusa_b')
        r9c = acao('ACAO|XYZ|texto', 1, s_res, lin8 + 1, 'recusa_c')
        reg('Z09 Recusas sem gravar nada: reativar um ponto ativo, texto vazio, ação desconhecida',
            r9a.startswith('ERRO|') and r9b.startswith('ERRO|') and r9c.startswith('ERRO|') and coment(q) == cz
            and L.inativar() == iz, {'reativar_ativo': r9a, 'texto_vazio': r9b, 'acao_desconhecida': r9c})

        # ------------------------------------------------------------ Z10 reversao
        base10 = L.eng_por_run()
        lin10 = next(d['linha'] for d in base10.values() if d['id'][0] == PD[2])
        run('mIntegracao.InativarPorId', PD[2], 'QA ADR-085: linha prévia', True, '', teto=120)   # sem ATUALIZAR
        c10 = coment_do(q, PD[2])
        r10 = acao('ACAO|NC|QA ADR-085: deve ser revertido', 1, s_res, lin10 + 1, 'reversao')
        c10b = coment_do(q, PD[2])
        reg('Z10 Reversão: o estado falha depois do comentário gravado -> o comentário volta exatamente ao que era '
            '(nada pela metade)',
            r10.startswith('ERRO|') and c10b == c10, {'retorno': r10, 'antes': [str(x) for x in c10],
                                                      'depois': [str(x) for x in c10b]})
        run('mIntegracao.ReativarPorId', PD[2], 'QA ADR-085: limpeza', teto=120)

        # ------------------------------------------------------------ Z11 Painel e filtros
        reg('Z11 Depois de cada ação (não conforme, comentário, tirar o X, inativar, só comentário) a aba ativa continua o '
            'Painel e os filtros (analito, período, lote, equipamento) são os mesmos',
            all(a == 'Painel' and f == fil0 for _, a, f in z11),
            {'antes': {k: str(v) for k, v in fil0.items()},
             'depois': [(n, a, {k: str(v) for k, v in f.items()}) for n, a, f in z11]})

        # ------------------------------------------------------------ Z12 coluna U
        u = [p.Range(f'U{r}').Formula for r in range(6, 7 + nlv)]
        erros = [c.Address for c in p.UsedRange if isinstance(c.Value, int) and c.Value < -2146820000]
        ep = q.nome('engPainel')
        reg('Z12 Coluna U do Painel sem o "Histórico" (U6 e uma linha por nível vazias), nenhum erro no Painel, '
            'engPainel continua com a coluna 21 (o motor ainda calcula; nada da auditoria apagado)',
            not any(u) and not erros and ep.Columns.Count >= 21,
            {'U6..': u, 'erros': erros[:10], 'engPainel_colunas': ep.Columns.Count})

        # ------------------------------------------------------------ Z13 Registros
        rw = q.ws('Registros')
        id13 = pts[len(pts) // 3][2]
        d13 = run('mIntegracao.DadosDoId', id13, teto=60).split('|')
        r13 = run('mCalibracao.RegistrarCalibracaoPorId', id13, 0, teto=120)
        l13 = int(r13.split('|')[1]) if r13.startswith('OK|') else 0
        row13 = rw.Range(f'B{l13}:L{l13}').Value[0] if l13 else ()
        aud13 = L.audit_ult()
        ra = run('mCalibracao.AssinarRegistroCom', l13, USUARIO_QA, SENHA_QA, teto=120)
        h13 = str(rw.Range(f'H{l13}').Value or '')
        aud13l = [a['acao'] for a in L.audit_desde(aud13)]
        rb = run('mCalibracao.AssinarRegistroCom', l13, USUARIO_QA, 'senha-errada', teto=120)
        h13b = str(rw.Range(f'H{l13}').Value or '')
        # trocar o ID da linha assinada (como o usuario: digitando na coluna K, com eventos)
        outro = pts[len(pts) // 3 + 1][2]
        rw.Unprotect(ii.SENHA)
        aud13c = L.audit_ult()
        rw.Range(f'K{l13}').Value = outro
        h13c = str(rw.Range(f'H{l13}').Value or '')
        i13c = str(rw.Range(f'I{l13}').Value or '')
        aud13cl = [a['acao'] for a in L.audit_desde(aud13c)]
        livre = next(r for r in range(4, 204) if not any(rw.Range(f'B{r}:L{r}').Value[0]))
        rc = run('mCalibracao.AssinarRegistroCom', livre, USUARIO_QA, SENHA_QA, teto=120)
        reg('Z13 Registros: o ID preenche data/hora, analito, RUN, equipamento e o valor; assinatura com login e senha '
            '(identidade + data + ID em H, ASSINATURA_REGISTRO); senha errada recusa; trocar o ID anula a assinatura '
            '(ASSINATURA_ANULADA); linha sem ID não assina',
            r13.startswith('OK|') and bool(row13) and row13[1] == d13[2] and row13[9] == d13[1]
            and row13[10] == d13[4] and len(d13) > 9 and f'= {d13[9]}' in str(row13[8])
            and ra.startswith('OK|') and f'login {USUARIO_QA}' in h13 and h13.endswith(f'ID {d13[1]}')
            and 'ASSINATURA_REGISTRO' in aud13l and rb.startswith('ERRO|') and h13b == h13
            and 'ASSINATURA_ANULADA' in aud13cl and h13c == USUARIO_QA and i13c == '' and rc.startswith('ERRO|'),
            {'calibracao': r13, 'linha': [str(x) for x in row13], 'dados_do_id': d13, 'assinar': ra, 'H': h13,
             'trilha': aud13l, 'senha_errada': rb, 'troca_id': {'H': h13c, 'I': i13c, 'trilha': aud13cl},
             'sem_id': rc})

        # ------------------------------------------------------------ Z14 tempos
        reg('Z14 Tempos medidos (s) das operações do Painel -- registrados como medidos, sem meta inventada',
            True, {'tempos_s': tempos})
    except Exception as e:                           # noqa: BLE001
        reg('Z99 Execução da suíte sem exceção', False, {'erro': repr(e), 'trace': traceback.format_exc()[-2000:]})
    finally:
        q.fechar(salvar=False)
    hash1 = sha256_arquivo(caminho)
    reg('Z00 Segurança do teste: arquivo de entrada intacto (SHA-256 igual antes e depois; nada salvo nele)',
        hash0 == hash1, {'arquivo': caminho, 'sha256_antes': hash0, 'sha256_depois': hash1})
    if saida:
        with open(saida, 'w', encoding='utf-8') as f:
            json.dump(RES, f, ensure_ascii=False, indent=1, default=str)
    falhas = [r for r in RES if r['resultado'] == 'FAIL']
    print(f'\n=== {produto} ADR-085 (painel tela única): {len(RES) - len(falhas)} PASS / {len(falhas)} FAIL ===', flush=True)
    return falhas


if __name__ == '__main__':
    falhas = executar(sys.argv[1], os.path.abspath(sys.argv[2]), sys.argv[3] if len(sys.argv) > 3 else None)
    sys.exit(1 if falhas else 0)
