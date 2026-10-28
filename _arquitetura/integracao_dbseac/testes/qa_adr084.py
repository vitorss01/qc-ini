# -*- coding: utf-8 -*-
"""QA do ADR-084 (28/10/2026): status sem "alvo TESTE", motivo livre na inativacao pelo grafico e o indicador de
processamento (mCarregando + frmCarregando). Roda numa COPIA instalada; nada e salvo no arquivo.

Uso: python qa_adr084.py <Bioquimica|Hematologia> <copia_instalada.xlsm> <saida.json>

  K01 Painel: nenhum status (J7..), nem o Q4, nem o A39 dizem "TESTE"; nenhuma formula deles tem "alvo TESTE"
  K02 a regra nao mudou: com o motor no lote/analito/equipamento da tela, o status de cada nivel = o veredito do motor
      (engPainel coluna 10) exatamente -- sem sufixo
  K03 frmInativar: campo livre txtMotivo (varias linhas, Enter quebra a linha, ate 1000 caracteres); a lista
      (cboMotivo) e o complemento sairam
  K04 frmCarregando instalado (titulo, etapa, trilho, barra, passo) e o modulo mCarregando
  K05 ATUALIZAR com o indicador aberto (automacao): as 6 etapas na ordem, "Etapa k de 6" nunca passa do total, fechado
      ao fim; a janela NAO abre com Interactive = False; estado do Excel devolvido (calculo automatico, eventos e tela)
  K06 ATUALIZAR que FALHA (DB_SEAC inexistente): devolve ERRO, o indicador fecha e o estado do Excel volta
  K07 com Interactive = True a janela aparece (sem modal), mostra a etapa e some no Fechar; um 2o Abrir nao cria
      outra janela
  K08 inativar pelo grafico (fachada mUI.InativarPontoUI com a resposta do teste) com motivo LIVRE de varias linhas:
      gravado igual na aba Inativar, o indicador passou por "Gravando" + as 6 etapas e fechou, a aba ativa continua
      o Painel; reativar pelo mesmo caminho
  K09 Cancelar: nada gravado, o indicador nem abre
  K00 seguranca do teste: arquivo de entrada intacto
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
from qa_inativar_ux import criar_usuario, fazer_login, preparar_painel, series, Leitor  # noqa: E402

USUARIO_QA = 'QA_INATIVAR_UX'           # o mesmo usuario de teste do qa_inativar_ux (senha de teste de la)
ETAPAS = ['Preparando as entradas', None, 'Conferindo os resultados', 'Atualizando Levey-Jennings',
          'Conferindo os vereditos de Westgard', 'Recalculando a estat']


def etapas_ok(hist, titulo, extra_inicio=()):
    """ABRIR:<titulo>|<extra...>|<6 etapas do nucleo>|FECHAR."""
    p = hist.split('|')
    if not p or p[0] != 'ABRIR:' + titulo or p[-1] != 'FECHAR':
        return False
    meio = p[1:-1]
    if list(meio[:len(extra_inicio)]) != list(extra_inicio):
        return False
    meio = meio[len(extra_inicio):]
    if len(meio) != len(ETAPAS):
        return False
    return all(e is None or m.startswith(e) for m, e in zip(meio, ETAPAS))


def estado_excel(q):
    xl = q.ex.xl
    return {'calculo_auto': xl.Calculation == -4105, 'eventos': bool(xl.EnableEvents),
            'tela': bool(xl.ScreenUpdating), 'barra_status_livre': xl.StatusBar in (False, None, '', 'FALSE')}   # o COM devolve 'FALSE' depois do VBA


def executar(produto, caminho, saida):
    hash0 = sha256_arquivo(caminho)
    q = QA(produto, caminho)
    L = Leitor(q)

    def run(macro, *a, teto=300):
        return str(q.run(macro, *a, teto=teto))

    try:
        p = q.ws('Painel')
        nlv = q.nlv
        # ------------------------------------------------------------ K01 / K02 status
        preparar_painel(q)
        cels = [f'J{6 + i}' for i in range(1, nlv + 1)] + ['Q4', 'A39']
        txts = {c: str(p.Range(c).Value or '') for c in cels}
        forms = {c: str(p.Range(c).Formula or '') for c in cels}
        reg('K01 Painel: status (J7..), Q4 e A39 sem "TESTE" no texto e nas fórmulas',
            not any('TESTE' in t.upper() for t in txts.values()) and not any('alvo TESTE' in f for f in forms.values()),
            {'textos': txts})
        eng = q.nome('engPainel').Value
        normais = {}
        for i in range(1, nlv + 1):
            ver = str(eng[i - 1][9] or '')
            st = txts[f'J{6 + i}']
            normais[f'N{i}'] = {'motor': ver, 'status': st}
        k02 = all(v['status'] == v['motor'] for v in normais.values() if v['motor'] in ('OK', 'REJEITADO'))
        reg('K02 Regra preservada: status de cada nível = veredito do motor (engPainel col. 10), sem sufixo',
            k02 and any(v['motor'] in ('OK', 'REJEITADO') for v in normais.values()), normais)

        # ------------------------------------------------------------ K03 / K04 janelas
        vbp = q.wb.VBProject
        comps = {c.Name: c for c in vbp.VBComponents}
        ctl = {}
        if 'frmInativar' in comps:
            try:
                comps['frmInativar'].Designer
            except Exception:                        # noqa: BLE001
                pass
            ctl = {c.Name: c for c in comps['frmInativar'].Designer.Controls}
        tm = ctl.get('txtMotivo')
        k03 = (tm is not None and 'cboMotivo' not in ctl and 'txtComplemento' not in ctl and bool(tm.MultiLine)
               and bool(tm.EnterKeyBehavior) and int(tm.MaxLength) == 1000)
        reg('K03 frmInativar: motivo em texto livre (txtMotivo, várias linhas, Enter quebra a linha, até 1000); '
            'a lista de frases e o complemento saíram', k03, {'controles': sorted(ctl)})
        cc = {}
        if 'frmCarregando' in comps:
            try:
                comps['frmCarregando'].Designer
            except Exception:                        # noqa: BLE001
                pass
            cc = {c.Name for c in comps['frmCarregando'].Designer.Controls}
        reg('K04 frmCarregando (título, etapa, trilho, barra, passo) e mCarregando instalados',
            {'lblTitulo', 'lblEtapa', 'lblTrilho', 'lblBarra', 'lblPasso'} <= cc and 'mCarregando' in comps,
            {'controles': sorted(cc), 'modulo': 'mCarregando' in comps})

        # ------------------------------------------------------------ K05 ATUALIZAR com o indicador (automacao)
        run('mCarregando.Abrir', 'Atualizando dados', 6, teto=60)
        t0 = time.time()
        r5 = run('mIntegracao.AtualizarDadosAutomatico', teto=1500)
        t5 = round(time.time() - t0, 1)
        h5 = run('mCarregando.Historico', teto=60)
        ab5 = run('mCarregando.Aberto', teto=60)
        vis5 = run('mCarregando.JanelaVisivel', teto=60)
        est5 = estado_excel(q)
        reg('K05 ATUALIZAR com o indicador: as 6 etapas reais na ordem, fechado ao fim, sem janela em automação, '
            'estado do Excel devolvido',
            r5.startswith('OK|') and etapas_ok(h5, 'Atualizando dados') and ab5 == 'False' and vis5 == 'False'
            and all(est5.values()),
            {'retorno': r5[:120], 'historico': h5, 'aberto_depois': ab5, 'janela': vis5, 'estado': est5, 'tempo_s': t5})

        # ------------------------------------------------------------ K06 ATUALIZAR que falha
        modo0 = q.cfg('MODO_FONTE')
        cam0 = q.cfg('CAMINHO_DB_SEAC')
        try:
            q.cfg('MODO_FONTE', 'SEAC')
            q.cfg('CAMINHO_DB_SEAC', r'C:\QA_ADR084_nao_existe\DB_SEAC_inexistente.xlsx')
            run('mCarregando.Abrir', 'Atualizando dados', 6, teto=60)
            r6 = run('mIntegracao.AtualizarDadosAutomatico', teto=900)
            h6 = run('mCarregando.Historico', teto=60)
            ab6 = run('mCarregando.Aberto', teto=60)
            est6 = estado_excel(q)
        finally:
            q.cfg('MODO_FONTE', modo0)
            q.cfg('CAMINHO_DB_SEAC', cam0)
        reg('K06 ATUALIZAR com o DB_SEAC inexistente: ERRO devolvido, indicador fechado, estado do Excel devolvido',
            r6.startswith('ERRO|') and ab6 == 'False' and h6.endswith('|FECHAR') and all(est6.values()),
            {'retorno': r6[:200], 'historico': h6, 'aberto_depois': ab6, 'estado': est6})

        # ------------------------------------------------------------ K07 janela de verdade (Interactive = True)
        xl = q.ex.xl
        vis_ant = xl.Visible
        try:
            xl.Visible = True
            xl.Interactive = True
            run('mCarregando.Abrir', 'QA ADR-084', 3, teto=60)
            run('mCarregando.Abrir', 'QA ADR-084 (2a chamada)', 3, teto=60)     # nao cria outra janela
            run('mCarregando.Etapa', 'Primeira etapa', teto=60)
            v7a = run('mCarregando.JanelaVisivel', teto=60)
            h7 = run('mCarregando.Historico', teto=60)
            run('mCarregando.Fechar', teto=60)
            v7b = run('mCarregando.JanelaVisivel', teto=60)
        finally:
            xl.Interactive = False
            xl.Visible = vis_ant
        reg('K07 Com Interactive = True a janela aparece sem modal, mostra a etapa e some ao fechar; 2º Abrir não '
            'abre outra janela',
            v7a == 'True' and v7b == 'False' and h7 == 'ABRIR:QA ADR-084|Primeira etapa',
            {'visivel_durante': v7a, 'visivel_depois': v7b, 'historico': h7})

        # ------------------------------------------------------------ K08 / K09 inativar com motivo livre
        criar_usuario(q, USUARIO_QA, 'ANALISTA')
        if fazer_login(q, USUARIO_QA).upper() != USUARIO_QA:
            raise RuntimeError('login do usuário de teste falhou')
        preparar_painel(q)
        p.Activate()
        s_res, _ = series(q, 1)
        base = L.eng_por_run()
        pts = [(rn, d['linha'], d['id'][0]) for rn, d in base.items() if d['cor'][0] is not None and d['id'][0]]
        if len(pts) < 4:
            raise RuntimeError(f'pré-condição: menos de 4 pontos N1 no LJ ({len(pts)})')
        alvo = pts[len(pts) // 2]
        inat0 = L.inativar()
        run('mUI.DefinirRespostaTeste', 'CANCELAR', teto=60)
        h9a = run('mCarregando.Historico', teto=60)
        r9 = run('mUI.InativarPontoUI', 1, s_res, alvo[1] + 1, teto=300)
        h9 = run('mCarregando.Historico', teto=60)
        inat9 = L.inativar()
        reg('K09 Cancelar: "CANCELADO", nada gravado na aba Inativar e o indicador nem abre',
            r9.startswith('CANCELADO|') and inat9 == inat0 and h9 == h9a, {'retorno': r9, 'historico': h9})

        time.sleep(2.2)                                  # mUI.BLOQUEIO_SEG: clique logo depois e ignorado
        mot = 'QA ADR-084: controle reconstituído às 7h40;\nrepetido após troca do frasco (ação — Δ ≥ 2s)'
        run('mUI.DefinirRespostaTeste', 'CONFIRMAR|' + mot + '|1', teto=60)
        aba0 = str(xl.ActiveSheet.Name)
        r8 = run('mUI.InativarPontoUI', 1, s_res, alvo[1] + 1, teto=900)
        aba1 = str(xl.ActiveSheet.Name)
        h8 = run('mCarregando.Historico', teto=60)
        inat8 = L.inativar()
        lin8 = next((x for x in inat8.values() if str(x['ID_REGISTRO']) == alvo[2]), {})
        f8 = L.cq({alvo[2]}).get(alvo[2], {})
        idc = alvo[2].split('-')[-1]
        reg('K08 Inativar pelo gráfico com motivo livre de 2 linhas: gravado igual (aba Inativar e Principal), '
            'indicador "Gravando" + 6 etapas e fechado, a aba ativa continua o Painel',
            r8.startswith('OK|') and lin8.get('MOTIVO') == mot and f8.get('MOTIVO_INATIVACAO') == mot
            and etapas_ok(h8, f'Inativando o resultado ID {idc}', ('Gravando na aba Inativar',))
            and aba0 == 'Painel' and aba1 == 'Painel',
            {'retorno': r8[:160], 'motivo_aba': lin8.get('MOTIVO'), 'motivo_principal': f8.get('MOTIVO_INATIVACAO'),
             'historico': h8, 'aba_antes': aba0, 'aba_depois': aba1})
        # limpeza pelo mesmo caminho (reativar), tambem conferida
        time.sleep(2.2)
        base8 = L.eng_por_run()
        alvo_x = None
        _, s_x = series(q, 1)
        for sx in s_x:
            for pt in range(1, 181):
                d = run('mIntegracao.DadosDoPonto', 1, sx, pt, teto=60)
                if d.startswith('OK|') and d.split('|')[1] == alvo[2]:
                    alvo_x = (sx, pt)
                    break
            if alvo_x:
                break
        if not alvo_x:
            raise RuntimeError(f'X do {alvo[2]} não encontrado para reativar (eixo {len(base8)})')
        run('mUI.DefinirRespostaTeste', 'CONFIRMAR|QA ADR-084: limpeza do teste|1', teto=60)
        r8b = run('mUI.InativarPontoUI', 1, alvo_x[0], alvo_x[1], teto=900)
        h8b = run('mCarregando.Historico', teto=60)
        f8b = L.cq({alvo[2]}).get(alvo[2], {})
        reg('K08b Reativar pelo mesmo caminho: ATIVO de novo e o indicador passou pelas etapas e fechou',
            r8b.startswith('OK|') and f8b.get('STATUS_ANALITICO') == 'ATIVO'
            and etapas_ok(h8b, f'Reativando o resultado ID {idc}', ('Gravando na aba Inativar',)),
            {'retorno': r8b[:160], 'historico': h8b, 'estado': f8b.get('STATUS_ANALITICO')})
        # ------------------------------------------------------------ K10..K14 calibracao pelo ID
        reg_ws = q.ws('Registros')
        calc = q.ws('Calc')
        linha3 = calc.Range('A3:EZ3').Formula[0]
        cols_cal = [k + 1 for k, f in enumerate(linha3) if 'regCalib' in str(f or '')]
        alvo_c = pts[len(pts) // 3]                   # outro ponto do grafico (analito e equipamento da tela)
        id_c = alvo_c[2]
        d_c = run('mIntegracao.DadosDoId', id_c, teto=60).split('|')
        aud_c0 = L.audit_ult()
        reg_antes = reg_ws.Range('B4:L203').Value
        r10 = run('mCalibracao.RegistrarCalibracaoPorId', id_c, 0, teto=120)
        lin10 = int(r10.split('|')[1]) if r10.startswith('OK|') else 0
        row10 = reg_ws.Range(f'B{lin10}:L{lin10}').Value[0] if lin10 else ()
        q.ex.xl.Calculate()
        dia = int(float(d_c[7].replace(',', '.'))) if len(d_c) > 7 else -1
        cv = calc.Range('C3:D182').Value
        rows_dia = [i for i, (c_, d_) in enumerate(cv) if isinstance(c_, (int, float)) and int(c_) == dia and d_ == 1]
        if not rows_dia:                                # datas do COM como datetime
            import datetime as _dt
            rows_dia = [i for i, (c_, d_) in enumerate(cv) if hasattr(c_, 'year') and d_ == 1 and
                        (c_.replace(tzinfo=None) - _dt.datetime(1899, 12, 30)).days == dia]
        medias = []
        icones = []
        for k in cols_cal:
            fm = re.search(r'\$([A-Z]+)\$1\b', str(calc.Cells(3, k).Formula))
            media = calc.Range(fm.group(0).replace('$', '')).Value if fm else None
            medias.append(media)
            icones.append([calc.Cells(3 + i, k).Value for i in rows_dia])
        k10_graf = bool(rows_dia) and all(all(isinstance(v, float) and media is not None and abs(v - media) < 1e-9
                                              for v in ic) for ic, media in zip(icones, medias))
        reg('K10 Calibração pelo ID: linha preenchida (data/hora, analito, Todos, RUN, Sim, ID, equipamento) e o ícone '
            'na MÉDIA do lote de cada nível, nas corridas desse dia (mesma regra da marcação manual)',
            r10.startswith('OK|') and row10 and row10[1] == d_c[2] and row10[2] == 'Todos' and row10[4] == 'Sim'
            and row10[9] == d_c[1] and row10[10] == d_c[4] and str(row10[3]).removesuffix('.0') == d_c[6] and k10_graf,
            {'retorno': r10, 'linha': row10, 'dados_do_id': d_c, 'corridas_do_dia': len(rows_dia),
             'icone_y': icones, 'media_do_lote': medias})
        # equipamento: outro equipamento na linha = o icone sai deste grafico
        q.ex.xl.EnableEvents = False
        reg_ws.Unprotect('qcini2025')
        reg_ws.Range(f'L{lin10}').Value = 'QA_OUTRO_EQUIPAMENTO'
        q.ex.xl.Calculate()
        some = all(not isinstance(calc.Cells(3 + i, k).Value, float) for k in cols_cal for i in rows_dia)
        reg_ws.Range(f'L{lin10}').Value = d_c[4]
        q.ex.xl.EnableEvents = True
        q.ex.xl.Calculate()
        volta = all(isinstance(calc.Cells(3 + i, k).Value, float) for k in cols_cal for i in rows_dia)
        reg('K11 Equipamento: com outro equipamento na linha o ícone some deste gráfico; com o do resultado, volta '
            '(linha sem equipamento = todos, como antes)', bool(rows_dia) and some and volta,
            {'sumiu_com_outro': some, 'voltou': volta})
        r12 = run('mCalibracao.RegistrarCalibracaoPorId', id_c.split('-')[-1], 0, teto=120)
        r12b = run('mCalibracao.RegistrarCalibracaoPorId', 'BIO-999999999' if q.bio else 'HEM-999999999', 0, teto=120)
        r12c = run('mCalibracao.RegistrarCalibracaoPorId', 'XYZ-12', 0, teto=120)
        r12d = run('mCalibracao.RegistrarCalibracaoPorId', '   ', 0, teto=120)
        reg_dep = reg_ws.Range('B4:L203').Value
        novas = sum(1 for a_, b_ in zip(reg_antes, reg_dep) if a_ != b_)
        reg('K12 Recusas sem gravar: ID repetido (só o número), ID inexistente, formato inválido e vazio -- só a '
            'linha do K10 mudou na aba',
            'já está registrado' in r12 and r12.startswith('ERRO|') and 'não encontrado' in r12b
            and 'formato inválido' in r12c and r12d.startswith('ERRO|') and novas == 1,
            {'repetido': r12, 'inexistente': r12b, 'formato': r12c, 'vazio': r12d, 'linhas_mudadas': novas})
        # digitado na celula (Worksheet_Change da aba Registros): outro ID valido e um invalido
        alvo_d = pts[len(pts) // 4]
        livre = next(r_ for r_ in range(4, 204) if not any(reg_ws.Range(f'B{r_}:L{r_}').Value[0]))
        reg_ws.Range(f'K{livre}').Value = alvo_d[2].split('-')[-1]
        ok13 = str(reg_ws.Range(f'K{livre}').Value) == alvo_d[2] and reg_ws.Range(f'F{livre}').Value == 'Sim'
        livre2 = next(r_ for r_ in range(4, 204) if not any(reg_ws.Range(f'B{r_}:L{r_}').Value[0]))
        reg_ws.Range(f'K{livre2}').Value = 'XYZ-12'
        ok13b = not reg_ws.Range(f'K{livre2}').Value and not reg_ws.Range(f'B{livre2}').Value
        reg('K13 Digitado na coluna K: a linha se preenche sozinha; ID inválido digitado volta vazio (nada gravado)',
            ok13 and ok13b, {'linha': livre, 'K': reg_ws.Range(f'K{livre}').Value, 'F': reg_ws.Range(f'F{livre}').Value,
                             'invalido_limpo': ok13b})
        a14 = L.audit_desde(aud_c0)
        acoes = [str(x['acao']) for x in a14]
        reg('K14 Audit_Log: CALIBRACAO_REGISTRADA (com o ID) e CALIBRACAO_RECUSADA para as recusas',
            acoes.count('CALIBRACAO_REGISTRADA') >= 2 and acoes.count('CALIBRACAO_RECUSADA') >= 4
            and any(id_c in str(x['parecer']) for x in a14),
            {'acoes': acoes[-10:]})
        rv = q.wb.Names('regView').RefersTo
        reg('K15 regView = B4:L203 (as 11 colunas vão para a RegistrosStore na troca de lote) e cabeçalhos K/L',
            rv.replace('$', '') == '=Registros!B4:L203' and reg_ws.Range('K3').Value == 'ID do resultado'
            and reg_ws.Range('L3').Value == 'Equipamento', {'regView': rv})
    except Exception as ex_geral:                    # noqa: BLE001
        reg('K99 Execução interrompida: os testes seguintes a este ponto não rodaram', False,
            {'erro': f'{type(ex_geral).__name__}: {ex_geral}', 'traceback': traceback.format_exc()[-1500:]})
    finally:
        q.fechar(salvar=False)
    hash1 = sha256_arquivo(caminho)
    reg('K00 Segurança do teste: arquivo de entrada intacto (SHA-256 igual antes e depois; nada salvo nele)',
        hash0 == hash1, {'arquivo': caminho, 'sha256_antes': hash0, 'sha256_depois': hash1})
    if saida:
        with open(saida, 'w', encoding='utf-8') as f:
            json.dump(RES, f, ensure_ascii=False, indent=1, default=str)
    falhas = [r for r in RES if r['resultado'] == 'FAIL']
    print(f'\n=== {produto} ADR-084 (status, motivo livre, indicador): {len(RES) - len(falhas)} PASS / {len(falhas)} FAIL ===',
          flush=True)
    return falhas


if __name__ == '__main__':
    if len(sys.argv) < 4:
        raise SystemExit(__doc__)
    falhas = executar(sys.argv[1], os.path.abspath(sys.argv[2]), sys.argv[3])
    sys.exit(1 if falhas else 0)
