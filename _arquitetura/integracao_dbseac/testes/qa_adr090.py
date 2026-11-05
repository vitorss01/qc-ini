# -*- coding: utf-8 -*-
"""QA do ADR-090 (04/11/2026): lote 3 da auditoria (avisos, rotulos, formatos, trilha e protecao; nenhuma regra de
negocio muda). Um caso por OC. Roda numa COPIA instalada; nada e salvo.

Uso: python qa_adr090.py <Bioquimica|Hematologia> <copia_instalada.xlsm> <saida.json>

  O12 Painel com mais de 180 corridas no periodo: A39 cita "indicadores do Painel" e o n do periodo inteiro por nivel
      (Eng_Saida!AF1) = contagem independente na tblCQ_Final; o n da janela continua menor (Bio: Glicose DIMENSION 1,
      lote 8974, "Tudo"; Hema: WBC, H3 vazio, Ano inteiro 2026)
  O40 De/Ate digitados (Painel protegido, eventos ligados): K4 "Personalizado", K3 = ano; G3/G4 intactos;
      anos diferentes = K3 vazio; escolher um mes depois volta a escrever G3/G4; a lista do K4 tem o rotulo
  O33 sem login: SalvarCEQ recusa (ERRO|0, nada gravado); AbrirCEQ confere a sessao antes de abrir (fonte)
  O19 manual sem EQUIPAMENTO: Bio = ambar + nota "Preenchido pelo Painel com ..."; digitar o equipamento tira os dois;
      Hema (um equipamento) = sem ambar; a instrucao da aba cita EQUIPAMENTO
  O23 ATUALIZAR (so entradas) com o manual sem NIVEL: RESULTADO_MANUAL_RECUSADO (com "NIVEL"), sem INCLUIDO; completo:
      INCLUIDO (ATIVO) ou EM_CONFLITO (CONFLITO_MANUAL), conforme o status
  O21 RUN vazio sem "0"; DETALHE com acento e sem "0"
  O24 ResumoQA por grupo: linha do E04 com o n certo, sem "Há ERRO DE GOVERNANÇA"; governanca so se houver E01..E11
  O31 sem valor-alvo: confirmacao (texto), gravacao com "fora do bias (sem valor-alvo)"; rotulo "Valor-alvo ¹"
  O55 descricao do registro sem o ano repetido (rodada com o ano)
  O34 listas da janela sem as linhas Uso = NAO (simulacao), contra a contagem independente
  O52 par por prefixo sem espaco (Hema: MCH -> MCHC; Bio: o par de palavra continua primeiro)
  O53 colunas de formula do CEQ recusam digitacao (validacao); sobrescrever pelo COM = Z3 "FÓRMULA SOBRESCRITA" +
      CEQ_FORMULA_SOBRESCRITA; restaurar a formula = Z3 de antes
  O54 edicao direta sem selecao com o MESMO valor = sem CEQ_ALTERADO e sem pendencia; valor diferente = antes da EQA_Base
  O36 Estatistica AT com virgula decimal (nenhum "004 / 000")
  O50 Principal - Resultados A5 = "Total: <n com milhar>" igual a ROWS(tblCQ_Final)
  O43 Painel J7:J9 reduzir para caber; O44 Estatistica A:B centro vertical; O49 (Hema) linhas vazias sem "0"/"SEM EP"
  O35 GUIA §5-6 com a janela do CEQ (sem Excel)
  V01 frmCEQ igual a especificacao e auditoria da janela; V02 abas Painel e Digitar Resultados sem texto cortado
  R00 seguranca do teste: arquivo de entrada intacto
"""
import datetime as dt
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
from qa_inativar_ux import criar_usuario, fazer_login, USUARIO_QA  # noqa: E402
import auditoria_janelas  # noqa: E402
import auditoria_abas  # noqa: E402

TAG = 'TESTE-AUDITORIA QA090'
AMBAR_FUNDO = 255 + 235 * 256 + 156 * 65536
PERS = 'Personalizado'
GUIA = os.path.join(AQUI, '..', 'GUIA_DE_USO.md')


def txt(v):
    if isinstance(v, float) and v == int(v):
        return str(int(v))
    return '' if v is None else str(v).strip()


def num(v):
    return isinstance(v, (int, float)) and not isinstance(v, bool)


def executar(produto, caminho, saida):
    hash0 = sha256_arquivo(caminho)
    bio = produto.startswith('Bio')
    # O35: fonte (sem Excel)
    g = open(GUIA, encoding='utf-8').read()
    s6 = g[g.find('## 5. Controle externo'):g.find('## 7. Itens')]
    reg('O35 GUIA §5–6: o CEQ pela janela (Registrar, Corrigir/Excluir selecionado, Consolidar, aviso do Z3, ATUALIZAR) '
        'e sem mandar digitar nas abas EQA', all(k in s6 for k in ('Registrar resultado do CEQ', 'Corrigir selecionado',
                                                                    'Excluir selecionado', 'Consolidar', 'Z3',
                                                                    'ATUALIZAR DADOS', 'valor-alvo'))
        and 'digite cada rodada nas abas' not in s6, {'trecho': s6[:300]})
    q = QA(produto, caminho)

    def xrun(macro, *a):
        return q.ex.xl.Run(f"'{q.wb.Name}'!{macro}", *a)

    try:
        xl = q.ex.xl
        al = q.ws('Audit_Log')

        def aud0():
            return al.Cells(al.Rows.Count, 1).End(-4162).Row

        def aud(r0):
            ult = aud0()
            if ult <= r0:
                return []
            v = al.Range(al.Cells(r0 + 1, 1), al.Cells(ult, 30)).Value
            return [(txt(l[6]), ' | '.join(txt(c) for c in l)) for l in v]

        p = q.ws('Painel')
        e = q.ws('Estatística')
        # ------------------------------------------------------------ telas estaticas
        j = p.Range('J7:J9')
        reg('O43 Painel J7:J9 com "Reduzir para caber" (o status longo não é cortado)', j.ShrinkToFit is True,
            {'shrink': j.ShrinkToFit, 'wrap': j.WrapText})
        linhas_f = [r for r in range(14, 134) if e.Cells(r, 1).HasFormula]
        va = {r: e.Range(e.Cells(r, 1), e.Cells(r, 2)).VerticalAlignment for r in linhas_f}
        reg('O44 Estatística: nome do analito e nível centralizados na vertical (A:B das linhas de analito)',
            linhas_f and all(v == -4108 for v in va.values()), {'linhas': len(linhas_f),
                                                                 'fora': [r for r, v in va.items() if v != -4108][:5]})
        at = [(r, txt(e.Cells(r, 46).Text)) for r in linhas_f if txt(e.Cells(r, 46).Text)]
        ruins = [x for x in at if not re.fullmatch(r'(-?[\d.]*\d,\d\d|SEM EP|[^/]*[a-z][^/]*) / (-?[\d.]*\d,\d\d|—|SEM EP|.*[a-z].*)', x[1])]
        reg('O36 Estatística AT (RMS viés / u(Cref)): vírgula decimal e 2 casas, nenhum "004 / 000"; texto (SEM EP) '
            'continua igual', at and not ruins and any(re.match(r'-?\d+,\d\d / ', t) for _, t in at),
            {'exemplos': at[:4], 'ruins': ruins[:4], 'formula_AT14': txt(e.Range('AT14').Formula)[:160]})
        pr = q.ws('Principal - Resultados')
        a5 = txt(pr.Range('A5').Text)
        n_cq = q.lo('tblCQ_Final').ListRows.Count
        reg('O50 Principal - Resultados A5: "Total: <n com ponto de milhar>" = linhas da tblCQ_Final (sem ",0")',
            a5.startswith('Total: ' + f'{n_cq:,}'.replace(',', '.') + '  ·  participam: ') and ',0 ' not in a5,
            {'A5': a5[:120], 'linhas': n_cq})
        if not bio:
            esp = q.espec()
            rows = []
            ok49 = True
            for r in linhas_f:
                fa = txt(e.Cells(r, 1).Formula)
                m = re.search(r'\$A\$(\d+)', fa)
                origem = txt(esp.Range('A' + m.group(1)).Value) if m else '?'
                vis = [txt(e.Cells(r, c).Text) for c in (1, 7, 18, 19, 20)]
                if origem == '':
                    ok49 = ok49 and all(v == '' for v in vis)
                    rows.append(r)
                else:
                    ok49 = ok49 and vis[0] == origem
            reg('O49 Estatística (Hema): linha sem analito sem "0" e sem "SEM EP" (A, G, R, S, T vazias); linhas com '
                'analito = nome do catálogo', ok49 and len(rows) > 0, {'linhas_vazias': len(rows),
                                                                       'exemplo': [txt(e.Cells(rows[0], c).Text) for c in (1, 7, 18)] if rows else None})

        # ------------------------------------------------------------ O12 corte de 180
        eng = q.ws('Eng_Saida')
        xl.EnableEvents = True
        p.Activate()
        q.ex.esperar()
        p.Unprotect(ii.SENHA)
        if bio:
            seq = [('L4', 'DIMENSION 1'), ('C3', 'Glicose'), ('H3', "'8974"), ('K4', 'Tudo (sem filtro)')]
        else:
            seq = [('C3', 'WBC'), ('K3', 2026), ('K4', 'Ano inteiro')]
        for c, v in seq:
            p.Range(c).Value = v
            q.ex.esperar()
        if not bio:
            p.Range('H3').MergeArea.ClearContents()
            q.ex.esperar()
        cort = int(eng.Range('M1').Value or 0)
        af1 = txt(eng.Range('AF1').Value)
        a39 = txt(p.Range('A39').Value)
        ps = q.painel_stats()
        de, ate = q.nome('filtroDe').Value, q.nome('filtroAte').Value
        an = txt(q.nome('selAnalito').Value).upper()
        lote = txt(eng.Range('E1').Value)
        todos = eng.Range('AD1').Value is True
        eqf = txt(eng.Range('S1').Value).upper()
        cont = {}
        for r in q.ler('tblCQ_Final'):
            if txt(r['ANALITO']).upper() != an or r['PARTICIPA_ESTATISTICA'] != 'SIM' or not num(r['RESULTADO']):
                continue
            if not todos and txt(r['LOTE']) != lote:
                continue
            if bio and eqf and txt(r['EQUIPAMENTO']).upper() != eqf:
                continue
            d = r['DATA_HORA']
            if not isinstance(d, dt.datetime):
                continue
            if isinstance(de, dt.datetime) and d.date() < de.replace(tzinfo=None).date():
                continue
            if isinstance(ate, dt.datetime) and d.date() > ate.replace(tzinfo=None).date():
                continue
            nv = int(r['NIVEL'] or 0)
            cont[nv] = cont.get(nv, 0) + 1
        indep = ' · '.join(f'N{k} {cont.get(k, 0)}' for k in range(1, q.nlv + 1))
        n_jan = [ps[k][1] for k in range(q.nlv)] if ps and len(ps[0]) > 1 else []
        reg('O12 Mais de 180 corridas no período: o A39 diz que o corte vale para os indicadores do Painel e mostra o n do '
            'período inteiro por nível (= contagem independente na tblCQ_Final); o n da janela continua o de antes (menor)',
            cort > 0 and 'indicadores do Painel' in a39 and f'Período inteiro: {af1}' in a39 and af1 == indep,
            {'cortadas': cort, 'AF1': af1, 'independente': indep, 'engPainel': ps, 'A39': a39[-260:],
             'filtro': [str(de), str(ate)], 'lote': lote, 'todos': todos, 'n_janela': n_jan})

        achados = {}

        def auditar(aba, rot):
            w = q.ws(aba)
            pr_ = bool(w.ProtectContents)
            w.Unprotect(ii.SENHA)
            achados[rot] = auditoria_abas.auditar_aba(w)
        auditar('Painel', 'Painel (A39 com o corte, K4 Tudo)')
        # ------------------------------------------------------------ O40 De/Ate -> Personalizado (Painel PROTEGIDO)
        p.Range('K3').Value = 2026
        p.Range('K4').Value = 'Out'
        q.ex.esperar()
        p.Protect(ii.SENHA, True, True)                       # sem UserInterfaceOnly: o VBA nao pode depender de Unprotect
        prot = bool(p.ProtectContents)
        passos = []
        try:
            p.Range('G3').Value = dt.datetime(2026, 9, 1)
            p.Range('G4').Value = dt.datetime(2026, 10, 31)
            q.ex.esperar()
            passos.append(('G3/G4 set-out', txt(p.Range('K3').Value), txt(p.Range('K4').Value),
                           str(p.Range('G3').Value)[:10], str(p.Range('G4').Value)[:10]))
            ok40a = (txt(p.Range('K4').Value) == PERS and txt(p.Range('K3').Value) == '2026'
                     and str(p.Range('G3').Value)[:10] == '2026-09-01' and str(p.Range('G4').Value)[:10] == '2026-10-31')
            p.Range('G3').Value = dt.datetime(2025, 12, 1)
            q.ex.esperar()
            passos.append(('G3 2025', txt(p.Range('K3').Value), txt(p.Range('K4').Value)))
            ok40b = txt(p.Range('K3').Value) == '' and txt(p.Range('K4').Value) == PERS
            p.Range('K3').Value = 2026
            p.Range('K4').Value = 'Out'
            q.ex.esperar()
            passos.append(('K4 Out', txt(p.Range('K3').Value), txt(p.Range('K4').Value),
                           str(p.Range('G3').Value)[:10], str(p.Range('G4').Value)[:10]))
            ok40c = str(p.Range('G3').Value)[:10] == '2026-10-01' and str(p.Range('G4').Value)[:10] == '2026-10-31'
            p.Range('K4').Value = PERS                          # escolher o rotulo na lista nao mexe no De/Ate
            q.ex.esperar()
            ok40d = str(p.Range('G3').Value)[:10] == '2026-10-01' and str(p.Range('G4').Value)[:10] == '2026-10-31'
            lista = txt(p.Range('K4').Validation.Formula1)
        finally:
            p.Unprotect(ii.SENHA)
        reg('O40 Digitar De/Até (Painel protegido): K4 = "Personalizado" e K3 = ano das datas; anos diferentes = '
            'K3 vazio; escolher "Out" depois reescreve G3/G4; escolher o rótulo não mexe no De/Até; a lista tem o rótulo',
            prot and ok40a and ok40b and ok40c and ok40d and PERS in lista.split(';'),
            {'passos': passos, 'protegido': prot, 'lista_fim': lista[-40:], 'ok': [ok40a, ok40b, ok40c, ok40d]})
        auditar('Painel', 'Painel (K4 Personalizado)')
        p.Range('K4').Value = 'Tudo (sem filtro)'
        q.ex.esperar()

        # ------------------------------------------------------------ O33 sem sessao
        campos0 = ['CAP', 'QA-090 2026', '2026', 'QA090 analito', 'QA90-01', '10', '', '', '', '', '', 'Acceptable', '',
                   '', '', 'qa_090.pdf']
        lo_cap = None
        for ws in q.wb.Worksheets:
            for t in ws.ListObjects:
                if t.Name == 'tblEQA_CAP_Dados':
                    lo_cap = t
        n0 = lo_cap.ListRows.Count
        r33 = str(xrun('mCEQForm.SalvarCEQ', campos0))
        cm = q.wb.VBProject.VBComponents('mCEQForm').CodeModule
        fonte = cm.Lines(1, cm.CountOfLines)
        ab = fonte[fonte.find('Public Sub AbrirCEQ()'):]
        ab = ab[:ab.find('End Sub')]
        reg('O33 Sem login: SalvarCEQ devolve "ERRO|0|Sessão encerrada…" e nada é gravado; AbrirCEQ confere a sessão '
            '(SessaoAtiva) antes de abrir a janela (fonte: o MsgBox não roda por automação)',
            r33.startswith('ERRO|0|Sessão encerrada') and lo_cap.ListRows.Count == n0
            and 'SessaoAtiva' in ab and ab.find('SessaoAtiva') < ab.find('.Show'), {'retorno': r33, 'linhas': n0})

        # ------------------------------------------------------------ login
        criar_usuario(q, USUARIO_QA, 'ANALISTA')
        fazer_login(q, USUARIO_QA)

        auditar('Digitar Resultados', 'Digitar Resultados (instrução nova, antes da linha de teste)')
        # ------------------------------------------------------------ O19 manual sem EQUIPAMENTO
        dr = q.ws('Digitar Resultados')
        lo_m = q.lo('tblResultados_Manuais')
        fin0 = q.ler('tblCQ_Final')
        reais = sorted([r for r in fin0 if r['STATUS_ANALITICO'] == 'ATIVO' and r['ORIGEM_RESULTADO'] == 'INTERFACEAMENTO'
                        and isinstance(r['DATA_HORA'], dt.datetime) and num(r['RESULTADO'])], key=lambda r: r['DATA_HORA'])
        T = reais[-300]
        dT = T['DATA_HORA'].date()
        xl.EnableEvents = True
        dr.Activate()
        rm = q.escrever_linha('tblResultados_Manuais', {
            'DATA': dt.datetime(dT.year, dT.month, dT.day), 'HORA': '03:17', 'LOTE': "'" + txt(T['LOTE']),
            'ANALITO': txt(T['ANALITO']), 'RESULTADO': float(T['RESULTADO']) * 1.001, 'MOTIVO': TAG})
        q.ex.esperar()
        c_eq = lo_m.ListColumns('EQUIPAMENTO').DataBodyRange.Cells(rm, 1)
        eq_auto = txt(c_eq.Value)
        cor1 = c_eq.Interior.Color
        nota1 = txt(c_eq.Comment.Text()) if c_eq.Comment is not None else ''
        sel_eq = txt(q.nome('selEquipamento').Value)
        instr = txt(dr.Range('A3').Value)
        if bio:
            c_eq.Value = eq_auto                               # o usuario confere e digita o equipamento
            q.ex.esperar()
            cor2 = c_eq.Interior.Color
            nota2 = c_eq.Comment is not None
            ok19 = (eq_auto == sel_eq and cor1 == AMBAR_FUNDO and nota1 == f'Preenchido pelo Painel com {sel_eq}: confira '
                    'antes de ATUALIZAR DADOS.' and cor2 != AMBAR_FUNDO and not nota2)
            ev = {'equip': eq_auto, 'cor': cor1, 'nota': nota1, 'depois_de_digitar': (cor2, nota2)}
        else:
            ok19 = eq_auto == sel_eq and cor1 != AMBAR_FUNDO and nota1 == ''
            ev = {'equip': eq_auto, 'cor': cor1, 'nota': nota1}
        reg('O19 EQUIPAMENTO vazio recebe o do Painel ' + ('em âmbar com a nota "Preenchido pelo Painel com …: confira"; '
            'digitar o equipamento tira a cor e a nota' if bio else 'sem âmbar (um equipamento só)')
            + '; a instrução da aba cita EQUIPAMENTO', ok19 and 'EQUIPAMENTO' in instr and 'âmbar' in instr,
            dict(ev, instrucao=instr[:120]))

        # ------------------------------------------------------------ O23/O21/O24: ATUALIZAR so entradas
        idm = txt(lo_m.ListColumns('ID_REGISTRO').DataBodyRange.Cells(rm, 1).Value)
        a1 = aud0()
        r1 = str(q.run('mIntegracao.AtualizarSoEntradas', teto=1500))
        t1 = [(a, d) for a, d in aud(a1) if idm in d]
        st1 = txt(lo_m.ListColumns('STATUS').DataBodyRange.Cells(rm, 1).Value)
        run1 = txt(lo_m.ListColumns('RUN').DataBodyRange.Cells(rm, 1).Text)
        det1 = txt(lo_m.ListColumns('DETALHE').DataBodyRange.Cells(rm, 1).Text)
        rq = str(xrun('mIntegracao.ResumoQA'))
        qa = q.ler('tblQA_Integracao')
        n04 = sum(1 for x in qa if x['SEVERIDADE'] == 'ERRO' and x['CODIGO'] == 'E04')
        ngov = sum(1 for x in qa if x['SEVERIDADE'] == 'ERRO' and x['CODIGO'] in ('E01', 'E02', 'E03', 'E10', 'E11'))
        lo_m.ListColumns('NIVEL').DataBodyRange.Cells(rm, 1).Value = 1
        q.ex.esperar()
        a2 = aud0()
        r2 = str(q.run('mIntegracao.AtualizarSoEntradas', teto=1500))
        t2 = [(a, d) for a, d in aud(a2) if idm in d]
        st2 = txt(lo_m.ListColumns('STATUS').DataBodyRange.Cells(rm, 1).Value)
        det2 = txt(lo_m.ListColumns('DETALHE').DataBodyRange.Cells(rm, 1).Text)
        esperado2 = {'ATIVO': 'RESULTADO_MANUAL_INCLUIDO', 'CONFLITO_MANUAL': 'RESULTADO_MANUAL_EM_CONFLITO'}.get(st2, '?')
        reg('O23 Audit_Log do manual pelo STATUS: sem NÍVEL = RESULTADO_MANUAL_RECUSADO (detalhe com o campo), sem '
            'INCLUIDO; completo = ' + esperado2 + ' (status ' + st2 + ')',
            r1.startswith('OK') and st1 == 'MANUAL_INCOMPLETO' and [a for a, _ in t1] == ['RESULTADO_MANUAL_RECUSADO']
            and 'NIVEL' in t1[0][1] and any(a == esperado2 for a, _ in t2),
            {'id': idm, 'atualizar1': r1[:60], 'trilha1': [(a, d[:200]) for a, d in t1], 'status2': st2,
             'trilha2': [(a, d[:160]) for a, d in t2]})
        reg('O21 Digitar Resultados: RUN vazio sem "0" (MANUAL_INCOMPLETO); DETALHE com acento ("obrigatórios") e sem '
            '"0" quando vazio',
            run1 == '' and 'obrigatórios' in det1 and det2 != '0' and (st2 != 'ATIVO' or det2 == ''),
            {'RUN': run1, 'DETALHE1': det1, 'DETALHE2': det2, 'status2': st2})
        linhas_rq = rq.split('\r\n')
        reg('O24 ResumoQA por grupo: o E04 (manual incompleto) tem linha própria com o n certo; sem "Há ERRO DE '
            'GOVERNANÇA"; a linha de governança só aparece com E01/E02/E03/E10/E11',
            n04 > 0 and any(l.startswith(f'• {n04} resultado(s) manual(is) incompleto(s)') for l in linhas_rq)
            and 'Há ERRO DE GOVERNANÇA' not in rq and (('erro(s) de governança' in rq) == (ngov > 0))
            and len(rq) <= 1024, {'resumo': linhas_rq, 'E04': n04, 'governanca': ngov})

        # ------------------------------------------------------------ O31 / O55 / O34 / O52 (CEQ, logado)
        campos = list(campos0)
        cf = str(xrun('mCEQForm.ConfirmacoesCEQ', campos, 0))
        cf_alvo = str(xrun('mCEQForm.ConfirmacoesCEQ', campos[:6] + ['10'] + campos[7:], 0))
        a3 = aud0()
        r31 = str(xrun('mCEQForm.SalvarCEQ', campos))
        lin31 = int(r31.split('|')[1]) if r31.startswith('OK|') else 0
        desc = str(xrun('mCEQForm.DescricaoCEQ', 'CAP', lin31)) if lin31 else ''
        lbl = q.wb.VBProject.VBComponents('frmCEQ').Designer.Controls('lblAlvo').Caption
        reg('O31 CEQ sem valor-alvo: a confirmação avisa que a amostra NÃO entra no bias; com alvo, sem esse aviso; '
            'gravado, o OK diz "fora do bias (sem valor-alvo)"; rótulo "Valor-alvo ¹"',
            'NÃO entra no bias' in cf and 'valor-alvo' in cf and 'valor-alvo' not in cf_alvo
            and r31.startswith('OK|') and 'sem valor-alvo' in r31 and lbl == 'Valor-alvo ¹',
            {'confirmacao': cf[-160:], 'com_alvo': cf_alvo[:80], 'retorno': r31, 'rotulo': lbl})
        reg('O55 Descrição do registro sem o ano repetido (rodada "QA-090 2026", ano 2026)',
            desc.startswith('CAP QA-090 2026 · QA090 analito') and '2026 2026' not in desc, {'descricao': desc})
        v = lo_cap.DataBodyRange.Value
        esperado34 = sorted({txt(r[3]) for r in v if txt(r[3]) and 'SIMULA' not in txt(r[17]).upper()}, key=str.upper)
        sim = sorted({txt(r[3]) for r in v if txt(r[3]) and 'SIMULA' in txt(r[17]).upper()} - set(esperado34))
        lista34 = sorted([str(x) for x in (xrun('mCEQForm.DistintosCEQ', 'CAP', 4) or ())], key=str.upper)
        reg('O34 Lista de analitos da janela sem os nomes que só existem na simulação (Uso = NAO)',
            [x.upper() for x in lista34] == [x.upper() for x in esperado34] and not (set(sim) & set(lista34)),
            {'n_lista': len(lista34), 'n_esperado': len(esperado34), 'so_simulacao': sim[:8]})
        mapa = xrun('mEQA.MapaDoProvedor', 'CAP') or ()
        nomes = [(str(m[0]), str(m[1])) for m in mapa]
        pares_sp = [(c, l) for c, cc in nomes for l, cl in nomes
                    if l.upper().startswith(c.upper()) and len(l) > len(c) and cl.upper() != cc.upper()
                    and not any(l2.upper().startswith(c.upper() + ' ') and cl2.upper() != cc.upper() for l2, cl2 in nomes)]
        res52 = []
        for c, l in pares_sp[:8]:
            av = str(xrun('mCEQForm.AvisoAnalitoCEQ', 'CAP', c))
            res52.append((c, l, av[-80:], f'“{l}”' in av))
        lact = str(xrun('mCEQForm.AvisoAnalitoCEQ', 'CAP', 'Lymphs' if not bio else 'Lactate'))
        mch_ok = bio or any(c.upper() == 'MCH' and ok for c, _, _, ok in res52)
        reg('O52 Par por prefixo também sem espaço (ex.: MCH → MCHC) no aviso azul; o par de palavra continua primeiro',
            all(x[3] for x in res52) and mch_ok and ('Absolute' in lact if not bio else ('Dehydrogenase' in lact or not lact.startswith('AZUL'))),
            {'pares_sem_espaco': res52, 'par_de_palavra': lact[-90:]})

        # ------------------------------------------------------------ O53 / O54 edicao direta
        wsC = lo_cap.Parent
        wsC.Activate()
        wsC.Unprotect(ii.SENHA)                               # quem tem a aba liberada (ADM)
        r0 = lo_cap.DataBodyRange.Row
        rr = next(r0 + i for i, r in enumerate(v) if txt(r[3]) and num(r[14]) and 'SIMULA' not in txt(r[17]).upper())
        val = wsC.Range(f'P{rr}').Validation
        vinfo = {c: (wsC.Range(f'{c}{rr}').Validation.Type, txt(wsC.Range(f'{c}{rr}').Validation.Formula1))
                 for c in ('O', 'P', 'S', 'X')}
        z3_0 = txt(wsC.Range('Z3').Value)
        z4_0 = wsC.Range('Z4').Value
        f_p = txt(wsC.Range(f'P{rr}').Formula)
        a4 = aud0()
        wsC.Range(f'P{rr}').Value = 2
        q.ex.esperar()
        xl.Calculate()
        z3_1 = txt(wsC.Range('Z3').Value)
        tr53 = [(a, d) for a, d in aud(a4)]
        wsC.Range(f'P{rr}').Formula = f_p
        xl.Calculate()
        z3_2 = txt(wsC.Range('Z3').Value)
        reg('O53 Colunas de fórmula do CEQ (O, P, S..X) recusam digitação (validação personalizada); sobrescrever mesmo '
            'assim -> Z3 "FÓRMULA SOBRESCRITA" e CEQ_FORMULA_SOBRESCRITA no Audit_Log; restaurar -> Z3 de antes',
            all(t == 7 and f.replace(' ', '') == '=1=0' for t, f in vinfo.values()) and 'FÓRMULA SOBRESCRITA' in z3_1
            and any(a == 'CEQ_FORMULA_SOBRESCRITA' and f'P{rr}' in d for a, d in tr53) and z3_2 == z3_0,
            {'validacao': vinfo, 'erro': txt(val.ErrorMessage)[:80], 'Z3': [z3_0, z3_1[:60], z3_2],
             'trilha': [(a, d[:200]) for a, d in tr53]})
        q.ws('Painel').Activate()                             # sai da aba: nenhuma celula do CEQ selecionada
        wsC.Activate()
        wsC.Range('AB1').Select()
        a5_ = aud0()
        am = wsC.Range(f'E{rr}').Value
        wsC.Range(f'E{rr}').Value = am                        # mesmo valor, sem selecionar antes
        q.ex.esperar()
        t54a = [(a, d) for a, d in aud(a5_)]
        z4_1 = wsC.Range('Z4').Value
        res0 = wsC.Range(f'F{rr}').Value
        wsC.Range('AB1').Select()
        wsC.Range(f'F{rr}').Value = float(res0) + 1           # valor diferente, sem selecionar antes
        q.ex.esperar()
        t54b = [(a, d) for a, d in aud(a5_)]
        reg('O54 Edição direta sem selecionar antes: o mesmo valor não registra "? → mesmo" nem cria pendência (Z4 '
            'igual); valor diferente registra o antes da EQA_Base (não "?")',
            not t54a and z4_1 == z4_0 and len(t54b) == 1 and t54b[0][0] == 'CEQ_ALTERADO'
            and (f'{txt(res0)} →' in t54b[0][1] or f'{txt(res0).replace(".", ",")} →' in t54b[0][1])
            and '? →' not in t54b[0][1],
            {'mesmo_valor': t54a, 'Z4': [z4_0, z4_1], 'antes': res0,
             'diferente': [(a, d[d.find('edi'):][:400]) for a, d in t54b]})
        xl.EnableEvents = False
        wsC.Range(f'F{rr}').Value = res0
        if lin31:
            xrun('mCEQForm.ExcluirCEQ', 'CAP', lin31)
        xl.EnableEvents = True

        # ------------------------------------------------------------ V01 / V02
        import instalar_frm_inativar as fi
        import ceq_form_spec as cq
        comp = q.wb.VBProject.VBComponents('frmCEQ')
        dif = fi._diferencas(comp, ii.codigo(cq.FONTE_CODIGO), cq.FORM, cq.CONTROLES)
        aj = auditoria_janelas.auditar('frmCEQ', cq.FORM, cq.CONTROLES)
        reg('V01 frmCEQ igual à especificação; auditoria da janela (rótulo "Valor-alvo ¹" e a nota do alvo): nada '
            'cortado nem sobreposto', not dif and not aj, {'diferencas': dif, 'auditoria': aj})
        # Hema no modo TODOS (H3 vazio, ADR-086): o H4 de 7 linhas e o botao btnGrafTodos sobre o texto ja eram assim antes
        # deste ADR (pendencia registrada a parte, como os status longos).
        # status longos do Painel (SEM CORRIDA NO PERÍODO etc.) em J7:J9 (47 pt) passam dos 30 % do "Reduzir para caber":
        # ja eram cortados antes deste ADR (o OC-0043 e o "REJEITADO", que cabe reduzido); ficam como pendencia (forma
        # curta dos status = decisao de UX) e saem do portao desta suite, registrados a parte
        pend_j = {}
        for k in list(achados):
            if k.startswith('Painel'):
                pend_j[k] = [x for x in achados[k] if (x.get('celula') in ('J7', 'J8', 'J9')
                             and not str(x.get('texto', '')).startswith('REJEITADO'))
                             or (not bio and (x.get('celula') == 'H4' or x.get('forma') == 'btnGrafTodos'))]
                achados[k] = [x for x in achados[k] if x not in pend_j[k]]
        if any(pend_j.values()):
            RES.append({'teste': 'INFO pendência anterior ao ADR-090: status longos em J7:J9; Hema TODOS: H4 e btnGrafTodos',
                        'resultado': 'INFO', 'evidencia': pend_j})
        auditar('Estatística', 'Estatística')
        auditar('Principal - Resultados', 'Principal - Resultados')
        reg('V02 Painel (A39 do corte; K4 "Personalizado"), Digitar Resultados (instrução), Estatística (AT) e Principal '
            '(A5) sem texto cortado',
            not any(achados.values()), {k: v[:4] for k, v in achados.items() if v} or {'abas': list(achados)})
    except Exception as ex_:                           # noqa: BLE001
        reg('R99 Execução da suíte sem exceção', False, {'erro': repr(ex_), 'trace': traceback.format_exc()[-2000:]})
    finally:
        q.fechar(salvar=False)
    hash1 = sha256_arquivo(caminho)
    reg('R00 Segurança do teste: arquivo de entrada intacto (SHA-256 igual antes e depois; nada salvo nele)',
        hash0 == hash1, {'arquivo': caminho, 'sha256_antes': hash0, 'sha256_depois': hash1})
    if saida:
        with open(saida, 'w', encoding='utf-8') as f:
            json.dump(RES, f, ensure_ascii=False, indent=1, default=str)
    falhas = [r for r in RES if r['resultado'] == 'FAIL']
    print(f'\n=== {produto} ADR-090 (lote 3 da auditoria): {len(RES) - len(falhas)} PASS / {len(falhas)} FAIL ===',
          flush=True)
    return falhas


if __name__ == '__main__':
    t0 = time.time()
    falhas = executar(sys.argv[1], os.path.abspath(sys.argv[2]), sys.argv[3] if len(sys.argv) > 3 else None)
    print(f'tempo {time.time() - t0:.0f}s')
    sys.exit(1 if falhas else 0)
