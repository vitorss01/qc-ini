# -*- coding: utf-8 -*-
"""QA do ADR-089 (04/11/2026): lote 2 da auditoria (CEQ: corrigir/excluir pela janela, consolidacao e Z2/Z3, analito
fora do mapa, par %/#, trilha) e OC-0051 (H3 apagado depois de lote + periodo). Roda numa COPIA instalada; nada e salvo.

Uso: python qa_adr089.py <Bioquimica|Hematologia> <copia_instalada.xlsm> <saida.json>

  L51 OC-0051 (sessao SEM o UserInterfaceOnly, como a automacao e o arquivo recem-aberto): lote em H3 -> K3/K4 ->
      apagar H3: o motor acompanha cada passo (E1 = lote, nunca DESATUALIZADO) e o H3 vazio entra no modo TODOS
      (Eng_Saida!AD1, H4 "TODOS", G6 "do método")
  C05 (G5) ANALISTA registra 110 e corrige para 105 pela janela (CorrigirCEQ): mesma linha, EQA_Base consolidada na
      hora com 105, G da Estatistica muda, CEQ_ALTERADO "110 -> 105"; a duplicidade aponta o botao Corrigir
  C09 Excluir selecionado (ExcluirCEQ): linha limpa, chave fora da EQA_Base, G volta EXATAMENTE ao de antes do registro,
      CEQ_EXCLUIDO com os valores apagados; EQA_BASE_CONSOLIDADA com antes/depois nos tres passos
  C06 (G6) edicao direta na tabela (eventos ligados): Z3 "HÁ EDIÇÃO NÃO CONSOLIDADA", pendencia, CEQ_ALTERADO com o
      antes e o depois; a consolidacao do ATUALIZAR (ConsolidarSePendente) leva o valor a EQA_Base, Z2 = Z1 novo e Z3
      volta a "base em dia"; o ATUALIZAR chama a consolidacao antes do motor (fonte)
  C07 (G7) nomes fora do mapa (CAP e Controllab): aviso ambar, confirmacao antes de gravar e aviso "fora do bias" no OK;
      nome do mapa = aviso azul com o canonico
  C08 (G8) pares do mapa (nome curto = prefixo de outro nome com outro canonico: os 7 %/# da Hema; Lactate / Lactate
      Dehydrogenase na Bio; o nRBC/100 WBC x nRBC Absolute nao e de prefixo e "nRBC" sozinho
      fica ambar): o nome curto mostra o canonico e o par; gravar o nome curto com a unidade do par pede
      confirmacao de unidade (Hema: Lymphs x 10E9/L -> LYMPH%)
  C10 PROVEDOR_ALTERADO com antes e depois (Bio: Painel M4; Hema: Estatistica R4); o provedor volta ao de antes
  V03 janela de verdade (modulo de teste so na copia): Corrigir/Excluir desabilitados sem selecao; analito da lista ->
      aviso azul no lblMsg
  V01 frmCEQ igual a especificacao; auditoria das janelas sem corte nem sobreposicao (textos novos medidos)
  R00 seguranca do teste: arquivo de entrada intacto
"""
import json
import os
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

SETA = '→'
CANON = {'Bioquimica': 'Proteína total', 'Hematologia': 'RBC'}
FORA = {'Bioquimica': [('CAP', 'Protein, total'), ('CAP', 'Hemoglobin A1c'), ('Controllab', 'Proteína total')],
        'Hematologia': [('CAP', 'Red Blood Cell'), ('Controllab', 'RBC')]}
SEQ51 = {'Bioquimica': [('L4', 'DIMENSION 2'), ('C3', 'Hemoglobina glicada'), ('H3', "'8590"), ('K3', 2026), ('K4', 'Set')],
         'Hematologia': [('C3', 'WBC'), ('H3', "'602911"), ('K3', 2026), ('K4', 'Mar')]}


def txt(v):
    if isinstance(v, float) and v == int(v):
        return str(int(v))
    return '' if v is None else str(v).strip()


def executar(produto, caminho, saida):
    hash0 = sha256_arquivo(caminho)
    q = QA(produto, caminho)
    bio = produto.startswith('Bio')

    def run(macro, *a, teto=600):
        return q.run(macro, *a, teto=teto)

    def xrun(macro, *a):
        return q.ex.xl.Run(f"'{q.wb.Name}'!{macro}", *a)

    try:
        xl = q.ex.xl
        al = q.ws('Audit_Log')

        def aud0():
            return al.Cells(al.Rows.Count, 1).End(-4162).Row

        def aud(r0):
            ult = aud0()
            return [(txt(al.Cells(r, 7).Value), txt(al.Cells(r, 24).Value)) for r in range(r0 + 1, ult + 1)]

        # ------------------------------------------------------------ L51 (antes de qualquer login: sem UserInterfaceOnly)
        cfg = q.wb.Names('loteParam').RefersToRange.Worksheet
        uio0 = bool(cfg.ProtectionMode)
        p = q.ws('Painel')
        p.Activate()
        q.ex.esperar()
        eng = q.ws('Eng_Saida')
        xl.EnableEvents = True
        passos = []

        def tela(rot):
            q.ex.esperar()
            d = {'passo': rot, 'H3': txt(p.Range('H3').Value), 'E1': txt(eng.Range('E1').Value),
                 'AD1': eng.Range('AD1').Value, 'J7': txt(p.Range('J7').Value), 'H4': txt(p.Range('H4').Value)[:60],
                 'G6': txt(p.Range('G6').Value)}
            passos.append(d)
            return d
        p.Unprotect(ii.SENHA)
        for c, v in SEQ51[produto]:
            p.Range(c).Value = v
            tela('apos_' + c)
        p.Range('H3').MergeArea.ClearContents()
        fim = tela('h3_vazio')
        lote = SEQ51[produto][[c for c, _ in SEQ51[produto]].index('H3')][1].strip("'")
        com_lote = [d for d in passos if d['passo'] != 'h3_vazio']
        ok51 = (all(d['E1'] == lote for d in com_lote if d['passo'] != 'apos_L4' and d['passo'] != 'apos_C3')
                and not any(str(d['J7']).startswith('DESATUAL') for d in passos)
                and fim['AD1'] is True and 'TODOS' in fim['H4'] and 'método' in fim['G6'])
        reg('L51 OC-0051: sem o UserInterfaceOnly da sessão, lote em H3 -> K3/K4 -> apagar H3: o motor acompanha cada '
            'passo (E1 = lote, nunca DESATUALIZADO) e o H3 vazio entra no modo TODOS (AD1, H4, G6 "do método")',
            ok51 and not uio0, {'configuracao_userinterfaceonly_no_inicio': uio0, 'passos': passos})

        # ------------------------------------------------------------ preparo do CEQ (ANALISTA)
        criar_usuario(q, USUARIO_QA, 'ANALISTA')
        fazer_login(q, USUARIO_QA)
        papel = txt(q.nome('currentUser').Value)
        mapa = xrun('mEQA.MapaDoProvedor', 'CAP') or ()
        canon = CANON[produto]
        nome_cap = next((m[0] for m in mapa if str(m[1]).strip().upper() == canon.upper()), None)
        e = q.ws('Estatística')

        def linha_est():
            for r in range(5, 200):
                if txt(e.Cells(r, 1).Value).upper() == canon.upper() and txt(e.Cells(r, 2).Value) in ('1', 'N1'):
                    return r
            for r in range(5, 200):
                if txt(e.Cells(r, 1).Value).upper() == canon.upper():
                    return r
            return None
        rE = linha_est()

        def g():
            q.ex.esperar()
            xl.Calculate()
            q.ex.esperar()
            return e.Cells(rE, 7).Value if rE else None
        try:
            ano = int(float(q.wb.Names('eqAnoEP').RefersToRange.Value))
        except Exception:                                # noqa: BLE001
            ano = 2026
        rod = 'QA-089 R1'
        base = q.ws('EQA_Base')
        lo_cap = next(t for w in q.wb.Worksheets for t in w.ListObjects if t.Name == 'tblEQA_CAP_Dados')
        wsc = lo_cap.Parent

        def campos(**kw):
            c = {'prov': 'CAP', 'rod': rod, 'ano': str(ano), 'an': nome_cap, 'am': 'QA089-01', 'res': '110',
                 'alvo': '100', 'sd': '5', 'sdi': '', 'linf': '', 'lsup': '', 'aval': 'Acceptable', 'nlabs': '',
                 'unid': '', 'pag': '', 'arq': 'TESTE-AUDITORIA qa_adr089'}
            c.update(kw)
            return [c[k] for k in ('prov', 'rod', 'ano', 'an', 'am', 'res', 'alvo', 'sd', 'sdi', 'linf', 'lsup', 'aval',
                                   'nlabs', 'unid', 'pag', 'arq')]

        def base_por_chave(ch):
            v = base.Range('A2:U5001').Value
            return [r for r in v if txt(r[20]).upper() == ch.upper()]
        chave = f'CAP|{ano}|{rod}|{nome_cap}|QA089-01'
        z1_0 = txt(base.Range('Z1').Value)

        # ------------------------------------------------------------ C05 registrar e corrigir
        g0 = g()
        a0 = aud0()
        r1 = str(xrun('mCEQForm.SalvarCEQ', campos()))
        lin = int(r1.split('|')[1]) if r1.startswith('OK|') else 0
        g1 = g()
        b1 = base_por_chave(chave)
        dup = str(xrun('mCEQForm.SalvarCEQ', campos(res='105')))
        lista = xrun('mCEQForm.RegistrosDaRodada', 'CAP', rod, str(ano)) or ()
        lin_lista = int(float(lista[0][4])) if lista and len(lista[0]) > 4 else 0
        r2 = str(xrun('mCEQForm.CorrigirCEQ', lin_lista, campos(res='105')))
        g2 = g()
        b2 = base_por_chave(chave)
        tr2 = aud(a0)
        alt = [d for a, d in tr2 if a == 'CEQ_ALTERADO']
        prot_aba = bool(wsc.ProtectContents)
        reg('C05 (G5) ANALISTA registra 110 e corrige para 105 pela janela: mesma linha, EQA_Base com 105 sem outro passo, '
            'G muda, CEQ_ALTERADO "110 -> 105"; a duplicidade aponta o botão Corrigir selecionado',
            bool(nome_cap) and r1.startswith('OK|') and lin_lista == lin and r2.startswith(f'OK|{lin}|')
            and txt(wsc.Cells(lin, 6).Value) == '105' and len(b2) == 1 and txt(b2[0][6]) == '105'
            and dup.startswith('ERRO|5|') and 'Corrigir selecionado' in dup and 'edite a linha' not in dup
            and any(f'110 {SETA} 105' in d for d in alt) and rE and g1 != g0 and g2 != g1,
            {'papel': papel, 'analito_cap': nome_cap, 'linha': lin, 'registro': r1, 'duplicidade': dup,
             'correcao': r2, 'base_apos_registro': [txt(x) for x in b1[0][:8]] if b1 else None,
             'base_apos_correcao': [txt(x) for x in b2[0][:8]] if b2 else None, 'G0': g0, 'G_registro': g1,
             'G_correcao': g2, 'trilha': tr2[-4:], 'aba_protegida': prot_aba})

        # ------------------------------------------------------------ C06 edicao direta + consolidacao do ATUALIZAR
        a1 = aud0()
        wsc.Unprotect(ii.SENHA)                          # quem edita direto e quem tem a aba liberada (ADM)
        wsc.Activate()
        xl.EnableEvents = True
        wsc.Cells(lin, 6).Select()                       # SelectionChange guarda o "antes"
        wsc.Cells(lin, 6).Value = 107
        q.ex.esperar()
        xl.Calculate()
        z3_ed = txt(wsc.Range('Z3').Value)
        pend = run('mEQA.ConsolidacaoPendente')
        b_ed = base_por_chave(chave)
        tr_ed = aud(a1)
        consol = run('mEQA.ConsolidarSePendente', teto=900)
        xl.Calculate()
        b_c = base_por_chave(chave)
        z1_c, z2_c, z3_c = txt(base.Range('Z1').Value), txt(wsc.Range('Z2').Value), txt(wsc.Range('Z3').Value)
        pend2 = run('mEQA.ConsolidacaoPendente')
        g3 = g()
        cm = q.wb.VBProject.VBComponents('mIntegracao').CodeModule
        fonte = cm.Lines(1, cm.CountOfLines)
        chama = 'mEQA.ConsolidarSePendente' in fonte and fonte.index('mEQA.ConsolidarSePendente') < fonte.index(
            'mDados.IniciarSnapshot', fonte.index('mEQA.ConsolidarSePendente'))
        reg('C06 (G6) Edição direta na tabela: Z3 "HÁ EDIÇÃO NÃO CONSOLIDADA", pendência, CEQ_ALTERADO com antes e depois; a '
            'consolidação do ATUALIZAR leva o valor à EQA_Base, Z2 = Z1 novo e Z3 volta a "base em dia"',
            z3_ed.startswith('⚠ HÁ EDIÇÃO NÃO CONSOLIDADA') and pend is True and txt(b_ed[0][6]) == '105'
            and any(a == 'CEQ_ALTERADO' and f'105 {SETA} 107' in d for a, d in tr_ed) and consol is True
            and txt(b_c[0][6]) == '107' and z2_c == z1_c and z1_c != z1_0 and z3_c == 'base em dia' and pend2 is False
            and chama and g3 != g2,
            {'Z3_editado': z3_ed, 'pendente': pend, 'trilha_edicao': tr_ed[-2:], 'consolidou': consol,
             'base_depois': txt(b_c[0][6]) if b_c else None, 'Z1': z1_c[:80], 'Z2': z2_c[:80], 'Z3': z3_c,
             'pendente_depois': pend2, 'ATUALIZAR_chama_antes_do_motor': chama, 'G_107': g3})
        xl.EnableEvents = False
        q.run('mSeguranca.ReprotectAll')

        # ------------------------------------------------------------ C09 excluir
        a2 = aud0()
        d9 = str(xrun('mCEQForm.DescricaoCEQ', 'CAP', lin))
        r9 = str(xrun('mCEQForm.ExcluirCEQ', 'CAP', lin))
        g4 = g()
        b9 = base_por_chave(chave)
        tr9 = aud(a2)
        vazia = not any(txt(x) for x in wsc.Range(wsc.Cells(lin, 1), wsc.Cells(lin, 14)).Value[0])
        tr_tudo = aud(a0)
        cons = [d for a, d in tr_tudo if a == 'EQA_BASE_CONSOLIDADA']
        reg('C09 Excluir selecionado: linha limpa, chave fora da EQA_Base, G volta exatamente ao de antes do registro, '
            'CEQ_EXCLUIDO com os valores apagados; EQA_BASE_CONSOLIDADA com origem, antes e depois em cada passo',
            r9.startswith(f'OK|{lin}|') and vazia and not b9 and g4 == g0 and nome_cap in d9
            and any(a == 'CEQ_EXCLUIDO' and 'valores apagados' in d and '107' in d for a, d in tr9)
            and len(cons) >= 4 and all('antes:' in d and 'depois:' in d and 'origem:' in d for d in cons),
            {'descricao': d9, 'exclusao': r9, 'G0': g0, 'G_depois': g4, 'trilha': tr9[-2:],
             'consolidacoes': len(cons), 'exemplo': cons[-1][:300] if cons else None})

        # ------------------------------------------------------------ C07 fora do mapa
        res7 = []
        ok7 = True
        for i, (prov, nm) in enumerate(FORA[produto]):
            av = str(xrun('mCEQForm.AvisoAnalitoCEQ', prov, nm))
            cf = str(xrun('mCEQForm.ConfirmacoesCEQ', campos(prov=prov, an=nm, am=f'QA089-F{i}'), 0))
            rs = str(xrun('mCEQForm.SalvarCEQ', campos(prov=prov, an=nm, am=f'QA089-F{i}')))
            ln = int(rs.split('|')[1]) if rs.startswith('OK|') else 0
            ex_ = str(xrun('mCEQForm.ExcluirCEQ', prov, ln)) if ln else ''
            okk = (av.startswith('AMBAR|⚠') and 'NÃO entra no bias' in av and 'não está no mapa' in cf
                   and rs.startswith('OK|') and 'fora do bias' in rs and ex_.startswith('OK|'))
            ok7 = ok7 and okk
            res7.append({'provedor': prov, 'nome': nm, 'aviso': av[:90], 'confirmacao': cf[:80], 'gravacao': rs,
                         'limpeza': ex_, 'ok': okk})
        azul = str(xrun('mCEQForm.AvisoAnalitoCEQ', 'CAP', nome_cap))
        cf_ok = str(xrun('mCEQForm.ConfirmacoesCEQ', campos(am='QA089-X'), 0))
        reg('C07 (G7) Nome fora do mapa (CAP e Controllab): aviso âmbar, confirmação antes de gravar e "fora do bias" no '
            'OK; nome do mapa = aviso azul com o canônico e sem confirmação',
            ok7 and azul.startswith(f'AZUL|No QC_INI: {canon}.') and cf_ok == '',
            {'fora_do_mapa': res7, 'mapeado': azul, 'confirmacao_mapeado': cf_ok})

        # ------------------------------------------------------------ C08 pares %/#
        pares = []
        nomes = [(str(m[0]), str(m[1])) for m in mapa]
        for curto, ccurto in nomes:
            for longo, clongo in nomes:
                if longo.upper().startswith(curto.upper() + ' ') and clongo.upper() != ccurto.upper():
                    pares.append((curto, ccurto, longo, clongo))
        lo_rows = lo_cap.DataBodyRange.Value

        def unidade_de(cn):
            us = {}
            for r in lo_rows:
                if any(txt(r[3]).upper() == n.upper() and c.upper() == cn.upper() for n, c in nomes) and txt(r[13]):
                    us[txt(r[13])] = us.get(txt(r[13]), 0) + 1
            return max(us, key=us.get) if us else ''
        res8, ok8 = [], True
        for curto, ccurto, longo, clongo in pares:
            av = str(xrun('mCEQForm.AvisoAnalitoCEQ', 'CAP', curto))
            u_curto, u_longo = unidade_de(ccurto), unidade_de(clongo)
            cf = str(xrun('mCEQForm.ConfirmacoesCEQ', campos(an=curto, unid=u_longo, am='QA089-P'), 0))
            precisa_conf = bool(u_curto and u_longo and u_curto.replace(' ', '').upper() != u_longo.replace(' ', '').upper())
            okk = (av.startswith(f'AZUL|No QC_INI: {ccurto}.') and f'“{longo}” ({clongo})' in av
                   and ((not precisa_conf) or ('unidade' in cf and f'“{longo}”' in cf)))
            ok8 = ok8 and okk
            res8.append({'curto': curto, 'canonico': ccurto, 'par': longo, 'canonico_par': clongo, 'unid_curto': u_curto,
                         'unid_par': u_longo, 'aviso': av[:140], 'confirmacao': cf[:160], 'ok': okk})
        lymphs = next((x for x in res8 if x['curto'].upper() == 'LYMPHS'), None)
        cf_ly = str(xrun('mCEQForm.ConfirmacoesCEQ', campos(an='Lymphs', unid='x 10E9/L', am='QA089-L'), 0)) if lymphs else ''
        # o 7o par do Tester (nRBC/100 WBC x nRBC Absolute) nao e de prefixo: o nome curto "nRBC" nao esta no mapa -> ambar
        nrbc = '' if bio else str(xrun('mCEQForm.AvisoAnalitoCEQ', 'CAP', 'nRBC'))
        reg('C08 (G8) Pares %/# do mapa: o nome curto mostra o canônico e o par; com a unidade do "#" a gravação pede '
            'confirmação de unidade' + (' (Bio: sem par %/#; o prefixo Lactate / Lactate Dehydrogenase também avisa)' if bio
                                        else ' (Lymphs x 10E9/L -> LYMPH%)'),
            ok8 and (bio or (len(pares) >= 6 and lymphs is not None and 'LYMPH%' in cf_ly and 'Lymphs Absolute' in cf_ly
                             and nrbc.startswith('AMBAR|'))),
            {'pares': len(pares), 'detalhe': res8, 'lymphs_x10E9': cf_ly, 'nRBC_curto': nrbc[:80]})

        # ------------------------------------------------------------ C10 PROVEDOR_ALTERADO
        a3 = aud0()
        xl.EnableEvents = True
        if bio:
            p.Activate()
            p.Unprotect(ii.SENHA)
            p.Range('C3').Value = canon
            q.ex.esperar()
            antes_p = txt(p.Range('M4').Value)
            outro = 'Controllab' if antes_p.upper() != 'CONTROLLAB' else 'CAP'
            p.Range('M4').Value = outro
            q.ex.esperar()
            p.Range('M4').Value = antes_p
            q.ex.esperar()
            depois_p = txt(p.Range('M4').Value)
        else:
            e.Activate()
            e.Unprotect(ii.SENHA)
            e.Range('R4').Select()
            antes_p = txt(e.Range('R4').Value)
            outro = 'Controllab' if antes_p.upper() != 'CONTROLLAB' else 'CAP'
            e.Range('R4').Value = outro
            q.ex.esperar()
            e.Range('R4').Select()
            e.Range('R4').Value = antes_p
            q.ex.esperar()
            depois_p = txt(e.Range('R4').Value)
        xl.EnableEvents = False
        tr10 = [d for a, d in aud(a3) if a == 'PROVEDOR_ALTERADO']
        reg('C10 PROVEDOR_ALTERADO com antes e depois (' + ('Painel M4, por analito' if bio else 'Estatística R4, todos os '
            'analitos') + '); o provedor volta ao de antes',
            len(tr10) == 2 and f'{antes_p} {SETA} {outro}' in tr10[0] and f'{outro} {SETA} {antes_p}' in tr10[1]
            and depois_p == antes_p, {'antes': antes_p, 'trilha': tr10})

        # ------------------------------------------------------------ V03 a janela (modulo de teste so na copia)
        cod = '''Public Function QA089_Janela(ByVal nome As String) As String
    Dim f As Object, s As String
    Set f = VBA.UserForms.Add("frmCEQ")
    s = "corr=" & f.Controls("btnCorrigir").Enabled & "|exc=" & f.Controls("btnExcluir").Enabled
    f.Controls("cboAnalito").Text = nome
    s = s & "|idx=" & f.Controls("cboAnalito").ListIndex & "|cor=" & f.Controls("lblMsg").ForeColor & "|msg=" & f.Controls("lblMsg").Caption
    Unload f
    QA089_Janela = s
End Function
'''
        mod = q.wb.VBProject.VBComponents.Add(1)
        mod.Name = 'mQA089'
        mod.CodeModule.AddFromString(cod)
        jan = str(xrun('mQA089.QA089_Janela', nome_cap))
        reg('V03 Janela: Corrigir/Excluir desabilitados sem seleção; analito escolhido da lista -> aviso azul com o '
            'canônico no lblMsg',
            'corr=Falso' in jan.replace('False', 'Falso') and 'exc=Falso' in jan.replace('False', 'Falso')
            and f'msg=No QC_INI: {canon}.' in jan and f'cor={31 + 72 * 256 + 121 * 65536}' in jan, {'janela': jan})

        # ------------------------------------------------------------ V01
        import instalar_frm_inativar as fi
        import ceq_form_spec as cq
        comp = q.wb.VBProject.VBComponents('frmCEQ')
        dif = fi._diferencas(comp, ii.codigo(cq.FONTE_CODIGO), cq.FORM, cq.CONTROLES)
        aj = auditoria_janelas.auditar('frmCEQ', cq.FORM, cq.CONTROLES)
        reg('V01 frmCEQ igual à especificação (botões Corrigir/Excluir, lista com a linha); auditoria da janela com os textos '
            'novos: nada cortado nem sobreposto', not dif and not aj, {'diferencas': dif, 'auditoria': aj})
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
    print(f'\n=== {produto} ADR-089 (CEQ: corrigir/excluir, consolidação, mapa, trilha; OC-0051): '
          f'{len(RES) - len(falhas)} PASS / {len(falhas)} FAIL ===', flush=True)
    return falhas


if __name__ == '__main__':
    t0 = time.time()
    falhas = executar(sys.argv[1], os.path.abspath(sys.argv[2]), sys.argv[3] if len(sys.argv) > 3 else None)
    print(f'tempo {time.time() - t0:.0f}s')
    sys.exit(1 if falhas else 0)
