# -*- coding: utf-8 -*-
"""instalar_adr089.py -- ADR-089 (04/11/2026): lote 2 de correcoes da auditoria (ciclo 2026-10-31_ciclo01): o CEQ.

Uso:  python instalar_adr089.py <Bioquimica|Hematologia> <arquivo.xlsm> [--sem-salvar]

Idempotente (a 2a execucao registra 0 mudancas). Passos:
  1. janela frmCEQ refeita pela especificacao (ceq_form_spec.py): botoes "Corrigir selecionado" e "Excluir
     selecionado", lista com a linha da aba (coluna oculta), confirmacoes do analito fora do mapa / unidade (OC-0027,
     OC-0029, OC-0030);
  2. todo o VBA atual das fontes (codigo_atual.py): mCEQForm (corrigir, excluir, avisos, trilha), mEQA (consolidacao
     com trilha EQA_BASE_CONSOLIDADA, instante numerico em EQA_Base!AA1, pendencia), modulo das abas EQA.CAP_Dados e
     EQA.Controllab_Dados (EQADados.cls: edicao direta = marca Z4 + CEQ_ALTERADO), mIntegracao (ATUALIZAR consolida o
     pendente), mEstatistica (PROVEDOR_ALTERADO no M4 da Bio), Estatistica.cls da Hema (PROVEDOR_ALTERADO no R4) e
     mLotes (OC-0051: troca do lote no Painel sem o UserInterfaceOnly da sessao); EXIGE compilacao;
  3. abas de digitacao do CEQ: Z2 = carimbo da EQA_Base por formula (recalcula sozinho; antes era UDF sem dependencia),
     Z3 avisa a edicao nao consolidada (marca Z4 mais nova que EQA_Base!AA1) antes do teste de quantidade de sempre,
     Z4 e AA1 com formato oculto, realce ambar do aviso novo (OC-0028);
  4. fumaca: frmCEQ igual a especificacao; os dois modulos de aba instalados; Z2 = Z1 da EQA_Base; sem pendencia na
     copia limpa; aviso azul para um nome do mapa e ambar para um nome fora dele; ConfirmacoesCEQ pede confirmacao fora
     do mapa. Nada e gravado na EQA_Base nem nas tabelas do CEQ.
  5. devolve a protecao de cada aba e a da estrutura como estavam; salva com os eventos desligados e confere a data de
     gravacao. Qualquer falha para ANTES de salvar.
"""
import os
import sys

AQUI = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, AQUI)
sys.path.insert(0, os.path.join(AQUI, '..', 'etapa1_multilote'))
import xlh  # noqa: E402
import instalar_integracao as ii  # noqa: E402
import codigo_atual  # noqa: E402
import instalar_frm_inativar  # noqa: E402
import ceq_form_spec  # noqa: E402
from instalar_seguranca_usuarios import foto_protecao, devolver_protecao  # noqa: E402

MUDANCAS = []
ABAS_CEQ = {'EQA.CAP_Dados': 'CAP', 'EQA.Controllab_Dados': 'Controllab'}
Z2 = '=IF(EQA_Base!$Z$1="","ainda nao consolidada",EQA_Base!$Z$1)'
AVISO_Z3 = '⚠ HÁ EDIÇÃO NÃO CONSOLIDADA — clique em Consolidar na EQA_Base'
PREFIXO_Z3 = f'=IF(N($Z$4)>N(EQA_Base!$AA$1),"{AVISO_Z3}",'
OCULTO = ';;;'
AMBAR = 0x00579C                                   # BGR de RGB(156, 87, 0)
CF_AMBAR = '=ESQUERDA($Z$3;1)="⚠"'                 # FormatConditions.Add usa a lingua da interface (como o ADR-034)


def log(m):
    print(m, flush=True)


def norm(f):
    return str(f or '').replace('_xlfn.', '').replace(' ', '').upper()


def z3_antigo(prov):
    return ('=IF(COUNTA($D$2:$D$5000)=COUNTIF(EQA_Base!$A$2:$A$5001,"%s"),"base em dia",'
            '"BASE DESATUALIZADA — clique em Consolidar na EQA_Base")' % prov)


def z3_novo(prov):
    return PREFIXO_Z3 + z3_antigo(prov)[1:] + ')'


def abas_ceq(wb):
    for aba, prov in ABAS_CEQ.items():
        ws = wb.Worksheets(aba)
        if str(ws.Range('Z1').Value or '') != 'ESTADO DA CONSOLIDAÇÃO':
            raise SystemExit(f'{aba}!Z1 inesperada: {ws.Range("Z1").Value!r} -- nada foi salvo')
        z2 = str(ws.Range('Z2').Formula)
        if norm(z2) != norm(Z2):
            if norm(z2) != norm('=mEQA.CarimboEQA()'):
                raise SystemExit(f'{aba}!Z2 inesperada: {z2!r} -- nada foi salvo')
            ws.Range('Z2').Formula = Z2
            MUDANCAS.append(f'{aba} Z2 (carimbo por formula)')
        z3 = str(ws.Range('Z3').Formula)
        if norm(z3) != norm(z3_novo(prov)):
            if norm(z3) != norm(z3_antigo(prov)):
                raise SystemExit(f'{aba}!Z3 inesperada: {z3[:160]!r} -- nada foi salvo')
            ws.Range('Z3').Formula = z3_novo(prov)
            MUDANCAS.append(f'{aba} Z3 (edição não consolidada)')
        if ws.Range('Z4').Formula not in ('', None) and not isinstance(ws.Range('Z4').Value, float):
            raise SystemExit(f'{aba}!Z4 ocupada: {ws.Range("Z4").Formula!r} -- nada foi salvo')
        if str(ws.Range('Z4').NumberFormat) != OCULTO:
            ws.Range('Z4').NumberFormat = OCULTO
            MUDANCAS.append(f'{aba} Z4 (marca de edição, oculta)')
        fc = ws.Range('Z3').FormatConditions
        tem = False
        for i in range(1, fc.Count + 1):
            try:
                tem = tem or '"⚠"' in str(fc(i).Formula1)
            except Exception:                              # noqa: BLE001
                pass
        if not tem:
            c = ws.Range(ws.Cells(3, 26), ws.Cells(3, 26)).FormatConditions.Add(2, None, CF_AMBAR)
            c.Font.Color = AMBAR
            c.Font.Bold = True
            MUDANCAS.append(f'{aba} Z3 (realce âmbar)')
    b = wb.Worksheets('EQA_Base')
    aa1 = b.Range('AA1')
    if aa1.Formula not in ('', None) and not isinstance(aa1.Value, float):
        raise SystemExit(f'EQA_Base!AA1 ocupada: {aa1.Formula!r} -- nada foi salvo')
    if str(aa1.NumberFormat) != OCULTO:
        aa1.NumberFormat = OCULTO
        MUDANCAS.append('EQA_Base AA1 (instante da consolidação, oculto)')


def fumaca(ex, wb):
    nome = wb.Name
    comp = wb.VBProject.VBComponents('frmCEQ')
    dif = instalar_frm_inativar._diferencas(comp, ii.codigo(ceq_form_spec.FONTE_CODIGO), ceq_form_spec.FORM,
                                            ceq_form_spec.CONTROLES)
    if dif:
        raise SystemExit(f'fumaca: frmCEQ diferente da especificacao: {dif} -- nada foi salvo')
    fonte = codigo_atual.normalizar(ii.codigo(os.path.join(codigo_atual.SRC, 'comum', 'EQADados.cls'))).lower()
    for aba in ABAS_CEQ:
        cm = wb.VBProject.VBComponents(wb.Worksheets(aba).CodeName).CodeModule
        if codigo_atual.normalizar(cm.Lines(1, cm.CountOfLines)).lower() != fonte:
            raise SystemExit(f'fumaca: modulo da aba {aba} diferente de EQADados.cls -- nada foi salvo')
    ex.xl.Calculation = -4105
    ex.xl.Calculate()
    z1 = str(wb.Worksheets('EQA_Base').Range('Z1').Value or '')
    for aba in ABAS_CEQ:
        ws = wb.Worksheets(aba)
        if str(ws.Range('Z2').Value or '') != (z1 or 'ainda nao consolidada'):
            raise SystemExit(f'fumaca: {aba}!Z2 = {ws.Range("Z2").Value!r}, EQA_Base!Z1 = {z1!r}')
        log(f'   {aba}: Z2 = Z1 da EQA_Base; Z3 = {ws.Range("Z3").Value!r}')
    pend = ex.run(f"'{nome}'!mEQA.ConsolidacaoPendente", teto=120)
    log(f'   ConsolidacaoPendente = {pend}')
    mapa = ex.run(f"'{nome}'!mEQA.MapaDoProvedor", 'CAP', teto=120)
    if not mapa:
        raise SystemExit('fumaca: o mapa do CAP veio vazio -- nada foi salvo')
    a1 = str(ex.run(f"'{nome}'!mCEQForm.AvisoAnalitoCEQ", 'CAP', mapa[0][0], teto=120))
    a2 = str(ex.run(f"'{nome}'!mCEQForm.AvisoAnalitoCEQ", 'CAP', 'QA-089 nome fora do mapa', teto=120))
    campos = ['CAP', 'QA', '2026', 'QA-089 nome fora do mapa', 'QA-01', '1', '', '', '', '', '', 'Acceptable', '', '',
              '', '']
    c2 = str(ex.xl.Run(f"'{nome}'!mCEQForm.ConfirmacoesCEQ", campos, 0))
    log(f'   aviso mapeado: {a1[:90]!r}; fora do mapa: {a2[:60]!r}; confirmação: {c2[:60]!r}')
    if not a1.startswith('AZUL|No QC_INI: ' + str(mapa[0][1])) or not a2.startswith('AMBAR|') or 'não está no mapa' not in c2:
        raise SystemExit('fumaca: avisos do analito no CEQ nao responderam como esperado -- nada foi salvo')


def main(produto, caminho, salvar=True):
    try:
        sys.stdout.reconfigure(errors='replace')
    except Exception:                                   # noqa: BLE001
        pass
    caminho = os.path.abspath(caminho)
    ex = xlh.Excel()
    log(f'EXCEL_PID {ex.pid}')
    try:
        wb = ex.abrir(caminho)
        if wb.ReadOnly:
            raise SystemExit(f'arquivo aberto SOMENTE LEITURA (outra instancia o segura): {caminho}')
        estrutura = bool(wb.ProtectStructure)
        foto = foto_protecao(wb)
        ii.desproteger_tudo(wb)
        ex.xl.EnableEvents = False
        log('1. janela frmCEQ (Corrigir/Excluir selecionado, avisos do analito)')
        if instalar_frm_inativar.criar_form(wb, log, nome='frmCEQ', fonte=ceq_form_spec.FONTE_CODIGO,
                                            form=ceq_form_spec.FORM, controles=ceq_form_spec.CONTROLES):
            MUDANCAS.append('frmCEQ')
        log('2. VBA (todo o codigo atual das fontes: codigo_atual.py)')
        codigo_atual.instalar(ex, wb, produto, log)
        ex.xl.Calculation = -4135
        log('3. abas de digitacao do CEQ: Z2 por formula, Z3 com a edicao nao consolidada, Z4/AA1 ocultas')
        abas_ceq(wb)
        log('4. fumaca')
        fumaca(ex, wb)
        ex.xl.EnableEvents = False
        ex.xl.Calculation = -4105
        devolver_protecao(wb, foto)
        if estrutura:
            wb.Protect(ii.SENHA, True, False)
        log(f'MUDANCAS {len(MUDANCAS)}' + (': ' + '; '.join(MUDANCAS) if MUDANCAS else ''))
        if salvar:
            # eventos desligados: com eles o Workbook_BeforeSave cancela o Save e engole o erro
            antes = os.path.getmtime(wb.FullName)
            wb.Save()
            if os.path.getmtime(wb.FullName) == antes:
                raise SystemExit('o Save nao chegou ao arquivo (data de gravacao igual) -- conferir')
            log('salvo')
    finally:
        ex.fechar()


if __name__ == '__main__':
    a = sys.argv[1:]
    if len(a) < 2:
        raise SystemExit(__doc__)
    main(a[0], a[1], salvar='--sem-salvar' not in a)
