# -*- coding: utf-8 -*-
"""instalar_adr090.py -- ADR-090 (04/11/2026): lote 3 de correcoes da auditoria (ciclo 2026-10-31_ciclo01 e reteste 02).
Nenhum item muda regra de negocio: avisos, rotulos, formatos, trilha e protecao contra digitacao.

Uso:  python instalar_adr090.py <Bioquimica|Hematologia> <arquivo.xlsm> [--sem-salvar]

Idempotente (a 2a execucao registra 0 mudancas). Passos:
  1. janela frmCEQ pela especificacao (ceq_form_spec.py): rotulo "Valor-alvo ¹" e a nota do alvo no lblSub (OC-0031);
  2. todo o VBA atual das fontes (codigo_atual.py), com compilacao exigida: mEstatistica (n do periodo inteiro em
     Eng_Saida!AF1, OC-0012), mIntegracao (EQUIPAMENTO preenchido pelo Painel em ambar com nota, OC-0019; ResumoQA por
     grupo, OC-0024; RESULTADO_MANUAL_RECUSADO / _EM_CONFLITO, OC-0023), mCEQForm (sessao, OC-0033; sem alvo, OC-0031;
     listas sem Uso = NAO, OC-0034; par sem espaco, OC-0052; coluna de formula sobrescrita, OC-0053; antes pela
     EQA_Base, OC-0054; rodada sem o ano repetido, OC-0055), mApp + Painel.cls (De/Ate -> Personalizado, OC-0040);
  3. planilhas:
     - Painel A39: o aviso do corte de 180 corridas cita os indicadores do Painel e o n do periodo inteiro (OC-0012);
       J7:J9 "Reduzir para caber" (OC-0043); lista do K4 com "Personalizado" e K4 reduzir para caber (OC-0040);
     - Estatistica: coluna AT com FIXED no lugar de TEXT(x,"0.00") (OC-0036); A14:B133 centralizado na vertical
       (OC-0044); Hematologia: coluna A vazia sem "0" (OC-0049);
     - Principal - Resultados A5 com FIXED (OC-0050);
     - Digitar Resultados: instrucao com EQUIPAMENTO (OC-0019); RUN e DETALHE sem o "0" e DETALHE com acentos (OC-0021);
     - EQA.CAP_Dados / EQA.Controllab_Dados: colunas de formula (O, P, S..X) com validacao que recusa digitacao e Z3
       com o aviso "FÓRMULA SOBRESCRITA" (OC-0053);
  4. fumaca (nada e gravado nas tabelas);
  5. devolve a protecao como estava e salva com os eventos desligados, conferindo a data de gravacao.
Qualquer falha para ANTES de salvar.
"""
import os
import re
import sys

AQUI = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, AQUI)
sys.path.insert(0, os.path.join(AQUI, '..', 'etapa1_multilote'))
import xlh  # noqa: E402
import instalar_integracao as ii  # noqa: E402
import codigo_atual  # noqa: E402
import instalar_frm_inativar  # noqa: E402
import ceq_form_spec  # noqa: E402
import camada_dados as cd  # noqa: E402
import painel as painel_def  # noqa: E402
from instalar_seguranca_usuarios import foto_protecao, devolver_protecao  # noqa: E402

MUDANCAS = []
A39_ANTIGO = '"⚠ "&engCortadas&" corrida(s) antigas do período fora do gráfico (máx. 180). "'
A39_NOVO = ('"⚠ "&engCortadas&" corrida(s) antigas do período fora do gráfico e dos indicadores do Painel (máx. 180). '
            'Período inteiro: "&Eng_Saida!$AF$1&" — veja a aba Estatística. "')
ABAS_CEQ = {'EQA.CAP_Dados': 'CAP', 'EQA.Controllab_Dados': 'Controllab'}
COLS_FORMULA = ['O', 'P', 'S', 'T', 'U', 'V', 'W', 'X']       # Bias, |Bias| e as de apoio (montar_modulo_eqa + ADR-071)
VAL_RECUSA = '=1=0'                                          # personalizada que recusa toda digitacao (sem funcao: lingua neutra)
VAL_TITULO = 'Coluna calculada'
VAL_MSG = ('Esta coluna é calculada por fórmula (Bias, |Bias| e colunas de apoio). Não digite aqui: corrija o Seu '
           'Resultado ou o Valor-alvo da linha.')
AVISO_FORMULA = ('⚠ FÓRMULA SOBRESCRITA em Bias (%) ou |Bias| (%): o bias usaria o valor digitado — restaure a '
                 'fórmula da coluna')
PREFIXO_Z3 = ('=IF(SUMPRODUCT(($D$2:$D$5000<>"")*((NOT(ISFORMULA($O$2:$O$5000)))+(NOT(ISFORMULA($P$2:$P$5000)))))>0,'
              f'"{AVISO_FORMULA}",')
ROTULO_PERS = painel_def.PERSONALIZADO
ABA_MAN = 'Digitar Resultados'


def log(m):
    print(m, flush=True)


def norm(f):
    return str(f or '').replace('_xlfn.', '').replace(' ', '').upper()


def setf(rng, f, oque, formula2=False):
    atual = rng.Formula2 if formula2 else rng.Formula
    if norm(atual) != norm(f):
        if formula2:
            rng.Formula2 = f
        else:
            rng.Formula = f
        MUDANCAS.append(oque)


def text_para_fixed(f):
    """TEXT(e,"0.00") -> IFERROR(FIXED(e,2,TRUE),e): o mesmo resultado para numero, texto ("SEM EP") e erro, com o
    separador do idioma do Excel (o codigo "0.00" do TEXT e lido no idioma: pt-BR "004")."""
    alvo = ',"0.00")'
    while True:
        i = f.upper().find('TEXT(')
        while i >= 0 and i > 0 and (f[i - 1].isalnum() or f[i - 1] in '_.'):
            i = f.upper().find('TEXT(', i + 1)
        if i < 0:
            return f
        nivel, j = 0, i + 4
        while j < len(f):
            if f[j] == '(':
                nivel += 1
            elif f[j] == ')':
                nivel -= 1
                if nivel == 0:
                    break
            elif f[j] == '"':
                j = f.index('"', j + 1)
            j += 1
        trecho = f[i:j + 1]
        if not trecho.endswith(alvo):
            raise SystemExit(f'TEXT( inesperado na coluna AT: {trecho[:120]!r} -- nada foi salvo')
        e = trecho[5:-len(alvo)]
        f = f[:i] + f'IFERROR(FIXED({e},2,TRUE),{e})' + f[j + 1:]


# ------------------------------------------------------------------ Painel
def painel(wb):
    p = wb.Worksheets('Painel')
    f = str(p.Range('A39').Formula)
    if A39_NOVO not in f:
        if f.count(A39_ANTIGO) != 1:
            raise SystemExit(f'Painel!A39 sem o aviso de corte esperado: {f[-400:]!r} -- nada foi salvo')
        p.Range('A39').Formula = f.replace(A39_ANTIGO, A39_NOVO)
        MUDANCAS.append('Painel A39 (aviso do corte cita os indicadores e o n do período)')
    j = p.Range('J7:J9')
    if j.ShrinkToFit is not True or j.WrapText is not False:
        j.WrapText = False
        j.ShrinkToFit = True
        MUDANCAS.append('Painel J7:J9 reduzir para caber')
    k4 = p.Range('K4')
    if k4.ShrinkToFit is not True:
        k4.ShrinkToFit = True                                   # "Tudo (sem filtro)" ja passava 6 pt da largura
        MUDANCAS.append('Painel K4 reduzir para caber')
    v = k4.Validation
    f1 = str(v.Formula1)
    sep = p.Application.International[5 - 1]
    itens = f1.split(sep)
    if ROTULO_PERS not in itens:
        if 'Tudo' not in f1:
            raise SystemExit(f'Painel!K4: lista inesperada {f1!r} -- nada foi salvo')
        f1 = sep.join(i for i in itens if not i.startswith('Personalizado'))
        props = {n: getattr(v, n) for n in ('IgnoreBlank', 'InCellDropdown', 'InputTitle', 'InputMessage', 'ErrorTitle',
                                             'ErrorMessage', 'ShowInput', 'ShowError')}
        tipo, alerta = v.Type, v.AlertStyle
        v.Delete()
        k4.Validation.Add(tipo, alerta, 1, f1 + sep + ROTULO_PERS)
        for n, x in props.items():
            try:
                setattr(k4.Validation, n, x)
            except Exception:                              # noqa: BLE001 -- titulo vazio etc.
                pass
        MUDANCAS.append('Painel K4 (lista com Personalizado)')


# ------------------------------------------------------------------ Estatistica
def estatistica(wb, produto):
    e = wb.Worksheets('Estatística')
    hema = produto.startswith('Hema')
    n_at = n_a = 0
    for r in range(14, 134):
        c = e.Cells(r, 46)                                       # AT
        f = str(c.Formula)
        if f.startswith('=') and 'TEXT(' in f.upper().replace('_XLFN.', ''):
            c.Formula = text_para_fixed(f)
            n_at += 1
        a = e.Cells(r, 1)
        fa = str(a.Formula)
        if hema:
            m = re.fullmatch(r"=(Especificações!\$A\$\d+)", fa)
            if m:
                a.Formula = f'=IF({m.group(1)}="","",{m.group(1)})'
                n_a += 1
            elif fa and not re.fullmatch(r'=IF\((Especificações!\$A\$\d+)="","",\1\)', fa):
                raise SystemExit(f'Estatística!A{r} inesperada: {fa!r} -- nada foi salvo')
    if n_at:
        MUDANCAS.append(f'Estatística AT: FIXED no lugar de TEXT "0.00" ({n_at} linha(s))')
    if n_a:
        MUDANCAS.append(f'Estatística A (Hematologia): linha vazia sem "0" ({n_a} linha(s))')


def estatistica_vertical(wb):
    e = wb.Worksheets('Estatística')
    n = 0
    for r in range(14, 134):
        if e.Cells(r, 1).HasFormula:
            ab = e.Range(e.Cells(r, 1), e.Cells(r, 2))
            if ab.VerticalAlignment != -4108:
                ab.VerticalAlignment = -4108
                n += 1
    if n:
        MUDANCAS.append(f'Estatística A:B centro vertical ({n} linha(s))')


# ------------------------------------------------------------------ Principal / Digitar Resultados
def saidas(wb):
    nome = 'Principal - Resultados'
    wp = wb.Worksheets(nome)
    setf(wp.Range('A5'), cd.formulas_resumo()[nome], f'{nome} A5 (FIXED)', formula2=True)
    ws = wb.Worksheets(ABA_MAN)
    _, tit, sub, instr = [a for a in cd.ABAS if a[0] == ABA_MAN][0]
    if [str(ws.Range(f'A{i}').Value or '') for i in range(1, 5)] != [tit, sub] + list(instr[:2]):
        cd.textos_topo(wb, ABA_MAN)
        MUDANCAS.append(f'{ABA_MAN}: instrução com EQUIPAMENTO')
    lo = ws.ListObjects('tblResultados_Manuais')
    for col, fmt in (('RUN', cd.FMT_RUN),):                # o DETALHE perde o "0" pela formula (&"")
        rg = lo.ListColumns(col).DataBodyRange
        if str(rg.NumberFormat) != fmt:
            rg.NumberFormat = fmt
            MUDANCAS.append(f'{ABA_MAN} {col}: formato sem o 0')
    rg = lo.ListColumns('DETALHE').DataBodyRange
    if norm(rg.Cells(1, 1).Formula).replace('[@[ID_REGISTRO]]', '[@ID_REGISTRO]') != norm(cd.FORMULA_DETALHE):
        # como o camada_dados: [@X] por COM so entra como Tabela[[#This Row],[X]] (Formula2)
        rg.Formula2 = re.sub(r'\[@([^\]]+)\]', lambda m: f'{lo.Name}[[#This Row],[{m.group(1)}]]', cd.FORMULA_DETALHE)
        MUDANCAS.append(f'{ABA_MAN} DETALHE: acentos')


# ------------------------------------------------------------------ CEQ
def abas_ceq(wb):
    for aba in ABAS_CEQ:
        ws = wb.Worksheets(aba)
        z3 = str(ws.Range('Z3').Formula)
        if not norm(z3).startswith(norm(PREFIXO_Z3)):
            if not norm(z3).startswith(norm('=IF(N($Z$4)>N(EQA_Base!$AA$1),')):
                raise SystemExit(f'{aba}!Z3 inesperada (falta o ADR-089): {z3[:160]!r} -- nada foi salvo')
            ws.Range('Z3').Formula = PREFIXO_Z3 + z3[1:] + ')'
            MUDANCAS.append(f'{aba} Z3 (fórmula sobrescrita)')
        for col in COLS_FORMULA:
            rg = ws.Range(f'{col}2:{col}5000')
            ok = True
            for c in (rg.Cells(1, 1), rg.Cells(rg.Rows.Count, 1)):
                try:
                    vv = c.Validation
                    ok = ok and vv.Type == 7 and str(vv.Formula1).replace(' ', '') == VAL_RECUSA and str(vv.ErrorMessage) == VAL_MSG
                except Exception:                              # noqa: BLE001 -- sem validacao
                    ok = False
            if not ok:
                rg.Validation.Delete()
                rg.Validation.Add(7, 1, 1, VAL_RECUSA)         # xlValidateCustom, xlValidAlertStop
                rg.Validation.ErrorTitle = VAL_TITULO
                rg.Validation.ErrorMessage = VAL_MSG
                rg.Validation.ShowError = True
                MUDANCAS.append(f'{aba} {col}: recusa digitação (coluna de fórmula)')


def eng_saida(wb):
    eng = wb.Worksheets('Eng_Saida')
    for c, ok in (('AE1', ('', 'n do período:')), ('AF1', None)):
        v = eng.Range(c).Value
        if ok is not None and str(v or '') not in ok:
            raise SystemExit(f'Eng_Saida!{c} ocupada: {v!r} -- nada foi salvo')
        if c == 'AF1' and v not in (None, '') and not str(v).startswith('N1 '):
            raise SystemExit(f'Eng_Saida!AF1 ocupada: {v!r} -- nada foi salvo')


# ------------------------------------------------------------------ fumaca
def fumaca(ex, wb, produto):
    nome = wb.Name
    comp = wb.VBProject.VBComponents('frmCEQ')
    dif = instalar_frm_inativar._diferencas(comp, ii.codigo(ceq_form_spec.FONTE_CODIGO), ceq_form_spec.FORM,
                                            ceq_form_spec.CONTROLES)
    if dif:
        raise SystemExit(f'fumaca: frmCEQ diferente da especificacao: {dif} -- nada foi salvo')
    ex.xl.Calculation = -4105
    ex.xl.Calculate()
    for aba in ABAS_CEQ:
        z3 = str(wb.Worksheets(aba).Range('Z3').Value or '')
        if 'FÓRMULA SOBRESCRITA' in z3:
            raise SystemExit(f'fumaca: {aba}!Z3 acusa fórmula sobrescrita na cópia limpa: {z3!r}')
        log(f'   {aba}: Z3 = {z3!r}')
    e = wb.Worksheets('Estatística')
    at = [str(e.Cells(r, 46).Text) for r in range(14, 134) if str(e.Cells(r, 46).Text)]
    if any(re.match(r'^0\d\d / ', t) for t in at) or not any(',' in t for t in at):
        raise SystemExit(f'fumaca: Estatística AT ainda com o TEXT do idioma: {at[:3]}')
    log(f'   Estatística AT14 = {at[0] if at else ""!r}')
    a5 = str(wb.Worksheets('Principal - Resultados').Range('A5').Text)
    if re.search(r'Total: \d+,0', a5):
        raise SystemExit(f'fumaca: Principal A5 = {a5!r}')
    log(f'   Principal A5 = {a5[:60]!r}')
    rq = str(ex.run(f"'{nome}'!mIntegracao.ResumoQA", teto=120))
    if 'GOVERNANÇA' in rq:
        raise SystemExit(f'fumaca: ResumoQA ainda com o texto antigo: {rq!r}')
    log(f'   ResumoQA = {rq[:120]!r}')
    ra = str(ex.run(f"'{nome}'!mCEQForm.RodadaAno", 'C-A 2025', '2025', teto=60))
    if ra != 'C-A 2025':
        raise SystemExit(f'fumaca: RodadaAno = {ra!r}')
    sv = str(ex.xl.Run(f"'{nome}'!mCEQForm.SalvarCEQ", ['CAP', 'QA-090', '2026', 'x', 'y', '1', '', '', '', '', '', 'A',
                                                         '', '', '', '']))
    if not sv.startswith('ERRO|0|Sessão encerrada'):
        raise SystemExit(f'fumaca: SalvarCEQ sem sessao nao recusou: {sv!r}')
    log(f'   SalvarCEQ sem sessão = {sv[:60]!r}')


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
        log('1. janela frmCEQ (nota do valor-alvo)')
        if instalar_frm_inativar.criar_form(wb, log, nome='frmCEQ', fonte=ceq_form_spec.FONTE_CODIGO,
                                            form=ceq_form_spec.FORM, controles=ceq_form_spec.CONTROLES):
            MUDANCAS.append('frmCEQ')
        log('2. VBA (todo o codigo atual das fontes: codigo_atual.py)')
        codigo_atual.instalar(ex, wb, produto, log)
        ex.xl.Calculation = -4135
        log('3. planilhas')
        eng_saida(wb)
        painel(wb)
        estatistica(wb, produto)
        estatistica_vertical(wb)
        saidas(wb)
        abas_ceq(wb)
        log('4. fumaca')
        fumaca(ex, wb, produto)
        ex.xl.EnableEvents = False
        ex.xl.Calculation = -4105
        devolver_protecao(wb, foto)
        if estrutura:
            wb.Protect(ii.SENHA, True, False)
        log(f'MUDANCAS {len(MUDANCAS)}' + (': ' + '; '.join(MUDANCAS) if MUDANCAS else ''))
        if salvar:
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
