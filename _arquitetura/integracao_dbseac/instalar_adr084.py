# -*- coding: utf-8 -*-
"""instalar_adr084.py -- ADR-084 (28/10/2026): calibracao pelo ID do resultado, indicador de processamento, motivo
livre na inativacao e status sem "alvo TESTE".

Uso:  python instalar_adr084.py <Bioquimica|Hematologia> <arquivo.xlsm> [--sem-salvar]

Idempotente (a 2a execucao registra 0 mudancas). Passos:
  1. janelas declarativas: frmInativar (motivo livre) e frmCarregando (indicador) -- instalar_frm_*.criar_form;
  2. todo o VBA atual das fontes (codigo_atual.py: mCalibracao, mCarregando, Registros.cls...) -- EXIGE compilacao;
  3. aba Registros: K3 "ID do resultado" (digitavel; o mCalibracao preenche a linha) e L3 "Equipamento", com o mesmo
     formato do cabecalho/das linhas de J; A2 explica o uso; regView passa a B4:L203; nomes regID (K) e regEquip (L);
     RegistrosStore ganha os cabecalhos J1/K1 (cada lote guarda as 11 colunas; blocos antigos ficam com K:L vazias);
  4. Calc: as colunas do icone "Calibracao" (formula com regCalib) passam a exigir o equipamento do grafico quando a
     linha da aba Registros tem equipamento -- (COUNTIFS(...,regEquip,"") + COUNTIFS(...,regEquip,engEquip)) > 0. O
     resto da regra NAO muda: mesmo analito, mesmo dia, "Sim", y = Media do lote do nivel;
  5. Painel: o status sem " (alvo TESTE)" vem do instalar_adr078 (status_nivel/q4/a39), reaplicado aqui;
  6. devolve a protecao de cada aba e a da estrutura como estavam. Qualquer falha para ANTES de salvar.
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
import instalar_frm_carregando  # noqa: E402
import instalar_adr078 as a78  # noqa: E402
from instalar_seguranca_usuarios import foto_protecao, devolver_protecao  # noqa: E402

MUDANCAS = []
R0, R1 = 4, 203
CAB_ID = 'ID do resultado'
CAB_EQ = 'Equipamento'
TEXTO_A2 = ('="REGISTROS DO LOTE EM USO:  "&loteAtivo&"   ·   Calibração: digite na coluna K o ID do resultado '
            'medido logo depois da calibração (ex.: BIO-365293 ou só o número) -- data, analito, RUN e equipamento '
            'são preenchidos e o ícone de calibração aparece na média do lote, no gráfico desse equipamento. '
            '\'Foi calibrado? = Sim\' digitado à mão continua valendo. Repetições/não conformes: aba Inativar."')
# Calc: o COUNTIFS da calibracao (como o instalar_integracao deixou) -> o mesmo + equipamento
RX_CAL = re.compile(r'COUNTIFS\((regAnalito,selAnalito,regData,">="&INT\(N\(\$C(\d+)\)\),regData,"<"&\(INT\(N\(\$C\2\)\)\+1\),'
                    r'regCalib,"Sim")\)>0')


def log(m):
    print(m, flush=True)


def nome(wb, nm, ref):
    alvo = ref.replace("'", '').replace('$', '').upper()
    for n in list(wb.Names):
        if n.Name == nm:
            if str(n.RefersTo).replace("'", '').replace('$', '').upper() == alvo:
                return
            n.Delete()
            break
    wb.Names.Add(nm, ref)
    MUDANCAS.append(f'nome {nm}')


def registros(wb):
    ws = wb.Worksheets('Registros')
    if str(ws.Range('J3').Value or '').strip() != 'Status':
        raise SystemExit(f'Registros!J3 inesperado ({ws.Range("J3").Value!r}): layout diferente -- nada foi feito')
    for cel, txt in (('K3', CAB_ID), ('L3', CAB_EQ)):
        if str(ws.Range(cel).Value or '') != txt:
            ws.Range('J3').Copy(ws.Range(cel))
            ws.Range(cel).Value = txt
            MUDANCAS.append(f'Registros {cel}')
    corpo = ws.Range(f'K{R0}:L{R1}')
    if ws.Range(f'K{R0}').NumberFormat != '@':
        ws.Range(f'J{R0}:J{R1}').Copy()
        corpo.PasteSpecial(-4122)                      # xlPasteFormats
        wb.Application.CutCopyMode = False
        ws.Range(f'K{R0}:K{R1}').NumberFormat = '@'    # ID e texto (BIO-365293; o numero puro nao vira 365293,0)
        MUDANCAS.append('Registros K:L formato')
    if ws.Range(f'K{R0}').Locked:
        ws.Range(f'K{R0}:K{R1}').Locked = False        # digitavel (como as demais colunas de entrada)
        MUDANCAS.append('Registros K destravada')
    for col, larg in (('K', 18), ('L', 16)):
        if abs(float(ws.Columns(col).ColumnWidth) - larg) > 0.5:
            ws.Columns(col).ColumnWidth = larg
            MUDANCAS.append(f'Registros largura {col}')
    if a78.norm(ws.Range('A2').Formula) != a78.norm(TEXTO_A2):
        ws.Range('A2').Formula = TEXTO_A2
        MUDANCAS.append('Registros A2')
    nome(wb, 'regView', f'=Registros!$B${R0}:$L${R1}')
    nome(wb, 'regID', f'=Registros!$K${R0}:$K${R1}')
    nome(wb, 'regEquip', f'=Registros!$L${R0}:$L${R1}')
    st = wb.Worksheets('RegistrosStore')
    for cel, txt in (('J1', 'ID'), ('K1', 'Equip')):
        if str(st.Range(cel).Value or '') != txt:
            st.Range(cel).Value = txt
            MUDANCAS.append(f'RegistrosStore {cel}')


def calc(wb):
    c = wb.Worksheets('Calc')
    linha3 = c.Range('A3:EZ3').Formula[0]
    alvo = [k + 1 for k, f in enumerate(linha3) if 'regCalib' in str(f or '')]
    if not alvo:
        raise SystemExit('Calc: nenhuma coluna com regCalib (icone de calibracao) -- layout diferente, nada foi feito')
    feitas = []
    for k in alvo:
        f3 = str(c.Cells(3, k).Formula)
        if 'regEquip' in f3:
            continue
        m = RX_CAL.search(f3)
        if not m:
            raise SystemExit(f'Calc {a78.col(k)}3: formula de calibracao inesperada -- nada foi feito: {f3[:200]}')
        novo = RX_CAL.sub(lambda mm: f'(COUNTIFS({mm.group(1)},regEquip,"")+COUNTIFS({mm.group(1)},regEquip,'
                                     f'TRIM(""&engEquip)))>0', f3)
        c.Range(c.Cells(3, k), c.Cells(182, k)).Formula = novo        # relativo: $C3 -> $C4 ... como o FillDown
        feitas.append(a78.col(k))
    if feitas:
        MUDANCAS.append('Calc calibracao+equipamento ' + ','.join(feitas))
    return alvo


CAB_UC = ['LOGIN', 'NOME', 'CARGO', 'PAPEL', 'SENHA_HASH', 'ASSINATURA', 'CONSELHO', 'SITUACAO', 'SINCRONIZADO_EM']
CFG_CAMINHO = ('CAMINHO_USUARIOS', '', 'ADR-084: caminho do cadastro central de profissionais (QC_INI_cadastrar '
               'usuários.xlsm). Vazio = a mesma pasta deste arquivo.')


def usuarios_central(wb):
    """Aba tecnica Usuarios_Central (copia do cadastro central; o mUsuariosCentral grava). Oculta e protegida."""
    nomes = [w.Name for w in wb.Worksheets]
    if 'Usuarios_Central' not in nomes:
        ws = wb.Worksheets.Add(After=wb.Worksheets(wb.Worksheets.Count))
        ws.Name = 'Usuarios_Central'
        MUDANCAS.append('aba Usuarios_Central criada')
    ws = wb.Worksheets('Usuarios_Central')
    if str(ws.Range('A1').Value or '') != 'Cópia do cadastro central em:':
        ws.Range('A1').Value = 'Cópia do cadastro central em:'
        ws.Range('C1').Value = 'arquivo:'
        ws.Range('A2').Value = ('ADR-084: gravada pelo mUsuariosCentral.Sincronizar (login). Não editar: altere o '
                                'cadastro no arquivo central.')
        MUDANCAS.append('Usuarios_Central A1:A2')
    atual = [str(v or '') for v in ws.Range('A3:I3').Value[0]]
    if atual != CAB_UC:
        ws.Range('A3:I3').Value = tuple(CAB_UC)
        ws.Range('A3:I3').Font.Bold = True
        ws.Columns('E').NumberFormat = '@'
        MUDANCAS.append('Usuarios_Central cabecalho')
    if ws.Visible != 2:
        ws.Visible = 2                                   # xlSheetVeryHidden
        MUDANCAS.append('Usuarios_Central oculta')
    # Usuarios local: o rotulo do hash
    us = wb.Worksheets('Usuarios')
    if str(us.Range('D3').Value or '') != 'Hash da senha (PBKDF2)':
        us.Range('D3').Value = 'Hash da senha (PBKDF2)'
        MUDANCAS.append('Usuarios D3')
    a2 = ('Usuários locais (login, senha, papel). Os profissionais do cadastro central (QC_INI_cadastrar usuários.xlsm) '
          'entram com o login e a senha de lá. Senha gravada só como hash PBKDF2-SHA256 com salt. Só o ADM cadastra.')
    if str(us.Range('A2').Value or '') != a2:
        us.Range('A2').Value = a2
        MUDANCAS.append('Usuarios A2')
    # chave do caminho do cadastro central
    lo = wb.Worksheets('Cfg_Integracao').ListObjects('tblConfigIntegracao')
    ks = [str(r[0] or '').strip().upper() for r in lo.DataBodyRange.Value]
    if CFG_CAMINHO[0] not in ks:
        lin = lo.ListRows.Add()
        lin.Range.Cells(1, 1).Value = CFG_CAMINHO[0]
        lin.Range.Cells(1, 2).Value = CFG_CAMINHO[1]
        if lo.ListColumns.Count >= 3:
            lin.Range.Cells(1, 3).Value = CFG_CAMINHO[2]
        MUDANCAS.append('Cfg CAMINHO_USUARIOS')


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
        nlv = 2 if produto.startswith('Bio') else 3
        log('1. janelas: frmInativar (motivo livre) e frmCarregando (indicador)')
        n = instalar_frm_inativar.criar_form(wb, log) + instalar_frm_carregando.criar_form(wb, log)
        if n:
            MUDANCAS.append(f'{n} janela(s)')
        log('2. VBA (todo o codigo atual das fontes: codigo_atual.py)')
        codigo_atual.instalar(ex, wb, produto, log)
        ex.xl.Calculation = -4135
        registros(wb)
        log('3. Registros: K "ID do resultado", L "Equipamento", regView B4:L203, regID, regEquip; RegistrosStore J1/K1')
        cols = calc(wb)
        log(f'4. Calc: icone de calibracao com equipamento nas colunas {[a78.col(k) for k in cols]}')
        usuarios_central(wb)
        log('4b. Usuarios_Central (oculta), Cfg CAMINHO_USUARIOS, rótulos da aba Usuarios (PBKDF2)')
        a78.MUDANCAS.clear()
        a78.painel(wb, a78.mapa(wb, nlv), nlv)
        MUDANCAS.extend(a78.MUDANCAS)
        log('5. Painel: status J7.., Q4 e A39 sem "alvo TESTE"')
        ex.xl.Calculation = -4105
        ex.xl.Calculate()
        ex.esperar()
        p = wb.Worksheets('Painel')
        st = [str(p.Range(f'J{6 + i}').Value) for i in range(1, nlv + 1)]
        if any('TESTE' in s.upper() for s in st):
            raise SystemExit(f'fumaca: status ainda com TESTE: {st} -- nada foi salvo')
        log(f'   fumaca: status {st}')
        log(f'MUDANCAS {len(MUDANCAS)}' + (': ' + '; '.join(MUDANCAS[:14]) if MUDANCAS else ''))
        devolver_protecao(wb, foto)
        uc = wb.Worksheets('Usuarios_Central')
        if not uc.ProtectContents:
            uc.Protect(ii.SENHA, True, True, True)      # aba tecnica: so o VBA grava (LiberarEscrita)
        if estrutura:
            wb.Protect(ii.SENHA, True, False)
        if salvar:
            wb.Save()
            log('salvo')
    finally:
        ex.fechar()


if __name__ == '__main__':
    a = sys.argv[1:]
    if len(a) < 2:
        raise SystemExit(__doc__)
    main(a[0], a[1], salvar='--sem-salvar' not in a)
