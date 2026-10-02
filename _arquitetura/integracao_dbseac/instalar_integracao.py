# -*- coding: utf-8 -*-
"""instalar_integracao.py -- ADR-057: instala a NOVA ARQUITETURA DE DADOS num QC_INI.

    DB_SEAC --PQ--> DB (recebimento) -> DB_ORGANIZADO
                         + Digitar Resultados + Inativar + COMENTARIOS_TECNICOS
                         -> Principal - Resultados (tblCQ_Final, single source of truth) -> LJ + Estatistica

Uso: python instalar_integracao.py <Bioquimica|Hematologia> <arquivo.xlsm> [--sem-salvar]

Passos (todos conferidos; qualquer falha para ANTES de salvar):
  1. camada de dados (abas, tabelas, consultas, carga)      -> camada_dados.montar
  2. VBA: tira mBanco/mImportar/mOperacao/formularios; troca os modulos por
     src/<produto>/ (gerados por patch_vba.py a partir deste mesmo arquivo);
     eventos das abas de entrada
  3. nomes engXN*, engEquip, selEquipamento; Calc (X vermelho, calibracao pelo
     dia, guarda de equipamento); Painel (seletor de equipamento na Bioquimica)
  4. Estatistica: bloco RESULTADOS NAO CONFORMES / REPETICOES
  5. Registros: so calibracao (sai Rep 1-3) + RegistrosStore alinhada
  6. legado: abas Importar, Resultados, EQC_Dados (LEGADO ADR-034),
     DB_Resultados, Cfg_Status; nomes r*/capBanco; Inicio reapontado
  7. navegacao: '+ Lancar' vira 'Dados CQ'; botao ATUALIZAR DADOS
  8. atualizacao completa pelo proprio VBA (mIntegracao.AtualizarDados)
  9. salva; depois, com o Excel fechado, o X dos graficos vira vermelho (XML)
"""
import os
import re
import shutil
import subprocess
import sys
import time
import zipfile

AQUI = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, AQUI)
sys.path.insert(0, os.path.join(AQUI, '..', 'etapa1_multilote'))
import xlh  # noqa: E402
import pqlib  # noqa: E402
import tema  # noqa: E402
import camada_dados as cd  # noqa: E402
import config_produtos as cp  # noqa: E402

SENHA = 'qcini2025'
REMOVER_VBA = ['mBanco', 'mImportar', 'mOperacao', 'frmCorrida', 'frmMassa', 'frmExcluir', 'frmConfigEstatistica']
ABAS_LEGADO = ['Importar', 'Resultados', 'EQC_Dados', 'DB_Resultados', 'Cfg_Status']
NOMES_LEGADO = ['rRUN', 'rData', 'rNivel', 'rAnalito', 'rValor', 'rStatus', 'rLote', 'rFirst', 'rRunUnico', 'capBanco',
                'regRep1', 'regRep2', 'regRep3']
ENTRADA_ABAS = ['Digitar Resultados', 'Inativar', 'COMENTARIOS_TECNICOS']
NOME_X = 'Não conforme (X)'


def log(*a):
    print(*a, flush=True)


def codigo(caminho):
    with open(caminho, encoding='utf-8', newline='') as f:
        linhas = f.read().replace('\r\n', '\n').split('\n')
    return '\r\n'.join(l for l in linhas if not l.startswith('Attribute ')).strip('\r\n') + '\r\n'


def substituir_codigo(vbp, comp, texto, criar=True):
    nomes = [c.Name for c in vbp.VBComponents]
    if comp not in nomes:
        if not criar:
            raise SystemExit(f'componente {comp} nao existe')
        novo = vbp.VBComponents.Add(1)          # vbext_ct_StdModule
        novo.Name = comp
    cm = vbp.VBComponents(comp).CodeModule
    n = cm.CountOfLines
    if n:
        cm.DeleteLines(1, n)
    cm.AddFromString(texto)


def nome(wb, n, ref):
    try:
        wb.Names(n).Delete()
    except Exception:
        pass
    wb.Names.Add(n, ref)
    return wb.Names(n).RefersTo


def sheet(wb, n):
    for ws in wb.Worksheets:
        if ws.Name == n:
            return ws
    return None


def desproteger_tudo(wb):
    try:
        wb.Unprotect(SENHA)
    except Exception:
        pass
    for ws in wb.Worksheets:
        try:
            ws.Unprotect(SENHA)
        except Exception:
            pass


def col_letra(n):
    s = ''
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


# =============================================================================
#  2. VBA
# =============================================================================
def instalar_vba(wb, produto):
    vbp = wb.VBProject
    pasta = os.path.join(AQUI, 'src', 'bio' if produto.startswith('Bio') else 'hema')
    if not os.path.isdir(pasta) or not os.listdir(pasta):
        raise SystemExit(f'rode patch_vba.py antes: {pasta} vazio')
    existentes = [c.Name for c in vbp.VBComponents]
    for m in REMOVER_VBA:
        if m in existentes:
            vbp.VBComponents.Remove(vbp.VBComponents(m))
            log(f'  VBA removido: {m}')
    for arq in sorted(os.listdir(pasta)):
        comp = os.path.splitext(arq)[0]
        substituir_codigo(vbp, comp, codigo(os.path.join(pasta, arq)), criar=arq.endswith('.bas'))
    log(f'  VBA: {len(os.listdir(pasta))} modulos substituidos/criados')
    # eventos das abas de entrada
    for nm in ENTRADA_ABAS:
        ws = sheet(wb, nm)
        substituir_codigo(vbp, ws.CodeName, (
            "Option Explicit\r\n"
            "' ADR-057: aba de ENTRADA da camada de dados. O evento so prepara a linha\r\n"
            "' (ID normalizado, carimbo, caixa REGISTRAR - LJ marcada por padrao, ID\r\n"
            "' MAN_nnnn); quem decide o que o dado e, e o Power Query (DB_CQ_FINAL).\r\n"
            "Private Sub Worksheet_Change(ByVal Target As Range)\r\n"
            "    mIntegracao.EntradaMudou Me, Target\r\n"
            "End Sub\r\n"), criar=False)
    log('  VBA: eventos das abas de entrada')


# =============================================================================
#  3. nomes, Calc, Painel
# =============================================================================
def instalar_calc(wb, produto):
    bio = produto.startswith('Bio')
    nlv = 2 if bio else 3
    nome(wb, 'engXN1', '=Eng_Saida!$AD$3:$AF$182')
    nome(wb, 'engXN2', '=Eng_Saida!$AG$3:$AI$182')
    nome(wb, 'engXN3', '=Eng_Saida!$AJ$3:$AL$182')
    nome(wb, 'engEquip', '=Eng_Saida!$S$1')
    if bio:
        nome(wb, 'selEquipamento', '=Painel!$L$4')
    else:
        nome(wb, 'selEquipamento', '=""')
    eng = sheet(wb, 'Eng_Saida')
    for t in range(3):
        for s in range(3):
            eng.Cells(2, 30 + t * 3 + s).Value = f'N{t + 1} X{s + 1}'
    eng.Range('AD1').Value = 'X VERMELHO: resultado inativado com REGISTRAR - LJ (ADR-057) — nunca entra em cálculo'
    calc = sheet(wb, 'Calc')
    # bloco de 22 colunas por nivel a partir de F: rep = colunas 19..21 do bloco (X:Z), calib = 22 (AA)
    for t in range(nlv):
        c0 = 6 + t * 22
        for s in range(3):
            c = c0 + 18 + s
            letra = col_letra(c)
            calc.Range(f'{letra}3:{letra}182').Formula = (
                f'=IF(OR($B3="",$D3<>1),NA(),IF(ISNUMBER(INDEX(engXN{t + 1},$A3,{s + 1})),'
                f'INDEX(engXN{t + 1},$A3,{s + 1}),NA()))')
            calc.Cells(2, c).Value = f'N{t + 1} x{s + 1}'
        cc = col_letra(c0 + 21)
        f = calc.Range(f'{cc}3').Formula
        if 'regData,$C3,' not in f:
            raise SystemExit(f'Calc!{cc}3 inesperada: {f}')
        f = f.replace('regData,$C3,', 'regData,">="&INT(N($C3)),regData,"<"&(INT(N($C3))+1),')
        calc.Range(f'{cc}3:{cc}182').Formula = f
    # guarda do equipamento em Calc!B (o motor tem de ser do equipamento em tela)
    fb = calc.Range('B3').Formula
    if 'engEquip' not in fb:
        if '(""&engLote)<>(""&loteAnalise),' not in fb:
            raise SystemExit(f'Calc!B3 inesperada: {fb}')
        fb = fb.replace('(""&engLote)<>(""&loteAnalise),', '(""&engLote)<>(""&loteAnalise),(""&engEquip)<>(""&selEquipamento),')
        calc.Range('B3:B182').Formula = fb
    log('  Calc: X vermelho em engXN*, calibração pelo dia, guarda de equipamento')
    if bio:
        pai = sheet(wb, 'Painel')
        l3, l4 = pai.Range('L3'), pai.Range('L4')
        l3.Value = 'Equipamento'
        m3 = pai.Range('M3')
        l3.Font.Bold, l3.Font.Size, l3.Font.Color = True, m3.Font.Size or 9, m3.Font.Color
        l3.HorizontalAlignment = -4131
        if str(l4.Value or '') not in ('DIMENSION 1', 'DIMENSION 2'):
            l4.Value = cp.PRODUTOS['Bioquimica']['cfg']['EQUIPAMENTO_PADRAO'][0]
        l4.Validation.Delete()
        l4.Validation.Add(3, 1, 1, ','.join(cp.PRODUTOS['Bioquimica']['equipamentos']))
        l4.Interior.Color = tema.CAMPO
        l4.Font.Bold = True
        l4.Locked = False
        for b in (7, 8, 9, 10):
            l4.Borders(b).LineStyle = 1
            l4.Borders(b).Color = tema.CAMPO_BORDA
        log('  Painel: seletor de equipamento em L4 (selEquipamento)')


# =============================================================================
#  4. Estatistica: bloco de nao conformes / repeticoes
# =============================================================================
def instalar_bloco_estatistica(wb, produto):
    bio = produto.startswith('Bio')
    est = sheet(wb, 'Estatística')
    if not bio:
        # mesmos nomes da Bioquimica para a formula ser identica (ano De/Ate da Hematologia).
        # Nome com FUNCAO e recusado pelo Names.Add neste Excel; a conta fica em duas celulas
        # da propria aba (AF3/AF4, fora da area usada) e o nome aponta para elas.
        est.Range('AE3').Value = 'Início efetivo (Ano De)'
        est.Range('AE4').Value = 'Fim efetivo (Ano Até)'
        est.Range('AF3').Formula = '=IF(N($B$3)=0,0,DATE($B$3,1,1))'
        est.Range('AF4').Formula = '=IF(N($B$3)=0,0,DATE(MAX($B$3,$D$3),12,31))'
        est.Range('AF3:AF4').NumberFormat = 'dd/mm/yyyy'
        est.Range('AE3:AF4').Font.Size = tema.NOTA
        est.Range('AE3:AF4').Font.Color = tema.TEXTO_FRACO
        nome(wb, 'Estat_Lote', "='Estatística'!$B$4")
        nome(wb, 'Estat_Ini_Efetiva', "='Estatística'!$AF$3")
        nome(wb, 'Estat_Fim_Efetiva', "='Estatística'!$AF$4")
    r0 = 142
    est.Range(f'A{r0}:C{r0 + 45}').Clear()
    est.Range(f'A{r0}').Value = 'RESULTADOS NÃO CONFORMES / REPETIÇÕES'
    est.Range(f'A{r0}').Font.Bold = True
    est.Range(f'A{r0}').Font.Size = tema.SECAO
    est.Range(f'A{r0}').Font.Color = tema.TINTA
    est.Range(f'A{r0 + 1}').Value = ('Quantidade de resultados de controle retirados da população estatística principal e '
                                     'tratados como repetição/não conformidade (inativados na aba Inativar), no lote e período '
                                     'desta aba' + (' e no equipamento do Painel' if bio else '') + '. Não é contagem de erro '
                                     'laboratorial nem de resultado de paciente.')
    est.Range(f'A{r0 + 1}').Font.Size = tema.NOTA
    est.Range(f'A{r0 + 1}').Font.Color = tema.TEXTO_FRACO
    est.Range(f'A{r0 + 2}').Value = 'ANALITO'
    est.Range(f'B{r0 + 2}').Value = 'QUANTIDADE DE REPETIÇÕES'
    h = est.Range(f'A{r0 + 2}:B{r0 + 2}')
    h.Font.Bold, h.Font.Size, h.Font.Color, h.Interior.Color = True, tema.MIUDO, tema.SOBRE_MARCA, tema.MARCA
    h.WrapText = True
    # UMA formula (matriz dinamica): os analitos na ordem do cadastro e a contagem
    # de cada um. Le so a classificacao que o Power Query ja fez (STATUS_ANALITICO).
    # Periodo POR DIA nas duas pontas (INT), como o motor: inicio com hora (ex. 03:00)
    # nao pode tirar do bloco os resultados do proprio primeiro dia.
    f = ('=LET(a,FILTER(Analitos!$A$4:$A$43,Analitos!$A$4:$A$43<>""),'
         'lt,Estat_Lote,ini,INT(N(Estat_Ini_Efetiva)),fim,N(Estat_Fim_Efetiva),'
         'q,COUNTIFS(tblCQ_Final[ANALITO],a,tblCQ_Final[STATUS_ANALITICO],"INATIVADO",'
         'tblCQ_Final[EQUIPAMENTO],IF(selEquipamento="","<>",selEquipamento),'
         'tblCQ_Final[LOTE],IF(lt="","<>",lt),'
         'tblCQ_Final[DATA],">="&ini,tblCQ_Final[DATA],"<"&IF(fim=0,2958466,INT(fim)+1)),'
         'HSTACK(a,q))')
    est.Range(f'A{r0 + 3}').Formula2 = f
    est.Range(f'B{r0 + 3}:B{r0 + 45}').HorizontalAlignment = -4108
    try:
        nome(wb, 'estNaoConformes', f"='Estatística'!$A${r0 + 3}#")
    except Exception:
        pass
    v = est.Range(f'A{r0 + 3}').Value
    log(f'  Estatística: bloco em A{r0}; primeira linha = {v!r}')


# =============================================================================
#  5. Registros: so calibracao
# =============================================================================
def instalar_registros(wb):
    reg = sheet(wb, 'Registros')
    cab = [reg.Cells(3, c).Value for c in range(1, 14)]
    if cab[5:8] == ['Rep 1', 'Rep 2', 'Rep 3']:
        # RegistrosStore guarda regView (B:M) em A:L; Rep 1-3 = colunas E:G da store
        store = sheet(wb, 'RegistrosStore')
        hs = [store.Cells(1, c).Value for c in range(1, 13)]
        store.Range('E:G').Delete()
        reg.Range('F:H').Delete()
        log(f'  Registros: Rep 1-3 removidas (store antes: {hs[3:8]})')
    else:
        log(f'  Registros: já sem Rep 1-3 ({cab[4:9]})')
    rv = wb.Names('regView').RefersTo
    if '$J$203' not in rv:
        raise SystemExit(f'regView inesperado depois da remocao: {rv}')
    reg.Range('A1').Value = 'REGISTROS / OBSERVAÇÕES — CALIBRAÇÃO'
    reg.Range('A2').Formula = ('="REGISTROS DO LOTE EM USO:  "&loteAtivo&"   ·   \'Foi calibrado? = Sim\' plota o ícone de '
                               'calibração na média, no dia da calibração. Repetições/não conformes agora são inativados '
                               'por ID (aba Inativar) e aparecem como X vermelho."')
    for n in ('regRep1', 'regRep2', 'regRep3'):
        try:
            wb.Names(n).Delete()
        except Exception:
            pass


# =============================================================================
#  6. legado
# =============================================================================
def referencias_a(wb, abas):
    """Formulas (fora das proprias abas) que citam alguma das abas -- prova de dependencia."""
    achados = []
    # nome EXATO da aba: 'Resultados'! ou Resultados! sem nada colado antes -- senao
    # 'Principal - Resultados'! (aba nova) seria confundida com Resultados (legado)
    pat = re.compile('|'.join("'" + re.escape(a) + "'!|" + r"(?<![\w'. -])" + re.escape(a) + "!" for a in abas))
    for ws in wb.Worksheets:
        if ws.Name in abas:
            continue
        try:
            ur = ws.UsedRange
            fs = ur.Formula
        except Exception:
            continue
        if not isinstance(fs, tuple):
            fs = ((fs,),)
        for i, linha in enumerate(fs):
            for j, f in enumerate(linha):
                if isinstance(f, str) and f.startswith('=') and pat.search(f):
                    achados.append(f'{ws.Name}!{col_letra(ur.Column + j)}{ur.Row + i}: {f[:100]}')
    for n in wb.Names:
        try:
            rt = n.RefersTo
        except Exception:
            continue
        if pat.search(rt or ''):
            achados.append(f'nome {n.Name}: {rt[:100]}')
    return achados


def remover_legado(wb, produto):
    ini = sheet(wb, 'Início')
    # Inicio: "Ultima corrida" e "Resultados no banco" passam a ler a fonte oficial
    for r in range(20, 32):
        rot = str(ini.Cells(r, 2).Value or '')
        if rot.startswith('Última corrida'):
            ini.Cells(r, 2).Value = 'Último resultado recebido'
            ini.Cells(r, 3).Formula = '=IF(COUNT(tblCQ_Final[DATA_HORA])=0,"nenhum",TEXT(MAX(tblCQ_Final[DATA_HORA]),"dd/mm/aaaa hh:mm"))'
        elif rot.startswith('Resultados no banco'):
            ini.Cells(r, 2).Value = 'Resultados na DB_CQ_FINAL'
            ini.Cells(r, 3).Formula = ('=TEXT(ROWS(tblCQ_Final),"#.##0")&" · na estatística: "&'
                                       'TEXT(COUNTIFS(tblCQ_Final[PARTICIPA_ESTATISTICA],"SIM"),"#.##0")&'
                                       '" · inativados: "&COUNTIFS(tblCQ_Final[STATUS_ANALITICO],"INATIVADO")')
        elif rot.startswith('Lote em uso (lançamentos)'):
            ini.Cells(r, 2).Value = 'Lote em uso'
        txt = str(ini.Cells(r, 2).Value or '')
        if txt.startswith('2.  Todo dia'):
            ini.Cells(r, 2).Value = ('2.  Todo dia:  "⟳ ATUALIZAR DADOS" (Painel ou Dados do CQ). Os resultados chegam do '
                                     'DB_SEAC; resultado fora do padrão vai para "Inativar". Se o interfaceamento falhar, '
                                     'lance em "Resultados manuais".')
    for r in range(30, 45):
        txt = str(ini.Cells(r, 2).Value or '')
        if txt.startswith('2.  Todo dia'):
            ini.Cells(r, 2).Value = ('2.  Todo dia:  "⟳ ATUALIZAR DADOS" (Painel ou Dados do CQ). Os resultados chegam do '
                                     'DB_SEAC; resultado fora do padrão vai para "Inativar". Se o interfaceamento falhar, '
                                     'lance em "Resultados manuais".')
    leg = sheet(wb, 'Audit_Legenda')
    if leg is not None:
        for r in range(1, 32):
            v = str(leg.Cells(r, 2).Value or '')
            if 'Cfg_Status' in v:
                leg.Cells(r, 2).Value = v.replace('Cfg_Status', 'tblCQ_Final (PARTICIPA_ESTATISTICA, decidido no Power Query)')
    presentes = [a for a in ABAS_LEGADO if sheet(wb, a) is not None]
    deps = referencias_a(wb, presentes)
    # nomes r* apontam para DB_Resultados por definicao: saem junto
    deps = [d for d in deps if not any(d.startswith(f'nome {n}:') for n in NOMES_LEGADO)]
    if deps:
        raise SystemExit('DEPENDENCIA de aba legada -- nada removido:\n  ' + '\n  '.join(deps[:40]))
    for n in NOMES_LEGADO:
        try:
            wb.Names(n).Delete()
        except Exception:
            pass
    # shapes que chamam macros que sairam
    sairam = ('IrResultados', 'AbrirFormCorrida', 'AbrirFormMassa', 'AbrirFormExcluir', 'RegistrarImportacao',
              'IrParaImportar', 'AtualizarOperacao', 'ConferirFlagsBanco', 'TestarCapacidade')
    for ws in wb.Worksheets:
        if ws.Name in presentes:
            continue
        for s in list(ws.Shapes):
            try:
                oa = s.OnAction or ''
            except Exception:
                oa = ''
            if any(oa.endswith(m) for m in sairam):
                log(f'  shape {ws.Name}!{s.Name} chamava {oa} -> removido')
                s.Delete()
    wb.Application.DisplayAlerts = False
    for a in presentes:
        ws = sheet(wb, a)
        ws.Visible = -1
        ws.Delete()
        log(f'  aba removida: {a}')
    return presentes


# =============================================================================
#  7. navegacao
# =============================================================================
def instalar_navegacao(wb, produto):
    """Um botao por aba de dados, com o MESMO nome da aba (pedido do usuario):
    DB, Inativar, Digitar Resultados, Principal - Resultados. 'Principal - Resultados'
    entra na barra de TODAS as telas no lugar do antigo '+ Lancar'; nas abas de dados
    (e no Painel) os outros ficam ao lado, junto com o ATUALIZAR DADOS."""
    import ux
    import painel
    ux.NAV[:] = [('principal', 'IrPrincipal', 'Principal - Resultados') if k == 'lancar' else (k, m, r)
                 for k, m, r in ux.NAV]
    for k in ('Importar', 'Resultados'):
        ux.ATIVO.pop(k, None)
    ux.ATIVO['Principal - Resultados'] = 'principal'

    def botao(nome_shape, macro, texto):
        return (nome_shape, macro, texto, None)      # largura medida na hora (ver barra abaixo)
    atualizar = botao('nav_atualizar', 'AtualizarDados', '⟳ ATUALIZAR DADOS')
    dados = [botao('nav_db', 'IrDB', 'DB'), botao('nav_inativar', 'IrInativar', 'Inativar'),
             botao('nav_digitar', 'IrDigitarResultados', 'Digitar Resultados')]
    extra = {'Inativar': [botao('nav_comentarios', 'IrComentarios', 'COMENTARIOS_TECNICOS')],
             'Principal - Resultados': [botao('nav_qa', 'IrQA', 'QA_INTEGRACAO')]}
    for a in ('DB', 'DB_ORGANIZADO', 'Digitar Resultados', 'Inativar', 'COMENTARIOS_TECNICOS',
              'Principal - Resultados', 'QA_INTEGRACAO', 'Cfg_Integracao'):
        ux.EXTRAS[a] = [atualizar] + dados + extra.get(a, [])
    ux.EXTRAS['Painel'] = list(ux.EXTRAS.get('Painel', [])) + [atualizar, botao('nav_inativar', 'IrInativar', 'Inativar')]

    original = ux.barra_navegacao

    def barra(ws, extras=(), inicio=None):
        """ux.barra_navegacao com largura MEDIDA por botao: 'Principal - Resultados'
        nao cabe nos 66 pt fixos e o nome do botao tem de ser o nome da aba, inteiro."""
        f = ws.Range('A1').Font

        def larg(t, base):
            return max(base, ux.larg_texto(ws, t, tema.FONTE, 8, True) + 16)
        extras = [(n, m, t, w if w else larg(t, 50)) for n, m, t, w in extras]
        for s in list(ws.Shapes):
            if s.Name.startswith('nav_'):
                s.Delete()
        a1 = ws.Range('A1').MergeArea
        topo = ws.Range('A1').Top
        alt = a1.Height if a1.Height >= 24 else ws.Range('A1:A2').Height
        h, gap = ux.NAV_H, ux.NAV_GAP
        larguras = [larg(rot, ux.NAV_W) for _k, _m, rot in ux.NAV]
        total = sum(w + gap for w in larguras) + sum(e[3] + 8 for e in extras)
        direita = a1.Left + a1.Width
        larg_titulo = ux.larg_texto(ws, str(ws.Range('A1').Value or ''), str(f.Name), float(f.Size or 11),
                                    bool(f.Bold)) + 18
        if inicio:
            left0 = max(ws.Range(inicio).Left, a1.Left + larg_titulo)
        else:
            left0 = max(a1.Left + larg_titulo, direita - total - 6)
        top = topo + max(1, (alt - h) / 2)
        ativo = ux.ATIVO.get(ws.Name, '')
        x = left0
        for nome_shape, macro, texto, w in extras:
            eh = texto == ws.Name                       # o botao da propria aba fica destacado
            ux._botao(ws, nome_shape, macro, texto, x, top, w, h, ux.BTN_ATIVO if eh else ux.rgb(31, 111, 60),
                      ux.AZUL_ESCURO if eh else ux.BRANCO, tam=8)
            x += w + 8
        for (k, macro, rot), w in zip(ux.NAV, larguras):
            eh = (k == ativo)
            ux._botao(ws, f'nav_{k}', macro, rot, x, top, w, h, ux.BTN_ATIVO if eh else ux.BTN,
                      ux.AZUL_ESCURO if eh else ux.BRANCO, tam=8)
            x += w + gap
        fim_barra = x + 8
        if fim_barra > direita:
            c0 = a1.Column + a1.Columns.Count
            c = c0
            while ws.Cells(1, c).Left < fim_barra and c < c0 + 60:
                c += 1
            cor = ws.Range('A1').Interior.Color
            linhas = 2 if ws.Range('A2').Interior.Color == cor else 1
            linhas = max(linhas, a1.Rows.Count)
            ws.Range(ws.Cells(1, c0), ws.Cells(linhas, c)).Interior.Color = cor
    ux.barra_navegacao = barra
    try:
        pai = sheet(wb, 'Painel')
        col_nav = painel.coluna_da_barra(pai)
        ux.barras(wb, col_nav + '1')
    finally:
        ux.barra_navegacao = original
    ini = sheet(wb, 'Início')
    for s in list(ini.Shapes):
        if s.Name in ('tile_lancar', 'tile_principal'):
            s.OnAction = 'IrPrincipal'
            sub = 'resultados prontos · DB · Inativar · Digitar Resultados'
            s.TextFrame.Characters().Text = 'Principal - Resultados\n' + sub
            try:
                ch = s.TextFrame.Characters(len('Principal - Resultados') + 2, len(sub))
                ch.Font.Size, ch.Font.Bold = 8, False
            except Exception:
                pass
            s.Name = 'tile_principal'
        elif s.Name == 'tile_reg':
            s.TextFrame.Characters().Text = 'Calibração\nmarcada no gráfico (ícone na média)'
            try:
                ch = s.TextFrame.Characters(len('Calibração') + 2, 60)
                ch.Font.Size, ch.Font.Bold = 8, False
            except Exception:
                pass
    log('  navegação: botões DB · Inativar · Digitar Resultados · Principal - Resultados (mesmo nome das abas)')


def compilar_vba(ex, wb, teto=120):
    """Depurar > Compilar VBAProject, com cao de guarda: erro de compilacao abre um
    modal no VBE (que trava a automacao e aparece na tela do usuario) -- aqui ele
    derruba a instalacao em vez de passar adiante."""
    import threading
    ctl = wb.VBProject.VBE.CommandBars.FindControl(1, 578)
    if ctl is None:
        log('  (comando Compilar nao encontrado; compilacao sob demanda)')
        return None
    feito = threading.Event()
    morto = []

    def cao():
        if not feito.wait(teto):
            morto.append(True)
            ex.matar()
    threading.Thread(target=cao, daemon=True).start()
    try:
        if ctl.Enabled:
            ctl.Execute()
    except Exception:
        if morto:
            raise SystemExit('ERRO DE COMPILACAO DO VBA (modal no VBE) -- instalacao interrompida')
        raise
    finally:
        feito.set()
    if morto:
        raise SystemExit('ERRO DE COMPILACAO DO VBA (modal no VBE) -- instalacao interrompida')
    ok = not ctl.Enabled
    log(f'  VBA compilado: {"OK" if ok else "comando ainda habilitado (conferir)"}')
    return ok


# =============================================================================
#  9. X vermelho nos graficos (XML, com o Excel fechado -- ADR-055 §4)
# =============================================================================
def x_vermelho_nos_graficos(caminho):
    tmp = caminho + '.tmp'
    trocados = 0
    with zipfile.ZipFile(caminho) as zin, zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            dados = zin.read(item.filename)
            if item.filename.startswith('xl/charts/chart') and item.filename.endswith('.xml'):
                x = dados.decode('utf-8')
                tam_ok = None
                for s in re.findall(r'<c:ser>.*?</c:ser>', x, re.S):
                    if '<c:v>OK</c:v>' in s:
                        m = re.search(r'<c:marker>.*?<c:size val="(\d+)"/>', s, re.S)
                        tam_ok = m.group(1) if m else None

                def troca_ser(m):
                    nonlocal trocados
                    s = m.group(0)
                    if '<c:v>Repetição</c:v>' not in s and f'<c:v>{NOME_X}</c:v>' not in s:
                        return s
                    s = s.replace('<c:v>Repetição</c:v>', f'<c:v>{NOME_X}</c:v>')
                    s = s.replace('<a:srgbClr val="7030A0"/>', '<a:srgbClr val="FF0000"/>')
                    if tam_ok:
                        s = re.sub(r'(<c:marker><c:symbol val="x"/><c:size val=")\d+("/>)', r'\g<1>' + tam_ok + r'\2', s)
                    trocados += 1
                    return s
                x = re.sub(r'<c:ser>.*?</c:ser>', troca_ser, x, flags=re.S)
                dados = x.encode('utf-8')
            zout.writestr(item, dados)
    os.replace(tmp, caminho)
    return trocados


# =============================================================================
def main(produto, caminho, salvar=True):
    caminho = os.path.abspath(caminho)
    ex = xlh.Excel()
    log(f'EXCEL_PID {ex.pid}')
    try:
        wb = ex.abrir(caminho)
        ex.xl.Calculation = -4135
        desproteger_tudo(wb)
        log('1. camada de dados')
        tempos = cd.montar(wb, produto, carregar=True)
        log('2. VBA')
        instalar_vba(wb, produto)
        compilar_vba(ex, wb)
        log('3. Calc / nomes / Painel')
        instalar_calc(wb, produto)
        log('5. Registros')
        instalar_registros(wb)
        log('6. legado')
        removidas = remover_legado(wb, produto)
        log('4. Estatística')
        ex.xl.Calculation = -4105
        instalar_bloco_estatistica(wb, produto)
        log('7. navegação')
        instalar_navegacao(wb, produto)
        log('8. atualização completa (mIntegracao.AtualizarDados via VBA)')
        ex.xl.EnableEvents = True
        # o MESMO caminho do botao, sem a caixa de dialogo final (AtualizarDadosCore)
        t0 = time.time()
        if os.environ.get('QC_PARAR_ANTES_ETAPA8'):
            wb.SaveCopyAs(caminho.replace('.xlsm', '_pre8.xlsm'))
            log('  copia antes da etapa 8 salva (QC_PARAR_ANTES_ETAPA8)')
            raise SystemExit(0)
        try:
            resumo = str(ex.run("'" + wb.Name + "'!mIntegracao.AtualizarDadosAutomatico", teto=1500))
            if not resumo.startswith('OK|'):
                raise RuntimeError('ATUALIZAR DADOS falhou: ' + resumo)
            resumo = resumo[3:]
        except Exception:
            wb.SaveCopyAs(caminho.replace('.xlsm', '_falha8.xlsm'))
            log('  ETAPA 8 FALHOU -- copia para depuracao: ' + caminho.replace('.xlsm', '_falha8.xlsm'))
            raise
        log(f'  ATUALIZAR DADOS ({time.time() - t0:.1f}s):\n    ' + str(resumo).replace('\n', '\n    '))
        ex.xl.EnableEvents = False
        # protecao final (o mesmo que ReprotectAll faz no login)
        for ws in wb.Worksheets:
            try:
                ws.Protect(SENHA, False, True, False, True)    # DrawingObjects:=False, Contents:=True, Scenarios:=False, UserInterfaceOnly:=True
            except Exception:
                pass
        if salvar:
            wb.Save()
            log('salvo')
    finally:
        ex.fechar()
    if salvar:
        n = x_vermelho_nos_graficos(caminho)
        log(f'9. gráficos: {n} série(s) de X agora vermelhas e com o nome "{NOME_X}"')
    return tempos


if __name__ == '__main__':
    main(sys.argv[1], sys.argv[2], salvar='--sem-salvar' not in sys.argv)
