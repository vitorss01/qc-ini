# -*- coding: utf-8 -*-
"""repro_listas_r4.py -- Tester/QA, ciclo 2026-11-14_reteste04, modo REPRODUZIR (P-01 Bio / P-02 Hema).

Escrito pelo Tester/QA, sem reaproveitar os scripts do Usuario Simulado. Usa so a infraestrutura oficial
(`qa_final.QA`: abre a copia, revela as abas, imprime EXCEL_PID; fecha so o Excel que criou).

O que faz, numa COPIA nova da producao (em %TEMP%/qcw/auditoria/2026-11-14_reteste04/tester/):
  1. le o separador de lista do Excel instalado (International(xlListSeparator));
  2. varre TODAS as validacoes de lista (xlValidateList) de todas as abas e marca as que tem lista literal com ","
     num Excel cujo separador nao e "," (lista de um item so);
  3. na tblResultados_Manuais, numa linha vazia e com eventos desligados, escreve cada item esperado e o texto
     inteiro da Formula1 e le Validation.Value (a avaliacao do proprio Excel), limpando a celula depois;
  4. (so Bio, --consequencia) digita dois manuais sinteticos com eventos ligados -- A com os itens do menu
     (EQUIPAMENTO/MATRIZ/MOTIVO = texto inteiro) e B controle com "DIMENSION 1"/"SORO"/"Outro" --, roda o ATUALIZAR e
     confere STATUS, tblCQ_Final, tblQA_Integracao e as duas series do Painel (L4 = DIMENSION 1 / DIMENSION 2).
Nada e salvo. Nenhuma captura de tela. Valores sinteticos: 123,4567 (A) e 123,4568 (B), MOTIVO/obs TESTE-AUDITORIA.

Uso: python -u repro_listas_r4.py Bioquimica|Hematologia [--consequencia]
"""
import datetime as dt
import hashlib
import json
import os
import shutil
import sys
import time

AQUI = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.abspath(os.path.join(AQUI, '..', '..', '..'))
TESTES = os.path.join(REPO, '_arquitetura', 'integracao_dbseac', 'testes')
sys.path.insert(0, TESTES)
sys.path.insert(0, os.path.join(REPO, '_arquitetura', 'integracao_dbseac'))
import qa_final  # noqa: E402

CICLO = '2026-11-14_reteste04'
PASTA = os.path.join(os.environ['TEMP'], 'qcw', 'auditoria', CICLO, 'tester')
SAIDA = os.path.join(REPO, '_arquitetura', 'agentes', 'relatorios', CICLO)

ESPERADO = {
    'Bioquimica': {'EQUIPAMENTO': ['DIMENSION 1', 'DIMENSION 2'], 'MATRIZ': ['SORO', 'URINA']},
    'Hematologia': {'EQUIPAMENTO': ['XN1000'], 'MATRIZ': ['SANGUE TOTAL']},
}
MOTIVOS = ['Falha do interfaceamento', 'Resultado não transmitido', 'Equipamento em contingência', 'Outro']
VAL_A, VAL_B = 123.4567, 123.4568


def sha(p):
    h = hashlib.sha256()
    with open(p, 'rb') as f:
        for b in iter(lambda: f.read(1 << 20), b''):
            h.update(b)
    return h.hexdigest()[:12]


R = {}


def reg(k, v):
    R[k] = v
    print('##', k, json.dumps(v, ensure_ascii=False, default=str)[:1500], flush=True)


def val(cel):
    try:
        v = cel.Validation
        return {'tipo': v.Type, 'formula1': v.Formula1, 'estilo': v.AlertStyle, 'msg': v.ErrorMessage,
                'mostra_erro': v.ShowError, 'menu': v.InCellDropdown}
    except Exception as e:
        return {'erro': str(e)[:120]}


def senha_planilha():
    """Senha de protecao das abas lida do instalador (nao repetida neste arquivo versionado)."""
    import re
    t = open(os.path.join(REPO, '_arquitetura', 'integracao_dbseac', 'camada_dados.py'), encoding='utf-8').read()
    return re.search(r"^SENHA\s*=\s*'([^']+)'", t, re.M).group(1)


def varrer_listas(wb, sep):
    achados, vistos, erros = [], set(), {}
    s = senha_planilha()
    for ws in wb.Worksheets:
        try:
            if ws.ProtectContents:
                ws.Unprotect(s)                      # so na copia, que nunca e salva (SpecialCells falha com a aba protegida)
            rng = ws.Cells.SpecialCells(-4174)       # xlCellTypeAllValidation
        except Exception as e:
            erros[ws.Name] = str(e)[:80]
            continue
        for a in rng.Areas:
            for j in range(1, a.Columns.Count + 1):
                c = a.Cells(1, j)
                v = val(c)
                if v.get('tipo') != 3:
                    continue
                f1 = str(v['formula1'] or '')
                chave = (ws.Name, f1, a.Columns(j).Address)
                if chave in vistos:
                    continue
                vistos.add(chave)
                literal = not f1.startswith('=')
                um_item = literal and (',' in f1) and sep != ','
                achados.append({'aba': ws.Name, 'intervalo': a.Columns(j).Address.replace('$', ''), 'formula1': f1,
                                'estilo': v['estilo'], 'msg': v['msg'],
                                'veredito': 'UM_ITEM_SO' if um_item else ('NOME/INTERVALO' if not literal else 'OK')})
    reg('varredura_abas_sem_validacao_ou_erro', erros)
    return achados


def testar_coluna(qa, lo, col, linha, candidatos):
    cel = lo.ListColumns(col).DataBodyRange.Cells(linha, 1)
    antes = cel.Value
    assert antes in (None, ''), f'celula {col} da linha {linha} nao esta vazia'
    out = {'celula': cel.Address.replace('$', ''), **val(cel), 'aceita': {}}
    for c in candidatos:
        cel.Value = c
        out['aceita'][str(c)] = bool(cel.Validation.Value)
        cel.ClearContents()
    return out


def linha_vazia(lo):
    vals = lo.DataBodyRange.Value
    for i, r in enumerate(vals, 1):
        if all(x in (None, '') for x in r[:10]):
            return i
    raise RuntimeError('sem linha vazia')


def main():
    produto = sys.argv[1]
    consequencia = '--consequencia' in sys.argv
    os.makedirs(PASTA, exist_ok=True)
    src = os.path.join(REPO, f'QC_{produto}.xlsm')
    dst = os.path.join(PASTA, f'{produto}_repro_listas_{dt.datetime.now():%H%M%S}.xlsm')
    reg('producao_sha256', sha(src))
    shutil.copy2(src, dst)
    reg('copia', {'arquivo': os.path.basename(dst), 'sha256': sha(dst)})
    qa = qa_final.QA(produto, dst)
    reg('EXCEL_PID', qa.ex.pid)
    try:
        xl = qa.ex.xl
        intl = xl.International          # por esta automacao vem a tupla inteira (indice 0 = xlCountryCode)
        intl = intl if isinstance(intl, (tuple, list)) else [xl.International(i) for i in range(1, 46)]
        sep = str(intl[5 - 1])
        reg('idioma', {'xlListSeparator': sep, 'xlDecimalSeparator': str(intl[3 - 1]),
                       'LanguageID_UI': xl.LanguageSettings.LanguageID(2)})
        reg('varredura_listas', varrer_listas(qa.wb, sep))

        lo = qa.lo('tblResultados_Manuais')
        cab = [c.Name for c in lo.ListColumns]
        reg('manuais_cabecalho', cab)
        reg('manuais_linhas_preenchidas_antes',
            sum(1 for r in lo.DataBodyRange.Value if any(x not in (None, '') for x in r[:10])))
        n = linha_vazia(lo)
        xl.EnableEvents = False
        try:
            lotes = [r[0] for r in qa.nome('lstLotes').Value if r[0] not in (None, '')]
            ans = [r[0] for r in qa.nome('lstAnalitos').Value if r[0] not in (None, '')]
            res = {}
            for col in ('EQUIPAMENTO', 'MATRIZ', 'MOTIVO'):
                esp = ESPERADO[produto].get(col, MOTIVOS)
                f1 = lo.ListColumns(col).DataBodyRange.Cells(n, 1).Validation.Formula1
                res[col] = testar_coluna(qa, lo, col, n, esp + [f1] if f1 not in esp else esp)
            res['LOTE'] = testar_coluna(qa, lo, 'LOTE', n, [lotes[0], 'LOTE-INEXISTENTE-TESTE'])
            res['ANALITO'] = testar_coluna(qa, lo, 'ANALITO', n, [ans[0], 'ANALITO-INEXISTENTE-TESTE'])
            reg('validacao_por_coluna', res)
        finally:
            xl.EnableEvents = True
        # o mesmo teste em outra linha (ultima da folga), para afastar efeito de linha
        n2 = lo.ListRows.Count
        xl.EnableEvents = False
        try:
            reg('validacao_ultima_linha', {c: testar_coluna(qa, lo, c, n2, (ESPERADO[produto].get(c, MOTIVOS))[:1])
                                           for c in ('EQUIPAMENTO', 'MATRIZ', 'MOTIVO')})
        finally:
            xl.EnableEvents = True

        if produto == 'Bioquimica':
            pai = qa.ws('Painel')
            l4 = pai.Range('L4')
            reg('painel_L4', {**val(l4), 'valor': l4.Value, 'aceita_valor_atual': bool(l4.Validation.Value)})

        if consequencia and produto == 'Bioquimica':
            consequencia_bio(qa, lo, lotes, ans)
    finally:
        json.dump(R, open(os.path.join(SAIDA, f'tester_repro_listas_{produto}.json'), 'w', encoding='utf-8'),
                  ensure_ascii=False, indent=1, default=str)
        qa.fechar(False)
        print('FECHADO (sem salvar) PID', R.get('EXCEL_PID'), flush=True)


def consequencia_bio(qa, lo, lotes, ans):
    xl = qa.ex.xl
    pai = qa.ws('Painel')
    gli = next((a for a in ans if str(a).upper().startswith('GLICOSE')), ans[0])
    pai.Range('L4').Value = 'DIMENSION 1'
    qa.ex.esperar()
    pai.Range('C3').Value = gli               # C3 = analito pelo nome (Painel.cls: AnalitoEscolhido); B3 e o indice
    qa.ex.esperar()
    time.sleep(1)
    lote = str(qa.nome('loteAnalise').Value)
    reg('painel_selecao', {'B3': pai.Range('B3').Value, 'C3': pai.Range('C3').Value, 'loteAnalise': lote,
                           'L4': pai.Range('L4').Value})
    f_eq = lo.ListColumns('EQUIPAMENTO').DataBodyRange.Cells(1, 1).Validation.Formula1
    f_mat = lo.ListColumns('MATRIZ').DataBodyRange.Cells(1, 1).Validation.Formula1
    f_mot = lo.ListColumns('MOTIVO').DataBodyRange.Cells(1, 1).Validation.Formula1
    hoje = dt.datetime(2026, 11, 14)
    linhas = {}
    for rot, eq, mat, mot, valor, hora in (('A_menu', f_eq, f_mat, f_mot, VAL_A, 3 / 24 + 33 / 1440),
                                            ('B_controle', 'DIMENSION 1', 'SORO', 'Outro', VAL_B, 3 / 24 + 34 / 1440)):
        n = linha_vazia(lo)
        ordem = [('DATA', hoje), ('HORA', hora), ('EQUIPAMENTO', eq), ('MATRIZ', mat), ('LOTE', lote), ('NIVEL', 1),
                 ('ANALITO', gli), ('RESULTADO', valor), ('MOTIVO', mot)]
        for col, v in ordem:              # celula a celula, eventos ligados (como o usuario)
            lo.ListColumns(col).DataBodyRange.Cells(n, 1).Value = v
            qa.ex.esperar()
        cels = {c.Name: lo.ListColumns(c.Name).DataBodyRange.Cells(n, 1) for c in lo.ListColumns}
        linhas[rot] = {'linha': n, 'valores': {k: c.Value for k, c in cels.items()},
                       'validacao_ok': {k: bool(cels[k].Validation.Value) for k in ('EQUIPAMENTO', 'MATRIZ', 'MOTIVO')}}
    reg('manuais_digitados', linhas)
    t0 = time.time()
    r, t = qa.atualizar(teto=1800)
    reg('atualizar', {'retorno': r[:300], 'segundos': round(t, 1)})
    pos = {}
    for rot, d in linhas.items():
        n = d['linha']
        pos[rot] = {c.Name: lo.ListColumns(c.Name).DataBodyRange.Cells(n, 1).Value for c in lo.ListColumns}
    reg('manuais_depois', pos)
    ids = {rot: str(v.get('ID_REGISTRO')) for rot, v in pos.items()}
    fin = [x for x in qa.ler('tblCQ_Final') if str(x.get('ID_REGISTRO')) in ids.values()]
    reg('cq_final', fin)
    qa_rows = [x for x in qa.ler('tblQA_Integracao')
               if any(i in json.dumps(x, ensure_ascii=False, default=str) for i in ids.values())]
    reg('qa_integracao_com_os_ids', qa_rows)
    reg('qa_integracao_total', len(qa.ler('tblQA_Integracao')))
    # series do Painel: procura os valores sinteticos na aba Calc com cada equipamento
    series = {'filtro_antes': {a: pai.Range(a).Value for a in ('G3', 'G4', 'K3', 'K4', 'H3')}}
    pai.Range('K4').Value = 'Tudo (sem filtro)'      # a 1a rodada mostrou o motor parado em 30/10 (filtro de periodo)
    qa.ex.esperar()
    series['filtro_depois'] = {a: pai.Range(a).Value for a in ('G3', 'G4', 'K3', 'K4', 'H3')}
    for eq in ('DIMENSION 1', 'DIMENSION 2'):
        pai.Range('L4').Value = eq
        qa.ex.esperar()
        xl.CalculateFull()
        qa.ex.esperar()
        achou = {}
        for aba in ('Eng_Saida', 'Calc'):
            sh = qa.ws(aba)
            ur = sh.UsedRange
            v2 = ur.Value2
            ach = {'A_menu': [], 'B_controle': []}
            for i, linha in enumerate(v2):
                for j, x in enumerate(linha):
                    if isinstance(x, float):
                        if abs(x - VAL_A) < 1e-9:
                            ach['A_menu'].append(sh.Cells(ur.Row + i, ur.Column + j).Address.replace('$', ''))
                        elif abs(x - VAL_B) < 1e-9:
                            ach['B_controle'].append(sh.Cells(ur.Row + i, ur.Column + j).Address.replace('$', ''))
            achou[aba] = ach
        es = qa.eng_saida()
        series[eq] = {'L4': pai.Range('L4').Value, 'C3': pai.Range('C3').Value, 'engAnalito': str(qa.nome('engAnalito').Value),
                      'engEquip': str(qa.nome('engEquip').Value), 'eng_n': es['n'],
                      'eng_ultimos_runs': es['runs'][-3:], 'eng_ultimos_val': es['val'][-3:], 'achados': achou}
    reg('series_painel', series)


if __name__ == '__main__':
    main()
