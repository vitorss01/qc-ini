# -*- coding: utf-8 -*-
"""QA do ADR-095 (14/11/2026, reteste 04): listas de escolha com o separador do Excel, acentos do DETALHE, cabecalho da
Estatistica da Bio em "Tudo" e o aviso do corte de 180 corridas na Hema. Roda numa COPIA instalada; nada e salvo.

Uso: python qa_adr095.py <Bioquimica|Hematologia> <copia_instalada.xlsm> <saida.json>

  L01 (R4-01, Alto) Digitar Resultados: EQUIPAMENTO, MATRIZ e MOTIVO -- a lista tem UM item por valor do cadastro
      (config_produtos / MOTIVOS_MANUAL); cada item escrito na celula e aceito (Validation.Value) e o texto com todos
      os itens juntos por virgula e recusado
  L02 (R4-01) Registros H4 "Calibrado?": Sim e Não aceitos, "Sim,Não" recusado; Bio: Painel L4 aceita cada equipamento
  L03 (R4-01) nenhuma validacao de lista do arquivo gravada como item unico com virgula (x12ac:list no XML) -- a trava
      contra o defeito voltar por outro instalador
  L04 nenhum formato de data com o ano literal (barra-y no XML: "14/11/yyyy" na tela) -- instalar_adr095
  D01 (OC-0021) DETALHE: a formula da coluna = camada_dados.FORMULA_DETALHE; os textos literais dos dois motores
      (mEtlCQ.bas e DB_CQ_FINAL.m) depois dos acentos nao tem nenhuma palavra de SEM_ACENTO_PROIBIDO; a trava reprova
      um texto sintetico sem acento
  E01 (Bio, P-09) Estatistica com o Painel sem De/Ate ("Tudo"): D4, B5 e D5 vazios (sem "Ano 2026" e sem
      "00/01/1900"); com De/Ate: B5/D5 = as datas e D4 = o ano do Painel
  E02 (Bio, reteste 05 R5-01) so o De / so o Ate: H5/J5 dizem "primeiro dado"/"ultimo dado" em vez de "00/01/1900"
  A01 (OC-0012, efeito colateral) Painel A39: Bio aponta a Estatistica so com lote no H3 (R5V-01: com todos os lotes
      diz o recorte); Hema nunca aponta (diz o recorte: todos os lotes / lote)
  P01 (reteste 05, P-03) linha colada com o ID de outra ganha ID novo; bloco colado: a de cima fica; mover nao muda
  Y00 seguranca do teste: arquivo de entrada intacto
"""
import json
import os
import re
import sys
import traceback
import zipfile

AQUI = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, AQUI)
from qa_final import QA, reg, RES, sha256_arquivo  # noqa: E402
sys.path.insert(0, os.path.join(AQUI, '..'))
sys.path.insert(0, os.path.join(AQUI, '..', '..', 'etapa1_multilote'))
import instalar_integracao as ii  # noqa: E402
import camada_dados as cd  # noqa: E402
import config_produtos as cp  # noqa: E402

RAIZ = os.path.join(AQUI, '..')


def txt(v):
    return '' if v is None else str(v).strip()


def listas_x12ac(caminho):
    """(aba, sqref, lista) de cada validacao gravada como item unico (x12ac:list) -- lido direto do XML."""
    z = zipfile.ZipFile(caminho)
    wbx = z.read('xl/workbook.xml').decode('utf8')
    rels = z.read('xl/_rels/workbook.xml.rels').decode('utf8')
    alvo = {re.search(r'Id="([^"]+)"', r).group(1): re.search(r'Target="([^"]+)"', r).group(1)
            for r in re.findall(r'<Relationship [^>]*/>', rels)}
    out = []
    for nome, rid in re.findall(r'<sheet [^>]*name="([^"]+)"[^>]*r:id="([^"]+)"', wbx):
        x = z.read('xl/' + alvo[rid].replace('/xl/', '')).decode('utf8')
        for dv in re.findall(r'<dataValidation [^>]*>.*?</dataValidation>', x, re.S):
            m = re.search(r'<x12ac:list>([^<]*)</x12ac:list>', dv)
            if m:
                out.append((nome, re.search(r'sqref="([^"]+)"', dv).group(1)[:40], m.group(1)[:120]))
    return out


def literais_motores():
    """Textos literais do codigo (sem comentarios) dos dois motores que gravam o MOTIVO_EXCLUSAO_AUTOMATICA."""
    lits = []
    for arq, com in ((os.path.join(RAIZ, 'src', 'comum', 'mEtlCQ.bas'), r"^\s*'.*$"),
                     (os.path.join(RAIZ, 'pq', 'DB_CQ_FINAL.m'), r'^\s*//.*$')):
        t = '\n'.join(re.sub(com, '', ln) for ln in open(arq, encoding='utf-8').read().splitlines())
        lits += [s.replace('""', '"') for s in re.findall(r'"((?:[^"]|"")*)"', t)]
    return lits


def sem_acento(textos):
    ruins = []
    for s in textos:
        x = ' ' + s + ' '
        for a, b in cd.ACENTOS_DETALHE:
            x = x.replace(a, b)
        for w in cd.SEM_ACENTO_PROIBIDO:
            if re.search(r'(?<![A-Za-zÀ-ÿ_])' + w + r'(?![A-Za-zÀ-ÿ_])', x):
                ruins.append(f'{w}: {s[:90]}')
        for pad in cd.SEM_ACENTO_PADROES:                      # reteste 05: o verbo "é" e o ordinal (padrao generico)
            if re.search(pad, x.strip()):
                ruins.append(f'{pad}: {s[:90]}')
    return ruins


def aceita(cel, valor):
    cel.Value = valor
    try:
        return bool(cel.Validation.Value)
    except Exception:                                   # noqa: BLE001
        return None


def executar(produto, caminho, saida):
    hash0 = sha256_arquivo(caminho)
    bio = produto.startswith('Bio')
    p_cfg = cp.PRODUTOS[produto]
    q = QA(produto, caminho)
    try:
        xl, wb = q.ex.xl, q.wb
        xl.EnableEvents = False
        sep = cd.sep_lista(xl)

        # ------------------------------------------------------------ L01
        lo = wb.Worksheets('Digitar Resultados').ListObjects('tblResultados_Manuais')
        ws = lo.Parent
        ws.Unprotect(ii.SENHA)
        matrizes = sorted({m for m, *_ in p_cfg['depara'] if m}) or [p_cfg['cfg']['MATRIZ_PADRAO'][0]]
        esperado = {'EQUIPAMENTO': list(p_cfg['equipamentos']), 'MATRIZ': matrizes, 'MOTIVO': list(cd.MOTIVOS_MANUAL)}
        ev, ok1 = {}, True
        for col, itens in esperado.items():
            rng = lo.ListColumns(col).DataBodyRange
            cel = rng.Cells(rng.Rows.Count, 1)              # a ultima linha (a das linhas novas) tambem tem a lista
            antes = cel.Formula
            f1 = txt(cel.Validation.Formula1)
            menu = f1.split(sep)
            ac = {i: aceita(cel, i) for i in itens}
            junto = aceita(cel, ','.join(itens)) if len(itens) > 1 else False
            cel.Formula = antes
            okc = menu == itens and all(v is True for v in ac.values()) and junto is False
            ok1 = ok1 and okc
            ev[col] = {'menu': menu, 'aceitos': ac, 'texto_junto_aceito': junto, 'ok': okc}
        reg('L01 Digitar Resultados: EQUIPAMENTO, MATRIZ e MOTIVO com um item por valor; cada item é aceito e o texto '
            'com os itens juntos por vírgula é recusado (R4-01)', ok1, ev)

        # ------------------------------------------------------------ L02
        r = wb.Worksheets('Registros')
        r.Unprotect(ii.SENHA)
        h4 = r.Range('H4')
        a = h4.Formula
        ev2 = {'H4_menu': txt(h4.Validation.Formula1).split(sep), 'Sim': aceita(h4, 'Sim'), 'Não': aceita(h4, 'Não'),
               'Sim,Não': aceita(h4, 'Sim,Não')}
        h4.Formula = a
        ok2 = ev2['Sim'] is True and ev2['Não'] is True and ev2['Sim,Não'] is False
        if bio:
            p = wb.Worksheets('Painel')
            p.Unprotect(ii.SENHA)
            l4 = p.Range('L4')
            a = l4.Formula
            ev2['L4'] = {e: aceita(l4, e) for e in p_cfg['equipamentos']}
            l4.Formula = a
            ok2 = ok2 and all(v is True for v in ev2['L4'].values())
            ev2['L4_msg'] = txt(l4.Validation.ErrorMessage)
            ok2 = ok2 and 'na seta da célula' in ev2['L4_msg'] and bool(l4.Validation.ShowError)
        # reteste 05 (P-02): os filtros e as colunas com lista recusam com a mensagem do sistema, nao a do Excel
        import instalar_adr087 as i87
        msgs = {f'{chr(64 + c)}4': txt(r.Cells(4, c).Validation.ErrorMessage) for c in i87.MSG_FILTRO}
        lo_r = r.ListObjects(i87.TB)
        for idx in i87.MSG_TABELA:
            try:
                msgs[f'tbl col {idx}'] = txt(lo_r.ListColumns(idx).DataBodyRange.Validation.ErrorMessage)
            except Exception:                           # noqa: BLE001  (coluna sem lista)
                msgs[f'tbl col {idx}'] = None
        ev2['mensagens'] = msgs
        ok2 = ok2 and all(m and 'na seta da célula' in m for k, m in msgs.items() if not k.startswith('tbl'))             and msgs.get('tbl col 6') and 'na seta da célula' in msgs['tbl col 6']
        reg('L02 Registros H4 "Calibrado?" aceita Sim e Não e recusa "Sim,Não"; Bio: Painel L4 aceita cada equipamento; '
            'filtros do Registros, coluna "Foi calibrado?" e L4 recusam com "na seta da célula" (reteste 05, P-02)',
            ok2, ev2)

        # ------------------------------------------------------------ L03 (XML da copia, sem Excel)
        x12 = listas_x12ac(caminho)
        reg('L03 Nenhuma validação de lista gravada como item único com vírgula (x12ac:list) em nenhuma aba', not x12,
            {'x12ac': x12})

        # ------------------------------------------------------------ L04 (XML da copia, sem Excel)
        import instalar_adr095 as i95
        lit = i95.ano_literal(caminho)
        reg('L04 Nenhuma célula ou coluna com o ano literal no formato de data ("14/11/yyyy": NumberFormat "yyyy" gravado '
            'por automação no Excel pt-BR), em nenhuma aba (ADR-095, revisão independente)', not lit,
            {a_: {'celulas': len(d['celulas']), 'colunas': len(d['colunas']), 'faixas': [f[:3] for f in i95.faixas(d['celulas'])[:6]]}
             for a_, d in lit.items()})

        # ------------------------------------------------------------ D01
        lo_m = lo
        fd = txt(lo_m.ListColumns('DETALHE').DataBodyRange.Cells(1, 1).Formula)
        faltam = [b for _, b in cd.ACENTOS_DETALHE if f'"{b}"' not in fd]
        ruins = sem_acento(literais_motores())
        trava = sem_acento(['fora do periodo historico (texto de teste)'])   # palavras sem par na lista: a trava pega
        # reteste 05: textos sinteticos SEM par na lista (verbo "é" depois de um ID e o ordinal) -- a trava tem de pegar
        # os dois; e a conjuncao no meio do texto ("fora da estatistica e do grafico") nao pode ser pega
        trava_ctx = sem_acento([' e do interfaceamento (texto de teste)', 'vale a 2a corrida (texto de teste)'])
        conj = sem_acento(['fora da estatística e do gráfico (texto de teste)'])
        reg('D01 DETALHE: a fórmula da coluna tem a lista de acentos do ADR-095; nenhum texto dos dois motores sobra sem '
            'acento (palavras e o verbo "é"/ordinal com contexto); a trava reprova um texto sintético sem acento (OC-0021)',
            not faltam and not ruins and len(trava) == 2 and len(trava_ctx) == 2 and not conj,
            {'acentos_fora_da_formula': faltam, 'sobram_sem_acento': ruins[:10], 'trava_sintetica': trava,
             'trava_verbo_ordinal': trava_ctx, 'conjuncao_pega': conj,
             'n_pares': len(cd.ACENTOS_DETALHE), 'formula_inicio': fd[:160]})

        # ------------------------------------------------------------ P01 (reteste 05, P-03)
        # Linha colada com o ID de outra ganha ID novo (Worksheet_Change -> mIntegracao.EntradaMudou, eventos LIGADOS
        # como na mao do usuario). So a coluna ID_REGISTRO e escrita; nada e salvo na copia.
        idc = lo.ListColumns('ID_REGISTRO').DataBodyRange
        vids = [txt(v[0]) for v in idc.Value]
        r0 = max([k for k, v in enumerate(vids, start=1) if v] or [0]) + 2      # duas linhas livres depois da ultima
        if r0 + 8 > len(vids):
            lo.Resize(lo.Range.Resize(lo.Range.Rows.Count + 10))
            idc = lo.ListColumns('ID_REGISTRO').DataBodyRange
        cel = lambda k: idc.Cells(k, 1)                                    # noqa: E731
        ev7 = {}
        xl.EnableEvents = True
        try:
            cel(r0).Value = 'MAN_9001'                                     # 1a: unica, fica
            cel(r0 + 1).Value = 'MAN_9001'                                 # colada com o ID de outra: ID novo
            bloco = ws.Range(cel(r0 + 2), cel(r0 + 3))
            bloco.Value = (('MAN_9100',), ('MAN_9100',))                   # bloco: a de cima fica, a de baixo muda
            ev7['antes_mover'] = [txt(cel(r0 + k).Value) for k in range(6)]
            cel(r0 + 2).Cut(cel(r0 + 5))                                   # mover (recortar e colar): nao muda
            ev7['depois_mover'] = [txt(cel(r0 + k).Value) for k in range(6)]
            # ID fora do padrao MAN_ repetido NAO e trocado: o motor continua a dar MAN_INVALIDO/E04 (regra intacta)
            cel(r0 + 6).Value = 'ABC9'
            cel(r0 + 7).Value = 'ABC9'
            ev7['invalido_repetido'] = [txt(cel(r0 + k).Value) for k in (6, 7)]
        finally:
            xl.EnableEvents = False
        a7, d7 = ev7['antes_mover'], ev7['depois_mover']
        novo = lambda s: s.startswith('MAN_') and s not in ('MAN_9001', 'MAN_9100')    # noqa: E731
        ok7 = (a7[0] == 'MAN_9001' and novo(a7[1]) and a7[2] == 'MAN_9100' and novo(a7[3]) and a7[3] != a7[1]
               and d7[2] == '' and d7[5] == 'MAN_9100' and ev7['invalido_repetido'] == ['ABC9', 'ABC9'])
        reg('P01 Digitar Resultados: linha colada com o ID de outra ganha ID novo (D05); num bloco colado a de cima fica '
            'com o ID; mover a linha (recortar e colar) não muda o ID; ID fora do padrão MAN_ repetido não é trocado (o motor '
            'o recusa como antes) (reteste 05, P-03)', ok7, ev7)

        # ------------------------------------------------------------ E01 (Bio)
        p = wb.Worksheets('Painel')
        p.Unprotect(ii.SENHA)
        if bio:
            e = wb.Worksheets('Estatística')
            g3, g4, k3 = p.Range('G3').Formula, p.Range('G4').Formula, p.Range('K3').Formula
            try:
                p.Range('G3').Value = None
                p.Range('G4').Value = None
                xl.Calculate()
                tudo = {c: txt(e.Range(c).Text) for c in ('B5', 'D5', 'D4', 'H5')}
                import datetime as dt
                p.Range('G3').Value = dt.datetime(2026, 9, 1)
                p.Range('G4').Value = dt.datetime(2026, 9, 30)
                p.Range('K3').Value = 2026
                xl.Calculate()
                setembro = {c: txt(e.Range(c).Text) for c in ('B5', 'D5', 'D4')}
                # E02 (R5-01): so o De / so o Ate -- o lado aberto da janela em texto, nunca "00/01/1900" nem "####"
                p.Range('G4').Value = None
                xl.Calculate()
                so_de = {c: txt(e.Range(c).Text) for c in ('H5', 'J5')}
                so_de['J5_v'] = e.Range('J5').Value2
                p.Range('G3').Value = None
                p.Range('G4').Value = dt.datetime(2026, 9, 30)
                xl.Calculate()
                so_ate = {c: txt(e.Range(c).Text) for c in ('H5', 'J5')}
                so_ate['H5_v'] = e.Range('H5').Value2
                f_h5, f_j5 = str(e.Range('H5').Formula), str(e.Range('J5').Formula)
            finally:
                p.Range('G3').Formula, p.Range('G4').Formula, p.Range('K3').Formula = g3, g4, k3
                xl.Calculate()
            ok5 = (tudo['B5'] == '' and tudo['D5'] == '' and tudo['D4'] == '' and tudo['H5'] == 'histórico inteiro do lote'
                   and setembro == {'B5': '01/09/2026', 'D5': '30/09/2026', 'D4': '2026'})
            reg('E01 Estatística com o Painel em "Tudo": B5, D5 e D4 vazios (sem "00/01/1900" nem "Ano 2026"); com De/Até '
                'set/2026: as datas e o ano (P-09)', ok5, {'tudo': tudo, 'set_2026': setembro})
            import instalar_adr092 as a92
            ok5b = (so_de['H5'] == '01/09/2026' and so_de['J5'] == 'último dado' and not so_de['J5_v']
                    and so_ate['H5'] == 'primeiro dado' and so_ate['J5'] == '30/09/2026' and not so_ate['H5_v']
                    and ii._norm_vba(f_h5) == ii._norm_vba(a92.H5) and ii._norm_vba(f_j5) == ii._norm_vba(a92.J5))
            reg('E02 Estatística H5/J5 com só o De / só o Até: "01/09/2026 a último dado" e "primeiro dado a 30/09/2026" '
                '(sem "00/01/1900" nem "####"; valor 0 e fórmulas do ADR-092 intactos) (reteste 05, R5-01)', ok5b,
                {'so_de': so_de, 'so_ate': so_ate, 'H5': f_h5, 'J5': f_j5})

        # ------------------------------------------------------------ A01
        a39 = str(p.Range('A39').Formula)
        if bio:
            # ADR-096 (R5V-01): com lote no H3 aponta a Estatistica; com todos os lotes diz o recorte (a Estatistica
            # segue o lote em analise e mostraria outro n)
            ok6 = ('veja a aba Estatística' in a39 and '": "&Eng_Saida!$AF$1&" — veja a aba Estatística. "' in a39
                   and 'IF(engTodos," (todos os lotes): "&Eng_Saida!$AF$1&" resultados. "' in a39)
        else:
            ok6 = ('Estatística' not in a39 and '" (todos os lotes)"," do lote")' in a39
                   and '&Eng_Saida!$AF$1&" resultados. ' in a39)
        reg('A01 Painel A39 (corte de 180 corridas): Bio aponta a aba Estatística; Hema não aponta e diz o recorte '
            '(todos os lotes ou o lote) -- efeito colateral da OC-0012', ok6, {'A39_fim': a39[-420:]})
    except Exception as e:                                     # noqa: BLE001
        reg('F99 Execução interrompida: os testes seguintes a este ponto não rodaram', False,
            {'erro': f'{type(e).__name__}: {e}', 'traceback': traceback.format_exc()[-1500:]})
    finally:
        q.fechar(salvar=False)
    hash1 = sha256_arquivo(caminho)
    reg('Y00 Segurança do teste: arquivo de entrada intacto (SHA-256 igual antes e depois; nada salvo nele)', hash0 == hash1,
        {'arquivo': caminho, 'sha256_antes': hash0, 'sha256_depois': hash1})
    if saida:
        with open(saida, 'w', encoding='utf-8') as f:
            json.dump(RES, f, ensure_ascii=False, indent=1, default=str)
    falhas = [r for r in RES if r['resultado'] == 'FAIL']
    print(f'\n=== {produto} ADR-095 (listas, acentos, Estatística, aviso do corte): {len(RES) - len(falhas)} PASS / '
          f'{len(falhas)} FAIL ===', flush=True)
    return falhas


if __name__ == '__main__':
    falhas = executar(sys.argv[1], os.path.abspath(sys.argv[2]), sys.argv[3] if len(sys.argv) > 3 else None)
    sys.exit(1 if falhas else 0)
