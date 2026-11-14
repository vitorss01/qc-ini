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
  D01 (OC-0021) DETALHE: a formula da coluna = camada_dados.FORMULA_DETALHE; os textos literais dos dois motores
      (mEtlCQ.bas e DB_CQ_FINAL.m) depois dos acentos nao tem nenhuma palavra de SEM_ACENTO_PROIBIDO; a trava reprova
      um texto sintetico sem acento
  E01 (Bio, P-09) Estatistica com o Painel sem De/Ate ("Tudo"): D4, B5 e D5 vazios (sem "Ano 2026" e sem
      "00/01/1900"); com De/Ate: B5/D5 = as datas e D4 = o ano do Painel
  A01 (OC-0012, efeito colateral) Painel A39: Bio aponta a Estatistica; Hema nao (diz o recorte: todos os lotes / lote)
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
        reg('L02 Registros H4 "Calibrado?" aceita Sim e Não e recusa "Sim,Não"; Bio: Painel L4 aceita cada equipamento',
            ok2, ev2)

        # ------------------------------------------------------------ L03 (XML da copia, sem Excel)
        x12 = listas_x12ac(caminho)
        reg('L03 Nenhuma validação de lista gravada como item único com vírgula (x12ac:list) em nenhuma aba', not x12,
            {'x12ac': x12})

        # ------------------------------------------------------------ D01
        lo_m = lo
        fd = txt(lo_m.ListColumns('DETALHE').DataBodyRange.Cells(1, 1).Formula)
        faltam = [b for _, b in cd.ACENTOS_DETALHE if f'"{b}"' not in fd]
        ruins = sem_acento(literais_motores())
        trava = sem_acento(['fora do periodo historico (texto de teste)'])   # palavras sem par na lista: a trava pega
        reg('D01 DETALHE: a fórmula da coluna tem a lista de acentos do ADR-095; nenhum texto dos dois motores sobra sem '
            'acento; a trava reprova um texto sintético sem acento (OC-0021)',
            not faltam and not ruins and len(trava) == 2,
            {'acentos_fora_da_formula': faltam, 'sobram_sem_acento': ruins[:10], 'trava_sintetica': trava,
             'n_pares': len(cd.ACENTOS_DETALHE), 'formula_inicio': fd[:160]})

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
            finally:
                p.Range('G3').Formula, p.Range('G4').Formula, p.Range('K3').Formula = g3, g4, k3
                xl.Calculate()
            ok5 = (tudo['B5'] == '' and tudo['D5'] == '' and tudo['D4'] == '' and tudo['H5'] == 'histórico inteiro do lote'
                   and setembro == {'B5': '01/09/2026', 'D5': '30/09/2026', 'D4': '2026'})
            reg('E01 Estatística com o Painel em "Tudo": B5, D5 e D4 vazios (sem "00/01/1900" nem "Ano 2026"); com De/Até '
                'set/2026: as datas e o ano (P-09)', ok5, {'tudo': tudo, 'set_2026': setembro})

        # ------------------------------------------------------------ A01
        a39 = str(p.Range('A39').Formula)
        if bio:
            ok6 = 'veja a aba Estatística' in a39 and 'Período inteiro: "&Eng_Saida!$AF$1' in a39
        else:
            ok6 = 'Estatística' not in a39 and '(todos os lotes)' in a39 and ' do lote' in a39 and 'resultados.' in a39
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
