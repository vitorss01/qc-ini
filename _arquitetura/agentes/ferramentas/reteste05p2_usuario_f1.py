# -*- coding: utf-8 -*-
"""reteste05p2_usuario_f1.py -- Usuario Simulado, ciclo 2026-11-15_reteste05_parte2, frente F1, RETESTE (Bioquimica).

Ocorrencias (correcao declarada: ADR-095, entrega 42, Bio b20fe00f):
  OC-0021  DETALHE da Digitar Resultados sem acento (IMPLAUSIVEL, regras de lote ADR-091, troca de analito ADR-093,
           substituicao ADR-082). Fluxo: manuais digitados celula a celula (eventos ligados) que geram cada texto do
           MOTIVO_EXCLUSAO_AUTOMATICA possivel num manual; ATUALIZAR DADOS pelo nucleo do botao
           (mIntegracao.AtualizarDadosAutomatico, le o DB_SEAC so para leitura); leitura do DETALHE/RUN/STATUS (Text) e da
           tblCQ_Final. Oraculo proprio: lista de palavras sem acento (independente do camada_dados), verbo "e" -> "é",
           e DETALHE == MOTIVO com os acentos aplicados por mim.
  OC-0068  Listas de um item so: Digitar Resultados EQUIPAMENTO/MATRIZ/MOTIVO, Registros H4 "Calibrado?", Painel L4.
           Validation.Formula1 partida pelo separador do Excel; Validation.Value de cada item valido e de invalidos;
           ErrorMessage; filtro H4 aplicado (so muda a vista: n de linhas e valores da tblRegistros iguais); manual com
           o EQUIPAMENTO escolhido no item do menu "DIMENSION 2" -> entra na serie DIMENSION 2 do Painel (n +1, valor nos
           pontos do Eng_Saida) e a serie DIMENSION 1 nao muda.
  OC-0072  Ano literal nos formatos de data: Text/NumberFormat apos o login (Configuracao D/AD 26-125, Eng_Saida AC3:AC182,
           Registros E4/F4 com datas digitadas, Westgard_Fotos F, Sync_Log A/B/L/M, Estatistica E8:F9/C10, LotesStore U1)
           antes e depois do ATUALIZAR; SaveCopyAs da copia em uso -> varredura do XML (xml_formatos_listas.py).
A MESMA sequencia roda em duas copias, uma de cada vez (um Excel por vez):
  base41 = copia do backup anterior a entrega 42 (_backup_pre_integracao_2026-11-15_0011) -- o "antes";
  nova   = copia da producao atual (raiz).
Regras: so COPIAS em %TEMP%/qcw/auditoria/2026-11-15_reteste05_parte2/; a copia em uso nunca e salva (o SaveCopyAs grava
um terceiro arquivo na mesma pasta, so para ler o XML); fechar() encerra so o PID do QA; nenhuma captura de tela (caixas
lidas pelo texto das janelas do PID desta sessao e fechadas). Dados de teste: manuais com MOTIVO "Outro: TESTE-AUDITORIA",
valores sinteticos 123,456 / 111,111 / 12,345 / 1,234 / -1; usuario TESTE_AUDITORIA_R5P2 (ANALISTA) criado so na copia.
Uso: python -u reteste05p2_usuario_f1.py <base41|nova>  -> Bioquimica_F1_<rotulo>.json na pasta do ciclo
"""
import collections
import datetime as dt
import hashlib
import os
import re
import shutil
import sys
import time
import traceback
import unicodedata
import zipfile

AQUI = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, AQUI)
_argv = sys.argv[:]
sys.argv = [sys.argv[0], 'Bioquimica']
import valida_adr090_091 as V  # noqa: E402
sys.argv = _argv
import repro_comum as C  # noqa: E402
import xml_formatos_listas as X  # noqa: E402
from qa_inativar_ux import criar_usuario, fazer_login  # noqa: E402

PROD = 'Bioquimica'
CICLO = '2026-11-15_reteste05_parte2'
PASTA = os.path.join(os.environ['TEMP'], 'qcw', 'auditoria', CICLO)
C.PASTA = PASTA
BACKUP = os.path.join(C.REPO, '_backup_pre_integracao_2026-11-15_0011', f'QC_{PROD}.xlsm')
USR = 'TESTE_AUDITORIA_R5P2'
MOTIVO = 'Outro: TESTE-AUDITORIA'
MESES = ['Jan', 'Fev', 'Mar', 'Abr', 'Mai', 'Jun', 'Jul', 'Ago', 'Set', 'Out', 'Nov', 'Dez']
t, up, num = V.t, V.up, V.num

# ---- oraculo proprio de acentos (nao usa camada_dados): palavra sem acento -> com acento
MEUS_ACENTOS = {'nao': 'não', 'ja': 'já', 'so': 'só', 'estatistica': 'estatística', 'grafico': 'gráfico',
                'implausivel': 'implausível', 'lancado': 'lançado', 'lancamento': 'lançamento', 'catalogo': 'catálogo',
                'substituido': 'substituído', 'obrigatorio': 'obrigatório', 'obrigatorios': 'obrigatórios',
                'invalido': 'inválido', 'invalidos': 'inválidos', 'ambiguo': 'ambíguo', 'virgula': 'vírgula',
                'padrao': 'padrão', 'numerico': 'numérico', 'retransmissao': 'retransmissão', 'periodo': 'período',
                'codigo': 'código', 'numero': 'número', 'historico': 'histórico', 'analise': 'análise',
                'unico': 'único', 'proprio': 'próprio', 'automatico': 'automático', 'invalida': 'inválida',
                'conteudo': 'conteúdo', 'medio': 'médio', 'serie': 'série', 'apos': 'após', 'ate': 'até'}
RX_PAL = re.compile(r"[A-Za-zÀ-ÿ]+")


def sem_acento_achados(s):
    """Palavras sem acento (ignora tokens TODOS EM MAIUSCULAS = nomes de coluna/status) e o verbo 'e' sem acento."""
    achados = []
    for m in RX_PAL.finditer(s or ''):
        w = m.group(0)
        if w.isupper() and len(w) > 1:
            continue
        if w.lower() in MEUS_ACENTOS:
            achados.append(w)
    for m in re.finditer(r'\S+ e (resultado manual|retransmiss\w+)', s or ''):
        achados.append('verbo "e" -> "é": ' + m.group(0)[:60])
    if re.search(r'\b1a linha\b', s or ''):
        achados.append('"1a" (ordinal)')
    return achados


def acentuar(s):
    def f(m):
        w = m.group(0)
        if w.isupper() and len(w) > 1:
            return w
        return MEUS_ACENTOS.get(w.lower(), w) if w.islower() else w
    return RX_PAL.sub(f, s or '')


def tira_acento(s):
    return ''.join(c for c in unicodedata.normalize('NFD', s or '') if unicodedata.category(c) != 'Mn')


class Caixas(V.Vigia):
    """Como o V.Vigia, mas so caixas de mensagem (#32770): nao fecha o frmCarregando (UserForm) do ATUALIZAR."""

    def run(self):
        import win32con
        import win32gui
        import win32process
        t0 = time.time()
        while not self.parar.is_set() and time.time() - t0 < self.teto:
            achou = []

            def cb(h, _):
                try:
                    _, pid = win32process.GetWindowThreadProcessId(h)
                    if pid == self.pid and win32gui.GetClassName(h) == '#32770' and win32gui.IsWindowVisible(h):
                        achou.append(h)
                except Exception:  # noqa: BLE001
                    pass
                return True
            try:
                win32gui.EnumWindows(cb, None)
            except Exception:  # noqa: BLE001
                pass
            for h in achou:
                filhos = []

                def cc(ch, _):
                    try:
                        tx = win32gui.GetWindowText(ch)
                        if tx:
                            filhos.append(tx)
                    except Exception:  # noqa: BLE001
                        pass
                    return True
                try:
                    win32gui.EnumChildWindows(h, cc, None)
                except Exception:  # noqa: BLE001
                    pass
                self.vistos.append({'titulo': win32gui.GetWindowText(h), 'textos': filhos,
                                    'quando': time.strftime('%H:%M:%S')})
                try:
                    win32gui.PostMessage(h, win32con.WM_CLOSE, 0, 0)
                except Exception:  # noqa: BLE001
                    pass
            time.sleep(0.6 if achou else 0.25)


def secao(L, nome, fn, *a):
    print('>>', nome, time.strftime('%H:%M:%S'), flush=True)
    return V.secao(L, nome, fn, *a)


# ------------------------------------------------------------------------------------------- OC-0072 leitura
RX_DATA = re.compile(r'^\d{2}/\d{2}/\d{4}$')
RX_DATA_HORA = re.compile(r'^\d{2}/\d{2}/\d{4} \d{2}:\d{2}:\d{2}$')


def ler_datas(qa, aba, ender, maxc=400):
    """Text/NumberFormat/NumberFormatLocal/Value2 das celulas nao vazias (no maximo maxc) do intervalo."""
    ws = qa.ws(aba)
    rng = ws.Range(ender)
    v2 = rng.Value2
    if not isinstance(v2, tuple):
        v2 = ((v2,),)
    out, ruins = [], []
    n_nv = 0
    for i, lin in enumerate(v2):
        for j, x in enumerate(lin):
            if x in (None, ''):
                continue
            n_nv += 1
            if len(out) >= maxc:
                continue
            c = rng.Cells(i + 1, j + 1)
            tx = t(c.Text)
            d = {'ref': c.Address.replace('$', ''), 'v2': x, 'txt': tx, 'fmt': t(c.NumberFormat),
                 'fmt_local': t(c.NumberFormatLocal)}
            if isinstance(x, float):
                tem_hora = abs(x - int(x)) > 1e-9
                ok = bool(RX_DATA_HORA.match(tx)) if tem_hora else bool(RX_DATA.match(tx) or RX_DATA_HORA.match(tx))
            else:
                ok = not re.search(r'y{2,4}|a{4}', tx)
            if re.search(r'yyyy|aaaa|/y|/a', tx):
                ok = False
            d['ok'] = ok
            out.append(d)
            if not ok:
                ruins.append(d)
    try:
        fmt_unico = t(rng.NumberFormat)
    except Exception:  # noqa: BLE001
        fmt_unico = None
    return {'intervalo': f'{aba}!{ender}', 'nao_vazias': n_nv, 'lidas': len(out), 'fmt_do_intervalo': fmt_unico,
            'ruins': ruins[:10], 'n_ruins': len(ruins), 'amostra': out[:3] + out[-2:]}


def alvos_0072(qa):
    out = {}
    out['Configuracao_D'] = ler_datas(qa, 'Configuração', 'D26:D125')
    out['Configuracao_AD'] = ler_datas(qa, 'Configuração', 'AD26:AD125')
    out['Eng_Saida_AC'] = ler_datas(qa, 'Eng_Saida', 'AC3:AC182')
    ws = qa.ws('Westgard_Fotos')
    n = ws.Cells(ws.Rows.Count, 6).End(-4162).Row
    out['Westgard_Fotos_F_inicio'] = ler_datas(qa, 'Westgard_Fotos', 'F4:F40')
    out['Westgard_Fotos_F_fim'] = ler_datas(qa, 'Westgard_Fotos', f'F{max(4, n - 40)}:F{n}')
    out['Westgard_Fotos_F_fmt_coluna'] = t(ws.Range('F:F').NumberFormat)
    sl = qa.ws('Sync_Log')
    m = sl.Cells(sl.Rows.Count, 1).End(-4162).Row
    for col in ('A', 'B', 'L', 'M'):
        out[f'Sync_Log_{col}'] = ler_datas(qa, 'Sync_Log', f'{col}4:{col}{max(4, m)}')
    out['Sync_Log_ultima_linha'] = m
    e = qa.ws('Estatística')
    out['Estatistica_E8F9C10'] = {a: {'f': t(e.Range(a).Formula)[:120], 'txt': t(e.Range(a).Text), 'v2': e.Range(a).Value2,
                                      'fmt': t(e.Range(a).NumberFormat), 'fmt_local': t(e.Range(a).NumberFormatLocal)}
                                  for a in ('E8', 'F8', 'E9', 'F9', 'C10', 'D8', 'D9', 'B10')}
    out['LotesStore_U1'] = ler_datas(qa, 'LotesStore', 'U1')
    out['n_ruins_total'] = sum(v.get('n_ruins', 0) for v in out.values() if isinstance(v, dict))
    return out


# ------------------------------------------------------------------------------------------- OC-0068 listas
def avaliar_lista(qa, cel, validos, invalidos, sep):
    vv = cel.Validation
    d = {'tipo': vv.Type, 'formula1': t(vv.Formula1)[:200], 'alertstyle': vv.AlertStyle, 'showerror': bool(vv.ShowError),
         'incelldropdown': bool(vv.InCellDropdown), 'msg': t(vv.ErrorMessage)[:240], 'titulo': t(vv.ErrorTitle)}
    f1 = d['formula1']
    d['itens_menu'] = f1.split(sep) if not f1.startswith('=') else f1
    antes = cel.Formula
    ev = qa.ex.xl.EnableEvents
    qa.ex.xl.EnableEvents = False
    res = {}
    try:
        for x in list(validos) + list(invalidos):
            cel.Value = x
            res[x] = bool(cel.Validation.Value)
        cel.Formula = antes
    finally:
        qa.ex.xl.EnableEvents = ev
    d['aceita'] = res
    d['ok'] = (all(res[x] for x in validos) and not any(res[x] for x in invalidos)
               and (isinstance(d['itens_menu'], str) or sorted(d['itens_menu']) == sorted(validos)))
    return d


def s_listas(qa, senha):
    sep = t(qa.ex.xl.International[4])
    out = {'sep_lista_excel': sep}
    lo, _ = V.manual_cols(qa)
    _, r = C.linha_livre_manual(qa)
    eqs = ['DIMENSION 1', 'DIMENSION 2']
    mzs = ['SORO', 'URINA']
    mot = ['Falha do interfaceamento', 'Resultado não transmitido', 'Equipamento em contingência', 'Outro']
    def av(*a):
        try:
            return avaliar_lista(*a)
        except Exception as ex:  # noqa: BLE001
            return {'erro': repr(ex)[:240], 'ok': False, 'msg': ''}
    out['DR_EQUIPAMENTO'] = av(qa, lo.ListColumns('EQUIPAMENTO').DataBodyRange.Cells(r, 1), eqs,
                                          ['DIMENSION 9', ','.join(eqs), 'DIMENSION'], sep)
    out['DR_MATRIZ'] = av(qa, lo.ListColumns('MATRIZ').DataBodyRange.Cells(r, 1), mzs,
                                     ['PLASMA', ','.join(mzs)], sep)
    out['DR_MOTIVO'] = av(qa, lo.ListColumns('MOTIVO').DataBodyRange.Cells(r, 1), mot,
                                     ['Outro: TESTE-AUDITORIA', ','.join(mot), 'Falha'], sep)
    # vizinhos que nao eram literais (devem continuar como estavam)
    viz = {}
    for col in ('LOTE', 'ANALITO', 'NIVEL', 'DATA', 'RESULTADO'):
        c = lo.ListColumns(col).DataBodyRange.Cells(r, 1)
        try:
            viz[col] = [c.Validation.Type, t(c.Validation.Formula1)[:60], t(c.Validation.ErrorMessage)[:120]]
        except Exception as ex:  # noqa: BLE001
            viz[col] = repr(ex)[:60]
    out['DR_vizinhos'] = viz
    rg = qa.ws('Registros')
    out['Registros_H4'] = av(qa, rg.Range('H4'), ['Sim', 'Não'], ['Sim,Não', 'Talvez', 'S'], sep)
    out['Registros_F7_foi_calibrado'] = av(qa, rg.Range('F7'), ['Sim', 'Não'], ['Sim,Não'], sep)
    p = qa.ws('Painel')
    out['Painel_L4'] = av(qa, p.Range('L4'), eqs, ['DIMENSION 9', ','.join(eqs)], sep)
    out['Painel_K4_vizinho'] = t(p.Range('K4').Validation.Formula1)[:200].split(sep)[:6]
    out['Painel_M4_vizinho'] = t(p.Range('M4').Validation.Formula1).split(sep)
    out['ok'] = all(out[k]['ok'] for k in ('DR_EQUIPAMENTO', 'DR_MATRIZ', 'DR_MOTIVO', 'Registros_H4', 'Painel_L4'))
    out['msg_seta'] = {k: ('na seta da célula' in out[k]['msg']) for k in ('DR_EQUIPAMENTO', 'DR_MATRIZ', 'DR_MOTIVO',
                                                                              'Registros_H4', 'Painel_L4')}
    return out


def tbl_hash(lo):
    if lo.ListRows.Count == 0:
        return (0, '')
    v = lo.DataBodyRange.Value2
    return (len(v), hashlib.sha256(repr(v).encode()).hexdigest()[:16])


def s_registros(qa, senha):
    """Filtro H4 'Calibrado?' e datas De/Ate (E4/F4) digitadas como o usuario: so a vista muda."""
    rg = qa.ws('Registros')
    rg.Activate()
    qa.ex.esperar()
    lo = qa.lo('tblRegistros')
    out = {'tbl_antes': tbl_hash(lo)}

    def vis():
        n = lo.ListRows.Count
        return sum(1 for i in range(1, n + 1) if not lo.DataBodyRange.Rows(i).EntireRow.Hidden), n
    out['vis_antes'] = vis()
    col_cal = [str(c.Name) for c in lo.ListColumns]
    out['colunas'] = col_cal
    vals = lo.DataBodyRange.Value2 if lo.ListRows.Count else ()
    out['valores_col_calibrado'] = dict(collections.Counter(t(l[7]) for l in vals)) if vals else {}
    passos = []
    for rot, cel, v in (('H4=Sim', 'H4', 'Sim'), ('H4=Não', 'H4', 'Não'), ('H4 vazio', 'H4', None),
                        ('E4 De 01/10/2026 digitado', 'E4', '01/10/2026'), ('F4 Até 31/10/2026 digitado', 'F4', '31/10/2026')):
        try:
            if v is None:
                rg.Range(cel).ClearContents()
                qa.ex.esperar()
            else:
                C.por(qa, rg, cel, v, local=True)
            time.sleep(0.5)
            passos.append({'passo': rot, 'vis_tot': vis(), 'tbl': tbl_hash(lo), 'H4_txt': t(rg.Range('H4').Text),
                           'E4': [t(rg.Range('E4').Text), t(rg.Range('E4').NumberFormat), t(rg.Range('E4').NumberFormatLocal),
                                  rg.Range('E4').Value2],
                           'F4': [t(rg.Range('F4').Text), t(rg.Range('F4').NumberFormat), rg.Range('F4').Value2],
                           'linha5': t(rg.Range('A5').Text)[:200]})
        except Exception as ex:  # noqa: BLE001
            passos.append({'passo': rot, 'erro': repr(ex)[:200]})
    for cel in ('E4', 'F4'):
        try:
            rg.Range(cel).ClearContents()
            qa.ex.esperar()
        except Exception:  # noqa: BLE001
            pass
    out['passos'] = passos
    out['vis_depois_limpar'] = vis()
    out['tbl_depois'] = tbl_hash(lo)
    out['so_vista'] = out['tbl_antes'] == out['tbl_depois'] and all(x.get('tbl') == out['tbl_antes'] for x in passos
                                                                    if 'tbl' in x)
    qa.ws('Painel').Activate()
    qa.ex.esperar()
    return out


# ------------------------------------------------------------------------------------------- manuais
def hora_livre(fin, dia, analito, usadas, gap=40):
    ocup = [r['DATA_HORA'].hour * 60 + r['DATA_HORA'].minute for r in fin
            if isinstance(r['DATA_HORA'], dt.datetime) and r['DATA_HORA'].date() == dia and up(r['ANALITO']) == up(analito)]
    for g in (gap, 32):
        for m in range(60, 23 * 60, 3):
            if all(abs(m - u) >= g for u in ocup) and all(abs(m - u) >= 3 for u in usadas):
                usadas.append(m)
                return f'{m // 60:02d}:{m % 60:02d}'
    raise RuntimeError('sem hora livre')


def painel_serie(qa, eq, lote, ano, mes):
    p = qa.ws('Painel')
    for c, v in (('L4', eq), ('C3', 'Glicose'), ('H3', "'" + lote), ('K3', ano), ('K4', MESES[mes - 1])):
        C.por(qa, p, c, v, local=False)
    time.sleep(1.2)
    qa.ex.esperar()
    es = qa.eng_saida()
    vals1 = [l[0] for l in es['val']] if es['n'] else []
    return {'L4': t(p.Range('L4').Value), 'engPainel': qa.painel_stats(), 'eng_S1': t(es.get('equip')),
            'eng_E1': t(es.get('lote')), 'n_pontos': es['n'], 'tem_123456_N1': any(isinstance(x, float) and
                                                                                   abs(x - 123.456) < 1e-9 for x in vals1)}


def s_manuais(qa, fin, sep):
    out = {}
    p = qa.ws('Painel')
    lo, cab = V.manual_cols(qa)
    out['colunas'] = cab
    # regras de lote e de/para reais da copia
    regras = [dict(zip([str(c.Name) for c in qa.lo('tblLotesMaterial').ListColumns], [t(x) for x in l]))
              for l in (qa.lo('tblLotesMaterial').DataBodyRange.Value if qa.lo('tblLotesMaterial').ListRows.Count else ())]
    fora = [r['LOTE'] for r in regras if up(r.get('REGRA')) == 'FORA_DO_CQ']
    so = [(r['LOTE'], {up(a) for a in r.get('ANALITOS', '').split(';') if a.strip()}) for r in regras
          if up(r.get('REGRA')) == 'SO_ANALITOS']
    mzu = [r['LOTE'] for r in regras if up(r.get('REGRA')) == 'MATRIZ' and up(r.get('MATRIZ')) == 'URINA']
    dp = qa.ler('tblDeParaAnalitos')
    urina_qcini = {up(d.get('ANALITO_QCINI')) for d in dp if up(d.get('MATRIZ')) == 'URINA'}
    urina_orig = {up(d.get('ANALITO_ORIGEM')) for d in dp if up(d.get('MATRIZ')) == 'URINA'}
    cat = [t(x[0]) for x in qa.espec().Range('A4:A43').Value if x[0]]
    an_mzsem = next(a for a in cat if up(a) not in urina_qcini and up(a) not in urina_orig and 'COLESTEROL' in up(a)) \
        if any('COLESTEROL' in up(a) and up(a) not in urina_qcini for a in cat) else \
        next(a for a in cat if up(a) not in urina_qcini and up(a) not in urina_orig)
    lote_so, an_so_lista = so[0]
    an_so = 'Glicose' if 'GLICOSE' not in an_so_lista else next(a for a in cat if up(a) not in an_so_lista)
    lote_u = '6354' if '6354' in mzu else mzu[0]
    out['escolhas'] = {'lote_FORA': fora[0], 'lote_SO': lote_so, 'analito_fora_do_SO': an_so, 'lote_URINA': lote_u,
                       'analito_sem_depara_urina': an_mzsem, 'n_regras': len(regras)}
    # referencias reais (interfaceamento, ATIVO, Glicose N1)
    def reais(eq):
        return sorted([r for r in fin if r['STATUS_ANALITICO'] == 'ATIVO' and r['ORIGEM_RESULTADO'] == 'INTERFACEAMENTO'
                       and t(r['ANALITO']) == 'Glicose' and int(r['NIVEL'] or 0) == 1
                       and isinstance(r['DATA_HORA'], dt.datetime) and t(r['EQUIPAMENTO']) == eq],
                      key=lambda r: r['DATA_HORA'])
    r1, r2 = reais('DIMENSION 1'), reais('DIMENSION 2')
    R, S1, S2, S3 = r1[-40], r1[-41], r1[-42], r1[-43]
    R2 = r2[-30]
    dup = [r for r in fin if r['STATUS_ANALITICO'] == 'DUPLICIDADE_ORIGEM' and r['ORIGEM_RESULTADO'] == 'INTERFACEAMENTO']
    ref = lambda r: {k: t(r[k]) for k in ('ID_REGISTRO', 'EQUIPAMENTO', 'LOTE', 'NIVEL', 'ANALITO', 'RESULTADO',  # noqa: E731
                                          'DATA_HORA', 'RUN', 'STATUS_ANALITICO')}
    out['referencias'] = {'R': ref(R), 'S1': ref(S1), 'S2': ref(S2), 'S3': ref(S3), 'R2_DIM2': ref(R2),
                          'DUP': ref(dup[0]) if dup else None}
    numid = lambda r: re.sub(r'\D', '', t(r['ID_REGISTRO']))  # noqa: E731
    dia, lote = R['DATA_HORA'].date(), t(R['LOTE'])
    dia2, lote2 = R2['DATA_HORA'].date(), t(R2['LOTE'])
    sdia, sdia2 = dia.strftime('%d/%m/%Y'), dia2.strftime('%d/%m/%Y')
    us, us2 = [], []
    # Painel ANTES (OC-0068: a serie do manual DIMENSION 2)
    out['serie_antes'] = {'DIM2': painel_serie(qa, 'DIMENSION 2', lote2, dia2.year, dia2.month),
                          'DIM1': painel_serie(qa, 'DIMENSION 1', lote2, dia2.year, dia2.month)}
    C.por(qa, p, 'L4', 'DIMENSION 1', local=False)
    C.por(qa, p, 'C3', 'Glicose', local=False)
    try:
        qa.run('mIntegracao.IrDigitarResultados', teto=120)
    except Exception as ex:  # noqa: BLE001
        out['ir_erro'] = repr(ex)[:200]
    out['aba_ativa'] = t(qa.wb.ActiveSheet.Name)
    qa.ex.xl.EnableEvents = True
    linhas = {}
    vig = Caixas(qa.ex.pid, teto=1800)
    vig.start()
    # o item do menu que tem "DIMENSION 2" (na base e o item unico "DIMENSION 1,DIMENSION 2")
    _, rr = C.linha_livre_manual(qa)
    f1 = t(lo.ListColumns('EQUIPAMENTO').DataBodyRange.Cells(rr, 1).Validation.Formula1)
    item_dim2 = next(x for x in f1.split(sep) if 'DIMENSION 2' in x)
    out['item_menu_DIM2'] = item_dim2
    eq1 = ('EQUIPAMENTO', 'DIMENSION 1')

    def novo(rot, campos):
        _, r = C.linha_livre_manual(qa)
        for col, v in campos:
            V.escrever_celula(qa, r, col, v)
            time.sleep(0.2)
        linhas[rot] = r
        return r

    L1 = "'" + lote
    try:
        novo('M01_ok_menu_DIM2', [('DATA', sdia2), ('HORA', hora_livre(fin, dia2, 'Glicose', us2)),
                                  ('EQUIPAMENTO', item_dim2), ('LOTE', "'" + lote2), ('NIVEL', '1'), ('ANALITO', 'Glicose'),
                                  ('RESULTADO', '123,456'), ('MOTIVO', MOTIVO)])
        novo('M02_implausivel', [('DATA', sdia), ('HORA', hora_livre(fin, dia, 'Glicose', us)), eq1, ('LOTE', L1),
                                 ('NIVEL', '1'), ('ANALITO', 'Glicose'), ('RESULTADO', '-1'), ('MOTIVO', MOTIVO)])
        novo('M03_lote_FORA', [('DATA', sdia), ('HORA', hora_livre(fin, dia, 'Glicose', us)), eq1, ('LOTE', "'" + fora[0]),
                               ('NIVEL', '1'), ('ANALITO', 'Glicose'), ('RESULTADO', '123,456'), ('MOTIVO', MOTIVO)])
        novo('M04_lote_SO', [('DATA', sdia), ('HORA', hora_livre(fin, dia, an_so, us)), eq1, ('LOTE', "'" + lote_so),
                             ('NIVEL', '1'), ('ANALITO', an_so), ('RESULTADO', '123,456'), ('MOTIVO', MOTIVO)])
        novo('M05_urina_sem_depara', [('DATA', sdia), ('HORA', hora_livre(fin, dia, an_mzsem, us)), eq1,
                                      ('LOTE', "'" + lote_u), ('NIVEL', '1'), ('ANALITO', an_mzsem), ('RESULTADO', '1,234'),
                                      ('MOTIVO', MOTIVO)])
        novo('M06_urina_Creatinina', [('DATA', sdia), ('HORA', hora_livre(fin, dia, 'CRE2', us)), eq1,
                                      ('LOTE', "'" + lote_u), ('NIVEL', '1'), ('ANALITO', 'Creatinina'),
                                      ('RESULTADO', '12,345'), ('MOTIVO', MOTIVO)])
        novo('M07_subst_ok', [('SUBSTITUI_ID', numid(S1)), ('RESULTADO', '123,456'), ('MOTIVO', MOTIVO)])
        r8 = novo('M08_subst_dup', [('SUBSTITUI_ID', numid(S1))])
        hs1 = S1['DATA_HORA'] + dt.timedelta(minutes=2)
        V.escrever_celula(qa, r8, 'HORA', hs1.strftime('%H:%M'))
        V.escrever_celula(qa, r8, 'RESULTADO', '111,111')
        V.escrever_celula(qa, r8, 'MOTIVO', MOTIVO)
        novo('M09_subst_inexistente', [('SUBSTITUI_ID', '999999999'), ('DATA', sdia),
                                       ('HORA', hora_livre(fin, dia, 'Glicose', us)), eq1, ('LOTE', L1), ('NIVEL', '2'),
                                       ('ANALITO', 'Glicose'), ('RESULTADO', '123,456'), ('MOTIVO', MOTIVO)])
        id_m02 = V.ler_manual(qa, linhas['M02_implausivel'])['ID_REGISTRO'][0]
        novo('M10_subst_id_manual', [('SUBSTITUI_ID', id_m02), ('DATA', sdia), ('HORA', hora_livre(fin, dia, 'Glicose', us)),
                                     eq1, ('LOTE', L1), ('NIVEL', '2'), ('ANALITO', 'Glicose'), ('RESULTADO', '123,456'),
                                     ('MOTIVO', MOTIVO)])
        r11 = novo('M11_subst_outra_chave', [('SUBSTITUI_ID', numid(S2))])
        V.escrever_celula(qa, r11, 'NIVEL', '2')
        V.escrever_celula(qa, r11, 'RESULTADO', '123,456')
        V.escrever_celula(qa, r11, 'MOTIVO', MOTIVO)
        novo('M12_subst_sem_motivo', [('SUBSTITUI_ID', numid(S3)), ('RESULTADO', '123,456')])
        if dup:
            novo('M13_subst_retransmissao', [('SUBSTITUI_ID', numid(dup[0])), ('RESULTADO', '123,456'), ('MOTIVO', MOTIVO)])
        novo('M14_data_futura', [('DATA', '31/12/2026'), ('HORA', '03:00'), eq1, ('LOTE', L1), ('NIVEL', '1'),
                                 ('ANALITO', 'Glicose'), ('RESULTADO', '123,456'), ('MOTIVO', MOTIVO)])
        hc = (R['DATA_HORA'] + dt.timedelta(minutes=5)).strftime('%H:%M')
        novo('M15_conflito', [('DATA', sdia), ('HORA', hc), eq1, ('LOTE', L1), ('NIVEL', '1'), ('ANALITO', 'Glicose'),
                              ('RESULTADO', '111,111'), ('MOTIVO', MOTIVO)])
        hrep = hora_livre(fin, dia, 'Glicose', us)
        for rot in ('M16_repetido_1', 'M17_repetido_2'):
            novo(rot, [('DATA', sdia), ('HORA', hrep), eq1, ('LOTE', L1), ('NIVEL', '2'), ('ANALITO', 'Glicose'),
                       ('RESULTADO', '123,456'), ('MOTIVO', MOTIVO)])
        novo('M18_sem_nivel', [('DATA', sdia), ('HORA', hora_livre(fin, dia, 'Glicose', us)), eq1, ('LOTE', L1),
                               ('ANALITO', 'Glicose'), ('RESULTADO', '123,456'), ('MOTIVO', MOTIVO)])
        out['avisos_ao_digitar'] = list(vig.vistos)
        n_av = len(vig.vistos)
        out['pre'] = {k: {c: V.ler_manual(qa, r)[c][0] for c in ('ID_REGISTRO', 'DATA', 'HORA', 'EQUIPAMENTO', 'MATRIZ',
                                                                 'LOTE', 'NIVEL', 'ANALITO', 'RESULTADO', 'MOTIVO',
                                                                 'SUBSTITUI_ID')} for k, r in linhas.items()}
        t0 = time.time()
        ret = t(qa.run('mIntegracao.AtualizarDadosAutomatico', teto=1500))
        out['atualizar'] = {'ret': ret[:700], 's': round(time.time() - t0, 1)}
        out['avisos_no_atualizar'] = vig.vistos[n_av:]
    finally:
        vig.parar.set()
    leit = {k: V.ler_manual(qa, r) for k, r in linhas.items()}
    fin1 = qa.final_por_id()
    res = {}
    for k, v in leit.items():
        idm = v['ID_REGISTRO'][0]
        f = fin1.get(idm, {})
        det_txt = v['DETALHE'][1]
        mot = t(f.get('MOTIVO_EXCLUSAO_AUTOMATICA'))
        res[k] = {'ID': idm, 'STATUS': v['STATUS'][0], 'RUN_txt': v['RUN'][1], 'RUN_v': v['RUN'][0],
                  'DETALHE_txt': det_txt, 'final': {c: t(f.get(c)) for c in ('STATUS_ANALITICO', 'PARTICIPA_ESTATISTICA',
                                                                             'ANALITO', 'MATRIZ', 'EQUIPAMENTO', 'RUN')},
                  'MOTIVO_EXCL': mot, 'sem_acento': sem_acento_achados(det_txt),
                  'DETALHE_igual_MOTIVO_acentuado_por_mim': det_txt == acentuar(mot),
                  'MOTIVO_tem_acento': mot != tira_acento(mot)}
    out['manuais'] = res
    # o resultado substituido (S1) na tblCQ_Final
    s1 = fin1.get(t(S1['ID_REGISTRO']), {})
    out['S1_depois'] = {c: t(s1.get(c)) for c in ('STATUS_ANALITICO', 'PARTICIPA_ESTATISTICA', 'MOTIVO_EXCLUSAO_AUTOMATICA',
                                                  'RUN')}
    # varredura de todas as linhas da aba (RUN/DETALHE "0", palavras sem acento)
    n = lo.ListRows.Count
    vr, vd = lo.ListColumns('RUN').DataBodyRange, lo.ListColumns('DETALHE').DataBodyRange
    vi = lo.ListColumns('ID_REGISTRO').DataBodyRange.Value
    zeros, sa = [], []
    for i in range(1, min(n, 400) + 1):
        if vi[i - 1][0] in (None, ''):
            continue
        a, b = t(vr.Cells(i, 1).Text), t(vd.Cells(i, 1).Text)
        if a == '0' or b == '0':
            zeros.append((t(vi[i - 1][0]), a, b))
        if sem_acento_achados(b):
            sa.append((t(vi[i - 1][0]), sem_acento_achados(b), b[:160]))
    out['varredura'] = {'linhas_com_id': sum(1 for x in vi if x[0] not in (None, '')), 'com_zero': zeros,
                        'com_palavra_sem_acento': sa}
    # Painel DEPOIS (OC-0068)
    out['serie_depois'] = {'DIM2': painel_serie(qa, 'DIMENSION 2', lote2, dia2.year, dia2.month),
                           'DIM1': painel_serie(qa, 'DIMENSION 1', lote2, dia2.year, dia2.month)}
    out['M01_serie'] = {'dia': sdia2, 'lote': lote2,
                        'n_DIM2_antes_depois': [out['serie_antes']['DIM2']['engPainel'][0][1],
                                                out['serie_depois']['DIM2']['engPainel'][0][1]],
                        'n_DIM1_antes_depois': [out['serie_antes']['DIM1']['engPainel'][0][1],
                                                out['serie_depois']['DIM1']['engPainel'][0][1]],
                        'DIM1_igual': out['serie_antes']['DIM1']['engPainel'] == out['serie_depois']['DIM1']['engPainel'],
                        'valor_no_grafico_DIM2': out['serie_depois']['DIM2']['tem_123456_N1'],
                        'valor_no_grafico_DIM1': out['serie_depois']['DIM1']['tem_123456_N1']}
    return out


# ------------------------------------------------------------------------------------------- execucao
def rodar(rotulo):
    origem = BACKUP if rotulo == 'base41' else os.path.join(C.REPO, f'QC_{PROD}.xlsm')
    destino = os.path.join(PASTA, f'{PROD}_F1_{rotulo}.xlsm')
    if os.path.exists(destino):
        os.remove(destino)
    h0 = C.sha(origem)
    shutil.copy2(origem, destino)
    L = C.Log(os.path.join(PASTA, f'{PROD}_F1_{rotulo}.json'))
    L['ambiente'] = {'origem': os.path.relpath(origem, C.REPO), 'sha_origem_12': h0, 'copia': os.path.basename(destino),
                     'sha_copia_12': C.sha(destino)}
    senha = C.senha_planilha()
    qa = C.qa_final.QA(PROD, destino)
    L['excel_pid'] = qa.ex.pid
    try:
        p, e = qa.ws('Painel'), qa.ws('Estatística')
        p.Activate()
        qa.ex.esperar()
        fin = qa.ler('tblCQ_Final')
        L['linhas_tblCQ_Final'] = len(fin)
        criar_usuario(qa, USR, 'ANALISTA')
        L['login'] = fazer_login(qa, USR)
        for w in (p, e):
            try:
                w.Unprotect(senha)
            except Exception:  # noqa: BLE001
                pass
        sep = t(qa.ex.xl.International[4])
        secao(L, 'O72_apos_login', alvos_0072, qa)
        secao(L, 'O68_listas', s_listas, qa, senha)
        secao(L, 'O68_O72_registros', s_registros, qa, senha)
        secao(L, 'O21_O68_manuais', s_manuais, qa, fin, sep)
        secao(L, 'O72_depois_atualizar', alvos_0072, qa)
        if rotulo == 'nova':
            sc = os.path.join(PASTA, f'{PROD}_F1_nova_SaveCopyAs.xlsm')
            if os.path.exists(sc):
                os.remove(sc)
            qa.wb.SaveCopyAs(sc)
            r = X.resumo(sc)
            z = zipfile.ZipFile(sc)
            L['O72_O68_xml_depois_do_uso'] = {'arquivo': os.path.basename(sc),
                                              'numFmt_com_letra_literal': r['numFmt_com_letra_literal'],
                                              'uso_literais': r['uso_literais'], 'x12ac_list': r['x12ac_list'],
                                              'formatos_de_data': r['formatos_de_data'],
                                              'validacoes_DR_Registros_Painel': [x for x in r['validacoes_lista'] if
                                                                                 x[0] in ('Digitar Resultados', 'Registros',
                                                                                          'Painel')]}
            z.close()
    except Exception as ex:  # noqa: BLE001
        L['ERRO_FATAL'] = {'erro': repr(ex), 'trace': traceback.format_exc()[-3000:]}
    finally:
        try:
            qa.fechar(salvar=False)
        except Exception:  # noqa: BLE001
            pass
    L['origem_intacta'] = {'sha_antes': h0, 'sha_depois': C.sha(origem), 'igual': h0 == C.sha(origem)}
    print('FIM', rotulo, flush=True)


if __name__ == '__main__':
    t0 = time.time()
    rodar(sys.argv[1])
    print(f'tempo {time.time() - t0:.0f}s', flush=True)
