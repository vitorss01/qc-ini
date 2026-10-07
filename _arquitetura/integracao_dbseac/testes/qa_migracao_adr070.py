# -*- coding: utf-8 -*-
"""QA da MIGRACAO do ADR-070.

Uso: python qa_migracao_adr070.py <Bioquimica|Hematologia> <arquivo.xlsm> [saida.json] [--original <producao.xlsm>]

  <arquivo.xlsm> ... sem --original: um arquivo AINDA NO LAYOUT ANTIGO (antes do instalar_adr070.py).
                     com --original (o entregar.py: a copia INSTALADA): a migracao e testada sobre a producao
                     original (layout antigo) e a copia instalada e conferida contra ela (M06).
  --original ....... a producao ANTES da instalacao (tambem pela variavel QC_PROD_ORIGINAL). Nunca e aberta: a
                     suite trabalha em copias proprias.

Nenhum arquivo recebido e alterado (SHA-256 conferido): tudo acontece em copias _mig_tmp/_m06_tmp, apagadas no fim.
A migracao roda SEMPRE em MODO_FONTE = HISTORICO na copia (revisao 07/10/2026): em SEAC o instalador nao roda o
ATUALIZAR DADOS e o M03 compararia a tblCQ_Final de ANTES da migracao com ela mesma -- PASS vazio. Se a consulta
nova nao rodou sobre a tabela migrada, o M03 FALHA.
  1. na copia, com o VBA e as consultas ANTIGAS, inativa pelo caminho real (evento ligado) como o laboratorio fez ate
     hoje: A (so o numero, com comentario tecnico), B (com comentario e caixa DESMARCADA), uma linha VAZIA no meio,
     C (sem comentario: E01), um ID inexistente (E03); ATUALIZAR DADOS; foto da tabela e da tblCQ_Final; salva a copia;
  2. roda instalar_adr070.py na copia (subprocesso, como o entregar.py);
  3. M01 layout novo sem formula; M02 cada linha antiga na MESMA posicao com ID, caixa, data e usuario identicos; a linha
     vazia continua vazia; ANALITO das linhas antigas = o da tblCQ_Final (vazio no ID inexistente); MOTIVO vazio;
     M03 tblCQ_Final: os mesmos INATIVADOS com a mesma plotagem, justificativa, governanca, data e usuario da
     inativacao; QA: E01 do C e E03 do inexistente continuam, nenhum E10/E11;
  4. M04 idempotencia: o instalador de novo nao migra nada e a tabela fica identica.
  5. M06 (com --original) as inativacoes REAIS sobrevivem: a copia instalada tem o mesmo conjunto (ID, REGISTRAR - LJ,
     DATA_INATIVACAO, USUARIO) da producao original, todo ID com resultado ficou com ANALITO, e um ATUALIZAR DADOS
     (HISTORICO, consulta nova) numa copia dela da os mesmos INATIVADOS com a mesma plotagem, sem E10/E11 nas linhas
     migradas. Producao ja no layout novo: M01-M04 nao se aplicam (dito no log) e o M06 roda.
"""
import json
import os
import shutil
import subprocess
import sys
import traceback

AQUI = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, AQUI)
from qa_final import QA, reg, RES, NAO, sha256_arquivo  # noqa: E402
sys.path.insert(0, os.path.join(AQUI, '..'))
import instalar_adr070 as i70  # noqa: E402

SIM = 'SIM'
TB = 'tblInativacao_NaoConformes'
CAB_NOVO = ['ID_REGISTRO', 'ANALITO', 'REGISTRAR - LJ', 'MOTIVO', 'DATA_INATIVACAO', 'USUARIO']
ANTIGAS = ['ID_REGISTRO', 'REGISTRAR - LJ', 'DATA_INATIVACAO', 'USUARIO']
COLS_FIN = ('STATUS_ANALITICO', 'TIPO_PLOTAGEM_LJ', 'PARTICIPA_ESTATISTICA', 'INATIVACAO_REGISTRADA', 'TEM_JUSTIFICATIVA',
            'GOVERNANCA', 'COMENTARIO_TECNICO', 'DATA_INATIVACAO', 'USUARIO_INATIVACAO')


def linhas_tabela(q, n):
    """As n primeiras linhas da tabela, em Value2 (data = numero de serie, sem fuso)."""
    lo = q.lo(TB)
    out = []
    for k in range(1, n + 1):
        out.append({c.Name: c.DataBodyRange.Cells(k, 1).Value2 for c in lo.ListColumns})
    return out


def tabela_inteira(q):
    lo = q.lo(TB)
    return [c.Name for c in lo.ListColumns], lo.DataBodyRange.Value2


def foto_final(q, ids):
    f = q.final_por_id()
    return {i: ({k: f[i][k] for k in COLS_FIN} if i in f else None) for i in ids}


def instalar(produto, caminho, log):
    env = dict(os.environ, PYTHONIOENCODING='utf-8')
    with open(log, 'w', encoding='utf-8') as fh:
        r = subprocess.run([sys.executable, '-u', os.path.join(AQUI, '..', 'instalar_adr070.py'), produto, caminho],
                           cwd=os.path.join(AQUI, '..'), env=env, stdout=fh, stderr=subprocess.STDOUT)
    with open(log, encoding='utf-8', errors='replace') as fh:
        texto = fh.read()
    return r.returncode == 0, texto


def layout_novo(produto, arquivo, pasta):
    tmp = os.path.join(pasta, 'layout_tmp_' + os.path.basename(arquivo))
    shutil.copy2(arquivo, tmp)
    q = QA(produto, tmp)
    try:
        return 'MOTIVO' in [c.Name for c in q.lo(TB).ListColumns]
    finally:
        q.fechar(salvar=False)
        try:
            os.remove(tmp)
        except OSError:
            pass


def inativados_final(q):
    return {i: r['TIPO_PLOTAGEM_LJ'] for i, r in q.final_por_id().items() if r['STATUS_ANALITICO'] == 'INATIVADO'}


def m06_sobrevivem(produto, instalado, original):
    """Achado 14: na entrega (inclusive em SEAC) as inativacoes reais da producao sobrevivem a migracao."""
    pasta = os.path.dirname(instalado)
    t_orig = os.path.join(pasta, 'm06_orig_tmp.xlsm')
    t_inst = os.path.join(pasta, 'm06_inst_tmp.xlsm')
    shutil.copy2(original, t_orig)
    shutil.copy2(instalado, t_inst)
    q = None
    try:
        q = QA(produto, t_orig)
        conj0 = i70.conjunto_inativacao(q.wb)
        inat0 = inativados_final(q)
        fin0 = set(q.final_por_id())
        q.fechar(salvar=False)
        q = QA(produto, t_inst)
        conj1 = i70.conjunto_inativacao(q.wb)
        lo = q.lo(TB)
        pref = i70.prefixo_id(q.wb)
        conhecidos = i70.ids_com_resultado(q.wb)
        ids = [r[0] for r in lo.ListColumns('ID_REGISTRO').DataBodyRange.Value]
        ans = [r[0] for r in lo.ListColumns('ANALITO').DataBodyRange.Value]
        sem_an = [(k, i70.cd._norm_id(i, pref)) for k, (i, a) in enumerate(zip(ids, ans), start=1)
                  if i70.cd._norm_id(i, pref) in conhecidos and a in (None, '')]
        modo0 = q.cfg('MODO_FONTE', 'HISTORICO')
        _, t = q.atualizar()
        inat1 = inativados_final(q)
        fin1 = set(q.final_por_id())
        migr = {c[0] for c in conj0 if c[0]}
        e1011 = [x for x in q.ler('tblQA_Integracao') if x['CODIGO'] in ('E10', 'E11')
                 and str(x['ID_REGISTRO'] or '').split('~')[0] in migr]
        comuns = fin0 & fin1
        a0 = {i: p for i, p in inat0.items() if i in comuns}
        a1 = {i: p for i, p in inat1.items() if i in comuns}
        perdidas = sorted(set(a0) - set(a1))[:10]
        novas = sorted(set(a1) - set(a0))[:10]
        plot = [(i, a0[i], a1[i]) for i in set(a0) & set(a1) if a0[i] != a1[i]][:10]
        reg('M06 Inativações REAIS sobrevivem à entrega: a cópia instalada tem o mesmo (ID, REGISTRAR - LJ, DATA, USUARIO) '
            'da produção original; todo ID com resultado ficou com ANALITO; ATUALIZAR DADOS (HISTORICO, consulta nova) dá '
            'os mesmos INATIVADOS com a mesma plotagem; nenhum E10/E11 nas linhas migradas',
            conj1 == conj0 and not sem_an and not perdidas and not novas and not plot and not e1011 and len(conj0) > 0,
            {'linhas_inativar_original': len(conj0), 'saiu': sorted(conj0 - conj1, key=str)[:5],
             'entrou': sorted(conj1 - conj0, key=str)[:5], 'ID_sem_ANALITO': sem_an[:5],
             'inativados_original/depois (IDs nas duas tabelas)': [len(a0), len(a1)],
             'perdidos': perdidas, 'novos': novas, 'plotagem_mudou': plot, 'E10_E11_migradas': e1011[:5],
             'MODO_FONTE_da_copia': modo0, 'atualizacao_s': round(t, 1)})
    except Exception as ex:                                    # noqa: BLE001
        reg('M06 Inativações reais sobrevivem à entrega: execução interrompida', False,
            {'erro': f'{type(ex).__name__}: {ex}', 'traceback': traceback.format_exc()[-1200:]})
    finally:
        if q is not None:
            q.fechar(salvar=False)
        for f in (t_orig, t_inst):
            try:
                os.remove(f)
            except OSError:
                pass


def executar(produto, caminho, saida, original=None):
    hash0 = sha256_arquivo(caminho)
    hash_o = sha256_arquivo(original) if original else None
    fonte = original or caminho
    pasta = os.path.dirname(caminho)
    tmp = os.path.join(pasta, os.path.basename(fonte).replace('.xlsm', '_mig_tmp.xlsm'))
    q = None
    migrar = True
    if original:
        try:
            migrar = not layout_novo(produto, original, pasta)
        except Exception as ex:                                 # noqa: BLE001
            reg('M00 Execução interrompida: não abriu a produção original', False, {'erro': f'{type(ex).__name__}: {ex}'})
            migrar = False
        if not migrar:
            print('   produção original JÁ no layout do ADR-070: M01-M04 não se aplicam (nada a migrar); M06 confere a '
                  'cópia instalada contra ela', flush=True)
    try:
        if not migrar:
            raise StopIteration
        shutil.copy2(fonte, tmp)
        # ------------------------------------------------------------ 1. inativacoes no layout antigo
        q = QA(produto, tmp)
        lo = q.lo(TB)
        cab0 = [c.Name for c in lo.ListColumns]
        faixa0 = float(lo.Parent.Range('A1').MergeArea.Width)        # a barra de navegacao foi montada sobre ela
        if 'MOTIVO' in cab0:
            raise RuntimeError(f'pré-condição ausente: o arquivo já está no layout do ADR-070: {cab0} '
                               f'(na entrega passe --original <produção antes da instalação>)')
        # revisao 07/10/2026 (achado 13): a migracao e o M03 so valem com o ATUALIZAR da consulta nova -- em SEAC o
        # instalador nao o roda. A COPIA vai para HISTORICO (o historico ja recebido, sem rede)
        modo_antes = q.cfg('MODO_FONTE', 'HISTORICO')
        fin0 = q.final_por_id()
        pref = str(q.cfg('PREFIXO_ID') or '').strip().upper()
        cand, vistos = [], set()
        for r in fin0.values():
            if r['ORIGEM_RESULTADO'] == 'INTERFACEAMENTO' and r['STATUS_ANALITICO'] == 'ATIVO' and r['ANALITO'] not in vistos \
                    and str(r['ID_REGISTRO']).startswith(pref + '-'):
                vistos.add(r['ANALITO'])
                cand.append(r)
            if len(cand) == 3:
                break
        A, B, C = cand
        inexistente = f'{pref}-999999998'
        ocupadas = [v[0] for v in lo.ListColumns('ID_REGISTRO').DataBodyRange.Value]
        r0 = next(k for k in range(1, len(ocupadas) - 6) if all(v in (None, '') for v in ocupadas[k - 1:k + 6]))
        pos = {'A': r0, 'B': r0 + 1, 'vazia': r0 + 2, 'C': r0 + 3, 'X': r0 + 4}
        idcol = lo.ListColumns('ID_REGISTRO').DataBodyRange
        print(f'   MODO_FONTE da copia: {modo_antes!r} -> HISTORICO', flush=True)
        idcol.Cells(pos['A'], 1).Value = A['ID_REGISTRO'].split('-')[-1]          # so o numero: o evento normaliza
        idcol.Cells(pos['B'], 1).Value = B['ID_REGISTRO']
        idcol.Cells(pos['C'], 1).Value = C['ID_REGISTRO']
        idcol.Cells(pos['X'], 1).Value = inexistente
        lo.ListColumns('REGISTRAR - LJ').DataBodyRange.Cells(pos['B'], 1).Value = False
        q.escrever_linha('tblComentariosTecnicos', {'ID_REGISTRO': A['ID_REGISTRO'], 'COMENTARIO_TECNICO': 'QA migração: A'})
        q.escrever_linha('tblComentariosTecnicos', {'ID_REGISTRO': B['ID_REGISTRO'], 'COMENTARIO_TECNICO': 'QA migração: B'})
        q.atualizar()
        ids = [A['ID_REGISTRO'], B['ID_REGISTRO'], C['ID_REGISTRO']]
        antes_fin = foto_final(q, ids)
        n_lin = pos['X']
        antes_tab = linhas_tabela(q, n_lin)
        inat_antes = {i: r['TIPO_PLOTAGEM_LJ'] for i, r in q.final_por_id().items() if r['STATUS_ANALITICO'] == 'INATIVADO'}
        pre = (antes_fin[A['ID_REGISTRO']]['STATUS_ANALITICO'] == 'INATIVADO'
               and antes_fin[B['ID_REGISTRO']]['TIPO_PLOTAGEM_LJ'] == 'NAO_PLOTAR'
               and antes_fin[C['ID_REGISTRO']]['GOVERNANCA'] != 'OK')
        if not pre:
            raise RuntimeError(f'pré-condição ausente: inativação no layout antigo não ficou como esperado: {antes_fin}')
        # salva SO a copia temporaria. Eventos DESLIGADOS: com eles, o Workbook_BeforeSave (ADR-065) cancela a
        # gravacao do Excel e grava a propria -- por automacao (COM) ela nao chega ao disco (medido: mtime igual)
        t_antes = os.path.getmtime(tmp)
        q.ex.xl.EnableEvents = False
        q.wb.Save()
        q.fechar(salvar=False)
        if os.path.getmtime(tmp) == t_antes:
            raise RuntimeError('pré-condição ausente: a cópia com as inativações do layout antigo não foi gravada')
        q = None

        # ------------------------------------------------------------ 2. instalador
        ok1, log1 = instalar(produto, tmp, tmp.replace('.xlsm', '_instalar1.log'))
        if not ok1:
            raise RuntimeError('instalar_adr070.py falhou na cópia: ' + log1[-1500:])

        # ------------------------------------------------------------ 3. conferencias
        q = QA(produto, tmp)
        lo = q.lo(TB)
        cab1 = [c.Name for c in lo.ListColumns]
        faixa1 = float(lo.Parent.Range('A1').MergeArea.Width)
        formulas = [c.Name for c in lo.ListColumns if c.DataBodyRange.Cells(1, 1).HasFormula]
        try:
            caixa = int(lo.ListColumns('REGISTRAR - LJ').DataBodyRange.Cells(1, 1).CellControl.Type)
        except Exception:                                      # noqa: BLE001
            caixa = None
        reg('M01 Migração: a tabela fica só com ID | ANALITO | REGISTRAR - LJ | MOTIVO | DATA_INATIVACAO | USUARIO, sem fórmula; '
            'a faixa do título (onde ficam os botões) mantém a largura',
            cab1 == CAB_NOVO and not formulas and abs(faixa1 - faixa0) < 6,
            {'antes': cab0, 'depois': cab1, 'formulas': formulas, 'caixa_nativa_tipo': caixa,
             'largura_faixa_do_titulo': [faixa0, faixa1]})
        depois_tab = linhas_tabela(q, n_lin)
        dif = []
        for k, (a, d) in enumerate(zip(antes_tab, depois_tab), start=1):
            for c in ANTIGAS:
                if a.get(c) != d.get(c) and not (a.get(c) in (None, '') and d.get(c) in (None, '')):
                    dif.append((k, c, a.get(c), d.get(c)))
            if d.get('MOTIVO') not in (None, ''):
                dif.append((k, 'MOTIVO', None, d.get('MOTIVO')))
        an_esp = {pos['A']: A['ANALITO'], pos['B']: B['ANALITO'], pos['C']: C['ANALITO'], pos['X']: None,
                  pos['vazia']: None}
        an_dif = [(k, an_esp[k], depois_tab[k - 1].get('ANALITO')) for k in an_esp
                  if (an_esp[k] or None) != (depois_tab[k - 1].get('ANALITO') or None)]
        vazia_ok = all(v in (None, '', False) for v in depois_tab[pos['vazia'] - 1].values())
        reg('M02 Migração preserva as linhas: mesma posição, ID, caixa, data e usuário idênticos; a linha vazia continua '
            'vazia; ANALITO = o da Principal - Resultados (vazio no ID inexistente); MOTIVO vazio',
            not dif and not an_dif and vazia_ok,
            {'diferencas': dif[:10], 'analito_divergente': an_dif, 'linha_vazia': depois_tab[pos['vazia'] - 1],
             'linhas_depois': depois_tab[r0 - 1:]})
        # a consulta nova RODOU sobre a tabela migrada? (senao o M03 compararia a tabela de antes com ela mesma)
        modo_dep = str(q.cfg('MODO_FONTE') or '')
        rodou = 'ATUALIZAR DADOS em modo historico' in log1 and 'inativado(s), mesmos IDs' in log1
        depois_fin = foto_final(q, ids)
        inat_depois = {i: r['TIPO_PLOTAGEM_LJ'] for i, r in q.final_por_id().items() if r['STATUS_ANALITICO'] == 'INATIVADO'}
        qa = q.ler('tblQA_Integracao')
        e1011 = [x for x in qa if x['CODIGO'] in ('E10', 'E11')]
        e01c = [x for x in qa if x['CODIGO'] == 'E01' and x['ID_REGISTRO'] == C['ID_REGISTRO']]
        e03x = [x for x in qa if x['CODIGO'] == 'E03' and x['ID_REGISTRO'] == inexistente]
        reg('M03 Depois da migração, com a consulta NOVA rodada sobre a tabela migrada (ATUALIZAR em HISTORICO no '
            'instalador), os inativados são os mesmos (mesma plotagem, justificativa pelo comentário antigo, governança, '
            'data e usuário); E01 do sem-comentário e E03 do inexistente continuam; nenhum E10/E11',
            rodou and depois_fin == antes_fin and inat_depois == inat_antes and not e1011 and bool(e01c) and bool(e03x),
            {'consulta_nova_rodou': rodou, 'MODO_FONTE': modo_dep, 'antes': antes_fin, 'depois': depois_fin,
             'inativados': [len(inat_antes), len(inat_depois)],
             'E10_E11': e1011[:5], 'E01_C': len(e01c), 'E03_inexistente': len(e03x)})
        tab1 = tabela_inteira(q)
        q.fechar(salvar=False)
        q = None

        # ------------------------------------------------------------ 4. idempotencia
        ok2, log2 = instalar(produto, tmp, tmp.replace('.xlsm', '_instalar2.log'))
        q = QA(produto, tmp)
        tab2 = tabela_inteira(q)
        reg('M04 Idempotência: o instalador de novo não migra nada e a tabela da Inativar fica idêntica',
            ok2 and 'nada migrado' in log2 and tab1 == tab2,
            {'instalador_ok': ok2, 'nada_migrado': 'nada migrado' in log2, 'tabela_identica': tab1 == tab2,
             'log_fim': log2[-400:]})
    except StopIteration:
        pass
    except Exception as ex_geral:
        reg('M00 Execução interrompida: os testes seguintes a este ponto não rodaram', False,
            {'erro': f'{type(ex_geral).__name__}: {ex_geral}', 'traceback': traceback.format_exc()[-1500:]})
    finally:
        if q is not None:
            q.fechar(salvar=False)
    for f in (tmp,):
        try:
            os.remove(f)
        except OSError:
            pass
    if original and os.path.abspath(original) != os.path.abspath(caminho):
        m06_sobrevivem(produto, caminho, original)
    hash1 = sha256_arquivo(caminho)
    hash_o1 = sha256_arquivo(original) if original else None
    reg('M05 Segurança do teste: arquivos de entrada intactos (SHA-256 igual antes e depois)',
        hash0 == hash1 and hash_o == hash_o1,
        {'arquivo': caminho, 'sha256_antes': hash0, 'sha256_depois': hash1, 'original': original,
         'original_intacto': hash_o == hash_o1})
    if saida:
        with open(saida, 'w', encoding='utf-8') as fh:
            json.dump(RES, fh, ensure_ascii=False, indent=1, default=str)
    falhas = [r for r in RES if r['resultado'] == 'FAIL']
    print(f'\n=== {produto} migração ADR-070: {len(RES) - len(falhas)} PASS / {len(falhas)} FAIL ===', flush=True)
    return falhas


if __name__ == '__main__':
    a = sys.argv[1:]
    orig = os.environ.get('QC_PROD_ORIGINAL') or None
    if '--original' in a:
        i = a.index('--original')
        orig = a[i + 1]
        del a[i:i + 2]
    falhas = executar(a[0], os.path.abspath(a[1]), a[2] if len(a) > 2 else None,
                      os.path.abspath(orig) if orig else None)
    sys.exit(1 if falhas else 0)
