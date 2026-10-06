# -*- coding: utf-8 -*-
"""QA da MIGRACAO do ADR-070 -- roda sobre um arquivo AINDA NO LAYOUT ANTIGO (antes do instalar_adr070.py).

Uso: python qa_migracao_adr070.py <Bioquimica|Hematologia> <arquivo_pre_adr070.xlsm> [saida.json]

O arquivo recebido NAO e alterado (SHA-256 conferido): tudo acontece em <nome>_mig_tmp.xlsm, apagada no fim.
  1. na copia, com o VBA e as consultas ANTIGAS, inativa pelo caminho real (evento ligado) como o laboratorio fez ate
     hoje: A (so o numero, com comentario tecnico), B (com comentario e caixa DESMARCADA), uma linha VAZIA no meio,
     C (sem comentario: E01), um ID inexistente (E03); ATUALIZAR DADOS; foto da tabela e da tblCQ_Final; salva a copia;
  2. roda instalar_adr070.py na copia (subprocesso, como o entregar.py);
  3. M01 layout novo sem formula; M02 cada linha antiga na MESMA posicao com ID, caixa, data e usuario identicos; a linha
     vazia continua vazia; ANALITO das linhas antigas = o da tblCQ_Final (vazio no ID inexistente); MOTIVO vazio;
     M03 tblCQ_Final: os mesmos INATIVADOS com a mesma plotagem, justificativa, governanca, data e usuario da
     inativacao; QA: E01 do C e E03 do inexistente continuam, nenhum E10/E11;
  4. M04 idempotencia: o instalador de novo nao migra nada e a tabela fica identica.
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


def executar(produto, caminho, saida):
    hash0 = sha256_arquivo(caminho)
    tmp = caminho.replace('.xlsm', '_mig_tmp.xlsm')
    shutil.copy2(caminho, tmp)
    q = None
    try:
        # ------------------------------------------------------------ 1. inativacoes no layout antigo
        q = QA(produto, tmp)
        lo = q.lo(TB)
        cab0 = [c.Name for c in lo.ListColumns]
        if 'MOTIVO' in cab0:
            raise RuntimeError(f'pré-condição ausente: o arquivo já está no layout do ADR-070: {cab0}')
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
        q.fechar(salvar=True)                                  # salva SO a copia temporaria
        q = None

        # ------------------------------------------------------------ 2. instalador
        ok1, log1 = instalar(produto, tmp, tmp.replace('.xlsm', '_instalar1.log'))
        if not ok1:
            raise RuntimeError('instalar_adr070.py falhou na cópia: ' + log1[-1500:])

        # ------------------------------------------------------------ 3. conferencias
        q = QA(produto, tmp)
        lo = q.lo(TB)
        cab1 = [c.Name for c in lo.ListColumns]
        formulas = [c.Name for c in lo.ListColumns if c.DataBodyRange.Cells(1, 1).HasFormula]
        try:
            caixa = int(lo.ListColumns('REGISTRAR - LJ').DataBodyRange.Cells(1, 1).CellControl.Type)
        except Exception:                                      # noqa: BLE001
            caixa = None
        reg('M01 Migração: a tabela fica só com ID | ANALITO | REGISTRAR - LJ | MOTIVO | DATA_INATIVACAO | USUARIO, sem fórmula',
            cab1 == CAB_NOVO and not formulas,
            {'antes': cab0, 'depois': cab1, 'formulas': formulas, 'caixa_nativa_tipo': caixa})
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
        depois_fin = foto_final(q, ids)
        inat_depois = {i: r['TIPO_PLOTAGEM_LJ'] for i, r in q.final_por_id().items() if r['STATUS_ANALITICO'] == 'INATIVADO'}
        qa = q.ler('tblQA_Integracao')
        e1011 = [x for x in qa if x['CODIGO'] in ('E10', 'E11')]
        e01c = [x for x in qa if x['CODIGO'] == 'E01' and x['ID_REGISTRO'] == C['ID_REGISTRO']]
        e03x = [x for x in qa if x['CODIGO'] == 'E03' and x['ID_REGISTRO'] == inexistente]
        reg('M03 Depois da migração os inativados são os mesmos (mesma plotagem, justificativa pelo comentário antigo, '
            'governança, data e usuário); E01 do sem-comentário e E03 do inexistente continuam; nenhum E10/E11',
            depois_fin == antes_fin and inat_depois == inat_antes and not e1011 and bool(e01c) and bool(e03x),
            {'antes': antes_fin, 'depois': depois_fin, 'inativados': [len(inat_antes), len(inat_depois)],
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
    hash1 = sha256_arquivo(caminho)
    reg('M05 Segurança do teste: arquivo de entrada intacto (SHA-256 igual antes e depois)', hash0 == hash1,
        {'arquivo': caminho, 'sha256_antes': hash0, 'sha256_depois': hash1})
    if saida:
        with open(saida, 'w', encoding='utf-8') as fh:
            json.dump(RES, fh, ensure_ascii=False, indent=1, default=str)
    falhas = [r for r in RES if r['resultado'] == 'FAIL']
    print(f'\n=== {produto} migração ADR-070: {len(RES) - len(falhas)} PASS / {len(falhas)} FAIL ===', flush=True)
    return falhas


if __name__ == '__main__':
    falhas = executar(sys.argv[1], os.path.abspath(sys.argv[2]), sys.argv[3] if len(sys.argv) > 3 else None)
    sys.exit(1 if falhas else 0)
