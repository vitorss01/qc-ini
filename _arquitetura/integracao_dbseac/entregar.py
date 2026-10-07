# -*- coding: utf-8 -*-
"""entregar.py -- instala, TESTA e so entao entrega os .xlsm de producao (ADR-057/058/059).

Uso:  python entregar.py [Hematologia] [Bioquimica] [--modo SEAC|HISTORICO] [--caminho <DB_SEAC.xlsm>] [--sem-testes]

  --caminho ......... grava CAMINHO_DB_SEAC nos QCs (o DB_SEAC do laboratorio).

  --modo SEAC ....... rede do laboratorio (DB_SEAC acessivel). Padrao.
  --modo HISTORICO .. fora da rede: reprocessa so o historico ja recebido (ADR-058).
  --sem-testes ...... instala e blinda sem rodar as suites (NAO entrega na producao).

Para cada produto:
  1. copia o .xlsm de producao para _entrega_<data>/<produto>/ (a producao nao e tocada);
  2. na copia: MODO_FONTE (058), lotes automaticos e validade (060, registra os lotes ja recebidos),
     papeis e sessao (059), graficos: zoom (061) e todos a vista (063), troca de analito rapida (062),
     incerteza de medicao e correcoes do CEQ (064/067), ETp ausente = "-" (068), selecao de rodadas do CEQ e
     Sigma com CV pooled de N meses (071) e blindagem.
     Todos idempotentes; cada um instala o VBA inteiro das fontes (codigo_atual.py);
  3. guarda os bytes instalados e roda, cada suite numa copia propria deles:
     qa_seguranca, qa_casos_extremos, qa_troca_lote, qa_lotes_auto, qa_graficos (Excel VISIVEL),
     qa_incerteza (recalculo independente), qa_desempenho, qa_final, qa_etl_recebimento, qa_etl, qa_rodadas_cv;
  4. SO SE TODAS PASSAREM E A PRODUCAO NAO TIVER SIDO GRAVADA DURANTE OS TESTES: backup da producao em _backup_pre_integracao_<data>/ conferido por
     SHA-256, e troca pelos MESMOS bytes que passaram nos testes (SHA-256 conferido de novo).
Falhou qualquer passo: a producao fica como estava, e o relatorio diz onde parou.

Leva ~25 min por produto (qa_final faz varios ATUALIZAR DADOS). Excel COM precisa de uma
instancia de pe nesta maquina -- ver o comentario de Novo-Excel em scripts_fase3/blindar_artefato.ps1.
"""
import hashlib
import os
import shutil
import subprocess
import sys
import time

AQUI = os.path.dirname(os.path.abspath(__file__))
RAIZ = os.path.abspath(os.path.join(AQUI, '..', '..'))
ARQ = {'Hematologia': 'QC_Hematologia.xlsm', 'Bioquimica': 'QC_Bioquimica.xlsm'}
SUITES = ['qa_seguranca.py', 'qa_casos_extremos.py', 'qa_troca_lote.py', 'qa_lotes_auto.py',
          'qa_graficos.py', 'qa_incerteza.py', 'qa_desempenho.py', 'qa_final.py',
          'qa_etl_recebimento.py', 'qa_etl.py',          # ADR-066: QA-ETL-001, gate obrigatorio da camada de dados
          'qa_rodadas_cv.py']                            # ADR-071: selecao de rodadas do CEQ e Sigma com CV pooled


def sha(p):
    h = hashlib.sha256()
    with open(p, 'rb') as f:
        for b in iter(lambda: f.read(1 << 20), b''):
            h.update(b)
    return h.hexdigest()


def rodar(log, cwd, *args):
    env = dict(os.environ, PYTHONIOENCODING='utf-8')
    with open(log, 'w', encoding='utf-8') as f:
        r = subprocess.run([sys.executable, '-u', '-W', 'ignore', *args], cwd=cwd, env=env,
                           stdout=f, stderr=subprocess.STDOUT, text=True, encoding='utf-8', errors='replace')
    with open(log, encoding='utf-8', errors='replace') as f:
        fim = [l.rstrip() for l in f if l.startswith(('===', 'FAIL', 'INCONCLUSIVO', 'Traceback', 'VBA NAO', '   XML'))]
    return r.returncode == 0, fim


def entregar(produto, modo, testar, pasta, origem=None):
    prod = os.path.join(RAIZ, ARQ[produto])
    trab = os.path.join(pasta, produto)
    os.makedirs(trab, exist_ok=True)
    alvo = os.path.join(trab, ARQ[produto])
    shutil.copy2(prod, alvo)
    sha_inicio = sha(alvo)
    print(f'\n## {produto}\n   producao sha256 {sha_inicio[:16]} -> copia de trabalho', flush=True)

    # cada instalador instala TODO o VBA atual (codigo_atual.py) e cuida da sua parte de planilha
    passos = [('modo_fonte', ['instalar_modo_historico.py', produto, alvo, '--modo', modo]
               + (['--caminho', origem] if origem else [])),
              ('lotes', ['instalar_adr060.py', produto, alvo, '--registrar']),
              ('seguranca', ['instalar_seguranca_usuarios.py', produto, alvo]),
              ('graficos', ['instalar_adr063.py', produto, alvo]),
              ('desempenho', ['instalar_adr062.py', produto, alvo]),
              ('incerteza', ['instalar_adr064.py', produto, alvo]),
              ('espec_etp', ['instalar_adr069.py', produto, alvo]),
              ('sem_etp', ['instalar_adr068.py', produto, alvo]),
              ('rodadas_cv', ['instalar_adr071.py', produto, alvo]),     # ADR-071: rodadas do CEQ + Sigma CV pooled
              ('blindagem', ['blindar_entrega.py', alvo])]
    for nome, args in passos:
        ok, fim = rodar(os.path.join(trab, f'{nome}.log'), AQUI, *args)
        print(f'   {nome:12s} {"OK" if ok else "FALHOU"}  {" | ".join(fim)[:200]}', flush=True)
        if not ok:
            return False, f'{nome} falhou (ver {trab})'
    instalado = sha(alvo)
    print(f'   instalado sha256 {instalado[:16]}', flush=True)
    if not testar:
        return False, 'instalado e blindado SEM testes: producao nao foi tocada'

    for suite in SUITES:
        copia = os.path.join(trab, 'teste_' + suite.replace('.py', ''), ARQ[produto])
        os.makedirs(os.path.dirname(copia), exist_ok=True)
        shutil.copy2(alvo, copia)
        ok, fim = rodar(os.path.join(trab, suite.replace('.py', '.log')), os.path.join(AQUI, 'testes'),
                        suite, produto, copia, os.path.join(trab, suite.replace('.py', '.json')))
        print(f'   {suite:22s} {"OK" if ok else "FALHOU"}  {" | ".join(fim)[:300]}', flush=True)
        if not ok:
            return False, f'{suite} falhou (ver {trab})'
    if sha(alvo) != instalado:
        return False, 'o arquivo instalado mudou durante os testes'

    # a producao foi gravada no laboratorio durante as horas de teste: trocar agora apagaria o que foi lancado
    # (inativacoes, manuais, comentarios). Nao troca; rodar de novo com o arquivo fechado (06/10/2026)
    if sha(prod) != sha_inicio:
        return False, ('a producao foi gravada durante os testes (alguem salvou o arquivo): NAO trocada. '
                       'Rodar de novo com o arquivo fechado')
    bk = os.path.join(RAIZ, '_backup_pre_integracao_' + time.strftime('%Y-%m-%d_%H%M'))
    os.makedirs(bk, exist_ok=True)
    shutil.copy2(prod, os.path.join(bk, ARQ[produto]))
    if sha(os.path.join(bk, ARQ[produto])) != sha(prod):
        return False, 'backup divergente: producao NAO foi trocada'
    shutil.copy2(alvo, prod)
    if sha(prod) != instalado:
        shutil.copy2(os.path.join(bk, ARQ[produto]), prod)
        return False, 'copia para a producao divergiu: producao devolvida do backup'
    return True, f'ENTREGUE sha256 {instalado[:16]} (backup em {bk})'


def main(argv):
    modo = argv[argv.index('--modo') + 1].upper() if '--modo' in argv else 'SEAC'
    produtos = [a for a in argv if a in ARQ] or ['Hematologia', 'Bioquimica']
    pasta = os.path.join(RAIZ, '_entrega_' + time.strftime('%Y-%m-%d_%H%M'))
    origem = argv[argv.index('--caminho') + 1] if '--caminho' in argv else None
    res = {p: entregar(p, modo, '--sem-testes' not in argv, pasta, origem) for p in produtos}
    print(f'\nMODO_FONTE = {modo}. Logs e evidencias em {pasta}')
    for p, (ok, msg) in res.items():
        print(f'  {p}: {msg}')
    return 0 if all(ok for ok, _ in res.values()) else 1


if __name__ == '__main__':
    sys.exit(main(sys.argv[1:]))
