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
     incerteza de medicao e correcoes do CEQ (064/067), ETp ausente = "-" (068), aba Inativar simples e ID na
     dica do LJ (070), selecao de rodadas do CEQ e Sigma com CV pooled de N meses (071), limites por lote com
     rastreabilidade (073), aba Analitos com um bloco por lote e anos de referencia CLIA/VB (074; antes, o resgate
     das digitacoes que nao chegaram a LotesStore -- 078), linhas do alvo que nao dependem das corridas e troca de
     lote sem gravar na LotesStore (078), camada de dados em VBA com Sync_Log (079), eixo X = corrida do lote por nivel (080),
     manual que substitui um resultado do interfaceamento por escolha (SUBSTITUI_ID, 082) e blindagem. ANTES de todos,
     o relatorio da mudanca de veredicto de Westgard (083: regras do mesmo nivel na sequencia de ativos), que precisa
     do VBA antigo da producao.
     Todos idempotentes; cada um instala o VBA inteiro das fontes (codigo_atual.py);
  3. guarda os bytes instalados e roda, cada suite numa copia propria deles:
     qa_seguranca, qa_casos_extremos, qa_troca_lote, qa_lotes_auto, qa_graficos (Excel VISIVEL),
     qa_incerteza (recalculo independente), qa_desempenho, qa_final, qa_etl_recebimento, qa_etl, qa_rodadas_cv,
     qa_inativar, qa_migracao_adr070, qa_lotes_limites, qa_blocos_lote, qa_varredura_lotes, qa_linhas_alvo, qa_sync, qa_corrida_lote,
     qa_inativar_ux, qa_manual_substitui, qa_westgard_sequencia. A producao ORIGINAL (copiada para <produto>/producao_original/ antes de
     instalar, conferida por SHA-256) vai para as suites pela variavel QC_PROD_ORIGINAL: o R00 do qa_rodadas_cv
     compara a copia instalada com ela (regressao do gate) e o qa_migracao_adr070 testa a migracao sobre ela e
     confere que as inativacoes reais sobrevivem (M06) -- nenhum caminho de rede no codigo;
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
# QC_RAIZ: pasta dos .xlsm de producao quando a entrega roda de uma COPIA LIMPA do codigo commitado (21/10/2026: a
# entrega pegou do src/ alteracoes em andamento de outro ADR e falhou -- a producao nao foi tocada)
RAIZ = os.environ.get('QC_RAIZ') or os.path.abspath(os.path.join(AQUI, '..', '..'))
ARQ = {'Hematologia': 'QC_Hematologia.xlsm', 'Bioquimica': 'QC_Bioquimica.xlsm'}
SUITES = ['qa_seguranca.py', 'qa_casos_extremos.py', 'qa_troca_lote.py', 'qa_lotes_auto.py',
          'qa_graficos.py', 'qa_incerteza.py', 'qa_desempenho.py', 'qa_final.py',
          'qa_etl_recebimento.py', 'qa_etl.py',          # ADR-066: QA-ETL-001, gate obrigatorio da camada de dados
          'qa_rodadas_cv.py',                            # ADR-071: selecao de rodadas do CEQ e Sigma com CV pooled
          'qa_inativar.py',                              # ADR-070: Inativar por ID + analito, ID na dica do LJ
          'qa_migracao_adr070.py',                       # ADR-070: migracao e inativacoes reais contra a producao
          'qa_lotes_limites.py',                         # ADR-073: limites por lote + rastreabilidade
          'qa_blocos_lote.py',                           # ADR-074: aba Analitos com um bloco por lote
          'qa_varredura_lotes.py',                       # ADR-077/078: todos os lotes no Painel com media e limites
          'qa_linhas_alvo.py',                           # ADR-078: linhas do alvo antes da 1a corrida, troca sem gravar
          'qa_sync.py',                                  # ADR-079: incremental, idempotente, falhas, reserva PQ
          'qa_corrida_lote.py',                          # ADR-080/081: corrida do lote (so ativos), X entre corridas
          'qa_inativar_ux.py',                           # ADR-081: inativar/reativar pelo grafico, simples
          'qa_manual_substitui.py',                      # ADR-082: manual substitui o automatico (SUBSTITUI_ID)
          'qa_westgard_sequencia.py',                    # ADR-083: Westgard do mesmo nivel na sequencia de ativos
          'qa_adr084.py',                                # ADR-084: status, motivo livre, indicador, calibracao pelo ID
          'qa_usuarios_adr084.py',                       # ADR-084: cadastro central, PBKDF2, login e assinatura
          'qa_adr085.py']                                # ADR-085: painel tela unica (acoes no ponto, comentario, Westgard)


def sha(p):
    h = hashlib.sha256()
    with open(p, 'rb') as f:
        for b in iter(lambda: f.read(1 << 20), b''):
            h.update(b)
    return h.hexdigest()


def rodar(log, cwd, *args, original=None):
    env = dict(os.environ, PYTHONIOENCODING='utf-8')
    if original:
        env['QC_PROD_ORIGINAL'] = original
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
    # a producao como estava ANTES de instalar: referencia do R00 (qa_rodadas_cv) e do M06 (qa_migracao_adr070)
    original = os.path.join(trab, 'producao_original', ARQ[produto])
    os.makedirs(os.path.dirname(original), exist_ok=True)
    shutil.copy2(prod, original)
    if sha(original) != sha_inicio:
        return False, 'a producao mudou enquanto era copiada: rodar de novo com o arquivo fechado'
    print(f'\n## {produto}\n   producao sha256 {sha_inicio[:16]} -> copia de trabalho', flush=True)

    # cada instalador instala TODO o VBA atual (codigo_atual.py) e cuida da sua parte de planilha
    # ADR-083 PRIMEIRO: o relatorio da mudanca de veredicto de Westgard precisa do VBA ANTIGO da producao (qualquer
    # outro instalador ja levaria o mEstatistica novo pelo codigo_atual e a troca sairia sem relatorio)
    passos = [('westgard_seq', ['instalar_adr083.py', produto, alvo]),
              ('modo_fonte', ['instalar_modo_historico.py', produto, alvo, '--modo', modo]
               + (['--caminho', origem] if origem else [])),
              ('lotes', ['instalar_adr060.py', produto, alvo, '--registrar']),
              ('seguranca', ['instalar_seguranca_usuarios.py', produto, alvo]),
              ('graficos', ['instalar_adr063.py', produto, alvo]),
              ('desempenho', ['instalar_adr062.py', produto, alvo]),
              ('incerteza', ['instalar_adr064.py', produto, alvo]),
              ('espec_etp', ['instalar_adr069.py', produto, alvo]),
              ('sem_etp', ['instalar_adr068.py', produto, alvo]),
              ('inativar', ['instalar_adr070.py', produto, alvo]),       # ADR-070: Inativar simples + ID no LJ
              ('rodadas_cv', ['instalar_adr071.py', produto, alvo]),     # ADR-071: rodadas do CEQ + Sigma CV pooled
              ('limites_lote', ['instalar_adr073.py', produto, alvo]),   # ADR-073: central de limites por lote
              ('espec_hema', ['instalar_adr075.py', produto, alvo]),     # ADR-075: Hematologia no modelo da Bioquimica
              ('blocos_lote', ['instalar_adr074.py', produto, alvo]),    # ADR-074: um bloco por lote na aba Analitos
              ('linhas_alvo', ['instalar_adr078.py', produto, alvo]),    # ADR-078: linhas do alvo sem depender das corridas
              ('motor_dados', ['instalar_adr079.py', produto, alvo]),    # ADR-079: camada de dados em VBA + Sync_Log
              ('corrida_lote', ['instalar_adr080.py', produto, alvo]),   # ADR-080/081: eixo X = corrida do lote (ativos)
              ('inativar_ux', ['instalar_frm_inativar.py', produto, alvo]),  # ADR-081: inativar/reativar pelo grafico
              ('manual_subst', ['instalar_adr082.py', produto, alvo]),   # ADR-082: SUBSTITUI_ID (manual no lugar do automatico)
              ('adr084', ['instalar_adr084.py', produto, alvo]),         # ADR-084: calibracao pelo ID, indicador, usuarios central
              ('adr085', ['instalar_adr085.py', produto, alvo]),         # ADR-085: painel tela unica
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
                        suite, produto, copia, os.path.join(trab, suite.replace('.py', '.json')), original=original)
        print(f'   {suite:22s} {"OK" if ok else "FALHOU"}  {" | ".join(fim)[:300]}', flush=True)
        if not ok:
            return False, f'{suite} falhou (ver {trab})'
    if sha(alvo) != instalado:
        return False, 'o arquivo instalado mudou durante os testes'
    if sha(original) != sha_inicio:
        return False, 'a copia da producao original mudou durante os testes (uma suite a alterou)'

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
    try:
        rel = relatorio_westgard(trab, prod)
    except OSError as e:                     # a producao ja foi trocada: o relatorio continua em trab
        rel = [f'NAO copiado ({e}); ver {trab}']
    return True, f'ENTREGUE sha256 {instalado[:16]} (backup em {bk})' + (f'; Westgard: {", ".join(rel)}' if rel else '')


def relatorio_westgard(trab, prod):
    """ADR-083: o relatorio da mudanca de veredicto de Westgard (instalar_adr083.py, ao lado da copia de trabalho) vai
    tambem para o lado do arquivo de producao -- e a lista para o laboratorio decidir sobre a revisao de laudos. Um que
    ja esteja la nunca e sobrescrito: o mesmo conteudo nao e copiado de novo; conteudo diferente ganha data/hora no nome,
    e a mensagem diz o nome que o Audit_Log cita e o nome final (os dois aparecem no resumo da entrega). Previa
    (_PREVIA, --sem-salvar) nao vai."""
    base = os.path.splitext(os.path.basename(prod))[0] + '_ADR083_mudancas_westgard'
    out = []
    for nome in sorted(os.listdir(trab)):
        if nome.startswith(base) and nome.endswith('.csv') and not nome.endswith('_PREVIA.csv'):
            orig = os.path.join(trab, nome)
            dest = os.path.join(os.path.dirname(prod), nome)
            if os.path.exists(dest):
                if sha(dest) == sha(orig):
                    out.append(f'{dest} (ja estava)')
                    continue
                dest = dest[:-4] + '_' + time.strftime('%Y%m%d_%H%M%S') + '.csv'
                shutil.copy2(orig, dest)
                out.append(f'{dest} (o Audit_Log cita {nome}; ja havia outro com esse nome ao lado da producao)')
                continue
            shutil.copy2(orig, dest)
            out.append(dest)
    return out


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
