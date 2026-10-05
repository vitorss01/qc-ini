# -*- coding: utf-8 -*-
"""alterar_db_seac.py -- DB_SEAC v2.4: preserva a chave de origem do SIPEC.

O CSV do SIPEC (relat207 / relat209) traz em cada linha um identificador unico
do resultado (coluna 'id'), alem de idAmostra, item_Id e Unidade. As consultas
DB_HEMATOLOGIA e DB_BIOQUIMICA descartavam esses campos -- e sem o 'id' nao ha
chave estavel: no proprio arquivo, 527 grupos da Hematologia (retransmissoes) e
1.566 da Bioquimica (replicas no mesmo segundo) repetem a chave natural
equipamento+lote+nivel+data/hora+analito.

O que este instalador faz, e so isso:
  1. acrescenta ID_ORIGEM, ID_AMOSTRA, ITEM_ID e UNIDADE NO FIM das duas
     tabelas -- nenhuma coluna existente muda de nome, posicao ou tipo, entao as
     tabelas dinamicas, as segmentacoes e VerificarInterfaceamento (que acham as
     colunas pelo NOME) continuam iguais;
  2. Bioquimica: ID_ORIGEM entra como desempate na ordenacao do PRIMEIRO_DO_DIA,
     que antes era arbitraria entre resultados do mesmo segundo;
  3. registra a mudanca como ALTERACAO APROVADA, do jeito que o proprio arquivo
     exige (VALIDACAO secao 5 + assinatura regravada + trilha de auditoria),
     sobe VERSAO para 2.4 e roda a rotina oficial ATUALIZAR.

Uso:  python alterar_db_seac.py "<caminho do DB_SEAC.xlsm>"
"""
import datetime
import getpass
import math
import os
import socket
import sys

AQUI = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(AQUI, '..', 'etapa1_multilote'))
import xlh  # noqa: E402
import pywintypes  # noqa: E402

# Senha de administracao do DB_SEAC (constante SENHA de mdlControle, documentada na aba
# VALIDACAO do proprio arquivo). NUNCA em texto aqui: o repositorio e publico. Vem da
# variavel de ambiente DB_SEAC_SENHA ou e pedida no terminal.
SENHA = os.environ.get('DB_SEAC_SENHA') or getpass.getpass('Senha de administracao do DB_SEAC: ')
VERSAO_NOVA = '2.4'
CAMPOS_NOVOS = ['ID_ORIGEM', 'ID_AMOSTRA', 'ITEM_ID', 'UNIDADE']


def assinatura(s):
    """Porta fiel de mdlControle.Assinatura (checksum de deteccao, nao criptografico)."""
    s = s.replace('\r', '').replace('\n', '').replace(' ', '')
    n = len(s)
    h = 0.0
    for ch in s:
        c = ord(ch)
        if c > 32767:
            c -= 65536                   # AscW devolve Integer com sinal
        h = h * 31 + c
        if h > 2147483647.0:
            h = h - math.floor(h / 2147483647.0) * 2147483647.0
    return '%05d-%08X' % (n, int(h) & 0xFFFFFFFF)


def troca(txt, velho, novo, rotulo):
    k = txt.count(velho)
    if k != 1:
        raise RuntimeError(f'{rotulo}: esperado 1 trecho, achados {k}. A consulta mudou desde a auditoria -- nada foi gravado.')
    return txt.replace(velho, novo)


def patch_hema(m):
    m = troca(m,
              'Renomeado = Table.RenameColumns(NaJanela, {{"Column2", "EQUIPAMENTO"}, {"Column9", "ANALITO"}}),',
              'Renomeado = Table.RenameColumns(NaJanela, {{"Column2", "EQUIPAMENTO"}, {"Column9", "ANALITO"},\n'
              '        {"Column1", "ID_ORIGEM"}, {"Column3", "ID_AMOSTRA"}, {"Column4", "ITEM_ID"}, {"Column6", "UNIDADE"}}),',
              'hema/rename')
    m = troca(m,
              '{"EQUIPAMENTO", "LOTE", "NIVEL", "NIVEL_DESC", "DATA", "DATA_HORA", "ANALITO", "VALOR", "FLAG"}),\n\n    Tipos',
              '{"EQUIPAMENTO", "LOTE", "NIVEL", "NIVEL_DESC", "DATA", "DATA_HORA", "ANALITO", "VALOR", "FLAG",\n'
              '         // v2.4: chave de origem do SIPEC e rastreabilidade, sempre NO FIM\n'
              '         "ID_ORIGEM", "ID_AMOSTRA", "ITEM_ID", "UNIDADE"}),\n\n    Tipos',
              'hema/saida')
    m = troca(m,
              '{"VALOR", type nullable number}, {"FLAG", type nullable text}})\nin',
              '{"VALOR", type nullable number}, {"FLAG", type nullable text},\n'
              '        {"ID_ORIGEM", Int64.Type}, {"ID_AMOSTRA", type text}, {"ITEM_ID", Int64.Type}, {"UNIDADE", type nullable text}})\nin',
              'hema/tipos')
    return m


def patch_bioq(m):
    m = troca(m,
              'Renomeado = Table.RenameColumns(NaJanela, {{"Column9", "ANALITO"}}),',
              'Renomeado = Table.TransformColumnTypes(\n'
              '        Table.RenameColumns(NaJanela, {{"Column9", "ANALITO"},\n'
              '            {"Column1", "ID_ORIGEM"}, {"Column3", "ID_AMOSTRA"}, {"Column4", "ITEM_ID"}, {"Column6", "UNIDADE"}}),\n'
              '        {{"ID_ORIGEM", Int64.Type}, {"ITEM_ID", Int64.Type}}),',
              'bioq/rename')
    m = troca(m,
              '{{"K", Order.Ascending}, {"DATA_HORA", Order.Ascending}}',
              '{{"K", Order.Ascending}, {"DATA_HORA", Order.Ascending}, {"ID_ORIGEM", Order.Ascending}}',
              'bioq/ordenacao')
    m = troca(m,
              '"DATA", "DATA_HORA", "ANALITO", "PRIMEIRO_DO_DIA", "VALOR", "FLAG"}),',
              '"DATA", "DATA_HORA", "ANALITO", "PRIMEIRO_DO_DIA", "VALOR", "FLAG",\n'
              '         // v2.4: chave de origem do SIPEC e rastreabilidade, sempre NO FIM\n'
              '         "ID_ORIGEM", "ID_AMOSTRA", "ITEM_ID", "UNIDADE"}),',
              'bioq/saida')
    m = troca(m,
              '{"FLAG", type nullable text}})\nin',
              '{"FLAG", type nullable text},\n'
              '        {"ID_ORIGEM", Int64.Type}, {"ID_AMOSTRA", type text}, {"ITEM_ID", Int64.Type}, {"UNIDADE", type nullable text}})\nin',
              'bioq/tipos')
    return m


def main(caminho):
    caminho = os.path.abspath(caminho)
    xl = xlh.Excel(visivel=False, eventos=False)
    try:
        wb = xl.abrir(caminho)
        q = {qq.Name: qq for qq in wb.Queries}

        # ---- 0. sanidade: minha porta da assinatura confere com as referencias intocadas
        ctrl = wb.Worksheets('CONTROLE')
        refs = {}
        r = 3
        while ctrl.Cells(r, 1).Value:
            refs[ctrl.Cells(r, 1).Value] = (r, ctrl.Cells(r, 2).Value)
            r += 1
        for nome in ('CFG_MESES_HEMA', 'CFG_BYTES_HEMA', 'CFG_CORTE_HEMA', 'DB_HEMATOLOGIA'):
            calc = assinatura(q[nome].Formula)
            if calc != refs[nome][1]:
                raise RuntimeError(f'porta da Assinatura divergente em {nome}: {calc} x {refs[nome][1]}')
        print('assinatura: porta conferida contra 4 referencias intocadas')
        sit_antes = {n: ('CONFORME' if assinatura(q[n].Formula) == refs[n][1] else 'ALTERADA') for n in refs}
        print('integridade ANTES:', sit_antes)

        # ---- 1. novas consultas
        mh0, mb0 = q['DB_HEMATOLOGIA'].Formula, q['DB_BIOQUIMICA'].Formula
        if 'ID_ORIGEM' in mh0 or 'ID_ORIGEM' in mb0:
            raise RuntimeError('o arquivo ja tem ID_ORIGEM -- instalador ja aplicado')
        mh, mb = patch_hema(mh0), patch_bioq(mb0)
        os.makedirs(os.path.join(AQUI, 'db_seac'), exist_ok=True)
        for nome, txt in (('DB_HEMATOLOGIA_v2.3_original.m', mh0), ('DB_BIOQUIMICA_v2.3_original.m', mb0),
                          ('DB_HEMATOLOGIA.m', mh), ('DB_BIOQUIMICA.m', mb)):
            with open(os.path.join(AQUI, 'db_seac', nome), 'w', encoding='utf-8', newline='\n') as f:
                f.write(txt)

        wb.Unprotect(SENHA)
        for sh in ('CONTROLE', 'LOG_AUDITORIA', 'VALIDACAO', 'INICIO'):
            wb.Worksheets(sh).Unprotect(SENHA)
        q['DB_HEMATOLOGIA'].Formula = mh
        q['DB_BIOQUIMICA'].Formula = mb

        # ---- 2. VERSAO no codigo
        cm = wb.VBProject.VBComponents('mdlControle').CodeModule
        achou = False
        for i in range(1, cm.CountOfLines + 1):
            ln = cm.Lines(i, 1)
            if ln.strip().startswith('Private Const VERSAO'):
                cm.ReplaceLine(i, ln.replace('"2.3"', '"%s"' % VERSAO_NOVA))
                achou = '"%s"' % VERSAO_NOVA in cm.Lines(i, 1)
                break
        if not achou:
            raise RuntimeError('constante VERSAO nao encontrada/alterada em mdlControle')

        # ---- 3. historico de versoes (VALIDACAO secao 5)
        val = wb.Worksheets('VALIDACAO')
        ult = None
        for rr in range(1, val.UsedRange.Rows.Count + val.UsedRange.Row + 1):
            v = val.Cells(rr, 1).Value
            if v == 2.3 or str(v) == '2.3':
                ult = rr
        if ult is None:
            raise RuntimeError('linha da versao 2.3 nao encontrada na VALIDACAO')
        val.Rows(ult + 1).Insert()
        val.Rows(ult).Copy(val.Rows(ult + 1))
        hoje = datetime.date.today()
        val.Cells(ult + 1, 1).Value = 2.4
        val.Cells(ult + 1, 2).Value = pywintypes.Time(datetime.datetime(hoje.year, hoje.month, hoje.day))
        val.Cells(ult + 1, 2).NumberFormat = 'dd/mm/yyyy'
        val.Cells(ult + 1, 3).Value = 'Vítor Santos da Silva'
        val.Cells(ult + 1, 5).Value = (
            'Chave de origem preservada: DB_HEMATOLOGIA e DB_BIOQUIMICA passam a levar, NO FIM das tabelas, '
            'ID_ORIGEM (coluna id do CSV do SIPEC, unica por resultado), ID_AMOSTRA, ITEM_ID e UNIDADE. Nenhuma '
            'coluna existente mudou de nome, posicao ou tipo. Na bioquimica, ID_ORIGEM desempata a ordenacao do '
            'PRIMEIRO_DO_DIA entre resultados do mesmo segundo. Motivo: integracao com o QC_INI (DB_CQ_FINAL), que '
            'exige ID estavel para inativacao e rastreabilidade.')
        val.Cells(ult + 1, 6).Value = 'SIM (gestor, 01/10/2026)'
        val.Cells(ult + 1, 7).Value = 'SIM'

        # ---- 4. assinaturas regravadas
        user = getpass.getuser()
        for nome, (rr, _) in refs.items():
            ctrl.Cells(rr, 2).Value = assinatura(q[nome].Formula)
            ctrl.Cells(rr, 3).Value = pywintypes.Time(datetime.datetime.now().replace(microsecond=0))
            ctrl.Cells(rr, 3).NumberFormat = 'dd/mm/yyyy hh:mm:ss'
            ctrl.Cells(rr, 4).Value = user

        # ---- 5. trilha (mesmo layout de mdlControle.RegistrarLog)
        log = wb.Worksheets('LOG_AUDITORIA')
        r = log.Cells(log.Rows.Count, 1).End(-4162).Row + 1
        log.Cells(r, 1).Value = pywintypes.Time(datetime.datetime.now().replace(microsecond=0))
        log.Cells(r, 1).NumberFormat = 'dd/mm/yyyy hh:mm:ss'
        log.Cells(r, 2).Value = user
        log.Cells(r, 3).Value = socket.gethostname()
        log.Cells(r, 4).Value = 'REGRAVACAO DE ASSINATURAS'
        log.Cells(r, 9).Value = 'REDEFINIDA'
        log.Cells(r, 10).NumberFormat = '@'
        log.Cells(r, 10).Value = VERSAO_NOVA
        divergentes = ', '.join(f'{k}={v}' for k, v in sit_antes.items() if v != 'CONFORME') or 'todas CONFORME'
        log.Cells(r, 11).Value = ('alteracao aprovada v2.4 (VALIDACAO secao 5): ID_ORIGEM/ID_AMOSTRA/ITEM_ID/UNIDADE '
                                  'no fim de tbHematologia e tbBioquimica; integridade antes da alteracao: ' + divergentes)

        # ---- 6. rotina oficial de atualizacao (retranca tudo no fim)
        print('ATUALIZAR (SIPEC)...')
        xl.run('ATUALIZAR', teto=900)

        # ---- 7. verificacao
        out = {}
        originais = {
            'tbHematologia': ['EQUIPAMENTO', 'LOTE', 'NIVEL', 'NIVEL_DESC', 'DATA', 'DATA_HORA', 'ANALITO', 'VALOR', 'FLAG'],
            'tbBioquimica': ['EQUIPAMENTO', 'MATRIZ', 'BLOCO', 'LOTE', 'NIVEL', 'NIVEL_DESC', 'DATA', 'DATA_HORA',
                             'ANALITO', 'PRIMEIRO_DO_DIA', 'VALOR', 'FLAG']}
        for tb, cols0 in originais.items():
            ncols_antes = len(cols0)
            lo = None
            for sh in wb.Worksheets:
                for l in sh.ListObjects:
                    if l.Name == tb:
                        lo = l
            nomes = [c.Name for c in lo.ListColumns]
            if nomes[:ncols_antes] != cols0 or nomes[ncols_antes:] != CAMPOS_NOVOS:
                raise RuntimeError(f'{tb}: colunas inesperadas {nomes} -- arquivo NAO salvo')
            ids = lo.ListColumns('ID_ORIGEM').DataBodyRange.Value
            ids = [x[0] for x in ids]
            nulos = sum(1 for x in ids if x in (None, ''))
            out[tb] = dict(colunas=len(nomes), linhas=len(ids), ids_unicos=len(set(ids)), ids_nulos=nulos,
                           primeiras_colunas=nomes[:ncols_antes])
        out['pivots'] = sum(sh.PivotTables().Count for sh in wb.Worksheets)
        out['segmentacoes'] = xl.run('ContarObjetosDeFiltro')
        out['integridade'] = xl.run('VerificarIntegridade')
        r = log.Cells(log.Rows.Count, 1).End(-4162).Row
        out['ultima_trilha'] = [log.Cells(r, c).Value for c in (4, 5, 6, 7, 8, 9, 10, 11)]
        for k, v in out.items():
            print(k, v)
        for tb in ('tbHematologia', 'tbBioquimica'):
            if out[tb]['ids_unicos'] != out[tb]['linhas'] or out[tb]['ids_nulos']:
                raise RuntimeError(f'{tb}: ID_ORIGEM nao e unico/completo -- arquivo NAO salvo')
        if out['integridade'] != 'CONFORME':
            raise RuntimeError('integridade nao ficou CONFORME -- arquivo NAO salvo')
        wb.Save()
        print('SALVO')
    finally:
        xl.fechar(salvar=False)


if __name__ == '__main__':
    main(sys.argv[1])
