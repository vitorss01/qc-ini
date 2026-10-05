# -*- coding: utf-8 -*-
"""Parametros por produto da integracao DB_SEAC -> DB_CQ_FINAL.

Tudo o que difere entre Hematologia e Bioquimica esta AQUI; as consultas M sao
as mesmas nos dois arquivos e leem estes valores da tabela tblConfigIntegracao.

De/para de analitos: o equipamento chama 'GLUC' o que o cadastro do QC_INI chama
'Glicose'. A correspondencia foi PROVADA, nao suposta: cada par abaixo foi
conferido valor a valor contra o historico lancado no QC_INI (DB_Resultados,
lote 8974, Dimension 1) -- ver testes/prova_depara.py e o relatorio.
"""

CAMINHO_DB_SEAC = (r'C:\Users\vitor.santos\OneDrive - MSFT\INI_DB_AJUSTES\OneDrive - MSFT\Desktop'
                   r'\QC_INI_PRONTO\qc-ini\INTERFACEAMENTO_DB_COPIA\DB_SEAC.xlsm')

COMUM = {
    'CAMINHO_DB_SEAC': (CAMINHO_DB_SEAC, 'Arquivo de origem (DB_SEAC.xlsm). Trocar aqui se o arquivo mudar de pasta.'),
    'GAP_CORRIDA_MIN': (10, 'Intervalo maximo (min) entre resultados consecutivos de uma mesma corrida (agrupa os niveis).'),
    'TOLERANCIA_CONFLITO_MIN': (30, 'Resultado manual a ate N min de um do interfaceamento, mesma chave = possivel duplicidade.'),
    # ADR-058: SEAC e o padrao e o unico modo de producao. HISTORICO existe para
    # quando o DB_SEAC esta fora de alcance (fora da rede do laboratorio): o
    # historico ja recebido continua sendo reprocessado, nada novo entra, e o
    # proprio sistema avisa (QA A07 + Audit_Log) enquanto o modo estiver ligado.
    'MODO_FONTE': ('SEAC', 'SEAC = recebe do DB_SEAC (se inacessivel, a atualizacao PARA). '
                           'HISTORICO = nao le o DB_SEAC; reprocessa so o que ja foi recebido '
                           '(inativacoes, manuais, comentarios). Voltar para SEAC na rede do laboratorio.'),
}

PRODUTOS = {
    'Hematologia': {
        'arquivo': 'QC_Hematologia.xlsm',
        'cfg': {
            'SETOR': ('HEMATOLOGIA', 'Setor deste arquivo. So os dados deste setor sao recebidos.'),
            'TABELA_ORIGEM': ('tbHematologia', 'Tabela do setor dentro do DB_SEAC.'),
            'PREFIXO_ID': ('HEM', 'Prefixo do ID_REGISTRO do interfaceamento (HEM-<id SIPEC>). Manuais: MAN-HEM-...'),
            'MATRIZ_PADRAO': ('SANGUE TOTAL', 'Matriz atribuida quando a origem nao informa.'),
            'EQUIPAMENTO_PADRAO': ('XN1000', 'Equipamento analisado por padrao e sugerido no lancamento manual.'),
            'NIVEIS': (3, 'Numero de niveis de controle do setor.'),
        },
        # (MATRIZ, ANALITO_ORIGEM, ANALITO_QCINI, BLOCO) na ordem do cadastro
        'depara': [(None, a, a, None) for a in
                   ['WBC', 'RBC', 'HGB', 'HCT', 'MCV', 'MCH', 'MCHC', 'PLT', 'NEUT%', 'LYMPH%', 'MONO%', 'EO%',
                    'BASO%', 'NEUT#', 'LYMPH#', 'MONO#', 'EO#', 'BASO#', 'IG%', 'IG#', 'NRBC%', 'NRBC#',
                    'RDW-SD', 'RDW-CV', 'MPV', 'RET%', 'RET#', 'IRF']]
                  + [(None, a, None, None) for a in ['RET-HE', 'IPF', 'IPF#']],
        'equipamentos': ['XN1000'],
    },
    'Bioquimica': {
        'arquivo': 'QC_Bioquimica.xlsm',
        'cfg': {
            'SETOR': ('BIOQUIMICA', 'Setor deste arquivo. So os dados deste setor sao recebidos.'),
            'TABELA_ORIGEM': ('tbBioquimica', 'Tabela do setor dentro do DB_SEAC.'),
            'PREFIXO_ID': ('BIO', 'Prefixo do ID_REGISTRO do interfaceamento (BIO-<id SIPEC>). Manuais: MAN-BIO-...'),
            'MATRIZ_PADRAO': ('SORO', 'Matriz atribuida quando a origem nao informa.'),
            'EQUIPAMENTO_PADRAO': ('DIMENSION 1', 'Equipamento analisado por padrao e sugerido no lancamento manual.'),
            'NIVEIS': (2, 'Numero de niveis de controle do setor.'),
        },
        'depara': [
            ('SORO', 'LA', 'Lactato', 'PAINEL'), ('SORO', 'URCA', 'Ácido úrico', 'PAINEL'),
            ('SORO', 'ALB', 'Albumina', 'PAINEL'), ('SORO', 'DBI', 'Bilirrubina direta', 'PAINEL'),
            ('SORO', 'TBI', 'Bilirrubina total', 'PAINEL'), ('SORO', 'CA', 'Cálcio', 'PAINEL'),
            ('SORO', 'IBCT', 'Capacidade de fixação do ferro', 'PAINEL'), ('SORO', 'CHOL', 'Colesterol total', 'PAINEL'),
            ('SORO', 'AHDL', 'HDL colesterol', 'PAINEL'), ('SORO', 'CRE2', 'Creatinina', 'PAINEL'),
            ('SORO', 'IRON', 'Ferro', 'PAINEL'), ('SORO', 'PHOS', 'Fósforo', 'PAINEL'),
            ('SORO', 'GLUC', 'Glicose', 'PAINEL'), ('SORO', 'A1C', 'Hemoglobina glicada', 'HBA1C'),
            ('SORO', 'MG', 'Magnésio', 'PAINEL'), ('SORO', 'TP', 'Proteína total', 'PAINEL'),
            ('SORO', 'TGL', 'Triglicerídeos', 'PAINEL'), ('SORO', 'BUN', 'Ureia', 'PAINEL'),
            ('SORO', 'AMY', 'Amilase', 'PAINEL'), ('SORO', 'CKI', 'Creatina fosfoquinase', 'PAINEL'),
            ('SORO', 'ALPI', 'Fosfatase alcalina', 'PAINEL'), ('SORO', 'GGT', 'GGT', 'PAINEL'),
            ('SORO', 'LDI', 'Lactato desidrogenase', 'PAINEL'), ('SORO', 'LIP', 'Lipase', 'PAINEL'),
            ('SORO', 'AST', 'AST (TGO)', 'PAINEL'), ('SORO', 'ALTI', 'ALT (TGP)', 'PAINEL'),
            ('SORO', 'RCRP', 'Proteína C-reativa', 'PCR'), ('SORO', 'TCO2', 'Bicarbonato', 'PAINEL'),
            ('SORO', 'NA', 'Sódio', 'PAINEL'), ('SORO', 'K', 'Potássio', 'PAINEL'), ('SORO', 'CL', 'Cloro', 'PAINEL'),
            # recebidos e guardados, ainda sem cadastro no QC_INI (ANALITO_QCINI vazio = nao mapeado)
            ('SORO', 'FERR', None, 'FERR'), ('SORO', 'TNIH', None, 'TNIH'), ('SORO', 'UCFP', None, 'UCFP'),
            ('URINA', 'CRE2', None, 'URINA'), ('URINA', 'PHOS', None, 'URINA'), ('URINA', 'MALB', None, 'URINA'),
            ('URINA', 'UCFP', None, 'URINA'),
        ],
        'equipamentos': ['DIMENSION 1', 'DIMENSION 2'],
    },
}


def linhas_cfg(produto):
    p = PRODUTOS[produto]
    out = [[k, v, d] for k, (v, d) in p['cfg'].items()]
    out += [[k, v, d] for k, (v, d) in COMUM.items()]
    return out


def linhas_depara(produto):
    out = []
    for i, (mz, orig, qc, bloco) in enumerate(PRODUTOS[produto]['depara'], start=1):
        obs = '' if qc else 'sem cadastro no QC_INI: guardado na base, fora do Painel'
        out.append([mz, orig, qc, bloco, i, obs])
    return out
