// =====================================================================
//  SEAC_ORIGEM -- leitura da tabela do setor no DB_SEAC.xlsm (somente conexao)
//
//  Le APENAS a tabela do setor deste arquivo (CFG TABELA_ORIGEM):
//  tbHematologia no QC de Hematologia, tbBioquimica no QC de Bioquimica.
//  Os dois setores nunca se misturam fisicamente.
//
//  As colunas saem COM O NOME E NA ORDEM DO DB_SEAC (quem conhece a aba
//  DADOS_HEMATOLOGIA / DADOS_BIOQUIMICA reconhece o recebimento). So os tipos
//  sao fixados e uma coluna e acrescentada na frente: a chave.
//  Nenhum filtro de qualidade, nenhum calculo, nenhuma linha descartada.
//
//  ID_REGISTRO = <PREFIXO_ID>-<ID_ORIGEM>, onde ID_ORIGEM e a coluna 'id' do CSV
//  do SIPEC: sequencial do sistema de origem, unico por resultado, levado
//  adiante pelo DB_SEAC v2.4. Nao e indice de linha: o mesmo resultado tem o
//  mesmo ID em qualquer atualizacao, em qualquer ordenacao.
//  O prefixo (HEM/BIO) separa os setores; os IDs manuais usam MAN_nnnn.
// =====================================================================
let
    C = CFG_INTEGRACAO,
    Fonte = Excel.Workbook(File.Contents(C[CAMINHO_DB_SEAC]), null, true),
    Tabela = Fonte{[Item = C[TABELA_ORIGEM], Kind = "Table"]}[Data],
    Presentes = Table.ColumnNames(Tabela),

    // tipos das colunas conhecidas que existirem nesta tabela (Hematologia nao tem MATRIZ/BLOCO)
    TiposConhecidos = {
        {"EQUIPAMENTO", type text}, {"MATRIZ", type text}, {"BLOCO", type text},
        {"LOTE", type text}, {"NIVEL", Int64.Type}, {"NIVEL_DESC", type text},
        {"DATA", type date}, {"DATA_HORA", type datetime}, {"ANALITO", type text},
        {"PRIMEIRO_DO_DIA", type text}, {"VALOR", type nullable number}, {"FLAG", type nullable text},
        {"ID_ORIGEM", Int64.Type}, {"ID_AMOSTRA", type text}, {"ITEM_ID", Int64.Type},
        {"UNIDADE", type nullable text}},
    Tipos = Table.TransformColumnTypes(Tabela, List.Select(TiposConhecidos, each List.Contains(Presentes, _{0}))),

    ComId = Table.AddColumn(Tipos, "ID_REGISTRO", each C[PREFIXO_ID] & "-" & Text.From([ID_ORIGEM]), type text),
    Saida = Table.ReorderColumns(ComId, {"ID_REGISTRO"} & Presentes)
in
    Saida
