// =====================================================================
//  DB_RECEBIMENTO -- CAMADA 1, INGESTAO  (tabela tblDB_Recebimento)
//
//  Recebe os resultados do DB_SEAC e os GUARDA, com as colunas do DB_SEAC
//  (mesmos nomes, mesma ordem) + ID_REGISTRO na frente + RECEBIDO_EM no fim.
//
//  Por que acumula, e nao espelha: o DB_SEAC carrega uma janela movel de 12
//  meses (CFG_MESES_*). Um espelho perderia todo resultado que saisse da
//  janela -- e com ele o historico do Levey-Jennings, a estatistica de um lote
//  longo e a propria inativacao (o ID inativado deixaria de existir).
//  Receber e acrescentar; nunca apagar.
//
//  Como acumula sem VBA: a consulta le a propria tabela ja carregada
//  (Excel.CurrentWorkbook) e acrescenta so os ID_REGISTRO que ainda nao estao
//  la (anti-juncao SO sobre o que ja foi recebido -- e o que torna o refresh
//  idempotente). Linha ja recebida nunca e reescrita: se a origem mudar o
//  valor de um ID antigo, o valor RECEBIDO fica, e a divergencia aparece no
//  QA (QA_INTEGRACAO, codigo A03). RECEBIDO_EM carimba a primeira chegada.
//
//  Se a leitura do DB_SEAC falhar, a atualizacao falha e o Excel mantem a
//  tabela como estava -- nada e perdido.
//
//  Sem regra de qualidade, sem calculo, sem exclusao, sem inativacao.
// =====================================================================
let
    TiposConhecidos = {
        {"ID_REGISTRO", type text}, {"EQUIPAMENTO", type text}, {"MATRIZ", type text}, {"BLOCO", type text},
        {"LOTE", type text}, {"NIVEL", Int64.Type}, {"NIVEL_DESC", type text},
        {"DATA", type date}, {"DATA_HORA", type datetime}, {"ANALITO", type text},
        {"PRIMEIRO_DO_DIA", type text}, {"VALOR", type nullable number}, {"FLAG", type nullable text},
        {"ID_ORIGEM", Int64.Type}, {"ID_AMOSTRA", type text}, {"ITEM_ID", Int64.Type},
        {"UNIDADE", type nullable text}, {"RECEBIDO_EM", type datetime}},
    Tipar = (t as table) as table =>
        Table.TransformColumnTypes(t, List.Select(TiposConhecidos, each List.Contains(Table.ColumnNames(t), _{0}))),

    // ---- o que o DB_SEAC tem agora ----
    Origem = SEAC_ORIGEM,

    // ---- o que ja foi recebido (a propria tabela) ----
    Lido = try Excel.CurrentWorkbook(){[Name = "tblDB_Recebimento"]}[Content] otherwise null,

    // colunas = as da origem; coluna que um dia saia do DB_SEAC continua aqui (nada some)
    DaOrigem = List.RemoveItems(Table.ColumnNames(Origem), {"ID_REGISTRO"}),
    // so colunas de DADO: o texto provisorio que o Excel poe na tabela durante a
    // primeira carga ("DadosExternos_1: Obtendo dados...") nao pode virar coluna
    JaTinha = if Lido = null then {} else
              List.Select(List.RemoveItems(Table.ColumnNames(Lido), {"ID_REGISTRO", "RECEBIDO_EM"}),
                          each not Text.Contains(_, ":") and not Text.StartsWith(_, "DadosExternos")
                               and not Text.StartsWith(_, "ExternalData") and not Text.StartsWith(_, "Column")),
    Colunas = {"ID_REGISTRO"} & List.Distinct(DaOrigem & JaTinha) & {"RECEBIDO_EM"},

    Existente =
        if Lido = null then #table(Colunas, {})
        else Table.SelectRows(Table.SelectColumns(Lido, Colunas, MissingField.UseNull),
                              each [ID_REGISTRO] <> null and Text.Trim(Text.From([ID_REGISTRO])) <> ""),
    ExistenteT = Tipar(Existente),

    Novos = Table.RemoveColumns(
                Table.NestedJoin(Origem, {"ID_REGISTRO"},
                                 Table.SelectColumns(ExistenteT, {"ID_REGISTRO"}), {"ID_REGISTRO"},
                                 "_ja", JoinKind.LeftAnti),
                {"_ja"}),
    Carimbo = DateTime.LocalNow(),
    NovosC = Table.AddColumn(Novos, "RECEBIDO_EM", each Carimbo, type datetime),

    Tudo = Table.Combine({ExistenteT, Tipar(Table.SelectColumns(NovosC, Colunas, MissingField.UseNull))}),
    Ordenado = Table.Sort(Tudo, {{"DATA_HORA", Order.Ascending}, {"ID_ORIGEM", Order.Ascending}})
in
    Ordenado
