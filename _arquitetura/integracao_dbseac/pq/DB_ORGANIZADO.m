// =====================================================================
//  DB_ORGANIZADO -- visao HORIZONTAL de conferencia  (tabela tblDB_Organizado)
//
//  Reproduz a grade das tabelas dinamicas do DB_SEAC:
//    HEMA_NIVEL_1/2/3 ..... linhas = DATA_HORA (instante da corrida),
//                           colunas = ANALITO, valor = MEDIA de VALOR,
//                           filtros = LOTE e NIVEL
//    BIOQ_<EQ>_<NIVEL> .... idem, por EQUIPAMENTO e BLOCO (painel, HbA1c, PCR, urina...)
//  Aqui numa tabela so -- DATA, DATA_HORA, EQUIPAMENTO, [MATRIZ, BLOCO], LOTE,
//  NIVEL na frente e um analito por coluna -- para o operador ver todas as
//  corridas, niveis e analitos lado a lado e filtrar como no DB_SEAC.
//
//  E camada de CONFERENCIA, nao de calculo: nem o Levey-Jennings nem a
//  Estatistica leem daqui (eles leem tblCQ_Final, vertical, uma linha por
//  resultado). O valor da celula e a media do instante, como no DB_SEAC;
//  quando houve mais de um resultado do mesmo analito no mesmo instante
//  (replica ou retransmissao), RESULTADOS_AGREGADOS diz quantos foram
//  somados -- a conferencia nao esconde que agregou.
//
//  As colunas de analito nascem do dado (Table.Pivot): analito novo vira
//  coluna sozinho. Ordem: a do cadastro de de/para (ORDEM), depois alfabetica.
// =====================================================================
let
    R0 = Excel.CurrentWorkbook(){[Name = "tblDB_Recebimento"]}[Content],
    Presentes = Table.ColumnNames(R0),
    TiposConhecidos = {
        {"EQUIPAMENTO", type text}, {"MATRIZ", type text}, {"BLOCO", type text},
        {"LOTE", type text}, {"NIVEL", Int64.Type}, {"NIVEL_DESC", type text},
        {"DATA", type date}, {"DATA_HORA", type datetime}, {"ANALITO", type text},
        {"VALOR", type nullable number}},
    R = Table.TransformColumnTypes(
            Table.SelectRows(R0, each [ID_REGISTRO] <> null),
            List.Select(TiposConhecidos, each List.Contains(Presentes, _{0}))),

    // chave da linha: as colunas de filtro/linha das dinamicas do DB_SEAC que existirem
    Chave = List.Select({"DATA", "DATA_HORA", "EQUIPAMENTO", "MATRIZ", "BLOCO", "LOTE", "NIVEL", "NIVEL_DESC"},
                        each List.Contains(Presentes, _)),

    // quantos resultados compoem cada linha e quantos foram agregados (>1 no mesmo analito)
    Contagem = Table.Group(R, Chave & {"ANALITO"}, {{"_n", each Table.RowCount(_), Int64.Type}}),
    PorLinha = Table.Group(Contagem, Chave, {
        {"N_RESULTADOS", each List.Sum([_n]), Int64.Type},
        {"RESULTADOS_AGREGADOS", each List.Sum(List.Transform([_n], each if _ > 1 then _ else 0)), Int64.Type}}),

    // ordem das colunas de analito
    DP0 = try Excel.CurrentWorkbook(){[Name = "tblDeParaAnalitos"]}[Content] otherwise #table({"ANALITO_ORIGEM", "ORDEM"}, {}),
    DP = Table.Buffer(Table.SelectColumns(DP0, {"ANALITO_ORIGEM", "ORDEM"}, MissingField.UseNull)),
    NomesDP = List.Buffer(DP[ANALITO_ORIGEM]),
    Analitos = List.Distinct(R[ANALITO]),
    OrdemDe = (a as text) as number =>
        let p = List.PositionOf(NomesDP, a)
        in if p >= 0 and DP{p}[ORDEM] <> null then Number.From(DP{p}[ORDEM]) else 100000,
    Ordenados = List.Sort(Analitos, (x, y) =>
        let ox = OrdemDe(x), oy = OrdemDe(y)
        in if ox <> oy then Value.Compare(ox, oy) else Value.Compare(x, y)),

    Base = Table.SelectColumns(R, Chave & {"ANALITO", "VALOR"}),
    Grade = Table.Pivot(Base, Ordenados, "ANALITO", "VALOR", List.Average),

    // NestedJoin + expand: Table.Join recusaria as colunas-chave repetidas
    Junta = Table.NestedJoin(Grade, Chave, PorLinha, Chave, "_c", JoinKind.LeftOuter),
    ComContagem = Table.ExpandTableColumn(Junta, "_c", {"N_RESULTADOS", "RESULTADOS_AGREGADOS"}),
    Saida = Table.SelectColumns(ComContagem, Chave & Ordenados & {"N_RESULTADOS", "RESULTADOS_AGREGADOS"}),
    Ordenado = Table.Sort(Saida, List.Transform(List.Select({"DATA_HORA", "EQUIPAMENTO", "BLOCO", "NIVEL"},
                                                             each List.Contains(Chave, _)), each {_, Order.Ascending}))
in
    Ordenado
