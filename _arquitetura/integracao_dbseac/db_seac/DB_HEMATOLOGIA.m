// =====================================================================
//  DB_HEMATOLOGIA - CQI da Hematologia | Sysmex XN-1000
//  Niveis: 01 NORMAL | 02 MEDIUM | 03 HIGH
//  Fonte: SIPEC / relat207.csv
//
//  TRES OTIMIZACOES, EM ORDEM DE IMPACTO:
//  1. Range HTTP: baixa so a cauda recente do CSV. O arquivo guarda todo o
//     historico desde 2009 (17 MB) e cresce todo dia; o IIS responde 206.
//  2. Ordem dos filtros: o teste mais seletivo e mais barato vem primeiro
//     (Text.Length do idAmostra). Filtrar por amostra antes do resto e
//     2,8x mais rapido que a ordem inversa - medido.
//  3. Conversao de tipos por ultimo, sobre ~20% das linhas iniciais.
// =====================================================================
let
    // Se o servidor ignorar o Range, ou o arquivo for menor que o pedido,
    // vem tudo e o Table.Skip(1) descarta o cabecalho. Se vier so o trecho,
    // o Skip(1) descarta a primeira linha, que chega cortada ao meio.
    Bruto = Web.Contents(
        "https://sipec.ini.fiocruz.br/relatorios/relatorio/lab/Hematologia/Indicadores%20Controle%20XN1000/relat207.csv",
        [Headers = [Range = "bytes=-" & Text.From(CFG_BYTES_HEMA)]]),

    Origem = Csv.Document(Bruto, [Delimiter = ",", Columns = 9, Encoding = 1252, QuoteStyle = QuoteStyle.Csv]),
    SemCabecalho = Table.Skip(Origem, 1),

    // (1) So o CQI de rotina: "QC-" + 6 digitos de lote + 2 de nivel (01/02/03).
    //     O filtro antigo (StartsWith "QC" + EndsWith "01") deixava passar os
    //     materiais de linearidade e de reagente - QC-WBC176 01, QC-RBC176-01,
    //     QC-LL176 01, QC-PLT176 01, QC-REAS 01 - que contaminavam o Nivel 1.
    //     Tambem exclui XbarM1/XbarM2, que sao media movel e nao CQI de lote.
    SoQC = Table.SelectRows(SemCabecalho,
        each Text.Length([Column3]) = 11
            and Text.StartsWith([Column3], "QC-")
            and List.Contains({"01", "02", "03"}, Text.End([Column3], 2))
            and Text.Length(Text.Select(Text.Middle([Column3], 3), {"0" .. "9"})) = 8),

    // (2) Fora os scattergrams e histogramas (valor "PNG&R&..."): sao 20% do
    //     arquivo e nao tem valor numerico. SCAT_*/DIST_* equivale exatamente
    //     ao conjunto PNG - conferido nos 222.151 registros da fonte.
    SoNumerico = Table.SelectRows(SoQC,
        each not Text.StartsWith([Column9], "SCAT_")
            and not Text.StartsWith([Column9], "DIST_")),

    // (3) Recorte do periodo ainda em texto (comparacao barata)
    NoPeriodo = Table.SelectRows(SoNumerico,
        each Text.Length([Column8]) >= 19
            and Number.FromText(Text.Middle([Column8], 6, 4)) >= Date.Year(CFG_CORTE_HEMA)),

    // (4) Agora sim os tipos. "----" e "++++" sao flags do XN-1000
    //     (parametro nao reportado / fora de faixa): viram FLAG, nao erro.
    Campos = Table.AddColumn(NoPeriodo, "r", each
        let niv = Text.End([Column3], 2) in
        [ DH    = DateTime.FromText([Column8], [Format = "dd/MM/yyyy HH:mm:ss", Culture = "pt-BR"]),
          LOTE  = Text.Middle([Column3], 3, 6),
          NIVEL = Number.FromText(niv),
          ND    = if niv = "01" then "1 - NORMAL"
                  else if niv = "02" then "2 - MEDIUM" else "3 - HIGH",
          V     = try Number.FromText([Column5], "pt-BR") otherwise null,
          FL    = if [Column5] = "----" or [Column5] = "++++" then [Column5] else null ]),

    Exp = Table.ExpandRecordColumn(Campos, "r", {"DH", "LOTE", "NIVEL", "ND", "V", "FL"},
        {"DATA_HORA", "LOTE", "NIVEL", "NIVEL_DESC", "VALOR", "FLAG"}),

    NaJanela = Table.SelectRows(Exp, each DateTime.Date([DATA_HORA]) >= CFG_CORTE_HEMA),

    Renomeado = Table.RenameColumns(NaJanela, {{"Column2", "EQUIPAMENTO"}, {"Column9", "ANALITO"},
        {"Column1", "ID_ORIGEM"}, {"Column3", "ID_AMOSTRA"}, {"Column4", "ITEM_ID"}, {"Column6", "UNIDADE"}}),
    ComData = Table.AddColumn(Renomeado, "DATA", each DateTime.Date([DATA_HORA]), type date),

    Saida = Table.SelectColumns(ComData,
        {"EQUIPAMENTO", "LOTE", "NIVEL", "NIVEL_DESC", "DATA", "DATA_HORA", "ANALITO", "VALOR", "FLAG",
         // v2.4: chave de origem do SIPEC e rastreabilidade, sempre NO FIM
         "ID_ORIGEM", "ID_AMOSTRA", "ITEM_ID", "UNIDADE"}),

    Tipos = Table.TransformColumnTypes(Saida, {
        {"EQUIPAMENTO", type text}, {"LOTE", type text}, {"NIVEL", Int64.Type}, {"NIVEL_DESC", type text},
        {"DATA", type date}, {"DATA_HORA", type datetime}, {"ANALITO", type text},
        {"VALOR", type nullable number}, {"FLAG", type nullable text},
        {"ID_ORIGEM", Int64.Type}, {"ID_AMOSTRA", type text}, {"ITEM_ID", Int64.Type}, {"UNIDADE", type nullable text}})
in
    Tipos