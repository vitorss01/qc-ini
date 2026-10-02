// =====================================================================
//  DB_BIOQUIMICA - CQI da Bioquimica | DIMENSION 1 e DIMENSION 2
//
//  IDENTIFICACAO DO LOTE:  LLLL N -EE     ex.: 8974 1 -01
//     LLLL = lote do kit
//     N    = nivel  (1 NORMAL, 2 PATOLOGICO)
//     EE   = matriz + equipamento:
//              -01  SORO  na DIMENSION 1        -03  URINA na DIMENSION 1
//              -02  SORO  na DIMENSION 2        -04  URINA na DIMENSION 2
//
//  BLOCO agrupa os analitos por programa de controle, porque cada um tem
//  material, lote e horario proprios e ficaria ilegivel numa grade unica:
//     URINA   -> tudo que vem em matriz urina
//     HBA1C   -> A1C
//     PCR     -> RCRP
//     PAINEL  -> demais analitos do painel bioquimico em soro
//
//  ---- A QUESTAO DA DATA UNICA ----
//  Cada analito da Dimension e liberado individualmente e carrega o seu
//  proprio DATA_HORA. Nos dados reais: 69% dos dias tem mais de um horario
//  e 88% tem pelo menos um analito repetido no mesmo dia. Por isso a linha
//  do painel e o INSTANTE de liberacao, nao o dia. Ver aba LEIA-ME.
//
//  ---- PRIMEIRO DO DIA ----
//  PRIMEIRO_DO_DIA rotula cada resultado como "1 - PRIMEIRO DO DIA" ou
//  "2 - REPETICAO NO DIA", conforme seja ou nao o mais antigo daquela
//  combinacao equipamento + lote + nivel + matriz + analito + dia. Serve para
//  o usuario alternar entre ver todos os resultados ou um ponto por dia,
//  sem que nenhum dado seja excluido ou alterado - e apenas um rotulo.
//  O prefixo numerico existe so para que a segmentacao liste o primeiro no topo.
//  Calculado por comparacao com a linha anterior sobre a tabela ordenada,
//  em uma unica passagem; agrupar custaria varios segundos.
// =====================================================================
let
    Bruto = Web.Contents(
        "https://sipec.ini.fiocruz.br/relatorios/relatorio/lab/Hematologia/Indicadores%20Controle%20EXL200/relat209.csv",
        [Headers = [Range = "bytes=-" & Text.From(CFG_BYTES_BIOQ)]]),

    Origem = Csv.Document(Bruto, [Delimiter = ",", Columns = 9, Encoding = 1252, QuoteStyle = QuoteStyle.Csv]),
    SemCabecalho = Table.Skip(Origem, 1),

    // (1) Filtro barato primeiro: o lote do CQI da Dimension tem 8 caracteres,
    //     com o sufixo de matriz/equipamento. Descarta de imediato os outros
    //     programas que convivem no mesmo relatorio (troponina, IMT) e os
    //     lotes legados sem sufixo, que nao permitem identificar o equipamento.
    Candidatos = Table.SelectRows(SemCabecalho,
        each (Text.Length([Column3]) = 8 and Text.Middle([Column3], 5, 1) = "-")
            or ([Column9] = "TNIH" and Text.Length([Column3]) = 10 and Text.Middle([Column3], 7, 1) = "-")),

    NoPeriodo = Table.SelectRows(Candidatos,
        each Text.Length([Column8]) >= 19
            and Number.FromText(Text.Middle([Column8], 6, 4)) >= Date.Year(CFG_CORTE_BIOQ)),

    // (2) Valida e decompoe o lote
    Campos = Table.AddColumn(NoPeriodo, "r", each
        let
            t      = [Column3],
            trop   = [Column9] = "TNIH",
            corpo  = if trop then Text.Start(t, Text.Length(t) - 3) else Text.Start(t, 5),
            suf    = Text.End(t, 2),
            nivRaw = Text.End(corpo, 1),
            niv    = if trop then (if nivRaw = "5" then "1" else if nivRaw = "6" then "2" else nivRaw) else nivRaw,
            an     = [Column9],
            urina  = suf = "03" or suf = "04",
            valido = Text.Length(Text.Select(corpo, {"0" .. "9"})) = Text.Length(corpo)
                     and (niv = "1" or niv = "2")
                     and List.Contains({"01", "02", "03", "04"}, suf),
            num    = try Number.FromText([Column5], "pt-BR")
        in
            if not valido then null
            else [ DH     = DateTime.FromText([Column8], [Format = "dd/MM/yyyy HH:mm:ss", Culture = "pt-BR"]),
                   LOTE   = if trop then corpo else Text.Start(corpo, 4),
                   NIVEL  = Number.FromText(niv),
                   ND     = if niv = "1" then "1 - NORMAL" else "2 - PATOLOGICO",
                   EQ     = if suf = "01" or suf = "03" then "DIMENSION 1" else "DIMENSION 2",
                   MTZ    = if urina then "URINA" else "SORO",
                   BLOCO  = if urina then "URINA"
                            else if an = "A1C" then "HBA1C"
                            else if an = "RCRP" then "PCR"
                            else if an = "FERR" then "FERR"
                            else if an = "TNIH" then "TNIH"
                            else if an = "UCFP" then "UCFP"
                            else "PAINEL",
                   V      = if num[HasError] then null else num[Value],
                   FL     = if num[HasError] then [Column5] else null ]),

    SoValidos = Table.SelectRows(Campos, each [r] <> null),

    Exp = Table.ExpandRecordColumn(SoValidos, "r",
        {"DH", "LOTE", "NIVEL", "ND", "EQ", "MTZ", "BLOCO", "V", "FL"},
        {"DATA_HORA", "LOTE", "NIVEL", "NIVEL_DESC", "EQUIPAMENTO", "MATRIZ", "BLOCO", "VALOR", "FLAG"}),

    NaJanela = Table.SelectRows(Exp, each DateTime.Date([DATA_HORA]) >= CFG_CORTE_BIOQ),

    Renomeado = Table.RenameColumns(NaJanela, {{"Column9", "ANALITO"}}),
    ComData = Table.AddColumn(Renomeado, "DATA", each DateTime.Date([DATA_HORA]), type date),

    // (3) chave do "primeiro do dia" e ordenacao cronologica dentro dela
    ComChave = Table.AddColumn(ComData, "K", each
        [EQUIPAMENTO] & "|" & [LOTE] & "|" & Text.From([NIVEL]) & "|" & [MATRIZ] & "|"
        & [ANALITO] & "|" & Date.ToText([DATA], [Format = "yyyyMMdd"]), type text),

    Ordenado = Table.Buffer(Table.Sort(ComChave,
        {{"K", Order.Ascending}, {"DATA_HORA", Order.Ascending}})),

    // Comparacao com a linha anterior por indice sobre a lista de chaves
    // bufferizada: uma passagem, acesso O(1). Medido em 63.881 linhas -
    // esta forma custa 0,3 s; a variante com List.Zip custava 15 s.
    ChavesB = List.Buffer(Ordenado[K]),
    ComIdx  = Table.AddIndexColumn(Ordenado, "I", 0, 1, Int64.Type),
    ComMarca = Table.AddColumn(ComIdx, "PRIMEIRO_DO_DIA",
        each if [I] = 0 or ChavesB{[I] - 1} <> [K]
             then "1 - PRIMEIRO DO DIA" else "2 - REPETICAO NO DIA", type text),

    Saida = Table.SelectColumns(ComMarca,
        {"EQUIPAMENTO", "MATRIZ", "BLOCO", "LOTE", "NIVEL", "NIVEL_DESC",
         "DATA", "DATA_HORA", "ANALITO", "PRIMEIRO_DO_DIA", "VALOR", "FLAG"}),

    Tipos = Table.TransformColumnTypes(Saida, {
        {"EQUIPAMENTO", type text}, {"MATRIZ", type text}, {"BLOCO", type text},
        {"LOTE", type text}, {"NIVEL", Int64.Type}, {"NIVEL_DESC", type text},
        {"DATA", type date}, {"DATA_HORA", type datetime}, {"ANALITO", type text},
        {"PRIMEIRO_DO_DIA", type text}, {"VALOR", type nullable number},
        {"FLAG", type nullable text}})
in
    Tipos