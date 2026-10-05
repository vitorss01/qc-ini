// =====================================================================
//  QA_INTEGRACAO -- verificacoes de governanca da camada de dados  (tblQA_Integracao)
//
//  Uma linha por achado. SEVERIDADE:
//    ERRO ..... processo inadequado; precisa de acao (ex.: inativacao sem justificativa)
//    ALERTA ... o sistema protegeu o dado, mas alguem deve olhar (ex.: duplicidade)
//    INFO ..... registro informativo (ex.: analito sem cadastro no QC_INI)
//
//  Le as tabelas JA carregadas (tblCQ_Final e as de entrada) -- roda depois do
//  DB_CQ_FINAL e nao refaz nenhuma regra: so CONFERE o que ele decidiu.
//  Unica releitura da origem: A03, que compara o recebido com o DB_SEAC atual.
// =====================================================================
let
    C = CFG_INTEGRACAO,
    PREFIXO = Text.Upper(Text.Trim(Text.From(C[PREFIXO_ID]))),          // D15
    Txt = (x as any) as nullable text => if x = null then null else
            let t = Text.Trim(Text.From(x)) in if t = "" then null else t,
    SIM = "SIM", NAO = "NÃO",
    SoDigitos = (t as text) as logical => t <> "" and Text.Length(Text.Select(t, {"0".."9"})) = Text.Length(t),
    // mesma normalizacao do DB_CQ_FINAL (D14): numero puro = ID do setor, sem zeros a esquerda; sem NBSP; MAN_7 -> MAN_0007
    SemZeros = (d as text) as text => let z = Text.TrimStart(d, "0") in if z = "" then "0" else z,
    NormId = (x as any) as nullable text =>
        let t = Txt(x) in
        if t = null then null
        else let u = Text.Upper(Text.Remove(t, {" ", Character.FromNumber(160)})),
                 pre = PREFIXO & "-" in
             if SoDigitos(u) then pre & SemZeros(u)
             else if Text.StartsWith(u, pre) and SoDigitos(Text.Middle(u, Text.Length(pre))) then pre & SemZeros(Text.Middle(u, Text.Length(pre)))
             else if Text.StartsWith(u, "MAN_") and SoDigitos(Text.Middle(u, 4)) then "MAN_" & Text.PadStart(Text.Middle(u, 4), 4, "0")
             else u,
    // D03: tabela obrigatoria ausente PARA (o QA nunca confere "vazio" no lugar de uma tabela perdida)
    Ler = (nome as text, cols as list) as table =>
        let t = try Excel.CurrentWorkbook(){[Name = nome]}[Content]
                otherwise error Error.Record("TABELA_AUSENTE", "Tabela obrigatoria ausente ou renomeada: " & nome, nome)
        in Table.SelectColumns(t, cols, MissingField.UseNull),
    DataHoraTxt = (v as any) as text => try DateTime.ToText(DateTime.From(v), "dd/MM/yyyy HH:mm") otherwise "?",
    Segundos = (v as any) as nullable number => try Number.Round(Number.From(DateTime.From(v)) * 86400) otherwise null,

    F0 = Excel.CurrentWorkbook(){[Name = "tblCQ_Final"]}[Content],
    F = Table.Buffer(Table.SelectRows(F0, each [ID_REGISTRO] <> null or [ORIGEM_RESULTADO] = "MANUAL")),
    Achado = (sev as text, cod as text, teste as text, t as table, detalhe as function) as table =>
        Table.FromRecords(List.Transform(Table.ToRecords(t), each [
            SEVERIDADE = sev, CODIGO = cod, TESTE = teste,
            ID_REGISTRO = Record.FieldOrDefault(_, "ID_REGISTRO", null),
            ANALITO = Record.FieldOrDefault(_, "ANALITO", null),
            NIVEL = Record.FieldOrDefault(_, "NIVEL", null),
            DATA_HORA = Record.FieldOrDefault(_, "DATA_HORA", null),
            DETALHE = detalhe(_)]),
            {"SEVERIDADE", "CODIGO", "TESTE", "ID_REGISTRO", "ANALITO", "NIVEL", "DATA_HORA", "DETALHE"}),

    // E01 -- inativado sem justificativa tecnica (ERRO DE GOVERNANCA)
    E01 = Achado("ERRO", "E01", "Resultado inativado sem justificativa tecnica",
            Table.SelectRows(F, each [STATUS_ANALITICO] = "INATIVADO" and [TEM_JUSTIFICATIVA] <> SIM),
            each "ERRO DE GOVERNANÇA: " & _[ID_REGISTRO] & " | " & (_[ANALITO] ?? "?") & " | nivel " & Text.From(_[NIVEL] ?? "?") &
                 " | " & (try DateTime.ToText(_[DATA_HORA], "dd/MM/yyyy HH:mm") otherwise "?") &
                 " -- inativado sem COMENTARIO_TECNICO; registrar a justificativa na aba COMENTARIOS_TECNICOS"),

    // E02 / E03 -- tabela de inativacao: ID repetido, ID inexistente
    I0 = Ler("tblInativacao_NaoConformes", {"ID_REGISTRO"}),
    I1 = Table.SelectRows(Table.AddColumn(Table.AddIndexColumn(I0, "_linha", 1, 1), "_IDN", each NormId([ID_REGISTRO])), each [_IDN] <> null),
    Rep = Table.SelectRows(Table.Group(I1, {"_IDN"}, {{"_n", each Table.RowCount(_)}, {"_linhas", each Text.Combine(List.Transform([_linha], Text.From), ", ")}}), each [_n] > 1),
    E02 = Achado("ERRO", "E02", "ID_REGISTRO duplicado na tabela de inativacao",
            Table.RenameColumns(Rep, {{"_IDN", "ID_REGISTRO"}}),
            each "aparece " & Text.From(_[_n]) & " vezes (linhas " & _[_linhas] & " da tabela); vale so a primeira, a contagem nao duplica"),
    IdsFinal = List.Buffer(F[ID_REGISTRO]),
    Orfaos = Table.SelectRows(Table.Distinct(I1, {"_IDN"}), each not List.Contains(IdsFinal, [_IDN])),
    E03 = Achado("ERRO", "E03", "ID inativado que nao existe na base",
            Table.RenameColumns(Table.RemoveColumns(Orfaos, {"ID_REGISTRO"}), {{"_IDN", "ID_REGISTRO"}}),
            each "ID digitado na linha " & Text.From(_[_linha]) & " da inativacao nao corresponde a nenhum resultado recebido ou manual"),

    // E04 -- resultado manual incompleto
    E04 = Achado("ERRO", "E04", "Resultado manual com campo obrigatorio ausente",
            Table.SelectRows(F, each [STATUS_ANALITICO] = "MANUAL_INCOMPLETO"),
            each _[MOTIVO_EXCLUSAO_AUTOMATICA]),

    // E05 -- corrida acima de 99 no dia (RUN nao representavel)
    E05 = Achado("ERRO", "E05", "Mais de 99 corridas no mesmo dia",
            Table.SelectRows(F, each [CORRIDA_NO_DIA] <> null and [CORRIDA_NO_DIA] > 99),
            each "RUN nao atribuido; revisar o dado do dia"),

    // A01 -- duplicidade manual x interfaceamento / manual x manual
    A01 = Achado("ALERTA", "A01", "Possivel duplicidade de resultado manual",
            Table.SelectRows(F, each [STATUS_ANALITICO] = "CONFLITO_MANUAL"),
            each _[MOTIVO_EXCLUSAO_AUTOMATICA] & (if _[ID_RELACIONADO] <> null then " -> " & _[ID_RELACIONADO] else "") &
                 ". O manual ficou FORA da estatistica; confirme e inative um deles ou corrija o manual."),

    // A02 -- retransmissao do interfaceamento
    A02 = Achado("ALERTA", "A02", "Retransmissao do interfaceamento (duplicidade de origem)",
            Table.SelectRows(F, each [STATUS_ANALITICO] = "DUPLICIDADE_ORIGEM"),
            each "copia de " & (_[ID_RELACIONADO] ?? "?") & " (mesma amostra, item, instante e valor) -- nao entra em calculo nem no grafico"),

    // A03 -- valor recebido diferente do valor atual na origem
    // D13 (QA-ETL-001): compara tambem ANALITO, LOTE, NIVEL e DATA_HORA -- ID reaproveitado ou corrigido na
    // origem com outro analito/lote/instante nao passa mais em silencio. DATA_HORA comparada ao segundo (a ida
    // e volta pela celula do Excel arredonda a fracao).
    Rec = Ler("tblDB_Recebimento", {"ID_REGISTRO", "VALOR", "FLAG", "ANALITO", "LOTE", "NIVEL", "DATA_HORA"}),
    Origem = SEAC_ORIGEM,
    Agora = Table.SelectColumns(Table.SelectRows(Origem, each [ID_REGISTRO] <> null),
                {"ID_REGISTRO", "VALOR", "FLAG", "ANALITO", "LOTE", "NIVEL", "DATA_HORA"}, MissingField.UseNull),
    Cmp = Table.ExpandTableColumn(Table.NestedJoin(Agora, {"ID_REGISTRO"}, Rec, {"ID_REGISTRO"}, "_r", JoinKind.Inner),
            "_r", {"VALOR", "FLAG", "ANALITO", "LOTE", "NIVEL", "DATA_HORA"},
            {"VALOR_RECEBIDO", "FLAG_RECEBIDA", "ANALITO_RECEBIDO", "LOTE_RECEBIDO", "NIVEL_RECEBIDO", "DH_RECEBIDO"}),
    Difs = (r as record) as list => List.RemoveNulls({
            if (try Number.From(r[VALOR_RECEBIDO]) otherwise null) <> r[VALOR] then "VALOR " & Text.From(r[VALOR_RECEBIDO] ?? "vazio") &
                " -> " & Text.From(r[VALOR] ?? "vazio") else null,
            if Txt(r[FLAG_RECEBIDA]) <> Txt(r[FLAG]) then "FLAG" else null,
            if Txt(r[ANALITO_RECEBIDO]) <> Txt(r[ANALITO]) then "ANALITO " & (Txt(r[ANALITO_RECEBIDO]) ?? "vazio") & " -> " &
                (Txt(r[ANALITO]) ?? "vazio") else null,
            if Txt(r[LOTE_RECEBIDO]) <> Txt(r[LOTE]) then "LOTE " & (Txt(r[LOTE_RECEBIDO]) ?? "vazio") & " -> " & (Txt(r[LOTE]) ?? "vazio") else null,
            if (try Number.From(r[NIVEL_RECEBIDO]) otherwise null) <> (try Number.From(r[NIVEL]) otherwise null) then "NIVEL" else null,
            if Segundos(r[DH_RECEBIDO]) <> Segundos(r[DATA_HORA]) then "DATA_HORA " & DataHoraTxt(r[DH_RECEBIDO]) & " -> " &
                DataHoraTxt(r[DATA_HORA]) else null}),
    Div = Table.SelectRows(Table.AddColumn(Cmp, "_difs", each Difs(_)), each not List.IsEmpty([_difs])),
    A03 = Achado("ALERTA", "A03", "Valor na origem diferente do valor recebido",
            Div,
            each "mudou na origem: " & Text.Combine(_[_difs], "; ") &
                 ". Vale o RECEBIDO (registro imutavel); investigar a correcao na origem."),

    // E06 (D02/P-02) -- linha da origem SEM ID: nao entrou no recebimento; nunca some em silencio
    SemId = Table.SelectRows(Origem, each [ID_REGISTRO] = null or Text.Trim(Text.From([ID_REGISTRO])) = ""),
    E06 = Achado("ERRO", "E06", "Linha da origem sem ID (rejeitada)",
            SemId,
            each "linha do DB_SEAC sem ID valido -- NAO entrou no recebimento: ID_AMOSTRA " & ((try Txt(_[ID_AMOSTRA]) otherwise null) ?? "?") &
                 " | " & ((try Txt(_[ANALITO]) otherwise null) ?? "?") & " | nivel " & ((try Text.From(_[NIVEL]) otherwise null) ?? "?") &
                 " | " & DataHoraTxt(try _[DATA_HORA] otherwise null) & " | ID_ORIGEM informado: " &
                 (try (Txt(_[ID_ORIGEM]) ?? "vazio") otherwise "invalido") & ". Corrigir o ID na origem (SIPEC/DB_SEAC)."),

    // E07 (D01) -- ID repetido na origem (mesma carga) ou ja duplicado no recebimento: entra uma vez so
    RepOrig = Table.SelectRows(Table.Group(Agora, {"ID_REGISTRO"}, {{"_n", each Table.RowCount(_), Int64.Type},
                {"ANALITO", each List.First([ANALITO])}, {"NIVEL", each List.First([NIVEL])},
                {"DATA_HORA", each List.Min([DATA_HORA])}}), each [_n] > 1),
    RepRec = Table.SelectRows(Table.Group(Table.SelectRows(Rec, each [ID_REGISTRO] <> null), {"ID_REGISTRO"},
                {{"_n", each Table.RowCount(_), Int64.Type}, {"ANALITO", each List.First([ANALITO])},
                 {"NIVEL", each List.First([NIVEL])}, {"DATA_HORA", each List.Min([DATA_HORA])}}), each [_n] > 1),
    Rep7 = Table.Distinct(Table.Combine({Table.AddColumn(RepOrig, "_onde", each "na origem (DB_SEAC)"),
                                         Table.AddColumn(RepRec, "_onde", each "no recebimento")}), {"ID_REGISTRO"}),
    E07 = Achado("ERRO", "E07", "ID repetido na origem",
            Rep7,
            each "o ID aparece " & Text.From(_[_n]) & " vezes " & _[_onde] & "; entrou UMA vez (a de menor ITEM_ID). " &
                 "Conferir a origem: dois resultados nunca podem ter o mesmo ID."),

    // E08 (D15) -- prefixo dos IDs recebidos diferente do CFG: trocar o prefixo duplicaria o historico
    PreErr = Table.SelectRows(Rec, each [ID_REGISTRO] <> null and not Text.StartsWith(Text.Upper(Text.From([ID_REGISTRO])), PREFIXO & "-")),
    PreGrp = Table.Group(Table.AddColumn(PreErr, "_pref", each Text.BeforeDelimiter(Text.From([ID_REGISTRO]), "-")), {"_pref"},
                {{"_n", each Table.RowCount(_), Int64.Type}, {"ID_REGISTRO", each List.Min([ID_REGISTRO])}}),
    E08 = Achado("ERRO", "E08", "Prefixo do ID recebido diferente do CFG",
            PreGrp,
            each Text.From(_[_n]) & " ID(s) recebido(s) com prefixo '" & _[_pref] & "-' e o CFG PREFIXO_ID = '" & PREFIXO &
                 "'. Conferir o CFG: com o prefixo trocado, todo o DB_SEAC entraria de novo como resultado novo."),

    // A04 -- inativacao de resultado que ja estava fora por regra automatica
    A04 = Achado("ALERTA", "A04", "Inativacao redundante",
            Table.SelectRows(F, each [INATIVACAO_REGISTRADA] = SIM and [STATUS_ANALITICO] <> "INATIVADO"),
            each "o resultado ja estava fora da estatistica por " & _[STATUS_ANALITICO] & "; a inativacao nao muda nada"),

    // A05 -- comentario tecnico sem inativacao correspondente
    K0 = Ler("tblComentariosTecnicos", {"ID_REGISTRO", "COMENTARIO_TECNICO"}),
    K1 = Table.SelectRows(Table.AddColumn(K0, "_IDN", each NormId([ID_REGISTRO])), each [_IDN] <> null and Txt([COMENTARIO_TECNICO]) <> null),
    Inativados = List.Buffer(Table.SelectRows(F, each [INATIVACAO_REGISTRADA] = SIM)[ID_REGISTRO]),
    SemInat = Table.SelectRows(Table.Distinct(K1, {"_IDN"}), each not List.Contains(Inativados, [_IDN])),
    A05 = Achado("INFO", "A05", "Comentario tecnico de resultado nao inativado",
            Table.RenameColumns(Table.RemoveColumns(SemInat, {"ID_REGISTRO"}), {{"_IDN", "ID_REGISTRO"}}),
            each "comentario mantido como historico (resultado reativado ou ID nunca inativado)"),

    // I01 -- analito recebido sem cadastro no QC_INI (um achado por analito)
    NC = Table.Group(Table.SelectRows(F, each [ANALITO_CADASTRADO] = NAO), {"ANALITO"}, {{"_n", each Table.RowCount(_)}}),
    I01 = Achado("INFO", "I01", "Analito recebido sem cadastro no QC_INI", NC,
            each Text.From(_[_n]) & " resultado(s) guardados na base; nao aparecem no Painel ate o analito ser cadastrado/mapeado"),

    // I02 -- resultados sem valor numerico (um achado por analito)
    SV = Table.Group(Table.SelectRows(F, each [STATUS_ANALITICO] = "SEM_VALOR"), {"ANALITO"}, {{"_n", each Table.RowCount(_)}}),
    I02 = Achado("INFO", "I02", "Resultado sem valor numerico (flag do equipamento)", SV,
            each Text.From(_[_n]) & " resultado(s) com flag e sem numero; guardados, fora de calculo e grafico"),

    // A06 -- manual ATIVO e interfaceamento no mesmo dia/chave, FORA da tolerancia:
    //        nao e tratado como conflito automatico (podem ser duas corridas), mas
    //        nunca passa em silencio
    Chave6 = {"EQUIPAMENTO", "MATRIZ", "LOTE", "NIVEL", "ANALITO", "DATA"},
    ManAtivo = Table.SelectRows(F, each [ORIGEM_RESULTADO] = "MANUAL" and [STATUS_ANALITICO] = "ATIVO"),
    IntDia = Table.SelectColumns(Table.SelectRows(F, each [ORIGEM_RESULTADO] = "INTERFACEAMENTO" and
                                     ([STATUS_ANALITICO] = "ATIVO" or [STATUS_ANALITICO] = "INATIVADO")), Chave6 & {"ID_REGISTRO"}),
    MesmoDia = Table.SelectRows(Table.NestedJoin(ManAtivo, Chave6, IntDia, Chave6, "_d", JoinKind.LeftOuter),
                                each not Table.IsEmpty([_d])),
    A06 = Achado("ALERTA", "A06", "Resultado manual e do interfaceamento no mesmo dia (fora da tolerancia)",
            MesmoDia,
            each "mesmo equipamento/lote/nivel/analito no dia: " & Text.Combine(List.FirstN(_[_d][ID_REGISTRO], 5), ", ") &
                 ". Os dois participam (corridas diferentes). Se forem o MESMO resultado, inative o manual."),

    // I03 -- lote recebido que nao esta no cadastro de lotes (Configuracao!C26:C125):
    //        os resultados estao na base, mas o Painel so mostra lote cadastrado (e com media/DP)
    Cad0 = try Excel.CurrentWorkbook(){[Name = "regLoteCol"]}[Content] otherwise #table({"Column1"}, {}),
    LotesCad = List.Buffer(List.RemoveNulls(List.Transform(Table.Column(Cad0, Table.ColumnNames(Cad0){0}),
                    each let t = Txt(_) in if t = null then null else (if _ is number then Number.ToText(_, "0") else t)))),
    PorLote = Table.Group(Table.SelectRows(F, each [ANALITO_CADASTRADO] = SIM and [ORIGEM_RESULTADO] = "INTERFACEAMENTO"),
                {"EQUIPAMENTO", "LOTE"}, {{"_n", each Table.RowCount(_)}, {"_ini", each List.Min([DATA_HORA])}, {"_fim", each List.Max([DATA_HORA])}}),
    SemCad = Table.SelectRows(PorLote, each not List.Contains(LotesCad, [LOTE])),
    I03 = Achado("INFO", "I03", "Lote recebido sem cadastro no QC_INI",
            Table.RenameColumns(SemCad, {{"_fim", "DATA_HORA"}}),
            each "lote " & _[LOTE] & " (" & _[EQUIPAMENTO] & "): " & Text.From(_[_n]) & " resultado(s) de " &
                 DateTime.ToText(_[_ini], "dd/MM/yyyy") & " a " & DateTime.ToText(_[DATA_HORA], "dd/MM/yyyy") &
                 ". Cadastre o lote (Configuracao) e a media/DP para ve-lo no Painel."),

    // A07 -- MODO_FONTE = HISTORICO (ADR-058): o DB_SEAC NAO foi lido. Alerta em toda
    //        atualizacao enquanto o modo estiver ligado -- esquecer o modo ligado na
    //        rede do laboratorio nunca passa em silencio
    ModoF = Text.Upper(Text.Trim(Text.From(Record.FieldOrDefault(C, "MODO_FONTE", "SEAC") ?? "SEAC"))),
    UltRec = try List.Max(Ler("tblDB_Recebimento", {"DATA_HORA"})[DATA_HORA]) otherwise null,
    A07 = Achado("ALERTA", "A07", "MODO HISTORICO: DB_SEAC nao lido nesta atualizacao",
            if ModoF = "HISTORICO" or ModoF = "HISTÓRICO" then #table({"DATA_HORA"}, {{UltRec}}) else #table({"DATA_HORA"}, {}),
            each "nenhum resultado novo foi recebido; vale o historico ate " &
                 (try DateTime.ToText(DateTime.From(_[DATA_HORA]), "dd/MM/yyyy HH:mm") otherwise "?") &
                 ". Na rede do laboratorio, voltar Cfg_Integracao MODO_FONTE = SEAC."),

    // E09 (D16) -- resultado do interfaceamento sem DATA_HORA: fora da corrida e do calculo (SEM_DATA_HORA)
    E09 = Achado("ERRO", "E09", "Resultado do interfaceamento sem DATA_HORA",
            Table.SelectRows(F, each [ORIGEM_RESULTADO] = "INTERFACEAMENTO" and [DATA_HORA] = null),
            each "sem DATA_HORA na origem: fora da corrida (RUN vazio), do grafico e da estatistica. Corrigir na origem."),

    // A08 (D06) -- "REGISTRAR - LJ" com valor nao reconhecido: vale NAO (nao plota), mas nunca em silencio
    IL = Ler("tblInativacao_NaoConformes", {"ID_REGISTRO", "REGISTRAR - LJ"}),
    LjConhecido = (v as any) as logical =>
        v = null or v is logical or v is number or
        List.Contains({"", "SIM", "S", "TRUE", "VERDADEIRO", "V", "X", "1", "NÃO", "NAO", "N", "FALSE", "FALSO", "F", "0"},
                      Text.Upper(Text.Trim(Text.From(v)))),
    LjEstranho = Table.SelectRows(Table.AddColumn(IL, "_IDN", each NormId([ID_REGISTRO])),
                    each [_IDN] <> null and not LjConhecido([#"REGISTRAR - LJ"])),
    A08 = Achado("ALERTA", "A08", "REGISTRAR - LJ com valor nao reconhecido",
            Table.RenameColumns(Table.RemoveColumns(LjEstranho, {"ID_REGISTRO"}), {{"_IDN", "ID_REGISTRO"}}),
            each "valor '" & Text.From(_[#"REGISTRAR - LJ"]) & "' na inativacao: tratado como NAO (o X nao aparece). " &
                 "Marque ou desmarque a caixa."),

    // A09 (D09, decisao conservadora) -- manual que coincide com interfaceamento INATIVADO: continua fora da
    // estatistica (CONFLITO_MANUAL) ate o RT decidir se "inativar o automatico e lancar o correto" e permitido
    IdsInat = List.Buffer(Table.SelectRows(F, each [STATUS_ANALITICO] = "INATIVADO")[ID_REGISTRO]),
    A09 = Achado("ALERTA", "A09", "Manual coincide com resultado do interfaceamento inativado",
            Table.SelectRows(F, each [ORIGEM_RESULTADO] = "MANUAL" and [STATUS_ANALITICO] = "CONFLITO_MANUAL" and
                                     [ID_RELACIONADO] <> null and List.Contains(IdsInat, [ID_RELACIONADO])),
            each "o automatico " & _[ID_RELACIONADO] & " esta INATIVADO e este manual tem a mesma chave e instante: o manual " &
                 "NAO participa (decisao do RT pendente: manual pode substituir um automatico inativado?)."),

    // A10 (D12) -- Hematologia (sem seletor de equipamento): o Painel, a Estatistica e o Westgard sao do
    // EQUIPAMENTO_PADRAO do CFG; resultado de outro equipamento fica na base, fora dessas series -- nunca em silencio
    EqPad = Text.Upper(Text.Trim(Text.From(Record.FieldOrDefault(C, "EQUIPAMENTO_PADRAO", "") ?? ""))),
    SemSeletor = Text.Upper(Text.Trim(Text.From(Record.FieldOrDefault(C, "SETOR", "") ?? ""))) = "HEMATOLOGIA",
    OutroEq = if not SemSeletor or EqPad = "" then #table({"EQUIPAMENTO"}, {}) else
              Table.Group(Table.SelectRows(F, each [PARTICIPA_ESTATISTICA] = SIM and Text.Upper(Text.Trim(Text.From([EQUIPAMENTO] ?? ""))) <> EqPad),
                          {"EQUIPAMENTO", "ANALITO"}, {{"_n", each Table.RowCount(_), Int64.Type}, {"DATA_HORA", each List.Max([DATA_HORA])}}),
    A10 = Achado("ALERTA", "A10", "Resultado de outro equipamento fora do Painel (Hematologia)",
            OutroEq,
            each Text.From(_[_n]) & " resultado(s) do equipamento '" & Text.From(_[EQUIPAMENTO]) & "': o Painel, a Estatistica e o " &
                 "Westgard da Hematologia sao do equipamento padrao (CFG EQUIPAMENTO_PADRAO = " & EqPad & "). Se for o mesmo " &
                 "aparelho com outro nome, corrija o lancamento."),

    Tudo = Table.Combine({E01, E02, E03, E04, E05, E06, E07, E08, E09, A01, A02, A03, A04, A05, A06, A07, A08, A09, A10, I01, I02, I03}),
    Peso = (s as text) as number => if s = "ERRO" then 1 else if s = "ALERTA" then 2 else 3,
    Ordenado = Table.Sort(Tudo, {{each Peso([SEVERIDADE]), Order.Ascending}, {"CODIGO", Order.Ascending}, {"DATA_HORA", Order.Ascending}}),
    Tipos = Table.TransformColumnTypes(Ordenado, {{"SEVERIDADE", type text}, {"CODIGO", type text}, {"TESTE", type text},
            {"ID_REGISTRO", type nullable text}, {"ANALITO", type nullable text}, {"NIVEL", Int64.Type},
            {"DATA_HORA", type nullable datetime}, {"DETALHE", type nullable text}})
in
    Tipos
