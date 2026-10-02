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
    PREFIXO = C[PREFIXO_ID],
    Txt = (x as any) as nullable text => if x = null then null else
            let t = Text.Trim(Text.From(x)) in if t = "" then null else t,
    SIM = "SIM", NAO = "NÃO",
    SoDigitos = (t as text) as logical => t <> "" and Text.Length(Text.Select(t, {"0".."9"})) = Text.Length(t),
    // mesma normalizacao do DB_CQ_FINAL: numero puro = ID do setor; MAN_7 -> MAN_0007
    NormId = (x as any) as nullable text =>
        let t = Txt(x) in
        if t = null then null
        else let u = Text.Upper(Text.Remove(t, {" "})) in
             if SoDigitos(u) then PREFIXO & "-" & u
             else if Text.StartsWith(u, "MAN_") and SoDigitos(Text.Middle(u, 4)) then "MAN_" & Text.PadStart(Text.Middle(u, 4), 4, "0")
             else u,
    Ler = (nome as text, cols as list) as table =>
        let t = try Excel.CurrentWorkbook(){[Name = nome]}[Content] otherwise #table(cols, {})
        in Table.SelectColumns(t, cols, MissingField.UseNull),

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
    Rec = Ler("tblDB_Recebimento", {"ID_REGISTRO", "VALOR", "FLAG"}),
    Agora = Table.SelectColumns(SEAC_ORIGEM, {"ID_REGISTRO", "VALOR", "FLAG", "ANALITO", "NIVEL", "DATA_HORA"}),
    Cmp = Table.ExpandTableColumn(Table.NestedJoin(Agora, {"ID_REGISTRO"}, Rec, {"ID_REGISTRO"}, "_r", JoinKind.Inner),
            "_r", {"VALOR", "FLAG"}, {"VALOR_RECEBIDO", "FLAG_RECEBIDA"}),
    Div = Table.SelectRows(Cmp, each (try Number.From([VALOR_RECEBIDO]) otherwise null) <> [VALOR] or Txt([FLAG_RECEBIDA]) <> Txt([FLAG])),
    A03 = Achado("ALERTA", "A03", "Valor na origem diferente do valor recebido",
            Div,
            each "recebido " & Text.From(_[VALOR_RECEBIDO] ?? "vazio") & ", origem agora " & Text.From(_[VALOR] ?? "vazio") &
                 ". Vale o RECEBIDO (registro imutavel); investigar a correcao na origem."),

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

    Tudo = Table.Combine({E01, E02, E03, E04, E05, A01, A02, A03, A04, A05, A06, I01, I02, I03}),
    Peso = (s as text) as number => if s = "ERRO" then 1 else if s = "ALERTA" then 2 else 3,
    Ordenado = Table.Sort(Tudo, {{each Peso([SEVERIDADE]), Order.Ascending}, {"CODIGO", Order.Ascending}, {"DATA_HORA", Order.Ascending}}),
    Tipos = Table.TransformColumnTypes(Ordenado, {{"SEVERIDADE", type text}, {"CODIGO", type text}, {"TESTE", type text},
            {"ID_REGISTRO", type nullable text}, {"ANALITO", type nullable text}, {"NIVEL", Int64.Type},
            {"DATA_HORA", type nullable datetime}, {"DETALHE", type nullable text}})
in
    Tipos
