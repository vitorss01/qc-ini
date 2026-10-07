// =====================================================================
//  DB_CQ_FINAL -- CAMADA 2 -> 3: FONTE OFICIAL DOS RESULTADOS DO QC_INI  (tblCQ_Final)
//
//  SINGLE SOURCE OF TRUTH. Unica tabela lida pelo Levey-Jennings, Westgard,
//  Estatistica, Sigma, indicadores e Power BI. Construida SO por esta
//  consulta, a partir de:
//    tblDB_Recebimento ........... resultados do interfaceamento (DB_SEAC)
//    tblResultados_Manuais ....... resultados digitados (staging; falha de interface)
//    tblInativacao_NaoConformes .. ID + ANALITO + MOTIVO retirados da populacao estatistica (soft-delete)
//    tblComentariosTecnicos ...... justificativa tecnica por ID
//    tblDeParaAnalitos ........... nome do equipamento -> nome do cadastro
//
//  GRANULARIDADE: UMA LINHA = UM RESULTADO ANALITICO (ID_REGISTRO unico).
//  O recebimento ja e vertical (um resultado por linha), entao a final nasce
//  dele direto -- sem pivot/unpivot de ida e volta. A grade horizontal e so a
//  visao de conferencia (DB_ORGANIZADO).
//
//  NADA E APAGADO. Todo registro recebido ou digitado sai aqui, com a
//  classificacao que diz o que ele e. A inativacao e LEFT OUTER JOIN com a
//  tabela de inativacao (nunca anti-join): o inativado CONTINUA na tabela,
//  com PARTICIPA_ESTATISTICA = NÃO, e pode aparecer no LJ como X vermelho.
//  A decisao "participa da estatistica?" e "como aparece no LJ?" existe SO
//  aqui; VBA, grafico, Estatistica e Power BI leem, nunca reinterpretam.
//
//  STATUS_ANALITICO (precedencia de cima para baixo)
//    MANUAL_INCOMPLETO ... linha manual sem campo obrigatorio, ID fora do
//                          padrao MAN_nnnn ou DATA/HORA posterior ao lancamento
//                                                                   -> nao participa, nao plota
//    DUPLICIDADE_ORIGEM .. retransmissao do interfaceamento: mesma
//                          amostra, item, instante e valor de um ID
//                          menor (ID_RELACIONADO)                   -> nao participa, nao plota
//    CONFLITO_MANUAL ..... resultado manual que coincide com um do
//                          interfaceamento (mesma chave logica, ate
//                          TOLERANCIA_CONFLITO_MIN) ou repete outro
//                          manual -- nunca duplica em silencio      -> nao participa, nao plota
//    SEM_DATA_HORA ....... interfaceamento sem DATA_HORA (QA E09)  -> nao participa, nao plota, sem RUN
//    SEM_VALOR ........... sem numero (flag '----' do XN-1000)      -> nao participa, nao plota
//    INATIVADO ........... ID na tblInativacao_NaoConformes COM o   -> nao participa;
//                          ANALITO da linha igual ao do resultado
//                          (ADR-070; analito vazio ou de outro ID:
//                          a inativacao NAO vale e o QA acusa E10/E11)
//                            REGISTRAR - LJ marcado (padrao) ...... X_VERMELHO
//                            REGISTRAR - LJ desmarcado ............ NAO_PLOTAR
//    ATIVO ............... todo o resto (interface ou manual)       -> participa, ponto NORMAL
//
//  JUSTIFICATIVA (ADR-070): o MOTIVO escrito na propria linha da aba Inativar OU um
//  comentario na aba COMENTARIOS_TECNICOS. Inativado sem nenhum dos dois = GOVERNANCA ERRO (QA E01).
//
//  CORRIDA e RUN
//    O instante de cada nivel e diferente (o XN-1000 mede os tres niveis com
//    ~1 min de intervalo; na Dimension cada analito e liberado na sua hora) e
//    ha duas ou mais corridas por dia. Por isso a corrida NAO e o dia:
//      1. por (setor, equipamento, matriz, lote, analito, dia), os resultados
//         em ordem de DATA_HORA formam um BLOCO enquanto o intervalo entre um e
//         o seguinte nao passa de GAP_CORRIDA_MIN -- e o bloco que junta os
//         niveis de uma mesma corrida;
//      2. dentro do bloco, o 2o resultado do MESMO nivel (replica ou repeticao
//         imediata) abre a posicao seguinte, para que todo resultado tenha o
//         seu ponto -- nunca dois pontos do mesmo nivel na mesma posicao;
//      3. CORRIDA_NO_DIA = ordem de (bloco, posicao) no dia;
//         RUN = yymmdd*100 + CORRIDA_NO_DIA  (ex. 26093001).
//    O RUN depende so dos resultados do dia -- inativar ou reativar NAO muda
//    numeracao (inativados contam na formacao da corrida; so retransmissoes e
//    conflitos, que nao sao eventos reais, ficam fora).
//
//    Implementacao SEM ESTADO ENCADEADO: cada passo marca onde comeca um
//    segmento (comparando a linha com a anterior, acesso O(1) em lista
//    bufferizada) e completa para baixo com Table.FillDown. A versao anterior acumulava o
//    contador num List.Generate; os campos de registro do M sao preguicosos e
//    cada contador apontava para o anterior -- uma cadeia de ~44 mil
//    avaliacoes aninhadas que estourava a pilha quando a ordenacao seguinte
//    pedia o ultimo valor primeiro.
// =====================================================================
let
    C = CFG_INTEGRACAO,
    GAP = Number.From(C[GAP_CORRIDA_MIN]),
    TOL = Number.From(C[TOLERANCIA_CONFLITO_MIN]),
    PREFIXO = Text.Upper(Text.Trim(Text.From(C[PREFIXO_ID]))),          // D15: "hem" no CFG nao vira outro setor
    SIM = "SIM", NAO = "NÃO",

    Txt = (x as any) as nullable text => if x = null then null else
            let t = Text.Trim(Text.From(x)) in if t = "" then null else t,
    SoDigitos = (t as text) as logical => t <> "" and Text.Length(Text.Select(t, {"0".."9"})) = Text.Length(t),
    // ID digitado: maiusculas, sem espacos; numero puro = ID do interfaceamento deste setor;
    // MAN_7 / man_0007 -> MAN_0007 (manual)
    // D14 (QA-ETL-001): tambem sem NBSP (colado de e-mail/web) e sem zeros a esquerda no numero
    // ("00123" e "HEM-00123" = HEM-123). Mesma regra no QA_INTEGRACAO e no VBA (mIntegracao.NormalizarId).
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
    Logico = (v as any, padrao as logical) as logical =>
        if v = null then padrao
        else if v is logical then v
        else if v is number then v <> 0
        else let t = Text.Upper(Text.Trim(Text.From(v))) in
             if t = "" then padrao else List.Contains({"SIM", "S", "TRUE", "VERDADEIRO", "V", "X", "1"}, t),
    // D03 (QA-ETL-001): tabela de ENTRADA obrigatoria (criada pelo instalador) ausente ou renomeada PARA a
    // atualizacao. Antes virava tabela vazia: todo inativado voltava a ATIVO e entrava na estatistica sem erro.
    Ler = (nome as text, cols as list) as table =>
        let t = try Excel.CurrentWorkbook(){[Name = nome]}[Content]
                otherwise error Error.Record("TABELA_AUSENTE", "Tabela obrigatoria ausente ou renomeada: " & nome &
                                             " (a atualizacao parou; nada foi alterado)", nome)
        in Table.SelectColumns(t, cols, MissingField.UseNull),

    // ------------------------------------------------------------ de/para
    DP0 = Ler("tblDeParaAnalitos", {"MATRIZ", "ANALITO_ORIGEM", "ANALITO_QCINI", "BLOCO", "ORDEM"}),
    DP = Table.Buffer(Table.SelectRows(
            Table.TransformColumns(DP0, {{"MATRIZ", Txt}, {"ANALITO_ORIGEM", Txt}, {"ANALITO_QCINI", Txt}, {"BLOCO", Txt}}),
            each [ANALITO_ORIGEM] <> null)),
    ParaQcini = (matriz as nullable text, origem as nullable text) as nullable text =>
        let r = Table.SelectRows(DP, each [ANALITO_ORIGEM] = origem and ([MATRIZ] = null or [MATRIZ] = matriz))
        in if Table.IsEmpty(r) then null else r{0}[ANALITO_QCINI],
    ParaOrigem = (matriz as nullable text, qcini as nullable text) as nullable record =>
        let r = Table.SelectRows(DP, each [ANALITO_QCINI] = qcini and ([MATRIZ] = null or [MATRIZ] = matriz))
        in if Table.IsEmpty(r) then null else r{0},
    Cadastrados = List.Buffer(List.RemoveNulls(List.Distinct(DP[ANALITO_QCINI]))),

    // ------------------------------------------------------------ interfaceamento
    // colunas do DB_SEAC; as que um setor nao tem (Hematologia: MATRIZ, BLOCO,
    // PRIMEIRO_DO_DIA) entram com o padrao do CFG / vazias
    ColsRec = {"ID_REGISTRO", "EQUIPAMENTO", "MATRIZ", "BLOCO", "LOTE", "NIVEL", "NIVEL_DESC", "DATA", "DATA_HORA",
               "ANALITO", "PRIMEIRO_DO_DIA", "VALOR", "FLAG", "ID_ORIGEM", "ID_AMOSTRA", "ITEM_ID", "UNIDADE", "RECEBIDO_EM"},
    R0 = Excel.CurrentWorkbook(){[Name = "tblDB_Recebimento"]}[Content],
    // Buffer: R1 e lido por Mapa, R2 e NivelDesc (3 leituras da tabela de 44 mil linhas)
    R1 = Table.Buffer(Table.TransformColumnTypes(
        Table.RenameColumns(
            Table.ReplaceValue(Table.Distinct(Table.SelectColumns(Table.SelectRows(R0, each [ID_REGISTRO] <> null), ColsRec,
                                              MissingField.UseNull), {"ID_REGISTRO"}),          // D01: defesa, 1 linha por ID
                               null, C[MATRIZ_PADRAO], Replacer.ReplaceValue, {"MATRIZ"}),
            {{"ANALITO", "ANALITO_ORIGEM"}}), {
        {"ID_REGISTRO", type text}, {"ID_ORIGEM", Int64.Type},
        {"EQUIPAMENTO", type text}, {"MATRIZ", type text}, {"BLOCO", type nullable text},
        {"LOTE", type text}, {"NIVEL", Int64.Type}, {"NIVEL_DESC", type text},
        {"DATA", type date}, {"DATA_HORA", type datetime}, {"ANALITO_ORIGEM", type text},
        {"PRIMEIRO_DO_DIA", type nullable text}, {"VALOR", type nullable number}, {"FLAG", type nullable text},
        {"UNIDADE", type nullable text}, {"ID_AMOSTRA", type text}, {"ITEM_ID", Int64.Type},
        {"RECEBIDO_EM", type datetime}})),
    Mapa = Table.Buffer(Table.AddColumn(
                Table.Distinct(Table.SelectColumns(R1, {"MATRIZ", "ANALITO_ORIGEM"})),
                "ANALITO_QCINI", each ParaQcini([MATRIZ], [ANALITO_ORIGEM]), type text)),
    // Juncao PLANA (Table.Join) com as chaves do lado direito renomeadas: o par
    // NestedJoin + Expand monta uma tabela aninhada por linha e custava caro em
    // 44 mil linhas.
    R2 = Table.RemoveColumns(Table.Join(R1, {"MATRIZ", "ANALITO_ORIGEM"},
            Table.RenameColumns(Mapa, {{"MATRIZ", "_M_MZ"}, {"ANALITO_ORIGEM", "_M_AO"}}), {"_M_MZ", "_M_AO"}, JoinKind.LeftOuter),
            {"_M_MZ", "_M_AO"}),
    Interface = Table.AddColumn(Table.AddColumn(Table.AddColumn(R2,
        "ANALITO", each if [ANALITO_QCINI] <> null then [ANALITO_QCINI] else [ANALITO_ORIGEM], type text),
        "ORIGEM_RESULTADO", each "INTERFACEAMENTO", type text),
        "_LINHA_MANUAL", each null, Int64.Type),

    // retransmissao: mesma amostra, item, instante, analito, nivel e valor (ID SIPEC diferente).
    // Replica REAL (outro ITEM_ID, outro instante ou outro valor) nao cai aqui: e resultado.
    // Agrupamento NATIVO pelas colunas tipadas (sem montar texto linha a linha).
    // Nulos agrupam e casam com nulos (FLAG vazia, ITEM_ID ausente).
    ChaveDup = {"EQUIPAMENTO", "MATRIZ", "LOTE", "NIVEL", "ANALITO_ORIGEM", "DATA_HORA", "ID_AMOSTRA", "ITEM_ID", "VALOR", "FLAG"},
    ChaveDupR = List.Transform(ChaveDup, each "_K_" & _),
    // so as chaves que REPETEM (retransmissao e rara): o lado direito da juncao fica pequeno
    Canon = Table.RenameColumns(
                Table.SelectRows(Table.Group(Interface, ChaveDup, {{"_ID_CANON", each List.Min([ID_ORIGEM]), Int64.Type},
                                                                   {"_N", each Table.RowCount(_), Int64.Type}}), each [_N] > 1),
                List.Zip({ChaveDup, ChaveDupR})),
    IJ = Table.RemoveColumns(Table.Join(Interface, ChaveDup, Canon, ChaveDupR, JoinKind.LeftOuter), ChaveDupR),
    InterfaceD = Table.AddColumn(Table.AddColumn(IJ,
        "_DUP", each [_N] <> null and [ID_ORIGEM] <> [_ID_CANON], type logical),
        "_ID_DUP_DE", each if [_N] <> null and [ID_ORIGEM] <> [_ID_CANON] then PREFIXO & "-" & Text.From([_ID_CANON]) else null, type text),

    // ------------------------------------------------------------ manuais (staging -> append)
    // Mesmo schema logico do interfaceamento: DATA + HORA formam DATA_HORA; o RUN
    // nao e digitado -- e calculado abaixo pela MESMA regra dos automaticos.
    M0 = Ler("tblResultados_Manuais", {"ID_REGISTRO", "DATA", "HORA", "EQUIPAMENTO", "MATRIZ", "LOTE", "NIVEL",
                                       "ANALITO", "RESULTADO", "UNIDADE", "MOTIVO", "USUARIO", "REGISTRADO_EM"}),
    M1 = Table.SelectRows(Table.AddIndexColumn(M0, "_LINHA_MANUAL", 1, 1, Int64.Type), each
            List.NonNullCount(List.Transform({[ID_REGISTRO], [DATA], [HORA], [LOTE], [NIVEL], [ANALITO], [RESULTADO]}, Txt)) > 0),
    NivelDesc = Table.Buffer(Table.Distinct(Table.SelectColumns(R1, {"NIVEL", "NIVEL_DESC"}), {"NIVEL"})),
    M2 = Table.AddColumn(M1, "_r", each
        let
            id = NormId([ID_REGISTRO]),
            dt = try Date.From([DATA]) otherwise null,
            hr = try (if [HORA] is number then Time.From(Number.Mod([HORA], 1)) else Time.From([HORA])) otherwise null,
            eq = Txt([EQUIPAMENTO]) ?? C[EQUIPAMENTO_PADRAO],
            mz = Txt([MATRIZ]) ?? C[MATRIZ_PADRAO],
            lt = let v = [LOTE] in if v is number then Number.ToText(v, "0") else Txt(v),
            nv = try Int64.From([NIVEL]) otherwise null,
            an = Txt([ANALITO]),
            // D07: texto com ponto e sem virgula ("1.5") e AMBIGUO em pt-BR (ponto = milhar: viraria 15) -- nunca converte
            rtx = if [RESULTADO] is number then null else Txt([RESULTADO]),
            amb = rtx <> null and Text.Contains(rtx, ".") and not Text.Contains(rtx, ","),
            vl = if amb then null
                 else try (if [RESULTADO] is number then [RESULTADO] else Number.FromText(Text.Trim(Text.From([RESULTADO])), "pt-BR")) otherwise null,
            orig = ParaOrigem(mz, an),
            nd = let p = List.PositionOf(NivelDesc[NIVEL], nv) in if p >= 0 then NivelDesc{p}[NIVEL_DESC] else (if nv = null then null else Text.From(nv)),
            dh = if dt = null or hr = null then null else DateTime.From(dt) + (hr - #time(0, 0, 0)),
            // gate 4.4 (data futura): o resultado nao pode ser posterior ao proprio
            // lancamento. Compara com REGISTRADO_EM (carimbo do evento de digitacao),
            // nao com "agora": a classificacao nao muda sozinha com o passar do tempo
            reg = (try DateTime.From([REGISTRADO_EM]) otherwise null) ?? DateTime.LocalNow(),
            faltando = List.Select({
                {"ID_REGISTRO (padrao MAN_0001)", id = null or not (Text.StartsWith(id, "MAN_") and SoDigitos(Text.Middle(id, 4)))},
                {"DATA", dt = null}, {"HORA", hr = null}, {"LOTE", lt = null},
                {"DATA/HORA posterior ao lancamento", dh <> null and dh > reg},
                {"NIVEL", nv = null or nv < 1 or nv > Number.From(C[NIVEIS])},
                {"ANALITO", an = null}, {"RESULTADO ambiguo (use virgula decimal: 1,5)", amb},
                {"RESULTADO", vl = null and not amb}}, each _{1})
        in [
            ID_REGISTRO = if id <> null and Text.StartsWith(id, "MAN_") and SoDigitos(Text.Middle(id, 4)) then id
                          else "MAN_INVALIDO_L" & Text.From([_LINHA_MANUAL]), ID_ORIGEM = null,
            ID_DIGITADO = id,
            EQUIPAMENTO = eq, MATRIZ = mz,
            BLOCO = if orig = null then null else orig[BLOCO], LOTE = lt, NIVEL = nv, NIVEL_DESC = nd,
            DATA = dt, DATA_HORA = dh,
            ANALITO_ORIGEM = if orig = null then an else orig[ANALITO_ORIGEM], PRIMEIRO_DO_DIA = null,
            VALOR = vl, FLAG = null, UNIDADE = Txt([UNIDADE]), ID_AMOSTRA = null, ITEM_ID = null,
            RECEBIDO_EM = try DateTime.From([REGISTRADO_EM]) otherwise null,
            ANALITO_QCINI = if List.Contains(Cadastrados, an) then an else null, ANALITO = an,
            ORIGEM_RESULTADO = "MANUAL", _LINHA_MANUAL = [_LINHA_MANUAL],
            _MANUAL_FALTA = Text.Combine(List.Transform(faltando, each _{0}), ", "),
            MOTIVO_MANUAL = Txt([MOTIVO]), USUARIO_MANUAL = Txt([USUARIO])
        ]),
    // lista de colunas explicita: sem linha manual, a tabela vazia ainda tem as colunas
    ColsManual = {"ID_REGISTRO", "ID_ORIGEM", "EQUIPAMENTO", "MATRIZ", "BLOCO", "LOTE", "NIVEL", "NIVEL_DESC",
                  "DATA", "DATA_HORA", "ANALITO_ORIGEM", "PRIMEIRO_DO_DIA", "VALOR", "FLAG", "UNIDADE", "ID_AMOSTRA", "ITEM_ID",
                  "RECEBIDO_EM", "ANALITO_QCINI", "ANALITO", "ORIGEM_RESULTADO", "_LINHA_MANUAL", "_MANUAL_FALTA",
                  "MOTIVO_MANUAL", "USUARIO_MANUAL", "ID_DIGITADO"},
    Manuais0 = Table.FromRecords(M2[_r], ColsManual, MissingField.UseNull),
    // ID repetido entre manuais: vale a 1a linha; as seguintes ganham sufixo ~L<n>
    // (o ID_REGISTRO da DB_CQ_FINAL e UNICO sempre) e viram CONFLITO_MANUAL
    // D08 (QA-ETL-001): com o mesmo MAN_ numa linha incompleta e numa completa, vale a COMPLETA, qualquer que
    // seja a ordem fisica (antes a incompleta de cima derrubava a boa para CONFLITO)
    PrimeiroId = Table.Group(Manuais0, {"ID_REGISTRO"}, {{"_PRIM_ID", each
                    let ok = List.Min(Table.SelectRows(_, each [_MANUAL_FALTA] = "")[_LINHA_MANUAL]) in ok ?? List.Min([_LINHA_MANUAL]),
                    Int64.Type}}),
    Manuais1 = Table.ExpandTableColumn(Table.NestedJoin(Manuais0, {"ID_REGISTRO"}, PrimeiroId, {"ID_REGISTRO"}, "_a", JoinKind.LeftOuter), "_a", {"_PRIM_ID"}),
    Manuais = Table.RemoveColumns(Table.AddColumn(Table.RenameColumns(Manuais1, {{"ID_REGISTRO", "_ID0"}}), "ID_REGISTRO", each
                if [_PRIM_ID] = [_LINHA_MANUAL] then [_ID0] else [_ID0] & "~L" & Text.From([_LINHA_MANUAL]), type text), {"_ID0"}),
    ManuaisOk = Table.SelectRows(Manuais, each [_MANUAL_FALTA] = ""),

    // manual x manual: o mesmo ID ou o mesmo resultado (chave + instante) repetido
    KMan = (r as record) as text => Text.Combine({r[EQUIPAMENTO], r[MATRIZ], r[LOTE], Text.From(r[NIVEL]), r[ANALITO],
                                                  DateTime.ToText(r[DATA_HORA], "yyyyMMddHHmmss")}, "|"),
    MK = Table.AddColumn(ManuaisOk, "_KMAN", each KMan(_), type text),
    PrimeiroK = Table.Group(MK, {"_KMAN"}, {{"_PRIM_K", each List.Min([_LINHA_MANUAL]), Int64.Type},
                                            {"_ID_PRIM_K", each Table.Sort(_, {"_LINHA_MANUAL"}){0}[ID_REGISTRO], type text}}),
    MM = Table.ExpandTableColumn(Table.NestedJoin(MK, {"_KMAN"}, PrimeiroK, {"_KMAN"}, "_b", JoinKind.LeftOuter), "_b", {"_PRIM_K", "_ID_PRIM_K"}),

    // manual x interfaceamento: mesma chave logica no mesmo dia, ate TOL minutos -> CONFLITO
    // (o interfaceamento prevalece; o manual fica guardado, fora do calculo, e vai ao QA)
    // Buffer: sem ele cada linha manual reavaliava a juncao sobre ~44 mil linhas (medido: 130 manuais = 6 min)
    InterfaceValida = Table.Buffer(Table.SelectColumns(
            Table.SelectRows(InterfaceD, each not [_DUP] and [VALOR] <> null and [DATA_HORA] <> null),
            {"ID_REGISTRO", "EQUIPAMENTO", "MATRIZ", "LOTE", "NIVEL", "ANALITO", "DATA", "DATA_HORA"})),
    ChaveMI = {"EQUIPAMENTO", "MATRIZ", "LOTE", "NIVEL", "ANALITO"},
    // D09c (QA-ETL-001): a tolerancia e medida no DATA_HORA e atravessa a meia-noite -- um manual as 00:02 e o
    // interfaceamento as 23:58 da vespera sao o mesmo resultado (antes os dois participavam). Candidatos: o proprio
    // dia, a vespera e o dia seguinte (TOL < 24 h), cada um por juncao de chave + DATA (hash, pequena).
    MMd = Table.AddColumn(Table.AddColumn(MM, "_DANT", each try Date.AddDays([DATA], -1) otherwise null, type nullable date),
                          "_DSEG", each try Date.AddDays([DATA], 1) otherwise null, type nullable date),
    MI = if Table.IsEmpty(MM) then Table.AddColumn(MM, "_i", each #table({"ID_REGISTRO", "DATA_HORA"}, {}))
         else Table.Buffer(Table.RemoveColumns(Table.AddColumn(
                Table.NestedJoin(Table.NestedJoin(Table.NestedJoin(MMd,
                    ChaveMI & {"DATA"}, InterfaceValida, ChaveMI & {"DATA"}, "_i0", JoinKind.LeftOuter),
                    ChaveMI & {"_DANT"}, InterfaceValida, ChaveMI & {"DATA"}, "_ia", JoinKind.LeftOuter),
                    ChaveMI & {"_DSEG"}, InterfaceValida, ChaveMI & {"DATA"}, "_is", JoinKind.LeftOuter),
                "_i", each Table.Combine({[_i0], [_ia], [_is]})), {"_i0", "_ia", "_is", "_DANT", "_DSEG"})),
    MI2 = Table.AddColumn(MI, "_conf", each
        let
            cand = Table.AddColumn(Table.SelectRows([_i], (x) => x[DATA_HORA] <> null and [DATA_HORA] <> null),
                                   "_d", (x) => Number.Abs(Duration.TotalMinutes(x[DATA_HORA] - [DATA_HORA]))),
            perto = if Table.IsEmpty(cand) then null else Table.Min(cand, "_d")
        in
            if [_PRIM_ID] <> [_LINHA_MANUAL] then [m = "ID_REGISTRO manual repetido (vale a 1a linha)", id = null]
            else if [_PRIM_K] <> [_LINHA_MANUAL] then [m = "resultado manual repetido (mesma chave e instante)", id = [_ID_PRIM_K]]
            else if perto <> null and perto[_d] <= TOL then
                [m = "coincide com resultado do interfaceamento (" & Number.ToText(perto[_d], "0") & " min)", id = perto[ID_REGISTRO]]
            else null),
    ManuaisC = Table.RemoveColumns(MI2, {"_i", "_KMAN", "_PRIM_ID", "_PRIM_K", "_ID_PRIM_K"}, MissingField.Ignore),
    ManuaisFalha = Table.SelectRows(Manuais, each [_MANUAL_FALTA] <> ""),

    // ------------------------------------------------------------ uniao (UNION ALL)
    Comuns = {"ID_REGISTRO", "ID_ORIGEM", "EQUIPAMENTO", "MATRIZ", "BLOCO", "LOTE", "NIVEL", "NIVEL_DESC",
              "DATA", "DATA_HORA", "ANALITO", "ANALITO_ORIGEM", "ANALITO_QCINI", "PRIMEIRO_DO_DIA", "VALOR", "FLAG", "UNIDADE",
              "ID_AMOSTRA", "ITEM_ID", "RECEBIDO_EM", "ORIGEM_RESULTADO", "_LINHA_MANUAL",
              "_DUP", "_ID_DUP_DE", "_MANUAL_FALTA", "_conf", "MOTIVO_MANUAL", "USUARIO_MANUAL"},
    // Buffer: o UNICO buffer largo. U alimenta a corrida (tabela estreita) e a
    // saida; sem ele a uniao (leitura + duplicidade) seria avaliada duas vezes.
    U = Table.Buffer(Table.Combine({
        Table.SelectColumns(Table.AddColumn(InterfaceD, "_MANUAL_FALTA", each ""), Comuns, MissingField.UseNull),
        Table.SelectColumns(Table.AddColumn(ManuaisC, "_DUP", each false), Comuns, MissingField.UseNull),
        Table.SelectColumns(Table.AddColumn(ManuaisFalha, "_DUP", each false), Comuns, MissingField.UseNull)})),

    // ------------------------------------------------------------ inativacao (soft-delete) e comentarios
    // ADR-070: a linha da Inativar traz ID + ANALITO (conferencia) + MOTIVO (justificativa). O ANALITO tem de ser
    // o do resultado: ID de outro analito (digitado errado) ou sem analito NAO inativa nada (QA E10/E11).
    // Tabela ainda no layout antigo (sem a coluna MOTIVO; antes da migracao do instalar_adr070): regra antiga,
    // so pelo ID -- sem isso, reinstalar a consulta antes da migracao reativaria todo inativado (e o Audit_Log
    // registraria reativacoes que ninguem fez).
    InatLegado = not (try List.Contains(Table.ColumnNames(Excel.CurrentWorkbook(){[Name = "tblInativacao_NaoConformes"]}[Content]),
                                        "MOTIVO") otherwise true),
    AnKey = (x as any) as nullable text => let t = Txt(x) in if t = null then null else Text.Upper(t),
    I0 = Ler("tblInativacao_NaoConformes", {"ID_REGISTRO", "ANALITO", "REGISTRAR - LJ", "MOTIVO", "DATA_INATIVACAO", "USUARIO"}),
    I1 = Table.SelectRows(Table.AddIndexColumn(
            Table.AddColumn(I0, "_IDN", each NormId([ID_REGISTRO]), type text), "_ord", 1, 1, Int64.Type),
            each [_IDN] <> null),
    I2 = Table.Distinct(Table.Buffer(Table.Sort(I1, {"_ord"})), {"_IDN"}),          // 1a ocorrencia vale; repeticao vai para o QA
    Inat = Table.SelectColumns(Table.AddColumn(Table.AddColumn(Table.AddColumn(Table.AddColumn(Table.AddColumn(I2,
            // D06 (QA-ETL-001): vazio = SIM so na linha nova (sem DATA_INATIVACAO); numa inativacao ja carimbada a
            // celula vazia aparece DESMARCADA e vale NAO. Valor nao reconhecido: NAO e ALERTA A08 no QA.
            "_LJ", each Logico([#"REGISTRAR - LJ"], (try DateTime.From([DATA_INATIVACAO]) otherwise null) = null), type logical),
            "_DTI", each try DateTime.From([DATA_INATIVACAO]) otherwise null, type nullable datetime),
            "_USRI", each Txt([USUARIO]), type nullable text),
            "_ANI", each if InatLegado then null else AnKey([ANALITO]), type nullable text),
            "_MOT", each if InatLegado then null else Txt([MOTIVO]), type nullable text),
            {"_IDN", "_LJ", "_DTI", "_USRI", "_ANI", "_MOT"}),

    K0 = Ler("tblComentariosTecnicos", {"ID_REGISTRO", "COMENTARIO_TECNICO", "DATA", "USUARIO"}),
    K1 = Table.SelectRows(Table.AddColumn(Table.AddColumn(K0, "_IDN", each NormId([ID_REGISTRO]), type text),
            "_TXT", each Txt([COMENTARIO_TECNICO]), type text), each [_IDN] <> null and [_TXT] <> null),
    Coment = Table.Group(K1, {"_IDN"}, {{"_COMENT", each Text.Combine([_TXT], " | "), type text}}),

    // LEFT OUTER: todo resultado continua; quem casa com a inativacao so muda de estado
    UI0 = Table.Join(U, {"ID_REGISTRO"}, Table.RenameColumns(Inat, {{"_IDN", "_INAT_ID"}}), {"_INAT_ID"}, JoinKind.LeftOuter),
    // ADR-070: a inativacao so VALE com o analito informado igual ao do resultado (layout antigo: so o ID).
    // A que nao vale nao deixa rastro na linha (data, usuario, motivo): o QA conta por que (E10/E11).
    UI = Table.AddColumn(UI0, "_INAT_OK", each [_INAT_ID] <> null and
            (InatLegado or ([_ANI] <> null and [_ANI] = AnKey([ANALITO]))), type logical),
    UK = Table.RemoveColumns(Table.Join(UI, {"ID_REGISTRO"}, Table.RenameColumns(Coment, {{"_IDN", "_IDK"}}), {"_IDK"},
                                        JoinKind.LeftOuter), {"_IDK"}),

    // ------------------------------------------------------------ classificacao
    Classif = Table.AddColumn(UK, "_cls", each
        let
            inat = [_INAT_OK],
            st = if [_MANUAL_FALTA] <> "" then "MANUAL_INCOMPLETO"
                 else if [_DUP] then "DUPLICIDADE_ORIGEM"
                 else if [_conf] <> null then "CONFLITO_MANUAL"
                 else if [DATA_HORA] = null then "SEM_DATA_HORA"            // D16: fora da corrida e do calculo; QA E09
                 else if [VALOR] = null then "SEM_VALOR"
                 else if inat then "INATIVADO"
                 else "ATIVO",
            lj = st = "ATIVO" or (st = "INATIVADO" and [_LJ]),
            motivo = if st = "MANUAL_INCOMPLETO" then "campos obrigatorios ausentes/invalidos: " & [_MANUAL_FALTA]
                     else if st = "DUPLICIDADE_ORIGEM" then "retransmissao do interfaceamento"
                     else if st = "CONFLITO_MANUAL" then [_conf][m]
                     else if st = "SEM_DATA_HORA" then "resultado do interfaceamento sem DATA_HORA (fora da corrida)"
                     else if st = "SEM_VALOR" then "resultado sem valor numerico (flag " & ([FLAG] ?? "vazia") & ")"
                     else null
        in [
            STATUS_ANALITICO = st,
            PARTICIPA_ESTATISTICA = if st = "ATIVO" then SIM else NAO,
            REGISTRAR_RESULTADO_NO_LJ = if lj then SIM else NAO,
            TIPO_PLOTAGEM_LJ = if st = "ATIVO" then "NORMAL" else if lj then "X_VERMELHO" else "NAO_PLOTAR",
            MOTIVO_EXCLUSAO_AUTOMATICA = motivo,
            ID_RELACIONADO = if st = "DUPLICIDADE_ORIGEM" then [_ID_DUP_DE] else if st = "CONFLITO_MANUAL" then [_conf][id] else null,
            INATIVACAO_REGISTRADA = if inat then SIM else NAO,
            CORRIDA_REAL = List.Contains({"ATIVO", "INATIVADO", "SEM_VALOR"}, st)
        ]),
    Exp = Table.ExpandRecordColumn(Classif, "_cls", {"STATUS_ANALITICO", "PARTICIPA_ESTATISTICA",
            "REGISTRAR_RESULTADO_NO_LJ", "TIPO_PLOTAGEM_LJ", "MOTIVO_EXCLUSAO_AUTOMATICA", "ID_RELACIONADO",
            "INATIVACAO_REGISTRADA", "CORRIDA_REAL"}),
    // ADR-070: justificativa = MOTIVO da linha da Inativar (so da inativacao que vale) OU comentario tecnico.
    // Data, usuario e motivo da inativacao so aparecem na linha quando a inativacao vale.
    ComJust = Table.AddColumn(Table.AddColumn(Table.AddColumn(Table.AddColumn(Table.AddColumn(Exp,
        "_MOTV", each if [_INAT_OK] then [_MOT] else null, type nullable text),
        "_DTIV", each if [_INAT_OK] then [_DTI] else null, type nullable datetime),
        "_USRIV", each if [_INAT_OK] then [_USRI] else null, type nullable text),
        "TEM_JUSTIFICATIVA", each if [_COMENT] <> null or [_MOTV] <> null then SIM else NAO, type text),
        "GOVERNANCA", each if [STATUS_ANALITICO] = "INATIVADO" and [_COMENT] = null and [_MOTV] = null
                           then "ERRO DE GOVERNANÇA: INATIVADO SEM JUSTIFICATIVA TÉCNICA" else "OK", type text),

    // ------------------------------------------------------------ corrida e RUN
    // Calculada numa tabela ESTREITA (so as chaves) e devolvida pelo ID_REGISTRO:
    // ordenar e bufferizar 45 colunas x 44 mil linhas tres vezes estourava a
    // memoria do conteiner do Power Query e o tempo crescia mais que a base.
    // Mesma condicao de CORRIDA_REAL: fora so incompleto, retransmissao e conflito.
    // D16: resultado sem DATA_HORA nao entra na corrida (quebraria o corte por GAP); o QA reporta (E09)
    Reais = Table.SelectColumns(Table.SelectRows(U, each [_MANUAL_FALTA] = "" and not [_DUP] and [_conf] = null and [DATA_HORA] <> null),
                {"ID_REGISTRO", "EQUIPAMENTO", "MATRIZ", "LOTE", "ANALITO", "DATA", "DATA_HORA", "NIVEL", "ITEM_ID"}),

    // Cada passo: ordena, poe ao lado de cada linha os valores da linha ANTERIOR
    // (coluna deslocada, montada de uma vez a partir da lista -- leitura
    // sequencial, sem acesso por posicao), marca a linha que ABRE um segmento
    // com o proprio indice e completa para baixo (Table.FillDown). Tudo nativo
    // e linear; nenhum valor depende do anterior por avaliacao encadeada.
    ComAnterior = (t as table, colunas as list) as table =>
        Table.FromColumns(Table.ToColumns(t) & List.Transform(colunas, (c) => {null} & List.RemoveLastN(Table.Column(t, c), 1)),
                          Table.ColumnNames(t) & List.Transform(colunas, (c) => "_ANT_" & c)),

    // (1) bloco: muda quando muda o grupo do dia ou o intervalo passa de GAP minutos
    Ordem1 = Table.Buffer(Table.AddIndexColumn(Table.Sort(
                Table.AddColumn(Reais, "_GK", each Text.Combine({[EQUIPAMENTO], [MATRIZ], [LOTE], [ANALITO],
                                                                 Date.ToText([DATA], "yyyyMMdd")}, "|"), type text),
                {{"_GK", Order.Ascending}, {"DATA_HORA", Order.Ascending}, {"NIVEL", Order.Ascending},
                 {"ITEM_ID", Order.Ascending}, {"ID_REGISTRO", Order.Ascending}}), "_i1", 0, 1, Int64.Type)),
    T1 = Table.FillDown(Table.AddColumn(ComAnterior(Ordem1, {"_GK", "DATA_HORA"}), "_BL", each
            if [_ANT__GK] = null or [_GK] <> [_ANT__GK] or Duration.TotalMinutes([DATA_HORA] - [_ANT_DATA_HORA]) > GAP
            then [_i1] else null, Int64.Type), {"_BL"}),

    // (2) posicao dentro de (bloco, nivel): o 2o resultado do mesmo nivel abre posicao nova
    Ordem2 = Table.Buffer(Table.AddIndexColumn(Table.Sort(
                Table.SelectColumns(T1, {"ID_REGISTRO", "_GK", "_BL", "NIVEL", "DATA", "DATA_HORA", "ITEM_ID"}),
                {{"_BL", Order.Ascending}, {"NIVEL", Order.Ascending}, {"DATA_HORA", Order.Ascending},
                 {"ITEM_ID", Order.Ascending}, {"ID_REGISTRO", Order.Ascending}}), "_i2", 0, 1, Int64.Type)),
    T2a = Table.FillDown(Table.AddColumn(ComAnterior(Ordem2, {"_BL", "NIVEL"}), "_I2INI", each
            if [_ANT__BL] = null or [_BL] <> [_ANT__BL] or [NIVEL] <> [_ANT_NIVEL] then [_i2] else null, Int64.Type), {"_I2INI"}),
    T2 = Table.Buffer(Table.SelectColumns(Table.AddColumn(T2a, "_POS", each [_i2] - [_I2INI] + 1, Int64.Type),
                {"ID_REGISTRO", "_GK", "_BL", "_POS", "DATA"})),

    // (3) corrida = ordem de (bloco, posicao) dentro do dia: numera os segmentos
    Seg = Table.Buffer(Table.AddIndexColumn(Table.Sort(Table.Distinct(Table.SelectColumns(T2, {"_GK", "_BL", "_POS"})),
                {{"_GK", Order.Ascending}, {"_BL", Order.Ascending}, {"_POS", Order.Ascending}}), "_s", 0, 1, Int64.Type)),
    SegC = Table.AddColumn(Table.FillDown(Table.AddColumn(ComAnterior(Seg, {"_GK"}), "_SINI", each
                if [_ANT__GK] = null or [_GK] <> [_ANT__GK] then [_s] else null, Int64.Type), {"_SINI"}),
            "CORRIDA_NO_DIA", each [_s] - [_SINI] + 1, Int64.Type),
    ComCorrida = Table.Join(T2, {"_GK", "_BL", "_POS"},
                    Table.RenameColumns(Table.SelectColumns(SegC, {"_GK", "_BL", "_POS", "CORRIDA_NO_DIA"}),
                                        {{"_GK", "_S_GK"}, {"_BL", "_S_BL"}, {"_POS", "_S_POS"}}),
                    {"_S_GK", "_S_BL", "_S_POS"}, JoinKind.LeftOuter),

    RunPorId = Table.Buffer(Table.SelectColumns(Table.AddColumn(ComCorrida, "RUN", each
        if [CORRIDA_NO_DIA] > 99 then null else Number.From(Date.ToText([DATA], "yyMMdd")) * 100 + [CORRIDA_NO_DIA], Int64.Type),
        {"ID_REGISTRO", "CORRIDA_NO_DIA", "RUN"})),
    RunPorIdU = Table.Distinct(RunPorId, {"ID_REGISTRO"}),                    // D01: o join abaixo nunca multiplica
    Junto = Table.RemoveColumns(Table.Join(ComJust, {"ID_REGISTRO"}, Table.RenameColumns(RunPorIdU, {{"ID_REGISTRO", "_RID"}}),
                                           {"_RID"}, JoinKind.LeftOuter), {"_RID"}),

    // ------------------------------------------------------------ saida
    Final = Table.AddColumn(Table.AddColumn(Table.AddColumn(Table.AddColumn(Junto,
        "SETOR", each C[SETOR], type text),
        "HORA", each if [DATA_HORA] = null then null else DateTime.Time([DATA_HORA]), type nullable time),
        "ANALITO_CADASTRADO", each if [ANALITO_QCINI] <> null then SIM else NAO, type text),
        "RESULTADO", each [VALOR], type nullable number),
    Renom = Table.RenameColumns(Final, {{"_COMENT", "COMENTARIO_TECNICO"}, {"_DTIV", "DATA_INATIVACAO"},
                                        {"_USRIV", "USUARIO_INATIVACAO"}, {"_MOTV", "MOTIVO_INATIVACAO"}}),
    Saida = Table.SelectColumns(Renom, {
        "ID_REGISTRO", "ORIGEM_RESULTADO", "SETOR", "EQUIPAMENTO", "MATRIZ", "BLOCO",
        "DATA", "HORA", "DATA_HORA", "CORRIDA_NO_DIA", "RUN", "NIVEL", "NIVEL_DESC", "LOTE",
        "ANALITO", "ANALITO_ORIGEM", "ANALITO_CADASTRADO", "RESULTADO", "UNIDADE", "FLAG",
        "STATUS_ANALITICO", "PARTICIPA_ESTATISTICA", "REGISTRAR_RESULTADO_NO_LJ", "TIPO_PLOTAGEM_LJ",
        "INATIVACAO_REGISTRADA", "COMENTARIO_TECNICO", "TEM_JUSTIFICATIVA", "GOVERNANCA",
        "DATA_INATIVACAO", "USUARIO_INATIVACAO", "MOTIVO_INATIVACAO", "MOTIVO_EXCLUSAO_AUTOMATICA", "ID_RELACIONADO",
        "ID_ORIGEM", "ID_AMOSTRA", "ITEM_ID", "PRIMEIRO_DO_DIA", "MOTIVO_MANUAL", "USUARIO_MANUAL", "RECEBIDO_EM"}),
    Tipos = Table.TransformColumnTypes(Saida, {
        {"ID_REGISTRO", type text}, {"ORIGEM_RESULTADO", type text}, {"SETOR", type text}, {"EQUIPAMENTO", type text},
        {"MATRIZ", type text}, {"BLOCO", type nullable text}, {"DATA", type date}, {"HORA", type time},
        {"DATA_HORA", type datetime}, {"CORRIDA_NO_DIA", Int64.Type}, {"RUN", Int64.Type}, {"NIVEL", Int64.Type},
        {"NIVEL_DESC", type text}, {"LOTE", type text}, {"ANALITO", type text}, {"ANALITO_ORIGEM", type text},
        {"ANALITO_CADASTRADO", type text}, {"RESULTADO", type nullable number}, {"UNIDADE", type nullable text},
        {"FLAG", type nullable text}, {"STATUS_ANALITICO", type text}, {"PARTICIPA_ESTATISTICA", type text},
        {"REGISTRAR_RESULTADO_NO_LJ", type text}, {"TIPO_PLOTAGEM_LJ", type text}, {"INATIVACAO_REGISTRADA", type text},
        {"COMENTARIO_TECNICO", type nullable text}, {"TEM_JUSTIFICATIVA", type text}, {"GOVERNANCA", type text},
        {"DATA_INATIVACAO", type nullable datetime}, {"USUARIO_INATIVACAO", type nullable text},
        {"MOTIVO_INATIVACAO", type nullable text}, {"MOTIVO_EXCLUSAO_AUTOMATICA", type nullable text}, {"ID_RELACIONADO", type nullable text},
        {"ID_ORIGEM", Int64.Type}, {"ID_AMOSTRA", type nullable text}, {"ITEM_ID", Int64.Type},
        {"PRIMEIRO_DO_DIA", type nullable text}, {"MOTIVO_MANUAL", type nullable text}, {"USUARIO_MANUAL", type nullable text},
        {"RECEBIDO_EM", type nullable datetime}}),
    Ordenado = Table.Sort(Tipos, {{"DATA_HORA", Order.Ascending}, {"ANALITO", Order.Ascending},
                                  {"NIVEL", Order.Ascending}, {"ID_REGISTRO", Order.Ascending}})
in
    Ordenado
