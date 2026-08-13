# QC_INI — Power Query (M)

Cole cada bloco em **Página Inicial → Transformar dados → Nova Consulta → Consulta em branco → Editor Avançado**.

Defina primeiro o parâmetro do caminho: **Gerenciar Parâmetros → Novo**, nome
`pCaminhoQC`, tipo Texto, valor = caminho completo do `.xlsm`. Nunca deixe o
caminho embutido nas consultas — muda por máquina e por usuário.

---

## 1. `Fato_QC` — a tabela fato

Lê o `ListObject` **pelo nome**, não por faixa. É o que permite o banco crescer
sem tocar na consulta.

```m
let
    Fonte = Excel.Workbook(File.Contents(pCaminhoQC), null, true),
    Tabela = Fonte{[Item="tblBI_Fato", Kind="Table"]}[Data],
    Tipada = Table.TransformColumnTypes(Tabela, {
        {"ID_Result", type text}, {"ID_Corrida", type text},
        {"Data", type date}, {"Ano", Int64.Type}, {"Mes", Int64.Type},
        {"Trimestre", Int64.Type}, {"Competencia", type text},
        {"ID_Analito", type text}, {"Analito", type text},
        {"Area", type text}, {"Unidade", type text},
        {"ID_Lote", type text}, {"Lote", type text},
        {"Nivel", Int64.Type}, {"RUN", Int64.Type},
        {"Resultado", type number}, {"Status", type text}, {"Ativo", Int64.Type},
        {"Media_Alvo", type number}, {"DP_Alvo", type number}, {"Z", type number},
        {"Lim_m3s", type number}, {"Lim_m2s", type number}, {"Lim_m1s", type number},
        {"Lim_p1s", type number}, {"Lim_p2s", type number}, {"Lim_p3s", type number},
        {"CVtp_pct", type number}, {"BIAStp_pct", type number}, {"ETp_pct", type number},
        {"W_1_3s", Int64.Type}, {"W_2_2s", Int64.Type}, {"W_R_4s", Int64.Type},
        {"Veredito", type text}
    }),
    // Só o que está Ativo entra na análise. A exclusão lógica é decisão do
    // laboratório e tem trilha de auditoria; o BI a respeita, não a revisita.
    Ativos = Table.SelectRows(Tipada, each [Ativo] = 1),
    // Nível como texto rotulado: em visual, "1" vira número e o Power BI tenta
    // somá-lo.
    ComRotulo = Table.AddColumn(Ativos, "Nivel_Rotulo",
        each "N" & Text.From([Nivel]), type text)
in
    ComRotulo
```

---

## 2. `Dim_Data` — calendário

Cobre exatamente o intervalo dos dados. Um calendário fixo de 1900–2100 infla o
modelo sem servir a nada.

```m
let
    Min = List.Min(Fato_QC[Data]),
    Max = List.Max(Fato_QC[Data]),
    // até o fim do mês do último dado: evita mês parcial em visual mensal
    FimMes = Date.EndOfMonth(Max),
    N = Duration.Days(FimMes - Min) + 1,
    Dias = List.Dates(Min, N, #duration(1,0,0,0)),
    Tab = Table.FromList(Dias, Splitter.SplitByNothing(), {"Data"}),
    Tipada = Table.TransformColumnTypes(Tab, {{"Data", type date}}),
    Cols = Table.AddColumn(Tipada, "Ano", each Date.Year([Data]), Int64.Type),
    Cols2 = Table.AddColumn(Cols, "Mes", each Date.Month([Data]), Int64.Type),
    Cols3 = Table.AddColumn(Cols2, "Trimestre", each Date.QuarterOfYear([Data]), Int64.Type),
    Cols4 = Table.AddColumn(Cols3, "Competencia",
        each Date.ToText([Data], "yyyy-MM"), type text),
    Cols5 = Table.AddColumn(Cols4, "MesNome",
        each Date.ToText([Data], "MMM", "pt-BR"), type text),
    // chave numérica para ordenar MesNome corretamente no visual
    Cols6 = Table.AddColumn(Cols5, "AnoMes",
        each Date.Year([Data]) * 100 + Date.Month([Data]), Int64.Type)
in
    Cols6
```

Depois de carregar: **Modelagem → Marcar como tabela de datas** → `Data`.
E ordenar `MesNome` por `Mes` (Coluna → Classificar por coluna).

---

## 3. `Dim_Analito`

```m
let
    Base = Table.SelectColumns(Fato_QC, {"ID_Analito", "Analito", "Area", "Unidade"}),
    Unicos = Table.Distinct(Base, {"ID_Analito"})
in
    Unicos
```

## 4. `Dim_Lote`

```m
let
    Base = Table.SelectColumns(Fato_QC, {"ID_Lote", "Lote"}),
    Unicos = Table.Distinct(Base, {"ID_Lote"})
in
    Unicos
```

## 5. `Dim_Nivel`

```m
let
    Base = Table.SelectColumns(Fato_QC, {"Nivel", "Nivel_Rotulo"}),
    Unicos = Table.Distinct(Base, {"Nivel"}),
    Ord = Table.Sort(Unicos, {{"Nivel", Order.Ascending}})
in
    Ord
```

---

## 6. Relações a criar no modelo

| De | Para | Cardinalidade | Direção |
|---|---|---|---|
| `Fato_QC[Data]` | `Dim_Data[Data]` | muitos‑para‑um | simples |
| `Fato_QC[ID_Analito]` | `Dim_Analito[ID_Analito]` | muitos‑para‑um | simples |
| `Fato_QC[ID_Lote]` | `Dim_Lote[ID_Lote]` | muitos‑para‑um | simples |
| `Fato_QC[Nivel]` | `Dim_Nivel[Nivel]` | muitos‑para‑um | simples |

**Todas simples, nenhuma bidirecional.** Filtro cruzado bidirecional numa estrela
cria caminhos ambíguos e medidas que mudam de valor conforme o visual — e num
painel de CQ isso significa dois números diferentes para a mesma pergunta.

Ocultar do relatório: `ID_Result`, `ID_Corrida`, `ID_Analito`, `ID_Lote`. São
chaves; existem para o modelo, não para o usuário arrastar.

---

## 7. Atualização

- **Desktop:** Página Inicial → Atualizar.
- **Service:** o `.xlsm` está no OneDrive; usar o conector OneDrive/SharePoint
  dispensa gateway. Arquivo em disco local exigiria *On-premises data gateway
  (personal mode)*.
- O `.xlsm` precisa estar **fechado** ou salvo antes da atualização — o Power
  Query lê o arquivo em disco, não o que está aberto na memória do Excel.
