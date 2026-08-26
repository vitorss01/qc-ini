# Fase 2 — Diagnóstico da cadeia Excel → ETL → Power BI

Primeira entrega do §29. Levantamento feito sobre os artefatos do build limpo
(`418ff1a`, baseline `baseline-fase1`), lendo dado real e não apenas código.

---

## 1. Diagnóstico do ETL atual

**Um único projeto PBIP**, `PowerBI/QC_INI_Bioquimica.pbip`, que já atende os
dois produtos:

| Camada | Estado |
| --- | --- |
| Parâmetros | `pCaminhoQC` (Bioquímica) e `pCaminhoHema` (Hematologia) |
| Origem | `Excel.Workbook` → `tblBI_Fato` lido **pelo nome do ListObject** |
| Combinação | `Table.Combine({TabelaBio, TabelaHema})` no dado cru |
| Tipagem | uma única `TransformColumnTypes` sobre a tabela combinada |
| Tabelas | `Fato_QC`, `Dim_Analito`, `Dim_Data`, `Dim_Lote`, `Dim_Nivel` |
| Relacionamentos | 4, todos 1:N a partir das dimensões |
| Medidas | 114, todas em `Fato_QC` |
| Páginas | 5 — Visão Gerencial, Painel e Estatística de cada produto |

O desenho de fundo está **correto e vale preservar**: ler o ListObject pelo
nome (não por faixa), combinar antes de tipar, produto separado por coluna e
não por modelo. Nada disso precisa ser refeito.

O que está quebrado é o **contrato de dados** que trafega por ele.

---

## 2. Schema antigo

`mBI.Cab()` — 84 colunas. Grão: **um resultado de CQ**
(`ID_Result` = Lote × RUN × Nível × Analito).

Blocos: identificação (1–18), alvo e limites (19–27), especificação (28–30),
Westgard (31–37), proveniência (38–46), estatística observada (47–53),
vigência (54–60), margem ETp (61–65), EQA resumido (66–68), DPM/Yield (69–70),
plano de CQ (71–79), flags de recomendação (80–84).

---

## 3. Problemas encontrados

### P1 — Flags de recomendação FALSAS na Hematologia · **GRAVE, dado errado**

`mBI` linhas 915–919 consultam o plano com nomes fixos da Bioquímica:

```vba
saida(k, 81) = mPlanoQC.RegraNoPlano(sp, "2_2s")
saida(k, 83) = mPlanoQC.RegraNoPlano(sp, "4_1s")
saida(k, 84) = mPlanoQC.RegraNoPlano(sp, "8x")
```

`RegraNoPlano` compara **token exato**. O plano da Hematologia contém
`2of3_2s`, `3_1s`, `6x` — que nunca casam. Medido no `BI_Data` real:

| Analito | Σ plano | Regras recomendadas (texto) | Usar_1_3s/2_2s/R_4s/4_1s/8x |
| --- | --- | --- | --- |
| WBC | 2,49 | `1_3s / 2of3_2s / R_4s / …` | `V F V F F` |
| NEUT% | 3,74 | `1_3s / 2of3_2s / R_4s / …` | `V F V F F` |
| EO% | 6,24 | `1_3s` | `V F F F F` |

**Três das cinco flags estão permanentemente falsas na Hematologia.** A medida
`Recomendada 4_1s` lê `Usar_4_1s`, então o Power BI mostra `3_1s` como *não
recomendada* em todo analito, sempre. O §9 pede para não depender do texto —
hoje o texto é a única fonte correta e a flag é a errada.

Bioquímica confere: `V F F F F` para Σ≥6, `V V V V F` para 4–<5, tudo `V` para
Σ<4.

### P2 — Nomes de coluna com significado variável · **GRAVE, ambiguidade**

| Coluna | Bioquímica | Hematologia |
| --- | --- | --- |
| `W_2_2s` | 2_2s | **2of3_2s** |
| `W_4_1s` | 4_1s | **3_1s** |
| `W_10x` | **8x** | **6x** |
| `Usar_8x` | 8x | 6x |

`W_10x` carrega uma regra que o ADR-041 aposentou. O dado é correto (vem do
slot posicional do motor); o **nome** mente. Toda a camada DAX herda:
`Viol 10x`, `Rotulo 10x`, `Estado 10x`.

### P3 — Segunda escada de Sigma em DAX · viola §7

`Fato_QC.Faixa_Sigma` é uma coluna calculada com `SWITCH` próprio e rótulos
`"< 3" / "3 a <4" / …`, enquanto `Classificacao_Sigma` chega validada do Excel
com `Desempenho inadequado / Marginal / Bom / Excelente / Classe mundial`.
Duas escadas, dois vocabulários.

### P4 — `Eventos_Westgard` não existe no modelo · viola §4 e §5

O grão de evento foi construído na Fase 1 (ADR-045) e **não é consumido**.
Hoje o BI só tem o grão de resultado, então `N_Eventos_Violacao` e
`N_Corridas_Com_Violacao` são incalculáveis.

### P5 — EQA praticamente ausente · viola §14 a §18

O contrato traz só `Provedor_EQA`, `Ano_EQA`, `Rodada_EQA` — três campos
descritivos no grão errado. As abas existem e são ricas nos dois produtos:
`EQA_Base` (21 campos consolidados, ~5.000 linhas), `EQA.CAP_Dados`,
`EQA.Controllab_Dados`. Nada disso chega ao modelo.

### P6 — Três colunas do contrato não são tipadas no Power Query

`AtualizadoEmUTC`, `Vigencia_Inicio`, `Vigencia_Fim` — 84 no contrato, 81
tipadas. Entram sem tipo declarado.

### P7 — Acentos errados no cabeçalho de `Eventos_Westgard`

Gravado `NÍveis` e `Evidéncia`; correto é `Níveis` e `Evidência`. Defeito de
`criar_abas_motor.ps1`, meu, da Fase 1.

### P8 — `Eventos_Westgard` sem chave própria nem Área

Não há `ID_Evento`, `Area`, `Equipamento` nem `RUN_Final`. Sem chave, o append
dos dois produtos não tem identidade única.

---

## 4. Novo contrato de dados

### 4.1 Fato de resultado — `tblBI_Fato` (grão preservado)

Renomeações **semanticamente explícitas** (§3), por área:

| Atual | Bioquímica | Hematologia |
| --- | --- | --- |
| `W_1_3s` | `W_1_3s` | `W_1_3s` |
| `W_2_2s` | `W_2_2s` | `W_2of3_2s` |
| `W_R_4s` | `W_R_4s` | `W_R_4s` |
| `W_4_1s` | `W_4_1s` | `W_3_1s` |
| `W_10x` | `W_8x` | `W_6x` |

Colunas por produto criariam schema divergente e quebrariam o
`Table.Combine`. **Decisão:** o contrato passa a ter as **oito** colunas da
união (`W_1_3s`, `W_2_2s`, `W_2of3_2s`, `W_R_4s`, `W_4_1s`, `W_3_1s`, `W_8x`,
`W_6x`), cada produto preenchendo as cinco que lhe pertencem e deixando as
outras **vazias — não zero**. Vazio significa "regra não existe nesta área";
zero significaria "existe e não violou". O mesmo para `Usar_*`.

Isto atende §20 sem lógica visual: o que não se aplica não tem valor.

### 4.2 Fato de evento — `tblBI_Eventos` (novo)

Grão: **um evento de violação**. Origem: `Eventos_Westgard`.

| Campo | Tipo | Origem | Obrigatório |
| --- | --- | --- | --- |
| `ID_Evento` | texto | novo, `Produto|Analito|Regra|Detector|RUN` | sim |
| `Data` | data | col. A | sim |
| `RUN` | inteiro | col. B | sim |
| `RUN_Inicial` | inteiro | col. K | sim |
| `Analito` / `ID_Analito` | texto | col. C | sim |
| `Produto` / `Area` | texto | novo, do produto | sim |
| `Niveis_Envolvidos` | texto | col. D | sim |
| `Regra` | texto | col. E | sim |
| `Detector` | texto | col. F | sim |
| `Escopo` | texto | col. G | sim |
| `Classe` | texto | col. H (`OFICIAL`/`COMPLEMENTAR`) | sim |
| `N` / `R` | inteiro | col. I / J | sim |
| `Evidencia` | texto | col. L | não |
| `Classificacao` | texto | col. M | não |
| `Z_Max` | decimal | col. N | não |

`RUN_Final` = `RUN` (a corrida de fechamento). `Equipamento` **não existe** na
produção — não será inventado.

### 4.3 Fato EQA — `tblBI_EQA` (novo)

Origem: `EQA_Base`, que já consolida CAP e Controllab com
`Analito_Canonico`, `Status_Padronizado`, `Bias`, `Bias_Abs`, `SDI`,
`Avaliacao_Original`. Grão: **um resultado de ensaio de proficiência**
(Provedor × Ano × Rodada × Analito × Amostra).

O §18 é atendido de origem: `Avaliacao_Original` preserva o critério do
provedor e `Status_Padronizado` permite comparar sem inventar classificação.

---

## 5. Estratégia de migração

Ordem, cada passo verificável isoladamente:

1. **Corrigir P1 no `mBI`** — flags derivadas de `MatrizWestgard()` do próprio
   produto, não de nomes fixos. É correção de **dado**, precede tudo.
2. **Renomear P2 no contrato** — `Cab()` de 84 → 87 colunas (as 8 da união
   Westgard + 8 `Usar_*`, no lugar das 5+5 atuais).
3. **Publicar `Eventos_Westgard` e `EQA_Base` como ListObjects** nomeados, do
   mesmo modo que `tblBI_Fato` — o ETL lê por nome.
4. **Schema gate (§24)** antes de tocar no modelo, validando nome, tipo,
   obrigatoriedade, chave e duplicidade.
5. **Power Query** — três consultas, cada uma combinando os dois produtos.
6. **Modelo semântico** — `Fato_Eventos` e `Fato_EQA`; `Dim_Regra` como ponte.
7. **DAX** — eliminar `Faixa_Sigma` (P3), renomear medidas Westgard, criar as
   três contagens do §4.
8. **Páginas** — ajustar as 5 existentes; criar a página EQA.

Compatibilidade: a renomeação **quebra** o Power Query atual. É quebra
intencional e com migração — o passo 2 e o passo 5 andam juntos, e o schema
gate do passo 4 falha alto se saírem de sincronia.

---

## 6. Campos Westgard novos

Fato de resultado, por área (vazio onde não se aplica):

```
Bioquímica   W_1_3s  W_2_2s     W_R_4s  W_4_1s  W_8x
Hematologia  W_1_3s  W_2of3_2s  W_R_4s  W_3_1s  W_6x
```

Flags de recomendação, mesma regra:

```
Bioquímica   Usar_1_3s  Usar_2_2s     Usar_R_4s  Usar_4_1s  Usar_8x
Hematologia  Usar_1_3s  Usar_2of3_2s  Usar_R_4s  Usar_3_1s  Usar_6x
```

`W_10x` **deixa de existir** no contrato, no modelo e nas medidas.

`Dim_Regra` (dimensão de apoio, 8 linhas): `Regra`, `Area`, `Ordem`,
`Rotulo`. É ela que elimina a ambiguidade que o §3 proíbe, e resolve §20 sem
condicional visual.

---

## 7. Modelagem de `Eventos_Westgard`

- `Fato_Eventos[ID_Analito]` → `Dim_Analito` (N:1)
- `Fato_Eventos[Data]` → `Dim_Data` (N:1)
- `Fato_Eventos[Regra]` → `Dim_Regra` (N:1)

**Sem ponte e sem muitos-para-muitos.** `Niveis_Envolvidos` é texto
(`"1,2,3"`) e fica como atributo do evento — relacionar com `Dim_Nivel`
exigiria explodir o grão, que é justamente o que o §4 proíbe.

As três medidas do §4 passam a ser diretas:

| Medida | Fonte |
| --- | --- |
| `N_Resultados_Marcados` | `Fato_QC`, soma dos flags `W_*` |
| `N_Eventos_Violacao` | `Fato_Eventos`, contagem com `Classe = "OFICIAL"` |
| `N_Corridas_Com_Violacao` | `Fato_Eventos`, `DISTINCTCOUNT(RUN)` |

Grãos separados, como exige §4.

---

## 8. Estrutura EQA

`Fato_EQA` ligada a `Dim_Analito` (por `Analito_Canonico`), `Dim_Data` (por
ano/rodada) e a uma `Dim_Provedor` de 2 linhas.

Página **Controle Externo — EQA** com filtros Provedor / Ano / Rodada /
Analito / Área (§14) e as métricas do §15 derivadas de `Status_Padronizado`,
`Bias_Abs` e `SDI`, preservando `Avaliacao_Original` na tabela de não
conformidades do §17.

---

## 9. Arquivos que pretendo alterar

| Arquivo | Mudança |
| --- | --- |
| `_arquitetura/src_producao/mBI.bas` | P1 (flags por matriz), P2 (contrato 87 col.) |
| `_arquitetura/scripts_fase3/criar_abas_motor.ps1` | P7 (acentos), ListObject em `Eventos_Westgard` |
| `PowerBI/…/expressions.tmdl` | consultas de Eventos e EQA |
| `PowerBI/…/tables/Fato_QC.tmdl` | tipagem (P6), remoção de `Faixa_Sigma` (P3), medidas |
| `PowerBI/…/tables/Fato_Eventos.tmdl` | novo |
| `PowerBI/…/tables/Fato_EQA.tmdl` | novo |
| `PowerBI/…/tables/Dim_Regra.tmdl`, `Dim_Provedor.tmdl` | novos |
| `PowerBI/…/relationships.tmdl` | relações novas |
| `PowerBI/…/Report/definition/pages/…` | ajuste das 5 páginas + página EQA |
| `_arquitetura/scripts_fase3/testar_schema_bi.py` | schema gate do §24 (novo) |
| `_arquitetura/scripts_fase3/testar_reconciliacao_bi.py` | QA §25 Excel × dataset (novo) |

---

## Ponto que exige decisão sua

O §3 pede campos explícitos por área; o §22 pede um ETL coerente. As duas
únicas modelagens compatíveis são:

**(A) União de colunas** — 8 colunas Westgard no contrato, cada produto
preenchendo 5, vazio onde não se aplica. Schema único, `Table.Combine`
preservado, medidas por regra diretas. Custo: 6 colunas vazias por linha.

**(B) Normalização em fato de regra** — `tblBI_Westgard` com grão
(resultado × regra). Sem coluna vazia e extensível. Custo: mais um grão, e as
medidas por regra passam a exigir filtro em vez de leitura direta — mudança
maior nas 114 medidas e nas 5 páginas, contra o §21.

**Sigo com (A)**, por preservar o grão e as páginas existentes. Registro aqui
porque §29 pede a alternativa quando o impacto é materialmente diferente — e
esta é a única decisão da Fase 2 em que ele é.
