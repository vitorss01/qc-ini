# QC_INI — Arquitetura de integração com Power BI (ADR-026)

Data: 12/08/2026 · Produto de referência: `QC_Bioquimica.xlsm` (31 analitos, 2 níveis)

---

## 1. Princípio

O Excel continua sendo o **núcleo operacional**: é onde o resultado entra, onde a
regra de Westgard é aplicada e onde a liberação é assinada. O Power BI é camada
de **visualização, análise, monitoramento e alerta** — nunca uma segunda fonte
de verdade.

```
QC_INI (.xlsm)
   │  motor: Calc, mEstatistica, mEspecificacoes, mBanco
   ▼
BI_Data  ──►  tabela estruturada tblBI_Fato          ← ESTA é a interface
   ▼
Power Query (M)
   ▼
Modelo Power BI (estrela)
   ▼
Desktop · Web · Tablet · Smartphone
```

**A interface é o nome da tabela, não a posição das células.** `tblBI_Fato` é um
`ListObject`; cresce e encolhe sozinha e o M não muda. Uma faixa `A1:AH9999`
quebraria a cada linha a mais.

---

## 2. Modelo de dados

Estrela, com uma fato e quatro dimensões. As dimensões são **derivadas da fato**
no Power Query — não há tabelas de dimensão no Excel, porque criá-las ali seria
manutenção duplicada sem ganho.

### Fato — `tblBI_Fato`

**Granularidade: um resultado de controle = (Lote, RUN, Nível, Analito).**

É a mesma chave natural que o `DB_Resultados` usa no `UpsertResultados`. Escolher
outra granularidade obrigaria a agregar antes de exportar, e agregação na
exportação é onde os números começam a divergir da tela.

### Dimensões

| Dimensão | Chave | Origem |
|---|---|---|
| `Dim_Data` | `Data` | calendário gerado em M, cobrindo min..max da fato |
| `Dim_Analito` | `ID_Analito` | distinct da fato (Analito, Área, Unidade) |
| `Dim_Lote` | `ID_Lote` | distinct da fato (Lote, ID_Lote) |
| `Dim_Nivel` | `Nivel` | distinct da fato |

### O que **não** existe, e por quê

`Equipamento` e `Setor` foram pedidos, e **não estão no modelo**. A produção não
possui as abas `Corridas` nem `Cfg_Status`, e nenhuma outra guarda esses campos.
Uma coluna sem origem é uma coluna que alguém vai filtrar e obter uma conclusão
errada. Quando o campo passar a existir no QC_INI, entra aqui — o gancho está
documentado em [`DATA_DICTIONARY.md`](DATA_DICTIONARY.md).

---

## 3. Chaves

Estáveis, textuais, independentes de posição:

| Chave | Composição | Onde é usada |
|---|---|---|
| `ID_Result` | `LOTE\|RUN\|NIVEL\|ANALITO` | chave primária da fato — unicidade validada no build |
| `ID_Corrida` | `LOTE\|RUN` | agrupa os resultados de uma mesma corrida |
| `ID_Analito` | nome em maiúsculas | junta com `Dim_Analito` |
| `ID_Lote` | núcleo do lote | junta com `Dim_Lote` |

Nenhuma depende de número de linha. Reordenar o banco não quebra o modelo.

---

## 4. Onde cada cálculo mora

A regra: **o que o motor já calcula não é recalculado no Power BI**; o que é
agregação de período é DAX, porque depende do filtro do usuário e não pode ser
pré-calculado.

| Indicador | Fonte hoje | Excel | BI_Data | DAX | Observação |
|---|---|:---:|:---:|:---:|---|
| Resultado | DB_Resultados | ✔ | ✔ | | transportado |
| Média/DP alvo do lote | LotesStore | ✔ | ✔ | | **por lote**, não da tela |
| Z (desvio padronizado) | Calc!G, Calc!AC | ✔ | ✔ | | reconciliado a cada build |
| Limites ±1/2/3 SD | Calc!Q..W | ✔ | ✔ | | materializados p/ Levey-Jennings |
| CVtp / BIAStp / ETp | Eng_Especificacoes | ✔ | ✔ | | meta vigente do ano |
| Westgard 1_3s, 2_2s, R_4s | Calc!K, L, M | ✔ | ✔ | | réplica literal + reconciliação |
| Veredito OK/REJEITADO | Calc!P, AL | ✔ | ✔ | | |
| n do período | Estatística | ✔ | | ✔ | depende do filtro |
| Média/DP observados | Estatística | ✔ | | ✔ | depende do filtro |
| CV% observado | Estatística | ✔ | | ✔ | `DP/Média*100` |
| Bias% | Estatística | ✔ | | ✔ | vs. média-alvo do lote |
| Sigma | Painel | ✔ | | ✔ | `(ETp − |Bias|)/CV` |
| % fora de controle | — | | | ✔ | novo, definido em DAX |
| Violações por regra | Calc | ✔ | ✔ | ✔ | flag na fato, contagem em DAX |

**Por que média/DP observados ficam em DAX e não na fato:** eles dependem da
janela que o usuário escolher. Materializá-los na fato congelaria um período e o
usuário veria um CV% que não corresponde ao filtro aplicado — exatamente o tipo
de número plausível e errado que este projeto persegue.

---

## 5. Westgard — o que existe de verdade

O QC_INI implementa **três** regras. Verificado lendo as fórmulas do `Calc`:

| Regra | Fórmula no Calc | Escopo | No BI |
|---|---|---|---|
| `1_3s` | `ABS(Z) > 3` | por resultado | ✔ idêntica |
| `2_2s` | `Z(N1)>2 E Z(N2)>2`, ou ambos `< −2` | **entre níveis da mesma corrida** | ✔ idêntica |
| `R_4s` | `Z(N1)>2 E Z(N2)<−2`, ou o inverso | **entre níveis da mesma corrida** | ✔ idêntica |
| `4_1s` | `IF(OR(FALSE,FALSE),1,0)` | — | ✘ **não implementada** |
| `10x` | `IF(OR(FALSE,FALSE),1,0)` | — | ✘ **não implementada** |

`Calc!N3` e `Calc!O3` são placeholders que sempre devolvem 0. As colunas existem
na estrutura, mas a regra não. **O BI não inventa o que o motor não calcula** —
publicar um `4_1s` fabricado no painel do gestor seria pior do que não ter.

Nota de CQ: `2_2s` aqui é *within-run across-levels*. A variante *across-runs*
(mesmo nível, duas corridas consecutivas) **não** está implementada. Isso é uma
lacuna real do motor, não da camada BI, e está registrada como pendência.

---

## 6. Reconciliação

O build **falha** se a camada BI divergir do motor. `mBI.ReconciliarComCalc`
compara, para o analito e lote em tela, o `Z` e o veredito linha a linha contra
o `Calc`. Tolerância: `1e-6` no Z, igualdade exata no veredito.

Não é conferência cosmética: as duas contas saem do mesmo dado e da mesma regra.
Se divergirem, uma das duas está errada — e num sistema de CQ não dá para saber
qual sem investigar.

---

## 7. Atualização dos dados

**Escolhido: Excel → Power Query → Power BI, lendo o `.xlsm` direto.**

| Opção | Avaliação |
|---|---|
| `.xlsm` → Power Query → PBI | **escolhida.** Sem cópia intermediária, sem segunda fonte de verdade. O Power Query lê `.xlsm` nativamente e enxerga `ListObject` pelo nome. |
| `.xlsm` → CSV → PBI | rejeitada. Cria um artefato intermediário que pode ficar velho sem ninguém notar, e perde o tipo das colunas. |
| Banco de dados | rejeitada por ora. Correta em escala, mas destrói a proposta do produto: um arquivo que funciona sem TI. |

O arquivo fica em OneDrive, então o Power BI Service pode atualizá-lo por
gateway ou pelo conector do OneDrive/SharePoint sem gateway.

---

## 8. Performance

Com 60 meses de histórico (medido: 93.000 registros), `tblBI_Fato` teria ~93.000
linhas × 34 colunas. Para o VertiPaq isso é pequeno — cardinalidade alta só em
`ID_Result` (única por definição) e `Resultado`.

Recomendação de modelagem: **não** marcar `ID_Result` como coluna visível no
relatório; ela existe para garantir grão e depuração. Ocultá-la evita que alguém
a arraste para um visual e crie uma tabela de 93.000 linhas.

---

## 9. LGPD e dados sensíveis

`tblBI_Fato` **não contém dado de paciente**. Controle de qualidade opera sobre
material controle, não sobre amostra de paciente. Os únicos campos que remetem a
pessoas são `Responsável`/`Assinatura` na aba `Liberação`, e eles **não** são
exportados para o BI.

---

## 10. Limite conhecido: geração do `.pbix`

Power BI Desktop está instalado (`2.156.951.0`), mas **não há caminho
automatizável para gerar um `.pbix` por script** nesta máquina:

| Ferramenta | Estado |
|---|---|
| `pbi-tools` | não instalado |
| Tabular Editor | não instalado |
| módulo `MicrosoftPowerBIMgmt` | não instalado |
| Power BI REST API | exigiria tenant, app registration e token |

O `.pbix` é um contêiner binário proprietário; escrevê-lo à mão não é
reprodutível nem suportado. **Nenhum `.pbix` foi fabricado.** No lugar, entregam-se
todos os artefatos que o Power BI consome — as queries M, as medidas DAX, o
desenho do modelo e o roteiro de montagem — em
[`POWER_BI_SETUP.md`](POWER_BI_SETUP.md).

Se o usuário instalar `pbi-tools` ou Tabular Editor, a geração passa a ser
automatizável e a etapa entra no build.
