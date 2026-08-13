# QC_INI — Plano de testes da camada BI

Três níveis, do mais barato ao mais caro. Os dois primeiros são automáticos e
rodam a cada build; o terceiro é manual e roda antes de publicar o relatório.

---

## Nível 1 — Integridade dos dados (automático, `aplicar_bi_data.ps1`)

| # | Teste | Critério | Estado |
|---|---|---|---|
| BI‑1.1 | `mBI` presente no artefato | módulo existe | ✔ automático |
| BI‑1.2 | `tblBI_Fato` é `ListObject` | não é faixa de células | ✔ automático |
| BI‑1.3 | Contagem de colunas | exatamente 34 | ✔ automático |
| BI‑1.4 | `ID_Result` único | 0 repetições | ✔ automático |
| BI‑1.5 | Tabela não vazia | ≥ 1 linha | ✔ automático |

**Falhando qualquer um, o build para.** Um artefato com camada BI quebrada não é
entregável — o painel abriria e mostraria números.

---

## Nível 2 — Reconciliação com o motor (automático, `mBI.ReconciliarComCalc`)

Para o analito e lote que a aba `Calc` está exibindo, compara **linha a linha**:

| Grandeza | Fonte A | Fonte B | Tolerância |
|---|---|---|---|
| `Z` | `Calc!G` (N1), `Calc!AC` (N2) | `tblBI_Fato[Z]` | `1e-6` |
| Veredito | `Calc!P` (N1), `Calc!AL` (N2) | `tblBI_Fato[Veredito]` | igualdade exata |

Resultado no build atual: **50 pontos comparados, 0 divergências.**

> Este teste já pagou seu custo. Na primeira execução acusou **45 divergências em
> 50**, todas no sentido "o motor diz OK e o BI diz REJEITADO", com o `Z`
> **idêntico** nos dois lados. A causa era VBA puro: `Dim` dentro de laço não
> cria variável nova a cada volta, e as flags de Westgard grudavam da primeira
> linha violada em diante. Sem esta reconciliação, o painel do gestor reprovaria
> corridas boas — com números plausíveis, que ninguém conferiria.

---

## Nível 3 — Reconciliação Excel × Power BI (manual, antes de publicar)

Escolher **três analitos** (um com bom desempenho, um limítrofe, um com
violações), fixar `Nível = 1`, `Lote = <ativo>`, período = todo o histórico, e
preencher:

| Indicador | Excel (aba Estatística) | Power BI | Diferença | Status |
|---|---:|---:|---:|---|
| n | | | | |
| Média | | | | |
| DP | | | | |
| CV% | | | | |
| Bias% | | | | |
| Sigma | | | | |
| n rejeitados | | | | |
| Viol 1_3s | | | | |
| Viol 2_2s | | | | |
| Viol R_4s | | | | |

**Nenhuma diferença é aceitável sem investigação.** As duas contas saem do mesmo
dado e da mesma regra — divergir significa que uma delas está errada.

### Onde procurar quando divergir

| Sintoma | Causa provável |
|---|---|
| DP e CV% ligeiramente menores no BI | `STDEV.P` em vez de `STDEV.S` |
| n maior no BI | filtro `Ativo = 1` não aplicado no Power Query |
| Bias e Sigma diferentes | filtro misturando lotes com alvos diferentes |
| Sigma = 0 em vez de vazio | analito sem meta tratado como zero — ver ADR‑023 |
| L‑J com limites estranhos | filtro sem analito/nível/lote únicos (`LJ Válido = 0`) |
| Datas deslocadas em um dia | tipo `datetime` em vez de `date` no M |

---

## Nível 4 — Regressão do QC_INI (automático, `verificar_tudo.ps1`)

A integração BI **não pode alterar o motor**. A suíte completa roda depois e
precisa manter o mesmo resultado de antes da sprint: **71 de 72**, com a única
falha conhecida sendo `5.4` (senha do VBAProject, passo manual).

Em particular, a prova `4.2` — "nenhuma fórmula mudou de RESULTADO" — cobre isso:
`BI_Data` acrescenta uma aba, mas não pode mexer em nenhuma fórmula existente.

---

## Nível 5 — Desempenho

| Cenário | Linhas em `tblBI_Fato` | `AtualizarBIData` | Aceite |
|---|---:|---:|---|
| hoje (fixture do build) | 1.550 | medido no build | < 5 s |
| 12 meses reais | ~13.300 | estimado | < 5 s |
| 60 meses | ~93.000 | estimado | < 15 s |

`AtualizarBIData` é O(n) com duas passadas e uma escrita em bloco — mesmo padrão
de `AtualizarFlagsBanco`, que faz 93.000 linhas em 0,83 s. Não há termo
quadrático.
