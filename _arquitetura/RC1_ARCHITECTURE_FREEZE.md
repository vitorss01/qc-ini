# RC1 — Architecture Freeze

**Sistema QC_INI · Hematologia** · congelado em 03/08/2026, após o Sprint HARDENING 1
(Marcos 1 a 5).

Este documento define a arquitetura da RC1 e as regras que **não podem ser quebradas**
sem decisão explícita do gestor. Serve para que, daqui a dois meses, ninguém volte a
escrever no `Calc` por conveniência.

Evidência de tudo o que está afirmado aqui:
[`src_hardening1/rc1_mapa_escritas.csv`](src_hardening1/rc1_mapa_escritas.csv) (varredura
do VBA) e [`src_hardening1/marco5_adr019.csv`](src_hardening1/marco5_adr019.csv)
(varredura das fórmulas).

---

## 1. Camadas

```
         ENTRADA                      Resultados  (view de digitação)
            │                              │
            ▼                              ▼
  ┌──────────────────┐            ┌──────────────────┐
  │     mOperacao    │───────────▶│      mDados      │
  │  (orquestração)  │            │ (camada de dados)│
  └──────────────────┘            └────────┬─────────┘
                                           ▼
                                   ╔═══════════════════╗
                                   ║   DB_Resultados   ║  única fonte operacional
                                   ╚═════════╤═════════╝  schema A:G
                                             │ leitura
                                             ▼
                                   ┌───────────────────┐
                                   │   mEstatistica    │  MOTOR
                                   │  única camada de  │  média · DP · CV · bias CQI
                                   │      cálculo      │  ET · Sigma · Z · Westgard
                                   └─────────┬─────────┘
                                             │ escrita
                                             ▼
                                   ╔═══════════════════╗
                                   ║     Eng_Saida     ║  camada de saída
                                   ╚═════════╤═════════╝  (very hidden)
                     ┌───────────────────────┼───────────────────────┐
                     │ fórmula               │ fórmula               │ fórmula
                     ▼                       ▼                       ▼
              ┌────────────┐          ┌────────────┐          ┌──────────────┐
              │    Calc    │          │   Painel   │          │  Estatística │
              │  seleção,  │─────────▶│apresentação│          │ apresentação │
              │filtro,plot │  gráfico │            │          │  + bias EQC  │
              └────────────┘          └────────────┘          └──────┬───────┘
                                                                     │
                                                              ┌──────▼───────┐
                                                              │  EQC_Dados   │
                                                              │ (trueness)   │
                                                              └──────────────┘
```

`Registros`, `Eventos_Westgard`, `Usuarios`, `LotesStore`, `LiberStore` e
`RegistrosStore` são abas de dado/histórico, fora da cadeia de cálculo.

---

## 2. Quem escreve em quem

Levantado por varredura do VBA, não por memória. 36 operações de escrita no total.

| Módulo | Escreve em | Ops | Papel |
|---|---|---|---|
| `mDados` | `DB_Resultados` | 6 | **Único** que grava no banco |
| `mEstatistica` | `Eng_Saida` | 12 | Motor: publica todo cálculo |
| `mEstatistica` | `Eventos_Westgard` | 3 | Histórico auditável de violações |
| `mOperacao` | `Resultados` | 2 | View de digitação |
| `mSeguranca` | `Usuarios` | 5 | Login e papéis |
| `mLotes` | `LotesStore`, `LiberStore`, `RegistrosStore` | 8 | Persistência de lotes |

| Aba | Recebe escrita de VBA? |
|---|---|
| **`Calc`** | **Ninguém** |
| **`Painel`** | **Ninguém** |
| **`Estatística`** | **Ninguém** |
| `Eng_Saida` | Só `mEstatistica` |
| `DB_Resultados` | Só `mDados` |

---

## 3. Quem lê o quê

| Componente | Lê |
|---|---|
| `mEstatistica` (motor) | `DB_Resultados`, `Analitos`, `Cfg_Status`, `Registros`, `Eng_Saida` |
| `Calc` (fórmulas) | `DB_Resultados` via nomes `r*` — **só seleção e filtro**; `Eng_Saida` para as regras |
| `Painel` (fórmulas) | `Eng_Saida`, `Estatística` (bias EQC), `Analitos` |
| `Estatística` (fórmulas) | `Eng_Saida`, `EQC_Dados` (bias externo) |
| `clsCht` | `Calc` (séries do gráfico) |
| `mUI` | `Painel` (leitura e orquestração) |

**O motor não lê `Calc`, `Painel` nem `Estatística`.** Foi assim que o ciclo
motor → planilha → motor deixou de existir.

---

## 4. Invariantes congeladas

1. **`DB_Resultados` só é escrito por `mDados`.** Formulário nenhum grava direto.
2. **Todo cálculo estatístico vive em `mEstatistica` e é publicado em `Eng_Saida`.**
   Nenhuma fórmula de interface calcula média, DP, CV, bias do CQI, Erro Total, Sigma
   ou regra de Westgard.
3. **A planilha pode selecionar, filtrar, ordenar e localizar** — `COUNTIFS`, `SUMIFS`,
   `MAXIFS`, `INDEX/MATCH`, `AGGREGATE` de posição. É o que o Excel faz bem e é o que
   mantém o seletor de analito vivo.
4. **O bias que alimenta Erro Total, Sigma e Classificação vem do EQC**
   (`EQC_Dados`), não do alvo do lote. Trueness se estima por comparação
   interlaboratorial. Decisão do gestor, Marco 4.
5. **`Cfg_Status` é a única autoridade sobre elegibilidade.** Acrescentar um estado é
   digitar uma linha, sem tocar em código.
6. **O `.xlsm` é artefato; a fonte do código é `snapshot_producao/<produto>/vba/`**
   mais os patches em `src_hardening1/`. Toda alteração nasce em texto versionado.
7. **`aS` e `aM` são proibidos como identificadores.** VBA é insensível a maiúsculas:
   `aS` é a palavra reservada `As` e o módulo não compila. O lint de
   `gerar_mEstatistica.ps1` barra a classe inteira.

---

## 5. Parecer arquitetural — achados

Revisão feita sobre o candidato a RC1, sem implementar nada.

### Ciclos

**Nenhum.** O motor deixou de ler as abas de interface no Marco 3. A ordem
`AtualizarCalc → AtualizarPainelEng → AtualizarEstatisticaAba` é respeitada por
`AtualizarEstatistica` e por `RecalcularAnalitoAtual`.

### CRÍTICO — 1. `Eng_Saida` fica obsoleto ao trocar de analito

Regressão introduzida pelo Marco 2. As colunas de regra do `Calc` leem `Eng_Saida`
casando por `MATCH` sobre o RUN — mas o RUN identifica a **corrida**, e é compartilhado
entre analitos. Trocar o analito no `Painel` dispara apenas `AtualizarEixos`
([`Planilha7.cls:10`](snapshot_producao/Hematologia/vba/Planilha7.cls:10)); o motor não
roda, o `MATCH` encontra o mesmo RUN e devolve o veredicto **do analito anterior**.

Reproduzido: com o motor rodado para `RDW-CV` e trocando para `WBC` sem reexecutar,
a linha do RUN 6 exibe `REJEITADO` quando o valor real do WBC é `OK`.

Na produção isso não acontecia, porque o `Calc` calculava as regras por conta própria.

**Correção recomendada, em duas camadas:**
- *À prova de falha:* prefixar as fórmulas de regra com `IF(engAnalito<>selAnalito,"",…)`.
  O painel passa a exibir vazio em vez de valor errado. Não exige VBA.
- *Correta:* `Planilha7.Worksheet_Change` chamar `RecalcularAnalitoAtual` quando
  `selAnalito` mudar. Custo medido: ~0,27 s.

**Bloqueia o congelamento até ser corrigido.**

### ALTO — 2. `UpsertResultados` ressuscita registro excluído

[`mDados.bas:97`](snapshot_producao/Hematologia/vba/mDados.bas:97) força
`Status = Ativo` em toda atualização. Reenviar a mesma chave traz de volta uma linha
excluída, sem registro. Item 2.3 do gate, confirmado em código.

### ALTO — 3. `RegistrarLog` é um stub vazio

[`mDados.bas:197`](snapshot_producao/Hematologia/vba/mDados.bas:197). A função existe
na camada certa — dentro de `mDados`, como o item 3.2 exige — e não faz nada. Itens
3.1 e 3.2.

### MÉDIO — 4. RUN colapsa duas corridas no mesmo dia

[`mDados.bas:50`](snapshot_producao/Hematologia/vba/mDados.bas:50): `NovoRUN` é função
de (Data, lote). Turno e pós-calibração viram a mesma corrida. Item 2.4.

### MÉDIO — 5. `CarregarDB` lê o bloco `A:G` do banco

[`mDados.bas:39`](snapshot_producao/Hematologia/vba/mDados.bas:39). As duas tabelas de
log da Sprint NC precisam ficar em blocos deslocados à direita, com intervalos nomeados
próprios. Restrição concreta de projeto, não opinião.

### MÉDIO — 6. 1.620 fórmulas do `Calc` dependem de `regRep1/2/3`

A Sprint NC elimina as colunas `Rep 1/2/3` da aba `Registros`. Essas 1.620 fórmulas
quebram junto. Precisa entrar no planejamento daquela sprint, não ser descoberto durante.

### BAIXO — 7. `AtualizarPainelEng` executa duas vezes

`AtualizarOperacao` e `AtualizarTudo` chamam `AtualizarEstatistica` (que já roda
`AtualizarPainelEng`) e em seguida `AtualizarPainel`, que roda de novo. Não gera erro —
custa ~0,16 s por execução redundante. `mUI.AtualizarPainel` também é o único caminho
que chamaria `AtualizarPainelEng` sem `AtualizarCalc` antes; hoje nenhum chamador faz
isso, mas a fragilidade está armada.

### BAIXO — 8. Z-score calculado em dois lugares

`Calc` G/AC/AY (540 células) e o motor. Concordam por construção — mesmas entradas,
valor e alvo lidos de `Analitos`. Não é motor de regras duplicado. Simplificação
possível numa v2: apontar também o Z para `Eng_Saida`.

### BAIXO — 9. `AtualizarBanco` não atualiza banco nenhum

[`mDados.bas:192`](snapshot_producao/Hematologia/vba/mDados.bas:192): o corpo é
`Application.Calculate`. Nome engana quem lê a cadeia de orquestração.

---

## 6. Simplificações possíveis — backlog v2, não RC1

- Unificar o Z-score em `Eng_Saida` (achado 8) e aposentar 540 fórmulas.
- Remover a chamada redundante de `AtualizarPainelEng` (achado 7).
- Renomear `AtualizarBanco` para `RecalcularPlanilha` (achado 9).
- Fazer as 1.160 fórmulas que fixam `rStatus="Ativo"` consultarem `Cfg_Status`.
  Hoje o efeito é nulo porque só `Ativo` é elegível; vira defeito no dia em que outro
  estado for marcado como elegível.

Nenhum desses é impedimento para a RC1.
