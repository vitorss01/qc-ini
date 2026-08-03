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
8. **A chave lógica da saída do motor é `ANALITO|RUN`, nunca o RUN sozinho.**
   O RUN identifica a corrida, e a corrida cobre vários analitos. Qualquer consulta
   futura a `Eng_Saida` casa por `engChave`. Quando a chave não bate, o resultado é
   vazio — falha visível, jamais dado de outro analito.
9. **Módulo de documento (`Planilha*.cls`) se aplica sem o cabeçalho.**
   `AddFromFile` insere o arquivo inteiro como código, e `VERSION`/`BEGIN`/`Attribute`
   viram erro de sintaxe que derruba a compilação do **projeto todo** — inclusive
   procedimentos de outros módulos, que passam a falhar com `0x800A9C68`.
   `aplicar_vba.ps1` remove o cabeçalho antes de aplicar.

---

## 5. Parecer arquitetural — achados

Revisão feita sobre o candidato a RC1, sem implementar nada.

### Ciclos

**Nenhum.** O motor deixou de ler as abas de interface no Marco 3. A ordem
`AtualizarCalc → AtualizarPainelEng → AtualizarEstatisticaAba` é respeitada por
`AtualizarEstatistica` e por `RecalcularAnalitoAtual`.

### RESOLVIDO — 1. `Eng_Saida` ficava obsoleto ao trocar de analito

Regressão introduzida pelo Marco 2, encontrada por esta revisão e **corrigida antes do
congelamento**.

*O defeito:* as colunas de regra do `Calc` casavam por `MATCH` sobre o RUN — mas o RUN
identifica a **corrida**, e a mesma corrida cobre vários analitos. Trocar o analito no
`Painel` disparava apenas `AtualizarEixos`; o motor não rodava, o `MATCH` encontrava o
mesmo RUN e devolvia o veredicto **do analito anterior**. Reproduzido: motor rodado para
`RDW-CV`, trocando para `WBC` sem reexecutar, a linha do RUN 6 exibia `REJEITADO` quando
o WBC real é `OK`. Na produção não ocorria, porque o `Calc` calculava as regras sozinho.

*A correção, estrutural, em duas camadas:*

1. **Chave lógica `ANALITO|RUN`**, publicada pelo motor na coluna AB de `Eng_Saida`
   (nome `engChave`). As fórmulas casam por `MATCH(selAnalito&"|"&RC2, engChave, 0)`.
   Se `Eng_Saida` for de outro analito, o `MATCH` **falha** e o `IFERROR` devolve vazio.
   O erro passa a ser visível, e nunca um veredicto falso.
2. **`Planilha7.Worksheet_Change` chama `RecalcularAnalitoAtual`** quando muda o analito
   (B3) ou o filtro (G3, G4, M3, M4, N3, N4) — todos alteram o que o motor publica.

*Verificado:* com `Eng_Saida` de RDW-CV e a planilha pedindo WBC, as 10 primeiras linhas
devolvem vazio — **0 veredictos falsos**, contra 1 em 10 antes. Com eventos ligados, o
seletor repopula sozinho e o veredicto é o do analito certo. A regressão consolidada
permanece em 233 divergências com a mesma decomposição: nenhuma nova foi introduzida.

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
