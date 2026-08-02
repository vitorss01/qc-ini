## BLOQUEADOR RESOLVIDO — causa raiz provada

### Ponto exato do congelamento

**Não é uma linha *dentro* de `RegistrarEventosWestgard` — é a *chamada* para ela.**

`mEstatistica.bas:885` — `Dim aM As Double, aS As Double, etp As Double`

O identificador **`aS` é a palavra reservada `As`** (identificadores VBA são *case-insensitive*). O parser lê `Dim aM As Double, ` e então encontra a keyword `As` onde exige um identificador → **erro de sintaxe**.

Mesma falha em mais dois pontos: `mEstatistica.bas:683` (`AtualizarPainelEng`) e `mEstatistica.bas:754` (`AtualizarEstatisticaAba`).

### Por que ficou latente e por que "congela sem erro"

1. O código foi instalado por script via `CodeModule.AddFromString`, que **não faz verificação de sintaxe** na inserção (diferente de digitar no VBE).
2. O VBA opera em **Compile On Demand** (padrão): compila um procedimento só quando ele é chamado pela primeira vez. Por isso `D_Alvo` e `D_Calc` rodaram — a conclusão "o projeto compila" era falsa, só os procedimentos exercitados compilavam.
3. No momento da chamada o VBE levanta um **diálogo MODAL**. Sob `/automation` com `Visible=False` esse diálogo é invisível → o Excel bloqueia indefinidamente com **0% de CPU**.
4. `On Error GoTo` **não intercepta erro de compilação** — daí "trava, não lança erro".

Mesma classe do Erro 9 já corrigido (`val()` sombreando `Val()`): identificador injetado por script colidindo com palavra reservada.

### Cadeia de evidências

| # | Evidência | Resultado |
|---|---|---|
| 1 | Trace em arquivo (sobrevive ao kill) | Só `=== D_REW INICIO ===`. O marcador `S0 entrada` — **primeira instrução do corpo** — nunca gravou. Execução jamais entrou no procedimento. |
| 2 | Amostragem de CPU/memória por 300 s | `cpu%=0,0` constante, `mem=341MB` estável, threads caindo 56→40. Excel **bloqueado**, não calculando. |
| 3 | Enumeração de janelas durante o travamento | `#32770` "Microsoft Visual Basic for Applications", `vis=1`, Static = **`'Erro de compilação:\n\nErro de sintaxe'`**, sobre `[mEstatistica (Código)]`, com `XLMAIN … en=0` (desabilitada = modal ativo). |
| 4 | Rotina **original** (sem meu clone) | Mesmo diálogo, mesma mensagem. |
| 5 | **Repro mínimo isolado** (pasta de trabalho NOVA, um módulo) | `Dim aM As Double, aS As Double, etp As Double` → congela com o diálogo idêntico. `aSd` → `RESULT 'OK 6' (0.01s)`. |

Corroboração independente: os **três** procedimentos que contêm `aS` são exatamente os três que nunca completaram (Painel `B7:J9` e Estatística `C7:M126` ainda com as fórmulas originais salvas, `Eventos_Westgard` vazia). `AtualizarCalc`, que **não** contém `aS`, sempre funcionou.

### Descartes (na ordem exigida)

- **(a) `Application.Calculation` automático** — DESCARTADO. CPU 0,0% por 300 s; tempestade de recálculo consumiria um núcleo a 100%. Após a correção a rotina completa em 0,066 s com `Calculation = -4105` (automático) **intocado**. A teoria das ~1,5×10⁸ leituras por recálculo não se sustenta.
- **(b) Cascata de eventos de planilha** — DESCARTADO por dupla evidência: congelou com `EnableEvents=False`; após a correção roda em 0,066 s com `EnableEvents=True`. Além disso a classe de `Eventos_Westgard` não tem handler algum (só Planilha2/7/12 têm, nenhuma toca essa aba).
- **(c) Recursão `AgregadoWestgard` ↔ `RegistrarEventosWestgard`** — DESCARTADO. A execução nunca chegou ao corpo da rotina (evidência 1).
- **(d) `Dictionary` de `Collection`** — DESCARTADO e **NÃO alterado**. A estrutura aninhada está intacta e processa 1.575 linhas / 63 grupos em 66 ms.

### Correção aplicada

Renomeação de `aS` → `aSd`. **7 linhas, 9 tokens, um único módulo.** Regex *case-sensitive* `(?<![A-Za-z0-9_])aS(?![A-Za-z0-9_])` — nunca casa com a keyword `As`.

Diff completo confirmado contra o original: as únicas diferenças de código são essas 7 linhas em `mEstatistica.bas`. (Os `Attribute VB_Base` dos 5 `.frm` foram re-carimbados pelo Excel ao salvar via COM — comportamento normal, nenhum código de formulário mudou.) 33 módulos antes e depois; nenhum resíduo de diagnóstico.

### Tempos medidos

| Rotina | Antes | Depois |
|---|---|---|
| `RegistrarEventosWestgard` | **congela** (morto a 240 s e a 300 s, CPU 0%) | **0,066 s** |
| `AtualizarPainelEng` | congela | 0,027 s |
| `AtualizarEstatisticaAba` | congela | 0,090 s |
| `AtualizarEstatistica` (orquestração) | congela | **0,172 s** (3 execuções seguidas: 0,172 / 0,160 / 0,176) |

`Eventos_Westgard` populada e coerente: **9 eventos**, `J2 = 9`, regras `{13s:2, 22s:2, 41s:3, 10x:2}`. Verificações de coerência: 0 violações `13s` com |z|≤3, 0 violações `22s` com |z|≤2; |z| entre 0,0387 e 3,5974. Resultado determinístico e idempotente entre execuções. Zero diálogos modais capturados. `Calculation`/`EnableEvents`/`ScreenUpdating` preservados.

Varredura de palavras reservadas nos 3 produtos: **QC_Bioquimica 0 colisões, QC_Imunologia 0 colisões**, QC_Hematologia 3 (as corrigidas). Os dois outros arquivos não foram abertos nem tocados.

### Entregáveis

- Build corrigido e validado: `C:\Users\vitor\OneDrive - MSFT\Desktop\QC_INI\_arquitetura\fase3a_corrigido\QC_Hematologia.xlsm`
- Backup do estado anterior: `C:\Users\vitor\OneDrive - MSFT\Desktop\QC_INI\_backup_pre_fase3b\QC_Hematologia.xlsm`
- **O arquivo vivo `QC_Hematologia.xlsm` NÃO foi modificado** (Regra 4). Promoção: `Copy-Item "_arquitetura\fase3a_corrigido\QC_Hematologia.xlsm" "QC_Hematologia.xlsm" -Force`

Scripts de diagnóstico e validação em `C:\Users\vitor\AppData\Local\Temp\claude\C--Users-vitor-OneDrive---MSFT-Desktop-QC-INI\919a6d1a-4dbb-471e-9759-5f896c9a1c9b\scratchpad\` — `wins.py` (enumera janelas modais durante o travamento) e `probe_run.py` + `probes/p_as.bas` (repro mínimo) são os dois que provaram a causa; `validate.py` é o harness de aceitação.

### Dois pontos que NÃO corrigi (Regras 1 e 3) — precisam de decisão do gestor

1. **`priR`/`priRun`/`ultR`/`ultRun` não são reinicializados por grupo** (`mEstatistica.bas:900`). Em VBA o escopo de `Dim` é o *procedimento*, não o laço; a linha 901 só zera `nv` e `maxZ`. A partir do primeiro grupo com violação, os 62 grupos seguintes gravam em `agg` a "primeira/última regra e RUN" do grupo anterior. Isso vai direto ao Painel em `AtualizarPainelEng:738`. Defeito provado por construção, hoje mascarado porque `AtualizarPainelEng` nunca havia executado. Correção: acrescentar `priR = "": priRun = 0: ultR = "": ultRun = 0` na linha 901.

2. **Primeira execução real destrói fórmulas permanentemente.** Agora que o caminho VBA executa, `AtualizarPainelEng` sobrescreve `Painel!B7:J9` e `AtualizarEstatisticaAba` sobrescreve `Estatística!C7:M126` — células que hoje contêm fórmulas — com valores literais. É o desenho da Fase 3 (motor único em VBA), mas é irreversível e ainda não aconteceu no arquivo entregue (validei sem salvar). Convém decidir conscientemente antes de rodar no arquivo de produção.
---

# PASSO 3 — Vazamento de estado: CORRIGIDO

`priR`/`priRun`/`ultR`/`ultRun` agora são reiniciados por grupo, junto com `nv` e `maxZ`.

**Validação:** `grupos_com_0_violacoes_reportando_regra = 0` (esperado 0).
Antes da correção, todo grupo processado após o primeiro com violação herdava a
regra e o RUN do grupo anterior — falso positivo silencioso em até 62 grupos.

# PASSO 4 — Destruição de fórmulas: CONFIRMADA

Contagem de células com fórmula, antes e depois de uma execução de `AtualizarEstatistica`:

| Aba | Antes | Depois | Perda |
|---|---|---|---|
| Painel | 58 | 13 | **45** |
| Estatística | 2.160 | 840 | **1.320** |
| Resultados | 0 | 0 | 0 |
| Liberação | 403 | 403 | 0 |
| Analitos | 658 | 658 | 0 |
| DB_Resultados | 45.190 | 45.190 | 0 |

**Perda total: 1.365 fórmulas — irreversível.**

`AtualizarPainelEng` e `AtualizarEstatisticaAba` gravam valores literais sobre células
que continham fórmulas. É coerente com o desenho da Fase 3 (motor único em VBA, Excel
guarda só o resultado), mas **não foi decisão explicitamente aprovada** e não tem volta:
uma vez executado no arquivo de produção, a camada de fórmulas deixa de existir.

**Tempos após as correções:** `RegistrarEventosWestgard` 0,098 s · `AtualizarEstatistica` 0,341 s.

## Decisão necessária antes de promover
- **(A)** Aceitar: motor VBA é a única fonte de cálculo. Painel e Estatística viram saída.
  Exige que o motor rode sempre, senão as abas ficam com dados congelados sem aviso.
- **(B)** Preservar fórmulas: motor escreve em área própria e as abas passam a referenciá-la.
  Mantém a planilha auditável sem macro, mas duplica a lógica de apresentação.
