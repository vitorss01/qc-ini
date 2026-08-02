# Fase 2 — Operação Laboratorial

**Escopo:** fluxo operacional diário via UserForms. Sem auditoria, backup, logs, estatística, painel, EQC ou gráficos (fases 3–6).
**Backup pré-fase:** `_backup_pre_fase2/`

---

## 1. Arquitetura

```
┌─────────────────────────── INTERFACE ───────────────────────────┐
│  frmCorrida        frmMassa            frmExcluir               │
│  (1 corrida)       (importação)        (exclusão lógica)        │
└────────────────────────────┬────────────────────────────────────┘
                             │  nenhum form escreve em célula direto
┌────────────────────────────▼────────────────────────────────────┐
│  mEntrada   ParseNum · ParseData · CodigoLote · NucleoLote      │
│             (parsing e validação de entrada)                     │
└────────────────────────────┬────────────────────────────────────┘
┌────────────────────────────▼────────────────────────────────────┐
│  mDados     CarregarDB · NovoRUN · UpsertResultados             │
│             ExcluirLogico · RunsDoLote · ListaAnalitos          │
└────────────────────────────┬────────────────────────────────────┘
┌────────────────────────────▼────────────────────────────────────┐
│  DB_Resultados   ← ÚNICA FONTE OPERACIONAL (vertical)           │
│  RUN │ Data │ Nível │ Lote │ Analito │ Resultado │ Status        │
└────────────────────────────┬────────────────────────────────────┘
┌────────────────────────────▼────────────────────────────────────┐
│  mOperacao  AtualizarViewResultados → AtualizarOperacao         │
└────────────────────────────┬────────────────────────────────────┘
┌────────────────────────────▼────────────────────────────────────┐
│  Resultados (VIEW, somente leitura)   N1 A:G │ N2 K:Q │ N3 U:AA │
│  Painel · Estatística  (intocados — fases 3+)                   │
└─────────────────────────────────────────────────────────────────┘
```

### Fluxo de gravação

```
Usuário digita/cola
      ↓
Validação (acumula TODAS as inconsistências: linha · coluna · motivo)
      ↓  erro? → lista completa, nada é gravado
      ↓  ok?
NovoRUN(Data, Lote)  → reutiliza RUN existente ou cria o próximo
      ↓
Horizontal → Vertical  (1 linha × N analitos vira N registros)
      ↓
UpsertResultados  → Dictionary chave RUN|Nível|Analito
      ├── existe → ATUALIZA (nunca duplica)
      └── novo   → acumula em buffer
      ↓
Escrita em LOTE (uma única atribuição de Range)
      ↓
AtualizarOperacao → view + estatística + painel
```

---

## 2. DB_Resultados e a view

A antiga aba `Resultados` foi **renomeada** para `DB_Resultados` — isso faz o Excel reapontar
sozinho todos os nomes definidos (`rRUN`, `rData`, `rStatus`, …) e fórmulas do `Calc`,
`Estatística` e `Liberação`, sem tocar em nada da Fase 1.

Uma aba **nova** `Resultados` foi criada como camada de visualização:

| Bloco | Colunas | Conteúdo |
|---|---|---|
| Nível 1 | A:G | RUN · Data · Lote · Analito · Resultado · Status |
| Nível 2 | K:Q | idem |
| Nível 3 | U:AA | idem (só Hematologia) |
| Botões | I:J | Nova Corrida · Inserção em Massa · Excluir Resultados · Atualizar View |

- Coluna **NC (X) removida** definitivamente.
- A view é **reconstruída em lote** (uma escrita por bloco), filtrada pelo lote em uso.
- Toda a aba fica com células bloqueadas — é somente leitura.
- Nenhum botão fica sobre células com dado, título ou cor.

---

## 3. UserForms

### frmCorrida — Nova Corrida
Data · Nível · **Lote** (novo) → carrega automaticamente todos os analitos cadastrados.
O RUN aparece em tempo real conforme Data/Lote mudam. Confirmação antes de gravar.
Atalho `Alt+S` para salvar; `Esc` cancela.

### frmExcluir — Exclusão lógica
Nível → lista os analitos com **checkbox por item** (ListBox em modo opção).
Botões **Selecionar todos** e **Limpar seleção**. Combo de **Corrida (RUN)** já filtrado
pelo lote em uso. Confirmação explícita avisando que nada é apagado — apenas
`Status = Excluído`.

> Implementei a lista com `ListBox` em `ListStyle = fmListStyleOption` em vez de 40 CheckBoxes
> dinâmicos: o efeito visual é o mesmo (um checkbox por analito), mas rola bem, não estoura o
> limite prático de controles do form e carrega instantaneamente.

### frmMassa — Inserção em massa
Área de colagem que aceita **Ctrl+V direto do Excel** (colunas separadas por TAB), na ordem
`Data | Nível | Lote | analitos…` exibida no cabeçalho. Aceita 1, 5, 30, 50+ linhas.

Botões **Validar** (só confere), **Gravar** e **Limpar grade**. Barra de progresso durante a
importação. A lista de erros mostra **linha · coluna · motivo** e **não para na primeira
inconsistência** — acumula todas.

> Optei por área de colagem + lista de erros em vez de uma grade editável célula a célula:
> uma grade real exigiria ~2.000 controles (50 linhas × 43 colunas), o que trava o VBA.
> A colagem direta do Excel é o caminho que o analista já usa no dia a dia.

**Validações:** data válida · nível dentro do intervalo · lote preenchido **e cadastrado** ·
resultado numérico (aceita `12,5` e `12.5`) · analito existente.

---

## 4. Chave de unicidade — decisão de projeto

O pedido dizia "mesmo RUN + Analito → atualizar". A chave implementada é
**RUN + Nível + Analito**.

Motivo: o RUN identifica a *corrida*, que abrange **todos os níveis**. Com RUN+Analito apenas,
gravar o Nível 2 sobrescreveria o resultado do Nível 1 do mesmo analito na mesma corrida —
perda silenciosa de dado. Com o Nível na chave, o comportamento pedido (não duplicar,
atualizar o existente) é preservado sem esse efeito colateral.

---

## 5. Performance

- Leitura do banco inteiro para memória em **uma** operação (`CarregarDB`).
- Índice em `Scripting.Dictionary` para o upsert — sem varredura repetida.
- Inserções acumuladas em buffer e gravadas com **uma** atribuição de `Range`.
- View reconstruída com **uma escrita por bloco**.
- Sem `Select`, sem `Activate`, sem `Cells` dentro de laço de escrita.
- `ScreenUpdating` desligado apenas nos trechos de escrita.

---

## 6. Módulos

| Módulo | Situação | Papel |
|---|---|---|
| `mDados` | **reescrito** | Banco `DB_Resultados`, RUN, upsert, exclusão lógica, listas |
| `mOperacao` | **novo** | View em blocos, `AtualizarOperacao`, abertura dos forms |
| `mEntrada` | **novo** | Parsers tolerantes a locale e helpers de lote |
| `frmCorrida` | **reconstruído** | + seletor de Lote, grava pela camada de dados |
| `frmExcluir` | **novo** | Exclusão lógica |
| `frmMassa` | **novo** | Importação em massa |
| `mSeguranca`, `mLotes`, `mUI` | intocados | Fase 1 preservada |
| `clsCht`, `frmAssinar`, `frmDev` | intocados | Fase 1 preservada |

---

## 7. Testes executados

| Teste | Resultado |
|---|---|
| Fase 1 intacta (Painel n antes/depois) | 25 → 25 |
| Nomes definidos reapontados p/ DB_Resultados | rRUN/rData/rStatus/rLote OK |
| View popula por bloco | 525 linhas/bloco (21 analitos × 25 corridas) |
| Carga dos 3 forms | frmCorrida 67 · frmExcluir 12 · frmMassa 11 controles |
| frmExcluir lista analitos e RUNs | 28 analitos · 25 RUNs |
| Inserção de 2 registros novos | novos=2 · atualizados=0 |
| **Reenvio da mesma chave** | novos=0 · atualizados=2 · **delta de linhas = 2** |
| Valor efetivamente atualizado | 99,9 ✔ |
| Exclusão lógica | 1 registro marcado |
| Varredura de erros de fórmula | **NENHUM** |

---

## 8. Armadilha registrada

**`Worksheet.Move(After=...)` apaga a aba.** Pelo win32com, o argumento nomeado não vincula;
o Excel entende `Move()` sem argumentos, cujo comportamento padrão é mover a aba para uma
**pasta de trabalho nova** — ela desaparece da atual, sem erro nenhum. Foi assim que a view
sumiu na primeira tentativa. Correção: usar posicional (`ws.Move(wsAntes)`) e conferir depois
que a aba continua existindo.

---

## 9. Não implementado (próximas fases)

Auditoria, backup, autosave, logs de usuário, melhorias estatísticas, painel, EQC e gráficos —
conforme combinado. `RegistrarLog` existe como gancho vazio, já chamado nos pontos de
gravação e exclusão, para a Fase 5 apenas preencher.
