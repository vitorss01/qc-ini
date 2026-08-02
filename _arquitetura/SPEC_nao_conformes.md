# Especificação — RUN Não Conforme e Auditoria de Exclusões

**Origem:** requisito do gestor, 02/08/2026 (revisão 2 — substitui a versão anterior).
**Escopo de gate:** não cria item novo. Concretiza **3.1** (trilha de auditoria),
**3.2** (log em todo caminho de gravação) e a pendência **F3-3**.
**Sprint:** HARDENING 3.

---

## 1. Regra que tudo abaixo respeita

Nenhum formulário escreve em célula. Toda inclusão, alteração, movimentação ou exclusão
passa por `mDados` (ADR-001), e o log é emitido **de dentro do `mDados`** — nenhum
caminho de gravação escapa. Não existe exclusão sem rastreabilidade.

---

## 2. `Registros` passa a ser aba derivada

**Estrutura nova.** Saem `Rep 1 · Rep 2 · Rep 3`. Entra **uma coluna só: `RUN NC`**.

Layout a partir de **B4** (primeira linha gravável):

`Nº · Data · RUN NC · Nível · Analito · Resultado · Lote · Tipo de NC · Parecer Técnico ·
Usuário · Data/Hora do registro`

**Digitação manual deixa de existir.** Todo conteúdo chega automaticamente pelo
formulário da aba `Resultados`. A aba fica protegida contra edição direta. Isso remove a
classe inteira de erro em que um valor aparece em `Registros` sem corresponder a nenhuma
linha do banco.

---

## 3. `frmRunNC` — "Registrar RUN Não Conforme" (aba Resultados)

Padrão visual do `frmExcluir`.

| Etapa | Comportamento |
|---|---|
| Informar o **RUN** | é a chave: o formulário **preenche sozinho** Data, Lote e Nível |
| Selecionar analitos | um, vários ou **todos** |
| Puxar resultados | automático — o RUN identifica a corrida, então o valor vem do banco |
| Conferência | exibe Data · RUN · Nível · Lote · Analito · Resultado antes de confirmar |
| Tipo de NC | combo lido do `Cfg_Status` — nenhum estado escrito em código |
| **Parecer Técnico** | multilinha, obrigatório, **mínimo 5 palavras**; abaixo do resultado |

Sem parecer válido o botão salvar fica desabilitado. A contagem ignora espaços múltiplos
e exige palavras reais, não 5 caracteres soltos.

**Ao salvar:**

1. o registro **sai da aba `Resultados`** — `Status` passa a um estado não elegível
2. a linha **permanece no `DB_Resultados`**, com RUN, lote e valor original intactos
3. deixa de entrar em média, DP, CV%, Bias, Erro Total, Sigma, Z-Score e Westgard
4. aparece em `Registros`, na primeira linha livre a partir de **B4**

> **Por que a linha não sai do banco.** Movida fisicamente, ela perderia o vínculo com a
> corrida que a gerou e o registro que se quis preservar deixaria de ser rastreável —
> contrário ao ADR-003 e à ISO 15189 §8.4.2 (valor original recuperável após emenda).
> `Registros` é vitrine; `DB_Resultados` é a verdade.

---

## 4. `frmExcluirRegistro` — "Excluir Registro NC" (aba Registros)

O analista pode se enganar ao classificar. O formulário desfaz isso com controle:

- confirmação explícita: *"Deseja realmente excluir este registro?"*
- **Parecer Técnico obrigatório**, mesma regra de 5 palavras
- some da aba `Registros`
- **continua no `DB_Resultados`** — nada é apagado

---

## 5. Duas tabelas de auditoria no `DB_Resultados`

Conforme decidido pelo gestor, as duas tabelas ficam **na própria aba `DB_Resultados`**:

| Tabela | Recebe |
|---|---|
| `LOG_Resultados` | exclusões e marcações de NC feitas na aba `Resultados` |
| `LOG_Registros` | exclusões feitas na aba `Registros` |

**Implementação sem quebrar a leitura do banco.** `CarregarDB` lê o bloco `A:G` e
`UltimaLinhaBanco` mede a coluna A. As duas tabelas ficam em **blocos de colunas
deslocados à direita**, com intervalos nomeados próprios, então a leitura do banco
permanece intacta e o requisito é atendido literalmente.

Colunas, idênticas nas duas (para que uma visão de auditoria possa uni-las):

`ID_Auditoria` · `Data/Hora da operação` · `Tipo de operação` · `Aba de origem` · `RUN` ·
`Data da corrida` · `Nível` · `Analito` · `Lote` · `Resultado` · `Status anterior` ·
`Status novo` · `Parecer Técnico` · `Usuário do sistema` · `Usuário Office` ·
`Usuário Windows` · `Nome do computador` · `Nome do arquivo` · `Versão do sistema`

`ID_Auditoria` é único e nunca reaproveitado. Ambas são **append-only** e não editáveis
pela interface. Captura sem dependência externa: `Application.UserName`,
`Environ$("USERNAME")`, `Environ$("COMPUTERNAME")`, `ThisWorkbook.Name` e a constante
`VERSAO_SISTEMA`.

---

## 6. Aplicação global

Todo formulário que exclua qualquer coisa — `frmExcluir`, `frmExcluirRegistro`,
`frmRunNC` e os que vierem — obedece ao mesmo contrato: confirmação explícita, Parecer
Técnico obrigatório de no mínimo 5 palavras, log completo, gravação só por `mDados`.

---

## 7. Impactos mapeados antes de codificar

1. **Gráfico.** `mEstatistica` lê `Registros` para os marcadores (X roxo = repetição,
   ícone = calibração), em `mEstatistica.bas:461` e `:537`. Repetições deixam de existir
   nessa aba; o marcador passa a significar **não conformidade** e o código acompanha.
2. **Dados existentes.** 2 linhas de demonstração em `Registros`, 3 em `RegistrosStore`.
   Migração trivial, mas obrigatória e verificada nos três produtos.
3. **`RegistrosStore`.** A aba é por lote (ADR de armazém). O armazém continua; muda o
   layout que ele guarda.
4. **Proteção.** `Registros` passa a protegida de verdade — depende do item 3.3 do gate.

---

## 8. Validação obrigatória

1. `mDados` recusa exclusão sem parecer, e recusa parecer com menos de 5 palavras
2. `ID_Auditoria` único sob execuções repetidas
3. RUN informado preenche Data, Lote, Nível e Resultado corretamente, sem digitação
4. NC → sai da view, sai do cálculo, aparece em `Registros`, gera log
5. Média/DP/CV/Bias/ET/Sigma/Z/Westgard recalculados **sem** o valor marcado
6. Excluir da aba `Registros` **não** remove do `DB_Resultados`
7. `Registros` recusa digitação manual
8. Migração `Rep`→`RUN NC` sem perda, nos três produtos
9. Marcadores do gráfico coerentes com a nova semântica
10. Regressão completa das Fases 1 e 2 · idempotência (4.8) · persistência (4.9)
