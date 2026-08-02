# Especificação — Resultados Não Conformes e Auditoria de Exclusões

**Origem:** requisito do gestor, 02/08/2026.
**Escopo de gate:** não cria item novo. Concretiza os já abertos **3.1** (trilha de
auditoria), **3.2** (log em todo caminho de gravação) e a pendência **F3-3** (`frmNaoConforme`).
**Sprint:** HARDENING 3.

---

## 1. Regra de arquitetura que tudo abaixo respeita

Nenhum formulário escreve em célula. Toda inclusão, alteração, movimentação ou exclusão
passa por `mDados` (ADR-001). O log é emitido **de dentro do `mDados`** — nenhum caminho
de gravação escapa.

---

## 2. O que "transferir para Registros" significa aqui

O requisito diz *transferir para `Registros`* e também *nada é apagado do banco;
a exclusão ocorre só na interface*. Os dois se conciliam de uma forma só:

- a linha **permanece** em `DB_Resultados`, com o mesmo RUN, o mesmo lote, o valor
  original intacto, e o `Status` alterado para um estado **não elegível** do `Cfg_Status`;
- por não ser elegível, some da view `Resultados` e sai de todos os cálculos;
- uma **cópia de exibição** é escrita em `Registros`, para leitura e histórico.

**Por que não mover fisicamente.** A linha movida perderia o vínculo com a corrida que a
gerou — RUN, lote, nível — e o registro que se quis preservar deixaria de ser rastreável.
É o oposto do objetivo, e contraria o ADR-003 e a ISO 15189 §8.4.2 (valor original
recuperável após a emenda). `Registros` é vitrine; `DB_Resultados` é a verdade.

---

## 3. Formulário `frmNaoConforme` (aba Resultados)

Botão **"Registrar Resultado Não Conforme"**, mesmo padrão visual do `frmExcluir`.

| Campo | Comportamento |
|---|---|
| RUN | combo, via `RunsDoLote` (chave principal) |
| Nível | combo 1..NLV |
| Analito | lista de seleção |
| **Painel de conferência** | exibe Data · RUN · Nível · Lote · Analito · Resultado · Status atual **antes** de confirmar |
| **Tipo de não conformidade** | combo lido de `Cfg_Status` — nenhum estado no código |
| **Parecer Técnico da Exclusão** | multilinha, **obrigatório**; sem ele o botão confirmar fica desabilitado |

Sugestões oferecidas na interface (o usuário pode editar e complementar): controle
contaminado · erro comprovado de pipetagem · material hemolisado · repetição após
recalibração · resultado lançado incorretamente · falha operacional · controle vencido ·
troca de lote · equipamento com erro durante a corrida · outro (descrever).

---

## 4. Aba `Registros` — nova estrutura

`Rep 1 · Rep 2 · Rep 3` → **`NC 1 · NC 2 · NC 3`**

- o resultado não conforme entra automaticamente em **NC 1**
- **NC 2** e **NC 3** ficam livres para repetições posteriores, digitadas à mão
- inserção a partir de **B4**, primeira linha livre
- campos transferidos: Data · RUN · Analito · Nível · Resultado · Lote · Equipamento
  (quando existir) · tipo de NC

**Impacto a tratar na implementação:** `mEstatistica` lê `Registros` para os marcadores
do gráfico (repetição = X roxo; calibração = ícone), em `mEstatistica.bas:461` e `:537`.
A semântica muda de *repetição* para *não conformidade* e o marcador precisa acompanhar.
Dados existentes: 2 linhas de demonstração — migração trivial, mas obrigatória e
verificada nos três produtos (`RegistrosStore` por lote).

---

## 5. `frmExcluirRegistro` (aba Registros)

Espelha o `frmExcluir`. Confirmação explícita antes de executar. Exige Parecer Técnico.
Remove da vitrine `Registros`; **nunca** do `DB_Resultados`.

---

## 6. `DB_Exclusoes` — trilha de auditoria

Aba própria, **append-only**, nunca editável pela interface.

> Aba separada, não uma segunda tabela dentro de `DB_Resultados`: o banco é lido em
> bloco único (`CarregarDB`), e uma tabela vizinha na mesma planilha quebraria essa
> leitura.

Colunas:

`ID_Auditoria` · `Data/Hora da operação` · `Tipo de operação` · `RUN` · `Data da corrida` ·
`Nível` · `Analito` · `Lote` · `Equipamento` · `Resultado` · `Status anterior` ·
`Status novo` · `Parecer Técnico` · `Usuário do sistema` · `Usuário Office` ·
`Usuário Windows` · `Nome do computador` · `Nome do arquivo` · `Versão do sistema`

`ID_Auditoria` é único e não reaproveitável. Captura sem API externa:
`Application.UserName`, `Environ$("USERNAME")`, `Environ$("COMPUTERNAME")`,
`ThisWorkbook.Name`, mais uma constante `VERSAO_SISTEMA`.

---

## 7. Aplicação global

Todo formulário que exclua qualquer coisa — `frmExcluir`, `frmExcluirRegistro`,
`frmNaoConforme` e os que vierem — obedece ao mesmo contrato: confirmação explícita,
Parecer Técnico obrigatório, log completo em `DB_Exclusoes`, gravação só por `mDados`.
**Não existe exclusão sem rastreabilidade.**

---

## 8. Validação obrigatória

1. Unitário: `mDados` recusa exclusão sem parecer
2. Unitário: `ID_Auditoria` único sob execuções repetidas
3. Integração: NC → sai da view, sai do cálculo, aparece em `Registros`, gera log
4. Estatística: média/DP/CV/Bias/ET/Sigma/Z/Westgard recalculados **sem** o valor NC
5. Interface: confirmar impossível com o Parecer vazio ou só com espaços
6. Migração `Rep`→`NC` sem perda, nos três produtos
7. Marcadores do gráfico coerentes com a nova semântica
8. Regressão completa das Fases 1 e 2
9. Idempotência e persistência (itens 4.8 e 4.9)
