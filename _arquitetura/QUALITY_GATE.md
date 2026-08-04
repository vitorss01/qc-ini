# QUALITY GATE — Sistema QC (CQI laboratorial)

Documento oficial de liberação do projeto. Enquanto houver item ⏳ ou ❌ **não se avança
para a fase seguinte**. Cada ✅ exige evidência verificável — marcação sem evidência
não vale.

**Última atualização:** 03/08/2026
**Sprint corrente:** HARDENING (nenhuma funcionalidade nova)

---

## 1. Motor estatístico

| # | Item | Status | Evidência |
|---|---|---|---|
| 1.1 | Projeto compila | ✅ | Execução de procedimento força compilação; 3 erros de sintaxe (`aS`) corrigidos. **Este item é a causa raiz do 2.1** — ver nota abaixo |
| 1.2 | Nenhum travamento | ✅ | `RegistrarEventosWestgard` 0,098 s · `AtualizarEstatistica` 0,341 s |
| 1.3 | Média / DP / CV% | ✅ | Idêntico a referência Python em 6 casas: 14,000000 · 3,162278 · 22,587698 |
| 1.4 | Bias / ET / Sigma / Z | ✅ | 40,000000 · 21,500000 · 5,000000 · 1,000000 |
| 1.5 | Westgard 12s/13s/22s/R4s/41s/10x | ✅ | 9/9 asserções, incluindo níveis opostos → R4s e **não** 22s |
| 1.6 | Elegibilidade CLSI EP05/C24 | ✅ | 6/6 estados; desconhecido nunca entra no cálculo |
| 1.7 | Robustez (n=0, n=1, DP=0, alvo=0) | ✅ | Todos devolvem 0, sem erro |
| 1.8 | Sem vazamento de estado entre grupos | ✅ | `grupos_com_0_violacoes_reportando_regra = 0` |

> **Nota — por que as fórmulas foram destruídas (liga 1.1 a 2.1).**
> `aS` não é identificador válido: VBA é insensível a maiúsculas, então `aS` é o
> mesmo token que a palavra reservada `As`, e o módulo não compila. Como o VBA
> compila sob demanda, procedimento a procedimento, o defeito ficava invisível —
> `AtualizarCalc` rodava normalmente enquanto `AtualizarPainelEng`,
> `AtualizarEstatisticaAba` e `RegistrarEventosWestgard` morriam com "não foi
> possível executar a macro". **Na produção esses três nunca rodaram, e é por
> isso que as fórmulas do `Painel` e da `Estatística` sobreviveram lá.** Ao
> corrigir o `aS`, a Fase 3A deu vida aos três e eles passaram por cima das
> fórmulas. O gate registrava a correção em 1.1 e o estrago em 2.1 sem ligar um
> ao outro. `gerar_mEstatistica.ps1` agora renomeia `aS`/`aM` para
> `alvoS`/`alvoM` e tem lint que barra identificador colidindo com reservada.

## 2. Integridade dos dados

| # | Item | Status | O que fecha |
|---|---|---|---|
| 2.1 | Fórmulas preservadas | ✅ | Marcos 1–5. Motor escreve só em `Eng_Saida`. Diff célula a célula contra a produção: `AUSENTE 0 · ALTERADA 4362 · VALOR 0 · EXTRA 9`, 61.425 fórmulas. As 4362 são lista fechada e versionada (`src_hardening1/marco*_celulas_redirecionadas.csv`): 3240 Calc + 42 Painel + 1080 Estatística. As 9 EXTRA são `Painel!S:U`, onde o motor gravava valor direto |
| 2.1b | ADR-019 verificado | ✅ | `varredura_adr019.ps1`: produção tinha **120** fórmulas de interface calculando parâmetro estatístico sobre o banco (`Estatística!D`, `AVERAGEIFS`); build tem **0**. As 900 restantes no `Calc` são seleção/filtro, permitidas. Estatística agregada que sobra na interface é só a cadeia do bias EQC (840, mantida por decisão) e 8 células de limite de eixo |
| 2.2 | Sem sobrescrita silenciosa | ⏳ | Reenviar mesma Data+lote sobrescreve valor sem versionar. Exige versionamento: linha original nunca alterada |
| 2.3 | Exclusão lógica não é revertida sozinha | ✅ | `UpsertResultados` força `Status = Ativo`, ressuscitando excluído. Reverter exclusão deve ser ação própria, com justificativa |
| 2.4 | RUN representa a corrida analítica | ⏳ | Hoje RUN = (Data + lote) colapsa turnos e pós-calibração. CLSI C24 trata como eventos distintos |
| 2.5 | `Cfg_Status` não altera histórico retroativamente | ⏳ | Uma célula redefine o que entra na estatística de todo o histórico. Exige versionamento da tabela + log |


> **Por que 2.1 e 2.1b voltaram para ⏳ (03/08/2026).**
> A evidência que os fechou era verdadeira sobre as *planilhas* e falsa sobre o
> *sistema*: as 4.362 células foram mesmo redirecionadas para `Eng_Saida`, mas o
> **motor corrigido nunca entrou no `.xlsm`** — nem aqui, nem na máquina do
> trabalho. Verificado lendo o VBA de dentro dos três artefatos: `mEstatistica`
> com 1.028 linhas, 8 ocorrências de `aS` e zero referência a `Eng_Saida`,
> quando o módulo versionado tem 1.049 linhas e nenhuma `aS`. As fórmulas
> apontavam para uma camada que ninguém abastecia.
>
> Quatro defeitos no pipeline, nenhum em VBA nem em lógica de CQI:
> 1. `Select-Object -First` **encerra o script produtor** (`StopUpstreamCommandsException`).
>    As etapas morriam depois das primeiras mensagens e **antes** do `Save()`.
>    `criar_eng_saida` usava `-Last`, que bufferiza — por isso só ela persistia.
> 2. `$xl.Quit()` não mata o processo enquanto o PowerShell segura referências COM.
>    O Excel órfão travava o arquivo e a etapa seguinte o abria em **somente leitura**,
>    sem aviso, porque `DisplayAlerts = $false`.
> 3. Nenhum dos 12 scripts tinha `$ErrorActionPreference = 'Stop'`: erro de COM é
>    não-terminante, então falha virava mensagem de sucesso.
> 4. `aplicar_vba.ps1` imprimia "VBA aplicado" sem conferir nada.
>
> **Lição estrutural:** havia lint de VBA, diff célula a célula e varredura de
> arquitetura — e nenhum controle verificava se a etapa anterior tinha acontecido.
> O pipeline reportava intenção, não resultado. É o mesmo princípio que o gestor
> exigiu para as fórmulas: não basta restaurar, precisa comparar.
>
> Reabrir estes dois itens é o certo. Gate que aceita evidência sem lastro dá
> confiança falsa, que é pior do que não ter gate.

## 3. Auditoria e segurança

| # | Item | Status | O que fecha |
|---|---|---|---|
| 3.1 | Trilha de auditoria | ✅ | `RegistrarLog` é stub vazio. Aba `Audit_Log` append-only: data/hora, usuário, papel, ação, chave, valor ANTES, valor DEPOIS, justificativa |
| 3.2 | Log em todo caminho de gravação | ✅ | Chamar de **dentro do `mDados`**, não dos formulários — senão algum caminho escapa |
| 3.3 | Proteção persistida no arquivo salvo | ✅ | Nenhuma das 18 abas tem `<sheetProtection>`. `UserInterfaceOnly:=True` **não persiste** entre sessões |
| 3.4 | Projeto VBA com senha | ⏳ | DPB decodifica para payload vazio — sem senha |
| 3.5 | Arquivo distribuído em estado bloqueado | ✅ | Salvo em sessão autenticada: abrir com macros desabilitadas dá acesso total a `DB_Resultados` |


> **Reabertos e agora fechados com evidência que se sustenta (03/08/2026).**
> Os itens 2.1 e 2.1b tinham sido rebaixados porque o motor validado nunca
> estivera dentro do `.xlsm`. Agora fecham porque `verificar_tudo.ps1` prova, na
> **mesma execução**, que (a) o VBA dentro do arquivo é idêntico byte a byte à
> fonte versionada — `mEstatistica`, `mDados` e `mAuditoria`, por hash — e (b) o
> diff das fórmulas dá `AUSENTE 0 · ALTERADA 4362 · VALOR 0 · EXTRA 9`. Sem (a),
> o (b) não significava nada.
>
> **3.1 e 3.2** fecham com o teste que dá sentido à trilha: a suíte grava no log,
> **adultera uma linha fora do sistema** e exige que a verificação acuse a linha
> certa — `QUEBRADO|5|conteudo alterado` — e que a cadeia volte a fechar quando o
> valor é restaurado. "Append-only" deixou de ser promessa e virou propriedade
> demonstrável diante de um auditor.
>
> **2.3 fecha com teste, não com leitura de código.** Os itens 3.6 e 3.7 da suíte
> percorrem o vetor de fraude inteiro: marcam um registro real como `Excluído`,
> **reenviam a mesma chave pela camada de dados** — como um formulário faria — e
> exigem que o status continue `Excluído` e que a tentativa apareça no log como
> `UPSERT_BLOQUEADO`. Depois restauram o estado original.
>
> **Como reproduzir tudo isto:** `.\_arquitetura\scripts_fase3erificar_tudo.ps1`
> — um comando, 21 verificações, relatório em `VERIFICACAO.md` e código de saída
> diferente de zero se algo falhar. Roda sem agente.


> **3.3 e 3.5 fecham; 3.4 continua ⏳ de propósito (03/08/2026).**
> A proteção era aplicada só em tempo de execução: `Workbook_Open` chama
> `LockApp`, que usa `UserInterfaceOnly:=True`. Duas falhas somadas —
> **`Workbook_Open` não roda com macros desabilitadas**, e `UserInterfaceOnly`
> é atributo de sessão, não persiste. Como o arquivo era salvo em sessão de ADM
> (que roda `UnprotectAll`), ele saía **distribuído destravado**.
>
> Agora `blindar_artefato.ps1` produz o entregável já travado: 21 abas
> protegidas com senha **gravada no arquivo**, 20 em `veryHidden` (só `Login`
> visível) e estrutura da pasta protegida. Com macros habilitadas nada muda para
> o usuário — o login revela o que o papel permite.
>
> **A verificação lê o `.xlsm` como zip e inspeciona o XML**, em vez de perguntar
> ao Excel. Se a pergunta é "o que o auditor vê com macros desligadas",
> perguntar ao Excel com macros ligadas não responde.
>
> **3.4 é passo manual e vai continuar sendo.** Pôr senha no projeto VBA exige
> escrever os campos `DPB`/`DPx` do `vbaProject.bin`, e errar corrompe o projeto.
> Faça no VBE: *Ferramentas → Propriedades do VBAProject → Proteção*. A suíte
> detecta e reporta enquanto não for feito.
>
> **Limite declarado:** proteção de planilha é barreira, não cofre — a senha está
> em texto no VBA e a marcação sai descompactando o arquivo. Ela detém o acidente
> e o curioso, não a fraude decidida. A garantia de integridade está na trilha
> encadeada por hash.

## 4. Qualidade e testes

| # | Item | Status | O que fecha |
|---|---|---|---|
| 4.1 | Testes unitários do motor | ✅ | Suíte em `scripts_fase1/` e `scripts_fase2/` |
| 4.2 | QA integrado | ⏳ | **Nunca executou** — agente interrompido por limite de gastos |
| 4.3 | Regressão Fases 1 e 2 pós-Fase 3 | ⏳ | RUN, exclusão lógica, troca de lote, login, 3 UserForms, upsert sem duplicar |
| 4.4 | Casos extremos | ⏳ | Banco vazio · 1 resultado · 2 resultados · analito sem média/DP · lote sem resultado · status inválido · RUN duplicado · data futura |
| 4.5 | Teste de estresse | ⏳ | 5.000 RUNs · 40 analitos × 3 níveis · 50 lotes · importação de 500 linhas |
| 4.6 | Zero erro de fórmula | ⏳ | Revalidar após as correções de 2.x |
| 4.7 | Estado global restaurado | ⏳ | `ScreenUpdating`/`EnableEvents`/`Calculation` ao fim de toda rotina |
| 4.8 | Idempotência | ⏳ | `AtualizarEstatistica` + `RegistrarEventosWestgard` executados 1×, 10× e 100× seguidos sobre o mesmo banco produzem saída **byte a byte idêntica**. Pega vazamento de estado, acúmulo em coleções e linhas duplicadas em `Eventos_Westgard` |
| 4.9 | Persistência entre sessões | ⏳ | Gerar → salvar → fechar o Excel → reabrir → gerar de novo → resultado idêntico. Expõe dependência de estado em memória, variável `Static`, `UserInterfaceOnly` e cache não persistido |

## 5. Cobertura por produto

| Produto | Fase 1 | Fase 2 | Fase 3 | Observação |
|---|---|---|---|---|
| Hematologia | ✅ | ✅ | ⏳ | Motor instalado; correção validada **não promovida** |
| Bioquímica | ✅ | ✅ | ❌ | Fase 3 nunca aplicada |
| Imunologia | ✅ | ✅ | ❌ | Fase 3 nunca aplicada |

## 6. UX

| # | Item | Status | Observação |
|---|---|---|---|
| 6.1 | Fluxo diário validado com usuário | ⏳ | Simulação feita por agente; falta validação real |
| 6.2 | Padronização "Corrida" → "RUN" | ⏳ | Pendente em gráficos, mensagens e cabeçalhos |
| 6.3 | Eixo X alinhado ao RUN | ⏳ | Pendente |
| 6.4 | 52 achados de UX/arquitetura triados | ⏳ | `FASE3A_achados_estaticos.md` |

---

## Resumo

> **Regra de merge:** nada entra na `main` enquanto houver ⏳ ou ❌.
> A branch `fase3a-motor-cqi` é de trabalho; se um PR for aberto, deve ser **draft**.

**22 ✅ · 20 ⏳ · 4 ❌**  (contagem inclui a tabela de cobertura por produto)

O **motor** está pronto e validado. O **sistema** ainda não é auditável nem confiável.

---

# Anexo — Redesenho proposto do RUN (item 2.4)

Modelo atual: `RUN = f(Data, Lote)` — determinístico, bom para deduplicação, mas
**incapaz de representar duas corridas no mesmo dia** (turno, pós-calibração,
pós-manutenção), que a CLSI C24 trata como eventos distintos.

Modelo proposto: **RUN = inteiro sequencial**, com os atributos da corrida em colunas
próprias:

```
RUN | Data | Hora | Turno | Lote | Equipamento | Usuário
```

Assim `RUN 501 (10/08 08:00)` e `RUN 502 (10/08 13:20)` coexistem naturalmente.

### Pré-requisito antes de escrever código
O custo aqui **não é o código, é a migração**. Os três produtos já têm dados gravados
com o RUN atual (Hematologia 1.575 linhas, Bioquímica e Imunologia 1.000 cada).
Trocar a semântica da chave exige:

1. Script de migração que preserve a correspondência RUN antigo → RUN novo
2. Backup obrigatório antes (a migração é irreversível)
3. Revalidar tudo que consome RUN: `Calc`, `Liberação`, `Eventos_Westgard`, os 3 UserForms
4. Decidir o que fazer com corridas históricas sem hora registrada
   (sugestão: hora nula = turno único, preservando o comportamento atual)

Planejar a migração **antes** de executar, não durante.

---

# Ordem de execução da Sprint HARDENING

1. **2.1 Fórmulas** — muda onde o motor escreve e afeta o teste de todo o resto. Primeiro.
2. **2.4 RUN** — mexe no modelo de dados. Depois das fórmulas, antes da auditoria.
3. **3.1–3.2 Auditoria** — precisa da chave final do RUN para logar corretamente.
4. **2.2–2.3 Versionamento e exclusão** — apoiam-se na trilha de auditoria.
5. **3.3–3.5 Proteção** — por último nas correções: muda o estado de distribuição.
6. **4.2–4.7 QA completo** — só faz sentido com tudo acima fechado.
7. **Propagar para Bioquímica e Imunologia** — o pipeline é paramétrico; propagar apenas
   depois de Hematologia passar em todo o gate.

---

# Roadmap até o Release Candidate

**Arquitetura congelada.** Nenhuma funcionalidade nova — nem dashboard, nem indicador,
nem UserForm, nem gráfico — até RC1. Só correção, teste e validação.

### Sprint HARDENING 1 — Fórmulas
- Criar `Eng_Saida`; motor deixa de escrever em `Painel` e `Estatística`
- Restaurar as 1.365 fórmulas destruídas (fonte: `_backup_pre_fase3/`)
- Painel e Estatística voltam a consumir por fórmula, a partir de `Eng_Saida` (ADR-019)
- **Validar a restauração, não confiar nela.** Diff automatizado célula a célula contra o
  backup: endereço, fórmula em R1C1 e resultado. Aceite = 100% idêntico; qualquer
  divergência é listada, não resumida em contagem
- **Varredura ADR-019:** nenhuma fórmula de interface referencia `DB_Resultados`. Esperado: zero
- **Aceite:** contagem **e conteúdo** idênticos · zero referência direta ao banco · regressão completa

> **Histórico não é reprocessado.** A unificação do Westgard e a correção de
> ET/Sigma mudam veredictos passados. Decisão do gestor: **não** reprocessar o
> histórico — o efeito cascata é maior que o benefício. A partir da RC1, as
> regras Westgard e o cálculo de Erro Total e Sigma passam a ser determinados
> **exclusivamente pelo motor**. Evidência do impacto medido fica em
> `src_hardening1/marco2_impacto_westgard.csv`.

### Sprint HARDENING 2 — Modelo do RUN
- RUN sequencial + `Data | Hora | Turno | Lote | Equipamento | Usuário` (ADR-011)
- Script de migração dos 3.575 registros existentes, com backup obrigatório
- **Aceite:** unicidade garantida · 3 UserForms validados · nada órfão

### Sprint HARDENING 3 — Auditoria (ISO 15189)
- `Audit_Log` append-only: usuário, timestamp, ação, chave, valor antes, valor depois, motivo
- Log chamado de dentro de `mDados` — nenhum caminho de gravação escapa
- **Aceite:** toda gravação e exclusão rastreável · log não editável pela interface

### Sprint HARDENING 4 — Validação
- Estresse · regressão · usabilidade · desempenho
- **Idempotência (4.8):** 1× · 10× · 100× execuções consecutivas → saída idêntica
- **Persistência (4.9):** gerar → fechar → reabrir → gerar → resultado idêntico
- **Teste com macros desabilitadas** (o cenário do auditor)
- Teste de recuperação após falha
- **Aceite:** Quality Gate 100% verde

### Release Candidate (RC1)
Somente quando **todos** forem verdadeiros:
- Quality Gate 100% verde
- Nenhum bug crítico aberto
- Regressão 100%
- Auditoria completa e funcional
- Desempenho validado sob carga
- **Hematologia, Bioquímica e Imunologia executando arquitetura idêntica**

A partir do RC1: implantação piloto, e toda mudança passa a ser evolução funcional —
não mais correção estrutural.
