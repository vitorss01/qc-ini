# QUALITY GATE — Sistema QC (CQI laboratorial)

Documento oficial de liberação do projeto. Enquanto houver item ⏳ ou ❌ **não se avança
para a fase seguinte**. Cada ✅ exige evidência verificável — marcação sem evidência
não vale.

**Última atualização:** 01/08/2026
**Sprint corrente:** HARDENING (nenhuma funcionalidade nova)

---

## 1. Motor estatístico

| # | Item | Status | Evidência |
|---|---|---|---|
| 1.1 | Projeto compila | ✅ | Execução de procedimento força compilação; 3 erros de sintaxe (`aS`) corrigidos |
| 1.2 | Nenhum travamento | ✅ | `RegistrarEventosWestgard` 0,098 s · `AtualizarEstatistica` 0,341 s |
| 1.3 | Média / DP / CV% | ✅ | Idêntico a referência Python em 6 casas: 14,000000 · 3,162278 · 22,587698 |
| 1.4 | Bias / ET / Sigma / Z | ✅ | 40,000000 · 21,500000 · 5,000000 · 1,000000 |
| 1.5 | Westgard 12s/13s/22s/R4s/41s/10x | ✅ | 9/9 asserções, incluindo níveis opostos → R4s e **não** 22s |
| 1.6 | Elegibilidade CLSI EP05/C24 | ✅ | 6/6 estados; desconhecido nunca entra no cálculo |
| 1.7 | Robustez (n=0, n=1, DP=0, alvo=0) | ✅ | Todos devolvem 0, sem erro |
| 1.8 | Sem vazamento de estado entre grupos | ✅ | `grupos_com_0_violacoes_reportando_regra = 0` |

## 2. Integridade dos dados

| # | Item | Status | O que fecha |
|---|---|---|---|
| 2.1 | Fórmulas preservadas | ⏳ | Motor escreve em `Eng_Saida`; Painel/Estatística referenciam por fórmula. **Hoje destrói 1.365** (Painel 58→13, Estatística 2160→840). Teste: contagem antes/depois idêntica |
| 2.2 | Sem sobrescrita silenciosa | ⏳ | Reenviar mesma Data+lote sobrescreve valor sem versionar. Exige versionamento: linha original nunca alterada |
| 2.3 | Exclusão lógica não é revertida sozinha | ⏳ | `UpsertResultados` força `Status = Ativo`, ressuscitando excluído. Reverter exclusão deve ser ação própria, com justificativa |
| 2.4 | RUN representa a corrida analítica | ⏳ | Hoje RUN = (Data + lote) colapsa turnos e pós-calibração. CLSI C24 trata como eventos distintos |
| 2.5 | `Cfg_Status` não altera histórico retroativamente | ⏳ | Uma célula redefine o que entra na estatística de todo o histórico. Exige versionamento da tabela + log |

## 3. Auditoria e segurança

| # | Item | Status | O que fecha |
|---|---|---|---|
| 3.1 | Trilha de auditoria | ⏳ | `RegistrarLog` é stub vazio. Aba `Audit_Log` append-only: data/hora, usuário, papel, ação, chave, valor ANTES, valor DEPOIS, justificativa |
| 3.2 | Log em todo caminho de gravação | ⏳ | Chamar de **dentro do `mDados`**, não dos formulários — senão algum caminho escapa |
| 3.3 | Proteção persistida no arquivo salvo | ⏳ | Nenhuma das 18 abas tem `<sheetProtection>`. `UserInterfaceOnly:=True` **não persiste** entre sessões |
| 3.4 | Projeto VBA com senha | ⏳ | DPB decodifica para payload vazio — sem senha |
| 3.5 | Arquivo distribuído em estado bloqueado | ⏳ | Salvo em sessão autenticada: abrir com macros desabilitadas dá acesso total a `DB_Resultados` |

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

**8 ✅ · 24 ⏳ · 2 ❌**

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
