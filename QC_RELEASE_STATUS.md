# QC_RELEASE_STATUS — QC_INI

Documento de acompanhamento do fechamento de versão. **Ler no início de toda sessão; atualizar no fim.** Mantido pelo
`qc-release-manager` (`.claude/agents/qc-release-manager.md`, procedimento `/qc-release`).

## Release em curso: entrega 46 (candidata) — atualizado em 23/11/2026, 20h50

**Fase:** ETAPA 5 (testes direcionados, em cópia montada em casa) → próxima: ETAPA 6 (validação integrada no PC do
trabalho). Feature freeze desde 23/11/2026.

**Base em produção (validada):** entrega 45 (Bio sha `a9ba3515…`, Hema `1d0d9fe8…`, 18/11/2026).

**Conteúdo da candidata (fechado — nada mais entra além de P0/P1):**

| Item | Tipo | Estado |
|---|---|---|
| ADR-099 coluna Material + avisos do CEQ | pedido do usuário (antes do freeze) | testado em cópia; Hema qa_adr099 8/0; Bio M04 ajustado no teste (base sem ATUALIZAR) — reteste na v3 |
| ADR-100 alvos pelos 40 primeiros (ferramenta de dados) | pedido do usuário | ferramenta: todas as conferências A00–A10 PASS nos dois (cópia); **aplicar no trabalho após revisão do CSV pelo RT** |
| ADR-101 cores/classe do Sigma (6 faixas), Sigma negativo = 0, legenda | pedido do usuário (23/11, antes do freeze) | instalado nos dois (fumaça OK); qa_incerteza Bio 15/0 com o piso 0 |
| **Liberação: assinatura acompanha o RUN** (mLotes.AtualizarListaLiberacao + LiberStore) | **P1 corrigido** (integridade: assinatura ao lado do RUN errado) | qa_fluxo_manual F05e PASS (Bio, v2); correção complementar (assinada nunca sai da lista de 200) em reteste na v3 |

## Critérios do gate

| Critério | Estado | Evidência |
|---|---|---|
| Inserção manual integra ao interfaceamento | ✅ em cópia | qa_fluxo_manual Bio 19/20 e Hema 18/20 (falhas: F05e = defeito da Liberação, corrigido; Hema F04 pré-condição de dados) |
| Bias / ET / Sigma do CEQ = oráculo | ✅ em cópia | qa_fluxo_ceq: G/T/AC/AD/ET/Sigma = oráculo (diferenças só de rótulo de classe/tolerância/coluna U do teste, corrigidas no teste) |
| Equivalência VBA × Power Query | ✅ (12/11, ADR-092) | equivalencia_cq --fixtures 0 divergências — repetir no gate do trabalho |
| Regressão das áreas tocadas | ⏳ v2/v3 em execução | ver "Testes" |
| 0 P0 / 0 P1 conhecidos | ⏳ | único P1 encontrado (Liberação) corrigido, aguardando reteste |
| Validação integrada (entregar --direcionada) | ❌ pendente | só no PC do trabalho (precisa da produção e do DB_SEAC) |
| Documentação | ✅ | ADR-099/100/101, GUIA_DE_USO, PRONTIDAO_LIBERACAO_2026-11-23, este arquivo, BACKLOG |

## Testes (cópia de casa = arquivos de 05/10 + cadeia completa do entregar, modo HISTÓRICO)

- **Passaram:**
  - qa_adr087, qa_adr090, qa_adr092 e qa_incerteza (Bio, v2 com ADR-101);
  - qa_adr099 (Hema);
  - qa_linhas_alvo (Bio, com alvos).
- **Falhas de ambiente ou de dados da cópia (não são do produto):**
  - Excel 16.0.20625 caindo (0xc0000005, "Membro não encontrado"): qa_final, qa_lotes_auto e qa_westgard_sequencia na
    Bio. Repetir.
  - Dados antigos da cópia: qa_adr088 R12c, qa_blocos_lote BN (6354 Creatinina) e qa_corrida_lote.
  - Sem DB_SEAC ou cadastro central em casa: qa_sync e qa_usuarios_adr084.
- **Em execução:** v2 Hema e v3 (código final, nos dois produtos): qa_adr099, qa_fluxo_manual, qa_fluxo_ceq, qa_final
  e qa_lotes_auto.

## Pendências

- **P0:** nenhuma conhecida.
- **P1 (corrigida, NÃO validada; conta como aberta para o gate):** Liberação, assinatura ao lado do RUN errado
  (`src/{bio,hema}/mLotes.bas`, ADR-101 §2).
  - Provado até agora: só o F05e da **Bio** (v2).
  - Falta provar: o F05e da **Hema** e a parte complementar nos dois produtos (assinada nunca sai da lista de 200,
    "(sem resultado ativo)", coluna E da LiberStore).
  - O defeito também existe na **entrega 45, em produção**. Por isso o plano B do PRONTIDAO §6 ("45 + ADR-100") leva
    esse P1 junto. Se a 46 não fechar, o plano B passa a ser a 45 + só a correção da Liberação, ou uma orientação
    provisória: assinar a Liberação só depois de entrarem todas as corridas do dia.
  - No trabalho, antes de trocar a produção: conferir se alguma assinatura já está ao lado do RUN errado (aba
    Liberação × Audit_Log).
- **Pré-condições operacionais (não são defeito):**
  - `ADR-101.md` e `instalar_adr101.py` não estão no git, e o `entregar.py` chama a etapa `adr101`. Precisam entrar no
    commit.
  - O `.gitignore` dos backups do ADR-100 já está no commit 070f43a (`git check-ignore` confirma). O item do PRONTIDAO
    §3 está resolvido.
- **Testes que faltam:**
  - ADR-101: `qa_incerteza` na Hema, com o oráculo max(0, …);
  - ADR-099: Bio M04;
  - repetir sozinhas, na Bio, as suítes que caíram por ambiente (`qa_final`, `qa_lotes_auto`, `qa_westgard_sequencia`).
- **P2 (laboratório ou orientação, não bloqueiam):**
  - OC-0011 (especificações 2026);
  - OC-0039 (lote em uso, Configuração C20);
  - logins do cadastro central;
  - Ano EP N4 = 2026 quando chegar o 1º CEQ de 2026;
  - ver `PRONTIDAO_LIBERACAO_2026-11-23.md` §2 e §4.

## Próxima ação objetiva

1. Terminar a v3 em casa, com prioridade para o `qa_fluxo_manual` nos dois produtos: F05e e a parte complementar da
   Liberação. Depois, o `qa_incerteza` da Hema e as suítes que caíram por ambiente, cada uma sozinha.
2. Commit e push, incluindo `ADR-101.md`, `instalar_adr101.py`, os dois `mLotes.bas`, os testes, este arquivo e o
   BACKLOG. Conferir o diff: sem senha, sem caminho do DB_SEAC, sem `.xlsm`, sem `~$*`.
3. **No trabalho:**
   1. `git pull`. Conferir as assinaturas da Liberação na produção (ver Pendências);
   2. `entregar.py Hematologia Bioquimica --direcionada --com qa_fluxo_manual.py,qa_fluxo_ceq.py,qa_adr099.py,qa_incerteza.py,qa_westgard_sequencia.py,qa_adr091.py,qa_adr094.py,qa_adr097.py,qa_blocos_lote.py`;
   3. `aplicar_atualizar.py`;
   4. `aplicar_alvos_ciq.py --simular`, revisão do CSV pelo RT e então aplicar;
   5. validação independente (Tester);
   6. relatório de release.
