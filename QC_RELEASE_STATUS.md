# QC_RELEASE_STATUS — QC_INI

Documento de acompanhamento do fechamento de versão. **Ler no início de toda sessão; atualizar no fim.** Mantido pelo
`qc-release-manager` (`.claude/agents/qc-release-manager.md`, procedimento `/qc-release`).

## Release em curso: entrega 46 (candidata) — atualizado em 23/11/2026, 23h10

**Fase:** ETAPA 5 concluída em casa (testes direcionados em cópia). **Próxima: ETAPA 6** — validação integrada no PC do
trabalho (precisa da produção e do DB_SEAC). Feature freeze desde 23/11/2026.

**Base em produção (validada):** entrega 45 (Bio sha `a9ba3515…`, Hema `1d0d9fe8…`, 18/11/2026).

**Conteúdo da candidata (fechado — nada mais entra além de P0/P1):**

| Item | Tipo | Estado em cópia |
|---|---|---|
| ADR-099 coluna Material + avisos do CEQ | pedido do usuário (antes do freeze) | qa_adr099 Bio 8/0 e Hema 8/0 |
| ADR-100 alvos pelos 40 primeiros (ferramenta de dados) | pedido do usuário | A00–A11 PASS nos dois. Com os alvos: qa_blocos_lote BO Bio 126 pares/0 divergência, Hema 15/0; qa_westgard_sequencia Bio e Hema 29/0; qa_linhas_alvo 10/0 nos dois. **Aplicar no trabalho após o RT revisar o CSV** |
| ADR-101 cores/classe do Sigma (6 faixas), Sigma negativo = 0, legenda (V5:X8) | pedido do usuário (23/11, antes do freeze) | fumaça do instalador OK nos dois; qa_incerteza Bio 15/0 e Hema 16/0 com o piso 0; qa_adr090 Bio 21/0 e Hema 23/0 (legenda fora dos gráficos) |
| **Liberação: a assinatura acompanha o RUN** (`src/{bio,hema}/mLotes.bas`) | **P1 corrigido** | qa_fluxo_manual Bio 20/0 e Hema 19/0. F05e: o marcador acompanha o RUN. F07b: a Liberação volta ao estado inicial depois da retirada |

## Critérios do gate

| Critério | Estado | Evidência |
|---|---|---|
| Inserção manual integra ao interfaceamento | ✅ em cópia | qa_fluxo_manual Bio 20/0, Hema 19/0 + 1 NAO_EXECUTADO (F04: falta a pré-condição nos dados) |
| Bias, ET e Sigma do CEQ = oráculo | ✅ em cópia | qa_fluxo_ceq Bio e Hema 13/0 + 1 ALERTA (OC-0069, já no backlog BK-06) |
| Fluxo principal e ATUALIZAR | ✅ em cópia | qa_final Bio e Hema 29/0; qa_lotes_auto Bio 14/0 (v3) e Hema 14/0 (v2) |
| Equivalência VBA × Power Query | ✅ (12/11) | equivalencia_cq --fixtures: 0 divergências. O motor não mudou desde então; repetir no gate do trabalho |
| 0 P0 e 0 P1 conhecidos | ✅ em cópia | o único P1 (Liberação) foi corrigido e retestado nos dois produtos |
| Validação integrada (entregar --direcionada) | ❌ pendente | só no PC do trabalho |
| Documentação | ✅ | ADR-099/100/101, GUIA_DE_USO, PRONTIDAO_LIBERACAO_2026-11-23, este arquivo, BACKLOG |

## Testes que não fecharam em casa (não são defeito do produto)

- **Ambiente:** o Excel 16.0.20625 desta máquina cai ou trava (0xc0000005, "Membro não encontrado", timeout).
  - qa_lotes_auto Hema estourou os 1.500 s (na v2 levou 1.381 s e passou);
  - várias suítes caíram uma vez e passaram ao rodar de novo.
- **Dados antigos da cópia (arquivos de 05/10):**
  - qa_adr088 R12c;
  - qa_blocos_lote BN (6354 Creatinina) e B00 (anos "yyyy" na biblioteca);
  - qa_corrida_lote.
- **Exigem o trabalho:** qa_sync (DB_SEAC) e qa_usuarios_adr084 (cadastro central).
- **Não testado especificamente:** o caminho "assinatura de corrida que perdeu todo resultado ativo" (vai para o fim da
  lista com aviso) e a corrida assinada além da 200ª.

## Pendências

- **P0:** nenhuma.
- **P1:** nenhuma aberta. A da Liberação foi corrigida e retestada em cópia. O defeito também existe na entrega 45, em
  produção: **no trabalho, antes de trocar, conferir se alguma assinatura já está ao lado do RUN errado** (aba
  Liberação × Audit_Log `ASSINATURA_*`, que guarda o RUN).
- **P2 (laboratório ou orientação, não bloqueiam):**
  - OC-0011 (especificações 2026);
  - OC-0039 (lote em uso, Configuração C20);
  - logins do cadastro central;
  - Ano EP N4 = 2026 quando chegar o 1º CEQ de 2026;
  - ver PRONTIDAO §2 e §4.

## Próxima ação objetiva (no trabalho)

1. `git pull`.
2. Conferir as assinaturas da Liberação na produção.
3. Fechar os outros Excel e pausar a tarefa DB_SEAC.
4. Rodar:
   ```
   python _arquitetura/integracao_dbseac/entregar.py Hematologia Bioquimica --direcionada --com qa_westgard_sequencia.py,qa_adr091.py,qa_adr094.py,qa_blocos_lote.py,qa_lotes_auto.py
   ```
   As suítes novas já estão no gate: qa_adr099, qa_fluxo_manual, qa_fluxo_ceq e qa_incerteza (SUITES_DO_LOTE).
5. `aplicar_atualizar.py` nos dois produtos.
6. `aplicar_alvos_ciq.py <produto> --simular`, revisão do CSV pelo RT (critérios no PRONTIDAO §4, passo 6), depois
   aplicar sem flag.
7. Validação independente (tester-qa) do que mudou.
8. Relatório de release e declaração do RC, se o gate passar.
