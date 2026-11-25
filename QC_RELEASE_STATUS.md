# QC_RELEASE_STATUS — QC_INI

Documento de acompanhamento do fechamento de versão. **Ler no início de toda sessão; atualizar no fim.** Mantido pelo
`qc-release-manager` (`.claude/agents/qc-release-manager.md`, procedimento `/qc-release`).

## Release: entrega 47 — EM PRODUÇÃO, VALIDADA, RC APROVADO — atualizado em 25/11/2026

**Fase:** ETAPA 7 concluída (relatório de release:
`_arquitetura/agentes/relatorios/2026-11-25_RELEASE_entrega47.md`). Veredito: **QC_INI — RELEASE CANDIDATE APROVADO
PARA ENTREGA** (release técnica do código). A **entrega 47** (commit c362922) está em produção nos dois produtos:
- `entregar.py --direcionada`: 21/21 suítes PASS por produto, blindagem TRAVADO;
- ATUALIZAR sem mudança de status depois da entrega (MANIFESTO 24/11);
- validada pelo Tester (`_arquitetura/agentes/relatorios/2026-11-24_validacao_entrega47/TESTER.md`): A-01, instalar_adr101,
  ferramenta ADR-100 (em cópia) e regressão mínima VALIDADAS.

O O-1 (gravação) foi medido em 25/11 pelo Orquestrador: o comando da interface `ExecuteMso("FileSave")` grava nos dois
produtos, sem e com sessão de login; só `wb.Save()` por automação com eventos ligados não grava. Fica P3 (BK-21).

**Fica fora da release técnica (etapas humanas posteriores):**
- a aplicação do ADR-100, que aguarda a aprovação escrita do RT. A Bio só vem depois de limpar o N1 dos lotes 8974/8977,
  simular de novo e remover a trava `~$`;
- a aprovação institucional/regulatória do laboratório.

**Base em produção (validada):** entrega 47 + ATUALIZAR. Bio sha `27526469…` (entrega `42c69669…`), Hema `53fe8cc8…`
(entrega `898d5412…`), 24/11/2026.

**Conteúdo da release (entregas 46 + 47; fechado):**

| Item | Tipo | Estado em cópia |
|---|---|---|
| ADR-099 coluna Material + avisos do CEQ | pedido do usuário (antes do freeze) | qa_adr099 Bio 8/0 e Hema 8/0 |
| ADR-100 alvos pelos 40 primeiros (ferramenta de dados) | pedido do usuário | A00–A11 PASS nos dois. Com os alvos: qa_blocos_lote BO Bio 126 pares/0 divergência, Hema 15/0; qa_westgard_sequencia Bio e Hema 29/0; qa_linhas_alvo 10/0 nos dois. **Aplicar no trabalho após o RT revisar o CSV** |
| ADR-101 cores/classe do Sigma (6 faixas), Sigma negativo = 0, legenda (V5:X8) | pedido do usuário (23/11, antes do freeze) | fumaça do instalador OK nos dois; qa_incerteza Bio 15/0 e Hema 16/0 com o piso 0; qa_adr090 Bio 21/0 e Hema 23/0 (legenda fora dos gráficos) |
| **Liberação: a assinatura acompanha o RUN** (`src/{bio,hema}/mLotes.bas`) | **P1 corrigido** | qa_fluxo_manual 19/0 nos dois + F04 NAO_EXECUTADO (F05e e F07b). qa_liberacao_run Bio 4/0 e Hema 4/0 (24/11): assinatura de corrida sem resultado ativo vai para o fim com aviso e 1 trilha; com o lote acima de 200 corridas a linha é reservada; a 200ª continua no lugar; tudo volta ao lugar |
| **A-01** troca de lote sem login não apaga as assinaturas da Liberação; instalar_adr101 reinstalável (c362922, entrega 47) | P2 corrigido | Tester 47: repro_e46_liberacao 8/8 nos dois com e sem login (antes Hema 4/8); repro_e47_troca_lote P01–P06 PASS; reinstalação 0 diferenças |


## Critérios do gate (entrega 47)

| Critério | Estado | Evidência |
|---|---|---|
| Funções essenciais (CIQ automático e manual, CEQ, inativação, lotes, Painel, Liberação) | ✅ produção | log da entrega 47: qa_final 29/0, qa_lotes_auto 14/0, qa_fluxo_manual 19/0 (F04 NAO_EXECUTADO, BK-20), qa_fluxo_ceq 13/0 + 1 ALERTA (BK-06), qa_liberacao_run 4/0, qa_westgard_sequencia 29/0, nos dois |
| Cálculos (bias, ET, Sigma, incerteza) = oráculo | ✅ | qa_fluxo_ceq, qa_incerteza (Bio 15/0, Hema 16/0); oráculo do Tester repro_e46_sigma 5/0 |
| Recebimento/ATUALIZAR/integração | ✅ | qa_sync OK, qa_etl Bio 64/0 Hema 63/0, qa_equiv_qa_adr094 3/0; ATUALIZAR pós-47 sem mudança |
| Equivalência VBA × PQ (motor completo) | ✅ 25/11 | `equivalencia_cq --fixtures` em cópias da produção nova: Bio 37.009 e Hema 49.266 linhas, 0 divergências |
| 0 P0 e 0 P1 | ✅ | relatório de release §5 |
| Validação integrada | ✅ PASS | entrega 47, 21/21 nos dois (MANIFESTO 24/11) |
| Validação independente | ✅ VALIDADA | TESTER entregas 46 e 47 |
| Proteção/recuperação | ✅ | blindagem TRAVADO; backups `_backup_pre_integracao_2026-11-24_2102/2345` e `_backup_pre_atualizar_2026-11-24_2346/2347`; sha no MANIFESTO |
| Documentação | ✅ | ADR-099/100/101, GUIA_DE_USO, PRONTIDAO_LIBERACAO_2026-11-23, relatório de release, BACKLOG |

## Testes que não fecharam em casa (não são defeito do produto)

- **Ambiente:** o Excel 16.0.20625 desta máquina cai ou trava (0xc0000005, "Membro não encontrado", timeout).
  - qa_lotes_auto Hema travou duas vezes (Excel ocupado/timeout) e passou na 3ª tentativa, na mesma cópia (842 s);
  - várias suítes caíram uma vez e passaram ao rodar de novo.
- **Dados antigos da cópia (arquivos de 05/10):**
  - qa_adr088 R12c;
  - qa_blocos_lote BN (6354 Creatinina) e B00 (anos "yyyy" na biblioteca);
  - qa_corrida_lote.
- **Rodadas no trabalho:** qa_sync OK nos dois (entrega 47); qa_usuarios_adr084 12/0 nos dois (25/11, cópias). O cadastro central ainda não está ligado na produção (`CAMINHO_USUARIOS` vazio).
- O caminho "assinatura de corrida sem resultado ativo" foi testado em 24/11 (`qa_liberacao_run`). Ele mostrou e
  corrigiu uma brecha: com o lote acima de 200 corridas, a assinatura não tinha linha e sumia da tela. Agora a linha é
  reservada.


## Pendências (nenhuma P0/P1)

- **P0:** nenhuma. **P1:** nenhuma (o P1 da Liberação foi corrigido na 46; o A-01 foi corrigido na 47; os dois estão
  validados em produção).
- **Ambiente/operacional:** a trava `~$QC_Bioquimica.xlsm` de 17/11 sobrou (nenhum processo segura o arquivo). Removê-la
  com o Excel da Bio fechado antes do ADR-100 na Bio.
- **P2, laboratório ou orientação (não bloqueiam):**
  - 4 E01 de Lactato N2 13/10 (Bio, inativação sem MOTIVO);
  - OC-0011 (especificações 2026);
  - OC-0039 (lote em uso, C20);
  - logins do cadastro central;
  - Ano EP N4 = 2026 quando chegar o 1º CEQ de 2026.
- **P2 com prazo:** OC-0067/BK-05 (senha das abas no repositório público), em entrega própria em cerca de 2 semanas.
- **P3 (backlog):** O-1 → BK-21; ferramenta ADR-100 → BK-22; F04 → BK-20; triagem do `--simular` → BK-19.
- **Limitações documentadas:**
  - o F04 não foi executado (falta o caso nos dados);
  - ADR-101 na Bio: os dados só têm 4 das 6 faixas (a Hema tem as 6).

## Próxima ação objetiva

1. **(Humano: RT)** revisar o CSV do `aplicar_alvos_ciq.py Hematologia --simular` (critérios: PRONTIDAO §4, passo 6;
   ressalvas do parecer). Com a aprovação escrita, aplicar sem flag na Hema.
2. Bio: limpar o N1 dos lotes 8974/8977, remover a trava `~$` (Excel da Bio fechado), simular de novo e levar ao RT.
3. **(Humano: qualidade/RT)** aprovação institucional documentada da versão 47.
4. Depois da release: BK-05 (senha das abas) como a próxima entrega.
