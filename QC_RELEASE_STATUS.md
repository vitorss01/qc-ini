# QC_RELEASE_STATUS — QC_INI

Documento de acompanhamento do fechamento de versão. **Ler no início de toda sessão; atualizar no fim.** Mantido pelo
`qc-release-manager` (`.claude/agents/qc-release-manager.md`, procedimento `/qc-release`).

## Release: entrega 48 — EM PRODUÇÃO (atualizado em 25/11/2026)

**Fase:** a ETAPA 7 está concluída. O relatório é
`_arquitetura/agentes/relatorios/2026-11-25_RELEASE_entrega48.md`. Veredito: **QC_INI — RELEASE CANDIDATE APROVADO
PARA ENTREGA** (release técnica do código).

A entrega 48 é o ADR-102: alvos do laboratório por lote e período; transferência Média/DP; exceção temporária;
vigilância bula/Unity no Painel, da linha 97 em diante. É uma mudança de escopo autorizada pelo dono em 25/11. Ela está
em produção desde 25/11, no commit `58d2847`. A implementação está no `8e38d99`, e a correção do gate no `58d2847`.

**Os dois commits ainda não foram enviados ao remoto.** O push vem com o diff conferido.

**Base em produção (sha da troca, antes do 1º ATUALIZAR):**

| Produto | sha256 | Bytes | Backup anterior |
|---|---|---|---|
| Bioquímica | `0eca3fac6734fc7a…` | 36.525.733 | `_backup_pre_integracao_2026-11-25_2009` |
| Hematologia | `846b2bf4b0b3b06c…` | 41.073.728 | `_backup_pre_integracao_2026-11-25_2222` |

O sha completo está no MANIFESTO.

### Critérios do gate (entrega 48)

| Critério | Estado | Evidência |
|---|---|---|
| Validação integrada | ✅ PASS | `entregar.py --direcionada --com qa_desempenho,qa_usuarios_adr084,qa_adr095`: 18/18 suítes por produto, 36 OK, 0 FAIL. Logs em `_entrega_2026-11-25_1806/` (não versionado) |
| O gate barra a regressão | ✅ | a 1ª tentativa (8e38d99) foi barrada pelo qa_adr090 V02 nos dois produtos, e a produção não foi trocada. A causa foi medida e corrigida só na posição dos botões (ADR-102, "Correção do gate"). Em cópia: qa_adr090 Bio 21/0, Hema 23/0 |
| Funções essenciais | ✅ | qa_final 29/0, qa_lotes_auto 14/0, qa_westgard_sequencia 29/0, qa_troca_lote 3/0, qa_blocos_lote 14/0, nos dois produtos |
| Cálculos = oráculo | ✅ | qa_alvos_manual 35/0 nos dois (oráculo `calcular_alvos` em Python); Tester X01/X05a = oráculo; X05b (LJ) e X05c (Westgard 1-3s) usam o alvo gravado. O motor e as regras estatísticas não mudaram |
| Recebimento/ATUALIZAR/integração | ✅ (em cópia) | qa_sync OK (S4, S5, S6), qa_etl Bio 64/0 e Hema 63/0, qa_migracao_adr070 2/0. **O ATUALIZAR na produção depois da 48 ainda não foi conferido** (a tarefa agendada roda sozinha) |
| Segurança/blindagem | ✅ | qa_seguranca 20/0; qa_usuarios_adr084 12/0; blindagem TRAVADO (42 abas, só a Login visível) |
| Validação independente | ✅ VALIDADA | Tester: fase 1 com ressalvas baixas (D-01 a D-03 → BK-23); fase 2 `repro_e48_alvos.py` Bio 14/14 e Hema 14/14, com o sha da produção inalterado (Y00) |
| 0 P0 e 0 P1 | ✅ | relatório 48, §5 |
| Proteção/recuperação | ✅ | backups acima; sha no MANIFESTO |
| Documentação | ✅ | ADR-102 (adendo, revisão e correção do gate), GUIA_DE_USO 4c, MANIFESTO, BACKLOG, relatório 48 |

**Limitação de dados:** na Hema, a faixa IDEAL (≥ 90) não foi exercitada com dado real. O maior N1 do lote de teste é
83, por isso a cópia usou `ALVO_N_IDEAL = 78`. Na produção, o valor instalado é 90 (BK-24).

### Pendências

- **P0:** nenhuma. **P1:** nenhuma.
- **Verificação pendente (não é defeito conhecido):** a verificação visual da seção do Painel, da linha 97 em diante, e
  dos botões. A política proíbe captura de tela; o qa_adr090 V02 prova a posição.
- ~~1º ATUALIZAR depois da 48~~ feito em 26/11 nos dois produtos (Bio +191, Hema +124; sem erro novo de QA: Bio 4 E01
  já conhecidos, Hema 0), junto com a **ação pontual** pedida pelo usuário: inativar só em nov/2026 os resultados com
  1-3s que foram repetidos no mesmo dia dentro de ±3DP (Bio 13, Hema 31; MANIFESTO 26/11). Não é regra nova do sistema.
  Escopo escolhido pelo usuário: nov/2026. A simulação do histórico inteiro (out/2025 em diante) deu 783 na Bio e 275 na Hema,
  e nada disso foi aplicado.
- **Para o RT (dado/alvo, não defeito de código):** em nov/2026 há violações persistentes **sem repetição**, compatíveis
  com alvo que não corresponde ao lote: Cálcio N1 z≈−15 diário; LDH e Magnésio N1 z≈+30 em 06/11; Cloro e Lipase
  rondando −3 a −4; na Hema RDW-SD N2 +4 a +15 quase diário. Conferir os alvos desses analitos no Painel (ADR-102).
- **P2:**
  - OC-0067/BK-05: a senha das abas está no repositório público, inclusive em `testes/qa_final.py:64`. A aceitação do
    risco ainda não foi registrada.
  - 4 E01 de governança na Bio: inativações sem MOTIVO.
  - OC-0011, OC-0039, logins do cadastro central, Ano EP N4.
  - Trava `~$` da Bio (ambiente).
- **P3 (backlog):**
  - entrega 48: BK-23 (D-01 a D-03), BK-24 (L11 a L15), BK-25 (spinner);
  - anteriores: BK-20, BK-21; BK-19 (Alta) e BK-22 (Média), sem efeito enquanto a gravação em massa do ADR-100
    estiver aposentada (ADR-102); BK-10 superado pelo ADR-102 §1.
- **Fora da release técnica:** a aprovação do RT e da qualidade, uma etapa humana documentada.

### Próxima ação objetiva

1. **(Usuário)** Conferir visualmente a seção do Painel da linha 97 em diante, nos dois produtos. Pontos a ver:
   - os botões não cobrem texto ("Aplicar Alvos ao Lote", "Liberar digitação manual", "Gravar valores manuais",
     "Registrar parecer");
   - a exceção aparece TRANCADA;
   - a grade bula/Unity está legível.
2. ~~Push~~ feito em 25/11 (`be185df`, com `8e38d99` e `58d2847`), diff conferido: sem senha, caminho de rede, `.xlsm` ou dados.
3. ~~1º ATUALIZAR depois da 48~~ feito em 26/11 (ver Pendências). **RT:** conferir os alvos de Cálcio N1, LDH, Magnésio N1, Cloro, Lipase (Bio) e RDW-SD N2 (Hema).
4. **ADR-100 em massa aposentado:** não aplicar o `aplicar_alvos_ciq.py` sem `--simular` em produção. O alvo passa a ser
   definido pelo analista no Painel (ADR-102). A pendência "RT revisar o CSV" da entrega 47 fica sem efeito, salvo uma
   decisão nova do usuário.
5. Depois: BK-05 (senha das abas) como a próxima entrega própria.

## Histórico: entrega 47 (RC aprovado em 25/11/2026, substituída pela 48)

- **Relatório:** `_arquitetura/agentes/relatorios/2026-11-25_RELEASE_entrega47.md`.
- **Código:** commit `c362922`.
- **Gate:** 21/21 suítes por produto, blindagem TRAVADO.
- **Tester 47:** VALIDADA (A-01, instalar_adr101, ferramenta ADR-100 em cópia).
- **Produção 47 (depois do ATUALIZAR de 24/11):** Bio `27526469…`, Hema `53fe8cc8…`.
- **Conteúdo:**
  - ADR-099 (Material e avisos do CEQ);
  - ADR-101 (Sigma em 6 faixas, piso 0);
  - P1 da Liberação (a assinatura acompanha o RUN) corrigido na 46;
  - A-01 (troca de lote sem login não apaga as assinaturas) corrigido na 47.
- **Equivalência VBA × PQ (gate 47, `--fixtures`):** Bio 37.009 e Hema 49.266 linhas, 0 divergências.
- **O-1 (gravação):** `ExecuteMso("FileSave")` grava; `wb.Save()` pela COM com eventos ligados não grava. Fica P3
  (BK-21).
- **Limitações registradas na 47, ainda válidas:**
  - F04 do qa_fluxo_manual NAO_EXECUTADO, por falta do caso nos dados (BK-20);
  - na Bio, o ADR-101 só exercita 4 das 6 faixas.

## Testes que não fecharam em casa (ambiente, não defeito do produto)

- O Excel 16.0.20625 de casa cai ou trava (0xc0000005, "Membro não encontrado", timeout). Nesses casos, rodar a suíte
  sozinha de novo.
- Dados antigos de cópias de 05/10 afetam o qa_adr088 R12c, o qa_blocos_lote BN/B00 e o qa_corrida_lote (BK-15).
