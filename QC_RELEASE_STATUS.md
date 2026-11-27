# QC_RELEASE_STATUS — QC_INI

Documento de acompanhamento do fechamento de versão. **Ler no início de toda sessão; atualizar no fim.** Mantido pelo
`qc-release-manager` (`.claude/agents/qc-release-manager.md`, procedimento `/qc-release`).

## Release: entrega 49 — HEMA EM PRODUÇÃO, BIO AGUARDANDO A TROCA (atualizado em 27/11/2026)

**Fase:** a ETAPA 7 está concluída. O relatório é
`_arquitetura/agentes/relatorios/2026-11-27_RELEASE_entrega49.md`. Veredito: **QC_INI — RELEASE CANDIDATE APROVADO
PARA ENTREGA** (release técnica do código), com a **Bioquímica aguardando a troca**.

A entrega 49 é o ADR-103, mudança de escopo autorizada pelo dono em 26/11 ("3. ok — pode!"):
1. regra institucional do Sigma, literal, 5 faixas (**σ = 4 é Marginal**, decisão do usuário);
2. "Aplicar a vários analitos" (`frmAlvosLote` + `AplicarAlvosVarios`, mesmo núcleo do botão de um);
3. Painel compacto (C4:D4 removido, H4:I4 mesclado, status por nível → coluna AB oculta, Classe Sigma em J);
4. tabela de digitação da bula em N114 + "TRANSFERIR VALORES PARA REFERÊNCIA INICIAL";
5. "INADEQUADO" (Bio Painel A57): **pendente de definição do usuário, nada alterado**.

BI fora (BK-26). Produto: `355781e`, `c8137a3`. Só testes: `eb13620`, `a5a263b`. MANIFESTO: `b97bf3b`.
**Os cinco commits ainda não foram enviados ao remoto.**

**Base em produção:**

| Produto | Estado | sha256 | Backup anterior |
|---|---|---|---|
| Hematologia | entrega 49, desde 27/11 05:18 | `b8c501144f1c8e0f…` (41.003.831 bytes) | `_backup_pre_integracao_2026-11-27_0518` |
| Bioquímica | **ainda entrega 48** (+ ATUALIZAR/13s de 26/11); validado `44f713c92b3542bf…` aguarda troca | `e9589911bebc0ff3…` | `_backup_pre_integracao_2026-11-27_0156` |

### Critérios do gate (entrega 49)

| Critério | Estado | Evidência |
|---|---|---|
| Validação integrada | ✅ PASS | 2ª tentativa (eb13620), `--direcionada` + 10 `--com`: 26/26 por produto. Bio `_entrega_2026-11-26_2333`; Hema `_entrega_2026-11-27_0157` (G03 do qa_graficos = ambiente, passou na retomada 8/0; qa_etl 63/0 no rerun isolado após a5a263b). Logs não versionados |
| O gate barra a regressão | ✅ | 1ª tentativa (c8137a3) barrada no qa_final T17 nos dois; produção não trocada. Causa: premissa do teste × estado da produção (repetições reais na Bio; Hema no lote 625311 com Ano 2025). Correção só do teste (eb13620); rerun 29/0 nos dois |
| Funções essenciais | ✅ | qa_final, qa_lotes_auto, qa_westgard_sequencia, qa_troca_lote, qa_blocos_lote, qa_fluxo_manual, qa_alvos_manual OK nos dois |
| Cálculos = oráculo | ✅ | Tester: Sigma = oráculo em 22 fronteiras; Estatística M e Painel J = oráculo em todas as linhas (Bio 58/57, Hema 31/28). Fórmula do Sigma, motor e Westgard/LJ inalterados |
| Recebimento/ATUALIZAR/integração | ✅ (em cópia) | qa_sync OK, qa_etl Bio 64/0 e Hema 63/0, qa_migracao_adr070 2/0 |
| Segurança/blindagem | ✅ | qa_seguranca 20/0; qa_adr084 17/0; qa_usuarios_adr084 12/0; blindagem TRAVADO nos dois |
| Validação independente | ✅ VALIDADA | Tester `repro_e49_adr103.py`: Bio 24/0, Hema 23/0, itens 1–5; Y00 sha inalterado |
| 0 P0 e 0 P1 | ✅ | relatório 49, §5 |
| Proteção/recuperação | ✅ | backups acima; Bio intacta (sha = backup); sha no MANIFESTO |
| Documentação | ✅ | ADR-103, MANIFESTO (Hema), BACKLOG, relatório 49 |

### Pendências

- **P0:** nenhuma. **P1:** nenhuma.
- **Troca da Bio (operacional, não defeito):** `shutil.copy2` → WinError 32; a produção Bio está aberta num Excel do
  laboratório desde a madrugada. Um script em segundo plano espera o fechamento (até 10 h) e roda `trocar_producao.py`,
  que recusa se a produção tiver sido gravada desde o início dos testes. **Se recusar, refazer o gate Bio.** Depois da
  troca, registrar a linha da Bio no MANIFESTO.
- **Versões mistas** até a troca: Bio classifica o Sigma pelas 6 faixas do ADR-101, Hema pela regra institucional (o
  número do Sigma é o mesmo).
- **Tarefa 5 "INADEQUADO":** aguarda a definição do usuário.
- **Verificação visual (não é defeito conhecido):** Painel compacto (topo, J6:J9, H4:I4) e tabela N114 com o botão, nos
  dois produtos. Política sem captura de tela; qa_adr103 e o Tester provam a estrutura.
- **Para o usuário/RT (configuração, não defeito):** a produção Hema está no lote 625311 com Estatística Ano De/Até
  2025; o lote não tem dados nesse ano, então a tela de Estatística da Hema aparece vazia até ajustar Ano De/Até para 2026.
- **Para o RT (herdado da 48):** conferir os alvos de Cálcio N1, LDH, Magnésio N1, Cloro, Lipase (Bio) e RDW-SD N2 (Hema).
- **P2 (herdadas, sem mudança):** OC-0067/BK-05 (senha das abas no repositório); 4 E01 de governança na Bio; OC-0011,
  OC-0039, logins do cadastro central, Ano EP N4; trava `~$` da Bio (ambiente).
- **P3 (backlog):** entrega 49: BK-26 (BI), BK-27 (cores Analitos), BK-28 (lint), BK-29 (largura N114 Hema), BK-30
  (qa_fluxo_ceq Hema), BK-31 (V07 do Tester). Anteriores: BK-23 a BK-25 e demais.
- **Fora da release técnica:** a aprovação do RT e da qualidade, uma etapa humana documentada.

### Próxima ação objetiva

1. **(Automático/Claude)** Conferir o resultado do script de troca da Bio: se trocou, registrar sha e backup no MANIFESTO;
   se recusou (produção gravada), refazer o gate Bio (`entregar.py --direcionada` com as mesmas `--com`).
2. **(Usuário)** Ajustar Ano De/Até da Estatística da Hema para 2026; conferir visualmente o Painel compacto e a N114;
   definir a Tarefa 5.
3. Push dos commits 355781e..b97bf3b (e da linha Bio do MANIFESTO) com o diff conferido: sem senha, caminho de rede,
   `.xlsm` ou dados.

## Histórico: entrega 48 (RC aprovado em 25/11/2026; Hema substituída pela 49 em 27/11, Bio até a troca)

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
