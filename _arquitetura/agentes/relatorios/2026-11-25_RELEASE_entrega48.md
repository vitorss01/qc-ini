# Relatório de release — QC_INI, entrega 48 / ADR-102 (25/11/2026)

Emitido pelo `qc-release-manager` (procedimento `/qc-release`, seção 7). Fontes:
- `_arquitetura/integracao_dbseac/ADR-102.md`, incluindo a seção "Correção do gate da entrega 48";
- `QC_RELEASE_STATUS.md` e `BACKLOG_POS_RELEASE.md` (BK-23 a BK-25);
- `_arquitetura/agentes/relatorios/2026-11-25_validacao_entrega48/tester_validar.md` (fases 1 e 2);
- `MANIFESTO_PRODUCAO.md`, as duas linhas de 25/11;
- `git log` até 58d2847;
- os logs da entrega em `_entrega_2026-11-25_1806/{Bioquimica,Hematologia}/` (não versionado). As contagens do §3 foram
  conferidas nesses logs.

Este relatório não traz senhas, caminhos de rede nem dados de paciente.

**Mudança de escopo durante o freeze.** A mudança de escopo foi autorizada pelo dono do sistema em 25/11 e está
registrada no ADR-102. Ela inclui a troca do critério do alvo: sai a regra dos "40 primeiros" do ADR-100 e entra o período escolhido
pelo analista. O ciclo 48 foi aberto depois do RC 47.

## 1. O que foi corrigido / entregue

| Item | Tipo | Commit | Produtos |
|---|---|---|---|
| ADR-102 §1: alvos do laboratório (Média/DP) por lote e período no rodapé do Painel (linha 97 em diante). Usa só resultados ativos, de todos os equipamentos, com datas inclusivas e DP n−1. Regra do N: < 2 bloqueia, 2 a 19 é temporário (pede 2ª confirmação), 20 a 89 é aceitável, ≥ 90 é ideal. A gravação é estática, com auditoria por nível em LotesStore AA:AO e no Audit_Log | funcionalidade nova autorizada pelo usuário | 8e38d99 | Bio, Hema |
| Adendo: escolha do que transferir (1 = Média e DP, 2 = só a Média, 3 = só o DP) e exceção temporária manual (Média da bula, DP do lote anterior ou da bula). A exceção fica trancada até ser liberada e exige justificativa. A marca temporária é registrada por campo | idem | 8e38d99 | Bio, Hema |
| ADR-102 §2: referência bula/Unity (abas `RefBulaStore`, `RefBulaObs`, `RefBulaParecer`). Gera observações automáticas, só informativas, e tem o parecer técnico do analista. **Nunca grava na LotesStore** | idem | 8e38d99 | Bio, Hema |
| `aplicar_alvos_ciq.py` pula a origem `MANUAL`. A aplicação em massa do ADR-100 fica **aposentada**; a ferramenta continua para o `--simular` e como oráculo | ferramenta | 8e38d99 | — |
| **Gate da 1ª tentativa:** o botão "Registrar parecer" ficava sobre o texto "Último parecer" (A181). Correção só de posição: `instalar_adr102.alinhar_com_aba_ativa` e `mAlvosRef.AlinharBotoes`, chamado no fim do `mUI.PainelEncaixar` | defeito do gate, corrigido antes da troca | 58d2847 | Bio, Hema |

O cálculo estatístico validado não mudou. Média e DP usam `mEstatistica.CalcularMedia/CalcularDP`, que já existiam.
O motor de dados (`mEtlCQ` = `DB_CQ_FINAL.m`) e as regras de Westgard e LJ também não mudaram. LJ e Westgard leem o alvo
gravado, como antes (X05b e X05c).

## 2. Testes executados

- **Antes da entrega, em cópias:**
  - `qa_alvos_manual`: Bio 35/0 e Hema 35/0. Na Hema, o `ALVO_N_IDEAL` foi rebaixado **só na cópia**: o lote do teste
    tem N1 = 83 (< 90), e o A00 registra `ideal_ajustado_na_copia = 78`. Na produção, o valor instalado é 90.
  - Reinstalação: MUDANCAS 0.
  - `instalar_adr102 --autoteste`: 16/16.
  - `aplicar_alvos_ciq --autoteste`: 20/20.
- **Entrega 48, 1ª tentativa** (8e38d99, 25/11 13h59):
  - resultado: 17/18 suítes OK por produto;
  - o `qa_adr090` V02 falhou nos dois produtos;
  - a **produção não foi trocada**.
  - Causa medida: a posição foi gravada com o Painel inativo, e o Excel a recalculou ao ativar a aba (o botão descia
    26 pt).
  - Primeira hipótese (o PainelEncaixar sozinho): testada e descartada.
  - Depois da correção final, em cópia: `qa_adr090` Bio 21/0 e Hema 23/0.
- **Entrega 48, 2ª tentativa** (58d2847, início às 18h06):
  - comando: `entregar.py Bioquimica Hematologia --modo SEAC --direcionada --com
    qa_desempenho.py,qa_usuarios_adr084.py,qa_adr095.py`;
  - todos os instaladores OK, incluindo o adr102 (MUDANCAS 14, fumaça OK, exceção TRANCADA);
  - blindagem TRAVADO nos dois: 42 abas, só a Login visível, nenhuma sem proteção, estrutura travada.
- **Validação independente (Tester):**
  - fase 1, estática: VALIDADA COM RESSALVAS, sem achado crítico ou alto; D-01 a D-03 são baixos;
  - fase 2: `repro_e48_alvos.py` em cópias da produção nova, nunca salvas.

## 3. O que passou

**Gate da 2ª tentativa:** 18/18 suítes por produto, 36 OK no total, 0 FAIL.

| Suíte | Bio | Hema |
|---|---|---|
| qa_seguranca | 20/0 | 20/0 |
| qa_alvos_manual (ADR-102) | 35/0 | 35/0 |
| qa_lotes_limites | 8/0 | 8/0 |
| qa_blocos_lote | 14/0 | 14/0 |
| qa_linhas_alvo | 10/0 | 10/0 |
| qa_varredura_lotes | 4/0 | 4/0 |
| qa_troca_lote | 3/0 | 3/0 |
| qa_lotes_auto | 14/0 | 14/0 |
| qa_westgard_sequencia | 29/0 | 29/0 |
| qa_graficos | 8/0 | 8/0 |
| qa_adr090 | 21/0 | 23/0 |
| qa_desempenho | 5/0 | 5/0 |
| qa_final | 29/0 | 29/0 |
| qa_etl | 64/0 | 63/0 |
| qa_migracao_adr070 | 2/0 | 2/0 |
| qa_sync (DB_SEAC, S4 trava de encolhimento, S5 recuperação, S6 reserva PQ) | OK | OK |
| qa_usuarios_adr084 | 12/0 | 12/0 |
| qa_adr095 | 10/0 | 8/0 |

As 12 suítes que fazem a verificação (adr090, adr095, alvos_manual, blocos_lote, etl, final, linhas_alvo, lotes_limites,
migracao_adr070, usuarios_adr084, varredura_lotes, westgard_sequencia) conferiram o arquivo de entrada intacto, com o
SHA-256 igual ao da produção nova. As outras 6 (desempenho, graficos, lotes_auto, seguranca, sync, troca_lote) não
registram o sha.

**Tester, fase 2: Bio 14/14 APROVADO e Hema 14/14 APROVADO.**
- X00: pré-condições.
- X01: um resultado inativado no período não entra (= oráculo).
- X02: entram todos os equipamentos.
- X03: as datas são inclusivas.
- X04: os valores são estáticos; trocar filtros ou rodar o ATUALIZAR não muda o alvo.
- X05a: o Aplicar no modo 1 dá o mesmo valor do oráculo, e a condição segue a regra do N.
- X05b: o LJ usa o alvo gravado.
- X05c: o Westgard 1-3s usa o alvo gravado.
- X06: exceção só do DP, seguida do Aplicar no modo 3.
- X07: trocar o analito tranca a exceção de novo; o spinner sem evento recusa gravar.
- X08: um resultado igual ao LI ou ao LS fica dentro da faixa.
- X09: uma observação PENDENTE de resultado inativado vira RETIRADA no ATUALIZAR.
- X10: um resultado inativado abaixo do LI não gera observação.
- Y00: o sha da produção é o mesmo antes e depois.

**Veredito do Tester: VALIDADA.**

## 4. O que não foi feito e por quê

| Item | Motivo | Consequência |
|---|---|---|
| Verificação visual humana da seção nova do Painel (linha 97 em diante) e dos botões | a política de dados de paciente proíbe captura de tela | a posição está provada só pelo `qa_adr090` V02 (forma sobre texto). A leitura visual fica com o usuário (próxima ação) |
| Uso real pelo analista | etapa humana | — |
| ATUALIZAR na produção depois da entrega 48 | não foi rodado nesta sessão. A tarefa agendada do DB_SEAC roda sozinha. Em cópia, o gancho no fim do ATUALIZAR foi exercitado (qa_alvos_manual R06, qa_final, qa_sync, X04 e X09) | a primeira execução real fica registrada pela tarefa agendada. Conferir no Sync_Log |
| Lacunas L11 a L15 do Tester: DP = 0, bit TESTE nos modos 2 e 3, rebloqueio no Open/BeforeClose, IDEAL com N ≥ 90 na Hema, sessão com login | falta dado sintético ou dado real (na Hema, nenhum lote tem N1 ≥ 90). O rebloqueio no Open/BeforeClose só foi conferido no código | BK-24 |
| Suíte completa | o ciclo usa a direcionada. A completa é periódica | — |
| `equivalencia_cq --fixtures` | o motor não mudou nesta entrega. A última prova foi no gate 47: 0 divergências | o qa_etl e o qa_sync cobrem o recebimento |
| Aprovação do RT e da qualidade | etapa humana documentada | fora da release técnica |

## 5. Pendências que ficam: nenhuma é P0 ou P1

| ID | Componente | Evidência | Prior. | Conduta |
|---|---|---|---|---|
| D-01 a D-03 | `src/comum/mAlvosRef.bas`: rastreabilidade da transferência parcial | Tester, fase 1. O Audit_Log guarda o histórico completo; só o resumo em AA:AO e W fica impreciso | P3 | BK-23 |
| L11 a L15 | `testes/qa_alvos_manual.py` | Tester, fase 1 §2 | P3 | BK-24 |
| UX do spinner | Painel, linhas 100 a 102 | a tela fica defasada; o Aplicar recalcula pelo analito atual e nada é gravado errado (X07) | P3 | BK-25 |
| VISUAL-97 | seção do Painel e botões | sem captura de tela, por política | verificação pendente (não é defeito conhecido) | o usuário confere na produção |
| OC-0067 / BK-05 | senha das abas no repositório público; `testes/qa_final.py:64` tem a senha literal | conhecido, anterior a esta entrega | P2 com prazo | entrega própria. A aceitação do risco até lá ainda não foi registrada pelo usuário |
| E01 de governança | Bio, aba Inativar | o resumo do QA no qa_final e no qa_etl mostra 4 erros de governança (inativação sem MOTIVO), como na entrega 47 | P2 (laboratório) | o laboratório informa o motivo. Não altera resultado |
| Herdadas da 47 | OC-0011, OC-0039, logins do cadastro central, Ano EP N4, trava `~$` da Bio | `QC_RELEASE_STATUS.md` | P2 / ambiente | sem mudança |

## 6. Adiados / backlog

Acréscimos deste ciclo: BK-23 (D-01 a D-03), BK-24 (lacunas do `qa_alvos_manual`) e BK-25 (apuração defasada pelo
spinner). O ADR-102 aposenta a aplicação em massa do ADR-100. Com isso, o BK-19 (triagem do `--simular`) e o BK-22
(robustez da gravação do `aplicar_alvos_ciq`) perdem a urgência: a ferramenta não grava mais em produção, a não ser por
uma decisão nova. O BK-10 está superado (ADR-102 §1).

## 7. Status final

Conferência do gate:
- **0 P0 e 0 P1 conhecidos.**
- Validação integrada PASS: 18/18 suítes nos dois produtos, blindagem TRAVADO.
- O gate funcionou: a 1ª tentativa foi barrada e a produção não foi trocada.
- Validação independente VALIDADA: o Tester deu 14/14 nos dois produtos, em cópias da produção nova.
- Cálculo: N, Média e DP batem com o oráculo independente (`calcular_alvos`, em Python). LJ e Westgard usam o alvo
  gravado. O motor e as regras estatísticas não mudaram.
- Proteção e recuperação: os backups são `_backup_pre_integracao_2026-11-25_2009` (Bio) e `_2222` (Hema); o sha está no
  MANIFESTO.
- Documentação: ADR-102 com adendo, revisão e correção do gate; GUIA_DE_USO 4c; MANIFESTO; BACKLOG; este relatório.

**QC_INI — RELEASE CANDIDATE APROVADO PARA ENTREGA** (entrega 48)

Esta aprovação vale para a release técnica do código. Ficam de fora, como etapas humanas:
1. a verificação visual da seção do Painel pelo usuário;
2. a aprovação institucional e regulatória do RT e da qualidade.

Os commits 8e38d99 e 58d2847 ainda **não foram enviados ao remoto**. O push vem depois deste relatório, com o diff
conferido.

## 8. Versão recomendada

**Entrega 48**, código no commit `58d2847`, que está em produção desde 25/11/2026:

| Produto | sha256 em produção | Bytes | Backup anterior |
|---|---|---|---|
| QC_Bioquimica.xlsm | `0eca3fac6734fc7a9334e04d28c8cc6e3483f601da2211abd3e5ce76fcf224e0` | 36.525.733 | `_backup_pre_integracao_2026-11-25_2009` |
| QC_Hematologia.xlsm | `846b2bf4b0b3b06c5509187a86dfb3442f920c5c3deed0b517dae919443a3af0` | 41.073.728 | `_backup_pre_integracao_2026-11-25_2222` |

Para voltar à entrega 47, use esses backups. O sha muda no primeiro ATUALIZAR da tarefa agendada, o que é esperado. O
sha acima é o da troca.
