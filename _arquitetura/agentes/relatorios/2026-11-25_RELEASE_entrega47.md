# Relatório de release — QC_INI, entrega 47 (25/11/2026)

Emitido pelo `qc-release-manager` (procedimento `/qc-release`, seção 7). As fontes são estas:
- `QC_RELEASE_STATUS.md` e `BACKLOG_POS_RELEASE.md`;
- `MANIFESTO_PRODUCAO.md`, linhas de 24/11;
- os TESTER.md das entregas 46 e 47;
- `git log` até ca91be6;
- o log da entrega 47 (`%TEMP%\qcw\entrega47.log`, fora do repositório);
- as medições do O-1 (`%TEMP%\qcw\o1\{Bioquimica,Hematologia}\resultado.json`).

Este relatório não inclui senhas, caminhos de rede nem dados de paciente.

## 1. O que foi corrigido

O ciclo 46/47 é o feature freeze desde 23/11. As decisões estão nos ADR-099, 100, 101 e 046.

| Item | Tipo | Commit | Produtos |
|---|---|---|---|
| **Liberação: a assinatura acompanha o RUN** (`src/{bio,hema}/mLotes.bas`), com linha reservada para a assinatura de uma corrida sem resultado ativo | P1 corrigido (entrega 46) | e2dfaff / 9c4021a | Bio, Hema |
| **A-01**: a troca de lote sem login apagava as assinaturas da Liberação (Hema). A correção libera e restaura a proteção em `SalvarViewNoBloco`/`CarregarBlocoNaView`, como manda o ADR-046 | P2 latente, anterior à 46, corrigido (entrega 47) | c362922 | Bio, Hema |
| `instalar_adr101.py` reinstalável: ele recusava reinstalar por causa da legenda e travaria a entrega | correção de ferramenta (entrega 47) | c362922 | — |
| ADR-099 (coluna Material e avisos do CEQ), ADR-101 (Sigma em 6 faixas, Sigma negativo = 0, legenda V5:X8) | pedidos do usuário anteriores ao freeze (entrega 46) | 58cabc3…9c4021a | Bio, Hema |
| ADR-100: ferramenta `aplicar_alvos_ciq.py` (alvos pelos 40 primeiros) | ferramenta de dados. **Ainda não foi aplicada em produção** | — | — |

O cálculo estatístico não mudou. O motor de dados (`mEtlCQ` = `DB_CQ_FINAL.m`) não foi tocado nas entregas 46 e 47.

## 2. Testes executados

- **Validação integrada (gate):** entrega 47, `entregar.py --direcionada`, com 21 suítes por produto mais 32 verificações
  do instalador e a blindagem.
  - Bio: ENTREGUE, sha `42c69669…`, backup `_backup_pre_integracao_2026-11-24_2102`.
  - Hema: ENTREGUE, sha `898d5412…`, backup `_backup_pre_integracao_2026-11-24_2345`.
  - Depois veio o ATUALIZAR sem mudança de status: Bio `27526469…` (ATIVO 36.826) e Hema `53fe8cc8…` (ATIVO 48.546).
- **Validação independente (Tester):**
  - entrega 46: 4 itens;
  - entrega 47: A-01, instalar_adr101, ferramenta ADR-100 e regressão mínima.
  - Tudo rodou em cópias, e o sha da produção ficou igual no início e no fim.
- **Medições do Orquestrador:**
  - A-01 antes/depois (24/11, 18h33–18h42, antes do commit c362922): em cópias da produção da **entrega 46**
    (`9693db90…`/`a8d6fcb2…`) com o código corrigido instalado. Na produção nova, quem mediu foi o Tester 47.
  - O-1 (gravação), 25/11, em cópias da produção nova, nos dois produtos, **sem e com sessão de login**
    (`%TEMP%\qcw\o1\` e `%TEMP%\qcw\o1login\`).
  - Trava `~$QC_Bioquimica.xlsm`, 25/11: `QC_Bioquimica.xlsm` abriu com `FileShare None` (medição feita no terminal,
    sem arquivo de log).

## 3. O que passou

| Área | Evidência |
|---|---|
| Segurança, blindagem | qa_seguranca 20/0 nos dois. XML: 39 abas, só a Login visível, nenhuma aba sem proteção, estrutura travada |
| CIQ automático, ATUALIZAR, fluxo principal | qa_final 29/0 nos dois; qa_lotes_auto 14/0 nos dois; qa_sync OK nos dois (DB_SEAC, no trabalho) |
| ETL e equivalência do QA | qa_etl Bio 64/0, Hema 63/0; qa_equiv_qa_adr094 3/0 nos dois (VBA = PQ); qa_migracao_adr070 2/0 |
| CIQ manual | qa_fluxo_manual 19/0 nos dois, com o F04 NAO_EXECUTADO (ver §4) |
| CEQ: bias, ET, Sigma pela janela | qa_fluxo_ceq 13/0 + 1 ALERTA nos dois. O ALERTA é o OC-0069, já no backlog como BK-06 |
| Incerteza e Sigma com piso 0 | qa_incerteza Bio 15/0, Hema 16/0; oráculo próprio do Tester `repro_e46_sigma` 5/0 nos dois |
| Westgard por sequência | qa_westgard_sequencia 29/0 nos dois |
| Lotes, blocos, material | qa_blocos_lote 14/0, qa_adr091 7/0, qa_adr094 7/0, qa_adr099 8/0 nos dois |
| Painel e Estatística | qa_adr087 14/0, qa_adr090 Bio 21/0 e Hema 23/0, qa_adr092 7/0, qa_adr095 Bio 10/0 e Hema 8/0, qa_adr097 Bio 29/0 e Hema 30/0 |
| Liberação pelo RUN | qa_liberacao_run 4/0 nos dois, na entrega e no Tester 47. `repro_e46_liberacao` 8/8 nos dois, com e sem login |
| A-01 | antes: Hema 4/8, Bio 8/8. Depois: 8/8 nos dois. Cenário próprio do Tester `repro_e47_troca_lote`: P01 a P06 PASS nos dois (trocas seguidas, Save, reabertura, proteção ativa). A 1ª rodada deu 4/2 (P05/P06) porque o `wb.Save()` pela COM não gravou (O-1); a 2ª chamou a sequência do BeforeSave passo a passo, sem o `Me.Save` aninhado nem o vigia dos gráficos — prova que o A-01 persiste no disco, não o caminho de gravação do usuário (este está na linha O-1) |
| instalar_adr101 reinstalável | duas instalações na mesma cópia: rc 0 e 0 diferenças na foto, nos dois produtos |
| Ferramenta ADR-100 (em cópia) | autoteste 19/19; recusa com trava `~$`; `--simular` não grava nada; `--so-copia` 14 PASS com 0 divergência; troca em cópia com backup e sha conferidos; recusa reaplicar sem `--forcar` |
| Integridade (T00) | status documentados, inativações da aba = INATIVADO, nenhum não-ATIVO na estatística, sha da produção inalterado pelos testes |
| O-1, gravação pelo caminho da interface | comando da interface `ExecuteMso("FileSave")` com eventos ligados **gravou no disco nos dois produtos, sem sessão e com sessão de login** (usuário de teste só na cópia). `wb.Save()` pela COM com eventos ligados não grava: sem sessão fica `Saved=False`; com sessão fica `Saved=True` (enganoso, só em automação). Com eventos desligados, como fazem as ferramentas (`aplicar_*` desligam os eventos; os instaladores conferem a data de gravação), grava nos dois produtos com sessão. Limite: foi o comando da interface acionado por automação, não um Ctrl+S digitado por pessoa (sem captura de tela, por regra) |
| Equivalência VBA × PQ, motor completo (`equivalencia_cq --fixtures`) | 25/11, em cópias da produção nova: Bio 37.009 = 37.009 linhas, Hema 49.266 = 49.266, **0 divergências** nos dois (`%TEMP%\qcw\gate47\`) |
| Usuários e assinatura (ADR-084, `qa_usuarios_adr084`) | 25/11, cópia do QC + cópia do cadastro central: 12/0 nos dois. O central ainda não está ligado na produção (`CAMINHO_USUARIOS` vazio): é a pendência LOGINS |

## 4. O que não foi feito e por quê

| Item | Motivo | Consequência |
|---|---|---|
| F04 do qa_fluxo_manual (R_4s com manual) | NAO_EXECUTADO nos dois: os dados reais não têm nenhuma corrida C2 com o N1 em 1,2 ≤ \|z\| ≤ 2,0 | não se declara PASS. A regra R_4s em si é coberta pelo qa_westgard_sequencia (29/0). O caso sintético fica no BK-20 |
| ADR-101 com as 6 faixas na Bio | os dados da Bio só têm 4 das 6 faixas | a Hema exercita as 6. O plano `tblPlanoQC_Sigma` tem as 6 faixas nos dois produtos (G02) |
| Aplicação do ADR-100 em produção | é uma ação de dados que exige aprovação **escrita** do RT. O parecer do Especialista em CQ (fora do repositório) diz: Hema aplicar com ressalvas; Bio **não** aplicar antes de limpar o N1 dos lotes 8974/8977 e simular de novo | etapa humana posterior, fora desta release de código |
| Suíte completa (40 suítes) | o ciclo usa a direcionada (21). A completa é periódica | — |
| Teste de interface com captura de tela | proibido (dados de paciente). O O-1 foi resolvido por `ExecuteMso`, sem capturar a tela | — |

## 5. Pendências que ficam: nenhuma é P0 ou P1

| ID | Componente | Evidência | Prior. | Conduta |
|---|---|---|---|---|
| O-1 | `EstaPastaDeTrabalho.cls` / qa_seguranca S19 | `o1\*` e `o1login\*\resultado.json`: o comando da interface grava (sem e com login); COM+eventos não grava. O Tester 47 o deixou como "P2 a confirmar"; a confirmação com login foi feita depois, pelo Orquestrador (não houve nova validação do Tester) | **P3** | BK-21 (o S19b passa a provar a gravação) |
| BK-06 / OC-0069 | aviso de provedor do CEQ | ALERTA do qa_fluxo_ceq nos dois | **P3** | backlog |
| Ano EP N4 | Configuração | trocar para 2026 quando chegar o 1º CEQ de 2026 | **P2 (orientação)** | laboratório |
| TRAVA-BIO | arquivo `~$QC_Bioquimica.xlsm` de 17/11, que sobrou | medido: QC_Bioquimica.xlsm abre com FileShare None, então nenhum processo o segura | **ambiente/operacional** | remover, com o Excel da Bio fechado, antes de aplicar o ADR-100 na Bio. Sem isso a ferramenta recusa (Y00), o que é seguro |
| E01-LAC | Bio, aba Inativar | 4 E01 do QA_INTEGRACAO: Lactato N2 de 13/10 inativado sem MOTIVO. Já existiam antes da 46 (MANIFESTO 24/11) | **P2 (laboratório)** | o laboratório informa o motivo. Não altera resultado: as linhas estão INATIVADO |
| OC-0011 | especificações 2026 | OCORRENCIAS | **P2 (laboratório)** | decisão ou cadastro do laboratório |
| OC-0039 | lote em uso, Configuração C20 | OCORRENCIAS | **P2 (orientação)** | orientação |
| LOGINS | cadastro central (ADR-084) | logins dos usuários ainda não criados | **P2 (laboratório)** | o ADMIN cria os logins |
| OC-0067 / BK-05 | senha das abas em texto puro no repositório público | conhecida. A barreira real é o login PBKDF2 mais a permissão da pasta | **P2** (com prazo: Alta, cerca de 2 semanas) | entrega própria pós-release (BK-05). **Aceitar o risco até lá é decisão do usuário, ainda não registrada** |
| BK-22 | `aplicar_alvos_ciq.py` | Tester 47: o sha do backup não é conferido; `makedirs` sem `exist_ok` (falha antes da troca). O `--so-copia` e a troca em cópia rodaram **só na Hema**; na Bio só o `--simular` (gravação real com pares "insuficiente" e 2 níveis só no autoteste sem Excel; `--forcar` não testado) | **P3** | backlog; **antes de aplicar na Bio de produção, rodar `--so-copia` na Bio** |
| ADR-100 | dados (LotesStore) | parecer do Especialista em CQ | **não é defeito**: é uma etapa humana | aprovação escrita do RT, depois `--simular` e a aplicação (ver o STATUS) |

## 6. Adiados / backlog

BK-01 a BK-22 estão em `BACKLOG_POS_RELEASE.md`. Os acréscimos deste ciclo são o BK-19 (triagem do `--simular`), o BK-20
(F04 sintético), o BK-21 (O-1/S19b) e o BK-22 (robustez da ferramenta do ADR-100). Nada foi adiado de novo neste
relatório.

## 7. Status final

Conferência do gate:
- **0 P0 e 0 P1 conhecidos.**
- Validação integrada PASS: 21/21 suítes nos dois produtos, blindagem TRAVADO.
- Validação independente: o Tester marcou VALIDADA todos os itens das entregas 46 e 47 (o item 3 da 46 "com o achado
  A-01", corrigido na 47 e validado de novo).
- Proteção e recuperação preservadas: backups, sha no MANIFESTO, blindagem.
- Documentação atualizada: ADRs, GUIA, PRONTIDAO, STATUS, BACKLOG, MANIFESTO.
- As limitações estão documentadas no §4.

**QC_INI — RELEASE CANDIDATE APROVADO PARA ENTREGA**

A entrega 47 entrou em produção em 24/11 (o gate do `entregar.py` só troca a produção com todas as suítes OK); esta
declaração confirma, depois da validação independente, a versão que já está em produção. Ela vale para a release
técnica do código. Ficam explícitas três etapas **humanas e posteriores**, que esta
aprovação não cobre:
1. **ADR-100:** aplicar os alvos pelos 40 primeiros só com a aprovação escrita do RT. Na Bio, isso vem só depois de
   limpar o N1 dos lotes 8974/8977, simular de novo e remover a trava `~$`.
2. **Aprovação institucional e regulatória:** o responsável técnico e a qualidade do laboratório aprovam e documentam.
   "Release técnico aprovado" não é essa aprovação.
3. As pendências P2 do laboratório: E01 do Lactato, OC-0011, OC-0039 e os logins.

## 8. Versão recomendada

**Entrega 47**, código no commit `c362922`, que é o que está em produção:

| Produto | sha256 em produção (após o ATUALIZAR de 24/11) | Entrega (código) |
|---|---|---|
| QC_Bioquimica.xlsm | `2752646953b2209939b86dc9ae86a3b098bc5e09ec1ef1ddb08f0dee2bbf7ffc` | `42c696693baf4ddd…` |
| QC_Hematologia.xlsm | `53fe8cc8ea48efad8d4dfdb76d84f10abf6c0545c65206a01aa1c44c8f3dc30d` | `898d541263c33848…` |

Para voltar atrás, use os backups `_backup_pre_integracao_2026-11-24_2102` (Bio) e `_2345` (Hema), além dos
`_backup_pre_atualizar_2026-11-24_2346/2347`.
