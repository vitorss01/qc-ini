# Continuar em casa: situação em 11/11/2026, às 11h

A validação foi interrompida às 10h45 a pedido do usuário, porque o PC do trabalho estava lento. Este arquivo diz o que já está pronto, o que falta e como retomar sem começar do zero.

## 1. O que está em produção (no PC do trabalho)
| Arquivo | sha256 | O que tem |
|---|---|---|
| QC_Bioquimica.xlsm | `8b1cda1345d2d6ff…` | entrega 38 (ADR-091) + ATUALIZAR de 11/11, 08h18 |
| QC_Hematologia.xlsm | `f1907656b663ea6b…` | entrega 38 (ADR-091) + ATUALIZAR de 11/11, 08h21 |

O histórico completo está em `_arquitetura/integracao_dbseac/MANIFESTO_PRODUCAO.md`.

**Os `.xlsm` não vão para o GitHub**, porque o repositório é público e eles têm dados do laboratório. Os backups `_backup_*` também ficaram de fora. Para ter os arquivos em casa, use o OneDrive corporativo (a pasta já está nele), ou traga-os por um meio autorizado pelo serviço.

## 2. Validação do ADR-090 (19 ocorrências) e do ADR-091 (material pelo lote)

**Tester (VALIDAR):** concluído. O relatório está em `_arquitetura/agentes/relatorios/2026-11-11_reteste03/tester_validar.md`. O script dele é `_arquitetura/agentes/ferramentas/valida_adr090_091.py`, e as evidências estão em `.../2026-11-11_reteste03/evidencias/`.

| Resultado do Tester | Ocorrências |
|---|---|
| VALIDADA | OC-0012, 0019, 0021, 0023, 0024, 0031, 0033, 0034, 0035, 0036, 0044, 0050, 0052 (Hema), 0053, 0054, 0055; **ADR-091 Bio e Hema** |
| NAO_SE_APLICA | OC-0052 (Bio) |
| **NAO_VALIDADA** | **OC-0040** (Estatística D4 ainda mostra "Ano 0"; o ADR-090 corrigiu a OC-0041 com o rótulo errado), **OC-0043** (os status longos ainda ficam cortados em J7:J9), **OC-0049** (B98:B133 da Hema ainda mostram o nível nas linhas vazias) |

**Especialista em CQ:** concluído. A regra do ADR-091 foi VALIDADA tecnicamente. O parecer está em `.../2026-11-11_reteste03/parecer_cq_adr091.md`, com os riscos R1–R5 (lote novo de urina sem cadastro, sufixo -EE, 19 resultados de PCR/HbA1c com lote provavelmente digitado errado, CQ de urina sem estatística).

**Usuário Simulado (reteste):** **NÃO FOI EXECUTADO**, por causa da interrupção. Pela regra do projeto, nenhuma ocorrência fica VALIDADA só com o Tester. Por isso as 19 continuam como `AGUARDA RETESTE` no `OCORRENCIAS.md`.

**Regressão:**
- Bio: qa_final com 27 aprovados e 2 falhas (T15 e T19), que o Tester atribui a resultados novos do DB_SEAC chegando durante a suíte, ainda a confirmar. qa_rodadas_cv 12/0, qa_adr090 22/0, qa_adr091 7/0, equivalência VBA×PQ OK.
- Hema: qa_final 29/0, rodadas 12/0, adr090 23/0, adr091 7/0. A equivalência da Hema foi interrompida no meio.

## 3. Ocorrências novas encontradas pelo Tester (registrar no OCORRENCIAS.md)
| ID provisório | Prioridade | Resumo |
|---|---|---|
| N-01 | **Alto** | Um manual digitado num lote de URINA com o nome do analito de soro (ex.: 6354 + "Creatinina") entra na estatística do soro. O `mEtlCQ.bas:1019` não acha o analito de urina no de/para. Hoje está latente (0 manuais assim na produção) |
| N-02 | Médio | `CorrigirCEQ`/`ExcluirCEQ` gravam sem sessão (`mCEQForm.bas:409/:542` não chamam `SessaoAtiva`) |
| N-03 | Baixo | O Audit_Log grava RESULTADO_MANUAL_INCLUIDO para manual IMPLAUSIVEL/LOTE_* (`mIntegracao.bas:557-561`) |
| N-04 | Baixo | Status longos cortados em J7:J9 (forma curta = decisão de UX) |
| N-05 | Baixo | O ADR-090 chama de OC-0040 a correção que é da OC-0041 (só documentação) |

O Tester já sugeriu os casos de teste: T07 no qa_adr091, O33b e O23b no qa_adr090.

## 4. Como retomar (ordem sugerida)
1. **Confirmar T15/T19 do qa_final Bio** numa cópia sem ATUALIZAR concorrente, para saber se é só dado novo chegando.
2. **Corrigir OC-0040, OC-0043, OC-0049 e N-01..N-03** (ADR-092). A N-01 é Alta. Pela regra, ela precisa de decisão de CQ sobre o manual em lote de urina: recusar ou mapear para o analito de urina.
3. **Entregar** (entregar.py, com todas as suítes) e depois rodar o ATUALIZAR na produção com `_arquitetura/agentes/ferramentas/aplicar_atualizar.py`.
4. **Rodar o Usuário Simulado** (reteste) das 19 ocorrências do ADR-090, das corrigidas e do ADR-091. Depois, o Tester VALIDAR só do que mudou.
5. Marcar VALIDADA no `OCORRENCIAS.md` só onde o Tester disse VALIDADA **e** o Usuário Simulado disse RESOLVIDA.

Para rodar os agentes em casa:
- `python _arquitetura/agentes/instalar_agentes.py <pasta onde você abre o Claude Code>`, e depois abra uma sessão nova.
- A skill `/auditoria-qc` usa o modo `reteste`.
- Os prompts usados hoje estão no histórico do workflow `validacao-adr090-091`, e o resumo deles está neste arquivo.

## 5. Contexto da sessão (memória)
`_arquitetura/agentes/memoria_sessao/` tem as notas que o assistente acumulou: regras do usuário, armadilhas de Excel/COM, decisões ADR-067..091. O caminho de rede foi trocado por SERVIDOR/COMPARTILHAMENTO.

Regras que valem sempre:
- repositório público: nada de senha, caminho de rede interno ou `.xlsm`;
- trabalhar só na main;
- testar em cópias;
- produção só pelo entregar;
- um Excel por vez;
- não mexer no BI;
- nunca capturar a tela (dados de paciente);
- não alterar fórmula, regra ou critério estatístico sem autorização;
- filtros só mudam a visualização.
