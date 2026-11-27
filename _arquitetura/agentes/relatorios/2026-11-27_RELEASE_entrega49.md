# Relatório de release — QC_INI, entrega 49 / ADR-103 (27/11/2026)

Emitido pelo `qc-release-manager` (procedimento `/qc-release`, seção 7). Fontes:
- `_arquitetura/integracao_dbseac/ADR-103.md`;
- `QC_RELEASE_STATUS.md` e `BACKLOG_POS_RELEASE.md` (BK-26 a BK-31);
- `MANIFESTO_PRODUCAO.md`, linha de 27/11 (Hematologia);
- `git log` até b97bf3b (diff dos commits só de teste eb13620 e a5a263b conferido);
- os logs dos gates em `_entrega_2026-11-26_1931/`, `_entrega_2026-11-26_2333/` (Bio) e `_entrega_2026-11-27_0157/`
  (Hema), não versionados; as contagens do §3 vêm desses logs, conforme o relato do Claude principal;
- Tester VALIDAR, script `_arquitetura/agentes/ferramentas/repro_e49_adr103.py`.

Este relatório não traz senhas, caminhos de rede nem dados de paciente.

**Mudança de escopo durante o freeze.** Autorizada pelo dono do sistema em 26/11 ("3. ok — pode!") e registrada no
ADR-103. A Tarefa 1 (regra Sigma) é uma decisão literal do usuário sobre classificação, com oráculo independente; a
fórmula do Sigma, as regras de Westgard/N/run size e o motor de dados não mudaram.

## 1. O que foi corrigido / entregue

| Item | Tipo | Commit | Produtos |
|---|---|---|---|
| ADR-103 §1: regra institucional da Métrica Sigma, literal, 5 faixas (σ ≥ 6 · 5 < σ < 6 · 4 < σ ≤ 5 · 3 ≤ σ ≤ 4 · σ < 3). **σ = 4 é Marginal** por decisão do usuário (26/11). Escada única em `mQualidade.bas` (agora versionado); a classificação deixou de ler a `tblPlanoQC_Sigma` (que segue só para a escolha de regras). Corrige o resumo Bio Estatística A98:A102 ("Desempenho inadequado" contava 0 desde o ADR-101) | regra decidida pelo usuário | 355781e | Bio, Hema |
| ADR-103 §2: "Aplicar a vários analitos" (`frmAlvosLote` + `mAlvosRef.AplicarAlvosVarios`). Mesmo lote/período para todos, uma pergunta de modo, uma confirmação, uma ciência de TEMPORÁRIO; gravação por analito pelo **mesmo núcleo** do botão de um; resumo `PARAMETRO_LOTE_CALCULADO_VARIOS` no Audit_Log. O botão de um analito não mudou | funcionalidade autorizada | 355781e | Bio, Hema |
| ADR-103 §3: Painel compacto. C4:D4 removido; H4:I4 mesclado; status J6:J9 recortado (fórmulas inalteradas) para a coluna técnica AB oculta; Classe Sigma em J; `painel_layout.py` para instaladores e suítes | interface autorizada | 355781e, c8137a3 | Bio, Hema |
| ADR-103 §4: tabela de digitação da bula/Unity em N114 e botão "TRANSFERIR VALORES PARA REFERÊNCIA INICIAL" (`TransferirReferenciaCore`): valida tudo antes, recusa inteira se houver inválido, copia só valores, entrega ao `GradeEditada`; nunca toca a LotesStore | funcionalidade autorizada | 355781e | Bio, Hema |
| Ajuste do gate 1: Classe Sigma em J:K mesclados (larguras inalteradas), Estatística M "reduzir para caber", separador TAB na janela (analitos com `#`) | acabamento antes da troca | c8137a3 | Bio, Hema |
| ADR-103 §5: "INADEQUADO" (Bio Painel A57) | **pendente de definição do usuário — nada alterado** | — | — |
| Testes: qa_final T17 conta o acréscimo do bloco de repetições e garante filtro com dados do lote (eb13620); qa_etl ajusta Ano De/Até **na cópia** para os anos do lote e E16 respeita o equipamento padrão da Hema (a5a263b) | só teste (pré-condição) | eb13620, a5a263b | — |
| MANIFESTO: Hema em produção | documentação | b97bf3b | Hema |

O cálculo estatístico validado (média, DP, CV, bias, ET, fórmula do Sigma), o motor `mEtlCQ` = `DB_CQ_FINAL.m` e as
regras de Westgard/LJ não mudaram. O status por nível do Painel mudou só de lugar (fórmulas recortadas, mesmas linhas).

## 2. Testes executados

- **Gate, 1ª tentativa** (c8137a3, `_entrega_2026-11-26_1931`): barrada no qa_final T17 **nos dois produtos**; a
  produção não foi trocada. Causa medida: premissa do teste × estado da produção, não defeito do produto. Bio já tinha
  repetições reais (inativações 1-3s de nov/2026, ação pontual de 26/11); na Hema o laboratório passou o Painel ao lote
  novo 625311 às 09:57 de 26/11 com Estatística Ano 2025 (lote sem dados nesse ano; confirmado que a ferramenta 13s
  deixou 522611). Correção só do teste (eb13620); rerun isolado 29/0 nos dois.
- **Gate, 2ª tentativa** (eb13620), `entregar.py --direcionada` + 10 `--com`, 26 suítes por produto:
  - Bioquímica `_entrega_2026-11-26_2333`: 26/26 OK; instalado sha `44f713c92b3542bf…`; blindagem TRAVADO.
    **A troca não ocorreu**: `shutil.copy2` → WinError 32 (o `QC_Bioquimica.xlsm` de produção está aberto num Excel do
    laboratório desde a madrugada). Produção intacta: sha `e9589911…` = original = backup
    `_backup_pre_integracao_2026-11-27_0156`. Script em segundo plano espera o arquivo fechar (até 10 h) e roda
    `trocar_producao.py`, que recusa se a produção tiver sido gravada desde o início dos testes.
  - Hematologia `_entrega_2026-11-27_0157`: 35 instaladores OK; blindagem TRAVADO; 9 suítes OK e qa_graficos G03
    falhou (zoom 130% não reagiu em ~1 s; o mesmo código passou na tentativa 1 e nas entregas 47/48 → **ambiente**).
    Retomada na mesma cópia (`rodar_suites --a-partir qa_graficos`): 16 OK, incl. qa_graficos 8/0 e qa_adr103 21/0;
    qa_etl falhou em E07/E13a (sem população: mesmo lote 625311 × Ano 2025) e E16 (cobrava no LJ o manual do equipamento
    B, fora do LJ por desenho — A10/E14c — exposto porque o dia do E14 passou a cair na janela do lote). Correção só do
    teste (a5a263b); rerun isolado 63/0. Troca por `trocar_producao.py` com as 26 exigidas todas OK.
- **Validação independente (Tester VALIDAR)**, `repro_e49_adr103.py`, em cópias, nunca salvas.

**Conferência da correção dos testes (diff lido):** a5a263b não afrouxa o que é verificado — só ajusta o recorte de
anos na cópia quando o lote em B4 não tem nenhum resultado nos anos B3..D3, e no E16 classifica como
`OUTRO_EQUIPAMENTO` o resultado de equipamento diferente do padrão da Hema, que o Painel não mostra por desenho (ADR do
alerta A10, conferido no E14c). eb13620 conta o acréscimo do bloco em vez de exigir contagem fixa. Aceitas.

## 3. O que passou

**Gate da 2ª tentativa: 26/26 suítes por produto.**

| Suíte | Bio | Hema |
|---|---|---|
| qa_seguranca | 20/0 | OK |
| qa_alvos_manual | 35/0 | OK |
| qa_lotes_limites | 8/0 | OK |
| qa_blocos_lote | 14/0 | OK |
| qa_linhas_alvo | 10/0 | OK |
| qa_varredura_lotes | 4/0 | OK |
| qa_troca_lote | 3/0 | OK |
| qa_lotes_auto | 14/0 | OK |
| qa_westgard_sequencia | 29/0 | OK |
| qa_graficos | 8/0 | 8/0 (na retomada; G03 ambiente) |
| qa_adr090 | 21/0 | OK |
| qa_adr103 | 21/0 | 21/0 |
| qa_incerteza | 15/0 | OK |
| qa_desempenho | 5/0 | OK |
| qa_final | 29/0 | OK (29/0 no rerun isolado da tentativa 1) |
| qa_etl | 64/0 | 63/0 (rerun isolado após a5a263b) |
| qa_migracao_adr070 | 2/0 | OK |
| qa_sync | OK | OK |
| qa_adr084 | 17/0 | OK |
| qa_usuarios_adr084 | 12/0 | OK |
| qa_adr085 | 15/0 | OK |
| qa_adr088 | 16/0 | OK |
| qa_adr089 | 10/0 | OK |
| qa_adr092 | 7/0 | OK |
| qa_adr095 | 10/0 | OK |
| qa_fluxo_manual | 19/0 | OK |

Na Hema, as contagens por suíte não foram repassadas a este relatório além das citadas; o `trocar_producao.py` exige as
26 OK e aceitou a troca.

**Tester VALIDAR: VALIDADA nos dois produtos, itens 1–5. Bio 24/0, Hema 23/0.**
- Sigma: oráculo em 22 fronteiras; Estatística M e Painel J = oráculo em todas as linhas (Bio 58/57, Hema 31/28).
- Painel compacto sem `#REF`.
- Transferência: só valores; recusa total em 4 casos inválidos; cancelar não grava.
- `AplicarAlvosVarios` = `AplicarAlvos` isolado em duas cópias; Audit_Log com 1 `PARAMETRO_LOTE_CALCULADO_VARIOS`.
- Regressão: qa_adr103 21/0 e qa_alvos_manual 35/0 nos dois; Y00 sha da produção inalterado.
- Observações sem defeito: Fonte em formato texto; dados reais da Bio sem σ > 5 (as faixas altas foram provadas na
  Hema e nas fronteiras). Proposta V07 (cópia A × B no modo 2) → BK-31.

## 4. O que não foi feito e por quê

| Item | Motivo | Consequência |
|---|---|---|
| **Troca da Bioquímica** | produção aberta num Excel do laboratório (WinError 32) | Bio segue na entrega 48 (`e9589911…`). A troca roda sozinha quando o arquivo fechar; se a produção for gravada antes, o `trocar_producao.py` recusa e o **gate Bio é refeito** |
| Tarefa 5 "INADEQUADO" (Bio Painel A57) | aguarda definição do usuário | nada alterado; não é a regra Sigma |
| Verificação visual do Painel compacto e da tabela N114 | política sem captura de tela | provado por qa_adr103 e pelo Tester (sem `#REF`, posições); leitura visual fica com o usuário |
| BI (modelo e visuais com as 5 classes) | decisão do usuário ("sem desenvolver BI por enquanto") | BK-26; a BI_Data já recebe os rótulos novos |
| Suíte completa / `equivalencia_cq --fixtures` | o ciclo usa a direcionada; o motor não mudou (última prova: gate 47, 0 divergências) | qa_etl e qa_sync cobrem o recebimento |
| Push | não autorizado nesta etapa | commits 355781e..b97bf3b só locais |
| Aprovação do RT e da qualidade | etapa humana documentada | fora da release técnica |

## 5. Pendências que ficam: nenhuma é P0 ou P1

| ID | Componente | Evidência | Prior. | Conduta |
|---|---|---|---|---|
| TROCA-BIO | produção Bio | WinError 32 no gate `_entrega_2026-11-26_2333`; sha produção = backup | pendência operacional (não é defeito) | aguardar o fechamento; conferir no MANIFESTO/log do script. Se recusar, refazer o gate Bio |
| VERSÕES-MISTAS | Bio (48) × Hema (49) em produção | MANIFESTO 27/11 | observação | enquanto a Bio não trocar, a Bio classifica o Sigma pelas 6 faixas do ADR-101 e a Hema pela regra institucional; sem efeito no número do Sigma |
| HEMA-ANO | Estatística da Hema na produção | Painel no lote 625311 com Ano De/Até 2025 (lote sem dados em 2025) | dado de configuração do usuário, não defeito | **recomendar ao usuário/RT** ajustar Ano De/Até para 2026; até lá a Estatística da Hema aparece vazia |
| TAREFA-5 | Bio Painel A57 | ADR-103 §5 | aguarda o usuário | nada alterado |
| V07 | testes/qa_adr103 | proposta do Tester | P3 | BK-31 |
| BI | PowerBI/ | ADR-103 §1 "Fora" | P3 | BK-26 |
| Cores/lint/largura/CEQ Hema | Analitos, lint_vba, N114:X154, qa_fluxo_ceq | ADR-103 | P3 | BK-27 a BK-30 (já no backlog) |
| OC-0067 / BK-05 e herdadas | senha das abas; E01 de governança; OC-0011, OC-0039 etc. | `QC_RELEASE_STATUS.md` | P2 | sem mudança |

## 6. Adiados / backlog

Acréscimo deste ciclo: **BK-31** (V07 do Tester na qa_adr103). BK-26 a BK-30 já tinham sido registrados durante a
implementação. Nada adiado é P0/P1.

## 7. Status final

Conferência do gate:
- **0 P0 e 0 P1 conhecidos.**
- Validação integrada PASS: 26/26 suítes nos dois produtos; blindagem TRAVADO. As duas falhas de suíte da 2ª tentativa
  na Hema foram classificadas antes de qualquer alteração (G03 = ambiente; qa_etl = pré-condição do teste) e a correção
  ficou restrita ao teste, com rerun isolado.
- O gate funcionou: a 1ª tentativa foi barrada e a produção não foi trocada; na Bio a troca não forçou o arquivo aberto.
- Cálculo: a regra Sigma institucional bate com o oráculo independente em 22 fronteiras e em todas as linhas reais das
  duas abas; o resto do cálculo e o motor não mudaram.
- Integridade: Y00 sha inalterado; `AplicarAlvosVarios` idêntico ao botão de um; transferência nunca toca a LotesStore.
- Proteção e recuperação: Hema backup `_backup_pre_integracao_2026-11-27_0518`; Bio intacta (backup
  `_backup_pre_integracao_2026-11-27_0156`); sha no MANIFESTO.
- Documentação: ADR-103, MANIFESTO (Hema), BACKLOG, STATUS, este relatório.

**QC_INI — RELEASE CANDIDATE APROVADO PARA ENTREGA** (entrega 49, release técnica do código)

Condições registradas:
1. **Hematologia entregue** (b8c50114, 27/11 05:18). **Bioquímica aguardando a troca** — o arquivo validado
   (`44f713c9…`) só vai à produção pelo `trocar_producao.py` quando a produção for fechada; se houver gravação na
   produção nesse meio tempo, o gate Bio é refeito antes da troca.
2. Ficam de fora, como etapas humanas: a verificação visual do Painel compacto e da tabela N114 pelo usuário; a
   definição da Tarefa 5; a aprovação institucional e regulatória do RT e da qualidade.
3. Os commits 355781e, c8137a3, eb13620, a5a263b e b97bf3b ainda **não foram enviados ao remoto**; o push vem com o
   diff conferido.

## 8. Versão recomendada

**Entrega 49**, produto no commit `c8137a3` (testes em `a5a263b`):

| Produto | Estado | sha256 | Backup anterior |
|---|---|---|---|
| QC_Hematologia.xlsm | em produção desde 27/11 05:18 | `b8c501144f1c8e0faa717b52e0a7364c9aa4de65b88bbc01409e911af7d82496` (41.003.831 bytes) | `_backup_pre_integracao_2026-11-27_0518` |
| QC_Bioquimica.xlsm | aguardando troca (produção ainda na 48 + ATUALIZAR de 26/11) | validado `44f713c92b3542bf…`; produção atual `e9589911bebc0ff3…` | `_backup_pre_integracao_2026-11-27_0156` |

Para voltar a Hema à entrega 48, use o backup acima. O sha muda no primeiro ATUALIZAR da tarefa agendada, o que é
esperado; o sha acima é o da troca.
