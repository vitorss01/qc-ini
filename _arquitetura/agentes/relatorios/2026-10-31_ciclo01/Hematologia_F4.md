# Hematologia, F4 (períodos, lotes, filtros, gráficos e usabilidade): ciclo 2026-10-31_ciclo01

Agente: Usuário Simulado. Modo: completo. Data: 31/10/2026, das 22h40 às 23h00.

## Resumo

**O Painel está certo.** Os filtros abaixo calcularam exatamente sobre os dados brutos filtrados, no motor
(Eng_Saida) e não só na tela:

- todos os meses com dado, out/2025 a out/2026, nos dois lotes dos meses de troca;
- os cinco trimestres, S1, S2, "Ano inteiro" 2025 e 2026 e "Tudo";
- os lotes A → B → H3 vazio;
- oito analitos.

Foram 0 divergências, com diferença relativa máxima de 7,2e-16. Também bateram ponto a ponto com o recálculo
independente:

- o Levey-Jennings dos 3 níveis: pontos, corrida do lote do ADR-080 e limites;
- o Westgard N=3 do ADR-083: 1_3s, 2of3_2s, R_4s, 3_1s, 6x, o alerta 1_2s e o veredito.

Isso vale em 5 cenários de lote e no modo TODOS: 1.261 pontos comparados, 0 divergências.

**Ficaram reprovados:**

- **A aba Estatística não acompanha o lote nem o ano (OC-HF4-1, Crítico).** Troquei o lote no Painel. B4 passou a
  dizer 619611 "(refletido do Painel)", mas n, média, DP e CV continuaram os do lote anterior (522611). O mesmo
  acontece quando mudo Ano De/Até na própria aba e quando entro nela pelo botão. Só o ATUALIZAR DADOS recalcula. O
  cálculo em si está certo: depois do recálculo interno, 0 divergências.
- **H3 vazio com período longo (OC-HF4-2).** O n e as contagens de Westgard são só das 180 corridas mais recentes, e
  o aviso fala só do gráfico.

Há ainda achados de coerência e usabilidade, OC-HF4-3 a OC-HF4-9. Vários repetem os da Bioquímica.

## 1. Ambiente e evidência

| Item | Valor |
|---|---|
| Produção copiada | `QC_Hematologia.xlsm`, sha256 curto `0b45fa270d2a`, gravado em 30/10/2026 22:11. O hash da produção é o mesmo antes e depois: ela não foi aberta |
| Cópia | `%TEMP%\qcw\auditoria\2026-10-31_ciclo01\hema_f4\Hematologia_F4.xlsm` (`shutil.copy2`). O hash ficou `0b45fa270d2a` depois das 3 sessões, ou seja, nada foi salvo. A cópia foi apagada no fim |
| Sessões do Excel (uma por vez, cada uma encerrada no `finally` pelo próprio PID) | PID 3484: leitura do layout (Painel, Estatística, nomes, séries). PID 23680: roteiro principal, 22h44–22h47. PID 5428: conferência da Estatística pelo botão de navegação e lista dos módulos VBA |
| Base do cálculo independente | `cq.pkl`: a `tblCQ_Final` da mesma cópia lida por openpyxl, sem Excel, com 44.392 linhas (43.710 ativas, todas XN1000). O T00 confere que os IDs são idênticos aos do COM. A base vai de 01/10/2025 a 01/10/2026, a mesma da F3: a produção não foi atualizada depois da entrega (ver F1, dúvida 3). Esta frente não roda o ATUALIZAR |
| Ferramentas | `_arquitetura/agentes/ferramentas/f4_hema.py` (novo; reaproveita o cálculo do `f4_periodos_lotes.py` e acrescenta 3 níveis, Westgard N=3, Estatística por ano e modo TODOS) e `f4_hema_conf.py` (complemento da OC-HF4-1) |
| Saídas | `hema_f4\saida\f4h.json`, `f4h.log`, `f4h_conf.json`, `explorar.json`. Imagens exportadas pelo próprio Excel da cópia: `saida\Painel_set_619611.png`, `saida\Painel_todos_fevmar.png` e `saida\Estatística.png`. A tela do usuário não foi capturada |
| Dados de teste | Nenhum. A F4 só mexe em filtros |

**Caminho do usuário**

- **Painel:** C3, H3, K3, K4, G3 e G4 são desbloqueadas na aba protegida (`mSeguranca.ReprotectAll`). Escrevi nelas
  com os eventos ligados.
- **Estatística:** B3 e D3 (Ano De/Até) são desbloqueadas e também foram escritas assim.
- **Fora do caminho do usuário:** o `mEstatistica.AtualizarEstatistica`, que é o trecho do ATUALIZAR que refaz a
  Estatística, foi chamado direto nos testes E-INTERNO e E-ANO-INVERTIDO. Esses testes valem como conferência do
  cálculo da aba, não como caminho do usuário.

**Hematologia × Bioquímica**

- Não há seletor de equipamento: tudo é XN1000 (`EquipFiltro` = EQUIPAMENTO_PADRAO). A combinação "outro
  equipamento" é **não aplicável**.
- A Estatística da Hema filtra por **ano** (B3/D3), não por mês ou trimestre (`EstatBasica`).
- Nenhum lote tem mais de 180 corridas (o maior, 528211, tem 89). O corte da janela no lote escolhido não se aplica,
  mas aparece no H3 vazio (OC-HF4-2).

## 2. Testes executados

Tolerância: 1e-9 relativo sobre o `Value2`.

**Cálculo independente:** `statistics` sobre as linhas brutas com PARTICIPA_ESTATISTICA = SIM, filtradas por analito,
XN1000, lote, nível e dia de DATA_HORA. O DP é amostral (n−1). O alvo do lote vem do cadastro
(`mLotes.MediaDoLote`/`DPDoLote`).

### 2.1 Períodos (WBC), pelo Painel (K3 Ano, K4 Período, lote em H3)

| ID | Período | Lote | Estado | Painel N1 / N2 / N3: n; média; CV % | Maior dif. rel. |
|---|---|---|---|---|---|
| P-M202510-522611 | out/2025 | 522611 | APROVADO | 50; 3,0446; 1,80 / 51; 6,7765; 1,50 / 51; 16,7516; 0,91 | 4,4e-16 |
| P-M202510-528211 | out/2025 (2º lote) | 528211 | APROVADO | 2; 3,015 / 2; 6,925 / 2; 16,34 | 0 |
| P-M202511-528211 | nov/2025 | 528211 | APROVADO | 45; 2,9993; 1,48 / 49; 6,8171 / 45; 16,4880 | 3,3e-16 |
| P-M202512-528211 | dez/2025 | 528211 | APROVADO | 36; 2,9817; 2,39 / 35 / 34 | 3,7e-16 |
| P-M202512-533811 | dez/2025 (2º lote) | 533811 | APROVADO | 8; 3,2088 / 8 / 7 | 0 |
| P-M202601-533811 | jan/2026 | 533811 | APROVADO | 52; 3,1569; 1,97 / 43 / 41 | 3,4e-16 |
| P-M202602-536411 | fev/2026 | 536411 | APROVADO | 21; 3,0043 / 22 / 21 | 1,7e-16 |
| P-M202602-533811 | fev/2026 (2º lote) | 533811 | APROVADO | 16; 3,2006; **6,19** / 16 / 18; CV 3,06 | 4,4e-16 |
| P-M202603-536411 | mar/2026 | 536411 | APROVADO | 37 / 35 / 34 | 4,5e-16 |
| P-M202603-602911 | mar/2026 (2º lote) | 602911 | APROVADO | 17 / 22 / 16 | 3,7e-16 |
| P-M202604-602911 | abr/2026 | 602911 | APROVADO | 22 / 22 / 22 | 1,1e-16 |
| P-M202604-608511 | abr/2026 (2º lote) | 608511 | APROVADO | 13 / 11 / 10 | 1,4e-16 |
| P-M202605-608511 | mai/2026 | 608511 | APROVADO | 30 / 26 / 28 | 3,8e-16 |
| P-M202606-614011 | jun/2026 | 614011 | APROVADO | 21 / 17 / 16 | 4,2e-16 |
| P-M202606-608511 | jun/2026 (2º lote) | 608511 | APROVADO | 14 / 13 / 12 | 2,1e-16 |
| P-M202607-614011 | jul/2026 | 614011 | APROVADO | 34 / 38 / 35 | 2,9e-16 |
| P-M202608-619611 | ago/2026 | 619611 | APROVADO | 23 / 22 / 24 | 2,1e-16 |
| P-M202608-614011 | ago/2026 (2º lote) | 614011 | APROVADO | 5 / 5 / 5 | 2,1e-16 |
| P-M202609-619611 | set/2026 | 619611 | APROVADO | 31; 3,1274; 2,42 / 46; 7,0033; 1,21 / 25; 16,8016; 0,76 | 2,1e-16 |
| P-M202610-625311 | out/2026 (só 01/10) | 625311 | APROVADO | 1; 3,18; DP/CV vazios (n < 2) / 2; 7,33 / 1; 17,44 | 0 |
| P-T20254 | T4/2025 | 528211 | APROVADO | 83; 2,9920; 1,925 / 86; 6,8033; 1,459 / 81; 16,4441; 1,373 | 2,2e-16 |
| P-T20261 | T1/2026 | 533811 | APROVADO | 68; 3,1672; 3,472 / 59; 6,8515; 1,718 / 59; 15,5712; 2,029 | 4,4e-16 |
| P-T20262 | T2/2026 | 608511 | APROVADO | 57; 3,1291; 2,418 / 50 / 50 | 4,9e-16 |
| P-T20263 | T3/2026 | 619611 | APROVADO | 54; 3,1206; 2,275 / 68; 6,9947; 1,283 / 49; 16,8612; 0,874 | 4,3e-16 |
| P-T20264 | T4/2026 | 625311 | APROVADO | igual a out/2026 | 0 |
| P-LS12026 | S1/2026 | 533811 | APROVADO | igual a T1 (o lote só tem jan–fev) | 4,4e-16 |
| P-LS22026 | S2/2026 | 619611 | APROVADO | igual a T3 | 4,3e-16 |
| P-LAno 2025 | "Ano inteiro" 2025 | 528211 | APROVADO | igual a T4/2025 | 2,2e-16 |
| P-LAno 2026 | "Ano inteiro" 2026 | 533811 | APROVADO | igual a T1/2026 | 4,4e-16 |
| P-LTudo2026 | "Tudo (sem filtro)" | 528211 | APROVADO | 83 / 86 / 81, todo o histórico do lote (89 corridas, sem corte) | 2,2e-16 |
| G3/G4 | todos os meses e trimestres | — | APROVADO | O atalho Ano/Período preencheu De/Até com o 1º e o último dia certos em todos os casos | — |
| P-TODOS-Ano | H3 vazio + "Ano inteiro" 2026 | todos | **REPROVADO** | n do Painel 141 / 155 / 130. Todos os resultados de 2026: **337 / 340 / 308**. O n do Painel é exatamente o da janela das 180 corridas RUN\|LOTE mais recentes, com 213 corridas do período cortadas. Ver OC-HF4-2 | n |
| P-TODOS-Tudo | H3 vazio + "Tudo" | todos | **REPROVADO** | 141 / 155 / 130 contra **478 / 485 / 447**, com 366 corridas cortadas. Ver OC-HF4-2 | n |

**O trimestre é calculado dos dados brutos do trimestre, e não como média dos meses.** Conferido nos cinco
trimestres: o Painel bate com os dados brutos. A "média dos meses" daria outro número:

- **T4/2025, lote 528211, N1:** média das médias mensais 2,9987 e média dos CV mensais 1,369 %, contra o Painel com
  2,9920 e 1,925 %;
- **T1/2026, lote 533811, N1:** 3,1788 e 4,082 %, contra 3,1672 e 3,472 %.

Nenhum ADR manda outra coisa. **Tempos:** K4 de 0,1 a 0,3 s; K3 de 0,1 a 0,2 s; G3/G4 de 0,1 a 0,2 s; H3 de 1,3 a
3,0 s; C3 de 0,15 a 3,0 s.

### 2.2 Aba Estatística (84 linhas analito × nível, colunas C:F)

| ID | Teste | Estado | Evidência |
|---|---|---|---|
| E-BASE-522611 | Estado de abertura (lote 522611, Ano 2025–2025) | APROVADO | 336 comparações, 0 divergências, máx. 7,9e-16. As únicas 3 "divergências" do JSON são uma linha de catálogo vazio (analito "0", n vazio), artefato da ferramenta. WBC 50 / 51 / 51, igual à F3 |
| E-LOTE-619611 | O usuário troca o lote no Painel (H3 = 619611): a Estatística acompanha? | **REPROVADO** | B4 = "619611 (refletido do Painel)". C:F **continuam os do 522611**: WBC N1 n = 50, média 3,0446, DP 0,054929, CV 1,8041 %, idênticos ao recálculo do 522611. O 619611 não tem nenhum resultado em 2025 e deveria mostrar n = 0. Ver OC-HF4-1 |
| E-ANO-2026-619611 | O usuário muda Ano De/Até para 2026/2026 | **REPROVADO** | B3/D3 aceitos (0,1 s), mas os números não mudam: 339 de 339 comparações divergem. WBC N1: tela 50; 3,0446 contra 54; 3,12056 do 619611 em 2026 |
| E-NAV (f4_hema_conf) | Entrar na aba pelo botão "Estatística" (`mApp.IrEstatistica`), sair e voltar | **REPROVADO** | Continua com os números do 522611 (50; 3,0446) e B4 = 619611. O módulo da planilha Estatística (Planilha10) tem **0 linhas** de código: não há evento que recalcule a aba |
| E-INTERNO-619611 | Depois de `mEstatistica.AtualizarEstatistica` (passo interno do ATUALIZAR) | APROVADO (cálculo) | 336 comparações, 0 divergências, máx. 1,1e-15. WBC 54; 3,12056; 0,070988 / 68; 6,99471 / 49; 16,86122 |
| E-ANO-INVERTIDO | Ano Até (2025) < Ano De (2026) | APROVADO (cálculo) | Vale só o Ano De, como manda o `EstatBasica`. A tela não avisa e continua mostrando "Ano Até 2025" (ver OC-HF4-9) |
| L-*/Estat | A Estatística durante os testes de lote (536411 → 602911 → H3 vazio) | **REPROVADO** (confirma a OC-HF4-1) | B4 = 536411, depois 602911, depois 522611 (H3 vazio → `loteAnalise` = lote em uso). Os números ficaram sempre os do 619611/2026 do último recálculo |

### 2.3 Lotes: A → B → H3 vazio (WBC, De/Até 01/02–31/03/2026, período com 3 lotes)

| ID | Teste | Estado | Evidência |
|---|---|---|---|
| L-536411 | Lote 536411: estatística só do lote | APROVADO | N1 58; 3,016897; DP 0,061877. N2 57; 7,095789. N3 55; 17,288545. Máx. 4,3e-16 |
| L-536411-MOTOR | Eng_Saida só com pontos do lote, analito, nível e período, e com todos eles | APROVADO | 170 pontos, todos do 536411. 0 fora, 0 faltando |
| L-602911 (+MOTOR) | Lote 602911 | APROVADO | Estatística e motor só do lote. 0 fora, 0 faltando |
| L-TODOS | H3 vazio (ADR-086) | APROVADO no motor | engTodos = VERDADEIRO. H4 "H3 vazio = TODOS os lotes do período...". n = 91 / 95 / 89, igual ao independente. Pontos: 533811 = 50, 536411 = 170, 602911 = 55. Média/DP/CV/ET vazios no `engPainel`. **Mas** a tela mostra Bias 1,86 % e Sigma 2,27 / 3,58 / 5,68 (OC-HF4-3). O A39 fala do lote 522611, que não tem ponto no período, e a "Últ. violação" é R_4s · RUN 25102801, de out/2025 (OC-HF4-4/5) |
| L-TODOS-ALVO | Cada ponto com a Média/DP do seu lote (Eng_Saida BW..CB) | APROVADO | 0 divergências contra o cadastro |
| L-SOBREPOSTOS | H3 vazio, 15–31/12/2025, com 528211 e 533811 rodando em paralelo | APROVADO | Os RUN 25122201, 25122301 e 25122302 existem nos dois lotes e viraram **2 corridas cada** (RUN\|LOTE). n = 22 / 22 / 21, igual ao independente. 0 fora, 0 faltando |
| L-SOBREPOSTOS-X | Eixo X do modo TODOS: corrida 1, 2, 3... por nível, atravessando os lotes | APROVADO | 0 divergências contra o recálculo das chaves RUN\|LOTE de todo o histórico |
| L-AJUSTE-1 | Lote corrente 619611 em Set/2026 → lote em uso 522611 (sem ponto em set) → volta ao 619611 | APROVADO (ajuste) | No 522611, o período vira "Tudo" com aviso em H4 e A39 ("o filtro anterior não mostrava nenhum resultado dele -- ajustado: período Tudo"). Ao **voltar ao 619611, o "Set" não volta** (fica "Tudo"). Segue o ADR-078, mas o usuário perde a escolha (OC-HF4-6) |
| L-AJUSTE-2 | Lote em uso 522611 em Out/2025 → 619611 (sem ponto) → volta ao lote em uso | APROVADO | Ajusta para "Tudo" no 619611. De volta ao 522611, Out/2025 e 01–31/10/2025 voltam, com o A39 "Filtro da rotina restaurado (de volta ao lote em uso)", como diz o GUIA §4b |

### 2.4 Outras combinações (lote 619611, set/2026; equipamento único)

| ID | Analito | Estado | Painel N1 / N2 / N3: n; média; DP (máx. dif.) |
|---|---|---|---|
| C-RBC (+MOTOR) | RBC | APROVADO | 31; 2,3155; 0,02873 / 46; 4,1974 / 25; 4,944 (7,2e-16) |
| C-HGB (+MOTOR) | HGB | APROVADO | 31; 5,7065 / 46; 11,8261 / 25; 15,604 (3,5e-16) |
| C-PLT (+MOTOR) | PLT | APROVADO | 31; 83,7419 / 46; 232,2391 / 25; 546,16 (4,5e-16) |
| C-MCV (+MOTOR) | MCV, ETp de VB | APROVADO | 31; 72,2065 / 46; 79,287 / 25; 88,44 (2,6e-16) |
| C-NEUT% (+MOTOR) | NEUT% | APROVADO | 31; 42,5161 / 46; 44,8783 / 25; 47,776 (2,6e-16) |
| C-IG% (+MOTOR) | IG%, sem ETp | APROVADO | ETp "-", Sigma "-" (ADR-068). Bias 3,009 % e ET 8,14 / 7,52 / 6,73 % |
| C-RET% (+MOTOR) | RET%, ETp sem CEQ | APROVADO no cálculo; **confirma a OC-H2 da F3** | G = "SEM EP" e H = **5,8099 % = 1,65 × CV 3,52114**, ou seja, ET sem bias. Sigma vazio. Ver a confirmação na seção 3 |
| C-RDW-SD (+MOTOR) | RDW-SD, ETp sem CEQ | APROVADO no cálculo; confirma a OC-H2 | G "SEM EP", H = 1,5478 = 1,65 × 0,93806 |

Em todos os analitos, o motor tem 0 pontos fora do escopo e 0 faltando.

### 2.5 Levey-Jennings e Westgard N=3 (5 cenários + modo TODOS)

**Definição usada no Westgard independente**, conferida em `AvaliarWestgard`, nos detectores e na tabela
`DETECTORES` da Hema:

| Regra | O que dispara | O que marca |
|---|---|---|
| 1_3s | \|z\| > 3 | o ponto |
| 2of3_2s WITHIN_RUN | com ≥ 2 níveis presentes no RUN, 2 ou mais do mesmo lado além de 2s | os níveis que passam |
| R_4s | qualquer par de níveis do RUN com \|z1 − z2\| > 4 | os dois níveis |
| 3_1s N3_R1 | RUN completo, os 3 níveis além do mesmo 1s | os 3 níveis |
| 6x N3_R2 | 2 RUNs completos seguidos no eixo compactado das corridas elegíveis, 6 observações do mesmo lado da média | os 3 níveis do 2º RUN |
| alerta 1_2s | \|z\| > 2 | o ponto |

Os detectores complementares SAME_LEVEL têm Ativo = 0 e não marcam.

| Cenário | EIXO (motor = janela) | PONTOS N1 / N2 / N3: n, X da corrida do lote | LIMITES (7 linhas = Média + k·DP do cadastro) | Tabela × gráfico | WESTGARD por ponto | Contagens M7:Q9 e status |
|---|---|---|---|---|---|---|
| WBC 619611, Set/2026 | APROVADO (77) | APROVADO: 31, X 24..54 / 46, X 23..68 / 25, X 25..49, sem buraco | APROVADO: 3,12056 ± 0,07935; 6,99471 ± 0,1368; 16,86122 ± 0,25395 | APROVADO | APROVADO: 171 pontos, 0 divergências | APROVADO: 0 em todas as regras, OK/OK/OK |
| WBC 614011, T3/2026 | APROVADO (73) | APROVADO: 39, X 22..60 / 43, X 18..60 / 40, X 17..56 | APROVADO | APROVADO | APROVADO: 176, 0 | APROVADO: 0, OK |
| PLT 608511, Mai/2026 | APROVADO (47) | APROVADO: 30, X 14..43 / 26, X 12..37 / 28, X 11..38 | APROVADO: 79,789 ± 7,029; 227,78 ± 16,314; 530,88 ± 22,277 | APROVADO | APROVADO: 118, 0 | APROVADO: 6x = 6 em cada nível, REJEITADO ×3 |
| HGB 533811, Jan/2026 | APROVADO (62) | APROVADO: 52, X 9..60 / 43, X 9..51 / 41, X 8..48 | APROVADO | APROVADO | APROVADO: 159, 0 | APROVADO: 6x = 9 por nível, REJEITADO ×3 |
| RBC 522611, Tudo | APROVADO (56) | APROVADO: 50 / 51 / 51, X a partir de 1 | APROVADO: 2,3066 ± 0,0219; 4,3769 ± 0,035; 5,1228 ± 0,0345 | APROVADO | APROVADO: 152, 0 | APROVADO: N1 [1_3s 40, 2of3 8, R4s 39, 3_1s 1, 6x 2], N2 [2, 0, 39, 1, 2], N3 [4, 8, 18, 1, 2] |
| WBC H3 vazio, fev–mar/2026 (ADR-086) | APROVADO: 180 chaves RUN\|LOTE iguais (lotes 528211, 533811, 536411, 602911; 128 corridas antes da janela) | APROVADO: 91 / 95 / 89, X 194..284 sem buraco atravessando 3 lotes | APROVADO: os degraus = Média + k·DP do lote de cada ponto, 0 divergências | — | APROVADO: lote a lote, 485 pontos, 0 divergências | APROVADO: N1 [2, 0, 2, 2, 7], N2 [1, 1, 2, 2, 7], N3 [3, 1, 1, 2, 7] |

**Notas**

- **Ordem das colunas na tela:** M (1-3S), N (3-1S), O (6X), P (2of3-2S), Q (R4S) foram conferidas contra o
  `engPainel` 13..17, que está na ordem 1_3s, 2of3, R_4s, 3_1s, 6x.
- **Corrida do lote:** cada nível tem a sua sequência (ADR-080). No WBC 619611, o N2 tem um controle extra à tarde,
  o que dá 68 corridas, contra 54 do N1. As colunas BK..BM do Eng_Saida bateram com o recálculo em todos os cenários.

### 2.6 Usabilidade (imagens exportadas da cópia e `auditoria_abas.auditar_aba`)

| ID | Teste | Estado | Evidência |
|---|---|---|---|
| U-AUDIT | Textos cortados e formas sobrepostas (Painel e Estatística) | APROVADO | `auditar_aba`: 0 achados nas duas abas |
| U-K4 | Rótulo do período depois de digitar De/Até | REPROVADO (Médio) | Com De/Até 01/02–31/03/2026, K3 continua "2025" e K4 "Tudo (sem filtro)". Ver `Painel_todos_fevmar.png` e OC-HF4-7 |
| U-ULTVIOL | "Últ. violação" sob o título "WESTGARD — violações no período" | REPROVADO (Médio) | WBC 619611, Set/2026: 0 violações no período, mas a coluna S diz "6x · RUN 26081801 · Erro Sistemático" (agosto). HGB Jan/2026: "6x · RUN 26020604" (fevereiro). No H3 vazio: "R_4s · RUN 25102801", de out/2025. Ver OC-HF4-5 |
| U-IMG-LJ | Inspeção das imagens do LJ | REPROVADO (Médio) | Linha suavizada, nenhum valor no eixo Y, linhas de ±k DP sem rótulo. No H3 vazio, os degraus viram curvas. É a OC-F4-8 da Bio, ver OC-HF4-8 |
| U-LOTE-USO | Lote em uso × lote corrente | REPROVADO (Médio) | C4 "Lançamentos no lote: 522611" (Configuração C20, validade 31/03/2026, vencida). O lote corrente é o 625311 desde 01/10/2026. Ver OC-HF4-4 |
| U-BOTOES | Botões da Estatística | Observação | Só a barra de navegação. Não há botão para recalcular a aba (ver OC-HF4-1) |
| U-TEMPOS | Tempo de resposta dos filtros | APROVADO | Todos ≤ 3 s. É mais rápido que a Bio, que não tem a troca de equipamento |

## 3. Ocorrências

### OC-HF4-1: A Estatística não recalcula quando o usuário troca o lote no Painel nem quando muda Ano De/Até; B4 mostra o lote novo com os números do lote antigo (Crítico; tipo `funcional`)

- **Tela/módulo:** aba Estatística (C:F leem `engEstat` = Eng_Saida 190:309, publicado por
  `mEstatistica.AtualizarEstatisticaAba`).
  - **Quem chama o recálculo:** só o `AtualizarEstatistica`. Quem chama esse, por sua vez, é o
    `mIntegracao.AtualizarDadosCore` (o ATUALIZAR) e o `EquipMudou`, que na Hema só roda quando o ajuste
    automático do ADR-077 muda o filtro.
  - **Troca de lote:** o `TrocarLoteAnalise` → `RecalcularAnalitoAtual` refaz só o motor do Painel.
  - **Evento da aba:** o módulo da planilha Estatística (Planilha10) está vazio, então mudar B3/D3 não dispara nada.
- **Ação:** Painel H3 = 619611 (antes 522611), depois a aba Estatística pelo botão, depois Ano De/Até = 2026/2026.
- **Esperado:** os textos da própria aba (B4 "(refletido do Painel)"; F4 "O lote é o escolhido no Painel (H3): média,
  DP e CV nunca misturam lotes"; "Ano De/Até filtram um intervalo") levam a esperar n/média/DP/CV do lote 619611 nos
  anos escolhidos. Ou então um aviso de "desatualizado".
- **Observado:**
  - **Lote trocado no Painel:** B4 = 619611, mas WBC N1 = n 50; média 3,0446; DP 0,054929; CV 1,8041 %. São
    exatamente os números do 522611 em 2025, e o 619611 não tem nenhum resultado em 2025.
  - **Ano De/Até mudado para 2026/2026:** nada muda. Deveria dar 54; 3,12056; 0,070988; CV 2,2749 %.
  - **Navegando pelo botão, saindo e voltando:** nada muda.
  - **Alcance:** vale para as 84 linhas.
  - **O cálculo em si está certo:** com o recálculo interno do ATUALIZAR (E-INTERNO), 0 divergências.
  - **No H3 vazio:** B4 vira o lote em uso (522611) e os números continuam os do último recálculo.
- **Dificuldade:** nada na aba avisa. O usuário vê o lote que escolheu escrito em B4 e números plausíveis.
- **Consequência:**
  - estatística mostrada sob o lote errado: associação errada entre lote e estatística;
  - o ET % (H = 1,65·F + |G|), a Margem, o Status CV (Q) e o "N (medições)" da linha usam o F do lote antigo;
  - o Painel não é afetado: B:F e H vêm do motor do Painel, e G e I (Bias do CEQ e Sigma do CVp) não dependem do lote;
  - o relatório tirado da Estatística depois de trocar o lote ou o ano sai com números de outro lote ou ano.
- **Sugestão (decisão do Orquestrador, do Especialista em CQ e do usuário):**
  - recalcular a Estatística no `TrocarLoteAnalise` e num `Worksheet_Change` de B3:D3;
  - ou, no mínimo, mostrar "DESATUALIZADO, use ATUALIZAR" quando B4/B3/D3 diferirem do que foi calculado. O
    Eng_Saida poderia guardar o lote e os anos do último cálculo, como o J7:J9 do Painel já faz com `engLote`.

  A Bioquímica tem outra estrutura (`mEstatPeriodo`, UDF e Planilha7), e lá a F4 não viu isso.
- **Evidência:** `f4h.json` (E-BASE, E-LOTE-619611 com `se_fosse_o_lote_anterior_522611`, E-ANO-2026-619611,
  E-INTERNO-619611 e os `Estat_B4`/`Estat_WBC` dos testes L-*); `f4h_conf.json` (`apos_H3`, `apos_IrEstatistica`,
  `apos_B3D3`, `apos_ir_e_voltar` e `vbcomp`, Planilha10 = 0 linhas). Código: `src/hema/mEstatistica.bas`
  (`AtualizarEstatistica`, `RecalcularAnalitoAtual`, `EquipMudou`) e `src/hema/mLotes.bas` (`TrocarLoteAnalise`).
- **Prioridade:** Crítico, porque é estatística mostrada sob o lote errado (REGRAS §4). O Orquestrador pode baixar
  para Alto se considerar que a rotina sempre começa com ATUALIZAR, mas a troca de lote e de ano é justamente o uso
  desta aba.

### OC-HF4-2: H3 vazio com período longo: n e contagens de Westgard são só das 180 corridas mais recentes, e o aviso fala só do gráfico (Médio; tipo `calculo`)

- **Tela/módulo:** Painel B7:B9 e M7:R9 no modo TODOS (`AtualizarCalc`, janela NK = 180 de chaves RUN|LOTE).
- **Ação:** WBC, H3 vazio, K4 = "Ano inteiro" 2026, e depois "Tudo (sem filtro)".
- **Observado:**
  - **"Ano inteiro" 2026:** n = 141 / 155 / 130, contra 337 / 340 / 308 resultados no ano (393 corridas RUN|LOTE).
    O Painel é idêntico ao recálculo só da janela.
  - **"Tudo":** 141 / 155 / 130 contra 478 / 485 / 447.
  - **O aviso do A39:** "⚠ 213 corrida(s) antigas do período fora do **gráfico** (máx. 180)".
- **Consequência:** o n e as violações "do período" cobrem só os últimos ~5 meses.
- **Por que Médio, e não Alto como a OC-F4-1 da Bio:**
  - no modo TODOS, média, DP e CV ficam vazios por decisão do ADR-086 (decisão 6: "no máximo as 180 corridas mais
    recentes", de propósito);
  - com lote escolhido, nenhum lote da Hema passa de 180 corridas.

  O defeito é o aviso dizer só "gráfico".
- **Sugestão:** escrever "fora do gráfico **e das contagens**" e mostrar o n total do período, ou calcular o n e as
  contagens sobre o período inteiro.
- **Evidência:** `f4h.json`, P-TODOS-Ano e P-TODOS-Tudo (`n_so_da_janela_180`, `engM1`, A39).

### OC-HF4-3: No H3 vazio, o Painel ainda mostra Bias % e Sigma (e o card "SIGMA DO PLANO"), contra a decisão 4 do ADR-086 (Médio; tipo `duvida_metodo`)

É a mesma OC-F4-3 da Bioquímica.

- **Observado:** WBC fev–mar/2026 com H3 vazio:
  - C:E e H vazios, e o `engPainel` vazio, como o ADR diz;
  - mas G = 1,86 % (Bias do CEQ) e I = 2,27 / 3,58 / 5,68, com "SIGMA DO PLANO 2,27 σ".
  - G e I leem a Estatística, que no H3 vazio está no lote em uso (522611) e, pela OC-HF4-1, possivelmente
    desatualizada. Bias e Sigma não dependem do lote, então os valores coincidem por construção.
- **Sugestão:** o Especialista decide se Bias e Sigma do analito aparecem no modo TODOS. Depois, o ADR ou o Painel
  se alinham.
- **Evidência:** `f4h.json` L-TODOS (`linha_painel`, `tela_G_H_I_vazia` = falso) e a imagem `Painel_todos_fevmar.png`.

### OC-HF4-4: "Lote em uso" 522611 (validade 31/03/2026) aparece em C4 e no A39 do H3 vazio, e o lote corrente é o 625311 (Médio; tipo `usabilidade`)

Repete a OC-F4-4 da Bioquímica.

- **Observado:**
  - **C4:** "Lançamentos no lote: 522611". `loteAtivo` = Configuração C20 = 522611, lote de out/2025 com validade
    31/03/2026.
  - **No H3 vazio:** o A39 diz "Lote 522611 · alvo N1 3,0038 ± 0,05290 ...", e esse lote não tem ponto em
    fev–mar/2026. A Estatística B4 vira 522611.
  - **Os lançamentos manuais e o filtro da rotina (ADR-078)** se referem a esse lote.
- **Consequência:** o usuário não sabe qual lote está vendo nem para qual lote vão os manuais. O alvo mostrado não é
  o de nenhum ponto.
- **Sugestão:**
  - **No H3 vazio:** listar no A39 os lotes do período (o Eng_Saida O1 já tem "533811 (19) · 536411 (59) · 602911 (23)").
  - **No laboratório:** atualizar o lote em uso da Hematologia na Configuração.
- **Evidência:** `f4h.json` L-TODOS (A39, C4, O1) e `lote_em_uso`.

### OC-HF4-5: "Últ. violação" e "Classificação" sob o título "violações no período" mostram uma violação fora do período, com o total 0 ao lado (Médio; tipo `usabilidade`)

- **Tela/módulo:** Painel S7:T9 (`engPainel` 19–20 = `AgregadoWestgard`, os eventos do lote inteiro em
  Eventos_Westgard).
- **Observado:**
  - **WBC 619611, Set/2026:** M:R = 0 nos três níveis, mas S = "6x · RUN 26081801" e T = "Erro Sistemático". O RUN
    é de agosto.
  - **HGB 533811, Jan/2026:** S = "6x · RUN 26020604", de fevereiro, depois do fim do período.
  - **No H3 vazio (fev–mar/2026):** S = "R_4s · RUN 25102801", do lote em uso de out/2025, que nem está na tela.
- **Consequência:** quem lê a linha entende que houve um erro sistemático no período filtrado.
- **Sugestão:** filtrar S/T pelo período, ou renomear para "Últ. violação do lote (histórico)". No H3 vazio, usar os
  lotes do período.
- **Evidência:** `f4h.json`, G-CONTAGEM-* (`ultima_violacao_S`) e L-TODOS. Imagens `Painel_set_619611.png` e
  `Painel_todos_fevmar.png`.

### OC-HF4-6: Ao voltar de outro lote, o período escolhido não volta quando o lote de partida não é o "lote em uso" (Médio; tipo `usabilidade`)

Repete a OC-F4-7 da Bioquímica.

- **Observado:**
  - **Partindo do 619611:** 619611 Set/2026 → 522611 (ajusta para "Tudo") → 619611 fica em "Tudo".
  - **Partindo do lote em uso:** 522611 Out/2025 → 619611 → 522611 devolve Out/2025.
  - É o ADR-078. Mas, como o lote em uso está desatualizado (OC-HF4-4), quem trabalha no lote corrente perde a
    escolha.
- **Sugestão:** guardar o filtro sempre que o ajuste automático o troca, ou resolver junto com a OC-HF4-4.
- **Evidência:** `f4h.json` L-AJUSTE-1 e L-AJUSTE-2.

### OC-HF4-7: Depois de digitar De/Até, Ano (K3) e Período (K4) continuam com o rótulo antigo (Médio; tipo `usabilidade`)

Repete a OC-F4-6 da Bioquímica.

- **Observado:** com G3:G4 = 01/02–31/03/2026, o Painel mostra "Ano 2025 · Período Tudo (sem filtro)".
- **Sugestão:** ao digitar G3/G4, limpar K3/K4 ou pôr "Personalizado".
- **Evidência:** `f4h.json` `K4_apos_digitar_G3G4` e a imagem `Painel_todos_fevmar.png`.

### OC-HF4-8: Levey-Jennings com linha suavizada, sem valores no eixo Y e sem rótulo nas linhas de ±k DP; no H3 vazio, os degraus de limite viram curvas (Médio; tipo `usabilidade`)

Repete a OC-F4-8 da Bioquímica nos 3 gráficos da Hema.

- **Observado:**
  - Em `Painel_set_619611.png`, a linha entre os pontos é curva e passa acima dos pontos (N1, corridas 47–50).
  - Não há números no eixo Y.
  - Em `Painel_todos_fevmar.png`, a troca 533811 → 536411 → 602911 aparece como ondas nas 7 linhas.
  - Os valores dos pontos e dos limites estão certos: G-PONTOS, G-LIMITES e G-TODOS-N* APROVADOS.
- **Sugestão:** linha reta, rótulos de ±k DP e valores no eixo Y.

### OC-HF4-9: Ano Até menor que Ano De é aceito sem aviso, e a aba calcula só o Ano De (Baixo; tipo `usabilidade`)

- **Observado:** B3 = 2026 e D3 = 2025. O cálculo usa só 2026, como manda o `EstatBasica`, mas a tela continua "Ano
  Até 2025", sem aviso.
- **Sugestão:** validação (D3 ≥ B3) ou um texto "usando só 2026".
- **Evidência:** `f4h.json` E-ANO-INVERTIDO e `estat_ano_invertido_tela`.

### Confirmação (sem ID novo): ET do Painel com bias "SEM EP" = 1,65 × CV (é a OC-H2 da F3)

- **RET% (619611, set/2026):** G7:G9 = "SEM EP", H7 = 5,8099 % = 1,65 × 3,52114, H8 = 5,0733 e H9 = 7,1078. A
  Estatística deixa H vazio quando não há CEQ.
- **RDW-SD:** H = 1,5478 / 1,6458 / 1,6779.
- **O que a F4 acrescenta:** a confirmação na tela, que a F3 tinha pedido.

## 4. Dúvidas de regra de negócio (ao Especialista em CQ, pelo Orquestrador)

1. **Estatística ao trocar lote ou ano:** a aba deve recalcular na hora, ou só no ATUALIZAR com um aviso de
   desatualizada (OC-HF4-1)?
2. **H3 vazio:** o Bias (CEQ) e o Sigma (CVp), que são do analito, devem aparecer no modo TODOS (OC-HF4-3)? E o n e
   as contagens devem ser do período inteiro ou da janela de 180 corridas (OC-HF4-2)?
3. **"Últ. violação":** deve ser do período filtrado ou do histórico do lote (OC-HF4-5)?
4. **Lote em uso:** o da Hematologia (Configuração C20) é 522611, com validade 31/03/2026. Quem atualiza esse campo, e
   ele deve acompanhar o lote corrente (OC-HF4-4)?
5. **Dados para o laboratório conferir (não são defeitos do sistema):**
   - **RBC 522611:** o alvo cadastrado N1 2,3066 ± 0,0219 fica cerca de 4,4 DP abaixo da média real do lote
     (2,4022), e 40 dos 50 pontos dão 1_3s.
   - **DP repetido entre lotes:** os lotes 614011 e 619611 de WBC têm DP idênticos no cadastro (0,07935 / 0,1368 /
     0,25395). Pode ser o DP de TESTE de 1,5 × o lote em uso (GUIA §4b), ainda não substituído.
   - **Lotes sobrepostos:** 528211 e 533811 rodaram juntos em 22–23/12/2025. O ADR-080 diz que "nunca há dois lotes
     no mesmo nível ao mesmo tempo". O motor tratou bem (L-SOBREPOSTOS), mas a premissa do ADR não vale nesta base.

## 5. Impressões de uso (Relatório A)

- **Filtros do Painel:** rápidos (até 3 s) e diretos. Na Hema há um equipamento só, então há menos passos que na
  Bio. O atalho Ano/Período preenche De/Até sem erro, e o ajuste automático ao trocar de lote avisa com clareza.
- **Confiança nos números:** ao longo de 31 períodos, 8 analitos e 6 cenários de LJ (1.261 pontos de Westgard),
  tudo o que o Painel mostra bate com o recálculo dos dados brutos. A tabela e o gráfico dizem a mesma coisa, e o
  H3 vazio separa corretamente os lotes paralelos.
- **O que mais confunde:**
  - **A aba Estatística:** parece seguir o lote do Painel (B4 muda), mas os números ficam parados até o próximo
    ATUALIZAR, e mudar o ano na própria aba não faz nada.
  - **O texto "Últ. violação · Erro Sistemático":** aparece ao lado de "0 violações no período".
  - **O lote em uso antigo (522611):** aparece em C4 e no A39.
- **H3 vazio:** útil para ver a troca de lote, com os degraus de limite certos. Mas, num período longo, o n é só das
  últimas 180 corridas, e o aviso fala só do gráfico.

## 6. Limpeza

- **Arquivos apagados:** a cópia `.xlsm` e o `cq.pkl`. Antes de apagar, a cópia tinha o hash da produção
  (`0b45fa270d2a`): nada foi salvo.
- **O que ficou em `%TEMP%\qcw\auditoria\2026-10-31_ciclo01\hema_f4\`:** só JSON, logs e imagens. Conferi que não
  há senha nem caminho de rede.
- **Excel:** criei e encerrei só os PIDs 3484, 23680 e 5428. Um EXCEL.EXE de outra origem, provavelmente a tarefa
  agendada, apareceu durante a 1ª sessão e não foi tocado. No fim não havia nenhum Excel aberto.
- **Produção:** continua com sha `0b45fa270d2a`.
