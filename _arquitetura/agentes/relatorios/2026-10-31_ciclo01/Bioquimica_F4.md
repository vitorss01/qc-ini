# Bioquímica, F4 (períodos, lotes, filtros, gráficos e usabilidade): ciclo 2026-10-31_ciclo01

Agente: Usuário Simulado. Modo: completo. Execução: 01/11/2026, das 02h22 às 02h36 (3º passe, que substitui o
relatório de 31/10 21h37). O 3º passe refez do zero os dois roteiros anteriores numa cópia nova da produção (mesmo hash)
e acrescentou um roteiro novo (`f4c_todos_periodos.py`) para o que ainda não tinha sido coberto.

Resumo:

- **Reprodução:** os 79 testes do roteiro principal e os 5 do complemento deram **o mesmo estado e os mesmos números**
  do 2º passe. A única diferença nos JSON está nos tempos de resposta.
- **Filtros de período e lote:** mês, trimestre, semestre, "Ano inteiro" e lote A → B → H3 vazio, em 3 analitos e nos
  2 equipamentos (Glicose DIM 1, Ureia DIM 2 e Hemoglobina glicada DIM 1, esta com 3 lotes em sequência), calcularam
  **exatamente** sobre os dados brutos filtrados. Foram 66 comparações de período: 64 com 0 divergências (diferença
  relativa máxima de 9,8e-16) e as 2 de "Tudo (sem filtro)" reprovadas (OC-F4-1). O motor (Eng_Saida) usa só os dados filtrados, e não apenas a tela.
- **Trimestre e semestre:** são calculados dos dados brutos do intervalo, e não como média dos meses. Conferido em
  todos os casos.
- **Levey-Jennings e Westgard:** batem ponto a ponto com o recálculo independente em 7 cenários (4 no modo de um lote e
  3 com H3 vazio). Pontos, eixo X (corrida do lote, ADR-080, e corrida atravessando os lotes, ADR-086), as 7 linhas de
  limite (inclusive os degraus por lote), as marcas por ponto e as contagens M:Q conferem. Foram 1.287 marcas de
  Westgard comparadas, com 0 divergências.
- **Reprovados:**
  - o período "Tudo (sem filtro)" num lote com mais de 180 corridas (OC-F4-1);
  - o lote de urina 5659 dentro da Glicose (OC-F4-2), que este passe mostrou ser **um padrão** (OC-F4-11): os lotes de
    urina 6353 e 6354, o lote de PCR 8353 e os lotes 8177 e 8794 também estão ativos dentro de analitos séricos. No H3
    vazio, isso destrói a escala do gráfico (Ureia N2: eixo Y até 1.932 com o alvo em 97,7);
  - a lista de lotes do H3, que não filtra pelo analito, e o Painel, que troca o analito sozinho (OC-F4-12);
  - a coluna "Últ. violação/Classificação", que não é do período (OC-F4-13);
  - o ponto de violação cortado na borda do LJ (OC-F4-14);
  - os achados de coerência e usabilidade já vistos (OC-F4-3 a OC-F4-10), que persistem.

## 1. Ambiente e evidência

| Item | Valor |
|---|---|
| Produção copiada | `QC_Bioquimica.xlsm`, sha256 curto `fe10dba1f1fd`, gravada em 30/10/2026 16:33. O hash é o mesmo antes e depois da cópia e no fim da frente: a produção não foi aberta |
| Cópia | `%TEMP%\qcw\auditoria\2026-10-31_ciclo01\f4r3\Bioquimica_F4.xlsm`. O hash continua `fe10dba1f1fd` depois das três sessões: nada foi salvo |
| Sessões do Excel (uma por vez, sem outro Excel aberto, cada uma encerrada no `finally` pelo próprio PID) | PID 2488: roteiro principal. PID 25940: complemento. PID 18528: roteiro novo do 3º passe |
| Base do cálculo independente | `cq.pkl`: a `tblCQ_Final` da mesma cópia lida por openpyxl, sem Excel, com 34.254 linhas (34.244 elegíveis). O teste T00 confere que os IDs são idênticos aos lidos pelo COM |
| Ferramentas | `_arquitetura/agentes/ferramentas/f4_periodos_lotes.py` (roteiro P/E/L/C/G), `f4b_complemento.py` ("Ano inteiro", "Tudo", 5659, células de filtro, Estatística, imagens) e **`f4c_todos_periodos.py`** (novo neste passe: períodos em Ureia DIM 2 e Hemoglobina glicada DIM 1; H3 vazio ponto a ponto com degraus e Westgard lote a lote; lista do H3) |
| Saídas | Em `%TEMP%\qcw\auditoria\2026-10-31_ciclo01\f4r3\`: `f4.json` e `f4.log`; `img\f4b.json` e `f4b.log`; `c\f4c.json` e `f4c.log`. Imagens exportadas pelo próprio Excel da cópia (PDF → PNG): `img\Painel_out_8977.png`, `img\Painel_t3_8977.png`, `img\Painel_todos_setout.png`, `img\Estatística.png`, `c\Painel_todos_GLI1-JUNAGO.png`, `c\Painel_todos_HBA1-JUN.png` e `c\Painel_todos_UR1-DEZ.png`. A tela do usuário não foi capturada |
| Reprodução do 2º passe | Comparação campo a campo de `f4r2\f4.json` × `f4r3\f4.json`: 79 testes com o mesmo estado e a mesma evidência. As únicas diferenças são `tempo_s`/`tempos_s` |
| Dados de teste | Nenhum. A F4 só mexe em filtros e não grava dados |

**Correções feitas na ferramenta no 2º passe** (conferidas no texto do `mEstatistica.AtualizarCalc` e mantidas):

1. `janela()`: com a janela de 180 corridas cheia, as corridas que só têm X saem. É a regra do D10/QA-ETL-001.
2. O Westgard independente roda no eixo **compactado** das corridas elegíveis (`pDeE`).
3. A evidência do Westgard serializa sem chaves de tupla.
4. Os rótulos reais do K4 são "Ano inteiro" e "Tudo (sem filtro)".
5. A aba Estatística usa MENSAL e TRIMESTRAL (não há lista em B4).

**Definição usada no H3 vazio (ADR-086)**, implementada no `f4c` e conferida no texto do ADR:

- chave da corrida = RUN|LOTE;
- janela = as 180 corridas elegíveis mais recentes até o fim do período, de todos os lotes juntos;
- X = posição 1, 2, 3... por nível em ordem cronológica, atravessando os lotes, no histórico inteiro;
- alvo de cada ponto = Média/DP do lote do ponto, lida do cadastro;
- Westgard: uma passada por lote, só com os pontos desse lote dentro da janela.

**Caminho do usuário:**

- As células de filtro do Painel (C3, L4, H3, K3, K4, G3 e G4) são **desbloqueadas** na aba protegida. Foram escritas
  pela própria célula, com os eventos ligados e o Painel protegido (`mSeguranca.ReprotectAll`). É o mesmo caminho do
  usuário.
- O Painel só foi desprotegido para o `PageSetup` da exportação das imagens, um passo técnico; depois foi protegido
  de novo.
- **Exceção:** na aba Estatística, B4, D4, F4 e H4 são **bloqueadas**, e os testes E-* escreveram nelas por automação,
  um passo que o usuário não consegue dar (ver OC-F4-5).

## 2. Testes executados

As seções 2.1 a 2.6 são do roteiro principal e do complemento, refeitos neste passe com resultado idêntico ao do
2º passe. Os números de cálculo são idênticos nos dois passes. A coluna "Tempo" da 2.1 é a do 2º passe; os tempos do
3º passe diferem em até 0,15 s e estão em U-TEMPOS. As seções 2.7 a 2.9 são novas.

Tolerância: 1e-9 relativo sobre o `Value2`. O cálculo independente usa `statistics` sobre as linhas brutas com
PARTICIPA_ESTATISTICA = SIM, filtradas por analito, equipamento, lote, nível e dia, com DP amostral (n−1). O alvo do
lote vem do cadastro (`mLotes.MediaDoLote`/`DPDoLote`).

### 2.1 Períodos (Glicose, DIMENSION 1)

Os períodos foram escolhidos pelo Painel: K3 (Ano) e K4 (Período), com o lote de cada mês em H3.

| ID | Período | Lote | Estado | Painel N1 / N2: n, média, DP, CV | Maior dif. relativa | Tempo |
|---|---|---|---|---|---|---|
| P-M202510 | out/2025 | 8974 | APROVADO | 24; 103,2083; 56,6307; 54,87 % / 22; 275,5455; 41,9918; 15,24 % | 2,6e-16 | 0,6 s |
| P-M202511 | nov/2025 | 8974 | APROVADO | 14; 104,2143; 55,2076 / 12; 290,4167; 4,0555 | 2,6e-16 | 0,5 s |
| P-M202512 | dez/2025 | 8974 | APROVADO | 15; 90,3333; 5,7776 / 13; 289,4615; 1,8980 | 1,7e-16 | 0,6 s |
| P-M202601 | jan/2026 | 8974 | APROVADO | 17; 87,7647; 0,7524 / 16; 281,8125; 4,7218 | 1,5e-16 | 0,6 s |
| P-M202602 | fev/2026 | 8974 | APROVADO | 17; 87,4706; 1,0676 / 16; 282,6875; 2,5747 | 2,4e-16 | 0,6 s |
| P-M202603 | mar/2026 | 8974 | APROVADO | 25; 88,2800; 1,4866 / 20; 282,0500; 3,4864 | 1,8e-16 | 0,6 s |
| P-M202604 | abr/2026 | 8974 | APROVADO | 20; 98,6500; 44,1055 / 17; 281,5882; 2,8736 | 4,4e-16 | 0,6 s |
| P-M202605 | mai/2026 | 8974 | APROVADO | 21; 86,8571; 1,3522 / 24; 279,0833; 4,2006 | 2,1e-16 | 0,6 s |
| P-M202606 | jun/2026 | 8974 | APROVADO | 18; 85,3333; 1,1376 / 18; 278,9444; 3,0577 | 0 | 0,6 s |
| P-M202607 | jul/2026 | 8977 | APROVADO | 12; 74,9167; 0,7930 / 13; 271,5385; 3,7995 | 0 | 0,6 s |
| P-M202608 | ago/2026 | 8977 | APROVADO | 20; 78,5000; 1,6059 / 19; 277,7368; 4,3696 | 2,0e-16 | 0,6 s |
| P-M202609 | set/2026 | 8977 | APROVADO | 23; 77,3478; 1,7218 / 24; 270,9167; 8,8804 | 0 | 0,6 s |
| P-M202610 | out/2026 | 8977 | APROVADO | 21; 77,4286; 1,1650 / 20; 243,3500; 83,3952; 34,27 % | 3,8e-16 | 0,6 s |
| P-T20254 | T4/2025 | 8974 | APROVADO | 53; 99,8302; 47,1796 / 47; 283,1915; 29,3692 | 1,7e-16 | 0,7 s |
| P-T20261 | T1/2026 | 8974 | APROVADO | 59; 87,8983; 1,2275; CV 1,3965 % / 52; 282,1731; 3,6284; CV 1,2859 % | 1,7e-16 | 0,6 s |
| P-T20262 | T2/2026 | 8974 | APROVADO | 59; 90,3898; 25,9662 / 59; 279,7627; 3,6594 | 5,1e-16 | 0,6 s |
| P-T20263 | T3/2026 | 8977 | APROVADO | 55; 77,2364; 2,0089; CV 2,6010 % / 56; 273,3750; 7,2377; CV 2,6475 % | 6,8e-16 | 0,6 s |
| P-T20264 | T4/2026 | 8977 | APROVADO | igual a out/2026: só há outubro | 3,8e-16 | 0,6 s |
| P-LS1 | S1/2026 | 8974 | APROVADO | 118; 89,1441; 18,3453 / 111; 280,8919; 3,8242 | 9,8e-16 | 0,7 s |
| P-LAno inteiro | 2026 | 8974 | APROVADO | 121; 88,9835; 18,1475 / 114; 280,7807; 3,8553 | 3,2e-16 | 0,6 s |
| P-LTudo (sem filtro) | todo o histórico | 8974 | **REPROVADO** | Painel 140; 89,1500; 16,9680; 19,03 % / 131; 281,9847; 4,9225. Independente, todos os dados: **174; 92,2874; 30,3737; 32,91 % / 161; 281,4845; 16,1149**. O Painel é idêntico ao recálculo **só da janela das 180 corridas** (M1 = 39 corridas cortadas). Ver OC-F4-1 | 0,44 no DP | 0,7 s |
| G3/G4 | todos os meses e trimestres | — | APROVADO | O atalho Ano/Período preencheu De/Até com o 1º e o último dia certos em todos os casos | — | — |

**O trimestre é calculado dos dados brutos, e não como média dos meses.** Conferido em todos os trimestres, porque o
Painel bate com os dados brutos do trimestre. A "média dos meses" daria outro número:

- **T1/2026, N1:** média das médias mensais 87,8384 e média dos CV mensais 1,2539 %, contra o Painel com 87,8983 e
  1,3965 %;
- **T3/2026, N1:** 76,9215 e 1,7767 %, contra 77,2364 e 2,6010 %.

O Painel mostra, corretamente, o número dos dados brutos. Nenhum ADR manda outra coisa.

### 2.2 Aba Estatística (62 linhas, analito × nível, DIMENSION 1)

| ID | Visão | Lote | Estado | Evidência |
|---|---|---|---|---|
| E-MENS-202603-8974 | MENSAL, mar/2026 | 8974 | APROVADO (cálculo) | Janela H5:J5 01/03–31/03/2026. 232 comparações de n, média, DP e CV contra os dados brutos, 0 divergências, máx. 1,3e-12. Os valores vêm de UDF com cache de somas, daí o 1e-12, que fica abaixo de 1e-9 |
| E-MENS-202609-8977 | MENSAL, set/2026 | 8977 | APROVADO (cálculo) | 232 comparações, 0 divergências, máx. 1,1e-12 |
| E-TRIM-202601-8974 | TRIMESTRAL, T1/2026 | 8974 | APROVADO (cálculo) | Janela 01/01–31/03/2026. 240 comparações, 0 divergências, máx. 1,2e-11 |
| E-TRIM-202603-8977 | TRIMESTRAL, T3/2026 | 8977 | APROVADO (cálculo) | 232 comparações, 0 divergências, máx. 1,8e-12 |
| E2-USUARIO | Visão pela célula, como o usuário | — | **REPROVADO** | B4, D4, F4 e H4 estão **bloqueadas** (aba protegida), e só a visão PERSONALIZADO, refletida do Painel G3/G4, é alcançável. D4 = `=Painel!I3`, que está vazio, e a aba mostra "Ano 0". Forçando MENSAL com D4 refletido, a janela fica vazia e a aba mostra o **histórico inteiro do lote** (Glicose 8977 N1 n = 76 em vez de 23), sem aviso. Ver OC-F4-5 |
| E-PERS | PERSONALIZADO, refletindo set/2026 do Painel | 8977 | APROVADO | Glicose N1 23; 77,34783; 1,72175 e N2 24; 270,91667; 8,88045, iguais ao Painel e ao recálculo independente |

### 2.3 Lotes: A → B → H3 vazio

Glicose, DIMENSION 1, De/Até 01/06–31/08/2026, um período que cruza a troca de lote.

| ID | Teste | Estado | Evidência |
|---|---|---|---|
| L-8974 | Lote 8974: estatística só do lote | APROVADO | N1 21; 84,9524; 1,7457. N2 21; 278,6190; 3,0899. Máx. 4,3e-16 |
| L-8974-MOTOR | Eng_Saida só com pontos do lote, analito, equipamento, nível e período, e com todos eles | APROVADO | 42 pontos, todos do 8974. 0 fora do escopo, 0 faltando |
| L-8977 | Lote 8977 | APROVADO | N1 32; 77,1562; 2,2159. N2 32; 275,2188; 5,1226 |
| L-8977-MOTOR | Motor só do lote 8977 | APROVADO | 64 pontos, 0 fora, 0 faltando |
| L-TODOS | H3 vazio (ADR-086) | APROVADO no motor | engTodos = VERDADEIRO. H4 "H3 vazio = TODOS os lotes do período...". n = 54 e 53, iguais ao independente. Pontos: 8974 = 42, 8977 = 64 e 8353 = 1. Média/DP/CV/ET vazios no `engPainel`. **Mas** o Painel continua mostrando Bias 2,73 % e Sigma 3,19 / 0,55 (OC-F4-3), e o A39 fala do lote 8974 (OC-F4-4) |
| L-TODOS-ALVO | Cada ponto com a Média/DP do seu lote (Eng_Saida BW..) | APROVADO | 0 divergências contra o cadastro |
| B-TODOS-SETOUT | H3 vazio, 01/09–31/10/2026 | APROVADO no motor; achado de dado | n = 60 e 56, iguais ao independente. Pontos por lote: **8977 = 44 + 44 e 5659 = 16 + 12**, todos com MATRIZ = SORO. O lote 5659 é o controle de urina (OC-F4-2) |
| B-5659 | Lote 5659 em H3 | APROVADO no cálculo | N1 16; 55,9375; DP 9,1686; CV 16,39 %. N2 12; 33,25; 0,7538. Iguais ao independente. Alvo do cadastro N1 59,67 ± 1,5428 e N2 33,0 ± 5,0270, os DP idênticos aos do 8977 |
| L-AJUSTE | Trocar para um lote sem ponto no filtro (8974 em out/2026) | APROVADO (ajuste) | Período ajustado para "Tudo (sem filtro)", com aviso em H4 e A39 ("o filtro anterior não mostrava nenhum resultado dele -- ajustado..."). Tempo 2,7 s. Ao **voltar ao 8977**, o "Out" escolhido **não volta** (OC-F4-7) |
| B-XTODOS | Eixo X no H3 vazio atravessando os lotes (ADR-086) | APROVADO (pela imagem e pela contagem) | Imagem `Painel_todos_setout.png`: N1 vai até a corrida 267 e N2 até 249. O recálculo independente das chaves RUN\|LOTE ativas de todo o histórico deu 267 e 249, com a 1ª corrida do período em 208 (N1) e 194 (N2) |

### 2.4 Outras combinações (lote 8977, set/2026)

| ID | Analito / equipamento | Estado | Evidência |
|---|---|---|---|
| C-Glicos-2 (+MOTOR) | Glicose / DIMENSION 2 | APROVADO | N1 21; 77,2381; 1,2611. N2 20; 273,1; 6,7270. Motor 0 fora e 0 faltando |
| C-Ureia-2 (+MOTOR) | Ureia / DIMENSION 2 | APROVADO | N1 23; 32,6087; 1,5880. N2 22; 96,0; 3,7670 |
| C-Proteí-1 (+MOTOR) | Proteína total / DIMENSION 1 | APROVADO | N1 24; 6,6292; 0,6224. N2 25; 4,6040; 0,0676 |
| C-Creati-1 (+MOTOR) | Creatinina / DIMENSION 1 | APROVADO | N1 25; 2,7292; 0,1445. N2 28; 6,2586; 1,4423 (CV 23 %, dado) |
| C-Colest-2 (+MOTOR) | Colesterol total / DIMENSION 2 | APROVADO | N1 21; 238,5714; 1,9893. N2 20; 93,25; 1,1642 |
| C-HBA1C | HbA1c / DIMENSION 2, analito sem CEQ: o ET do Painel | **REPROVADO (confirma OC-7 da F2 e OC-2 da F3)** | Painel G8 = "SEM EP", mas H8 = **2,7740 %** = 1,65 × CV 1,68119. O `engPainel` ainda calcula Sigma 4,7585 com bias 0, que fica oculto. Fórmula H7: `=IF($E7="","",$E7*1.65+ABS(IF(ISNUMBER($G7),$G7,0)))`. Esta é a confirmação na tela que a F3 deixou para a F4 |

### 2.5 Levey-Jennings e Westgard (4 cenários)

| Cenário | EIXO (motor = janela independente) | PONTOS N1 / N2 (gráfico = ativos do período; X = corrida do lote) | LIMITES (7 linhas = Média + k·DP do cadastro) | Tabela × gráfico | WESTGARD por ponto | Contagens do Painel M:Q e status |
|---|---|---|---|---|---|---|
| Glicose DIM 1, 8977, out/2026 | APROVADO (102 = 102) | APROVADO: 21 = 21, X 56..76; 20 = 20, X 57..76, sem buraco | APROVADO: N1 79,3140 ± 1,5428; N2 274,7759 ± 5,0270 | APROVADO | APROVADO: 152 pontos, 0 divergências | APROVADO: N1 [0,2,0,0,0], N2 [2,4,0,0,0], REJEITADO/REJEITADO |
| Glicose DIM 1, 8977, T3/2026 | APROVADO (76) | APROVADO: 55 = 55, X 1..55; 56 = 56, X 1..56 (lote novo começa em 1) | APROVADO | APROVADO | APROVADO: 111, 0 divergências | APROVADO: N1 [6,18,0,6,1], N2 [1,5,0,6,1] |
| Ureia DIM 2, 8977, set/2026 | APROVADO (93) | APROVADO: 23, X 49..71; 22, X 43..64 (as sequências por nível diferem, como manda o ADR-080) | APROVADO: N1 33,1929 ± 1,7094; N2 96,3481 ± 4,8 | APROVADO | APROVADO: 135, 0 divergências | APROVADO: [0,0,0,1,2] nos dois níveis |
| Glicose DIM 1, 8974, mar/2026 | APROVADO (137) | APROVADO: 25, X 88..112; 20, X 80..99 | APROVADO: N1 87,1 ± 1,0285; N2 285,6364 ± 3,3513 | APROVADO | APROVADO: 211, 0 divergências | APROVADO: N1 [1,3,0,0,0], N2 [1,0,0,0,0] |

**Definição usada no Westgard independente**, conferida em `AvaliarWestgard` e nos detectores:

- 1_3s: |z| > 3;
- 2_2s WITHIN_RUN: os dois níveis do RUN;
- 2_2s ACROSS: duas observações ativas seguidas do nível (ADR-083);
- R_4s: |z1 − z2| > 4 no RUN;
- 4_1s N2_R2 e 8x N2_R4: corridas completas seguidas, no eixo compactado;
- os complementares SAME_LEVEL não marcam (`DetectorAtivo` = 0);
- a 5ª coluna do bloco ("R10x" no comentário do código) é o 8x N2_R4.

### 2.6 Usabilidade (imagens exportadas da cópia e `auditoria_abas.auditar_aba`)

| ID | Teste | Estado | Evidência |
|---|---|---|---|
| U-TEMPOS | Tempo de resposta dos filtros (3º passe) | APROVADO | C3: 0,14–0,21 s (a 1ª troca 1,84 s). H3: 1,2–3,5 s; H3 vazio 2,3 s. K3/K4: 0,4–0,85 s. L4 (equipamento): **5,2–6,1 s**, o mais lento. Estatística: 2,3–2,8 s por célula |
| U-AUDIT-PAINEL | Textos e formas do Painel | REPROVADO (Baixo) | J7/J8 "REJEITADO" cortado: 48 pt de texto em 47 pt de coluna (OC-F4-9) |
| U-AUDIT-ESTAT | Textos e formas da Estatística | APROVADO | 0 achados. Na imagem, porém, o nome do analito aparece desalinhado da linha (OC-F4-10) |
| U-IMG-LJ | Inspeção das imagens do LJ | REPROVADO (Médio) | Linha suavizada que passa do valor real; nenhum valor no eixo Y nem rótulo nas linhas de DP. No H3 vazio, os degraus de limite viram ondas (OC-F4-8) |
| U-K4 | Rótulo do período depois de digitar De/Até | REPROVADO (Médio) | K4 continua "S1" com De/Até 01/09–31/10/2026 (OC-F4-6) |

### 2.7 Períodos em outro analito e no outro equipamento (3º passe, `f4c`)

Mesmo método da 2.1: K3/K4 pelo Painel, com o lote majoritário de cada mês em H3. Na Hemoglobina glicada, o lote
minoritário dos meses de troca também foi testado. Tolerância 1e-9 relativo; todos os testes deram G3/G4 certos.

**Ureia / DIMENSION 2 (lotes 8974 → 8977):** 20 testes, todos APROVADOS, com diferença relativa máxima de 5,6e-16.

| Período | Lote | Painel N1: n; média; DP | Painel N2: n; média; DP |
|---|---|---|---|
| out/2025 a jun/2026 (9 meses) | 8974 | ex.: out/2025 25; 32,36; 2,0591. mar/2026 20; 31,25; **7,9596** | out/2025 24; 95,4583; 4,0430. mar/2026 23; 92,1304; **20,4635** |
| jul/2026 a out/2026 (4 meses) | 8977 | ex.: jul/2026 25; 34,56; **12,7380**. set/2026 23; 32,6087; 1,5880 | jul/2026 19; 94,8421; 2,3396. out/2026 21; 93,0; **11,5931** |
| T4/2025, T1, T2 e S1/2026 | 8974 | T1: 56; 32,6786; 4,9730. S1: 123; 32,7236; 3,4884 | T1: 63; 94,0476; 12,4712. S1: 134; 93,4030; 10,1449 |
| T3, T4 e S2/2026 | 8977 | T3: 71; 33,3803; 7,6108. S2: 89; 33,0; 6,9101 | T3: 64; 95,8594; 2,9540. S2: 85; 95,1529; 6,3312 |

Os DP em negrito vêm dos valores implausíveis já registrados na F1 (OC-1 da F1). O cálculo está certo; o dado é que
infla o DP.

**Hemoglobina glicada / DIMENSION 1 (8588 → 8589 → 8590, com meses de sobreposição):** 22 testes, todos APROVADOS,
com diferença relativa máxima de 4,4e-16.

| Período | Lote | N1: n; média; DP | N2: n; média; DP | Observação |
|---|---|---|---|---|
| dez/2025 | 8589 | 8; 5,1675; 0,1219 | 9; 9,1922; 1,5596 | mês com dois lotes |
| dez/2025 | 8588 | 8; 4,9525; 0,1427 | 8; 9,8425; 0,1074 | o lote minoritário só com os seus pontos |
| jun/2026 | 8589 | 23; 5,1278; 0,1230 | 27; 9,7485; 0,2403 | mês com dois lotes |
| jun/2026 | 8590 | 1; 5,08; — | 1; 9,68; — | n = 1: DP vazio, como deve ser |
| T2/2026 | 8589 | 74; 5,2088; 0,5600 | 76; 9,7457; 0,2040 | |
| T3/2026 | 8590 | 76; 5,1451; 0,1433 | 78; 9,8347; 0,2260 | |
| os demais 11 meses, T4/2025, T1/2026, T4/2026, S1 e S2 | majoritário | — | — | todos APROVADOS |

**Trimestre ≠ média dos meses:** a média das médias mensais dá outro número, e o Painel mostra o dos dados brutos. Por
exemplo:

- Ureia DIM 2, T1/2026, N1: Painel 32,6786 contra a média das médias mensais 32,75;
- Hemoglobina glicada, T1/2026, N2: 9,6548 contra 9,7044.

### 2.8 H3 vazio, ponto a ponto (3º passe, `f4c`)

Em cada cenário, De/Até foram digitados em G3/G4 e H3 foi apagado, pelas células, como o usuário faz.

| Cenário | Escopo e n | Pontos e X (gráfico = independente) | Degraus (7 linhas × ponto = Média + k·DP do lote do ponto) | Westgard por ponto (lote a lote na janela RUN\|LOTE) | Westgard = visão do lote | Contagens M:Q |
|---|---|---|---|---|---|---|
| **GLI1-JUNAGO:** Glicose DIM 1, 01/06–31/08/2026 (8974 → 8977, mais 1 ponto do 8353) | APROVADO: n 54/53 = independente. Média/DP/CV/ET vazios; engTodos = VERDADEIRO. 0 fora do escopo, 0 faltando | APROVADO: N1 54 pontos, X 154..207; N2 53 pontos, X 141..193 | APROVADO: 378 + 371 conferências, 0 divergências | APROVADO: 265 marcas, 0 divergências. Janela de 180 corridas no motor = 180 no independente | APROVADO: 0 divergências, nem na borda | APROVADO: N1 [10, 26, 0, 5, 2] e N2 [2, 8, 0, 5, 2] |
| **HBA1-JUN:** Hemoglobina glicada DIM 1, jun/2026 (8589 e 8590 no mesmo mês) | APROVADO: n 24/28; O1 = "8589 (32) · 8590 (1)" | APROVADO: N1 24, X 173..196; N2 28, X 168..195 | APROVADO: 168 + 196, 0 divergências | APROVADO: 288 marcas, 0 divergências | APROVADO | APROVADO: N1 0; N2 [1, 0, 0, 0, 0] (1_3s em BIO-300301 = 10,53, z = +4,1 no 8589) |
| **UR1-DEZ:** Ureia DIM 1, dez/2025 (8974, mais 2 linhas do lote de urina 6353) | APROVADO no motor: n 21/17 = independente, **com** as linhas 871 e 1830 do 6353 | APROVADO: N1 21, X 46..66; N2 17, X 43..59 | APROVADO: 147 + 119, 0 divergências. O alvo do 6353 é N1 871 ± 1,7094 e N2 1830 ± 4,8 (ver OC-F4-11) | APROVADO: 125 marcas, 0 divergências | APROVADO | APROVADO: N1 [4, 5, 1, 3, 0] e N2 [3, 7, 1, 3, 0] |

**O cálculo do modo TODOS está correto e segue o ADR-086 à risca.** O que falha é o que entra nele (OC-F4-11) e o que
o cerca na tela (OC-F4-3, OC-F4-4, OC-F4-13 e OC-F4-14).

**Achados visuais nas imagens** (`c\Painel_todos_*.png`):

- **UR1-DEZ:** os eixos Y vão de −12,6 a 918,4 (N1) e de 0,3 a 1.932,2 (N2). Os pontos do 8974 ficam numa linha
  achatada no pé do gráfico, e um pico enorme marca o ponto de urina. O gráfico fica inútil (OC-F4-11).
- **HBA1-JUN:** o y_max do N2 é 10,53, exatamente o valor do ponto de violação BIO-300301. O ponto vermelho fica cortado
  na borda de cima, sobre o texto "Duplo clique num ponto..." (OC-F4-14). A coluna "Últ. violação" diz
  "1_3s · RUN 25102201", de 22/10/2025, e não a violação de junho (OC-F4-13).
- **GLI1-JUNAGO:** o K4 continua "Jun", com De/Até 01/06–31/08 (OC-F4-6). O ponto −1 do lote 8353 (corrida 202) puxa o
  eixo Y do N1 até −10. A linha suavizada passa por baixo do eixo em UR1-DEZ e faz as bordas dos degraus parecerem
  ondas (OC-F4-8).

### 2.9 Lista de lotes do H3 (3º passe, `f4c`)

| ID | Teste | Estado | Evidência |
|---|---|---|---|
| H-LISTA | A lista do H3 mostra só os lotes do analito e do equipamento? | **REPROVADO** | A validação é `=lstLotes`, com 22 lotes, a mesma para todos os analitos. Com Glicose / DIM 1, só 4 deles têm resultado (5659, 8353, 8974 e 8977). Os outros 18 (8588, 8589 e 8590 da Hemoglobina glicada; 8352, 8353 e 8356 da PCR; 6353 e 6354 de urina etc.) aparecem sem distinção. Ver OC-F4-12 |
| H-OUTRO | Glicose + H3 = 8588 (lote da Hemoglobina glicada) | APROVADO (avisa), com achado | O Painel **troca o analito** para Hemoglobina glicada e o período para "Tudo (sem filtro)". O H4 diz "ver o aviso abaixo dos gráficos", e o A39 explica: "o filtro anterior não mostrava nenhum resultado dele -- ajustado: analito Hemoglobina glicada; período Tudo". Os números mostrados (8588: N1 51; 4,9780; 0,1397) são os certos do 8588. O usuário que queria Glicose passa a ver outro analito (OC-F4-12) |

## 3. Ocorrências

### OC-F4-1: "Tudo (sem filtro)" num lote com mais de 180 corridas: n, média, DP e CV do Painel são só das 180 corridas mais recentes, e o aviso fala só do gráfico (Alto; tipo `calculo`)

- **Tela/módulo:** Painel, indicadores por nível (`mEstatistica.AtualizarCalc`, janela NK = 180, e `AtualizarPainelEng`).
- **Ação:** Glicose, DIMENSION 1, H3 = 8974, K4 = "Tudo (sem filtro)". O mesmo acontece quando o Painel ajusta o
  período sozinho para "Tudo" ao trocar de lote (L-AJUSTE).
- **Esperado:** a estatística do período escolhido, isto é, todo o histórico ativo do lote, ou então um aviso claro de
  que ela também foi cortada.
- **Observado:**
  - **Painel:** N1 n = 140, média 89,1500, DP 16,9680, CV 19,03 %. N2 n = 131, média 281,9847, DP 4,9225, CV 1,75 %.
  - **Todos os dados brutos do lote:** N1 174; 92,2874; 30,3737; 32,91 %. N2 161; 281,4845; 16,1149; 5,72 %.
  - **De onde vem o número do Painel:** é idêntico ao recálculo das 180 corridas elegíveis mais recentes. Ficam fora
    39 corridas (Eng_Saida M1 = 39).
  - **O aviso do A39:** "⚠ 39 corrida(s) antigas do período fora do **gráfico** (máx. 180)".
- **Dificuldade:** o usuário lê "fora do gráfico" e entende que a tabela cobre o período inteiro.
- **Consequência:** n, média, DP e CV de "Tudo" não são do lote inteiro. Em lote longo, o CV mostrado é outro (aqui,
  19 % contra 33 % no N1).
- **Sugestão:** ao Especialista em CQ e ao usuário:
  - calcular a estatística do Painel sobre todo o período, independente da janela do gráfico;
  - ou, se a janela for decisão de projeto (desempenho), dizer no aviso "fora do gráfico **e da estatística**" e mostrar
    o n total.
- **Evidência:** `f4.json`, teste P-LTudo (sem filtro); `f4b.json`, A-Tudo. Os dois reproduzem bit a bit.
- **Prioridade:** Alto, e não Crítico, porque o corte é proposital (ADR-049/086, "a mesma janela de sempre") e há aviso;
  o defeito é o aviso dizer só "gráfico". Se o Especialista entender que a estatística tem de ser do período, sobe para
  Crítico.

### OC-F4-2: Controle de urina (lote 5659) entra como Glicose sérica, ganha bloco de lote e se mistura com o 8977 no H3 vazio (Alto; tipo `dados`)

- **Tela/módulo:** Painel (H3 vazio e H3 = 5659), recebimento e cadastro automático de lote, Configuração.
- **Ação:** Glicose, DIMENSION 1, De/Até 01/09–31/10/2026: primeiro com H3 vazio, depois com H3 = 5659.
- **Esperado:** controle de urina fora da Glicose sérica.
- **Observado:**
  - **Origem dos dados:** na `tblCQ_Final`, o lote 5659 tem UCFP e GLUC de urina (MATRIZ = URINA, GLUC sem cadastro).
    Também tem **28 linhas GLUC com MATRIZ = SORO**, mapeadas para "Glicose" e ATIVAS: N1 16, média 55,9; N2 12,
    média 33,3. Os valores típicos de controle de urina (33 e 60 mg/dL) contrastam com 77 e 274 do soro 8977.
  - **Cadastro automático:** a Configuração registrou "Lote 22 · 5659 · automático em 04/10/2026 · 6 resultado(s) de
    24/09 a 28/09/2026 · DIMENSION 1".
  - **Alvo de TESTE:** N1 59,67 ± 1,5428 e N2 33,0 ± 5,0270. Os DP são idênticos aos do 8977, ou seja, foram copiados.
    O CV do alvo do N2 fica em 15 %, e o Westgard quase nunca dispara.
  - **No H3 vazio:** o Painel junta 8977 (44 + 44) e 5659 (16 + 12). O LJ (imagem `Painel_todos_setout.png`) alterna
    os dois materiais, com as linhas de limite em degrau oscilando entre 60 e 78 (N1) e entre 33 e 275 (N2).
  - **Níveis trocados:** em 27/10/2026, BIO-369955 = 32 e BIO-369956 = 33 entraram como **N1** do 5659, valores de N2.
    Por isso o N1 do 5659 tem DP 9,17 e status REJEITADO.
- **Consequência:**
  - **No H3 vazio:** o n e as contagens de Westgard ficam inflados (n N1 60 em vez de 44; violações N1 18).
  - **No gráfico:** fica ilegível.
  - **No lote 5659:** a estatística dele é de um material de outra matriz, sob o nome Glicose.
  - **O que não é afetado:** a estatística do 8977 com o lote escolhido em H3.
- **Sugestão:**
  - **No laboratório:** conferir no SIPEC o material do lote 5659 (MATRIZ enviada = SORO) e inativar ou corrigir as 28
    linhas.
  - **Ao Especialista:** avaliar uma trava de plausibilidade por matriz e lote, por exemplo um lote com resultado de
    matriz URINA não virar lote de analito sérico.
  - **Na tela:** avisar no Painel quando o H3 vazio junta lotes de níveis muito diferentes.
- **Evidência:** `img\f4b.json`, testes B-TODOS-SETOUT e B-5659 e `config_5659`. IDs: BIO-355958, BIO-355959,
  BIO-356592, BIO-356596, BIO-369955, BIO-369956 e BIO-370948/370952 (os últimos, 28/10). Script: `f4b_complemento.py`.
- **Prioridade:** Alto. Pela tabela do §4 seria Crítico, por ser associação errada de lote e analito, mas a origem é o
  dado de entrada, e a estatística do lote sérico escolhido não é afetada. Fica para o Orquestrador decidir.

### OC-F4-3: No H3 vazio, o Painel ainda mostra Bias % e Sigma, contra a decisão 4 do ADR-086, e o ET fica vazio (Médio; tipo `duvida_metodo`)

- **Tela/módulo:** Painel G7:I8 (fórmulas que leem a aba Estatística, e não o `engPainel`).
- **Ação:** apagar H3 (Glicose, DIMENSION 1, jun–ago e set–out/2026).
- **Esperado (ADR-086, decisão 4):** "Média/DP/CV/Bias/ET/Sigma do período: ficam vazios no modo TODOS".
- **Observado:** média, DP, CV e ET vazios. Mas Bias = 2,73 % (o bias do CEQ, que não depende do lote) e Sigma = 3,19 /
  **0,55**, com o card "SIGMA DO PLANO 0,55 σ" em destaque. O `engPainel` está vazio, como o ADR diz. O que aparece vem
  de `G7 = INDEX(Estatística!G...)` e `I7 = INDEX(Estatística!L...)`.
- **Consequência:** a linha fica incoerente: Sigma sem ET. E o ADR e o comportamento divergem.
- **Sugestão:** o Especialista decide se Bias e Sigma, que são do analito (CEQ e CV pooled), devem aparecer no modo
  TODOS. Depois, o ADR ou o Painel se alinham.
- **Evidência:** `f4.json` L-TODOS (`linha_painel`) e `img\f4b.json` B-TODOS-SETOUT. Imagem `Painel_todos_setout.png`.

### OC-F4-4: O Painel mostra dois lotes diferentes: "Lançamentos no lote: 8974" (lote em uso) junto de "Lote 8977 · lote em análise"; no H3 vazio, o A39 dá o alvo do 8974 (Médio; tipo `usabilidade`)

- **Tela/módulo:** Painel C4 (texto sob o analito), H3/H4 e A39.
- **Observado:**
  - **Com H3 = 8977:** sob o analito aparece "Lançamentos no lote: 8974".
  - **No H3 vazio:** o A39 diz "Lote 8974 · alvo N1 87,1 ± 1,03 · N2 285,6 ± 3,35. Filtro 01/09/26 a 31/10/26",
    e o 8974 não tem nenhum ponto no período. O Eng_Saida E1 (lote em uso) = 8974, e a Glicose do DIMENSION 1 roda no
    8977 desde 07/07/2026.
- **Consequência:** o usuário não sabe qual lote está vendo nem qual recebe os lançamentos manuais. O alvo mostrado no
  A39 não é o de nenhum ponto da tela.
- **Sugestão:**
  - **No H3 vazio:** o A39 deve listar os lotes do período (o Eng_Saida O1 já tem "8977 (55) · 5659 (18)").
  - **Rótulo:** trocar "Lançamentos no lote" por "Lote em uso (digitação)".
  - **Laboratório:** conferir se o lote em uso da Bioquímica ainda deve ser o 8974 (ver Dúvidas).
- **Evidência:** `img\Painel_out_8977.png` e `img\f4b.json` (`todos_setout_estado`: E1 = 8974, O1).

### OC-F4-5: Na Estatística, as visões MENSAL e TRIMESTRAL são inalcançáveis e quebradas: B4/F4 bloqueadas, "Ano 0" refletido de Painel!I3 vazio, e a janela vazia cai no histórico inteiro sem aviso (Médio; tipo `funcional`)

- **Tela/módulo:** Estatística B4:J5 (`JanelaInicio`/`JanelaFim`, `EstatPeriodo`).
- **Ação:** tentar ver a Estatística por mês ou trimestre, como sugerem os campos "Visão", "Ano" e "parte".
- **Esperado:** escolher a visão ou, se ela não for para o usuário, não mostrar campos mortos.
- **Observado:**
  - **Células bloqueadas:** B4 (Visão), D4, F4 e H4 estão bloqueadas na aba protegida. Só vale o PERSONALIZADO,
    replicado do Painel.
  - **Ano 0:** D4 = `=Painel!I3`, e o I3 do Painel está vazio. A aba mostra "Ano **0** (refletido do Painel)".
  - **Visão forçada:** com MENSAL e D4 refletido, H5/J5 ficam vazios, e a aba calcula sobre **todo o histórico do lote**:
    Glicose 8977 N1 n = 76 em vez dos 23 de setembro. O "Período efetivo" não avisa.
  - **O cálculo em si está certo:** com ano e parte válidos (E-MENS/E-TRIM), 0 divergências.
- **Consequência:** confusão ("Ano 0"). Um administrador que destrave a célula ganha números do histórico inteiro com
  cara de mensais.
- **Sugestão:** apontar D4 para Painel!K3 (o ano atual do Painel), ou retirar os campos Visão/Ano/parte, deixando claro
  que a aba segue o período do Painel.
- **Evidência:** `img\f4b.json` (`celulas_filtro` e `estatistica_caminho_usuario`), `f4.json` (`est_cabecalho`,
  `est_linha14`) e a imagem `img\Estatística.png`.

### OC-F4-6: Depois de digitar De/Até, o campo Período (K4) continua com o rótulo antigo (Médio; tipo `usabilidade`)

- **Tela/módulo:** Painel K3/K4 × G3/G4.
- **Ação:** com K4 = "S1", digitar De 01/09/2026 e Até 31/10/2026.
- **Observado:** K4 continua "S1" e G3:G4 = 01/09–31/10. A tela diz "S1" e mostra setembro e outubro.
- **Consequência:** o usuário pode imprimir ou ler o período errado.
- **Sugestão:** ao digitar G3/G4, limpar K4 ou pôr "Personalizado".
- **Evidência:** `img\f4b.json` B-5659 (`K4: "S1"`, G3/G4) e a imagem `Painel_todos_setout.png`.

### OC-F4-7: Ao voltar de outro lote, o período escolhido não volta quando o lote de partida não é o "lote em uso" (Médio; tipo `usabilidade`)

- **Tela/módulo:** Painel H3 (ADR-077/078; o filtro da rotina fica no Eng_Saida AA1).
- **Ação:** H3 = 8977, K4 = Out → H3 = 8974. O Painel ajusta para "Tudo", com aviso. Depois, H3 = 8977 de novo.
- **Esperado pelo usuário (GUIA §4b: "Ao voltar ao lote em uso, o filtro de antes volta sozinho"):** voltar a "Out".
- **Observado:** fica em "Tudo (sem filtro)". Pelo ADR-078, o filtro só é guardado quando o lote que sai é o lote em
  uso, que aqui é o 8974. Portanto, o comportamento segue o ADR. Mas, para quem trabalha no lote corrente 8977, que
  não é o "em uso", a escolha se perde.
- **Sugestão:** guardar o filtro sempre que o ajuste automático o troca, qualquer que seja o lote de partida. Ou ligar
  isto à revisão do lote em uso (OC-F4-4).
- **Evidência:** `f4.json` L-AJUSTE (antes, depois_8974, volta_8977).

### OC-F4-8: Levey-Jennings com linha suavizada que ultrapassa os pontos, sem valores no eixo Y e sem rótulo nas linhas de DP; no H3 vazio, os degraus viram ondas (Médio; tipo `usabilidade`)

- **Tela/módulo:** gráficos do Painel (série "Resultado" e linhas de limite).
- **Observado:**
  - **Linha suavizada:** em `Painel_out_8977.png`, no N2, a linha entre a corrida 73 (OK) e a 74 sobe **acima** dos
    dois pontos. A suavização inventa valores entre as corridas.
  - **Eixo Y e linhas de DP:** os dois gráficos não mostram números no eixo Y (o eixo tem mínimo e máximo, mas sem
    rótulos), e as 7 linhas tracejadas não dizem qual é ±1, ±2 ou ±3 DP. A legenda não as lista.
  - **No H3 vazio** (`Painel_todos_setout.png`): os limites em degrau, suavizados, viram senoides. O ADR-086 só previa
    "segmento inclinado".
  - **Efeito da OC-1 da F1:** os valores −1 e 1 achatam o N2 de out/2026, e as 7 linhas ficam coladas no topo.
- **Consequência:** a leitura visual do LJ fica pior, e não dá para dizer pelo gráfico a quantos DP o ponto está.
- **Sugestão:** linha reta entre os pontos, sem suavização; rótulos de ±k DP na ponta direita das linhas; valores no
  eixo Y.
- **Evidência:** as imagens em `%TEMP%\qcw\auditoria\2026-10-31_ciclo01\f4r3\img\` e, para o H3 vazio, em `f4r3\c\`. Os números dos pontos e dos
  limites estão certos (G-PONTOS e G-LIMITES APROVADOS).

### OC-F4-9: "REJEITADO" cortado em J7:J8 do Painel (Baixo; tipo `usabilidade`)

- **Observado:** `auditar_aba` mediu 48 pt de texto numa coluna de 47 pt (Segoe UI).
- **Sugestão:** alargar J em 1–2 pt ou usar "reduzir para caber".
- **Evidência:** `img\f4b.json`, `achados_painel`.

### OC-F4-10: Na Estatística, o nome do analito fica no pé da célula e os números no meio, e o nome parece pertencer à linha de baixo (Baixo; tipo `usabilidade`)

- **Observado:** em `img\Estatística.png`, nas linhas mais altas (Lactato N2, Ácido úrico N2, Cálcio), o nome aparece
  colado à linha seguinte.
- **Sugestão:** alinhar o nome ao centro na vertical.
- **Evidência:** imagem `Estatística.png`. A auditoria automática não acusa porque não mede o alinhamento vertical.

### Confirmação (sem ID novo): ET do Painel com bias "SEM EP" = 1,65 × CV (é a OC-7 da F2 e a OC-2 da F3)

- **Caso:** HbA1c, DIMENSION 2, lote 8977, set/2026.
- **Observado:** Painel G8 = "SEM EP", H8 = 2,7740 % = 1,65 × 1,68119. A Estatística deixa H vazio.
- **O que a F4 acrescenta:** a confirmação na tela, que a F3 tinha pedido.

### OC-F4-11: Lotes que não são do analito (urina 6353/6354, PCR 8353, 8177, 8794 e o 8974 na Hemoglobina glicada) estão ativos dentro de analitos séricos, ganham alvo copiado e, no H3 vazio, destroem o gráfico (Alto; tipo `dados`)

- **Tela/módulo:** recebimento e cadastro automático de lote (`tblCQ_Final`, LotesStore) e Painel com H3 vazio.
- **Ação:** Ureia, DIMENSION 1, K3 = 2025, K4 = Dez, H3 apagado (cenário UR1-DEZ). Depois, varredura da `tblCQ_Final`
  da cópia por lote × analito.
- **Esperado:** o modo TODOS junta só os lotes de controle daquele analito sérico.
- **Observado:**
  - **Linhas de lotes de urina:** todas com MATRIZ = SORO, origem INTERFACEAMENTO e PARTICIPA_ESTATISTICA = SIM,
    mapeadas para analitos séricos. Os valores são de controle de urina:
    - Ureia DIM 1, lote 6353, em 28/12/2025: BIO-230952 = 871 e BIO-230958 = 1830 (no soro: 33 e 95);
    - Creatinina DIM 2, lote 6354: N2 cerca de 236 (BIO-353874, BIO-354443), contra 6,4 no soro;
    - Sódio, Potássio, Cloro, Fósforo e Cálcio nos lotes 6353 e 6354, em dez/2025, ago/2026 e set/2026. Exemplos:
      Sódio N1 = 57 em BIO-225460 e Potássio N2 = 115 em BIO-225468.
  - **Lote 8353 (PCR):** um ponto em quase todos os analitos em 24/08/2026 08:38, com valores absurdos: Glicose −1
    (BIO-340790), Ureia −1 (BIO-340789), Creatinina −0,32 (BIO-340793) e Potássio 0 (BIO-340766).
  - **Lotes 8177 e 8794:** um ponto cada, em 16/09/2026 e 01/12/2025, no DIMENSION 2 de vários analitos.
  - **Lote 8974 na Hemoglobina glicada:** é o lote da química geral, e tem 14 linhas ativas de Hemoglobina glicada. A
    BIO-210299 (= 4,85) entrou como N2.
  - **Alvo criado com o próprio valor:** o cadastro automático deu ao 6353 da Ureia N1 = 871 ± 1,7094 e N2 = 1830 ± 4,8.
    A média é o único resultado, e o DP é idêntico ao de outro lote. Por isso z = 0 e o Westgard nunca dispara.
  - **No Painel (H3 vazio):** o n inclui as linhas de urina (N1 21 = 20 do 8974 + 1 do 6353). Os eixos Y vão até 918
    (N1) e 1.932 (N2), e os pontos do lote real ficam achatados numa linha. No caso da Glicose jun–ago, o ponto −1 do
    8353 puxa o eixo do N1 até −10.
  - **O que fica fora:** a lista do H3 (OC-F4-12) oferece todos esses lotes ao usuário. Com o lote escolhido em H3, a
    estatística do lote sérico **não** é afetada (2.1, 2.7). O CVp do Sigma também os deixa fora (F3: "6354 n = 5 e
    8177 n = 1 fora").
- **Dificuldade:** o usuário não tem como saber, pelo Painel, que o pico é de outro material. A ficha do ponto mostra o
  lote, mas não a matriz enviada.
- **Consequência:**
  - **Modo TODOS:** inutilizável nesses períodos;
  - **Lotes alheios:** ganham "estatística" e alvo falsos, com nome de analito sérico;
  - **Contagens:** o n e as contagens do período misturam materiais.
- **Sugestão:**
  - **No laboratório:** conferir no sistema de origem a configuração dos lotes 6353/6354 (urina), 8353 (PCR), 8177,
    8794 e 8974 (Hemoglobina glicada), e inativar as linhas.
  - **Ao Especialista e ao Orquestrador:** uma trava de plausibilidade no recebimento. Por exemplo: um lote de material
    de urina não vira lote de analito sérico, um lote com um único resultado não ganha alvo automático e um resultado
    muito longe do alvo dos lotes do analito é sinalizado.
  - **No modo TODOS:** avisar quando o período junta lotes com alvos muito diferentes.
- **Evidência:**
  - **Varredura offline:** a `cq.pkl` da cópia (sha `fe10dba1f1fd`), com os IDs citados.
  - **No Excel:** `c\f4c.json`, testes W-UR1-DEZ-* (`pontos_por_lote_nivel_matriz`, `alvos`, `eixo`) e
    W-GLI1-JUNAGO-*.
  - **Imagens:** `c\Painel_todos_UR1-DEZ.png` e `c\Painel_todos_GLI1-JUNAGO.png`.
- **Prioridade:** Alto, pelo mesmo raciocínio da OC-F4-2, da qual esta é a generalização. Pela tabela do §4 seria
  Crítico, por ser associação errada de lote e analito. Mas a origem é o dado de entrada, e a estatística com o lote
  escolhido não é afetada. Fica para o Orquestrador decidir.

### OC-F4-12: A lista do H3 tem os 22 lotes de todos os analitos, e escolher um lote de outro analito troca o analito do Painel sozinho (Médio; tipo `usabilidade`)

- **Tela/módulo:** Painel H3 (validação `=lstLotes`) e o ajuste automático do ADR-077/078.
- **Ação:** Glicose, DIMENSION 1, set/2026, H3 = 8977. Abrir a lista do H3 e escolher 8588.
- **Esperado:** a lista oferece os lotes do analito e do equipamento selecionados, ou ao menos os destaca. Escolher um
  lote sem resultado não muda o analito sem confirmação.
- **Observado:**
  - **A lista:** traz 22 lotes na ordem do cadastro (8974, 8973, 8588, 8352, 5352, 8353, ...). Só 4 têm resultado de
    Glicose DIM 1.
  - **Ao escolher 8588:** C3 passa a "Hemoglobina glicada" e K4 a "Tudo (sem filtro)". A explicação fica no A39,
    **abaixo dos gráficos**; o H4 diz "ver o aviso abaixo dos gráficos".
- **Dificuldade:** o usuário não sabe qual lote é de qual analito. Um clique errado muda o analito e o período, e o
  aviso fica fora da área visível.
- **Consequência:** risco de ler a tela de um analito achando que é outro, e mais cliques para voltar (analito, lote e
  período).
- **Sugestão:**
  - **Na lista:** filtrar o H3 pelos lotes com resultado do analito e equipamento (o Eng_Saida O1 já sabe quais são),
    ou mostrar "lote · analito · período".
  - **No ajuste:** quando o lote for de outro analito, perguntar antes de trocar o analito, ou mostrar o aviso no topo.
- **Evidência:** `c\f4c.json`, testes H-LISTA e H-OUTRO (antes e depois de C3, K4, H4 e A39).

### OC-F4-13: Na tabela "WESTGARD — violações no período", as colunas "Últ. violação" e "Classificação" não são do período; no H3 vazio, são de um lote que não está na tela (Médio; tipo `usabilidade`)

- **Tela/módulo:** Painel S7:T8, que vem de `engPainel` col. 19/20 = `AgregadoWestgard(analito, nível)`
  (`mEstatistica.bas`, perto da linha 2034). Esse agregado lê a aba Eventos_Westgard, que tem o **histórico inteiro**
  do lote do Painel, sem janela nem período. No modo TODOS, esse lote é o lote em uso.
- **Ação:** qualquer período que não seja o mais recente.
- **Esperado:** a última violação **do período**, coerente com as contagens M:Q ao lado e com o título da tabela.
- **Observado:**

  | Caso | Período | S7 / S8 mostrado | De quando é |
  |---|---|---|---|
  | Glicose DIM 1, lote 8974 | mar/2026 | "2_2s · RUN 26081201" / "2_2s · RUN 26070601" | ago e jul/2026 |
  | Glicose DIM 1, lote 8977 | T3/2026 (jul–set) | "2_2s · RUN 26102802" / "2_2s · RUN 26102806" | out/2026 |
  | Hemoglobina glicada DIM 1, H3 vazio | jun/2026 | N2 "1_3s · RUN 25102201", "Erro Aleatório" | 22/10/2025, lote 8974, que não tem ponto na tela |
  | Ureia DIM 1, H3 vazio | dez/2025 | "1_3s · RUN 26070601" / "8x · RUN 26052801" | jul e mai/2026 |

  No caso da Hemoglobina glicada, a violação real de junho (1_3s em BIO-300301, contada em M8) não é a que aparece em
  S8.
- **Dificuldade:** o título diz "violações no período", e a coluna ao lado das contagens parece explicá-las.
- **Consequência:** o usuário investiga a corrida errada, ou registra a classificação (sistemático/aleatório) de outro
  momento ou de outro lote.
- **Sugestão:** calcular "Últ. violação/Classificação" sobre os pontos do período mostrado, como as contagens M:Q. Se o
  agregado do lote for proposital, mudar o rótulo para "Últ. violação do lote (histórico)" e, no modo TODOS, deixar
  vazio.
- **Evidência:** `f4.json`, testes G-CONTAGEM-* (`ultima_violacao_S7_S8` e `periodo`); `c\f4c.json`, W-*-ESCOPO
  (`linha_painel`); imagem `c\Painel_todos_HBA1-JUN.png`.

### OC-F4-14: O ponto de violação fica cortado na borda superior do LJ, porque o máximo do eixo Y é igual ao valor do ponto (Baixo; tipo `usabilidade`)

- **Tela/módulo:** gráfico N2 do Painel no modo TODOS (moldura do eixo Y pelos degraus, ADR-086 §3).
- **Ação:** Hemoglobina glicada, DIMENSION 1, jun/2026, H3 vazio.
- **Observado:** y_max = 10,53, exatamente o valor de BIO-300301, o único 1_3s do período. Os +3s do 8589 e do 8590
  ficam em 10,315 e 10,302. O marcador vermelho sai metade para fora da área do gráfico e fica sobre o texto "Duplo
  clique num ponto...".
- **Consequência:** o ponto mais importante do gráfico é o que pior se vê, e fica mais difícil de dar o duplo clique
  nele.
- **Sugestão:** deixar uma margem (por exemplo 5 % da amplitude) além do maior e do menor ponto ao montar a moldura.
- **Evidência:** `c\f4c.json`, W-HBA1-JUN-PONTOS-N2 (`eixo.y_max` = 10,53); imagem `c\Painel_todos_HBA1-JUN.png`.

### Persistência das ocorrências do 2º passe

Como a produção é a mesma, todas as ocorrências OC-F4-1 a OC-F4-10 se reproduziram, com os mesmos números. O 3º passe
acrescentou estas evidências:

- **OC-F4-3:** também em Ureia DIM 1 dez/2025 (Bias 5,91 %, Sigma 0,25 / 0,19, "SIGMA DO PLANO 0,19 σ") e Glicose
  jun–ago.
- **OC-F4-4:** no H3 vazio da Hemoglobina glicada, o A39 diz "Lote 8974 · alvo N1 5,07 ± 0,097", de um lote sem ponto
  no período. O C4 diz "Lançamentos no lote: 8974" também para a Hemoglobina glicada, que roda no 8590.
- **OC-F4-6:** K4 "Jun" com De/Até 01/06–31/08/2026.

## 4. Dúvidas de regra de negócio (ao Especialista em CQ, pelo Orquestrador)

1. **Janela de 180 corridas:** a estatística do Painel para períodos longos ("Tudo", e "Ano inteiro" em lotes maiores)
   deve ser da janela de 180 corridas ou do período inteiro? Hoje é da janela (OC-F4-1).
2. **H3 vazio:** o Bias (CEQ) e o Sigma (CV pooled), que são do analito e não do lote, devem aparecer no modo TODOS
   (OC-F4-3)?
3. **Lote em uso:** o da Bioquímica é 8974 (Eng_Saida E1 e "Lançamentos no lote"), mas a Glicose roda no 8977 desde
   julho. O lote em uso é global ou por analito, e quem o atualiza? Isso decide para onde vão os lançamentos manuais e
   o filtro da rotina (OC-F4-4 e OC-F4-7).
4. **Filtro de período:** o motor filtra por **corrida**, pela data da 1ª linha dela (`PassaFiltro(dts)`), e não por
   resultado. Nesta base não houve divergência, porque nenhuma corrida atravessou a virada de mês. Fica registrado
   como definição a confirmar.
5. **Plausibilidade:** cabe uma trava por matriz e lote, para que o controle de urina não vire lote de analito sérico
   (OC-F4-2)?

6. **Lotes de outro material no analito sérico:** o recebimento deve aceitar como ativo um resultado com lote de
   urina (6353, 6354, 5659) ou de outro analito (8353 da PCR, 8974 na Hemoglobina glicada), e criar alvo com um
   único ponto (OC-F4-11)?
7. **"Últ. violação":** deve ser do período mostrado, como as contagens M:Q, ou do histórico do lote? E no H3 vazio
   (OC-F4-13)?
8. **Lista do H3:** deve ser filtrada pelo analito e equipamento? E o ajuste automático pode trocar o analito sem
   confirmação (OC-F4-12)?

## 5. Impressões de uso (Relatório A)

- **Filtros do Painel:** são rápidos e diretos (lista de analito, lote, ano e período pelo nome, "T1", "Jan"). O
  atalho Ano/Período preenche De/Até sem erro. Trocar o equipamento é o passo mais lento (5–7 s).
- **Confiança nos números:** em 22 períodos, 7 combinações e 4 cenários de LJ, tudo o que o Painel e a Estatística
  mostram bate com o recálculo dos dados brutos. A tabela e o gráfico dizem a mesma coisa.
- **H3 vazio:** é útil para ver a troca de lote. Mas, com dois materiais no período (o caso do 5659), o gráfico fica
  ilegível, e o A39 fala de um terceiro lote.
- **O que mais atrapalha a leitura:** os valores implausíveis (−1, 1 e os N1 de 100+ em out/nov/2025 e abr/2026 do
  8974; OC-1 da F1). Eles achatam os gráficos e inflam DP e CV (N1 do 8974: CV 54,9 % em out/2025, 44,7 % em abr/2026).
- **Estatística:** os campos Visão/Ano/parte parecem editáveis e não são. "Ano 0" passa a impressão de erro.
- **Aviso de corte:** o "⚠ 39 corrida(s) antigas do período fora do gráfico" passa despercebido e não diz que a
  estatística também foi cortada.
- **3º passe, outros analitos:** com Ureia no DIMENSION 2 e com a Hemoglobina glicada (3 lotes, meses com dois
  lotes), os filtros se comportam igual à Glicose, e os números conferem. Escolher o lote minoritário do mês mostra só
  os pontos dele.
- **H3 vazio:** o cálculo é exato e o gráfico em degrau ajuda a ver a troca de lote. Mas basta uma linha de outro
  material no período (urina, PCR) para o gráfico virar uma linha achatada com um pico. A tabela continua mostrando
  Bias e Sigma, e a "Últ. violação" é de outro lote e de outra data.
- **Lista de lotes:** 22 lotes sem dizer de qual analito. Um clique errado troca o analito, e o aviso fica abaixo
  dos gráficos, fora da vista.

## 6. Limpeza

Nenhum dado gravado. A cópia `f4r3\Bioquimica_F4.xlsm` ficou com o hash da produção (`fe10dba1f1fd`): nada foi
salvo. A produção manteve o mesmo hash do início ao fim.

Cada Excel foi encerrado pelo próprio PID (2488, 25940 e 18528), um por vez. Antes de cada sessão, conferi que não
havia outro Excel aberto, e nenhum outro processo do Excel foi tocado. As saídas e as imagens ficam em `%TEMP%`, fora
do repositório.
