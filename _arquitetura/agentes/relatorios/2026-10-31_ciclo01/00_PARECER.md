# Parecer consolidado: ciclo 2026-10-31_ciclo01

Agente: Usuário Simulado (parecer consolidado). Data: 03/11/2026. Modo: consolidação. O Excel não foi aberto neste passo.
Tudo aqui vem dos 8 relatórios das frentes desta pasta, de `frentes.json` e de `ocorrencias_todas.json`. Nenhum número
foi recalculado agora: os valores são os que as frentes mediram e registraram.

| Item | Valor |
|---|---|
| Produção auditada | `QC_Bioquimica.xlsm` sha256 curto `fe10dba1f1fd`; `QC_Hematologia.xlsm` sha256 curto `0b45fa270d2a`. Os hashes são os mesmos antes e depois em todas as frentes: a produção não foi aberta para escrita |
| Código-fonte | `src` sem mudança desde o commit `0a25f5a` (F3 dos dois produtos) |
| Frentes | F1 entrada de CIQ manual e automática; F2 CEQ; F3 auditoria matemática independente; F4 períodos, lotes, filtros, gráficos e usabilidade. Cada uma nos dois produtos |
| Relatórios de origem | `Bioquimica_F1.md` a `Bioquimica_F4.md`, `Hematologia_F1.md` a `Hematologia_F4.md` |
| Dados de teste | Só nas cópias, marcados `TESTE-AUDITORIA`, e as cópias foram apagadas no fim de cada frente |
| Ferramentas reutilizáveis | `_arquitetura/agentes/ferramentas/calc_independente.py`, `conferencia_cruzada_hema.py`, `segunda_opiniao_hema.py`, `f4_hema.py`, `f4_hema_conf.py`, `f4_periodos_lotes.py` |

**Resumo em números**

| | Bioquímica | Hematologia | Total |
|---|---|---|---|
| Testes registrados (entradas; vários agregam dezenas de subtestes) | 125 | 120 | 245 |
| Aprovados / reprovados / parciais / não executados | 85 / 32 / 5 / 3 | 87 / 30 / 0 / 3 | 172 / 62 / 5 / 6 |
| Cálculos no Relatório B (aprovados / reprovados) | 88 (83 / 5) | 86 (78 / 8) | 174 (161 / 13) |
| Ocorrências brutas (frentes) | 41 | 44 | 85 |
| Ocorrências deduplicadas (P-xx) | 38 | 42 | 80 |
| Críticas / Altas / Médias / Baixas (deduplicadas) | 2 / 10 / 19 / 7 | 2 / 8 / 21 / 11 | 4 / 18 / 40 / 18 |
| Dúvidas de regra de negócio levantadas | 24 | 25 | 49 |

Os 5 "parciais" são da Bio F2 (T17, T26, T27, T35, T36). Eles foram aprovados no cálculo ou no isolamento e reprovados
no fluxo ou no ET do Painel. Os 13 cálculos reprovados não são erro de fórmula. Em todos eles o QC_INI calcula certo
sobre a população errada ou mostra um número que não corresponde ao filtro da tela (ver B.3).

---

## Relatório A: experiência do usuário

### A.1 Funcionalidades testadas

| Frente | Bioquímica | Hematologia |
|---|---|---|
| F1: CIQ manual e automático | 33 testes (25 A, 8 R). Tela Digitar Resultados, 16 validações (vazio, texto, ponto/vírgula, data futura e inválida, nível, hora, inexistentes, duplicidade, minúsculas, equipamento vazio), SUBSTITUI_ID, recebimento DB_SEAC (+331), Sync_Log, idempotência, histórico, plausibilidade, equivalência manual × automático | 37 testes (30 A, 7 R). Os mesmos fluxos, mais 32 casos de validação de dados × digitação, espaço sobrando e manual implausível; recebimento +2.387 |
| F2: CEQ | 43 testes em 3 passes (25 A, 12 R, 5 parciais, 1 NE). Glicose, Bilirrubina direta (VB) e Proteína total (2 lotes): janela `frmCEQ`, 12 validações, duplicidade, inclusão, correção, Consolidar, isolamento por provedor, ano, lote e equipamento, permissão, trilha, primeiro CEQ de um analito novo, reversão | 26 testes (14 A, 11 R, 1 NE). RBC (CLIA) e LYMPH# (VB): os mesmos fluxos, mais nome do analito vizinho, provedor global R4 e provedor na coluna A da tabela do CAP |
| F3: matemática independente | 25 testes (18 A, 5 R, 2 NE). 62 linhas × 19 colunas (1.178 comparações) e segunda opinião do zero em 6 analitos (96/96) | 28 testes (23 A, 4 R, 1 NE). 84 linhas × 27 colunas, 45 células do Painel, segunda e terceira implementação (15 níveis; 459/459), contador U, blocos de lote (2.268), alvos do lote |
| F4: períodos, lotes, filtros, gráficos | 24 entradas (17 A, 7 R) que agregam mais de 150 testes: 3 analitos, 2 equipamentos, 66 períodos, lote A → B → H3 vazio, 7 cenários de LJ/Westgard (1.287 marcas), lista do H3, tempos, imagens exportadas | 29 testes (20 A, 8 R, 1 NE por não se aplicar). 31 períodos, 8 analitos, 6 cenários de LJ/Westgard N=3 (1.261 pontos), lotes paralelos, Estatística por ano, tempos, imagens exportadas |

### A.2 Facilidade de uso

- **Entrada manual (F1): fácil no caso comum, nos dois produtos.** O ID e o carimbo nascem sozinhos, a vírgula
  decimal funciona, o ponto é recusado com mensagem clara e as listas guiam o preenchimento. O SUBSTITUI_ID é o
  recurso mais amigável do sistema: preenche a chave inteira a partir do número do ID. As mensagens do QA são
  claras (DETALHE com o campo, A01, A06, A10, I04, E04).
- **Automático (F1): transparente para o usuário.** Um clique no ATUALIZAR traz tudo. Nada se perde nem se duplica, o
  ID é preservado e o histórico sobrevive à janela móvel do DB_SEAC. O manual e o automático dão a mesma estatística
  bit a bit.
- **Janela do CEQ (F2): boa.** Segue a ordem do relatório do CAP, traz o ano e a unidade sozinha, valida 12 casos com
  o campo certo e acha duplicidade mesmo com caixa e espaços diferentes. O que se acabou de registrar aparece na
  lista da rodada.
- **Painel (F4): rápido e direto.** Analito em 0,1–0,2 s, período em 0,1–0,8 s, lote em 1–3,5 s. Só o equipamento da
  Bioquímica é lento (5–6 s). O H3 vazio mostra os lotes em degrau e o Westgard lote a lote, ambos exatos.

### A.3 Dificuldades encontradas

1. **Erros silenciosos na entrada.** São os pontos de maior risco para o usuário, porque nada avisa:
   - equipamento ou analito em minúsculas troca um resultado real no Painel (P-01, P-03);
   - data colada inválida rola para outro mês (P-05, P-16);
   - equipamento vazio herda o do Painel (P-06);
   - valores colados fora do cadastro ficam ATIVO e invisíveis (P-23, P-42);
   - resultado −1/0 entra ATIVO (P-02, P-15).
2. **CEQ nas bordas, também em silêncio:**
   - nome fora do mapa, Controllab e analito novo são gravados "OK" e nunca entram no bias (P-09, P-19);
   - o nome do analito vizinho cai no bias errado (P-20);
   - registro sem alvo é aceito (P-25, P-45).
3. **Corrigir um CEQ é impossível para o ANALISTA.** A janela manda editar uma linha bloqueada (P-07, P-17). Para o
   ADM, a correção só vale depois do "Consolidar", e o indicador diz "base em dia" (P-08, P-18).
4. **A aba Estatística não acompanha a tela.**
   - Na Hematologia, mostra o lote novo com os números do anterior até o ATUALIZAR (P-04).
   - Na Bioquímica, no H3 vazio, calcula um lote que ninguém escolheu (P-28).
   - Na Bioquímica, as visões MENSAL e TRIMESTRAL estão inalcançáveis (P-37).
5. **Resultados sem sentido entram no Sigma.** Na Bioquímica, 253 resultados suspeitos ativos (P-02) fazem quase tudo
   aparecer como "Desempenho inadequado", e a tela não aponta qual resultado causou isso.
6. **Espera sem retorno.** O ATUALIZAR leva 15–25 s (Hema: 66 s com 2.387 novos) e cada gravação do CEQ leva 15–24 s,
   sem indicador de progresso. O usuário só vê o resultado do manual depois do ATUALIZAR.

### A.4 Navegação

- O caminho principal está bem resolvido: Digitar Resultados → ATUALIZAR → Painel. Os filtros ficam numa linha só
  (C3, H3, K3/K4, G3/G4, L4).
- A lista do H3 tem os 22 lotes de todos os analitos. Escolher um lote de outro analito troca o analito sozinho, e o
  aviso fica abaixo dos gráficos (P-40).
- Ao voltar de outro lote, o período escolhido se perde quando o lote de partida não é o "lote em uso" (P-30, P-60).
- O CEQ tem dois caminhos: a janela e as abas de dados com o botão Consolidar. O GUIA ensina só o caminho antigo, que
  está bloqueado para o ANALISTA (P-29, P-49).
- Na Hematologia, entrar na Estatística pelo botão não recalcula nada (P-04).

### A.5 Apresentação

- **Números certos com rótulos que confundem:**
  - "Lançamentos no lote: 8974/522611" mostra o lote em uso, vencido na Hema, em telas de outro lote (P-36, P-58);
  - "Últ. violação" de fora do período aparece ao lado de 0 violações (P-41, P-59);
  - o Status fica REJEITADO por uma regra antiga (P-44);
  - o K3/K4 mantém o rótulo antigo depois de digitar De/Até (P-38, P-61);
  - Bias e Sigma aparecem no modo TODOS (P-35, P-57).
- **Indicadores que se contradizem:**
  - ET "dentro do orçamento" ao lado de Sigma "inadequado", porque cada um usa um CV diferente;
  - três ou quatro Sigmas para o mesmo analito: Estatística, bloco do lote e Eng_Saida (P-33, P-50, P-76);
  - ET do Painel sem bias quando o analito não tem CEQ (P-10, P-21).
- **Gráfico:** o LJ tem linha suavizada que passa dos pontos, não tem valores no eixo Y e as linhas de DP não têm
  rótulo. No H3 vazio, os degraus viram ondas (P-39, P-62), e o ponto de violação fica cortado na borda (P-69).
  Os dados por trás estão certos.
- **Acabamento:**
  - coluna AT ilegível ("004 / 000") (P-31, P-51);
  - "REJEITADO" cortado (P-67);
  - nome do analito desalinhado (P-68);
  - "0" nas colunas RUN/DETALHE e nas linhas vazias (P-63, P-72, P-78);
  - contagens com decimal (P-79);
  - textos sem acento nas mensagens do QA.

### A.6 Sugestões do usuário (para o Orquestrador priorizar; nenhuma foi implementada)

1. **Normalizar a grafia** de EQUIPAMENTO e ANALITO do manual para a do cadastro no código comum. Uma correção cobre
   P-01 e P-03. O motor nunca deve sobrescrever um valor no mesmo RUN e nível sem alerta.
2. **Triagem de plausibilidade no QA** (ALERTA com o ID) para valor ≤ 0, nível trocado e |z| muito alto, tanto no
   automático quanto no manual (P-02, P-15). O laboratório revisa os 56 + 253 IDs já identificados.
3. **Recusar no ATUALIZAR** o que a validação do Excel já recusa quando digitado: data inválida, analito, lote ou
   equipamento fora do cadastro. Tornar o EQUIPAMENTO obrigatório na Bioquímica (P-05, P-06, P-16, P-23, P-42).
4. **CEQ:**
   - editar e excluir pela própria janela, com papel e trilha;
   - consolidar sozinho depois da edição;
   - avisar quando o nome não está no mapa, quando não há alvo e quando a unidade não bate com a do analito;
   - permitir incluir um nome no mapa W:Y pela tela (P-07 a P-09, P-17 a P-20, P-25 a P-27, P-45 a P-47).
5. **Estatística:** recalcular ou avisar "desatualizada" ao trocar o lote ou o ano (P-04), e seguir o H3 vazio (P-28).
6. **ET do Painel:** vazio quando o bias for "SEM EP", como na Estatística (P-10, P-21).
7. **Painel:**
   - lista do H3 filtrada pelo analito e equipamento (P-40);
   - restaurar o período na volta (P-30, P-60);
   - K3/K4 coerentes com De/Até (P-38, P-61);
   - "Últ. violação" do período (P-41, P-59);
   - LJ sem suavização, com eixo Y numerado (P-39, P-62).
8. **Atualizar o GUIA §5–6** para a janela do CEQ, o Consolidar e a correção (P-29, P-49). Dizer também que apagar no
   DB_SEAC não remove do QC_INI.
9. **Indicador de progresso** no ATUALIZAR e na gravação do CEQ.

---

## Relatório B: validação matemática

**Fontes das linhas.**

- As listas `calculos` do `frentes.json`: 93 linhas.
- As tabelas do 1º e do 2º passe da Bio F2 (Glicose e Bilirrubina direta, `Bioquimica_F2.md` §3 e §3b), que o JSON
  não traz.
- Os testes agregados de cálculo da Bio F4 (`frentes.json`, testes P-*, Q-*, E-*, L-*, C-*, G-*, W-*).
- As tabelas do relatório da Hema F4 (§2.1 a §2.5), porque o JSON da Hema F4 não tem lista de cálculos.
- As conferências globais das F3 (testes F3-T22/T23/T24/T26/T27).

**Precisão e método.**

- A tolerância é 1e-9 relativo sobre o `Value2` (REGRAS_COMUNS §3.5).
- Exceção: o DPM com Sigma > 6,8 usa tolerância absoluta de 1e-9, justificada em `Hematologia_F3.md` §4.1.
- Fórmulas de referência, conferidas no código: bias % = (X − alvo)/alvo × 100. Bias do ET/Sigma = média, por rodada,
  das médias de |bias| (`mCEQ.BiasEQ`). ET = 1,65 × CV do lote no período + |bias|. Sigma = (ETp − |bias|)/CVp, com o
  CVp pooled de 6 meses dos lotes com n ≥ 20.
- O DP é amostral (n − 1).

### B.1 Bioquímica: 88 cálculos, 83 aprovados, 5 reprovados

#### Bioquímica — CIQ: n, média, DP, CV (6: 4 aprovados, 2 reprovados)

| Frente | Indicador | Analito | Equipamento | Lote | Nível | Período | Entradas | Fórmula | Independente | QC_INI | Diferença | Status |
|---|---|---|---|---|---|---|---|---|---|---|---|---|
| F1 | n/media/DP/CV | Glicose | DIMENSION 1 | 8977 | N1 | 01/10/2026-31/10/2026 | tblCQ_Final PARTICIPA=SIM + MAN_0001 77,111 | fmean; stdev (n-1); CV=DP/media*100 | n 26; media 77,35042307692308; DP 1,0921468096580027; CV 1,4119467822068532 | engPainel n 26; 77,35042307692308; 1,092146809658003; 1,4119467822068537 (Estatistica DP 1,0921468096578573) | rel max 1,3e-13 | APROVADO |
| F1 | n/media/DP/CV | Glicose | DIMENSION 1 | 8977 | N2 | 01/10/2026-31/10/2026 | tblCQ_Final PARTICIPA=SIM + MAN_0002 244,222 | fmean; stdev (n-1); CV | n 25; 247,76888; 74,94913215881823; 30,24961494713066 | engPainel n 25; 247,76888; 74,94913215881823; 30,249614947130663 | <1e-14 | APROVADO |
| F1 | n/media/DP/CV apos SUBSTITUI_ID | Glicose | DIMENSION 1 | 8977 | N2 | 01/10/2026-31/10/2026 | BIO-371184 (1) substituido por MAN_0001 270,111 | fmean; stdev; CV | n 24; 259,129625; 55,681958305777144; 21,488071194398227 | engPainel n 24; 259,129625; 55,68195830577715; 21,48807119439823 | rel 3,3e-15 | APROVADO |
| F1 | n/media/DP/CV com equipamento em minusculas (L01) | Glicose | DIMENSION 1 | 8977 | N2 | 01/10/2026-31/10/2026 | manual 'dimension 1' 244,222 em 29/10 23:36 | fmean; stdev; CV, populacao do filtro (equipamento exato) | exato: n 24; 247,91666666666666; CV 30,880297740178257. Sem diferenciar maiusculas: n 25; 247,76888 | engPainel n 24; 246,55091666666667 (troca do 277 de BIO-371977); Estatistica n 25; 247,76888 | Painel x Estatistica divergem em n e media; Painel nao corresponde a nenhuma populacao valida | **REPROVADO** |
| F1 | n/media/DP/CV com analito em minusculas (L02) | Glicose | DIMENSION 1 | 8977 | N2 | 01/10/2026-31/10/2026 | manual 'glicose' 244,222 em 28/10 23:36 (+ L01) | fmean; stdev; CV | exato: n 24; 247,91666666666666. Sem diferenciar maiusculas: n 26; 247,63246153846154 | engPainel n 24; 245,2268333333333 (troca tambem o 276 de BIO-371038); Estatistica n 26; 247,6324615384615 | Painel x Estatistica divergem | **REPROVADO** |
| F3 | CV % | Proteína total | DIMENSION 2 | 8977 | 1 | 01-31/10/2026 | n=14 (BIO-359368..370822) | DP(n-1)/média*100 | 1.860147520339105 | 1.860147520339105 | 5.8e-13 | APROVADO |

#### Bioquímica — Equivalência manual × automático (1: 1 aprovados, 0 reprovados)

| Frente | Indicador | Analito | Equipamento | Lote | Nível | Período | Entradas | Fórmula | Independente | QC_INI | Diferença | Status |
|---|---|---|---|---|---|---|---|---|---|---|---|---|
| F1 | equivalencia manual x automatico | Glicose | DIMENSION 1 | 8977 | N1/N2 | outubro/2026 | 77,111/244,222 via manual (A) e via DB_SEAC sintetico BIO-9000001/2 (B) | comparacao bit a bit engPainel, Estatistica, Eng_Saida | identico nas duas copias | identico (Eng_Saida difere so no N1 ID) | 0 | APROVADO |

#### Bioquímica — Períodos e lotes (filtros do Painel e da Estatística) (12: 10 aprovados, 2 reprovados)

| Frente | Indicador | Analito | Equipamento | Lote | Nível | Período | Entradas | Fórmula | Independente | QC_INI | Diferença | Status |
|---|---|---|---|---|---|---|---|---|---|---|---|---|
| F4 | n/média/DP (período) | Glicose | DIMENSION 1 | 8974 | 1 | T1/2026 | 59 resultados ativos 01/01–31/03/2026 | statistics.fmean / stdev (n−1) sobre os dados brutos do trimestre | 59; 87,8983; 1,2275 | 59; 87,8983; 1,2275 (engPainel) | ≤1,7e-16 rel. | APROVADO |
| F4 | n/média/DP (período) | Glicose | DIMENSION 1 | 8974 | 1 | Tudo (sem filtro) | 174 resultados ativos do lote | statistics sobre todo o histórico ativo do lote | 174; 92,2874; 30,3737 | 140; 89,1500; 16,9680 (= janela de 180 corridas) | DP 0,44 rel. | **REPROVADO** |
| F4 | n/média/DP (período) | Ureia | DIMENSION 2 | 8974 | 2 | T1/2026 | 63 resultados ativos | statistics sobre os dados brutos do trimestre | 63; 94,0476; 12,4712 | 63; 94,0476; 12,4712 | ≤1,4e-16 rel. | APROVADO |
| F4 | n/média/DP (período) | Hemoglobina glicada | DIMENSION 1 | 8588 | 1 | dez/2025 (lote minoritário) | 8 resultados ativos do 8588 no mês | statistics, filtro por lote | 8; 4,9525; 0,1427 | 8; 4,9525; 0,1427 | ≤2,0e-16 rel. | APROVADO |
| F4 (testes agregados) | n/média/DP/CV por mês | Glicose | DIMENSION 1 | lote do mês (H3) | N1 e N2 | 13 meses, out/2025 a out/2026 (K3/K4) | linhas ativas do mês | statistics.fmean/stdev (n−1) | 13 meses recalculados | Painel = independente; G3/G4 certos | máx. 4,4e-16 | APROVADO |
| F4 (testes agregados) | n/média/DP/CV trimestre, semestre, ano | Glicose | DIMENSION 1 | lote do período | N1 e N2 | T4/2025, T1–T4/2026, S1, Ano inteiro | dados brutos do período (não média dos meses) | statistics | T1/2026 N1 87,8983 (média das médias daria 87,8384) | idêntico | máx. 9,8e-16 | APROVADO |
| F4 (testes agregados) | n/média/DP/CV (Q-UR2, 20 testes) | Ureia | DIMENSION 2 | lote do período | N1 e N2 | 13 meses, T4/2025–T4/2026, S1, S2 | dados brutos | statistics | 20 recálculos | idêntico | máx. 5,6e-16 | APROVADO |
| F4 (testes agregados) | n/média/DP/CV (Q-HBA1, 22 testes) | Hemoglobina glicada | DIMENSION 1 | 8588 → 8589 → 8590 | N1 e N2 | meses, trimestres, semestres, inclusive lote minoritário dos meses de troca | dados brutos | statistics | 22 recálculos | idêntico | máx. 4,4e-16 | APROVADO |
| F4 (testes agregados) | Estatística MENSAL/TRIMESTRAL (62 linhas) | todos | DIMENSION 1 | lote em análise | todos | mensal e trimestral | tblCQ_Final | statistics | 936 valores | idêntico (UDF com cache) | máx. 1,2e-11 | APROVADO |
| F4 (testes agregados) | Lote A → lote B (estatística e Eng_Saida) | Glicose | DIMENSION 1 | 8974 / 8977 | N1 e N2 | período do teste | linhas do lote filtrado | statistics + escopo do motor | 0 pontos fora do escopo, 0 faltando | idêntico | máx. 4,3e-16 | APROVADO |
| F4 (testes agregados) | Combinações analito × equipamento (10 testes) | Glicose DIM 2, Ureia DIM 2, Proteína total DIM 1, Creatinina DIM 1, Colesterol DIM 2 | DIMENSION 1/2 | 8977 | N1 e N2 | set/2026 | dados brutos | statistics + motor | 10 recálculos | idêntico | 0 divergências | APROVADO |
| F4 (testes agregados) | ET do Painel com bias SEM EP (C-HBA1C) | Hemoglobina glicada | DIMENSION 1 | 8590 | N2 | período do teste | CV do lote; bias SEM EP | ADR-068: sem bias, sem ET | ET vazio (Estatística H vazio) | Painel H8 = 2,7740 % = 1,65 × CV | ET exibido sem base | **REPROVADO** |

#### Bioquímica — Bias do CEQ (23: 23 aprovados, 0 reprovados)

| Frente | Indicador | Analito | Equipamento | Lote | Nível | Período | Entradas | Fórmula | Independente | QC_INI | Diferença | Status |
|---|---|---|---|---|---|---|---|---|---|---|---|---|
| F2 (1º passe) | Bias CEQ (G) | Glicose | DIMENSION 2 | 8977 | 1 e 2 | CEQ CAP, etapa B0 (base) | 15 amostras, 3 rodadas | G | 2,7332944432981674 | 2,733294443298167 | 1,6e-16 | APROVADO |
| F2 (1º passe) | Bias CEQ (G) | Glicose | DIMENSION 2 | 8977 | 1 e 2 | CEQ CAP, etapa B1 inclusão (+10 %) | 16 / 4 | G = (1,49097+1,43622+5,27270+10)/4 | 4,549970832473625 | 4,549970832473625 | 0 | APROVADO |
| F2 (1º passe) | Bias com sinal (T) | Glicose | DIMENSION 2 | 8977 | 1 e 2 | CEQ CAP, etapa B1 | idem | T (com sinal) | 1,5051721711534922 | 1,5051721711534924 | 1,5e-16 | APROVADO |
| F2 (1º passe) | Bias CEQ (G) | Glicose | DIMENSION 2 | 8977 | 1 e 2 | CEQ CAP, etapa B2 correção (+5 %) | 16 / 4 | G | 3,2999708324736257 | 3,299970832473625 | 1,3e-16 | APROVADO |
| F2 (1º passe) | Bias com sinal (T) | Glicose | DIMENSION 2 | 8977 | 1 e 2 | CEQ CAP, etapa B2 | idem | T | 0,2551721711534922 | 0,25517217115349244 | 8,7e-16 | APROVADO |
| F2 (1º passe) | Bias CEQ (G) | Glicose | DIMENSION 2 | 8977 | 1 e 2 | CEQ CAP, etapa N4 = 2026 | 1 amostra (R2) | G | 30 | 30 | 0 | APROVADO |
| F2 (1º passe) | Bias CEQ (G) | Glicose | DIMENSION 2 | 8977 | 1 e 2 | CEQ CAP, etapa N4 = TODOS | 17 amostras, 5 rodadas | G | 8,639976665978901 | 8,639976665978901 | 0 | APROVADO |
| F2 (2º passe) | Bias CEQ (G) | Bilirrubina direta | DIMENSION 2 | 8977 | 1 e 2 | CEQ CAP, etapa B0 (base) | 15 amostras, 3 rodadas | G | 30,770933248452042 | 30,77093324845204 | 1,2e-16 | APROVADO |
| F2 (2º passe) | Bias CEQ (G) | Bilirrubina direta | DIMENSION 2 | 8977 | 1 e 2 | CEQ CAP, etapa B1 inclusão (+10 %) | 16 / 4 | G = (44,8571+21,6335+25,8222+10)/4 | 25,578199936339033 | 25,57819993633903 | 1,4e-16 | APROVADO |
| F2 (2º passe) | Bias com sinal (T) | Bilirrubina direta | DIMENSION 2 | 8977 | 1 e 2 | CEQ CAP, etapa B1 | idem | T (com sinal) | −20,578199936339033 | −20,57819993633903 | 1,7e-16 | APROVADO |
| F2 (2º passe) | Bias CEQ (G) | Bilirrubina direta | DIMENSION 2 | 8977 | 1 e 2 | CEQ CAP, etapa B2 correção (+5 %) | 16 / 4 | G | 24,328199936339033 | 24,32819993633903 | 1,5e-16 | APROVADO |
| F2 (2º passe) | Bias com sinal (T) | Bilirrubina direta | DIMENSION 2 | 8977 | 1 e 2 | CEQ CAP, etapa B2 | idem | T | −21,828199936339033 | −21,82819993633903 | 1,6e-16 | APROVADO |
| F2 (2º passe) | Bias / ET / Sigma | Bilirrubina direta | DIMENSION 2 | 8977 | 1 / 2 | CEQ CAP, etapa N4 = 2026 | 1 amostra (R2) | G / H / L | 30 / 36,343095717491046 (N1) / 1,3428082542927773 (N2) | idem | 0 | APROVADO |
| F2 (2º passe) | Bias (G) / Sigma (L) | Bilirrubina direta | DIMENSION 2 | 8977 | 1 / 2 | CEQ CAP, etapa N4 = TODOS | 17 amostras, 5 rodadas | G / L | 25,462559949071228 / 0,6872842624405966 (N1) | 25,462559949071224 / 0,6872842624405968 | ≤ 3e-16 | APROVADO |
| F2 (2º passe) | Bias CEQ (G) | Bilirrubina total (controle) | DIMENSION 2 | 8977 | 1 e 2 | CEQ CAP, etapa N4 = TODOS | 15 amostras, 3 rodadas | G | 10,82231673715532 | 10,82231673715532 | 0 | APROVADO |
| F2 (3º passe) | Bias CEQ (G) B0 | Proteína total | DIMENSION 2 | 8977 | 1 e 2 | Ano EP 2025, rodadas TODAS | 15 amostras, 3 rodadas CAP | média das médias de \|bias\| por rodada (mCEQ.BiasEQ) | 3.3793363600341007 | 3.3793363600341007 | 0 | APROVADO |
| F2 (3º passe) | Bias CEQ (G) após inclusão | Proteína total | DIMENSION 2 | 8977 | 1 e 2 | Ano EP 2025 | 16 amostras, 4 rodadas (+TESTE-AUDITORIA R1 110/100) | (4,61503+2,35503+3,16794+10)/4 | 5.034502270025575 | 5.034502270025575 | 0 | APROVADO |
| F2 (3º passe) | Bias CEQ (G) após correção + Consolidar | Proteína total | DIMENSION 2 | 8977 | 1 e 2 | Ano EP 2025 | R1 105/100 | média das médias de \|bias\| | 3.7845022700255755 | 3.7845022700255755 | 0 | APROVADO |
| F2 (3º passe) | Bias CEQ (G) N4=TODOS | Proteína total | DIMENSION 2 | 8977 | 1 e 2 | 2025-2026 | 17 amostras, 5 rodadas | média das médias de \|bias\| | 9.02760181602046 | 9.02760181602046 | 0 | APROVADO |
| F2 (3º passe) | Bias CEQ (G) controle N4=TODOS | Albumina | DIMENSION 2 | 8977 | 1 e 2 | 2025-2026 | 15 amostras, 3 rodadas | média das médias de \|bias\| | 5.039521617989425 | 5.039521617989425 | 0 | APROVADO |
| F3 | Bias EQC abs % | Proteína total | DIMENSION 2 | 8977 | 1 | CAP 2025, 3 rodadas | 15 amostras, C-A 4,615032 / C-B 2,355033 / C-C 3,167944 | média das médias por rodada de \|(X-alvo)/alvo*100\| | 3.3793363600341007 | 3.3793363600341007 | 0 | APROVADO |
| F3 | Bias EQC abs % | Bilirrubina direta | DIMENSION 2 | 8977 | 1 | CAP 2025 | C-A 44,8571 / C-B 21,6335 / C-C 25,8222 (inclui 2 amostras com resultado 0 = -100 %) | duas etapas | 30.77093324845204 | 30.77093324845204 | 1.2e-16 | APROVADO |
| F3 | Bias / ET / Sigma | Hemoglobina glicada | DIMENSION 2 | 8977 | 1 e 2 | 01-31/10/2026 | sem CEQ; n=0 no lote 8977 | SEM EP -> sem ET/Sigma (ADR-068) | SEM EP / - / - | SEM EP / - / - | igual | APROVADO |

#### Bioquímica — ET % (15: 15 aprovados, 0 reprovados)

| Frente | Indicador | Analito | Equipamento | Lote | Nível | Período | Entradas | Fórmula | Independente | QC_INI | Diferença | Status |
|---|---|---|---|---|---|---|---|---|---|---|---|---|
| F2 (1º passe) | ET (H) | Glicose | DIMENSION 2 | 8977 | 1 | CEQ CAP, etapa B1 | F = 1,2997 | H = 1,65·F + G | 6,694469325367574 | 6,694469325367574 | 0 | APROVADO |
| F2 (1º passe) | ET (H) | Glicose | DIMENSION 2 | 8977 | 2 | CEQ CAP, etapa B1 | F = 1,4462 | H | 6,936230926868154 | 6,936230926868154 | 0 | APROVADO |
| F2 (1º passe) | ET (H) | Glicose | DIMENSION 2 | 8977 | 1 | CEQ CAP, etapa B2 | idem | H | 5,444469325367574 | 5,444469325367574 | 0 | APROVADO |
| F2 (1º passe) | ET (H) | Glicose | DIMENSION 2 | 8977 | 2 | CEQ CAP, etapa B2 | idem | H | 5,686230926868154 | 5,686230926868154 | 0 | APROVADO |
| F2 (1º passe) | ET (H) / Sigma (L) | Glicose | DIMENSION 2 | 8977 | 1 / 2 | CEQ CAP, etapa N4 = 2026 |  | H / L | 32,14449849289395 / −3,1215897725166273 (N2) | idem | 0 | APROVADO |
| F2 (2º passe) | ET (H) / Sigma (L) | Bilirrubina direta | DIMENSION 2 | 8977 | 1 / 2 | CEQ CAP, etapa B0 | F, CVp | H / L | 37,114028965943085 / 1,238400231033604 (N2) | 37,114028965943085 / 1,2384002310336044 | ≤ 3,8e-16 | APROVADO |
| F2 (2º passe) | ET (H) | Bilirrubina direta | DIMENSION 2 | 8977 | 1 | CEQ CAP, etapa B1 | F = 3,8443 | H = 1,65·F + G | 31,921295653830075 | 31,92129565383007 | 1,1e-16 | APROVADO |
| F2 (2º passe) | ET (H) | Bilirrubina direta | DIMENSION 2 | 8977 | 2 | CEQ CAP, etapa B1 | F = 1,7457 | H | 28,458648487281312 | 28,45864848728131 | 1,2e-16 | APROVADO |
| F2 (2º passe) | ET (H) | Bilirrubina direta | DIMENSION 2 | 8977 | 1 / 2 | CEQ CAP, etapa B2 | idem | H | 30,671295653830075 / 27,208648487281312 | 30,67129565383007 / 27,20864848728131 | ≤ 1,3e-16 | APROVADO |
| F2 (3º passe) | ET (H) após inclusão | Proteína total | DIMENSION 2 | 8977 | 1 / 2 | 01-31/10/2026 | F N1 1,86015, N2 2,20128 | 1,65·CV + \|bias\| | 8.103745678586874 / 8.666611564082336 | 8.103745678586874 / 8.666611564082336 | 0 | APROVADO |
| F2 (3º passe) | ET (H) após correção | Proteína total | DIMENSION 2 | 8977 | 1 / 2 | 01-31/10/2026 | idem | 1,65·CV + \|bias\| | 6.853745678586875 / 7.416611564082336 | 6.853745678586875 / 7.416611564082336 | 0 | APROVADO |
| F2 (3º passe) | ET (H) com H3=8974 e H3 vazio | Proteína total | DIMENSION 2 | 8974 | 1 | Tudo (período perdido) | F 94,9776; G 3,7845 | 1,65·F + G (G e L não mudam com o lote) | 160.49753916758107 | 160.49753916758107 | 0 | APROVADO |
| F3 | ET % | Proteína total | DIMENSION 2 | 8977 | 1 | 01-31/10/2026 | CV 1,860148; bias 3,379336 | 1,65*CV + \|bias\| | 6.448579768593623 | 6.448579768593623 | 2.8e-13 | APROVADO |
| F3 | ET % | Ureia | DIMENSION 2 | 8977 | 1 | 01-31/10/2026 | CV 7,563602; bias 5,910451 (3 rodadas) | 1,65*CV + \|bias\| | 18.39039314 | 18.39039314 | 0 | APROVADO |
| F3 | ET % | Potássio | DIMENSION 2 | 8977 | 2 | 01-31/10/2026 | CV 1,632448; bias 2,579837 | 1,65*CV + \|bias\| | 5.273376204921652 | 5.273376204924055 | 4.6e-13 | APROVADO |

#### Bioquímica — ETp (especificação) (1: 1 aprovados, 0 reprovados)

| Frente | Indicador | Analito | Equipamento | Lote | Nível | Período | Entradas | Fórmula | Independente | QC_INI | Diferença | Status |
|---|---|---|---|---|---|---|---|---|---|---|---|---|
| F3 | ETp % | Bilirrubina direta | DIMENSION 2 | 8977 | 1 | biblioteca 2025 | VB MIN f=0,75; CVi 20,9; CVg 31,1 | 1,65*f*CVi + f/2*raiz(CVi^2+CVg^2) | 39.91509557613611 | 39.91509557613611 | 0 | APROVADO |

#### Bioquímica — Sigma, classe e DPM (18: 18 aprovados, 0 reprovados)

| Frente | Indicador | Analito | Equipamento | Lote | Nível | Período | Entradas | Fórmula | Independente | QC_INI | Diferença | Status |
|---|---|---|---|---|---|---|---|---|---|---|---|---|
| F2 (1º passe) | Sigma (L) | Glicose | DIMENSION 2 | 8977 | 1 | CEQ CAP, etapa B1 | CVp 22,0341 | L = (10 − G)/CVp | 0,24734552311288047 | 0,24734552311288047 | 0 | APROVADO |
| F2 (1º passe) | Sigma (L) | Glicose | DIMENSION 2 | 8977 | 2 | CEQ CAP, etapa B1 | CVp 6,4070 | L | 0,850637765463382 | 0,850637765463382 | 0 | APROVADO |
| F2 (1º passe) | Sigma (L) | Glicose | DIMENSION 2 | 8977 | 1 | CEQ CAP, etapa B2 | idem | L | 0,3040758440684709 | 0,3040758440684709 | 0 | APROVADO |
| F2 (1º passe) | Sigma (L) | Glicose | DIMENSION 2 | 8977 | 2 | CEQ CAP, etapa B2 | idem | L | 1,0457371262456712 | 1,045737126245671 | ~1e-16 | APROVADO |
| F2 (1º passe) | Sigma (L) | Glicose | DIMENSION 2 | 8977 | 1 / 2 | CEQ CAP, etapa N4 = TODOS |  | L | 0,06172364819688728 / 0,21227174649321132 | idem | 0 | APROVADO |
| F2 (2º passe) | Sigma (L) | Bilirrubina direta | DIMENSION 2 | 8977 | 1 | CEQ CAP, etapa B1 | CVp 21,0285 | L = (39,9151 − G)/CVp | 0,6817850514087859 | 0,681785051408786 | 1,6e-16 | APROVADO |
| F2 (2º passe) | Sigma (L) | Bilirrubina direta | DIMENSION 2 | 8977 | 2 | CEQ CAP, etapa B1 | CVp 7,3839 | L | 1,9416556964301084 | 1,9416556964301088 | 2,3e-16 | APROVADO |
| F2 (2º passe) | Sigma (L) | Bilirrubina direta | DIMENSION 2 | 8977 | 1 / 2 | CEQ CAP, etapa B2 | idem | L | 0,7412282764745606 / 2,110944061325536 | 0,7412282764745607 / 2,1109440613255366 | ≤ 2,1e-16 | APROVADO |
| F2 (3º passe) | Sigma (L) após inclusão | Proteína total | DIMENSION 2 | 8977 | 1 / 2 | CVp 29/04-28/10/2026 | ETp 10 (CLIA); CVp 3,32573 / 5,92399 | (ETp − \|bias\|)/CVp | 1.4930545183334316 / 0.8382017183904314 | 1.4930545183334316 / 0.8382017183904314 | 0 | APROVADO |
| F2 (3º passe) | Sigma (L) após correção | Proteína total | DIMENSION 2 | 8977 | 1 / 2 | CVp 6 meses | idem | (ETp − \|bias\|)/CVp | 1.868911733341443 / 1.0492081884293243 | 1.868911733341443 / 1.0492081884293243 | 0 | APROVADO |
| F2 (3º passe) | Sigma (L) N4=2026 (\|bias\| > ETp) | Proteína total | DIMENSION 2 | 8977 | 1 / 2 | Ano EP 2026 | 1 amostra, bias 30 | (ETp − \|bias\|)/CVp | -6.0137154401281805 / -3.3761035206222867 | -6.0137154401281805 / -3.3761035206222867 | 0 | APROVADO |
| F3 | Sigma | Proteína total | DIMENSION 2 | 8977 | 1 | CVp 29/04-28/10/2026 | ETp CLIA 10; CVp 3,325731 (8974 n=44, 8977 n=90) | (ETp-\|bias\|)/CVp | 1.9907393577787527 | 1.9907393577779087 | 4.2e-13 | APROVADO |
| F3 | Sigma | Proteína total | DIMENSION 2 | 8977 | 2 | CVp 29/04-28/10/2026 | ETp 10; CVp 5,923989 | (ETp-\|bias\|)/CVp | 1.1176022911872259 | 1.1176022911872419 | 1.4e-14 | APROVADO |
| F3 | Sigma | Bilirrubina direta | DIMENSION 2 | 8977 | 1 | CVp 29/04-28/10/2026 | ETp 39,9151; bias 30,7709 | (ETp-\|bias\|)/CVp | 0.4348467994 | 0.4348467994 | 7.0e-15 | APROVADO |
| F3 | Sigma | Ureia | DIMENSION 2 | 8977 | 2 | CVp 29/04-28/10/2026 | ETp CLIA 9 | (ETp-\|bias\|)/CVp | 0.4004936761 | 0.4004936761 | 3.5e-15 | APROVADO |
| F3 | Sigma | Glicose | DIMENSION 2 | 8977 | 1 | CVp 29/04-28/10/2026 | ETp 10; bias 2,733294; CVp 22,03 (inclui BIO-319726 = 277) | (ETp-\|bias\|)/CVp | 0.3297940308 | 0.3297940308 | 1.2e-15 | APROVADO |
| F3 | Sigma | Lactato | DIMENSION 2 | 8977 | 1 | CVp 29/04-28/10/2026 | ETp VB ÓTI 15,209694; CVp 7,314608 | (ETp-\|bias\|)/CVp | 1.7101053181939625 | 1.7101053181939767 | 8.3e-15 | APROVADO |
| F3 | Sigma | Lipase | DIMENSION 2 | 8977 | 1 | CVp 29/04-28/10/2026 | ETp VB MIN 18,845471; CVp 4,533547 | (ETp-\|bias\|)/CVp | 2.47520034398127 | 2.4752003439812365 | 1.4e-14 | APROVADO |

#### Bioquímica — Levey-Jennings e Westgard (5: 5 aprovados, 0 reprovados)

| Frente | Indicador | Analito | Equipamento | Lote | Nível | Período | Entradas | Fórmula | Independente | QC_INI | Diferença | Status |
|---|---|---|---|---|---|---|---|---|---|---|---|---|
| F4 | Westgard por ponto (H3 vazio, lote a lote) | Glicose | DIMENSION 1 | 8974+8977+8353 | 1 e 2 | 01/06–31/08/2026 | 180 corridas RUN\|LOTE, alvo do cadastro por lote | ADR-083/086: 1_3s, 2_2s within/across, R_4s, 4_1s, 8x por lote na janela | 265 marcas; contagens N1 [10,26,0,5,2] N2 [2,8,0,5,2] | idênticas (Eng_Saida e Painel M:Q) | 0 | APROVADO |
| F4 | Linhas de limite em degrau | Ureia | DIMENSION 1 | 8974+6353 | 1 e 2 | dez/2025 | Média/DP do lote de cada ponto | M_lote + k·DP_lote, k = −3..+3 | 266 valores | idênticos nas 7 séries | 0 | APROVADO |
| F4 (testes agregados) | LJ e Westgard (G-*, 4 cenários, 36 testes) | Glicose e outros | DIMENSION 1/2 | lote do cenário | N1 e N2 | 4 cenários | Eng_Saida + cadastro do lote | eixo, pontos, corrida do lote (ADR-080), 7 limites, Westgard por ponto (ADR-083), contagens M:Q | 609 marcas de Westgard | idêntico | 0 divergências | APROVADO |
| F4 (testes agregados) | Westgard por ponto, H3 vazio (W-HBA1-JUN) | Hemoglobina glicada | DIMENSION 1 | 8589 + 8590 simultâneos | N1 e N2 | jun/2026 | Eng_Saida | ADR-083/086 lote a lote | 288 marcas | idêntico | 0 divergências | APROVADO |
| F4 (testes agregados) | Westgard por ponto, H3 vazio (W-UR1-DEZ) | Ureia | DIMENSION 1 | 8974 + 6353 (urina) | N1 e N2 | dez/2025 | Eng_Saida | ADR-083/086 lote a lote | 125 marcas | idêntico (eixo Y até 1.932 pela urina, ver OC-F4-11) | 0 divergências | APROVADO |

#### Bioquímica — Resumos e contadores (1: 0 aprovados, 1 reprovados)

| Frente | Indicador | Analito | Equipamento | Lote | Nível | Período | Entradas | Fórmula | Independente | QC_INI | Diferença | Status |
|---|---|---|---|---|---|---|---|---|---|---|---|---|
| F3 | Resumo Sem EP | todos | DIMENSION 2 | 8977 | todos | - | COUNTIF(G14:G93,"SEM EP") | contagem de níveis sem CEQ | 4 | 22 | 18 linhas vazias (A=0) contadas | **REPROVADO** |

#### Bioquímica — Conferência global (calculadora independente) (6: 6 aprovados, 0 reprovados)

| Frente | Indicador | Analito | Equipamento | Lote | Nível | Período | Entradas | Fórmula | Independente | QC_INI | Diferença | Status |
|---|---|---|---|---|---|---|---|---|---|---|---|---|
| F2 (1º passe) | Estatística completa (19 colunas) | Todos (62 linhas) | DIMENSION 2 | 8977 | 1 e 2 | CEQ CAP, etapa B0, S1, S2 | tblCQ_Final + EQA_Base | calc_independente (19 colunas) |  |  | máx. < 1e-9 | APROVADO (0/1.178 divergências em cada) |
| F2 (2º passe) | Estatística completa (19 colunas) | Todos (62 linhas) | DIMENSION 2 | 8977 | 1 e 2 | CEQ CAP, etapa S1, S2 (Run A e Run B) | tblCQ_Final + EQA_Base | calc_independente (19 colunas) |  |  | máx. < 1e-9 | APROVADO (0/1.178 em cada um dos 4 retratos) |
| F2 (3º passe) | Estatística completa (19 colunas) | Todos (62 linhas) | DIMENSION 2 | 8977 | 1 e 2 | 01-31/10/2026 | tblCQ_Final + EQA_Base (S1 e S2) | calc_independente.py | 0 divergências | 1.178 comparações por retrato | < 1e-9 | APROVADO |
| F3 | Global (19 colunas) | todos (31 analitos) | DIMENSION 2 | 8977 | todos | 01-31/10/2026; CVp 29/04-28/10/2026 | 62 linhas analito x nível | C,D,E,F,G,H,J,K,L,M,N,O,T,V,W,AC,AD,AF,AR | 1178 comparações | 1178 comparações | 0 divergências (máx 1,9e-11) | APROVADO |
| F3 (teste F3-T23) | Segunda opinião independente | 4 analitos | DIMENSION 2 | 8977 | N1 e N2 | 01-31/10/2026 | idem | idem | 56 valores | 56 valores | 56/56 (máx. 5,8e-13) | APROVADO |
| F3 (teste F3-T24) | Segunda opinião do zero (script sem a calculadora) | Lactato, Lipase, Creatinina, Potássio, Bicarbonato, Colesterol total | DIMENSION 2 | 8977 | N1 e N2 | 01-31/10/2026; CVp 29/04-28/10/2026 | tblCQ_Final + EQA_Base + biblioteca | n, média, DP, CV, bias, ET, ETp, Sigma | 96 valores | 96 valores | 96/96 dentro de 1e-9 (máx. 8,9e-13) | APROVADO |

### B.2 Hematologia: 86 cálculos, 78 aprovados, 8 reprovados

#### Hematologia — CIQ: n, média, DP, CV (9: 7 aprovados, 2 reprovados)

| Frente | Indicador | Analito | Equipamento | Lote | Nível | Período | Entradas | Fórmula | Independente | QC_INI | Diferença | Status |
|---|---|---|---|---|---|---|---|---|---|---|---|---|
| F1 | media/DP/CV | WBC | XN1000 | 625311 | N1 | Tudo (Painel) / ano 2026 (Estatística) | 22 resultados PARTICIPA=SIM (21 + MAN_0001 3,111) | fmean; stdev n-1; CV=DP/média*100 | n 22; 3,120045454545455; 0,05290645587821089; 1,6956950355044937 | engPainel = Estatística: n 22; 3,120045454545455; 0,05290645587821089; 1,6956950355044937 | 0 | APROVADO |
| F1 | media/DP/CV | WBC | XN1000 | 625311 | N2 | Tudo / 2026 | 40 resultados (39 + MAN_0002 7,322) | fmean; stdev n-1; CV | n 40; 7,336799999999999; 0,11213709558004521; 1,5284196867850457 | n 40; 7,336799999999999; 0,11213709558004521; 1,5284196867850455 | 1,5e-16 relativo no CV | APROVADO |
| F1 | media/DP/CV | WBC | XN1000 | 625311 | N3 | Tudo / 2026 | 22 resultados (21 + MAN_0003 17,433) | fmean; stdev n-1; CV | n 22; 17,385590909090908; 0,1750675132602679; 1,0069690134531193 | idêntico | 0 | APROVADO |
| F1 | media após SUBSTITUI_ID | WBC | XN1000 | 625311 | N1 | Tudo / 2026 | 21 resultados, 3,11 de HEM-372173 trocado por 3,111 | (21*3,1204761904761904 - 3,11 + 3,111)/21 | 3,1205238095238097 | 3,1205238095238097 (Painel = Estatística) | 0 | APROVADO |
| F1 | media/DP/CV com equipamento xn1000 | WBC | XN1000 | 625311 | N3 | Tudo / 2026 | 21 reais + MAN_0001 17,433 com equipamento xn1000 (RUN 26102901) | independente com equipamento exato: n 21; um valor por RUN reproduz o Painel; sem diferenciar maiúsculas reproduz a Estatística | exato n 21 média 17,383333333333333; troca n 21 média 17,376333333333335 | Painel n 21 média 17,37633333333333 DP 0,1737853081630701; Estatística n 22 média 17,385590909090908 | Painel e Estatística divergem entre si (n 21 x 22) e do esperado | **REPROVADO** |
| F1 | media/DP/CV com analito wbc | WBC | XN1000 | 625311 | N3 | Tudo / 2026 | base do D02 + MAN_0002 17,433 com analito wbc (RUN 26102801) | 364,903 - 17,14 + 17,433 = 365,196 / 21 | 17,390285714285714 (troca) | Painel n 21 média 17,390285714285714; Estatística n 23 média 17,38765217391304 | Painel x Estatística n 21 x 23 | **REPROVADO** |
| F1 | media/DP/CV com espaço sobrando | WBC | XN1000 | 625311 | N3 | Tudo / 2026 | 21 + 2 manuais 17,433 ('WBC ' e 'XN1000 ') | fmean; stdev; CV | n 23; 17,387652173913043; 0,17132786450005755; 0,9853421427253033 | n 23; 17,38765217391304; 0,17132786450005755; 0,9853421427253034 | 1e-16 relativo | APROVADO |
| F1 | media/DP/CV com manual implausível | WBC | XN1000 | 625311 | N1 | Tudo / 2026 | 21 reais + manual -1 | fmean; stdev; CV | n 22; 2,933181818181818; 0,8800778448449961; 30,004203605439194 | n 22; 2,933181818181818; 0,8800778448449961; 30,004203605439198 | 1,3e-16 relativo (cálculo correto; dado implausível aceito, OC-HF1-7) | APROVADO |
| F3 | CV % | RBC | XN1000 | 522611 | 1 | CIQ ano 2025 | n=50, média 2,4022, DP 0,030256724669982 | DP(n-1)/média*100 | 1.259542280825165 | 1.2595422808251646 | 3.5e-16 | APROVADO |

#### Hematologia — Equivalência manual × automático (1: 1 aprovados, 0 reprovados)

| Frente | Indicador | Analito | Equipamento | Lote | Nível | Período | Entradas | Fórmula | Independente | QC_INI | Diferença | Status |
|---|---|---|---|---|---|---|---|---|---|---|---|---|
| F1 | equivalência manual x automático | WBC | XN1000 | 625311 | N1-N3 | Tudo / 2026 | mesmo trio por MAN_0001-0003 (cópia A) e HEM-9000001-0003 (cópia B) | comparação bit a bit de engPainel, Estatística e independente | igual nos dois caminhos | igual nos dois caminhos; Eng_Saida difere só no N1 ID | 0 | APROVADO |

#### Hematologia — Períodos e lotes (filtros do Painel e da Estatística) (24: 18 aprovados, 6 reprovados)

| Frente | Indicador | Analito | Equipamento | Lote | Nível | Período | Entradas | Fórmula | Independente | QC_INI | Diferença | Status |
|---|---|---|---|---|---|---|---|---|---|---|---|---|
| F4 (relatório §2) | n/média/CV por mês | WBC | XN1000 | 10 lotes (522611 → 625311), os dois nos meses de troca | N1–N3 | 20 combinações mês × lote, out/2025 a out/2026 | linhas PARTICIPA=SIM do mês e do lote | statistics (n−1) | ex.: set/2026 619611 N1 31; 3,1274; CV 2,42 | Painel idêntico; G3/G4 certos | máx. 4,5e-16 | APROVADO |
| F4 (relatório §2) | n/média/CV trimestre | WBC | XN1000 | 528211 | N1–N3 | T4/2025 | dados brutos do trimestre | statistics | 83; 2,9920; 1,925 / 86; 6,8033; 1,459 / 81; 16,4441; 1,373 (média dos meses daria 2,9987; 1,369) | idêntico | 2,2e-16 | APROVADO |
| F4 (relatório §2) | n/média/CV trimestre | WBC | XN1000 | 533811 | N1–N3 | T1/2026 | dados brutos do trimestre | statistics | 68; 3,1672; 3,472 / 59; 6,8515; 1,718 / 59; 15,5712; 2,029 | idêntico | 4,4e-16 | APROVADO |
| F4 (relatório §2) | n/média/CV trimestre | WBC | XN1000 | 608511 | N1–N3 | T2/2026 | dados brutos do trimestre | statistics | 57; 3,1291; 2,418 / 50 / 50 | idêntico | 4,9e-16 | APROVADO |
| F4 (relatório §2) | n/média/CV trimestre | WBC | XN1000 | 619611 | N1–N3 | T3/2026 | dados brutos do trimestre | statistics | 54; 3,1206; 2,275 / 68; 6,9947; 1,283 / 49; 16,8612; 0,874 | idêntico | 4,3e-16 | APROVADO |
| F4 (relatório §2) | n/média/CV períodos longos | WBC | XN1000 | 533811 / 619611 / 528211 / 625311 | N1–N3 | T4/2026, S1/2026, S2/2026, Ano inteiro 2025 e 2026, Tudo | dados brutos | statistics | iguais aos trimestres; Tudo 528211 = 83/86/81 (89 corridas, sem corte) | idêntico | máx. 4,4e-16 | APROVADO |
| F4 (relatório §2) | n (H3 vazio + Ano inteiro) | WBC | XN1000 | todos | N1–N3 | Ano inteiro 2026 | todos os resultados de 2026 | contagem dos dados do período | 337 / 340 / 308 | 141 / 155 / 130 (janela de 180 corridas RUN\|LOTE) | 213 corridas cortadas | **REPROVADO** |
| F4 (relatório §2) | n (H3 vazio + Tudo) | WBC | XN1000 | todos | N1–N3 | Tudo (sem filtro) | todo o histórico | contagem dos dados do período | 478 / 485 / 447 | 141 / 155 / 130 | 366 corridas cortadas | **REPROVADO** |
| F4 (relatório §2) | Estatística C:F, estado de abertura | todos (84 linhas) | XN1000 | 522611 | N1–N3 | Ano 2025–2025 | tblCQ_Final | statistics | 336 valores (WBC 50/51/51) | idêntico | máx. 7,9e-16 | APROVADO |
| F4 (relatório §2) | Estatística após troca do lote no Painel (E-LOTE) | WBC | XN1000 | 619611 | N1 | Ano 2025 | linhas do 619611 em 2025 | statistics | n = 0 (o 619611 não tem resultado em 2025) | B4 = 619611, mas 50; 3,0446; DP 0,054929; CV 1,8041 (números do 522611) | número de outro lote | **REPROVADO** |
| F4 (relatório §2) | Estatística após Ano De/Até = 2026 (E-ANO) | WBC (339 valores) | XN1000 | 619611 | N1–N3 | Ano 2026 | linhas do 619611 em 2026 | statistics | N1 54; 3,12056 | 50; 3,0446 (não recalcula) | 339/339 divergem | **REPROVADO** |
| F4 (relatório §2) | Estatística após AtualizarEstatistica (E-INTERNO, fora do caminho do usuário) | todos (84 linhas) | XN1000 | 619611 | N1–N3 | Ano 2026 | tblCQ_Final | statistics | 54; 3,12056; 0,070988 / 68; 6,99471 / 49; 16,86122 | idêntico | máx. 1,1e-15 | APROVADO |
| F4 (relatório §2) | Lote A → lote B (estatística e motor) | WBC | XN1000 | 536411 | N1–N3 | 01/02–31/03/2026 | linhas do lote | statistics + escopo do motor | N1 58; 3,016897; DP 0,061877 / N2 57; 7,095789 / N3 55; 17,288545 | idêntico; 170 pontos, 0 fora, 0 faltando | 4,3e-16 | APROVADO |
| F4 (relatório §2) | Lote A → lote B (estatística e motor) | WBC | XN1000 | 602911 | N1–N3 | 01/02–31/03/2026 | linhas do lote | statistics + escopo do motor | só o lote | idêntico; 0 fora, 0 faltando | 0 | APROVADO |
| F4 (relatório §2) | H3 vazio (motor, ADR-086) | WBC | XN1000 | 533811 + 536411 + 602911 | N1–N3 | 01/02–31/03/2026 | todos os lotes do período | contagem e escopo | n 91 / 95 / 89; pontos 50 / 170 / 55 | idêntico (tela mostra Bias e Sigma, OC-HF4-3) | 0 | APROVADO |
| F4 (relatório §2) | H3 vazio com lotes paralelos (RUN\|LOTE) | WBC | XN1000 | 528211 + 533811 | N1–N3 | 15–31/12/2025 | RUN 25122201/25122301/25122302 em 2 lotes | corrida = RUN\|LOTE | n 22 / 22 / 21; 2 corridas por RUN | idêntico | 0 | APROVADO |
| F4 (relatório §2) | n/média/DP por analito | RBC | XN1000 | 619611 | N1–N3 | set/2026 | dados brutos | statistics + motor | 31; 2,3155; 0,02873 / 46; 4,1974 / 25; 4,944 | idêntico | 7,2e-16 | APROVADO |
| F4 (relatório §2) | n/média/DP por analito | HGB | XN1000 | 619611 | N1–N3 | set/2026 | dados brutos | statistics + motor | 31; 5,7065 / 46; 11,8261 / 25; 15,604 | idêntico | 3,5e-16 | APROVADO |
| F4 (relatório §2) | n/média/DP por analito | PLT | XN1000 | 619611 | N1–N3 | set/2026 | dados brutos | statistics + motor | 31; 83,7419 / 46; 232,2391 / 25; 546,16 | idêntico | 4,5e-16 | APROVADO |
| F4 (relatório §2) | n/média/DP por analito | MCV | XN1000 | 619611 | N1–N3 | set/2026 | dados brutos | statistics + motor | 31; 72,2065 / 46; 79,287 / 25; 88,44 | idêntico | 2,6e-16 | APROVADO |
| F4 (relatório §2) | n/média/DP por analito | NEUT% | XN1000 | 619611 | N1–N3 | set/2026 | dados brutos | statistics + motor | 31; 42,5161 / 46; 44,8783 / 25; 47,776 | idêntico | 2,6e-16 | APROVADO |
| F4 (relatório §2) | ETp ausente (ADR-068) | IG% | XN1000 | 619611 | N1–N3 | set/2026 | sem ETp | ETp e Sigma '-' | ETp '-', Sigma '-'; bias 3,009; ET 8,14 / 7,52 / 6,73 | idêntico | 0 | APROVADO |
| F4 (relatório §2) | ET do Painel com bias SEM EP | RET% | XN1000 | 619611 | N1–N3 | set/2026 | CV 3,52114; bias SEM EP | ADR-068: sem bias, sem ET | ET vazio | H = 5,8099 % = 1,65 × CV | ET exibido sem base (confirma OC-H2) | **REPROVADO** |
| F4 (relatório §2) | ET do Painel com bias SEM EP | RDW-SD | XN1000 | 619611 | N1–N3 | set/2026 | CV 0,93806; bias SEM EP | ADR-068: sem bias, sem ET | ET vazio | H = 1,5478 = 1,65 × CV | ET exibido sem base (confirma OC-H2) | **REPROVADO** |

#### Hematologia — Bias do CEQ (7: 7 aprovados, 0 reprovados)

| Frente | Indicador | Analito | Equipamento | Lote | Nível | Período | Entradas | Fórmula | Independente | QC_INI | Diferença | Status |
|---|---|---|---|---|---|---|---|---|---|---|---|---|
| F2 | Bias abs (G) | RBC | XN1000 | 522611 | 1-3 | CEQ CAP 2025, S1 inclusao | 16 amostras / 4 rodadas (FH9-A/B/C + TESTE-AUDITORIA R1 110/100) | media das medias por rodada de \|bias\| | 3.6502150939575695 | 3.650215093957569 | 1.2e-16 | APROVADO |
| F2 | Bias abs (G) | RBC | XN1000 | 522611 | 1-3 | S2 correcao (105) | 16 / 4 | duas etapas | 2.4002150939575695 | 2.400215093957569 | 1.9e-16 | APROVADO |
| F2 | Bias abs (G) | RBC | XN1000 | 522611 | 1-3 | N4 = TODOS | 17 / 5 | duas etapas | 7.920172075166056 | 7.920172075166056 | 0 | APROVADO |
| F2 | Bias abs (G) | LYMPH# | XN1000 | 522611 | 1-3 | S1 inclusao | 16 / 4 | duas etapas | 4.9528950092632815 | 4.9528950092632815 | 0 | APROVADO |
| F2 | Bias abs (G) | LYMPH# | XN1000 | 522611 | 1-3 | S2 correcao | 16 / 4 | duas etapas | 3.7028950092632815 | 3.7028950092632815 | 0 | APROVADO |
| F2 | Bias abs (G) | LYMPH% (vizinho) | XN1000 | 522611 | 1-3 | depois de 'Lymphs' (TA-03 150/100) | 16 / 4 | duas etapas | 15.222013284930423 | 15.222013284930423 | 0 | APROVADO |
| F3 | Bias abs % | RBC | XN1000 | 522611 | 1 | CAP 2025, TODAS | FH9-A 0,805702; FH9-B 1,481856; FH9-C 2,313302 (5 amostras cada) | média das médias por rodada de abs((X-alvo)/alvo*100) | 1.533620125276759 | 1.533620125276759 | 0 | APROVADO |

#### Hematologia — ET % (3: 3 aprovados, 0 reprovados)

| Frente | Indicador | Analito | Equipamento | Lote | Nível | Período | Entradas | Fórmula | Independente | QC_INI | Diferença | Status |
|---|---|---|---|---|---|---|---|---|---|---|---|---|
| F2 | ET (H) | RBC | XN1000 | 522611 | 1 | S1 | F 1,2595422808251646; G 3,6502 | 1,65*F + \|G\| | 5.728459857319091 | 5.72845985731909 | 1.6e-16 | APROVADO |
| F3 | ET % | RBC | XN1000 | 522611 | 1 | CIQ 2025 + CAP 2025 | CV 1,259542; bias 1,533620 | 1,65*CV + abs(bias) | 3.611864888638281 | 3.61186488863828 | 2.5e-16 | APROVADO |
| F3 | ET % | MONO% | XN1000 | 522611 | 3 | CIQ 2025 + CAP 2025 | CV 7,324166; bias 6,491155 | 1,65*CV + abs(bias) | 18.5760299086878 | 18.5760299086878 | 0 | APROVADO |

#### Hematologia — Sigma, classe e DPM (31: 31 aprovados, 0 reprovados)

| Frente | Indicador | Analito | Equipamento | Lote | Nível | Período | Entradas | Fórmula | Independente | QC_INI | Diferença | Status |
|---|---|---|---|---|---|---|---|---|---|---|---|---|
| F2 | Sigma (L) | RBC | XN1000 | 522611 | 1 | S1, CVp 6 m | ETp 4; CVp 0,9757600160893026 | (ETp - \|G\|)/CVp | 0.35847431773676797 | 0.3584743177367684 | 1.2e-15 | APROVADO |
| F2 | Sigma (L) | RBC | XN1000 | 522611 | 1 | S2 | ETp 4; CVp 0,97576 | (ETp - \|G\|)/CVp | 1.6395270144949416 | 1.639527014494942 | 2.7e-16 | APROVADO |
| F2 | Sigma (L) | LYMPH# | XN1000 | 522611 | 1 | S1, ETp VB | ETp 22,758; CVp 4,988470640249724 | (ETp - \|G\|)/CVp | 3.5692477770989917 | 3.5692477770989917 | 0 | APROVADO |
| F3 | Sigma | RBC | XN1000 | 522611 | 1 | CVp 02/04-01/10/2026 | ETp CLIA 4; bias 1,533620; CVp 0,975760 (gl 187) | (ETp - abs(bias))/CVp 6 m | 2.52765007179079 | 2.52765007179 | 1.8e-12 | APROVADO |
| F3 | DPM | RBC | XN1000 | 522611 | 1 | CVp 02/04-01/10/2026 | Sigma 2,52765 | (1-Phi(Sigma-1,5))*1e6 | 152057.2308 | 152057.2308 | 7.2e-12 | APROVADO |
| F3 | Sigma | HGB | XN1000 | 522611 | 1 | CVp 02/04-01/10/2026 | n=50; CV=1,091389; bias=1,443113; ETp=4 (CLIA); CVp6=0,894071 gl=187; L20 | (ETp - abs(bias))/CVp 6 m | 2.85982679462 | 2.85982679462 | 1.1e-12 | APROVADO |
| F3 | Sigma | HGB | XN1000 | 522611 | 2 | CVp 02/04-01/10/2026 | CVp6=0,675070 gl=194; L21 | (ETp - abs(bias))/CVp 6 m | 3.78758856942 | 3.78758856939 | 7.0e-12 | APROVADO |
| F3 | Sigma | HGB | XN1000 | 522611 | 3 | CVp 02/04-01/10/2026 | CVp6=0,529493 gl=171; L22 | (ETp - abs(bias))/CVp 6 m | 4.82893413352 | 4.82893413356 | 9.4e-12 | APROVADO |
| F3 | Sigma | MCH | XN1000 | 522611 | 1 | CVp 02/04-01/10/2026 | bias=1,330225; ETp=2,630170 (VB MIN); CVp6=1,175173; L29 | (ETp - abs(bias))/CVp 6 m | 1.10617279384 | 1.10617279384 | 3.9e-13 | APROVADO |
| F3 | Sigma | MCH | XN1000 | 522611 | 2 | CVp 02/04-01/10/2026 | CVp6=0,844022; L30 | (ETp - abs(bias))/CVp 6 m | 1.54017908677 | 1.54017908677 | 2.0e-12 | APROVADO |
| F3 | Sigma | MCH | XN1000 | 522611 | 3 | CVp 02/04-01/10/2026 | CVp6=0,828970; L31 | (ETp - abs(bias))/CVp 6 m | 1.56814522097 | 1.56814522097 | 1.6e-13 | APROVADO |
| F3 | Sigma | MCHC | XN1000 | 522611 | 1 | CVp 02/04-01/10/2026 | bias=1,606529; ETp=1,737874 (VB MIN); CVp6=1,475015; L32 | (ETp - abs(bias))/CVp 6 m | 0.0890465264878 | 0.0890465264876 | 1.5e-12 | APROVADO |
| F3 | Sigma | MCHC | XN1000 | 522611 | 2 | CVp 02/04-01/10/2026 | CVp6=1,245132; L33 | (ETp - abs(bias))/CVp 6 m | 0.105486745022 | 0.105486745022 | 1.3e-13 | APROVADO |
| F3 | Sigma | MCHC | XN1000 | 522611 | 3 | CVp 02/04-01/10/2026 | CVp6=1,181613; L34 | (ETp - abs(bias))/CVp 6 m | 0.111157337079 | 0.111157337079 | 1.1e-12 | APROVADO |
| F3 | Sigma | NEUT% | XN1000 | 522611 | 1 | CVp 02/04-01/10/2026 | bias=2,465174; ETp=27,582810 (VB MIN); CVp6=2,632297; L38 | (ETp - abs(bias))/CVp 6 m | 9.54209706918 | 9.54209706918 | 2.6e-13 | APROVADO |
| F3 | Sigma | NEUT% | XN1000 | 522611 | 2 | CVp 02/04-01/10/2026 | CVp6=1,935995; L39 | (ETp - abs(bias))/CVp 6 m | 12.9740209644 | 12.9740209644 | 4.4e-13 | APROVADO |
| F3 | Sigma | NEUT% | XN1000 | 522611 | 3 | CVp 02/04-01/10/2026 | CVp6=1,686134; L40 | (ETp - abs(bias))/CVp 6 m | 14.8965847837 | 14.8965847838 | 2.1e-13 | APROVADO |
| F3 | Sigma | MONO% | XN1000 | 522611 | 1 | CVp 02/04-01/10/2026 | bias=6,491155; ETp=26,163428 (VB MIN); CVp6=12,175156; L44 | (ETp - abs(bias))/CVp 6 m | 1.61577168492 | 1.61577168492 | 1.7e-14 | APROVADO |
| F3 | Sigma | MONO% | XN1000 | 522611 | 2 | CVp 02/04-01/10/2026 | CVp6=6,415953; L45 | (ETp - abs(bias))/CVp 6 m | 3.06614959368 | 3.06614959368 | 1.9e-15 | APROVADO |
| F3 | Sigma | MONO% | XN1000 | 522611 | 3 | CVp 02/04-01/10/2026 | CVp6=4,083639 gl=171; L46 | (ETp - abs(bias))/CVp 6 m | 4.81733867981 | 4.81733867981 | 3.7e-15 | APROVADO |
| F3 | Sigma | LYMPH# | XN1000 | 522611 | 1 | CVp 02/04-01/10/2026 | bias=3,270527; ETp=22,757983 (VB MIN); CVp6=4,988471; L56 | (ETp - abs(bias))/CVp 6 m | 3.90649910152 | 3.90649910152 | 1.0e-14 | APROVADO |
| F3 | Sigma | LYMPH# | XN1000 | 522611 | 2 | CVp 02/04-01/10/2026 | CVp6=2,863348; L57 | (ETp - abs(bias))/CVp 6 m | 6.80582850679 | 6.80582850679 | 3.6e-13 | APROVADO |
| F3 | Sigma | LYMPH# | XN1000 | 522611 | 3 | CVp 02/04-01/10/2026 | CVp6=1,954109; L58 | (ETp - abs(bias))/CVp 6 m | 9.97255507035 | 9.97255507035 | 2.5e-13 | APROVADO |
| F3 | Sigma | EO# | XN1000 | 522611 | 1 | CVp 02/04-01/10/2026 | bias=8,181631; ETp=43,555633 (VB MIN); CVp6=7,081119; L62 | (ETp - abs(bias))/CVp 6 m | 4.99553816283 | 4.99553816283 | 3.1e-14 | APROVADO |
| F3 | Sigma | EO# | XN1000 | 522611 | 2 | CVp 02/04-01/10/2026 | CVp6=7,345442; L63 | (ETp - abs(bias))/CVp 6 m | 4.8157756701 | 4.8157756701 | 1.8e-14 | APROVADO |
| F3 | Sigma | EO# | XN1000 | 522611 | 3 | CVp 02/04-01/10/2026 | CVp6=6,179441; L64 | (ETp - abs(bias))/CVp 6 m | 5.7244665354 | 5.7244665354 | 3.1e-14 | APROVADO |
| F3 | Sigma | MPV | XN1000 | 522611 | 1 | CVp 02/04-01/10/2026 | bias=2,883455; ETp=5,474090 (VB MIN); CVp6=3,053153; L86 | (ETp - abs(bias))/CVp 6 m | 0.848511570987 | 0.848511570987 | 1.0e-13 | APROVADO |
| F3 | Sigma | MPV | XN1000 | 522611 | 2 | CVp 02/04-01/10/2026 | CVp6=1,919028; L87 | (ETp - abs(bias))/CVp 6 m | 1.3499726048 | 1.3499726048 | 3.0e-15 | APROVADO |
| F3 | Sigma | MPV | XN1000 | 522611 | 3 | CVp 02/04-01/10/2026 | CVp6=1,508053; L88 | (ETp - abs(bias))/CVp 6 m | 1.71786781307 | 1.71786781307 | 6.0e-13 | APROVADO |
| F3 | Sigma | RDW-CV | XN1000 | 522611 | 1-3 | CVp 02/04-01/10/2026 | bias=1,023083; fonte CLIA sem ETp na biblioteca 2025 | '-' quando ETp ausente (ADR-068) | - | - | 0 | APROVADO |
| F3 | Sigma do bloco do lote (Analitos BD) | MONO% | XN1000 | 522611 | 3 | cadastro do lote | Média 10,650980; DP 1,170143 (TESTE = 1,5 x DP do lote); bias 6,491155; ETp 26,163428 | (ETp - abs(bias))/(100*DP/Média do cadastro) | 1.791 | 1.791 | <1e-15 | APROVADO |

#### Hematologia — Levey-Jennings e Westgard (6: 6 aprovados, 0 reprovados)

| Frente | Indicador | Analito | Equipamento | Lote | Nível | Período | Entradas | Fórmula | Independente | QC_INI | Diferença | Status |
|---|---|---|---|---|---|---|---|---|---|---|---|---|
| F4 (relatório §2) | LJ + Westgard N=3 por ponto | WBC | XN1000 | 619611 | N1–N3 | Set/2026 | Eng_Saida + cadastro do lote | eixo, pontos, corrida do lote, 7 limites, 5 regras + 1_2s | 171 pontos; limites 3,12056 ± 0,07935 / 6,99471 ± 0,1368 / 16,86122 ± 0,25395; contagens 0 | idêntico (OK/OK/OK) | 0 | APROVADO |
| F4 (relatório §2) | LJ + Westgard N=3 por ponto | WBC | XN1000 | 614011 | N1–N3 | T3/2026 | idem | idem | 176 pontos | idêntico | 0 | APROVADO |
| F4 (relatório §2) | LJ + Westgard N=3 por ponto | PLT | XN1000 | 608511 | N1–N3 | Mai/2026 | idem | idem | 118 pontos; 6x = 6 por nível, REJEITADO ×3 | idêntico | 0 | APROVADO |
| F4 (relatório §2) | LJ + Westgard N=3 por ponto | HGB | XN1000 | 533811 | N1–N3 | Jan/2026 | idem | idem | 159 pontos; 6x = 9 por nível | idêntico | 0 | APROVADO |
| F4 (relatório §2) | LJ + Westgard N=3 por ponto | RBC | XN1000 | 522611 | N1–N3 | Tudo | idem | idem | 152 pontos; N1 1_3s 40, R4s 39 (alvo do cadastro, ver OC-H13) | idêntico | 0 | APROVADO |
| F4 (relatório §2) | LJ + Westgard lote a lote, H3 vazio | WBC | XN1000 | 528211 + 533811 + 536411 + 602911 | N1–N3 | fev–mar/2026 | Eng_Saida | 180 chaves RUN\|LOTE; degraus = Média + k·DP do lote do ponto | 485 pontos; contagens N1 [2,0,2,2,7] | idêntico | 0 | APROVADO |

#### Hematologia — Conferência global (calculadora independente) (5: 5 aprovados, 0 reprovados)

| Frente | Indicador | Analito | Equipamento | Lote | Nível | Período | Entradas | Fórmula | Independente | QC_INI | Diferença | Status |
|---|---|---|---|---|---|---|---|---|---|---|---|---|
| F2 | Estatistica 84x27 + Painel 45 | todos | XN1000 | 522611 | 1-3 | S1 e S2, runs A e B | tblCQ_Final + EQA_Base | calc_independente.py | - | - | 0 divergencias (Painel M47 B-S2 com a tolerancia absoluta da F3) | APROVADO |
| F3 (teste F3-T22) | 2ª implementação (conferencia_cruzada_hema.py) | WBC, RBC, HCT, MCV, PLT | XN1000 | 522611 | N1-N3 | idem | idem | script independente da calculadora | 15 níveis | 15 níveis | máx. 3,8e-12 | APROVADO |
| F3 (teste F3-T24) | 3ª implementação do zero (segunda_opiniao_hema.py) | HGB, MCH, MCHC, NEUT%, MONO%, LYMPH#, EO#, MPV, RDW-CV | XN1000 | 522611 | N1-N3 | idem | idem; ETp recalculado da biblioteca; DPM por erfc | script do zero | 459 valores | 459 valores | 459/459 (máx. rel. 1,6e-10, DPM HGB N3) | APROVADO |
| F3 (teste F3-T26/T27) | Contador de alertas U e bloco do lote (Analitos AU:BD) | todos | XN1000 | 9 lotes | N1-N3 | cadastro do lote | Estatística U14:U133; Analitos AU:BD | definição do próprio bloco (CV do cadastro) | 120 + 2.268 valores | idem | 0 divergências (definição difere da Estatística, OC-H14) | APROVADO |
| F3 (testes F3-T03..T11, T23) | Estatística completa + Painel | todos (28 analitos, 84 linhas) | XN1000 | 522611 | N1-N3 | CIQ ano 2025; CVp 02/04-01/10/2026 | tblCQ_Final + EQA_Base + Especificações | calc_independente.py (C:F, G, H, I:K, L, M:W, AF:AI, AR, AU) | 84 × 27 + 45 células do Painel | idem | 0 divergências (máx. 9,4e-12; DPM com tolerância absoluta 1e-9 em 3 linhas, F3 §4.1) | APROVADO |

### B.3 Os 13 cálculos reprovados (nenhum é erro de fórmula)

| Produto | Cálculo | O que aconteceu | Ocorrência |
|---|---|---|---|
| Bioquímica | L01/L02: Glicose DIM 1, 8977, N2, manual em minúsculas | O Painel troca um resultado real pelo manual (n 24, média 246,5509/245,2268), enquanto a Estatística conta os dois (n 25/26). Nenhuma das duas é a população válida | P-01 |
| Bioquímica | Resumo "Sem EP" (Estatística B103) | Mostra 22, o independente dá 4: as 18 linhas vazias são contadas | P-32 |
| Bioquímica | "Tudo (sem filtro)", Glicose DIM 1, lote 8974 | O Painel dá 140; 89,15; DP 16,968, só da janela de 180 corridas. Todos os dados do lote dão 174; 92,2874; 30,3737 | P-12 |
| Bioquímica | C-HBA1C: ET do Painel sem CEQ | H8 = 2,7740 % = 1,65 × CV. Pelo ADR-068, sem bias não há ET | P-10 |
| Hematologia | WBC N3, manual com xn1000 / wbc | O Painel dá n 21 com um valor trocado; a Estatística dá n 22/23 | P-03 |
| Hematologia | H3 vazio + Ano inteiro 2026 / + Tudo | n 141/155/130, da janela de 180 corridas, contra 337/340/308 e 478/485/447 | P-56 |
| Hematologia | Estatística depois de trocar o lote (E-LOTE) e o ano (E-ANO) | Mostra o 619611 com os números do 522611; 339/339 valores divergem. Depois do recálculo interno, 0 divergências | P-04 |
| Hematologia | ET do Painel do RET% e do RDW-SD sem CEQ | H = 1,65 × CV com bias SEM EP | P-21 |

### B.4 Diferenças de método (não reprovadas; vão ao Especialista em CQ)

- **Seleção das amostras.** O bias do ET/Sigma (`BiasEQ`) inclui amostras com resultado 0 (−100 %) e NAO AVALIADO,
  que a incerteza (`ViesEQ`) exclui. Na Bilirrubina direta, sem essas amostras o bias seria 19,99 % e não 30,77 %, e
  o Sigma N1 seria 0,95 e não 0,43 (P-11, P-74).
- **CVs diferentes.** O ET usa o CV do lote no período e o Sigma usa o CVp de 6 meses. O bloco do lote usa o CV do
  cadastro, e o Eng_Saida guarda um Sigma legado com bias contra o alvo do lote. O resultado são três ou quatro Sigmas
  para o mesmo analito (P-33, P-50, P-76).
- **CVp.** Junta todos os lotes do equipamento, e a janela termina no último resultado de qualquer equipamento
  (MU_Fim global). Fica como dúvida da Bio F3.
- **ETp "CLIA" da Hematologia.** WBC 7 % e PLT 15 % (P-53). No RDW, a especificação está no RDW-SD e o CEQ no RDW-CV
  (P-52).
- **Sinal oposto entre rodadas** reduz o |bias| médio. Um Sigma negativo (|bias| > ETp) é mostrado como número.
  Ambos ficam como dúvidas da Bio F2 e da Hema F2.

### B.5 Caminho do dado (conferido ponto a ponto nas F2/F3)

- **CEQ.** Janela `SalvarCEQ` → tabela do provedor (O/P) → `mEQA.AtualizarEQABase`, que mapeia o nome para o
  canônico (W:Y) e grava Uso e EQA_Base → `mCEQ.BiasEQ` (G, T, AC, AD) → H (ET), L (Sigma), N/O/V/W → Painel B7:I9 e
  J/M45:47. **A correção na tabela para antes da EQA_Base até alguém clicar em Consolidar (P-08, P-18).**
- **CIQ.** tblCQ_Final (PARTICIPA = SIM) → C:F por lote e período → H (ET). Em paralelo, o CVp de 6 meses
  (`mIncerteza.IncertezaCIQ`) → L (Sigma) e a u(Rw) de 12 meses → AF.

---

## Relatório C: ocorrências

**Deduplicação.** As 85 ocorrências das frentes viraram 80, mesclando 5 pares em que a mesma falha apareceu em duas
frentes do mesmo produto:

| ID | Produto | Frentes mescladas |
|---|---|---|
| P-02 | Bio | F1 OC-1 + F3 OC-3 |
| P-10 | Bio | F2 OC-7 + F3 OC-2, mais a confirmação da F4 sem ID novo |
| P-30 | Bio | F2 reprodução da OC-F4-7 + F4 OC-F4-7 |
| P-04 | Hema | F2 OC-HF2-6 + F4 OC-HF4-1 |
| P-21 | Hema | F2 OC-HF2-7 + F3 OC-H2, mais a confirmação da F4 sem ID novo |

**Regras da tabela.**

- A prioridade da mesclada é a maior das duas. A P-04 subiu de Alto para Crítico pela F4.
- Falhas iguais nos dois produtos **não** foram mescladas, como pede a definição. A última coluna aponta o P-xx
  correspondente, porque em vários casos uma correção no código comum (`src/comum`) cobre as duas.
- A ordem é por prioridade, depois por produto (Bio, Hema), depois pela frente.
- Os IDs P-xx são provisórios. O registro oficial em `OCORRENCIAS.md` é do Orquestrador.

### C.1 Tabela

| ID | Produto | Prioridade | Tipo | Módulo | Resumo | Frentes | Evidência (relatório + seção) | Mesmo defeito no outro produto |
|---|---|---|---|---|---|---|---|---|
| P-01 | Bioquímica | **Crítico** | `calculo` | Digitar Resultados -> ATUALIZAR (mEtlCQ.CalcularCorridas Option Compare Binary x mEstatistica.GarantirIndice/EquipFiltro UCase) | Equipamento ou analito em minusculas faz o manual substituir um resultado real no Painel/LJ/Westgard e a Estatistica o conta a mais | F1 | Bioquimica_F1.md §4 OC-9 | P-03 |
| P-02 | Bioquímica | **Crítico** | `dados` | ATUALIZAR -> tblCQ_Final -> Painel/Estatistica / tblCQ_Final -> CVp (Estatística L, Painel C46:C47/I7:I8), u(Rw) AF, plano de CQ | Resultados implausíveis (valores ≤ 0, −1/1, resultado de outro nível) do interfaceamento ficam ATIVO e entram na estatística, no Westgard, no CVp do Sigma e na u(Rw) | F1 + F3 (mesclada) | Bioquimica_F1.md §4 OC-1; Bioquimica_F3.md §5 OC-3 | P-15 |
| P-03 | Hematologia | **Crítico** | `calculo` | Digitar Resultados -> ATUALIZAR (mEtlCQ.CalcularCorridas x mEstatistica.GarantirIndice) | Equipamento ou analito em minúsculas faz o manual tomar o lugar de um resultado real no Painel/LJ, e a Estatística o conta a mais | F1 | Hematologia_F1.md §4 OC-HF1-1 | P-01 |
| P-04 | Hematologia | **Crítico** | `funcional` | Painel H3 / mLotes.TrocarLoteAnalise / Estatistica / Estatística / mEstatistica.AtualizarEstatisticaAba / mLotes.TrocarLoteAnalise | A Estatística não recalcula ao trocar o lote no Painel nem ao mudar Ano De/Até: B4 mostra o lote novo com n/média/DP/CV/ET do lote anterior; só o ATUALIZAR recalcula | F2 + F4 (mesclada) | Hematologia_F2.md §4 OC-HF2-6; Hematologia_F4.md §3 OC-HF4-1 | — |
| P-05 | Bioquímica | **Alto** | `dados` | Digitar Resultados -> mEtlCQ.ParaData/ParaDataHora | Data invalida colada 31/02/2026 vira 03/03/2026 e o manual entra ATIVO | F1 | Bioquimica_F1.md §4 OC-2 | P-16 |
| P-06 | Bioquímica | **Alto** | `funcional` | Digitar Resultados (mIntegracao.PrepararLinha/EquipPadrao=selEquipamento) | EQUIPAMENTO vazio recebe o equipamento do Painel sem aviso; instrucoes nao pedem equipamento | F1 | Bioquimica_F1.md §4 OC-3 | — |
| P-07 | Bioquímica | **Alto** | `funcional` | frmCEQ / mCEQForm.SalvarCEQ / EQA.CAP_Dados | ANALISTA não consegue corrigir nem excluir um CEQ; a janela manda 'edite a linha existente' numa aba toda bloqueada | F2 | Bioquimica_F2.md §4 OC-1 | P-17 |
| P-08 | Bioquímica | **Alto** | `funcional` | EQA.CAP_Dados Z1:Z3 / mEQA.AtualizarEQABase | Correção na tabela só chega ao bias/ET/Sigma depois do 'Consolidar na EQA_Base'; o indicador diz 'base em dia' | F2 | Bioquimica_F2.md §4 OC-2 | P-18 |
| P-09 | Bioquímica | **Alto** | `funcional` | frmCEQ / SalvarCEQ / mEQA mapa W:Y | CEQ com nome fora do mapa, do Controllab ou de analito novo é gravado OK e nunca entra no bias; analito sem CEQ não consegue ganhar o primeiro | F2 | Bioquimica_F2.md §4 OC-3 | P-19 |
| P-10 | Bioquímica | **Alto** | `calculo` | Painel H7:H8 / Painel, INDICADORES POR NÍVEL, H7:H8 | Com bias "SEM EP" (sem CEQ ou provedor sem mapa), o Painel H7:H8 mostra ET = 1,65 × CV (bias tratado como 0), enquanto a Estatística deixa o ET vazio | F2 + F3 (mesclada) | Bioquimica_F2.md §4 OC-7; Bioquimica_F3.md §5 OC-2 + Bioquimica_F4.md §3 "Confirmação (sem ID novo)" (C-HBA1C) | P-21 |
| P-11 | Bioquímica | **Alto** | `duvida_metodo` | Estatística G/H/L (mCEQ.BiasEQ) x AR (mCEQ.ViesEQ) | Bias do ET/Sigma inclui amostras do CEQ com resultado 0 (-100 %) que o projeto exclui da incerteza | F3 | Bioquimica_F3.md §5 OC-1 | P-74 |
| P-12 | Bioquímica | **Alto** | `calculo` | Painel / mEstatistica.AtualizarCalc (janela NK=180) | 'Tudo (sem filtro)' num lote com mais de 180 corridas dá estatística só das 180 mais recentes, e o aviso fala só do gráfico | F4 | Bioquimica_F4.md §3 OC-F4-1 | P-56 |
| P-13 | Bioquímica | **Alto** | `dados` | Recebimento, cadastro automático de lote, Painel H3 vazio | Lote de urina 5659 entra como Glicose sérica e se mistura no H3 vazio | F4 | Bioquimica_F4.md §3 OC-F4-2 | — |
| P-14 | Bioquímica | **Alto** | `dados` | Recebimento / cadastro automático de lote / Painel H3 vazio | Lotes que não são do analito (urina 6353/6354, PCR 8353, 8177, 8794, 8974 na Hemoglobina glicada) ativos em analitos séricos, com alvo criado do próprio valor; no H3 vazio destroem o gráfico | F4 | Bioquimica_F4.md §3 OC-F4-11 | — |
| P-15 | Hematologia | **Alto** | `dados` | Digitar Resultados (RESULTADO decimal >= -1E+307) -> ATUALIZAR/QA | Manual com resultado implausível (-1, 0) entra ATIVO na estatística sem alerta | F1 | Hematologia_F1.md §4 OC-HF1-7 | P-02 |
| P-16 | Hematologia | **Alto** | `dados` | Digitar Resultados -> mEtlCQ.ParaData | Data inválida colada (31/02/2026) vira 03/03/2026 e entra ATIVO no lote | F1 | Hematologia_F1.md §4 OC-HF1-2 | P-05 |
| P-17 | Hematologia | **Alto** | `funcional` | frmCEQ / mCEQForm.SalvarCEQ / EQA.CAP_Dados | ANALISTA nao consegue corrigir nem excluir um CEQ; a janela manda editar uma linha bloqueada | F2 | Hematologia_F2.md §4 OC-HF2-1 | P-07 |
| P-18 | Hematologia | **Alto** | `funcional` | EQA.CAP_Dados Z2:Z3 / mEQA.AtualizarEQABase | Correcao na tabela nao chega ao bias/ET/Sigma ate o Consolidar; Z3 diz 'base em dia' e Z2 fica parado | F2 | Hematologia_F2.md §4 OC-HF2-2 | P-08 |
| P-19 | Hematologia | **Alto** | `funcional` | SalvarCEQ / mEQA.Consolidar (mapa W:Y) | Nome fora do mapa e qualquer CEQ Controllab gravados OK sem aviso e fora do bias; R4 = Controllab zera os 28 analitos | F2 | Hematologia_F2.md §4 OC-HF2-3 | P-09 |
| P-20 | Hematologia | **Alto** | `funcional` | frmCEQ (Analito em texto livre) / mapa W:Y | CEQ do LYMPH# registrado como 'Lymphs' entra calado no bias do LYMPH% e derruba o Sigma dele | F2 | Hematologia_F2.md §4 OC-HF2-4 | — |
| P-21 | Hematologia | **Alto** | `calculo` | Painel H7:H9 | Com bias "SEM EP" (sem CEQ ou R4 = Controllab), o Painel H7:H9 mostra ET = 1,65 × CV (bias tratado como 0), enquanto a Estatística deixa o ET vazio | F2 + F3 (mesclada) | Hematologia_F2.md §4 OC-HF2-7; Hematologia_F3.md §5 OC-H2 + Hematologia_F4.md §3 "Confirmação (sem ID novo)" (C-RET%/C-RDW-SD) | P-10 |
| P-22 | Hematologia | **Alto** | `dados` | Analitos, bloco do lote 522611 (E:J) -> LotesStore -> limites LJ/Westgard | Alvo (Média/DP) do lote 522611 incompatível com os dados em 12 pares (IG# com valores de IG%; RDW-SD N2 = 14,78 fL) | F3 | Hematologia_F3.md §5 OC-H13 | — |
| P-23 | Bioquímica | Médio | `funcional` | Digitar Resultados -> mEtlCQ/mEtlQA | Manual com analito, lote ou equipamento inexistente fica ATIVO e invisivel | F1 | Bioquimica_F1.md §4 OC-4 | P-42 |
| P-24 | Bioquímica | Médio | `dados` | QA_INTEGRACAO (E01) | Quatro inativacoes sem motivo fazem toda atualizacao dizer ERRO DE GOVERNANCA | F1 | Bioquimica_F1.md §4 OC-5 | P-43 |
| P-25 | Bioquímica | Médio | `funcional` | frmCEQ / SalvarCEQ | Registro sem valor-alvo aceito sem aviso; a amostra conta nos limites e não no bias | F2 | Bioquimica_F2.md §4 OC-4 | P-45 |
| P-26 | Bioquímica | Médio | `funcional` | EQA.CAP_Dados / mEQA / Painel M4 | Edição, consolidação e troca de provedor do CEQ sem trilha no Audit_Log | F2 | Bioquimica_F2.md §4 OC-5 | P-46 |
| P-27 | Bioquímica | Médio | `funcional` | mCEQForm.SalvarCEQ / AbrirCEQ | SalvarCEQ grava sem sessão nem papel | F2 | Bioquimica_F2.md §4 OC-6 | P-47 |
| P-28 | Bioquímica | Médio | `usabilidade` | Estatística H4 (=loteAnalise) × Painel H3/H4/B7:I8 | Com H3 vazio (TODOS), a Estatística calcula o lote em uso (8974) e diz 'refletido do Painel' | F2 | Bioquimica_F2.md §4 OC-10 | — |
| P-29 | Bioquímica | Médio | `usabilidade` | GUIA_DE_USO §5-6 | O GUIA §5-6 manda digitar nas abas de dados; não cita a janela, o Consolidar nem a correção | F2 | Bioquimica_F2.md §4 OC-9 | P-49 |
| P-30 | Bioquímica | Médio | `usabilidade` | Painel H3 / filtro da rotina (ADR-077/078) / Painel H3 (ADR-077/078) | Ao voltar de outro lote, o período escolhido não volta quando o lote de partida não é o "lote em uso" | F2 + F4 (mesclada) | Bioquimica_F2.md OC-F4-7 (§2 F2-T37; §4 abertura); Bioquimica_F4.md §3 OC-F4-7 | P-60 |
| P-31 | Bioquímica | Médio | `funcional` | Estatística AT14:AT75 (RMS do viés / u(Cref)) | Coluna AT da Estatística mostra '004 / 000' em vez de '3,69 / 0,xx' | F3 | Bioquimica_F3.md §5 OC-4 | P-51 |
| P-32 | Bioquímica | Médio | `calculo` | Estatística B103 | Resumo 'Sem EP (sem bias -> sem Sigma)' conta 22 níveis, mas são 4 | F3 | Bioquimica_F3.md §5 OC-5 | P-78 |
| P-33 | Bioquímica | Médio | `duvida_metodo` | Estatística H/L, Analitos bloco do lote AU:BD, Eng_Saida 185:186 e 190:269 | Três/quatro Sigmas diferentes para o mesmo analito, e ET x Sigma com CVs diferentes | F3 | Bioquimica_F3.md §5 OC-6 | P-50 |
| P-34 | Bioquímica | Médio | `dados` | Especificações, biblioteca anual (linhas 55-89) | Biblioteca de especificações 2026-2029 sem a VB da Lipase e sem o ETp CLIA da PCR | F3 | Bioquimica_F3.md §5 OC-7 | P-55 |
| P-35 | Bioquímica | Médio | `duvida_metodo` | Painel G7:I8 (lê a aba Estatística) | No H3 vazio o Painel mostra Bias e Sigma, contra a decisão 4 do ADR-086 | F4 | Bioquimica_F4.md §3 OC-F4-3 | P-57 |
| P-36 | Bioquímica | Médio | `usabilidade` | Painel C4, A39 | 'Lançamentos no lote: 8974' e A39 com o alvo do 8974 em telas de outro lote | F4 | Bioquimica_F4.md §3 OC-F4-4 | P-58 |
| P-37 | Bioquímica | Médio | `funcional` | Estatística B4:J5 | Visões MENSAL/TRIMESTRAL da Estatística inalcançáveis e quebradas ('Ano 0', janela vazia = histórico inteiro) | F4 | Bioquimica_F4.md §3 OC-F4-5 | — |
| P-38 | Bioquímica | Médio | `usabilidade` | Painel K3/K4 × G3/G4 | K4 mantém o rótulo antigo depois de digitar De/Até | F4 | Bioquimica_F4.md §3 OC-F4-6 | P-61 |
| P-39 | Bioquímica | Médio | `usabilidade` | Gráficos do Painel | LJ com linha suavizada que passa dos pontos, sem valores no eixo Y e sem rótulo nas linhas de DP | F4 | Bioquimica_F4.md §3 OC-F4-8 | P-62 |
| P-40 | Bioquímica | Médio | `usabilidade` | Painel H3 (=lstLotes) e ajuste ADR-077/078 | Lista do H3 com os 22 lotes de todos os analitos; escolher lote de outro analito troca o analito sozinho | F4 | Bioquimica_F4.md §3 OC-F4-12 | — |
| P-41 | Bioquímica | Médio | `usabilidade` | Painel S7:T8 = engPainel col 19/20 = AgregadoWestgard (Eventos_Westgard) | 'Últ. violação' e 'Classificação' na tabela 'violações no período' vêm do histórico inteiro do lote do Painel (no H3 vazio, do lote em uso), não do período | F4 | Bioquimica_F4.md §3 OC-F4-13 | P-59 |
| P-42 | Hematologia | Médio | `funcional` | Digitar Resultados -> ATUALIZAR/QA | Manual com analito ou lote inexistente fica ATIVO sem aparecer no Painel (equipamento inexistente gera A10) | F1 | Hematologia_F1.md §4 OC-HF1-3 | P-23 |
| P-43 | Hematologia | Médio | `usabilidade` | mIntegracao.ResumoQA (mensagem final do ATUALIZAR) | Linha manual incompleta (E04) faz o ATUALIZAR dizer 'Há ERRO DE GOVERNANÇA' | F1 | Hematologia_F1.md §4 OC-HF1-8 | P-24 |
| P-44 | Hematologia | Médio | `usabilidade` | Painel J7:J9 (mEstatistica.AtualizarPainelEng) | Status do Painel fica REJEITADO nos 3 níveis por regra 6x de 09/10 mesmo após corridas OK | F1 | Hematologia_F1.md §4 OC-HF1-4 | — |
| P-45 | Hematologia | Médio | `funcional` | frmCEQ / SalvarCEQ | Registro sem valor-alvo aceito sem aviso; conta nos limites (S) e nao no bias (AC) | F2 | Hematologia_F2.md §4 OC-HF2-5 | P-25 |
| P-46 | Hematologia | Médio | `funcional` | EQA.CAP_Dados / AtualizarEQABase / Estatistica R4 | Correcao, consolidacao e troca do provedor global R4 sem rastro no Audit_Log | F2 | Hematologia_F2.md §4 OC-HF2-9 | P-26 |
| P-47 | Hematologia | Médio | `funcional` | mCEQForm.SalvarCEQ / AbrirCEQ | SalvarCEQ grava sem sessao nem papel | F2 | Hematologia_F2.md §4 OC-HF2-10 | P-27 |
| P-48 | Hematologia | Médio | `dados` | tblEQA_CAP_Dados coluna A / mEQA.Consolidar | 90 linhas 'Controllab' (simulacao) na tabela do CAP sao gravadas como CAP na EQA_Base | F2 | Hematologia_F2.md §4 OC-HF2-11 | — |
| P-49 | Hematologia | Médio | `usabilidade` | GUIA_DE_USO §5-6 | GUIA §5-6 manda digitar nas abas bloqueadas e nao cita a janela, o Consolidar nem a correcao | F2 | Hematologia_F2.md §4 OC-HF2-12 | P-29 |
| P-50 | Hematologia | Médio | `duvida_metodo` | Analitos AU:BD (todos os blocos de lote) | Bloco do lote (Analitos AU:BD) mostra outro ET e outro Sigma para o mesmo lote; em 21 pares com DP de TESTE | F3 | Hematologia_F3.md §5 OC-H14 | P-33 |
| P-51 | Hematologia | Médio | `funcional` | Estatística AT14:AT97 | Coluna AT da Estatística mostra '002 / 000' em vez de '2,49 / 0,xx' | F3 | Hematologia_F3.md §5 OC-H1 | P-31 |
| P-52 | Hematologia | Médio | `duvida_metodo` | Especificações x EQA_Base (de/para) x Estatística | RDW avaliado no CAP como RDW-CV, mas a especificação VB está só no RDW-SD; nenhum dos dois tem Sigma | F3 | Hematologia_F3.md §5 OC-H3 | — |
| P-53 | Hematologia | Médio | `duvida_metodo` | Especificações, biblioteca 2025 linha 48 | ETp 'CLIA' 2025 de WBC (7 %) e PLT (15 %) diferente dos limites CLIA publicados | F3 | Hematologia_F3.md §5 OC-H4 | — |
| P-54 | Hematologia | Médio | `usabilidade` | Estatística B3/D3 x B4 -> C:F, H, N:Q | Filtro Ano De/Até da Estatística não acompanha o lote e corta lotes que atravessam o ano | F3 | Hematologia_F3.md §5 OC-H5 | — |
| P-55 | Hematologia | Médio | `dados` | Especificações linhas 55-89 | Biblioteca de especificações 2026-2029 vazia na Hematologia | F3 | Hematologia_F3.md §5 OC-H9 | P-34 |
| P-56 | Hematologia | Médio | `calculo` | Painel B7:B9 e M7:R9 / AtualizarCalc (NK = 180, modo TODOS) | H3 vazio com período longo: n e contagens de Westgard só das 180 corridas mais recentes; o aviso fala só do gráfico | F4 | Hematologia_F4.md §3 OC-HF4-2 | P-12 |
| P-57 | Hematologia | Médio | `duvida_metodo` | Painel G7:I9 | No H3 vazio, o Painel mostra Bias % e Sigma (e o SIGMA DO PLANO), contra o ADR-086 dec. 4 | F4 | Hematologia_F4.md §3 OC-HF4-3 | P-35 |
| P-58 | Hematologia | Médio | `usabilidade` | Painel C4/A39; Configuração C20 | Lote em uso 522611 (validade 31/03/2026) em C4 e no A39 do H3 vazio; o lote corrente é o 625311 | F4 | Hematologia_F4.md §3 OC-HF4-4 | P-36 |
| P-59 | Hematologia | Médio | `usabilidade` | Painel S7:T9 (AgregadoWestgard) | 'Últ. violação/Classificação' sob 'violações no período' mostra uma violação fora do período, com total 0 ao lado | F4 | Hematologia_F4.md §3 OC-HF4-5 | P-41 |
| P-60 | Hematologia | Médio | `usabilidade` | Painel H3 (ADR-077/078) | Voltar de outro lote não devolve o período quando o lote de partida não é o lote em uso | F4 | Hematologia_F4.md §3 OC-HF4-6 | P-30 |
| P-61 | Hematologia | Médio | `usabilidade` | Painel K3/K4 × G3/G4 | Depois de digitar De/Até, Ano (K3) e Período (K4) mantêm o rótulo antigo | F4 | Hematologia_F4.md §3 OC-HF4-7 | P-38 |
| P-62 | Hematologia | Médio | `usabilidade` | Gráficos do Painel | LJ com linha suavizada, sem valores no eixo Y e sem rótulo de ±k DP; no H3 vazio, os degraus viram curvas | F4 | Hematologia_F4.md §3 OC-HF4-8 | P-39 |
| P-63 | Bioquímica | Baixo | `usabilidade` | Digitar Resultados (XLOOKUP) | Colunas RUN e DETALHE da aba Digitar Resultados mostram 0 | F1 | Bioquimica_F1.md §4 OC-6 | P-72 |
| P-64 | Bioquímica | Baixo | `usabilidade` | QA_INTEGRACAO (I01) | I01 lista NA/K/CL/CRE2/PHOS sem dizer que sao da URINA | F1 | Bioquimica_F1.md §4 OC-7 | — |
| P-65 | Bioquímica | Baixo | `funcional` | Audit_Log (mIntegracao) | Audit_Log grava RESULTADO_MANUAL_INCLUIDO para linhas incompletas e em conflito | F1 | Bioquimica_F1.md §4 OC-8 | P-71 |
| P-66 | Bioquímica | Baixo | `usabilidade` | frmCEQ / DistintosCEQ | Listas da janela misturam a simulação (A, B, C; nomes em português) com o CAP real | F2 | Bioquimica_F2.md §4 OC-8 | P-73 |
| P-67 | Bioquímica | Baixo | `usabilidade` | Painel J7:J8 | 'REJEITADO' cortado em J7:J8 | F4 | Bioquimica_F4.md §3 OC-F4-9 | — |
| P-68 | Bioquímica | Baixo | `usabilidade` | Estatística | Nome do analito desalinhado na vertical na Estatística | F4 | Bioquimica_F4.md §3 OC-F4-10 | — |
| P-69 | Bioquímica | Baixo | `usabilidade` | Gráfico N2 do Painel, modo TODOS (moldura pelos degraus) | Ponto de violação cortado na borda superior do LJ (y_max = valor do ponto) | F4 | Bioquimica_F4.md §3 OC-F4-14 | — |
| P-70 | Hematologia | Baixo | `usabilidade` | Digitar Resultados, coluna MOTIVO | MOTIVO só aceita os 4 itens da lista; 'Outro' sem complemento | F1 | Hematologia_F1.md §4 OC-HF1-9 | — |
| P-71 | Hematologia | Baixo | `funcional` | mIntegracao (Audit_Log) | Audit_Log grava RESULTADO_MANUAL_INCLUIDO para linhas incompletas/em conflito | F1 | Hematologia_F1.md §4 OC-HF1-5 | P-65 |
| P-72 | Hematologia | Baixo | `usabilidade` | Digitar Resultados (fórmulas XLOOKUP RUN/DETALHE) | Colunas RUN e DETALHE da aba Digitar Resultados mostram 0 | F1 | Hematologia_F1.md §4 OC-HF1-6 | P-63 |
| P-73 | Hematologia | Baixo | `usabilidade` | mCEQForm.DistintosCEQ | Listas da janela misturam nomes da simulacao (RBC, rodadas 1/2/3, OK) com os do CAP real | F2 | Hematologia_F2.md §4 OC-HF2-8 | P-66 |
| P-74 | Hematologia | Baixo | `duvida_metodo` | Estatística G/H (BiasEQ) x AR/AU (ViesEQ) | Bias do ET inclui amostras NAO AVALIADO que a incerteza exclui (IG%, IG#, NRBC%, NRBC#) | F3 | Hematologia_F3.md §5 OC-H6 | P-11 |
| P-75 | Hematologia | Baixo | `dados` | tblCQ_Final -> AF e L | Resultados isolados de 3-4 DP ativos pesam na u(Rw) da HGB N3 (+20 %) | F3 | Hematologia_F3.md §5 OC-H7 | — |
| P-76 | Hematologia | Baixo | `duvida_metodo` | Eng_Saida I185:I187 e L190:L309 | Sigma legado com bias contra o alvo do lote continua gravado no Eng_Saida | F3 | Hematologia_F3.md §5 OC-H8 | — |
| P-77 | Hematologia | Baixo | `dados` | Especificações linha 53 | 'FAB CVTp' 2025 com valores que não parecem CV % (RDW-SD 48,1; RDW-CV 15,1; MPV 9,4; RET% 1,2) | F3 | Hematologia_F3.md §5 OC-H10 | — |
| P-78 | Hematologia | Baixo | `usabilidade` | Estatística A98:T133 | Linhas vazias da Estatística mostram '0' e 'SEM EP' | F3 | Hematologia_F3.md §5 OC-H11 | P-32 |
| P-79 | Hematologia | Baixo | `usabilidade` | Principal - Resultados A5 | Cabeçalho da tblCQ_Final mostra contagens com decimal | F3 | Hematologia_F3.md §5 OC-H12 | — |
| P-80 | Hematologia | Baixo | `usabilidade` | Estatística B3/D3 | Ano Até menor que Ano De aceito sem aviso; a aba usa só o Ano De | F4 | Hematologia_F4.md §3 OC-HF4-9 | — |

### C.2 Detalhe das ocorrências Críticas e Altas

**P-01 (Crítico, Bioquímica, `calculo`): Equipamento ou analito em minusculas faz o manual substituir um resultado real no Painel/LJ/Westgard e a Estatistica o conta a mais**

- Observado: ETL da ao manual o RUN da 1a corrida real do dia (26102901 / 26102801); o motor junta em maiusculas e, no Eng_Saida (um valor por RUN e nivel), o manual 244,222 substitui o N2 real BIO-371977 (277) / BIO-371038 (276): Painel n 24 media 246,5509 / 245,2268; Estatistica conta os dois (n 25 media 247,7689 / n 26 media 247,6325); vereditos de Westgard do RUN real OK->REJEITADO; ET da linha de Sigma muda 53,69->52,65->51,67; nenhum QA. Mesmo defeito da OC-HF1-1 da Hema
  - Consequência: Estatistica incorreta e divergente entre Painel e Estatistica, LJ com ponto trocado e data errada, Westgard avaliado sobre valor errado numa corrida real
  - Sugestão do relatório: Normalizar EQUIPAMENTO/ANALITO do manual para o cadastro (PrepararLinha ou mEtlCQ, VBA e PQ) ou mesma comparacao no ETL e no motor; motor nunca sobrescrever valor no mesmo RUN/nivel sem alerta. Uma correcao no codigo comum cobre OC-HF1-1 e esta

**P-02 (Crítico, Bioquímica, `dados`): Resultados implausíveis (valores ≤ 0, −1/1, resultado de outro nível) do interfaceamento ficam ATIVO e entram na estatística, no Westgard, no CVp do Sigma e na u(Rw)**

- [F1] Observado: BIO-371183=-1 e BIO-371184=1 ATIVO: N2 CV 30,88% (2,14% sem eles), ET 53,69%, REJEITADO; 56 ativos <=0 na producao (43 negativos, 13 zeros, 55 de analitos cadastrados, desde 12/2025)
  - Consequência: Estatistica, ET, Sigma e Westgard incorretos ate alguem inativar
  - Sugestão do relatório: Triagem de plausibilidade no QA (ALERTA com ID); laboratorio revisar os 56 IDs; consultar SIPEC sobre -1/0/1 sem flag
- [F3] Observado: 253/253 suspeitos de 13/10 continuam ativos (PARTICIPA=SIM); 287 na janela de 6 m. Ex.: Glicose N1 BIO-319726 = 277 (lote 8977) -> Sigma 0,33 em vez de 4,32; Proteína total BIO-232288 = 101,5 e BIO-268082 = -0,5 -> u(Rw) 12 m = 81,2 %; Amilase BIO-215815 = 496 em N1. 53 de 58 níveis 'Desempenho inadequado'
  - Consequência: Plano de CQ recomendado sobre Sigma artificial em quase todos os analitos; incerteza superestimada
  - Sugestão do relatório: Laboratório revisar e inativar os IDs pela aba Inativar; com autorização, alerta no QA_INTEGRACAO para nível trocado ou valor <= 0

**P-03 (Crítico, Hematologia, `calculo`): Equipamento ou analito em minúsculas faz o manual tomar o lugar de um resultado real no Painel/LJ, e a Estatística o conta a mais**

- Observado: RUN da 1ª corrida real do dia; o Painel/LJ troca o N3 real (17,58 HEM-371290; 17,14 HEM-370716) pelo manual 17,433 (n 21); a Estatística conta os dois (n 22/23); sem QA. Reproduzido bit a bit no 2º passe; Range.Validation.Value confirma que 'wbc' passa digitado (lstAnalitos não diferencia maiúsculas) e 'xn1000' é recusado (lista literal)
  - Consequência: estatística divergente, LJ com ponto trocado e data errada, Westgard sobre o valor errado
  - Sugestão do relatório: normalizar EQUIPAMENTO/ANALITO do manual no ETL (o ETL já apara espaços, F02) ou usar a mesma comparação no ETL e no motor; o motor não sobrescreve valor do mesmo RUN/nível sem alerta

**P-04 (Crítico, Hematologia, `funcional`): A Estatística não recalcula ao trocar o lote no Painel nem ao mudar Ano De/Até: B4 mostra o lote novo com n/média/DP/CV/ET do lote anterior; só o ATUALIZAR recalcula**

- [F2] Observado: H3 522611 -> 619611: B4 = 619611, mas 0 celulas mudaram (RBC N1 n 50, CV 1,2595, ET 4,478 do lote 522611), enquanto o Painel mostra o 619611 (n 54, CV 1,1344, ET 4,272). TrocarLoteAnalise chama RecalcularAnalitoAtual, nao AtualizarEstatisticaAba
  - Consequência: Tabela x Painel divergentes; numeros de um lote com o rotulo de outro
  - Sugestão do relatório: Atualizar a aba na troca de lote, ou aviso de desatualizada; confirmar na F4
- [F4] Observado: H3 = 619611: B4 = 619611 (refletido do Painel), mas WBC N1 = 50; 3,0446; DP 0,054929; CV 1,8041 %, idênticos ao 522611/2025 (o 619611 não tem nenhum resultado em 2025). Mudar B3/D3 para 2026 e entrar pelo botão não mudam nada. O módulo da planilha Estatística está vazio; só o ATUALIZAR (AtualizarEstatistica) recalcula. Depois do recálculo interno, 0 divergências
  - Consequência: Estatística mostrada sob o lote errado; o ET %, a margem e o Status CV da aba usam o CV de outro lote; relatório errado. O Painel não é afetado
  - Sugestão do relatório: Recalcular no TrocarLoteAnalise e num Worksheet_Change de B3:D3, ou mostrar DESATUALIZADO quando lote/anos diferirem do último cálculo

**P-05 (Alto, Bioquímica, `dados`): Data invalida colada 31/02/2026 vira 03/03/2026 e o manual entra ATIVO**

- Observado: MAN_0015 DATA_HORA 03/03/2026 08:00 ATIVO RUN 26030301, sem aviso; Westgard retroativo do lote muda
  - Consequência: Resultado em data errada num mes fechado
  - Sugestão do relatório: Conferir dia/mes/ano apos DateSerial; MANUAL_INCOMPLETO se divergir (tambem no PQ)

**P-06 (Alto, Bioquímica, `funcional`): EQUIPAMENTO vazio recebe o equipamento do Painel sem aviso; instrucoes nao pedem equipamento**

- Observado: MAN_0016 ganha DIMENSION 2 e entra ATIVO na serie errada; sem mensagem
  - Consequência: Associacao errada entre equipamentos
  - Sugestão do relatório: Incluir EQUIPAMENTO nas instrucoes e GUIA; torna-lo obrigatorio na Bio ou destacar

**P-07 (Alto, Bioquímica, `funcional`): ANALISTA não consegue corrigir nem excluir um CEQ; a janela manda 'edite a linha existente' numa aba toda bloqueada**

- Observado: Reproduzida no 3º passe (Proteína total): ERRO|5|... edite a linha existente; ProtectContents=Verdadeiro e célula F547 Locked
  - Consequência: CEQ errado fica no bias/ET/Sigma até um ADM corrigir
  - Sugestão do relatório: Corrigir/Excluir na janela ou desbloquear as colunas digitáveis com consolidação automática

**P-08 (Alto, Bioquímica, `funcional`): Correção na tabela só chega ao bias/ET/Sigma depois do 'Consolidar na EQA_Base'; o indicador diz 'base em dia'**

- Observado: Reproduzida: O=5 na tabela, mas EQA_Base 110/10 e 0 células mudam na Estatística; o Z2 continua com o carimbo de 22/08 mesmo depois de consolidar, enquanto o Z1 da EQA_Base mostra 01/11 02:11:45
  - Consequência: ET e Sigma calculados com um CEQ já corrigido
  - Sugestão do relatório: Consolidar ao editar/ATUALIZAR; Z3 comparar conteúdo; Z2 volátil

**P-09 (Alto, Bioquímica, `funcional`): CEQ com nome fora do mapa, do Controllab ou de analito novo é gravado OK e nunca entra no bias; analito sem CEQ não consegue ganhar o primeiro**

- Observado: 'Protein, total' e Controllab 'Proteína total' com Canônico vazio. Novo: HbA1c pelo nome do CAP ('Hemoglobin A1c') e pelo nome canônico ('Hemoglobina glicada') → OK|551| e OK|552|; MapearAnalito='' para os dois; HbA1c continua SEM EP sem aviso
  - Consequência: Analito novo ou provedor Controllab fica sem bias/ET/Sigma em silêncio
  - Sugestão do relatório: Conferir MapearAnalito no SalvarCEQ e avisar; incluir o Controllab e uma tela de mapeamento (decisão do Orquestrador)

**P-10 (Alto, Bioquímica, `calculo`): Com bias "SEM EP" (sem CEQ ou provedor sem mapa), o Painel H7:H8 mostra ET = 1,65 × CV (bias tratado como 0), enquanto a Estatística deixa o ET vazio**

- [F2] Observado: Reproduzida na Proteína total: M4=Controllab → Painel H7:H8 3,069/3,632 % com bias 'SEM EP'; Estatística H vazia; ET real com CAP 6,85/7,42 %
  - Consequência: ET subestimado aparenta desempenho ótimo
  - Sugestão do relatório: Deixar H7:H8 vazio quando G não for número
- [F3] Observado: H7 =IF($E7="","",$E7*1.65+ABS(IF(ISNUMBER($G7),$G7,0))): com G7='SEM EP' o ET vira 1,65*CV (HbA1c 2,54 % / 2,78 %; PCR 2,37 % / 2,01 %), enquanto C63 fica vazio
  - Consequência: ET subestimado exibido para analitos sem CEQ; dois ETs diferentes na mesma tela
  - Sugestão do relatório: ET do Painel seguir a Estatística (vazio/SEM EP quando G não for número)

**P-11 (Alto, Bioquímica, `duvida_metodo`): Bias do ET/Sigma inclui amostras do CEQ com resultado 0 (-100 %) que o projeto exclui da incerteza**

- Observado: CHM-04 (C-A) e CHM-13 (C-C) = 0,0 contra alvos 0,05/0,03. BiasEQ = 30,77 % (com elas); ViesEQ exclui (n=13). Sem elas o bias seria 19,99 %, ET N1 26,3 % em vez de 37,1 % e Sigma N1 0,95 em vez de 0,43
  - Consequência: Bias, ET e Sigma distorcidos por critério que o próprio projeto considera inadequado em outra coluna
  - Sugestão do relatório: Especialista em CQ decidir se BiasEQ ABS aplica o filtro do ViesEQ e se deve haver regra de resolução/limite inferior

**P-12 (Alto, Bioquímica, `calculo`): 'Tudo (sem filtro)' num lote com mais de 180 corridas dá estatística só das 180 mais recentes, e o aviso fala só do gráfico**

- Observado: Glicose DIM1 8974: Painel N1 140; 89,15; DP 16,968 contra todos os dados 174; 92,2874; 30,3737. A39: '39 corrida(s) ... fora do gráfico'
  - Consequência: n/média/DP/CV de 'Tudo' não são do lote inteiro
  - Sugestão do relatório: Calcular sobre o período inteiro, ou dizer 'fora do gráfico e da estatística' com o n total

**P-13 (Alto, Bioquímica, `dados`): Lote de urina 5659 entra como Glicose sérica e se mistura no H3 vazio**

- Observado: 28 linhas GLUC MATRIZ=SORO ativas no lote 5659 (N1 55,9; N2 33,3), alvo com DP copiado do 8977; no H3 vazio junta com o 8977
  - Consequência: n e Westgard inflados no H3 vazio; gráfico ilegível
  - Sugestão do relatório: Laboratório inativar/corrigir; trava de plausibilidade por matriz/lote

**P-14 (Alto, Bioquímica, `dados`): Lotes que não são do analito (urina 6353/6354, PCR 8353, 8177, 8794, 8974 na Hemoglobina glicada) ativos em analitos séricos, com alvo criado do próprio valor; no H3 vazio destroem o gráfico**

- Observado: Ureia DIM1 lote 6353: BIO-230952=871 e BIO-230958=1830 (MATRIZ=SORO, ativas), alvo 871±1,7094 / 1830±4,8; no H3 vazio dez/2025 eixo Y até 918 (N1) e 1.932 (N2). Creatinina 6354 N2 ~236; Na/K/Cl/P/Ca em 6353/6354; lote 8353 com Glicose −1, Ureia −1 em 24/08/2026; 14 linhas de Hemoglobina glicada no lote 8974
  - Consequência: H3 vazio inutilizável nesses períodos; estatística/alvo falsos com nome de analito sérico. Com o lote escolhido em H3 não há efeito
  - Sugestão do relatório: Laboratório conferir e inativar; trava de plausibilidade (matriz, alvo com 1 ponto, distância ao alvo do analito); aviso no modo TODOS

**P-15 (Alto, Hematologia, `dados`): Manual com resultado implausível (-1, 0) entra ATIVO na estatística sem alerta**

- Observado: a validação aceita; ATIVO, PARTICIPA=SIM, corrida própria, nenhum código de QA; CV N1 1,736% -> 30,004%, N2 1,548% -> 16,292%; só o Westgard 1-3s REJEITADO sinaliza
  - Consequência: média/DP/CV (e ET/Sigma) errados até alguém inativar; mesma família da OC-1 da Bio
  - Sugestão do relatório: triagem de plausibilidade (<= 0 ou |z| extremo) com ALERTA/MANUAL_INCOMPLETO, e validação RESULTADO > 0 na tela; decisão do Especialista em CQ/RT

**P-16 (Alto, Hematologia, `dados`): Data inválida colada (31/02/2026) vira 03/03/2026 e entra ATIVO no lote**

- Observado: MAN_0016 03/03/2026 08:00 ATIVO RUN 26030301, N1 n 22->23, sem aviso; digitada, a validação recusa (Validation.Value False)
  - Consequência: resultado num mês errado da série do lote
  - Sugestão do relatório: conferir dia/mês/ano depois de converter, nos dois motores

**P-17 (Alto, Hematologia, `funcional`): ANALISTA nao consegue corrigir nem excluir um CEQ; a janela manda editar uma linha bloqueada**

- Observado: ERRO|5 'edite a linha existente', mas a aba esta protegida e as 24 colunas Locked; a janela nao tem editar nem excluir
  - Consequência: CEQ errado fica no bias/ET/Sigma ate um ADM corrigir, ou vira duplicata
  - Sugestão do relatório: Corrigir/Excluir na janela, ou desbloquear A..N/Q/R e consolidar ao sair

**P-18 (Alto, Hematologia, `funcional`): Correcao na tabela nao chega ao bias/ET/Sigma ate o Consolidar; Z3 diz 'base em dia' e Z2 fica parado**

- Observado: O467 = 5, EQA_Base 110/10, 0 celulas na Estatistica (G 3,6502 em vez de 2,4002); Z3 'base em dia'; Z2 = carimbo de 24/08 depois de 7 consolidacoes
  - Consequência: ET e Sigma com um CEQ ja corrigido, sem sinal na tela
  - Sugestão do relatório: Consolidar ao editar ou no ATUALIZAR; Z3 por conteudo; Z2 volatil

**P-19 (Alto, Hematologia, `funcional`): Nome fora do mapa e qualquer CEQ Controllab gravados OK sem aviso e fora do bias; R4 = Controllab zera os 28 analitos**

- Observado: 'Red Blood Cell' e Controllab 'RBC' com Canonico vazio; o mapa tem 31 entradas, todas do CAP; com R4 = Controllab todos os analitos ficam SEM EP
  - Consequência: Bias, ET e Sigma somem em silencio (na Hema, o produto inteiro)
  - Sugestão do relatório: Conferir MapearAnalito no SalvarCEQ e avisar; cadastrar o mapa do Controllab

**P-20 (Alto, Hematologia, `funcional`): CEQ do LYMPH# registrado como 'Lymphs' entra calado no bias do LYMPH% e derruba o Sigma dele**

- Observado: OK|469| sem aviso; Canonico = LYMPH%; G LYMPH% 3,63 -> 15,22; Sigma N1 4,33 -> 1,70 (Bom -> Desempenho inadequado); unidade x 10E9/L aceita para um analito em %
  - Consequência: Associacao errada entre analitos no bias/ET/Sigma e no plano de CQ do vizinho (7 pares desse tipo no mapa)
  - Sugestão do relatório: Mostrar o canonico (MapearAnalito) ao lado do campo; avisar quando a unidade divergir da ja registrada; lista fechada opcional

**P-21 (Alto, Hematologia, `calculo`): Com bias "SEM EP" (sem CEQ ou R4 = Controllab), o Painel H7:H9 mostra ET = 1,65 × CV (bias tratado como 0), enquanto a Estatística deixa o ET vazio**

- [F2] Observado: R4 = Controllab: Painel bias SEM EP e ET 2,078/1,576/1,519 % (RBC) e 7,205/6,990/5,638 % (LYMPH#), = 1,65 x CV; Estatistica H e L vazios
  - Consequência: ET subestimado que esconde a falta de CEQ
  - Sugestão do relatório: Formula do Painel igual a da Estatistica
- [F3] Observado: H7 = IF($E7="","",$E7*1.65+ABS(IF(ISNUMBER($G7),$G7,0))); RET% mostraria 5,60/6,50/7,68 % e RDW-SD 1,69/1,43/2,91 %
  - Consequência: ET subestimado exibido em 6 níveis
  - Sugestão do relatório: ET do Painel seguir a Estatística (vazio/SEM EP quando G não é número)

**P-22 (Alto, Hematologia, `dados`): Alvo (Média/DP) do lote 522611 incompatível com os dados em 12 pares (IG# com valores de IG%; RDW-SD N2 = 14,78 fL)**

- Observado: IG# N1-N3 alvo 10,29/11,05/11,88 vs médias 0,304/0,780/2,009 (z -28,6/-31,6/-29,5; alvo tem a ordem do IG%); RDW-SD N2 alvo 14,7793 vs 47,38 (z +398); RDW-CV N2 +25,7; HGB N2 -14,6; MCH N2 -12,4; MCV N2 -7,3; RDW-CV N3 -6,5; HCT N2 -5,7; RET# N1 +5,6; RET% N1 +5,2. Outros 8 lotes: 0 pares. O aviso de TESTE (ADR-076, 21 pares) não cobre esses 12
  - Consequência: Limites do LJ e Westgard deslocados: Eventos_Westgard do lote 522611 (6.155 eventos) dominada pelos afetados (IG# 714 eventos, 152 de 1_3s; RDW-SD 661; RDW-CV 609; RET# 606; RET% 595; MCV 428; HGB 413; MCH 408; HCT 350). Estatística C:W não é afetada
  - Sugestão do relatório: Laboratório conferir a bula do lote 522611 e corrigir pela aba Analitos (com trilha); avaliar aviso de plausibilidade (média do lote a > k DP do alvo após n>=20). Confirmação visual no LJ com a F4

---

## Relatório D: parecer final

Critério: "validado" só onde houve teste executado e aprovado. O resto é "não validado" (teste reprovado) ou "não
executado" (com o motivo).

### D.1 A rotina funciona?

**Parcialmente. O caminho comum funciona; há 4 falhas críticas.**

**Validado nos dois produtos:**

- o recebimento automático do DB_SEAC: nada perdido nem duplicado, ID `BIO-`/`HEM-<id>` preservado, associação exata
  em 10 campos, idempotência, histórico intacto (F1 A01–A05);
- o manual válido com carimbo, gravação e consulta (M01);
- a atualização do Painel, do LJ e do Westgard (M03, G-*);
- o SUBSTITUI_ID (S01–S05);
- o registro do CEQ pela janela com o nome certo (Bio F2-T34, Hema HF2-T05/T06).

**Não validado.** A rotina não está pronta sem ressalvas por causa de:

- P-01 e P-03: manual em minúsculas troca um resultado real;
- P-02: resultados implausíveis ativos na Bio;
- P-04: Estatística da Hema parada no lote anterior;
- 18 ocorrências Altas, a maior parte no CEQ e na entrada.

### D.2 Bias, ET e Sigma estão validados?

**O cálculo está validado. A exibição e alguns casos de borda não estão.**

**Validado:**

- o bias do CEQ (abs, com sinal, n de amostras e de rodadas), o ET, o ETp/CVTp do cadastro (CLIA e VB), o CVp de
  6 meses, o Sigma, a margem, a classe, o DPM e a u(Rw);
- conferência: Bio 62 × 19 = 1.178 comparações com 0 divergências, mais 96/96 e 56/56 de segundas opiniões; Hema
  84 × 27 e 45 células do Painel com 0 divergências, mais 459/459 numa terceira implementação do zero;
- maior diferença: 1,9e-11 relativo, e 1,6e-10 no DPM por erfc;
- inclusão e correção de CEQ mudam bias, ET e Sigma só do analito certo, com diferença 0 em relação ao independente
  (Bio Glicose, Bilirrubina direta e Proteína total; Hema RBC e LYMPH#);
- sinal e escala percentual (Bio F3-T11, Hema F3-T12);
- provedor, ano e rodada certos (Bio F3-T12, Hema F3-T13);
- 2 lotes no CVp (Bio F3-T15, Hema F3-T16);
- ETp ausente aparece como "-" (ADR-068) na **Hematologia** (F3-T14, 21 linhas).

**Não validado:**

- o ET do Painel quando o bias é SEM EP (P-10, P-21);
- o resumo "Sem EP" (P-32);
- a coluna AT, ilegível (P-31, P-51);
- o Sigma da Bio, que é numericamente correto mas está contaminado por resultados implausíveis ativos (P-02). Exemplo:
  Glicose N1 0,33 em vez de 4,32.

**Não executado:** ETp ausente na **Bioquímica** (Bio F3-T13). Nenhum analito da Bio está sem ETp no ano de
referência 2025, e provocar o caso exige o Excel.

**Pendente de método** (Especialista em CQ): P-11, P-74, P-33, P-50, P-76, P-52 e P-53.

### D.3 A integração CIQ + CEQ está validada?

**Validada no caminho feliz. Não validada nas bordas do CEQ.**

**Validado:**

- o CEQ entra no bias do analito certo, em todos os níveis, e muda só ele (Bio F2-T34, Hema HF2-T07/T22);
- o isolamento por provedor (Controllab sintético não mexe no CAP) e por ano (N4) (Bio F2-T36, Hema HF2-T12/T13);
- o bias não depende do lote do CIQ nem do período (Bio F2-T37, Hema HF2-T17 na parte do CEQ);
- o CIQ alimenta o CV e o CVp do ET e do Sigma (F3);
- o caminho do dado está mapeado em B.5.

**Não validado:**

- o ANALISTA não consegue corrigir (P-07, P-17);
- a correção só vale depois do Consolidar, com indicador enganoso (P-08, P-18);
- nome fora do mapa, Controllab e analito novo são gravados e ignorados (P-09, P-19);
- o nome do vizinho cai no bias errado (P-20);
- registro sem alvo é aceito (P-25, P-45);
- SalvarCEQ grava sem sessão (P-27, P-47) e não deixa trilha (P-26, P-46);
- linhas Controllab gravadas como CAP (P-48).

**Não executado:** CAP × Controllab com dado **real**, porque o Controllab tem 0 linhas nos dois produtos. O isolamento
foi provado só com dados sintéticos.

**Por projeto, pendente de confirmação:** o mesmo bias do CEQ vale para o DIMENSION 1 e o DIMENSION 2 (ADR-087).

### D.4 Os filtros de período e de lote estão corretos?

**No Painel, sim, com duas exceções. Na aba Estatística, não.**

**Validado (Painel e motor):**

- mês a mês: 13 meses na Bio, 20 combinações mês × lote na Hema;
- trimestres calculados dos dados brutos, e não da média dos meses (Bio T1/2026 87,8983 contra 87,8384; Hema
  T4/2025 2,9920 contra 2,9987);
- semestres, ano inteiro e "Tudo" em lote com até 180 corridas;
- lote A → lote B, com 0 pontos fora do escopo e 0 faltando;
- H3 vazio com lotes em degrau, Westgard lote a lote e lotes paralelos;
- combinações de analito e equipamento;
- LJ (eixo, corrida do lote do ADR-080, 7 limites) e Westgard (ADR-083) por ponto: 2.548 marcas e pontos, 0
  divergências.

**Não validado:**

- "Tudo" em lote com mais de 180 corridas na Bio (P-12) e H3 vazio em período longo na Hema (P-56): a estatística vem
  só da janela de 180 corridas;
- a Estatística da Hema não segue o lote nem o ano (P-04);
- a Estatística da Bio no H3 vazio calcula o lote em uso (P-28);
- as visões MENSAL e TRIMESTRAL da Bio estão quebradas (P-37);
- o filtro Ano De/Até da Hema corta lotes que atravessam o ano (P-54);
- Bias e Sigma no H3 vazio contrariam o ADR-086 dec. 4 (P-35, P-57; dúvida de método).

**Usabilidade dos filtros:** P-30, P-60, P-38, P-61, P-40, P-41 e P-59.

**Não executado / não aplicável:** outro equipamento na Hema, que tem só o XN1000.

### D.5 A entrada manual e a automática estão validadas?

**Automática: transferência validada; plausibilidade não validada.**

- **Validado nos dois produtos:** recebimento, ID, associação, Sync_Log, idempotência e histórico.
- **Hema:** 0 valores ≤ 0 ativos (F1 A06).
- **Bio, não validado:** 56 ativos ≤ 0 na produção e 253 suspeitos ativos (P-02).

**Manual: validado no caso comum; não validado nas entradas coladas ou fora do padrão.**

- **Validado:**
  - campos vazios, texto, ponto, "1.5", data futura, nível inválido, hora 25:99, duplicidade manual × manual e
    manual × interfaceamento, SUBSTITUI_ID com e sem MOTIVO, espaço sobrando (Hema);
  - equivalência manual × automático bit a bit nos dois produtos (E01).
- **Não validado:**
  - minúsculas (P-01, P-03);
  - data inválida colada (P-05, P-16);
  - analito, lote ou equipamento inexistente colado (P-23, P-42);
  - equipamento vazio na Bio (P-06);
  - −1 e 0 digitados na Hema (P-15).

### D.6 Há divergência matemática?

**Nas fórmulas, não.** São 0 divergências em milhares de comparações, nos dois produtos e em três implementações
independentes.

**No que é exibido, sim, em 13 cálculos (B.3).** São população errada ou exibição desatualizada, não conta errada:

- Painel × Estatística com minúsculas (P-01, P-03);
- ET do Painel sem bias (P-10, P-21);
- resumo "Sem EP" com 22 contra 4 (P-32);
- janela de 180 corridas (P-12, P-56);
- Estatística da Hema parada no lote anterior (P-04).

**Diferenças de método** sem reprovação: ver B.4.

### D.7 Há problema operacional?

**Sim.**

- **CEQ:**
  - correção impossível para o ANALISTA (P-07, P-17);
  - Consolidar escondido, com indicador enganoso (P-08, P-18);
  - gravação sem sessão e sem trilha (P-26, P-27, P-46, P-47);
  - GUIA desatualizado (P-29, P-49).
- **Cadastro e dados:**
  - alvo do lote 522611 da Hema incompatível com os dados em 12 pares, gerando centenas de rejeições de Westgard que
    vêm do cadastro (P-22);
  - lote em uso 522611 vencido (P-58);
  - lotes de urina e de outro analito ativos em analitos séricos da Bio (P-13, P-14);
  - biblioteca de especificações 2026–2029 vazia na Hema e incompleta na Bio: um lote novo de 2026 perde o ETp e o
    Sigma sem aviso (P-34, P-55).
- **Alarmes e mensagens:**
  - "ERRO DE GOVERNANÇA" permanente na Bio, por 4 inativações sem motivo (P-24);
  - o mesmo texto na Hema para uma simples linha incompleta (P-43).
- **Tempo:** ATUALIZAR de 15 a 66 s e gravação do CEQ de 15 a 24 s, sem indicador de progresso.
- **Estado da base (dúvidas a confirmar com o laboratório):**
  - a base da Hema na produção vai só até 01/10/2026 e a Sync_Log está vazia;
  - o DB_SEAC não tinha resultado novo desde 30/10 em 01/11, nos dois produtos.

### D.8 O que fica pendente

1. **Corrigir e retestar**, nesta ordem: as 4 Críticas (P-01 a P-04), depois as 18 Altas (P-05 a P-22). Cada correção
   precisa do parecer do Especialista em CQ e da autorização do usuário. O reteste refaz exatamente o fluxo original
   numa cópia nova e também o fluxo vizinho.
2. **Não executados:**
   - janela `frmCEQ` vista na tela (Bio F2-T22, Hema HF2-T25): é modal e a captura de tela é proibida; pode ser
     coberta pela suíte `qa_adr087` V01 ou por exportação do próprio Excel;
   - ETp ausente "-" na Bioquímica (Bio F3-T13);
   - CAP × Controllab com dado real.
   - Os testes Bio F3-T14 e Hema F3-T15 (mudança ao incluir ou corrigir CEQ/CIQ) **foram cobertos** pelas F1/F2:
     Bio F2-T34/T35 e F1 M02; Hema HF2-T06/T10 e F1 M02.
3. **49 dúvidas de regra de negócio para o Especialista em CQ, via Orquestrador.** As que bloqueiam correções:
   - triagem de plausibilidade: bloqueia ou só alerta, e com qual limite (P-02, P-15);
   - BiasEQ deve excluir resultado 0 e NAO AVALIADO? (P-11, P-74);
   - qual CV entra no ET e qual Sigma é o oficial (P-33, P-50);
   - janela de 180 corridas ou período inteiro (P-12, P-56);
   - Bias e Sigma no H3 vazio (P-35, P-57);
   - EQUIPAMENTO obrigatório na Bio (P-06);
   - provedor global R4 da Hema;
   - Sigma negativo: mostrar o número ou "ETp excedido";
   - papel mínimo para registrar e corrigir CEQ;
   - ETp de WBC e PLT e do RDW (P-52, P-53).
4. **Ações do laboratório** (dados, não código):
   - revisar e inativar os 56 + 253 IDs suspeitos da Bio;
   - conferir o alvo do lote 522611 da Hema (P-22) e atualizar o lote em uso (P-58);
   - preencher a biblioteca 2026–2029 (P-34, P-55);
   - confirmar a falta de controle no DB_SEAC desde 30/10 e a base da Hema até 01/10;
   - dar motivo às 4 inativações da Bio (P-24).
