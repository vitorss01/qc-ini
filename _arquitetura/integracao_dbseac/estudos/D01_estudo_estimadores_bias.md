# D-01: estudo comparativo dos estimadores de viés do CEQ para Sigma e Erro Total

**Data:** 06/10/2026 · **Autor:** gestor técnico e da qualidade (agente) · **Tipo:** estudo somente leitura, sem mudança em produção
**Status da pendência D-01:** continua aberta, **aguardando decisão formal do RT** (a seção 12 traz a minuta)
**Script:** `estudos/d01_estudo.py` · **Resultados completos:** `estudos/d01_resultados_bioquimica.csv`, `estudos/d01_resultados_hematologia.csv`

---

## 1. Resumo executivo

1. **O estudo reproduz exatamente o que está em produção.** O estimador A recalculado bate com a coluna "Bias EQC (abs) %" em 100% das linhas (Bio 58/58, Hema 72/72). Sigma, classe e regras de Westgard também batem 100% (Bio 58/58, Hema 57/57). O impacto registrado no ADR-064 se confirma: trocando A por B, **Bio 5/58 e Hema 15/57 linhas sobem de classe**, e MCH N3 vai de **1,74σ para 3,25σ**.
2. **Os dados não sustentam B (|média com sinal|) puro como viés do Sigma.** As **20 de 20** subidas de classe que B produziria têm duas coisas em comum: **o IC95% do viés inclui zero** e **o sinal da média de rodada se inverte** entre as 3 rodadas de 2025. B "melhora" essas linhas porque cancela rodadas positivas contra negativas, não porque mostrou um viés pequeno.
3. **MCH é o caso que mais pesa.** As médias por rodada foram +1,58%, −0,52% e −1,64%. A ANOVA entre rodadas dá **p = 5×10⁻⁵**. A dispersão dentro de cada rodada (0,72%) é do tamanho do CV do CIQ (razão 0,96). O laboratório estava deslocado em cada rodada, e o deslocamento mudou de direção de uma rodada para outra. **B apaga um erro sistemático real que variou no tempo.**
4. **A crítica ao estimador A ("conta a imprecisão duas vezes") vale só em parte.** Ela se confirma onde a dispersão de amostra única domina, ou seja, no diferencial leucocitário (EO, MONO, LYMPH): ali A chega a ser ≥ 2×B. Não se confirma onde há deslocamento entre rodadas. Na mediana, A fica muito perto do estimador por rodada G (G/A = 1,00 na Bio e 0,90 na Hema).
5. **Ajuste do gestor: acrescentei o estimador G** (RMS das médias com sinal de cada rodada). G calcula a média com sinal **dentro** de cada rodada, o que dilui o erro de amostra única por √m. Entre rodadas ele usa o RMS, que **não cancela** deslocamentos. Contra A, G muda **Bio 3↑/1↓** e **Hema 8↑/2↓**. MCH N3 fica em **1,71σ (inadequado)**.
6. **Recomendação provisória ao RT** (seções 11 e 12): não trocar por B. Adotar G como "viés sistemático" do Sigma/ET tradicional, com salvaguardas de transição. Rebaixar A a indicador ("magnitude média do desvio no CEQ"). Manter C (RMS por amostra) só na incerteza (ADR-067). Deixar D (regressão) para uma fase 2. **Até a decisão do RT, a produção continua com A**, que é o estado atual, sem mudança.

---

## 2. Avaliação crítica das orientações recebidas

Conferi as referências no PubMed. Os textos completos de Ercan 2022 e Coskun 2022 foram lidos no PMC.

| Orientação | Avaliação do gestor |
|---|---|
| Guardar o bias **com sinal**; no Sigma tradicional bilateral entra \|viés sistemático estimado\|: σ = (TEa − \|Bias_sist\|)/CV (Westgard 2018) | **Concordo.** A fórmula da planilha já é essa (`L = (K − ABS(G))/F`). O que está em discussão é **qual número vai em G**, não a fórmula. |
| Mean(\|bias\|) não é viés sistemático; deve ficar como indicador separado | **Concordo no conceito.** É a média do erro absoluto por amostra e mistura erro sistemático com erro aleatório. **Ressalva empírica:** nestes dados, A fica próximo de um estimador que dilui o aleatório (G). Rebaixá-lo a indicador está certo, mas o motivo mais forte para não usá-lo é conceitual: ele não é estimativa de viés. O "excesso" numérico é pequeno na maioria dos analitos (seção 6). |
| Coskun 2022 criticou o RMS de Ercan 2022 e não recomendou a média com sinal | **Confirmado pelo texto completo.** Ercan 2022 critica a média aritmética com sinal (o cancelamento) e propõe o RMS. Coskun 2022 responde que o RMS é uma versão simplificada da fórmula de variâncias agrupadas, inadequada para o caso. Também lembra que viés é a média de réplicas menos o valor de referência (VIM), que é preciso avaliar a **significância** do viés e que o uso linear do viés no Sigma é questionável. **A carta não traz o argumento E\|X\| = σ√(2/π).** O ADR-064 atribuiu esse argumento a Coskun 2022 por engano: é uma propriedade matemática da distribuição normal dobrada e não precisa de citação. A seção 13 traz o texto de correção. |
| Coskun 2024: significância; viés constante/proporcional por regressão | **Concordo e incorporei.** IC95% por amostra **e por rodada**, regressão Xlab × Xref (estimador D) e o estimador F (viés "desprezado se não significativo") como sensibilidade. |
| Wauthier 2022: o estimador muda muito o Sigma | **Concordo, com uma nuance.** No resumo, Wauthier conclui que a fonte do TEa pesou mais e que "a natureza do viés foi menos decisiva". Aqui a mudança de estimador mexe em 9% (Bio) a 26% (Hema) das linhas com B. Na Bio, 43 de 58 linhas ficam "inadequadas" com **qualquer** estimador, porque o CV domina (seção 10). |
| Não trocar a fórmula ainda nem manter a antiga definitivamente: fazer primeiro o estudo comparativo (A–E) | **Concordo, com ajustes** (seção 3.4): (i) acrescentei G e a ANOVA entre rodadas, porque com 3 rodadas o fenômeno que decide é a **heterogeneidade no tempo**; (ii) IC95% também pela rodada (unidade de agrupamento), já que o IC por amostra trata 15 amostras como independentes quando são 3 grupos de 5; (iii) duas seleções de amostras (S1 = a do Sigma de hoje e S2 = a do ADR-064), porque elas **não são iguais**; (iv) D só por interpolação; (v) F apenas informativo. |
| Separar "qual viés" da controvérsia "o modelo linear de Sigma é adequado?" (Coskun 2019) | **Concordo.** Fica fora do escopo desta decisão (seção 14). |

---

## 3. Método

### 3.1 Dados

- Cópias byte a byte dos `.xlsm` de produção da entrega de 05/10/2026. O SHA-256 das cópias é igual ao dos originais. A leitura foi feita com openpyxl (`data_only`, `read_only`), sem Excel e sem gravar nada nas pastas.
- CEQ: aba `EQA_Base`, provedor **CAP**, **Ano EP 2025**, rodada TODAS (nomes `eqAnoEP`, `eqProvedor`, `eqRodada` da própria pasta). Na Bioquímica o provedor vem por analito (`Analitos!AR`) e é CAP em todos.
- Estrutura real: **15 amostras em 3 rodadas (5 por rodada)** por analito. Bio: 29 analitos com CEQ. Hema: 24 analitos com CEQ, dos quais 19 têm Sigma.
- TEa = coluna "ETp %"; CV = coluna "CV %" da aba `Estatística` (os mesmos valores do Sigma de hoje).
- Classe, regras, N, run size e frequência saem da `tblPlanoQC_Sigma` (aba `Cfg_PlanoQC`) de cada pasta, pela regra `Sigma_Min ≤ s < Sigma_Max` do `mPlanoQC`.

### 3.2 Seleção de amostras

| Seleção | Regra | Uso |
|---|---|---|
| **S1** | Exatamente a do `mCEQ.BiasEQ`: analito canônico, `Uso_Analitico ≠ "NAO"`, provedor, rodada, ano vigente (maior ano ≤ Ano EP) e `Bias_Abs` "numérico" no sentido do VBA (célula vazia conta como 0; não houve nenhum caso) | Principal: é a que reproduz o Sigma de hoje |
| **S2** | A do `mCEQ.ViesEQ` (ADR-064): S1 + amostra avaliada pelo provedor (≠ NAO AVALIADO) + resultado e alvo numéricos e ≠ 0 + chave única | Sensibilidade: é a regra auditada do ADR-064 |

> **Achado:** o `BiasEQ`, que alimenta o Sigma, **não aplica** os filtros que o ADR-064 pôs no `ViesEQ`. Ele aceita amostras NAO AVALIADO, resultado ou alvo zero e chave duplicada. Hoje isso **não muda nenhuma classe** (S1→S2: 0 linhas com A, B, C ou G). Muda números: Bilirrubina direta passa de A = 30,77% (15 amostras) para A = 19,99% (13 amostras). Na Hema, IG e NRBC caem para 0 amostra em S2, mas não têm Sigma. Na implementação da decisão, a seleção deve ser unificada em S2.

### 3.3 Estimadores (todos por analito; o Sigma é por analito × nível)

| | Definição | Observação |
|---|---|---|
| **A** | Média, por rodada, de \|bias\|; depois a média das rodadas (2 etapas, ADR-035) | É o de hoje (`BiasEQ "ABS"`) |
| **B** | \|média por rodada do bias com sinal; depois média das rodadas\| (`BiasEQ "SIGNED"` em módulo). Também a média por amostra, DP, EP e IC95% **por amostra** (t, n−1) e **por rodada** (t, k−1) | A versão por amostra bate com o `ViesEQ "MEDIA"` (Bio 58/58, Hema 60/60) |
| **C** | RMS = √(média dos bias²) por amostra | Nordtest; já entra no u(bias) (ADR-067) |
| **D** | OLS Xlab = a + b·Xref por analito. Viés no nível = (a + (b−1)·Xc)/Xc·100, com Xc = média do controle e IC95% do valor ajustado | Só com ≥ 8 amostras, ≥ 3 concentrações, max/min ≥ 1,5 e Xc dentro de [min/1,25; max×1,25] (sem extrapolação, que também protege contra unidade diferente entre CEQ e CIQ) |
| **E** | Bootstrap estratificado por rodada (2000 réplicas, semente 20261006) de A, B e G: IC 2,5–97,5% do Sigma e fração das réplicas na mesma classe ("estabilidade") | Reamostra **dentro** da rodada. Não captura a incerteza **entre** rodadas (seção 15) |
| **F** | B se o IC95% por amostra exclui 0; senão 0 | Informativo: leitura literal de "viés não significativo pode ser desprezado" |
| **G** | √(média das médias de rodada²), com as médias de rodada com sinal | Acrescentado pelo gestor. Não cancela deslocamentos entre rodadas e dilui o erro de amostra única por √m |

Diagnósticos por analito: ANOVA de 1 via do bias entre rodadas (heterogeneidade), inversão de sinal entre médias de rodada, outliers (|x − mediana| > 3,5·MAD·1,4826) e DP dentro da rodada (DPw) comparado ao CV do CIQ.

### 3.4 Ajustes ao plano original e por quê

1. **Incluí G e a ANOVA.** Com 3 rodadas, a diferença entre A e B vem quase toda da variação **entre** rodadas. Sem medir isso, a comparação A × B não separa "aleatório virando viés" de "viés que muda no tempo".
2. **IC95% pela rodada.** As 5 amostras de uma rodada compartilham calibração e dia, então o IC por amostra é otimista. Com k = 3 (t = 4,30), o IC por rodada mostra o tamanho real da incerteza.
3. **Duas seleções.** Os filtros do Sigma e da incerteza são diferentes hoje (3.2).
4. **D só interpolado.** A unidade do CEQ nem sempre é a do cadastro (ex.: Lactato).

### 3.5 Validação (reprodução da produção)

| | Bioquímica | Hematologia |
|---|---|---|
| A = coluna "Bias EQC (abs) %" | 58/58 | 72/72 |
| B com sinal (2 etapas) = "Bias EQC (sinal) %" | 58/58 | 72/72 |
| n amostras / rodadas = colunas AC/AD | 58/58 | 72/72 |
| Sigma_A = "SIX SIGMA" | 58/58 | 57/57 |
| Classe_A = "Status sigma" | 58/58 | 57/57 |
| Regras_A = "Regras Westgard recomendadas" | 58/58 | 57/57 |
| B por amostra em S2 = "Viés CEQ médio % (com sinal)" | 58/58 | 60/60 |

---

## 4. Resultados agregados

### 4.1 Mudanças de classe contra A (seleção S1; S2 dá resultado idêntico)

| Estimador | Bio: comparáveis / sobe / desce | Hema: comparáveis / sobe / desce |
|---|---|---|
| B = \|Mean(bias)\| | 58 / **5** / 0 | 57 / **15** / 0 |
| C = RMS | 58 / 0 / **8** | 57 / 0 / **6** |
| D = regressão no nível | 52 / **7** / 0 | 30 / **6** / 1 |
| F = B só se significativo | 58 / **6** / 0 | 57 / **16** / 0 |
| **G = RMS das médias de rodada** | 58 / **3** / **1** | 57 / **8** / **2** |

### 4.2 Distribuição das classes (linhas com Sigma)

| | Inadequado | Marginal | Bom | Excelente | Classe mundial |
|---|---|---|---|---|---|
| Bio A | 43 | 3 | 4 | 2 | 6 |
| Bio B | 42 | 3 | 4 | 1 | 8 |
| Bio C | 46 | 3 | 2 | 3 | 4 |
| Bio D (52 linhas) | 35 | 4 | 3 | 0 | 10 |
| Bio G | 43 | 3 | 3 | 2 | 7 |
| Hema A | 27 | 4 | 6 | 8 | 12 |
| Hema B | 24 | 7 | 0 | 8 | 18 |
| Hema C | 29 | 2 | 9 | 6 | 11 |
| Hema D (30 linhas) | 11 | 5 | 2 | 3 | 9 |
| Hema G | 26 | 5 | 5 | 5 | 16 |

### 4.3 Matrizes de mudança (só as células fora da diagonal)

**Bioquímica**

- A→B: Inadequado→Marginal 1; Marginal→Bom 1; Bom→Excelente 1; Excelente→Mundial 2.
- A→C: Marginal→Inadequado 2; Bom→Inadequado 1; Bom→Marginal 2; Excelente→Bom 1; Mundial→Excelente 2.
- A→D: Inadequado→Marginal 2; Marginal→Bom 1; Bom→Mundial 2; Excelente→Mundial 2.
- A→G: Marginal→Inadequado 1 (Bilirrubina total N2); Inadequado→Marginal 1 (LDH N2); Bom→Excelente 1 (ALT N2); Excelente→Mundial 1 (Ferro N2).

**Hematologia**

- A→B: Inadequado→Marginal 3; Bom→Excelente 6; Excelente→Mundial 6.
- A→C: Marginal→Inadequado 2; Excelente→Bom 3; Mundial→Excelente 1.
- A→D: Inadequado→Marginal 1; Bom→Excelente 2; Excelente→Bom 1; Excelente→Mundial 3.
- A→G: Inadequado→Marginal 2 (MONO% N3, MONO# N3); Bom→Excelente 2 (EO% N2, EO# N2); Excelente→Mundial 4 (LYMPH% N3, EO% N1, EO% N3, EO# N1); Marginal→Inadequado 1 (PLT N2); Excelente→Bom 1 (PLT N3).

### 4.4 Diagnóstico por analito (saída do script)

| | Bio (29 analitos com CEQ) | Hema (24 analitos com CEQ) |
|---|---|---|
| Rodadas heterogêneas (ANOVA p < 0,05) | 16 | 7 |
| Inversão de sinal entre médias de rodada | 17 | 12 |
| IC95% **por amostra** inclui 0 | 13 | 13 |
| IC95% **por rodada** inclui 0 | 25 | 19 |
| Viés proporcional (IC da inclinação exclui 1) / constante | 10 / 5 | 7 / 1 |
| Com outlier | 8 | 2 (NRBC%, NRBC#, sem Sigma) |
| A ≥ 2·B | 10 | 9 |
| Mediana A/B · C/A · G/A | 1,46 · 1,24 · 1,00 | 1,62 · 1,21 · 0,90 |
| Mediana DPw/CV do CIQ (linhas; coluna `S1_razao_DPw_CV`) | 0,44 (14/58 > 1) | 0,93 (25/57 > 1) |

**O que explica as subidas de classe com B:**

| | Bio | Hema |
|---|---|---|
| Linhas que sobem com B | 5 | 15 |
| ... com IC95% por amostra incluindo 0 | 5 | 15 |
| ... com IC95% por rodada incluindo 0 | 5 | 15 |
| ... com inversão de sinal entre rodadas | 5 | 15 |
| ... com rodadas heterogêneas (p < 0,05) | 2 (Glicose N2, CK N2) | 1 (MCH N3) |

### 4.5 Estabilidade (bootstrap) e disponibilidade de D

- Linhas com menos de 80% das réplicas na mesma classe: Bio A 7/58, B 6/58, G 8/58; Hema A 10/57, B 10/57, G 11/57. A instabilidade é parecida nos três estimadores e se concentra nas linhas perto de um limite de classe.
- D calculável no nível: Bio 52/62 linhas. Sem D: 4 por falta de amostras com Xlab/Xref (analitos sem CEQ), 4 por faixa de concentração estreita (Sódio, Cloro) e 2 por extrapolação (Lactato N2, Ureia N2). Hema: 39/84 linhas (30 com Sigma). Sem D: 33 por faixa estreita (MCV, MCH, MCHC, RDW-CV, MPV, NEUT%, LYMPH%, MONO%, EO%, BASO%, IG%) e 12 sem CEQ (RDW-SD, RET%, RET#, IRF).

---

## 5. Casos críticos (seleção S1)

Legenda: σ = Sigma com o estimador. "σB pior" = Sigma de B no limite desfavorável do IC95% **por rodada**. Estab. = fração das réplicas bootstrap na mesma classe (A/B/G). Classe: Inad. = inadequado; Marg. = marginal; Exc. = excelente; Mund. = classe mundial.

### 5.1 Hematologia

| Analito N | CV | TEa | A | B com sinal [IC95 amostra] [IC95 rodada] | C | G | D | σA | σB | σC | σD | σG | σB pior | Estab. A/B/G | Sinais |
|---|---|---|---|---|---|---|---|---|---|---|---|---|---|---|---|
| **MCH N3** | 0,75 | 2,63 | 1,33 | −0,20 [−1,05; 0,65] [−4,25; 3,86] | 1,49 | 1,35 | — | **1,74 Inad.** | **3,25 Marg.** | 1,52 Inad. | — | **1,71 Inad.** | −2,17 | 1,00/0,88/1,00 | rodadas heterogêneas (p = 5×10⁻⁵); inversão de sinal |
| MCH N2 | 0,83 | 2,63 | 1,33 | −0,20 (idem) | 1,49 | 1,35 | — | 1,57 Inad. | 2,95 Inad. | 1,38 Inad. | — | 1,55 Inad. | −1,97 | 1,00/0,63/1,00 | idem |
| MCH N1 | 1,18 | 2,63 | 1,33 | −0,20 (idem) | 1,49 | 1,35 | — | 1,10 Inad. | 2,07 Inad. | 0,96 Inad. | — | 1,09 Inad. | −1,38 | 1,00/1,00/1,00 | idem |
| PLT N2 | 3,15 | 15,00 | 4,96 | +4,52 [2,01; 7,04] [−5,55; 14,59] | 6,30 | 5,60 | +4,82 | 3,18 Marg. | 3,33 Marg. | 2,76 Inad. | 3,23 Marg. | 2,98 Inad. | 0,13 | 0,79/0,91/0,56 | rodadas heterogêneas (p = 0,006); inversão; viés proporcional |
| PLT N3 | 1,98 | 15,00 | 4,96 | +4,52 (idem) | 6,30 | 5,60 | +5,71 | 5,07 Exc. | 5,29 Exc. | 4,39 Bom | 4,69 Bom | 4,75 Bom | 0,21 | 0,58/0,75/0,71 | idem |
| WBC N3 | 0,91 | 7,00 | 1,86 | −1,86 [−2,81; −0,92] [−2,63; −1,10] | 2,49 | 1,88 | −0,90 | 5,67 Exc. | 5,67 Exc. | 4,97 Bom | 6,73 Mund. | 5,65 Exc. | 4,82 | 0,66/0,66/0,69 | viés significativo e estável |
| EO% N1 | 6,27 | 43,56 | 7,69 | +0,18 [−4,58; 4,94] [−5,75; 6,11] | 8,31 | 1,96 | — | 5,72 Exc. | 6,92 Mund. | 5,62 Exc. | — | 6,64 Mund. | 5,98 | 0,99/1,00/0,92 | inversão; A ≥ 2B; DPw/CV = 1,4 |
| EO% N3 | 6,93 | 43,56 | 7,69 | +0,18 (idem) | 8,31 | 1,96 | — | 5,17 Exc. | 6,26 Mund. | 5,09 Exc. | — | 6,00 Mund. | 5,40 | 0,94/0,66/0,12 | idem; G encostado no limite 6 |
| LYMPH# N3 | 3,42 | 22,76 | 3,27 | +0,33 [−1,99; 2,66] [−7,09; 7,76] | 4,07 | 2,46 | +1,29 | 5,70 Exc. | 6,56 Mund. | 5,47 Exc. | 6,28 Mund. | 5,94 Exc. | 4,39 | 0,96/0,99/0,78 | inversão |
| LYMPH% N1 | 4,08 | 22,76 | 3,63 | +2,21 [−0,19; 4,61] [−5,47; 9,88] | 4,73 | 3,35 | — | 4,69 Bom | 5,04 Exc. | 4,42 Bom | — | 4,76 Bom | 3,16 | 0,96/0,59/0,90 | inversão |
| MONO# N3 | 6,98 | 26,16 | 6,12 | +0,20 [−4,47; 4,87] [−5,42; 5,82] | 8,16 | 1,86 | +1,36 | 2,87 Inad. | 3,72 Marg. | 2,58 Inad. | 3,55 Marg. | 3,48 Marg. | 2,91 | 0,75/0,99/0,83 | inversão; A ≥ 2B |
| EO# N2 | 7,72 | 43,56 | 8,18 | −1,65 [−6,73; 3,43] [−6,92; 3,61] | 9,01 | 2,39 | −0,76 | 4,58 Bom | 5,43 Exc. | 4,47 Bom | 5,54 Exc. | 5,33 Exc. | 4,75 | 1,00/0,93/0,67 | inversão; A ≥ 2B |

As demais linhas que sobem com B (LYMPH% N2/N3, MONO% N3, EO% N2, LYMPH# N1/N2, EO# N1/N3) estão no CSV, com o mesmo padrão: inversão de sinal e IC incluindo zero. WBC N2 (Marginal → Inadequado) e BASO# N1 (Mundial → Excelente) mudam **só com C**.

### 5.2 Bioquímica

| Analito N | CV | TEa | A | B com sinal [IC95 amostra] [IC95 rodada] | C | G | D | σA | σB | σC | σD | σG | σB pior | Estab. A/B/G | Sinais |
|---|---|---|---|---|---|---|---|---|---|---|---|---|---|---|---|
| Glicose N2 | 1,37 | 10,00 | 2,73 | −1,33 [−3,42; 0,77] [−9,83; 7,18] | 3,89 | 3,09 | +0,19 | 5,29 Exc. | 6,32 Mund. | 4,45 Bom | 7,15 Mund. | 5,03 Exc. | 0,12 | 0,74/0,76/0,44 | rodadas heterogêneas (p = 0,005); inversão |
| CK N2 | 4,02 | 20,00 | 6,69 | −3,61 [−8,55; 1,33] [−22,83; 15,61] | 9,35 | 7,27 | +0,09 | 3,31 Marg. | 4,07 Bom | 2,65 Inad. | 4,95 Bom | 3,16 Marg. | −0,70 | 0,80/0,61/0,68 | rodadas heterogêneas (p = 0,01); inversão |
| ALT N2 | 2,90 | 20,00 | 6,19 | −2,72 [−8,85; 3,41] [−16,92; 11,49] | 11,03 | 5,40 | +0,23 | 4,76 Bom | 5,95 Exc. | 3,09 Marg. | 6,81 Mund. | 5,03 Exc. | 1,06 | 0,41/0,34/0,32 | inversão; viés proporcional; 1 outlier |
| Bilirrubina total N2 | 3,03 | 20,00 | 10,82 | −8,08 [−14,83; −1,32] [−38,52; 22,36] | 14,29 | 12,86 | −7,99 | 3,03 Marg. | 3,94 Marg. | 1,89 Inad. | 3,97 Marg. | 2,36 Inad. | −6,12 | 0,53/0,46/0,81 | rodadas heterogêneas (p = 0,0005); inversão; viés proporcional |
| LDH N2 | 4,53 | 15,00 | 1,60 | −0,38 [−1,36; 0,61] [−3,42; 2,66] | 1,76 | 1,07 | +0,30 | 2,96 Inad. | 3,23 Marg. | 2,93 Inad. | 3,25 Marg. | 3,08 Marg. | 2,56 | 0,84/1,00/0,82 | inversão |
| Ferro N2 | 2,33 | 15,00 | 1,15 | −0,05 [−0,96; 0,85] [−1,68; 1,57] | 1,58 | 0,54 | −0,61 | 5,93 Exc. | 6,40 Mund. | 5,75 Exc. | 6,17 Mund. | 6,20 Mund. | 5,71 | 0,72/0,99/0,75 | inversão; 1 outlier |
| AST N2 | 2,06 | 20,00 | 9,94 | +9,94 [4,05; 15,84] [5,64; 14,24] | 14,30 | 10,04 | +4,81 | 4,88 Bom | 4,88 Bom | 2,76 Inad. | 7,37 Mund. | 4,83 Bom | 2,79 | 0,27/0,27/0,24 | viés significativo e estável; proporcional (b = 1,037 [1,014; 1,060]) |
| Bilirrubina direta N2 | 6,65 | 39,92 | 30,77 | −30,77 [−47,90; −13,64] [−61,52; −0,02] | 42,90 | 32,39 | −17,08 | 1,38 Inad. | 1,38 Inad. | −0,45 Inad. | 3,44 Marg. | 1,13 Inad. | −3,25 | 0,94/0,94/0,98 | proporcional (b = 0,848); 2 outliers; S2 → A = 19,99 |

---

## 6. Casos especiais

| Situação | Exemplo real | Leitura |
|---|---|---|
| **Mean\|Bias\| alto com Mean(Bias) ≈ 0, rodadas homogêneas** | EO%: A = 7,69, B = 0,18, ANOVA p não significativa, DPw = 9,0 contra CV 6–7 | O ruído de amostra única domina, e aqui A **conta o aleatório duas vezes** (a crítica do ADR-064 vale). B e G concordam que o viés sistemático é pequeno. G = 1,96 ainda carrega o ruído residual σw/√5. |
| **Mean\|Bias\| alto com Mean(Bias) ≈ 0, rodadas heterogêneas** | MCH: rodadas +1,58 / −0,52 / −1,64, p = 5×10⁻⁵, DPw/CV ≈ 1 | **B cancela erro sistemático real que variou no tempo** (deriva de calibração/lote ou efeito do material da rodada). A e G concordam (1,33 e 1,35). É o caso que desqualifica B puro. |
| **Mean(Bias) alto e estável** | AST N2 (+9,94; IC por rodada [5,64; 14,24]); WBC (−1,86; IC por rodada [−2,63; −1,10]) | Os estimadores convergem: A = B ≈ G. Viés significativo e relevante deve ser **investigado/corrigido** (Coskun 2024), não só "absorvido" no Sigma. |
| **Tendência com a concentração** | AST (b = 1,037), Bilirrubina direta (b = 0,848), PLT (b = 1,063) | O viés médio do CEQ não é o viés no nível do controle. D muda a classe de AST N2 (4,88 → 7,37) e de PLT N3 (5,07 → 4,69). É argumento para a fase 2 com D. |
| **Inversão de sinal no tempo** | 17/29 analitos na Bio e 12/24 na Hema | Com 3 rodadas, a média com sinal é frágil: um sinal trocado basta para zerar o viés. |
| **Outliers** | Bilirrubina direta (2), AST (2), NRBC (6) | NRBC e Bilirrubina direta perto do limite de detecção: o bias relativo explode (±100%). Já são cortados em S2 (alvo/resultado 0, NAO AVALIADO). |
| **IC95% inclui ou não zero** | Por amostra: inclui em 13/29 (Bio) e 13/24 (Hema). Por rodada: 25/29 e 19/24 | Com k = 3 rodadas, o teste por rodada quase nunca rejeita zero. "Não significativo" aqui é **falta de poder**, não ausência de viés. Por isso F (zerar) é rejeitado. |

---

## 7. O que cada estimador representa e seus riscos

| | Representa | Riscos | Adequação ao CEQ |
|---|---|---|---|
| **A** Mean(\|bias\|) | Erro absoluto médio por amostra (sistemático + aleatório) | Com viés zero dá ≈ 0,8·σ da dispersão das amostras; conta a imprecisão de novo (já está no CV) | Conservador. Não é viés no sentido do VIM. Útil como **indicador de desempenho no CEQ** |
| **B** \|Mean(bias)\| | Viés constante médio do período, se o viés for estável | **Cancelamento de sinal** entre rodadas; com 3 rodadas, IC enorme; esconde deriva no tempo | Correto em teoria para viés estável. Frágil com poucas rodadas e viés que muda |
| **C** RMS por amostra | √(viés² + dispersão²): viés mais toda a dispersão das amostras | Conta o aleatório **inteiro** (mais que A); Coskun 2022 critica como uso indevido de variâncias agrupadas | Bom para **incerteza** (Nordtest, ADR-067), onde a dispersão deve entrar. Inadequado no Sigma, onde o CV já está no denominador |
| **D** Regressão | Viés constante + proporcional, avaliado na concentração do controle | Extrapolação; unidade CEQ × cadastro; faixa estreita (índices hematimétricos); pontos influentes | O mais informativo quando há faixa (CAP cobre várias concentrações). Exige harmonizar unidades |
| **E** Bootstrap | Estabilidade da estimativa e da classe | Com 5 amostras por rodada e reamostragem dentro da rodada, **subestima** a incerteza entre rodadas | Ferramenta de decisão de transição ("só muda se for estável"), não estimador |
| **F** B se significativo | "Viés não significativo é desprezado" | Com baixo poder, zera viés real. Foi o que mais promoveu classes (Bio 6, Hema 16) | **Rejeitado** para Sigma |
| **G** RMS das médias de rodada | Viés sistemático **incluindo a variação entre rodadas**, com o ruído de amostra única reduzido por √m | Ainda carrega σw/√m (conservador); com 3 rodadas é ruidoso; não é estimador consagrado na literatura | Equilíbrio: concorda com B quando o viés é estável e com A quando há deriva entre rodadas |

---

## 8. Analitos mais sensíveis à escolha

- **Hematologia:** MCH (N1–N3), diferencial leucocitário (EO%, EO#, LYMPH%, LYMPH#, MONO% N3, MONO# N3), PLT N2/N3, WBC N2/N3 (com C) e BASO# N1.
- **Bioquímica:** Glicose N2, CK N2, ALT N2, Bilirrubina total N2, LDH N2, Ferro N2, AST N2 (com D), Creatinina N2, Triglicerídeos N2 e HDL N1.

## 9. Impacto na segurança e no controle estatístico

- **Com B:** 20 linhas teriam as regras **relaxadas**. Exemplos: EO% N1/N3, LYMPH% N3, LYMPH# N3 e EO# N1/N3 passariam de `1_3s/2of3_2s/R_4s` para **só `1_3s`**. CK N2 iria de N = 6 e run size 45 para N = 4 e run size 200. ALT N2 iria de N = 4 e run size 200 para N = 2 e run size 450. Glicose N2 e Ferro N2 iriam para `1_3s` com run size 1000. **Todas** essas linhas têm IC95% do viés incluindo zero e inversão de sinal: é flexibilização apoiada em cancelamento.
- **Com G:** 11 linhas relaxam (EO, MONO, LYMPH% N3, Ferro N2, LDH N2, ALT N2) e **3 apertam** (PLT N2 → inadequado, PLT N3 → Bom, Bilirrubina total N2 → inadequado). As que apertam são justamente as que têm deriva entre rodadas.
- **MCH:** com B, N3 sairia de "CQ estatístico isolado pode ser insuficiente — investigar/melhorar" para plano marginal (N = 6). Com A ou G, continua sinalizado para investigação, que é o que os dados pedem: há deslocamento entre rodadas.

---

## 10. Achados colaterais (fora do D-01, registrados para o RT/gestor)

1. **Lipase sem ETp** (TEa = 0 nas duas linhas): o Sigma sai negativo (−1,14 e −2,31) e a classe "inadequado" é artefato de cadastro, não desempenho. Cadastrar o ETp.
2. **Bioquímica: CV do CIQ muito alto no nível 1** em vários analitos (ex.: HDL N2 68,4%, Amilase N1 62,0%, AST N1 33,0%, CK N1 31,0%, Bilirrubina direta N1 26,3%, Glicose N1 20,4%, Cálcio N1 15,1%). Isso deixa 43 de 58 linhas "inadequadas" com **qualquer** estimador de viés. O CV domina o Sigma da Bio e precisa de investigação própria (mistura de lote ou de nível no período? dados do DB_SEAC?) antes de qualquer conclusão sobre o viés.
3. **Período diferente entre CV e viés:** o CV é do período de 2026 e o viés vem do CEQ de 2025 (Ano EP vigente). É uma limitação comum a todos os estimadores.
4. **Seleção do BiasEQ ≠ ViesEQ** (seção 3.2): unificar em S2 na implementação.

---

## 11. Recomendação provisória ao RT

**Qual estimador.** Para o viés sistemático do Sigma/ET tradicional, **G**: a média com sinal de cada rodada, agregada pelo RMS entre rodadas. Manter o sinal guardado e exibido (direção), com B e o IC95% como indicador de direção e significância. Rebaixar A a indicador, com o nome **"magnitude média do desvio no CEQ"**. C continua só na incerteza (ADR-067). D vira verificação de fase 2 nos analitos com viés proporcional.

**Por quê.**
1. B puro produziria 20 promoções, todas sustentadas por cancelamento entre rodadas com IC incluindo zero. No MCH, apagaria um deslocamento estatisticamente comprovado.
2. A não é estimativa de viés e superestima onde o ruído de amostra única domina (diferencial leucocitário).
3. G atende às três críticas da literatura de uma vez: usa a média de amostras por rodada (Coskun 2022: viés é média, não resultado isolado), não cancela rodadas (Ercan 2022) e não soma a dispersão inteira das amostras (crítica ao RMS por amostra e ao E|X|).

**Limitações.**
- G é proposta do gestor, não fórmula consagrada.
- Só há 3 rodadas (IC por rodada com t = 4,30), e o bootstrap subestima a incerteza entre rodadas.
- O CV e o viés vêm de períodos diferentes.
- A regressão fica indisponível para índices e percentuais (faixa estreita).
- Na Bio o CV domina, então a escolha do viés muda pouco a gestão.

**Impacto.** Bio 3↑/1↓ e Hema 8↑/2↓ contra a produção, com menos relaxamento que B (5 e 15) e com aperto onde há deriva.

## 12. Decisão a submeter formalmente ao RT (minuta)

> **D-01. Viés do CEQ usado no Sigma e no Erro Total** (ISO 15189, controle de mudanças)
>
> 1. ( ) Aprovo que o viés sistemático do Sigma/ET seja **G = √(média das médias de rodada²)**, com as médias de rodada com sinal, sobre a seleção de amostras do ADR-064 (avaliadas, alvo/resultado ≠ 0, chave única, Uso_Analitico = SIM).
> 2. ( ) Aprovo que **Mean(|bias|)** passe a ser exibido como indicador "magnitude média do desvio no CEQ", fora da fórmula.
> 3. ( ) **Transição:** nenhuma linha tem o plano de CQ relaxado automaticamente. As 11 linhas que sobem com G (lista no CSV, colunas `S1_classe_A` × `S1_classe_G`) só mudam de regra após revisão linha a linha pelo RT. A mudança não se aplica se a estabilidade bootstrap for < 80%, nem se o Status SDI ou o Status limites do CEQ estiverem "FORA/NÃO OK" (aí vale o plano mais conservador entre A e G). As 3 linhas que apertam (PLT N2/N3, Bilirrubina total N2) mudam imediatamente.
> 4. ( ) Reavaliar a escolha com ≥ 4 rodadas (incluir 2026) e, na fase 2, avaliar D (regressão no nível do controle) para os analitos com viés proporcional.
> 5. ( ) Alternativa conservadora: **manter A** na fórmula até haver ≥ 6 rodadas, já com o rótulo de indicador e sem flexibilização.
> 6. Rejeitadas: B puro (cancelamento entre rodadas), C no Sigma (conta a dispersão inteira) e F (zerar viés não significativo com baixo poder).
>
> Responsável técnico: ____________________ Data: ___/___/_____

Até a assinatura: **produção inalterada (A)**.

## 13. Proposta de texto de correção para o ADR-064 (não aplicada)

Na seção "Pendência D-01", substituir:

> "...conta a imprecisão duas vezes e **subestima o Sigma** (Coskun 2022, texto completo; confirmado pelo revisor)... Recomendação: adotar a média com sinal + EP, como na triagem deste ADR."

por:

> "...com viés verdadeiro zero, a média de |bias| ainda vale ≈ 0,8 × a dispersão das amostras. É uma propriedade da normal dobrada, E|X| = σ·√(2/π), e não um resultado de Coskun 2022. Coskun 2022 (PMID 36277430) é uma carta que critica o RMS proposto por Ercan 2022 (PMID 35966254), lembra que viés é a média de réplicas menos o valor de referência e que sua significância deve ser avaliada, e questiona o uso linear do viés no Sigma. A carta não recomenda a média com sinal. O estudo D-01 (`estudos/D01_estudo_estimadores_bias.md`, 06/10/2026) mostrou que, com 3 rodadas, a média com sinal promove 20 linhas por cancelamento entre rodadas (todas com IC95% incluindo zero) e apaga a deriva comprovada do MCH. Recomendação revisada: G (RMS das médias de rodada) com transição controlada, submetida ao RT."

No `QUALITY_GATE.md` 14.9, trocar "Superestima o viés (≈0,8σ sem viés real)" por "Mistura erro aleatório com sistemático; o efeito é grande no diferencial leucocitário e pequeno onde há deriva entre rodadas (G/A mediano 1,00 Bio / 0,90 Hema). Estudo D-01 e decisão do RT pendentes."

## 14. Fora do escopo: o modelo linear de Sigma

Coskun et al. 2019 (PMID 30591816) e Coskun 2022 argumentam que somar o viés linearmente, σ = (TEa − |viés|)/CV, subestima o desempenho, e que o viés deveria entrar pela distribuição normal (ou ser corrigido quando conhecido). Esta decisão **não** muda o modelo. Ela escolhe o melhor número de viés **dentro** do modelo tradicional, que é o adotado por Westgard 2018 (PMID 30022879) e pela `tblPlanoQC_Sigma`. Rever o modelo seria outra mudança, com outro impacto nas faixas e nas regras, e deve ser tratada em pendência própria se o RT quiser.

## 15. Limitações do estudo

- Só **3 rodadas** de um provedor (CAP 2025) e 5 amostras por rodada. Todas as inferências entre rodadas têm baixo poder.
- O bootstrap reamostra dentro da rodada: mede a sensibilidade às amostras, não às rodadas. Por isso também apresento o "σB pior" pelo IC por rodada.
- A regressão é OLS (o alvo do CEQ, média de muitos laboratórios, tem erro pequeno frente ao resultado único). Deming não foi usado.
- O CV usado é o da planilha (período de 2026); o viés é de 2025.
- Analitos sem CEQ (Hemoglobina glicada, PCR, RDW-SD, RET, IRF) ficam sem Sigma com qualquer estimador.

## 16. Reprodutibilidade

```
python d01_estudo.py --bio <cópia de QC_Bioquimica.xlsm> --hema <cópia de QC_Hematologia.xlsm> [--saida <pasta>] [--boot 2000] [--seed 20261006]
```

O script lê só as cópias, não grava nada nos `.xlsm` e não contém caminhos de rede nem senhas. O stdout traz a validação, as matrizes, os diagnósticos e a tabela por linha. Os CSV (`;`, decimal com vírgula) trazem para S1 e S2: n, rodadas, A, B (com sinal, 1 e 2 etapas, DP, EP, IC95% por amostra e por rodada), C, F, G, D (a, b, IC, viés no nível), ANOVA, flags e, para cada estimador, Sigma, classe, regras, N, run size, frequência, ET, bootstrap e "σB pior".

## 17. Referências (verificadas no PubMed em 06/10/2026)

- Westgard S, Bayat H, Westgard JO. Analytical Sigma metrics: a review of Six Sigma implementation tools for medical laboratories. *Biochem Med* 2018;28(2):020502. PMID 30022879. [doi:10.11613/BM.2018.020502](https://doi.org/10.11613/BM.2018.020502)
- Ercan Ş. Bias estimation for Sigma metric calculation: arithmetic mean versus quadratic mean. *Biochem Med* 2022;32(3):030401. PMID 35966254. [doi:10.11613/BM.2022.030401](https://doi.org/10.11613/BM.2022.030401)
- Coşkun A. Bias, the unfinished symphony. *Biochem Med* 2022;32(3):030402. PMID 36277430. [doi:10.11613/BM.2022.030402](https://doi.org/10.11613/BM.2022.030402)
- Coskun A. Bias in laboratory medicine: the dark side of the moon. *Ann Lab Med* 2024;44(1):6-20. PMID 37665281. [doi:10.3343/alm.2024.44.1.6](https://doi.org/10.3343/alm.2024.44.1.6)
- Wauthier L, Di Chiaro L, Favresse J. Sigma metrics in laboratory medicine: a call for harmonization. *Clin Chim Acta* 2022;532:13-20. PMID 35594921. [doi:10.1016/j.cca.2022.05.012](https://doi.org/10.1016/j.cca.2022.05.012)
- Coskun A, Serteser M, Ünsal I. Sigma metric revisited: true known mistakes. *Biochem Med* 2019;29(1):010902. PMID 30591816. [doi:10.11613/BM.2019.010902](https://doi.org/10.11613/BM.2019.010902)
- Magnusson B et al. Nordtest TR 537 (incerteza top-down; base do ADR-067).
