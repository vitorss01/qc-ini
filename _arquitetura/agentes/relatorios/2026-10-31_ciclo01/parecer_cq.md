# Parecer do Especialista em CQ — ciclo 2026-10-31_ciclo01

**Agente:** Especialista em Controle de Qualidade · **Data:** 03/11/2026 · **Entrada:** `rota_cq.json` (33 ocorrências,
tipos calculo/dados/duvida_metodo) e os relatórios `<produto>_<frente>.md` deste ciclo.
**Limites respeitados:** não editei `src/`, instaladores nem planilhas. Não abri Excel. Li, com openpyxl em modo somente
leitura, a cópia já existente `%TEMP%\qcw\auditoria\2026-10-31_ciclo01\Bioquimica_F4.xlsm` (produção sha256 curto
`fe10dba1f1fd`). Não havia cópia da Hematologia; para ela usei os números dos relatórios.

## 0. Conferências independentes que fiz (Python puro, script no scratchpad da sessão)

| O quê | Resultado | Confere com |
|---|---|---|
| Bias de duas etapas da Bilirrubina direta (EQA_Base, CAP 2025, 15 amostras) | Com CHM-04/CHM-13 (resultado 0): rodadas 44,8571 / 21,6335 / 25,8222 → **30,770933248452** %. Sem elas: 31,0714 / 21,6335 / 7,2778 → **19,994209755156** % | G20 = 30,77093324845204 (Bio F3) e 19,99 % do relatório |
| Outras amostras de CEQ da Bio com resultado/alvo 0 ou NAO AVALIADO (Uso ≠ NAO, 2025) | **Só a Bilirrubina direta** (2 amostras) | — |
| ATIVO com PARTICIPA = SIM e RESULTADO ≤ 0 na tblCQ_Final da Bio | **56** (43 negativos, 13 zeros), todos ATIVO, o mais antigo de 05/12/2025 | 56 / 43 / 13 do Bio F1 |
| CVp de 6 m da Glicose N1, DIMENSION 2 (29/04–28/10/2026, lotes n ≥ 20) | **22,0341 %** (8974: n 45, CV 1,2064; 8977: n 84, CV 27,2416). Sem BIO-319726 (= 277): **1,6841 %** (8977: CV 1,8914) | 22,03 → 1,68 % do Bio F3; Sigma 0,33 → 4,32 (a razão fecha) |
| Glicose DIM1 lote 8977 N2, mês civil out/2026 | n 20, CV **34,27 %**; sem BIO-371183 (−1) e BIO-371184 (1): n 18, CV **2,08 %** | 30,88 % → 2,14 % do Bio F1 (a diferença é de janela: o F1 usou o período do Painel, não a fórmula) |
| Lotes de outro material com n ≥ 20 dentro de um analito sérico canônico (entrariam no CVp) | **Nenhum.** Glicose 5659 = 28 linhas em N1+N2 nos dois equipamentos (< 20 por chave); os lotes de urina 6353/6354 com n ≥ 20 estão com códigos de origem (CL, CRE2, K, NA, PHOS, MALB), fora do catálogo | Impacto da G5 limitado ao modo TODOS |
| Chaves (equip, lote, nível, analito, RUN) com > 1 participante, em maiúsculas, na produção | **0** (o defeito da G1 só aparece com digitação em minúsculas, o que ainda não aconteceu na produção) | — |

## Convenção usada

- O **veredito** fala do software e do método adotado.
- Nas ocorrências de **dados**, quando o software faz o que o ADR manda e o problema é o registro, a linha diz
  **"dado PROCEDE"**. A ação fica com o laboratório. Nesses casos a ocorrência não deve ser fechada sem a ação.
- **Autorização = true** em toda mudança de fórmula, critério estatístico, regra de negócio ou dado cadastrado (ETp,
  alvos, especificações, exclusão/inativação de amostras).
- **Tolerância:** 1e-9 relativo (REGRAS §3.5), salvo quando indicado.

---

## G1 — Manual com EQUIPAMENTO/ANALITO em minúsculas substitui um resultado real (#1 Bio F1, #16 Hema F1)

- **Veredito:** `ERRO_IMPLEMENTACAO`
- **Justificativa:**
  - O ADR-057 (decisão 6, `ADR-057.md:56-58`) garante **no máximo um participante por (equipamento, lote, nível,
    analito, RUN)**. É a condição para EstatBasica, AtualizarCalc e Westgard não divergirem.
  - O ETL monta a chave da corrida com o texto cru, em comparação binária:
    - `src/comum/mEtlCQ.bas:21` tem `Option Compare Binary`;
    - em `:1042-1046`, `gk` concatena `F_EQ`/`F_AN` sem normalizar.
  - O motor indexa em maiúsculas (`src/bio/mEstatistica.bas:210,213`, `GarantirIndice`/`EquipFiltro`). As duas
    camadas discordam sobre o que é "o mesmo analito/equipamento".
  - O manual recebe o RUN da 1ª corrida real e, no Eng_Saida (um valor por RUN × nível), toma o lugar do resultado
    real. A regra do ADR não muda; a implementação viola a garantia dele.
- **Impacto:**
  - Divergência Painel × Estatística: Bio n 24 × 25; Hema n 21 × 22.
  - O LJ mostra o ponto errado. O Westgard de uma corrida real foi de OK para REJEITADO, e o ET da linha de Sigma mudou
    de 53,69 para 51,67.
  - Vale para qualquer analito, nos dois produtos. Hoje não há caso na produção (0 chaves repetidas), mas a validação
    da Hema aceita "wbc" digitado.
- **Prioridade final:** **Crítico** (confirmada).
- **Correção:** no ETL (VBA e PQ, código comum), normalizar EQUIPAMENTO e ANALITO do manual para a grafia do cadastro
  (comparação sem diferenciar maiúsculas, gravando a forma canônica). Se não houver correspondência, gerar
  MANUAL_INCOMPLETO. Como defesa, o motor nunca sobrescreve um valor no mesmo RUN × nível em silêncio (QA de erro).
- **Autorização:** **false**. A correção restaura a garantia do ADR-057 e não muda nenhuma regra estatística.
- **Critério de aceite:**
  1. Na cópia, refazer E1–E3 (Bio: "dimension 1" e "glicose") e D1–D3 (Hema: "wbc"). O manual passa a ter o
     EQUIPAMENTO/ANALITO do cadastro e uma corrida própria (ou MANUAL_INCOMPLETO).
  2. O resultado real continua no Eng_Saida: BIO-371977 = 277 e HEM-371290 = 17,58.
  3. Painel n = Estatística n, e médias iguais (1e-9): Bio 25 / 247,7689; Hema 22.
  4. O veredito de Westgard do RUN real é igual ao da linha de base (OK).
  5. Invariante em toda a tblCQ_Final, com chave em maiúsculas: 0 chaves (equip, lote, nível, analito, RUN) com mais
     de um PARTICIPA = SIM.
  6. A suíte `equivalencia_cq.py` dá PQ = VBA.

## G2 — Data inválida colada (31/02/2026) vira 03/03/2026 e entra ATIVO (#3 Bio F1, #18 Hema F1)

- **Veredito:** `ERRO_IMPLEMENTACAO`
- **Justificativa:**
  - `src/comum/mEtlCQ.bas:211-223` (`ParaData`) e `:225-241` (`ParaDataHora`) usam `DateSerial`, que "rola" um dia
    inexistente para o mês seguinte, e não conferem dia, mês e ano depois da conversão.
  - O padrão do projeto para entrada ambígua é recusar: ADR-066 D07 ("1.5" em texto → MANUAL_INCOMPLETO + E04).
  - A validação da célula recusa a data digitada, mas a colagem passa por fora dela.
- **Impacto:** um resultado entra num mês e num RUN que não existem para ele. Isso altera a série do lote, o Westgard
  retroativo e a estatística de um mês já fechado. Afeta os dois produtos e os dois motores.
- **Prioridade final:** **Alto** (confirmada).
- **Correção:** depois de converter, exigir `Day = d`, `Month = m` e `Year = a` (também na hora: 0–23 e 0–59). Se
  divergir, MANUAL_INCOMPLETO com DETALHE "DATA". Fazer igual no PQ.
- **Autorização:** **false**.
- **Critério de aceite:**
  1. Colar, um caso por vez: 31/02/2026, 29/02/2026 (2026 não é bissexto), 00/01/2026, 32/01/2026, 15/13/2026 e
     "10/10/2026 25:00". Todos dão MANUAL_INCOMPLETO + E04 DATA, PARTICIPA = NÃO, e o n do nível não muda: Bio igual
     à linha de base; Hema N1 22 (não 23).
  2. 29/02/2028 e 28/02/2026 são aceitos.
  3. VBA e PQ dão o mesmo resultado (`equivalencia_cq.py`).

## G3 — Painel H7:H8/H9: ET = 1,65·CV quando o bias é "SEM EP" (#5 Bio F2, #7 Bio F3, #20 Hema F2, #23 Hema F3)

- **Veredito:** `ERRO_IMPLEMENTACAO` (regressão)
- **Justificativa:**
  - O ADR-030 corrigiu exatamente este defeito no Painel: "H7 e I7 … agora os dois exigem bias numérico"
    (`ARCHITECTURE_DECISIONS.md:973-976`).
  - O ADR-033 diz: "Sem EP não vira zero … H, L … ficam vazias" (`:1119`). A Estatística H obedece.
  - O guarda `ABS(IF(ISNUMBER($G7),$G7,0))` voltou pelo gerador `_arquitetura/etapa1_multilote/painel.py:705`, que
    reintroduziu a fórmula anterior ao ADR-030.
  - O I7 da linha 706 foi depois substituído pelo INDEX/MATCH do ADR-071 D2, e por isso só o H ficou errado.
  - A regra está certa; a implementação reintroduziu o zero.
- **Impacto:**
  - ET subestimado na linha principal do Painel de todo nível sem CEQ: Bio HbA1c e PCR (4 níveis), Hema RET% e RDW-SD
    (6 níveis).
  - Com o provedor **Controllab** (que não tem nenhuma linha), vale para **todos** os analitos dos dois produtos. Ex.:
    Proteína total 3,07/3,63 % contra o ET real com CAP de 6,85/7,42 %.
  - A mesma tela mostra outro ET (bloco ERRO TOTAL vazio).
  - Sigma e plano não são afetados (leem a Estatística). O Eng_Saida (`engPainel`) ainda calcula um Sigma oculto com
    bias 0 (ver G9).
- **Prioridade final:** **Crítico** (revisada; era Alto). Pela REGRAS §4, estatística incorreta exibida é Crítico. A
  correção é pequena.
- **Correção:** `H7 = IF(OR($E7="",NOT(ISNUMBER($G7))),"",$E7*1.65+ABS($G7))`, nos dois produtos (H7:H8 Bio, H7:H9 Hema).
  Corrigir também o gerador `painel.py` e os instaladores que reescrevem o Painel, para a regressão não voltar.
- **Autorização:** **false** (restaura o ADR-030).
- **Critério de aceite:**
  1. Com G não numérico, H fica `""`:
     - Bio HbA1c lote 8590 DIM2 out/2026 (eram 2,54/2,78);
     - Bio PCR lote 8353 (eram 2,37/2,01);
     - Bio Proteína total com M4 = Controllab (eram 3,069/3,632);
     - Hema R4 = Controllab, RBC e LYMPH#;
     - Hema RET% e RDW-SD.
  2. Com G numérico, H = 1,65·E + |G| (1e-9). Ex.: Proteína total CAP igual ao valor anterior à correção.
  3. O teste estático do instalador procura `IF(ISNUMBER($G` com `,0)` nas fórmulas do Painel e deve achar 0
     ocorrências.

## G4 — Resultados implausíveis (≤ 0, nível/material trocado, outliers) ativos sem nenhuma triagem (#2 Bio F1, #6 Bio F3, #17 Hema F1, #29 Hema F3)

- **Veredito:** `METODO_QUESTIONAVEL` (lacuna de regra) + **dado PROCEDE**
- **Justificativa:**
  - O software segue os ADR: a classificação do ADR-057 (`ADR-057.md:52-56`) não tem critério de plausibilidade, e o
    GUIA §6 deixa a revisão para o laboratório, pela aba Inativar.
  - Não inativar automaticamente é **correto**: a rejeição de corrida é do Westgard, e a exclusão de um resultado é ato
    técnico registrado com motivo, como pedem a ISO 15189:2022 e a CLSI C24.
  - O frágil é **não haver nenhum alerta**. O projeto já tem o critério escrito e usado num estudo:
    `estudos/cv_investigacao.py:9-10` (valor ≤ 0, ou |x − mediana do lote| / (1,4826·MAD) > 5).
  - Mesmo assim, 253/253 suspeitos de 13/10 continuam ativos, e há 56 ativos ≤ 0 na Bio desde 12/2025 (conferido).
  - A tela de digitação aceita RESULTADO = −1.
- **Impacto (reproduzido):**
  - Um único resultado de N2 gravado como N1 leva o CVp da Glicose N1 de 1,68 % para 22,03 % e o Sigma de 4,32 para
    0,33.
  - 53/58 níveis da Bio aparecem como "Desempenho inadequado", e o plano de CQ (regras, N, run size) é recomendado
    sobre um Sigma artificial.
  - u(Rw) de 12 m da Proteína total = 81,2 %.
  - Hema: o manual de teste −1 leva o WBC N1 de CV 1,736 % para 30,0 %. Os outliers da HGB N3 pesam +20 % na u(Rw)
    (#29, efeito pequeno).
  - Os vereditos operacionais (classe Sigma e plano) mudam em quase todos os analitos da Bio.
- **Prioridade final:** **Crítico** (#2, #6). Os itens #17 (Alto) e #29 (Baixo) ficam dentro do grupo.
- **Correção:**
  1. **Laboratório (agora):** revisar e inativar com motivo, pela aba Inativar, os 56 IDs ≤ 0 da Bio e a lista da
     F3 (287 na janela de 6 m), começando por BIO-319726, BIO-215815/215820, BIO-232288/232289, BIO-268082 e
     BIO-371183/371184. Hema: HEM-241016, 309465/309466 e 314337/314338. Consultar o SIPEC sobre os valores −1/0/1
     sem flag.
  2. **Desenvolvimento (com autorização):** um código novo no QA_INTEGRACAO, do tipo ALERTA, com o ID, para:
     - RESULTADO < 0;
     - RESULTADO = 0, exceto nos analitos marcados como "aceita zero" (ex.: Hema NRBC, IG, BASO);
     - |z robusto| > 5 contra a mediana do lote × nível × equipamento (o mesmo critério do estudo).

     O alerta **nunca** muda PARTICIPA. Na tela de digitação, RESULTADO < 0 dá MANUAL_INCOMPLETO.
- **Autorização:** **true** (inativação de amostras e regra nova de triagem).
- **Critério de aceite:**
  1. Na cópia `fe10dba1f1fd`, o alerta lista exatamente os 56 IDs ≤ 0 da Bio (43 < 0, 13 = 0), em igualdade de
     conjunto com uma varredura independente. A lista de |z| > 5 é igual à de `cv_investigacao.py` na mesma janela.
  2. n e PARTICIPA ficam idênticos antes e depois da instalação.
  3. Manual Hema WBC = −1 dá MANUAL_INCOMPLETO.
  4. Depois da inativação simulada na cópia:
     - Glicose N1 DIM2: CVp 6 m = 1,6841 % (1e-9 contra o independente) e Sigma = (ETp − |bias|)/1,6841;
     - Glicose DIM1 8977 N2, out/2026 civil: CV 2,0849 %;
     - Hema WBC N1 volta a 1,736 %.

## G5 — Lotes de outro material (urina, PCR) ativos como analito sérico, com alvo tirado do próprio valor (#13 Bio F4, #15 Bio F4)

- **Veredito:** `METODO_QUESTIONAVEL` (nenhuma trava de lote × analito/matriz) + **dado PROCEDE**
- **Justificativa:**
  - A origem manda o controle de urina como GLUC/URE com MATRIZ = SORO. Conferido: 28 linhas da Glicose no lote 5659,
    2 da Ureia no 6353 e os 14 da HbA1c no 8974.
  - O recebimento cria o lote e o preenchimento de TESTE do ADR-076 dá a ele um alvo tirado do próprio valor
    (871 ± 1,71). Isso faz o ponto parecer "no alvo".
  - O modo TODOS do ADR-086 junta todo lote do analito.
  - Nenhum ADR exige que o lote pertença ao analito (bula/LotesStore). A lacuna é de regra.
- **Impacto:**
  - Não chega ao CVp, à u(Rw) nem à estatística por lote: conferi que nenhum desses lotes atinge n ≥ 20 por chave em
    analito canônico (ADR-071 D5).
  - Afeta o modo H3 vazio: n e Westgard inflados, e o eixo Y vai a 918/1.932 na Ureia dez/2025.
  - Com o lote escolhido em H3, nada muda.
- **Prioridade final:** **Alto** (confirmada).
- **Correção:**
  1. **Laboratório:** inativar com motivo os IDs dos lotes 5659/6353/6354/8353 nos analitos séricos e os 14 de HbA1c
     no 8974, e pedir a correção da matriz/código no SIPEC.
  2. **Com autorização:**
     - ALERTA no recebimento quando o lote não tem Média/DP de bula para o analito na LotesStore e o alvo veio do
       preenchimento de TESTE com n < 20;
     - ou uma lista de lotes permitidos por analito.
     - O modo TODOS avisa quando junta um lote marcado como TESTE com n < 20.
- **Autorização:** **true**.
- **Critério de aceite:**
  1. Na cópia, o alerta lista: Glicose 5659 (28 linhas), Ureia 6353 (BIO-230952/230958), Creatinina/Na/K/Cl/P/Ca em
     6353/6354, Glicose e Ureia em 8353 e HbA1c em 8974 (14). Não lista os lotes 8974/8977 dos analitos séricos.
  2. Depois da inativação simulada, Ureia DIM1 dez/2025 em H3 vazio não tem nenhum ponto com valor > 3 DP do alvo do
     seu lote, e o eixo Y fica na faixa do controle sérico.

## G6 — Alvo do lote 522611 incompatível com os dados em 12 pares (#22 Hema F3)

- **Veredito:** `METODO_QUESTIONAVEL` (sem verificação de plausibilidade do alvo; o aviso do ADR-076 só cobre os pares
  de TESTE) + **dado PROCEDE**
- **Justificativa:**
  - Um RDW-SD de 14,78 fL não é fisiológico para material de controle (a média medida é 47,38 fL, z = +398).
  - O IG# tem alvo na ordem de grandeza do IG% (z ≈ −30).
  - A série vermelha do N2 está a −5,7 a −14,6 DP do alvo.
  - Cadastro de nível/analito trocado é "associação errada entre analitos/níveis" (REGRAS §4). O cálculo está certo
    sobre um dado errado.
- **Impacto:**
  - Limites do LJ e eventos de Westgard do lote 522611: 6.155 eventos, com o topo exatamente nos 12 pares (IG# 714,
    dos quais 152 de 1_3s).
  - A Estatística C:W não usa o alvo e não é afetada.
  - O lote é antigo (out/2025, validade 31/03/2026). O efeito é sobre registros oficiais de Westgard
    (Westgard_Fotos/Eventos), não sobre a rotina atual.
- **Prioridade final:** **Alto** (confirmada).
- **Correção:**
  1. **Laboratório:** conferir na bula do lote 522611 os valores de IG# N1–N3, RDW-SD/RDW-CV N2, HGB/MCH/MCV/HCT N2 e
     RET%/RET# N1, e corrigir pela aba Analitos, com trilha. Depois, regerar os eventos do lote.
  2. **Com autorização:** ALERTA quando |média do lote − alvo| / DP do alvo > 5, com n ≥ 20.
- **Autorização:** **true** (alvo cadastrado).
- **Critério de aceite:**
  1. Depois da correção, nenhum par do lote 522611 tem |z| > 5 (hoje são 12).
  2. O alerta (se autorizado) marca exatamente esses 12 pares antes da correção e 0 pares nos outros 8 lotes.
  3. Eventos_Westgard do 522611 regerada: IG# 1_3s cai de 152.

## G7 — O bias do ET/Sigma inclui amostras que o próprio projeto exclui da incerteza: resultado 0 e NAO AVALIADO (#8 Bio F3, #28 Hema F3)

- **Veredito:** `METODO_QUESTIONAVEL`
- **Justificativa:**
  - `mCEQ.BiasEQ` (`src/comum/mCEQ.bas:697-837`) faz o que o ADR-030/035 manda: média das |bias| por rodada, depois a
    média das rodadas, sobre **todas** as amostras com Uso ≠ NAO. Não é defeito.
  - Mas a auditoria do ADR-064 (A-01, `ADR-064.md:100`; e A-09, `:108`) reconheceu que amostra NAO AVALIADO ou com
    resultado/alvo 0 "distorce" (±100 % perto do limite de detecção), e corrigiu **só** o `ViesEQ` (`mCEQ.bas:1192-1197`).
  - O projeto aplica dois critérios de seleção diferentes ao mesmo CEQ, e o mais frágil alimenta justamente ET, Sigma e
    plano.
  - CHM-04 tem alvo 0,05 mg/dL com DP do grupo 0,05 (CV do grupo 100 %), abaixo de qualquer faixa quantificável. Um
    bias de −100 % ali descreve a resolução do laudo (0,1 mg/dL), não o método.
  - A evidência técnica é forte e vem do próprio projeto, por isso reabro.
- **Impacto (reproduzido):**
  - **Bio:** só a Bilirrubina direta. Bias 30,7709 → 19,9942 %; ET N1 37,1 → 26,3 %; Sigma N1 0,43 → 0,95 (CVp
    21,03 %). A classe Sigma não muda (< 3), mas o status da margem sai de "Margem crítica" para "dentro do orçamento"
    (ETp 39,92 %).
  - **Hema:** IG%, IG#, NRBC% e NRBC#, todas com as 15 amostras NAO AVALIADO. G = 3,0094 % vira SEM EP, e o ET exibido
    (7,5–8,1 %) some. Não há Sigma afetado (sem ETp, ADR-068).
- **Prioridade final:** **Alto** (#8; o #28 era Baixo e entra no grupo).
- **Correção:** no `BiasEQ`, para todos os modos (ABS, SIGNED, N, NRODADAS, DETALHE), aplicar o mesmo filtro de amostra
  do `ViesEQ`: fora NAO AVALIADO, resultado e alvo ≠ 0, cada chave uma vez. As duas etapas e a média de |bias| ficam
  como estão (decisão D-01 não reaberta).

  Recomendação adicional, a decidir pelo RT: um limite inferior de concentração por analito para o bias % (ex.: amostras
  com alvo abaixo do limite de quantificação declarado). Fica fora desta correção.
- **Autorização:** **true** (critério de seleção de amostras do bias).
- **Critério de aceite:**
  1. **Bio:**
     - G20 = 19,994209755156 % (1e-9);
     - AC20 = 13 amostras, AD20 = 3 rodadas;
     - H20/H21 = 1,65·F + 19,9942;
     - L20 = (39,9151 − 19,9942)/CVp.
  2. As outras 28 linhas de analito da Bio com CEQ ficam com G, T, AC e AD **idênticos** aos de antes (diferença 0).
  3. **Hema:** G de IG%, IG#, NRBC% e NRBC# = "SEM EP" e H = ""; os demais analitos ficam idênticos.
  4. O `qa_incerteza.py` I04 continua passando (o ViesEQ não muda).

## G8 — Resumo "Sem EP" conta 22 níveis, mas são 4 (#9 Bio F3)

- **Veredito:** `ERRO_IMPLEMENTACAO`
- **Justificativa:**
  - Estatística B103 = `COUNTIF($G$14:$G$93,"SEM EP")` conta as 18 linhas sem analito (A76:A93 = 0), onde
    `BiasEQ("0")` devolve SEM EP.
  - O ADR-064 A-08 (`ADR-064.md:107`) já fixou o princípio de que linha sem analito não produz indicador. O resumo
    ficou de fora.
- **Impacto:** é um indicador de gestão: sugere que um terço da rotina está sem CEQ. Não muda o cálculo.
- **Prioridade final:** **Médio** (confirmada).
- **Correção:** `B103 = COUNTIFS($G$14:$G$93,"SEM EP",$A$14:$A$93,"<>0",$A$14:$A$93,"<>")`. Conferir as demais
  contagens do resumo (B98:B102, F98:F100) e o equivalente da Hema (A98:A133 = 0, OC-H11).
- **Autorização:** **false**.
- **Critério de aceite:**
  1. Na cópia `fe10dba1f1fd` com CAP, B103 = 4 (HbA1c e PCR × 2).
  2. Com o provedor Controllab (sem dados), B103 = 62 (31 analitos × 2).
  3. As classes 0/0/3/2/53 e o orçamento 46/2/10 ficam inalterados.

## G9 — Vários Sigmas para o mesmo analito; ET e Sigma com CVs diferentes; Sigma legado no Eng_Saida (#10 Bio F3, #24 Hema F3, #30 Hema F3)

- **Veredito:** `DIFERENCA_METODO` (defensável), com limpeza de saída legada
- **Justificativa:**
  - **O Sigma oficial é um só e está documentado:** Estatística L = (ETp − |bias|)/CVp de N meses (ADR-071,
    `ADR-071.md:48-58`). O Painel I7 e o `Cfg_PlanoQC!B1` leem esse valor (D2).
  - **ET com o CV do período e Sigma com o CVp:** é decisão registrada (ADR-071 D6, `:55`). Ela é defensável:
    - o ET do período mede o desempenho no período;
    - o Sigma de 6 m mede a capacidade do método para escolher o plano.
  - As 41 contradições "dentro do orçamento × inadequado" são, na maior parte, consequência do CVp contaminado (G4).
    Por isso não há, hoje, evidência para reabrir a D6.
  - **Bloco AU:BD da aba Analitos (ADR-074):** CV = DP/Média do **cadastro**. É uma grandeza de planejamento, diferente
    do Sigma medido, e o rótulo não diz isso. Nos 21 pares de TESTE (ADR-076), o DP é 1,5 × o real, e o Sigma do bloco
    sai 1,5 vez menor.
  - **Eng_Saida I185:I187 (`engPainel`) e L190:L309 (`engEstat`):**
    - calculam Sigma com o bias **contra o alvo do lote**, que o ADR-030 rejeitou como bias de ET/Sigma
      (`ARCHITECTURE_DECISIONS.md:922-940`);
    - também calculam com bias 0 quando falta CEQ (4,7585 na HbA1c, Bio F4);
    - nenhuma fórmula nem o `mBI` lê essas colunas (grep em `src/*/mBI.bas` sem referência). São saídas mortas com uma
      definição proibida.
- **Impacto:** nenhum veredito operacional muda hoje. O plano lê L. O risco é de decisão sobre o número errado: os
  Sigmas diferem em classe (MONO% N3: 1,79 × 4,82; Proteína total N1: 4,00/2,50/5,03 × 1,99).
- **Prioridade final:** **Médio** (#10, #24); o #30 era Baixo.
- **Correção:**
  1. Rotular o bloco AU:BD como "Sigma (CV do cadastro) — informativo, não é o Sigma do plano".
  2. Nos pares de TESTE, o bloco mostra "TESTE" no lugar de ET/Sigma.
  3. Retirar (ou esvaziar com o rótulo LEGADO) Eng_Saida G:I 185:187 e o Sigma/classe do `engEstat`.
  4. Reavaliar a D6 depois da limpeza da G4. Se, com o CVp limpo, mais de 10 % dos níveis continuarem "dentro do
     orçamento" e "inadequados" ao mesmo tempo, propor ao usuário que o status de orçamento use o CVp. Isso exige
     autorização.
- **Autorização:**
  - **false** para os rótulos, para o "TESTE" no bloco e para a retirada das colunas legadas sem leitor;
  - **true** para mudar o CV do ET/orçamento ou retirar o Sigma do bloco (ADR-074).
- **Critério de aceite:**
  1. Painel I7 = Estatística L (1e-9) nos dois produtos.
  2. O cabeçalho do bloco contém "CV do cadastro".
  3. Nos 21 pares de TESTE da Hema (lista `teste_pairs.py`), AX/BA/BD = "TESTE".
  4. Eng_Saida I185:I187 e L190:L309 ficam vazios ou rotulados.
  5. A busca no XML de todas as abas e no `src/` não acha nenhuma leitura deles.
  6. A Estatística C:AU fica idêntica à de antes (diferença 0).

## G10 — Biblioteca de especificações 2026–2029 incompleta (Bio: Lipase VB e PCR CLIA) ou vazia (Hema) (#11 Bio F3, #27 Hema F3)

- **Veredito:** `NAO_PROCEDE` quanto ao software + **dado PROCEDE**
- **Justificativa:**
  - Sem especificação, o ETp vira "-" e o Sigma some. É o comportamento decidido no ADR-068 (`ADR-068.md:14`), e o
    ADR-071 item 7 já marca "⚠ SIGMA EM k DE N NÍVEIS".
  - A lacuna é de cadastro (GUIA §6: o laboratório mantém a biblioteca de cada ano).
- **Impacto:** nenhum hoje (todos os lotes têm o ano de referência 2025). Quando um lote nascer com 2026, a Lipase e a
  PCR (Bio) e **todos** os analitos (Hema) perdem ETp, Sigma e plano.
- **Prioridade final:** **Médio** (confirmada).
- **Correção:**
  1. O laboratório preenche 2026 antes do primeiro lote com esse ano. Copiar 2025 só se a fonte não mudou, e registrar
     a fonte.
  2. Opcional, sem mudança de regra: um aviso no QA/bloco do lote quando o ano escolhido não tem a especificação da
     fonte do analito.
- **Autorização:** **true** (dado cadastrado); false para o aviso.
- **Critério de aceite:**
  1. Para cada analito e cada ano de 2025 a 2029, a especificação da fonte em uso (Especificações S) é numérica e
     maior que 0, ou existe um registro de "sem especificação" decidido.
  2. Simulando na cópia um lote com ano 2026, Estatística K de Lipase e PCR (Bio) e de WBC (Hema) é numérico.

## G11 — Janela de 180 corridas: "Tudo (sem filtro)" e o H3 vazio mostram n/estatística/contagens só das 180 corridas mais recentes, e o aviso fala só do gráfico (#12 Bio F4, #32 Hema F4)

- **Veredito:** `DIFERENCA_METODO` (a janela é decisão documentada; o aviso é insuficiente)
- **Justificativa:**
  - A janela de 180 corridas é desenho do motor:
    - ADR-049 decisão 5: "a faixa do Painel diz quantas ficaram de fora" (`ARCHITECTURE_DECISIONS.md:2611-2613`);
    - ADR-066 D10: o n, a média e o DP do Painel são da janela;
    - ADR-086 decisão 6: o modo TODOS trabalha com no máximo 180 corridas.
  - O que falha é a transparência: o texto A39 diz "fora do gráfico", mas o corte também vale para n, média, DP, CV e
    as contagens de Westgard.
  - A estatística oficial (aba Estatística, por período) não tem a janela.
  - O número "do lote inteiro" do relatório (Glicose 8974 N1, DP 30,37) também está contaminado por resultados
    implausíveis (G4).
- **Impacto:** os descritivos e as contagens do Painel não cobrem o período pedido quando ele passa de 180 corridas. ET,
  Sigma e plano não são afetados.
- **Prioridade final:** **Alto** na Bio (#12, o rótulo "Tudo" promete o lote inteiro); **Médio** na Hema (#32).
- **Correção:** trocar o A39 por "N corrida(s) fora do gráfico **e dos indicadores do Painel** (máx. 180). Período
  inteiro: n = X — veja a aba Estatística". Calcular os descritivos do Painel sobre o período inteiro seria mudança de
  regra e de desempenho, e depende de autorização.
- **Autorização:** **false** para o aviso; **true** para mudar o cálculo.
- **Critério de aceite:**
  1. Com H3 = 8974 e K4 = "Tudo" (Glicose DIM1), o A39 contém "indicadores" e o n total do período (174 no N1).
  2. Com a Hema WBC em H3 vazio e "Ano inteiro 2026", o aviso mostra n total 337/340/308.
  3. Os números do Painel ficam iguais aos de antes (140 / 89,15 / 16,968 e 141/155/130).

## G12 — No H3 vazio o Painel mostra Bias % e Sigma (e o Sigma do plano), contra o ADR-086 decisão 4 (#14 Bio F4, #33 Hema F4)

- **Veredito:** `ERRO_IMPLEMENTACAO` (a tela não segue o texto do ADR-086 dec. 4, `ADR-086.md:52-53`), com a
  recomendação de **corrigir o ADR**, não a tela
- **Justificativa técnica:**
  - O motivo do ADR-086 ("juntar lotes não tem sentido metrológico") vale para média, DP, CV e ET do período, porque
    misturam materiais com alvos diferentes.
  - Não vale para o **Bias**: ele vem do CEQ e é do método, não do lote de controle.
  - Nem para o **Sigma**: o CVp já é agrupado por lote pelos graus de liberdade, em todos os lotes da janela (ADR-071),
    e não depende de H3.
  - Os dois são indicadores do método, metrologicamente válidos no modo TODOS. O ET (H) usa o CV do lote, e por isso
    fica vazio.
- **Impacto:** a linha fica incoerente (Sigma sem ET) e o ADR diverge da tela. Nenhum veredito operacional muda.
- **Prioridade final:** **Médio** (confirmada).
- **Correção:**
  - **Opção A (recomendada):** emendar o ADR-086 dec. 4. No modo TODOS ficam Bias e Sigma, com o rótulo "do método —
    independem do lote (Sigma: CVp N m; ETp do lote em uso)"; continuam vazios Média, DP, CV e ET.
  - **Opção B:** esvaziar G:I (e O4 "Sigma do plano") no modo TODOS, como o ADR diz hoje.
- **Autorização:** **true** para a Opção A (muda uma decisão documentada); false para a Opção B.
- **Critério de aceite (A):**
  1. Em H3 vazio, Painel G = Estatística G e I = Estatística L (1e-9). Ex.: Hema WBC fev–mar/2026, G = 1,86 % e
     I = 2,27/3,58/5,68.
  2. C:E e H ficam vazios.
  3. O rótulo contém "do método".
  4. O ADR-086 é atualizado.

  **Critério de aceite (B):** em H3 vazio, C:I e O4 ficam vazios nos dois produtos.

## G13 — Troca de lote no Painel deixa a Estatística com o B4 novo e os números do lote anterior (#19 Hema F2)

- **Veredito:** `ERRO_IMPLEMENTACAO`
- **Justificativa:**
  - `TrocarLoteAnalise` (`src/hema/mLotes.bas:398-444`; a linha 38 do procedimento) chama só
    `RecalcularAnalitoAtual`, que não chama `AtualizarEstatisticaAba` (`src/hema/mEstatistica.bas:2089`).
  - Na Hema, a Estatística C:F é gravada pelo VBA a partir de B4. O B4 muda e os valores não: números de um lote com o
    rótulo de outro. O ADR-035 ("o Painel não recalcula nada; lê a Estatística") pressupõe uma Estatística coerente
    com o filtro.
  - Na Bio, a mesma rotina chama a mesma sub (`src/bio/mLotes.bas`), mas a Estatística da Bio usa UDF que depende de
    H4. É provável que recalcule; o Tester deve confirmar.
- **Impacto:** n, média, DP, CV, ET, margem e status do lote anterior, exibidos como sendo do lote novo (RBC N1: n 50,
  CV 1,2595, ET 4,478 do 522611, rotulados 619611). Isso é associação errada entre lotes.
- **Prioridade final:** **Crítico** (revisada; era Alto). REGRAS §4: "associação errada entre lotes".
- **Correção:** na troca de lote, chamar `AtualizarEstatisticaAba`, ou marcar a aba como desatualizada, com aviso
  visível e os números ocultos, até o próximo ATUALIZAR.
- **Autorização:** **false**.
- **Critério de aceite:**
  1. Hema, H3 522611 → 619611: Estatística B4 = 619611 e RBC N1 n = 54, CV = 1,1344 %, ET = 4,272, iguais ao Painel e
     ao independente (1e-9).
  2. De volta ao 522611: n = 50, CV = 1,2595 %, ET = 4,478.
  3. O mesmo roteiro na Bio (H3 8974 ↔ 8977) dá C:H iguais ao independente.

## G14 — 90 linhas "Controllab" de simulação na tabela do CAP gravadas como CAP na EQA_Base (#21 Hema F2)

- **Veredito:** `NAO_PROCEDE` (comportamento decidido)
- **Justificativa:**
  - O ADR-034 decidiu manter as 90 linhas de simulação visíveis, com Uso_Analitico = NAO, e não apagar histórico
    (`ARCHITECTURE_DECISIONS.md:1205-1212`).
  - O provedor vem da aba de digitação (`mEQA.Consolidar`, `src/comum/mEQA.bas:308-342`). Hoje elas não entram em
    nenhuma conta: o Bio F3-T12 provou AD = 3 rodadas, e não 6.
  - O risco existe só se alguém trocar o Uso para SIM. O ADR-064 A-11 preserva essa decisão manual por chave.
- **Impacto:** nenhum hoje.
- **Prioridade final:** **Baixo** (revisada; era Médio).
- **Correção (salvaguarda, opcional):** um aviso na consolidação quando a coluna A diverge do provedor da aba, e
  recusar Uso = SIM nessas linhas.
- **Autorização:** **false** para o aviso; **true** para apagar ou mover as linhas (dado).
- **Critério de aceite:**
  1. A consolidação relata "90 linhas com provedor da coluna A ≠ aba (CAP)".
  2. Com Uso = SIM forçado numa dessas linhas, o Uso volta a NAO (ou aparece ERRO), e o bias do CAP de WBC continua
     −1,86 % (inalterado).

## G15 — RDW: o CAP avalia o RDW-CV, mas a VB está só no RDW-SD; nenhum dos dois tem Sigma (#25 Hema F3)

- **Veredito:** `INCONCLUSIVO` (depende do documento-fonte do CVi/CVg)
- **Justificativa:**
  - O de/para do CAP leva "RDW" (%) para o RDW-CV, o que é coerente com a unidade em %.
  - Os CVi 1,6 / CVg 4,2 estão no RDW-SD, mas não sei afirmar, sem a fonte cadastrada, se o dado de variação biológica
    é de RDW-CV ou de RDW-SD. Não invento a referência.
  - Se a fonte for de "RDW (%)", a especificação está no analito errado. Se for de RDW-SD, falta o CEQ do RDW-SD (o que
    é normal: o CAP avalia RDW em %).
- **Impacto:** o RDW fica sem Sigma e sem plano de CQ. Nenhum número errado é exibido.
- **Prioridade final:** **Médio** (confirmada).
- **Correção:** o laboratório confere na fonte (base EFLM de variação biológica, ou a usada em 2025) a que medida se
  referem CVi/CVg, registra a fonte e, se for RDW (%), move a especificação para o RDW-CV.
- **Autorização:** **true** (especificação cadastrada).
- **Critério de aceite:** com a especificação no RDW-CV, Estatística K(RDW-CV) = ETp VB recalculado (1,65·f·CVi +
  f/2·√(CVi² + CVg²), com o f do desempenho escolhido; 1e-9) e L(RDW-CV) numérico. Com a decisão de manter no RDW-SD,
  registrar o motivo no cadastro.

## G16 — ETp "CLIA" 2025 de WBC (7 %) e PLT (15 %) diferente do CLIA (#26 Hema F3)

- **Veredito:** `INCONCLUSIVO` (provável divergência de cadastro ou de rótulo)
- **Justificativa:**
  - Pelo meu conhecimento, a regra final do CLIA para hematologia (42 CFR 493.941, publicada em 2022 e em vigor a
    partir de 2024/2025) fixa WBC ±10 %, PLT ±25 % e RBC/HGB/HCT ±4 %. Os critérios anteriores eram WBC ±15 %,
    PLT ±25 %, RBC ±6 %, HGB ±7 % e HCT ±6 %.
  - O cadastro tem RBC/HGB/HCT = 4, iguais à regra atual, mas WBC = 7 e PLT = 15 não batem com nenhuma das duas
    versões.
  - Não tenho como abrir o texto oficial neste ambiente. A confirmação na fonte é obrigatória.
  - O QC_INI usa corretamente o que está cadastrado.
- **Impacto:** se o valor for engano, o Sigma de WBC (2,27/3,58/5,68 com ETp 7) e o de PLT ficam subestimados, e o
  plano de CQ fica mais pesado do que o necessário. Se for meta interna, o rótulo "CLIA" engana.
- **Prioridade final:** **Médio** (confirmada).
- **Correção:** o laboratório confere no 42 CFR 493.941 vigente. Se for meta própria, troca a fonte para FAB/meta
  interna, com registro. Se for erro, corrige para 10 / 25.
- **Autorização:** **true** (ETp cadastrado).
- **Critério de aceite:**
  1. Especificações B48 e I48 = valor da fonte declarada, com a fonte registrada.
  2. Estatística K14:K16 e K35:K37 = o novo ETp.
  3. L = (ETp − |bias|)/CVp (1e-9). Com ETp 10, o WBC N1 passa a (10 − 1,86)/CVp.

## G17 — "FAB CVTp" 2025 da Hema com valores que não são CV % (#31 Hema F3)

- **Veredito:** `NAO_PROCEDE` quanto ao software + **dado PROCEDE**
- **Justificativa:** RDW-SD 48,1, RDW-CV 15,1, MPV 9,4 e RET% 1,2 têm cara de valor de controle (ADR-075 moveu o antigo
  "CV Fab" para a linha 53). Nenhum analito usa FAB, e a linha "FAB ETp" está vazia.
- **Impacto:** nenhum hoje. Se a fonte FAB for escolhida, o CVTp de 48,1 % fica absurdo.
- **Prioridade final:** **Baixo** (confirmada).
- **Correção:** o laboratório revisa ou esvazia X53:AA53 antes de usar FAB.
- **Autorização:** **true** (especificação cadastrada).
- **Critério de aceite:** X53:AA53 vazios ou com CV % plausível (< 30 %) e fonte registrada.

## G18 — Quatro inativações sem motivo fazem toda atualização acusar E01 (#4 Bio F1)

- **Veredito:** `NAO_PROCEDE`
- **Justificativa:** o E01 é o comportamento correto (`src/comum/mEtlQA.bas:194-203`). Inativar sem justificativa
  técnica fere a rastreabilidade (ISO 15189), e o alarme deve persistir até o motivo ser registrado. Não é ruído: são 4
  pendências reais (BIO-366027/366033/366054/366164, Lactato N2, 13/10).
- **Impacto:** fadiga de alarme enquanto a pendência existir.
- **Prioridade final:** **Baixo** (revisada; era Médio).
- **Correção:** o laboratório preenche o MOTIVO desses 4 IDs na aba Inativar. Não é preciso mudar o software.
- **Autorização:** **false** (ação do laboratório, sem mudança no sistema).
- **Critério de aceite:** depois do motivo registrado, o ATUALIZAR na cópia dá 0 E01, e o E01 volta a aparecer ao se
  inativar um ID de teste sem motivo.

---

## Resumo das prioridades revisadas

- **G3** (ET do Painel) e **G13** (troca de lote): sobem de Alto para **Crítico**.
- **G14** (simulação na tabela do CAP) e **G18** (E01): descem para **Baixo**.
- **G7:** o #28 sobe de Baixo para Alto por estar junto do #8.
- As demais ficam como estavam.

## Ordem sugerida

1. Correções sem autorização: G1, G3, G13, G2, G8. Rótulos e legado da G9, aviso da G11.
2. Pedir ao usuário, em um único pacote: G4 (triagem + inativações), G7 (filtro do BiasEQ), G5/G6 (travas e cadastro),
   G12 opção A.
3. Laboratório: inativações (G4/G5), bula do 522611 (G6), biblioteca 2026 (G10), fontes de RDW/CLIA/FAB (G15–G17),
   motivos do E01 (G18).
4. Reavaliar a D6 (CV do ET) depois da limpeza da G4.
