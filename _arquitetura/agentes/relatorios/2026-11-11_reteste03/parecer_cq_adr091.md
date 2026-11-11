# Parecer do Especialista em CQ: ADR-091 (o lote define o material)

Ciclo: 2026-11-11_reteste03 · Data: 11/11/2026 · Produto: QC_Bioquimica (a Hema não tem regra cadastrada)

## 0. Como o parecer foi feito (evidência)

- **Sem Excel.** A produção da Bio (sha `8b1cda1345d2d6ff`) foi copiada com `shutil.copy2` para
  `%TEMP%\qcw\auditoria\2026-11-11_reteste03\cq_parecer\`. A cópia tem o mesmo sha. Ela foi lida com `openpyxl` em
  modo `read_only, data_only`, sem gravar nada. Foram lidas a `tblCQ_Final` (aba "Principal - Resultados",
  A6:AN35472, 35.466 linhas), a `tblLotesMaterial` (Cfg_Integracao L6:P23), a `tblDeParaAnalitos` e a
  `tblQA_Integracao`. O sha da produção foi conferido de novo no fim e não mudou. A cópia foi apagada depois.
- `relatorios/efeito_adr091/Bioquimica.json` e `Hematologia.json`: participantes por lote|analito antes e depois, e
  contagem por STATUS, do ATUALIZAR de 11/11.
- `relat209 (14).csv` de 05/11/2026 (o mesmo arquivo da análise do ADR-091; cópia local, não versionada). Só foram
  feitas agregações por lote, sufixo e mne. Nenhum dado de paciente foi lido nem registrado.
- `config_produtos.py`: a carga inicial (17 regras) é igual, linha por linha, à tabela da produção.
- Os scripts são Python puro, rodaram com `python -I` e ficaram no scratchpad da sessão. Eles fazem a reconciliação
  do JSON, a matriz por lote e sufixo, as medianas por grupo de sufixo e a busca de lotes residuais.

## 1. Veredito: **VALIDADA** (a regra está tecnicamente correta e foi aplicada como o ADR manda)

### 1.1 Urina vs soro (regra MATRIZ): correta, e a evidência numérica é forte

O sufixo `-EE` é digitado pelo técnico, e o relat209 confirma que ele não é confiável. O 6354 tem 61 resultados com
sufixo de soro e 711 com sufixo de urina. O 5659 tem 98 de soro e 33 de urina. O 6351 tem 248 de soro e 8 de urina.
O que decide a questão é a **magnitude** dos resultados com sufixo de soro nos lotes de urina. Eles são de urina,
não de soro:

| Analito (código) | Lote/nível | Sufixo 01/02 (mediana) | Sufixo 03/04 (mediana) | Soro 8977 (mediana) |
|---|---|---|---|---|
| CRE2 (mg/dL) | 6354 N1 / N2 | 73,99 (n 1) / 235,95 (n 7) | 76,28 (n 68) / 236,94 (n 62) | 2,66 / 6,46 |
| K (mmol/L) | 6354 N1 / N2 | 21,9 / 123,4 | 22,0 / 119,2 | 3,8 / 6,4 |
| NA (mmol/L) | 6354 N2 | 164 | 163 | 130 |
| PHOS (mg/dL) | 6354 N2 | 82,73 | 83,54 | 7,15 |
| CRE2 | 6353 N1 / N2 | 76,87 / 235,28 | (sem 03/04) | — |
| GLUC (mg/dL) | 5659 N1 / N2 | 59 / 33 | 59 / 33 | 77 / 275 |

Os resultados com sufixo de soro são estatisticamente indistinguíveis dos que têm sufixo de urina. Eles ficam a uma
ou duas ordens de grandeza do soro: a creatinina de 76 mg/dL fica a cerca de 28 vezes o nível 1 sérico. Antes do
ADR-091, esses valores entravam em "Creatinina", "Potássio", "Glicose" etc. do soro como outliers grosseiros. Inflavam
o DP e o CV, derrubavam o Sigma e disparavam regras de Westgard (1-3s e R-4s) sem causa analítica. A associação errada
entre material e analito é Crítico pela §4. A correção é a esperada pela boa prática: o CIQ é avaliado por
**material/lote de controle**, cada um com sua média e seu DP (CLSI C24-Ed4, 2016; ISO 15189:2022, 7.3.7.2). Por
isso, o identificador confiável do material deve prevalecer sobre um campo digitado e inconsistente.

Conferência na tblCQ_Final da produção:
- Nenhuma linha dos 10 lotes de urina tem MATRIZ SORO. Nenhuma tem ANALITO_CADASTRADO = SIM, ou seja, nenhuma está
  num analito sérico do Painel.
- Nenhuma linha fora da família 635x/565x tem MATRIZ URINA ou sufixo 03/04. A regra não "puxou" soro para urina.
- 5659: 28 participantes saíram de Glicose (soro) e foram para GLUC/URINA. 6353: saíram 30 (Cloro 4, Creatinina 7,
  Cálcio 2, Fósforo 7, Potássio 4, Sódio 4, Ureia 2) e o mesmo número entrou nos códigos de urina. 6354: saíram 31.
  Total de 89, como na fumaça do ADR (5659 = 28; 6353 = 30; 6354 = 31).

### 1.2 PCR no lote 8353 (SO_ANALITOS): correta

- No relat209, o 8353 tem 975 resultados, quase todos RCRP. A família 835x é de PCR: 8351, 8352 e 8353 só têm RCRP.
  Há um único painel completo (30 analitos), no dia 24/08/2026, N1, Dimension 1, e mais um analito no dia 02/07/2026.
- Na produção ficam 965 ATIVO + 1 IMPLAUSIVEL em Proteína C-reativa. Saem 30 resultados como LOTE_OUTRO_MATERIAL: 17
  participavam e 13 eram ≤ 0. A PCR do 8353 **não perdeu nenhum ponto**: passou de 928 para 965, com os resultados
  novos.
- Do ponto de vista de CQ, um painel bioquímico rodado num controle de PCR não tem valor-alvo nem DP válido. Na
  estatística por lote, ele criaria um "lote" sérico com n = 1, sem sentido metrológico. Excluir é o certo.

### 1.3 Lotes digitados errado (FORA_DO_CQ): correta, e a escolha de não remapear é defensável

- 8177 (29 resultados, 16/09/2026, N2, Dimension 2) e 8794 (30, 01/12/2025, N1, Dimension 2) são painéis completos
  únicos, com número a um dígito do lote vigente. Os outros quatro lotes da lista (7974, 8975, 8773, 8970) estão no
  relat209, mas antes da janela da base (que começa em 01/10/2025). Hoje não têm linha na base e a regra fica
  preventiva.
- Não remapear para o lote "certo" é a decisão correta. Reatribuir um resultado de CQ a outro lote por inferência
  altera o registro primário sem evidência documental e quebra a rastreabilidade (ISO 15189:2022, 7.11 e 8.4,
  controle de registros). O resultado fica guardado, visível no A13 e corrigível na origem.

### 1.4 Reconciliação: nenhum resultado sérico legítimo saiu da estatística

O cálculo foi independente, sobre o `Bioquimica.json`:
- A soma dos deltas de participantes nas chaves que mudaram é **+1.136**, igual ao delta de ATIVO (34.187 → 35.323).
- 1.212 resultados novos − 76 que saíram (17 do 8353 e 59 do 8177/8794) = 1.136. Todos os novos entraram ATIVO e
  nenhum outro resultado deixou de participar.
- Há 88 chaves lote|analito com perda. **Todas** são de lotes com regra (8353, 8177, 8794, 5659, 6353, 6354) e de
  analito sérico. Não há perda em 8974, 8977, 8352, 858x/859x, 404x nem 100310x.
- Não houve ganho em analito sérico de lote de urina.
- STATUS antes/depois: IMPLAUSIVEL caiu de 56 para 43 (os 13 do 8353 viraram LOTE_OUTRO_MATERIAL, que vem antes pela
  precedência). LOTE_FORA_DO_CQ = 59 e LOTE_OUTRO_MATERIAL = 30. Os 89 estão com PARTICIPA = NÃO, LJ = NÃO e
  NAO_PLOTAR, com o motivo fixo do ADR. O QA dá **A13 = 89** e **A14 = 0**.
- A pendência do ADR (inativação com "Creatinina" nos lotes de urina virar E10) **não se materializou**. Os 11
  INATIVADO são todos do 8977 sérico (Proteína total, Lactato, Glicose). Os 4 E01 do QA são inativações de Lactato do
  8977 sem motivo (13/10/2026), sem relação com o ADR-091.
- Hematologia: STATUS idêntico antes e depois, sem nenhum LOTE_*. A única mudança é o lote 625311 (de 4 para 123), que
  é só dado novo (3.689 resultados novos = delta de ATIVO).

### 1.5 Impacto no CQ

Mudam no 1º ATUALIZAR, como o ADR previa: média, DP, CV, Sigma e Westgard de **Creatinina, Potássio, Sódio, Cloro,
Fósforo, Cálcio, Ureia e Glicose** nos recortes que continham os resultados de urina (período ou "todos os lotes"),
mais os analitos séricos que tinham o ponto único dos 8353/8177/8794. Os lotes vigentes 8974/8977 não perderam
nenhum ponto. A mudança vai no sentido correto (o DP e o CV caem, o Sigma sobe), porque o que saiu são outliers de
outro material. Nenhum veredito operacional piora por causa da regra.

## 2. Riscos residuais e ações

| # | Risco | Prioridade | Quem age | Precisa de autorização? |
|---|---|---|---|---|
| R1 | **Lote novo de urina não cadastrado** na tblLotesMaterial: volta ao sufixo e, com 01/02, contamina o soro em silêncio (é o mesmo defeito Crítico que o ADR corrigiu). Hoje não existe alerta para "lote sem regra com analito típico de urina (MALB, UCFP) ou com resultados ora 01/02, ora 03/04" | **Alto** | Laboratório: cadastrar o lote de urina **antes** do 1º uso (fazer parte da troca de lote). Projeto: novo alerta no QA (heurística acima) e texto no GUIA (pendência do ADR) | Sim, para o alerta (regra nova de QA) |
| R2 | **Sufixo -EE** continua digitado: a mesma falha pode atingir equipamento (01 vs 02 = Dimension 1/2) e não só a matriz. Exemplo: o 6354 tem resultados com sufixo 01 **e** 02 no mesmo material | Médio | Laboratório: rever a digitação do idAmostra (código de barras, ou lista fechada no equipamento) | Não (é ação de processo) |
| R3 | **Lotes digitados errado que a lista não cobre**, de PCR e HbA1c, com ATIVO e analito cadastrado: 8335 (3 PCR, inclusive 1 novo, de 05/11), 8356 (3), 8553 (2), 8383 (2), 5352 (1); HbA1c 8989 (4), 8359 (2), 8566 (1), 8990 (1). São 19 resultados. Não contaminam a estatística dos lotes vigentes (a estatística é por lote), mas: (a) somem do lote verdadeiro (faltam corridas no LJ e no Westgard daquele dia); (b) entram no recorte "todos os lotes"; (c) podem criar lote fantasma. O relat209 tem ainda o 87974 (painel completo de 17/09/2025, hoje fora da janela) | Médio | Laboratório: confirmar cada um e corrigir na origem, ou pedir a inclusão como FORA_DO_CQ (a tabela é editável e não exige código) | Não para a edição da tabela pelo laboratório. Sim, se o projeto quiser alertar automaticamente |
| R4 | **CQ de urina sem estatística**: CRE2, PHOS, MALB, UCFP, NA, K, CL, BUN, CA e GLUC de urina ficam "sem cadastro". O 6354 tem 859 participantes e o 5659 tem 154, mas nenhum Painel, Westgard ou Sigma. BUN, CA, CL, K, NA e GLUC de URINA nem estão no de/para. É uma lacuna de escopo, não defeito do ADR-091 | Médio | Laboratório/RT: decidir se o CQ de urina passa a ser monitorado no QC_INI (cadastro, ETp de urina) | Sim (escopo novo) |
| R5 | **Corridas sem CQ avaliado**: se o 8177 era o 8977 (16/09/2026, N2, D2) e o 8794 era o 8974 (01/12/2025, N1, D2), o QC_INI não avaliou essas corridas | Baixo | Laboratório: conferir no registro do dia se a corrida foi liberada com CQ aceito | Não |

Contexto, sem ocorrência nova: há cerca de 327 resultados ATIVO dos lotes séricos (8974: 223; 8977: 89) com valor
mais perto da mediana do **outro nível**. É troca de nível (o dígito N do idAmostra), já registrada na OC-0004 como
ação do laboratório. Esse ponto não tem relação com o ADR-091. Ele é citado porque tem a mesma causa raiz (o campo é
digitado) e reforça o R2.

## 3. Critérios de aceite (para o reteste do Tester e do Usuário Simulado)

1. Em todas as linhas da tblCQ_Final cujo LOTE está numa regra MATRIZ, MATRIZ = URINA e ANALITO_CADASTRADO = NÃO. A
   contagem de exceções é **0**, exata.
2. Fora dos lotes da regra MATRIZ, MATRIZ = URINA e sufixo 03/04 aparecem em **0** linhas.
3. Lote 8353: (STATUS = LOTE_OUTRO_MATERIAL) ⇔ (ANALITO ≠ Proteína C-reativa) e (STATUS ∉ {INATIVADO e anteriores}).
   Na produção de 11/11 são 30 linhas. PCR: 965 ATIVO + 1 IMPLAUSIVEL.
4. Lotes FORA_DO_CQ: todas as linhas com STATUS = LOTE_FORA_DO_CQ (59 em 11/11), PARTICIPA = NÃO e
   TIPO_PLOTAGEM_LJ = NAO_PLOTAR.
5. A13 = número de linhas LOTE_* (89) e A14 = 0.
6. Reconciliação do ATUALIZAR: Δ ATIVO = novos − (LOTE_* que eram ATIVO). Toda chave lote|analito com perda de
   participantes pertence a lote com regra. Tolerância **0** (contagens inteiras).
7. Estatística (Usuário Simulado): em Creatinina, Potássio, Sódio, Cloro, Fósforo e Glicose de soro, nenhum ponto do
   recorte por período passa de 3× a mediana do lote/nível. Hoje esse ponto existiria só se houvesse urina. Média, DP
   e CV recalculados em Python sobre os participantes conferem com a célula, com tolerância de 1e-9 relativo
   (REGRAS_COMUNS §3.5).

## 4. Conclusão

**ADR-091 VALIDADO do ponto de vista técnico de CQ.**
- A regra que usa o lote como discriminador do material é correta e necessária.
- A implementação na produção faz exatamente o que o ADR manda.
- Nenhum resultado sérico legítimo saiu da estatística e nenhum resultado de urina ficou num analito sérico.

A validação **operacional** das OC continua com o Tester e o Usuário Simulado (REGRAS_COMUNS §3.3).

Ações do laboratório:
- R1: cadastrar todo lote novo de urina antes do 1º uso.
- R2: corrigir a digitação do idAmostra.
- R3: confirmar os 9 lotes residuais de PCR/HbA1c.
- R4: decidir o escopo do CQ de urina.

Ações do projeto, que dependem de autorização explícita do usuário:
- o alerta do R1;
- a descrição da tabela no GUIA (pendência já prevista no ADR).
