# Bioquímica, F1 (entrada de CIQ manual e automática): ciclo 2026-10-31_ciclo01

Agente: Usuário Simulado. Modo: completo. Data: 01/11/2026.

Este relatório é o do **3º passe** da frente, das 01h44 às 02h04 de 01/11. Os passes anteriores usaram a mesma produção
(sha `fe10dba1f1fd`): o 1º das 16h06 às 16h37 de 31/10 e o 2º das 20h42 às 21h00. Neste passe:

- refiz as sessões A (automático, manual e validações), B (equivalência) e C (SUBSTITUI_ID) em cópias novas, com um
  DB_SEAC congelado novo (01/11 01:44);
- acrescentei a sessão **E**: equipamento e analito digitados em **minúsculas** na Bioquímica. É a pendência que a F1
  da Hematologia deixou na OC-HF1-1 ("Vale conferir se a Bioquímica tem o mesmo comportamento"). **Tem.** Ver a **OC-9**.

Todos os números das sessões A, B e C se repetiram bit a bit em relação ao 2º passe. A comparação automática campo a
campo dos JSON achou diferença só na data do V05 (amanhã = 02/11) e na contagem de linhas da origem (ver A01). As
ocorrências OC-1 a OC-8 foram **reproduzidas**, e a OC-9 é nova.

## 1. Ambiente e evidência

| Item | Valor |
|---|---|
| Produção copiada | `QC_Bioquimica.xlsm`, sha256 curto `fe10dba1f1fd`, última gravação 30/10/2026 16:33. O hash é o mesmo antes e depois: a produção não foi aberta |
| Cópias (uma por sessão, todas novas) | `%TEMP%\qcw\auditoria\2026-10-31_ciclo01\f1r3\Bioquimica_F1_{A,B,C,E}.xlsm`. Todas com sha `fe10dba1f1fd` antes e depois, ou seja, nada foi salvo |
| Excel | PIDs 24804 (A), 2824 (B), 26368 (C) e 23896 (E), um de cada vez, todos encerrados no `finally`. Nenhum outro processo foi tocado |
| Origem automática | Cópia de arquivo do DB_SEAC de produção (`<UNC>`, só leitura), feita em 01/11 01:44: `DB_SEAC_congelado.xlsm`, sha `2c9c0d9f307d`. A Cfg_Integracao das cópias foi apontada para ela, como o administrador faria |
| Configuração da cópia | MODO_FONTE=SEAC, MOTOR_DADOS=VBA, PREFIXO_ID=BIO, EQUIPAMENTO_PADRAO=DIMENSION 1, MATRIZ_PADRAO=SORO, GAP_CORRIDA_MIN=10, TOLERANCIA_CONFLITO_MIN=30, NIVEIS=2 |
| Ferramenta | `_arquitetura/agentes/ferramentas/f1_entrada_ciq.py`, sessões A, B, C e **E** (criada neste passe, documentada no cabeçalho). Saídas `f1_{A,B,C,E}.json` e `.log` em `%TEMP%\qcw\auditoria\2026-10-31_ciclo01\f1r3\`. Leituras extras sem Excel (openpyxl na cópia A): `sumidos.json` (A01) e `naopos.json` (OC-1) |
| Dados de teste | MOTIVO `Outro: TESTE-AUDITORIA` e valores terminados em ,111 ou ,222: N1 = 77,111 e N2 = 244,222; validações 90,111; substituição 270,111 e 270,222. Na cópia B, IDs sintéticos 9000001 e 9000002, com ID_AMOSTRA `TESTE-AUDITORIA` |

**Como os passos foram executados**

- **Digitação:** escrevi na tabela `tblResultados_Manuais` célula a célula, com os eventos ligados, e o
  `Worksheet_Change` carimbou cada linha. Para simular a digitação no Excel pt-BR, usei `FormulaLocal`.
- **ATUALIZAR:** rodou por `mIntegracao.AtualizarDadosAutomatico`, o núcleo do botão ATUALIZAR DADOS sem a caixa de
  diálogo.
- **Filtros do Painel (H3, L4, C3 e K3):** mudei pela própria célula, com os eventos ligados. Para isso tirei a
  proteção da aba Painel, um passo técnico.

**Limites do método**

1. A escrita por COM não passa pela validação de dados do Excel. As caixas de erro da validação, como "Data
   inválida." e "Resultado numérico.", **não foram vistas**: só li as regras. Os casos V06, V08, V11, V13 e E02/E03
   representam o usuário que **cola** um valor. A validação de lista do Excel não diferencia maiúsculas de minúsculas,
   então "dimension 1" e "glicose" provavelmente passam **mesmo digitados**. Isso não foi exercitado.
2. O aviso de "ID inexistente" do SUBSTITUI_ID é uma caixa de mensagem, que não aparece em automação
   (`Interactive = False`). Conferi só o efeito: nada foi preenchido.
3. A cópia foi aberta pelo `qa_final.QA`, que revela as abas sem login. Por isso, USUARIO e USUARIO_MANUAL ficam
   "(sem login)". A identidade do usuário logado (ADR-084) **não** foi exercitada nesta frente.

## 2. Testes executados

| ID | Teste | Estado | Evidência |
|---|---|---|---|
| T01 | Tela Digitar Resultados: campos, validação, proteção e instruções | APROVADO (com ressalvas, ver OC-3 e OC-6) | 17 colunas: ID_REGISTRO, DATA, HORA, EQUIPAMENTO, MATRIZ, LOTE, NIVEL, ANALITO, RESULTADO, UNIDADE, MOTIVO, USUARIO, REGISTRADO_EM, RUN, STATUS, DETALHE, SUBSTITUI_ID. RUN, STATUS e DETALHE são `XLOOKUP` na tblCQ_Final. A aba é protegida. Validações: DATA > 01/01/2020 ("Data inválida."), sem limite superior; EQUIPAMENTO "DIMENSION 1,DIMENSION 2"; MATRIZ "SORO,URINA"; LOTE `=lstLotes`; ANALITO `=lstAnalitos`; NIVEL inteiro de 1 a 2; RESULTADO decimal ("Resultado numérico."); MOTIVO em lista. HORA e SUBSTITUI_ID não têm validação. As instruções A1:A4 não citam EQUIPAMENTO |
| A01 | Recebimento DB_SEAC → `tblDB_Recebimento` | APROVADO | A origem tem agora 31.386 linhas (eram 31.577 às 20h42 de 31/10), com 31.386 IDs válidos, nenhum sem ID e nenhum repetido. A janela da origem passou a começar em 03/11/2025 08:10: a virada do mês tirou o resto de outubro/2025. Das 31.055 linhas já recebidas pela produção nessa janela, **0** faltam na origem atual (`sumidos.json`). O QC guarda o histórico que saiu da janela. Recebimento: 34.254 → 34.585 (+331). 0 IDs não recebidos e 0 duplicados. Novos: BIO-371188 a BIO-372097, os mesmos do 2º passe |
| A02 | Preservação do ID e associação: origem × recebimento × tblCQ_Final | APROVADO | 0 divergências nas colunas origem × recebimento. 0 divergências nos 10 campos origem × final dos 331 IDs: EQUIPAMENTO, LOTE, NIVEL, ANALITO→ANALITO_ORIGEM, DATA, DATA_HORA, VALOR→RESULTADO, MATRIZ, ID_ORIGEM e ID_AMOSTRA. ID_REGISTRO = `BIO-<ID_ORIGEM>`. Os 331 entram ATIVO, com ORIGEM=INTERFACEAMENTO. O de-para está correto (GLUC→Glicose, A1C→Hemoglobina glicada, TCO2→Bicarbonato...). As urinas e FERR, TNIH, UCFP e MALB ficam ANALITO_CADASTRADO=NÃO, com I01 (por projeto) |
| A03 | Sync_Log e QA | APROVADO | Sync_Log: OK, VBA, SEAC, LIDOS_ORIGEM 31.386 (= origem), SEM_ID 0, NOVOS 331, TOTAL_RECEBIDO = LINHAS_CQ = 34.585, DATA_MAX 30/10/2026 07:31:08 (= máximo da origem), 44,2 s. QA: 4 E01 (preexistentes, OC-5) e 10 I01. Audit_Log: 95 linhas novas (Westgard das corridas novas) |
| A04 | Histórico intacto | APROVADO | 0 linhas antigas alteradas no recebimento. 0 IDs antigos com RUN ou STATUS mudado. Final com 34.585 linhas e 34.585 IDs únicos |
| A05 | Idempotência (2º ATUALIZAR com o mesmo arquivo) | APROVADO | Sync_Log OK_SEM_MUDANCA, NOVOS 0, GRAVADAS 0. Final idêntica (ID, RUN, STATUS, RESULTADO). 15,2 s |
| A06 | Plausibilidade do que entra pelo automático | REPROVADO | Glicose N2, lote 8977, DIMENSION 1: BIO-371183 = **−1** e BIO-371184 = **1** (28/10 13:59:55) continuam ATIVO na estatística de outubro. Na produção há **56** resultados ATIVO com valor ≤ 0 (43 negativos e 13 zeros), 55 deles em analitos cadastrados. Ver **OC-1** |
| M01 | Manual válido N1 e N2 (Glicose, lote 8977, DIMENSION 1, 30/10 08:31 e 08:32): ID, carimbo, gravação e consulta | APROVADO | MAN_0001 e MAN_0002 nascem ao digitar a 1ª célula, com REGISTRADO_EM, EQUIPAMENTO (do Painel) e MATRIZ (SORO). Antes do ATUALIZAR: STATUS "—", DETALHE "atualize os dados". Depois (24,5 s): ATIVO, RUN 26103005, CORRIDA_NO_DIA 5. Na tblCQ_Final: ORIGEM=MANUAL, PARTICIPA=SIM, NORMAL, MOTIVO_MANUAL preservado. `DadosDoId` devolve ID, analito, nível, equipamento, lote, RUN, data/hora, status e valor. Audit_Log: RESULTADO_MANUAL_INCLUIDO para os dois. QA: A06 (há interfaceamento no mesmo dia: BIO-372036/372037 e BIO-372072/372073) |
| M02 | Média, DP e CV atualizados (Painel `engPainel` e Estatística) × cálculo independente | APROVADO | Ver a seção 3. O n subiu exatamente 1 por nível. A maior diferença relativa foi de 1,3e-13, abaixo de 1e-9 |
| M03 | Levey-Jennings e Westgard atualizados | APROVADO | Eng_Saida, slot 109, RUN 26103005: N1=77,111, N2=244,222, N1 ID = MAN_0001. N1 R-4s; N2 1-3s, 2-2s, R-4s e alerta 1-2s → REJEITADO. Painel: "R_4s · RUN 26103005" e "2_2s · RUN 26103005". Audit_Log: WESTGARD_EVENTO_CORRIDA_NOVA |
| V01 | Campos vazios | APROVADO | MANUAL_INCOMPLETO, "campos obrigatorios ausentes/invalidos: DATA, HORA, LOTE, NIVEL", E04 |
| V02 | Resultado "abc" | APROVADO | MANUAL_INCOMPLETO "...: RESULTADO", E04 |
| V03 | Resultado "90.5" (ponto) | APROVADO | Fica texto. MANUAL_INCOMPLETO "RESULTADO ambiguo (use virgula decimal: 1,5)", E04 |
| V04 | Resultado "1.5" (risco de virar data) | APROVADO | Não virou data. MANUAL_INCOMPLETO "ambiguo", E04 |
| V04b | Resultado "77,111" (vírgula) | APROVADO | Gravado como número 77,111, exibido como "77,111" |
| V05 | Data futura (02/11/2026) | APROVADO | A tela aceita (só exige > 01/01/2020). O ATUALIZAR barra: MANUAL_INCOMPLETO "DATA/HORA posterior ao lancamento", E04 |
| V06 | Analito inexistente "Glicoze" (colado) | REPROVADO | **ATIVO**, PARTICIPA=SIM, ANALITO_CADASTRADO=NÃO. Não aparece no Painel. Só um I01 agregado. Ver **OC-4** |
| V07 | Nível 3 | APROVADO | MANUAL_INCOMPLETO "NIVEL", E04 |
| V08 | Lote inexistente "99999" (colado) | REPROVADO | **ATIVO**, PARTICIPA=SIM, sem nenhuma mensagem no QA. Ver **OC-4** |
| V09 | Duplicidade manual × manual | APROVADO | MAN_0011 fica CONFLITO_MANUAL "resultado manual repetido (mesma chave e instante)", ID_RELACIONADO MAN_0001, A01. O n do Painel não muda |
| V10 | Duplicidade manual × interfaceamento (+5 min, mesmo valor) | APROVADO | MAN_0012 fica CONFLITO_MANUAL "coincide com resultado do interfaceamento", ID_RELACIONADO BIO-372036, A01 com a orientação do SUBSTITUI_ID |
| V11 | Equipamento inexistente "DIMENSION 9" (colado) | REPROVADO | **ATIVO**, PARTICIPA=SIM, com evento de Westgard próprio e sem nenhuma mensagem no QA. Ver **OC-4** |
| V12 | Hora "25:99" | APROVADO | MANUAL_INCOMPLETO "HORA", E04 |
| V13 | Data "31/02/2026" (colada; fica texto) | REPROVADO | Vira **03/03/2026 08:00**, **ATIVO**, RUN 26030301, sem aviso. O resumo do N1 no Painel passou de "36 ev / 24 res, maior Z −3,44" para "38 ev / 26 res, maior Z +7,00". Causa no código: `mEtlCQ.ParaData` (linhas 211–223) monta a data com `DateSerial` sem conferir dia e mês. Ver **OC-2** |
| V15 | EQUIPAMENTO vazio com o Painel em DIMENSION 2 | REPROVADO (risco) | A linha ganhou **DIMENSION 2** antes mesmo do ATUALIZAR e entrou ATIVO na série do DIMENSION 2 (MAN_0016). O padrão é `selEquipamento` (Painel), não o EQUIPAMENTO_PADRAO. Ver **OC-3** |
| E01 | Equivalência: o mesmo par N1/N2 pelo caminho manual (A) e pelo automático (B, DB_SEAC sintético, BIO-9000001/2) | APROVADO | Mesmo RUN 26103005 e mesma CORRIDA_NO_DIA 5. `engPainel` (2 × 21 colunas), as linhas da Estatística e o cálculo independente idênticos bit a bit. As 15 últimas linhas do Eng_Saida (39 colunas) diferem só na coluna "N1 ID" (MAN_0001 × BIO-9000001) |
| S01 | SUBSTITUI_ID: digitar só o número do ID (371184) numa linha livre preenche a chave | APROVADO | Antes do ATUALIZAR: MAN_0001, SUBSTITUI_ID normalizado para "BIO-371184", DATA 28/10/2026, HORA 13:59:55, DIMENSION 1, lote 8977, NIVEL 2, Glicose |
| S02 | SUBSTITUI_ID inexistente (987654321) | APROVADO | Nada preenchido: sem ID, data, hora, equipamento, lote, nível nem analito. A caixa de aviso não foi vista (limite 2) |
| S03 | Substituição com MOTIVO (270,111 no lugar do "1" de BIO-371184) | APROVADO | MAN_0001 ATIVO, PARTICIPA=SIM, RUN 26102806, corrida 6 (os do automático). BIO-371184 fica SUBSTITUIDO_MANUAL, NÃO, sem RUN, ID_RELACIONADO MAN_0001. 0 outros RUN ou corridas mudaram. QA I04 com os dois valores, a diferença e como desfazer. A06 lista BIO-371038/371039/371183 como "continuam valendo". Dica: `MAN_0001 ... 26102806 ... ATIVO 270,111` |
| S04 | SUBSTITUI_ID sem MOTIVO (BIO-371183, 270,222) | APROVADO | MAN_0002 MANUAL_INCOMPLETO "MOTIVO (obrigatorio para substituir um resultado)", E04. BIO-371183 continua ATIVO, RUN 26102805, valor −1 |
| S05 | Estatística depois da substituição × cálculo independente | APROVADO | Ver a seção 3. N2: n 24, média 259,129625, CV 21,49%. Diferença relativa máxima 3,3e-15 |
| **L01** | **Equipamento colado/digitado em minúsculas "dimension 1"** (manual N2 válido, 29/10 23:36, 244,222) | **REPROVADO** | MAN_0001 ATIVO, sem nenhuma mensagem no QA. Ganhou o **RUN 26102901**, o da 1ª corrida real do dia. No Painel e no LJ, o manual **substituiu** o N2 real BIO-371977 (277) desse RUN. A Estatística contou os dois. Os vereditos de Westgard do RUN real passaram de OK para REJEITADO. Ver seção 3 e **OC-9** |
| **L02** | **Analito colado/digitado em minúsculas "glicose"** (equipamento certo, 28/10 23:36) | **REPROVADO** | MAN_0002 ATIVO, ANALITO_CADASTRADO=NÃO, sem nenhuma mensagem no QA para o ID. Ganhou o **RUN 26102801** e substituiu no Painel o N2 real BIO-371038 (276) desse RUN. A Estatística ficou com n 26. Ver **OC-9** |

## 3. Estatística: Painel × Estatística × cálculo independente

Filtro: Glicose, lote 8977, DIMENSION 1, de 01/10/2026 a 31/10/2026 (Painel G3:G4, igual a Estat_Ini/Fim_Efetiva).

- **População independente:** linhas da tblCQ_Final com PARTICIPA=SIM nesse filtro.
- **Cálculo:** média por `statistics.fmean`, DP amostral (n−1) por `statistics.stdev`, CV = DP/média × 100.
- **Tolerância:** 1e-9 relativo (REGRAS_COMUNS §3.5).

### Depois do manual (M02, cópia A; idêntico na cópia B)

| Nível | Fonte | n | Média | DP | CV % |
|---|---|---|---|---|---|
| N1 antes | engPainel = independente | 25 | 77,36 | 1,1135528725660044 | 1,4394426998009364 |
| N1 depois | engPainel | 26 | 77,35042307692308 | 1,092146809658003 | 1,4119467822068537 |
| N1 depois | Estatística | 26 | 77,35042307692308 | 1,0921468096578573 | 1,4119467822066654 |
| N1 depois | independente | 26 | 77,35042307692308 | 1,0921468096580027 | 1,4119467822068532 |
| N2 antes | engPainel = independente | 24 | 247,91666666666666 | 76,55740481419193 | 30,880297740178257 |
| N2 depois | engPainel | 25 | 247,76888 | 74,94913215881823 | 30,249614947130663 |
| N2 depois | Estatística | 25 | 247,76888 | 74,94913215881826 | 30,249614947130674 |
| N2 depois | independente | 25 | 247,76888 | 74,94913215881823 | 30,24961494713066 |

### Depois da substituição de BIO-371184 (S05, cópia C)

| Nível | Fonte | n | Média | DP | CV % |
|---|---|---|---|---|---|
| N2 depois | engPainel | 24 | 259,129625 | 55,68195830577715 | 21,48807119439823 |
| N2 depois | Estatística | 24 | 259,129625 | 55,68195830577723 | 21,488071194398263 |
| N2 depois | independente | 24 | 259,129625 | 55,681958305777144 | 21,488071194398227 |

Conferência manual: soma antes = 24 × 247,91666… = 5.950,0. Trocando 1 por 270,111, a soma vai a 6.219,111, e
6.219,111 / 24 = 259,129625.

**Status: APROVADO** nas duas tabelas. A diferença entre as fontes está na 13ª a 15ª casa significativa, efeito da
ordem de soma em ponto flutuante.

### Minúsculas (L01/L02, cópia E): o caso reprovado

N2 de Glicose, lote 8977, DIMENSION 1, outubro. A linha de base (E1) é a mesma do "N2 antes" acima.

| Passo | Fonte | O que o QC_INI fez | n | Média | DP | CV % |
|---|---|---|---|---|---|---|
| L01 ("dimension 1") | engPainel | trocou o 277 de BIO-371977 (RUN 26102901) pelo 244,222 manual | 24 | 246,55091666666667 | 76,30797925792253 | 30,950190852905987 |
| L01 | independente, "um valor por RUN" (244,222 no lugar do 277) | reproduz o Painel | 24 | 246,55091666666667 | 76,30797925792254 | 30,95019085290599 |
| L01 | Estatística | contou os dois | 25 | 247,76888 | 74,94913215881826 | 30,249614947130674 |
| L01 | independente, equipamento e analito sem diferenciar maiúsculas | reproduz a Estatística | 25 | 247,76888 | 74,94913215881823 | 30,24961494713066 |
| L01 | independente, equipamento exato ("DIMENSION 1") | o que o ETL considera | 24 | 247,91666666666666 | 76,55740481419193 | 30,880297740178257 |
| L02 ("glicose") | engPainel | trocou também o 276 de BIO-371038 (RUN 26102801) | 24 | 245,2268333333333 | 76,0500337842115 | 31,012117536435248 |
| L02 | independente, "um valor por RUN" (as duas trocas) | reproduz o Painel | 24 | 245,22683333333336 | 76,0500337842115 | 31,01211753643524 |
| L02 | Estatística | contou os três | 26 | 247,6324615384615 | 73,4381465916623 | 29,65610652796266 |
| L02 | independente, sem diferenciar maiúsculas | reproduz a Estatística | 26 | 247,63246153846154 | 73,43814659166216 | 29,656106527962596 |

Eng_Saida, slot 103, RUN 26102901:

- antes: N1 78, N2 277, Data 29/10 14:59:24, N1 ID BIO-371928, vereditos N1 OK e N2 OK;
- depois do L01: N2 **244,222**, Data **29/10 23:36**, vereditos N1 **REJEITADO** e N2 **REJEITADO**.

No slot 97 (RUN 26102801), depois do L02, o N2 ficou 244,222, a data 28/10 23:36 e os dois vereditos REJEITADO.

O ET da linha de Sigma da Estatística também andou com os resultados contados a mais: 53,69 → 52,65 → 51,67.

**Status: REPROVADO.** O mesmo resultado entra de três jeitos: no RUN do ETL, como corrida própria; no Painel e no LJ,
como substituto do real; na Estatística, como resultado a mais. As reconstruções independentes batem com o QC_INI em
todas as casas. O defeito, portanto, não é de aritmética: é de **associação** (qual resultado pertence a qual série e
corrida).

## 4. Ocorrências

As OC-1 a OC-8 foram reproduzidas neste passe com os mesmos IDs e valores. A OC-9 é nova.

### OC-9 (nova): Equipamento ou analito em minúsculas faz o manual "tomar o lugar" de um resultado real no Painel, no LJ e no Westgard, e a Estatística o conta a mais (Crítico; tipo `calculo`)

É o mesmo defeito da OC-HF1-1 da Hematologia, agora confirmado na Bioquímica.

- **Tela/módulo:** Digitar Resultados → ATUALIZAR.
  - `mEtlCQ.CalcularCorridas` (`src/comum/mEtlCQ.bas`, linha 1031 em diante, `Option Compare Binary` na linha 21)
    monta a chave da corrida com o texto exato de EQUIPAMENTO e ANALITO.
  - O motor da Bio (`src/bio/mEstatistica.bas`, `EquipFiltro` na linha 122 e `GarantirIndice` nas linhas 210 e
    213) compara equipamento e analito em **maiúsculas**. O Eng_Saida guarda um valor por RUN e nível.
- **Ação:** colar (ou, provavelmente, digitar, ver limite 1) "dimension 1" no EQUIPAMENTO (L01) ou "glicose" no
  ANALITO (L02) de um manual N2 válido. Depois, preencher o resto e clicar em ATUALIZAR.
- **Esperado:** uma destas três saídas:
  - o texto é normalizado (DIMENSION 1, Glicose) e o manual entra na corrida certa como resultado próprio;
  - o manual fica MANUAL_INCOMPLETO;
  - o manual sai com um ALERTA.

  Em nenhum caso um resultado real pode sumir do gráfico, nem o Painel e a Estatística podem divergir.
- **Observado:**
  - O manual de 29/10 23:36 recebeu o RUN **26102901**, o da 1ª corrida real do dia, porque para o ETL "dimension 1"
    é outra série e ali ele é a corrida 1 (CORRIDA_NO_DIA 1).
  - O motor junta "dimension 1" ao DIMENSION 1. No RUN 26102901, o N2 manual (244,222) **substitui** o N2 real
    BIO-371977 (277):
    - o LJ mostra 244,222 com data 23:36 no lugar do ponto real;
    - o Painel recalcula com a troca (n 24, média 246,5509);
    - os vereditos de Westgard dessa corrida real viram REJEITADO.
  - A Estatística conta os dois (n 25, média 247,7689), e o ET da linha de Sigma muda.
  - Nenhum alerta: `qa` vazio para o ID. A dica do ponto mostra "dimension 1", e o STATUS é ATIVO.
  - No L02, "glicose" fica ANALITO_CADASTRADO=NÃO, mas mesmo assim entra na Glicose do Painel, onde troca o 276 de
    BIO-371038, e na da Estatística (n 26). Não sai nenhum I01 para o ID.
- **Dificuldade:**
  - O STATUS diz ATIVO, e o usuário acha que deu certo.
  - A validação de lista do Excel aceita a grafia em minúsculas.
  - O valor real some do gráfico sem nenhum registro.
- **Consequência:**
  - estatística incorreta e divergente entre o Painel e a Estatística (n diferente para o mesmo filtro);
  - LJ com um ponto trocado e data errada;
  - Westgard avaliado sobre o valor errado numa corrida real, que é rejeitada;
  - o resultado do interfaceamento continua ATIVO na base, mas invisível no gráfico.

  É associação errada entre resultado e corrida e estatística incorreta, por isso Crítico (REGRAS_COMUNS §4).
- **Sugestão (decisão do Orquestrador, do Especialista em CQ e do usuário, não do agente):**
  - normalizar EQUIPAMENTO e ANALITO do manual para a grafia do cadastro no `PrepararLinha` ou no `mEtlCQ`, nos dois
    motores (VBA e PQ);
  - ou usar a mesma comparação no ETL e no motor;
  - e, no motor, nunca sobrescrever um valor já presente no mesmo RUN e nível sem alerta.

  Como o código é comum (`mEtlCQ`) e os motores dos dois produtos usam `UCase`, uma correção única deve cobrir a
  OC-HF1-1 e a OC-9.
- **Evidência:** `f1r3\f1_E.json`:
  - E1, E2 e E3: `painel`, `estat`, `ind_exato`, `ind_maiusc`, `pop_maiusc`, `eng_run`, `final`, `qa` = [], `dica`;
  - `E_dias`: d1 29/10 com RUNs 26102901/02; d2 28/10 com RUNs 26102801/02/05/06.

  A reconstrução está na seção 3.

### OC-1: Valores de controle impossíveis (−1, 1, zero, negativos) do interfaceamento entram ATIVO na estatística e no Westgard (Crítico; tipo `dados`)

- **Tela/módulo:** ATUALIZAR DADOS → tblCQ_Final (mEtlCQ) → Painel e Estatística.
- **Ação:** ATUALIZAR com o DB_SEAC de produção (cópia congelada em 01/11 01:44) e conferir Glicose, lote 8977,
  DIMENSION 1, outubro.
- **Esperado:** um valor fisicamente impossível (glicose −1 ou 1 mg/dL num controle de cerca de 270) não entra calado
  na média, no DP, no CV, no ET, no Sigma e no Westgard. No mínimo, o QA dá um alerta com o ID.
- **Observado:**
  - BIO-371183 (VALOR −1) e BIO-371184 (VALOR 1), N2, 28/10/2026 13:59:55, sem FLAG na origem. Ficam ATIVO e
    PARTICIPA=SIM, e já estão na produção.
  - Para o N2, Painel e Estatística mostram n 24, DP 76,557, CV 30,88%, Westgard REJEITADO ("maior Z −54,86"),
    ET 53,69%, "ETp excedido" e "Desempenho inadequado".
  - Sem os dois valores: DP 5,78 e CV 2,14% (1º passe).
  - **Novo neste passe** (`naopos.json`, tblCQ_Final da produção): **56** resultados ATIVO com valor ≤ 0, sendo 43
    negativos e 13 zeros. 55 deles são de analitos cadastrados: Creatinina 6; Colesterol total, CTLF, Lipase e Ureia
    5 cada; CK, Albumina e Lactato 4 cada; e outros 13 analitos. Todos são de 12/2025 em diante, com picos em
    03/2026 (20) e 08/2026 (13). Exemplo: BIO-224545/224547/224549/224551, Colesterol total N2, DIMENSION 2, lote
    8974, 05/12/2025, valores −8, −8, −10 e −8, em corridas seguidas.
- **Dificuldade:** o usuário só percebe pelo CV absurdo ou pelo Z extremo. Nada no QA aponta o valor.
- **Consequência:** estatística, ET, Sigma e veredito de Westgard incorretos para esses analitos, lotes e níveis
  enquanto ninguém inativa. Risco de decisão errada de liberação e de plano de CQ.
- **Remédios que já existem (S03):** inativar o valor ou substituí-lo por um manual (SUBSTITUI_ID). Os dois
  funcionam, mas dependem de alguém notar o valor.
- **Sugestão (decisão do Especialista em CQ e do RT, não do agente):**
  1. uma triagem de plausibilidade no QA (VALOR ≤ 0, ou |z| acima de um limite em relação ao alvo do lote) que gere
     ALERTA com o ID, sem excluir sozinha;
  2. o laboratório revisar os 56 IDs;
  3. perguntar ao SIPEC o que significam −1, 0 e 1 sem flag.
- **Evidência:** `f1r3\f1_A.json` (A1_plausibilidade, A3_painel_antes, A3_estat_antes, A3_ind_antes),
  `f1r3\f1_C.json` (C_alvos_antes) e `f1r3\naopos.json`.

### OC-2: Data inválida colada ("31/02/2026") vira 03/03/2026 e o manual entra ATIVO no histórico do lote (Alto; tipo `dados`)

- **Tela/módulo:** Digitar Resultados → ATUALIZAR (`mEtlCQ.ParaData` / `ParaDataHora`, linhas 211–243).
- **Ação:** colar "31/02/2026" na DATA (a célula guarda texto), preencher o resto e clicar em ATUALIZAR.
- **Esperado:** MANUAL_INCOMPLETO "DATA", como acontece com a hora "25:99".
- **Observado:**
  - MAN_0015 ficou com DATA_HORA 03/03/2026 08:00, ATIVO, PARTICIPA=SIM e RUN 26030301, sem aviso.
  - O Westgard registrou uma corrida nova em março no lote 8977 / DIMENSION 1, e o resumo do N1 no Painel mudou.
  - Causa: o `DateSerial(ano, mês, dia)` não confere se o resultado tem o mesmo dia e mês, e rola 31/02 para 03/03.
- **Dificuldade:** o STATUS mostra ATIVO, e o usuário acha que deu certo.
- **Consequência:** o resultado fica numa data errada, dentro de um mês já fechado, e muda a estatística e o Westgard
  retroativos do lote.
- **Mitigação existente:** ao **digitar**, a validação "Data inválida." recusa o texto (regra lida, não exercitada).
- **Sugestão:** depois de converter, conferir dia, mês e ano. Se não baterem, MANUAL_INCOMPLETO "DATA". O mesmo no
  motor PQ de reserva.
- **Evidência:** `f1r3\f1_A.json` → A4_validacoes.casos["V13 ..."], A4_validacoes.painel e audit (AUD com chave
  `26030301|1|GLICOSE`).

### OC-3: EQUIPAMENTO vazio é preenchido com o equipamento do Painel, sem aviso, e as instruções não pedem o equipamento (Alto; tipo `funcional`)

- **Tela/módulo:** Digitar Resultados (`mIntegracao.PrepararLinha` → `EquipPadrao` = `selEquipamento`, e só se ele
  estiver vazio o EQUIPAMENTO_PADRAO).
- **Ação:** com o Painel em DIMENSION 2, digitar um manual sem EQUIPAMENTO, seguindo as instruções da aba (A3) e o
  GUIA §3.1, que não citam o equipamento.
- **Esperado:** o equipamento vem do usuário, ou o campo é obrigatório. Um padrão fixo também serviria, desde que
  avisado.
- **Observado:** a linha ganhou "DIMENSION 2" ao ser criada e entrou ATIVO na série do DIMENSION 2 (MAN_0016).
  Nenhuma mensagem. O A06 só saiu porque havia resultados do DIMENSION 2 no mesmo dia.
- **Consequência:** associação errada entre equipamentos: o resultado entra na estatística e no Westgard do
  equipamento errado.
- **Sugestão:** incluir EQUIPAMENTO nas instruções (A3 e GUIA §3) e torná-lo obrigatório na Bioquímica, ou pelo menos
  destacar o valor preenchido sozinho.
- **Evidência:** `f1r3\f1_A.json` → A4_validacoes.casos["V15 ..."]: digitado_antes_de_atualizar.EQUIPAMENTO =
  "DIMENSION 2"; cfg EQUIPAMENTO_PADRAO = DIMENSION 1.

### OC-4: Manual com analito, lote ou equipamento inexistente fica ATIVO, a tela diz "entrou na estatística" e nada aparece no Painel (Médio; tipo `funcional`)

- **Tela/módulo:** Digitar Resultados → ATUALIZAR (mEtlCQ, mEtlQA).
- **Ação:** colar "Glicoze" (V06), o lote "99999" (V08) e o equipamento "DIMENSION 9" (V11).
- **Esperado:** MANUAL_INCOMPLETO, ou no mínimo um ALERTA com o ID.
- **Observado:** os três ficam ATIVO e PARTICIPA=SIM, e nenhum aparece no Painel. Para o analito, só sai o I01
  agregado. Para o lote e o equipamento, nenhuma linha no QA. O GUIA §3 diz "ATIVO = entrou na estatística".
- **Mitigação existente:** ao digitar, as listas recusam o valor (regra lida). O risco é colar.
- **Consequência:** resultado perdido em silêncio.
- **Sugestão:** analito, lote ou equipamento fora do cadastro → MANUAL_INCOMPLETO (com o campo no DETALHE) ou ALERTA
  com o ID. A grafia em minúsculas (OC-9) é o caso perigoso deste mesmo tema, porque nela o resultado não some:
  entra no lugar errado.
- **Evidência:** `f1r3\f1_A.json` → A4_validacoes.casos V06, V08, V11 e qa_codigos.

### OC-5: Quatro inativações sem motivo na produção fazem toda atualização dizer "Há ERRO DE GOVERNANÇA" (Médio; tipo `dados`)

- **Observado:** E01 em BIO-366027, BIO-366033, BIO-366054 e BIO-366164 (Lactato N2, 13/10/2026, "inativado sem
  justificativa"). Toda atualização deste passe (A1 a A4, B1, B2, C1, C5 e E1 a E3) termina com "Há ERRO DE
  GOVERNANÇA".
- **Consequência:** o alarme permanente vira ruído, e um erro novo de verdade passa despercebido.
- **Sugestão:** o laboratório escrever o MOTIVO dessas 4 linhas na aba Inativar. Não é defeito de código.
- **Evidência:** `f1r3\f1_A.json` → A1_automatico.qa_codigos e resumo.

### OC-6: As colunas RUN e DETALHE da aba Digitar Resultados mostram "0" (Baixo; tipo `usabilidade`)

- **Observado:** o XLOOKUP de célula vazia devolve 0. Linhas ATIVO mostram DETALHE = 0, inclusive as de L01 e L02.
  Linhas MANUAL_INCOMPLETO ou CONFLITO_MANUAL mostram RUN = 0. O DETALHE vem sem acento ("obrigatorios",
  "lancamento", "ambiguo").
- **Sugestão:** a fórmula devolver "" para vazio, e acentuar o texto.
- **Evidência:** `f1r3\f1_A.json` → A3_manual.linhas_digitadas e A4_validacoes.casos[*].linha_depois;
  `f1r3\f1_E.json` → E2.linha e E3.linha.

### OC-7: O I01 lista "NA", "K", "CL", "CRE2" e "PHOS" como analito sem cadastro sem dizer que são da URINA (Baixo; tipo `usabilidade`)

- **Observado:** "CL: ... resultado(s) guardados na base; nao aparecem no Painel ate o analito ser
  cadastrado/mapeado". O Cloro de soro está cadastrado e aparece. Quem lê pode achar que o Cloro sumiu.
- **Sugestão:** incluir a matriz no texto, por exemplo "CL (URINA)".
- **Evidência:** `f1r3\f1_A.json` → A1_automatico.mapa_analito_origem_final (NA/K/CL → "NÃO" 3 cada; NA/K/CL →
  Sódio/Potássio/Cloro "SIM" 10 cada).

### OC-8: O Audit_Log grava RESULTADO_MANUAL_INCLUIDO também para linhas incompletas e em conflito (Baixo; tipo `funcional`)

- **Observado:** as linhas incompletas e em conflito da cópia A geraram RESULTADO_MANUAL_INCLUIDO com chave
  "0|…|GLICOSE" e RUN vazio (por exemplo, AUD-000514 com chave "0|0|GLICOSE"). O mesmo vale para o SUBSTITUI_ID sem
  motivo na cópia C.
- **Sugestão:** uma ação própria (RESULTADO_MANUAL_RECUSADO ou _EM_CONFLITO), ou o status no detalhe.
- **Evidência:** `f1r3\f1_A.json` → A4_validacoes.audit; `f1r3\f1_C.json` → C.audit.

## 5. Dúvidas de regra de negócio (ao Especialista em CQ, pelo Orquestrador)

1. Deve haver triagem de plausibilidade para valores do interfaceamento sem flag (−1, 0, 1, negativos)? Hoje os 56
   entram ATIVO (OC-1). A triagem seria alerta, exclusão automática ou só inativação manual?
2. Resultados em par no mesmo segundo, mesmo nível, com ITEM_ID diferente e valores sistematicamente diferentes (ex.:
   Glicose N2 277 × 267 em 22/10) contam como duas corridas (ADR-057 §6). O laboratório confirma que isso está certo?
3. Na Bioquímica, o EQUIPAMENTO do manual deve ser obrigatório (OC-3)?
4. Em 01/11 01:44 (domingo), o resultado mais recente do DB_SEAC continua sendo **30/10 07:31:08**, o mesmo desde
   31/10 16h04. A mensagem do ATUALIZAR mostra isso corretamente (GUIA §1.2). O laboratório confirma que não houve
   controle de bioquímica depois das 07:31 de sexta? Se houve, o problema está no SIPEC ou no DB_SEAC, e não no QC_INI.
5. O DB_SEAC é uma janela móvel. Na virada para novembro, ele deixou de trazer outubro/2025 (191 linhas a menos). O
   QC_INI guarda o que já recebeu: nada se perdeu, e o histórico não mudou (A01 e A04). O laboratório confirma que é
   isso que quer? Um corolário: se alguém **apagar** um resultado errado no DB_SEAC, o QC_INI continua com ele ATIVO.
   A correção tem de ser feita no QC (Inativar ou SUBSTITUI_ID), e o GUIA não diz isso.

## 6. Impressões de uso (Relatório A)

- **É fácil inserir?** Sim para o caso comum:
  - a linha nasce com ID, carimbo e matriz;
  - a vírgula decimal funciona, e o ponto é recusado com mensagem clara;
  - as listas guiam a digitação;
  - o SUBSTITUI_ID é o fluxo mais amigável da tela: digitar só o número do ID preenche o resto.
- **O usuário sabe o que preencher?** Quase. As instruções não falam do **EQUIPAMENTO**, que se preenche sozinho com
  o do Painel (OC-3).
- **Há risco de preencher errado?** Sim, e o maior apareceu neste passe:
  - **maiúsculas e minúsculas** no equipamento ou no analito (OC-9). A lista do Excel aceita "dimension 1" e
    "glicose", a linha fica ATIVO, e o QC troca um resultado real pelo manual no gráfico e no Westgard. Ao mesmo
    tempo, a Estatística conta os dois;
  - o equipamento implícito (OC-3);
  - a data futura, que a tela aceita (o ATUALIZAR barra, com mensagem clara);
  - colar valores, que contorna as listas e gera resultados "ATIVO" invisíveis (OC-4) ou com a data trocada (OC-2).
- **As mensagens de erro são claras?** Sim:
  - o DETALHE diz o campo exato;
  - o A01 ensina o SUBSTITUI_ID;
  - o I04 explica a troca e diz como desfazer;
  - o E04 do SUBSTITUI_ID sem motivo é explícito.

  Ressalvas:
  - o texto vem sem acento;
  - ATIVO aparece para um resultado que não está no Painel (OC-4) ou que está no lugar errado (OC-9);
  - a mensagem final sempre fala em "ERRO DE GOVERNANÇA" por causa de 4 linhas antigas (OC-5).
- **Dá para ver o que foi salvo?** Sim. STATUS, RUN e DETALHE aparecem na linha depois do ATUALIZAR (15 a 25 s), e a
  dica do ponto mostra ID, corrida e valor. Antes disso, a linha diz "atualize os dados".
- **Automático:** confiável na transferência. Nada perdido, nada duplicado, ID preservado, associação exata, Sync_Log
  coerente e idempotente. A janela móvel do DB_SEAC não apaga histórico. O problema está no **conteúdo** que a origem
  manda (OC-1).
- **Manual × automático:** equivalentes. O mesmo dado dá a mesma estatística, a mesma corrida e o mesmo Westgard
  (E01), desde que digitado com a grafia exata do cadastro (OC-9).

## 7. Limpeza

- As cópias `.xlsm` deste passe (sha `fe10dba1f1fd`, iguais à produção antes e depois), o DB_SEAC congelado, o DB_SEAC
  sintético e o cache da origem foram apagados.
- Ficam em `%TEMP%\qcw\auditoria\2026-10-31_ciclo01\f1r3\` só os JSON e os logs. Uma busca confirmou que não há senha
  nem caminho de rede neles.
- Criei e encerrei só os PIDs 24804, 2824, 26368 e 23896. No fim, não havia nenhum EXCEL.EXE aberto.
- A produção continua com sha `fe10dba1f1fd`.
- Os JSON dos passes anteriores (`f1_A.json`/`f1_B.json` na raiz da pasta do ciclo e `f1r2\`) foram mantidos.
- Ferramenta: `f1_entrada_ciq.py` ganhou a sessão E. Depois da execução, corrigi um detalhe de saída: o campo
  `linhas_do_run` comparava RUN float com int e saiu vazio. A reconstrução da seção 3 foi feita a partir de
  `pop_maiusc`, que estava completo.
