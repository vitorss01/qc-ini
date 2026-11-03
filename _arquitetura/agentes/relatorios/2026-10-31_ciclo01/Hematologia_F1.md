# Hematologia, F1 (entrada de CIQ manual e automática): ciclo 2026-10-31_ciclo01, 2º passe

Agente: Usuário Simulado. Modo: completo. Data: 01/11/2026, das 02h42 às 03h09.

Este passe refez a F1 inteira da Hematologia numa produção que não mudou desde o 1º passe (31/10, 21h40–22h10):
automático, manual, validações, equivalência, SUBSTITUI_ID e minúsculas (sessões A, B, C e D). Todos os números do
1º passe se repetiram bit a bit. Acrescentei duas sessões para fechar os limites que o 1º passe tinha deixado:

- **F:** o que a validação de dados do Excel aceita quando o usuário **digita** (até aqui, só as regras tinham sido
  lidas), e manuais com espaço sobrando no analito ou no equipamento;
- **G:** resultado implausível (−1 e 0) digitado num manual.

Saíram daí três ocorrências novas (OC-HF1-7 a 9) e evidência nova para a OC-HF1-1: **"wbc" em minúsculas passa pela
validação mesmo digitado**.

## 1. Ambiente e evidência

| Item | Valor |
|---|---|
| Produção copiada | `QC_Hematologia.xlsm`, sha256 curto `0b45fa270d2a`, última gravação 30/10/2026 22:11 (a mesma do 1º passe). O hash é o mesmo antes e depois do passe: a produção não foi aberta |
| Cópias | Uma cópia nova da produção por sessão: `%TEMP%\qcw\auditoria\2026-10-31_ciclo01\hema_f1_r2\Hematologia_F1_<A,B,C,D,F,G>.xlsm`. As 6 tinham `0b45fa270d2a` antes e depois, ou seja, nada foi salvo. PIDs do Excel: A 27756, B 23016, C 18544, D 12248, F 20728 e G 10740, cada um encerrado no `finally` |
| Origem automática | Cópia de arquivo do DB_SEAC de produção (`<UNC>`, só leitura), feita em 01/11 às 02h43: `DB_SEAC_congelado.xlsm`, sha `1769a4f546b8`, gravado pela tarefa agendada às 02h42. Tem 41.788 linhas: 403 a menos que no 1º passe, porque a janela de 12 meses andou. O resultado mais recente ainda é de **30/10 08:06:36**. A Cfg_Integracao de cada cópia foi apontada para ela, como o administrador faria |
| Configuração da cópia | SETOR=HEMATOLOGIA, MODO_FONTE=SEAC, MOTOR_DADOS=VBA, PREFIXO_ID=HEM, EQUIPAMENTO_PADRAO=XN1000, MATRIZ_PADRAO=SANGUE TOTAL, GAP_CORRIDA_MIN=10, TOLERANCIA_CONFLITO_MIN=30, NIVEIS=3 |
| Ferramenta | `_arquitetura/agentes/ferramentas/f1_entrada_ciq.py --produto Hematologia`. Neste passe ganhou as **sessões F e G**, documentadas no cabeçalho, e as sessões A a E não mudaram. Runner: `hema_f1_r2\rodar.py` (copia a produção, roda a sessão e confere os hashes). Saídas: `f1_A.json` … `f1_G.json` e os `.log`, em `%TEMP%\qcw\auditoria\2026-10-31_ciclo01\hema_f1_r2\` |
| Dados de teste | Todos levam MOTIVO `Outro: TESTE-AUDITORIA`. Os valores são sintéticos à vista: 3 casas decimais (a Hema manda no máximo 2), ou −1 e 0. WBC N1 = 3,111, N2 = 7,322 e N3 = 17,433. Na cópia B, os IDs sintéticos são HEM-9000001 a 9000003 |

**Como os passos foram executados**

- **Digitação:** escrevi na `tblResultados_Manuais` célula a célula (`FormulaLocal`, como no Excel pt-BR), com os
  eventos ligados. É assim que o `Worksheet_Change` carimba o ID.
- **ATUALIZAR:** rodou por `mIntegracao.AtualizarDadosAutomatico`, o núcleo do botão sem a caixa de diálogo.
- **Filtros:** H3 (lote), C3 (analito), K3 (ano) e K4 = "Tudo (sem filtro)" pela própria célula do Painel; Ano
  De/Até (B3/D3) na aba Estatística.
- **Validação × digitação (sessão F):** escrevi cada texto numa linha livre com os eventos desligados e li
  `Range.Validation.Value`, que é a avaliação da regra feita pelo próprio Excel (True = a regra aceita; com
  `ShowError`, False = a caixa de erro bloquearia a digitação). Depois limpei a célula. Conferi que a linha terminou
  vazia.

**Limites do método**

1. A caixa de erro da validação continua sem ser vista na tela (automação COM). O que este passe acrescenta é o
   veredicto da regra pelo próprio Excel (`Validation.Value`). Colar, ou escrever por COM, passa por cima da
   validação. Por isso os casos marcados "colado" representam o usuário que cola o valor.
2. A caixa de aviso do SUBSTITUI_ID inexistente não aparece em automação. Conferi só o efeito.
3. As cópias foram abertas pelo `qa_final.QA`, sem login, e por isso USUARIO fica "(sem login)". O ADR-084 não foi
   exercitado nesta frente.
4. O MOTIVO de teste (`Outro: TESTE-AUDITORIA`) **não passa** pela validação de lista quando digitado (sessão F). Ele
   só entra colado ou por COM. Isso não altera nenhum resultado (o ETL aceita qualquer MOTIVO não vazio), mas é a
   origem da OC-HF1-9.

## 2. Testes executados

| ID | Teste | Estado | Evidência |
|---|---|---|---|
| T01 | Tela Digitar Resultados: campos, validação, proteção e instruções | APROVADO | 17 colunas e aba protegida. Validações: DATA ≥ 01/01/2020, sem limite superior; EQUIPAMENTO por lista **literal** "XN1000"; MATRIZ "SANGUE TOTAL"; LOTE `=lstLotes` (9 lotes, de 522611 a 625311); ANALITO `=lstAnalitos` (28); NIVEL inteiro de 1 a 3; RESULTADO decimal ≥ −1E+307; MOTIVO por lista literal ("Falha do interfaceamento", "Resultado não transmitido", "Equipamento em contingência", "Outro"). HORA e SUBSTITUI_ID sem validação. As instruções A3:A4 citam DATA, HORA, LOTE, NÍVEL, ANALITO, RESULTADO e MOTIVO |
| A01 | Recebimento DB_SEAC → `tblDB_Recebimento` | APROVADO | Origem: 41.788 linhas e 41.788 IDs válidos, nenhum sem ID e nenhum repetido. Recebimento: 44.392 → 46.779 (+2.387). 0 IDs da origem não recebidos e 0 duplicados. As 403 linhas que saíram da janela da origem desde o 1º passe continuam no recebimento, que guarda o histórico |
| A02 | Preservação do ID e associação: origem × recebimento × tblCQ_Final | APROVADO | 0 divergências origem × recebimento nos 2.387 novos. 0 divergências nos 10 campos origem × final (EQUIPAMENTO, LOTE, NIVEL, ANALITO→ANALITO_ORIGEM, DATA, DATA_HORA, VALOR→RESULTADO, MATRIZ, ID_ORIGEM e ID_AMOSTRA). ID_REGISTRO = `HEM-<ID_ORIGEM>`. Os 2.387 entram ATIVO, todos do XN1000 |
| A03 | Sync_Log e QA | APROVADO | OK, VBA, SEAC, NOVOS 2.387, TOTAL_RECEBIDO = LINHAS_CQ = 46.779, DATA_MAX 30/10/2026 08:06 (= máximo da origem), 66,2 s, dos quais 52,6 s no motor. QA: 0 erros, 589 A02 (retransmissões fora do cálculo), 3 I01 e 31 I02 |
| A04 | Histórico intacto | APROVADO | 0 IDs antigos com RUN ou STATUS mudado. Final com 46.779 linhas e 46.779 IDs únicos |
| A05 | Idempotência (2º ATUALIZAR com o mesmo arquivo) | APROVADO | OK_SEM_MUDANCA, final idêntica, 18,5 s |
| A06 | Plausibilidade do automático (valor ≤ 0) | APROVADO | 0 resultados ATIVO com valor ≤ 0, nos novos e na base inteira |
| M01 | Manual válido N1, N2 e N3 (WBC, lote 625311, XN1000, 30/10 09:36–09:38) | APROVADO | MAN_0001 a 0003 nascem ao digitar, com REGISTRADO_EM, XN1000 e SANGUE TOTAL. Antes do ATUALIZAR: "—" / "atualize os dados". Depois (23,8 s): ATIVO, RUN 26103002, CORRIDA_NO_DIA 2, ORIGEM=MANUAL, PARTICIPA=SIM. `DadosDoId` devolve o ponto completo. Audit_Log: RESULTADO_MANUAL_INCLUIDO "26103002\|n\|WBC". QA: 3 **ALERTA A06** (o manual no mesmo dia do automático, em corrida diferente), com a orientação do SUBSTITUI_ID. Correção do 1º passe: o A06 é ALERTA, e não informativo |
| M02 | n, média, DP e CV (Painel `engPainel` e Estatística) × cálculo independente | APROVADO | Seção 3. O n subiu exatamente 1 por nível. Diferença relativa máxima entre as três fontes: 1,5e-16 |
| M03 | Levey-Jennings e Westgard | APROVADO | Eng_Saida, slot 41, RUN 26103002: N1=3,111, N2=7,322, N3=17,433, N1 ID MAN_0001, veredicto OK nos 3 níveis |
| V01 | Campos vazios | APROVADO | MANUAL_INCOMPLETO "...: DATA, HORA, LOTE, NIVEL", E04 |
| V02 | Resultado "abc" | APROVADO | MANUAL_INCOMPLETO "...: RESULTADO", E04. A validação também recusa ao digitar (F) |
| V03 | Resultado "90.5" | APROVADO | MANUAL_INCOMPLETO "RESULTADO ambiguo (use virgula decimal: 1,5)", E04. A validação recusa ao digitar (F) |
| V04 | Resultado "1.5" (risco de virar data) | APROVADO | Não virou data. "ambiguo", E04 |
| V04b | Resultado "3,111" | APROVADO | Número 3,111, exibido como "3,111" |
| V05 | Data futura (02/11/2026) | APROVADO | A validação aceita (F: 01/11/2027 também passa). O ATUALIZAR barra: MANUAL_INCOMPLETO "DATA/HORA posterior ao lancamento", E04 |
| V06 | Analito inexistente "Glicoze" (colado) | REPROVADO | ATIVO, PARTICIPA=SIM, ANALITO_CADASTRADO=NÃO, fora do Painel, só um I01 agregado. **Digitado, a validação recusa** (F). OC-HF1-3 |
| V07 | Nível 4 | APROVADO | MANUAL_INCOMPLETO "NIVEL", E04. A validação recusa 4, 0, 1,5 e "abc" (F) |
| V08 | Lote inexistente "99999" (colado) | REPROVADO | ATIVO, PARTICIPA=SIM, sem mensagem no QA. Digitado, a validação recusa (F). OC-HF1-3 |
| V09 | Duplicidade manual × manual | APROVADO | MAN_0012 fica CONFLITO_MANUAL, ID_RELACIONADO MAN_0001, A01. O n do Painel não muda |
| V10 | Duplicidade manual × interfaceamento (+5 min, mesmo valor) | APROVADO | MAN_0013 fica CONFLITO_MANUAL "coincide com resultado do interfaceamento (5 min)", ID_RELACIONADO HEM-372173, A01 |
| V11 | Equipamento inexistente "XN9999" (colado) | APROVADO (com ressalva, OC-HF1-3) | ATIVO, fora do Painel, ALERTA A10 com o nome do equipamento e como corrigir. Digitado, a validação recusa (F) |
| V12 | Hora "25:99" | APROVADO | MANUAL_INCOMPLETO "HORA", E04. HORA não tem validação de dados: o texto fica na célula até o ATUALIZAR |
| V13 | Data "31/02/2026" (colada) | REPROVADO | Vira 03/03/2026 08:00, ATIVO, RUN 26030301, sem aviso, e entra no N1 do lote (n 22 → 23). Digitada, a validação recusa (F). OC-HF1-2 |
| V15 | EQUIPAMENTO vazio | APROVADO | Recebe "XN1000" (EQUIPAMENTO_PADRAO) e entra ATIVO na série certa (RUN 26102904) |
| V16 | Equipamento "xn1000" (colado) | REPROVADO | ATIVO, sem QA, RUN 26102901 (o da 1ª corrida real do dia). O N3 do Painel fica com n 22, com a média trocada. Digitado, a validação **recusa** (F, lista literal). OC-HF1-1 |
| E01 | Equivalência: o mesmo trio pelo manual (A) e pelo automático (B, DB_SEAC sintético HEM-9000001 a 9000003) | APROVADO | Mesmo RUN 26103002 e mesma corrida 2. `engPainel`, as linhas da Estatística e o recálculo independente idênticos. As 15 últimas linhas do Eng_Saida (39 colunas) diferem só na coluna "N1 ID" (MAN_0001 × HEM-9000001). Sync_Log de B: NOVOS 3, OK |
| S01 | SUBSTITUI_ID: digitar só o número (372173) preenche a chave | APROVADO | MAN_0001, "HEM-372173", 30/10/2026 08:05, XN1000, SANGUE TOTAL, 625311, N1, WBC |
| S02 | SUBSTITUI_ID inexistente (987654321) | APROVADO | Nada preenchido |
| S03 | Substituição com MOTIVO (3,111 no lugar de 3,11) | APROVADO | MAN_0001 fica ATIVO, RUN 26103001, corrida 1. HEM-372173 fica SUBSTITUIDO_MANUAL, sem RUN. 0 outros RUN mudaram. I04 com os dois valores, a diferença e como desfazer. Audit_Log: SUBSTITUIDO_POR_MANUAL + MANUAL_INCLUIDO. Eng_Saida, slot 40: N1=3,111 |
| S04 | SUBSTITUI_ID sem MOTIVO | APROVADO | MAN_0002 fica MANUAL_INCOMPLETO "MOTIVO (obrigatorio para substituir...)", E04. HEM-372248 continua ATIVO. A mensagem final diz "Há ERRO DE GOVERNANÇA" (OC-HF1-8) |
| S05 | Estatística depois da substituição × cálculo independente | APROVADO | N1: n 21, média 3,1205238095238097 = (21 × 3,120476… − 3,11 + 3,111)/21. Painel = Estatística = independente |
| D01 | Linha de base da sessão D | APROVADO | N3: n 21, média 17,383333…, DP 0,17906237274573683, CV 1,0300807636379876 nas três fontes. RUN 26102901: HEM-371215 (N2 7,39), HEM-371290 (N3 17,58), HEM-371365 (N1 3,21) |
| D02 | Manual N3 com EQUIPAMENTO "xn1000" (29/10 23:36, 17,433) | REPROVADO | RUN 26102901, corrida 1, ATIVO, sem QA. O Painel troca o 17,58 real pelo 17,433 (n 21, média 17,376333…). A Estatística conta os dois (n 22). O Eng_Saida do RUN mostra N3 17,433 com data 23:36. OC-HF1-1 |
| D03 | Manual N3 com ANALITO "wbc" (28/10 23:36) | REPROVADO | ATIVO, ANALITO_CADASTRADO=NÃO, RUN 26102801. O Painel troca o 17,14 real (HEM-370716) pelo 17,433 (n 21, média 17,390285…). A Estatística fica com n 23. **"wbc" passa pela validação digitado** (F1). OC-HF1-1 |
| F01 | Validação de dados × o que o usuário digita (32 casos) | APROVADO (como registro). Achados nas OC-HF1-1, 7 e 9 | Tabela na seção 3.3. Destaques: EQUIPAMENTO por lista literal **diferencia maiúsculas** ("xn1000" e "Xn1000" recusados); ANALITO por `lstAnalitos` **não diferencia** ("wbc" aceito); RESULTADO aceita −1; MOTIVO só aceita os 4 itens da lista |
| F02 | Espaço sobrando: ANALITO "WBC " (27/10 23:36) e EQUIPAMENTO "XN1000 " (26/10 23:36), N3 | APROVADO | O ETL apara os dois. Na tblCQ_Final ficam "WBC", ANALITO_CADASTRADO=SIM, e "XN1000". Cada um entra numa corrida própria (RUN 26102703 e 26102603, CORRIDA_NO_DIA 3), com A06. Os RUN reais do dia ficam intactos (26102701: N3 17,2; 26102601: N3 17,22). O Painel e a Estatística vão a n 23, média 17,38765217391304, iguais ao independente. A validação aceita "XN1000 " e recusa "WBC " ao digitar |
| G01 | Resultado implausível digitado: N1 = −1 (25/10 23:36) e N2 = 0 (24/10 23:36), com analito, lote e equipamento certos | REPROVADO | Os dois ficam **ATIVO**, PARTICIPA=SIM, com corrida própria (RUN 26102501 e 26102401), **sem nenhuma mensagem no QA**, e a validação aceita os dois ao digitar. N1: CV de 1,736% para **30,004%**, média 2,9332, n 22. N2: CV de 1,548% para **16,292%**. O Westgard marca REJEITADO (1-3s) nas duas corridas. OC-HF1-7 |

## 3. Estatística: Painel × Estatística × cálculo independente

Filtro: WBC, lote 625311 (H3 = Estatística B4), XN1000. No Painel, Período "Tudo (sem filtro)". Na Estatística, Ano
De = Até = 2026 (o lote só tem dados de 2026).

**Como calculei**

- **População:** linhas da tblCQ_Final com PARTICIPA=SIM, ANALITO = "WBC", LOTE = "625311", EQUIPAMENTO = "XN1000" e o
  nível. Na Estatística, filtrei ainda pelo ano de DATA_HORA (`src/hema/mEstatistica.EstatBasica`).
- **Cálculo:** `statistics.fmean`, `statistics.stdev` (n−1) e CV = DP/média × 100.
- **Tolerância:** 1e-9 relativo, com comparação pelo `Value2`.

### 3.1 Depois do manual (M02, cópia A; idêntico na cópia B, automático)

| Nível | Fonte | n | Média | DP | CV % |
|---|---|---|---|---|---|
| N1 antes | as três | 21 | 3,1204761904761904 | 0,05417344279960346 | 1,7360633279286932 |
| N1 depois | engPainel = Estatística = independente | 22 | 3,120045454545455 | 0,05290645587821089 | 1,6956950355044937 |
| N2 antes | as três | 39 | 7,337179487179487 | 0,11357697871693422 | 1,5479651126893008 |
| N2 depois | engPainel = Estatística | 40 | 7,336799999999999 | 0,11213709558004521 | 1,5284196867850455 |
| N2 depois | independente | 40 | 7,336799999999999 | 0,11213709558004521 | 1,5284196867850457 |
| N3 antes | as três | 21 | 17,383333333333333 | 0,17906237274573683 | 1,0300807636379876 |
| N3 depois | engPainel = Estatística = independente | 22 | 17,385590909090908 | 0,1750675132602679 | 1,0069690134531193 |

**Status: APROVADO.** A maior diferença relativa foi de 1,5e-16.

### 3.2 Casos com divergência ou dado errado

| Caso | Fonte | n | Média | DP | CV % |
|---|---|---|---|---|---|
| D02 "xn1000" | Painel (trocou o 17,58 de HEM-371290 pelo 17,433) | 21 | 17,37633333333333 | 0,1737853081630701 | 1,0001264641355299 |
| D02 | Estatística (contou os dois) | 22 | 17,385590909090908 | 0,1750675132602679 | 1,0069690134531193 |
| D02 | Independente, equipamento exato | 21 | 17,383333333333333 | 0,17906237274573686 | 1,0300807636379876 |
| D03 "wbc" | Painel (trocou o 17,14 de HEM-370716) | 21 | 17,390285714285714 | 0,16542313709307477 | 0,9512387536978966 |
| D03 | Estatística | 23 | 17,38765217391304 | 0,17132786450005755 | 0,9853421427253034 |
| F02 espaços | Painel = Estatística = independente | 23 | 17,38765217391304 | 0,17132786450005755 | 0,9853421427253034 |
| G01 N1 com −1 | Painel = Estatística = independente | 22 | 2,933181818181818 | 0,8800778448449961 | 30,004203605439198 |
| G01 N2 com 0 | Painel = Estatística = independente | 40 | 7,15375 | 1,165514495174358 | 16,292357087881992 |

**Como reconstruir:**

- **D02:** soma dos 21 N3 = 17,383333… × 21 = 365,05; − 17,58 + 17,433 = 364,903; ÷ 21 = 17,376333…, o que bate com o
  Painel.
- **D03:** a partir do D02, 364,903 − 17,14 + 17,433 = 365,196; ÷ 21 = 17,390285…, o que também bate com o Painel.
- **F02:** os mesmos 23 valores do D03, mas aqui o ETL aparou os espaços e o Painel concorda com a Estatística.
- **G01:** a conta está certa nas três fontes. O defeito é o dado, que entrou sem barreira (OC-HF1-7).

### 3.3 Validação de dados × digitação (sessão F, `f1_F.json` → F1_validacao)

| Coluna | Digitado | Excel guarda | Regra aceita? |
|---|---|---|---|
| EQUIPAMENTO | XN1000 / " XN1000" / "XN1000 " | texto | sim / sim / sim |
| EQUIPAMENTO | xn1000 / Xn1000 / XN9999 | texto | **não** / não / não |
| ANALITO | WBC / **wbc** | texto | sim / **sim** |
| ANALITO | "WBC " / Glicoze | texto | não / não |
| LOTE | 625311 / 99999 / "625311 " | texto | sim / não / não |
| DATA | 30/10/2026 / 2026-10-30 / 01/11/2027 | data | sim / sim / sim (futura passa) |
| DATA | 31/02/2026 / abc | texto | não / não |
| HORA | 08:05 / 25:99 / abc | hora / texto / texto | sem validação |
| NIVEL | 1 / 4 / 0 / 1,5 / abc | número / … / texto | sim / não / não / não / não |
| RESULTADO | 3,111 / **−1** | número | sim / **sim** |
| RESULTADO | 90.5 / abc | texto | não / não |
| MOTIVO | "Outro: TESTE-AUDITORIA" / texto livre | texto | **não** / não |

**Leitura:**

- A lista literal do EQUIPAMENTO diferencia maiúsculas. A lista por intervalo (`lstAnalitos`, `lstLotes`) não
  diferencia.
- A OC-HF1-1 é alcançável por **digitação** pelo analito, e pelo equipamento só colando. Na Hema o equipamento vem
  preenchido sozinho, então o caminho realista é o analito.
- V06, V08, V11, V13 e V16 só acontecem colando.

## 4. Ocorrências

### OC-HF1-1: Equipamento ou analito em minúsculas faz o manual "tomar o lugar" de um resultado real no Painel e no Levey-Jennings, e a Estatística o conta a mais (Crítico; tipo `calculo`)

Reproduzida bit a bit neste passe. **Evidência nova:** "wbc" passa pela validação do Excel digitado.

- **Tela/módulo:** Digitar Resultados → ATUALIZAR.
  - `mEtlCQ.CalcularCorridas` monta a chave da corrida (`_GK`) com o texto exato de EQUIPAMENTO e ANALITO
    (`Option Compare Binary`, linhas 21 e 1042–1046).
  - O motor da Hema (`mEstatistica.GarantirIndice`, `EquipFiltro`) compara em maiúsculas e guarda um valor por RUN e
    nível no Eng_Saida.
- **Ação:** digitar "wbc" no ANALITO (D03; aceito pela validação) ou colar "xn1000" no EQUIPAMENTO (D02) de um manual
  N3 válido, numa hora livre de um dia com corrida real, e clicar em ATUALIZAR.
- **Esperado:** o texto é normalizado para a grafia do cadastro, ou o manual fica MANUAL_INCOMPLETO, ou sai um
  ALERTA. Em nenhum caso um resultado real pode sumir do gráfico, nem o Painel e a Estatística podem divergir.
- **Observado:**
  - O manual recebe o RUN da **1ª corrida real do dia** (26102901 / 26102801), porque para o ETL ele é outra série.
  - No Painel e no LJ, o N3 manual **substitui** o N3 real desse RUN (17,58 de HEM-371290; 17,14 de HEM-370716), e o
    ponto passa a ter a data 23:36.
  - A Estatística conta os dois (n 22 / 23, contra 21 no Painel).
  - Nenhuma mensagem no QA.
  - No D03 a linha diz ANALITO_CADASTRADO=NÃO, mas entra no WBC do Painel e da Estatística.
- **Dificuldade:** o STATUS diz ATIVO e a dica mostra o valor manual, então o usuário acha que deu certo. O valor real
  some do gráfico sem registro.
- **Consequência:**
  - estatística divergente entre o Painel e a Estatística para o mesmo filtro;
  - LJ com ponto trocado e data errada;
  - Westgard avaliado sobre o valor errado na corrida real.

  É associação errada entre resultado e corrida (REGRAS_COMUNS §4).
- **Sugestão (decisão do Especialista em CQ, do Orquestrador e do usuário):**
  - normalizar EQUIPAMENTO e ANALITO do manual para a grafia do cadastro no ETL, nos dois motores;
  - ou usar a mesma comparação no ETL e no motor;
  - no motor, nunca sobrescrever um valor já presente no mesmo RUN e nível sem alerta.

  O F02 mostra que o ETL já apara espaços: falta tratar a caixa das letras. A Bio tem o mesmo defeito (OC-9 da
  Bioquimica_F1).
- **Evidência:** `hema_f1_r2\f1_D.json` (D1, D2, D3: painel, estat, ind_exato, ind_maiusc, eng_run, linhas_do_run,
  final, qa = []); `f1_F.json` → F1_validacao (ANALITO "wbc" com validacao_aceita = true; EQUIPAMENTO "xn1000" com
  false); `f1_A.json` → A4_validacoes.casos["V16 ..."].

### OC-HF1-7 (nova): Manual com resultado implausível (−1, 0) entra ATIVO na estatística sem nenhum alerta (Alto; tipo `dados`)

- **Tela/módulo:** Digitar Resultados (validação de RESULTADO = decimal ≥ −1E+307) → ATUALIZAR (`mEtlCQ`, QA).
- **Ação:** digitar um manual WBC N1 com RESULTADO "−1" e outro N2 com "0", com data, hora, lote, equipamento,
  analito e motivo certos, e clicar em ATUALIZAR.
- **Esperado:** a tela ou o ATUALIZAR recusa ou alerta um valor fisicamente impossível (WBC ≤ 0). No mínimo, o QA dá
  um ALERTA com o ID.
- **Observado:**
  - A validação aceita os dois ao digitar.
  - Depois do ATUALIZAR, ficam ATIVO, PARTICIPA=SIM, com corrida própria (RUN 26102501 e 26102401). QA: 0 erros e
    nenhum código para os IDs.
  - O N1 do lote vai de CV 1,736% para **30,004%** (média 2,933), e o N2 de 1,548% para **16,292%**.
  - O Westgard marca REJEITADO (1-3s) nas duas corridas, e é o único sinal.
- **Dificuldade:** um erro de digitação (sinal trocado, zero no lugar do valor) passa por todas as barreiras.
- **Consequência:** média, DP, CV e, por consequência, ET e Sigma do lote ficam errados até alguém inativar. É a mesma
  família da OC-1 da Bioquímica (valores ≤ 0 do interfaceamento). Na Hema, o automático tem hoje 0 valores ≤ 0
  (A06), e o risco vem do manual.
- **Sugestão (Especialista em CQ e RT):** a mesma triagem de plausibilidade sugerida na OC-1 da Bio, aplicada também
  ao manual:
  - resultado ≤ 0, ou |z| extremo em relação ao alvo do lote, gera ALERTA com o ID, ou MANUAL_INCOMPLETO;
  - validação de RESULTADO > 0 na tela.
- **Evidência:** `hema_f1_r2\f1_G.json` (G_antes, G_depois, G_linhas, G_qa_codigos); `f1_F.json` → F1_validacao
  (RESULTADO "-1" aceito).

### OC-HF1-2: Data inválida colada ("31/02/2026") vira 03/03/2026 e o manual entra ATIVO no lote (Alto; tipo `dados`)

É a mesma OC-2 da Bioquímica (código comum, `mEtlCQ.ParaData`). Reproduzida.

- **Observado:** MAN_0016 fica com 03/03/2026 08:00, ATIVO, RUN 26030301, sem aviso, e entra no N1 do lote 625311
  com Período "Tudo" (n 22 → 23).
- **Novo:** confirmei pelo Excel que, **digitada**, a data é recusada (`Validation.Value` = False, "Data inválida.").
  O risco fica restrito a colar.
- **Sugestão:** conferir dia, mês e ano depois de converter. Se não baterem, MANUAL_INCOMPLETO "DATA", nos dois
  motores.
- **Evidência:** `f1_A.json` → A4_validacoes.casos["V13 ..."] e painel; `f1_F.json` → DATA "31/02/2026".

### OC-HF1-3: Manual com analito ou lote inexistente fica ATIVO sem aparecer no Painel; o equipamento inexistente ao menos gera A10 (Médio; tipo `funcional`)

É a mesma OC-4 da Bioquímica. Reproduzida.

- **Observado:**
  - "Glicoze" e o lote "99999" (colados) ficam ATIVO e PARTICIPA=SIM, fora do Painel. O analito gera só o I01
    agregado, e o lote não gera nada.
  - "XN9999" gera o ALERTA A10.
  - Os três são **recusados pela validação quando digitados** (sessão F), então o caminho é colar.
- **Consequência:** resultado perdido em silêncio, enquanto a instrução A4 diz "ATIVO = entrou na estatística".
- **Sugestão:** MANUAL_INCOMPLETO ou um ALERTA com o ID, como o A10 já faz com o equipamento.
- **Evidência:** `f1_A.json` → A4_validacoes.casos V06, V08 e V11, e qa_agregados; `f1_F.json`.

### OC-HF1-8 (nova): Linha manual incompleta faz o ATUALIZAR dizer "Há ERRO DE GOVERNANÇA" (Médio; tipo `usabilidade`)

- **Tela/módulo:** mensagem final do ATUALIZAR (`mIntegracao.ResumoQA`, linhas 370–375). Ela acrescenta a frase
  sempre que houver **qualquer** ERRO no QA.
- **Ação:** deixar uma linha manual incompleta (por exemplo o SUBSTITUI_ID sem MOTIVO, S04, ou qualquer V01–V12) e
  clicar em ATUALIZAR.
- **Esperado:** uma mensagem que diga que há manual incompleto, e em qual linha ou ID (o E04 já tem o texto certo).
- **Observado:** "QA da integração: 1 erro(s)… Há ERRO DE GOVERNANÇA: veja a aba QA_INTEGRACAO (ex.: inativação sem
  motivo, ou ID que não é do analito informado -- essa não foi aplicada)." Isso aparece na cópia C (1 E04) e na
  cópia A depois das validações (7 E04). O erro real é só o E04 do manual.
- **Dificuldade:** o usuário procura uma inativação que não existe e pode achar que fez algo grave.
- **Sugestão:** falar em governança só quando os códigos forem de governança (E01, E10 etc.), e para o E04 dizer
  "N resultado(s) manual(is) incompleto(s): veja o DETALHE na aba Digitar Resultados". Relaciona-se com a OC-5 da
  Bioquímica, mesma função.
- **Evidência:** `f1_C.json` → C.resumo2 e C.qa (só E04 como ERRO); `f1_A.json` → A4_validacoes.resumo e qa_codigos
  ("ERRO | E04": 7).

### OC-HF1-4: A coluna Status do Painel fica "REJEITADO" nos 3 níveis por uma regra 6x de 09/10, mesmo depois de corridas OK (Médio; tipo `usabilidade`)

Reproduzida.

- **Observado:** J7:J9 = REJEITADO, "6x · RUN 26100902, Erro Sistemático", com Período "Tudo", também depois da
  corrida manual 26103002, OK nos 3 níveis.
- **Sugestão:** deixar claro que é o pior status do período, ou mostrar também o da última corrida (UX/UI e
  Especialista em CQ).
- **Evidência:** `f1_A.json` → A3_manual.painel_depois (J7:J9 e engPainel colunas 10 e 19–21).

### OC-HF1-9 (nova): MOTIVO só aceita os 4 itens da lista; "Outro" não tem onde dizer qual é o motivo (Baixo; tipo `usabilidade`)

- **Tela/módulo:** Digitar Resultados, coluna MOTIVO (lista literal "Falha do interfaceamento, Resultado não
  transmitido, Equipamento em contingência, Outro").
- **Observado:**
  - Digitar "Outro: TESTE-AUDITORIA" ou qualquer texto livre é recusado pela validação ("Escolha o motivo da lista.").
  - Colado, o mesmo texto entra e aparece no I04 e no MOTIVO_MANUAL.
  - A tabela não tem nenhuma coluna de observação.
- **Consequência:** numa substituição (ADR-082) com motivo "Outro", a trilha registra só "Outro". Já a inativação
  aceita motivo em texto livre (ADR-084).
- **Sugestão:** permitir "Outro: <texto>" (validação por "começa com um item da lista", ou sem erro bloqueante), ou
  criar uma coluna de observação (UX/UI e RT).
- **Evidência:** `f1_F.json` → F1_validacao (MOTIVO, os dois casos com validacao_aceita = false); `f1_A.json` →
  telas.MOTIVO_validacao.

### OC-HF1-5: O Audit_Log grava RESULTADO_MANUAL_INCLUIDO também para linhas incompletas e em conflito (Baixo; tipo `funcional`)

É a mesma OC-8 da Bioquímica. Reproduzida: MAN_0002 da cópia C (sem MOTIVO) gera RESULTADO_MANUAL_INCLUIDO com chave
"0|2|WBC" e RUN vazio.

- **Evidência:** `f1_C.json` → C.audit; `f1_A.json` → A4_validacoes.audit.

### OC-HF1-6: As colunas RUN e DETALHE da aba Digitar Resultados mostram "0" (Baixo; tipo `usabilidade`)

É a mesma OC-6 da Bioquímica. Reproduzida: linhas ATIVO mostram DETALHE = 0, e linhas incompletas ou em conflito
mostram RUN = 0. O DETALHE vem sem acento.

- **Evidência:** `f1_A.json` → A3_manual.linhas_digitadas; `f1_C.json` → C.linha_sem_motivo.

## 5. Dúvidas de regra de negócio (ao Especialista em CQ, pelo Orquestrador)

1. O EQUIPAMENTO e o ANALITO do manual devem ser normalizados para a grafia do cadastro, ou recusados (OC-HF1-1)? O
   motor do Painel pode guardar dois valores do mesmo nível num RUN?
2. Qual é a triagem de plausibilidade aceitável para o manual (≤ 0? |z| acima de quanto?), e ela deve bloquear ou só
   alertar (OC-HF1-7)?
3. Uma corrida REJEITADA pelo Westgard (G01: 1-3s com −1) continua entrando no n, na média e no DP do Painel e da
   Estatística. É por projeto, até o usuário inativar?
4. A coluna Status do Painel deve ser o pior status do período, ou o da última corrida (OC-HF1-4)?
5. O MOTIVO do manual precisa aceitar texto livre ou um complemento quando for "Outro" (OC-HF1-9)?
6. O arquivo de produção (30/10 22:11) tem a base só até 01/10 e a Sync_Log vazia. O primeiro ATUALIZAR de uma cópia
   recebe 2.387 resultados (02/10 a 30/10). O laboratório não atualizou depois da entrega, ou usa outro arquivo?
7. Em 01/11 às 02h42, o resultado mais recente do DB_SEAC da Hema ainda era de **30/10 08:06:36**. Não houve controle
   na sexta à tarde, no sábado e no domingo, ou o interfaceamento parou?

## 6. Impressões de uso (Relatório A)

- **É fácil inserir?** Sim:
  - um equipamento só, preenchido sozinho;
  - a matriz vem sozinha;
  - o ID nasce ao digitar;
  - a vírgula decimal funciona, e o ponto é recusado com mensagem clara.

  O SUBSTITUI_ID continua o fluxo mais amigável: o número preenche toda a chave.
- **O usuário sabe o que preencher?** Sim. As instruções A3:A4 listam os campos, e as listas suspensas guiam lote,
  analito, nível e motivo.
- **Há risco de preencher errado?** Sim, e agora medido pelo próprio Excel:
  - **analito em minúsculas** passa pela validação e leva o resultado à corrida errada (OC-HF1-1, o mais grave);
  - **resultado −1 ou 0** passa pela validação e pelo ATUALIZAR, e contamina o CV (OC-HF1-7);
  - data futura passa na tela, mas o ATUALIZAR barra;
  - data inválida, analito, lote ou equipamento inexistentes e equipamento em minúsculas só entram **colados**
    (OC-HF1-2 e 3);
  - espaço sobrando é tratado corretamente (F02).
- **As mensagens de erro são claras?** Na maior parte, sim:
  - o DETALHE diz o campo;
  - o A01 e o A06 ensinam o SUBSTITUI_ID;
  - o I04 mostra os dois valores e como desfazer;
  - o A10 é um bom alerta.

  Ressalvas:
  - a frase "ERRO DE GOVERNANÇA" para um manual incompleto (OC-HF1-8);
  - "0" nas colunas RUN e DETALHE;
  - texto sem acento;
  - ATIVO para resultado invisível;
  - "Outro" sem complemento (OC-HF1-9).
- **Dá para ver o que foi salvo?** Sim, depois do ATUALIZAR (18 a 25 s; 66 s no primeiro, com 2.387 novos): STATUS,
  RUN, DETALHE e a dica do ponto.
- **Automático:** confiável na transferência:
  - nada perdido, nada duplicado;
  - ID preservado e associação exata;
  - histórico preservado mesmo depois de sair da janela da origem;
  - idempotente;
  - nenhum valor ≤ 0.
- **Manual × automático:** equivalentes bit a bit para o mesmo dado. **Exceções:** grafia diferente da do cadastro
  (OC-HF1-1) e valor implausível, que nenhum dos dois caminhos barra.
- **Reprodutibilidade:** as sessões A, B, C e D deram exatamente os mesmos resultados do 1º passe (Painel, Estatística,
  vereditos e status de cada caso de validação), com uma origem 403 linhas menor.

## 7. Limpeza

- As 6 cópias `.xlsm` (todas com `0b45fa270d2a` antes de apagar), o DB_SEAC congelado, o cache da origem e o DB_SEAC
  sintético foram apagados.
- Ficam em `%TEMP%\qcw\auditoria\2026-10-31_ciclo01\hema_f1_r2\` só os JSON, os logs e o `rodar.py`. Conferi por
  busca que não há caminho de rede nem senha.
- Criei e encerrei só os PIDs 27756, 23016, 18544, 12248, 20728 e 10740. Os Excel da tarefa agendada (02h42, 03h06)
  não foram tocados, e no fim não havia nenhum EXCEL.EXE meu aberto.
- A produção continua com sha `0b45fa270d2a`. Nenhum arquivo de `src/` foi alterado. Mudou só a ferramenta do agente
  (`ferramentas/f1_entrada_ciq.py`, sessões F e G).
