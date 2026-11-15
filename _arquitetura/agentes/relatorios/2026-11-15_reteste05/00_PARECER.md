# Parecer consolidado: ciclo 2026-11-15_reteste05

Agente: Usuário Simulado (parecer consolidado). Data: 15/11/2026. Modo: reteste, consolidação. Neste passo o Excel não
foi aberto e nenhum número foi recalculado. Tudo aqui vem dos 2 relatórios das frentes desta pasta
(`Bioquimica_F1.md`, `Hematologia_F1.md`) e do resumo estruturado que cada frente devolveu. Os valores são os que as
frentes mediram e registraram.

| Item | Valor |
|---|---|
| Produção auditada (entrega 42) | `QC_Bioquimica.xlsm` sha256 curto **b20fe00f7e51**; `QC_Hematologia.xlsm` sha256 curto **7bab8fc008f6**. Nas duas frentes o hash foi o mesmo antes e depois: a produção não foi aberta para escrita |
| Base de comparação (entrega 41) | Backups de antes da entrega 42: Bio `_backup_pre_integracao_2026-11-15_0011` sha256 curto **8ec815ab653d**; Hema `_backup_pre_integracao_2026-11-15_0458` sha256 curto **aab0a54f464e**. São as produções que o reteste 04 testou. Hash igual antes e depois |
| Lista do reteste | 2 itens: **OC-0071** (Bio: cabeçalho da Estatística com o Painel em "Tudo", ADR-095 §3) e **OC-0012** (Hema: aviso do corte de 180 corridas, ADR-095 §4) |
| Frentes executadas | Só a F1, nos dois produtos, no modo reteste. A F2 (CEQ), a F3 (auditoria matemática independente de bias, ET e Sigma) e a F4 completa (LJ, Westgard, corrida do lote) **não fizeram parte deste reteste** |
| Dados de teste | Bio: nenhum (só filtros do Painel). Hema: nenhum nas cópias nova, base e Bio. Na cópia `Hematologia_F1_lote.xlsm`, 565 linhas de WBC (lotes 533811, 614011 e 619611) ganharam o LOTE "528211" para exercitar o ramo "do lote" do aviso (dado sintético). Nenhuma cópia foi salva e todas foram apagadas no fim. Ficam os JSON e os logs em `%TEMP%\qcw\auditoria\2026-11-15_reteste05\` |
| Excel | Um por vez, aberto e encerrado pelo `qa_final.QA`. Bio F1: PIDs 29768 (base) e 30020 (nova). Hema F1: 11036 (nova), 9732 (base), 4032 (Bio, regressão), 19288 e 30144 (rodadas do lote sintético interrompidas por aba protegida e por falta de diagnóstico) e 18360 (rodada válida). O Excel 13084 da tarefa agendada e o 4984 (que não era das frentes) não foram tocados. Nenhuma captura de tela |
| Ferramentas | `_arquitetura/agentes/ferramentas/reteste05_usuario_f1.py` (Bio) e `reteste05_usuario_f1_hema.py` (Hema, modos padrão e `--lote-sintetico`), documentados no cabeçalho |
| tblCQ_Final | Bio: 35.906 linhas, hash do Value2 `471c3b76f80f008c`, igual na base e na nova. Hema: 48.391 linhas (47.709 participam), hash `ee71bc38576d73a3`, igual na base e na nova; último resultado 13/11/2026 13:10:20 |

**Resumo em números**

| | Bioquímica | Hematologia | Total |
|---|---|---|---|
| Itens do reteste: RESOLVIDA / PERSISTE / EFEITO_COLATERAL | 1 / 0 / 0 (OC-0071) | 1 / 0 / 0 (OC-0012) | 2 / 0 / 0 |
| Testes vizinhos e de regressão: APROVADO / REPROVADO / NAO_EXECUTADO | 2 / 1 / 1 (caso "só o Até") | 2 / 0 / 0 (um deles é a regressão na Bio) | 4 / 1 / 1 |
| Cenários de filtro rodados (base × nova) | 9 (S1–S9) | 7 com dados reais + 4 sintéticos (L1–L4) + 3 de regressão na Bio (B1–B3) | 23 |
| Diferenças numéricas base × nova | 0 (só o D4 de exibição nos cenários "Tudo", que era o objetivo da correção) | 0 (só o carimbo `Now` do Eng_Saida!G1 e o texto do A39, que era o objetivo) | 0 |
| Cálculos no Relatório B (aprovados / com diferença explicada / reprovados) | 4 (4 / 0 / 0) | 12 (11 / 1 / 0) | 16 (15 / 1 / 0) |
| Ocorrências novas das frentes / deduplicadas (P-xx) | 1 / 1 | 0 / 0 | 1 / 1 |
| Altas / Médias / Baixas | 0 / 0 / 1 | 0 / 0 / 0 | 0 / 0 / 1 (nenhuma Crítica) |
| Dúvidas de regra de negócio | 1 | 3 (mais 1 que já é a OC-0039) | 4 |

**Situação dos 2 itens do reteste**

| Item | Bio | Hema | Consolidado |
|---|---|---|---|
| OC-0071 Estatística da Bio em "Tudo": B5/D5 "00/01/1900" e D4 "2026" | **RESOLVIDA** | — (a Estatística da Hema tem filtro próprio) | **RESOLVIDA**. O fluxo vizinho "só o De" mostra o mesmo defeito numa célula vizinha (J5), que já existia na entrega 41: P-01, Baixo |
| OC-0012 Aviso do corte de 180 corridas (efeito colateral da Hema no reteste 04) | Regressão APROVADA: texto e números iguais ao reteste 04, com "— veja a aba Estatística", como o ADR-095 §4 manda | **RESOLVIDA** | **RESOLVIDA**. O ramo "Período inteiro do lote" só pôde ser exercitado com dado sintético (D-02) |

---

## Relatório A: experiência do usuário

### A.1 Funcionalidades testadas

| Frente | Bioquímica | Hematologia |
|---|---|---|
| F1 | OC-0071 refeita pelo caminho do usuário: células de filtro do Painel (L4, C3, H3, K3, K4, G3/G4), escritas com os eventos ligados, e depois a aba Estatística aberta. Em cada cenário foram lidos Formula, Text, Value2 e NumberFormat de A4:J5, as linhas 1–12 (busca de "1900", "yyyy" e "Ano 0"), o Value2 de A1:AU200 e o `engPainel`. 9 cenários: "Tudo", Set/2026, T3/2026, "Ano inteiro", H3 vazio com "Tudo" e com Out/2026, De/Até atravessando o ano, só o De, volta a "Tudo" | OC-0012 refeita pelo caminho do usuário: células de filtro do Painel (C3, H3, K3, K4, G3/G4), com o De/Até gravado como número de série. Em cada cenário foram lidos o A39 (valor e fórmula), A6:Q12, `Eng_Saida!M1`, `AF1`, `engPainel`, `Eng_Saida!A1:AF190` e `Calc!A1:AZ190` (origem do LJ). 7 cenários reais (WBC e HGB, H3 vazio e lote 528211, "Tudo", "Ano inteiro", De/Até, Out/2026), 4 sintéticos para o ramo do lote e 3 de regressão na Bio |

Único passo fora do caminho do usuário: na cópia do lote sintético da Hema, a troca do LOTE foi feita na tblCQ_Final
com os eventos desligados, seguida de `mEstatistica.InvalidarCache`. O motivo é que nenhum lote real tem mais de
180 corridas (o maior, 528211, tem 89).

### A.2 Facilidade de uso

- **Cabeçalho da Estatística da Bio (OC-0071).** Em "Tudo", o Ano e o De/Até ficam vazios, ao lado de "histórico
  inteiro do lote". Com mês, trimestre ou ano, as datas aparecem em dd/mm/aaaa, iguais ao Painel. Atravessando o ano,
  o Ano fica vazio, como o K3 do Painel. Quem chega à Estatística pela dica do aviso de 180 corridas da Bio não vê mais
  a data de 1900. Nenhum "yyyy" literal nem "Ano 0" apareceu nas linhas 1–12, inclusive E8:F9 e C10 (OC-0072).
- **Aviso do corte na Hema (OC-0012).** O aviso se explica sozinho:
  - diz quantas corridas ficaram fora;
  - diz que o corte vale para o gráfico **e** para os indicadores;
  - dá o n do período inteiro por nível, com o recorte ("todos os lotes" ou "do lote");
  - não manda mais o usuário a uma aba que mostra outro ano e outro lote.
- **Troca de lote com período sem dados (Hema).** Com o lote 528211 e o K3 2026, o Painel mudou sozinho para "Tudo" e
  disse por quê ("o filtro anterior não mostrava nenhum resultado dele -- ajustado: período Tudo…"). Isso ajuda.

### A.3 Dificuldades encontradas

- **Só uma das datas preenchida (P-01).** Quem apaga só o Até no Painel vê na Estatística "01/09/2026" a
  "00/01/1900" (J5). Não dá para saber até quando os dados foram considerados. É pouco provável na rotina, mas acontece
  quando alguém apaga só uma das datas. O caso simétrico (só o Até) não foi executado.
- **REJEITADO no modo TODOS da Hema (continua do reteste 04, D-14).** Com o H3 vazio, os três níveis seguem
  REJEITADO, inclusive em Out/2026 (n 21/39/21), com o Painel sem Média, DP e CV (ADR-086). Quem não conhece o ADR
  fica sem saber de onde vem o status.
- **Alvo de um lote no aviso de "todos os lotes" (Hema).** Com o H3 vazio, o A39 começa com "Lote 522611 · alvo …" e
  termina com "(todos os lotes)". Já é a OC-0039 (ABERTA); não foi registrado de novo.

### A.4 Navegação

Nada de novo neste ciclo. O caminho foi o do reteste 04 (filtros do Painel e aba Estatística), e o aviso da Hema
deixou de mandar o usuário à Estatística, o que elimina uma navegação que levava a outro recorte.

### A.5 Apresentação

- A39 da Hema: cerca de 275 caracteres na área mesclada A39:U40, com quebra de linha (33 pt). A tela não foi vista
  (não pode ser capturada). Pelo tamanho, cabe em duas linhas. Não é ocorrência.
- O texto do aviso na Bio continua com "— veja a aba Estatística", e na Hema não. É a diferença que o ADR-095 §4
  manda (a Estatística da Bio segue o Painel e a da Hema não).

### A.6 Tempos observados

As frentes não mediram tempo neste reteste: só filtros, sem ATUALIZAR nem CEQ.

### A.7 Sugestões do usuário (para o Orquestrador priorizar; nenhuma foi implementada)

1. Estatística da Bio: o mesmo tratamento do B5/D5 no J5 (e no H5, para o caso "só o Até"): formato
   `dd/mm/aaaa;;` ou uma condição por célula, sem mexer no `JanelaFim`. Outra opção é o Painel avisar quando só uma das
   datas estiver preenchida (P-01).
2. Hema, modo TODOS: explicar de onde vem o REJEITADO quando o Painel não mostra Média, DP e CV (D-04).
3. Hema, A39 com o H3 vazio: tratar o alvo de um lote só no início do aviso (OC-0039, já aberta).

---

## Relatório B: validação matemática

**Escopo deste ciclo.** Não houve F3: bias, ET e Sigma **não** foram recalculados de forma independente. O que se
validou matematicamente foi:
1. as contagens do aviso do corte de 180 corridas (cortadas, n da janela e n do período inteiro), contra um oráculo
   próprio em Python puro sobre a tblCQ_Final da cópia;
2. a **invariância** dos números entre a entrega 41 e a 42, célula a célula, na precisão total do `Value2` e com
   tolerância **zero**. Zero é a tolerância certa porque é o mesmo motor sobre a mesma tblCQ_Final (hash igual).

**Definição usada pelo oráculo (Hema F1, a mesma conferida no reteste 04):**
- **Resultado que conta:** PARTICIPA_ESTATISTICA = SIM e RESULTADO numérico, do analito (e do lote, se houver), com
  data dentro de [De, Até].
- **Corrida:** o par (RUN, LOTE) distinto (ADR-080). Por RUN só, o fluxo original daria 452 em vez de 458.
- **Janela:** as 180 corridas mais recentes do período. Cortadas = total − 180, e o n da janela é contado por nível.
- **AF1 esperado:** "N1 a · N2 b · N3 c" com o n do período inteiro.

### B.1 Hematologia (origem: Hematologia_F1.md)

| Analito | Equipamento | Lote | Nível | Período | Dados de entrada | Fórmula encontrada | Resultado independente | Resultado QC_INI (Value2) | Diferença | Status |
|---|---|---|---|---|---|---|---|---|---|---|
| WBC | XN1000 | H3 vazio (todos) | N1/N2/N3 | 2026, "Ano inteiro" (fluxo original) | tblCQ_Final; 458 corridas (RUN, LOTE) | corte = corridas − 180; janela = 180 mais recentes; AF1 = n do período | cortadas 278; janela 120/158/112; período 374/400/340 | M1 278; engPainel 120/158/112; AF1 "N1 374 · N2 400 · N3 340"; A39 "Período inteiro (todos os lotes): N1 374 · N2 400 · N3 340 resultados." | 0 | APROVADO |
| WBC | XN1000 | H3 vazio | N1/N2/N3 | "Tudo (sem filtro)" (H4) | tblCQ_Final | idem | período 515/545/479; janela 120/158/112 | M1 431; AF1 "N1 515 · N2 545 · N3 479"; n 120/158/112 | 0 em AF1 e n (o relatório da frente não traz o oráculo explícito do M1 431) | APROVADO |
| WBC | XN1000 | H3 vazio | N1/N2/N3 | De/Até 01/01–31/10/2026, "Personalizado" (H5) | tblCQ_Final | idem | período 357/377/328; janela 125/157/116 | M1 251; AF1 "N1 357 · N2 377 · N3 328"; n 125/157/116 | 0 | APROVADO |
| WBC | XN1000 | H3 vazio | N1/N2/N3 | Out/2026 (H6, sem corte) | tblCQ_Final | n do período | 21/39/21 | M1 0; AF1 vazio; n 21/39/21; sem aviso | 0 | APROVADO |
| HGB | XN1000 | H3 vazio | N1/N2/N3 | De/Até 01/01–31/10/2026 (H7) | tblCQ_Final | idem H5 | período 357/377/328; janela 125/157/116 | M1 251; AF1 "N1 357 · N2 377 · N3 328"; n 125/157/116 | 0 | APROVADO (os números iguais aos do WBC são coerentes: HGB e WBC saem das mesmas corridas do hemograma) |
| WBC | XN1000 | 528211 | N1/N2/N3 | "Tudo" (89 corridas) | tblCQ_Final | n do período, sem corte | 83/86/81 | M1 0; n 83/86/81; sem aviso | 0 | APROVADO |
| WBC | XN1000 | 528211 | N1/N2/N3 | K3 2026 → ajuste automático para "Tudo" | tblCQ_Final | idem | 83/86/81 | idem, com o aviso do ajuste | 0 | APROVADO |
| WBC | XN1000 | 528211 (sintético, 806 resultados) | N1/N2/N3 | "Tudo" (L1) | cópia com 565 linhas relabeladas | corte; AF1 | cortadas 136; AF1 273/281/252 | M1 136; A39 "Período inteiro do lote: N1 273 · N2 281 · N3 252 resultados." | 0 | APROVADO |
| WBC | XN1000 | 528211 (sintético) | N1/N2/N3 | 2026, "Ano inteiro" (L2) | idem | idem | cortadas 42; AF1 182/187/164 | M1 42; A39 "Período inteiro do lote: N1 182 · N2 187 · N3 164 resultados." | 0 | APROVADO |
| WBC | XN1000 | H3 vazio (cópia sintética) | N1/N2/N3 | 2026 (L3) | idem | idem | cortadas 277; AF1 374/400/340 | M1 277; A39 "Período inteiro (todos os lotes): …" | 0 | APROVADO (277 e não 278 porque a troca de rótulo juntou dois pares (RUN, LOTE) da corrida 26080701) |
| WBC | XN1000 | 528211 (sintético) | N1/N2/N3 | Out/2026 (L4) | idem | sem corrida no período | cortadas 0; AF1 vazio | M1 0; AF1 vazio; "Nenhuma corrida deste lote no período (com dados: 625311 (40))" | 0 | APROVADO |
| WBC | XN1000 | sintético (L1–L3) | N1/N2/N3 | n da janela | idem | um ponto por nível em cada corrida (ADR-080) | ex.: 142/153/132 | ex.: 141/152/131 | −1 por nível | DIFERENÇA EXPLICADA, não é defeito: a corrida 26080701 passou a ter dois resultados por nível no mesmo "lote" (dois lotes reais rodaram nela em paralelo) e o motor guarda um ponto por nível. Conferido ID por ID: nenhum ID só no motor; só no oráculo HEM-328296, HEM-328371 e HEM-328446. Com os dados reais o n bateu exatamente |

**Invariância Hema (entrega 41 × 42).** Nos 7 cenários de dados reais: `Painel!A6:Q12` 0 células diferentes;
`Calc!A1:AZ190` (origem do LJ) 0 células diferentes; `engPainel`, M1 e AF1 iguais; `Eng_Saida!A1:AF190` com uma
célula diferente por cenário, a G1 (carimbo `Now`), o que é esperado. Só o texto do A39 mudou, no trecho do corte.

### B.2 Bioquímica (origem: Bioquimica_F1.md e a regressão do Hematologia_F1.md)

| Analito | Equipamento | Lote | Nível | Período | Dados de entrada | Fórmula encontrada | Resultado independente | Resultado QC_INI (Value2) | Diferença | Status |
|---|---|---|---|---|---|---|---|---|---|---|
| Glicose (e todos os analitos da tabela) | DIMENSION 1 | 8974 e H3 vazio | N1/N2 | 9 cenários S1–S9 | Estatística A1:AU200 da entrega 41 (mesma tblCQ_Final) | invariância: mesmo motor, mesmos dados | valores da base 41 | valores da nova 42 | 0 em todos os números (n, média, DP, CV, bias, ET, Sigma, status). Só o D4 de exibição nos cenários "Tudo" (S1, S5, S9): 2026 → "" | APROVADO |
| (Painel) | DIMENSION 1 | 8974 e H3 vazio | N1/N2 | 9 cenários | `engPainel` da base 41 | invariância | idem | idem | 0 | APROVADO |
| Glicose | DIMENSION 1 | H3 vazio | N1/N2 | "Tudo" (B1, regressão do aviso) | oráculo do reteste 04 (não refeito) | corte; período inteiro | reteste 04: 163 + 1 corrida só com X vermelho; período 264/249 | A39 "⚠ 164 corrida(s) … Período inteiro: N1 264 · N2 249 — veja a aba Estatística." | 0 em relação ao reteste 04 | APROVADO (regressão de texto; o oráculo numérico não foi refeito neste ciclo) |
| Glicose | DIMENSION 1 | 8974 | N1/N2 | "Tudo" (B2) e Out/2026 (B3) | oráculo do reteste 04 (não refeito) | idem | reteste 04: 39 cortadas; período 174/161; Out sem corte | B2 "⚠ 39 corrida(s) … Período inteiro: N1 174 · N2 161 — veja a aba Estatística."; B3 sem aviso | 0 em relação ao reteste 04 | APROVADO (idem) |

**Cabeçalho da Estatística (só exibição, conferido junto).** S2 Set/2026: B5 01/09/2026, D5 30/09/2026, D4 2026;
S3 T3: 01/07–30/09/2026; S4 ano: 01/01–31/12/2026; S6 H3 vazio + Out/2026: 01/10–31/10/2026; S7 01/12/2025–31/01/2026
com D4 vazio. Todos iguais ao Painel. Observação da frente: as datas escritas pelo COM no G3/G4 em S7 e S8 entraram
com 03:00 (Value2 `…,125`) por causa do fuso; a entrada foi a mesma na base e na nova, então a comparação vale. Um
usuário que digita a data grava 00:00.

### B.3 Etapas para reproduzir

1. Copiar a produção atual e o backup da entrega 41 para `%TEMP%\qcw\auditoria\<ciclo>\` e conferir o sha256 (Bio
   b20fe00f7e51 e 8ec815ab653d; Hema 7bab8fc008f6 e aab0a54f464e).
2. Bio: rodar `_arquitetura/agentes/ferramentas/reteste05_usuario_f1.py`. Ele abre uma cópia por vez com o
   `qa_final.QA`, aplica S1–S9 pelas células de filtro do Painel, abre a Estatística e grava em `Bioquimica_F1.json`
   a leitura de A4:J5, das linhas 1–12 e de A1:AU200. Depois compara base × nova com tolerância zero.
3. Hema: rodar `reteste05_usuario_f1_hema.py` (nova, base e Bio) e `reteste05_usuario_f1_hema.py --lote-sintetico`.
   O oráculo lê a tblCQ_Final da cópia e conta resultados, pares (RUN, LOTE), janela de 180 e n por nível. Depois
   compara com M1, AF1, `engPainel` e o A39.

### B.4 Caminho do dado (conferido neste ciclo)

- **CIQ → janela → Painel/LJ:** tblCQ_Final (PARTICIPA = SIM) → corridas (RUN na Bio; par RUN × LOTE no modo TODOS da
  Hema, ADR-080) → janela das 180 mais recentes → `Eng_Saida` (M1 cortadas, AF1 n do período) → `engPainel` (n,
  indicadores) → `Calc` (origem do LJ) → A39 (texto do aviso).
- **Painel → Estatística da Bio:** Painel G3/G4/K3 → Estatística B5/D5/D4 (exibição) e H5/J5 (`JanelaFim`). O motor
  lê B5/D5 com B4 = PERSONALIZADO (ADR-095 §3); o D4 é só exibição.
- **CEQ → bias → ET → Sigma:** não exercitado neste ciclo. Os números que dependem dele ficaram iguais bit a bit aos da
  entrega 41.

### B.5 Diferenças de método e observações (não reprovadas)

- **n da janela 1 abaixo por nível no dado sintético:** artefato da troca de rótulo (B.1, última linha). Com os dados
  reais o n bate.
- **M1 da Bio com H3 vazio (164 contra 163):** já explicado no reteste 04 (`mEstatistica` D10, uma corrida só com X
  vermelho). Os n batem.

---

## Relatório C: ocorrências

**Deduplicação.** As frentes trouxeram 1 ocorrência nova (R5-01, Bio F1). Ficou 1. O ID P-xx é provisório; o ID
definitivo (OC-xxxx) é do Orquestrador. Nenhuma ocorrência é Crítica, Alta ou Média. A Hema não trouxe ocorrência nova.
A observação "Lote 522611 · alvo …" com o H3 vazio já é a **OC-0039 (ABERTA)** e não foi registrada de novo.

### C.1 Tabela

| ID | Produto | Prioridade | Tipo | Tela/módulo | Resumo | Origem | ID anterior | Outro produto |
|---|---|---|---|---|---|---|---|---|
| P-01 | Bioquímica | Baixo | `usabilidade` | aba Estatística, cabeçalho "PERÍODO ANALISADO", J5 (`JanelaFim`, formato dd/mm/aaaa) | Com só o De preenchido no Painel (Até vazio), o J5 mostra "00/01/1900" | Bio F1 R5-01 | — (vizinho da OC-0071; o defeito já existia na entrega 41) | não se aplica (a Estatística da Hema tem filtro próprio) |

### C.2 Detalhe

**P-01 (Bio, Baixo, usabilidade) — J5 "00/01/1900" com só o De**
- **Tela/módulo:** aba Estatística, cabeçalho "PERÍODO ANALISADO", J5:
  `=IF(AND(Painel!$G$3="",Painel!$G$4=""),"",JanelaFim($B$4,$D$4,$F$4,$B$5,$D$5))`, formato `dd/mm/aaaa`.
- **Ação:** no Painel, em Set/2026, apagar o Até (G4) e deixar só o De 01/09/2026. O K4 vira "Personalizado". Abrir a
  Estatística.
- **Esperado:** J5 vazio, ou o fim real da janela (por exemplo, o último dado). Nunca uma data sem sentido.
- **Observado:** D5 vazio (o `;;` do ADR-095 resolveu esse), H5 01/09/2026 e **J5 "00/01/1900"** (Value2 0). Na
  entrega 41, D5 e J5 mostravam "00/01/1900". A condição `SEM_DEATE` só trata De e Até vazios ao mesmo tempo.
- **Dificuldade:** quem confere ou imprime a Estatística lê "até 00/01/1900".
- **Consequência:** só de leitura. Não se sabe até quando os dados foram considerados. Os números não mudam (0
  diferenças com a base no S8).
- **Sugestão:** aplicar ao J5 (e ao H5, para o caso "só o Até") o formato `dd/mm/aaaa;;` ou uma condição por célula
  (`Painel!$G$4=""`), sem mexer no `JanelaFim`. Outra opção é o Painel avisar quando só uma das datas estiver
  preenchida.
- **Evidência:** `%TEMP%\qcw\auditoria\2026-11-15_reteste05\Bioquimica_F1.json`: `NOVA.cen.S8_so_De.hdr.J5` (texto
  "00/01/1900", Value2 0.0, formato `dd/mm/aaaa`) e `BASE.cen.S8_so_De.hdr` (D5 e J5 "00/01/1900"). Script
  `_arquitetura/agentes/ferramentas/reteste05_usuario_f1.py`. Produção b20fe00f7e51.
- **Não executado:** o caso simétrico (só o Até preenchido).

### C.3 Dúvidas de regra de negócio (para o Orquestrador encaminhar)

| ID | Produto | Para | Dúvida |
|---|---|---|---|
| D-01 | Bio | Orquestrador / UX (e Tester, para confirmar no código) | Com só o De preenchido no Painel, o motor considera "do De até o último dado"? O que o cabeçalho (J5) deve mostrar nesse caso: vazio, "até o último dado" ou a data do último resultado? (ligada ao P-01) |
| D-02 | Hema | Orquestrador | O ramo "Período inteiro do lote" do aviso do corte foi validado **só com dado sintético**: nenhum lote real passa de 89 corridas. Aceitar essa validação, ou refazer com dado real quando um lote passar de 180 corridas (cerca de 2 meses a 3 corridas por dia)? |
| D-03 | Hema | Tester/QA | No cenário sintético L4 (lote sem corrida no período), o `Eng_Saida!A1:BJ200` ainda tinha IDs HEM- da publicação anterior, embora M1, n e `Calc!D3:D182` estivessem zerados e o A39 dissesse "Nenhuma corrida deste lote no período". Essas colunas chegam a algum gráfico? O LJ mostra só as linhas do alvo nesse caso (ADR-078)? Não investigado |
| D-04 | Hema | UX | Continua a D-14 do reteste 04: no modo TODOS, o Painel dá REJEITADO sem mostrar Média, DP e CV (ADR-086). Explicar de onde vem o status? |

---

## Relatório D: parecer final

Critério: "validado" só onde houve teste executado e aprovado neste ciclo. O resto é "não validado" (teste reprovado)
ou "não executado" (com o motivo). Este ciclo foi um reteste de 2 itens, só com a F1.

### D.1 A rotina funciona?

**Sim, no que foi refeito.**
- **Validado:**
  - OC-0071 (Bio): cabeçalho da Estatística coerente em "Tudo", mês, trimestre, ano, H3 vazio e De/Até atravessando o
    ano.
  - OC-0012 (Hema): o aviso do corte diz o recorte e o n do período inteiro sem apontar para a Estatística, com os
    números iguais ao oráculo.
  - Regressão na Bio: o aviso não mudou.
  - Os números da entrega 42 são iguais bit a bit aos da entrega 41 em todos os cenários (23), nos dois produtos.
- **Não validado:** o cabeçalho da Estatística da Bio com só o De preenchido (P-01, Baixo, defeito antigo; os números
  estão certos).
- **Não executado:** o resto da rotina (CEQ, entrada manual e automática, LJ, Westgard). Ficou fora da lista deste
  reteste.

### D.2 Bias, ET e Sigma estão validados?

**Não executado neste ciclo** (não houve F3 nem F2). O que foi executado e aprovado é a **invariância**: na Bio, o
bias, o ET, o Sigma, o n, a média, o DP, o CV e os status da Estatística A1:AU200 são idênticos bit a bit aos da
entrega 41 em 9 cenários. A última validação independente de bias, ET e Sigma é a do reteste 04 (47 cálculos e
32.640 comparações, 0 divergências), feita sobre a entrega 41. Na Hema, a frente comparou o Painel, o Calc e o
Eng_Saida, mas não a Estatística (A1:AU…). Lá, a invariância do bias, ET e Sigma da Estatística não foi medida neste
ciclo.

### D.3 A integração CIQ + CEQ está validada?

**Não executado neste ciclo.** Nenhum CEQ foi incluído nem corrigido. Só se pode dizer que, sem mudança de dados, os
números que combinam CIQ e CEQ na Estatística da Bio não mudaram com a entrega 42.

### D.4 Os filtros de período e de lote estão corretos?

**Validado nos casos executados:**
- **Bio:** "Tudo", Set/2026, T3/2026, "Ano inteiro" 2026, De/Até atravessando o ano, H3 vazio com "Tudo" e com
  Out/2026, e a volta a "Tudo" depois do personalizado. Cabeçalho igual ao Painel e números iguais aos da entrega 41.
- **Hema:** WBC e HGB com H3 vazio em "Tudo", "Ano inteiro", De/Até 01/01–31/10/2026 e Out/2026; lote 528211 em
  "Tudo" e com o ajuste automático do período. A janela de 180 corridas, o n da janela e o n do período batem com o
  oráculo (com o par (RUN, LOTE) do ADR-080). O ramo do lote foi validado só com dado sintético (D-02).

**Não validado:** o cabeçalho com só o De (P-01, exibição).

**Não executado neste ciclo:**
- o caso "só o Até";
- a troca de equipamento na Bio;
- a F4 completa: pontos e limites do LJ, Westgard e corrida do lote. O `Calc!A1:AZ190` (origem do LJ) da Hema é igual
  ao da entrega 41, mas o gráfico em si não foi conferido.

### D.5 A entrada manual e a automática estão validadas?

**Não executado neste ciclo.** Nenhum resultado manual foi digitado e nenhum ATUALIZAR foi rodado. A tblCQ_Final das
cópias é a mesma da produção (último resultado da Hema em 13/11/2026 13:10:20). A última validação é a do reteste 04,
com as listas de validação quebradas (P-01/P-02 daquele ciclo, tratadas no ADR-095 §1 como OC-0068). O reteste da
OC-0068 não fez parte desta lista.

### D.6 Há divergência matemática?

**Não**, nos cálculos executados.
- 16 linhas do Relatório B: 15 aprovadas com diferença 0, e 1 com diferença explicada (−1 por nível no n da janela,
  só no dado sintético, por um artefato da troca de rótulo, conferido ID por ID).
- 0 diferenças numéricas entre a entrega 41 e a 42 nos 23 cenários, com tolerância zero.
- Bias, ET e Sigma não foram recalculados de forma independente neste ciclo (D.2).

### D.7 Há problema operacional?

**Não**, nada que impeça a rotina neste escopo. Fica só o P-01 (Baixo), de exibição. As produções ficaram intactas:
mesmo hash antes e depois de todas as execuções.

Na automação de teste (não é defeito do QC_INI), duas rodadas do lote sintético da Hema pararam antes de concluir (aba
protegida e falta de diagnóstico). Foram refeitas e a terceira rodada (PID 18360) é a válida.

### D.8 O que fica pendente

1. Decidir e corrigir o **P-01** (Baixo) e responder a D-01. No reteste, cobrir também o caso "só o Até".
2. Encaminhar D-02 (Orquestrador), D-03 (Tester/QA) e D-04 (UX). A OC-0039 continua aberta.
3. Seguem abertas, do ADR-095 §6: **OC-0069** (aviso do CEQ com o Controllab mapeado) e **OC-0070** (Painel em lote com
   regra da tblLotesMaterial).
4. Fora da lista deste reteste, e por isso sem estado neste parecer: OC-0068 (listas de escolha), OC-0021 (acentos do
   DETALHE) e OC-0072 (ano literal nos formatos). Da OC-0072, só se viu de carona que não há "yyyy" nas linhas 1–12 da
   Estatística da Bio.
5. **Não executados, a cobrir num próximo ciclo:**
   - F3 sobre a entrega 42 (bias, ET e Sigma, ETp ausente na Bio, ETp VB/FAB);
   - F2 (CEQ: inclusão, correção, exclusão, seleção por provedor, ano e rodada);
   - F1 completa (manual, recebimento de resultado novo do DB_SEAC, equivalência manual × automático);
   - F4 completa (LJ, Westgard, corrida do lote, equipamentos);
   - os pendentes do reteste 04 que continuam sem caso real (regra MATRIZ na Hema, código ambíguo do ADR-093), o
     clique Sim/Não do frmCEQ e a digitação física com validação.

Dados pessoais: nenhum dado de usuário, login ou máquina consta deste parecer nem dos relatórios das frentes.
