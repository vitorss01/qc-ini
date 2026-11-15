# Reteste 05: Usuário Simulado, Hematologia, frente F1

Data: 15/11/2026 · Agente: AGENTE_USUARIO_SIMULADO · Modo: RETESTE (1 ocorrência: OC-0012) · Produção: entrega 42.

## Ambiente e método

| Item | Valor |
|---|---|
| Produção atual da Hema (só copiada) | `QC_Hematologia.xlsm` sha256 curto **7bab8fc008f6**, igual antes e depois |
| Base de comparação (antes da entrega 42) | `_backup_pre_integracao_2026-11-15_0458/QC_Hematologia.xlsm` sha256 curto **aab0a54f464e**. É a entrega 41, a mesma que o reteste 04 testou. Igual antes e depois |
| Produção atual da Bio (regressão) | `QC_Bioquimica.xlsm` sha256 curto **b20fe00f7e51**, igual antes e depois |
| Cópias (novas, nenhuma salva, apagadas no fim) | `%TEMP%\qcw\auditoria\2026-11-15_reteste05\Hematologia_F1_nova.xlsm`, `Hematologia_F1_base41.xlsm`, `Bioquimica_F1_bio.xlsm` e `Hematologia_F1_lote.xlsm` (lote sintético, ver abaixo) |
| Excel | Um por vez, cada um aberto e encerrado pelo `qa_final.QA`: PIDs 11036 (nova), 9732 (base), 4032 (Bio), 19288 e 30144 (lote sintético; essas duas rodadas pararam antes de concluir, por aba protegida e por falta de diagnóstico) e 18360 (lote sintético, a rodada válida). Não toquei no Excel 13084, que já estava aberto (tarefa agendada), nem no 4984, que apareceu no fim e não é meu. Nenhuma captura de tela |
| Script | `_arquitetura/agentes/ferramentas/reteste05_usuario_f1_hema.py`, documentado no cabeçalho. Modos: o padrão (nova, base e Bio) e `--lote-sintetico` |
| Saídas (fora do repositório) | `%TEMP%\qcw\auditoria\2026-11-15_reteste05\Hematologia_F1{,_lote}.{json,log}` |
| tblCQ_Final | Hema: **a mesma** na nova e na base (48.391 linhas, 47.709 participam, hash do Value2 `ee71bc38576d73a3` nas duas, último resultado 13/11/2026 13:10:20). Bio: 35.906 linhas |
| Dados de teste | Nenhum nas cópias nova, base e Bio (só filtros). Na cópia do lote sintético, 565 linhas de WBC trocaram de LOTE (ver abaixo) |

**Caminho usado.** É o do usuário: as células de filtro do Painel (C3, H3, K3, K4, G3/G4; na Bio também o L4), escritas
com os eventos ligados. O De/Até foi escrito como número de série (Value2), para não herdar o fuso do COM. Em cada
cenário li:
- do Painel: o A39 (valor e fórmula) e A6:Q12;
- do motor: `Eng_Saida!M1` (corridas cortadas), `AF1` (n do período inteiro), `engPainel` e `Eng_Saida!A1:AF190`;
- a origem do Levey-Jennings: `Calc!A1:AZ190`.

A comparação base × nova foi feita célula a célula, na precisão total do `Value2`, com tolerância zero, porque é o
mesmo motor sobre a mesma tblCQ_Final.

**Oráculo próprio** (Python puro, sobre a tblCQ_Final lida da cópia):
- conta os resultados com PARTICIPA_ESTATISTICA = SIM e RESULTADO numérico, do analito (e do lote, se houver), com data
  dentro de [De, Até];
- AF1 esperado = "N1 a · N2 b · N3 c";
- corridas = pares (RUN, LOTE) distintos (ADR-080, como conferi no reteste 04);
- janela = as 180 corridas mais recentes do período;
- cortadas = total − 180, e o n da janela é contado por nível.

**O ramo "do lote" exigiu sair do caminho do usuário.** Nenhum lote real da Hema tem mais de 180 corridas: o maior,
528211, tem 89 (WBC, RET-HE, RBC, PLT…). Por isso o aviso com lote no H3 **não aparece com os dados reais**. Para
exercitar o ramo, fiz o seguinte numa cópia nova da produção, descartada no fim:
- dei às 565 linhas de WBC dos lotes 533811, 614011 e 619611 o LOTE "528211", com o mesmo tipo texto, na tblCQ_Final e
  com os eventos desligados (dado sintético de teste);
- chamei `mEstatistica.InvalidarCache`, o único passo fora do caminho do usuário;
- filtrei o Painel como o usuário.

## Resultado por ocorrência

| OC | Estado | Evidência |
|---|---|---|
| OC-0012 | **RESOLVIDA** | Detalhe abaixo |

### Fluxo original: WBC, H3 vazio (todos os lotes), K3 2026 + "Ano inteiro"

| Item | Base (entrega 41) | Nova (entrega 42) | Oráculo |
|---|---|---|---|
| A39 (trecho do corte) | "⚠ 278 corrida(s) antigas do período fora do gráfico e dos indicadores do Painel (máx. 180). Período inteiro: N1 374 · N2 400 · N3 340 — veja a aba Estatística." | "⚠ 278 corrida(s) antigas do período fora do gráfico e dos indicadores do Painel (máx. 180). **Período inteiro (todos os lotes): N1 374 · N2 400 · N3 340 resultados.**" | — |
| Cita a aba Estatística? | sim | **não** (nem o texto nem a fórmula do A39) | — |
| `Eng_Saida!M1` | 278 | 278 | 278 (458 corridas (RUN, LOTE); por RUN só, seriam 452) |
| `Eng_Saida!AF1` | "N1 374 · N2 400 · N3 340" | igual | "N1 374 · N2 400 · N3 340" |
| n do Painel (engPainel) | 120 / 158 / 112 | igual | janela 120 / 158 / 112 |

O A39 inteiro da nova diz: "Lote 522611 · alvo N1 3,0038 ± 0,05290 · N2 6,8052 ± 0,09120 · N3 16,5238 ± 0,16930. ⚠ 278
corrida(s) … Período inteiro (todos os lotes): N1 374 · N2 400 · N3 340 resultados. Filtro 01/01/26 a 31/12/26".

### Fluxo original, 2ª parte: um lote no H3

O lote de WBC com mais resultados é o 528211 (250 que participam). Coloquei-o no H3 com o K3 2026 + "Ano inteiro".
O Painel ajustou sozinho para "Tudo" e avisou: "FILTRO: Lote 528211: o filtro anterior não mostrava nenhum resultado
dele -- ajustado: período Tudo…". Em "Tudo", o lote tem 89 corridas: M1 = 0, sem aviso de corte, e n 83/86/81, igual ao
oráculo. É o comportamento certo, e base e nova ficaram idênticas.

**Ramo "do lote" (dado sintético, cópia `Hematologia_F1_lote.xlsm`).** O WBC do 528211 ficou com 806 resultados que
participam.

| Cenário | A39 (trecho do corte) | M1 = oráculo | AF1 = oráculo |
|---|---|---|---|
| L1: H3 528211 + "Tudo" | "⚠ 136 corrida(s) antigas do período fora do gráfico e dos indicadores do Painel (máx. 180). **Período inteiro do lote: N1 273 · N2 281 · N3 252 resultados.**" | 136 = 136 | sim |
| L2: + K3 2026 "Ano inteiro" | "⚠ 42 corrida(s) … **Período inteiro do lote: N1 182 · N2 187 · N3 164 resultados.**" | 42 = 42 | sim |
| L3: H3 vazio (todos), 2026 | "⚠ 277 corrida(s) … **Período inteiro (todos os lotes): N1 374 · N2 400 · N3 340 resultados.**" | 277 = 277 | sim |
| L4: H3 528211 + Out/2026 | sem aviso de corte ("Nenhuma corrida deste lote no período (com dados: 625311 (40))") | 0 = 0 | AF1 vazio, como esperado |

Nenhum dos A39 cita a aba Estatística.

**Diferença de 1 por nível, só nos cenários sintéticos.** Nos três, o n da janela do Painel ficou 1 abaixo do oráculo
por nível: 141/152/131 contra 142/153/132. Conferi a causa comparando, ID por ID, os resultados que o motor publicou
no Eng_Saida com os da janela do oráculo:
- não há nenhum ID só no motor;
- ficam só no oráculo HEM-328296, HEM-328371 e HEM-328446, da corrida 26080701;
- com a troca de rótulo, essa corrida passou a ter dois resultados por nível no mesmo "lote", porque dois lotes
  reais rodaram nela em paralelo;
- o motor põe um ponto por nível em cada corrida (ADR-080) e descarta o repetido.

É um artefato do dado sintético, que não existe nos dados reais, e não é defeito. Com os dados reais (H1, H4, H5, H7),
o n da janela bateu exatamente com o oráculo.

### Vizinhos (dados reais, nova × base)

| Cenário | M1 | AF1 (= oráculo) | n do Painel (= oráculo) | A39 nova |
|---|---|---|---|---|
| H4: WBC, todos, "Tudo (sem filtro)" | 431 | N1 515 · N2 545 · N3 479 | 120/158/112 | "Período inteiro (todos os lotes): N1 515 · N2 545 · N3 479 resultados." |
| H5: WBC, todos, De/Até 01/01–31/10/2026 (K4 "Personalizado") | 251 | N1 357 · N2 377 · N3 328 | 125/157/116 | "Período inteiro (todos os lotes): N1 357 · N2 377 · N3 328 resultados." |
| H6: WBC, todos, Out/2026 (sem corte) | 0 | vazio | 21/39/21 = n do período | sem aviso de corte |
| H7: HGB, todos, De/Até 01/01–31/10/2026 | 251 | N1 357 · N2 377 · N3 328 | 125/157/116 | idem, sem citar a Estatística |

**Indicadores e gráfico inalterados.** Nos 7 cenários de dados reais, a base (entrega 41) e a nova (entrega 42)
ficaram iguais nestes pontos:
- `Painel!A6:Q12`: 0 células diferentes;
- `Calc!A1:AZ190`, a origem do Levey-Jennings: 0 células diferentes;
- `engPainel`: igual;
- `M1` e `AF1`: iguais;
- `Eng_Saida!A1:AF190`: uma célula diferente por cenário, a G1 (carimbo `Now` da publicação). Isso é esperado.

O único texto que mudou foi o do A39, no trecho do corte.

**Regressão na Bio** (cópia da produção atual, DIMENSION 1, Glicose):
- B1, H3 vazio + "Tudo": "⚠ 164 corrida(s) … Período inteiro: N1 264 · N2 249 — veja a aba Estatística.";
- B2, lote 8974 + "Tudo": "⚠ 39 corrida(s) … Período inteiro: N1 174 · N2 161 — veja a aba Estatística.". São os
  mesmos 174/161 que o reteste 04 achou na Estatística da Bio;
- B3, lote 8974 + Out/2026: sem aviso de corte.

A fórmula do A39 da Bio continua com "— veja a aba Estatística". A Bio não mudou, como o ADR-095 §4 manda. Não refiz
o oráculo na Bio, porque o filtro por equipamento não está no meu oráculo: a regressão pedida é de texto.

**Critério de aceite.** Todos os itens foram atendidos:
- texto com H3 vazio;
- texto com lote, só com dado sintético, porque não há lote real com mais de 180 corridas;
- sem citar a Estatística;
- contagens = AF1 = dados de origem;
- indicadores e gráfico inalterados;
- Bio inalterada.

## Ocorrências novas

Nenhuma.

## Dúvidas de regra de negócio

1. **Ramo "do lote" sem dado real.** Hoje nenhum lote da Hema passa de 89 corridas, então o texto "Período inteiro do
   lote" só aparece com dado sintético. Ele vai aparecer quando um lote durar mais de 180 corridas (cerca de 2 meses
   a 3 corridas por dia). Fica registrado que esse ramo foi validado só com dado sintético.
2. **"Lote 522611 · alvo …" no início do A39 com H3 vazio (todos os lotes).** O mesmo aviso diz "(todos os lotes)" no
   fim e mostra o alvo do lote em uso no começo. Isso já é a OC-0039 (ABERTA). Não registrei de novo.
3. **Observação não investigada (fora do escopo):** no cenário sintético L4 (lote sem corrida no período), o
   `Eng_Saida!A1:BJ200` ainda tinha IDs HEM- da publicação anterior, embora M1, n e `Calc!D3:D182` estivessem zerados
   e o A39 dissesse "Nenhuma corrida deste lote no período". Não confirmei se essas colunas chegam a algum gráfico.
   Fica como pergunta ao Tester, que pode conferir se o LJ mostra só as linhas do alvo nesse caso (ADR-078).

## Impressões de uso (Relatório A)

- **Aviso do corte.** O aviso agora é autossuficiente. Diz quantas corridas ficaram fora, que o corte vale para o
  gráfico **e** para os indicadores, e quantos resultados o período inteiro tem, por nível, e de qual recorte (todos
  os lotes ou o lote). Já não manda o usuário a uma aba que mostra outro ano e outro lote. Os números conferem com a
  base de dados.
- **Tamanho do A39.** O texto tem cerca de 275 caracteres numa área mesclada A39:U40, com quebra de linha (33 pt de
  altura). Não vi a tela, que não pode ser capturada. Pelo tamanho, cabe em duas linhas.
- **Troca de lote com período sem dados.** Ao escolher um lote antigo com K3 2026, o Painel mudou sozinho para "Tudo"
  e disse por quê. Isso ajuda.
- **REJEITADO com o Painel sem Média, DP e CV.** No modo TODOS, os três níveis continuam REJEITADO, inclusive em
  Out/2026 (n 21/39/21), com o Painel sem Média, DP e CV (ADR-086). A dúvida de onde vem o REJEITADO, anotada no
  reteste 04, continua para quem não conhece o ADR.

## Arquivos

- Script: `_arquitetura/agentes/ferramentas/reteste05_usuario_f1_hema.py`.
- Saídas (fora do repositório): `%TEMP%\qcw\auditoria\2026-11-15_reteste05\Hematologia_F1.json/.log` e
  `Hematologia_F1_lote.json/.log`.
