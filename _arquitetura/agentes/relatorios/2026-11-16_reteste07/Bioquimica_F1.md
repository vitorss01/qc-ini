# Reteste 07: Usuário Simulado, Bioquímica, frente F1

Data: 16/11/2026 · Agente: AGENTE_USUARIO_SIMULADO · Modo: RETESTE da OC-0078 · Produção: entrega 44 (ADR-097).

## Ambiente e método

| Item | Valor |
|---|---|
| Produção atual (só copiada) | `QC_Bioquimica.xlsm` sha256 curto **fdc173906fff**, igual antes e depois (`origem_intacta.igual = true`) |
| Base de comparação ("antes") | `_backup_pre_integracao_2026-11-16_0729/QC_Bioquimica.xlsm` sha256 curto **bd5314426ab8** (entrega 43), igual antes e depois |
| Cópias (novas, nenhuma salva, apagadas no fim) | `%TEMP%\qcw\auditoria\2026-11-16_reteste07\Bioquimica_F1_nova.xlsm` e `Bioquimica_F1_base.xlsm` |
| Excel | Um por vez, abertos e encerrados pelo `qa_final.QA`: PID 30228 (nova) e 12976 (base). Antes de cada execução não havia nenhum outro EXCEL.EXE; nenhum outro processo foi tocado. Nenhuma captura de tela |
| Script (novo, documentado no cabeçalho) | `_arquitetura/agentes/ferramentas/reteste07_usuario_f1.py <nova\|base>` (reaproveita os auxiliares do `reteste06_usuario_f1.py`) |
| Saídas (fora do repositório) | `Bioquimica_F1_nova.json`, `Bioquimica_F1_base.json` e os `.log`, na pasta do ciclo no `%TEMP%` |
| Usuários (só na cópia, ANALISTA, login por `mSeguranca.DoLogin`) | U1 = `TESTE_AUDITORIA_R7` (digitou os originais e **apagou**); U2 = `TESTE_AUDITORIA_R7B` (digitou nas linhas apagadas, recortou, fez o Ctrl+D, editou, substituiu) |
| Dados de teste | Glicose, DIMENSION 1, lote 8977, 16/11/2026, MOTIVO "Outro". Valores sintéticos de 3 decimais calculados do alvo do lote (`mLotes.MediaDoLote`/`DPDoLote`: N1 79,3140 ± 1,5428; N2 274,7759 ± 5,0270), com \|z\| ≤ 0,3 para não disparar Westgard; o manual que substitui vale o real + 0,123 (77,123, z = −1,42). Tabela manual vazia no início (a produção não tem manual) |
| tblCQ_Final | 36.056 linhas no início. O 1º ATUALIZAR leu o DB_SEAC só para leitura e trouxe 2 resultados novos, de outros analitos/filtros (a conta à mão abaixo confirma que nenhum entrou no filtro) |

**Caminho do usuário:**
- Login; Painel com L4, C3, H3, K3 e K4 escritos com os eventos ligados.
- Entrada pelo botão (`mIntegracao.IrDigitarResultados`) e digitação célula a célula (`FormulaLocal`), com
  `Application.Interactive = True`, para que os avisos do `EntradaMudou` aparecessem. As caixas foram lidas pelo texto
  da janela do meu PID e fechadas no OK.
- Delete = `ClearContents` na seleção (ID…MOTIVO; DATA…MOTIVO; duas linhas de uma vez); recorte = `Range.Cut(destino)`;
  Ctrl+D = `Range.FillDown()` em ID…MOTIVO de 3 linhas, a de cima carimbada.
- ATUALIZAR DADOS pelo núcleo do botão (`AtualizarDadosAutomatico`), duas vezes.

**O que não executei:** o teclado e o mouse físicos (Delete, Ctrl+X/Ctrl+V e Ctrl+D de verdade). O cenário "dias
depois" foi reproduzido com 2,5 min entre o carimbo antigo e a nova digitação, como no reteste 06; o mecanismo
(DATA/HORA digitada posterior ao carimbo herdado) é o mesmo. Não refiz o "apagar só o ID de uma linha com dado"
(C04 do `qa_adr097`).

**Oráculo próprio:** Python puro (`statistics`) sobre o `Value2` da tblCQ_Final: PARTICIPA = SIM, Glicose,
DIMENSION 1, lote 8977, nível, DATA_HORA em nov/2026 → n, média, DP amostral, CV. E a conta à mão "antes − retirados −
substituído + manuais ATIVOS".

## Roteiro executado (mesmo roteiro nas duas cópias)

| Fase | Quem | O que |
|---|---|---|
| 1 | U1 | Digita A (N1 03:01), D (N2 03:02), M (N2 03:03), G (N1 03:04), F (N1 03:05) nas linhas 1–5 e o controle K (N2 03:08) na 8. ATUALIZAR #1 |
| 2a | U1 | Apaga A (Delete de **ID…MOTIVO**) e D (Delete só de **DATA…MOTIVO**, o ID fora da seleção) |
| 2b | U2 | 2,5 min depois do carimbo de A: botão Digitar Resultados; digita B (N1) onde o cursor parou e E (N2) na linha de D, as duas com DATA/HORA = carimbo antigo + 1 min (o caso que era recusado) |
| 2c | U2 | Recorta ID…MOTIVO de M (já integrado) para a linha 12; digita N (N2) na origem |
| 2d | U2 | Ctrl+D de F (linha 5) para as linhas 6 e 7; troca HORA e RESULTADO das duas de baixo |
| 2e | U2 | Corrige o RESULTADO de G (manual de U1, já integrado) na própria linha |
| 2f | U2 | ADR-082: digita só o número 379301 (BIO-379301, N1 real, 77,0) no SUBSTITUI_ID, depois RESULTADO 77,123 e MOTIVO. Vizinho: outra linha com SUBSTITUI_ID 379339 (N2) e Delete de ID…MOTIVO |
| 2g | U2 | Vizinhos: duas linhas digitadas e apagadas de uma vez; carimbo órfão legado (gravado com os eventos desligados) numa linha vazia |
| — | U2 | ATUALIZAR #2 |

## Resultado

| OC | Estado | Resumo |
|---|---|---|
| OC-0078 | **RESOLVIDA** (com uma ressalva no critério 3) | O Delete de ID…MOTIVO e o de DATA…MOTIVO limpam USUARIO e REGISTRADO_EM na hora (e o ID que sobrava, no segundo caso). O manual digitado depois na mesma linha, por outro usuário, ganha ID novo e o carimbo de quem digitou, agora, e fica ATIVO, entrando na Estatística, no Painel e no LJ (n e média exatos com o oráculo). A origem do recorte fica limpa e o manual digitado nela fica ATIVO. O Ctrl+D não muda a origem. Editar um manual mantém ID e carimbo. O ADR-082 funciona. Na base (entrega 43), o mesmo roteiro recusa B, E e N com "DATA/HORA posterior ao lançamento" e reaproveita o ID MAN_0002. **Ressalva:** a segunda metade do critério 3 ("o carimbo vai com a linha recortada") **não** é atendida. A linha movida fica com o carimbo de quem moveu, e o USUARIO_MANUAL do resultado já integrado muda sem registro no Audit_Log. O ADR-097 §5 declara isso como limite conhecido, e o comportamento é idêntico na base: não é efeito colateral. Ver R7-01 |

## Testes executados

| ID | Teste | Estado | Evidência (nova = entrega 44; base = entrega 43) |
|---|---|---|---|
| OC-0078 | Reteste do fluxo original e dos 6 critérios de aceite | **RESOLVIDA** | T01 a T08 abaixo; ressalva do critério 3 em T03b/R7-01 |
| T01 | Critério 1: Delete de ID…MOTIVO e manual novo na mesma linha | APROVADO | **Nova:** logo após o Delete, a linha 1 (MAN_0001, U1, 10:11:45) fica com ID, USUARIO e REGISTRADO_EM **vazios**. O botão Digitar Resultados leva a **B7 = linha 1**. U2 digita B (16/11 10:12, 79,083) às 10:14:15 → **MAN_0007**, USUARIO **TESTE_AUDITORIA_R7B**, REGISTRADO_EM **10:14:15**. ATUALIZAR #2: **ATIVO**, RUN 26111609, DETALHE vazio, USUARIO_MANUAL R7B, Audit "manual lancado por TESTE_AUDITORIA_R7B em 16/11/2026 10:14:15". MAN_0001 sai da tblCQ_Final (Audit RESULTADO_MANUAL_RETIRADO). LJ: Eng_Saida slot 134, N1 = MAN_0007 · 79,083. **Base:** a linha fica com U1/10:18:42; B herda isso e fica **MANUAL_INCOMPLETO "DATA/HORA posterior ao lançamento"** |
| T02 | Critério 2: Delete só de DATA…MOTIVO (o ID fica fora da seleção) | APROVADO | **Nova:** a linha 2 (MAN_0002) perde também o ID e o carimbo na hora. E (16/11 10:12, N2 275,379) digitado por U2 às 10:14:21 → **MAN_0008** (ID novo, não reaproveitado), R7B/10:14:21 → **ATIVO** RUN 26111609; MAN_0002 RETIRADO. **Base:** o ID MAN_0002 fica na linha; E é gravado com o **mesmo ID MAN_0002** e o carimbo de U1 e fica MANUAL_INCOMPLETO "posterior ao lançamento" (e o resultado antigo MAN_0002 é trocado pelo novo sob o mesmo ID) |
| T03a | Critério 3 (origem): recortar a linha e digitar outro manual na origem | APROVADO | **Nova:** logo após o recorte de MAN_0003 (linha 3 → 12), a origem fica com ID, USUARIO e REGISTRADO_EM vazios. N (16/11 10:13, N2 273,67), digitado por U2 às 10:14:34 → **MAN_0009**, R7B/10:14:34 → **ATIVO** RUN 26111610. **Base:** a origem fica com U1/10:18:55; N herda isso e fica MANUAL_INCOMPLETO "posterior ao lançamento" |
| T03b | Critério 3 (destino): o ID e o carimbo vão com a linha recortada | **REPROVADO** (só o carimbo; o ID vai) | **Nova:** o destino (linha 12) fica com **MAN_0003** (o ID vai, sem aviso), mas o carimbo passa a ser **TESTE_AUDITORIA_R7B / 10:14:29** (quem moveu, na hora do recorte), e não o original (R7, 10:11:57). Depois do ATUALIZAR #2, MAN_0003 continua ATIVO com o mesmo valor, mas o **USUARIO_MANUAL na tblCQ_Final muda de R7 para R7B**, e o Audit_Log **não** registra nada sobre o MAN_0003. **Base:** idêntico (destino R7B/10:21:26; USUARIO_MANUAL R7B). É o limite conhecido do ADR-097 §5: não é efeito colateral da entrega 44. Ver R7-01 |
| T04 | Critério 4: Ctrl+D a partir de uma linha carimbada | APROVADO | **Nova:** a origem (linha 5) continua **MAN_0005 / R7 / 10:12:08**. As de baixo ganham **MAN_0010 e MAN_0011** com R7B/10:14:44, e aparece **uma** caixa: "As 2 linhas coladas traziam IDs que já existem em outras linhas e receberam IDs novos (MAN_0010 a MAN_0011). … Atenção: com a mesma data, hora, nível e analito de outra linha, ela fica como repetida (CONFLITO_MANUAL), mesmo com outro resultado. Se for uma nova dosagem, use a hora real dela." Trocadas a HORA (03:21/03:22) e o RESULTADO, as três ficam ATIVO (RUN 26111602/03/04), e MAN_0005 mantém USUARIO_MANUAL R7. A base se comporta igual neste ponto |
| T05 | Critério 5: editar o resultado de um manual existente na própria linha | APROVADO | **Nova:** MAN_0004 (U1, 10:12:02, 79,005, já ATIVO) corrigido por U2 para 78,851: o ID e o carimbo **não mudam** (R7/10:12:02), e a tela continua ATIVO. ATUALIZAR #2: ATIVO, 78,851, USUARIO_MANUAL R7, Audit **RESULTADO_MANUAL_ALTERADO** 79,005 → 78,851. LJ: slot 126, N1 = MAN_0004 · 78,851 |
| T06 | Critério 6: ADR-082 (SUBSTITUI_ID) continua funcionando | APROVADO | **Nova:** digitando só "379301" no SUBSTITUI_ID de uma linha vazia, ele vira **BIO-379301**, e a linha se preenche sozinha (DIMENSION 1, SORO, 8977, N1, Glicose, 16/11 09:03), com ID MAN_0012 e carimbo R7B/10:15:15. O SUBSTITUI_ID **não** se perde ao digitar RESULTADO e MOTIVO. ATUALIZAR #2: MAN_0012 **ATIVO**, DETALHE "substitui o resultado do interfaceamento BIO-379301"; BIO-379301 **SUBSTITUIDO_MANUAL**, PARTICIPA NÃO, sem RUN e fora do Eng_Saida; I04 "vale o manual (77,123) … Diferença 0,123"; Audit RESULTADO_SUBSTITUIDO_POR_MANUAL |
| T07 | Vizinho (C11): Delete de ID…MOTIVO numa linha com SUBSTITUI_ID | APROVADO | **Nova:** a linha (MAN_0013, SUBSTITUI_ID BIO-379339) fica **toda** vazia, com SUBSTITUI_ID e carimbo. BIO-379339 continua ATIVO (slot 133, 271,0). **Base:** sobram o SUBSTITUI_ID BIO-378769 e o carimbo R7B na linha "vazia" |
| T08 | Vizinhos: Delete de 2 linhas de uma vez; carimbo órfão legado | APROVADO | **Nova:** as duas linhas (MAN_0014/0015) ficam vazias, sem carimbo. O órfão gravado com os eventos desligados (TESTE_AUDITORIA_ORFAO) é limpo pela passada do ATUALIZAR. **Base:** o carimbo fica nas duas linhas e o órfão não é limpo |
| T09 | Estatística, Painel e LJ conferidos com o oráculo | APROVADO | Relatório B abaixo: n e média exatos; DP com diferença relativa ≤ 3,3e-13 |
| T10 | Trilha da retirada (fluxo vizinho do Delete) | APROVADO (com dúvida D2) | MAN_0001 e MAN_0002, apagados por U1, saem da tblCQ_Final no ATUALIZAR #2 com Audit RESULTADO_MANUAL_RETIRADO. O usuário da linha do Audit é **R7B**, quem rodou o ATUALIZAR, e não R7, quem apagou (ver dúvida 2) |
| T11 | Base (entrega 43), mesmo roteiro | APROVADO (prova de que o roteiro pega o defeito) | B, E e N ficam MANUAL_INCOMPLETO "DATA/HORA posterior ao lançamento", com USUARIO_MANUAL R7 (quem apagou). E reaproveita o ID MAN_0002. O SUBSTITUI_ID e o carimbo ficam nas linhas apagadas, e o órfão não é limpo. Os 3 manuais recusados ficam fora da estatística (a base é o arquivo de 15/11, com outros dados reais, por isso os n não se comparam com os da nova) |
| T12 | Apagar só o ID de uma linha com dado (vizinho) | NAO_EXECUTADO | Fora do roteiro do usuário deste reteste; coberto pelo C04 do `qa_adr097` (Tester) |

## Relatório B: conferência numérica (Painel e Estatística × oráculo)

Filtro do Painel: L4 DIMENSION 1 · C3 Glicose · H3 lote 8977 · K3 2026 · K4 Nov. Aba Estatística: 01/11/2026 a
30/11/2026. Fonte: `Value2` da tblCQ_Final da cópia `Bioquimica_F1_nova.xlsm` (produção fdc173906fff).

| Analito | Equip. | Lote | Nível | Período | Momento | Dados de entrada | Fórmula | Independente (n · média · DP · CV%) | QC_INI engPainel / Painel A7:E8 | QC_INI Estatística (n · média · DP) | Diferença relativa | Status |
|---|---|---|---|---|---|---|---|---|---|---|---|---|
| Glicose | DIMENSION 1 | 8977 | N1 | nov/2026 | antes | 12 reais, soma 930,0 | média = Σx/n; DP amostral (n−1); CV = 100·DP/média | 12 · 77,5 · 0,7977240352 · 1,0293213358 | 12 · 77,5 · 0,7977240352 · 1,0293213358 | 12 · 77,5 · 0,7977240352 | 0 | APROVADO |
| Glicose | DIMENSION 1 | 8977 | N2 | nov/2026 | antes | 12 reais, soma 3272,0 | idem | 12 · 272,6666666667 · 6,3149653901 · 2,3160019768 | 12 · 272,6666666667 · 6,3149653901 · 2,3160019768 | 12 · 272,6666666667 · 6,3149653901 | DP 1,4e-16 (Painel) / 1,4e-13 (Estat.) | APROVADO |
| Glicose | DIMENSION 1 | 8977 | N1 | nov/2026 | ATUALIZAR #2 | − BIO-379301 (77,0, substituído) + MAN_0004 78,851 (editado), MAN_0005 79,545, MAN_0010 79,7, MAN_0011 78,928, MAN_0012 77,123, MAN_0007 79,083 (o que era recusado). MAN_0001 retirado (estava no 1º) | (930 − 77 + 473,230)/17 | 17 · **78,0135294118** · 1,0547901520 · 1,3520605463 | 17 · 78,0135294118 · 1,0547901520 · 1,3520605463 | 17 · 78,0135294118 · 1,0547901520 | média 0; DP 0 (Painel) / 3,3e-13 (Estat.); CV 1,6e-16 | APROVADO |
| Glicose | DIMENSION 1 | 8977 | N2 | nov/2026 | ATUALIZAR #2 | + MAN_0003 275,781 (movido), MAN_0006 275,027, MAN_0008 275,379 (o que era recusado), MAN_0009 273,67 (digitado na origem do recorte). MAN_0002 retirado. BIO-379339 continua | (3272 + 1099,857)/16 | 16 · **273,2410625** · 5,5197913785 · 2,0201178139 | 16 · 273,2410625 · 5,5197913785 · 2,0201178139 | 16 · 273,2410625 · 5,5197913785 | média 0; DP 0 (Painel) / 5,5e-14 (Estat.) | APROVADO |

**Como reproduzir:**
1. Filtre a tblCQ_Final por PARTICIPA_ESTATISTICA = "SIM", ANALITO = "Glicose", EQUIPAMENTO = "DIMENSION 1",
   LOTE = "8977", NIVEL e DATA_HORA em novembro de 2026.
2. Calcule `statistics.fmean` e `statistics.stdev`, e o CV = 100·DP/média.
3. Faça a conta à mão: n₁ = 12 − 1 + 6 = 17, Σ₁ = 930 − 77 + 473,230 = 1326,230 (N1); n₂ = 12 + 4 = 16,
   Σ₂ = 3272 + 1099,857 = 4371,857 (N2). As duas contas batem com o oráculo, ou seja, nenhum dos 2 resultados reais
   novos do ATUALIZAR #1 caiu neste filtro.

Os números estão em `Bioquimica_F1_nova.json` (`antes`, `depois2`).

O que **não** pode entrar não entrou: o MAN_0001 e o MAN_0002 apagados, o BIO-379301 substituído, as linhas
apagadas (MAN_0013/0014/0015) e o 79,005 antigo do MAN_0004. A tolerância é a padrão (1e-9 relativo). O maior desvio,
3,3e-13 no DP da aba Estatística, é arredondamento de ponto flutuante.

**LJ (Eng_Saida, a fonte do gráfico), pontos de teste depois do ATUALIZAR #2:**

| Slot | Corrida | N1 (ID · valor) | N2 (ID · valor) | Marcas |
|---|---|---|---|---|
| 126 | 26111601 | MAN_0004 · 78,851 | MAN_0003 · 275,781 | nenhuma |
| 127 | 26111602 | MAN_0005 · 79,545 | MAN_0006 · 275,027 | nenhuma |
| 128 | 26111603 | MAN_0010 · 79,7 | — | nenhuma |
| 129 | 26111604 | MAN_0011 · 78,928 | — | nenhuma |
| 131 | 26111606 | MAN_0012 · 77,123 (substitui BIO-379301) | — | nenhuma |
| 133 | 26111608 | — | BIO-379339 · 271,0 | nenhuma |
| 134 | 26111609 | **MAN_0007 · 79,083** | **MAN_0008 · 275,379** | nenhuma |
| 135 | 26111610 | — | **MAN_0009 · 273,67** | nenhuma |

As corridas do dia foram renumeradas depois da retirada de MAN_0001/0002 (MAN_0003 passou de 26111602 para 26111601),
como o ADR-080 prevê. Nenhum ponto de teste tem \|z\| ≥ 2 em relação ao alvo do lote (o maior é o MAN_0012, z = −1,42; os demais
ficam abaixo de 0,4), então nenhuma marca de Westgard nova é esperada, e nenhuma aparece. O status do Painel continua OK.

## Ocorrência nova (fluxo vizinho)

### R7-01: recortar e colar um manual já integrado troca o dono (USUARIO_MANUAL) para quem moveu, sem registro no Audit_Log

- **Tela/módulo:** aba Digitar Resultados (`mIntegracao.EntradaMudou`/`PrepararLinha`/`Carimbar`) e motor
  (USUARIO_MANUAL da tblCQ_Final, Audit_Log).
- **Ação:**
  1. U1 digitou MAN_0003 (N2, 275,781), carimbado R7 às 10:11:57, e rodou o ATUALIZAR. O USUARIO_MANUAL ficou R7.
  2. U2 selecionou ID…MOTIVO, recortou e colou 9 linhas abaixo.
  3. U2 rodou o ATUALIZAR.
- **Esperado:** pelo critério de aceite 3 da OC-0078 ("o ID e o carimbo vão com ela"), o destino mantém R7 /
  10:11:57, e o resultado continua atribuído a quem o lançou.
- **Observado:**
  - O ID vai (MAN_0003). O carimbo do destino passa a ser **R7B / 10:14:29**, o de quem moveu, na hora do recorte.
  - Depois do ATUALIZAR, a tblCQ_Final mostra **USUARIO_MANUAL = TESTE_AUDITORIA_R7B** para o MAN_0003, e o Audit_Log
    não tem nenhuma linha sobre o MAN_0003: nem ALTERADO, nem "movido".
  - O valor, a data/hora e o status não mudam. O resultado continua ATIVO e no LJ.
  - É idêntico na base (entrega 43).
  - O ADR-097 §5 registra isso como "limite conhecido", porque USUARIO e REGISTRADO_EM não fazem parte de ID…MOTIVO.
- **Dificuldade:** nenhuma para o usuário, que nem percebe. As colunas de carimbo são travadas, e nada avisa que o
  lançamento mudou de dono.
- **Consequência:** a trilha do resultado de CQ fica errada sem deixar rastro. O sistema passa a dizer que quem lançou
  o MAN_0003 foi quem apenas o moveu, e o lançamento original deixa de ser registrado. Não afeta nenhum número. Mover
  uma linha manual é raro.
- **Sugestão (para decisão do Orquestrador):**
  - Quando o destino recebe um ID MAN_ que já está na tblCQ_Final e está sem carimbo, recuperar dali o
    USUARIO_MANUAL e o instante do lançamento.
  - Alternativa: o motor registrar no Audit_Log a troca de USUARIO_MANUAL de um manual existente.
  - Alternativa mínima: o GUIA §3 dizer "não recorte linhas manuais; para corrigir, edite na própria linha".
  - Também alinhar o critério de aceite 3 da OC-0078 ao ADR-097 §5, ou o contrário.
- **Evidência:** `Bioquimica_F1_nova.json`, `passos.F2c_recorte` (`M_antes` × `destino_logo_apos_recorte`) e
  `atualizar2.tblCQ_Final_meus` (MAN_0003), que se compara com `atualizar1` (USUARIO_MANUAL R7) e
  `atualizar2.auditoria` (nenhuma linha do MAN_0003). Base: `Bioquimica_F1_base.json`, mesmos campos.
- **Prioridade:** Médio. A trilha fica incorreta num caso raro e documentado; não há erro de cálculo nem perda de
  resultado. **Tipo:** funcional.

## Dúvidas de regra de negócio

1. **R7-01 / critério 3:** o carimbo de uma linha manual movida deve ser o do lançamento original (critério de aceite
   da OC-0078) ou o de quem moveu (ADR-097 §5)? Hoje é o de quem moveu, e o USUARIO_MANUAL do resultado já integrado
   muda em silêncio.
2. **Retirada de um manual (T10):** o Audit_Log grava RESULTADO_MANUAL_RETIRADO no nome de quem rodou o ATUALIZAR, e
   não de quem apagou a linha (no teste, R7 apagou e R7B atualizou). É o esperado? Se a trilha precisa dizer quem
   retirou, o Delete teria de ser registrado no momento em que acontece.

## Impressões de uso (Relatório A)

- **Reaproveitar uma linha apagada agora funciona como o usuário espera.** O botão Digitar Resultados leva à linha
  apagada, e o que se digita nela fica com o nome e a hora de quem digitou e entra na estatística. Ninguém precisa
  saber que existiam colunas de carimbo. Na entrega 43, o mesmo gesto perdia o resultado com um motivo impossível de
  corrigir.
- **Apagar ficou consistente.** Seja de ID…MOTIVO, de DATA…MOTIVO ou de duas linhas de uma vez, a linha fica
  realmente vazia, sem ID solto, carimbo ou SUBSTITUI_ID esquecido.
- **O Ctrl+D é seguro:** a linha de cima não muda de dono, e as de baixo ganham ID novo, com um aviso só. Um detalhe:
  o aviso fala em "linhas **coladas**" também no Ctrl+D. É compreensível, mas "copiadas" ou "repetidas" descreveria
  melhor os dois gestos (acabamento, não abro ocorrência).
- **Corrigir na própria linha** mantém ID e dono e fica no Audit_Log como alteração. É o que o GUIA manda.
- **SUBSTITUI_ID:** digitar só o número e ver a linha se preencher continua muito prático, e o I04 explica bem a
  troca.
- **Ponto de atenção:** recortar uma linha manual muda o dono dela sem nenhum aviso (R7-01). O usuário comum não tem
  como perceber.
- **Números:** n, média, DP e CV do Painel e da aba Estatística batem com a minha conta antes e depois. Os três
  manuais que a entrega 43 recusava (B, E, N) aparecem no LJ na corrida certa.
