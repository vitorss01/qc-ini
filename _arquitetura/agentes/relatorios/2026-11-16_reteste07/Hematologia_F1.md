# Reteste 07: Usuário Simulado, Hematologia, frente F1

Data: 17/11/2026 · Agente: AGENTE_USUARIO_SIMULADO · Modo: RETESTE da OC-0078 · Produção: entrega 44 (ADR-097).

## Ambiente e método

| Item | Valor |
|---|---|
| Produção atual (só copiada) | `QC_Hematologia.xlsm` sha256 curto **5b0e78ace7b9**, igual antes e depois das duas execuções na produção (`origem_intacta.igual = true`) |
| Base de comparação ("antes") | `_backup_pre_integracao_2026-11-16_0959/QC_Hematologia.xlsm` sha256 curto **49f68fba99d5** (entrega 43), igual antes e depois |
| Cópias (novas, nenhuma salva, apagadas no fim) | `%TEMP%\qcw\auditoria\2026-11-16_reteste07\Hematologia_F1_nova.xlsm`, `Hematologia_F1_base.xlsm` e `Hematologia_F1_c1.xlsm` |
| Excel | Um por vez, aberto e encerrado pelo `qa_final.QA`: PID 3472 (nova), 5700 (base) e 21728 (complemento c1). Os outros EXCEL.EXE da máquina (o Excel aberto do usuário e o da tarefa agendada) não foram tocados. Nenhuma captura de tela |
| Execução anterior interrompida | Uma primeira tentativa em 16/11 parou no 1º ATUALIZAR, sem erro gravado. O PID dela (27492) já não existia. Arquivei o log e o JSON em `anterior_interrompida\`, apaguei a cópia e o `~$` que sobraram, e refiz tudo numa cópia nova da produção. A produção tinha o mesmo hash |
| Scripts | `_arquitetura/agentes/ferramentas/reteste07_usuario_f1_hema.py <nova\|base>` (o mesmo roteiro do reteste da Bioquímica, com os auxiliares da Hema do `reteste06_usuario_f1_hema.py`) e **novo** `reteste07_usuario_f1_hema_c1.py` (complemento do critério 1 e do C04; o motivo está no cabeçalho) |
| Saídas (fora do repositório) | `Hematologia_F1_nova.json`, `Hematologia_F1_base.json`, `Hematologia_F1_c1.json` e os `.log`, na pasta do ciclo no `%TEMP%` |
| Usuários (só na cópia, ANALISTA, login por `mSeguranca.DoLogin`) | Roteiro: U1 = `TESTE_AUDITORIA_R7H` (digitou e **apagou**); U2 = `TESTE_AUDITORIA_R7HB` (digitou nas linhas apagadas, recortou, fez o Ctrl+D, editou e substituiu). Complemento: `TESTE_AUDITORIA_R7HC` (digitou e apagou) e `TESTE_AUDITORIA_R7HD` |
| Dados de teste | WBC, XN1000, lote 625311, 17/11/2026, MOTIVO "Outro". Valores sintéticos de 3 decimais calculados do alvo do lote (`mLotes.MediaDoLote`/`DPDoLote`: N1 3,18 ± 0,07935; N2 7,33 ± 0,1368; N3 17,44 ± 0,25395), com \|z\| ≤ 0,3. O manual que substitui vale o real + 0,123 (3,273). A tabela manual estava vazia no início (a produção não tem manual) |
| tblCQ_Final | 48.484 linhas no início. O 1º ATUALIZAR leu o DB_SEAC só para leitura e trouxe **155** resultados novos, de 16/11 à tarde e de 17/11 até 08:10:26. Cinco deles caem no filtro e entram na conta à mão abaixo |

**Caminho do usuário:** o mesmo do reteste da Bioquímica.
- Login; Painel com C3, H3, K3 e K4 escritos com os eventos ligados; aba Estatística com Ano De/Até (B3/D3) postos em
  2026, como o usuário faria.
- Entrada pelo botão (`mIntegracao.IrDigitarResultados`) e digitação célula a célula (`FormulaLocal`), com
  `Application.Interactive = True`, para que os avisos do `EntradaMudou` aparecessem. As caixas foram lidas pelo texto
  da janela do meu PID e fechadas no OK.
- Delete = `ClearContents` na seleção (ID…MOTIVO; DATA…MOTIVO; só a célula do ID; duas linhas de uma vez). Recorte =
  `Range.Cut(destino)`. Ctrl+D = `Range.FillDown()` em ID…MOTIVO de 3 linhas, com a de cima carimbada.
- ATUALIZAR DADOS pelo núcleo do botão (`AtualizarDadosAutomatico`), duas vezes em cada cópia.

**O que não executei:**
- o teclado e o mouse físicos (Delete, Ctrl+X/Ctrl+V e Ctrl+D de verdade);
- o cenário "dias depois", que reproduzi com 2,5 min entre o carimbo antigo e a nova digitação, como no reteste 06. O
  mecanismo (DATA/HORA digitada posterior ao carimbo herdado) é o mesmo.

**Oráculo próprio:** Python puro (`statistics`) sobre o `Value2` da tblCQ_Final, com PARTICIPA = SIM, WBC, XN1000,
lote 625311, por nível. No Painel, DATA_HORA em nov/2026; na aba Estatística, o ano 2026. Ele calcula n, média, DP
amostral e CV. Além disso, fiz a conta à mão "antes − retirados − substituído + manuais ATIVOS + reais novos do
DB_SEAC".

## Roteiro executado

| Fase | Quem | O que |
|---|---|---|
| 1 | U1 | Digita A (N1 03:01), D (N2 03:02), M (N3 03:03), G (N1 03:04) e F (N1 03:05) nas linhas 1 a 5, e o controle K (N2 03:08) na linha 8. ATUALIZAR #1: os 6 ficam ATIVO |
| 2a | U1 | Apaga A (Delete de **ID…MOTIVO**) e D (Delete só de **DATA…MOTIVO**, com o ID fora da seleção) |
| 2b | U2 | 2,5 min depois do carimbo de A, clica no botão Digitar Resultados. Digita B (N1) onde o cursor parou e E (N2) na linha de D. As duas têm DATA/HORA = carimbo antigo + 1 min (08:40), o caso que era recusado |
| 2c | U2 | Recorta ID…MOTIVO de M (já integrado) para a linha 12 e digita N (N3) na origem |
| 2d | U2 | Ctrl+D de F (linha 5) para as linhas 6 e 7; troca HORA e RESULTADO das duas de baixo |
| 2e | U2 | Corrige o RESULTADO de G (manual de U1, já integrado) na própria linha |
| 2f | U2 | ADR-082: digita só o número 379002 (HEM-379002, N1 real, 3,15) no SUBSTITUI_ID, depois RESULTADO 3,273 e MOTIVO. Vizinho C11: outra linha com SUBSTITUI_ID 379077 (N2) e Delete de ID…MOTIVO |
| 2g | U2 | Vizinhos: duas linhas digitadas e apagadas de uma vez; carimbo órfão legado (gravado com os eventos desligados) numa linha vazia |
| — | U2 | ATUALIZAR #2 |
| c1 | R7HC → R7HD | Complemento: R7HC digita A (N1 03:11) e C (N2 03:12) e roda o ATUALIZAR #1. Apaga A (ID…MOTIVO) e **só o ID** de C (C04). Depois de 2,5 min, R7HD clica no botão e digita B (N1) onde o cursor parou, com DATA/HORA = carimbo + 1 min (08:51), a 40,6 min do resultado real N1 mais próximo. ATUALIZAR #2 |

**Por que houve o complemento c1:** no roteiro principal, o B (critério 1) ganhou ID e carimbo novos e **não** foi
recusado. Mesmo assim, ficou **CONFLITO_MANUAL "coincide com resultado do interfaceamento (30 min)"**, com o A01
apontando para o HEM-379943 (N1 real de 17/11 08:10:26, chegado no 1º ATUALIZAR). As 08:40 do B ficam a 29,6 min
dele. É a regra correta do motor (ADR-082/GUIA §3), e não o defeito, mas o critério 1 pede "fica ATIVO e entra na
Estatística/LJ". O c1 refaz só esse caso com um horário a mais de 30 min de qualquer real do dia.

## Resultado

| OC | Estado | Resumo |
|---|---|---|
| OC-0078 | **RESOLVIDA** (com a mesma ressalva da Bioquímica no critério 3) | Na Hematologia, o Delete de ID…MOTIVO, o de DATA…MOTIVO e o de duas linhas limpam USUARIO e REGISTRADO_EM na hora. O Delete de DATA…MOTIVO também tira o ID que sobrava, e o Delete numa linha com SUBSTITUI_ID tira o SUBSTITUI_ID. O manual digitado depois na mesma linha, por outro usuário, ganha ID novo e o carimbo de quem digitou, agora. Ele fica ATIVO e entra no Painel, na Estatística e no LJ, com n e média exatos pelo oráculo (c1: MAN_0004; roteiro: MAN_0008 e MAN_0009). A origem do recorte fica limpa. O Ctrl+D não muda a origem, e editar um manual mantém o ID e o carimbo. O ADR-082 funciona, e o órfão legado é limpo no ATUALIZAR. Na base (entrega 43), o mesmo roteiro recusa B, E e N com "DATA/HORA posterior ao lançamento" no nome de quem apagou e reaproveita o ID MAN_0002. **Ressalva:** a linha recortada leva o ID, mas não o carimbo. Ela fica com o carimbo de quem moveu, e o USUARIO_MANUAL do resultado já integrado muda sem registro no Audit_Log. É o limite declarado no ADR-097 §5, e a base faz o mesmo: não é efeito colateral (OC nova abaixo, a mesma R7-01 da Bioquímica) |

## Testes executados

| ID | Teste | Estado | Evidência (nova = entrega 44, cópia de 5b0e78ace7b9; base = entrega 43) |
|---|---|---|---|
| OC-0078 | Reteste do fluxo original e dos 6 critérios de aceite | **RESOLVIDA** | T01 a T09 abaixo; a ressalva do critério 3 está em T03b e na ocorrência nova |
| T01 | Critério 1: Delete de ID…MOTIVO e manual novo na mesma linha, ATIVO e na Estatística/LJ | APROVADO | **c1:** logo após o Delete, a linha 1 (MAN_0001, R7HC, 08:50:36) fica com ID, USUARIO e REGISTRADO_EM **vazios**. O botão Digitar Resultados leva a **B7 = linha 1**. R7HD digita B (17/11 08:51, N1 3,168) às 08:53:06, e a linha recebe **MAN_0004**, USUARIO **TESTE_AUDITORIA_R7HD** e REGISTRADO_EM **08:53:06**. No ATUALIZAR #2, o MAN_0004 fica **ATIVO**, RUN 26111704, DETALHE vazio, USUARIO_MANUAL R7HD, e o Audit grava "manual lancado por TESTE_AUDITORIA_R7HD em 17/11/2026 08:53:06". O MAN_0001 sai (RESULTADO_MANUAL_RETIRADO). No LJ (Eng_Saida), o slot 73, corrida do lote 42, tem N1 = MAN_0004 · 3,168, alvo 3,18/0,07935, sem marca. O Painel N1 tem n = 21, igual ao oráculo (Relatório B). **Roteiro principal:** o mesmo gesto dá MAN_0007, R7HB, 08:41:31, sem recusa por "posterior ao lançamento"; o MAN_0007 fica CONFLITO_MANUAL pela regra dos 30 min (ver acima). **Base:** a linha fica com R7H/08:44:22, o B herda isso e fica **MANUAL_INCOMPLETO "DATA/HORA posterior ao lançamento"**, com USUARIO_MANUAL R7H (quem apagou) |
| T02 | Critério 2: Delete só de DATA…MOTIVO (o ID fica fora da seleção) | APROVADO | **Nova:** a linha 2 (MAN_0002) perde também o ID e o carimbo na hora. O E (17/11 08:40, N2 7,346), digitado por R7HB às 08:41:36, recebe **MAN_0008** (ID novo, não reaproveitado) e R7HB/08:41:37, e fica **ATIVO**, RUN 26111707, corrida 67. O MAN_0002 é retirado. **Base:** o ID MAN_0002 fica na linha com R7H/08:44:28, o E é gravado com o **mesmo ID MAN_0002** e o carimbo de R7H e fica MANUAL_INCOMPLETO "posterior ao lançamento" |
| T03a | Critério 3 (origem): recortar a linha e digitar outro manual na origem | APROVADO | **Nova:** logo após o recorte de MAN_0003 (linha 3 → 12), a origem fica com ID, USUARIO e REGISTRADO_EM vazios. O N (17/11 08:41, N3 17,384), digitado por R7HB às 08:41:46, recebe **MAN_0009** e R7HB/08:41:46, e fica **ATIVO**, RUN 26111707, corrida 37. **Base:** a origem fica com R7H/08:44:33, e o N recebe MAN_0008 com o carimbo de R7H e fica MANUAL_INCOMPLETO "posterior ao lançamento" |
| T03b | Critério 3 (destino): o ID e o carimbo vão com a linha recortada | **REPROVADO** (só o carimbo; o ID vai) | **Nova:** o destino (linha 12) fica com **MAN_0003**, mas o carimbo passa a ser **TESTE_AUDITORIA_R7HB / 08:41:43** (quem moveu, na hora do recorte), e não o original (R7H, 08:39:12). No ATUALIZAR #2, o MAN_0003 continua ATIVO com o mesmo valor (17,491, slot 70), mas o **USUARIO_MANUAL muda de R7H (ATUALIZAR #1) para R7HB**, e o Audit_Log não tem nenhuma linha do MAN_0003. **Base:** idêntico (R7HB/08:47:04; USUARIO_MANUAL R7HB). É o limite conhecido do ADR-097 §5, e não efeito colateral |
| T04 | Critério 4: Ctrl+D a partir de uma linha carimbada | APROVADO | **Nova:** a origem (linha 5) continua **MAN_0005 / R7H / 08:39:23**. As de baixo ganham **MAN_0010 e MAN_0011** com R7HB/08:41:52, e aparece **uma** caixa: "As 2 linhas coladas traziam IDs que já existem em outras linhas e receberam IDs novos (MAN_0010 a MAN_0011). … Atenção: com a mesma data, hora, nível e analito de outra linha, ela fica como repetida (CONFLITO_MANUAL), mesmo com outro resultado. Se for uma nova dosagem, use a hora real dela." Com a HORA (03:21/03:22) e o RESULTADO trocados, as três ficam ATIVO (RUN 26111702/03/04, corridas 41/42/43), e o MAN_0005 mantém USUARIO_MANUAL R7H. Na base, a origem também não muda (o aviso é o texto antigo, sem a frase do CONFLITO_MANUAL) |
| T05 | Critério 5: editar o resultado de um manual existente na própria linha | APROVADO | **Nova:** o MAN_0004 (R7H, 08:39:17, 3,164, já ATIVO) foi corrigido por R7HB para 3,156. O ID e o carimbo **não mudam**, e a tela continua ATIVO, RUN 26111702. No ATUALIZAR #2, ele fica ATIVO, 3,156, USUARIO_MANUAL R7H, e o Audit grava **RESULTADO_MANUAL_ALTERADO** 3,164 → 3,156 ("manual lancado por TESTE_AUDITORIA_R7H em 17/11/2026 08:39:17"). No LJ, o slot 70 tem N1 = MAN_0004 · 3,156 |
| T06 | Critério 6: o ADR-082 (SUBSTITUI_ID) continua funcionando | APROVADO | **Nova:** digitando só "379002" no SUBSTITUI_ID de uma linha vazia, ele vira **HEM-379002**, e a linha se preenche sozinha (XN1000, SANGUE TOTAL, 625311, N1, WBC, 16/11 08:09), com o ID MAN_0012, o carimbo R7HB/08:42:03 e o DETALHE "atualize os dados". O SUBSTITUI_ID **não** se perde ao digitar o RESULTADO e o MOTIVO. No ATUALIZAR #2, o MAN_0012 fica **ATIVO**, RUN 26111601, DETALHE "substitui o resultado do interfaceamento HEM-379002". O HEM-379002 fica **SUBSTITUIDO_MANUAL**, PARTICIPA NÃO, sem RUN. O I04 diz "vale o manual (3,273) … Diferença manual - automatico: 0,123", e o Audit grava RESULTADO_SUBSTITUIDO_POR_MANUAL e INCLUIDO. No LJ, o slot 68 tem N1 = MAN_0012 · 3,273 |
| T07 | Vizinho (C11): Delete de ID…MOTIVO numa linha com SUBSTITUI_ID | APROVADO | **Nova:** a linha (MAN_0013, SUBSTITUI_ID HEM-379077) fica **toda** vazia, sem SUBSTITUI_ID e sem carimbo. O HEM-379077 continua ATIVO (slot 68, N2 7,36). **Base:** sobram o SUBSTITUI_ID HEM-378911 e o carimbo R7HB na linha "vazia" |
| T08 | Vizinhos: Delete de 2 linhas de uma vez; carimbo órfão legado | APROVADO | **Nova:** as duas linhas (MAN_0014/0015) ficam vazias, sem carimbo. O órfão gravado com os eventos desligados (TESTE_AUDITORIA_ORFAO) é limpo pela passada do ATUALIZAR. **Base:** o carimbo fica nas duas linhas e o órfão não é limpo |
| T09 | Vizinho (C04, ADR-097 §4): apagar só o ID de uma linha com dado | APROVADO | **c1:** o MAN_0002 (R7HC, 08:50:41, N2 7,314, já ATIVO), sem o ID, ganha **MAN_0003** na hora e **mantém** o carimbo R7HC/08:50:41, sem caixa. No ATUALIZAR #2, o MAN_0003 fica ATIVO (RUN 26111701, slot 70), o Audit grava INCLUIDO "manual lancado por TESTE_AUDITORIA_R7HC em 17/11/2026 08:50:41", e o MAN_0002 é RETIRADO. Não há duplicidade: o 7,314 conta uma vez só (conta à mão abaixo) |
| T10 | Estatística, Painel e LJ conferidos com o oráculo | APROVADO | Relatório B abaixo: n e média exatos, e DP com diferença relativa ≤ 2,2e-16 |
| T11 | Trilha da retirada (fluxo vizinho do Delete) | APROVADO (com a dúvida 2) | MAN_0001 e MAN_0002, apagados por R7H, saem da tblCQ_Final no ATUALIZAR #2 com RESULTADO_MANUAL_RETIRADO. O usuário gravado na linha do Audit é **R7HB**, quem rodou o ATUALIZAR, e não R7H, quem apagou. No c1 é igual: apagou R7HC, e o Audit grava R7HD |
| T12 | Base (entrega 43), mesmo roteiro | APROVADO (prova de que o roteiro pega o defeito) | B, E e N ficam MANUAL_INCOMPLETO "DATA/HORA posterior ao lançamento", com USUARIO_MANUAL R7H (quem apagou), e o Audit grava RESULTADO_MANUAL_RECUSADO. O E reaproveita o ID MAN_0002. O SUBSTITUI_ID e os carimbos ficam nas linhas apagadas, e o órfão não é limpo |

## Relatório B: conferência numérica (Painel e Estatística × oráculo)

Filtro do Painel: C3 WBC · H3 lote 625311 · K3 2026 · K4 Nov (G3:G4 = 01/11/2026 a 30/11/2026). Aba Estatística: lote
B4 625311, Ano De/Até 2026. Fonte: `Value2` da tblCQ_Final das cópias de 5b0e78ace7b9. "Painel" é o `engPainel`, igual
ao Painel A7:E9. Fórmula em todas as linhas: média = Σx/n; DP amostral (n−1); CV = 100·DP/média.

| Analito | Equip. | Lote | Nível | Período | Momento | Dados de entrada | Independente (n · média · DP · CV%) | QC_INI Painel (n · média · DP · CV%) | QC_INI Estatística 2026 (n · média · DP) × oráculo do ano | Diferença relativa máx. | Status |
|---|---|---|---|---|---|---|---|---|---|---|---|
| WBC | XN1000 | 625311 | N1 | nov/2026 | antes | 18 reais, Σ 56,44 | 18 · 3,1355555556 · 0,0700606647 · 2,2343939826 | 18 · 3,1355555556 · 0,0700606647 · 2,2343939826 | 39 · 3,1274358974 · 0,0616320981 = oráculo | 1,1e-16 | APROVADO |
| WBC | XN1000 | 625311 | N2 | nov/2026 | antes | 24 reais, Σ 175,81 | 24 · 7,3254166667 · 0,1043813729 · 1,4249206247 | idem | 63 · 7,3326984127 · 0,1094591741 = oráculo | 2,4e-16 | APROVADO |
| WBC | XN1000 | 625311 | N3 | nov/2026 | antes | 13 reais, Σ 227,34 | 13 · 17,4876923077 · 0,2380179911 · 1,3610600355 | idem | 34 · 17,4232352941 · 0,2065990200 = oráculo | 2,0e-16 | APROVADO |
| WBC | XN1000 | 625311 | N1 | nov/2026 | roteiro, ATUALIZAR #2 | 56,44 − HEM-379002 3,15 + MAN_0012 3,273 + MAN_0004 3,156 (editado) + MAN_0005 3,192 + MAN_0010 3,2 + MAN_0011 3,16 + reais HEM-379713 3,09 e HEM-379943 3,13. MAN_0001 retirado; MAN_0007 CONFLITO fora | 24 · **3,1454583333** · 0,0692299126 · 2,2009483278 (Σ 75,491) | 24 · 3,1454583333 · 0,0692299126 · 2,2009483278 | 45 · 3,1338 · 0,0632310338 = oráculo | 0 | APROVADO |
| WBC | XN1000 | 625311 | N2 | nov/2026 | roteiro, ATUALIZAR #2 | 175,81 + MAN_0006 7,337 + MAN_0008 7,346 (o que era recusado) + reais HEM-379548 7,4 e HEM-379788 7,3. MAN_0002 retirado; HEM-379077 continua | 28 · **7,3283214286** · 0,0975868288 · 1,3316395813 (Σ 205,193) | 28 · 7,3283214286 · 0,0975868288 · 1,3316395813 | 67 · 7,3334776119 · 0,1065004422 = oráculo | 1,3e-16 | APROVADO |
| WBC | XN1000 | 625311 | N3 | nov/2026 | roteiro, ATUALIZAR #2 | 227,34 + MAN_0003 17,491 (movido) + MAN_0009 17,384 (digitado na origem do recorte) + real HEM-379863 17,67 | 16 · **17,4928125** · 0,2196072915 · 1,2554144255 (Σ 279,885) | 16 · 17,4928125 · 0,2196072915 · 1,2554144255 | 37 · 17,4306756757 · 0,2023143942 = oráculo | 2,0e-16 | APROVADO |
| WBC | XN1000 | 625311 | N1 | nov/2026 | c1, ATUALIZAR #2 | 56,44 + reais 3,09 e 3,13 + **MAN_0004 3,168** (o manual do critério 1). MAN_0001 retirado | 21 · **3,1346666667** · 0,0657976697 · 2,0990324223 (Σ 65,828) | 21 · 3,1346666667 · 0,0657976697 · 2,0990324223 | 42 · 3,1275714286 · 0,0599585804 = oráculo | 2,1e-16 | APROVADO |
| WBC | XN1000 | 625311 | N2 | nov/2026 | c1, ATUALIZAR #2 | 175,81 + reais 7,4 e 7,3 + MAN_0003 7,314 (C04: o mesmo dado com ID novo, contado uma vez). MAN_0002 retirado | 27 · **7,3268148148** · 0,0993985904 · 1,3566412269 (Σ 197,824) | 27 · 7,3268148148 · 0,0993985904 · 1,3566412269 | 66 · 7,3329393939 · 0,1073302422 = oráculo | 2,4e-16 | APROVADO |
| WBC | XN1000 | 625311 | N3 | nov/2026 | c1, ATUALIZAR #2 | 227,34 + real 17,67 | 14 · 17,5007142857 · 0,2338133472 · 1,3360217381 (Σ 245,01) | idem | 35 · 17,4302857143 · 0,2077680510 = oráculo | 2,0e-16 | APROVADO |

**Como reproduzir:**
1. Filtre a tblCQ_Final por PARTICIPA_ESTATISTICA = "SIM", ANALITO = "WBC", EQUIPAMENTO = "XN1000", LOTE = "625311",
   NIVEL e DATA_HORA em novembro de 2026 (Painel) ou no ano de 2026 (Estatística).
2. Calcule `statistics.fmean` e `statistics.stdev`, e o CV = 100·DP/média.
3. Faça a conta à mão. Os cinco reais novos do DB_SEAC no filtro estão listados em `Hematologia_F1_c1.json`
   (`filtro_depois` − `filtro_antes`): HEM-379548 (N2, 16/11 12:36), HEM-379713 (N1, 17/11 07:54), HEM-379788 (N2,
   07:55), HEM-379863 (N3, 07:56) e HEM-379943 (N1, 08:10).
   - **Roteiro:** n₁ = 18 − 1 + 5 + 2 = 24, Σ₁ = 75,491; n₂ = 24 + 2 + 2 = 28, Σ₂ = 205,193; n₃ = 13 + 2 + 1 = 16,
     Σ₃ = 279,885.
   - **c1:** n₁ = 18 + 2 + 1 = 21, Σ₁ = 65,828; n₂ = 24 + 2 + 1 = 27, Σ₂ = 197,824; n₃ = 14.

   As somas batem com o oráculo, e o oráculo bate com o QC_INI.

O que **não** pode entrar não entrou:
- os apagados MAN_0001 e MAN_0002;
- o HEM-379002 substituído;
- o MAN_0007, CONFLITO pela regra dos 30 min;
- as linhas apagadas MAN_0013/0014/0015;
- o 3,164 antigo do MAN_0004;
- no c1, o MAN_0002 antigo do C04.

A tolerância é a padrão (1e-9 relativo). O maior desvio, 2,4e-16, é arredondamento de ponto flutuante.

**LJ (Eng_Saida, a fonte do gráfico), pontos de teste depois do ATUALIZAR #2 do roteiro:**

| Slot | RUN | Corrida do lote (N1/N2/N3) | N1 (ID · valor) | N2 (ID · valor) | N3 (ID · valor) | Marcas |
|---|---|---|---|---|---|---|
| 68 | 26111601 | 39 / 63 / 34 | MAN_0012 · 3,273 (substitui HEM-379002) | HEM-379077 · 7,36 | HEM-379157 · 17,78 | OK |
| 70 | 26111701 | 40 / 65 / 35 | MAN_0004 · 3,156 | MAN_0006 · 7,337 | MAN_0003 · 17,491 | OK |
| 71 | 26111702 | 41 / — / — | MAN_0005 · 3,192 | — | — | OK |
| 72 | 26111703 | 42 / — / — | MAN_0010 · 3,2 | — | — | OK |
| 73 | 26111704 | 43 / — / — | MAN_0011 · 3,16 | — | — | OK |
| 76 | 26111707 | — / 67 / 37 | — | **MAN_0008 · 7,346** | **MAN_0009 · 17,384** | OK |

A corrida do lote termina em 45/67/37, que é o n da aba Estatística em 2026 (45/67/37), como o ADR-080 prevê. Os
reais de 17/11 (07:54 a 08:10) ocupam as corridas 44/45 (N1), 66 (N2) e 36 (N3). Nenhum ponto de teste tem \|z\| ≥ 2
pelo alvo do lote: o maior é o MAN_0012, com z = (3,273 − 3,18)/0,07935 = +1,17. Portanto, nenhuma marca de
Westgard nova era esperada, e nenhuma apareceu (status do Painel OK nos 3 níveis).

## Ocorrência nova (fluxo vizinho)

### Hema: recortar e colar um manual já integrado troca o dono (USUARIO_MANUAL) para quem moveu, sem registro no Audit_Log

É a mesma R7-01 da Bioquímica, agora confirmada na Hematologia. Pode ser registrada como uma ocorrência só, "Ambos".

- **Tela/módulo:** aba Digitar Resultados (`mIntegracao.EntradaMudou`/`PrepararLinha`/`Carimbar`) e motor
  (USUARIO_MANUAL da tblCQ_Final, Audit_Log).
- **Ação:**
  1. R7H digitou o MAN_0003 (N3, 17,491), carimbado R7H às 08:39:12, e rodou o ATUALIZAR. O USUARIO_MANUAL ficou R7H.
  2. R7HB selecionou ID…MOTIVO, recortou e colou 9 linhas abaixo.
  3. R7HB rodou o ATUALIZAR.
- **Esperado:** pelo critério de aceite 3 da OC-0078 ("o ID e o carimbo vão com ela"), o destino mantém R7H/08:39:12,
  e o resultado continua atribuído a quem o lançou.
- **Observado:**
  - O ID vai (MAN_0003). O carimbo do destino passa a ser **R7HB/08:41:43**, o de quem moveu.
  - Depois do ATUALIZAR, a tblCQ_Final mostra **USUARIO_MANUAL = TESTE_AUDITORIA_R7HB**, e o Audit_Log não tem nenhuma
    linha sobre o MAN_0003.
  - O valor, a data/hora, o status e o ponto do LJ não mudam.
  - É idêntico na base (entrega 43).
  - O ADR-097 §5 registra isso como "limite conhecido".
- **Dificuldade:** nenhuma para o usuário, que nem percebe. As colunas de carimbo são travadas, e nada avisa.
- **Consequência:** a trilha de quem lançou o resultado de CQ fica errada sem deixar rastro. Nenhum número muda.
  Mover uma linha manual é raro.
- **Sugestão (para decisão do Orquestrador):**
  - Recuperar da tblCQ_Final o USUARIO_MANUAL e o instante do lançamento quando o destino recebe um MAN_ que já existe.
  - Alternativa: o motor registrar no Audit_Log a troca de USUARIO_MANUAL.
  - Alternativa mínima: o GUIA §3 dizer "não recorte linhas manuais; corrija na própria linha".
  - Também alinhar o critério 3 da OC-0078 ao ADR-097 §5.
- **Evidência:** `Hematologia_F1_nova.json`:
  - `passos.F2c_recorte` (`M_antes` × `destino_logo_apos_recorte`);
  - `atualizar1.tblCQ_Final_meus` (MAN_0003, USUARIO_MANUAL R7H) × `atualizar2.tblCQ_Final_meus` (R7HB);
  - `atualizar2.auditoria`, sem linha do MAN_0003;
  - na base, `Hematologia_F1_base.json`, mesmos campos.
- **Prioridade:** Médio. **Tipo:** funcional.

## Dúvidas de regra de negócio

1. **Critério 3 × ADR-097 §5:** o carimbo de uma linha manual movida deve ser o do lançamento original (critério de
   aceite) ou o de quem moveu (ADR)? Hoje é o de quem moveu, e o USUARIO_MANUAL do resultado já integrado muda em
   silêncio.
2. **Retirada de um manual (T11):** o Audit_Log grava RESULTADO_MANUAL_RETIRADO no nome de quem rodou o ATUALIZAR, e
   não de quem apagou a linha. É o esperado?
3. **Linha reaproveitada e a regra dos 30 min (só registro):** um manual digitado com a hora de "agora" pode cair a
   menos de 30 min de um resultado do interfaceamento do mesmo nível e ficar CONFLITO_MANUAL (o MAN_0007 do roteiro,
   a 29,6 min do HEM-379943). A regra está certa e documentada, e o DETALHE e o A01 explicam o que fazer. Fica o
   lembrete de que, na rotina, o manual do mesmo nível logo depois de um automático tende a cair nesse caso. Não abro
   ocorrência.

## Impressões de uso (Relatório A)

- **Reaproveitar uma linha apagada funciona como o usuário espera, também na Hematologia.** O botão Digitar Resultados
  leva à linha apagada, e o que se digita ali fica no nome e na hora de quem digitou e entra no Painel, na Estatística
  e no LJ. Na entrega 43, o mesmo gesto perdia o resultado com "DATA/HORA posterior ao lançamento", um motivo que o
  usuário não tem como corrigir.
- **Apagar ficou consistente:** de ID…MOTIVO, de DATA…MOTIVO, de duas linhas ou com SUBSTITUI_ID, a linha fica
  realmente vazia.
- **Apagar só o ID** gera outro na hora, sem aviso. O dado não se perde nem duplica. Quem olha a tela vê o número
  mudar sem saber por quê, mas o GUIA §3 já manda apagar o ID nesse caso.
- **Ctrl+D:** a linha de cima não muda de dono, e o aviso único diz quais IDs foram gerados e lembra de usar a hora
  real. O aviso fala em "linhas coladas" também no Ctrl+D (acabamento, como na Bioquímica).
- **SUBSTITUI_ID:** digitar só o número e ver a linha se preencher (inclusive a MATRIZ SANGUE TOTAL) é prático, e o
  I04 explica bem a troca.
- **Ponto de atenção:** recortar uma linha manual muda o dono dela em silêncio (ocorrência acima).
- **Números:** n, média, DP e CV do Painel e da aba Estatística batem com a minha conta antes e depois, nos 3 níveis.
  A corrida do lote no LJ fecha com o n da Estatística.
