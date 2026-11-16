# Parecer consolidado: ciclo 2026-11-15_reteste06

Agente: Usuário Simulado (parecer consolidado). Data: 15/11/2026. Modo: reteste, consolidação.

**De onde vêm os dados.**
- Neste passo o Excel não foi aberto.
- Tudo vem dos 2 relatórios das frentes desta pasta (`Bioquimica_F1.md`, `Hematologia_F1.md`) e do resumo
  estruturado que cada frente devolveu.
- Para conferir o Relatório B, li só para leitura os JSON das frentes na pasta do ciclo no `%TEMP%`
  (`Bioquimica_F1_nova.json`, `Hematologia_F1_nova.json`, blocos `antes`, `depois1`, `depois2`). Recalculei a
  diferença relativa oráculo × `engPainel` em todos os níveis e momentos, e refiz à mão as somas "antes + novos" e os z
  do Westgard (B.3).

| Item | Valor |
|---|---|
| Produção auditada (entrega 43, ADR-096) | `QC_Bioquimica.xlsm` sha256 curto **bd5314426ab8**; `QC_Hematologia.xlsm` sha256 curto **49f68fba99d5**. Nas duas frentes, o hash foi igual antes e depois de todas as execuções: a produção não foi aberta para escrita |
| Base de comparação (entrega 42) | Bio: `_backup_pre_integracao_2026-11-15_1802`, sha256 curto **b20fe00f7e51**. Hema: `_backup_pre_integracao_2026-11-15_2059`, sha256 curto **7bab8fc008f6**. Nas duas, hash igual antes e depois |
| Lista do reteste | 1 item nos dois produtos: **OC-0073** (linha manual copiada e colada com o ID_REGISTRO de outra aparece ATIVA na tela, mas o motor a deixa CONFLITO_MANUAL, fora da estatística, do LJ e do Westgard). Tratada no ADR-096 |
| Frentes executadas | Só a F1, nos dois produtos, em modo reteste, com os fluxos vizinhos. **Não fizeram parte deste reteste:** a F2 (CEQ), a F3 (bias, ET e Sigma) e a F4 completa (mês a mês, trimestres, troca de lote, H3 vazio). As outras OC da entrega 43 que estão em AGUARDA RETESTE (OC-0074, OC-0075, OC-0076) **não** foram retestadas neste ciclo |
| Caminho | O do usuário: login (`mSeguranca.DoLogin`), filtros do Painel com os eventos ligados, botão Digitar Resultados (`mIntegracao.IrDigitarResultados`, que parou em B7), digitação célula a célula (`FormulaLocal`), colar com `Range.Copy(destino)` e recortar com `Range.Cut(destino)`, `Application.Interactive = True` para a caixa de aviso aparecer (lida pelo texto da janela do PID da frente e fechada no OK), ATUALIZAR DADOS pelo núcleo do botão (`AtualizarDadosAutomatico`). Na Hema, o Ano De/Até da aba Estatística (B3/D3) escrito como o usuário faria |
| Fora do caminho do usuário (declarado) | Não foram executados o teclado e o mouse físicos (Ctrl+V de verdade, clique no cabeçalho da linha). A proteção que o teclado obedece foi **emulada** com `Protect` sem `UserInterfaceOnly`. O cenário "dias depois" da P-01 foi reproduzido com 2,5 min de intervalo (o mecanismo é o mesmo). O repetido "com eventos desligados" (T08) foi criado de propósito por automação, para simular dado antigo |
| Dados de teste | Bio: Glicose, DIMENSION 1, lote 8977, 14/11/2026; N1 79,468 / 79,005 / 79,777 / 78,697 / 79,391 / 79,237; N2 **294,884 (média + 4 DP)** / 273,268 / 275,781 / 275,279. Hema: WBC, XN1000, lote 625311, 14/11/2026; N1 3,188 / 3,148 / 3,204 / 3,164 / 3,184 / 3,176; N2 **7,877 (média + 4 DP)** / 7,357 / 7,344; N3 17,364. ID inválido `TESTEAUD9`. MOTIVO "Outro". Usuários `TESTE_AUDITORIA_R6`, `_R6B`, `_R6C` (Bio) e `_R6H`, `_R6HB`, `_R6HC` (Hema), ANALISTA, criados só nas cópias |
| Cópias | Em `%TEMP%\qcw\auditoria\2026-11-15_reteste06\`: Bio `nova`, `base`, `carimbo`, `carimbo_base`; Hema `nova` (2 execuções), `base`, `carimbo`, `carimbo_base`. Nenhuma salva, todas apagadas no fim. Ficam os JSON e os logs |
| Excel | Um por vez, aberto e encerrado pelo `qa_final.QA`. Bio: PIDs 27380, 5392, 31420, 28200. Hema: 29952, 9308, 21892, 28536, 23340. Os Excel de outros processos e da tarefa agendada não foram tocados. Nenhuma captura de tela |
| Ferramentas | Em `_arquitetura/agentes/ferramentas/`, documentadas no cabeçalho: `reteste06_usuario_f1.py <nova\|base>`, `reteste06_carimbo_residual.py [nova\|base]` (Bio) e `reteste06_usuario_f1_hema.py <nova\|base\|carimbo\|carimbo_base>` (Hema) |
| tblCQ_Final | Bio 35.906 linhas; Hema 48.391 linhas no início. Os ATUALIZAR leram o DB_SEAC só para leitura e não trouxeram resultado novo (último real: Bio 13/11 08:39; Hema WBC 13/11 13:10:20) |

**Resumo em números**

| | Bioquímica | Hematologia | Total |
|---|---|---|---|
| Itens do reteste: RESOLVIDA / PERSISTE / EFEITO_COLATERAL | 1 / 0 / 0 | 1 / 0 / 0 | OC-0073 **RESOLVIDA** nos dois produtos |
| Testes (T01..T17): APROVADO / REPROVADO / NAO_EXECUTADO | 15 / 2 (T12, T13) / 0 | 14 / 3 (T12, T13, T17) / 0 | 29 / 5 / 0 |
| REPROVADOs que já existiam na entrega 42 (base) | 2 de 2 | T12 e T13 sim; T17 não comparado com a base | Nenhum REPROVADO é efeito colateral da entrega 43 |
| Cálculos no Relatório B (aprovados / reprovados) | 5 + 1 Westgard (6 / 0) | 14 + 1 Westgard (15 / 0) | 21 / 0 |
| Maior diferença relativa (n, média, DP, CV) | 5,0e-13 (DP da aba Estatística); 0 no `engPainel` | 2,4e-16 | Todas abaixo de 1e-9 |
| Ocorrências das frentes / deduplicadas (P-xx) | 4 / — | 5 / — | 9 / 5 |
| Críticas / Altas / Médias / Baixas (deduplicadas) | — | — | 0 / 1 / 3 / 1 |
| Dúvidas de regra de negócio (deduplicadas) | 2 | 3 | 3 |

**Situação do item do reteste**

| Item | Bio | Hema | Consolidado |
|---|---|---|---|
| OC-0073: linha colada com o ID de outra | **RESOLVIDA**. Toda colagem ou digitação de um ID MAN_ existente gera ID novo na hora, com aviso (A:K, bloco de 3 com uma caixa só, ID digitado, inclusive "man_1"). O colado e editado (294,884) fica ATIVO, entra no Painel, na Estatística, no LJ (slot 126) e dispara o 1-3s (N2 REJEITADO). A origem não muda. Mover mantém o ID. O repetido com eventos desligados é apontado pela mensagem do ATUALIZAR e pelo A01 novo ("Não inative o MAN_0001"); apagar o ID resolve em 1 passo. Base: 11 linhas `MAN_0001~L<n>` CONFLITO mostradas ATIVO, sem 1-3s | **RESOLVIDA**. Mesmo roteiro e mesmo resultado; colado e editado (7,877) ATIVO, LJ slot 68, corrida do lote N2 63, 1-3s (z = 3,9985), J8 REJEITADO. Base: defeito reproduzido (N1 18, N2 24, N3 12 contra 20, 25, 13 na nova) | **RESOLVIDA** nos dois produtos, com prova pela base de que o roteiro pega o defeito. Resíduo previsto no ADR-096 e coberto pela mensagem: o repetido que entra com os eventos desligados aparece ATIVO na tela até ser corrigido |

---

## Relatório A: experiência do usuário

### A.1 Funcionalidades testadas

| Frente | Bioquímica | Hematologia |
|---|---|---|
| F1: colar e digitar ID na Digitar Resultados | Colar sem editar, colar e editar (média + 4 DP), colar e trocar só o resultado, colar e trocar hora e resultado, bloco de 3, ID MAN_ digitado (linha vazia e "man_1" sobre linha preenchida), mover (recortar), ID inválido repetido, repetido com eventos desligados e sua correção | Os mesmos cenários, com 3 níveis, a MATRIZ preenchida sozinha e o XN1000 como equipamento único |
| F1: proteção da aba | Colar a linha inteira (A:Q) × o bloco ID..MOTIVO (A:K) com a proteção que o teclado obedece | Idem |
| F1: carimbo (USUARIO, REGISTRADO_EM) | Apagar uma linha manual e digitar outro manual nela, com controle numa linha nunca usada; apagar a linha colada; trilha de quem colou | Idem |
| F1: ATUALIZAR DADOS | Mensagem (linha nova do ID repetido), A01 novo, E04, tblCQ_Final, Audit_Log (INCLUIDO, RETIRADO, WESTGARD_EVENTO_CORRIDA_NOVA) | Idem |
| F1: o que o usuário vê depois | Painel (engPainel e A7:T8), Eng_Saida (pontos do LJ, alvo, marcas do Westgard), aba Estatística de novembro | Painel (A7:T9), Eng_Saida de 3 níveis com a corrida do lote, aba Estatística por ano (B3/D3) |

### A.2 Facilidade de uso

**Colar ficou seguro.**
- O aviso aparece na hora, com título "Digitar Resultados": diz o ID antigo, o novo, que o ID é gerado pelo sistema e o
  que fazer antes do ATUALIZAR.
- No bloco colado é uma caixa só para as 3 linhas, sem fila de caixas.
- O ID novo não muda quando a linha é editada depois de colada.
- Na Hema, a MATRIZ se preenche sozinha, e o equipamento único não confunde.

**A tela bate com o motor** para tudo o que entra com os eventos ligados: cada linha tem o próprio ID e o próprio
STATUS, RUN e DETALHE.

**Dá para ver o que foi salvo.** No Painel, o n e a média mudam na hora; o ponto colado e editado aparece no LJ e
dispara o 1-3s (REJEITADO, "Erro Aleatório"). Na entrega 42 esse ponto ficava fora sem aviso.

**Mensagens novas.** A linha do ATUALIZAR para o ID repetido ("Apague o ID da linha repetida (um novo é gerado) e
atualize de novo") e o A01 novo ("Não inative o MAN_0001: isso tiraria o resultado que vale") são claros e levam à
correção em 1 passo.

### A.3 Dificuldades encontradas

**Reaproveitar uma linha apagada (P-01, Alto).**
- Depois de apagar ID..MOTIVO, o USUARIO e o REGISTRADO_EM ficam na linha (as colunas são travadas).
- O botão Digitar Resultados leva o cursor exatamente para essa linha.
- O manual digitado ali herda o carimbo antigo e é recusado: "DATA/HORA posterior ao lançamento".
- O usuário não tem como corrigir: o DETALHE manda olhar a data/hora, que está certa, e o carimbo não é editável.
- O ADR-096 e o A01 novo mandam "apague a linha", o que torna esse caminho mais comum.

**Copiar a linha inteira (P-02, Médio).**
- O GUIA §3 diz "pode copiar uma linha e colar embaixo". O gesto natural (selecionar a linha toda ou clicar no número
  da linha) inclui as colunas travadas, e o Excel recusa com uma mensagem que fala em senha.
- Só o bloco ID..MOTIVO funciona, e nada na tela diz isso.

**ID fora do padrão (P-03, Médio).**
- A mensagem do ATUALIZAR manda "veja o DETALHE", mas o DETALHE da linha mostra "atualize os dados" e o STATUS "—",
  mesmo depois de atualizar duas vezes. O motivo só aparece na QA_INTEGRACAO.

**Repetir o controle no mesmo minuto trocando só o valor (P-04, Baixo).**
- O aviso da colagem diz "mude o que for diferente (nível, hora, resultado)", mas trocar só o resultado deixa a linha
  repetida (a chave não inclui o valor).
- O A01 desse ramo ainda diz "inative um deles", o mesmo tropeço que o ADR-096 corrigiu no ramo do ID.

**Aba Estatística da Hema abre em 2025 (P-05, Médio).**
- B3 = D3 = 2025 no arquivo de produção; o lote em uso (625311) é de 2026 e aparece com n 0, média e DP vazios.
- Quem confere o lançamento pela Estatística acha que o resultado não entrou. Com 2026, os números batem.

### A.4 Navegação

- O botão Digitar Resultados leva à primeira linha "livre", que pode ser uma linha apagada com carimbo residual (P-01).
- O Painel e o Eng_Saida refletem o ATUALIZAR sem passo extra.
- Na Hema, a Estatística exige trocar o ano à mão (P-05).

### A.5 Apresentação

- O STATUS do repetido antigo (eventos desligados) continua ATIVO na tela até ser corrigido, porque o XLOOKUP acha a 1ª
  linha. É o resíduo previsto no ADR-096, e a mensagem do ATUALIZAR o cobre.
- A mensagem do ATUALIZAR não conta os CONFLITO por chave repetida (P-04), só os de ID repetido e os incompletos.

### A.6 Tempos observados

As frentes não mediram tempo neste reteste.

### A.7 Sugestões do usuário

Para o Orquestrador priorizar. Nenhuma foi implementada.

1. **P-01:** no `EntradaMudou` da aba manual, limpar USUARIO e REGISTRADO_EM quando ID..MOTIVO e SUBSTITUI_ID ficarem
   vazios (como a aba Inativar já faz no ADR-070); no `PrepararEntradas`, limpar o carimbo das linhas sem dado;
   opcionalmente, carimbar de novo quando a linha passa de vazia para preenchida.
2. **P-02:** GUIA §3 e ADR-096: "selecione de ID até MOTIVO (as colunas brancas), copie e cole na próxima linha";
   opcionalmente, uma dica na célula do ID.
3. **P-03:** fazer o DETALHE cair no `MAN_INVALIDO_L<linha>` quando o ID da linha não for MAN_ válido, ou mostrar "ID
   fora do padrão MAN_0001: apague o ID (um novo é gerado)"; ou a mensagem do ATUALIZAR citar o E04 de ID.
4. **P-04:** aviso: "mude a hora ou o nível; só trocar o resultado deixa a linha como repetida"; A01 desse ramo:
   "apague a linha repetida <ID>; não inative o <ID que vale>".
5. **P-05:** iniciar B3/D3 no ano do lote em análise ou no corrente, ou avisar com n 0 ("o lote <B4> não tem resultado
   em <ano>; ajuste o Ano De/Até"). Tratar junto da OC-0048.

---

## Relatório B: validação matemática

**Escopo deste ciclo.** Só a F1. Bias, ET e Sigma **não** foram recalculados de forma independente (não houve F3 nem
F2). O que se validou:
1. n, média, DP e CV das séries que recebem os manuais colados e digitados, no Painel (`engPainel`) e na aba
   Estatística, antes do primeiro ATUALIZAR, depois dele e depois do segundo (com a correção do repetido);
2. que os valores que **não** podem entrar (CONFLITO por chave, `MAN_INVALIDO_L*`, o repetido antes da correção) não
   entraram;
3. os pontos novos no Eng_Saida (fonte do LJ), o alvo do lote em cada ponto e a única marca de Westgard esperada.

**Oráculo das frentes:** Python puro (`statistics.fmean`, `statistics.stdev`, CV = 100·DP/média) sobre o `Value2` da
tblCQ_Final da cópia, filtrando PARTICIPA_ESTATISTICA = "SIM", analito, equipamento, lote, nível e a janela de
DATA_HORA. Tolerância: a padrão, 1e-9 relativo (REGRAS §3.5).

### B.1 Tabela

Bio: cópia `Bioquimica_F1_nova.xlsm` (produção bd5314426ab8). Painel: L4 DIMENSION 1 · C3 Glicose · H3 lote 8977 ·
K3 2026 · K4 Nov. Estatística: janela 01/11/2026–30/11/2026.
Hema: cópia `Hematologia_F1_nova.xlsm` (produção 49f68fba99d5). Painel: C3 WBC · H3 lote 625311 · K3 2026 · K4 Nov
(G3/G4 01/11/2026–30/11/2026). Estatística: B3 = D3 = 2026, B4 = 625311.

Formato dos resultados: n · média · DP (· CV % quando a frente o registrou).

| Analito | Equipamento | Lote | Nível | Período | Dados de entrada | Fórmula encontrada | Resultado independente | Resultado QC_INI (Value2) | Diferença (relativa) | Status |
|---|---|---|---|---|---|---|---|---|---|---|
| Glicose | DIMENSION 1 | 8977 | N1 | nov/2026, antes | 10 reais, soma 775,0 | média = Σx/n; DP amostral; CV = 100·DP/média | 10 · 77,5 · 0,8498365856 · 1,0965633 | engPainel e Estatística: 10 · 77,5 · 0,8498365856 · 1,0965633 | 0 | APROVADO |
| Glicose | DIMENSION 1 | 8977 | N2 | nov/2026, antes | 10 reais, soma 2718,0 | idem | 10 · 271,8 · 5,9777364724 · 2,1993144 | engPainel e Estatística: idem | engPainel 0; DP Estatística 9,6e-14 | APROVADO |
| Glicose | DIMENSION 1 | 8977 | N1 | nov/2026, 1º ATUALIZAR | + MAN_0001 79,468; MAN_0008 79,005; MAN_0007 79,777 | (775 + 238,25)/13 | 13 · 77,9423076923 · 1,1284025866 · 1,4477408 | engPainel e Estatística: 13 · 77,9423076923 · 1,1284025866 · 1,4477408 | engPainel 0; DP Estatística 5,0e-13 | APROVADO |
| Glicose | DIMENSION 1 | 8977 | N2 | nov/2026, 1º ATUALIZAR | + MAN_0005 **294,884**; MAN_0010 273,268; MAN_0014 275,781 | (2718 + 843,933)/13 | 13 · 273,9948461538 · 8,2150720006 · 2,9982579 | engPainel e Estatística: idem | engPainel 0; DP Estatística 6,2e-14 | APROVADO |
| Glicose | DIMENSION 1 | 8977 | N2 | nov/2026, 2º ATUALIZAR | + MAN_0017 275,279 (repetido corrigido) | 3837,212/14 | 14 · 274,0865714286 · 7,9002441901 · 2,8823901 | engPainel e Estatística: idem | engPainel 0; DP Estatística 1,2e-13 | APROVADO |
| WBC | XN1000 | 625311 | N1 | nov/2026, antes | 17 reais, soma 53,29 | idem | 17 · 3,1347058824 · 0,0721212215 · 2,3007333 | engPainel: idem | média 0; DP 0; CV 1,9e-16 | APROVADO |
| WBC | XN1000 | 625311 | N2 | nov/2026, antes | 23 reais, soma 168,45 | idem | 23 · 7,3239130435 · 0,1064612234 · 1,4536112 | engPainel: idem | média 2,4e-16; DP 0; CV 3,1e-16 | APROVADO |
| WBC | XN1000 | 625311 | N3 | nov/2026, antes | 12 reais, soma 209,56 | idem | 12 · 17,4633333333 · 0,2310581719 · 1,3231046 | engPainel: idem | 0 | APROVADO |
| WBC | XN1000 | 625311 | N1 | nov/2026, 1º ATUALIZAR | + MAN_0001 3,188; MAN_0008 3,164; MAN_0007 3,204 | (53,29 + 9,556)/20 | 20 · 3,1423 · 0,0690423973 · 2,1971931 | engPainel: idem | 0 | APROVADO |
| WBC | XN1000 | 625311 | N2 | nov/2026, 1º ATUALIZAR | + MAN_0005 **7,877**; MAN_0014 7,357 | (168,45 + 15,234)/25 | 25 · 7,34736 · 0,1503611985 · 2,0464656 | engPainel: idem | média 1,2e-16; DP 0 | APROVADO |
| WBC | XN1000 | 625311 | N3 | nov/2026, 1º ATUALIZAR | + MAN_0010 17,364 | (209,56 + 17,364)/13 | 13 · 17,4556923077 · 0,2229302524 · 1,2771207 | engPainel: idem | 0 | APROVADO |
| WBC | XN1000 | 625311 | N2 | nov/2026, 2º ATUALIZAR | + MAN_0017 7,344 (repetido corrigido) | 191,028/26 | 26 · 7,3472307692 · 0,1473247590 · 2,0051740 | engPainel: idem | média 2,4e-16; DP 0; CV 2,2e-16 | APROVADO |
| WBC | XN1000 | 625311 | N1 | ano 2026 (Estatística), antes | 38 reais | idem | 38 · 3,1268421053 · 0,0623462547 | Estatística: idem | 0; DP 1,1e-16 | APROVADO |
| WBC | XN1000 | 625311 | N2 | ano 2026 (Estatística), antes | 62 reais | idem | 62 · 7,3322580645 · 0,1102964612 | Estatística: idem | média 1,2e-16; DP 0 | APROVADO |
| WBC | XN1000 | 625311 | N3 | ano 2026 (Estatística), antes | 33 reais | idem | 33 · 17,4124242424 · 0,1997972457 | Estatística: idem | média 2,0e-16; DP 1,4e-16 | APROVADO |
| WBC | XN1000 | 625311 | N1 | ano 2026 (Estatística), 1º ATUALIZAR | + os 3 do N1 acima | (118,82 + 9,556)/41 | 41 · 3,1311219512 · 0,0620774497 | Estatística: idem | 0 | APROVADO |
| WBC | XN1000 | 625311 | N2 | ano 2026 (Estatística), 1º ATUALIZAR | + 7,877; 7,357 | (454,6 + 15,234)/64 | 64 · 7,34115625 · 0,1281351245 | Estatística: idem | média 2,4e-16; DP 0 | APROVADO |
| WBC | XN1000 | 625311 | N3 | ano 2026 (Estatística), 1º ATUALIZAR | + 17,364 | (574,61 + 17,364)/34 | 34 · 17,411 · 0,1969219196 | Estatística: idem | média 2,0e-16; DP 0 | APROVADO |
| WBC | XN1000 | 625311 | N2 | ano 2026 (Estatística), 2º ATUALIZAR | + 7,344 | 477,178/65 | 65 · 7,3412 · 0,1271306169 | Estatística: idem | média 2,4e-16; DP 2,2e-16 | APROVADO |
| Glicose (Westgard) | DIMENSION 1 | 8977 | N2 | corrida 26111401 | MAN_0005 294,884; alvo do lote 274,7758621 ± 5,0269933 | z = (x − média do alvo)/DP do alvo; 1-3s se \|z\| > 3 | z = 4,00003 → 1-3s; demais pontos novos \|z\| < 0,5 → nenhuma outra marca | Eng_Saida slot 126: marca 1_3s = 1 no N2, só ela; Painel M8 0 → 1, J8 REJEITADO, S8 "1_3s · RUN 26111401" | nenhuma | APROVADO |
| WBC (Westgard) | XN1000 | 625311 | N2 | corrida 26111401 | MAN_0005 7,877; alvo 7,33 ± 0,1368; N1 da mesma corrida z 0,10 | idem; R-4s se amplitude entre níveis > 4; 2of3-2s | z = 3,9985 → 1-3s; amplitude 3,90 < 4 → sem R-4s; só 1 nível > 2 DP → sem 2of3-2s; demais \|z\| ≤ 0,30 | Eng_Saida slot 68: 1_3s = 1 no N2, só ela; J8 REJEITADO; Audit "N2 z=4,00" | nenhuma | APROVADO |

**Valores que não podem entrar e não entraram** (nos dois produtos): os CONFLITO por chave repetida (MAN_0002, 0003,
0004, 0006, 0011, 0012, 0013); os `MAN_INVALIDO_L13/L14` (Bio 79,391 e 79,237; Hema 3,184 e 3,176); o repetido com
eventos desligados no 1º ATUALIZAR (Bio 275,279; Hema 7,344). O n bate exatamente com a contagem dos ATIVOS.

**Base (entrega 42):** os números também batem com o oráculo, mas sem os resultados colados (Bio N1 e N2 com n 11;
Hema N1 18, N2 24, N3 12). O defeito da OC-0073 era de entrada, não de cálculo.

### B.2 LJ (Eng_Saida, a fonte do gráfico) depois do 2º ATUALIZAR

| Produto | Slot | Corrida | N1 (ID · valor) | N2 (ID · valor) | N3 | Marcas |
|---|---|---|---|---|---|---|
| Bio | 126 | 26111401 | MAN_0001 · 79,468 | MAN_0005 · 294,884 | — | N2 **1_3s** |
| Bio | 127 | 26111402 | MAN_0008 · 79,005 | — | — | — |
| Bio | 128 | 26111403 | — | MAN_0010 · 273,268 | — | — |
| Bio | 129 | 26111404 | — | MAN_0014 · 275,781 | — | — |
| Bio | 130 | 26111405 | — | MAN_0017 · 275,279 (só no 2º) | — | — |
| Bio | 131 | 26111406 | MAN_0007 · 79,777 | — | — | — |
| Hema | 68 | 26111401 | MAN_0001 · 3,188 · lote 39 | MAN_0005 · 7,877 · lote 63 | — | N2 **1_3s** |
| Hema | 69 | 26111402 | MAN_0008 · 3,164 · lote 40 | — | — | — |
| Hema | 70 | 26111403 | — | — | MAN_0010 · 17,364 · lote 34 | — |
| Hema | 71 | 26111404 | — | MAN_0014 · 7,357 · lote 64 | — | — |
| Hema | 72 | 26111405 | — | MAN_0017 · 7,344 · lote 65 (só no 2º) | — | — |
| Hema | 73 | 26111406 | MAN_0007 · 3,204 · lote 41 | — | — | — |

- O alvo do lote está em cada ponto (Bio N1 79,3140496 ± 1,5427784, N2 274,7758621 ± 5,0269933; Hema N1 3,18 ±
  0,07935, N2 7,33 ± 0,1368, N3 17,44 ± 0,25395).
- Na Hema, a corrida do lote continua a sequência real (N1 38 → 39, N2 62 → 63, N3 33 → 34), como manda o ADR-080.
- O MAN_0007 passou de 26111405 para 26111406 quando o MAN_0017 (15:00) entrou antes dele no dia: as corridas
  seguintes do dia andam +1, como diz o GUIA.

### B.3 Etapas para reproduzir e conferência do consolidado

1. Copiar a produção atual e o backup da entrega 42 para `%TEMP%\qcw\auditoria\<ciclo>\` e conferir o sha256: Bio
   bd5314426ab8 e b20fe00f7e51; Hema 49f68fba99d5 e 7bab8fc008f6.
2. Rodar, uma cópia por vez: `reteste06_usuario_f1.py nova`, `... base`, `reteste06_carimbo_residual.py nova`,
   `... base` (Bio) e `reteste06_usuario_f1_hema.py nova|base|carimbo|carimbo_base` (Hema).
3. Oráculo: filtrar a tblCQ_Final por PARTICIPA_ESTATISTICA = "SIM", ANALITO, EQUIPAMENTO, LOTE, NIVEL e DATA_HORA
   (nov/2026 para o Painel; nov/2026 na Bio e ano 2026 na Hema para a Estatística); `statistics.fmean`,
   `statistics.stdev`, CV = 100·DP/média.
4. Conta à mão: n₁ = n₀ + k; média₁ = (Σ₀ + Σ novos)/n₁, com Σ₀ = n₀·média₀.

**O que este consolidado conferiu de novo (sem Excel):**
- Somas e médias, refeitas à mão: Bio 1013,25/13 = 77,9423077; 3561,933/13 = 273,9948462; 3837,212/14 = 274,0865714.
  Hema 62,846/20 = 3,1423; 183,684/25 = 7,34736; 191,028/26 = 7,3472308; 226,924/13 = 17,4556923; ano 2026:
  128,376/41 = 3,1311220; 469,834/64 = 7,3411563; 477,178/65 = 7,3412; 591,974/34 = 17,411. Todas iguais às
  publicadas.
- Diferença relativa oráculo × `engPainel`, lida dos JSON (`antes`, `depois1`, `depois2`, todos os níveis): n idêntico;
  média, DP e CV com diferença 0 na Bio e no máximo 3,1e-16 na Hema.
- z do Westgard: (294,884 − 274,7758621)/5,0269933 = 4,00003; (7,877 − 7,33)/0,1368 = 3,99854.
- Os números da aba Estatística (Bio e Hema) foram tomados dos relatórios das frentes; não os reli nos JSON.

### B.4 Caminho do dado (conferido neste ciclo)

- **Manual até o Painel e o LJ.** Digitar Resultados (colar ou digitar) → `Worksheet_Change`/`EntradaMudou` (ID novo
  quando o ID já existe, aviso, carimbo USUARIO/REGISTRADO_EM) → ATUALIZAR (`mEtlCQ`: chave, status, RUN; ID repetido
  → `~L<n>` CONFLITO; chave e instante repetidos → CONFLITO; ID fora do padrão → `MAN_INVALIDO_L<n>`) → tblCQ_Final
  (PARTICIPA = SIM) → série equipamento × analito × lote × nível × período → Eng_Saida (pontos, alvo, corrida do lote,
  marcas) → engPainel e Painel (n, média, DP, CV, status do Westgard) e aba Estatística.
- **Trilha.** Audit_Log RESULTADO_MANUAL_INCLUIDO / RETIRADO e WESTGARD_EVENTO_CORRIDA_NOVA; QA_INTEGRACAO A01 e E04.
  O "lançado por … em …" vem do carimbo da linha (por isso a P-01 também afeta a trilha).
- **CEQ → bias → ET → Sigma e CIQ → CV → ET/Sigma:** não exercitados neste ciclo.

### B.5 Diferenças de método e observações (não reprovadas)

1. **Bias, ET e Sigma fora do escopo.** As colunas G:I do Painel (A7:T8/T9) e as colunas equivalentes do `engPainel`
   mudam com os manuais (o CV entra no ET e no Sigma), mas nenhuma frente as conferiu. O `engPainel` traz um bias e um
   Sigma próprios, diferentes dos mostrados no Painel (o "Sigma oculto" já descrito no `parecer_cq` do ciclo
   2026-10-31_ciclo01). Não é ocorrência nova; fica para a F3.
2. **Estatística da Bio com 5e-13 no DP.** É acima do 1e-16 do `engPainel`, mas muito abaixo da tolerância de 1e-9:
   ordem de grandeza compatível com uma fórmula de planilha (soma de quadrados) contra o `stdev` do Python.
3. **Hema, ano da Estatística.** Os números da Estatística só foram comparados com B3/D3 = 2026, digitados pela frente.
   Com o ano que vem no arquivo (2025), a aba mostra n 0 para o lote 625311 (P-05): não é erro de cálculo, é filtro de
   ano sem dado do lote.

---

## Relatório C: ocorrências

**Deduplicação.** As frentes trouxeram 9 ocorrências; ficaram 5:
- R6-01 (Bio) e R6-01 Hema → **P-01**.
- R6-02 (Bio) e R6-02 Hema → **P-02**.
- R6-03 (Bio) e R6-03 Hema → **P-03**.
- R6-04 (Bio) e R6-04 Hema → **P-04**.
- R6H-05 (Hema) → **P-05**.

Os IDs P-xx são provisórios; o ID definitivo (OC-xxxx) é dado pelo Orquestrador. Nenhuma é Crítica. **Nenhuma é efeito
colateral da entrega 43:** P-01 e P-02 foram reproduzidas na base dos dois produtos; P-03 e P-04 são idênticas na
base; P-05 é estado do arquivo (Ano De/Até gravado), sem relação com o ADR-096, e não foi comparada com a base.

### C.1 Tabela

| ID | Produto | Prioridade | Tipo | Tela/módulo | Resumo | Origem | Relação | Base (entrega 42) |
|---|---|---|---|---|---|---|---|---|
| P-01 | Ambos | **Alto** | `funcional` | Digitar Resultados (`mIntegracao.PrepararLinha`/`Carimbar`, `EntradaMudou`) + motor (`mEtlCQ`, regra "DATA/HORA posterior ao lançamento") | Apagar ou mover uma linha manual deixa USUARIO e REGISTRADO_EM; o próximo manual digitado nela herda o carimbo, é recusado (MANUAL_INCOMPLETO, E04) e a trilha o atribui ao usuário antigo | Bio R6-01; Hema R6-01 | vizinho da OC-0073 (o ADR-096 manda "apague a linha") | Igual nos dois |
| P-02 | Ambos | Médio | `usabilidade` | Digitar Resultados (proteção, L:P travadas) / GUIA §3 / ADR-096 | "Copiar uma linha e colar embaixo": a linha inteira é recusada com mensagem que fala em senha; só o bloco ID..MOTIVO funciona | Bio R6-02; Hema R6-02 | vizinho da OC-0073 | Igual nos dois |
| P-03 | Ambos | Médio | `usabilidade` | Digitar Resultados (XLOOKUP de STATUS/RUN/DETALHE pelo ID) e mensagem do ATUALIZAR | Manual com ID fora do padrão mostra STATUS "—" e DETALHE "atualize os dados" depois do ATUALIZAR, enquanto a mensagem manda ver o DETALHE | Bio R6-03; Hema R6-03 | — | Igual nos dois |
| P-04 | Ambos | Baixo | `usabilidade` | Aviso do `EntradaMudou` (ADR-096) e `mEtlQA.TextoA01` (ramo "resultado manual repetido") | Linha colada trocando só o resultado fica CONFLITO por chave e instante; o aviso sugere que trocar o resultado basta; o A01 manda "inative um deles"; a mensagem do ATUALIZAR não conta | Bio R6-04; Hema R6-04 | resíduo de texto do mesmo tropeço que a OC-0073 corrigiu no ramo do ID | Igual nos dois |
| P-05 | Hematologia | Médio | `usabilidade` | Aba Estatística (B3/D3 Ano De/Até, B4 lote em análise) | Abre com 2025; o lote em uso (625311, 2026) aparece com n 0, média e DP vazios, sem aviso | Hema R6H-05 | ligada à OC-0048 (aberta) | Não comparado |

### C.2 Detalhe

**P-01 (Ambos, Alto, funcional): o carimbo de uma linha apagada fica e recusa o próximo manual digitado nela**
- **Tela/módulo:** aba Digitar Resultados (`mIntegracao.PrepararLinha`/`Carimbar`, `EntradaMudou`) e motor (`mEtlCQ`,
  teste "DATA/HORA posterior ao lançamento").
- **Ação:**
  1. Digitar o manual A (14/11 07:40). Carimbo: Bio 21:18:30; Hema 21:47:02.
  2. Selecionar ID..MOTIVO (A:K), o que o usuário consegue selecionar, e apagar.
  3. Esperar 2,5 min e clicar no botão Digitar Resultados: o cursor vai para B7, a linha apagada.
  4. Digitar ali o manual B (15/11, depois do carimbo de A).
  5. Controle: o manual C igual, numa linha nunca usada.
  6. ATUALIZAR DADOS.
- **Esperado:** B ganha carimbo próprio (usuário atual, hora atual) e fica ATIVO, como o C.
- **Observado:**
  - Depois do Delete, a linha mantém USUARIO e REGISTRADO_EM. O `Carimbar` só preenche o vazio, e B herda o carimbo.
  - B fica **MANUAL_INCOMPLETO**, DETALHE "campos obrigatórios ausentes/inválidos: DATA/HORA posterior ao lançamento",
    E04 na QA; a mensagem conta "1 resultado(s) manual(is) incompleto(s)".
  - O controle C fica ATIVO.
  - O mesmo resíduo fica na origem de um recorte (T06) e na linha colada apagada (T14).
  - Igual na base dos dois produtos.
- **Dificuldade:** o usuário não tem como corrigir: as colunas do carimbo são travadas e o DETALHE aponta a data/hora,
  que está certa. A única saída é usar outra linha, e nada diz isso.
- **Consequência:**
  - No uso real (apagar hoje e digitar dias depois na mesma linha, que é para onde o botão leva), todo manual digitado
    nessa linha com data/hora posterior ao carimbo é recusado: fica fora da estatística, do LJ e do Westgard.
  - A trilha atribui o lançamento a outro usuário e a outro momento (Audit_Log "manual lancado por <usuário antigo> em
    <data antiga>").
  - O ADR-096, o A01 e a mensagem do ATUALIZAR agora mandam "apague … a linha", o que aumenta a exposição.
- **Sugestão:** limpar USUARIO e REGISTRADO_EM no `EntradaMudou` quando ID..MOTIVO e SUBSTITUI_ID ficarem vazios (como a
  Inativar no ADR-070); limpar no `PrepararEntradas` o carimbo das linhas sem dado; opcionalmente, carimbar de novo
  quando a linha passa de vazia para preenchida. Decisão do Orquestrador.
- **Evidência:** `%TEMP%\qcw\auditoria\2026-11-15_reteste06\Bioquimica_F1_carimbo.json`, `Bioquimica_F1_carimbo_base.json`,
  `Hematologia_F1_carimbo.json`, `Hematologia_F1_carimbo_base.json` (A_apagado, botao_parou_em, B_digitado,
  atualizar). Scripts `ferramentas/reteste06_carimbo_residual.py` e `ferramentas/reteste06_usuario_f1_hema.py carimbo`.
  Produções bd5314426ab8 e 49f68fba99d5; bases b20fe00f7e51 e 7bab8fc008f6.
- **Não executado:** o intervalo de dias (reproduzido com 2,5 min; o mecanismo é o mesmo).

**P-02 (Ambos, Médio, usabilidade): copiar a linha inteira, como o GUIA §3 sugere, é recusado pela proteção**
- **Tela/módulo:** aba Digitar Resultados (proteção; USUARIO, REGISTRADO_EM, RUN, STATUS e DETALHE travados em L:P);
  GUIA §3; ADR-096.
- **Ação:** copiar a linha inteira da tabela (A:Q), ou pelo número da linha, e colar na próxima linha livre, com a
  proteção como o teclado do usuário a vê.
- **Esperado:** colagem aceita, com ID novo e aviso, como o GUIA promete.
- **Observado:** o Excel recusa e não cola nada: "A célula ou gráfico que você está tentando alterar está em uma
  planilha protegida… Pode ser necessário inserir uma senha." O bloco A:K é aceito. Pela COM (caminho de automação) a
  linha inteira é aceita e traz o carimbo da origem (T15). Igual na base.
- **Consequência:** o usuário desiste do atalho ou acha que precisa de senha. Sem risco para os dados.
- **Sugestão:** GUIA §3 e ADR-096: "selecione de ID até MOTIVO (as colunas brancas), copie e cole na próxima linha";
  opcionalmente, uma dica na célula do ID.
- **Evidência:** `Bioquimica_F1_nova.json`, `Hematologia_F1_nova.json` e as bases, passo `P1c_protecao_do_usuario`
  (`linha_inteira` RECUSADA, `bloco_AK` ACEITO).
- **Não executado:** o Ctrl+V físico (emulação por `Protect` sem `UserInterfaceOnly`).

**P-03 (Ambos, Médio, usabilidade): manual com ID inválido mostra "atualize os dados" no DETALHE depois do ATUALIZAR**
- **Tela/módulo:** aba Digitar Resultados (STATUS, RUN e DETALHE por XLOOKUP do ID) e mensagem do ATUALIZAR.
- **Ação:** duas linhas completas com ID "TESTEAUD9" (fora do padrão MAN_); ATUALIZAR (duas vezes na Hema).
- **Esperado:** a linha diz por que não entrou.
- **Observado:** o motor recusa corretamente (`MAN_INVALIDO_L13/L14`, MANUAL_INCOMPLETO, "ID_REGISTRO (padrao
  MAN_0001)", E04) e a mensagem diz "2 resultado(s) manual(is) incompleto(s): veja o DETALHE". Na tela, STATUS e RUN
  "—" e DETALHE "atualize os dados": o XLOOKUP procura "TESTEAUD9", e a tblCQ_Final tem `MAN_INVALIDO_L13`. Igual na
  base.
- **Consequência:** o usuário atualiza de novo, em vão, e não descobre que o problema é o ID. O manual fica fora da
  estatística.
- **Sugestão:** DETALHE cair em `MAN_INVALIDO_L<linha>` quando o ID não for MAN_ válido, ou mostrar "ID fora do padrão
  MAN_0001: apague o ID (um novo é gerado)"; ou a mensagem do ATUALIZAR citar o E04 de ID.
- **Evidência:** `*_F1_nova.json`, `atualizar1.tela` e `atualizar2.tela` (P6a/P6b), `tblCQ_Final`, `qa_integracao`; o
  mesmo nas bases.

**P-04 (Ambos, Baixo, usabilidade): repetido por chave e instante; aviso e A01 induzem ao erro**
- **Tela/módulo:** aviso do `EntradaMudou` (ADR-096) e `mEtlQA.TextoA01`, ramo "resultado manual repetido (mesma chave e
  instante)".
- **Ação:** colar a origem e trocar só o RESULTADO (Bio 78,697; Hema 3,148); ATUALIZAR.
- **Esperado:** pelo aviso ("Mude o que for diferente (nível, hora, resultado)"), trocar o resultado bastaria.
- **Observado:** CONFLITO_MANUAL "resultado manual repetido (mesma chave e instante)", porque a chave não inclui o valor
  (regra do motor; a tela mostra isso certo). O A01 dessa linha e das coladas sem edição diz "-> MAN_0001 … confirme e
  inative um deles ou corrija o manual". A mensagem do ATUALIZAR não conta esses CONFLITO. Igual na base.
- **Consequência:** quem repete o controle no mesmo minuto trocando só o valor perde o resultado; "inative um deles"
  pode levar a inativar o MAN_0001, que é o que vale (efeito não testado).
- **Sugestão:** aviso: "mude a hora ou o nível; só trocar o resultado deixa a linha como repetida"; A01 desse ramo:
  "apague a linha repetida <ID>; não inative o <ID que vale>".
- **Evidência:** `*_F1_nova.json`, passo `P2c_so_resultado` e `atualizar1.qa_integracao` (A01 do MAN_0006).

**P-05 (Hema, Médio, usabilidade): aba Estatística abre em 2025 e mostra n 0 para o lote em uso**
- **Tela/módulo:** aba Estatística da Hematologia (B3/D3 Ano De/Até; B4 lote em análise).
- **Ação:** lançar e atualizar os manuais; abrir a Estatística sem mexer no ano; depois pôr 2026 em B3 e D3.
- **Esperado:** ver o lote em uso (625311, de 2026) com o manual novo, ou um aviso de que o ano não tem dado do lote.
- **Observado:** B3 = D3 = 2025 no arquivo de produção. WBC N1 a N3: n 0, média e DP vazios, antes e depois dos
  manuais. Com 2026, os números aparecem e batem com o oráculo. O K5 diz "Bias do EP: CAP | ano 2025" (ano do CEQ,
  filtro separado).
- **Consequência:** quem confere o lançamento pela Estatística acha que o resultado não entrou ou que o lote não tem
  estatística. Sem erro de cálculo.
- **Sugestão:** iniciar B3/D3 no ano do lote em análise ou no corrente; ou, com n 0 para o lote, mostrar "o lote <B4>
  não tem resultado em <ano>; ajuste o Ano De/Até". Tratar junto da OC-0048.
- **Evidência:** `Hematologia_F1_nova_exec1.json` (`estatistica_cab` B3/D3 = 2025; `comparacao.estatistica` = [0, "",
  ""] nos 3 níveis) e `Hematologia_F1_nova.json` (`estatistica_cab.antes_de_digitar_o_ano` = 2025).
- **Não executado:** a mesma leitura na base da Hema; o estado da Bio (a Bio usa janela de datas, não Ano De/Até).

### C.3 Dúvidas de regra de negócio (para o Orquestrador encaminhar)

| ID | Produto | Para | Dúvida |
|---|---|---|---|
| D-01 | Ambos | Orquestrador / Especialista em CQ | (P-01) O REGISTRADO_EM deve ser o momento em que **esta** linha recebeu dados, ou o primeiro carimbo da célula? O motor o usa como "lançamento" na regra "DATA/HORA posterior ao lançamento", e o Audit_Log como "manual lançado por … em …". A leitura das duas frentes é que tem de ser o da linha atual |
| D-02 | Ambos | Especialista em CQ | (P-04) Dois manuais com a mesma chave e o mesmo instante, com valores diferentes, devem continuar como "repetido" (vale o primeiro)? Na Hema isso pesa mais, porque o mesmo controle às vezes é repassado em seguida. Se for intencional, o aviso do ADR-096 precisa dizer que trocar só o resultado não basta |
| D-03 | Hematologia | Orquestrador / UX | (P-05) Qual ano a aba Estatística da Hema deve mostrar ao abrir: o do lote em análise, o corrente ou o último usado? |

---

## Relatório D: parecer final

Critério: "validado" só onde houve teste executado e aprovado neste ciclo. O resto é "não validado" (teste reprovado)
ou "não executado" (com o motivo). Este ciclo foi o reteste de 1 item (OC-0073), só com a F1 e os fluxos vizinhos.

### D.1 A rotina funciona?

**Sim, no que foi refeito: a OC-0073 está RESOLVIDA nos dois produtos, sem efeito colateral.** Os fluxos vizinhos
acharam um risco operacional anterior à entrega 43 (P-01, Alto).

**Validado:**
- colar ou digitar um ID MAN_ existente gera ID novo na hora, com aviso claro (A:K, bloco de 3 com uma caixa, ID
  digitado, "man_1");
- o colado e editado entra no Painel, na Estatística, no LJ e no Westgard (1-3s), e a origem não muda;
- mover mantém o ID; o ID inválido repetido continua recusado como antes;
- o repetido criado com os eventos desligados é apontado pela mensagem do ATUALIZAR e pelo A01 novo, e apagar o ID
  resolve em 1 passo;
- a base (entrega 42) reproduz o defeito com o mesmo roteiro, o que prova que o teste o pega.

**Não validado:**
- reaproveitar uma linha apagada (T13, P-01, Alto): o manual novo é recusado;
- copiar a linha inteira com a proteção do usuário (T12, P-02, Médio);
- a Estatística da Hema com o ano que vem no arquivo (T17, P-05, Médio).

**Não executado:** teclado e mouse físicos (Ctrl+V real, clique no cabeçalho da linha); o intervalo de dias da P-01.

### D.2 Bias, ET e Sigma estão validados?

**Não executado neste ciclo** (não houve F3 nem F2). Os números de bias, ET e Sigma do Painel mudaram com os manuais
(o CV entra neles), mas nenhum foi recalculado de forma independente. A última validação independente é a do reteste 04
(entrega 41).

### D.3 A integração CIQ + CEQ está validada?

**Não executado neste ciclo.** Nenhum CEQ foi incluído nem corrigido, e nenhuma leitura de bias foi conferida.

### D.4 Os filtros de período e de lote estão corretos?

**Validado, no escopo executado:**
- Painel com lote no H3 e mês (nov/2026): as séries Glicose DIMENSION 1 lote 8977 (N1, N2) e WBC XN1000 lote 625311
  (N1 a N3) recebem só os ATIVOS do lote e do mês, com n, média, DP e CV iguais ao oráculo;
- aba Estatística da Bio (janela 01–30/11/2026) e da Hema (ano 2026, lote 625311): iguais ao oráculo.

**Não validado:** a Estatística da Hema com o Ano De/Até do arquivo (2025) mostra n 0 para o lote em uso (P-05; não é
erro de cálculo).

**Não executado:** mês a mês além de novembro, trimestres, troca de lote, H3 vazio (todos os lotes), outros
equipamentos. A F4 não fez parte deste reteste.

### D.5 A entrada manual e a automática estão validadas?

**Entrada manual: validada nos cenários da OC-0073** (colar, colar e editar, bloco, ID digitado, mover, ID inválido,
repetido com eventos desligados e correção, apagar a linha colada), nos dois produtos. Status, RUN, tblCQ_Final,
Audit_Log, A01 e E04 conferidos; números iguais ao oráculo.

**Não validado na entrada manual:**
- digitar numa linha apagada (P-01, Alto): o manual válido é recusado e a trilha fica errada;
- colar a linha inteira pelo teclado (P-02);
- a tela do ID inválido (P-03) e o aviso/A01 do repetido por chave (P-04).

**Entrada automática (DB_SEAC → tblCQ_Final): não executada.** O ATUALIZAR rodou, mas não havia resultado novo no
DB_SEAC (últimos de 13/11/2026).

**Equivalência manual × automático:** não executada.

### D.6 Há divergência matemática?

**Não**, nos cálculos executados.
- Relatório B: 21 linhas, 21 aprovadas, 0 reprovadas (19 de n/média/DP/CV e 2 de Westgard).
- Maior diferença relativa: 5,0e-13 (DP da aba Estatística da Bio); no `engPainel`, 0 na Bio e no máximo 3,1e-16 na
  Hema. Todas abaixo de 1e-9.
- Valores que não podem entrar (CONFLITO, `MAN_INVALIDO_L*`, repetido antes da correção) não entraram.
- Bias, ET e Sigma não foram recalculados (D.2).

### D.7 Há problema operacional?

**Sim, um relevante, anterior à entrega 43: a P-01 (Alto).** Apagar uma linha manual e digitar outro manual nela (é
para onde o botão leva) faz o resultado válido ser recusado, sem que o usuário consiga corrigir, e atribui o lançamento
a outro usuário e momento. O ADR-096 manda "apague a linha", o que aumenta a exposição. Os números ficam certos (o
recusado simplesmente não entra), mas o resultado de controle fica fora do CQ.

Também: P-02, P-03 e P-05 (Médio) e P-04 (Baixo), de usabilidade. As produções ficaram intactas: mesmo hash antes e
depois de todas as execuções.

### D.8 O que fica pendente

1. **P-01 (Alto):** decidir e corrigir o carimbo residual (com a D-01). No reteste, refazer o roteiro do
   `reteste06_carimbo_residual.py` nos dois produtos, mais o recorte (T06) e a linha colada apagada (T14).
2. **P-02, P-03, P-05 (Médio) e P-04 (Baixo):** decidir e corrigir (P-02 é texto do GUIA/ADR-096; P-05 junto da
   OC-0048).
3. **Dúvidas:** encaminhar D-01 (o que é o REGISTRADO_EM), D-02 (repetido por chave com valor diferente) e D-03 (ano
   inicial da Estatística da Hema).
4. **OC da entrega 43 que seguem em AGUARDA RETESTE e não foram retestadas neste ciclo:** OC-0074, OC-0075 e OC-0076.
5. **Seguem abertas, de ciclos anteriores:** entre outras, OC-0045, OC-0046, OC-0047, OC-0048, OC-0064, OC-0069,
   OC-0070, OC-0077 e as dúvidas do reteste 05 (formato 14, RUN do IMPLAUSIVEL, texto colado por cima da validação,
   "Outro" sem texto livre).
6. **Não executados, a cobrir num próximo ciclo:**
   - F3 sobre a entrega 43 (bias, ET e Sigma; ETp ausente; ETp VB/FAB);
   - F2 (CEQ);
   - recebimento automático com resultado novo do DB_SEAC e equivalência manual × automático;
   - F4 completa (mês a mês, trimestres, troca de lote, H3 vazio, LJ e Westgard em outros cenários);
   - teclado e mouse físicos, que pedem uma sessão interativa.

Dados pessoais: nenhum login de pessoa, máquina ou caminho de rede consta deste parecer. Os usuários citados
(`TESTE_AUDITORIA_R6*`) são de teste, criados só nas cópias.
