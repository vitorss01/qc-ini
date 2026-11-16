# Reteste 06: Usuário Simulado, Hematologia, frente F1

Data: 15/11/2026 · Agente: AGENTE_USUARIO_SIMULADO · Modo: RETESTE da OC-0073 · Produção: entrega 43 (ADR-096).

## Ambiente e método

| Item | Valor |
|---|---|
| Produção atual (só copiada) | `QC_Hematologia.xlsm`, sha256 curto **49f68fba99d5**. Ficou igual antes e depois das 3 execuções na produção (`origem_intacta.igual = true`) e foi conferida de novo no fim |
| Base de comparação ("antes") | `_backup_pre_integracao_2026-11-15_2059/QC_Hematologia.xlsm`, sha256 curto **7bab8fc008f6** (entrega 42), igual antes e depois |
| Cópias (novas, nenhuma salva, todas apagadas no fim) | `%TEMP%\qcw\auditoria\2026-11-15_reteste06\Hematologia_F1_nova.xlsm` (2 execuções), `Hematologia_F1_base.xlsm`, `Hematologia_F1_carimbo.xlsm` e `Hematologia_F1_carimbo_base.xlsm` |
| Excel | Um por vez, abertos e encerrados pelo `qa_final.QA`. Nova: PID 29952 (1ª execução) e 9308 (2ª). Base: 21892. Carimbo: 28536 (nova) e 23340 (base). Não toquei no 13084 nem nos Excel da tarefa agendada (por exemplo, o 8640, aberto às 21:54:06). Nenhuma captura de tela |
| Script (novo, documentado no cabeçalho) | `_arquitetura/agentes/ferramentas/reteste06_usuario_f1_hema.py <nova\|base\|carimbo\|carimbo_base>`. É o mesmo roteiro do irmão da Bio (`reteste06_usuario_f1.py` e `reteste06_carimbo_residual.py`), com o que muda na Hema: 3 níveis, XN1000 sem L4, coluna MATRIZ, Painel A7:T9, Estatística por ano (B3/D3) e Eng_Saida de 3 níveis |
| Saídas (fora do repositório) | `Hematologia_F1_{nova,nova_exec1,base,carimbo,carimbo_base}.json` e os `.log`, na pasta do ciclo no `%TEMP%` |
| Usuários | `TESTE_AUDITORIA_R6H` digitou a origem. `TESTE_AUDITORIA_R6HB` fez as colagens. `TESTE_AUDITORIA_R6HC` fez o teste do carimbo. Todos ANALISTA, criados só na cópia, com login por `mSeguranca.DoLogin` |
| Dados de teste | WBC, XN1000, lote **625311** (o lote corrente, com os 3 níveis em nov/2026), 14/11/2026 (o último dado real de WBC é de 13/11/2026 13:10:20), MOTIVO "Outro", MATRIZ preenchida sozinha (SANGUE TOTAL). Valores sintéticos de 3 decimais, calculados do alvo do lote (`mLotes.MediaDoLote`/`DPDoLote`: N1 3,18 ± 0,07935; N2 7,33 ± 0,1368; N3 17,44 ± 0,25395). N1: 3,188 / 3,148 / 3,204 / 3,164 / 3,184 / 3,176. N2: **7,877 (média + 4 DP, para provocar 1-3s)** / 7,357 / 7,344. N3: 17,364. ID inválido `TESTEAUD9`. Carimbo: 3,181 / 3,182 / 7,333 em 14 e 15/11 |
| tblCQ_Final | 48.391 linhas no início. Os ATUALIZAR leram o DB_SEAC só para leitura e não trouxeram resultado novo ("novos nesta atualização: 0") |

**Caminho do usuário:**
- Fiz o login. No Painel, escrevi C3, H3, K3 e K4 com os eventos ligados.
- Na aba Estatística, escrevi o Ano De/Até (B3/D3, as células destravadas) = 2026.
- Entrei pelo botão (`mIntegracao.IrDigitarResultados`, que parou em B7) e digitei célula a célula (`FormulaLocal`).
- Copiei e colei com `Range.Copy(destino)`, o equivalente ao Ctrl+C/Ctrl+V, e recortei com `Range.Cut(destino)`.
- Durante a digitação deixei `Application.Interactive = True`, para que o aviso do `EntradaMudou` aparecesse. As
  caixas foram lidas pelo texto da janela do meu PID e fechadas no OK.
- Rodei o **ATUALIZAR DADOS** pelo núcleo do botão (`AtualizarDadosAutomatico`).

O que não executei:
- o teclado e o mouse físicos (Ctrl+V de verdade e o clique no cabeçalho da linha);
- o cenário "dias depois" do R6-01, que reproduzi com 2,5 minutos de intervalo. O mecanismo é o mesmo.

**Leituras:**
- linha da tela (STATUS, RUN e DETALHE em `Text`);
- tblCQ_Final, tblQA_Integracao (A01 e E04) e tblAuditoria nova;
- Painel: engPainel e A7:T9 (contagens do Westgard e status);
- Eng_Saida: pontos do LJ por ID (AM..AO), corrida do lote (BK..BM), alvo (BW..CB) e marcas do Westgard (C..W);
- aba Estatística: WBC N1 a N3, com o lote do B4 = 625311.

**Oráculo próprio:** Python puro (`statistics`) sobre o `Value2` da tblCQ_Final. Ele pega PARTICIPA = SIM, WBC, XN1000
e lote 625311, e calcula n, média, DP amostral e CV em duas janelas: DATA_HORA em nov/2026 (Painel) e no ano de 2026
(aba Estatística). Também confere à mão a soma "antes + valores novos".

**Duas execuções na cópia nova.** A 1ª (PID 29952) leu a aba Estatística com o Ano De/Até que vem no arquivo: B3 =
D3 = 2025. Nesse ano o lote 625311 não tem nenhum dado, e o n deu 0 nos 3 níveis (ver R6H-05). Rodei de novo numa
cópia nova (PID 9308), agora pondo B3/D3 = 2026 como o usuário faria. O resto das duas execuções é idêntico: os mesmos
IDs, avisos, status, RUN e números do Painel. Os números deste relatório são os da 2ª execução.

## Resultado

| OC | Estado | Resumo |
|---|---|---|
| OC-0073 | **RESOLVIDA** | Toda linha colada com o ID de outra ganha um ID novo na hora, com um aviso claro. Isso vale para ID…MOTIVO (A:K), para a linha inteira pela COM, para o bloco de 3 e para o ID digitado, inclusive "man_1". Colar e editar fica ATIVO e entra no Painel, na aba Estatística, no LJ e no Westgard (1-3s acusado no N2). A origem não muda. Mover mantém o ID. O ID inválido repetido continua recusado, como na base. O repetido que entra com os eventos desligados é explicado pela mensagem do ATUALIZAR e pelo A01 novo, e apagar o ID resolve. Na base (entrega 42), o mesmo roteiro reproduz o defeito |

Os fluxos vizinhos **confirmam na Hematologia** as 4 ocorrências que a Bio achou neste ciclo (R6-01 a R6-04). Nenhuma
foi causada pela entrega 43: a R6-01 e a R6-02 também aparecem na base da Hema, e as outras duas são idênticas na
base. Achei também **1 observação nova (R6H-05)**: a aba Estatística da Hema vem com o ano de 2025.

## Testes executados

| ID | Teste | Estado | Evidência (nova = entrega 43; base = entrega 42) |
|---|---|---|---|
| OC-0073 | Reteste do fluxo original e dos critérios de aceite | **RESOLVIDA** | Itens T01 a T09 abaixo |
| T01 | Colar sem editar (ID…MOTIVO da origem na linha livre) | APROVADO | O MAN_0001 colado vira **MAN_0002** na hora. O aviso, com o título "Digitar Resultados", diz: "A linha trazia o ID MAN_0001, que já existe em outra linha. Ela recebeu o ID novo MAN_0002 (o ID é único e gerado pelo sistema). Mude o que for diferente (nível, hora, resultado) antes de clicar em ATUALIZAR DADOS." Depois do ATUALIZAR: CONFLITO_MANUAL, DETALHE "resultado manual repetido (mesma chave e instante)", RUN vazio. É a regra do motor (mesma chave e instante), e a tela mostra isso. A origem MAN_0001 fica ATIVO, RUN 26111401. Nenhum `~L` na tblCQ_Final. **Base:** a colada fica com MAN_0001, a tela mostra **ATIVO/26111401** e a tblCQ_Final tem `MAN_0001~L2` CONFLITO "vale a 1a linha" |
| T02 | Colar e editar (NIVEL 2, RESULTADO 7,877 = média + 4 DP) | APROVADO | ID **MAN_0005** logo depois de colar, e ele não muda com a edição. Fica ATIVO, PARTICIPA SIM, RUN 26111401 (a mesma corrida da origem, pela mesma hora). **LJ:** Eng_Saida, slot 68, corrida 26111401: N1 = MAN_0001 3,188 e N2 = **MAN_0005 7,877**, corrida do lote N2 = 63. **Westgard:** marca 1_3s = 1 no N2, com z = (7,877 − 7,33)/0,1368 = 3,9985. Painel N2: M8 vai de 0 a 1, J8 de OK a **REJEITADO**, S8 "1_3s · RUN 26111401", T8 "Erro Aleatório". O Audit_Log tem WESTGARD_EVENTO_CORRIDA_NOVA "1_3s INDIVIDUAL … N2 z=4,00". **Base:** o mesmo valor fica `MAN_0001~L4` CONFLITO. O N2 fica com n 24 sem ele, M8 = 0 e o status OK |
| T03 | A linha original não muda | APROVADO | A linha da origem, lida depois de cada colagem, é idêntica à de antes (`origem_igual = true` em P1 e P1b; `origem_intacta_antes_atualizar = true`). No fim: MAN_0001, 3,188, TESTE_AUDITORIA_R6H, REGISTRADO_EM 21:44:00, ATIVO, RUN 26111401 |
| T04 | Digitar um ID MAN_ que já existe | APROVADO | Numa linha vazia, com o ID digitado antes dos dados, "MAN_0001" vira **MAN_0008**, com aviso. Sobre uma linha preenchida (N3, ID automático MAN_0009), digitei "man_1" (minúsculo, sem zeros) e ele virou **MAN_0010**, com aviso. Os dois ficam ATIVOS depois do ATUALIZAR: MAN_0008 N1 com RUN 26111402 e MAN_0010 N3 com RUN 26111403. **Base:** os dois ficam MAN_0001, `~L8` e `~L9` CONFLITO, e o N3 fica com n 12, sem o 17,364 |
| T05 | Bloco de 3 linhas colado de uma vez | APROVADO | MAN_0001, 0002 e 0003 viram **MAN_0011, MAN_0012 e MAN_0013**, em sequência, com **uma** caixa: "As 3 linhas coladas traziam IDs que já existem em outras linhas e receberam IDs novos (MAN_0011 a MAN_0013)." A marca qcUltimoIdManual vai de 10 a 13. **Base:** as 3 ficam MAN_0001 |
| T06 | Mover (recortar ID…MOTIVO e colar 4 linhas abaixo) | APROVADO | O ID MAN_0014 se mantém, sem aviso. A origem ID…MOTIVO fica vazia. Depois do ATUALIZAR: ATIVO, RUN 26111404. Observação: USUARIO e REGISTRADO_EM ficam na linha de origem (R6-01) |
| T07 | ID fora do padrão MAN_ repetido ("TESTEAUD9" em 2 linhas) | APROVADO | Não é trocado. Na tblCQ_Final ficam `MAN_INVALIDO_L13` e `MAN_INVALIDO_L14`, MANUAL_INCOMPLETO, "ID_REGISTRO (padrao MAN_0001)", e a QA tem 2 **E04**. É **idêntico na base**. A tela, porém, não mostra o motivo (R6-03) |
| T08 | Repetido criado com os eventos desligados (dado antigo) e a correção | APROVADO | **1º ATUALIZAR:** `MAN_0001~L15` (N2, 15:00, 7,344) fica CONFLITO "vale a 1a linha". A mensagem traz "• 1 linha(s) da Digitar Resultados repete(m) o ID de outra (ex.: MAN_0001) e ficou(aram) fora da estatística. Apague o ID da linha repetida (um novo é gerado) e atualize de novo." O A01 diz "…apague o ID da linha repetida (um novo sera gerado) ou a linha. Nao inative o MAN_0001: isso tiraria o resultado que vale." Antes da correção, a tela dessa linha mostrava ATIVO/26111401, porque o XLOOKUP acha a 1ª linha. É o resíduo previsto no ADR-096, e a mensagem o cobre. **Correção:** ao apagar o ID, a linha recebeu **MAN_0017**. **2º ATUALIZAR:** ATIVO, RUN 26111405. Não há A01 nem a linha na mensagem, e o N2 vai de 25 para 26 no Painel e de 64 para 65 na Estatística. O Audit_Log grava RESULTADO_MANUAL_RETIRADO de `MAN_0001~L15` e RESULTADO_MANUAL_INCLUIDO na corrida 26111405. **Base:** a mensagem não cita o caso, o A01 manda "confirme e inative um deles" e, depois de apagar o ID, a linha vira MAN_0006 ATIVO |
| T09 | Painel, Estatística e LJ conferidos com o oráculo | APROVADO | Relatório B abaixo. n e média são exatos, a diferença relativa é ≤ 2,5e-16 e os pontos novos estão no Eng_Saida com o alvo do lote |
| T10 | Fluxo vizinho: colar e trocar só o RESULTADO (mesma hora e nível) | APROVADO (com R6-04) | Recebe ID novo (MAN_0006). Fica CONFLITO_MANUAL "resultado manual repetido (mesma chave e instante)", fora da estatística, pela regra do motor (a chave não inclui o valor). O aviso sugere que trocar o resultado basta, e o A01 diz "confirme e inative um deles … -> MAN_0001" (R6-04, igual à Bio) |
| T11 | Fluxo vizinho: colar e trocar HORA e RESULTADO | APROVADO | MAN_0007 fica ATIVO, com corrida própria: 26111405, e depois 26111406, quando o MAN_0017 das 15:00 entrou. As corridas seguintes do dia andam +1, como diz o GUIA |
| T12 | Fluxo vizinho: colar a **linha inteira** com a proteção que o usuário vê | REPROVADO (R6-02) | A aba é protegida, e USUARIO, REGISTRADO_EM, RUN, STATUS e DETALHE são travados (L:P). Com a proteção sem UserInterfaceOnly, que é a que o teclado do usuário obedece, colar a linha inteira dá: "A célula ou gráfico que você está tentando alterar está em uma planilha protegida… Pode ser necessário inserir uma senha." Nada é colado. O bloco ID…MOTIVO (A:K) é aceito. **Igual na base** |
| T13 | Fluxo vizinho: apagar uma linha manual (ID…MOTIVO) e digitar outro manual nela | REPROVADO (R6-01) | **Nova:** digitei A (14/11 07:40); o carimbo saiu às 21:47:02. Apaguei A:K, e USUARIO e REGISTRADO_EM ficaram. Depois de 2,5 min, o botão levou o cursor para **B7, a linha apagada**. Ali digitei B (15/11 21:48), às 21:49:37. B **herdou o REGISTRADO_EM 21:47:02** e ficou MANUAL_INCOMPLETO "DATA/HORA posterior ao lançamento", com um E04 na QA. O controle C (igual, N2, numa linha nunca usada, REGISTRADO_EM 21:49:42) ficou ATIVO. **Base:** o mesmo resultado (carimbo 21:51:23, B MANUAL_INCOMPLETO, C ATIVO) |
| T14 | Fluxo vizinho: apagar a linha colada (ID…MOTIVO) e atualizar | APROVADO | O MAN_0002 some da tblCQ_Final, e o Audit_Log grava RESULTADO_MANUAL_RETIRADO "ID MAN_0002". Nada mais muda. Restam USUARIO e REGISTRADO_EM na linha (R6-01) |
| T15 | Trilha: quem lançou a linha colada | APROVADO (A:K) / observação (linha inteira) | Na colagem de A:K, o USUARIO é o usuário 2 (R6HB) e o REGISTRADO_EM é o momento da colagem. Na colagem da linha inteira pela COM, o USUARIO e o REGISTRADO_EM da **origem** vêm junto: a linha MAN_0003 mostra TESTE_AUDITORIA_R6H 21:44:00, e o Audit_Log diz "manual lancado por TESTE_AUDITORIA_R6H em 15/11/2026 21:44:00", mas quem colou foi o R6HB. Pelo teclado isso não acontece, porque a colagem é recusada (T12). Igual à Bio |
| T16 | Base (entrega 42), mesmo roteiro | APROVADO (prova de que o teste pega o defeito) | Nenhum aviso e nenhum ID novo. 11 linhas `MAN_0001~L<n>` CONFLITO "vale a 1a linha" aparecem **ATIVO/26111401** na tela. Painel em nov/2026: N1 n 18, N2 n 24, N3 n 12 (nova: 20, 25 e 13). Não há 1-3s. A mensagem não cita o caso, e o A01 manda "confirme e inative um deles" |
| T17 | Aba Estatística com o ano que vem no arquivo | Observação (R6H-05) | B3 = D3 = **2025** (sem mexer). Com o lote 625311 (B4), WBC N1 a N3 = **0 / vazio / vazio**, antes e depois dos manuais. Com B3/D3 = 2026, os números batem com o oráculo (Relatório B) |

## Relatório B: conferência numérica (Painel e Estatística × oráculo)

Filtro do Painel: C3 WBC · H3 lote 625311 · K3 2026 · K4 Nov (G3/G4 = 01/11/2026–30/11/2026). Na aba Estatística:
B3 = D3 = 2026 e B4 = 625311. Os dados vêm do `Value2` da tblCQ_Final da cópia `Hematologia_F1_nova.xlsm` (produção
49f68fba99d5).

| Analito | Equip. | Lote | Nível | Período | Momento | Dados de entrada (novos ATIVOS) | Fórmula | Independente (n · média · DP) | QC_INI engPainel / Painel (n · média · DP · CV%) | Diferença relativa (média; DP) | Status |
|---|---|---|---|---|---|---|---|---|---|---|---|
| WBC | XN1000 | 625311 | N1 | nov/2026 | antes | 17 reais, soma 53,29 | média = Σx/n; DP amostral | 17 · 3,1347058824 · 0,0721212215 | 17 · 3,1347058824 · 0,0721212215 · 2,3007333 | 0; 0 | APROVADO |
| WBC | XN1000 | 625311 | N2 | nov/2026 | antes | 23 reais, soma 168,45 | idem | 23 · 7,3239130435 · 0,1064612234 | 23 · 7,3239130435 · 0,1064612234 · 1,4536112 | 2,4e-16; 0 | APROVADO |
| WBC | XN1000 | 625311 | N3 | nov/2026 | antes | 12 reais, soma 209,56 | idem | 12 · 17,4633333333 · 0,2310581719 | 12 · 17,4633333333 · 0,2310581719 · 1,3231046 | 0; 0 | APROVADO |
| WBC | XN1000 | 625311 | N1 | nov/2026 | 1º ATUALIZAR | + MAN_0001 3,188; MAN_0008 3,164; MAN_0007 3,204 | (53,29 + 9,556)/20 | 20 · **3,1423** · 0,0690423973 | 20 · 3,1423 · 0,0690423973 · 2,1971931 | 0; 0 | APROVADO |
| WBC | XN1000 | 625311 | N2 | nov/2026 | 1º ATUALIZAR | + MAN_0005 **7,877**; MAN_0014 7,357 | (168,45 + 15,234)/25 | 25 · **7,34736** · 0,1503611985 | 25 · 7,34736 · 0,1503611985 · 2,0464656 | 1,2e-16; 0 | APROVADO |
| WBC | XN1000 | 625311 | N3 | nov/2026 | 1º ATUALIZAR | + MAN_0010 17,364 | (209,56 + 17,364)/13 | 13 · **17,4556923077** · 0,2229302524 | 13 · 17,4556923077 · 0,2229302524 · 1,2771207 | 0; 0 | APROVADO |
| WBC | XN1000 | 625311 | N2 | nov/2026 | 2º ATUALIZAR | + MAN_0017 7,344 (o P7 corrigido) | 191,028/26 | 26 · **7,3472307692** · 0,1473247590 | 26 · 7,3472307692 · 0,1473247590 · 2,0051740 | 2,4e-16; 0 | APROVADO |
| WBC | XN1000 | 625311 | N1 | ano 2026 (Estatística) | antes | 38 reais | idem | 38 · 3,1268421053 · 0,0623462547 | Estatística: 38 · 3,1268421053 · 0,0623462547 | 0; 1,1e-16 | APROVADO |
| WBC | XN1000 | 625311 | N2 | ano 2026 (Estatística) | antes | 62 reais | idem | 62 · 7,3322580645 · 0,1102964612 | 62 · 7,3322580645 · 0,1102964612 | 1,2e-16; 0 | APROVADO |
| WBC | XN1000 | 625311 | N3 | ano 2026 (Estatística) | antes | 33 reais | idem | 33 · 17,4124242424 · 0,1997972457 | 33 · 17,4124242424 · 0,1997972457 | 2,0e-16; 1,4e-16 | APROVADO |
| WBC | XN1000 | 625311 | N1 | ano 2026 (Estatística) | 1º ATUALIZAR | + os 3 do N1 acima | (118,82 + 9,556)/41 | 41 · **3,1311219512** · 0,0620774497 | 41 · 3,1311219512 · 0,0620774497 | 0; 0 | APROVADO |
| WBC | XN1000 | 625311 | N2 | ano 2026 (Estatística) | 1º ATUALIZAR | + 7,877; 7,357 | (454,6 + 15,234)/64 | 64 · **7,34115625** · 0,1281351245 | 64 · 7,34115625 · 0,1281351245 | 2,4e-16; 0 | APROVADO |
| WBC | XN1000 | 625311 | N3 | ano 2026 (Estatística) | 1º ATUALIZAR | + 17,364 | (574,61 + 17,364)/34 | 34 · **17,411** · 0,1969219196 | 34 · 17,411 · 0,1969219196 | 2,0e-16; 0 | APROVADO |
| WBC | XN1000 | 625311 | N2 | ano 2026 (Estatística) | 2º ATUALIZAR | + 7,344 | 477,178/65 | 65 · **7,3412** · 0,1271306169 | 65 · 7,3412 · 0,1271306169 | 2,4e-16; 2,2e-16 | APROVADO |

Os valores que **não** podem entrar não entraram:
- os CONFLITO por chave repetida: MAN_0002/3/4/6/11/12/13;
- os MAN_INVALIDO_L13/L14 (3,184 e 3,176);
- o P7 (7,344) no 1º ATUALIZAR.

O n bate exatamente com a contagem dos ATIVOS. A tolerância é a padrão (1e-9 relativo), e o maior desvio, 2,4e-16, é
arredondamento de ponto flutuante. Na base, os números também batem com o oráculo, mas com os resultados colados
fora (N1 18, N2 24 e depois 25, N3 12): o defeito era de entrada, não de cálculo.

**Como reproduzir:**
1. Filtre a tblCQ_Final por PARTICIPA_ESTATISTICA = "SIM", ANALITO = "WBC", EQUIPAMENTO = "XN1000", LOTE = "625311",
   NIVEL e DATA_HORA em nov/2026 (Painel) ou em 2026 (Estatística).
2. Calcule `statistics.fmean` e `statistics.stdev`, e o CV = 100·DP/média.
3. Faça a conta à mão: n₁ = n₀ + k e média₁ = (Σ₀ + Σ novos)/n₁, com Σ₀ = n₀·média₀.

Os números estão em `Hematologia_F1_nova.json` (`antes`, `depois1`, `depois2`, chave `comparacao`).

**LJ e Westgard (Eng_Saida, a fonte do gráfico), depois do 2º ATUALIZAR:**

| Slot | Corrida | N1 (ID · valor · corrida do lote) | N2 | N3 | Marcas |
|---|---|---|---|---|---|
| 68 | 26111401 | MAN_0001 · 3,188 · 39 | MAN_0005 · 7,877 · 63 | — | N2 **1_3s** (REJEITADO) |
| 69 | 26111402 | MAN_0008 · 3,164 · 40 | — | — | — |
| 70 | 26111403 | — | — | MAN_0010 · 17,364 · 34 | — |
| 71 | 26111404 | — | MAN_0014 · 7,357 · 64 | — | — |
| 72 | 26111405 | — | MAN_0017 · 7,344 · 65 (só no 2º) | — | — |
| 73 | 26111406 | MAN_0007 · 3,204 · 41 | — | — | — |

O alvo do lote está em cada ponto: N1 3,18 ± 0,07935, N2 7,33 ± 0,1368 e N3 17,44 ± 0,25395. O z de cada ponto novo,
fora o MAN_0005, tem módulo ≤ 0,30. Na corrida 26111401, a amplitude entre níveis é 4,00 − 0,10 = 3,90 < 4, então não
há R-4s, e só um nível passa de 2 DP, então não há 2of3-2s. A única marca esperada é o 1-3s do MAN_0005, e é a única que
aparece. A corrida do lote continua a sequência real: N1 38 → 39, N2 62 → 63 e N3 33 → 34 (ADR-080).

## Ocorrências (fluxos vizinhos)

As quatro primeiras são as da Bio neste ciclo, confirmadas na Hema com o mesmo roteiro. Para não duplicar no
registro, sugiro tratá-las como "Ambos".

### R6-01 (Hema): apagar ou mover uma linha manual deixa USUARIO e REGISTRADO_EM, e o próximo manual digitado nela é recusado

- **Tela/módulo:** aba Digitar Resultados (`mIntegracao.PrepararLinha`/`Carimbar`, `EntradaMudou`) e motor (regra
  "DATA/HORA posterior ao lancamento").
- **Ação:**
  1. Digitei o manual A (WBC N1, 14/11 07:40). O carimbo saiu às 21:47:02.
  2. Apaguei ID…MOTIVO (A:K). As colunas L:P são travadas.
  3. Esperei 2,5 min e cliquei no botão Digitar Resultados. O cursor foi para B7, a linha apagada.
  4. Digitei ali o manual B (15/11 21:48).
  5. Fiz o controle: o manual C, numa linha nunca usada.
  6. Rodei o ATUALIZAR.
- **Esperado:** B tem carimbo próprio e fica ATIVO, como o C.
- **Observado:**
  - B herda USUARIO e REGISTRADO_EM 21:47:02.
  - Depois do ATUALIZAR, B fica **MANUAL_INCOMPLETO** "campos obrigatórios ausentes/inválidos: DATA/HORA posterior ao
    lançamento", com E04. O C fica ATIVO.
  - Na base (entrega 42), acontece o mesmo.
  - O resíduo também fica depois do Delete da linha colada (T14) e na origem de um recorte (T06).
- **Dificuldade:** as colunas do carimbo são travadas, e o DETALHE manda olhar a data/hora, que está certa. O usuário
  não tem como corrigir.
- **Consequência:**
  - Um manual válido fica fora da estatística, do LJ e do Westgard.
  - A trilha atribui o lançamento a outra pessoa, em outro momento.
  - O ADR-096 manda "apague … a linha", o que torna esse caminho mais comum.
- **Sugestão:** a mesma da Bio. Limpar USUARIO e REGISTRADO_EM quando ID…MOTIVO e SUBSTITUI_ID ficam vazios, como a
  aba Inativar já faz (ADR-070), e fazer o `PrepararEntradas` limpar o carimbo das linhas sem dado.
- **Evidência:** `Hematologia_F1_carimbo.json` e `Hematologia_F1_carimbo_base.json` (A_apagado, botao_parou_em,
  B_digitado, atualizar).
- **Prioridade:** Alto. **Tipo:** funcional.

### R6-02 (Hema): copiar a linha inteira, como o GUIA §3 sugere, é recusado pela proteção

- **Tela/módulo:** aba Digitar Resultados (proteção) e GUIA §3.
- **Ação:** copiei a linha inteira da tabela (A:Q) e colei na linha livre, com a proteção como o teclado do usuário a
  enxerga.
- **Esperado:** a colagem funciona e a linha ganha ID novo, pelo GUIA ("Pode copiar uma linha e colar embaixo…").
- **Observado:**
  - O Excel recusa com "A célula ou gráfico que você está tentando alterar está em uma planilha protegida… Pode ser
    necessário inserir uma senha." Nada é colado.
  - Só funciona copiar ID…MOTIVO (A:K).
  - É igual na base.
  - Pela COM, a linha inteira é aceita e traz o carimbo da origem (T15).
- **Consequência:** o usuário desiste do atalho ou acha que precisa de senha. Não há risco para os dados.
- **Sugestão:** o GUIA §3 e o ADR-096 devem dizer "selecione de ID até MOTIVO (as colunas brancas), copie e cole na
  próxima linha".
- **Evidência:** `Hematologia_F1_nova.json`, passo `P1c_protecao_do_usuario` (`linha_inteira` RECUSADA, `bloco_AK`
  ACEITO), e o mesmo passo em `Hematologia_F1_base.json`. É uma emulação: o Ctrl+V físico não foi executado.
- **Prioridade:** Médio. **Tipo:** usabilidade.

### R6-03 (Hema): o manual com ID inválido mostra "atualize os dados" no DETALHE depois do ATUALIZAR

- **Tela/módulo:** aba Digitar Resultados (o STATUS, o RUN e o DETALHE vêm por XLOOKUP do ID) e a mensagem do
  ATUALIZAR.
- **Ação:** duas linhas completas com ID "TESTEAUD9"; depois, ATUALIZAR, duas vezes.
- **Esperado:** a tela diz por que as linhas não entraram.
- **Observado:**
  - O motor recusa corretamente (MAN_INVALIDO_L13/L14, MANUAL_INCOMPLETO, E04).
  - A mensagem diz "2 resultado(s) manual(is) incompleto(s): **veja o DETALHE** na aba Digitar Resultados".
  - Na tela, porém, o STATUS e o RUN mostram "—" e o DETALHE "**atualize os dados**", mesmo depois de 2 ATUALIZAR.
  - É igual na base.
- **Consequência:** o usuário atualiza de novo, em vão, e o manual fica fora da estatística.
- **Sugestão:** o DETALHE pode cair no `MAN_INVALIDO_L<linha>`, ou mostrar "ID fora do padrão MAN_0001: apague o ID
  (um novo é gerado)".
- **Evidência:** `Hematologia_F1_nova.json`, `atualizar1.tela` e `atualizar2.tela` (P6a/P6b), `tblCQ_Final` e
  `qa_integracao`.
- **Prioridade:** Médio. **Tipo:** usabilidade.

### R6-04 (Hema): colar sem mudar hora nem nível; o aviso sugere que trocar o resultado basta, e o A01 manda "inative um deles"

- **Tela/módulo:** o aviso do `EntradaMudou` (ADR-096) e o texto do A01 para "resultado manual repetido (mesma chave e
  instante)".
- **Ação:** colei a origem e troquei **só** o RESULTADO (3,148).
- **Observado:**
  - A linha fica CONFLITO_MANUAL "resultado manual repetido (mesma chave e instante)", fora da estatística. É a regra
    do motor, e a tela mostra isso.
  - Os A01 dessa linha (MAN_0006) e das coladas sem edição (MAN_0002/3/4/11/12/13) dizem "-> MAN_0001 … confirme e
    **inative um deles** ou corrija o manual".
  - A mensagem do ATUALIZAR não conta esses CONFLITO.
- **Consequência:** quem repete o controle no mesmo minuto e troca só o valor perde o resultado. E "inative um deles"
  pode levar a inativar o MAN_0001, o resultado que vale. É o mesmo tropeço que o ADR-096 corrigiu no ramo do ID.
- **Sugestão:**
  - No aviso: "Mude a hora ou o nível; só trocar o resultado deixa a linha como repetida."
  - No A01 desse ramo: "apague a linha repetida <ID>; não inative o <ID que vale>".
- **Evidência:** `Hematologia_F1_nova.json`, passo `P2c_so_resultado` e `atualizar1.qa_integracao`.
- **Prioridade:** Baixo. **Tipo:** usabilidade.

### R6H-05 (nova, só Hema): a aba Estatística vem com Ano De/Até = 2025 e mostra n 0 para o lote em uso

- **Tela/módulo:** aba Estatística da Hematologia (B3/D3 Ano De/Até, B4 lote em análise).
- **Ação:**
  1. Lancei e atualizei os manuais.
  2. Abri a aba Estatística sem mexer no ano.
  3. Depois, pus 2026 em B3 e D3.
- **Esperado:** ver os números do lote em uso (625311, que é de 2026), incluindo o manual que acabei de lançar, ou um
  aviso de que o ano não tem dado do lote.
- **Observado:**
  - B3 = D3 = **2025** no arquivo de produção.
  - Com o B4 = 625311, WBC N1 a N3 = **n 0, média e DP vazios**, antes e depois dos manuais.
  - Com B3/D3 = 2026, os números aparecem e batem com o oráculo.
  - O K5 também diz "Bias do EP: CAP | ano 2025", que é o ano do CEQ, um filtro separado (N4).
- **Dificuldade:** o Painel está em nov/2026 e a Estatística em 2025. Nada na aba diz que o lote não tem dado nesse ano.
- **Consequência:** o usuário que confere o lançamento pela aba Estatística acha que o resultado não entrou, ou que o
  lote não tem estatística. Não há erro de cálculo.
- **Sugestão:**
  - Iniciar B3/D3 no ano do lote em análise, ou no ano corrente.
  - Ou, com n 0 para o lote, mostrar "o lote <B4> não tem resultado em <ano>; ajuste o Ano De/Até".
  - Pode ser tratada junto com a OC-0048 (Ano De/Até corta lotes que atravessam o ano), que já está aberta.
- **Evidência:** `Hematologia_F1_nova_exec1.json`: `estatistica_cab` B3/D3 = 2025 e `comparacao.estatistica` = [0,
  "", ""] nos 3 níveis, antes e depois. `Hematologia_F1_nova.json`: `estatistica_cab.antes_de_digitar_o_ano` = 2025.
- **Prioridade:** Médio. **Tipo:** usabilidade.

## Dúvidas de regra de negócio

1. **R6-01:** o REGISTRADO_EM deve ser o momento em que **esta** linha recebeu dados, ou o primeiro carimbo da célula?
   É a mesma dúvida da Bio.
2. **R6-04:** dois manuais com a mesma chave e o mesmo instante, mas com valores diferentes, devem continuar como
   "repetido"? Na Hema isso pesa mais, porque o mesmo controle às vezes é repassado em seguida.
3. **R6H-05:** qual é o ano esperado na abertura da aba Estatística da Hema: o do lote em análise, o corrente ou o
   último usado?

## Impressões de uso (Relatório A)

- **É fácil inserir?** É. Com a correção, colar ficou seguro na Hema também:
  - o aviso aparece na hora e diz o ID antigo, o novo e o que fazer;
  - no bloco, é uma caixa só;
  - a MATRIZ se preenche sozinha (SANGUE TOTAL), e o EQUIPAMENTO, sendo único, não confunde.
- **A tela bate com o motor?**
  - Para tudo o que entra com os eventos ligados, sim: cada linha tem o próprio ID e o próprio STATUS, RUN e DETALHE.
  - O resíduo é o repetido que entra com os eventos desligados. Ele aparece ATIVO na tela até a correção, mas a
    mensagem do ATUALIZAR o aponta, e apagar o ID resolveu em 1 passo.
- **Mensagens:**
  - O aviso da colagem e a linha nova do ATUALIZAR são claros, e o A01 do ID repetido não manda mais inativar o
    resultado bom.
  - Ficam confusos o "atualize os dados" do ID inválido (R6-03) e o "inative um deles" da chave repetida (R6-04).
- **Riscos de preencher errado:**
  - Apagar uma linha e reaproveitá-la, que é para onde o botão leva, recusa o manual novo com um motivo que o usuário
    não consegue corrigir (R6-01).
  - Copiar a linha inteira esbarra na proteção, com uma mensagem que fala em senha (R6-02).
- **Dá para ver o que foi salvo?** No Painel, sim: o n e a média mudaram na hora, o ponto colado e editado apareceu no
  LJ e disparou o 1-3s (REJEITADO), o que na entrega 42 não acontecia. Na aba Estatística, só depois de trocar o ano
  para 2026 (R6H-05).
- **Números:** n, média, DP e CV do Painel e da aba Estatística batem com a minha conta em todos os momentos e nos 3
  níveis.
