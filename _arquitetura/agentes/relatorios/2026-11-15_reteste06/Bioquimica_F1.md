# Reteste 06: Usuário Simulado, Bioquímica, frente F1

Data: 15/11/2026 · Agente: AGENTE_USUARIO_SIMULADO · Modo: RETESTE da OC-0073 · Produção: entrega 43 (ADR-096).

## Ambiente e método

| Item | Valor |
|---|---|
| Produção atual (só copiada) | `QC_Bioquimica.xlsm` sha256 curto **bd5314426ab8**, igual antes e depois das 2 execuções (`origem_intacta.igual = true`) |
| Base de comparação ("antes") | `_backup_pre_integracao_2026-11-15_1802/QC_Bioquimica.xlsm` sha256 curto **b20fe00f7e51** (entrega 42), igual antes e depois |
| Cópias (novas, nenhuma salva, todas apagadas no fim) | `%TEMP%\qcw\auditoria\2026-11-15_reteste06\Bioquimica_F1_nova.xlsm`, `Bioquimica_F1_base.xlsm`, `Bioquimica_F1_carimbo.xlsm` e `Bioquimica_F1_carimbo_base.xlsm` |
| Excel | Um por vez, abertos e encerrados pelo `qa_final.QA`: PID 27380 (nova), 5392 (carimbo, nova), 31420 (base) e 28200 (carimbo, base). Os Excel de outros processos (13084 e os da tarefa agendada, que trocam de PID a cada 2 min) não foram tocados. Nenhuma captura de tela |
| Scripts (novos, documentados no cabeçalho) | `_arquitetura/agentes/ferramentas/reteste06_usuario_f1.py <nova\|base>` e `_arquitetura/agentes/ferramentas/reteste06_carimbo_residual.py [nova\|base]` |
| Saídas (fora do repositório) | `Bioquimica_F1_{nova,base,carimbo,carimbo_base}.json` e os `.log`, na pasta do ciclo no `%TEMP%` |
| Usuários | `TESTE_AUDITORIA_R6` digitou a origem. `TESTE_AUDITORIA_R6B` fez as colagens. `TESTE_AUDITORIA_R6C` fez o teste do carimbo. Todos ANALISTA, criados só na cópia; login por `mSeguranca.DoLogin` |
| Dados de teste | Glicose, DIMENSION 1, lote 8977, 14/11/2026 (o último dado real é de 13/11 08:39), MOTIVO "Outro". Valores sintéticos de 3 decimais calculados do alvo do lote (`mLotes.MediaDoLote`/`DPDoLote`: N1 79,314 ± 1,543; N2 274,776 ± 5,027). N1: 79,468 / 79,005 / 79,777 / 78,697 / 79,391 / 79,237. N2: **294,884 (média + 4,00 DP, para provocar 1-3s)** / 273,268 / 275,781 / 275,279. ID inválido `TESTEAUD9`. Carimbo: 79,321 / 79,654 / 274,321 em 14 e 15/11 |
| tblCQ_Final | 35.906 linhas no início. Os ATUALIZAR leram o DB_SEAC só para leitura e não trouxeram resultado novo |

**Caminho do usuário:**
- Fiz o login. No Painel, escrevi L4, C3, H3, K3 e K4 com os eventos ligados.
- Entrei pelo botão (`mIntegracao.IrDigitarResultados`) e digitei célula a célula (`FormulaLocal`).
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
- Painel: engPainel e A7:T8 (contagem do Westgard e status);
- Eng_Saida: pontos do LJ por ID, alvo e marcas do Westgard;
- aba Estatística: Glicose N1 e N2 (B5:J5 = 01/11/2026 a 30/11/2026).

**Oráculo próprio:** Python puro (`statistics`) sobre o `Value2` da tblCQ_Final. Ele pega PARTICIPA = SIM, Glicose,
DIMENSION 1, lote 8977 e DATA_HORA em nov/2026, e calcula n, média, DP amostral e CV. Também confere à mão a soma
"antes + valores novos".

## Resultado

| OC | Estado | Resumo |
|---|---|---|
| OC-0073 | **RESOLVIDA** | Toda linha colada com o ID de outra ganha um ID novo na hora, com um aviso claro. Isso vale para A:K, linha inteira, bloco de 3 e ID digitado, inclusive "man_1". Colar e editar fica ATIVO e entra na estatística, no LJ e no Westgard (1-3s acusado). A origem não muda. Mover mantém o ID. O ID inválido repetido continua recusado como na base. O repetido que entra com os eventos desligados é explicado pela mensagem do ATUALIZAR e pelo A01 novo, e apagar o ID resolve. Na base (entrega 42) o mesmo roteiro reproduz o defeito |

Os fluxos vizinhos acharam **4 ocorrências novas**, nenhuma causada pela entrega 43; as 4 já existem na base. A de
maior peso é a **R6-01 (Alto)**: apagar uma linha manual deixa o carimbo e recusa o próximo manual digitado nela.

## Testes executados

| ID | Teste | Estado | Evidência (nova = entrega 43; base = entrega 42) |
|---|---|---|---|
| OC-0073 | Reteste do fluxo original e dos critérios de aceite | **RESOLVIDA** | Itens T01 a T09 abaixo |
| T01 | Colar sem editar (A:K da origem na linha livre) | APROVADO | O MAN_0001 colado vira **MAN_0002** na hora. Aviso: "A linha trazia o ID MAN_0001, que já existe em outra linha. Ela recebeu o ID novo MAN_0002 (o ID é único e gerado pelo sistema). Mude o que for diferente (nível, hora, resultado) antes de clicar em ATUALIZAR DADOS." Depois do ATUALIZAR: CONFLITO_MANUAL, DETALHE "resultado manual repetido (mesma chave e instante)", RUN vazio. A origem MAN_0001 continua ATIVO, RUN 26111401. Nenhum `~L` na tblCQ_Final. Base: a colada fica com MAN_0001, a tela mostra ATIVO e a tblCQ_Final tem `MAN_0001~L2` CONFLITO "vale a 1a linha" |
| T02 | Colar e editar (NIVEL 2, RESULTADO 294,884 = média + 4 DP) | APROVADO | ID **MAN_0005** logo depois de colar, e ele não muda com a edição. ATIVO, PARTICIPA SIM, RUN 26111401 (a mesma corrida da origem, pela mesma hora). **LJ:** Eng_Saida, slot 126, ID N2 = MAN_0005, valor 294,884. **Westgard:** marca 1_3s = 1 no N2 (z = (294,884 − 274,7759)/5,0270 = 4,00003). Painel M8 vai de 0 a 1, J8 de OK a **REJEITADO**, S8 "1_3s · RUN 26111401". Audit_Log WESTGARD_EVENTO_CORRIDA_NOVA "1_3s INDIVIDUAL… N2 z=4,00". **Base:** o mesmo valor fica `MAN_0001~L4` CONFLITO, o N2 fica com n 11 sem ele, M8 = 0 e o status OK |
| T03 | A linha original não muda | APROVADO | A linha da origem lida depois de cada colagem é idêntica à de antes (`origem_igual = true` em todos os passos, `origem_intacta_antes_atualizar = true`). No fim: MAN_0001, 79,468, usuário 1, REGISTRADO_EM 21:13:32, ATIVO |
| T04 | Digitar um ID MAN_ que já existe | APROVADO | Em linha vazia, ID digitado antes dos dados: "MAN_0001" vira **MAN_0008**, com aviso. Sobre uma linha preenchida (o ID automático MAN_0009), digitei "man_1" (minúsculo, sem zeros) e ele virou **MAN_0010**, com aviso. Os dois ficam ATIVOS depois do ATUALIZAR (RUN 26111402 e 26111403). Base: os dois ficam MAN_0001, `~L8` e `~L9` CONFLITO |
| T05 | Bloco de 3 linhas colado de uma vez | APROVADO | MAN_0001, 0002 e 0003 viram **MAN_0011, MAN_0012 e MAN_0013**, em sequência, com **uma** caixa: "As 3 linhas coladas traziam IDs que já existem em outras linhas e receberam IDs novos (MAN_0011 a MAN_0013)." A marca qcUltimoIdManual vai de 10 a 13 |
| T06 | Mover (recortar A:K e colar 3 linhas abaixo) | APROVADO | O ID MAN_0014 se mantém, sem aviso. A origem A:K fica vazia. Depois do ATUALIZAR: ATIVO, RUN 26111404. Observação: USUARIO e REGISTRADO_EM ficam na linha de origem (ver R6-01) |
| T07 | ID fora do padrão MAN_ repetido ("TESTEAUD9" em 2 linhas) | APROVADO | Não é trocado. Na tblCQ_Final: `MAN_INVALIDO_L13` e `MAN_INVALIDO_L14`, MANUAL_INCOMPLETO, "ID_REGISTRO (padrao MAN_0001)". A QA tem 2 **E04**. É **idêntico na base**. A tela, porém, não mostra o motivo (ver R6-03) |
| T08 | Repetido criado com os eventos desligados (dado antigo) e a correção | APROVADO | 1º ATUALIZAR: `MAN_0001~L15` CONFLITO "vale a 1a linha". A mensagem traz "• 1 linha(s) da Digitar Resultados repete(m) o ID de outra (ex.: MAN_0001) e ficou(aram) fora da estatística. Apague o ID da linha repetida (um novo é gerado) e atualize de novo." O A01 traz "…apague o ID da linha repetida (um novo sera gerado) ou a linha. Nao inative o MAN_0001: isso tiraria o resultado que vale." Ao apagar o ID, a linha recebeu **MAN_0017**. 2º ATUALIZAR: ATIVO, RUN 26111405, sem A01 e sem a linha na mensagem, e o N2 vai de 13 para 14. Antes da correção, a tela dessa linha ainda mostrava ATIVO/26111401 (o XLOOKUP acha a 1ª linha). É o resíduo previsto no ADR-096, coberto pela mensagem |
| T09 | Estatística, Painel e LJ conferidos com o oráculo | APROVADO | Relatório B abaixo. n e média são exatos, o DP vem com diferença relativa ≤ 5e-13 e os novos pontos estão no Eng_Saida |
| T10 | Fluxo vizinho: colar e trocar só o RESULTADO (mesma hora e nível) | APROVADO (com R6-04) | Recebe ID novo (MAN_0006). Fica CONFLITO_MANUAL "resultado manual repetido (mesma chave e instante)", fora da estatística: é a regra do motor (chave sem o valor). O aviso sugere que trocar o resultado basta (R6-04) |
| T11 | Fluxo vizinho: colar e trocar HORA e RESULTADO | APROVADO | MAN_0007, ATIVO, corrida própria (26111405; 26111406 depois que o MAN_0017 das 15:00 entrou, porque as corridas seguintes do dia andam +1, como diz o GUIA) |
| T12 | Fluxo vizinho: colar a **linha inteira** com a proteção que o usuário vê | REPROVADO | Aba protegida, com USUARIO, REGISTRADO_EM, RUN, STATUS e DETALHE travados. Com a proteção sem UserInterfaceOnly, que é a regra que o teclado do usuário obedece, colar a linha inteira dá: "A célula ou gráfico que você está tentando alterar está em uma planilha protegida… Pode ser necessário inserir uma senha." Nada é colado. O bloco A:K é aceito. O mesmo acontece na base. Ver R6-02 |
| T13 | Fluxo vizinho: apagar uma linha manual (A:K) e digitar outro manual nela | REPROVADO | USUARIO e REGISTRADO_EM ficam na linha. O botão leva o cursor exatamente para ela (B7). O manual novo herda o carimbo antigo e é recusado: MANUAL_INCOMPLETO "DATA/HORA posterior ao lançamento". O controle igual, numa linha nunca usada, fica ATIVO. O mesmo acontece na base. Ver R6-01 |
| T14 | Fluxo vizinho: apagar a linha colada (A:K) e atualizar | APROVADO | O MAN_0002 some da tblCQ_Final, e o Audit_Log grava RESULTADO_MANUAL_RETIRADO. Nada mais muda. Restam USUARIO e REGISTRADO_EM na linha (R6-01) |
| T15 | Trilha: quem lançou a linha colada | APROVADO (A:K) / observação (linha inteira) | Na colagem A:K, USUARIO = usuário 2 e REGISTRADO_EM = o momento da colagem. Na colagem da linha inteira pela COM (o caminho do Tester), o USUARIO e o REGISTRADO_EM da **origem** vêm junto: o Audit_Log diz "manual lancado por TESTE_AUDITORIA_R6 em 21:13:32", e quem colou foi o R6B. Pelo teclado isso não acontece, porque a colagem é recusada (T12) |
| T16 | Base (entrega 42), mesmo roteiro | APROVADO (prova de que o teste pega o defeito) | Nenhum aviso e nenhum ID novo. 11 linhas `MAN_0001~L<n>` CONFLITO "vale a 1a linha" aparecem **ATIVO** na tela. N1 e N2 ficam com n 11 (nova: 13). Sem 1-3s. A mensagem não cita o caso, e o A01 manda "inative um deles" |

## Relatório B: conferência numérica (Painel e Estatística × oráculo)

Filtro do Painel: L4 DIMENSION 1 · C3 Glicose · H3 lote 8977 · K3 2026 · K4 Nov. Aba Estatística: janela
01/11/2026–30/11/2026. Fonte dos dados: `Value2` da tblCQ_Final da cópia `Bioquimica_F1_nova.xlsm` (produção
bd5314426ab8).

| Analito | Equip. | Lote | Nível | Período | Momento | Dados de entrada (novos ATIVOS) | Fórmula | Independente (n · média · DP) | QC_INI engPainel (n · média · DP · CV%) | QC_INI Estatística (n · média · DP) | Diferença relativa | Status |
|---|---|---|---|---|---|---|---|---|---|---|---|---|
| Glicose | DIMENSION 1 | 8977 | N1 | nov/2026 | antes | 10 resultados reais, soma 775,0 | média = Σx/n; DP amostral | 10 · 77,5 · 0,8498365856 | 10 · 77,5 · 0,8498365856 · 1,0965633 | 10 · 77,5 · 0,8498365856 | 0 / 0 | APROVADO |
| Glicose | DIMENSION 1 | 8977 | N2 | nov/2026 | antes | 10 reais, soma 2718,0 | idem | 10 · 271,8 · 5,9777364724 | 10 · 271,8 · 5,9777364724 · 2,1993144 | 10 · 271,8 · 5,9777364724 | DP Estat. 9,6e-14 | APROVADO |
| Glicose | DIMENSION 1 | 8977 | N1 | nov/2026 | 1º ATUALIZAR | + MAN_0001 79,468; MAN_0008 79,005; MAN_0007 79,777 | (775 + 238,25)/13 | 13 · **77,9423076923** · 1,1284025866 | 13 · 77,9423076923 · 1,1284025866 · 1,4477408 | 13 · 77,9423076923 · 1,1284025866 | DP Estat. 5,0e-13 | APROVADO |
| Glicose | DIMENSION 1 | 8977 | N2 | nov/2026 | 1º ATUALIZAR | + MAN_0005 **294,884**; MAN_0010 273,268; MAN_0014 275,781 | (2718 + 843,933)/13 | 13 · **273,9948461538** · 8,2150720006 | 13 · 273,9948461538 · 8,2150720006 · 2,9982579 | 13 · 273,9948461538 · 8,2150720006 | DP Estat. 6,2e-14 | APROVADO |
| Glicose | DIMENSION 1 | 8977 | N2 | nov/2026 | 2º ATUALIZAR | + MAN_0017 275,279 (o P7 corrigido) | 3837,212/14 | 14 · **274,0865714286** · 7,9002441901 | 14 · 274,0865714286 · 7,9002441901 · 2,8823901 | 14 · 274,0865714286 · 7,9002441901 | DP Estat. 1,2e-13 | APROVADO |

Os valores que **não** podem entrar não entraram:
- os CONFLITO por chave repetida (MAN_0002/3/4/6/11/12/13);
- os MAN_INVALIDO_L13/L14 (79,391 e 79,237);
- o P7 no 1º ATUALIZAR (275,279).

O n bate exatamente com a contagem dos ATIVOS. A tolerância é a padrão (1e-9 relativo), e o maior desvio, 5e-13 no DP da
aba Estatística, é arredondamento de ponto flutuante.

**Como reproduzir:**
1. Filtre a tblCQ_Final por PARTICIPA_ESTATISTICA = "SIM", ANALITO = "Glicose", EQUIPAMENTO = "DIMENSION 1",
   LOTE = "8977", NIVEL e DATA_HORA em novembro de 2026.
2. Calcule `statistics.fmean` e `statistics.stdev`, e o CV = 100·DP/média.
3. Faça a conta à mão: n₁ = n₀ + k e média₁ = (Σ₀ + Σ novos)/n₁, com Σ₀ = n₀·média₀.

Os números estão em `Bioquimica_F1_nova.json` (`antes`, `depois1`, `depois2`).

**LJ e Westgard (Eng_Saida, a fonte do gráfico):**

| Slot | Corrida | N1 (ID · valor) | N2 (ID · valor) | Marcas N1 / N2 |
|---|---|---|---|---|
| 126 | 26111401 | MAN_0001 · 79,468 | MAN_0005 · 294,884 | — / **1_3s** (REJEITADO) |
| 127 | 26111402 | MAN_0008 · 79,005 | — | — |
| 128 | 26111403 | — | MAN_0010 · 273,268 | — |
| 129 | 26111404 | — | MAN_0014 · 275,781 | — |
| 130 | 26111405 | — | MAN_0017 · 275,279 (só no 2º) | — |
| 131 | 26111406 | MAN_0007 · 79,777 | — | — |

Alvo do lote em cada linha: N1 79,3140496 ± 1,5427784 e N2 274,7758621 ± 5,0269933. O z de cada ponto novo, fora o
MAN_0005, fica abaixo de 0,5, então a única marca esperada é o 1-3s do MAN_0005, e é a única que aparece.

## Ocorrências novas (fluxos vizinhos)

### R6-01: apagar ou mover uma linha manual deixa USUARIO e REGISTRADO_EM, e o próximo manual digitado nela é recusado

- **Tela/módulo:** aba Digitar Resultados (`mIntegracao.PrepararLinha`/`Carimbar`, `EntradaMudou`) e motor
  (`mEtlCQ`, teste "DATA/HORA posterior ao lancamento").
- **Ação:**
  1. Digitei o manual A (14/11 07:40). O carimbo foi às 21:18:30.
  2. Selecionei ID…MOTIVO (A:K), o que o usuário consegue selecionar, porque L:P são travadas, e apaguei.
  3. Esperei 2,5 min e cliquei no botão Digitar Resultados. O cursor foi para **B7, a linha apagada**.
  4. Digitei ali o manual B, de 15/11 21:19, às 21:21:06.
  5. Fiz o controle: o manual C igual, numa linha nunca usada.
  6. Rodei o ATUALIZAR DADOS.
- **Esperado:** B tem carimbo próprio (o usuário de agora, 21:21) e fica ATIVO, como o C.
- **Observado:**
  - Depois do Delete, a linha fica com USUARIO = TESTE_AUDITORIA_R6C e REGISTRADO_EM = 15/11/2026 21:18:30.
  - O manual B **herda** esse carimbo, porque o `Carimbar` só preenche o vazio.
  - Depois do ATUALIZAR, B fica **MANUAL_INCOMPLETO**, com DETALHE "campos obrigatórios ausentes/inválidos: DATA/HORA
    posterior ao lançamento", fora da estatística. A mensagem conta "1 resultado(s) manual(is) incompleto(s)".
  - O controle C (REGISTRADO_EM 21:21:11) fica **ATIVO**.
  - O mesmo acontece na base (entrega 42): é anterior à entrega 43.
  - Recortar A:K para mover a linha deixa o mesmo resíduo na linha de origem (T06). A próxima digitação nela, o P6a,
    herdou o REGISTRADO_EM do P5.
- **Dificuldade:** o usuário não tem como corrigir. As colunas USUARIO e REGISTRADO_EM são travadas, e o DETALHE manda
  olhar a data/hora, que está certa. A única saída é usar outra linha, e nada diz isso.
- **Consequência:**
  - No uso real (apagar hoje e digitar dias depois na mesma linha, que é para onde o botão leva), **todo** manual
    digitado nessa linha com data/hora depois do apagamento é recusado. O resultado de controle fica fora da
    estatística, do LJ e do Westgard.
  - A trilha atribui o lançamento a quem digitou a linha apagada, na data antiga. O Audit_Log diz "manual lancado por
    <usuário antigo> em <data antiga>". Isso vale também quando outro usuário digita.
  - O ADR-096, o A01 e a mensagem do ATUALIZAR agora mandam "apague … a linha", o que torna esse caminho mais comum.
- **Sugestão (para decisão do Orquestrador):**
  - No `EntradaMudou` da aba manual, quando as colunas de dado (ID…MOTIVO e SUBSTITUI_ID) ficarem todas vazias, limpar
    USUARIO e REGISTRADO_EM, como a aba Inativar já faz no ADR-070.
  - Fazer o `PrepararEntradas` limpar o carimbo das linhas sem dado. Isso cobre o que já existe.
  - Opcional: carimbar de novo quando a linha passa de vazia para preenchida.
- **Evidência:** `Bioquimica_F1_carimbo.json` (A_apagado, botao_parou_em, B_digitado, atualizar) e
  `Bioquimica_F1_carimbo_base.json`. Script: `ferramentas/reteste06_carimbo_residual.py`.
- **Prioridade:** **Alto**. É risco operacional relevante: um manual válido fica fora do CQ sem que o usuário consiga
  corrigir, e a trilha fica errada. **Tipo:** funcional.

### R6-02: "copiar uma linha e colar embaixo" (GUIA §3) é recusado pela proteção quando a linha inteira é copiada

- **Tela/módulo:** aba Digitar Resultados (proteção da aba, colunas travadas L:P) e GUIA §3.
- **Ação:** copiei a linha inteira da tabela (A:Q) e colei na linha livre, com a proteção como o teclado do usuário a
  enxerga (sem UserInterfaceOnly).
- **Esperado:** pelo GUIA ("Pode copiar uma linha e colar embaixo…"), a colagem funciona e a linha ganha ID novo.
- **Observado:**
  - O Excel recusa e não cola nada: "A célula ou gráfico que você está tentando alterar está em uma planilha protegida.
    Para fazer uma alteração, desproteja a planilha. Pode ser necessário inserir uma senha."
  - Só funciona copiar o bloco **ID…MOTIVO (A:K)**.
  - Clicar no número da linha também inclui células travadas, e o resultado é o mesmo.
  - Igual na base.
  - Pela COM (o caminho do Tester), a linha inteira é aceita e traz USUARIO e REGISTRADO_EM da origem (T15).
- **Dificuldade:** o gesto natural, que é selecionar a linha toda, dá erro e fala em senha. O usuário não sabe que tem
  de selecionar só de ID até MOTIVO.
- **Consequência:** o usuário desiste do atalho ou acha que precisa de senha. Não há risco para os dados.
- **Sugestão:** o GUIA §3 e o ADR-096 devem dizer "selecione de ID até MOTIVO (as colunas brancas), copie e cole na
  próxima linha". Outra opção é uma dica na célula do ID.
- **Evidência:** `Bioquimica_F1_nova.json`, passo `P1c_protecao_do_usuario` (`linha_inteira` RECUSADA, `bloco_AK`
  ACEITO), e o mesmo passo em `Bioquimica_F1_base.json`. O teste é uma emulação: o Ctrl+V físico não foi executado.
- **Prioridade:** Médio. **Tipo:** usabilidade.

### R6-03: manual com ID inválido mostra "atualize os dados" no DETALHE depois do ATUALIZAR, e a mensagem manda ver o DETALHE

- **Tela/módulo:** aba Digitar Resultados (STATUS, RUN e DETALHE por XLOOKUP do ID) e a mensagem do ATUALIZAR.
- **Ação:** duas linhas com dados completos e ID "TESTEAUD9" (fora do padrão MAN_); depois, ATUALIZAR.
- **Esperado:** a tela dessas linhas diz por que não entraram.
- **Observado:**
  - O motor as recusa corretamente: MAN_INVALIDO_L13/L14, MANUAL_INCOMPLETO, E04.
  - A mensagem diz "2 resultado(s) manual(is) incompleto(s): **veja o DETALHE** na aba Digitar Resultados".
  - Na tela, porém, o STATUS e o RUN mostram "—" e o DETALHE mostra "**atualize os dados**", mesmo depois de
    atualizar. O XLOOKUP procura "TESTEAUD9", e a tblCQ_Final tem "MAN_INVALIDO_L13".
  - O motivo só aparece na QA_INTEGRACAO.
  - Igual na base.
- **Dificuldade e consequência:** o usuário atualiza de novo, em vão, e não descobre que o problema é o ID que
  digitou. O manual fica fora da estatística. O motor está certo; a tela é que não explica.
- **Sugestão:**
  - O DETALHE pode cair no ID `MAN_INVALIDO_L<linha>` quando o ID da linha não é MAN_ válido, ou mostrar "ID fora do
    padrão MAN_0001: apague o ID (um novo é gerado)".
  - Outra opção é a mensagem do ATUALIZAR citar o E04 de ID.
- **Evidência:** `Bioquimica_F1_nova.json`, `atualizar1.tela` (P6a/P6b), `tblCQ_Final` e `qa_integracao`.
- **Prioridade:** Médio. **Tipo:** usabilidade.

### R6-04: linha colada sem mudar hora nem nível: o aviso sugere que trocar o resultado basta, e o A01 manda "inative um deles"

- **Tela/módulo:** aviso do `EntradaMudou` (ADR-096) e texto do A01 para "resultado manual repetido (mesma chave e
  instante)".
- **Ação:** colei a origem e troquei **só** o RESULTADO (78,697).
- **Esperado:** pelo aviso ("Mude o que for diferente (nível, hora, resultado)"), mudar o resultado bastaria.
- **Observado:**
  - A linha fica CONFLITO_MANUAL "resultado manual repetido (mesma chave e instante)", fora da estatística. É a regra do
    motor: a chave é equipamento, matriz, lote, nível, analito e instante, sem o valor. A tela mostra isso
    corretamente.
  - O A01 desse caso, e também o de colar sem editar, diz "confirme e **inative um deles** ou corrija o manual" e
    aponta "-> MAN_0001", o que vale.
  - A mensagem do ATUALIZAR não conta esses CONFLITO.
- **Consequência:** é menor, porque o STATUS na tela está certo. Mas quem lança uma repetição do controle no mesmo
  minuto e troca só o valor perde o resultado. E "inative um deles" pode levar a inativar o MAN_0001, que é o que vale.
  É o mesmo tropeço que o ADR-096 corrigiu no ramo do ID. Não testei o efeito de inativar o MAN_0001.
- **Sugestão:**
  - No aviso: "Mude a hora ou o nível; só trocar o resultado deixa a linha como repetida."
  - No A01 desse ramo: "apague a linha repetida <ID>; não inative o <ID que vale>".
- **Evidência:** `Bioquimica_F1_nova.json`, passo `P2c_so_resultado` e `atualizar1.qa_integracao` (A01 do MAN_0006).
- **Prioridade:** Baixo. **Tipo:** usabilidade.

## Dúvidas de regra de negócio

1. **R6-01:** o REGISTRADO_EM deve ser o momento em que **esta** linha recebeu dados, ou o primeiro carimbo da célula?
   O motor usa esse campo como o "lançamento" na regra "DATA/HORA posterior ao lançamento", e o Audit_Log o usa como
   "manual lançado por … em …". Minha leitura é que tem de ser o da linha atual.
2. **R6-04:** dois manuais com a mesma chave e o mesmo instante e valores diferentes devem continuar como "repetido"?
   Hoje vale o primeiro, e é o que a regra do motor diz. Se for intencional, o aviso do ADR-096 precisa dizer que o
   resultado sozinho não basta.

## Impressões de uso (Relatório A)

- **É fácil inserir?** É. Com a correção, colar ficou seguro. O aviso aparece na hora, diz o ID antigo e o novo e o que
  fazer. No bloco, é uma caixa só para as 3 linhas, sem fila de caixas.
- **A tela bate com o motor?**
  - Para tudo o que entra com os eventos ligados, sim. Cada linha tem o próprio ID, e o STATUS, o RUN e o DETALHE são
    os dela.
  - O resíduo é o dado antigo, ou entrado com os eventos desligados. Ele aparece ATIVO na tela até a correção, mas a
    mensagem do ATUALIZAR agora o aponta e diz exatamente o que fazer. Apagar o ID resolveu em 1 passo.
- **Mensagens:**
  - O aviso da colagem e a linha nova da mensagem do ATUALIZAR são claros. O A01 novo não manda mais inativar o
    resultado bom.
  - Ficaram dois textos confusos fora do escopo da OC: o DETALHE "atualize os dados" no ID inválido (R6-03) e o "inative
    um deles" no repetido por chave (R6-04).
- **Riscos de preencher errado:**
  - O maior é o R6-01: apagar uma linha e reaproveitá-la, que é exatamente para onde o botão leva, recusa o manual
    novo com um motivo que o usuário não consegue corrigir.
  - O segundo é o R6-02: o gesto natural de copiar a linha inteira esbarra na proteção, com uma mensagem que fala em
    senha.
- **Números:** n, média, DP e CV do Painel e da aba Estatística batem com a minha conta em todos os momentos. O ponto
  colado e editado aparece no LJ e dispara o 1-3s, coisa que na entrega 42 não acontecia.
