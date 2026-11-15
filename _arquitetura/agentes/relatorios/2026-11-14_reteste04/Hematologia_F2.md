# Reteste 04 — Usuário Simulado — Hematologia, frente F2

Data: 14/11/2026 · Agente: AGENTE_USUARIO_SIMULADO · Modo: RETESTE (OC-0049, OC-0057, OC-0058, OC-0066, ADR-091,
ADR-094) + pedido do ciclo: incluir e corrigir um CEQ TESTE-AUDITORIA e conferir bias, ET e Sigma do analito certo, e
só dele. Produção: entrega 41.

## Ambiente e método

| Item | Valor |
|---|---|
| Produção copiada (nunca aberta) | `QC_Hematologia.xlsm` sha256 **aab0a54f464e0232** (= entrega 41), igual antes e depois das duas execuções (`producao_intacta.igual = true`) |
| Cópias novas (nenhuma salva) | `%TEMP%\qcw\auditoria\2026-11-14_reteste04\Hematologia_F2.xlsm` (parte 1: OC-0049, CEQ, OC-0057) e `Hematologia_F2_p2.xlsm` (parte 2: ADR-091, manuais, ATUALIZAR, ADR-094, Painel). Uma cópia de inspeção (`Hematologia_F2_insp.xlsm`), lida só por openpyxl |
| Retratos para a calculadora | `SaveCopyAs` com eventos desligados: `Hematologia_F2_{E0,S0,S1,S2,S3,RX,S4,R1}.xlsm`. Cópias e retratos apagados no fim (REGRAS §2.3); ficam os JSON |
| Excel | um por vez, aberto e fechado pelo `qa_final.QA`: PID 22512 (parte 1) e 10616 (parte 2). Antes de cada execução conferi que não havia Excel nem python rodando |
| Script | `_arquitetura/agentes/ferramentas/reteste04_usuario_f2_hema.py` (modos: padrão, `--parte2`, `--so-calc`), documentado no cabeçalho. Reaproveita `qa_final.QA`, `repro_comum`, `valida_adr090_091` (Audit, Vigia, manual, `abrir_ceq_sem_sessao`) e, do irmão da Bio `reteste04_usuario_f2.py`, a varredura `s_adr094` e a diferença célula a célula da Estatística. Os oráculos das regras ADR-091 e do Painel são meus. Matemática: `calc_independente.py` (`auditar_hema`) |
| Saídas (fora do repositório) | `Hematologia_F2.json/.log`, `Hematologia_F2_p2.json/.log`, `Hematologia_F2_calc.json`, `Hematologia_F2_calcR.json`, `calcH_*.json` |
| Usuário | `TESTE_AUDITORIA_F2H`, papel ANALISTA, criado só na cópia; login por `mSeguranca.DoLogin` |
| Dados de teste | CEQ: rodada "TESTE-AUDITORIA F2H R1", amostras TA-F2H-*, arquivo TESTE-AUDITORIA.pdf, HGB ("Hemoglobin") 14,3 → 13,65 sobre alvo 13 (+10 % → +5 %); 26/13 (+100 %) no Controllab e no CAP 2024. Linha de mapa "TESTE-AUDITORIA HGB CTL" → HGB (Controllab) na EQA_Base W33:Y33 da cópia, apagada depois. Analito "TESTE-AUDITORIA X" na Especificações!A32, apagado depois. Regras da tblLotesMaterial com NOTA "TESTE-AUDITORIA" (533811 FORA_DO_CQ; 536411 SO_ANALITOS WBC). Manuais MAN_0001..0006 com MOTIVO "Outro: TESTE-AUDITORIA" (−1 / 7,77 / 4,44). Inativação de HEM-374033 e comentário em HEM-373958, só na cópia |
| tblCQ_Final | 48.391 linhas; nenhum resultado novo do DB_SEAC durante o teste (último recebido 13/11/2026 13:10:20) |

**Caminho usado.**
- **Filtros:** Painel C3/H3/K4 e Estatística R4 escritos nas células, com os eventos ligados, como o usuário.
- **CEQ:** pelo núcleo que os botões da janela chamam (`mCEQForm.AvisoAnalitoCEQ`, `ConfirmacoesCEQ`, `SalvarCEQ`,
  `CorrigirCEQ`, `ExcluirCEQ`, `DescricaoCEQ`, `RegistrosDaRodada`) e pelo `AbrirCEQ` real.
- **Manuais:** na aba Digitar Resultados (botão `mIntegracao.IrDigitarResultados`, digitação célula a célula, eventos
  ligados), seguidos do ATUALIZAR DADOS pelo núcleo do botão (`mIntegracao.AtualizarDadosAutomatico`; o DB_SEAC só é
  lido).
- **Inativar e comentar:** pelo núcleo da janela do ponto (`InativarPorId`, `GravarComentarioTecnico`).

**Pelo VBA interno, sem o caminho do usuário:**
- **Tarefas de administrador**, feitas com a aba liberada:
  - a linha de mapa do Controllab na EQA_Base (eventos desligados);
  - o analito de teste na Especificações!A32 (eventos desligados + `Calculate`);
  - as duas regras na tblLotesMaterial (aba Cfg_Integracao).
- **Não executado:** o clique nas caixas Sim/Não do frmCEQ (não há interação com a tela).

**Incidente de execução:** na parte 1, a seção de regras e manuais parou num índice do meu script antes do ATUALIZAR.
O WBC N1 do lote 625311 tem só 38 resultados, e o script pedia o 41º de trás para frente. Corrigi o índice e refiz a
seção inteira numa **segunda cópia nova** (parte 2). Nada da parte 1 foi aproveitado para esses itens.

## Resultado por item do reteste

| ID | Estado | Evidência |
|---|---|---|
| OC-0049 | **RESOLVIDA** | Arquivo recém-aberto: **nenhum texto** em A98:AU133 (as 36 linhas sem analito; 0 de 1.692 células com `Text`, 0 valores). B98 = `=IF($A98="","",1)`, e G, R, S e T testam `$A98=""`. As linhas 14..97 têm os 28 analitos, com os níveis 1/2/3 na ordem. Vizinho: cadastrar "TESTE-AUDITORIA X" na Especificações!A32 faz as linhas 98..100 mostrarem o nome, os níveis 1/2/3, "SEM EP" em G/R/S/T e "-" em K/L (ETp ausente, ADR-068). A linha 101 continua vazia. Ao apagar o analito, as linhas voltam a ficar vazias. A14:AU97 fica igual bit a bit antes e depois. A calculadora sobre o retrato E0 deu 2.268 comparações + 45 do Painel, **0 divergências** |
| OC-0057 | **RESOLVIDA** | Registro com sessão: "OK\|467\|" (14,3). Depois do `Logout` (SessaoAtiva False), Salvar, **Corrigir** e **Excluir** devolvem "ERRO\|0\|Sessão encerrada: faça o login de novo. Nada foi gravado." F467 continua 14,3, a linha não é apagada (A467 "CAP"), a tabela fica com 467 → 467 linhas e não há nenhuma trilha CEQ_*. `AbrirCEQ` sem login mostra a caixa "Controle externo (CEQ)" com "Faça o login para registrar resultados do CEQ." (lida pelo texto da janela do meu PID), e a janela não abre. Vizinho: com novo login, Corrigir dá "OK\|467\|" (F = 13) e Excluir dá "OK\|467\|", com trilha CEQ_ALTERADO + EQA_BASE_CONSOLIDADA + CEQ_EXCLUIDO + EQA_BASE_CONSOLIDADA. Nenhuma linha TESTE-AUDITORIA sobra na EQA_Base |
| OC-0058 | **RESOLVIDA** | Trilha de cada manual no 1º ATUALIZAR: −1 no 625311 (IMPLAUSIVEL, MAN_0001), 533811 WBC 7,77 e 533811 −1 (LOTE_FORA_DO_CQ, MAN_0002/0003; a regra do lote vem antes do IMPLAUSIVEL) e 536411 RBC (LOTE_OUTRO_MATERIAL, MAN_0004) gravaram **RESULTADO_MANUAL_FORA_DA_ESTATISTICA**. Os ATIVOS (536411 WBC, MAN_0005; 625311 WBC N2, MAN_0006) gravaram **RESULTADO_MANUAL_INCLUIDO**. Nenhuma trilha fora dos manuais. O 2º ATUALIZAR (só entradas, 20,7 s) não repetiu nenhuma trilha de manual |
| OC-0066 | **RESOLVIDA** | O cenário é o exato da OC: na produção da Hema a coluna tblCQ_Final[DATA_INATIVACAO] estava **vazia**. Antes do teste, 0 linhas preenchidas, largura 17,71, formato "dd/mm/aaaa hh:mm:ss" na 1ª e na última linha. Depois de inativar HEM-374033 e rodar o ATUALIZAR, a coluna mostra "14/11/2026 16:47:59" (sem "####"). Sync_Log: 2 → 3 linhas, e a nova tem INICIO "14/11/2026 16:49:36" e FIM "14/11/2026 16:49:57". DATA_MIN/MAX "01/10/2025 06:37:53" / "13/11/2026 13:10:20". Nenhum "####" na tblSyncLog |
| ADR-091 | **RESOLVIDA** (regra) — ver **R4-F2-02 (Hema)** (mensagem do Painel) | Produção: tblLotesMaterial com o cabeçalho LOTE/REGRA/MATRIZ/ANALITOS/NOTA e uma linha vazia, e **0** linhas LOTE_* (ADR-091 §5: "Hema: nenhuma regra"). Com as regras de teste e o ATUALIZAR, o meu oráculo sobre a tblCQ_Final dá **0 violações**. 533811 (FORA_DO_CQ): 6.574 linhas, 6.481 LOTE_FORA_DO_CQ + 93 SEM_VALOR (precede), nenhuma participando. 536411 (SO_ANALITOS WBC): WBC 171 ATIVO + 9 DUPLICIDADE_ORIGEM; os outros 27 analitos têm 5.101 LOTE_OUTRO_MATERIAL + 270 DUPLICIDADE_ORIGEM. Nenhum lote sem regra tem LOTE_*. QA: **A13 = 11.582 = total LOTE_***, A14 = 0. Principal A5: "Total: 48.397 · participam: 36.131 · X vermelho: 1 · inativados: 1 · manuais: 6". Painel 536411 + WBC: n 59/57/55 = minha contagem dos que participam. Efeito na estatística: u(Rw) 12 m do HGB passou de 8 para 6 lotes (AF N1 1,17396 → 0,93500), igual à calculadora. MATRIZ: não executado (ver abaixo) |
| ADR-094 | **RESOLVIDA** | Varredura de **todas** as tabelas de todas as abas depois do inativar, do comentar, dos manuais e dos dois ATUALIZAR: nenhuma data-hora sem segundos e nenhum "####" nas abas do escopo. Valores lidos: Principal (DATA_HORA "09/10/2026 04:11:00", RECEBIDO_EM "14/11/2026 16:47:45", DATA "09/10/2026"); Inativar (largura 18) e Comentários (17) "14/11/2026 16:47:59"; Digitar Resultados REGISTRADO_EM "14/11/2026 16:47:57"; DB A5 "último resultado: 13/11/2026 13:10:20 · última chegada: 14/11/2026 15:20:18"; QA e Sync_Log com segundos; Registros "26/05/2025" (só data). Fora do escopo do ADR (ver Observações): BI_Data e Audit_Log Data/DataCorrida |

### Pedido do ciclo: incluir e corrigir um CEQ TESTE-AUDITORIA (HGB)

| ID | Teste | Estado | Evidência |
|---|---|---|---|
| CEQ-INC | HGB, CAP "Hemoglobin", Ano EP 2025, rodada TODAS, lote 522611 (B4), anos 2025–2025: incluir 14,3/13 (+10 %) | **APROVADO** | "OK\|467\|"; aviso "AZUL\|No QC_INI: HGB. Este resultado entra no bias deste analito."; EQA_Base com Analito_Canonico HGB, Bias 10,000000000000005, Uso SIM; lista da rodada "Hemoglobin · TA-F2H-01 · 14,3 · Acceptable · 467"; Z3 "base em dia"; K5 "4 rodada(s), 361 linha(s)". Mudaram **72 células, todas nas linhas 20–22 (HGB)**, mais a coluna U ("ordem crítico") de 14 linhas de outros analitos, renumerada porque o HGB entrou na lista de críticos (U = contagem acumulada de "ETp excedido/Margem crítica"; nenhum número de outro analito mudou). Bias 1,44311 → 3,58233; ET N1 3,24390 → 5,38313; Sigma N1 2,96712 → 0,48468. Painel B7:I9 (C3 = HGB) com os mesmos valores. A Estatística atualiza sem `Calculate` (o resultado é o mesmo com e sem ele) |
| CEQ-COR | Corrigir para 13,65 (+5 %) na mesma linha | **APROVADO** | "OK\|467\|"; trilha "Seu resultado: 14,3 → 13,65" + consolidação (alteradas 1). 52 células do HGB + U renumerado. Bias 2,33233; ET N1 4,13313; Sigma N1 1,93523, N2 2,37319, N3 2,81464; classe P N2/N3 volta a "Dentro do orçamento" |
| CEQ-ANO | CAP **2024** (+100 %) com Ano EP 2025 | **APROVADO** | Estatística sem nenhuma mudança |
| CEQ-PROV | Controllab do HGB: sem mapa e com mapa de teste (R4 = CAP); depois R4 = Controllab e volta | **APROVADO** (cálculo) — ver **R4-F2-01 (Hema)** (mensagem) | Sem mapa: aviso âmbar "não está no mapa do Controllab: este resultado NÃO entra no bias…" e Salvar "OK\|2\|Gravado fora do bias (analito não mapeado)." Com o mapa: aviso azul "entra no bias" e Salvar "OK\|3\|". **0 células mudaram** com R4 = CAP. R4 = Controllab: K5 "Controllab \| ano 2025 \| 1 rodada(s), 1 linha(s)", só o HGB tem bias (100), e os outros 27 analitos ficam SEM EP. Ao voltar para R4 = CAP, a Estatística fica **igual bit a bit** ao S3. CAP e Controllab não se misturam. Trilha PROVEDOR_ALTERADO nas duas trocas |
| CEQ-DUP | Duplicidade | **APROVADO** | "ERRO\|5\|Já existe este registro (CAP TESTE-AUDITORIA F2H R1 2025 · Hemoglobin · TA-F2H-01) na linha 467. Para corrigir, selecione-o em “Registros desta rodada” e clique em Corrigir selecionado." |
| CEQ-EXC | Excluir tudo (4 registros) + apagar o mapa | **APROVADO** | Estatística A14:AU133 **igual bit a bit** ao S0. EQA_Base sem nenhuma linha TESTE-AUDITORIA. K5 de volta a "3 rodada(s), 360 linha(s)". 4 trilhas CEQ_EXCLUIDO, cada uma com EQA_BASE_CONSOLIDADA |
| CEQ-CALC | Calculadora independente sobre os 8 retratos | **APROVADO** | `calc_independente.auditar_hema`: 2.268 comparações (84 linhas × 27 colunas) + 45 do Painel por retrato, **0 divergências** em E0, S0, S1, S2, S3, RX, S4 e R1 (tolerância relativa de 1e-9; V com a tolerância absoluta já justificada no cabeçalho da calculadora) |
| ADR-091-MZ | Regra MATRIZ na Hema | **NAO_EXECUTADO** | A Hema tem uma matriz só (SANGUE TOTAL), o de/para não tem outra, e a carga do ADR-091 não tem regra. Inventar uma matriz para a Hema seria testar um cenário que o laboratório não tem. A regra MATRIZ foi testada na Bio (Bioquimica_F2.md) |

## Relatório B — validação matemática do CEQ (HGB)

**Definição** (conferida em `src/comum/mCEQ.bas` `BiasEQ`, `AnoVigente` e `AmostraDoVies`, e nas fórmulas G/H/L da
Estatística da Hema antes de usar):
- **Amostras:** as da EQA_Base com Analito_Canonico = HGB, Uso_Analitico ≠ NAO e Provedor = Estatística!R4 (global na
  Hema), no ano N4 = 2025, com rodada P4 = TODAS. Ficam fora as NAO AVALIADO, as com resultado ou alvo 0 e as chaves
  repetidas.
- **bias_i** = (Resultado − Alvo)/Alvo × 100, recalculado por mim do Resultado e do Alvo.
- **Consolidação em duas etapas:** a média dos |bias| de cada rodada e, depois, a média das rodadas.
- **ET** = 1,65 × F + |G|, com F = CV do lote B4 = 522611 (Painel H3/Configuração C20) nos anos B3..D3 = 2025.
- **Sigma** = (K − |G|)/CVp, com CVp = CV pooled da janela 14/05–13/11/2026 (6 meses até o último resultado), XN1000,
  lotes com n ≥ 20.
- **ETp** = 4 % (CLIA, Especificações S/T) nos três níveis.

**Etapas:**
- Médias de |bias| por rodada, 5 amostras cada: FH9-A 2025 = 2,3593073054506553; FH9-B = 1,1137368437415163;
  FH9-C = 0,8562947636618198.
- As 15 linhas da simulação (rodadas "1/2/3", Uso NAO) não entram.
- CVp, da calculadora (igual ao Painel C45:C47, conferido nas 45 comparações do Painel): N1 0,8617391643040556,
  N2 0,7027095883994274, N3 0,5924957915732495.
- F: N1 1,091388640832613 (n 50), N2 0,6959305974249244 (n 51), N3 0,6648898448424729 (n 51).

| Analito | Equip. | Lote | Nível | Período | Dados de entrada | Fórmula | Independente | QC_INI (Value2) | Dif. rel. | Status |
|---|---|---|---|---|---|---|---|---|---|---|
| HGB | XN1000 | (bias não depende) | 1–3 | CEQ CAP 2025 | 3 rodadas × 5 | (2,35931 + 1,11374 + 0,85629)/3 | 1,4431129709513304 | 1,4431129709513304 | 0 | APROVADO |
| HGB | XN1000 | 522611 | 1 | 2025 | F 1,09139; G 1,44311 | H = 1,65·F + G | 3,2439042283251407 | 3,2439042283251416 | 3e-16 | APROVADO |
| HGB | XN1000 | — | 1 / 2 / 3 | Sigma 14/05–13/11/2026 | CVp acima | L = (4 − G)/CVp | 2,96712408460 / 3,63861127165 / 4,31545179799 | 2,967124084592441 / 3,638611271652592 / 4,315451798019849 | ≤ 6e-12 | APROVADO |
| HGB | XN1000 | — | 1–3 | + TA-F2H-01 14,3/13 | + rodada TESTE 10,000000000000005 (1) | (4,32934 + 10)/4 | 3,582334728213499 | 3,582334728213499 | 0 | APROVADO |
| HGB | XN1000 | 522611 | 1 / 2 / 3 | idem | F | H | 5,383125985587309 / 4,730620213964625 / 4,6794029722035795 | 5,383125985587311 / 4,730620213964624 / 4,6794029722035795 | ≤ 4e-16 | APROVADO |
| HGB | XN1000 | — | 1 / 2 / 3 | idem | CVp | L | 0,4846771379177 / 0,5943639857510 / 0,7049252968995 | 0,48467713791667294 / 0,5943639857509238 / 0,7049252969037589 | ≤ 6e-12 | APROVADO |
| HGB | XN1000 | — | 1–3 | idem | com sinal | T | 3,512232107163426 | 3,512232107163426 | 0 | APROVADO |
| HGB | XN1000 | — | 1–3 | corrigido 13,65/13 | rodada TESTE 5,000000000000003 | (4,32934 + 5)/4 | 2,3323347282134983 | 2,3323347282134987 | 2e-16 | APROVADO |
| HGB | XN1000 | 522611 | 1 / 2 / 3 | idem | F | H | 4,133125985587308 / 3,480620213964624 / 3,429402972203579 | 4,13312598558731 / 3,480620213964624 / 3,429402972203579 | ≤ 5e-16 | APROVADO |
| HGB | XN1000 | — | 1 / 2 / 3 | idem | CVp | L | 1,9352320758606 / 2,3731927090748 / 2,8146449232295 | 1,9352320758563868 / 2,3731927090746896 / 2,8146449232463637 | ≤ 6e-12 | APROVADO |
| HGB | — | — | 1–3 | + CAP 2024 (+100 %) | ano ≠ N4 | não entra | sem mudança | sem mudança (0 células) | 0 | APROVADO |
| HGB | — | — | 1–3 | + Controllab 26/13 mapeado, R4 = CAP | provedor ≠ R4 | não entra | sem mudança | sem mudança (0 células) | 0 | APROVADO |
| HGB | XN1000 | 522611 | 1 / 2 / 3 | R4 = Controllab | 1 amostra, 100 % | G; H; L | 100; 101,80079125737382 / 101,14828548575113 / 101,09706824399008; −111,40261923401 / −136,61404595127 / −162,02646730214 | 100; idem H; −111,40261923376998 / −136,61404595126515 / −162,026467303112 | ≤ 6e-12 | APROVADO |
| HGB | — | — | 1–3 | tudo excluído | 3 rodadas × 5 | idem base | 1,4431129709513304 | 1,4431129709513304 (S4 = S0 bit a bit) | 0 | APROVADO |
| HGB | XN1000 | 6 lotes | 1 / 2 / 3 | u(Rw) 12 m, depois das regras ADR-091 | 533811 e 536411 fora | CV pooled 14/11/2025–13/11/2026 | 0,934995864944362 / 0,6888645385444551 / 0,5844415458004871 | 0,9349958649446772 / 0,6888645385452185 / 0,5844415457960841 | ≤ 8e-12 | APROVADO |
| Todas (84 linhas) | XN1000 | 522611 | 1–3 | E0, S0..S4, RX, R1 | — | calc_independente | — | — | 0 divergência em 8 × 2.313 | APROVADO |

A diferença de até 6e-12 relativo no Sigma vem da ordem de soma do CV pooled; ela é igual à da F3 do ciclo 01 e fica
dentro de 1e-9.

**Caminho do dado (conferido nas trilhas e nos retratos):**
1. A janela do CEQ grava na aba EQA.CAP_Dados (linha 467) ou EQA.Controllab_Dados.
2. A EQA_Base é consolidada sozinha: Analito_Canonico pelo mapa W:Y, e Bias e |Bias| recalculados.
3. `BiasEQ` (provedor R4, ano N4, rodada P4) alimenta a Estatística G/T/AC/AD.
4. Daí saem H (ET, com o F do lote/ano), L (Sigma, com o CVp de 6 meses), M e P (classes), U (ordem dos críticos),
   V/W (DPM e rendimento), X..AA (plano), AJ..AU (incerteza e viés) e K5.
5. O Painel B7:I9 (G/I leem a Estatística, H recalculado) mostra o analito do C3.

No CIQ, a tblCQ_Final (PARTICIPA = SIM) alimenta o Eng_Saida engEstat (F) e o `IncertezaCIQ` (CVp e u(Rw)).

## Ocorrências novas (extensões para a Hema de achados da Bio deste ciclo)

### R4-F2-01 (Hema) — Janela do CEQ: nome do Controllab mapeado aparece como "entra no bias" com o R4 = CAP

- **Prioridade:** Médio · **Tipo:** usabilidade · **Tela/módulo:** janela "Registrar resultado do CEQ"
  (`mCEQForm.AvisoAnalitoCEQ`, `SalvarCEQ`), Hematologia. É a mesma causa do R4-F2-01 da Bio, que lá ficou como "não
  testado na Hema".
- **Ação:** com o R4 da Estatística = CAP (provedor de todos os analitos), registrar uma amostra do Controllab de um
  nome mapeado para HGB.
- **Esperado:** o aviso dizer que, com o provedor atual (CAP), o resultado fica gravado **fora** do bias, e que só
  entra se o R4 for trocado para Controllab.
- **Observado:** "AZUL\|No QC_INI: HGB. Este resultado entra no bias deste analito." e `SalvarCEQ` "OK\|3\|", sem nota.
  Na Estatística, 0 células mudam, o que está certo. Só com o R4 = Controllab o HGB passa a usar o resultado (G = 100).
- **Dificuldade / consequência:** o analista acha que o CEQ do Controllab já pesa no bias, no ET e no Sigma, e não
  percebe que o número continua sendo o do CAP. Hoje é **latente**: o mapa do Controllab da Hema está vazio, e
  qualquer nome cai no aviso âmbar correto, que confirmei.
- **Sugestão:** comparar o provedor digitado com o R4: "No QC_INI: HGB. **O bias de todos os analitos usa hoje o CAP
  (R4)**: este resultado fica gravado e só entra se o R4 for Controllab."
- **Evidência:** `Hematologia_F2.json` (`ceq.aviso_ctl_mapeado`, `ceq.S3_vizinhos.controllab_mapeado`,
  `ceq.S3_vizinhos.diff_S2` = [], `ceq.RX_controllab`); `src/comum/mCEQForm.bas` (`AvisoAnalitoCEQ`).

### R4-F2-02 (Hema) — Painel em lote com regra FORA_DO_CQ: "SEM CORRIDA" e "O lote não tem resultado de WBC", sem citar a regra

- **Prioridade:** Médio · **Tipo:** usabilidade · **Tela:** Painel (J7:J9, Q4, A39), ADR-091. É a mesma causa do
  R4-F2-02 da Bio.
- **Ação:** com a regra de teste 533811 FORA_DO_CQ, escolher WBC e o lote 533811 com K4 "Tudo (sem filtro)".
- **Esperado:** nenhum ponto, o que está certo, e um texto que diga por quê ("lote fora do CQ pela tblLotesMaterial:
  resultados guardados, fora da estatística"), como já dizem o DETALHE e o QA A13.
- **Observado:** J7:J9 "SEM CORRIDA", Q4 "SEM CORRIDA DESTE ANALITO" e A39 "Lote 533811 · alvo … O lote não tem
  resultado de WBC no XN1000." O lote **tem** 214 resultados de WBC (80/68/66 por nível), todos LOTE_FORA_DO_CQ.
- **Consequência:** o usuário procura um resultado que "não chegou". Não muda a estatística. Hoje é latente na Hema,
  porque a tabela está vazia na produção.
- **Sugestão:** a mesma da Bio. Com regra ativa para o lote, o A39/Q4 dizem "Lote 533811 fora do CQ (tblLotesMaterial)".
- **Evidência:** `Hematologia_F2_p2.json` (`painel_lotes.WBC_533811`).

## Observações (não são ocorrências desta frente)

- **Lote SO_ANALITOS + analito fora da lista:** com C3 = RBC e H3 = 536411 (regra "só WBC"), o Painel trocou sozinho o
  analito para WBC. O A39 avisou "FILTRO: Lote 536411: o filtro anterior não mostrava nenhum resultado dele --
  ajustado: analito WBC", e o H4 disse "ver o aviso abaixo dos gráficos". Ajuda a achar os dados, mas o usuário que
  queria o RBC não fica sabendo que o RBC foi excluído pela regra (ver R4-F2-02).
- **PROVEDOR_ALTERADO "(não registrado) → Controllab":** o "antes" é guardado no `Worksheet_SelectionChange` do R4. Eu
  escrevi no R4 pelo COM sem selecioná-lo, e por isso o antes ficou vazio na 1ª troca. A 2ª troca registrou
  "Controllab → CAP" certo. Pela tela o usuário sempre seleciona o R4 antes de trocar, então isso é artefato do teste
  (o mesmo padrão do OC-0054 do CEQ).
- **BI_Data** (aba do BI, fora do escopo do ADR-094 e proibida de mexer): Usar_1_3s, Usar_R_4s, Usar_3_1s e Usar_6x
  mostram "##########" (largura 8,86–10,86) e a Data está em aaaa-mm-dd. É o mesmo da Bio.
- **Audit_Log Data/DataCorrida** só com data: decisão do ADR-094 ("O que NÃO muda", hash).
- **DETALHE sem acento** nos textos ADR-091/IMPLAUSIVEL ("estatistica", "grafico", "implausivel"): já é a OC-0021
  PERSISTE do F1 deste ciclo.
- **Tempos:** Salvar/Corrigir/Excluir do CEQ entre 15 e 25 s (trilha), dentro do GUIA §5. O ATUALIZAR completo com as
  regras novas levou **74 s** (motor 51,5 s; westgard_fotos 40,5 s, com 468 mudanças de status a refotografar). O só
  entradas levou 20,7 s.

## Dúvidas de regra de negócio

1. **R4-F2-01 na Hema:** como o R4 é global e pode ser trocado, o aviso deve dizer "entra no bias se o R4 for
   Controllab" ou recusar o mapa do Controllab enquanto o R4 for CAP?
2. **Sigma negativo:** com |bias| > ETp (R4 = Controllab, bias 100 %), a Estatística e o Painel mostram Sigma −111,40
   e a classe "Desempenho inadequado". A conta segue a definição, (ETp − |bias|)/CVp, e a calculadora confere. Mostrar
   um Sigma negativo é o desejado, ou deveria aparecer "< 0"/"ETp excedido pelo bias"? Pergunta para o Especialista em
   CQ.
3. **BI_Data visível com "####"** (o mesmo da Bio): ocultar para quem não é ADM?
4. **Lote SO_ANALITOS no Painel:** trocar o analito sozinho (RBC → WBC) é o comportamento desejado para um lote de um
   analito só, ou o Painel deveria ficar no RBC e dizer "lote só de WBC (regra)"?

## Impressões de uso (Relatório A)

- **CEQ, inclusão e correção:**
  - A janela diz para qual analito o resultado vai, e a lista da rodada mostra o que acabou de ser gravado.
  - A duplicidade diz a linha e o botão.
  - O número muda na hora e só no HGB. A coluna "ordem crítico" dos outros analitos se renumera, o que é correto,
    mas um usuário que compare duas impressões pode estranhar.
  - Excluir volta exatamente ao que era.
  - É o fluxo mais confiável da Hema.
- **Provedor global (R4):**
  - A troca é imediata e reversível, e o K5 diz claramente o que entrou ("Controllab \| 1 rodada, 1 linha").
  - A troca fica no Audit_Log.
  - Falta a janela do CEQ saber do R4 (R4-F2-01).
- **Sessão:** sem login a janela nem abre; com a sessão encerrada no meio, nada é gravado e a mensagem é clara.
- **Estatística:** as linhas sem analito ficaram limpas (sem "0" e sem "SEM EP"). Um analito novo aparece sozinho na
  1ª linha livre, com os três níveis.
- **Regras de lote:**
  - O DETALHE de cada manual diz a regra que o excluiu, e o QA A13 conta certo.
  - O ATUALIZAR ficou mais lento na 1ª vez depois de uma regra nova (74 s contra cerca de 30 s), porque o Westgard
    refotografa os pontos que mudaram de status. Vale avisar no GUIA.
  - O Painel ainda não explica a regra (R4-F2-02).
- **Datas:** segundos em todas as abas do dia a dia, sem "####", inclusive na DATA_INATIVACAO que estava vazia na
  produção e na linha nova do Sync_Log.

## Arquivos

- **Script:** `_arquitetura/agentes/ferramentas/reteste04_usuario_f2_hema.py`.
- **Calculadora:** `_arquitetura/agentes/ferramentas/calc_independente.py`.
- **Saídas (fora do repositório):** `%TEMP%\qcw\auditoria\2026-11-14_reteste04\Hematologia_F2*.{json,log}` e
  `calcH_*.json`. Os retratos `.xlsm` foram apagados; o script os recria.
- **Dados pessoais:** o Audit_Log da cópia grava o nome, o login do Windows e a máquina de quem roda. Esses dados foram
  omitidos deste relatório.
