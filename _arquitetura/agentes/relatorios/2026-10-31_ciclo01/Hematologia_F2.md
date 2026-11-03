# Hematologia, F2 (CEQ): ciclo 2026-10-31_ciclo01

Agente: Usuário Simulado. Modo: completo. Data: 31/10/2026, das 22h05 às 22h45 (Excel das 22h19 às 22h23 e das
22h24 às 22h28, um de cada vez).

A frente refez na Hematologia o roteiro da F2 da Bioquímica, com uma ferramenta própria (o layout da Hema é outro):

- **Run A:** **RBC** ("Red Blood Cell Count" no CAP), ETp CLIA de 4 %, o analito do Painel. O controle é o **WBC**
  ("White Blood Cell Count", o nome mais parecido).
- **Run B:** **LYMPH#** ("Lymphs Absolute"), ETp de **outra fonte (VB, 22,758 %)**. O controle é o **LYMPH%**
  ("Lymphs"), o vizinho mais fácil de confundir.

Nos dois runs, um resultado de CEQ `TESTE-AUDITORIA` foi **incluído** (110 / alvo 100, bias +10 %) e depois
**corrigido** (105, bias +5 %). Conferi que o bias, o ET e o Sigma do analito certo, e só dele, mudam como a definição
manda.

## 1. Ambiente e evidência

| Item | Valor |
|---|---|
| Produção copiada | `QC_Hematologia.xlsm`, sha256 curto `0b45fa270d2a` (o mesmo da F1 e da F3). O hash é igual antes e depois: a produção não foi aberta |
| Cópias | `%TEMP%\qcw\auditoria\2026-10-31_ciclo01\hema_f2\A\Hematologia_F2_A.xlsm` (Excel PID **28268**) e `...\hema_f2\B\Hematologia_F2_B.xlsm` (PID **19392**). Cada uma é uma cópia nova da produção (`shutil.copy2`), com hash igual antes e depois (nunca salva). Cada PID foi criado e encerrado pelo próprio script, no `finally`. Nenhum outro Excel estava aberto antes nem ficou aberto depois |
| Retratos para a calculadora | `f2h_S1.xlsm` (depois da inclusão) e `f2h_S2.xlsm` (depois da correção) de cada run, por `SaveCopyAs` com os eventos desligados. Cópias e retratos foram apagados no fim |
| Ferramenta nova (reutilizável) | `_arquitetura/agentes/ferramentas/f2_ceq_hema.py`, com os passos P0 a P7 documentados no cabeçalho. É a adaptação do `f2_ceq.py` à Hema: Estatística A14:AU97, 3 níveis, provedor global em Estatística!R4, Painel B7:J9, lote H3, período Ano De/Até. O `f2_ceq.py` da Bio não foi alterado |
| Comandos | `python -u f2_ceq_hema.py <copia_A> <pasta_A>` (padrão RBC / "Red Blood Cell Count" / WBC / lote alternativo 619611) e `python -u f2_ceq_hema.py <copia_B> <pasta_B> "LYMPH#" "Lymphs Absolute" "LYMPH%" 619611 "x 10E9/L"`. Depois, `calc_independente.py` sobre S1 e S2 de cada run. Saídas (só JSON, TXT e log) em `...\hema_f2\A\` e `...\hema_f2\B\` |
| Usuário | Um usuário de teste `TA_F2H_ANALISTA`, papel **ANALISTA**, criado só na cópia, com login pelo `mSeguranca.DoLogin`. A senha de teste vem da suíte e não aparece aqui |
| Dados de teste | Rodadas `TESTE-AUDITORIA R1` (2025), `TESTE-AUDITORIA R2` (2026) e `TESTE-AUDITORIA CTL1` (Controllab). Amostras `TA-01`, `TA-R2`, `TA-03`, `TA-04`, `TA-C1` e `TA-SEMLOGIN`. Arquivo-fonte `TESTE-AUDITORIA.pdf`. Valores redondos e sintéticos à vista: um RBC de 110 × 10¹²/L é impossível. Resultado 110 → 105, 120, 130, 150 e 100; alvo 100, SD 2, SDI 5, limites 90/110, nº lab 50 |
| Filtros da cópia | Painel: C3 = RBC (run A) / LYMPH# (run B), H3 = 522611, K3 = 2025, K4 = "Tudo (sem filtro)". Estatística: Ano De/Até B3:D3 = 2025, Ano EP N4 = 2025, Rodada P4 = TODAS, Provedor R4 = CAP, Sigma em 6 meses (02/04–01/10/2026). Equipamento único XN1000 |

**Como os passos foram executados**

- **Gravação:** pelo núcleo da janela, `mCEQForm.SalvarCEQ`, que é o que o botão "Salvar e próximo" do `frmCEQ` chama,
  com os 16 campos na ordem da janela. A janela não foi aberta: ela é modal e travaria a automação. O comportamento
  dela (unidade automática, lista da rodada) foi lido em `frmCEQ.frm.txt`, e as funções que ela usa foram chamadas.
- **Correção:** escrevi direto na célula da tabela `tblEQA_CAP_Dados`, com os eventos ligados, como a mensagem da janela
  manda. Depois rodei o botão "Consolidar na EQA_Base" (OnAction `AtualizarEQABase`, em J1 da aba de dados).
- **Filtros:** Painel C3/H3 e Estatística N4/R4/B3/D3 foram mudados pela própria célula, com os eventos ligados. Isso
  exigiu tirar a proteção dessas abas, um passo técnico. Em B3/D3, chamei também `mEstatistica.AtualizarEstatisticaAba`,
  porque a aba Estatística da Hema não tem evento de planilha. Desvio do caminho do usuário, declarado aqui.
- **Limite do método:** uma escrita por COM numa aba protegida com `UserInterfaceOnly` passa mesmo em célula bloqueada.
  O bloqueio que o usuário encontraria foi conferido pelas propriedades `Locked` e `ProtectContents`, e não pelo
  teclado.

## 2. Testes executados

| ID | Teste | Estado | Evidência |
|---|---|---|---|
| HF2-T01 | Telas do CEQ: onde se registra, botões, proteção | APROVADO (com ressalvas: OC-HF2-1 e OC-HF2-8) | "Registrar resultado do CEQ" (`mCEQForm.AbrirCEQ`) fica em A3 de EQA.CAP_Resumo e de EQA.Controllab_Resumo. "Consolidar na EQA_Base" (`AtualizarEQABase`) fica em J1 das duas abas de dados. O Início tem o bloco "Controle externo (EQA)" (G15). `frmCEQ` instalado. Nenhuma aba EQA tem evento de planilha. A tabela do CAP tem 24 colunas (A..R digitadas e S..X auxiliares), sem validação de dados, e todas `Locked` para o ANALISTA, na tabela e abaixo dela |
| HF2-T02 | Preenchimento automático e listas da janela | APROVADO (ressalva OC-HF2-8) | A rodada traz o ano (`FH9-C 2025` → 2025). A rodada inexistente não traz nada. O analito traz a unidade ("Red Blood Cell Count" → x 10E12/L). Rodadas em ordem decrescente: FH9-C, FH9-B, FH9-A 2025, 3, 2, 1. A lista da FH9-C 2025 tem 125 registros; a de ano 2026, nenhum |
| HF2-T03 | Validações antes de gravar (12 casos) | APROVADO | Cada caso foi recusado com o campo certo: resultado vazio (6), resultado em texto (6), alvo em texto (7), limite inferior > superior (10), ano "25" (3), SD −1 (8), avaliação vazia (12), provedor "PNCQ" (1), rodada vazia (2), amostra vazia (5) e analito vazio (4). A tabela ficou idêntica (465 linhas antes e depois), nos dois runs |
| HF2-T04 | Duplicidade de uma chave **real** (FH9-A 2025 · Basophils · FH9-01), também com outra caixa e espaços ("fh9-a 2025", "basophils ", " fh9-01") | APROVADO | Os dois casos devolvem "Já existe este registro (...) na linha 2 da aba EQA.CAP_Dados. Para corrigir, edite a linha existente." Nada gravado |
| HF2-T05 | Inclusão: gravação, consolidação, trilha, lista da rodada (run A) | APROVADO | `OK\|467\|` em 23,9 s (run B: 18,8 s). A tabela ganhou uma linha (465 → 466), com O = P = 10 e as colunas auxiliares preenchidas. A EQA_Base recebeu a linha com Canônico = RBC, Uso = SIM, Bias 10 e chave `CAP\|2025\|TESTE-AUDITORIA R1\|Red Blood Cell Count\|TA-01`. Carimbo "CAP 466 \| total 466". Audit_Log: CEQ_REGISTRADO (AUD-000030). A lista da rodada mostra `Red Blood Cell Count · TA-01 · 110 · Acceptable`. EQA.CAP_Resumo mostra a rodada nova. K5: 3 → 4 rodadas, 360 → 361 linhas |
| HF2-T06 | Inclusão: o bias, o ET e o Sigma do RBC mudam como deveriam | APROVADO | Seção 3. Bias de 1,5336 % para 3,6502 % (4 rodadas, 16 amostras). ET N1/N2/N3 de 3,612 / 3,110 / 3,052 % para 5,728 / 5,226 / 5,169 %. Sigma de 2,528 / 3,092 / 3,055 para 0,358 / 0,439 / 0,433. Diferença máxima em relação ao independente: 1,3e-15 |
| HF2-T07 | Inclusão: **só** o RBC muda | APROVADO (com observação) | Mudaram só RBC\|1..3 (G, H, L, M, N, O, P, R, S, T, U, V, W, Y, Z, AC, AD, AJ–AM, AQ, AR, AT, AU). WBC (controle) ficou idêntico. Nas outras linhas, só a coluna **U** ("ordem crítico") mudou, em 14 células: HCT, MCV, MCH, MCHC e MPV, todos +2. O RBC entrou em "ETp excedido" nos 3 níveis e empurrou a ordenação. É o mesmo efeito da Bio (F2-T25): uma ordenação, não contaminação. G, H, L e as demais estatísticas desses analitos não mudaram. O Painel B7:J9, O4, J/M/O45:47 e C/H61:63 acompanharam a Estatística. C45:C47 (CVp) e B:E (n, média, DP, CV) não mudaram |
| HF2-T08 | Correção pelo caminho do ANALISTA | REPROVADO | A janela recusa a mesma chave e manda editar a linha (`ERRO\|5\|... edite a linha existente`). A aba EQA.CAP_Dados está protegida, e a célula F467 está `Locked`. A janela não tem editar nem excluir. Ver **OC-HF2-1** |
| HF2-T09 | A correção chega ao bias sem outro passo? | REPROVADO | Depois da edição (F467 = 105), O467 passa a 5 na aba de dados. Mas a EQA_Base continua com 110/10, e a Estatística e o Painel não mudam (0 células). Z3 diz "base em dia", e Z2 continua com o carimbo de 24/08/2026. Ver **OC-HF2-2** |
| HF2-T10 | Correção depois do "Consolidar na EQA_Base" | APROVADO | Em 21,8 s, a EQA_Base passou a 105/5. Bias 2,4002 %; ET 4,478 / 3,976 / 3,919 %; Sigma 1,640 / 2,006 / 1,982. Tudo igual ao independente (máx. 3,4e-16). Só as linhas do RBC mudaram. A lista da rodada mostra 105 |
| HF2-T11 | Calculadora independente (84 linhas × 27 colunas + 45 células do Painel) em S1 e S2 | APROVADO | Run A S1 e S2: **0 divergências** (2.268 + 45 comparações cada; n_eqa = 466). Run B S1: 0 divergências. Run B S2: 0 na Estatística e 1 no Painel, em M47 (DPM do LYMPH# N3, Sigma 9,75): QC 1,1e-10 contra 0. É o cancelamento de 1 − Φ com Sigma > 6,5 já justificado na F3 §4.1 (tolerância absoluta de 1e-9 DPM). A mesma grandeza na Estatística (V) passou. Não é ocorrência |
| HF2-T12 | Isolamento de provedor: Controllab do mesmo analito, com R4 = CAP | APROVADO | `OK\|2\|` na tabela do Controllab. A Estatística ficou idêntica (0 células), e K5 também. Mas a linha entrou com Analito_Canonico **vazio** (OC-HF2-3) |
| HF2-T13 | Isolamento de ano: CAP 2026 com Ano EP = 2025 | APROVADO | Estatística idêntica (0 células). Com N4 = 2026: só o RBC tem bias (G = 30, a amostra de 2026); os outros 27 analitos ficam SEM EP, como manda o GUIA §5. K5 "ano 2026 \| 1 rodada(s), 1 linha(s)". Com N4 = TODOS: 5 rodadas, G = 7,9202 %, igual ao independente; o WBC (controle) também é igual ao independente (1,8648; 3 rodadas). N4 restaurado: idêntico |
| HF2-T14 | Nome do analito fora do mapa: "Red Blood Cell" no lugar de "Red Blood Cell Count" | REPROVADO | `OK\|469\|` sem aviso. Entra na EQA_Base com Canônico vazio; "sem analito canonico" passa de 16 para 17. Bias inalterado (0 células). Ver **OC-HF2-3** |
| HF2-T15 | **Nome trocado para o do vizinho (run B):** "Lymphs" no lugar de "Lymphs Absolute", com unidade x 10E9/L | REPROVADO | `OK\|469\|` sem aviso. A linha é mapeada para **LYMPH%** e entra no bias dele: G 3,6294 → **15,2220 %**; Sigma N1 4,33 → 1,70 ("Bom" → "Desempenho inadequado"); R "FORA (\|SDI\| max 5,00)"; S "NAO OK (1 fora dos limites)"; AC 15 → 16. O LYMPH# não muda. Ver **OC-HF2-4** |
| HF2-T16 | Registro sem valor-alvo | REPROVADO | `OK\|470\|` sem aviso. O bias da linha fica vazio, e AC continua 16. A coluna S passou de "16 dentro dos limites" para "17". A amostra conta nos limites e não conta no bias. Ver **OC-HF2-5** |
| HF2-T17 | O bias do CEQ não depende do período nem do lote do CIQ | APROVADO no CEQ; REPROVADO na Estatística depois da troca de lote | Ano De/Até = 2026: G, T, AC, AD, R e S iguais em todas as linhas; mudaram C:F, H, N, O, P, Q e U (o lote 522611 não tem CIQ em 2026). H3 = 619611 no Painel: G/T/AC/AD iguais e Painel B7:E9 do lote novo. **Mas a Estatística mostra B4 = 619611 com C:F e H do lote 522611** (0 células mudaram). Ver **OC-HF2-6**. Restaurados: idênticos |
| HF2-T18 | Provedor global trocado (R4 = Controllab) | APROVADO no isolamento; REPROVADO no ET do Painel | K5 "Controllab \| ... 0 rodada(s), 0 linha(s)". Os **28 analitos** ficam SEM EP, mesmo com o registro Controllab do RBC (o mapa não tem nenhum nome do Controllab: OC-HF2-3). Na Estatística, H e L ficam vazios. **O Painel H7:H9 mostra ET 2,078 / 1,576 / 1,519 % (= 1,65 × CV) com bias "SEM EP"** (run B: 7,205 / 6,990 / 5,638 %). Confirma no Excel a OC-H2 da F3. Ver **OC-HF2-7**. R4 restaurado: idêntico |
| HF2-T19 | Trilha de auditoria do CEQ | REPROVADO | Cada gravação pela janela gera CEQ_REGISTRADO (AUD-000030 a 034). A edição da linha, as consolidações e a troca de R4, que muda o provedor do bias de todos os analitos, não geram nenhuma linha. Ver **OC-HF2-9** |
| HF2-T20 | Permissão: gravar sem sessão | REPROVADO | Depois do `mSeguranca.Logout` (SessaoAtiva = Falso, usuário vazio), `SalvarCEQ` devolveu `OK\|471\|` e gravou. Ver **OC-HF2-10** |
| HF2-T21 | Reversão (exclusão como o usuário: limpar as células digitadas e consolidar) | APROVADO | Run A e run B: 6 linhas limpas (5 CAP, 1 Controllab). Depois do Consolidar, a Estatística A14:AU97, K5, o Painel e as 465 linhas da EQA_Base ficaram **bit a bit iguais ao B0**. A tabela do CAP fica com 470 linhas (5 vazias no fim), o que é inofensivo |
| HF2-T22 | Run B (LYMPH#, ETp VB): inclusão, correção, só o analito certo | APROVADO | Seção 3b. Inclusão: G 3,2705 → 4,9529 %; T 0,3339 → 2,7504; ET N1 10,476 → 12,158 %; Sigma N1 3,906 → 3,569. Correção: G 3,7029; Sigma N1 3,820. Tudo igual ao independente (máx. 1,9e-16). O LYMPH% (vizinho) ficou idêntico na inclusão e na correção: 0 células de outros analitos mudaram, nem a U |
| HF2-T23 | Dados da produção: o provedor da coluna A da tabela do CAP | REPROVADO (dados) | A `tblEQA_CAP_Dados` tem 375 linhas "CAP" e **90 linhas "Controllab"** (a simulação, rodadas 1/2/3, Uso = NAO). A consolidação grava as 90 como **CAP** na EQA_Base. Ver **OC-HF2-11** |
| HF2-T24 | O manual (GUIA §5–6) ensina o fluxo do CEQ que a tela tem? | REPROVADO | O §6 manda digitar nas abas EQA.CAP_Dados / EQA.Controllab_Dados. Não cita a janela, nem o "Consolidar", nem a correção, e para o ANALISTA as abas estão bloqueadas. Mesma OC-9 da Bio. Ver **OC-HF2-12** |
| HF2-T25 | Janela `frmCEQ` aberta na tela (foco, marca vermelha, "Salvar e próximo") | NAO_EXECUTADO | A janela é modal e travaria a automação, e não posso capturar a tela. O comportamento foi lido no código (unidade automática ao escolher o analito, contexto mantido) |
| HF2-T26 | Segurança do teste | APROVADO | Produção `0b45fa270d2a` antes e depois. Cópias A e B com hash igual antes e depois. Só os PIDs 28268 e 19392 foram criados e encerrados, um de cada vez. Nenhuma captura de tela. Cópias e retratos apagados |

## 3. Relatório B: o efeito do CEQ no bias, no ET e no Sigma (RBC, lote 522611, XN1000)

**Dados de entrada brutos:** EQA_Base com Analito_Canonico = RBC, Uso ≠ NAO, Provedor = Estatística!R4 (CAP) e
Ano = N4 (2025).

**Definição conferida no código** (`mCEQ.BiasEQ`/`Casa`/`AnoVigente`; ADR-035/071/072; fórmulas G, H, L da
Estatística):

- bias_i = (Resultado − Alvo)/Alvo × 100, recalculado de Resultado e Alvo, sem ler a coluna Bias.
- Etapa 1: média dos |bias| por rodada (ano\|rodada). Etapa 2: média das médias.
- H = 1,65 × F + |G|. L = (K − |G|)/CVp.
- O CVp não depende do CEQ: (K − G)/L na linha de base dá N1 0,9757600160893026, N2 0,7976139072742076 e
  N3 0,8072347361497517, iguais ao Painel C45:C47 (`IncertezaCIQ "CV"`, validado de forma independente na F3).
- ETp = 4 (CLIA). F (CV do lote em 2025): N1 1,2595422808251646, N2 0,9550870524007767, N3 0,9204085695584352.

**Médias por rodada (|bias|, %):** FH9-A 2025 = 0,8057023199 (5); FH9-B 2025 = 1,4818564618 (5);
FH9-C 2025 = 2,3133015941 (5); TESTE-AUDITORIA R1 = 10 (S1) e depois 5 (S2) (1); TESTE-AUDITORIA R2 (2026) = 30
(só com N4 = 2026 ou TODOS). Todas as amostras do RBC estão acima do alvo, por isso T = G.

| Analito | Equip. | Lote | Nível | Etapa | Dados de entrada | Fórmula | Independente | QC_INI (Value2) | Dif. rel. | Status |
|---|---|---|---|---|---|---|---|---|---|---|
| RBC | XN1000 | 522611 | 1–3 | B0 (base) | 15 amostras / 3 rodadas | G | 1,533620125276759 | 1,533620125276759 | 0 | APROVADO |
| RBC | XN1000 | 522611 | 1–3 | S1 inclusão (+10 %) | 16 / 4 | G = (0,80570+1,48186+2,31330+10)/4 | 3,6502150939575695 | 3,650215093957569 | 1,2e-16 | APROVADO |
| RBC | XN1000 | 522611 | 1–3 | S1 | idem | T | 3,6502150939575695 | 3,650215093957569 | 1,2e-16 | APROVADO |
| RBC | XN1000 | 522611 | 1 / 2 / 3 | S1 | F | H = 1,65·F + G | 5,728459857319091 / 5,226108730418851 / 5,168889233728987 | 5,72845985731909 / 5,226108730418851 / 5,168889233728987 | ≤ 1,6e-16 | APROVADO |
| RBC | XN1000 | 522611 | 1 / 2 / 3 | S1 | CVp | L = (4 − G)/CVp | 0,35847431773676797 / 0,4385391263271689 / 0,43331250549349654 | 0,3584743177367684 / 0,43853912632716946 / 0,4333125054934971 | ≤ 1,3e-15 | APROVADO |
| RBC | XN1000 | 522611 | 1–3 | S2 correção (+5 %) | 16 / 4 | G | 2,4002150939575695 | 2,400215093957569 | 1,9e-16 | APROVADO |
| RBC | XN1000 | 522611 | 1 / 2 / 3 | S2 | idem | H | 4,478459857319091 / 3,9761087304188507 / 3,9188892337289873 | 4,47845985731909 / idem / idem | ≤ 2e-16 | APROVADO |
| RBC | XN1000 | 522611 | 1 / 2 / 3 | S2 | idem | L | 1,6395270144949416 / 2,005713405260935 / 1,9818088028184793 | 1,639527014494942 / 2,0057134052609356 / 1,98180880281848 | ≤ 3,4e-16 | APROVADO |
| RBC | XN1000 | 522611 | 1–3 | N4 = 2026 | 1 amostra (R2) | G / H N1 / L N1 | 30 / 32,078244763361525 / −26,645896092570013 | idem | 0 | APROVADO |
| RBC | XN1000 | 522611 | 1–3 | N4 = TODOS | 17 amostras / 5 rodadas | G / L N1 | 7,920172075166056 / −4,017557606918049 | idem | 0 | APROVADO |
| WBC (controle) | XN1000 | 522611 | 1–3 | N4 = TODOS | 15 / 3 | G | 1,8648146011737141 | 1,8648146011737141 | 0 | APROVADO |
| Todos (84 linhas) | XN1000 | 522611 | 1–3 | S1, S2 | tblCQ_Final + EQA_Base | calc_independente (27 colunas + 45 do Painel) | | | máx. < 1e-9 | APROVADO (0 divergências) |

**Caminho do dado (confirmado):**

1. A janela (`SalvarCEQ`) grava na tabela do provedor. Como a Hema não tem linha vazia, ela faz `ListRows.Add`; O/P e
   S..X saem por fórmula.
2. `mEQA.AtualizarEQABase` mapeia o nome do provedor para o canônico (W:Y), calcula o Uso e grava os valores. Depois
   vêm `InvalidarEQ`, `MarcarCEQSujo` e `RecalcularIncerteza`.
3. A Estatística recalcula G/T/AC/AD → H, L, M..W e a incerteza (AJ..AU).
4. O Painel G7:G9, I7:I9, O4, J/M/O45:47 e C/H61:63 leem a Estatística. O Painel H7:H9 é calculado no próprio Painel.

A correção feita na tabela para no passo 1 até alguém clicar em "Consolidar" (OC-HF2-2).

## 3b. Relatório B: LYMPH# (ETp VB 22,758 %), lote 522611, XN1000

CVp derivado da linha de base (não depende do CEQ), igual ao Painel C45:C47: N1 4,988470640249724,
N2 2,863348092688884, N3 1,9541086448368994. F: N1 4,366816642059906, N2 4,236274252445276, N3 3,4168926222310985.
Médias por rodada (|bias| / com sinal): FH9-A = 2,73847 / 2,71563; FH9-B = 3,36482 / −3,02042; FH9-C = 3,70829 / 1,30646
(5 amostras cada). TESTE-AUDITORIA R1 = 10 e depois 5.

| Analito | Equip. | Lote | Nível | Etapa | Dados de entrada | Fórmula | Independente | QC_INI (Value2) | Dif. rel. | Status |
|---|---|---|---|---|---|---|---|---|---|---|
| LYMPH# | XN1000 | 522611 | 1–3 | B0 | 15 / 3 | G / T | 3,2705266790177085 / 0,3338915076796763 | 3,2705266790177085 / 0,3338915076796766 | ≤ 9e-16 | APROVADO |
| LYMPH# | XN1000 | 522611 | 1–3 | S1 inclusão | 16 / 4 | G / T | 4,9528950092632815 / 2,750418630759757 | 4,9528950092632815 / 2,7504186307597576 | ≤ 1,6e-16 | APROVADO |
| LYMPH# | XN1000 | 522611 | 1 / 2 / 3 | S1 | F, CVp | H | 12,158142468662126 / 11,942747525797987 / 10,590767835944593 | idem | 0 | APROVADO |
| LYMPH# | XN1000 | 522611 | 1 / 2 / 3 | S1 | | L | 3,5692477770989917 / 6,21827565754141 / 9,11161607665935 | 3,5692477770989917 / 6,21827565754141 / 9,111616076659352 | ≤ 1,9e-16 | APROVADO |
| LYMPH# | XN1000 | 522611 | 1–3 | S2 correção | 16 / 4 | G / T | 3,7028950092632815 / 1,5004186307597571 | 3,7028950092632815 / 1,5004186307597576 | ≤ 3e-16 | APROVADO |
| LYMPH# | XN1000 | 522611 | 1 / 2 / 3 | S2 | | H / L | 10,908142468662126 / 3,8198255774200587 (N1); L N2 6,6548275400008565; L N3 9,75129391816664 | idem | 0 | APROVADO |
| LYMPH# | XN1000 | 522611 | 1–3 | N4 = 2026 / TODOS | 1 / 17 amostras | G | 30 / 8,962316007410625 | idem | 0 | APROVADO |
| LYMPH% (vizinho) | XN1000 | 522611 | 1–3 | depois de "Lymphs" (TA-03, 150/100) | 16 / 4 | G | 15,222013284930423 | 15,222013284930423 | 0 | APROVADO no cálculo (o QC_INI faz o que o mapa manda); ver OC-HF2-4 |
| Todos (84 linhas) | XN1000 | 522611 | 1–3 | S1, S2 | | calc_independente | | | máx. < 1e-9 (V/M47 com a tolerância absoluta da F3) | APROVADO |

**Observação:** no LYMPH#, a etapa 1 com sinal mostra que a FH9-B tem viés de −3,02 % e as outras de +2,7 e +1,3 %. Com
o |bias| (3,27 %), o sinal não se cancela, como o ADR-035 quer. T = 0,33 mostra a oscilação.

## 4. Ocorrências

### OC-HF2-1: O ANALISTA não consegue corrigir nem excluir um CEQ registrado, e a janela manda "edite a linha existente" (Alto; `funcional`; reproduz a OC-1 da Bio)
- **Tela/módulo:** frmCEQ (`mCEQForm.SalvarCEQ`) e aba EQA.CAP_Dados / EQA.Controllab_Dados.
- **Ação:** registrar 110 por engano e tentar registrar 105 na mesma rodada, analito e amostra.
- **Esperado:** um caminho de correção para quem registra. O ADR-087 diz "digitar direto na tabela continua possível".
- **Observado:** `ERRO|5|Já existe este registro (...) na linha 467 da aba EQA.CAP_Dados. Para corrigir, edite a linha existente.` A aba está protegida para o ANALISTA, e todas as 24 colunas, na tabela e abaixo dela, estão `Locked`. A janela não tem editar nem excluir.
- **Dificuldade:** a mensagem manda o usuário para uma célula que o Excel recusa.
- **Consequência:** um CEQ errado fica no bias, no ET e no Sigma até um ADM corrigir. Outra saída é registrar outra amostra, o que gera um registro duplicado.
- **Sugestão:** "Corrigir/Excluir" na própria janela, a partir da lista da rodada, com a mesma validação, consolidação e trilha. Ou desbloquear A..N, Q e R e consolidar ao sair da aba. A decisão é do Orquestrador e do UX/UI.
- **Evidência:** `hema_f2\A\f2h.json` → P0.tabela_CAP.colunas (Locked), P4.janela_mesma_chave, P4.celula_bloqueada = true. O mesmo no run B.

### OC-HF2-2: A correção na tabela não chega ao bias, ao ET nem ao Sigma até o "Consolidar", e o indicador diz "base em dia" (Alto; `funcional`; reproduz a OC-2 da Bio)
- **Tela/módulo:** EQA.CAP_Dados (Z2:Z3), `mEQA.AtualizarEQABase`, Estatística, Painel.
- **Ação:** corrigir 110 → 105 na linha da tabela e olhar a Estatística e o Painel.
- **Esperado:** o bias acompanha a correção, ou a tela avisa que falta consolidar.
- **Observado:**
  - O467 passa a 5, mas a EQA_Base fica com 110/10, e a Estatística e o Painel ficam com 0 células mudadas. G continua 3,6502 quando deveria ser 2,4002, e o Sigma N1 continua 0,358 em vez de 1,640.
  - Z3 continua "base em dia", porque só compara a quantidade de linhas.
  - Z2 (`=mEQA.CarimboEQA()`) mostra "consolidado: 2026-08-24 15:20:53" mesmo depois de 7 consolidações (o carimbo real em EQA_Base!Z1 é de 31/10).
- **Consequência:** ET e Sigma calculados com um CEQ já corrigido, e nada na tela mostra a divergência.
- **Sugestão:** consolidar ao editar a tabela ou no ATUALIZAR DADOS. Z3 por conteúdo, e Z2 volátil ou gravado como valor.
- **Evidência:** `f2h.json` (A e B) → P4.O_bias_na_tabela = 5, EQA_Base_linha_apos_editar (110/10), estat_mudou_sem_consolidar = [], Z3_apos_editar, Z2_apos_consolidar.

### OC-HF2-3: Nome fora do mapa e qualquer CEQ do Controllab são gravados "OK" e nunca entram no bias, sem aviso (Alto; `funcional`; reproduz a OC-3 da Bio)
- **Tela/módulo:** frmCEQ / `SalvarCEQ` → `mEQA.Consolidar` (mapa W:Y da EQA_Base).
- **Ação:** registrar "Red Blood Cell" no CAP; registrar "RBC" no Controllab; trocar R4 para Controllab.
- **Observado:**
  - Os dois registros voltam `OK`, com Canônico vazio, e "sem analito canonico" sobe de 15 para 16 e 17.
  - O mapa tem 31 entradas, **todas do CAP**. Com R4 = Controllab, os 28 analitos da Hematologia ficam SEM EP: sem ET e sem Sigma na Estatística, mesmo com um CEQ Controllab do RBC registrado.
- **Consequência:** um laboratório que passe a usar o Controllab, ou que digite o nome como está no PDF, fica sem bias em silêncio. Na Hema, isso vale para o produto inteiro, porque R4 é global.
- **Sugestão:** no `SalvarCEQ`, conferir `mEQA.MapearAnalito` e avisar quando voltar vazio. Cadastrar o mapa do Controllab.
- **Evidência:** `f2h.json` → P5.I2_controllab, P5.I4_nome_fora_do_mapa (run A) e P5.I8_provedor_controllab.analitos_G_numerico = [].

### OC-HF2-4: Um CEQ do LYMPH# registrado como "Lymphs" entra calado no bias do LYMPH% e derruba o Sigma dele (Alto; `funcional`; nova)
- **Tela/módulo:** frmCEQ (campo Analito em texto livre) / `SalvarCEQ` / mapa W:Y.
- **Ação (run B):** registrar um resultado de "Lymphs Absolute" com o nome "Lymphs" (uma palavra a menos), com a unidade x 10E9/L. Resultado 150, alvo 100, na rodada TESTE-AUDITORIA R1.
- **Esperado:** pelo menos um aviso. O nome é de outro analito canônico (LYMPH%), e a unidade (x 10E9/L) não é a que esse analito tem nos registros (%). Ou a janela mostra o analito do QC_INI (o canônico) que vai receber o resultado.
- **Observado:**
  - `OK|469|` sem aviso. A EQA_Base grava Canônico = **LYMPH%**.
  - O LYMPH% passa de G 3,63 % para **15,22 %**, ET N1 de 10,36 % para 21,95 %, Sigma N1 de 4,33 para **1,70** ("Bom" → "Desempenho inadequado"), margem "Margem critica", R "FORA (|SDI| max 5,00 > 2)" e S "NAO OK (1 fora dos limites)".
  - O LYMPH#, o analito pretendido, não muda.
  - O cálculo está certo pela definição (independente = 15,222013284930423). O problema é a falta de proteção contra a troca.
- **Dificuldade:** o mapa do CAP tem 7 pares desse tipo na Hema (Lymphs, Basophils, Eosinophils, Monocytes, Neut/Gran e IG, cada um com o seu "... Absolute", e nRBC/100 WBC × nRBC Absolute). Na janela, a unidade só vem sozinha se o campo estiver vazio ou automático; uma unidade digitada antes não é trocada.
- **Consequência:** associação errada entre analitos no bias, no ET e no Sigma, com efeito no plano de CQ do analito vizinho. Pelo §4 das REGRAS, associação errada é Crítico. Classifiquei como Alto porque a troca é erro de digitação do usuário, e não do cálculo. A decisão final é do Orquestrador.
- **Sugestão:** mostrar, ao lado do Analito, o nome canônico que vai receber o resultado (`MapearAnalito`). Avisar quando a unidade digitada diferir da unidade já registrada para aquele nome. Opcional: lista fechada com os nomes do mapa.
- **Evidência:** `hema_f2\B\f2h.json` → P5.I4_nome_fora_do_mapa (nome usado "Lymphs", 80 células mudadas, todas nas linhas LYMPH%) e P5.I6_N4_TODOS.controle_independente (G = 15,2220, 16 amostras, 4 rodadas).

### OC-HF2-5: Registro sem valor-alvo é aceito sem aviso; a amostra conta nos limites e some do bias (Médio; `funcional`; reproduz a OC-4 da Bio)
- **Observado:** `OK|470|`. O bias da linha fica vazio, e AC continua 16. A coluna S passou de "OK (16 dentro dos limites)" para "OK (17 dentro dos limites)". As duas colunas da mesma linha falam de números de amostras diferentes.
- **Consequência:** o usuário acha que registrou a rodada inteira, mas o bias é calculado com menos amostras.
- **Sugestão:** Valor-alvo obrigatório, ou um aviso ao gravar sem ele.
- **Evidência:** `f2h.json` → P5.I5_sem_alvo e P5.I5_linha_analito (A e B).

### OC-HF2-6: Depois de trocar o lote no Painel (H3), a Estatística mostra o lote novo em B4 com n, média, DP, CV e ET do lote anterior (Alto; `dados`; nova, fora do escopo do CEQ: encaminhar à F4)
- **Tela/módulo:** Painel H3 → `mLotes.TrocarLoteAnalise` → Estatística (B4, C:F, H).
- **Ação:** trocar o lote em análise de 522611 para 619611 no Painel e olhar a Estatística.
- **Esperado:** a Estatística acompanha o lote de B4 (o texto da linha 4 diz "O lote é o escolhido no Painel (H3)"), ou avisa que precisa de ATUALIZAR DADOS.
- **Observado:**
  - B4 passou a "619611", mas 0 células da Estatística mudaram. RBC N1 continua n 50, média 2,4022 e CV 1,2595 % (do lote 522611), e o ET 4,478 %.
  - O Painel, no mesmo momento, mostra o lote 619611: n 54, média 2,3148, CV 1,1344 % e ET 4,272 %.
  - Pelo código, o `TrocarLoteAnalise` chama `RecalcularAnalitoAtual` (Calc, Westgard, Painel) e não `AtualizarEstatisticaAba`. A aba Estatística da Hema não tem evento.
  - O bias do CEQ, o Sigma (CVp de 6 meses) e o restauro ficaram certos.
- **Consequência:** tabela × Painel divergentes. Quem lê a Estatística (ou o plano de CQ a partir dela) vê os números de um lote com o rótulo de outro.
- **Sugestão:** chamar `AtualizarEstatisticaAba` na troca de lote, ou mostrar "Estatística desatualizada — clique em ATUALIZAR DADOS" quando `engEstat` for de outro lote. A confirmação e a extensão (Bio, H3 vazio) ficam com a F4.
- **Evidência:** `hema_f2\A\f2h.json` → P5.I7b_lote (colunas_que_mudaram = [], B4 = 619611, linha_analito_depois × painel.B7:J9). O mesmo no run B. `src/hema/mLotes.bas` linhas 398–444.

### OC-HF2-7: Com o bias em "SEM EP", o Painel mostra ET = 1,65 × CV, enquanto a Estatística deixa o ET vazio (Alto; `calculo`; confirma no Excel a OC-H2 da F3 e a OC-7 da Bio)
- **Tela/módulo:** Painel H7:H9 (`=$E7*1.65+ABS(IF(ISNUMBER($G7),$G7,0))`).
- **Ação:** R4 = Controllab, um provedor sem CEQ utilizável.
- **Observado:**
  - Run A: o Painel mostra bias "SEM EP" e ET 2,078 / 1,576 / 1,519 % (= 1,65 × 1,2595 / 0,9551 / 0,9204), contra o ETp de 4 %.
  - Run B: 7,205 / 6,990 / 5,638 %.
  - Na Estatística, H e L ficam vazios. O bias ausente é tratado como 0 no Painel.
- **Consequência:** o ET parece pequeno e esconde a falta do CEQ. No caso do RBC, com o CAP, o ET real era de 3,6 a 5,7 %.
- **Sugestão:** sem bias, não mostrar o ET (como na Estatística), ou mostrar "ET só de imprecisão".
- **Evidência:** `f2h.json` (A e B) → P5.I8_provedor_controllab.painel.B7:J9.

### OC-HF2-8: As listas da janela misturam os nomes da simulação com os do CAP real (Baixo; `usabilidade`; reproduz a OC-8 da Bio)
- **Observado:**
  - Rodadas: "3", "2" e "1" (a simulação) ao lado de FH9-A/B/C 2025.
  - Analitos: "RBC", "WBC", "HCT", "HGB", "MCH" e "MCV" (os nomes da simulação, mapeados) ao lado de "Red Blood Cell Count", "Hematocrit" etc.
  - Avaliação: "OK" e "See Note [26]".
  - O "RBC" não traz unidade.
- **Sugestão:** montar as listas só com as linhas de Uso = SIM.
- **Evidência:** `f2h.json` → P1.

### OC-HF2-9: Correção, consolidação e troca do provedor global (R4) não deixam rastro no Audit_Log (Médio; `funcional`; reproduz a OC-5 da Bio)
- **Observado:** só a inclusão gera CEQ_REGISTRADO. A edição 110 → 105, as consolidações (P4, P7), a limpeza das linhas e a troca de R4 não geram nenhuma linha. Na Hema, R4 muda a fonte do bias, do ET e do Sigma dos 28 analitos de uma vez.
- **Consequência:** não há como responder qual era o CEQ antes da correção, nem quem trocou o provedor.
- **Sugestão:** CEQ_ALTERADO / CEQ_EXCLUIDO (com o valor anterior e o novo) e PROVEDOR_EQ_ALTERADO.
- **Evidência:** `f2h.json` → P4.trilha_da_edicao = [], P4.trilha_consolidar = [], P5.I8_provedor_controllab.trilha = [].

### OC-HF2-10: O registro do CEQ não confere sessão nem papel (Médio; `funcional`; reproduz a OC-6 da Bio)
- **Observado:** depois do Logout (SessaoAtiva = Falso, usuário vazio), `SalvarCEQ` devolveu `OK|471|` e gravou. `AbrirCEQ` e `SalvarCEQ` são públicas e não chamam `SessaoAtiva`.
- **Sugestão:** exigir sessão (e um papel mínimo configurável) no `AbrirCEQ` e no `SalvarCEQ`.
- **Evidência:** `f2h.json` → P6 (A e B).

### OC-HF2-11: Na tabela do CAP há 90 linhas com provedor "Controllab" que a consolidação grava como CAP (Médio; `dados`; nova, só na Hema)
- **Tela/módulo:** `tblEQA_CAP_Dados` (coluna A) → `mEQA.Consolidar` (o provedor vem da aba, não da coluna A).
- **Observado:**
  - A tabela do CAP tem 375 linhas "CAP" e 90 linhas **"Controllab"**: a simulação pré-ADR-034, rodadas 1/2/3, arquivo "EQC_Dados (simulação ...)".
  - Na EQA_Base, as 90 estão como **CAP**, com Uso = NAO. A tabela do Controllab tem só uma linha modelo vazia.
  - Hoje não entram em conta (Uso = NAO). Mas o cabeçalho do `mEQA` diz "quem discordar troca a célula para SIM". Se alguém fizer isso, um Controllab entra no bias do **CAP**, a mistura que o ADR-071 proíbe.
  - A janela também lista "Controllab" como valor da coluna Provedor da tabela do CAP.
- **Consequência:** dado de provedor contraditório (coluna A × EQA_Base) e risco latente de misturar provedores.
- **Sugestão:** mover essas 90 linhas para a tabela do Controllab, ou corrigir a coluna A para "CAP (simulação)". Na consolidação, avisar quando a coluna A divergir da aba. Decisão do Orquestrador e do Especialista.
- **Evidência:** `f2h.json` → P0.tabela_CAP.provedor_na_coluna_A = {CAP: 375, Controllab: 90}; P0.EQA_Base.provedores = ['CAP']; P1.provedor_col_A_CAP.

### OC-HF2-12: O manual manda digitar o CEQ nas abas de dados e não cita a janela, o "Consolidar" nem a correção (Médio; `usabilidade`; reproduz a OC-9 da Bio)
- **Observado:** o GUIA §6 diz "digite cada rodada nas abas EQA.CAP_Dados / EQA.Controllab_Dados". Para o ANALISTA, essas abas estão bloqueadas (OC-HF2-1). Não há menção ao botão "Registrar resultado do CEQ", ao "Consolidar na EQA_Base", ao mapa de nomes (OC-HF2-3/4) nem ao alvo obrigatório para o bias (OC-HF2-5). O §5 está certo para a Hema ("o programa é escolhido em R4"), mas não avisa que R4 vale para todos os analitos ao mesmo tempo.
- **Sugestão:** atualizar o §5–6 depois das decisões sobre a OC-HF2-1/2.
- **Evidência:** `GUIA_DE_USO.md` linhas 199–232; `f2h.json` → P0.telas.

## 5. Dúvidas de regra de negócio (ao Especialista em CQ, pelo Orquestrador)

1. **Provedor global na Hema (R4):** o laboratório pode um dia ter analitos no CAP e outros no Controllab? Hoje R4 vale
   para os 28 de uma vez (a Bio escolhe por analito). Trocar R4 sem o mapa do Controllab zera o bias de todos
   (HF2-T18).
2. **Sigma negativo exibido:** com bias > ETp (N4 = 2026, um bias de 30 % contra o ETp de 4 %), L = −26,6 e o status é
   "Desempenho inadequado". O valor negativo deve aparecer, ou o certo é "< 0" / "ETp excedido"?
3. **Amostra sem alvo:** ela deve contar na coluna S (limites) quando não conta no bias (AC)? (OC-HF2-5)
4. **Amostra "NAO AVALIADO" / "See Note":** entra no bias do ET/Sigma (o `BiasEQ` não filtra avaliação; o `ViesEQ`
   filtra). Mesma dúvida da Bio e da F3 (OC-H6).
5. **Papel mínimo:** quem pode registrar e quem pode corrigir um CEQ? Hoje não há conferência (OC-HF2-10), e só o ADM
   corrige (OC-HF2-1).

## 6. Impressões de uso (Relatório A)

- **A sequência faz sentido?** Sim. A janela segue a ordem do relatório do CAP (rodada → analito/amostra → resultado
  → avaliação/fonte). As validações apontam o campo certo, com mensagens claras. A duplicidade é achada mesmo com caixa
  e espaços diferentes. A rodada existente traz o ano, e o analito traz a unidade.
- **Dá para achar a rodada e ver o que acabou de registrar?** Sim. A lista "Registros desta rodada" e o EQA.CAP_Resumo
  mostram o registro logo depois de gravar. Não há campo de lote do material do CEQ (o modelo não tem).
- **O CEQ entra no bias certo, e só nele?** Pela janela, com o nome certo, sim, nos dois runs:
  - só o analito mapeado muda, nos 3 níveis, e o vizinho (WBC, LYMPH%) não muda;
  - provedor e ano são respeitados;
  - período e lote do CIQ não mexem no bias;
  - o Painel acompanha;
  - os números batem com o cálculo independente até 1e-15.

  O risco está nas bordas, todas silenciosas:
  - nome do vizinho (OC-HF2-4, o mais grave desta frente: um "Absolute" esquecido mudou o Sigma de outro analito de
    4,3 para 1,7);
  - nome fora do mapa e Controllab (OC-HF2-3);
  - registro sem alvo (OC-HF2-5);
  - correção que o ANALISTA não consegue fazer e que, feita pelo ADM, só vale depois de um botão, com o indicador
    dizendo "base em dia" (OC-HF2-1/2).
- **Tempo:** cada gravação leva de 18 a 24 s (consolidação e incerteza dos 28 analitos), e o Consolidar leva de 18 a
  22 s. Sem indicador de progresso, parece travado.
- **Achado fora do CEQ:** a troca de lote no Painel deixa a Estatística com o lote novo no rótulo e os números do
  anterior (OC-HF2-6). Encaminhar à F4.

## 7. Limpeza

Cada run removeu as suas linhas TESTE-AUDITORIA pelo caminho do usuário (limpar as células e consolidar) e voltou bit a
bit ao B0: Estatística A14:AU97, K5, Painel e as 465 linhas da EQA_Base (HF2-T21). As cópias `Hematologia_F2_A.xlsm` e
`Hematologia_F2_B.xlsm` e os quatro retratos foram apagados. Ficam em
`%TEMP%\qcw\auditoria\2026-10-31_ciclo01\hema_f2\{A,B}\` só JSON, TXT e log, sem senha nem caminho de rede. Do Excel,
só os PIDs 28268 e 19392, um de cada vez, criados e encerrados pelo script. A produção continua com o sha
`0b45fa270d2a`.
