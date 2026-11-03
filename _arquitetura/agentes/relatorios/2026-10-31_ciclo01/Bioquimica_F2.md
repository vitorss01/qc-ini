# Bioquímica, F2 (CEQ): ciclo 2026-10-31_ciclo01

Agente: Usuário Simulado. Modo: completo. Data: 31/10/2026.

Este relatório é o do **2º passe** da frente. O 1º passe foi das 16h44 às 17h05 (Excel das 16h52 às 16h56) e o 2º, das
21h04 às 21h25, sobre a mesma produção (sha `fe10dba1f1fd`). No 2º passe:

- **Run A (reprodução):** refiz a frente inteira (P0 a P7) com o Glicose numa cópia nova da produção. Todos os passos
  deram o mesmo resultado do 1º passe, bit a bit, descontados só os tempos, os IDs da trilha e os carimbos de data.
  As ocorrências OC-1 a OC-8 foram **reproduzidas**.
- **Run B (analito novo):** refiz a frente inteira com a **Bilirrubina direta** ("Bilirubin, direct" no CAP), que tem
  ETp de **outra fonte (VB, desempenho MIN, 39,915 %)**, bias negativo grande (−30,8 %) e um vizinho de nome parecido
  (Bilirrubina total) como controle. Assim, a inclusão e a correção foram conferidas também fora do CLIA, e o
  isolamento foi conferido contra o analito mais fácil de confundir.
- Li o GUIA_DE_USO §5–6 contra o que a tela faz (OC-9, nova).

As seções 2 a 4 trazem os testes e as ocorrências do 1º passe, que continuam valendo (reproduzidos no Run A). O que é
novo do 2º passe está marcado com **[2º passe]** e na seção 3b.

**[3º passe, 01/11/2026, 02h08–02h30; Excel das 02h11m31 às 02h15m51]** Mesma produção (`fe10dba1f1fd`). Run C com a
**Proteína total** ("Protein, total, serum"), escolhida por risco: é o analito do Painel e tem **2 lotes** (8974 e
8977) no CV pooled do Sigma. A Albumina foi o controle. O roteiro `f2_ceq.py` ganhou dois passos opcionais, descritos
no cabeçalho:

- **I9:** Painel H3 → outro lote (8974) → H3 vazio (ADR-086) → 8977. O CEQ pode depender do lote?
- **I10:** primeiro CEQ de um analito que não tem CEQ nem entrada no mapa (Hemoglobina glicada), registrado pelo nome
  do CAP e pelo nome canônico.

Resultado do 3º passe:

- A inclusão e a correção mudaram o bias, o ET e o Sigma só da Proteína total, com diferença **0** em relação ao
  cálculo independente. A calculadora independente deu 0 divergências em 1.178 comparações em S1 e em S2.
- O bias do CEQ, e com ele o Sigma, não depende do lote nem do H3 vazio.
- OC-1 a OC-7 foram reproduzidas. A OC-3 ficou maior: um analito sem CEQ não consegue ganhar o seu primeiro CEQ pela
  janela.
- Uma ocorrência nova, a OC-10.

Seções novas: **[3º passe]** e 3c.

## 1. Ambiente e evidência

| Item | Valor |
|---|---|
| Produção copiada | `QC_Bioquimica.xlsm`, sha256 curto `fe10dba1f1fd` (o mesmo da F1 e da F3). O hash é igual antes e depois: a produção não foi aberta |
| Cópia de trabalho | `%TEMP%\qcw\auditoria\2026-10-31_ciclo01\Bioquimica_F2.xlsm`. Hash igual antes e depois (`fe10dba1f1fd`): ela nunca foi salva. Excel PID **6320**, criado e encerrado pelo script no `finally`. Nenhum outro Excel foi tocado (o `/automation -Embedding` da tarefa agendada ficou intacto) |
| Retratos para a calculadora | `f2_S1.xlsm` (depois da inclusão) e `f2_S2.xlsm` (depois da correção), gravados por `SaveCopyAs` com os eventos desligados. A cópia aberta não foi alterada por isso. Os três `.xlsm` foram apagados no fim |
| Ferramenta nova (reutilizável) | `_arquitetura/agentes/ferramentas/f2_ceq.py` (passos P0 a P7 documentados no cabeçalho). Saídas em `%TEMP%\qcw\auditoria\2026-10-31_ciclo01\f2\`: `f2.json`, `calc_S1.json`, `calc_S2.json`, `f2_calc_B0.json` e o log `f2.log` |
| Calculadora independente | `ferramentas/calc_independente.py` (da F3), rodada sem Excel sobre a linha de base, S1 e S2 |
| Usuário | Um usuário de teste com papel **ANALISTA**, criado só na cópia, com login pelo `mSeguranca.DoLogin`. A senha de teste vem da suíte e não aparece aqui |
| Dados de teste | Rodadas `TESTE-AUDITORIA R1` (2025), `TESTE-AUDITORIA R2` (2026) e `TESTE-AUDITORIA CTL1` (Controllab); amostras `TA-01`, `TA-R2`, `TA-03`, `TA-04`, `TA-C1` e `TA-SEMLOGIN`; arquivo-fonte `TESTE-AUDITORIA.pdf`. Valores redondos sintéticos: resultado 110 → 105, 120, 130, 150 e 100, alvo 100, SD 2, SDI 5, limites 90/110, nº lab 50 |
| **[2º passe]** Cópias | `%TEMP%\qcw\auditoria\2026-10-31_ciclo01\f2r2\A\Bioquimica_F2_A.xlsm` (Glicose; Excel PID **18460**) e `...\f2r2\B\Bioquimica_F2_B.xlsm` (Bilirrubina direta; PID **29336**), cópias novas da produção (`fe10dba1f1fd`). Hash de cada cópia igual antes e depois (nunca salvas). Cada PID foi criado e encerrado pelo próprio script, no `finally`; os Excel `/automation -Embedding` da tarefa agendada (PIDs 15412 e 11980) não foram tocados. Produção com o mesmo hash no fim |
| **[2º passe]** Comandos | `python -u f2_ceq.py <copia_A> <pasta_A>` (padrão Glicose / "Glucose, serum" / controle Ureia) e `python -u f2_ceq.py <copia_B> <pasta_B> "Bilirrubina direta" "Bilirubin, direct" "Bilirrubina total"`. Depois, `calc_independente.py` sobre os retratos `f2_S1.xlsm` e `f2_S2.xlsm` de cada run. Saídas (só JSON, TXT e log) em `...\f2r2\A\` e `...\f2r2\B\`. Os quatro retratos e as duas cópias foram apagados no fim |
| **[3º passe]** Cópia | `%TEMP%\qcw\auditoria\2026-10-31_ciclo01\f2r3\C\Bioquimica_F2_C.xlsm`, uma cópia nova da produção (`fe10dba1f1fd`), com hash igual antes e depois (nunca salva). Excel PID **28540**, criado e encerrado pelo script no `finally`. Antes do run não havia Excel aberto. O Excel PID 15844 que apareceu depois, às 02h18m03, é da tarefa agendada e não foi tocado. A produção continua com `fe10dba1f1fd` no fim |
| **[3º passe]** Comando | `set F2_LOTE_ALT=8974` e `set F2_SEM_MAPA=Hemoglobin A1c\|Hemoglobina glicada`; depois `python -u f2_ceq.py <copia_C> <pasta_C> "Proteína total" "Protein, total, serum" "Albumina" "Protein, total" "g/dL"`. Depois, `calc_independente.py` sobre `f2_S1.xlsm` e `f2_S2.xlsm`. Saídas (JSON, TXT e log) em `...\f2r3\C\`. Cópia e retratos apagados |
| **[3º passe]** Dados de teste | As mesmas rodadas e amostras TESTE-AUDITORIA dos passes anteriores, mais `TA-H1` e `TA-H2` (HbA1c, unidade %). Unidade g/dL. Valores sintéticos: 110 → 105, 120, 130, 150 e 100, alvo 100 |
| Filtros da cópia | **[3º passe]** Painel: C3 = Proteína total (o valor da produção), H3 = 8977, L4 = DIMENSION 2, M4 = CAP, período "Out" (01–31/10/2026). Estatística: N4 = 2025, P4 = TODAS. Passes anteriores: Painel C3 = Glicose (trocado de Proteína total, como o usuário faria), H3 = 8977, L4 = DIMENSION 2, M4 = CAP, 01–31/10/2026. Estatística: Ano EP N4 = 2025, Rodada P4 = TODAS, Sigma em 6 meses (29/04–28/10/2026) |

Como os passos foram executados:

- **Gravação:** pelo núcleo da janela, `mCEQForm.SalvarCEQ`. É exatamente o que o botão "Salvar e próximo" do `frmCEQ` chama, com os 16 campos na ordem da janela. A janela em si **não** foi aberta: ela é modal e travaria a automação.
- **Leitura da janela:** o comportamento dela (marcar o campo, manter o contexto, lista da rodada) foi lido no código (`frmCEQ.frm.txt`), e as funções que ela usa foram chamadas: `DistintosCEQ`, `ValorRegistrado` e `RegistrosDaRodada`.
- **Correção:** escrevi direto na célula da tabela `tblEQA_CAP_Dados`, com os eventos ligados, como a mensagem da janela manda. Depois rodei o "Consolidar na EQA_Base" (OnAction `AtualizarEQABase`, botão em J1 da aba de dados).
- **Filtros:** Painel C3/L4/M4 e Estatística N4 foram mudados pela própria célula, com os eventos ligados. Isso exigiu tirar a proteção dessas duas abas, um passo técnico.
- **Limite do método:** uma escrita por COM numa aba protegida com `UserInterfaceOnly` passa mesmo em célula bloqueada. Por isso, o bloqueio que o usuário encontraria foi conferido pelas propriedades `Locked` e `ProtectContents`, e não pelo teclado.

## 2. Testes executados

| ID | Teste | Estado | Evidência |
|---|---|---|---|
| F2-T01 | Telas do CEQ: onde se registra, botões, proteção | APROVADO (com ressalvas: OC-1 e OC-8) | Botão "Registrar resultado do CEQ" (`mCEQForm.AbrirCEQ`) em A3 de EQA.CAP_Resumo e de EQA.Controllab_Resumo. Botão "Consolidar na EQA_Base" (`AtualizarEQABase`) em J1 de EQA.CAP_Dados e de EQA.Controllab_Dados. A EQA_Base não tem botão. O Início leva ao EQA pelo bloco "Controle externo (EQA)". `frmCEQ` instalado. Nenhuma das abas EQA tem evento de planilha. A tabela não tem validação de dados em nenhuma coluna (A..R) |
| F2-T02 | Preenchimento automático e listas da janela | APROVADO | A rodada existente traz o ano (`C-C 2025` → 2025). A rodada inexistente não traz nada. O analito traz a unidade (`Glucose, serum` → mg/dL). Rodadas em ordem decrescente: C-C 2025, C-B 2025, C-A 2025, C, B, A. A lista da rodada C-C 2025 tem 150 registros e a de ano 2026 tem 0. Ressalva de conteúdo das listas: OC-8 |
| F2-T03 | Validações antes de gravar (12 casos) | APROVADO | Cada caso foi recusado com o campo certo: resultado vazio (6), resultado em texto (6), alvo em texto (7), limite inferior > superior (10), ano "25" (3), SD −1 (8), avaliação vazia (12), provedor "PNCQ" (1), rodada vazia (2), amostra vazia (5) e analito vazio (4). A tabela do provedor ficou idêntica: 545 linhas antes e depois |
| F2-T04 | Duplicidade de uma chave **real** (C-A 2025 · Glucose, serum · CHM-01) | APROVADO | "Já existe este registro (...) na linha 77 da aba EQA.CAP_Dados. Para corrigir, edite a linha existente." Nada gravado |
| F2-T05 | Inclusão: gravação, consolidação, trilha, lista da rodada | APROVADO | `OK\|547\|` em 16,6 s. A linha 547 tem os 16 campos, O = 10 e P = 10. A EQA_Base ganhou a linha com Analito_Canonico = Glicose, Uso = SIM, Bias 10 e chave `CAP\|2025\|TESTE-AUDITORIA R1\|Glucose, serum\|TA-01`. Carimbo "CAP 546 \| total 546". Audit_Log: CEQ_REGISTRADO com o usuário e o papel. A lista da rodada mostra `Glucose, serum · TA-01 · 110 · Acceptable`. EQA.CAP_Resumo mostra a rodada nova |
| F2-T06 | Inclusão: bias, ET e Sigma do analito mudam como deveriam | APROVADO | Seção 3, etapa B1. Bias de 2,7333 % para 4,5500 % (4 rodadas, 16 amostras). ET N1 de 4,878 % para 6,694 %; N2 de 5,120 % para 6,936 %. Sigma N1 de 0,330 para 0,247; N2 de 1,134 para 0,851. Diferença máxima em relação ao independente: 1,5e-16 |
| F2-T07 | Inclusão: **só** o analito certo muda | APROVADO | Mudaram só Glicose\|1 e Glicose\|2 (G, H, L, N, O, R, T, V, W, AC, AD, AJ–AM, AQ, AR, AT, AU). Nas outras 60 linhas da Estatística (A14:AU80), 0 células mudaram. O resumo Six Sigma (A90:F110) não mudou. Painel B7:I8, J46:J47 e M46:M47 acompanham. CVp (C46:C47) e CV/n/média não mudaram. K5: "3 rodada(s), 435 linha(s)" → "4 rodada(s), 436 linha(s)" |
| F2-T08 | Correção pelo caminho do usuário ANALISTA | REPROVADO | A janela recusa a mesma chave e manda editar a linha (`ERRO\|5\|... edite a linha existente`). Na aba EQA.CAP_Dados, protegida para o ANALISTA, todas as colunas A..R estão `Locked`, tanto na tabela quanto abaixo dela. Não há edição nem exclusão na janela. Ver **OC-1** |
| F2-T09 | A correção chega ao bias sem outro passo? | REPROVADO | Depois da edição (F547 = 105), O547 passa a 5 na aba de dados, mas a EQA_Base continua com 110 e bias 10, e a Estatística/Painel não mudam (0 células). Z3 diz "base em dia", e Z2 mostra o carimbo de 22/08/2026. Ver **OC-2** |
| F2-T10 | Correção depois do "Consolidar na EQA_Base" | APROVADO | Em 15,5 s, a EQA_Base passou a 105 e bias 5. Bias 3,3000 %, ET N1 5,444 %, N2 5,686 %, Sigma N1 0,304, N2 1,046 = independente (máx. 8,7e-16, sinal). Só as linhas do Glicose mudaram (G, H, L, N, O, T, V, W, ...). A lista da rodada mostra 105 |
| F2-T11 | Calculadora independente (62 linhas × 19 colunas) na linha de base, em S1 e em S2 | APROVADO | B0, S1 e S2: **0 divergências** em C, D, E, F, G, H, J, K, L, M, N, O, T, V, W, AC, AD, AF e AR (1.178 comparações cada). n_eqa = 545 / 546 / 546 |
| F2-T12 | Isolamento de provedor: Controllab do mesmo analito, com o provedor do analito = CAP | APROVADO | `OK\|2\|` na tabela do Controllab. A Estatística ficou idêntica (0 células) e K5 também. Não mistura. Mas a linha entrou com Analito_Canonico **vazio** (ver OC-3) |
| F2-T13 | Isolamento de ano: CAP 2026 com o Ano EP = 2025 | APROVADO | Estatística idêntica (0 células). Com N4 = 2026: G = 30 (só a amostra de 2026), K5 "ano 2026 \| 1 rodada(s), 1 linha(s)". Os outros 30 analitos ficaram em SEM EP, como manda o GUIA §5. Com N4 = TODOS: 5 rodadas, G = 8,6400 %, igual ao independente; Ureia (controle) igual ao independente (3 rodadas). N4 restaurado: idêntico |
| F2-T14 | Nome do analito fora do mapa ("Glucose" em vez de "Glucose, serum") | REPROVADO | `OK\|549\|` sem aviso. Entra na EQA_Base com Canônico vazio, e "sem analito canonico" passa de 21 para 22. Bias inalterado (0 células). Ver **OC-3** |
| F2-T15 | Registro sem valor-alvo | REPROVADO | `OK\|550\|` sem aviso. Bias vazio na tabela e na EQA_Base. A amostra não entra no bias (AC continua 16). Ver **OC-4** |
| F2-T16 | Equipamento: o CEQ entra no bias dos dois DIMENSION? | APROVADO (segue o ADR-087) | Com L4 = DIMENSION 1, G, T, AC e AD ficaram iguais em todas as linhas. Mudaram C:F, H, L e a incerteza (CIQ de outro equipamento). O modelo do CEQ não tem equipamento (ADR-087). Dúvida 1 |
| F2-T17 | Provedor do analito trocado no Painel (M4 = Controllab) | APROVADO no isolamento; REPROVADO no ET do Painel | Especificações AR do Glicose = Controllab. As linhas do Glicose ficaram em SEM EP (G, T, AC, AD, R, S), H e L vazios na Estatística, e nenhum outro analito mudou. M4 voltou à fórmula. Restaurado: idêntico. **Mas o Painel H7:H8 mostrou ET 2,144 / 2,386 % (= 1,65 × CV), com bias "SEM EP".** Ver **OC-7**. Sem trilha: OC-5 |
| F2-T18 | Trilha de auditoria do CEQ | REPROVADO | A inclusão gera CEQ_REGISTRADO. A edição da linha, as 7 consolidações e a troca de provedor no M4 não geram nenhuma linha no Audit_Log (420 linhas em S1 e em S2). Ver **OC-5** |
| F2-T19 | Permissão: gravar sem sessão | REPROVADO | Depois do `mSeguranca.Logout` (SessaoAtiva = Falso), `SalvarCEQ` devolveu `OK\|551\|` e gravou. Ver **OC-6** |
| F2-T20 | Reversão (exclusão como o usuário: limpar as células digitadas e consolidar) | APROVADO | 6 linhas limpas (5 CAP e 1 Controllab). Depois do Consolidar, a Estatística A14:AU80 ficou **bit a bit igual ao B0**. K5, o Painel e as 545 linhas da EQA_Base também ficaram iguais |
| F2-T21 | Segurança do teste | APROVADO | Produção e cópia com hash `fe10dba1f1fd` antes e depois. Só o PID 6320 foi criado e encerrado. Nenhuma captura de tela |
| F2-T22 | Janela `frmCEQ` aberta na tela (foco, marca vermelha, "Salvar e próximo") | NAO_EXECUTADO | A janela é modal e travaria a automação, e não posso capturar a tela. Comportamento lido no código. A suíte `qa_adr087` V01 cobre a instalação e o layout (não reexecutada aqui) |
| F2-T23 **[2º passe]** | Reprodução da frente inteira (Glicose, Run A) numa cópia nova | APROVADO | `f2r2\A\f2.json` comparado com `f2\f2.json`: P0, P1, P2, P3, P4, P6 e P7 idênticos; em P5, retornos, mudanças, K5 e linhas da EQA_Base idênticos (só o tempo de cada gravação e os IDs AUD- da trilha mudam). Inclusão `OK\|547\|`; bias 2,7333 → 4,5500 → 3,3000 %; ET N1 4,878 → 6,694 → 5,444 %; Sigma N1 0,330 → 0,247 → 0,304. 0 células de outros analitos mudaram. Reversão bit a bit igual ao B0 |
| F2-T24 **[2º passe]** | Inclusão de CEQ na Bilirrubina direta (ETp VB 39,915 %): bias, ET e Sigma | APROVADO | `OK\|547\|` em 18,7 s. EQA_Base: Canônico = Bilirrubina direta, Uso = SIM, bias 10. G de 30,7709 → 25,5782 % (4 rodadas, 16 amostras); T de −30,7709 → −20,5782; ET N1 37,114 → 31,921 %, N2 33,651 → 28,459 %; Sigma N1 0,4348 → 0,6818, N2 1,2384 → 1,9417. Diferença máxima em relação ao independente: 2,3e-16. Seção 3b |
| F2-T25 **[2º passe]** | Inclusão na Bilirrubina direta: **só** o analito certo muda (vizinho Bilirrubina total intacto) | APROVADO (com observação) | Mudaram as linhas Bilirrubina direta\|1 e \|2 (G, H, L, N, O, P, R, S, T, U, V, W, AC, AD, AJ–AM, AR, AT, AU). A Bilirrubina total ficou idêntica. Nas outras linhas, só a coluna **U** (ordem da "Margem crítica") mudou em 10 células, de 1 em 1 (Ureia\|1 3 → 2, AST\|1 5 → 4, Sódio, Potássio, Cloro, Bicarbonato), e o resumo A90:F110 mudou. Motivo: a Bilirrubina direta N1 saiu da "Margem crítica" (O de 7,0 % → 20,0 %) e liberou a 2ª posição. É o efeito esperado de uma ordenação, não contaminação: G, H, L e as demais estatísticas desses analitos não mudaram |
| F2-T26 **[2º passe]** | Correção (110 → 105) na Bilirrubina direta, antes e depois do Consolidar | APROVADO no cálculo; REPROVADO no fluxo (OC-1, OC-2 reproduzidas) | A janela recusa a mesma chave ("edite a linha existente"); a célula está `Locked` numa aba protegida. Editada a linha: O = 5 na tabela, mas a EQA_Base fica com 110/10 e 0 células mudam na Estatística; Z2 continua com o carimbo de 22/08. Depois do Consolidar (17,1 s): G 24,3282 %, T −21,8282, ET N1 30,671 / N2 27,209 %, Sigma N1 0,7412 / N2 2,1109 = independente (máx. 2,1e-16). Só as linhas da Bilirrubina direta mudaram |
| F2-T27 **[2º passe]** | Isolamento da Bilirrubina direta: Controllab, ano 2026, Ano EP 2026/TODOS, equipamento, provedor no M4 | APROVADO no isolamento; REPROVADO no ET do Painel (OC-7) | Controllab: 0 células, Canônico vazio (OC-3). CAP 2026 com N4 = 2025: 0 células. N4 = 2026: G = 30, ET N1 36,343, Sigma N1 0,4715 = independente. N4 = TODOS: G 25,46256 (17 amostras, 5 rodadas) = independente; Bilirrubina total (controle) = independente (10,8223; 3 rodadas). L4 = DIMENSION 1: G/T/AC/AD iguais em todas as linhas. M4 = Controllab: só a Bilirrubina direta fica em SEM EP; **o Painel mostra ET 6,343 % (N1) e 2,880 % (N2) = 1,65 × CV com bias "SEM EP"** (OC-7 reproduzida em outro analito). Tudo restaurado idêntico |
| F2-T28 **[2º passe]** | Registro sem alvo na Bilirrubina direta | REPROVADO (OC-4 reproduzida) | `OK\|550\|` sem aviso; bias vazio; AC continua 16. Detalhe novo: a coluna S passou de "16 dentro dos limites" para "17": a amostra sem alvo **conta** na conferência dos limites e **não conta** no bias. As duas colunas falam de um número de amostras diferente |
| F2-T29 **[2º passe]** | Calculadora independente (62 linhas × 19 colunas) em S1 e S2 dos dois runs | APROVADO | A-S1, A-S2, B-S1 e B-S2: **0 divergências** em C, D, E, F, G, H, J, K, L, M, N, O, T, V, W, AC, AD, AF e AR (1.178 comparações cada; n_eqa = 546) |
| F2-T30 **[2º passe]** | Permissão e trilha (Run A e Run B) | REPROVADO (OC-5, OC-6 reproduzidas) | Nos dois runs: `SalvarCEQ` sem sessão devolve `OK\|551\|`; edição, consolidação e troca de provedor no M4 não geram linha no Audit_Log; só a inclusão gera CEQ_REGISTRADO |
| F2-T31 **[2º passe]** | O manual (GUIA §5–6) ensina o fluxo do CEQ que a tela tem? | REPROVADO | O GUIA §6 manda "digite cada rodada nas abas EQA.CAP_Dados / EQA.Controllab_Dados". Não cita a janela "Registrar resultado do CEQ" (ADR-087), nem o botão "Consolidar na EQA_Base", nem como corrigir. Para o ANALISTA, digitar nessas abas não é possível (OC-1). Ver **OC-9** |
| F2-T32 **[2º passe]** | Segurança do teste | APROVADO | Produção `fe10dba1f1fd` antes e depois. Cópias A e B com hash igual antes e depois. PIDs 18460 e 29336 criados e encerrados pelo script. Nenhuma captura de tela. Cópias e retratos apagados |
| F2-T33 **[3º passe]** | Telas, listas e validações com a Proteína total (P0 a P2) | APROVADO (OC-1 e OC-8 reproduzidas) | Login ANALISTA. A unidade vem do analito (`Protein, total, serum` → g/dL). As rodadas são C-C, C-B e C-A 2025, mais A, B e C, a simulação (OC-8). Os 12 casos de validação foram recusados com o campo certo. Duplicidade de uma chave real: "Já existe este registro (...) na linha 122". A tabela ficou idêntica (545 → 545) |
| F2-T34 **[3º passe]** | Inclusão de CEQ na Proteína total (2 lotes; analito do Painel): bias, ET e Sigma, e só desse analito | APROVADO | `OK\|547\|` em 17,1 s. EQA_Base: Canônico = Proteína total, Uso = SIM, bias 10. G de 3,3793 → 5,0345 % (4 rodadas, 16 amostras). T de 2,2801 → 4,2101. ET N1 de 6,4486 → 8,1037 %; N2 de 7,0114 → 8,6666 %. Sigma N1 de 1,9907 → 1,4931; N2 de 1,1176 → 0,8382. Diferença em relação ao independente: **0** nas 12 comparações. Mudaram só as linhas Proteína total\|1 e \|2. Nas outras 60 linhas: **0 células**. O resumo Six Sigma não mudou. Painel B7:I8 e J46:J47 acompanham. CVp (C46:C47) inalterado. Trilha CEQ_REGISTRADO (1 linha) |
| F2-T35 **[3º passe]** | Correção 110 → 105 na Proteína total | APROVADO no cálculo; REPROVADO no fluxo (OC-1 e OC-2 reproduzidas) | A janela recusa a mesma chave ("edite a linha existente"). A célula está `Locked` numa aba protegida. Editada a linha, O = 5 na tabela, mas a EQA_Base fica com 110/10 e a Estatística muda em 0 células. Depois do "Consolidar" (17,4 s): G 3,7845 %, T 2,9601, ET N1 6,8537 / N2 7,4166 %, Sigma N1 1,8689 / N2 1,0492, todos iguais ao independente (dif. 0). Só as linhas da Proteína total mudaram. Detalhe novo da OC-2: o carimbo da EQA_Base (Z1) passou a "2026-11-01 02:11:45", enquanto o Z2 da aba de dados continua "2026-08-22 18:48:37" depois da consolidação |
| F2-T36 **[3º passe]** | Isolamento da Proteína total: Controllab, ano 2026, N4 = 2026 e TODOS, equipamento, provedor no M4 | APROVADO no isolamento; REPROVADO no ET do Painel (OC-7) e no Controllab (OC-3) | Controllab: `OK\|2\|`, 0 células, Canônico vazio. CAP 2026 com N4 = 2025: 0 células. N4 = 2026: G = 30, ET N1 33,0692, Sigma N1 −6,0137 / N2 −3,3761, iguais ao independente (Dúvida 6). N4 = TODOS: G 9,0276 (17 amostras, 5 rodadas) = independente, e só a Proteína total mudou o G. A Albumina, o controle, ficou igual ao independente (5,0395; 3 rodadas). L4 = DIMENSION 1: G/T/AC/AD iguais em todas as linhas. M4 = Controllab: só a Proteína total fica em SEM EP, mas o **Painel mostra ET 3,069 / 3,632 % = 1,65 × CV com bias "SEM EP"** (OC-7). Os três filtros foram restaurados e a Estatística voltou idêntica |
| F2-T37 **[3º passe, novo]** | O CEQ depende do lote? Painel H3 8977 → 8974 → vazio (ADR-086) → 8977 | APROVADO no CEQ; observações OC-F4-7 e OC-10 | Com H3 = 8974 e com H3 vazio, **G, T, AC e AD ficaram iguais em todas as 62 linhas** (0 mudanças), e o **Sigma (L) também** (o CV pooled junta os lotes). O ET acompanha o CV do lote: H = 1,65·F + G com diferença 0 (Proteína total N1: F 94,98 → H 160,50). Mudaram só C, D, E, F, H, N, O, P, Q e U. Observações: (1) com o 8974, que não tem ponto em out/2026, o período virou "Tudo" (B5/D5 = 0), e na volta ao 8977 o "Out" não voltou, como na OC-F4-7. Por isso, a Estatística restaurada difere da de antes só nessas colunas do CIQ. (2) Com H3 vazio, a Estatística H4 (`=loteAnalise`) mostra "8974 (refletido do Painel)" e calcula média, CV e ET do 8974, enquanto o Painel está em TODOS com média, CV e ET vazios (OC-10). O F de 95 % da Proteína total no período "Tudo" vem dos valores impossíveis já registrados (OC-3 da F3, OC-1 da F1) |
| F2-T38 **[3º passe, novo]** | Primeiro CEQ de um analito sem CEQ e sem entrada no mapa (Hemoglobina glicada) | REPROVADO (amplia a OC-3) | Pelo nome do CAP ("Hemoglobin A1c"): `OK\|551\|`. Pelo nome canônico ("Hemoglobina glicada"): `OK\|552\|`. Nos dois casos, Canônico vazio na EQA_Base, e o "sem analito canonico" foi de 22 → 23 → 24. `mEQA.MapearAnalito("CAP", ...)` = "" para os dois nomes. A HbA1c continua "SEM EP" (G, T, AC, AD), sem ET e sem Sigma, e nada avisa. Nem o nome canônico entra |
| F2-T39 **[3º passe]** | Registro sem alvo na Proteína total | REPROVADO (OC-4 reproduzida) | `OK\|550\|` sem aviso, bias vazio, AC continua 16. A coluna S passou de "16 dentro dos limites" para "17" |
| F2-T40 **[3º passe]** | Permissão e trilha | REPROVADO (OC-5 e OC-6 reproduzidas) | Depois do Logout (SessaoAtiva = Falso), `SalvarCEQ` devolveu `OK\|553\|`. A edição da linha, a consolidação e a troca de provedor no M4 não geram linha no Audit_Log |
| F2-T41 **[3º passe]** | Calculadora independente (62 linhas × 19 colunas) em S1 e S2 do Run C | APROVADO | S1 e S2: **0 divergências** em C, D, E, F, G, H, J, K, L, M, N, O, T, V, W, AC, AD, AF e AR (1.178 comparações cada; n_eqa = 546; lote 8977, DIMENSION 2, 01–31/10/2026) |
| F2-T42 **[3º passe]** | Reversão (limpar as 8 linhas TESTE-AUDITORIA e consolidar) | APROVADO no CEQ, com ressalva | 7 linhas CAP e 1 Controllab limpas. EQA_Base: 545 linhas **bit a bit iguais ao B0**. K5 igual ao B0. A Estatística não voltou ao B0 só pelas colunas do CIQ do período (C, D, E, F, H, N, O, P, Q, U), porque o "Out" se perdeu no I9 (OC-F4-7). O bias depende só da EQA_Base e dos filtros do CEQ, que voltaram iguais. A cópia foi descartada |
| F2-T43 **[3º passe]** | Segurança do teste | APROVADO | Produção `fe10dba1f1fd` antes e depois. Cópia com hash igual antes e depois. Só o PID 28540 foi criado e encerrado. Nenhuma captura de tela. Cópia e retratos apagados. O JSON não tem senha nem caminho de rede |

## 3. Relatório B: o efeito do CEQ no bias, no ET e no Sigma (Glicose, lote 8977, DIMENSION 2)

**Dados de entrada brutos:** EQA_Base, linhas com Analito_Canonico = Glicose, Uso ≠ NAO, Provedor = CAP (Especificações AR) e Ano = N4.

**Definição conferida no código** (`mCEQ.BiasEQ`; ADR-035/071/072):

- bias_i = (Resultado − Alvo)/Alvo × 100. Recalculado aqui a partir de Resultado e Alvo, sem ler a coluna Bias.
- Etapa 1: média dos |bias| de cada rodada (chave ano\|rodada).
- Etapa 2: média das médias.
- ET (H) = 1,65 × CV do lote no período (F) + |bias|.
- Sigma (L) = (ETp − |bias|)/CVp.
- O CVp não depende do CEQ: (K − G)/L na linha de base dá N1 22,0340724139129 e N2 6,406991775820687, os mesmos valores do Painel C46:C47 (o CVp foi validado de forma independente na F3).
- ETp = 10 (CLIA). F: N1 1,299696056299363 e N2 1,4462182390269869.

**Médias por rodada (|bias|, %):**

| Rodada | Média | n |
|---|---|---|
| C-A 2025 | 1,4909704188 | 5 |
| C-B 2025 | 1,4362161107 | 5 |
| C-C 2025 | 5,2726968003 | 5 |
| TESTE-AUDITORIA R1 | 10 (B1) e depois 5 (B2) | 1 |
| TESTE-AUDITORIA R2 (2026) | 30 (só com N4 = 2026 ou TODOS) | 1 |

| Analito | Equip. | Lote | Nível | Etapa | Dados de entrada | Fórmula | Independente | QC_INI (Value2) | Dif. rel. | Status |
|---|---|---|---|---|---|---|---|---|---|---|
| Glicose | DIM 2 | 8977 | 1 e 2 | B0 (base) | 15 amostras, 3 rodadas | G | 2,7332944432981674 | 2,733294443298167 | 1,6e-16 | APROVADO |
| Glicose | DIM 2 | 8977 | 1 e 2 | B1 inclusão (+10 %) | 16 / 4 | G = (1,49097+1,43622+5,27270+10)/4 | 4,549970832473625 | 4,549970832473625 | 0 | APROVADO |
| Glicose | DIM 2 | 8977 | 1 e 2 | B1 | idem | T (com sinal) | 1,5051721711534922 | 1,5051721711534924 | 1,5e-16 | APROVADO |
| Glicose | DIM 2 | 8977 | 1 | B1 | F = 1,2997 | H = 1,65·F + G | 6,694469325367574 | 6,694469325367574 | 0 | APROVADO |
| Glicose | DIM 2 | 8977 | 2 | B1 | F = 1,4462 | H | 6,936230926868154 | 6,936230926868154 | 0 | APROVADO |
| Glicose | DIM 2 | 8977 | 1 | B1 | CVp 22,0341 | L = (10 − G)/CVp | 0,24734552311288047 | 0,24734552311288047 | 0 | APROVADO |
| Glicose | DIM 2 | 8977 | 2 | B1 | CVp 6,4070 | L | 0,850637765463382 | 0,850637765463382 | 0 | APROVADO |
| Glicose | DIM 2 | 8977 | 1 e 2 | B2 correção (+5 %) | 16 / 4 | G | 3,2999708324736257 | 3,299970832473625 | 1,3e-16 | APROVADO |
| Glicose | DIM 2 | 8977 | 1 e 2 | B2 | idem | T | 0,2551721711534922 | 0,25517217115349244 | 8,7e-16 | APROVADO |
| Glicose | DIM 2 | 8977 | 1 | B2 | idem | H | 5,444469325367574 | 5,444469325367574 | 0 | APROVADO |
| Glicose | DIM 2 | 8977 | 2 | B2 | idem | H | 5,686230926868154 | 5,686230926868154 | 0 | APROVADO |
| Glicose | DIM 2 | 8977 | 1 | B2 | idem | L | 0,3040758440684709 | 0,3040758440684709 | 0 | APROVADO |
| Glicose | DIM 2 | 8977 | 2 | B2 | idem | L | 1,0457371262456712 | 1,045737126245671 | ~1e-16 | APROVADO |
| Glicose | DIM 2 | 8977 | 1 e 2 | N4 = 2026 | 1 amostra (R2) | G | 30 | 30 | 0 | APROVADO |
| Glicose | DIM 2 | 8977 | 1 / 2 | N4 = 2026 | | H / L | 32,14449849289395 / −3,1215897725166273 (N2) | idem | 0 | APROVADO |
| Glicose | DIM 2 | 8977 | 1 e 2 | N4 = TODOS | 17 amostras, 5 rodadas | G | 8,639976665978901 | 8,639976665978901 | 0 | APROVADO |
| Glicose | DIM 2 | 8977 | 1 / 2 | N4 = TODOS | | L | 0,06172364819688728 / 0,21227174649321132 | idem | 0 | APROVADO |
| Todos (62 linhas) | DIM 2 | 8977 | 1 e 2 | B0, S1, S2 | tblCQ_Final + EQA_Base | calc_independente (19 colunas) | | | máx. < 1e-9 | APROVADO (0/1.178 divergências em cada) |

**Caminho do dado** (confirmado nesta frente):

- **Inclusão:** janela (`SalvarCEQ`) → tabela do provedor (O/P por fórmula) → `mEQA.AtualizarEQABase`. Este passo mapeia o nome do provedor para o canônico (W:Y), calcula o Uso e copia os valores para a EQA_Base. Depois vêm `mCEQ.InvalidarEQ` e `MarcarCEQSujo`, e `RecalcularIncerteza` recalcula a Estatística G/T/AC/AD → H, L, N, O, V, W e a incerteza. O Painel B7:I8 e J46:M47 leem a Estatística.
- **Correção na tabela:** o caminho para na tabela até alguém clicar em "Consolidar na EQA_Base" (OC-2).

## 3b. [2º passe] Relatório B: Bilirrubina direta (ETp VB), lote 8977, DIMENSION 2

**Entradas:** a mesma definição da seção 3 (`mCEQ.BiasEQ`, duas etapas), com ETp = 39,91509557613611 % (VB,
desempenho MIN, recalculado na F3 a partir da biblioteca 2025). CVp (derivado da linha de base, não depende do CEQ):
N1 21,02846873831051 e N2 7,383850631271355. CV do lote no período (F): N1 3,8443004348429066 e N2 1,745726394510911.

**Médias por rodada (|bias|, %):** C-A 2025 = 44,857091772879365 (5); C-B 2025 = 21,63348166491173 (5);
C-C 2025 = 25,82222630756503 (5); TESTE-AUDITORIA R1 = 10 (B1) e depois 5 (B2); TESTE-AUDITORIA R2 (2026) = 30.

| Analito | Equip. | Lote | Nível | Etapa | Dados de entrada | Fórmula | Independente | QC_INI (Value2) | Dif. rel. | Status |
|---|---|---|---|---|---|---|---|---|---|---|
| Bilirrubina direta | DIM 2 | 8977 | 1 e 2 | B0 (base) | 15 amostras, 3 rodadas | G | 30,770933248452042 | 30,77093324845204 | 1,2e-16 | APROVADO |
| Bilirrubina direta | DIM 2 | 8977 | 1 / 2 | B0 | F, CVp | H / L | 37,114028965943085 / 1,238400231033604 (N2) | 37,114028965943085 / 1,2384002310336044 | ≤ 3,8e-16 | APROVADO |
| Bilirrubina direta | DIM 2 | 8977 | 1 e 2 | B1 inclusão (+10 %) | 16 / 4 | G = (44,8571+21,6335+25,8222+10)/4 | 25,578199936339033 | 25,57819993633903 | 1,4e-16 | APROVADO |
| Bilirrubina direta | DIM 2 | 8977 | 1 e 2 | B1 | idem | T (com sinal) | −20,578199936339033 | −20,57819993633903 | 1,7e-16 | APROVADO |
| Bilirrubina direta | DIM 2 | 8977 | 1 | B1 | F = 3,8443 | H = 1,65·F + G | 31,921295653830075 | 31,92129565383007 | 1,1e-16 | APROVADO |
| Bilirrubina direta | DIM 2 | 8977 | 2 | B1 | F = 1,7457 | H | 28,458648487281312 | 28,45864848728131 | 1,2e-16 | APROVADO |
| Bilirrubina direta | DIM 2 | 8977 | 1 | B1 | CVp 21,0285 | L = (39,9151 − G)/CVp | 0,6817850514087859 | 0,681785051408786 | 1,6e-16 | APROVADO |
| Bilirrubina direta | DIM 2 | 8977 | 2 | B1 | CVp 7,3839 | L | 1,9416556964301084 | 1,9416556964301088 | 2,3e-16 | APROVADO |
| Bilirrubina direta | DIM 2 | 8977 | 1 e 2 | B2 correção (+5 %) | 16 / 4 | G | 24,328199936339033 | 24,32819993633903 | 1,5e-16 | APROVADO |
| Bilirrubina direta | DIM 2 | 8977 | 1 e 2 | B2 | idem | T | −21,828199936339033 | −21,82819993633903 | 1,6e-16 | APROVADO |
| Bilirrubina direta | DIM 2 | 8977 | 1 / 2 | B2 | idem | H | 30,671295653830075 / 27,208648487281312 | 30,67129565383007 / 27,20864848728131 | ≤ 1,3e-16 | APROVADO |
| Bilirrubina direta | DIM 2 | 8977 | 1 / 2 | B2 | idem | L | 0,7412282764745606 / 2,110944061325536 | 0,7412282764745607 / 2,1109440613255366 | ≤ 2,1e-16 | APROVADO |
| Bilirrubina direta | DIM 2 | 8977 | 1 / 2 | N4 = 2026 | 1 amostra (R2) | G / H / L | 30 / 36,343095717491046 (N1) / 1,3428082542927773 (N2) | idem | 0 | APROVADO |
| Bilirrubina direta | DIM 2 | 8977 | 1 / 2 | N4 = TODOS | 17 amostras, 5 rodadas | G / L | 25,462559949071228 / 0,6872842624405966 (N1) | 25,462559949071224 / 0,6872842624405968 | ≤ 3e-16 | APROVADO |
| Bilirrubina total (controle) | DIM 2 | 8977 | 1 e 2 | N4 = TODOS | 15 amostras, 3 rodadas | G | 10,82231673715532 | 10,82231673715532 | 0 | APROVADO |
| Todos (62 linhas) | DIM 2 | 8977 | 1 e 2 | S1, S2 (Run A e Run B) | tblCQ_Final + EQA_Base | calc_independente (19 colunas) | | | máx. < 1e-9 | APROVADO (0/1.178 em cada um dos 4 retratos) |

**Observações do Run B:**

- **O sentido da mudança é o esperado pela definição.** A Bilirrubina direta tem bias negativo grande (−30,8 %). Uma
  rodada nova com +10 % **reduz** o |bias| médio (30,77 → 25,58 %), porque a etapa 2 é a média das médias de |bias|
  de cada rodada. ET cai, Sigma sobe, e o analito sai da "Margem crítica" no N1. Está de acordo com o ADR-035/071; não
  é ocorrência, mas mostra que **uma única rodada com sinal oposto melhora o Sigma** de um analito com erro
  sistemático persistente. O viés com sinal (T) deixa isso visível (−30,8 → −20,6).
- **O Painel e a Estatística diferem na 15ª casa** no ET (Painel H7 31,921295653829823 × Estatística H 31,92129565383007,
  dif. rel. 7,7e-15). O Painel calcula o CV pelo seu próprio caminho. A diferença está muito abaixo da tolerância de
  1e-9 e da casa mostrada na tela. Não é ocorrência.
- **Cascata na coluna U:** ver F2-T25. É uma ordenação; nenhum número de outro analito mudou.

## 3c. [3º passe] Relatório B: Proteína total (2 lotes, ETp CLIA 10 %), lote 8977, DIMENSION 2

**Entradas:** a mesma definição da seção 3 (`mCEQ.BiasEQ`, em duas etapas). ETp = 10 (CLIA, Especificações). CVp vem
da linha de base e não depende do CEQ: (K − G)/L dá N1 3,3257310225463055 e N2 5,923988964744061, os mesmos valores
do Painel C46:C47. O CV do lote no período (F) é N1 1,8601475203401814 e N2 2,201278360034401 (n = 14 e 15).

**Médias por rodada (|bias|, %):**

| Rodada | Média | n |
|---|---|---|
| C-A 2025 | 4,6150323451197774 | 5 |
| C-B 2025 | 2,3550331099484563 | 5 |
| C-C 2025 | 3,167943625034068 | 5 |
| TESTE-AUDITORIA R1 | 10 (B1), depois 5 (B2) | 1 |
| TESTE-AUDITORIA R2 (2026) | 30 | 1 |

| Analito | Equip. | Lote | Nível | Etapa | Dados de entrada | Fórmula | Independente | QC_INI (Value2) | Dif. rel. | Status |
|---|---|---|---|---|---|---|---|---|---|---|
| Proteína total | DIM 2 | 8977 | 1 e 2 | B0 (base) | 15 amostras, 3 rodadas | G | 3,3793363600341007 | 3,3793363600341007 | 0 | APROVADO |
| Proteína total | DIM 2 | 8977 | 1 / 2 | B0 | F, CVp | H / L | 6,448579768595399 / 1,1176022911872419 (N2) | idem | 0 | APROVADO |
| Proteína total | DIM 2 | 8977 | 1 e 2 | B1 inclusão (+10 %) | 16 / 4 | G = (4,61503+2,35503+3,16794+10)/4 | 5,034502270025575 | 5,034502270025575 | 0 | APROVADO |
| Proteína total | DIM 2 | 8977 | 1 e 2 | B1 | idem | T (com sinal) | 4,210059101601026 | 4,210059101601026 | 0 | APROVADO |
| Proteína total | DIM 2 | 8977 | 1 / 2 | B1 | F | H = 1,65·F + G | 8,103745678586874 / 8,666611564082336 | idem | 0 | APROVADO |
| Proteína total | DIM 2 | 8977 | 1 / 2 | B1 | CVp | L = (10 − G)/CVp | 1,4930545183334316 / 0,8382017183904314 | idem | 0 | APROVADO |
| Proteína total | DIM 2 | 8977 | 1 e 2 | B2 correção (+5 %) | 16 / 4 | G / T | 3,7845022700255755 / 2,960059101601025 | idem | 0 | APROVADO |
| Proteína total | DIM 2 | 8977 | 1 / 2 | B2 | idem | H | 6,853745678586875 / 7,416611564082336 | idem | 0 | APROVADO |
| Proteína total | DIM 2 | 8977 | 1 / 2 | B2 | idem | L | 1,868911733341443 / 1,0492081884293243 | idem | 0 | APROVADO |
| Proteína total | DIM 2 | 8977 | 1 / 2 | N4 = 2026 | 1 amostra (R2) | G / H / L | 30 / 33,0692434085613 (N1) / −6,0137154401281805 (N1), −3,3761035206222867 (N2) | idem | 0 | APROVADO (Dúvida 6) |
| Proteína total | DIM 2 | 8977 | 1 / 2 | N4 = TODOS | 17 amostras, 5 rodadas | G / L | 9,02760181602046 / 0,2923862986475179 (N1), 0,16414584661900203 (N2) | idem | 0 | APROVADO |
| Albumina (controle) | DIM 2 | 8977 | 1 e 2 | N4 = TODOS | 15 amostras, 3 rodadas | G / T | 5,039521617989425 / −1,0916112161005735 | idem | 0 | APROVADO |
| Proteína total | DIM 2 | **8974** e **H3 vazio** | 1 e 2 | B2 + I9 | idem | G, T, AC, AD, L inalterados; H = 1,65·F + G | N1: F 94,97759811973062 → H 160,49753916758107 | 160,49753916758107 | 0 | APROVADO (CEQ independe do lote) |
| Todos (62 linhas) | DIM 2 | 8977 | 1 e 2 | S1, S2 | tblCQ_Final + EQA_Base | calc_independente (19 colunas) | | | máx. < 1e-9 | APROVADO (0/1.178 em cada) |

**Observações:**

- **Painel × Estatística:** CV N1 1,860147520339105 × 1,8601475203401814 e ET N1 6,448579768593623 × 6,448579768595399.
  A diferença relativa é de até 5,7e-13. É maior que a da Glicose (1e-15), porque o Painel calcula o CV pelo seu
  próprio caminho, mas fica muito abaixo da tolerância de 1e-9 e da casa mostrada. Não é ocorrência.
- **Caminho do dado:** é o mesmo da seção 3. O passo I9 confirmou que o lote do Painel entra só no CV do período (F)
  e, por ele, no ET (H). O bias (G) e o Sigma (L, com o CV pooled de todos os lotes com n ≥ 20) não mudam com o lote.

## 4. Ocorrências

**[3º passe]** OC-1 a OC-7 foram reproduzidas no Run C (Proteína total), e OC-8 também, pelas listas. A OC-3 foi
ampliada (F2-T38). A OC-10 é nova. A OC-F4-7 da F4 também foi reproduzida, no passo I9.

**[2º passe]** OC-1 a OC-8 foram reproduzidas no Run A (Glicose) com os mesmos números. OC-1, OC-2, OC-3, OC-4, OC-5,
OC-6 e OC-7 foram reproduzidas também no Run B (Bilirrubina direta). A OC-9 é nova.

### OC-1: O usuário não-ADM não consegue corrigir nem excluir um CEQ registrado, e a janela manda "edite a linha existente" (Alto; tipo `funcional`)
- **Tela/módulo:** frmCEQ (`mCEQForm.SalvarCEQ`) e aba EQA.CAP_Dados / EQA.Controllab_Dados.
- **Ação:** um ANALISTA registra 110 por engano e tenta registrar 105 na mesma rodada, analito e amostra.
- **Esperado:** um caminho de correção para quem registra. O ADR-087 diz que "digitar direto na tabela continua possível".
- **Observado:** a janela recusa: `ERRO|5|Já existe este registro (...) na linha 547 da aba EQA.CAP_Dados. Para corrigir, edite a linha existente.`. Mas, para o ANALISTA, a aba está protegida (`ProtectContents` = Verdadeiro, pelo `ReprotectAll` do login) e todas as colunas A..R, da tabela e abaixo dela, estão `Locked` = Verdadeiro. A janela não tem editar nem excluir. Só o ADM (`UnprotectAll`) consegue mexer na linha. Pelo mesmo motivo, digitar uma rodada direto na tabela também não é possível para quem não é ADM.
- **Dificuldade:** o usuário é mandado para um lugar onde o Excel diz que a célula está protegida.
- **Consequência:** um CEQ errado fica no bias, no ET e no Sigma até um ADM corrigir. Ou o usuário contorna o problema registrando outra amostra/rodada, o que gera um registro duplicado com chave diferente e distorce o bias.
- **Sugestão:** decisão do Orquestrador e do UX/UI:
  1. "Corrigir/Excluir" na própria janela, a partir da lista da rodada, passando pela mesma validação, consolidação e trilha;
  2. ou desbloquear A..N, Q e R da tabela e consolidar ao sair da aba.

  Ajustar a mensagem conforme a decisão.
- **Evidência:** `f2.json` → P0.tabela_cap.colunas (Locked de A..R), P4.janela_mesma_chave, P4.protegida/celula_bloqueada. A escrita por COM passou só por causa do `UserInterfaceOnly`.

### OC-2: A correção feita na tabela não chega ao bias, ao ET nem ao Sigma até o "Consolidar na EQA_Base", e o indicador de estado diz "base em dia" (Alto; tipo `funcional`)
- **Tela/módulo:** EQA.CAP_Dados (Z1:Z3), `mEQA.AtualizarEQABase`, Estatística/Painel.
- **Ação:** corrigir o resultado na linha da tabela (110 → 105) e olhar a Estatística e o Painel.
- **Esperado:** o bias acompanha a correção, ou a tela avisa claramente que falta consolidar.
- **Observado:**
  - A aba de dados mostra o bias novo (O547 = 5), mas a EQA_Base continua com 110 e bias 10.
  - A Estatística e o Painel não mudam (0 células): G continua 4,5500, quando deveria ser 3,3000. ET N1 6,694 em vez de 5,444. Sigma N1 0,247 em vez de 0,304.
  - Z3 continua "base em dia", porque só compara a **quantidade** de linhas (`COUNTA` × `COUNTIF`). Uma edição de valor nunca o deixa "desatualizado".
  - Z2 (`=mEQA.CarimboEQA()`, uma função sem dependência) continua mostrando "consolidado: 2026-08-22 18:48:37" mesmo depois das 7 consolidações do teste.
  - O ATUALIZAR DADOS não consolida o CEQ: `AtualizarEQABase` só é chamado pelo `SalvarCEQ` e pelo botão (leitura do `src/`).
- **Dificuldade:** a aba de dados e o resumo mostram o valor novo, e a Estatística, o antigo. Nada indica a divergência.
- **Consequência:** ET e Sigma calculados com um CEQ que já foi corrigido. Risco de decisão errada no plano de CQ.
- **Sugestão:** consolidar automaticamente ao editar a tabela (evento da aba) ou ao clicar em ATUALIZAR DADOS. Fazer o Z3 comparar o conteúdo (por exemplo, um hash ou a data da última edição contra a do carimbo) e o Z2 recalcular (Volatile, ou o carimbo gravado como valor).
- **Evidência:** `f2.json` → P4 (O_bias_na_tabela = 5; EQA_Base_linha_apos_editar com 110/10; estat_mudou_sem_consolidar = []; estado_Z2_apos_editar e estado_Z2_apos_consolidar = carimbo de 22/08); Z3 = "base em dia" em S1 e S2.

### OC-3: Um CEQ com nome de analito fora do mapa, ou de qualquer analito do Controllab, é gravado "OK" e nunca entra no bias, sem aviso (Alto; tipo `funcional`)
- **Tela/módulo:** frmCEQ / `SalvarCEQ` → `mEQA.Consolidar` (mapa W:Y da EQA_Base).
- **Ação:** registrar "Glucose" (em vez de "Glucose, serum") no CAP. Registrar "Glicose" no Controllab.
- **Esperado:** aviso de que o analito não está mapeado e não vai alimentar o bias, com a orientação de cadastrar no mapa ou escolher o nome da lista.
- **Observado:**
  - Os dois casos devolvem `OK|...|` sem aviso, e a janela mostraria "✓ Gravado".
  - Na EQA_Base, Analito_Canonico fica vazio, e "sem analito canonico" sobe de 20 para 21 e 22. O bias não muda.
  - O mapa W:Y tem **37 entradas, todas do CAP, e nenhuma do Controllab**. Com o provedor do Glicose trocado para Controllab no Painel (M4), o Glicose ficou em SEM EP, sem ET e sem Sigma, mesmo havendo um CEQ Controllab registrado.
  - O nome do analito na janela é texto livre (combo sem lista fixa).
- **Dificuldade:** o registro aparece na lista da rodada como se estivesse tudo certo.
- **Consequência:** um laboratório que passe a usar o Controllab, ou que digite o nome do analito como está no PDF, fica sem bias, ET e Sigma em silêncio.
- **Sugestão:** no `SalvarCEQ`, conferir `mEQA.MapearAnalito(prov, an)` e avisar (ou pedir confirmação) quando voltar vazio. Incluir as entradas do Controllab no mapa. Decisão do Orquestrador.
- **Evidência:** `f2.json` → P5.I2_controllab.linha_EQA_Base (canônico null), P5.I4_nome_fora_do_mapa (OK\|549\|, carimbo "sem analito canonico 22", n_mudancas 0), P5.I8_provedor_controllab.
- **[3º passe]:** reproduzida na Proteína total. O nome "Protein, total", sem ", serum", deu `OK|549|` com Canônico
  vazio. O Controllab "Proteína total" também ficou com Canônico vazio. **Ampliação (F2-T38):** um analito que ainda
  não tem CEQ (Hemoglobina glicada) não consegue ganhar o primeiro, nem pelo nome do CAP ("Hemoglobin A1c") nem pelo
  próprio nome canônico. Os dois casos devolvem `OK`, e `MapearAnalito` devolve "" para os dois. A HbA1c continua
  "SEM EP", e nada avisa. A janela não tem como cadastrar o nome no mapa W:Y da EQA_Base. Hoje, só quem edita a
  EQA_Base (o ADM) consegue dar bias, ET e Sigma a um analito novo, e o manual não ensina isso (OC-9). Evidência:
  `f2r3\C\f2.json` → P5.I10_1_Hemoglobin A1c, P5.I10_2_Hemoglobina glicada e P5.I10_mapear.

### OC-4: Registro sem valor-alvo é aceito sem aviso e a amostra some do bias (Médio; tipo `funcional`)
- **Tela/módulo:** frmCEQ / `SalvarCEQ`.
- **Ação:** registrar resultado 100 sem o valor-alvo.
- **Esperado:** o alvo é indispensável ao bias. Ele deveria ser obrigatório, ou pelo menos gerar um aviso de que a amostra não entra no bias.
- **Observado:** `OK|550|`. O bias da linha fica vazio (O/P = "" e Bias/Bias_Abs vazios na EQA_Base), e AC (amostras no bias) continua 16. Nada avisa.
- **Consequência:** o usuário acha que registrou a rodada inteira, mas o bias é calculado com menos amostras.
- **Sugestão:** marcar o Valor-alvo com * (ou avisar ao gravar sem ele).
- **Evidência:** `f2.json` → P5.I5_sem_alvo.
- **[2º passe]:** reproduzida no Run A e no Run B. No Run B, a coluna S (limites) passou de "16 dentro dos limites" para "17": a amostra sem alvo entra na contagem dos limites e não entra no bias (AC = 16).

### OC-5: Correção do CEQ, consolidação e troca do provedor do analito não deixam rastro no Audit_Log (Médio; tipo `funcional`)
- **Tela/módulo:** EQA.CAP_Dados, `mEQA.AtualizarEQABase`, Painel M4 (`mEstatistica.GravarProvedorAnalito`).
- **Observado:**
  - Só a inclusão pela janela gera CEQ_REGISTRADO.
  - A edição 110 → 105, as consolidações e a troca de provedor CAP → Controllab → CAP, que muda a fonte do bias do analito, não geraram nenhuma linha: o Audit_Log tem 420 linhas tanto em S1 quanto em S2.
  - A limpeza de linhas (exclusão) também não deixa rastro.
- **Consequência:** não há como responder "qual era o resultado do CEQ antes da correção, quem mudou e quando". Isso vale também para a troca do provedor que alimenta o ET e o Sigma.
- **Sugestão:** registrar CEQ_ALTERADO/CEQ_EXCLUIDO, com o valor anterior e o novo, e PROVEDOR_ANALITO_ALTERADO.
- **Evidência:** `f2.json` → P4.trilha_da_edicao = [], P4.trilha_consolidar = [], P5.I8_provedor_controllab.trilha = []; Audit_Log de S1/S2.

### OC-6: O registro do CEQ não confere sessão nem papel (Médio; tipo `funcional`)
- **Tela/módulo:** `mCEQForm.SalvarCEQ` / `AbrirCEQ`.
- **Ação:** fazer logout e chamar o núcleo da janela. Pela lista de macros (Alt+F8), `mCEQForm.AbrirCEQ` é pública.
- **Observado:** `SessaoAtiva` = Falso, e mesmo assim `SalvarCEQ` devolveu `OK|551|` e gravou. O código não confere `SessaoAtiva` nem o papel, ao contrário da inativação (`PodeInativar`, ADR-081).
- **Consequência:** um CEQ que mexe no bias, no ET e no Sigma pode ser gravado sem identificação.
- **Sugestão:** exigir `SessaoAtiva()` (e um papel mínimo configurável, como `PAPEL_MIN_INATIVAR`) no `AbrirCEQ` e no `SalvarCEQ`.
- **Evidência:** `f2.json` → P6.

### OC-7: Com o bias em "SEM EP", o Painel mostra ET = 1,65 × CV, enquanto a Estatística deixa o ET vazio (Alto; tipo `calculo`; confirma a OC-2 da F3/Bioquímica)
- **Tela/módulo:** Painel H7:H8.
- **Ação:** trocar o provedor do Glicose para Controllab no Painel (M4). Esse provedor não tem CEQ utilizável.
- **Observado:** a Estatística H38:H39 fica vazia e L, vazio. O Painel mostra bias "SEM EP", mas ET **2,144 %** (N1) e **2,386 %** (N2), que é exatamente 1,65 × CV (1,2997 e 1,4462). O bias ausente é tratado como 0.
- **Esperado:** sem bias, não há ET (como na Estatística), ou a tela mostra "ET só de imprecisão".
- **Consequência:** o ET aparece muito menor que o real e não deixa ver que falta o CEQ. Comparado com o ETp de 10 %, parece desempenho ótimo.
- **Evidência:** `f2.json` → P5.I8_provedor_controllab.painel.B7:I8. A F3 tinha deduzido pela fórmula; aqui está confirmado no Excel.
- **[3º passe]:** reproduzida na Proteína total. Painel H7:H8 = 3,069 % e 3,632 % (= 1,65 × 1,8601 e 1,65 × 2,2013)
  com bias "SEM EP". Com o CEQ do CAP, o ET real era 6,85 % e 7,42 %. A fórmula do Painel é
  `H7 = E7*1,65 + ABS(IF(ISNUMBER(G7), G7, 0))`: o bias ausente vira 0.
- **[2º passe]:** reproduzida na Bilirrubina direta: Painel H7:H8 = 6,343 % e 2,880 % (= 1,65 × 3,8443 e 1,65 × 1,7457) com bias "SEM EP", contra um ETp de 39,9 %. Com o CEQ do CAP, o ET real era 31,9 % e 28,5 %.

### OC-8: As listas da janela misturam os nomes da simulação com os do CAP real (Baixo; tipo `usabilidade`)
- **Observado:**
  - As rodadas incluem "A", "B" e "C", que são as 90 linhas de simulação (Uso = NAO).
  - Os analitos trazem "Glicose", "Ureia", "Creatinina" etc. (os nomes da simulação, que estão mapeados) ao lado de "Glucose, serum", "Urea Nitrogen" etc.
  - A avaliação traz "OK".
  - Quem escolhe "Glicose" grava uma linha que entra no bias (mapa CAP\|Glicose), o que é inofensivo. Quem escolhe a rodada "A" mistura o registro com a simulação no nome.
- **Sugestão:** montar as listas só com as linhas de Uso = SIM.
- **Evidência:** `f2.json` → P1.

### OC-9 [2º passe]: O manual manda digitar o CEQ nas abas de dados, não cita a janela nem o "Consolidar", e não diz como corrigir (Médio; tipo `usabilidade`)
- **Tela/módulo:** GUIA_DE_USO §5–6 × botão "Registrar resultado do CEQ" (A3 das abas EQA.*_Resumo) e botão "Consolidar na EQA_Base" (J1 das abas EQA.*_Dados).
- **Ação:** um analista segue o manual para lançar e depois corrigir uma rodada do CEQ.
- **Esperado:** o manual descreve o caminho que a tela oferece (ADR-087): a janela para registrar, a lista "Registros desta rodada" para conferir, como corrigir e quando clicar em "Consolidar na EQA_Base".
- **Observado:** o §6 diz "digite cada rodada nas abas EQA.CAP_Dados / EQA.Controllab_Dados". Não há menção à janela, ao botão Consolidar, à correção, ao mapa de nomes do analito (OC-3) nem ao campo alvo indispensável (OC-4). Para o ANALISTA, digitar nas abas de dados não é possível: todas as células estão bloqueadas (OC-1).
- **Dificuldade:** quem segue o manual bate na proteção da aba. Quem acha a janela não sabe que, depois de corrigir, precisa consolidar.
- **Consequência:** reforça a OC-1 e a OC-2: um CEQ corrigido pode ficar fora do bias, do ET e do Sigma sem ninguém perceber.
- **Sugestão:** atualizar o §6 (e o §5) com o fluxo da janela, a correção e o "Consolidar", depois da decisão sobre a OC-1/OC-2.
- **Evidência:** `GUIA_DE_USO.md` linhas 199–232 (sem "frmCEQ", "Registrar resultado do CEQ" ou "Consolidar"); `f2r2\A\f2.json` e `f2r2\B\f2.json` → P0.telas (botões e OnAction) e P0.tabela_cap.colunas (Locked).
- **Prioridade:** Médio.

### OC-10 [3º passe]: Com o H3 vazio (TODOS), a Estatística continua calculando um lote só (o lote em uso, 8974) e diz "refletido do Painel" (Médio; tipo `usabilidade`)
- **Tela/módulo:** Estatística H4 (`=loteAnalise`) e C:F/H/N:Q, × Painel H3/H4 e B7:I8 (ADR-086).
- **Ação:** no Painel, com a Proteína total, apagar o H3 (todos os lotes) e atualizar a Estatística.
- **Esperado:** a Estatística segue o Painel. Ou mostra "todos os lotes" com média, CV e ET vazios, como o Painel
  (ADR-086, decisão 4), ou diz claramente qual lote está calculando e por quê.
- **Observado:**
  - O Painel H4 diz "H3 vazio = TODOS os lotes do período", e B7:I8 tem média, CV e ET vazios (n = 134 e 131).
  - A Estatística H4 mostra **"8974 (refletido do Painel)"** e calcula o lote 8974: Proteína total N1 n = 178,
    CV 94,98 %, ET 160,50 %, igual ao que mostrava com H3 = 8974.
  - O `loteAnalise` com o H3 vazio cai no lote em uso (8974; ver OC-F4-4).
  - O bias e o Sigma não mudam (F2-T37), mas o ET, a margem e o "Orçamento" (N:Q) da Estatística passam a falar de
    um lote que o usuário não escolheu.
- **Dificuldade:** o rótulo "refletido do Painel" contradiz o Painel, que está em TODOS.
- **Consequência:** quem abre a Estatística depois de olhar "todos os lotes" no Painel lê ET, margem e orçamento de
  outro lote como se fossem da seleção.
- **Sugestão:** no H3 vazio, a Estatística mostra "TODOS (ver Painel)" e deixa C:F/H vazios, como o Painel, ou diz
  "lote em uso 8974 (H3 vazio no Painel)". É a mesma raiz da OC-HF4-1 (Hematologia) e da OC-F4-4. Decisão do
  Orquestrador.
- **Evidência:** `f2r3\C\f2.json` → P5.I9_H3_vazio (H3 null; cab_estat H4 = "8974"; conferencia_analito; painel
  B7:I8); fórmula Estatística!H4 = `=loteAnalise`.
- **Prioridade:** Médio.

## 5. Dúvidas de regra de negócio (ao Especialista em CQ, pelo Orquestrador)

1. **Equipamento:** o CEQ não tem equipamento (ADR-087), e o mesmo bias entra no ET e no Sigma do DIMENSION 1 e do DIMENSION 2 (F2-T16). A amostra do CAP é dosada em um equipamento só. Isso está confirmado pelo laboratório? Ou o registro deveria dizer em qual equipamento a amostra rodou?
2. **Período:** o bias segue o Ano EP (N4) e não o período da Estatística (B5:D5). A instrução da linha 7 da Estatística diz que as exclusões de período tiram resultados de "n, Média, DP, CV, **Bias** e ET". O bias do CEQ, porém, não depende de data. O texto deve ser corrigido, ou o bias deveria respeitar as exclusões?
3. **Amostra "Unacceptable" ou "NAO AVALIADO":** essa amostra entra no bias do ET/Sigma (o `BiasEQ` não filtra a avaliação, e o `ViesEQ` filtra; ver OC-1 da F3). Está correto?
4. **Papel mínimo:** quem pode registrar e quem pode corrigir um CEQ? Hoje não há conferência (OC-6), e só o ADM corrige (OC-1).
5. **[2º passe] Sinal oposto entre rodadas:** na Bilirrubina direta (bias persistente de −30,8 %), uma rodada com +10 % baixa o |bias| do ET/Sigma para 25,6 % e tira o N1 da "Margem crítica". Isso é o que a média das médias de |bias| (ADR-035/071) manda, e o QC_INI está certo pela definição. A pergunta é se o Especialista quer um alerta quando as rodadas têm sinais opostos (o viés com sinal T já mostra −20,6 %).

6. **[3º passe] Sigma negativo com ETp cadastrado:** com |bias| > ETp (Proteína total, N4 = 2026, bias 30 % e
   ETp 10 %), a Estatística mostra Sigma −6,01 (N1) e −3,38 (N2), o que é coerente com (ETp − |bias|)/CVp. O
   ADR-068 proíbe Sigma negativo só quando falta o ETp. Mostrar um Sigma negativo é a decisão desejada, ou a tela
   deveria mostrar "< 0" / "ETp excedido pelo bias"? O mesmo aconteceu na Glicose (−3,12 no N2).
7. **[3º passe] Primeiro CEQ de um analito novo:** quem deve poder incluir um nome de analito no mapa W:Y da
   EQA_Base, e por qual tela (F2-T38, OC-3)?

## 6. Impressões de uso (Relatório A)

**[3º passe]:** a experiência com a Proteína total foi igual à dos passes anteriores:

- A sequência da janela é boa e as validações são claras.
- O que se acabou de registrar aparece na lista da rodada.
- O CEQ entra no bias do analito certo, e só nele, também com dois lotes.

O que mais incomoda o usuário continua nas bordas, e tudo em silêncio:

- **Analito novo:** não ganha o primeiro CEQ e não há aviso (F2-T38).
- **Correção:** só vale depois de consolidar (OC-2).
- **Troca de lote no Painel:** é um bom teste de confiança no bias, que não muda. Mas a Estatística no H3 vazio fala
  de um lote que ninguém escolheu (OC-10), e o período escolhido se perde na volta (OC-F4-7).

Cada gravação leva de 15 a 21 s.

- **A sequência faz sentido?** Sim. A janela segue a ordem do relatório do provedor (rodada → analito/amostra → resultado → avaliação/fonte). As validações apontam o campo certo, com mensagens claras em português. O "Salvar e próximo" mantém o contexto. A rodada existente traz o ano, e o analito traz a unidade.
- **Dá para achar a rodada e ver o que acabou de registrar?** Sim. A lista "Registros desta rodada" mostra analito, amostra, resultado e avaliação logo depois de gravar, e a contagem por rodada também. Não existe campo de **lote do material** do CEQ (o modelo não tem, ADR-087). O usuário que procura o lote não o encontra.
- **O CEQ entra no bias certo, e só nele?** Para inclusão pela janela, sim:
  - só o analito mapeado muda, nos dois níveis;
  - o provedor e o ano são respeitados;
  - o Painel acompanha;
  - os números batem com o cálculo independente até 1e-15.

  O problema está nas bordas, todas silenciosas:
  - nome fora do mapa e qualquer registro Controllab (OC-3);
  - registro sem alvo (OC-4);
  - correção, que o usuário comum não consegue fazer (OC-1) e que, feita pelo ADM, só vale depois de um botão em outra aba, enquanto o indicador diz "base em dia" (OC-2).
- **Tempo:** cada gravação leva de 15 a 24 s, por causa da consolidação e do recálculo da incerteza. É aceitável, mas sem um indicador de progresso parece travado.
- **Trilha:** a inclusão é rastreável. A correção, a exclusão e a troca de provedor não são (OC-5).

## 7. Limpeza

As linhas TESTE-AUDITORIA foram removidas da cópia pelo próprio caminho do usuário (limpar e consolidar), com retorno bit a bit ao estado original (F2-T20). Depois, a cópia `Bioquimica_F2.xlsm`, a cópia de exploração e os retratos `f2_S1.xlsm` e `f2_S2.xlsm` foram apagados. Ficam em `%TEMP%\qcw\auditoria\2026-10-31_ciclo01\f2\` só JSON e o log (sem senha nem caminho de rede). Criei e encerrei um único Excel, o PID 6320.

**[2º passe]:** cada run removeu as suas linhas TESTE-AUDITORIA pelo mesmo caminho e voltou bit a bit ao B0 (Estatística, K5, Painel e as 545 linhas da EQA_Base). As cópias `Bioquimica_F2_A.xlsm` e `Bioquimica_F2_B.xlsm` e os quatro retratos foram apagados. Ficam em `%TEMP%\qcw\auditoria\2026-10-31_ciclo01\f2r2\{A,B}\` só JSON, TXT e log. Excel: só os PIDs 18460 e 29336, um de cada vez, criados e encerrados pelo script. A produção continua com o sha `fe10dba1f1fd`.

**[3º passe]:** o Run C removeu as suas 8 linhas TESTE-AUDITORIA (7 CAP, 1 Controllab) pelo mesmo caminho. A EQA_Base
e o K5 voltaram bit a bit ao B0. A Estatística ficou diferente só nas colunas do CIQ do período, por causa do período
perdido no I9 (F2-T42). A cópia `Bioquimica_F2_C.xlsm` e os retratos `f2_S1.xlsm` e `f2_S2.xlsm` foram apagados. Ficam
em `%TEMP%\qcw\auditoria\2026-10-31_ciclo01\f2r3\C\` só JSON, TXT e log. Excel: só o PID 28540, criado e encerrado
pelo script. A produção continua com o sha `fe10dba1f1fd`. A ferramenta `f2_ceq.py` ganhou os passos opcionais I9 e
I10 e os argumentos `nome_fora_do_mapa` e `unidade`. Sem eles, o comportamento é o mesmo de antes.
