# Reteste 04 — Usuário Simulado — Bioquímica, frente F2

Data: 14/11/2026 · Agente: AGENTE_USUARIO_SIMULADO · Modo: RETESTE (OC-0040, 0043, 0056, 0057, 0058, 0059, 0066,
ADR-091, ADR-093, ADR-094) + pedido do ciclo: incluir e corrigir um CEQ TESTE-AUDITORIA e conferir bias, ET e Sigma do
analito certo, e só dele. Produção: entrega 41.

## Ambiente e método

| Item | Valor |
|---|---|
| Produção copiada (nunca aberta) | `QC_Bioquimica.xlsm` sha256 **8ec815ab653db405** (= entrega 41), igual antes e depois das duas execuções (`producao_intacta.igual = true`) |
| Cópias novas (nenhuma salva) | `%TEMP%\qcw\auditoria\2026-11-14_reteste04\Bioquimica_F2.xlsm` (parte 1) e `Bioquimica_F2_p2.xlsm` (parte 2) |
| Retratos para a calculadora | `SaveCopyAs` com eventos desligados: `Bioquimica_F2_{S_mes,S_trim,S0..S4,G0..G4}.xlsm` na mesma pasta; cópias e retratos apagados no fim (REGRAS §2.3), ficam os JSON |
| Excel | um por vez, aberto e fechado pelo `qa_final.QA`: PID 6920 (parte 1) e 3800 (parte 2). Os Excel `/automation` da tarefa agendada (ex.: 16560) não foram tocados. Nenhum outro agente com Excel aberto (sem python rodando) |
| Script | `_arquitetura/agentes/ferramentas/reteste04_usuario_f2.py` (modos: padrão, `--parte2`, `--so-calc`), documentado no cabeçalho. Usa a infraestrutura (`qa_final.QA`, `repro_comum`, `Audit`/`Vigia`/leitura de manual do `valida_adr090_091`); os oráculos (contagem do período, de/para do ADR-093, regras da tblLotesMaterial, varredura de datas, diferença célula a célula da Estatística) são meus. Matemática: `calc_independente.py` (minha, F3 do ciclo 01) |
| Saídas (fora do repositório) | `Bioquimica_F2.json/.log`, `Bioquimica_F2_p2.json/.log`, `Bioquimica_F2_calc.json`, `Bioquimica_F2_calcG.json`, `calc_*.json`, `calcG_*.json` |
| Usuário | `TESTE_AUDITORIA_F2`, papel ANALISTA, criado só na cópia; login por `mSeguranca.DoLogin` |
| Dados de teste | CEQ: rodadas "TESTE-AUDITORIA F2 R1" (Proteína total, CAP 2025 e 2024, Controllab) e "TESTE-AUDITORIA F2 R2" (Glicose CAP e Controllab), amostras TA-F2-*, arquivo TESTE-AUDITORIA.pdf; valores 7,7 / 7,35 / 14 sobre alvo 7 e 110 / 105 / 200 sobre alvo 100. Uma linha de mapa "TESTE-AUDITORIA GLU CTL" → Glicose (Controllab) na EQA_Base W39:Y39 da cópia, apagada depois. Manuais MAN_0001..0007 com MOTIVO "Outro: TESTE-AUDITORIA" (123,456 / 12,345 / 1,234 / −1 / 50). Inativação e comentário técnico de teste em BIO-356720 e BIO-356623 (só na cópia) |
| tblCQ_Final | 35.906 linhas; nenhum resultado novo do DB_SEAC durante o teste (último recebido 13/11/2026 08:39:01) |

**Caminho usado.** Filtros do Painel pelas células livres da aba **protegida** (L4, C3, H3, K3, K4, G3/G4), com os eventos
ligados, como o usuário. CEQ pelo núcleo que os botões da janela chamam (`mCEQForm.AvisoAnalitoCEQ`,
`ConfirmacoesCEQ`, `SalvarCEQ`, `CorrigirCEQ`, `ExcluirCEQ`, `DescricaoCEQ`, `RegistrosDaRodada`) e o `AbrirCEQ` real.
Manuais na aba Digitar Resultados (navegação `mIntegracao.IrDigitarResultados`, digitação célula a célula, eventos
ligados) e ATUALIZAR DADOS pelo núcleo do botão (`mIntegracao.AtualizarDadosAutomatico`, DB_SEAC só lido).
Inativar e comentar pelo núcleo da janela do ponto (`mIntegracao.InativarPorId`, `GravarComentarioTecnico`).

**Pelo VBA interno, sem caminho de usuário:** a linha de mapa do Controllab na EQA_Base (o mapa é mantido pelo
administrador; escrevi com eventos desligados, como o F1 fez com o CAP). **Não executado:** o clique nas caixas
Sim/Não do frmCEQ (sem interação com a tela).

## Resultado por item do reteste

| ID | Estado | Evidência |
|---|---|---|
| OC-0040 | **RESOLVIDA** | Estatística protegida; A4 "Visão (segue o Painel)", B4 "PERSONALIZADO" travada; D4 `=IF(Painel!K3="","",Painel!K3)`. Pelo Painel: Set/2026 → D4 2026, B5/D5 e H5/J5 01/09–30/09/2026; T3/2026 → 01/07–30/09/2026; "Tudo (sem filtro)" → H5 "histórico inteiro do lote", J5 vazio; De/Até 01/12/2025–31/01/2026 → K3 vazio, D4 vazio. **Nenhum "Ano 0"** em 6 estados. Escrever "MENSAL" em B4 é recusado (célula travada). Números: Glicose · 8977 · DIMENSION 1, n/média/DP/CV de N1 e N2 = minha contagem da tblCQ_Final no mês (23/24), no trimestre (55/56) e no histórico (90/88), dif. rel. máx. 1,5e-13; calculadora: 1.178 comparações sem divergência em S_mes e S_trim. O trimestre vem dos dados brutos (média N1 77,236), não da média dos meses (76,921) |
| OC-0043 | **RESOLVIDA** | J7 e J8: 48 pt úteis, Segoe UI 10 negrito, ShrinkToFit, sem quebra. Fator de redução de cada forma: REJEITADO 0,928, ALERTA/OK 1,0, DESATUAL. 0,914, BASE VAZIA 0,865, SEM MÉD/DP 0,771, AGUARDA 1ª 0,810, SEM CORRIDA 0,736, PER. VAZIO 0,901, SEM PONTO 0,842 → todas ≤ 30 % de redução (pior 26,4 %). Estados reais provocados pelo filtro: Glicose 8977 Out/2026 "REJEITADO" (Q4 REJEITADO); Jan/2025 "PER. VAZIO" (Q4 "SEM CORRIDA NO PERÍODO"); lote 6354 "SEM MÉD/DP" (Q4 "SEM MÉDIA/DP"); cores próprias |
| OC-0056 / ADR-093 | **RESOLVIDA** | Lote 6354 (MATRIZ URINA) + "Creatinina" 12,345 N1 → MAN_0002 **ATIVO**, PARTICIPA SIM, ANALITO/ANALITO_ORIGEM **CRE2**, ANALITO_QCINI e BLOCO vazios, MATRIZ URINA; DETALHE "analito 'Creatinina' digitado num lote de URINA: lancado como 'CRE2' (…regra MATRIZ da tblLotesMaterial)"; QA **I05** (INFO) do MAN_0002; trilha RESULTADO_MANUAL_INCLUIDO com a nota. Vizinho recusado: 6354 + "Lactato" (código LA nunca recebido na urina) → **LOTE_OUTRO_MATERIAL**, fora da estatística e do LJ, QA **A13**, trilha FORA_DA_ESTATISTICA. Oráculo próprio (de/para × interfaceamento): Creatinina → CRE2 aceito, Lactato recusado, como o sistema fez. Nenhuma Creatinina de soro participa no 6354 (oráculo ADR-091) |
| OC-0057 | **RESOLVIDA** | Registro com sessão (linha 547, 7,7) → `Logout` (SessaoAtiva False) → Salvar, **Corrigir** e **Excluir**: "ERRO\|0\|Sessão encerrada: faça o login de novo. Nada foi gravado."; F547 7,7 → 7,7; linha não apagada; nenhuma trilha CEQ_*. `AbrirCEQ` sem login: caixa "Controle externo (CEQ)" "Faça o login para registrar resultados do CEQ." (lida pelo texto da janela do meu PID), janela não abriu. Vizinho: login de novo → Corrigir "OK\|547\|" (F = 7,0) e Excluir "OK\|547\|", trilhas CEQ_ALTERADO/CEQ_EXCLUIDO + consolidação |
| OC-0058 | **RESOLVIDA** | Trilha por manual no 1º ATUALIZAR: IMPLAUSIVEL (−1, MAN_0004), LOTE_FORA_DO_CQ (8794, MAN_0005), LOTE_OUTRO_MATERIAL por SO_ANALITOS (8353 + Glicose, MAN_0006) e por MATRIZ (6354 + Lactato, MAN_0003) → **RESULTADO_MANUAL_FORA_DA_ESTATISTICA**; ATIVOS (Glicose 8977, CRE2, PCR 8353) → INCLUIDO. 2º ATUALIZAR (só entradas): nenhuma trilha de manual repetida |
| OC-0059 | **RESOLVIDA** | `ADR-090.md:20` "OC-0041 (…registrado aqui como OC-0040 até o ADR-092, N-05) \| UX G23" e `:36` "OC-0041 (G23; antes rotulado OC-0040)". O registro bate com o OCORRENCIAS.md (OC-0040 = G22, OC-0041 = G23). Resta no `ADR-090.md:106` a frase histórica "OC-0043: os status longos … continuam cortados", já superada pelo ADR-092 (não confunde a numeração) |
| OC-0066 / ADR-094 | **RESOLVIDA** | Parte 2: inativar BIO-356720, comentar BIO-356623, manual MAN_0001 e ATUALIZAR completo (Sync_Log 18 → 19 linhas). Varredura de **todas** as tabelas de todas as abas (1ª, última e até 40 linhas espalhadas de cada coluna, e as linhas novas): nenhum "####" e nenhuma data-hora sem segundos nas abas do escopo. tblCQ_Final[DATA_INATIVACAO] do BIO-356720 "14/11/2026 16:00:55"; Inativar "14/11/2026 16:00:55" (largura 18,14); Comentários "14/11/2026 16:00:55"; Digitar Resultados REGISTRADO_EM "14/11/2026 16:00:52"; Sync_Log INICIO/FIM "14/11/2026 16:00:56" / "16:01:19", DATA_MIN/MAX com segundos; QA DATA_HORA "dd/mm/aaaa hh:mm:ss"; DB A5 "último resultado: 13/11/2026 08:39:01 · última chegada: 14/11/2026 14:34:33"; DATA (só data) "dd/mm/aaaa"; Registros "09/10/2026". Fora do escopo do ADR (ver Observações): BI_Data e Audit_Log Data/DataCorrida |
| ADR-091 | **RESOLVIDA** (regra) — ver **R4-F2-02** (mensagem do Painel) | 17 regras. Oráculo próprio sobre a tblCQ_Final depois do ATUALIZAR: 0 violações. 8353 SO_ANALITOS: PCR 978 ATIVO + 1 IMPLAUSIVEL, outros 31 LOTE_OUTRO_MATERIAL, nenhum participando; 8177 (29) e 8794 (31) todos LOTE_FORA_DO_CQ; MATRIZ URINA 6353 (44), 6354 (929), 5657 (7), 5658 (15), 5659 (161): todas as linhas URINA, nenhum analito de soro participando. Principal A5 "Total: 35.913 · participam: 35.765 · … inativados: 12 · manuais: 7". Painel: 8353 + PCR mostra os pontos (janela de 180 corridas, aviso de 84 cortadas e "Período inteiro: N1 249 · N2 238" = minha contagem, ver nota¹); 8353 + Glicose e 8794 + Glicose: nenhum ponto (correto), mas a mensagem é enganosa (R4-F2-02) |
| ADR-094 | **RESOLVIDA** | Ver OC-0066 acima (mesma execução) |

¹ A minha contagem de PCR 8353 N1 (250) foi feita sobre a tblCQ_Final da parte 1, que incluía o manual de teste MAN_0007
(PCR 8353 N1, ATIVO); o Painel da parte 2 (sem esse manual) diz 249. Os dois batem.

### Pedido do ciclo: incluir e corrigir um CEQ TESTE-AUDITORIA

| ID | Teste | Estado | Evidência |
|---|---|---|---|
| CEQ-PT | Proteína total (CAP, "Protein, total, serum"), Ano EP 2025, rodada TODAS: incluir 7,7/7 (+10 %) e corrigir para 7,35 (+5 %) | **APROVADO** | Só as linhas 44–45 (Proteína total N1/N2) mudam: G, L, S, T, V, W, AC, AD, AJ..AM, AR, AT, AU (30 células na inclusão, 22 na correção); nenhuma célula de outro analito. Bias 3,37934 → 5,03450 → 3,78450; Sigma N1 1,41593 → 1,06195 → 1,32928; N2 2,99888 → 2,24916 → 2,81535; nEQA 15 → 16, rodadas 3 → 4. Os valores aparecem **sem** `Calculate` explícito (igual com e sem) |
| CEQ-GLI | Glicose (CAP "Glucose, serum"), lote 8977, DIMENSION 1, 2026 (CIQ no período, para o ET): incluir 110/100 e corrigir para 105 | **APROVADO** | Só Glicose N1/N2 (linhas 38–39) e o resumo de classes (B101 "Marginal" 6 → 4, B102 "Desempenho inadequado" 41 → 43, consequência da reclassificação) mudam. Bias 2,73329 → 4,54997 → 3,29997; ET N1 6,34568 → 8,16236 → 6,91236; ET N2 6,86198 → 8,67866 → 7,42866; Sigma N1 3,47195 → 2,60397 → 3,20120; N2 3,26871 → 2,45153 → 3,01381; Painel B7:I8 com os mesmos valores (ET do Painel difere 3e-14 relativo: o Painel recalcula o DP) |
| CEQ-ANO | Mesmo analito, CAP **2024** (+100 %), com Ano EP 2025 | **APROVADO** | Estatística sem nenhuma mudança (0 células) |
| CEQ-PROV | Controllab (+100 %) do mesmo analito: Proteína total sem mapa; Glicose com mapa de teste | **APROVADO** (cálculo) — ver **R4-F2-01** (mensagem) | Proteína total: "OK\|2\|Gravado fora do bias (analito não mapeado)." Glicose com mapa: EQA_Base Analito_Canonico "Glicose", Provedor Controllab, Bias 100 → Estatística **sem nenhuma mudança** (o provedor da Glicose é o CAP, Especificações!AR). CAP e Controllab não se misturam |
| CEQ-DUP | Duplicidade | **APROVADO** | "ERRO\|5\|Já existe este registro (CAP TESTE-AUDITORIA F2 R1 2025 · Protein, total, serum · TA-F2-01) na linha 547. Para corrigir, selecione-o em “Registros desta rodada” e clique em Corrigir selecionado." |
| CEQ-EXC | Excluir tudo | **APROVADO** | Estatística A14:AU133 **igual bit a bit** ao retrato inicial (S4 = S0; G4 = G0); EQA_Base sem linha TESTE-AUDITORIA; trilha CEQ_EXCLUIDO com os valores apagados + EQA_BASE_CONSOLIDADA |
| CEQ-TRILHA | Trilha e consolidação | **APROVADO** | CEQ_REGISTRADO / CEQ_ALTERADO ("Seu resultado: 7,7 → 7,35") / CEQ_EXCLUIDO, cada um com EQA_BASE_CONSOLIDADA "incluídas 1 / alteradas 1 / removidas 1"; Z3 "base em dia" |
| CEQ-CALC | Calculadora independente sobre os 12 retratos | **APROVADO** | `calc_independente.py`: 1.178 comparações por retrato (colunas C..O, T, V, W, AC, AD, AF, AR; 80 linhas), **0 divergências** em S_mes, S_trim, S0..S4 e G0..G4 (tolerância 1e-9 relativo) |
| ADR-093-AMB | Código ambíguo (mais de um ANALITO_ORIGEM) | **NAO_EXECUTADO** | Nenhum analito do catálogo da Bio tem mais de um código no de/para (varredura própria); o caminho só existe nas fixtures do Tester |

## Relatório B — validação matemática do CEQ (ADR-035/064/069/071/072/088, conferido no código)

Definição confirmada em `src/comum/mCEQ.bas` (`BiasEQ`, `AmostraDoVies`) antes de usar: amostras da EQA_Base do analito
canônico, provedor de Especificações!AR, ano vigente = Ano EP explícito (N4 = 2025), rodada TODAS; fora NAO AVALIADO,
resultado ou alvo 0 e chave repetida; bias_i = (Resultado − Alvo)/Alvo × 100; **duas etapas**: média dos |bias| de cada
rodada, depois média das rodadas. ET = 1,65 × CV do período (Estatística F) + |bias|. Sigma = (ETp − |bias|)/CV pooled
(janela de 6 meses, 14/05–13/11/2026, lotes com n ≥ 20). ETp = 10 (CLIA, Especificações S/T) para os dois analitos.

| Analito | Equip. | Lote | Nível | Período | Dados de entrada | Fórmula | Independente | QC_INI (Value2) | Dif. rel. | Status |
|---|---|---|---|---|---|---|---|---|---|---|
| Proteína total | DIM 1 | (bias não depende) | N1/N2 | CEQ 2025 | rodadas C-A 4,61503 · C-B 2,35503 · C-C 3,16794 (5 amostras cada) | média das médias | 3,3793363600341007 | 3,3793363600341007 | 0 | APROVADO |
| Proteína total | DIM 1 | — | N1/N2 | + TA-F2-01 7,7/7 | + rodada TESTE 10,0 | (10,13801 + 10)/4 | 5,034502270025575 | 5,034502270025576 | 2e-16 | APROVADO |
| Proteína total | DIM 1 | — | N1/N2 | corrigido 7,35/7 | rodada TESTE 5,0 | (10,13801 + 5)/4 | 3,7845022700255755 | 3,784502270025574 | 4e-16 | APROVADO |
| Proteína total | — | — | N1 | Sigma, janela 6 m | CVp N1 = (10 − 3,37934)/1,41593 = 4,67584 | (10 − 5,03450)/4,67584 | 1,0619486593 | 1,0619486593062732 | < 1e-9 | APROVADO |
| Proteína total | — | — | — | + CAP 2024 (+100 %) | ano ≠ Ano EP | não entra | sem mudança | sem mudança | 0 | APROVADO |
| Glicose | DIM 1 | 8977 | N1/N2 | CEQ 2025 | C-A 1,49097 · C-B 1,43622 · C-C 5,27270 | média das médias | 2,7332944432981674 | 2,733294443298167 | 2e-16 | APROVADO |
| Glicose | DIM 1 | 8977 | N1 | 01/01–31/12/2026 | n 90, CV 2,189325928446349 | 1,65·CV + 2,73329 | 6,3456822252 | 6,3456822252346425 | < 1e-12 | APROVADO |
| Glicose | DIM 1 | 8977 | N1/N2 | + TA-F2-G1 110/100 | + rodada TESTE 10,0 | (8,19988 + 10)/4 | 4,549970832473625 | 4,549970832473625 | 0 | APROVADO |
| Glicose | DIM 1 | 8977 | N1 | idem | CV 2,18933 | 1,65·CV + 4,54997 | 8,1623586144101 | 8,1623586144101 | 0 | APROVADO |
| Glicose | DIM 1 | 8977 | N2 | idem | CV 2,502236 | 1,65·CV + 4,54997 | 8,67866038695 | 8,67866038694817 | < 1e-12 | APROVADO |
| Glicose | DIM 1 | — | N1 | Sigma, 14/05–13/11/2026 | CVp N1 = (10 − 2,73329)/3,47195 = 2,09297 | (10 − 4,54997)/2,09297 | 2,6039655689620536 | 2,6039655689620536 | 0 | APROVADO |
| Glicose | DIM 1 | 8977 | N1/N2 | corrigido 105/100 | rodada TESTE 5,0 | (8,19988 + 5)/4; ET N1 1,65·2,18933 + 3,29997 | 3,299970832473625; 6,912358614410101 | 3,299970832473625; 6,912358614410101 | 0 | APROVADO |
| Glicose | — | — | — | + Controllab 200/100 mapeado | provedor ≠ AR (CAP) | não entra | sem mudança | sem mudança (0 células) | 0 | APROVADO |
| Glicose | DIM 1 | 8977 | N1/N2 | Set/2026 · T3/2026 · histórico | contagem própria da tblCQ_Final | n, média, DP (n−1), CV | 23/24 · 55/56 · 90/88 | iguais | ≤ 1,5e-13 | APROVADO |

**Caminho do dado (conferido nas trilhas e nos retratos):** janela do CEQ → aba EQA.CAP_Dados (linha 547) →
consolidação automática da EQA_Base (Analito_Canonico pelo mapa W:Y; Bias e |Bias| recalculados) → `BiasEQ`
(provedor AR, ano N4, rodada P4) → Estatística G/T/AC/AD → H (ET, com o CV do período F) e L (Sigma, com o CV pooled) →
M (classe), V/W (DPM, rendimento), Y/Z (plano), AJ..AM/AR/AT (incerteza e viés) → Painel B7:I8 do analito selecionado e
resumo de classes B101:B102.

## Ocorrências novas

### R4-F2-01 — Janela do CEQ: nome do Controllab mapeado para um analito cujo provedor é o CAP aparece como "entra no bias", e não entra

- **Prioridade:** Médio · **Tipo:** usabilidade · **Tela/módulo:** janela "Registrar resultado do CEQ"
  (`mCEQForm.AvisoAnalitoCEQ`, `SalvarCEQ`), Bioquímica (o código é comum; a Hema com R4 = CAP deve ter o mesmo
  comportamento, não testado).
- **Ação:** com um nome do Controllab mapeado para Glicose na EQA_Base (W:Y), registrar uma amostra do Controllab.
- **Esperado:** a janela avisar que o resultado **não** entra no bias, porque a Glicose usa o CAP (Especificações!AR),
  como já faz com o nome fora do mapa ("Gravado fora do bias").
- **Observado:** aviso azul "No QC_INI: Glicose. Este resultado entra no bias deste analito."; `SalvarCEQ` "OK\|2\|"
  sem nota. Na Estatística, nada muda (correto: CAP e Controllab não se misturam). O `AvisoAnalitoCEQ` olha só o mapa
  do provedor, sem conferir o provedor configurado do analito.
- **Dificuldade / consequência:** o analista acha que o CEQ do Controllab já está no bias, no ET e no Sigma, e não
  percebe que precisa trocar o AR (ou que o resultado não conta). Hoje é **latente**: o mapa do Controllab da Bio está
  vazio, e qualquer nome do Controllab cai no aviso âmbar correto.
- **Sugestão:** no aviso e no retorno do Salvar, comparar o provedor digitado com o AR do analito canônico (Bio) ou com
  o R4 (Hema): "No QC_INI: Glicose. **O bias da Glicose usa o CAP**: este resultado fica gravado, fora do bias."
- **Evidência:** `Bioquimica_F2_p2.json` (`ceq_glicose.aviso_ctl`, `G3_controllab.salvar`, `diff_G2` = []);
  `src/comum/mCEQForm.bas:291-309`.

### R4-F2-02 — Painel em lote com regra FORA_DO_CQ ou SO_ANALITOS: "SEM CORRIDA / SEM MÉD/DP" e "digite a Média e o DP", sem citar a regra

- **Prioridade:** Médio · **Tipo:** usabilidade · **Tela:** Painel (J7:J8, Q4, A39), ADR-091.
- **Ação:** escolher o lote 8794 (FORA_DO_CQ) com Glicose, ou o 8353 (só PCR) com Glicose.
- **Esperado:** nenhum ponto (correto) e um texto que diga por quê: "lote fora do CQ" / "este lote só vale para PCR"
  (regra da tblLotesMaterial), como o DETALHE e o QA A13 dizem.
- **Observado:** J7 "SEM CORRIDA", J8 "SEM MÉD/DP", Q4 "SEM MÉDIA/DP", A39 "⛔ LOTE 8794 · Glicose: SEM MÉDIA/DP em N2 —
  digite a Média e o DP no bloco do lote (aba Analitos). … O lote não tem resultado de Glicose no DIMENSION 1." Os dois
  lotes **têm** 31 resultados de Glicose (LOTE_FORA_DO_CQ / LOTE_OUTRO_MATERIAL).
- **Consequência:** o usuário procura o resultado que "não chegou" ou digita Média/DP num lote que não é de CQ. Não
  muda estatística. Parente da OC-0061 (lotes só de urina).
- **Sugestão:** no A39/Q4, com a regra do lote ativa, mostrar "Lote 8794 fora do CQ (tblLotesMaterial): resultados
  guardados, fora da estatística" e não pedir Média/DP.
- **Evidência:** `Bioquimica_F2_p2.json` (`painel_lotes.2_Glicose_com_8353`, `3_lote_8794_FORA_DO_CQ`);
  `Bioquimica_F2.json` (`adr091.por_regra`).

### R4-F2-03 — Estatística D4 mostra "2026 (refletido do Painel)" com o Painel em "Tudo (sem filtro)"

- **Prioridade:** Baixo · **Tipo:** usabilidade · **Tela:** Estatística A4:J5.
- **Ação:** no Painel, K3 2026 e depois K4 "Tudo (sem filtro)".
- **Observado:** K3 continua 2026, e a Estatística mostra "Ano 2026 (refletido do Painel)" ao lado de H5 "histórico
  inteiro do lote" (e de B5/D5 "00/01/1900", o R4-02 do F1, reproduzido aqui). O cálculo é do histórico inteiro (n 90/88
  = minha contagem).
- **Consequência:** quem imprime a aba lê "Ano 2026" num número que inclui outros anos, se o lote tiver.
- **Sugestão:** com G3:G4 vazios, D4 vazio (ou "todos"), junto com a correção do B5/D5 proposta no R4-02.
- **Evidência:** `Bioquimica_F2.json` (`O40.3_tudo`).

## Observações (não são ocorrências desta frente)

- **BI_Data** (visível a todos depois do login, `mSeguranca.UnlockApp` "Case Else"): a coluna Usar_1_3s mostra
  "##########" (largura 10,71) e as datas estão em aaaa-mm-dd. A aba é do BI, que é proibido mexer (REGRAS §2.7) e está
  fora da lista do ADR-094. Fica como dúvida (abaixo).
- **Audit_Log Data/DataCorrida** só com data, embora guardem hora: decisão do ADR-094 ("O que NÃO muda", hash).
- **DETALHE sem acento** ("lancado", "catalogo", "estatistica", "grafico") nos textos do ADR-091/093 e do IMPLAUSIVEL:
  já registrado pelo F1 deste ciclo (OC-0021 PERSISTE).
- O tempo de cada Salvar/Corrigir/Excluir do CEQ ficou em ~15–20 s (trilha), dentro dos 15–25 s do GUIA §5. O
  ATUALIZAR completo levou 23–26 s; o só-entradas 15 s.

## Dúvidas de regra de negócio

1. **Urina sem tela:** o manual aceito pelo ADR-093 (CRE2) e toda a série de urina participam da estatística
   (PARTICIPA SIM), mas nenhuma tela os mostra: a lista do C3 é `=lstAnalitos` (catálogo) e a Estatística lista só o
   catálogo. No Painel, 6354 + "Creatinina" mostra o alvo do bloco da Creatinina (73,990 ± 0,0909) e "SEM CORRIDA". Onde
   o laboratório acompanha o CQ de urina (e confere o manual que digitou)? Liga-se à OC-0061 e à decisão do ADR-091
   ("fora do Painel").
2. **Provedor no aviso do CEQ (R4-F2-01):** o aviso deve considerar o AR (Bio) / R4 (Hema)? Ou o mapa do Controllab da
   Bio deve ficar vazio enquanto o AR for CAP?
3. **BI_Data visível com "####":** deve ficar oculta para quem não é ADM (como Calc e Cfg_Integracao)? Decisão fora do
   escopo do BI, mas é uma aba que o analista vê.
4. **ADR-093, código ambíguo:** não há caso real na Bio; o caminho recusado por ambiguidade só existe nas fixtures. Manter
   assim ou documentar no GUIA que hoje todo analito do catálogo tem um código só?
5. **Lote de urina + "Glicose":** pelo de/para atual, um manual "Glicose" num lote 635x vira GLUC (urina), porque o
   equipamento já mandou GLUC na urina (lotes 565x). É o esperado pelo laboratório?

## Impressões de uso (Relatório A)

- **CEQ, inclusão e correção:** a sequência é a do relatório do provedor e o aviso azul "No QC_INI: Proteína total. Este
  resultado entra no bias deste analito" dá segurança. A lista "Registros desta rodada" mostra o que acabou de ser
  gravado (analito, amostra, resultado, avaliação, linha), e a duplicidade explica onde está o registro e o que fazer.
  O número muda na hora (sem precisar de ATUALIZAR ou recalcular), só no analito certo, e volta exatamente ao que era
  quando se exclui. É o ponto mais sólido do sistema hoje.
- **Sessão:** sem login, a janela nem abre e a mensagem diz por quê; com a sessão encerrada no meio, nada é gravado e a
  mensagem é clara.
- **Período da Estatística:** a aba agora diz que segue o Painel, sem campos mortos e sem "Ano 0". Mês e trimestre se
  escolhem no Painel (K3/K4) e a Estatística acompanha. O "Tudo (sem filtro)" ainda deixa o cabeçalho confuso (R4-F2-03
  e R4-02).
- **Status do Painel:** as formas curtas cabem e o Q4 dá a forma longa. Mas, nos lotes com regra da tblLotesMaterial, o
  texto manda digitar Média/DP em vez de dizer que o lote está fora do CQ (R4-F2-02).
- **Manuais de urina:** o DETALHE explica a troca para CRE2 e o QA dá o I05; o recusado diz o motivo e o A13 orienta.
  Falta onde conferir a série de urina depois (dúvida 1).
- **Datas:** com segundos em todas as abas do dia a dia, sem "####", inclusive na linha nova do Sync_Log e na
  DATA_INATIVACAO recém-gravada.

## Arquivos

- Script: `_arquitetura/agentes/ferramentas/reteste04_usuario_f2.py`; calculadora: `ferramentas/calc_independente.py`.
- Saídas (fora do repositório): `%TEMP%\qcw\auditoria\2026-11-14_reteste04\Bioquimica_F2*.{json,log}`, `calc_*.json`,
  `calcG_*.json` (os retratos `.xlsm` foram apagados; o script os recria).
- O Audit_Log da cópia grava nome, login do Windows e máquina de quem roda; esses dados foram omitidos deste relatório.
