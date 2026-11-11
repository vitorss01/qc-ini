# Tester/QA, modo VALIDAR: ADR-090 (entrega 37) e ADR-091 (entrega 38) — reteste03, 11/11/2026

## 0. Resumo

| OC | Produto | Veredito | Evidência curta |
|---|---|---|---|
| OC-0012 | Bio | **VALIDADA** | Glicose · DIMENSION 1 · 8974 · "Tudo": M1 = 39 cortadas, AF1 "N1 174 · N2 161" = contagem independente; janela do Painel 140/131 (igual ao ciclo 01); A39 "…fora do gráfico e dos indicadores do Painel (máx. 180). Período inteiro: N1 174 · N2 161…". Ampliação Cloro 8974: 116 cortadas, "N1 250 · N2 224" = independente |
| OC-0012 | Hema | **VALIDADA** | WBC, H3 vazio, Ano inteiro 2026: 272 cortadas, "N1 372 · N2 394 · N3 338" = independente, janela 123/157/114. Ampliação RBC: idem |
| OC-0019 | Bio | **VALIDADA** | Painel em DIMENSION 2, manual sem EQUIPAMENTO: "DIMENSION 2", fundo 10284031 (= RGB 255,235,156), fonte 22428 (= RGB 156,87,0), nota exata "Preenchido pelo Painel com DIMENSION 2: confira antes de ATUALIZAR DADOS."; digitar DIMENSION 1 tira cor e nota; instrução A3 cita EQUIPAMENTO |
| OC-0019 | Hema | **VALIDADA** | XN1000 preenchido sem âmbar (fundo branco) e sem nota; mesma instrução |
| OC-0021 | Bio, Hema | **VALIDADA** | RUN vazio (Text "") em MANUAL_INCOMPLETO (2) e CONFLITO_MANUAL; DETALHE "campos obrigatórios ausentes/inválidos: NIVEL"; ATIVO com DETALHE ""; 300 linhas varridas sem "0" e sem texto sem acento |
| OC-0023 | Bio, Hema | **VALIDADA** | sem NÍVEL e sem HORA → só RESULTADO_MANUAL_RECUSADO (detalhe com o campo); conflito → RESULTADO_MANUAL_EM_CONFLITO; completar o NÍVEL → ATIVO, RUN 26092201 (Bio) / 26092901 (Hema) e RESULTADO_MANUAL_INCLUIDO. Ampliação: ver N-03 |
| OC-0024 | Hema | **VALIDADA** | 2 E04 e nenhum outro ERRO: "• 2 resultado(s) manual(is) incompleto(s)…", sem "governança" (178 caracteres) |
| OC-0024 | Bio | **VALIDADA** | 2 E04 + 4 E01 reais: linha do E04 com 2 e linha de governança com exatamente 4; sem "Há ERRO DE GOVERNANÇA" (323 caracteres) |
| OC-0031 | Bio, Hema | **VALIDADA** (função) | sem alvo: confirmação "…NÃO entra no bias (nem no ET e no Sigma)…", OK "Gravado fora do bias (sem valor-alvo).", G **inalterado** (Bio 3,3793363600341007; Hema 1,533620125276759); com alvo, G muda (5,0345… / 3,6502…); "Valor-alvo ¹" e nota no lblSub |
| OC-0033 | Bio, Hema | **VALIDADA** | sem login: SalvarCEQ "ERRO\|0\|Sessão encerrada…", 0 linhas gravadas; AbrirCEQ mostrou a caixa "Faça o login para registrar resultados do CEQ." (lida pelo texto da janela do meu PID) e a janela não abriu. Ampliação: ver N-02 |
| OC-0034 | Bio, Hema | **VALIDADA** | rodada/analito/avaliação da janela = oráculo (Uso ≠ simulação): Bio 3/31/2, Hema 3/25/3; "A/B/C", "1/2/3", "OK", "RBC/WBC/HGB/HCT", "Glicose/Creatinina…" da simulação fora |
| OC-0035 | Bio, Hema | **VALIDADA** (texto) | GUIA §3 cita equipamento e âmbar; §5–6 citam a janela, login, valor-alvo, Corrigir/Excluir selecionado, Z3, Consolidar, ATUALIZAR, FÓRMULA SOBRESCRITA e o R4 da Hema; nenhum "digite nas abas"; todos os botões citados existem no arquivo |
| OC-0036 | Bio, Hema | **VALIDADA** | 62 (Bio) e 84 (Hema) textos "x,xx / y,yy", nenhum "004 / 000"; Bio AT14 "3,69 / 0,26", Hema AT14 "2,49 / 0,04" (valores do parecer) |
| OC-0040 | Bio | **NAO_VALIDADA** | Estatística D4 continua `=Painel!I3` e mostra "0" ("Ano 0 (refletido do Painel)") com o Painel em K3 = 2026; B4 "PERSONALIZADO" bloqueada. O ADR-090 implementou o G23 (registrado como OC-0041) com o rótulo "OC-0040" |
| OC-0043 | Bio, Hema | **NAO_VALIDADA** | "REJEITADO" cabe (51,8 pt em 48 pt, redução 7 %); mas o critério do UX é **cada status** com redução ≤ 30 %, e "SEM CORRIDA NO PERÍODO", "AGUARDANDO 1ª CORRIDA" e "SEM CORRIDA DESTE ANALITO" precisariam de 62–67 % (fator 0,33–0,38). O ADR-090 deixou esses três como pendência |
| OC-0044 | Bio | **VALIDADA** | 96 linhas de analito com A:B centralizadas (−4108), inclusive Lactato N2, Ácido úrico N2 e Cálcio (linhas de 33 pt) |
| OC-0049 | Hema | **NAO_VALIDADA** | o "0" e o "SEM EP" sumiram (A, G, R, S, T vazias nas 36 linhas sem analito, 98 a 133), mas B98:B133 continua mostrando o nível 1/2/3 (constante), contra o critério "A98:T133 vazias" |
| OC-0050 | Bio, Hema | **VALIDADA** | "Total: 35.466 · participam: 35.323" e "Total: 48.081 · participam: 47.399" = contagem da tblCQ_Final |
| OC-0052 | Hema | **VALIDADA** | "MCH" → "Atenção: existe também “MCHC” (MCHC)"; o par de palavra continua primeiro ("Lymphs" → "Lymphs Absolute") |
| OC-0052 | Bio | **NAO_SE_APLICA** | o mapa do CAP da Bio (37 nomes) não tem nenhum par por prefixo sem espaço (varredura própria); o par de palavra "Lactate" → "Lactate Dehydrogenase" continua |
| OC-0053 | Bio, Hema | **VALIDADA** | O, P, S..X com validação `=1=0` (tipo 7), 2,55 reprovado (Validation.Value = False); colando 2,55 em P2: Z3 "⚠ FÓRMULA SOBRESCRITA…" e CEQ_FORMULA_SOBRESCRITA "P2 (\|Bias\| (%)) = 2,55"; restaurar a fórmula volta a "base em dia"; G inalterado nos três momentos |
| OC-0054 | Bio, Hema | **VALIDADA** | mesmo valor sem selecionar: nenhuma trilha, Z3 "base em dia"; valor diferente: CEQ_ALTERADO "Seu Resultado: 120 → 121" (Bio) / "4,8 → 5,8" (Hema), sem "?" |
| OC-0055 | Bio, Hema | **VALIDADA** | rodada "TESTE-AUDITORIA C-A 2025", ano 2025: descrição "CAP TESTE-AUDITORIA C-A 2025 · …", nenhum "2025 2025" na descrição, na lista da rodada e na trilha (registro e exclusão) |
| ADR-091 | Bio | **VALIDADA** | 17 regras (10 MATRIZ, 6 FORA, 1 SO); 8353: 965 PCR ATIVOS participando, 30 de 29 outros analitos LOTE_OUTRO_MATERIAL e 0 participando; 8177 (29) e 8794 (30) LOTE_FORA_DO_CQ, 0 participando; 5657/5658/5659/6353/6354: 1.081 linhas, todas URINA, 0 em analito sérico; A13 = 89 = LOTE_*, A14 = 0; equivalência VBA × PQ (--fixtures) 35.503 linhas, 0 divergências. Efeito colateral: N-01 |
| ADR-091 | Hema | **VALIDADA** | tabela presente e vazia, 0 LOTE_*, A13 = A14 = 0 |

Regressão (cópias da produção atual, uma suíte por vez): ver §4. Bio: qa_final **27 PASS / 2 FAIL** (T15 e T19, causa: resultados
novos do DB_SEAC chegando durante a suíte; ver §4), qa_rodadas_cv 12/0, qa_adr090 22/0, qa_adr091 7/0, equivalencia_cq OK.
Hema: qa_final 29/0, (demais em §4).

**Ocorrências novas (para o Orquestrador registrar):** N-01 (Alto), N-02 (Médio), N-03 (Baixo), N-04 (Baixo), N-05 (Baixo).

Quem corrigiu foi o agente programador (ADR-090: commits bbcd5a4/105ecbc; ADR-091: bfcc5db). Eu não editei `src/`,
instaladores nem suítes. O script de validação é meu (`_arquitetura/agentes/ferramentas/valida_adr090_091.py`) e não
reaproveita o `qa_adr090.py` nem o `qa_adr091.py`; a regressão rodou pelo `ferramentas/regressao_reteste03.py`.

## 1. Ambiente

- **Produção:** a de 11/11/2026, depois da entrega 38 e do ATUALIZAR. Bio `8b1cda1345d2d6ff`, Hema `f1907656b663ea6b`
  (iguais ao MANIFESTO). Só copiada (`shutil.copy2`); o sha continuou o mesmo no fim de cada execução (conferido pelo
  script: `producao_intacta.igual = true` em todos os JSON) e no fim da regressão.
- **Cópias:** `%TEMP%\qcw\auditoria\2026-11-11_reteste03\` (`<produto>_valida.xlsm`, `regressao\<produto>\copia_<suíte>\`).
  Nada foi salvo nelas.
- **Excel:** um por vez, aberto e encerrado pelo `qa_final.QA` (PIDs 16912, 31572, 23360 na Bio; 17400 na Hema; os da
  regressão estão nos logs). O Excel 31528 (e um 16400 que apareceu e saiu sozinho durante a rodada da Hema) não era
  meu e não foi tocado. Nenhuma captura de tela: a caixa do AbrirCEQ foi lida pelo texto das janelas `#32770` do
  **meu** PID (`win32gui`) e fechada com WM_CLOSE.
- **Dados de teste (só nas cópias):** usuário `TESTE_AUDITORIA_R3` (ANALISTA); manuais MAN_0001..0008 com MOTIVO
  "Outro: TESTE-AUDITORIA", valores 123,456 / 111,111 / 12,345 / 1,234 / −1; CEQ rodada "TESTE-AUDITORIA C-A 2025",
  amostras TA-31/TA-55, 110 com alvo 100 e sem alvo; P2 sobrescrito com 2,55; F2 +1.
- **Logs e JSON:** `valida_bio_rodada1.log` / `valida_Bioquimica_rodada1.json` (1ª rodada inteira),
  `valida_Bioquimica_ADR091_manuais.json` e `valida_Bioquimica_manuais.json` (rodadas 2 e 3, ver abaixo),
  `valida_hema.log` / `valida_Hematologia.json`, `regressao.log` e `regressao\<produto>\*.log|json`.
- **Rodadas do meu script na Bio (o que mudou e por quê):**
  1. 1ª rodada completa. Dois defeitos **do meu script**:
     - o oráculo do de/para da URINA esperava ANALITO_QCINI, mas na Cfg_Integracao a URINA está "sem cadastro"
       (ANALITO_QCINI vazio, o ANALITO fica o da origem: CRE2, UCFP…);
     - o manual "incompleto" foi feito sem MOTIVO, e o MOTIVO só é obrigatório com SUBSTITUI_ID
       (`mEtlCQ.bas:1046-1048`), então ele ficou ATIVO.
  2. Corrigi o oráculo e passei a fazer o incompleto **sem NÍVEL** (é o cenário do ciclo 01, chave "0|0|GLICOSE") e
     um 2º **sem HORA**. Repeti só as seções ADR091 e manuais (`--secoes`).
  3. Repeti a seção de manuais com a trilha sem truncar (o "NIVEL" do detalhe estava depois do corte de 300 caracteres
     do meu log).

  A Hema rodou uma vez, já com as correções.

## 2. Por ocorrência

### OC-0012: janela de 180 corridas (parecer CQ G11)

Oráculo: resultados com PARTICIPA = SIM e RESULTADO numérico do analito, lote (no modo TODOS, todos), equipamento (Bio)
e datas De/Até do Painel, contados na tblCQ_Final da cópia.

| Cenário | M1 (cortadas) | AF1 | Independente | n da janela (Painel) |
|---|---|---|---|---|
| Bio Glicose · DIM 1 · 8974 · Tudo (critério 1 do CQ) | 39 | N1 174 · N2 161 | N1 174 · N2 161 | 140 / 131 |
| Bio Cloro · DIM 1 · 8974 · Tudo (ampliação) | 116 | N1 250 · N2 224 | N1 250 · N2 224 | 148 / 135 |
| Hema WBC · H3 vazio · Ano inteiro 2026 (critério 2) | 272 | N1 372 · N2 394 · N3 338 | igual | 123 / 157 / 114 |
| Hema RBC · H3 vazio · Ano inteiro 2026 (ampliação) | 272 | N1 372 · N2 394 · N3 338 | igual | 123 / 157 / 114 |

O A39 diz "⚠ N corrida(s) antigas do período fora do gráfico e dos indicadores do Painel (máx. 180). Período inteiro: …
— veja a aba Estatística." O n da janela da Glicose é o do ciclo 01 (140/131): o cálculo não mudou (critério 3).
O critério 2 do CQ citava 337/340/308; hoje são 372/394/338 porque a base cresceu (48.081 linhas).

### OC-0019: EQUIPAMENTO do manual (UX G01)

Bio, com o Painel em DIMENSION 2 (o critério do UX): o 1º toque na linha grava "DIMENSION 2", fundo 10284031 =
255 + 235·256 + 156·65536 e fonte 22428 = 156 + 87·256, e a nota tem o texto exato. Escrevendo "DIMENSION 1" por cima
(eventos ligados), o fundo vira 16777215 e a nota sai. Na Hema, "XN1000" entra sem cor e sem nota. A instrução A3 cita
"EQUIPAMENTO vazio recebe o do Painel (fica em âmbar): confira" nos dois.

Observação que continua (já vista no ciclo 01): a validação de dados do EQUIPAMENTO da Bio (lista
"DIMENSION 1,DIMENSION 2") dá `Validation.Value = False` para "DIMENSION 1" escrito pelo COM, em todas as 8 linhas.
Conferir digitando, pelo Usuário Simulado.

### OC-0021, OC-0023, OC-0024: manuais e trilha (UX G03, G05, G06)

Após `AtualizarSoEntradas` (eventos ligados, dados digitados com FormulaLocal):

| Manual | STATUS | RUN (Text) | DETALHE (Text) | Trilha |
|---|---|---|---|---|
| M1 sem NÍVEL (e sem EQUIPAMENTO) | MANUAL_INCOMPLETO | "" | campos obrigatórios ausentes/inválidos: NIVEL | RESULTADO_MANUAL_RECUSADO ("…invalidos: NIVEL \| manual lancado por TESTE_AUDITORIA_R3…") |
| M8 sem HORA | MANUAL_INCOMPLETO | "" | …: HORA | RESULTADO_MANUAL_RECUSADO (HORA) |
| M2 = real +5 min (Hema +1 min) | CONFLITO_MANUAL | "" | coincide com resultado do interfaceamento (5 min) | RESULTADO_MANUAL_EM_CONFLITO |
| M1 depois de digitar o NÍVEL | ATIVO | 26092201 (Hema 26092901) | "" | RESULTADO_MANUAL_INCLUIDO |

- A varredura das 300 linhas da tabela não achou "0" em RUN/DETALHE nem texto sem acento.
- O texto da trilha continua sem acento ("obrigatorios"). Ele vem do MOTIVO_EXCLUSAO_AUTOMATICA, que o ADR-090 deixou
  sem acento de propósito (os dois motores).

ResumoQA:
- **Hema:** "QA da integração: 2 erro(s)…" + "• 2 resultado(s) manual(is) incompleto(s): veja o DETALHE…", sem
  "governança".
- **Bio:** a linha do E04 (2) e "• 4 erro(s) de governança…". Os 4 são E01 **reais** da produção, inativações de
  Lactato N2 sem MOTIVO de 13/10: BIO-366027, 366033, 366054 e 366164. É dado do laboratório. Não é o E10 que o ADR-091
  previa.

### OC-0031, OC-0033, OC-0034, OC-0055: janela do CEQ (UX G13, G15, G16; reteste02)

- **OC-0031:**
  - Fluxo, com G = bias da Estatística dos níveis do analito:
    - B0;
    - gravar TA-31 **sem** alvo → G igual a B0, bit a bit;
    - gravar TA-55 **com** alvo 100 → G muda;
    - excluir.
  - G da Bio: 3,3793363600341007 → 3,3793363600341007 → 5,034502270025575. Da Hema: 1,533620125276759 → igual →
    3,650215093957569.
  - Fica de fora a caixa "Gravar mesmo assim?" (Sim/Não) do formulário. Ela não é exercitável por COM: validei a função
    `ConfirmacoesCEQ` e o retorno do `SalvarCEQ`.
- **OC-0033:**
  - SalvarCEQ sem sessão foi recusado no arquivo recém-aberto e de novo depois do `mSeguranca.Logout`.
  - `AbrirCEQ` sem sessão:
    - abriu uma caixa `#32770` com o título "Controle externo (CEQ)" e o texto "Faça o login para registrar resultados
      do CEQ.";
    - nenhum `ThunderDFrame` (o frmCEQ) apareceu;
    - a caixa foi fechada em 0,1–0,3 s.
- **OC-0034:**
  - O oráculo é o conjunto distinto de cada coluna nas linhas cujo Arquivo fonte (coluna R) não contém "SIMULA".
  - O resultado é igual ao `DistintosCEQ` nas 3 colunas.
  - Só o CAP tem dados: o Controllab está vazio nos dois produtos.
- **OC-0055:** a descrição, os "Registros desta rodada" e a trilha (CEQ_REGISTRADO, CEQ_EXCLUIDO) não repetem o ano.

### OC-0035: GUIA (UX G17)

- **Texto conferido:** §3 (equipamento e âmbar) e §5–6 (todos os termos do critério). Nenhuma instrução manda digitar
  nas abas EQA.
- **Botões que o GUIA cita, procurados no arquivo:**
  - "Registrar resultado do CEQ" nos dois Resumos;
  - "Consolidar" nas abas de dados do CEQ;
  - "Salvar e próximo", "Corrigir selecionado" e "Excluir selecionado" no frmCEQ;
  - "ATUALIZAR DADOS" no Painel, na DB e na Digitar Resultados.
- **Fora daqui:** a parte do critério "o Usuário Simulado segue só o GUIA" é dele.

### OC-0036, OC-0044, OC-0050: formatos

- **AT:** são textos de UDF (`mCEQ.ViesEQ`), então não recalculei o RMS. Conferi três coisas:
  - o formato "n,nn / n,nn" nas 62/84 linhas;
  - os valores do critério (Bio AT14 3,69; Hema AT14 2,49);
  - a ausência de "0dd" sem vírgula.
- **A5:** igual a ROWS e a COUNTIFS(SIM) contados por mim na tabela.

### OC-0040: Estatística MENSAL/TRIMESTRAL e "Ano 0" (UX G22) — NAO_VALIDADA

- **O que a cópia mostra:**
  - Estatística!C4:E4 = "Ano" | `=Painel!I3` → "0" | "(refletido do Painel)";
  - Painel!I3 vazio e Painel!K3 = 2026;
  - B4 "PERSONALIZADO" e D4 bloqueadas (`Locked = True`).
- **Critérios do G22 que não se cumprem:** "nenhum Ano 0", "D4 = Painel!K3" e o rótulo da Visão.
- **Causa, no registro:** o ADR-090 (`ADR-090.md:20`) chama de OC-0040 a correção do **G23**, mas no OCORRENCIAS.md a
  OC-0040 é o G22 ("Visões MENSAL/TRIMESTRAL inalcançáveis e Ano 0"). O G23 é a OC-0041.
- **Correção proposta no parecer UX (sem autorização pendente):** D4 = `=Painel!K3`.
- **Bônus (OC-0041 = G23, que de fato foi corrigido):** com o Painel protegido sem UserInterfaceOnly, testei esta
  sequência:
  1. K4 "S1";
  2. digitar De/Até 01/09–31/10/2026, que deu K4 "Personalizado" e K3 2026;
  3. escolher K4 "Out", que reescreveu G3/G4 para 01/10–31/10.

  Comportamento conforme o critério do G23 (com o rótulo curto que o ADR justificou). Proponho ao Orquestrador
  registrar a OC-0041 como AGUARDA RETESTE/validada por esta evidência e manter a OC-0040 aberta.

### OC-0043: J7:J9 (UX G25) — NAO_VALIDADA pelo critério

Medida com a fonte efetiva (`DisplayFormat`: Segoe UI 10 negrito) e `auditoria_abas.largura_pt`; largura útil = J7
mesclada − 3 pt = 48 pt; ShrinkToFit = True em J7:J8 (Bio) e J7:J9 (Hema).

| Status | Texto (pt) | Fator de redução necessário |
|---|---|---|
| REJEITADO | 51,8 | 0,928 (cabe) |
| ALERTA / OK | 36,0 / 13,5 | 1,0 |
| SEM CORRIDA NO PERÍODO | 129,0 | 0,372 |
| AGUARDANDO 1ª CORRIDA | 127,5 | 0,376 |
| SEM CORRIDA DESTE ANALITO | 143,2 | 0,335 |

O "REJEITADO" da ocorrência está resolvido. O critério do UX ("cada status … redução ≤ 30 %; se não, a forma curta ou
quebra") não se cumpre nos três longos, que ficam ilegíveis (≈ 3,7 pt efetivos). O ADR-090 já os listou como pendência
de UX.

### OC-0049: linhas vazias da Estatística da Hema (UX G31) — NAO_VALIDADA pelo critério

36 linhas (98 a 133) sem analito: A, G, R, S e T vazias. A varredura de A..AU nessas linhas só achou texto em **B**
("1", "2", "3", constante da planilha). O critério do UX era "A98:T133 vazias". O que a ocorrência descrevia (o "0" e o
"SEM EP") não aparece mais.

### OC-0052, OC-0053, OC-0054: CEQ (reteste02)

- **OC-0052:**
  - O oráculo é o mapa lido da EQA_Base!W:Y: todos os nomes que são prefixo literal de outro, de canônico diferente,
    sem par de palavra.
  - Na Hema, 1 par (MCH → MCHC), citado no aviso.
  - Na Bio, 0 pares.
- **OC-0053:**
  - Linha 2 do CAP (Bio: ALT (SGPT) → ALT (TGP); Hema: Basophils → BASO%).
  - Z3: "base em dia" → "⚠ FÓRMULA SOBRESCRITA em Bias (%) ou |Bias| (%): o bias usaria o valor digitado — restaure a
    fórmula da coluna" → "base em dia".
  - O G do canônico não mudou: Bio 6,185176166653275; Hema 2,5085549222470034.
  - Não executado: Consolidar com a fórmula ainda sobrescrita (o ADR diz que o valor digitado seguiria para o bias; o
    aviso e a trilha são a proteção aceita).
- **OC-0054:**
  - A célula de destino não estava selecionada (seleção em AB1, vinda de outra aba).
  - Mesmo valor: nenhuma linha no Audit_Log.
  - Valor diferente: uma CEQ_ALTERADO com o antes da EQA_Base, e o Z3 passa a "HÁ EDIÇÃO NÃO CONSOLIDADA" (pendência
    real).

### ADR-091: material por lote

**Bioquímica, tblCQ_Final da cópia:** 35.466 linhas. Os status são idênticos ao `efeito_adr091/Bioquimica.json`:
ATIVO 35.323, LOTE_FORA_DO_CQ 59, IMPLAUSIVEL 43, LOTE_OUTRO_MATERIAL 30, INATIVADO 11.

| Regra | Lote | Linhas | O que a cópia tem |
|---|---|---|---|
| SO_ANALITOS PCR | 8353 | 996 | PCR: 965 ATIVO (todos participam) + 1 IMPLAUSIVEL (BIO-224563, 0,0); outros 29 analitos: 30 LOTE_OUTRO_MATERIAL, 0 participam |
| FORA_DO_CQ | 8177 / 8794 | 29 / 30 | todos LOTE_FORA_DO_CQ, 0 participam |
| FORA_DO_CQ | 7974, 8975, 8773, 8970 | 0 | sem dados |
| MATRIZ URINA | 5657 / 5658 / 5659 | 7 / 15 / 155 | todos URINA; UCFP, GLUC; 0 em analito sérico |
| MATRIZ URINA | 6353 / 6354 | 44 / 860 | todos URINA; CRE2, MALB, PHOS, UCFP, NA, K, CL, BUN; 0 em analito sérico |
| MATRIZ URINA | 5653, 5654, 5656, 6351, 6352 | 0 | sem dados |

- **Colunas derivadas dos 89 LOTE_*:** PARTICIPA "NÃO", REGISTRAR_RESULTADO_NO_LJ "NÃO", NAO_PLOTAR e
  MOTIVO_EXCLUSAO_AUTOMATICA com o texto fixo. Todos têm RUN (contam na corrida).
- **Precedência:** os 13 resultados ≤ 0 do 8353 em outros analitos estão LOTE_OUTRO_MATERIAL, não IMPLAUSIVEL. O manual
  −1 no 8794 (MAN de teste) ficou LOTE_FORA_DO_CQ.
- **De/para:** nas linhas de urina, o ANALITO é o da origem, porque a URINA é "sem cadastro" na Cfg_Integracao. 0
  divergências.
- **Cenário original da OC-0007:** 0 resultados dos lotes 5659, 6353, 6354, 8353 (fora a PCR), 8177 e 8794 participando
  em analito do catálogo.
- **QA:**
  - A13: 89 ALERTAs, um por ID LOTE_* (89/89 IDs cobertos);
  - A14: 0;
  - `LotesMaterialResumo`: "MATRIZ=10;FORA_DO_CQ=6;SO_ANALITOS=1|" (sem aviso).
- **Manuais de teste:**
  - 8353 + Glicose → LOTE_OUTRO_MATERIAL;
  - 8353 + PCR → ATIVO;
  - 8794 + −1 → LOTE_FORA_DO_CQ;
  - 6354 + Creatinina: ver **N-01**.
- **Equivalência VBA × PQ** (`estudos/equivalencia_cq.py --fixtures --n 1`, cópia nova, refresh do PQ incluído):
  - 35.503 linhas, 0 colunas com divergência;
  - ADR-082 OK, ADR-088 OK;
  - ADR-091 OK nos lotes livres das fixtures (SO 8589, FORA 8590, MATRIZ 8588);
  - LOTE_FORA_DO_CQ 505 e LOTE_OUTRO_MATERIAL 579 no PQ;
  - 739 s.

**Hematologia:** tblLotesMaterial presente na Cfg_Integracao com as 5 colunas, sem linhas. "MATRIZ=0;FORA_DO_CQ=0;
SO_ANALITOS=0|"; 0 LOTE_*; A13 = A14 = 0; status ATIVO 47.399, DUPLICIDADE_ORIGEM 589, SEM_VALOR 93 (= efeito_adr091).

**Ampliação (lotes fora da tabela, para o laboratório decidir, não é defeito do software):**
- **PCR participando em lotes que não são da família 835x:**
  - 8974 (2; é o lote do painel de soro);
  - 8553 (2), 8383 (2), 8335 (3), 5352 (1);
  - 8352 (144) e 8356 (3) são da família 835x.
- **HbA1c:** 8359 (2).
- **Nenhum outro "painel rodado uma vez":** nenhum lote com 8 ou mais analitos e no máximo 4 resultados por analito
  fora da tabela.

## 3. Ocorrências novas

| Id | Prior. | Produto | Título | Evidência | Causa provável (não editei) |
|---|---|---|---|---|---|
| N-01 | **Alto** | Bio | Manual digitado num lote de URINA com o nome do analito de SORO entra na estatística do soro | Manual MAN de teste: lote 6354, analito "Creatinina", 12,345, N1, DIMENSION 1. Na tblCQ_Final ficou ATIVO, PARTICIPA SIM, ANALITO "Creatinina", MATRIZ URINA. Com o Painel em Creatinina · 6354 · Tudo, engPainel N1 n = 1, média 12,345, J7 "REJEITADO"; a Estatística mostra Creatinina N1 n = 1. As listas da aba aceitam a combinação (`lstLotes`, `lstAnalitos`: Validation.Value = True), e o lote 6354 tem bloco de Creatinina na aba Analitos (A39 "alvo N1 73,990 ± 0,0909"). Assim, quem digita à mão o controle de urina escolhe "Creatinina", porque é o único da lista | `src/comum/mEtlCQ.bas:1005` aplica a MATRIZ do lote ao manual. Em `:1019`, `ParaOrigem(DP, nD, "URINA", "Creatinina")` não acha linha (a URINA não tem ANALITO_QCINI), e o manual fica com o analito de soro. O motor filtra por ANALITO + LOTE, sem MATRIZ. É o defeito da OC-0007 pelo caminho manual, e hoje está latente (0 manuais na produção). Decisão de regra: CQ |
| N-02 | Médio | Ambos | CorrigirCEQ e ExcluirCEQ gravam sem sessão | Depois do `Logout`, SessaoAtiva False: SalvarCEQ "ERRO\|0", mas `CorrigirCEQ(547, …115)` deu "OK\|547\|" (F 110 → 115, CEQ_ALTERADO e consolidação) e `ExcluirCEQ('CAP', 547)` deu "OK\|547\|" (linha apagada, CEQ_EXCLUIDO). Na Hema, o mesmo com a linha 467 | `src/comum/mCEQForm.bas:409` (CorrigirCEQ) e `:542` (ExcluirCEQ) não chamam `mSeguranca.SessaoAtiva`; só o `SalvarCEQ` (`:402`) chama. Pela janela, isso só acontece se a sessão expirar com o frmCEQ aberto, mas o texto "Sessão encerrada" do SalvarCEQ foi feito exatamente para esse caso |
| N-03 | Baixo | Ambos | Audit_Log grava RESULTADO_MANUAL_INCLUIDO para manual que não entrou na estatística (IMPLAUSIVEL, LOTE_FORA_DO_CQ, LOTE_OUTRO_MATERIAL) | Bio: manuais −1 (IMPLAUSIVEL), 8794 −1 (LOTE_FORA_DO_CQ) e 8353 Glicose (LOTE_OUTRO_MATERIAL), todos com RESULTADO_MANUAL_INCLUIDO (o detalhe traz "MANUAL:IMPLAUSIVEL" etc.). Hema: manual −1, a mesma coisa | `src/comum/mIntegracao.bas:557-561` (`AcaoManualNovo`): só MANUAL_INCOMPLETO e CONFLITO_MANUAL têm ação própria. É o padrão da OC-0023 nos status novos |
| N-04 | Baixo | Ambos | Status longos cortados em J7:J9 | §2 OC-0043: fatores 0,33–0,38 com Reduzir para caber | Forma curta dos status (decisão de UX, como o ADR-090 registrou) |
| N-05 | Baixo | Bio | Registro: o ADR-090 chama de OC-0040 a correção do G23 (OC-0041) | ADR-090.md:20 e :36 × OCORRENCIAS.md (OC-0040 = G22, OC-0041 = G23) | Documentação. O G22 continua aberto (ver OC-0040) |

**Testes propostos:**
- N-01 → `qa_adr091.py`, caso **T07**: manual em lote MATRIZ com analito do catálogo de outra matriz. Esperado: não
  participa (ou E/A no QA).
- N-02 → `qa_adr090.py`, caso **O33b**: Corrigir/Excluir sem sessão devolvem "ERRO|0" e nada muda.
- N-03 → `qa_adr090.py`, caso **O23b**: ação da trilha por status (IMPLAUSIVEL e LOTE_*).

**Observações (não são OC):**
- O manual comum sem MOTIVO entra ATIVO, embora a instrução A3 diga "Preencha … MOTIVO". O MOTIVO só é exigido com
  SUBSTITUI_ID (ADR-082). A decisão sobre isso é do CQ/UX, e se liga à OC-0026.
- A validação do EQUIPAMENTO da Bio recusa "DIMENSION 1" escrito pelo COM, como já se viu no ciclo 01.
- A trilha do Audit_Log grava o nome, o login do Windows e a máquina de quem roda. É assim por desenho (ISO 15189), mas
  quem copia evidência para relatório versionado precisa omitir esses dados. Neste relatório eles foram omitidos.

## 4. Regressão (cópias novas da produção atual, uma por suíte, `regressao_reteste03.py`)

| Suíte | Bio | Tempo | Hema | Tempo | Entrega 38 (10/11) |
|---|---|---|---|---|---|
| qa_final | **27 PASS / 2 FAIL** (T15, T19) | 1.203 s | **29 / 0** | 1.809 s | 29/0 · 29/0 |
| qa_rodadas_cv | **12 / 0** | 125 s | HEMA_RODADAS | — | 12/0 · 12/0 |
| qa_adr090 | **22 / 0** | 257 s | HEMA_ADR090 | — | 22/0 · 23/0 |
| qa_adr091 | **7 / 0** | 272 s | HEMA_ADR091 | — | 7/0 · 7/0 |
| equivalencia_cq --fixtures | **OK**: 35.503 linhas, 0 divergências; ADR-082/088/091 OK | 739 s | HEMA_EQUIV | — | Bio 34.291 / Hema 44.430, 0 div. |

**qa_final Bio, T15 e T19: FAIL de ambiente, não de produto (reprovados, não aprovados).**
- **Causa:** o MODO_FONTE é o DB_SEAC de produção, e hoje (dia útil) chegaram resultados novos durante a suíte. A
  tblCQ_Final tinha 35.466 linhas na cópia, 35.503 no T02, 35.620 no T14 e 35.622 → 35.625 entre a 1ª e a 4ª
  atualização do T19.
- **T19:** compara contagem e IDs entre 4 ATUALIZAR, e os 3 resultados que chegaram no meio derrubam a igualdade.
- **T15:**
  - A suíte fixa o De/Até do Painel no início (G4 = último dia com dado = 10/11, `qa_final.py:469`). Depois, coloca o
    manual às 23:59 do último dia com dado **depois** das atualizações, que já é 11/11.
  - O manual ficou correto (ATIVO, PARTICIPA SIM, RUN 26111102), mas fora da janela do Painel (`no_LJ: false`).
- **Mesmo padrão registrado antes:** a Bio barrada pelo qa_sync em 29/10 ("dados chegaram ao DB_SEAC no meio do teste").
- **Sem defeito do código:** a minha validação e as suítes do ADR-090/091 cobrem o caminho do manual sem defeito de
  código.
- **Para fechar:** repetir o qa_final da Bio fora do horário de chegada de dados. Proposta para a suíte (não editei):
  recalcular o G4 depois do último ATUALIZAR no T15, e no T19 comparar só os IDs que existiam antes.

O R00 do qa_rodadas_cv recebeu como "original" outra cópia da mesma produção (a correção já estava entregue). Por isso o
R00 aqui é uma auto-comparação e só prova que a suíte não altera nada; os R01..R12 são os que contam.

Produção conferida no fim da regressão: Bio `8b1cda1345d2` e Hema `f1907656b663`, iguais.
