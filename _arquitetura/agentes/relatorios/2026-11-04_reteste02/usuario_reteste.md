# Reteste 02: Usuário Simulado (ADR-089)

Data: 04/11/2026 · Agente: AGENTE_USUARIO_SIMULADO · Independente do Tester (não usei `valida_adr089.py` nem `repro_*`).

## Ambiente e método

| Item | Valor |
|---|---|
| Produção copiada (nunca aberta) | `QC_Bioquimica.xlsm` **161134b2f49d**, `QC_Hematologia.xlsm` **f3f225776701**, iguais antes e depois de cada etapa |
| Cópias | `%TEMP%\qcw\auditoria\2026-11-04_reteste02\usuario\<produto>\` (`<produto>_ceq.xlsm` e `<produto>_h3.xlsm`, cópias novas; nenhuma salva; retratos `S_*.xlsm` gerados por SaveCopyAs com os eventos desligados) |
| DB_SEAC | o congelado do reteste 01 (cópia de arquivo), apontado só nas cópias. O 1º ATUALIZAR trouxe 448 resultados novos na Bio (70 s) e 2.542 na Hema (71 s) |
| Excel | um por vez, em sequência. PIDs 27432, 11272 e 8588 (Bio) e 15252 e 4312 (Hema), todos encerrados pelo `fechar`. Nenhum Excel de outro processo foi tocado. Nenhuma captura de tela |
| Runner | `_arquitetura/agentes/ferramentas/reteste02_ceq.py <produto> [ceq,h3,calc]` (novo, documentado no cabeçalho). Ele refaz o fluxo do `f2_ceq.py`/`f2_ceq_hema.py` e do `reteste01_painel_h3.py` |
| Usuário | `TA_R02_ANALISTA`, papel **ANALISTA**, com login pelo `mSeguranca.DoLogin`. A etapa h3 roda **sem login** (sem UserInterfaceOnly), como no reteste 01 |
| Alvos | Bio: Glicose ("Glucose, serum", CAP, Ano EP 2025), com Ureia de controle. Hema: LYMPH# ("Lymphs Absolute", x 10E9/L, CAP), com LYMPH% de controle. São os mesmos do ciclo 01 |
| Dados de teste | rodada `TESTE-AUDITORIA R1`/`R9`, amostras TA-01..04, arquivo `TESTE-AUDITORIA.pdf`, resultados 110/105/120/150 com alvo 100. Tudo foi excluído no fim, e a Estatística voltou bit a bit ao estado inicial |

**Caminho usado:** chamei o núcleo que os botões da janela chamam (`mCEQForm.SalvarCEQ`, `CorrigirCEQ`, `ExcluirCEQ`, `AvisoAnalitoCEQ`, `ConfirmacoesCEQ`, `RegistrosDaRodada`, `LinhaCEQ`, `DescricaoCEQ`). Fiz a edição direta na tabela com os eventos ligados, selecionando a célula antes, como faz quem digita. O ATUALIZAR foi pelo núcleo do botão (`AtualizarDadosAutomatico`). Os filtros do Painel e o R4/M4 foram pelas células.

**Erro de script na 1ª passada da Bio, refeita depois:** na 1ª passada, o meu passo "Página sem seleção" escreveu "2" na coluna **P (|Bias| (%), fórmula)** em vez da Q. A passada foi refeita do zero, com a coluna certa, e é a que vale nas tabelas abaixo. Esse erro mostrou um risco real, registrado como R02-U1. A evidência ficou guardada em `run1_*`.

## O que não deu para fazer por automação

- **As caixas Sim/Não (`MsgBox`):** a do "Excluir o registro …?" (padrão Não) e a do "Gravar mesmo assim?" (fora do mapa e unidade diferente). Confirmei por leitura de `frmCEQ` (linhas 116 e 265, `vbDefaultButton2`) que o padrão é **Não**, e exercitei o que vem antes delas (`ConfirmacoesCEQ`, que devolve o texto da caixa) e depois do Sim (`SalvarCEQ`, `ExcluirCEQ`). Os cliques "Não" (nada gravado) e "Sim" na própria caixa **não foram executados**.
- **O estado da janela em execução:**
  - Corrigir e Excluir desabilitados sem seleção;
  - "Salvar correção" e "Cancelar correção";
  - provedor travado durante a correção;
  - o `lblMsg` azul ou âmbar.

  Do desenho (Designer), li o seguinte: `btnCorrigir` "Corrigir selecionado" (Top 420, Left 12), `btnExcluir` "Excluir selecionado" (Top 420, Left 138), `lstRegs` com 5 colunas e a 5ª de largura 0. As regras de habilitação vêm do código (linhas 56–57 e 102), mas não as vi rodando. Fica **NAO_EXECUTADO** para um teste manual com a janela aberta.
- **O botão "Consolidar" da EQA_Base:** não cliquei. A consolidação foi conferida pelo ATUALIZAR (o caminho novo), e a janela consolida sozinha.
- **OC-0051 com o Excel interativo:** a caixa "Nao foi possivel trocar o lote em analise" não aparece em automação (`Interactive = False`). Conferi o efeito: o modo TODOS entrou e não houve DESATUALIZADO.

## Resultado por ocorrência

| OC | Produto | Estado | Evidência |
|---|---|---|---|
| OC-0027 | Bio | **RESOLVIDA** | Logado como ANALISTA. Registro → `OK\|547\|`. A lista da rodada traz a linha na 5ª coluna (`[…, 547]`), e o `LinhaCEQ` devolve os 16 campos. A duplicidade diz: "Já existe este registro (CAP TESTE-AUDITORIA R1 2025 · Glucose, serum · TA-01) na linha 547. Para corrigir, selecione-o em “Registros desta rodada” e clique em Corrigir selecionado." `CorrigirCEQ(547, 105)` → `OK\|547\|`, na **mesma** linha e sem duplicar (EQA_Base 546 linhas). Corrigir para a chave de um registro real (linha 2) é recusado com ERRO\|5. Excluir → `OK`. Excluir de novo → "A linha 547 não é um registro…". A linha vazia é reaproveitada (`OK\|547\|`). A Estatística volta **bit a bit** ao estado inicial (A14:AU97, 0 diferenças) |
| OC-0027 | Hema | **RESOLVIDA** | O mesmo fluxo no LYMPH#: `OK\|467\|`, correção na linha 467, recusa da chave real (linha 2, Basophils), 4 exclusões `OK`, reuso da linha 467, Estatística bit a bit igual ao início |
| OC-0028 | Bio | **RESOLVIDA** | Edição direta 105 → 120 (aba liberada, como o ADM faria). Z3 = "⚠ HÁ EDIÇÃO NÃO CONSOLIDADA — clique em Consolidar na EQA_Base", em negrito âmbar (cor 22428). Z4 (46330,88734) > AA1 (46330,88710). G continua em 3,29997 (ainda não consolidado). ATUALIZAR (22 s) → trilha "origem: ATUALIZAR DADOS (edição pendente) … alteradas 1". Z2 passa a "consolidado: 2026-11-04 21:17:58 …", Z3 volta a "base em dia" e G vai a **7,04997** = independente. Um 2º ATUALIZAR sem pendência não consolida de novo |
| OC-0028 | Hema | **RESOLVIDA** | Z3 âmbar com o mesmo texto. G do LYMPH# fica em 3,70290 até o ATUALIZAR (30 s) e depois vai a **7,45290** = independente. O LYMPH% não muda. Z2 é atualizado. O 2º ATUALIZAR não consolida |
| OC-0029 | Bio | **RESOLVIDA** (aviso, como decidido) | `AvisoAnalitoCEQ(CAP,"Glucose")` = AMBAR "⚠ “Glucose” não está no mapa do CAP: este resultado NÃO entra no bias, no ET nem no Sigma…". O Controllab ("Glucose, serum") recebe o mesmo AMBAR. A confirmação explica o mapa (EQA_Base W:Y). Depois do Sim, o retorno é `OK\|548\|Gravado fora do bias (analito não mapeado).` (CAP) e `OK\|2\|…` (Controllab). A Estatística da Glicose fica idêntica. O nome certo recebe AZUL "No QC_INI: Glicose…" e nenhuma confirmação |
| OC-0029 | Hema | **RESOLVIDA** | "Lymphs Abs" (CAP) e "Lymphs Absolute" (Controllab) recebem AMBAR, a confirmação e o OK com "Gravado fora do bias". O LYMPH# fica idêntico |
| OC-0030 | Hema | **RESOLVIDA** (aviso) | "Lymphs" → AZUL "No QC_INI: LYMPH%. … Atenção: existe também “Lymphs Absolute” (LYMPH#).". Com x 10E9/L, a confirmação diz: "A unidade “x 10E9/L” é diferente da já registrada para LYMPH% (“%”). Se o resultado é do LYMPH#, use o nome “Lymphs Absolute”.". Se o usuário responde **Sim** mesmo assim, o resultado entra no LYMPH% (G 3,629 → **15,222**, Sigma N1 4,42 → 1,74, igual ao ciclo 01, e o independente bate). Isso é o comportamento decidido no ADR (só aviso), e o OK final avisa "No bias do LYMPH%, com unidade diferente…" |
| OC-0030 | Bio (vizinho) | **RESOLVIDA** | "Lactate" → AZUL com o par "Lactate Dehydrogenase" (Lactato desidrogenase). U/L × mmol/L pede confirmação e sugere o nome |
| OC-0032 | Bio | **RESOLVIDA** | Todos os eventos aparecem no Audit_Log com o usuário e o papel (ANALISTA): CEQ_REGISTRADO; CEQ_ALTERADO "Seu resultado: 110 → 105 \| pela janela"; CEQ_ALTERADO da edição direta "Seu Resultado: 105 → 120 \| falta consolidar"; EQA_BASE_CONSOLIDADA com a origem (janela registro/correção/exclusão, ATUALIZAR), as linhas por provedor, incluídas/alteradas/removidas e o carimbo de antes e de depois; CEQ_EXCLUIDO com todos os valores apagados; PROVEDOR_ALTERADO "Painel M4 · analito Glicose \| provedor do CEQ: CAP → Controllab" e a volta. Ver R02-U2 |
| OC-0032 | Hema | **RESOLVIDA** | Os mesmos eventos. PROVEDOR_ALTERADO "Estatística R4 (vale para todos os analitos) \| provedor do CEQ: CAP → Controllab" e a volta, com o antes certo (R4 selecionado antes) |
| OC-0051 | Bio | **RESOLVIDA** | Mesma sequência do reteste 01, sem login: L4 DIMENSION 2 → C3 HbA1c → H3 8590 → K3 2026 → K4 Set. O resultado é N1 n 23 / média 4,97957 / CV 2,7136 e N2 28 / 9,5225 / 2,3760, iguais ao independente (tblCQ_Final do retrato: 23/4,979565/2,713600 e 28/9,5225/2,375969). Apagar o H3 → H4 "▲ H3 vazio = TODOS os lotes do período…", G6 "Bias % do método (CEQ)", I6 "Sigma do método", n **24/32** = independente (8590 + 8977: 23+1, 28+4), média/DP/CV vazios, **sem DESATUALIZADO**, Configuração!G2 = 8974. Vizinho: 8974 → 8977 (n 90/83, média 79,744/274,265, CV 26,43/2,077 = independente) → vazio (TODOS) → 8974 (volta certo) |
| OC-0051 | Hema | **RESOLVIDA** | C3 WBC → H3 602911 → K3 2026 → K4 Mar → apagar o H3: H4 TODOS, G6/I6 "do método", n 54/57/50, sem DESATUALIZADO. Vizinho: 536411 → 602911 → vazio → 536411, todos coerentes |

### Bias recalculado depois da correção (calculadora independente)

O `calc_independente.py` foi rodado sobre os retratos e recalcula do zero, a partir da EQA_Base e da tblCQ_Final (regra do ADR-088 dec. 6, conferida no `mCEQ.AmostraDoVies`). A tolerância é 1e-9 relativo.

| Analito | Equip. | Lote | Nível | Momento (retrato) | Entrada | Fórmula | Independente G | QC_INI G | Status |
|---|---|---|---|---|---|---|---|---|---|
| Glicose | DIMENSION 2 | 8977 | 1, 2 | S_corr (janela, 105) | 3 rodadas reais (média 2,733294) + R1 = \|5\| | média das médias por rodada | 3,2999708324736257 | 3,299970832473625 | OK |
| Glicose | DIMENSION 2 | 8977 | 1, 2 | S_atu (edição 120 + ATUALIZAR) | + R1 = \|20\| | idem | 7,049970832473625 | 7,049970832473625 | OK |
| Glicose | DIMENSION 2 | 8977 | 1, 2 | S_fim (excluído) | 3 rodadas reais | idem | 2,7332944432981674 | 2,733294443298167 | OK |
| LYMPH# | XN1000 | 522611 | 1–3 | S_corr / S_atu / S_fim | 3,270527 + R1 = 5 / 20 / — | idem | 3,70290 / 7,45290 / 3,27053 | iguais | OK |
| LYMPH% | XN1000 | 522611 | 1–3 | S_par ("Lymphs" com Sim) | + 150/100 | idem | 15,222013284930423 | 15,222013284930423 | OK (é o que o mapa manda) |

O resumo das linhas e colunas foi o seguinte:
- Bio: 62 linhas × 19 colunas, **0 DIF** em S_corr, S_atu e S_fim.
- Hema: 84 linhas × 27 colunas e Painel 45/45, **0 DIF** em S_corr, S_atu, S_par e S_fim.

O Sigma (L) acompanha: Glicose N1 0,3085 → 0,1358 → 0,3346, e o independente bate em todos.

## Impressão de usuário

- **Corrigir/Excluir (OC-0027):** a mensagem da duplicidade agora diz onde está o registro e qual botão usar, e isso resolve a confusão do ciclo 01. Os dois botões ficam na mesma linha do Salvar, à esquerda. Pelo desenho são fáceis de achar, mas não vi a janela aberta. Um ponto confunde: na descrição do registro o ano aparece duas vezes ("CAP C-A 2025 2025 · ALT…") quando o nome da rodada já traz o ano (R02-U4).
- **Edição direta + Z3 (OC-0028):** o aviso é claro e chama atenção (âmbar e negrito, "⚠ HÁ EDIÇÃO NÃO CONSOLIDADA"). Ele só aparece em Z3, que pode estar fora da área visível de quem digita na coluna A–R. Ele manda "Consolidar na EQA_Base", mas o ATUALIZAR também resolve e o texto não diz isso. Seria mais simples para o usuário: "clique em ATUALIZAR DADOS (ou Consolidar na EQA_Base)".
- **Fora do mapa (OC-0029):** o texto é claro e diz a consequência (fora do bias, do ET e do Sigma) e o que fazer. O "peça ao administrador para mapear" deixa o analista sem saída imediata, mas é o que foi decidido.
- **Lymphs (OC-0030):** o aviso azul com "existe também “Lymphs Absolute” (LYMPH#)" e a confirmação de unidade pegariam o meu erro do ciclo 01. Se eu responder Sim por hábito, o erro passa igual. Fica como risco residual aceito no ADR.
- **Trilha (OC-0032):** completa e legível, com antes → depois.
- **OC-0051:** o Painel agora entra no TODOS sem precisar de nada mais. Um ponto confunde: no modo TODOS, a A39 continua escrevendo "Lote 8974 · alvo N1 …" (o lote em uso) logo abaixo do H4 que diz "TODOS os lotes" (R02-U3).
- **Pendência conhecida (G17):** o GUIA §5–6 ainda não cita Corrigir/Excluir nem o aviso do Z3. Não verifiquei de novo, porque está fora deste lote.

## Ocorrências novas

**R02-U1: Sobrescrever a coluna |Bias| (%) (P, fórmula) na tabela do provedor muda o bias do ET/Sigma sem aviso nem trilha (Médio; `dados`)**
- **Tela/módulo:** EQA.CAP_Dados (edição direta com a aba liberada) → `mEQA` (D_BIASABS = 16 → EQA_Base Q) → `mCEQ.BiasEQ` (EQ_C_BIASABS).
- **Ação:** na 1ª passada da Bio, digitei "2" em P547 por engano (|Bias| de fórmula `=IF(O547="","",ABS(O547))`), junto com a edição do resultado para 120, e depois rodei o ATUALIZAR.
- **Esperado:** o bias vem do resultado e do alvo (ou a coluna de fórmula não aceita digitação), e a edição é marcada e registrada.
- **Observado:**
  - G da Glicose = **2,54997** (usou |bias| = 2), enquanto o T (com sinal, da coluna O) usou +20. O independente dá 7,04997: diferença relativa de 0,64 no G e de 0,60 no Sigma (0,343 × 0,136).
  - A edição de P não entra no Z4/Z3 e não gera CEQ_ALTERADO, porque só A..N, Q e R são vigiadas. A fórmula some da linha.
- **Consequência:** o bias, o ET e o Sigma ficam errados em silêncio, com Z3 "base em dia". Só acontece com a aba liberada (ADM).
- **Sugestão:** a consolidação recalcula |bias| a partir de O (ou de F/G) em vez de copiar P, ou trava O/P mesmo para o ADM. O Z3/Z4 também deveria vigiar O/P.
- **Evidência:** `Bioquimica\run1_ceq.json` (C3), `run1_S_atu.xlsm` (P547 = 2), `run1_S_atu_calc.json` (DIF G/H/L/N/O/V/W só nas linhas Glicose 38–39).

**R02-U2: A edição direta sem selecionar antes registra "? → valor" mesmo quando o valor não mudou, e marca a pendência (Baixo; `usabilidade`)**
- **Ação:** escrevi o mesmo "TA-01" na Amostra sem selecionar a célula.
- **Observado:** CEQ_ALTERADO "Amostra: ? → TA-01" e Z4 atualizado. Isso está previsto no ADR (riscos), mas gera trilha e pendência falsas.
- **Sugestão:** comparar com o valor da EQA_Base quando não houver o "antes".

**R02-U3: No modo TODOS (H3 vazio), a A39 mostra "Lote <em uso> · alvo …" logo abaixo do H4 "TODOS os lotes" (Baixo; `usabilidade`)**
- **Observado:** Bio A39 "Lote 8974 · alvo N1 5,07 …" e Hema "Lote 522611 · alvo …", com o H4 dizendo "cada ponto com o alvo do seu lote".
- **Sugestão:** no modo TODOS, escrever "Todos os lotes (alvo de cada lote)".

**R02-U4: A descrição do registro repete o ano ("CAP C-A 2025 2025 · ALT …") (Baixo; `usabilidade`)**
- **Onde aparece:** na mensagem de duplicidade, na confirmação da exclusão e no Audit_Log, quando a rodada já traz o ano.

Observação sem OC: os eventos de usuário do CEQ saem com Categoria "SISTEMA", mas o usuário e o papel estão preenchidos.

## Arquivos

- Saídas: `%TEMP%\qcw\auditoria\2026-11-04_reteste02\usuario\{Bioquimica,Hematologia}\`:
  - `ceq.json`/`ceq.log`;
  - `h3.json`/`h3.log`;
  - `calc.json`/`calc.log`, `S_*_calc.json`;
  - `trilha_bio.json`/`trilha_hema.json`;
  - na Bio, `run1_*` (a passada com o erro de coluna).
- Ferramenta nova: `_arquitetura/agentes/ferramentas/reteste02_ceq.py`.
