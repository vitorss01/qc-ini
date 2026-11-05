# Tester/QA, modo VALIDAR: ADR-089 (reteste02, 04/11/2026)

## 0. Resumo

| OC | Produto | Veredito | Evidência curta |
|---|---|---|---|
| OC-0027 | Bio | **VALIDADA** | 110 → G 5,0345; Corrigir 105 → G 3,7845022700255755 (igual ao ciclo 01); Excluir → G volta exatamente a 3,3793363600341007, 0 linhas da Estatística diferentes |
| OC-0027 | Hema | **VALIDADA** | RBC: 1,5336 → 3,6502 → 2,400215093957569 (critério do UX: 2,4002) → 1,533620125276759 exato |
| OC-0028 | Bio | **VALIDADA** | edição direta F547 → Z3 "⚠ HÁ EDIÇÃO NÃO CONSOLIDADA", G parado; ATUALIZAR completo (68 s) consolida, G = G da correção, Z2 = Z1, Z3 "base em dia" |
| OC-0028 | Hema | **VALIDADA** | idem, F467; ATUALIZAR 68 s; G 2,400215093957569 = G da correção |
| OC-0029 | Bio | **VALIDADA** (função) | "Protein, total", Controllab "Proteína total", "Hemoglobin A1c": AMBAR + confirmação + "Gravado fora do bias", 0 linhas mudadas |
| OC-0029 | Hema | **VALIDADA** (função) | "Red Blood Cell", Controllab "RBC", "nRBC": idem |
| OC-0030 | Hema | **VALIDADA** (função) | "Lymphs": aviso azul LYMPH% + "Lymphs Absolute" (LYMPH#); confirmação de unidade x 10E9/L × %; os 6 pares %/# iguais |
| OC-0030 | Bio (ampliação) | **VALIDADA** (função) | Lactate / Lactate Dehydrogenase: aviso do par e confirmação U/L × mmol/L |
| OC-0032 | Bio | **VALIDADA** | CEQ_REGISTRADO, CEQ_ALTERADO (janela e edição direta, antes → depois), CEQ_EXCLUIDO, EQA_BASE_CONSOLIDADA (janela, botão, ATUALIZAR), PROVEDOR_ALTERADO (M4, ida e volta); cadeia `OK|641` |
| OC-0032 | Hema | **VALIDADA** | idem; PROVEDOR_ALTERADO no R4 "vale para todos os analitos", ida e volta; cadeia `OK|249` |
| OC-0051 | Bio | **VALIDADA** | sem UserInterfaceOnly: HbA1c 8590 → Set → H3 vazio = TODOS (AD1 True, H4 "TODOS", G6 "do método"), E1 acompanha cada passo, nunca DESATUALIZADO |
| OC-0051 | Hema | **VALIDADA** | WBC 602911 → Mar → vazio, e 536411 → 602911 → vazio: TODOS, sem DESATUALIZADO, sem e com UIO |

Nenhum veredito NAO_VALIDADA. Quem corrigiu foi o agente programador (commits 85f5b85 e 2facec9). Eu não editei
`src/` nem instaladores. O script de validação é meu (`ferramentas/valida_adr089.py`) e não reaproveita o `qa_adr089.py`.

## 1. Ambiente

- **Produção:** entregue (entrega 36, commit 2facec9). Bio `161134b2f49d2517`, Hema `f3f2257767017380`. São os
  mesmos hashes do MANIFESTO e dos `.xlsm` em `_entrega_2026-11-04_1405/`. A produção foi só copiada (`shutil.copy2`)
  e continuou com o mesmo hash no fim dos dois scripts e da regressão.
- **Cópias:** `%TEMP%\qcw\auditoria\2026-11-04_reteste02\` (`<produto>_repro.xlsm`, e `regressao\<produto>\...`).
  Nada foi salvo.
- **Excel:** um por vez. PIDs Bio 16488 e Hema (impresso no log), encerrados pelo `qa_final.QA.fechar()`. Nenhuma
  captura de tela.
- **Dados de teste:** rodada `TESTE-AUDITORIA R2` / `TESTE-AUDITORIA CTL2`, ano 2025, amostras `TA-*`, arquivo
  `TESTE-AUDITORIA.pdf`, resultados 110 / 105 / 120 / 131, alvo 100.
- **Logs e JSON:** `valida_bio.log`, `valida_hema.log`, `valida_adr089_<produto>.json`, `auditoria_janelas.json`,
  `regressao.log` e `regressao\<produto>\*.log|json`, todos na pasta das cópias.
- **Estado da sessão:**
  - A 1ª parte (L51 e C1–C5) rodou com o arquivo recém-aberto, **sem** UserInterfaceOnly
    (`Configuração.ProtectionMode = False`). É o caso em que a OC-0051 falhava, e ele prova que `CorrigirCEQ` e
    `ExcluirCEQ` liberam a proteção sozinhos.
  - Depois, `mSeguranca.ReprotectAll` (o estado da sessão com login) para a edição direta, o ATUALIZAR e o provedor.
- **ATUALIZAR:** `mIntegracao.AtualizarDadosAutomatico`, o ATUALIZAR **completo**. Ele lê o DB_SEAC, sem gravar nada
  lá. Não usei `ConsolidarSePendente` direto, como faz o `qa_adr089` C06: assim o gancho do ADR fica provado no caminho
  real.

## 2. Por ocorrência

### OC-0027: Corrigir e Excluir pela janela (G5 do ciclo 01)

Passos iguais aos do G5/G6 do ciclo 01. Bio: Proteína total, "Protein, total, serum", g/dL, Painel em 8977 /
DIMENSION 2 / 2026 / Out. Hema: RBC, "Red Blood Cell Count", x 10E12/L.

| Passo | Bio (linha 547) | Hema (linha 467) |
|---|---|---|
| B0 (G, todos os níveis) | 3,3793363600341007 | 1,533620125276759 |
| `SalvarCEQ` 110 | `OK|547|`, G 5,034502270025575 | `OK|467|`, G 3,650215093957569 |
| `SalvarCEQ` 105, mesma chave | `ERRO|5|Já existe este registro (…) na linha 547. Para corrigir, selecione-o em “Registros desta rodada” e clique em Corrigir selecionado.` | idem, linha 467 |
| `RegistrosDaRodada` | 1 × 5, a 5ª coluna = 547 | 1 × 5, a 5ª = 467 |
| `CorrigirCEQ(547, …105)` | `OK|547|`, F547 = 105, EQA_Base 105, G **3,7845022700255755** | `OK|467|`, G **2,400215093957569** |
| `CorrigirCEQ` com resultado "abc" | `ERRO|6|Seu resultado: digite um número.`, F inalterado, sem trilha | idem |
| `ExcluirCEQ` | `OK|547|`, A/D/F/R vazios, 0 linhas TESTE na EQA_Base, G = B0 **exato**, 0 linhas da Estatística diferentes do B0 (G, H, L) | idem, G = B0 exato |
| `ExcluirCEQ` de novo | `ERRO|0|A linha 547 não é um registro…` | idem |

- **Conferência com o ciclo 01:** os G de 110 e 105 são **idênticos**, nos 16 dígitos, aos que o ciclo 01 obteve pelo
  caminho editar + Consolidar (Bio 5,0345… / 3,7845…, Hema 2,4002). O critério do UX G09 (Hema 2,4002) foi atendido.
  O valor da Glicose Bio (3,30) citado pelo UX é de outro analito e não foi medido: usei a Proteína total do meu G5.
- **Permissão:** `CorrigirCEQ` e `ExcluirCEQ` gravaram numa aba protegida **sem** UserInterfaceOnly. Logo, a
  aba bloqueada não impede mais o ANALISTA.
- **`auditoria_janelas.py`:** frmCEQ OK, frmInativar OK, frmCarregando OK (rc 0). Os botões "Corrigir selecionado" e
  "Excluir selecionado" existem na especificação (`ceq_form_spec.py` 99–103). A habilitação deles (`BotoesSelecao`,
  `frmCEQ.frm.txt` 53–58) eu conferi só no código.
- **Só no nível da função:** a caixa Sim/Não do Excluir (padrão Não, `frmCEQ.frm.txt` 116–120) e o clique nos botões.
  Pelo código, o Não sai antes do `ExcluirCEQ`.

### OC-0028: edição direta, Z3, ATUALIZAR (G6 do ciclo 01)

| Passo | Bio | Hema |
|---|---|---|
| `SalvarCEQ` 110 (reaproveita a linha excluída) | `OK|547|` (a mesma linha) | `OK|467|` |
| `F.Select` + `F = 105` (eventos ligados) | Z3 = "⚠ HÁ EDIÇÃO NÃO CONSOLIDADA — clique em Consolidar na EQA_Base", em negrito; Z4 46330,82996 > AA1 46330,82979; `ConsolidacaoPendente` True; **G parado** (5,0345…) | idem; Z3 em negrito, fonte 22428 = RGB(156,87,0), âmbar |
| ATUALIZAR completo | `OK|Atualização concluída…`, 68,0 s, 618 resultados novos do DB_SEAC | 67,9 s, 2852 novos |
| Depois | G **3,7845022700255755** = G da correção pela janela; Z2 = Z1 (carimbo novo); Z3 "base em dia"; sem pendência | G **2,400215093957569** = G da correção; Z2 = Z1; "base em dia" |
| Trilha do ATUALIZAR | `EQA_BASE_CONSOLIDADA` "origem: ATUALIZAR DADOS (edição pendente)" | idem |
| `ConsolidarSePendente` sem pendência | False, nenhuma linha de trilha | idem |

- **Ampliação (Controllab):** registro Controllab, F = 131 editado direto → Z3 da EQA.Controllab_Dados avisa, e a
  pendência fica True. O botão Consolidar (`mEQA.AtualizarEQABase`) leva a "base em dia", Z2 = Z1, com trilha
  "origem: botão Consolidar". Bio e Hema iguais.
- **Diferença do critério do UX G10:** o UX pedia a consolidação ao **sair da aba**. O ADR-089 decidiu não fazer isso
  e deixou o aviso no Z3 e o ATUALIZAR. Validei o critério do Orquestrador (Z3 → ATUALIZAR → bias).

### OC-0029: fora do mapa (G7 do ciclo 01)

| Nome | Aviso | `ConfirmacoesCEQ` | `SalvarCEQ` | Estatística |
|---|---|---|---|---|
| Bio CAP "Protein, total" | `AMBAR|⚠ “Protein, total” não está no mapa do CAP: este resultado NÃO entra no bias…` | "O analito … não está no mapa do CAP. O registro será gravado, mas NÃO entra…" | `OK|548|Gravado fora do bias (analito não mapeado).` | 0 linhas mudadas |
| Bio Controllab "Proteína total" | AMBAR "…mapa do Controllab" | idem | `OK|2|Gravado fora do bias…` | 0 |
| Bio CAP "Hemoglobin A1c" (analito novo) | AMBAR | idem | `OK|549|Gravado fora do bias…` | 0 |
| Hema CAP "Red Blood Cell" | AMBAR | idem | `OK|468|Gravado fora do bias…` | 0 |
| Hema Controllab "RBC" | AMBAR | idem | `OK|2|Gravado fora do bias…` | 0 |
| Hema CAP "nRBC" | AMBAR | idem | `OK|469|Gravado fora do bias…` | 0 |
| Nome da lista (Bio "Protein, total, serum", Hema "Red Blood Cell Count") | `AZUL|No QC_INI: <canônico>. Este resultado entra no bias deste analito.` | "" (sem falso positivo) | — | — |

- **Só no nível da função:** a caixa "Gravar mesmo assim?" (padrão Não) e a cor do `lblMsg`. Pelo código
  (`frmCEQ.frm.txt` 263–270), o Não sai com "Nada foi gravado" antes do `SalvarCEQ`. O "Não → nada gravado" do
  critério do UX não foi clicado.
- **Bio M4 = Controllab:** a Proteína total fica "SEM EP" mesmo com o registro Controllab gravado. Isso é coerente com
  "fora do bias" (o Controllab não tem mapa; G11 item 4 continua pendente por decisão).

### OC-0030: pares %/# (G8 do ciclo 01)

- **"Lymphs", x 10E9/L, 110:**
  - aviso `AZUL|No QC_INI: LYMPH%. Este resultado entra no bias deste analito. Atenção: existe também “Lymphs Absolute” (LYMPH#).`;
  - confirmação `A unidade “x 10E9/L” é diferente da já registrada para LYMPH% (“%”). Se o resultado é do LYMPH#, use o nome “Lymphs Absolute”.`;
  - quem confirma Sim recebe `OK|470|No bias do LYMPH%, com unidade diferente da já registrada (“%”).`, e o LYMPH%
    muda nos 3 níveis. É o comportamento decidido no ADR (só aviso).
- **Varredura independente dos pares:** calculei os pares a partir do W:Y da cópia, com a unidade mais usada de cada
  canônico na tabela do CAP.
  - Na Hema, os 6 pares %/# (Basophils, Eosinophils, IG, Lymphs, Monocytes, Neut/Gran) dão aviso com o par,
    confirmação com a unidade do par e **nenhuma** confirmação com a unidade própria (%).
  - O nRBC não é de prefixo e cai no âmbar (ver OC-0029).
- **Ampliação, Bio:** Lactate (mmol/L) × Lactate Dehydrogenase (U/L). Aviso do par, confirmação com U/L e nada com
  mmol/L.
- **Só no nível da função:** o Não da confirmação ("nada muda no LYMPH%", critério do UX G12), pelo mesmo caminho do
  código da OC-0029.

### OC-0032: trilha no Audit_Log

Cada linha tem carimbo, Windows (usuário e máquina), arquivo, versão e hash encadeado.
`mAuditoria.VerificarIntegridadeLog` dá `OK|641` (Bio) e `OK|249` (Hema) no fim.

| Evento | Exemplo (Bio) |
|---|---|
| `CEQ_ALTERADO` (janela) | "CAP / TESTE-AUDITORIA R2 / 2025 / Protein, total, serum / TA-11 / Seu resultado: 110 → 105 / pela janela / aba EQA.CAP_Dados linha 547" |
| `CEQ_ALTERADO` (edição direta) | "edição direta na tabela / aba EQA.CAP_Dados / linha 547 (…) · Seu Resultado: 110 → 105 / falta consolidar…" (o antes veio do SelectionChange) |
| `CEQ_EXCLUIDO` | "valores apagados: Seu resultado 105; Valor-alvo 100; SD 2; SDI 5; …; Arquivo fonte TESTE-AUDITORIA.pdf / pela janela / linha 547" |
| `EQA_BASE_CONSOLIDADA` | origem (janela registro/correção/exclusão, botão, ATUALIZAR), linhas por provedor, incluídas/alteradas/removidas, carimbo antes e depois |
| `PROVEDOR_ALTERADO` | Bio "Painel M4 · analito Proteína total / CAP → Controllab", e a volta. Hema "Estatística R4 (vale para todos os analitos) / CAP → Controllab", e a volta |

- **Usuário:** na automação, os campos do login saem "(sem login) / (sem papel)", e o usuário do Windows é gravado.
  O usuário do login **não foi verificado** (`NAO_EXECUTADO`: exige o login interativo; fica para o Usuário Simulado).

### OC-0051: H3 vazio = TODOS

Os critérios, a cada passo: E1 = H3 enquanto há lote; nenhum J7:J(6+n) "DESATUALIZADO"; no fim, AD1 = True, H4 com
"TODOS" e G6 "Bias % do método (CEQ)".

| Sequência | Sem UIO | Com UIO |
|---|---|---|
| Bio: L4 DIMENSION 2, C3 HbA1c, H3 8590, K3 2026, K4 Set, apagar H3 | OK (E1 8590 → 8590 → TODOS) | OK |
| Bio: C3 Glicose, L4 DIMENSION 1, H3 8974 → 8977, K3/K4 Out, apagar H3 | OK | — |
| Hema: C3 WBC, H3 602911, K3 2026, K4 Mar, apagar H3 | OK | OK |
| Hema: 536411 → 602911 → Mar → apagar H3 | OK | — |

- O `Configuração!G2` (loteParam) acompanhou cada troca: 8590 → 8974 na Bio, e 602911 → 522611 na Hema, no TODOS.
- O J7 mostra OK/REJEITADO, o status de Westgard da tela, e nunca DESATUALIZADO.
- No 8974 (sem resultado em Out), o Painel ajustou o período para "Tudo" e gravou `PAINEL_FILTRO_AJUSTADO`. É um
  comportamento anterior, não deste ADR.

## 3. Regressão

- **Completa, na entrega:** `entregar.py`, `_entrega_2026-11-04_1405/`, sobre o mesmo binário (os hashes conferem).
  As 29 suítes passaram nos dois produtos, inclusive `qa_sync` (S6 do PQ: 121 s Bio, 235 s Hema, dentro do histórico
  de 170–1130 s ou abaixo dele), `qa_desempenho` 5/0 e `qa_final` 29/0. Quem rodou foi a entrega, não eu.
- **Repetida por mim, nas suítes afetadas,** sobre cópias novas da produção atual (`regressao.log`):

| Suíte | Bio | Hema |
|---|---|---|
| `qa_adr087` (CEQ, janelas) | 14 PASS / 0 FAIL | 14 / 0 |
| `qa_adr086` (H3 vazio) | 11 / 0 | 11 / 0 |
| `qa_rodadas_cv` (EQA_Base, R00) | 12 / 0 | 12 / 0 |
| `qa_troca_lote` (mLotes) | 3 / 0 | 3 / 0 |
| `qa_final` (ATUALIZAR completo) | 29 / 0 (1100 s) | 29 / 0 (1223 s) |
| `auditoria_janelas.py` | OK (os 3 formulários) | — |

A produção ficou com o mesmo hash antes e depois (Bio `161134b2f49d2517`, Hema `f3f2257767017380`). Os tempos do `qa_final` (18–20 min) estão na faixa da entrega (cerca de 18 min no log de `_entrega_2026-11-04_1405`).

## 4. Efeitos colaterais e observações (nenhum bloqueia)

1. **MCH / MCHC (Hema, Baixo):**
   - "MCH" é prefixo literal de "MCHC", com outro canônico. O aviso azul não cita o par, porque o `ParDoMapa` exige
     nome + espaço (`src/comum/mCEQForm.bas` 255–265), e o ADR-089 §4 diz só "prefixo".
   - A confirmação de unidade pg × g/dL dispara, sem sugerir o nome.
   - Não é um par %/#, então o comportamento é razoável. O texto do ADR é que deveria dizer "prefixo de palavra".
2. **Rótulo do campo na trilha (Baixo):** a janela grava "Seu resultado", e a edição direta grava "Seu Resultado", o
   cabeçalho da tabela. Filtrar o Audit_Log pelo nome do campo exige os dois.
3. **Controllab sempre fora do bias:** é o esperado enquanto não existir o mapa do Controllab (G11 item 4, fora do
   lote). Hoje todo CEQ Controllab recebe o aviso âmbar e a confirmação.
4. **Unidade divergente com Sim:** o registro entra no bias do canônico errado, por decisão ("só aviso"). Fica
   registrado no OK e no Audit_Log, mas não há bloqueio.
5. **Fluxos vizinhos conferidos sem regressão:**
   - registro novo e duplicidade;
   - reaproveitamento da linha excluída (547 / 467);
   - validação numérica na correção;
   - exclusão repetida;
   - ATUALIZAR sem pendência (não consolida de novo);
   - troca de provedor de ida e volta (G volta igual);
   - cadeia de hash do Audit_Log íntegra;
   - Painel em lote único, nos dois estados de proteção.
6. **Tempo:** o ATUALIZAR completo levou 68 s nos dois produtos, com a consolidação pendente incluída.

## 5. Teste não executado

- **Os cliques nas caixas** (Sim/Não do Excluir e do "Gravar mesmo assim?") e a cor real do `lblMsg` na janela aberta:
  `NAO_EXECUTADO` por automação, porque o `MsgBox` é modal. Ficam com o Usuário Simulado. O núcleo (`ConfirmacoesCEQ`,
  `CorrigirCEQ`, `ExcluirCEQ`, `AvisoAnalitoCEQ`) foi exercitado, e o desvio do Não foi conferido no código.
- **O usuário do login na trilha:** `NAO_EXECUTADO` (exige login interativo).
