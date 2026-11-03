# Tester/QA, modo REPRODUZIR: ciclo 2026-10-31_ciclo01

Agente: Tester/QA. Data: 03/11/2026, das 07h10 às 08h00. Entrada: as 26 ocorrências críticas e altas de
`rota_tester.json`, com a evidência dos relatórios `<produto>_F1..F4.md`.

## 0. Resumo

As 26 ocorrências são **14 falhas distintas** (G1 a G14). Reproduzi cada uma uma vez por produto, com scripts próprios.
Não reaproveitei nenhum script do Usuário Simulado.

| Veredito | Ocorrências |
|---|---|
| REPRODUZIDA | 25 |
| PARCIAL | 1: Bio F3 "outro nível e valores ≤ 0". O Sigma da Glicose e os 253 suspeitos foram reproduzidos; a u(Rw) de 81,2 % da Proteína total não foi reconferida |
| NAO_REPRODUZIDA | 0 |

Os números das cópias de 03/11 batem **bit a bit** com os dos agentes em todos os casos em que a base é a mesma: CEQ
da Bio e da Hema, Sigma da Glicose, bias da Bilirrubina direta, alvos do lote 522611, "Tudo" no lote 8974, troca de lote
da Hema e ET do Painel com bias em "SEM EP".

Nos manuais (G1, G2 da Hema, G3), a produção ainda não tem os resultados de 29–30/10 que a F1 recebeu de um DB_SEAC
congelado. Por isso usei os dias mais recentes que a produção tem. Os números mudam, mas o defeito é o mesmo e foi
reconstruído pelo cálculo independente com diferença relativa ≤ 1e-15.

## 1. Ambiente e método

| Item | Valor |
|---|---|
| Produção (só copiada, nunca aberta) | `QC_Bioquimica.xlsm` sha `fe10dba1f1fd`; `QC_Hematologia.xlsm` sha `0b45fa270d2a`. São os mesmos hashes do ciclo e continuaram iguais no fim |
| Cópias | `%TEMP%\qcw\auditoria\2026-10-31_ciclo01\tester\`: `Bio_dados.xlsm` e `Hema_dados.xlsm` (openpyxl); `Bioquimica_repro.xlsm` e `Hematologia_repro.xlsm` (Excel; cópia nova a cada sessão). Nenhuma foi salva: o hash ficou igual antes e depois em todas. Todas foram apagadas no fim |
| Excel | Um por vez, nenhum outro agente usando Excel. PIDs criados e encerrados pelo próprio script: 28976 (Bio, parou na seleção de dias), **20440** (Bio, manuais/Painel), **26244** (Bio, CEQ), 5456 (Hema, tentativa 1: só a G13), **15792** (Hema, sessão completa). Nenhum outro processo foi tocado. A tela não foi capturada |
| Scripts (novos, do Tester) | `_arquitetura/agentes/ferramentas/repro_dados.py` (openpyxl: G2, G10, G12, G14, fórmulas da G9, proteção da G5); `repro_comum.py` (utilitários sobre `qa_final.QA`); `repro_excel_bio.py` (G1, G3, G4, G9, G11, G5–G7, ampliação da G12); `repro_excel_hema.py` (G13, G1, G2, G3, G9, G5–G8, ampliação da G14) |
| Logs e JSON | `%TEMP%\qcw\auditoria\2026-10-31_ciclo01\tester\`: `repro_dados.json` e `.log`; `repro_excel_bio_manual.json` e `.log`; `repro_excel_bio_ceq.json` e `.log`; `repro_excel_hema.json` e `.log`; `repro_excel_hema_tentativa1.json` e `.log` |
| Dados de teste | Manuais com MOTIVO `Outro: TESTE-AUDITORIA` e valores sintéticos: Bio 77,111 e 244,222; Hema 17,433, 3,111, −1 e 0. CEQ com rodada `TESTE-AUDITORIA R1` / `TESTE-AUDITORIA CTL1`, amostras `TA-01..05` e `TA-C1`, arquivo `TESTE-AUDITORIA.pdf`. Só existiram nas cópias |

**Caminho do usuário**

- **Digitar Resultados:** escrevi célula a célula com `FormulaLocal` e os eventos ligados. O `Worksheet_Change` carimbou
  o ID, o EQUIPAMENTO padrão e a MATRIZ.
- **ATUALIZAR:** rodei `mIntegracao.AtualizarSoEntradas`, o núcleo do ATUALIZAR **sem ler o DB_SEAC** (ADR-081). Ele
  aplica os manuais e roda o motor inteiro: ETL, QA, Estatística, Westgard e incerteza. Desvio declarado: a F1 usou o
  ATUALIZAR completo com um DB_SEAC congelado. Com este caminho, os dados novos de 29–30/10 não entram e o DB_SEAC de
  produção não é tocado.
- **Painel e Estatística:** mudei pela própria célula. As abas foram desprotegidas, um passo técnico.
- **CEQ:** usei `mCEQForm.SalvarCEQ`, o núcleo do botão "Salvar e próximo" da `frmCEQ`, e o botão "Consolidar"
  (`mEQA.AtualizarEQABase`).
- **Validação da tela:** li `Range.Validation.Value` numa linha livre, com os eventos desligados, e limpei a célula
  depois.

**Achado de método**

A 1ª sessão de CEQ da Bio parou num diálogo "Erro em tempo de execução 1004 (planilha protegida)". O Excel aberto por
automação não roda o `Workbook_Open`, e por isso as abas ficam protegidas **sem** `UserInterfaceOnly`. Nesse estado, o
`mEQA.AtualizarEQABase` (`src/comum/mEQA.bas` 217–263), que não tem tratamento de erro nem `LiberarEscrita`, para no
`ClearContents` da EQA_Base.

- Li o texto do diálogo pelas janelas do meu PID (texto, sem imagem) e encerrei a macro pelo botão "Fim".
- Na sessão seguinte, chamei `mSeguranca.ReprotectAll`, o mesmo estado de uma sessão de ANALISTA depois do login
  (`UnlockApp`).
- Para o usuário logado isso não acontece, então **não** é ocorrência do produto. Fica a nota: o botão "Consolidar"
  depende do estado de proteção da sessão.
- Atenção às suítes que chamam o `SalvarCEQ` sem login: lá dentro, a falha da consolidação fica calada
  (`On Error Resume Next`, "aviso" no retorno).

## 2. Mapa ocorrência → grupo

| # | Ocorrência (rota_tester) | Produto | Grupo | Veredito |
|---|---|---|---|---|
| 1 | Minúsculas no EQUIPAMENTO/ANALITO do manual | Bio | G1 | REPRODUZIDA |
| 2 | Valores impossíveis do interfaceamento ATIVOS | Bio | G2 | REPRODUZIDA |
| 3 | Data 31/02/2026 vira 03/03/2026 | Bio | G3 | REPRODUZIDA |
| 4 | EQUIPAMENTO vazio recebe o do Painel | Bio | G4 | REPRODUZIDA |
| 5 | OC-1 F2: ANALISTA não corrige o CEQ | Bio | G5 | REPRODUZIDA |
| 6 | OC-2 F2: correção só depois do Consolidar | Bio | G6 | REPRODUZIDA |
| 7 | OC-3 F2: nome fora do mapa, Controllab, HbA1c | Bio | G7 | REPRODUZIDA |
| 8 | OC-7 F2: ET = 1,65×CV com SEM EP | Bio | G9 | REPRODUZIDA |
| 9 | F3: outro nível e ≤ 0 derrubam Sigma/u(Rw) | Bio | G2 | PARCIAL (u(Rw) não reconferida) |
| 10 | F3: ET do Painel com bias 0 | Bio | G9 | REPRODUZIDA |
| 11 | F3: bias com amostras de resultado 0 | Bio | G10 | REPRODUZIDA |
| 12 | OC-F4-1: "Tudo" com mais de 180 corridas | Bio | G11 | REPRODUZIDA |
| 13 | OC-F4-2: urina 5659 na Glicose | Bio | G12 | REPRODUZIDA |
| 14 | OC-F4-11: lotes que não são do analito | Bio | G12 | REPRODUZIDA |
| 15 | Minúsculas no EQUIPAMENTO/ANALITO do manual | Hema | G1 | REPRODUZIDA |
| 16 | Manual implausível (−1, 0) ATIVO | Hema | G2 | REPRODUZIDA |
| 17 | Data 31/02/2026 vira 03/03/2026 | Hema | G3 | REPRODUZIDA |
| 18 | ANALISTA não corrige o CEQ | Hema | G5 | REPRODUZIDA |
| 19 | Correção só depois do Consolidar | Hema | G6 | REPRODUZIDA |
| 20 | Fora do mapa, Controllab, R4 = Controllab | Hema | G7 | REPRODUZIDA |
| 21 | "Lymphs" entra no LYMPH% | Hema | G8 | REPRODUZIDA |
| 22 | Troca de lote: Estatística do lote anterior (F2) | Hema | G13 | REPRODUZIDA |
| 23 | ET = 1,65×CV com SEM EP (F2) | Hema | G9 | REPRODUZIDA |
| 24 | OC-H13: alvo do lote 522611 incompatível | Hema | G14 | REPRODUZIDA |
| 25 | OC-H2: ET do Painel com bias 0 | Hema | G9 | REPRODUZIDA |
| 26 | Estatística não recalcula (lote, Ano De/Até) | Hema | G13 | REPRODUZIDA |

## 3. Reproduções por grupo

### G1: EQUIPAMENTO ou ANALITO em minúsculas no manual (Crítico): REPRODUZIDA nos dois produtos

**Bioquímica** (sessão 20440). Filtro: Glicose, lote 8977, DIMENSION 1, Out/2026. O N2 da produção: n 20, média
243,35, CV 34,2697 %.

- **MAN_0002**, EQUIPAMENTO "dimension 1", 28/10 23:36, 244,222:
  - RUN **26102801**, CORRIDA_NO_DIA 1, ATIVO;
  - toma o lugar do N2 real **BIO-371038 (276)**.
- **MAN_0003**, ANALITO "glicose", 27/10 23:36, 244,222:
  - ANALITO_CADASTRADO = NÃO, RUN **26102701**, ATIVO;
  - toma o lugar do N2 real **BIO-370428 (279)**.

| Fonte (N2) | n | Média | CV % |
|---|---|---|---|
| Painel (`engPainel`) | 20 | 240,0222 | 34,40700544375547 |
| Independente "um valor por RUN" (o manual vence) | 20 | 240,0222 | 34,407005443755466 |
| Estatística | **22** | 243,42927272727272 | 32,58651208371281 |
| Independente sem diferenciar maiúsculas | 22 | 243,42927272727275 | 32,58651208371278 |
| Independente exato (o que o ETL separa) | 20 | 243,35 | 34,26966854739166 |

- **Eng_Saida:** no RUN 26102801, o N2 passa de 276 para 244,222, a Data de 12:57:38 para 23:36:00, e os vereditos N1/N2
  de OK para **REJEITADO**. O mesmo acontece no RUN 26102701 (279 → 244,222).
- **ET da Estatística** (Glicose\|2): 59,278 → 56,501.
- **QA:** nenhum código para MAN_0002 nem para MAN_0003.
- **Validação da tela:** "glicose" é **aceita** (`Validation.Value` = True).
- O controle "DIMENSION 1" também deu False (lista `x12ac` literal "DIMENSION 1,DIMENSION 2"). Por isso, a leitura de
  `Validation.Value` no EQUIPAMENTO da Bio é **inconclusiva**. Precisa de conferência manual na tela; anoto para o UX/UI.

**Hematologia** (sessão 15792). Filtro: WBC, lote 619611, "Tudo", Estatística 2026. O N3 da produção: n 49, média
16,861224.

- **MAN_0001**, EQUIPAMENTO "xn1000", 30/09 23:36, 17,433:
  - RUN 26093001, corrida 1;
  - troca o N3 real **HEM-358501 (16,84)**.
- **MAN_0002**, ANALITO "wbc", 29/09:
  - ANALITO_CADASTRADO = NÃO, RUN 26092901;
  - troca **HEM-357677 (16,94)**.
- **Números do N3:**
  - Painel n 49, média 16,883387755 = independente "um por RUN";
  - Estatística **n 51**, média 16,883647059 = independente sem caixa.
- **Eng_Saida 26092901:** N3 = 17,433, Data 23:36, N1/N3 REJEITADO.
- **QA:** 0 códigos.
- **Validação:** "wbc" é **aceito** e "xn1000" é recusado. A lista do EQUIPAMENTO é o literal "XN1000"; o controle
  "XN1000" deu True.

**Causa (confirmada no código, sem alterar):**

- `src/comum/mEtlCQ.bas` linha 21 (`Option Compare Binary`) monta a corrida com o texto exato;
- `src/bio/mEstatistica.bas` 122, 210 e 213 (`UCase$`) junta em maiúsculas;
- o Eng_Saida guarda um valor por RUN e nível.

**Ampliação:**

- os dois produtos; N2 da Bio e N3 da Hema; dias diferentes dos da F1;
- vale para qualquer analito, porque o caminho é comum;
- na Hema, só o ANALITO passa digitado; o EQUIPAMENTO, só colado.

**Teste automatizado proposto:** `qa_manual_substitui.py` **M20**. Manual com EQUIPAMENTO e ANALITO em minúsculas
(par Bio/Hema). Aceita uma destas saídas: normalizado para o cadastro, MANUAL_INCOMPLETO ou ALERTA. Exige que o
n do `engPainel` seja igual ao n da Estatística, e que nenhum valor real do mesmo RUN suma do Eng_Saida.

### G2: valores implausíveis ATIVOS (Crítico/Alto): REPRODUZIDA (Bio F1, Hema F1); PARCIAL (Bio F3)

**Bio, interfaceamento** (openpyxl, `repro_dados.json` D1):

- **BIO-371183 = −1** (RUN 26102805) e **BIO-371184 = 1** (RUN 26102806): Glicose N2, lote 8977, DIMENSION 1,
  28/10 13:59:55, sem FLAG, ATIVO, PARTICIPA=SIM.
- **Glicose N2 de outubro** na produção: CV **34,27 %** com os dois (n 20) contra 2,08 % sem eles (n 18).
- **Varredura:** **56** ATIVO ≤ 0, todos do INTERFACEAMENTO:
  - 43 negativos e 13 zeros; 55 de analitos cadastrados;
  - de 05/12/2025 a 28/10/2026;
  - Creatinina 6; Colesterol total, Lipase, CTLF e Ureia 5 cada; e outros.

**Bio, F3** (D2):

- **253/253** suspeitos de `estudos/cv_suspeitos_bioquimica.csv` continuam ATIVO, e nenhum ID sumiu.
- **Glicose N1, Sigma da Estatística = 0,32979403081717**:
  - recalculado pelo método do `mIncerteza.IncertezaCIQ` (DIMENSION 2, janela 29/04–28/10, lotes com n ≥ 20):
    CVp 22,0341, Sigma 0,329794030817174, ou seja, bate;
  - sem **BIO-319726 = 277** (N1, lote 8977, DIMENSION 2): CVp 1,6841 e Sigma **4,315**;
  - CV do lote 8977 = **27,2416 %** e do 8974 = **1,2064 %**, como na F3.
- **IDs citados, todos ATIVO:** BIO-232288 = 101,5 (Proteína total N1), BIO-268082 = −0,5 (Proteína total N2) e
  BIO-215815.
- **Status sigma:** **53** dos 58 níveis com status estão em "Desempenho inadequado" (2 Marginal, 3 Bom).
- **Não reconferido:** a u(Rw) de 12 m = 81,2 % da Proteína total, que dá o PARCIAL.

**Hema, manual** (sessão 15792):

- **MAN_0003**, WBC N1 = **−1**, 28/09 23:36: ATIVO, RUN 26092803.
- **MAN_0004**, N2 = **0**, 25/09: ATIVO, RUN 26092503.
- **CV:** N1 de 2,2749 % para **18,2159 %**; N2 de 1,2834 % para **12,2838 %**. Painel, Estatística e independente
  iguais.
- **QA:** só um **A06**, "manual e interfaceamento no mesmo dia". Esse alerta não é de plausibilidade e saiu porque havia
  resultado automático no dia. Não há nenhum alerta sobre o valor.
- **Validação:** −1 e 0 são aceitos (regra decimal ≥ −1E+307).

**Ampliação:**

- O automático da Hema tem hoje 0 valores ≤ 0.
- O manual da **Bio** também aceita −1 e 0 na validação, e por isso o risco do manual vale para os dois produtos.

**Testes propostos:**

- `qa_casos_extremos.py` **X12**: manual ≤ 0 deve gerar ALERTA/MANUAL_INCOMPLETO com o ID;
- `qa_etl.py` **E17**: varredura de ATIVO ≤ 0 na tblCQ_Final deve gerar ALERTA no QA com o ID;
- regra de plausibilidade a definir pelo Especialista em CQ.

### G3: data inválida colada, 31/02/2026 (Alto): REPRODUZIDA nos dois produtos

- **Bio, MAN_0004:**
  - a célula guarda o texto "31/02/2026";
  - depois do ATUALIZAR: DATA_HORA **03/03/2026 08:00**, RUN **26030301**, ATIVO, PARTICIPA=SIM, sem código no QA.
- **Hema, MAN_0005:** idêntico (03/03/2026 08:00, RUN 26030301, ATIVO).
- **Digitada:** a validação recusa nos dois produtos (`Validation.Value` = False).
- **Causa:** `src/comum/mEtlCQ.bas`, `ParaData`, linha 219, e `ParaDataHora`, linha 234: o `DateSerial` sem conferir dia
  e mês.
- **Teste proposto:** `qa_etl.py` **E18**, com datas texto 31/02, 30/02, 29/02/2026 e 31/04. Exige MANUAL_INCOMPLETO
  "DATA" nos motores VBA e PQ.

### G4: EQUIPAMENTO vazio recebe o equipamento do Painel (Alto): REPRODUZIDA (Bio)

- **Antes de digitar:** Painel L4 = DIMENSION 2; `EQUIPAMENTO_PADRAO` da Cfg = DIMENSION 1.
- **MAN_0001**, sem EQUIPAMENTO: a linha ganhou **"DIMENSION 2"** já ao ser digitada.
- **Depois do ATUALIZAR:** DIMENSION 2, RUN 26102703, ATIVO.
- **QA:** só o **A06**, porque havia BIO-369987/369988 no mesmo dia. Não há aviso de equipamento.
- **Causa:** `src/comum/mIntegracao.bas` 711–712 (`EquipPadrao()`).
- **Ampliação:** a Hema tem um equipamento só (XN1000), então não se aplica hoje.
- **Teste proposto:** `qa_manual_substitui.py` **M21**. Com o Painel no equipamento ≠ padrão, um manual sem EQUIPAMENTO
  deve exigir o campo ou avisar.

### G5: o ANALISTA não consegue corrigir nem excluir um CEQ (Alto): REPRODUZIDA nos dois produtos

**Bio** (sessão 26244, estado de ANALISTA com `UserInterfaceOnly`):

- **Inclusão:** `OK|547|`.
- **Mesma chave com 105:** `ERRO|5|Já existe este registro (...) na linha 547 da aba EQA.CAP_Dados. Para corrigir,
  edite a linha existente.`
- **A aba que a mensagem manda editar:** `ProtectContents` = True, `ProtectionMode` (UserInterfaceOnly) = True, e as
  colunas **A..R todas `Locked`**.
- **openpyxl:** aba protegida, colunas A..R Locked, 0 validações.
- **`frmCEQ.frm.txt`:** não tem editar, corrigir nem excluir.

**Hema:** `OK|467|`, depois `ERRO|5|... linha 467 ...`. Aba protegida e A..R Locked.

**Teste proposto:** `qa_adr087.py` **C05**. Com login de ANALISTA, registrar e depois corrigir o mesmo CEQ pela
janela. Exige um caminho de correção, ou a coluna digitável desbloqueada com a consolidação automática.

### G6: a correção na tabela só chega ao bias depois do "Consolidar" (Alto): REPRODUZIDA nos dois produtos

**Bio** (Proteína total, lote 8977, DIMENSION 2):

| Passo | G (bias) | ET N1 / N2 | Sigma N1 / N2 |
|---|---|---|---|
| B0 | 3,3793363600341007 | 6,4486 / 7,0114 | 1,9907 / 1,1176 |
| Inclusão 110 | 5,034502270025575 | 8,103745678586874 / 8,666611564082336 | 1,4930545183334316 / 0,8382017183904314 |
| Célula F547 = 105 | **inalterado (0 linhas da Estatística mudaram)** | | |
| Depois do Consolidar | 3,7845022700255755 | 6,853745678586875 / 7,416611564082336 | 1,868911733341443 / 1,0492081884293243 |

Os valores são idênticos aos do relatório B da F2.

- Depois da edição, a tabela mostra O547 = 5, mas a EQA_Base continua com **110/10**.
- **Z2** = "consolidado: 2026-08-22 18:48:37 | CAP 545 ..." e **Z3** = "base em dia", antes e **depois** do Consolidar.
- O Z1 da EQA_Base passou a "2026-11-03 07:49:04 | CAP 546".

**Hema** (RBC):

- G: 1,5336 → 3,6502 (inclusão) → inalterado depois da edição (O467 = 5; EQA_Base 110/10) → **2,4002** depois do
  Consolidar.
- Sigma N1: 2,528 → 0,358 → 1,640.
- Z2 continua com o carimbo de **2026-08-24 15:20:53**, e Z3 diz "base em dia".

**Teste proposto:** `qa_adr087.py` **C06**. Editar o resultado de uma linha do CEQ deve mudar a Estatística G sem outro
passo, ou o Z3 deve dizer "falta consolidar". O Z2 deve acompanhar o Z1.

### G7: CEQ com nome fora do mapa, do Controllab ou de analito novo, gravado OK e fora do bias (Alto): REPRODUZIDA nos dois produtos

**Bio:** quatro gravações, todas com Analito_Canonico vazio na EQA_Base e **0 linhas da Estatística mudadas**. A HbA1c
continua "SEM EP".

| Nome gravado | Retorno | `MapearAnalito` |
|---|---|---|
| "Protein, total" (sem ", serum") | `OK|548|` | "" |
| Controllab "Proteína total" | `OK|2|` | "" |
| "Hemoglobin A1c" | `OK|549|` | "" |
| "Hemoglobina glicada" | `OK|550|` | "" |

**Hema:**

- "Red Blood Cell" deu `OK|468|`, e o Controllab "RBC" deu `OK|2|`. Os dois ficaram com canônico vazio e 0 linhas
  mudadas.
- O mapa W:Y tem **31 entradas, todas do CAP** (D8).
- **Provedor R4 = Controllab:**
  - SEM EP passa de 12 para **84 de 84** linhas, e o H fica vazio nas 84;
  - K5 = "Controllab | ... 0 rodada(s), 0 linha(s)";
  - isso acontece mesmo com o registro Controllab do RBC gravado.

**Causa:** o `mCEQForm.SalvarCEQ` (`src/comum/mCEQForm.bas` 169–268) grava e consolida sem consultar o
`mEQA.MapearAnalito`.

**Teste proposto:** `qa_adr087.py` **C07**. `SalvarCEQ` com um nome fora do mapa deve devolver aviso ("não mapeado").
Com o provedor Controllab, deve existir mapa ou aviso.

### G8: "Lymphs" (sem Absolute) entra no bias do LYMPH% (Alto): REPRODUZIDA (Hema)

- **Gravação:** "Lymphs", unidade **x 10E9/L**, 110/100, deu `OK|469|` sem aviso, com `MapearAnalito` = **"LYMPH%"**.
- **LYMPH%:**
  - G 3,6294 → **5,2220** nos 3 níveis;
  - Sigma N1 4,3264 → 3,9662 ("Bom" → "Marginal");
  - ET N1 11,353 → 12,946.
- **LYMPH#:** idêntico.
- A unidade x 10E9/L foi aceita para um analito em %.
- **Magnitude:** a F2 viu G = 15,22, e eu vi 5,22. A diferença vem do desenho do dado: aqui, uma amostra única numa
  rodada nova, (3 × 3,63 + 10) / 4. A associação errada é a mesma.

**Ampliação:**

- O mapa tem **7 pares** "nome curto → %" / "nome + Absolute → #": Basophils, Eosinophils, IG, Lymphs, Monocytes,
  Neut/Gran e nRBC. Os 7 têm o mesmo risco.
- Na Bio, o mapa não tem pares desse tipo.

**Teste proposto:** `qa_adr087.py` **C08**. Para cada par % / #, gravar o nome curto com a unidade do "#" deve gerar
aviso de unidade divergente ou mostrar o canônico antes de gravar.

### G9: ET do Painel = 1,65 × CV quando o bias é "SEM EP" (Alto): REPRODUZIDA nos dois produtos (4 ocorrências)

**Fórmula** (openpyxl, data_only=False):

- Painel!H7 = `IF($E7="","",$E7*1.65+ABS(IF(ISNUMBER($G7),$G7,0)))`, igual na Bio (H7:H8) e na Hema (H7:H9);
- a Estatística!H14 vale `""` quando G não é número.

**Valores no Excel:**

| Produto | Caso | Painel H (ET %) | 1,65 × CV | Estatística H | Bloco ERRO TOTAL |
|---|---|---|---|---|---|
| Bio | HbA1c, 8590, DIM2, Out/2026 | 2,5382 / 2,7842 | idêntico | vazio | C61:C63 vazio |
| Bio | PCR, 8353, DIM2, Out/2026 | 2,3723 / 2,0146 | idêntico | vazio | vazio |
| Bio | Proteína total, M4 = Controllab | 3,0692 / 3,6321 (antes 6,4486 / 7,0114) | idêntico | — | — |
| Hema | RET%, 619611, 2026 | 5,8606 / 5,5504 / 8,2934 | idêntico | vazio | — |
| Hema | RBC, R4 = Controllab | 1,8718 / 1,2478 / 1,3286 | idêntico | vazio (84/84) | — |

- **Linhas SEM EP:** na Bio, 4 (HbA1c e PCR); na Hema, 12 (RDW-SD, RET%, RET# e IRF).
- **Números da Hema diferentes dos da F3/F2:** lá era o lote 522611; aqui, o 619611. A fórmula e o defeito são os mesmos.

**Teste proposto:** `qa_rodadas_cv.py` **R12**. Para todo analito com G = "SEM EP", exige Painel H7:H9 vazio (ou
rotulado), igual à Estatística H.

### G10: o bias do ET/Sigma inclui amostras de resultado 0 (−100 %) (Alto, dúvida de método): REPRODUZIDA (Bio)

- **Amostras com resultado 0:** `CAP|2025|C-A 2025|Bilirubin, direct|CHM-04` (0 contra 0,05) e
  `...|C-C 2025|...|CHM-13` (0 contra 0,03), com Bias_Abs = 100.
- **Recálculo pelo estimador do `BiasEQ` (duas etapas):**
  - **com** as duas amostras: 30,770933248452042, **igual** à Estatística G20 (30,77093324845204);
  - **sem** elas (filtro do `ViesEQ`): **19,99420975515574**.
- **Rodadas:** C-A 44,857 → 31,071; C-B 21,633; C-C 25,822 → 7,278.
- **Sem os zeros:** ET N1 26,34 % (37,11 hoje); ET N2 22,87 % (33,65); Sigma N1 **0,947** (0,435); Sigma N2 2,698
  (1,238).
- **Ampliação, Hema:** a EQA_Base tem **10 amostras** com resultado 0: NRBC# 5 e NRBC% 5 (FH9-04, 06, 10, 11 e 13,
  Bias_Abs 100), que entram no G do NRBC. O NRBC não tem ETp hoje, então o efeito em ET/Sigma fica latente.
- **Teste proposto:** `qa_incerteza.py` **I17**, a montar depois da decisão do Especialista em CQ. Exige o mesmo
  critério de amostra no BiasEQ "ABS" e no ViesEQ.

### G11: "Tudo (sem filtro)" num lote com mais de 180 corridas (Alto): REPRODUZIDA (Bio)

Filtro: Glicose, DIMENSION 1, H3 = 8974, K4 = "Tudo (sem filtro)".

| Nível | Painel: n / média / DP | Todos os dados (independente) |
|---|---|---|
| N1 | 140 / 89,15 / 16,967987 | 174 / 92,287356 / 30,373744 |
| N2 | 131 / 281,984733 / 4,922452 | 161 / 281,484472 / 16,114863 |

- Eng_Saida I1 = 180 slots.
- A39 diz "⚠ 39 corrida(s) antigas do período fora do **gráfico** (máx. 180)", e não fala da estatística.
- **Ampliação, Hema:** o maior lote × analito × nível tem 86 resultados (528211), abaixo de 180, então não se aplica
  hoje. O código da janela (NK = 180, `AtualizarCalc`) é comum.
- **Teste proposto:** `qa_adr086.py` **Y11**. Com "Tudo" num lote de mais de 180 corridas, exige que o n do Painel seja
  o n total, ou um aviso explícito de que a estatística também foi cortada.

### G12: lotes que não são do analito, ativos em analitos séricos (Alto): REPRODUZIDA (Bio, 2 ocorrências)

**Lote de urina 5659 na Glicose:**

- **28** linhas GLUC, MATRIZ = SORO, todas ATIVAS: N1 n 16, média 55,94; N2 n 12, média 33,25. Ex.: BIO-355958 = 33.
- **Alvo cadastrado:** N1 59,67 ± **1,5428** e N2 33 ± **5,0270**. Os dois DPs são **iguais aos do lote 8977** (DP N1
  1,5428 e N2 5,0270): o DP foi copiado.

**Lote 6353 na Ureia:**

- BIO-230952 = 871 e BIO-230958 = 1830, ativos.
- Alvo **871 ± 1,7094 / 1830 ± 4,8**: o alvo é o próprio valor.

**Lote 6354:**

- Creatinina N2 com média 196,69 (n 6) contra o alvo 176,55 ± 0,147 (z = +137).
- Os lotes 6353/6354 estão ativos em Na, K, Cl, P, Ca, Creatinina e Ureia.

**Lote 8353 (PCR):**

- **30** linhas de outros analitos ativas, com CV 316 % no N1.
- Alvo da Glicose e da Ureia = **−1**, criado do valor −1.

**Lote 8974:** **14** linhas de Hemoglobina glicada.

**Lote 8794:** um ponto por analito, com alvo = o próprio valor em todos (z = 0).

**Ampliação:** varredura de 257 lote × analito × nível com alvo:

- 6 com |z| > 5, entre eles 8974 Amilase N1 com z = +7,0, n 373, a examinar;
- vários com n ≤ 2 e alvo idêntico ao valor (8794, 6353).

**Teste proposto:** `qa_varredura_lotes.py` **V05**. Exige alerta para um lote cujo alvo foi criado com n ≤ 2, para
|média − alvo| > 5 DP depois de n ≥ 20 e para MATRIZ/lote de urina num analito sérico. Também vale para o modo TODOS
(ADR-086).

### G13: a Estatística não recalcula na troca de lote nem no Ano De/Até (Crítico/Alto): REPRODUZIDA (Hema, 2 ocorrências)

| Passo | B4 | WBC N1: n / média / CV % | RBC N1: n / CV % / ET % | Células mudadas |
|---|---|---|---|---|
| Antes (522611, 2025) | 522611 | 50 / 3,0446 / 1,8041 | 50 / 1,2595 / 3,6119 | — |
| Painel H3 = 619611 | **619611** | 50 / 3,0446 / 1,8041 | 50 / 1,2595 / 3,6119 | **0** |
| Botão Estatística (`mApp.IrEstatistica`) | 619611 | idem | idem | **0** |
| B3 = D3 = 2026 | 619611 | idem | idem | **0** |
| Independente 619611 / 2026 | — | 54 / 3,120556 / 2,2749 | 54 / 1,1344 | — |
| Independente 619611 / 2025 | — | **n 0** | n 0 | — |
| Depois de `mEstatistica.AtualizarEstatisticaAba` | 619611 | 54 / 3,120556 / 2,2749 | 54 / 1,1344 / 3,4054 | iguais ao independente |

- Enquanto isso, o Painel já mostrava o lote 619611 (RBC N1 n 54, CV 1,1344, ET 3,4054).
- **Causa:** `src/hema/mLotes.bas` 398–444. O `TrocarLoteAnalise` chama só o `RecalcularAnalitoAtual`, e a aba
  Estatística não tem evento de planilha.
- **Ampliação, Bio:** não afetada. No G9, a Estatística da Bio acompanhou a troca de analito, lote e período do Painel
  (HbA1c\|1: n 15, CV 1,5383 = Painel). Isso confirma a F2-T37 da Bio.
- **Teste proposto:** `qa_troca_lote.py` **L04**. Depois de trocar o H3, e depois de mudar B3/D3, exige que a
  Estatística C:F seja igual à do independente para o lote de B4, ou que apareça "DESATUALIZADO".

### G14: alvo do lote 522611 incompatível com os dados (Alto): REPRODUZIDA (Hema)

Openpyxl (D7) e Excel (`mLotes.MediaDoLote`/`DPDoLote`) dão o mesmo resultado: **12 pares com |z| > 5**, todos do lote
522611.

| Analito, nível | Alvo | Média medida | z |
|---|---|---|---|
| RDW-SD N2 | 14,7793 | 47,3765 | **+398,0** |
| IG# N2 | 11,0517 | 0,7798 | −31,6 |
| IG# N3 | 11,8759 | 2,0094 | −29,5 |
| IG# N1 | 10,2897 | 0,3036 | −28,6 |
| RDW-CV N2 | 9,8345 | 17,5373 | +25,7 |
| HGB N2 | 12,2759 | 11,0647 | −14,6 |
| MCH N2 | 28,0552 | 25,449 | −12,4 |
| MCV N2 | 81,8759 | 76,5157 | −7,3 |
| RDW-CV N3 | 14,9448 | 14,349 | −6,5 |
| HCT N2 | 35,8345 | 33,2588 | −5,7 |
| RET# N1 | 0,1068 | 0,1444 | +5,6 |
| RET% N1 | 4,6307 | 6,0104 | +5,2 |

- Nos 12 pares, `mLotes.AlvoDeTeste` = **False**: o aviso de TESTE do ADR-076 não cobre nenhum deles.
- **Ampliação:** 700 pares lote × analito × nível avaliados, nos 9 lotes (522611, 528211, 533811, 536411, 602911,
  608511, 614011, 619611 e 625311). **0 pares** fora dos 12 do 522611. Na Bio, a mesma varredura está na G12.
- **Teste proposto:** `qa_lotes_limites.py` **L09**. Para cada lote, exige alerta quando |média − alvo| > 5 DP do alvo
  com n ≥ 20.

## 4. Achados novos e observações (não estão na rota)

1. **Bio, validação do EQUIPAMENTO:** `Validation.Value` dá False também para "DIMENSION 1" (lista `x12ac` literal
   "DIMENSION 1,DIMENSION 2"). Não sei dizer se a tela recusa o valor certo digitado ou se é um efeito da leitura por
   COM. **Não confirmado.** Sugiro que o UX/UI confira à mão, digitando, numa cópia.
2. **Botão "Consolidar na EQA_Base" sem tratamento de erro:** ver o achado de método na seção 1. Só aparece sem
   `UserInterfaceOnly` (automação sem login). Prioridade baixa: é de robustez, não do usuário.
3. **Bio, 8974 Amilase N1:** média 79,54 contra o alvo 74,05 ± 0,785 (z = +7,0, n 373). Pode ser deriva real ou alvo
   desatualizado. Vai ao Especialista em CQ.

## 5. Segurança e limpeza

- A produção não foi aberta: `fe10dba1f1fd` e `0b45fa270d2a`, iguais antes e depois.
- Todas as cópias ficaram com o hash igual antes e depois de cada sessão (nada salvo) e foram apagadas no fim. Ficaram
  só os JSON e os logs em `%TEMP%\...\tester\`.
- Só os PIDs criados pelos meus scripts foram encerrados: 28976, 20440, 26244, 5456 e 15792. No 20440, a macro parada
  foi encerrada pelo botão "Fim" do diálogo do próprio processo.
- Nenhuma captura de tela, nenhum caminho de rede e nenhuma senha neste relatório nem nos scripts. A senha das abas é
  lida do instalador em tempo de execução.
- `src/` não foi alterado.
