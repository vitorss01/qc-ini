# Parecer consolidado: ciclo 2026-11-14_reteste04

Agente: Usuário Simulado (parecer consolidado). Data: 14/11/2026. Modo: reteste, consolidação. Neste passo o Excel não
foi aberto e nenhum número foi recalculado. Tudo aqui vem dos 4 relatórios das frentes desta pasta e da lista do
reteste (`ocorrencias_reteste.json`). Os valores são os que as frentes mediram e registraram.

| Item | Valor |
|---|---|
| Produção auditada (entrega 41) | `QC_Bioquimica.xlsm` sha256 `8ec815ab653db405`; `QC_Hematologia.xlsm` sha256 `aab0a54f464e0232`. Em todas as execuções o hash foi o mesmo antes e depois (`producao_intacta.igual = true`): a produção não foi aberta para escrita |
| Frentes executadas | F1 (entrada de CIQ manual/automática e janela do CEQ) e F2 (CEQ, períodos da Estatística, regras de lote, datas), nos dois produtos. Neste reteste, a F3 (auditoria matemática) rodou **dentro da F2**: calculadora independente sobre 12 retratos da Bio e 8 da Hema. A F4 completa (LJ, Westgard, corrida do lote) **não fazia parte deste reteste** |
| Relatórios de origem | `Bioquimica_F1.md`, `Bioquimica_F2.md`, `Hematologia_F1.md`, `Hematologia_F2.md` |
| Lista do reteste | 27 itens: 24 OC (OC-0012 a OC-0066) e 3 ADR (ADR-091, ADR-093, ADR-094) |
| Dados de teste | Só nas cópias de `%TEMP%\qcw\auditoria\2026-11-14_reteste04\`, marcados `TESTE-AUDITORIA` (usuários, manuais, rodadas de CEQ, mapas, regras da tblLotesMaterial, analito de teste). Nenhuma cópia foi salva e todas foram apagadas no fim. Ficam os JSON e os logs |
| Excel | Um por vez, aberto e encerrado pelo `qa_final.QA` de cada frente: Bio F1 PIDs 2020, 23060 e 7148; Bio F2 6920 e 3800; Hema F1 5424 e 1740; Hema F2 22512 e 10616. Os Excel da tarefa agendada não foram tocados. Nenhuma captura de tela |
| Ferramentas | `_arquitetura/agentes/ferramentas/reteste04_usuario_f1.py`, `reteste04_usuario_f1_hema.py`, `reteste04_usuario_f2.py`, `reteste04_usuario_f2_hema.py` e `calc_independente.py` |

**Resumo em números**

| | Bioquímica | Hematologia | Total |
|---|---|---|---|
| Entradas de teste | 39 (F1 20, F2 19) | 30 (F1 16, F2 14) | 69 |
| RESOLVIDA / PERSISTE / EFEITO_COLATERAL | 25 / 1 / 0 | 18 / 1 / 1 | 43 / 2 / 1 |
| Testes novos e vizinhos: APROVADO / REPROVADO / NAO_EXECUTADO | 10 / 2 / 1 | 8 / 1 / 1 | 18 / 3 / 2 |
| Cálculos no Relatório B (aprovados / reprovados) | 23 (23 / 0) | 24 (24 / 0) | 47 (47 / 0) |
| Comparações da calculadora independente | 12 retratos × 1.178 = 14.136, 0 divergências | 8 retratos × 2.313 = 18.504, 0 divergências | 32.640, 0 divergências |
| Ocorrências das frentes / deduplicadas (P-xx) | 6 / 5 | 5 / 5 | 11 / 10 |
| Altas / Médias / Baixas (deduplicadas) | 1 / 2 / 2 | 1 / 3 / 1 | 2 / 5 / 3 (nenhuma Crítica) |
| Dúvidas de regra de negócio (deduplicadas) | | | 15 |

**Situação dos 27 itens do reteste (os dois produtos juntos)**

| Item | Bio | Hema | Consolidado |
|---|---|---|---|
| OC-0012 Painel janela 180 | RESOLVIDA | EFEITO_COLATERAL | **EFEITO_COLATERAL** (só na Hema: P-05) |
| OC-0019 EQUIPAMENTO vazio sem aviso (Bio) | RESOLVIDA | — | **RESOLVIDA** no que a OC pedia. O fluxo vizinho ("digite o certo") é barrado pelo P-01 |
| OC-0021 RUN/DETALHE com 0, DETALHE sem acento | PERSISTE | PERSISTE | **PERSISTE** (parcial: o "0" foi resolvido; o acento falta no IMPLAUSIVEL e na nota do ADR-093) |
| OC-0023 Audit_Log de manual incompleto/em conflito | RESOLVIDA | RESOLVIDA | **RESOLVIDA** |
| OC-0024 ResumoQA com E04 de manual | RESOLVIDA | RESOLVIDA | **RESOLVIDA** |
| OC-0031 frmCEQ sem valor-alvo | RESOLVIDA | RESOLVIDA | **RESOLVIDA** (o clique na caixa Sim/Não não foi executado) |
| OC-0033 frmCEQ sem sessão | RESOLVIDA | RESOLVIDA | **RESOLVIDA** |
| OC-0034 Listas do frmCEQ com nomes da simulação | RESOLVIDA | RESOLVIDA | **RESOLVIDA** |
| OC-0035 GUIA §5–6 | RESOLVIDA | RESOLVIDA | **RESOLVIDA** (dúvidas de texto D-03 e D-15) |
| OC-0036 Estatística AT "004 / 000" | RESOLVIDA | RESOLVIDA | **RESOLVIDA** |
| OC-0040 Estatística MENSAL/TRIM e "Ano 0" (Bio) | RESOLVIDA | — | **RESOLVIDA** (vizinho "Tudo": P-09, Baixo) |
| OC-0041 Painel K3/K4 depois de De/Até | RESOLVIDA | RESOLVIDA | **RESOLVIDA** |
| OC-0043 Painel J7:J8 cortado (Bio) | RESOLVIDA | — | **RESOLVIDA** |
| OC-0044 Estatística: alinhamento vertical (Bio) | RESOLVIDA | — | **RESOLVIDA** |
| OC-0049 Estatística: linhas vazias (Hema) | — | RESOLVIDA | **RESOLVIDA** (também conferida de carona no Hema F1) |
| OC-0050 Principal A5 "Total" | RESOLVIDA | RESOLVIDA | **RESOLVIDA** |
| OC-0052 Aviso do par de nomes no mapa | RESOLVIDA | RESOLVIDA | **RESOLVIDA** |
| OC-0054 Edição direta do CEQ sem selecionar | RESOLVIDA | RESOLVIDA | **RESOLVIDA** |
| OC-0055 Descrição do CEQ com o ano repetido | RESOLVIDA | RESOLVIDA | **RESOLVIDA** |
| OC-0056 Manual de urina com nome de soro (Bio) | RESOLVIDA | — | **RESOLVIDA** |
| OC-0057 Corrigir/Excluir CEQ sem sessão | RESOLVIDA | RESOLVIDA | **RESOLVIDA** |
| OC-0058 Trilha de manual fora da estatística | RESOLVIDA | RESOLVIDA | **RESOLVIDA** |
| OC-0059 Documentação do ADR-090 (Bio) | RESOLVIDA | — | **RESOLVIDA** |
| OC-0066 Datas com segundos em coluna vazia e Sync_Log | RESOLVIDA | RESOLVIDA | **RESOLVIDA** |
| ADR-091 Material definido pelo lote | RESOLVIDA (regra) | RESOLVIDA (regra) | **RESOLVIDA** na regra e na estatística. A mensagem do Painel fica em P-04/P-07. Não executados: MATRIZ na Hema e código ambíguo |
| ADR-093 Manual de urina aceito como urina (Bio) | RESOLVIDA | — | **RESOLVIDA** (o caso do código ambíguo não foi executado: não há caso real) |
| ADR-094 Data-hora com segundos, sem #### | RESOLVIDA | RESOLVIDA | **RESOLVIDA** no escopo do ADR. A BI_Data fica fora (D-07) |

---

## Relatório A: experiência do usuário

### A.1 Funcionalidades testadas

| Frente | Bioquímica | Hematologia |
|---|---|---|
| F1 | 16 OC refeitas pelo caminho do usuário: Painel (L4, C3, H3, K3, K4, G3/G4, com o Painel protegido), Digitar Resultados (botão `IrDigitarResultados`, digitação célula a célula, ATUALIZAR pelo núcleo do botão), janela do CEQ (`AbrirCEQ` real e o núcleo dos botões), edição direta da EQA.CAP_Dados. Também os vizinhos ADR-093 e OC-0053, e as listas de validação (achado novo P-01) | 14 OC refeitas pelos mesmos caminhos, mais o vizinho das listas de validação (P-02) e o OC-0053 de carona |
| F2 | 10 itens (OC-0040, 0043, 0056, 0057, 0058, 0059, 0066, ADR-091, 093, 094) + pedido do ciclo: incluir, corrigir e excluir CEQ de Proteína total e de Glicose, CAP 2024, Controllab mapeado e duplicidade. Calculadora sobre 12 retratos | 6 itens (OC-0049, 0057, 0058, 0066, ADR-091, ADR-094) + o mesmo pedido com HGB, troca do R4 para Controllab e de volta, e analito novo na Especificações. Calculadora sobre 8 retratos |

### A.2 Facilidade de uso

- **CEQ (os dois produtos): é o fluxo mais sólido do sistema.**
  - A sequência segue o relatório do provedor.
  - O aviso azul diz para qual analito o resultado vai, e a lista "Registros desta rodada" mostra o que acabou de ser
    gravado.
  - A duplicidade diz a linha e o botão para corrigir.
  - O número muda na hora e só no analito certo. Excluir volta a Estatística ao estado inicial, igual bit a bit.
  - Sem login a janela nem abre e a caixa diz por quê. Com a sessão encerrada no meio, nada é gravado.
  - O registro sem valor-alvo avisa a consequência antes de gravar e não mexe no bias.
- **Manuais (os dois produtos):**
  - O botão leva direto à linha livre.
  - STATUS e DETALHE dizem o campo que falta, e o conflito diz os minutos.
  - O RUN fica em branco quando o manual não entra.
  - A mensagem do ATUALIZAR separa os manuais incompletos dos erros de governança.
  - Na Bio, o âmbar com nota no EQUIPAMENTO preenchido pelo Painel chama atenção.
- **Período:** K3/K4 mostram "Personalizado" depois de De/Até, e a Estatística da Bio segue o Painel sem "Ano 0".
- **Janela de 180 corridas:** o aviso agora diz que o corte vale para o gráfico **e** para os indicadores, e dá o n
  do período inteiro, que confere.
- **Datas:** segundos em todas as abas do dia a dia, sem "####".

### A.3 Dificuldades encontradas

- **Lançar um manual com o equipamento ou o motivo certos (P-01, P-02).** O âmbar pede "confira (ou digite o
  certo)", mas a validação da coluna recusa "DIMENSION 1" e cada um dos 4 motivos. A seta oferece uma única opção com
  os itens colados por vírgula, e essa opção grava um equipamento inexistente que entra ATIVO. Hoje, quem precisa
  lançar um manual na Bio só consegue deixando o EQUIPAMENTO vazio com o Painel já no equipamento certo, o que o GUIA
  não diz. Na Hema, nenhum motivo de substituição (ADR-082) passa pela validação.
- **Painel em lote com regra da tblLotesMaterial (P-04, P-07).** O Painel diz "SEM CORRIDA / SEM MÉD/DP", "digite a
  Média e o DP" e "o lote não tem resultado", quando o lote tem dezenas de resultados excluídos pela regra. O
  DETALHE e o QA A13 explicam, mas o Painel não.
- **Aviso da janela de 180 na Hema (P-05).** O "veja a aba Estatística" leva a outro recorte (lote 522611, ano 2025),
  e o n citado (374/400/340) não aparece em lugar nenhum.
- **Aviso do CEQ do Controllab (P-03, P-06).** Um nome do Controllab mapeado aparece como "entra no bias", e não entra
  enquanto o provedor do analito (AR na Bio, R4 na Hema) for o CAP. Hoje é latente: o mapa do Controllab está vazio
  nos dois produtos.
- **Cabeçalho da Estatística da Bio em "Tudo (sem filtro)" (P-09).** De/Até mostram "00/01/1900" e o Ano diz 2026.
- **Painel da Hema no modo TODOS** (impressão do Hema F1, sem ocorrência): mostra n e Sigma sem Média, DP, CV e ET
  (ADR-086) e dá REJEITADO nos três níveis. Quem não conhece o ADR-086 fica sem saber de onde vem o REJEITADO.

### A.4 Navegação

- Os botões citados pelo GUIA existem: "Registrar resultado do CEQ" nos Resumos, "Consolidar na EQA_Base" nas abas de
  dados, ATUALIZAR no Painel, no DB, no DB_ORGANIZADO e na Digitar Resultados.
- O GUIA §5 diz "o botão Consolidar da EQA_Base", mas o botão fica nas abas de dados, não na EQA_Base (D-15).
- A numeração das seções do GUIA está fora de ordem (§6, §7, §6b, §6c, §7b, §8), o que atrapalha achar o item (Bio
  F1).
- Lote SO_ANALITOS na Hema: o Painel troca o analito sozinho (RBC → WBC) e avisa no A39. Isso ajuda a achar os dados,
  mas não diz que o RBC foi excluído pela regra (D-13).
- Urina na Bio: os manuais aceitos como urina (CRE2) e a série de urina participam da estatística, mas nenhuma tela
  os mostra (D-05, ligada à OC-0061).

### A.5 Apresentação

- AT "x,xx / y,yy", A5 "Total: … · participam: …", J7:J8 sem corte (pior redução 26,4 %), A:B centralizados, linhas
  vazias limpas na Hema: tudo legível e conferido.
- Ainda há textos sem acento ao lado de textos acentuados: o DETALHE do IMPLAUSIVEL e a nota do ADR-093 (P-08, P-10).
- A instrução A3 da Digitar Resultados (texto comum) promete "fica em âmbar" também na Hema, onde não acontece (D-10).
- A BI_Data, visível a todos depois do login, mostra "####" e datas aaaa-mm-dd. Fica fora do escopo, porque o BI não
  pode ser mexido (D-07).
- Na Hema, a coluna U ("ordem crítico") de outros analitos se renumera quando um CEQ muda a classe do HGB. Está
  correto, mas quem compara duas impressões pode estranhar.

### A.6 Tempos observados

- Salvar, Corrigir e Excluir do CEQ: 15–25 s, dentro do que o GUIA §5 promete.
- ATUALIZAR completo: Bio 23–26 s; Hema 27–32 s. Só entradas: Bio 15 s; Hema 20,7 s.
- Na Hema, o 1º ATUALIZAR depois de uma regra nova na tblLotesMaterial levou **74 s** (o Westgard refotografou 468
  mudanças de status). Vale avisar no GUIA.

### A.7 Sugestões do usuário (para o Orquestrador priorizar; nenhuma foi implementada)

1. Listas da Digitar Resultados com o separador do idioma ou apontando para um intervalo/nome, e reinstalar nas duas
   produções (P-01, P-02). Também um QA para EQUIPAMENTO fora da lista do produto.
2. Painel A39/Q4: com regra ativa para o lote, dizer "lote fora do CQ / só PCR (tblLotesMaterial)" e não pedir
   Média/DP (P-04, P-07).
3. Janela do CEQ: comparar o provedor digitado com o AR (Bio) ou o R4 (Hema) no aviso e no retorno do Salvar (P-03,
   P-06).
4. Hema: trocar o "veja a aba Estatística" do aviso da janela de 180 quando o H3 está vazio (P-05).
5. Estatística da Bio: B5/D5 e D4 vazios com o Painel em "Tudo (sem filtro)" (P-09).
6. Acentuar o DETALHE na origem ou completar a lista de acentuação (P-08, P-10).
7. GUIA: numeração das seções, nome e local do botão Consolidar, o atalho do equipamento vazio enquanto o P-01 não é
   corrigido, e o aviso do 1º ATUALIZAR mais lento depois de uma regra nova.

---

## Relatório B: validação matemática

**Definição usada (conferida no código pelas frentes F2 antes de calcular):** `src/comum/mCEQ.bas` (`BiasEQ`,
`AnoVigente`, `AmostraDoVies`).
- **Amostras:** as da EQA_Base do analito canônico, com o provedor de Especificações!AR (Bio) ou Estatística!R4 (Hema,
  global), o ano vigente N4 = 2025 e a rodada P4 = TODAS. Ficam fora as NAO AVALIADO, as com resultado ou alvo 0 e as
  chaves repetidas.
- **bias_i** = (Resultado − Alvo)/Alvo × 100. **Bias** = média dos |bias| de cada rodada e, depois, média das
  rodadas (duas etapas).
- **ET** = 1,65 × CV do período (Estatística F) + |bias|.
- **Sigma** = (ETp − |bias|)/CVp, com o CVp = CV pooled da janela de 6 meses (14/05–13/11/2026), lotes com n ≥ 20.
- **ETp:** Bio 10 % (CLIA) para Proteína total e Glicose; Hema 4 % (CLIA) para HGB.
- **CIQ:** n, média, DP amostral (n − 1) e CV = DP/média × 100 das linhas da tblCQ_Final com PARTICIPA = SIM.

Tolerância: 1e-9 relativo sobre o `Value2` (REGRAS §3.5). A maior diferença foi 8e-12, que vem da ordem de soma do CV
pooled.

### B.1 Bioquímica: 23 cálculos, 23 aprovados (origem: Bioquimica_F2.md, Relatório B; Bioquimica_F1.md)

| Analito | Equipamento | Lote | Nível | Período | Dados de entrada | Fórmula encontrada | Resultado independente | Resultado QC_INI (Value2) | Diferença | Status |
|---|---|---|---|---|---|---|---|---|---|---|
| Proteína total | DIMENSION 1 | (bias não depende) | N1/N2 | CEQ CAP 2025, rodada TODAS | C-A 4,61503 · C-B 2,35503 · C-C 3,16794 (5 amostras cada) | média das médias de \|bias\| por rodada | 3,3793363600341007 | 3,3793363600341007 | 0 | APROVADO |
| Proteína total | DIMENSION 1 | — | N1/N2 | + TA-F2-01 7,7/7 | + rodada TESTE 10,0 | (10,13801 + 10)/4 | 5,034502270025575 | 5,034502270025576 | 2e-16 | APROVADO |
| Proteína total | DIMENSION 1 | — | N1/N2 | corrigido 7,35/7 | rodada TESTE 5,0 | (10,13801 + 5)/4 | 3,7845022700255755 | 3,784502270025574 | 4e-16 | APROVADO |
| Proteína total | DIMENSION 1 | pooled | N1 | Sigma, janela 6 m | CVp N1 4,67584 = (10 − 3,37934)/1,41593; ETp 10 | (10 − 5,03450)/4,67584 | 1,0619486593 | 1,0619486593062732 | < 1e-9 | APROVADO |
| Proteína total | DIMENSION 1 | pooled | N1 | Sigma: base → inclusão → correção | idem | (ETp − \|bias\|)/CVp | 1,41593 → 1,06195 → 1,32928 | 1,4159315457416979 → 1,0619486593062732 → 1,329280535448261 | < 1e-9 | APROVADO |
| Proteína total | — | — | — | + CAP 2024 (+100 %) | ano ≠ Ano EP (N4 2025) | não entra | sem mudança | sem mudança (0 células) | 0 | APROVADO |
| Proteína total | DIMENSION 1 | — | N1/N2 | F1 OC-0031: CEQ sem alvo | TA-55 sem alvo | não entra | bias inalterado | 3,3793363600341007 (igual bit a bit) | 0 | APROVADO |
| Proteína total | DIMENSION 1 | — | N1/N2 | F1 OC-0031: com alvo 110/100 | + rodada TESTE 10,0 | (10,13801 + 10)/4 | 5,034502270025575 (o independente da F2, mesma configuração) | 5,034502270025575 | 0 | APROVADO |
| Glicose | DIMENSION 1 | 8977 | N1/N2 | CEQ CAP 2025 | C-A 1,49097 · C-B 1,43622 · C-C 5,27270 | média das médias | 2,7332944432981674 | 2,733294443298167 | 2e-16 | APROVADO |
| Glicose | DIMENSION 1 | 8977 | N1 | ET, 01/01–31/12/2026 | n 90, CV 2,189325928446349 | 1,65·CV + 2,73329 | 6,3456822252 | 6,3456822252346425 | < 1e-12 | APROVADO |
| Glicose | DIMENSION 1 | 8977 | N1/N2 | + TA-F2-G1 110/100 | + rodada TESTE 10,0 | (8,19988 + 10)/4 | 4,549970832473625 | 4,549970832473625 | 0 | APROVADO |
| Glicose | DIMENSION 1 | 8977 | N1 | ET, idem | CV 2,18933 | 1,65·CV + 4,54997 | 8,1623586144101 | 8,1623586144101 | 0 | APROVADO |
| Glicose | DIMENSION 1 | 8977 | N2 | ET, idem | CV 2,502236 | 1,65·CV + 4,54997 | 8,67866038695 | 8,67866038694817 | < 1e-12 | APROVADO |
| Glicose | DIMENSION 1 | pooled | N1 | Sigma, 14/05–13/11/2026 | CVp N1 2,09297 = (10 − 2,73329)/3,47195 | (10 − 4,54997)/2,09297 | 2,6039655689620536 | 2,6039655689620536 | 0 | APROVADO |
| Glicose | DIMENSION 1 | 8977 / pooled | N1 | corrigido 105/100 | rodada TESTE 5,0 | bias (8,19988 + 5)/4; ET 1,65·2,18933 + 3,29997; Sigma (10 − 3,29997)/2,09297 | 3,299970832473625; 6,912358614410101; 3,20120 | 3,299970832473625; 6,912358614410101; 3,201202182042403 | 0 / 0 / < 1e-9 | APROVADO |
| Glicose | — | — | — | + Controllab 200/100 mapeado | provedor ≠ AR (CAP) | não entra | sem mudança | sem mudança (0 células) | 0 | APROVADO |
| Glicose | DIMENSION 1 | 8977 | N1/N2 | Set/2026 · T3/2026 · histórico | contagem própria da tblCQ_Final | n, média, DP (n − 1), CV | n 23/24 · 55/56 · 90/88 | iguais | ≤ 1,5e-13 | APROVADO |
| Glicose | DIMENSION 1 | 8977 | N1 | T3/2026: trimestre dos dados brutos | linhas do trimestre | média das linhas (não a média dos meses) | 77,236 (a média dos meses daria 76,921) | 77,236 | — | APROVADO |
| Glicose | DIMENSION 1 | 8974 | N1/N2 | "Tudo (sem filtro)", janela de 180 | 219 RUNs com resultado que participa | as 180 corridas mais recentes; período inteiro | 39 cortadas; janela 140/131; período 174/161 | M1 39; engPainel 140/131; AF1 "N1 174 · N2 161" = Estatística 174/161 | 0 | APROVADO |
| Glicose | DIMENSION 1 | 8974 | N1/N2 | Out/2025 | contagem própria | n | 24/22, sem corte | M1 0, n 24/22 | 0 | APROVADO |
| Glicose | DIMENSION 1 | H3 vazio | N1/N2 | "Tudo (sem filtro)" | contagem própria | janela e período | 163 cortadas (+1 corrida só com X vermelho, que o motor também conta); janela 129/130; período 264/249 | M1 164; janela 129/130; AF1 "N1 264 · N2 249" | 0 (M1: diferença de definição explicada, `mEstatistica` D10) | APROVADO |
| PCR | DIMENSION 1 | 8353 | N1/N2 | período inteiro (ADR-091) | contagem da tblCQ_Final | n | N1 249 · N2 238 (N1 250 na parte 1 com o manual de teste MAN_0007) | "Período inteiro: N1 249 · N2 238" | 0 | APROVADO |
| (todos) | — | — | — | Principal A5 | tblCQ_Final da cópia | contagem | 35.906 linhas, 35.763 SIM | "Total: 35.906 · participam: 35.763" | 0 | APROVADO |

Conferência global: `calc_independente.py` sobre os retratos S_mes, S_trim, S0..S4 e G0..G4 (colunas C..O, T, V, W,
AC, AD, AF, AR; 80 linhas): 1.178 comparações por retrato, **0 divergências**. Isolamento: na inclusão e na correção
do CEQ, a comparação célula a célula mostrou que só mudam as linhas do analito (Proteína total 44–45; Glicose 38–39,
mais o resumo de classes B101:B102). Ao excluir tudo, A14:AU133 voltou a ser igual bit a bit (S4 = S0, G4 = G0).

### B.2 Hematologia: 24 cálculos, 24 aprovados (origem: Hematologia_F2.md, Relatório B; Hematologia_F1.md)

Entradas comuns: médias de |bias| por rodada (5 amostras cada): FH9-A 2,3593073054506553, FH9-B 1,1137368437415163,
FH9-C 0,8562947636618198. As 15 linhas da simulação (rodadas "1/2/3", Uso NAO) não entram. CVp: N1
0,8617391643040556, N2 0,7027095883994274, N3 0,5924957915732495. F (lote 522611, 2025): N1 1,091388640832613 (n 50),
N2 0,6959305974249244 (n 51), N3 0,6648898448424729 (n 51).

| Analito | Equipamento | Lote | Nível | Período | Dados de entrada | Fórmula encontrada | Resultado independente | Resultado QC_INI (Value2) | Diferença | Status |
|---|---|---|---|---|---|---|---|---|---|---|
| HGB | XN1000 | (bias não depende) | 1–3 | CEQ CAP 2025, rodada TODAS | 3 rodadas × 5 amostras | (2,35931 + 1,11374 + 0,85629)/3 | 1,4431129709513304 | 1,4431129709513304 | 0 | APROVADO |
| HGB | XN1000 | 522611 | 1 | ET 2025 | F 1,09139; G 1,44311 | 1,65·F + \|G\| | 3,2439042283251407 | 3,2439042283251416 | 3e-16 | APROVADO |
| HGB | XN1000 | pooled | 1 / 2 / 3 | Sigma 14/05–13/11/2026 | CVp; ETp 4 | (4 − \|G\|)/CVp | 2,96712408460 / 3,63861127165 / 4,31545179799 | 2,967124084592441 / 3,638611271652592 / 4,315451798019849 | ≤ 6e-12 | APROVADO |
| HGB | XN1000 | — | 1–3 | + TA-F2H-01 14,3/13 | + rodada TESTE 10,000000000000005 | (4,32934 + 10)/4 | 3,582334728213499 | 3,582334728213499 | 0 | APROVADO |
| HGB | XN1000 | 522611 | 1 / 2 / 3 | ET, idem | F | 1,65·F + \|G\| | 5,383125985587309 / 4,730620213964625 / 4,6794029722035795 | 5,383125985587311 / 4,730620213964624 / 4,6794029722035795 | ≤ 4e-16 | APROVADO |
| HGB | XN1000 | pooled | 1 / 2 / 3 | Sigma, idem | CVp | (4 − \|G\|)/CVp | 0,4846771379177 / 0,5943639857510 / 0,7049252968995 | 0,48467713791667294 / 0,5943639857509238 / 0,7049252969037589 | ≤ 6e-12 | APROVADO |
| HGB | XN1000 | — | 1–3 | bias com sinal, idem | amostras | média com sinal (T) | 3,512232107163426 | 3,512232107163426 | 0 | APROVADO |
| HGB | XN1000 | — | 1–3 | corrigido 13,65/13 | rodada TESTE 5,000000000000003 | (4,32934 + 5)/4 | 2,3323347282134983 | 2,3323347282134987 | 2e-16 | APROVADO |
| HGB | XN1000 | 522611 | 1 / 2 / 3 | ET, idem | F | 1,65·F + \|G\| | 4,133125985587308 / 3,480620213964624 / 3,429402972203579 | 4,13312598558731 / 3,480620213964624 / 3,429402972203579 | ≤ 5e-16 | APROVADO |
| HGB | XN1000 | pooled | 1 / 2 / 3 | Sigma, idem | CVp | (4 − \|G\|)/CVp | 1,9352320758606 / 2,3731927090748 / 2,8146449232295 | 1,9352320758563868 / 2,3731927090746896 / 2,8146449232463637 | ≤ 6e-12 | APROVADO |
| HGB | — | — | 1–3 | + CAP 2024 (+100 %) | ano ≠ N4 | não entra | sem mudança | sem mudança (0 células) | 0 | APROVADO |
| HGB | — | — | 1–3 | + Controllab 26/13 mapeado, R4 = CAP | provedor ≠ R4 | não entra | sem mudança | sem mudança (0 células) | 0 | APROVADO |
| HGB | XN1000 | 522611 | 1 / 2 / 3 | R4 = Controllab | 1 amostra, +100 % | G; 1,65·F + \|G\|; (4 − \|G\|)/CVp | G 100; H 101,80079125737382 / 101,14828548575113 / 101,09706824399008; L −111,40261923401 / −136,61404595127 / −162,02646730214 | G 100; H idem; L −111,40261923376998 / −136,61404595126515 / −162,026467303112 | ≤ 6e-12 | APROVADO (Sigma negativo: dúvida D-12) |
| HGB | — | — | 1–3 | tudo excluído | 3 rodadas × 5 | idem base | 1,4431129709513304 | 1,4431129709513304 (S4 = S0 bit a bit) | 0 | APROVADO |
| HGB | XN1000 | 6 lotes (533811 e 536411 fora pela tblLotesMaterial) | 1 / 2 / 3 | u(Rw) 12 m, 14/11/2025–13/11/2026, depois das regras ADR-091 | tblCQ_Final depois do ATUALIZAR | CV pooled (n ≥ 20) | 0,934995864944362 / 0,6888645385444551 / 0,5844415458004871 | 0,9349958649446772 / 0,6888645385452185 / 0,5844415457960841 | ≤ 8e-12 | APROVADO |
| HGB | XN1000 | 8 → 6 lotes | 1 | u(Rw): efeito da regra | antes e depois | CV pooled | 1,17396 → 0,93500 | AF N1 1,17396 → 0,93500 | < 1e-9 | APROVADO |
| RBC | XN1000 | — | 1–3 | F1 OC-0031: CEQ sem alvo | TA-55 sem alvo | não entra | bias inalterado | 1,533620125276759 (igual bit a bit) | 0 | APROVADO |
| WBC | XN1000 | H3 vazio (TODOS) | 1 / 2 / 3 | 2026, Ano inteiro, janela de 180 | corridas = par (RUN, LOTE) com resultado que participa (ADR-080) | as 180 corridas mais recentes; período inteiro | 278 cortadas; janela 120/158/112; período 374/400/340 | M1 278; engPainel 120/158/112; AF1 "N1 374 · N2 400 · N3 340" | 0 | APROVADO (contando só por RUN daria 272 e 121/159/113: um RUN com dois lotes conta como duas corridas) |
| WBC | XN1000 | 528211 | 1 / 2 / 3 | "Tudo (sem filtro)" | contagem própria | n | 83/86/81, sem corte | M1 0; n 83/86/81 = Estatística | 0 | APROVADO |
| WBC | XN1000 | 528211 | 1 / 2 / 3 | Nov/2025 | contagem própria | n | 45/49/45 | 45/49/45 | 0 | APROVADO |
| WBC | XN1000 | 536411 (SO_ANALITOS WBC) | 1 / 2 / 3 | Tudo | contagem dos que participam | n | 59/57/55 | 59/57/55 | 0 | APROVADO |
| (todos) | — | 533811 / 536411 | — | regras ADR-091 | tblCQ_Final depois do ATUALIZAR | contagem LOTE_* | 11.582 (6.481 + 5.101) | QA A13 = 11.582; A14 = 0 | 0 | APROVADO |
| (todos) | — | — | — | Principal A5 | tblCQ_Final (openpyxl) | contagem | 48.391 linhas, 47.709 SIM | "Total: 48.391 · participam: 47.709" | 0 | APROVADO |
| (28 analitos) | XN1000 | 522611 | 1–3 | retratos E0, S0, S1, S2, S3, RX, S4, R1 | cópias da produção aab0a54f464e | `calc_independente.auditar_hema` | 8 × 2.313 comparações | idem | 0 divergências | APROVADO |

Isolamento: a inclusão mudou 72 células, todas nas linhas 20–22 (HGB), e a coluna U ("ordem crítico", contagem
acumulada) de 14 linhas de outros analitos, renumerada porque o HGB entrou na lista de críticos. Nenhum número de
outro analito mudou. O analito novo "TESTE-AUDITORIA X" (sem ETp) mostrou "SEM EP" e "-" em K/L (ADR-068), e A14:AU97
ficou igual bit a bit antes e depois.

**Não recalculado neste ciclo:** o bias do RBC com alvo (3,650215093957569, Hema F1 OC-0031) é igual ao do reteste 03,
mas não foi recalculado de forma independente agora.

### B.3 Etapas para reproduzir

1. Copiar a produção para `%TEMP%\qcw\auditoria\<ciclo>\` e conferir o sha256 (Bio `8ec815ab653d`, Hema
   `aab0a54f464e`).
2. Rodar `_arquitetura/agentes/ferramentas/reteste04_usuario_f2.py` (Bio) ou `reteste04_usuario_f2_hema.py` (Hema).
   O script faz login com um usuário de teste, registra/corrige/exclui os CEQ TESTE-AUDITORIA pelo núcleo da janela e
   tira os retratos com `SaveCopyAs`.
3. `calc_independente.py` lê a EQA_Base e a tblCQ_Final de cada retrato com openpyxl e recalcula em Python puro:
   bias_i, as médias por rodada e das rodadas, F (CV do período), ET, CVp de 6 meses (lotes com n ≥ 20), Sigma e
   u(Rw). Depois compara com o `Value2` da Estatística e do Painel.
4. Os F1 recontam os n da janela de 180 corridas a partir da tblCQ_Final (corrida = RUN na Bio; par RUN × LOTE no modo
   TODOS da Hema, ADR-080).

### B.4 Caminho do dado (conferido nas trilhas e nos retratos)

- **CEQ → bias → ET → Sigma:**
  1. A janela do CEQ grava na EQA.CAP_Dados ou na EQA.Controllab_Dados.
  2. A EQA_Base é consolidada sozinha: Analito_Canonico pelo mapa W:Y, e Bias e |Bias| recalculados.
  3. `BiasEQ` (provedor AR/R4, ano N4, rodada P4) alimenta a Estatística G/T/AC/AD.
  4. Daí saem H (ET, com o F do período ou lote), L (Sigma, com o CVp), M/P (classes), U (ordem dos críticos), V/W
     (DPM, rendimento), X..AA (plano), AJ..AU (incerteza e viés) e K5.
  5. O Painel B7:I9 mostra o analito do C3, e a Bio também tem o resumo de classes B101:B102.
- **CIQ → CV → ET e Sigma:** a tblCQ_Final (PARTICIPA = SIM, já com as regras ADR-091 aplicadas) alimenta o Eng_Saida
  (engEstat F, engPainel) e o `IncertezaCIQ` (CVp, u(Rw)).
- Os valores aparecem sem `Calculate` explícito (o resultado é o mesmo com e sem ele).

### B.5 Diferenças de método e observações (não reprovadas)

- **Sigma negativo** quando |bias| > ETp (Hema, R4 = Controllab): a conta segue a definição e a calculadora confere.
  Fica a dúvida de exibição D-12 para o Especialista em CQ.
- **M1 da Bio com H3 vazio** (164 contra 163): o motor conta como corrida uma corrida que só tem X vermelho
  (`mEstatistica` D10). Os n da janela e do período batem.
- **Corrida no modo TODOS da Hema:** é o par (RUN, LOTE), conforme o ADR-080. Com essa definição o oráculo bate.
- **ET do Painel × Estatística (Bio Glicose):** diferença de 3e-14 relativo, porque o Painel recalcula o DP.

---

## Relatório C: ocorrências

**Deduplicação.** As frentes trouxeram 11 ocorrências. Ficaram 10: na Bio, o R4-02 (F1) e o R4-F2-03 (F2) são a mesma
tela na mesma condição ("Tudo (sem filtro)" no cabeçalho da Estatística) e viraram o P-09. Falhas iguais nos dois
produtos **não** foram mescladas. A coluna "Outro produto" aponta o par, porque em vários casos uma correção comum
cobre os dois. A ordem é por prioridade, depois produto (Bio, Hema). Os IDs P-xx são provisórios. O ID definitivo
(OC-xxxx) é do Orquestrador. Nenhuma ocorrência é Crítica. O Bio F1 pergunta se o P-01 deveria ser Crítico (D-04).

### C.1 Tabela

| ID | Produto | Prioridade | Tipo | Tela/módulo | Resumo | Origem | ID anterior | Outro produto |
|---|---|---|---|---|---|---|---|---|
| P-01 | Bioquímica | **Alto** | `funcional` | Digitar Resultados / validação de dados (`camada_dados.validacoes`, linhas 532–537) | As listas de EQUIPAMENTO, MATRIZ e MOTIVO têm um item só (separador "," num Excel pt-BR): o valor certo digitado é recusado, e o item do menu grava um equipamento inexistente que entra ATIVO, sem QA | Bio F1 R4-01 | — (vizinho da OC-0019) | P-02 |
| P-02 | Hematologia | **Alto** | `funcional` | Digitar Resultados / `camada_dados.validacoes` (linha 537) | A lista do MOTIVO tem um item só: cada um dos 4 motivos é recusado, e o motivo de substituição (ADR-082) não pode ser registrado certo | Hema F1 VIZ-LISTAS | — | P-01 |
| P-03 | Bioquímica | Médio | `usabilidade` | frmCEQ / `mCEQForm.AvisoAnalitoCEQ`, `SalvarCEQ` | Nome do Controllab mapeado para um analito cujo provedor (AR) é o CAP aparece como "entra no bias", e não entra (latente: o mapa do Controllab está vazio) | Bio F2 R4-F2-01 | — | P-06 |
| P-04 | Bioquímica | Médio | `usabilidade` | Painel J7:J8, Q4, A39 (ADR-091) | Lote com regra FORA_DO_CQ ou SO_ANALITOS: "SEM CORRIDA / SEM MÉD/DP", "digite a Média e o DP", "o lote não tem resultado", sem citar a regra | Bio F2 R4-F2-02 | — (parente da OC-0061) | P-07 |
| P-05 | Hematologia | Médio | `usabilidade` | Painel A39 × aba Estatística | O aviso da janela de 180 (modo TODOS) manda olhar a Estatística, que mostra outro recorte (lote 522611, 2025): 374/400/340 não aparecem em lugar nenhum | Hema F1 R4-H01 | OC-0012 (EFEITO_COLATERAL) | — (na Bio confere) |
| P-06 | Hematologia | Médio | `usabilidade` | frmCEQ / `mCEQForm.AvisoAnalitoCEQ`, `SalvarCEQ` | Com o R4 = CAP, um nome do Controllab mapeado aparece como "entra no bias", e não entra (latente) | Hema F2 R4-F2-01 | — | P-03 |
| P-07 | Hematologia | Médio | `usabilidade` | Painel J7:J9, Q4, A39 (ADR-091) | Lote FORA_DO_CQ: "SEM CORRIDA" e "O lote não tem resultado de WBC", embora tenha 214 (latente: a tabela está vazia na produção) | Hema F2 R4-F2-02 | — | P-04 |
| P-08 | Bioquímica | Baixo | `usabilidade` | Digitar Resultados (DETALHE) / `mEtlCQ` MOTIVO_EXCLUSAO_AUTOMATICA | O DETALHE do IMPLAUSIVEL e a nota do ADR-093 continuam sem acento ("implausivel", "estatistica", "grafico", "lancado") | Bio F1 (OC-0021 PERSISTE) | OC-0021 | P-10 |
| P-09 | Bioquímica | Baixo | `usabilidade` | Estatística A4:J5 | Com o Painel em "Tudo (sem filtro)": B5/D5 "00/01/1900" e D4 "Ano 2026 (refletido do Painel)", ao lado de H5 "histórico inteiro do lote" | Bio F1 R4-02 + Bio F2 R4-F2-03 (mesclada) | — (vizinho da OC-0040) | não testado (a Estatística da Hema tem filtro próprio) |
| P-10 | Hematologia | Baixo | `usabilidade` | Digitar Resultados (DETALHE) / `mEtlCQ` TXT_IMPLAUSIVEL | O DETALHE do IMPLAUSIVEL continua sem acento | Hema F1 (OC-0021 PERSISTE) | OC-0021 | P-08 |

### C.2 Detalhe

**P-01 (Bio, Alto, funcional) — listas de validação com um item só**
- **Ação:** depois do âmbar da OC-0019, corrigir o EQUIPAMENTO digitando "DIMENSION 1" (GUIA §3: "confira (ou digite
  o certo)") ou escolher no menu. O mesmo vale para escolher MATRIZ ou MOTIVO.
- **Esperado:** listas com DIMENSION 1 e DIMENSION 2, SORO e URINA, e os 4 motivos, aceitando qualquer um.
- **Observado:**
  - `Validation.Formula1` = "DIMENSION 1,DIMENSION 2", estilo Parar, mensagem "Escolha o equipamento da lista.".
  - `Validation.Value`: "DIMENSION 1" Falso, "DIMENSION 2" Falso; só o texto inteiro dá Verdadeiro. Idem "SORO,URINA"
    e os 4 motivos. LOTE e ANALITO (`=lstLotes`, `=lstAnalitos`) estão certos.
  - Manual com o item do menu → MAN_0001 ATIVO, PARTICIPA SIM, RUN 26101401, equipamento "DIMENSION 1,DIMENSION 2",
    DETALHE vazio, nenhum achado no QA, fora das duas séries do Painel.
- **Dificuldade:** o usuário vê o âmbar, tenta corrigir e é barrado. A única opção do menu parece certa.
- **Consequência:** a correção que a OC-0019 pede fica impossível por digitação, e o MOTIVO (obrigatório no
  SUBSTITUI_ID) não pode ser escolhido. O menu induz a gravar um equipamento inexistente que participa da base sem
  aviso (agrava o G02 do ciclo 01). É a mesma causa do G18/G32: separador gravado pelo COM e lido pelo idioma do Excel.
- **Sugestão:** criar as listas com o separador do idioma (`Application.International(xlListSeparator)`) ou, melhor,
  apontá-las para um intervalo/nome; reinstalar nas duas produções; adicionar um QA para EQUIPAMENTO fora da lista do
  produto.
- **Evidência:** `Bioquimica_F1_val.json` (`listas`, `consequencia`), `Bioquimica_F1_comp.json` (`O19_validacao`),
  `Bioquimica_F1.json` (`O19_validacao_DIMENSION1` = false); produção 8ec815ab653d.
- **Não executado:** a digitação física (sem interação com a tela). A avaliação é a do próprio Excel
  (`Range.Validation.Value`), que é o critério que a caixa Parar aplica.

**P-02 (Hema, Alto, funcional) — lista do MOTIVO com um item só**
- **Ação:** digitar ou escolher o MOTIVO de um manual (obrigatório para substituir, ADR-082).
- **Esperado:** lista com os 4 motivos, aceitando qualquer um.
- **Observado:** `Formula1` = "Falha do interfaceamento,Resultado não transmitido,Equipamento em contingência,Outro",
  estilo Parar, "Escolha o motivo da lista.". `Validation.Value` dá Falso para cada motivo. EQUIPAMENTO ("XN1000") e
  MATRIZ ("SANGUE TOTAL") estão certos porque têm um item só.
- **Consequência:** o motivo da substituição não pode ser registrado certo pela tela. Quem usa o menu grava um texto
  sem sentido no Audit_Log e no I04. Não muda estatística.
- **Sugestão:** a mesma do P-01.
- **Evidência:** `Hematologia_F1.json` (`listas_validacao.MOTIVO`); `camada_dados.py:537`; produção aab0a54f464e.

**P-03 / P-06 (Médio, usabilidade) — aviso do CEQ ignora o provedor do analito**
- **Ação:** com um nome do Controllab mapeado para Glicose (Bio) ou HGB (Hema), registrar uma amostra do Controllab
  enquanto o AR/R4 é CAP.
- **Esperado:** aviso de que o resultado fica gravado **fora** do bias, porque o analito usa o CAP.
- **Observado:** "AZUL|No QC_INI: Glicose/HGB. Este resultado entra no bias deste analito." e `SalvarCEQ` "OK|n|" sem
  nota. A Estatística não muda (0 células), o que está certo: CAP e Controllab não se misturam.
- **Consequência:** o analista acha que o CEQ do Controllab já pesa no bias, no ET e no Sigma. Hoje é latente: os
  mapas do Controllab estão vazios, e qualquer nome cai no aviso âmbar correto.
- **Sugestão:** comparar o provedor digitado com o AR (Bio) ou o R4 (Hema) no aviso e no retorno do Salvar.
- **Evidência:** `Bioquimica_F2_p2.json` (`ceq_glicose.aviso_ctl`, `G3_controllab`, `diff_G2` = []);
  `Hematologia_F2.json` (`ceq.aviso_ctl_mapeado`, `ceq.S3_vizinhos`, `ceq.RX_controllab`);
  `src/comum/mCEQForm.bas:291-309`.

**P-04 / P-07 (Médio, usabilidade) — Painel em lote com regra da tblLotesMaterial**
- **Ação:** Bio: H3 8794 (FORA_DO_CQ) ou 8353 (só PCR) com Glicose. Hema: regra de teste 533811 FORA_DO_CQ, WBC,
  "Tudo (sem filtro)".
- **Esperado:** nenhum ponto (correto) e um texto que diga "lote fora do CQ / só PCR (tblLotesMaterial)", sem pedir
  Média/DP.
- **Observado:** Bio J7 "SEM CORRIDA", J8 "SEM MÉD/DP", A39 "…digite a Média e o DP no bloco do lote… O lote não tem
  resultado de Glicose no DIMENSION 1"; cada lote tem 31 resultados LOTE_FORA_DO_CQ ou LOTE_OUTRO_MATERIAL. Hema J7:J9
  "SEM CORRIDA", Q4 "SEM CORRIDA DESTE ANALITO", A39 "O lote não tem resultado de WBC no XN1000."; o lote tem 214.
- **Consequência:** o usuário procura um resultado "que não chegou" ou digita Média/DP num lote que não é de CQ. Não
  muda estatística.
- **Sugestão:** com regra ativa para o lote, o A39/Q4 dizem o motivo da regra.
- **Evidência:** `Bioquimica_F2_p2.json` (`painel_lotes.2_Glicose_com_8353`, `3_lote_8794_FORA_DO_CQ`);
  `Hematologia_F2_p2.json` (`painel_lotes.WBC_533811`).

**P-05 (Hema, Médio, usabilidade) — efeito colateral da OC-0012**
- **Ação:** WBC, H3 vazio, K3 2026 + Ano inteiro; abrir a Estatística, como o aviso manda.
- **Esperado:** achar o n do período inteiro citado (374/400/340), como na Bio.
- **Observado:** a Estatística da Hema tem filtro próprio (Ano 2025–2025, lote B4 522611 "refletido do Painel", com o
  H3 do Painel vazio) e mostra WBC n 50/51/51.
- **Consequência:** o aviso aponta para um número que não confere, justamente com o Painel em REJEITADO.
- **Sugestão:** na Hema, trocar o fim do texto ("escolha um lote no H3 para ver a estatística do período") ou omiti-lo
  com o H3 vazio. O AF1 está certo.
- **Evidência:** `Hematologia_F1.json` (`O12.original_WBC_TODOS_2026`, `O12.estatistica_WBC`),
  `Hematologia_F1_comp.json` (`tela`).

**P-08 / P-10 (Baixo, usabilidade) — OC-0021 PERSISTE em parte**
- **Resolvido:** o RUN fica vazio nos MANUAL_INCOMPLETO e CONFLITO, o DETALHE fica vazio no ATIVO, não há nenhum "0"
  nas 300 linhas, e "campos obrigatórios ausentes/inválidos" está acentuado.
- **Persiste:** Bio MAN_0005 "resultado <= 0: implausivel para material de controle (fora da estatistica e do
  grafico)" e MAN_0007 "… lancado como 'CRE2' …"; Hema MAN_0004, o mesmo texto do IMPLAUSIVEL.
- **Causa:** o ADR-090 acentua só uma lista fixa de palavras (SUBSTITUTE).
- **Sugestão:** acentuar na origem nos dois motores (VBA e PQ) ou completar a lista (D-02).
- **Evidência:** `Bioquimica_F1.json` (`manuais.manuais_1`); `Hematologia_F1.json` (`manuais.varredura.sem_acento`).

**P-09 (Bio, Baixo, usabilidade) — cabeçalho da Estatística em "Tudo (sem filtro)"**
- **Observado:** B5 `=Painel!G3` e D5 `=Painel!G4` mostram "00/01/1900". O K3 continua 2026 e o D4 diz "Ano 2026
  (refletido do Painel)", ao lado de H5 "histórico inteiro do lote". O número é do histórico inteiro (n 90/88
  conferido).
- **Sugestão:** `=IF(Painel!G3="","",Painel!G3)` (e G4), como na D4 do ADR-092, e D4 vazio sem De/Até.
- **Evidência:** `Bioquimica_F1.json` (`O12.estatistica_cabecalho`), `Bioquimica_F1_comp.json` (`estat_B5_D5_tudo`),
  `Bioquimica_F2.json` (`O40.3_tudo`).

### C.3 Dúvidas de regra de negócio (para o Orquestrador encaminhar)

| ID | Produto | Para | Dúvida |
|---|---|---|---|
| D-01 | Bio | Especialista em CQ / RT | O manual com o equipamento automático mantido entra ATIVO (o âmbar é só aviso). O EQUIPAMENTO obrigatório na Bio (item 3 do G01) segue sem decisão |
| D-02 | Ambos | Orquestrador | Acentuar o MOTIVO_EXCLUSAO_AUTOMATICA na origem (VBA e PQ juntos) ou manter a lista fixa de acentuação atualizada a cada texto novo (ADR-088, ADR-093)? |
| D-03 | Ambos | UX | O Z3 diz "clique em Consolidar na EQA_Base" e a trilha diz "Consolidar na EQA_Base ou ATUALIZAR DADOS": unificar? |
| D-04 | Bio | Orquestrador | P-01 deveria ser Crítico (associação a equipamento inexistente com PARTICIPA SIM e sem QA) ou Alto (depende de o usuário usar o menu)? As frentes registraram Alto |
| D-05 | Bio | Especialista em CQ / RT | Urina sem tela: o manual aceito como CRE2 e a série de urina participam, mas nenhuma tela os mostra. Onde o laboratório acompanha o CQ de urina? (OC-0061, ADR-091) |
| D-06 | Ambos | Orquestrador | O aviso do CEQ deve considerar o AR (Bio) / R4 (Hema), dizendo "entra se o R4 for Controllab", ou o mapa do Controllab deve ser recusado enquanto o provedor for o CAP? |
| D-07 | Ambos | Orquestrador (BI fora do escopo) | A BI_Data é visível a todos depois do login e mostra "####" e datas aaaa-mm-dd: ocultar para quem não é ADM? |
| D-08 | Bio | Orquestrador | ADR-093, código ambíguo: não há caso real na Bio. Documentar que hoje todo analito do catálogo tem um código só? |
| D-09 | Bio | Especialista em CQ / RT | Um manual "Glicose" em lote 635x vira GLUC (urina) pelo de/para atual: é o esperado? |
| D-10 | Hema | UX | A instrução A3 (texto comum) promete "fica em âmbar", mas na Hema o equipamento entra sem âmbar (ADR-090): um texto por produto? |
| D-11 | Hema | Especialista em CQ | Um manual IMPLAUSIVEL (−1) recebeu o RUN 26100901 e, pela regra (ADR-088, `mEtlCQ.bas:958`), abre corrida e renumera as do dia, mesmo fora da estatística. É o desejado? (A renumeração dos reais não foi medida) |
| D-12 | Ambos | Especialista em CQ | Sigma negativo (\|bias\| > ETp): mostrar o número negativo ou "< 0" / "ETp excedido pelo bias"? |
| D-13 | Hema | UX | Lote SO_ANALITOS: o Painel troca o analito sozinho (RBC → WBC). Manter, ou ficar no RBC e explicar a regra? |
| D-14 | Hema | UX | O Painel no modo TODOS dá REJEITADO sem mostrar Média, DP, CV e ET (ADR-086). Explicar de onde vem o status? |
| D-15 | Ambos | UX | O GUIA §5 diz "o botão Consolidar da EQA_Base", mas o botão fica nas abas de dados. A numeração do GUIA está fora de ordem (§6, §7, §6b, §6c, §7b, §8) |

---

## Relatório D: parecer final

Critério: "validado" só onde houve teste executado e aprovado neste ciclo. O resto é "não validado" (teste reprovado)
ou "não executado" (com o motivo).

### D.1 A rotina funciona?

**Sim no que foi refeito, com uma ressalva Alta na entrada manual.**

- **Validado nos dois produtos:**
  - CEQ pela janela: registro, correção, exclusão, duplicidade, sessão, sem alvo, listas, descrição e trilha.
  - Manuais: incompletos, em conflito, implausíveis, completados depois, com trilha e QA certos.
  - ATUALIZAR: idempotente, sem trilha repetida.
  - Regras de lote (ADR-091), datas (ADR-094) e o Principal A5.
  - 25 dos 27 itens do reteste RESOLVIDOS.
- **Não validado:**
  - Corrigir o EQUIPAMENTO (Bio) ou registrar o MOTIVO (Bio e Hema) pela tela é impossível (P-01, P-02, Alto).
  - O manual lançado pelo item do menu entra ATIVO com um equipamento inexistente e sem QA.
- **Persiste:** OC-0021 (acento, Baixo).
- **Efeito colateral:** OC-0012 na Hema (P-05, Médio).

### D.2 Bias, ET e Sigma estão validados?

**Validado** nos dois produtos para os casos executados:
- **Bio:** Proteína total e Glicose (DIMENSION 1, lote 8977).
- **Hema:** HGB (XN1000, lote 522611); o caso extremo R4 = Controllab com bias de 100 %.
- **Inclusão e correção de CEQ:** mudam bias, ET e Sigma só do analito certo. A diferença máxima foi 6e-12 relativo
  no Sigma e 4e-16 no bias e no ET.
- **Exclusão:** volta a Estatística igual bit a bit.
- **Calculadora:** 32.640 comparações, 0 divergências.
- **Seleção das amostras:** ano errado (CAP 2024) e provedor errado (Controllab com AR/R4 = CAP) não entram.
- **ETp ausente aparece como "-":** validado na **Hema** (analito de teste sem ETp: "SEM EP" e "-").
- **Não executado neste ciclo:**
  - ETp ausente na Bio. A frente não provocou o caso, e nenhum analito da Bio está sem ETp no ano de referência.
  - ETp de fonte diferente de CLIA (VB/FAB): os analitos refeitos são todos CLIA. O caso foi validado no ciclo 01,
    mas não neste.
- **Pendente de método:** D-12 (exibição do Sigma negativo).

### D.3 A integração CIQ + CEQ está validada?

**Validada nos casos executados:**
- Glicose 8977: o ET usa o CV do período do CIQ e o bias do CEQ, e muda com a inclusão e a correção do CEQ.
- HGB: o ET usa o F do lote 522611 e o Sigma usa o CVp de 6 meses.
- O u(Rw) do HGB reage às regras ADR-091 (8 → 6 lotes), igual ao independente.
- O caminho do dado (B.4) foi conferido nas trilhas e nos retratos.

**Ressalva de usabilidade (não muda número):** o aviso da janela do CEQ ignora o provedor do analito (P-03, P-06).
Hoje é latente.

### D.4 Os filtros de período e de lote estão corretos?

**Validado nos casos executados:**
- **Cálculos:** mês (Set/2026, Out/2025, Nov/2025), trimestre (T3/2026, dos dados brutos: 77,236 e não a média dos
  meses, 76,921), ano, "Tudo (sem filtro)", De/Até atravessando o ano (sem "Ano 0") e K3/K4 "Personalizado".
- **Lotes:** 8974, 8977, 8353, 528211, 536411 e H3 vazio nos dois produtos.
- **Janela de 180 corridas:** o n da janela e o do período batem com o oráculo nos dois produtos.
- **Não validado:**
  - A apresentação do cabeçalho da Estatística da Bio em "Tudo" (P-09, Baixo; o número está certo).
  - O ponteiro do aviso da Hema para a Estatística (P-05, Médio).
- **Não executado neste ciclo:** a F4 completa (Levey-Jennings, Westgard, corrida do lote, troca de equipamento na
  Bio). Não fazia parte deste reteste. A última validação é a do ciclo 01 e das entregas.

### D.5 A entrada manual e a automática estão validadas?

- **Manual:**
  - **Validado:** status, DETALHE, RUN, trilha (OC-0023, OC-0058), QA/ResumoQA (OC-0024), urina (ADR-093/OC-0056) e
    regras de lote, nos dois produtos.
  - **Não validado:** as listas de EQUIPAMENTO, MATRIZ e MOTIVO (P-01, P-02).
  - **Não executado:** a digitação física numa célula com validação (avaliada pelo `Validation.Value`, sem interação
    com a tela) e o clique Sim/Não no frmCEQ.
- **Automática:**
  - **Validado:** o ATUALIZAR idempotente e a reclassificação das linhas existentes pelas regras ADR-091 (0 violações
    no oráculo, A13 = total LOTE_*).
  - **Não executado neste ciclo:** o recebimento de resultado **novo** do DB_SEAC. Nenhum chegou durante os testes
    (últimos: Bio 13/11/2026 08:39:01, Hema 13/11/2026 13:10:20).
  - **Não executado neste ciclo:** a equivalência manual × automático (fora da lista do reteste; validada no ciclo
    01).

### D.6 Há divergência matemática?

**Não.**
- 47 cálculos do Relatório B aprovados.
- 32.640 comparações da calculadora com 0 divergências. A maior diferença foi 8e-12 relativo (u(Rw), ordem de soma do
  CV pooled), dentro de 1e-9.
- As diferenças de contagem (M1 164 × 163 na Bio; RUN × par RUN/LOTE na Hema) são de definição, estão explicadas no
  código e no ADR-080, e os n batem.

### D.7 Há problema operacional?

**Sim, um Alto nos dois produtos:** P-01 e P-02, a entrada manual com listas de validação quebradas. O laboratório
precisa saber disso antes de usar o SUBSTITUI_ID ou corrigir o equipamento.

**Médios:**
- P-05 (aviso da Hema que leva a outro recorte).
- P-03/P-06 e P-04/P-07: latentes enquanto o mapa do Controllab e, na Hema, a tblLotesMaterial estiverem vazios. Na Bio,
  o P-04 já acontece com os lotes 8794 e 8353.

**Outros:**
- O 1º ATUALIZAR da Hema depois de uma regra nova levou 74 s.
- A BI_Data mostra "####" (D-07).

As produções ficaram intactas: mesmo hash antes e depois de todas as execuções.

### D.8 O que fica pendente

1. Corrigir e retestar **P-01 e P-02** (Alto) nas duas produções. No reteste, digitar cada item e conferir que um
   EQUIPAMENTO fora da lista é apontado pelo QA.
2. Decidir e corrigir os Médios **P-03 a P-07** e os Baixos **P-08 a P-10**. A OC-0021 continua aberta (P-08/P-10). A
   OC-0012 fica com efeito colateral na Hema (P-05).
3. Levar ao Especialista em CQ e ao RT: D-01, D-05, D-09, D-11 e D-12. Ao UX: D-03, D-10, D-13, D-14 e D-15. Ao
   Orquestrador: D-02, D-04, D-06, D-07 e D-08.
4. **Não executados, a cobrir num próximo ciclo:**
   - recebimento de resultado novo do DB_SEAC;
   - equivalência manual × automático;
   - ETp ausente na Bio;
   - ETp VB/FAB;
   - regra MATRIZ na Hema (não há caso real);
   - código ambíguo do ADR-093 (não há caso real);
   - clique nas caixas Sim/Não do frmCEQ;
   - digitação física com validação;
   - a F4 completa (LJ, Westgard, corrida do lote, equipamentos).
5. GUIA: ordem das seções, local do botão Consolidar, o atalho do equipamento vazio enquanto o P-01 não é corrigido e
   o aviso do 1º ATUALIZAR mais lento depois de uma regra nova.

Dados pessoais: o Audit_Log das cópias grava nome, login do Windows e máquina de quem roda. Esses dados não constam
deste parecer nem dos relatórios das frentes.
