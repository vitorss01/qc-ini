# Parecer consolidado: ciclo 2026-11-15_reteste05_parte2

Agente: Usuário Simulado (parecer consolidado). Data: 15/11/2026. Modo: reteste, consolidação.

**De onde vêm os dados.**
- Neste passo o Excel não foi aberto e nenhum número foi recalculado.
- Tudo vem dos 2 relatórios das frentes desta pasta (`Bioquimica_F1.md`, `Hematologia_F1.md`) e do resumo
  estruturado que cada frente devolveu.
- Para confirmar as linhas de origem citadas, li `src/comum/mEtlCQ.bas` só para leitura: 808, 816 e 1244.

| Item | Valor |
|---|---|
| Produção auditada (entrega 42) | `QC_Bioquimica.xlsm` sha256 curto **b20fe00f7e51**; `QC_Hematologia.xlsm` sha256 curto **7bab8fc008f6**. Nas duas frentes, o hash foi igual antes e depois: a produção não foi aberta para escrita |
| Base de comparação (entrega 41) | Bio: `_backup_pre_integracao_2026-11-15_0011`, sha256 curto **8ec815ab653d**. Hema: `_backup_pre_integracao_2026-11-15_0458`, sha256 curto **aab0a54f464e**. Nas duas, hash igual antes e depois |
| Lista do reteste | 3 itens nos dois produtos: **OC-0021** (DETALHE sem acento), **OC-0068** (listas de escolha com um item só) e **OC-0072** (ano literal "yyyy" nos formatos de data). Os três foram tratados no ADR-095 |
| Frentes executadas | Só a F1, nos dois produtos, em modo reteste. **Não fizeram parte deste reteste:** a F2 (CEQ), a F3 (auditoria independente de bias, ET e Sigma) e a F4 completa (Levey-Jennings, Westgard, corrida do lote, trimestres) |
| Caminho | O do usuário: login, filtros do Painel com os eventos ligados, navegação do botão (`mIntegracao.IrDigitarResultados`), manuais digitados célula a célula com `FormulaLocal` e eventos ligados, ATUALIZAR DADOS pelo núcleo do botão (`mIntegracao.AtualizarDadosAutomatico`) e filtros H4 e De/Até do Registros digitados. A Hema também colou uma linha inteira (`Range.Copy`) |
| Fora do caminho do usuário (declarado) | Hema: 2 regras de lote cadastradas só na cópia, com NOTA `TESTE-AUDITORIA`: 602911 FORA_DO_CQ e 522611 SO_ANALITOS = RBC. Foi preciso porque a Hema não tem regra na `tblLotesMaterial`. Nas duas frentes, o texto "Outro: TESTE-AUDITORIA" no MOTIVO foi gravado por automação, por cima da validação; pela tela, o usuário não consegue gravá-lo |
| Dados de teste | 17 linhas manuais por produto (MAN_0001..), valores sintéticos 123,456, 111,111, 12,345, 1,234 e −1, e SUBSTITUI_ID 999999999 (inexistente). Usuários `TESTE_AUDITORIA_R5P2` e `TESTE_AUDITORIA_R5P2H` (ANALISTA) criados só nas cópias. Nenhuma cópia foi salva sobre si mesma, e todas foram apagadas no fim. O único `SaveCopyAs` foi feito na pasta do ciclo, só para ler o XML depois do uso. Ficam os JSON e os logs em `%TEMP%\qcw\auditoria\2026-11-15_reteste05_parte2\` |
| Excel | Um por vez, aberto e encerrado pelo `qa_final.QA`. Bio: PIDs 30212 (tentativa interrompida por erro do script antes de qualquer leitura; cópia descartada), 24532 (base) e 12032 (nova). Hema: 16116 (base) e 29532 (nova). Os Excel da tarefa agendada (13084 e os que ela abriu durante o trabalho) não foram tocados. Nenhuma captura de tela |
| Ferramentas | Em `_arquitetura/agentes/ferramentas/`, documentadas no cabeçalho: `reteste05p2_usuario_f1.py` (Bio), `reteste05p2_usuario_f1_hema.py` (Hema) e `xml_formatos_listas.py` (lê o XML direto, sem Excel nem openpyxl) |
| tblCQ_Final | Bio: 35.906 linhas. Hema: 48.391 linhas. Iguais na base e na nova. Os ATUALIZAR leram o DB_SEAC só para leitura e não trouxeram resultado novo. Último recebido: Bio 13/11/2026 08:39:01; Hema 13/11/2026 13:10:20 |

**Resumo em números**

| | Bioquímica | Hematologia | Total |
|---|---|---|---|
| Itens do reteste: RESOLVIDA / PERSISTE / EFEITO_COLATERAL | 2 / 1 / 0 | 2 / 1 / 0 | 6 / 2 / 0. Por OC, nos dois produtos: OC-0068 e OC-0072 resolvidas; OC-0021 persiste com resíduo Baixo |
| Testes vizinhos: APROVADO / REPROVADO / NAO_EXECUTADO | 3 / 0 / 1 (caixas de aviso e clique físico) | 3 / 0 / 1 (a mesma limitação, declarada no relatório) | 6 / 0 / 2 |
| Linhas manuais lançadas e conferidas (base × nova) | 17 | 17 (16 digitadas e 1 colada) | 34 |
| Diferenças base × nova em STATUS, MOTIVO_EXCLUSAO_AUTOMATICA, RUN e tblCQ_Final dos manuais | 0. A exceção é o EQUIPAMENTO do MAN_0001, que é o efeito esperado da OC-0068 | 0 | 0 não esperadas |
| Diferenças de valor `<v>` no XML, base × nova, nos alvos da OC-0072 | 0. A única diferença é a linha nova do Sync_Log, gravada pelo ATUALIZAR da entrega | 0 (idem) | 0 |
| Cálculos no Relatório B (aprovados / com ressalva / reprovados) | 3 (2 / 1 / 0) | 3 (2 / 1 / 0) | 6 (4 / 2 / 0) |
| Ocorrências das frentes / deduplicadas (P-xx) | 2 / — | 3 / — | 5 / 3 |
| Críticas / Altas / Médias / Baixas (deduplicadas) | — | — | 0 / 0 / 1 / 2 |
| Dúvidas de regra de negócio (deduplicadas) | 3 | 3 | 4 |

**Situação dos 3 itens do reteste**

| Item | Bio | Hema | Consolidado |
|---|---|---|---|
| OC-0021: DETALHE sem acento | **PERSISTE** (resíduo). Os 9 textos sem acento da base estão acentuados. Sobra "MAN_0002 **e** resultado manual" | **PERSISTE** (resíduo). Sobram "MAN_0002 **e** resultado manual" e "HEM-212660 **e** retransmissão de HEM-212425", os dois provocados | **PERSISTE**, com resíduo Baixo (P-01). Status e números não mudaram |
| OC-0068: listas de escolha com um item só | **RESOLVIDA**. EQUIPAMENTO, MATRIZ, MOTIVO, Registros H4 e Painel L4 com um item por valor. O "DIMENSION 2" do menu entra na série certa. 0 `x12ac:list` | **RESOLVIDA**. MOTIVO e H4 com um item por valor. O motivo escolhido vai certo para a trilha do substituído. 0 `x12ac:list` | **RESOLVIDA**. Resíduo à parte: a recusa do H4 e do L4 usa a mensagem genérica do Excel (P-02, Baixo) |
| OC-0072: ano literal nos formatos de data | **RESOLVIDA**. 0 células com "yyyy" e 0 numFmt literais no XML, antes e depois do uso. Valores iguais | **RESOLVIDA**. Idem, inclusive nas 5.000 datas do Eventos_Westgard | **RESOLVIDA**. Fica uma dúvida sobre o formato interno 14 (D-01) |

---

## Relatório A: experiência do usuário

### A.1 Funcionalidades testadas

| Frente | Bioquímica | Hematologia |
|---|---|---|
| F1: entrada manual (Digitar Resultados) | 17 manuais que geram cada texto possível de MOTIVO_EXCLUSAO_AUTOMATICA. Os casos: IMPLAUSIVEL; LOTE_FORA_DO_CQ 8177; SO_ANALITOS 8353; MATRIZ URINA 6354 sem de/para; troca para CRE2 (ADR-093); SUBSTITUI_ID válido, já substituído, inexistente, de manual, de outra chave e sem MOTIVO; data futura; conflito; repetido; sem NIVEL | 17 linhas com os mesmos cenários, exceto a troca para CRE2 (só existe na Bio). Mais duas: SUBSTITUI_ID de uma retransmissão (DUPLICIDADE_ORIGEM) e uma linha inteira colada |
| F1: ATUALIZAR DADOS | Mensagem com os grupos E04/E12 e governança; S1 SUBSTITUIDO_MANUAL; tblCQ_Final dos manuais | Idem. Os 5 E12 da QA_INTEGRACAO batem com a mensagem |
| F1: listas de escolha | EQUIPAMENTO, MATRIZ e MOTIVO da Digitar Resultados; Registros H4; Painel L4. Itens válidos e inválidos avaliados pelo `Validation.Value`; mensagens de erro; XML | MOTIVO, EQUIPAMENTO e MATRIZ; Registros H4 e F7. Idem. O Painel L4 não existe na Hema |
| F1: datas | Configuração D/AD, Eng_Saida AC, Registros E4/F4, Westgard_Fotos F, Sync_Log, Estatística E8:F9/C10 e LotesStore U1: Text, NumberFormat e XML, depois do login e depois do ATUALIZAR | Os mesmos alvos, mais o Eventos_Westgard A4:A5003 |
| F1: filtros do Registros | H4 Sim/Não/vazio e De/Até: só mudam a vista (hash do Value2 da tblRegistros constante) | Idem |

### A.2 Facilidade de uso

**Listas.**
- A seta mostra uma opção por linha, e digitar o valor certo funciona, como o GUIA §3 promete ("escolha na seta da
  célula").
- Na Bio, corrigir o equipamento âmbar ficou possível sem truque. O manual com "DIMENSION 2" entra na série do
  DIMENSION 2, e a média muda o quanto deve mudar.
- Na Hema, o MOTIVO escolhido ("Resultado não transmitido") vai certo para o resultado substituído e para a trilha.
  Antes, a trilha gravava os 4 motivos colados.

**DETALHE.**
- O texto ficou legível e acentuado (implausível, estatística, gráfico, catálogo, lançado, já, não, só, retransmissão).
- Ele diz por que a substituição não valeu: ID inexistente, ID manual, outra chave, já substituído ou retransmissão.
- O texto casa com a mensagem do ATUALIZAR, que separa incompletos, SUBSTITUI_ID inválido e erros de governança.

**Datas.**
- Mostram dd/mm/aaaa nos alvos: De/Até do Registros, Configuração, Eventos_Westgard, Sync_Log e Westgard_Fotos.
- A coluna de data/hora do Eng_Saida deixou de mostrar "#####" e mostra os segundos.

**Mensagens das listas da Digitar Resultados.** São claras: "Escolha … na seta da célula.", e a do MOTIVO lembra que
ele é obrigatório na substituição.

### A.3 Dificuldades encontradas

**Colar uma linha inteira (P-03, Médio, Hema).**
- É o atalho natural de quem digita vários manuais parecidos, mas a linha colada leva junto o ID (MAN_0007).
- Depois do ATUALIZAR, o sistema deixa a cópia fora da estatística (`MAN_0007~L17`, CONFLITO_MANUAL).
- A tela, porém, mostra as duas linhas como ATIVO, com o mesmo RUN e o mesmo DETALHE.
- O usuário confere a aba, como o GUIA §3 manda, e conclui que as duas entraram.
- O único sinal é um alerta A01 no meio de 8.914 da QA_INTEGRACAO. A mensagem do ATUALIZAR não cita o caso.
- Foi o único ponto desta rodada em que a tela contradiz o que o sistema fez.

**"X e resultado manual" / "X e retransmissão de Y" (P-01, Baixo).**
- Sem o acento do "é", a segunda frase se lê como uma lista de dois IDs, não como a razão da recusa.

**Recusa genérica no Registros H4 e no Painel L4 (P-02, Baixo).**
- A recusa funciona, mas o Excel mostra o texto padrão dele, sem a orientação "na seta da célula" que as listas da
  Digitar Resultados têm.

**"Outro" sem texto livre (D-04).**
- Com o MOTIVO só da lista (estilo Parar), "Outro: falha no LIS" é recusado.
- Quem escolhe "Outro" não tem onde dizer qual foi o motivo.

### A.4 Navegação

Nada de novo. O botão da Digitar Resultados leva à aba certa (B7 na Hema). O filtro do Registros (H4 e De/Até) só
muda a vista e volta ao estado inicial quando é limpo.

### A.5 Apresentação

- **Configuração AD26:** o Text fica vazio nos dois produtos porque a coluna AD é oculta (`hidden="1"`). Não é defeito.
- **Eventos_Westgard (Hema):** mostra só a data, sem a hora, por decisão do ADR-094.
- **Formato interno 14:** Configuração D/AD, Registros E4/F4, Estatística E8:F9/C10 (Bio) e Eventos_Westgard A
  (Hema) passaram a usá-lo. Ele segue a configuração regional do Windows (D-01).

### A.6 Tempos observados

As frentes não mediram tempo neste reteste.

### A.7 Sugestões do usuário

Para o Orquestrador priorizar. Nenhuma foi implementada.

1. **P-01:** pares com contexto no `ACENTOS_DETALHE` (`' e resultado manual'` → `' é resultado manual'`,
   `' e retransmiss'` → `' é retransmiss'`, `'1a linha'` → `'1ª linha'`), ou o texto reescrito na origem nos dois
   motores (VBA e PQ). Estender a trava `qa_adr095` D01 a esses padrões.
2. **P-02:** ErrorMessage "Escolha na seta da célula." nos filtros literais do `instalar_adr087.FILTROS` (Registros H4
   e F7:F9) e no L4 do Painel (`instalar_integracao`).
3. **P-03:** carimbar um ID novo quando a colagem trouxer um ID que já existe, ou casar o STATUS/RUN/DETALHE da tela
   pela chave da linha (`~L<n>`), e citar o caso na mensagem do ATUALIZAR.
4. **D-04:** dar um lugar para o texto livre de "Outro": uma coluna de observação, ou o estilo Aviso só para esse item.

---

## Relatório B: validação matemática

**Escopo deste ciclo.** Não houve F3 nem F2: bias, ET e Sigma **não** foram recalculados de forma independente.

O que se validou matematicamente:
1. **A estatística da série que recebe o manual lançado pelo item do menu** (critério de aceite da OC-0068):
   - n e média, recalculados por conta própria nos dois produtos;
   - DP e CV, recalculados em Python puro só na Hema.
2. **A série vizinha:** que ela não muda.
3. **A invariância entre as entregas 41 e 42:** status e colunas da tblCQ_Final dos manuais, e valores `<v>` brutos do
   XML nos alvos de formato. A tolerância é zero, porque é o mesmo motor sobre a mesma tblCQ_Final.

### B.1 Tabela

| Analito | Equipamento | Lote | Nível | Período | Dados de entrada | Fórmula encontrada | Resultado independente | Resultado QC_INI (Value2) | Diferença | Status |
|---|---|---|---|---|---|---|---|---|---|---|
| Glicose (Bio) | DIMENSION 2 | 8977 | N1 | Out/2026 | Série antes: n 23, média 76,217391304 (soma 1753,000). Mais o MAN_0001 = 123,456 (07/10/2026 01:00, ATIVO, RUN 26100701) | média = Σx/n | (1753 + 123,456)/24 = **78,1856667** | engPainel: n **24**, média **78,185666667** | 0 até a 7ª casa (a conta parte da média publicada, com 9 casas) | APROVADO (n e média) |
| Glicose (Bio) | DIMENSION 2 | 8977 | N1 | Out/2026 | Idem | DP amostral; CV = DP/média × 100 | **não recalculado** pela frente | DP 9,682841402; CV 12,38442008 % | — | NÃO EXECUTADO (DP e CV sem conferência independente) |
| Glicose (Bio) | DIMENSION 1 | 8977 | N1/N2 | Out/2026 | engPainel antes e depois do MAN_0001 e da troca do L4 | invariância | N1 n 25, média 77,36; N2 n 22 | idênticos antes e depois | 0 | APROVADO (o manual não vazou para a outra série) |
| WBC (Hema) | XN1000 | 625311 | N1 | Out/2026, antes do manual | Pontos do Eng_Saida da série (n 21) | média; DP amostral (`statistics.stdev`); CV = DP/média × 100 | n 21; média 3,1204761904761904; DP 0,054173442799603455; CV 1,736063327928693 % | n 21; média 3,12047619 | 0 na 9ª casa | APROVADO (ressalva B.5.1) |
| WBC (Hema) | XN1000 | 625311 | N1 | Out/2026, depois do MAN_0001 = 123,456 (09/10/2026 01:00, ATIVO, RUN 26100901) | Pontos do Eng_Saida (n 22, com o 123,456) | idem | n 22; média 8,590272727272728; DP 25,655674350877398; CV 298,65960215005487 %. Conta da média: (21 × 3,12047619 + 123,456)/22 = 8,5902727 | n 22; média 8,590272727; DP 25,655674351; CV 298,65960215 % | 0 na 9ª casa | APROVADO (ressalva B.5.1) |
| WBC (Hema) | XN1000 | 625311 | N2/N3 | Out/2026 | engPainel antes e depois do manual N1 | invariância | N2 n 39; N3 n 21 | idênticos | 0 | APROVADO |

### B.2 Invariância entrega 41 × 42 (tolerância zero)

**Manuais (as 17 linhas de cada produto).** Comparei STATUS, MOTIVO_EXCLUSAO_AUTOMATICA e RUN da tela, mais estas
colunas da tblCQ_Final: STATUS_ANALITICO, PARTICIPA_ESTATISTICA, ANALITO, MATRIZ, EQUIPAMENTO e RUN. Na Hema, também
LOTE e NIVEL.
- As colunas são **idênticas** nos dois produtos.
- A exceção é o EQUIPAMENTO do MAN_0001 da Bio. Na base, ele grava o texto juntado "DIMENSION 1,DIMENSION 2", que não
  entra em nenhuma série. Na nova, grava "DIMENSION 2". É o efeito esperado da OC-0068.
- O DETALHE só difere por acento.
- Nenhum RUN ou DETALHE "0".

**Valores `<v>` brutos do XML nos alvos da OC-0072:**

| Alvo | Bio (iguais / comparadas) | Hema (iguais / comparadas) |
|---|---|---|
| Eventos_Westgard A | não é alvo na Bio | 5.000 / 5.000 |
| Configuração D/AD 26-125 | 200 / 200 | 200 / 200 |
| Eng_Saida AC3:AC182 | 180 / 180 | 180 / 180 |
| Registros E4/F4 | 2 / 2 | 2 / 2 |
| Estatística E8:F9, C10 | 5 / 5 | 5 / 5 |
| Westgard_Fotos F | 13.701 / 13.701. As 33 a mais da base são células vazias, só com estilo | 4.356 / 4.356 |
| LotesStore U1 | 1 / 1 | 1 / 1 |
| Sync_Log | 16.076 células comuns, todas iguais. A nova tem a linha 22 a mais, gravada pelo ATUALIZAR da entrega | 16.012 células comuns, todas iguais. As 4 diferentes são a linha 6, gravada pelo ATUALIZAR da entrega |

**Filtros do Registros.** O hash do Value2 da tblRegistros fica constante em todos os passos do filtro: Bio
`fb1e82c46115fbb2`, Hema `1f2166a5755048ff`. A contagem de linhas visíveis bate com os dados:
- Bio: H4 = Sim mostra 1 de 3, H4 = Não 1 de 3, vazio 3 de 3.
- Hema: Sim 2 de 3, Não 2 de 3, vazio 3 de 3, e De/Até de outubro de 2026 mostra 1 de 3 (os registros são de 2025).

### B.3 Etapas para reproduzir

1. Copiar a produção atual e o backup da entrega 41 para `%TEMP%\qcw\auditoria\<ciclo>\` e conferir o sha256: Bio
   b20fe00f7e51 e 8ec815ab653d; Hema 7bab8fc008f6 e aab0a54f464e.
2. **Bio:** rodar `_arquitetura/agentes/ferramentas/reteste05p2_usuario_f1.py base41` e depois `... nova`.
3. **Hema:** rodar `reteste05p2_usuario_f1_hema.py base41`, depois `nova`, depois `xml`.
4. O que cada script faz, numa cópia por vez, aberta pelo `qa_final.QA`:
   - faz o login;
   - escreve os filtros do Painel com os eventos ligados;
   - lança os 17 manuais célula a célula e roda o ATUALIZAR;
   - lê o DETALHE (`Text`), o RUN, o STATUS, a tblCQ_Final e o engPainel antes e depois do MAN_0001;
   - grava tudo em `<produto>_F1_<cópia>.json`.
5. **Média:** (n_antes × média_antes + 123,456)/(n_antes + 1).
6. **Na Hema, DP e CV:** `statistics.mean` e `statistics.stdev` sobre os pontos da série no Eng_Saida, com
   CV = DP/média × 100.
7. **XML:** `xml_formatos_listas.py` sobre as cópias sem Excel e sobre o `SaveCopyAs` feito depois do uso.
   - Procura `x12ac:list` (OC-0068).
   - Procura numFmt com letra escapada (`dd/mm/\y\y\y\y`), para a OC-0072.
   - Compara os `<v>` base × nova célula a célula.

### B.4 Caminho do dado (conferido neste ciclo)

- **Entrada manual até o Painel.** Digitar Resultados (item da lista) → `Worksheet_Change` (carimbo do ID) → ATUALIZAR
  (`mEtlCQ`: status, MOTIVO, RUN) → tblCQ_Final (PARTICIPA = SIM) → série por equipamento × analito × lote × nível ×
  período → Eng_Saida (pontos) → engPainel (n, média, DP, CV).
- **Substituição (ADR-082).** SUBSTITUI_ID + MOTIVO → o S1 vira SUBSTITUIDO_MANUAL e PARTICIPA NÃO, com o motivo
  escolhido na trilha.
  - Quando a substituição é inválida, o manual continua ATIVO (ID inexistente, ID manual) ou fica CONFLITO_MANUAL
    (outra chave, já substituído, retransmissão), e nada é substituído.
- **CEQ → bias → ET → Sigma e CIQ → CV → ET/Sigma:** não exercitados neste ciclo.

### B.5 Diferenças de método e observações (não reprovadas)

1. **Precisão da comparação na Hema (ressalva).**
   - A frente comparou com o engPainel "no 9º decimal", porque o leitor do script arredonda em 9 casas. Não comparou
     com a precisão total do `Value2`.
   - Para os valores grandes (CV 298,66 %), 9 casas são mais apertadas que o 1e-9 relativo das REGRAS §3.5.
   - Para o DP de 0,054 (antes do manual), 1e-9 absoluto equivale a cerca de 2e-8 relativo, mais folgado que a regra.
   - Não há indício de divergência, mas a conferência estrita a 1e-9 relativo não foi feita.
2. **Independência dos pontos (ressalva).** Na Hema, os pontos vieram do Eng_Saida, que é saída do QC_INI, e não da
   tblCQ_Final bruta.
   - A conta confere a estatística a partir dos pontos publicados.
   - A seleção dos pontos foi conferida só indiretamente: o n bate (21 → 22), o 123,456 está entre os pontos, e N2 e
     N3 não mudaram.
3. **Bio: só n e média.** A média foi conferida a partir da média publicada (76,217391304 × 23 = 1753,000). O DP e o
   CV da Bio não foram recalculados.

---

## Relatório C: ocorrências

**Deduplicação.** As frentes trouxeram 5 ocorrências; ficaram 3:
- R5P2-01 (Bio) e R5P2-H01 (Hema) são o mesmo texto de origem comum (`src/comum/mEtlCQ.bas`). Viraram **P-01**.
- R5P2-02 (Bio) e R5P2-H02 (Hema) são a mesma falta de ErrorMessage nos filtros literais. Viraram **P-02**.
- R5P2-H03 (Hema) virou **P-03**.

Os IDs P-xx são provisórios; o ID definitivo (OC-xxxx) é dado pelo Orquestrador. Nenhuma ocorrência é Crítica ou Alta.
Nenhuma é efeito colateral da entrega 42: o P-02 e o P-03 já existiam na entrega 41, e o P-01 é o resíduo da OC-0021.

### C.1 Tabela

| ID | Produto | Prioridade | Tipo | Tela/módulo | Resumo | Origem | ID anterior | Outro produto |
|---|---|---|---|---|---|---|---|---|
| P-01 | Ambos | Baixo | `usabilidade` | Digitar Resultados, coluna DETALHE (`mEtlCQ.RazaoSubstituicao`, `mEtlCQ.bas:808/816`; `camada_dados.ACENTOS_DETALHE`) | O verbo "é" sem acento em "MAN_0002 **e** resultado manual" e em "HEM-212660 **e** retransmissão de HEM-212425". Na tblCQ_Final e na QA_INTEGRACAO, "vale a 1a linha" (`mEtlCQ.bas:1244`) | Bio F1 R5P2-01 e Hema F1 R5P2-H01 | OC-0021 (resíduo; a OC-0021 fica PERSISTE) | Bio: o caso da retransmissão não foi provocado (não há DUPLICIDADE_ORIGEM); o código é comum |
| P-02 | Ambos | Baixo | `usabilidade` | Registros H4 e F7:F9 (`instalar_adr087.FILTROS`); Painel L4 da Bio (`instalar_integracao`) | A recusa de valor inválido funciona, mas o `Validation.ErrorMessage` está vazio: aparece a mensagem genérica do Excel, sem "na seta da célula" | Bio F1 R5P2-02 e Hema F1 R5P2-H02 | ligado à OC-0068 (que fica RESOLVIDA); já era assim na entrega 41 | O L4 só existe na Bio |
| P-03 | Hematologia (código comum) | **Médio** | `usabilidade` | Digitar Resultados: STATUS/RUN/DETALHE por XLOOKUP do ID; `Worksheet_Change`; motor `mEtlCQ.bas:1244` | A linha colada com o ID junto aparece ATIVA na tela, mas o motor a deixa fora da estatística (`MAN_0007~L17`, CONFLITO_MANUAL) | Hema F1 R5P2-H03 | — (já era assim na entrega 41) | Bio não executado; o motor é comum (`src/comum/mEtlCQ.bas`) |

### C.2 Detalhe

**P-01 (Ambos, Baixo, usabilidade): DETALHE da substituição com o verbo "é" sem acento**
- **Tela/módulo:** aba Digitar Resultados, coluna DETALHE.
  - Origem: `mEtlCQ.RazaoSubstituicao` (`src/comum/mEtlCQ.bas`, linhas 808 e 816) e o mesmo texto no PQ.
  - A tela acentua pelo `camada_dados.ACENTOS_DETALHE`, que troca só palavras inteiras. Trocar " e " quebraria a
    conjunção ("da estatística e do gráfico").
- **Ação:** digitar no SUBSTITUI_ID o ID de outro manual (MAN_0002), ou de um resultado retransmitido (HEM-212660, só
  na Hema). Completar a linha e clicar em ATUALIZAR DADOS.
- **Esperado:**
  - "SUBSTITUI_ID inválido: MAN_0002 **é** resultado manual (…)";
  - "… HEM-212660 **é** retransmissão de HEM-212425 (informe o ID original) …".
- **Observado:**
  - "SUBSTITUI_ID inválido: MAN_0002 e resultado manual (SUBSTITUI_ID aceita só resultado do interfaceamento) (este
    manual não substitui nada)". Visto no MAN_0010 da Bio e no MAN_0008 da Hema.
  - "coincide com resultado do interfaceamento (0 min); SUBSTITUI_ID inválido: HEM-212660 e retransmissão de
    HEM-212425 (informe o ID original) (este manual não substitui nada)". Visto no MAN_0011 da Hema.
  - O "ID_REGISTRO manual repetido (vale a 1a linha)" aparece na tblCQ_Final e na QA_INTEGRACAO. Não aparece na tela
    (ver P-03).
  - A trava `qa_adr095` D01 não pega esses casos, porque o "e" não está em `SEM_ACENTO_PROIBIDO`.
- **Dificuldade:** pouca. A segunda frase parece uma lista de dois IDs.
- **Consequência:** só de leitura. STATUS, MOTIVO, RUN e números são iguais aos da base.
- **Sugestão:**
  - pares com contexto no `ACENTOS_DETALHE` (`' e resultado manual'`, `' e retransmiss'`, `'1a linha'` → `'1ª linha'`);
  - ou o texto reescrito na origem, nos dois motores;
  - e a trava D01 estendida a esses padrões.
- **Evidência:**
  - `Bioquimica_F1_nova.json`: `O21_O68_manuais.manuais.M10_subst_id_manual`.
  - `Hematologia_F1_nova.json`: `O21_O68_manuais.manuais.M10_subst_id_manual` e `.M13_subst_retransmissao`.
  - Produções b20fe00f7e51 e 7bab8fc008f6.

**P-02 (Ambos, Baixo, usabilidade): filtros do Registros e L4 do Painel recusam com a mensagem genérica do Excel**
- **Tela:** Registros H4 ("Calibrado?") e F7:F9 ("Foi calibrado?"), nos dois produtos; Painel L4, só na Bio.
- **Ação:** digitar um valor fora da lista: "Talvez" ou "S" no H4, "DIMENSION 9" no L4.
- **Esperado:** "Escolha … na seta da célula.", como nas listas da Digitar Resultados (ADR-095 §1).
- **Observado:** `Validation.ErrorMessage` vazio, na base e na nova. A recusa funciona: `Validation.Value` = Falso.
- **Consequência:** nenhuma nos dados. A orientação é pior.
- **Sugestão:** ErrorMessage "Escolha na seta da célula." no `instalar_adr087.FILTROS` e no L4 do
  `instalar_integracao`.
- **Evidência:**
  - Bio: `O68_listas.Registros_H4.msg` e `Painel_L4.msg`.
  - Hema: `O68_listas.Registros_H4.msg` e `Registros_F7_foi_calibrado.msg`.
- **Não executado:** ver a caixa na tela (Excel com `Interactive = False`).

**P-03 (Hema, Médio, usabilidade): linha manual colada com o ID junto aparece ATIVA, mas fica fora da estatística**
- **Tela/módulo:** aba Digitar Resultados.
  - STATUS, RUN e DETALHE vêm de XLOOKUP pelo ID_REGISTRO, que acha sempre a 1ª linha.
  - O `Worksheet_Change` não carimba ID novo na colagem.
  - O motor marca a repetida em `mEtlCQ.bas:1244`.
- **Ação:** copiar a linha inteira do MAN_0007, colar na próxima linha livre e clicar em ATUALIZAR DADOS.
- **Esperado:** uma de duas coisas:
  - um ID novo carimbado na colagem;
  - ou a tela mostrar a linha colada como CONFLITO_MANUAL ("ID_REGISTRO manual repetido").
- **Observado:**
  - A linha colada manteve o ID MAN_0007, sem aviso.
  - Na tblCQ_Final: `MAN_0007` (ATIVO, SIM, RUN 26092901) e `MAN_0007~L17` (CONFLITO_MANUAL, NÃO, sem RUN, "ID_REGISTRO
    manual repetido (vale a 1a linha)").
  - Na QA_INTEGRACAO: o A01 "… confirme e inative um deles ou corrija o manual".
  - Na tela: as duas linhas com STATUS ATIVO, RUN 26092901 e o mesmo DETALHE.
  - A mensagem do ATUALIZAR não cita o caso.
  - Base e nova se comportam igual.
- **Dificuldade:** o usuário confere a aba, vê ATIVO nas duas linhas e conclui que as duas entraram. "Inative um
  deles" é ambíguo, porque os dois têm o mesmo ID na tela.
- **Consequência:**
  - Os números estão certos: a cópia não entra no Painel.
  - Uma correção feita na linha colada não vale, e a tela continua dizendo ATIVO.
  - Pela REGRAS §4, é inconsistência de visualização: Médio.
- **Sugestão:**
  - carimbar ID novo quando a colagem trouxer um ID existente;
  - ou casar a tela pela chave da linha (`~L<n>`);
  - e citar o caso na mensagem do ATUALIZAR. Decisão do Orquestrador.
- **Evidência:** `Hematologia_F1_nova.json`:
  - `O21_O68_manuais.M19_final`;
  - `.manuais.M19_linha_colada_de_M09`;
  - `.pre.M19_linha_colada_de_M09`;
  - `.qa_integracao_manuais` (A01 `MAN_0007~L17`).
  - Mesmo resultado em `Hematologia_F1_base41.json`.
- **Não executado:** o mesmo fluxo na Bio.

### C.3 Dúvidas de regra de negócio (para o Orquestrador encaminhar)

| ID | Produto | Para | Dúvida |
|---|---|---|---|
| D-01 | Ambos | Orquestrador / UX | OC-0072: Configuração D/AD, Registros E4/F4, Estatística E8:F9/C10 (Bio) e Eventos_Westgard A (Hema) passaram ao formato interno 14, a data abreviada do Windows. Ele mostra dd/mm/aaaa em pt-BR, mas segue a configuração regional da máquina. O Eng_Saida, o Sync_Log e o Westgard_Fotos usam `dd/mm/yyyy` explícito. Todas as máquinas do laboratório estão em pt-BR, ou é melhor usar o explícito também nessas células? |
| D-02 | Ambos | Especialista em CQ | IMPLAUSIVEL e LOTE_* (fora da estatística) mostram o RUN do dia (Bio 26092501, Hema 26092901), igual à base. É intencional, para mostrar a corrida em que o resultado foi lançado? Continua em aberto a dúvida do reteste 04: um IMPLAUSIVEL abre corrida e renumera as corridas reais do dia? A renumeração não foi medida |
| D-03 | Ambos | Orquestrador | Risco residual do ADR-095 §1 ("não feito"): colar texto por cima da validação grava um equipamento fora do cadastro, e o manual entra ATIVO, fora de qualquer série. A regra ou o QA que impede isso depende de autorização |
| D-04 | Ambos | Orquestrador / Especialista em CQ | Com o MOTIVO só da lista (estilo Parar), "Outro: <texto livre>" é recusado. O ADR-082 trata o motivo como justificativa de auditoria da substituição. Falta um lugar para o texto livre de "Outro" (coluna de observação, ou estilo Aviso), ou bastam os 4 itens? |

**Observação sem ocorrência.** Na Hema, o Eventos_Westgard caiu de 5.000 para 1.763 linhas depois do ATUALIZAR, igual
na base e na nova. A explicação provável são as 2 regras TESTE, que tiraram 8.311 resultados reais da estatística da
cópia. Não foi investigado, porque não afeta o formato. A Orquestração pode pedir ao Tester/QA a confirmação.

---

## Relatório D: parecer final

Critério: "validado" só onde houve teste executado e aprovado neste ciclo. O resto é "não validado" (teste reprovado
ou que persiste) ou "não executado" (com o motivo). Este ciclo foi um reteste de 3 itens, só com a F1.

### D.1 A rotina funciona?

**Sim, no que foi refeito, com um resíduo de texto e uma inconsistência de tela já existente.**

**Validado:**
- OC-0068 nos dois produtos:
  - listas com um item por valor;
  - válidos aceitos e inválidos recusados;
  - o manual escolhido no menu entra na série certa (Bio DIMENSION 2, Hema WBC 625311);
  - o motivo vai certo para a trilha da substituição;
  - 0 `x12ac:list`.
- OC-0072 nos dois produtos: 0 "yyyy" na tela e 0 numFmt literais no XML, antes e depois do uso, com os valores iguais
  aos da base.
- Os fluxos vizinhos da OC-0021, OC-0068 e OC-0072 (6 testes), sem efeito colateral.

**Não validado:**
- OC-0021 **PERSISTE** com resíduo (P-01, Baixo, só texto).
- A tela da linha colada contradiz o motor (P-03, Médio, já existia).

**Não executado:**
- as caixas de aviso do SUBSTITUI_ID e o clique físico na seta, porque o Excel de automação roda com
  `Interactive = False`;
- CEQ, LJ e Westgard, que ficaram fora da lista.

### D.2 Bias, ET e Sigma estão validados?

**Não executado neste ciclo** (não houve F3 nem F2). Nenhum número de bias, ET ou Sigma foi lido ou recalculado.

A última validação independente é a do reteste 04 (47 cálculos, 0 divergências), feita sobre a entrega 41. A
invariância bit a bit da Estatística da Bio entre as entregas 41 e 42 foi medida no reteste 05 (parte 1).

### D.3 A integração CIQ + CEQ está validada?

**Não executado neste ciclo.** Nenhum CEQ foi incluído nem corrigido, e nenhuma leitura de bias foi feita.

### D.4 Os filtros de período e de lote estão corretos?

**Validado, no escopo executado:**
- os filtros do Registros (H4 Sim/Não/vazio e De/Até) só mudam a vista, com o hash da tblRegistros constante e as
  contagens visíveis coerentes;
- a troca do L4 no Painel da Bio altera só o engPainel, e a série DIMENSION 1 fica idêntica;
- a série Out/2026 do lote 8977 (Bio) e do 625311 (Hema) recebe o manual com n e média conferidos.

**Não executado:** mês a mês, trimestres, troca de lote, H3 vazio, LJ, Westgard e corrida do lote. A F4 não fez parte
deste reteste.

### D.5 A entrada manual e a automática estão validadas?

**Entrada manual: validada nos 34 casos executados** (17 por produto). Status, MOTIVO, RUN e tblCQ_Final são iguais
aos da base, menos a correção esperada do equipamento. A substituição válida e as inválidas funcionam. O manual do
menu entra na série certa, com n e média conferidos nos dois produtos, e DP e CV conferidos na Hema.

**Ressalvas:**
- O DETALHE tem o resíduo do P-01.
- A linha colada aparece ATIVA na tela sem estar (P-03).
- As caixas de aviso não foram vistas.

**Entrada automática (DB_SEAC → tblCQ_Final): não executada.** O ATUALIZAR rodou nos dois produtos, mas não havia
resultado novo no DB_SEAC (últimos de 13/11/2026), então não houve recebimento a conferir.

**Equivalência manual × automático:** não executada.

### D.6 Há divergência matemática?

**Não**, nos cálculos executados.
- **Relatório B:** 6 linhas, 4 aprovadas sem ressalva, 0 reprovadas. As 2 com ressalva:
  - DP e CV da Bio sem conferência independente;
  - na Hema, comparação no 9º decimal e não na precisão total do `Value2`, com os pontos lidos do Eng_Saida.
- **Invariância base × nova:** 0 diferenças não esperadas, com tolerância zero, nos manuais e nos valores do XML.
- Bias, ET e Sigma não foram recalculados neste ciclo (D.2).

### D.7 Há problema operacional?

**Nada que impeça a rotina.** O ponto operacional é o P-03 (Médio): colar uma linha inteira gera uma linha que parece
ATIVA e não entra na estatística. Os números ficam certos, mas o usuário pode achar que lançou o que não lançou, ou
corrigir a linha errada.

Também há o P-01 e o P-02 (Baixo), de texto. As produções ficaram intactas: mesmo hash antes e depois de todas as
execuções.

### D.8 O que fica pendente

1. **P-01** (resíduo da OC-0021, que fica PERSISTE): corrigir o texto e estender a trava D01. No reteste, provocar a
   retransmissão também na Bio, se houver caso.
2. **P-02** (Baixo) e **P-03** (Médio): decidir e corrigir. No reteste do P-03, refazer a colagem também na Bio.
3. **Dúvidas:** encaminhar D-01 (formato 14 × pt-BR), D-02 (RUN no IMPLAUSIVEL e renumeração), D-03 (texto colado por
   cima da validação) e D-04 ("Outro" sem texto livre).
4. **Comparação estrita:** refazer a conferência do DP e do CV com a precisão total do `Value2` (1e-9 relativo), com os
   pontos tirados da tblCQ_Final bruta. Incluir o DP e o CV da Bio.
5. **Seguem abertas, de ciclos anteriores:**
   - do reteste 05 (parte 1): o P-01 do J5 "00/01/1900" e as dúvidas D-01 a D-04;
   - a OC-0039;
   - do ADR-095 §6: a OC-0069 e a OC-0070.
6. **Não executados, a cobrir num próximo ciclo:**
   - F3 sobre a entrega 42 (bias, ET e Sigma, ETp ausente, ETp VB/FAB);
   - F2 (CEQ);
   - recebimento automático com resultado novo do DB_SEAC e equivalência manual × automático;
   - F4 completa (mês a mês, trimestres, lotes, LJ, Westgard, corrida do lote);
   - caixas de aviso do SUBSTITUI_ID e clique físico, que pedem uma sessão interativa.

Dados pessoais: nenhum login, máquina ou caminho de rede consta deste parecer. As frentes omitiram o login do Windows
gravado no Audit_Log e no Sync_Log das cópias.
