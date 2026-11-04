# Reteste do ciclo 01 — Usuário Simulado (ADR-088)

Data: 04/11/2026 · Agente: AGENTE_USUARIO_SIMULADO · Independente do Tester (não usei `valida_adr088.py` nem `repro_*`).

## Ambiente e método

| Item | Valor |
|---|---|
| Produção copiada (nunca aberta) | `QC_Bioquimica.xlsm` **06bdd4ec53b7**, `QC_Hematologia.xlsm` **4c0b41f0c0a6**: iguais antes e depois de cada etapa |
| Cópias | `%TEMP%\qcw\auditoria\2026-11-04_reteste01\usuario\<produto>\` (uma cópia nova por sessão; nada salvo, exceto a `<produto>_prep.xlsm`, salva só em %TEMP%) |
| DB_SEAC | congelado por cópia de arquivo (Bio da2afa4a161f), apontado só nas cópias |
| 1º ATUALIZAR depois da entrega | feito em todas as cópias (`AtualizarDadosAutomatico`, o núcleo do botão). Bio: 448 novos, 74 s; Hema: 2.542 novos, 72 s |
| Excel | um por vez, sempre sequencial; só os PIDs dos meus scripts foram encerrados (pelo `fechar`); nenhuma captura de tela |
| Runner | `_arquitetura/agentes/ferramentas/reteste01_usuario.py` (novo; documentado no cabeçalho). Ele chama os scripts do ciclo 01: `f1_entrada_ciq.py` (sessões A, D/E, G), `f4_periodos_lotes.py` / `f4_hema.py` e `calc_independente.py` |
| Alvo das sessões D/E/G | o **mesmo do ciclo 01** (Bio Glicose/8977/DIMENSION 1/30-10; Hema WBC/625311/XN1000/30-10), tirado do `f1_A.json` de 31/10. A sessão A de hoje escolhe o mês corrente (novembro, Glicose 5659/DIMENSION 2, o lote de urina da OC-F4-2), que não tem dois dias anteriores com corrida |
| Calculadora | `calc_independente.py` atualizada **antes** de comparar, com a regra nova conferida no código (`mCEQ.AmostraDoVies`, chamada no `BiasEQ` depois da escolha do ano vigente): fora NAO AVALIADO, resultado ou alvo = 0, cada CHAVE uma vez. `CALC_BIAS_PRE_ADR088=1` refaz a regra antiga. O ET do Painel da Hema passou a ser vazio sem bias numérico (ADR-088 dec. 2) |

Ajustes feitos durante o reteste: (1) a 1ª passada da sessão E (Bio) falhou porque o alvo de novembro não tinha dois dias anteriores; por isso refiz D/E/G com o alvo do ciclo 01 (a G da Bio com o alvo de hoje ficou guardada em `f1_G_alvo_hoje.json` e dá o mesmo resultado). (2) `REGISTRAR_LJ` não existe na tblCQ_Final (o script passou a tolerar isso). (3) H3 é célula mesclada (`MergeArea.ClearContents`, como o f4).

## Resultado por ocorrência

| OC | Produto | Estado | Evidência |
|---|---|---|---|
| OC-0001 | Bio | **RESOLVIDA** | Sessão E: "dimension 1" vira **DIMENSION 1** já na digitação (linha antes do ATUALIZAR). O manual N2 244,222 de 29/10 23:36 ganha o **RUN próprio 26102903** (corrida 3), e não o 26102901. O RUN 26102901 continua com o N2 real **277** (BIO-371977), veredito OK. A Estatística conta **um** a mais: n 23 → 24, média 258,13425, igual ao independente, com o equipamento exato e em maiúsculas. QA: A06 (resultado do mesmo dia). "glicose" (28/10) → Glicose, RUN 26102807, n 25. `f1_E.json` |
| OC-0001 | Hema | **RESOLVIDA** | Sessão D: "xn1000" → **XN1000**, RUN **26102904** (corrida 4). O RUN 26102901 continua com o N3 real **17,58**. N3 n 22 → 23, média 17,384913… = independente = Estatística. "wbc" → WBC, RUN 26102803, n 24. A12 = 0 nas duas cópias preparadas. Na sessão A, V16 ("xn1000") também foi normalizado. `f1_D.json` |
| OC-0002 | Bio | **RESOLVIDA** | Painel com HbA1c / DIMENSION 2 / 8590 / Set-2026, aba ativa e filtros pelas células: N1 n 23, CV 2,7136, G "SEM EP", **H vazio** (no ciclo 01, H8 = 2,7740 = 1,65·CV). Fórmula de H7: `IF(OR($E7="",NOT(ISNUMBER($G7))),"",…)`. Estatística H = independente (62/62). `painel2.json`, `calc.json` |
| OC-0002 | Hema | **RESOLVIDA** | IG% (SEM EP pelo ADR-088): CV 3,11 / 2,85 / 2,74, G "SEM EP", **H7:H9 vazios**. Com bias numérico (RBC), H7:H9 = 1,65·CV + \|bias\| (calculadora: Painel **45/45 OK**). `painel.json`, `calc.log` |
| OC-0003 | Hema | **RESOLVIDA** | `f4_hema.py`, mesmo fluxo do ciclo 01: H3 = 619611 → B4 619611, e o WBC passa a **n 0** (o lote não tem CIQ em 2025; antes ficava 50/3,0446 do 522611). Ano De/Até 2026 → **54 / 3,12056** = independente. Divergência real: 0 (máx. 1,1e-15). Os 3 "divergentes" são as linhas vazias da tabela (analito "0", "" × 0), o mesmo artefato do script no ciclo 01. Tempo da troca: 1,46 s |
| OC-0003 | Bio (vizinho) | **RESOLVIDA** | A Estatística da cópia preparada é igual ao independente em 62 linhas × 19 colunas. f4 da Bio: 78 aprovados; o único reprovado é o P-LTudo (OC-F4-1, das 180 corridas, fora do escopo) |
| OC-0004 | Bio | **RESOLVIDA** | Automático: depois do ATUALIZAR, **56 IMPLAUSIVEL** (PARTICIPA NÃO, NAO_PLOTAR), A11 = 56 com os mesmos IDs (inclui BIO-371183 = −1, o A06 do ciclo 01), **0 participantes ≤ 0**. Manual (sessão G, alvo do ciclo 01): −1 e 0 → STATUS **IMPLAUSIVEL**, A11, n/média/DP de N1 e N2 idênticos antes e depois. Recebem RUN (26102501, 26102401), como o ADR manda |
| OC-0004 | Hema | **RESOLVIDA** | Automático: 0 na produção (participantes ≤ 0 = 0). Manual −1 e 0 → IMPLAUSIVEL, A11, estatística idêntica (N1 22 / N2 42 / N3 22) |
| OC-0005 | Bio e Hema | **RESOLVIDA** | Sessão A, V13: "31/02/2026" colada → **MANUAL_INCOMPLETO**, DETALHE "campos obrigatorios ausentes/invalidos: DATA", RUN 0, PARTICIPA NÃO, DATA vazia na tblCQ_Final (antes virava 03/03 ATIVO) |
| OC-0006 | Bio | **RESOLVIDA** | Bilirrubina direta G = **19,99420975515574** (QC_INI = independente, dif 0), AC 13, AD 3; H e L com dif ≤ 5e-14. Com a regra antiga, a mesma calculadora dá 30,77 / AC 15, ou seja, a mudança está nos números. Triglicerídeos e os outros: 62 linhas × 19 colunas OK |
| OC-0006 | Hema | **RESOLVIDA** | IG%, IG#, NRBC% e NRBC# × 3 níveis: G = "SEM EP", H vazio, L "-", AC/AD "SEM EP". Calculadora: 84 linhas, todas as colunas OK |
| OC-0009 | Bio | **RESOLVIDA** | B103 = **4**. Independente: 4 linhas sem bias (HbA1c N1/N2, PCR N1/N2). B98:B102 = 0/0/3/2/53 = contagem independente das classes (0/0/3/2/53) |
| OC-0013 | Bio e Hema | **RESOLVIDA** (no fluxo original) | `f4` L-TODOS (8974 → 8977 → H3 vazio; Hema 536411 → 602911 → vazio): H4 "▲ H3 vazio = TODOS os lotes…", média/DP/CV/ET vazios, Bias e Sigma visíveis (Bio 2,7333 / 1,2553). G6 e I6 usam o mesmo `engTodos` do H4 (`=IF(engTodos,"Bias % do método (CEQ)",…)`, `"Sigma do método"`), então os rótulos trocam junto com o H4. Ver a OC nova abaixo |

## Impressão de usuário: a mudança ficou clara na tela?

- **OC-0001:** sim, e sem esforço. A grafia é corrigida na hora em que eu digito, antes do ATUALIZAR, e o manual aparece como uma corrida a mais. Só que nada avisa que o texto foi corrigido. Um usuário que digitou "dimension 1" não fica sabendo que a lista tem a grafia certa (Baixo, acabamento).
- **OC-0002:** sim. O ET em branco ao lado de "SEM EP" fica coerente com a Estatística. Falta um texto do tipo "sem CEQ, sem ET": o vazio pode parecer falha de cálculo para quem não conhece a regra.
- **OC-0003:** sim. A Estatística acompanha o lote e o ano na hora (1,5 s). Um lote sem dado no ano mostra n 0 e células vazias, o que é honesto.
- **OC-0004:** clara no manual: a linha diz **IMPLAUSIVEL** e "resultado <= 0: implausivel para material de controle (fora da estatistica e do grafico)". Pouco clara no automático: o ATUALIZAR só diz "56 alerta(s)", e para saber quais resultados saíram é preciso abrir a QA_INTEGRACAO (A11). O Painel e o LJ não mostram que houve exclusão, e o Audit_Log não registra a passagem (pendência d do ADR).
- **OC-0005:** clara o bastante: STATUS MANUAL_INCOMPLETO e "ausentes/invalidos: DATA". Não diz que 31/02 não existe, mas aponta o campo.
- **OC-0006:** a mudança é só nos números, sem nenhum sinal na tela. O K5 ainda conta as amostras pela regra antiga (pendência c do ADR), e o laboratório pode ver no K5 um número de amostras diferente do AC.
- **OC-0009:** sim. "Sem EP (sem bias → sem Sigma)" = 4 bate com o que eu vejo nas linhas.
- **OC-0013:** sim, no fluxo que funcionou. O H4 explica o modo TODOS, e os rótulos "do método" deixam claro que Bias e Sigma não são do lote.

## Ocorrências novas

**R01-U1 — H3 vazio às vezes não entra no modo TODOS e deixa o Painel "DESATUALIZADO" (Médio; `funcional`; a confirmar pelo Tester)**
- Tela/módulo: Painel, H3 (ADR-086).
- Ação, numa cópia preparada, com a aba Painel ativa e os eventos ligados:
  - Bio: L4 DIMENSION 2 → C3 Hemoglobina glicada → H3 8590 → K3 2026 → K4 Set (e, numa 2ª tentativa, também G3/G4 digitados) → apagar H3;
  - Hema: C3 WBC → H3 602911 → K3 2026 → K4 Mar → apagar H3.
- Esperado: H4 "H3 vazio = TODOS", motor com todos os lotes e G6/I6 "do método".
- Observado (4 execuções, com pausa de 3 s): H4 continua "lote em análise". J7:J9 mostra **DESATUALIZADO**. B:I continuam com os números do lote anterior. A39 passa a "Lote 8974 / 522611 (lote em uso) · Nenhuma corrida deste lote no período". Um usuário leria números de um lote junto com o alvo de outro.
- O fluxo do f4 (lote A → lote B → apagar H3) entra no TODOS nos dois produtos, e não isolei a diferença. Não comparei com uma cópia anterior ao ADR-088 (o `TrocarLoteAnalise` ganhou o `AtualizarEstatisticaAba` nesta entrega), por isso não sei se é EFEITO_COLATERAL.
- Evidência: `usuario\<produto>\painel.json` e `painel2.json`; script `ferramentas/reteste01_painel_h3.py <Bioquimica|Hematologia>` (G34=1 digita também G3/G4); ele usa a `<produto>_prep.xlsm`.

Observações sem OC nova: o ADR-088 previa o resumo Sigma da Bio 6/1/3/7/41, mas a tela e o independente dão 0/0/3/2/53, porque o texto do ADR está errado e o número está certo (o Tester já anotou isso). A sessão A escolhe como "alvo de Glicose" do mês o lote 5659 (urina), o que reforça a OC-F4-2 já aberta.

## Arquivos

- Saídas: `%TEMP%\qcw\auditoria\2026-11-04_reteste01\usuario\{Bioquimica,Hematologia}\` (`f1_A/D/E/G.json`, `prep.json`, `painel.json`, `painel2.json`, `f4.log`/`f4h.log`, `calc.json`/`calc.log`, `runner_*.log`).
- Ferramentas: `reteste01_usuario.py` (novo); `calc_independente.py` (bias ADR-088 dec. 6 e ET do Painel da Hema dec. 2; `CALC_BIAS_PRE_ADR088=1` para a regra antiga).

