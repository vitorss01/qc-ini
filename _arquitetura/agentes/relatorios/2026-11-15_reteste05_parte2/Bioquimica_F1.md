# Reteste 05, parte 2: Usuário Simulado, Bioquímica, frente F1

Data: 15/11/2026 · Agente: AGENTE_USUARIO_SIMULADO · Modo: RETESTE (OC-0021, OC-0068, OC-0072) · Produção: entrega 42.

## Ambiente e método

| Item | Valor |
|---|---|
| Produção atual (só copiada) | `QC_Bioquimica.xlsm` sha256 curto **b20fe00f7e51**, igual antes e depois (`origem_intacta.igual = true`) |
| Base de comparação ("antes") | `_backup_pre_integracao_2026-11-15_0011/QC_Bioquimica.xlsm` sha256 curto **8ec815ab653d** (entrega 41), igual antes e depois |
| Cópias (novas, nenhuma salva, apagadas no fim) | `%TEMP%\qcw\auditoria\2026-11-15_reteste05_parte2\Bioquimica_F1_base41.xlsm` e `Bioquimica_F1_nova.xlsm`; para o XML sem Excel, `Bioquimica_F1_xml_{base41,nova}.xlsm`; e `Bioquimica_F1_nova_SaveCopyAs.xlsm`, gravado da cópia em uso só para ler o XML depois do uso |
| Excel | um por vez, abertos e encerrados pelo `qa_final.QA`: PID 30212 (1ª tentativa da base, que parou num erro do meu script antes de qualquer leitura; cópia descartada), PID 24532 (base) e PID 12032 (nova). Os Excel 13084, 21368 e 30164 (tarefa agendada) não foram tocados. Nenhuma captura de tela |
| Scripts (novos, documentados no cabeçalho) | `_arquitetura/agentes/ferramentas/reteste05p2_usuario_f1.py <base41\|nova>` e `_arquitetura/agentes/ferramentas/xml_formatos_listas.py` (leitura direta do XML, sem Excel nem openpyxl) |
| Saídas (fora do repositório) | `Bioquimica_F1_base41.json/.log`, `Bioquimica_F1_nova.json/.log`, `xml_resumo.json`, `xml_alvos_comparacao.json`, na pasta do ciclo no `%TEMP%` |
| Usuário | `TESTE_AUDITORIA_R5P2`, papel ANALISTA, criado só na cópia; login pelo `mSeguranca.DoLogin` |
| Dados de teste | 17 manuais MAN_0001..0017 com MOTIVO "Outro: TESTE-AUDITORIA" (um, de propósito, sem motivo), valores 123,456 / 111,111 / 12,345 / 1,234 / −1; SUBSTITUI_ID 999999999 (inexistente) |
| tblCQ_Final | 35.906 linhas nas duas cópias; os dois ATUALIZAR leram o DB_SEAC só para leitura e não trouxeram resultado novo (último recebido 13/11/2026 08:39:01) |

**Caminho usado.** Fiz o caminho do usuário:
1. Login.
2. Painel: L4, C3, H3, K3 e K4 escritos com os eventos ligados.
3. `mIntegracao.IrDigitarResultados`, que é a navegação do botão.
4. Manuais digitados célula a célula (`FormulaLocal`, eventos ligados). Na substituição, digitei primeiro o SUBSTITUI_ID, como manda o GUIA §3.
5. **ATUALIZAR DADOS** pelo núcleo do botão (`mIntegracao.AtualizarDadosAutomatico`).
6. Na aba Registros: filtro H4 e De/Até em E4/F4, digitados.

A mesma sequência rodou na base (entrega 41) e na nova (entrega 42). Isso permite separar o que a entrega mudou do que
já era assim.

**Não executado (sem interação com a tela):**
- O clique físico na seta da validação e a caixa "Parar". Usei o veredito do próprio Excel (`Range.Validation.Value`),
  que é o critério que a caixa aplica.
- As caixas de aviso do SUBSTITUI_ID (ID inexistente, ID manual, chave diferente). O `xlh` abre o Excel com
  `Interactive = False`, e o evento só mostra o aviso quando `Application.Interactive` é verdadeiro. Nenhuma caixa
  apareceu nas duas cópias.

## Resultado por ocorrência

| OC | Estado | Resumo |
|---|---|---|
| OC-0021 | **PERSISTE** (residual pequeno) | Os 9 textos que estavam sem acento no reteste 04 agora aparecem acentuados, e status, MOTIVO e RUN ficaram idênticos aos da base. Sobra um texto da substituição (ADR-082) com o verbo "é" sem acento ("MAN_0002 **e** resultado manual"). Por leitura do código, há mais dois iguais ("… **e** retransmissão de …", "vale a **1a** linha") |
| OC-0068 | **RESOLVIDA** | Uma opção por valor nas 5 listas. Válidos aceitos, inválidos e o texto juntado recusados. O manual com "DIMENSION 2" escolhido no menu entra na série DIMENSION 2 (n 23 → 24, média conferida). Nenhum `x12ac:list` no XML, nem antes nem depois do uso. O filtro do Registros só muda a vista |
| OC-0072 | **RESOLVIDA** | Nenhuma célula-alvo mostra "yyyy". As datas aparecem em dd/mm/aaaa, com :ss onde há hora. O XML da cópia, antes e depois do uso, não tem nenhum formato com letra literal. Os valores das células são iguais aos da base |

### OC-0021: DETALHE sem acento

**Fluxo original.** Lancei manuais que geram cada texto possível de MOTIVO_EXCLUSAO_AUTOMATICA num manual. Depois
rodei ATUALIZAR e li o DETALHE (`Text`), o RUN (`Text`), o STATUS e a linha da tblCQ_Final.

Referências reais usadas:
- R = BIO-356720 (Glicose N1, DIMENSION 1, lote 8977, 25/09/2026 11:35:49);
- S1 = BIO-356623, S2 = BIO-355801 e S3 = o 43º mais recente da mesma série;
- R2 = BIO-362158 (DIMENSION 2, 07/10/2026);
- regras de lote da cópia: FORA_DO_CQ 8177, SO_ANALITOS 8353 e MATRIZ URINA 6354 (17 regras).

| Manual | Cenário | STATUS (igual na base) | DETALHE na entrega 42 (Text) | Base (entrega 41) |
|---|---|---|---|---|
| MAN_0002 | −1, lote comum (ADR-088) | IMPLAUSIVEL | "resultado <= 0: implausível para material de controle (fora da estatística e do gráfico)" | implausivel / estatistica / grafico |
| MAN_0003 | lote 8177 FORA_DO_CQ (ADR-091) | LOTE_FORA_DO_CQ | "lote fora do CQ (regra FORA_DO_CQ da tblLotesMaterial; fora da estatística e do gráfico)" | sem acento |
| MAN_0004 | lote 8353 SO_ANALITOS + Glicose | LOTE_OUTRO_MATERIAL | "lote de outro material neste analito (regra SO_ANALITOS …; fora da estatística e do gráfico)" | sem acento |
| MAN_0005 | lote 6354 (URINA) + Colesterol total, sem de/para | LOTE_OUTRO_MATERIAL | "manual com analito do catálogo sem cadastro na matriz do lote (…; fora da estatística e do gráfico)" | catalogo / estatistica / grafico |
| MAN_0006 | lote 6354 + "Creatinina" (ADR-093) | ATIVO, ANALITO CRE2 | "analito 'Creatinina' digitado num lote de URINA: lançado como 'CRE2' (…)" | lancado |
| MAN_0007 | SUBSTITUI_ID válido (S1) | ATIVO; S1 → SUBSTITUIDO_MANUAL | "substitui o resultado do interfaceamento BIO-356623" | igual |
| MAN_0008 | mesmo SUBSTITUI_ID de outro manual, +2 min | CONFLITO_MANUAL | "SUBSTITUI_ID inválido: outro manual (MAN_0007) já substitui BIO-356623 (este manual não substitui nada)" | ja / nao |
| MAN_0009 | SUBSTITUI_ID 999999999 | ATIVO | "SUBSTITUI_ID inválido: o ID BIO-999999999 não existe nos resultados do interfaceamento (este manual não substitui nada)" | nao ×2 |
| MAN_0010 | SUBSTITUI_ID = MAN_0002 | ATIVO | "SUBSTITUI_ID inválido: MAN_0002 **e** resultado manual (SUBSTITUI_ID aceita só resultado do interfaceamento) (este manual não substitui nada)" | e / so / nao |
| MAN_0011 | SUBSTITUI_ID S2 e NIVEL trocado para 2 | CONFLITO_MANUAL | "coincide com resultado do interfaceamento (3 min); SUBSTITUI_ID inválido: BIO-355801 tem outra chave: … (este manual não substitui nada)" | nao |
| MAN_0012 | SUBSTITUI_ID S3 sem MOTIVO | MANUAL_INCOMPLETO | "campos obrigatórios ausentes/inválidos: MOTIVO (obrigatório para substituir um resultado)" | igual (já acentuado pelo ADR-090) |
| MAN_0013 | data 31/12/2026 | MANUAL_INCOMPLETO | "campos obrigatórios ausentes/inválidos: DATA/HORA posterior ao lançamento" | igual |
| MAN_0014 | +5 min de R | CONFLITO_MANUAL | "coincide com resultado do interfaceamento (4 min)" | igual |
| MAN_0016 | repetição exata de MAN_0015 | CONFLITO_MANUAL | "resultado manual repetido (mesma chave e instante)" | igual |
| MAN_0017 | sem NIVEL | MANUAL_INCOMPLETO | "campos obrigatórios ausentes/inválidos: NIVEL" | igual |
| MAN_0001 / MAN_0015 | ATIVO comum | ATIVO | "" (vazio) | igual |

**Oráculo próprio.** Ele não usa o `camada_dados`:
- uma lista de 33 palavras sem acento, procuradas por palavra inteira, maiúscula ou minúscula, ignorando tokens todos em
  maiúsculas (nomes de coluna e de status);
- o padrão "`X e resultado manual`" / "`X e retransmiss…`", com o verbo "é" sem acento;
- a comparação `DETALHE == MOTIVO` com os meus acentos aplicados.

Resultado na entrega 42:
- `DETALHE == MOTIVO acentuado por mim` vale nos 17 manuais;
- nenhuma palavra da lista sobrou;
- a única marca encontrada é "MAN_0002 e resultado manual" (MAN_0010).

Na base, 9 manuais tinham palavra sem acento.

**Critério "status inalterado".** Comparei base × nova manual a manual:
- STATUS, MOTIVO_EXCLUSAO_AUTOMATICA, RUN e as colunas da tblCQ_Final (STATUS_ANALITICO, PARTICIPA_ESTATISTICA,
  ANALITO, MATRIZ, EQUIPAMENTO, RUN) são **idênticos** nos 17 manuais. A exceção é o EQUIPAMENTO do MAN_0001, que é o
  efeito esperado da OC-0068.
- O DETALHE só difere por acento (`DETALHE` sem acento da nova = DETALHE da base).
- O MOTIVO_EXCLUSAO_AUTOMATICA continua **sem acento** nos dois arquivos, como o ADR-090/095 manda.
- O S1 substituído é igual nas duas: "substituido pelo manual MAN_0007 (motivo: Outro: TESTE-AUDITORIA)", fora da
  estatística, sem RUN.

**RUN.**
- Nenhum RUN ou DETALHE "0" nas 17 linhas (varredura da aba).
- O RUN fica vazio nos CONFLITO e nos MANUAL_INCOMPLETO.
- IMPLAUSIVEL e LOTE_* têm o RUN do dia (26092501), igual à base. É a corrida em que o resultado foi lançado, não um 0
  indevido.

**O que persiste.** O verbo "é" fica sem acento no texto do ADR-082 "`<ID> e resultado manual (SUBSTITUI_ID aceita só
resultado do interfaceamento)`". A origem é o `mEtlCQ.RazaoSubstituicao`. O `ACENTOS_DETALHE` só troca palavras, e
trocar " e " por " é " quebraria a conjunção ("da estatística e do gráfico", "mesma chave e instante").

Pela mesma leitura do `mEtlCQ.bas`, há mais dois casos que eu não consegui provocar:
- "`<ID> e retransmissao de <ID> (informe o ID original)`": vira "e retransmissão", ainda sem o "é". Não existe nenhum
  DUPLICIDADE_ORIGEM na cópia.
- "`ID_REGISTRO manual repetido (vale a 1a linha)`": exige um MAN_ repetido, e o ID é carimbado sozinho.

A trava `qa_adr095` D01 não pega esses casos, porque o "e" não está em `SEM_ACENTO_PROIBIDO`. Ver **R5P2-01**.

**Fluxo vizinho:**
- Os textos que o ADR-090 já acentuava continuam iguais: MANUAL_INCOMPLETO (NIVEL, MOTIVO, DATA/HORA posterior ao
  lançamento), CONFLITO e repetido.
- A mensagem do ATUALIZAR separa os grupos: "3 resultado(s) manual(is) incompleto(s)", "4 manual(is) com SUBSTITUI_ID
  inválido: nada foi substituído — veja o DETALHE", "4 erro(s) de governança". Os 4 E12 correspondem a MAN_0008, 0009,
  0010 e 0011.
- Status, PARTICIPA e RUN são idênticos aos da base: não houve efeito colateral.

### OC-0068: listas de escolha com um item só

**Fluxo original** (`O68_listas`). O separador de lista do Excel da sessão é ";".

| Lista | Base (entrega 41) | Nova (entrega 42) |
|---|---|---|
| Digitar Resultados EQUIPAMENTO | menu `["DIMENSION 1,DIMENSION 2"]`; "DIMENSION 1"/"DIMENSION 2" **recusados**; o texto juntado aceito; msg "Escolha o equipamento da lista." | menu `["DIMENSION 1","DIMENSION 2"]`; os dois aceitos; "DIMENSION 9", "DIMENSION" e o texto juntado recusados; msg "Escolha o equipamento **na seta da célula**." (estilo Parar) |
| Digitar Resultados MATRIZ | `["SORO,URINA"]`; SORO e URINA recusados | `["SORO","URINA"]`; aceitos; "PLASMA" e "SORO,URINA" recusados; msg "Escolha a matriz na seta da célula." |
| Digitar Resultados MOTIVO | um item só; os 4 motivos recusados | 4 itens; os 4 aceitos; "Outro: TESTE-AUDITORIA", "Falha" e o texto juntado recusados; msg "Escolha o motivo na seta da célula. Ele é obrigatório quando o manual substitui um resultado." |
| Registros H4 "Calibrado?" | `["Sim,Não"]`; Sim e Não recusados | `["Sim","Não"]`; aceitos; "Sim,Não", "Talvez" e "S" recusados |
| Painel L4 | já estava certa: 2 itens | 2 itens; os dois aceitos; inválidos recusados |

Vizinhos sem mudança: LOTE (`=lstLotes`), ANALITO (`=lstAnalitos`), NIVEL (1 a 2), DATA (≥ 01/01/2020), RESULTADO
(decimal), Registros F7 "Foi calibrado?" (Sim;Não), Painel K4 e M4.

**Manual pelo item do menu** (critério "entra na série certa"):
- **MAN_0001**: EQUIPAMENTO = o item do menu que contém "DIMENSION 2"; Glicose N1, lote 8977, 07/10/2026 01:00,
  123,456.
- **Nova:** o item é "DIMENSION 2" e o manual fica ATIVO, PARTICIPA SIM, EQUIPAMENTO DIMENSION 2, RUN 26100701. Painel
  em DIMENSION 2 · Glicose · 8977 · Out/2026:
  - N1 passou de n 23, média 76,217391304 para **n 24, média 78,185666667, DP 9,682841402, CV 12,38442008 %**;
  - o valor 123,456 está entre os pontos do Eng_Saida.
- **Conferência independente da média:** 23 × 76,217391304 = 1753,000; (1753 + 123,456)/24 = 78,1856667. Bate.
- **Série DIMENSION 1 do mesmo lote e mês:** engPainel idêntico antes e depois (N1 n 25, média 77,36; N2 n 22).
- **Base (o defeito):** o item é "DIMENSION 1,DIMENSION 2". O MAN_0001 fica ATIVO, PARTICIPA SIM, com equipamento
  inexistente, e as duas séries continuam com n 23 e n 25. Isso reproduz o R4-01.

**XML (L03, independente).** A busca por `x12ac:list`:
- encontrou 4 na base (Registros H4, Digitar Resultados D/E/K);
- encontrou **0** na cópia da produção nova;
- encontrou **0** no `SaveCopyAs` feito depois do login, das listas, do filtro, dos 17 manuais e do ATUALIZAR.

**Filtros só mudam a vista** (`O68_O72_registros`). A tblRegistros tem 3 linhas: 2 registros de Proteína total sem
"Foi calibrado?" e a linha livre.
- H4 = Sim deixa 1 de 3 visível; H4 = Não, 1 de 3; H4 vazio, 3 de 3.
- O hash do `Value2` da tblRegistros é o mesmo em todos os passos (`fb1e82c46115fbb2`).
- No Painel, trocar L4 entre DIMENSION 1 e 2 muda só o engPainel. A série DIMENSION 1 é idêntica antes e depois.

### OC-0072: ano literal nos formatos de data

**Fluxo original.** Depois do login, li `Text`, `NumberFormat` e `NumberFormatLocal` das células não vazias de cada
alvo. Repeti a leitura depois do ATUALIZAR, que grava uma linha nova no Sync_Log, refaz o Westgard_Fotos e o Eng_Saida.

| Alvo | Base (entrega 41) | Nova (entrega 42), após login e após ATUALIZAR |
|---|---|---|
| Configuração D26 (30/06/2026) | "30/06/yyyy" | "30/06/2026" (`dd/mm/aaaa`) |
| Configuração AD26 | Text "" (coluna AD oculta no XML: `hidden="1"`) | idem. Não é defeito: a coluna é oculta |
| Eng_Saida AC3:AC182 (126 → 110 datas) | "##############" (o texto com "yyyy" não cabe na coluna) | "07/07/2026 08:37:54" … (`dd/mm/aaaa hh:mm:ss`), 0 ruins |
| Registros E4/F4 (De/Até digitados 01/10/2026 e 31/10/2026) | "01/10/yyyy" e "31/10/yyyy" | "01/10/2026" e "31/10/2026" |
| Westgard_Fotos F (início e fim, 78 células) | já estava certo nas linhas com dado; só o estilo da coluna tinha ano literal | "14/11/2026 07:16:19" … e, depois do ATUALIZAR, "15/11/2026 05:45:46"; formato da coluna inteira `dd/mm/aaaa hh:mm:ss` |
| Sync_Log A/B/L/M (linhas 4–22; 23 depois do ATUALIZAR) | linhas com dado certas; o estilo das linhas vazias 4006+ tinha ano literal | todas `dd/mm/aaaa hh:mm:ss`; a linha nova "15/11/2026 05:45:20" |
| Estatística E8:F9, C10 | vazias, formato `dd/mm/yyyy` (literal) | vazias, formato `dd/mm/aaaa` |
| LotesStore U1 (achado a mais do XML) | estilo com ano literal | `dd/mm/aaaa hh:mm:ss` |

**XML (L04, independente: `xml_formatos_listas.py`).**
- **Base:** 3 numFmt com letra escapada (`dd/mm/\y\y\y\y`, também com ` hh:mm` e ` hh:mm:ss`), usados em 8 estilos e 453
  células ou colunas: Configuração 200, Eng_Saida 180, Registros 2, Estatística 5, Sync_Log 64, Westgard_Fotos F3 e
  coluna F, LotesStore U1.
- **Nova e SaveCopyAs depois do uso:** **0** numFmt com letra literal e **0** células.

**Valores inalterados.** Comparei o `<v>` bruto do XML célula a célula, base × nova, nos alvos:
- Configuração: 200 de 200 iguais.
- Eng_Saida: 180 de 180.
- Registros: 2 de 2.
- Estatística: 5 de 5.
- Westgard_Fotos F: 13.701 de 13.701. As 33 a mais da base são células vazias só com estilo.
- LotesStore: 1 de 1.
- Sync_Log: 16.076 células comuns, todas iguais. A nova tem uma linha a mais (22), que é o ATUALIZAR da própria entrega
  (15/11/2026 00:17).

**Fluxo vizinho.** O `SaveCopyAs` feito depois de login, listas, filtros, manuais e ATUALIZAR mostra que nenhum desses
passos recria o ano literal em nenhuma aba. No Digitar Resultados, a DATA e o REGISTRADO_EM dos 17 manuais também não
têm formato literal. Os formatos de data do arquivo são todos `dd/mm/yyyy[ hh:mm:ss]`, `yyyy-mm-dd` ou o interno 14.

## Ocorrências novas

### R5P2-01: DETALHE da substituição com o verbo "é" sem acento (resíduo da OC-0021)

- **Prioridade:** Baixo · **Tipo:** usabilidade · **Tela/módulo:** aba Digitar Resultados, coluna DETALHE. A origem é o
  `mEtlCQ.RazaoSubstituicao` (e o mesmo texto no PQ). A tela usa `camada_dados.ACENTOS_DETALHE`.
- **Ação:** digitar no SUBSTITUI_ID o ID de outro manual (MAN_0002), completar a linha e clicar em ATUALIZAR DADOS.
- **Esperado:** "SUBSTITUI_ID inválido: MAN_0002 **é** resultado manual (SUBSTITUI_ID aceita só resultado do
  interfaceamento) (este manual não substitui nada)".
- **Observado:** "… MAN_0002 **e** resultado manual …". Pela leitura do código, há mais dois casos que não consegui
  provocar na cópia:
  - "`<ID> e retransmissão de <ID>`" (não há DUPLICIDADE_ORIGEM na base);
  - "`ID_REGISTRO manual repetido (vale a 1a linha)`".
- **Dificuldade:** pouca. Lido rápido, "MAN_0002 e resultado manual" parece uma frase incompleta.
- **Consequência:** só de leitura. Status, MOTIVO e números são iguais aos da base.
- **Sugestão:** acrescentar ao `ACENTOS_DETALHE` pares com contexto, que não pegam a conjunção: `(' e resultado
  manual', ' é resultado manual')`, `(' e retransmiss', ' é retransmiss')`, `('1a linha', '1ª linha')`. A outra opção é
  reescrever o texto na origem nos dois motores. Também cabe estender a trava D01 a esses padrões.
- **Evidência:** `Bioquimica_F1_nova.json`, `O21_O68_manuais.manuais.M10_subst_id_manual` (DETALHE_txt e `sem_acento`);
  `mEtlCQ.bas` linhas 808, 816 e 1240.

### R5P2-02: Registros H4 e Painel L4 recusam o valor inválido com a mensagem genérica do Excel

- **Prioridade:** Baixo · **Tipo:** usabilidade · **Tela:** Registros H4 ("Calibrado?") e F7:F9; Painel L4.
- **Ação:** digitar um valor fora da lista (ex.: "Talvez" no H4, "DIMENSION 9" no L4).
- **Esperado:** uma mensagem como a da Digitar Resultados ("Escolha … na seta da célula.").
- **Observado:** `Validation.ErrorMessage` vazio nas duas células (`msg = ''`), na base e na nova. O Excel mostra o texto
  padrão dele ("O valor não corresponde às restrições de validação…"). A recusa funciona: `Validation.Value` = Falso.
- **Consequência:** nenhuma nos dados. Só a orientação é pior que a das listas da Digitar Resultados.
- **Sugestão:** pôr `ErrorMessage` "Escolha na seta da célula." nos filtros literais (`instalar_adr087.FILTROS`) e no
  L4 (`instalar_integracao`).
- **Evidência:** `Bioquimica_F1_nova.json`, `O68_listas.Registros_H4.msg` e `Painel_L4.msg`.
- **Não executado:** ver a caixa na tela.

## Dúvidas de regra de negócio

1. **OC-0072:** Configuração D/AD, Registros E4/F4 e Estatística E8:F9/C10 passaram para o formato interno 14 (data
   abreviada do Windows). Ele mostra dd/mm/aaaa no Windows em português, mas segue a configuração regional da máquina. O
   Eng_Saida, o Sync_Log e o Westgard_Fotos usam `dd/mm/yyyy` explícito. Todas as máquinas do laboratório estão em
   pt-BR? Se não estiverem, o explícito é mais previsível.
2. **OC-0021:** o RUN do dia aparece também no IMPLAUSIVEL e nos LOTE_* (fora da estatística), como na base. Isso é
   intencional, para mostrar em que corrida o resultado foi lançado? Se for, não há o que mudar.
3. **Risco residual (já registrado no ADR-095 §1, "não feito"):** colar um texto por cima da validação continua gravando
   um equipamento fora do cadastro, e o manual entra ATIVO. A regra ou QA que impede isso depende de autorização.

## Impressões de uso (Relatório A)

- **Listas:** agora a seta mostra uma opção por linha e a digitação do valor certo funciona. É o que o GUIA §3 promete
  ("escolha o certo na seta da célula"). Corrigir o equipamento âmbar ficou possível sem truque. O manual entra na série
  certa do Painel, e a média muda o quanto deve mudar.
- **DETALHE:** o texto ficou legível e coerente com o resto da tela (implausível, estatística, gráfico, catálogo,
  lançado, já, não, só). O único tropeço é "MAN_0002 e resultado manual". O DETALHE da substituição diz exatamente por
  que nada foi substituído (ID inexistente, ID manual, outra chave, já substituído) e casa com a mensagem do ATUALIZAR
  ("4 manual(is) com SUBSTITUI_ID inválido").
- **Datas:** o De/Até do Registros mostra a data digitada (antes, "01/10/yyyy"). A coluna de data/hora do Eng_Saida, que
  antes mostrava "#####", ficou legível com os segundos. Sync_Log e Westgard_Fotos estão coerentes.
- **Mensagens:** as da Digitar Resultados são claras. As do H4 e do L4 ainda são as genéricas do Excel (R5P2-02).

## Arquivos

- Scripts: `_arquitetura/agentes/ferramentas/reteste05p2_usuario_f1.py` e
  `_arquitetura/agentes/ferramentas/xml_formatos_listas.py`.
- Saídas (fora do repositório): `%TEMP%\qcw\auditoria\2026-11-15_reteste05_parte2\Bioquimica_F1_{base41,nova}.{json,log}`,
  `xml_resumo.json`, `xml_alvos_comparacao.json`. As cópias `.xlsm` foram apagadas.
- O Audit_Log e o Sync_Log da cópia gravam o login do Windows de quem roda. Esse dado foi omitido deste relatório.
