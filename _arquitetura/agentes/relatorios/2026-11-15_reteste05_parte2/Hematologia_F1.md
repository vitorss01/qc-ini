# Reteste 05, parte 2: Usuário Simulado, Hematologia, frente F1

Data: 15/11/2026 · Agente: AGENTE_USUARIO_SIMULADO · Modo: RETESTE (OC-0021, OC-0068, OC-0072) · Produção: entrega 42.

## Ambiente e método

| Item | Valor |
|---|---|
| Produção atual (só copiada) | `QC_Hematologia.xlsm` sha256 curto **7bab8fc008f6**, igual antes e depois (`origem_intacta.igual = true`; conferido de novo no fim) |
| Base de comparação ("antes") | `_backup_pre_integracao_2026-11-15_0458/QC_Hematologia.xlsm` sha256 curto **aab0a54f464e** (entrega 41), igual antes e depois |
| Cópias (novas, nenhuma salva, apagadas no fim) | `%TEMP%\qcw\auditoria\2026-11-15_reteste05_parte2\Hematologia_F1_base41.xlsm` e `Hematologia_F1_nova.xlsm`; para o XML sem Excel, `Hematologia_F1_xml_{base41,nova}.xlsm`; e `Hematologia_F1_nova_SaveCopyAs.xlsm`, gravado da cópia em uso só para ler o XML depois do uso |
| Excel | um por vez, abertos e encerrados pelo `qa_final.QA`: PID 16116 (base) e PID 29532 (nova). Não toquei no 13084 nem nos que a tarefa agendada abriu durante o trabalho (11360, 644). Nenhuma captura de tela |
| Script (novo, documentado no cabeçalho) | `_arquitetura/agentes/ferramentas/reteste05p2_usuario_f1_hema.py <base41\|nova\|xml>`. Reaproveita do irmão da Bio (`reteste05p2_usuario_f1.py`) o oráculo de acentos, a leitura de datas, a avaliação de listas e o filtro do Registros; e o `xml_formatos_listas.py` |
| Saídas (fora do repositório) | `Hematologia_F1_base41.json/.log`, `Hematologia_F1_nova.json/.log`, `Hematologia_F1_xml.json`, na pasta do ciclo no `%TEMP%` |
| Usuário | `TESTE_AUDITORIA_R5P2H`, papel ANALISTA, criado só na cópia; login pelo `mSeguranca.DoLogin` |
| Dados de teste | 17 linhas manuais (MAN_0001..0016 e uma linha colada), WBC, valores 123,456 / 111,111 / 12,345 / −1; MOTIVO "Outro: TESTE-AUDITORIA", exceto nos dois casos da OC-0068, que usam o item do menu; SUBSTITUI_ID 999999999 (inexistente); duas regras de lote com NOTA "TESTE-AUDITORIA" (ver abaixo) |
| tblCQ_Final | 48.391 linhas nas duas cópias; o ATUALIZAR leu o DB_SEAC só para leitura e não trouxe resultado novo (último recebido 13/11/2026 13:10:20) |

**Caminho usado.**
1. Login.
2. Painel: C3, H3, K3 e K4 escritos com os eventos ligados.
3. `mIntegracao.IrDigitarResultados`, a navegação do botão. Levou à aba Digitar Resultados, célula B7.
4. Manuais digitados célula a célula (`FormulaLocal`, eventos ligados). Na substituição, digitei primeiro o
   SUBSTITUI_ID (GUIA §3).
5. Uma linha inteira copiada e colada (`Range.Copy`), como o Ctrl+C / Ctrl+V do usuário.
6. **ATUALIZAR DADOS** pelo núcleo do botão (`mIntegracao.AtualizarDadosAutomatico`).
7. Na aba Registros: filtro H4 e De/Até em E4/F4, digitados.

A mesma sequência rodou na base (entrega 41) e na nova (entrega 42).

**Fora do caminho do usuário comum (declarado):** a Hema não tem nenhuma regra na `tblLotesMaterial`. Sem regra, os
textos do ADR-091 não aparecem. Para provocá-los, cadastrei na cópia duas regras, como o GUIA §6 manda o
administrador fazer (aba Cfg_Integracao):
- `602911` FORA_DO_CQ;
- `522611` SO_ANALITOS = "RBC".

São lotes antigos de WBC. Cadastrei as mesmas duas regras nas duas cópias. O `mEtlCQ.LotesMaterialResumo` devolveu
"MATRIZ=0;FORA_DO_CQ=1;SO_ANALITOS=1". Efeito nos reais da cópia: 3.751 resultados do 602911 viraram LOTE_FORA_DO_CQ,
e 4.560 do 522611 (fora o RBC) viraram LOTE_OUTRO_MATERIAL. A troca de analito de urina (ADR-093) só existe na Bio e não
se aplica.

**Não executado (sem interação com a tela):**
- O clique físico na seta e a caixa "Parar". Usei o veredito do próprio Excel (`Range.Validation.Value`).
- As caixas de aviso do SUBSTITUI_ID. O `xlh` abre o Excel com `Interactive = False`, e nenhuma caixa apareceu nas duas
  cópias.

## Resultado por ocorrência

| OC | Estado | Resumo |
|---|---|---|
| OC-0021 | **PERSISTE** (residual pequeno) | Os textos sem acento da base agora aparecem acentuados: implausível, estatística, gráfico, já, não, só. STATUS, MOTIVO, RUN e tblCQ_Final são idênticos aos da base. Sobra o verbo "é" sem acento em **dois** textos, os dois provocados na cópia: "MAN_0002 **e** resultado manual" e "HEM-212660 **e** retransmissão de HEM-212425". O segundo a Bio não conseguiu provocar |
| OC-0068 | **RESOLVIDA** | MOTIVO e Registros H4 mostram um item por valor. Válidos aceitos; inválidos e o texto juntado recusados. EQUIPAMENTO e MATRIZ, que têm um item só, continuam certos. O manual com os itens do menu entra na série certa (WBC N1 625311 Out/2026: n 21 → 22, média conferida). A substituição com o motivo do menu grava o motivo certo. Nenhum `x12ac:list` no XML. O filtro do Registros só muda a vista. Resíduo: o H4 recusa com a mensagem genérica do Excel (R5P2-02, que vale também para a Hema) |
| OC-0072 | **RESOLVIDA** | Nenhuma célula-alvo mostra "yyyy". As datas aparecem em dd/mm/aaaa, com :ss onde o formato tem hora. O XML, antes e depois do uso, não tem formato com ano literal. Os valores das células são iguais aos da base |

### OC-0021: DETALHE sem acento

**Fluxo original.** Lancei manuais que geram cada texto possível de MOTIVO_EXCLUSAO_AUTOMATICA na Hema. Depois rodei
ATUALIZAR e li o DETALHE (`Text`), o RUN (`Text`), o STATUS e a linha da tblCQ_Final.

Referências reais usadas, todas WBC N1, XN1000:
- R = HEM-357971 (lote 619611, 29/09/2026 08:19:39, RUN 26092905);
- S1 = HEM-357896, S2 = HEM-357847 e S3 = HEM-357762, do mesmo lote e dia;
- R2 = HEM-364445 (lote 625311, 09/10/2026 10:32:44);
- DUP = HEM-212660 (BASO# N1, lote 522611, DUPLICIDADE_ORIGEM de HEM-212425). Há 589 DUPLICIDADE_ORIGEM na cópia.

| Manual | Cenário | STATUS (igual na base) | RUN | DETALHE na entrega 42 (Text) | Base (entrega 41) |
|---|---|---|---|---|---|
| MAN_0001 | itens do menu (OC-0068) | ATIVO | 26100901 | "" | igual |
| MAN_0002 | −1 (ADR-088) | IMPLAUSIVEL | 26092901 | "resultado <= 0: implausível para material de controle (fora da estatística e do gráfico)" | implausivel / estatistica / grafico |
| MAN_0003 | lote 602911 FORA_DO_CQ | LOTE_FORA_DO_CQ | 26092901 | "lote fora do CQ (regra FORA_DO_CQ da tblLotesMaterial; fora da estatística e do gráfico)" | estatistica / grafico |
| MAN_0004 | lote 522611 SO_ANALITOS (RBC) + WBC | LOTE_OUTRO_MATERIAL | 26092901 | "lote de outro material neste analito (regra SO_ANALITOS da tblLotesMaterial; fora da estatística e do gráfico)" | estatistica / grafico |
| MAN_0005 | SUBSTITUI_ID S1, motivo do menu | ATIVO; S1 → SUBSTITUIDO_MANUAL | 26092906 | "substitui o resultado do interfaceamento HEM-357896" | igual |
| MAN_0006 | mesmo SUBSTITUI_ID, +2 min | CONFLITO_MANUAL | "" | "SUBSTITUI_ID inválido: outro manual (MAN_0005) já substitui HEM-357896 (este manual não substitui nada)" | ja / nao |
| MAN_0007 | SUBSTITUI_ID 999999999 | ATIVO | 26092901 | "SUBSTITUI_ID inválido: o ID HEM-999999999 não existe nos resultados do interfaceamento (este manual não substitui nada)" | nao ×2 |
| MAN_0008 | SUBSTITUI_ID = MAN_0002 | ATIVO | 26092902 | "SUBSTITUI_ID inválido: MAN_0002 **e** resultado manual (SUBSTITUI_ID aceita só resultado do interfaceamento) (este manual não substitui nada)" | e / so / nao |
| MAN_0009 | SUBSTITUI_ID S2 e NIVEL 2 | CONFLITO_MANUAL | "" | "coincide com resultado do interfaceamento (22 min); SUBSTITUI_ID inválido: HEM-357847 tem outra chave: XN1000 \| SANGUE TOTAL \| 619611 \| 1 \| WBC (este manual: … \| 2 \| WBC) (este manual não substitui nada)" | nao |
| MAN_0010 | SUBSTITUI_ID S3 sem MOTIVO | MANUAL_INCOMPLETO | "" | "campos obrigatórios ausentes/inválidos: MOTIVO (obrigatório para substituir um resultado)" | igual |
| MAN_0011 | SUBSTITUI_ID = DUP (retransmissão) | CONFLITO_MANUAL | "" | "coincide com resultado do interfaceamento (0 min); SUBSTITUI_ID inválido: HEM-212660 **e** retransmissão de HEM-212425 (informe o ID original) (este manual não substitui nada)" | e / nao |
| MAN_0012 | data 31/12/2026 | MANUAL_INCOMPLETO | "" | "campos obrigatórios ausentes/inválidos: DATA/HORA posterior ao lançamento" | igual |
| MAN_0013 | +5 min de R | CONFLITO_MANUAL | "" | "coincide com resultado do interfaceamento (4 min)" | igual |
| MAN_0014 / MAN_0015 | repetição exata (N3) | ATIVO / CONFLITO_MANUAL | 26092901 / "" | "" / "resultado manual repetido (mesma chave e instante)" | igual |
| MAN_0016 | sem NIVEL | MANUAL_INCOMPLETO | "" | "campos obrigatórios ausentes/inválidos: NIVEL" | igual |
| linha 17 (colada do MAN_0007) | Ctrl+C / Ctrl+V da linha inteira | tela: ATIVO; tblCQ_Final: `MAN_0007~L17` CONFLITO_MANUAL | tela: 26092901 | tela: igual ao MAN_0007; tblCQ_Final: "ID_REGISTRO manual repetido (vale a 1a linha)" | igual (ver R5P2-H03) |

**Oráculo próprio** (o mesmo do irmão da Bio, que não usa o `camada_dados`):
- 33 palavras sem acento procuradas por palavra inteira, ignorando os tokens todos em maiúsculas;
- o padrão "`X e resultado manual`" / "`X e retransmiss…`" (verbo "é");
- o ordinal "1a linha";
- `DETALHE == MOTIVO` com os meus acentos aplicados.

Resultado na entrega 42:
- `DETALHE == MOTIVO acentuado por mim` vale nas 17 linhas;
- nenhuma palavra da lista sobrou;
- sobram só o "MAN_0002 e resultado manual" (MAN_0008) e o "HEM-212660 e retransmissão" (MAN_0011).

Na base, 9 linhas tinham palavra sem acento.

**Critério "status inalterado".** Comparei base × nova linha a linha:
- STATUS, RUN, MOTIVO_EXCLUSAO_AUTOMATICA e as colunas da tblCQ_Final (STATUS_ANALITICO, PARTICIPA_ESTATISTICA,
  ANALITO, MATRIZ, EQUIPAMENTO, LOTE, NIVEL, RUN) são **idênticos** nas 17 linhas.
- O DETALHE sem acento da nova é igual ao DETALHE da base.
- O MOTIVO_EXCLUSAO_AUTOMATICA continua sem acento nos dois arquivos, como o ADR-095 §2 manda.
- A mensagem do ATUALIZAR e a contagem do QA são iguais nas duas: "8 erro(s), 8914 alerta(s), 35 informativo(s)",
  "3 resultado(s) manual(is) incompleto(s)" e "5 manual(is) com SUBSTITUI_ID inválido". A contagem alta de alertas vem
  das regras TESTE.

**RUN.**
- Nenhum RUN ou DETALHE "0" nas 17 linhas (varredura da aba).
- O RUN fica vazio nos CONFLITO e nos MANUAL_INCOMPLETO.
- O IMPLAUSIVEL e os LOTE_* têm o RUN do dia (26092901), igual à base.

**O que persiste.** O verbo "é" fica sem acento em dois textos do `mEtlCQ.RazaoSubstituicao`:
- `mEtlCQ.bas:808`: "`<ID> e resultado manual (...)`";
- `mEtlCQ.bas:816`: "`<ID> e retransmissao de <ID> (informe o ID original)`". O `ACENTOS_DETALHE` acentua
  "retransmissão", mas não o "é".

O mesmo texto aparece na aba QA_INTEGRACAO, sem acento de propósito (E12, A01). A trava `qa_adr095` D01 não pega esse
caso, porque o "e" não está em `SEM_ACENTO_PROIBIDO`. Ver **R5P2-H01**, extensão do R5P2-01 da Bio.

**Fluxo vizinho:**
- Os textos que o ADR-090 já acentuava continuam iguais: MANUAL_INCOMPLETO, CONFLITO e repetido.
- Os E12 da QA_INTEGRACAO são 5 (MAN_0006, 0007, 0008, 0009, 0011) e casam com a mensagem do ATUALIZAR.
- Os A13 citam MAN_0003 e MAN_0004 com a regra; o A11 cita o MAN_0002.
- O caso "vale a 1a linha" (linha colada) foi provocado. O texto existe só na tblCQ_Final e na QA_INTEGRACAO
  (A01 `MAN_0007~L17`). A aba Digitar Resultados nunca o mostra: a linha colada exibe o status do MAN_0007 original
  (ATIVO). Isso não depende da entrega e é igual na base (**R5P2-H03**).

### OC-0068: listas de escolha com um item só

**Fluxo original** (`O68_listas`). O separador de lista do Excel da sessão é ";". Os itens válidos do oráculo vêm dos
dados (tblCQ_Final: equipamento XN1000, matriz SANGUE TOTAL) e do GUIA/ADR-082 (os 4 motivos), não da própria
validação.

| Lista | Base (entrega 41) | Nova (entrega 42) |
|---|---|---|
| Digitar Resultados MOTIVO | um item só (`"Falha do interfaceamento,Resultado não transmitido,Equipamento em contingência,Outro"`); os 4 motivos **recusados**, o texto juntado aceito; msg "Escolha o motivo da lista." | 4 itens; os 4 aceitos; "Outro: TESTE-AUDITORIA", "Falha", "outro" e o texto juntado recusados; estilo Parar; msg "Escolha o motivo na seta da célula. Ele é obrigatório quando o manual substitui um resultado." |
| Digitar Resultados EQUIPAMENTO | `["XN1000"]` (já certo, um equipamento só); msg "Escolha o equipamento da lista." | `["XN1000"]`; aceito; "XN9999", "XN" e "XN1000,XN1000" recusados; msg "Escolha o equipamento na seta da célula." |
| Digitar Resultados MATRIZ | `["SANGUE TOTAL"]` (já certo) | igual; "SORO", "PLASMA" e "SANGUE" recusados; msg "Escolha a matriz na seta da célula." |
| Registros H4 "Calibrado?" | `["Sim,Não"]`; Sim e Não recusados | `["Sim","Não"]`; aceitos; "Sim,Não", "Talvez" e "S" recusados; **msg vazia** (texto padrão do Excel) |
| Registros F7 "Foi calibrado?" | já estava certo (Sim;Não), msg vazia | igual |
| Painel L4 | não existe na Hema (sem validação, L3/L4 vazios). Não se aplica | idem |

Vizinhos sem mudança: LOTE (`=lstLotes`), ANALITO (`=lstAnalitos`), NIVEL (1 a 3), DATA (≥ 01/01/2020), RESULTADO
(decimal), SUBSTITUI_ID (sem validação); no Painel, C3, H3, K3 e K4.

**Manual pelo item do menu** (critério "entra na série certa"):
- **MAN_0001:** EQUIPAMENTO, MATRIZ e MOTIVO preenchidos com o item do menu. Na nova, os itens são "XN1000",
  "SANGUE TOTAL" e "Falha do interfaceamento". WBC N1, lote 625311, 09/10/2026 01:00, valor 123,456.
- **Nova:** o manual fica ATIVO, PARTICIPA SIM, RUN 26100901. Na tela, o MOTIVO gravado é "Falha do interfaceamento".
  No Painel (WBC · 625311 · 2026 · Out):
  - N1 passou de n 21, média 3,12047619 para **n 22, média 8,590272727, DP 25,655674351, CV 298,65960215 %**;
  - o 123,456 aparece entre os pontos do Eng_Saida;
  - N2 (n 39) e N3 (n 21) ficaram idênticos.
- **Conferência independente** (Python `statistics` sobre os pontos do Eng_Saida):
  - antes: n 21, média 3,1204761904761904, DP 0,054173442799603455, CV 1,736063327928693 %;
  - depois: n 22, média 8,590272727272728, DP 25,655674350877398, CV 298,65960215005487 %.
  - Batem com o engPainel no 9º decimal (o leitor arredonda em 9 casas). Pela conta, (21 × 3,12047619 + 123,456)/22 =
    8,5902727.
- **Base (o defeito):** o MOTIVO gravado é o texto juntado. Na Hema, isso não muda a série, porque há um equipamento
  só.

**Substituição com o motivo do menu** (vizinho, ADR-082). O MAN_0005 usou o SUBSTITUI_ID S1 e o MOTIVO "Resultado não
transmitido" escolhido no menu.
- **Nova:** S1 = SUBSTITUIDO_MANUAL, PARTICIPA NÃO, sem RUN, "substituido pelo manual MAN_0005 (motivo: Resultado não
  transmitido)".
- **Base:** "(motivo: Falha do interfaceamento,Resultado não transmitido,Equipamento em contingência,Outro)", no I04 e no
  Audit_Log. É a trilha sem sentido que o R4-01 previa, agora corrigida.

**XML (L03, independente).** A busca por `x12ac:list`:
- encontrou 2 na base (Registros H4; Digitar Resultados K7:K306, o MOTIVO);
- encontrou **0** na cópia da produção nova;
- encontrou **0** no `SaveCopyAs` feito depois do login, das listas, dos filtros, dos manuais e do ATUALIZAR.

Na nova, o XML grava F7:F9 e H4 na mesma regra `"Sim,Não"`. A vírgula é a forma de arquivo do Excel, não um item
único.

**Filtros só mudam a vista** (`O68_O72_registros`). A tblRegistros tem 3 linhas: 2 registros de WBC (lote 522611, um
"Não" e um "Sim" em "Foi calibrado?") e a linha livre.
- H4 = Sim deixa 2 de 3 visíveis (o "Sim" e a linha livre). H4 = Não, também 2 de 3. H4 vazio, 3 de 3.
- De/Até 01/10–31/10/2026 deixa 1 de 3 visível (só a linha livre: os registros são de 2025).
- O hash do `Value2` da tblRegistros é o mesmo em todos os passos (`1f2166a5755048ff`).

### OC-0072: ano literal nos formatos de data

**Fluxo original.** Depois do login, li `Text`, `NumberFormat` e `NumberFormatLocal` das células não vazias de cada
alvo. Repeti a leitura depois do ATUALIZAR, que reescreve o Eventos_Westgard, o Eng_Saida, o Westgard_Fotos e o
Sync_Log.

| Alvo | Base (entrega 41) | Nova (entrega 42), após login e após ATUALIZAR |
|---|---|---|
| Eventos_Westgard A4:A5003 (5.000 datas; 1.763 depois do ATUALIZAR) | "01/10/yyyy", "10/10/yyyy" … | "01/10/2025", "10/10/2025" … e, depois, "29/10/2026" … (`dd/mm/aaaa`). Só a data, sem hora: o ADR-094 tira a hora dessa coluna de propósito, porque o instante vem ao minuto |
| Configuração D26 (31/03/2026) | "31/03/yyyy" | "31/03/2026" (`dd/mm/aaaa`) |
| Configuração AD26 | Text "" (coluna AD oculta) | idem. Não é defeito: a coluna é oculta |
| Eng_Saida AC3:AC182 (56 → 41 datas) | "##############" (o texto com "yyyy" não cabe na coluna) | "01/10/2025 06:40:23" … e, depois, "01/10/2026 08:11:43" (`dd/mm/aaaa hh:mm:ss`); 0 ruins |
| Registros E4/F4 (De/Até digitados 01/10/2026 e 31/10/2026) | "01/10/yyyy" e "31/10/yyyy" | "01/10/2026" e "31/10/2026" |
| Westgard_Fotos F | linhas com dado já certas; o estilo da coluna (`<col 6>`) e o F3 com ano literal | "14/11/2026 10:24:53" … e, depois do ATUALIZAR, "15/11/2026 06:04:51"; formato da coluna inteira `dd/mm/aaaa hh:mm:ss` |
| Sync_Log A/B/L/M | já certas (`dd/mm/aaaa hh:mm:ss`) | iguais; a linha nova do ATUALIZAR com segundos |
| Estatística E8:F9, C10 | vazias, formato Geral (o alvo é só da Bio) | idem |
| LotesStore U1 (achado a mais do XML) | estilo com ano literal (a célula tem o texto "REGISTRADO_EM") | `dd/mm/aaaa hh:mm:ss` |

**XML (L04, independente: `xml_formatos_listas.py`).**
- **Base:** 2 numFmt com letra escapada (`dd/mm/\y\y\y\y` e `dd/mm/\y\y\y\y\ hh:mm`), usados em 5.384 células e 1
  coluna:
  - Eventos_Westgard: 5.000;
  - Configuração: 200;
  - Eng_Saida: 180;
  - Registros: 2;
  - LotesStore U1: 1;
  - Westgard_Fotos: F3 e coluna F.
- **Nova e SaveCopyAs depois do uso:** **0** numFmt com letra literal e **0** células. Os formatos de data do arquivo são
  todos `dd/mm/yyyy[ hh:mm:ss]`, `yyyy-mm-dd[ hh:mm:ss]` ou o interno 14.

**Valores inalterados.** Comparei o `<v>` bruto do XML célula a célula, base × nova, nos alvos:
- Eventos_Westgard: 5.000 de 5.000 iguais.
- Configuração: 200 de 200.
- Eng_Saida: 180 de 180.
- Registros: 2 de 2.
- Estatística: 5 de 5.
- Westgard_Fotos F: 4.356 de 4.356.
- LotesStore: 1 de 1.
- Sync_Log: 16.012 células comuns iguais. As 4 diferentes são a linha 6, vazia na base, que na nova é o ATUALIZAR da
  própria entrega (A6 = 15/11/2026 04:59).

**Fluxo vizinho.** O `SaveCopyAs` foi feito depois do login, das listas, dos filtros, dos manuais, da colagem e do
ATUALIZAR. Nenhum desses passos recria o ano literal em nenhuma aba.

A queda do Eventos_Westgard de 5.000 para 1.763 linhas depois do ATUALIZAR foi igual na base e na nova. A explicação
provável são as duas regras TESTE, que tiraram 8.311 resultados reais da estatística. Não investiguei além, porque não
muda o formato.

## Ocorrências novas

### R5P2-H01: DETALHE da substituição com o verbo "é" sem acento, incluindo a retransmissão (resíduo da OC-0021; extensão do R5P2-01 da Bio)

- **Prioridade:** Baixo · **Tipo:** usabilidade · **Tela/módulo:** aba Digitar Resultados, coluna DETALHE. A origem é o
  `mEtlCQ.RazaoSubstituicao` (`mEtlCQ.bas:808` e `:816`) e o mesmo texto do PQ. A tela usa
  `camada_dados.ACENTOS_DETALHE`.
- **Ação:** digitar no SUBSTITUI_ID o ID de outro manual (MAN_0002), ou o de um resultado retransmitido (HEM-212660),
  completar a linha e clicar em ATUALIZAR DADOS.
- **Esperado:**
  - "… MAN_0002 **é** resultado manual …";
  - "… HEM-212660 **é** retransmissão de HEM-212425 (informe o ID original) …".
- **Observado:**
  - "SUBSTITUI_ID inválido: MAN_0002 e resultado manual (SUBSTITUI_ID aceita só resultado do interfaceamento) (este
    manual não substitui nada)";
  - "coincide com resultado do interfaceamento (0 min); SUBSTITUI_ID inválido: HEM-212660 e retransmissão de
    HEM-212425 (informe o ID original) (este manual não substitui nada)".
  - O caso da retransmissão não tinha sido provocado na Bio, que não tem DUPLICIDADE_ORIGEM. Na Hema há 589.
- **Dificuldade:** pouca. "HEM-212660 e retransmissão de HEM-212425" parece uma lista de dois IDs, não a razão da
  recusa.
- **Consequência:** só de leitura. Status, MOTIVO e números são iguais aos da base.
- **Sugestão:** a do R5P2-01: pares com contexto no `ACENTOS_DETALHE` (`' e resultado manual'`, `' e retransmiss'`), ou
  o texto reescrito na origem nos dois motores; e a trava D01 estendida a esses padrões.
- **Evidência:** `Hematologia_F1_nova.json`, `O21_O68_manuais.manuais.M10_subst_id_manual` e `.M13_subst_retransmissao`
  (DETALHE_txt, `sem_acento`).

### R5P2-H02: Registros H4 recusa o valor inválido com a mensagem genérica do Excel (o R5P2-02 vale também para a Hema)

- **Prioridade:** Baixo · **Tipo:** usabilidade · **Tela:** Registros H4 ("Calibrado?") e F7:F9.
- **Ação:** digitar "Talvez" ou "S" no H4.
- **Esperado:** "Escolha … na seta da célula.", como nas listas da Digitar Resultados (texto do ADR-095 §1 e do
  critério de aceite da OC-0068).
- **Observado:** `Validation.ErrorMessage` vazio, na base e na nova. A recusa funciona: `Validation.Value` = Falso.
- **Consequência:** nenhuma nos dados. Só a orientação fica pior que a da Digitar Resultados.
- **Sugestão:** a do R5P2-02: ErrorMessage nos filtros literais do `instalar_adr087.FILTROS`. Pode ser unificada com o
  R5P2-02.
- **Evidência:** `Hematologia_F1_nova.json`, `O68_listas.Registros_H4.msg` e `.Registros_F7_foi_calibrado.msg`.

### R5P2-H03: linha colada com o ID junto aparece ATIVA na tela, mas o sistema a deixou fora da estatística

- **Prioridade:** Médio · **Tipo:** usabilidade (inconsistência de visualização) · **Tela/módulo:** aba Digitar
  Resultados (STATUS/RUN/DETALHE são XLOOKUP por ID_REGISTRO), `Worksheet_Change` da aba e motor (`mEtlCQ.bas:1244`).
  Não depende da entrega 42: base e nova são iguais. Testado só na Hema; o código é comum.
- **Ação:** copiar a linha inteira de um manual (MAN_0007) e colar na próxima linha livre para "aproveitar" os campos,
  depois clicar em ATUALIZAR DADOS.
- **Esperado:** uma de duas coisas:
  - a linha colada ganha um ID novo (MAN_0017) ao ser colada;
  - ou a tela mostra que ela ficou fora da estatística (CONFLITO_MANUAL, "ID_REGISTRO manual repetido").
- **Observado:**
  - A linha colada manteve o ID MAN_0007. Nenhum aviso apareceu ao colar.
  - Depois do ATUALIZAR, a tblCQ_Final tem `MAN_0007` (ATIVO, SIM, RUN 26092901) e `MAN_0007~L17` (CONFLITO_MANUAL,
    NÃO, sem RUN, "ID_REGISTRO manual repetido (vale a 1a linha)").
  - A QA_INTEGRACAO tem o A01: "MAN_0007~L17 … O manual ficou FORA da estatistica; confirme e inative um deles ou
    corrija o manual."
  - A **tela mostra as duas linhas iguais**: STATUS ATIVO, RUN 26092901 e o DETALHE do MAN_0007. O XLOOKUP pelo ID
    acha sempre a 1ª.
  - A mensagem do ATUALIZAR não cita o caso.
- **Dificuldade:** o usuário confere a aba, como o GUIA §3 manda, vê ATIVO nas duas linhas e conclui que as duas
  entraram. O único sinal é um alerta no meio de 8.914 da QA_INTEGRACAO.
- **Consequência:**
  - Os números estão certos: a cópia não entra na estatística nem no Painel.
  - Se o usuário corrigir o valor na linha colada achando que vale, a correção não entra, e a tela continua dizendo
    ATIVO.
  - "Inative um deles" é ambíguo: os dois têm o mesmo ID na tela.
- **Sugestão:**
  - carimbar ID novo quando a colagem trouxer um ID que já existe na tabela;
  - ou o STATUS/DETALHE da tela casar pela chave da linha (o `~L<n>`), e não só pelo ID;
  - e citar o caso na mensagem do ATUALIZAR, como já se faz com os incompletos.
  - Decisão do Orquestrador.
- **Evidência:** `Hematologia_F1_nova.json`:
  - `O21_O68_manuais.M19_final` (as duas linhas da tblCQ_Final);
  - `.manuais.M19_linha_colada_de_M09` (o que a tela mostra);
  - `.pre.M19_linha_colada_de_M09` (ID MAN_0007 na linha colada);
  - `.qa_integracao_manuais` (A01 `MAN_0007~L17`).
  - Mesmo resultado em `Hematologia_F1_base41.json`.

## Dúvidas de regra de negócio

1. **MOTIVO só da lista.** Com a OC-0068 resolvida, o MOTIVO aceita só os 4 itens, com estilo Parar. "Outro: falha no
   LIS" é recusado. O ADR-082 trata o motivo como justificativa de auditoria da substituição. Para "Outro", o laboratório
   quer um lugar para o texto livre, por exemplo uma coluna de observação ou o estilo Aviso? Ou bastam os 4 itens?
   (Nos meus testes, o texto "Outro: TESTE-AUDITORIA" foi gravado por automação, o que passa por cima da validação.
   Pela tela, o usuário não consegue.)
2. **OC-0072, formato interno 14.** Configuração D/AD, Registros E4/F4 e Eventos_Westgard A passaram para o formato
   interno 14 (data abreviada do Windows). Ele mostra dd/mm/aaaa em pt-BR, mas segue a configuração regional da
   máquina. É a mesma dúvida da Bio: todas as máquinas estão em pt-BR?
3. **RUN do dia no IMPLAUSIVEL e nos LOTE_*.** Os manuais fora da estatística de 29/09 às 01:00–01:06 levaram o RUN
   26092901, como na base. É a corrida em que foram lançados. A dúvida 3 do reteste 04 (um IMPLAUSIVEL abre corrida e
   renumera as do dia) continua em aberto; não medi a renumeração dos reais.

## Impressões de uso (Relatório A)

- **Listas:**
  - A seta do MOTIVO agora mostra quatro opções. A escolhida vai certa para o resultado substituído ("motivo: Resultado
    não transmitido") e para a trilha.
  - Antes, qualquer motivo digitado era recusado, e o menu gravava os quatro colados.
  - EQUIPAMENTO e MATRIZ têm um item só e já funcionavam.
  - O "Calibrado?" do Registros filtra certo.
- **DETALHE:** coerente e acentuado (implausível, estatística, gráfico, já, não, só, retransmissão). Os dois tropeços são
  "MAN_0002 e resultado manual" e "HEM-212660 e retransmissão de HEM-212425". Ler o segundo como "dois IDs" é fácil.
- **Datas:**
  - Eventos_Westgard, Configuração e De/Até do Registros mostram a data, e não mais "yyyy".
  - A coluna de data/hora do Eng_Saida deixou de ser "#####" e mostra os segundos.
- **Colar linha:** é um atalho natural para quem digita vários manuais parecidos. Hoje ele produz uma linha que parece
  ATIVA e não é (R5P2-H03). É o único ponto desta rodada em que a tela contradiz o que o sistema fez.
- **Mensagens:**
  - A do ATUALIZAR separa incompletos e SUBSTITUI_ID inválido, e as contagens batem com a QA_INTEGRACAO.
  - As das listas da Digitar Resultados orientam bem. A do H4 é a genérica do Excel.

## Arquivos

- Script: `_arquitetura/agentes/ferramentas/reteste05p2_usuario_f1_hema.py`. Usa `reteste05p2_usuario_f1.py` e
  `xml_formatos_listas.py`.
- Saídas (fora do repositório): `%TEMP%\qcw\auditoria\2026-11-15_reteste05_parte2\Hematologia_F1_{base41,nova}.{json,log}`
  e `Hematologia_F1_xml.json`. As cópias `.xlsm` foram apagadas.
- O Audit_Log e o Sync_Log da cópia gravam o login do Windows de quem roda. Esse dado foi omitido deste relatório.
