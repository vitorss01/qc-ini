# Tester/QA — ciclo 2026-11-14_reteste04 — modo REPRODUZIR

Reprodução independente das duas ocorrências altas do ciclo (P-01 Bio, P-02 Hema; origem Bio F1 R4-01 e Hema F1
R4-01 extensão). Não reaproveitei os scripts do Usuário Simulado (`reteste04_usuario_f1*.py`). Usei só `qa_final.QA`,
que abre a cópia, revela as abas e imprime o EXCEL_PID.

## 1. Ambiente e segurança

| Item | Valor |
|---|---|
| Script próprio | `_arquitetura/agentes/ferramentas/repro_listas_r4.py` (`python -u repro_listas_r4.py Bioquimica --consequencia` / `Hematologia`) |
| Produção copiada | `QC_Bioquimica.xlsm` sha256 `8ec815ab653d`, `QC_Hematologia.xlsm` sha256 `aab0a54f464e` (os mesmos antes e depois, conferido) |
| Cópias | `%TEMP%\qcw\auditoria\2026-11-14_reteste04\tester\<produto>_repro_listas_<hhmmss>.xlsm`, sempre nova a cada rodada, nunca salva e apagada no fim |
| Excel | os PIDs criados pelo script (21028, 30888, 30740, 20196, 8736) foram fechados um de cada vez pelo `qa.fechar(False)`. Os outros EXCEL.EXE vistos (18772, 31600) não eram meus e ficaram intactos |
| Idioma do Excel | `International(xlListSeparator)` = `;`, decimal `,`, UI 1046 (pt-BR) |
| Dados sintéticos (só na cópia) | MAN_0001 = 123,4567 e MAN_0002 = 123,4568, Glicose N1, lote 8977, 14/11/2026 03:33/03:34. Usuário "(sem login)" |
| Captura de tela | nenhuma |
| Logs/JSON | `tester_repro_listas_Bioquimica.log/.json` (rodada final), `_run1`/`_run2` (rodadas anteriores, ver §4), `tester_repro_listas_Hematologia.log/.json` |

O método: a validação foi avaliada pelo **próprio Excel** com `Validation.Value`. O script escreve o valor numa linha vazia da
`tblResultados_Manuais` com os eventos desligados, lê a validação e limpa a célula. O teste foi repetido na 1ª linha
livre (linha 7) e na última da folga (linha 306). A **digitação física** pelo teclado ficou `NAO_EXECUTADO` porque não houve
interação com a tela. Além disso, o script varreu **todas** as validações de lista de **todas** as abas
(`SpecialCells(xlCellTypeAllValidation)`, com as abas desprotegidas só na cópia).

## 2. P-01 — Bioquímica — **REPRODUZIDA**

### 2.1 Listas (`tester_repro_listas_Bioquimica.json`, `validacao_por_coluna`)

| Coluna (intervalo) | Formula1 | Estilo | Validation.Value |
|---|---|---|---|
| EQUIPAMENTO (D7:D306) | `DIMENSION 1,DIMENSION 2` | 1 = Parar, "Escolha o equipamento da lista." | "DIMENSION 1" **Falso** · "DIMENSION 2" **Falso** · texto inteiro Verdadeiro |
| MATRIZ (E7:E306) | `SORO,URINA` | Parar, "Escolha a matriz da lista." | "SORO" **Falso** · "URINA" **Falso** · texto inteiro Verdadeiro |
| MOTIVO (K7:K306) | `Falha do interfaceamento,Resultado não transmitido,Equipamento em contingência,Outro` | Parar, "Escolha o motivo da lista." | cada um dos 4 motivos **Falso** · texto inteiro Verdadeiro |
| LOTE (F7:F306) | `=lstLotes` | Parar | "8974" Verdadeiro · "LOTE-INEXISTENTE-TESTE" Falso (correto) |
| ANALITO (H7:H306) | `=lstAnalitos` | Parar | "Lactato" Verdadeiro · "ANALITO-INEXISTENTE-TESTE" Falso (correto) |

Na linha 306 o resultado foi o mesmo: "DIMENSION 1", "SORO" e "Falha do interfaceamento" deram Falso. O defeito não
depende da linha.

### 2.2 Consequência (rodada final, cópia `Bioquimica_repro_listas_170843.xlsm`)

O Painel ficou em Glicose (C3), DIMENSION 1 e lote 8977. Digitei dois manuais célula a célula, com os eventos ligados,
como um usuário faria:

- **A (itens do menu):** EQUIPAMENTO `DIMENSION 1,DIMENSION 2`, MATRIZ `SORO,URINA`, MOTIVO = os 4 motivos colados,
  RESULTADO 123,4567. O ID gerado foi **MAN_0001**.
- **B (controle):** EQUIPAMENTO `DIMENSION 1`, MATRIZ `SORO`, MOTIVO `Outro`, RESULTADO 123,4568. O ID gerado foi
  **MAN_0002**. Pela tela esses valores seriam recusados (Validation.Value Falso nos três).

Rodei o ATUALIZAR (`mIntegracao.AtualizarDadosAutomatico`, 20,3 s; 0 resultados novos do SEAC).

| | MAN_0001 (menu) | MAN_0002 (controle) |
|---|---|---|
| Digitar Resultados: STATUS / DETALHE / RUN | **ATIVO** / vazio / 26111401 | ATIVO / vazio / 26111401 |
| tblCQ_Final: EQUIPAMENTO | **`DIMENSION 1,DIMENSION 2`** | `DIMENSION 1` |
| tblCQ_Final: MATRIZ / **BLOCO** / ANALITO_ORIGEM | `SORO,URINA` / **vazio** / `Glicose` | `SORO` / `PAINEL` / `GLUC` |
| STATUS_ANALITICO / PARTICIPA_ESTATISTICA / GOVERNANCA | **ATIVO / SIM / OK** | ATIVO / SIM / OK |
| MOTIVO_MANUAL | os 4 motivos colados | `Outro` |
| tblQA_Integracao (148 linhas) | **nenhuma linha cita MAN_0001** | nenhuma (esperado) |
| Série DIMENSION 1 do Painel (Eng_Saida/Calc, K4 = "Tudo (sem filtro)") | **ausente** | **presente**: Eng_Saida!Y128, RUN 26111401; Calc!F128/H128/J128 |
| Série DIMENSION 2 do Painel | **ausente** | ausente (esperado) |

Todos os pontos da ocorrência se confirmam: o manual entra ATIVO e PARTICIPA SIM com um equipamento que não existe, o
DETALHE fica vazio, o QA não dá nenhum achado e o resultado não aparece em nenhuma das duas séries. O controle mostra
que o método enxerga um manual quando ele está certo. A única diferença é o RUN: aqui 26111401 (data de 14/11), lá
26101401 (data do Usuário Simulado). O mecanismo é o mesmo.

### 2.3 Ampliação

1. **A MATRIZ do menu também quebra o de/para.** É uma consequência nova, que a ocorrência não mencionava. Com MATRIZ
   `SORO,URINA`, o MAN_0001 sai com **BLOCO vazio** e ANALITO_ORIGEM `Glicose`, sem a tradução do de/para. O controle sai
   com BLOCO `PAINEL` e `GLUC`. O menu da MATRIZ grava uma matriz que não existe e o resultado perde o bloco. **Não
   isolei** se a ausência no LJ vem do EQUIPAMENTO (guarda `engEquip<>selEquipamento` do Calc), do BLOCO vazio ou dos
   dois, porque o manual A tinha os dois defeitos ao mesmo tempo.
2. **Nenhuma regra do QA pega equipamento fora do cadastro na Bio.** O A10 de `src/comum/mEtlQA.bas` (linhas
   656–676) só roda quando `SETOR = "HEMATOLOGIA"`. Na Bio, um EQUIPAMENTO que não existe passa calado, e isso explica
   o "nenhum achado no QA". Uma regra nova depende do parecer do Especialista em CQ.
3. **Outra aba, os dois produtos:** a varredura achou mais uma lista com um item só, o filtro **"Calibrado?" da aba
   Registros (H4)**, `Formula1 = "Sim,Não"`. A causa é `instalar_adr087.py:42-43` (`FILTROS`, `'Sim,Não'`). É só um
   filtro de vista, então o impacto é menor: o filtro não pode ser escolhido nem digitado (minha estimativa é Médio, a
   classificar). A Hema tem o mesmo.
4. **Não afetadas:** Painel L4 `DIMENSION 1;DIMENSION 2` (aceita o valor atual), Painel K4/M4, Especificações Q/S/AR,
   Estatística F4, Registros D7:D9/F7:F9/J7:J9 e Usuarios C4:C53/H8 usam `;`. As demais validações apontam para
   nomes ou intervalos.
5. **O modelo de correção já existe no repositório.** `instalar_adr064.py:151` e `instalar_adr075.py:173` montam a
   lista com `Application.International[5 - 1]` (xlListSeparator). Do lado do VBA, `GuardarValidacao`
   (`src/bio/mIntegracao.bas:885-900`) já separa a Formula1 por `,` **e** `;`. Uma lista com `;`, ou que aponte para um
   nome, não quebra o cadastro que o VBA monta a partir dela.

### 2.4 Causa (arquivo e linha; não editei nada)

`_arquitetura/integracao_dbseac/camada_dados.py`, função `validacoes`:
- linha 532: `lista('EQUIPAMENTO', ','.join(p['equipamentos']), ...)`;
- linha 534: `lista('MATRIZ', ','.join(matrizes), ...)`;
- linha 537: `lista('MOTIVO', ','.join(MOTIVOS_MANUAL), ...)`;
- a linha 515 (`r.Validation.Add(3, estilo, 1, src)`) grava o texto como está. Por COM, o Excel pt-BR lê esse texto
  com o separador local `;` e a lista fica com um item só.
- Mesmo padrão em `camada_dados.py:503` (fallback `'SIM,NÃO'` do REGISTRAR - LJ). Ele não está ativo nas produções
  porque a varredura não achou essa lista (o checkbox nativo foi usado), mas é latente.
- Mesmo padrão em `instalar_adr087.py:43` (`'Sim,Não'`, Registros!H4).

## 3. P-02 — Hematologia — **REPRODUZIDA**

`tester_repro_listas_Hematologia.json`, cópia `Hematologia_repro_listas_170422.xlsm`:

| Coluna | Formula1 | Validation.Value |
|---|---|---|
| MOTIVO (K7:K306) | `Falha do interfaceamento,Resultado não transmitido,Equipamento em contingência,Outro`, estilo Parar, "Escolha o motivo da lista." | "Falha do interfaceamento", "Resultado não transmitido", "Equipamento em contingência" e "Outro": **Falso**. Texto inteiro: Verdadeiro. Linha 306: igual |
| EQUIPAMENTO (D7:D306) | `XN1000` | "XN1000" Verdadeiro (certo, um item) |
| MATRIZ (E7:E306) | `SANGUE TOTAL` | Verdadeiro (certo, um item) |
| LOTE / ANALITO | `=lstLotes` / `=lstAnalitos` | "522611" e "WBC" Verdadeiro; inventados Falso (certo) |

**Ampliação na Hema:**
- O filtro "Calibrado?" da aba Registros (H4) também tem `Sim,Não`, com um item só (ver §2.3, item 3).
- As demais listas literais usam `;`: Registros D/F/J, Especificações Q/S, Estatística R4 e Usuarios.
- Não reproduzi na Hema a gravação de um manual com o MOTIVO do menu (substituição, Audit_Log, I04): `NAO_EXECUTADO`.
  Na Bio, porém, o MAN_0001 mostra que o texto colado chega inteiro à `tblCQ_Final.MOTIVO_MANUAL`. O código que grava o
  motivo é o mesmo (`src/comum`), então o efeito esperado na Hema é o mesmo.

## 4. Rodadas anteriores (transparência)

- `_run1`: o analito foi posto por engano em **B3**, o índice do Painel, e não em C3, o nome. Com isso o controle não
  apareceu na série e a verificação do LJ ficou inválida. As tabelas (STATUS, tblCQ_Final, QA) deram o mesmo resultado
  da rodada final.
- `_run2`: com C3 certo, mas com o Painel no filtro de período da produção (K4 "Out", G3:G4 01–31/10/2026), o motor
  parou em 30/10. Nem o controle de 14/11 apareceu, então o resultado foi inconclusivo.
- A rodada final pôs K4 em "Tudo (sem filtro)" na cópia. O controle apareceu (Eng_Saida!Y128) e o manual do menu não.
  É essa rodada que vale.
- Tempo do ATUALIZAR: 20,9 s, 20,6 s e 20,3 s, sem anomalia.

## 5. Testes automatizados propostos

| Suíte | ID proposto | Caso |
|---|---|---|
| `qa_manual_substitui.py` | **M16** | Nos dois produtos, para cada coluna de lista da `tblResultados_Manuais` (EQUIPAMENTO = `cp.PRODUTOS[...]['equipamentos']`, MATRIZ = matrizes do de/para, MOTIVO = `MOTIVOS_MANUAL`): `Validation.Value` Verdadeiro para **cada** item, Falso para o texto inteiro colado e Falso para um valor inventado. Linha 7 e última linha da folga |
| `qa_final.py` | **T27** | Varredura de todas as abas: nenhuma validação de lista literal (Formula1 sem `=`) com `,` quando `International(xlListSeparator) <> ","`. Pega o Registros!H4 e qualquer lista futura |
| `qa_adr087.py` | **R08** | O filtro "Calibrado?" (Registros!H4) aceita "Sim" e "Não" e filtra |
| `qa_etl.py` | **E17** (depende do parecer do CQ) | Manual da Bio com EQUIPAMENTO ou MATRIZ fora do cadastro gera um achado no QA, ou fica fora da estatística com DETALHE. Hoje sai ATIVO, SIM, GOVERNANCA OK, BLOCO vazio e sem achado |

O caso M16 entra no critério de aceite da correção. O T27 é a regressão estrutural. O E17 é uma regra nova de QA e
precisa da decisão do Especialista em CQ e da autorização do usuário.
