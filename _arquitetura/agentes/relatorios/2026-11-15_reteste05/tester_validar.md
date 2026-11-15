# Tester/QA, modo VALIDAR: ciclo 2026-11-15_reteste05 (com a parte 2)

Agente: Tester/QA · Data: 15/11/2026 · Correção validada: ADR-095 (commits 69a1f65, 39084f1 e e7490ba; a Hema foi
entregue a partir de 51bffc2, que só muda o instalador de VBA). Quem corrigiu foi o Orquestrador. **Eu não corrigi
nada** e não editei `src/`, instaladores, fórmulas nem a produção.

| Item | Valor |
|---|---|
| Produções validadas (entrega 42) | `QC_Bioquimica.xlsm` sha256 **b20fe00f7e51**; `QC_Hematologia.xlsm` **7bab8fc008f6**. O hash é o mesmo antes e depois de todas as rodadas |
| Bases (entrega 41) | Bio `_backup_pre_integracao_2026-11-15_0011` **8ec815ab653d**; Hema `_backup_pre_integracao_2026-11-15_0458` **aab0a54f464e** |
| Cópias | Todas novas, em `%TEMP%\qcw\auditoria\2026-11-15_reteste05\tester\`. Nenhuma foi salva sobre si mesma. O único `SaveCopyAs` gravou na mesma pasta e serviu só para ler o XML do estado em uso. As 34 cópias `.xlsm` foram apagadas no fim; ficam os JSON e os logs |
| tblCQ_Final | Bio 35.906 linhas, hash do Value2 `471c3b76f80f008c`; Hema 48.391 linhas, `ee71bc38576d73a3`. Iguais na base e na nova |
| Excel | Um por vez, aberto e encerrado pelo `qa_final.QA`. Minhas rodadas: PIDs 16100 (refeita depois de corrigir o meu script), 14584, 30416, 12432, 29964, 9212 e 20500. As suítes usaram 16 PIDs, todos encerrados por elas. O Excel 13084 da tarefa agendada e o 31296 (que não era meu) não foram tocados. Nenhuma captura de tela |
| Dados de teste (só nas cópias) | Usuário `TESTE_AUDITORIA_R5V` (ANALISTA), que fica carimbado no USUARIO dos manuais. Manuais de 14/11/2026 com valores sintéticos: Bio 123,456 / 98,765 / 111,111; Hema 12,345 / 11,111. Na cópia `lote_sint` da Hema, 388 linhas de WBC dos lotes 533811 e 619611 receberam o rótulo LOTE 528211 |
| Scripts (meus) | Em `_arquitetura/agentes/ferramentas/`: `valida_r5_excel.py` (Excel: telas, listas, Painel, manuais e colagem), `valida_r5_xml.py` (XML direto, sem Excel), `valida_r5_compara.py` (base × nova), `valida_r5_amao.py` (n e média contados à mão na tblCQ_Final do SaveCopyAs) e `valida_r5_lote_sint.py` (ramo "do lote" do aviso). Não reaproveitei os scripts do Usuário Simulado nem as suítes qa_adr09x |

## Vereditos

| Ocorrência | Produto | Veredito |
|---|---|---|
| OC-0012 (aviso do corte de 180 corridas) | Hema (Bio = regressão) | **VALIDADA** |
| OC-0068 (listas com um item só) | Ambos | **VALIDADA** |
| OC-0071 (cabeçalho da Estatística em "Tudo") | Bio | **VALIDADA** |
| OC-0072 (ano literal "yyyy") | Ambos | **VALIDADA** |

**Efeito colateral da correção:** nenhum. Em todos os cenários, os números da entrega 42 são idênticos aos da entrega
41: tolerância zero, porque é o mesmo motor sobre a mesma tblCQ_Final. Achei 2 observações. As duas já existiam na
entrega 41, então não são efeito colateral:
- **R5V-01:** o aviso da Bio com o H3 vazio manda o usuário à Estatística, que mostra outro recorte;
- **R5V-02:** existe um bloco de estatística por ano que nenhuma célula lê.

Ver §6.

---

## 1. OC-0012: aviso do corte na Hematologia (VALIDADA)

Oráculo próprio (Python puro, sobre o Value2 da tblCQ_Final da cópia):
- **o que conta:** PARTICIPA_ESTATISTICA = SIM, RESULTADO numérico, analito, equipamento do Eng_Saida!S1 e lote (com o
  H3 vazio, todos os lotes);
- **corrida:** o RUN, ou o par (RUN, LOTE) com o H3 vazio (ADR-080/086). A data da corrida é a da 1ª linha;
- **ordem:** as corridas vão por (data, RUN);
- **janela:** as 180 corridas elegíveis mais recentes até o Até;
- **M1:** cortadas do período mais as corridas só com X sem posição (D10 do QA-ETL-001);
- **AF1:** o n do período, por nível.

| Cenário (Hema) | M1 (sistema / oráculo) | AF1 = oráculo | Aviso no A39 (nova) | Igual à base |
|---|---|---|---|---|
| H1 WBC, H3 vazio, 2026 "Ano inteiro" | 278 / 278 | N1 374 · N2 400 · N3 340 | "Período inteiro (todos os lotes): N1 374 · N2 400 · N3 340 resultados." | engPainel, A6:Q12, Calc A1:AZ190 e Eng_Saida A1:AF190 (sem o G1) iguais; só o texto do A39 muda |
| H2 WBC, H3 vazio, "Tudo" | 431 / 431 | N1 515 · N2 545 · N3 479 | idem, "(todos os lotes)" | idem |
| H3 WBC, H3 vazio, De/Até 01/01–31/10/2026 | 251 / 251 | N1 357 · N2 377 · N3 328 | idem | idem |
| H4 PLT, H3 vazio, "Tudo" (ampliação: outro analito) | 431 / 431 | N1 515 · N2 545 · N3 479 | idem | idem |
| H5 RBC, H3 vazio, 2025 (sem corte) | 0 / 0 | vazio | sem aviso | tudo igual, A39 inclusive |
| H6 WBC, lote 528211, "Tudo" (89 corridas, sem corte) | 0 / 0 | vazio | sem aviso | tudo igual |
| H7 WBC, H3 vazio, Out/2026 (sem corte) | 0 / 0 | vazio | sem aviso | tudo igual |

- **n da janela por nível** (engPainel col. 2 × oráculo): H1, H2 e H4 120/158/112; H3 125/157/116, sem diferença.
- **Sem corte** (H5–H7), o n do Painel = o n do período do oráculo: 141/145/139, 83/86/81 e 21/39/21.
- **Média:** com lote (H6), a diferença relativa da média é ≤ 2,2e-16. No modo TODOS, o Painel não mostra a média (ADR-086).
- **Aba Estatística:** "Estatística" não aparece em nenhum A39 da Hema nova. Na base, aparecia nos 4 cenários com corte.

**Ramo "do lote":** nenhum lote real tem mais de 180 corridas (o maior, 528211, tem 89). Por isso o ramo foi
exercitado com dado sintético, em `valida_r5_lote_sint.py`, com lotes diferentes dos do Usuário Simulado:
- 388 linhas de WBC dos lotes 533811 (81 corridas) e 619611 (77) receberam o rótulo 528211;
- o resultado: M1 64 = oráculo 64; AF1 "N1 213 · N2 221 · N3 196" = oráculo;
- o A39 diz "⚠ 64 corrida(s) antigas … Período inteiro do lote: N1 213 · N2 221 · N3 196 resultados.", sem
  "Estatística";
- em 2026 (150 corridas) não há aviso. Com o H3 vazio, o texto volta a "(todos os lotes)" com 278 e 374/400/340;
- o n da janela fica 3 abaixo do oráculo por nível (152 × 155). É o mesmo artefato que o Usuário Simulado viu: os lotes
  reunidos rodaram nos mesmos RUN, e o motor guarda um ponto por nível em cada corrida. Com dado real, o n bateu
  exatamente.

**Bio (regressão do aviso):** em 6 cenários com corte, o A39 da nova = o da base, com "— veja a aba Estatística":
- B1 Glicose 8974 "Tudo": M1 39, AF1 N1 174 · N2 161;
- B4 H3 vazio "Tudo": M1 164 = 163 cortadas + 1 corrida só com X; AF1 264/249;
- B5 H3 vazio 2026: M1 101, AF1 211/202;
- B8 só o Até 31/05: M1 8, AF1 153/140;
- B9 DIMENSION 2, Hemoglobina glicada, H3 vazio: M1 212, AF1 303/343;
- B10 volta a "Tudo".

M1 e AF1 = oráculo em todos. A fórmula do A39 da Bio na nova é igual à da base (XML).

## 2. OC-0068: listas de escolha (VALIDADA nos dois produtos)

Método:
- `Validation.Formula1` foi partido por `Application.International(xlListSeparator)` (";" nas duas cópias);
- cada item foi escrito na célula com os eventos desligados e conferido por `Validation.Value`;
- 9 textos inválidos foram testados: as listas coladas por vírgula e por ";", "DIMENSION 3", "XN2000", "PLASMA",
  "Talvez" e "Outro: TESTE-AUDITORIA".

| Lista | Base (entrega 41) | Nova (entrega 42) |
|---|---|---|
| Bio Digitar Resultados EQUIPAMENTO | 1 item "DIMENSION 1,DIMENSION 2" | 2 itens, os 2 aceitos, os 9 inválidos recusados. Mensagem "Escolha o equipamento na seta da célula." |
| Bio MATRIZ | 1 item "SORO,URINA" | 2 itens (SORO, URINA), aceitos; inválidos recusados. "Escolha a matriz na seta da célula." |
| Bio e Hema MOTIVO | 1 item com os 4 motivos | 4 itens, aceitos; inválidos recusados. "Escolha o motivo na seta da célula. Ele é obrigatório quando o manual substitui um resultado." |
| Bio e Hema Registros H4 | 1 item "Sim,Não" | 2 itens (Sim, Não), aceitos; "Sim,Não" recusado |
| Bio Painel L4 | já era de 2 itens | 2 itens, aceitos; inválidos recusados |
| Hema EQUIPAMENTO / MATRIZ | 1 item (XN1000 / SANGUE TOTAL; um valor só no cadastro) | igual, com a mensagem "na seta da célula" |

- **XML:** a base tem 4 `<x12ac:list>` (Bio) e 2 (Hema), por exemplo Registros H4 `"Sim,Não"` e a coluna D da Digitar
  Resultados. A nova tem 0, inclusive no SaveCopyAs depois do login e depois do ATUALIZAR, nos dois produtos.
- **Filtros só mudam a vista.** Com o H4 em Sim, Não e vazio, as linhas visíveis da tabela do Registros foram Bio 3/1/1/3
  e Hema 3/2/2/3. O hash do Value2 da tabela e da tblCQ_Final ficou igual. Com o E4/F4 digitados (01/09 e 30/09/2026),
  o hash também ficou igual.
- **Manual com equipamento do menu (Bio).** MAN_0001: EQUIPAMENTO = 2º item do menu ("DIMENSION 2"), MATRIZ e MOTIVO do
  menu, Glicose, lote 8977, N1, 123,456, 14/11/2026 09:07:05. Depois do ATUALIZAR:
  - tblCQ_Final: ATIVO, SIM, EQUIPAMENTO DIMENSION 2, RUN 26111401;
  - Painel DIMENSION 2 / Glicose / 8977 / Nov-2026, N1: n passou de 11 (média 75,818181…) para 12 (média 79,788);
  - à mão, na tblCQ_Final do SaveCopyAs: n = 12, média = 79,788 (sem os manuais, 11 e 75,8181818181818). Na conta,
    (11 × 75,8181818… + 123,456) / 12 = 79,788, com diferença relativa 0;
  - a série do DIMENSION 1 não recebeu o MAN_0001.

## 3. OC-0071: Estatística da Bio com o Painel em "Tudo" (VALIDADA)

Foram 12 cenários (B1–B10, B6b e B6c), lidos na base e na nova.

| Cenário | B5 | D5 | D4 | H5 / I5 / J5 |
|---|---|---|---|---|
| "Tudo" (B1, B4, B9, B10), nova | "" (Value2 0) | "" (Value2 0) | "" | "histórico inteiro do lote" / "" / "" |
| "Tudo", base | "00/01/1900" | "00/01/1900" | "2026" | idem |
| Set/2026 (B2), nova | 01/09/2026 | 30/09/2026 | 2026 | 01/09/2026 a 30/09/2026 |
| Ano inteiro (B3, B5) | 01/01/2026 | 31/12/2026 | 2026 | 01/01/2026 a 31/12/2026 |
| Maio De/Até (B6b, B7) | 01/05/2026 | 31/05/2026 | 2026 | 01/05/2026 a 31/05/2026 |

- **Fórmulas:** o B5/D5 continua `=Painel!G3` / `=Painel!G4`. Na nova, o formato é `dd/mm/aaaa;;` (na base,
  `dd/mm/aaaa`). O D4 tem a condição de De/Até vazios (ADR-095 §3).
- **Números:** a grade A14:AU133 (Value2) não tem nenhuma célula diferente entre a base e a nova nos 12 cenários. O
  engPainel, o Painel A6:Q12, o Calc e o Eng_Saida também são iguais.
- **Textos:** nenhum "yyyy" nas linhas 1–12 da Estatística, e nenhum "1900" nos cenários "Tudo" e com De/Até
  completos. O "1900" que sobra é o R5-01 (§A).

## 4. OC-0072: ano literal nos formatos de data (VALIDADA nos dois produtos)

Quatro checagens, cada uma com controle positivo na base. Todas foram feitas depois do login do
`TESTE_AUDITORIA_R5V`.

1. **XML** (`valida_r5_xml.py`).
   - Base: Bio 3 numFmt literais (`dd/mm/\y\y\y\y`, com hh:mm e com hh:mm:ss) em 453 células (Configuração 200,
     Eng_Saida 180, Sync_Log 64, Estatística 5, Registros 2, LotesStore 1, Westgard_Fotos 1 e a coluna F). Hema 2
     numFmt em 5.384 células (Eventos_Westgard 5.000, Configuração 200, Eng_Saida 180, …).
   - Nova: **0** numFmt literais e **0** células, nos arquivos de produção, no SaveCopyAs depois do login e no SaveCopyAs
     depois do ATUALIZAR.
2. **Valores inalterados.** O `<v>` e o tipo das 453 (Bio) e das 5.384 (Hema) células que tinham formato literal na
   base foram comparados com a nova: **0 diferenças**, e nenhuma célula faltando.
3. **Text na tela.** Alvos lidos na nova:
   - Configuração D26 "30/06/2026" (Bio) e "31/03/2026" (Hema);
   - Eng_Saida AC3 "07/07/2026 08:37:54" (a base mostrava "##############");
   - Eventos_Westgard A4 "01/10/2025" (a base mostrava "01/10/yyyy");
   - Registros E4/F4 digitados "01/09/2026" / "30/09/2026" (a base mostrava "01/09/yyyy");
   - Sync_Log "15/11/2026 06:55:06".
   - Formatos `dd/mm/aaaa` e `dd/mm/aaaa hh:mm:ss`.
4. **Busca "yyyy"** (`Find` em valores e fórmulas) em todas as abas visíveis. Na nova, 0 achados, antes e depois do
   ATUALIZAR. Na base, Bio Configuração D26 e Hema Eventos_Westgard A4…, o que mostra que a busca pega o defeito.

## 5. Regressão

As suítes rodaram como o `entregar.py` as chama (`rodar_suites.py`: cada suíte numa cópia nova da produção). A versão
oficial das suítes é a do commit HEAD 29d27db, tirada por `git archive` para uma pasta limpa. Nesses 5 arquivos, ela é
igual à que foi entregue.

| Suíte | Bio | Tempo | Hema | Tempo |
|---|---|---|---|---|
| qa_adr095 | **8 PASS / 0 FAIL** | 26 s | **7 PASS / 0 FAIL** | 21 s |
| qa_adr092 | 7 / 0 | 58 s | 7 / 0 | 56 s |
| qa_adr090 | 21 / 0 | 251 s | 23 / 0 | 254 s |
| qa_adr087 | 14 / 0 | 206 s | 14 / 0 | 199 s |
| qa_seguranca | 20 / 0 (estrutura travada, só Login visível) | 124 s | 20 / 0 | 121 s |

Os tempos são os mesmos da 1ª rodada (262/214/123 s), sem lentidão.

**Aviso ao Orquestrador.** Rodei primeiro com o `qa_adr095.py` da cópia de trabalho, que não está commitado. Nessa
rodada, ele **falha** contra a produção:
- Bio: D01 (os pares " é resultado manual", " é retransmiss" e " 1ª linha" fora da fórmula);
- Hema: L02 (mensagem "na seta da célula" no H4/L4), D01 e P01 (ID novo na colagem).

São os testes do P-01, do P-02 e do P-03, para a próxima entrega, e a produção ainda não os atende. **Não é
regressão.** Mas esse arquivo não pode ser usado para validar a entrega 42.

## 6. Conferências pedidas (só relato)

### A. R5-01 e a janela da Estatística da Bio

**R5-01: REPRODUZIDA** na nova (e na base).
- Só o De (B6, De 01/09/2026, e B6c, De 01/05/2026): J5.Text = "00/01/1900", Value2 0. O B5 mostra a data, e o D5
  agora fica em branco.
- **Ampliação, caso simétrico (só o Até, B8 31/05/2026):** H5.Text = "00/01/1900", Value2 0, e J5 = "31/05/2026". É o
  mesmo defeito na célula vizinha. Também existe na base.

**Qual bloco alimenta a grade.** A grade da Estatística da Bio (C:F, linhas 14–133) é a UDF
`EstatPeriodo($A,$B,"N"|"MEDIA"|…, Estat_Ini_Efetiva, Estat_Fim_Efetiva, Estat_Exclusoes, $H$4)`:
- `Estat_Ini_Efetiva` = Estatística!H5 e `Estat_Fim_Efetiva` = J5;
- `mEstatPeriodo.Agregar` filtra **por dia**, e 0 = sem limite (`src/bio/mEstatPeriodo.bas` 166–167);
- o lote é o H4 (= `loteAnalise`).

O bloco do `AtualizarEstatisticaAba`, que filtra por **ano** (`src/bio/mEstatistica.bas` 2132–2140, anoDe/anoAte de
B5/D5), grava em Eng_Saida!C190:M269 (`engEstat`, linha 2183). **Nenhuma fórmula lê esse bloco**: 0 referências a
`engEstat` nas abas, contra 320 chamadas `EstatPeriodo`, e nenhum leitor no VBA. A janela usada no cálculo é a mesma
mostrada em H5:J5.

| Glicose, DIMENSION 1 (n N1 / N2) | Grade da Estatística | Oráculo pela janela de datas | Oráculo pelo ano (2026) | Painel |
|---|---|---|---|---|
| Lote 8977, Set/2026 (De 01/09, Até 30/09) | **23 / 24** | 23 / 24 | 90 / 88 | 23 / 24 |
| Lote 8977, só o De 01/09/2026 | **58 / 56** (01/09 → último dado, 13/11) | 58 / 56 | 90 / 88 | 58 / 56 |
| Lote 8977, Out/2026 · só o De 01/10 · só o Até 31/10 | 25/22 · 35/32 · 80/78 | iguais | 90 / 88 | iguais |
| Lote 8974, Set/2026 e só o De 01/09 | **0 / 0** (o lote não tem dado a partir de setembro) | 0 / 0 | 121 / 114 | 0 / 0 |
| Lote 8974, Mai/2026 · só o De 01/05 · só o Até 31/05 | 21/24 · 42/45 · 153/140 | iguais (média igual em precisão total, ex. 86,857142857…) | 121 / 114 | 21/24 · 42/45 · 145/134 (janela de 180; AF1 153/140) |

Resposta à dúvida do UX:
- com De 01/09 e Até 30/09, o n é o de **setembro**, e não o de 2026;
- com só o De, a janela vai do De até o último dado, no Painel e na Estatística;
- logo, o texto "último dado" proposto para o J5 (e "primeiro dado" para o H5) seria verdadeiro;
- única diferença: com só o Até e mais de 180 corridas, o Painel mostra a janela de 180 (145/134) e a Estatística o
  período inteiro (153/140). É o que o aviso do corte já explica.

**R5V-02 (informativo, para o Orquestrador e o Especialista em CQ).** O `AtualizarEstatisticaAba` da Bio calcula e grava
um bloco por ano (Eng_Saida 190–269) que não chega à tela. Não muda número exibido nenhum. Se o bloco for lido por
alguém no futuro, a janela dele (ano) é diferente da janela do H5:J5. Recomendo documentar que o bloco é legado ou
removê-lo, mas só com o parecer do Especialista. Não alterei nada.

**R5V-01 (observação, Médio sugerido; já existia na entrega 41).** Na Bio, com o H3 vazio (todos os lotes), o aviso do
corte termina em "Período inteiro: N1 264 · N2 249 — veja a aba Estatística". A Estatística, porém, segue o
`loteAnalise` (H4 = 8974) e mostra 174/161.
- Os outros casos: B5 (2026) aviso 211/202 × Estatística 121/114; B9 (DIMENSION 2, Hemoglobina glicada) aviso 303/343
  × Estatística 4/4.
- É a mesma contradição que o ADR-095 §4 corrigiu na Hema. A premissa "a Estatística da Bio segue o Painel" não vale
  com o H3 vazio.
- Fonte: `instalar_adr078.corte_periodo(nlv=2)`, que mantém o texto fixo da Bio, e a célula Estatística!H4
  (`=loteAnalise`).
- Proposta de teste: `qa_adr095` A02 (Bio, H3 vazio e corte > 0): o A39 não aponta a Estatística, ou a Estatística
  mostra o mesmo recorte.

### B. P-03: linha manual copiada e colada com o mesmo ID (REPRODUZIDA nos dois produtos)

Caminho seguido:
- botão `mIntegracao.IrDigitarResultados`, com `Application.EnableEvents = True` conferido antes e depois;
- manuais digitados célula a célula com `FormulaLocal`;
- colagem por `ListRow.Range.Copy(destino)` na próxima linha livre, o equivalente ao Ctrl+C / Ctrl+V da linha;
- `Worksheet_Change` disparado; nenhuma caixa apareceu, conforme o vigia do meu PID;
- ATUALIZAR pelo núcleo do botão.

O clique e o teclado físicos foram **NAO_EXECUTADO**, porque não há sessão interativa.

| | Bio (Glicose, DIMENSION 1, lote 8977, 14/11/2026 10:20:30) | Hema (WBC, XN1000, lote 625311, 14/11/2026 10:20:30) |
|---|---|---|
| Origem (digitada) | MAN_0002, N1, 98,765 → ATIVO, RUN 26111401 | MAN_0001, N1, 12,345 → ATIVO, RUN 26111401 |
| Colada sem editar | ID **MAN_0002** (o mesmo). Tela: ATIVO, RUN 26111401, DETALHE vazio. tblCQ_Final: `MAN_0002~L3` **CONFLITO_MANUAL**, NÃO, NAO_PLOTAR, "ID_REGISTRO manual repetido (vale a 1a linha)" | ID **MAN_0001**. Tela: ATIVO, RUN 26111401. tblCQ_Final: `MAN_0001~L2` CONFLITO_MANUAL |
| Colada e editada (NIVEL 2, RESULTADO trocado) | ID MAN_0002. Tela: **ATIVO**, RUN 26111401, NIVEL 2, 111,111. tblCQ_Final: `MAN_0002~L4` **CONFLITO_MANUAL**, PARTICIPA NÃO | ID MAN_0001. Tela: **ATIVO**, NIVEL 2, 11,111. tblCQ_Final: `MAN_0001~L3` CONFLITO_MANUAL |
| O editado entra na estatística? | **Não.** Painel Nov/2026, N2: n 10, média 271,8 antes e depois. À mão: n 10, média 271,8; com o editado seriam n 11, média 257,19 | **Não.** N2: n 23, média 7,3239130… antes e depois (à mão igual; diferença relativa 1e-15) |
| Série do N1 | sobe só com a origem: n 10 → 11; média 77,5 → 79,4331818… = (10 × 77,5 + 98,765) / 11; à mão igual | n 17 → 18; média 3,1347059 → 3,6463889 = (17 × 3,1347059 + 12,345) / 18; à mão igual |
| Marca `qcUltimoIdManual` | =2 (a colagem não consumiu ID) | =1 |
| Sinal para o usuário | Só 2 ALERTAS A01 na QA_INTEGRACAO ("… confirme e inative um deles …", que manda inativar o resultado bom). A mensagem do ATUALIZAR não cita o caso. Ela diz "4 erro(s) de governança": são 4 E01 de inativações reais da produção sem motivo (Lactato N2, 13/10/2026), sem relação com o teste | 2 A01; a mensagem do ATUALIZAR não cita o caso (0 erro) |

**Conclusão B.** O P-03 se confirma, e o cenário colar-e-editar, que era inferência do UX, agora foi executado:
- o resultado editado (outro nível, outro valor) fica **fora** da estatística, do LJ e do Westgard;
- a tela mostra ATIVO, com o RUN e o DETALHE da linha original;
- causa (sem edição minha): `PrepararLinha` (`src/bio|hema/mIntegracao.bas` 768–775) só carimba ID novo quando o ID está
  vazio, e a colagem traz o ID junto. O motor (`src/comum/mEtlCQ.bas` 1243–1244) testa "ID repetido" antes de
  qualquer outra coisa e não olha o conteúdo.

Concordo com a prioridade **Alta**. Teste automatizado proposto: `qa_adr095` P01, que já existe na cópia de trabalho.
Ele deve incluir:
- o caso colar-e-editar: ID novo, ATIVO, n do N2 + 1, média conferida;
- o caso colar sem editar: CONFLITO pelo motivo "resultado manual repetido (mesma chave e instante)";
- o mover, sem ID novo.

## 7. Efeitos colaterais

Nenhum nos números: os 12 cenários da Bio e os 7 da Hema dão os mesmos engPainel, Painel A6:Q12, Calc, Eng_Saida e grade
da Estatística da base. Só mudaram, como o ADR-095 previa:
- o A39 da Hema;
- B5, D5 e D4 da Estatística da Bio em "Tudo";
- os formatos de data;
- as listas.

Observações que já existiam na entrega 41: R5V-01, R5V-02, R5-01 com o caso simétrico do H5, e a recusa do H4/L4 sem
"na seta da célula" (P-02, conhecido).

## 8. Evidência (fora do repositório)

Em `%TEMP%\qcw\auditoria\2026-11-15_reteste05\tester\`:
- `valida_<produto>_{base,nova}.json` e os `.log`;
- `valida_Bioquimica_nova_painel_lote8977.json` e `valida_Hematologia_lote_sint.json`;
- `xml_Bio.json` e `xml_Hema.json`;
- `regressao_head.log` (a regressão oficial) e `regressao.log` (a 1ª rodada, com o `qa_adr095` da cópia de trabalho);
- os `regr_*/saida*/*.json` das suítes.
