# Tester/QA: modo VALIDAR, ciclo 2026-11-16_reteste07 (ADR-097, entrega 44)

Correção validada: ADR-097, commit `7eb389c`, entrega 44 em produção desde 16/11. **Não fui eu quem a implementou.**
O Tester/QA só validou. A execução anterior do Tester neste ciclo foi interrompida às 19h26 de 17/11, pelo limite de
uso da conta, antes do relatório. Esta retomada (17/11 23h45 a 18/11 00h45) conferiu o que ela tinha executado, refez
só o que estava incompleto e rodou a regressão da Hematologia que tinha caído.

## Ambiente

| Item | Valor |
|---|---|
| Produção | `QC_Bioquimica.xlsm` sha256 `fdc173906fff6985…` e `QC_Hematologia.xlsm` `5b0e78ace7b98421…` (entrega 44 + ATUALIZAR, MANIFESTO linhas 59–60). Conferida antes e depois de **cada** rodada, da primeira (13:38 de 17/11) à última (`sha256sum` completo às 00:44 de 18/11): **não mudou** |
| Código das suítes | `git archive 7eb389c` (`tester_validar4\codigo`; a rodada das 18h36 usou `tester_validar3\codigo` = HEAD `f9962ad`, que só difere de `7eb389c` no `MANIFESTO_PRODUCAO.md`). **Não** é o working tree (ADR-098) |
| Cópias | todas novas, em `%TEMP%\qcw\auditoria\2026-11-16_reteste07\tester_validar*\`, nunca salvas, apagadas no fim (`copia_apagada: true`; `rm` das `teste_*` e `base` no fim de cada cadeia) |
| Excel | um por vez. PIDs meus: validação Bio 14964, Hema 5280; complementos 13024, 12640, 15472 (17/11) e 5332, 14508, 1360, 14608 (17/11 23h50); regressão Hema 13880, 1128, 4200, 15488, 19024, 12932, 4568. Todos encerrados pelo `qa_final.QA`/suíte. **Não toquei** nos PIDs 17836 e 11872 (não são meus: tarefa agendada do DB_SEAC ou outro processo). Nenhuma captura de tela |
| Scripts (meus) | `_arquitetura/agentes/ferramentas/valida_r7_adr097.py` (OC-0078..0083 e OC-0048, oráculo próprio), `valida_r7_estat_hema.py` (OC-0048, lote só de 2026 e lote de dois anos), `valida_r7_registros_com_dado.py` (OC-0082 com dado) e, nesta retomada, `valida_r7_lote_discrimina.py` (OC-0083 discriminante). **Não** reaproveitei `reteste07_usuario_f1*.py` nem a suíte `qa_adr097` nos critérios |
| Dados de teste | só nas cópias: usuários `TESTE_AUDITORIA_R7VA` e `_R7VB` (ANALISTA); manuais de 14/11/2026 (madrugada, a mais de 40 min de qualquer real) e de hoje (depois do carimbo antigo), valores sintéticos de 3 decimais. **Bio**: Lactato, lote 8977, DIMENSION 1. **Hema**: EO%, lote 625311, XN1000. IDs inválidos `TESTE1`, `MAN_1,5`, `MAN_12%`, `MAN_1E3`, `MAN_-12`. Registros: equipamento `TESTE-AUDITORIA EQ-A/B`, status `TESTE-AUDITORIA ST-1/2`, lote trocado só na cópia |
| Saídas | `tester_validar\valida_r7_<Produto>.json/.log`, `valida_r7_estat_Hematologia.json`, `valida_r7_registros_<Produto>.json`, `valida_r7_lote_discrimina_<Produto>.json`; `tester_validar3\` (cadeia de 17/11 18h36, regressão Bio) e `tester_validar4\` (complementos e regressão Hema desta retomada), tudo fora do repositório |

## Vereditos

| ID | Produto | Veredito | Resumo |
|---|---|---|---|
| OC-0078 | Bioquímica | **VALIDADA** | Itens (1)–(6) conferidos. Manual novo na linha apagada com carimbo novo e ATIVO; nenhum manual antigo perde o carimbo; n/média = oráculo. O carimbo do destino do recorte é a OC-0084 (limite do ADR-097 §1.5), que **não** reprova |
| OC-0078 | Hematologia | **VALIDADA** | Igual, com 3 níveis |
| OC-0079 | Ambos | **VALIDADA** | GUIA §3 no commit; aviso da colagem com o texto do ADR-097; colar ID..MOTIVO na aba **protegida**, sem senha, funciona |
| OC-0080 | Ambos | **VALIDADA** | 25 IDs avaliados na fórmula do DETALHE; inválidos dizem "ID fora do padrão: apague-o (um novo é gerado)", válidos dizem "atualize os dados"; apagar o ID inválido gera MAN_0023..0027, que ficam ATIVO |
| OC-0081 | Ambos | **VALIDADA** | CONFLITO_MANUAL mantido; A01 novo, igual no VBA e no PQ; a mensagem do ATUALIZAR conta a linha; corrigir a HORA a deixa ATIVO |
| OC-0082 | Ambos | **VALIDADA** | Lista vazia vira "(nenhum registrado)", que não filtra; texto livre recusado. Com dado, a lista é a da coluna e o filtro bate com o oráculo (refeito nesta retomada, ver abaixo) |
| OC-0083 | Ambos | **VALIDADA** | Lote digitado só com números fica texto, é aceito e filtra 1 de 2 (oráculo 1) nos dois lotes; fora da lista, recusado (complemento discriminante desta retomada) |
| OC-0048 | Hematologia | **VALIDADA** | D3 < B3 recusado com mensagem; A5 vermelho nos 4 casos (sem resultado no intervalo, também fora dele, B3 > D3, vazio quando cobre); grade = oráculo (≤ 1,1e-15) |

Nenhum `NAO_VALIDADA` e nenhum efeito colateral nas suítes rodadas.

### OC-0078: carimbo da linha manual (Ambos), VALIDADA

Roteiro próprio, na aba Digitar Resultados **protegida**, com os eventos ligados. O usuário A lança 10 manuais em
14/11; depois o usuário B faz os passos abaixo.

| Item | Passo | Bioquímica | Hematologia |
|---|---|---|---|
| (1) | Delete de ID..MOTIVO (MAN_0001; e duas áreas de uma vez, MAN_0004 + MAN_0005) | linha inteira vazia, sem USUARIO/REGISTRADO_EM; idem nas duas áreas | igual |
| (1) | B digita na linha apagada, **hoje, 13:41–13:43**, depois do carimbo antigo (13:40), o caso que antes era recusado | MAN_0020/0021/0022, carimbo **R7VB 13:44:04–12**, ATIVO no ATUALIZAR, sem aviso | MAN_0020/0021/0022, R7VB 14:10:00–06, ATIVO |
| (2) | Delete de DATA..MOTIVO, deixando o ID (MAN_0002) | ID **e** carimbo saem (linha vazia) | igual |
| (3) | Recortar MAN_0003 (já integrado) para outra linha | origem vazia, **sem carimbo**; destino com MAN_0003, ATIVO, RUN 26111402 e carimbo de quem moveu (R7VB): é a **OC-0084**, limite do ADR-097 §1.5, que não reprova | igual |
| (4) | Ctrl+D de MAN_0006 para 2 linhas | origem mantém MAN_0006 e o carimbo de A; as novas ganham MAN_0014/0015 (B), com o aviso do ADR-097 | igual |
| (5) | Editar o RESULTADO de MAN_0007 na própria linha | ID e carimbo iguais (R7VA 13:40:16), sem aviso | igual |
| (6) | SUBSTITUI_ID válido digitado só com o número | MAN_0017 ATIVO, "substitui o resultado do interfaceamento BIO-372856"; BIO-372856 SUBSTITUIDO_MANUAL, PARTICIPA NÃO | MAN_0017 ↔ HEM-372378, igual |
| (6) | Delete de ID..MOTIVO com o SUBSTITUI_ID fora da seleção | o SUBSTITUI_ID sai junto; BIO-373513 continua ATIVO | HEM-372453 ATIVO |
| (6) | Linha só com SUBSTITUI_ID inexistente + MOTIVO/EQUIPAMENTO | nada some; aviso "ID 999999987 não encontrado nos resultados…" | igual |
| — | Apagar só o ID (MAN_0008) | ID novo MAN_0019, carimbo de A mantido | igual |
| — | Carimbo órfão legado (gravado com os eventos desligados) | limpo no ATUALIZAR | igual |
| — | Nenhum manual antigo perde o carimbo | `carimbo_antigo_antes_atualizar2`: todas as linhas não tocadas iguais (a única "mudança" é a linha do "apagar só o ID", que trocou o ID e manteve o carimbo); `carimbos_vs_ref` do ATUALIZAR igual, exceto a linha do órfão (limpo, esperado) | igual |

**n e média contra o oráculo** (Python sobre a `tblCQ_Final`: PARTICIPA = SIM, analito, lote, nível, equipamento,
nov/2026). Tolerância de 1e-9 relativo; todas as diferenças ≤ 3,4e-16.

| | Antes | Depois do ATUALIZAR 2 | Depois do ATUALIZAR 3 (IDs inválidos corrigidos, HORA da colada corrigida) |
|---|---|---|---|
| Bio Painel N1 / N2 | 15 / 10 | **24 / 13** (oráculo 24 / 13) | **28 / 15** (oráculo 28 / 15) |
| Bio Estatística (H4 8977, 01–30/11) | — | 24 / 13 = oráculo | 28 / 15 = oráculo |
| Hema Painel N1 / N2 / N3 | 20 / 27 / 14 | **29 / 30 / 14** (oráculo igual) | **33 / 32 / 14** (oráculo igual) |
| Hema Estatística (lote 625311, 2026) | — | 50 / 69 / 35 = oráculo | 54 / 71 / 35 = oráculo |

Reserva **PQ** no mesmo estado: `diferencas_vba_pq = {}` nos dois produtos (36.240 e 48.690 linhas iguais). O PQ
levou 998 s na Bio e 1956 s na Hema. O VBA leva de 30 a 38 s. Os 1956 s da Hema ficam acima do histórico do M15 do `qa_manual_substitui` (só o ATUALIZAR
pelo PQ: 779 a 1236 s), mas é uma medida só, e o tempo do PQ varia muito (S6 do `qa_etl`: 170 a 1130 s). Não chamo de
defeito sem repetir; fica como observação.

### OC-0079: GUIA e colagem (Ambos), VALIDADA

- `git show 7eb389c -- GUIA_DE_USO.md` §3 orienta selecionar **de ID até MOTIVO**, Ctrl+C/Ctrl+V, usar a **hora real**,
  avisa que a mesma data, hora, nível e analito dá CONFLITO_MANUAL mesmo com outro resultado, diz para não copiar a
  linha inteira ("as colunas de USUARIO a DETALHE são do sistema…") e para corrigir o manual na própria linha.
- Travas (as duas): ID..MOTIVO e SUBSTITUI_ID livres; USUARIO, REGISTRADO_EM, RUN, STATUS e DETALHE travadas; a aba
  fica protegida o tempo todo.
- `Range(ID..MOTIVO).Copy(destino)` na aba protegida, sem desproteger: funcionou, sem pedir senha. Gerou o ID novo e
  o aviso "…Mude o que for diferente antes de clicar em ATUALIZAR DADOS. Atenção: com a mesma data, hora, nível e
  analito de outra linha, ela fica como repetida (CONFLITO_MANUAL), mesmo com outro resultado. Se for uma nova dosagem,
  use a hora real dela." (singular e plural, no Ctrl+D).

### OC-0080: DETALHE do ID fora do padrão (Ambos), VALIDADA

- Fórmula da coluna DETALHE igual à da `camada_dados.py` (`igual_camada_dados: true`), avaliada para 25 IDs: `MAN_0012`,
  `' man_7 '`, `man_00012`, `'MAN_12 '` e `MAN_0` dão "atualize os dados". `MAN_1,5`, `MAN_12%`, `MAN_1E3`, `MAN_-12`,
  `TESTE1`, `MAN_`, `MAN_12A`, `MAN_1.5`, `MAN_+12`, `MAN_ 12`, `HEM-12345`, `BIO-12345`, `12345`, `MAN12`,
  `MAN_(12)`, `MAN_$12`, `MAN_1/2`, `MAN_12/11`, `MAN_1e3` e `MAN_0x1A` dão "ID fora do padrão: apague-o (um novo é
  gerado)". Os 25 bateram com o esperado nos dois produtos.
- Nas linhas reais, antes do ATUALIZAR: os 5 inválidos mostram o texto novo, e MAN_0012/MAN_0013 mostram "atualize os
  dados". No ATUALIZAR, os inválidos ficam fora (E04 `MAN_INVALIDO_L16..L20`, "5 resultado(s) manual(is)
  incompleto(s)").
- Apagar o ID inválido: as linhas ganham MAN_0023..MAN_0027, mantêm o carimbo de A e ficam ATIVO no ATUALIZAR 3. Os
  E04 somem (Bio: 9 → 4 erros, os 4 restantes são de governança e já existiam; Hema: 5 → 0).

### OC-0081: colar e trocar só o RESULTADO (Ambos), VALIDADA

- Colar ID..MOTIVO de MAN_0010 e trocar só o RESULTADO: MAN_0016 fica **CONFLITO_MANUAL**, PARTICIPA NÃO (a regra do
  motor não mudou). O carimbo da origem não muda.
- A01, **idêntico no VBA e no PQ** (`A01_igual: true` nos dois produtos): "resultado manual repetido (mesma chave e
  instante) -> MAN_0010. O manual ficou FORA da estatistica; resultado repetido na aba Digitar Resultados (mesma data,
  hora, nivel e analito de MAN_0010): se for outra dosagem, corrija a HORA da linha MAN_0016 para a hora real; se for
  digitacao dobrada, apague a linha MAN_0016. Nao inative o MAN_0010: isso tiraria o resultado que vale."
- A mensagem do ATUALIZAR (e o `ResumoQA` do VBA e do PQ) traz "• 1 linha(s) da Digitar Resultados repete(m) data,
  hora, nível e analito de outra (ex.: MAN_0016)… Se for outra dosagem, corrija a HORA; se for digitação dobrada,
  apague a linha. Não inative a outra."
- Corrigir a HORA: MAN_0016 fica ATIVO no ATUALIZAR 3, a linha some da mensagem, e MAN_0010 continua ATIVO o tempo
  todo.

### OC-0082: filtros do Registros sem lista vazia (Ambos), VALIDADA

- **Lista vazia** (produção sem equipamento e sem status nos registros; na Hema também forçado na cópia): G4 e J4
  mostram `(nenhum registrado)`. `TESTE-AUDITORIA-R7V`, `DIMENSION 3`, `XN2000` e `ATIVO` são recusados com "Escolha
  o … na seta da célula. Para ver todos, apague o filtro (tecla Delete) ou clique em Limpar filtros.". Escolher
  `(nenhum registrado)` não filtra: K4 "Sem filtro · 2 registro(s)".
- **Com dado** (complemento): a 1a execução, às 18h36 de 17/11, deu `ok: false`. Diagnóstico: o **oráculo do meu
  script** contava a linha de entrada vazia da `tblRegistros` (3 linhas, 2 registros), e o sistema estava certo ("1
  de 2"). O script foi corrigido às 18h37 (mesmo critério do `ContarRegistros`), mas não chegou a rodar de novo. **Refiz
  nesta retomada** (23h50, cópias novas), e deu `ok: true` nos dois produtos. A lista de G4/J4 é a dos valores da
  coluna (`TESTE-AUDITORIA EQ-A/EQ-B`, `ST-1/ST-2`), sem `(nenhum registrado)`. Fora da lista é recusado (inclusive
  `(nenhum registrado)`). Cada item filtra 1 de 2, que é o oráculo. Apagar a coluna volta a `(nenhum registrado)`. A
  1a execução está guardada como `valida_r7_registros_<Produto>_run1_oraculo_contava_linha_vazia.json`.
- Com o dado devolvido, as listas de B4/C4/G4/J4 ficam iguais às de antes (`igual_antes: true`).

### OC-0083: lote digitado só com números (Ambos), VALIDADA

- `valida_r7_adr097.py`: B4 tem formato `@`; o lote digitado (`FormulaLocal`) fica `str`, é aceito, e o K4 mostra
  "Filtros ativos: Lote 8974 · 2 de 2" (Hema 522611, 2 de 2), igual ao oráculo. `99999`, `8977-EE` e `1` são
  recusados.
- Ressalva da evidência: todos os registros eram do mesmo lote, e "2 de 2" não distingue "filtra" de "não filtra".
  **Complemento desta retomada** (`valida_r7_lote_discrimina.py`): na cópia, troquei o lote do 1º registro (Bio 5352,
  Hema 528211) e digitei cada lote. Os dois mostram **1 de 2**, igual ao oráculo, com o valor em texto e a validação
  aceita. `99999`, `1` e o lote com um dígito a mais são recusados. `ok: true` nos dois produtos.

### OC-0048: Estatística da Hema, ano De/Até (Hematologia), VALIDADA

| Caso | A5 (fonte 192 = vermelho, negrito) | Grade |
|---|---|---|
| Lote 522611 (só 2025), 2025–2025 | vazio | = oráculo (máx 6,3e-16) |
| 522611, 2026–2026 | "⚠ O lote 522611 não tem resultado em 2026 (tem em 2025). Ajuste o Ano De/Até (B3/D3)." | vazia, como o oráculo |
| 522611, 2024–2024 | idem, "em 2024 (tem em 2025)" | vazia |
| 522611, 2025–2026 | vazio | = 2025–2025 (mesmo hash) |
| 522611, B3 2026 > D3 2025 (digitado depois) | "⚠ Ano Até (2025) menor que o Ano De (2026): a Estatística usa só 2026. O lote 522611 não tem resultado em 2026…" | = 2026–2026 |
| 625311 (só 2026), 2025–2025 / 2026–2026 / 2025–2026 / 2026–2025 | aviso "não tem em 2025 (tem em 2026)" / vazio / vazio / "…usa só 2026." | n total 0 / 3808 / 3808 / 3808, = oráculo |
| 533811 (2025 e 2026), 2025–2025 / 2026–2026 / 2025–2026 | "também tem resultado em 2026, fora do Ano De/Até: os números abaixo não são o lote inteiro." / idem "em 2025" / vazio | 644 / 5208 / 5852, = oráculo (máx 1,1e-15) |

D3 tem validação `>= $B$3` com "Ano Até deve ser igual ou maior que Ano De.". D3 recusa 2025 e 2024 com B3 = 2026 e
aceita 2026 e 2027. Os números da grade não mudam por causa do aviso.

## Regressão (código das suítes = `7eb389c`, cada suíte numa cópia nova da produção)

| Suíte | Bioquímica (17/11 18h38–19h24, `tester_validar3`) | Hematologia (17/11 23h52 – 18/11 00h44, `tester_validar4`) |
|---|---|---|
| qa_adr097 | **PASS** 18/0 (75 s) | **PASS** 19/0 (78 s, rodada de 17/11 19h26, `tester_validar3`) |
| qa_adr095 | **PASS** 10/0 (26 s) | **PASS** 8/0 (28 s) |
| qa_adr090 | **PASS** 21/0 (244 s) | **PASS** 23/0 (242 s) |
| qa_manual_substitui | **PASS** 20/0 (430 s) | **PASS** 20/0 (1298 s; a suíte inteira levou 1209 s na rodada de 17/11 10h27; o M15, do PQ, sozinho variou de 779 a 1236 s no histórico) |
| qa_etl | **PASS** 64/0 (792 s) | **PASS** 63/0 (857 s; a Hema tem um caso a menos, como na rodada de 10h42) |
| qa_equiv_qa_adr094 | **PASS** 3/0 (908 s) | **PASS** 3/0 (396 s) |
| qa_adr087 | **PASS** 14/0 (209 s) | **PASS** 14/0 (205 s) |
| qa_inativar | **PASS** 12/0 (87 s) | **PASS** 12/0 (105 s) |
| **Total** | **8/8 OK** | **8/8 OK** |

- **Hematologia de 17/11 19h26 (7 suítes "FALHOU (0 s)"):** confirmei que os 7 logs (`tester_validar3\reg\Hematologia\
  qa_*.log`) têm **0 bytes** e nenhum JSON. A cadeia caiu com a sessão (limite de uso) logo depois do `qa_adr097`.
  Não conto isso como falha nem como aprovação; as 7 foram refeitas acima.
- **Rodada descartada de 17/11 09h07–10h42 (`tester_validar\regressao`):** usou as suítes do **working tree** (ADR-098
  em andamento). Por isso o `qa_adr097` da Hema falhou em 5 casos novos do ADR-098, entre eles o R01 da OC-0084, que
  ainda não tem correção na produção. O `qa_etl` da Bio teve um E08g: 4 resultados novos (BIO-380233..36) chegaram da
  fonte viva durante o teste, um efeito ambiental, e a rodada das 18h36 deu 64/0. A rodada também foi interrompida às
  10h42. Não vale como evidência para `7eb389c`.
- **Não executado pelo Tester (fora do escopo pedido):** as outras suítes da lista oficial (`entregar.SUITES`, 35 no
  total), inclusive o NUCLEO (`qa_seguranca`, `qa_final`, `qa_sync`, `qa_migracao_adr070`) e o `qa_adr092`.
  `NAO_EXECUTADO` por mim. O gate da entrega 44 (`entregar.py --direcionada`) as rodou: 18 suítes PASS por produto
  (MANIFESTO linhas 57–58). O ADR-097 mexe no texto do QA (A01, `ResumoQA`, `QA_INTEGRACAO.m`) e no evento da aba
  manual, não no cálculo estatístico; a equivalência VBA = PQ (qa_equiv, qa_etl e a reserva PQ do meu script) passou.

## Produção intacta

**Sim.** Bio `fdc173906fff6985dcf72bd98704db01b878e8720cd30bf530c15a315aa657c6` e Hema
`5b0e78ace7b984218f2b516b7830c6c5ad3b78a8182d98be2e2be9fb6d49b1fc`, iguais antes e depois de todas as rodadas (última
conferência às 00:44 de 18/11). Nada foi salvo na raiz nem no DB_SEAC.

## Novos achados

| ID | Prioridade | Achado |
|---|---|---|
| T7V-01 | **Médio** (processo/privacidade) | Na raiz do repositório há um `~$QC_Bioquimica.xlsm` (arquivo de dono do Excel) de **17/11 10:29**, sem Excel com ele aberto agora. Ele **não** está no `.gitignore` (`git check-ignore` não o pega) e aparece como `??` no `git status`. Um `git add -A` o levaria para o repositório público, e esse arquivo guarda o nome de quem abriu a planilha. O arquivo também indica que alguém abriu a produção da Bio no Excel às 10h29 e não fechou de forma limpa (o hash não mudou). Sugestão ao Orquestrador: confirmar quem abriu, apagar o `~$` com o Excel fechado e pôr `~$*` no `.gitignore`. Não mexi nele |
| T7V-02 | Baixo (infra de teste) | O `qa_etl` grava `testes/resultados/etl_rastreabilidade_<Produto>.json` na árvore de onde roda. A rodada das 9h com o working tree deixou os dois arquivos modificados (`M` no `git status`, 48+/96−). Rodando do archive não suja o repositório. Sugestão: o `qa_etl` gravar na pasta de saída da suíte, ou o Orquestrador descartar essas mudanças antes de commitar |
| (conhecido) | Médio | OC-0084 confirmada de novo, nos dois produtos: no recorte, o destino fica com o carimbo de quem moveu (R7VB). O ADR-098 (commit `21c8ebd`, entrega 45) trata disso; não faz parte desta validação |

Nenhum achado novo no produto (cálculo, dados ou telas) nas ocorrências validadas.
