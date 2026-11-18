# Tester/QA — VALIDAR — ciclo 2026-11-18_reteste08

**Ocorrência:** OC-0084 (Médio, Ambos). Um manual já integrado, recortado de ID a MOTIVO, mudava de dono.
**Correção validada:** ADR-098, commit 21c8ebd, entrega 45. **Quem corrigiu não fui eu:** este agente não escreveu
nenhuma linha do ADR-098, do `src/` nem dos instaladores.
**Base:** cópias novas da produção entregue, sem instalar nada (a entrega já aconteceu):
- `QC_Bioquimica.xlsm`, sha256 `a9ba35159655b20b`;
- `QC_Hematologia.xlsm`, sha256 `1d0d9fe8e720798f`.

Suítes rodadas a partir do `git archive 21c8ebd`. Pasta de trabalho: `%TEMP%\qcw\auditoria\2026-11-18_reteste08_tester\`.
Um Excel por vez, e só PIDs criados por mim.

## Veredito

| Produto | Veredito |
|---|---|
| Bioquímica | **VALIDADA** |
| Hematologia | **VALIDADA** |

Os 6 critérios de aceite foram aprovados nos dois produtos com o meu roteiro. **Ressalva da Hematologia:** o M15 do `qa_manual_substitui` (PQ × VBA, fora do escopo do ADR-098) ficou NAO_EXECUTADO por FailFast ambiental do Mashup, duas vezes (T08-03).

**Controle negativo:** o mesmo roteiro rodou numa cópia do backup anterior à entrega 45
(`_backup_pre_integracao_2026-11-18_0253`, Bioquímica da entrega 44). Ele **reprova** K1, K2, K3, K6 e K7, ou seja,
reproduz o defeito:
- o destino fica no nome de "(sem login)";
- o Q é recusado, e o n do Lactato N1 cai de 28 para 27;
- o S perde a substituição e fica CONFLITO_MANUAL;
- aparecem linhas RESULTADO_MANUAL_STATUS no Audit_Log.

O roteiro, portanto, distingue o código corrigido do antigo.

## Roteiro próprio

Script: `_arquitetura/agentes/ferramentas/repro_OC-0084.py`. É meu: não usa o script do Usuário Simulado nem o
`qa_adr097`, só `qa_final.QA`. Logs e JSON: `%TEMP%\qcw\auditoria\2026-11-18_reteste08_tester\repro_{Bio,Hema}.log` e
`repro_OC-0084_<produto>.json`.

**Pela interface.** O SendKeys não foi usado, porque mandaria teclas para a janela em foco do usuário. No lugar dele,
o roteiro reproduz a sequência da interface:
1. **Fase A.** Digita célula a célula, com os eventos ligados. Troca o dono e a hora do carimbo com os eventos
   desligados (TESTE-AUDITORIA_U1 a U4, como lançamento de outra pessoa). Roda o ATUALIZAR e grava a cópia.
2. **Fase B.** Reabre a cópia, sem retrato. Roda `mSeguranca.ReprotectAll`, que deixa a aba Digitar Resultados
   **protegida**, como depois do login.
3. **Seleção.** Uma célula é selecionada (`Select`). O `Worksheet_SelectionChange` tira o 1º retrato. É justamente o
   caminho que o ADR-098 declara sem teste automático.
4. **Recorte.** Ctrl+X/Ctrl+V é reproduzido como `Selection.Cut` + `ActiveSheet.Paste`, com os eventos ligados e a
   aba protegida.
5. **Range.Cut.** O K3 usa `Range.Cut(destino)`.

**Dados.** O analito afetado está no lote, no equipamento e na janela da Estatística:
- Bioquímica: Lactato, DIMENSION 2, lote 8977, out/2026;
- Hematologia: BASO#, XN1000, lote 522611, 2025.

Os valores são sintéticos: resultado real × fator, com 3 casas.

| Caso | Critério | Bio | Hema | Evidência principal |
|---|---|---|---|---|
| A0 | pré-condição | APROVADO | APROVADO | MAN_0001..0004 ATIVO no nome de U1..U4; interfaceamento substituído (BIO-371660 / HEM-213776) SUBSTITUIDO_MANUAL |
| K1 | 1 | APROVADO | APROVADO | MAN_0001 da linha 3 para a 4 (vazia), aba protegida: destino U1 com o carimbo original (Bio 27/10/2026 03:40; Hema 29/10/2025 03:40), origem sem carimbo; após o ATUALIZAR, USUARIO_MANUAL = U1, ATIVO, mesmo valor (4,606 / 0,150) |
| K2 | 2 | APROVADO | APROVADO | bloco P (U2, carimbo 03:25) e Q (U3, carimbo 05:00, DATA/HORA 03:30) sobe uma linha: P vai para a linha 7 e Q para a 8 (a antiga do P), cada um com o próprio carimbo; linha 9 limpa; após o ATUALIZAR os dois ficam ATIVO, sem "posterior", com o mesmo USUARIO_MANUAL |
| K3 | 3 | APROVADO | APROVADO | MAN_0004 (U4, SUBSTITUI_ID BIO-371660 / HEM-213776) da linha 12 para a 13: SUBSTITUI_ID e carimbo vão junto, origem limpa; após o ATUALIZAR o interfaceamento continua SUBSTITUIDO_MANUAL com ID_RELACIONADO = MAN_0004 |
| K4 | 4 | APROVADO | APROVADO | ID de MAN_0004 redigitado, igual e como "man_4": carimbo e SUBSTITUI_ID iguais. MAN_0005 não integrado (U5, carimbo 01/01/2020) redigitado igual e como "man_5": carimbo igual; após o ATUALIZAR, MANUAL_INCOMPLETO "DATA/HORA posterior ao lancamento" (a regra do motor continua valendo) |
| K5 | 5 | APROVADO | APROVADO | ver a lista abaixo |
| K6 | 6 | APROVADO | APROVADO | ver a tabela abaixo |
| K7 | trilha | APROVADO | APROVADO | nenhuma linha nova no Audit_Log sobre os IDs movidos no ATUALIZAR #2 |

**K5, fluxos vizinhos** (sem ATUALIZAR):
- manual novo: ganha carimbo novo, "(sem login)" e a hora atual;
- Delete de ID…MOTIVO de MAN_0004: a linha fica limpa, inclusive o SUBSTITUI_ID;
- cópia da linha de P colada numa linha vazia: ID novo (MAN_0007) com o carimbo de quem colou, e P mantém U2;
- Ctrl+D (`FillDown`) a partir de MAN_0006: a origem mantém ID e carimbo, e as linhas de baixo ganham IDs novos;
- RESULTADO editado na linha do Q: carimbo U3 mantido;
- SUBSTITUI_ID digitado numa linha vazia: a linha se preenche com o analito, o nível 1 e a data.

**K6, números conferidos com o oráculo próprio** (tblCQ_Final, PARTICIPA=SIM, mesmo filtro, Python `statistics`):

| | antes (produção) | depois do ATUALIZAR #1 | depois do recorte + ATUALIZAR #2 | conta à mão |
|---|---|---|---|---|
| Bio Lactato N1 | n 25, média 4,4488 | n 28, média 4,464678571428571 | n 28, média 4,464678571428571 | (111,22 + 4,606 + 4,597 + 4,588) / 28 |
| Bio Lactato N2 | n 17, média 1,188235294 | n 17, média 1,189176470588235 | igual | (20,2 − 1,3 + 1,316) / 17 |
| Hema BASO# N1 | n 50, média 0,1454 | n 53, média 0,145622641509434 | igual | +3 manuais |
| Hema BASO# N2 | n 51, média 0,326862745 | n 51, média 0,326941176470588 | igual | troca 1 por 1 |

Detalhes do K6:
- a Estatística é igual ao oráculo (tolerância relativa de 1e-9);
- nenhuma linha da Estatística (A14:F133) mudou entre o ATUALIZAR #1 e o #2;
- a `calc_independente.py` rodou nas cópias gravadas depois do #1 (`snapA1_*`) e do #2 (`snapB2_*`), com 0
  diferença em todas as colunas: Bio 62 linhas × 19 colunas; Hema 84 linhas × 27 colunas e o Painel 45/45.

## Regressão

Cada suíte rodou numa cópia nova da produção, uma por vez, com o código de 21c8ebd. Executor:
`regressao.py` (cópia do `rodar` do `entregar.py`, com `QC_PROD_ORIGINAL`). O gate da entrega 45 serviu só de
referência.

| Suíte | Bioquímica | Hematologia | Referência do gate (Bio / Hema) |
|---|---|---|---|
| qa_adr097.py (C01..C12, D01, M01, Q01, R01..R09c, F01, F02, Y00) | **28 PASS / 1 FAIL (R06)**, ver o achado T08-01 | 30 PASS / 0 FAIL | 29/0 e 30/0 |
| qa_manual_substitui.py | 20 PASS / 0 FAIL (849 s) | **18 PASS (M01..M14, M16..M18, M00); M15 NAO_EXECUTADO** (2 tentativas, FailFast do Mashup), ver T08-03 | 20/0 e 20/0 (retomada, 825 s) |
| qa_adr095.py | 10 PASS / 0 FAIL | 8 PASS / 0 FAIL | 10/0 e 8/0 |
| qa_inativar.py | 12 PASS / 0 FAIL | 12 PASS / 0 FAIL | 12/0 e 12/0 |

**R06 na Bioquímica: falha da fixture com os dados novos, não do produto.**

A pré-condição (`pre: false`) não se formou, porque o ATUALIZAR de 18/11 trouxe para a produção dois
interfaceamentos com a mesma chave e o mesmo instante:
- BIO-380496: Albumina, DIMENSION 2, lote 8977, N1, 17/11/2026 16:19:06, resultado 3,56;
- BIO-380497: a mesma chave e o mesmo instante, resultado 3,89.

O `qa_adr097.py` escolhe o 380496 como `ref` (R05) e o 380497 como `ref6` (R06). Os dois manuais de substituição
herdam a mesma chave e o mesmo instante. O motor marca o 2º como CONFLITO_MANUAL "resultado manual repetido (mesma
chave e instante)". Essa é a regra do ADR-097 Q01, funcionando como desenhada.

**Prova.** Rodei uma variante de diagnóstico, só na cópia do arquivo de código, em
`codigo\...\testes\qa_adr097_diag_tester.py`. A única diferença: `ref6` e `ref9` pulam o interfaceamento com a mesma
chave e o mesmo instante do `ref` (linhas 342 e 360 do `qa_adr097.py`). Resultado: **29 PASS / 0 FAIL**, com o R06
aprovado. O meu K3 cobre o mesmo comportamento com outro interfaceamento, e o R06 da Hematologia passou. Essa
variante não conta como execução oficial.

**Produção intacta:** sha256 conferido antes e depois de cada roteiro e suíte. Bio continua `a9ba35159655b20b...` e
Hema `1d0d9fe8e720798f...`. Nada foi salvo na raiz.

## Novos achados

| ID | Prioridade | Descrição | Encaminhamento |
|---|---|---|---|
| T08-01 | **Baixo** (teste) | `testes/qa_adr097.py` ~342 e ~360: `ref6` e `ref9` são escolhidos só por ID diferente. Com dois interfaceamentos reais na mesma chave e no mesmo instante (agora presentes na produção da Bio), o R06 reprova sem defeito do produto, e o gate da próxima entrega da Bio vai reprovar. Proposta: excluir de `ref6` e `ref9` os registros com a mesma (DATA_HORA, ANALITO, NIVEL, EQUIPAMENTO) dos já escolhidos (exemplo pronto na variante de diagnóstico) | Orquestrador / quem mantém as suítes |
| T08-02 | **Baixo** (dúvida de regra) | Na produção da Bio existem dois resultados de interfaceamento ATIVO com a mesma chave e o mesmo segundo e valores diferentes (BIO-380496 = 3,56 e BIO-380497 = 3,89, Albumina N1). Se o usuário quiser substituir os dois por manuais (ADR-082), o 2º manual cai em "resultado manual repetido" (ADR-097 Q01). Não é defeito do ADR-098, e não mudei nenhuma regra | Especialista em CQ (parecer) |
| T08-03 | **Info** (ambiente) | O `qa_manual_substitui.py` da Hema caiu **duas vezes** no M15 (PQ × VBA, refresh do Power Query numa instância nova), em 1737 s e em 1752 s, com "servidor RPC não disponível". O log de eventos mostra o FailFast 0x80131623 do Microsoft.Mashup.Container.Loader às 06:50:03 e às 07:22:58–07:23:12. Na 2ª vez, o Excel da tarefa agendada DB_SEAC estava aberto em paralelo, com os seus containers Mashup. M01..M14, M16..M18 e M00 passaram nas duas. **O M15 fica NAO_EXECUTADO.** Ele confere o motor PQ contra o VBA, que o ADR-098 não muda. O gate da entrega 45 o aprovou no mesmo código (retomada, 20/0, 825 s). Proposta: o Orquestrador o reexecuta com a máquina ociosa. Logs em `reg\Hematologia\qa_manual_substitui_tentativa1_failfast\` e `reg\Hematologia\qa_manual_substitui\` | Orquestrador (reexecutar o M15 da Hema) |

**Proposta de teste automatizado.** Incluir no `qa_adr097.py` o caso **R10**, com o caminho que o ADR-098 deixou sem
teste:
1. gravar a cópia com os eventos desligados e reabrir;
2. rodar `ReprotectAll`;
3. selecionar uma célula (o retrato vem do SelectionChange);
4. fazer `Selection.Cut` + `ActiveSheet.Paste` com a aba protegida.

O roteiro `repro_OC-0084.py` (fases A/B, K1–K3) serve de modelo. Ele mostrou que o caminho funciona e que o
controle negativo o reprova.

## Efeitos colaterais

Nenhum no produto. Os fluxos vizinhos (K5) e as 4 suítes afetadas passaram. A única reprovação, o R06 da Bio, vem da
fixture do teste com os dados novos (T08-01).

**Observações do método:**
- o pywin32 desloca em +3 h o `datetime` ingênuo escrito por `.Value`, então o roteiro grava as datas como serial
  em `Value2`;
- o `Workbook_BeforeSave` não grava por COM, e a cópia foi gravada com os eventos desligados. Isso já está registrado
  no `qa_migracao_adr070`.
