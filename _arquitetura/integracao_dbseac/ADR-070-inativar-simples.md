# ADR-070 — Aba Inativar simples (ID + analito + motivo) e o ID na dica do Levey-Jennings

**Data:** 06/10/2026 · **Status:** aceito · **Instalador:** `instalar_adr070.py` (no `entregar.py`, antes da blindagem)
**Mesma família:** ADR-057 (soft-delete por ID, LEFT OUTER JOIN), ADR-066 (QA-ETL-001, D06/D14)

## Pedido do usuário

1. Ao passar o mouse (ou clicar) num ponto do Levey-Jennings do Painel, ver o **ID** do resultado, além de resultado,
   corrida e data. Assim ele escolhe no gráfico o que inativar.
2. A aba Inativar tinha "um monte de fórmulas sem funcionar". Ele quer a aba **simples e direta**, como a planilha de
   exemplo (uma tabela onde se digita o índice; o Power Query faz o resto).
3. Inativar = informar o **ANALITO** e o **ID**.

## Causa das "fórmulas sem funcionar"

As colunas de conferência (XLOOKUP na tblCQ_Final, COUNTIF nos comentários) tinham formato de célula **texto** (`@`).
Com esse formato, o Excel guarda a fórmula como **texto literal** e a célula mostra `=IF(tblInativacao…`. O mesmo
defeito estava em **Digitar Resultados** (STATUS e DETALHE, justamente as colunas que a instrução manda conferir) e em
**COMENTARIOS_TECNICOS** (ANALITO e INATIVADO). As de formato número/data (NIVEL, RUN, DATA_HORA) calculavam.

## Decisão

**Aba Inativar** (`tblInativacao_NaoConformes`) tem só o que o usuário preenche e a trilha de auditoria:

| ID_REGISTRO | ANALITO | REGISTRAR - LJ | MOTIVO | DATA_INATIVACAO | USUARIO |
|---|---|---|---|---|---|
| digitado: o número que aparece na dica do gráfico (o evento normaliza para `BIO-`/`HEM-`) | lista do cadastro (`lstAnalitos`) | caixa, marcada por padrão (X vermelho); desmarcada = não plota | texto | carimbo do VBA, travado | carimbo do VBA, travado |

- As oito colunas de fórmula saíram (ANALITO por XLOOKUP, NIVEL, LOTE, DATA_HORA, RUN, RESULTADO, PLOTAGEM_LJ,
  JUSTIFICATIVA). Ninguém as lia (VBA, Power Query, testes, BI); a conferência agora é a dica do gráfico e o QA.
- **O ANALITO é conferência.** A inativação só vale quando o analito da linha é o do resultado (maiúsculas e espaços
  nas pontas não contam). Se o ID não é daquele analito, ou se a linha não tem analito, **a inativação não é aplicada**
  (o resultado continua ATIVO) e o QA acusa **ERRO**: **E10** "ID não pertence ao analito informado" e **E11**
  "Inativação sem analito informado". É a opção segura: um número digitado errado não tira da estatística o resultado
  de outro analito. A correção é editar a linha. ID inexistente continua E03; ID repetido continua E02 (vale a 1ª linha).
- **MOTIVO é a justificativa.** `TEM_JUSTIFICATIVA = SIM` com o MOTIVO preenchido **ou** com um comentário técnico.
  Inativado sem nenhum dos dois continua sendo **E01** (ERRO DE GOVERNANÇA). A aba COMENTARIOS_TECNICOS continua
  existindo para outros comentários, e o comentário lá continua aceito como justificativa.
- tblCQ_Final ganha a coluna **MOTIVO_INATIVACAO**. Ela, DATA_INATIVACAO e USUARIO_INATIVACAO só aparecem quando a
  inativação vale. O Audit_Log grava o motivo junto do comentário.
- **Inativado continua na base** (LEFT OUTER JOIN, nunca anti-join): PARTICIPA_ESTATISTICA = NÃO, X_VERMELHO ou
  NAO_PLOTAR. Nada muda nisso.
- **Reativar = apagar a linha.** Ao apagar a célula do ID, o evento limpa a linha inteira (caixa, carimbo, analito e
  motivo); nada fica esperando o próximo ID digitado ali.

**Dica do gráfico.** A classe `clsCht` (eventos dos gráficos do Painel) só existia dentro dos .xlsm e lia a aba
antiga Calc: mostrava "RUN <índice do ponto>", a data e o valor. Agora ela é fonte versionada
(`src/comum/clsCht.cls`, instalada pelo `codigo_atual.py`) e chama `mUI.DicaPonto(nivel, serie, ponto)`:

- `ID 314216 · RUN 25103103 · 25/10/2025 14:32 · 1,234` (ponto normal);
- `ID 314216 · INATIVADO · RUN … · … · …` (X vermelho);
- `RUN … · data` (linhas de limite e calibração).

O motor (`mEstatistica.AtualizarCalc`) publica, junto das séries, o ID e a DATA_HORA de cada ponto e de cada X no
Eng_Saida (**AM:AO** ID do ponto N1..N3, **AP:AX** ID de cada X, **AY:BA** e **BB:BJ** as datas/horas). O número mostrado
é o curto, sem prefixo: é o que se digita.

**Duplo clique no ponto** (opcional do pedido, feito): pergunta o motivo e grava a linha na aba Inativar (ID, analito
em tela, caixa marcada, motivo, carimbo) por `mIntegracao.RegistrarInativacao`. Não roda a atualização sozinho: o
usuário confere e clica ⟳ ATUALIZAR DADOS. No X vermelho, avisa que já está inativado. O mesmo ID não entra duas vezes.

## Migração (arquivos em uso)

`instalar_adr070.py` refaz a tabela no layout novo **preservando cada linha na mesma posição**: ID, REGISTRAR - LJ,
DATA_INATIVACAO e USUARIO idênticos. O ANALITO das linhas antigas vem da tblCQ_Final pelo ID (o que a fórmula antiga
buscava). O MOTIVO fica vazio: a justificativa antiga continua valendo pelos comentários (regra OU). Em MODO HISTÓRICO
o instalador roda o ATUALIZAR DADOS e exige o **mesmo conjunto de inativados, com a mesma plotagem**, e nenhum E10/E11.
Em SEAC, a consulta nova vale no próximo ATUALIZAR DADOS. Enquanto a tabela estiver no layout antigo (sem a coluna
MOTIVO), a consulta nova aplica a regra antiga, só pelo ID: reinstalar a consulta antes de migrar não reativa ninguém.
É idempotente: na 2ª vez não migra nada.

Digitar Resultados e COMENTARIOS_TECNICOS: as colunas de fórmula passam a ter formato Geral e as fórmulas são
reescritas; elas voltam a calcular. Em `camada_dados.py`, coluna de fórmula nunca mais tem formato `@`.

## Prova

- `testes/qa_inativar.py` N01–N09: estrutura e fórmulas calculando; dica de ponto conhecido (ID/RUN/data/valor iguais
  à Principal - Resultados); (1) ID + analito certo aplica (PARTICIPA = NÃO, X na corrida certa); (3) MOTIVO satisfaz o
  E01 sem comentário; (2) analito errado não aplica + E10, sem analito + E11; (5) dica do X; duplo clique (núcleo);
  corrigir o analito aplica; apagar o ID reativa e limpa a linha. Entrou na lista de suítes do `entregar.py`.
- `testes/qa_migracao_adr070.py` M01–M05 (roda sobre um arquivo ainda no layout antigo): (4) as linhas antigas ficam
  na mesma posição, com o analito preenchido, os mesmos inativados e a mesma governança; idempotência.
- `qa_final`, `qa_etl` e `qa_casos_extremos` ajustados: as inativações levam o analito. O E07 confere
  INATIVACAO_REGISTRADA = (1ª linha do ID na Inativar com o analito certo) ∩ final.
