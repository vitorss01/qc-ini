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
| digitado: o número que aparece na dica do gráfico (o evento normaliza para `BIO-`/`HEM-`) | lista do cadastro (`lstAnalitos`), em estilo **Aviso** (revisão 07/10: aceita analito sem cadastro) | caixa, marcada por padrão (X vermelho); desmarcada = não plota | texto | carimbo do VBA, travado | carimbo do VBA, travado |

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

## Revisão adversarial (07/10/2026)

Uma revisão confirmou seis defeitos nesta entrega. Todos corrigidos; o último é uma limitação registrada.

1. **Analito sem cadastro tinha deixado de ser inativável (regressão).** A validação de ANALITO era *Parar* com a
   lista do cadastro. O resultado de um analito sem cadastro (Bio FERR, TNIH, UCFP, CRE2/PHOS/MALB de urina; Hema
   RET-HE, IPF, IPF#) continua ATIVO na tblCQ_Final com ANALITO = nome de origem, e esse nome não podia ser digitado.
   **Agora a validação é *Aviso*:** a lista continua sugerindo o cadastro, e um nome fora dela pede confirmação
   (a mensagem manda digitar como está na coluna ANALITO da Principal - Resultados, ex.: FERR). A conferência
   ID × analito continua no Power Query (`_ANI = AnKey(ANALITO)`; senão E10/E11). Prova: `qa_inativar` N10.
2. **Linha "livre" com carimbo residual.** `RegistrarInativacao` (duplo clique) reaproveitava a primeira linha com ID,
   ANALITO e MOTIVO vazios. Se ela tivesse data/usuário de sobra, o carimbo antigo ficava e a caixa não era marcada, e
   o ponto virava NAO_PLOTAR em vez de X vermelho. **Agora só reaproveita linha totalmente vazia** (as seis colunas,
   caixa desmarcada). Também o `EntradaMudou` percorre todas as áreas de uma seleção com Ctrl, que era uma das origens
   da sobra. Prova: N11.
3. **Dica do LJ desatualizada depois de republicar.** A caixa "tip" só era reescrita quando o elemento sob o mouse
   mudava. Trocar o período, o lote ou o analito com o mouse parado num ponto deixava na tela o ID de outra corrida.
   **Agora o `AtualizarCalc` chama `mUI.LimparDicas`** a cada publicação: apaga o texto da caixa em todos os gráficos e
   muda a geração das dicas. O `clsCht` põe a geração na chave e reescreve no próximo movimento. Prova: N12.
4. **M03 passava sem testar nada em SEAC.** Em SEAC o instalador não roda o ATUALIZAR, e o M03 comparava a tblCQ_Final
   de antes da migração com ela mesma. **Agora a `qa_migracao_adr070` põe a cópia em HISTÓRICO** antes de inativar no
   layout antigo, o instalador roda a consulta nova sobre a tabela migrada, e o M03 **falha** se o log não mostrar
   essa atualização.
5. **Na entrega em SEAC nada conferia que as inativações reais sobrevivem.**
   - O instalador confere em **qualquer modo** o conjunto (ID, REGISTRAR - LJ, DATA_INATIVACAO, USUARIO), lido direto
     da tabela antes e depois, e aborta se divergir. Na migração ele também exige o ANALITO de todo ID que tem
     resultado.
   - A migração também busca o analito do **manual digitado e ainda não atualizado** (`tblResultados_Manuais`). Antes,
     esse MAN_ inativado migrava com ANALITO vazio e voltava a ATIVO no próximo ATUALIZAR.
   - A `qa_migracao_adr070.py` entrou nas suítes do `entregar.py`. O entregar copia a produção original antes de
     instalar (`<produto>/producao_original/`, SHA-256 conferido) e a passa às suítes por `QC_PROD_ORIGINAL`. Não há
     caminho de rede no código.
   - A suíte testa a migração sobre essa cópia (M01–M04) e acrescenta o **M06**: a cópia instalada tem o mesmo
     conjunto da produção; todo ID com resultado tem ANALITO; e um ATUALIZAR (HISTÓRICO, consulta nova) numa cópia dela
     dá os mesmos INATIVADOS, com a mesma plotagem e sem E10/E11 nas linhas migradas. Se a produção já estiver no
     layout novo, o log diz que M01–M04 não se aplicam e o M06 roda.
6. **InputBox ANSI (limitação, sem UserForm).** O motivo do duplo clique passa pelo `InputBox` do VBA, que é um
   diálogo ANSI (cp1252 no Windows pt-BR). Símbolos fora dele (≥, σ, Δ…) chegam como "?" ou trocados por "melhor
   aproximação", e o MOTIVO é a justificativa oficial. O que foi feito:
   - o texto mostrado no diálogo e nos avisos passa por `mUI.TextoAnsi`, que troca os símbolos conhecidos por texto
     visível (">=", "sigma", "delta"…) e o resto por "[?]", nunca por um "?" mudo;
   - o prompt pede que esses símbolos sejam escritos por extenso;
   - se o motivo digitado tiver "?", o sistema pergunta antes de gravar (Não = cancelar).
   Quem precisa do símbolo digita o motivo direto na aba Inativar, que aceita Unicode. Os acentos do português existem
   no cp1252 e não são afetados.
