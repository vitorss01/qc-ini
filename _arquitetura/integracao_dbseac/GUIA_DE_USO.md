# QC_INI — Guia de uso (implantação, outubro/2026)

QC_Hematologia.xlsm e QC_Bioquimica.xlsm · origem dos resultados: DB_SEAC (interfaceamento SIPEC), que a tarefa
agendada atualiza a cada 2 minutos.

## 1. Rotina diária

1. Abra o QC e faça o login.
2. Clique em **⟳ ATUALIZAR DADOS**. Leva de 15 a 25 segundos (ADR-079; antes, 2 a 4 minutos). A mensagem final
   diz quantos resultados novos entraram e a **data do resultado mais recente recebido** — se ela estiver velha, o
   problema está no DB_SEAC ou no SIPEC, não no QC.
3. No **Painel**, escolha o analito pelas setas ou pela lista. Os gráficos de Levey-Jennings mostram os níveis lado a
   lado.
   - **Trocar de lote (H3):** se o filtro atual não mostra nada do lote escolhido (outro ano, outro equipamento, analito
     que o lote não tem), o Painel ajusta sozinho o período ("Tudo"), o equipamento e o analito. A barra de status diz
     o que mudou.
   - **Eixo X = corrida do lote (ADR-080):** cada nível tem a sua sequência 1, 2, 3... desde o primeiro resultado
     daquele lote (cada passagem do controle é uma corrida, mesmo à noite). Trocou de lote, recomeça do 1. Com filtro
     de período, aparece o número real da corrida no lote.
4. Passe o mouse num ponto do gráfico. Aparecem o **ID**, a corrida (RUN), a data/hora e o valor. Nos pontos já
   inativados (X vermelho) aparece também **INATIVADO**.
5. Confira o Status da corrida, as regras de Westgard violadas e o histórico do analito. As regras do mesmo nível (2-2s
   entre corridas) olham os resultados ativos seguidos daquele nível, mesmo com corrida só de outro nível no meio
   (ADR-083; ver a seção 2).
6. Olhe a aba **QA_INTEGRACAO**:
   - **ERRO**: precisa de ação;
   - **ALERTA**: o sistema protegeu o dado, mas alguém deve olhar;
   - **INFO**: só registro.

## 2. Inativar um resultado (não conforme / repetição) — ADR-081

**Pelo gráfico (o jeito simples):**

1. No Painel, dê **duplo clique** (ou clique com o botão direito) no ponto.
2. Abre a janela **Inativar o resultado ID …** com o analito, o nível, a corrida, a data/hora, o valor e o lote —
   confira que é o ponto certo.
3. Escolha o **motivo** na lista (em "Outro", escreva o complemento). A caixa **Mostrar como X vermelho no gráfico**
   já vem marcada; desmarque se o resultado não deve aparecer no gráfico.
4. Clique em **Inativar**. O gráfico se atualiza sozinho em alguns segundos e mostra "✓ ID … inativado". Não precisa
   clicar em ATUALIZAR.

**Reativar:** duplo clique no **X vermelho** → janela **Reativar** → motivo → **Reativar**. Resultado inativado sem X
(caixa desmarcada) não aparece no gráfico: reative pela aba Inativar (apague a linha e clique em ATUALIZAR).

**Como as corridas ficam:** a corrida do lote conta só os resultados **válidos**. Se a ID 54565 era a corrida 3 e foi
inativada, o resultado seguinte passa a ser a corrida 3 — o gráfico continua 1, 2, 3… sem lacuna. O X vermelho fica
**entre** as corridas vizinhas (ex.: entre a 2 e a 3), sem número. O resultado é sempre identificado pela **ID**, que
nunca muda; a dica do ponto mostra a corrida, a ID e o RUN do dia. Reativar devolve a numeração exatamente como era.

**Atenção — na Bioquímica, inativar ou reativar pode mudar o veredicto de corridas POSTERIORES do mesmo nível
(ADR-083).** As regras de Westgard do mesmo nível comparam os resultados **ativos seguidos** daquele nível: o 2-2s liga
dois N1 seguidos acima de +2DP mesmo que haja uma corrida só do N2 entre eles. Por isso:

- inativar um N1 que estava entre dois N1 acima de +2DP **cria** um 2-2s na corrida seguinte (que pode já ter sido
  liberada) e ela vira REJEITADO;
- reativar esse N1, ou digitar entre os dois um manual **dentro de ±2DP**, **desfaz** o 2-2s (um manual acima de +2DP
  forma outro 2-2s).

Na **Hematologia** as regras do mesmo nível são só complementares: aparecem na aba Eventos_Westgard como COMPLEMENTAR,
mas nenhum veredicto muda.

A regra oficial que entra ou sai fica no **Audit_Log**, com o RUN, o analito, o nível e a regra:
WESTGARD_EVENTO_NOVO / WESTGARD_EVENTO_REMOVIDO quando a corrida já tinha sido avaliada (o parecer começa com
`[RETROATIVO]`) e WESTGARD_EVENTO_CORRIDA_NOVA quando a regra disparou numa corrida nova (rotina). **A trilha cobre
todos os lotes e, na Bioquímica, os dois equipamentos** — não só o que está no Painel: o sistema guarda uma foto das
regras oficiais de cada lote × equipamento (aba técnica oculta) e, a cada ATUALIZAR ou inativação pelo gráfico, audita
o lote cujo resultado mudou (inclusive o inativado pela aba Inativar, de qualquer lote). Não precisa trocar o lote do
Painel antes de inativar. Depois de inativar, confira o status das corridas seguintes do mesmo nível no Painel.
As regras entre níveis (R-4s, 2-2s e 2 de 3 dentro da corrida, 4-1s, 3-1s, 8x e 6x oficiais) continuam pela corrida do
dia.

**Pela aba Inativar (quem preferir digitar):** na próxima linha livre, digite a **ID** (o analito se preenche
sozinho), marque ou não **Registrar-LJ** (X vermelho), escreva o **motivo** e clique em ⟳ **ATUALIZAR DADOS**.

- O resultado não é apagado: continua na base, com PARTICIPA = NÃO, e a trilha de auditoria registra data, usuário,
  motivo e a corrida que ele tinha.
- Se o ID não for do analito informado, o QA acusa (E10) e a inativação não é aplicada.
- Quem pode inativar: qualquer usuário com login (a chave PAPEL_MIN_INATIVAR da aba Cfg_Integracao restringe a um
  papel mínimo, se o laboratório quiser).

## 3. Digitar um resultado (falha do interfaceamento)

1. Na aba **Digitar Resultados**, preencha data, hora, lote, nível, analito, resultado e motivo. O ID (MAN_0001…) é
   gerado sozinho.
2. Atualize e confira o STATUS (e a coluna DETALHE):
   - **ATIVO** = entrou na estatística;
   - **CONFLITO_MANUAL** = o mesmo resultado já veio pelo interfaceamento (mesma chave, até 30 min) e vale o
     automático. O manual fica guardado, fora do cálculo. Para o manual valer no lugar dele, use o SUBSTITUI_ID
     (abaixo);
   - **MANUAL_INCOMPLETO** = falta campo (o DETALHE diz qual).

**Substituir um resultado do interfaceamento pelo manual (ADR-082 — por escolha, nunca sozinho):**

1. Passe o mouse no ponto do gráfico e anote o número do **ID** do resultado que vai ser substituído.
2. Na próxima linha livre da aba **Digitar Resultados**, digite primeiro esse número na coluna **SUBSTITUI_ID** (a
   última). Equipamento, lote, nível, analito, data e hora se preenchem sozinhos com os do resultado indicado (só o que
   estiver vazio). Se o ID não existir, aparece um aviso e nada é preenchido. Se algum desses campos já estava
   preenchido com outro valor (por exemplo o equipamento padrão), aparece um aviso dizendo qual: corrija-o, senão nada
   é substituído (QA E12 "tem outra chave").
3. Digite o **RESULTADO** correto e o **MOTIVO** (obrigatório para substituir) e clique em ⟳ **ATUALIZAR DADOS**.
4. Confira: o manual fica **ATIVO** (DETALHE "substitui o resultado do interfaceamento …"); o automático fica
   **SUBSTITUIDO_MANUAL** — guardado na base, fora da estatística e do gráfico, sem corrida. A dica do ponto mostra
   "ID MAN_0007 (substitui 54565)". A aba QA_INTEGRACAO registra a troca (INFO **I04**: os dois IDs, os dois valores, a
   diferença e o motivo) e o Audit_Log grava RESULTADO_SUBSTITUIDO_POR_MANUAL.

- Só vale o ID de um resultado do **interfaceamento** com o **mesmo** equipamento, matriz, lote, nível e analito, que
  não seja retransmissão e que nenhum outro manual já substitua. Fora disso nada é substituído e o QA dá **ERRO E12**
  dizendo por quê. Dois manuais para o mesmo ID: vale o primeiro; o segundo fica CONFLITO_MANUAL.
- Resultado do interfaceamento **sem valor** ou **sem data/hora** também pode ser substituído (é o caso típico).
- Resultado já **inativado** pelo usuário continua **INATIVADO** (X vermelho e não conformidade preservados); o manual
  vale do mesmo jeito. O inativado não ganha corrida própria: o X vermelho aparece na corrida do manual e nenhuma
  outra corrida do dia muda (nem se você inativar ou reativar o automático depois).
- O manual entra na corrida pela **própria** data/hora (a que se preencheu sozinha é a do resultado substituído). Se
  você mudar a hora para fora do bloco dos outros níveis (mais de 10 min), ele abre uma corrida própria e as seguintes
  do dia andam +1.
- **Inativar o manual NÃO traz o automático de volta** (a janela de inativar avisa). Para voltar ao automático,
  apague o SUBSTITUI_ID (ou a linha do manual) e clique em ATUALIZAR DADOS; o Audit_Log grava
  SUBSTITUICAO_POR_MANUAL_DESFEITA. Se o automático também estava inativado, reative-o (o I04 diz isso).
- Digitou o manual errado, inativou-o e digitou outro com o mesmo SUBSTITUI_ID? O primeiro continua substituindo e o
  segundo fica CONFLITO_MANUAL (o QA E12 explica): apague o SUBSTITUI_ID do primeiro e atualize — o segundo passa a
  valer (Audit_Log: SUBSTITUICAO_POR_MANUAL_ALTERADA).
- Corrigir depois o valor ou a hora de um manual fica no Audit_Log (RESULTADO_MANUAL_ALTERADO).
- Manuais antigos (sem SUBSTITUI_ID) continuam como estavam: nada muda no histórico.

## 4. Especificações de qualidade (aba Especificações, nos dois QCs)

A antiga aba Analitos agora se chama **Especificações**. Ela guarda:
- o catálogo de analitos;
- o desempenho (Q) e a fonte do ETp (S);
- a **biblioteca anual** de especificações, nas linhas 46–89 (2025 a 2029): CLIA ETp, VB-CVi, VB-CVg, FAB ETp e FAB
  CVTp de cada ano.

A coluna **S** escolhe a especificação de cada analito:
- **CLIA** usa a coluna K;
- **VB** (variação biológica: CVi e CVg nas colunas O e P, desempenho na Q) usa a coluna R;
- **FAB** usa a coluna M.

A coluna T mostra o ETp em uso. "-" = especificação não cadastrada, e nesse caso o Sigma fica "-".

As colunas K..P mostram as especificações dos **anos de referência do lote em análise** (o do Painel). **S2** é o ano
padrão de lote novo.

**Hematologia:** a biblioteca de 2025 recebeu os valores que estavam digitados na tabela. Os anos 2026–2029 ficam para o
laboratório preencher (ADR-075).

## 4b. Média/DP e limites de cada lote (aba Analitos, um bloco por lote — ADR-074)

- **Um bloco por lote do cadastro.** Lote e validade vêm da Configuração; lote novo ganha o bloco sozinho.
- **Abrir e recolher:**
  - ▼/▶ abre ou recolhe um bloco;
  - na linha 1, **Abrir todos** e **Recolher todos**;
  - **Lote em análise ▸** vai ao bloco do lote do Painel;
  - **Especificações ▸** abre a biblioteca.
- **Média/DP:** digite nas células amarelas (E:J) do bloco do lote. São valores, sem fórmula. Grava na hora, com trilha
  (a LotesStore é onde o sistema guarda e lê).
- **Valores de TESTE (15/10/2026):** os lotes que não tinham Média/DP receberam a **média dos próprios resultados** e
  **DP = 1,5 × o do lote em uso**, só para os limites aparecerem.
  - A LotesStore marca esses valores por nível (colunas TESTE_N1..N3).
  - O resumo do bloco avisa "contém valores de TESTE".
  - O Painel mostra **"(alvo TESTE)"** no status do nível. O veredito com alvo de teste não serve para liberar
    resultado.
  - **Substitua pelos valores corretos**: basta digitar por cima. A marca sai do nível quando a Média e o DP de
    teste foram trocados (trocar só um deles mantém o aviso).
  - Os limites (Média ± 1, 2 e 3 DP), o CV, o ET e o Sigma do lote aparecem ao lado.
  - O gráfico do Painel usa esses valores quando o lote é escolhido em H3.
  - **Lote sem Média/DP = gráfico sem limites e Westgard sem avaliar.**
  - O botão **Preencher vazios pelo CIQ (20 primeiros)** do bloco calcula só o que está vazio. Revise os valores.
- **Linhas do alvo antes da 1ª corrida (ADR-078):** com a Média/DP digitada, o gráfico mostra a média e as linhas
  −3DP..+3DP mesmo sem nenhuma corrida no período ou na base. Sem corrida, o eixo vai de 0 a 21 corridas. Confira o
  alvo antes da 1ª corrida do lote novo.
- **Status de cada nível no Painel (J7..):**

  | Status | O que quer dizer |
  |---|---|
  | OK / REJEITADO | veredito do Westgard no período |
  | SEM MÉDIA/DP | o nível não tem Média/DP no bloco do lote: sem linhas e sem Westgard |
  | AGUARDANDO 1ª CORRIDA | o lote ainda não tem resultado; as linhas mostram o alvo cadastrado |
  | SEM CORRIDA DESTE ANALITO | o lote tem resultados, mas não deste analito neste equipamento |
  | SEM CORRIDA NO PERÍODO | há corridas, mas nenhuma no período do filtro |
  | SEM PONTO NO PERÍODO | este nível não tem ponto no período |
  | DESATUALIZADO | o motor ainda está no lote, analito ou equipamento anterior (aguarde ou troque de novo) |
  | BASE SEM DADOS | a base de resultados está vazia ou não carregou: use ATUALIZAR DADOS |

- **Trocar de lote no Painel não grava nada:** só mostra. Quem grava é a digitação no bloco (ou na Especificações).
- **Filtro ajustado:** se o lote escolhido não tem resultado no filtro atual, o Painel ajusta o período, o analito ou,
  na Bioquímica, o equipamento. O aviso aparece em H4 e abaixo dos gráficos. Ao voltar ao lote em uso, o filtro de
  antes volta sozinho. Se você mexer no filtro, vale o que você escolheu.
- **Limitação:** a Média/DP é por lote × analito. Os dois DIMENSION usam o mesmo alvo do lote.
- **Média digitada longe dos pontos achata o eixo** de propósito: o eixo sempre mostra o alvo e os pontos.
- **Ano de referência:** abaixo da tabela de cada bloco ficam o **CLIA** (D:E) e a **Variação Biológica** (J:K).
  - O ano escolhe as especificações do bloco e, com elas, o ETp, o ET e o Sigma daquele lote.
  - Lote novo nasce com o ano padrão (Especificações S2).
  - A troca de ano fica na trilha de auditoria.
- **Registro:** cada gravação fica no `LotesStore` com a especificação do lote naquele momento (fonte, "CLIA aaaa · VB
  aaaa", ETp, CVTp), data, usuário e origem. Esse é o histórico de limites de cada lote.

## 5. Controle externo (CEQ) e Sigma (aba Estatística)

**Bias do CEQ.** O programa é escolhido em R4 na Hematologia e por analito em Especificações!AR na Bioquímica. CAP e
Controllab nunca se misturam. As opções:

| Ano | Rodada | Bias |
|---|---|---|
| qualquer | uma rodada da lista | só aquela rodada |
| um ano (ex. 2025) | TODAS | as rodadas daquele ano juntas (sem dado no ano = SEM EP) |
| TODOS | TODAS | as rodadas de todos os anos juntas |

A célula K5 diz o que entrou: rodadas e linhas.

**CV do Sigma.**
- É o **CV pooled** do CIQ: lotes com n ≥ 20, combinados pelos graus de liberdade.
- A janela padrão é de **6 meses**, terminando no último resultado. Ela fica na linha 11 da Estatística e pode ser
  mudada para 7, 8, 9 meses ou mais.

**Sigma.** Sigma = (ETp − |Bias|) / CV pooled. O "Sigma do plano" no Painel usa esse valor e avisa quando algum nível
ficou sem Sigma.

**Incerteza de medição** (Nordtest, ADR-067).
- Usa só o CIQ (12 meses) e o CEQ, sem certificado de calibrador.
- Sem CEQ suficiente para o analito, o U não é estimado, e a coluna "Situação" diz o motivo.

## 6. O que o laboratório mantém

- **CEQ:** digite cada rodada nas abas EQA.CAP_Dados / EQA.Controllab_Dados, com o **DP do grupo e o nº de
  laboratórios** de cada amostra. Sem esses dois campos, a incerteza não sai.
- **Cadastro:** ETp, CVI e CVG de analito novo e da biblioteca de cada ano (Bioquímica: aba Especificações);
  **Média/DP e anos de referência de cada lote novo** (seção 4b).
- **Investigação do CV:** revisar os IDs de `estudos/cv_suspeitos_<produto>.csv` (relatório
  `estudos/CV_INVESTIGACAO_2026-10-13.md`) e inativar os não conformes com motivo.

## 7. Itens que ficam com o laboratório (senhas não ficam com o desenvolvimento)

- **Senha do projeto VBA:** no Excel, Alt+F11 → Ferramentas → Propriedades de VBAProject → Proteção → "Bloquear projeto
  para exibição" e a senha. Salve e reabra.
- **Senha de proteção das abas:** a senha atual está no código público do repositório desde as versões anteriores.
  Para trocá-la:
  1. altere a constante `SENHA` em `instalar_integracao.py` e nos testes;
  2. faça uma entrega (`entregar.py`);
  3. não publique a senha nova no GitHub. Use uma variável de ambiente; o desenvolvimento pode preparar isso se for
     pedido.

## 7b. Registro de cada atualização (aba Sync_Log — ADR-079)

Cada ATUALIZAR DADOS deixa uma linha na aba **Sync_Log**: início e fim, resultado (OK, OK_SEM_MUDANCA ou ERRO),
quantos resultados foram lidos do DB_SEAC, quantos eram novos, total recebido, datas mínima e máxima do recebido,
erros/alertas do QA, o tempo de cada etapa (s) e, em falha, a etapa e o erro. Não há senha nem valor de resultado no
registro. Ficam as 2000 execuções mais recentes. A aba é só leitura.

## 8. Se algo der errado

- **ATUALIZAR DADOS falhou:** a mensagem diz a etapa. Até a etapa "validar antes de gravar", **nada foi gravado**: os
  gráficos e a Estatística continuam com a atualização anterior. A linha ERRO do Sync_Log guarda a mensagem.
- **"a Principal - Resultados cairia de X para Y linhas":** a atualização parou para não apagar resultados (trava do
  ADR-079: a tabela não pode encolher mais de 10% de uma vez). Confira as abas Digitar Resultados e Inativar. Se a
  redução for esperada, mude Cfg_Integracao **MOTOR_DADOS = PQ**, atualize uma vez e volte para **VBA**.
- **Conferir o cálculo (reconciliação) ou contornar uma falha do cálculo em VBA:** Cfg_Integracao
  **MOTOR_DADOS = PQ** faz o ATUALIZAR usar o Power Query, como antes (2 a 4 min). Os dois dão a mesma tabela. Volte
  para **VBA** depois.
- **Aba DB_ORGANIZADO (visão horizontal, só consulta):** deixou de ser refeita a cada ATUALIZAR. Para refazê-la:
  Alt+F8 → `mEtlSync.AtualizarOrganizado` → Executar (cerca de 1 min).

- **"DB_SEAC não encontrado" / ATUALIZAR DADOS falha:** veja o log diário em `AUTOMACAO_DB_SEAC\logs`. Desde
  13/10/2026, a automação restaura sozinha o último backup se o arquivo sumir; procure "RESTAURADO" no log.
- **Voltar uma versão do QC:** os backups de cada entrega ficam em `_backup_pre_integracao_<data>`, e os hashes em
  `MANIFESTO_PRODUCAO.md`.
- **Atualização sem rede (fora do laboratório):** MODO_FONTE = HISTORICO na aba Cfg_Integracao. O sistema reprocessa só
  o que já foi recebido. Volte para SEAC na rede.
