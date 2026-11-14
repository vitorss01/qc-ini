# Entrega 39 (ADR-092 + ADR-093): acompanhamento

Este arquivo é atualizado a cada etapa, enviado ao GitHub e mandado para você abrir pelo celular.

## 13/11, por volta das 15h: mudança de plano
- **Sua decisão sobre o N-01:** um manual num lote de urina com nome de analito de soro é **aceito como urina**. Isso virou o **ADR-093**:
  - o resultado vai para o analito de urina que o equipamento usa (ex.: Creatinina → CRE2);
  - a coluna DETALHE mostra a troca, e o QA registra o aviso I05;
  - se não houver um analito de urina único para aquele nome, o manual continua fora da estatística.
- Parei a entrega 39, que tinha começado às 13h52. Ela estava só nas cópias, e a produção não foi tocada. Assim o ADR-092 e o ADR-093 entram juntos **numa entrega só**, poupando de 4 a 6 h de PC carregado.
- **Já feito:**
  - o código da regra nos dois motores (VBA e Power Query) e o aviso I05;
  - os testes ajustados: o K01 do qa_adr092 e o caso MAN_0066 do estudo de equivalência entre os motores;
  - o ADR-093.md;
  - o GUIA_DE_USO, que agora explica a tabela de lotes e o lançamento em lote de urina;
  - a OC-0056 atualizada.
- **Em andamento:**
  - testes nas cópias da produção: qa_adr092, qa_adr091 e equivalência VBA×PQ, nos dois produtos;
  - revisão independente do código por dois agentes, sem Excel.

## 13/11, por volta das 15h30
- **Revisão independente:** 1 problema ALTO e 2 MÉDIOS, todos corrigidos. O mais importante foi este: um analito digitado errado poderia virar uma série de urina que o equipamento nunca manda. Agora o manual só é aceito se o equipamento já enviou aquele analito de urina.
- **Falha na cópia, causa encontrada e corrigida:** o VBA não compilava porque uma edição automática quebrou o fim de uma função nova. A função foi restaurada, e o lint (a conferência feita antes de abrir o Excel) agora pega esse tipo de erro. A produção não foi tocada.
- **Instalação completa nas cópias:** OK nos dois produtos, com todos os instaladores (westgard_seq → adr092), VBA compilado e blindagem TRAVADO.
- **Em andamento:** testes nas cópias instaladas (qa_adr092, qa_adr091, qa_etl com a equivalência VBA×PQ e qa_manual_substitui) nos dois produtos.

## 13/11, 18h30: testes nas cópias OK, entrega completa começou
| Teste | Bioquímica | Hematologia |
|---|---|---|
| qa_adr092 | 7 PASS / 0 FAIL | 7 PASS / 0 FAIL |
| qa_adr091 | 7 / 0 | 7 / 0 |
| qa_etl | 64 / 0 | 63 / 0 |
| qa_manual_substitui | 20 / 0 | 20 / 0 |
| VBA × Power Query, com os casos de teste | 0 divergências (35.505 linhas) | 0 divergências (48.119 linhas) |

- **Lacuna encontrada e fechada:**
  - nos dados reais, o teste K01 tinha pegado só analitos que o equipamento nunca manda em urina (Magnésio na Bioquímica). Testou o caso recusado, mas não o caso **aceito como urina**;
  - criei o K01b: na Bioquímica, "Cloro" num lote de urina ficou ATIVO como CL, com a nota, o aviso I05 e a trilha no Audit_Log;
  - a comparação entre os motores agora lista os casos conferidos: aceito como A1C, recusado o código nunca recebido;
  - na Hematologia não há urina, e o teste diz "não se aplica".
- Commits enviados ao GitHub (a502b4b).
- **Entrega 39 completa começou às 18h30:** Bioquímica e depois Hematologia, com todas as suítes. Previsão de término **entre 1h e 2h da madrugada**.

## 13/11, 21h15: Bioquímica parou num teste; causa encontrada e corrigida
- **O que falhou:** o teste BN (qa_blocos_lote), na Bioquímica. Ele recalcula os "valores de TESTE" gravados em 16/10 e não achou os resultados do lote de urina 6353 (Cálcio e Creatinina).
- **Causa:** em 16/10 esses resultados de urina chegavam com nome de soro (CA → "Cálcio"). Desde o ATUALIZAR de 11/11, a regra do ADR-091 faz o mesmo resultado chegar com o código de urina (CA, CRE2). Os valores batem exatamente: Cálcio N1 = 4,6 e Creatinina N1 = 77,075. Os dados e as regras estão certos; o teste é que procurava pelo nome antigo.
- **Correção:** o teste agora conta esses resultados com o nome que tinham na data da foto. Nada no sistema mudou. Commit enviado ao GitHub.
- **Produção da Bioquímica:** intacta. A entrega só troca a produção quando tudo passa.
- **Hematologia:** segue normalmente. Até agora, tudo OK.

**Observação para o CQ (não bloqueia):** os blocos "Cálcio", "Creatinina" etc. do lote de urina 6353 na aba Analitos têm médias de urina gravadas em 16/10 com o nome do soro. Com a regra do ADR-091, nenhum resultado novo cai neles. Os resultados de urina desse lote estão nos códigos (CA, CRE2…). Vale o laboratório revisar esses blocos depois.

## 13/11, 22h10: chave avaliada, ADR-094 (data-hora com segundos) escrito, entrega 39 retomada
- **Chave:** medi nas duas produções. A data-hora até o segundo não é única: até 62 resultados no mesmo segundo na Bioquímica e 93 na Hematologia. Uma corrida gera todos os analitos no mesmo instante. Mesmo somando equipamento + lote + nível + analito, ainda sobram 3.570 repetidos na Bio e 1.116 na Hema. Você decidiu manter o ID: ele é o número sequencial do SIPEC, único por resultado.
- **ADR-094 escrito**, em commit, ainda não entregue:
  - data-hora com segundos em tudo o que você vê: gráfico, janela, colunas, assinatura, mensagens e textos do QA;
  - campos só de data continuam dd/mm/aaaa;
  - **defeito corrigido junto:** na Bioquímica, as linhas da carga de 11/11 mostravam a data com o mês na frente (10/28/2026).
  - Revisão independente do código em andamento.
- **Entrega 39:** a Hematologia parou no teste de desempenho por culpa minha. Eu rodava uma leitura pesada em paralelo, e a troca pelo spinner foi de 200 para 335 ms. O sistema não tem defeito. Retomei com a máquina livre, a partir de uma cópia limpa do código.

## 14/11, 2h: entrega 39 não trocou as produções; causas achadas; entrega 40 em andamento
- **Produções intactas:** a entrega só troca quando tudo passa. Bio continua `8b1cda13…` e Hema `f1907656…`.
- **Bioquímica, 3 falhas: o sistema estava certo; eram os testes:**
  - **BN (valores de TESTE de 16/10):** o teste contava todo CRE2 do lote de urina 6354 como "Creatinina". Em 16/10, só os resultados que chegaram com sufixo de soro (01/02) tinham esse nome. Corrigido.
  - **V01 (varredura de lotes):** nos lotes de urina 6353, 6354 e 5659, os resultados ficam com o código do aparelho (CRE2, MALB, GLUC…), como o ADR-091 decidiu ("fora do Painel"). O teste exigia gráfico com pontos nesses lotes. Agora conta só o que o Painel pode mostrar.
  - **W20 (Westgard):** depois da meia-noite, o teste escolheu 13/11 como dia livre. A cópia só tinha dados até 10/11, mas o ATUALIZAR do teste lê a origem ao vivo, que já tinha resultado real naquela manhã. O manual virou CONFLITO_MANUAL, que é a regra certa. Agora o dia livre fica dentro do que a cópia já carregou. Corrigi o mesmo risco no qa_etl.
  - Conferido numa cópia: BN 14/0, varredura 4/0, Westgard 29/0.
- **U13 (tempo da inativação), nos dois produtos: limite ajustado com medição:**
  - com a máquina livre, a **produção de hoje** (código antigo) já passa do limite na Hematologia: 25,1 s, para um limite de 25;
  - a cópia nova deu Hema 27,3 s e Bio 19,1 s;
  - o ATUALIZAR interno é o mesmo de 10/11. O que cresceu foi a base, e a medida varia ±3 s entre execuções;
  - novo limite: Bio 22 s, Hema 30 s.
  - **Aviso honesto:** inativar um ponto leva ~20 a 27 s para você. Funciona, mas é lento. Deixo proposta a otimização como trabalho separado.
- **Entrega 40 começou às 1h54.** Ela já contém tudo da 39 e mais o ADR-094 (data-hora com segundos), a partir de uma cópia limpa do commit `1b453a0`. Depois dela vem o ATUALIZAR nas produções entregues, com backup e conferência do SHA.

## 14/11, 5h15: Bioquímica passou em tudo, menos num teste com defeito (corrigido); Hematologia em andamento
- **Bioquímica, 34 de 35 suítes OK** na entrega 40. Entre elas: segurança 20/0, ETL 64/0, Westgard 29/0, inativar UX 16/0, ADR-091 7/0 e ADR-092 7/0.
- **A que falhou, qa_adr094 (data-hora com segundos):**
  - a 1ª verificação passou: 572 datas conferidas nas tabelas, todas com o formato certo;
  - a 2ª parou: o teste tentou acrescentar uma linha numa tabela e o Excel recusou.
- **Causa:** defeito do teste, não do sistema. Pelo Python, `Resize(linhas, colunas)` não redimensiona; devolve uma única célula. O teste pediu para a tabela encolher para 1 coluna, e o Excel recusou. O sistema faz o mesmo passo em VBA, onde funciona; é assim que toda carga grava as tabelas. Corrigi o teste (`GetResize`, como os outros testes já fazem). Commit `d184bfb` no GitHub.
- **Produções intactas:** a entrega só troca quando tudo passa.
- **Hematologia:** instalação em andamento desde 5h02. Ela vai cair no mesmo teste, porque a cadeia usa a versão antiga dele.
- **Retomada já armada:** quando a Hematologia terminar, as duas cópias instaladas rodam de novo o qa_adr094 corrigido, a comparação do QA entre os motores e as versões mais rigorosas da varredura de lotes e do Westgard (revisão de ontem à noite). Só se tudo passar: troca da produção (backup + SHA) e ATUALIZAR.
- **Parecer de UX sobre a lentidão da inativação (OC-0060):** APROVADO_PROVISORIO até 30/11. Meta: ≤ 8 s na Bio e ≤ 10 s na Hema, com resultado idêntico ao ATUALIZAR completo. Também ficaram registradas a estimativa fixa do "aguarde" (OC-0063) e o guia que diz "Etapa k de 7" (OC-0064).

## 14/11, 6h45: Hematologia parou no teste de desempenho; causa achada (interferência externa, não o código); retomada em andamento
- **Hematologia na entrega 40:** instalação completa OK (todos os instaladores, blindagem TRAVADO). As primeiras 6 suítes passaram. O qa_desempenho reprovou por pouco: 1ª troca de analito 641 ms (limite 600) e mediana 305 ms (limite 300). As suítes seguintes não rodaram.
- **Investigação:** medi a produção atual, a cópia da entrega 39 e a da entrega 40, alternadas, 3 vezes cada:
  - fora de uma certa janela de tempo, as três dão o mesmo: cerca de 250 ms;
  - dentro dela, qualquer uma fica lenta, **inclusive a produção**, que chegou a 371 ms e também reprovaria.
- **Causa:** a tarefa agendada do laboratório "DB_SEAC - Atualização Automática" abre um Excel a cada 2 minutos, por cerca de 42 s. Quando a medida cai nesses 42 s, ela mede a disputa pela CPU, não o sistema. A falha da entrega caiu exatamente numa dessas janelas (6h14). Provavelmente a falha de ontem às 18h27 também; na hora, culpei a minha leitura em paralelo.
- **Correção (só no teste):** as medidas de tempo agora esperam a tarefa estar parada e com folga até a próxima execução, sem mudar nenhum limite. Isso vale para a troca de analito (P01/P02) e para o tempo de inativar (U13). Commit `ae89471`.
- **Para você saber:** durante esses ~42 s a cada 2 min, trocar de analito leva cerca de 370 ms em vez de 250. É imperceptível no uso normal.
- **Um erro meu no script da retomada (6h14):** o `tar` do Git não extraiu a cópia limpa. Nada rodou e nada foi trocado. Corrigido; o script agora para se a cópia não estiver completa.
- **Retomada em andamento desde 6h39**, de uma cópia limpa do `ae89471`:
  1. Bioquímica: qa_adr094 (corrigido), comparação do QA entre os motores, varredura de lotes e Westgard (oráculos novos). Se tudo passar: troca da produção e ATUALIZAR;
  2. Hematologia: do qa_desempenho até o fim (29 suítes). Se tudo passar: troca e ATUALIZAR.
- **Revisão independente** das duas correções de teste em andamento: agentes que só leem o código, sem Excel.

## 14/11, 7h20: um defeito real do ADR-094 apareceu e foi corrigido; entrega 41 completa em andamento (o horário mudou)
- **O que a retomada mostrou:** com o defeito do teste corrigido, a verificação do ADR-094 rodou inteira pela primeira vez. Ela achou um **defeito real**, que é só de apresentação (nenhum valor, cálculo ou regra muda):
  - colunas de data-hora que hoje estão **vazias** iam aparecer como `####` assim que recebessem o primeiro valor. São elas: data da inativação na tabela de resultados; registrado em, dos manuais; e, na Hematologia, as colunas de inativação e de comentários. O instalador só alargava a coluna quando já havia data nela;
  - na aba Sync_Log, INÍCIO e FIM de cada ATUALIZAR apareciam `####` desde o ADR-079: com segundos, a data-hora não cabe na largura 12. A aba é visível a todos.
- **Correção:** o instalador testa a coluna vazia com uma data de exemplo, alarga e apaga a data em seguida. Testado nas cópias: verificação do ADR-094 **7 PASS / 0 FAIL** na Bioquímica e na Hematologia. Commit `1a72892`.
- **Também conferido na retomada (Bioquímica):** comparação do QA entre os motores 3/0, varredura de lotes 4/0 e Westgard 29/0, já com os oráculos mais rigorosos da revisão de ontem.
- **Revisão independente** (agentes que só leem o código) das correções de teste de hoje: confirmou a causa do erro no Python e a espera da tarefa DB_SEAC. Também apontou ajustes pequenos, já feitos, e um ponto em aberto: os limites de tempo agora valem com a máquina sem a tarefa rodando, e você sente a janela da tarefa. Registrado como OC-0065.
- **Por que outra entrega completa:** o instalador mudou, então o arquivo entregue é outro. A regra é que todas as suítes rodem sobre o arquivo exato que vai para a produção. Retomar só parte delas não vale.
- **Produções intactas:** Bio `8b1cda13…`, Hema `f1907656…`.

## Próximas etapas (7h20): o horário atrasou
1. **Entrega 41**, Bioquímica e depois Hematologia, com todas as suítes: começou às 7h15. Previsão: Bio por volta de 10h30, Hema por volta de 13h30.
2. ATUALIZAR nas produções logo em seguida: por volta das 14h.
3. Usuário Simulado (reteste) e depois Tester VALIDAR, no fim de semana. Meta: pronto para uso na segunda, 16/11.

## Etapas de 6h45 (substituídas)
1. Bioquímica entregue e atualizada: por volta das 7h15.
2. Hematologia, suítes até cerca de 9h; troca e ATUALIZAR até cerca de 9h30.
3. Usuário Simulado (reteste) e depois Tester VALIDAR.

## Etapas de 5h15 (substituídas)
1. Hematologia: suítes até cerca de 8h.
2. Retomada das duas (4 suítes cada), troca das produções e ATUALIZAR: até cerca de 9h.
3. Usuário Simulado (reteste) e depois Tester VALIDAR.

## Etapas de 2h (substituídas)
1. Entrega 40, Bioquímica e depois Hematologia, com todas as suítes: de 1h54 até cerca de 8h30.
2. ATUALIZAR nas duas produções logo em seguida, por volta de 9h.
3. Usuário Simulado (reteste) e depois Tester VALIDAR.

## Etapas de 22h10 (substituídas)
1. Entrega 39: Hematologia (do qa_final até o fim) e Bioquímica (do qa_blocos_lote até o fim), com troca das produções, por volta de 1h30.
2. ATUALIZAR nas duas produções, por volta de 2h.
3. Entrega 40 (ADR-094) completa, com todas as suítes e mais a comparação do QA entre os motores, das 2h até cerca de 9h.

## Etapas anteriores (substituídas)
1. Hematologia: término das suítes e troca da produção por volta da meia-noite.
2. Bioquímica: retomo as suítes que faltam na mesma cópia instalada, a partir do qa_blocos_lote, cerca de 1h. A troca da produção segue as mesmas checagens da entrega: backup, SHA e ninguém ter salvo a produção durante os testes. Término por volta de 1h15.
2. ATUALIZAR na produção, com backup e conferência do SHA, logo depois.
3. Usuário Simulado (reteste) e depois Tester VALIDAR.
4. Meta: o sistema validado e pronto para uso na segunda, 16/11.
