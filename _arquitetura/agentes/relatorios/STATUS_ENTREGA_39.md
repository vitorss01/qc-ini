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

## Próximas etapas (novos horários)
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
