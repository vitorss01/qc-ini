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

## Próximas etapas
1. Se os testes e a revisão passarem: commit e entrega 39 completa, com todas as suítes. Previsão de 4 a 6 h, terminando hoje à noite.
2. ATUALIZAR na produção.
3. Usuário Simulado (reteste) e depois Tester VALIDAR.
4. Meta: o sistema validado e pronto para uso na segunda, 16/11.
