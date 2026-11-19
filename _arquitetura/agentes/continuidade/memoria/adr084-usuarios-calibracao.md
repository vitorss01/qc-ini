---
name: adr084-usuarios-calibracao
description: "ADR-084 (28/10/2026) — motivo livre, frmCarregando, status sem \"alvo TESTE\", calibração pelo ID, cadastro central + PBKDF2"
metadata:
  node_type: memory
  type: project
  originSessionId: cc5cf662-a93c-4a60-a8b5-23064798503c
  modified: 2026-10-29T09:01:59.688Z
---

Decisões do ADR-084 (pedido grande de 28/10/2026; prazo sexta 30/10):

- O usuário mandou tirar "(alvo TESTE)" do status: "pode deixar apenas média, não precisa demonstrar que é provisório". A marca continua na LotesStore. Risco avisado: o lote em uso 8974 tinha Média/DP de teste.
- As outras três perguntas ficaram sem resposta, então valeram os padrões recomendados:
  - login pelo central, com a cópia sincronizada (hash PBKDF2 copiado para a aba oculta Usuarios_Central);
  - assinatura numa pasta `assinaturas` ao lado do central, com o nome do arquivo na coluna E;
  - calibração marcada em todos os níveis.
- A sincronização do central é feita por VBA+ACE, não por Power Query: o PQ não roda com MOTOR_DADOS=VBA. Essa justificativa está no ADR.
- PBKDF2 com 310 mil iterações: 600 mil mediram 1,6 s por login nesta máquina.
- `InsertPictureInCell` falhou neste Excel. A imagem da assinatura é um shape `qcAssin_*`, redesenhado pelo login gravado no texto da linha.
- Navegação depois de inativar: NÃO reproduzida (o Painel fica). Pedir ao usuário o passo exato se ele insistir.
- Heredoc do bash com texto acentuado corrompe o Python (vira "�"). Usar Write/Edit para conteúdo não ASCII.

- Entregas 28 a 31 (28 e 29/10) pararam por testes que dependiam dos dados do dia: a primeira linha da tabela de inativação, a corrida só com o N1, as corridas contadas por equipamento, lote sem Média/DP. A saída foi rodar o resto das suítes ("cauda") numa cópia instalada antes de cada nova entrega.
- Dois defeitos reais achados em 29/10 (commit 67c7bf5):
  - colunas Q/R do Audit_Log: "3,924" virava 3924 → corrigido com formato "@";
  - manual que substitui o 2º de dois resultados no mesmo segundo renumerava o 1º → corrigido com o ITEM_ID do indicado na ordem, em mEtlCQ e no PQ.
- ENTREGUE em 29/10, na entrega 32: os dois QCs com o commit 67c7bf5, 24 suítes, 365 casos cada e 0 FAIL. O central também foi instalado (backup `_backup_pre_adr084_2026-10-29_0559` e pasta `assinaturas`). A coluna D está vazia nas 8 linhas e o laboratório ainda precisa fazer o primeiro acesso do ADMIN.

**Why:** reabrir isso sem contexto leva a refazer escolhas que o usuário já aceitou ou deixou no padrão.
**How to apply:** mantenha esses padrões; se o usuário responder às perguntas pendentes, ajuste. Ver [[manual-nao-substitui]], [[westgard-sequencia-adr083]], [[entrega-limpa-sozinha]].
