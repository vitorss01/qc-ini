---
name: previsao-em-toda-mensagem
description: toda mensagem ao usuário termina com previsão de término (horário) e duração aproximada
metadata:
  node_type: memory
  type: feedback
  originSessionId: cc5cf662-a93c-4a60-a8b5-23064798503c
  modified: 2026-11-14T13:28:05.131Z
---

Em TODAS as mensagens ao usuário (chat e status pelo celular), terminar com uma linha de previsão: horário de término do que está em andamento e quanto tempo ainda vai durar aproximadamente (ex.: "Previsão: termina por volta das 14h30, daqui a ~4h").

**Why:** pedido explícito do usuário em 14/11/2026, depois de perguntar várias vezes "quando termina"; ele acompanha pelo celular. Previsões anteriores atrasaram (Hema estimada com o tempo da Bio).

**How to apply:** estimar a partir das durações medidas de entregas anteriores (pastas _entrega_*, horários dos .json) — Hema ≈ 1,28 × Bio; ATUALIZAR ~5 min por produto. Se a previsão mudar, dizer quanto e por quê. Ver [[meta-pronto-semana-1611]].
