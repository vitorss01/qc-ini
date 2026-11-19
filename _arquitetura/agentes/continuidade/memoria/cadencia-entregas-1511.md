---
name: cadencia-entregas-1511
description: "Decisões do usuário em 15/11/2026 sobre ritmo de correção, suítes direcionadas e validação só pelo Tester em baixo/médio"
metadata:
  node_type: memory
  type: feedback
  originSessionId: cc5cf662-a93c-4a60-a8b5-23064798503c
  modified: 2026-11-15T16:58:38.796Z
---

Decisões do usuário em 15/11/2026 (pergunta "como ser mais ágil"):
- Começa a usar o QC só depois da entrega 43 validada (não quis usar a 42 com contorno).
- Cadência: MANTER — todo achado é corrigido e entregue na hora (não quis lote semanal para médio/baixo).
- Entrega de correção: pode rodar só as suítes afetadas + essenciais (seguranca, etl, sync, final...) ≈ 2–3 h em vez das 35 (~8 h); as 35 completas rodam periodicamente (ex.: numa cópia, à noite, sem bloquear).
- Validação: itens BAIXO e MÉDIO só pelo Tester; CRÍTICO e ALTO continuam Usuário Simulado RESOLVIDA + Tester VALIDADA. Nunca valida quem implementou.

**Why:** cada achado disparava ~8 h de entrega + ~3,5 h de reteste; o usuário quer terminar mais rápido sem perder a verificação independente.

**How to apply:** em entregar.py usar o modo de suítes direcionadas para correções; rotear o reteste pela prioridade (baixo/médio → só tester-qa VALIDAR). Ver [[meta-pronto-semana-1611]] e [[previsao-em-toda-mensagem]].
