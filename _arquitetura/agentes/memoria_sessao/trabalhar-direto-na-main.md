---
name: trabalhar-direto-na-main
description: Usuário quer todo trabalho do QC_INI direto na main (produção) — sem branches de teste
metadata:
  node_type: memory
  type: feedback
  originSessionId: cc5cf662-a93c-4a60-a8b5-23064798503c
  modified: 2026-10-02T19:26:31.695Z
---

A partir de 02/10/2026 o usuário pediu: "só trabalhe em produção a partir de agora — nada mais em branchs de testes — só produção". A linha `integracao-db-seac` foi incorporada à `main` (fast-forward, commit 0b58219) e apagada.

**Why:** o usuário considera a main o produto; branches paralelas viraram ruído.
**How to apply:** commitar direto na `main` do repo `qc-ini`; testar sempre em CÓPIAS dos .xlsm (scratchpad), nunca em branch; manter backup sha256 antes de instalar nos arquivos de produção. Push: ele pediu commit + push nesta entrega; antes de cada push conferir segredos (o remoto é público — ex.: a senha do DB_SEAC e o caminho de rede interno do laboratório (servidor/compartilhamento do DB_SEAC) nunca vão para o repo; o caminho entra só por `entregar.py --caminho`). Ver [[integracao-db-seac]], [[qc-ini-contexto]].
