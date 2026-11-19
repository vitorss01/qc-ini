---
name: lote-define-material
description: No relat209 (Bio) o material vem só do sufixo -EE do idAmostra digitado pelo técnico (inconsistente); o LOTE é o discriminador (ADR-091)
metadata:
  type: project
---

O relat209.csv do SIPEC (Bioquímica) não tem campo de material nem de lote. O DB_SEAC decompõe o `idAmostra`
`LLLLN-EE`:
- LLLL = lote;
- N = nível;
- EE = matriz e equipamento: 01/02 = SORO (Dimension 1/2), 03/04 = URINA (Dimension 1/2).

O técnico digita o EE errado com frequência nos lotes de urina:
- 6351: 248 SORO × 8 URINA;
- 6353: tudo SORO;
- 5659: 98 SORO × 33 URINA.

Também há lotes digitados com um dígito trocado, rodados uma única vez:
- 8794 e 7974, provavelmente 8974;
- 8177, provavelmente 8977;
- 8773 e 8970, provavelmente 8973;
- 8975, provavelmente 8974.

**Why:** em 05/11/2026 o usuário decidiu que quem define o material é o LOTE ("o que define se é urina é o lote") e
autorizou tomar a iniciativa das correções.

**How to apply:** a regra fica na `tblLotesMaterial` (ADR-091; carga em `config_produtos.py`):
- MATRIZ URINA para 635x e 565x;
- SO_ANALITOS para o 8353 (PCR);
- FORA_DO_CQ para os lotes digitados errado.

Lote novo de urina ou com erro de digitação vira uma linha nova na tabela, sem mudar código. Ver
[[agentes-auditoria-qc]].
