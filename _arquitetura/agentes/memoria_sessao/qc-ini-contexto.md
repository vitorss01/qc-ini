---
name: qc-ini-contexto
description: "Onde fica o QC_INI, como os .xlsm são construídos e as armadilhas de automação Excel já conhecidas"
metadata:
  node_type: memory
  type: project
  originSessionId: cc5cf662-a93c-4a60-a8b5-23064798503c
  modified: 2026-10-01T19:11:18.575Z
---

Repo clonado em `QC_INI_PRONTO\qc-ini` (remoto github vitorss01/qc-ini, público). Sistema de CQ laboratorial em Excel/VBA; .xlsm na raiz são artefatos, fontes/instaladores em `_arquitetura/` (etapa mais recente: `etapa1_multilote/`, helper COM `xlh.py`).
DB_SEAC.xlsm é Power Query sobre CSVs do SIPEC (relat207 Hema, relat209 Bioq) com janela móvel de 12 meses; o CSV tem coluna `id` única por resultado.

Armadilhas documentadas no projeto: VBA compila sob demanda (chamar rotina nova antes de salvar); `aS`/`aM` proibidos; Excel guarda cor em BGR; formatar elementos de gráfico abre modal não suprimível; matar Excel e limpar Resiliency antes de rodar automação; OneDrive AutoSave ativo na pasta.

**Why:** evita redescobrir o contexto a cada sessão.
**How to apply:** ler `_arquitetura/QUALITY_GATE.md`, `ARCHITECTURE_DECISIONS.md` (ADR até 056) antes de mudar arquitetura. Ver [[integracao-db-seac]].
