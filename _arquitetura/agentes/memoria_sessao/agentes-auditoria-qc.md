---
name: agentes-auditoria-qc
description: "Arquitetura de agentes do QC_INI criada em 31/10/2026 (usuario-simulado, especialista-cq, ux-ui, tester-qa + Orquestrador /auditoria-qc) e armadilhas de registro"
metadata:
  node_type: memory
  type: project
  originSessionId: cc5cf662-a93c-4a60-a8b5-23064798503c
  modified: 2026-10-31T18:57:31.968Z
---

Em 31/10/2026, a pedido do usuário, criei os agentes persistentes do QC_INI (commit 0a25f5a):
- fonte versionada em `qc-ini/.claude/{agents,skills/auditoria-qc,workflows/auditoria-qc.js}`;
- regras, registro e relatórios em `_arquitetura/agentes/`:
  - `REGRAS_COMUNS.md`;
  - `OCORRENCIAS.md` (só o Orquestrador escreve; IDs OC-xxxx);
  - `relatorios/<ciclo>/`;
  - `ferramentas/`.

O Orquestrador é a sessão principal, rodando a skill `/auditoria-qc`. Os subagentes não acionam outros agentes, não
pedem autorização ao usuário e não entregam.

**Why:** o usuário quer um ciclo testar → corrigir → retestar até validar. Fórmulas e regras só mudam com o parecer
do especialista-cq **e** a autorização do usuário. A validação é sempre feita por um agente diferente de quem corrigiu.

**How to apply:**
- A sessão abre em `QC_INI_PRONTO`, acima do repositório. Depois de editar `qc-ini/.claude`, rodar
  `python _arquitetura/agentes/instalar_agentes.py`, que copia para `QC_INI_PRONTO/.claude`.
- Agentes, skill e workflow novos só entram no registro na sessão seguinte:
  - o workflow roda por `scriptPath`;
  - a função `rodar` do workflow usa um agente comum como plano B quando o `agentType` falta.
- O `.gitignore` foi ajustado para versionar só `agents/`, `skills/` e `workflows/` de `.claude`.
- O ciclo 01 (os dois produtos, F1–F4) foi disparado em 31/10, por volta das 16h; o resultado deve ser registrado em
  `OCORRENCIAS.md`.

Ver [[adr086-087-telas]], [[entrega-limpa-sozinha]], [[excel-agendado-producao]].

**Estado em 03/11/2026, quando o limite de uso foi atingido:**
- O ciclo 01 está completo e commitado (`b98327d`): 50 OC registradas, 25/26 críticas e altas reproduzidas.
- Decisões do usuário em 03/11:
  - resultado ≤ 0 sai da estatística;
  - o BiasEQ usa o mesmo filtro do ViesEQ;
  - com H3 vazio, o Bias e o Sigma continuam visíveis, rotulados "do método";
  - o ANALISTA pode corrigir e excluir CEQ, com trilha;
  - o resto fica só com aviso; o formato do LJ não muda.
- O agente programador do lote 1 (ADR-088: OC-0001..0006, 0009, 0013) foi **interrompido no meio**. Pode haver
  edições parciais não commitadas em `src/`, `instalar_adr088.py`, `qa_adr088.py`, `painel.py` e `entregar.py`.
- **Antes de continuar:**
  1. rodar `git status`;
  2. revisar ou descartar essas edições;
  3. conferir se há EXCEL órfão (só cópias em `%TEMP%\qcw\adr088`);
  4. depois retomar o lote 1, validar (tester-qa + usuario-simulado), entregar e seguir para o lote 2 (CEQ:
     OC-0027..0030, 0032).
