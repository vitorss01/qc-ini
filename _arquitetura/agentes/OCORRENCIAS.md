# Registro de ocorrências da auditoria do QC_INI

Mantido só pelo Orquestrador (`/auditoria-qc`). Uma linha por ocorrência e IDs que nunca se repetem. A evidência fica
no relatório do ciclo indicado.

**Status possíveis:**

1. ABERTA
2. EM ANÁLISE
3. AGUARDA AUTORIZAÇÃO
4. EM CORREÇÃO
5. AGUARDA RETESTE
6. VALIDADA
7. ENCERRADA SEM CORREÇÃO

**Prioridade:** Crítico, Alto, Médio ou Baixo (REGRAS_COMUNS §4).

| ID | Ciclo | Produto | Módulo | Prior. | Descrição | Evidência | Analisa | Corrige | Status | Reteste | Validação final |
|---|---|---|---|---|---|---|---|---|---|---|---|

## Histórico dos ciclos

| Ciclo | Modo | Produtos | Frentes | Ocorrências novas | Resultado |
|---|---|---|---|---|---|
| 2026-10-31_ensaio | ensaio (sem Excel) | Bioquimica | F1–F4 | 0 (4 fictícias [ENSAIO], descartadas) | Ligação conferida: 4 frentes, parecer consolidado, roteamento CQ/UX/Tester correto; agentes fora do registro rodaram pela definição (8/8) |
