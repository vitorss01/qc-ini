---
name: integracao-db-seac
description: "Reestruturação 2026-10 do QC_INI: DB_SEAC -> Power Query -> Principal - Resultados (tblCQ_Final) como fonte única; decisões do usuário e onde está o código"
metadata:
  node_type: memory
  type: project
  originSessionId: cc5cf662-a93c-4a60-a8b5-23064798503c
  modified: 2026-10-02T15:13:34.301Z
---

Etapa iniciada em 01/10/2026; entregue e enviada à `main` em 02/10/2026 (commit 0b58219; checkpoint: tag `checkpoint-pre-integracao-dbseac`). Decisão registrada como ADR-057 (`_arquitetura/integracao_dbseac/ADR-057.md`, relatório `RELATORIO_ADR057.md`).

Código: `_arquitetura/integracao_dbseac/` — `pq/*.m` (consultas), `camada_dados.py` (abas/tabelas/consultas), `patch_vba.py` (gera VBA de src/<bio|hema> a partir do próprio .xlsm, com conferência de referência residual e de sombreamento de nome), `instalar_integracao.py` (instala tudo, compila VBA com cão de guarda), `instalar_producao_integracao.py` (backup sha256 + instala + rollback), `testes/qa_final.py` (26 testes PASS/FAIL).

Nomes pedidos pelo usuário (02/10/2026): abas e botões **DB** (recebimento), **Inativar** (coluna checkbox "REGISTRAR - LJ", padrão marcado = X vermelho), **Digitar Resultados** (IDs MAN_0001…), **Principal - Resultados** (tblCQ_Final). Mais DB_ORGANIZADO, COMENTARIOS_TECNICOS, QA_INTEGRACAO, Cfg_Integracao.

Decisões do usuário (01/10/2026):
- Histórico antigo do DB_Resultados é IGNORADO; abas sem função podem ser apagadas (após provar ausência de dependência). Removidas: Importar, Resultados, EQC_Dados (LEGADO ADR-034), DB_Resultados, Cfg_Status, Rep 1-3 de Registros. EQA.* preservadas.
- Bioquímica: separar por equipamento (DIMENSION 1 / 2), seletor Painel!L4 `selEquipamento`, padrão D1.
- Cópia do DB_SEAC.xlsm levada à v2.4 (colunas ID_ORIGEM/ID_AMOSTRA/ITEM_ID/UNIDADE). O DB_SEAC de PRODUÇÃO (rede, atualizado pela tarefa agendada) ainda é v2.3 — precisa do patch antes de apontar CAMINHO_DB_SEAC para ele.

Lições: o VBA não diferencia maiúsculas (variável local escondendo função = erro 450, que abre o VBE na tela do usuário mesmo em Excel invisível); Names.Add com função falha neste Excel pt-BR (usar célula auxiliar); `[@Col]` é recusado via COM (usar `tbl[[#This Row],[Col]]`); NumberFormat "General" recusado via COM.

**Why:** usuário quer fonte única de verdade e continuidade quando o interfaceamento falha.
**How to apply:** trabalho direto na main (ver [[trabalhar-direto-na-main]]); o remoto do repo é público (INTERFACEAMENTO_DB_COPIA/ fica no .gitignore). Ver [[qc-ini-contexto]], [[excel-agendado-producao]].
