---
name: qc-release
description: Procedimento padrão de FECHAMENTO de versão do QC_INI (feature freeze): auditoria final consolidada, classificação P0–P3 com evidência, plano mínimo, correções, testes direcionados, uma validação integrada, relatório de release e atualização do QC_RELEASE_STATUS.md e do BACKLOG_POS_RELEASE.md. Use quando o usuário pedir para finalizar, estabilizar, entregar, liberar, fazer release, avaliar prontidão, corrigir bugs para a entrega ou decidir se uma mudança entra na versão.
---

# QC_INI — procedimento de release (`/qc-release`)

Quem conduz é o agente **`qc-release-manager`** (`.claude/agents/qc-release-manager.md`). A sessão principal implementa;
o release manager classifica, decide o que entra e confere as evidências.

## 0. Retomar sem refazer

1. Ler `QC_RELEASE_STATUS.md` e `BACKLOG_POS_RELEASE.md` (raiz do repositório).
2. Se a fase registrada não for "diagnóstico", **continuar dela**. Não reabrir componente marcado como testado sem
   evidência de que uma mudança posterior o afetou.

## 1. Diagnóstico consolidado (uma vez por ciclo)

Acione `qc-release-manager` com: o status, a última avaliação de prontidão (`_arquitetura/agentes/relatorios/PRONTIDAO_*`),
as ocorrências abertas (`_arquitetura/agentes/OCORRENCIAS.md`) e os resultados de suíte mais recentes. Ele devolve a
tabela de problemas com evidência e prioridade. **Sem evidência, não é problema.**

## 2–3. Classificação e plano

- P0/P1: entram. P2: só se a correção for objetiva, de baixo risco e sem mudança estrutural. P3: backlog.
- Falha de ambiente (Excel/PQ caindo, sem DB_SEAC, dados antigos da cópia) não é defeito: rodar a suíte de novo sozinha.
- Plano = lista curta de correções mínimas, cada uma com o teste que a prova.

## 4. Correções (Claude principal)

- Pequenas, isoladas, reversíveis; ADR curto; sem refatoração; sem tocar cálculo estatístico validado sem autorização
  explícita do usuário e oráculo independente.
- Antes de abrir o Excel: `python _arquitetura/integracao_dbseac/estudos/lint_vba.py`.
- Sempre em cópia; um Excel por vez; matar só o PID que o script criou; nunca capturar a tela.

## 5. Testes direcionados

`_arquitetura/integracao_dbseac/testes/rodar_suites.py <produto> <copia_instalada.xlsm> <saida> --suites a.py,b.py`
(uma suíte por vez, cada uma numa cópia nova). Rodar as suítes do que mudou e das dependências diretas.

## 6. Validação integrada final (uma)

No PC do trabalho, com a produção atual: `python _arquitetura/integracao_dbseac/entregar.py Hematologia Bioquimica
--direcionada` (NUCLEO + SUITES_DO_LOTE). Produção só é trocada se tudo passar. Ações de dados (ex.: alvos do ADR-100)
pela ferramenta própria, com backup e sha.

## 7. Relatório de release

1. O que foi corrigido · 2. testes executados · 3. que passaram · 4. que não puderam ser feitos e por quê ·
5. problemas críticos remanescentes · 6. melhorias adiadas (backlog) · 7. status final · 8. versão recomendada.

Gate atendido → **"QC_INI — RELEASE CANDIDATE APROVADO PARA ENTREGA"** (release técnico; a aprovação do responsável do
laboratório é etapa humana documentada). Não declarar nada sem evidência.

## 8. Fechar a sessão

Atualizar `QC_RELEASE_STATUS.md` (fase, critérios, testados, correções, pendências P0/P1/P2, testes que faltam, backlog,
próxima ação) e `BACKLOG_POS_RELEASE.md` (melhoria, componente, justificativa, benefício, prioridade futura, risco).
