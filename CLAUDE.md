# QC_INI — instruções permanentes do projeto

Sistema de controle estatístico da qualidade do laboratório clínico (Bioquímica e Hematologia; Excel/VBA + Power Query),
de uso diário. Repositório público.

## Fase atual: FECHAMENTO DE VERSÃO (feature freeze desde 23/11/2026)

- **Não implementar funcionalidade nova**, não redesenhar módulo que funciona, não refatorar por preferência.
- Só entram: bugs que comprometem o funcionamento; erro de cálculo estatístico/regra de CQ; integridade de dados;
  falha de automação/integração/importação; segurança/permissões; erro que impede uma operação necessária; problema
  crítico de interface.
- Toda melhoria não essencial vai para `BACKLOG_POS_RELEASE.md` — não é implementada agora.
- **Se está funcionando, foi testado e não tem problema real, não mexa.**

## RELEASE MANAGER (obrigatório)

Use o agente **`qc-release-manager`** (`.claude/agents/`) e a skill **`/qc-release`** (`.claude/skills/qc-release/`)
sempre que o trabalho envolver: finalização, estabilização, correção de bugs, auditoria de release, testes de regressão,
prontidão para entrega, ou qualquer mudança/ampliação de escopo.

- Antes de alterar algo, o problema é classificado (P0 crítico · P1 alto · P2 moderado · P3 melhoria) **com evidência**.
  P0/P1 corrige; P2 só se objetivo e de baixo risco; P3 vai para o backlog.
- Ciclo: diagnóstico consolidado (uma vez) → classificação → plano → correções → testes direcionados → **uma**
  validação integrada → relatório. Não reiniciar auditoria completa a cada correção; não reabrir componente aprovado sem
  evidência de que mudança posterior o afetou.
- **Início de toda sessão:** ler `QC_RELEASE_STATUS.md` e continuar da fase registrada. **Fim de sessão:** atualizar o
  `QC_RELEASE_STATUS.md` (e o backlog).

## Regras que sempre valem

- Repositório público: nada de senha, caminho de rede interno (DB_SEAC), `.xlsm`, `~$*`, CSV/JSON com dados do
  laboratório. Conferir o diff antes de cada push. Trabalhar direto na `main`.
- Teste sempre em **cópia**; produção só por `entregar.py`, `aplicar_atualizar.py`, `trocar_producao.py` ou ferramenta com
  backup + sha. Um Excel por vez; matar só o PID que o script criou; nunca capturar a tela (dados de paciente).
- Não alterar fórmula, regra ou critério estatístico sem autorização explícita do usuário (com ADR e oráculo independente).
- Falha de ambiente (Excel/Power Query caindo, sem DB_SEAC em casa) não é defeito do produto: rodar a suíte sozinha de novo.
- Mensagens ao usuário em português, terminando com a previsão de término e a duração aproximada.
- Contexto técnico: `_arquitetura/agentes/continuidade/README.md`, `CONTINUAR_EM_CASA.md`, ADRs em
  `_arquitetura/integracao_dbseac/`, `GUIA_DE_USO.md`, `_arquitetura/agentes/OCORRENCIAS.md`.
