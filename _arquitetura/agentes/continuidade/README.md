# Continuidade: tudo o que só existia na máquina do trabalho (19/11/2026)

O código, os testes, os ADRs, o GUIA, o manifesto, as ocorrências e os relatórios **já estavam no GitHub**. Esta pasta
traz o resto, sanitizado para repositório público: o caminho do DB_SEAC virou `<CAMINHO_DB_SEAC>`, a pasta do
repositório virou `<QC_RAIZ>` e o perfil do usuário virou `%USERPROFILE%`. A senha não está aqui.

| Pasta / arquivo | O que é | Como usar em casa |
|---|---|---|
| `memoria/` | As anotações de contexto das sessões do Claude Code: regras combinadas, armadilhas, decisões do usuário. `MEMORY.md` é o índice. | Copie os `.md` para `%USERPROFILE%\.claude\projects\<pasta do projeto>\memory\`. A pasta é criada na 1ª sessão do Claude Code no repositório. Ou peça ao Claude: "leia `_arquitetura/agentes/continuidade/memoria/` antes de começar". |
| `scripts/entregar45.ps1` | Modelo da entrega (git archive + QC_RAIZ + `entregar.py --direcionada`). | Preencha `$origem` **só no trabalho**. Em casa use `--modo HISTORICO`, sem DB_SEAC. |
| `scripts/regressao_total*.ps1` | Regressão completa: as 36 suítes, cópia nova da produção, um Excel por vez; `_b`/`_c` retomam com `--a-partir`. | Troque `<QC_RAIZ>` pela pasta do repositório. |
| `scripts/retomar45h.ps1` | Retoma suítes que caíram, na mesma cópia instalada (`rodar_suites.py`). | Depois, `_arquitetura/agentes/ferramentas/trocar_producao.py ... --direcionada --com ...`. |
| `scripts/t46.ps1`, `scripts/teste_copia45.py` | Testes direcionados numa cópia. | Idem. |
| `regressao_2026-11-18_run.log` | Log da regressão completa de 18–19/11. | Ver o estado abaixo. |
| `logs/entrega44.log`, `entrega45.log`, `retomar45h.log` | Logs das entregas 44 e 45. | Referência. |

## O que NÃO foi para o GitHub, e por quê

O repositório é **público**.

- **Os `.xlsm` de produção** (`QC_Bioquimica.xlsm`, `QC_Hematologia.xlsm`) e as pastas `_backup_*` e `_entrega_*`: têm
  os dados do laboratório.
  - **Leve por OneDrive ou pen drive.** Eles já ficam na pasta do OneDrive do trabalho, ao lado do repositório.
  - Versão atual: Bio sha256 `a9ba35159655b20b…`, Hema `1d0d9fe8e720798f…` (entrega 45 + ATUALIZAR de 18/11), conferida
    em 19/11.
- **O caminho do DB_SEAC e a senha das abas:** ficam só na máquina do trabalho. A troca da senha é a OC-0067.

## Estado em 19/11/2026, 15h

- **Em produção, validado:** entregas 43, 44 e 45.
  - Ocorrências validadas: OC-0012, 0021, 0048, 0068, 0071–0084.
  - Todas as Críticas e Altas do código estão validadas.
  - A entrega 45 (ADR-098) trata do recorte de linhas manuais.
- **Regressão completa (18–19/11, código 21c8ebd, cópias das produções):**
  - Bioquímica: **36/36 sem falha**.
  - Hematologia: **26/36 sem falha**.
  - Faltam 10 suítes da Hematologia, a partir de `qa_manual_substitui.py`: `qa_westgard_sequencia`, `qa_adr084`,
    `qa_usuarios_adr084`, `qa_adr085`, `qa_adr086`, `qa_adr088`, `qa_adr089`, `qa_adr091` e `qa_adr094`. A sessão caiu
    antes delas.
  - Retomar com `regressao_total_c.ps1`.
- **Próximo passo para declarar "validado":**
  1. Terminar a regressão.
  2. Rodar o ciclo completo do `/auditoria-qc`: frentes F1–F4 nos dois produtos, modo `completo`. A última rodada
     completa foi o `2026-10-31_ciclo01`.
  3. Corrigir o que aparecer, sempre corrigido e entregue na hora.
- **Pendências** (detalhe em `_arquitetura/agentes/OCORRENCIAS.md`):
  - **Laboratório:**
    - OC-0008 (Alta): conferir na bula do lote 522611 (Hema) os alvos de IG# N1–N3, RDW-SD e RDW-CV N2, e
      HGB/MCH/MCV/HCT N2;
    - OC-0011, 0015, 0016 e 0017: especificações a cadastrar ou conferir.
  - **Reconferir na auditoria:**
    - OC-0007: corrigida pelo ADR-091; falta só o reteste;
    - abertas desde 31/10: OC-0010, 0020, 0025, 0037, 0038, 0039, 0045, 0046, 0060, 0065, 0069 e 0070 (Médias);
      OC-0014, 0022, 0026, 0047 e 0061–0064 (Baixas).
  - **Testes:**
    - OC-0085: o Tester confirmar o R06;
    - OC-0087: falta o teste automático do retrato tirado na seleção (R10).
  - **Com o usuário:** OC-0067, a troca da senha.
  - **Encerrada pelo usuário em 18/11:** OC-0086 (opção A).

## Regras que valem em casa também

O detalhe está em `memoria/`.

- **Dados:**
  - teste sempre numa **cópia**;
  - a produção só muda por `entregar.py`, `aplicar_atualizar.py` ou `trocar_producao.py` (backup + sha);
  - nunca altere fórmula, regra de negócio ou critério estatístico sem autorização explícita.
- **Excel:**
  - nenhum Excel em paralelo durante uma entrega;
  - lint (`estudos/lint_vba.py`) antes de abrir o Excel;
  - o Mashup do Power Query às vezes derruba o Excel (FailFast 0x80131623): é ambiental, rode a suíte sozinha de novo.
- **Validação:**
  - quem implementa não valida;
  - Crítico e Alto precisam do Usuário Simulado RESOLVIDA + Tester VALIDADA;
  - Médio e Baixo, só o Tester (decisão de 15/11).
- **Repositório:**
  - trabalhe direto na `main`;
  - antes de cada push, confira o diff: nada de senha, caminho do DB_SEAC, `.xlsm` ou arquivo `~$*` (já no
    `.gitignore`).
- **Mensagens ao usuário:** sempre em português, terminando com a previsão de término e a duração aproximada.
