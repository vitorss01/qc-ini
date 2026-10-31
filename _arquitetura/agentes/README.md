# Arquitetura de agentes do QC_INI

Criada em 31/10/2026. Antes dela, os papéis existiam só nas conversas: gestor/orquestrador, especialista em CQ,
programador e UX/UI (ADR-081), e o auditor independente "Agente 09" (QUALITY_GATE §15). Agora cada papel tem
instruções persistentes, versionadas no repositório. Elas não dependem de uma conversa para funcionar.

## Agentes

| Agente | Onde está definido | Papel | Abre Excel? | Altera código? |
|---|---|---|---|---|
| **Orquestrador** | `.claude/skills/auditoria-qc/SKILL.md` (a sessão principal roda `/auditoria-qc`) | Dispara o ciclo, consolida os pareceres, registra `OCORRENCIAS.md`, pede autorização ao usuário, corrige como programador e entrega | sim (correção e entrega) | sim |
| **Usuário Simulado** (`AGENTE_USUARIO_SIMULADO`) | `.claude/agents/usuario-simulado.md` | Usa o QC numa cópia como um profissional do laboratório; recalcula bias, ET e Sigma de forma independente; produz os Relatórios A–D | sim (cópia) | não |
| **Especialista em CQ** | `.claude/agents/especialista-cq.md` | Parecer de metodologia: erro de implementação, diferença de método ou método questionável; critério de aceite numérico | não | não |
| **UX/UI** | `.claude/agents/ux-ui.md` | Parecer de usabilidade e correção visual concreta; critério de aceite visual | não | não |
| **Tester/QA** | `.claude/agents/tester-qa.md` | Reproduz as ocorrências críticas e altas sozinho; valida correções (nunca as próprias); regressão | sim (cópia) | não |

O Orquestrador é a sessão principal, e não um subagente, por três motivos:

- só a sessão principal aciona outros agentes;
- só ela fala com o usuário, para pedir autorizações;
- só ela roda a entrega em produção.

## Comunicação

```
usuário ──/auditoria-qc──▶ ORQUESTRADOR ──Workflow "auditoria-qc"──┐
                                                                   ▼
        Testar:     usuario-simulado F1 → F2 → F4 (Excel, em série)  ‖  F3 matemática (sem Excel, em paralelo)
        Consolidar: usuario-simulado → relatorios/<ciclo>/00_PARECER.md (Relatórios A, B, C, D)
        Analisar:   especialista-cq (cálculo/dados/método) ‖ ux-ui (usabilidade) ‖ tester-qa (reproduz críticas/altas)
                                                                   │
ORQUESTRADOR ◀── resultado estruturado (JSON) + relatórios .md ────┘
   ├─ registra OC-xxxx em OCORRENCIAS.md, cruza os pareceres, pede autorização quando a regra muda
   ├─ corrige (src → lint → cópia → commit → entregar.py)
   └─ Workflow modo "reteste" + tester-qa VALIDAR ─▶ VALIDADA ou volta para análise
```

- **Os agentes devolvem dados estruturados** (schema do workflow) e gravam o relatório em Markdown na pasta do ciclo.
  O Orquestrador é o único que escreve em `OCORRENCIAS.md`.
- **Fora do ciclo:** o Orquestrador pode chamar um agente avulso (`Agent` com `subagent_type`), por exemplo para
  pedir mais evidência sobre uma ocorrência em que os pareceres divergem.

## Instalação e uso

- **Fonte única** (versionada): `qc-ini/.claude/` (`agents/`, `skills/auditoria-qc/`, `workflows/auditoria-qc.js`).
- **Pasta de abertura:** o Claude Code é aberto em `QC_INI_PRONTO`, um nível acima do repositório, e procura `.claude/`
  lá. Por isso, depois de editar a fonte, rode:
  - `python _arquitetura/agentes/instalar_agentes.py` para copiar a fonte para a pasta de abertura;
  - `python _arquitetura/agentes/instalar_agentes.py --conferir` para saber se as cópias estão em dia.
- **Quando passa a valer:** agentes, skill e workflow novos ou alterados só entram no registro na sessão seguinte.
  Até lá, o workflow funciona pelo caminho do arquivo (`Workflow({scriptPath: ...})`), e cada agente fora do registro
  roda como um agente comum que carrega a própria definição (função `rodar` do workflow).
- **Uso:**
  - `/auditoria-qc` para o ciclo completo;
  - `/auditoria-qc reteste` para retestar as ocorrências em `AGUARDA RETESTE`;
  - `args.ensaio = true` para conferir a ligação sem Excel.

## Pastas

- `REGRAS_COMUNS.md`: o que todo agente obedece (segurança dos dados, método, prioridades).
- `OCORRENCIAS.md`: o registro de todas as ocorrências e o histórico dos ciclos.
- `relatorios/<AAAA-MM-DD_cicloNN>/`: os relatórios de cada ciclo:
  - `<produto>_<frente>.md`;
  - `00_PARECER.md`;
  - `parecer_cq.md`;
  - `parecer_ux.md`;
  - `tester_*.md`.
- `ferramentas/`: os scripts dos agentes, reutilizáveis entre ciclos (calculadora independente, reproduções
  `repro_<id>.py`).

## Frentes do Usuário Simulado

| Frente | Cobre |
|---|---|
| F1 | CIQ manual e automático (DB_SEAC → recebimento), equivalência entre os dois, rastreabilidade |
| F2 | CEQ: registro, edição, consulta, duplicidade, e o efeito de incluir ou corrigir um CEQ no bias, ET e Sigma |
| F3 | Auditoria matemática independente: média, DP, CV, bias, ET, Sigma, ETp/fonte, CIQ+CEQ (somente leitura) |
| F4 | Mês, trimestre, lote, H3 vazio, equipamento e nível; Levey-Jennings, Westgard, tabela × gráfico; usabilidade |
