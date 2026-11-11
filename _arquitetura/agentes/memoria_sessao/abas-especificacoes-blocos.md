---
name: abas-especificacoes-blocos
description: "ADR-074/075 (14/10/2026) — motor virou aba \"Especificações\"; aba \"Analitos\" = um bloco por lote; achar o motor pelo nome"
metadata:
  node_type: memory
  type: project
  originSessionId: cc5cf662-a93c-4a60-a8b5-23064798503c
  modified: 2026-10-14T22:05:26.545Z
---

Desde o ADR-074 (Bioquímica) e o ADR-075 + 074 (Hematologia), a antiga aba **Analitos** chama **Especificações** e é
o MOTOR:
- catálogo A4:A43;
- fonte em S;
- biblioteca anual 2025–2029 em A46:AO89;
- lote em análise.

A aba **Analitos** nova tem um bloco por lote. O bloco começa na linha 2 + 54·k, com os anos de referência
CLIA/VB na linha +45. Ela é mantida por `mLotesBlocos`.

**Why:** o usuário quis Média/DP, limites e o ano CLIA/VB por lote, no layout de um exemplo dele. Renomear a aba fez o
Excel reescrever as cerca de 3.300 referências. Deslocar linhas teria quebrado VBA, instaladores e testes.

**How to apply:**
- Código novo nunca usa 'Analitos' para o motor. O nome se resolve em cada camada:
  - VBA: `mLotes.AbaEspec()`;
  - Python: `ii.aba_espec(wb)`;
  - testes: `QA.espec()`.
- Média/DP continuam só na LotesStore. Os blocos leem por fórmula e gravam por `mLotes.GravarParametroStore`.
- A Hematologia ganhou o layout de colunas da Bioquímica (ADR-075). As duas são iguais em K..W.

Armadilhas descobertas:
- **LET:** `Names.Add` com LET é recusado via COM; nesses casos, ponha a fórmula numa célula e aponte o nome para ela.
  Variável de LET não pode ser `r` ou `c`.
- **FormatConditions:** usam o idioma local. Use `xlCellValue` com números, ou traduza via `FormulaLocal`.

Ver [[qc-ini-contexto]] e [[integracao-db-seac]].
