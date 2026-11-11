# Validação final: situação em 11/11/2026, às 08h30

## 1. ATUALIZAR na produção: concluído às 08h21
Os dois arquivos de produção foram atualizados pelo caminho seguro: primeiro numa cópia, depois blindagem, backup e troca conferida por sha.

| | Bioquímica | Hematologia |
|---|---|---|
| Arquivo de produção (sha) | `62c98a50…` → `8b1cda13…` | `e3efc3c9…` → `f1907656…` |
| Resultados | 34.254 → 35.466 | 44.392 → 48.081 |
| Efeito da regra de lotes | 59 resultados fora do CQ (lotes digitados errado); 30 do lote 8353 tirados dos analitos que não são PCR; urina saiu dos analitos séricos e foi para os de urina | nenhum (a Hema não tem regra cadastrada) |
| Outras mudanças | resultados novos que chegaram | resultados novos que chegaram (o lote 625311 passou de 4 para 123 corridas) |

O detalhe por lote e analito está em `relatorios/efeito_adr091/`.

## 2. Validação independente: em andamento
- **Tester:** modo VALIDAR, com script próprio, cobrindo as 19 ocorrências do ADR-090 e a regra de lotes, mais as suítes de regressão.
- **Usuário Simulado:** reteste dos fluxos e recálculo independente de bias, ET e Sigma. Roda depois do Tester, porque só pode haver um Excel por vez.
- **Especialista em CQ:** parecer técnico sobre a regra de lotes. Roda ao mesmo tempo, sem Excel.

Previsão: entre 13h e 14h. Uma ocorrência só fica VALIDADA quando o Tester disser VALIDADA **e** o Usuário Simulado disser RESOLVIDA.
