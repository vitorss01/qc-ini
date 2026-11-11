# Entrega 38 (ADR-091, material definido pelo lote): situação em 10/11/2026

## Hematologia: ENTREGUE às 15h11
- As 31 suítes passaram, sem nenhuma falha.
- A produção nova tem sha256 `e3efc3c9f42fdf3a`.
- O backup ficou em `_backup_pre_integracao_2026-11-10_1511`.

## Bioquímica: ENTREGUE às 19h15 (3ª tentativa)
- As 31 suítes passaram, sem nenhuma falha.
- A produção nova tem sha256 `62c98a500ed06856`.
- O backup ficou em `_backup_pre_integracao_2026-11-10_1915`.
- As duas entregas estão registradas no MANIFESTO_PRODUCAO (commit `edc54ef`).

### Histórico das tentativas
- **Na primeira tentativa, 1 teste falhou em 14.** Foi o BN da suíte de blocos por lote, que confere os valores de TESTE do lote 8353.
- **Causa:** a falha estava no teste, não no sistema. Os valores de TESTE foram calculados em 16/10 com os resultados que contavam naquela data. Depois disso, duas regras novas tiraram alguns desses resultados da estatística:
  - o ADR-088, que exclui valores ≤ 0;
  - o ADR-091, pelo qual o lote 8353 é só de PCR.

  O teste não sabia disso e passou a esperar uma média diferente.
- **Correção:** o teste agora recalcula a média considerando os resultados que valiam em 16/10. Nenhuma fórmula nem regra do sistema mudou (commit `6c12f22`). Na cópia, o BN passou, com 194 pares conferidos e nenhuma divergência.
- **2ª tentativa (15h15):** o BN passou. Desta vez falhou o G03, que mede quanto tempo os gráficos levam para se ajustar depois de um zoom de 160%. Rodei o teste de novo numa cópia 3 vezes e passou nas 3 (8 aprovados, 0 falhas). Foi uma falha intermitente de tempo de resposta, não um defeito.
- **3ª tentativa:** em andamento. Ela roda as suítes completas da Bioquímica e deve levar cerca de 3 h.
- **Produção:** a produção da Bioquímica continua a mesma de antes (`e188e160`) e só será trocada se todas as suítes passarem.

## Próximos passos (prazo: quarta, 11/11)
1. Rodar o ATUALIZAR na produção e medir o efeito da regra de lotes.
2. O Tester e o Usuário Simulado validam as 19 ocorrências do ADR-090 e a regra de lotes do ADR-091.
