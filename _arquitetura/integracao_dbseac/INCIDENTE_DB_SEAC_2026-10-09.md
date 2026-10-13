# Incidente — DB_SEAC apagado durante a gravação (09/10/2026 13:50)

**Detectado:** 13/10/2026 (a entrega dos QCs não encontrou a origem). **Status:** resolvido e com prevenção.

## O que aconteceu
- **A falha:** o orquestrador da atualização automática (`Atualizar-DBSEAC.ps1`, tarefa agendada a cada 2 min)
  encerrou o Excel por TIMEOUT (120 s) no meio de um `Save`. O Excel grava assim: escreve um `.tmp`, apaga o
  original e renomeia o `.tmp`. Ele foi encerrado entre esses passos e o `DB_SEAC.xlsm` ficou apagado. O log mostra
  `TIMEOUT` às 13:50 e, a partir das 13:52, `Arquivo nao encontrado` a cada ciclo.
- **Impacto:**
  - 4 dias sem DB_SEAC;
  - o ATUALIZAR DADOS dos QCs em `MODO_FONTE = SEAC` parou, como manda o ADR-058;
  - nenhum dado foi perdido, porque o DB_SEAC é refeito a partir dos CSVs do SIPEC (janela de 12 meses).
- **Não relacionado** às mudanças dos QCs (ADR-068 a 072).

## Correção (13/10/2026)
1. **Restauração:** o backup automático de 09/10 00:00 foi restaurado (cópia conferida por SHA-256). O ciclo
   seguinte (11:22) atualizou e salvou normalmente. Os `.tmp` deixados pela falha não foram apagados.
2. **Prevenção, com aprovação do usuário:**
   - **Autorrecuperação:** se o DB_SEAC estiver ausente, o orquestrador copia o backup mais recente de
     `_BACKUP_DB_SEAC`, confere o SHA-256 e registra `RESTAURADO`.
   - **Conferência após TIMEOUT:** depois de um TIMEOUT, o orquestrador confere se o arquivo ainda existe e o
     restaura na hora se não existir.
   - **TIMEOUT de 120 para 180 s:** os ciclos normais levam de 30 a 50 s, mas sob carga chegam a 100 s.
   - A versão anterior do script ficou guardada ao lado, como `.bak_20261013`.
   - Testado numa cópia local, com pastas temporárias e outro mutex, antes de entrar em produção. Os ciclos seguintes
     em produção deram OK.
