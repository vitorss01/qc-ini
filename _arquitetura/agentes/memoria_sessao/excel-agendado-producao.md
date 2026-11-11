---
name: excel-agendado-producao
description: Tarefa agendada abre Excel por COM a cada 2 min para atualizar o DB_SEAC de produção — nunca matar EXCEL.EXE por horário de criação
metadata:
  node_type: memory
  type: project
  originSessionId: cc5cf662-a93c-4a60-a8b5-23064798503c
  modified: 2026-10-02T13:11:56.906Z
---

Nesta máquina a tarefa agendada "DB_SEAC - Atualização Automática" roda a cada 2 minutos (wscript → Executar-DBSEAC-Invisivel.vbs na rede `\\SERVIDOR\COMPARTILHAMENTO\INTERFACEAMENTO_DB -  TESTE\AUTOMACAO_DB_SEAC`) e abre um EXCEL.EXE `/automation -Embedding` para rodar ATUALIZAR no DB_SEAC de produção e salvar. Log diário em `...\AUTOMACAO_DB_SEAC\logs\DB_SEAC_<data>.log`.

Em 02/10/2026 matei instâncias "criadas nos últimos N minutos" e derrubei um ciclo (log: "ERRO | O servidor RPC não está disponível"); o ciclo seguinte voltou OK.

**Why:** matar o Excel da tarefa no meio do Save pode corromper o DB_SEAC de produção.
**How to apply:** só encerrar o PID que o próprio script criou (xlh.Excel guarda `self.pid`); registrar PIDs dos testes; nunca filtrar EXCEL.EXE por horário. Ver [[qc-ini-contexto]].

Incidente 09/10/2026 13:50: o TIMEOUT do orquestrador matou o Excel no meio do Save e o DB_SEAC.xlsm sumiu por 4 dias. Em 13/10 restaurei o backup de _BACKUP_DB_SEAC e, com aprovação do usuário, o Atualizar-DBSEAC.ps1 ganhou autorrecuperação (restaura o último backup se o arquivo sumir; confere após TIMEOUT) e TIMEOUT de 180 s (backup do script: .bak_20261013). Doc: INCIDENTE_DB_SEAC_2026-10-09.md. Se a entrega disser ORIGEM_NAO_ENCONTRADA, olhar primeiro o log da automação.
