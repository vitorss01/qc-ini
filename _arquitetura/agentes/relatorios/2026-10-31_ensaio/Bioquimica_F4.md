# Bioquimica — F4 (ensaio de integração, ciclo 2026-10-31_ensaio)

Ensaio: nada foi executado. Nenhum Excel aberto, nenhuma cópia feita. Todos os testes ficam NAO_EXECUTADO (motivo: ensaio).

No ciclo real eu faria:
1. Copiaria `QC_Bioquimica.xlsm` para `%TEMP%\qcw\auditoria\2026-10-31_ensaio\Bioquimica_F4.xlsm`, anotaria o sha256 curto e abriria com `qa_final.QA` (um Excel só, matando apenas o EXCEL_PID impresso).
2. Períodos: filtraria mês a mês e por trimestre no Painel e recalcularia em Python puro média/DP/CV do trimestre a partir dos dados brutos (não pela média dos meses).
3. Lotes: lote A, lote B e H3 vazio (ADR-086), conferindo que Calc/Eng_Saida (não só a tela) usam só os dados filtrados.
4. Combinaria equipamentos, níveis e analitos diferentes.
5. Levey-Jennings: pontos x tblCQ_Final, limites (Calc 184:185, ADR-078), corrida do lote por nível (ADR-080) e violações de Westgard por sequência (ADR-083).
6. Compararia tabela x gráfico e registraria a usabilidade (cliques, textos cortados, filtros, mensagens).

## Ocorrências
- [ENSAIO] Trimestre deve ser calculado dos dados brutos? — tipo duvida_metodo, prioridade Baixo. Fictícia, só para validar o fluxo.
