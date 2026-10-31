# Bioquimica F1 — ciclo 2026-10-31_ensaio (ENSAIO, nada executado)

1. Copiaria `QC_Bioquimica.xlsm` da raiz para `%TEMP%\qcw\auditoria\2026-10-31_ensaio\Bioquimica_F1.xlsm` e anotaria o sha256 curto.
2. Abriria a cópia com `qa_final.QA` (um Excel por vez), anotaria o `EXCEL_PID` e encerraria só esse PID no `finally`.
3. Manual: preencheria a tabela de entrada de CIQ com dados `TESTE-AUDITORIA` e testaria vazio, texto, vírgula/ponto, data futura, analito/nível/lote inexistente e duplicidade.
4. Conferiria o registro gravado e a atualização de média, DP, CV, Levey-Jennings e Westgard.
5. Automático: rodaria `mIntegracao.AtualizarDadosCore` e conferiria o ID `BIO-<id>`, a associação equipamento/analito/lote/nível/data/hora em `tblCQ_Final`, Sync_Log e Audit_Log.
6. Equivalência: entraria o mesmo valor pelos dois caminhos, em cópias separadas, e compararia a estatística em `Value2`.
7. Registraria a usabilidade (clareza dos campos e mensagens) e descartaria as cópias.

Testes: todos NAO_EXECUTADO (motivo: ensaio). Ocorrência: uma fictícia [ENSAIO], funcional, Alto.
