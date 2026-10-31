# Bioquimica — F2 (CEQ) — ciclo 2026-10-31_ensaio (ENSAIO, nada executado)

1. Copiaria `QC_Bioquimica.xlsm` da raiz para `%TEMP%\qcw\auditoria\2026-10-31_ensaio\` (shutil.copy2) e abriria só a cópia, um Excel, guardando o PID.
2. Registraria um resultado de CEQ pela janela do CEQ com operador/comentário `TESTE-AUDITORIA` e valores sintéticos (programa, rodada, lote, analito, resultado, alvo, unidade).
3. Anotaria antes/depois o bias (`mCEQ.BiasEQ`), o ET e o Sigma do analito escolhido e de um analito vizinho de controle.
4. Corrigiria o mesmo registro pela edição e conferiria que só o analito/equipamento/período certos mudaram, sem duplicidade.
5. Avaliaria sequência de preenchimento, busca da rodada/lote, permissões e preenchimento automático.
6. Encerraria apenas o meu PID e descartaria a cópia; sem captura de tela, sem senha ou UNC, sem tocar em src/ ou fórmulas.
